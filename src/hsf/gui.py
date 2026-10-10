# -*- coding: utf-8 -*-
"""图形界面：拖拽/选择源表与模板 -> 选择输出目录 -> 执行 -> 进度与错误展示。

设计要点：
  * 耗时流程放在后台线程，主线程只负责 UI，避免界面假死
  * 后台线程所有 UI 更新通过 `root.after(0, ...)` 投递回主线程
  * tkinterdnd2 可选：装了就支持拖拽，没装自动降级为纯按钮选择
"""
from __future__ import annotations

import os
import queue
import re
import threading
import time
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, simpledialog, ttk

from .config import Settings
from .pipeline import Pipeline
from .util import extract_prefix

TITLE = 'HSF 均质表自动填写工具'
GEOMETRY = '760x620'

FILE_TYPES_XLS = [('Excel 97-2003 源表', '*.xls'), ('所有文件', '*.*')]
FILE_TYPES_XLSX = [('Excel 模板', '*.xlsx'), ('所有文件', '*.*')]


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(TITLE)
        self.root.geometry(GEOMETRY)
        self.root.minsize(680, 560)

        self.src_var = tk.StringVar()
        self.tpl_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.status_var = tk.StringVar(value='就绪：先选择源表与模板')
        self.running = False
        self._prefix = None          # 由 start() 预检时确定
        self._q: queue.Queue = queue.Queue()

        self._build()
        self._init_dnd()
        self.root.after(80, self._pump)

    # ------------------------------------------------------------ UI 构建
    def _build(self):
        root = self.root
        pad = {'padx': 10, 'pady': (6, 4)}

        top = ttk.Frame(root)
        top.pack(fill='x', **pad)

        ttk.Label(top, text='Step 1 · 源材料成分表 (.xls，内含检测报告 PDF)',
                  font=('Microsoft YaHei UI', 9, 'bold')).pack(anchor='w')
        f1 = ttk.Frame(top)
        f1.pack(fill='x', pady=(2, 0))
        self.src_entry = ttk.Entry(f1, textvariable=self.src_var)
        self.src_entry.pack(side='left', fill='x', expand=True)
        ttk.Button(f1, text='选择…', width=10,
                   command=self.pick_src).pack(side='left', padx=(6, 0))

        ttk.Label(top, text='Step 2 · 均质创建导入模板 (.xlsx)',
                  font=('Microsoft YaHei UI', 9, 'bold')).pack(anchor='w', pady=(10, 0))
        f2 = ttk.Frame(top)
        f2.pack(fill='x', pady=(2, 0))
        self.tpl_entry = ttk.Entry(f2, textvariable=self.tpl_var)
        self.tpl_entry.pack(side='left', fill='x', expand=True)
        ttk.Button(f2, text='选择…', width=10,
                   command=self.pick_tpl).pack(side='left', padx=(6, 0))

        ttk.Label(top, text='Step 3 · 输出目录（全平铺，每次运行都需指定）',
                  font=('Microsoft YaHei UI', 9, 'bold')).pack(anchor='w', pady=(10, 0))
        f3 = ttk.Frame(top)
        f3.pack(fill='x', pady=(2, 0))
        self.out_entry = ttk.Entry(f3, textvariable=self.out_var)
        self.out_entry.pack(side='left', fill='x', expand=True)
        ttk.Button(f3, text='选择…', width=10,
                   command=self.pick_out).pack(side='left', padx=(6, 0))

        bar = ttk.Frame(root)
        bar.pack(fill='x', padx=10, pady=(12, 2))
        self.run_btn = ttk.Button(bar, text='开始生成', command=self.start)
        self.run_btn.pack(side='left')
        self.open_btn = ttk.Button(bar, text='打开输出目录', command=self.open_out,
                                   state='disabled')
        self.open_btn.pack(side='left', padx=(8, 0))
        ttk.Button(bar, text='清空日志', command=self.clear_log).pack(side='right')

        pf = ttk.Frame(root)
        pf.pack(fill='x', padx=10, pady=(8, 2))
        self.prog = ttk.Progressbar(pf, mode='determinate', maximum=100)
        self.prog.pack(fill='x')
        ttk.Label(pf, textvariable=self.status_var,
                  foreground='#555555').pack(anchor='w', pady=(2, 0))

        nb = ttk.Notebook(root)
        nb.pack(fill='both', expand=True, padx=10, pady=(8, 10))

        self.log_box = tk.Text(nb, height=8, wrap='none', font=('Consolas', 9))
        sb = ttk.Scrollbar(self.log_box, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.log_box.pack(fill='both', expand=True)
        nb.add(self.log_box, text='运行日志')

        self.err_box = tk.Text(nb, height=8, wrap='none', font=('Consolas', 9))
        sb2 = ttk.Scrollbar(self.err_box, command=self.err_box.yview)
        self.err_box.configure(yscrollcommand=sb2.set)
        sb2.pack(side='right', fill='y')
        self.err_box.pack(fill='both', expand=True)
        nb.add(self.err_box, text='错误 / 告警 (0)')

        for box in (self.log_box, self.err_box):
            box.tag_config('INFO', foreground='#333333')
            box.tag_config('WARN', foreground='#b26a00')
            box.tag_config('ERROR', foreground='#c00000')
            box.tag_config('OK', foreground='#1a7f37')

    # ------------------------------------------------------------ 拖拽
    def _init_dnd(self):
        self.dnd_ok = False
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
            for widget in (self.src_entry, self.tpl_entry, self.out_entry,
                           self.root):
                try:
                    widget.drop_target_register(DND_FILES)
                except Exception:
                    pass
            self.src_entry.dnd_bind('<<Drop>>', lambda e: self._drop(e, 'src'))
            self.tpl_entry.dnd_bind('<<Drop>>', lambda e: self._drop(e, 'tpl'))
            self.out_entry.dnd_bind('<<Drop>>', lambda e: self._drop(e, 'out'))
            self.root.dnd_bind('<<Drop>>', self._drop_auto)
            self.dnd_ok = True
        except Exception:
            self.dnd_ok = False

    @staticmethod
    def _clean_path(raw: str) -> str:
        s = (raw or '').strip()
        if s.startswith('{') and s.endswith('}'):
            s = s[1:-1]
        return s.strip('"').strip()

    def _drop(self, event, kind: str):
        path = self._clean_path(event.data)
        if not path:
            return
        if kind == 'src':
            self.src_var.set(path)
        elif kind == 'tpl':
            self.tpl_var.set(path)
        else:
            self.out_var.set(path)
        self._hint_out_dir()

    def _drop_auto(self, event):
        """拖到窗口空白处：按扩展名自动判断落在哪个输入框。"""
        path = self._clean_path(event.data)
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext == '.xls':
            self.src_var.set(path)
        elif ext == '.xlsx':
            self.tpl_var.set(path)
        elif os.path.isdir(path):
            self.out_var.set(path)
        else:
            self.src_var.set(path)
        self._hint_out_dir()

    def _hint_out_dir(self):
        """源表选定后，默认把输出目录放到源表同级的「输出」子目录。"""
        src = self.src_var.get()
        if src and not self.out_var.get():
            base = os.path.dirname(src)
            self.out_var.set(os.path.join(base, 'HSF输出'))

    # ------------------------------------------------------------ 选择
    def pick_src(self):
        p = filedialog.askopenfilename(title='选择材料成分表 (.xls)',
                                       filetypes=FILE_TYPES_XLS)
        if p:
            self.src_var.set(p)
            self._hint_out_dir()

    def pick_tpl(self):
        p = filedialog.askopenfilename(title='选择均质创建导入模板 (.xlsx)',
                                       filetypes=FILE_TYPES_XLSX)
        if p:
            self.tpl_var.set(p)

    def pick_out(self):
        p = filedialog.askdirectory(title='选择输出目录')
        if p:
            self.out_var.set(p)

    def open_out(self):
        d = self.out_var.get()
        if d and os.path.isdir(d):
            try:
                os.startfile(d)  # noqa: S606 - Windows only convenience
            except Exception:
                messagebox.showinfo('输出目录', d)

    def clear_log(self):
        for box in (self.log_box, self.err_box):
            box.configure(state='normal')
            box.delete('1.0', 'end')
            box.configure(state='disabled')
        self.err_count = 0
        try:
            self.root.nametowidget(self.err_box)  # keep ref
        except Exception:
            pass
        self._refresh_err_tab(0)

    # ------------------------------------------------------------ 日志输出
    def _write(self, box: tk.Text, text: str, tag: str = 'INFO'):
        box.configure(state='normal')
        box.insert('end', text + '\n', (tag,))
        box.see('end')
        box.configure(state='disabled')

    def _refresh_err_tab(self, n: int):
        idx = 1
        self.root.after(0, lambda: None)
        try:
            nb = self.err_box.master
            cur = nb.index(nb.select()) if nb.select() else 0
            nb.tab(idx, text=f'错误 / 告警 ({n})')
        except Exception:
            pass

    def log(self, text: str, tag: str = 'INFO'):
        self._write(self.log_box, text, tag)
        if tag in ('WARN', 'ERROR'):
            self._write(self.err_box, text, tag)

    # ------------------------------------------------------------ 执行
    def _validate(self) -> str | None:
        src, tpl, out = self.src_var.get(), self.tpl_var.get(), self.out_var.get()
        if not src or not os.path.isfile(src):
            return '请选择有效的源材料成分表 (.xls)'
        if os.path.splitext(src)[1].lower() not in ('.xls', '.xlsx'):
            return '源表必须是 .xls（Excel 97-2003，含内嵌 PDF）'
        if not tpl or not os.path.isfile(tpl):
            return '请选择有效的均质创建导入模板 (.xlsx)'
        if os.path.splitext(tpl)[1].lower() != '.xlsx':
            return '模板必须是 .xlsx 文件'
        if not out:
            return '请选择输出目录'
        return None

    def _ask_prefix(self, src: str) -> str | None:
        """源文件名里抓不到 RR 编号时，主线程弹窗问用户。

        必须在启动后台线程之前做——Tk 不是线程安全的，从工作线程弹对话框会崩。
        返回 None 表示用户取消。
        """
        fname = os.path.basename(src)
        val = simpledialog.askstring(
            '需要芯片编号',
            '源文件名里没有识别出芯片编号（形如 RR12345）：\n\n'
            f'    {fname}\n\n'
            '请手动输入编号，将用于均质名称与 PDF 文件名\n'
            '（例如输入 RR12345，得到 RR12345-Die1）：',
            parent=self.root, initialvalue='RR')
        if val is None:
            return None
        val = re.sub(r'\s+', '', val).upper()
        return val or None

    def start(self):
        if self.running:
            return
        err = self._validate()
        if err:
            messagebox.showwarning('无法开始', err)
            return

        # 芯片编号预检：抓不到就先问清楚再开工（2026-08-31）
        prefix = extract_prefix(self.src_var.get())
        if not prefix:
            prefix = self._ask_prefix(self.src_var.get())
            if not prefix:
                messagebox.showwarning(
                    '已取消', '未提供芯片编号，无法生成均质名称。\n'
                              '建议把源表文件名改成包含编号的形式，例如 '
                              '「材料成分表 RR12345.xls」。')
                return
        self._prefix = prefix

        self.running = True
        self.run_btn.configure(state='disabled', text='正在生成…')
        self.open_btn.configure(state='disabled')
        self.prog['value'] = 0
        self.status_var.set('正在启动…')
        self.clear_log()
        self.log('=' * 72)
        self.log(f'源表   ：{self.src_var.get()}')
        self.log(f'模板   ：{self.tpl_var.get()}')
        self.log(f'输出   ：{self.out_var.get()}')
        self.log('=' * 72, 'OK')

        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        t0 = time.time()
        try:
            pipe = Pipeline(
                self.src_var.get(), self.tpl_var.get(), self.out_var.get(),
                Settings(),
                progress=lambda msg, pct: self._q.put(('prog', msg, pct)),
                logger=lambda line, level: self._q.put(('log', line, level)),
                prefix_override=self._prefix)
            res = pipe.run()
            self._q.put(('done', res, time.time() - t0))
        except Exception:
            self._q.put(('fatal', traceback.format_exc(), time.time() - t0))

    # ------------------------------------------------------------ 主线程泵
    def _pump(self):
        try:
            while True:
                item = self._q.get_nowait()
                kind = item[0]
                if kind == 'prog':
                    _, msg, pct = item
                    self.prog['value'] = max(0, min(100, pct))
                    self.status_var.set(msg)
                    self.log(f'[{pct:3d}%] {msg}')
                elif kind == 'log':
                    _, line, level = item
                    tag = {'WARN': 'WARN', 'ERROR': 'ERROR'}.get(level, 'INFO')
                    self.log(line, tag)
                elif kind == 'done':
                    self._on_done(item[1], item[2])
                elif kind == 'fatal':
                    self._on_fatal(item[1])
        except queue.Empty:
            pass
        self.root.after(80, self._pump)

    def _on_fatal(self, tb: str):
        self.running = False
        self.run_btn.configure(state='normal', text='开始生成')
        self.prog['value'] = 100
        self.status_var.set('运行失败，详见「错误 / 告警」页')
        self.log(tb, 'ERROR')
        self._finish_err_tab()
        messagebox.showerror('运行失败', '处理过程抛出异常，请查看「错误 / 告警」页。')

    def _on_done(self, res, elapsed: float):
        self.running = False
        self.run_btn.configure(state='normal', text='开始生成')
        self.open_btn.configure(state='normal')
        self.prog['value'] = 100

        for e in res.errors:
            self.log('ERROR: ' + e, 'ERROR')
        n_warn = 0
        for rec in res.records:
            for w in rec.warnings:
                n_warn += 1
                self.log(f'WARN  [{rec.name}] {w}', 'WARN')

        if res.errors:
            self.status_var.set(f'完成但有 {len(res.errors)} 个致命错误，'
                                f'耗时 {elapsed:.1f}s')
            self.log(f'完成（含 {len(res.errors)} 个致命错误），耗时 {elapsed:.1f}s',
                     'ERROR')
        else:
            n_skip = len(getattr(res, 'skipped', []) or [])
            tail = f'，{n_warn} 条告警' if n_warn else ''
            tail += f'，跳过 {n_skip} 个空块' if n_skip else ''
            self.status_var.set(f'完成：表 {res.n_xlsx} 份 / HSF PDF {res.n_hsf} 份 / '
                                f'MSDS PDF {res.n_msds} 份，耗时 {elapsed:.1f}s' + tail)
            self.log(f'完成：生成表 {res.n_xlsx} 份，HSF PDF {res.n_hsf} 份，'
                     f'MSDS PDF {res.n_msds} 份，耗时 {elapsed:.1f}s', 'OK')
            if n_skip:
                self.log(f'跳过 {n_skip} 个空块（详见汇总报告「跳过的空块」页）', 'WARN')
            self.log(f'输出目录：{res.out_dir}', 'OK')

        self._finish_err_tab(len(res.errors) + n_warn)

    def _finish_err_tab(self, n=None):
        try:
            nb = self.err_box.master
            if n is None:
                txt = self.err_box.get('1.0', 'end')
                n = sum(1 for ln in txt.split('\n') if ln.startswith(('ERROR', 'WARN')))
            nb.tab(1, text=f'错误 / 告警 ({n})')
        except Exception:
            pass


def launch():
    """入口：优先用 TkinterDnD 根窗口（支持拖拽），否则退回普通 Tk。"""
    try:
        from tkinterdnd2 import TkinterDnD  # type: ignore
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    launch()
