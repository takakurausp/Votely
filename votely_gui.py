#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
votely_gui.py — Votely 管理ツール GUI 版（完全ローカル）

Google API・認証設定は不要。
起動方法: python votely_gui.py
"""

import io
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

sys.path.insert(0, str(Path(__file__).parent))

# ─── ライブラリチェック ───────────────────────────────────────────────────────
_MISSING: list[str] = []
try:
    import qrcode  # noqa: F401
except ImportError:
    _MISSING.append('qrcode[pil]')
try:
    from PIL import Image  # noqa: F401
except ImportError:
    if 'qrcode[pil]' not in _MISSING:
        _MISSING.append('Pillow')
try:
    from reportlab.lib.pagesizes import A4  # noqa: F401
except ImportError:
    _MISSING.append('reportlab')

# ─── カラーパレット ───────────────────────────────────────────────────────────
C_BG     = '#f5f7fa'
C_PANEL  = '#ffffff'
C_ACCENT = '#3b82f6'
C_HOVER  = '#2563eb'
C_OK     = '#16a34a'
C_ERR    = '#dc2626'
C_WARN   = '#d97706'
C_TEXT   = '#1e293b'
C_SUB    = '#64748b'
C_BORDER = '#e2e8f0'
C_LOG    = '#0f172a'

F_TITLE  = ('Yu Gothic UI', 14, 'bold')
F_NORMAL = ('Yu Gothic UI', 10)
F_SMALL  = ('Yu Gothic UI', 9)
F_MONO   = ('Consolas', 9)

CONFIG_FILE = Path(__file__).parent / 'votely_config.json'


# =============================================================================
# ログリダイレクター
# =============================================================================

class _QWriter(io.TextIOBase):
    def __init__(self, q: queue.Queue):
        self._q = q
    def write(self, s: str) -> int:
        if s and s.strip():
            self._q.put(('log', s.strip()))
        return len(s)
    def flush(self): pass


# =============================================================================
# メインウィンドウ
# =============================================================================

class VotelyApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('Votely 管理ツール')
        self.geometry('700x620')
        self.minsize(600, 520)
        self.configure(bg=C_BG)

        self._q: queue.Queue = queue.Queue()
        self._cfg: dict = {}
        self._busy = False

        self._build_ui()
        self._load_config()

        if _MISSING:
            self._show_missing()

        self.after(100, self._poll)

    # ─── UI 構築 ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ヘッダー
        hdr = tk.Frame(self, bg=C_ACCENT, pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text='🗳  Votely 管理ツール',
                 font=('Yu Gothic UI', 14, 'bold'),
                 fg='white', bg=C_ACCENT).pack(side='left', padx=16)

        # 投票ページ URL（タブ共通）
        url_frame = self._section(self, '投票ページ URL')
        url_frame.pack(fill='x', padx=12, pady=(10, 4))

        uf = tk.Frame(url_frame, bg=C_PANEL)
        uf.pack(fill='x', padx=8, pady=8)
        uf.columnconfigure(1, weight=1)

        tk.Label(uf, text='URL:', font=F_SMALL, bg=C_PANEL,
                 fg=C_TEXT).grid(row=0, column=0, padx=(0,8), sticky='e')
        self._url_var = tk.StringVar()
        tk.Entry(uf, textvariable=self._url_var, font=F_NORMAL,
                 relief='flat', bg='#f1f5f9',
                 highlightthickness=1, highlightbackground=C_BORDER
                 ).grid(row=0, column=1, sticky='ew')
        self._btn(uf, '保存', self._save_config, small=True
                  ).grid(row=0, column=2, padx=(8,0))

        tk.Label(uf,
                 text='例: https://your-id.github.io/Votely　（末尾スラッシュなし）',
                 font=('Yu Gothic UI', 8), bg=C_PANEL, fg=C_SUB
                 ).grid(row=1, column=1, sticky='w', pady=(2,0))

        # タブ
        style = ttk.Style(self)
        style.configure('TNotebook', background=C_BG, borderwidth=0)
        style.configure('TNotebook.Tab', font=F_NORMAL, padding=[14, 6])
        style.map('TNotebook.Tab',
                  background=[('selected', C_PANEL), ('!selected', '#dde3ec')],
                  foreground=[('selected', C_ACCENT)])

        nb = ttk.Notebook(self)
        nb.pack(fill='x', padx=12, pady=4)

        t1 = tk.Frame(nb, bg=C_PANEL)
        t2 = tk.Frame(nb, bg=C_PANEL)
        nb.add(t1, text='①  メールCSV → トークンCSV')
        nb.add(t2, text='②  当日参加者チケット')

        self._build_tab_tokens(t1)
        self._build_tab_guests(t2)

        # ログ
        log_frame = self._section(self, 'ログ')
        log_frame.pack(fill='both', expand=True, padx=12, pady=(4, 4))

        self._log = scrolledtext.ScrolledText(
            log_frame, font=F_MONO, bg=C_LOG, fg='#94a3b8',
            relief='flat', bd=0, wrap='word', height=10,
            state='disabled',
        )
        self._log.pack(fill='both', expand=True, padx=8, pady=(6,2))
        self._log.tag_config('info',    foreground='#94a3b8')
        self._log.tag_config('ok',      foreground='#4ade80')
        self._log.tag_config('err',     foreground='#f87171')
        self._log.tag_config('warn',    foreground='#fbbf24')
        self._log.tag_config('head',    foreground='#60a5fa',
                             font=(*F_MONO[:2], 'bold'))

        br = tk.Frame(log_frame, bg=C_PANEL)
        br.pack(anchor='e', padx=8, pady=(0,6))
        self._btn(br, 'ログをクリア', self._clear_log, small=True).pack()

        # ステータスバー
        self._status = tk.StringVar(value='● 待機中')
        tk.Label(self, textvariable=self._status,
                 font=F_SMALL, fg=C_SUB, bg=C_BORDER,
                 anchor='w', padx=10).pack(fill='x', side='bottom')

    # ─── タブ①: トークン一括発行 ─────────────────────────────────────────────

    def _build_tab_tokens(self, parent):
        frm = tk.Frame(parent, bg=C_PANEL)
        frm.pack(fill='x', padx=16, pady=12)
        frm.columnconfigure(1, weight=1)

        # 入力 CSV
        tk.Label(frm, text='入力 CSV:', font=F_SMALL, bg=C_PANEL,
                 fg=C_TEXT, anchor='e').grid(row=0, column=0, sticky='e',
                                              padx=(0,8), pady=4)
        r0 = tk.Frame(frm, bg=C_PANEL)
        r0.grid(row=0, column=1, sticky='ew', pady=4)
        r0.columnconfigure(0, weight=1)
        self._in_csv_var = tk.StringVar()
        tk.Entry(r0, textvariable=self._in_csv_var, font=F_NORMAL,
                 relief='flat', bg='#f1f5f9',
                 highlightthickness=1, highlightbackground=C_BORDER
                 ).grid(row=0, column=0, sticky='ew')
        self._btn(r0, '参照...', self._browse_input_csv,
                  small=True).grid(row=0, column=1, padx=(6,0))
        tk.Label(frm, text='1 列目にメールアドレスが入った CSV',
                 font=('Yu Gothic UI', 8), bg=C_PANEL, fg=C_SUB
                 ).grid(row=1, column=1, sticky='w')

        # 出力 CSV
        tk.Label(frm, text='出力 CSV:', font=F_SMALL, bg=C_PANEL,
                 fg=C_TEXT, anchor='e').grid(row=2, column=0, sticky='e',
                                              padx=(0,8), pady=(10,4))
        r2 = tk.Frame(frm, bg=C_PANEL)
        r2.grid(row=2, column=1, sticky='ew', pady=(10,4))
        r2.columnconfigure(0, weight=1)
        self._out_csv_var = tk.StringVar(value='tokens.csv')
        tk.Entry(r2, textvariable=self._out_csv_var, font=F_NORMAL,
                 relief='flat', bg='#f1f5f9',
                 highlightthickness=1, highlightbackground=C_BORDER
                 ).grid(row=0, column=0, sticky='ew')
        self._btn(r2, '参照...', self._browse_output_csv,
                  small=True).grid(row=0, column=1, padx=(6,0))
        tk.Label(frm,
                 text='スプレッドシートへのインポート用 CSV（メールアドレス / トークン / URL / FALSE）',
                 font=('Yu Gothic UI', 8), bg=C_PANEL, fg=C_SUB
                 ).grid(row=3, column=1, sticky='w')

        self._btn_tok = self._btn(
            parent, '▶  トークンを一括発行', self._run_tokens, width=22)
        self._btn_tok.pack(pady=(4, 14))

    # ─── タブ②: 当日参加者チケット ───────────────────────────────────────────

    def _build_tab_guests(self, parent):
        frm = tk.Frame(parent, bg=C_PANEL)
        frm.pack(fill='x', padx=16, pady=12)
        frm.columnconfigure(1, weight=1)

        def row(r, label, widget_fn):
            tk.Label(frm, text=label, font=F_SMALL, bg=C_PANEL,
                     fg=C_TEXT, anchor='e').grid(
                         row=r, column=0, sticky='e', padx=(0,8), pady=4)
            widget_fn(r)

        def entry(r, var):
            tk.Entry(frm, textvariable=var, font=F_NORMAL,
                     relief='flat', bg='#f1f5f9',
                     highlightthickness=1, highlightbackground=C_BORDER
                     ).grid(row=r, column=1, sticky='ew', pady=4)

        def file_row(r, var, browse_fn):
            f = tk.Frame(frm, bg=C_PANEL)
            f.grid(row=r, column=1, sticky='ew', pady=4)
            f.columnconfigure(0, weight=1)
            tk.Entry(f, textvariable=var, font=F_NORMAL,
                     relief='flat', bg='#f1f5f9',
                     highlightthickness=1, highlightbackground=C_BORDER
                     ).grid(row=0, column=0, sticky='ew')
            self._btn(f, '参照...', browse_fn,
                      small=True).grid(row=0, column=1, padx=(6,0))

        # 発行枚数
        self._g_num = tk.IntVar(value=30)
        row(0, '発行枚数:', lambda r: tk.Spinbox(
            frm, from_=1, to=500, textvariable=self._g_num,
            width=7, font=F_NORMAL, relief='flat', bg='#f1f5f9'
        ).grid(row=r, column=1, sticky='w', pady=4))

        # タイトル・説明
        self._g_title = tk.StringVar(value='Votely 投票チケット')
        self._g_desc  = tk.StringVar(value='このQRコードを読み取って投票してください')
        row(1, 'タイトル:', lambda r: entry(r, self._g_title))
        row(2, '説明文:',   lambda r: entry(r, self._g_desc))

        # 出力 CSV
        self._g_csv = tk.StringVar(value='guest_tokens.csv')
        row(3, '出力 CSV:', lambda r: file_row(r, self._g_csv,
                                                self._browse_guest_csv))

        # 出力 PDF
        self._g_pdf = tk.StringVar(value='guest_tickets.pdf')
        row(4, '出力 PDF:', lambda r: file_row(r, self._g_pdf,
                                                self._browse_guest_pdf))

        self._btn_guest = self._btn(
            parent, '🎫  チケットを発行', self._run_guests, width=22)
        self._btn_guest.pack(pady=(4, 14))

    # ─── ウィジェットヘルパー ─────────────────────────────────────────────────

    def _section(self, parent, title: str) -> tk.LabelFrame:
        return tk.LabelFrame(parent, text=f'  {title}  ',
                             font=F_SMALL, bg=C_PANEL, fg=C_SUB,
                             relief='flat', bd=1,
                             highlightthickness=1,
                             highlightbackground=C_BORDER)

    def _btn(self, parent, text, cmd, small=False, width=None):
        kw = dict(text=text, command=cmd,
                  font=F_SMALL if small else F_NORMAL,
                  bg=C_ACCENT, fg='white',
                  activebackground=C_HOVER, activeforeground='white',
                  relief='flat', cursor='hand2',
                  padx=10 if small else 16,
                  pady=4  if small else 8)
        if width:
            kw['width'] = width
        b = tk.Button(parent, **kw)
        b.bind('<Enter>', lambda _: b.configure(bg=C_HOVER))
        b.bind('<Leave>', lambda _: b.configure(bg=C_ACCENT))
        return b

    # ─── ファイルダイアログ ───────────────────────────────────────────────────

    def _browse_input_csv(self):
        p = filedialog.askopenfilename(
            title='メールアドレス CSV を選択',
            filetypes=[('CSV', '*.csv'), ('すべて', '*.*')])
        if p:
            self._in_csv_var.set(p)

    def _browse_output_csv(self):
        p = filedialog.asksaveasfilename(
            title='出力 CSV の保存先',
            defaultextension='.csv',
            filetypes=[('CSV', '*.csv')],
            initialfile=self._out_csv_var.get())
        if p:
            self._out_csv_var.set(p)

    def _browse_guest_csv(self):
        p = filedialog.asksaveasfilename(
            title='ゲストトークン CSV の保存先',
            defaultextension='.csv',
            filetypes=[('CSV', '*.csv')],
            initialfile=self._g_csv.get())
        if p:
            self._g_csv.set(p)

    def _browse_guest_pdf(self):
        p = filedialog.asksaveasfilename(
            title='チケット PDF の保存先',
            defaultextension='.pdf',
            filetypes=[('PDF', '*.pdf')],
            initialfile=self._g_pdf.get())
        if p:
            self._g_pdf.set(p)

    # ─── 設定 ────────────────────────────────────────────────────────────────

    def _load_config(self):
        try:
            if CONFIG_FILE.exists():
                self._cfg = json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
                self._url_var.set(self._cfg.get('pages_url', ''))
        except Exception:
            pass

    def _save_config(self):
        self._cfg['pages_url'] = self._url_var.get().strip().rstrip('/')
        try:
            CONFIG_FILE.write_text(
                json.dumps(self._cfg, ensure_ascii=False, indent=2),
                encoding='utf-8')
            self._log_put('[設定] 投票ページ URL を保存しました。', 'ok')
        except Exception as e:
            self._log_put(f'[エラー] 設定保存失敗: {e}', 'err')

    # ─── バリデーション ───────────────────────────────────────────────────────

    def _check_url(self) -> str | None:
        url = self._url_var.get().strip().rstrip('/')
        if not url:
            messagebox.showwarning('未入力', '投票ページ URL を入力してください。')
            return None
        if not url.startswith('http'):
            messagebox.showwarning('URL エラー', 'URL は http:// または https:// で始めてください。')
            return None
        if _MISSING:
            messagebox.showerror('ライブラリ不足',
                                 '以下をインストールしてください:\n\n'
                                 f"pip install {' '.join(_MISSING)}")
            return None
        return url

    # ─── 実行（トークン） ─────────────────────────────────────────────────────

    def _run_tokens(self):
        url = self._check_url()
        if not url:
            return
        in_csv  = self._in_csv_var.get().strip()
        out_csv = self._out_csv_var.get().strip() or 'tokens.csv'
        if not in_csv:
            messagebox.showwarning('未入力', '入力 CSV を選択してください。')
            return
        if not Path(in_csv).exists():
            messagebox.showerror('ファイルなし', f'ファイルが見つかりません:\n{in_csv}')
            return
        self._log_put('━' * 48, 'head')
        self._log_put('▶ トークン一括発行を開始します', 'head')
        self._set_status('発行中...', C_ACCENT)
        self._run_bg(self._task_tokens, in_csv, url, out_csv)

    def _task_tokens(self, in_csv, url, out_csv):
        old = sys.stdout
        sys.stdout = _QWriter(self._q)
        try:
            from votely_tool import generate_tokens
            generate_tokens(in_csv, url, out_csv)
            self._q.put(('ok', f'完了 → {out_csv}'))
        except Exception as e:
            self._q.put(('err', f'[エラー] {e}'))
        finally:
            sys.stdout = old

    # ─── 実行（ゲストチケット） ───────────────────────────────────────────────

    def _run_guests(self):
        url = self._check_url()
        if not url:
            return
        num = self._g_num.get()
        if num <= 0:
            messagebox.showwarning('未入力', '発行枚数は 1 以上を入力してください。')
            return
        self._log_put('━' * 48, 'head')
        self._log_put(f'🎫 チケット発行を開始します（{num} 枚）', 'head')
        self._set_status('発行中...', C_ACCENT)
        self._run_bg(self._task_guests, num, url,
                     self._g_title.get().strip(),
                     self._g_desc.get().strip(),
                     self._g_csv.get().strip() or 'guest_tokens.csv',
                     self._g_pdf.get().strip() or 'guest_tickets.pdf')

    def _task_guests(self, num, url, title, desc, out_csv, out_pdf):
        old = sys.stdout
        sys.stdout = _QWriter(self._q)
        try:
            from votely_tool import generate_guest_tickets
            generate_guest_tickets(num, url, title, desc, out_csv, out_pdf)
            self._q.put(('ok', f'完了 → {out_pdf}'))
            self._q.put(('open', out_pdf))
        except Exception as e:
            self._q.put(('err', f'[エラー] {e}'))
        finally:
            sys.stdout = old

    # ─── バックグラウンド実行 ─────────────────────────────────────────────────

    def _run_bg(self, fn, *args):
        def _wrap():
            try:
                fn(*args)
            except Exception as e:
                self._q.put(('err', f'[エラー] {e}'))
            finally:
                self._q.put(('done', None))
        self._set_busy(True)
        threading.Thread(target=_wrap, daemon=True).start()

    def _set_busy(self, v: bool):
        self._busy = v
        state = 'disabled' if v else 'normal'
        self._btn_tok.configure(state=state)
        self._btn_guest.configure(state=state)

    # ─── キューポーリング ─────────────────────────────────────────────────────

    def _poll(self):
        try:
            while True:
                kind, val = self._q.get_nowait()
                if kind == 'log':
                    self._classify(val)
                elif kind == 'ok':
                    self._log_put(f'✅ {val}', 'ok')
                    self._set_status(val, C_OK)
                elif kind == 'err':
                    self._log_put(val, 'err')
                    self._set_status('エラーが発生しました', C_ERR)
                elif kind == 'done':
                    self._set_busy(False)
                elif kind == 'open':
                    self._open_file(val)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _classify(self, msg: str):
        if any(k in msg for k in ('エラー', 'Error', '失敗')):
            tag = 'err'
        elif any(k in msg for k in ('完了', '✅', '成功', '保存', '出力')):
            tag = 'ok'
        elif any(k in msg for k in ('警告', '注意')):
            tag = 'warn'
        else:
            tag = 'info'
        self._log_put(msg, tag)

    def _log_put(self, msg: str, tag: str = 'info'):
        self._log.configure(state='normal')
        self._log.insert('end', msg + '\n', tag)
        self._log.configure(state='disabled')
        self._log.see('end')

    def _clear_log(self):
        self._log.configure(state='normal')
        self._log.delete('1.0', 'end')
        self._log.configure(state='disabled')

    def _set_status(self, msg: str, color: str = C_SUB):
        self._status.set(f'● {msg}')

    def _open_file(self, path: str):
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            self._log_put(f'[警告] PDF を自動で開けません: {e}', 'warn')

    def _show_missing(self):
        self._log_put('━' * 48, 'head')
        self._log_put('⚠  必要なライブラリが不足しています', 'warn')
        self._log_put(f"  pip install {' '.join(_MISSING)}", 'warn')
        self._log_put('━' * 48, 'head')


# =============================================================================

if __name__ == '__main__':
    VotelyApp().mainloop()
