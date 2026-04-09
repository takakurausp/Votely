#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
votely_tool.py — Votely 管理ツール（外部 Python スクリプト版）

GAS の実行時間制限（6 分）を回避するため、トークン生成や PDF 出力を
ローカル Python で処理し、結果だけ Google スプレッドシートに書き込む。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
必要ライブラリ（初回のみ実行）:
  pip install gspread google-auth google-auth-oauthlib \
              qrcode[pil] reportlab Pillow

認証ファイルの準備（どちらか一方）:
  [A] サービスアカウント（推奨・自動化向き）
      Google Cloud Console でサービスアカウントを作成 → JSON キーをダウンロード
      → スプレッドシートをそのアカウントのメールアドレスに「編集者」として共有

  [B] OAuth2 クライアントシークレット（初回のみブラウザ認証）
      Google Cloud Console で「デスクトップアプリ」用 OAuth2 クライアント ID を作成
      → JSON をダウンロード → 初回実行時にブラウザが開き認証完了

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
使い方:

  # 名簿シートにトークンを一括発行（メール未送信の行のみ）
  python votely_tool.py generate-tokens \\
      --spreadsheet-id <SPREADSHEET_ID> \\
      --credentials service_account.json

  # 当日参加者チケット 30 枚を発行して PDF を生成
  python votely_tool.py guest-tickets 30 \\
      --spreadsheet-id <SPREADSHEET_ID> \\
      --credentials service_account.json \\
      --title "懇親会投票チケット" \\
      --desc "このQRコードを読み取って投票してください" \\
      --output guest_tickets.pdf

  # スプレッドシート ID を config ファイルに保存して省略する
  python votely_tool.py config \\
      --spreadsheet-id <SPREADSHEET_ID> \\
      --credentials service_account.json
  python votely_tool.py generate-tokens   # 以降は --spreadsheet-id 省略可
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import argparse
import io
import json
import os
import re
import sys
import uuid
from datetime import datetime
from pathlib import Path

# ─── サードパーティ ───────────────────────────────────────────────────────────
try:
    import gspread
    from google.oauth2.service_account import Credentials as SACredentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    import google.auth
    import pickle
except ImportError as e:
    print(f"[エラー] 必要なライブラリが不足しています: {e}")
    print("以下を実行してインストールしてください:")
    print("  pip install gspread google-auth google-auth-oauthlib qrcode[pil] reportlab Pillow")
    sys.exit(1)

try:
    import qrcode
    from PIL import Image
except ImportError as e:
    print(f"[エラー] qrcode / Pillow が不足しています: {e}")
    print("  pip install qrcode[pil] Pillow")
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, Image as RLImage,
                                    KeepTogether)
    from reportlab.lib.enums import TA_CENTER
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except ImportError as e:
    print(f"[エラー] reportlab が不足しています: {e}")
    print("  pip install reportlab")
    sys.exit(1)

# =============================================================================
# 定数
# =============================================================================

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly',
]

# スプレッドシートのシート名（コード.gs と一致させる）
SHEET_SETTINGS   = '設定'
SHEET_ROSTER     = '名簿とトークン'

# 名簿シートの列インデックス（0 始まり）
COL_EMAIL = 0   # A列: メールアドレス / ゲストラベル
COL_TOKEN = 1   # B列: トークン
COL_URL   = 2   # C列: 投票用URL
COL_VOTED = 3   # D列: 投票済みフラグ

# 設定シートの値列（B列 = インデックス 1、0始まり）
# 行番号は 0 始まり: B1=index0, B2=index1, ...
SETTINGS_ROW_PAGES_URL = 4   # B5（0-indexed row 4）= GitHub Pages URL
SETTINGS_ROW_GAS_URL   = 3   # B4（0-indexed row 3）= GAS URL

# config ファイルのパス（プロジェクトルートに保存）
CONFIG_FILE = Path(__file__).parent / 'votely_config.json'

# OAuth2 トークンキャッシュ
OAUTH_TOKEN_CACHE = Path(__file__).parent / '.votely_oauth_token.pkl'

# =============================================================================
# 設定ファイル（votely_config.json）管理
# =============================================================================

def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
        except Exception:
            return {}
    return {}

def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"[設定保存] {CONFIG_FILE}")

# =============================================================================
# Google 認証
# =============================================================================

def build_gspread_client(credentials_path: str) -> gspread.Client:
    """
    credentials_path がサービスアカウント JSON か OAuth2 クライアントシークレット JSON かを
    自動判別し、gspread クライアントを返す。
    """
    cred_data = json.loads(Path(credentials_path).read_text(encoding='utf-8'))

    # サービスアカウント
    if cred_data.get('type') == 'service_account':
        creds = SACredentials.from_service_account_file(credentials_path, scopes=SCOPES)
        return gspread.authorize(creds)

    # OAuth2 クライアントシークレット
    creds = None
    if OAUTH_TOKEN_CACHE.exists():
        with open(OAUTH_TOKEN_CACHE, 'rb') as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(OAUTH_TOKEN_CACHE, 'wb') as f:
            pickle.dump(creds, f)
        print(f"[認証] OAuth2 トークンを {OAUTH_TOKEN_CACHE} に保存しました。")

    return gspread.authorize(creds)

# =============================================================================
# スプレッドシート操作ユーティリティ
# =============================================================================

def open_spreadsheet(gc: gspread.Client, spreadsheet_id: str) -> gspread.Spreadsheet:
    try:
        return gc.open_by_key(spreadsheet_id)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"[エラー] スプレッドシート ID '{spreadsheet_id}' が見つかりません。")
        print("  → サービスアカウントのメールアドレスに「編集者」として共有されているか確認してください。")
        sys.exit(1)

def get_pages_url(spreadsheet: gspread.Spreadsheet) -> str:
    """設定シートの B5（GitHub Pages URL）を取得する。"""
    ws = spreadsheet.worksheet(SHEET_SETTINGS)
    # B 列 = 列番号 2（1-indexed）、行番号 5 = B5
    url = str(ws.cell(5, 2).value or '').rstrip('/')
    if not url:
        print("[警告] 設定シートの B5（フロントエンドURL）が空です。")
        url = input("  GitHub Pages の URL を入力してください（例: https://your-id.github.io/Votely）: ").strip().rstrip('/')
    return url

def _pad_num(n: int, width: int = 3) -> str:
    return str(n).zfill(width)

def _get_next_guest_number(roster_data: list[list]) -> int:
    """
    名簿データ（ヘッダ行を含む全行）から guest_NNN の最大番号 + 1 を返す。
    1 件もなければ 1 を返す。
    """
    max_n = 0
    for row in roster_data[1:]:  # ヘッダスキップ
        label = str(row[COL_EMAIL]) if len(row) > COL_EMAIL else ''
        m = re.match(r'^guest_(\d+)$', label)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1

# =============================================================================
# コマンド: config — 設定の保存
# =============================================================================

def cmd_config(args):
    cfg = load_config()
    if args.spreadsheet_id:
        cfg['spreadsheet_id'] = args.spreadsheet_id
    if args.credentials:
        cfg['credentials'] = str(Path(args.credentials).resolve())
    save_config(cfg)
    print("[完了] 設定を保存しました。次回から --spreadsheet-id / --credentials を省略できます。")
    for k, v in cfg.items():
        print(f"  {k}: {v}")

# =============================================================================
# コマンド: generate-tokens — 名簿へのトークン一括発行
# =============================================================================

def cmd_generate_tokens(args):
    """
    名簿シートを読み込み、トークンが未発行（B列が空）の行に UUID トークンと
    投票用 URL を生成してバッチ書き込みする。

    GAS の generateTokensOnly() と同等だが、Sheets API の batchUpdate を
    使うため数百行でも高速（6 分制限なし）。
    """
    gc = build_gspread_client(args.credentials)
    ss = open_spreadsheet(gc, args.spreadsheet_id)
    pages_url = get_pages_url(ss)

    roster_ws = ss.worksheet(SHEET_ROSTER)
    print("[読込] 名簿シートを読み込んでいます...")
    all_values = roster_ws.get_all_values()

    if len(all_values) < 2:
        print("[情報] 名簿にデータがありません（ヘッダのみ）。")
        return

    # 更新が必要な行を収集
    updates = []   # list of (row_1indexed, token, url)
    for i, row in enumerate(all_values[1:], start=2):  # 1-indexed, ヘッダ=1
        email = str(row[COL_EMAIL]).strip() if len(row) > COL_EMAIL else ''
        token = str(row[COL_TOKEN]).strip() if len(row) > COL_TOKEN else ''
        if email and not token:
            new_token = str(uuid.uuid4())
            vote_url  = f"{pages_url}/index.html?token={new_token}"
            updates.append((i, new_token, vote_url))

    if not updates:
        print("[情報] 未発行の行はありません。すべてのトークンは発行済みです。")
        return

    print(f"[発行] {len(updates)} 件のトークンを生成しています...")

    # gspread の batch_update でまとめて書き込む（API 呼び出し 1 回で完結）
    cell_updates = []
    for row_idx, token, url in updates:
        cell_updates.append(gspread.Cell(row_idx, COL_TOKEN + 1, token))  # B列
        cell_updates.append(gspread.Cell(row_idx, COL_URL   + 1, url))    # C列

    # BATCH_SIZE ごとに分割して update（1 リクエストに大量セルを詰めると上限超過の恐れ）
    BATCH_SIZE = 500
    for start in range(0, len(cell_updates), BATCH_SIZE):
        batch = cell_updates[start:start + BATCH_SIZE]
        roster_ws.update_cells(batch, value_input_option='USER_ENTERED')
        done = min(start + BATCH_SIZE, len(cell_updates)) // 2  # 2 cells per token
        print(f"  → {done}/{len(updates)} 件書き込み完了")

    print(f"\n[完了] {len(updates)} 件のトークンを発行しました。")
    print(f"  スプレッドシートの「{SHEET_ROSTER}」シートを確認してください。")

# =============================================================================
# コマンド: guest-tickets — 当日参加者チケット発行 + PDF 生成
# =============================================================================

def cmd_guest_tickets(args):
    """
    当日参加者用トークンを N 枚発行してスプレッドシートに追記し、
    QR コード付き A4 印刷用 PDF を生成する。

    GAS の createGuestTickets() + _buildGuestTicketsDialogHtml() と同等。
    """
    num = args.num
    if num <= 0:
        print("[エラー] 発行枚数は 1 以上を指定してください。")
        sys.exit(1)

    gc = build_gspread_client(args.credentials)
    ss = open_spreadsheet(gc, args.spreadsheet_id)
    pages_url = get_pages_url(ss)

    roster_ws = ss.worksheet(SHEET_ROSTER)
    print("[読込] 名簿シートを読み込んでいます...")
    all_values = roster_ws.get_all_values()

    # 次の guest 番号を決定
    start_num = _get_next_guest_number(all_values)
    print(f"[採番] guest_{_pad_num(start_num)} から {num} 枚発行します...")

    # トークン生成
    tickets = []
    rows_to_append = []
    for i in range(num):
        n         = start_num + i
        label     = f"guest_{_pad_num(n)}"
        token     = str(uuid.uuid4())
        vote_url  = f"{pages_url}/index.html?token={token}"
        tickets.append({'label': label, 'token': token, 'url': vote_url})
        rows_to_append.append([label, token, vote_url, False])

    # スプレッドシートに一括追記
    print(f"[書込] {len(rows_to_append)} 行をスプレッドシートに追記しています...")
    # append_rows は末尾にまとめて追記するため append_row ループより格段に速い
    roster_ws.append_rows(rows_to_append, value_input_option='USER_ENTERED')
    print(f"  → 「{SHEET_ROSTER}」シートに追記完了")

    # PDF 生成
    output_path = args.output or f"guest_tickets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    print(f"[PDF] チケット PDF を生成しています → {output_path}")
    _generate_tickets_pdf(
        tickets    = tickets,
        output_path= output_path,
        title      = args.title,
        desc       = args.desc,
    )

    print(f"\n[完了] {num} 枚のチケットを発行しました。")
    print(f"  スプレッドシート: 「{SHEET_ROSTER}」シートに追記済み")
    print(f"  PDF: {output_path}")
    print("  ※ PDF を開いてブラウザまたは PDF リーダーから A4 印刷してください。")

# =============================================================================
# PDF 生成
# =============================================================================

def _qr_image_buffer(url: str, size: int = 200) -> io.BytesIO:
    """URL から QR コード画像を生成して BytesIO で返す。"""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white').convert('RGB')
    img = img.resize((size, size), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


def _register_japanese_font() -> str | None:
    """
    システムにインストールされている日本語フォントを探して reportlab に登録する。
    登録できたフォント名を返す。見つからなければ None。
    """
    # Windows / macOS / Linux のよくある日本語フォントパスを探す
    candidates = [
        # Windows
        r'C:\Windows\Fonts\msgothic.ttc',
        r'C:\Windows\Fonts\meiryo.ttc',
        r'C:\Windows\Fonts\YuGothM.ttc',
        r'C:\Windows\Fonts\NotoSansCJK-Regular.ttc',
        # macOS
        '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc',
        '/System/Library/Fonts/Hiragino Sans GB.ttc',
        '/Library/Fonts/Arial Unicode MS.ttf',
        # Linux (noto-fonts-cjk)
        '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/fonts-japanese-gothic.ttf',
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('JapaneseFont', path))
                return 'JapaneseFont'
            except Exception:
                continue
    return None


def _generate_tickets_pdf(tickets: list[dict], output_path: str,
                           title: str = '', desc: str = ''):
    """
    tickets: [{'label': str, 'token': str, 'url': str}, ...]
    A4 縦、2 列 × N 行のチケットグリッドを生成する。
    """
    # ── フォント設定 ────────────────────────────────────────────────────────
    jp_font = _register_japanese_font()
    base_font   = jp_font if jp_font else 'Helvetica'
    bold_font   = jp_font if jp_font else 'Helvetica-Bold'

    # ── ドキュメント設定 ────────────────────────────────────────────────────
    doc = SimpleDocTemplate(
        output_path,
        pagesize    = A4,
        leftMargin  = 10 * mm,
        rightMargin = 10 * mm,
        topMargin   = 12 * mm,
        bottomMargin= 10 * mm,
    )

    page_w = A4[0] - 20 * mm   # 使用可能幅
    card_w = (page_w - 6 * mm) / 2   # カード幅（2列、中央余白 6mm）
    qr_size= 42 * mm                  # QR コード画像サイズ

    # ── スタイル ─────────────────────────────────────────────────────────────
    title_style = ParagraphStyle(
        'CardTitle',
        fontName  = bold_font,
        fontSize  = 9,
        leading   = 12,
        alignment = TA_CENTER,
        textColor = colors.HexColor('#1a3050'),
    )
    label_style = ParagraphStyle(
        'CardLabel',
        fontName  = bold_font,
        fontSize  = 11,
        leading   = 14,
        alignment = TA_CENTER,
        textColor = colors.HexColor('#333333'),
    )
    desc_style = ParagraphStyle(
        'CardDesc',
        fontName  = base_font,
        fontSize  = 7,
        leading   = 9,
        alignment = TA_CENTER,
        textColor = colors.HexColor('#555555'),
    )
    url_style = ParagraphStyle(
        'CardUrl',
        fontName  = base_font,
        fontSize  = 5.5,
        leading   = 7,
        alignment = TA_CENTER,
        textColor = colors.HexColor('#888888'),
    )

    # ── チケットセルを組み立てる ──────────────────────────────────────────────
    def make_card_cell(ticket: dict) -> list:
        """1 枚分のチケット内容を reportlab のフローアブルリストで返す。"""
        items = []
        if title:
            items.append(Paragraph(title, title_style))
            items.append(Spacer(1, 1 * mm))

        # QR コード
        buf = _qr_image_buffer(ticket['url'], size=300)
        qr_img = RLImage(buf, width=qr_size, height=qr_size)
        items.append(qr_img)
        items.append(Spacer(1, 1 * mm))

        # ゲストラベル
        items.append(Paragraph(f"<b>{ticket['label']}</b>", label_style))

        if desc:
            items.append(Spacer(1, 0.8 * mm))
            items.append(Paragraph(desc, desc_style))

        # URL（小さく表示）
        items.append(Spacer(1, 0.8 * mm))
        # URL が長い場合は途中で改行
        short_url = ticket['url']
        items.append(Paragraph(short_url, url_style))

        return items

    # ── 2 列テーブルにまとめる ────────────────────────────────────────────────
    story = []

    # ページ上部タイトル
    page_title_style = ParagraphStyle(
        'PageTitle',
        fontName  = bold_font,
        fontSize  = 13,
        leading   = 18,
        alignment = TA_CENTER,
        textColor = colors.HexColor('#1a3050'),
        spaceAfter= 4 * mm,
    )
    header_text = title or 'Votely 投票チケット'
    story.append(Paragraph(header_text, page_title_style))

    if desc:
        page_desc_style = ParagraphStyle(
            'PageDesc',
            fontName  = base_font,
            fontSize  = 9,
            leading   = 12,
            alignment = TA_CENTER,
            textColor = colors.HexColor('#555555'),
            spaceAfter= 4 * mm,
        )
        story.append(Paragraph(desc, page_desc_style))

    # チケットを 2 列ずつテーブルに配置
    # カード間に薄い枠線を描画する
    cell_pad   = 4 * mm
    card_inner = card_w - 2 * cell_pad

    # ページ内タイトル/説明を除いたカード本体の高さを固定しない（可変）ため
    # KeepTogether + Table を使って 2 列グリッドにする
    rows = []
    for i in range(0, len(tickets), 2):
        left_cell  = make_card_cell(tickets[i])
        right_cell = make_card_cell(tickets[i + 1]) if i + 1 < len(tickets) else ['']

        t = Table(
            [[left_cell, right_cell]],
            colWidths = [card_w, card_w],
        )
        t.setStyle(TableStyle([
            ('BOX',         (0, 0), (-1, -1), 0.5, colors.HexColor('#aaaaaa')),
            ('INNERGRID',   (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
            ('VALIGN',      (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN',       (0, 0), (-1, -1), 'CENTER'),
            ('TOPPADDING',  (0, 0), (-1, -1), cell_pad),
            ('BOTTOMPADDING',(0,0), (-1,-1),  cell_pad),
            ('LEFTPADDING', (0, 0), (-1, -1), cell_pad),
            ('RIGHTPADDING',(0, 0), (-1, -1), cell_pad),
            ('BACKGROUND',  (0, 0), (-1, -1), colors.HexColor('#fafafa')),
        ]))
        story.append(KeepTogether(t))
        story.append(Spacer(1, 3 * mm))

    doc.build(story)

# =============================================================================
# CLI エントリポイント
# =============================================================================

def resolve_args(args, cfg: dict):
    """
    CLI 引数に値がなければ config ファイルからフォールバックする。
    最終的に必須値が埋まっていなければエラーで終了する。
    """
    if not getattr(args, 'spreadsheet_id', None):
        args.spreadsheet_id = cfg.get('spreadsheet_id', '')
    if not getattr(args, 'credentials', None):
        args.credentials = cfg.get('credentials', '')

    if not args.spreadsheet_id:
        print("[エラー] --spreadsheet-id が必要です。")
        print("  スプレッドシートの URL に含まれる ID を指定してください。")
        print("  例: https://docs.google.com/spreadsheets/d/<ここがID>/edit")
        sys.exit(1)
    if not args.credentials:
        print("[エラー] --credentials が必要です（サービスアカウント or OAuth2 JSON のパス）。")
        sys.exit(1)
    if not Path(args.credentials).exists():
        print(f"[エラー] 認証ファイルが見つかりません: {args.credentials}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description='Votely 管理ツール — トークン発行・PDF チケット生成',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--spreadsheet-id', default='',
                        help='対象スプレッドシートの ID（URL の /d/<ID>/ の部分）')
    common.add_argument('--credentials', default='',
                        help='サービスアカウント or OAuth2 クライアントシークレット JSON のパス')

    sub = parser.add_subparsers(dest='command', required=True)

    # ── config ──────────────────────────────────────────────────────────────
    p_cfg = sub.add_parser('config', parents=[common],
                           help='スプレッドシート ID と認証ファイルパスを保存する')

    # ── generate-tokens ──────────────────────────────────────────────────────
    p_gen = sub.add_parser('generate-tokens', parents=[common],
                           help='名簿シートの未発行行にトークンを一括発行する')

    # ── guest-tickets ─────────────────────────────────────────────────────────
    p_guest = sub.add_parser('guest-tickets', parents=[common],
                             help='当日参加者用トークンを発行して QR コード付き PDF を生成する')
    p_guest.add_argument('num', type=int, help='発行枚数')
    p_guest.add_argument('--title', default='Votely 投票チケット',
                         help='チケットのタイトル（デフォルト: "Votely 投票チケット"）')
    p_guest.add_argument('--desc',  default='',
                         help='チケットの説明文（任意）')
    p_guest.add_argument('--output', default='',
                         help='出力 PDF ファイル名（デフォルト: guest_tickets_YYYYMMDD_HHMMSS.pdf）')

    args = parser.parse_args()
    cfg  = load_config()

    if args.command == 'config':
        cmd_config(args)
        return

    resolve_args(args, cfg)

    if args.command == 'generate-tokens':
        cmd_generate_tokens(args)
    elif args.command == 'guest-tickets':
        cmd_guest_tickets(args)


if __name__ == '__main__':
    main()
