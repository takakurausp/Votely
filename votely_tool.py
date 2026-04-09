#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
votely_tool.py — Votely トークン生成ツール（完全ローカル版）

Google API・認証設定は一切不要。
メールアドレス CSV を読み込み、トークン付き CSV を出力する。
当日参加者チケットは CSV + QR コード付き PDF を生成する。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
必要ライブラリ（初回のみ）:
  pip install -r requirements.txt

使い方:
  # メールアドレス CSV → トークン付き CSV
  python votely_tool.py generate-tokens emails.csv \\
      --pages-url https://your-id.github.io/Votely \\
      --output tokens.csv

  # 当日参加者チケット（CSV + PDF）
  python votely_tool.py guest-tickets 30 \\
      --pages-url https://your-id.github.io/Votely \\
      --title "懇親会投票チケット" \\
      --desc "このQRコードを読み取って投票してください" \\
      --output-csv guest_tokens.csv \\
      --output-pdf guest_tickets.pdf
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import argparse
import csv
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
    import qrcode
    from PIL import Image
except ImportError as e:
    print(f"[エラー] {e}")
    print("  pip install qrcode[pil] Pillow")
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, Image as RLImage,
                                    KeepTogether)
    from reportlab.lib.enums import TA_CENTER
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except ImportError as e:
    print(f"[エラー] {e}")
    print("  pip install reportlab")
    sys.exit(1)

# =============================================================================
# 定数・設定
# =============================================================================

CONFIG_FILE = Path(__file__).parent / 'votely_config.json'


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
        except Exception:
            return {}
    return {}


def save_config(cfg: dict):
    CONFIG_FILE.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8'
    )

# =============================================================================
# コア機能
# =============================================================================

def generate_tokens(input_csv: str, pages_url: str, output_csv: str) -> int:
    """
    メールアドレス CSV を読み込み、各行にトークンと投票 URL を付与して
    output_csv に書き出す。

    入力 CSV: 1 列目がメールアドレス（ヘッダ行は自動スキップ）
    出力 CSV: メールアドレス, トークン, 投票用URL, 投票済みフラグ
    """
    base = pages_url.rstrip('/')

    # ── 入力 CSV 読み込み ─────────────────────────────────────────────────────
    emails: list[str] = []
    with open(input_csv, encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f)
        for i, row in enumerate(reader):
            if not row:
                continue
            email = row[0].strip()
            # 1 行目がヘッダ文字列なら無視
            if i == 0 and email.lower() in ('メールアドレス', 'email', 'mail', 'e-mail'):
                continue
            if email:
                emails.append(email)

    if not emails:
        print('[エラー] メールアドレスが 1 件も見つかりませんでした。')
        print(f'  確認: {input_csv} の 1 列目にメールアドレスが入っていますか？')
        return 0

    print(f'[読込] {len(emails)} 件のメールアドレスを読み込みました。')
    print(f'[発行] トークンを生成しています...')

    # ── トークン生成 ──────────────────────────────────────────────────────────
    rows = []
    for email in emails:
        token   = str(uuid.uuid4())
        url     = f'{base}/index.html?token={token}'
        rows.append([email, token, url, 'FALSE'])

    # ── 出力 CSV 書き込み（BOM 付き UTF-8 → Excel で文字化けしない） ──────────
    with open(output_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['メールアドレス', 'トークン', '投票用URL', '投票済みフラグ'])
        writer.writerows(rows)

    print(f'[完了] {len(rows)} 件を出力しました → {output_csv}')
    return len(rows)


def generate_guest_tickets(num: int, pages_url: str, title: str, desc: str,
                            output_csv: str, output_pdf: str) -> list[dict]:
    """
    当日参加者用トークンを num 枚発行し、CSV と PDF を生成する。
    既存の output_csv が存在する場合は続き番号から採番して追記する。
    """
    base      = pages_url.rstrip('/')
    start_num = _next_guest_num(output_csv)

    print(f'[発行] guest_{str(start_num).zfill(3)} から {num} 枚生成します...')

    tickets: list[dict] = []
    rows: list[list]    = []
    for i in range(num):
        n     = start_num + i
        label = f'guest_{str(n).zfill(3)}'
        token = str(uuid.uuid4())
        url   = f'{base}/index.html?token={token}'
        tickets.append({'label': label, 'token': token, 'url': url})
        rows.append([label, token, url, 'FALSE'])

    # ── CSV 出力（既存ファイルがあれば追記） ──────────────────────────────────
    file_exists = Path(output_csv).exists()
    with open(output_csv, 'a' if file_exists else 'w',
              encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['ゲストラベル', 'トークン', '投票用URL', '投票済みフラグ'])
        writer.writerows(rows)
    print(f'[CSV] {len(rows)} 件を {"追記" if file_exists else "出力"} しました → {output_csv}')

    # ── PDF 出力 ──────────────────────────────────────────────────────────────
    _generate_tickets_pdf(tickets, output_pdf, title, desc)
    print(f'[PDF] チケット PDF を生成しました → {output_pdf}')

    return tickets


def _next_guest_num(csv_path: str) -> int:
    """既存 CSV の guest_NNN ラベルから次の番号を返す。なければ 1。"""
    if not Path(csv_path).exists():
        return 1
    max_n = 0
    try:
        with open(csv_path, encoding='utf-8-sig', newline='') as f:
            for row in csv.reader(f):
                if row:
                    m = re.match(r'^guest_(\d+)$', str(row[0]).strip())
                    if m:
                        max_n = max(max_n, int(m.group(1)))
    except Exception:
        pass
    return max_n + 1 if max_n > 0 else 1

# =============================================================================
# PDF 生成
# =============================================================================

def _register_japanese_font() -> str | None:
    candidates = [
        r'C:\Windows\Fonts\msgothic.ttc',
        r'C:\Windows\Fonts\meiryo.ttc',
        r'C:\Windows\Fonts\YuGothM.ttc',
        '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc',
        '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('JapaneseFont', path))
                return 'JapaneseFont'
            except Exception:
                continue
    return None


def _qr_buffer(url: str, px: int = 300) -> io.BytesIO:
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M,
                       box_size=10, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white').convert('RGB')
    img = img.resize((px, px), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


def _generate_tickets_pdf(tickets: list[dict], output_path: str,
                           title: str = '', desc: str = ''):
    jp = _register_japanese_font()
    bf = jp or 'Helvetica'
    bfb = jp or 'Helvetica-Bold'

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=12*mm, bottomMargin=10*mm,
    )

    page_w  = A4[0] - 20*mm
    card_w  = (page_w - 6*mm) / 2
    qr_size = 42*mm
    pad     = 4*mm

    def sty(name, font, size, color='#333333', align=TA_CENTER, leading=None):
        return ParagraphStyle(name, fontName=font, fontSize=size,
                              leading=leading or size * 1.3,
                              alignment=align,
                              textColor=colors.HexColor(color))

    s_title = sty('T', bfb, 9,  '#1a3050')
    s_label = sty('L', bfb, 11, '#333333')
    s_desc  = sty('D', bf,  7,  '#555555')
    s_url   = sty('U', bf,  5,  '#999999')
    s_page  = sty('P', bfb, 13, '#1a3050')
    s_pdesc = sty('PD', bf, 9,  '#555555')

    def card(t: dict) -> list:
        items = []
        if title:
            items += [Paragraph(title, s_title), Spacer(1, 1*mm)]
        items.append(RLImage(_qr_buffer(t['url']), qr_size, qr_size))
        items += [Spacer(1, 1*mm), Paragraph(f"<b>{t['label']}</b>", s_label)]
        if desc:
            items += [Spacer(1, 0.8*mm), Paragraph(desc, s_desc)]
        items += [Spacer(1, 0.8*mm), Paragraph(t['url'], s_url)]
        return items

    story = [Paragraph(title or 'Votely 投票チケット', s_page)]
    if desc:
        story.append(Paragraph(desc, s_pdesc))
    story.append(Spacer(1, 4*mm))

    for i in range(0, len(tickets), 2):
        left  = card(tickets[i])
        right = card(tickets[i+1]) if i+1 < len(tickets) else ['']
        t = Table([[left, right]], colWidths=[card_w, card_w])
        t.setStyle(TableStyle([
            ('BOX',          (0,0),(-1,-1), 0.5, colors.HexColor('#aaaaaa')),
            ('INNERGRID',    (0,0),(-1,-1), 0.5, colors.HexColor('#cccccc')),
            ('VALIGN',       (0,0),(-1,-1), 'MIDDLE'),
            ('ALIGN',        (0,0),(-1,-1), 'CENTER'),
            ('TOPPADDING',   (0,0),(-1,-1), pad),
            ('BOTTOMPADDING',(0,0),(-1,-1), pad),
            ('LEFTPADDING',  (0,0),(-1,-1), pad),
            ('RIGHTPADDING', (0,0),(-1,-1), pad),
            ('BACKGROUND',   (0,0),(-1,-1), colors.HexColor('#fafafa')),
        ]))
        story += [KeepTogether(t), Spacer(1, 3*mm)]

    doc.build(story)

# =============================================================================
# CLI
# =============================================================================

def main():
    cfg = load_config()

    parser = argparse.ArgumentParser(
        description='Votely トークン生成ツール（完全ローカル版）')

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--pages-url', default='',
                        help='投票ページの URL（例: https://your-id.github.io/Votely）')

    sub = parser.add_subparsers(dest='command', required=True)

    # generate-tokens
    p1 = sub.add_parser('generate-tokens', parents=[common],
                        help='メールアドレス CSV にトークンを付与して出力')
    p1.add_argument('input', help='入力 CSV（メールアドレス列）')
    p1.add_argument('--output', default='tokens.csv', help='出力 CSV（デフォルト: tokens.csv）')

    # guest-tickets
    p2 = sub.add_parser('guest-tickets', parents=[common],
                        help='当日参加者チケットを発行して CSV + PDF を生成')
    p2.add_argument('num', type=int, help='発行枚数')
    p2.add_argument('--title', default='Votely 投票チケット')
    p2.add_argument('--desc',  default='')
    p2.add_argument('--output-csv', default='guest_tokens.csv')
    p2.add_argument('--output-pdf', default='')

    args = parser.parse_args()

    # pages_url: 引数 → config の順で解決
    pages_url = args.pages_url or cfg.get('pages_url', '')
    if not pages_url:
        pages_url = input('投票ページ URL を入力してください: ').strip()
    pages_url = pages_url.rstrip('/')

    if args.command == 'generate-tokens':
        generate_tokens(args.input, pages_url, args.output)

    elif args.command == 'guest-tickets':
        pdf_path = args.output_pdf or \
            f"guest_tickets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        generate_guest_tickets(
            args.num, pages_url, args.title, args.desc,
            args.output_csv, pdf_path,
        )


if __name__ == '__main__':
    main()
