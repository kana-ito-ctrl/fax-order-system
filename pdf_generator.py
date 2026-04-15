"""PDF generator - Sylvia v12 / Haruna final version (Windows compatible)"""
import os
import json
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO

# Font registration - try multiple paths (Windows / Linux / macOS)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATHS = [
    # リポジトリ同梱フォント（Render等のクラウド環境用）
    os.path.join(BASE_DIR, "data", "msgothic.ttc"),
    # Windows
    "C:/Windows/Fonts/msgothic.ttc",
    "C:/Windows/Fonts/meiryo.ttc",
    "C:/Windows/Fonts/YuGothR.ttc",
    "C:/Windows/Fonts/YuGothM.ttc",
    # Linux (Streamlit Cloud)
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    # macOS
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
]
FONT = "JP"
_font_registered = False
for fp in FONT_PATHS:
    if os.path.exists(fp):
        try:
            pdfmetrics.registerFont(TTFont(FONT, fp))
            _font_registered = True
            break
        except Exception:
            continue

if not _font_registered:
    print("WARNING: Japanese font not found. PDF output may have missing characters.")

# Theme Colors
SYL_PRIMARY = colors.HexColor("#8B4513")
SYL_LIGHT = colors.HexColor("#F5E6D3")
HAR_PRIMARY = colors.HexColor("#1B5E8C")
HAR_LIGHT = colors.HexColor("#D6EAF8")

# Data directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")


def load_staff():
    with open(os.path.join(DATA_DIR, "staff.json"), "r", encoding="utf-8") as f:
        return json.load(f)


def safe_str(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s


def tw(text, fs):
    from reportlab.pdfbase.pdfmetrics import stringWidth
    return stringWidth(str(text), FONT, fs)


def draw_clipped(c, text, x, y, max_w, fs, lh=None):
    text = safe_str(text)
    if not text:
        return 1
    text = text.replace('\n', ' ')
    if lh is None:
        lh = fs + 3
    if tw(text, fs) <= max_w:
        c.setFont(FONT, fs)
        c.drawString(x, y, text)
        return 1
    for s in range(1, 4):
        nfs = fs - s
        if nfs < 6:
            break
        if tw(text, nfs) <= max_w:
            c.setFont(FONT, nfs)
            c.drawString(x, y, text)
            return 1
    nfs = max(fs - 2, 6)
    c.setFont(FONT, nfs)
    lines, cur = [], ""
    for ch in text:
        if tw(cur + ch, nfs) > max_w:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    for i, ln in enumerate(lines[:3]):
        c.drawString(x, y - i * lh, ln)
    return min(len(lines), 3)


def draw_header_text(c, text, x, y, max_w, fs):
    text = safe_str(text)
    if not text:
        return 0
    text = text.replace('\n', ' ')
    if tw(text, fs) <= max_w:
        c.setFont(FONT, fs)
        c.drawString(x, y, text)
        return 0
    for nfs in range(fs - 1, 5, -1):
        if tw(text, nfs) <= max_w:
            c.setFont(FONT, nfs)
            c.drawString(x, y, text)
            return 0
    nfs = max(fs - 2, 6)
    c.setFont(FONT, nfs)
    line_h = nfs + 2
    lines, cur = [], ""
    for ch in text:
        if tw(cur + ch, nfs) > max_w:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    for i, ln in enumerate(lines[:2]):
        c.drawString(x, y - i * line_h, ln)
    extra_lines = min(len(lines), 2) - 1
    return extra_lines * line_h


def cell_mid_y(row_top, row_h, font_size):
    return row_top - (row_h / 2) - (font_size * 0.3)


def draw_table_row(c, ml, row_top, col_defs, uw, row_h, fill_color=None):
    row_y = row_top - row_h
    if fill_color:
        c.setFillColor(fill_color)
        c.rect(ml, row_y, uw, row_h, fill=True, stroke=False)
        c.setFillColor(colors.black)
    c.setStrokeColor(colors.HexColor("#999999"))
    c.setLineWidth(0.5)
    c.rect(ml, row_y, uw, row_h, fill=False, stroke=True)
    for i, (xo, _) in enumerate(col_defs):
        if i > 0:
            c.line(ml + xo * mm, row_y, ml + xo * mm, row_y + row_h)
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)


# ========== Sylvia PDF ==========
def gen_sylvia_pdf(order, items, staff_name="伊藤"):
    buf = BytesIO()
    staff_data = load_staff()
    company = staff_data["company"]
    staff = next((s for s in staff_data["staff"] if s["name"] == staff_name), staff_data["staff"][0])

    c = canvas.Canvas(buf, pagesize=landscape(A4))
    w, h = landscape(A4)
    ml = 12 * mm
    mr = 12 * mm
    uw = w - ml - mr
    pad = 4

    # Color bar
    bar_h = 5 * mm
    c.setFillColor(SYL_PRIMARY)
    c.rect(0, h - bar_h, w, bar_h, fill=True, stroke=False)
    c.setFillColor(colors.white)
    c.setFont(FONT, 8)
    c.drawString(ml, h - bar_h + 1.5 * mm, "シルビア様 発注書（出荷指示書）")
    c.setFillColor(colors.black)

    # Header
    y = h - bar_h - 10 * mm
    c.setFont(FONT, 13)
    c.drawString(ml, y, "（株）シルビア 本社")
    rx = w - mr - 75 * mm
    c.setFont(FONT, 11)
    c.drawString(rx, y, "日付　%s" % safe_str(order.get('order_date', '')))
    y -= 16
    c.setFont(FONT, 11)
    c.drawString(ml, y, "ご担当者様")

    addr_text = "\u3012%s %s" % (safe_str(company.get('postal', '')), safe_str(company.get('address', '')))
    extra_drop = draw_header_text(c, addr_text, rx, y, 75 * mm, 10)
    y -= (14 + extra_drop)

    c.setFont(FONT, 10)
    c.drawString(ml, y, "FAX：0587-95-5120　　TEL：0587-95-2725")
    c.drawString(rx, y, "%s　担当：%s" % (safe_str(company.get('name', '')), safe_str(staff.get('name', ''))))
    y -= 20

    # Title
    c.setFont(FONT, 18)
    c.setFillColor(SYL_PRIMARY)
    c.drawString(ml, y, "発　注　書（出荷指示書）")
    c.setFillColor(colors.black)
    y -= 22

    # Order info
    subtotal = sum(i.get('amount', 0) for i in items)
    tax = int(subtotal * 0.08)
    total = subtotal + tax

    c.setFont(FONT, 12)
    c.drawString(ml, y, "オーダーNO：")
    c.setFont(FONT, 13)
    c.drawString(ml + 35 * mm, y, safe_str(order.get('order_no', '')))
    c.setFont(FONT, 12)
    c.drawString(ml + 80 * mm, y, "納品日：")
    c.setFont(FONT, 13)
    c.drawString(ml + 100 * mm, y, safe_str(order.get('delivery_date', '')))
    c.setFont(FONT, 13)
    c.drawString(w - mr - 75 * mm, y, "発注額：\xa5%s（税込）" % "{:,}".format(total))
    y -= 24

    # Destination table (3行レイアウト: 納品先 / 郵便番号+住所 / 電話+FAX)
    dest_row_h = 22
    label_w = 28 * mm
    val_w = uw - label_w

    def _syl_dest_row(label, value, row_top, fs=12):
        c.setFillColor(SYL_LIGHT)
        c.rect(ml, row_top - dest_row_h, label_w, dest_row_h, fill=True, stroke=False)
        c.setFillColor(SYL_PRIMARY)
        c.setFont(FONT, 10)
        c.drawString(ml + pad, cell_mid_y(row_top, dest_row_h, 10), label)
        c.setFillColor(colors.black)
        c.setStrokeColor(colors.HexColor("#999999"))
        c.setLineWidth(0.5)
        c.rect(ml, row_top - dest_row_h, uw, dest_row_h, fill=False, stroke=True)
        c.line(ml + label_w, row_top - dest_row_h, ml + label_w, row_top)
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        draw_clipped(c, value, ml + label_w + pad, cell_mid_y(row_top, dest_row_h, fs), val_w - pad * 2, fs)

    _syl_dest_row("納品先", order.get('delivery_dest', ''), y, 13)
    y -= dest_row_h
    postal = safe_str(order.get('postal', ''))
    address = safe_str(order.get('address', '')).replace('\n', ' ')
    addr_val = ("〒%s %s" % (postal, address)) if postal else address
    _syl_dest_row("住所", addr_val, y, 12)
    y -= dest_row_h
    tel = safe_str(order.get('tel', ''))
    fax = safe_str(order.get('fax', ''))
    tel_val = "TEL: %s" % tel if tel else ""
    if fax:
        tel_val += "　　FAX: %s" % fax
    _syl_dest_row("電話番号", tel_val, y, 11)
    y -= dest_row_h + 8

    # Product table
    uw_mm = uw / mm  # usable width in mm
    prod_cols = [
        (0, 30), (30, 72), (102, 20), (122, 18),
        (140, 18), (158, 22), (180, 22), (202, uw_mm - 202),
    ]
    prod_labels = ["JANコード", "商品名", "規格", "配送荷姿", "1袋単価", "CS単価", "数量(CS)", "金額"]
    prod_hdr_h = 22
    prod_row_h = 28

    c.setFillColor(SYL_PRIMARY)
    c.rect(ml, y - prod_hdr_h, uw, prod_hdr_h, fill=True, stroke=False)
    c.setFillColor(colors.white)
    c.setFont(FONT, 10)
    hdr_ty = cell_mid_y(y, prod_hdr_h, 10)
    for idx, ((xo, wc), label) in enumerate(zip(prod_cols, prod_labels)):
        if idx >= 4:
            c.drawRightString(ml + (xo + wc) * mm - pad, hdr_ty, label)
        else:
            c.drawString(ml + xo * mm + pad, hdr_ty, label)
    c.setFillColor(colors.black)
    y -= prod_hdr_h

    for i, item in enumerate(items):
        bg = colors.HexColor("#FFF8F0") if i % 2 == 0 else None
        draw_table_row(c, ml, y, prod_cols, uw, prod_row_h, fill_color=bg)
        ty = cell_mid_y(y, prod_row_h, 11)

        draw_clipped(c, item.get('jan', ''), ml + prod_cols[0][0] * mm + pad, ty, 29 * mm, 10)
        draw_clipped(c, item.get('name', ''), ml + prod_cols[1][0] * mm + pad, ty, 77 * mm, 11)
        draw_clipped(c, item.get('spec', ''), ml + prod_cols[2][0] * mm + pad, ty, 19 * mm, 9)
        draw_clipped(c, item.get('pack', ''), ml + prod_cols[3][0] * mm + pad, ty, 17 * mm, 9)

        c.setFont(FONT, 11)
        c.drawRightString(ml + (prod_cols[4][0] + prod_cols[4][1]) * mm - pad, ty,
                          safe_str(item.get('unit_price', '')))
        cs_price = item.get('cs_price', '')
        if cs_price != '' and cs_price is not None:
            try:
                c.drawRightString(ml + (prod_cols[5][0] + prod_cols[5][1]) * mm - pad, ty,
                                  "{:,}".format(int(cs_price)))
            except (ValueError, TypeError):
                c.drawRightString(ml + (prod_cols[5][0] + prod_cols[5][1]) * mm - pad, ty, safe_str(cs_price))
        c.setFont(FONT, 12)
        c.drawRightString(ml + (prod_cols[6][0] + prod_cols[6][1]) * mm - pad, ty,
                          safe_str(item.get('quantity', '')))
        amt = item.get('amount', '')
        if amt and amt is not None:
            try:
                c.setFont(FONT, 11)
                c.drawRightString(ml + (prod_cols[7][0] + prod_cols[7][1]) * mm - pad, ty,
                                  "{:,}".format(int(amt)))
            except (ValueError, TypeError):
                pass
        y -= prod_row_h

    for _ in range(max(0, 4 - len(items))):
        draw_table_row(c, ml, y, prod_cols, uw, prod_row_h)
        y -= prod_row_h

    # Totals
    y -= 10
    tx = ml + 155 * mm
    c.setFont(FONT, 12)
    c.drawString(tx, y, "小計")
    c.drawRightString(w - mr - 5 * mm, y, "\xa5%s" % "{:,}".format(subtotal))
    y -= 18
    c.drawString(tx, y, "消費税(8%)")
    c.drawRightString(w - mr - 5 * mm, y, "\xa5%s" % "{:,}".format(tax))
    y -= 18
    c.setFont(FONT, 14)
    c.setFillColor(SYL_PRIMARY)
    c.drawString(tx, y, "合計")
    c.drawRightString(w - mr - 5 * mm, y, "\xa5%s" % "{:,}".format(total))
    c.setFillColor(colors.black)

    syl_remarks = safe_str(order.get('remarks', ''))
    if syl_remarks:
        y -= 30
        c.setFillColor(SYL_PRIMARY)
        c.setFont(FONT, 14)
        c.drawString(ml, y, "※ %s" % syl_remarks)
        c.setFillColor(colors.black)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


# ========== Haruna PDF ==========
def gen_haruna_pdf(order, ddc, staff_name="伊藤"):
    buf = BytesIO()
    staff_data = load_staff()
    company = staff_data["company"]
    staff = next((s for s in staff_data["staff"] if s["name"] == staff_name), staff_data["staff"][0])

    c = canvas.Canvas(buf, pagesize=landscape(A4))
    w, h = landscape(A4)
    ml = 12 * mm
    mr = 12 * mm
    uw = w - ml - mr
    pad = 4

    # Color bar
    bar_h = 5 * mm
    c.setFillColor(HAR_PRIMARY)
    c.rect(0, h - bar_h, w, bar_h, fill=True, stroke=False)
    c.setFillColor(colors.white)
    c.setFont(FONT, 8)
    c.drawString(ml, h - bar_h + 1.5 * mm, "ハルナプロデュース様 発注書")
    c.setFillColor(colors.black)

    # Header
    y = h - bar_h - 10 * mm
    c.setFont(FONT, 13)
    c.drawString(ml, y, "ハルナプロデュース㈱")
    rx = w - mr - 80 * mm
    c.setFont(FONT, 11)
    c.drawString(rx, y, "日付　%s" % safe_str(order.get('order_date', '')))
    y -= 15
    c.setFont(FONT, 11)
    c.drawString(ml, y, "受注ご担当者様")

    addr_text = "\u3012%s %s" % (safe_str(company.get('postal', '')), safe_str(company.get('address', '')))
    extra_drop = draw_header_text(c, addr_text, rx, y, 80 * mm, 10)
    y -= (14 + extra_drop)

    c.setFont(FONT, 10)
    c.drawString(rx, y, safe_str(company.get('name', '')))
    y -= 14
    c.drawString(rx, y, "FAX：%s　TEL：%s" % (safe_str(company.get('fax', '')), safe_str(company.get('tel', ''))))
    y -= 14
    c.drawString(rx, y, "担当：%s（携帯：%s）" % (safe_str(staff.get('name', '')), safe_str(staff.get('phone', ''))))
    y -= 18

    # Title
    c.setFont(FONT, 18)
    c.setFillColor(HAR_PRIMARY)
    c.drawString(ml, y, "発　注　書")
    c.setFillColor(colors.black)
    c.setFont(FONT, 11)
    c.drawString(ml + 80 * mm, y, "下記の通り、注文いたします。")
    y -= 22

    # Order info
    c.setFont(FONT, 12)
    c.drawString(ml, y, "オーダーNO：")
    c.setFont(FONT, 13)
    c.drawString(ml + 35 * mm, y, safe_str(order.get('order_no', '')))
    c.setFont(FONT, 12)
    c.drawString(ml + 85 * mm, y, "納品日：")
    c.setFont(FONT, 13)
    c.drawString(ml + 105 * mm, y, safe_str(order.get('delivery_date', '')))
    y -= 24

    # Destination table (3行レイアウト: 納品先 / 郵便番号+住所 / 電話+FAX+入荷時間)
    dest_row_h = 22
    label_w = 28 * mm
    val_w = uw - label_w

    def _har_dest_row(label, value, row_top, fs=12):
        c.setFillColor(HAR_LIGHT)
        c.rect(ml, row_top - dest_row_h, label_w, dest_row_h, fill=True, stroke=False)
        c.setFillColor(HAR_PRIMARY)
        c.setFont(FONT, 10)
        c.drawString(ml + pad, cell_mid_y(row_top, dest_row_h, 10), label)
        c.setFillColor(colors.black)
        c.setStrokeColor(colors.HexColor("#999999"))
        c.setLineWidth(0.5)
        c.rect(ml, row_top - dest_row_h, uw, dest_row_h, fill=False, stroke=True)
        c.line(ml + label_w, row_top - dest_row_h, ml + label_w, row_top)
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        draw_clipped(c, value, ml + label_w + pad, cell_mid_y(row_top, dest_row_h, fs), val_w - pad * 2, fs)

    _har_dest_row("納品先", order.get('delivery_dest', ''), y, 13)
    y -= dest_row_h
    postal = safe_str(ddc.get('postal', ''))
    address = safe_str(ddc.get('address', '')).replace('\n', ' ')
    addr_val = ("〒%s %s" % (postal, address)) if postal else address
    _har_dest_row("住所", addr_val, y, 12)
    y -= dest_row_h
    tel = safe_str(ddc.get('tel', ''))
    fax = safe_str(ddc.get('fax', ''))
    time_val = safe_str(ddc.get('time', ''))
    tel_parts = []
    if tel:
        tel_parts.append("TEL: %s" % tel)
    if fax:
        tel_parts.append("FAX: %s" % fax)
    if time_val:
        tel_parts.append("入荷時間: %s" % time_val)
    _har_dest_row("電話番号", "　　".join(tel_parts), y, 11)
    y -= dest_row_h + 4

    # Palette info table (Haruna only)
    info_cols = [(0, 50), (50, 60), (110, 88)]
    info_hdr = ["バース予約", "パレット条件", "備考"]
    info_h = 22

    c.setFillColor(HAR_LIGHT)
    c.rect(ml, y - 16, uw, 16, fill=True, stroke=False)
    c.setFillColor(HAR_PRIMARY)
    c.setFont(FONT, 9)
    for (xo, wc), label in zip(info_cols, info_hdr):
        c.drawString(ml + xo * mm + pad, y - 12, label)
    c.setFillColor(colors.black)
    y -= 16

    draw_table_row(c, ml, y, info_cols, uw, info_h)
    ty = cell_mid_y(y, info_h, 10)
    draw_clipped(c, safe_str(ddc.get('berse', '無')), ml + pad, ty, 49 * mm, 10)
    draw_clipped(c, safe_str(ddc.get('palette', '')), ml + 50 * mm + pad, ty, 59 * mm, 10)
    jpr = safe_str(ddc.get('jpr', ''))
    method = safe_str(ddc.get('method', ''))
    if jpr:
        draw_clipped(c, "JPRコード：%s" % jpr, ml + 110 * mm + pad, ty, 87 * mm, 10)
    elif method:
        draw_clipped(c, method, ml + 110 * mm + pad, ty, 87 * mm, 10)
    y -= info_h + 6

    # Product table
    prod_cols = [(0, 90), (90, 30), (120, 26), (146, 26), (172, 26)]
    prod_labels = ["商品名", "規格", "配送荷姿", "発注数量(CS)", "総バラ数(本)"]
    prod_hdr_h = 22
    prod_row_h = 28

    c.setFillColor(HAR_PRIMARY)
    c.rect(ml, y - prod_hdr_h, uw, prod_hdr_h, fill=True, stroke=False)
    c.setFillColor(colors.white)
    c.setFont(FONT, 10)
    hdr_ty = cell_mid_y(y, prod_hdr_h, 10)
    for idx, ((xo, wc), label) in enumerate(zip(prod_cols, prod_labels)):
        if idx >= 3:
            c.drawRightString(ml + (xo + wc) * mm - pad, hdr_ty, label)
        else:
            c.drawString(ml + xo * mm + pad, hdr_ty, label)
    c.setFillColor(colors.black)
    y -= prod_hdr_h

    qty = order['quantity']
    bara = qty * 24
    bg = colors.HexColor("#EBF5FB")
    draw_table_row(c, ml, y, prod_cols, uw, prod_row_h, fill_color=bg)
    ty = cell_mid_y(y, prod_row_h, 11)
    draw_clipped(c, "2Water Ceramide", ml + prod_cols[0][0] * mm + pad, ty, 89 * mm, 12)
    draw_clipped(c, "500ml\xd724本", ml + prod_cols[1][0] * mm + pad, ty, 29 * mm, 10)
    draw_clipped(c, "24本/cs", ml + prod_cols[2][0] * mm + pad, ty, 25 * mm, 10)
    c.setFont(FONT, 13)
    c.drawRightString(ml + (prod_cols[3][0] + prod_cols[3][1]) * mm - pad, ty, str(qty))
    c.drawRightString(ml + (prod_cols[4][0] + prod_cols[4][1]) * mm - pad, ty, "{:,}".format(bara))
    y -= prod_row_h

    for _ in range(3):
        draw_table_row(c, ml, y, prod_cols, uw, prod_row_h)
        y -= prod_row_h

    y -= 10
    c.setFont(FONT, 12)
    c.drawRightString(ml + 172 * mm - pad, y, "小　　計")
    c.setFont(FONT, 13)
    c.drawRightString(ml + 198 * mm - pad, y, "{:,} 本".format(bara))

    har_remarks = safe_str(order.get('remarks', ''))
    if har_remarks:
        y -= 30
        c.setFillColor(HAR_PRIMARY)
        c.setFont(FONT, 14)
        c.drawString(ml, y, "※ %s" % har_remarks)
        c.setFillColor(colors.black)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf
