"""PDF生成モジュール - シルビアv12 / ハルナ最終版"""
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

# フォント登録（Streamlit Cloud用に複数パスを試行）
FONT_PATHS = [
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
]
FONT = "JP"
for fp in FONT_PATHS:
    if os.path.exists(fp):
        try:
            pdfmetrics.registerFont(TTFont(FONT, fp))
            break
        except:
            continue

# ========== Theme Colors ==========
SYL_PRIMARY = colors.HexColor("#8B4513")
SYL_LIGHT = colors.HexColor("#F5E6D3")
HAR_PRIMARY = colors.HexColor("#1B5E8C")
HAR_LIGHT = colors.HexColor("#D6EAF8")

# ========== Load Masters ==========
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

def load_staff():
    with open(os.path.join(DATA_DIR, "staff.json"), "r", encoding="utf-8") as f:
        return json.load(f)

def load_product_master():
    return pd.read_csv(
        os.path.join(DATA_DIR, "product_master.csv"),
        dtype={"JANコード": str}
    )

def load_ddc_master():
    return pd.read_csv(os.path.join(DATA_DIR, "ddc_master.csv"))

# ========== Utility Functions ==========

def safe_str(val):
    """値を安全な文字列に変換。nan/None/NaN -> 空文字"""
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
    """セル内にテキスト描画。収まらない場合: 縮小->改行"""
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
    """ヘッダー用テキスト描画。長い文字列はフォント自動縮小 -> 折り返し。
    戻り値: 追加で下がったピクセル数（折り返しがなければ0）"""
    text = safe_str(text)
    if not text:
        return 0
    text = text.replace('\n', ' ')
    # まず指定サイズで収まるか
    if tw(text, fs) <= max_w:
        c.setFont(FONT, fs)
        c.drawString(x, y, text)
        return 0
    # 段階的にフォントサイズを下げる（最小6pt）
    for nfs in range(fs - 1, 5, -1):
        if tw(text, nfs) <= max_w:
            c.setFont(FONT, nfs)
            c.drawString(x, y, text)
            return 0
    # それでも収まらない場合は折り返し（最大2行）
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
            c.line(ml + xo*mm, row_y, ml + xo*mm, row_y + row_h)
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)


# ========== シルビア v12 ==========
def gen_sylvia_pdf(order, items, staff_name="伊藤"):
    """シルビア向け発注書PDF生成。BytesIOを返す。パレット条件行なし。"""
    buf = BytesIO()
    staff_data = load_staff()
    company = staff_data["company"]
    staff = next((s for s in staff_data["staff"] if s["name"] == staff_name), staff_data["staff"][0])

    c = canvas.Canvas(buf, pagesize=landscape(A4))
    w, h = landscape(A4)
    ml = 12*mm
    mr = 12*mm
    uw = w - ml - mr
    pad = 4

    # カラーバー
    bar_h = 5*mm
    c.setFillColor(SYL_PRIMARY)
    c.rect(0, h - bar_h, w, bar_h, fill=True, stroke=False)
    c.setFillColor(colors.white)
    c.setFont(FONT, 8)
    c.drawString(ml, h - bar_h + 1.5*mm, "シルビア様 発注書（出荷指示書）")
    c.setFillColor(colors.black)

    # ヘッダー
    y = h - bar_h - 10*mm
    c.setFont(FONT, 13)
    c.drawString(ml, y, "（株）シルビア 本社")
    rx = w - mr - 75*mm
    c.setFont(FONT, 11)
    c.drawString(rx, y, "日付　%s" % safe_str(order.get('order_date', '')))
    y -= 16
    c.setFont(FONT, 11)
    c.drawString(ml, y, "ご担当者様")

    # TWO住所（長文対応: 自動縮小/折り返し）
    addr_text = "\u3012%s %s" % (safe_str(company.get('postal', '')), safe_str(company.get('address', '')))
    addr_max_w = 75*mm
    extra_drop = draw_header_text(c, addr_text, rx, y, addr_max_w, 10)
    y -= (14 + extra_drop)

    c.setFont(FONT, 10)
    c.drawString(ml, y, "FAX：0587-95-5120　　TEL：0587-95-2725")
    c.drawString(rx, y, "%s　担当：%s" % (safe_str(company.get('name', '')), safe_str(staff.get('name', ''))))
    y -= 20

    # タイトル
    c.setFont(FONT, 18)
    c.setFillColor(SYL_PRIMARY)
    c.drawString(ml, y, "発　注　書（出荷指示書）")
    c.setFillColor(colors.black)
    y -= 22

    # 注文情報
    subtotal = sum(i.get('amount', 0) for i in items)
    tax = int(subtotal * 0.08)
    total = subtotal + tax

    c.setFont(FONT, 12)
    c.drawString(ml, y, "オーダーNO：")
    c.setFont(FONT, 13)
    c.drawString(ml + 35*mm, y, safe_str(order.get('order_no', '')))
    c.setFont(FONT, 12)
    c.drawString(ml + 80*mm, y, "納品日：")
    c.setFont(FONT, 13)
    c.drawString(ml + 100*mm, y, safe_str(order.get('delivery_date', '')))
    c.setFont(FONT, 13)
    c.drawString(w - mr - 75*mm, y, "発注額：\xa5%s（税込）" % "{:,}".format(total))
    y -= 24

    # 納品先テーブル（パレット条件行なし）
    dest_cols = [(0, 40), (40, 22), (62, 78), (140, 28), (168, 30)]
    dest_labels = ["納品先", "郵便番号", "住所", "電話番号", "FAX番号"]
    hdr_h = 20
    row_h = 34

    c.setFillColor(SYL_LIGHT)
    c.rect(ml, y - hdr_h, uw, hdr_h, fill=True, stroke=False)
    c.setFillColor(SYL_PRIMARY)
    c.setFont(FONT, 10)
    hdr_ty = cell_mid_y(y, hdr_h, 10)
    for (xo, wc), label in zip(dest_cols, dest_labels):
        c.drawString(ml + xo*mm + pad, hdr_ty, label)
    c.setFillColor(colors.black)
    y -= hdr_h

    draw_table_row(c, ml, y, dest_cols, uw, row_h)
    ty = cell_mid_y(y, row_h, 10)
    draw_clipped(c, order.get('delivery_dest', ''), ml + pad, ty, 39*mm, 10, 12)
    draw_clipped(c, order.get('postal', ''), ml + 40*mm + pad, ty, 21*mm, 10, 12)
    draw_clipped(c, safe_str(order.get('address', '')).replace('\n', ' '), ml + 62*mm + pad, ty, 77*mm, 10, 12)
    draw_clipped(c, order.get('tel', ''), ml + 140*mm + pad, ty, 27*mm, 10, 12)
    draw_clipped(c, order.get('fax', ''), ml + 168*mm + pad, ty, 29*mm, 10, 12)
    y -= row_h + 8

    # 商品テーブル
    prod_cols = [
        (0, 30), (30, 78), (108, 18), (126, 16),
        (142, 14), (156, 14), (170, 14), (184, 14),
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
            c.drawRightString(ml + (xo + wc)*mm - pad, hdr_ty, label)
        else:
            c.drawString(ml + xo*mm + pad, hdr_ty, label)
    c.setFillColor(colors.black)
    y -= prod_hdr_h

    for i, item in enumerate(items):
        bg = colors.HexColor("#FFF8F0") if i % 2 == 0 else None
        draw_table_row(c, ml, y, prod_cols, uw, prod_row_h, fill_color=bg)
        ty = cell_mid_y(y, prod_row_h, 11)

        draw_clipped(c, item.get('jan', ''), ml + prod_cols[0][0]*mm + pad, ty, 29*mm, 10)
        draw_clipped(c, item.get('name', ''), ml + prod_cols[1][0]*mm + pad, ty, 77*mm, 11)
        draw_clipped(c, item.get('spec', ''), ml + prod_cols[2][0]*mm + pad, ty, 17*mm, 9)
        draw_clipped(c, item.get('pack', ''), ml + prod_cols[3][0]*mm + pad, ty, 15*mm, 9)

        c.setFont(FONT, 11)
        c.drawRightString(ml + (prod_cols[4][0] + prod_cols[4][1])*mm - pad, ty, safe_str(item.get('unit_price', '')))
        cs_price = item.get('cs_price', '')
        if cs_price != '' and cs_price is not None:
            try:
                c.drawRightString(ml + (prod_cols[5][0] + prod_cols[5][1])*mm - pad, ty, "{:,}".format(int(cs_price)))
            except (ValueError, TypeError):
                c.drawRightString(ml + (prod_cols[5][0] + prod_cols[5][1])*mm - pad, ty, safe_str(cs_price))
        c.setFont(FONT, 12)
        c.drawRightString(ml + (prod_cols[6][0] + prod_cols[6][1])*mm - pad, ty, safe_str(item.get('quantity', '')))
        amt = item.get('amount', '')
        if amt and amt is not None:
            try:
                c.setFont(FONT, 11)
                c.drawRightString(ml + (prod_cols[7][0] + prod_cols[7][1])*mm - pad, ty, "{:,}".format(int(amt)))
            except (ValueError, TypeError):
                pass
        y -= prod_row_h

    for _ in range(max(0, 4 - len(items))):
        draw_table_row(c, ml, y, prod_cols, uw, prod_row_h)
        y -= prod_row_h

    y -= 10
    tx = ml + 155*mm
    c.setFont(FONT, 12)
    c.drawString(tx, y, "小計")
    c.drawRightString(w - mr - 5*mm, y, "\xa5%s" % "{:,}".format(subtotal))
    y -= 18
    c.drawString(tx, y, "消費税(8%)")
    c.drawRightString(w - mr - 5*mm, y, "\xa5%s" % "{:,}".format(tax))
    y -= 18
    c.setFont(FONT, 14)
    c.setFillColor(SYL_PRIMARY)
    c.drawString(tx, y, "合計")
    c.drawRightString(w - mr - 5*mm, y, "\xa5%s" % "{:,}".format(total))

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


# ========== ハルナ 最終版 ==========
def gen_haruna_pdf(order, ddc, staff_name="伊藤"):
    """ハルナ向け発注書PDF生成。BytesIOを返す。パレット条件行はハルナのみ表示。"""
    buf = BytesIO()
    staff_data = load_staff()
    company = staff_data["company"]
    staff = next((s for s in staff_data["staff"] if s["name"] == staff_name), staff_data["staff"][0])

    c = canvas.Canvas(buf, pagesize=landscape(A4))
    w, h = landscape(A4)
    ml = 12*mm
    mr = 12*mm
    uw = w - ml - mr
    pad = 4

    # カラーバー
    bar_h = 5*mm
    c.setFillColor(HAR_PRIMARY)
    c.rect(0, h - bar_h, w, bar_h, fill=True, stroke=False)
    c.setFillColor(colors.white)
    c.setFont(FONT, 8)
    c.drawString(ml, h - bar_h + 1.5*mm, "ハルナプロデュース様 発注書")
    c.setFillColor(colors.black)

    # ヘッダー
    y = h - bar_h - 10*mm
    c.setFont(FONT, 13)
    c.drawString(ml, y, "ハルナプロデュース㈱")
    rx = w - mr - 80*mm
    c.setFont(FONT, 11)
    c.drawString(rx, y, "日付　%s" % safe_str(order.get('order_date', '')))
    y -= 15
    c.setFont(FONT, 11)
    c.drawString(ml, y, "受注ご担当者様")

    # TWO住所（長文対応: 自動縮小/折り返し）
    addr_text = "\u3012%s %s" % (safe_str(company.get('postal', '')), safe_str(company.get('address', '')))
    addr_max_w = 80*mm
    extra_drop = draw_header_text(c, addr_text, rx, y, addr_max_w, 10)
    y -= (14 + extra_drop)

    c.setFont(FONT, 10)
    c.drawString(rx, y, safe_str(company.get('name', '')))
    y -= 14
    c.drawString(rx, y, "FAX：%s　TEL：%s" % (safe_str(company.get('fax', '')), safe_str(company.get('tel', ''))))
    y -= 14
    c.drawString(rx, y, "担当：%s（携帯：%s）" % (safe_str(staff.get('name', '')), safe_str(staff.get('phone', ''))))
    y -= 18

    # タイトル
    c.setFont(FONT, 18)
    c.setFillColor(HAR_PRIMARY)
    c.drawString(ml, y, "発　注　書")
    c.setFillColor(colors.black)
    c.setFont(FONT, 11)
    c.drawString(ml + 80*mm, y, "下記の通り、注文いたします。")
    y -= 22

    # 注文情報
    c.setFont(FONT, 12)
    c.drawString(ml, y, "オーダーNO：")
    c.setFont(FONT, 13)
    c.drawString(ml + 35*mm, y, safe_str(order.get('order_no', '')))
    c.setFont(FONT, 12)
    c.drawString(ml + 85*mm, y, "納品日：")
    c.setFont(FONT, 13)
    c.drawString(ml + 105*mm, y, safe_str(order.get('delivery_date', '')))
    remarks = safe_str(order.get('remarks', ''))
    if remarks:
        c.setFont(FONT, 11)
        c.drawString(w - mr - 70*mm, y, "備考：%s" % remarks)
    y -= 24

    # 納品先テーブル
    dest_cols = [(0, 42), (42, 22), (64, 72), (136, 24), (160, 22), (182, 16)]
    dest_labels = ["納品先", "郵便番号", "住所", "電話番号", "FAX番号", "入荷時間"]
    hdr_h = 20
    row_h = 32

    c.setFillColor(HAR_LIGHT)
    c.rect(ml, y - hdr_h, uw, hdr_h, fill=True, stroke=False)
    c.setFillColor(HAR_PRIMARY)
    c.setFont(FONT, 10)
    hdr_ty = cell_mid_y(y, hdr_h, 10)
    for (xo, wc), label in zip(dest_cols, dest_labels):
        c.drawString(ml + xo*mm + pad, hdr_ty, label)
    c.setFillColor(colors.black)
    y -= hdr_h

    draw_table_row(c, ml, y, dest_cols, uw, row_h)
    ty = cell_mid_y(y, row_h, 10)
    draw_clipped(c, order.get('delivery_dest', ''), ml + pad, ty, 41*mm, 10, 12)
    draw_clipped(c, ddc.get('postal', ''), ml + 42*mm + pad, ty, 21*mm, 10)
    draw_clipped(c, safe_str(ddc.get('address', '')).replace('\n', ' '), ml + 64*mm + pad, ty, 71*mm, 10, 12)
    draw_clipped(c, ddc.get('tel', ''), ml + 136*mm + pad, ty, 23*mm, 10)
    draw_clipped(c, ddc.get('fax', ''), ml + 160*mm + pad, ty, 21*mm, 10)
    draw_clipped(c, ddc.get('time', ''), ml + 182*mm + pad, ty, 15*mm, 9)
    y -= row_h + 4

    # パレット条件テーブル（ハルナのみ表示）
    info_cols = [(0, 50), (50, 60), (110, 88)]
    info_hdr = ["バース予約", "パレット条件", "備考"]
    info_h = 22

    c.setFillColor(HAR_LIGHT)
    c.rect(ml, y - 16, uw, 16, fill=True, stroke=False)
    c.setFillColor(HAR_PRIMARY)
    c.setFont(FONT, 9)
    for (xo, wc), label in zip(info_cols, info_hdr):
        c.drawString(ml + xo*mm + pad, y - 12, label)
    c.setFillColor(colors.black)
    y -= 16

    draw_table_row(c, ml, y, info_cols, uw, info_h)
    ty = cell_mid_y(y, info_h, 10)
    draw_clipped(c, safe_str(ddc.get('berse', '無')), ml + pad, ty, 49*mm, 10)
    draw_clipped(c, safe_str(ddc.get('palette', '')), ml + 50*mm + pad, ty, 59*mm, 10)
    jpr = safe_str(ddc.get('jpr', ''))
    method = safe_str(ddc.get('method', ''))
    if jpr:
        draw_clipped(c, "JPRコード：%s" % jpr, ml + 110*mm + pad, ty, 87*mm, 10)
    elif method:
        draw_clipped(c, method, ml + 110*mm + pad, ty, 87*mm, 10)
    y -= info_h + 6

    # 商品テーブル
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
            c.drawRightString(ml + (xo + wc)*mm - pad, hdr_ty, label)
        else:
            c.drawString(ml + xo*mm + pad, hdr_ty, label)
    c.setFillColor(colors.black)
    y -= prod_hdr_h

    qty = order['quantity']
    bara = qty * 24
    bg = colors.HexColor("#EBF5FB")
    draw_table_row(c, ml, y, prod_cols, uw, prod_row_h, fill_color=bg)
    ty = cell_mid_y(y, prod_row_h, 11)
    draw_clipped(c, "2Water Ceramide", ml + prod_cols[0][0]*mm + pad, ty, 89*mm, 12)
    draw_clipped(c, "500ml\xd724本", ml + prod_cols[1][0]*mm + pad, ty, 29*mm, 10)
    draw_clipped(c, "24本/cs", ml + prod_cols[2][0]*mm + pad, ty, 25*mm, 10)
    c.setFont(FONT, 13)
    c.drawRightString(ml + (prod_cols[3][0] + prod_cols[3][1])*mm - pad, ty, str(qty))
    c.drawRightString(ml + (prod_cols[4][0] + prod_cols[4][1])*mm - pad, ty, "{:,}".format(bara))
    y -= prod_row_h

    for _ in range(3):
        draw_table_row(c, ml, y, prod_cols, uw, prod_row_h)
        y -= prod_row_h

    y -= 10
    c.setFont(FONT, 12)
    c.drawRightString(ml + 172*mm - pad, y, "小　　計")
    c.setFont(FONT, 13)
    c.drawRightString(ml + 198*mm - pad, y, "{:,} 本".format(bara))

    c.showPage()
    c.save()
    buf.seek(0)
    return buf
