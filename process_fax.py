"""FAX Order Processing Script - Main entry point

Usage:
    python process_fax.py <pdf_path> [--staff NAME] [--remarks TEXT]
    python process_fax.py <pdf_path1> <pdf_path2> ...  (multiple files)

Output:
    - CSV file with order data (output/ directory)
    - Sylvia order PDF (if applicable)
    - Haruna order PDF (if applicable)
"""
import os
import sys
import csv
import argparse

# Windows terminal UTF-8 output
if sys.stdout.encoding and sys.stdout.encoding.lower() in ('cp932', 'shift_jis', 'mbcs'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from datetime import date

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ocr_module import process_fax_pdf, load_staff
from pdf_generator import gen_sylvia_pdf, gen_haruna_pdf


OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


def ensure_output_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def results_to_csv(results, pdf_name):
    """Write OCR+matching results to CSV"""
    ensure_output_dir()
    csv_name = os.path.splitext(pdf_name)[0] + "_result.csv"
    csv_path = os.path.join(OUTPUT_DIR, csv_name)

    rows = []
    for page_result in results:
        if "error" in page_result:
            continue
        ocr = page_result["ocr_raw"]
        ddc = page_result["ddc_match"]
        candidates = ddc.get("candidates", [])

        # 候補列（最大3件）
        def cand_cell(idx):
            if idx < len(candidates):
                c = candidates[idx]
                return f"{c['name']} ({c['score']:.0%})"
            return ""

        low_conf = ddc.get("low_confidence", False)
        match_status = "OK" if ddc.get("matched") else "NG"
        if ddc.get("matched") and low_conf:
            match_status = "要確認"

        for item in page_result["matched_items"]:
            rows.append({
                "ページ": page_result["page"],
                "オーダーNO": ocr.get("order_no", ""),
                "納品日": ocr.get("delivery_date", ""),
                "発注元": ocr.get("sender", ""),
                "納品先(OCR)": ocr.get("delivery_dest", ""),
                "納品先(マスタ)": ddc.get("name", "") if ddc.get("matched") else "",
                "納品先住所": ddc.get("address", "") if ddc.get("matched") else "",
                "納品先TEL": ddc.get("tel", "") if ddc.get("matched") else "",
                "DDCマッチ": match_status,
                "DDC候補1": cand_cell(0),
                "DDC候補2": cand_cell(1),
                "DDC候補3": cand_cell(2),
                "商品名(OCR)": item.get("ocr_name", ""),
                "商品名(マスタ)": item.get("master_name", ""),
                "JANコード": item.get("jan", ""),
                "商品コード": item.get("code", ""),
                "規格": item.get("spec", ""),
                "配送荷姿": item.get("pack", ""),
                "CS単価": item.get("cs_price", ""),
                "数量(CS)": item.get("quantity", ""),
                "金額": item.get("amount", ""),
                "出力先": item.get("output_dest", ""),
                "商品マッチ": "OK" if item.get("matched") else "NG",
                "備考": ocr.get("notes", ""),
            })

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        if rows:
            writer = csv.DictWriter(f, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(rows)
        else:
            f.write("データなし\n")

    return csv_path


# ── ネクストエンジン汎用受注CSV ──

NE_HEADERS = [
    "店舗伝票番号", "受注日", "受注郵便番号", "受注住所１", "受注住所２",
    "受注名", "受注名カナ", "受注電話番号", "受注メールアドレス",
    "発送郵便番号", "発送先住所１", "発送先住所２", "発送先名", "発送先カナ",
    "発送電話番号", "支払方法", "発送方法", "商品計", "税金", "発送料",
    "手数料", "ポイント", "その他費用", "合計金額", "ギフトフラグ",
    "時間帯指定", "日付指定", "作業者欄", "備考",
    "商品名", "商品コード", "商品価格", "受注数量", "商品オプション",
    "出荷済フラグ", "顧客区分", "顧客コード", "消費税率（%）",
    "のし", "ラッピング", "メッセージ",
]

# TWO会社情報（受注側の固定値）
_TWO_POSTAL = "1500012"
_TWO_ADDRESS = "東京都品川区西五反田2-24-4 THE CROSS GOTANDA 5階"
_TWO_NAME = "株式会社ＴＷＯ"
_TWO_TEL = "03-6839-0010"

def _get_irisuu(item):
    """商品マッチ結果から入数を取得する。case_quantity があればそれを使い、なければ1。"""
    cq = item.get("case_quantity", 0)
    try:
        cq = int(cq or 0)
    except (ValueError, TypeError):
        cq = 0
    return cq if cq > 0 else 1


def results_to_ne_csv(results, pdf_name):
    """OCR結果をネクストエンジン汎用受注CSV形式に変換して出力する。

    NE仕様:
    - 同一伝票番号の複数商品は行を分けて記載（ヘッダー部分は同じ値を繰り返す）
    - エンコーディング: Shift-JIS (cp932)
    - 店舗伝票番号 = オーダーNO
    - 発送先 = 納品先(DDCマッチ結果)
    - 受注側 = TWO固定情報
    """
    ensure_output_dir()
    csv_name = os.path.splitext(pdf_name)[0] + "_NE受注.csv"
    csv_path = os.path.join(OUTPUT_DIR, csv_name)

    rows = []
    for page_result in results:
        if "error" in page_result:
            continue
        ocr = page_result["ocr_raw"]
        ddc = page_result.get("ddc_match", {})

        order_no = ocr.get("order_no", "")
        delivery_date = ocr.get("delivery_date", "")
        # 受注日: YYYY/MM/DD 形式に変換
        order_date_str = ""
        if delivery_date:
            order_date_str = str(date.today()).replace("-", "/") + " 00:00:00"

        # 納品先情報
        dest_name = ddc.get("name", ocr.get("delivery_dest", ""))
        dest_postal = (ddc.get("postal", "") or "").replace("-", "")
        dest_address = ddc.get("address", "") or ""
        dest_tel = ddc.get("tel", "") or ""

        # 日付指定 = 納品日
        date_spec = ""
        if delivery_date:
            date_spec = delivery_date.replace("-", "/")

        # 商品行を生成
        matched_items = page_result.get("matched_items", [])
        if not matched_items:
            continue

        # 商品計・合計を計算（バラ単価 × バラ数量）
        total_amount = 0
        for item in matched_items:
            if item.get("matched"):
                irisuu = _get_irisuu(item)
                cs_p = item.get("cs_price", 0)
                qty = item.get("quantity", 0)
                try:
                    unit_p = int(float(cs_p or 0)) // irisuu if irisuu > 0 else int(float(cs_p or 0))
                    q = int(float(qty or 0)) * irisuu
                    total_amount += unit_p * q
                except (ValueError, TypeError):
                    pass

        tax = int(total_amount * 0.1)

        for item in matched_items:
            if not item.get("matched"):
                continue

            qty_cs = item.get("quantity", 0)
            cs_price = item.get("cs_price", 0)
            try:
                price = int(float(cs_price or 0))
            except (ValueError, TypeError):
                price = 0
            try:
                qty_cs_int = int(float(qty_cs or 0))
            except (ValueError, TypeError):
                qty_cs_int = 0

            jan = item.get("jan", "")
            # CS数 → バラ数に換算（case_quantityから動的取得）
            irisuu = _get_irisuu(item)
            quantity_bara = qty_cs_int * irisuu
            # バラ単価 = CS単価 / 入数
            unit_price_bara = int(price / irisuu) if irisuu > 0 else price

            product_name = item.get("master_name", item.get("ocr_name", ""))
            product_code = item.get("code", "")
            if not product_code:
                raise ValueError(f"商品コード未登録: {product_name} (JAN:{jan})")
            # 備考 = 出荷元（シルビア/ハルナ/自社倉庫）
            output_dest = item.get("output_dest", "")
            notes = output_dest

            row = {
                "店舗伝票番号": order_no,
                "受注日": order_date_str,
                "受注郵便番号": dest_postal,
                "受注住所１": dest_address,
                "受注住所２": "",
                "受注名": dest_name,
                "受注名カナ": "",
                "受注電話番号": dest_tel,
                "受注メールアドレス": "",
                "発送郵便番号": dest_postal,
                "発送先住所１": dest_address,
                "発送先住所２": "",
                "発送先名": dest_name,
                "発送先カナ": "",
                "発送電話番号": dest_tel,
                "支払方法": "請求書払い",
                "発送方法": "",
                "商品計": total_amount,
                "税金": tax,
                "発送料": 0,
                "手数料": 0,
                "ポイント": 0,
                "その他費用": 0,
                "合計金額": total_amount + tax,
                "ギフトフラグ": 0,
                "時間帯指定": "時間帯指定[午前中]",
                "日付指定": date_spec,
                "作業者欄": "",
                "備考": notes,
                "商品名": product_name,
                "商品コード": product_code,
                "商品価格": unit_price_bara,
                "受注数量": quantity_bara,
                "商品オプション": "",
                "出荷済フラグ": "",
                "顧客区分": 9,
                "顧客コード": "",
                "消費税率（%）": "",
                "のし": "",
                "ラッピング": "",
                "メッセージ": "",
            }
            rows.append(row)

    with open(csv_path, "w", newline="", encoding="cp932", errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=NE_HEADERS)
        writer.writeheader()
        writer.writerows(rows)

    return csv_path


# ── クーラ（COOLA/ロジザード）用CSV ──

# 卸先マスタCSVパス
_OROSHISAKI_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "oroshisaki_master.csv")

# クーラCSVヘッダー（83カラム）
COOLA_HEADERS = [
    "ショップコード", "ショップ名", "表示受注番号", "行no",
    "配送先名", "配送先郵便番号", "配送先都道府県", "配送先住所１", "配送先住所２", "配送先住所３",
    "配送先電話番号", "配送先fax番号", "配送先eメール", "配送先法人名", "配送先所属部署・役職",
    "出荷備考", "出荷日", "配送会社id", "配送指定日", "配送時間帯",
    "顧客id", "発注日", "税種区分", "送料", "代引手数料",
    "消費税", "ポイント使用額", "ポイントステータス", "獲得ポイント", "使用ポイント",
    "合計ポイント", "請求額合計", "決済区分id", "決済方法", "配送方法",
    "領収書flg", "領収書宛名", "ギフトメッセージ", "包装1", "包装2",
    "包装料1", "包装料2", "のし", "荷送人名指定flg",
    "購入者氏名", "購入者ふりがな", "購入者郵便番号", "購入者都道府県",
    "購入者住所１", "購入者住所２", "購入者住所３",
    "購入者法人名", "購入者電話番号", "購入者eメール", "コメント",
    "品番", "品名", "形式/型番", "色id", "色名", "サイズid", "サイズ名",
    "項目選択肢", "受注数", "単価", "行備考", "出荷行備考",
    "送状備考1", "送状備考2", "出荷ステータス", "予備1", "ギフトフラグ",
    "発注時間", "優先フラグ", "クーポン金額", "納品書コメント",
    "NE配送会社ID", "認証番号", "8％消費税", "10％消費税", "8％金額合計", "10％金額合計", "消費税率",
]

_oroshisaki_cache = None


def _load_oroshisaki_master():
    """卸先マスタCSVを読み込み、卸先名→レコードのdictを返す"""
    global _oroshisaki_cache
    if _oroshisaki_cache is not None:
        return _oroshisaki_cache
    _oroshisaki_cache = {}
    if not os.path.exists(_OROSHISAKI_CSV):
        return _oroshisaki_cache
    import unicodedata
    for enc in ['shift-jis', 'cp932', 'utf-8-sig', 'utf-8']:
        try:
            with open(_OROSHISAKI_CSV, encoding=enc) as f:
                reader = csv.DictReader(f)
                for row in reader:
                    name = row.get("卸先名", "").strip()
                    if name:
                        norm = unicodedata.normalize('NFKC', name)
                        _oroshisaki_cache[name] = row
                        _oroshisaki_cache[norm] = row
            break
        except (UnicodeDecodeError, KeyError):
            continue
    return _oroshisaki_cache


def _split_address(address):
    """住所を都道府県/住所1(50文字)/住所2(15文字)/住所3に分割"""
    import re
    pref = ""
    rest = address
    m = re.match(r'(東京都|北海道|(?:京都|大阪)府|.{2,3}県)', address)
    if m:
        pref = m.group(1)
        rest = address[len(pref):]
    addr1 = rest[:50]
    addr2 = rest[50:65] if len(rest) > 50 else ""
    addr3 = rest[65:] if len(rest) > 65 else ""
    return pref, addr1, addr2, addr3


# 都道府県 → BtoB最安配送会社のマッピング
# 福山通運が安い地域は福山、それ以外はヤマト
_PREF_TO_BTOB_CARRIER_120 = {}
_PREF_TO_BTOB_CARRIER_140 = {}
# 120サイズ比較: 福山140 vs ヤマト120 → 福山が安い地域
for p in ["青森", "岩手", "秋田", "山形", "宮城", "福島",  # 北東北・南東北
          "茨城", "栃木", "群馬", "埼玉", "千葉", "東京", "神奈川", "山梨",  # 関東
          "新潟", "長野", "富山", "石川", "福井",  # 信越・北陸
          "静岡", "愛知", "岐阜", "三重",  # 東海
          "滋賀", "京都", "大阪", "兵庫", "奈良", "和歌山",  # 関西
          "鳥取", "島根", "岡山", "広島", "山口", "徳島", "香川", "愛媛", "高知"]:  # 中国・四国
    _PREF_TO_BTOB_CARRIER_120[p] = "福山通運"
# ヤマト120が安い地域: 北海道・九州・沖縄
for p in ["北海道", "福岡", "佐賀", "長崎", "熊本", "大分", "宮崎", "鹿児島", "沖縄"]:
    _PREF_TO_BTOB_CARRIER_120[p] = "ヤマト運輸"
# 140サイズ比較: 福山140 vs ヤマト140 → 同じ地域区分（九州でヤマト140が安い）
_PREF_TO_BTOB_CARRIER_140 = dict(_PREF_TO_BTOB_CARRIER_120)


def _select_shipping_method(product_name, total_cs, dest_pref):
    """商品名・CS数・都道府県から最適な配送方法を選択する。

    ルール（物流費まとめ 2026/3/9版）:
    - Snack 1-4cs: 佐川急便 60サイズ
    - Snack 5cs以上: BtoB最安（福山/ヤマト120）
    - Gummy 1cs: 佐川急便 80サイズ
    - Gummy 2cs以上: BtoB最安（福山/ヤマト120）
    - Energy: BtoB最安（福山/ヤマト140）
    - Water: BtoB最安（福山/ヤマト120）
    - その他: ヤマト運輸（デフォルト）
    """
    pn = product_name.upper() if product_name else ""
    pref_clean = (dest_pref or "").replace("都", "").replace("府", "").replace("県", "").replace("道", "")
    if pref_clean == "北海":
        pref_clean = "北海道"

    def btob_best(size="120"):
        mapping = _PREF_TO_BTOB_CARRIER_140 if size == "140" else _PREF_TO_BTOB_CARRIER_120
        carrier = mapping.get(pref_clean, "福山通運")
        if carrier == "福山通運":
            return "福山通運"
        return f"ヤマト(発払い)B2v6"

    if "SNACK" in pn or "サブレ" in pn or "トリュフ" in pn or "ガーリック" in pn or "ガトーショコラ" in pn:
        if total_cs < 5:
            return "佐川急便"
        return btob_best("120")
    elif "GUMMY" in pn or "グミ" in pn:
        if total_cs <= 1:
            return "佐川急便"
        return btob_best("120")
    elif "ENERGY" in pn or "エナジー" in pn:
        if total_cs <= 1:
            return "佐川急便"
        return btob_best("140")
    elif "WATER" in pn or "ウォーター" in pn or "CERAMIDE" in pn or "セラミド" in pn:
        return btob_best("120")
    else:
        return "ヤマト(発払い)B2v6"


def _classify_product_group(product_name):
    """商品名から伝票分割グループを判定する。
    Returns: (group_key, suffix)
      - Snack 4品 → ("S", "-S") ※混載二重梱包
      - Energy    → ("E", "-E")
      - Water     → ("W", "-W")
      - Gummy     → ("G", "-G")
      - その他    → ("X", "-X")
    """
    pn = (product_name or "").upper()
    if "SNACK" in pn or "サブレ" in pn or "トリュフ" in pn or "ガーリック" in pn or "ガトーショコラ" in pn:
        return "S", "-S"
    elif "ENERGY" in pn or "エナジー" in pn:
        return "E", "-E"
    elif "WATER" in pn or "ウォーター" in pn or "CERAMIDE" in pn or "セラミド" in pn:
        return "W", "-W"
    elif "GUMMY" in pn or "グミ" in pn:
        return "G", "-G"
    else:
        return "X", "-X"


def _build_group_remarks(group_items):
    """グループ内の商品から出荷備考を生成する。

    Snack例: 和紅茶47cs賞味2027/01/05 トリュフ47cs賞味2027/01/05 [二重梱包]
    その他例: 賞味2027/06/01
    """
    parts = []
    has_double_pack = False
    for it in group_items:
        name = it.get("master_name", it.get("ocr_name", ""))
        # 商品略称を生成（ブランド名を除去して短縮）
        short = name
        for prefix in ["2Snack ", "2Water ", "2Gummy ", "RNL 2Energy", "2Energy ", "TWO "]:
            short = short.replace(prefix, "")
        short = short.strip()
        if not short:
            short = name

        qty_cs = it.get("quantity", 0)
        try:
            qty_cs = int(float(qty_cs or 0))
        except (ValueError, TypeError):
            qty_cs = 0

        expiry = it.get("expiry_date", "")
        if expiry:
            expiry_fmt = expiry.replace("-", "/")
            parts.append(f"{short}賞味{expiry_fmt}")
        else:
            parts.append(short)

        if it.get("double_pack"):
            has_double_pack = True

    remarks = " ".join(parts)
    if has_double_pack:
        remarks += " [二重梱包]"
    return remarks


def results_to_coola_csv(results, pdf_name):
    """OCR結果をクーラ（COOLA/ロジザード）用CSV形式に変換して出力する。

    伝票分割ルール:
    - 商品グループ（Snack/Energy/Water/Gummy/その他）ごとに伝票を分割
    - 伝票番号にサフィックス（-S/-E/-W/-G/-X）を自動付番
    - 出荷備考にグループ単位の賞味期限・二重梱包指示を記載
    - Snack 4品は混載1伝票（二重梱包）
    """
    ensure_output_dir()
    csv_name = os.path.splitext(pdf_name)[0] + "_COOLA.csv"
    csv_path = os.path.join(OUTPUT_DIR, csv_name)

    import unicodedata
    oroshisaki = _load_oroshisaki_master()

    rows = []
    for page_result in results:
        if "error" in page_result:
            continue
        ocr = page_result["ocr_raw"]
        ddc = page_result.get("ddc_match", {})

        order_no = ocr.get("order_no", "")
        delivery_date = ocr.get("delivery_date", "")

        # 納品先情報
        dest_name = ddc.get("name", ocr.get("delivery_dest", ""))
        dest_postal = (ddc.get("postal", "") or "").replace("-", "")
        dest_address = ddc.get("address", "") or ""
        dest_tel = ddc.get("tel", "") or ""

        # 卸先マスタから情報取得
        oroshi = oroshisaki.get(dest_name) or oroshisaki.get(unicodedata.normalize('NFKC', dest_name)) or {}
        shop_code = "4"
        shop_name = "卸"
        payment = oroshi.get("支払方法", "その他")

        # 住所分割
        pref, addr1, addr2, addr3 = _split_address(dest_address)

        # マッチ済み商品をグループに分類
        matched_items = [it for it in page_result.get("matched_items", []) if it.get("matched")]
        if not matched_items:
            continue

        groups = {}  # group_key → [items]
        for item in matched_items:
            product_name = item.get("master_name", item.get("ocr_name", ""))
            gkey, _ = _classify_product_group(product_name)
            groups.setdefault(gkey, []).append(item)

        # グループごとに伝票を生成
        suffix_map = {"S": "-S", "E": "-E", "W": "-W", "G": "-G", "X": "-X"}
        for gkey, group_items in groups.items():
            suffix = suffix_map.get(gkey, "")
            slip_no = f"{order_no}{suffix}" if len(groups) > 1 else order_no

            # グループ内の合計金額
            group_total = 0
            for item in group_items:
                irisuu = _get_irisuu(item)
                cs_p = item.get("cs_price", 0)
                qty = item.get("quantity", 0)
                try:
                    unit_p = int(float(cs_p or 0)) // irisuu if irisuu > 0 else int(float(cs_p or 0))
                    q = int(float(qty or 0)) * irisuu
                    group_total += unit_p * q
                except (ValueError, TypeError):
                    pass
            group_tax = int(group_total * 0.1)

            # 出荷備考を生成
            remarks = _build_group_remarks(group_items)

            # グループ内の全CS数合計（配送方法判定用）
            total_cs_in_group = 0
            for item in group_items:
                try:
                    total_cs_in_group += int(float(item.get("quantity", 0) or 0))
                except (ValueError, TypeError):
                    pass

            line_no = 0
            for item in group_items:
                line_no += 1

                qty_cs = item.get("quantity", 0)
                cs_price = item.get("cs_price", 0)
                try:
                    price = int(float(cs_price or 0))
                except (ValueError, TypeError):
                    price = 0
                try:
                    qty_cs_int = int(float(qty_cs or 0))
                except (ValueError, TypeError):
                    qty_cs_int = 0

                irisuu = _get_irisuu(item)
                quantity_bara = qty_cs_int * irisuu
                unit_price_bara = int(price / irisuu) if irisuu > 0 else price

                product_name = item.get("master_name", item.get("ocr_name", ""))
                product_code = item.get("code", "")

                # 配送方法: Snackは混載合計CSで判定、他は個別CS
                if gkey == "S":
                    shipping_method = _select_shipping_method(product_name, total_cs_in_group, pref)
                else:
                    shipping_method = _select_shipping_method(product_name, qty_cs_int, pref)

                row = {h: "" for h in COOLA_HEADERS}
                row.update({
                    "ショップコード": shop_code,
                    "ショップ名": shop_name,
                    "表示受注番号": slip_no,
                    "行no": line_no,
                    "配送先名": dest_name,
                    "配送先郵便番号": dest_postal,
                    "配送先都道府県": pref,
                    "配送先住所１": addr1,
                    "配送先住所２": addr2,
                    "配送先住所３": addr3,
                    "配送先電話番号": dest_tel,
                    "出荷備考": remarks,
                    "出荷日": "",
                    "配送会社id": "",
                    "配送指定日": f"{delivery_date} 00:00:00" if delivery_date else "",
                    "配送時間帯": "午前中",
                    "顧客id": "",
                    "税種区分": "1",
                    "送料": 0,
                    "代引手数料": 0,
                    "消費税": group_tax,
                    "ポイント使用額": 0,
                    "請求額合計": group_total + group_tax,
                    "決済区分id": "",
                    "決済方法": payment,
                    "配送方法": shipping_method,
                    "購入者氏名": dest_name,
                    "購入者郵便番号": dest_postal,
                    "購入者都道府県": pref,
                    "購入者住所１": addr1,
                    "購入者住所２": addr2,
                    "購入者住所３": addr3,
                    "購入者電話番号": dest_tel,
                    "品名": product_name,
                    "形式/型番": product_code,
                    "受注数": quantity_bara,
                    "単価": unit_price_bara,
                    "送状備考1": "",
                    "ギフトフラグ": 0,
                    "消費税率": 10,
                })
                rows.append(row)

    if not rows:
        return None

    with open(csv_path, "w", newline="", encoding="cp932", errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=COOLA_HEADERS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        writer.writerows(rows)

    return csv_path


def results_to_excel(results, pdf_name):
    """Write OCR+matching results to Excel (.xlsx) with DDC dropdown for uncertain rows.

    - OK行: 緑背景、納品先(DDC)は固定テキスト
    - 要確認行: 黄背景、納品先(DDC)にドロップダウン（候補から選択）
    - NG行: 赤背景、納品先(DDC)にドロップダウン
    """
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    ensure_output_dir()
    xlsx_name = os.path.splitext(pdf_name)[0] + "_result.xlsx"
    xlsx_path = os.path.join(OUTPUT_DIR, xlsx_name)

    wb = Workbook()
    ws = wb.active
    ws.title = "受注データ"

    # カラーパレット
    C_HEADER  = "2F5496"  # ヘッダー：濃青
    C_OK      = "E2EFDA"  # OK：薄緑
    C_CAUTION = "FFF2CC"  # 要確認：薄黄
    C_NG      = "FCE4D6"  # NG：薄赤

    # 列定義: (列名, 幅)
    COLUMNS = [
        ("ページ",        5),
        ("オーダーNO",   14),
        ("納品日",        12),
        ("発注元",        22),
        ("納品先(OCR)",   28),
        ("納品先(DDC)",   30),   # ← 要確認/NG にドロップダウン
        ("DDCマッチ",      8),
        ("DDC候補1",      26),
        ("DDC候補2",      26),
        ("DDC候補3",      26),
        ("商品名(OCR)",   26),
        ("商品名(マスタ)", 26),
        ("JANコード",     14),
        ("商品コード",    10),
        ("規格",          14),
        ("配送荷姿",      10),
        ("CS単価",        10),
        ("数量(CS)",       8),
        ("金額",          10),
        ("出力先",        10),
        ("商品マッチ",     8),
        ("備考",          20),
    ]
    headers = [c[0] for c in COLUMNS]
    col_map = {h: i + 1 for i, (h, _) in enumerate(COLUMNS)}

    # ヘッダー行を書き込む
    hdr_fill = PatternFill("solid", fgColor=C_HEADER)
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    for col_idx, (header, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 22

    # 候補を格納する非表示シート（ドロップダウン参照用）
    ws_cand = wb.create_sheet("DDC候補リスト")
    ws_cand.sheet_state = "hidden"
    cand_row_idx = 0  # 候補シートの現在行

    # データ行を構築して書き込む
    ddc_col_letter = get_column_letter(col_map["納品先(DDC)"])

    for page_result in results:
        if "error" in page_result:
            continue
        ocr = page_result["ocr_raw"]
        ddc = page_result["ddc_match"]
        candidates = ddc.get("candidates", [])

        def cand_name(idx):
            return candidates[idx]["name"] if idx < len(candidates) else ""

        def cand_str(idx):
            if idx < len(candidates):
                c = candidates[idx]
                return f"{c['name']} ({c['score']:.0%})"
            return ""

        low_conf = ddc.get("low_confidence", False)
        match_status = "OK" if ddc.get("matched") else "NG"
        if ddc.get("matched") and low_conf:
            match_status = "要確認"

        # ドロップダウン用候補名リスト（空除外）
        pure_cands = [cand_name(i) for i in range(3) if cand_name(i)]

        for item in page_result["matched_items"]:
            row_data = {
                "ページ":        page_result["page"],
                "オーダーNO":    ocr.get("order_no", ""),
                "納品日":        ocr.get("delivery_date", ""),
                "発注元":        ocr.get("sender", ""),
                "納品先(OCR)":   ocr.get("delivery_dest", ""),
                "納品先(DDC)":   ddc.get("name", cand_name(0)) if ddc.get("matched") else cand_name(0),
                "DDCマッチ":     match_status,
                "DDC候補1":      cand_str(0),
                "DDC候補2":      cand_str(1),
                "DDC候補3":      cand_str(2),
                "商品名(OCR)":   item.get("ocr_name", ""),
                "商品名(マスタ)": item.get("master_name", ""),
                "JANコード":     item.get("jan", ""),
                "商品コード":    item.get("code", ""),
                "規格":          item.get("spec", ""),
                "配送荷姿":      item.get("pack", ""),
                "CS単価":        item.get("cs_price", ""),
                "数量(CS)":      item.get("quantity", ""),
                "金額":          item.get("amount", ""),
                "出力先":        item.get("output_dest", ""),
                "商品マッチ":    "OK" if item.get("matched") else "NG",
                "備考":          ocr.get("notes", ""),
            }

            # 背景色を選択
            if match_status == "OK":
                row_fill = PatternFill("solid", fgColor=C_OK)
            elif match_status == "要確認":
                row_fill = PatternFill("solid", fgColor=C_CAUTION)
            else:
                row_fill = PatternFill("solid", fgColor=C_NG)

            # 現在の書き込み行（ヘッダー=1なのでデータは2行目〜）
            data_row = ws.max_row + 1

            for col_name in headers:
                cell = ws.cell(row=data_row, column=col_map[col_name], value=row_data[col_name])
                cell.fill = row_fill
                cell.alignment = Alignment(vertical="center")

            # 要確認/NG行にドロップダウンを追加
            if match_status in ("要確認", "NG") and pure_cands:
                cand_row_idx += 1
                # 候補を非表示シートに縦方向に書き込む
                for c_col, name in enumerate(pure_cands, 1):
                    ws_cand.cell(row=cand_row_idx, column=c_col, value=name)

                # データ検証: 候補シートの該当行を参照
                end_col = get_column_letter(len(pure_cands))
                formula = f"'DDC候補リスト'!$A${cand_row_idx}:${end_col}${cand_row_idx}"
                dv = DataValidation(
                    type="list",
                    formula1=formula,
                    showErrorMessage=False,
                    showInputMessage=True,
                    promptTitle="DDC選択",
                    prompt="正しい納品先DDCを選択してください",
                )
                dv.add(f"{ddc_col_letter}{data_row}")
                ws.add_data_validation(dv)

    # ヘッダー固定・フィルター
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    wb.save(xlsx_path)
    return xlsx_path


def generate_pdfs(results, pdf_name, staff_name="伊藤", remarks=""):
    """Generate Sylvia and/or Haruna order PDFs from processing results"""
    ensure_output_dir()
    generated = []

    for page_result in results:
        if "error" in page_result:
            continue

        ocr = page_result["ocr_raw"]
        ddc = page_result["ddc_match"]
        sylvia_items = page_result["sylvia_items"]
        haruna_items = page_result["haruna_items"]
        order_no = ocr.get("order_no", "")
        dest_name = ddc.get("name", ocr.get("delivery_dest", "unknown"))
        page_num = page_result["page"]
        base_name = os.path.splitext(pdf_name)[0]

        # Sylvia PDF
        if sylvia_items:
            items_data = []
            for i in sylvia_items:
                items_data.append({
                    "jan": i.get("jan", ""),
                    "name": i["master_name"],
                    "spec": i.get("spec", ""),
                    "pack": i.get("pack", ""),
                    "unit_price": int(i.get("unit_price", 0)),
                    "cs_price": int(i.get("cs_price", 0)),
                    "quantity": i["quantity"],
                    "amount": int(i.get("amount", 0)),
                })
            order_data = {
                "order_no": order_no,
                "order_date": str(date.today()),
                "delivery_date": ocr.get("delivery_date", ""),
                "delivery_dest": dest_name,
                "postal": ddc.get("postal", ""),
                "address": ddc.get("address", ""),
                "tel": ddc.get("tel", ""),
                "fax": ddc.get("fax", ""),
                "remarks": remarks,
            }
            pdf_buf = gen_sylvia_pdf(order_data, items_data, staff_name)
            out_name = f"シルビア_{base_name}_p{page_num}.pdf"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            with open(out_path, "wb") as f:
                f.write(pdf_buf.read())
            generated.append(("シルビア", out_path))

        # Haruna PDF
        if haruna_items:
            total_qty = sum(i["quantity"] for i in haruna_items)
            order_data = {
                "order_no": order_no,
                "order_date": str(date.today()),
                "delivery_date": ocr.get("delivery_date", ""),
                "delivery_dest": dest_name,
                "quantity": total_qty,
                "remarks": remarks,
            }
            ddc_data = {
                "postal": ddc.get("postal", ""),
                "address": ddc.get("address", ""),
                "tel": ddc.get("tel", ""),
                "fax": ddc.get("fax", ""),
                "time": ddc.get("time", ""),
                "berse": ddc.get("berse", "無"),
                "palette": ddc.get("palette", ""),
                "jpr": ddc.get("jpr", ""),
                "method": ddc.get("method", ""),
            }
            pdf_buf = gen_haruna_pdf(order_data, ddc_data, staff_name)
            out_name = f"ハルナ_{base_name}_p{page_num}.pdf"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            with open(out_path, "wb") as f:
                f.write(pdf_buf.read())
            generated.append(("ハルナ", out_path))

    return generated


def print_summary(results, csv_path, xlsx_path, generated_pdfs):
    """Print processing summary"""
    print("\n" + "=" * 60)
    print("処理結果サマリー")
    print("=" * 60)

    for page_result in results:
        page = page_result["page"]
        if "error" in page_result:
            print(f"\n  Page {page}: ERROR - {page_result['error']}")
            continue

        ocr = page_result["ocr_raw"]
        ddc = page_result["ddc_match"]
        print(f"\n  Page {page}:")
        print(f"    オーダーNO : {ocr.get('order_no', '(なし)')}")
        print(f"    納品日     : {ocr.get('delivery_date', '(なし)')}")
        print(f"    発注元     : {ocr.get('sender', '(なし)')}")
        print(f"    納品先(OCR): {ocr.get('delivery_dest', '(なし)')}")
        if ddc.get("matched"):
            score_str = ""
            if "match_score" in ddc:
                score_str = f" (類似度: {ddc['match_score']:.0%})"
            conf_str = " ⚠要確認" if ddc.get("low_confidence") else ""
            print(f"    納品先(DDC): {ddc['name']}{score_str}{conf_str}")
        else:
            print(f"    納品先(DDC): *** 未マッチ ***")

        candidates = ddc.get("candidates", [])
        if candidates:
            print(f"    DDC候補:")
            for i, c in enumerate(candidates, 1):
                print(f"      {i}. {c['name']} ({c['score']:.0%})")

        print(f"    商品:")
        for item in page_result["matched_items"]:
            if item.get("matched"):
                print(f"      OK  {item['master_name']} x {item['quantity']}CS → {item['output_dest']}")
            else:
                print(f"      NG  {item.get('ocr_name', '?')} x {item.get('quantity', '?')}CS *** 未マッチ ***")

        if page_result["sylvia_items"]:
            total = sum(i["amount"] for i in page_result["sylvia_items"])
            print(f"    → シルビア: {len(page_result['sylvia_items'])}商品, 合計 ¥{total:,.0f}")
        if page_result["haruna_items"]:
            total_cs = sum(i["quantity"] for i in page_result["haruna_items"])
            print(f"    → ハルナ: {total_cs}CS ({total_cs * 24:,}本)")

    print(f"\n  CSV出力:   {csv_path}")
    print(f"  Excel出力: {xlsx_path}")
    for dest, path in generated_pdfs:
        print(f"  {dest}PDF: {path}")
    print("=" * 60)


def main():
    parser = argparse.ArgumentParser(description="FAX発注書 自動処理スクリプト")
    parser.add_argument("pdf_files", nargs="+", help="FAX PDFファイルのパス")
    parser.add_argument("--staff", default="伊藤", help="担当者名 (default: 伊藤)")
    parser.add_argument("--remarks", default="", help="イレギュラーリクエスト/備考")
    parser.add_argument("--model", default="claude-sonnet-4-6", help="Claude model for OCR")
    args = parser.parse_args()

    for pdf_path in args.pdf_files:
        if not os.path.exists(pdf_path):
            print(f"ERROR: ファイルが見つかりません: {pdf_path}")
            continue

        pdf_name = os.path.basename(pdf_path)
        print(f"\n{'=' * 60}")
        print(f"Processing: {pdf_name}")
        print(f"{'=' * 60}")

        # 1. OCR + matching
        results = process_fax_pdf(pdf_path, model=args.model)

        # 2. CSV + Excel output
        csv_path = results_to_csv(results, pdf_name)
        xlsx_path = results_to_excel(results, pdf_name)

        # 3. PDF generation
        generated = generate_pdfs(results, pdf_name, args.staff, args.remarks)

        # 4. Summary
        print_summary(results, csv_path, xlsx_path, generated)


if __name__ == "__main__":
    main()
