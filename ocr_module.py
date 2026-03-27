"""OCR module - Claude Vision API for FAX PDF reading"""
import os
import sys
import json
import base64
import unicodedata

# Windows terminal UTF-8 output
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

import anthropic
import pandas as pd
from difflib import SequenceMatcher


def normalize(s):
    """NFKC normalization: convert full-width chars to half-width (e.g. ＤＤＣ→DDC)"""
    t = unicodedata.normalize('NFKC', str(s)).strip()
    # ノーブレークスペース(\xa0)等の特殊スペースを通常スペースに変換
    t = t.replace('\xa0', ' ').replace('\u3000', ' ')
    return t


def normalize_company(s):
    """会社名の表記ゆれを統一してからNFKC正規化する。
    「株式会社」「(株)」「㈱」→ 空文字に統一し、スペースも正規化。
    DDCマッチ時のスコア比較専用（表示には使わない）。
    """
    import re
    t = normalize(s)
    # 株式会社の表記ゆれをスペースに変換（中間にある場合のセパレータ役割を保持）
    # 空文字にすると「伊藤忠食品(株)三郷物流センター」→「伊藤忠食品三郷物流センター」になり分割不能
    t = re.sub(r'株式会社|（株）|\(株\)|㈱', ' ', t)
    # 有限会社
    t = re.sub(r'有限会社|（有）|\(有\)|㈲', ' ', t)
    # 角括弧内のメモを除去 例: ＜TWO＞ → '' / <TWO> → ''
    t = re.sub(r'[<＜][^>＞]*[>＞]', '', t)
    # 末尾の「宛」「様」「御中」「行」を除去
    t = re.sub(r'[宛様行]$', '', t)
    t = re.sub(r'御中$', '', t)
    # 連続スペースを1つに、前後の空白除去
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def get_api_key():
    """Get Anthropic API key from environment variable"""
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        key_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".api_key")
        if os.path.exists(key_file):
            with open(key_file, "r") as f:
                key = f.read().strip()
    return key


def _auto_rotate_image(img):
    """FAX画像の向きを自動検出して正立方向に回転する。

    FAXのPDFには横向き（ランドスケープ）で格納されているページがある。
    縦書き方向の分散比率 (h_cv/v_cv) で横向きを検出し、
    上部の暗ピクセル密度でCW/CCWどちらに回転すべきかを判定する。

    検出は低解像度サムネイルで行い、実際の回転はオリジナル解像度で適用する。

    Args:
        img: PIL.Image オブジェクト

    Returns:
        回転補正後の PIL.Image オブジェクト
    """
    import numpy as np

    # 検出用に縮小（長辺600px相当）— 精度は十分、速度を優先
    thumb_size = 600
    ratio = thumb_size / max(img.width, img.height)
    thumb = img.resize(
        (max(1, int(img.width * ratio)), max(1, int(img.height * ratio))),
        resample=0  # NEAREST: 最速
    )

    arr = np.array(thumb.convert('L'))
    dark = (arr < 128).astype(np.float32)
    hproj = dark.sum(axis=1)
    vproj = dark.sum(axis=0)
    h_cv = float(np.std(hproj)) / (float(np.mean(hproj)) + 1e-6)
    v_cv = float(np.std(vproj)) / (float(np.mean(vproj)) + 1e-6)
    h_score = h_cv / (v_cv + 1e-6)

    # h_score < 0.8 → テキスト行が縦方向 → 横向きページ
    if h_score >= 0.8:
        return img  # 正立方向、回転不要

    # CCW (+90°) と CW (-90°) どちらが正しいか：
    # FAX文書は上部にヘッダがあるため、正立方向では上部の暗ピクセルが多い
    def top_bottom_ratio(rotated_thumb):
        a = np.array(rotated_thumb.convert('L'))
        h = a.shape[0]
        top = float((a[:h // 4] < 128).sum())
        bot = float((a[3 * h // 4:] < 128).sum())
        return top / (bot + 1e-6)

    thumb_ccw = thumb.rotate(90, expand=True)
    thumb_cw = thumb.rotate(-90, expand=True)
    if top_bottom_ratio(thumb_ccw) >= top_bottom_ratio(thumb_cw):
        return img.rotate(90, expand=True)
    return img.rotate(-90, expand=True)


def pdf_to_images(pdf_bytes, dpi=300):
    """Convert PDF to list of per-page base64 PNG images (with auto-rotation)"""
    import fitz  # PyMuPDF
    from PIL import Image
    import io
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        img = _auto_rotate_image(img)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        images.append({"page": page_num + 1, "base64": b64})
    doc.close()
    return images


OCR_PROMPT = """このFAX発注書の画像を正確に読み取り、以下のJSON形式で情報を抽出してください。

{
  "order_no": "オーダーNO/発注NO",
  "delivery_date": "納品日/入荷日 (YYYY-MM-DD)",
  "sender": "発注者/発注元の会社名",
  "sender_contact": "担当者名",
  "sender_fax": "発注元FAX番号",
  "delivery_dest": "納品先/入荷場所の名称",
  "delivery_address": "納品先住所",
  "delivery_tel": "納品先電話番号",
  "delivery_fax": "納品先FAX番号",
  "items": [
    {
      "product_name": "商品名",
      "jan_code": "JANコード（13桁）",
      "product_code": "商品コード",
      "quantity_cs": "ケース数（CS単位の数値のみ）",
      "unit_price": "単価（数値、あれば）",
      "amount": "金額（数値、あれば）"
    }
  ],
  "notes": "備考・特記事項"
}

【重要】この発注書で扱う主な商品は以下のとおりです：
1. ２Snack香り華やぐ和紅茶サブレ（JAN: 4589570801393）入数12袋/cs
2. ２Snack香るトリュフ（JAN: 4589570801416）入数12袋/cs
3. ２Snack焦がしガーリック（JAN: 4589570801430）入数12袋/cs
4. ２Snack濃厚ガトーショコラ風サブレ（JAN: 4589570801454）入数12袋/cs
5. 2Water Ceramide（JAN: 4589570801485）入数24本/cs
6. 2Gummy LIPOSOME VC（JAN: 4589570801331）入数48袋/cs
7. 2Energy 250ml（JAN: 4589570801348）入数30本/cs
8. 2Energy 250ml 26RN（JAN: 4589570801621）入数30本/cs

読み取りルール：
- 商品名が読みにくい場合、JANコード（4589570801で始まる13桁）から上記リストで特定してください。
- quantity_csは必ず「ケース数(CS)」です。入数（12袋、24本等）や数量バラ（個数）ではありません。
  例：「12袋×3cs」→ quantity_cs は 3。「24本/cs × 10」→ quantity_cs は 10。
  数量(CS)列、発注数量(CS)列、函数列の値を読み取ってください。
- 「函数」「函」もケース数(CS)と同じ意味です。例：「8函」→ quantity_cs は 8。
- 手書き修正がある場合（取り消し線で数字が消されている場合）は、最終的な記載数字を使用してください。
- itemsの並び順は、FAX原本の表に記載されている上から下の順番を維持してください。
- delivery_dateは納品日（入荷日）です。製造日・賞味期限の日付と混同しないこと。
  年が不明な場合は2026年としてください。
- 数量が0または空欄の商品行はitemsに含めないでください。

【読み取り注意】
- 「函数」「函」はケース数(CS)と同じ意味です。「入数」（12入など）は1CS内の袋/本数なので混同しないこと。
- 商品名に「TWO」ブランドが先頭に付く場合（例：TWO 2Snack 焦がしガーリック）、上記商品リストで照合してください。
- FAX画質の劣化により「昭」「鷺」「鷹」は字形が似ているため慎重に判断してください。
- 「埼京」と「埼玉」は1文字違いです（右側の字が「京」か「玉」か正確に判断）。
- 「DDC」「SDC」「RDC」等のアルファベットコードは正確に読み取ってください。

【発注書フォーマットへの対応】ページ上部に「発注書」というタイトルがある場合：
- 「納入先」欄は右側に配置されているので、その会社名＋施設名を delivery_dest に読み取ること。
- 左上の「○○商事（株）御中」はFAXの宛先（受信者）なので delivery_dest には使わないこと。
- 「函数」列の数字が quantity_cs です（「入数」列の12入等と混同しないこと）。

【入荷場所が複数記載されている場合】
- 伝票に入荷場所・納品先が複数箇所に記載されているフォーマットがあります。
- その場合は **一番上（最初）に記載されている入荷場所** を delivery_dest として読み取ってください。
- 下段の入荷場所コードやライフ番号等は無視してください。

- JSONのみを返してください。説明文やマークダウンは不要です。"""


def ocr_fax_page(image_b64, model="claude-sonnet-4-6"):
    """Read one FAX page image with Claude Vision API and return parsed JSON"""
    # Windows環境でのエンコーディング問題を回避
    import io as _io
    if sys.stdout and hasattr(sys.stdout, 'encoding') and (sys.stdout.encoding or '').lower() in ('ascii', 'cp932', 'shift_jis', 'mbcs'):
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    if sys.stderr and hasattr(sys.stderr, 'encoding') and (sys.stderr.encoding or '').lower() in ('ascii', 'cp932', 'shift_jis', 'mbcs'):
        sys.stderr = _io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)

    api_key = get_api_key()
    if not api_key:
        return {"error": "API key not set. Set ANTHROPIC_API_KEY environment variable."}

    client = anthropic.Anthropic(api_key=api_key)

    try:
        response = client.messages.create(
            model=model,
            max_tokens=2000,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": image_b64
                        }
                    },
                    {"type": "text", "text": OCR_PROMPT}
                ]
            }]
        )

        text = response.content[0].text.strip()
        # Extract JSON from response
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0].strip()
        elif "```" in text:
            text = text.split("```")[1].split("```")[0].strip()

        return json.loads(text)

    except json.JSONDecodeError:
        return {"error": "JSON parse failed", "raw": text}
    except UnicodeEncodeError as e:
        return {"error": "Encoding error: " + repr(e)}
    except Exception as e:
        try:
            msg = repr(e)
        except Exception:
            msg = type(e).__name__
        return {"error": "OCR error: " + msg}


def load_product_master():
    """Load product master (Supabase優先、CSVフォールバック)"""
    from supabase_client import load_product_master as _load_from_supabase
    return _load_from_supabase()


def load_ddc_master():
    """Load DDC master (Supabase torihikisaki_master優先、CSVフォールバック)"""
    from supabase_client import load_torihikisaki_master_from_supabase
    df = load_torihikisaki_master_from_supabase()
    if df is not None and len(df) > 0:
        return df
    # フォールバック: ローカルCSV
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "ddc_master.csv")
    print(f"  [CSV] 納品先マスタ読み込み: {csv_path}")
    return pd.read_csv(csv_path)


def load_staff():
    """Load staff JSON"""
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "staff.json")
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


def match_product(ocr_item, product_master):
    """Match a single OCR item to product master.
    Returns dict with matched product data or None."""
    jan = str(ocr_item.get("jan_code", "")).strip()
    ocr_name = normalize(ocr_item.get("product_name", ""))

    # 1. JAN code exact match
    if jan and len(jan) == 13:
        m = product_master[product_master["JANコード"] == jan]
        if len(m) > 0:
            return _product_row_to_dict(m.iloc[0], ocr_item)

    # 2. Product name exact match (normalized)
    if ocr_name:
        norm_pm = product_master["商品名"].apply(normalize)
        exact_idx = norm_pm[norm_pm == ocr_name].index
        if len(exact_idx) > 0:
            return _product_row_to_dict(product_master.loc[exact_idx[0]], ocr_item)

    # 3. Fuzzy match on product name (both sides normalized)
    if ocr_name:
        best_ratio = 0
        best_row = None
        for _, row in product_master.iterrows():
            master_name = normalize(row["商品名"])
            score = 0.0
            if ocr_name in master_name or master_name in ocr_name:
                score = 0.85
            seq_ratio = SequenceMatcher(None, ocr_name, master_name).ratio()
            final_score = max(score, seq_ratio)
            if final_score > best_ratio and final_score >= 0.5:
                best_ratio = final_score
                best_row = row
        if best_row is not None:
            return _product_row_to_dict(best_row, ocr_item)

    return None


def _product_row_to_dict(row, ocr_item):
    """Convert a product master row to a standardized dict"""
    try:
        qty = int(float(str(ocr_item.get("quantity_cs", 0)).strip() or "0"))
    except (ValueError, TypeError):
        qty = 0
    cs_price = float(row.get("CS単価", 0) or 0)
    case_qty = int(row.get("入数", 0) or 0)
    return {
        "matched": True,
        "ocr_name": str(ocr_item.get("product_name", "")),
        "master_name": str(row["商品名"]),
        "jan": str(row["JANコード"]) if pd.notna(row.get("JANコード")) else "",
        "code": str(row.get("商品コード", "")),
        "spec": str(row.get("規格", "")),
        "pack": str(row.get("配送荷姿", "")),
        "unit_price": float(row.get("1袋単価", 0) or 0),
        "cs_price": cs_price,
        "case_quantity": case_qty,
        "output_dest": str(row.get("出力先", "")),
        "quantity": qty,
        "amount": qty * cs_price,
    }


def _normalize_dc(s):
    """DDC/SDC/RDC等を統一（スコアリング専用）: [A-Z]DC → _DC
    OCRが DDC を SDC と誤読するケースを吸収する。"""
    import re
    return re.sub(r'[A-Z]DC', '_DC', s)


def _parts_score(ocr, master):
    """OCR と master の部分文字列マッチスコアを算出（パーツ比較ヘルパー）"""
    best = 0.0
    # OCR側がスペース区切り
    if ' ' in ocr:
        for part in [p for p in ocr.split(' ') if len(p) >= 2]:
            r = SequenceMatcher(None, part, master).ratio()
            s = r * 0.85
            if len(part) >= 4 and (part in master or master in part):
                s = max(s, 0.70)
            best = max(best, s)
    # マスタ側がスペース区切り
    if ' ' in master:
        for mpart in [p for p in master.split(' ') if len(p) >= 2]:
            r = SequenceMatcher(None, ocr, mpart).ratio()
            s = r * 0.85
            if len(mpart) >= 4 and (ocr in mpart or mpart in ocr):
                s = max(s, 0.70)
            best = max(best, s)
    return best


def _score_ddc(dest_name_clean, dest_company, master_name, master_company):
    """DDCマッチスコアを算出する（内部関数）"""
    # 通常の正規化でスコア算出
    score = 0.0
    if dest_name_clean in master_name or master_name in dest_name_clean:
        score = 0.85
    seq_ratio = SequenceMatcher(None, dest_name_clean, master_name).ratio()

    # 会社名正規化版でもスコア算出（株式会社等の表記ゆれを吸収）
    score_company = 0.0
    if dest_company in master_company or master_company in dest_company:
        score_company = 0.85
    seq_ratio_company = SequenceMatcher(None, dest_company, master_company).ratio()

    # マスタ名が長い場合がOCRの長い文字列に含まれているかチェック
    if len(master_company) >= 5 and master_company in dest_company:
        score_company = max(score_company, 0.90)

    best = max(score, seq_ratio, score_company, seq_ratio_company)

    # パーツマッチング（スペース区切りの各部分を比較）
    best = max(best, _parts_score(dest_company, master_company))

    # DDC/SDC/RDC 1文字誤読吸収: [A-Z]DC を _DC に統一してパーツ比較
    # 例: OCR「埼京SDC」→「埼京_DC」 vs マスタ「ローソン埼京DDC」→「ローソン埼京_DC」
    dest_dc = _normalize_dc(dest_company)
    master_dc = _normalize_dc(master_company)
    if dest_dc != dest_company or master_dc != master_company:
        best = max(best, _parts_score(dest_dc, master_dc))

    return best


def match_ddc(dest_name, ddc_master, max_candidates=3, sender=""):
    """Match delivery destination name to DDC master.

    Args:
        dest_name: OCRで読み取った納品先名
        ddc_master: 納品先マスタ DataFrame
        max_candidates: 返す候補数
        sender: 発注元会社名（OCRから取得）。マスタ名の先頭と一致すればスコアをブースト。

    Returns dict with:
        - matched: True/False
        - name, address, tel, etc. (1位の候補データ)
        - match_score: 1位のスコア
        - low_confidence: スコア0.92未満ならTrue（0.90は要確認ゾーン）
        - candidates: 上位N件のリスト [{name, score, address}, ...]
    """
    dest_name_clean = normalize(dest_name)
    if not dest_name_clean:
        return {"matched": False, "name": dest_name, "candidates": []}

    dest_company = normalize_company(dest_name)
    # 発注元の正規化（"三菱食品株式会社 加食G" → "三菱食品 加食G" → 先頭4文字以上で判定）
    sender_norm = normalize_company(sender) if sender else ""
    # 卸先コードプレフィックス（"三菱食品 加食G" の "三菱食品" 部分だけ取る）
    sender_key = sender_norm.split(' ')[0] if ' ' in sender_norm else sender_norm

    # 1. Exact match (normalized)
    norm_master = ddc_master["納品先名"].apply(normalize)
    exact_idx = norm_master[norm_master == dest_name_clean].index
    if len(exact_idx) > 0:
        result = _ddc_row_to_dict(ddc_master.loc[exact_idx[0]])
        result["match_score"] = 1.0
        result["low_confidence"] = False
        result["candidates"] = [{"name": result["name"], "score": 1.0}]
        return result

    # 1b. Exact match (会社名正規化版)
    comp_master = ddc_master["納品先名"].apply(normalize_company)
    exact_comp_idx = comp_master[comp_master == dest_company].index
    if len(exact_comp_idx) > 0:
        result = _ddc_row_to_dict(ddc_master.loc[exact_comp_idx[0]])
        result["match_score"] = 1.0
        result["low_confidence"] = False
        result["candidates"] = [{"name": result["name"], "score": 1.0}]
        return result

    # 2. Fuzzy scoring → 上位N件を候補として返す
    DDC_FUZZY_THRESHOLD = 0.35  # 候補収集は広めに（採用判定は0.7以上）
    scored = []
    sender_matched = []  # 発注元プレフィックスが一致するマスタ（base scoreに関わらず収集）
    for _, row in ddc_master.iterrows():
        master_name = normalize(row["納品先名"])
        master_company = normalize_company(row["納品先名"])
        base_score = _score_ddc(dest_name_clean, dest_company, master_name, master_company)

        is_sender_match = (
            sender_key and len(sender_key) >= 3
            and master_company.startswith(sender_key)
        )

        # 施設名一致チェック: 「会社名 施設名」または「会社名(施設名)」形式の場合
        # 会社名が一致していても施設名が違うマスタには高スコアを付けない
        facility_ok = None  # None=判定不可, True=施設名一致, False=施設名不一致
        import re as _re

        def _split_facility(s):
            """「会社名 施設名」または「会社名(施設名)」を (head, tail) に分割。不可なら (s, None)。"""
            if ' ' in s:
                h = s.split(' ')[0]
                t = s[len(h) + 1:]
                if len(t) >= 2:
                    return h, t
            m = _re.match(r'^(.+?)\((.+)\)$', s)
            if m and len(m.group(1)) >= 2 and len(m.group(2)) >= 2:
                return m.group(1).strip(), m.group(2).strip()
            return s, None

        ocr_head, ocr_tail = _split_facility(dest_company)
        mst_head, mst_tail = _split_facility(master_company)
        if ocr_tail is not None and mst_tail is not None:
            # 両方「会社名 施設名」形式: 会社名一致なら施設名を比較
            head_r = SequenceMatcher(None, ocr_head, mst_head).ratio()
            tail_r = SequenceMatcher(None, ocr_tail, mst_tail).ratio()
            if head_r >= 0.8:
                facility_ok = tail_r >= 0.85
        elif ocr_tail is None and mst_tail is not None and is_sender_match:
            # OCRが施設名のみ（「静岡物流センター」）、マスタが「会社名 施設名」の場合
            # sender boost が効くケースでのみ、OCR全体とマスタ施設部分を比較
            tail_r = SequenceMatcher(None, dest_company, mst_tail).ratio()
            if tail_r < 0.85:
                facility_ok = False

        # 発注元ブースト: マスタ名の先頭が発注元社名と一致する場合にスコアを加算
        # 例: sender="三菱食品", master="三菱食品 ローソン埼京DDC" → +0.25
        final_score = base_score
        if is_sender_match and base_score >= 0.15:
            final_score = min(base_score + 0.25, 0.95)

        # 施設名が不一致の場合: ブーストしても要確認止まりに制限
        # 例: OCR=「船橋物流センター」, マスタ=「昭島センター」→ 0.88以下に制限
        if facility_ok is False:
            final_score = min(final_score, 0.88)

        if final_score >= DDC_FUZZY_THRESHOLD:
            scored.append((final_score, row))

        # OCR誤読対策: 発注元マッチのマスタはbaseスコアが低くても別収集しておく
        if is_sender_match:
            sender_matched.append((min(base_score + 0.25, 0.95), row))

    # スコア降順でソート
    scored.sort(key=lambda x: x[0], reverse=True)

    # 発注元マッチ強制候補: OCRが完全に誤読した場合でも候補に出す
    # 例: 昭島→鷹殿 のような誤読でも「伊藤忠食品（株）昭島センター」が候補に残る
    if sender_matched:
        sender_matched.sort(key=lambda x: x[0], reverse=True)
        best_sender_score, best_sender_row = sender_matched[0]
        if best_sender_score >= 0.30:
            existing_names = {str(r.get("納品先名", "")) for _, r in scored}
            if str(best_sender_row.get("納品先名", "")) not in existing_names:
                scored.append((best_sender_score, best_sender_row))
                scored.sort(key=lambda x: x[0], reverse=True)

    # 候補リスト作成
    candidates = []
    for s, row in scored[:max_candidates]:
        candidates.append({
            "name": str(row.get("納品先名", "")),
            "score": round(s, 3),
            "address": str(row.get("住所", "")),
        })

    # 1位が0.7以上なら採用
    if scored and scored[0][0] >= 0.7:
        best_score, best_row = scored[0]
        result = _ddc_row_to_dict(best_row)
        result["match_score"] = best_score
        # 0.90未満は低信頼度（候補確認推奨）
        result["low_confidence"] = best_score < 0.92
        result["candidates"] = candidates
        return result

    # 未マッチ（候補は残す）
    return {"matched": False, "name": dest_name, "candidates": candidates}


def _ddc_row_to_dict(row):
    """Convert a DDC master row to dict"""
    def safe(val):
        s = str(val).strip() if val is not None else ""
        return "" if s.lower() in ("nan", "none", "null") else s
    return {
        "matched": True,
        "name": safe(row.get("納品先名", "")),
        "nohinsaki_code": safe(row.get("納品先コード", "")),
        "torihikisaki": safe(row.get("取引先名", "")),
        "torihikisaki_code": safe(row.get("取引先コード", "")),
        "postal": safe(row.get("郵便番号", "")),
        "address": safe(row.get("住所", "")),
        "tel": safe(row.get("電話番号", "")),
        "fax": safe(row.get("FAX番号", "")),
        "time": safe(row.get("入荷時間", "")),
        "berse": safe(row.get("バース予約", "")),
        "palette": safe(row.get("パレット条件", "")),
        "jpr": safe(row.get("JPRコード", "")),
        "method": safe(row.get("納品方法", "")),
    }


def process_fax_pdf(pdf_path, model="claude-sonnet-4-6"):
    """Process a FAX PDF file end-to-end.

    Args:
        pdf_path: Path to the FAX PDF file
        model: Claude model to use for OCR

    Returns:
        list of dicts, one per page, each containing:
            - ocr_raw: raw OCR result
            - matched_items: list of matched product items
            - ddc_match: matched delivery destination
            - sylvia_items: items for Sylvia
            - haruna_items: items for Haruna
    """
    pm = load_product_master()
    ddc = load_ddc_master()

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    pages = pdf_to_images(pdf_bytes, dpi=300)
    results = []

    for page_info in pages:
        print(f"  Page {page_info['page']}/{len(pages)} - OCR processing...")
        ocr_result = ocr_fax_page(page_info["base64"], model=model)

        if "error" in ocr_result:
            results.append({"page": page_info["page"], "error": ocr_result["error"], "ocr_raw": ocr_result})
            continue

        # 年の異常値補正: OCRが製造年や賞味期限年を発注日と誤読する場合がある
        delivery_date = ocr_result.get("delivery_date", "")
        if delivery_date and len(delivery_date) >= 4 and delivery_date[:4].isdigit():
            year = int(delivery_date[:4])
            if year < 2020 or year > 2030:
                ocr_result["delivery_date"] = "2026" + delivery_date[4:]

        # Match products
        matched_items = []
        for ocr_item in ocr_result.get("items", []):
            match = match_product(ocr_item, pm)
            if match:
                matched_items.append(match)
            else:
                matched_items.append({
                    "matched": False,
                    "ocr_name": ocr_item.get("product_name", ""),
                    "jan": str(ocr_item.get("jan_code", "")),
                    "quantity": int(float(str(ocr_item.get("quantity_cs") or 0).strip() or "0")),
                })

        # Match DDC（発注元情報もヒントとして渡す）
        ddc_match = match_ddc(
            ocr_result.get("delivery_dest", ""),
            ddc,
            sender=ocr_result.get("sender", ""),
        )

        # Split by output destination
        sylvia_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "シルビア"]
        haruna_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "ハルナ"]
        warehouse_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "自社倉庫"]

        results.append({
            "page": page_info["page"],
            "ocr_raw": ocr_result,
            "matched_items": matched_items,
            "ddc_match": ddc_match,
            "sylvia_items": sylvia_items,
            "haruna_items": haruna_items,
            "warehouse_items": warehouse_items,
        })

    return results
