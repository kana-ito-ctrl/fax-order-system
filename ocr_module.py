"""OCR モジュール - Claude Vision APIでFAX PDFを読み取り"""
import os
import json
import base64
import anthropic
import streamlit as st


def get_api_key():
    """APIキーを取得（Streamlit Secrets → 環境変数の順）"""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except:
        return os.environ.get("ANTHROPIC_API_KEY", "")


def pdf_to_images(pdf_bytes):
    """PDFをページごとの画像(base64)リストに変換"""
    import fitz  # PyMuPDF
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        # 200 DPI for good OCR quality
        mat = fitz.Matrix(200/72, 200/72)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        images.append({"page": page_num + 1, "base64": b64})
    doc.close()
    return images


def ocr_fax_page(image_b64):
    """1ページ分のFAX画像をClaude Vision APIで読み取り"""
    api_key = get_api_key()
    if not api_key:
        return {"error": "APIキーが設定されていません"}

    client = anthropic.Anthropic(api_key=api_key)

    prompt = """このFAX発注書の画像を読み取り、以下のJSON形式で情報を抽出してください。

{
  "order_no": "オーダーNO/発注NO",
  "order_date": "発注日 (YYYY-MM-DD)",
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
      "jan_code": "JANコード（4589570で始まる13桁）",
      "product_code": "商品コード",
      "quantity_cs": "ケース数（数値）",
      "quantity_bara": "バラ数/入数（数値）",
      "unit_price": "単価（数値、あれば）",
      "amount": "金額（数値、あれば）",
      "best_before": "賞味期限条件（あれば）"
    }
  ],
  "notes": "備考・特記事項"
}

注意事項：
- JANコードは4589570801で始まる13桁の数字です。必ず正確に読み取ってください。
- 数量はケース(CS)単位の数値で返してください。
- 金額や単価がない場合は空文字にしてください。
- JSONのみを返してください。説明文は不要です。"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
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
                    {"type": "text", "text": prompt}
                ]
            }]
        )

        text = response.content[0].text.strip()
        # JSON部分を抽出
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0].strip()
        elif "```" in text:
            text = text.split("```")[1].split("```")[0].strip()

        return json.loads(text)

    except json.JSONDecodeError:
        return {"error": "OCR結果のJSON解析に失敗しました", "raw": text}
    except Exception as e:
        return {"error": f"OCRエラー: {str(e)}"}


def match_products(ocr_items, product_master):
    """OCR結果を商品マスタとマッチング"""
    matched = []
    for item in ocr_items:
        jan = str(item.get("jan_code", "")).strip()
        match = None

        # 1. JANコード完全一致
        if jan and len(jan) == 13:
            m = product_master[product_master["JANコード"] == jan]
            if len(m) > 0:
                match = m.iloc[0]

        # 2. 商品名ファジーマッチ（JANで見つからない場合）
        if match is None:
            from difflib import SequenceMatcher
            name = item.get("product_name", "")
            best_ratio = 0
            for _, row in product_master.iterrows():
                ratio = SequenceMatcher(None, name, row["商品名"]).ratio()
                if ratio > best_ratio and ratio >= 0.6:
                    best_ratio = ratio
                    match = row

        if match is not None:
            matched.append({
                "matched": True,
                "ocr_name": item.get("product_name", ""),
                "master_name": match["商品名"],
                "jan": match["JANコード"] if pd.notna(match["JANコード"]) else "",
                "code": match.get("商品コード", ""),
                "spec": match["規格"],
                "pack": match["配送荷姿"],
                "unit_price": float(match["1袋単価"]),
                "cs_price": float(match["CS単価"]),
                "output_dest": match["出力先"],
                "quantity": int(item.get("quantity_cs", 0)),
                "amount": int(item.get("quantity_cs", 0)) * float(match["CS単価"]),
            })
        else:
            matched.append({
                "matched": False,
                "ocr_name": item.get("product_name", ""),
                "jan": jan,
                "quantity": int(item.get("quantity_cs", 0)),
            })

    return matched


def match_ddc(dest_name, ddc_master):
    """納品先名をDDCマスタとマッチング"""
    from difflib import SequenceMatcher

    best_ratio = 0
    best_match = None
    for _, row in ddc_master.iterrows():
        ratio = SequenceMatcher(None, dest_name, row["納品先名"]).ratio()
        if ratio > best_ratio and ratio >= 0.5:
            best_ratio = ratio
            best_match = row

    if best_match is not None:
        return {
            "matched": True,
            "name": best_match["納品先名"],
            "postal": str(best_match.get("郵便番号", "")),
            "address": str(best_match.get("住所", "")),
            "tel": str(best_match.get("電話番号", "")),
            "fax": str(best_match.get("FAX番号", "")),
            "time": str(best_match.get("入荷時間", "")),
            "berse": str(best_match.get("バース予約", "無")),
            "palette": str(best_match.get("パレット条件", "")),
            "jpr": str(best_match.get("JPRコード", "")),
            "method": str(best_match.get("納品方法", "")),
        }
    else:
        return {"matched": False, "name": dest_name}


import pandas as pd  # needed for match_products
