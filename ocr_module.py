"""OCR モジュール - Claude Vision APIでFAX PDFを読み取り"""
import os
import json
import base64
import anthropic
import streamlit as st


def get_api_key():
    """APIキーを取得（Streamlit Secrets -> 環境変数の順）"""
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

    prompt = (
        "このFAX発注書の画像を読み取り、以下のJSON形式で情報を抽出してください。\n\n"
        "{\n"
        '  "order_no": "オーダーNO/発注NO",\n'
        '  "order_date": "発注日 (YYYY-MM-DD)",\n'
        '  "delivery_date": "納品日/入荷日 (YYYY-MM-DD)",\n'
        '  "sender": "発注者/発注元の会社名",\n'
        '  "sender_contact": "担当者名",\n'
        '  "sender_fax": "発注元FAX番号",\n'
        '  "delivery_dest": "納品先/入荷場所の名称",\n'
        '  "delivery_address": "納品先住所",\n'
        '  "delivery_tel": "納品先電話番号",\n'
        '  "delivery_fax": "納品先FAX番号",\n'
        '  "items": [\n'
        "    {\n"
        '      "product_name": "商品名",\n'
        '      "jan_code": "JANコード（4589570で始まる13桁）",\n'
        '      "product_code": "商品コード",\n'
        '      "quantity_cs": "ケース数（数値）",\n'
        '      "quantity_bara": "バラ数/入数（数値）",\n'
        '      "unit_price": "単価（数値、あれば）",\n'
        '      "amount": "金額（数値、あれば）",\n'
        '      "best_before": "賞味期限条件（あれば）"\n'
        "    }\n"
        "  ],\n"
        '  "notes": "備考・特記事項"\n'
        "}\n\n"
        "注意事項：\n"
        "- JANコードは4589570801で始まる13桁の数字です。必ず正確に読み取ってください。\n"
        "- 数量はケース(CS)単位の数値で返してください。\n"
        "- 金額や単価がない場合は空文字にしてください。\n"
        "- JSONのみを返してください。説明文は不要です。"
    )

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
        return {"error": "OCRエラー: {}".format(str(e))}


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
    """納品先名をDDCマスタとマッチング（従来互換）"""
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


def match_ddc_candidates(dest_name, ddc_master, max_candidates=5):
    """
    納品先名をDDCマスタとマッチング（候補付き）

    完全一致 -> 自動マッチ
    部分一致・類似度 -> 候補リスト返却（最大max_candidates件）
    該当なし -> 空リスト

    Returns:
        dict:
            exact_match (bool): 完全一致があったか
            matched_row (dict|None): 完全一致のデータ
            candidates (list[dict]): 候補リスト（類似度順）
              各要素: {"name": str, "score": float, "row_data": dict}
    """
    from difflib import SequenceMatcher

    dest_name_clean = dest_name.strip()
    if not dest_name_clean:
        return {"exact_match": False, "matched_row": None, "candidates": []}

    # --- 1. 完全一致チェック ---
    exact = ddc_master[ddc_master["納品先名"] == dest_name_clean]
    if len(exact) > 0:
        row = exact.iloc[0]
        return {
            "exact_match": True,
            "matched_row": _row_to_dict(row),
            "candidates": [],
        }

    # --- 2. 部分一致 + 類似度でスコアリング ---
    scored = []
    for _, row in ddc_master.iterrows():
        master_name = str(row["納品先名"])
        score = 0.0

        # 部分一致ボーナス（どちらかが含まれている場合）
        # 例：「福岡物流センター」が「山星屋福岡物流センター」に含まれる -> 0.85
        if dest_name_clean in master_name or master_name in dest_name_clean:
            score = 0.85

        # SequenceMatcherの類似度
        seq_ratio = SequenceMatcher(None, dest_name_clean, master_name).ratio()

        # 最終スコア = 部分一致とSequenceMatcherの大きい方
        final_score = max(score, seq_ratio)

        if final_score >= 0.3:  # 最低閾値
            scored.append({
                "name": master_name,
                "score": final_score,
                "row_data": _row_to_dict(row),
            })

    # スコア降順でソート -> 上位N件
    scored.sort(key=lambda x: x["score"], reverse=True)
    candidates = scored[:max_candidates]

    return {
        "exact_match": False,
        "matched_row": None,
        "candidates": candidates,
    }


def _row_to_dict(row):
    """DDCマスタの1行をdictに変換"""
    return {
        "name": str(row.get("納品先名", "")),
        "postal": str(row.get("郵便番号", "")),
        "address": str(row.get("住所", "")),
        "tel": str(row.get("電話番号", "")),
        "fax": str(row.get("FAX番号", "")),
        "time": str(row.get("入荷時間", "")),
        "berse": str(row.get("バース予約", "無")),
        "palette": str(row.get("パレット条件", "")),
        "jpr": str(row.get("JPRコード", "")),
        "method": str(row.get("納品方法", "")),
    }


import pandas as pd  # needed for match_products
