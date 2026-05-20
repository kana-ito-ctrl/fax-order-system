"""Google Drive FAX自動処理スクリプト（v2: Supabase連携対応）

Google Driveの共有フォルダからFAX PDFを読み取り、OCR処理・出荷判定・PDF生成を行う。
OCR結果を Supabase (fax_orders_scm / fax_order_items_scm) に保存する。

【新フォルダ構成（2026-04 〜）】
  _自動取込/新着/     ← GAS が自動保存する新着PDF（監視対象）
  _自動取込/処理済み/ ← 処理完了後に自動で移動
  _自動取込/エラー/   ← OCRエラー・商品マッチ不能を隔離

【旧フォルダ構成（レガシーモード、既に未使用）】
  05_卸納品実績/n8n/印刷前/
  05_卸納品実績/n8n/印刷前/修正済み/
  05_卸納品実績/n8n/印刷前/処理済み/

出力先：
  受注管理/output/     ← 発注書PDF（シルビア・ハルナ）
  受注管理/受注一覧.csv ← 受注一覧（Google Sheetsで開く）
  Supabase fax_orders_scm / fax_order_items_scm

使い方：
  python drive_processor.py                  # 新フォルダ(_自動取込/新着)を処理
  python drive_processor.py --folder 修正済み  # レガシーモード: 修正済みフォルダを処理
  python drive_processor.py --mode legacy    # レガシーモード: 旧フォルダを監視
  python drive_processor.py --dry-run        # Supabase書込をスキップ（テスト用）
  python drive_processor.py --limit 1        # 最新1件のみ処理（動作確認用）
"""
import os
import sys
import json
import shutil
import csv
import argparse
import time
from datetime import datetime

if sys.stdout.encoding and sys.stdout.encoding.lower() in ('cp932', 'shift_jis', 'mbcs'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ocr_module import process_fax_pdf
from pdf_generator import gen_sylvia_pdf, gen_haruna_pdf
from shipping_judge import judge_shipping, judge_warehouse, check_lot_alert, check_lead_time_alert
from supabase_client import insert_fax_order, check_fax_order_exists

# ==================== フォルダ設定 ====================
# 新フォルダ（2026-04〜）: GAS経由で自動取込。
#   _自動取込/新着/     : 監視対象
#   _自動取込/処理済み/ : 処理後移動先
#   _自動取込/エラー/   : エラー隔離
AUTO_BASE_FOLDER     = r"G:\共有ドライブ\TWO\SCM\31_受発注確認\受注受信\_自動取込"
NEW_INPUT_FOLDER     = os.path.join(AUTO_BASE_FOLDER, "新着")
NEW_PROCESSED_FOLDER = os.path.join(AUTO_BASE_FOLDER, "処理済み")
NEW_ERROR_FOLDER     = os.path.join(AUTO_BASE_FOLDER, "エラー")

# レガシーフォルダ（旧n8nフォルダ、現在未使用）
LEGACY_INPUT_FOLDER  = r"G:\共有ドライブ\TWO\SCM\05_卸納品実績\n8n\印刷前"
CORRECTED_SUBFOLDER  = "修正済み"   # 手書き訂正後のPDF格納先（レガシーモード）
PROCESSED_SUBFOLDER  = "処理済み"   # 処理完了後の移動先（レガシーモード）

# デフォルトで新フォルダを使う（main()で --mode により切替）
FAX_INPUT_FOLDER = NEW_INPUT_FOLDER

OUTPUT_FOLDER = r"G:\共有ドライブ\TWO\SCM\受注管理\output"
ORDER_LIST_CSV = r"G:\共有ドライブ\TWO\SCM\受注管理\受注一覧.csv"
# ======================================================

ORDER_LIST_HEADERS = [
    "処理日時", "入力元", "FAXファイル名", "ページ",
    "オーダーNO", "納品日", "発注元",
    "納品先(OCR)", "納品先(マスタ)", "DDCマッチ",
    "DDC候補1", "DDC候補2", "DDC候補3",
    "商品名", "CS数", "金額",
    "出荷区分", "倉庫", "ロットアラート", "ステータス",
    "シルビアPDF", "ハルナPDF", "備考",
]


def ensure_folders(mode: str = "new"):
    """必要なフォルダを作成する

    Args:
        mode: "new" (新フォルダ構成) or "legacy" (旧構成)
    """
    folders = [
        OUTPUT_FOLDER,
        os.path.dirname(ORDER_LIST_CSV),
    ]
    if mode == "legacy":
        folders.extend([
            os.path.join(LEGACY_INPUT_FOLDER, CORRECTED_SUBFOLDER),
            os.path.join(LEGACY_INPUT_FOLDER, PROCESSED_SUBFOLDER),
        ])
    else:
        folders.extend([
            NEW_INPUT_FOLDER,
            NEW_PROCESSED_FOLDER,
            NEW_ERROR_FOLDER,
        ])
    for folder in folders:
        os.makedirs(folder, exist_ok=True)


def get_fax_files(subfolder=None, days=None, limit=None, extensions=(".pdf", ".csv")):
    """処理対象のファイル一覧を取得する（更新日時が古い順）

    Args:
        subfolder: サブフォルダ名（例: "修正済み"）
        days: 直近N日以内のファイルのみ対象（例: 3）
        limit: 最大件数（例: 20）
        extensions: 拾う拡張子のタプル。デフォルトは PDF と CSV。
    """
    if subfolder:
        target_folder = os.path.join(FAX_INPUT_FOLDER, subfolder)
    else:
        target_folder = FAX_INPUT_FOLDER

    if not os.path.exists(target_folder):
        return []

    cutoff = None
    if days is not None:
        cutoff = time.time() - days * 86400

    files = []
    exts_lower = tuple(e.lower() for e in extensions)
    for f in os.listdir(target_folder):
        flower = f.lower()
        # .meta.json は除外
        if flower.endswith('.meta.json'):
            continue
        if flower.startswith('~'):
            continue
        if not flower.endswith(exts_lower):
            continue
        full_path = os.path.join(target_folder, f)
        if not os.path.isfile(full_path):
            continue
        if cutoff is not None and os.path.getmtime(full_path) < cutoff:
            continue
        files.append(full_path)

    files.sort(key=lambda x: os.path.getmtime(x))

    if limit is not None:
        files = files[-limit:]  # 最新N件

    return files


def init_order_list_csv():
    """受注一覧CSVが存在しない場合、ヘッダーを書き込んで初期化する"""
    if not os.path.exists(ORDER_LIST_CSV):
        with open(ORDER_LIST_CSV, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=ORDER_LIST_HEADERS)
            writer.writeheader()


def append_to_order_list(rows):
    """受注一覧CSVに行を追記する"""
    with open(ORDER_LIST_CSV, 'a', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=ORDER_LIST_HEADERS)
        writer.writerows(rows)


def read_meta_json(pdf_path: str) -> dict:
    """PDFと同じディレクトリにある .meta.json サイドカーファイルを読み取る。

    GASが保存したGmailメタ情報（gmail_message_id / from / subject / received_at 等）を返す。
    見つからない場合は空dict。
    """
    meta_path = os.path.splitext(pdf_path)[0] + ".meta.json"
    if not os.path.exists(meta_path):
        return {}
    try:
        with open(meta_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"  [meta.json] 読み取り失敗 {os.path.basename(meta_path)}: {e}")
        return {}


def _move_pdf_with_meta(pdf_path: str, dest_folder: str) -> str:
    """PDFと同名の .meta.json サイドカーも一緒に移動する。衝突時はタイムスタンプ付与。"""
    os.makedirs(dest_folder, exist_ok=True)
    base_name = os.path.basename(pdf_path)
    dest = os.path.join(dest_folder, base_name)
    if os.path.exists(dest):
        name, ext = os.path.splitext(base_name)
        dest = os.path.join(dest_folder, f"{name}_{datetime.now().strftime('%H%M%S')}{ext}")
    shutil.move(pdf_path, dest)

    # サイドカーも同じ命名規則で移動（存在すれば）
    meta_src = os.path.splitext(pdf_path)[0] + ".meta.json"
    if os.path.exists(meta_src):
        meta_dest = os.path.splitext(dest)[0] + ".meta.json"
        try:
            shutil.move(meta_src, meta_dest)
        except OSError:
            pass  # サイドカー移動失敗は致命的ではない
    return dest


def move_to_processed(pdf_path: str, mode: str = "new") -> str:
    """処理済みPDFを処理済みフォルダへ移動する（.meta.json も一緒に）"""
    if mode == "legacy":
        processed_folder = os.path.join(LEGACY_INPUT_FOLDER, PROCESSED_SUBFOLDER)
    else:
        processed_folder = NEW_PROCESSED_FOLDER
    return _move_pdf_with_meta(pdf_path, processed_folder)


def move_to_error(pdf_path: str) -> str:
    """エラーPDFを エラーフォルダへ移動する（.meta.json も一緒に）"""
    return _move_pdf_with_meta(pdf_path, NEW_ERROR_FOLDER)


def build_row(ocr, ddc, item, page_num, pdf_name, output_dest, total_cs,
              sylvia_pdf_name="", haruna_pdf_name="", remarks="", source_label="FAX"):
    """受注一覧CSV用の1行データを生成する"""
    alert, alert_msg = check_lot_alert(output_dest, total_cs)
    shipping = judge_shipping(output_dest, total_cs)
    warehouse = judge_warehouse(output_dest, total_cs)

    ddc_matched = ddc.get("matched", False)
    candidates = ddc.get("candidates", [])

    def cand_cell(idx):
        if idx < len(candidates):
            c = candidates[idx]
            return f"{c['name']} ({c['score']:.0%})"
        return ""

    # リードタイムアラート（直送・自社倉庫ともに評価）
    lead_alert, lead_msg = (False, "")
    order_date_str = datetime.now().strftime("%Y-%m-%d")
    delivery_date_str = ocr.get("delivery_date", "")
    lead_alert, lead_msg = check_lead_time_alert(order_date_str, delivery_date_str, output_dest, shipping)

    status_flags = []
    if alert:
        status_flags.append("ロット未満")
    if lead_alert:
        status_flags.append("リードタイム")
    # DDC未マッチは直送（シルビア・ハルナ）の場合のみアラート対象
    if not ddc_matched and shipping == "直送":
        status_flags.append("DDC未マッチ")
    elif ddc.get("low_confidence") and shipping == "直送":
        status_flags.append("DDC低信頼度")
    if not item.get("matched", True):
        status_flags.append("マスタ外商品")
    status = ("⚠️要確認: " + "/".join(status_flags)) if status_flags else "✅完了"

    return {
        "処理日時": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "入力元": source_label,
        "FAXファイル名": pdf_name,
        "ページ": page_num,
        "オーダーNO": ocr.get("order_no", ""),
        "納品日": ocr.get("delivery_date", ""),
        "発注元": ocr.get("sender", ""),
        "納品先(OCR)": ocr.get("delivery_dest", ""),
        "納品先(マスタ)": ddc.get("name", "") if ddc_matched else "",
        "DDCマッチ": ("要確認" if ddc.get("low_confidence") else "OK") if ddc_matched else "NG",
        "DDC候補1": cand_cell(0),
        "DDC候補2": cand_cell(1),
        "DDC候補3": cand_cell(2),
        "商品名": item.get("master_name", item.get("ocr_name", "")),
        "CS数": item.get("quantity", total_cs),
        "金額": int(item.get("amount", 0)) if item.get("amount") else "",
        "出荷区分": shipping,
        "倉庫": warehouse,
        "ロットアラート": " / ".join(filter(None, [alert_msg if alert else "", lead_msg if lead_alert else ""])),
        "ステータス": status,
        "シルビアPDF": sylvia_pdf_name,
        "ハルナPDF": haruna_pdf_name,
        "備考": remarks or ocr.get("notes", ""),
    }


def _detect_error_flags(page_result: dict) -> list[str]:
    """OCR結果からエラー判定フラグを抽出する（発注書以外の可能性を検出）"""
    flags = []
    ocr = page_result.get("ocr_raw", {}) or {}
    matched = page_result.get("matched_items", []) or []
    warehouse_items = page_result.get("warehouse_items", []) or []

    all_items = matched + warehouse_items
    if not all_items:
        flags.append("商品なし")
    else:
        matched_count = sum(1 for i in all_items if i.get("matched"))
        if matched_count == 0:
            flags.append("商品マッチ率0%")
        elif len(all_items) > 0 and matched_count / len(all_items) < 0.5:
            flags.append("商品マッチ率50%未満")

    if not ocr.get("delivery_date"):
        flags.append("納品日未読")
    if not ocr.get("delivery_dest") and not page_result.get("ddc_match", {}).get("matched"):
        flags.append("納品先未読")

    return flags


def _build_supabase_payload(page_result: dict, pdf_path: str, meta: dict, source_label: str) -> tuple[dict, list[dict], list[str]]:
    """OCR結果をSupabase INSERT用のheader/items形式に変換する。

    Returns:
        (header_dict, items_list, status_flags)
    """
    ocr = page_result.get("ocr_raw", {}) or {}
    ddc = page_result.get("ddc_match", {}) or {}
    ddc_matched = bool(ddc.get("matched"))

    # warehouse_items は matched_items の自社倉庫商品サブセット。
    # 加算すると重複INSERTになるため matched_items のみ使用する。（バグ修正 2026-05-01）
    matched = page_result.get("matched_items", []) or []
    all_items = matched

    flags = _detect_error_flags(page_result)
    status = "error" if ("商品なし" in flags or "商品マッチ率0%" in flags) else "draft"

    # 出荷倉庫判定（matched_itemsから最多出力先を採用）
    output_dests = [i.get("output_dest") for i in matched if i.get("output_dest")]
    primary_dest = max(set(output_dests), key=output_dests.count) if output_dests else ""

    total_cs = sum((i.get("quantity") or 0) for i in matched)
    shipping_type = judge_shipping(primary_dest, total_cs) if primary_dest else ""
    warehouse = judge_warehouse(primary_dest, total_cs) if primary_dest else ""

    grand_total = 0
    for i in all_items:
        amt = i.get("amount")
        if amt:
            try:
                grand_total += float(amt)
            except (TypeError, ValueError):
                pass

    # 納品日をISO形式に（YYYY-MM-DD）
    def _to_date(s):
        if not s:
            return None
        s = str(s).strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
            try:
                return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return None

    header = {
        "trade_id_system": None,
        "slip_number": ocr.get("order_no") or None,
        "source_channel": source_label,
        "order_type": "wholesale",
        "status": status,
        "order_date": datetime.now().strftime("%Y-%m-%d"),
        "delivery_date": _to_date(ocr.get("delivery_date")),
        "ship_date": None,
        "slip_date": None,
        "partner_code": None,
        "partner_name": (ocr.get("sender") or "")[:200] or None,
        "delivery_location_code": ddc.get("code") if ddc_matched else None,
        "delivery_location_name": (ddc.get("name") if ddc_matched else ocr.get("delivery_dest")) or None,
        "delivery_location_address": ddc.get("address") if ddc_matched else None,
        "grand_total": grand_total if grand_total else None,
        "notes": ocr.get("notes") or None,
        "source_file_path": pdf_path,
        "source_file_name": os.path.basename(pdf_path),
        "source_email_id": meta.get("gmail_message_id") or None,
        "source_email_from": meta.get("from") or None,
        "source_email_received_at": meta.get("received_at") or None,
        "ocr_page_no": page_result.get("page"),
        "ocr_raw": ocr,
        "ocr_confidence": None,
        "ddc_matched": ddc_matched,
        "ddc_match_candidates": ddc.get("candidates") or None,
        "warehouse": warehouse or None,
        "shipping_type": shipping_type or None,
        "status_flags": flags or None,
    }

    items = []
    for idx, i in enumerate(all_items, start=1):
        qty = i.get("quantity")
        try:
            qty = float(qty) if qty is not None else None
        except (TypeError, ValueError):
            qty = None
        items.append({
            "slip_detail_id_system": str(idx),
            "line_no": idx,
            "product_code": i.get("code") or None,
            "product_code_raw": i.get("code_raw") or i.get("ocr_name") or None,
            "product_name": i.get("master_name") or i.get("ocr_name") or None,
            "spec": i.get("spec") or None,
            "quantity": qty,
            "unit": i.get("unit") or "CS",
            "unit_price": i.get("cs_price") or i.get("unit_price") or None,
            "amount": i.get("amount") or None,
            "tax": 0,
            "subtotal": i.get("amount") or None,
            "tax_category": "",
            "tax_type": "外税",
            "jan_code": i.get("jan") or None,
            "product_master_matched": bool(i.get("matched")),
            "ocr_raw_name": i.get("ocr_name") or None,
            "ocr_confidence": None,
        })

    return header, items, flags


def save_to_supabase(results: list, pdf_path: str, source_label: str = "efax") -> dict:
    """OCR結果をページごとにSupabase（fax_orders_scm + fax_order_items_scm）へINSERT。

    Returns:
        {"inserted": int, "errors": int, "skipped": int, "error_pages": list}
    """
    stats = {"inserted": 0, "errors": 0, "skipped": 0, "error_pages": []}
    meta = read_meta_json(pdf_path)
    email_id = meta.get("gmail_message_id")

    for page_result in results:
        if "error" in page_result:
            stats["errors"] += 1
            stats["error_pages"].append(page_result.get("page"))
            continue

        page_num = page_result.get("page")
        # 重複チェック（同じメール×ページが既に入っていたらスキップ）
        if email_id and page_num is not None:
            if check_fax_order_exists(email_id, page_num):
                print(f"  [Supabase] 重複スキップ: page={page_num} email_id={email_id[:8]}")
                stats["skipped"] += 1
                continue

        header, items, flags = _build_supabase_payload(page_result, pdf_path, meta, source_label)
        order_id = insert_fax_order(header, items)
        if order_id:
            flag_label = f" (⚠️ {', '.join(flags)})" if flags else ""
            print(f"  [Supabase] INSERT成功 page={page_num} status={header['status']}{flag_label}")
            stats["inserted"] += 1
        else:
            print(f"  [Supabase] INSERT失敗 page={page_num}")
            stats["errors"] += 1
            stats["error_pages"].append(page_num)

    return stats


def process_file(pdf_path, staff_name="伊藤", remarks="", source_label="FAX", model="claude-sonnet-4-6"):
    """1件のFAX PDFを処理してCSV行リストと生成PDFパスを返す"""
    pdf_name = os.path.basename(pdf_path)
    print(f"\n処理中: {pdf_name}")

    results = process_fax_pdf(pdf_path, model=model)

    csv_rows = []
    generated_pdfs = []

    for page_result in results:
        if "error" in page_result:
            print(f"  Page {page_result['page']}: ERROR - {page_result['error']}")
            continue

        ocr = page_result["ocr_raw"]
        ddc = page_result["ddc_match"]
        page_num = page_result["page"]
        order_no = ocr.get("order_no", "")
        base_name = os.path.splitext(pdf_name)[0]
        now_date = datetime.now().strftime("%Y-%m-%d")

        ddc_matched = ddc.get("matched", False)
        dest_name = ddc.get("name", ocr.get("delivery_dest", "不明")) if ddc_matched else ocr.get("delivery_dest", "不明")

        sylvia_items = page_result["sylvia_items"]
        haruna_items = page_result["haruna_items"]
        unmatched_items = [i for i in page_result["matched_items"] if not i.get("matched")]

        sylvia_pdf_name = ""
        haruna_pdf_name = ""

        # --- シルビア処理 ---
        if sylvia_items:
            total_cs = sum(i["quantity"] for i in sylvia_items)
            shipping = judge_shipping("シルビア", total_cs)
            alert, alert_msg = check_lot_alert("シルビア", total_cs)

            if shipping == "直送":
                items_data = [{
                    "jan": i.get("jan", ""),
                    "name": i["master_name"],
                    "spec": i.get("spec", ""),
                    "pack": i.get("pack", ""),
                    "unit_price": int(i.get("unit_price", 0)),
                    "cs_price": int(i.get("cs_price", 0)),
                    "quantity": i["quantity"],
                    "amount": int(i.get("amount", 0)),
                } for i in sylvia_items]

                order_data = {
                    "order_no": order_no,
                    "order_date": now_date,
                    "delivery_date": ocr.get("delivery_date", ""),
                    "delivery_dest": dest_name,
                    "postal": ddc.get("postal", "") if ddc_matched else "",
                    "address": ddc.get("address", "") if ddc_matched else "",
                    "tel": ddc.get("tel", "") if ddc_matched else "",
                    "fax": ddc.get("fax", "") if ddc_matched else "",
                    "remarks": remarks,
                }
                pdf_buf = gen_sylvia_pdf(order_data, items_data, staff_name)
                sylvia_pdf_name = f"シルビア_{base_name}_p{page_num}.pdf"
                out_path = os.path.join(OUTPUT_FOLDER, sylvia_pdf_name)
                with open(out_path, "wb") as f:
                    f.write(pdf_buf.read())
                generated_pdfs.append(("シルビア", out_path))
                print(f"  → シルビアPDF生成: {sylvia_pdf_name} ({total_cs}CS / {shipping})")
            else:
                alert_label = f" ⚠️{alert_msg}" if alert else ""
                print(f"  → シルビア: {total_cs}CS → {shipping}{alert_label}")

            for i in sylvia_items:
                csv_rows.append(build_row(
                    ocr, ddc, i, page_num, pdf_name, "シルビア", total_cs,
                    sylvia_pdf_name=sylvia_pdf_name, remarks=remarks, source_label=source_label
                ))

        # --- ハルナ処理 ---
        if haruna_items:
            total_cs = sum(i["quantity"] for i in haruna_items)
            shipping = judge_shipping("ハルナ", total_cs)
            alert, alert_msg = check_lot_alert("ハルナ", total_cs)

            if shipping == "直送":
                order_data = {
                    "order_no": order_no,
                    "order_date": now_date,
                    "delivery_date": ocr.get("delivery_date", ""),
                    "delivery_dest": dest_name,
                    "quantity": total_cs,
                    "remarks": remarks,
                }
                ddc_data = {
                    "postal": ddc.get("postal", "") if ddc_matched else "",
                    "address": ddc.get("address", "") if ddc_matched else "",
                    "tel": ddc.get("tel", "") if ddc_matched else "",
                    "fax": ddc.get("fax", "") if ddc_matched else "",
                    "time": ddc.get("time", "") if ddc_matched else "",
                    "berse": ddc.get("berse", "無") if ddc_matched else "無",
                    "palette": ddc.get("palette", "") if ddc_matched else "",
                    "jpr": ddc.get("jpr", "") if ddc_matched else "",
                    "method": ddc.get("method", "") if ddc_matched else "",
                    "notes": ddc.get("notes", "") if ddc_matched else "",
                }
                pdf_buf = gen_haruna_pdf(order_data, ddc_data, staff_name)
                haruna_pdf_name = f"ハルナ_{base_name}_p{page_num}.pdf"
                out_path = os.path.join(OUTPUT_FOLDER, haruna_pdf_name)
                with open(out_path, "wb") as f:
                    f.write(pdf_buf.read())
                generated_pdfs.append(("ハルナ", out_path))
                print(f"  → ハルナPDF生成: {haruna_pdf_name} ({total_cs}CS / {shipping})")
            else:
                alert_label = f" ⚠️{alert_msg}" if alert else ""
                print(f"  → ハルナ: {total_cs}CS → {shipping}{alert_label}")

            # ハルナはCS合計で1行
            haruna_summary_item = {
                "matched": True,
                "master_name": "2Water Ceramide",
                "ocr_name": "2Water Ceramide",
                "quantity": total_cs,
                "amount": None,
            }
            csv_rows.append(build_row(
                ocr, ddc, haruna_summary_item, page_num, pdf_name, "ハルナ", total_cs,
                haruna_pdf_name=haruna_pdf_name, remarks=remarks, source_label=source_label
            ))

        # --- 自社倉庫商品（2Gummy, 2Energy等）---
        warehouse_items = page_result.get("warehouse_items", [])
        if warehouse_items:
            for i in warehouse_items:
                total_cs = i.get("quantity", 0)
                csv_rows.append(build_row(
                    ocr, ddc, i, page_num, pdf_name, "自社倉庫", total_cs,
                    remarks=remarks, source_label=source_label
                ))
                print(f"  → 自社倉庫: {i.get('master_name', '?')} x {total_cs}CS")

        # --- マスタ外商品 ---
        for i in unmatched_items:
            csv_rows.append(build_row(
                ocr, ddc, i, page_num, pdf_name, "対象外", 0,
                remarks=remarks, source_label=source_label
            ))
            print(f"  → マスタ外: {i.get('ocr_name', '?')} x {i.get('quantity', '?')}CS")

    return csv_rows, generated_pdfs


def print_summary(total_files, all_rows, elapsed):
    """処理結果サマリーを表示する"""
    alerts = [r for r in all_rows if "⚠️" in r.get("ステータス", "")]
    direct = [r for r in all_rows if r.get("出荷区分") == "直送"]
    warehouse = [r for r in all_rows if r.get("出荷区分") == "自社倉庫"]

    print(f"\n{'='*55}")
    print(f"完了: {total_files}件のFAXを処理 ({elapsed:.1f}秒)")
    print(f"  直送: {len(direct)}件  自社倉庫: {len(warehouse)}件")
    if alerts:
        print(f"\n⚠️ 要確認 {len(alerts)}件:")
        for r in alerts:
            print(f"   {r['FAXファイル名']} p{r['ページ']} | {r['納品先(OCR)']} | {r['ステータス']}")
    print(f"\n受注一覧: {ORDER_LIST_CSV}")
    print(f"出力PDF: {OUTPUT_FOLDER}")
    print(f"{'='*55}")


def run_import(
    limit: int | None = None,
    days: int | None = None,
    folder: str | None = None,
    mode: str = "new",
    no_supabase: bool = False,
    source_channel: str = "efax",
    model: str = "claude-sonnet-4-6",
    staff: str = "伊藤",
    remarks: str = "",
    print_log: bool = True,
    progress_cb=None,
) -> dict:
    """新着フォルダから対象ファイルを処理し、Supabase登録までを行う。

    Web UI からも main() からも呼べる共通関数。

    Args:
        progress_cb: 進捗通知コールバック。state dict を受け取る。
            state = {"phase": "scanning"|"processing"|"done",
                     "total": int, "current": int, "current_filename": str,
                     "inserted": int, "skipped": int, "errors": int,
                     "moved_to_error": int}

    Returns: 統計情報の dict（処理件数、Supabase登録件数等）
    """
    global FAX_INPUT_FOLDER
    if mode == "legacy":
        FAX_INPUT_FOLDER = LEGACY_INPUT_FOLDER
    else:
        FAX_INPUT_FOLDER = NEW_INPUT_FOLDER

    ensure_folders(mode=mode)
    init_order_list_csv()

    def _notify(state):
        if progress_cb:
            try:
                progress_cb(state)
            except Exception as e:
                if print_log:
                    print(f"  [progress_cb] {e}")

    _notify({"phase": "scanning", "total": 0, "current": 0,
             "current_filename": "", "inserted": 0, "skipped": 0,
             "errors": 0, "moved_to_error": 0})

    pdf_files = get_fax_files(subfolder=folder, days=days, limit=limit)

    if not pdf_files:
        if print_log:
            folder_label = folder if folder else ("新着" if mode == "new" else "印刷前")
            print(f"処理対象のPDFがありません: {FAX_INPUT_FOLDER}\\{folder_label}")
        result = {
            "status": "no_files",
            "processed": 0,
            "sb_inserted": 0,
            "sb_errors": 0,
            "sb_skipped": 0,
            "error_files": 0,
            "elapsed_sec": 0,
            "files": [],
        }
        _notify({"phase": "done", "total": 0, "current": 0,
                 "current_filename": "", "inserted": 0, "skipped": 0,
                 "errors": 0, "moved_to_error": 0})
        return result

    if print_log:
        mode_label = "新フォルダ(_自動取込)" if mode == "new" else "レガシー(n8n)"
        folder_label = (
            f"{folder}（手書き訂正）" if folder
            else ("新着" if mode == "new" else "印刷前（新着）")
        )
        print(f"処理対象: {len(pdf_files)}件 [{mode_label} / {folder_label}]")
        if no_supabase:
            print(f"  ⚠️  --no-supabase: Supabase書込はスキップします")

    all_rows = []
    start = time.time()
    sb_total = {"inserted": 0, "errors": 0, "skipped": 0}
    moved_to_error_count = 0
    files_log = []

    for idx, file_path in enumerate(pdf_files, start=1):
        fname = os.path.basename(file_path)
        _notify({"phase": "processing", "total": len(pdf_files), "current": idx,
                 "current_filename": fname, "inserted": sb_total["inserted"],
                 "skipped": sb_total["skipped"], "errors": sb_total["errors"],
                 "moved_to_error": moved_to_error_count})

        ext = os.path.splitext(file_path)[1].lower()
        # 拡張子で source_channel を決定（meta.json の source_channel を優先）
        meta = read_meta_json(file_path)
        meta_channel = meta.get("source_channel")
        if meta_channel:
            file_source_channel = meta_channel
        elif ext == ".csv":
            file_source_channel = "paltac"  # CSV は PALTAC とみなす（インフォマートは Supabase直結のためここに来ない）
        else:
            file_source_channel = source_channel  # PDF はデフォルトで efax

        # 解析処理: PDF→OCR / CSV→parse_paltac_csv
        try:
            if ext == ".pdf":
                results = process_fax_pdf(file_path, model=model)
            elif ext == ".csv":
                from process_fax import parse_paltac_csv as _parse_paltac_csv
                with open(file_path, "rb") as f:
                    csv_bytes = f.read()
                results = _parse_paltac_csv(csv_bytes, fname)
                if not results:
                    print(f"\n⚠️ CSV解析: 取込対象データなし（届先区分=発注 の行が0件）: {fname}")
                    if mode == "new":
                        moved = move_to_error(file_path)
                        print(f"  → エラーフォルダへ移動: {os.path.basename(moved)}")
                        moved_to_error_count += 1
                    files_log.append({"file": fname, "status": "csv_no_data"})
                    continue
            else:
                print(f"\n⚠️ 未対応拡張子をスキップ: {fname}")
                files_log.append({"file": fname, "status": "unsupported_ext"})
                continue
        except Exception as e:
            print(f"\n❌ 解析失敗: {fname}: {e}")
            if mode == "new":
                moved = move_to_error(file_path)
                print(f"  → エラーフォルダへ移動: {os.path.basename(moved)}")
                moved_to_error_count += 1
            files_log.append({"file": fname, "status": "parse_error", "error": str(e)})
            continue

        # Phase 4 (D1): drive_processor では発注書PDF生成しない。
        # PDF/CSV/受注一覧.csv の生成は確認アプリ（/api/confirm）で確定時に行う。
        # 旧 _process_results() の呼び出しは削除済み（関数定義は当面残す＝レガシー参照用）。

        # Supabase書込（--no-supabase でスキップ）
        file_sb_stats = None
        if not no_supabase and mode == "new":
            file_sb_stats = save_to_supabase(results, file_path, source_label=file_source_channel)
            sb_total["inserted"] += file_sb_stats["inserted"]
            sb_total["errors"] += file_sb_stats["errors"]
            sb_total["skipped"] += file_sb_stats["skipped"]
            # 全ページがエラー扱いならエラーフォルダへ移動
            total_pages = len([r for r in results if "error" not in r])
            if total_pages > 0 and file_sb_stats["errors"] == total_pages and file_sb_stats["inserted"] == 0:
                moved = move_to_error(file_path)
                print(f"  → エラーフォルダへ移動: {os.path.basename(moved)}")
                moved_to_error_count += 1
                files_log.append({"file": fname, "status": "supabase_all_errors",
                                  "stats": file_sb_stats})
                continue

        # 処理済みへ移動
        moved = move_to_processed(file_path, mode=mode)
        print(f"  → 処理済みへ移動: {os.path.basename(moved)}")
        files_log.append({"file": fname, "status": "ok", "stats": file_sb_stats})

    # Phase 4 (D1): _process_results() を呼ばないので all_rows は常に空。
    # 受注一覧.csv への追記は Phase 4 確認アプリで実施する。

    elapsed = time.time() - start
    if print_log:
        # PDF件数・処理結果の簡易サマリ（旧 print_summary は all_rows 前提だったので簡略化）
        print(f"\n{'='*55}")
        print(f"完了: {len(pdf_files)}件のFAXを処理 ({elapsed:.1f}秒)")
        print(f"  ※ 発注書PDF/CSV はSupabase確定時に生成されます（Phase 4）")
        print(f"{'='*55}")
    if not no_supabase and mode == "new":
        print(f"\nSupabase: 登録{sb_total['inserted']}件 / 重複スキップ{sb_total['skipped']}件 / エラー{sb_total['errors']}件")
        if moved_to_error_count:
            print(f"エラーフォルダ移動: {moved_to_error_count}件")

    result = {
        "status": "ok",
        "processed": len(pdf_files),
        "sb_inserted": sb_total["inserted"],
        "sb_errors": sb_total["errors"],
        "sb_skipped": sb_total["skipped"],
        "error_files": moved_to_error_count,
        "elapsed_sec": round(elapsed, 1),
        "files": files_log,
    }
    _notify({"phase": "done", "total": len(pdf_files), "current": len(pdf_files),
             "current_filename": "", "inserted": sb_total["inserted"],
             "skipped": sb_total["skipped"], "errors": sb_total["errors"],
             "moved_to_error": moved_to_error_count})
    return result


def main():
    """CLIエントリーポイント。argparseで引数をパースして run_import() を呼ぶ。"""
    parser = argparse.ArgumentParser(description="Google Drive FAX自動処理")
    parser.add_argument("--mode", choices=["new", "legacy"], default="new",
                        help="new: _自動取込/新着 を処理 / legacy: 旧n8nフォルダを処理")
    parser.add_argument("--folder", default=None,
                        help="サブフォルダ名を指定（レガシーで '修正済み' 等）")
    parser.add_argument("--days", type=int, default=None,
                        help="直近N日以内のファイルのみ対象")
    parser.add_argument("--limit", type=int, default=None,
                        help="処理する最大件数")
    parser.add_argument("--no-supabase", action="store_true",
                        help="Supabase書込をスキップ（テスト用）")
    parser.add_argument("--source-channel", default="efax",
                        help="PDFのデフォルトsource_channel (default: efax)")
    parser.add_argument("--model", default="claude-sonnet-4-6",
                        help="OCR用モデル名")
    parser.add_argument("--staff", default="伊藤", help="発注書PDFの担当者名")
    parser.add_argument("--remarks", default="", help="発注書の備考")
    args = parser.parse_args()

    run_import(
        limit=args.limit,
        days=args.days,
        folder=args.folder,
        mode=args.mode,
        no_supabase=args.no_supabase,
        source_channel=args.source_channel,
        model=args.model,
        staff=args.staff,
        remarks=args.remarks,
    )


def _process_results(results, pdf_path, staff_name="伊藤", remarks="", model="claude-sonnet-4-6"):
    """既存 process_file() の OCR後ロジックを、results 入力版として切り出した薄いラッパー。

    旧 process_file(pdf_path) は内部で OCR を再実行してしまうため、
    main() で既に取得した results を再利用するためのラッパー。
    """
    # 旧 process_file() と同じ後処理を実行するため、process_file の内部で process_fax_pdf を呼ぶ箇所を
    # パススルーできるよう results を渡せるヘルパーとして残す。直接 process_file() は使わない。
    # 実装は process_file の中身をそのまま再利用したいが、巨大なので既存関数を利用してOCRだけ冪等にする。
    # → シンプルに process_file を呼び、OCRを再実行する代わりに results をキャッシュするような実装は複雑なので、
    #   今回は OCR を1回で済ませるためにこの関数で process_file と同じ処理を再現する形にする。
    from datetime import date as _date
    csv_rows = []
    generated_pdfs = []
    pdf_name = os.path.basename(pdf_path)

    for page_result in results:
        if "error" in page_result:
            print(f"  Page {page_result['page']}: ERROR - {page_result['error']}")
            continue

        ocr = page_result["ocr_raw"]
        ddc = page_result["ddc_match"]
        page_num = page_result["page"]
        order_no = ocr.get("order_no", "")
        base_name = os.path.splitext(pdf_name)[0]
        now_date = datetime.now().strftime("%Y-%m-%d")

        ddc_matched = ddc.get("matched", False)
        dest_name = ddc.get("name", ocr.get("delivery_dest", "不明")) if ddc_matched else ocr.get("delivery_dest", "不明")

        sylvia_items = page_result.get("sylvia_items", [])
        haruna_items = page_result.get("haruna_items", [])
        unmatched_items = [i for i in page_result.get("matched_items", []) if not i.get("matched")]

        sylvia_pdf_name = ""
        haruna_pdf_name = ""

        # --- シルビア処理 ---
        if sylvia_items:
            total_cs = sum(i["quantity"] for i in sylvia_items)
            shipping = judge_shipping("シルビア", total_cs)
            alert, alert_msg = check_lot_alert("シルビア", total_cs)

            if shipping == "直送":
                items_data = [{
                    "jan": i.get("jan", ""),
                    "name": i["master_name"],
                    "spec": i.get("spec", ""),
                    "pack": i.get("pack", ""),
                    "unit_price": int(i.get("unit_price", 0)),
                    "cs_price": int(i.get("cs_price", 0)),
                    "quantity": i["quantity"],
                    "amount": int(i.get("amount", 0)),
                } for i in sylvia_items]

                order_data = {
                    "order_no": order_no,
                    "order_date": now_date,
                    "delivery_date": ocr.get("delivery_date", ""),
                    "delivery_dest": dest_name,
                    "postal": ddc.get("postal", "") if ddc_matched else "",
                    "address": ddc.get("address", "") if ddc_matched else "",
                    "tel": ddc.get("tel", "") if ddc_matched else "",
                    "fax": ddc.get("fax", "") if ddc_matched else "",
                    "remarks": remarks,
                }
                pdf_buf = gen_sylvia_pdf(order_data, items_data, staff_name)
                sylvia_pdf_name = f"シルビア_{base_name}_p{page_num}.pdf"
                out_path = os.path.join(OUTPUT_FOLDER, sylvia_pdf_name)
                with open(out_path, "wb") as f:
                    f.write(pdf_buf.read())
                generated_pdfs.append(("シルビア", out_path))
                print(f"  → シルビアPDF生成: {sylvia_pdf_name} ({total_cs}CS / {shipping})")
            else:
                alert_label = f" ⚠️{alert_msg}" if alert else ""
                print(f"  → シルビア: {total_cs}CS → {shipping}{alert_label}")

            for i in sylvia_items:
                csv_rows.append(build_row(
                    ocr, ddc, i, page_num, pdf_name, "シルビア", total_cs,
                    sylvia_pdf_name=sylvia_pdf_name, remarks=remarks, source_label="FAX"
                ))

        # --- ハルナ処理 ---
        if haruna_items:
            total_cs = sum(i["quantity"] for i in haruna_items)
            shipping = judge_shipping("ハルナ", total_cs)

            if shipping == "直送":
                order_data = {
                    "order_no": order_no,
                    "order_date": now_date,
                    "delivery_date": ocr.get("delivery_date", ""),
                    "delivery_dest": dest_name,
                    "quantity": total_cs,
                    "remarks": remarks,
                }
                ddc_data = {
                    "postal": ddc.get("postal", "") if ddc_matched else "",
                    "address": ddc.get("address", "") if ddc_matched else "",
                    "tel": ddc.get("tel", "") if ddc_matched else "",
                    "fax": ddc.get("fax", "") if ddc_matched else "",
                    "time": ddc.get("time", "") if ddc_matched else "",
                    "berse": ddc.get("berse", "無") if ddc_matched else "無",
                    "palette": ddc.get("palette", "") if ddc_matched else "",
                    "jpr": ddc.get("jpr", "") if ddc_matched else "",
                    "method": ddc.get("method", "") if ddc_matched else "",
                    "notes": ddc.get("notes", "") if ddc_matched else "",
                }
                pdf_buf = gen_haruna_pdf(order_data, ddc_data, staff_name)
                haruna_pdf_name = f"ハルナ_{base_name}_p{page_num}.pdf"
                out_path = os.path.join(OUTPUT_FOLDER, haruna_pdf_name)
                with open(out_path, "wb") as f:
                    f.write(pdf_buf.read())
                generated_pdfs.append(("ハルナ", out_path))
                print(f"  → ハルナPDF生成: {haruna_pdf_name} ({total_cs}CS / {shipping})")

            haruna_summary_item = {
                "matched": True,
                "master_name": "2Water Ceramide",
                "ocr_name": "2Water Ceramide",
                "quantity": total_cs,
                "amount": None,
            }
            csv_rows.append(build_row(
                ocr, ddc, haruna_summary_item, page_num, pdf_name, "ハルナ", total_cs,
                haruna_pdf_name=haruna_pdf_name, remarks=remarks, source_label="FAX"
            ))

        # --- 自社倉庫商品 ---
        for i in page_result.get("warehouse_items", []):
            total_cs = i.get("quantity", 0)
            csv_rows.append(build_row(
                ocr, ddc, i, page_num, pdf_name, "自社倉庫", total_cs,
                remarks=remarks, source_label="FAX"
            ))
            print(f"  → 自社倉庫: {i.get('master_name', '?')} x {total_cs}CS")

        # --- マスタ外商品 ---
        for i in unmatched_items:
            csv_rows.append(build_row(
                ocr, ddc, i, page_num, pdf_name, "対象外", 0,
                remarks=remarks, source_label="FAX"
            ))
            print(f"  → マスタ外: {i.get('ocr_name', '?')} x {i.get('quantity', '?')}CS")

    return csv_rows, generated_pdfs


if __name__ == "__main__":
    main()
