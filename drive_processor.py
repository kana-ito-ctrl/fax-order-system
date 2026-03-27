"""Google Drive FAX自動処理スクリプト

Google Driveの共有フォルダからFAX PDFを読み取り、OCR処理・出荷判定・PDF生成を行う。

フォルダ構成：
  印刷前/          ← 新着FAXが自動で格納される
  印刷前/修正済み/ ← 手書き訂正後のPDFを手動で格納
  印刷前/処理済み/ ← 処理完了後に自動で移動

出力先：
  受注管理/output/     ← 発注書PDF（シルビア・ハルナ）
  受注管理/受注一覧.csv ← 受注一覧（Google Sheetsで開く）

使い方：
  python drive_processor.py              # 印刷前フォルダを処理
  python drive_processor.py --folder 修正済み  # 手書き訂正ファイルを処理
  python drive_processor.py --staff 伊藤 --remarks "要冷蔵対応"
"""
import os
import sys
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

# ==================== フォルダ設定 ====================
FAX_INPUT_FOLDER = r"G:\共有ドライブ\TWO\SCM\05_卸納品実績\n8n\印刷前"
CORRECTED_SUBFOLDER = "修正済み"   # 手書き訂正後のPDF格納先
PROCESSED_SUBFOLDER = "処理済み"   # 処理完了後の移動先
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


def ensure_folders():
    """必要なフォルダを作成する"""
    for folder in [
        OUTPUT_FOLDER,
        os.path.join(FAX_INPUT_FOLDER, CORRECTED_SUBFOLDER),
        os.path.join(FAX_INPUT_FOLDER, PROCESSED_SUBFOLDER),
        os.path.dirname(ORDER_LIST_CSV),
    ]:
        os.makedirs(folder, exist_ok=True)


def get_fax_files(subfolder=None, days=None, limit=None):
    """処理対象のFAX PDFファイル一覧を取得する（更新日時が古い順）

    Args:
        subfolder: サブフォルダ名（例: "修正済み"）
        days: 直近N日以内のファイルのみ対象（例: 3）
        limit: 最大件数（例: 20）
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

    pdf_files = []
    for f in os.listdir(target_folder):
        if f.lower().endswith('.pdf') and not f.startswith('~'):
            full_path = os.path.join(target_folder, f)
            if os.path.isfile(full_path):
                if cutoff is None or os.path.getmtime(full_path) >= cutoff:
                    pdf_files.append(full_path)

    pdf_files.sort(key=lambda x: os.path.getmtime(x))

    if limit is not None:
        pdf_files = pdf_files[-limit:]  # 最新N件

    return pdf_files


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


def move_to_processed(pdf_path):
    """処理済みPDFを 印刷前/処理済み/ フォルダへ移動する"""
    processed_folder = os.path.join(FAX_INPUT_FOLDER, PROCESSED_SUBFOLDER)
    os.makedirs(processed_folder, exist_ok=True)
    dest = os.path.join(processed_folder, os.path.basename(pdf_path))
    if os.path.exists(dest):
        name, ext = os.path.splitext(os.path.basename(pdf_path))
        dest = os.path.join(processed_folder, f"{name}_{datetime.now().strftime('%H%M%S')}{ext}")
    shutil.move(pdf_path, dest)
    return dest


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


def main():
    parser = argparse.ArgumentParser(description="Google Drive FAX自動処理スクリプト")
    parser.add_argument("--staff", default="伊藤", help="担当者名 (default: 伊藤)")
    parser.add_argument("--remarks", default="", help="イレギュラーリクエスト/備考")
    parser.add_argument("--folder", default=None, choices=["修正済み"],
                        help="処理対象サブフォルダ（省略時: 印刷前）")
    parser.add_argument("--days", type=int, default=None,
                        help="直近N日以内のファイルのみ処理（例: --days 3）")
    parser.add_argument("--limit", type=int, default=None,
                        help="最新N件のみ処理（例: --limit 20）")
    parser.add_argument("--model", default="claude-sonnet-4-6", help="Claude model for OCR")
    args = parser.parse_args()

    ensure_folders()
    init_order_list_csv()

    pdf_files = get_fax_files(subfolder=args.folder, days=args.days, limit=args.limit)

    if not pdf_files:
        folder_label = f"印刷前/{args.folder}" if args.folder else "印刷前"
        print(f"処理対象のPDFがありません: {FAX_INPUT_FOLDER}\\{folder_label}")
        return

    folder_label = f"修正済み（手書き訂正）" if args.folder else "印刷前（新着）"
    print(f"処理対象: {len(pdf_files)}件 [{folder_label}]")

    all_rows = []
    start = time.time()

    for pdf_path in pdf_files:
        rows, pdfs = process_file(
            pdf_path,
            staff_name=args.staff,
            remarks=args.remarks,
            model=args.model,
        )
        all_rows.extend(rows)

        moved = move_to_processed(pdf_path)
        print(f"  → 処理済みへ移動: {os.path.basename(moved)}")

    if all_rows:
        append_to_order_list(all_rows)

    print_summary(len(pdf_files), all_rows, time.time() - start)


if __name__ == "__main__":
    main()
