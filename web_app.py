"""FAX受注処理 ブラウザUI

Usage:
    python web_app.py
    → ブラウザが自動で開きます (http://localhost:5000)
"""
import os
import sys
import json
import uuid
import base64
import threading
import webbrowser

# Windows terminal UTF-8 (ascii含む全非UTF-8環境を補正)
import io as _io
for _stream_name in ('stdout', 'stderr'):
    _stream = getattr(sys, _stream_name, None)
    if _stream and hasattr(_stream, 'buffer'):
        try:
            if (_stream.encoding or 'ascii').lower() not in ('utf-8', 'utf8'):
                setattr(sys, _stream_name, _io.TextIOWrapper(
                    _stream.buffer, encoding='utf-8', errors='replace', line_buffering=True))
        except Exception:
            pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, jsonify, send_file, render_template_string
from datetime import datetime, date

# アプリ本体ファイルの最終更新日時を取得（起動時に確定）
_APP_FILE = os.path.abspath(__file__)
APP_LAST_UPDATED = datetime.fromtimestamp(os.path.getmtime(_APP_FILE)).strftime("%Y-%m-%d %H:%M")
from ocr_module import (
    process_fax_pdf, pdf_to_images, load_ddc_master, load_product_master,
    match_ddc, match_product, load_staff, ocr_fax_page, normalize, normalize_company,
    _ddc_row_to_dict,
)
from process_fax import generate_pdfs, results_to_excel, results_to_csv, results_to_ne_csv, results_to_coola_csv, parse_paltac_csv, parse_infomart_csv, parse_smacla_csv, ensure_output_dir, OUTPUT_DIR
from pdf_generator import gen_sylvia_pdf, gen_haruna_pdf

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# In-memory session store (single-user local app)
sessions = {}

# TWO受注NO採番（YYMMDDNNN: 日付6桁+連番3桁）
_two_order_date = ""
_two_order_seq = 0


def generate_two_order_no():
    """TWO受注NOを自動採番する。YYMMDDNNN形式（9桁数字）"""
    global _two_order_date, _two_order_seq
    today = date.today().strftime("%y%m%d")
    if today != _two_order_date:
        _two_order_date = today
        _two_order_seq = 0
    _two_order_seq += 1
    return f"{today}{_two_order_seq:03d}"


# Pre-loaded master data (loaded once at startup)
_ddc_master_df = None
_product_master_df = None
_ddc_list_cache = None


def get_ddc_master():
    global _ddc_master_df
    if _ddc_master_df is None:
        _ddc_master_df = load_ddc_master()
    return _ddc_master_df


def get_product_master():
    global _product_master_df
    if _product_master_df is None:
        _product_master_df = load_product_master()
    return _product_master_df


def get_ddc_list():
    """Return DDC master as list of dicts for JSON API"""
    global _ddc_list_cache
    if _ddc_list_cache is None:
        df = get_ddc_master()
        _ddc_list_cache = []
        for _, row in df.iterrows():
            def s(v):
                v2 = str(v).strip() if v is not None else ""
                return "" if v2.lower() in ("nan", "none", "null") else v2
            _ddc_list_cache.append({
                "name": s(row.get("納品先名", "")),
                "nohinsaki_code": s(row.get("納品先コード", "")),
                "address": s(row.get("住所", "")),
                "tel": s(row.get("電話番号", "")),
                "postal": s(row.get("郵便番号", "")),
                "torihikisaki": s(row.get("取引先名", "")),
                "torihikisaki_code": s(row.get("取引先コード", "")),
            })
    return _ddc_list_cache


# ─── API Endpoints ───


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE, app_last_updated=APP_LAST_UPDATED)


@app.route("/pending")
def pending_orders_page():
    """保留中受注一覧ページ（read-only / Supabase fax_orders_scm から取得）"""
    return render_template_string(PENDING_ORDERS_TEMPLATE, app_last_updated=APP_LAST_UPDATED)


@app.route("/confirmed")
def confirmed_orders_page():
    """Phase 4: 確定済み受注一覧ページ（再編集可・Supabase fax_orders_scm から取得）"""
    return render_template_string(CONFIRMED_ORDERS_TEMPLATE, app_last_updated=APP_LAST_UPDATED)


@app.route("/api/confirmed_orders")
def api_confirmed_orders():
    """確定済み受注（status='confirmed'）の一覧をJSONで返す"""
    from supabase_client import fetch_confirmed_orders
    source = request.args.get("source") or None
    days_param = request.args.get("days")
    if days_param == "":
        days = None  # 空文字明示 = 全期間
    elif days_param:
        try:
            days = int(days_param)
        except ValueError:
            days = 90
    else:
        days = 90  # クエリ未指定時のデフォルト
    try:
        rows = fetch_confirmed_orders(limit=200, source_channel=source, days=days)
        return jsonify({"orders": rows, "count": len(rows)})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/pending_orders")
def api_pending_orders():
    """保留中受注（status='draft' or 'error'）の一覧をJSONで返す"""
    from supabase_client import fetch_pending_orders
    source = request.args.get("source") or None
    try:
        rows = fetch_pending_orders(limit=200, source_channel=source)
        return jsonify({"orders": rows, "count": len(rows)})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/pending_orders/<order_id>")
def api_pending_order_detail(order_id):
    """指定受注の詳細（ヘッダー＋明細）をJSONで返す"""
    from supabase_client import fetch_order_with_items
    try:
        data = fetch_order_with_items(order_id)
        if not data:
            return jsonify({"error": "受注が見つかりません"}), 404
        return jsonify(data)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/ddc_list")
def api_ddc_list():
    return jsonify(get_ddc_list())


@app.route("/api/bulk_csv_export", methods=["POST"])
def api_bulk_csv_export():
    """Phase 4 改修: 複数の確定済 fax_orders_scm レコードから NE/COOLA CSV を一括生成。

    Body: {"order_ids": [uuid, ...], "csv_type": "ne" | "coola"}
    Returns: {"success": True, "filename": "...", "url": "/api/download/...", "skipped": N}

    旧運用（複数受注を1つのCSVに）に近い、まとめアップロードフロー用。
    """
    data = request.get_json(silent=True) or {}
    order_ids = data.get("order_ids") or []
    csv_type = (data.get("csv_type") or "").lower()
    if not order_ids:
        return jsonify({"error": "order_ids が空です"}), 400
    if csv_type not in ("ne", "coola"):
        return jsonify({"error": "csv_type は 'ne' または 'coola'"}), 400

    from supabase_client import fetch_order_with_items
    pm_df = get_product_master()
    ddc_df = get_ddc_master()

    # 各 order_id を fetch して内部 results 形式に変換
    pm_jan_index = {}
    if pm_df is not None and "JANコード" in pm_df.columns:
        for _, prow in pm_df.iterrows():
            j = str(prow.get("JANコード") or "").strip()
            if j:
                pm_jan_index[j] = prow

    all_results = []
    skipped = 0
    for oid in order_ids:
        try:
            data_row = fetch_order_with_items(oid)
        except Exception:
            data_row = None
        if not data_row:
            skipped += 1
            continue
        header = data_row["header"]
        items = data_row["items"]

        matched_items = []
        for it in items:
            jan = str(it.get("jan_code") or "").strip()
            prow = pm_jan_index.get(jan)
            output_dest = str(prow.get("出力先") or "").strip() if prow is not None else ""
            case_qty = int(prow.get("入数") or 0) if prow is not None else 0
            cs_price = float(prow.get("CS単価") or 0) if prow is not None else float(it.get("unit_price") or 0)
            try:
                qty = int(float(it.get("quantity") or 0))
            except (TypeError, ValueError):
                qty = 0
            matched_items.append({
                "jan": jan,
                "code": it.get("product_code") or "",
                "master_name": it.get("product_name") or "",
                "ocr_name": it.get("ocr_raw_name") or it.get("product_name") or "",
                "spec": it.get("spec") or (str(prow.get("規格") or "") if prow is not None else ""),
                "pack": str(prow.get("配送荷姿") or "") if prow is not None else "",
                "quantity": qty,
                "unit": it.get("unit") or "CS",
                "unit_price": float(prow.get("1袋単価") or 0) if prow is not None else float(it.get("unit_price") or 0),
                "cs_price": cs_price,
                "amount": float(it.get("amount") or 0),
                "case_quantity": case_qty,
                "output_dest": output_dest,
                "matched": bool(it.get("product_master_matched")),
            })

        # DDC情報（住所・電話等）
        from ocr_module import _ddc_row_to_dict
        ddc_match = {
            "matched": bool(header.get("ddc_matched")),
            "name": header.get("delivery_location_name") or "",
            "code": header.get("delivery_location_code") or "",
            "address": header.get("delivery_location_address") or "",
            "postal": "", "tel": "", "fax": "",
            "time": "", "berse": "無", "palette": "", "jpr": "", "method": "",
        }
        nohinsaki_code = ddc_match["code"]
        if nohinsaki_code and ddc_df is not None and "納品先コード" in ddc_df.columns:
            mask = ddc_df["納品先コード"].astype(str) == str(nohinsaki_code)
            if mask.any():
                full_ddc = _ddc_row_to_dict(ddc_df[mask].iloc[0])
                for k in ("postal", "address", "tel", "fax", "time", "berse",
                          "palette", "jpr", "method"):
                    ddc_match[k] = full_ddc.get(k, ddc_match.get(k, ""))

        ocr_raw = {
            "order_no": header.get("slip_number") or "",
            "delivery_date": header.get("delivery_date") or "",
            "delivery_dest": ddc_match["name"],
            "sender": header.get("partner_name") or "",
            "notes": header.get("notes") or "",
            "items": [],
        }
        all_results.append({
            "page": header.get("ocr_page_no") or 1,
            "two_order_no": header.get("slip_number") or "",
            "ocr_raw": ocr_raw,
            "matched_items": matched_items,
            "ddc_match": ddc_match,
            # 確定時に保存された送料を反映
            "shipping_fee": int(header.get("shipping_fee") or 0) if not header.get("shipping_fee_two_burden") else 0,
            # warehouse_direct は CSV フィルタには使わなくなったため省略
        })

    if not all_results:
        return jsonify({"error": "対象レコードがありません",
                        "skipped": skipped}), 400

    from datetime import datetime as _dt
    timestamp = _dt.now().strftime("%Y%m%d_%H%M%S")

    try:
        ensure_output_dir()
        if csv_type == "ne":
            outfile_name = f"NE一括_{timestamp}_{len(all_results)}件.csv"
            outfile = results_to_ne_csv(all_results, outfile_name)
        else:  # coola
            outfile_name = f"COOLA一括_{timestamp}_{len(all_results)}件.csv"
            outfile = results_to_coola_csv(all_results, outfile_name)

        if not outfile:
            return jsonify({"error": "CSV生成失敗（対象データなし or 直送のみで COOLA出力対象が0件）",
                            "skipped": skipped, "processed": len(all_results)}), 400

        filename = os.path.basename(outfile)
        return jsonify({
            "success": True,
            "filename": filename,
            "url": f"/api/download/{filename}",
            "processed": len(all_results),
            "skipped": skipped,
            "csv_type": csv_type,
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"生成エラー: {repr(e)}"}), 500


@app.route("/api/calculate_shipping_fee", methods=["POST"])
def api_calculate_shipping_fee():
    """Phase 4: 送料計算プレビュー用 API（ページ毎個別計算）。

    Body: {pages: [{items, ddc_address, shipping_type}, ...]}
    Returns: {pages: [shipping_result, ...], total_fee_all_pages: int}
    """
    data = request.get_json(silent=True) or {}
    pages_in = data.get("pages") or []
    try:
        from shipping_fee_calculator import calculate_shipping_fee
        page_results = []
        total = 0
        for p in pages_in:
            items = p.get("items", []) or []
            ddc_address = p.get("ddc_address", "") or ""
            shipping_type = p.get("shipping_type")
            r = calculate_shipping_fee(items, ddc_address, shipping_type)
            page_results.append(r)
            total += int(r.get("total_fee") or 0)
        return jsonify({"pages": page_results, "total_fee_all_pages": total})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e), "pages": [], "total_fee_all_pages": 0}), 500


@app.route("/api/staff_list")
def api_staff_list():
    """Phase 4: 担当者プルダウン用の一覧を返す。

    data/staff.json から取得（Supabase staff_members は他チーム共用テーブルのため使わない）。
    """
    try:
        from ocr_module import load_staff
        sd = load_staff()
        names = [{"name": (s.get("name") or "").strip()}
                 for s in sd.get("staff", [])
                 if s.get("name")]
        return jsonify({"staff": names, "count": len(names)})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"staff": [], "count": 0, "error": str(e)}), 500


@app.route("/api/load_draft", methods=["POST"])
def api_load_draft():
    """Phase 4: Supabase の draft レコードを読み込んでセッションに展開する。

    Body: {"order_id": "<uuid>"}
    Returns: session_id ほか（/api/upload と同じ形式）+ from_draft / draft_id

    フロントは /?draft_id=xxx でこの API を叩いてセッション化し、
    既存の OCR結果 UI を「ドラフト編集モード」として再利用する。
    """
    data = request.get_json(silent=True) or {}
    order_id = data.get("order_id")
    if not order_id:
        return jsonify({"error": "order_id required"}), 400

    from supabase_client import fetch_order_with_items
    try:
        order_data = fetch_order_with_items(order_id)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Supabase fetch error: {e}"}), 500

    if not order_data:
        return jsonify({"error": "受注が見つかりません"}), 404

    header = order_data["header"]
    items = order_data["items"]

    # 商品マスタから出力先・1袋単価等を引き当て
    pm_df = get_product_master()
    pm_jan_index = {}
    if pm_df is not None and "JANコード" in pm_df.columns:
        for _, prow in pm_df.iterrows():
            j = str(prow.get("JANコード") or "").strip()
            if j:
                pm_jan_index[j] = prow

    matched_items = []
    for it in items:
        jan = str(it.get("jan_code") or "").strip()
        prow = pm_jan_index.get(jan)
        output_dest = str(prow.get("出力先") or "").strip() if prow is not None else ""
        case_qty = int(prow.get("入数") or 0) if prow is not None else 0
        bag_unit_price = float(prow.get("1袋単価") or 0) if prow is not None else 0.0
        cs_price = float(prow.get("CS単価") or 0) if prow is not None else float(it.get("unit_price") or 0)
        spec_pm = str(prow.get("規格") or "").strip() if prow is not None else ""
        pack_pm = str(prow.get("配送荷姿") or "").strip() if prow is not None else ""

        try:
            qty = int(float(it.get("quantity") or 0))
        except (TypeError, ValueError):
            qty = 0

        matched_items.append({
            "jan": jan,
            "code": it.get("product_code") or "",
            "code_raw": it.get("product_code_raw") or "",
            "master_name": it.get("product_name") or "",
            "ocr_name": it.get("ocr_raw_name") or it.get("product_name") or "",
            "spec": it.get("spec") or spec_pm,
            "pack": pack_pm,
            "quantity": qty,
            "unit": it.get("unit") or "CS",
            "unit_price": bag_unit_price or float(it.get("unit_price") or 0),
            "cs_price": cs_price,
            "amount": float(it.get("amount") or 0),
            "case_quantity": case_qty,
            "output_dest": output_dest,
            "matched": bool(it.get("product_master_matched")),
        })

    # DDC マスタから完全な納品先情報を引き当て（住所・電話・ハルナ条件等）
    ddc_match = {
        "matched": bool(header.get("ddc_matched")),
        "name": header.get("delivery_location_name") or "",
        "code": header.get("delivery_location_code") or "",
        "address": header.get("delivery_location_address") or "",
        "candidates": header.get("ddc_match_candidates") or [],
        "postal": "", "tel": "", "fax": "",
        "time": "", "berse": "無", "palette": "", "jpr": "", "method": "",
        "notes": "",
        "low_confidence": False,
    }
    nohinsaki_code = ddc_match["code"]
    if nohinsaki_code:
        from ocr_module import _ddc_row_to_dict
        ddc_df = get_ddc_master()
        if ddc_df is not None and "納品先コード" in ddc_df.columns:
            mask = ddc_df["納品先コード"].astype(str) == str(nohinsaki_code)
            if mask.any():
                full_ddc = _ddc_row_to_dict(ddc_df[mask].iloc[0])
                for k in ("postal", "address", "tel", "fax", "time", "berse",
                          "palette", "jpr", "method", "notes"):
                    ddc_match[k] = full_ddc.get(k, ddc_match.get(k, ""))

    # OCR raw 復元（保存済み ocr_raw を優先、フィールドはヘッダーから補完）
    ocr_raw_stored = header.get("ocr_raw") or {}
    ocr_raw = {
        "order_no": header.get("slip_number") or ocr_raw_stored.get("order_no") or "",
        "delivery_date": header.get("delivery_date") or ocr_raw_stored.get("delivery_date") or "",
        "delivery_dest": ddc_match["name"] or ocr_raw_stored.get("delivery_dest") or "",
        "sender": header.get("partner_name") or ocr_raw_stored.get("sender") or "",
        "notes": header.get("notes") or ocr_raw_stored.get("notes") or "",
        "items": ocr_raw_stored.get("items") or [],
    }

    # 出力先別グルーピング（Phase 4 確認画面で使う）
    sylvia_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "シルビア"]
    haruna_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "ハルナ"]

    page_result = {
        "page": header.get("ocr_page_no") or 1,
        "source": header.get("source_channel") or "draft",
        "two_order_no": generate_two_order_no(),
        "ocr_raw": ocr_raw,
        "matched_items": matched_items,
        "ddc_match": ddc_match,
        "sylvia_items": sylvia_items,
        "haruna_items": haruna_items,
        # Phase 4 ドラフト追跡情報
        "draft_id": order_id,
        "draft_status": header.get("status"),
        "draft_confirmation_count": header.get("confirmation_count") or 0,
        "draft_confirmed_at": header.get("confirmed_at"),
    }

    session_id = str(uuid.uuid4())[:8]
    sessions[session_id] = {
        "pdf_bytes": None,
        "pdf_name": header.get("source_file_name") or f"draft_{order_id[:8]}.pdf",
        "pages": [],
        "results": [page_result],
        "source": header.get("source_channel") or "draft",
        # Phase 4 ドラフト追跡情報
        "draft_id": order_id,
        "draft_status": header.get("status"),
    }

    return jsonify({
        "session_id": session_id,
        "filename": sessions[session_id]["pdf_name"],
        "page_count": 1,
        "source": sessions[session_id]["source"],
        "auto_results": True,
        "from_draft": True,
        "draft_id": order_id,
        "draft_status": header.get("status"),
        "draft_confirmation_count": header.get("confirmation_count") or 0,
        "draft_confirmed_at": header.get("confirmed_at"),
        "draft_slip_number": header.get("slip_number"),
    })


@app.route("/api/manual_entry", methods=["POST"])
def api_manual_entry():
    """手入力モード: 空の受注データを作成"""
    session_id = str(uuid.uuid4())[:8]
    results = [{
        "page": 1,
        "source": "manual",
        "two_order_no": generate_two_order_no(),
        "ocr_raw": {
            "order_no": "",
            "delivery_date": "",
            "delivery_dest": "",
            "sender": "",
            "notes": "",
            "items": [],
        },
        "matched_items": [],
        "ddc_match": {"matched": False, "name": "", "candidates": []},
        "sylvia_items": [],
        "haruna_items": [],
    }]
    sessions[session_id] = {
        "pdf_bytes": None,
        "pdf_name": "manual_entry.csv",
        "pages": [],
        "results": results,
        "source": "manual",
    }
    return jsonify({
        "session_id": session_id,
        "filename": "手入力",
        "page_count": 1,
        "source": "manual",
        "auto_results": True,
    })


@app.route("/api/product_list")
def api_product_list():
    """商品マスタの商品名リストを返す（プルダウン用）"""
    df = get_product_master()
    items = []
    for _, row in df.iterrows():
        name = str(row.get("商品名", "")).strip()
        jan = str(row.get("JANコード", "")).strip()
        code = str(row.get("商品コード", "")).strip()
        case_qty = int(row.get("入数", 0) or 0)
        output_dest = str(row.get("出力先", "")).strip()
        cs_price = float(row.get("CS単価", 0) or 0)
        spec = str(row.get("規格", "")).strip()
        pack = str(row.get("配送荷姿", "")).strip()
        unit_price = float(row.get("1袋単価", 0) or 0)
        if name:
            items.append({
                "name": name, "jan": jan, "code": code,
                "case_quantity": case_qty, "output_dest": output_dest,
                "cs_price": cs_price, "spec": spec, "pack": pack,
                "unit_price": unit_price,
            })
    return jsonify(items)


@app.route("/api/upload_multi_csv", methods=["POST"])
def api_upload_multi_csv():
    """複数のCSVファイルを一括アップロードして1セッションにまとめる"""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    all_results = []
    filenames = []
    sources_used = set()

    for f in files:
        fname = f.filename.lower()
        if not fname.endswith(".csv"):
            continue

        csv_bytes = f.read()

        # フォーマット判定
        try:
            first_line = csv_bytes.decode('shift-jis', errors='replace').split('\n')[0]
        except Exception:
            first_line = ""

        if first_line.startswith('"H"') or first_line.startswith('H,'):
            source = "infomart"
            results = parse_infomart_csv(csv_bytes, f.filename)
        elif '送信者ID' in first_line and 'メッセージ識別ID' in first_line:
            source = "smacla"
            results = parse_smacla_csv(csv_bytes, f.filename)
        else:
            source = "paltac"
            results = parse_paltac_csv(csv_bytes, f.filename)

        if results:
            sources_used.add(source)
            filenames.append(f.filename)
            # ページ番号を全体で連番化
            for r in results:
                r["page"] = len(all_results) + 1
                r["two_order_no"] = generate_two_order_no()
                r["source_file"] = f.filename
                all_results.append(r)

    if not all_results:
        return jsonify({"error": "発注データが見つかりません"}), 400

    session_id = str(uuid.uuid4())[:8]
    # 複数フォーマット混在時はpaltac扱い（OCRスキップさえできればOK）
    if sources_used == {"infomart"}:
        main_source = "infomart"
    elif sources_used == {"smacla"}:
        main_source = "smacla"
    else:
        main_source = "paltac"
    combined_name = f"combined_{len(filenames)}files.csv"

    sessions[session_id] = {
        "pdf_bytes": None,
        "pdf_name": combined_name,
        "pages": [],
        "results": all_results,
        "source": main_source,
    }
    return jsonify({
        "session_id": session_id,
        "filename": combined_name,
        "page_count": len(all_results),
        "source": main_source,
        "auto_results": True,
        "file_count": len(filenames),
    })


@app.route("/api/upload", methods=["POST"])
def api_upload():
    """Upload PDF or PALTAC CSV"""
    if "file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    f = request.files["file"]
    fname = f.filename.lower()

    # CSV処理（PALTAC or インフォマート自動判別）
    if fname.endswith(".csv"):
        csv_bytes = f.read()
        session_id = str(uuid.uuid4())[:8]

        # インフォマート判定: 1行目が "H" で始まる
        try:
            first_line = csv_bytes.decode('shift-jis', errors='replace').split('\n')[0]
        except Exception:
            first_line = ""

        if first_line.startswith('"H"') or first_line.startswith('H,'):
            source = "infomart"
            results = parse_infomart_csv(csv_bytes, f.filename)
            error_msg = "インフォマートCSVにデータが見つかりません"
        elif '送信者ID' in first_line and 'メッセージ識別ID' in first_line:
            source = "smacla"
            results = parse_smacla_csv(csv_bytes, f.filename)
            error_msg = "スマクラCSVにデータが見つかりません"
        else:
            source = "paltac"
            results = parse_paltac_csv(csv_bytes, f.filename)
            error_msg = "発注データが見つかりません（届先区分「発注」のみ取込対象）"

        if not results:
            return jsonify({"error": error_msg}), 400

        # TWO受注NOを採番
        for r in results:
            r["two_order_no"] = generate_two_order_no()

        sessions[session_id] = {
            "pdf_bytes": None,
            "pdf_name": f.filename,
            "pages": [],
            "results": results,
            "source": source,
        }
        return jsonify({
            "session_id": session_id,
            "filename": f.filename,
            "page_count": len(results),
            "source": source,
            "auto_results": True,
        })

    if not fname.endswith(".pdf"):
        return jsonify({"error": "PDFまたはCSVファイルを選択してください"}), 400

    pdf_bytes = f.read()
    session_id = str(uuid.uuid4())[:8]

    # Convert to page images
    pages = pdf_to_images(pdf_bytes, dpi=150)

    sessions[session_id] = {
        "pdf_bytes": pdf_bytes,
        "pdf_name": f.filename,
        "pages": pages,  # [{page, base64}, ...]
        "results": None,
    }

    return jsonify({
        "session_id": session_id,
        "filename": f.filename,
        "page_count": len(pages),
    })


@app.route("/api/page_image/<session_id>/<int:page>")
def api_page_image(session_id, page):
    """Serve a page image as PNG"""
    sess = sessions.get(session_id)
    if not sess:
        return "Session not found", 404
    for p in sess["pages"]:
        if p["page"] == page:
            img_bytes = base64.b64decode(p["base64"])
            from io import BytesIO
            return send_file(BytesIO(img_bytes), mimetype="image/jpeg")
    return "Page not found", 404


@app.route("/api/ocr", methods=["POST"])
def api_ocr():
    """Run OCR + matching on uploaded PDF, or return pre-parsed PALTAC results"""
    data = request.get_json()
    session_id = data.get("session_id")
    sess = sessions.get(session_id)
    if not sess:
        return jsonify({"error": "セッションが見つかりません"}), 404

    # CSV（PALTAC/インフォマート）/ 手入力: 既にパース済みの結果を返す
    if sess.get("source") in ("paltac", "infomart", "manual", "smacla") and sess.get("results"):
        return jsonify({"pages": sess["results"]})

    # Phase 4: ドラフト編集モード（/api/load_draft 経由）も事前展開済み
    if sess.get("draft_id") and sess.get("results"):
        return jsonify({"pages": sess["results"]})

    pm = get_product_master()
    ddc = get_ddc_master()
    pages = sess["pages"]
    results = []

    for page_info in pages:
        ocr_result = ocr_fax_page(page_info["base64"])
        if "error" in ocr_result:
            results.append({
                "page": page_info["page"],
                "error": ocr_result["error"],
                "ocr_raw": ocr_result,
            })
            continue

        # Year correction
        dd = ocr_result.get("delivery_date", "")
        if dd and len(dd) >= 4 and dd[:4].isdigit():
            yr = int(dd[:4])
            if yr < 2020 or yr > 2030:
                ocr_result["delivery_date"] = "2026" + dd[4:]

        # Match products
        matched_items = []
        for ocr_item in ocr_result.get("items", []):
            match = match_product(ocr_item, pm)
            if match:
                matched_items.append(match)
            else:
                try:
                    qty = int(float(str(ocr_item.get("quantity_cs") or 0).strip() or "0"))
                except (ValueError, TypeError):
                    qty = 0
                matched_items.append({
                    "matched": False,
                    "ocr_name": ocr_item.get("product_name", ""),
                    "jan": str(ocr_item.get("jan_code", "")),
                    "quantity": qty,
                })

        # Match DDC
        ddc_match = match_ddc(
            ocr_result.get("delivery_dest", ""),
            ddc,
            sender=ocr_result.get("sender", ""),
        )

        # Split by output
        sylvia_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "シルビア"]
        haruna_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "ハルナ"]

        # TWO受注NO自動採番
        two_order_no = generate_two_order_no()

        results.append({
            "page": page_info["page"],
            "ocr_raw": ocr_result,
            "two_order_no": two_order_no,
            "matched_items": matched_items,
            "ddc_match": ddc_match,
            "sylvia_items": sylvia_items,
            "haruna_items": haruna_items,
        })

    sess["results"] = results
    return jsonify({"session_id": session_id, "pages": results})


@app.route("/api/update_ddc", methods=["POST"])
def api_update_ddc():
    """User selected a different DDC - look up full record"""
    data = request.get_json()
    ddc_name = data.get("ddc_name", "")
    df = get_ddc_master()
    for _, row in df.iterrows():
        if str(row.get("納品先名", "")).strip() == ddc_name:
            return jsonify(_ddc_row_to_dict(row))
    return jsonify({"matched": False, "name": ddc_name})


@app.route("/api/confirm", methods=["POST"])
def api_confirm():
    """Confirm and generate output PDFs"""
    data = request.get_json()
    session_id = data.get("session_id")
    sess = sessions.get(session_id)
    if not sess or not sess["results"]:
        return jsonify({"error": "先にOCR処理を実行してください"}), 400

    results = sess["results"]
    pdf_name = sess["pdf_name"]
    staff_name = data.get("staff_name", "伊藤")
    remarks = data.get("remarks", "")

    # Apply user edits to results
    user_pages = data.get("pages", [])
    for up in user_pages:
        page_num = up.get("page")
        for r in results:
            if r["page"] == page_num and "error" not in r:
                # Update DDC if changed
                if up.get("ddc_name"):
                    df = get_ddc_master()
                    for _, row in df.iterrows():
                        if str(row.get("納品先名", "")).strip() == up["ddc_name"]:
                            r["ddc_match"] = _ddc_row_to_dict(row)
                            r["ddc_match"]["match_score"] = 1.0
                            r["ddc_match"]["low_confidence"] = False
                            break
                # Update items (quantities, expiry, product changes, additions)
                if up.get("items"):
                    # 手入力で商品数が変わった場合、matched_itemsを再構築
                    new_items = []
                    for ui in up["items"]:
                        idx = ui.get("index")
                        if idx is not None and idx < len(r["matched_items"]):
                            mi = r["matched_items"][idx]
                        else:
                            mi = {}
                        # フロントから送られた値で更新
                        if ui.get("matched"):
                            mi["matched"] = True
                            mi["master_name"] = ui.get("master_name", mi.get("master_name", ""))
                            mi["jan"] = ui.get("jan", mi.get("jan", ""))
                            mi["code"] = ui.get("code", mi.get("code", ""))
                            mi["output_dest"] = ui.get("output_dest", mi.get("output_dest", ""))
                            mi["case_quantity"] = ui.get("case_quantity", mi.get("case_quantity", 0))
                            mi["cs_price"] = ui.get("cs_price", mi.get("cs_price", 0))
                            mi["spec"] = ui.get("spec", mi.get("spec", ""))
                            mi["pack"] = ui.get("pack", mi.get("pack", ""))
                            mi["unit_price"] = ui.get("unit_price", mi.get("unit_price", 0))
                        if "quantity" in ui:
                            mi["quantity"] = ui["quantity"]
                            if mi.get("cs_price"):
                                mi["amount"] = mi["quantity"] * mi["cs_price"]
                        if "expiry_date" in ui:
                            mi["expiry_date"] = ui["expiry_date"]
                        if "double_pack" in ui:
                            mi["double_pack"] = ui["double_pack"]
                        new_items.append(mi)
                    r["matched_items"] = new_items
                # Update delivery date
                if up.get("delivery_date"):
                    r["ocr_raw"]["delivery_date"] = up["delivery_date"]
                # Update order_no
                if up.get("order_no"):
                    r["ocr_raw"]["order_no"] = up["order_no"]
                # Update TWO受注NO
                if up.get("two_order_no"):
                    r["two_order_no"] = up["two_order_no"]
                # オーダーNOが空の場合、TWO受注NOを補完
                if not r["ocr_raw"].get("order_no"):
                    r["ocr_raw"]["order_no"] = r.get("two_order_no", "")
                # Update remarks (per page)
                if "remarks" in up:
                    r["remarks"] = up["remarks"]
                # Update warehouse_direct flag
                if "warehouse_direct" in up:
                    r["warehouse_direct"] = up["warehouse_direct"]
                    if up["warehouse_direct"]:
                        r["ddc_match"] = {
                            "matched": True,
                            "name": "株式会社ベルーナ・ジーエフ・ロジスティクス",
                            "postal": "3620066",
                            "address": "埼玉県上尾市領家丸山30-1",
                            "tel": "048-725-0179",
                            "fax": "",
                            "time": "",
                            "berse": "",
                            "palette": "",
                            "jpr": "",
                            "method": "",
                            "match_score": 1.0,
                            "low_confidence": False,
                            "candidates": [],
                        }
                # Re-split sylvia/haruna
                r["sylvia_items"] = [i for i in r["matched_items"] if i.get("matched") and i.get("output_dest") == "シルビア"]
                r["haruna_items"] = [i for i in r["matched_items"] if i.get("matched") and i.get("output_dest") == "ハルナ"]

    # Phase 4: 送料計算（ページ毎個別計算）
    # フロントから shipping_fee_two_burden / shipping_fee_manual_overrides が来る前提
    # manual_overrides 形式: {"<pageIdx>": {"brand": int}}
    shipping_fee_two_burden = bool(data.get("shipping_fee_two_burden", False))
    manual_overrides = data.get("shipping_fee_manual_overrides") or {}

    shipping_fee_total = 0
    shipping_fee_breakdown = None
    try:
        from shipping_fee_calculator import calculate_shipping_fee, append_lot_break_history
        per_page_results = []
        total_all_pages = 0
        for idx, r in enumerate(results):
            if "error" in r:
                per_page_results.append(None)
                continue
            items = r.get("matched_items", [])
            addr = r.get("ddc_match", {}).get("address", "") or ""
            ship_type = "自社倉庫" if r.get("warehouse_direct") else None
            shipping_result = calculate_shipping_fee(items, addr, ship_type)
            # 手動入力（Gummy/Water）の上書きを反映（ページ別）
            page_overrides = manual_overrides.get(str(idx)) or manual_overrides.get(idx) or {}
            for g in shipping_result.get("groups", []):
                if g.get("needs_manual"):
                    brand = g["brand"]
                    if brand in page_overrides:
                        try:
                            override_amt = int(page_overrides[brand])
                            g["fee"] = override_amt
                            g["needs_manual"] = False
                            g["manual_input_value"] = override_amt
                            shipping_result["total_fee"] += override_amt
                        except (ValueError, TypeError):
                            pass
            shipping_result["needs_manual_input_overall"] = any(
                g.get("needs_manual") for g in shipping_result.get("groups", [])
            )
            page_fee = 0 if shipping_fee_two_burden else int(shipping_result.get("total_fee") or 0)
            r["shipping_fee"] = page_fee
            total_all_pages += page_fee
            per_page_results.append(shipping_result)

            # lot_break_history.csv へ追記（2Snack の自動計算分のみ、ページ毎に1回）
            for g in shipping_result.get("groups", []):
                if g.get("method") == "snack_lot_break" and g.get("detail", {}).get("applicable"):
                    try:
                        append_lot_break_history(g["detail"], irisuu=12)
                    except Exception as he:
                        print(f"[shipping_fee] lot_break_history 追記失敗: {he}")

        shipping_fee_total = total_all_pages
        shipping_fee_breakdown = {"pages": per_page_results, "total_fee_all_pages": total_all_pages}
    except Exception as se:
        import traceback
        traceback.print_exc()
        print(f"[shipping_fee] 計算エラー: {se}")
        # 送料計算に失敗しても確定処理は続行（送料 0 で出力）

    # Generate outputs
    try:
        ensure_output_dir()
        generated = generate_pdfs(results, pdf_name, staff_name)
        csv_path = results_to_csv(results, pdf_name)
        xlsx_path = results_to_excel(results, pdf_name)
        # NE CSV: 全件出力（売上管理用、直送・自社倉庫経由問わず）
        ne_csv_path = results_to_ne_csv(results, pdf_name) if results else None
        # COOLA CSV: 自社倉庫出荷システム用。直送商品（10cs以上）は results_to_coola_csv 内でフィルタ済
        coola_csv_path = results_to_coola_csv(results, pdf_name) if results else None
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"生成エラー: {repr(e)}"}), 500

    # Phase 4: ドラフト経由なら Supabase の status を confirmed に更新 + スナップショット保存
    draft_id = sess.get("draft_id")
    confirm_result = None
    if draft_id:
        try:
            from supabase_client import confirm_fax_order
            # スナップショット作成（編集後の最新状態を保存）
            snapshot = _build_confirmation_snapshot(results, staff_name, remarks,
                                                    shipping_fee_total, shipping_fee_two_burden)
            # 編集されたヘッダー項目（最終状態）
            r0 = next((r for r in results if "error" not in r), None)
            edited_header = None
            if r0:
                ocr0 = r0.get("ocr_raw", {})
                ddc0 = r0.get("ddc_match", {})
                edited_header = {
                    "slip_number": ocr0.get("order_no") or None,
                    "delivery_date": ocr0.get("delivery_date") or None,
                    "partner_name": (ocr0.get("sender") or "")[:200] or None,
                    "delivery_location_code": ddc0.get("code") if ddc0.get("matched") else None,
                    "delivery_location_name": ddc0.get("name") if ddc0.get("matched") else None,
                    "delivery_location_address": ddc0.get("address") if ddc0.get("matched") else None,
                    "notes": ocr0.get("notes") or None,
                    "warehouse": r0.get("ddc_match", {}).get("name") if r0.get("warehouse_direct") else None,
                }
            confirm_result = confirm_fax_order(
                draft_id,
                confirmed_by=staff_name,
                snapshot=snapshot,
                shipping_fee=shipping_fee_total,
                shipping_fee_two_burden=shipping_fee_two_burden,
                shipping_fee_breakdown=shipping_fee_breakdown,
                edited_header=edited_header,
            )
        except Exception as ce:
            import traceback
            traceback.print_exc()
            print(f"[confirm] Supabase更新失敗: {ce}")
            # Supabase 更新失敗してもファイル生成は完了しているので 200 で返す（エラーフラグ付き）

    output_files = []
    if ne_csv_path:
        output_files.append({"name": os.path.basename(ne_csv_path), "type": "ne_csv", "label": "NE受注CSV"})
    if coola_csv_path:
        output_files.append({"name": os.path.basename(coola_csv_path), "type": "coola_csv", "label": "クーラCSV"})
    if csv_path:
        output_files.append({"name": os.path.basename(csv_path), "type": "csv"})
    if xlsx_path:
        output_files.append({"name": os.path.basename(xlsx_path), "type": "xlsx"})
    for label, path in generated:
        output_files.append({"name": os.path.basename(path), "type": "pdf", "label": label})

    return jsonify({
        "success": True,
        "files": output_files,
        "shipping_fee_total": shipping_fee_total,
        "shipping_fee_two_burden": shipping_fee_two_burden,
        "shipping_fee_breakdown": shipping_fee_breakdown,
        "draft_updated": bool(confirm_result),
        "confirmation_count": (confirm_result or {}).get("confirmation_count"),
    })


def _build_confirmation_snapshot(results, staff_name, remarks, shipping_fee, two_burden):
    """確定時のスナップショットを生成（confirmation_history に append される）"""
    pages = []
    for r in results:
        if "error" in r:
            continue
        ocr = r.get("ocr_raw", {})
        ddc = r.get("ddc_match", {})
        pages.append({
            "page": r.get("page"),
            "order_no": ocr.get("order_no"),
            "two_order_no": r.get("two_order_no"),
            "delivery_date": ocr.get("delivery_date"),
            "partner_name": ocr.get("sender"),
            "delivery_location_code": ddc.get("code"),
            "delivery_location_name": ddc.get("name"),
            "delivery_location_address": ddc.get("address"),
            "warehouse_direct": bool(r.get("warehouse_direct")),
            "notes": ocr.get("notes"),
            "remarks_per_page": r.get("remarks"),
            "shipping_fee": int(r.get("shipping_fee") or 0),
            "items": [{
                "jan": i.get("jan"),
                "code": i.get("code"),
                "master_name": i.get("master_name"),
                "ocr_name": i.get("ocr_name"),
                "spec": i.get("spec"),
                "pack": i.get("pack"),
                "quantity": i.get("quantity"),
                "unit_price": i.get("unit_price"),
                "cs_price": i.get("cs_price"),
                "amount": i.get("amount"),
                "output_dest": i.get("output_dest"),
                "matched": bool(i.get("matched")),
                "double_pack": bool(i.get("double_pack")),
                "expiry_date": i.get("expiry_date"),
            } for i in r.get("matched_items", [])],
        })
    return {
        "staff_name": staff_name,
        "remarks_global": remarks,
        "shipping_fee": int(shipping_fee or 0),
        "shipping_fee_two_burden": bool(two_burden),
        "pages": pages,
    }


## ─── Drive 新着PDF取込（非同期実行） ───

_import_state = {
    "running": False,
    "phase": "idle",          # idle / scanning / processing / done / error
    "total": 0,
    "current": 0,
    "current_filename": "",
    "inserted": 0,
    "skipped": 0,
    "errors": 0,
    "moved_to_error": 0,
    "started_at": None,       # ISO文字列
    "finished_at": None,
    "result": None,           # run_import の戻り値（完了時のみ）
    "error_message": None,    # 例外時のメッセージ
}
_import_state_lock = threading.Lock()


def _run_import_thread(limit: int | None):
    """drive_processor.run_import をバックグラウンドで走らせるスレッド本体。"""
    from drive_processor import run_import as _run_import

    def progress_cb(state):
        with _import_state_lock:
            _import_state.update({
                "phase": state.get("phase", _import_state["phase"]),
                "total": state.get("total", _import_state["total"]),
                "current": state.get("current", _import_state["current"]),
                "current_filename": state.get("current_filename", ""),
                "inserted": state.get("inserted", _import_state["inserted"]),
                "skipped": state.get("skipped", _import_state["skipped"]),
                "errors": state.get("errors", _import_state["errors"]),
                "moved_to_error": state.get("moved_to_error", _import_state["moved_to_error"]),
            })

    try:
        result = _run_import(
            limit=limit,
            mode="new",
            no_supabase=False,
            source_channel="efax",
            progress_cb=progress_cb,
        )
        with _import_state_lock:
            _import_state["result"] = result
            _import_state["phase"] = "done"
    except Exception as e:
        import traceback
        traceback.print_exc()
        with _import_state_lock:
            _import_state["phase"] = "error"
            _import_state["error_message"] = str(e)
    finally:
        with _import_state_lock:
            _import_state["running"] = False
            _import_state["finished_at"] = datetime.now().isoformat(timespec="seconds")


@app.route("/api/import_drive", methods=["POST"])
def api_import_drive():
    """Drive 新着フォルダの PDF/CSV を非同期で取込開始。即座に200を返す。"""
    with _import_state_lock:
        if _import_state["running"]:
            return jsonify({
                "started": False,
                "message": "既に取込処理が実行中です",
                "state": {k: v for k, v in _import_state.items() if k != "result"},
            }), 409

        # 状態リセット
        _import_state.update({
            "running": True,
            "phase": "scanning",
            "total": 0,
            "current": 0,
            "current_filename": "",
            "inserted": 0,
            "skipped": 0,
            "errors": 0,
            "moved_to_error": 0,
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "finished_at": None,
            "result": None,
            "error_message": None,
        })

    # オプション: ?limit=5 で件数制限（テスト用）
    limit_arg = request.args.get("limit", type=int)

    t = threading.Thread(target=_run_import_thread, args=(limit_arg,), daemon=True)
    t.start()
    return jsonify({"started": True, "message": "取込を開始しました"})


@app.route("/api/import_status")
def api_import_status():
    """現在の取込状態を返す。フロントが2-3秒おきにポーリング。"""
    with _import_state_lock:
        return jsonify(dict(_import_state))


@app.route("/api/reload_master", methods=["POST"])
def api_reload_master():
    """マスタデータのキャッシュをクリアして再読込"""
    global _ddc_master_df, _product_master_df, _ddc_list_cache
    from supabase_client import clear_all_caches
    from process_fax import _oroshisaki_cache
    import process_fax
    clear_all_caches()
    _ddc_master_df = None
    _product_master_df = None
    _ddc_list_cache = None
    process_fax._oroshisaki_cache = None
    # 再読込
    ddc_count = len(get_ddc_master())
    prod_count = len(get_product_master())
    return jsonify({"success": True, "message": f"マスタ再読込完了（納品先: {ddc_count}件, 商品: {prod_count}件）"})


@app.route("/api/download/<filename>")
def api_download(filename):
    """Download an output file"""
    path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "File not found", 404


# ─── Pending Orders HTML Template (Phase 4a) ───

PENDING_ORDERS_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>保留中受注一覧 - FAX受注処理システム</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; margin: 0; background: #f5f5f5; color: #222; }
  .header { background: #2E7D32; color: white; padding: 12px 20px; display: flex; align-items: center; gap: 8px; }
  .header h1 { margin: 0; font-size: 18px; }
  .header a { color: white; text-decoration: none; padding: 6px 14px; border-radius: 4px; background: rgba(255,255,255,0.2); font-size: 14px; }
  .header a:hover { background: rgba(255,255,255,0.3); }
  .ver { font-size: 12px; opacity: 0.7; margin-left: 12px; }
  .container { padding: 20px; }
  .toolbar { display: flex; gap: 12px; margin-bottom: 12px; align-items: center; flex-wrap: wrap; }
  .toolbar button { padding: 6px 14px; border: 1px solid #ccc; background: white; cursor: pointer; border-radius: 4px; font-size: 14px; }
  .toolbar button:hover { background: #f0f0f0; }
  .toolbar button:disabled { opacity: 0.6; cursor: not-allowed; }
  .toolbar button.active { background: #1976D2; color: white; border-color: #1976D2; }
  .toolbar button.import-btn { background: #2E7D32; color: white; border-color: #2E7D32; font-weight: 600; }
  .toolbar button.import-btn:hover:not(:disabled) { background: #1B5E20; }
  .summary { font-size: 14px; color: #555; }
  .import-status { background: #FFF8E1; border: 1px solid #FFD54F; border-radius: 4px; padding: 10px 14px; margin-bottom: 12px; font-size: 13px; display: none; }
  .import-status.error { background: #FFEBEE; border-color: #EF9A9A; }
  .import-status.done { background: #E8F5E9; border-color: #A5D6A7; }
  .import-status .progress-bar { height: 6px; background: #E0E0E0; border-radius: 3px; margin-top: 6px; overflow: hidden; }
  .import-status .progress-bar > div { height: 100%; background: #2E7D32; transition: width 0.3s; }
  .import-status .stats { color: #555; font-size: 12px; margin-top: 4px; }
  .toast { position: fixed; bottom: 24px; right: 24px; background: #323232; color: white; padding: 12px 20px; border-radius: 4px; font-size: 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.3); z-index: 200; animation: toast-in 0.3s; }
  .toast.success { background: #2E7D32; }
  .toast.error { background: #C62828; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: none; } }
  table { width: 100%; background: white; border-collapse: collapse; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid #eee; font-size: 13px; }
  th { background: #fafafa; font-weight: 600; color: #333; position: sticky; top: 0; }
  tr:hover { background: #f5f9ff; cursor: pointer; }
  tr.selected { background: #e3f2fd; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
  .badge-efax { background: #E3F2FD; color: #1565C0; }
  .badge-paltac { background: #FFF3E0; color: #E65100; }
  .badge-infomart { background: #F3E5F5; color: #6A1B9A; }
  .badge-draft { background: #E8F5E9; color: #2E7D32; }
  .badge-error { background: #FFEBEE; color: #C62828; }
  .flag { display: inline-block; padding: 1px 6px; background: #FFE0B2; color: #BF360C; border-radius: 3px; font-size: 11px; margin-right: 4px; }
  .empty { text-align: center; padding: 40px; color: #888; }
  .loading { text-align: center; padding: 40px; color: #666; }

  /* 詳細パネル */
  .detail-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.4); display: none; z-index: 100; }
  .detail-panel { position: fixed; top: 0; right: 0; bottom: 0; width: 600px; max-width: 95vw; background: white; box-shadow: -2px 0 10px rgba(0,0,0,0.2); padding: 20px; overflow-y: auto; z-index: 101; display: none; }
  .detail-panel h2 { margin-top: 0; }
  .detail-panel .close { position: absolute; top: 12px; right: 16px; background: none; border: none; font-size: 24px; cursor: pointer; color: #888; }
  .detail-section { margin-bottom: 20px; }
  .detail-section h3 { margin: 0 0 8px; color: #1976D2; font-size: 14px; border-bottom: 1px solid #eee; padding-bottom: 4px; }
  .kv { display: grid; grid-template-columns: 140px 1fr; gap: 4px 12px; font-size: 13px; }
  .kv .k { color: #666; }
  .kv .v { color: #222; }
  .item-row { padding: 8px 0; border-bottom: 1px solid #f0f0f0; font-size: 13px; }
  .item-row .name { font-weight: 600; }
  .item-row .meta { color: #666; font-size: 12px; }
</style>
</head>
<body>

<div class="header">
  <h1>📋 保留中受注一覧</h1>
  <span class="ver" title="web_app.py の最終更新日時">ver. {{ app_last_updated }}</span>
  <div style="flex: 1"></div>
  <a href="/confirmed">✅ 確定済み一覧</a>
  <a href="/">← FAX受注処理に戻る</a>
</div>

<div class="container">
  <div class="toolbar">
    <button onclick="loadOrders('')" id="btn-all" class="active">すべて</button>
    <button onclick="loadOrders('efax')" id="btn-efax">eFax</button>
    <button onclick="loadOrders('paltac')" id="btn-paltac">PALTAC</button>
    <button onclick="startImport()" id="btn-import" class="import-btn" style="margin-left: auto">📥 新着PDF取込</button>
    <button onclick="loadOrders(null, true)" id="btn-reload">🔄 更新</button>
    <span class="summary" id="summary"></span>
  </div>

  <!-- 検索バー -->
  <div class="search-bar" style="display:flex; gap:12px; margin-bottom:12px; padding:10px 14px; background:white; border-radius:6px; box-shadow:0 1px 3px rgba(0,0,0,0.08); align-items:center;">
    <span style="font-size:13px; color:#555; font-weight:600">🔍 検索:</span>
    <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#444;">
      取引先
      <input type="text" id="filterPartner" oninput="applyFilter()" placeholder="例: ネクスコ"
             style="padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; width:160px;">
    </label>
    <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#444;">
      納品日
      <input type="date" id="filterDeliveryDate" onchange="applyFilter()" oninput="applyFilter()"
             style="padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px;">
    </label>
    <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#444;">
      伝票No
      <input type="text" id="filterSlip" oninput="applyFilter()" placeholder="例: 00868"
             style="padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; width:140px;">
    </label>
    <button onclick="clearFilters()"
            style="padding:5px 12px; border:1px solid #ccc; background:#f8f8f8; border-radius:4px; cursor:pointer; font-size:12px;">
      クリア
    </button>
  </div>

  <div class="import-status" id="importStatus"></div>

  <div id="content">
    <div class="loading">読み込み中...</div>
  </div>
</div>

<!-- 詳細パネル -->
<div class="detail-overlay" id="detailOverlay" onclick="closeDetail()"></div>
<div class="detail-panel" id="detailPanel">
  <button class="close" onclick="closeDetail()">×</button>
  <div id="detailBody"></div>
</div>

<script>
let currentSource = '';
let currentOrders = [];   // 取得した全件
let filteredOrders = [];  // 検索フィルタ適用後（描画用）

async function loadOrders(source, isReload) {
  if (source !== undefined) currentSource = source;

  // ボタンのactive状態更新
  document.querySelectorAll('.toolbar button').forEach(b => b.classList.remove('active'));
  const btnId = currentSource === '' ? 'btn-all' : `btn-${currentSource}`;
  const btn = document.getElementById(btnId);
  if (btn) btn.classList.add('active');

  document.getElementById('content').innerHTML = '<div class="loading">読み込み中...</div>';

  try {
    const url = '/api/pending_orders' + (currentSource ? `?source=${currentSource}` : '');
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    currentOrders = data.orders || [];
    applyFilter();  // 既存検索条件を適用して描画
  } catch (err) {
    document.getElementById('content').innerHTML = `<div class="empty">取得失敗: ${err.message}</div>`;
  }
}

// 検索バー: 取引先・納品日・伝票No でクライアント側フィルタ
function applyFilter() {
  const partnerQ = (document.getElementById('filterPartner').value || '').toLowerCase().trim();
  const dateQ = (document.getElementById('filterDeliveryDate').value || '').trim();
  const slipQ = (document.getElementById('filterSlip').value || '').toLowerCase().trim();

  filteredOrders = currentOrders.filter(o => {
    if (partnerQ) {
      const partner = (o.partner_name || '').toLowerCase();
      if (!partner.includes(partnerQ)) return false;
    }
    if (dateQ) {
      if ((o.delivery_date || '') !== dateQ) return false;
    }
    if (slipQ) {
      const slip = (o.slip_number || '').toLowerCase();
      if (!slip.includes(slipQ)) return false;
    }
    return true;
  });

  renderTable();
  const totalText = currentOrders.length;
  const visibleText = filteredOrders.length;
  const filterActive = partnerQ || dateQ || slipQ;
  document.getElementById('summary').textContent =
    filterActive ? `${visibleText}件 / 全${totalText}件` : `${totalText}件`;
}

function clearFilters() {
  document.getElementById('filterPartner').value = '';
  document.getElementById('filterDeliveryDate').value = '';
  document.getElementById('filterSlip').value = '';
  applyFilter();
}

function renderTable() {
  if (filteredOrders.length === 0) {
    if (currentOrders.length === 0) {
      document.getElementById('content').innerHTML = '<div class="empty">保留中の受注はありません 🎉</div>';
    } else {
      document.getElementById('content').innerHTML = '<div class="empty">該当する受注がありません（検索条件を変えてみてください）</div>';
    }
    return;
  }

  const rows = filteredOrders.map(o => {
    const channelBadge = `<span class="badge badge-${o.source_channel || 'unknown'}">${o.source_channel || '-'}</span>`;
    const statusBadge = `<span class="badge badge-${o.status}">${o.status}</span>`;
    const flags = (o.status_flags || []).map(f => `<span class="flag">${f}</span>`).join('');
    const ddc = o.ddc_matched ? '✅' : '❌';
    const total = o.grand_total != null ? Number(o.grand_total).toLocaleString() + '円' : '-';
    const partner = (o.partner_name || '').substring(0, 24);
    const dest = (o.delivery_location_name || '').substring(0, 30);
    return `<tr onclick="showDetail('${o.id}', this)" data-id="${o.id}">
      <td>${channelBadge}</td>
      <td>${o.slip_number || '-'}</td>
      <td>${o.delivery_date || '-'}</td>
      <td>${partner}</td>
      <td>${dest} ${ddc}</td>
      <td>${o.warehouse || '-'}</td>
      <td>${total}</td>
      <td>${statusBadge} ${flags}</td>
      <td style="font-size:11px;color:#888">${(o.created_at || '').substring(0, 16).replace('T', ' ')}</td>
    </tr>`;
  }).join('');

  document.getElementById('content').innerHTML = `
    <table>
      <thead>
        <tr>
          <th>経路</th><th>伝票No</th><th>納品日</th><th>取引先</th><th>納品先(DDC)</th>
          <th>倉庫</th><th>合計</th><th>状態</th><th>登録日時</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;
}

async function showDetail(orderId, rowEl) {
  document.querySelectorAll('tr.selected').forEach(r => r.classList.remove('selected'));
  if (rowEl) rowEl.classList.add('selected');

  document.getElementById('detailOverlay').style.display = 'block';
  document.getElementById('detailPanel').style.display = 'block';
  document.getElementById('detailBody').innerHTML = '<div class="loading">読み込み中...</div>';

  try {
    const res = await fetch(`/api/pending_orders/${orderId}`);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    renderDetail(data);
  } catch (err) {
    document.getElementById('detailBody').innerHTML = `<div class="empty">取得失敗: ${err.message}</div>`;
  }
}

function renderDetail(data) {
  const h = data.header;
  const items = data.items || [];

  const kv = (k, v) => `<div class="k">${k}</div><div class="v">${v != null && v !== '' ? v : '-'}</div>`;

  // Phase 4: 編集して確定 ボタン（draft / error のみ表示）
  const canEdit = h.status === 'draft' || h.status === 'error';
  const editBtnHtml = canEdit
    ? `<div style="margin: 12px 0 18px; display: flex; gap: 10px; align-items: center;">
         <button onclick="openDraftEditor('${h.id}')"
                 style="padding: 10px 20px; background: #2E7D32; color: white; border: none;
                        border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer;
                        display: inline-flex; align-items: center; gap: 6px;">
           ✏️ 編集して確定
         </button>
         <span style="color: #666; font-size: 12px">
           → 確認画面で内容修正・確定 → 発注書PDF・NE/COOLA CSV を生成
         </span>
       </div>`
    : '';

  let html = `
    <h2>${h.slip_number || '(伝票No未設定)'} <small style="color:#888">${h.source_channel}</small></h2>
    ${editBtnHtml}

    <div class="detail-section">
      <h3>受注情報</h3>
      <div class="kv">
        ${kv('ステータス', h.status)}
        ${kv('発注日', h.order_date)}
        ${kv('納品日', h.delivery_date)}
        ${kv('取引先', h.partner_name)}
        ${kv('納品先(OCR)', h.delivery_location_name)}
        ${kv('住所', h.delivery_location_address)}
        ${kv('DDCマッチ', h.ddc_matched ? '✅ マッチ' : '❌ 未マッチ')}
        ${kv('倉庫判定', h.warehouse)}
        ${kv('出荷区分', h.shipping_type)}
        ${kv('合計金額', h.grand_total != null ? Number(h.grand_total).toLocaleString() + '円' : '')}
        ${kv('要確認', (h.status_flags || []).join(', '))}
      </div>
    </div>

    <div class="detail-section">
      <h3>明細 (${items.length}件)</h3>`;

  if (items.length === 0) {
    html += '<div style="color:#888">明細なし</div>';
  } else {
    html += items.map(i => `
      <div class="item-row">
        <div class="name">${i.product_name || '-'} ${i.product_master_matched ? '✅' : '❌'}</div>
        <div class="meta">
          JAN: ${i.jan_code || '-'} / ${i.quantity || 0} ${i.unit || ''} ×
          ${i.unit_price != null ? Number(i.unit_price).toLocaleString() + '円' : '?'}
          = ${i.amount != null ? Number(i.amount).toLocaleString() + '円' : '?'}
          ${i.spec ? ' / ' + i.spec : ''}
        </div>
      </div>`).join('');
  }

  html += `
    </div>

    <div class="detail-section">
      <h3>取込元情報</h3>
      <div class="kv">
        ${kv('元ファイル', h.source_file_name)}
        ${kv('元メール送信元', h.source_email_from)}
        ${kv('元メール受信日時', h.source_email_received_at ? h.source_email_received_at.replace('T', ' ').substring(0, 16) : '')}
        ${kv('OCRページ', h.ocr_page_no)}
      </div>
    </div>

    <div style="font-size:11px;color:#999;margin-top:20px">
      ID: ${h.id}<br>
      作成: ${(h.created_at || '').replace('T', ' ').substring(0, 19)}<br>
      更新: ${(h.updated_at || '').replace('T', ' ').substring(0, 19)}
    </div>`;

  document.getElementById('detailBody').innerHTML = html;
}

// Phase 4: ドラフト編集モードでトップページを開く
function openDraftEditor(orderId) {
  // 単純に / にリダイレクト → / 側で URL パラメータを読んで /api/load_draft を呼ぶ
  window.location.href = '/?draft_id=' + encodeURIComponent(orderId);
}

function closeDetail() {
  document.getElementById('detailOverlay').style.display = 'none';
  document.getElementById('detailPanel').style.display = 'none';
  document.querySelectorAll('tr.selected').forEach(r => r.classList.remove('selected'));
}

// ESCキーで閉じる
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeDetail(); });

// ─── 新着PDF取込（非同期＋ポーリング） ───
let importPollTimer = null;

function showToast(msg, type) {
  const t = document.createElement('div');
  t.className = 'toast' + (type ? ' ' + type : '');
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 4000);
}

async function startImport() {
  const btn = document.getElementById('btn-import');
  btn.disabled = true;
  try {
    const res = await fetch('/api/import_drive', { method: 'POST' });
    const data = await res.json();
    if (!res.ok) {
      // 409: 既に実行中なら、状態だけ取りに行く（多重起動防止）
      if (res.status === 409) {
        showToast(data.message || '既に取込処理が実行中', 'error');
        startPolling();
        return;
      }
      throw new Error(data.error || data.message || ('HTTP ' + res.status));
    }
    showToast(data.message || '取込を開始しました');
    startPolling();
  } catch (err) {
    btn.disabled = false;
    showToast('取込開始失敗: ' + err.message, 'error');
  }
}

function startPolling() {
  if (importPollTimer) clearInterval(importPollTimer);
  // 即座に1回 + その後 2秒おき
  pollImportStatus();
  importPollTimer = setInterval(pollImportStatus, 2000);
}

function stopPolling() {
  if (importPollTimer) {
    clearInterval(importPollTimer);
    importPollTimer = null;
  }
}

async function pollImportStatus() {
  try {
    const res = await fetch('/api/import_status');
    const s = await res.json();
    renderImportStatus(s);

    if (!s.running) {
      stopPolling();
      document.getElementById('btn-import').disabled = false;

      if (s.phase === 'done') {
        const r = s.result || {};
        if (r.status === 'no_files') {
          showToast('新着PDFはありません', 'success');
        } else {
          showToast(
            `取込完了: ${r.processed}件処理 / 登録${r.sb_inserted} / 重複${r.sb_skipped} / エラー${r.sb_errors}`,
            'success'
          );
          // 一覧自動更新
          loadOrders(null, true);
        }
      } else if (s.phase === 'error') {
        showToast('取込失敗: ' + (s.error_message || '不明なエラー'), 'error');
      }
    }
  } catch (err) {
    console.error('ポーリング失敗:', err);
  }
}

function renderImportStatus(s) {
  const el = document.getElementById('importStatus');
  // 起動前 (idle / 結果なし) は非表示
  if (s.phase === 'idle' && !s.running) {
    el.style.display = 'none';
    return;
  }

  el.style.display = 'block';
  el.classList.remove('error', 'done');

  let html = '';
  if (s.phase === 'scanning') {
    html = '🔍 新着フォルダをスキャン中...';
  } else if (s.phase === 'processing') {
    const pct = s.total > 0 ? Math.round(s.current / s.total * 100) : 0;
    html = `⚙️ 処理中 (${s.current} / ${s.total}件)`;
    if (s.current_filename) html += `<div style="font-size:12px;color:#666;margin-top:2px">📄 ${s.current_filename}</div>`;
    html += `<div class="progress-bar"><div style="width:${pct}%"></div></div>`;
    html += `<div class="stats">登録 ${s.inserted}件 / 重複スキップ ${s.skipped}件 / エラー ${s.errors}件 / エラー移動 ${s.moved_to_error}件</div>`;
  } else if (s.phase === 'done') {
    el.classList.add('done');
    const r = s.result || {};
    if (r.status === 'no_files') {
      html = `✅ 新着PDFはありません（${s.finished_at || ''}）`;
    } else {
      html = `✅ 取込完了 — ${r.processed}件処理 (${r.elapsed_sec}秒)`;
      html += `<div class="stats">登録 ${r.sb_inserted}件 / 重複スキップ ${r.sb_skipped}件 / エラー ${r.sb_errors}件 / エラー移動 ${r.error_files}件</div>`;
    }
  } else if (s.phase === 'error') {
    el.classList.add('error');
    html = '❌ 取込失敗: ' + (s.error_message || '不明');
  }

  el.innerHTML = html;
}

// 初期ロード
loadOrders('');
// ページを開いた時点で実行中なら、状況を引き継いで表示
(async () => {
  try {
    const res = await fetch('/api/import_status');
    const s = await res.json();
    if (s.running) {
      document.getElementById('btn-import').disabled = true;
      renderImportStatus(s);
      startPolling();
    }
  } catch (e) { /* 無視 */ }
})();
</script>

</body>
</html>"""


# ─── Confirmed Orders Page (Phase 4) ───

CONFIRMED_ORDERS_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>確定済み受注一覧 - FAX受注処理システム</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; margin: 0; background: #f5f5f5; color: #222; }
  .header { background: #1565C0; color: white; padding: 12px 20px; display: flex; align-items: center; gap: 8px; }
  .header h1 { margin: 0; font-size: 18px; }
  .header a { color: white; text-decoration: none; padding: 6px 14px; border-radius: 4px; background: rgba(255,255,255,0.2); font-size: 14px; }
  .header a:hover { background: rgba(255,255,255,0.3); }
  .ver { font-size: 12px; opacity: 0.7; margin-left: 12px; }
  .container { padding: 20px; }
  .toolbar { display: flex; gap: 12px; margin-bottom: 12px; align-items: center; flex-wrap: wrap; }
  .toolbar button { padding: 6px 14px; border: 1px solid #ccc; background: white; cursor: pointer; border-radius: 4px; font-size: 14px; }
  .toolbar button:hover { background: #f0f0f0; }
  .toolbar button.active { background: #1976D2; color: white; border-color: #1976D2; }
  .toolbar select { padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; background: white; }
  .summary { font-size: 14px; color: #555; }
  table { width: 100%; background: white; border-collapse: collapse; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid #eee; font-size: 13px; }
  th { background: #fafafa; font-weight: 600; color: #333; position: sticky; top: 0; }
  tr:hover { background: #f5f9ff; cursor: pointer; }
  tr.selected { background: #e3f2fd; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
  .badge-efax { background: #E3F2FD; color: #1565C0; }
  .badge-paltac { background: #FFF3E0; color: #E65100; }
  .badge-infomart { background: #F3E5F5; color: #6A1B9A; }
  .badge-confirmed { background: #C8E6C9; color: #1B5E20; }
  .badge-count { background: #FFE0B2; color: #BF360C; padding: 1px 7px; border-radius: 10px; font-size: 11px; font-weight: 700; }
  .empty { text-align: center; padding: 40px; color: #888; }
  .loading { text-align: center; padding: 40px; color: #666; }

  /* 詳細パネル */
  .detail-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.4); display: none; z-index: 100; }
  .detail-panel { position: fixed; top: 0; right: 0; bottom: 0; width: 600px; max-width: 95vw; background: white; box-shadow: -2px 0 10px rgba(0,0,0,0.2); padding: 20px; overflow-y: auto; z-index: 101; display: none; }
  .detail-panel h2 { margin-top: 0; }
  .detail-panel .close { position: absolute; top: 12px; right: 16px; background: none; border: none; font-size: 24px; cursor: pointer; color: #888; }
  .detail-section { margin-bottom: 20px; }
  .detail-section h3 { margin: 0 0 8px; color: #1565C0; font-size: 14px; border-bottom: 1px solid #eee; padding-bottom: 4px; }
  .kv { display: grid; grid-template-columns: 140px 1fr; gap: 4px 12px; font-size: 13px; }
  .kv .k { color: #666; }
  .kv .v { color: #222; }
  .item-row { padding: 8px 0; border-bottom: 1px solid #f0f0f0; font-size: 13px; }
  .item-row .name { font-weight: 600; }
  .item-row .meta { color: #666; font-size: 12px; }
  .history-row { padding: 8px 12px; background: #FFF8E1; border-radius: 4px; margin-bottom: 6px; font-size: 12px; }
  .history-row .at { color: #BF360C; font-weight: 600; }
</style>
</head>
<body>

<div class="header">
  <h1>✅ 確定済み受注一覧</h1>
  <span class="ver" title="web_app.py の最終更新日時">ver. {{ app_last_updated }}</span>
  <div style="flex: 1"></div>
  <a href="/pending">📋 保留中一覧</a>
  <a href="/">← FAX受注処理に戻る</a>
</div>

<div class="container">
  <div class="toolbar">
    <button onclick="loadOrders('')" id="btn-all" class="active">すべて</button>
    <button onclick="loadOrders('efax')" id="btn-efax">eFax</button>
    <button onclick="loadOrders('paltac')" id="btn-paltac">PALTAC</button>
    <span style="margin-left: 16px; font-size: 13px; color: #666;">期間:</span>
    <select id="daysSelect" onchange="loadOrders(null, true)">
      <option value="30">直近 30日</option>
      <option value="90" selected>直近 90日</option>
      <option value="180">直近 180日</option>
      <option value="">全期間</option>
    </select>
    <button onclick="loadOrders(null, true)" style="margin-left: auto">🔄 更新</button>
    <span class="summary" id="summary"></span>
  </div>

  <!-- 一括CSV出力バー -->
  <div class="bulk-bar" style="display:flex; gap:10px; margin-bottom:12px; padding:10px 14px; background:#E8F5E9; border-left:4px solid #2E7D32; border-radius:6px; align-items:center; flex-wrap:wrap;">
    <span style="font-size:13px; color:#1B5E20; font-weight:600">📦 まとめCSV出力:</span>
    <button onclick="selectAllVisible(true)" style="padding:5px 10px; border:1px solid #66BB6A; background:white; border-radius:4px; cursor:pointer; font-size:12px;">全件選択</button>
    <button onclick="selectAllVisible(false)" style="padding:5px 10px; border:1px solid #ccc; background:white; border-radius:4px; cursor:pointer; font-size:12px;">選択クリア</button>
    <button onclick="selectToday()" style="padding:5px 10px; border:1px solid #66BB6A; background:white; border-radius:4px; cursor:pointer; font-size:12px;">今日確定分のみ</button>
    <span id="selectionCount" style="font-size:12px; color:#555; margin-left:8px;">選択0件</span>
    <button onclick="bulkExport('ne')" id="btnBulkNe"
            style="padding:6px 14px; border:none; background:#2E7D32; color:white; border-radius:4px; cursor:pointer; font-size:13px; font-weight:600; margin-left:auto;" disabled>
      🔄 NE受注CSVをまとめてダウンロード
    </button>
    <button onclick="bulkExport('coola')" id="btnBulkCoola"
            style="padding:6px 14px; border:none; background:#E65100; color:white; border-radius:4px; cursor:pointer; font-size:13px; font-weight:600;" disabled>
      🏭 COOLA CSVをまとめてダウンロード
    </button>
  </div>

  <!-- 検索バー -->
  <div class="search-bar" style="display:flex; gap:12px; margin-bottom:12px; padding:10px 14px; background:white; border-radius:6px; box-shadow:0 1px 3px rgba(0,0,0,0.08); align-items:center;">
    <span style="font-size:13px; color:#555; font-weight:600">🔍 検索:</span>
    <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#444;">
      取引先
      <input type="text" id="filterPartner" oninput="applyFilter()" placeholder="例: ネクスコ"
             style="padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; width:160px;">
    </label>
    <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#444;">
      納品日
      <input type="date" id="filterDeliveryDate" onchange="applyFilter()" oninput="applyFilter()"
             style="padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px;">
    </label>
    <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#444;">
      伝票No
      <input type="text" id="filterSlip" oninput="applyFilter()" placeholder="例: 00868"
             style="padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; width:140px;">
    </label>
    <button onclick="clearFilters()"
            style="padding:5px 12px; border:1px solid #ccc; background:#f8f8f8; border-radius:4px; cursor:pointer; font-size:12px;">
      クリア
    </button>
  </div>

  <div id="content">
    <div class="loading">読み込み中...</div>
  </div>
</div>

<!-- 詳細パネル -->
<div class="detail-overlay" id="detailOverlay" onclick="closeDetail()"></div>
<div class="detail-panel" id="detailPanel">
  <button class="close" onclick="closeDetail()">×</button>
  <div id="detailBody"></div>
</div>

<script>
let currentSource = '';
let currentOrders = [];   // 取得した全件
let filteredOrders = [];  // 検索フィルタ適用後（描画用）

async function loadOrders(source, isReload) {
  if (source !== undefined && source !== null) currentSource = source;

  document.querySelectorAll('.toolbar button').forEach(b => b.classList.remove('active'));
  const btnId = currentSource === '' ? 'btn-all' : `btn-${currentSource}`;
  const btn = document.getElementById(btnId);
  if (btn) btn.classList.add('active');

  document.getElementById('content').innerHTML = '<div class="loading">読み込み中...</div>';

  try {
    const days = document.getElementById('daysSelect').value;
    const params = new URLSearchParams();
    if (currentSource) params.set('source', currentSource);
    if (days) params.set('days', days);
    const url = '/api/confirmed_orders' + (params.toString() ? '?' + params.toString() : '');
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    currentOrders = data.orders || [];
    applyFilter();  // 既存検索条件を適用して描画
  } catch (err) {
    document.getElementById('content').innerHTML = `<div class="empty">取得失敗: ${err.message}</div>`;
  }
}

// 検索バー: 取引先・納品日・伝票No でクライアント側フィルタ
function applyFilter() {
  const partnerQ = (document.getElementById('filterPartner').value || '').toLowerCase().trim();
  const dateQ = (document.getElementById('filterDeliveryDate').value || '').trim();
  const slipQ = (document.getElementById('filterSlip').value || '').toLowerCase().trim();

  filteredOrders = currentOrders.filter(o => {
    if (partnerQ) {
      const partner = (o.partner_name || '').toLowerCase();
      if (!partner.includes(partnerQ)) return false;
    }
    if (dateQ) {
      // delivery_date は 'YYYY-MM-DD' 形式と仮定
      if ((o.delivery_date || '') !== dateQ) return false;
    }
    if (slipQ) {
      const slip = (o.slip_number || '').toLowerCase();
      if (!slip.includes(slipQ)) return false;
    }
    return true;
  });

  renderTable();
  const totalText = currentOrders.length;
  const visibleText = filteredOrders.length;
  const filterActive = partnerQ || dateQ || slipQ;
  document.getElementById('summary').textContent =
    filterActive ? `${visibleText}件 / 全${totalText}件` : `${totalText}件`;
}

function clearFilters() {
  document.getElementById('filterPartner').value = '';
  document.getElementById('filterDeliveryDate').value = '';
  document.getElementById('filterSlip').value = '';
  applyFilter();
}

function renderTable() {
  if (filteredOrders.length === 0) {
    if (currentOrders.length === 0) {
      document.getElementById('content').innerHTML = '<div class="empty">確定済み受注はありません</div>';
    } else {
      document.getElementById('content').innerHTML = '<div class="empty">該当する受注がありません（検索条件を変えてみてください）</div>';
    }
    return;
  }

  const rows = filteredOrders.map(o => {
    const channelBadge = `<span class="badge badge-${o.source_channel || 'unknown'}">${o.source_channel || '-'}</span>`;
    const confirmBadge = `<span class="badge badge-confirmed">確定</span>`;
    const countBadge = (o.confirmation_count || 0) > 1
      ? `<span class="badge-count">${o.confirmation_count}回</span>` : '';
    const total = o.grand_total != null ? Number(o.grand_total).toLocaleString() + '円' : '-';
    const partner = (o.partner_name || '').substring(0, 24);
    const dest = (o.delivery_location_name || '').substring(0, 30);
    const ddc = o.ddc_matched ? '✅' : '❌';
    const ship = o.shipping_fee_two_burden
      ? '<span style="color:#888">TWO負担</span>'
      : (o.shipping_fee != null ? `¥${Number(o.shipping_fee).toLocaleString()}` : '-');
    const confirmedAt = (o.confirmed_at || '').replace('T', ' ').substring(0, 16);
    const checked = selectedOrderIds.has(o.id) ? 'checked' : '';
    return `<tr data-id="${o.id}">
      <td onclick="event.stopPropagation()" style="text-align:center; padding:0 4px;">
        <input type="checkbox" ${checked} onchange="toggleOrderSelect('${o.id}', this.checked)"
               style="width:16px; height:16px; cursor:pointer;">
      </td>
      <td onclick="showDetail('${o.id}', this.parentElement)">${channelBadge}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)">${o.slip_number || '-'}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)">${o.delivery_date || '-'}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)">${partner}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)">${dest} ${ddc}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)">${total}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)" style="font-size:12px">${ship}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)">${confirmBadge} ${countBadge}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)" style="font-size:11px;color:#888">${confirmedAt}</td>
      <td onclick="showDetail('${o.id}', this.parentElement)" style="font-size:11px;color:#666">${o.confirmed_by || '-'}</td>
    </tr>`;
  }).join('');

  document.getElementById('content').innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="width:40px;text-align:center;"><input type="checkbox" id="selectAllChk" onchange="selectAllVisible(this.checked)"></th>
          <th>経路</th><th>伝票No</th><th>納品日</th><th>取引先</th><th>納品先(DDC)</th>
          <th>合計</th><th>送料</th><th>状態</th><th>確定日時</th><th>確定者</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;
  updateSelectionUI();
}

async function showDetail(orderId, rowEl) {
  document.querySelectorAll('tr.selected').forEach(r => r.classList.remove('selected'));
  if (rowEl) rowEl.classList.add('selected');

  document.getElementById('detailOverlay').style.display = 'block';
  document.getElementById('detailPanel').style.display = 'block';
  document.getElementById('detailBody').innerHTML = '<div class="loading">読み込み中...</div>';

  try {
    const res = await fetch(`/api/pending_orders/${orderId}`);  // /api/pending_orders/<id> は status 関係なく取得可能
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    renderDetail(data);
  } catch (err) {
    document.getElementById('detailBody').innerHTML = `<div class="empty">取得失敗: ${err.message}</div>`;
  }
}

function renderDetail(data) {
  const h = data.header;
  const items = data.items || [];

  const kv = (k, v) => `<div class="k">${k}</div><div class="v">${v != null && v !== '' ? v : '-'}</div>`;

  const editBtnHtml = `
    <div style="margin: 12px 0 18px; display: flex; gap: 10px; align-items: center;">
      <button onclick="openDraftEditor('${h.id}')"
              style="padding: 10px 20px; background: #FF9800; color: white; border: none;
                     border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer;">
        ✏️ 再編集する
      </button>
      <span style="color: #666; font-size: 12px">
        → 内容修正・再確定で発注書PDF/CSV を再生成
      </span>
    </div>`;

  let html = `
    <h2>${h.slip_number || '(伝票No未設定)'} <small style="color:#888">${h.source_channel}</small></h2>
    ${editBtnHtml}

    <div class="detail-section">
      <h3>確定情報</h3>
      <div class="kv">
        ${kv('ステータス', h.status)}
        ${kv('確定日時', (h.confirmed_at || '').replace('T', ' ').substring(0, 19))}
        ${kv('確定者', h.confirmed_by)}
        ${kv('確定回数', h.confirmation_count || 1)}
        ${kv('送料(税抜)', h.shipping_fee_two_burden ? 'TWO負担' :
          (h.shipping_fee != null ? Number(h.shipping_fee).toLocaleString() + '円' : '-'))}
      </div>
    </div>

    <div class="detail-section">
      <h3>受注情報</h3>
      <div class="kv">
        ${kv('発注日', h.order_date)}
        ${kv('納品日', h.delivery_date)}
        ${kv('取引先', h.partner_name)}
        ${kv('納品先', h.delivery_location_name)}
        ${kv('住所', h.delivery_location_address)}
        ${kv('DDCマッチ', h.ddc_matched ? '✅ マッチ' : '❌ 未マッチ')}
        ${kv('倉庫', h.warehouse)}
        ${kv('出荷区分', h.shipping_type)}
        ${kv('合計金額', h.grand_total != null ? Number(h.grand_total).toLocaleString() + '円' : '')}
      </div>
    </div>

    <div class="detail-section">
      <h3>明細 (${items.length}件)</h3>`;

  if (items.length === 0) {
    html += '<div style="color:#888">明細なし</div>';
  } else {
    html += items.map(i => `
      <div class="item-row">
        <div class="name">${i.product_name || '-'} ${i.product_master_matched ? '✅' : '❌'}</div>
        <div class="meta">
          JAN: ${i.jan_code || '-'} / ${i.quantity || 0} ${i.unit || ''} ×
          ${i.unit_price != null ? Number(i.unit_price).toLocaleString() + '円' : '?'}
          = ${i.amount != null ? Number(i.amount).toLocaleString() + '円' : '?'}
          ${i.spec ? ' / ' + i.spec : ''}
        </div>
      </div>`).join('');
  }
  html += '</div>';

  // 確定履歴（confirmation_history）
  const history = h.confirmation_history || [];
  if (history.length > 0) {
    html += `<div class="detail-section"><h3>確定履歴 (${history.length}回)</h3>`;
    history.slice().reverse().forEach((entry, idx) => {
      const at = (entry.at || '').replace('T', ' ').substring(0, 19);
      const by = entry.by || '-';
      const isLatest = idx === 0;
      const label = isLatest ? '最新' : `${history.length - idx}回目`;
      html += `<div class="history-row">
        <span class="at">[${label}]</span> ${at} by ${by}
      </div>`;
    });
    html += '</div>';
  }

  html += `
    <div style="font-size:11px;color:#999;margin-top:20px">
      ID: ${h.id}<br>
      作成: ${(h.created_at || '').replace('T', ' ').substring(0, 19)}<br>
      更新: ${(h.updated_at || '').replace('T', ' ').substring(0, 19)}
    </div>`;

  document.getElementById('detailBody').innerHTML = html;
}

function openDraftEditor(orderId) {
  // 確定済みでも /api/load_draft で読み込めば編集可能（再確定フロー）
  window.location.href = '/?draft_id=' + encodeURIComponent(orderId);
}

// ─── 一括CSV出力（複数受注をまとめてダウンロード） ───
const selectedOrderIds = new Set();

function toggleOrderSelect(orderId, checked) {
  if (checked) selectedOrderIds.add(orderId);
  else selectedOrderIds.delete(orderId);
  updateSelectionUI();
}

function selectAllVisible(check) {
  if (check) {
    filteredOrders.forEach(o => selectedOrderIds.add(o.id));
  } else {
    filteredOrders.forEach(o => selectedOrderIds.delete(o.id));
  }
  // チェックボックスのDOM状態を反映
  document.querySelectorAll('input[type="checkbox"][onchange^="toggleOrderSelect"]').forEach(chk => {
    const tr = chk.closest('tr');
    const id = tr ? tr.getAttribute('data-id') : null;
    if (id) chk.checked = selectedOrderIds.has(id);
  });
  const allChk = document.getElementById('selectAllChk');
  if (allChk) allChk.checked = check;
  updateSelectionUI();
}

function selectToday() {
  const today = new Date().toISOString().substring(0, 10);  // YYYY-MM-DD
  // 確定日時 (confirmed_at) が今日のレコードを選択
  selectedOrderIds.clear();
  filteredOrders.forEach(o => {
    const at = (o.confirmed_at || '').substring(0, 10);
    if (at === today) selectedOrderIds.add(o.id);
  });
  // 表示中のチェックボックス更新
  document.querySelectorAll('input[type="checkbox"][onchange^="toggleOrderSelect"]').forEach(chk => {
    const tr = chk.closest('tr');
    const id = tr ? tr.getAttribute('data-id') : null;
    if (id) chk.checked = selectedOrderIds.has(id);
  });
  updateSelectionUI();
}

function updateSelectionUI() {
  const n = selectedOrderIds.size;
  document.getElementById('selectionCount').textContent = `選択 ${n}件`;
  const enabled = n > 0;
  document.getElementById('btnBulkNe').disabled = !enabled;
  document.getElementById('btnBulkCoola').disabled = !enabled;
  // ボタンスタイル更新
  document.getElementById('btnBulkNe').style.opacity = enabled ? '1' : '0.5';
  document.getElementById('btnBulkCoola').style.opacity = enabled ? '1' : '0.5';
}

async function bulkExport(csvType) {
  const ids = Array.from(selectedOrderIds);
  if (ids.length === 0) {
    alert('受注を1件以上選択してください');
    return;
  }
  const btnId = csvType === 'ne' ? 'btnBulkNe' : 'btnBulkCoola';
  const btn = document.getElementById(btnId);
  const originalLabel = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '⏳ 生成中...';
  try {
    const res = await fetch('/api/bulk_csv_export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ order_ids: ids, csv_type: csvType }),
    });
    const data = await res.json();
    if (!res.ok || !data.success) {
      alert('生成失敗: ' + (data.error || ('HTTP ' + res.status)));
      return;
    }
    // ダウンロード（新規タブで開く）
    window.open(data.url, '_blank');
    let msg = `✅ ${csvType.toUpperCase()} CSV を生成しました（${data.processed}件）`;
    if (data.skipped) msg += ` / スキップ ${data.skipped}件`;
    setTimeout(() => alert(msg), 200);
  } catch (err) {
    alert('エラー: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalLabel;
    updateSelectionUI();
  }
}

function closeDetail() {
  document.getElementById('detailOverlay').style.display = 'none';
  document.getElementById('detailPanel').style.display = 'none';
  document.querySelectorAll('tr.selected').forEach(r => r.classList.remove('selected'));
}

document.addEventListener('keydown', e => { if (e.key === 'Escape') closeDetail(); });

// 初期ロード
loadOrders('');
</script>

</body>
</html>"""


# ─── HTML Template ───

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FAX受注処理システム</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', 'Yu Gothic UI', 'Meiryo', sans-serif; background: #f0f2f5; color: #333; }

.header { background: #2F5496; color: white; padding: 6px 14px; display: flex; align-items: center; justify-content: space-between; }
.header h1 { font-size: 17px; font-weight: 600; }
.header .status { font-size: 12px; opacity: 0.8; }

.upload-zone {
    margin: 10px 14px; padding: 20px; background: white; border: 2px dashed #aaa;
    border-radius: 6px; text-align: center; cursor: pointer; transition: all 0.2s;
}
.upload-zone:hover, .upload-zone.drag-over { border-color: #2F5496; background: #f0f4ff; }
.upload-zone input { display: none; }
.upload-zone p { color: #666; font-size: 14px; }

.main-area { display: flex; margin: 0 12px 12px; gap: 10px; height: calc(100vh - 140px); }
.main-area.hidden { display: none; }

.left-panel {
    flex: 0 0 36%; background: #1a1a1a; border-radius: 6px; overflow: hidden;
    display: flex; flex-direction: column;
}
.left-panel .page-nav {
    background: #2a2a2a; padding: 5px 8px; display: flex; align-items: center;
    justify-content: center; gap: 10px; color: #ccc; font-size: 12px;
}
.left-panel .page-nav button {
    background: #444; color: white; border: none; padding: 3px 10px; border-radius: 3px; cursor: pointer; font-size: 12px;
}
.left-panel .page-nav button:hover { background: #555; }
.left-panel .page-nav button:disabled { opacity: 0.3; cursor: default; }
.left-panel .preview-img {
    flex: 1; overflow: auto; display: flex; align-items: flex-start; justify-content: center; padding: 6px;
}
.left-panel .preview-img img { max-width: 100%; height: auto; }

.right-panel { flex: 1; overflow-y: auto; display: flex; flex-direction: column; gap: 8px; min-width: 0; }

.card {
    background: white; border-radius: 6px; padding: 10px 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    position: relative;
}
/* DDC検索カードはドロップダウンが他カードの上に被さるよう前面に */
.card.ddc-card { z-index: 50; }
.card.ddc-card.has-dropdown-open { z-index: 1000; }
.card h3 { font-size: 14px; color: #2F5496; margin-bottom: 6px; border-bottom: 1px solid #e0e0e0; padding-bottom: 4px; }

.field-grid { display: grid; grid-template-columns: 84px 1fr; gap: 4px 8px; align-items: center; }
.field-grid label { font-size: 12px; color: #666; font-weight: 600; text-align: right; }
.field-grid input, .field-grid select {
    padding: 5px 8px; border: 1px solid #ddd; border-radius: 3px; font-size: 13px; width: 100%;
}
.field-grid input:focus { outline: none; border-color: #2F5496; box-shadow: 0 0 0 2px rgba(47,84,150,0.15); }

.badge { display: inline-block; padding: 2px 8px; border-radius: 8px; font-size: 11px; font-weight: 600; }
.badge-ok { background: #E2EFDA; color: #2d7a2d; }
.badge-review { background: #FFF2CC; color: #b8860b; }
.badge-ng { background: #FCE4D6; color: #c0392b; }

/* DDC search dropdown */
.ddc-search-wrap { position: relative; }
.ddc-search-wrap input { width: 100%; }
.ddc-dropdown {
    position: absolute; top: 100%; left: 0; right: 0; z-index: 100;
    background: white; border: 1px solid #ccc; border-top: none; border-radius: 0 0 6px 6px;
    max-height: 240px; overflow-y: auto; box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    display: none;
}
.ddc-dropdown.show { display: block; }
.ddc-item {
    padding: 5px 10px; cursor: pointer; border-bottom: 1px solid #f0f0f0; transition: background 0.1s;
}
.ddc-item:hover, .ddc-item.active { background: #e8f0fe; }
.ddc-item .ddc-name { font-size: 13px; font-weight: 600; }
.ddc-item .ddc-addr { font-size: 11px; color: #888; margin-top: 2px; }

.items-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.items-table th { background: #f5f5f5; padding: 3px 4px; text-align: left; font-weight: 600; border-bottom: 2px solid #ddd; white-space: nowrap; font-size: 11px; }
.items-table td { padding: 3px 4px; border-bottom: 1px solid #eee; }
.items-table tr.matched { }
.items-table tr.unmatched { background: #FCE4D6; }
.items-table input[type="number"] { width: 44px; padding: 2px; border: 1px solid #ddd; border-radius: 3px; text-align: right; font-size: 12px; }
.items-table input[type="date"] { font-size: 11px; padding: 1px 2px; border: 1px solid #ddd; border-radius: 3px; max-width: 100%; }
.items-table select { width: 100%; max-width: 100%; min-width: 130px; font-size: 12px; padding: 2px; }
/* 列幅: 商品名は残り全部。他列は最小限。合計 ~530px に収めて余裕を持って右パネル内へ */
.items-table th:nth-child(1), .items-table td:nth-child(1) { min-width: 140px; }                      /* 商品名 */
.items-table th:nth-child(2), .items-table td:nth-child(2) { width: 92px;  white-space: nowrap; }     /* JAN */
.items-table th:nth-child(3), .items-table td:nth-child(3) { width: 50px;  white-space: nowrap; }     /* 数量 */
.items-table th:nth-child(4), .items-table td:nth-child(4) { width: 100px; white-space: nowrap; }     /* 賞味期限 */
.items-table th:nth-child(5), .items-table td:nth-child(5) { width: 60px;  white-space: nowrap; }     /* 梱包 */
.items-table th:nth-child(6), .items-table td:nth-child(6) { width: 54px;  white-space: nowrap; }     /* 出力先 */
.items-table th:nth-child(7), .items-table td:nth-child(7) { width: 40px;  text-align: center; }      /* マッチ */
.items-table th:nth-child(8), .items-table td:nth-child(8) { width: 26px; }                            /* 削除 */

.footer-bar {
    margin: 0 14px 8px; padding: 6px 12px; background: white; border-radius: 6px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1); display: flex; align-items: center; gap: 10px;
}
.footer-bar.hidden { display: none; }
.footer-bar label { font-size: 13px; color: #666; font-weight: 600; }
.footer-bar input { padding: 5px 10px; border: 1px solid #ddd; border-radius: 3px; font-size: 13px; }
.btn-confirm {
    margin-left: auto; background: #2F5496; color: white; border: none; padding: 6px 20px;
    border-radius: 5px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.2s;
}
.btn-confirm:hover { background: #1e3a6e; }
.btn-confirm:disabled { background: #aaa; cursor: default; }

.output-panel { margin: 0 24px 24px; }
.output-panel.hidden { display: none; }
.output-panel a {
    display: inline-block; margin: 4px 8px 4px 0; padding: 8px 16px; background: #2F5496;
    color: white; text-decoration: none; border-radius: 4px; font-size: 13px;
}
.output-panel a:hover { background: #1e3a6e; }

.spinner { display: inline-block; width: 18px; height: 18px; border: 3px solid #ccc; border-top-color: #2F5496; border-radius: 50%; animation: spin 0.8s linear infinite; margin-right: 8px; vertical-align: middle; }
@keyframes spin { to { transform: rotate(360deg); } }

.overlay {
    position: fixed; inset: 0; background: rgba(0,0,0,0.5); display: flex;
    align-items: center; justify-content: center; z-index: 200;
}
.overlay.hidden { display: none; }
.overlay-box { background: white; padding: 32px 48px; border-radius: 12px; text-align: center; font-size: 16px; }
</style>
</head>
<body>

<div class="header">
    <h1>FAX受注処理システム</h1>
    <span style="font-size:12px;color:#999;margin-left:12px" title="web_app.py の最終更新日時">ver. {{ app_last_updated }}</span>
    <div class="status" id="statusText"></div>
    <a href="/pending" style="background:#1976D2;color:#fff;border:none;border-radius:4px;padding:6px 14px;cursor:pointer;font-size:14px;margin-left:8px;text-decoration:none" title="Supabaseに保存された保留中受注を一覧表示">📋 保留中受注一覧</a>
    <a href="/confirmed" style="background:#1B5E20;color:#fff;border:none;border-radius:4px;padding:6px 14px;cursor:pointer;font-size:14px;margin-left:8px;text-decoration:none" title="確定済み受注を一覧表示・再編集可能">✅ 確定済み一覧</a>
    <button onclick="reloadMaster()" style="background:#FF8F00;color:#fff;border:none;border-radius:4px;padding:6px 14px;cursor:pointer;font-size:14px;margin-left:8px" title="Supabaseからマスタを再読込">マスタ更新</button>
</div>

<div class="upload-zone" id="uploadZone" onclick="document.getElementById('fileInput').click()">
    <input type="file" id="fileInput" accept=".pdf,.csv" multiple>
    <p>PDF / CSV（PALTAC・インフォマート）をドラッグ＆ドロップ、またはクリックして選択<br><span style="font-size:13px;color:#999">※ 複数CSVを同時選択すると1セッションにまとめて取込</span></p>
    <div style="margin-top:16px" onclick="event.stopPropagation()">
        <button onclick="startManualEntry()" style="padding:10px 24px;background:#2E7D32;color:#fff;border:none;border-radius:6px;font-size:15px;cursor:pointer">手入力で受注登録</button>
    </div>
</div>

<div class="main-area hidden" id="mainArea">
    <div class="left-panel">
        <div class="page-nav">
            <button id="prevBtn" onclick="changePage(-1)" disabled>&lt; 前</button>
            <span id="pageInfo">1 / 1</span>
            <button id="nextBtn" onclick="changePage(1)" disabled>次 &gt;</button>
            <button onclick="rotateImage()" style="margin-left:12px;padding:4px 10px;background:#666;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px" title="画像を180°回転">↻ 回転</button>
        </div>
        <div class="preview-img">
            <img id="previewImg" src="" alt="PDF Preview" style="transition:transform 0.3s">
        </div>
    </div>
    <div class="right-panel" id="rightPanel">
        <!-- Dynamically populated per page -->
    </div>
</div>

<!-- Phase 4: 送料計算パネル -->
<div id="shippingPanel" class="hidden" style="margin: 8px 14px; padding: 8px 12px; background: #FFF8E1; border-left: 4px solid #FFA726; border-radius: 6px; font-size: 12px;">
  <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
    <strong style="font-size:15px;">🚚 送料計算（ロット割れ自社倉庫出荷）</strong>
    <button onclick="recalculateShipping()"
            style="padding:4px 10px; border:1px solid #FFA726; background:white; border-radius:4px; cursor:pointer; font-size:12px;">
      🔄 再計算
    </button>
    <span id="shippingTotalDisplay" style="margin-left:auto; font-size:18px; font-weight:bold; color:#E65100;"></span>
  </div>
  <div id="shippingBreakdown" style="font-size:13px; color:#444;"></div>
  <div id="shippingManualInputs" style="margin-top:8px;"></div>
  <div style="margin-top:10px;">
    <label style="display:flex; align-items:center; gap:8px; cursor:pointer;">
      <input type="checkbox" id="shippingTwoBurden" onchange="updateShippingDisplay()">
      <span style="font-size:13px;">TWO負担とする（CSV「発送料/送料」に 0 を入れる）</span>
    </label>
  </div>
</div>

<div class="footer-bar hidden" id="footerBar">
    <label>担当者</label>
    <select id="staffInput" onchange="saveStaffSelection()" style="padding:6px 10px;border:1px solid #ccc;border-radius:4px;font-size:14px;min-width:120px">
      <!-- options are populated by JavaScript via /api/staff_list -->
    </select>
    <button class="btn-confirm" id="confirmBtn" onclick="doConfirm()">確定 &amp; PDF生成</button>
</div>

<div class="output-panel hidden" id="outputPanel">
    <div class="card">
        <h3>出力ファイル</h3>
        <div id="outputFiles"></div>
    </div>
</div>

<div class="overlay hidden" id="loadingOverlay">
    <div class="overlay-box">
        <div class="spinner"></div>
        <span id="loadingText">処理中...</span>
    </div>
</div>

<script>
// ─── State ───
let sessionId = null;
let pageResults = [];
let currentPage = 0;
// 商品名→賞味期限のメモリ（セッション中保持、ロット切替時は手動変更可）
const expiryMemory = {};
// 商品マスタリスト（プルダウン用）
let productMaster = [];
let ddcMaster = [];
// Phase 4: ドラフト編集モード時のメタ情報
let draftMode = null;  // null or { draft_id, draft_status, draft_confirmation_count, draft_confirmed_at, draft_slip_number }

// Phase 4: 担当者プルダウン用 localStorage キー
const STAFF_LS_KEY = 'fax_order_system.last_staff';

// ─── Init: load DDC list + 商品マスタ + 担当者一覧 + ドラフト ───
Promise.all([
    fetch('/api/ddc_list').then(r => r.json()),
    fetch('/api/product_list').then(r => r.json()),
    fetch('/api/staff_list').then(r => r.json()),
]).then(async ([ddcData, prodData, staffData]) => {
    ddcMaster = ddcData;
    productMaster = prodData;
    document.getElementById('statusText').textContent = `DDC: ${ddcData.length}件 / 商品: ${prodData.length}件 読込完了`;

    // Phase 4: 担当者プルダウン構築 + localStorage から前回値を復元
    populateStaffSelect(staffData.staff || []);

    // Phase 4: URL に ?draft_id=xxx があればドラフト編集モードで起動
    const params = new URLSearchParams(window.location.search);
    const draftId = params.get('draft_id');
    if (draftId) {
        await loadDraftEditor(draftId);
    }
});

// 担当者プルダウン構築
function populateStaffSelect(staffList) {
    const sel = document.getElementById('staffInput');
    if (!sel) return;
    sel.innerHTML = '';
    if (!staffList || staffList.length === 0) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = '(担当者未設定)';
        sel.appendChild(opt);
        return;
    }
    staffList.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s.name;
        opt.textContent = s.name;
        sel.appendChild(opt);
    });
    // localStorage から前回選択を復元（リスト内に存在する場合のみ）
    const last = localStorage.getItem(STAFF_LS_KEY);
    if (last && staffList.some(s => s.name === last)) {
        sel.value = last;
    }
}

// 担当者選択時に localStorage 保存
function saveStaffSelection() {
    const sel = document.getElementById('staffInput');
    if (sel && sel.value) {
        localStorage.setItem(STAFF_LS_KEY, sel.value);
    }
}

// ─── Phase 4: 送料計算パネル（ページ毎個別計算） ───
let lastShippingResults = null;  // {pages: [...], total_fee_all_pages: int}
const shippingManualOverrides = {};  // {pageIdx: {brand: int}}

async function recalculateShipping() {
    if (!pageResults || pageResults.length === 0) {
        document.getElementById('shippingPanel').classList.add('hidden');
        return;
    }

    // ページ毎に items + 住所 + 出荷区分 をまとめる
    const pagesPayload = pageResults.map(pr => {
        if (pr.error) return { items: [], ddc_address: '', shipping_type: null };
        return {
            items: (pr.matched_items || []).map(it => ({
                jan: it.jan || '',
                quantity: it.quantity || 0,
                matched: it.matched || false,
                master_name: it.master_name || '',
                ocr_name: it.ocr_name || '',
            })),
            ddc_address: pr.ddc_match ? (pr.ddc_match.address || '') : '',
            shipping_type: pr.warehouse_direct ? '自社倉庫' : null,
        };
    });

    const hasMatched = pagesPayload.some(p =>
        p.items.some(i => i.matched && (i.quantity || 0) > 0));
    if (!hasMatched) {
        document.getElementById('shippingPanel').classList.add('hidden');
        return;
    }

    try {
        const res = await fetch('/api/calculate_shipping_fee', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pages: pagesPayload }),
        });
        const data = await res.json();
        lastShippingResults = data;
        renderShippingPanel(data);
    } catch (err) {
        console.error('送料計算エラー:', err);
    }
}

function renderShippingPanel(result) {
    const panel = document.getElementById('shippingPanel');
    const breakdownDiv = document.getElementById('shippingBreakdown');
    const manualDiv = document.getElementById('shippingManualInputs');
    const totalDiv = document.getElementById('shippingTotalDisplay');

    panel.classList.remove('hidden');

    const pages = result.pages || [];

    if (pages.length === 0) {
        breakdownDiv.innerHTML = '<div style="color:#777;">該当ページがありません</div>';
        manualDiv.innerHTML = '';
        totalDiv.textContent = '¥0';
        updateShippingDisplay();
        return;
    }

    // ページ毎セクション（全ページ表示。送料発生なしページも一目で分かるように）
    let html = '';
    pages.forEach((p, pageIdx) => {
        const pr = (pageResults || [])[pageIdx] || {};
        const ddcName = pr.ddc_match ? (pr.ddc_match.name || '') : '';
        const ddcCode = pr.ddc_match
            ? (pr.ddc_match.nohinsaki_code || pr.ddc_match.code || '')
            : '';
        const destLabel = ddcName
            ? `${ddcName}${ddcCode ? ` [${ddcCode}]` : ''}`
            : (pr.ocr_raw && pr.ocr_raw.delivery_dest) || '(納品先未確定)';
        const addrLabel = p.prefecture
            ? `${p.prefecture} / ${p.zone || '-'}`
            : '<span style="color:#C62828">都道府県未確定</span>';

        const groups = p.groups || [];
        const lotBreakGroups = groups.filter(g => g.is_lot_break);
        const pageFee = (p.total_fee || 0);
        const headerColor = lotBreakGroups.length > 0 ? '#E65100' : '#2E7D32';
        const headerMark = lotBreakGroups.length > 0 ? '🚚' : '✅';
        const headerNote = lotBreakGroups.length > 0
            ? `<span style="color:${headerColor}; font-weight:600;">¥${pageFee.toLocaleString()}</span>`
            : `<span style="color:${headerColor};">送料なし（ロット成立）</span>`;

        html += `<div style="margin:6px 0; padding:8px; background:rgba(255,167,38,0.08); border-radius:4px;">
          <div style="margin-bottom:4px; font-size:13px;">
            <span style="font-weight:600;">${headerMark} P.${pageIdx + 1}</span>
            <span style="color:#333; margin-left:6px;">${destLabel}</span>
            <span style="color:#666; font-size:11px; margin-left:6px;">(${addrLabel})</span>
            <span style="margin-left:8px;">${headerNote}</span>
          </div>`;

        if (groups.length > 0) {
            html += `<table style="width:100%; border-collapse:collapse; font-size:12px;">
                <thead>
                  <tr style="background:#FFE0B2;">
                    <th style="padding:4px 8px; text-align:left;">brand</th>
                    <th style="padding:4px 8px; text-align:right;">CS</th>
                    <th style="padding:4px 8px; text-align:right;">送料(税抜)</th>
                    <th style="padding:4px 8px; text-align:left;">判定</th>
                  </tr>
                </thead><tbody>`;
            groups.forEach(g => {
                const fee = g.fee != null ? `¥${(g.fee || 0).toLocaleString()}` : '<em style="color:#C62828">手動入力必要</em>';
                const rowBg = g.is_lot_break ? '#FFFFFF' : '#F5F5F5';
                html += `<tr style="background:${rowBg};">
                  <td style="padding:4px 8px;">${g.brand}</td>
                  <td style="padding:4px 8px; text-align:right;">${g.total_cs}</td>
                  <td style="padding:4px 8px; text-align:right;">${fee}</td>
                  <td style="padding:4px 8px; color:#666;">${g.label || ''}</td>
                </tr>`;
            });
            html += '</tbody></table>';
        }
        if (p.warnings && p.warnings.length) {
            html += p.warnings.map(w => `<div style="color:#C62828;font-size:12px;margin-top:4px;">⚠️ ${w}</div>`).join('');
        }
        html += '</div>';
    });
    breakdownDiv.innerHTML = html;

    // 手動入力欄（ページ毎 × brand）
    let mHtml = '';
    pages.forEach((p, pageIdx) => {
        const manualGroups = (p.groups || []).filter(g => g.needs_manual);
        if (manualGroups.length === 0) return;
        if (!mHtml) {
            mHtml = '<div style="border-top:1px solid #FFD180; padding-top:8px; margin-top:8px;">';
            mHtml += '<div style="color:#BF360C; font-weight:600; margin-bottom:6px;">⚠️ 以下は外部アプリで計算 → 手動入力してください</div>';
        }
        manualGroups.forEach(g => {
            const url = g.external_tool_url || '';
            const stored = (shippingManualOverrides[pageIdx] || {})[g.brand];
            const value = stored != null ? stored : '';
            mHtml += `<div style="margin:4px 0; display:flex; gap:8px; align-items:center; font-size:13px;">
              <span style="font-weight:600;">P.${pageIdx + 1} ${g.brand}</span>
              <span>${g.total_cs}CS</span>
              <a href="${url}" target="_blank" style="color:#1565C0; text-decoration:underline;">🔗 計算アプリを開く</a>
              <input type="number" min="0" placeholder="送料(税抜)"
                     value="${value}"
                     oninput="setShippingManualOverride(${pageIdx}, '${g.brand}', this.value)"
                     style="width:100px; padding:3px 6px; border:1px solid #ccc; border-radius:4px;">
              <span>円</span>
            </div>`;
        });
    });
    if (mHtml) {
        mHtml += '</div>';
        manualDiv.innerHTML = mHtml;
    } else {
        manualDiv.innerHTML = '';
    }

    updateShippingDisplay();
}

function setShippingManualOverride(pageIdx, brand, value) {
    const n = parseInt(value);
    if (!isNaN(n) && n >= 0) {
        if (!shippingManualOverrides[pageIdx]) shippingManualOverrides[pageIdx] = {};
        shippingManualOverrides[pageIdx][brand] = n;
    } else if (shippingManualOverrides[pageIdx]) {
        delete shippingManualOverrides[pageIdx][brand];
        if (Object.keys(shippingManualOverrides[pageIdx]).length === 0) {
            delete shippingManualOverrides[pageIdx];
        }
    }
    updateShippingDisplay();
}

function updateShippingDisplay() {
    if (!lastShippingResults) return;
    const pages = lastShippingResults.pages || [];
    let total = pages.reduce((sum, p) => sum + (p.total_fee || 0), 0);
    pages.forEach((p, pageIdx) => {
        (p.groups || []).forEach(g => {
            if (g.needs_manual) {
                const v = (shippingManualOverrides[pageIdx] || {})[g.brand];
                if (v != null) total += v;
            }
        });
    });
    const twoBurden = document.getElementById('shippingTwoBurden').checked;
    const totalDiv = document.getElementById('shippingTotalDisplay');
    if (twoBurden) {
        totalDiv.innerHTML = `<span style="text-decoration:line-through; color:#888;">¥${total.toLocaleString()}</span> <span style="color:#1565C0;">→ TWO負担: ¥0</span>`;
    } else {
        totalDiv.textContent = `合計: ¥${total.toLocaleString()}（税抜）`;
    }
}

// Phase 4: ドラフト編集モードでセッション化 → 結果UI表示
async function loadDraftEditor(orderId) {
    showLoading('ドラフト読み込み中...');
    // 前回の状態を完全リセット（送料計算の手動入力値などが残らないように）
    resetShippingPanel();
    try {
        const res = await fetch('/api/load_draft', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ order_id: orderId }),
        });
        const data = await res.json();
        if (!res.ok || !data.session_id) {
            throw new Error(data.error || `HTTP ${res.status}`);
        }
        sessionId = data.session_id;
        draftMode = {
            draft_id: data.draft_id,
            draft_status: data.draft_status,
            draft_confirmation_count: data.draft_confirmation_count,
            draft_confirmed_at: data.draft_confirmed_at,
            draft_slip_number: data.draft_slip_number,
            filename: data.filename,
        };

        // OCR(展開済み)を取得して表示
        const ocrRes = await fetch('/api/ocr', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: sessionId }),
        });
        const ocrData = await ocrRes.json();
        pageResults = ocrData.pages;
        currentPage = 0;
        renderDraftBanner();
        showResults();
        // アップロードゾーンを隠す
        const uz = document.getElementById('uploadZone');
        if (uz) uz.style.display = 'none';
        // 「手入力で受注登録」も隠す（混乱回避）
        const startBtns = document.querySelectorAll('button[onclick="startManualEntry()"]');
        startBtns.forEach(b => b.style.display = 'none');
        // 確定ボタンの文言を「確定 & 再生成」に変更
        const confirmBtn = document.getElementById('confirmBtn');
        if (confirmBtn) confirmBtn.textContent = '確定 & 再生成';
        hideLoading();
    } catch (err) {
        hideLoading();
        alert('ドラフト読み込み失敗: ' + err.message + '\n\n通常モードに切り替えます。');
        // URL から draft_id を除去
        history.replaceState(null, '', window.location.pathname);
        draftMode = null;
    }
}

// Phase 4: 編集モード中の上部バナー
function renderDraftBanner() {
    if (!draftMode) return;
    let banner = document.getElementById('draftBanner');
    if (!banner) {
        banner = document.createElement('div');
        banner.id = 'draftBanner';
        banner.style.cssText = `
            background: linear-gradient(135deg, #FFF3E0 0%, #FFE0B2 100%);
            border: 2px solid #F57C00;
            border-radius: 8px;
            padding: 12px 18px;
            margin: 12px 24px;
            font-size: 14px;
            color: #E65100;
            display: flex;
            align-items: center;
            gap: 10px;
            box-shadow: 0 2px 4px rgba(245,124,0,0.2);
        `;
        // ヘッダー直下に挿入
        const header = document.querySelector('.header');
        if (header && header.nextElementSibling) {
            header.parentNode.insertBefore(banner, header.nextElementSibling);
        } else {
            document.body.insertBefore(banner, document.body.firstChild);
        }
    }
    const slipText = draftMode.draft_slip_number ? `伝票No. ${draftMode.draft_slip_number}` : '(伝票No未設定)';
    const statusBadge = draftMode.draft_status === 'confirmed'
        ? `<span style="background:#4CAF50;color:white;padding:2px 8px;border-radius:10px;font-size:11px;">CONFIRMED</span>`
        : draftMode.draft_status === 'error'
        ? `<span style="background:#D32F2F;color:white;padding:2px 8px;border-radius:10px;font-size:11px;">ERROR</span>`
        : `<span style="background:#FFA726;color:white;padding:2px 8px;border-radius:10px;font-size:11px;">DRAFT</span>`;
    const countText = draftMode.draft_confirmation_count > 0
        ? `<span style="margin-left:auto;font-size:12px;color:#BF360C">確定回数: ${draftMode.draft_confirmation_count}回 / 最終確定: ${(draftMode.draft_confirmed_at || '').replace('T',' ').substring(0,16)}</span>`
        : '';
    banner.innerHTML = `
        <span style="font-size:18px">📝</span>
        <strong>ドラフト編集モード</strong>
        ${statusBadge}
        <span style="color:#5D4037">${slipText}</span>
        <span style="font-size:11px;color:#888">(${draftMode.filename})</span>
        ${countText}
    `;
}

// ─── File Upload ───
const uploadZone = document.getElementById('uploadZone');
const fileInput = document.getElementById('fileInput');

uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', e => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) handleFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', () => { if (fileInput.files.length > 0) handleFiles(fileInput.files); });

async function handleFiles(files) {
    const fileArr = Array.from(files);
    const allCsv = fileArr.length > 1 && fileArr.every(f => f.name.toLowerCase().endsWith('.csv'));
    // Phase 4: 新規アップロード前に送料計算状態をリセット
    resetShippingPanel();

    try {
        let data;
        if (allCsv) {
            // 複数CSV一括
            showLoading(`複数CSV一括アップロード中... (${fileArr.length}ファイル)`);
            const form = new FormData();
            for (const f of fileArr) form.append('files', f);
            const res = await fetch('/api/upload_multi_csv', { method: 'POST', body: form });
            data = await res.json();
        } else {
            // 単体ファイル
            const file = fileArr[0];
            showLoading(`アップロード中: ${file.name}`);
            const form = new FormData();
            form.append('file', file);
            const res = await fetch('/api/upload', { method: 'POST', body: form });
            data = await res.json();
        }
        if (data.error) { alert(data.error); hideLoading(); return; }

        sessionId = data.session_id;

        if (data.auto_results) {
            showLoading(`CSV取込中... (${data.page_count}件)`);
            const ocrRes = await fetch('/api/ocr', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: sessionId }),
            });
            const ocrData = await ocrRes.json();
            if (ocrData.error) { alert(ocrData.error); hideLoading(); return; }
            pageResults = ocrData.pages;
        } else {
            showLoading(`OCR処理中... (${data.page_count}ページ)`);
            const ocrRes = await fetch('/api/ocr', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: sessionId }),
            });
            const ocrData = await ocrRes.json();
            if (ocrData.error) { alert(ocrData.error); hideLoading(); return; }
            pageResults = ocrData.pages;
        }

        currentPage = 0;
        showResults();
        hideLoading();
    } catch (err) {
        alert('エラー: ' + err.message);
        hideLoading();
    }
}

// ─── 備考自動入力ルール ───
// DDC毎に「条件成立 → 備考に自動入力」するルール
const DDC_REMARKS_RULES = [
    { ddcCode: '880', requiresOutput: 'シルビア', text: '運送NO.をご案内ください。' },
];

// ─── 二重梱包 不可 DDC ───
// この納品先コードに対しては二重梱包ボタンを無効化し、強制的に「通常」固定にする
const NO_DOUBLE_PACK_DDC_CODES = [
    '976', // 北海道江別物流C（伊藤忠商事日本アクセス）
    '977', // ドンキ仙台港PDC（伊藤忠商事日本アクセス）
    '978', // ドンキ浦和預けC（伊藤忠商事日本アクセス）
    '979', // ドンキ加須預けC（伊藤忠商事日本アクセス）
    '980', // PPIH中部共配C（伊藤忠商事日本アクセス）
    '981', // PPIH山静共配C（日本アクセス）
    '982', // PPIH北陸共配C（伊藤忠商事日本アクセス）
    '983', // ドンキ泉北PDC（伊藤忠商事日本アクセス）
];

function isNoDoublePackPage(pageIdx) {
    const pr = pageResults[pageIdx];
    if (!pr || !pr.ddc_match) return false;
    const code = String(pr.ddc_match.nohinsaki_code || pr.ddc_match.code || '');
    return NO_DOUBLE_PACK_DDC_CODES.includes(code);
}

function applyNoDoublePackRule(pageIdx) {
    if (!isNoDoublePackPage(pageIdx)) return;
    const pr = pageResults[pageIdx];
    if (!pr || !pr.matched_items) return;
    let changed = false;
    pr.matched_items.forEach(it => {
        if (it.double_pack) {
            it.double_pack = false;
            changed = true;
        }
    });
    if (changed && pageIdx === currentPage) renderPage(currentPage);
}

function applyAutoRemarks(pageIdx) {
    const pr = pageResults[pageIdx];
    if (!pr || pr.error) return;
    // 既に備考が入力済みなら上書きしない
    if (pr.remarks && pr.remarks.trim()) return;
    const ddcCode = pr.ddc_match
        ? String(pr.ddc_match.nohinsaki_code || pr.ddc_match.code || '')
        : '';
    if (!ddcCode) return;
    for (const rule of DDC_REMARKS_RULES) {
        if (ddcCode !== rule.ddcCode) continue;
        const hasOutput = (pr.matched_items || []).some(it =>
            it.matched && it.output_dest === rule.requiresOutput);
        if (!hasOutput) continue;
        pr.remarks = rule.text;
        if (pageIdx === currentPage) renderPage(currentPage);
        break;
    }
}

// ─── Display Results ───
function showResults() {
    document.getElementById('mainArea').classList.remove('hidden');
    document.getElementById('footerBar').classList.remove('hidden');
    document.getElementById('uploadZone').style.display = 'none';
    // 備考自動入力（DDC + 出力先ルール）を全ページ適用
    (pageResults || []).forEach((_, i) => applyAutoRemarks(i));
    // 二重梱包不可DDC ルールも全ページ適用
    (pageResults || []).forEach((_, i) => applyNoDoublePackRule(i));
    renderPage(currentPage);
    updatePageNav();
    // Phase 4: 結果表示時に送料を自動計算
    recalculateShipping();
}

function changePage(delta) {
    currentPage += delta;
    if (currentPage < 0) currentPage = 0;
    if (currentPage >= pageResults.length) currentPage = pageResults.length - 1;
    // ページ切替時に回転リセット
    imageRotation = 0;
    document.getElementById('previewImg').style.transform = '';
    renderPage(currentPage);
    updatePageNav();
}

function updatePageNav() {
    document.getElementById('pageInfo').textContent = `${currentPage + 1} / ${pageResults.length}`;
    document.getElementById('prevBtn').disabled = currentPage <= 0;
    document.getElementById('nextBtn').disabled = currentPage >= pageResults.length - 1;
}

function renderPage(idx) {
    const pr = pageResults[idx];
    // Preview image (PALTAC CSV / 手入力 / ドラフト編集モードはプレビューなし)
    const previewImg = document.getElementById('previewImg');
    const previewContainer = previewImg.parentElement;
    let placeholder = document.getElementById('previewPlaceholder');
    if (pr.source === 'paltac' || pr.source === 'infomart' || pr.source === 'manual' || pr.source === 'smacla' || draftMode) {
        previewImg.src = '';
        previewImg.style.display = 'none';
        // プレースホルダーメッセージを表示
        if (!placeholder) {
            placeholder = document.createElement('div');
            placeholder.id = 'previewPlaceholder';
            placeholder.style.cssText = 'color:#aaa;text-align:center;padding:40px 20px;font-size:13px;line-height:1.6;';
            previewContainer.appendChild(placeholder);
        }
        let msg;
        if (draftMode) {
            msg = '📝 ドラフト編集モード<br><span style="font-size:11px">元PDF プレビューは非表示<br>(Supabase保存済みのデータを編集中)</span>';
        } else if (pr.source === 'manual') {
            msg = '✏️ 手入力モード<br><span style="font-size:11px">PDF元データなし</span>';
        } else {
            msg = '📊 CSV取込<br><span style="font-size:11px">PDF元データなし</span>';
        }
        placeholder.innerHTML = msg;
        placeholder.style.display = '';
    } else {
        previewImg.style.display = '';
        previewImg.src = `/api/page_image/${sessionId}/${pr.page}`;
        if (placeholder) placeholder.style.display = 'none';
    }

    const panel = document.getElementById('rightPanel');

    if (pr.error) {
        panel.innerHTML = `<div class="card"><h3>エラー</h3><p>${pr.error}</p></div>`;
        return;
    }

    const ocr = pr.ocr_raw;
    const ddc = pr.ddc_match;
    const items = pr.matched_items;
    const lowConf = ddc.low_confidence || false;
    const matched = ddc.matched || false;
    let status = matched ? (lowConf ? '要確認' : 'OK') : 'NG';
    let badgeClass = status === 'OK' ? 'badge-ok' : status === '要確認' ? 'badge-review' : 'badge-ng';

    let candidatesHtml = '';
    if (ddc.candidates && ddc.candidates.length > 0) {
        candidatesHtml = '<div style="margin-top:8px;font-size:14px;color:#666;">候補: ';
        for (const c of ddc.candidates) {
            candidatesHtml += `<span style="cursor:pointer;text-decoration:underline;margin-right:10px;" onclick="selectCandidate(${idx},'${c.name.replace(/'/g, "\\'")}')">${c.name} (${Math.round(c.score*100)}%)</span>`;
        }
        candidatesHtml += '</div>';
    }

    let itemsRows = '';
    for (let i = 0; i < items.length; i++) {
        const it = items[i];
        // 記憶済みの賞味期限を自動適用（未入力の場合のみ）
        const masterName = it.matched ? it.master_name : it.ocr_name;
        if ((!it.expiry_date || it.expiry_date === '') && masterName && expiryMemory[masterName]) {
            it.expiry_date = expiryMemory[masterName];
        }
        if (!it.expiry_date) it.expiry_date = '';
        if (it.double_pack === undefined) it.double_pack = false;
        const cls = it.matched ? 'matched' : 'unmatched';
        const name = it.matched ? it.master_name : it.ocr_name;
        const matchBadge = it.matched ? '<span class="badge badge-ok">OK</span>' : '<span class="badge badge-ng">NG</span>';
        // 二重梱包 不可DDCなら強制リセット + ボタン無効化
        const dpDisabled = isNoDoublePackPage(idx);
        if (dpDisabled) it.double_pack = false;
        const dpChecked = it.double_pack ? 'checked' : '';
        let dpStyle, dpLabel, dpTitle, dpAttr;
        if (dpDisabled) {
            dpStyle = 'background:#ccc;color:#666;cursor:not-allowed;text-decoration:line-through;';
            dpLabel = '不可';
            dpTitle = 'この納品先(DDC)は二重梱包不可です';
            dpAttr = 'disabled';
        } else {
            dpStyle = it.double_pack ? 'background:#E65100;color:#fff;' : 'background:#eee;color:#999;';
            dpLabel = it.double_pack ? '二重梱包' : '通常';
            dpTitle = '';
            dpAttr = '';
        }
        // 商品プルダウン生成
        let prodOptions = `<option value="">-- 選択 --</option>`;
        for (const pm of productMaster) {
            const sel = (pm.name === name) ? ' selected' : '';
            prodOptions += `<option value="${pm.name}"${sel}>${pm.name}</option>`;
        }
        itemsRows += `<tr class="${cls}">
            <td><select data-page="${idx}" data-item="${i}" onchange="updateProduct(this)" style="font-size:13px;width:100%;padding:3px">${prodOptions}</select></td>
            <td>${it.jan || ''}</td>
            <td><input type="number" value="${it.quantity || 0}" min="0" data-page="${idx}" data-item="${i}" onchange="updateQty(this)" style="width:60px"></td>
            <td><input type="date" value="${it.expiry_date}" data-page="${idx}" data-item="${i}" onchange="updateExpiry(this)" style="width:140px;font-size:14px"></td>
            <td><button onclick="toggleDoublePack(${idx},${i},this)" ${dpAttr} title="${dpTitle}" style="border:none;border-radius:4px;padding:5px 10px;font-size:13px;cursor:pointer;${dpStyle}" id="dpBtn-${idx}-${i}">${dpLabel}</button></td>
            <td>${it.output_dest || ''}</td>
            <td>${matchBadge}</td>
            <td><button onclick="removeProductRow(${idx},${i})" style="background:#d32f2f;color:#fff;border:none;border-radius:3px;padding:2px 8px;cursor:pointer;font-size:12px">✕</button></td>
        </tr>`;
    }

    panel.innerHTML = `
    <div class="card">
        <h3>注文情報 <span class="badge ${badgeClass}" id="ddcBadge-${idx}">${status}</span></h3>
        <div class="field-grid">
            <label>オーダーNO</label>
            <input type="text" value="${ocr.order_no || pr.two_order_no || ''}" data-page="${idx}" data-field="order_no" onchange="editField(this)">
            <label>TWO受注NO</label>
            <input type="text" value="${pr.two_order_no || ''}" data-page="${idx}" data-field="two_order_no" onchange="editTwoOrderNo(this)" style="background:#FFF8E1;font-weight:bold">
            <label>納品日</label>
            <input type="date" value="${ocr.delivery_date || ''}" data-page="${idx}" data-field="delivery_date" onchange="editField(this)">
            <label>発注元</label>
            <input type="text" value="${ocr.sender || ''}" readonly style="background:#f5f5f5">
            <label>納品先(OCR)</label>
            <input type="text" value="${ocr.delivery_dest || ''}" readonly style="background:#f5f5f5">
            <label>備考</label>
            <input type="text" value="${pr.remarks || ''}" data-page="${idx}" data-field="remarks" onchange="editRemarks(this)" placeholder="この注文の備考">
            <label>自社倉庫向け</label>
            <div style="display:flex;align-items:center;gap:8px">
                <input type="checkbox" id="warehouseChk-${idx}" ${pr.warehouse_direct ? 'checked' : ''} onchange="toggleWarehouse(${idx}, this)" style="width:20px;height:20px">
                <span style="font-size:13px;color:#666">${pr.warehouse_direct ? 'ベルーナ宛（CSV出力なし）' : 'チェックでベルーナ宛に設定'}</span>
            </div>
        </div>
    </div>
    <div class="card ddc-card" id="ddcCard-${idx}">
        <h3>納品先(DDC) マッチング</h3>
        <div class="ddc-search-wrap">
            <input type="text" id="ddcSearch-${idx}" value="${ddc.name || ''}"
                   placeholder="納品先を入力して検索..." autocomplete="off"
                   oninput="filterDdc(${idx}, this.value)"
                   onfocus="filterDdc(${idx}, this.value)"
                   data-page="${idx}">
            <div class="ddc-dropdown" id="ddcDropdown-${idx}"></div>
        </div>
        ${candidatesHtml}
        <div style="margin-top:8px;font-size:14px;color:#888;" id="ddcInfo-${idx}">
            ${ddc.address ? '住所: ' + ddc.address : ''}
            ${ddc.tel ? ' / TEL: ' + ddc.tel : ''}
        </div>
    </div>
    <div class="card">
        <h3>商品明細</h3>
        <table class="items-table">
            <thead><tr><th>商品名</th><th>JAN</th><th>数量(CS)</th><th>賞味期限</th><th>梱包</th><th>出力先</th><th>マッチ</th><th></th></tr></thead>
            <tbody>${itemsRows}</tbody>
        </table>
        <button onclick="addProductRow(${idx})" style="margin-top:8px;padding:6px 16px;background:#1565C0;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:14px">＋ 商品追加</button>
    </div>`;
}

// ─── DDC Search ───
let ddcDebounce = null;
function filterDdc(pageIdx, query) {
    clearTimeout(ddcDebounce);
    ddcDebounce = setTimeout(() => _filterDdc(pageIdx, query), 100);
}

function _filterDdc(pageIdx, query) {
    const dd = document.getElementById(`ddcDropdown-${pageIdx}`);
    const card = document.getElementById(`ddcCard-${pageIdx}`);
    if (!query || query.length < 1) {
        dd.classList.remove('show');
        if (card) card.classList.remove('has-dropdown-open');
        return;
    }

    const q = query.toLowerCase();
    const matches = ddcMaster.filter(d =>
        d.name.toLowerCase().includes(q) || d.address.toLowerCase().includes(q)
    ).slice(0, 30);

    if (matches.length === 0) {
        dd.innerHTML = '<div class="ddc-item" style="color:#999">該当なし</div>';
        dd.classList.add('show');
        if (card) card.classList.add('has-dropdown-open');
        return;
    }

    dd.innerHTML = matches.map((d, i) => `
        <div class="ddc-item" data-idx="${i}">
            <div class="ddc-name">${highlightMatch(d.name, q)}</div>
            <div class="ddc-addr">${d.address}</div>
        </div>
    `).join('');
    // クリックハンドラを直接バインド（特殊文字エスケープの問題を回避）
    dd.querySelectorAll('.ddc-item[data-idx]').forEach((el, i) => {
        el.addEventListener('click', (e) => {
            e.stopPropagation();
            selectDdc(pageIdx, matches[i].name);
        });
    });
    dd.classList.add('show');
    if (card) card.classList.add('has-dropdown-open');
}

function highlightMatch(text, query) {
    const idx = text.toLowerCase().indexOf(query);
    if (idx < 0) return text;
    return text.substring(0, idx) + '<b style="color:#2F5496">' + text.substring(idx, idx + query.length) + '</b>' + text.substring(idx + query.length);
}

function selectDdc(pageIdx, name) {
    const input = document.getElementById(`ddcSearch-${pageIdx}`);
    const dd = document.getElementById(`ddcDropdown-${pageIdx}`);
    const card = document.getElementById(`ddcCard-${pageIdx}`);
    input.value = name;
    dd.classList.remove('show');
    if (card) card.classList.remove('has-dropdown-open');

    // Store selection in pageResults
    const entry = ddcMaster.find(d => d.name === name);
    if (entry) {
        pageResults[pageIdx].ddc_match.name = name;
        pageResults[pageIdx].ddc_match.matched = true;
        pageResults[pageIdx].ddc_match.low_confidence = false;
        pageResults[pageIdx].ddc_match.address = entry.address;
        pageResults[pageIdx].ddc_match.tel = entry.tel;
        pageResults[pageIdx].ddc_match.nohinsaki_code = entry.nohinsaki_code || '';
        pageResults[pageIdx]._userDdc = name;

        document.getElementById(`ddcInfo-${pageIdx}`).textContent =
            `住所: ${entry.address} / TEL: ${entry.tel}`;
        const badge = document.getElementById(`ddcBadge-${pageIdx}`);
        if (badge) { badge.className = 'badge badge-ok'; badge.textContent = 'OK'; }
        applyAutoRemarks(pageIdx);
        applyNoDoublePackRule(pageIdx);
        recalculateShipping();
    }
}

function selectCandidate(pageIdx, name) {
    selectDdc(pageIdx, name);
    document.getElementById(`ddcSearch-${pageIdx}`).value = name;
}

// Close dropdown when clicking outside
document.addEventListener('click', e => {
    if (!e.target.closest('.ddc-search-wrap')) {
        document.querySelectorAll('.ddc-dropdown').forEach(d => d.classList.remove('show'));
        document.querySelectorAll('.ddc-card.has-dropdown-open').forEach(c => c.classList.remove('has-dropdown-open'));
    }
});

// ─── Edit fields ───
function editField(el) {
    const idx = parseInt(el.dataset.page);
    const field = el.dataset.field;
    pageResults[idx].ocr_raw[field] = el.value;
}

function updateQty(el) {
    const idx = parseInt(el.dataset.page);
    const itemIdx = parseInt(el.dataset.item);
    const qty = parseInt(el.value) || 0;
    const item = pageResults[idx].matched_items[itemIdx];
    item.quantity = qty;
    if (item.cs_price) {
        item.amount = qty * item.cs_price;
        const amtCell = document.querySelector(`.item-amount-${idx}-${itemIdx}`);
        if (amtCell) amtCell.textContent = '¥' + item.amount.toLocaleString();
    }
    recalculateShipping();
}
function editTwoOrderNo(el) {
    const idx = parseInt(el.dataset.page);
    pageResults[idx].two_order_no = el.value;
}
function editRemarks(el) {
    const idx = parseInt(el.dataset.page);
    pageResults[idx].remarks = el.value;
}
function toggleWarehouse(pageIdx, chk) {
    const pr = pageResults[pageIdx];
    pr.warehouse_direct = chk.checked;
    if (chk.checked) {
        // ベルーナの納品先情報を自動設定
        pr._userDdc = '株式会社ベルーナ・ジーエフ・ロジスティクス';
        pr.ddc_match = {
            matched: true,
            name: '株式会社ベルーナ・ジーエフ・ロジスティクス',
            postal: '3620066',
            address: '埼玉県上尾市領家丸山30-1',
            tel: '048-725-0179',
            fax: '',
            match_score: 1.0,
            low_confidence: false,
            candidates: [],
        };
    }
    renderPage(pageIdx);
    recalculateShipping();
}
function updateProduct(el) {
    const idx = parseInt(el.dataset.page);
    const itemIdx = parseInt(el.dataset.item);
    const selectedName = el.value;
    const item = pageResults[idx].matched_items[itemIdx];
    // 商品マスタから該当商品の情報を取得して更新
    const pm = productMaster.find(p => p.name === selectedName);
    if (pm) {
        item.matched = true;
        item.master_name = pm.name;
        item.jan = pm.jan || '';
        item.code = pm.code || '';
        item.output_dest = pm.output_dest || '';
        item.case_quantity = pm.case_quantity || 0;
        item.cs_price = pm.cs_price || 0;
        item.spec = pm.spec || '';
        item.pack = pm.pack || '';
        item.unit_price = pm.unit_price || 0;
        if (item.quantity && item.cs_price) {
            item.amount = item.quantity * item.cs_price;
        }
        // 賞味期限: メモリから自動適用させるためリセット
        delete item.expiry_date;
    }
    applyAutoRemarks(idx);
    renderPage(idx);
    recalculateShipping();
}
function updateExpiry(el) {
    const idx = parseInt(el.dataset.page);
    const itemIdx = parseInt(el.dataset.item);
    const item = pageResults[idx].matched_items[itemIdx];
    item.expiry_date = el.value;
    // 商品名→賞味期限をメモリに保存（次の注文で自動適用）
    const name = item.master_name || item.ocr_name;
    if (name && el.value) {
        expiryMemory[name] = el.value;
    }
}
function toggleDoublePack(pageIdx, itemIdx, btn) {
    if (isNoDoublePackPage(pageIdx)) return;  // 不可DDCでは無視
    const item = pageResults[pageIdx].matched_items[itemIdx];
    item.double_pack = !item.double_pack;
    if (item.double_pack) {
        btn.textContent = '二重梱包';
        btn.style.background = '#E65100';
        btn.style.color = '#fff';
    } else {
        btn.textContent = '通常';
        btn.style.background = '#eee';
        btn.style.color = '#999';
    }
}

// ─── Confirm ───
async function doConfirm() {
    const btn = document.getElementById('confirmBtn');

    // Phase 4: 再確定（confirmation_count > 0）の場合は警告ダイアログ
    if (draftMode && (draftMode.draft_confirmation_count || 0) > 0) {
        const lastAt = (draftMode.draft_confirmed_at || '').replace('T', ' ').substring(0, 16);
        const ok = confirm(
            `⚠️ このオーダーは既に確定済みです（${draftMode.draft_confirmation_count}回目 / ${lastAt}）。\n\n` +
            '再確定すると新しい発注書PDF・NE/COOLA CSV が生成されます（旧版は履歴フォルダへ自動退避）。\n\n' +
            'もし旧版のFAXを既に送付済みの場合、二重発注にならないようご注意ください。\n\n' +
            '続行しますか？'
        );
        if (!ok) return;
    }

    btn.disabled = true;
    showLoading(draftMode ? '確定&再生成中...' : '発注PDF生成中...');

    const pages = pageResults.map((pr, idx) => ({
        page: pr.page,
        ddc_name: pr._userDdc || pr.ddc_match.name || '',
        order_no: pr.ocr_raw.order_no || '',
        two_order_no: pr.two_order_no || '',
        delivery_date: pr.ocr_raw.delivery_date || '',
        remarks: pr.remarks || '',
        warehouse_direct: pr.warehouse_direct || false,
        items: pr.matched_items.map((it, i) => ({
            index: i, quantity: it.quantity, expiry_date: it.expiry_date || '', double_pack: it.double_pack || false,
            master_name: it.master_name || '', jan: it.jan || '', code: it.code || '',
            output_dest: it.output_dest || '', case_quantity: it.case_quantity || 0,
            cs_price: it.cs_price || 0, spec: it.spec || '', pack: it.pack || '',
            unit_price: it.unit_price || 0, matched: it.matched || false,
        })),
    }));

    try {
        const res = await fetch('/api/confirm', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                session_id: sessionId,
                staff_name: document.getElementById('staffInput').value,
                remarks: '',
                pages: pages,
                // Phase 4: 送料計算パネルからの値を反映
                shipping_fee_two_burden: document.getElementById('shippingTwoBurden').checked,
                shipping_fee_manual_overrides: shippingManualOverrides,
            }),
        });
        const data = await res.json();
        if (data.error) { alert(data.error); hideLoading(); btn.disabled = false; return; }

        // Show output files
        const outDiv = document.getElementById('outputFiles');
        let html = data.files.map(f => {
            const icon = f.type === 'pdf' ? '📄' : f.type === 'xlsx' ? '📊' : f.type === 'ne_csv' ? '🔄' : f.type === 'coola_csv' ? '🏭' : '📝';
            const label = f.label ? `${f.label} ` : '';
            const style = (f.type === 'ne_csv' || f.type === 'coola_csv') ? ' style="background:#E65100;font-weight:bold"' : '';
            return `<a href="/api/download/${encodeURIComponent(f.name)}" target="_blank"${style}>${icon} ${label}${f.name}</a>`;
        }).join('');
        // Phase 4: 送料サマリ + 確定情報を表示
        if (data.shipping_fee_total || data.shipping_fee_breakdown) {
            const fee = data.shipping_fee_total || 0;
            const tw = data.shipping_fee_two_burden ? '（TWO負担、CSVには未反映）' : '';
            html += `<div style="margin-top:10px;padding:8px 12px;background:#F0F4FF;border-radius:4px;font-size:13px;">🚚 送料: ¥${fee.toLocaleString()}（税抜）${tw}</div>`;
        }
        if (data.draft_updated) {
            html += `<div style="margin-top:6px;padding:8px 12px;background:#E8F5E9;border-radius:4px;font-size:13px;">✅ Supabaseの確定状態を更新しました（確定回数: ${data.confirmation_count}回目）</div>`;
        }
        const nextLabel = draftMode ? '一覧に戻る' : '次のPDFを処理';
        const nextHandler = draftMode ? "window.location.href='/pending'" : "resetForNext()";
        html += `<br><button onclick="${nextHandler}" style="margin-top:12px;padding:10px 28px;background:#2E7D32;color:#fff;border:none;border-radius:6px;font-size:15px;cursor:pointer">${nextLabel}</button>`;
        outDiv.innerHTML = html;
        document.getElementById('outputPanel').classList.remove('hidden');
        hideLoading();
        btn.disabled = false;
    } catch (err) {
        alert('エラー: ' + err.message);
        hideLoading();
        btn.disabled = false;
    }
}

// ─── Reload master data ───
async function reloadMaster() {
    document.getElementById('statusText').textContent = 'マスタ再読込中...';
    try {
        const res = await fetch('/api/reload_master', { method: 'POST' });
        const data = await res.json();
        if (data.success) {
            document.getElementById('statusText').textContent = data.message;
            // DDCリスト・商品リストも再取得
            const [ddcRes, prodRes] = await Promise.all([
                fetch('/api/ddc_list'),
                fetch('/api/product_list'),
            ]);
            ddcMaster = await ddcRes.json();
            productMaster = await prodRes.json();
        } else {
            document.getElementById('statusText').textContent = 'マスタ更新失敗';
        }
    } catch (err) {
        document.getElementById('statusText').textContent = 'マスタ更新エラー: ' + err.message;
    }
}

// ─── Reset for next PDF ───
function resetForNext() {
    sessionId = null;
    pageResults = [];
    currentPage = 0;
    totalPages = 0;
    // expiryMemory は保持（賞味期限の引継ぎ）
    document.getElementById('outputPanel').classList.add('hidden');
    document.getElementById('outputFiles').innerHTML = '';
    document.getElementById('rightPanel').innerHTML = '';
    document.getElementById('mainArea').classList.add('hidden');
    document.getElementById('footerBar').classList.add('hidden');
    document.getElementById('uploadZone').style.display = '';
    document.getElementById('fileInput').value = '';
    // 回転リセット
    imageRotation = 0;
    const img = document.getElementById('previewImg');
    if (img) img.style.transform = '';
    // Phase 4: 送料計算の状態を完全リセット（前回の手動入力値が次の受注に紛れ込まないように）
    resetShippingPanel();
    // ドラフト編集モードフラグもリセット
    draftMode = null;
    const banner = document.getElementById('draftBanner');
    if (banner) banner.remove();
}

// 送料計算パネルの状態を完全リセット
function resetShippingPanel() {
    lastShippingResults = null;
    // shippingManualOverrides は const なので個別キー削除
    Object.keys(shippingManualOverrides).forEach(k => delete shippingManualOverrides[k]);
    const panel = document.getElementById('shippingPanel');
    if (panel) {
        panel.classList.add('hidden');
        const breakdown = document.getElementById('shippingBreakdown');
        if (breakdown) breakdown.innerHTML = '';
        const manual = document.getElementById('shippingManualInputs');
        if (manual) manual.innerHTML = '';
        const total = document.getElementById('shippingTotalDisplay');
        if (total) total.innerHTML = '';
    }
    const twoBurden = document.getElementById('shippingTwoBurden');
    if (twoBurden) twoBurden.checked = false;
}

// ─── Manual entry ───
async function startManualEntry() {
    showLoading('手入力モード準備中...');
    // 前回の状態を完全リセット
    resetShippingPanel();
    try {
        const res = await fetch('/api/manual_entry', { method: 'POST' });
        const data = await res.json();
        sessionId = data.session_id;
        const ocrRes = await fetch('/api/ocr', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: sessionId }),
        });
        const ocrData = await ocrRes.json();
        pageResults = ocrData.pages;
        currentPage = 0;
        showResults();
        hideLoading();
    } catch (err) {
        alert('エラー: ' + err.message);
        hideLoading();
    }
}
function addProductRow(pageIdx) {
    const pr = pageResults[pageIdx];
    pr.matched_items.push({
        matched: false,
        ocr_name: '',
        master_name: '',
        jan: '',
        code: '',
        quantity: 0,
        cs_price: 0,
        amount: 0,
        output_dest: '',
        case_quantity: 0,
        expiry_date: '',
        double_pack: false,
    });
    renderPage(pageIdx);
    recalculateShipping();
}
function removeProductRow(pageIdx, itemIdx) {
    pageResults[pageIdx].matched_items.splice(itemIdx, 1);
    renderPage(pageIdx);
    recalculateShipping();
}

// ─── Image rotation ───
let imageRotation = 0;
function rotateImage() {
    imageRotation = (imageRotation + 180) % 360;
    document.getElementById('previewImg').style.transform = `rotate(${imageRotation}deg)`;
}

// ─── Loading overlay ───
function showLoading(text) {
    document.getElementById('loadingText').textContent = text;
    document.getElementById('loadingOverlay').classList.remove('hidden');
}
function hideLoading() {
    document.getElementById('loadingOverlay').classList.add('hidden');
}
</script>
</body>
</html>"""


# ─── Main ───
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = os.environ.get("HOST", "127.0.0.1")
    print(f"\n  FAX受注処理システム 起動中...")
    print(f"  http://{host}:{port}\n")

    # Pre-load masters
    get_ddc_master()
    get_product_master()

    # Open browser (local only)
    if host == "127.0.0.1":
        threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    app.run(host=host, port=port, debug=False)
