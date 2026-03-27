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
from ocr_module import (
    process_fax_pdf, pdf_to_images, load_ddc_master, load_product_master,
    match_ddc, match_product, load_staff, ocr_fax_page, normalize, normalize_company,
    _ddc_row_to_dict,
)
from process_fax import generate_pdfs, results_to_excel, results_to_csv, results_to_ne_csv, results_to_coola_csv, ensure_output_dir, OUTPUT_DIR
from pdf_generator import gen_sylvia_pdf, gen_haruna_pdf
from datetime import date

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# In-memory session store (single-user local app)
sessions = {}

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
    return render_template_string(HTML_TEMPLATE)


@app.route("/api/ddc_list")
def api_ddc_list():
    return jsonify(get_ddc_list())


@app.route("/api/upload", methods=["POST"])
def api_upload():
    """Upload PDF, convert to page images, return session_id + page list"""
    if "file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "PDFファイルを選択してください"}), 400

    pdf_bytes = f.read()
    session_id = str(uuid.uuid4())[:8]

    # Convert to page images
    pages = pdf_to_images(pdf_bytes, dpi=200)

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
            return send_file(BytesIO(img_bytes), mimetype="image/png")
    return "Page not found", 404


@app.route("/api/ocr", methods=["POST"])
def api_ocr():
    """Run OCR + matching on uploaded PDF"""
    data = request.get_json()
    session_id = data.get("session_id")
    sess = sessions.get(session_id)
    if not sess:
        return jsonify({"error": "セッションが見つかりません"}), 404

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

        results.append({
            "page": page_info["page"],
            "ocr_raw": ocr_result,
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
                # Update quantities if changed
                if up.get("items"):
                    for ui in up["items"]:
                        idx = ui.get("index")
                        if idx is not None and idx < len(r["matched_items"]):
                            mi = r["matched_items"][idx]
                            if "quantity" in ui:
                                mi["quantity"] = ui["quantity"]
                                if mi.get("cs_price"):
                                    mi["amount"] = mi["quantity"] * mi["cs_price"]
                            if "expiry_date" in ui:
                                mi["expiry_date"] = ui["expiry_date"]
                            if "double_pack" in ui:
                                mi["double_pack"] = ui["double_pack"]
                # Update delivery date
                if up.get("delivery_date"):
                    r["ocr_raw"]["delivery_date"] = up["delivery_date"]
                # Update order_no
                if up.get("order_no"):
                    r["ocr_raw"]["order_no"] = up["order_no"]
                # Re-split sylvia/haruna
                r["sylvia_items"] = [i for i in r["matched_items"] if i.get("matched") and i.get("output_dest") == "シルビア"]
                r["haruna_items"] = [i for i in r["matched_items"] if i.get("matched") and i.get("output_dest") == "ハルナ"]

    # Generate outputs
    try:
        ensure_output_dir()
        generated = generate_pdfs(results, pdf_name, staff_name, remarks)
        csv_path = results_to_csv(results, pdf_name)
        xlsx_path = results_to_excel(results, pdf_name)
        ne_csv_path = results_to_ne_csv(results, pdf_name)
        coola_csv_path = results_to_coola_csv(results, pdf_name)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"生成エラー: {repr(e)}"}), 500

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

    return jsonify({"success": True, "files": output_files})


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

.header { background: #2F5496; color: white; padding: 12px 24px; display: flex; align-items: center; justify-content: space-between; }
.header h1 { font-size: 22px; font-weight: 600; }
.header .status { font-size: 15px; opacity: 0.8; }

.upload-zone {
    margin: 16px 24px; padding: 32px; background: white; border: 2px dashed #aaa;
    border-radius: 8px; text-align: center; cursor: pointer; transition: all 0.2s;
}
.upload-zone:hover, .upload-zone.drag-over { border-color: #2F5496; background: #f0f4ff; }
.upload-zone input { display: none; }
.upload-zone p { color: #666; font-size: 17px; }

.main-area { display: flex; margin: 0 24px 24px; gap: 16px; height: calc(100vh - 200px); }
.main-area.hidden { display: none; }

.left-panel {
    flex: 0 0 42%; background: #1a1a1a; border-radius: 8px; overflow: hidden;
    display: flex; flex-direction: column;
}
.left-panel .page-nav {
    background: #2a2a2a; padding: 8px 12px; display: flex; align-items: center;
    justify-content: center; gap: 12px; color: #ccc; font-size: 13px;
}
.left-panel .page-nav button {
    background: #444; color: white; border: none; padding: 4px 12px; border-radius: 4px; cursor: pointer;
}
.left-panel .page-nav button:hover { background: #555; }
.left-panel .page-nav button:disabled { opacity: 0.3; cursor: default; }
.left-panel .preview-img {
    flex: 1; overflow: auto; display: flex; align-items: flex-start; justify-content: center; padding: 8px;
}
.left-panel .preview-img img { max-width: 100%; height: auto; }

.right-panel { flex: 1; overflow-y: auto; display: flex; flex-direction: column; gap: 12px; }

.card {
    background: white; border-radius: 8px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
.card h3 { font-size: 17px; color: #2F5496; margin-bottom: 10px; border-bottom: 1px solid #e0e0e0; padding-bottom: 6px; }

.field-grid { display: grid; grid-template-columns: 100px 1fr; gap: 6px 12px; align-items: center; }
.field-grid label { font-size: 15px; color: #666; font-weight: 600; text-align: right; }
.field-grid input, .field-grid select {
    padding: 8px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 15px; width: 100%;
}
.field-grid input:focus { outline: none; border-color: #2F5496; box-shadow: 0 0 0 2px rgba(47,84,150,0.15); }

.badge { display: inline-block; padding: 3px 12px; border-radius: 10px; font-size: 13px; font-weight: 600; }
.badge-ok { background: #E2EFDA; color: #2d7a2d; }
.badge-review { background: #FFF2CC; color: #b8860b; }
.badge-ng { background: #FCE4D6; color: #c0392b; }

/* DDC search dropdown */
.ddc-search-wrap { position: relative; }
.ddc-search-wrap input { width: 100%; }
.ddc-dropdown {
    position: absolute; top: 100%; left: 0; right: 0; z-index: 100;
    background: white; border: 1px solid #ccc; border-top: none; border-radius: 0 0 6px 6px;
    max-height: 280px; overflow-y: auto; box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    display: none;
}
.ddc-dropdown.show { display: block; }
.ddc-item {
    padding: 8px 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0; transition: background 0.1s;
}
.ddc-item:hover, .ddc-item.active { background: #e8f0fe; }
.ddc-item .ddc-name { font-size: 15px; font-weight: 600; }
.ddc-item .ddc-addr { font-size: 13px; color: #888; margin-top: 2px; }

.items-table { width: 100%; border-collapse: collapse; font-size: 15px; }
.items-table th { background: #f5f5f5; padding: 6px 8px; text-align: left; font-weight: 600; border-bottom: 2px solid #ddd; }
.items-table td { padding: 6px 8px; border-bottom: 1px solid #eee; }
.items-table tr.matched { }
.items-table tr.unmatched { background: #FCE4D6; }
.items-table input[type="number"] { width: 60px; padding: 4px; border: 1px solid #ddd; border-radius: 3px; text-align: right; }

.footer-bar {
    margin: 0 24px 16px; padding: 12px 16px; background: white; border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1); display: flex; align-items: center; gap: 16px;
}
.footer-bar.hidden { display: none; }
.footer-bar label { font-size: 15px; color: #666; font-weight: 600; }
.footer-bar input { padding: 8px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 15px; }
.btn-confirm {
    margin-left: auto; background: #2F5496; color: white; border: none; padding: 10px 28px;
    border-radius: 6px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background 0.2s;
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
    <div class="status" id="statusText"></div>
    <button onclick="reloadMaster()" style="background:#FF8F00;color:#fff;border:none;border-radius:4px;padding:6px 14px;cursor:pointer;font-size:14px;margin-left:8px" title="Supabaseからマスタを再読込">マスタ更新</button>
</div>

<div class="upload-zone" id="uploadZone" onclick="document.getElementById('fileInput').click()">
    <input type="file" id="fileInput" accept=".pdf" multiple>
    <p>PDFファイルをドラッグ＆ドロップ、またはクリックして選択</p>
</div>

<div class="main-area hidden" id="mainArea">
    <div class="left-panel">
        <div class="page-nav">
            <button id="prevBtn" onclick="changePage(-1)" disabled>&lt; 前</button>
            <span id="pageInfo">1 / 1</span>
            <button id="nextBtn" onclick="changePage(1)" disabled>次 &gt;</button>
        </div>
        <div class="preview-img">
            <img id="previewImg" src="" alt="PDF Preview">
        </div>
    </div>
    <div class="right-panel" id="rightPanel">
        <!-- Dynamically populated per page -->
    </div>
</div>

<div class="footer-bar hidden" id="footerBar">
    <label>担当者</label>
    <input type="text" id="staffInput" value="伊藤" style="width:80px">
    <label>備考</label>
    <input type="text" id="remarksInput" value="" style="width:200px">
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
let ddcMaster = [];

// ─── Init: load DDC list ───
fetch('/api/ddc_list').then(r => r.json()).then(data => {
    ddcMaster = data;
    document.getElementById('statusText').textContent = `DDCマスタ: ${data.length}件読込完了`;
});

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
    // Process first file for now
    const file = files[0];
    showLoading(`アップロード中: ${file.name}`);

    const form = new FormData();
    form.append('file', file);

    try {
        const res = await fetch('/api/upload', { method: 'POST', body: form });
        const data = await res.json();
        if (data.error) { alert(data.error); hideLoading(); return; }

        sessionId = data.session_id;
        showLoading(`OCR処理中... (${data.page_count}ページ)`);

        // Run OCR
        const ocrRes = await fetch('/api/ocr', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: sessionId }),
        });
        const ocrData = await ocrRes.json();
        if (ocrData.error) { alert(ocrData.error); hideLoading(); return; }

        pageResults = ocrData.pages;
        currentPage = 0;
        showResults();
        hideLoading();
    } catch (err) {
        alert('エラー: ' + err.message);
        hideLoading();
    }
}

// ─── Display Results ───
function showResults() {
    document.getElementById('mainArea').classList.remove('hidden');
    document.getElementById('footerBar').classList.remove('hidden');
    document.getElementById('uploadZone').style.display = 'none';
    renderPage(currentPage);
    updatePageNav();
}

function changePage(delta) {
    currentPage += delta;
    if (currentPage < 0) currentPage = 0;
    if (currentPage >= pageResults.length) currentPage = pageResults.length - 1;
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
    // Preview image
    document.getElementById('previewImg').src = `/api/page_image/${sessionId}/${pr.page}`;

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
        if (!it.expiry_date) it.expiry_date = '';
        if (it.double_pack === undefined) it.double_pack = false;
        const cls = it.matched ? 'matched' : 'unmatched';
        const name = it.matched ? it.master_name : it.ocr_name;
        const matchBadge = it.matched ? '<span class="badge badge-ok">OK</span>' : '<span class="badge badge-ng">NG</span>';
        const dpChecked = it.double_pack ? 'checked' : '';
        const dpStyle = it.double_pack ? 'background:#E65100;color:#fff;' : 'background:#eee;color:#999;';
        itemsRows += `<tr class="${cls}">
            <td>${name}</td>
            <td>${it.jan || ''}</td>
            <td><input type="number" value="${it.quantity || 0}" min="0" data-page="${idx}" data-item="${i}" onchange="updateQty(this)" style="width:60px"></td>
            <td><input type="date" value="${it.expiry_date}" data-page="${idx}" data-item="${i}" onchange="updateExpiry(this)" style="width:140px;font-size:14px"></td>
            <td><button onclick="toggleDoublePack(${idx},${i},this)" style="border:none;border-radius:4px;padding:5px 10px;font-size:13px;cursor:pointer;${dpStyle}" id="dpBtn-${idx}-${i}">${it.double_pack ? '二重梱包' : '通常'}</button></td>
            <td>${it.output_dest || ''}</td>
            <td>${matchBadge}</td>
        </tr>`;
    }

    panel.innerHTML = `
    <div class="card">
        <h3>注文情報 <span class="badge ${badgeClass}" id="ddcBadge-${idx}">${status}</span></h3>
        <div class="field-grid">
            <label>オーダーNO</label>
            <input type="text" value="${ocr.order_no || ''}" data-page="${idx}" data-field="order_no" onchange="editField(this)">
            <label>納品日</label>
            <input type="date" value="${ocr.delivery_date || ''}" data-page="${idx}" data-field="delivery_date" onchange="editField(this)">
            <label>発注元</label>
            <input type="text" value="${ocr.sender || ''}" readonly style="background:#f5f5f5">
            <label>納品先(OCR)</label>
            <input type="text" value="${ocr.delivery_dest || ''}" readonly style="background:#f5f5f5">
        </div>
    </div>
    <div class="card">
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
            <thead><tr><th>商品名</th><th>JAN</th><th>数量(CS)</th><th>賞味期限</th><th>梱包</th><th>出力先</th><th>マッチ</th></tr></thead>
            <tbody>${itemsRows}</tbody>
        </table>
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
    if (!query || query.length < 1) { dd.classList.remove('show'); return; }

    const q = query.toLowerCase();
    const matches = ddcMaster.filter(d =>
        d.name.toLowerCase().includes(q) || d.address.toLowerCase().includes(q)
    ).slice(0, 30);

    if (matches.length === 0) {
        dd.innerHTML = '<div class="ddc-item" style="color:#999">該当なし</div>';
        dd.classList.add('show');
        return;
    }

    dd.innerHTML = matches.map(d => `
        <div class="ddc-item" onclick="selectDdc(${pageIdx}, '${d.name.replace(/'/g, "\\'")}')">
            <div class="ddc-name">${highlightMatch(d.name, q)}</div>
            <div class="ddc-addr">${d.address}</div>
        </div>
    `).join('');
    dd.classList.add('show');
}

function highlightMatch(text, query) {
    const idx = text.toLowerCase().indexOf(query);
    if (idx < 0) return text;
    return text.substring(0, idx) + '<b style="color:#2F5496">' + text.substring(idx, idx + query.length) + '</b>' + text.substring(idx + query.length);
}

function selectDdc(pageIdx, name) {
    const input = document.getElementById(`ddcSearch-${pageIdx}`);
    const dd = document.getElementById(`ddcDropdown-${pageIdx}`);
    input.value = name;
    dd.classList.remove('show');

    // Store selection in pageResults
    const entry = ddcMaster.find(d => d.name === name);
    if (entry) {
        pageResults[pageIdx].ddc_match.name = name;
        pageResults[pageIdx].ddc_match.matched = true;
        pageResults[pageIdx].ddc_match.low_confidence = false;
        pageResults[pageIdx].ddc_match.address = entry.address;
        pageResults[pageIdx].ddc_match.tel = entry.tel;
        pageResults[pageIdx]._userDdc = name;

        document.getElementById(`ddcInfo-${pageIdx}`).textContent =
            `住所: ${entry.address} / TEL: ${entry.tel}`;
        const badge = document.getElementById(`ddcBadge-${pageIdx}`);
        if (badge) { badge.className = 'badge badge-ok'; badge.textContent = 'OK'; }
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
}
function updateExpiry(el) {
    const idx = parseInt(el.dataset.page);
    const itemIdx = parseInt(el.dataset.item);
    pageResults[idx].matched_items[itemIdx].expiry_date = el.value;
}
function toggleDoublePack(pageIdx, itemIdx, btn) {
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
    btn.disabled = true;
    showLoading('発注PDF生成中...');

    const pages = pageResults.map((pr, idx) => ({
        page: pr.page,
        ddc_name: pr._userDdc || pr.ddc_match.name || '',
        order_no: pr.ocr_raw.order_no || '',
        delivery_date: pr.ocr_raw.delivery_date || '',
        items: pr.matched_items.map((it, i) => ({ index: i, quantity: it.quantity, expiry_date: it.expiry_date || '', double_pack: it.double_pack || false })),
    }));

    try {
        const res = await fetch('/api/confirm', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                session_id: sessionId,
                staff_name: document.getElementById('staffInput').value,
                remarks: document.getElementById('remarksInput').value,
                pages: pages,
            }),
        });
        const data = await res.json();
        if (data.error) { alert(data.error); hideLoading(); btn.disabled = false; return; }

        // Show output files
        const outDiv = document.getElementById('outputFiles');
        outDiv.innerHTML = data.files.map(f => {
            const icon = f.type === 'pdf' ? '📄' : f.type === 'xlsx' ? '📊' : f.type === 'ne_csv' ? '🔄' : f.type === 'coola_csv' ? '🏭' : '📝';
            const label = f.label ? `${f.label} ` : '';
            const style = (f.type === 'ne_csv' || f.type === 'coola_csv') ? ' style="background:#E65100;font-weight:bold"' : '';
            return `<a href="/api/download/${encodeURIComponent(f.name)}" target="_blank"${style}>${icon} ${label}${f.name}</a>`;
        }).join('');
        outDiv.innerHTML += '<br><button onclick="resetForNext()" style="margin-top:12px;padding:10px 28px;background:#2E7D32;color:#fff;border:none;border-radius:6px;font-size:15px;cursor:pointer">次のPDFを処理</button>';
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
            // DDCリストも再取得
            const ddcRes = await fetch('/api/ddc_list');
            ddcMaster = await ddcRes.json();
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
    currentPage = 1;
    totalPages = 0;
    ocrResults = {};
    document.getElementById('outputPanel').classList.add('hidden');
    document.getElementById('outputFiles').innerHTML = '';
    document.getElementById('ocrResult').innerHTML = '<p style="color:#999">PDFをアップロードしてください</p>';
    document.getElementById('pageNav').innerHTML = '';
    document.getElementById('pageView').innerHTML = '';
    document.getElementById('fileInput').value = '';
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
