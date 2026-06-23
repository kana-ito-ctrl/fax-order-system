"""Supabase API クライアント（FAXシステム用）

products + fax_product_master + shipping_rules を取得する。
Supabase接続できない場合はローカルCSVにフォールバック。
"""
import os
import json
import urllib.request
import urllib.error
import pandas as pd

# ローカル開発用: .env ファイルを読み込む（依存ライブラリなし）
# Render等の本番環境では環境変数が直接設定されているので .env なし＝OK
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(_env_path):
    with open(_env_path, encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# Supabase設定（環境変数のみ参照・ハードコード fallback なし）
# ローカル: .env から読み込み済み / Render: ダッシュボードの Environment Variables から
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")

# ローカルCSVフォールバックパス
_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


def _supabase_get(table: str, query_params: str = "") -> list[dict] | None:
    """Supabase REST API GETリクエスト。失敗時はNoneを返す。"""
    url = f"{SUPABASE_URL}/rest/v1/{table}?{query_params}"
    req = urllib.request.Request(url, headers={
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
        print(f"  [Supabase] {table} 取得失敗: {e}")
        return None


def _supabase_post(table: str, data: dict | list[dict], return_representation: bool = True) -> list[dict] | None:
    """Supabase REST API POSTリクエスト（INSERT）。失敗時はNoneを返す。"""
    url = f"{SUPABASE_URL}/rest/v1/{table}"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
    }
    if return_representation:
        headers["Prefer"] = "return=representation"
    payload = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode("utf-8")
            return json.loads(body) if body else []
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace") if hasattr(e, "read") else str(e)
        print(f"  [Supabase] {table} INSERT失敗 HTTP{e.code}: {err_body[:300]}")
        return None
    except (urllib.error.URLError, TimeoutError) as e:
        print(f"  [Supabase] {table} INSERT失敗: {e}")
        return None


def _supabase_patch(table: str, query_params: str, data: dict) -> list[dict] | None:
    """Supabase REST API PATCHリクエスト（UPDATE）。"""
    url = f"{SUPABASE_URL}/rest/v1/{table}?{query_params}"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }
    payload = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers=headers, method="PATCH")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode("utf-8")
            return json.loads(body) if body else []
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace") if hasattr(e, "read") else str(e)
        print(f"  [Supabase] {table} UPDATE失敗 HTTP{e.code}: {err_body[:300]}")
        return None
    except (urllib.error.URLError, TimeoutError) as e:
        print(f"  [Supabase] {table} UPDATE失敗: {e}")
        return None


def confirm_fax_order(
    order_id: str,
    *,
    confirmed_by: str,
    snapshot: dict,
    shipping_fee: int | None = None,
    shipping_fee_two_burden: bool = False,
    shipping_fee_breakdown: dict | None = None,
    edited_header: dict | None = None,
) -> dict | None:
    """Phase 4: draft → confirmed の遷移処理。

    1) 現行レコードの confirmation_count を取得
    2) status='confirmed' / confirmed_at=now / confirmed_by / confirmation_count++
       confirmation_history に snapshot を append
       shipping_fee 関連も同時に更新
    3) edited_header で編集されたフィールド（partner_name 等）も更新

    Args:
        order_id: fax_orders_scm の id (UUID)
        confirmed_by: 確定者名（staff_name）
        snapshot: 確定時点のフルデータ（後で版間比較に使う）
        shipping_fee: 送料(税抜・円)。None なら更新しない
        shipping_fee_two_burden: True ならTWO負担（CSVには 0 が入る）
        shipping_fee_breakdown: calculate_shipping_fee の戻り値そのもの
        edited_header: 編集されたヘッダー項目（delivery_date, partner_name, ...）

    Returns:
        更新後のレコード dict、失敗時 None
    """
    from datetime import datetime, timezone, timedelta
    JST = timezone(timedelta(hours=9))

    # 現行 confirmation_count + history を取得
    current = _supabase_get(
        "fax_orders_scm",
        f"id=eq.{order_id}&select=confirmation_count,confirmation_history,status"
    )
    if not current:
        print(f"  [Supabase] confirm_fax_order: order_id={order_id} not found")
        return None

    cur = current[0]
    new_count = (cur.get("confirmation_count") or 0) + 1
    history = list(cur.get("confirmation_history") or [])
    confirmed_at = datetime.now(JST).isoformat(timespec="seconds")
    history.append({
        "at": confirmed_at,
        "by": confirmed_by,
        "snapshot": snapshot,
    })

    payload = {
        "status": "confirmed",
        "confirmed_at": confirmed_at,
        "confirmed_by": confirmed_by,
        "confirmation_count": new_count,
        "confirmation_history": history,
    }
    if shipping_fee is not None:
        payload["shipping_fee"] = int(shipping_fee)
    payload["shipping_fee_two_burden"] = bool(shipping_fee_two_burden)
    if shipping_fee_breakdown is not None:
        payload["shipping_fee_breakdown"] = shipping_fee_breakdown

    # 編集されたヘッダー項目をマージ（None 値は除外）
    if edited_header:
        for k, v in edited_header.items():
            if v is not None:
                payload[k] = v

    result = _supabase_patch("fax_orders_scm", f"id=eq.{order_id}", payload)
    if result:
        print(f"  [Supabase] confirm 完了: order_id={order_id[:8]}.. count={new_count}")
        return result[0] if isinstance(result, list) else result
    print(f"  [Supabase] confirm 失敗: order_id={order_id[:8]}..")
    return None


def insert_fax_order(header: dict, items: list[dict]) -> str | None:
    """fax_orders_scm にヘッダーをINSERT、fax_order_items_scm に明細をINSERTする。

    Args:
        header: fax_orders_scm 用の辞書（source_channel, status, slip_number等）
        items:  fax_order_items_scm 用の辞書リスト（order_idは自動設定）

    Returns:
        作成したorderのUUID。失敗時はNone。
    """
    # ① ヘッダーINSERT
    inserted = _supabase_post("fax_orders_scm", header)
    if not inserted:
        return None
    order_id = inserted[0].get("id")
    if not order_id:
        return None

    # ② 明細INSERT（order_idを埋め込む）
    if items:
        items_with_fk = [{**item, "order_id": order_id} for item in items]
        result = _supabase_post("fax_order_items_scm", items_with_fk)
        if result is None:
            # 明細失敗時はヘッダーも巻き戻す
            print(f"  [Supabase] 明細INSERT失敗のためヘッダーを削除: {order_id}")
            url = f"{SUPABASE_URL}/rest/v1/fax_orders_scm?id=eq.{order_id}"
            req = urllib.request.Request(url, headers={
                "apikey": SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
            }, method="DELETE")
            try:
                urllib.request.urlopen(req, timeout=10)
            except Exception:
                pass
            return None

    return order_id


def check_fax_order_exists(source_email_id: str, ocr_page_no: int) -> bool:
    """既に同じメール×ページの受注がfax_orders_scmに存在するかチェックする。"""
    query = f"source_email_id=eq.{source_email_id}&ocr_page_no=eq.{ocr_page_no}&select=id"
    rows = _supabase_get("fax_orders_scm", query)
    return rows is not None and len(rows) > 0


def fetch_pending_orders(limit: int = 200, source_channel: str | None = None) -> list[dict]:
    """確認待ち受注（status='draft' or 'error'）を新しい順に取得する。

    fax_orders_scm のみ対象（インフォマートは別画面で確認するため除外）。

    Args:
        limit: 取得上限件数
        source_channel: 'efax' / 'paltac' / None（すべて）

    Returns:
        ヘッダーレコードのリスト（明細は含まない、確認画面のサマリ用）
    """
    cols = (
        "id,source_channel,status,slip_number,order_date,delivery_date,"
        "partner_name,delivery_location_name,grand_total,ddc_matched,"
        "warehouse,shipping_type,status_flags,source_file_name,"
        "source_email_from,source_email_received_at,ocr_page_no,created_at"
    )
    query_parts = [
        f"select={cols}",
        "status=in.(draft,error)",
        f"order=created_at.desc",
        f"limit={limit}",
    ]
    if source_channel:
        query_parts.append(f"source_channel=eq.{source_channel}")
    rows = _supabase_get("fax_orders_scm", "&".join(query_parts))
    return rows if rows is not None else []


def fetch_confirmed_orders(limit: int = 200, source_channel: str | None = None,
                           days: int | None = 90) -> list[dict]:
    """確定済受注（status='confirmed'）を確定日時の新しい順に取得する。

    Phase 4 `/confirmed` ページ用。

    Args:
        limit: 取得上限件数
        source_channel: 'efax' / 'paltac' / None（すべて）
        days: 直近N日以内に絞る（None で無制限）。デフォルト90日

    Returns:
        ヘッダーレコードのリスト（confirmation_count, confirmed_at, shipping_fee 含む）
    """
    cols = (
        "id,source_channel,status,slip_number,order_date,delivery_date,"
        "partner_name,delivery_location_name,grand_total,ddc_matched,"
        "warehouse,shipping_type,status_flags,source_file_name,"
        "source_email_from,source_email_received_at,ocr_page_no,created_at,"
        "confirmed_at,confirmed_by,confirmation_count,"
        "shipping_fee,shipping_fee_two_burden"
    )
    query_parts = [
        f"select={cols}",
        "status=eq.confirmed",
        "order=confirmed_at.desc",
        f"limit={limit}",
    ]
    if source_channel:
        query_parts.append(f"source_channel=eq.{source_channel}")
    if days is not None:
        from datetime import datetime, timedelta, timezone
        # +00:00 を含めると URL でスペース化するため "Z" リテラルで UTC 表現
        cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
        query_parts.append(f"confirmed_at=gte.{cutoff}")
    rows = _supabase_get("fax_orders_scm", "&".join(query_parts))
    return rows if rows is not None else []


def fetch_order_with_items(order_id: str) -> dict | None:
    """ヘッダー1件と紐づく明細を取得する（確認画面の詳細表示用）"""
    header_rows = _supabase_get("fax_orders_scm", f"id=eq.{order_id}&select=*")
    if not header_rows:
        return None
    items = _supabase_get(
        "fax_order_items_scm",
        f"order_id=eq.{order_id}&select=*&order=line_no.asc"
    )
    return {
        "header": header_rows[0],
        "items": items if items is not None else [],
    }


def load_product_master_from_supabase() -> pd.DataFrame | None:
    """Supabaseから商品マスタを取得する。

    product_master_scm テーブル（is_wholesale=true, is_discontinued=false）を取得し、
    fax_product_master からCS単価・出力先を補完する。
    JAN未登録の冷凍商品等は商品名マッチング用に含める。
    """
    # product_master_scm を取得（卸対象・販売中のみ）
    pm_rows = _paginated_get(
        "product_master_scm",
        "select=product_code,product_name,jan_code,case_quantity,temperature_zone,notes"
        "&is_wholesale=eq.true&is_discontinued=eq.false"
    )
    if not pm_rows:
        return None

    # fax_product_master からCS単価・出力先を取得（シルビアPDF等で使用）
    fax_pm = _supabase_get("fax_product_master", "select=jan,cs_price,unit_price,spec,pack,output_dest,price_display")
    fax_map = {}
    if fax_pm:
        for fp in fax_pm:
            jan = fp.get("jan", "")
            if jan:
                fax_map[jan] = fp

    rows = []
    for pm in pm_rows:
        # ケース単位の商品（-cs, -case サフィックス）はOCRマッチ対象外
        pcode = pm.get("product_code", "")
        if pcode.endswith("-cs") or pcode.endswith("-case"):
            continue
        jan = pm.get("jan_code") or ""
        fax = fax_map.get(jan, {})
        cs_price = float(fax.get("cs_price") or 0)
        case_qty = pm.get("case_quantity") or 0
        try:
            case_qty = int(case_qty)
        except (ValueError, TypeError):
            case_qty = 0

        # fax_product_masterにない商品はnotesからメタ情報を取得
        notes = pm.get("notes") or ""
        def _notes_val(key):
            if key + ":" in notes:
                return notes.split(key + ":")[1].split()[0].strip()
            return ""

        output_dest = fax.get("output_dest", "") or _notes_val("output_dest")
        spec = fax.get("spec", "") or _notes_val("spec")
        pack = fax.get("pack", "") or _notes_val("pack")
        unit_price = float(fax.get("unit_price") or 0) or float(_notes_val("unit_price") or 0)
        if not cs_price:
            cs_price = float(_notes_val("cs_price") or 0)

        rows.append({
            "温度帯": pm.get("temperature_zone", ""),
            "商品コード": pm.get("product_code", ""),
            "商品名": pm.get("product_name", ""),
            "規格": spec,
            "配送荷姿": pack,
            "1袋単価": unit_price,
            "CS単価": cs_price,
            "改定単価": 0,
            "出力先": output_dest,
            "JANコード": jan,
            "入数": case_qty,
            "price_display": fax.get("price_display", "cs"),
        })

    if not rows:
        return None

    df = pd.DataFrame(rows)
    df["JANコード"] = df["JANコード"].astype(str)
    return df


def load_product_master() -> pd.DataFrame:
    """商品マスタを読み込む（Supabase優先、フォールバックはCSV）"""
    df = load_product_master_from_supabase()
    if df is not None and len(df) > 0:
        jan_count = (df["JANコード"].str.len() > 0).sum()
        print(f"  [Supabase] 商品マスタ(product_master_scm): {len(df)}件読み込み（JAN登録: {jan_count}件）")
        return df

    # フォールバック: ローカルCSV
    csv_path = os.path.join(_DATA_DIR, "product_master.csv")
    print(f"  [CSV] 商品マスタ読み込み: {csv_path}")
    return pd.read_csv(csv_path, dtype={"JANコード": str})


def load_shipping_rules_from_supabase() -> list[dict] | None:
    """Supabaseから出荷振分けルールを取得する"""
    return _supabase_get("shipping_rules", "select=*&order=priority.asc")


def load_shipping_expiry_rules_from_supabase() -> dict | None:
    """Supabaseから出荷賞味期限ルールを取得する。
    Returns: {jan: min_remaining_days} の dict、失敗時は None
    """
    data = _supabase_get("shipping_expiry_rules", "select=jan,min_remaining_days")
    if data is None:
        return None
    return {r["jan"]: r["min_remaining_days"] for r in data}


def load_haruna_conditions_from_supabase() -> dict | None:
    """Supabaseからハルナ出荷指示書用・納品先別条件を取得する。
    Returns: {nohinsaki: {arrival_time, basse_reservation, pallet_condition, jpr_code}} の dict
    """
    data = _supabase_get(
        "haruna_conditions",
        "select=nohinsaki,arrival_time,basse_reservation,pallet_condition,jpr_code,notes"
    )
    if data is None:
        return None
    return {r["nohinsaki"]: r for r in data}


_torihikisaki_cache: pd.DataFrame | None = None


def _paginated_get(table: str, query_params: str) -> list[dict] | None:
    """Supabaseデフォルト上限1000件を超えるテーブルをページネーションで全件取得する。"""
    all_rows = []
    page_size = 1000
    offset = 0
    separator = "&" if query_params else ""
    while True:
        batch = _supabase_get(
            table,
            f"{query_params}{separator}limit={page_size}&offset={offset}"
        )
        if batch is None:
            if all_rows:
                break  # 途中で失敗しても取得済み分で続行
            return None
        all_rows.extend(batch)
        if len(batch) < page_size:
            break
        offset += page_size
    return all_rows


def load_torihikisaki_master_from_supabase() -> pd.DataFrame | None:
    """nohinsaki_master_scm + torihikisaki_master_scm + haruna_conditions を結合して
    DDCマスタ相当のDataFrameを返す。

    nohinsaki_master_scm の nohinsaki_name を納品先名として使い、
    torihikisaki_master_scm から取引先名を取得、
    haruna_conditions からバース予約・パレット条件等をマージして返す。
    is_active=true のレコードのみ対象。
    """
    global _torihikisaki_cache
    if _torihikisaki_cache is not None:
        return _torihikisaki_cache

    # nohinsaki_master_scm（納品先マスタ）を全件取得
    rows = _paginated_get(
        "nohinsaki_master_scm",
        "select=nohinsaki_code,nohinsaki_name,torihikisaki_code,yubin_bango,jusho,denwa_bango&is_active=eq.true"
    )
    if not rows:
        return None

    # torihikisaki_master_scm（取引先マスタ）をコード→名前のマップに
    torihikisaki_rows = _paginated_get(
        "torihikisaki_master_scm",
        "select=torihikisaki_code,torihikisaki_name&is_active=eq.true"
    )
    torihikisaki_map = {}
    if torihikisaki_rows:
        for tr in torihikisaki_rows:
            code = tr.get("torihikisaki_code", "")
            if code:
                torihikisaki_map[code] = tr.get("torihikisaki_name", "")

    haruna = load_haruna_conditions_from_supabase() or {}

    # haruna_conditions のキーを正規化してマッチできるようにする
    # 「㈱」vs「株式会社」等の差異を吸収
    import re
    import unicodedata
    def _norm_key(s):
        t = unicodedata.normalize('NFKC', s).replace('\xa0', ' ').replace('\u3000', ' ')
        t = re.sub(r'株式会社|（株）|\(株\)|㈱', '', t)
        t = re.sub(r'\s+', ' ', t).strip()
        return t
    haruna_normalized = {}
    for key, val in haruna.items():
        haruna_normalized[_norm_key(key)] = val

    # 電話番号サニタイズ用: Slack の <tel:NUMBER|NUMBER> 形式を素の番号に戻す
    # マスタ登録時にユーザーが Slack のリッチリンクごとコピペしてしまうケースの保険
    _TEL_MARKUP_RE = re.compile(r"<tel:([^|>]+)(?:\|[^>]*)?>")
    def _sanitize_tel(s):
        if not s:
            return ""
        s = str(s).strip()
        m = _TEL_MARKUP_RE.search(s)
        if m:
            return m.group(1).strip()
        return s

    records = []
    for r in rows:
        nohinsaki = (r.get("nohinsaki_name") or "").replace('\xa0', ' ').replace('\u3000', ' ').strip()
        nohinsaki_code = (r.get("nohinsaki_code") or "").strip()
        if not nohinsaki:
            continue
        torihikisaki_code = r.get("torihikisaki_code", "")
        torihikisaki_name = torihikisaki_map.get(torihikisaki_code, "")
        hc = haruna_normalized.get(_norm_key(nohinsaki), {})
        # 表示名: 「納品先名 [コード] 住所」で区別しやすく
        jusho = (r.get("jusho") or "").strip()
        display_name = f"{nohinsaki} [{nohinsaki_code}]" if nohinsaki_code else nohinsaki
        if jusho:
            display_name += f" / {jusho[:30]}"
        records.append({
            "納品先名":   display_name,
            "納品先コード": nohinsaki_code,
            "取引先名":   torihikisaki_name,
            "取引先コード": torihikisaki_code,
            "郵便番号":   (r.get("yubin_bango") or ""),
            "住所":       (r.get("jusho") or ""),
            "電話番号":   _sanitize_tel(r.get("denwa_bango")),
            "FAX番号":    "",
            "入荷時間":   (hc.get("arrival_time") or ""),
            "バース予約": (hc.get("basse_reservation") or ""),
            "パレット条件": (hc.get("pallet_condition") or ""),
            "JPRコード":  (hc.get("jpr_code") or ""),
            "納品方法":   "",
        })

    if not records:
        return None

    df = pd.DataFrame(records)
    print(f"  [Supabase] 納品先マスタ(nohinsaki_master_scm): {len(df)}件読み込み（重複除去後）")
    _torihikisaki_cache = df
    return df


# キャッシュ
_haruna_conditions_cache: dict | None = None

def load_haruna_conditions() -> dict:
    """ハルナ条件をキャッシュ付きで返す。失敗時は空dict。"""
    global _haruna_conditions_cache
    if _haruna_conditions_cache is not None:
        return _haruna_conditions_cache
    result = load_haruna_conditions_from_supabase()
    if result is not None:
        print(f"  [Supabase] haruna_conditions: {len(result)}件読み込み")
        _haruna_conditions_cache = result
    else:
        print("  [Supabase] haruna_conditions 取得失敗: 空dictで継続")
        _haruna_conditions_cache = {}
    return _haruna_conditions_cache


def clear_all_caches():
    """全キャッシュをクリアする。次回アクセス時にSupabaseから再取得される。"""
    global _torihikisaki_cache, _haruna_conditions_cache
    _torihikisaki_cache = None
    _haruna_conditions_cache = None
    # shipping_rules cache も巻き戻す（warehouse 名変更等の即時反映用）
    try:
        import shipping_judge
        shipping_judge._shipping_rules_cache = None
    except Exception:
        pass
    # products / fax_product_master キャッシュも（shipping_fee_calculator）
    try:
        import shipping_fee_calculator
        shipping_fee_calculator._products_cache = None
    except Exception:
        pass
    print("  [Cache] 全キャッシュをクリアしました")
