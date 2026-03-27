"""Supabase API クライアント（FAXシステム用）

products + fax_product_master + shipping_rules を取得する。
Supabase接続できない場合はローカルCSVにフォールバック。
"""
import os
import json
import urllib.request
import urllib.error
import pandas as pd

# Supabase設定（環境変数から取得）
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


def load_product_master_from_supabase() -> pd.DataFrame | None:
    """Supabaseから商品マスタを取得する。

    product_master_scm テーブル（is_wholesale=true, is_discontinued=false）を取得し、
    fax_product_master からCS単価・出力先を補完する。
    JAN未登録の冷凍商品等は商品名マッチング用に含める。
    """
    # product_master_scm を取得（卸対象・販売中のみ）
    pm_rows = _paginated_get(
        "product_master_scm",
        "select=product_code,product_name,jan_code,case_quantity,temperature_zone"
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

        rows.append({
            "温度帯": pm.get("temperature_zone", ""),
            "商品コード": pm.get("product_code", ""),
            "商品名": pm.get("product_name", ""),
            "規格": fax.get("spec", ""),
            "配送荷姿": fax.get("pack", ""),
            "1袋単価": float(fax.get("unit_price") or 0),
            "CS単価": cs_price,
            "改定単価": 0,
            "出力先": fax.get("output_dest", ""),
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

    seen: set = set()
    records = []
    for r in rows:
        nohinsaki = (r.get("nohinsaki_name") or "").replace('\xa0', ' ').replace('\u3000', ' ').strip()
        if not nohinsaki or nohinsaki in seen:
            continue
        seen.add(nohinsaki)
        torihikisaki_code = r.get("torihikisaki_code", "")
        torihikisaki_name = torihikisaki_map.get(torihikisaki_code, "")
        hc = haruna_normalized.get(_norm_key(nohinsaki), {})
        records.append({
            "納品先名":   nohinsaki,
            "納品先コード": (r.get("nohinsaki_code") or ""),
            "取引先名":   torihikisaki_name,
            "取引先コード": torihikisaki_code,
            "郵便番号":   (r.get("yubin_bango") or ""),
            "住所":       (r.get("jusho") or ""),
            "電話番号":   (r.get("denwa_bango") or ""),
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
    print("  [Cache] 全キャッシュをクリアしました")
