"""出荷判定・ロットアラート・リードタイムアラートロジック

出荷振分けルール：
  Supabase shipping_rules テーブルから読み込み（フォールバック: ハードコード）
  - priority順に評価し、最初にマッチしたルールを適用
  - jan / ddc_code が NULL のルールは全商品/全納品先に適用

リードタイムルール（受注日=1日目、土日祝除く）：
- シルビア直送:        4日目（+3営業日）より早い納品日 → アラート
- ハルナ直送(〜2026/3): 5日目（+4営業日）より早い納品日 → アラート
- ハルナ直送(2026/4〜): 4日目（+3営業日）より早い納品日 → アラート
"""

from datetime import date, datetime, timedelta

try:
    import jpholiday
    HAS_JPHOLIDAY = True
except ImportError:
    HAS_JPHOLIDAY = False

DIRECT_SHIPPING_THRESHOLD = 10  # 直送条件: 10CS以上（フォールバック用）

# Supabaseルールのキャッシュ（セッション中1回のみ取得）
_shipping_rules_cache = None


def _load_shipping_rules():
    """Supabaseから出荷ルールを読み込む（フォールバック: None）"""
    global _shipping_rules_cache
    if _shipping_rules_cache is not None:
        return _shipping_rules_cache
    try:
        from supabase_client import load_shipping_rules_from_supabase
        rules = load_shipping_rules_from_supabase()
        if rules:
            _shipping_rules_cache = rules
            return rules
    except Exception:
        pass
    _shipping_rules_cache = []
    return []


def judge_shipping(output_dest: str, total_cs: int, jan: str = None, ddc_code: str = None) -> str:
    """出荷区分を判定する

    Args:
        output_dest: 出力先（"ハルナ" / "シルビア" / "自社倉庫"）
        total_cs: 合計CS数
        jan: JANコード（Supabaseルール用、省略可）
        ddc_code: 納品先コード（Supabaseルール用、省略可）

    Returns:
        "直送" または "自社倉庫"
    """
    # Supabaseルールがあれば使用
    rules = _load_shipping_rules()
    if rules:
        for rule in rules:
            # JANフィルタ
            if rule.get("jan") and rule["jan"] != jan:
                continue
            # DDCフィルタ
            if rule.get("ddc_code") and rule["ddc_code"] != ddc_code:
                continue
            # CS数範囲チェック
            min_cs = rule.get("min_cs", 0)
            max_cs = rule.get("max_cs")
            if total_cs < min_cs:
                continue
            if max_cs is not None and total_cs >= max_cs:
                continue
            # マッチ: warehouse が直送元名（シルビア/ハルナ）なら直送
            warehouse = rule.get("warehouse", "")
            if warehouse in ("シルビア", "ハルナ"):
                return "直送"
            return "自社倉庫"

    # フォールバック: ハードコードルール
    if output_dest in ("ハルナ", "シルビア"):
        return "直送" if total_cs >= DIRECT_SHIPPING_THRESHOLD else "自社倉庫"
    return "自社倉庫"


def judge_warehouse(output_dest: str, total_cs: int, jan: str = None, ddc_code: str = None) -> str:
    """具体的な出荷倉庫名を返す（直送の場合はシルビア/ハルナ、自社倉庫の場合はSBS/ベルーナ等）

    Args:
        output_dest: 出力先
        total_cs: 合計CS数
        jan: JANコード
        ddc_code: 納品先コード

    Returns:
        倉庫名 (例: "シルビア", "ハルナ", "SBS", "ベルーナ")
    """
    rules = _load_shipping_rules()
    if rules:
        for rule in rules:
            if rule.get("jan") and rule["jan"] != jan:
                continue
            if rule.get("ddc_code") and rule["ddc_code"] != ddc_code:
                continue
            min_cs = rule.get("min_cs", 0)
            max_cs = rule.get("max_cs")
            if total_cs < min_cs:
                continue
            if max_cs is not None and total_cs >= max_cs:
                continue
            return rule.get("warehouse", "自社倉庫")

    # フォールバック
    if output_dest in ("ハルナ", "シルビア"):
        if total_cs >= DIRECT_SHIPPING_THRESHOLD:
            return output_dest
        return "自社倉庫"
    return "自社倉庫"


def check_lot_alert(output_dest: str, total_cs: int):
    """ロット数アラートを確認する

    Args:
        output_dest: 出力先（"ハルナ" または "シルビア"）
        total_cs: 合計CS数

    Returns:
        (is_alert: bool, alert_message: str)
    """
    if output_dest in ("ハルナ", "シルビア"):
        if total_cs < DIRECT_SHIPPING_THRESHOLD:
            msg = (
                f"ロット未満アラート: {output_dest} {total_cs}CS"
                f"（直送条件: {DIRECT_SHIPPING_THRESHOLD}CS以上）"
            )
            return True, msg
    return False, ""


def _is_business_day(d: date) -> bool:
    """営業日かどうか判定（土日・祝日除く）"""
    if d.weekday() >= 5:  # 土曜=5, 日曜=6
        return False
    if HAS_JPHOLIDAY and jpholiday.is_holiday(d):
        return False
    return True


def _add_business_days(start: date, n: int) -> date:
    """startの翌営業日からn営業日後の日付を返す（start自身は含まない）"""
    current = start
    count = 0
    while count < n:
        current += timedelta(days=1)
        if _is_business_day(current):
            count += 1
    return current


def _parse_date(date_str: str):
    """日付文字列をdateオブジェクトに変換（複数フォーマット対応）"""
    if not date_str:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(date_str).strip(), fmt).date()
        except (ValueError, AttributeError):
            continue
    return None


def check_lead_time_alert(order_date_str: str, delivery_date_str: str, output_dest: str, shipping: str = "直送"):
    """リードタイムアラートを確認する

    受注日を1日目として（土日祝除く）:
    - シルビア直送:         4日目より早い → アラート（+3営業日）
    - ハルナ直送(〜2026/3): 5日目より早い → アラート（+4営業日）
    - ハルナ直送(2026/4〜): 4日目より早い → アラート（+3営業日）
    - 自社倉庫（全商品）:   3日目より早い → アラート（+2営業日）

    Returns:
        (is_alert: bool, alert_message: str)
    """
    order_dt = _parse_date(order_date_str)
    delivery_dt = _parse_date(delivery_date_str)

    if not order_dt or not delivery_dt:
        return False, ""

    if shipping == "自社倉庫":
        required_biz = 2  # 3日目
        day_label = "3"
    elif output_dest == "シルビア":
        required_biz = 3  # 4日目
        day_label = "4"
    elif output_dest == "ハルナ":
        if order_dt < date(2026, 4, 1):
            required_biz = 4  # 5日目（〜2026/3）
            day_label = "5"
        else:
            required_biz = 3  # 4日目（2026/4〜）
            day_label = "4"
    else:
        return False, ""

    earliest = _add_business_days(order_dt, required_biz)

    if delivery_dt < earliest:
        msg = (
            f"リードタイムアラート: 納品日 {delivery_dt} が"
            f"最短納品日 {earliest}（受注日から{day_label}日目）より早い"
        )
        return True, msg

    return False, ""
