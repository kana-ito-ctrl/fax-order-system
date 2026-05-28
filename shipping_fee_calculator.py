"""ロット割れ自社倉庫出荷の送料自動計算モジュール

商品マスタ (Supabase products テーブル) の brand と order_lot に基づき、
ブランド別の送料を算出する。

計算ロジック概要:
  - 2Snack（混載10cs以上）の <10CS    → lot-break-calculator (運賃+保管費+ピッキング費等の合計)
  - 2Snack（2cs単位、35gショコラ系）   → 一旦除外（将来対応）
  - 2Energy <2CS の場合 1CS=¥1,000   → 全国一律
  - 2Energy 2CS以上                  → ¥0（送料込み）
  - 2Gummy <2CS                      → 手動入力（外部Streamlit URL案内）
  - 2Water <10CS（自社倉庫出荷時のみ） → 手動入力（外部Streamlit URL案内）

Reference:
  G:/共有ドライブ/TWO/SCM/28_ロット割れ対応/.claude/skills/lot-break-calculator/SKILL.md
"""
import os
import re
import sys
import csv
from datetime import datetime
from typing import Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Windows ターミナル UTF-8 補正
if sys.stdout.encoding and sys.stdout.encoding.lower() in ('cp932', 'shift_jis', 'mbcs'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


# ========== 定数 ==========

STREAMLIT_URL = "https://logistics-cost-piu8hweufvy5pfm6mfvmcv.streamlit.app/"

LOT_BREAK_HISTORY_CSV = r"G:\共有ドライブ\TWO\SCM\28_ロット割れ対応\lot_break_history.csv"

# 都道府県 → ゾーン
_ZONE_MAP = {
    "北海道": "北海道",
    "青森県": "北東北", "秋田県": "北東北", "岩手県": "北東北",
    "宮城県": "関東", "山形県": "関東", "福島県": "関東",
    "茨城県": "関東", "栃木県": "関東", "群馬県": "関東",
    "埼玉県": "関東", "千葉県": "関東", "東京都": "関東",
    "神奈川県": "関東", "山梨県": "関東",
    "新潟県": "信越/東海/北陸", "長野県": "信越/東海/北陸",
    "富山県": "信越/東海/北陸", "石川県": "信越/東海/北陸", "福井県": "信越/東海/北陸",
    "静岡県": "信越/東海/北陸", "愛知県": "信越/東海/北陸",
    "岐阜県": "信越/東海/北陸", "三重県": "信越/東海/北陸",
    "京都府": "関西", "大阪府": "関西", "兵庫県": "関西",
    "滋賀県": "関西", "奈良県": "関西", "和歌山県": "関西",
    "鳥取県": "中国/四国", "島根県": "中国/四国", "岡山県": "中国/四国",
    "広島県": "中国/四国", "山口県": "中国/四国",
    "徳島県": "中国/四国", "香川県": "中国/四国",
    "愛媛県": "中国/四国", "高知県": "中国/四国",
    "福岡県": "九州", "佐賀県": "九州", "長崎県": "九州",
    "熊本県": "九州", "大分県": "九州", "宮崎県": "九州", "鹿児島県": "九州",
    "沖縄県": "沖縄",
}

# 1-4ケース 佐川60サイズ運賃（円/個口・税抜）
_SAGAWA_60_RATES = {
    "北海道": 530, "北東北": 500, "関東": 500, "信越/東海/北陸": 500,
    "関西": 500, "中国/四国": 500, "九州": 530, "沖縄": 1600,
}

# 5-8ケース 二重梱包BtoB最安運賃（円/個口・税抜）
_LOT_BREAK_5_8_RATES = {
    "北海道": (1350, "ヤマト120サイズ"),
    "北東北": (870, "福山通運"),
    "関東":   (670, "福山通運"),
    "信越/東海/北陸": (700, "福山通運"),
    "関西":   (740, "福山通運"),
    "中国/四国": (980, "福山通運"),
    "九州":   (1350, "ヤマト120サイズ"),
    "沖縄":   (3700, "ヤマト120サイズ"),
}

# 単価マスタ
_STORAGE_FEE_PER_CASE = 88     # 保管費
_PICKING_FEE_PER_CASE = 60     # ピッキング費
_PACKAGING_FEE_PER_PKG = 100   # 梱包手数料
_RECEIVING_FEE_PER_CASE = 60   # 入庫費
_DOUBLE_PACK_MATERIAL = 120    # 二重梱包用資材（5cs以上）
_DELIVERY_PACK_PER_PKG = 20    # デリバリーパック


# ========== 商品マスタ取得 ==========

_products_cache: list[dict] | None = None


def load_products_master() -> list[dict]:
    """Supabase products テーブルから全商品を取得（キャッシュ付き）"""
    global _products_cache
    if _products_cache is not None:
        return _products_cache

    try:
        from supabase_client import _supabase_get
        rows = _supabase_get("products", "select=product_code,jan,brand,product_name,case_qty,order_lot")
        _products_cache = rows or []
    except Exception as e:
        print(f"[shipping_fee] products取得失敗: {e}")
        _products_cache = []
    return _products_cache


def lookup_product_by_jan(jan: str) -> dict | None:
    """JANコードから商品マスタ情報を取得"""
    if not jan:
        return None
    products = load_products_master()
    for p in products:
        if p.get("jan") == jan:
            return p
    return None


# ========== 都道府県抽出 ==========

_PREF_NAMES = list(_ZONE_MAP.keys())
_PREF_PATTERN = re.compile("|".join(re.escape(p) for p in _PREF_NAMES))


def extract_prefecture(address: str) -> Optional[str]:
    """住所文字列から都道府県名を抽出"""
    if not address:
        return None
    m = _PREF_PATTERN.search(address)
    return m.group(0) if m else None


def get_zone(prefecture: str) -> Optional[str]:
    """都道府県からゾーンを判定"""
    return _ZONE_MAP.get(prefecture)


# ========== 2Snack ロット割れ計算 (lot-break-calculator 移植) ==========

def calculate_snack_lot_break(case_count: int, prefecture: str) -> dict:
    """2Snack のロット割れ送料を計算する。

    Args:
        case_count: 1〜9 (10以上はロット成立で送料なし)
        prefecture: 都道府県名（例: "東京都"）

    Returns:
        {
            "applicable": bool,         # 計算対象（1-9CS範囲内）か
            "case_count": int,
            "prefecture": str,
            "zone": str,
            "carrier": str,             # 配送会社
            "package_count": int,        # 個口数
            "carrier_fee": int,          # 運賃
            "storage_fee": int,          # 保管費
            "picking_fee": int,          # ピッキング費
            "packaging_fee": int,        # 梱包手数料
            "receiving_fee": int,        # 入庫費
            "double_pack_material": int, # 二重梱包資材
            "delivery_pack": int,        # デリバリーパック
            "subtotal_excl_tax": int,    # 小計（税抜）← CSV出力値
            "tax": int,
            "total_incl_tax": int,
            "error": str | None,
        }
    """
    result = {
        "applicable": False, "case_count": case_count, "prefecture": prefecture,
        "zone": None, "carrier": None, "package_count": 0,
        "carrier_fee": 0, "storage_fee": 0, "picking_fee": 0,
        "packaging_fee": 0, "receiving_fee": 0,
        "double_pack_material": 0, "delivery_pack": 0,
        "subtotal_excl_tax": 0, "tax": 0, "total_incl_tax": 0,
        "error": None,
    }

    if not (1 <= case_count <= 9):
        result["error"] = f"ケース数 {case_count} はロット割れ計算対象外（1〜9CSのみ）"
        return result

    zone = get_zone(prefecture)
    if not zone:
        result["error"] = f"都道府県 '{prefecture}' のゾーン判定不可"
        return result

    result["zone"] = zone
    result["applicable"] = True

    # 梱包ルール
    if 1 <= case_count <= 4:
        # 佐川60サイズ × ケース数個口
        package_count = case_count
        carrier = "佐川急便60サイズ"
        carrier_fee = _SAGAWA_60_RATES[zone] * case_count
        double_pack_material = 0
    elif 5 <= case_count <= 8:
        # 5-8: 二重梱包 1個口
        package_count = 1
        carrier_fee, carrier = _LOT_BREAK_5_8_RATES[zone]
        double_pack_material = _DOUBLE_PACK_MATERIAL
    else:  # 9
        # 9: 8cs（二重梱包1個口）+ 1cs（佐川60サイズ1個口）
        package_count = 2
        rate_5_8, carrier_5_8 = _LOT_BREAK_5_8_RATES[zone]
        rate_60 = _SAGAWA_60_RATES[zone]
        carrier_fee = rate_5_8 + rate_60
        carrier = f"{carrier_5_8} + 佐川急便60サイズ"
        double_pack_material = _DOUBLE_PACK_MATERIAL

    result["carrier"] = carrier
    result["package_count"] = package_count
    result["carrier_fee"] = carrier_fee
    result["storage_fee"] = _STORAGE_FEE_PER_CASE * case_count
    result["picking_fee"] = _PICKING_FEE_PER_CASE * case_count
    result["packaging_fee"] = _PACKAGING_FEE_PER_PKG * package_count
    result["receiving_fee"] = _RECEIVING_FEE_PER_CASE * case_count
    result["double_pack_material"] = double_pack_material
    result["delivery_pack"] = _DELIVERY_PACK_PER_PKG * package_count

    subtotal = (carrier_fee + result["storage_fee"] + result["picking_fee"] +
                result["packaging_fee"] + result["receiving_fee"] +
                double_pack_material + result["delivery_pack"])
    result["subtotal_excl_tax"] = subtotal
    result["tax"] = int(subtotal * 0.1)  # SKILL.md仕様: 切り捨て
    result["total_incl_tax"] = subtotal + result["tax"]

    return result


# ========== 全体計算 ==========

def _classify_brand(item: dict) -> tuple[str, str | None]:
    """itemから (brand, order_lot) を判定する。

    item は OCR の matched_items のフォーマット ({jan, quantity, ...}) を想定。
    """
    jan = item.get("jan") or ""
    p = lookup_product_by_jan(jan)
    if not p:
        return ("不明", None)
    return (p.get("brand") or "不明", p.get("order_lot"))


def calculate_shipping_fee(items: list[dict], ddc_address: str = "",
                           shipping_type: Optional[str] = None) -> dict:
    """受注全体の送料を算出する。

    Args:
        items: matched_items (OCR後の確定済み商品リスト)。
               各要素は {jan, quantity, master_name, ...} を持つ。
        ddc_address: 納品先住所（都道府県抽出に使用）
        shipping_type: 出荷区分 ("直送" / "自社倉庫" / None)。
                       2Water 自社倉庫出荷判定に使用。

    Returns:
        {
            "total_fee": int,         # 全brand合計の送料（税抜・円）
            "currency": "円(税抜)",
            "prefecture": str | None,
            "zone": str | None,
            "groups": [               # brand別の内訳
                {
                    "brand": str,
                    "items": [...],
                    "total_cs": int,
                    "threshold": int | None,
                    "is_lot_break": bool,
                    "fee": int | None,            # 自動算出値、Noneは手動必要
                    "method": str,                # "snack_lot_break" | "energy_flat" | "gummy_manual" | "water_manual" | "snack_skipped" | "no_fee"
                    "needs_manual": bool,
                    "external_tool_url": str | None,
                    "detail": dict,               # 計算内訳
                    "label": str,                 # UI表示用ラベル
                },
                ...
            ],
            "needs_manual_input_overall": bool,
            "warnings": [str, ...],
        }
    """
    prefecture = extract_prefecture(ddc_address)
    zone = get_zone(prefecture) if prefecture else None

    # 商品をbrand毎にグルーピング
    groups_by_brand: dict[str, dict] = {}
    for item in items:
        if not item.get("matched"):
            continue
        try:
            qty = int(float(item.get("quantity") or 0))
        except (ValueError, TypeError):
            qty = 0
        if qty <= 0:
            continue
        brand, order_lot = _classify_brand(item)
        g = groups_by_brand.setdefault(brand, {"items": [], "total_cs": 0, "order_lot": order_lot})
        g["items"].append({
            "jan": item.get("jan"),
            "name": item.get("master_name") or item.get("ocr_name"),
            "qty": qty,
            "order_lot": order_lot,
        })
        g["total_cs"] += qty
        # order_lot は同brand内で揃っていることが多いが、混在時は先勝ち
        if g["order_lot"] is None and order_lot:
            g["order_lot"] = order_lot

    groups_result = []
    total_fee = 0
    warnings = []
    needs_manual_overall = False

    for brand, g in groups_by_brand.items():
        cs = g["total_cs"]
        order_lot = g["order_lot"]
        group_data = {
            "brand": brand,
            "items": g["items"],
            "total_cs": cs,
            "threshold": None,
            "is_lot_break": False,
            "fee": 0,
            "method": "no_fee",
            "needs_manual": False,
            "external_tool_url": None,
            "detail": {},
            "label": "",
        }

        if brand == "2Snack":
            if order_lot == "混載10cs以上":
                group_data["threshold"] = 10
                if cs < 10:
                    group_data["is_lot_break"] = True
                    if not prefecture:
                        group_data["method"] = "snack_lot_break"
                    detail = calculate_snack_lot_break(cs, prefecture or "")
                    group_data["detail"] = detail
                    if detail.get("applicable"):
                        group_data["fee"] = detail["subtotal_excl_tax"]
                        group_data["method"] = "snack_lot_break"
                        group_data["label"] = f"2Snack {cs}CS / {detail['zone']} ロット割れ"
                        total_fee += group_data["fee"]
                    else:
                        group_data["method"] = "snack_lot_break_error"
                        group_data["label"] = f"2Snack {cs}CS（計算エラー: {detail.get('error')}）"
                        warnings.append(group_data["label"])
                else:
                    group_data["label"] = f"2Snack {cs}CS（10CS以上、送料なし）"
            elif order_lot == "2cs単位":
                # ショコラおかき/ラスク 35g等。一旦除外。
                group_data["threshold"] = 2
                group_data["method"] = "snack_skipped"
                if cs < 2:
                    group_data["is_lot_break"] = True
                    group_data["label"] = f"2Snack(2cs単位) {cs}CS（現在計算対象外、将来対応）"
                    warnings.append(group_data["label"])
                else:
                    group_data["label"] = f"2Snack(2cs単位) {cs}CS（2CS以上、送料なし）"
            else:
                group_data["label"] = f"2Snack {cs}CS（order_lot不明: {order_lot}）"
                warnings.append(group_data["label"])

        elif brand == "2Energy":
            group_data["threshold"] = 2
            if cs == 1:
                group_data["is_lot_break"] = True
                group_data["fee"] = 1000
                group_data["method"] = "energy_flat"
                group_data["detail"] = {"flat_rate_excl_tax": 1000, "note": "1CS全国一律"}
                group_data["label"] = f"2Energy 1CS（全国一律¥1,000）"
                total_fee += 1000
            else:
                group_data["label"] = f"2Energy {cs}CS（2CS以上、送料込み）"

        elif brand == "2Gummy":
            group_data["threshold"] = 2
            if cs < 2:
                group_data["is_lot_break"] = True
                group_data["needs_manual"] = True
                group_data["fee"] = None
                group_data["method"] = "gummy_manual"
                group_data["external_tool_url"] = STREAMLIT_URL
                group_data["label"] = f"2Gummy {cs}CS ロット割れ（外部アプリで計算→手動入力）"
                needs_manual_overall = True
            else:
                group_data["label"] = f"2Gummy {cs}CS（2CS以上、送料なし）"

        elif brand == "2Water":
            group_data["threshold"] = 10
            # 2Water は ハルナ商品なので judge_shipping により <10CS は自動的に自社倉庫経由
            if cs < 10:
                group_data["is_lot_break"] = True
                group_data["needs_manual"] = True
                group_data["fee"] = None
                group_data["method"] = "water_manual"
                group_data["external_tool_url"] = STREAMLIT_URL
                group_data["label"] = f"2Water {cs}CS ロット割れ（自社倉庫経由・外部アプリで計算→手動入力）"
                needs_manual_overall = True
            else:
                group_data["label"] = f"2Water {cs}CS（10CS以上、ハルナ直送・送料なし）"

        else:
            group_data["label"] = f"{brand} {cs}CS（送料計算対象外）"

        groups_result.append(group_data)

    # jan → is_lot_break マップを生成（NE CSV 備考の出荷経路判定で使用）
    jans_lot_break = {}
    for g in groups_result:
        is_lb = bool(g.get("is_lot_break"))
        for it in g.get("items", []):
            jan = it.get("jan")
            if jan:
                jans_lot_break[jan] = is_lb

    return {
        "total_fee": total_fee,
        "currency": "円(税抜)",
        "prefecture": prefecture,
        "zone": zone,
        "groups": groups_result,
        "jans_lot_break": jans_lot_break,
        "needs_manual_input_overall": needs_manual_overall,
        "warnings": warnings,
    }


# ========== lot_break_history.csv 追記 (Snackのみ) ==========

# lot_break_history.csv 仕様メモ:
# - 既存ファイルのヘッダー行は 19カラム（伝票発行費・賞味期限管理費 を含む旧フォーマット）
# - 一方、2026-04 以降の実データは SKILL.md 仕様（18カラム）で記録されている
# - つまりヘッダーと実データが不整合な状態
# - 本モジュールは「最新運用＝18カラム」に揃える方針で追記する
_HISTORY_HEADERS = [
    "日時", "都道府県", "地帯", "ケース数", "個口数", "配送サイズ",
    "運賃", "保管費", "ピッキング費", "梱包手数料", "入庫費",
    "二重梱包資材", "デリバリーパック",
    "小計(税抜)", "消費税", "合計(税込)", "総袋数", "袋単価",
]


def append_lot_break_history(snack_detail: dict, irisuu: int = 12) -> bool:
    """2Snack ロット割れ計算結果を lot_break_history.csv に追記する。

    Args:
        snack_detail: calculate_snack_lot_break() の戻り値
        irisuu: 入数（2Snackは通常 12袋/CS）

    Returns:
        書込成功時 True

    Notes:
        - エンコーディング: UTF-8 (BOMなし)
        - ファイル新規作成時のみヘッダー書込。既存ファイルには追記のみ
        - SKILL.md 仕様の 18カラムで記録（既存実データに合わせる）
    """
    if not snack_detail.get("applicable"):
        return False

    case_count = snack_detail["case_count"]
    total_bags = case_count * irisuu
    total_incl_tax = snack_detail["total_incl_tax"]
    bag_unit_price = round(total_incl_tax / total_bags, 1) if total_bags > 0 else 0

    row = {
        "日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "都道府県": snack_detail["prefecture"],
        "地帯": snack_detail["zone"],
        "ケース数": case_count,
        "個口数": snack_detail["package_count"],
        "配送サイズ": snack_detail["carrier"],
        "運賃": snack_detail["carrier_fee"],
        "保管費": snack_detail["storage_fee"],
        "ピッキング費": snack_detail["picking_fee"],
        "梱包手数料": snack_detail["packaging_fee"],
        "入庫費": snack_detail["receiving_fee"],
        "二重梱包資材": snack_detail["double_pack_material"],
        "デリバリーパック": snack_detail["delivery_pack"],
        "小計(税抜)": snack_detail["subtotal_excl_tax"],
        "消費税": snack_detail["tax"],
        "合計(税込)": total_incl_tax,
        "総袋数": total_bags,
        "袋単価": bag_unit_price,
    }

    file_exists = os.path.exists(LOT_BREAK_HISTORY_CSV)
    try:
        os.makedirs(os.path.dirname(LOT_BREAK_HISTORY_CSV), exist_ok=True)
        with open(LOT_BREAK_HISTORY_CSV, "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=_HISTORY_HEADERS)
            if not file_exists:
                writer.writeheader()
            writer.writerow(row)
        return True
    except Exception as e:
        print(f"[shipping_fee] lot_break_history.csv 書込失敗: {e}")
        return False


# ========== CLI / テスト ==========

def _print_result(result: dict):
    """計算結果をターミナルに表示"""
    print()
    print("=" * 60)
    print("  送料計算結果")
    print("=" * 60)
    print(f"納品先都道府県: {result['prefecture']} / ゾーン: {result['zone']}")
    print()
    for g in result["groups"]:
        marker = "🚚" if g["is_lot_break"] else "  "
        fee_str = f"¥{g['fee']:,}" if g['fee'] is not None else "(手動入力必要)"
        print(f"  {marker} [{g['brand']}] {g['label']}")
        print(f"      合計CS: {g['total_cs']} / 閾値: {g['threshold']} / 送料: {fee_str}")
        if g.get("external_tool_url"):
            print(f"      🔗 {g['external_tool_url']}")
        if g["method"] == "snack_lot_break" and g["detail"].get("applicable"):
            d = g["detail"]
            print(f"      [内訳] 運賃 {d['carrier_fee']:,} / 保管 {d['storage_fee']:,} / "
                  f"ピッキング {d['picking_fee']:,} / 梱包 {d['packaging_fee']:,} / 入庫 {d['receiving_fee']:,} / "
                  f"二重梱包 {d['double_pack_material']:,} / デリバリーパック {d['delivery_pack']:,}")
            print(f"             配送会社: {d['carrier']} / 個口数: {d['package_count']}")
        print()
    print(f"📊 合計送料（税抜）: ¥{result['total_fee']:,}")
    if result["needs_manual_input_overall"]:
        print("⚠️  一部 brand は手動入力が必要です（外部Streamlitアプリで計算してください）")
    if result["warnings"]:
        print("\n⚠️ 警告:")
        for w in result["warnings"]:
            print(f"  - {w}")


def main():
    """CLI動作確認用。サンプル受注で計算してみる。"""
    import argparse
    parser = argparse.ArgumentParser(description="送料計算モジュール（CLI動作確認用）")
    parser.add_argument("--scenario", default="snack3_tokyo",
                        choices=["snack3_tokyo", "snack5_osaka", "snack9_okinawa",
                                 "energy1", "energy3", "gummy1", "water5_warehouse",
                                 "mixed"],
                        help="テストシナリオ")
    args = parser.parse_args()

    # サンプルJANは products テーブルに登録されているもの
    scenarios = {
        "snack3_tokyo": (
            [{"jan": "4589570801454", "matched": True, "quantity": 3, "master_name": "ガトーショコラ風サブレ"}],
            "東京都新宿区西新宿1-1-1", None
        ),
        "snack5_osaka": (
            [{"jan": "4589570801454", "matched": True, "quantity": 3, "master_name": "ガトーショコラ風"},
             {"jan": "4589570801416", "matched": True, "quantity": 2, "master_name": "香るトリュフ"}],
            "大阪府大阪市北区梅田1-1-1", None
        ),
        "snack9_okinawa": (
            [{"jan": "4589570801454", "matched": True, "quantity": 9, "master_name": "ガトーショコラ風"}],
            "沖縄県那覇市西1-1-1", None
        ),
        "energy1": (
            [{"jan": "4589570801348", "matched": True, "quantity": 1, "master_name": "2Energy 250ml"}],
            "東京都新宿区西新宿1-1-1", None
        ),
        "energy3": (
            [{"jan": "4589570801348", "matched": True, "quantity": 3, "master_name": "2Energy 250ml"}],
            "東京都新宿区西新宿1-1-1", None
        ),
        "gummy1": (
            [{"jan": "4589570801331", "matched": True, "quantity": 1, "master_name": "2Gummy LIPOSOME VC 50g"}],
            "東京都新宿区西新宿1-1-1", None
        ),
        "water5_warehouse": (
            [{"jan": "4589570801485", "matched": True, "quantity": 5, "master_name": "2Water Ceramide 500ml"}],
            "東京都新宿区西新宿1-1-1", "自社倉庫"
        ),
        "mixed": (
            [
                {"jan": "4589570801454", "matched": True, "quantity": 3, "master_name": "ガトーショコラ風"},
                {"jan": "4589570801348", "matched": True, "quantity": 1, "master_name": "2Energy 250ml"},
                {"jan": "4589570801331", "matched": True, "quantity": 1, "master_name": "2Gummy"},
            ],
            "神奈川県横浜市西区高島1-1-1", None
        ),
    }

    items, address, shipping_type = scenarios[args.scenario]
    print(f"シナリオ: {args.scenario}")
    print(f"  納品先: {address}")
    print(f"  出荷区分: {shipping_type}")
    print(f"  商品: {[(i['master_name'], i['quantity']) for i in items]}")

    result = calculate_shipping_fee(items, address, shipping_type)
    _print_result(result)


if __name__ == "__main__":
    main()
