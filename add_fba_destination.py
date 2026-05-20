"""Amazon FBA 納品先 登録ツール（対話型）

Usage:
    python add_fba_destination.py             # 本番登録
    python add_fba_destination.py --dry-run   # 登録せずに動作確認のみ

機能:
    1. 取引先マスタ (torihikisaki_master_scm) に Amazon を登録（未登録の場合のみ）
    2. 納品先マスタ (nohinsaki_master_scm) に FBA 倉庫を登録（複数登録可）
    3. 任意で ハルナ条件 (haruna_conditions) も同時登録

明日 (2026-04-30) の業務担当者との打合せでその場で実行する想定。
"""
import os
import sys
import json
import argparse
import urllib.request
import urllib.error
from datetime import datetime

# Windows ターミナル UTF-8 補正
if sys.stdout.encoding and sys.stdout.encoding.lower() in ('cp932', 'shift_jis', 'mbcs'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from supabase_client import (
    SUPABASE_URL, SUPABASE_KEY,
    _supabase_get, _supabase_post,
)

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        f"add_fba_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")


def _print_header(text):
    print()
    print("=" * 60)
    print(f"  {text}")
    print("=" * 60)


def _input(prompt, default="", required=False):
    """対話入力。default を表示、required=True の場合は空入力を拒否。"""
    suffix = f" [{default}]" if default else ""
    while True:
        v = input(f"{prompt}{suffix}: ").strip()
        if not v and default:
            return default
        if not v and required:
            print("  ⚠️  必須項目です。入力してください。")
            continue
        return v


def _yn(prompt, default="y"):
    """y/n 確認。"""
    s = input(f"{prompt} ({'Y/n' if default == 'y' else 'y/N'}): ").strip().lower()
    if not s:
        return default == "y"
    return s in ("y", "yes")


def find_amazon_torihikisaki():
    """Amazon 取引先がマスタにあるか確認。"""
    rows = _supabase_get("torihikisaki_master_scm",
                         "select=torihikisaki_code,torihikisaki_name&is_active=eq.true")
    if not rows:
        return None
    for r in rows:
        name = (r.get("torihikisaki_name") or "").lower()
        if any(kw in name for kw in ["amazon", "アマゾン"]):
            return r
    return None


def get_next_torihikisaki_code():
    """次の取引先コード (T1201 形式) を取得。"""
    rows = _supabase_get("torihikisaki_master_scm",
                         "select=torihikisaki_code&torihikisaki_code=like.T*&order=torihikisaki_code.desc&limit=1")
    if not rows:
        return "T0001"
    last = rows[0]["torihikisaki_code"]
    try:
        n = int(last[1:]) + 1
        return f"T{n:04d}"
    except (ValueError, IndexError):
        return "T9000"  # フォールバック


def get_next_nohinsaki_code(start: int = 9001):
    """次の納品先コードを取得（既存数値コードのうち start 以上の最大値 + 1、最低 start）。

    nohinsaki_code は VARCHAR のため辞書順比較が効かない。Python側で数値比較する。
    """
    rows = _supabase_get("nohinsaki_master_scm", "select=nohinsaki_code")
    if not rows:
        return str(start)
    nums = []
    for r in rows:
        c = r.get("nohinsaki_code") or ""
        try:
            n = int(c)
            if n >= start:
                nums.append(n)
        except (ValueError, TypeError):
            continue
    if not nums:
        return str(start)
    return str(max(nums) + 1)


def register_amazon_torihikisaki(dry_run=False):
    """Amazon 取引先を登録（既存の場合はそれを返す）。"""
    existing = find_amazon_torihikisaki()
    if existing:
        print(f"✓ Amazon 取引先は既に登録済: {existing['torihikisaki_name']} ({existing['torihikisaki_code']})")
        return existing

    print("Amazon 取引先がマスタに未登録です。新規登録します。")
    next_code = get_next_torihikisaki_code()

    name = _input("取引先名", default="アマゾンジャパン合同会社", required=True)
    code = _input("取引先コード（既定値推奨）", default=next_code, required=True)

    payload = {
        "torihikisaki_code": code,
        "torihikisaki_name": name,
        "is_active": True,
    }
    print(f"\n登録内容: {payload}")
    if not _yn("この内容で登録しますか?"):
        print("中止しました。")
        return None

    if dry_run:
        print("[DRY-RUN] 取引先INSERTをスキップ")
        return payload

    result = _supabase_post("torihikisaki_master_scm", payload)
    if not result:
        print("❌ 取引先登録失敗。Supabaseエラーを確認してください。")
        return None
    print(f"✅ 取引先登録完了: {code} / {name}")
    return result[0] if result else payload


def register_fba_destination(torihikisaki_code: str, dry_run=False):
    """FBA 倉庫を1件登録。"""
    print()
    print("-" * 60)
    print("  FBA 納品先 登録")
    print("-" * 60)

    next_code = get_next_nohinsaki_code()
    code = _input("納品先コード（既定値推奨）", default=next_code, required=True)
    name = _input("納品先名 (例: Amazon FBA TYO9)", required=True)
    yubin = _input("郵便番号 (ハイフン無しでもOK / 例: 270-0000)").replace("-", "")
    jusho = _input("住所", required=True)
    tel = _input("電話番号").replace("-", "")

    print()
    print("📋 確認:")
    print(f"  納品先コード: {code}")
    print(f"  納品先名:     {name}")
    print(f"  取引先コード: {torihikisaki_code}")
    print(f"  郵便番号:     {yubin}")
    print(f"  住所:         {jusho}")
    print(f"  電話番号:     {tel}")
    if not _yn("\nこの内容で登録しますか?"):
        print("  → スキップしました")
        return None

    payload = {
        "nohinsaki_code": code,
        "nohinsaki_name": name,
        "torihikisaki_code": torihikisaki_code,
        "yubin_bango": yubin,
        "jusho": jusho,
        "denwa_bango": tel,
        "is_active": True,
    }

    if dry_run:
        print("[DRY-RUN] 納品先INSERTをスキップ")
    else:
        result = _supabase_post("nohinsaki_master_scm", payload)
        if not result:
            print("❌ 納品先登録失敗。Supabaseエラーを確認してください。")
            return None
        print(f"✅ 納品先登録完了: {code} / {name}")

    # ハルナ条件も登録するか確認
    if _yn("\nハルナ条件（入荷時間・バース予約等）も登録しますか?", default="n"):
        haruna_payload = register_haruna_condition(name, dry_run=dry_run)
        if haruna_payload:
            payload["_haruna_conditions"] = haruna_payload

    return payload


def register_haruna_condition(nohinsaki_name: str, dry_run=False):
    """ハルナ条件を登録。"""
    print()
    print("ハルナ条件（任意項目・空欄でEnter可）")
    arrival_time = _input("入荷時間 (例: 9:00-11:00)")
    basse = _input("バース予約 (有/無/任意の文字列)")
    pallet = _input("パレット条件")
    jpr = _input("JPRコード")

    payload = {
        "nohinsaki": nohinsaki_name,
        "arrival_time": arrival_time or None,
        "basse_reservation": basse or None,
        "pallet_condition": pallet or None,
        "jpr_code": jpr or None,
    }
    if dry_run:
        print(f"[DRY-RUN] ハルナ条件INSERTをスキップ: {payload}")
        return payload

    result = _supabase_post("haruna_conditions", payload)
    if result:
        print("✅ ハルナ条件登録完了")
    else:
        print("⚠️ ハルナ条件登録失敗（ハルナ条件なしで継続可能）")
    return payload


def main():
    parser = argparse.ArgumentParser(description="Amazon FBA 納品先 登録ツール")
    parser.add_argument("--dry-run", action="store_true",
                        help="Supabaseへの書込をスキップして動作確認のみ")
    args = parser.parse_args()

    _print_header("Amazon FBA 納品先 登録ツール")
    if args.dry_run:
        print("⚠️  DRY-RUN モード: 実際にはSupabaseに登録しません")

    print(f"接続先: {SUPABASE_URL}")
    print()

    # ステップ1: 取引先（Amazon）登録
    _print_header("Step 1/2: 取引先マスタ (Amazon) の確認")
    amazon = register_amazon_torihikisaki(dry_run=args.dry_run)
    if not amazon:
        print("取引先登録に失敗したため終了します。")
        return 1

    torihikisaki_code = amazon.get("torihikisaki_code")

    # ステップ2: FBA倉庫を複数登録
    _print_header("Step 2/2: FBA倉庫を納品先マスタに登録")
    print("（複数登録できます。終了するときは「もう1件登録?」でnを選んでください）")

    registered = []
    while True:
        result = register_fba_destination(torihikisaki_code, dry_run=args.dry_run)
        if result:
            registered.append(result)

        if not _yn("\nもう1件 FBA倉庫を登録しますか?"):
            break

    # サマリー
    _print_header("完了")
    print(f"登録件数: {len(registered)}件")
    for r in registered:
        print(f"  - {r.get('nohinsaki_code')} / {r.get('nohinsaki_name')}")

    # ログ保存
    if registered and not args.dry_run:
        log_data = {
            "timestamp": datetime.now().isoformat(),
            "torihikisaki": amazon,
            "registered_destinations": registered,
        }
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)
        print(f"\nログ保存: {LOG_FILE}")

    print("\n📌 次のステップ:")
    print("  1. FAX受注処理システム（Render版）でマスタ更新ボタンを押す")
    print("     https://fax-order-system.onrender.com/")
    print("  2. 「手入力で受注登録」→ 納品先プルダウンに新規登録した倉庫が出ることを確認")

    return 0


if __name__ == "__main__":
    sys.exit(main())
