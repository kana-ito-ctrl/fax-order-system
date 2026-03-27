"""
ネクストエンジンCSV → Supabase products テーブル アップロードスクリプト
"""
import csv
import json
import sys
import urllib.request
import urllib.error

sys.stdout.reconfigure(encoding='utf-8')

SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
API_KEY = os.environ.get("SUPABASE_KEY", "")
CSV_PATH = r'C:\Users\tw2407-044\Downloads\data2026030916103282656400.csv'

HEADERS_HTTP = {
    "apikey": API_KEY,
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json",
    "Prefer": "resolution=ignore-duplicates,return=minimal",
}


def get_temperature(tag: str) -> str:
    """商品分類タグから温度帯を判定"""
    if "冷凍" in tag:
        return "冷凍"
    elif "冷蔵" in tag:
        return "冷蔵"
    else:
        return "常温"


def parse_int(val: str) -> int | None:
    """文字列を整数に変換（空文字はNone）"""
    val = val.strip()
    if not val:
        return None
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None


def load_csv(path: str) -> list[dict]:
    """CSVを読み込んでSupabase用レコードリストを返す"""
    records = []
    with open(path, encoding='cp932', errors='replace') as f:
        reader = csv.reader(f)
        headers = next(reader)
        print(f"CSVヘッダー: {headers}")

        for row in reader:
            if len(row) < 20:
                continue
            product_code = row[0].strip()
            if not product_code:
                continue  # コードなしはスキップ

            product_name = row[1].strip()
            retail_price = parse_int(row[2])
            tag = row[21].strip() if len(row) > 21 else ""
            jan = row[19].strip() if len(row) > 19 else ""
            is_active = row[6].strip() == "取扱中"
            temperature = get_temperature(tag)

            record = {
                "product_code": product_code,
                "product_name": product_name,
                "seller": "TWO",
                "is_active": is_active,
                "is_water": False,
                "is_gummy": False,
                "sort_order": 0,
                "temperature": temperature,
                "retail_price": retail_price,
                "jan": jan if jan else None,
            }

            records.append(record)

    return records


def upsert_batch(records: list[dict], batch_size: int = 50) -> None:
    """バッチでupsert"""
    total = len(records)
    success = 0
    failed = 0

    for i in range(0, total, batch_size):
        batch = records[i:i + batch_size]
        data = json.dumps(batch).encode('utf-8')

        req = urllib.request.Request(
            f"{SUPABASE_URL}/rest/v1/products",
            data=data,
            headers=HEADERS_HTTP,
            method="POST",
        )
        try:
            with urllib.request.urlopen(req) as resp:
                success += len(batch)
                print(f"  OK: {i+1}〜{min(i+batch_size, total)}件目")
        except urllib.error.HTTPError as e:
            body = e.read().decode('utf-8', errors='replace')
            print(f"  ERROR ({e.code}) batch {i//batch_size + 1}: {body[:300]}")
            failed += len(batch)

    print(f"\n完了: 成功={success}, 失敗={failed}, 合計={total}")


def main():
    print("=== CSVを読み込み中 ===")
    records = load_csv(CSV_PATH)
    print(f"読み込み件数: {len(records)}")

    print("\n=== Supabaseへアップロード中 ===")
    upsert_batch(records)


if __name__ == "__main__":
    main()
