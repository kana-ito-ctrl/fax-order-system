"""OCR出力検証スクリプト - テスト用期待値との比較

使い方:
    python verify_ocr_output.py <result_csv_path>

例:
    python verify_ocr_output.py output/20260313_直送_result.csv
"""
import csv
import sys
import os

# ==================== 期待値定義 ====================
# 20260313_直送.pdf の正解データ（ページ画像から目視確認済み）
EXPECTED = {
    "20260313_直送": [
        # Page 1: 発注書 伊藤忠食品 船橋物流センター
        {
            "page": 1,
            "delivery_dest_contains": "船橋",
            "ddc_match": "OK",          # 伊藤忠食品㈱ 船橋物流センター (oroshisaki_code 601)
            "items": [
                {"name_contains": "ガーリック", "qty": 8},
                {"name_contains": "トリュフ",   "qty": 5},
            ],
        },
        # Page 2: 発注書 加藤産業 所沢物流センター
        {
            "page": 2,
            "delivery_dest_contains": "所沢",
            "ddc_match": "OK",
            "items": [
                {"name_contains": "ガーリック", "qty_min": 1},
            ],
        },
        # Page 3: 発注書 伊藤忠食品 昭島物流センター
        {
            "page": 3,
            "delivery_dest_contains": "昭島",
            "ddc_match": "OK",
            "items": [
                {"name_contains": "ガトーショコラ", "qty": 5},
                {"name_contains": "ガーリック",     "qty": 2},
                {"name_contains": "和紅茶",         "qty": 4},
                {"name_contains": "トリュフ",       "qty": 5},
            ],
        },
        # Pages 4-6: ハルナ (2Water Ceramide)
        {
            "page": 4,
            "items": [{"name_contains": "Ceramide", "qty_min": 1}],
        },
        {
            "page": 5,
            "items": [{"name_contains": "Ceramide", "qty_min": 1}],
        },
        {
            "page": 6,
            "items": [{"name_contains": "Ceramide", "qty_min": 1}],
        },
        # Page 7: FAX発注票 三菱食品 埼京SDC → ローソン埼京DDC
        {
            "page": 7,
            "delivery_dest_contains": "埼",          # 埼京 or 埼玉
            "ddc_match_not": "NG",                   # OK or 要確認 (三菱食品 ローソン埼京DDC)
            "items": [
                {"name_contains": "Ceramide", "qty": 96},
            ],
        },
    ]
}
# ====================================================


def load_csv(path):
    rows = []
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def check_file(csv_path, base_name):
    if base_name not in EXPECTED:
        print(f"  ⚠ 期待値定義なし: {base_name}")
        return True

    rows = load_csv(csv_path)
    if not rows:
        print("  ✗ CSVが空です（処理失敗）")
        return False

    expected_pages = {e["page"]: e for e in EXPECTED[base_name]}
    passed = 0
    failed = 0

    for page_num, exp in sorted(expected_pages.items()):
        page_rows = [r for r in rows if str(r.get("ページ", "")) == str(page_num)]
        if not page_rows:
            print(f"  ✗ Page {page_num}: 行なし")
            failed += 1
            continue

        errors = []

        # DDCマッチ確認
        if "ddc_match" in exp:
            actual_ddc = page_rows[0].get("DDCマッチ", "")
            if actual_ddc != exp["ddc_match"]:
                errors.append(f"DDCマッチ: 期待={exp['ddc_match']}, 実際={actual_ddc}")

        if "ddc_match_not" in exp:
            actual_ddc = page_rows[0].get("DDCマッチ", "")
            if actual_ddc == exp["ddc_match_not"]:
                cands = page_rows[0].get("DDC候補1", "")
                errors.append(f"DDCマッチ=NG（候補: {cands}）")

        # 納品先確認
        if "delivery_dest_contains" in exp:
            actual_dest = page_rows[0].get("納品先(OCR)", "")
            if exp["delivery_dest_contains"] not in actual_dest:
                errors.append(
                    f"納品先: '{exp['delivery_dest_contains']}' が "
                    f"'{actual_dest}' に含まれない"
                )

        # 商品・数量確認
        for item_exp in exp.get("items", []):
            kw = item_exp["name_contains"]
            matched_rows = [
                r for r in page_rows
                if kw in r.get("商品名(OCR)", "") or kw in r.get("商品名(マスタ)", "")
            ]
            if not matched_rows:
                errors.append(f"商品「{kw}」が見つからない")
                continue
            if "qty" in item_exp:
                actual_qty = int(matched_rows[0].get("数量(CS)", 0) or 0)
                if actual_qty != item_exp["qty"]:
                    errors.append(
                        f"商品「{kw}」数量: 期待={item_exp['qty']}, 実際={actual_qty}"
                    )
            if "qty_min" in item_exp:
                actual_qty = int(matched_rows[0].get("数量(CS)", 0) or 0)
                if actual_qty < item_exp["qty_min"]:
                    errors.append(
                        f"商品「{kw}」数量: 期待>={item_exp['qty_min']}, 実際={actual_qty}"
                    )

        if errors:
            print(f"  ✗ Page {page_num}:")
            for e in errors:
                print(f"      - {e}")
            failed += 1
        else:
            dest = page_rows[0].get("納品先(OCR)", "")[:25]
            ddc = page_rows[0].get("DDCマッチ", "")
            print(f"  ✓ Page {page_num}: {dest} [{ddc}]")
            passed += 1

    print(f"\n  結果: {passed}✓ / {failed}✗ (計{passed+failed}ページ)")
    return failed == 0


def main():
    if len(sys.argv) < 2:
        print("使い方: python verify_ocr_output.py <result_csv_path>")
        print("例:     python verify_ocr_output.py output/20260313_直送_result.csv")
        sys.exit(1)

    csv_path = sys.argv[1]
    if not os.path.exists(csv_path):
        print(f"ファイルが見つかりません: {csv_path}")
        sys.exit(1)

    base_name = os.path.basename(csv_path).replace("_result.csv", "")
    print(f"検証中: {base_name}")
    print("=" * 50)

    ok = check_file(csv_path, base_name)
    print("=" * 50)
    if ok:
        print("✅ 全チェック通過")
    else:
        print("⚠️  要確認項目あり（上記参照）")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
