# FAX発注書 自動処理システム

株式会社TWO 事業管理部向けの、FAX発注書自動処理Webアプリケーション。

## 機能

- **FAX読取モード**：FAX PDFをアップロード → Claude Vision APIで自動読み取り → 発注書PDF生成
- **手動入力モード**：手動でデータ入力 → 発注書PDF生成
- **2種類の出力**：ハルナプロデュース向け（青テーマ）/ シルビア向け（茶テーマ）
- **マスタ自動照合**：商品マスタ（JAN/商品名）、DDCマスタ（納品先）

## セットアップ

### Streamlit Community Cloudでデプロイ

1. このリポジトリをGitHubにプッシュ
2. [share.streamlit.io](https://share.streamlit.io) でデプロイ
3. Settings → Secrets に以下を設定：

```
ANTHROPIC_API_KEY = "sk-ant-..."
```

### ローカル実行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ファイル構成

```
├── app.py                 # メインアプリ
├── pdf_generator.py       # PDF生成（シルビアv12 / ハルナ最終版）
├── ocr_module.py          # OCR（Claude Vision API）
├── requirements.txt       # Python依存関係
├── packages.txt           # システム依存関係（日本語フォント）
├── .streamlit/
│   └── config.toml        # テーマ設定
└── data/
    ├── product_master.csv  # 商品マスタ
    ├── ddc_master.csv      # 納品先マスタ
    └── staff.json          # 担当者情報
```
