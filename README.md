# PDCA分析 AIエージェント

ふるさと納税 返礼品の売上データを自動分析する Streamlit Webアプリです。

## 起動方法

```bash
# 1. 依存ライブラリをインストール
pip install -r requirements.txt

# 2. アプリを起動
streamlit run app.py
```

ブラウザで `http://localhost:8501` が自動的に開きます。

## 使い方

1. 受注実績 Excel ファイル（.xlsx）をアップロード
2. 自動分析が実行されKPIとプレビューが表示される
3. 「完成Excelをダウンロード」ボタンをクリック

## 出力Excelシート構成

| シート | 内容 |
|--------|------|
| ダッシュボード | KPI・月別バーチャート・カテゴリ別サマリー |
| 商品データ | 全商品 + ABCランク + 累積比率 + 4象限分類 |
| OG別ABC分析 | 受注・売上・粗利を横並び比較（累積比率付き） |
| 販売年数別ABC分析 | 1年生/2年生/3年生グループ内ランク付け |
| 販売年数別円グラフ | 構成比 円グラフ × 3指標 |
| 4象限スター分析 | Q1スター/Q2高収益/Q3量販型/Q4要改善 |

## クラウドデプロイ（Streamlit Cloud）

1. GitHubリポジトリにこのフォルダをプッシュ
2. [share.streamlit.io](https://share.streamlit.io) でリポジトリを接続
3. `app.py` をメインファイルとして指定
4. デプロイ完了（無料で外部公開可能）

## ファイル構成

```
pdca_app/
├── app.py           # Streamlit アプリ本体
├── analysis.py      # データ読み込み・分析・Excel生成ロジック
├── requirements.txt # 依存ライブラリ
└── README.md        # このファイル
```
