"""
PDCA分析 AIエージェント
ふるさと納税 返礼品 ABC分析ツール
"""
import streamlit as st
import pandas as pd
import io
from analysis import load_data, compute_analysis, build_excel

# ─────────────────────────────────────────────
# ページ設定
# ─────────────────────────────────────────────
st.set_page_config(
    page_title='PDCA分析 AIエージェント',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='collapsed',
)

# ─────────────────────────────────────────────
# カスタムCSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ── 全体 ── */
html, body, [data-testid="stAppViewContainer"] {
    background: #f0f4f8;
    font-family: 'Arial', sans-serif;
}

/* ── ヘッダーバナー ── */
.banner {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 60%, #375623 100%);
    border-radius: 14px;
    padding: 36px 44px;
    margin-bottom: 24px;
    box-shadow: 0 6px 24px rgba(31,78,121,0.25);
    display: flex;
    align-items: center;
    gap: 28px;
}
.banner-icon { font-size: 64px; line-height:1; }
.banner-text h1 {
    color: #fff;
    font-size: 2.0rem;
    font-weight: 700;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
}
.banner-text p {
    color: rgba(255,255,255,0.82);
    font-size: 1.0rem;
    margin: 0;
}

/* ── ステップカード ── */
.step-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
    margin-bottom: 28px;
}
.step-card {
    background: #fff;
    border-radius: 12px;
    padding: 22px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.07);
    border-top: 4px solid;
    text-align: center;
}
.step-card.s1 { border-color: #2E75B6; }
.step-card.s2 { border-color: #ED7D31; }
.step-card.s3 { border-color: #375623; }
.step-icon { font-size: 2.2rem; margin-bottom: 10px; }
.step-title { font-weight: 700; font-size: 1.05rem; color: #1F4E79; margin-bottom: 6px; }
.step-desc  { font-size: 0.87rem; color: #555; }

/* ── アップロードエリア ── */
.upload-area {
    background: #fff;
    border-radius: 14px;
    padding: 32px 36px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    margin-bottom: 24px;
}
.upload-title {
    font-size: 1.2rem;
    font-weight: 700;
    color: #1F4E79;
    margin-bottom: 4px;
}
.upload-sub {
    font-size: 0.88rem;
    color: #777;
    margin-bottom: 18px;
}

/* ── KPIカード ── */
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 14px;
    margin-bottom: 28px;
}
.kpi-card {
    background: #fff;
    border-radius: 12px;
    padding: 20px 16px;
    text-align: center;
    box-shadow: 0 2px 10px rgba(0,0,0,0.07);
    border-bottom: 4px solid;
}
.kpi-label { font-size: 0.82rem; color: #777; font-weight: 600; margin-bottom: 8px; }
.kpi-value { font-size: 1.55rem; font-weight: 800; }

/* ── 結果テーブル ── */
.result-card {
    background: #fff;
    border-radius: 12px;
    padding: 24px 28px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    margin-bottom: 20px;
}
.result-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #1F4E79;
    border-left: 5px solid #2E75B6;
    padding-left: 12px;
    margin-bottom: 16px;
}

/* ── ダウンロードボタン ── */
.dl-section {
    background: linear-gradient(135deg, #375623 0%, #4E7A2E 100%);
    border-radius: 14px;
    padding: 36px 44px;
    text-align: center;
    box-shadow: 0 6px 20px rgba(55,86,35,0.30);
    margin-top: 8px;
}
.dl-title {
    color: #fff;
    font-size: 1.4rem;
    font-weight: 700;
    margin-bottom: 8px;
}
.dl-sub {
    color: rgba(255,255,255,0.8);
    font-size: 0.95rem;
    margin-bottom: 24px;
}
div[data-testid="stDownloadButton"] > button {
    background: #fff !important;
    color: #375623 !important;
    font-weight: 800 !important;
    font-size: 1.2rem !important;
    padding: 18px 60px !important;
    border-radius: 50px !important;
    border: none !important;
    box-shadow: 0 4px 16px rgba(0,0,0,0.20) !important;
    transition: transform 0.15s ease;
}
div[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.25) !important;
}

/* ── フッター ── */
.footer {
    text-align: center;
    color: #aaa;
    font-size: 0.8rem;
    padding: 24px 0 8px;
}

/* ── アラート ── */
.info-box {
    background: #EBF5FB;
    border-left: 5px solid #2E75B6;
    border-radius: 0 10px 10px 0;
    padding: 16px 20px;
    margin-bottom: 18px;
    font-size: 0.92rem;
    color: #1A5276;
}

/* ── ランクバッジ ── */
.badge-a { background:#C6EFCE; color:#276221; border-radius:6px; padding:2px 10px; font-weight:700; font-size:0.9rem; }
.badge-b { background:#FFEB9C; color:#9C6500; border-radius:6px; padding:2px 10px; font-weight:700; font-size:0.9rem; }
.badge-c { background:#FFC7CE; color:#9C0006; border-radius:6px; padding:2px 10px; font-weight:700; font-size:0.9rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# バナー
# ─────────────────────────────────────────────
st.markdown("""
<div class="banner">
  <div class="banner-icon">📊</div>
  <div class="banner-text">
    <h1>PDCA分析 AIエージェント</h1>
    <p>ふるさと納税 返礼品 売上データを自動分析 ― ABC分析・4象限スター分析・月別推移を一括出力</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ステップガイド
# ─────────────────────────────────────────────
st.markdown("""
<div class="step-grid">
  <div class="step-card s1">
    <div class="step-icon">📥</div>
    <div class="step-title">Step 1 — データをアップロード</div>
    <div class="step-desc">受注実績Excelファイル（.xlsx）を選択してください</div>
  </div>
  <div class="step-card s2">
    <div class="step-icon">⚙️</div>
    <div class="step-title">Step 2 — 自動分析</div>
    <div class="step-desc">AIが自動でABC分析・4象限分析・月別推移を計算</div>
  </div>
  <div class="step-card s3">
    <div class="step-icon">📤</div>
    <div class="step-title">Step 3 — ダウンロード</div>
    <div class="step-desc">6シート構成の完成Excelファイルをダウンロード</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ファイルアップロード
# ─────────────────────────────────────────────
st.markdown('<div class="upload-area">', unsafe_allow_html=True)
st.markdown('<div class="upload-title">📂 Excelファイルをアップロード</div>', unsafe_allow_html=True)
st.markdown('<div class="upload-sub">対応形式: .xlsx ／ 列構成: 商品コード・OG・販売年数・月別受注件数/売上金額/粗利益・合計列</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    label='ファイルをドラッグ&ドロップ、またはクリックして選択',
    type=['xlsx'],
    key='file_uploader',
    label_visibility='collapsed',
)
st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# メイン処理
# ─────────────────────────────────────────────
if uploaded is not None:
    with st.spinner('🔄 データを読み込んでいます...'):
        try:
            # ファイルを2回読めるようにバッファ化
            file_bytes = uploaded.read()
            buf1 = io.BytesIO(file_bytes)
            buf2 = io.BytesIO(file_bytes)

            data, months, sheet_name = load_data(buf1)

            if len(data) == 0:
                st.error('❌ データが見つかりませんでした。ファイル形式を確認してください。')
                st.stop()

        except Exception as e:
            st.error(f'❌ ファイルの読み込みに失敗しました: {e}')
            st.stop()

    with st.spinner('⚙️ 分析を実行中...'):
        try:
            data, stats = compute_analysis(data, months)
        except Exception as e:
            st.error(f'❌ 分析処理中にエラーが発生しました: {e}')
            st.stop()

    # 成功通知
    st.success(f'✅ 分析完了！  {stats["TOTAL_ITEMS"]}商品 ／ シート「{sheet_name}」を処理しました。')

    # ── KPIカード ──
    gp_rate_pct = stats['GP_RATE'] * 100
    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card" style="border-color:#1F4E79">
        <div class="kpi-label">🛍️ 総商品数</div>
        <div class="kpi-value" style="color:#1F4E79">{stats['TOTAL_ITEMS']:,} 商品</div>
      </div>
      <div class="kpi-card" style="border-color:#2E75B6">
        <div class="kpi-label">📦 年間受注件数</div>
        <div class="kpi-value" style="color:#2E75B6">{stats['TOTAL_ORDERS']:,} 件</div>
      </div>
      <div class="kpi-card" style="border-color:#375623">
        <div class="kpi-label">💰 年間売上金額</div>
        <div class="kpi-value" style="color:#375623">¥{stats['TOTAL_SALES']:,}</div>
      </div>
      <div class="kpi-card" style="border-color:#7B3F00">
        <div class="kpi-label">📈 年間粗利益</div>
        <div class="kpi-value" style="color:#7B3F00">¥{stats['TOTAL_GP']:,}</div>
      </div>
      <div class="kpi-card" style="border-color:#C00000">
        <div class="kpi-label">📊 粗利率</div>
        <div class="kpi-value" style="color:#C00000">{gp_rate_pct:.1f}%</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── タブ別プレビュー ──
    tab1, tab2, tab3, tab4 = st.tabs([
        '🏆 売上トップ10',
        '🔵 OG別 ABC分析',
        '📅 販売年数別サマリー',
        '🎯 4象限スター分析'
    ])

    with tab1:
        st.markdown('<div class="result-card"><div class="result-title">売上金額 トップ10商品</div>', unsafe_allow_html=True)
        top10 = data.sort_values('合計売上金額', ascending=False).head(10)[
            ['返礼品名','OG','販売年数','合計受注件数','合計売上金額','合計粗利益','粗利率','売上ランク']
        ].copy()
        top10['合計売上金額'] = top10['合計売上金額'].map('¥{:,.0f}'.format)
        top10['合計粗利益']   = top10['合計粗利益'].map('¥{:,.0f}'.format)
        top10['粗利率']        = top10['粗利率'].map('{:.1%}'.format)
        top10['合計受注件数']  = top10['合計受注件数'].map('{:,.0f}'.format)
        top10 = top10.rename(columns={
            '返礼品名':'商品名','合計受注件数':'受注件数',
            '合計売上金額':'売上金額','合計粗利益':'粗利益','売上ランク':'ランク'
        })
        st.dataframe(top10, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="result-card"><div class="result-title">OG別 売上金額ランキング</div>', unsafe_allow_html=True)
        og_df = data.groupby('OG').agg(
            商品数=('商品コード','count'),
            受注件数=('合計受注件数','sum'),
            売上金額=('合計売上金額','sum'),
            粗利益=('合計粗利益','sum')
        ).reset_index()
        og_df['粗利率'] = og_df['粗利益'] / og_df['売上金額']
        og_df['売上ランク'] = og_df['売上金額'] / og_df['売上金額'].sum()
        og_df = og_df.sort_values('売上金額', ascending=False)
        og_df['売上金額'] = og_df['売上金額'].map('¥{:,.0f}'.format)
        og_df['粗利益']   = og_df['粗利益'].map('¥{:,.0f}'.format)
        og_df['粗利率']   = og_df['粗利率'].map('{:.1%}'.format)
        og_df['売上ランク'] = og_df['売上ランク'].map('{:.1%}'.format)
        og_df = og_df.rename(columns={'売上ランク':'売上シェア','受注件数':'受注件数計'})
        st.dataframe(og_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="result-card"><div class="result-title">販売年数別 集計</div>', unsafe_allow_html=True)
        yr_df = data.groupby('販売年数').agg(
            商品数=('商品コード','count'),
            受注件数=('合計受注件数','sum'),
            売上金額=('合計売上金額','sum'),
            粗利益=('合計粗利益','sum')
        ).reindex(['1年生','2年生','3年生']).reset_index()
        yr_df['粗利率'] = yr_df['粗利益'] / yr_df['売上金額']
        yr_df['売上シェア'] = yr_df['売上金額'] / yr_df['売上金額'].sum()
        yr_df['売上金額'] = yr_df['売上金額'].map('¥{:,.0f}'.format)
        yr_df['粗利益']   = yr_df['粗利益'].map('¥{:,.0f}'.format)
        yr_df['粗利率']   = yr_df['粗利率'].map('{:.1%}'.format)
        yr_df['売上シェア'] = yr_df['売上シェア'].map('{:.1%}'.format)
        st.dataframe(yr_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab4:
        q_counts = data['象限分類'].value_counts()
        q_order  = ['Q1: スター','Q2: 高収益','Q3: 量販型','Q4: 要改善']
        q_colors_map = {
            'Q1: スター': ('#C6EFCE','#276221','⭐'),
            'Q2: 高収益': ('#BDD7EE','#1F4E79','💎'),
            'Q3: 量販型': ('#FFC7CE','#9C0006','📦'),
            'Q4: 要改善': ('#FFEB9C','#9C6500','🔧'),
        }
        cols4 = st.columns(4)
        for i, quad in enumerate(q_order):
            cnt = int(q_counts.get(quad, 0))
            bg, fg, icon = q_colors_map[quad]
            with cols4[i]:
                st.markdown(f"""
                <div style="background:{bg};border-radius:12px;padding:20px;text-align:center;border:1px solid {fg}33">
                  <div style="font-size:2rem">{icon}</div>
                  <div style="font-weight:700;color:{fg};font-size:1.05rem;margin:8px 0 4px">{quad}</div>
                  <div style="font-size:2rem;font-weight:800;color:{fg}">{cnt}</div>
                  <div style="color:{fg};font-size:0.85rem">商品</div>
                </div>
                """, unsafe_allow_html=True)

    st.markdown('<br>', unsafe_allow_html=True)

    # ── Excel生成 & ダウンロード ──
    with st.spinner('📝 Excelファイルを生成中...'):
        try:
            title_prefix = uploaded.name.replace('.xlsx','').replace('_','　')[:20]
            excel_buf = build_excel(data, stats, months, title_prefix=title_prefix)
            excel_bytes = excel_buf.getvalue()
        except Exception as e:
            st.error(f'❌ Excel生成中にエラーが発生しました: {e}')
            st.stop()

    # ダウンロードセクション
    out_filename = uploaded.name.replace('.xlsx', '') + '_ABC分析済.xlsx'

    st.markdown("""
    <div class="dl-section">
      <div class="dl-title">📤 分析完了！Excelファイルをダウンロード</div>
      <div class="dl-sub">6シート構成（ダッシュボード・商品データ・OG別・販売年数別・円グラフ・4象限スター分析）</div>
    </div>
    """, unsafe_allow_html=True)

    col_dl_l, col_dl_c, col_dl_r = st.columns([2, 3, 2])
    with col_dl_c:
        st.download_button(
            label='⬇️　完成Excelをダウンロード',
            data=excel_bytes,
            file_name=out_filename,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )

    # 出力シート一覧
    with st.expander('📋 出力シート構成を確認', expanded=False):
        st.markdown("""
| # | シート名 | 内容 |
|---|---------|------|
| 1 | **ダッシュボード** | KPIカード・月別売上/受注バーチャート・カテゴリ別サマリー |
| 2 | **商品データ** | 全商品 + ABC3ランク + 累積比率 + 4象限分類 |
| 3 | **OG別ABC分析** | 26 OGの受注・売上・粗利を横並び比較（累積比率付き） |
| 4 | **販売年数別ABC分析** | 1年生/2年生/3年生グループ内でそれぞれランク付け |
| 5 | **販売年数別円グラフ** | 受注件数・売上金額・粗利益の構成比 円グラフ3枚 |
| 6 | **4象限スター分析** | Q1スター/Q2高収益/Q3量販型/Q4要改善 色分け一覧 |
        """)

else:
    # アップロード前のガイド
    st.markdown("""
    <div class="info-box">
      <strong>📌 ご利用方法</strong><br>
      上の「ファイルをドラッグ&ドロップ」エリアに、受注実績Excelファイル（.xlsx）をアップロードしてください。<br>
      ファイルは自動で分析され、プレビューと完成Excelダウンロードボタンが表示されます。
    </div>
    """, unsafe_allow_html=True)

    # 期待するフォーマット説明
    with st.expander('📋 対応しているExcelファイルのフォーマット', expanded=True):
        st.markdown("""
**必須列（標準的な北国からの贈り物フォーマット）**

| 列 | 内容 |
|----|------|
| A | 商品コード |
| B | 返礼品名 |
| C | OG（オリジン） |
| D | OGファミリー |
| E | カテゴリ |
| F | 分類 |
| G | 販売年数（1年生/2年生/3年生） |
| H〜K | 寄付額・返礼額・商品原価・単位粗利益 |
| L〜AV | 月別（1月〜12月）受注件数・売上金額・粗利益 |
| AV〜AX | 合計 受注件数・売上金額・粗利益 |

> データは3行目以降（行1=タイトル、行2=ヘッダー、行3=サブヘッダー）から読み込まれます。
        """)

# ─────────────────────────────────────────────
# フッター
# ─────────────────────────────────────────────
st.markdown("""
<div class="footer">
  PDCA分析 AIエージェント ― Powered by Claude ／ Built with Streamlit & openpyxl
</div>
""", unsafe_allow_html=True)
