# -*- coding: utf-8 -*-
"""
app.py
部品発注リスト自動補完システム ― Streamlit WebUI

Google Cloud / 認証情報(credentials.json)は不要。
発注履歴スプレッドシートを「リンクを知っている全員が閲覧可」にして、
そのURLを貼り付けるだけで最新データを取得します(方式A)。
ローカルの履歴ファイル(.xlsx/.csv)を読む方式Bも選べます。

起動方法:
    streamlit run app.py
"""

import base64
import os
import traceback

import pandas as pd
import streamlit as st

import order_core as core

st.set_page_config(page_title="部品発注リスト自動補完", page_icon="📦", layout="wide")


# ---------------------------------------------------------------------------
# 物流倉庫テーマ(カスタムCSS)
# ---------------------------------------------------------------------------
def inject_theme_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=M+PLUS+1p:wght@400;500;700;800&display=swap');

        html, body, [class*="css"], .stApp, [data-testid="stSidebar"],
        input, textarea, button, select {
            font-family: 'M PLUS 1p', 'Yu Gothic', 'Meiryo', sans-serif !important;
        }

        /* 背景: コンクリート＋うっすら棚グリッド */
        .stApp {
            background-color: #eef1f5;
            background-image:
                linear-gradient(rgba(100,116,139,.06) 1px, transparent 1px),
                linear-gradient(90deg, rgba(100,116,139,.06) 1px, transparent 1px);
            background-size: 42px 42px;
            color: #1f2a37;
        }
        /* 上端のハザードストライプ(安全色) */
        .stApp::before {
            content: ""; position: fixed; top: 0; left: 0; right: 0; height: 7px; z-index: 5;
            background: repeating-linear-gradient(45deg, #FFC400 0 14px, #1a1a1a 14px 28px);
        }

        /* 文字色 */
        .stApp, .stMarkdown, .stMarkdown p, .stMarkdown li, .stApp p, .stApp li,
        label, [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p,
        .stRadio label, .stCheckbox label, .stSlider label,
        [data-testid="stExpander"] summary, [data-testid="stMarkdownContainer"] {
            color: #1f2a37 !important;
        }
        [data-testid="stCaptionContainer"], .stCaption, small { color: #5b6675 !important; }
        [data-testid="stHeader"] { background: transparent; }

        /* 見出し: 濃いスレート＋アンバーの下線 */
        h1, h2, h3 {
            font-family: 'M PLUS 1p', sans-serif !important;
            color: #1f2a37 !important;
            font-weight: 800 !important;
            letter-spacing: .5px;
        }
        h2, h3 { border-left: 6px solid #F5A623; padding-left: 10px; }

        .block-container { padding-top: 2.2rem; max-width: 1150px; }

        /* サイドバー: 明るいスチール */
        [data-testid="stSidebar"] > div:first-child {
            background: #f4f6fa;
            border-right: 2px solid #cfd6e2;
        }
        [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { border-left-color:#2c5282; }

        /* ボタン: スチールスレート */
        .stButton > button, .stDownloadButton > button {
            border-radius: 8px;
            border: 1px solid #1f2a37 !important;
            background: #334155;
            color: #fff !important;
            font-weight: 700;
            padding: .5rem 1rem;
            box-shadow: 0 2px 6px rgba(31,42,55,.18);
            transition: all .15s ease;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            background: #1f2a37; transform: translateY(-1px);
            box-shadow: 0 5px 14px rgba(31,42,55,.28); color:#fff !important;
        }
        /* 主要ボタン(実行/DL): 安全アンバー */
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
            background: #F5A623; color:#1f2a37 !important; border-color:#b9791000 !important;
            border: 1px solid #b97910 !important;
        }
        .stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
            background: #e0940f; color:#1f2a37 !important;
        }
        /* ボタン内ラベル(markdown)の文字色を明示(全体ルールに負けないように) */
        .stButton > button p, .stButton > button [data-testid="stMarkdownContainer"],
        .stDownloadButton > button p, .stDownloadButton > button [data-testid="stMarkdownContainer"] {
            color: #ffffff !important;
        }
        .stButton > button[kind="primary"] p, .stButton > button[kind="primary"] [data-testid="stMarkdownContainer"],
        .stDownloadButton > button[kind="primary"] p, .stDownloadButton > button[kind="primary"] [data-testid="stMarkdownContainer"] {
            color: #1f2a37 !important;
        }

        /* メトリクス: 白カード＋アンバーの左帯 */
        [data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #d6dce6;
            border-left: 6px solid #F5A623;
            border-radius: 10px;
            padding: 14px 18px;
            box-shadow: 0 4px 12px rgba(31,42,55,.08);
        }
        [data-testid="stMetricValue"] { color:#1f2a37 !important; font-weight:800; }
        [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] p { color:#5b6675 !important; }

        /* 入力欄 */
        .stTextInput input, .stNumberInput input, .stTextArea textarea,
        [data-baseweb="select"] > div {
            border-radius: 8px !important;
            background: #fff !important;
            border: 1px solid #cbd5e1 !important;
            color: #1f2a37 !important;
        }
        .stTextInput input:focus { box-shadow: 0 0 0 3px rgba(245,166,35,.35) !important; border-color:#F5A623 !important; }

        /* ファイルアップローダ */
        [data-testid="stFileUploaderDropzone"] {
            background: #fbfcfe;
            border: 2px dashed #94a3b8;
            border-radius: 10px;
        }

        /* アラート */
        [data-testid="stAlert"] { border-radius: 10px; border: 1px solid #d6dce6; }

        /* expander: 白カード */
        [data-testid="stExpander"] {
            border: 1px solid #d6dce6;
            border-radius: 12px;
            background: #ffffff;
            box-shadow: 0 4px 12px rgba(31,42,55,.07);
        }
        [data-testid="stDataFrame"] { border: 1px solid #cbd5e1; border-radius: 8px; overflow: hidden; }

        hr { border-color: #d6dce6 !important; }

        /* ===== 倉庫の風景バナー(タイトル下) ===== */
        .wh-stage {
            position: relative; height: 140px; margin: .3rem 0 1.1rem;
            border: 1px solid #cfd6e2; border-radius: 12px; overflow: hidden;
            background: linear-gradient(#f8fafc, #e7ecf3);
            box-shadow: 0 6px 18px rgba(31,42,55,.08);
        }
        .wh-stage > * { position: absolute; }
        .wh-stage .floor  { left:0; right:0; bottom:0; height:26px; background:#c9d0db; border-top:2px solid #aab3c2; }
        .wh-stage .hazard { left:0; right:0; bottom:0; height:7px;
            background: repeating-linear-gradient(45deg, #FFC400 0 12px, #222 12px 24px); }

        /* スチール棚 */
        .wh-rack { bottom:26px; height:104px; width:230px; }
        .wh-rack .post  { position:absolute; top:0; width:7px; height:100%; border-radius:2px;
            background: linear-gradient(#5b6675,#3c4554); }
        .wh-rack .post.l { left:0; } .wh-rack .post.r { right:0; }
        .wh-rack .shelf { position:absolute; left:0; right:0; height:7px; border-radius:2px;
            background: linear-gradient(#6b7686,#49525f); }
        .wh-rack .shelf.s1 { top:0; } .wh-rack .shelf.s2 { top:48px; } .wh-rack .shelf.s3 { bottom:0; }
        .wh-box { position:absolute; background: linear-gradient(#d8a45f,#c2873f);
            border:1px solid #9c6a2b; border-radius:3px; }
        .wh-box::after { content:""; position:absolute; left:0; right:0; top:44%; height:3px; background:rgba(120,80,30,.45); }

        /* フォークリフト(走行) */
        .forklift { bottom:26px; left:-100px; width:80px; height:58px; animation: drive 12s linear infinite; }
        .forklift .cab  { position:absolute; bottom:12px; left:22px; width:40px; height:30px;
            background:#F5A623; border:2px solid #8a5e0c; border-radius:6px 10px 4px 4px; }
        .forklift .roof { position:absolute; bottom:34px; left:24px; width:34px; height:4px; background:#444; border-radius:2px; }
        .forklift .pillar { position:absolute; bottom:12px; right:18px; width:4px; height:24px; background:#444; }
        .forklift .mast { position:absolute; bottom:6px; left:12px; width:5px; height:46px; background:#3a3f47; }
        .forklift .fork { position:absolute; bottom:8px; left:-4px; width:18px; height:4px; background:#2b2f35; }
        .forklift .load { position:absolute; bottom:12px; left:-8px; width:22px; height:20px;
            background: linear-gradient(#d8a45f,#c2873f); border:2px solid #9c6a2b; border-radius:2px; }
        .forklift .w1, .forklift .w2 { position:absolute; bottom:0; width:16px; height:16px; border-radius:50%;
            background:#222; border:3px solid #6b7280; }
        .forklift .w1 { left:24px; } .forklift .w2 { left:44px; }
        @keyframes drive { 0% { left:-100px; } 100% { left:112%; } }

        @media (max-width: 900px){ .wh-rack { display:none; } }
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_theme_css()

# ---------------------------------------------------------------------------
# 背景画像
# ---------------------------------------------------------------------------
ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
BG_PATH = os.path.join(ASSETS_DIR, "background.png")


def apply_background(image_bytes: bytes, overlay: float):
    """画面に背景画像を適用する。overlay(0〜1)が大きいほど白く薄くなり文字が読みやすい。"""
    mime = "image/jpeg" if image_bytes[:2] == b"\xff\xd8" else "image/png"
    b64 = base64.b64encode(image_bytes).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: linear-gradient(
                rgba(238,241,245,{overlay}), rgba(238,241,245,{overlay})),
                url("data:{mime};base64,{b64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# タイトル
st.markdown(
    """
    <div style="display:flex; align-items:center; gap:12px; margin:.2rem 0 .3rem;">
      <span style="font-size:2rem;">📦</span>
      <span style="font-family:'M PLUS 1p',sans-serif; font-size:1.9rem; font-weight:800;
        color:#1f2a37; letter-spacing:.5px;">
        部品発注リスト 自動補完システム
      </span>
      <span style="background:#F5A623; color:#1f2a37; font-weight:700; font-size:.72rem;
        padding:3px 8px; border-radius:6px; border:1px solid #b97910;">LOGISTICS</span>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption("発注履歴を参照し、型式から単価・発注先・メーカー・仕様URLを自動入力します（Google Cloud不要）。")

# 倉庫の風景バナー(スチール棚＋段ボール箱＋走るフォークリフト)
st.markdown(
    """
    <div class="wh-stage">
      <div class="wh-rack" style="right:5%;">
        <div class="post l"></div><div class="post r"></div>
        <div class="shelf s1"></div><div class="shelf s2"></div><div class="shelf s3"></div>
        <div class="wh-box" style="left:16px;  top:10px;  width:46px; height:32px;"></div>
        <div class="wh-box" style="left:70px;  top:14px;  width:38px; height:28px;"></div>
        <div class="wh-box" style="left:120px; top:8px;   width:52px; height:34px;"></div>
        <div class="wh-box" style="left:22px;  top:58px;  width:40px; height:32px;"></div>
        <div class="wh-box" style="left:78px;  top:60px;  width:50px; height:30px;"></div>
        <div class="wh-box" style="left:140px; top:56px;  width:44px; height:34px;"></div>
      </div>
      <div class="wh-rack" style="left:6%;">
        <div class="post l"></div><div class="post r"></div>
        <div class="shelf s1"></div><div class="shelf s2"></div><div class="shelf s3"></div>
        <div class="wh-box" style="left:20px;  top:12px;  width:48px; height:30px;"></div>
        <div class="wh-box" style="left:90px;  top:9px;   width:50px; height:33px;"></div>
        <div class="wh-box" style="left:30px;  top:58px;  width:44px; height:32px;"></div>
        <div class="wh-box" style="left:110px; top:60px;  width:46px; height:30px;"></div>
      </div>
      <div class="forklift">
        <div class="load"></div><div class="fork"></div><div class="mast"></div>
        <div class="roof"></div><div class="pillar"></div><div class="cab"></div>
        <div class="w1"></div><div class="w2"></div>
      </div>
      <div class="floor"></div>
      <div class="hazard"></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 使い方ガイド(UI内)
# ---------------------------------------------------------------------------
with st.expander("📖 使い方ガイド（クリックで開閉）", expanded=True):
    st.markdown(
        """
        <div style="display:flex;gap:14px;flex-wrap:wrap;margin:.2rem 0 1rem;">
          <div style="flex:1;min-width:210px;background:#2c5282;
               border:1px solid #1f3a5f;border-radius:10px;padding:14px 16px;box-shadow:0 4px 12px rgba(31,42,55,.12);color:#fff;">
            <div style="font-size:1.4rem;font-weight:800;">①</div>
            <b style="color:#fff;">履歴データを指定</b><br>
            <span style="color:#dbe6f3;font-size:.9rem;">
            サイドバーで取得方式を選び、<b>Driveの共有フォルダURL</b>（または履歴ファイル）を入力。</span>
          </div>
          <div style="flex:1;min-width:210px;background:#F5A623;
               border:1px solid #b97910;border-radius:10px;padding:14px 16px;box-shadow:0 4px 12px rgba(31,42,55,.12);color:#1f2a37;">
            <div style="font-size:1.4rem;font-weight:800;">②</div>
            <b style="color:#1f2a37;">新規発注リストを用意</b><br>
            <span style="color:#4a3a12;font-size:.9rem;">
            下の「空テンプレート」をDLし、<b>型式</b>と<b>数量</b>を入力してアップロード。</span>
          </div>
          <div style="flex:1;min-width:210px;background:#475569;
               border:1px solid #2f3a49;border-radius:10px;padding:14px 16px;box-shadow:0 4px 12px rgba(31,42,55,.12);color:#fff;">
            <div style="font-size:1.4rem;font-weight:800;">③</div>
            <b style="color:#fff;">実行してDL</b><br>
            <span style="color:#dde3ea;font-size:.9rem;">
            <b>🚀 自動補完を実行</b>を押し、結果を確認して<b>💾 補完済みExcel</b>をダウンロード。</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
**取得方式の選び方**

- **Driveフォルダを走査（推奨）** — 「発注履歴見積書」フォルダを丸ごと読み込みます。
  - フォルダを *共有 → リンクを知っている全員 → 閲覧者* にして、**フォルダURLを貼り付け**ます。
  - フォルダ内は **メーカー別サブフォルダ**にしておくと、その**フォルダ名が「発注先」**になります。
  - 中の **Excel / CSV / PDF見積書** を自動で読み取ります（※スキャンした画像PDFは読めません）。
- **公開URLから取得** — 単一のGoogleスプレッドシート（型式・品名・発注先・単価…）のURLを貼ります。
- **ローカル履歴ファイル** — 手元の履歴 .xlsx / .csv を直接アップロードします。

**補完のしくみ**：新規発注リストの **A列「型式」** をキーに、
- **単価(D列)・発注先(E列)** … 発注履歴（見積/発注元）から自動入力
- **メーカー(B列)** … 「型式→メーカー対応表」から自動入力（サイドバーで指定）
- **合計金額(F列)** … 単価×数量 を計算
- **仕様URL(G列)** … 型式で検索するリンクを自動生成（Google/MISUMI/モノタロウから選択）

> 💡 困ったら：実行後に出る「読み込んだファイル／スキップ／エラー」の表示で原因を確認できます。
        """
    )

# ---------------------------------------------------------------------------
# サイドバー: 履歴データの取得元
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 履歴データの取得元")
    source_mode = st.radio(
        "取得方式",
        ["Driveフォルダを走査（推奨）", "公開URLから取得", "ローカル履歴ファイル"],
        help="Driveフォルダ方式は『Googleドライブ パソコン版』で同期されたフォルダを丸ごと読み込みます",
    )

    sheet_url = ""
    gid_override = ""
    hist_file = None
    folder_path = ""
    recursive = True

    if source_mode == "Driveフォルダを走査（推奨）":
        folder_path = st.text_input(
            "発注履歴見積書フォルダ（URL または ローカルパス）",
            placeholder="https://drive.google.com/drive/folders/... または G:\\マイドライブ\\発注履歴見積書",
            help="共有フォルダのURL、または『Googleドライブ パソコン版』で同期したローカルパスのどちらでも可。",
        )
        recursive = st.checkbox("サブフォルダも含める（ローカルパス時のみ）", value=True)
        st.caption("対応: .xlsx / .csv / Googleスプレッドシート / .pdf(見積書)")
        st.info("📌 URL指定の場合は、フォルダを『共有→リンクを知っている全員→閲覧者』にしてください。")
    elif source_mode == "公開URLから取得":
        sheet_url = st.text_input(
            "発注履歴シートのURL",
            placeholder="https://docs.google.com/spreadsheets/d/.../edit#gid=0",
            help="ブラウザのアドレスバーのURLをそのまま貼り付け。タブ(gid)も自動認識します。",
        )
        gid_override = st.text_input(
            "タブのgid（任意）", value="",
            help="特定タブを指定したい場合のみ。URLに#gid=...が含まれていれば空欄でOK。",
        )
        st.info("📌 シートは『共有 → リンクを知っている全員 → 閲覧者』に設定してください。")
    else:
        hist_file = st.file_uploader(
            "発注履歴ファイル (.xlsx / .csv)", type=["xlsx", "csv"], key="hist"
        )
        st.caption("列構成: A型式 B品名 C発注先 D単価 E備考")

    st.divider()
    st.subheader("🏭 型式→メーカー対応表")
    st.download_button(
        "📄 対応表テンプレートをDL",
        data=core.build_maker_template(),
        file_name="型式メーカー対応表_テンプレート.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    maker_file = st.file_uploader(
        "対応表 (.xlsx / .csv)", type=["xlsx", "csv"], key="maker",
        help="A列=型式, B列=メーカー の表。B列「メーカー」の自動入力に使います。",
    )
    st.caption("列構成: A型式 B メーカー（見出しがあれば名前でも自動判別）")

    st.divider()
    spec_engine_label = st.selectbox(
        "G列『仕様URL』の検索先",
        ["Google", "MISUMI", "モノタロウ"],
        help="型式で検索するリンクを自動生成します。",
    )
    fill_total = st.checkbox("合計金額(単価×数量)を計算する", value=True)

    # -----------------------------------------------------------------------
    # 表示設定: 背景画像
    # -----------------------------------------------------------------------
    st.divider()
    with st.expander("🎨 背景画像の設定"):
        bg_upload = st.file_uploader(
            "背景画像を選択 (png / jpg)", type=["png", "jpg", "jpeg"], key="bg"
        )
        overlay = st.slider(
            "背景の薄さ（大きいほど白く・文字が読みやすい）",
            min_value=0.0, max_value=1.0, value=0.6, step=0.05,
        )
        col_set, col_clear = st.columns(2)
        with col_set:
            if bg_upload is not None and st.button("背景を適用＆保存", use_container_width=True):
                os.makedirs(ASSETS_DIR, exist_ok=True)
                with open(BG_PATH, "wb") as f:
                    f.write(bg_upload.getvalue())
                st.success("背景を保存しました")
        with col_clear:
            if st.button("背景をクリア", use_container_width=True):
                if os.path.exists(BG_PATH):
                    os.remove(BG_PATH)
                st.success("背景を解除しました")

# 背景画像の適用(アップロード直後はそのプレビュー、以降は保存済み画像)
_bg_bytes = None
if bg_upload is not None:
    _bg_bytes = bg_upload.getvalue()
elif os.path.exists(BG_PATH):
    with open(BG_PATH, "rb") as f:
        _bg_bytes = f.read()
if _bg_bytes:
    apply_background(_bg_bytes, overlay)

# ---------------------------------------------------------------------------
# メイン: テンプレート配布 & 入力ファイルアップロード
# ---------------------------------------------------------------------------
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("① 新規発注リスト（入力）")
    st.download_button(
        "📄 空テンプレートをダウンロード",
        data=core.build_template(),
        file_name="新規発注リスト_テンプレート.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    uploaded = st.file_uploader("新規発注リスト (.xlsx) をアップロード", type=["xlsx"])

with col_b:
    st.subheader("② 列構成（目安）")
    st.markdown(
        "| 列 | 内容 | 入力 |\n"
        "|---|---|---|\n"
        "| A | 型式 | **手入力(キー)** |\n"
        "| B | メーカー | 自動(対応表) |\n"
        "| C | 数量 | **手入力** |\n"
        "| D | 単価 | 自動(履歴) |\n"
        "| E | 発注先 | 自動(履歴) |\n"
        "| F | 合計金額 | 自動 |\n"
        "| G | 仕様URL | 自動(検索リンク) |"
    )

st.divider()

# ---------------------------------------------------------------------------
# 実行
# ---------------------------------------------------------------------------
run = st.button("🚀 自動補完を実行", type="primary", use_container_width=True)

if run:
    # 入力チェック
    if source_mode == "Driveフォルダを走査（推奨）" and not folder_path.strip():
        st.error("発注履歴見積書フォルダのパスを入力してください。")
        st.stop()
    if source_mode == "公開URLから取得" and not sheet_url.strip():
        st.error("発注履歴シートのURLを入力してください。")
        st.stop()
    if source_mode == "ローカル履歴ファイル" and hist_file is None:
        st.error("発注履歴ファイル(.xlsx/.csv)をアップロードしてください。")
        st.stop()
    if uploaded is None:
        st.error("新規発注リスト(.xlsx)をアップロードしてください。")
        st.stop()

    try:
        scan_report = None
        with st.spinner("発注履歴を取得中..."):
            if source_mode == "Driveフォルダを走査（推奨）":
                import folder_loader
                target = folder_path.strip()
                if target.lower().startswith("http"):
                    history, scan_report = folder_loader.scan_drive_folder_url(target)
                else:
                    history, scan_report = folder_loader.scan_folder(
                        target, recursive=recursive
                    )
            elif source_mode == "公開URLから取得":
                gid = gid_override.strip() or None
                history = core.load_history_from_csv_url(sheet_url.strip(), gid=gid)
            else:
                if hist_file.name.lower().endswith(".csv"):
                    # CSVはテキストとして読み込む
                    import io, csv
                    text = hist_file.getvalue().decode("utf-8-sig", errors="replace")
                    rows = list(csv.reader(io.StringIO(text)))
                    history = core._build_history_from_rows(rows)
                else:
                    history = core.load_history_from_file(hist_file.getvalue())

        # フォルダ走査のレポート表示
        if scan_report is not None:
            st.info(
                f"フォルダ走査: 表 {scan_report.tables_loaded} / "
                f".gsheet {scan_report.gsheets_loaded} / PDF {scan_report.pdfs_loaded} ファイル、"
                f"履歴 {scan_report.records_total} 件(うちPDF由来 {scan_report.pdf_items} 件)"
            )
            if scan_report.sources:
                with st.expander("読み込んだファイル"):
                    for s in scan_report.sources:
                        st.write("•", s)
            if scan_report.skipped:
                with st.expander(f"⚠️ スキップ {len(scan_report.skipped)} 件"):
                    for n, why in scan_report.skipped:
                        st.write(f"- {n}: {why}")
            if scan_report.errors:
                with st.expander(f"❌ エラー {len(scan_report.errors)} 件"):
                    for n, why in scan_report.errors:
                        st.write(f"- {n}: {why}")
        else:
            st.info(f"発注履歴を {len(history)} 件読み込みました。")

        if not history:
            st.warning("履歴データが0件でした。フォルダパス・列構成・共有設定を確認してください。")
            st.stop()

        # 型式→メーカー対応表(任意)
        maker_map = {}
        if maker_file is not None:
            if maker_file.name.lower().endswith(".csv"):
                import io as _io, csv as _csv
                _text = maker_file.getvalue().decode("utf-8-sig", errors="replace")
                maker_map = core.load_maker_map_from_rows(list(_csv.reader(_io.StringIO(_text))))
            else:
                maker_map = core.load_maker_map_from_file(maker_file.getvalue())
            st.info(f"型式→メーカー対応表を {len(maker_map)} 件読み込みました。")

        _engine = {"Google": "google", "MISUMI": "misumi", "モノタロウ": "monotaro"}.get(
            spec_engine_label, "google"
        )

        with st.spinner("照合・補完中..."):
            output_bytes, result = core.process_workbook(
                uploaded.getvalue(), history,
                maker_map=maker_map, fill_total=fill_total, spec_engine=_engine,
            )

        # サマリ
        st.success("補完が完了しました 🎉")
        m1, m2, m3 = st.columns(3)
        m1.metric("対象行数", result.total_rows)
        m2.metric("ヒット(補完)", result.matched)
        m3.metric("新規(未ヒット)", result.new_items)

        if result.new_models:
            with st.expander(f"⚠️ 履歴に無い型式 {len(result.new_models)} 件(新規)"):
                st.write(", ".join(map(str, result.new_models)))

        # プレビュー
        if result.preview:
            st.subheader("プレビュー")
            st.dataframe(pd.DataFrame(result.preview), use_container_width=True, hide_index=True)

        # ダウンロード
        base = os.path.splitext(uploaded.name)[0]
        st.download_button(
            "💾 補完済み Excel をダウンロード",
            data=output_bytes,
            file_name=f"{base}_補完済み.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    except PermissionError as e:
        st.error(str(e))
    except Exception as e:  # noqa: BLE001  画面に分かりやすく出すため広く捕捉
        st.error(f"エラーが発生しました: {e}")
        with st.expander("詳細(トラブルシュート用)"):
            st.code(traceback.format_exc())
        st.markdown(
            "**よくある原因**\n"
            "- シートが『リンクを知っている全員が閲覧可』になっていない\n"
            "- URL / タブ(gid) の誤り\n"
            "- 履歴の列構成(A型式 B品名 C発注先 D単価 E備考)が違う"
        )
