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
# モダンなダークテーマ(カスタムCSS)
# ---------------------------------------------------------------------------
def inject_theme_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Poppins:wght@600;800&display=swap');

        html, body, [class*="css"], .stApp, [data-testid="stSidebar"] {
            font-family: 'Inter', 'Segoe UI', sans-serif;
        }

        /* 背景: 深い宇宙的グラデーション */
        .stApp {
            background: radial-gradient(1200px 600px at 15% -10%, #243049 0%, rgba(15,23,42,0) 60%),
                        radial-gradient(1000px 500px at 100% 0%, #2a1e4a 0%, rgba(15,23,42,0) 55%),
                        linear-gradient(160deg, #0b1220 0%, #0f172a 50%, #060a14 100%);
            background-attachment: fixed;
            color: #e2e8f0;
        }

        /* 文字色(config.toml が無くても読めるよう明示) */
        .stApp, .stMarkdown, .stMarkdown p, .stMarkdown li, .stApp p, .stApp li,
        label, [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p,
        .stRadio label, .stCheckbox label, .stSlider label,
        [data-testid="stExpander"] summary, [data-testid="stMetricLabel"],
        [data-testid="stMarkdownContainer"] {
            color: #e2e8f0 !important;
        }
        [data-testid="stCaptionContainer"], .stCaption, small { color: #94a3b8 !important; }
        /* 上部ヘッダーバーを透過 */
        [data-testid="stHeader"] { background: transparent; }

        /* タイトルをグラデーション文字に */
        h1 {
            font-family: 'Poppins', sans-serif !important;
            font-weight: 800 !important;
            letter-spacing: .5px;
            background: linear-gradient(90deg, #a78bfa 0%, #60a5fa 45%, #22d3ee 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            text-shadow: 0 0 30px rgba(139,92,246,.25);
        }
        h2, h3 { font-family: 'Poppins', sans-serif !important; font-weight: 600 !important; color:#e9edf5; }

        /* メインコンテナを少し中央寄せに */
        .block-container { padding-top: 2.2rem; max-width: 1200px; }

        /* サイドバー: すりガラス */
        [data-testid="stSidebar"] > div:first-child {
            background: rgba(17, 25, 40, 0.75);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border-right: 1px solid rgba(148,163,184,0.12);
        }

        /* 通常ボタン */
        .stButton > button, .stDownloadButton > button {
            border-radius: 12px;
            border: 1px solid rgba(148,163,184,0.25);
            background: rgba(30,41,59,0.55);
            color: #e2e8f0;
            font-weight: 600;
            padding: .5rem 1rem;
            transition: all .18s ease;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            transform: translateY(-2px);
            border-color: #a78bfa;
            box-shadow: 0 8px 22px rgba(124,58,237,.32);
            color: #fff;
        }
        /* 主要ボタン(実行/DL): グラデーション */
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
            background: linear-gradient(95deg, #7c3aed 0%, #4f46e5 50%, #2563eb 100%);
            border: none; color: #fff;
            box-shadow: 0 6px 18px rgba(79,70,229,.35);
        }
        .stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 28px rgba(99,102,241,.55);
        }

        /* メトリクスをカード化 */
        [data-testid="stMetric"] {
            background: linear-gradient(180deg, rgba(40,52,78,0.6), rgba(24,32,52,0.6));
            border: 1px solid rgba(148,163,184,0.18);
            border-radius: 16px;
            padding: 16px 18px;
            backdrop-filter: blur(6px);
            box-shadow: 0 8px 24px rgba(2,6,23,.35);
        }
        [data-testid="stMetricValue"] { color: #c4b5fd; font-weight: 700; }

        /* 入力欄 */
        .stTextInput input, .stNumberInput input, .stTextArea textarea {
            border-radius: 10px !important;
            background: rgba(15,23,42,0.6) !important;
            border: 1px solid rgba(148,163,184,0.22) !important;
            color: #e2e8f0 !important;
        }
        .stTextInput input:focus { border-color:#8b5cf6 !important; box-shadow:0 0 0 2px rgba(139,92,246,.25)!important; }

        /* ファイルアップローダ */
        [data-testid="stFileUploaderDropzone"] {
            background: rgba(30,41,59,0.4);
            border: 1.5px dashed rgba(139,92,246,0.45);
            border-radius: 14px;
        }

        /* アラート(info/success/warning) 角丸 */
        [data-testid="stAlert"] { border-radius: 12px; }

        /* expander / dataframe を角丸 */
        [data-testid="stExpander"] {
            border: 1px solid rgba(148,163,184,0.15);
            border-radius: 14px;
            background: rgba(24,32,52,0.45);
        }
        [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }

        /* タブ・ラジオのアクセント */
        .stRadio [data-baseweb="radio"] div[aria-checked="true"] { border-color:#8b5cf6 !important; }

        /* 区切り線を淡く */
        hr { border-color: rgba(148,163,184,0.15) !important; }
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
    """画面に背景画像を適用する。overlay(0〜1)が大きいほど暗くなり文字が読みやすい(ダーク基調)。"""
    mime = "image/jpeg" if image_bytes[:2] == b"\xff\xd8" else "image/png"
    b64 = base64.b64encode(image_bytes).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: linear-gradient(
                rgba(8,12,22,{overlay}), rgba(8,12,22,{overlay})),
                url("data:{mime};base64,{b64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


st.markdown(
    """
    <div style="margin-bottom:.4rem;">
      <span style="font-family:'Poppins',sans-serif;font-size:2.1rem;font-weight:800;
        background:linear-gradient(90deg,#a78bfa,#60a5fa,#22d3ee);
        -webkit-background-clip:text;-webkit-text-fill-color:transparent;">
        📦 部品発注リスト 自動補完システム
      </span>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption("発注履歴を参照し、型式から品名・発注先・単価を自動入力します（Google Cloud不要）。")

# ---------------------------------------------------------------------------
# 使い方ガイド(UI内)
# ---------------------------------------------------------------------------
with st.expander("📖 使い方ガイド（クリックで開閉）", expanded=True):
    st.markdown(
        """
        <div style="display:flex;gap:14px;flex-wrap:wrap;margin:.2rem 0 1rem;">
          <div style="flex:1;min-width:210px;background:linear-gradient(180deg,rgba(124,58,237,.18),rgba(37,99,235,.10));
               border:1px solid rgba(148,163,184,.18);border-radius:16px;padding:14px 16px;">
            <div style="font-size:1.4rem;">①</div>
            <b style="color:#c4b5fd;">履歴データを指定</b><br>
            <span style="color:#cbd5e1;font-size:.9rem;">
            サイドバーで取得方式を選び、<b>Driveの共有フォルダURL</b>（または履歴ファイル）を入力。</span>
          </div>
          <div style="flex:1;min-width:210px;background:linear-gradient(180deg,rgba(124,58,237,.18),rgba(37,99,235,.10));
               border:1px solid rgba(148,163,184,.18);border-radius:16px;padding:14px 16px;">
            <div style="font-size:1.4rem;">②</div>
            <b style="color:#c4b5fd;">新規発注リストを用意</b><br>
            <span style="color:#cbd5e1;font-size:.9rem;">
            下の「空テンプレート」をDLし、<b>型式</b>と<b>数量</b>を入力してアップロード。</span>
          </div>
          <div style="flex:1;min-width:210px;background:linear-gradient(180deg,rgba(124,58,237,.18),rgba(37,99,235,.10));
               border:1px solid rgba(148,163,184,.18);border-radius:16px;padding:14px 16px;">
            <div style="font-size:1.4rem;">③</div>
            <b style="color:#c4b5fd;">実行してDL</b><br>
            <span style="color:#cbd5e1;font-size:.9rem;">
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

**補完のしくみ**：新規発注リストの **A列「型式」** をキーに履歴を検索し、
**品名・発注先・単価** を自動入力。**合計金額＝単価×数量** を計算します。
履歴に無い型式は空欄のまま **状態列に「新規」**（黄色）を立てます。

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
            "背景の暗さ（大きいほど暗く・文字が読みやすい）",
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
        "| B | 品名 | 自動 |\n"
        "| C | 発注先 | 自動 |\n"
        "| D | 単価 | 自動 |\n"
        "| E | 数量 | **手入力** |\n"
        "| F | 合計金額 | 自動 |\n"
        "| G | 状態 | 自動(OK/新規) |"
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

        with st.spinner("照合・補完中..."):
            output_bytes, result = core.process_workbook(
                uploaded.getvalue(), history, fill_total=fill_total
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
