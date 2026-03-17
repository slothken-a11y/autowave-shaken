"""
車検満期管理・進捗可視化システム
===================================
セッション完結型：ブラウザを閉じればデータは消去されます。
クラウドDB不使用 / ローカルCSVアップロード専用。
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from io import BytesIO
import calendar
import os
from pathlib import Path

# ──────────────────────────────────────────────
# ページ設定
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="車検満期管理システム",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# パスワード認証
# ──────────────────────────────────────────────
def check_password():
    """パスワード認証画面を表示し、認証済みかどうかを返す"""
    # Streamlit Cloudのsecretsからパスワードを取得
    # ローカル実行時はパスワードなしでスルー
    try:
        correct_password = st.secrets["password"]
    except Exception:
        return True  # secrets未設定（ローカル環境）はスルー

    # 認証済みセッション
    if st.session_state.get("authenticated"):
        return True

    # ── ログイン画面 ──
    st.markdown("""
    <div style="
        max-width:400px; margin:80px auto; padding:2rem;
        background:#fff; border-radius:12px;
        box-shadow:0 4px 20px rgba(0,0,0,0.1);
        text-align:center;
    ">
        <div style="font-size:3rem;">🚗</div>
        <h2 style="color:#1a237e; margin:0.5rem 0;">車検満期管理システム</h2>
        <p style="color:#666; font-size:0.9rem;">株式会社オートウェーブ</p>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        st.markdown("<div style='max-width:400px;margin:0 auto;'>", unsafe_allow_html=True)
        password = st.text_input("パスワード", type="password", placeholder="パスワードを入力してください")
        submitted = st.form_submit_button("ログイン", use_container_width=True, type="primary")
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if password == correct_password:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("パスワードが違います。")

    return False

if not check_password():
    st.stop()

# ──────────────────────────────────────────────
# カスタムCSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
    /* メインヘッダー */
    .main-header {
        background: linear-gradient(135deg, #1a237e 0%, #283593 100%);
        color: white;
        padding: 1.2rem 1.5rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; letter-spacing: 0.05em; }
    .main-header p  { margin: 0.3rem 0 0 0; font-size: 0.85rem; opacity: 0.85; }

    /* KPIカード */
    .kpi-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 4px solid #1a237e;
        text-align: center;
    }
    .kpi-card .label { font-size: 0.78rem; color: #666; margin-bottom: 0.3rem; }
    .kpi-card .value { font-size: 1.8rem; font-weight: 700; color: #1a237e; }
    .kpi-card .sub   { font-size: 0.75rem; color: #999; margin-top: 0.2rem; }

    /* ステータスバッジ */
    .badge-pass    { background: #e8f5e9; color: #2e7d32; padding: 3px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; }
    .badge-warn    { background: #fff3e0; color: #e65100; padding: 3px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; }
    .badge-danger  { background: #ffebee; color: #c62828; padding: 3px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; }

    /* セクション */
    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1a237e;
        border-bottom: 2px solid #1a237e;
        padding-bottom: 0.4rem;
        margin: 1.5rem 0 1rem 0;
    }

    /* セキュリティバナー */
    .security-banner {
        background: #e8f5e9;
        border: 1px solid #a5d6a7;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        font-size: 0.8rem;
        color: #2e7d32;
        text-align: center;
        margin-bottom: 1rem;
    }

    /* アップロードエリア */
    .upload-status {
        padding: 0.5rem;
        border-radius: 6px;
        font-size: 0.85rem;
        margin: 0.3rem 0;
    }
    .upload-ok   { background: #e8f5e9; color: #2e7d32; }
    .upload-wait { background: #fff3e0; color: #e65100; }

    div[data-testid="stMetricValue"] { font-size: 1.5rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background: #f5f5f5;
        border-radius: 8px 8px 0 0;
        padding: 8px 20px;
    }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# ヘッダー
# ──────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🚗 車検満期管理・進捗可視化システム</h1>
    <p>セッション完結型 ─ データはブラウザ終了時に自動消去されます</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="security-banner">
    🔒 セキュリティ：アップロードされたデータはサーバーメモリ上のみで処理され、永続化されません。
</div>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# ユーティリティ関数
# ──────────────────────────────────────────────
def to_csv_bytes(df: pd.DataFrame) -> bytes:
    """DataFrameをCSVバイト列に変換（ダウンロード用）"""
    buf = BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
    return buf.getvalue()


def parse_date_col(series: pd.Series) -> pd.Series:
    """日付列を柔軟にパース"""
    return pd.to_datetime(series, errors="coerce", infer_datetime_format=True)


def progress_color(rate: float, target: float = 0.70) -> str:
    """進捗率に応じた色クラスを返す"""
    if rate >= target:
        return "badge-pass"
    elif rate >= target * 0.6:
        return "badge-warn"
    else:
        return "badge-danger"


def progress_label(rate: float, target: float = 0.70) -> str:
    if rate >= target:
        return "✅ 合格"
    elif rate >= target * 0.6:
        return "⚠️ 要注意"
    else:
        return "🚨 警告"


# ──────────────────────────────────────────────
# サイドバー：CSVアップロード
# ──────────────────────────────────────────────
# ──────────────────────────────────────────────
# GoogleドライブファイルID定義
# ──────────────────────────────────────────────
GDRIVE_IDS = {
    "master"     : "147knWIQ1AP09lnPINo1xl0TzDxWDfmDI",
    "history"    : "17fbI7_5OOf6tSGq4qntL7BpOF-dj3sFq",
    "reservation": "1oqPp9cRyPGfkFEeXVpBgtdJJWs-uirtZ",
    "kikan"      : None,  # Excelは手動アップロードのみ
}

def gdrive_url(file_id: str) -> str:
    """GoogleドライブダウンロードURLを生成"""
    return f"https://drive.google.com/uc?export=download&id={file_id}"

@st.cache_data(ttl=300, show_spinner=False)  # 5分キャッシュ
def load_from_gdrive(file_id: str, name: str) -> pd.DataFrame:
    """GoogleドライブからCSVを読み込む"""
    url = gdrive_url(file_id)
    for enc in ["cp932", "utf-8-sig", "utf-8", "shift_jis"]:
        try:
            return pd.read_csv(url, encoding=enc)
        except Exception:
            continue
    raise ValueError(f"{name} をGoogleドライブから読み込めませんでした。")

# ──────────────────────────────────────────────
# /data フォルダパス定義（ローカル起動用）
# ──────────────────────────────────────────────
DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

DATA_FILES = {
    "master"     : DATA_DIR / "Master_Data.csv",
    "history"    : DATA_DIR / "Past_History.csv",
    "reservation": DATA_DIR / "Reservation_Data.csv",
    "kikan"      : DATA_DIR / "基幹予約データ.xlsx",
}

def file_mtime(path: Path) -> str:
    """ファイルの最終更新日時を文字列で返す（存在しない場合は空文字）"""
    if path.exists():
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y/%m/%d %H:%M:%S")
    return ""

# ──────────────────────────────────────────────
# サイドバー
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 データ読み込み")

    # ── GoogleドライブCSV状況 ──
    st.markdown("#### ☁️ Googleドライブ（クラウド自動読み込み）")
    gdrive_labels = {
        "master"     : "Master_Data.csv",
        "history"    : "Past_History.csv",
        "reservation": "Reservation_Data.csv",
    }
    for key, label in gdrive_labels.items():
        fid = GDRIVE_IDS.get(key)
        if fid:
            st.markdown(
                f'<div class="upload-status upload-ok">☁️ {label}（Googleドライブ）</div>',
                unsafe_allow_html=True,
            )

    # ── 更新ボタン ──
    st.markdown("")
    if st.button("🔄 データを再読み込み", use_container_width=True, type="primary"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")

    # ── /data フォルダの自動読み込み状況（ローカル用）──
    st.markdown("#### 📁 /data フォルダ（ローカル起動時）")
    auto_status = {}
    for key, path in DATA_FILES.items():
        exists = path.exists()
        auto_status[key] = exists
        label = path.name
        mtime = file_mtime(path)
        opt   = "（任意）" if key == "kikan" else ""
        if exists:
            st.markdown(
                f'<div class="upload-status upload-ok">✅ {label}<br>'
                f'<span style="font-size:0.75rem;opacity:0.8;">{mtime}</span></div>',
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # ── 手動アップロード（フォールバック）──
    st.markdown("#### 📤 手動アップロード（フォールバック）")
    st.caption("GoogleドライブやローカルにCSVがない場合のみ使用")

    f_master      = st.file_uploader("① Master_Data.csv",      type=["csv"],        key="master",      help="Googleドライブが優先されます")
    f_history     = st.file_uploader("② Past_History.csv",     type=["csv"],        key="history",     help="Googleドライブが優先されます")
    f_reservation = st.file_uploader("③ Reservation_Data.csv", type=["csv"],        key="reservation", help="Googleドライブが優先されます")
    f_kikan       = st.file_uploader("④ 基幹予約データ.xlsx（任意）", type=["xlsx","xls"], key="kikan",       help="手動アップロードのみ対応")

    st.markdown("---")
    st.markdown("### ⚙️ 分析期間設定")
    today = datetime.today()

    st.caption("開始")
    col_sy, col_sm = st.columns(2)
    with col_sy:
        start_year = st.selectbox("開始年", range(today.year - 1, today.year + 3), index=1)
    with col_sm:
        start_month = st.selectbox("開始月", range(1, 13), index=today.month - 1, format_func=lambda x: f"{x}月")

    st.caption("終了")
    col_ey, col_em = st.columns(2)
    with col_ey:
        end_year = st.selectbox("終了年", range(today.year - 1, today.year + 3), index=2, key="end_y")
    with col_em:
        end_month = st.selectbox("終了月", range(1, 13), index=2, format_func=lambda x: f"{x}月", key="end_m")

    analysis_start = datetime(start_year, start_month, 1)
    analysis_end = datetime(end_year, end_month, calendar.monthrange(end_year, end_month)[1])

    if analysis_end < analysis_start:
        st.error("⚠️ 終了月が開始月より前です。期間を修正してください。")

    st.info(f"分析期間：{analysis_start.strftime('%Y/%m')}〜{analysis_end.strftime('%Y/%m')}")


# ──────────────────────────────────────────────
# データ読み込みと処理
# ──────────────────────────────────────────────

# /data ファイルを優先、なければ手動アップロードを使用
def resolve_source(data_key: str, uploaded_file):
    """
    優先順位：
    1. 手動アップロード（最優先・最新データ）
    2. Googleドライブ（クラウド自動）
    3. /data フォルダ（ローカル起動時）
    """
    # 1. 手動アップロード優先
    if uploaded_file is not None:
        return ("upload", uploaded_file)
    # 2. Googleドライブ
    fid = GDRIVE_IDS.get(data_key)
    if fid:
        return ("gdrive", fid)
    # 3. /dataフォルダ
    path = DATA_FILES[data_key]
    if path.exists():
        return ("file", str(path))
    return None

src_master      = resolve_source("master",      f_master)
src_history     = resolve_source("history",     f_history)
src_reservation = resolve_source("reservation", f_reservation)
src_kikan       = resolve_source("kikan",       f_kikan)

# 必須3ファイルが揃っていない場合
if not all([src_master, src_history, src_reservation]):
    st.markdown("---")
    missing = []
    if not src_master:      missing.append("Master_Data.csv")
    if not src_history:     missing.append("Past_History.csv")
    if not src_reservation: missing.append("Reservation_Data.csv")
    st.warning(
        f"以下のファイルが見つかりません：**{'、'.join(missing)}**\n\n"
        f"Googleドライブのリンクを確認するか、サイドバーから手動アップロードしてください。"
    )

    # サンプルCSVの生成・ダウンロードセクション
    with st.expander("📋 サンプルCSVをダウンロード（テスト用）"):
        st.caption("動作確認用のダミーデータをダウンロードできます。")

        # --- サンプルデータ生成 ---
        np.random.seed(42)
        n_records = 200
        stores = ["S001", "S002", "S003", "S004", "S005"]
        base_date = datetime(2026, 4, 1)

        sample_master = pd.DataFrame({
            "顧客ID": [f"C{str(i).zfill(5)}" for i in range(1, n_records + 1)],
            "車両ID": [f"V{str(i).zfill(5)}" for i in range(1, n_records + 1)],
            "登録番号": [f"品川{np.random.randint(100,999)}-{np.random.choice(['あ','い','う','え','お'])}-{str(np.random.randint(1,9999)).zfill(4)}" for i in range(n_records)],
            "車検満了日": [(base_date + timedelta(days=np.random.randint(0, 365))).strftime("%Y/%m/%d") for _ in range(n_records)],
            "入庫店舗ID": [np.random.choice(stores) for _ in range(n_records)],
        })

        # 70%をリピート対象にする
        repeat_ids = sample_master["車両ID"].sample(frac=0.70, random_state=42).tolist()
        sample_history = pd.DataFrame({
            "車両ID": repeat_ids,
            "前回車検実施日": [(datetime(2024, 4, 1) + timedelta(days=np.random.randint(0, 365))).strftime("%Y/%m/%d") for _ in repeat_ids],
        })

        # 予約状況（40%が本予約、20%が仮予約）
        reserved_nums = sample_master["登録番号"].sample(frac=0.60, random_state=42).tolist()
        statuses = (["本予約"] * int(len(reserved_nums) * 0.67)) + (["仮予約"] * (len(reserved_nums) - int(len(reserved_nums) * 0.67)))
        np.random.shuffle(statuses)
        sample_reservation = pd.DataFrame({
            "登録番号": reserved_nums,
            "予約ステータス": statuses[:len(reserved_nums)],
            "最終更新日": [(datetime(2026, 1, 1) + timedelta(days=np.random.randint(0, 90))).strftime("%Y/%m/%d") for _ in reserved_nums],
        })

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("Master_Data.csv", to_csv_bytes(sample_master), "Master_Data.csv", "text/csv")
        with c2:
            st.download_button("Past_History.csv", to_csv_bytes(sample_history), "Past_History.csv", "text/csv")
        with c3:
            st.download_button("Reservation_Data.csv", to_csv_bytes(sample_reservation), "Reservation_Data.csv", "text/csv")

    st.stop()


# ── CSVをDataFrameに読み込み（エンコーディング自動判定）──
@st.cache_data(show_spinner=False)
def read_csv_auto_path(path: str, name: str) -> pd.DataFrame:
    """ファイルパスからCSVを読み込む（/dataフォルダ用・キャッシュあり）"""
    for enc in ["cp932", "utf-8-sig", "utf-8", "shift_jis", "latin1"]:
        try:
            return pd.read_csv(path, encoding=enc)
        except (UnicodeDecodeError, Exception):
            continue
    raise ValueError(f"{name} を読み込めませんでした。エンコーディングを確認してください。")

def read_csv_auto(uploaded_file, name: str) -> pd.DataFrame:
    """アップロードファイルオブジェクトからCSVを読み込む"""
    for enc in ["cp932", "utf-8-sig", "utf-8", "shift_jis", "latin1"]:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc)
        except (UnicodeDecodeError, Exception):
            continue
    raise ValueError(f"{name} を読み込めませんでした。エンコーディングを確認してください。")

def load_df(src, name: str) -> pd.DataFrame:
    """resolve_source の結果に応じてDataFrameを返す"""
    kind, obj = src
    if kind == "gdrive":
        return load_from_gdrive(obj, name)
    elif kind == "file":
        return read_csv_auto_path(obj, name)
    else:
        return read_csv_auto(obj, name)

try:
    df_master      = load_df(src_master,      "Master_Data")
    df_history     = load_df(src_history,     "Past_History")
    df_reservation = load_df(src_reservation, "Reservation_Data")
except Exception as e:
    st.error(f"CSV読み込みエラー: {e}")
    st.stop()

# ── カラム名の正規化（全角スペース・前後空白を除去）──
for df in [df_master, df_history, df_reservation]:
    df.columns = df.columns.str.strip().str.replace("\u3000", "")

# ── Master_Data 列名マッピング（元データ列名 → システム内部列名）──
MASTER_COL_MAP = {
    "SYARYO_ID"    : "車両ID",
    "TOROKUBANGO"  : "登録番号",
    "MANRYOBI"     : "車検満了日",
    "SYARYO_KYOTEN": "入庫店舗ID",
    "KYOTEN_ID"    : "拠点ID",
    "SHAMEI"       : "車名",
    "TSUSHOMEI"    : "通称名",
    "KATASHIKI"    : "型式",
    "TOROKUBI"     : "登録日",
    "SHONENDO"     : "初年度",
}
# 元データ列名が含まれている場合のみリネーム（既に日本語列名の場合はスキップ）
if "SYARYO_ID" in df_master.columns:
    df_master = df_master.rename(columns=MASTER_COL_MAP)
    # 登録番号のスペースを除去（予約データとの結合キー）
    df_master["登録番号"] = df_master["登録番号"].astype(str).str.replace(" ", "", regex=False)
    # 顧客IDが元データにない場合は車両IDで代用
    if "顧客ID" not in df_master.columns:
        df_master["顧客ID"] = df_master["車両ID"]

# ── 日付変換 ──
df_master["車検満了日"] = parse_date_col(df_master["車検満了日"])
df_history["前回車検実施日"] = parse_date_col(df_history["前回車検実施日"])
df_reservation["最終更新日"] = parse_date_col(df_reservation["最終更新日"])

# ── 車検満了日 異常値クレンジング ──
# 現実的な範囲（2020年〜2035年）以外はNULLに変換して除外
MANRYOBI_MIN = pd.Timestamp("2020-01-01")
MANRYOBI_MAX = pd.Timestamp("2035-12-31")
abnormal_mask = (
    df_master["車検満了日"].notna()
    & ((df_master["車検満了日"] < MANRYOBI_MIN) | (df_master["車検満了日"] > MANRYOBI_MAX))
)
n_abnormal = abnormal_mask.sum()
if n_abnormal > 0:
    st.warning(
        f"⚠️ 車検満了日が異常値（2020年未満または2035年超）のレコードが **{n_abnormal:,}件** あります。"
        f"これらはデータ入力ミスとして分析対象から除外します。"
    )
    df_master.loc[abnormal_mask, "車検満了日"] = pd.NaT

# ──────────────────────────────────────────────
# ロジック1: リピート/新規の識別
# 優先①：Master に「リピート」列（〇フラグ）があればそれを使用
# 優先②：ない場合は Past_History の車両IDマッチで判定
# ──────────────────────────────────────────────
if "リピート" in df_master.columns:
    # K列〇フラグ優先（元データのVLOOKUP結果をそのまま使用）
    df_master["顧客種別"] = df_master["リピート"].apply(
        lambda x: "リピート" if str(x).strip() == "〇" else "新規"
    )
    st.info("ℹ️ リピート判定：Master_Dataの「リピート」列（〇フラグ）を使用しています。")
else:
    # フォールバック：車両IDマッチ（int64統一でマッチング漏れ防止）
    df_master["車両ID_key"] = pd.to_numeric(df_master["車両ID"], errors="coerce").astype("Int64")
    df_history["車両ID_key"] = pd.to_numeric(df_history["車両ID"], errors="coerce").astype("Int64")
    repeat_vehicle_ids = set(df_history["車両ID_key"].dropna().unique())
    df_master["顧客種別"] = df_master["車両ID_key"].apply(
        lambda x: "リピート" if x in repeat_vehicle_ids else "新規"
    )
    st.info("ℹ️ リピート判定：Past_Historyの車両IDマッチで判定しています（リピート列なし）。")

# ──────────────────────────────────────────────
# ロジック2: 流出判定
# ──────────────────────────────────────────────
df_master["流出済"] = df_master["車検満了日"].apply(
    lambda d: True if pd.notna(d) and d > analysis_end else False
)

# ──────────────────────────────────────────────
# ロジック3: 分析期間内のターゲット抽出
# ──────────────────────────────────────────────
mask_period = (
    (df_master["車検満了日"] >= analysis_start)
    & (df_master["車検満了日"] <= analysis_end)
)
df_target = df_master[mask_period].copy()

# ──────────────────────────────────────────────
# ロジック4: 予約状況の紐付け
# ★重要：予約CSVはステータス付与のみ。分母（行数）は絶対に変えない
#   → 車両IDベースで本予約セットを構築し、mapで1対1付与
# ──────────────────────────────────────────────

# 店舗名寄せマップ（基幹システム → Master店舗名）
KIKAN_STORE_MAP = {
    "車検　上総君津店" : "ｵｰﾄｳｪｰﾌﾞ上総君津店",
    "車検　冨里店"    : "ｵｰﾄｳｪｰﾌﾞ富里店",
    "車検　浜野店"    : "ｵｰﾄｳｪｰﾌﾞ浜野店",
    "車検　柏沼南店"  : "ｵｰﾄｳｪｰﾌﾞ柏沼南店",
    "車検　千種"     : "ｵｰﾄｳｪｰﾌﾞ千種店",
    "車検　茂原店"    : "ｵｰﾄｳｪｰﾌﾞ茂原店",
    "車検　宮野木"    : "ｵｰﾄｳｴｰﾌﾞ宮野木店",
}

# ── 予約CSV：登録番号 → 車両ID に変換して本予約IDセットを構築 ──
df_reservation["登録番号"] = df_reservation["登録番号"].astype(str).str.replace(" ", "", regex=False)
df_reservation_dedup = (
    df_reservation
    .sort_values("最終更新日", ascending=False)
    .drop_duplicates(subset="登録番号", keep="first")
)
plate_to_vid = df_master.set_index("登録番号")["車両ID_key"].to_dict()     if "車両ID_key" in df_master.columns     else df_master.assign(車両ID_key=pd.to_numeric(df_master["車両ID"], errors="coerce").astype("Int64"))          .set_index("登録番号")["車両ID_key"].to_dict()

res_booked_vids = set()
res_plate_map   = {}
for _, row in df_reservation_dedup.iterrows():
    plate = row["登録番号"]
    vid   = plate_to_vid.get(plate)
    if pd.notna(vid):
        res_booked_vids.add(vid)
    res_plate_map[plate] = row["予約ステータス"]

# ── 基幹システムExcel（任意）：車両IDで本予約IDセットを追加 ──
kikan_booked_vids = set()
if src_kikan is not None:
    try:
        _kikan_kind, _kikan_obj = src_kikan
        df_kikan = pd.read_excel(_kikan_obj)
        df_kikan.columns = df_kikan.columns.str.strip().str.replace("　", "")
        # A列=車両ID, B列=店舗（予約項目）, G列=登録番号
        vid_col   = df_kikan.columns[0]   # A列
        store_col = df_kikan.columns[1]   # B列
        plate_col = df_kikan.columns[6]   # G列
        df_kikan["車両ID_key"] = pd.to_numeric(df_kikan[vid_col], errors="coerce").astype("Int64")
        df_kikan["店舗名_master"] = df_kikan[store_col].map(KIKAN_STORE_MAP)
        valid_kikan = df_kikan[df_kikan["車両ID_key"].notna() & df_kikan["店舗名_master"].notna()]
        kikan_booked_vids = set(valid_kikan["車両ID_key"].unique())
        st.info(f"ℹ️ 基幹システム予約データ：{len(df_kikan):,}件読込 → 有効{len(valid_kikan):,}件（車両ID {len(kikan_booked_vids):,}件）を予約済みとして統合")
    except Exception as e:
        st.warning(f"基幹システムデータの読込に失敗しました: {e}")

# ── 統合：予約CSV ∪ 基幹Excel の車両IDセット ──
combined_booked_vids = res_booked_vids | kikan_booked_vids

# ── Masterに車両IDキーが未設定なら付与 ──
if "車両ID_key" not in df_target.columns:
    df_target["車両ID_key"] = pd.to_numeric(df_target["車両ID"], errors="coerce").astype("Int64")

# ── 行数を変えずに予約ステータスを付与 ──
def get_status(row):
    vid   = row.get("車両ID_key")
    plate = row.get("登録番号", "")
    # 車両IDで基幹Excelにマッチ → 本予約
    if pd.notna(vid) and vid in kikan_booked_vids:
        return "本予約"
    # 登録番号で予約CSVにマッチ → 本予約
    if plate in res_plate_map:
        return res_plate_map[plate]
    return "未予約"

df_target["予約ステータス"] = df_target.apply(get_status, axis=1)
df_target["最終更新日"]    = df_target["登録番号"].map(
    df_reservation_dedup.set_index("登録番号")["最終更新日"].to_dict()
)

# ──────────────────────────────────────────────
# ロジック5: ステータス分類
# ──────────────────────────────────────────────
def classify_status(row):
    kind = row["顧客種別"]
    res = row["予約ステータス"]
    if kind == "リピート":
        if res == "本予約":
            return "①リピート本予約済"
        elif res == "仮予約":
            return "②リピート仮予約済"
        else:
            return "③リピート未決"
    else:  # 新規
        if res == "本予約":
            return "④新規本予約済"
        elif res == "仮予約":
            return "⑤新規仮予約済"
        else:
            return "⑥新規未決"


df_target["ステータス"] = df_target.apply(classify_status, axis=1)

# 満了月カラム
df_target["満了月"] = df_target["車検満了日"].dt.to_period("M").astype(str)

# ──────────────────────────────────────────────
# タブ構成
# ──────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 ダッシュボード",
    "🎯 営業アプローチ抽出",
    "🆕 新規開拓管理",
    "📋 全データ一覧",
])


# ══════════════════════════════════════════════
# TAB 1 : ダッシュボード
# ══════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">📊 全体サマリー</div>', unsafe_allow_html=True)

    # 全体KPI
    total_target = len(df_target)
    total_repeat = len(df_target[df_target["顧客種別"] == "リピート"])
    total_new = len(df_target[df_target["顧客種別"] == "新規"])
    total_reserved = len(df_target[df_target["予約ステータス"].isin(["本予約", "仮予約"])])
    total_confirmed = len(df_target[df_target["予約ステータス"] == "本予約"])
    overall_rate = total_reserved / total_repeat if total_repeat > 0 else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="label">分析対象 合計</div>
            <div class="value">{total_target:,}</div>
            <div class="sub">分析期間内</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="label">リピート対象</div>
            <div class="value">{total_repeat:,}</div>
            <div class="sub">目標分母</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="label">新規対象</div>
            <div class="value">{total_new:,}</div>
            <div class="sub">新規開拓枠</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="label">予約済（本＋仮）</div>
            <div class="value">{total_reserved:,}</div>
            <div class="sub">本予約: {total_confirmed:,}</div>
        </div>""", unsafe_allow_html=True)
    with k5:
        badge_cls = progress_color(overall_rate)
        st.markdown(f"""
        <div class="kpi-card">
            <div class="label">リピート予約進捗率</div>
            <div class="value">{overall_rate:.1%}</div>
            <div class="sub"><span class="{badge_cls}">{progress_label(overall_rate)}</span></div>
        </div>""", unsafe_allow_html=True)

    st.markdown("")

    # ── 満了月選択 ──
    st.markdown('<div class="section-title">📅 月別 × 店舗別 進捗管理</div>', unsafe_allow_html=True)

    available_months = sorted(df_target["満了月"].unique())
    if not available_months:
        st.warning("分析期間内に該当データがありません。")
        st.stop()

    selected_month = st.selectbox("満了月を選択", available_months, index=0)
    df_month = df_target[df_target["満了月"] == selected_month]

    # 3ヶ月前ルール判定
    selected_dt = pd.Period(selected_month, freq="M").to_timestamp()
    months_until = (selected_dt.year - today.year) * 12 + (selected_dt.month - today.month)
    deadline_msg = ""
    if months_until <= 3:
        deadline_msg = "🚨 3ヶ月以内 → 目標70%必達月"
    elif months_until <= 6:
        deadline_msg = "⚠️ 6ヶ月以内 → 予約加速フェーズ"
    else:
        deadline_msg = "ℹ️ 事前アプローチフェーズ"

    st.info(f"満了月まで残り **{months_until}ヶ月** ── {deadline_msg}")

    # ── 店舗別テーブル ──
    stores = sorted(df_month["入庫店舗ID"].unique())

    rows = []
    for store in stores:
        ds = df_month[df_month["入庫店舗ID"] == store]
        repeat = ds[ds["顧客種別"] == "リピート"]
        n_repeat = len(repeat)
        n_confirmed = len(repeat[repeat["予約ステータス"] == "本予約"])
        n_tentative = len(repeat[repeat["予約ステータス"] == "仮予約"])
        n_reserved = n_confirmed + n_tentative
        rate = n_reserved / n_repeat if n_repeat > 0 else 0
        n_new = len(ds[ds["顧客種別"] == "新規"])
        n_new_res = len(ds[(ds["顧客種別"] == "新規") & (ds["予約ステータス"].isin(["本予約", "仮予約"]))])
        rows.append({
            "店舗ID": store,
            "リピート目標数": n_repeat,
            "本予約": n_confirmed,
            "仮予約": n_tentative,
            "予約計": n_reserved,
            "進捗率": rate,
            "判定": progress_label(rate),
            "新規対象": n_new,
            "新規予約": n_new_res,
        })

    df_store = pd.DataFrame(rows)

    # 表示用にフォーマット
    df_display = df_store.copy()
    df_display["進捗率"] = df_display["進捗率"].apply(lambda x: f"{x:.1%}")

    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "店舗ID": st.column_config.TextColumn("店舗ID", width="small"),
            "判定": st.column_config.TextColumn("判定", width="small"),
        },
    )

    # ── 棒グラフ ──
    st.markdown('<div class="section-title">📈 店舗別 進捗率チャート</div>', unsafe_allow_html=True)

    chart_df = df_store[["店舗ID", "進捗率"]].copy()
    chart_df["目標(70%)"] = 0.70
    chart_df = chart_df.set_index("店舗ID")
    st.bar_chart(chart_df, color=["#1a237e", "#e57373"])

    # ── ステータス分布 ──
    st.markdown('<div class="section-title">📊 ステータス分布</div>', unsafe_allow_html=True)
    status_counts = df_month["ステータス"].value_counts().sort_index()
    st.bar_chart(status_counts, color="#3949ab")


# ══════════════════════════════════════════════
# TAB 2 : 営業アプローチ抽出
# ══════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">🎯 アプローチ対象セグメント抽出</div>', unsafe_allow_html=True)
    st.caption("満了月・店舗・条件を指定し、営業リストをCSVダウンロードできます。")

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        seg_month = st.selectbox("対象満了月", available_months, key="seg_month")
    with col_f2:
        all_stores = ["全店舗"] + sorted(df_target["入庫店舗ID"].unique().tolist())
        seg_store = st.selectbox("対象店舗", all_stores, key="seg_store")

    seg_condition = st.radio(
        "抽出条件",
        [
            "リピーター × 未予約（仮予約すらなし）",
            "リピーター × 仮予約のみ（本予約に昇格すべき）",
            "全リピーター × 未決（③リピート未決）",
            "カスタム（下の詳細で指定）",
        ],
        index=0,
    )

    # フィルタ適用
    df_seg = df_target[df_target["満了月"] == seg_month].copy()
    if seg_store != "全店舗":
        df_seg = df_seg[df_seg["入庫店舗ID"] == seg_store]

    if seg_condition.startswith("リピーター × 未予約"):
        df_seg = df_seg[(df_seg["顧客種別"] == "リピート") & (df_seg["予約ステータス"] == "未予約")]
    elif seg_condition.startswith("リピーター × 仮予約"):
        df_seg = df_seg[(df_seg["顧客種別"] == "リピート") & (df_seg["予約ステータス"] == "仮予約")]
    elif seg_condition.startswith("全リピーター"):
        df_seg = df_seg[df_seg["ステータス"] == "③リピート未決"]
    # カスタムの場合はフィルタなし

    # 満了日までの残月数を計算
    df_seg["満了まで(月)"] = df_seg["車検満了日"].apply(
        lambda d: (d.year - today.year) * 12 + (d.month - today.month) if pd.notna(d) else None
    )

    # 4ヶ月以内ハイライト
    st.markdown(f"**抽出件数：{len(df_seg):,}件**")
    urgent = df_seg[df_seg["満了まで(月)"].fillna(99) <= 4]
    if len(urgent) > 0:
        st.error(f"🚨 満了4ヶ月以内の緊急対象：{len(urgent)}件")

    display_cols = ["顧客ID", "車両ID", "登録番号", "車検満了日", "入庫店舗ID",
                    "顧客種別", "予約ステータス", "ステータス", "満了まで(月)"]
    existing_cols = [c for c in display_cols if c in df_seg.columns]

    st.dataframe(
        df_seg[existing_cols].sort_values("車検満了日"),
        use_container_width=True,
        hide_index=True,
    )

    if len(df_seg) > 0:
        st.download_button(
            "📥 アプローチリストをCSVダウンロード",
            to_csv_bytes(df_seg[existing_cols]),
            f"approach_list_{seg_month}_{seg_store}.csv",
            "text/csv",
            use_container_width=True,
        )


# ══════════════════════════════════════════════
# TAB 3 : 新規開拓管理
# ══════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">🆕 新規顧客 開拓管理</div>', unsafe_allow_html=True)
    st.caption("自社未利用（新規）客のうち、仮予約以上が取れているリストを管理します。")

    df_new = df_target[df_target["顧客種別"] == "新規"].copy()

    col_n1, col_n2 = st.columns([1, 2])
    with col_n1:
        new_status_filter = st.multiselect(
            "予約ステータス",
            ["本予約", "仮予約", "未予約"],
            default=["本予約", "仮予約"],
            key="new_status",
        )
    with col_n2:
        new_store_filter = st.multiselect(
            "店舗フィルタ",
            sorted(df_new["入庫店舗ID"].unique().tolist()),
            default=sorted(df_new["入庫店舗ID"].unique().tolist()),
            key="new_store",
        )

    df_new_filtered = df_new[
        (df_new["予約ステータス"].isin(new_status_filter))
        & (df_new["入庫店舗ID"].isin(new_store_filter))
    ]

    # KPI
    kn1, kn2, kn3, kn4 = st.columns(4)
    with kn1:
        st.metric("新規 全体", f"{len(df_new):,}")
    with kn2:
        n_new_hon = len(df_new[df_new["予約ステータス"] == "本予約"])
        st.metric("本予約", f"{n_new_hon:,}")
    with kn3:
        n_new_kari = len(df_new[df_new["予約ステータス"] == "仮予約"])
        st.metric("仮予約", f"{n_new_kari:,}")
    with kn4:
        n_new_none = len(df_new[df_new["予約ステータス"] == "未予約"])
        st.metric("未予約", f"{n_new_none:,}")

    st.markdown("")

    new_display_cols = ["顧客ID", "車両ID", "登録番号", "車検満了日", "入庫店舗ID",
                        "予約ステータス", "最終更新日", "満了月"]
    new_existing = [c for c in new_display_cols if c in df_new_filtered.columns]

    st.dataframe(
        df_new_filtered[new_existing].sort_values("車検満了日"),
        use_container_width=True,
        hide_index=True,
    )

    if len(df_new_filtered) > 0:
        st.download_button(
            "📥 新規開拓リストをCSVダウンロード",
            to_csv_bytes(df_new_filtered[new_existing]),
            "new_customer_list.csv",
            "text/csv",
            use_container_width=True,
        )


# ══════════════════════════════════════════════
# TAB 4 : 全データ一覧
# ══════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">📋 分析期間内 全データ一覧</div>', unsafe_allow_html=True)

    # 検索フィルタ
    search_text = st.text_input("🔍 フリーワード検索（登録番号・顧客IDなど）", key="search")

    df_all = df_target.copy()
    if search_text:
        mask_search = df_all.apply(lambda row: row.astype(str).str.contains(search_text, case=False).any(), axis=1)
        df_all = df_all[mask_search]

    st.markdown(f"**表示件数：{len(df_all):,}件**")
    st.dataframe(
        df_all.sort_values("車検満了日"),
        use_container_width=True,
        hide_index=True,
    )

    if len(df_all) > 0:
        st.download_button(
            "📥 全データCSVダウンロード",
            to_csv_bytes(df_all),
            "all_data_export.csv",
            "text/csv",
            use_container_width=True,
        )

    # 流出済データ
    st.markdown('<div class="section-title">⚠️ 流出済フラグ データ（分析期間外）</div>', unsafe_allow_html=True)
    df_lost = df_master[df_master["流出済"] == True].copy()
    st.markdown(f"分析期間より先に満了日がある車両：**{len(df_lost):,}件**")
    if len(df_lost) > 0:
        st.dataframe(df_lost.head(100), use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────
# フッター
# ──────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center; font-size:0.75rem; color:#999; padding:0.5rem 0;">
    車検満期管理システム v1.0 ─ セッション完結型 ─ データは永続化されません<br>
    AO S研究所理論準拠：3ヶ月前予約率70%目標管理
</div>
""", unsafe_allow_html=True)
