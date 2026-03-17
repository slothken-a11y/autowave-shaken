"""
予約データ CSVコンバーター
===========================
予約システム出力CSV（csv_order_xxxxx.csv）を
車検管理システム用 Reservation_Data.csv に変換します。

【変換ルール】
  - 予約状況が「キャンセル（店舗）」「キャンセル（アプリ）」→ 除外
  - それ以外（来店前・終了 等）→ 予約ステータス「本予約」
  - 登録番号のスペースを除去（例：「千葉 480 て 9579」→「千葉480て9579」）
  - 登録番号が空・電話番号（数字10〜11桁）のレコードは除外
  - 同一登録番号が複数ある場合は最新の予約受付日時を採用

【出力形式】
  登録番号, 予約ステータス, 最終更新日
"""

import pandas as pd
import re
import sys
from io import BytesIO


def is_valid_plate(s: str) -> bool:
    """登録番号として有効かチェック"""
    s = str(s).strip()
    if s in ["nan", "", "None"]:
        return False
    # 電話番号（数字のみ10〜11桁）は除外
    if re.fullmatch(r"\d{10,11}", s):
        return False
    return True


def convert_reservation(input_path: str, output_path: str = None) -> pd.DataFrame:
    """
    予約システムCSVをReservation_Data.csv形式に変換する

    Parameters
    ----------
    input_path : str  入力CSVパス
    output_path : str  出力CSVパス（Noneの場合は保存しない）

    Returns
    -------
    pd.DataFrame  変換後のDataFrame
    """

    # ── 読み込み（エンコーディング自動判定）──
    df = None
    for enc in ["cp932", "utf-8-sig", "utf-8", "shift_jis", "latin1"]:
        try:
            df = pd.read_csv(input_path, encoding=enc)
            break
        except Exception:
            continue
    if df is None:
        raise ValueError(f"ファイルを読み込めません: {input_path}")

    df.columns = df.columns.str.strip().str.replace("\u3000", "")
    total_input = len(df)

    # ── ① キャンセル除外 ──
    cancel_mask = df["予約状況"].astype(str).str.startswith("キャンセル")
    df_valid = df[~cancel_mask].copy()
    n_cancelled = cancel_mask.sum()

    # ── ② 登録番号クリーニング ──
    df_valid["登録番号"] = (
        df_valid["登録番号"]
        .astype(str)
        .str.replace(" ", "", regex=False)
        .str.strip()
    )

    # ── ③ 無効な登録番号を除外 ──
    valid_plate_mask = df_valid["登録番号"].apply(is_valid_plate)
    df_valid = df_valid[valid_plate_mask]
    n_invalid_plate = (~valid_plate_mask).sum()

    # ── ④ 予約ステータス = 本予約 ──
    df_valid["予約ステータス"] = "本予約"

    # ── ⑤ 最終更新日 = 予約受付日時 ──
    df_valid["最終更新日"] = df_valid["予約受付日時"]

    # ── ⑥ 必要列のみ抽出 ──
    df_out = df_valid[["登録番号", "予約ステータス", "最終更新日"]].copy()

    # ── ⑦ 重複登録番号 → 最新の受付日時を採用 ──
    df_out["最終更新日_dt"] = pd.to_datetime(df_out["最終更新日"], errors="coerce")
    df_out = (
        df_out.sort_values("最終更新日_dt", ascending=False)
        .drop_duplicates(subset="登録番号", keep="first")
        .drop(columns=["最終更新日_dt"])
        .sort_index()
        .reset_index(drop=True)
    )

    # ── 変換サマリ出力 ──
    print("=" * 50)
    print("  予約データ変換完了")
    print("=" * 50)
    print(f"  入力件数         : {total_input:,}件")
    print(f"  キャンセル除外   : {n_cancelled:,}件")
    print(f"  無効登録番号除外 : {n_invalid_plate:,}件")
    print(f"  重複除去後       : {len(df_out):,}件  ← 出力件数")
    print(f"  予約ステータス   : 本予約 {len(df_out):,}件")
    print("=" * 50)

    # ── 保存 ──
    if output_path:
        df_out.to_csv(output_path, index=False, encoding="utf-8-sig")
        print(f"  保存先: {output_path}")

    return df_out


# ──────────────────────────────────────────────
# Streamlit UI（単体起動時）
# ──────────────────────────────────────────────
if __name__ == "__main__":

    # コマンドライン実行
    if len(sys.argv) >= 2:
        input_file  = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) >= 3 else "Reservation_Data.csv"
        convert_reservation(input_file, output_file)
        sys.exit(0)

    # Streamlit UI
    try:
        import streamlit as st

        st.set_page_config(page_title="予約データ コンバーター", page_icon="🔄", layout="centered")

        st.markdown("""
        <div style="background:linear-gradient(135deg,#1a237e,#283593);color:white;
                    padding:1rem 1.5rem;border-radius:10px;margin-bottom:1.5rem;text-align:center;">
            <h2 style="margin:0;">🔄 予約データ CSVコンバーター</h2>
            <p style="margin:0.3rem 0 0;font-size:0.85rem;opacity:0.85;">
                予約システム出力CSV → Reservation_Data.csv 変換ツール
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        **変換ルール**
        - キャンセル（店舗・アプリ）→ **除外**
        - 来店前・終了 等 → **本予約**
        - 登録番号のスペース除去・無効値除外
        - 同一登録番号は最新受付日時を採用
        """)

        uploaded = st.file_uploader(
            "予約システムCSVをアップロード（csv_order_xxxxx.csv）",
            type=["csv"],
        )

        if uploaded:
            # 一時ファイルに書き出して変換
            import tempfile, os
            with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
                tmp.write(uploaded.read())
                tmp_path = tmp.name

            try:
                df_result = convert_reservation(tmp_path)
            finally:
                os.unlink(tmp_path)

            # 結果表示
            col1, col2, col3 = st.columns(3)
            col1.metric("出力件数", f"{len(df_result):,}件")
            col2.metric("予約ステータス", "本予約のみ")
            col3.metric("登録番号ユニーク", f"{df_result['登録番号'].nunique():,}件")

            st.dataframe(df_result.head(20), use_container_width=True, hide_index=True)

            # ダウンロード
            buf = BytesIO()
            df_result.to_csv(buf, index=False, encoding="utf-8-sig")
            st.download_button(
                "📥 Reservation_Data.csv をダウンロード",
                buf.getvalue(),
                "Reservation_Data.csv",
                "text/csv",
                use_container_width=True,
                type="primary",
            )

    except ImportError:
        print("Streamlit未インストール。コマンドライン実行してください。")
        print("使い方: python reservation_converter.py 入力.csv 出力.csv")
