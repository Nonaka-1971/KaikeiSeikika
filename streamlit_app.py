# streamlit_app.py

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="📊 仕訳日記帳 正規化＆集計ツール", layout="centered")

st.title("📊 仕訳日記帳 正規化＆集計ツール")
st.markdown("""
このツールでは、仕訳日記帳の複数ファイルを列名マスタに基づいて正規化し、  
貸借の組み合わせ単位で集計処理を行い、整合性チェックをしてExcelでダウンロードできます。
""")

# === 1. 列名マスタのアップロード ===
st.header("1️⃣ 列名マスタをアップロード")
master_file = st.file_uploader("『列名マスタテンプレート横.xlsx』をアップロードしてください", type=["xlsx"])

# === 2. 仕訳日記帳ファイルのアップロード ===
st.header("2️⃣ 仕訳日記帳ファイルをアップロード（複数可）")
uploaded_files = st.file_uploader(
    "仕訳日記帳ファイルを複数選択できます",
    type=["xlsx"],
    accept_multiple_files=True
)

if master_file and uploaded_files:
    # === 列名マスタの読み込み ===
    df_master = pd.read_excel(master_file)
    df_master.columns = df_master.columns.str.strip()

    column_mappings = {}
    for idx, row in df_master.iterrows():
        software_name = f"row{idx}"
        mapping = {str(v).strip(): col for col, v in row.items() if pd.notna(v)}
        column_mappings[software_name] = mapping

    for uploaded_file in uploaded_files:
        st.divider()
        st.subheader(f"📂 ファイル処理: {uploaded_file.name}")

        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()

        # 自動判定
        best_match = None
        best_score = 0
        for software, mapping in column_mappings.items():
            matched = len([c for c in df.columns if c in mapping])
            if matched > best_score:
                best_score = matched
                best_match = software
        colmap = column_mappings[best_match]
        st.write(f"✅ 推定形式: **{best_match}** （一致列数: {best_score}）")

        # 列名変換
        df_renamed = df.rename(columns=colmap)

        # 前処理
        for col in ["貸方科目コード", "貸方補助コード"]:
            if col in df_renamed.columns:
                df_renamed[col] = df_renamed[col].replace({0: np.nan, "0": np.nan})
        for col in ["貸方科目コード", "貸方科目名", "貸方補助コード", "貸方補助科目名"]:
            if col in df_renamed.columns:
                df_renamed[col] = df_renamed[col].fillna(method="ffill")

        # ソート
        sort_cols = [c for c in ["年", "月", "日", "伝票No"] if c in df_renamed.columns]
        if sort_cols:
            df_renamed = df_renamed.sort_values(by=sort_cols)

        # 出力対象カラム
        output_columns = [
            "要素内訳借方勘定科目コード", "要素内訳借方勘定科目名称", "要素内訳借方補助科目コード", "要素内訳借方補助科目名称",
            "要素内訳借方税区分", "要素内訳借方部門コード", "要素内訳借方部門名称", "要素内訳借方予備",
            "要素内訳貸方勘定科目コード", "要素内訳貸方勘定科目名称", "要素内訳貸方補助科目コード", "要素内訳貸方補助科目名称",
            "要素内訳貸方税区分", "要素内訳貸方部門コード", "要素内訳貸方部門名称", "要素内訳貸方予備",
            "借方金額", "貸方金額", "借方消費税金額", "貸方消費税金額"
        ]

        # グループキーと集計列
        group_keys = [c for c in output_columns if c not in ["借方金額", "貸方金額", "借方消費税金額", "貸方消費税金額"] and c in df_renamed.columns]
        sum_columns = [c for c in ["借方金額", "貸方金額", "借方消費税金額", "貸方消費税金額"] if c in df_renamed.columns]

        if not group_keys:
            st.warning("⚠️ 正規カラムが不足しているためスキップします。")
            st.write("現在のカラム:", df_renamed.columns.tolist())
            continue

        for col in sum_columns:
            df_renamed[col] = pd.to_numeric(df_renamed[col], errors="coerce")

        # 集計
        df_grouped = df_renamed.groupby(group_keys, dropna=False)[sum_columns].sum().reset_index()

        # 出力列順
        final_columns = [col for col in output_columns if col in df_grouped.columns]
        df_final = df_grouped[final_columns]

        # 合計行追加
        total_row = [df_final[col].sum() if col in sum_columns else "" for col in df_final.columns]
        df_final = pd.concat([df_final, pd.DataFrame([total_row], columns=df_final.columns)], ignore_index=True)

        # 整合性チェック
        original_combinations = df_renamed[group_keys].drop_duplicates().shape[0]
        grouped_combinations = df_grouped.shape[0]

        st.write(f"- 貸借組み合わせ件数（元データ）: {original_combinations}")
        st.write(f"- 貸借組み合わせ件数（集計後）: {grouped_combinations}")
        if original_combinations == grouped_combinations:
            st.success("✔️ 件数一致")
        else:
            st.error("❌ 件数不一致")

        for col in sum_columns:
            original_total = df_renamed[col].sum()
            grouped_total = df_grouped[col].sum()
            match = "一致 ✔️" if np.isclose(original_total, grouped_total, rtol=1e-5) else "不一致 ❌"
            st.write(f"- {col} 合計: 元 = {original_total:,.0f} / 集計後 = {grouped_total:,.0f} → {match}")

        # ダウンロード
        towrite = BytesIO()
        df_final.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button(
            label="⬇️ 正規化集計済ファイルをダウンロード",
            data=towrite,
            file_name=uploaded_file.name.replace(".xlsx", "") + "_正規化集計済.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.caption("👨‍💻 Powered by Streamlit")
