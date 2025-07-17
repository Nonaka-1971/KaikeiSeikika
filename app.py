# ✅ 必要なライブラリをインストール（最初の1回のみ）
# !pip install streamlit pandas openpyxl

import streamlit as st
import pandas as pd
from io import BytesIO

# ===============================
# ✅ ヘッダー
# ===============================
st.title("📊 要素内訳一覧抽出ツール（仕訳日記帳から）")
st.markdown("仕訳日記帳とテンプレートファイルをアップロードしてください。")

# ===============================
# ✅ ファイルアップロード
# ===============================
shiwake_file = st.file_uploader("📥 仕訳日記帳ファイル（正規化・集計済）をアップロード", type=["xlsx"])
template_file = st.file_uploader("📥 要素内訳一覧テンプレートファイル（空）をアップロード", type=["xlsx"])

if shiwake_file and template_file:
    # ===============================
    # ✅ ファイル読み込み
    # ===============================
    df_all = pd.read_excel(shiwake_file)
    df_all.columns = df_all.columns.str.strip()

    df_template = pd.read_excel(template_file)
    df_template.columns = df_template.columns.str.strip()

    # ===============================
    # ✅ モード判定
    # ===============================
    template_name = template_file.name
    if "原価" in template_name:
        mode = "原価"
    elif "支払" in template_name:
        mode = "支払"
    elif "売上" in template_name:
        mode = "売上"
    elif "入金" in template_name:
        mode = "入金"
    else:
        mode = "不明"

    st.success(f"🔍 判定結果 → 処理モード：{mode}")

    # ===============================
    # ✅ データ抽出
    # ===============================
    if mode == "原価":
        df_target = df_all[df_all["要素内訳貸方勘定科目名称"].isin(["工事未払金", "買掛金"])].copy()
        st.info("✅ 原価モードで貸方勘定科目が一致したデータを抽出します。")
    elif mode == "支払":
        df_target = df_all[df_all["要素内訳借方勘定科目名称"] == "工事未払金"].copy()
        st.info("✅ 支払モードで借方勘定科目が一致したデータを抽出します。")
    elif mode == "売上":
        df_target = df_all[df_all["要素内訳借方勘定科目名称"].isin(["工事未収金", "未収入金"])].copy()
        st.info("✅ 売上モードで借方勘定科目が一致したデータを抽出します。")
    elif mode == "入金":
        df_target = df_all[df_all["要素内訳貸方勘定科目名称"].isin(["未成工事受入金", "未収入金"])].copy()
        st.info("✅ 入金モードで貸方勘定科目が一致したデータを抽出します。")
    else:
        st.error(⚠️ テンプレートファイル名から処理モードが判別できません。")

    # ===============================
    # ✅ 列構造に合わせて再構成し、ダウンロード可能に
    # ===============================
    if not df_target.empty:
        for col in df_template.columns:
            if col not in df_target.columns:
                df_target[col] = ""

        df_target = df_target[df_template.columns]

        output = BytesIO()
        df_target.to_excel(output, index=False)
        output.seek(0)

        st.success(f"✅ {mode} モードの抽出結果が {len(df_target)} 件あります。")
        st.download_button(
            label="📤 抽出結果をダウンロード",
            data=output,
            file_name=f"{mode}_要素内訳一覧_抽出結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ 抽出対象のデータがありませんでした。")
