import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import openpyxl
from io import BytesIO
import os
import re
import datetime

# フォント自動検出（日本語フォントが確実に存在するものを複数候補に）
font_path_candidates = [
    "C:/Windows/Fonts/meiryo.ttc",
    "C:/Windows/Fonts/YuGothM.ttc",
    "C:/Windows/Fonts/msgothic.ttc",
    "C:/Windows/Fonts/msmincho.ttc",
]
font_path = None
for p in font_path_candidates:
    if os.path.isfile(p):
        font_path = p
        break
prop = fm.FontProperties(fname=font_path) if font_path else None

# グローバルにmatplotlibの日本語フォントを設定
if prop:
    plt.rcParams['font.family'] = prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False
else:
    st.warning("日本語フォントが見つかりません。グラフ凡例に□が出る可能性があります。")

st.set_page_config(layout="wide")
st.title('工数集計・可視化アプリ')

staff_to_branch = {
    "a.kani":      "東京",
    "a.murai":     "東京",
    "d.tajima":    "東京",
    "h.meguro":    "東京",
    "h.obata":     "大阪",
    "h.sato":      "東京",
    "k.fujita":    "大阪",
    "k.horio":     "大阪",
    "k.muraoka":   "大阪",
    "k.usami":     "東京",
    "m.maekawa":   "大阪",
    "m.moriguchi": "大阪",
    "m.okawa":     "大阪",
    "r.tanaka":    "大阪",
    "s.tawada":    "大阪",
    "y.hara":      "大阪",
    "y.nakai":     "東京",
    "y.nishitani": "大阪",
    "yuki.sato":   "東京",
}

work_content_list = [
    "移動", "納品試運転", "制御部更新", "点検", "年間保守", "訪問修理", "引取修理", "改造", "移設",
    "お客様対応(保証期限内)", "お客様対応(受注でない)", "社内サポート", "修繕", "貿易管理", "庶務",
    "教育", "標準化", "検査", "組立", "手配", "設計", "その他"
]

def extract_number(val):
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        m = re.match(r'^(\d+)', val)
        if m:
            return int(m.group(1))
    return pd.NA

def convert_month(month_raw):
    if month_raw is None:
        return ""
    if isinstance(month_raw, datetime.datetime):
        return month_raw.strftime("%Y_%m")
    if isinstance(month_raw, str):
        try:
            month_dt = pd.to_datetime(month_raw)
            return month_dt.strftime("%Y_%m")
        except Exception:
            return str(month_raw)
    if isinstance(month_raw, (int, float)):
        try:
            month_dt = pd.to_datetime('1899-12-30') + pd.to_timedelta(float(month_raw), unit='D')
            return month_dt.strftime("%Y_%m")
        except Exception:
            return str(month_raw)
    return str(month_raw)

# --- サイドバーで保存済みCSVデータのアップロード ---
st.sidebar.markdown("---")
st.sidebar.markdown("### 保存したCSVデータを再アップロード（積み上げ用）")
uploaded_csv = st.sidebar.file_uploader("保存済みCSVファイルを選択", type="csv", key="saved_csv")
df_saved = None
if uploaded_csv:
    df_saved = pd.read_csv(uploaded_csv)

uploaded_files = st.file_uploader(
    'Upload Excel files (multiple allowed)',
    type=['xlsx', 'xlsm'],
    accept_multiple_files=True
)

all_data = []

if uploaded_files:
    for up in uploaded_files:
        file_bytes = BytesIO(up.read())
        wb = openpyxl.load_workbook(file_bytes, data_only=True)
        ws = wb["①個人記入欄"]
        month_raw = ws["H2"].value
        month = convert_month(month_raw)
        staff = ws["F1"].value
        branch = staff_to_branch.get(staff, "")

        file_bytes.seek(0)
        df = pd.read_excel(
            file_bytes,
            sheet_name="①個人記入欄",
            header=None,
            skiprows=6,
            usecols=[3, 4, 7, 39],
            nrows=130
        )
        df.columns = ["作業分類", "作業内容", "作業分類元", "工数 [h]"]

        cond = pd.Series([True]*len(df))
        if len(df) > 46:
            cond.iloc[46:] = df.loc[46:, "作業分類"].notna() & (df.loc[46:, "作業分類"].astype(str).str.strip() != "")
        cond = cond & ~(
            df["作業分類"].astype(str).str.startswith("参考行", na=False) |
            df["作業内容"].astype(str).str.startswith("参考行", na=False) |
            df["作業分類元"].astype(str).str.startswith("参考行", na=False) |
            df["工数 [h]"].astype(str).str.startswith("参考行", na=False)
        )

        df_valid = df[cond].copy()
        df_valid["工数 [h]"] = df_valid["工数 [h]"].apply(extract_number)
        df_valid["工数 [h]"] = pd.to_numeric(df_valid["工数 [h]"], errors="coerce")
        df_valid["作業内容_分類"] = df_valid["作業内容"].apply(
            lambda x: x if x in work_content_list else "その他"
        )
        df_valid["月"] = month
        df_valid["担当者"] = staff
        df_valid["支店"] = branch

        all_data.append(df_valid)

    df_all = pd.concat(all_data, ignore_index=True)
else:
    df_all = pd.DataFrame()  # 空データフレーム

# --- 積み上げ（結合）＆重複排除 ---
if df_saved is not None and not df_all.empty:
    df_all = pd.concat([df_saved, df_all], ignore_index=True)
    df_all = df_all.drop_duplicates()
elif df_saved is not None:
    df_all = df_saved

if not df_all.empty:
    st.subheader("抽出・除外済みデータ（積み上げ・重複排除済み）")
    st.dataframe(df_all)

    # --- サイドバー・チェックボックスフィルター ---
    st.sidebar.header("Filter Settings")

    month_list = sorted(df_all["月"].dropna().unique())
    staff_list = sorted(df_all["担当者"].dropna().unique())
    branch_list = sorted(df_all["支店"].dropna().unique())

    selected_months = []
    st.sidebar.markdown("#### 月でフィルター")
    for val in month_list:
        if st.sidebar.checkbox(val, value=True, key=f"month_{val}"):
            selected_months.append(val)

    selected_staffs = []
    st.sidebar.markdown("#### 担当者でフィルター")
    for val in staff_list:
        if st.sidebar.checkbox(val, value=True, key=f"staff_{val}"):
            selected_staffs.append(val)

    selected_branches = []
    st.sidebar.markdown("#### 支店でフィルター")
    for val in branch_list:
        if st.sidebar.checkbox(val, value=True, key=f"branch_{val}"):
            selected_branches.append(val)

    work_content_sum = df_all.groupby("作業内容_分類")["工数 [h]"].sum().to_dict()
    selected_workcontent = []
    st.sidebar.markdown("#### 作業内容でフィルター")
    for name in work_content_list:
        default_check = bool(work_content_sum.get(name, 0))
        checked = st.sidebar.checkbox(name, value=default_check, key=f"work_{name}")
        if checked:
            selected_workcontent.append(name)

    filtered = df_all[
        (df_all["月"].isin(selected_months)) &
        (df_all["担当者"].isin(selected_staffs)) &
        (df_all["支店"].isin(selected_branches)) &
        (df_all["作業内容_分類"].isin(selected_workcontent))
    ]

    st.subheader("フィルタ後データ")
    st.dataframe(filtered)

    # ---- 棒グラフ表示 ----
    st.subheader("作業内容ごとの工数合計（棒グラフ）")
    plot_df = filtered.groupby("作業内容_分類")["工数 [h]"].sum().reset_index()
    plot_df = plot_df[plot_df["工数 [h]"] > 0]  # 0のものは除外

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(plot_df["作業内容_分類"], plot_df["工数 [h]"], label="工数 [h]")

    # 凡例をfontproperties付きで
    ax.legend(fontproperties=prop)
    plt.xticks(rotation=45, ha="right", fontproperties=prop)
    plt.yticks(fontproperties=prop)
    plt.tight_layout()
    st.pyplot(fig)
else:
    st.info("データをアップロードしてください。")
