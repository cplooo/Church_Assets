
import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="真耶穌教會斗六教會財產清單", layout="wide")

# ---------- Helpers ----------
@st.cache_data
def load_excel(file_obj_or_path):
    if isinstance(file_obj_or_path, (str, os.PathLike)):
        df = pd.read_excel(file_obj_or_path, sheet_name="清單", engine="openpyxl")
    else:
        df = pd.read_excel(file_obj_or_path, sheet_name="清單", engine="openpyxl")
    # Normalize & clean
    for col in ["編號", "物品名稱", "儲存位置", "資訊類產品"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    # Keep the 4 columns only, if extra columns exist
    keep_cols = [c for c in ["編號","物品名稱","儲存位置","資訊類產品"] if c in df.columns]
    df = df[keep_cols].dropna(how="all")
    # Sort by 儲存位置 then 編號 numeric if possible
    if "編號" in df.columns:
        try:
            df["_num"] = pd.to_numeric(df["編號"], errors="coerce")
            df = df.sort_values(["儲存位置","_num","編號"], na_position="last").drop(columns=["_num"])
        except Exception:
            pass
    return df

def show_location_view(df):
    locations = sorted([x for x in df["儲存位置"].dropna().unique() if str(x).strip() != ""])
    default_loc = locations[0] if locations else None

    col1, col2 = st.columns([2,1])
    with col1:
        loc = st.selectbox("選擇儲存位置", options=locations, index=0 if default_loc else None)
    with col2:
        info_only = st.checkbox("只顯示資訊類產品（Y）", value=False)

    q = df[df["儲存位置"] == loc].copy() if loc else df.copy()
    if info_only and "資訊類產品" in q.columns:
        q = q[q["資訊類產品"].str.upper() == "Y"]

    total = len(q)
    st.subheader(f"儲存位置：{loc}（{total} 筆）" if loc else f"全部位置（{total} 筆）")

    cols_to_show = [c for c in ["編號","物品名稱"] if c in q.columns]
    if cols_to_show:
        st.dataframe(q[cols_to_show].reset_index(drop=True), use_container_width=True, hide_index=True)
    else:
        st.info("找不到欄位：編號、物品名稱")

def show_item_locations_view(df):
    with st.expander("說明", expanded=False):
        st.write("此頁列出每一個 **物品名稱** 及其出現的所有 **儲存位置**。可用上方搜尋欄快速過濾。")

    left, right = st.columns([2,1], gap="large")
    with left:
        keyword = st.text_input("搜尋物品名稱（支援關鍵字）", value="")
    with right:
        info_only = st.checkbox("只顯示資訊類產品（Y）", value=False)

    q = df.copy()
    if info_only:
        q = q[q["資訊類產品"].str.upper() == "Y"]

    if keyword.strip():
        q = q[q["物品名稱"].str.contains(keyword.strip(), case=False, na=False)]

    def _unique_join(series):
        uniq = sorted({str(x).strip() for x in series if str(x).strip() != ""})
        return "、".join(uniq)

    grp = (q
           .groupby("物品名稱", dropna=False)
           .agg(儲存位置清單=("儲存位置", _unique_join),
                位置數量=("儲存位置", lambda s: len(set([str(x).strip() for x in s if str(x).strip() != ""]))),
                物品數量=("編號", "count"))
           .reset_index())

    grp = grp.sort_values(by=["物品名稱", "位置數量"], ascending=[True, False], ignore_index=True)

    st.caption(f"共 **{len(grp)}** 種物品名稱。")
    st.dataframe(grp, use_container_width=True, hide_index=True)

    csv_bytes = grp.to_csv(index=False).encode("utf-8-sig")
    st.download_button("下載彙總（CSV）", data=csv_bytes, file_name="物品名稱_對應儲存位置_彙總.csv", mime="text/csv")

# ---------- Data Source ----------
DEFAULT_XLSX = "教會財產盤點_依空間整理.xlsx"
df = None

if os.path.exists(DEFAULT_XLSX):
    try:
        df = load_excel(DEFAULT_XLSX)
    except Exception as e:
        st.warning(f"讀取預設檔案失敗：{e}")

if df is None:
    st.info("請上傳 Excel 檔案（需包含工作表『清單』，欄位：編號、物品名稱、儲存位置、資訊類產品）")
    up = st.file_uploader("上傳 .xlsx 檔", type=["xlsx"])
    if up is not None:
        try:
            df = load_excel(up)
        except Exception as e:
            st.error(f"讀取上傳檔案失敗：{e}")

# ---------- UI ----------
st.title("真耶穌教會斗六教會財產清單")  # fixed at top

if df is not None:
    mode = st.radio("選擇瀏覽模式", ["🧾 依物品名稱彙總", "📁 依儲存位置瀏覽"], horizontal=True, index=0)

    if mode.startswith("🧾"):
        show_item_locations_view(df)
    else:
        show_location_view(df)
else:
    st.stop()
