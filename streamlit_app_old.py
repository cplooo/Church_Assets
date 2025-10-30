
import streamlit as st
import pandas as pd
from io import BytesIO
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
    st.title("真耶穌教會斗六教會財產清單")

    # Sidebar: choose location
    locations = sorted([x for x in df["儲存位置"].dropna().unique() if str(x).strip() != ""])
    default_loc = locations[0] if locations else None
    st.sidebar.header("篩選")
    loc = st.sidebar.selectbox("選擇儲存位置", options=locations, index=0 if default_loc else None)

    # Optional: show only 資訊類產品 (Y)
    info_only = st.sidebar.checkbox("只顯示資訊類產品（Y）", value=False)

    # Filter
    q = df[df["儲存位置"] == loc].copy() if loc else df.copy()
    if info_only and "資訊類產品" in q.columns:
        q = q[q["資訊類產品"].str.upper() == "Y"]

    # Main: summary + items
    total = len(q)
    st.subheader(f"儲存位置：{loc}（{total} 筆）" if loc else f"全部位置（{total} 筆）")

    # Show just the item names list (with 編號 for clarity)
    cols_to_show = [c for c in ["編號","物品名稱"] if c in q.columns]
    if cols_to_show:
        st.dataframe(q[cols_to_show].reset_index(drop=True), use_container_width=True, hide_index=True)
    else:
        st.info("找不到欄位：編號、物品名稱")

# ---------- Data Source ----------
# 1) Try default file path in the same folder.
DEFAULT_XLSX = "教會財產盤點_依空間整理.xlsx"
df = None

if os.path.exists(DEFAULT_XLSX):
    try:
        df = load_excel(DEFAULT_XLSX)
    except Exception as e:
        st.warning(f"讀取預設檔案失敗：{e}")

# 2) If not found or failed, allow user to upload
if df is None:
    st.info("請上傳 Excel 檔案（需包含工作表『清單』，欄位：編號、物品名稱、儲存位置、資訊類產品）")
    up = st.file_uploader("上傳 .xlsx 檔", type=["xlsx"])
    if up is not None:
        try:
            df = load_excel(up)
        except Exception as e:
            st.error(f"讀取上傳檔案失敗：{e}")

if df is not None:
    show_location_view(df)
else:
    st.stop()
