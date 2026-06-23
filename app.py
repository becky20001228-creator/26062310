import re
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="컨트롤러 재고 대시보드", layout="wide")


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_matching_column(columns, keywords):
    for col in columns:
        col_text = str(col).lower()
        if all(keyword.lower() in col_text for keyword in keywords):
            return col
    return None


@st.cache_data(show_spinner=False)
def read_sheet_names(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def load_sheet(file_bytes, sheet_name):
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)


st.title("컨트롤러 재고 대시보드")
st.caption("엑셀을 업로드하면 자사 창고 및 외주사 창고의 컨트롤러 재고를 표로 확인할 수 있습니다.")

uploaded_file = st.file_uploader("관련 엑셀 파일 업로드", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("엑셀 파일을 업로드해 주세요.")
    st.stop()

file_bytes = uploaded_file.getvalue()
sheet_names = read_sheet_names(file_bytes)
sheet_name = st.selectbox("시트 선택", sheet_names)
df = load_sheet(file_bytes, sheet_name)

if df.empty:
    st.warning("선택한 시트에 데이터가 없습니다.")
    st.stop()

st.subheader("컬럼 매핑")
columns = list(df.columns)

default_warehouse = find_matching_column(columns, ["창고"]) or columns[0]
default_item = find_matching_column(columns, ["품목"]) or columns[min(1, len(columns) - 1)]
default_spec = find_matching_column(columns, ["규격"]) or columns[min(2, len(columns) - 1)]
default_qty = (
    find_matching_column(columns, ["양품", "재고"])
    or find_matching_column(columns, ["수량"])
    or columns[min(3, len(columns) - 1)]
)

col1, col2, col3, col4 = st.columns(4)
with col1:
    warehouse_col = st.selectbox("창고명 컬럼", columns, index=columns.index(default_warehouse))
with col2:
    item_col = st.selectbox("품목 컬럼", columns, index=columns.index(default_item))
with col3:
    spec_col = st.selectbox("규격 컬럼", columns, index=columns.index(default_spec))
with col4:
    qty_col = st.selectbox("양품 재고 수량 컬럼", columns, index=columns.index(default_qty))

st.subheader("필터 설정")
warehouse_filter = st.text_input(
    "창고명에 포함될 키워드(콤마로 구분, 비우면 전체 표시)",
    value="자사,외주사",
)
warehouse_keywords = [x.strip() for x in warehouse_filter.split(",") if x.strip()]

result = df.copy()
result[warehouse_col] = result[warehouse_col].map(normalize_text)
result[item_col] = result[item_col].map(normalize_text)
result[spec_col] = result[spec_col].map(normalize_text)
result[qty_col] = pd.to_numeric(result[qty_col], errors="coerce").fillna(0)

if warehouse_keywords:
    pattern = "|".join(re.escape(keyword) for keyword in warehouse_keywords)
    result = result[result[warehouse_col].str.contains(pattern, case=False, na=False, regex=True)]

result = result[[warehouse_col, item_col, spec_col, qty_col]].copy()
result.columns = ["창고명", "품목", "규격", "양품 재고 수량"]
result = result.groupby(["창고명", "품목", "규격"], as_index=False)["양품 재고 수량"].sum()
result = result.sort_values(["창고명", "품목", "규격"]).reset_index(drop=True)

st.subheader("재고 현황")
st.dataframe(result, use_container_width=True, hide_index=True)

col_a, col_b, col_c = st.columns(3)
with col_a:
    st.metric("행 수", len(result))
with col_b:
    st.metric("총 양품 재고", int(result["양품 재고 수량"].sum()) if not result.empty else 0)
with col_c:
    st.metric("창고 수", result["창고명"].nunique() if not result.empty else 0)

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    result.to_excel(writer, index=False, sheet_name="재고현황")
output.seek(0)

st.download_button(
    "결과 엑셀 다운로드",
    data=output,
    file_name="컨트롤러_재고_대시보드.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
