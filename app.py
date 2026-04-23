import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from copy import copy
from io import BytesIO

LIST_FILE = "list/ARTICLE PERMANENT GMS CADRE MIROIR MEUBLE.xlsx"
LOW_STOCK_THRESHOLD = 100

RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
RED_FONT   = Font(name="Calibri", size=11, color="FFFFFF")


def build_excel(list_wb, stock_lookup: dict) -> BytesIO:
    """Clone the list workbook, update Qte site6, color low-stock cells."""
    ws = list_wb.active

    # find the Qte site6 column index and Code column index from header row
    header = {cell.value: cell.column for cell in ws[1]}
    qte_col  = header.get("Qte site6")
    code_col = header.get("Code")

    if not qte_col or not code_col:
        raise ValueError("Could not find 'Code' or 'Qte site6' columns in list file.")

    for row in ws.iter_rows(min_row=2):
        code_cell = row[code_col - 1]
        qte_cell  = row[qte_col  - 1]

        code = str(code_cell.value).strip() if code_cell.value is not None else ""
        if code in stock_lookup:
            qte_cell.value = stock_lookup[code]

        qty = qte_cell.value
        if isinstance(qty, (int, float)) and qty < LOW_STOCK_THRESHOLD:
            qte_cell.fill = RED_FILL
            qte_cell.font = RED_FONT

    buf = BytesIO()
    list_wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
st.set_page_config(page_title="Low Stock Checker", layout="centered")
st.title("Low Stock Checker")
st.markdown("Upload the **stock file** — the app updates the article list with fresh quantities and highlights stock below **100** in red.")

uploaded = st.file_uploader("Upload stock.xlsx", type=["xlsx"])

if uploaded:
    # build stock lookup: Code -> Qte site6
    stock_df = pd.read_excel(uploaded, sheet_name=0, dtype={"Code": str})
    stock_df.columns = stock_df.columns.str.strip()
    stock_df["Code"] = stock_df["Code"].astype(str).str.strip()
    stock_lookup = dict(zip(stock_df["Code"], stock_df["Qte site6"]))

    # open the list workbook (preserves all formatting/structure)
    try:
        list_wb = openpyxl.load_workbook(LIST_FILE)
    except FileNotFoundError:
        st.error(f"Could not find the article list at: `{LIST_FILE}`")
        st.stop()

    excel_buf = build_excel(list_wb, stock_lookup)

    # quick stats for display
    ws = list_wb.active
    header = {cell.value: cell.column for cell in ws[1]}
    qte_col = header.get("Qte site6")
    quantities = [ws.cell(r, qte_col).value for r in range(2, ws.max_row + 1)]
    low_count = sum(1 for q in quantities if isinstance(q, (int, float)) and q < LOW_STOCK_THRESHOLD)

    c1, c2 = st.columns(2)
    c1.metric("Articles in list", ws.max_row - 1)
    c2.metric(f"Low stock (< {LOW_STOCK_THRESHOLD})", low_count)

    st.download_button(
        label="Download updated list",
        data=excel_buf,
        file_name="low_stock_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
