import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font
from io import BytesIO

LIST_FILE = "list/ARTICLE PERMANENT GMS CADRE MIROIR MEUBLE.xlsx"
LOW_STOCK_THRESHOLD = 100

RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
RED_FONT = Font(name="Calibri", size=11, color="FFFFFF")


def load_stock_lookup(file) -> dict:
    """Read stock file, return {code: qty} mapping."""
    wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    code_col = next((i for i, h in enumerate(headers) if h and str(h).strip() == "Code"), None)
    qte_col  = next((i for i, h in enumerate(headers) if h and str(h).strip() == "Qte site6"), None)
    if code_col is None or qte_col is None:
        raise ValueError("Stock file must have 'Code' and 'Qte site6' columns.")
    lookup = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[code_col]
        qty  = row[qte_col]
        if code is not None:
            lookup[str(code).strip()] = qty
    wb.close()
    return lookup


def build_excel(list_path: str, stock_lookup: dict) -> tuple[BytesIO, int, int]:
    """Clone list workbook, update Qte site6, colour low-stock cells."""
    wb = openpyxl.load_workbook(list_path)
    ws = wb.active

    headers  = {cell.value: cell.column for cell in ws[1]}
    qte_col  = headers.get("Qte site6")
    code_col = headers.get("Code")

    if not qte_col or not code_col:
        raise ValueError("List file must have 'Code' and 'Qte site6' columns.")

    total     = ws.max_row - 1
    low_count = 0

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
            low_count += 1

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, total, low_count


# ---------------------------------------------------------------------------
st.set_page_config(page_title="Low Stock Checker", layout="centered")
st.title("Low Stock Checker")
st.markdown(
    "Upload the **stock file** — the app updates the article list with fresh "
    "quantities and highlights stock below **100** in red."
)

uploaded = st.file_uploader("Upload stock.xlsx", type=["xlsx"])

if uploaded:
    try:
        stock_lookup = load_stock_lookup(uploaded)
    except ValueError as e:
        st.error(str(e))
        st.stop()

    try:
        excel_buf, total, low_count = build_excel(LIST_FILE, stock_lookup)
    except FileNotFoundError:
        st.error(f"Article list not found at: `{LIST_FILE}`")
        st.stop()
    except ValueError as e:
        st.error(str(e))
        st.stop()

    c1, c2 = st.columns(2)
    c1.metric("Articles in list", total)
    c2.metric(f"Low stock (< {LOW_STOCK_THRESHOLD})", low_count)

    st.download_button(
        label="Download updated list",
        data=excel_buf,
        file_name="low_stock_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
