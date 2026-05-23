import io
from typing import List

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# --- CONFIGURATION ---
st.set_page_config(page_title="Radiant Bridges Pro", page_icon="🛗", layout="wide")

DEFAULT_VENDORS = ["KONE", "TKE", "EEE", "AG MELCO"]

SPEC_LIST = [
    "CAPACITY", "SPEED", "DOOR TYPE", "DOOR SIZE (W X H)",
    "SHAFT SIZE (W X D)", "CABIN SIZE (W X D X H)",
    "OVER HEAD HEIGHT", "PIT DEPTH", "NO. OF LIFTS"
]

# --- INITIALIZATION ---
def make_default_df(vendors: List[str]) -> pd.DataFrame:
    df = pd.DataFrame(
        "",
        index=range(len(SPEC_LIST)),
        columns=["Specification", "Consultant"] + vendors
    )
    df["Specification"] = SPEC_LIST
    return df


def init_state():
    if "vendors" not in st.session_state:
        st.session_state.vendors = DEFAULT_VENDORS.copy()

    if "groups" not in st.session_state:
        st.session_state.groups = {
            "Tower A": {
                "PL1 & PL2": make_default_df(st.session_state.vendors)
            }
        }

    if "commercial_data" not in st.session_state:
        st.session_state.commercial_data = {
            "pricing": pd.DataFrame(columns=["Description"] + st.session_state.vendors),
            "payment": pd.DataFrame(columns=["Term"] + st.session_state.vendors),
            "delivery": pd.DataFrame(columns=["Milestone"] + st.session_state.vendors),
        }


init_state()

# --- EXCEL EXPORT LOGIC ---
def build_highlighted_excel(groups, commercial_data, vendors) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Report"

    # Styles
    header_fill = PatternFill("solid", fgColor="1F4E78")
    group_fill = PatternFill("solid", fgColor="BDD7EE")
    spec_fill = PatternFill("solid", fgColor="E2F0D9")
    match_fill = PatternFill("solid", fgColor="BBF7D0")
    deviation_fill = PatternFill("solid", fgColor="FECACA")
    missing_fill = PatternFill("solid", fgColor="FDE68A")

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def normalize(value):
        return str(value).strip().lower().replace(" ", "")

    def apply_style(cell, fill=None, bold=False, color="000000"):
        cell.border = border
        cell.font = Font(bold=bold, color=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fill:
            cell.fill = fill

    row_no = 1

    # 1. Technical Comparison
    for tower, group_data in groups.items():
        for group_name, df in group_data.items():
            ws.merge_cells(start_row=row_no, start_column=1, end_row=row_no, end_column=len(df.columns))
            title_cell = ws.cell(row_no, 1, f"{tower} - {group_name}")
            apply_style(title_cell, group_fill, True)
            row_no += 1

            for col_idx, col_name in enumerate(df.columns, 1):
                apply_style(ws.cell(row_no, col_idx, col_name), header_fill, True, "FFFFFF")
            row_no += 1

            for _, row in df.iterrows():
                consultant_value = str(row.get("Consultant", "")).strip()

                for col_idx, val in enumerate(row, 1):
                    col_name = df.columns[col_idx - 1]
                    fill = None
                    bold = False

                    if col_name == "Specification":
                        fill = spec_fill
                        bold = True

                    elif col_name in vendors and consultant_value:
                        vendor_value = str(val).strip()

                        if not vendor_value:
                            fill = missing_fill
                        elif normalize(vendor_value) == normalize(consultant_value):
                            fill = match_fill
                        else:
                            fill = deviation_fill

                    apply_style(ws.cell(row_no, col_idx, str(val)), fill, bold)

                row_no += 1

            row_no += 2

    # 2. Commercial Sections
    def write_section(ws, row, title, df):
        if df.empty:
            return row

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(1, len(df.columns)))
        title_cell = ws.cell(row, 1, title)
        apply_style(title_cell, group_fill, True)
        row += 1

        for i, col in enumerate(df.columns, 1):
            apply_style(ws.cell(row, i, col), header_fill, True, "FFFFFF")
        row += 1

        for _, r in df.iterrows():
            for i, val in enumerate(r, 1):
                apply_style(ws.cell(row, i, str(val)))
            row += 1

        return row + 2

    row_no = write_section(ws, row_no, "PRICING SUMMARY", commercial_data["pricing"])
    row_no = write_section(ws, row_no, "PAYMENT TERMS", commercial_data["payment"])
    row_no = write_section(ws, row_no, "DELIVERY PROGRAM", commercial_data["delivery"])

    # Column widths
    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 24
    ws.column_dimensions["A"].width = 32

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# --- UI ---
st.title("🛗 Radiant Bridges Pro Hub")

# Technical comparison editor
st.subheader("Technical Comparison")

tower = st.selectbox("Select Tower", list(st.session_state.groups.keys()))
group = st.selectbox("Select Group", list(st.session_state.groups[tower].keys()))

st.session_state.groups[tower][group] = st.data_editor(
    st.session_state.groups[tower][group],
    num_rows="dynamic",
    use_container_width=True,
    key=f"tech_editor_{tower}_{group}",
)

# Compliance preview
st.subheader("Deviation Preview")

def highlight_mismatches(row):
    styles = [""] * len(row)
    consultant = str(row.get("Consultant", "")).strip()

    if not consultant:
        return styles

    for i, col in enumerate(row.index):
        if col in st.session_state.vendors:
            vendor_value = str(row.get(col, "")).strip()

            if not vendor_value:
                styles[i] = "background-color: #fde68a"
            elif vendor_value.lower().replace(" ", "") == consultant.lower().replace(" ", ""):
                styles[i] = "background-color: #bbf7d0"
            else:
                styles[i] = "background-color: #fecaca"

    return styles

st.dataframe(
    st.session_state.groups[tower][group].style.apply(highlight_mismatches, axis=1),
    use_container_width=True,
)

# Commercial data input
with st.sidebar.expander("Commercial & Pricing Data", expanded=True):
    st.write("Pricing Summary")
    st.session_state.commercial_data["pricing"] = st.data_editor(
        st.session_state.commercial_data["pricing"],
        num_rows="dynamic",
        use_container_width=True,
        key="pricing_editor",
    )

    st.write("Payment Terms")
    st.session_state.commercial_data["payment"] = st.data_editor(
        st.session_state.commercial_data["payment"],
        num_rows="dynamic",
        use_container_width=True,
        key="payment_editor",
    )

    st.write("Delivery Program")
    st.session_state.commercial_data["delivery"] = st.data_editor(
        st.session_state.commercial_data["delivery"],
        num_rows="dynamic",
        use_container_width=True,
        key="delivery_editor",
    )

# Export
st.subheader("Export")

data = build_highlighted_excel(
    st.session_state.groups,
    st.session_state.commercial_data,
    st.session_state.vendors,
)

st.download_button(
    "📥 Download Full Highlighted Excel Report",
    data=data,
    file_name="Radiant_Bridges_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)
