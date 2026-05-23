import io
from typing import List

import pandas as pd
import streamlit as st

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


st.set_page_config(page_title="Radiant Bridges Pro Hub", page_icon="🛗", layout="wide")

DEFAULT_VENDORS = ["KONE", "TKE", "EEE", "AG MELCO"]

SPEC_LIST = [
    "CAPACITY",
    "SPEED",
    "DOOR TYPE",
    "DOOR SIZE (W X H)",
    "SHAFT SIZE (W X D)",
    "CABIN SIZE (W X D X H)",
    "OVER HEAD HEIGHT",
    "PIT DEPTH",
    "NO. OF LIFTS",
]


def make_default_df(vendors: List[str]) -> pd.DataFrame:
    df = pd.DataFrame(
        "",
        index=range(len(SPEC_LIST)),
        columns=["Specification", "Consultant"] + vendors,
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
            "notes": pd.DataFrame(columns=["Note"] + st.session_state.vendors),
        }

    if "attachments" not in st.session_state:
        st.session_state.attachments = []


init_state()


def normalize(value):
    return str(value).strip().lower().replace(" ", "")


def sync_vendor_columns():
    for tower in st.session_state.groups:
        for group in st.session_state.groups[tower]:
            df = st.session_state.groups[tower][group]

            required_cols = ["Specification", "Consultant"] + st.session_state.vendors

            for col in required_cols:
                if col not in df.columns:
                    df[col] = ""

            st.session_state.groups[tower][group] = df[required_cols]

    for key in st.session_state.commercial_data:
        df = st.session_state.commercial_data[key]
        first_col = df.columns[0] if len(df.columns) else "Description"
        required_cols = [first_col] + st.session_state.vendors

        for col in required_cols:
            if col not in df.columns:
                df[col] = ""

        st.session_state.commercial_data[key] = df[required_cols]


def highlight_main_table(row):
    styles = [""] * len(row)
    consultant = str(row.get("Consultant", "")).strip()

    if not consultant:
        return styles

    for i, col in enumerate(row.index):
        if col in st.session_state.vendors:
            vendor_value = str(row.get(col, "")).strip()

            if vendor_value and normalize(vendor_value) != normalize(consultant):
                styles[i] = "background-color: #fde68a"

    return styles


def build_excel(groups, commercial_data, vendors) -> bytes:
    if not OPENPYXL_AVAILABLE:
        output = io.StringIO()
        rows = []

        for tower, group_data in groups.items():
            for group, df in group_data.items():
                rows.append([f"{tower} - {group}"])
                rows.append(list(df.columns))
                for _, r in df.iterrows():
                    rows.append([r.get(c, "") for c in df.columns])
                rows.append([])

        for title, key in [
            ("PRICING SUMMARY", "pricing"),
            ("PAYMENT TERMS", "payment"),
            ("DELIVERY PROGRAM", "delivery"),
            ("NOTES", "notes"),
        ]:
            df = commercial_data[key]
            rows.append([title])
            rows.append(list(df.columns))
            for _, r in df.iterrows():
                rows.append([r.get(c, "") for c in df.columns])
            rows.append([])

        pd.DataFrame(rows).to_csv(output, index=False, header=False)
        return output.getvalue().encode("utf-8-sig")

    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Report"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    group_fill = PatternFill("solid", fgColor="BDD7EE")
    spec_fill = PatternFill("solid", fgColor="E2F0D9")
    deviation_fill = PatternFill("solid", fgColor="FDE68A")

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def apply_style(cell, fill=None, bold=False, color="000000"):
        cell.border = border
        cell.font = Font(bold=bold, color=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fill:
            cell.fill = fill

    row_no = 1

    for tower, group_data in groups.items():
        for group, df in group_data.items():
            ws.merge_cells(start_row=row_no, start_column=1, end_row=row_no, end_column=len(df.columns))
            apply_style(ws.cell(row_no, 1, f"{tower} - {group}"), group_fill, True)
            row_no += 1

            for col_idx, col_name in enumerate(df.columns, 1):
                apply_style(ws.cell(row_no, col_idx, col_name), header_fill, True, "FFFFFF")
            row_no += 1

            for _, r in df.iterrows():
                consultant = str(r.get("Consultant", "")).strip()

                for col_idx, col_name in enumerate(df.columns, 1):
                    value = str(r.get(col_name, ""))
                    fill = None
                    bold = False

                    if col_name == "Specification":
                        fill = spec_fill
                        bold = True

                    if col_name in vendors and consultant:
                        vendor_value = str(r.get(col_name, "")).strip()
                        if vendor_value and normalize(vendor_value) != normalize(consultant):
                            fill = deviation_fill

                    apply_style(ws.cell(row_no, col_idx, value), fill, bold)

                row_no += 1

            row_no += 2

    def write_section(title, df, row_no):
        if df.empty:
            return row_no

        ws.merge_cells(start_row=row_no, start_column=1, end_row=row_no, end_column=len(df.columns))
        apply_style(ws.cell(row_no, 1, title), group_fill, True)
        row_no += 1

        for col_idx, col_name in enumerate(df.columns, 1):
            apply_style(ws.cell(row_no, col_idx, col_name), header_fill, True, "FFFFFF")
        row_no += 1

        for _, r in df.iterrows():
            for col_idx, col_name in enumerate(df.columns, 1):
                apply_style(ws.cell(row_no, col_idx, str(r.get(col_name, ""))))
            row_no += 1

        return row_no + 2

    row_no = write_section("PRICING SUMMARY", commercial_data["pricing"], row_no)
    row_no = write_section("PAYMENT TERMS", commercial_data["payment"], row_no)
    row_no = write_section("DELIVERY PROGRAM", commercial_data["delivery"], row_no)
    row_no = write_section("NOTES", commercial_data["notes"], row_no)

    for col_idx in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 24
    ws.column_dimensions["A"].width = 32

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.title("🛗 Radiant Bridges Pro Hub")

if not OPENPYXL_AVAILABLE:
    st.warning("openpyxl is not installed. Excel will be exported as CSV fallback.")


with st.sidebar:
    st.header("Vendor Names")

    vendor_text = st.text_area(
        "Edit vendor names, one per line",
        "\n".join(st.session_state.vendors),
        height=160,
    )

    if st.button("Update Vendors"):
        new_vendors = [v.strip().upper() for v in vendor_text.splitlines() if v.strip()]
        if new_vendors:
            st.session_state.vendors = new_vendors
            sync_vendor_columns()
            st.rerun()

    st.divider()
    st.header("Tower / Group Setup")

    tower_names = list(st.session_state.groups.keys())
    selected_tower = st.selectbox("Select Tower", tower_names)

    new_tower_name = st.text_input("Add new tower name")

    if st.button("Add Tower"):
        if new_tower_name.strip():
            st.session_state.groups[new_tower_name.strip()] = {
                "Group 1": make_default_df(st.session_state.vendors)
            }
            st.rerun()

    rename_tower = st.text_input("Rename selected tower", selected_tower)

    if st.button("Rename Tower"):
        if rename_tower.strip() and rename_tower.strip() != selected_tower:
            st.session_state.groups[rename_tower.strip()] = st.session_state.groups.pop(selected_tower)
            st.rerun()

    if st.button("Remove Selected Tower"):
        if len(st.session_state.groups) > 1:
            del st.session_state.groups[selected_tower]
            st.rerun()
        else:
            st.warning("At least one tower is required.")

    group_names = list(st.session_state.groups[selected_tower].keys())
    selected_group = st.selectbox("Select Group", group_names)

    new_group_name = st.text_input("Add new group name")

    if st.button("Add Group"):
        if new_group_name.strip():
            st.session_state.groups[selected_tower][new_group_name.strip()] = make_default_df(st.session_state.vendors)
            st.rerun()

    rename_group = st.text_input("Rename selected group", selected_group)

    if st.button("Rename Group"):
        if rename_group.strip() and rename_group.strip() != selected_group:
            st.session_state.groups[selected_tower][rename_group.strip()] = st.session_state.groups[selected_tower].pop(selected_group)
            st.rerun()

    if st.button("Remove Selected Group"):
        if len(st.session_state.groups[selected_tower]) > 1:
            del st.session_state.groups[selected_tower][selected_group]
            st.rerun()
        else:
            st.warning("At least one group is required.")


st.subheader(f"Technical Comparison - {selected_tower} / {selected_group}")

current_df = st.session_state.groups[selected_tower][selected_group]

edited_df = st.data_editor(
    current_df,
    num_rows="dynamic",
    use_container_width=True,
    key=f"editor_{selected_tower}_{selected_group}",
)

st.session_state.groups[selected_tower][selected_group] = edited_df

st.markdown("### Highlighted Comparison")
st.caption("Yellow cells show vendor deviation from consultant specification.")

st.dataframe(
    edited_df.style.apply(highlight_main_table, axis=1),
    use_container_width=True,
)


st.divider()
st.subheader("Upload Specifications / Tender Offers")

uploaded_files = st.file_uploader(
    "Upload specs, tender offers, vendor quotations, or supporting files",
    type=["pdf", "xlsx", "xls", "csv", "docx"],
    accept_multiple_files=True,
)

if uploaded_files:
    for f in uploaded_files:
        st.session_state.attachments.append(
            {"name": f.name, "size": f.size, "type": f.type}
        )

if st.session_state.attachments:
    st.write("Uploaded Files")
    for item in st.session_state.attachments:
        st.write(f"📎 {item['name']} - {round(item['size']/1024, 1)} KB")


st.divider()
st.subheader("Pricing Summary")

st.session_state.commercial_data["pricing"] = st.data_editor(
    st.session_state.commercial_data["pricing"],
    num_rows="dynamic",
    use_container_width=True,
    key="pricing_editor",
)

st.subheader("Payment Terms")

st.session_state.commercial_data["payment"] = st.data_editor(
    st.session_state.commercial_data["payment"],
    num_rows="dynamic",
    use_container_width=True,
    key="payment_editor",
)

st.subheader("Delivery Program")

st.session_state.commercial_data["delivery"] = st.data_editor(
    st.session_state.commercial_data["delivery"],
    num_rows="dynamic",
    use_container_width=True,
    key="delivery_editor",
)

st.subheader("Notes")

st.session_state.commercial_data["notes"] = st.data_editor(
    st.session_state.commercial_data["notes"],
    num_rows="dynamic",
    use_container_width=True,
    key="notes_editor",
)


st.divider()
st.subheader("Generate Excel")

excel_data = build_excel(
    st.session_state.groups,
    st.session_state.commercial_data,
    st.session_state.vendors,
)

if OPENPYXL_AVAILABLE:
    st.download_button(
        "📥 Generate / Download Excel Report",
        data=excel_data,
        file_name="Radiant_Bridges_Comparison_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
else:
    st.download_button(
        "📥 Generate / Download CSV Report",
        data=excel_data,
        file_name="Radiant_Bridges_Comparison_Report.csv",
        mime="text/csv",
        use_container_width=True,
        type="primary",
    )
