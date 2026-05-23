import io
from typing import List

import pandas as pd
import streamlit as st

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


st.set_page_config(
    page_title="Radiant Bridges Pro Hub",
    page_icon="🛗",
    layout="wide",
)

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
    "Lift code",
    "Machine location",
    "Operation",
    "No. of Stops",
    "Travel Height",
    "Car wall",
    "Front Wall",
    "Ceiling",
    "Mirror",
    "Hand rail",
    "Skirting",
    "Decoration",
    "Door Material",
    "Sill Material",
    "COP Panel",
    "LOP",
    "Landing Jamb In Ground Floor",
    "Landing Jamb In Other Floors",
    "Hall Indicator",
    "Made",
]

SPEC_KEYWORDS = {
    "CAPACITY": ["capacity", "kg", "persons", "passenger"],
    "SPEED": ["speed", "m/s", "rated speed"],
    "DOOR TYPE": ["door type", "opening type", "center opening", "telescopic"],
    "DOOR SIZE (W X H)": ["door size", "door opening", "opening size"],
    "SHAFT SIZE (W X D)": ["shaft size", "hoistway", "well size"],
    "CABIN SIZE (W X D X H)": ["cabin size", "car size", "car internal"],
    "OVER HEAD HEIGHT": ["overhead", "over head", "oh"],
    "PIT DEPTH": ["pit depth", "pit"],
    "NO. OF LIFTS": ["no. of lifts", "number of lifts", "quantity"],
    "Lift code": ["lift code", "elevator code"],
    "Machine location": ["machine location", "machine room"],
    "Operation": ["operation", "control system"],
    "No. of Stops": ["stops", "landings"],
    "Travel Height": ["travel height", "travel"],
    "Car wall": ["car wall", "cabin wall"],
    "Front Wall": ["front wall"],
    "Ceiling": ["ceiling"],
    "Mirror": ["mirror"],
    "Hand rail": ["handrail", "hand rail"],
    "Skirting": ["skirting"],
    "Decoration": ["decoration", "finish"],
    "Door Material": ["door material", "landing door"],
    "Sill Material": ["sill material", "sill"],
    "COP Panel": ["cop", "car operating panel"],
    "LOP": ["lop", "landing operating panel"],
    "Landing Jamb In Ground Floor": ["ground floor jamb", "main floor jamb"],
    "Landing Jamb In Other Floors": ["other floor jamb", "typical floor jamb"],
    "Hall Indicator": ["hall indicator", "indicator"],
    "Made": ["made", "country of origin", "manufacturer"],
}


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

    if "tables" not in st.session_state:
        st.session_state.tables = [
            {
                "name": "Tower A - PL1 & PL2",
                "df": make_default_df(st.session_state.vendors),
            }
        ]

    if "commercial_data" not in st.session_state:
        st.session_state.commercial_data = {
            "pricing": pd.DataFrame(columns=["Description"] + st.session_state.vendors),
            "payment": pd.DataFrame(columns=["Term"] + st.session_state.vendors),
            "delivery": pd.DataFrame(columns=["Milestone"] + st.session_state.vendors),
            "notes": pd.DataFrame(columns=["Note"] + st.session_state.vendors),
        }

    if "editor_version" not in st.session_state:
        st.session_state.editor_version = 0


init_state()


def normalize(value):
    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
        .replace("/", "")
        .replace(":", "")
        .replace(".", "")
        .replace("(", "")
        .replace(")", "")
    )


def clean_line(text):
    return str(text).replace("\n", " ").replace("  ", " ").strip()


def sync_vendor_columns():
    required_cols = ["Specification", "Consultant"] + st.session_state.vendors

    for table in st.session_state.tables:
        df = table["df"]

        for col in required_cols:
            if col not in df.columns:
                df[col] = ""

        table["df"] = df[required_cols]

    for key in st.session_state.commercial_data:
        df = st.session_state.commercial_data[key]
        first_col = df.columns[0] if len(df.columns) else "Description"
        required_cols = [first_col] + st.session_state.vendors

        for col in required_cols:
            if col not in df.columns:
                df[col] = ""

        st.session_state.commercial_data[key] = df[required_cols]


def detect_target_column(file_name: str):
    name = normalize(file_name)

    if "consultant" in name or "spec" in name or "specification" in name:
        return "Consultant"

    for vendor in st.session_state.vendors:
        if normalize(vendor) in name:
            return vendor

    return None


def read_pdf_lines(uploaded_file):
    if not PDF_AVAILABLE:
        return []

    lines = []

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables or []:
                for row in table:
                    row_text = " | ".join([clean_line(x) for x in row if x])
                    if row_text:
                        lines.append(row_text)

            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    line = clean_line(line)
                    if line:
                        lines.append(line)

    return lines


def read_file_lines(uploaded_file):
    name = uploaded_file.name.lower()

    if name.endswith(".pdf"):
        return read_pdf_lines(uploaded_file)

    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file).fillna("")
        return df.astype(str).agg(" | ".join, axis=1).tolist()

    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded_file).fillna("")
        return df.astype(str).agg(" | ".join, axis=1).tolist()

    return []


def find_value_for_spec(lines, spec):
    keywords = SPEC_KEYWORDS.get(spec, [spec])
    keywords = [normalize(k) for k in keywords]

    for line in lines:
        line_clean = normalize(line)

        if any(k in line_clean for k in keywords):
            if "|" in line:
                parts = [p.strip() for p in line.split("|") if p.strip()]

                for idx, part in enumerate(parts):
                    part_clean = normalize(part)

                    if any(k in part_clean for k in keywords):
                        if idx + 1 < len(parts):
                            return parts[idx + 1]

                return parts[-1] if parts else ""

            for sep in [":", "-", "–"]:
                if sep in line:
                    parts = line.split(sep, 1)
                    if len(parts) == 2:
                        return parts[1].strip()

            return line.strip()

    return ""


def auto_fill_exact(rows, target_col, table_index):
    df = st.session_state.tables[table_index]["df"].copy()
    filled = 0

    for idx, row in df.iterrows():
        spec = str(row.get("Specification", "")).strip()

        if not spec:
            continue

        value = find_value_for_spec(rows, spec)

        if value:
            df.at[idx, target_col] = value
            filled += 1

    st.session_state.tables[table_index]["df"] = df
    st.session_state.editor_version += 1

    return filled


def build_excel():
    if not OPENPYXL_AVAILABLE:
        output = io.StringIO()
        rows = []

        for table in st.session_state.tables:
            rows.append([table["name"]])
            rows.append(list(table["df"].columns))

            for _, r in table["df"].iterrows():
                rows.append([r.get(c, "") for c in table["df"].columns])

            rows.append([])

        for title, key in [
            ("PRICING SUMMARY", "pricing"),
            ("PAYMENT TERMS", "payment"),
            ("DELIVERY PROGRAM", "delivery"),
            ("NOTES", "notes"),
        ]:
            df = st.session_state.commercial_data[key]
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
    table_fill = PatternFill("solid", fgColor="BDD7EE")
    spec_fill = PatternFill("solid", fgColor="E2F0D9")
    deviation_fill = PatternFill("solid", fgColor="FDE68A")

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style(cell, fill=None, bold=False, color="000000"):
        cell.border = border
        cell.font = Font(bold=bold, color=color)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        if fill:
            cell.fill = fill

    row_no = 1

    for table in st.session_state.tables:
        df = table["df"]

        ws.merge_cells(
            start_row=row_no,
            start_column=1,
            end_row=row_no,
            end_column=len(df.columns),
        )
        style(ws.cell(row_no, 1, table["name"]), table_fill, True)
        row_no += 1

        for col_idx, col_name in enumerate(df.columns, 1):
            style(ws.cell(row_no, col_idx, col_name), header_fill, True, "FFFFFF")

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

                if col_name in st.session_state.vendors and consultant:
                    vendor_value = str(r.get(col_name, "")).strip()

                    if vendor_value and normalize(vendor_value) != normalize(consultant):
                        fill = deviation_fill

                style(ws.cell(row_no, col_idx, value), fill, bold)

            row_no += 1

        row_no += 2

    for title, key in [
        ("PRICING SUMMARY", "pricing"),
        ("PAYMENT TERMS", "payment"),
        ("DELIVERY PROGRAM", "delivery"),
        ("NOTES", "notes"),
    ]:
        df = st.session_state.commercial_data[key]

        ws.merge_cells(
            start_row=row_no,
            start_column=1,
            end_row=row_no,
            end_column=max(1, len(df.columns)),
        )
        style(ws.cell(row_no, 1, title), table_fill, True)
        row_no += 1

        for col_idx, col_name in enumerate(df.columns, 1):
            style(ws.cell(row_no, col_idx, col_name), header_fill, True, "FFFFFF")

        row_no += 1

        for _, r in df.iterrows():
            for col_idx, col_name in enumerate(df.columns, 1):
                style(ws.cell(row_no, col_idx, str(r.get(col_name, ""))))

            row_no += 1

        row_no += 2

    for col_idx in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 24

    ws.column_dimensions["A"].width = 32

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.title("🛗 Radiant Bridges Pro Hub")

with st.sidebar:
    st.header("Vendor Names")

    vendor_text = st.text_area(
        "Edit vendor names, one per line",
        "\n".join(st.session_state.vendors),
        height=150,
    )

    if st.button("Update Vendors"):
        vendors = [x.strip().upper() for x in vendor_text.splitlines() if x.strip()]

        if vendors:
            st.session_state.vendors = vendors
            sync_vendor_columns()
            st.session_state.editor_version += 1
            st.rerun()


st.subheader("Comparison Tables")

if st.button("➕ Add Tower / Comparison Table", type="primary"):
    new_index = len(st.session_state.tables) + 1

    st.session_state.tables.append(
        {
            "name": f"Tower {new_index}",
            "df": make_default_df(st.session_state.vendors),
        }
    )

    st.session_state.editor_version += 1
    st.rerun()


for i, table in enumerate(st.session_state.tables):
    st.divider()

    col1, col2 = st.columns([5, 1])

    with col1:
        table["name"] = st.text_input(
            "Editable Table / Tower Name",
            table["name"],
            key=f"table_name_{i}_{st.session_state.editor_version}",
        )

    with col2:
        if st.button("Remove", key=f"remove_table_{i}"):
            if len(st.session_state.tables) > 1:
                st.session_state.tables.pop(i)
                st.session_state.editor_version += 1
                st.rerun()
            else:
                st.warning("At least one table is required.")

    st.session_state.tables[i]["df"] = st.data_editor(
        table["df"],
        num_rows="dynamic",
        use_container_width=True,
        key=f"table_editor_{i}_{st.session_state.editor_version}",
    )


st.divider()
st.subheader("Upload Specification / Vendor Offer")

uploaded_files = st.file_uploader(
    "Upload PDF / Excel / CSV files. File name should match Consultant / Spec / Vendor name.",
    type=["pdf", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

if uploaded_files:
    table_names = [table["name"] for table in st.session_state.tables]

    target_table_name = st.selectbox(
        "Select table to auto-fill",
        table_names,
    )

    target_table_index = table_names.index(target_table_name)

    if st.button("Apply Uploaded Files to Selected Table", type="primary"):
        for uploaded_file in uploaded_files:
            target_col = detect_target_column(uploaded_file.name)

            if not target_col:
                st.warning(f"Could not detect target column for {uploaded_file.name}")
                continue

            rows = read_file_lines(uploaded_file)

            if uploaded_file.name.lower().endswith(".pdf") and not PDF_AVAILABLE:
                st.warning("pdfplumber is not installed. PDF cannot be extracted.")
                continue

            filled = auto_fill_exact(rows, target_col, target_table_index)

            if filled:
                st.success(f"{uploaded_file.name}: {filled} values filled under {target_col}")
            else:
                st.warning(f"{uploaded_file.name}: no matching specification values found")

        st.rerun()


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

excel_data = build_excel()

st.download_button(
    "📥 Generate / Download Excel Report",
    data=excel_data,
    file_name="Radiant_Bridges_Comparison_Report.xlsx"
    if OPENPYXL_AVAILABLE
    else "Radiant_Bridges_Comparison_Report.csv",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if OPENPYXL_AVAILABLE
    else "text/csv",
    use_container_width=True,
    type="primary",
)
