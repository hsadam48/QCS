import io
import json
from datetime import date
from typing import Dict, List, Any

import pandas as pd
import streamlit as st
from fpdf import FPDF

# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Radiant Bridges - Pro Hub",
    page_icon="🛗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# DEFAULT DATA
# ============================================================
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
    "Landing Jamp In Ground Floor",
    "Landing Jamp In Other Floors",
    "Hall Indicator",
    "Made",
    "Remarks",
]

DEFAULT_VENDORS = ["KONE", "TKE", "EEE", "AG MELCO"]

DEFAULT_PROJECT_INFO = {
    "project": "RADIANT BRIDGES PROJECT",
    "document_title": "ELEVATOR TECHNICAL COMPARISON",
    "client": "Radiant Real Estate",
    "main_contractor": "ATGC Transport & General Contracting L.L.C.-SPC",
    "material_work": "Elevators",
    "pr_no": "",
    "revision": "Rev. 0",
    "comparison_date": date.today().isoformat(),
}

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def make_default_df(vendors: List[str]) -> pd.DataFrame:
    columns = ["Specification", "Consultant"] + vendors
    df = pd.DataFrame("", index=range(len(SPEC_LIST)), columns=columns)
    df["Specification"] = SPEC_LIST
    return df


def sync_vendor_columns(df: pd.DataFrame, vendors: List[str]) -> pd.DataFrame:
    required_cols = ["Specification", "Consultant"] + vendors

    if "Specification" not in df.columns:
        df.insert(0, "Specification", SPEC_LIST[: len(df)])

    if "Consultant" not in df.columns:
        df.insert(1, "Consultant", "")

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    return df[required_cols].fillna("").copy()


def init_state() -> None:
    if "vendors" not in st.session_state:
        st.session_state.vendors = DEFAULT_VENDORS.copy()

    if "project_info" not in st.session_state:
        st.session_state.project_info = DEFAULT_PROJECT_INFO.copy()

    if "groups" not in st.session_state:
        st.session_state.groups = {}

    if "attachments" not in st.session_state:
        st.session_state.attachments = []

    if "active_tower" not in st.session_state:
        st.session_state.active_tower = ""

    if "active_group" not in st.session_state:
        st.session_state.active_group = ""


def sync_all_groups() -> None:
    for tower in st.session_state.groups:
        for group in st.session_state.groups[tower]:
            st.session_state.groups[tower][group] = sync_vendor_columns(
                st.session_state.groups[tower][group],
                st.session_state.vendors,
            )


def ensure_valid_active_selection() -> None:
    """Keep active tower/group valid even when structure is empty."""
    if not st.session_state.groups:
        st.session_state.active_tower = ""
        st.session_state.active_group = ""
        return

    if st.session_state.active_tower not in st.session_state.groups:
        st.session_state.active_tower = list(st.session_state.groups.keys())[0]

    if not st.session_state.groups[st.session_state.active_tower]:
        st.session_state.active_group = ""
        return

    if st.session_state.active_group not in st.session_state.groups[st.session_state.active_tower]:
        st.session_state.active_group = list(st.session_state.groups[st.session_state.active_tower].keys())[0]


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file).fillna("")
    return pd.read_excel(uploaded_file).fillna("")


def build_formatted_excel(groups: Dict[str, Dict[str, pd.DataFrame]], vendors: List[str], project_info: Dict[str, Any]) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Comparison")
        writer.sheets["Comparison"] = worksheet

        worksheet.set_landscape()
        worksheet.set_paper(9)
        worksheet.fit_to_pages(1, 0)
        worksheet.set_margins(left=0.25, right=0.25, top=0.5, bottom=0.5)

        title_fmt = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#D9EAF7",
        })
        subtitle_fmt = workbook.add_format({
            "bold": True,
            "font_size": 12,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#EAF4DD",
        })
        label_fmt = workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#F2F2F2",
            "valign": "vcenter",
        })
        value_fmt = workbook.add_format({
            "border": 1,
            "valign": "vcenter",
            "text_wrap": True,
        })
        group_fmt = workbook.add_format({
            "bold": True,
            "font_size": 12,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#BDD7EE",
        })
        header_fmt = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#1F4E78",
            "font_color": "#FFFFFF",
            "text_wrap": True,
        })
        spec_fmt = workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#E2F0D9",
            "text_wrap": True,
            "valign": "top",
        })
        body_fmt = workbook.add_format({
            "border": 1,
            "text_wrap": True,
            "valign": "top",
        })
        note_fmt = workbook.add_format({
            "italic": True,
            "font_color": "#666666",
            "text_wrap": True,
        })

        columns = ["Specification", "Consultant"] + vendors
        last_col = len(columns) - 1

        worksheet.set_column(0, 0, 30)
        worksheet.set_column(1, last_col, 24)

        row = 0
        worksheet.merge_range(row, 0, row, last_col, project_info.get("document_title", "ELEVATOR TECHNICAL COMPARISON"), title_fmt)
        row += 1
        worksheet.merge_range(row, 0, row, last_col, project_info.get("project", "RADIANT BRIDGES PROJECT"), subtitle_fmt)
        row += 2

        info_rows = [
            ("Client", project_info.get("client", ""), "Date", project_info.get("comparison_date", "")),
            ("Main Contractor", project_info.get("main_contractor", ""), "Revision", project_info.get("revision", "")),
            ("Material / Work", project_info.get("material_work", ""), "PR No.", project_info.get("pr_no", "")),
        ]

        mid = max(2, last_col // 2)
        for left_label, left_value, right_label, right_value in info_rows:
            worksheet.write(row, 0, left_label, label_fmt)
            worksheet.merge_range(row, 1, row, mid, left_value, value_fmt)
            worksheet.write(row, mid + 1, right_label, label_fmt)
            worksheet.merge_range(row, mid + 2, row, last_col, right_value, value_fmt)
            row += 1

        row += 1

        for tower_name, tower_groups in groups.items():
            for group_name, df in tower_groups.items():
                df = sync_vendor_columns(df, vendors)

                worksheet.merge_range(row, 0, row, last_col, f"{tower_name} - {group_name}", group_fmt)
                worksheet.set_row(row, 24)
                row += 1

                for col_idx, col_name in enumerate(columns):
                    worksheet.write(row, col_idx, col_name, header_fmt)
                worksheet.set_row(row, 28)
                row += 1

                for _, data_row in df.iterrows():
                    worksheet.write(row, 0, data_row.get("Specification", ""), spec_fmt)
                    for col_idx, col_name in enumerate(columns[1:], start=1):
                        worksheet.write(row, col_idx, str(data_row.get(col_name, "")), body_fmt)
                    worksheet.set_row(row, 32)
                    row += 1

                row += 2

        worksheet.merge_range(
            row,
            0,
            row + 1,
            last_col,
            "Note: This comparison is generated from the Radiant Bridges Pro Hub. Please verify all values against official vendor offers, approved drawings, and project specifications before final approval.",
            note_fmt,
        )

        worksheet.freeze_panes(8, 1)

    output.seek(0)
    return output.getvalue()


def build_pdf_report(groups: Dict[str, Dict[str, pd.DataFrame]], vendors: List[str], project_info: Dict[str, Any]) -> bytes:
    """Build PDF report with Consultant Specification and Vendor Offer sections always included."""
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    columns = ["Specification", "Consultant"] + vendors
    usable_width = 277
    spec_width = 50
    consultant_width = 55
    vendor_width = (usable_width - spec_width - consultant_width) / max(1, len(vendors))
    col_widths = [spec_width, consultant_width] + [vendor_width] * len(vendors)

    def clean_text(value: Any) -> str:
        text = str(value).replace("
", " ").strip()
        return text.encode("latin-1", "ignore").decode("latin-1")

    def add_project_header() -> None:
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, clean_text(project_info.get("document_title", "ELEVATOR TECHNICAL COMPARISON")), ln=True, align="C")
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, clean_text(project_info.get("project", "RADIANT BRIDGES PROJECT")), ln=True, align="C")
        pdf.ln(4)

        pdf.set_font("Arial", "", 9)
        pdf.cell(45, 7, "Client", border=1)
        pdf.cell(95, 7, clean_text(project_info.get("client", "")), border=1)
        pdf.cell(35, 7, "Date", border=1)
        pdf.cell(95, 7, clean_text(project_info.get("comparison_date", "")), border=1, ln=True)

        pdf.cell(45, 7, "Main Contractor", border=1)
        pdf.cell(95, 7, clean_text(project_info.get("main_contractor", "")), border=1)
        pdf.cell(35, 7, "Revision", border=1)
        pdf.cell(95, 7, clean_text(project_info.get("revision", "")), border=1, ln=True)

        pdf.cell(45, 7, "Material / Work", border=1)
        pdf.cell(95, 7, clean_text(project_info.get("material_work", "")), border=1)
        pdf.cell(35, 7, "PR No.", border=1)
        pdf.cell(95, 7, clean_text(project_info.get("pr_no", "")), border=1, ln=True)

    def add_section_title(title: str) -> None:
        pdf.ln(4)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, clean_text(title), ln=True, align="C")
        pdf.ln(2)

    def add_table_header() -> None:
        pdf.set_font("Arial", "B", 7)
        for idx, col in enumerate(columns):
            pdf.cell(col_widths[idx], 7, clean_text(col)[:24], border=1, align="C")
        pdf.ln()

    def add_comparison_table(df: pd.DataFrame) -> None:
        pdf.set_font("Arial", "", 6.5)
        for _, data_row in df.iterrows():
            if pdf.get_y() > 185:
                pdf.add_page()
                add_table_header()

            y_start = pdf.get_y()
            x_start = pdf.get_x()
            row_height = 8

            for idx, col in enumerate(columns):
                value = clean_text(data_row.get(col, ""))[:55]
                pdf.set_xy(x_start + sum(col_widths[:idx]), y_start)
                pdf.multi_cell(col_widths[idx], 4, value, border=1)

            pdf.set_y(y_start + row_height)

    add_project_header()

    # SECTION 1: Consultant Specification only
    add_section_title("CONSULTANT SPECIFICATION")
    for tower_name, tower_groups in groups.items():
        for group_name, df in tower_groups.items():
            df = sync_vendor_columns(df, vendors)
            consultant_df = df[["Specification", "Consultant"]].copy()
            consultant_df["Vendor Offer"] = ""

            pdf.add_page()
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 8, clean_text(f"CONSULTANT SPECIFICATION - {tower_name} - {group_name}"), ln=True, align="C")
            pdf.ln(2)

            local_cols = ["Specification", "Consultant"]
            local_widths = [70, 200]

            pdf.set_font("Arial", "B", 8)
            for idx, col in enumerate(local_cols):
                pdf.cell(local_widths[idx], 7, clean_text(col), border=1, align="C")
            pdf.ln()

            pdf.set_font("Arial", "", 7)
            for _, row in consultant_df.iterrows():
                if pdf.get_y() > 185:
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 8)
                    for idx, col in enumerate(local_cols):
                        pdf.cell(local_widths[idx], 7, clean_text(col), border=1, align="C")
                    pdf.ln()
                    pdf.set_font("Arial", "", 7)

                y = pdf.get_y()
                pdf.multi_cell(local_widths[0], 5, clean_text(row.get("Specification", ""))[:60], border=1)
                pdf.set_xy(10 + local_widths[0], y)
                pdf.multi_cell(local_widths[1], 5, clean_text(row.get("Consultant", ""))[:160], border=1)
                pdf.set_y(max(pdf.get_y(), y + 8))

    # SECTION 2: Vendor Offers only
    for vendor in vendors:
        pdf.add_page()
        add_section_title(f"VENDOR OFFER - {vendor}")
        for tower_name, tower_groups in groups.items():
            for group_name, df in tower_groups.items():
                df = sync_vendor_columns(df, vendors)

                pdf.set_font("Arial", "B", 10)
                pdf.cell(0, 7, clean_text(f"{tower_name} - {group_name}"), ln=True)

                pdf.set_font("Arial", "B", 8)
                pdf.cell(70, 7, "Specification", border=1, align="C")
                pdf.cell(200, 7, clean_text(vendor), border=1, align="C")
                pdf.ln()

                pdf.set_font("Arial", "", 7)
                for _, row in df.iterrows():
                    if pdf.get_y() > 185:
                        pdf.add_page()
                        pdf.set_font("Arial", "B", 8)
                        pdf.cell(70, 7, "Specification", border=1, align="C")
                        pdf.cell(200, 7, clean_text(vendor), border=1, align="C")
                        pdf.ln()
                        pdf.set_font("Arial", "", 7)

                    y = pdf.get_y()
                    pdf.multi_cell(70, 5, clean_text(row.get("Specification", ""))[:60], border=1)
                    pdf.set_xy(80, y)
                    pdf.multi_cell(200, 5, clean_text(row.get(vendor, ""))[:160], border=1)
                    pdf.set_y(max(pdf.get_y(), y + 8))

                pdf.ln(3)

    # SECTION 3: Full comparison matrix
    pdf.add_page()
    add_section_title("FULL CONSULTANT SPECIFICATION VS VENDOR OFFER COMPARISON")
    for tower_name, tower_groups in groups.items():
        for group_name, df in tower_groups.items():
            df = sync_vendor_columns(df, vendors)
            pdf.add_page()
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 8, clean_text(f"{tower_name} - {group_name}"), ln=True, align="C")
            pdf.ln(2)
            add_table_header()
            add_comparison_table(df)

    return pdf.output(dest="S").encode("latin-1")


def export_backup() -> bytes:
    data = {
        "vendors": st.session_state.vendors,
        "project_info": st.session_state.project_info,
        "groups": {
            tower: {
                group: df.to_dict(orient="records")
                for group, df in group_data.items()
            }
            for tower, group_data in st.session_state.groups.items()
        },
    }
    return json.dumps(data, indent=2).encode("utf-8")


# ============================================================
# MAIN APP
# ============================================================
init_state()
sync_all_groups()
ensure_valid_active_selection()

with st.sidebar:
    st.header("⚙️ Project Info")
    pi = st.session_state.project_info
    pi["project"] = st.text_input("Project", pi.get("project", ""))
    pi["document_title"] = st.text_input("Document Title", pi.get("document_title", ""))
    pi["client"] = st.text_input("Client", pi.get("client", ""))
    pi["main_contractor"] = st.text_input("Main Contractor", pi.get("main_contractor", ""))
    pi["material_work"] = st.text_input("Material / Work", pi.get("material_work", ""))
    pi["pr_no"] = st.text_input("PR No.", pi.get("pr_no", ""))
    pi["revision"] = st.text_input("Revision", pi.get("revision", ""))
    pi["comparison_date"] = st.text_input("Date", pi.get("comparison_date", date.today().isoformat()))

    st.divider()
    st.header("🏢 General Structure")

    if st.session_state.groups:
        towers = list(st.session_state.groups.keys())
        st.session_state.active_tower = st.selectbox(
            "Tower",
            towers,
            index=towers.index(st.session_state.active_tower) if st.session_state.active_tower in towers else 0,
        )

        group_list = list(st.session_state.groups[st.session_state.active_tower].keys())
        if group_list:
            st.session_state.active_group = st.selectbox(
                "Group",
                group_list,
                index=group_list.index(st.session_state.active_group) if st.session_state.active_group in group_list else 0,
            )
        else:
            st.session_state.active_group = ""
            st.info("No group under this tower. Please add a group.")
    else:
        st.session_state.active_tower = ""
        st.session_state.active_group = ""
        st.info("No tower created yet. Please add a tower.")

    new_tower = st.text_input("Add Tower / Section", placeholder="Example: Tower A / Zone 1 / Package 1")
    if st.button("Add Tower / Section", use_container_width=True):
        name = new_tower.strip()
        if name and name not in st.session_state.groups:
            st.session_state.groups[name] = {}
            st.session_state.active_tower = name
            st.session_state.active_group = ""
            st.rerun()
        elif not name:
            st.warning("Enter a tower/section name.")
        else:
            st.warning("This tower/section already exists.")

    new_group = st.text_input("Add Group", placeholder="Example: PL1 & PL2 / Fireman Lift / Lift Package")
    if st.button("Add Group", use_container_width=True):
        name = new_group.strip()
        if not st.session_state.active_tower:
            st.warning("Please add/select a tower first.")
        elif name and name not in st.session_state.groups[st.session_state.active_tower]:
            st.session_state.groups[st.session_state.active_tower][name] = make_default_df(st.session_state.vendors)
            st.session_state.active_group = name
            st.rerun()
        elif not name:
            st.warning("Enter a group name.")
        else:
            st.warning("This group already exists.")

    c_del1, c_del2 = st.columns(2)
    with c_del1:
        if st.button("Remove Active Group", use_container_width=True):
            if st.session_state.active_tower and st.session_state.active_group:
                del st.session_state.groups[st.session_state.active_tower][st.session_state.active_group]
                ensure_valid_active_selection()
                st.rerun()
            else:
                st.warning("No active group to remove.")

    with c_del2:
        if st.button("Remove Active Tower", use_container_width=True):
            if st.session_state.active_tower:
                del st.session_state.groups[st.session_state.active_tower]
                ensure_valid_active_selection()
                st.rerun()
            else:
                st.warning("No active tower to remove.")

    st.caption("You can remove towers/groups up to the last one. Add new structure anytime.")

    st.divider()
    st.header("🏷️ Vendors")

    new_vendor = st.text_input("Add Vendor")
    if st.button("Add Vendor", use_container_width=True):
        name = new_vendor.strip().upper()
        if name and name not in st.session_state.vendors:
            st.session_state.vendors.append(name)
            sync_all_groups()
            st.rerun()

    for vendor in list(st.session_state.vendors):
        c1, c2 = st.columns([4, 1])
        c1.write(vendor)
        if c2.button("❌", key=f"delete_{vendor}"):
            st.session_state.vendors.remove(vendor)
            sync_all_groups()
            st.rerun()

    st.caption("Vendors can also be removed up to the last one. If no vendor remains, only Specification and Consultant columns will be shown.")

st.title("🛗 Radiant Bridges Comparison Pro Hub")
st.caption("Dynamic technical comparison editor with formatted Excel and PDF export.")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Towers", len(st.session_state.groups))
c2.metric("Groups", sum(len(x) for x in st.session_state.groups.values()))
c3.metric("Vendors", len(st.session_state.vendors))
c4.metric("Active", f"{st.session_state.active_tower or '-'} / {st.session_state.active_group or '-'}")

st.divider()

if not st.session_state.groups:
    st.warning("No structure available. Please add a Tower / Section from the sidebar.")
    st.stop()

if not st.session_state.active_tower or not st.session_state.groups.get(st.session_state.active_tower):
    st.warning("Please add at least one Group under the selected Tower / Section.")
    st.stop()

active_tower = st.session_state.active_tower
active_group = st.session_state.active_group

st.subheader(f"📊 {active_tower} > {active_group}")

current_df = sync_vendor_columns(
    st.session_state.groups[active_tower][active_group],
    st.session_state.vendors,
)

edited_df = st.data_editor(
    current_df,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key=f"editor_{active_tower}_{active_group}",
)

st.session_state.groups[active_tower][active_group] = sync_vendor_columns(edited_df, st.session_state.vendors)

st.divider()
st.subheader("📎 Attachments")
st.caption("Upload consultant specifications, vendor offers, drawings, PDFs, Excel files, or supporting documents. You can remove them anytime.")

attachment_uploads = st.file_uploader(
    "Upload attachments",
    type=["pdf", "xlsx", "xls", "csv", "docx", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
    key="attachment_uploads",
)

if attachment_uploads:
    existing_names = {item["name"] for item in st.session_state.attachments}
    for uploaded in attachment_uploads:
        if uploaded.name not in existing_names:
            st.session_state.attachments.append({
                "name": uploaded.name,
                "type": uploaded.type,
                "size": uploaded.size,
                "bytes": uploaded.getvalue(),
            })

if st.session_state.attachments:
    st.write("Current attachments:")
    for idx, item in enumerate(list(st.session_state.attachments)):
        a1, a2, a3 = st.columns([5, 2, 1])
        a1.write(f"📄 {item['name']}")
        a2.write(f"{round(item['size'] / 1024, 1)} KB")
        if a3.button("Remove", key=f"remove_attachment_{idx}"):
            st.session_state.attachments.pop(idx)
            st.rerun()

    if st.button("Remove All Attachments", use_container_width=True):
        st.session_state.attachments = []
        st.rerun()
else:
    st.info("No attachments uploaded yet.")

st.divider()
st.subheader("📂 Smart Import to Current Group")

uploaded_file = st.file_uploader("Upload vendor offer / comparison file", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        raw_df = read_uploaded_file(uploaded_file)
        st.dataframe(raw_df.head(10), use_container_width=True, hide_index=True)

        options = ["-- Do not import --"] + list(raw_df.columns)
        mapping = {}
        target_columns = ["Consultant"] + st.session_state.vendors

        st.write("Map uploaded columns:")
        cols = st.columns(min(3, len(target_columns)))
        for idx, target in enumerate(target_columns):
            with cols[idx % len(cols)]:
                mapping[target] = st.selectbox(f"Map to {target}", options, key=f"map_{target}")

        if st.button("Apply Mapping", type="primary"):
            df = st.session_state.groups[active_tower][active_group].copy()
            max_rows = min(len(df), len(raw_df))
            for target, source in mapping.items():
                if source != "-- Do not import --":
                    df.loc[: max_rows - 1, target] = raw_df[source].astype(str).head(max_rows).tolist()
            st.session_state.groups[active_tower][active_group] = sync_vendor_columns(df, st.session_state.vendors)
            st.success("Mapping applied successfully.")
            st.rerun()

    except Exception as exc:
        st.error(f"Unable to read file: {exc}")

st.divider()
st.subheader("📑 Full Preview")

with st.expander("Show all towers / groups"):
    for tower, group_data in st.session_state.groups.items():
        st.markdown(f"### {tower}")
        for group, df in group_data.items():
            st.markdown(f"**{group}**")
            st.dataframe(sync_vendor_columns(df, st.session_state.vendors), use_container_width=True, hide_index=True)

st.divider()
st.subheader("📤 Export")

excel_bytes = build_formatted_excel(
    st.session_state.groups,
    st.session_state.vendors,
    st.session_state.project_info,
)

pdf_bytes = build_pdf_report(
    st.session_state.groups,
    st.session_state.vendors,
    st.session_state.project_info,
)

col_x, col_p, col_b = st.columns(3)

with col_x:
    st.download_button(
        "📥 Download Excel",
        data=excel_bytes,
        file_name="RADIANT_BRIDGES_COMPARISON.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

with col_p:
    st.download_button(
        "📄 Download PDF",
        data=pdf_bytes,
        file_name="RADIANT_BRIDGES_COMPARISON.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with col_b:
    st.download_button(
        "💾 Backup JSON",
        data=export_backup(),
        file_name="radiant_bridges_backup.json",
        mime="application/json",
        use_container_width=True,
    )

with st.expander("Install / Run"):
    st.code(
        """pip install streamlit pandas openpyxl xlsxwriter fpdf
streamlit run radiant_bridges_pro_hub.py""",
        language="bash",
    )
