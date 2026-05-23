import io
import json
from datetime import date
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except Exception:
    PDF_AVAILABLE = False


st.set_page_config(
    page_title="Radiant Bridges Dynamic Comparison App",
    page_icon="🛗",
    layout="wide",
)

SPEC_LIST = [
    "CAPACITY", "SPEED", "DOOR TYPE", "DOOR SIZE (W X H)",
    "SHAFT SIZE (W X D)", "CABIN SIZE (W X D X H)",
    "OVER HEAD HEIGHT", "PIT DEPTH", "NO. OF LIFTS",
    "Lift code", "Machine location", "Operation", "No. of Stops",
    "Travel Height", "Car wall", "Front Wall", "Ceiling", "Mirror",
    "Hand rail", "Skirting", "Decoration", "Door Material",
    "Sill Material", "COP Panel", "LOP",
    "Landing Jamb In Ground Floor", "Landing Jamb In Other Floors",
    "Hall Indicator", "Made", "Remarks"
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


def make_default_df(vendors: List[str]) -> pd.DataFrame:
    columns = ["Specification", "Consultant"] + vendors
    df = pd.DataFrame("", index=range(len(SPEC_LIST)), columns=columns)
    df["Specification"] = SPEC_LIST
    return df


def sync_vendor_columns(df: pd.DataFrame, vendors: List[str]) -> pd.DataFrame:
    required_cols = ["Specification", "Consultant"] + vendors

    if "Specification" not in df.columns:
        df.insert(0, "Specification", "")

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
        st.session_state.groups = {
            "Tower A": {
                "PL1 & PL2": make_default_df(st.session_state.vendors)
            }
        }

    if "active_tower" not in st.session_state:
        st.session_state.active_tower = "Tower A"

    if "active_group" not in st.session_state:
        st.session_state.active_group = "PL1 & PL2"

    if "attachments" not in st.session_state:
        st.session_state.attachments = []


def sync_all_groups() -> None:
    for tower in st.session_state.groups:
        for group in st.session_state.groups[tower]:
            st.session_state.groups[tower][group] = sync_vendor_columns(
                st.session_state.groups[tower][group],
                st.session_state.vendors,
            )


def ensure_active_selection() -> None:
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
        st.session_state.active_group = list(
            st.session_state.groups[st.session_state.active_tower].keys()
        )[0]


def clean_text(value: Any) -> str:
    text = str(value).replace("\n", " ").strip()
    return text.encode("latin-1", "ignore").decode("latin-1")


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file).fillna("")
    return pd.read_excel(uploaded_file).fillna("")


def build_formatted_excel(groups, vendors, project_info) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        ws = workbook.add_worksheet("Comparison")
        writer.sheets["Comparison"] = ws

        ws.set_landscape()
        ws.fit_to_pages(1, 0)

        title_fmt = workbook.add_format({
            "bold": True, "font_size": 16, "align": "center",
            "valign": "vcenter", "border": 1, "bg_color": "#D9EAF7"
        })
        subtitle_fmt = workbook.add_format({
            "bold": True, "font_size": 12, "align": "center",
            "valign": "vcenter", "border": 1, "bg_color": "#EAF4DD"
        })
        label_fmt = workbook.add_format({
            "bold": True, "border": 1, "bg_color": "#F2F2F2"
        })
        value_fmt = workbook.add_format({
            "border": 1, "text_wrap": True
        })
        group_fmt = workbook.add_format({
            "bold": True, "font_size": 12, "align": "center",
            "border": 1, "bg_color": "#BDD7EE"
        })
        header_fmt = workbook.add_format({
            "bold": True, "align": "center", "border": 1,
            "bg_color": "#1F4E78", "font_color": "#FFFFFF",
            "text_wrap": True
        })
        spec_fmt = workbook.add_format({
            "bold": True, "border": 1, "bg_color": "#E2F0D9",
            "text_wrap": True, "valign": "top"
        })
        body_fmt = workbook.add_format({
            "border": 1, "text_wrap": True, "valign": "top"
        })

        columns = ["Specification", "Consultant"] + vendors
        last_col = len(columns) - 1

        ws.set_column(0, 0, 30)
        ws.set_column(1, last_col, 24)

        row = 0

        ws.merge_range(
            row, 0, row, last_col,
            project_info.get("document_title", "ELEVATOR TECHNICAL COMPARISON"),
            title_fmt
        )
        row += 1

        ws.merge_range(
            row, 0, row, last_col,
            project_info.get("project", "RADIANT BRIDGES PROJECT"),
            subtitle_fmt
        )
        row += 2

        info_rows = [
            ("Client", project_info.get("client", ""), "Date", project_info.get("comparison_date", "")),
            ("Main Contractor", project_info.get("main_contractor", ""), "Revision", project_info.get("revision", "")),
            ("Material / Work", project_info.get("material_work", ""), "PR No.", project_info.get("pr_no", "")),
        ]

        mid = max(2, last_col // 2)

        for left_label, left_value, right_label, right_value in info_rows:
            ws.write(row, 0, left_label, label_fmt)
            ws.merge_range(row, 1, row, mid, left_value, value_fmt)
            ws.write(row, mid + 1, right_label, label_fmt)
            ws.merge_range(row, mid + 2, row, last_col, right_value, value_fmt)
            row += 1

        row += 1

        for tower, group_data in groups.items():
            for group, df in group_data.items():
                df = sync_vendor_columns(df, vendors)

                ws.merge_range(row, 0, row, last_col, f"{tower} - {group}", group_fmt)
                row += 1

                for col_idx, col_name in enumerate(columns):
                    ws.write(row, col_idx, col_name, header_fmt)
                row += 1

                for _, data_row in df.iterrows():
                    ws.write(row, 0, data_row.get("Specification", ""), spec_fmt)

                    for col_idx, col_name in enumerate(columns[1:], start=1):
                        ws.write(row, col_idx, str(data_row.get(col_name, "")), body_fmt)

                    row += 1

                row += 2

    output.seek(0)
    return output.getvalue()


def build_pdf_report(groups, vendors, project_info) -> bytes:
    if not PDF_AVAILABLE:
        raise RuntimeError("PDF package is not installed. Add fpdf2==2.7.9 to requirements.txt")

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, clean_text(project_info.get("document_title", "")), ln=True, align="C")

    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, clean_text(project_info.get("project", "")), ln=True, align="C")
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

    for tower, group_data in groups.items():
        for group, df in group_data.items():
            df = sync_vendor_columns(df, vendors)

            pdf.add_page()
            pdf.set_font("Arial", "B", 12)
            pdf.cell(
                0, 8,
                clean_text(f"CONSULTANT SPECIFICATION - {tower} - {group}"),
                ln=True,
                align="C",
            )
            pdf.ln(2)

            pdf.set_font("Arial", "B", 8)
            pdf.cell(70, 7, "Specification", border=1, align="C")
            pdf.cell(200, 7, "Consultant Specification", border=1, align="C")
            pdf.ln()

            pdf.set_font("Arial", "", 7)

            for _, row in df.iterrows():
                if pdf.get_y() > 185:
                    pdf.add_page()

                y = pdf.get_y()
                pdf.multi_cell(70, 5, clean_text(row.get("Specification", ""))[:60], border=1)
                pdf.set_xy(80, y)
                pdf.multi_cell(200, 5, clean_text(row.get("Consultant", ""))[:160], border=1)
                pdf.set_y(max(pdf.get_y(), y + 8))

    for vendor in vendors:
        pdf.add_page()
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, clean_text(f"VENDOR OFFER - {vendor}"), ln=True, align="C")
        pdf.ln(2)

        for tower, group_data in groups.items():
            for group, df in group_data.items():
                df = sync_vendor_columns(df, vendors)

                pdf.set_font("Arial", "B", 10)
                pdf.cell(0, 7, clean_text(f"{tower} - {group}"), ln=True)

                pdf.set_font("Arial", "B", 8)
                pdf.cell(70, 7, "Specification", border=1, align="C")
                pdf.cell(200, 7, clean_text(vendor), border=1, align="C")
                pdf.ln()

                pdf.set_font("Arial", "", 7)

                for _, row in df.iterrows():
                    if pdf.get_y() > 185:
                        pdf.add_page()

                    y = pdf.get_y()
                    pdf.multi_cell(70, 5, clean_text(row.get("Specification", ""))[:60], border=1)
                    pdf.set_xy(80, y)
                    pdf.multi_cell(200, 5, clean_text(row.get(vendor, ""))[:160], border=1)
                    pdf.set_y(max(pdf.get_y(), y + 8))

                pdf.ln(3)

    for tower, group_data in groups.items():
        for group, df in group_data.items():
            df = sync_vendor_columns(df, vendors)
            columns = ["Specification", "Consultant"] + vendors

            pdf.add_page()
            pdf.set_font("Arial", "B", 12)
            pdf.cell(
                0, 8,
                clean_text(f"FULL COMPARISON - {tower} - {group}"),
                ln=True,
                align="C",
            )
            pdf.ln(2)

            usable_width = 277
            spec_width = 45
            other_width = (usable_width - spec_width) / max(1, len(columns) - 1)
            widths = [spec_width] + [other_width] * (len(columns) - 1)

            pdf.set_font("Arial", "B", 7)
            for i, col in enumerate(columns):
                pdf.cell(widths[i], 7, clean_text(col)[:22], border=1, align="C")
            pdf.ln()

            pdf.set_font("Arial", "", 6.5)

            for _, row in df.iterrows():
                if pdf.get_y() > 185:
                    pdf.add_page()

                y = pdf.get_y()
                x = pdf.get_x()

                for i, col in enumerate(columns):
                    pdf.set_xy(x + sum(widths[:i]), y)
                    pdf.multi_cell(widths[i], 4, clean_text(row.get(col, ""))[:45], border=1)

                pdf.set_y(y + 8)

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


init_state()
sync_all_groups()
ensure_active_selection()

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
    pi["comparison_date"] = st.text_input("Date", pi.get("comparison_date", ""))

    st.divider()
    st.header("🏢 Structure")

    if st.session_state.groups:
        tower_list = list(st.session_state.groups.keys())
        st.session_state.active_tower = st.selectbox(
            "Tower / Section",
            tower_list,
            index=tower_list.index(st.session_state.active_tower)
            if st.session_state.active_tower in tower_list else 0,
        )

        group_list = list(st.session_state.groups[st.session_state.active_tower].keys())

        if group_list:
            st.session_state.active_group = st.selectbox(
                "Group",
                group_list,
                index=group_list.index(st.session_state.active_group)
                if st.session_state.active_group in group_list else 0,
            )
        else:
            st.session_state.active_group = ""
            st.info("No group under this tower. Add a group.")

    new_tower = st.text_input("Add Tower / Section")
    if st.button("Add Tower / Section", use_container_width=True):
        name = new_tower.strip()
        if name and name not in st.session_state.groups:
            st.session_state.groups[name] = {}
            st.session_state.active_tower = name
            st.session_state.active_group = ""
            st.rerun()

    new_group = st.text_input("Add Group")
    if st.button("Add Group", use_container_width=True):
        name = new_group.strip()
        if not st.session_state.active_tower:
            st.warning("Add or select a tower first.")
        elif name and name not in st.session_state.groups[st.session_state.active_tower]:
            st.session_state.groups[st.session_state.active_tower][name] = make_default_df(
                st.session_state.vendors
            )
            st.session_state.active_group = name
            st.rerun()

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Remove Group", use_container_width=True):
            if st.session_state.active_tower and st.session_state.active_group:
                del st.session_state.groups[st.session_state.active_tower][st.session_state.active_group]
                ensure_active_selection()
                st.rerun()

    with c2:
        if st.button("Remove Tower", use_container_width=True):
            if st.session_state.active_tower:
                del st.session_state.groups[st.session_state.active_tower]
                ensure_active_selection()
                st.rerun()

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
        a, b = st.columns([4, 1])
        a.write(vendor)
        if b.button("❌", key=f"delete_vendor_{vendor}"):
            st.session_state.vendors.remove(vendor)
            sync_all_groups()
            st.rerun()


st.title("🛗 Radiant Bridges Dynamic Comparison App")
st.caption("Consultant specification, vendor offer comparison, Excel/PDF export, and attachment management.")

if not PDF_AVAILABLE:
    st.warning("PDF export package is missing. Add `fpdf2==2.7.9` to requirements.txt and reboot the app.")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Towers", len(st.session_state.groups))
m2.metric("Groups", sum(len(x) for x in st.session_state.groups.values()))
m3.metric("Vendors", len(st.session_state.vendors))
m4.metric("Active", f"{st.session_state.active_tower or '-'} / {st.session_state.active_group or '-'}")

st.divider()

if not st.session_state.groups:
    st.warning("No tower/section available. Add one from sidebar.")
    st.stop()

if not st.session_state.active_tower or not st.session_state.groups.get(st.session_state.active_tower):
    st.warning("Please add at least one group under selected tower/section.")
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

st.session_state.groups[active_tower][active_group] = sync_vendor_columns(
    edited_df,
    st.session_state.vendors,
)

st.divider()
st.subheader("📎 Attachments")

attachment_uploads = st.file_uploader(
    "Upload attachments",
    type=["pdf", "xlsx", "xls", "csv", "docx", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

if attachment_uploads:
    existing = {x["name"] for x in st.session_state.attachments}
    for file in attachment_uploads:
        if file.name not in existing:
            st.session_state.attachments.append({
                "name": file.name,
                "type": file.type,
                "size": file.size,
                "bytes": file.getvalue(),
            })

if st.session_state.attachments:
    for i, item in enumerate(list(st.session_state.attachments)):
        a, b, c = st.columns([5, 2, 1])
        a.write(f"📄 {item['name']}")
        b.write(f"{round(item['size'] / 1024, 1)} KB")
        if c.button("Remove", key=f"remove_attachment_{i}"):
            st.session_state.attachments.pop(i)
            st.rerun()

    if st.button("Remove All Attachments"):
        st.session_state.attachments = []
        st.rerun()
else:
    st.info("No attachments uploaded.")

st.divider()
st.subheader("📂 Smart Import to Current Group")

uploaded_file = st.file_uploader(
    "Upload Excel/CSV file to import values",
    type=["xlsx", "xls", "csv"],
    key="smart_import",
)

if uploaded_file:
    try:
        raw_df = read_uploaded_file(uploaded_file)
        st.dataframe(raw_df.head(10), use_container_width=True, hide_index=True)

        target_cols = ["Consultant"] + st.session_state.vendors
        options = ["-- Do not import --"] + list(raw_df.columns)

        mapping = {}
        cols = st.columns(min(3, len(target_cols)))

        for i, target in enumerate(target_cols):
            with cols[i % len(cols)]:
                mapping[target] = st.selectbox(
                    f"Map to {target}",
                    options,
                    key=f"map_{target}",
                )

        if st.button("Apply Mapping", type="primary"):
            df = st.session_state.groups[active_tower][active_group].copy()
            max_rows = min(len(df), len(raw_df))

            for target, source in mapping.items():
                if source != "-- Do not import --":
                    df.loc[:max_rows - 1, target] = raw_df[source].astype(str).head(max_rows).tolist()

            st.session_state.groups[active_tower][active_group] = sync_vendor_columns(
                df,
                st.session_state.vendors,
            )
            st.success("Import applied.")
            st.rerun()

    except Exception as exc:
        st.error(f"Unable to read file: {exc}")

st.divider()
st.subheader("📑 Full Preview")

with st.expander("Show all towers and groups"):
    for tower, group_data in st.session_state.groups.items():
        st.markdown(f"### {tower}")
        for group, df in group_data.items():
            st.markdown(f"**{group}**")
            st.dataframe(
                sync_vendor_columns(df, st.session_state.vendors),
                use_container_width=True,
                hide_index=True,
            )

st.divider()
st.subheader("📤 Export")

excel_bytes = build_formatted_excel(
    st.session_state.groups,
    st.session_state.vendors,
    st.session_state.project_info,
)

x1, x2, x3 = st.columns(3)

with x1:
    st.download_button(
        "📥 Download Excel",
        data=excel_bytes,
        file_name="RADIANT_BRIDGES_COMPARISON.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

with x2:
    if PDF_AVAILABLE:
        pdf_bytes = build_pdf_report(
            st.session_state.groups,
            st.session_state.vendors,
            st.session_state.project_info,
        )

        st.download_button(
            "📄 Download PDF",
            data=pdf_bytes,
            file_name="RADIANT_BRIDGES_COMPARISON.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    else:
        st.button("📄 Download PDF", disabled=True, use_container_width=True)

with x3:
    st.download_button(
        "💾 Backup JSON",
        data=export_backup(),
        file_name="radiant_bridges_backup.json",
        mime="application/json",
        use_container_width=True,
    )

with st.expander("Install / Run"):
    st.code(
        """pip install streamlit pandas openpyxl xlsxwriter fpdf2==2.7.9
streamlit run radiant_bridges_dynamic_comparison_app.py""",
        language="bash",
    )
