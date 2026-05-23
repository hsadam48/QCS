import io
import json
from datetime import date
from typing import List

import pandas as pd
import streamlit as st

try:
    import pdfplumber
    PDF_READ_AVAILABLE = True
except Exception:
    PDF_READ_AVAILABLE = False


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


def init_state():
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

    if "editor_version" not in st.session_state:
        st.session_state.editor_version = 0


def sync_all_groups():
    for tower in st.session_state.groups:
        for group in st.session_state.groups[tower]:
            st.session_state.groups[tower][group] = sync_vendor_columns(
                st.session_state.groups[tower][group],
                st.session_state.vendors,
            )


def ensure_active_selection():
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


def normalize_name(text: str) -> str:
    return (
        text.lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace(".", "")
    )


def detect_target_column(file_name: str, vendors: List[str]) -> str:
    name = normalize_name(file_name)

    if "spec" in name or "consultant" in name or "specification" in name:
        return "Consultant"

    for vendor in vendors:
        vendor_key = normalize_name(vendor)
        if vendor_key in name:
            return vendor

    return ""


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file).fillna("")

    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file).fillna("")

    if name.endswith(".pdf"):
        if not PDF_READ_AVAILABLE:
            return pd.DataFrame({"PDF Error": ["pdfplumber is not installed"]})

        rows = []

        with pdfplumber.open(uploaded_file) as pdf:
            for page_no, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()

                for table in tables or []:
                    for row in table:
                        rows.append(row)

                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        rows.append([f"Page {page_no}", line])

        if not rows:
            return pd.DataFrame({"PDF Content": ["No extractable text found"]})

        max_cols = max(len(r) for r in rows)
        rows = [r + [""] * (max_cols - len(r)) for r in rows]

        return pd.DataFrame(
            rows,
            columns=[f"Column {i+1}" for i in range(max_cols)]
        ).fillna("")

    raise ValueError("Unsupported file format")


def extract_combined_rows(raw_df: pd.DataFrame) -> List[str]:
    raw_df = raw_df.fillna("").astype(str)
    combined = raw_df.agg(" ".join, axis=1).tolist()
    return [x.strip() for x in combined if x.strip()]


def auto_fill_from_file(raw_df: pd.DataFrame, file_name: str):
    target_col = detect_target_column(file_name, st.session_state.vendors)

    if not target_col:
        return False, "Could not detect target column. Rename file like Consultant Specification.pdf, KONE Offer.pdf, AG MELCO Offer.pdf."

    active_tower = st.session_state.active_tower
    active_group = st.session_state.active_group

    df = st.session_state.groups[active_tower][active_group].copy()
    df = sync_vendor_columns(df, st.session_state.vendors)
    df = df.reset_index(drop=True)

    extracted_rows = extract_combined_rows(raw_df)

    if not extracted_rows:
        return False, "No readable data found in uploaded file."

    max_rows = min(len(df), len(extracted_rows))

    df.loc[0:max_rows - 1, target_col] = extracted_rows[:max_rows]

    st.session_state.groups[active_tower][active_group] = sync_vendor_columns(
        df,
        st.session_state.vendors,
    )

    st.session_state.editor_version += 1

    return True, f"Auto-filled extracted data under {target_col}."


def build_excel_or_csv(groups, vendors, project_info):
    rows = []

    rows.append([project_info.get("document_title", "")])
    rows.append([project_info.get("project", "")])
    rows.append(["Client", project_info.get("client", "")])
    rows.append(["Main Contractor", project_info.get("main_contractor", "")])
    rows.append(["Material / Work", project_info.get("material_work", "")])
    rows.append(["Revision", project_info.get("revision", "")])
    rows.append(["Date", project_info.get("comparison_date", "")])
    rows.append([])

    columns = ["Specification", "Consultant"] + vendors

    for tower, group_data in groups.items():
        for group, df in group_data.items():
            df = sync_vendor_columns(df, vendors)

            rows.append([f"{tower} - {group}"])
            rows.append(columns)

            for _, r in df.iterrows():
                rows.append([r.get(c, "") for c in columns])

            rows.append([])

    output = io.StringIO()
    pd.DataFrame(rows).to_csv(output, index=False, header=False)
    return output.getvalue().encode("utf-8-sig")


def make_pdf_from_lines(lines: List[str]) -> bytes:
    def esc(text):
        text = str(text).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        return text[:120]

    pages = [lines[i:i + 42] for i in range(0, len(lines), 42)]

    objects = []
    objects.append("<< /Type /Catalog /Pages 2 0 R >>")
    objects.append(
        f"<< /Type /Pages /Kids [{' '.join(f'{3 + i * 2} 0 R' for i in range(len(pages)))}] /Count {len(pages)} >>"
    )

    for page_index, page_lines in enumerate(pages):
        page_obj = 3 + page_index * 2
        content_obj = page_obj + 1

        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 842 595] "
            f"/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> "
            f"/Contents {content_obj} 0 R >>"
        )

        y = 560
        content = ["BT /F1 9 Tf"]

        for line in page_lines:
            content.append(f"40 {y} Td ({esc(line)}) Tj")
            content.append("0 -12 Td")
            y -= 12

        content.append("ET")

        stream = "\n".join(content)

        objects.append(
            f"<< /Length {len(stream.encode('latin-1', 'ignore'))} >>\n"
            f"stream\n{stream}\nendstream"
        )

    pdf = "%PDF-1.4\n"
    offsets = [0]

    for i, obj in enumerate(objects, start=1):
        offsets.append(len(pdf.encode("latin-1")))
        pdf += f"{i} 0 obj\n{obj}\nendobj\n"

    xref_pos = len(pdf.encode("latin-1"))
    pdf += f"xref\n0 {len(objects) + 1}\n"
    pdf += "0000000000 65535 f \n"

    for offset in offsets[1:]:
        pdf += f"{offset:010d} 00000 n \n"

    pdf += f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF"

    return pdf.encode("latin-1", "ignore")


def build_simple_pdf(groups, vendors, project_info) -> bytes:
    lines = []

    lines.append(project_info.get("document_title", "ELEVATOR TECHNICAL COMPARISON"))
    lines.append(project_info.get("project", "RADIANT BRIDGES PROJECT"))
    lines.append("")
    lines.append(f"Client: {project_info.get('client', '')}")
    lines.append(f"Main Contractor: {project_info.get('main_contractor', '')}")
    lines.append(f"Material / Work: {project_info.get('material_work', '')}")
    lines.append(f"Revision: {project_info.get('revision', '')}")
    lines.append(f"Date: {project_info.get('comparison_date', '')}")
    lines.append("")

    for tower, group_data in groups.items():
        for group, df in group_data.items():
            df = sync_vendor_columns(df, vendors)

            lines.append("=" * 90)
            lines.append(f"CONSULTANT SPECIFICATION - {tower} - {group}")
            lines.append("=" * 90)

            for _, r in df.iterrows():
                lines.append(f"{r.get('Specification', '')}: {r.get('Consultant', '')}")

            lines.append("")

            for vendor in vendors:
                lines.append("-" * 90)
                lines.append(f"VENDOR OFFER - {vendor} - {tower} - {group}")
                lines.append("-" * 90)

                for _, r in df.iterrows():
                    lines.append(f"{r.get('Specification', '')}: {r.get(vendor, '')}")

                lines.append("")

    return make_pdf_from_lines(lines)


def export_backup():
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
        st.session_state.active_tower = st.selectbox("Tower / Section", tower_list)

        group_list = list(st.session_state.groups[st.session_state.active_tower].keys())

        if group_list:
            st.session_state.active_group = st.selectbox("Group", group_list)
        else:
            st.session_state.active_group = ""

    new_tower = st.text_input("Add Tower / Section")

    if st.button("Add Tower / Section"):
        name = new_tower.strip()

        if name and name not in st.session_state.groups:
            st.session_state.groups[name] = {}
            st.session_state.active_tower = name
            st.rerun()

    new_group = st.text_input("Add Group")

    if st.button("Add Group"):
        name = new_group.strip()

        if st.session_state.active_tower and name:
            st.session_state.groups[st.session_state.active_tower][name] = make_default_df(
                st.session_state.vendors
            )
            st.session_state.active_group = name
            st.session_state.editor_version += 1
            st.rerun()

    c1, c2 = st.columns(2)

    if c1.button("Remove Group"):
        if st.session_state.active_tower and st.session_state.active_group:
            del st.session_state.groups[st.session_state.active_tower][st.session_state.active_group]
            ensure_active_selection()
            st.session_state.editor_version += 1
            st.rerun()

    if c2.button("Remove Tower"):
        if st.session_state.active_tower:
            del st.session_state.groups[st.session_state.active_tower]
            ensure_active_selection()
            st.session_state.editor_version += 1
            st.rerun()

    st.divider()
    st.header("🏷️ Vendors")

    new_vendor = st.text_input("Add Vendor")

    if st.button("Add Vendor"):
        name = new_vendor.strip().upper()

        if name and name not in st.session_state.vendors:
            st.session_state.vendors.append(name)
            sync_all_groups()
            st.session_state.editor_version += 1
            st.rerun()

    for vendor in list(st.session_state.vendors):
        a, b = st.columns([4, 1])
        a.write(vendor)

        if b.button("❌", key=f"del_{vendor}"):
            st.session_state.vendors.remove(vendor)
            sync_all_groups()
            st.session_state.editor_version += 1
            st.rerun()


st.title("🛗 Radiant Bridges Dynamic Comparison App")
st.caption("Auto-fills Consultant or Vendor data based on uploaded file name.")

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
    key=f"editor_{active_tower}_{active_group}_{st.session_state.editor_version}",
)

st.session_state.groups[active_tower][active_group] = sync_vendor_columns(
    edited_df,
    st.session_state.vendors,
)

st.divider()
st.subheader("📎 Upload Specification / Vendor Offer")

attachment_uploads = st.file_uploader(
    "Upload Consultant Specification or Vendor Offer",
    type=["pdf", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
    key="main_attachment_upload",
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
    st.write("Uploaded files:")

    for i, item in enumerate(list(st.session_state.attachments)):
        target = detect_target_column(item["name"], st.session_state.vendors)
        target_label = target if target else "Not detected"

        a, b, c, d = st.columns([5, 2, 2, 1])
        a.write(f"📄 {item['name']}")
        b.write(f"{round(item['size'] / 1024, 1)} KB")
        c.write(f"Target: **{target_label}**")

        if d.button("Remove", key=f"remove_{i}"):
            st.session_state.attachments.pop(i)
            st.rerun()

    st.divider()
    st.subheader("⚡ Auto Fill Data")

    selected_name = st.selectbox(
        "Select file to auto-fill",
        [x["name"] for x in st.session_state.attachments],
    )

    selected = next(x for x in st.session_state.attachments if x["name"] == selected_name)

    file_obj = io.BytesIO(selected["bytes"])
    file_obj.name = selected["name"]

    try:
        raw_df = read_uploaded_file(file_obj)
        raw_df = raw_df.reset_index(drop=True).fillna("")

        detected_target = detect_target_column(selected_name, st.session_state.vendors)

        if detected_target:
            st.success(f"Detected target column: {detected_target}")
        else:
            st.warning("Could not detect target column from file name.")

        with st.expander("Preview extracted data"):
            st.dataframe(raw_df.head(30), use_container_width=True, hide_index=True)

        if st.button("Auto Fill to Detected Column", type="primary"):
            ok, msg = auto_fill_from_file(raw_df, selected_name)

            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.warning(msg)

        st.caption("File name examples: Consultant Specification.pdf, KONE Offer.pdf, TKE Offer.xlsx, EEE Offer.pdf, AG MELCO Offer.pdf")

    except Exception as exc:
        st.error(f"Unable to read selected file: {exc}")

else:
    st.info("Upload Consultant Specification or Vendor Offer files above.")

st.divider()
st.subheader("📤 Export")

try:
    excel_bytes = build_excel_or_csv(
        st.session_state.groups,
        st.session_state.vendors,
        st.session_state.project_info,
    )

    st.download_button(
        "📥 Download Excel-Compatible CSV",
        data=excel_bytes,
        file_name="RADIANT_BRIDGES_COMPARISON.csv",
        mime="text/csv",
        use_container_width=True,
        type="primary",
    )

except Exception as exc:
    st.warning(f"Excel/CSV export unavailable: {exc}")

try:
    pdf_bytes = build_simple_pdf(
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

except Exception as exc:
    st.warning(f"PDF export unavailable: {exc}")

st.download_button(
    "💾 Backup JSON",
    data=export_backup(),
    file_name="radiant_bridges_backup.json",
    mime="application/json",
    use_container_width=True,
)

with st.expander("Requirements"):
    st.code(
        """streamlit
pandas
pdfplumber
openpyxl""",
        language="text",
    )
