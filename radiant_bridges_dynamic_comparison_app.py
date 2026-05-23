uploaded_files = st.file_uploader(
    "Upload PDF / Excel / CSV files",
    type=["pdf", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

if uploaded_files:
    selected_table_names = [table["name"] for table in st.session_state.tables]

    target_table_name = st.selectbox(
        "Select table to auto-fill",
        selected_table_names,
    )

    target_table_index = selected_table_names.index(target_table_name)

    for uploaded_file in uploaded_files:
        st.write(f"Processing: {uploaded_file.name}")

        file_name = uploaded_file.name.upper()

        if "CONSULTANT" in file_name or "SPEC" in file_name:
            target_col = "Consultant"
        else:
            target_col = None
            for vendor in st.session_state.vendors:
                if vendor.upper().replace(" ", "") in file_name.replace(" ", ""):
                    target_col = vendor
                    break

        if target_col is None:
            st.warning(f"Could not detect target column for {uploaded_file.name}")
            continue

        if uploaded_file.name.lower().endswith(".csv"):
            raw_df = pd.read_csv(uploaded_file).fillna("")
            extracted_rows = raw_df.astype(str).agg(" ".join, axis=1).tolist()

        elif uploaded_file.name.lower().endswith((".xlsx", ".xls")):
            raw_df = pd.read_excel(uploaded_file).fillna("")
            extracted_rows = raw_df.astype(str).agg(" ".join, axis=1).tolist()

        elif uploaded_file.name.lower().endswith(".pdf"):
            st.warning("PDF auto-fill needs pdfplumber logic added.")
            continue

        else:
            continue

        df = st.session_state.tables[target_table_index]["df"].copy()
        full_text = "\n".join(extracted_rows)

        filled = 0

        for idx, row in df.iterrows():
            spec = str(row.get("Specification", "")).strip()

            if not spec:
                continue

            for line in full_text.splitlines():
                if spec.lower().replace(" ", "") in line.lower().replace(" ", ""):
                    df.at[idx, target_col] = line
                    filled += 1
                    break

        st.session_state.tables[target_table_index]["df"] = df
        st.success(f"{uploaded_file.name}: {filled} rows filled under {target_col}")
