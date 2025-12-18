import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
import chardet

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Course Completion — CSV & XLSX", layout="wide")
st.title("📘 Course Completion — CSV (pivot) & XLSX (multi-sheet)")

uploaded_file = st.file_uploader(
    "Upload CSV (pivot) OR Excel (.xlsx multi-sheet)",
    type=["csv", "xlsx"]
)

if not uploaded_file:
    st.info("Upload a single-sheet pivot CSV (1 = pending) OR an Excel (.xlsx) with multiple sheets.")
    st.stop()

fname = uploaded_file.name.lower()

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def read_csv_smart(uploaded):
    uploaded.seek(0)
    raw = uploaded.read()
    enc = chardet.detect(raw[:20000])["encoding"] or "utf-8"
    text = raw.decode(enc, errors="replace")
    try:
        return pd.read_csv(StringIO(text), sep=None, engine="python")
    except Exception:
        return pd.read_csv(StringIO(text))


def df_to_excel_bytes(df, sheet_name="Sheet1"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


def normalize_columns(df):
    df.columns = df.columns.map(lambda x: str(x).strip() if pd.notna(x) else "Unnamed")
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        idxs = cols[cols == dup].index.tolist()
        for i, idx in enumerate(idxs):
            if i > 0:
                cols[idx] = f"{dup}.{i}"
    df.columns = cols
    return df

# --------------------------------------------------
# CSV FLOW (UNCHANGED)
# --------------------------------------------------
if fname.endswith(".csv"):
    df = read_csv_smart(uploaded_file)
    df = normalize_columns(df)
    df = df.dropna(axis=1, how="all")

    possible_name_cols = ["Employee Name", "Name of the Official", "Name", "Employee"]
    name_col = next((c for c in df.columns if c in possible_name_cols), df.columns[0])

    division_col = next(
        (c for c in df.columns if "division" in c.lower() or "unit" in c.lower()),
        None
    )

    exclude = {name_col}
    if division_col:
        exclude.add(division_col)

    for c in df.columns:
        if "s.no" in c.lower() or "emp" in c.lower():
            exclude.add(c)

    course_cols = [c for c in df.columns if c not in exclude]

    pending_mask = df[course_cols].applymap(
        lambda x: str(x).strip() == "1" if pd.notna(x) else False
    )

    total_courses = len(course_cols)

    st.metric("Overall Completion %",
              f"{round((1 - pending_mask.sum().sum() / (len(df) * total_courses)) * 100, 2)}%")

# --------------------------------------------------
# XLSX FLOW (MULTI-SHEET CONSOLIDATION)
# --------------------------------------------------
else:
    xls = pd.ExcelFile(uploaded_file)
    combined_df = pd.DataFrame()

    for sheet in xls.sheet_names:
        df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet)

        # Header fix
        df_sheet.columns = df_sheet.iloc[0]
        df_sheet = df_sheet[1:]
        df_sheet = df_sheet.dropna(axis=1, how="all")
        df_sheet = normalize_columns(df_sheet)

        # 🔹 NEW: Extract "Office of Working" from Column E
        office_col = df_sheet.columns[4]  # Column E (0-based index)
        df_sheet = df_sheet.rename(columns={office_col: "Office of Working"})

        division_col = next(
            (c for c in df_sheet.columns if "division" in c.lower() or "unit" in c.lower()),
            None
        )

        if not division_col:
            continue

        df_tp = df_sheet[df_sheet[division_col].astype(str).str.contains("RMS TP", case=False, na=False)]
        if df_tp.empty:
            continue

        df_tp = normalize_columns(df_tp)
        df_tp["Course Name"] = sheet

        name_col = next((c for c in df_tp.columns if "name" in c.lower()), None)
        if name_col:
            df_tp = df_tp.rename(columns={name_col: "Employee Name"})

        combined_df = pd.concat([combined_df, df_tp], ignore_index=True)

    if combined_df.empty:
        st.error("No RMS TP data found in Excel.")
        st.stop()

    combined_df = normalize_columns(combined_df)

    st.success("RMS TP data extracted successfully")
    st.dataframe(combined_df)

    # --------------------------------------------------
    # 🔹 UPDATED PIVOT WITH OFFICE OF WORKING
    # --------------------------------------------------
    pivot_df = combined_df.pivot_table(
        index=["Employee Name", "Office of Working"],
        columns="Course Name",
        aggfunc="size",
        fill_value=0
    ).reset_index()

    pivot_df["Total Courses"] = pivot_df.iloc[:, 2:].sum(axis=1)

    st.subheader("📊 Pivot: Employee vs Course (with Office)")
    st.dataframe(pivot_df)

    st.download_button(
        "📥 Download Pivot Excel",
        data=df_to_excel_bytes(pivot_df, "Pivot"),
        file_name="pivot_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
