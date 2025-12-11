import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
import chardet

st.set_page_config(page_title="Course Completion — CSV & XLSX", layout="wide")
st.title("📘 Course Completion — CSV (pivot) & XLSX (multi-sheet)")

uploaded_file = st.file_uploader("Upload CSV (pivot) OR Excel (.xlsx multi-sheet)", type=["csv", "xlsx"])
if not uploaded_file:
    st.info("Upload a single-sheet pivot CSV (1 = pending) OR an Excel (.xlsx) with multiple sheets.")
    st.stop()

fname = uploaded_file.name.lower()

# ---------- Helpers ----------
def read_csv_smart(uploaded):
    uploaded.seek(0)
    raw = uploaded.read()
    enc = chardet.detect(raw[:20000])["encoding"] or "utf-8"
    text = raw.decode(enc, errors="replace")
    try:
        df = pd.read_csv(StringIO(text), sep=None, engine="python", skip_blank_lines=True)
    except Exception:
        df = pd.read_csv(StringIO(text))
    return df

def to_excel_bytes_from_df(df_out):
    out = BytesIO()
    df_out.to_excel(out, index=False, sheet_name="Pending")
    out.seek(0)
    return out.getvalue()

# ---------- CSV flow (pivot-style where 1 = pending) ----------
if fname.endswith(".csv"):
    try:
        df = read_csv_smart(uploaded_file)
    except Exception as e:
        st.error("Could not parse CSV. Try saving as UTF-8 CSV. Error: " + str(e))
        st.stop()

    # basic cleaning
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(axis=1, how="all")

    # detect employee name column (fallback to first)
    possible_name_cols = ["Employee Name", "Name of the Official", "Name", "Employee"]
    name_col = next((c for c in df.columns if c in possible_name_cols), None) or df.columns[0]

    # detect division col (optional)
    division_col = next((c for c in df.columns if "division" in c.lower() or "unit" in c.lower() or "region" in c.lower()), None)

    # exclude typical non-course cols
    exclude = {name_col}
    for c in df.columns:
        low = c.lower()
        if "s.no" in low or "s.n" in low or "employee no" in low or "emp no" in low or "sr.no" in low:
            exclude.add(c)
    if division_col:
        exclude.add(division_col)

    course_cols = [c for c in df.columns if c not in exclude]
    if not course_cols:
        st.error("No course columns detected in CSV. Ensure first column is Employee Name and the rest are course columns.")
        st.stop()

    # pending mask: cell == '1'
    def is_pending(x):
        if pd.isna(x):
            return False
        return str(x).strip() == "1"

    pending_mask = df[course_cols].applymap(is_pending)
    total_courses = len(course_cols)

    def compute_completion_for_indexes(indexes):
        sub = pending_mask.loc[indexes]
        n_emps = sub.shape[0]
        if n_emps == 0 or total_courses == 0:
            return 0.0, 0, 0
        pending_slots = int(sub.values.sum())
        total_slots = n_emps * total_courses
        completed_slots = total_slots - pending_slots
        pct = round((completed_slots / total_slots) * 100, 2)
        return pct, pending_slots, total_slots

    # Top metric
    if division_col and division_col in df.columns:
        rms_idx = df[df[division_col].astype(str).str.contains("RMS TP", case=False, na=False)].index
        rms_pct, rms_pending, rms_total = compute_completion_for_indexes(rms_idx)
        st.metric("RMS TP Division completion %", f"{rms_pct}%", delta=f"{len(rms_idx)} employees")
    else:
        overall_pct, overall_pending, overall_total = compute_completion_for_indexes(df.index)
        st.metric("Overall completion % (no division column)", f"{overall_pct}%", delta=f"{df.shape[0]} employees")

    st.markdown("---")

    # ---------- NEW SEARCH UI: live-match when typing >= 4 chars ----------
    st.subheader("Search employee (type at least 4 characters)")

    name_list = df[name_col].astype(str).str.strip().dropna().unique().tolist()
    name_list_sorted = sorted(name_list)

    # typed input
    search_input = st.text_input("Enter employee name (partial, case-insensitive) — start typing (min 4 chars)")

    chosen_name = None
    query = search_input.strip()

    if len(query) >= 4:
        matches = [n for n in name_list_sorted if query.lower() in n.lower()]

        if len(matches) == 0:
            st.info("No matches found for that string.")
        else:
            st.write(f"Matches found: {len(matches)}")

            # If small number of matches, show as radio buttons for quick click selection
            if len(matches) <= 30:
                chosen_name = st.radio("Select employee from matches", options=matches, index=0)
            else:
                # For large lists: show the first 100 in a table and provide a selectbox for exact pick
                st.write("Showing first 100 matches:")
                st.dataframe(pd.DataFrame({"Matches": matches[:100]}))
                chosen_name = st.selectbox("Pick one of the matches", options=matches)

    # Fallback: if typed <4 chars, let user pick from full list
    if (not chosen_name) and len(query) < 4:
        st.caption("Or pick a name from the list (useful if you don't want to type).")
        chosen_from_list = st.selectbox("Select employee (full list)", ["-- none --"] + name_list_sorted)
        if chosen_from_list != "-- none --":
            chosen_name = chosen_from_list

    if not chosen_name:
        st.info("Select or search an employee (type at least 4 characters to get live matches).")
        st.stop()

    # ---------- Employee result display (CSV branch) ----------
    emp_rows = df[df[name_col].astype(str).str.strip() == chosen_name]
    emp_pending_series = pending_mask.loc[emp_rows.index].any(axis=0)
    pending_courses = emp_pending_series[emp_pending_series].index.tolist()
    pending_count = len(pending_courses)
    completed_count = total_courses - pending_count
    completion_pct = round((completed_count / total_courses) * 100, 2)

    c1, c2 = st.columns([3,1])
    with c1:
        st.markdown(f"### {chosen_name}")
        st.metric("Completion %", f"{completion_pct}%", f"{completed_count}/{total_courses} completed")
        st.markdown("**Pending courses:**")
        if pending_count:
            st.dataframe(pd.DataFrame({"Pending Course": pending_courses}))
        else:
            st.success("No pending courses — all completed!")

    with c2:
        st.markdown("**Quick summary**")
        st.write(f"- Total courses: **{total_courses}**")
        st.write(f"- Pending: **{pending_count}**")
        st.write(f"- Completed: **{completed_count}**")
        if division_col:
            emp_divs = emp_rows[division_col].astype(str).dropna().unique().tolist()
            st.write(f"- Division(s): **{', '.join(emp_divs)}**")

    # download pending
    if pending_count:
        pending_df = pd.DataFrame({"Employee Name":[chosen_name]*pending_count, "Pending Course": pending_courses})
        st.download_button("📥 Download pending courses (Excel)", data=to_excel_bytes_from_df(pending_df),
                           file_name=f"{chosen_name}_pending_courses.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")

    # division summary table (if available)
    if division_col:
        st.subheader("Division completion summary")
        divs = df[division_col].astype(str).fillna("Unknown").unique().tolist()
        rows = []
        for d in sorted(divs):
            idxs = df[df[division_col].astype(str).fillna("Unknown") == d].index
            pct, pending_slots, total_slots = compute_completion_for_indexes(idxs)
            rows.append({"Division": d, "Completion %": pct, "Employees": len(idxs), "Pending slots": pending_slots, "Total slots": total_slots})
        div_df = pd.DataFrame(rows).sort_values("Completion %", ascending=False).reset_index(drop=True)
        st.dataframe(div_df)

# ---------- XLSX flow (multi-sheet consolidation) ----------
else:
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error("Unable to read Excel file: " + str(e))
        st.stop()

    combined_df = pd.DataFrame()
    for sheet in xls.sheet_names:
        # header=1 when first row is merged title
        df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet, header=1)
        df_sheet.columns = df_sheet.columns.astype(str).str.strip()
        df_sheet = df_sheet.dropna(axis=1, how="all")
        # drop columns named Unnamed or empty
        df_sheet = df_sheet[[c for c in df_sheet.columns if not str(c).lower().startswith("unnamed")]]

        # detect division col
        division_col = next((c for c in df_sheet.columns if "division" in c.lower() or "unit" in c.lower()), None)
        if division_col and division_col in df_sheet.columns:
            df_tp = df_sheet[df_sheet[division_col].astype(str).str.contains("RMS TP", case=False, na=False)]
        else:
            # fallback: search all cells for RMS TP
            tp_mask = df_sheet.apply(lambda col: col.astype(str).str.contains("RMS TP", case=False, na=False))
            if tp_mask.any().any():
                df_tp = df_sheet[tp_mask.any(axis=1)]
            else:
                df_tp = pd.DataFrame()

        if df_tp.empty:
            continue

        df_tp["Course Name"] = sheet
        # normalize employee columns if present
        possible_name_cols = ["Employee Name", "Name of the Official", "Name", "Employee"]
        possible_empno_cols = ["Employee No.", "Employee No", "Employee Number", "Emp No"]
        name_col = next((c for c in df_tp.columns if c in possible_name_cols), None)
        empno_col = next((c for c in df_tp.columns if c in possible_empno_cols), None)
        if name_col:
            df_tp = df_tp.rename(columns={name_col: "Employee Name"})
        if empno_col:
            df_tp = df_tp.rename(columns={empno_col: "Employee No."})

        combined_df = pd.concat([combined_df, df_tp], ignore_index=True)

    if combined_df.empty:
        st.error("No RMS TP data found in any sheet.")
        st.stop()

    st.success("Data extracted successfully from sheets.")
    st.dataframe(combined_df)

    # create pivot
    if "Employee Name" not in combined_df.columns:
        st.error("Employee column not found after consolidation.")
        st.stop()

    pivot_df = combined_df.pivot_table(index="Employee Name", columns="Course Name",
                                       values="Employee No.", aggfunc="count", fill_value=0)
    pivot_df["Grand Total"] = pivot_df.sum(axis=1)
    pivot_df.loc["Grand Total"] = pivot_df.sum(numeric_only=True)

    st.subheader("Pivot (Employee vs Course)")
    st.dataframe(pivot_df)

    # download pivot as simple excel
    st.download_button("📥 Download Pivot Excel", data=pivot_df.reset_index().to_excel(index=False, sheet_name="Pivot", engine="openpyxl"),
                       file_name="pivot_summary.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
