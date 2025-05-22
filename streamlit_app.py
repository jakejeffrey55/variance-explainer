import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

uploaded_file = st.file_uploader("Upload your Excel file (Asset Review, Chart of Accounts)", type=["xlsx"])
trends_file = st.file_uploader("Upload your Trends file (Occupancy, Leasing, Move-ins/Move-outs, Unit Mix)", type=["xlsx"])
gl_file = st.file_uploader("Upload your General Ledger Report (.xlsx)", type=["xlsx"])

# Sentiment prompts
with st.expander("📣 Add Context for this Month"):
    major_event = st.text_input("Was there a major event this month? (e.g. vendor change, storm, freeze, emergency repair)")
    delay_note = st.text_input("Were there any billing delays, credits, or missing invoices?")
    moveout_note = st.text_input("Any known spike in move-outs or turnover pressure?")
    staffing_note = st.text_input("If payroll is under budget, is the site currently understaffed?")

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_asset = pd.read_excel(xls, sheet_name="Asset Review", skiprows=5)
        df_chart = pd.read_excel(xls, sheet_name="Chart of Accounts")

        st.write("✅ Asset Columns:", df_asset.columns.tolist())
        st.write("✅ Chart Columns:", df_chart.columns.tolist())

        df_asset["GL Code Raw"] = df_asset["Accounts"].astype(str).str.extract(r'(\d{4})')[0]
        df_asset["GL Code Num"] = pd.to_numeric(df_asset["GL Code Raw"], errors="coerce")
        df_asset["GL Type"] = df_asset["GL Code Num"].apply(lambda x: "Income" if 4000 <= x < 5000 else ("Expense" if 5000 <= x < 9000 else "Other") if pd.notna(x) else "Other")
        df_asset["GL Code"] = df_asset["GL Code Num"].apply(lambda x: str(int(x)).zfill(4) if pd.notna(x) else np.nan)

        df_chart = df_chart.rename(columns={
            'ACCOUNT NUMBER': 'GL Code',
            'ACCOUNT TITLE': 'Title',
            'ACCOUNT DESCRIPTION': 'Description'
        })
        df_chart["GL Code"] = pd.to_numeric(df_chart["GL Code"], errors="coerce").dropna().astype(int).astype(str).str.zfill(4)

        df_asset["GL Code"] = df_asset["GL Code"].astype(str).str.zfill(4)
        df_chart["GL Code"] = df_chart["GL Code"].astype(str).str.zfill(4)

        df_asset["$ Variance"] = pd.to_numeric(df_asset["$ Variance"], errors="coerce")
        df_asset["% Variance"] = pd.to_numeric(df_asset["% Variance"], errors="coerce")

        # Filter out totals and blanks
        df_asset = df_asset[~df_asset["Accounts"].astype(str).str.contains("(?i)total", na=False)]
        df_asset = df_asset[df_asset["Accounts"].astype(str).str.strip() != ""]

        # Highlight filter logic (simulate yellow highlight via formula logic)
        df_asset["Highlight"] = (
            (df_asset["$ Variance"].abs() >= 2000) & (df_asset["% Variance"].abs() >= 10)
        )

        df_asset["Explain"] = df_asset["Highlight"] & df_asset["GL Code"].notna()

        relevant_gl_codes = df_asset[df_asset["Explain"]]["GL Code"].dropna().unique()
        df_asset = df_asset[df_asset["GL Code"].isin(relevant_gl_codes)]

        if trends_file:
            t_xls = pd.ExcelFile(trends_file)
            unitmix_df = pd.read_excel(t_xls, sheet_name="Unit Mix")
            total_units = pd.to_numeric(unitmix_df[unitmix_df.columns[1]].dropna().iloc[-1], errors='coerce')
        else:
            total_units = np.nan

        if gl_file:
            gl_df_raw = pd.read_excel(gl_file, skiprows=8, header=None)
            gl_df_raw.columns = [
                "GL Code", "GL Name", "Post Date", "Effective Date", "Unused1", "Account Name",
                "Memo / Description", "Unused2", "Journal", "Unused3", "Debit", "Credit"
            ] + list(gl_df_raw.columns[12:])
            gl_df_raw["GL Code"] = gl_df_raw["GL Code"].astype(str).str.extract(r'(\d{4})')[0].str.zfill(4)
            gl_df_raw = gl_df_raw[gl_df_raw["GL Code"].isin(relevant_gl_codes)]
            st.write("✅ GL File Loaded Columns:", gl_df_raw.columns.tolist())
            st.write("✅ First few GL Codes:", gl_df_raw['GL Code'].dropna().unique()[:5])
        else:
            gl_df_raw = pd.DataFrame(columns=["GL Code", "Memo / Description"])

        df_merged = df_asset.merge(df_chart[["GL Code", "Description"]], how="left", on="GL Code")

        def generate_explanation(row):
    if not row["Explain"] or pd.isna(row["GL Code"]):
        return ""

    gl = row["GL Code"]
    desc = row.get("Description", "this account")
    actual = row.get("Actuals", np.nan)
    budget = row.get("Budget Reporting", np.nan)
    ytd_actual = row.get("YTD Actuals", np.nan)
    ytd_budget = row.get("YTD Budget", np.nan)
    var = row.get("$ Variance", 0)
    pct_var = row.get("% Variance", 0)
    direction = "unfavorable" if var < 0 else "favorable"

    explanation = f"GL {gl} – {desc}: This month's actuals of ${actual:,.0f} exceeded the ${budget:,.0f} budget by ${abs(var):,.0f} ({abs(pct_var):.1f}%). "

    # YTD variance context
    if pd.notna(ytd_actual) and pd.notna(ytd_budget):
        ytd_variance = ytd_actual - ytd_budget
        if abs(ytd_variance) > abs(var):
            explanation += f"The YTD variance of ${ytd_variance:,.0f} suggests a continuing trend. "
        else:
            explanation += "This appears to be a one-time deviation. "

    # GL context
    if not gl_df_raw.empty and gl in gl_df_raw["GL Code"].values:
        entries = gl_df_raw[gl_df_raw["GL Code"] == gl]
        entry_count = len(entries)
        entry_total = entries[["Debit", "Credit"]].fillna(0).sum(axis=1).sum()
        avg_entry = entry_total / entry_count if entry_count else 0
        max_entry = entries[["Debit", "Credit"]].fillna(0).sum(axis=1).max()

        if max_entry >= 2 * avg_entry:
            explanation += f"There is an unusually large journal entry of ${max_entry:,.0f} compared to the average of ${avg_entry:,.0f}. "
        elif entry_count > 5:
            explanation += f"A higher number of entries ({entry_count}) this month may have contributed. "

        memos = entries["Memo / Description"].dropna()
        if not memos.empty:
            top_memos = memos.value_counts().head(2).index.tolist()
            explanation += f" Top memo descriptions include: {', '.join(top_memos)}. "

    # Per-unit cost if available
    if not pd.isna(total_units) and total_units > 0 and pd.notna(actual):
        per_unit = actual / total_units
        explanation += f" Per-unit cost is approximately ${per_unit:,.2f}. "

    # Sentiment context
    if gl in ["5205", "5210"] and staffing_note:
        explanation += f" Staffing note: {staffing_note}."
    if moveout_note and gl in ["5601", "5671"]:
        explanation += f" Move-out context: {moveout_note}."
    if delay_note:
        explanation += f" Billing delay note: {delay_note}."
    if major_event:
        explanation += f" Event context: {major_event}."

    return explanation.strip()

        df_merged["Explanation"] = df_merged.apply(generate_explanation, axis=1)

        output_df = df_merged[df_merged["GL Code"].notna()][[col for col in [
            "GL Code", "Accounts", "Actuals", "Budget Reporting", "$ Variance",
            "% Variance", "YTD Actuals", "YTD Budget", "Explanation"
        ] if col in df_merged.columns]]

        st.success("Explanation generation complete ✅")
        st.dataframe(output_df, use_container_width=True)

        csv = output_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Results as CSV", data=csv, file_name="variance_explanations.csv")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
