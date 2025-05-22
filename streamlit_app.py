import streamlit as st
import pandas as pd
import numpy as np

# Streamlit app setup
st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

# Uploads
uploaded_file = st.file_uploader("Upload Asset Review file", type=["xlsx"])
trends_file = st.file_uploader("Upload Trends file (Unit Mix)", type=["xlsx"])
gl_file = st.file_uploader("Upload General Ledger file", type=["xlsx"])

# Context for sentiment
with st.expander("📣 Add Monthly Context"):
    major_event = st.text_input("Major event this month?")
    delay_note = st.text_input("Any billing delays or credits?")
    moveout_note = st.text_input("Spike in move-outs?")
    staffing_note = st.text_input("If payroll is under budget, is the site understaffed?")

# Globals
gl_df_raw = pd.DataFrame()
total_units = np.nan

def generate_explanation(row):
    global gl_df_raw, total_units, major_event, delay_note, moveout_note, staffing_note

    if not row.get("Explain") or pd.isna(row.get("GL Code")):
        return ""

    gl = row["GL Code"]
    desc = row.get("Description", "this account")
    actual = row.get("Actuals", np.nan)
    budget = row.get("Budget Reporting", np.nan)
    ytd_actual = row.get("YTD Actuals", np.nan)
    ytd_budget = row.get("YTD Budget", np.nan)
    var = row.get("$ Variance", 0)
    pct_var = row.get("% Variance", 0)

    explanation = f"GL {gl} – {desc}:\n"
    explanation += f"- Actuals: ${actual:,.0f} vs. Budget: ${budget:,.0f} → {'Over' if var > 0 else 'Under'} by ${abs(var):,.0f} ({abs(pct_var):.1f}%)\n"

    if pd.notna(ytd_actual) and pd.notna(ytd_budget):
        ytd_diff = ytd_actual - ytd_budget
        if abs(ytd_diff) > abs(var):
            explanation += f"- YTD variance of ${ytd_diff:,.0f} suggests a continuing trend.\n"
        else:
            explanation += "- This appears to be an isolated variance.\n"

    if not gl_df_raw.empty and gl in gl_df_raw["GL Code"].values:
        entries = gl_df_raw[gl_df_raw["GL Code"] == gl]
        entry_count = len(entries)
        entry_amounts = entries[["Debit", "Credit"]].fillna(0).sum(axis=1)
        total_posted = entry_amounts.sum()
        avg_posted = entry_amounts.mean()
        max_posted = entry_amounts.max()

        explanation += f"- {entry_count} journal entries totaling ${total_posted:,.0f}\n"
        if max_posted >= 2 * avg_posted:
            explanation += f"- One entry of ${max_posted:,.0f} is significantly larger than the average of ${avg_posted:,.0f}\n"

        memos = entries["Memo / Description"].dropna()
        if not memos.empty:
            top_memos = memos.value_counts().head(2).index.tolist()
            explanation += f"- Top memo descriptions: {', '.join(top_memos)}\n"

    if not pd.isna(total_units) and total_units > 0 and pd.notna(actual):
        explanation += f"- Per-unit cost: ${actual / total_units:,.2f}\n"

    if gl in ["5205", "5210"] and staffing_note:
        explanation += f"- Staffing note: {staffing_note}\n"
    if gl in ["5601", "5671"] and moveout_note:
        explanation += f"- Move-out context: {moveout_note}\n"
    if delay_note:
        explanation += f"- Billing delay note: {delay_note}\n"
    if major_event:
        explanation += f"- Event context: {major_event}\n"

    return explanation.strip()

# Main process
if uploaded_file:
    try:
        # Load files
        xls = pd.ExcelFile(uploaded_file)
        df_asset = pd.read_excel(xls, sheet_name="Asset Review", skiprows=5)
        df_chart = pd.read_excel(xls, sheet_name="Chart of Accounts")

        # Convert GL Codes
        df_asset["GL Code Raw"] = df_asset["Accounts"].astype(str).str.extract(r'(\d{4})')[0]
        df_asset["GL Code Num"] = pd.to_numeric(df_asset["GL Code Raw"], errors="coerce")
        df_asset["GL Code"] = df_asset["GL Code Num"].apply(lambda x: str(int(x)).zfill(4) if pd.notna(x) else np.nan)

        # Convert variances
        df_asset["$ Variance"] = pd.to_numeric(df_asset["$ Variance"], errors="coerce")
        df_asset["% Variance"] = pd.to_numeric(df_asset["% Variance"], errors="coerce")

        # Add YTD columns if missing
        for col in ["YTD Actuals", "YTD Budget"]:
            if col not in df_asset.columns:
                df_asset[col] = np.nan

        # Excel-style filter logic
        df_asset["Highlight"] = (
            (df_asset["$ Variance"].abs() >= 2000) &
            (df_asset["% Variance"].abs() >= 10) &
            (~df_asset["Accounts"].astype(str).str.startswith((
                "Total", "Net", "Income", "4011", "4012", "6", "7", "8999"
            )))
        )
        df_asset["Explain"] = df_asset["Highlight"] & df_asset["GL Code"].notna()

        # Debug
        st.subheader("🔍 Explanation Candidates")
        st.write(f"Rows flagged for explanation: {df_asset['Explain'].sum()}")
        st.dataframe(df_asset[df_asset["Explain"]][["Accounts", "$ Variance", "% Variance", "GL Code"]])

        # Trends (unit mix)
        if trends_file:
            t_xls = pd.ExcelFile(trends_file)
            unitmix_df = pd.read_excel(t_xls, sheet_name="Unit Mix")
            total_units = pd.to_numeric(unitmix_df[unitmix_df.columns[1]].dropna().iloc[-1], errors='coerce')

        # Load GL if present
        if gl_file:
            gl_df_raw = pd.read_excel(gl_file, skiprows=8, header=None)
            gl_df_raw.columns = [
                "GL Code", "GL Name", "Post Date", "Effective Date", "Unused1", "Account Name",
                "Memo / Description", "Unused2", "Journal", "Unused3", "Debit", "Credit"
            ] + list(gl_df_raw.columns[12:])
            gl_df_raw["GL Code"] = gl_df_raw["GL Code"].astype(str).str.extract(r'(\d{4})')[0].str.zfill(4)
            gl_df_raw = gl_df_raw[gl_df_raw["GL Code"].isin(df_asset["GL Code"].dropna().unique())]
        else:
            gl_df_raw = pd.DataFrame(columns=["GL Code", "Memo / Description"])

        # Merge and explain
        df_chart = df_chart.rename(columns={
            'ACCOUNT NUMBER': 'GL Code',
            'ACCOUNT TITLE': 'Title',
            'ACCOUNT DESCRIPTION': 'Description'
        })
        df_chart["GL Code"] = pd.to_numeric(df_chart["GL Code"], errors="coerce").dropna().astype(int).astype(str).str.zfill(4)

        df_merged = df_asset.merge(df_chart[["GL Code", "Description"]], how="left", on="GL Code")
        df_merged["Explanation"] = df_merged.apply(generate_explanation, axis=1)

        # Display
        output_df = df_merged[df_merged["Explanation"].str.strip() != ""][[
            "GL Code", "Accounts", "Actuals", "Budget Reporting", "$ Variance",
            "% Variance", "YTD Actuals", "YTD Budget", "Explanation"
        ]]

        st.subheader("📄 Generated Explanations")
        st.dataframe(output_df.style.set_properties(**{'white-space': 'pre-wrap'}), use_container_width=True)

        csv = output_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV", data=csv, file_name="variance_explanations.csv")

    except Exception as e:
        st.error(f"❌ Error: {e}")
