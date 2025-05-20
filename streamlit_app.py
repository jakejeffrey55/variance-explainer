import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

uploaded_file = st.file_uploader("Upload your Excel file (Asset Review, Chart of Accounts, Invoices)", type=["xlsx"])
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
        df_invoices = pd.read_excel(xls, sheet_name="Invoices ")

        df_asset["GL Code"] = df_asset["Accounts"].astype(str).str.extract(r'(\d{4})')[0]
        df_asset["GL Code"] = df_asset["GL Code"].astype(str).str.zfill(4)
        df_chart["GL Code"] = df_chart["GL Code"].astype(str).str.zfill(4)
        df_asset["$ Variance"] = pd.to_numeric(df_asset["$ Variance"], errors="coerce")
        df_asset["% Variance"] = pd.to_numeric(df_asset["% Variance"], errors="coerce")
        df_asset["GL Type"] = df_asset["GL Code"].astype(float).apply(
            lambda x: "Income" if 4000 <= x < 5000 else ("Expense" if 5000 <= x < 9000 else "Other")
        )

        df_chart = df_chart.rename(columns={
            'ACCOUNT NUMBER': 'GL Code',
            'ACCOUNT TITLE': 'Title',
            'ACCOUNT DESCRIPTION': 'Description'
        })

        df_invoices["GLCode"] = df_invoices["GLCode"].astype(str).str.zfill(4)
        df_invoices["SupplierInvoiceNumber"] = df_invoices["SupplierInvoiceNumber"].astype(str)

        invoice_stats = (
            df_invoices.groupby("GLCode")["SUM OF LineItemTotal"]
            .agg(['mean', 'max', 'idxmax'])
            .reset_index()
            .rename(columns={"mean": "Avg Invoice", "max": "Max Invoice"})
        )
        invoice_stats = invoice_stats.merge(
            df_invoices[["SupplierInvoiceNumber", "SUM OF LineItemTotal"]],
            left_on="idxmax", right_index=True, how="left"
        )

        invoice_totals = (
            df_invoices.groupby("GLCode")["SUM OF LineItemTotal"]
            .sum().reset_index()
            .rename(columns={"SUM OF LineItemTotal": "Total Invoiced"})
        )

        if trends_file:
            t_xls = pd.ExcelFile(trends_file)
            occ_df = pd.read_excel(t_xls, sheet_name="Occupancy vs Budget")
            leasing_df = pd.read_excel(t_xls, sheet_name="Leasing Trends")
            movein_df = pd.read_excel(t_xls, sheet_name="Move ins")
            moveout_df = pd.read_excel(t_xls, sheet_name="Move outs")
            unitmix_df = pd.read_excel(t_xls, sheet_name="Unit Mix")
            total_units = pd.to_numeric(unitmix_df[unitmix_df.columns[1]].dropna().iloc[-1], errors='coerce')
        else:
            occ_df = leasing_df = movein_df = moveout_df = unitmix_df = None
            total_units = np.nan

        if gl_file:
            gl_df_raw = pd.read_excel(gl_file, skiprows=7)
            gl_df_raw.columns.values[0:2] = ['GL Code', 'GL Name']
            gl_df_raw['GL Code'] = gl_df_raw['GL Code'].astype(str).str.extract(r'(\d{4})')[0].str.zfill(4)
        else:
            gl_df_raw = pd.DataFrame()

        def should_explain(row):
            gl_code = row["GL Code"]
            if pd.isna(gl_code): return False
            if pd.isna(row["Actuals"]) and pd.isna(row["$ Variance"]): return False
            return True

        df_asset["Explain"] = df_asset.apply(should_explain, axis=1)

        df_merged = df_asset.merge(df_chart[["GL Code", "Description"]], how="left", on="GL Code")
        df_merged = df_merged.merge(invoice_totals, how="left", left_on="GL Code", right_on="GLCode")
        df_merged = df_merged.merge(invoice_stats, how="left", left_on="GL Code", right_on="GLCode", suffixes=("", "_stat"))

        def generate_explanation(row):
            if not row["Explain"]:
                return ""

            gl = row["GL Code"]
            desc = row["Description"] if pd.notna(row["Description"]) else "this account"
            actual = row["Actuals"]
            budget = row["Budget Reporting"]
            ytd_actual = row.get("YTD Actuals", np.nan)
            ytd_budget = row.get("YTD Budget", np.nan)
            ytd_variance = ytd_actual - ytd_budget if pd.notna(ytd_actual) and pd.notna(ytd_budget) else np.nan
            var = row["$ Variance"]
            direction = "Unfavorable" if var < 0 else "Favorable"

            explanation = f"{direction} variance in {desc} (GL {gl}). "

            if pd.notna(ytd_variance) and abs(ytd_variance) > abs(var):
                explanation += f"YTD variance is growing (${ytd_variance:,.0f}), indicating a sustained overage pattern. "
            elif pd.notna(ytd_variance) and abs(ytd_variance) < abs(var):
                explanation += "This appears to be a one-time spike rather than an ongoing trend. "

            if pd.notna(row["Max Invoice"]) and row["Max Invoice"] >= 2 * row["Avg Invoice"]:
                explanation += f"Invoice #{row['SupplierInvoiceNumber']} for ${row['Max Invoice']:,.2f} is over 2× the average. "
            elif pd.isna(row["Total Invoiced"]) or row["Total Invoiced"] == 0:
                explanation += "No invoicing activity recorded this month. "
            else:
                explanation += f"Total invoiced: ${row['Total Invoiced']:,.2f}. "
                if not pd.isna(total_units) and total_units > 0:
                    per_unit = row["Total Invoiced"] / total_units
                    explanation += f"That equals approx. ${per_unit:,.2f} per unit. "

            if not gl_df_raw.empty and gl in gl_df_raw["GL Code"].values:
                memos = gl_df_raw[gl_df_raw["GL Code"] == gl]["Memo / Description"].dropna()
                if not memos.empty:
                    top_memos = memos.value_counts().head(2).index.tolist()
                    if top_memos:
                        explanation += f" Top GL memos: {', '.join(top_memos)}. "

            if gl in ["5205", "5210"] and staffing_note:
                explanation += f"Staffing note: {staffing_note}. "
            if moveout_note and gl in ["5601", "5671"]:
                explanation += f"Move-out context: {moveout_note}. "
            if delay_note:
                explanation += f"Billing delay note: {delay_note}. "
            if major_event:
                explanation += f"Event context: {major_event}. "

            return explanation.strip()

        df_merged["Explanation"] = df_merged.apply(generate_explanation, axis=1)

        output_df = df_merged[[
            "GL Code", "Accounts", "Actuals", "Budget Reporting", "$ Variance",
            "% Variance", "YTD Actuals", "YTD Budget", "Explanation"
        ]]

        st.success("Explanation generation complete ✅")
        st.dataframe(output_df, use_container_width=True)

        csv = output_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Results as CSV", data=csv, file_name="variance_explanations.csv")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
