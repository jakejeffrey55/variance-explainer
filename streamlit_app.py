import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

uploaded_file = st.file_uploader("Upload your Excel file (with Asset Review, Chart of Accounts, Invoices)", type=["xlsx"])
occupancy_file = st.file_uploader("Optional: Upload Occupancy Trends file", type=["xlsx"])

if uploaded_file:
    try:
        # Load required sheets
        xls = pd.ExcelFile(uploaded_file)
        df_asset = pd.read_excel(xls, sheet_name="Asset Review", skiprows=5)
        df_chart = pd.read_excel(xls, sheet_name="Chart of Accounts")
        df_invoices = pd.read_excel(xls, sheet_name="Invoices ")

        # Occupancy trends (optional)
        occupancy_trend = None
        if occupancy_file:
            occ_xls = pd.ExcelFile(occupancy_file)
            occupancy_trend = pd.read_excel(occ_xls, sheet_name=0)

        # Prep and classify GLs
        df_asset["GL Code"] = df_asset["Accounts"].astype(str).str.extract(r'(\d{4})')
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

        # Detect outliers by GL
        invoice_stats = (
            df_invoices.groupby("GLCode")["SUM OF LineItemTotal"]
            .agg(['mean', 'max', 'idxmax'])
            .reset_index()
        )
        invoice_stats = invoice_stats.rename(columns={"mean": "Avg Invoice", "max": "Max Invoice"})
        invoice_stats = invoice_stats.merge(
            df_invoices[["SupplierInvoiceNumber", "SUM OF LineItemTotal"]],
            left_on="idxmax", right_index=True, how="left"
        )

        invoice_totals = (
            df_invoices.groupby("GLCode")["SUM OF LineItemTotal"]
            .sum().reset_index()
            .rename(columns={"SUM OF LineItemTotal": "Total Invoiced"})
        )

        def should_explain(row):
            acct = str(row["Accounts"])
            if any([x in acct for x in ["Total", "Net", "Income"]]): return False
            if acct.startswith(("4011", "4012", "6", "7", "8999")): return False
            if abs(row["$ Variance"] or 0) < 1000: return False
            if abs(row["% Variance"] or 0) < 0.1: return False
            return True

        df_asset["Explain"] = df_asset.apply(should_explain, axis=1)

        # Merge GL descriptions and invoice summaries
        df_merged = df_asset.merge(df_chart[["GL Code", "Description"]], how="left", on="GL Code")
        df_merged = df_merged.merge(invoice_totals, how="left", left_on="GL Code", right_on="GLCode")
        df_merged = df_merged.merge(invoice_stats, how="left", left_on="GL Code", right_on="GLCode", suffixes=("", "_stat"))

        def generate_explanation(row):
            if not row["Explain"]:
                return ""
            gl = row['GL Code']
            desc = row['Description'] or "this account"
            var = row["$ Variance"]
            direction = "Unfavorable" if var < 0 else "Favorable"

            if row["GL Type"] == "Income":
                occ_note = ""
                if occupancy_trend is not None:
                    occ_note = " and potential under-occupancy trends"
                return f"{direction} variance in {desc} (GL {gl}) likely due to lower-than-expected income{occ_note}."

            if pd.notna(row["Max Invoice"]) and row["Max Invoice"] >= 2 * row["Avg Invoice"]:
                return f"{direction} variance in {desc} (GL {gl}) due to invoice #{row['SupplierInvoiceNumber']} for ${row['Max Invoice']:,.2f}, which exceeds 2× the typical average of ${row['Avg Invoice']:,.2f}."

            if pd.isna(row["Total Invoiced"]) or row["Total Invoiced"] == 0:
                return f"{direction} variance in {desc} (GL {gl}) likely due to missing or unrecorded invoices."

            return f"{direction} variance in {desc} (GL {gl}) explained by invoiced total of ${row['Total Invoiced']:,.2f}."

        df_merged["Explanation"] = df_merged.apply(generate_explanation, axis=1)

        output_df = df_merged[[
            "GL Code", "Accounts", "Actuals", "Budget Reporting", "$ Variance",
            "% Variance", "Explanation"
        ]]

        st.success("Explanation generation complete ✅")
        st.dataframe(output_df, use_container_width=True)

        csv = output_df.to_csv(index=False).encode('utf-8')
        st.download_button("⬇️ Download Results as CSV", data=csv, file_name="variance_explanations.csv")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
