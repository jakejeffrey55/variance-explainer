import streamlit as st
import pandas as pd

st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load required sheets
        xls = pd.ExcelFile(uploaded_file)
        df_asset = pd.read_excel(xls, sheet_name="Asset Review", skiprows=5)
        df_chart = pd.read_excel(xls, sheet_name="Chart of Accounts")
        df_invoices = pd.read_excel(xls, sheet_name="Invoices ")

        # Prep
        df_asset["GL Code"] = df_asset["Accounts"].astype(str).str.extract(r'(\d{4})')
        df_asset["$ Variance"] = pd.to_numeric(df_asset["$ Variance"], errors="coerce")
        df_asset["% Variance"] = pd.to_numeric(df_asset["% Variance"], errors="coerce")

        df_chart = df_chart.rename(columns={
            'ACCOUNT NUMBER': 'GL Code',
            'ACCOUNT TITLE': 'Title',
            'ACCOUNT DESCRIPTION': 'Description'
        })

        df_invoices["GLCode"] = df_invoices["GLCode"].astype(str).str.zfill(4)
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

        df_explained = df_asset.merge(df_chart[["GL Code", "Description"]], how="left", on="GL Code")
        df_explained = df_explained.merge(invoice_totals, how="left", left_on="GL Code", right_on="GLCode")

        def generate_explanation(row):
            if not row["Explain"]:
                return ""
            direction = "Unfavorable" if row["$ Variance"] < 0 else "Favorable"
            desc = row["Description"] if pd.notna(row["Description"]) else "this account"
            if pd.isna(row["Total Invoiced"]) or row["Total Invoiced"] == 0:
                return f"{direction} variance in {desc} (GL {row['GL Code']}) likely due to missing or unrecorded invoices."
            return f"{direction} variance in {desc} (GL {row['GL Code']}) explained by total invoicing of ${row['Total Invoiced']:,.2f}."

        df_explained["Explanation"] = df_explained.apply(generate_explanation, axis=1)

        # Output only necessary columns for display
        output_df = df_explained[[
            "GL Code", "Accounts", "Actuals", "Budget Reporting", "$ Variance",
            "% Variance", "Explanation"
        ]]

        st.success("Explanation generation complete ✅")
        st.dataframe(output_df, use_container_width=True)

        # Offer download
        csv = output_df.to_csv(index=False).encode('utf-8')
        st.download_button("⬇️ Download Results as CSV", data=csv, file_name="variance_explanations.csv")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
