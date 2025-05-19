import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

uploaded_file = st.file_uploader("Upload your Excel file (with Asset Review, Chart of Accounts, Invoices)", type=["xlsx"])
trends_file = st.file_uploader("Upload your Trends file (Occupancy, Leasing, Move-ins/outs, Unit Mix)", type=["xlsx"])

if uploaded_file:
    try:
        # Load base Excel workbook
        xls = pd.ExcelFile(uploaded_file)
        df_asset = pd.read_excel(xls, sheet_name="Asset Review", skiprows=5)
        df_chart = pd.read_excel(xls, sheet_name="Chart of Accounts")
        df_invoices = pd.read_excel(xls, sheet_name="Invoices ")

        # Extract & clean base sheets
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

        # Invoice stats
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

        # Trends import
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

        def should_explain(row):
            acct = str(row["Accounts"])
            if any([x in acct for x in ["Total", "Net", "Income"]]): return False
            if acct.startswith(("4011", "4012", "6", "7", "8999")): return False
            if abs(row["$ Variance"] or 0) < 1000: return False
            if abs(row["% Variance"] or 0) < 0.1: return False
            return True

        df_asset["Explain"] = df_asset.apply(should_explain, axis=1)

        # Merge descriptions, invoices, and stats
        df_merged = df_asset.merge(df_chart[["GL Code", "Description"]], how="left", on="GL Code")
        df_merged = df_merged.merge(invoice_totals, how="left", left_on="GL Code", right_on="GLCode")
        df_merged = df_merged.merge(invoice_stats, how="left", left_on="GL Code", right_on="GLCode", suffixes=("", "_stat"))

        # Prompt for payroll context if under budget
        understaffed_gls = df_merged[(df_merged["GL Code"].isin(["5205", "5210"])) & (df_merged["$ Variance"] < -1000)]
        if not understaffed_gls.empty:
            st.warning("⚠️ Payroll is under budget. Is the property currently understaffed?")
            understaffed_flag = st.radio("", ["Yes", "No"], index=1)
        else:
            understaffed_flag = "No"

        def generate_explanation(row):
            if not row["Explain"]:
                return ""
            gl = row['GL Code']
            desc = row['Description'] or "this account"
            var = row["$ Variance"]
            direction = "Unfavorable" if var < 0 else "Favorable"

            # Income logic
            if row["GL Type"] == "Income":
                occ_note = ""
                if occ_df is not None:
                    occ_row = occ_df.iloc[-1]  # assume latest
                    occ_note = f" Occupancy was {occ_row['Actual Occupancy']}% vs {occ_row['Budgeted Occupancy']}% budgeted."
                return f"{direction} variance in {desc} (GL {gl}) may be tied to income shortfalls.{occ_note}"

            # Payroll flag
            if gl in ["5205", "5210"] and understaffed_flag == "Yes":
                return f"{direction} variance in {desc} (GL {gl}) likely due to reduced staffing. Consider reviewing open positions."

            # Invoice outlier logic
            if pd.notna(row["Max Invoice"]) and row["Max Invoice"] >= 2 * row["Avg Invoice"]:
                return f"{direction} variance in {desc} (GL {gl}) due to invoice #{row['SupplierInvoiceNumber']} for ${row['Max Invoice']:,.2f}, which exceeds 2× the average of ${row['Avg Invoice']:,.2f}."

            # Move-out trend
            if moveout_df is not None and gl in ["5601", "5671"]:
                mo_row = moveout_df.iloc[-1]
                return f"{direction} variance in {desc} (GL {gl}) may relate to {mo_row['Move outs']} move-outs this month, which increases cleaning and service costs."

            if pd.isna(row["Total Invoiced"]) or row["Total Invoiced"] == 0:
                return f"{direction} variance in {desc} (GL {gl}) likely due to missing or unrecorded invoices."

            unit_note = f" Total invoiced: ${row['Total Invoiced']:,.2f}."
            if not pd.isna(total_units):
                per_unit = row['Total Invoiced'] / total_units
                unit_note += f" That equates to approx. ${per_unit:,.2f} per unit."

            return f"{direction} variance in {desc} (GL {gl}).{unit_note}"

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
