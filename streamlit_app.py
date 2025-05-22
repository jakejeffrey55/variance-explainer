import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Variance Explanation Generator", layout="wide")
st.title("📊 Cortland Asset Variance Explainer")

uploaded_file = st.file_uploader(
    "Upload your Excel file (Asset Review, Chart of Accounts)", type=["xlsx"]
)
trends_file = st.file_uploader(
    "Upload your Trends file (Occupancy, Leasing, Move-ins/Move-outs, Unit Mix)", type=["xlsx"]
)
gl_file = st.file_uploader(
    "Upload your General Ledger Report (.xlsx)", type=["xlsx"]
)

with st.expander("📣 Add Context for this Month"):
    major_event = st.text_input(
        "Was there a major event this month? (e.g. vendor change, storm, freeze, emergency repair)"
    )
    delay_note = st.text_input("Were there any billing delays, credits, or missing invoices?")
    moveout_note = st.text_input("Any known spike in move-outs or turnover pressure?")
    staffing_note = st.text_input(
        "If payroll is under budget, is the site currently understaffed?"
    )

if uploaded_file:
    try:
        # --- Load Asset Review & Chart of Accounts ---
        xls = pd.ExcelFile(uploaded_file)
        df_asset = pd.read_excel(xls, sheet_name="Asset Review", skiprows=5)
        df_chart = pd.read_excel(xls, sheet_name="Chart of Accounts")

        # Extract GL codes
        df_asset["GL Code Raw"] = df_asset["Accounts"].astype(str).str.extract(r"(\d{4})")[0]
        df_asset["GL Code"] = (
            pd.to_numeric(df_asset["GL Code Raw"], errors="coerce")
            .dropna().astype(int).astype(str).str.zfill(4)
        )
        df_asset["$ Variance"] = pd.to_numeric(df_asset["$ Variance"], errors="coerce")
        df_asset["% Variance"] = pd.to_numeric(df_asset["% Variance"], errors="coerce")

        # Clean up
        df_asset = df_asset[
            ~df_asset["Accounts"].astype(str).str.contains("(?i)total", na=False)
        ]
        df_asset = df_asset[df_asset["Accounts"].astype(str).str.strip() != ""]
        df_asset["Highlight"] = (
            (df_asset["$ Variance"].abs() >= 2000)
            | (df_asset["% Variance"].abs() >= 10)
        )
        df_asset = df_asset[df_asset["Highlight"]]

        # Trends for per-unit
        if trends_file:
            t_xls = pd.ExcelFile(trends_file)
            unitmix_df = pd.read_excel(t_xls, sheet_name="Unit Mix")
            total_units = pd.to_numeric(
                unitmix_df[unitmix_df.columns[1]].dropna().iloc[-1], errors="coerce"
            )
        else:
            total_units = np.nan

        # Load relevant GL entries
        if gl_file:
            gl_df_raw = pd.read_excel(gl_file, skiprows=8, header=None)
            cols = [
                "GL Code","GL Name","Post Date","Effective Date","Unused1",
                "Account Name","Memo / Description","Unused2","Journal",
                "Unused3","Debit","Credit"
            ]
            gl_df_raw.columns = cols + list(gl_df_raw.columns[12:])
            gl_df_raw["GL Code"] = (
                gl_df_raw["GL Code"].astype(str)
                .str.extract(r"(\d{4})")[0].str.zfill(4)
            )
            codes = df_asset["GL Code"].dropna().unique()
            gl_df_raw = gl_df_raw[gl_df_raw["GL Code"].isin(codes)]
        else:
            gl_df_raw = pd.DataFrame(columns=["GL Code","Memo / Description","Debit","Credit"])

        # Merge back just for indexing; we won't echo descriptions
        df_merged = df_asset[[
            "GL Code","Accounts","Actuals","Budget Reporting",
            "$ Variance","% Variance","YTD Actuals","YTD Budget"
        ]].copy()

        def generate_explanation(row):
            gl = row["GL Code"]
            actual = row.get("Actuals", np.nan)
            budget = row.get("Budget Reporting", np.nan)
            var = row.get("$ Variance", 0)
            pct = row.get("% Variance", 0)
            direction = "under" if var < 0 else "over"

            # Start with the core variance
            expl = (
                f"GL {gl}: actuals of ${actual:,.0f} came in {direction} budget "
                f"(${budget:,.0f}) by ${abs(var):,.0f} ({abs(pct):.1f}%). "
            )

            # YTD
            ytd_act = row.get("YTD Actuals", np.nan)
            ytd_bud = row.get("YTD Budget", np.nan)
            if pd.notna(ytd_act) and pd.notna(ytd_bud):
                ytdv = ytd_act - ytd_bud
                expl += (
                    f"YTD variance ${ytdv:,.0f} shows "
                    f"{'continuation' if abs(ytdv) > abs(var) else 'a one-off'} trend. "
                )

            # Invoice reversals
            if not gl_df_raw.empty:
                entries = gl_df_raw[gl_df_raw["GL Code"] == gl]
                revs = entries[
                    entries["Memo / Description"]
                    .str.contains("reverse|reversal", case=False, na=False)
                ]
                if not revs.empty:
                    count = len(revs)
                    expl += f"{count} reversal entr{'y' if count==1 else 'ies'} detected. "

                # Large or frequent invoices
                totals = entries[["Debit","Credit"]].fillna(0).sum(axis=1)
                if not totals.empty:
                    entry_count = len(totals)
                    sum_tot = totals.sum()
                    avg = sum_tot / entry_count
                    mx = totals.max()
                    if mx > avg * 1.5:
                        expl += (
                            f"An invoice of ${mx:,.0f} exceeded the average "
                            f"${avg:,.0f}, driving variance. "
                        )
                    elif entry_count > 3:
                        expl += f"{entry_count} postings this month contributed. "

            # Per-unit framing
            if pd.notna(total_units) and total_units > 0:
                per_u = actual / total_units
                expl += f"Equates to ${per_u:,.2f} per unit. "

            # Context notes
            if delay_note:
                expl += f"Billing delays: {delay_note}. "
            if major_event:
                expl += f"Event: {major_event}. "
            if moveout_note:
                expl += f"Turnover spike: {moveout_note}. "
            if staffing_note and pct < 0:
                expl += f"Staffing note: {staffing_note}. "

            return expl.strip()

        df_merged["Explanation"] = df_merged.apply(generate_explanation, axis=1)

        st.success("Explanation generation complete ✅")
        st.dataframe(df_merged, use_container_width=True)
        st.download_button(
            "⬇️ Download Results as CSV",
            data=df_merged.to_csv(index=False).encode("utf-8"),
            file_name="variance_explanations.csv"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
