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
        # --- Load Asset & Chart of Accounts ---
        xls = pd.ExcelFile(uploaded_file)
        df_asset = pd.read_excel(xls, sheet_name="Asset Review", skiprows=5)
        df_chart = pd.read_excel(xls, sheet_name="Chart of Accounts")

        # Normalize GL Code & Variances
        df_asset["GL Code"] = (
            pd.to_numeric(df_asset["Accounts"].astype(str).str.extract(r"(\d{4})")[0], errors="coerce")
            .dropna().astype(int).astype(str).str.zfill(4)
        )
        for col in ["$ Variance", "% Variance"]:
            df_asset[col] = pd.to_numeric(df_asset[col], errors="coerce")

        # Filter totals/blanks & highlight
        df_asset = df_asset[~df_asset["Accounts"].str.contains("(?i)total", na=False)]
        df_asset = df_asset[df_asset["Accounts"].str.strip() != ""]
        df_asset["Highlight"] = (
            (df_asset["$ Variance"].abs() >= 2000)
            | (df_asset["% Variance"].abs() >= 10)
        )
        df_asset = df_asset[df_asset["Highlight"]]

        # --- Load Trends sheets for story logic ---
        leasing_apps = prev_leasing_apps = None
        moveouts_cur = prev_moveouts = None
        if trends_file:
            t_xls = pd.ExcelFile(trends_file)
            # Applications trend
            if "Leasing" in t_xls.sheet_names:
                leasing_df = pd.read_excel(t_xls, sheet_name="Leasing")
                if "Applications" in leasing_df.columns:
                    vals = leasing_df["Applications"].dropna().astype(int)
                    if len(vals) >= 1: leasing_apps = vals.iloc[-1]
                    if len(vals) >= 2: prev_leasing_apps = vals.iloc[-2]
            # Move-Outs trend
            if "Move-ins/Move-outs" in t_xls.sheet_names:
                mo_df = pd.read_excel(t_xls, sheet_name="Move-ins/Move-outs")
                if "Move-Outs" in mo_df.columns:
                    mo_vals = mo_df["Move-Outs"].dropna().astype(int)
                    if len(mo_vals) >= 1: moveouts_cur = mo_vals.iloc[-1]
                    if len(mo_vals) >= 2: prev_moveouts = mo_vals.iloc[-2]
            # Unit mix for per-unit
            if "Unit Mix" in t_xls.sheet_names:
                um = pd.read_excel(t_xls, sheet_name="Unit Mix")
                total_units = pd.to_numeric(um[um.columns[1]].dropna().iloc[-1], errors="coerce")
            else:
                total_units = np.nan
        else:
            total_units = np.nan

        # --- Load GL journal entries ---
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
            codes = df_asset["GL Code"].unique()
            gl_df_raw = gl_df_raw[gl_df_raw["GL Code"].isin(codes)]
        else:
            gl_df_raw = pd.DataFrame(columns=["GL Code","Memo / Description","Debit","Credit"])

        # Merge description back for logic only
        df_merged = df_asset.merge(
            df_chart.rename(columns={
                "ACCOUNT NUMBER":"GL Code",
                "ACCOUNT DESCRIPTION":"Description"
            })[["GL Code","Description"]],
            how="left", on="GL Code"
        )

        def generate_explanation(row):
            gl = row["GL Code"]
            actual = row.get("Actuals", 0)
            budget = row.get("Budget Reporting", 0)
            var = row["$ Variance"]
            pct = row["% Variance"]
            direction = "under" if var < 0 else "over"
            expl = (
                f"GL {gl}: actuals ${actual:,.0f} {direction} budget "
                f"(${budget:,.0f}) by ${abs(var):,.0f} ({abs(pct):.1f}%). "
            )

            # YTD if present
            if "YTD Actuals" in row and "YTD Budget" in row and pd.notna(row["YTD Actuals"]):
                ytdv = row["YTD Actuals"] - row["YTD Budget"]
                trend = "continuing" if abs(ytdv) > abs(var) else "one-off"
                expl += f"YTD ${ytdv:,.0f} shows {trend} trend. "

            desc = str(row.get("Description","")).lower()

            # Application Fees story
            if "application" in desc and leasing_apps is not None:
                diff = leasing_apps - (prev_leasing_apps or leasing_apps)
                expl += (
                    f"We processed {leasing_apps} applications this period "
                    f"({diff:+} from prior), impacting application fee revenue. "
                )

            # Make-Ready costs story
            if "make ready" in desc and moveouts_cur is not None:
                diff_mo = moveouts_cur - (prev_moveouts or moveouts_cur)
                expl += (
                    f"Move-outs were {moveouts_cur} ({diff_mo:+}), "
                    f"driving increased make-ready spend. "
                )

            # Reversals & invoice‐size
            entries = gl_df_raw[gl_df_raw["GL Code"] == gl]
            if not entries.empty:
                revs = entries["Memo / Description"].str.contains(
                    "reverse|reversal", case=False, na=False
                ).sum()
                if revs:
                    expl += f"{revs} reversal entr{'y' if revs==1 else 'ies'}. "
                totals = entries[["Debit","Credit"]].fillna(0).sum(axis=1)
                if not totals.empty:
                    avg = totals.mean()
                    mx = totals.max()
                    if mx > avg * 1.5:
                        expl += (
                            f"One invoice of ${mx:,.0f} exceeded avg ${avg:,.0f}. "
                        )
                    elif len(totals) > 3:
                        expl += f"{len(totals)} postings contributed. "

            # Per-unit
            if pd.notna(total_units) and total_units > 0:
                expl += f"≈${actual/total_units:,.2f} per unit. "

            # Context
            if delay_note:
                expl += f"Billing delays: {delay_note}. "
            if major_event:
                expl += f"Event: {major_event}. "
            if moveout_note:
                expl += f"Turnover spike: {moveout_note}. "
            if staffing_note and var < 0:
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
