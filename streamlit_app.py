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

# Sentiment prompts
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
            .dropna()
            .astype(int)
            .astype(str)
            .str.zfill(4)
        )
        df_chart = df_chart.rename(columns={
            "ACCOUNT NUMBER": "GL Code",
            "ACCOUNT TITLE": "Title",
            "ACCOUNT DESCRIPTION": "Description",
        })
        df_chart["GL Code"] = (
            pd.to_numeric(df_chart["GL Code"], errors="coerce")
            .dropna()
            .astype(int)
            .astype(str)
            .str.zfill(4)
        )

        # Clean and filter Asset sheet
        df_asset["$ Variance"] = pd.to_numeric(df_asset["$ Variance"], errors="coerce")
        df_asset["% Variance"] = pd.to_numeric(df_asset["% Variance"], errors="coerce")
        df_asset = df_asset[
            ~df_asset["Accounts"].astype(str).str.contains("(?i)total", na=False)
        ]
        df_asset = df_asset[df_asset["Accounts"].astype(str).str.strip() != ""]
        df_asset["Highlight"] = (
            (df_asset["$ Variance"].abs() >= 2000)
            | (df_asset["% Variance"].abs() >= 10)
        )
        df_asset = df_asset[df_asset["Highlight"]]

        # Total units from Trends
        if trends_file:
            t_xls = pd.ExcelFile(trends_file)
            unitmix_df = pd.read_excel(t_xls, sheet_name="Unit Mix")
            total_units = pd.to_numeric(
                unitmix_df[unitmix_df.columns[1]].dropna().iloc[-1],
                errors="coerce"
            )
        else:
            total_units = np.nan

        # Load GL journal entries for only the highlighted codes
        if gl_file:
            gl_df_raw = pd.read_excel(gl_file, skiprows=8, header=None)
            gl_df_raw.columns = [
                "GL Code", "GL Name", "Post Date", "Effective Date", "Unused1",
                "Account Name", "Memo / Description", "Unused2", "Journal",
                "Unused3", "Debit", "Credit"
            ] + list(gl_df_raw.columns[12:])
            gl_df_raw["GL Code"] = (
                gl_df_raw["GL Code"].astype(str)
                .str.extract(r"(\d{4})")[0]
                .str.zfill(4)
            )
            codes = df_asset["GL Code"].dropna().unique()
            gl_df_raw = gl_df_raw[gl_df_raw["GL Code"].isin(codes)]
        else:
            gl_df_raw = pd.DataFrame(columns=["GL Code", "Memo / Description", "Debit", "Credit"])

        # Merge in descriptions
        df_merged = df_asset.merge(
            df_chart[["GL Code", "Description"]],
            how="left",
            on="GL Code"
        )

        # New explanation generator: we **use** the Description to guide the story
        def generate_explanation(row):
            gl = row["GL Code"]
            desc = row.get("Description") or row["Accounts"]
            desc_low = desc.lower()
            actual = row.get("Actuals", np.nan)
            budget = row.get("Budget Reporting", np.nan)
            ytd_actual = row.get("YTD Actuals", np.nan)
            ytd_budget = row.get("YTD Budget", np.nan)
            var = row.get("$ Variance", 0)
            pct_var = row.get("% Variance", 0)
            direction = "below" if var < 0 else "above"

            # Frame with what the account **is**
            explanation = (
                f"GL {gl} covers **{desc}**, an account that tracks {desc_low}. "
                f"This month’s actuals of ${actual:,.0f} came in {direction} "
                f"the ${budget:,.0f} budget by ${abs(var):,.0f} ({abs(pct_var):.1f}%). "
            )

            # YTD trend
            if pd.notna(ytd_actual) and pd.notna(ytd_budget):
                ytd_var = ytd_actual - ytd_budget
                trend = "continuing" if abs(ytd_var) > abs(var) else "one-off"
                explanation += (
                    f"The YTD variance of ${ytd_var:,.0f} suggests a {trend} trend. "
                )

            # Journal entry deep-dive (filter out phase splits)
            if gl and not gl_df_raw.empty:
                entries = gl_df_raw[gl_df_raw["GL Code"] == gl]
                if not entries.empty:
                    entry_count = len(entries)
                    totals = entries[["Debit", "Credit"]].fillna(0).sum(axis=1)
                    tot_sum = totals.sum()
                    avg_e = tot_sum / entry_count if entry_count else 0
                    max_e = totals.max()

                    explanation += (
                        f"A total of {entry_count} postings summed to ${tot_sum:,.0f}. "
                    )
                    if max_e >= 2 * avg_e:
                        explanation += (
                            f"One large posting of ${max_e:,.0f} (vs avg ${avg_e:,.0f}) drove much of it. "
                        )
                    elif entry_count > 5:
                        explanation += "Elevated entry volume also played a role. "

                    # Pull key invoices, remove Village splits
                    memos = entries["Memo / Description"].dropna()
                    core = (
                        memos[~memos.str.contains("Village", case=False)]
                        .str.replace(r"\s*-\s*Phase\s*\d+", "", regex=True)
                        .str.strip()
                    )
                    top = core.value_counts().head(2).index.tolist()
                    if top:
                        explanation += f"Key invoices: {', '.join(top)}. "

            # Per-unit perspective
            if pd.notna(total_units) and total_units > 0 and pd.notna(actual):
                per_u = actual / total_units
                explanation += f"That’s about ${per_u:,.2f} per unit. "

            # Contextual notes
            if desc_low.find("salary") != -1 and staffing_note:
                explanation += f"Staffing note: {staffing_note}. "
            if moveout_note and any(k in desc_low for k in ["turnover", "move-out"]):
                explanation += f"Move-out context: {moveout_note}. "
            if delay_note:
                explanation += f"Billing delay note: {delay_note}. "
            if major_event:
                explanation += f"Event context: {major_event}. "

            return explanation.strip()

        df_merged["Explanation"] = df_merged.apply(generate_explanation, axis=1)

        # Display results
        cols = [
            "GL Code", "Accounts", "Actuals", "Budget Reporting", "$ Variance",
            "% Variance", "YTD Actuals", "YTD Budget", "Explanation"
        ]
        output_df = df_merged[[c for c in cols if c in df_merged.columns]]

        st.success("Explanation generation complete ✅")
        st.dataframe(output_df, use_container_width=True)

        st.download_button(
            "⬇️ Download Results as CSV",
            data=output_df.to_csv(index=False).encode("utf-8"),
            file_name="variance_explanations.csv"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
