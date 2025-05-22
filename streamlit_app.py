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

        st.write("✅ Asset Columns:", df_asset.columns.tolist())
        st.write("✅ Chart Columns:", df_chart.columns.tolist())

        # --- Extract and classify GL codes ---
        df_asset["GL Code Raw"] = df_asset["Accounts"].astype(str).str.extract(r"(\d{4})")[0]
        df_asset["GL Code Num"] = pd.to_numeric(df_asset["GL Code Raw"], errors="coerce")
        df_asset["GL Code"] = df_asset["GL Code Num"].apply(
            lambda x: str(int(x)).zfill(4) if pd.notna(x) else np.nan
        )
        df_asset["GL Code"] = df_asset["GL Code"].astype(str).str.zfill(4)

        # --- Normalize Chart of Accounts GL codes ---
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

        # --- Convert variance columns ---
        df_asset["$ Variance"] = pd.to_numeric(df_asset["$ Variance"], errors="coerce")
        df_asset["% Variance"] = pd.to_numeric(df_asset["% Variance"], errors="coerce")

        # --- Remove totals & blanks ---
        df_asset = df_asset[
            ~df_asset["Accounts"].astype(str).str.contains("(?i)total", na=False)
        ]
        df_asset = df_asset[df_asset["Accounts"].astype(str).str.strip() != ""]

        # --- Highlight logic: either large $ OR large % ---
        df_asset["Highlight"] = (
            (df_asset["$ Variance"].abs() >= 2000)
            | (df_asset["% Variance"].abs() >= 10)
        )
        # Keep only highlighted rows
        df_asset = df_asset[df_asset["Highlight"]]

        # --- Trends file for total units ---
        if trends_file:
            t_xls = pd.ExcelFile(trends_file)
            unitmix_df = pd.read_excel(t_xls, sheet_name="Unit Mix")
            total_units = pd.to_numeric(
                unitmix_df[unitmix_df.columns[1]].dropna().iloc[-1],
                errors="coerce"
            )
        else:
            total_units = np.nan

        # --- Load GL journal entries, filtered to codes in df_asset ---
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

        # --- Merge with Chart descriptions ---
        df_merged = df_asset.merge(
            df_chart[["GL Code", "Description"]],
            how="left",
            on="GL Code"
        )

        # --- Explanation generator with story and split-filtering ---
        def generate_explanation(row):
            label = (
                f"GL {row['GL Code']}"
                if pd.notna(row["GL Code"]) else row["Accounts"]
            )
            desc = row.get("Description") or row["Accounts"]
            actual = row.get("Actuals", np.nan)
            budget = row.get("Budget Reporting", np.nan)
            ytd_actual = row.get("YTD Actuals", np.nan)
            ytd_budget = row.get("YTD Budget", np.nan)
            var = row.get("$ Variance", 0)
            pct_var = row.get("% Variance", 0)
            direction = "unfavorable" if var < 0 else "favorable"

            # Narrative intro
            explanation = (
                f"This variance story for {label} – {desc} – begins with actuals of "
                f"${actual:,.0f}, coming in {direction} of the ${budget:,.0f} budget "
                f"by ${abs(var):,.0f} ({abs(pct_var):.1f}%). "
            )

            # YTD context
            if pd.notna(ytd_actual) and pd.notna(ytd_budget):
                ytd_var = ytd_actual - ytd_budget
                if abs(ytd_var) > abs(var):
                    explanation += "Year-to-date variance indicates this is a continuing trend. "
                else:
                    explanation += "However, this appears to be a one-off deviation. "

            # Journal entries context
            if pd.notna(row["GL Code"]) and not gl_df_raw.empty:
                entries = gl_df_raw[gl_df_raw["GL Code"] == row["GL Code"]]
                if not entries.empty:
                    entry_count = len(entries)
                    totals = entries[["Debit", "Credit"]].fillna(0).sum(axis=1)
                    entry_total = totals.sum()
                    avg_entry = entry_total / entry_count if entry_count else 0
                    max_entry = totals.max()

                    explanation += (
                        f"Collectively, {entry_count} journal postings totaled "
                        f"${entry_total:,.0f}. "
                    )
                    if max_entry >= 2 * avg_entry:
                        explanation += (
                            f"One entry peaked at ${max_entry:,.0f}, more than twice the "
                            f"average ${avg_entry:,.0f}. "
                        )
                    elif entry_count > 5:
                        explanation += f"A higher-than-normal entry volume contributed. "

                    # Memo/story filtering out Village phase splits
                    memos = entries["Memo / Description"].dropna()
                    base_memos = (
                        memos[~memos.str.contains("Village", case=False)]
                        .str.replace(r"\s*-\s*Phase\s*\d+", "", regex=True)
                        .str.strip()
                    )
                    counts = base_memos.value_counts()
                    if not counts.empty:
                        top = counts.head(2).index.tolist()
                        explanation += f"Key invoices: {', '.join(top)}. "
                        splits = counts[counts > 1]
                        for inv, cnt in splits.items():
                            explanation += f"Invoice '{inv}' was split across {cnt} phases. "

            # Per-unit cost
            if pd.notna(total_units) and total_units > 0 and pd.notna(actual):
                per_unit = actual / total_units
                explanation += f"That equates to about ${per_unit:,.2f} per unit. "

            # Sentiment/context
            if row["GL Code"] in ["5205", "5210"] and staffing_note:
                explanation += f"Staffing note: {staffing_note}. "
            if row["GL Code"] in ["5601", "5671"] and moveout_note:
                explanation += f"Move-out context: {moveout_note}. "
            if delay_note:
                explanation += f"Billing delay note: {delay_note}. "
            if major_event:
                explanation += f"Event note: {major_event}. "

            return explanation.strip()

        df_merged["Explanation"] = df_merged.apply(generate_explanation, axis=1)

        # --- Output ---
        cols = [
            "GL Code", "Accounts", "Actuals", "Budget Reporting", "$ Variance",
            "% Variance", "YTD Actuals", "YTD Budget", "Explanation"
        ]
        output_df = df_merged[[c for c in cols if c in df_merged.columns]]

        st.success("Explanation generation complete ✅")
        st.dataframe(output_df, use_container_width=True)

        csv_bytes = output_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Results as CSV",
            data=csv_bytes,
            file_name="variance_explanations.csv"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
