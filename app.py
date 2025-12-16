import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Project Total Extractor",
    layout="wide"
)

st.title("Project Total Extractor")
st.write(
    "Upload an Excel expense report. "
    "The app extracts **Project Names** from rows that start with "
    "**'Total for'** in Column A and pulls the **total amount from Column G**."
)

# -----------------------------
# Deterministic extraction logic
# -----------------------------
TOTAL_FOR_RE = re.compile(r"^\s*Total for\s*(.*)\s*$", re.IGNORECASE)

def extract_project_totals(excel_bytes: bytes) -> pd.DataFrame:
    """
    Extract project totals deterministically:
    - Column A: 'Total for <Project Name>'
    - Column G: total amount
    """
    df = pd.read_excel(
        io.BytesIO(excel_bytes),
        header=None,
        engine="openpyxl"
    )

    results = []

    for _, row in df.iterrows():
        col_a = row.iloc[0] if len(row) > 0 else None

        if isinstance(col_a, str):
            match = TOTAL_FOR_RE.match(col_a)
            if match:
                project_name = match.group(1).strip()
                total_amount = row.iloc[6] if len(row) > 6 else None

                results.append({
                    "Project Name": project_name,
                    "Total Amount": total_amount
                })

    out_df = pd.DataFrame(results)

    if not out_df.empty:
        out_df["Total Amount"] = pd.to_numeric(
            out_df["Total Amount"],
            errors="coerce"
        )

    return out_df


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Project Summary"
        )
    return buffer.getvalue()

# -----------------------------
# Streamlit UI
# -----------------------------
uploaded_file = st.file_uploader(
    "Upload Excel file (.xlsx)",
    type=["xlsx"]
)

if uploaded_file:
    excel_bytes = uploaded_file.read()
    summary_df = extract_project_totals(excel_bytes)

    st.subheader("Extracted Project Totals")

    if summary_df.empty:
        st.warning("No 'Total for …' rows were found in Column A.")
    else:
        st.dataframe(summary_df, use_container_width=True)

        excel_out = dataframe_to_excel_bytes(summary_df)

        st.download_button(
            label="Download Project Summary Excel",
            data=excel_out,
            file_name="Project_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
