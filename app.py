import io
import re
import pandas as pd
import streamlit as st

# Optional OpenAI summary
from openai import OpenAI


st.set_page_config(page_title="Project Total Extractor", layout="wide")

st.title("Project Total Extractor (from 'Total for ...' rows)")
st.write(
    "Uploads an Excel report, finds rows where Column A contains **'Total for'**, "
    "extracts the project name, and pulls the total from **Column G**."
)

# -----------------------------
# Deterministic extraction logic
# -----------------------------
TOTAL_FOR_RE = re.compile(r"^\s*Total for\s*(.*)\s*$", re.IGNORECASE)

def extract_project_totals(excel_bytes: bytes) -> pd.DataFrame:
    """
    Deterministically extract project totals:
      - Find rows where Column A matches 'Total for ...'
      - Project Name = text after 'Total for'
      - Total Amount = Column G (index 6)
    """
    # Read with no header, because your file has report text and a header line inside
    df = pd.read_excel(io.BytesIO(excel_bytes), header=None, engine="openpyxl")

    rows = []
    for _, row in df.iterrows():
        col_a = row.iloc[0] if len(row) > 0 else None
        if isinstance(col_a, str):
            m = TOTAL_FOR_RE.match(col_a)
            if m:
                project_name = (m.group(1) or "").strip()
                total_amount = row.iloc[6] if len(row) > 6 else None  # Column G
                rows.append({"Project Name": project_name, "Total Amount": total_amount})

    out = pd.DataFrame(rows)

    # Normalize numeric column (keeps negatives)
    if not out.empty:
        out["Total Amount"] = pd.to_numeric(out["Total Amount"], errors="coerce")

    return out


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Project Summary")
    return output.getvalue()


# -----------------------------
# Optional GPT narrative summary
# -----------------------------
GPT_PROMPT = """You are a careful accounting assistant.

You will be given a JSON array of objects with:
- "Project Name" (string)
- "Total Amount" (number or null)

Your job:
1) Produce a short summary for humans.
2) Return STRICT JSON only that matches the provided schema.
3) Do not invent values; only use what is provided.
4) If totals are null, mention they are missing.
"""

GPT_SCHEMA = {
    "name": "project_total_summary",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "high_level": {"type": "string"},
            "largest_projects": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "project_name": {"type": "string"},
                        "total_amount": {"type": "number"},
                    },
                    "required": ["project_name", "total_amount"],
                },
            },
            "notes": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["high_level", "largest_projects", "notes"],
    },
}

def gpt_make_summary(project_df: pd.DataFrame) -> dict:
    """
    Optional: creates a consistent narrative summary using temperature=0 and a strict schema.
    Requires OPENAI_API_KEY configured in Streamlit secrets/environment.
    """
    client = OpenAI()  # reads OPENAI_API_KEY from env or Streamlit secrets

    payload = project_df.to_dict(orient="records")

    resp = client.responses.create(
        model="gpt-4.1-mini",   # pick what you have access to
        temperature=0,
        input=[
            {"role": "system", "content": GPT_PROMPT},
            {"role": "user", "content": f"Data:\n{payload}"},
        ],
        response_format={"type": "json_schema", "json_schema": GPT_SCHEMA},
    )
    return resp.output_parsed


# -----------------------------
# Streamlit UI
# -----------------------------
uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

use_gpt = st.toggle("Also generate a short GPT summary (optional)", value=False)
top_n = st.number_input("For GPT summary: show Top N largest projects", min_value=1, max_value=25, value=5, step=1)

if uploaded:
    excel_bytes = uploaded.read()
    summary_df = extract_project_totals(excel_bytes)

    st.subheader("Extracted Project Totals")
    if summary_df.empty:
        st.warning("No 'Total for ...' rows found in Column A.")
    else:
        st.dataframe(summary_df, use_container_width=True)

        # Download summary as Excel
        out_bytes = dataframe_to_excel_bytes(summary_df)
        st.download_button(
            label="Download Project Summary Excel",
            data=out_bytes,
            file_name="Project_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Optional GPT summary
        if use_gpt:
            st.subheader("GPT Summary (temperature=0, strict JSON)")
            # Create a trimmed view for "largest projects" selection
            df_for_gpt = summary_df.dropna(subset=["Total Amount"]).copy()
            df_for_gpt = df_for_gpt.sort_values("Total Amount", ascending=False).head(int(top_n))

            try:
                gpt_json = gpt_make_summary(df_for_gpt)
                st.json(gpt_json)
                st.markdown("**High-level:** " + gpt_json["high_level"])
            except Exception as e:
                st.error("GPT summary failed. Check your API key / model access.")
                st.exception(e)

st.caption("Deterministic extraction is used for the spreadsheet logic; GPT is optional for narrative only.")
