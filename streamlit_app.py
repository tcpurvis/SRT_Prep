import streamlit as st
import io
import re
from copy import copy
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock

st.set_page_config(page_title="Transcript Cleaner", layout="centered")
st.title("Transcript Cleaner")
st.caption("Cleans Source column and trims to TC In / TC Out / Source only.")

uploaded = st.file_uploader("Upload transcript .xlsx", type=["xlsx"])

KEEP_COLS = {"TC In", "TC Out", "Source"}


def process_workbook(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes), rich_text=True)
    ws = wb.active

    # --- Map header names to column indices (1-based) ---
    headers = {cell.value: cell.column for cell in ws[1] if cell.value}

    source_col = headers.get("Source")
    if source_col is None:
        raise ValueError("Could not find a 'Source' column in the spreadsheet.")

    # --- Process Source column: strip strikethrough, replace <> with [] ---
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[source_col - 1]
        val = cell.value

        if val is None:
            continue

        # Rich text (list of TextBlock objects)
        if isinstance(val, CellRichText):
            cleaned_blocks = []
            for block in val:
                if isinstance(block, TextBlock):
                    # Drop the whole run if strikethrough
                    if block.font and block.font.strike:
                        continue
                    # Replace <...> with [...] in remaining runs
                    new_text = re.sub(r"<([^>]*)>", r"[\1]", block.text)
                    new_block = TextBlock(copy(block.font), new_text)
                    cleaned_blocks.append(new_block)
                else:
                    # Plain string segment inside rich text
                    cleaned_blocks.append(re.sub(r"<([^>]*)>", r"[\1]", str(block)))
            cell.value = CellRichText(cleaned_blocks) if cleaned_blocks else None

        # Plain string
        elif isinstance(val, str):
            # Check cell-level strikethrough (entire cell is struck)
            if cell.font and cell.font.strike:
                cell.value = None
            else:
                cell.value = re.sub(r"<([^>]*)>", r"[\1]", val)

    # --- Delete columns that aren't TC In / TC Out / Source ---
    # Collect column indices to delete (high → low to avoid shifting issues)
    cols_to_delete = sorted(
        [col for name, col in headers.items() if name not in KEEP_COLS],
        reverse=True,
    )
    for col_idx in cols_to_delete:
        ws.delete_cols(col_idx)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


if uploaded:
    raw = uploaded.read()
    try:
        result = process_workbook(raw)
        st.success("File processed successfully!")
        st.download_button(
            label="⬇️ Download cleaned file",
            data=result,
            file_name=f"cleaned_{uploaded.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Something went wrong: {e}")
