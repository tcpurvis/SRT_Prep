import streamlit as st
import io
import re
import zipfile
from lxml import etree

st.set_page_config(page_title="Transcript Cleaner", layout="centered")
st.title("Transcript Cleaner")
st.caption("Cleans Source column and trims to TC In / TC Out / Source only.")

uploaded = st.file_uploader("Upload transcript .xlsx", type=["xlsx"])

NS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
TAG = f"{{{NS}}}"


def get_si_text(si):
    runs = si.findall(f"{TAG}r")
    if runs:
        return "".join((r.find(f"{TAG}t").text or "") for r in runs if r.find(f"{TAG}t") is not None)
    t = si.find(f"{TAG}t")
    return (t.text or "") if t is not None else ""


def process_workbook(file_bytes: bytes) -> bytes:
    all_files = {}
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        for name in z.namelist():
            all_files[name] = z.read(name)

    if "xl/sharedStrings.xml" not in all_files:
        raise ValueError("No shared strings found — is this a valid .xlsx transcript file?")

    sheet_tree = etree.fromstring(all_files["xl/worksheets/sheet1.xml"])
    ss_tree    = etree.fromstring(all_files["xl/sharedStrings.xml"])
    ss_list    = ss_tree.findall(f"{TAG}si")

    rows = sheet_tree.findall(f".//{TAG}row")
    if not rows:
        raise ValueError("Worksheet appears to be empty.")

    # --- Detect column letters for TC In, TC Out, Source from header row ---
    header_map = {}
    for cell in rows[0].findall(f"{TAG}c"):
        ref = cell.get("r", "")
        col_letter = re.match(r"^([A-Z]+)", ref).group(1)
        v = cell.find(f"{TAG}v")
        if v is not None and cell.get("t") == "s":
            header_map[get_si_text(ss_list[int(v.text)])] = col_letter

    needed = {"TC In", "TC Out", "Source"}
    missing = needed - set(header_map)
    if missing:
        raise ValueError(f"Could not find column(s): {', '.join(missing)}")

    source_col = header_map["Source"]
    keep_cols  = {header_map[n] for n in needed}

    # --- Collect shared string indices used by the Source column ---
    source_indices = set()
    for row in rows[1:]:
        for cell in row.findall(f"{TAG}c"):
            col = re.match(r"^([A-Z]+)", cell.get("r", "")).group(1)
            if col == source_col and cell.get("t") == "s":
                v = cell.find(f"{TAG}v")
                if v is not None:
                    source_indices.add(int(v.text))

    # --- Modify shared strings: remove strikethrough runs, replace <> with [] ---
    for idx in source_indices:
        si = ss_list[idx]
        runs = si.findall(f"{TAG}r")
        if runs:
            for r in list(runs):
                rpr = r.find(f"{TAG}rPr")
                if rpr is not None and rpr.find(f"{TAG}strike") is not None:
                    si.remove(r)
                    continue
                t = r.find(f"{TAG}t")
                if t is not None and t.text:
                    t.text = re.sub(r"<([^>]*)>", r"[\1]", t.text)
        else:
            t = si.find(f"{TAG}t")
            if t is not None and t.text:
                t.text = re.sub(r"<([^>]*)>", r"[\1]", t.text)

    # --- Remove non-keep columns from worksheet rows ---
    all_cols = set()
    for row in rows:
        for cell in row.findall(f"{TAG}c"):
            col = re.match(r"^([A-Z]+)", cell.get("r", "")).group(1)
            all_cols.add(col)

    for row in rows:
        for cell in list(row.findall(f"{TAG}c")):
            col = re.match(r"^([A-Z]+)", cell.get("r", "")).group(1)
            if col not in keep_cols:
                row.remove(cell)

    # --- Renumber column references to A, B, C ---
    sorted_keep = sorted(keep_cols)
    col_remap = {old: chr(ord("A") + i) for i, old in enumerate(sorted_keep)}

    for row in rows:
        for cell in row.findall(f"{TAG}c"):
            old_ref = cell.get("r", "")
            col = re.match(r"^([A-Z]+)", old_ref).group(1)
            row_num = re.search(r"(\d+)$", old_ref).group(1)
            cell.set("r", f"{col_remap.get(col, col)}{row_num}")

    # --- Serialize and repack ---
    new_sheet_xml = etree.tostring(sheet_tree, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_ss_xml    = etree.tostring(ss_tree,    xml_declaration=True, encoding="UTF-8", standalone=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            if name == "xl/worksheets/sheet1.xml":
                zout.writestr(name, new_sheet_xml)
            elif name == "xl/sharedStrings.xml":
                zout.writestr(name, new_ss_xml)
            else:
                zout.writestr(name, data)

    buf.seek(0)
    return buf.read()


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
