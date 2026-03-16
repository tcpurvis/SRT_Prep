import streamlit as st
import io
import re
import zipfile
from lxml import etree
from copy import deepcopy

st.set_page_config(page_title="Transcript Cleaner", layout="centered")
st.title("Transcript Cleaner")
st.caption("Cleans Source column and trims to TC In / TC Out / Source only.")

uploaded = st.file_uploader("Upload transcript .xlsx", type=["xlsx"])

NS           = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
TAG          = f"{{{NS}}}"
PALE_YELLOW  = "FFFFFF99"
SOURCE_WIDTH = 72.0
SOURCE_XF    = 2        # xf index used by Source column in original file


def get_si_text(si):
    runs = si.findall(f"{TAG}r")
    if runs:
        return "".join((r.find(f"{TAG}t").text or "") for r in runs if r.find(f"{TAG}t") is not None)
    t = si.find(f"{TAG}t")
    return (t.text or "") if t is not None else ""


def add_styles(styles_xml: bytes):
    """
    Adds pale-yellow fill + two new xf entries (wrap-only, wrap+yellow).
    Returns (new_styles_xml, xf_wrap_idx, xf_wrap_yellow_idx).
    """
    tree = etree.fromstring(styles_xml)

    # Add pale yellow fill
    fills_el = tree.find(f"{TAG}fills")
    yellow_fill = etree.SubElement(fills_el, f"{TAG}fill")
    pf = etree.SubElement(yellow_fill, f"{TAG}patternFill")
    pf.set("patternType", "solid")
    etree.SubElement(pf, f"{TAG}fgColor").set("rgb", PALE_YELLOW)
    etree.SubElement(pf, f"{TAG}bgColor").set("indexed", "64")
    yellow_fill_idx = len(fills_el) - 1
    fills_el.set("count", str(len(fills_el)))

    # Clone source xf as base template
    xfs_el = tree.find(f"{TAG}cellXfs")
    source_xf = list(xfs_el)[SOURCE_XF]

    # xf_wrap: plain clone (wrapText already set on source xf)
    xfs_el.append(deepcopy(source_xf))
    xf_wrap_idx = len(xfs_el) - 1

    # xf_wrap_yellow: clone with pale yellow fill
    xf_yellow = deepcopy(source_xf)
    xf_yellow.set("fillId", str(yellow_fill_idx))
    xf_yellow.set("applyFill", "1")
    xfs_el.append(xf_yellow)
    xf_wrap_yellow_idx = len(xfs_el) - 1

    xfs_el.set("count", str(len(xfs_el)))
    return (
        etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True),
        xf_wrap_idx,
        xf_wrap_yellow_idx,
    )


def process_workbook(file_bytes: bytes):
    all_files = {}
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        for name in z.namelist():
            all_files[name] = z.read(name)

    if "xl/sharedStrings.xml" not in all_files:
        raise ValueError("No shared strings found — is this a valid .xlsx transcript file?")

    sheet_tree = etree.fromstring(all_files["xl/worksheets/sheet1.xml"])
    ss_tree    = etree.fromstring(all_files["xl/sharedStrings.xml"])
    ss_list    = ss_tree.findall(f"{TAG}si")
    rows       = sheet_tree.findall(f".//{TAG}row")

    if not rows:
        raise ValueError("Worksheet appears to be empty.")

    # Detect column letters from header row
    header_map = {}
    for cell in rows[0].findall(f"{TAG}c"):
        col_letter = re.match(r"^([A-Z]+)", cell.get("r", "")).group(1)
        v = cell.find(f"{TAG}v")
        if v is not None and cell.get("t") == "s":
            header_map[get_si_text(ss_list[int(v.text)])] = col_letter

    needed = {"TC In", "TC Out", "Source"}
    missing = needed - set(header_map)
    if missing:
        raise ValueError(f"Could not find column(s): {', '.join(missing)}")

    source_col = header_map["Source"]
    keep_cols  = {header_map[n] for n in needed}

    # Map shared string index -> row numbers that reference it (Source col only)
    source_idx_to_rows = {}
    for row in rows[1:]:
        row_num = int(row.get("r", "0"))
        for cell in row.findall(f"{TAG}c"):
            col = re.match(r"^([A-Z]+)", cell.get("r", "")).group(1)
            if col == source_col and cell.get("t") == "s":
                v = cell.find(f"{TAG}v")
                if v is not None:
                    source_idx_to_rows.setdefault(int(v.text), set()).add(row_num)

    # Modify shared strings; track which indices changed
    modified_ss_indices = set()
    for idx, row_set in source_idx_to_rows.items():
        si = ss_list[idx]
        runs = si.findall(f"{TAG}r")
        changed = False
        if runs:
            for r in list(runs):
                rpr = r.find(f"{TAG}rPr")
                if rpr is not None and rpr.find(f"{TAG}strike") is not None:
                    si.remove(r)
                    changed = True
                    continue
                t = r.find(f"{TAG}t")
                if t is not None and t.text:
                    new_text = re.sub(r"<([^>]*)>", r"[\1]", t.text)
                    if new_text != t.text:
                        t.text = new_text
                        changed = True
        else:
            t = si.find(f"{TAG}t")
            if t is not None and t.text:
                new_text = re.sub(r"<([^>]*)>", r"[\1]", t.text)
                if new_text != t.text:
                    t.text = new_text
                    changed = True
        if changed:
            modified_ss_indices.add(idx)

    # Add new styles
    new_styles_xml, xf_wrap, xf_wrap_yellow = add_styles(all_files["xl/styles.xml"])
    all_files["xl/styles.xml"] = new_styles_xml

    # Drop non-keep columns
    for row in rows:
        for cell in list(row.findall(f"{TAG}c")):
            col = re.match(r"^([A-Z]+)", cell.get("r", "")).group(1)
            if col not in keep_cols:
                row.remove(cell)

    # Remap column letters to A, B, C
    sorted_keep = sorted(keep_cols)
    col_remap = {old: chr(ord("A") + i) for i, old in enumerate(sorted_keep)}

    for row in rows:
        for cell in row.findall(f"{TAG}c"):
            old_ref = cell.get("r", "")
            col = re.match(r"^([A-Z]+)", old_ref).group(1)
            row_num_str = re.search(r"(\d+)$", old_ref).group(1)
            cell.set("r", f"{col_remap.get(col, col)}{row_num_str}")

            # Apply wrap / highlight styles to Source column data cells
            if col == source_col and row_num_str != "1":
                v = cell.find(f"{TAG}v")
                ss_idx = int(v.text) if v is not None and cell.get("t") == "s" else None
                cell.set("s", str(xf_wrap_yellow if ss_idx in modified_ss_indices else xf_wrap))

    # Rebuild <cols> with correct widths for new A/B/C layout
    cols_el = sheet_tree.find(f"{TAG}cols")
    if cols_el is not None:
        cols_el.getparent().remove(cols_el)

    sheet_data = sheet_tree.find(f"{TAG}sheetData")
    cols_new = etree.Element(f"{TAG}cols")
    tc_col = etree.SubElement(cols_new, f"{TAG}col")
    tc_col.set("min", "1"); tc_col.set("max", "2")
    tc_col.set("width", "14.0"); tc_col.set("customWidth", "1")
    src_col = etree.SubElement(cols_new, f"{TAG}col")
    src_col.set("min", "3"); src_col.set("max", "3")
    src_col.set("width", str(SOURCE_WIDTH)); src_col.set("customWidth", "1")
    sheet_data.addprevious(cols_new)

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
    return buf.read(), len(modified_ss_indices)


if uploaded:
    raw = uploaded.read()
    try:
        result, n_changed = process_workbook(raw)
        st.success(f"Done — {n_changed} rows revised and highlighted.")
        st.download_button(
            label="⬇️ Download cleaned file",
            data=result,
            file_name=f"cleaned_{uploaded.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Something went wrong: {e}")
