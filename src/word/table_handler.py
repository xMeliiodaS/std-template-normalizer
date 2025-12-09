from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml.shared import OxmlElement, qn


def set_landscape_for_all_sections(docx_path: str, output_path: str = None):
    """
    Force all sections to Landscape without toggling.
    - Sets orientation to LANDSCAPE.
    - Swaps width/height only if the current page is Portrait (width < height).
    """
    document = Document(docx_path)

    for section in document.sections:
        # Always set orientation
        section.orientation = WD_ORIENT.LANDSCAPE

        # Only swap if the page is currently portrait
        if section.page_width < section.page_height:
            section.page_width, section.page_height = section.page_height, section.page_width

    document.save(output_path or docx_path)


def set_tables_autofit_to_window(docx_path: str,output_path: str = None, clear_column_widths: bool = True):
    """
    Applies Word's 'Layout -> AutoFit -> AutoFit to Window' to all tables in a .docx:
      - Set table preferred width to 100% (pct=5000, meaning 100%).
      - Remove fixed table layout (w:tblLayout w:type='fixed') to enable AutoFit behavior.
      - Optionally remove cell-level fixed widths so Word can reflow columns.

    Args:
        docx_path: str - path to input .docx
        output_path: str | None - path to save; overwrites input if None
        clear_column_widths: bool - remove <w:tcW> from cells to allow true autofit
    """
    doc = Document(docx_path)

    for table in doc.tables:
        tbl = table._tbl  # <w:tbl>
        tblPr = tbl.tblPr
        if tblPr is None:
            # Ensure <w:tblPr> exists
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)

        # ---- Preferred width: 100% (AutoFit to Window) ----
        # Look for existing <w:tblW>; create if missing
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW')
            tblPr.append(tblW)
        # Set as percentage (fiftieths of a percent): 100% -> 5000
        tblW.set(qn('w:type'), 'pct')
        tblW.set(qn('w:w'), '5000')

        # ---- Remove fixed table layout to allow AutoFit ----
        # If <w:tblLayout w:type="fixed"> exists, remove it
        tblLayout = tblPr.find(qn('w:tblLayout'))
        if tblLayout is not None:
            tblPr.remove(tblLayout)

        # ---- Optional: Clear per-cell fixed widths ----
        if clear_column_widths:
            for row in table.rows:
                for cell in row.cells:
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    tcW = tcPr.find(qn('w:tcW'))
                    if tcW is not None:
                        tcPr.remove(tcW)

    doc.save(output_path or docx_path)


from copy import deepcopy
from docx import Document
from docx.oxml.shared import qn
from docx.table import Table


def _row_is_empty(row) -> bool:
    """Return True if all cell texts in the row are empty/whitespace."""
    for cell in row.cells:
        if cell.text and cell.text.strip():
            return False
    return True


def _table_header_matches(table: Table, expected_headers: list[str], case_insensitive: bool = True) -> bool:
    """
    Check whether the table's first row cells match expected header labels.
    Supports partial matching when len(expected_headers) <= number of cells.
    """
    if len(table.rows) == 0:
        return False
    header_cells = table.rows[0].cells
    if len(header_cells) < len(expected_headers):
        return False

    for i, exp in enumerate(expected_headers):
        actual = header_cells[i].text.strip() if i < len(header_cells) else ""
        if case_insensitive:
            if exp.strip().lower() not in actual.lower():
                return False
        else:
            if exp.strip() not in actual:
                return False
    return True


def _find_table_by_header(dst_doc: Document, expected_headers: list[str]) -> Table:
    """
    Return the first table whose header row matches expected_headers.
    Example expected_headers: ["ID"] or ["ID", "Name", "Description"].
    """
    for table in dst_doc.tables:
        if _table_header_matches(table, expected_headers):
            return table
    raise ValueError(f"No table found with header(s): {expected_headers}")


def copy_table_rows_excluding_header_into_table_with_id(
        src_docx_path: str,
        dst_docx_path: str,
        output_path: str = None,
        src_table_index: int = 0,
        expected_target_headers: list[str] = ["ID"]
):
    """
    Copy all rows except the header from a source table and paste into the target table
    identified by its header row (e.g., header contains 'ID').
    Insertion starts at the first empty row below the header; if none, rows are appended.

    Args:
        src_docx_path: path to source .docx (table to copy from)
        dst_docx_path: path to target .docx (table to paste into)
        output_path: where to save updated target; overwrites dst_docx_path if None
        src_table_index: 0-based index of source table
        expected_target_headers: list of header labels to identify target table
    """
    # Load docs
    src = Document(src_docx_path)
    dst = Document(dst_docx_path)

    # Validate source table
    if src_table_index < 0 or src_table_index >= len(src.tables):
        raise IndexError(f"Source table index {src_table_index} out of range. Source has {len(src.tables)} tables.")

    src_table = src.tables[src_table_index]
    src_tbl_elm = src_table._tbl

    # Gather source rows excluding header (skip first <w:tr>)
    src_rows = src_tbl_elm.findall(qn('w:tr'))
    if len(src_rows) < 2:
        raise ValueError("Source table must have at least 2 rows (header + data).")
    rows_to_copy = [deepcopy(tr) for tr in src_rows[1:]]

    # Defensive: clear repeating header flags on copied rows
    for tr in rows_to_copy:
        trPr = tr.find(qn('w:trPr'))
        if trPr is not None:
            hdr = trPr.find(qn('w:tblHeader'))
            if hdr is not None:
                trPr.remove(hdr)

    # Find target table by its header row (e.g., header first cell contains 'ID')
    target_table = _find_table_by_header(dst, expected_headers=expected_target_headers)
    target_tbl_elm = target_table._tbl

    # Identify first empty data row (below header)
    empty_row_obj = None
    data_rows = list(target_table.rows)[1:] if len(target_table.rows) > 1 else []
    for r in data_rows:
        if _row_is_empty(r):
            empty_row_obj = r
            break

    if empty_row_obj is not None:
        # Replace the empty row with the first copied row, then insert remaining
        empty_tr = empty_row_obj._tr
        empty_tr.addnext(rows_to_copy[0])
        target_tbl_elm.remove(empty_tr)
        cursor = rows_to_copy[0]
        for tr in rows_to_copy[1:]:
            cursor.addnext(tr)
            cursor = tr
    else:
        # No empty row found → append at end
        for tr in rows_to_copy:
            target_tbl_elm.append(tr)

    dst.save(output_path or dst_docx_path)
