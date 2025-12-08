from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml.shared import OxmlElement, qn


def set_landscape_for_all_sections(docx_path: str):
    """
    Sets page orientation to landscape for all sections in a .docx file.
    If output_path is not provided, the input document is overwritten.
    """
    document = Document(docx_path)

    for section in document.sections:
        # Set orientation
        section.orientation = WD_ORIENT.LANDSCAPE

        # Swap width/height to match landscape
        section.page_width, section.page_height = section.page_height, section.page_width

    document.save(docx_path)


def set_tables_autofit_to_window(docx_path: str, clear_column_widths: bool = True):
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

    doc.save(docx_path)
