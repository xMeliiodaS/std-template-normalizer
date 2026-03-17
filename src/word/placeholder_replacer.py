import os
from docx import Document

from src.config.config_provider import ConfigProvider
from src.config.constants import DOCX_EXTENSION, WordPlaceholders, ConfigKeys
from src.config.logging_config import get_logger

logger = get_logger(__name__)


# doc_type from external config -> replacement values for DOC_TYPE, DOC_RECORD, DOC_TYPE_STx
_DOC_TYPE_REPLACEMENT_MAP = {
    "protocol": {
        # Protocol mode
        # ADD_DOC_TYPE -> Design
        WordPlaceholders.DOC_TYPE: "Design",
        # ADD_DOC_RECORD -> Protocol
        WordPlaceholders.DOC_RECORD: "Protocol",
        # ADD_DOC_STX -> STD
        WordPlaceholders.DOC_TYPE_STx: "(STD)",
    },
    "report": {
        # Report mode
        # ADD_DOC_TYPE -> Report
        WordPlaceholders.DOC_TYPE: "Report",
        # ADD_DOC_RECORD -> Report
        WordPlaceholders.DOC_RECORD: "Report",
        # ADD_DOC_STX -> STR
        WordPlaceholders.DOC_TYPE_STx: "(STR)",
    },
}


def get_doc_type_replacements(doc_type: str):
    """
    Get replacement values for DOC_TYPE, DOC_RECORD, and DOC_TYPE_STx placeholders
    based on the doc_type from the external config (e.g. "protocol", "report").

    :param doc_type: Value of "doc_type" from config (e.g. "protocol", "report").
    :return: Dict mapping placeholder keys to replacement strings, or None if doc_type is not mapped.
    """
    if not doc_type:
        return None
    return _DOC_TYPE_REPLACEMENT_MAP.get(doc_type.strip().lower())


def _replace_text_in_paragraph(paragraph, replacements: dict):
    if not paragraph.runs:
        return
    # Mutate only the run text that directly contains a placeholder.
    # This avoids reflowing text between runs and preserves line breaks,
    # spacing, and run-level formatting outside the exact replacement span.
    for run in paragraph.runs:
        run_text = run.text
        new_run_text = run_text

        for placeholder, value in replacements.items():
            if placeholder in new_run_text:
                new_run_text = new_run_text.replace(placeholder, value)

        if new_run_text != run_text:
            run.text = new_run_text

    def _replace_token_across_runs(token: str, replacement: str):
        if not token:
            return

        while True:
            run_texts = [run.text for run in paragraph.runs]
            full_text = "".join(run_texts)
            start_idx = full_text.find(token)
            if start_idx == -1:
                return

            end_idx = start_idx + len(token)

            # Map absolute paragraph offsets to run index and run-local offset.
            cumulative = 0
            start_run_idx = start_off = end_run_idx = end_off = 0
            start_found = end_found = False

            for idx, run_text in enumerate(run_texts):
                next_cumulative = cumulative + len(run_text)

                if not start_found and start_idx < next_cumulative:
                    start_run_idx = idx
                    start_off = start_idx - cumulative
                    start_found = True

                if not end_found and end_idx <= next_cumulative:
                    end_run_idx = idx
                    end_off = end_idx - cumulative
                    end_found = True
                    break

                cumulative = next_cumulative

            if not (start_found and end_found):
                return

            if start_run_idx == end_run_idx:
                run = paragraph.runs[start_run_idx]
                run.text = run.text[:start_off] + replacement + run.text[end_off:]
                continue

            start_run = paragraph.runs[start_run_idx]
            end_run = paragraph.runs[end_run_idx]

            start_prefix = start_run.text[:start_off]
            end_suffix = end_run.text[end_off:]

            start_run.text = start_prefix + replacement

            for mid_idx in range(start_run_idx + 1, end_run_idx):
                paragraph.runs[mid_idx].text = ""

            end_run.text = end_suffix

    # Replace placeholders individually while preserving existing run structure.
    for placeholder, value in replacements.items():
        _replace_token_across_runs(placeholder, value)



def replace_text_in_table(table, replacements: dict):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _replace_text_in_paragraph(paragraph, replacements)


def replace_placeholders_using_config(docx_path, output_path=None):
    config = ConfigProvider.load_config_json()

    # Ensure output_path has .docx extension if it doesn't
    if output_path and not output_path.endswith(DOCX_EXTENSION):
        output_path = output_path + DOCX_EXTENSION

    # Ensure docx_path exists and has .docx extension
    if not docx_path.endswith(DOCX_EXTENSION):
        if os.path.exists(docx_path + DOCX_EXTENSION):
            docx_path = docx_path + DOCX_EXTENSION
        else:
            raise ValueError(f"Document path must be a {DOCX_EXTENSION} file: {docx_path}")

    out = output_path or docx_path
    logger.info("Replacing placeholders. Input: %s, Output: %s", docx_path, out)

    doc = Document(docx_path)

    # Values from C# Template Normalizer (config key = field name → Word placeholder)
    protocol_number = config.get(ConfigKeys.PROTOCOL_NUMBER) or config.get(ConfigKeys.LEGACY_KEYS["DOC_STD"]) or ""
    stx_number = config.get(ConfigKeys.STX_NUMBER) or config.get(ConfigKeys.LEGACY_KEYS["STX_NUMBER"]) or ""
    stx_number = f"({stx_number})"
    protocol_number_display = f"{protocol_number}"
    std_name = config.get(ConfigKeys.STD_NAME) or config.get(ConfigKeys.LEGACY_KEYS["STD_NAME"]) or ""
    report_number = config.get(ConfigKeys.REPORT_NUMBER) or config.get(ConfigKeys.LEGACY_KEYS["REPORT_NUMBER"]) or ""
    test_plan = config.get(ConfigKeys.TEST_PLAN) or config.get(ConfigKeys.LEGACY_KEYS["PLAN_NUMBER"]) or ""
    prepared_by = config.get(ConfigKeys.PREPARED_BY) or config.get(ConfigKeys.LEGACY_KEYS["PREPARED_BY"]) or ""
    footer = config.get(ConfigKeys.FOOTER) or config.get(ConfigKeys.LEGACY_KEYS["FOOTER"]) or ""

    is_report = (config.get(ConfigKeys.DOC_TYPE) or "").strip().lower() == "report"

    add_doc_std_value = report_number if is_report else protocol_number_display

    replacements = {
        WordPlaceholders.DOC_TYPE: config.get(ConfigKeys.DOC_TYPE) or config.get(ConfigKeys.LEGACY_KEYS["DOC_TYPE"]) or "",
        WordPlaceholders.DOC_TYPE_STx: config.get(ConfigKeys.DOC_STX) or config.get(ConfigKeys.LEGACY_KEYS["DOC_TYPE_STX"]) or "",
        WordPlaceholders.DOC_RECORD: config.get(ConfigKeys.DOC_RECORD) or config.get(ConfigKeys.LEGACY_KEYS["DOC_RECORD"]) or "",
        WordPlaceholders.PROTOCOL_NUMBER: protocol_number_display,
        WordPlaceholders.REPORT_NUMBER: report_number,
        WordPlaceholders.STD_NAME: std_name,
        WordPlaceholders.PLAN_NUMBER: test_plan,
        WordPlaceholders.STX_NUMBER: stx_number,
        WordPlaceholders.PREPARED_BY: prepared_by,
        WordPlaceholders.FOOTER: footer,
        # Legacy placeholders (same values)

        "ADD_DOC_STD#": add_doc_std_value,
    }

    # Override DOC_TYPE, DOC_RECORD, DOC_TYPE_STx when doc_type is "protocol" or "report"
    doc_type_from_config = config.get(ConfigKeys.DOC_TYPE) or config.get(ConfigKeys.LEGACY_KEYS["DOC_TYPE"])
    doc_type_replacements = get_doc_type_replacements(doc_type_from_config)
    if doc_type_replacements:
        replacements.update(doc_type_replacements)

    # ---- Body ----
    for paragraph in doc.paragraphs:
        _replace_text_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        replace_text_in_table(table, replacements)

    # ---- Headers & Footers ----
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            _replace_text_in_paragraph(paragraph, replacements)

        for table in section.header.tables:
            replace_text_in_table(table, replacements)

        for paragraph in section.footer.paragraphs:
            _replace_text_in_paragraph(paragraph, replacements)

        for table in section.footer.tables:
            replace_text_in_table(table, replacements)

    # Save document
    doc.save(output_path or docx_path)
    logger.info("Output: Placeholders replaced successfully.")
