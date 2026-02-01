import json
import os
from docx import Document

from src.config.config_provider import ConfigProvider
from src.config.constants import DOCX_EXTENSION, WordPlaceholders, ConfigKeys


def _replace_text_in_paragraph(paragraph, replacements: dict):
    original_text = paragraph.text
    new_text = original_text

    for placeholder, value in replacements.items():
        if placeholder in new_text:
            new_text = new_text.replace(placeholder, value)

    if new_text == original_text:
        return  # nothing to change

    # Preserve style from the first run (Word standard practice)
    style_run = paragraph.runs[0] if paragraph.runs else None

    paragraph.clear()
    new_run = paragraph.add_run(new_text)

    if style_run:
        new_run.bold = style_run.bold
        new_run.italic = style_run.italic
        new_run.underline = style_run.underline
        new_run.font.name = style_run.font.name
        new_run.font.size = style_run.font.size
        new_run.font.color.rgb = style_run.font.color.rgb



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

    doc = Document(docx_path)

    replacements = {
        WordPlaceholders.DOC_STD: config.get(ConfigKeys.DOC_STD, config.get(ConfigKeys.LEGACY_KEYS["DOC_STD"], "")),
        WordPlaceholders.STD_NAME: config.get(ConfigKeys.STD_NAME, config.get(ConfigKeys.LEGACY_KEYS["STD_NAME"], "")),
        WordPlaceholders.PLAN_NUMBER: config.get(ConfigKeys.TEST_PLAN, config.get(ConfigKeys.LEGACY_KEYS["PLAN_NUMBER"], "")),
        WordPlaceholders.PREPARED_BY: config.get(ConfigKeys.PREPARED_BY, config.get(ConfigKeys.LEGACY_KEYS["PREPARED_BY"], "")),
        WordPlaceholders.TEST_PROTOCOL: config.get(ConfigKeys.TEST_PROTOCOL, config.get(ConfigKeys.LEGACY_KEYS["TEST_PROTOCOL"], "")),
        WordPlaceholders.FOOTER: config.get(ConfigKeys.FOOTER, config.get(ConfigKeys.LEGACY_KEYS["FOOTER"], "")),
    }

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
