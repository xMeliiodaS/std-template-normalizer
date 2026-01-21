import json
import os
from docx import Document

from src.config.config_provider import ConfigProvider


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
    if output_path and not output_path.endswith('.docx'):
        output_path = output_path + '.docx'
    
    # Ensure docx_path exists and has .docx extension
    if not docx_path.endswith('.docx'):
        if os.path.exists(docx_path + '.docx'):
            docx_path = docx_path + '.docx'
        else:
            raise ValueError(f"Document path must be a .docx file: {docx_path}")
    
    doc = Document(docx_path)

    replacements = {
        "ADD_DOC_STD#": config.get("doc_number", config.get("DOC_STD", "")),
        "ADD_STD_NAME": config.get("std_name", config.get("STD_name", "")),
        "ADD_PLAN_NUMBER": config.get("test_plan", config.get("PLAN-number", "")),
        "ADD_PREPARED_BY": config.get("prepared_by", config.get("Prepared_by", "")),
        "ADD_TEST_PROTOCOL": config.get("test_plan", config.get("Test_protocol", "")),
        "ADD_FOOTER": config.get("footer", config.get("Footer", "")),
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
