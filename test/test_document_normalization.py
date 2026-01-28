import os
import unittest
from src.config.config_provider import ConfigProvider
from src.word.table_handler import (
    set_paragraph_spacing,
    set_table_column_widths,
    set_tables_autofit_to_window,
    set_landscape_for_all_sections,
    copy_table_rows_excluding_header_into_table_with_id
)
from src.word.placeholder_replacer import replace_placeholders_using_config


class TestProtocolNormalization(unittest.TestCase):
    """Unit test for normalizing Word protocols and replacing placeholders."""

    def setUp(self):
        """Load configuration and set file paths for the test."""
        self.config = ConfigProvider.load_config_json()
        self.exported_word = self.config["Exported_STD"]
        self.template_ready_word = self.config["Template_protocol"]
        self.output_word = self.config["Normalized_protocol"]

        # Ensure the output file has a .docx extension
        if not self.output_word.endswith('.docx'):
            self.output_word += '.docx'

    def test_document_normalization(self):
        """
        Validate that Word tables are normalized and placeholders are replaced.

        Steps:
        1. Set all sections to landscape orientation.
        2. Autofit all tables to window width.
        3. Copy rows from exported Word into the template (excluding headers).
        4. Set column widths for all tables.
        5. Adjust paragraph spacing.
        6. Replace placeholders using configuration.
        """
        # Set landscape layout for all sections
        set_landscape_for_all_sections(self.exported_word, self.output_word)

        # Make tables autofit to window
        set_tables_autofit_to_window(self.exported_word, self.output_word)

        # Copy rows (excluding header) into template table
        copy_table_rows_excluding_header_into_table_with_id(
            self.exported_word,
            self.template_ready_word,
            self.output_word
        )

        # Adjust table column widths
        set_table_column_widths(
            self.output_word,
            self.output_word,
            widths_cm=[1.67, 3.07, 10.0, 10.5, 3.25, 3.0, 3.0, 4.55]
        )

        # Adjust paragraph spacing in the document
        set_paragraph_spacing(self.output_word, self.output_word)

        # Replace placeholders with values from config
        replace_placeholders_using_config(self.output_word, self.output_word)


if __name__ == "__main__":
    # Define appdata path for storing config
    appdata_folder = os.path.join(
        os.environ.get('APPDATA', os.path.expanduser('~\\AppData\\Roaming')),
        "ste_tool_studio"
    )
    config_path = os.path.join(appdata_folder, "config.json")

    # Load config
    config = ConfigProvider.load_config_json(config_path)

    # Create test suite and run
    suite = unittest.TestSuite()
    suite.addTest(TestProtocolNormalization('test_document_normalization'))
    unittest.TextTestRunner(verbosity=2).run(suite)
