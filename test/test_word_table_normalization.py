import os, unittest
from src.config.config_provider import ConfigProvider
from src.word.table_handler import (set_paragraph_spacing,
                                    set_table_column_widths,
                                    set_tables_autofit_to_window,
                                    set_landscape_for_all_sections,
                                    copy_table_rows_excluding_header_into_table_with_id)
from src.word.placeholder_replacer import replace_placeholders_using_config


class TestWordTableNormalization(unittest.TestCase):
    def setUp(self):
        self.config = ConfigProvider.load_config_json()
        self.exported_word = self.config["Exported_STD"]
        self.template_ready_word = self.config["Template_protocol"]
        self.output_word = self.config["Normalized_protocol"]
        # Ensure output_word has .docx extension
        if not self.output_word.endswith('.docx'):
            self.output_word = self.output_word + '.docx'

    def test_word_table_normalization(self):
        """Validate that Excel is consistent and generate violations HTML report."""
        set_landscape_for_all_sections(self.exported_word, self.output_word)

        set_tables_autofit_to_window(self.exported_word, self.output_word)

        copy_table_rows_excluding_header_into_table_with_id(
            self.exported_word,
            self.template_ready_word,
            self.output_word
        )

        set_table_column_widths(
            self.output_word,
            self.output_word,
            widths_cm=[1.67, 3.07, 10.0, 10.5, 3.25, 3.0, 3.0, 4.55]
            # table_index=3  # Change this based on what the debug shows
        )

        set_paragraph_spacing(self.output_word, self.output_word)

        replace_placeholders_using_config(self.output_word, self.output_word)

if __name__ == "__main__":
    appdata_folder = os.path.join(
        os.environ.get('APPDATA', os.path.expanduser('~\\AppData\\Roaming')),
        "ste_tool_studio"
    )
    config_path = os.path.join(appdata_folder, "config.json")

    config = ConfigProvider.load_config_json(config_path)

    suite = unittest.TestSuite()
    suite.addTest(TestWordTableNormalization('test_word_table_normalization'))
    unittest.TextTestRunner(verbosity=2).run(suite)
