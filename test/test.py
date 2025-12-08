import os, unittest

from src.config.config_provider import ConfigProvider
from src.word.table_handler import (set_landscape_for_all_sections,
                                    set_tables_autofit_to_window,
                                    copy_table_rows_excluding_header_into_table_with_id)


class TestExcelViolations(unittest.TestCase):
    def setUp(self):
        self.config = ConfigProvider.load_config_json()
        self.exported_word = self.config["Exported_word"]
        self.template_ready_word = self.config["Template_word"]

    def test_excel_violations(self):
        """Validate that Excel is consistent and generate violations HTML report."""
        set_landscape_for_all_sections(self.exported_word)
        # set_tables_autofit_to_window(self.exported_word)
        # copy_table_rows_excluding_header_into_table_with_id(self.exported_word, self.template_ready_word)

if __name__ == "__main__":
    appdata_folder = os.path.join(
        os.environ.get('APPDATA', os.path.expanduser('~\\AppData\\Roaming')),
        "TO_BE_CHANGED"
    )
    config_path = os.path.join(appdata_folder, "config.json")

    config = ConfigProvider.load_config_json(config_path)

    suite = unittest.TestSuite()
    suite.addTest(TestExcelViolations('test_excel_violations'))
    unittest.TextTestRunner(verbosity=2).run(suite)
