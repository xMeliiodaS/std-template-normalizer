"""
Constants module for the VSTS Azure Automation project.
Centralizes all hardcoded values to improve maintainability and clean code practices.
"""

# ============================================================================
# Application Configuration
# ============================================================================

APP_DATA_FOLDER_NAME = "ste_tool_studio"
CONFIG_FILE_NAME = "config.json"

# ============================================================================
# File Extensions
# ============================================================================

DOCX_EXTENSION = ".docx"
XLSX_EXTENSION = ".xlsx"

# ============================================================================
# Default Configuration Keys
# ============================================================================

class ConfigKeys:
    """Centralized JSON configuration keys."""

    EXPORTED_STD = "Exported_STD"
    TEMPLATE_PROTOCOL = "Template_protocol"
    NORMALIZED_PROTOCOL = "Normalized_protocol"

    # Config keys = C# Template Normalizer field names. Word placeholder in comments.
    DOC_TYPE = "doc_type"
    DOC_STX = "doc_stx"
    DOC_RECORD = "doc_record"
    PROTOCOL_NUMBER = "protocol_number"   # → ADD_PROTOCOL_NUMBER#
    STD_NAME = "std_name"                 # → ADD_STD_NAME
    REPORT_NUMBER = "report_number"       # → ADD_REPORT_NUMBER
    TEST_PLAN = "test_plan"               # → ADD_PLAN_NUMBER
    STX_NUMBER = "stx_number"             # → ADD_STX_NUMBER
    PREPARED_BY = "prepared_by"           # → ADD_PREPARED_BY
    # Legacy / optional
    FOOTER = "footer"
    TEST_PROTOCOL = "test_plan"

    LEGACY_KEYS = {
        "DOC_TYPE": "Doc_type",
        "DOC_TYPE_STX": "Doc_stx",
        "DOC_RECORD": "Doc_record",
        "DOC_STD": "doc_number",          # legacy key for Protocol number
        "STD_NAME": "STD_name",
        "PLAN_NUMBER": "PLAN-number",
        "PREPARED_BY": "Prepared_by",
        "TEST_PROTOCOL": "Test_protocol",
        "FOOTER": "Footer",
        "REPORT_NUMBER": "Report_number",
        "STX_NUMBER": "STx_number",
    }

# ============================================================================
# Word Placeholder Tokens
# ============================================================================

class WordPlaceholders:
    """Word placeholder tokens. Field (C#) = config key = placeholder name."""

    DOC_TYPE = "ADD_DOC_TYPE"              # doc_type → Design / Report
    DOC_TYPE_STx = "ADD_DOC_STX"           # doc_type → STD / STR
    DOC_RECORD = "ADD_DOC_RECORD"          # doc_type → Protocol / Report

    PROTOCOL_NUMBER = "ADD_PROTOCOL_NUMBER#"   # protocol_number
    REPORT_NUMBER = "ADD_REPORT_NUMBER"        # report_number
    STD_NAME = "ADD_STD_NAME"                  # std_name
    PLAN_NUMBER = "ADD_PLAN_NUMBER"            # test_plan
    STX_NUMBER = "ADD_STX_NUMBER"              # stx_number
    PREPARED_BY = "ADD_PREPARED_BY"            # prepared_by
    FOOTER = "ADD_FOOTER"                      # footer (optional)

# ============================================================================
# Word Table Handling Constants
# ============================================================================

class WordTableDefaults:
    """Defaults for Word table handling and normalization."""

    DEFAULT_TARGET_HEADERS = ["ID"]
    DEFAULT_PARAGRAPH_SPACING_BEFORE_PT = 0
    DEFAULT_PARAGRAPH_SPACING_AFTER_PT = 3

    DEFAULT_COLUMN_WIDTHS_CM = [1.67, 3.07, 10.0, 10.5, 3.25, 3.0, 3.0, 4.55]

# ============================================================================
# Word Orientation & Layout
# ============================================================================

class WordLayout:
    """Page layout and formatting defaults."""

    AUTOFIT_TABLE_PERCENT = "5000"
    AUTOFIT_TABLE_TYPE = "pct"
    FIXED_LAYOUT_TYPE = "fixed"

# ============================================================================
# XML Namespaces
# ============================================================================

class XmlTags:
    """OpenXML element tags."""

    TABLE_PROPERTIES = "w:tblPr"
    TABLE_WIDTH = "w:tblW"
    TABLE_LAYOUT = "w:tblLayout"
    GRID_COLUMN = "w:gridCol"
    TABLE_GRID = "w:tblGrid"
    ROW = "w:tr"
    ROW_PROPERTIES = "w:trPr"
    TABLE_HEADER = "w:tblHeader"
    CELL_WIDTH = "w:tcW"
    PARAGRAPH_PROPERTIES = "w:pPr"
    NUMBERING_PROPERTIES = "w:numPr"
    OUTLINE_LEVEL = "w:outlineLvl"
    PARAGRAPH_STYLE = "w:pStyle"

# ============================================================================
# Unit Test Configuration
# ============================================================================

class TestDefaults:
    """Unit testing defaults."""

    TEST_METHOD = "test_document_normalization"
    TEST_RUNNER_VERBOSITY = 2

# ============================================================================
# Error Messages
# ============================================================================

class ErrorMessages:
    """Standardized exception messages."""

    CONFIG_READ_ERROR = "Error reading config.json"
    DOCX_REQUIRED = "Document path must be a .docx file"
    TABLE_NOT_FOUND = "Table with headers not found"
    INVALID_WIDTHS = "widths_cm must be a non-empty list of column widths"
    COLUMN_MISMATCH = "Width count must match number of table columns"
    SOURCE_TABLE_INDEX_ERROR = "Source table index out of range"
    SOURCE_TABLE_TOO_SMALL = "Source table must have at least 2 rows"

# ============================================================================
# Logging & Debug Output
# ============================================================================

class LogMessages:
    """Centralized log and debug output messages."""

    DEFAULT_CONFIG_CREATED = "Default config.json not found. Creating empty config."


# ============================================================================
# Process output markers for C# StatusText (stdout = status, use these if needed)
# ============================================================================

class ProcessMarkers:
    """Markers for C# ProcessExecutionService progress parsing."""

    PROGRESS_TOTAL = "PROGRESS_TOTAL:"
    PROGRESS = "PROGRESS:"
    PROCESS_FINISHED = "PROCESS_FINISHED"
