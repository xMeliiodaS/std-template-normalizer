from typing import Dict, Optional

from src.config.config_provider import ConfigProvider
from src.config.constants import ConfigKeys, WordPlaceholders


DOC_STD_NUMBER = "ADD_DOC_STD#"
DELETE_ROW_MARKER = "TO_BE_DELETED_ROW"


_DOC_TYPE_REPLACEMENTS = {
    "protocol": {
        WordPlaceholders.DOC_TYPE: "Design",
        WordPlaceholders.DOC_RECORD: "Protocol",
        WordPlaceholders.DOC_TYPE_STx: "(STD)",
        WordPlaceholders.STX_TYPE: "STD",
    },
    "report": {
        WordPlaceholders.DOC_TYPE: "Report",
        WordPlaceholders.DOC_RECORD: "Report",
        WordPlaceholders.DOC_TYPE_STx: "(STR)",
        WordPlaceholders.STX_TYPE: "STR",
    },
}


def get_config_value(
    config: dict,
    key: str,
    legacy_key_name: Optional[str] = None,
) -> str:
    """
    Return the modern config value first.

    If the modern value is missing or empty, return the configured
    legacy value. Always return a string.
    """
    value = config.get(key)

    if value not in (None, ""):
        return str(value)

    if legacy_key_name:
        legacy_key = ConfigKeys.LEGACY_KEYS.get(legacy_key_name)

        if legacy_key:
            legacy_value = config.get(legacy_key)

            if legacy_value not in (None, ""):
                return str(legacy_value)

    return ""


def get_document_type(config: dict) -> str:
    """
    Return the normalized document type.

    Supported values:
    - protocol
    - report
    """
    return get_config_value(
        config,
        ConfigKeys.DOC_TYPE,
        "DOC_TYPE",
    ).strip().lower()


def is_report(config: dict) -> bool:
    return get_document_type(config) == "report"


def get_doc_type_replacements(doc_type: str) -> Dict[str, str]:
    """
    Return the document-type-dependent placeholder values.
    """
    normalized_doc_type = (doc_type or "").strip().lower()
    return dict(_DOC_TYPE_REPLACEMENTS.get(normalized_doc_type, {}))


def build_placeholder_replacements(
    config: Optional[dict] = None,
) -> Dict[str, str]:
    """
    Build the complete placeholder replacement dictionary.

    This is the single source of truth used by:
    - placeholder_replacer.py
    - docx_verifier.py
    """
    if config is None:
        config = ConfigProvider.load_config_json()

    protocol_number = get_config_value(
        config,
        ConfigKeys.PROTOCOL_NUMBER,
        "DOC_STD",
    )

    report_number = get_config_value(
        config,
        ConfigKeys.REPORT_NUMBER,
        "REPORT_NUMBER",
    )

    std_name = get_config_value(
        config,
        ConfigKeys.STD_NAME,
        "STD_NAME",
    )

    test_plan = get_config_value(
        config,
        ConfigKeys.TEST_PLAN,
        "PLAN_NUMBER",
    )

    stx_number = get_config_value(
        config,
        ConfigKeys.STX_NUMBER,
        "STX_NUMBER",
    )

    prepared_by = get_config_value(
        config,
        ConfigKeys.PREPARED_BY,
        "PREPARED_BY",
    )

    footer = get_config_value(
        config,
        ConfigKeys.FOOTER,
        "FOOTER",
    )

    replacements = {
        WordPlaceholders.PROTOCOL_NUMBER: protocol_number,
        WordPlaceholders.STD_NAME: std_name,
        WordPlaceholders.PLAN_NUMBER: test_plan,

        # Keep the current visible behavior during this refactor.
        # We will deliberately change this to the final format afterward.
        WordPlaceholders.STX_NUMBER: stx_number,

        WordPlaceholders.PREPARED_BY: prepared_by,
        WordPlaceholders.FOOTER: footer,

        # Current template placeholder for the active document number.
        DOC_STD_NUMBER: (
            report_number if is_report(config) else protocol_number
        ),

        # Structural marker, not a user-facing placeholder.
        DELETE_ROW_MARKER: "",
    }

    replacements.update(
        get_doc_type_replacements(get_document_type(config))
    )

    return replacements