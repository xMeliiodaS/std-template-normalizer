# STD Template Normalizer

A Python automation utility that turns an exported STD dataset (Excel or Word) plus a Word protocol/report template into a production-ready normalized `.docx`.

It was built to support a larger C# desktop workflow (STE Tool Studio) where users provide input files and metadata, then expect a consistently formatted output document with:

- table content copied from the exported STD source,
- strict template-preserving behavior,
- placeholder replacement driven by config,
- layout/format normalization for downstream review and release.

---

## Why this tool exists

In practice, teams had repetitive and error-prone manual work:

1. open a template,
2. copy STD rows into the right table,
3. reformat orientation/tables/spacing,
4. fill placeholders (`ADD_*` tokens),
5. validate no content was lost.

This project automates the full pipeline and adds **verification-grade checks** so that the result can be trusted in CI/manual validation.

---

## Core concept

The normalizer separates the process into two concerns:

1. **Transformation** (write path):
   - read source content,
   - mutate target `.docx`,
   - apply formatting and placeholder rules.

2. **Verification** (read-only path):
   - inspect source vs output,
   - enforce structure/content/format expectations,
   - detect unresolved placeholders or accidental template drift.

This dual design is important for FRS documentation because it defines both **functional behavior** and **acceptance criteria**.

---

## End-to-end workflow

The integration flow in `test/test_document_normalization.py` is effectively the reference pipeline:

1. Resolve input paths from config.
2. Set all sections to landscape.
3. Set all tables to AutoFit to Window.
4. Copy source rows into target table (Word or Excel source).
5. Apply explicit target table column widths.
6. Normalize paragraph spacing in the Section-6 table.
7. Replace placeholders across body/tables/header/footer.
8. Run comprehensive verification (`verify_normalized_protocol`).

---

## Inputs and outputs

### Inputs

- **Exported STD**: either `.xlsx` or `.docx`.
- **Template Protocol**: `.docx` with a target table (header-based discovery, default header prefix: `ID`).
- **Config JSON**: metadata values for placeholders and path settings.

### Output

- **Normalized Protocol**: finalized `.docx` containing source data + normalized formatting + resolved placeholders.

---

## Module map

### `src/word/table_handler.py`
Responsibilities:

- enforce landscape orientation,
- set table autofit behavior,
- copy source rows into target table using header detection,
- set second-column style to `Normal` and remove numbering metadata,
- enforce fixed target column widths,
- adjust paragraph spacing for the table after the Section 6 heading.

### `src/word/placeholder_replacer.py`
Responsibilities:

- replace placeholders in paragraphs/tables/header/footer,
- support placeholders split across Word runs,
- apply `doc_type` overrides (`protocol` vs `report`) for:
  - `ADD_DOC_TYPE`,
  - `ADD_DOC_RECORD`,
  - `ADD_DOC_STX`,
- remove rows marked `TO_BE_DELETED_ROW` for non-report mode.

### `src/excel/xlsx_reader.py`
Responsibilities:

- lightweight `.xlsx` parsing (ZIP + XML),
- resolve worksheet and shared strings,
- return a row matrix while preserving sparse column positions.

### `src/validation/docx_verifier.py`
Responsibilities:

- validate content integrity (source rows vs output target table),
- validate structural correctness,
- validate formatting rules,
- validate placeholder replacement completeness,
- validate template preservation outside the target table,
- provide strict/non-strict verification reporting.

### `src/config/*`
Responsibilities:

- constants for config keys/placeholders/default widths,
- config loading from `%APPDATA%\ste_tool_studio\config.json` by default,
- shared logging to `%APPDATA%\ste_tool_studio\ste_tool_studio.log`.

---

## Placeholder system

Primary placeholders are centralized in `WordPlaceholders`:

- `ADD_DOC_TYPE`
- `ADD_DOC_STX`
- `ADD_DOC_RECORD`
- `ADD_PROTOCOL_NUMBER#`
- `ADD_REPORT_NUMBER`
- `ADD_STD_NAME`
- `ADD_PLAN_NUMBER`
- `ADD_STX_NUMBER`
- `ADD_PREPARED_BY`
- `ADD_FOOTER`

Special behavior:

- `doc_type=protocol` maps to Design / Protocol / (STD).
- `doc_type=report` maps to Report / Report / (STR).
- legacy config keys are supported for backward compatibility.

---

## Configuration model

`ConfigKeys` defines both modern and legacy key names. Typical fields:

- `Exported_STD`
- `Template_protocol`
- `Normalized_protocol`
- `doc_type`, `doc_stx`, `doc_record`
- `protocol_number`, `report_number`, `std_name`, `test_plan`, `stx_number`, `prepared_by`, `footer`

A sample config exists at repository root (`config.json`) for local setup reference.

---

## Verification contract (useful for FRS)

`verify_normalized_protocol(...)` runs six check groups:

1. `content_integrity` (critical)
2. `structural_correctness` (critical)
3. `formatting` (warning in non-strict mode)
4. `placeholder_replacement` (critical)
5. `template_preservation` (warning in non-strict mode)
6. `body_paragraphs_preserved` (warning in non-strict mode)

In strict mode, any failure raises `VerificationError`.

---

## How to run

### Install dependencies

```bash
pip install -r requirements.txt
```

### Run tests / normalization flow

```bash
python -m unittest -v
```

> Note: the integration test depends on real files referenced in config. If they are missing, that test is skipped.

---

## Known assumptions and boundaries

- Target table detection assumes expected header prefix (default `ID`).
- Merged-cell tables are explicitly rejected by verifier for content integrity logic.
- Section-6 spacing normalization depends on detecting a heading like `6 ...` or `section 6`.
- Placeholder replacement assumes configured token vocabulary (`ADD_*`).
- Runtime paths and logging are Windows-oriented because this tool is integrated with a Windows C# app.

---

## FRS authoring guidance

If your teammate is writing an FRS from this repo, structure requirements around these categories:

1. **Input handling** (Word/Excel source, config keys, path resolution).
2. **Transformation requirements** (copy semantics, formatting normalization, placeholder mapping).
3. **Output requirements** (document structure and expected table shape).
4. **Verification requirements** (critical vs non-critical checks, strict mode behavior).
5. **Compatibility requirements** (legacy config key support, doc_type overrides).
6. **Operational requirements** (logging location, failure diagnostics, skip behavior when test fixtures missing).

This README is intended to be the high-level product understanding document for that FRS effort.
