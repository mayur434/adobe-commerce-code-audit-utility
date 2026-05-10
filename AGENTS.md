---
description: Agent instructions for working with the Adobe Commerce Code Audit Tool
applyTo: "**"
---

# Adobe Commerce Audit Tool — Agent Context

## Project Overview

This is a **standalone, reusable** Python tool that performs enterprise-grade static code analysis on any Adobe Commerce / Magento 2 project. It scans 52 audit categories (42 code + 10 DB dump) and generates a comprehensive Excel report.

## Tech Stack

- **Language:** Python 3.8+
- **Dependencies:** `openpyxl` (Excel generation with charts)
- **No framework** — pure Python, no web server, CLI-only

## Project Structure

```
adobe-commerce-audit/
├── audit.py              # CLI entry point (argparse, --module filter, DB graceful handling)
├── lib/
│   ├── __init__.py       # Package version: 3.3.0
│   ├── scanner.py        # AdobeCommerceAuditScanner class (42 code + 10 DB scan methods)
│   ├── expert.py         # Expert validation engine (templates + category routing rules)
│   ├── report.py         # AuditReportGenerator class (Excel — pure data sheets, no merged cells)
│   └── styles.py         # Excel styling constants + helper functions
├── config.json           # Project config (categories, thresholds, modules filter, DB path)
├── output/               # Generated reports (gitignored)
├── SKILL.md              # Copilot skill definition
├── AGENTS.md             # This file — agent context
├── requirements.txt      # openpyxl>=3.1.0
├── .gitignore
└── README.md
```

## Key Classes

### `AdobeCommerceAuditScanner` (lib/scanner.py)
- Constructor: `__init__(project_root, namespace="Custom", db_dump_path=None, categories=None, thresholds=None, modules=None)`
- Entry point: `scan()` → returns `dict[str, list[dict]]` (category → findings)
- 42 private `_scan_*` code methods + 10 `_dbscan_*` DB dump methods, each auto-registered
- DB table-to-module mapping via `_table_to_module()` (100+ prefix rules for Magento core tables)
- Helper methods: `_grep()`, `_read()`, `_add()`, `_db_add()`, `_context()`, `_module()`, `_rel()`, `_filter_selected_modules()`
- Accumulates findings in `self.findings` (defaultdict) and counts in `self.stats` (Counter)

### `AuditReportGenerator` (lib/report.py)
- Constructor: `__init__(findings, stats, project_name, project_root)`
- Entry point: `generate(output_path)` → saves `.xlsx` file
- 5 sheet generators: `_sheet_executive_summary()`, `_sheet_detail()`, `_sheet_recommendations()`, `_sheet_module_rollout_summary()`, `_sheet_module_plan()`
- All data sheets are pure tabular (header row 1, data row 2+, no merged cells, auto-filter enabled)

### `get_expert_recommendation()` (lib/expert.py)
- Template-based expert validation engine with category routing rules
- Returns expert recommendation string for each finding based on category, issue type, severity, and effort
- Templates cover: service contracts, plugin ordering, cache profiling, ACL, observer patterns, MSI API, admin security, business flow integrity, critical callback security, and more

### `lib/styles.py` (constants + helpers)
- Style constants: `HEADER_FONT`, `HEADER_FILL`, `THIN_BORDER`, `ZEBRA_FILL_*`, etc.
- Helper functions: `severity_fill()`, `severity_font()`, `style_header_row()`, `apply_zebra_and_borders()`, `color_severity_col()`, `color_priority_col()`

## Coding Conventions

- **No external dependencies** beyond openpyxl — keep it portable
- Finding format: `{"module", "file", "line", "type", "description", "code", "severity", "recommendation", "effort"}`
- Severity levels: `CRITICAL`, `HIGH`, `MEDIUM`, `LOW`, `INFO`
- Effort levels: `Low`, `Medium`, `High`, `Very High`
- All file paths stored as relative to project root via `_rel()`
- Scanner methods use `_grep()` for line-level regex and `_read()` for full-file content with caching

## How to Run

```bash
python3 audit.py --path /path/to/magento2-project --name "Project Name"
```

## How to Add a New Scanner

1. Add a method `_scan_new_category(self, php, xml, phtml)` to `AdobeCommerceAuditScanner`
2. Register it in the `scanners` list inside `scan()` as `("New Category", self._scan_new_category)`
3. Use `self._add(category, module, filepath, line, type, description, code, severity, recommendation, effort)` to record findings
4. The category will automatically appear in the report — no changes needed in `report.py`

## Testing

```bash
# Run against any Adobe Commerce project using config.json
python3 audit.py

# Run with CLI overrides
python3 audit.py --path /path/to/project --name "Project Name"

# Targeted re-run for specific modules
python3 audit.py --module Vendor_Checkout,Vendor_Payment
```

## Important Notes

- This tool does **static analysis only** — it reads files, never executes project code
- The scanner caches file reads in `self._php_cache` for performance
- DB dump analysis uses a streaming parser for multi-GB SQL files
- DB findings are mapped to owning Magento modules via table name prefix matching (100+ rules)
- The Recommendations sheet contains ~60 hardcoded best-practice items (not generated from findings)
- The Module Execution Plan lists every finding grouped by module, sorted by risk score
- The Module Rollout Summary provides wave-based deployment planning per module/domain
