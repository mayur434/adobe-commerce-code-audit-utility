# Adobe Commerce Code Audit Tool v3.0

Enterprise-grade static code analysis for **Adobe Commerce / Magento 2** projects.

Scans **27 audit categories** across PHP, XML, PHTML, and infrastructure config files, then generates a comprehensive Excel report with charts, recommendations, and a prioritized action plan.

---

## Features

- **27 scanners**: Exception Handling, Security, Database, Caching, Performance, Deprecated Code, Logging, File Storage, DI, Plugins, Cron, GraphQL, Queues, Config, Frontend, XML Config, WebAPI & ACL, DB Schema, Infrastructure, Cloud Deployment, PHP Deep Analysis, Event Observers, Module Architecture, Code Metrics, and more.
- **Excel report** with:
  - Executive Summary (severity & category breakdown, top risk modules)
  - Per-category detail sheets with color-coded severity, code context, and recommendations
  - Recommendations sheet (60+ actionable items grouped by area)
  - Prioritized Action Plan (P0–P4 with sprint mapping)
  - Charts sheet (pie, bar, stacked bar)
- **Zero dependencies on the scanned project** — pure Python, only needs `openpyxl`.

## Requirements

- Python 3.8+
- `openpyxl` (install via `pip`)

## Installation

```bash
git clone <repo-url> adobe-commerce-audit
cd adobe-commerce-audit
pip install -r requirements.txt
```

## Usage

```bash
# Basic — scans project root, outputs to output/ directory
python3 audit.py --path /path/to/magento2-project

# Custom project name
python3 audit.py --path /path/to/project --name "My Client"

# Custom output directory
python3 audit.py --path /path/to/project --output ./reports

# Custom namespace (default: Custom)
python3 audit.py --path /path/to/project --namespace VijaySales
```

### CLI Options

| Option         | Default     | Description                                  |
|---------------|-------------|----------------------------------------------|
| `--path`      | (required)  | Path to Adobe Commerce project root          |
| `--name`      | dir name    | Project name for the report title            |
| `--output`    | `output/`   | Output directory for the Excel report        |
| `--namespace` | `Custom`    | Custom module namespace to scan              |

## Project Structure

```
adobe-commerce-audit/
├── audit.py              # CLI entry point
├── lib/
│   ├── __init__.py       # Package metadata (version)
│   ├── scanner.py        # AdobeCommerceAuditScanner — 27 scan methods
│   ├── report.py         # AuditReportGenerator — Excel report + charts
│   └── styles.py         # Excel styles, colors, formatting helpers
├── output/               # Generated reports (gitignored)
├── requirements.txt
├── .gitignore
└── README.md
```

## Output

The tool generates a timestamped `.xlsx` file (e.g., `MyProject-code-audit-20250101_120000.xlsx`) with ~30 sheets:

1. **Executive Summary** — severity breakdown, category breakdown, top risk modules
2. **27 category sheets** — detailed findings with module, file, line, code context, severity, recommendation
3. **Recommendations** — 60+ actionable items with area color-coding, effort, priority
4. **Action Plan** — P0–P4 prioritized items mapped to sprints
5. **Charts** — severity pie, top modules bar, category stacked bar

## License

MIT
# adobe-commerce-audit
