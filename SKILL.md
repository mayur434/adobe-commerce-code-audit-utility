---
description: Run an enterprise Adobe Commerce / Magento 2 static code audit that scans 27 categories and generates an Excel report with charts, recommendations, and action plan.
---

# Adobe Commerce Code Audit Skill

## When to Use
- User asks to audit, scan, or review an Adobe Commerce / Magento 2 codebase
- User wants a code quality, security, or performance report for a Magento project
- User asks to find issues, vulnerabilities, or technical debt in custom modules
- User wants recommendations or an action plan for a Magento project

## How to Use

```bash
cd /path/to/adobe-commerce-audit
python3 audit.py --path <PROJECT_ROOT> --name "<Project Name>"
```

### CLI Options

| Option         | Default    | Description                             |
|---------------|------------|-----------------------------------------|
| `--path`      | (required) | Path to Adobe Commerce project root     |
| `--name`      | dir name   | Project name for the report title       |
| `--output`    | `output/`  | Output directory for the Excel report   |
| `--namespace` | `Custom`   | Custom module namespace to scan         |

## What It Scans (27 Categories)

1. Exception Handling — empty catch, generic catch, debug output, missing finally
2. Security — hardcoded credentials, SQL injection, file upload, superglobals, path traversal, CSP
3. Database — N+1 queries, save in loop, raw SQL, missing transactions, missing indexes
4. Caching — hardcoded TTL, GraphQL resolver cache, external API cache
5. Code Structure — god classes, fat constructors, business logic in controllers, missing strict_types
6. Performance — unbounded collections, sleep(), regex in loop, file_get_contents
7. Deprecated — setup_version, array() syntax, deprecated install scripts, Zend v1
8. Logging — no rotation, verbose logging, sensitive data in logs, static logger
9. File Storage — file writes, S3 ops, CSV generation, directory creation
10. Reusability — duplicate classes, duplicate config patterns
11. Test Coverage — modules without any tests
12. Dependency Injection — ObjectManager::getInstance(), legacy _objectManager
13. Plugin Architecture — around plugins (expensive)
14. Cron Jobs — every-minute crons, crons without lock
15. GraphQL — N+1 in resolvers, complex resolvers
16. Queue Processing — consumers without error handling, no max messages
17. Configuration — hardcoded URLs, magic numbers
18. Frontend Templates — heavy PHP in PHTML, inline JS, ObjectManager in templates
19. XML Configuration — plugin sortOrder, core preferences, sensitive fields, module sequence
20. WebAPI & ACL — anonymous endpoints, rate limiting, ACL granularity
21. DB Schema — missing indexes, varchar length, nullable defaults, wide tables
22. Infrastructure — php.ini, nginx security headers, docker health checks
23. Cloud Deployment — .magento.app.yaml, .magento.env.yaml, services.yaml
24. PHP Deep Analysis — weak hashing, exit/die, error suppression, DateTime timezone, TODO markers
25. Event Observers — hot event observers, heavy observers, global dispatch
26. Module Architecture — missing module.xml, missing API contracts, controller HTTP interfaces
27. Code Metrics — file size, method count, module size

## Output

Generates a timestamped `.xlsx` Excel report with ~31 sheets:
- **Executive Summary** — severity breakdown, category breakdown, top risk modules
- **27 category detail sheets** — per-finding: module, file, line, code context, severity, recommendation, effort
- **Recommendations** — 60+ actionable items grouped by area with priority
- **Action Plan** — P0–P4 prioritized items mapped to sprints
- **Charts** — severity pie, top modules bar, category stacked bar

## Prerequisites

- Python 3.8+
- `openpyxl>=3.1.0` (`pip install -r requirements.txt`)

## Architecture

```
lib/
├── scanner.py   # AdobeCommerceAuditScanner — 27 scan methods, file helpers, finding accumulator
├── report.py    # AuditReportGenerator — Excel workbook with 5 sheet generators + chart engine
├── styles.py    # All Excel styles, colors, fonts, borders, formatting helper functions
└── __init__.py  # Package version (3.0.0)
```

## Example Invocations

```bash
# Scan a project with defaults
python3 audit.py --path /home/user/magento2

# Scan with custom name and output
python3 audit.py --path /srv/www/magento --name "Acme Store" --output /tmp/reports

# Scan with custom namespace
python3 audit.py --path ./project --namespace AcmeVendor
```
