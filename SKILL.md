---
description: Run an enterprise Adobe Commerce / Magento 2 static code and DB audit that scans 52 categories and generates an Excel report with expert validation, recommendations, module rollout summary, and module execution plan.
---

# Adobe Commerce Code Audit Skill

## When to Use
- User asks to audit, scan, or review an Adobe Commerce / Magento 2 codebase
- User wants a code quality, security, or performance report for a Magento project
- User asks to find issues, vulnerabilities, or technical debt in custom modules
- User wants recommendations or a module-level execution plan for a Magento project
- User wants to audit a production database dump for schema, index, integrity, or performance issues

## How to Use

```bash
cd /path/to/adobe-commerce-audit
python3 audit.py --path <PROJECT_ROOT> --name "<Project Name>"
```

### CLI Options

| Option         | Default      | Description                                          |
|---------------|--------------|------------------------------------------------------|
| `--path`      | config.json  | Path to Adobe Commerce project root                  |
| `--db`        | config.json  | SQL dump path for DB audit                           |
| `--name`      | dir name     | Project name for the report title                    |
| `--output`    | `output/`    | Output directory for the Excel report                |
| `--namespace` | `Custom`     | Custom module namespace to scan                      |
| `--config`    | `config.json`| Path to config JSON                                  |
| `--module`    | (all)        | Optional targeted filter for re-runs only            |

## What It Scans (52 Categories)

### Code Audit (42 scanners)
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
28. Business Logic Identification — payment, checkout, order, customer, catalog business patterns
29. Business Customization Review — direct entity state, direct save, synchronous external API, hardcoded rules
30. Critical Commerce Flows — webhook/callback, around plugins on critical paths, collectTotals
31. MSI Inventory & Source Management — legacy stock writes, direct inventory tables, salable qty assumptions
32. Admin & Integration Security — admin ACL, WebAPI resources, webhook signatures, replay protection
33. Logical Flow & Cross-Module — circular deps, coupling, duplication, event flows, plugin chains, orphans
34–42. Coding Standards, Input Validation & XSS, Frontend Assets, Composer & Dependencies, Full Page Cache & Private Content, Backward Compatibility, Configuration & Scope, Layout & UI Components, XML Schema Validation

### DB Dump Audit (10 scanners)
43. DB: Table Structure — wide tables, missing PKs, composite PKs, high row counts
44. DB: Index Analysis — redundant/duplicate indexes, missing indexes, cardinality
45. DB: Column Analysis — column types, nullable defaults, varchar sizing, LOB patterns
46. DB: Foreign Keys — missing FKs, orphan references, cascade rules
47. DB: Naming Conventions — table/column naming patterns, reserved words
48. DB: Storage Engine — non-InnoDB tables
49. DB: Charset & Collation — mixed charsets, non-utf8mb4 tables
50. DB: Adobe Commerce Schema — Magento core schema deviations
51. DB: Data Integrity — SKU uniqueness, email uniqueness, orphan data patterns
52. DB: Performance — large tables, no clustered index, hot tables

## Output

Generates a timestamped `.xlsx` Excel report with ~47 sheets:
- **Executive Summary** — severity breakdown, category breakdown, top risk modules
- **44 category detail sheets** — per-finding: module, file, line, code context, severity, recommendation, expert validation, effort (pure data — filterable, sortable, groupable)
- **Recommendations** — 60+ actionable items grouped by area with priority
- **Module Rollout Summary** — wave-based deployment planning per module/domain
- **Module Execution Plan** — all findings grouped by module, sorted by risk score (pure data)

## Prerequisites

- Python 3.8+
- `openpyxl>=3.1.0` (`pip install -r requirements.txt`)

## Architecture

```
lib/
├── scanner.py   # AdobeCommerceAuditScanner — 42 code + 10 DB scan methods, DB table-to-module mapping
├── expert.py    # Expert validation engine — templates + category routing rules
├── report.py    # AuditReportGenerator — Excel workbook with 5 sheet generators (pure data, no merged cells)
├── styles.py    # All Excel styles, colors, fonts, borders, formatting helper functions
└── __init__.py  # Package version (3.3.0)
```

## Example Invocations

```bash
# Scan using config.json (recommended)
python3 audit.py

# Scan a project with CLI overrides
python3 audit.py --path /home/user/magento2 --name "Acme Store"

# Scan with DB dump
python3 audit.py --path /srv/www/magento --db /path/to/dump.sql --name "Acme Store"

# Scan with custom output
python3 audit.py --path ./project --output /tmp/reports

# Targeted re-run for specific modules after full audit
python3 audit.py --module Vendor_Checkout,Vendor_Payment
```
