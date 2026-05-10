# Adobe Commerce Code Audit Tool v3.2

Enterprise-grade static code analysis for **Adobe Commerce / Magento 2** projects.

The utility scans PHP, XML, PHTML, frontend assets, Cloud/infrastructure config, and optional SQL dumps, then generates a detailed Excel report with severity, remediation, expert validation, action planning, charts, and module-level rollout planning.

---

## Key capabilities

- **40+ code and DB audit areas**, including security, coding standards, exception handling, DI, plugins, observers, cron, queues, GraphQL, WebAPI/ACL, DB schema, database dump analysis, Cloud deployment, caching, frontend templates, UI/layout XML, Composer/dependencies, backward compatibility, and infrastructure.
- **Business customization review** for high-blast-radius flows such as checkout, quote, order, payment, refund, shipping, inventory, customer, promotions, and external integrations.
- **Critical commerce flow checks** for risky around plugins, `collectTotals()` use, webhook/idempotency gaps, direct entity state mutation, synchronous external calls, and notifications coupled to transactions.
- **MSI / inventory review** for legacy stock writes, direct inventory table access, salable quantity assumptions, reservations, multi-source behavior, backorders, cancellations, refunds, and shipment source deduction.
- **Admin and integration security review** for admin ACL, WebAPI resources, anonymous/customer API exposure, webhook signature validation, replay protection, and negative test expectations.
- **Full-project scanning by default** across all custom modules, with module-wise remediation and production rollout planning in the generated report.
- Optional `--module` / `scanner.modules` filter is retained only for targeted re-runs, debugging, or validating a specific remediation batch after the full audit is complete.
- **Expert validation column** in every detail sheet to validate whether the base recommendation is aligned, needs stronger controls, may be a false positive, or needs rollout caution.

## Requirements

- Python 3.8+
- `openpyxl`

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Scan full codebase using config.json
python3 audit.py

# Scan code only
python3 audit.py --path /path/to/magento2-project

# Optional targeted re-run after the full audit, only for validating a specific fix batch
python3 audit.py --path /path/to/magento2-project --module Vendor_Checkout,Vendor_Payment

# Scan code and DB dump
python3 audit.py --path /path/to/project --db /path/to/prod-dump.sql

# Custom output and name
python3 audit.py --path /path/to/project --name "Client Name" --output ./reports

# Use a custom config
python3 audit.py --config project-audit.json
```

## CLI options

| Option | Description |
|---|---|
| `--config` | Path to config JSON. Defaults to `config.json`. |
| `--path` | Adobe Commerce project root. Overrides `project.path`. |
| `--db` | SQL dump path. Overrides `database.dump_path`. Invalid CLI-provided DB path is treated as an error. Invalid config-only DB path is skipped with a warning. |
| `--name` | Project/report name. |
| `--output` | Output directory. |
| `--namespace` | Custom namespace hint. |
| `--module` | Optional targeted filter for re-runs only. Leave empty for the main audit so all modules are covered. |

## Module-by-module remediation planning

Run the main audit against the **entire Adobe Commerce project**. Do not split the audit by module, because cross-module plugins, observers, preferences, layout updates, shared services, WebAPI routes, queues, cron jobs, and database changes can create hidden dependencies.

Use the generated **Module Rollout Summary** and **Module Execution Plan** sheets to plan fixes and deployments module by module. Optional `--module` runs should be used only after the full audit, for targeted re-validation of a module or release batch.

Recommended rollout order:

1. Security, admin/API ACL, webhooks, payment callbacks, and secrets.
2. Checkout, quote, payment, order, invoice, shipment, credit memo, and refund flows.
3. Inventory/MSI, source selection, salable qty, reservations, and ERP integrations.
4. DB schema/indexes, performance, caching/FPC/private content, queues, and cron.
5. Code structure, coding standards, tests, frontend assets, and documentation.

## Output

The generated Excel workbook contains:

- Executive Summary
- Per-category finding sheets
- Expert Validation & Recommendation column on every detail sheet
- Recommendations
- Action Plan
- Module Rollout Summary
- Module Execution Plan
- Charts

## Notes

Static analysis can identify risky patterns and probable edge cases, but critical business findings must be validated against runtime behavior, store configuration, third-party modules, integrations, and production-like data. Treat generated recommendations as a review accelerator, not as a replacement for architecture review and regression testing.
