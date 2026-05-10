#!/usr/bin/env python3
"""
Adobe Commerce Code Audit Tool v3.0
=====================================
Enterprise-grade static code analysis for Adobe Commerce (Magento 2) projects.
Scans all configured categories and generates an Excel report with charts and module-wise remediation planning.
Supports optional DB dump analysis for full database audit.

Usage:
    python3 audit.py                            # uses config.json
    python3 audit.py --config my-config.json    # uses custom config
    python3 audit.py --path /override/path      # code audit only
    python3 audit.py --db /path/to/dump.sql     # DB audit only
    python3 audit.py --path /code --db /dump.sql # both code + DB audit
"""

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime

from lib.scanner import AdobeCommerceAuditScanner
from lib.report import AuditReportGenerator

DEFAULT_CONFIG = "config.json"


def load_config(config_path):
    """Load and return config dict from JSON file."""
    if not os.path.isfile(config_path):
        return {}
    try:
        with open(config_path, 'r') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        print(f"⚠️  Warning: Could not parse {config_path}: {e}")
        return {}


def main():
    parser = argparse.ArgumentParser(
        description="Adobe Commerce Enterprise Code Audit Tool v3.0",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
All settings can be configured in config.json. CLI arguments override config values.
At least one of --path or --db must be provided (or set in config.json).

Examples:
  python3 audit.py                                    # uses config.json
  python3 audit.py --config project-a.json            # uses custom config
  python3 audit.py --path /path/to/project            # code audit only
  python3 audit.py --db /path/to/dump.sql             # DB audit only
  python3 audit.py --path /path --db /path/dump.sql   # both code + DB audit
  python3 audit.py --path /path --name "Client" --output ./reports
        """,
    )
    parser.add_argument("--config", default=DEFAULT_CONFIG, help="Path to config JSON file (default: config.json)")
    parser.add_argument("--path", default=None, help="Path to Adobe Commerce project root (overrides config)")
    parser.add_argument("--db", default=None, help="Path to SQL dump file for database analysis (overrides config)")
    parser.add_argument("--name", default=None, help="Project name (overrides config)")
    parser.add_argument("--output", default=None, help="Output directory (overrides config)")
    parser.add_argument("--namespace", default=None, help="Custom module namespace (overrides config)")
    parser.add_argument("--module", default=None, help="Optional targeted re-run filter. Leave empty for the main full-project audit.")
    args = parser.parse_args()

    # Load config file
    cfg = load_config(args.config)
    project_cfg = cfg.get("project", {})
    output_cfg = cfg.get("output", {})
    scanner_cfg = cfg.get("scanner", {})
    thresholds = cfg.get("thresholds", {})
    db_cfg = cfg.get("database", {})

    # Resolve values: CLI > config > defaults
    project_path = args.path or project_cfg.get("path")
    db_path = args.db or db_cfg.get("dump_path")

    if not project_path and not db_path:
        print("❌ Error: No project path or DB dump provided.")
        print("   Set in config.json, or use --path and/or --db")
        sys.exit(1)

    # Validate project path if provided
    if project_path:
        project_path = os.path.abspath(project_path)
        if not os.path.isdir(project_path):
            print(f"❌ Error: Project path does not exist: {project_path}")
            sys.exit(1)
        if not os.path.isdir(os.path.join(project_path, "app", "code")):
            print(f"⚠️  Warning: No app/code directory found. This may not be an Adobe Commerce project.")

    # Validate DB dump path if provided
    if db_path:
        db_path = os.path.abspath(db_path)
        if not os.path.isfile(db_path):
            if args.db:
                print(f"❌ Error: DB dump file does not exist: {db_path}")
                sys.exit(1)
            print(f"⚠️  Warning: Configured DB dump does not exist, DB audit will be skipped: {db_path}")
            db_path = None

    project_name = args.name or project_cfg.get("name") or (os.path.basename(project_path) if project_path else "DB-Audit")
    output_dir = os.path.abspath(args.output or output_cfg.get("directory", "output"))
    namespace = args.namespace or scanner_cfg.get("namespace", "Custom")
    categories = scanner_cfg.get("categories", None)  # None = all categories
    module_filter = args.module or scanner_cfg.get("modules")
    if isinstance(module_filter, str):
        modules = [m.strip() for m in module_filter.split(",") if m.strip()]
    else:
        modules = list(module_filter or [])

    os.makedirs(output_dir, exist_ok=True)

    # Detect git branch from the scanned project
    branch = ""
    if project_path:
        try:
            result = subprocess.run(
                ["git", "rev-parse", "--abbrev-ref", "HEAD"],
                cwd=project_path, capture_output=True, text=True, timeout=5,
            )
            if result.returncode == 0:
                branch = result.stdout.strip().replace("/", "-")
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

    # Determine audit mode
    mode_parts = []
    if project_path:
        mode_parts.append("code")
    if db_path:
        mode_parts.append("db")
    audit_mode = "+".join(mode_parts)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    branch_part = f"-branch-{branch}" if branch else ""
    output_file = os.path.join(output_dir, f"{project_name}-audit-{audit_mode}-{timestamp}{branch_part}.xlsx")

    print(f"📄 Config: {args.config}")
    print(f"   Project: {project_name}")
    print(f"   Audit Mode: {audit_mode.upper()}")
    if project_path:
        print(f"   Code Path: {project_path}")
    if db_path:
        print(f"   DB Dump: {db_path}")
    print(f"   Output: {output_dir}")
    if project_path:
        print(f"   Namespace: {namespace}")
    if branch:
        print(f"   Git Branch: {branch}")
    if categories:
        print(f"   Categories: {len(categories)} selected")
    if modules:
        print(f"   Modules: {', '.join(modules)}")
        print("   ⚠️  Partial scan mode: use this only for targeted re-runs. Main audit should leave modules empty to cover all modules.")
    if thresholds:
        print(f"   Thresholds: {len(thresholds)} custom values")
    print()

    # Scan
    scanner = AdobeCommerceAuditScanner(
        project_root=project_path,
        namespace=namespace,
        thresholds=thresholds,
        categories=categories,
        db_dump_path=db_path,
        modules=modules,
    )
    findings = scanner.scan()

    # Report
    report = AuditReportGenerator(findings, scanner.stats, project_name,
                                   project_path or db_path)
    report.generate(output_file)

    print(f"\n📁 Report saved to: {output_file}")


if __name__ == "__main__":
    main()
