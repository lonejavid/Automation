"""
Automation runner script used by the web interface and CLI.

Supports downloading a single store or all stores for a given date.
"""

from pathlib import Path
import sys
import argparse
from datetime import datetime, timedelta
from typing import Optional

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.append(str(SRC_DIR))

from automation.hmecloud import (
    STORES,
    download_all_stores,
    download_single_store,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run HME Cloud automation for a single store or all stores."
    )
    parser.add_argument(
        "--mode",
        choices=["single", "all"],
        default="all",
        help="Run automation for a single store or all stores (default: all).",
    )
    parser.add_argument(
        "--store",
        help="Store name or index (1-based) when running in single mode.",
    )
    parser.add_argument(
        "--date",
        help="Report date in YYYY-MM-DD format (defaults to yesterday).",
    )
    return parser.parse_args()


def resolve_report_date(date_str: Optional[str]) -> datetime:
    if not date_str:
        return datetime.now() - timedelta(days=1)
    try:
        return datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        print("⚠️  Invalid date format. Using yesterday instead.")
        return datetime.now() - timedelta(days=1)


def resolve_store(store_arg: Optional[str]) -> str:
    if not store_arg:
        return STORES[0]

    # Allow numeric index (1-based)
    try:
        index = int(store_arg) - 1
        if 0 <= index < len(STORES):
            return STORES[index]
    except ValueError:
        pass

    # Fallback to exact string match
    if store_arg in STORES:
        return store_arg

    print("⚠️  Store not recognized. Defaulting to first store.")
    return STORES[0]


def main():
    args = parse_args()
    report_date = resolve_report_date(args.date)

    print("\n" + "=" * 80)
    print("🤖 HME CLOUD AUTOMATION")
    print("=" * 80)
    print(f"Report date: {report_date.strftime('%B %d, %Y')}")

    if args.mode == "single":
        store_name = resolve_store(args.store)
        print(f"Mode: SINGLE STORE\nStore: {store_name}")
        print("=" * 80)
        success = download_single_store(store_name, report_date=report_date)
    else:
        print("Mode: ALL STORES")
        print("=" * 80)
        success = download_all_stores(report_date=report_date)

    print("\n" + "=" * 80)
    if success:
        print("✅ AUTOMATION COMPLETED SUCCESSFULLY!")
    else:
        print("❌ AUTOMATION FINISHED WITH ERRORS")
    print("=" * 80)


if __name__ == "__main__":
    main()
