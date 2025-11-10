"""Automation helpers for the HME Cloud demos."""

from .hmecloud import (
    setup_chrome_driver,
    login_to_hmecloud,
    navigate_to_reports,
    select_store_and_date,
    download_store_report,
    download_all_stores,
)
from .run_macro import process_downloaded_file

__all__ = [
    "setup_chrome_driver",
    "login_to_hmecloud",
    "navigate_to_reports",
    "select_store_and_date",
    "download_store_report",
    "download_all_stores",
    "process_downloaded_file",
]
