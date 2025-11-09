"""Automation package for HME Cloud workflows."""

from .hmecloud import (
    setup_chrome_driver,
    login_to_hmecloud,
    navigate_to_reports,
    select_store_and_date,
    download_single_store,
    download_all_stores,
)
from .complete_automation import main as run_complete_automation
