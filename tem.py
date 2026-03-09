# %% [markdown]
# ===============================================================
# CNE Sales Report Generator
# ===============================================================
# This notebook processes:
#   • Financial Transactions (FT)
#   • New Sales
#
# It enriches them with reference data from BeinData and produces
# two Excel reports:
#   • Dealers
#   • Direct Sales
# ===============================================================


# %%
# ===============================================================
# Imports
# ===============================================================
import calendar
import shutil
import warnings

from datetime import datetime, timedelta

import pandas as pd

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from functions import create_folder_if_not_exists, search_files, sendMail


# Silence pandas FutureWarnings etc.
warnings.filterwarnings("ignore")


# %%
# ===============================================================
# Configuration
# ===============================================================

# ---- Date helpers ------------------------------------------------
today = datetime.now().date()
base_dir_format = datetime.now().strftime("%d.%m.%y")
base_dir = f"s:\\{base_dir_format}"

# ---- Output directory --------------------------------------------
result_directory = "s:\\playground"
create_folder_if_not_exists(result_directory)

# ---- Pandas display settings -------------------------------------
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", 50)
pd.set_option("display.float_format", "{:.2f}".format)


# %%
# ===============================================================
# Reporting Period Logic
# ===============================================================

if today.day == 1:
    # First day of month → full previous month
    last_month = today.month - 1 if today.month > 1 else 12
    last_month_year = today.year if today.month > 1 else today.year - 1

    FROM_DATE = pd.Timestamp(last_month_year, last_month, 1)
    TO_DATE = pd.Timestamp(
        last_month_year,
        last_month,
        calendar.monthrange(last_month_year, last_month)[1],
    )

elif today.weekday() == 6:
    # Sunday → include Friday + Saturday
    FROM_DATE = pd.Timestamp(today - timedelta(days=3))
    TO_DATE = pd.Timestamp(today - timedelta(days=1))

else:
    # Default → yesterday only
    FROM_DATE = pd.Timestamp(today - timedelta(days=1))
    TO_DATE = pd.Timestamp(today - timedelta(days=1))


# Manual override (disabled)
# FROM_DATE = "2026-01-01"
# TO_DATE   = "2026-12-26"


# %%
# ===============================================================
# Output Files
# ===============================================================

OUTPUT_FILE_DEALERS = (
    f"Dealers report {FROM_DATE:%Y-%m-%d} to {TO_DATE:%Y-%m-%d}.xlsx"
)

OUTPUT_FILE_DS = (
    f"Direct Sales report {FROM_DATE:%Y-%m-%d} to {TO_DATE:%Y-%m-%d}.xlsx"
)


# %%
# ===============================================================
# Locate Source Files
# ===============================================================

FT_PARTIAL_NAME = "NEWCNEFINTRANSRPT"
NEWSALES_PARTIAL_NAME = "CNENEWCAPTURERPT.CSV"
BEINDATA_PARTIAL_NAME = "BEINDATANEWRPT"

ft_found_files = search_files(base_dir, FT_PARTIAL_NAME)
newsales_found_files = search_files(base_dir, NEWSALES_PARTIAL_NAME)
beindata_found_files = search_files(base_dir, BEINDATA_PARTIAL_NAME)


# %%
# ===============================================================
# Validate Required Files
# ===============================================================

missing_files = []

if not ft_found_files:
    missing_files.append(f"FT file ({FT_PARTIAL_NAME})")

if not newsales_found_files:
    missing_files.append(f"New-sales file ({NEWSALES_PARTIAL_NAME})")

if len(beindata_found_files) < 2:
    missing_files.append(
        f"BeinData file ({BEINDATA_PARTIAL_NAME}) — "
        f"found {len(beindata_found_files)}, need at least 2"
    )


if missing_files:

    error_msg = (
        f"Missing source files in {base_dir}:\n"
        + "\n".join(f"  - {f}" for f in missing_files)
    )

    print(f"ERROR: {error_msg}")

    sendMail(
        subject="CNE Report - Missing Source Files",
        body=(
            "<p>The CNE Sales Report could not be generated."
            "<br><br><b>Missing files:</b></p><ul>"
            + "".join(f"<li>{f}</li>" for f in missing_files)
            + f"</ul><p>Search directory: <code>{base_dir}</code></p>"
        ),
    )

    raise SystemExit(error_msg)


# %%
# ===============================================================
# Log Discovered Files
# ===============================================================

print(f"FT file:        {ft_found_files[0]}")
print(f"New-sales file: {newsales_found_files[0]}")
print(f"BeinData file:  {beindata_found_files[1]}")
print(f"Current date:   {today}")