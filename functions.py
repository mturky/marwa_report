"""
functions.py – Shared utility functions for the CNE Sales Report pipeline.

Contains helpers for:
  • File system operations (search, list, folder creation, compression)
  • Logging
  • Email sending (via Exchange Web Services)
  • Data-formatting helpers used by other report scripts
"""

import os
import gzip
import shutil
import tarfile
import datetime as dtm
from datetime import datetime
from typing import List, Optional,Union

import pandas as pd
from exchangelib import (
    Account,
    Configuration,
    Credentials,
    DELEGATE,
    FileAttachment,
    HTMLBody,
    Mailbox,
    Message,
)



# ── Module-level constants ──────────────────────────────────────────────────
DATE_STR = datetime.now().strftime("%d-%m-%Y")
BASE_DIR_FORMAT = datetime.now().strftime("%d.%m.%y")
BASE_DIR = f"s:\\{BASE_DIR_FORMAT}"


# ── DataFrame helpers ───────────────────────────────────────────────────────

def insert_or_update(df: pd.DataFrame, row_id, new_value) -> pd.DataFrame:
    """Insert a new row or update the 'active' column for an existing date.

    Parameters
    ----------
    df : pd.DataFrame
        Must contain columns 'date' and 'active'.
    row_id : date-like
        The date value to match or insert.
    new_value
        The value to write into the 'active' column.

    Returns
    -------
    pd.DataFrame
        Updated DataFrame (the original is **not** mutated when a new row is
        added; always use the returned value).
    """
    if df["date"].eq(row_id).any():
        df.loc[df["date"] == row_id, "active"] = new_value
    else:
        new_row = pd.DataFrame({"date": [row_id], "active": [new_value]})
        df = pd.concat([df, new_row], ignore_index=True)
    return df


def addInitials(row: pd.Series) -> str:
    """Prefix an FT number with a document-type abbreviation if it doesn't
    already contain one (indicated by an underscore in the value).

    Supports: Invoice → INV, Payment → PMT, JV → JV,
              Debit Note → DN, Credit Note → CN.
    """
    ftnr = row["Ftnr"]
    if "_" in ftnr:
        return ftnr

    # Map doc-type labels to their short prefixes
    prefix_map = {
        "Invoice": "INV",
        "Payment": "PMT",
        "JV": "JV",
        "Debit Note": "DN",
        "Credit Note": "CN",
    }
    doc_type = row["Doc Type_x"]
    prefix = prefix_map.get(doc_type)
    return f"{prefix}_{ftnr}" if prefix else ftnr


# ── File system operations ──────────────────────────────────────────────────

def create_folder_if_not_exists(folder_path: str) -> None:
    """Create *folder_path* (and any intermediate parents) if it doesn't exist."""
    os.makedirs(folder_path, exist_ok=True)


def search_files(folder_path: str, file_name: str) -> List[str]:
    """Recursively walk *folder_path* and return paths whose filename contains
    *file_name* (case-insensitive match).
    """
    file_name_lower = file_name.lower()
    return [
        os.path.join(dirpath, fname)
        for dirpath, _, filenames in os.walk(folder_path)
        for fname in filenames
        if file_name_lower in fname.lower()
    ]


def list_files(directory: str) -> List[str]:
    """Return the names of all items in *directory*."""
    return os.listdir(directory)


def checkModificationDate(filepath: str) -> bool:
    """Return ``True`` if *filepath* was last modified today."""
    mod_time = os.path.getmtime(filepath)
    mod_date = dtm.datetime.fromtimestamp(mod_time).date()
    return mod_date == dtm.datetime.now().date()


def getModificationDate(filepath: str) -> float:
    """Return the modification timestamp (epoch seconds) of *filepath*."""
    return os.path.getmtime(filepath)


# ── Compression ─────────────────────────────────────────────────────────────

def compress_files_gzip(
    file_list: List[str],
    output_filename: str,
    compression_ratio: int,
) -> None:
    """Concatenate all files in *file_list* into a single gzip archive."""
    with open(output_filename, "wb") as out_f:
        with gzip.GzipFile(fileobj=out_f, compresslevel=compression_ratio) as gz:
            for file_name in file_list:
                with open(file_name, "rb") as in_f:
                    shutil.copyfileobj(in_f, gz)


def compress_files(
    output_filename: str,
    directory: str,
    file_list: List[str],
) -> None:
    """Create a *.tar.gz* archive from files in *directory* at maximum compression."""
    with tarfile.open(output_filename, "w:gz", compresslevel=9) as tar:
        for file_name in file_list:
            tar.add(os.path.join(directory, file_name))


# ── Logging ─────────────────────────────────────────────────────────────────

def write_log(message: str) -> str:
    """Append a timestamped *message* to today's log file.

    Returns ``'OK'`` on success, or the exception string on failure.
    """
    try:
        log_date = datetime.now().strftime("%d-%m-%Y")
        timestamp = datetime.now().strftime("[%d.%m.%y-%H:%M:%S] ")
        with open(f"{log_date} - log.txt", "a") as f:
            f.write(f"{timestamp}{message}\n")
        return "OK"
    except Exception as exc:
        return str(exc)


# ── Email ───────────────────────────────────────────────────────────────────


def _to_mailboxes(addresses: Union[str, List[str]]) -> List[Mailbox]:
    """Normalize string or list[str] into Mailbox objects."""
    if isinstance(addresses, str):
        addresses = [addresses]
    return [Mailbox(email_address=addr) for addr in addresses]

def sendMail(
    subject: str,
    body: str,
    to: Union[str, List[str]] = "mturky@cne.com.eg",
    cc: Optional[List[str]] = None,
    attachments: Optional[List[str]] = None,
    username: str = "mturky",
    password: str = "m0h@mmed",
) -> None:
    """Send an email via Exchange Web Services."""

    credentials = Credentials(username=username, password=password)

    config = Configuration(
        server="mail.cne.com.eg",
        credentials=credentials,
    )

    account = Account(
        primary_smtp_address="mturky@cne.com.eg",
        credentials=credentials,
        config=config,
        autodiscover=False,
    )

    msg = Message(
        account=account,
        subject=f"[Auto] {subject}",
        body=HTMLBody(body),
        to_recipients=_to_mailboxes(to),
        cc_recipients=_to_mailboxes(cc) if cc else [],
        bcc_recipients=[],
    )

    if attachments:
        for file_path in attachments:
            with open(file_path, "rb") as f:
                msg.attach(
                    FileAttachment(
                        name=file_path.split("/")[-1],
                        content=f.read(),
                    )
                )

    msg.send()


# ── DTH report helper ──────────────────────────────────────────────────────

def createDTHfile(beinfile: str, dth_file_name: str) -> None:
    """Generate a DTH (Direct-To-Home) subscriber CSV from a BeinData export.

    Filters for active subscriptions of recognised customer types, excludes
    promotional plans, and keeps only the most recent contract per customer.
    """
    df = pd.read_csv(beinfile, dtype="str")

    # Customer types that qualify as DTH subscribers
    DTH_SUB_TYPES = [
        "beIN Quartar Installment",
        "CNE Subscriber",
        "MCE staff (CNE staff)",
        "BeIN sports CC",
        "beIN Bi Installment",
        "Corporate Subscriber",
        "Temp",
        "Bein NC",
        "Bulk DTH customer",
        "beIN Installment Sub",
        "Charge Back",
    ]

    # Promotional / temporary plans to exclude
    PLANS_TO_EXCLUDE = [
        "AFCON23ADDNov",
        "EURO24ADDNov",
        "AFC23-ADD-Nov",
        "AFCON23-SA-Nov",
        "AFC23-SA-Nov",
        "EURO24SANov",
        "Temp01",
    ]

    # Filter & deduplicate
    dth = df.loc[
        (df["Status"] == "Active")
        & df["Customer Type"].isin(DTH_SUB_TYPES)
        & ~df["Plan"].isin(PLANS_TO_EXCLUDE)
    ].copy()

    dth["End Date"] = pd.to_datetime(dth["End Date"])
    dth = (
        dth.sort_values(["Customer Number", "End Date"], ascending=False)
        .drop_duplicates(subset=["Customer Number"], keep="first")
    )

    # Select output columns
    output_cols = [
        "Customer Number",
        "Customer Type",
        "Plan",
        "Decoder",
        "Item Description STB",
        "Smart Card",
        "Item Description SC",
    ]
    dth[output_cols].to_csv(
        f"{BASE_DIR}/{DATE_STR} {dth_file_name}.csv", index=False
    )