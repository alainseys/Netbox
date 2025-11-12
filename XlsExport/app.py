"""
Export:
 - IP Ranges
 - IP Addresses
 -> TO XLS => Email 
"""
import argparse
import sys
from pathlib import Path
from typing import List, Dict, Any
from email.message import EmailMessage
import smtplib

import requests
import urllib3
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import os
from dotenv import load_dotevn
# --------------------------------------------------------------------------- #
# Suppress SSL warnings
# --------------------------------------------------------------------------- #
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --------------------------------------------------------------------------- #
# Hardcoded Settings
# --------------------------------------------------------------------------- #

email_to = os.getenv("EMAIL_TO")
smtp_server = os.getenv("SMTP_SERVER")
smtp_port = os.getenv("SMTP_PORT")
from_email = os.getenv("FROM_EMAIL")
url_ip_ranges = os.getenv("URL_IP_RANGES")
url_ip_addresses = os.getenv("URL_IP_ADDRESSES")

OUTPUT_IP_RANGES = Path("ip_ranges.xlsx")
OUTPUT_IP_ADDRESSES = Path("ip_addresses.xlsx")

# --------------------------------------------------------------------------- #
# Safe string conversions
# --------------------------------------------------------------------------- #
def safe_string(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (str, int, float)):
        return str(value)
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, dict):
        return (
            value.get("display")
            or value.get("name")
            or value.get("label")
            or value.get("value")
            or str(value)
        )
    if isinstance(value, list):
        items = [
            item.get("display") or item.get("name") or item.get("label") or str(item)
            for item in value
        ]
        return ", ".join(filter(None, items))
    return str(value)


