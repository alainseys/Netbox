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

# --------------------------------------------------------------------------- #
# Flatten IP Range
# --------------------------------------------------------------------------- #
def flatten_ip_range(obj: Dict[str, Any]) -> Dict[str, Any]:
    flat = {
        "ID": obj.get("id"),
        "Display": obj.get("display"),
        "Start Address": obj.get("start_address"),
        "End Address": obj.get("end_address"),
        "Size": obj.get("size"),
        "Family": safe_string(obj.get("family")),
        "Status": safe_string(obj.get("status")),
        "VRF": safe_string(obj.get("vrf")),
        "Tenant": safe_string(obj.get("tenant")),
        "Role": safe_string(obj.get("role")),
        "Description": obj.get("description", ""),
        "Comments": obj.get("comments", ""),
        "Mark Utilized": obj.get("mark_utilized"),
        "Created": obj.get("created"),
        "Last Updated": obj.get("last_updated"),
    }
    flat["Tags"] = safe_string(obj.get("tags", []))
    for cf_key, cf_val in obj.get("custom_fields", {}).items():
        flat[f"CF: {cf_key}"] = safe_string(cf_val)
    return flat

# --------------------------------------------------------------------------- #
# Flatten IP Address
# --------------------------------------------------------------------------- #
def flatten_ip_address(obj: Dict[str, Any]) -> Dict[str, Any]:
    flat = {
        "ID": obj.get("id"),
        "Display": obj.get("display"),
        "Address": obj.get("address"),
        "Family": safe_string(obj.get("family")),
        "VRF": safe_string(obj.get("vrf")),
        "Tenant": safe_string(obj.get("tenant")),
        "Status": safe_string(obj.get("status")),
        "Role": safe_string(obj.get("role")),
        "Assigned To": safe_string(obj.get("assigned_object")),
        "DNS Name": obj.get("dns_name", ""),
        "Description": obj.get("description", ""),
        "Comments": obj.get("comments", ""),
        "NAT Inside": safe_string(obj.get("nat_inside")),
        "NAT Outside": safe_string(obj.get("nat_outside")),
        "Created": obj.get("created"),
        "Last Updated": obj.get("last_updated"),
    }
    flat["Tags"] = safe_string(obj.get("tags", []))
    for cf_key, cf_val in obj.get("custom_fields", {}).items():
        flat[f"CF: {cf_key}"] = safe_string(cf_val)
    return flat

# --------------------------------------------------------------------------- #
# Fetch all pages
# --------------------------------------------------------------------------- #
def fetch_all_pages(session: requests.Session, url: str) -> List[Dict[str, Any]]:
    results = []
    next_url = url
    while next_url:
        resp = session.get(next_url, verify=False)
        resp.raise_for_status()
        data = resp.json()
        results.extend(data["results"])
        next_url = data.get("next")
    return results


