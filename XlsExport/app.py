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
