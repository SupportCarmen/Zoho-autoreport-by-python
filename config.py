import os
from dotenv import load_dotenv

load_dotenv()

BASE = os.path.join(os.path.expanduser("~"), "Downloads")

WEBHOOK = "https://discord.com/api/webhooks/1481557932893274256/aPMr8fpYSMSAJ3MU3So3FocWdcYqOT2FW_8M_OVZlBB_wnzT7IeUmYnJGpLr0nxTzDyJ"
ZOHO_EMAIL = os.getenv("ZOHO_EMAIL")
ZOHO_PASSWORD = os.getenv("ZOHO_PASSWORD")
DASHBOARD_URL = "https://desk.zoho.com/agent/carmensoftware/carmen-software-support/dashboards/details/483929000025299144"

FOLDER = os.path.join(BASE, "captureReport")
REPORT_FOLDER = os.path.join(BASE, "report")

REPORTS = [
    {
        "url": "https://desk.zoho.com/supportapi/api/v1/reports/483929000037008035/export?orgId=710033074&includeDetails=true&from=0&limit=2000&format=xls",
        "name": "OpenAll"
    },
    {
        "url": "https://desk.zoho.com/supportapi/api/v1/reports/483929000029190842/export?orgId=710033074&includeDetails=true&from=0&limit=2000&format=xls",
        "name": "TicketToday"
    }
]
