"""
Configuration settings for the SAP Close-Up application
"""
import os

# Workflow steps in order
WORKFLOW_STEPS = [
    {
        "id": "service_order",
        "name": "Service Order Input",
        "template": "workflow_steps/service_order.html",
        "icon": "file-text"
    },
    {
        "id": "equipment_check",
        "name": "Equipment Verification",
        "template": "workflow_steps/equipment_check.html",
        "icon": "check-square"
    },
    {
        "id": "labor_on",
        "name": "Labor On",
        "template": "workflow_steps/labor_on.html",
        "icon": "play-circle"
    },
    {
        "id": "paperwork",
        "name": "Paperwork Verification",
        "template": "workflow_steps/paperwork.html",
        "icon": "clipboard"
    },
    {
        "id": "test_sheet",
        "name": "Test Sheet Verification",
        "template": "workflow_steps/test_sheet.html",
        "icon": "file-check"
    },
    {
        "id": "customer_req",
        "name": "Customer Requirements",
        "template": "workflow_steps/customer_req.html", 
        "icon": "users"
    },
    {
        "id": "notifications",
        "name": "Notifications",
        "template": "workflow_steps/notifications.html",
        "icon": "bell"
    },
    {
        "id": "warranty",
        "name": "Warranty Check",
        "template": "workflow_steps/warranty.html",
        "icon": "shield"
    },
    {
        "id": "findings",
        "name": "Findings Update",
        "template": "workflow_steps/findings.html",
        "icon": "edit"
    },
    {
        "id": "labor_off",
        "name": "Labor Off",
        "template": "workflow_steps/labor_off.html",
        "icon": "stop-circle"
    },
    {
        "id": "final_inspection",
        "name": "Final Inspection",
        "template": "workflow_steps/final_inspection.html",
        "icon": "search"
    },
    {
        "id": "complete",
        "name": "Complete",
        "template": "workflow_steps/complete.html",
        "icon": "check-circle"
    }
]

# SPEX customer numbers
SPEX_CUSTOMER_NUMBERS = [
    "PLANT1133",
    "SLSR01",
    "PLANT1057",
    "PLANT1052",
    "PLANT1013",
    "PLANT1103", 
    "PLANT1116",
    "PLANT1005"
]

# SAP Connection Configuration
# In a production environment, these would be environment variables
SAP_CONFIG = {
    "ashost": os.environ.get("SAP_HOST", ""),
    "sysnr": os.environ.get("SAP_SYSNR", ""),
    "client": os.environ.get("SAP_CLIENT", ""),
    "user": os.environ.get("SAP_USER", ""),
    "passwd": os.environ.get("SAP_PASSWORD", "")
}
