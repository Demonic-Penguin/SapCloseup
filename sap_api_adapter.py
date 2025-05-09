"""
SAP API Adapter - Bridge between Web App and SAP GUI

This script runs on a Windows machine with SAP GUI installed.
It exposes a REST API that the web application can call to communicate with SAP GUI.

Requirements:
- Python 3.6+ with Flask and win32com installed
- SAP GUI installed and running on the same machine
- SAP system accessible

Installation:
1. Install required packages: pip install flask requests
2. Ensure SAP GUI scripting is enabled
3. Run this script with: python sap_api_adapter.py
"""
import os
import sys
import logging
import time
import json
from datetime import datetime
from flask import Flask, request, jsonify

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("sap_adapter.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)

# Configure API security
API_KEY = os.environ.get("SAP_API_KEY", "change-this-in-production")

# Store SAP connection state
sap_state = {
    "connected": False,
    "logged_in": False,
    "session": None,
    "application": None,
    "connection": None
}

def authenticate_request():
    """Verify API key in request header"""
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return False
    
    token = auth_header.split(' ')[1]
    return token == API_KEY

def get_sap_session():
    """Get or create SAP session"""
    try:
        if sap_state["session"] is None:
            # Import COM objects for SAP GUI
            try:
                import win32com.client
            except ImportError:
                logger.error("win32com.client not available - are you running on Windows?")
                return None

            # Connect to SAP
            try:
                sap_state["application"] = win32com.client.GetObject("SAPGUI")
                if not sap_state["application"]:
                    return None
                
                sap_state["connection"] = sap_state["application"].Children(0)
                if not sap_state["connection"]:
                    return None
                
                sap_state["session"] = sap_state["connection"].Children(0)
                if not sap_state["session"]:
                    return None
                
                sap_state["connected"] = True
                return sap_state["session"]
            except Exception as e:
                logger.error(f"Error connecting to SAP: {e}")
                return None
        else:
            return sap_state["session"]
    except Exception as e:
        logger.error(f"Error in get_sap_session: {e}")
        return None

@app.route('/api/sap/status', methods=['GET'])
def status():
    """Check SAP connection status"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    return jsonify({
        "connected": sap_state["connected"],
        "logged_in": sap_state["logged_in"],
        "timestamp": datetime.now().isoformat()
    })

@app.route('/api/sap/connect', methods=['POST'])
def connect():
    """Connect to SAP"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    try:
        session = get_sap_session()
        
        if session:
            sap_state["connected"] = True
            return jsonify({"connected": True})
        else:
            return jsonify({"connected": False, "error": "Could not connect to SAP"}), 500
    except Exception as e:
        logger.error(f"Error connecting to SAP: {e}")
        return jsonify({"connected": False, "error": str(e)}), 500

@app.route('/api/sap/login', methods=['POST'])
def login():
    """Login to SAP"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    data = request.json
    username = data.get('username')
    password = data.get('password')
    
    try:
        session = get_sap_session()
        
        if not session:
            return jsonify({"logged_in": False, "error": "Not connected to SAP"}), 500
        
        # Check if we're already logged in - check user from session info
        try:
            current_user = session.Info.User
            if current_user and current_user.strip():
                sap_state["logged_in"] = True
                return jsonify({"logged_in": True, "user": current_user})
        except:
            # Not logged in, that's okay
            pass
        
        # Login logic would go here - typically not needed if SAP GUI is already open
        # This would be a placeholder for custom login screens if needed
        
        # Assume we're logged in if we can access user info
        try:
            current_user = session.Info.User
            if current_user and current_user.strip():
                sap_state["logged_in"] = True
                return jsonify({"logged_in": True, "user": current_user})
            else:
                return jsonify({"logged_in": False, "error": "Login failed - no user detected"}), 401
        except Exception as e:
            logger.error(f"Error checking login status: {e}")
            return jsonify({"logged_in": False, "error": str(e)}), 500
    except Exception as e:
        logger.error(f"Error in login: {e}")
        return jsonify({"logged_in": False, "error": str(e)}), 500

@app.route('/api/sap/service_order/<order_number>', methods=['GET'])
def get_service_order(order_number):
    """Get service order details"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    try:
        # Go to ZIWBN transaction
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZIWBN"
        session.findById("wnd[0]").sendVKey(0)
        
        # Input service order
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text = order_number
        session.findById("wnd[0]").sendVKey(0)
        
        # Get order data
        try:
            # Get customer number
            customer_number = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM").text
            
            # Get status
            subord_status = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").text
            
            # Get equipment info
            # Navigate to equipment tab
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H").select()
            
            # Try to get part number and serial number
            part_number = ""
            serial_number = ""
            try:
                # Try version 1 path
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").currentCellColumn = "MATNR"
                part_number = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "MATNR")
                serial_number = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "SERNR")
            except:
                # Try version 2 path (some SAP versions have a different path)
                try:
                    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").currentCellColumn = "MATNR"
                    part_number = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "MATNR")
                    serial_number = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "SERNR")
                except:
                    logger.warning(f"Could not get part/serial info for order {order_number}")
            
            # Go back to service order tab
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select()
            
            # Check for CONV in status
            is_converted = "CONV" in subord_status
            # Check for EXCH in status
            is_exchange = "EXCH" in subord_status
            # Check for SPEX customer
            is_spex = any(spex_cust in customer_number for spex_cust in ["PLANT1133", "SLSR01", "PLANT1057", "PLANT1052", "PLANT1013", "PLANT1103", "PLANT1116", "PLANT1005"])
            # Check for DPMI
            is_dpmi = "DPMI" in subord_status
            
            # Compile data
            order_data = {
                "order_number": order_number,
                "customer_number": customer_number,
                "part_number": part_number,
                "serial_number": serial_number,
                "is_spex": is_spex,
                "is_converted": is_converted,
                "is_exchange": is_exchange,
                "is_dpmi": is_dpmi,
                "status": subord_status
            }
            
            return jsonify(order_data)
            
        except Exception as e:
            logger.error(f"Error getting service order details: {e}")
            return jsonify({"error": f"Error getting service order details: {str(e)}"}), 500
    except Exception as e:
        logger.error(f"Error in ZIWBN transaction: {e}")
        return jsonify({"error": f"Error in ZIWBN transaction: {str(e)}"}), 500

@app.route('/api/sap/open_ziwbn', methods=['POST'])
def open_ziwbn():
    """Open ZIWBN transaction for a service order"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    data = request.json
    order_number = data.get('order_number')
    
    if not order_number:
        return jsonify({"error": "Order number is required"}), 400
    
    try:
        # Go to ZIWBN transaction
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZIWBN"
        session.findById("wnd[0]").sendVKey(0)
        
        # Input service order
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text = order_number
        session.findById("wnd[0]").sendVKey(0)
        
        return jsonify({"success": True})
    except Exception as e:
        logger.error(f"Error opening ZIWBN for {order_number}: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/sap/labor_on', methods=['POST'])
def labor_on():
    """Set labor on for service order"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    data = request.json
    order_number = data.get('order_number')
    
    if not order_number:
        return jsonify({"error": "Order number is required"}), 400
    
    try:
        # Implementation matches the VBS script from Close-Up Script Olathe.vbs
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellRow = -1
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn("LTXA1")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu()
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem("&FILTER")
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Close Up*"
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 6
        session.findById("wnd[1]").sendVKey(0)
        
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton("LABON")
        
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").setCurrentCell(-1, "LTXA1")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn("LTXA1")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu()
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem("&FILTER")
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = ""
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 0
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        return jsonify({"success": True})
    except Exception as e:
        logger.error(f"Error setting labor on for {order_number}: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/sap/labor_off', methods=['POST'])
def labor_off():
    """Set labor off for service order"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    data = request.json
    order_number = data.get('order_number')
    
    if not order_number:
        return jsonify({"error": "Order number is required"}), 400
    
    try:
        # Implementation based on the VBS script
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton("LABOFF")
        
        return jsonify({"success": True})
    except Exception as e:
        logger.error(f"Error setting labor off for {order_number}: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/sap/update_findings', methods=['POST'])
def update_findings():
    """Update findings for service order"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    data = request.json
    order_number = data.get('order_number')
    mods_in = data.get('mods_in', '')
    mods_out = data.get('mods_out', '')
    
    if not order_number:
        return jsonify({"error": "Order number is required"}), 400
    
    try:
        # This will need to be implemented based on the specific fields in your ZIWBN transaction
        # The implementation would navigate to the findings section and update fields
        
        # Example implementation:
        # Go to findings tab or section
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpFINDINGS").select()
        
        # Update mods in field
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpFINDINGS/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0214/txtEQBS_MODS_IN").text = mods_in
        
        # Update mods out field
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpFINDINGS/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0214/txtEQBS_MODS_OUT").text = mods_out
        
        # Save changes
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        
        return jsonify({"success": True})
    except Exception as e:
        logger.error(f"Error updating findings for {order_number}: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/sap/update_wanding', methods=['POST'])
def update_wanding():
    """Update wanding status for service order"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    data = request.json
    order_number = data.get('order_number')
    
    if not order_number:
        return jsonify({"error": "Order number is required"}), 400
    
    try:
        # Implementation based on the VBS script
        # This would be customized based on your exact SAP GUI screens
        
        # Example implementation:
        # Navigate to wanding status screen or field
        # session.findById("...").select()
        # Update status
        # session.findById("...").text = "Completed"
        # Save
        # session.findById("wnd[0]/tbar[0]/btn[11]").press()
        
        # Since the actual implementation depends on your specific SAP screens,
        # this is a placeholder for now
        logger.info(f"Updating wanding status for {order_number}")
        
        return jsonify({"success": True})
    except Exception as e:
        logger.error(f"Error updating wanding status for {order_number}: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/sap/update_comments', methods=['POST'])
def update_comments():
    """Update WSUPD comments for service order"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    data = request.json
    order_number = data.get('order_number')
    comments = data.get('comments', '')
    
    if not order_number:
        return jsonify({"error": "Order number is required"}), 400
    
    try:
        # Implementation based on the VBS script
        # This would be customized based on your exact SAP GUI screens
        
        # Example implementation:
        # Navigate to comments section
        # session.findById("...").select()
        # Update comments
        # session.findById("...").text = comments
        # Save
        # session.findById("wnd[0]/tbar[0]/btn[11]").press()
        
        # Since the actual implementation depends on your specific SAP screens,
        # this is a placeholder for now
        logger.info(f"Updating WSUPD comments for {order_number}: {comments}")
        
        return jsonify({"success": True})
    except Exception as e:
        logger.error(f"Error updating WSUPD comments for {order_number}: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/sap/print_report', methods=['POST'])
def print_report():
    """Print service report for service order"""
    if not authenticate_request():
        return jsonify({"error": "Unauthorized access"}), 401
    
    session = get_sap_session()
    if not session:
        return jsonify({"error": "Not connected to SAP"}), 500
    
    if not sap_state["logged_in"]:
        return jsonify({"error": "Not logged in to SAP"}), 401
    
    data = request.json
    order_number = data.get('order_number')
    
    if not order_number:
        return jsonify({"error": "Order number is required"}), 400
    
    try:
        # Implementation based on the VBS script
        # This would be customized based on your exact SAP GUI screens
        
        # Example implementation:
        # Navigate to print screen
        # session.findById("wnd[0]/tbar[0]/btn[45]").press()
        # Select report
        # session.findById("...").select()
        # Print
        # session.findById("...").press()
        
        # Since the actual implementation depends on your specific SAP screens,
        # this is a placeholder for now
        logger.info(f"Printing service report for {order_number}")
        
        return jsonify({"success": True})
    except Exception as e:
        logger.error(f"Error printing service report for {order_number}: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    # Default to port 5001 to avoid conflict with the main app
    port = int(os.environ.get("PORT", 5001))
    host = os.environ.get("HOST", "0.0.0.0")
    
    print(f"Starting SAP API Adapter on {host}:{port}")
    print(f"API Key is set to: {API_KEY}")
    print("Make sure SAP GUI is running and accessible")
    
    app.run(host=host, port=port, debug=False)