"""
Direct SAP GUI Connection Module

This module provides direct automation with SAP GUI without requiring script downloads
or separate API servers. It uses win32com to directly control SAP GUI.

Requirements:
- Python for Windows with pywin32 package installed
- SAP GUI for Windows installed with scripting enabled
- SAP system accessible

Usage:
1. Import DirectSapConnection class
2. Create an instance
3. Call methods directly

Example:
    sap = DirectSapConnection()
    sap.connect()
    sap.login(username, password)
    order_details = sap.get_service_order_details(order_number)
"""
import os
import logging
import time
import sys
import platform
from config import SPEX_CUSTOMER_NUMBERS

# Check if we're on Windows and can import win32com
WINDOWS_PLATFORM = platform.system() == "Windows"
WIN32COM_AVAILABLE = False

if WINDOWS_PLATFORM:
    try:
        import win32com.client
        WIN32COM_AVAILABLE = True
    except ImportError:
        logging.warning("win32com not available. Direct SAP integration will be disabled.")
else:
    logging.warning("Not running on Windows. Direct SAP integration will be disabled.")

logger = logging.getLogger(__name__)

class DirectSapConnectionBase:
    """Base class for SAP connection"""
    def __init__(self):
        self.connected = False
        self.logged_in = False
        self.sap_gui = None
        self.session = None
        logger.info("SAP Connection base initialized")
    
    def connect(self):
        """Connect to SAP"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def login(self, username=None, password=None):
        """Login to SAP"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def get_service_order_details(self, order_number):
        """Get service order details from SAP"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def open_ziwbn(self, order_number):
        """Open ZIWBN transaction"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def labor_on(self, order_number):
        """Set labor on process"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def labor_off(self, order_number):
        """Set labor off process"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def update_findings(self, order_number, mods_in, mods_out):
        """Update findings information"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def update_wanding_status(self, order_number):
        """Update wanding status"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def update_wsupd_comments(self, order_number, comments):
        """Update WSUPD comments"""
        raise NotImplementedError("This method must be implemented by subclasses")
    
    def print_service_report(self, order_number):
        """Print service report"""
        raise NotImplementedError("This method must be implemented by subclasses")


class DirectSapConnectionUnavailable(DirectSapConnectionBase):
    """Fallback class used when direct SAP connection is not available"""
    
    def __init__(self):
        super().__init__()
        if not WINDOWS_PLATFORM:
            self.reason = "Direct SAP connection is only available on Windows platforms"
        else:
            self.reason = "Direct SAP connection is not available (pywin32 not installed)"
        logger.warning(self.reason)
    
    def _not_available(self, method_name):
        """Helper to log and return appropriate value when a method is called"""
        logger.warning(f"Cannot execute {method_name}: {self.reason}")
        # For get_service_order_details specifically, return a structured error message
        if method_name == "get_service_order_details":
            return {
                "error": True,
                "message": self.reason,
                "details": "The Direct SAP Connection mode requires a Windows environment with SAP GUI installed. "
                           "Please switch to another connection mode or run this application on Windows."
            }
        return None
    
    def connect(self):
        """Report connection failure"""
        return self._not_available("connect")
    
    def login(self, username=None, password=None):
        """Report login failure"""
        return self._not_available("login")
    
    def get_service_order_details(self, order_number):
        """Report get_service_order_details failure"""
        return self._not_available("get_service_order_details")
    
    def open_ziwbn(self, order_number):
        """Report open_ziwbn failure"""
        return self._not_available("open_ziwbn")
    
    def labor_on(self, order_number):
        """Report labor_on failure"""
        return self._not_available("labor_on")
    
    def labor_off(self, order_number):
        """Report labor_off failure"""
        return self._not_available("labor_off")
    
    def update_findings(self, order_number, mods_in, mods_out):
        """Report update_findings failure"""
        return self._not_available("update_findings")
    
    def update_wanding_status(self, order_number):
        """Report update_wanding_status failure"""
        return self._not_available("update_wanding_status")
    
    def update_wsupd_comments(self, order_number, comments):
        """Report update_wsupd_comments failure"""
        return self._not_available("update_wsupd_comments")
    
    def print_service_report(self, order_number):
        """Report print_service_report failure"""
        return self._not_available("print_service_report")


class DirectSapConnection(DirectSapConnectionBase):
    """Direct SAP GUI Connection class that uses win32com to control SAP GUI directly"""
    
    def __init__(self):
        self.connected = False
        self.logged_in = False
        self.sap_gui = None
        self.session = None
        logger.info("Direct SAP Connection initialized")
    
    def connect(self):
        """Connect to SAP GUI directly"""
        logger.info("Connecting to SAP GUI directly...")
        
        try:
            # Get SAP GUI Scripting object
            try:
                # Try using GetObject first
                self.sap_gui = win32com.client.GetObject("SAPGUI")
            except Exception as e:
                logger.warning(f"Failed to connect with GetObject: {str(e)}")
                try:
                    # Try using Dispatch with correct ProgID
                    self.sap_gui = win32com.client.Dispatch("SAP.SAPGUI")
                except Exception as e2:
                    logger.warning(f"Failed to connect with Dispatch('SAP.SAPGUI'): {str(e2)}")
                    try:
                        # Try additional alternative ProgIDs
                        self.sap_gui = win32com.client.Dispatch("SAPGUI.ScriptingCtrl.1")
                    except Exception as e3:
                        logger.warning(f"Failed to connect with Dispatch('SAPGUI.ScriptingCtrl.1'): {str(e3)}")
                        # Final fallback
                        try:
                            self.sap_gui = win32com.client.Dispatch("SapROTWr.SapROTWrapper")
                        except Exception as e4:
                            logger.warning(f"Failed to connect with SapROTWrapper: {str(e4)}")
                
            # If we failed to get SAP GUI through normal means, try ROT (Running Object Table)
            if not self.sap_gui:
                try:
                    logger.info("Attempting to connect via ROT...")
                    # Try to use the ROT as a last resort
                    import pythoncom
                    from win32com.client import GetObject
                    
                    # Get the ROT
                    rot = pythoncom.GetRunningObjectTable()
                    rot_items = rot.EnumRunning()
                    
                    sap_found = False
                    for i in range(rot_items.GetCount()):
                        item = rot_items.Next()
                        name = rot.GetDisplayName(item, None)
                        logger.info(f"ROT item: {name}")
                        
                        # Look for SAP in the ROT
                        if "SAP" in name:
                            try:
                                # Attempt to get the SAP application from the ROT
                                self.sap_gui = GetObject(name)
                                sap_found = True
                                logger.info(f"Connected to SAP via ROT: {name}")
                                break
                            except Exception as e_rot:
                                logger.warning(f"Failed to connect to SAP via ROT item {name}: {str(e_rot)}")
                    
                    if not sap_found:
                        logger.warning("No SAP items found in the ROT")
                        
                except Exception as e_rot_overall:
                    logger.warning(f"Error accessing ROT: {str(e_rot_overall)}")
            
            # Final check if we have a valid SAP GUI reference
            if not self.sap_gui:
                error_msg = "Could not connect to SAP GUI. Make sure SAP GUI is installed and running."
                logger.error(error_msg)
                details = "Try one of the following:\n1. Ensure SAP GUI is running\n2. Restart SAP GUI\n3. Try a different connection mode"
                return {"error": True, "message": error_msg, "details": details}
                
            # Get the scripting engine
            try:
                application = self.sap_gui.GetScriptingEngine
            except Exception as e:
                # Try alternative property names for different SAP versions
                try:
                    # Some versions use ScriptingEngine instead of GetScriptingEngine
                    application = self.sap_gui.ScriptingEngine
                except Exception as e2:
                    error_msg = "Could not access SAP GUI Scripting engine. Make sure scripting is enabled in SAP."
                    logger.error(f"{error_msg} Error: {str(e)}, Alternate error: {str(e2)}")
                    return {"error": True, "message": error_msg, "details": "Ensure SAP scripting is enabled in SAP GUI options"}
            
            # Check if SAP GUI is running
            if application.Connections.Count > 0:
                # Use existing connection
                connection = application.Connections(0)
                self.session = connection.Children(0)
                self.connected = True
                logger.info("Connected to existing SAP GUI session")
            else:
                # SAP not running or no connections
                error_msg = "No active SAP connections found. Please log in to SAP first."
                logger.error(error_msg)
                return {"error": True, "message": error_msg}
                
        except Exception as e:
            error_msg = f"Error connecting to SAP GUI: {str(e)}"
            details = "Verify that: 1) SAP GUI is installed, 2) SAP GUI Scripting is enabled, 3) An active SAP session is open"
            logger.error(error_msg)
            return {"error": True, "message": error_msg, "details": details}
            
        return True
    
    def login(self, username=None, password=None):
        """Login to SAP directly"""
        if not self.connected:
            connection_result = self.connect()
            # Check if we got an error response
            if isinstance(connection_result, dict) and connection_result.get('error'):
                return connection_result
            elif not connection_result:
                return False
        
        try:
            # If we're using an existing connection, we might already be logged in
            # We'll just set the flag and proceed
            self.logged_in = True
            logger.info(f"Using existing SAP GUI session (login is assumed)")
            return True
        except Exception as e:
            error_msg = f"Error during login: {str(e)}"
            logger.error(error_msg)
            return {"error": True, "message": error_msg}
    
    def get_service_order_details(self, order_number):
        """Get service order details directly from SAP"""
        if not self.logged_in:
            login_result = self.login()
            if isinstance(login_result, dict) and login_result.get('error'):
                return login_result
            if not login_result:
                error_msg = "Not logged in to SAP"
                return {"error": True, "message": error_msg}
            
        logger.info(f"Getting service order details for {order_number} directly")
        
        try:
            # Navigate to IW33 - Service Order Display
            self.session.StartTransaction("IW33")
            
            # Enter service order number
            self.session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = order_number
            self.session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").CaretPosition = 12
            self.session.FindById("wnd[0]").SendVKey(0)  # ENTER
            
            # Get order details (adjust the fields as needed based on your SAP system)
            try:
                # Try to get customer number and other details
                customer_number = self.session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text
                part_number = self.session.FindById("wnd[0]/usr/tabsTABSTRIP_DETAIL/tabpINCL_DETAIL_TAB1/ssubDETAIL_1:SAPLITO0:0111/ctxtRIWO00-MATNR").Text
                serial_number = self.session.FindById("wnd[0]/usr/tabsTABSTRIP_DETAIL/tabpINCL_DETAIL_TAB1/ssubDETAIL_1:SAPLITO0:0111/ctxtRIWO00-SERIALNR").Text
                
                # SPEX determination logic
                is_spex = customer_number in SPEX_CUSTOMER_NUMBERS
                
                result = {
                    "order_number": order_number,
                    "customer_number": customer_number,
                    "is_spex": is_spex,
                    "part_number": part_number,
                    "serial_number": serial_number,
                    # Add more fields as needed
                }
                
                return result
            
            except Exception as e:
                error_msg = f"Error parsing service order details: {str(e)}"
                logger.error(error_msg)
                # Check if it's potentially a "not found" error
                if "FindById" in str(e) and "not found" in str(e).lower():
                    return {"error": True, "message": f"Service order {order_number} not found", 
                            "details": "The order number might be incorrect or not accessible."}
                return {"error": True, "message": error_msg}
        
        except Exception as e:
            error_msg = f"Error getting service order details: {str(e)}"
            logger.error(error_msg)
            
            # Add detailed troubleshooting info for SAP GUI errors
            if "StartTransaction" in str(e):
                details = "SAP GUI scripting may not be enabled. Please check your SAP configuration."
            elif "COM" in str(e) or "Dispatch" in str(e) or "GetObject" in str(e):
                details = "There appears to be an issue with the COM interface to SAP GUI. Verify SAP GUI is running properly."
            elif "syntax" in str(e).lower():
                details = "Invalid syntax error often indicates issues with SAP GUI scripting permissions or configuration."
            else:
                details = "Verify that SAP GUI is running, you are logged in, and that scripting is enabled."
                
            return {"error": True, "message": error_msg, "details": details}
    
    def open_ziwbn(self, order_number):
        """Open ZIWBN transaction directly"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Opening ZIWBN for {order_number} directly")
        
        try:
            # Navigate to ZIWBN
            self.session.StartTransaction("ZIWBN")
            
            # Enter service order
            self.session.FindById("wnd[0]/usr/ctxtS_AUFNR-LOW").Text = order_number
            self.session.FindById("wnd[0]/usr/ctxtS_AUFNR-LOW").CaretPosition = 12
            self.session.FindById("wnd[0]/tbar[1]/btn[8]").press()  # Execute
            
            # Select the order in the list (first row)
            self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
            self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell()
            
            return True
            
        except Exception as e:
            logger.error(f"Error opening ZIWBN for {order_number}: {str(e)}")
            return False
    
    def labor_on(self, order_number):
        """Set labor on directly"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Setting labor on for {order_number} directly")
        
        try:
            # Open the order in ZIWBN first
            if not self.open_ziwbn(order_number):
                return False
            
            # Navigate to labor tab
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0").Select()
            
            # Enter labor data for labor on
            current_date = time.strftime("%d.%m.%Y")  # Current date in DD.MM.YYYY format
            current_time = time.strftime("%H:%M:%S")  # Current time in HH:MM:SS format
            
            # Set date and time fields
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUB_0100:ZIWWB_CLOSEUP:0200/ctxtZTLABOR-DATUM").Text = current_date
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUB_0100:ZIWWB_CLOSEUP:0200/ctxtZTLABOR-UZEIT").Text = current_time
            
            # Click labor on button
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUB_0100:ZIWWB_CLOSEUP:0200/btnLABOR_ON").press()
            
            # Save
            self.session.FindById("wnd[0]/tbar[0]/btn[11]").press()
            
            return True
            
        except Exception as e:
            logger.error(f"Error setting labor on for {order_number}: {str(e)}")
            return False
    
    def labor_off(self, order_number):
        """Set labor off directly"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Setting labor off for {order_number} directly")
        
        try:
            # Open the order in ZIWBN first
            if not self.open_ziwbn(order_number):
                return False
            
            # Navigate to labor tab
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0").Select()
            
            # Enter labor data for labor off
            current_date = time.strftime("%d.%m.%Y")  # Current date in DD.MM.YYYY format
            current_time = time.strftime("%H:%M:%S")  # Current time in HH:MM:SS format
            
            # Set date and time fields
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUB_0100:ZIWWB_CLOSEUP:0200/ctxtZTLABOR-END_DATUM").Text = current_date
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUB_0100:ZIWWB_CLOSEUP:0200/ctxtZTLABOR-END_UZEIT").Text = current_time
            
            # Click labor off button
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUB_0100:ZIWWB_CLOSEUP:0200/btnLABOR_OFF").press()
            
            # Save
            self.session.FindById("wnd[0]/tbar[0]/btn[11]").press()
            
            return True
            
        except Exception as e:
            logger.error(f"Error setting labor off for {order_number}: {str(e)}")
            return False
    
    def update_findings(self, order_number, mods_in, mods_out):
        """Update findings directly"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating findings for {order_number} directly")
        
        try:
            # Open the order in ZIWBN first
            if not self.open_ziwbn(order_number):
                return False
            
            # Navigate to findings tab
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTAB3").Select()
            
            # Set findings fields
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTAB3/ssubSUB_0100:ZIWWB_CLOSEUP:0400/txtZIWBN_FINDINGS-MODS_IN").Text = mods_in
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTAB3/ssubSUB_0100:ZIWWB_CLOSEUP:0400/txtZIWBN_FINDINGS-MODS_OUT").Text = mods_out
            
            # Save
            self.session.FindById("wnd[0]/tbar[0]/btn[11]").press()
            
            return True
            
        except Exception as e:
            logger.error(f"Error updating findings for {order_number}: {str(e)}")
            return False
    
    def update_wanding_status(self, order_number):
        """Update wanding status directly"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating wanding status for {order_number} directly")
        
        try:
            # Open the order in ZIWBN first
            if not self.open_ziwbn(order_number):
                return False
            
            # Navigate to wanding tab
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTAB4").Select()
            
            # Check wanded checkbox
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTAB4/ssubSUB_0100:ZIWWB_CLOSEUP:0500/chkZIWBN_WANDING-WANDED").Selected = True
            
            # Set wanding date to current date
            current_date = time.strftime("%d.%m.%Y")  # Current date in DD.MM.YYYY format
            self.session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTAB4/ssubSUB_0100:ZIWWB_CLOSEUP:0500/ctxtZIWBN_WANDING-WANDING_DATE").Text = current_date
            
            # Save
            self.session.FindById("wnd[0]/tbar[0]/btn[11]").press()
            
            return True
            
        except Exception as e:
            logger.error(f"Error updating wanding status for {order_number}: {str(e)}")
            return False
    
    def update_wsupd_comments(self, order_number, comments):
        """Update WSUPD comments directly"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating WSUPD comments for {order_number} directly")
        
        try:
            # Navigate to WSUPD transaction
            self.session.StartTransaction("WSUPD")
            
            # Enter service order number
            self.session.FindById("wnd[0]/usr/ctxtVIQMEL-QMNUM").Text = order_number
            self.session.FindById("wnd[0]/usr/ctxtVIQMEL-QMNUM").CaretPosition = 12
            self.session.FindById("wnd[0]").SendVKey(0)  # ENTER
            
            # Add comments
            self.session.FindById("wnd[0]/usr/tblSAPLQS1LTCTRL_WRITER/txtLTXT-TDLINE[0,0]").Text = comments
            
            # Save
            self.session.FindById("wnd[0]/tbar[0]/btn[11]").press()
            
            return True
            
        except Exception as e:
            logger.error(f"Error updating WSUPD comments for {order_number}: {str(e)}")
            return False
    
    def print_service_report(self, order_number):
        """Print service report directly"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Printing service report for {order_number} directly")
        
        try:
            # Navigate to IW33 - Service Order Display
            self.session.StartTransaction("IW33")
            
            # Enter service order number
            self.session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = order_number
            self.session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").CaretPosition = 12
            self.session.FindById("wnd[0]").SendVKey(0)  # ENTER
            
            # Open print menu
            self.session.FindById("wnd[0]/tbar[1]/btn[44]").press()  # Print button
            
            # Select report type (may need adjustment based on your SAP system)
            self.session.FindById("wnd[1]/usr/cntlPRINT_PARAMETERS/shellcont/shell").selectedNode = "000001"
            self.session.FindById("wnd[1]/usr/cntlPRINT_PARAMETERS/shellcont/shell").doubleClickNode = "000001"
            
            # Print
            self.session.FindById("wnd[1]/tbar[0]/btn[13]").press()
            
            return True
            
        except Exception as e:
            logger.error(f"Error printing service report for {order_number}: {str(e)}")
            return False


# Factory function to create the appropriate direct SAP connection
def create_direct_sap_connection():
    """
    Create a direct SAP connection if the platform and requirements support it,
    otherwise return a placeholder implementation that logs warnings.
    
    This allows the application to gracefully handle the absence of Windows or win32com.
    """
    if WINDOWS_PLATFORM and WIN32COM_AVAILABLE:
        return DirectSapConnection()
    else:
        return DirectSapConnectionUnavailable()