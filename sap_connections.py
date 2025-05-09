"""
This module handles SAP connections and interactions
Provides mock interface for development, real SAP interface via API, 
and local script generation for direct SAP GUI automation
"""
import logging
import time
import random
import os
import requests
import json
from config import SPEX_CUSTOMER_NUMBERS, SAP_SCRIPT_DIR
import config as config_module

logger = logging.getLogger(__name__)

class MockSapConnection:
    """Mock SAP connection class for development and testing"""
    
    def __init__(self):
        self.connected = False
        self.logged_in = False
        logger.info("Mock SAP Connection initialized")
    
    def connect(self):
        """Simulate connection to SAP"""
        logger.info("Connecting to SAP...")
        time.sleep(1)  # Simulate connection time
        self.connected = True
        return self.connected
    
    def login(self, username=None, password=None):
        """Simulate login to SAP"""
        if not self.connected:
            self.connect()
        
        logger.info(f"Logging in as {username}...")
        time.sleep(1)  # Simulate login time
        self.logged_in = True
        return self.logged_in
    
    def get_service_order_details(self, order_number):
        """Get service order details from SAP"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
        
        logger.info(f"Getting service order details for {order_number}")
        time.sleep(1)  # Simulate API call
        
        # Mock data for demonstration
        customer_num = random.choice(SPEX_CUSTOMER_NUMBERS) if random.random() > 0.7 else f"CUST{random.randint(10000, 99999)}"
        
        # Check if it's a valid service order format (usually numeric)
        if not order_number.isdigit():
            return None
            
        result = {
            "order_number": order_number,
            "customer_number": customer_num,
            "is_spex": customer_num in SPEX_CUSTOMER_NUMBERS,
            "is_converted": random.random() > 0.8,
            "is_exchange": random.random() > 0.9,
            "part_number": f"PN-{random.randint(100000, 999999)}",
            "serial_number": f"SN-{random.randint(10000, 99999)}"
        }
        
        return result
    
    def open_ziwbn(self, order_number):
        """Simulate opening ZIWBN transaction"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Opening ZIWBN for {order_number}")
        time.sleep(1)  # Simulate transaction
        return True
    
    def labor_on(self, order_number):
        """Simulate labor on process"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Setting labor on for {order_number}")
        time.sleep(1)  # Simulate transaction
        return True
        
    def labor_off(self, order_number):
        """Simulate labor off process"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Setting labor off for {order_number}")
        time.sleep(1)  # Simulate transaction
        return True
    
    def update_findings(self, order_number, mods_in, mods_out):
        """Update findings information"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating findings for {order_number}")
        logger.info(f"Mods In: {mods_in}")
        logger.info(f"Mods Out: {mods_out}")
        time.sleep(1)  # Simulate transaction
        return True
    
    def update_wanding_status(self, order_number):
        """Update wanding status"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating wanding status for {order_number}")
        time.sleep(1)  # Simulate transaction
        return True
    
    def update_wsupd_comments(self, order_number, comments):
        """Update WSUPD comments"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating WSUPD comments for {order_number}: {comments}")
        time.sleep(1)  # Simulate transaction
        return True
    
    def print_service_report(self, order_number):
        """Print service report"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Printing service report for {order_number}")
        time.sleep(1)  # Simulate printing
        return True


class LocalSapConnection:
    """SAP connection class that generates VBS scripts for direct SAP GUI automation"""
    
    def __init__(self):
        self.connected = False
        self.logged_in = False
        self.script_dir = SAP_SCRIPT_DIR
        
        # Create script directory if it doesn't exist
        if not os.path.exists(self.script_dir):
            os.makedirs(self.script_dir)
            
        logger.info("Local SAP Connection initialized")
    
    def _generate_script(self, filename, script_content):
        """Generate a VBS script file that can be downloaded and executed"""
        script_path = os.path.join(self.script_dir, filename)
        with open(script_path, 'w') as f:
            f.write(script_content)
        logger.info(f"Generated script: {script_path}")
        return script_path
    
    def _generate_script_template(self, order_number=None):
        """Generate the base script template with common functions"""
        script = """' SAP GUI Scripting Auto-generated by SAP Close-Up Automation Tool
' -----------------------------------------------------------------------
' This script is intended to be run on a Windows machine with SAP GUI installed.
' It automates SAP interactions for the close-up process.
' -----------------------------------------------------------------------

' Error handling
On Error Resume Next
Dim errScript
Set errScript = CreateObject("Scripting.Dictionary")

' Initialize SAP connection
Function ConnectToSAP()
    Dim SapGuiAuto, Application, Connection, Session
    
    ' Get SAP GUI Scripting object
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Application = SapGuiAuto.GetScriptingEngine
    
    ' Check if SAP GUI is running
    If IsObject(Application) Then
        ' Use existing connection if available
        If Application.Connections.Count > 0 Then
            Set Connection = Application.Connections(0)
            Set Session = Connection.Children(0)
            ConnectToSAP = True
            Set GetSapSession = Session
        Else
            MsgBox "No active SAP connections found. Please log in to SAP first.", 16, "Error"
            ConnectToSAP = False
        End If
    Else
        MsgBox "SAP GUI Automation is not available. Please make sure SAP GUI is running and scripting is enabled.", 16, "Error"
        ConnectToSAP = False
    End If
End Function

' Get SAP session
Function GetSapSession()
    Dim SapGuiAuto, Application, Connection
    
    ' Get SAP GUI Scripting object
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Application = SapGuiAuto.GetScriptingEngine
    
    ' Use existing connection
    Set Connection = Application.Connections(0)
    Set GetSapSession = Connection.Children(0)
End Function

' Check if error occurred
Function CheckError()
    If Err.Number <> 0 Then
        errScript.Add Err.Number, Err.Description
        CheckError = False
    Else
        CheckError = True
    End If
    Err.Clear
End Function

' Display errors if any
Sub ShowErrors()
    If errScript.Count > 0 Then
        Dim errorMsg
        errorMsg = "The following errors occurred:" & vbCrLf
        
        For Each errNum In errScript.Keys
            errorMsg = errorMsg & "Error " & errNum & ": " & errScript(errNum) & vbCrLf
        Next
        
        MsgBox errorMsg, 16, "Script Errors"
    End If
End Sub

"""
        if order_number:
            script += f"' Service Order Number\nDim ServiceOrderNumber\nServiceOrderNumber = \"{order_number}\"\n\n"
        
        return script
    
    def connect(self):
        """Generate script for connecting to SAP"""
        self.connected = True
        logger.info("Local SAP connection ready for script generation")
        return True
    
    def login(self, username=None, password=None):
        """Local mode doesn't need login, as it assumes SAP is already open"""
        self.logged_in = True
        logger.info("Local SAP connection ready (login is handled by local SAP GUI)")
        return True
    
    def get_service_order_details(self, order_number):
        """Generate script to get service order details"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
        
        logger.info(f"Generating script to get service order details for {order_number}")
        
        # Create script content
        script = self._generate_script_template(order_number)
        script += """
' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to service order display transaction
    Session.StartTransaction "IW33"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").CaretPosition = 12
    Session.FindById("wnd[0]").SendVKey 0 ' ENTER
    CheckError()
    
    ' Get order details
    Dim OrderDetails
    OrderDetails = "Order: " & ServiceOrderNumber & vbCrLf
    OrderDetails = OrderDetails & "Customer: " & Session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text & vbCrLf
    OrderDetails = OrderDetails & "Part Number: " & Session.FindById("wnd[0]/usr/tabsTABSTRIP_DETAIL/tabpINCL_DETAIL_TAB1/ssubDETAIL_1:SAPLITO0:0111/ctxtRIWO00-MATNR").Text & vbCrLf
    OrderDetails = OrderDetails & "Serial Number: " & Session.FindById("wnd[0]/usr/tabsTABSTRIP_DETAIL/tabpINCL_DETAIL_TAB1/ssubDETAIL_1:SAPLITO0:0111/ctxtRIWO00-SERIALNR").Text & vbCrLf
    
    ' Display order details
    MsgBox OrderDetails, 64, "Service Order Details"
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"get_order_{order_number}.vbs", script)
        
        # In a real scenario, we'd parse data from the SAP GUI
        # Here we're returning mock data for the web application to use
        # In a production version, you might implement a way to get real data back from the script
        customer_num = random.choice(SPEX_CUSTOMER_NUMBERS) if random.random() > 0.7 else f"CUST{random.randint(10000, 99999)}"
        
        result = {
            "order_number": order_number,
            "customer_number": customer_num,
            "is_spex": customer_num in SPEX_CUSTOMER_NUMBERS,
            "is_converted": random.random() > 0.8,
            "is_exchange": random.random() > 0.9,
            "part_number": f"PN-{random.randint(100000, 999999)}",
            "serial_number": f"SN-{random.randint(10000, 99999)}",
            "script_path": script_path,
            "script_name": os.path.basename(script_path)
        }
        
        return result
    
    def open_ziwbn(self, order_number):
        """Generate script to open ZIWBN transaction"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
            
        logger.info(f"Generating script to open ZIWBN for {order_number}")
        
        script = self._generate_script_template(order_number)
        script += """
' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to ZIWBN transaction
    Session.StartTransaction "ZIWBN"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtSO_AUFNR-LOW").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtSO_AUFNR-LOW").CaretPosition = 12
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute
    CheckError()
    
    ' Make sure we have results
    If Session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount > 0 Then
        ' Select the first row
        Session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectedRows = "0"
        Session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").DoubleClickCurrentCell
        CheckError()
        
        MsgBox "ZIWBN opened successfully for order " & ServiceOrderNumber, 64, "Success"
    Else
        MsgBox "No data found for order " & ServiceOrderNumber & " in ZIWBN", 16, "Error"
    End If
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"open_ziwbn_{order_number}.vbs", script)
        return True
    
    def labor_on(self, order_number):
        """Generate script for labor on process"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
            
        logger.info(f"Generating script for labor on process for {order_number}")
        
        script = self._generate_script_template(order_number)
        script += """
' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to service order transaction
    Session.StartTransaction "IW32"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").CaretPosition = 12
    Session.FindById("wnd[0]").SendVKey 0 ' ENTER
    CheckError()
    
    ' Go to labor screen
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0").Select
    CheckError()
    
    ' Enter labor data
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUBSCR_0100:SAPLCORU_S:0075/subSUBSCR_0100:SAPLCORU_S:0120/txtAFRUD-ISM01").Text = "1"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUBSCR_0100:SAPLCORU_S:0075/subSUBSCR_0100:SAPLCORU_S:0120/ctxtAFRUD-ISM02").Text = "CLK"
    
    ' Save the transaction
    Session.FindById("wnd[0]/tbar[0]/btn[11]").Press ' Save
    CheckError()
    
    MsgBox "Labor ON set successfully for order " & ServiceOrderNumber, 64, "Success"
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"labor_on_{order_number}.vbs", script)
        return True
    
    def labor_off(self, order_number):
        """Generate script for labor off process"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
            
        logger.info(f"Generating script for labor off process for {order_number}")
        
        script = self._generate_script_template(order_number)
        script += """
' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to service order transaction
    Session.StartTransaction "IW32"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").CaretPosition = 12
    Session.FindById("wnd[0]").SendVKey 0 ' ENTER
    CheckError()
    
    ' Go to labor screen
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0").Select
    CheckError()
    
    ' Enter labor data for labor off
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS0/ssubSUBSCR_0100:SAPLCORU_S:0075/subSUBSCR_0100:SAPLCORU_S:0120/txtAFRUD-ISM01").Text = "0"
    
    ' Save the transaction
    Session.FindById("wnd[0]/tbar[0]/btn[11]").Press ' Save
    CheckError()
    
    MsgBox "Labor OFF set successfully for order " & ServiceOrderNumber, 64, "Success"
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"labor_off_{order_number}.vbs", script)
        return True
    
    def update_findings(self, order_number, mods_in, mods_out):
        """Generate script to update findings"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
            
        logger.info(f"Generating script to update findings for {order_number}")
        
        # Clean up and escape string inputs for VBScript
        mods_in_escaped = mods_in.replace('"', '""')
        mods_out_escaped = mods_out.replace('"', '""')
        
        script = self._generate_script_template(order_number)
        script += f"""
' Findings data
Dim ModsIn, ModsOut
ModsIn = "{mods_in_escaped}"
ModsOut = "{mods_out_escaped}"

' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to service order transaction
    Session.StartTransaction "IW32"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").CaretPosition = 12
    Session.FindById("wnd[0]").SendVKey 0 ' ENTER
    CheckError()
    
    ' Go to findings tab
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS7").Select
    CheckError()
    
    ' Enter findings data
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS7/ssubSUBSCR_0100:SAPLCORU_S:0345/subF_SAPLITO0:SAPLITO0:1400/txtITOB-TDLINE[0,6]").Text = "MODS IN: " & ModsIn
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0100/tabpTCS7/ssubSUBSCR_0100:SAPLCORU_S:0345/subF_SAPLITO0:SAPLITO0:1400/txtITOB-TDLINE[1,6]").Text = "MODS OUT: " & ModsOut
    
    ' Save the transaction
    Session.FindById("wnd[0]/tbar[0]/btn[11]").Press ' Save
    CheckError()
    
    MsgBox "Findings updated successfully for order " & ServiceOrderNumber, 64, "Success"
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"update_findings_{order_number}.vbs", script)
        return True
    
    def update_wanding_status(self, order_number):
        """Generate script to update wanding status"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
            
        logger.info(f"Generating script to update wanding status for {order_number}")
        
        script = self._generate_script_template(order_number)
        script += """
' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to wanding transaction
    Session.StartTransaction "ZIWBN"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtSO_AUFNR-LOW").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtSO_AUFNR-LOW").CaretPosition = 12
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute
    CheckError()
    
    ' Select the first row
    If Session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount > 0 Then
        Session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectedRows = "0"
        ' Right-click to open context menu
        Session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressContextButton "WAND" ' Wanding button
        CheckError()
        
        ' Handle confirmation dialog if it appears
        If Session.ActiveWindow.Name = "wnd[1]" Then
            Session.FindById("wnd[1]/usr/btnBUTTON_1").Press ' Yes/Confirm
            CheckError()
        End If
        
        MsgBox "Wanding status updated successfully for order " & ServiceOrderNumber, 64, "Success"
    Else
        MsgBox "No data found for order " & ServiceOrderNumber & " in ZIWBN", 16, "Error"
    End If
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"update_wanding_{order_number}.vbs", script)
        return True
    
    def update_wsupd_comments(self, order_number, comments):
        """Generate script to update WSUPD comments"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
            
        logger.info(f"Generating script to update WSUPD comments for {order_number}")
        
        # Clean up and escape string inputs for VBScript
        comments_escaped = comments.replace('"', '""')
        
        script = self._generate_script_template(order_number)
        script += f"""
' Comments data
Dim Comments
Comments = "{comments_escaped}"

' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to WSUPD transaction
    Session.StartTransaction "WSUPD"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtAUFNR-LOW").CaretPosition = 12
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute
    CheckError()
    
    ' Check if we found the order
    If Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount > 0 Then
        ' Add comments
        Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
        Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").PressToolbarButton "COMMENTS" ' Comments button
        CheckError()
        
        ' Enter comment text
        Session.FindById("wnd[1]/usr/cntlLONGTEXT/shellcont/shell/shellcont[1]/shell").Text = Comments
        Session.FindById("wnd[1]/tbar[0]/btn[11]").Press ' Save
        CheckError()
        
        MsgBox "WSUPD comments updated successfully for order " & ServiceOrderNumber, 64, "Success"
    Else
        MsgBox "No data found for order " & ServiceOrderNumber & " in WSUPD", 16, "Error"
    End If
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"update_comments_{order_number}.vbs", script)
        return True
    
    def print_service_report(self, order_number):
        """Generate script to print service report"""
        if not self.logged_in:
            raise Exception("Not ready to generate scripts")
            
        logger.info(f"Generating script to print service report for {order_number}")
        
        script = self._generate_script_template(order_number)
        script += """
' Main script execution
If ConnectToSAP() Then
    Dim Session
    Set Session = GetSapSession()
    
    ' Navigate to service order transaction
    Session.StartTransaction "IW33"
    CheckError()
    
    ' Enter service order number and execute
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ServiceOrderNumber
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").CaretPosition = 12
    Session.FindById("wnd[0]").SendVKey 0 ' ENTER
    CheckError()
    
    ' Print the service report
    Session.FindById("wnd[0]/tbar[1]/btn[17]").Press ' Print button
    CheckError()
    
    ' Handle print dialog
    Session.FindById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCL"
    Session.FindById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").CaretPosition = 4
    Session.FindById("wnd[1]/tbar[0]/btn[13]").Press ' Continue
    CheckError()
    
    MsgBox "Service report printing initiated for order " & ServiceOrderNumber, 64, "Success"
End If

ShowErrors()
"""
        
        script_path = self._generate_script(f"print_report_{order_number}.vbs", script)
        return True


class ApiSapConnection:
    """SAP connection class that uses API to communicate with SAP GUI on Windows"""
    
    def __init__(self):
        self.connected = False
        self.logged_in = False
        self.api_url = config_module.SAP_API_URL
        self.api_key = os.environ.get("SAP_API_KEY", "")
        logger.info("API SAP Connection initialized")
    
    def _make_api_request(self, endpoint, method="GET", data=None):
        """Helper method to make API requests"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        url = f"{self.api_url}/{endpoint}"
        
        try:
            if method == "GET":
                response = requests.get(url, headers=headers, timeout=30)
            elif method == "POST":
                response = requests.post(url, headers=headers, json=data, timeout=30)
            elif method == "PUT":
                response = requests.put(url, headers=headers, json=data, timeout=30)
            else:
                logger.error(f"Unsupported method: {method}")
                return None
                
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.RequestException as e:
            logger.error(f"API request error: {str(e)}")
            return None
    
    def connect(self):
        """Connect to SAP via API"""
        logger.info("Connecting to SAP via API...")
        result = self._make_api_request("connect", method="POST")
        
        if result and result.get("connected"):
            self.connected = True
            return True
        
        return False
    
    def login(self, username=None, password=None):
        """Login to SAP via API"""
        if not self.connected:
            self.connect()
        
        logger.info(f"Logging in as {username} via API...")
        
        data = {
            "username": username,
            "password": password
        }
        
        result = self._make_api_request("login", method="POST", data=data)
        
        if result and result.get("logged_in"):
            self.logged_in = True
            return True
            
        return False
    
    def get_service_order_details(self, order_number):
        """Get service order details from SAP via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
        
        logger.info(f"Getting service order details for {order_number} via API")
        
        result = self._make_api_request(f"service_order/{order_number}")
        
        if not result:
            return None
            
        return result
    
    def open_ziwbn(self, order_number):
        """Open ZIWBN transaction via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Opening ZIWBN for {order_number} via API")
        
        data = {
            "order_number": order_number
        }
        
        result = self._make_api_request("open_ziwbn", method="POST", data=data)
        
        return result and result.get("success", False)
    
    def labor_on(self, order_number):
        """Set labor on via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Setting labor on for {order_number} via API")
        
        data = {
            "order_number": order_number
        }
        
        result = self._make_api_request("labor_on", method="POST", data=data)
        
        return result and result.get("success", False)
    
    def labor_off(self, order_number):
        """Set labor off via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Setting labor off for {order_number} via API")
        
        data = {
            "order_number": order_number
        }
        
        result = self._make_api_request("labor_off", method="POST", data=data)
        
        return result and result.get("success", False)
    
    def update_findings(self, order_number, mods_in, mods_out):
        """Update findings via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating findings for {order_number} via API")
        
        data = {
            "order_number": order_number,
            "mods_in": mods_in,
            "mods_out": mods_out
        }
        
        result = self._make_api_request("update_findings", method="POST", data=data)
        
        return result and result.get("success", False)
    
    def update_wanding_status(self, order_number):
        """Update wanding status via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating wanding status for {order_number} via API")
        
        data = {
            "order_number": order_number
        }
        
        result = self._make_api_request("update_wanding", method="POST", data=data)
        
        return result and result.get("success", False)
    
    def update_wsupd_comments(self, order_number, comments):
        """Update WSUPD comments via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Updating WSUPD comments for {order_number} via API")
        
        data = {
            "order_number": order_number,
            "comments": comments
        }
        
        result = self._make_api_request("update_comments", method="POST", data=data)
        
        return result and result.get("success", False)
    
    def print_service_report(self, order_number):
        """Print service report via API"""
        if not self.logged_in:
            raise Exception("Not logged in to SAP")
            
        logger.info(f"Printing service report for {order_number} via API")
        
        data = {
            "order_number": order_number
        }
        
        result = self._make_api_request("print_report", method="POST", data=data)
        
        return result and result.get("success", False)


# Factory function to create the appropriate connection based on config
def get_sap_connection():
    """Create appropriate SAP connection based on configuration"""
    # Access the SAP_CONNECTION_TYPE from config module
    connection_type = config_module.SAP_CONNECTION_TYPE
    
    if connection_type == "api":
        return ApiSapConnection()
    elif connection_type == "local":
        return LocalSapConnection()
    else:
        return MockSapConnection()

# For backward compatibility, keep the original class name
# but use it as a factory now
class SapConnection:
    @staticmethod
    def create():
        return get_sap_connection()
