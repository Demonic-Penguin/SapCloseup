"""
This module handles SAP connections and interactions
Provides both a mock interface for development and a real SAP interface via API
"""
import logging
import time
import random
import os
import requests
from config import SPEX_CUSTOMER_NUMBERS, SAP_API_URL, SAP_CONNECTION_TYPE

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


class ApiSapConnection:
    """SAP connection class that uses API to communicate with SAP GUI on Windows"""
    
    def __init__(self):
        self.connected = False
        self.logged_in = False
        self.api_url = SAP_API_URL
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
    if SAP_CONNECTION_TYPE == "api":
        return ApiSapConnection()
    else:
        return MockSapConnection()

# For backward compatibility, keep the original class name
# but use it as a factory now
class SapConnection:
    @staticmethod
    def create():
        return get_sap_connection()
