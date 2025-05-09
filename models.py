# This file contains data models for the application
# In this case, we're using session-based data storage rather than a database
# since we're primarily interfacing with SAP

class ServiceOrder:
    """Model representing a service order with all its related data"""
    def __init__(self, order_number):
        self.order_number = order_number
        self.customer_number = None
        self.part_number = None
        self.serial_number = None
        self.is_spex = False
        self.is_converted = False
        self.is_exchange = False
        self.is_dpmi = False
        self.repair_level = None
        self.delivery_block = None
        self.zh_status = None
        self.zg_status = None
        self.mods_in = None
        self.mods_out = None
        self.sf_mods_in = None
        self.sf_mods_out = None
        self.sw_versions_in = None
        self.sw_versions_out = None
        self.workflow_state = "service_order"  # Starting state
        self.completed_steps = []
        
        # For local SAP GUI script generation
        self.script_name = None
        self.script_path = None

    def mark_step_complete(self, step_name):
        """Mark a workflow step as complete"""
        if step_name not in self.completed_steps:
            self.completed_steps.append(step_name)

    def is_step_complete(self, step_name):
        """Check if a workflow step is complete"""
        return step_name in self.completed_steps
