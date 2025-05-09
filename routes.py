import logging
from flask import render_template, request, redirect, url_for, session, flash, jsonify, send_file, abort
from app import app
from models import ServiceOrder
from sap_connections import SapConnection
from config import WORKFLOW_STEPS, SPEX_CUSTOMER_NUMBERS, SAP_SCRIPT_DIR
import os

logger = logging.getLogger(__name__)
# Create a SAP connection instance using the factory method
sap = SapConnection.create()

# Ensure SAP connection is active
@app.before_request
def before_request():
    if not sap.connected:
        sap.connect()
    if not sap.logged_in:
        sap.login()

# Helper to get current service order from session
def get_current_service_order():
    if 'service_order' not in session:
        return None
    
    order_data = session['service_order']
    order = ServiceOrder(order_data['order_number'])
    for key, value in order_data.items():
        if hasattr(order, key):
            setattr(order, key, value)
    return order

# Helper to save service order to session
def save_service_order(order):
    session['service_order'] = order.__dict__

# Route for home page
@app.route('/')
def index():
    # Clear any existing service order
    if 'service_order' in session:
        session.pop('service_order')
    return render_template('index.html')

# Route for configuration page
@app.route('/config', methods=['GET', 'POST'])
def config():
    from config import SAP_CONNECTION_TYPE, SAP_API_URL
    import os
    
    # Get current configuration
    current_config = {
        'connection_type': SAP_CONNECTION_TYPE,
        'api_url': SAP_API_URL,
        'api_key': os.environ.get('SAP_API_KEY', '')
    }
    
    # Handle form submission
    if request.method == 'POST':
        connection_type = request.form.get('connection_type')
        api_url = request.form.get('api_url')
        api_key = request.form.get('api_key')
        
        # Store in environment variables (this is temporary until app restart)
        os.environ['SAP_CONNECTION_TYPE'] = connection_type
        
        if connection_type == 'api':
            os.environ['SAP_API_URL'] = api_url
            if api_key:
                os.environ['SAP_API_KEY'] = api_key
        
        # This is a simplified approach for demonstration
        # In a production environment, we would store these in a config file or database
        
        # Show success message
        flash('Configuration saved successfully. Connection mode is now: ' + connection_type, 'success')
        
        # Update current config for display
        current_config = {
            'connection_type': connection_type,
            'api_url': api_url,
            'api_key': api_key
        }
        
        # Recreate SAP connection with new settings
        global sap
        sap = SapConnection.create()
        
    return render_template('config.html', 
                           current_config=current_config,
                           config={'SAP_CONNECTION_TYPE': SAP_CONNECTION_TYPE})

# Main workflow route
@app.route('/workflow/<step_id>', methods=['GET', 'POST'])
def workflow(step_id):
    # Get the current step information
    current_step = next((step for step in WORKFLOW_STEPS if step["id"] == step_id), None)
    if not current_step:
        flash("Invalid workflow step", "danger")
        return redirect(url_for('index'))
    
    # Get the current service order or redirect to start
    service_order = get_current_service_order()
    if not service_order and step_id != "service_order":
        flash("Please start with a service order", "warning")
        return redirect(url_for('workflow', step_id='service_order'))
    
    # Create a new service order if we're at the first step
    if step_id == "service_order" and request.method == "POST":
        order_number = request.form.get('service_order_number', '').strip()
        if not order_number:
            flash("Please enter a service order number", "warning")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS)
        
        # Create new service order and get details from SAP
        service_order = ServiceOrder(order_number)
        
        try:
            # Get service order details from SAP
            details = sap.get_service_order_details(order_number)
            if not details:
                flash("Service order not found in SAP", "danger")
                return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS)
            
            # Update service order with details from SAP
            for key, value in details.items():
                if hasattr(service_order, key):
                    setattr(service_order, key, value)
            
            # Open ZIWBN transaction
            sap.open_ziwbn(order_number)
            
            # Mark step as complete
            service_order.mark_step_complete(step_id)
            service_order.workflow_state = "equipment_check"
            save_service_order(service_order)
            
            # Move to next step
            return redirect(url_for('workflow', step_id='equipment_check'))
            
        except Exception as e:
            logger.error(f"Error getting service order details: {e}")
            flash(f"Error: {str(e)}", "danger")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS)
    
    # Handle equipment verification step
    elif step_id == "equipment_check" and request.method == "POST":
        pn_match = request.form.get('pn_match') == 'yes'
        sn_match = request.form.get('sn_match') == 'yes'
        
        if not pn_match or not sn_match:
            flash("Part or serial number doesn't match. Process cannot continue.", "danger")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
        
        service_order.mark_step_complete(step_id)
        service_order.workflow_state = "labor_on"
        save_service_order(service_order)
        return redirect(url_for('workflow', step_id='labor_on'))
    
    # Handle labor on step
    elif step_id == "labor_on" and request.method == "POST":
        try:
            # Set labor on in SAP
            sap.labor_on(service_order.order_number)
            
            service_order.mark_step_complete(step_id)
            service_order.workflow_state = "paperwork"
            save_service_order(service_order)
            return redirect(url_for('workflow', step_id='paperwork'))
        except Exception as e:
            logger.error(f"Error setting labor on: {e}")
            flash(f"Error: {str(e)}", "danger")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
    
    # Handle paperwork verification step
    elif step_id == "paperwork" and request.method == "POST":
        paperwork_complete = request.form.get('paperwork_complete') == 'yes'
        hardware_complete = request.form.get('hardware_complete') == 'yes'
        
        if not paperwork_complete or not hardware_complete:
            flash("Paperwork or hardware is not complete. Please ensure all requirements are met.", "warning")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
        
        # Determine next step based on whether it's a SPEX order
        service_order.mark_step_complete(step_id)
        if service_order.is_spex:
            service_order.workflow_state = "test_sheet"
            next_step = "test_sheet"
        else:
            service_order.workflow_state = "customer_req"
            next_step = "customer_req"
        
        save_service_order(service_order)
        return redirect(url_for('workflow', step_id=next_step))
    
    # Handle test sheet verification
    elif step_id == "test_sheet" and request.method == "POST":
        test_matches = request.form.get('test_matches') == 'yes'
        test_passes = request.form.get('test_passes') == 'yes'
        test_signed = request.form.get('test_signed') == 'yes'
        comments_addressed = request.form.get('comments_addressed') == 'yes'
        
        if not all([test_matches, test_passes, test_signed, comments_addressed]):
            flash("All test sheet requirements must be met before proceeding.", "warning")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
        
        service_order.mark_step_complete(step_id)
        
        # If SPEX, go to findings, otherwise continue normal flow
        if service_order.is_spex:
            service_order.workflow_state = "findings"
            next_step = "findings"
        else:
            service_order.workflow_state = "customer_req"
            next_step = "customer_req"
        
        save_service_order(service_order)
        return redirect(url_for('workflow', step_id=next_step))
    
    # Handle customer requirements
    elif step_id == "customer_req" and request.method == "POST":
        requirements_met = request.form.get('requirements_met') == 'yes'
        
        if not requirements_met:
            flash("Customer requirements must be addressed before proceeding.", "warning")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
        
        service_order.mark_step_complete(step_id)
        service_order.workflow_state = "notifications"
        save_service_order(service_order)
        return redirect(url_for('workflow', step_id='notifications'))
    
    # Handle notifications
    elif step_id == "notifications" and request.method == "POST":
        notifications_handled = request.form.get('notifications_handled') == 'yes'
        
        if not notifications_handled:
            flash("All notifications must be handled before proceeding.", "warning")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
        
        service_order.mark_step_complete(step_id)
        service_order.workflow_state = "warranty"
        save_service_order(service_order)
        return redirect(url_for('workflow', step_id='warranty'))
    
    # Handle warranty
    elif step_id == "warranty" and request.method == "POST":
        warranty_checked = request.form.get('warranty_checked') == 'yes'
        
        if not warranty_checked:
            flash("Warranty must be checked before proceeding.", "warning")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
        
        service_order.mark_step_complete(step_id)
        service_order.workflow_state = "findings"
        save_service_order(service_order)
        return redirect(url_for('workflow', step_id='findings'))
    
    # Handle findings update
    elif step_id == "findings" and request.method == "POST":
        mods_in = request.form.get('mods_in', '')
        mods_out = request.form.get('mods_out', '')
        
        try:
            # Update findings in SAP
            sap.update_findings(service_order.order_number, mods_in, mods_out)
            
            # Update service order
            service_order.mods_in = mods_in
            service_order.mods_out = mods_out
            
            # Update wanding status
            sap.update_wanding_status(service_order.order_number)
            
            # Update WSUPD comments
            comments = request.form.get('comments', '')
            sap.update_wsupd_comments(service_order.order_number, comments)
            
            service_order.mark_step_complete(step_id)
            service_order.workflow_state = "labor_off"
            save_service_order(service_order)
            return redirect(url_for('workflow', step_id='labor_off'))
        except Exception as e:
            logger.error(f"Error updating findings: {e}")
            flash(f"Error: {str(e)}", "danger")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
    
    # Handle labor off
    elif step_id == "labor_off" and request.method == "POST":
        try:
            # Set labor off in SAP
            sap.labor_off(service_order.order_number)
            
            service_order.mark_step_complete(step_id)
            service_order.workflow_state = "final_inspection"
            save_service_order(service_order)
            return redirect(url_for('workflow', step_id='final_inspection'))
        except Exception as e:
            logger.error(f"Error setting labor off: {e}")
            flash(f"Error: {str(e)}", "danger")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
    
    # Handle final inspection
    elif step_id == "final_inspection" and request.method == "POST":
        inspection_complete = request.form.get('inspection_complete') == 'yes'
        
        if not inspection_complete:
            flash("Final inspection must be completed before proceeding.", "warning")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
        
        try:
            # Print service report
            sap.print_service_report(service_order.order_number)
            
            service_order.mark_step_complete(step_id)
            service_order.workflow_state = "complete"
            save_service_order(service_order)
            return redirect(url_for('workflow', step_id='complete'))
        except Exception as e:
            logger.error(f"Error in final inspection: {e}")
            flash(f"Error: {str(e)}", "danger")
            return render_template('workflow.html', step=current_step, all_steps=WORKFLOW_STEPS, service_order=service_order)
    
    # Handle complete step
    elif step_id == "complete" and request.method == "POST":
        # Clear service order from session
        session.pop('service_order', None)
        return redirect(url_for('index'))
    
    # GET request - display the current step
    from config import SAP_CONNECTION_TYPE
    return render_template('workflow.html', 
                          step=current_step, 
                          all_steps=WORKFLOW_STEPS, 
                          service_order=service_order,
                          config={'SAP_CONNECTION_TYPE': SAP_CONNECTION_TYPE})

# Reset the workflow and start over
@app.route('/reset', methods=['POST'])
def reset_workflow():
    if 'service_order' in session:
        session.pop('service_order')
    return redirect(url_for('index'))

# Error handlers
@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

@app.errorhandler(500)
def server_error(e):
    logger.error(f"Server error: {e}")
    return render_template('500.html'), 500

# Route to download generated SAP scripts
@app.route('/scripts/<filename>')
def download_script(filename):
    from config import SAP_CONNECTION_TYPE, SAP_SCRIPT_DIR
    
    # Only allow downloads in local connection mode
    if SAP_CONNECTION_TYPE != 'local':
        flash("Script downloads are only available in Local SAP Connection mode", "warning")
        return redirect(url_for('config'))
    
    # Verify the script file exists
    script_path = os.path.join(SAP_SCRIPT_DIR, filename)
    if not os.path.exists(script_path):
        abort(404)
    
    try:
        # Send the file for download
        return send_file(script_path, as_attachment=True)
    except Exception as e:
        logger.error(f"Error downloading script: {e}")
        flash(f"Error downloading script: {str(e)}", "danger")
        return redirect(url_for('index'))
