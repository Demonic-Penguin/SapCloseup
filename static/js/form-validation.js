/**
 * Form validation script for SAP Close-Up Automation Tool
 */

document.addEventListener('DOMContentLoaded', function() {
    // Fetch all the forms we want to apply custom Bootstrap validation styles to
    const forms = document.querySelectorAll('.needs-validation');
  
    // Loop over them and prevent submission
    Array.from(forms).forEach(form => {
        form.addEventListener('submit', event => {
            if (!form.checkValidity()) {
                event.preventDefault();
                event.stopPropagation();
            }
        
            form.classList.add('was-validated');
        }, false);
    });
    
    // Service order form validation
    const serviceOrderForm = document.getElementById('serviceOrderForm');
    if (serviceOrderForm) {
        const serviceOrderInput = document.getElementById('service_order_number');
        
        serviceOrderInput.addEventListener('input', function() {
            // Remove any non-digit characters
            this.value = this.value.replace(/\D/g, '');
            
            // Check validity
            if (this.value.length > 0) {
                this.setCustomValidity('');
            } else {
                this.setCustomValidity('Please enter a valid service order number');
            }
        });
    }
    
    // Equipment check form validation
    const equipmentCheckForm = document.getElementById('equipmentCheckForm');
    if (equipmentCheckForm) {
        const pnMatchYes = document.getElementById('pn_match_yes');
        const pnMatchNo = document.getElementById('pn_match_no');
        const snMatchYes = document.getElementById('sn_match_yes');
        const snMatchNo = document.getElementById('sn_match_no');
        
        equipmentCheckForm.addEventListener('submit', function(event) {
            if (pnMatchNo.checked || snMatchNo.checked) {
                // Show alert
                alert('Part or serial number doesn\'t match. Process cannot continue. Please contact your supervisor.');
                event.preventDefault();
                event.stopPropagation();
            }
        });
    }
    
    // Test sheet form validation
    const testSheetForm = document.getElementById('testSheetForm');
    if (testSheetForm) {
        testSheetForm.addEventListener('submit', function(event) {
            const testMatchesNo = document.getElementById('test_matches_no');
            const testPassesNo = document.getElementById('test_passes_no');
            const testSignedNo = document.getElementById('test_signed_no');
            const commentsAddressedNo = document.getElementById('comments_addressed_no');
            
            if (testMatchesNo.checked || testPassesNo.checked || testSignedNo.checked || commentsAddressedNo.checked) {
                if (!confirm('One or more test sheet requirements are not met. Are you sure you want to proceed?')) {
                    event.preventDefault();
                    event.stopPropagation();
                }
            }
        });
    }
    
    // Findings form validation
    const findingsForm = document.getElementById('findingsForm');
    if (findingsForm) {
        const modsOut = document.getElementById('mods_out');
        
        findingsForm.addEventListener('submit', function(event) {
            if (!modsOut.value.trim()) {
                modsOut.setCustomValidity('Outgoing modifications are required');
            } else {
                modsOut.setCustomValidity('');
            }
        });
        
        modsOut.addEventListener('input', function() {
            if (this.value.trim()) {
                this.setCustomValidity('');
            } else {
                this.setCustomValidity('Outgoing modifications are required');
            }
        });
    }
});
