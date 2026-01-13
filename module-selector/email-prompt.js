/**
 * Email Prompt Dialog Logic
 * Handles email input and validation
 */

// Wait for Office.js to be ready
Office.initialize = function() {
    console.log("[EmailPrompt] Office initialized");
    
    // Get DOM elements
    const emailInput = document.getElementById('emailInput');
    const errorMessage = document.getElementById('errorMessage');
    const verifyButton = document.getElementById('verifyButton');
    const cancelButton = document.getElementById('cancelButton');
    
    // Focus the input
    emailInput.focus();
    
    // Handle Enter key
    emailInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            handleVerify();
        }
    });
    
    // Handle Escape key
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            handleCancel();
        }
    });
    
    // Handle verify button
    verifyButton.addEventListener('click', handleVerify);
    
    // Handle cancel button
    cancelButton.addEventListener('click', handleCancel);
    
    // Hide error on input
    emailInput.addEventListener('input', function() {
        errorMessage.classList.remove('show');
    });
    
    /**
     * Handle verify attempt
     */
    function handleVerify() {
        const email = emailInput.value.trim().toLowerCase();
        
        // Basic validation
        if (!email) {
            showError('Please enter an email address.');
            return;
        }
        
        // Email format validation
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
            showError('Please enter a valid email address.');
            return;
        }
        
        // Hide error
        errorMessage.classList.remove('show');
        
        // Disable button and show loading
        verifyButton.disabled = true;
        verifyButton.innerHTML = '<span class="loading"></span>Verifying...';
        
        // Send email back to parent
        setTimeout(() => {
            try {
                Office.context.ui.messageParent(JSON.stringify({
                    success: true,
                    email: email
                }));
            } catch (error) {
                console.error("[EmailPrompt] Error sending message:", error);
                showError('Failed to verify. Please try again.');
                verifyButton.disabled = false;
                verifyButton.innerHTML = 'Verify Access';
            }
        }, 300); // Small delay for UX
    }
    
    /**
     * Handle cancel
     */
    function handleCancel() {
        try {
            Office.context.ui.messageParent(JSON.stringify({
                success: false,
                email: null,
                cancelled: true
            }));
        } catch (error) {
            console.error("[EmailPrompt] Error sending cancel message:", error);
        }
    }
    
    /**
     * Show error message
     */
    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.add('show');
        emailInput.focus();
    }
};
