/**
 * Bootstrap - Global config sync from Supabase warehouse
 * 
 * THIS IS THE SOURCE OF TRUTH for config sync.
 * Called from module-selector/selector.js at add-in startup.
 * Runs BEFORE any module loads.
 * 
 * IMPORTANT: Uses warehouseRequest() for all API calls.
 * Do not call fetch() directly for warehouse. Use warehouseRequest().
 */

import { hasExcelRuntime, saveConfigValue } from "./gateway.js";
import { 
    warehouseRequest, 
    columnMapperRequest, 
    forceRefreshAuth, 
    debugAuthDump,
    getWarehouseAuthContext 
} from "./warehouse.js";

// ============================================================================
// EMAIL VALIDATION
// ============================================================================

const SESSION_EMAIL_KEY = "pf_authorized_email";

/**
 * Get authorized email from session storage
 * @returns {string|null}
 */
function getSessionEmail() {
    try {
        return sessionStorage.getItem(SESSION_EMAIL_KEY);
    } catch (e) {
        console.warn("[Session] Could not read sessionStorage:", e);
        return null;
    }
}

/**
 * Store authorized email in session storage
 * @param {string} email
 */
function setSessionEmail(email) {
    try {
        sessionStorage.setItem(SESSION_EMAIL_KEY, email);
        console.log("[Session] Email stored for session");
    } catch (e) {
        console.warn("[Session] Could not write to sessionStorage:", e);
    }
}

/**
 * Clear authorized email from session storage
 */
function clearSessionEmail() {
    try {
        sessionStorage.removeItem(SESSION_EMAIL_KEY);
        console.log("[Session] Email cleared from session");
    } catch (e) {
        console.warn("[Session] Could not clear sessionStorage:", e);
    }
}

// =============================================================================
// EMAIL AUTHORIZATION SYSTEM
// =============================================================================

/**
 * Validate email with server
 * @param {string} email - Email address to validate
 * @returns {Promise<{authorized: boolean, reason: string, error: string|null}>}
 */
async function validateEmailWithServer(email) {
    console.log("[EmailAuth] Validating email with server...");
    
    try {
        const result = await columnMapperRequest("validate_email", {
            email: email.toLowerCase(),
            installation_key: INSTALLATION_KEY
        }, "email_validation");
        
        if (!result.ok) {
            console.error("[EmailAuth] Server validation failed:", result.error);
            return {
                authorized: false,
                reason: "server_error",
                error: result.error || "Server validation failed"
            };
        }
        
        if (!result.data || !result.data.success) {
            console.error("[EmailAuth] Validation unsuccessful");
            return {
                authorized: false,
                reason: "validation_failed",
                error: "Email validation failed"
            };
        }
        
        console.log("[EmailAuth] Server validation result:", result.data.authorized ? "✓" : "✗");
        console.log("[EmailAuth] Authorization reason:", result.data.reason);
        
        return {
            authorized: result.data.authorized,
            reason: result.data.reason || "unknown",
            error: result.data.authorized ? null : "Email not authorized"
        };
        
    } catch (error) {
        console.error("[EmailAuth] Error validating email:", error);
        return {
            authorized: false,
            reason: "error",
            error: error.message
        };
    }
}

// =============================================================================
// CONFIGURATION
// =============================================================================

const INSTALLATION_KEY = "pf_install_9f3c2b1a_20251212";

// Field mapping: API response key -> SS_PF_Config field name
const FIELD_MAP = {
    "installation_key": "SS_Installation_Key",
    "crm_company_id": "SS_Company_ID",
    "ss_company_name": "SS_Company_Name",
    "ss_accounting_software": "SS_Accounting_Software",
    "pto_payroll_provider": "PTO_Payroll_Provider",
    "pr_payroll_provider": "PR_Payroll_Provider"
};

// Required fields that MUST exist in response
const REQUIRED_FIELDS = ["crm_company_id", "ss_company_name"];

// Cache
let bootstrapCache = null;
let bootstrapRanAt = null;

// =============================================================================
// MAIN ENTRY POINT
// =============================================================================

/**
 * Run bootstrap config sync. Called at global add-in startup.
 * 
 * @param {Object} options
 * @param {boolean} options.force - Force refresh even if cached
 * @returns {Promise<{success: boolean, data?: object, error?: string}>}
 */
export async function bootstrapConfigSync(options = {}) {
    const { force = false, email = null } = options;
    
    console.log("╔═══════════════════════════════════════════════════════════╗");
    console.log("║  BOOTSTRAP CONFIG SYNC - GLOBAL ENTRYPOINT                ║");
    console.log("╚═══════════════════════════════════════════════════════════╝");
    console.log("[Bootstrap] Force:", force, "| Cached:", !!bootstrapCache);
    
    // Skip if already run (unless forced)
    if (bootstrapCache && !force) {
        console.log("[Bootstrap] Using cached result from", bootstrapRanAt);
        return { success: true, data: bootstrapCache, cached: true };
    }
    
    // ========================================================================
    // STAGE 0: EMAIL AUTHORIZATION (CRITICAL SECURITY GATE)
    // ========================================================================
    console.log("\n[Bootstrap] ▶ STAGE 0: EMAIL AUTHORIZATION");
    
    // Check session storage first
    const sessionEmail = getSessionEmail();
    let authorizedEmail = email;
    
    if (sessionEmail) {
        console.log("[Bootstrap] ✓ Using email from session storage");
        authorizedEmail = sessionEmail;
    } else if (!email) {
        console.log("[Bootstrap] ❌ No email provided and no session email");
        return {
            success: false,
            error: "Email authorization required",
            stage: "authorization",
            needsEmailPrompt: true
        };
    } else {
        // New email provided - validate it
        console.log("[Bootstrap] Validating new email...");
        const emailValidation = await validateEmailWithServer(email);
        
        if (!emailValidation.authorized) {
            console.log("[Bootstrap] ❌ Email not authorized:", emailValidation.reason);
            return {
                success: false,
                error: emailValidation.error || "Email not authorized",
                stage: "authorization",
                unauthorized: true,
                reason: emailValidation.reason
            };
        }
        
        console.log("[Bootstrap] ✓ Email authorized:", email, "| Reason:", emailValidation.reason);
        
        // Store in session for future use
        setSessionEmail(email);
    }
    
    console.log("[Bootstrap] ✓ Using authorized email:", authorizedEmail);
    // ========================================================================
    
    // ALWAYS refresh auth before bootstrap
    console.log("\n[Bootstrap] Refreshing auth context...");
    await forceRefreshAuth();
    debugAuthDump();
    
    // STAGE 1: FETCH
    console.log("\n[Bootstrap] ▶ STAGE 1: FETCH FROM WAREHOUSE");
    let data;
    try {
        data = await fetchFromWarehouse();
        console.log("[Bootstrap] ✓ Fetch complete:", data.ss_company_name);
    } catch (fetchError) {
        console.error("[Bootstrap] ❌ Fetch failed:", fetchError.message);
        return { success: false, error: fetchError.message, stage: "fetch" };
    }
    
    // STAGE 2: VALIDATE
    console.log("\n[Bootstrap] ▶ STAGE 2: VALIDATE SCHEMA");
    try {
        validateSchema(data);
        console.log("[Bootstrap] ✓ Schema valid");
    } catch (validationError) {
        console.error("[Bootstrap] ❌ Validation failed:", validationError.message);
        return { success: false, error: validationError.message, stage: "validate" };
    }
    
    // STAGE 3: WRITE TO SS_PF_Config
    console.log("\n[Bootstrap] ▶ STAGE 3: WRITE TO SS_PF_Config");
    if (!hasExcelRuntime()) {
        console.warn("[Bootstrap] ⚠️ Excel not available - skipping write");
        bootstrapCache = data;
        bootstrapRanAt = new Date().toISOString();
        return { success: true, data, written: false, reason: "no_excel" };
    }
    
    try {
        const writeResult = await writeConfigValues(data);
        console.log("[Bootstrap] ✓ Write complete:", writeResult);
        
        // Cache result
        bootstrapCache = data;
        bootstrapRanAt = new Date().toISOString();
        
        return { success: true, data, written: true, writeResult };
    } catch (writeError) {
        console.error("[Bootstrap] ❌ Write failed:", writeError.message);
        return { success: false, error: writeError.message, stage: "write" };
    }
}

// =============================================================================
// FETCH - Uses warehouseRequest (not direct fetch)
// =============================================================================

/**
 * Fetch bootstrap data from warehouse.
 * Do not call fetch() directly for warehouse. Use warehouseRequest().
 */
async function fetchFromWarehouse() {
    console.log("[Bootstrap:Fetch] Installation key:", INSTALLATION_KEY);
    
    // Use centralized warehouse request
    const result = await columnMapperRequest("bootstrap", {
        installation_key: INSTALLATION_KEY
    }, "bootstrap_fetch");
    
    if (!result.ok) {
        throw new Error(`API error: ${result.error}`);
    }
    
    const data = result.data;
    
    if (!data.success) {
        throw new Error(`API returned success=false: ${data.error}`);
    }
    
    // Add installation_key to data so it gets written to SS_PF_Config
    data.installation_key = INSTALLATION_KEY;
    
    return data;
}

// =============================================================================
// VALIDATE
// =============================================================================

function validateSchema(data) {
    const missing = REQUIRED_FIELDS.filter(f => !data[f]);
    if (missing.length > 0) {
        throw new Error(`Missing required fields: ${missing.join(", ")}`);
    }
}

// =============================================================================
// WRITE
// =============================================================================

async function writeConfigValues(data) {
    const results = { written: 0, skipped: 0, errors: [] };
    
    for (const [apiKey, configField] of Object.entries(FIELD_MAP)) {
        const value = data[apiKey];
        
        if (!value) {
            console.log(`[Bootstrap:Write] Skip ${configField} (no value)`);
            results.skipped++;
            continue;
        }
        
        try {
            console.log(`[Bootstrap:Write] Writing ${configField} = "${value}"`);
            const written = await saveConfigValue(configField, value);
            if (written) {
                results.written++;
            } else {
                results.skipped++;
            }
        } catch (err) {
            console.error(`[Bootstrap:Write] Error writing ${configField}:`, err);
            results.errors.push({ field: configField, error: err.message });
        }
    }
    
    return results;
}

// =============================================================================
// EXPORTS
// =============================================================================

export function getBootstrapCache() {
    return bootstrapCache;
}

export function clearBootstrapCache() {
    bootstrapCache = null;
    bootstrapRanAt = null;
}

// Re-export warehouse utilities for convenience
export { forceRefreshAuth, debugAuthDump, getWarehouseAuthContext };

// Expose globally for console debugging
if (typeof window !== "undefined") {
    window.bootstrapConfigSync = bootstrapConfigSync;
    window.clearBootstrapCache = clearBootstrapCache;
}
