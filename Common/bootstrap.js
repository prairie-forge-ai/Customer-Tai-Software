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

// =============================================================================
// CONFIGURATION
// =============================================================================

const INSTALLATION_KEY = "pf_install_9f3c2b1a_20251212";

// Field mapping: API response key -> SS_PF_Config field name
const FIELD_MAP = {
    "installation_key": "SS_Installation_Key",
    "company_id": "SS_Company_ID",
    "ss_company_name": "SS_Company_Name",
    "ss_accounting_software": "SS_Accounting_Software",
    "pto_payroll_provider": "PTO_Payroll_Provider",
    "pr_payroll_provider": "PR_Payroll_Provider"
};

// Required fields that MUST exist in response
const REQUIRED_FIELDS = ["company_id", "ss_company_name"];

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
    const { force = false } = options;
    
    console.log("╔═══════════════════════════════════════════════════════════╗");
    console.log("║  BOOTSTRAP CONFIG SYNC - GLOBAL ENTRYPOINT                ║");
    console.log("╚═══════════════════════════════════════════════════════════╝");
    console.log("[Bootstrap] Force:", force, "| Cached:", !!bootstrapCache);
    
    // Skip if already run (unless forced)
    if (bootstrapCache && !force) {
        console.log("[Bootstrap] Using cached result from", bootstrapRanAt);
        return { success: true, data: bootstrapCache, cached: true };
    }
    
    // ALWAYS refresh auth before bootstrap
    console.log("[Bootstrap] Refreshing auth context...");
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
