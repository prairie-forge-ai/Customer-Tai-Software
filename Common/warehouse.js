/**
 * Warehouse API Client - SINGLE SOURCE OF TRUTH
 * 
 * ALL warehouse/edge function calls MUST go through warehouseRequest().
 * Do not call fetch() directly for warehouse. Use warehouseRequest().
 * 
 * This module provides:
 * - Centralized auth header management
 * - Preflight validation (throws BEFORE fetch if auth missing)
 * - Structured logging (no secrets exposed)
 * - Debug utilities for troubleshooting
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const WAREHOUSE_BASE_URL = "https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1";
const PROJECT_REF = "jgciqwzwacaesqjaoadc";

// Storage key - ONE constant used everywhere
const AUTH_STORAGE_KEY = "pf_warehouse_auth";

// The anon key - this is designed to be public/client-side safe
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE";

// =============================================================================
// AUTH STATE
// =============================================================================

// IMPORTANT: Initialize with built-in key synchronously to avoid race conditions.
// The async IIFE below can upgrade this from storage if available.
let cachedAuthToken = SUPABASE_ANON_KEY;
let tokenSource = "BUILTIN_ANON_KEY_SYNC";
let lastRefreshTime = new Date().toISOString();

/**
 * Force refresh auth from available sources.
 * Called by bootstrap before making requests.
 */
export async function forceRefreshAuth() {
    console.log("[Warehouse:Auth] Force refreshing auth...");
    
    // Priority 1: OfficeRuntime.storage (if available)
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
        try {
            const stored = await OfficeRuntime.storage.getItem(AUTH_STORAGE_KEY);
            if (stored) {
                cachedAuthToken = stored;
                tokenSource = "OfficeRuntime.storage";
                lastRefreshTime = new Date().toISOString();
                console.log("[Warehouse:Auth] ✓ Loaded from OfficeRuntime.storage");
                return;
            }
        } catch (e) {
            console.warn("[Warehouse:Auth] OfficeRuntime.storage read failed:", e.message);
        }
    }
    
    // Priority 2: Use the built-in anon key (always available)
    cachedAuthToken = SUPABASE_ANON_KEY;
    tokenSource = "BUILTIN_ANON_KEY";
    lastRefreshTime = new Date().toISOString();
    console.log("[Warehouse:Auth] ✓ Using built-in anon key");
}

/**
 * Store auth token persistently (if storage available)
 */
export async function storeAuthToken(token) {
    cachedAuthToken = token;
    
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
        try {
            await OfficeRuntime.storage.setItem(AUTH_STORAGE_KEY, token);
            tokenSource = "OfficeRuntime.storage";
            console.log("[Warehouse:Auth] ✓ Stored to OfficeRuntime.storage");
        } catch (e) {
            console.warn("[Warehouse:Auth] OfficeRuntime.storage write failed:", e.message);
            tokenSource = "MEMORY_ONLY";
        }
    } else {
        tokenSource = "MEMORY_ONLY";
    }
    
    lastRefreshTime = new Date().toISOString();
}

/**
 * Get current auth context (for debugging - no secrets exposed)
 */
export function getWarehouseAuthContext() {
    const hasToken = !!cachedAuthToken;
    const tokenPreview = hasToken ? `${cachedAuthToken.substring(0, 20)}...` : null;
    
    return {
        baseUrl: WAREHOUSE_BASE_URL,
        projectRef: PROJECT_REF,
        tokenSource: tokenSource,
        headers_present: hasToken ? ["Authorization", "apikey", "Content-Type"] : ["Content-Type"],
        isValid: hasToken && cachedAuthToken.length > 50,
        tokenPreview: tokenPreview,
        lastRefreshTime: lastRefreshTime,
        storageKeyUsed: AUTH_STORAGE_KEY
    };
}

/**
 * Debug dump - callable from console for troubleshooting
 */
export function debugAuthDump() {
    const ctx = getWarehouseAuthContext();
    
    console.log("╔═══════════════════════════════════════════════════════════╗");
    console.log("║  WAREHOUSE AUTH DEBUG DUMP                                ║");
    console.log("╚═══════════════════════════════════════════════════════════╝");
    console.log("[Debug] baseUrl:", ctx.baseUrl);
    console.log("[Debug] projectRef:", ctx.projectRef);
    console.log("[Debug] tokenSource:", ctx.tokenSource);
    console.log("[Debug] headers_present:", ctx.headers_present);
    console.log("[Debug] isValid:", ctx.isValid);
    console.log("[Debug] tokenPreview:", ctx.tokenPreview);
    console.log("[Debug] lastRefreshTime:", ctx.lastRefreshTime);
    console.log("[Debug] storageKeyUsed:", ctx.storageKeyUsed);
    console.log("[Debug] OfficeRuntime available:", typeof OfficeRuntime !== "undefined");
    
    return ctx;
}

// =============================================================================
// MAIN REQUEST FUNCTION
// =============================================================================

/**
 * Make a warehouse API request.
 * 
 * THIS IS THE ONLY ALLOWED WAY TO CALL WAREHOUSE/EDGE FUNCTIONS.
 * Do not call fetch() directly for warehouse. Use warehouseRequest().
 * 
 * @param {string} endpoint - The endpoint path (e.g., "column-mapper")
 * @param {object} body - Request body (will be JSON stringified)
 * @param {string} callsiteTag - Tag for logging (e.g., "bootstrap", "analyze")
 * @returns {Promise<{ok: boolean, status: number, data?: any, error?: string}>}
 */
export async function warehouseRequest(endpoint, body, callsiteTag = "unknown") {
    const url = `${WAREHOUSE_BASE_URL}/${endpoint}`;
    const ctx = getWarehouseAuthContext();
    
    // Log request (no secrets)
    console.log(`[DW] callsite=${callsiteTag} url=${url} tokenSource=${ctx.tokenSource} headers_present=[${ctx.headers_present.join(",")}]`);
    
    // PREFLIGHT INVARIANT: Auth must be valid
    if (!ctx.isValid) {
        const errorMsg = `[DW:PREFLIGHT] ❌ Authorization missing! callsite=${callsiteTag} tokenSource=${ctx.tokenSource}`;
        console.error(errorMsg);
        console.error("[DW:PREFLIGHT] Call forceRefreshAuth() or debugAuthDump() to troubleshoot");
        throw new Error(errorMsg);
    }
    
    // Build headers
    const headers = {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${cachedAuthToken}`,
        "apikey": cachedAuthToken
    };
    
    try {
        const response = await fetch(url, {
            method: "POST",
            headers: headers,
            body: JSON.stringify(body)
        });
        
        console.log(`[DW] callsite=${callsiteTag} status=${response.status}`);
        
        // Handle 401 specifically
        if (response.status === 401) {
            const errorBody = await response.text();
            console.error(`[DW:401] callsite=${callsiteTag} response_body=${errorBody}`);
            console.error(`[DW:401] tokenSource=${ctx.tokenSource} tokenPreview=${ctx.tokenPreview}`);
            return { ok: false, status: 401, error: `Unauthorized: ${errorBody}` };
        }
        
        if (!response.ok) {
            const errorBody = await response.text();
            console.error(`[DW:ERROR] callsite=${callsiteTag} status=${response.status} body=${errorBody}`);
            return { ok: false, status: response.status, error: errorBody };
        }
        
        const data = await response.json();
        console.log(`[DW] callsite=${callsiteTag} success=true`);
        return { ok: true, status: response.status, data };
        
    } catch (networkError) {
        console.error(`[DW:NETWORK] callsite=${callsiteTag} error=${networkError.message}`);
        return { ok: false, status: 0, error: networkError.message };
    }
}

// =============================================================================
// CONVENIENCE FUNCTIONS
// =============================================================================

/**
 * Call column-mapper endpoint with specific action
 */
export async function columnMapperRequest(action, payload, callsiteTag) {
    return warehouseRequest("column-mapper", { action, ...payload }, callsiteTag || action);
}

// =============================================================================
// INITIALIZATION
// =============================================================================

// Auto-initialize auth on module load
(async function initAuth() {
    await forceRefreshAuth();
})();

// =============================================================================
// GLOBAL EXPORTS FOR CONSOLE DEBUGGING
// =============================================================================

if (typeof window !== "undefined") {
    window.warehouseDebug = {
        debugAuthDump,
        forceRefreshAuth,
        getWarehouseAuthContext,
        testRequest: async () => {
            console.log("\n=== TEST: debugAuthDump ===");
            debugAuthDump();
            
            console.log("\n=== TEST: bootstrap ===");
            try {
                const result = await columnMapperRequest("bootstrap", { 
                    installation_key: "pf_install_9f3c2b1a_20251212" 
                }, "test_bootstrap");
                console.log("Bootstrap result:", result);
            } catch (e) {
                console.error("Bootstrap error:", e);
            }
            
            console.log("\n=== TEST: get_options ===");
            try {
                const result = await columnMapperRequest("get_options", { 
                    module: "payroll-recorder" 
                }, "test_get_options");
                console.log("Get options result:", result);
            } catch (e) {
                console.error("Get options error:", e);
            }
        }
    };
    
    console.log("[Warehouse] Debug utilities available: window.warehouseDebug.debugAuthDump(), .forceRefreshAuth(), .testRequest()");
}

