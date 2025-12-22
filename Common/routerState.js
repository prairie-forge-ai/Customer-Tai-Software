/**
 * Router State Persistence
 * 
 * Saves and restores the current view/step state to localStorage so that
 * the taskpane returns to the same screen after reload or coming back from
 * clicking elsewhere in Excel.
 * 
 * @module routerState
 */

// Storage key includes module prefix to avoid conflicts
const STORAGE_KEY_PREFIX = "pf.routeState.v1";

/**
 * Route state object
 * @typedef {Object} RouteState
 * @property {string} route - Route identifier (e.g., "payroll/home", "payroll/step/2")
 * @property {Object} params - Optional parameters (activeStepId, focusedIndex, etc.)
 * @property {number} ts - Timestamp for expiration checks
 */

// Max age before route expires (12 hours)
const MAX_AGE_MS = 1000 * 60 * 60 * 12;

/**
 * Get the storage key for a module
 * @param {string} moduleKey - Module identifier (e.g., "payroll-recorder")
 * @returns {string}
 */
function getStorageKey(moduleKey) {
    return `${STORAGE_KEY_PREFIX}.${moduleKey}`;
}

/**
 * Save the current route state to localStorage
 * 
 * @param {string} moduleKey - Module identifier
 * @param {string} route - Route identifier
 * @param {Object} params - Optional parameters
 */
export function saveRouteState(moduleKey, route, params = {}) {
    const payload = {
        route,
        params,
        ts: Date.now()
    };
    
    try {
        localStorage.setItem(getStorageKey(moduleKey), JSON.stringify(payload));
        console.debug(`[RouterState] Saved: ${route}`, params);
    } catch (e) {
        // Non-fatal: storage could be blocked or full
        console.warn("[RouterState] Failed to save route state", e);
    }
}

/**
 * Load the saved route state from localStorage
 * 
 * @param {string} moduleKey - Module identifier
 * @returns {RouteState|null} - Saved route state or null if not found/expired
 */
export function loadRouteState(moduleKey) {
    try {
        const raw = localStorage.getItem(getStorageKey(moduleKey));
        if (!raw) return null;
        
        const parsed = JSON.parse(raw);
        if (!parsed?.route) return null;
        
        // Check if route is expired
        if (Date.now() - parsed.ts > MAX_AGE_MS) {
            console.debug("[RouterState] Route expired, clearing");
            clearRouteState(moduleKey);
            return null;
        }
        
        console.debug(`[RouterState] Restored: ${parsed.route}`, parsed.params);
        return parsed;
    } catch (e) {
        console.warn("[RouterState] Failed to load route state", e);
        return null;
    }
}

/**
 * Clear the saved route state
 * 
 * @param {string} moduleKey - Module identifier
 */
export function clearRouteState(moduleKey) {
    try {
        localStorage.removeItem(getStorageKey(moduleKey));
        console.debug("[RouterState] Cleared");
    } catch (e) {
        console.warn("[RouterState] Failed to clear route state", e);
    }
}

/**
 * Build a route string from view and step info
 * 
 * @param {string} moduleKey - Module identifier (e.g., "payroll-recorder")
 * @param {string} activeView - Current view ("home", "config", "step")
 * @param {number|null} activeStepId - Current step ID (for step view)
 * @returns {string} - Route string (e.g., "payroll/home", "payroll/step/2")
 */
export function buildRouteString(moduleKey, activeView, activeStepId) {
    const prefix = moduleKey.replace("-recorder", "").replace("-", "/");
    
    switch (activeView) {
        case "config":
            return `${prefix}/config`;
        case "step":
            return `${prefix}/step/${activeStepId}`;
        case "home":
        default:
            return `${prefix}/home`;
    }
}

/**
 * Parse a route string into view and step info
 * 
 * @param {string} route - Route string
 * @returns {Object} - { activeView, activeStepId }
 */
export function parseRouteString(route) {
    if (!route) return { activeView: "home", activeStepId: null };
    
    const parts = route.split("/");
    const viewPart = parts[1] || "home";
    
    switch (viewPart) {
        case "config":
            return { activeView: "config", activeStepId: 0 };
        case "step": {
            const stepId = parseInt(parts[2], 10);
            return { 
                activeView: "step", 
                activeStepId: isNaN(stepId) ? null : stepId 
            };
        }
        case "home":
        default:
            return { activeView: "home", activeStepId: null };
    }
}

