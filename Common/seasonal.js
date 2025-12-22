/**
 * Prairie Forge Seasonal Effects - Database-Driven Auto-injection Script
 * 
 * Fetches seasonal effects from the centralized Supabase database,
 * allowing real-time control across all platforms (website + Excel add-ins).
 * 
 * USAGE:
 * <script src="../Common/seasonal.js"></script>
 * 
 * TO DISABLE:
 * - Add class "pf-no-holiday" to body or .pf-root
 * - Or set window.PF_DISABLE_SEASONAL = true before loading
 * 
 * ============================================================================
 * IMPORTANT: SUPABASE PROJECT CONFIGURATION
 * ============================================================================
 * This customer has ONE AND ONLY ONE Supabase project:
 *   Project ID: jgciqwzwacaesqjaoadc
 *   URL: https://jgciqwzwacaesqjaoadc.supabase.co
 * 
 * DO NOT use any other project ID. There is no separate "Prairie Forge" or
 * "centralized" Supabase project. All modules (module-selector, payroll-recorder,
 * pto-accrual) connect to the SAME project above.
 * 
 * See also: Common/supabase-config.js for the canonical configuration.
 * ============================================================================
 */

(function() {
    'use strict';

    // Supabase configuration - MUST match Common/supabase-config.js
    // Project ID: jgciqwzwacaesqjaoadc (the ONLY Supabase project for this customer)
    const SUPABASE_URL = 'https://jgciqwzwacaesqjaoadc.supabase.co';
    const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE';
    // Platform identifier - database uses 'excel' (not 'excel-addin')
    const PLATFORM = 'excel';

    // Check if disabled
    if (window.PF_DISABLE_SEASONAL) {
        console.log('[Seasonal] Effects disabled via PF_DISABLE_SEASONAL');
        return;
    }

    /**
     * Check if today falls within a seasonal date range
     * Database stores start_month, start_day, end_month, end_day (not full dates)
     */
    function isDateInRange(feature) {
        const now = new Date();
        const currentMonth = now.getMonth() + 1; // 1-12
        const currentDay = now.getDate();

        const startMonth = feature.start_month;
        const startDay = feature.start_day;
        const endMonth = feature.end_month;
        const endDay = feature.end_day;

        // Handle same-year ranges (e.g., Dec 1 - Dec 31)
        if (startMonth <= endMonth) {
            if (currentMonth < startMonth || currentMonth > endMonth) return false;
            if (currentMonth === startMonth && currentDay < startDay) return false;
            if (currentMonth === endMonth && currentDay > endDay) return false;
            return true;
        }
        
        // Handle year-wrap ranges (e.g., Dec 15 - Jan 5)
        if (currentMonth > startMonth || (currentMonth === startMonth && currentDay >= startDay)) {
            return true;
        }
        if (currentMonth < endMonth || (currentMonth === endMonth && currentDay <= endDay)) {
            return true;
        }
        return false;
    }

    /**
     * Check if a feature is enabled for this platform
     */
    function isPlatformEnabled(feature) {
        if (!feature.platforms || !Array.isArray(feature.platforms)) {
            return true; // No platform restriction
        }
        return feature.platforms.includes(PLATFORM);
    }

    // Fetch active seasonal features from Supabase
    async function fetchSeasonalFeatures() {
        // Query enabled features, then filter by platform and date client-side
        // (PostgREST array containment syntax is tricky, so we filter client-side)
        const url = `${SUPABASE_URL}/rest/v1/seasonal_features?` + 
            `enabled=eq.true` +
            `&order=priority.desc`;

        const response = await fetch(url, {
            headers: {
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const features = await response.json();
        
        // Filter by platform and current date (month/day based)
        return features.filter(f => isPlatformEnabled(f) && isDateInRange(f));
    }

    // Inject seasonal elements based on active features
    function injectSeasonalEffects(features) {
        if (!features || features.length === 0) {
            console.log('[Seasonal] No active seasonal features');
            return;
        }

        const inject = () => {
            const container = 
                document.querySelector('.pf-root') || 
                document.querySelector('.app-shell') ||
                document.querySelector('.pto-shell') ||
                document.body;

            if (!container) {
                console.warn('[Seasonal] No container found');
                return;
            }

            // Set season data attribute
            const hasWinter = features.some(f => 
                f.feature_key === 'winter_snow' || 
                f.feature_key === 'snow' || 
                f.feature_key === 'christmas_lights'
            );
            if (hasWinter) {
                document.body.dataset.season = 'winter';
            }

            // Map database feature keys to CSS class names
            // CSS expects: .pf-holiday-snow, .pf-holiday-lights
            const CSS_CLASS_MAP = {
                'snow': 'pf-holiday-snow',
                'christmas_lights': 'pf-holiday-lights',
                'winter_snow': 'pf-holiday-snow',
                'valentines_hearts': 'pf-valentines-hearts',
                'st_patricks': 'pf-st-patricks',
                'fall_leaves': 'pf-fall-leaves'
            };

            // Inject each feature
            features.forEach(feature => {
                const cssClass = CSS_CLASS_MAP[feature.feature_key] || 
                                 feature.css_class || 
                                 `pf-${feature.feature_key.replace(/_/g, '-')}`;
                
                // Remove existing if present to avoid duplicates
                const existing = container.querySelector(`.${cssClass}`);
                if (existing) {
                    existing.remove();
                }

                const element = document.createElement('div');
                element.className = cssClass;
                element.setAttribute('aria-hidden', 'true');
                container.insertBefore(element, container.firstChild);
                
                console.log(`[Seasonal] ✨ Injected: ${feature.feature_key} as .${cssClass}`);
            });
        };

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', inject);
        } else {
            inject();
        }

        // Re-inject for async SPAs and Excel add-ins
        setTimeout(inject, 500);
        setTimeout(inject, 1500);
        setTimeout(inject, 3000);
        
        // Set up MutationObserver to handle Excel add-in DOM changes
        const observer = new MutationObserver((mutations) => {
            let shouldReinject = false;
            mutations.forEach((mutation) => {
                if (mutation.type === 'childList') {
                    // Check if pf-root was modified (common in Excel add-ins)
                    mutation.addedNodes.forEach((node) => {
                        if (node.nodeType === Node.ELEMENT_NODE) {
                            if (node.classList?.contains('pf-root') || 
                                node.querySelector?.('.pf-root')) {
                                shouldReinject = true;
                            }
                        }
                    });
                }
            });
            
            if (shouldReinject) {
                setTimeout(inject, 100);
            }
        });
        
        // Observe body for changes
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }

    // Initialize
    fetchSeasonalFeatures()
        .then(injectSeasonalEffects)
        .catch(err => console.warn('[Seasonal] Failed to load:', err.message));

})();
