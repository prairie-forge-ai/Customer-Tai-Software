/**
 * Homepage Sheet Utilities
 * Creates and manages module landing pages with clean black backgrounds
 */

import { BRANDING } from "./constants.js";

// Ada Assistant Configuration - imported from constants
const ADA_IMAGE_URL = BRANDING.ADA_IMAGE_URL;

// Supabase configuration for Ada API calls
const SUPABASE_URL = "https://jgciqwzwacaesqjaoadc.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE";

/**
 * Standalone Ada API call for popup modal
 * Works independently of module-specific implementations
 */
async function callAdaApiStandalone(prompt, context, messageHistory) {
    const COPILOT_URL = `${SUPABASE_URL}/functions/v1/copilot`;

    try {
        // Get current module and step context from global state
        const prairieForgeContext = window.PRAIRIE_FORGE_CONTEXT || {};
        const currentModule = prairieForgeContext.module || "general";
        const currentStep = prairieForgeContext.step !== null ? String(prairieForgeContext.step) : "analysis";
        const stepName = prairieForgeContext.stepName || null;
        
        console.log("[Ada Popup] Calling copilot API with context:", {
            module: currentModule,
            step: currentStep,
            stepName: stepName
        });
        
        // Merge step name into context for logging
        const enrichedContext = {
            ...context,
            stepName: stepName
        };

        const response = await fetch(COPILOT_URL, {
            method: "POST",
            mode: "cors",
            credentials: "omit",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${SUPABASE_ANON_KEY}`,
                "apikey": SUPABASE_ANON_KEY
            },
            body: JSON.stringify({
                prompt: prompt,
                context: enrichedContext,
                module: currentModule,
                function: currentStep,
                history: messageHistory?.slice(-10) || []
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("[Ada Popup] API error:", response.status, errorText);
            throw new Error(`API request failed: ${response.status}`);
        }

        const data = await response.json();
        console.log("[Ada Popup] API response received");

        if (data.message || data.response) {
            return data.message || data.response;
        }

        return "I received your question but couldn't generate a response. Please try again.";

    } catch (error) {
        console.error("[Ada Popup] API call failed:", error);
        return `I'm having trouble connecting right now. Error: ${error.message}. Please try again in a moment.`;
    }
}

/**
 * Creates or activates a module homepage sheet with formatted styling
 * @param {string} sheetName - Name of the homepage sheet (e.g., "PR_Homepage")
 * @param {string} title - Module title to display (e.g., "Payroll Recorder")
 * @param {string} subtitle - Description/subtext for the module
 */
export async function activateHomepageSheet(sheetName, title, subtitle) {
    if (typeof Excel === "undefined") {
        console.warn("Excel runtime not available for homepage sheet");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            // Check if sheet exists
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject, name, visibility");
            await context.sync();
            
            let targetSheet;
            
            if (sheet.isNullObject) {
                // Create the sheet
                targetSheet = context.workbook.worksheets.add(sheetName);
                await context.sync();
                
                // Set up the homepage content
                await setupHomepageContent(context, targetSheet, title, subtitle);
            } else {
                targetSheet = sheet;
                
                // Make sure sheet is visible before activating (SS_ sheets may be hidden)
                if (targetSheet.visibility !== Excel.SheetVisibility.visible) {
                    targetSheet.visibility = Excel.SheetVisibility.visible;
                    await context.sync();
                }
                
                // Update content in case title/subtitle changed
                await setupHomepageContent(context, targetSheet, title, subtitle);
            }
            
            // Activate the sheet
            targetSheet.activate();
            targetSheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error(`Error activating homepage sheet ${sheetName}:`, error);
    }
}

/**
 * Sets up the homepage sheet content and formatting
 */
async function setupHomepageContent(context, sheet, title, subtitle) {
    // Clear existing content
    try {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("isNullObject");
        await context.sync();
        if (!usedRange.isNullObject) {
            usedRange.clear();
            await context.sync();
        }
    } catch (e) {
        // Ignore clear errors
    }
    
    // Hide gridlines
    sheet.showGridlines = false;
    
    // Set up column widths for a clean look
    sheet.getRange("A:A").format.columnWidth = 400; // Content area
    sheet.getRange("B:B").format.columnWidth = 50;  // Right padding
    
    // Set row heights
    sheet.getRange("1:1").format.rowHeight = 60;  // Title row
    sheet.getRange("2:2").format.rowHeight = 30;  // Subtitle row
    
    // Write content starting at A1
    const data = [
        [title, ""],
        [subtitle, ""],
        ["", ""],
        ["", ""]
    ];
    
    const dataRange = sheet.getRangeByIndexes(0, 0, 4, 2);
    dataRange.values = data;
    
    // Apply black background to entire visible area
    const backgroundRange = sheet.getRange("A1:Z100");
    backgroundRange.format.fill.color = "#0f0f0f";
    
    // Format title (A1)
    const titleCell = sheet.getRange("A1");
    titleCell.format.font.bold = true;
    titleCell.format.font.size = 36;
    titleCell.format.font.color = "#ffffff";
    titleCell.format.font.name = "Segoe UI Light";
    titleCell.format.verticalAlignment = "Center";
    
    // Format subtitle (A2)
    const subtitleCell = sheet.getRange("A2");
    subtitleCell.format.font.size = 14;
    subtitleCell.format.font.color = "#a0a0a0";
    subtitleCell.format.font.name = "Segoe UI";
    subtitleCell.format.verticalAlignment = "Top";
    
    // Freeze panes to prevent scrolling issues
    sheet.freezePanes.freezeRows(0);
    sheet.freezePanes.freezeColumns(0);
    
    await context.sync();
}

/**
 * Homepage configurations for each module
 */
export const HOMEPAGE_CONFIG = {
    "module-selector": {
        sheetName: "SS_Homepage",
        title: "ForgeSuite",
        subtitle: "Select a module from the side panel to get started."
    },
    "payroll-recorder": {
        sheetName: "PR_Homepage",
        title: "Payroll Recorder",
        subtitle: "Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."
    },
    "pto-accrual": {
        sheetName: "PTO_Homepage",
        title: "PTO Accrual",
        subtitle: "Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."
    }
};

/**
 * Get homepage config by module key
 */
export function getHomepageConfig(moduleKey) {
    return HOMEPAGE_CONFIG[moduleKey] || HOMEPAGE_CONFIG["module-selector"];
}

/**
 * Renders the floating Ada assistant button
 * Call this when showing the home view
 */
export function renderAdaFab() {
    // Remove existing FAB if present
    removeAdaFab();
    
    const fab = document.createElement("button");
    fab.className = "pf-ada-fab";
    fab.id = "pf-ada-fab";
    fab.setAttribute("aria-label", "Ask Ada");
    fab.setAttribute("title", "Ask Ada");
    fab.innerHTML = `
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${ADA_IMAGE_URL}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `;
    
    document.body.appendChild(fab);
    
    // Bind click event
    fab.addEventListener("click", showAdaModal);
    
    return fab;
}

/**
 * Removes the Ada FAB from the DOM
 * Call this when navigating away from home view
 */
export function removeAdaFab() {
    const existingFab = document.getElementById("pf-ada-fab");
    if (existingFab) {
        existingFab.remove();
    }

    // Also remove modal if open
    const existingModal = document.getElementById("pf-ada-modal-overlay");
    if (existingModal) {
        existingModal.remove();
    }
}

/**
 * Determine the appropriate Ada configuration based on current context
 */
function getAdaModalContext() {
    // Import copilot functions dynamically
    let renderCopilotCard, bindCopilotCard, createExcelContextProvider;

    const loadCopilotFunctions = async () => {
        try {
            const copilotModule = await import('./copilot.js');
            renderCopilotCard = copilotModule.renderCopilotCard;
            bindCopilotCard = copilotModule.bindCopilotCard;
            createExcelContextProvider = copilotModule.createExcelContextProvider;
        } catch (e) {
            console.warn('Could not load copilot functions:', e);
        }
    };

    // Always provide a working Ada interface using standalone API
    const config = {
        subtext: "Your AI-powered assistant",
        copilotHtml: "", // Will be set by bindFunction
        bindFunction: async () => {
            await loadCopilotFunctions();
            if (renderCopilotCard && bindCopilotCard) {
                const container = document.getElementById('ada-modal-copilot');
                if (container) {
                    container.innerHTML = renderCopilotCard({
                        id: "ada-modal-copilot",
                        heading: "Ada",
                        subtext: "Ask questions about your data",
                        welcomeMessage: "Hi! I'm Ada, your AI assistant. How can I help you today?",
                        placeholder: "Ask about your data, analysis, or insights...",
                        quickActions: [
                            { id: "help", label: "What can you do?", prompt: "What kinds of questions can you help me with? What data do you have access to?" },
                            { id: "overview", label: "Data Overview", prompt: "Give me an overview of the data available in this workbook." },
                            { id: "tips", label: "Best Practices", prompt: "What are some best practices for using this tool effectively?" }
                        ],
                        contextProvider: createExcelContextProvider ? createExcelContextProvider({
                            config: 'SS_PF_Config'
                        }) : null,
                        onPrompt: callAdaApiStandalone
                    });

                    // Bind after a short delay to ensure DOM is ready
                    setTimeout(() => {
                        bindCopilotCard(container, {
                            id: "ada-modal-copilot",
                            contextProvider: createExcelContextProvider ? createExcelContextProvider({
                                config: 'SS_PF_Config'
                            }) : null,
                            onPrompt: callAdaApiStandalone
                        });
                    }, 200);
                }
            }
        }
    };

    return config;
}


/**
 * Shows the Ada assistant modal
 */
export function showAdaModal() {
    // Remove existing modal if present
    const existingModal = document.getElementById("pf-ada-modal-overlay");
    if (existingModal) {
        existingModal.remove();
    }

    // Get context configuration
    const contextConfig = getAdaModalContext();

    const overlay = document.createElement("div");
    overlay.className = "pf-ada-modal-overlay";
    overlay.id = "pf-ada-modal-overlay";

    overlay.innerHTML = `
        <div class="pf-ada-modal pf-ada-modal--chat">
            <div class="pf-ada-modal__header">
                <span class="pf-ada-modal__beta-tag">BETA</span>
                <button class="pf-ada-modal__close" id="ada-modal-close" title="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${ADA_IMAGE_URL}" alt="Ada" />
                <h2 class="pf-ada-modal__title">Ask Ada</h2>
                <p class="pf-ada-modal__subtitle">Your AI-powered assistant to help you troubleshoot, answer questions and perform deeper analyses.</p>
            </div>
            <div class="pf-ada-modal__body">
                <div class="pf-ada-copilot-container" id="ada-modal-copilot">
                    <div class="pf-ada-loading">
                        <div class="pf-ada-typing"><span></span><span></span><span></span></div>
                        <p>Loading Ada...</p>
                    </div>
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(overlay);

    // Trigger animation
    requestAnimationFrame(() => {
        overlay.classList.add("is-visible");
    });

    // Bind close events
    const closeBtn = document.getElementById("ada-modal-close");
    closeBtn?.addEventListener("click", hideAdaModal);

    // Close on overlay click
    overlay.addEventListener("click", (e) => {
        if (e.target === overlay) {
            hideAdaModal();
        }
    });

    // Close on Escape key
    const handleEscape = (e) => {
        if (e.key === "Escape") {
            hideAdaModal();
            document.removeEventListener("keydown", handleEscape);
        }
    };
    document.addEventListener("keydown", handleEscape);

    // Bind the copilot functionality
    if (contextConfig.bindFunction) {
        contextConfig.bindFunction();
    }
}

/**
 * Hides the Ada assistant modal
 */
export function hideAdaModal() {
    const overlay = document.getElementById("pf-ada-modal-overlay");
    if (overlay) {
        overlay.classList.remove("is-visible");
        setTimeout(() => {
            overlay.remove();
        }, 300);
    }
}

