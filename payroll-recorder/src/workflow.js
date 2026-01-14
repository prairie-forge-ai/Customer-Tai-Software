import { VERSION, WORKFLOW_STEPS as STEP_DETAILS, SHEET_NAMES } from "./constants.js";
import { BRANDING } from "../../Common/constants.js";
import { applyModuleTabVisibility, showAllSheets, showAndActivateSheet } from "../../Common/tab-visibility.js";
import { renderCopilotCard, bindCopilotCard, createExcelContextProvider, DEFAULT_OPTIONS } from "../../Common/copilot.js";
import { activateHomepageSheet, getHomepageConfig, renderAdaFab, removeAdaFab } from "../../Common/homepage-sheet.js";
import { initDatePicker } from "../../Common/date-picker.js";
import { saveRouteState, loadRouteState, buildRouteString, parseRouteString } from "../../Common/routerState.js";
import * as XLSX from "xlsx";
import {
    HOME_ICON_SVG,
    MODULES_ICON_SVG,
    ARROW_LEFT_SVG,
    USERS_ICON_SVG,
    BOOK_ICON_SVG,
    ARROW_RIGHT_SVG,
    MENU_ICON_SVG,
    TABLE_ICON_SVG,
    LOCK_CLOSED_SVG,
    LOCK_OPEN_SVG,
    CHECK_ICON_SVG,
    X_CIRCLE_SVG,
    X_ICON_SVG,
    LINK_ICON_SVG,
    CALCULATOR_ICON_SVG,
    SAVE_ICON_SVG,
    DOWNLOAD_ICON_SVG,
    UPLOAD_ICON_SVG,
    REFRESH_ICON_SVG,
    TRASH_ICON_SVG,
    SETTINGS_ICON_SVG,
    FILE_TEXT_ICON_SVG,
    GLOBE_ICON_SVG,
    GRID_ICON_SVG,
    CLIPBOARD_LIST_ICON_SVG,
    CHECK_CIRCLE_SVG,
    INFO_CIRCLE_SVG,
    ALERT_TRIANGLE_SVG,
    getStepIconSvg
} from "../../Common/icons.js";
import {
    renderInlineNotes,
    renderSignoff,
    updateLockButtonVisual,
    updateSaveButtonState,
    initSaveTracking
} from "../../Common/notes-signoff.js";
import { canCompleteStep, showBlockedToast } from "../../Common/workflow-validation.js";
import { initializeOffice } from "../../Common/gateway.js";
import { formatSheetHeaders, formatCurrencyColumn, formatDateColumn, NUMBER_FORMATS } from "../../Common/sheet-formatting.js";
import { formatXlsxWorksheet, formatExpenseReviewSheet, setXlsxColumnWidths, XLSX_COLUMN_WIDTHS } from "../../Common/xlsx-formatting.js";

const MODULE_KEY = "payroll-recorder";
const MODULE_ALIAS_TOKENS = ["payroll", "payroll recorder", "payroll review", "pr"];
const MODULE_NAME = "Payroll Recorder";

// Make module and step context globally accessible for Ada
window.PRAIRIE_FORGE_CONTEXT = {
    module: MODULE_KEY,
    step: null, // Will be updated when navigating to steps
    moduleName: MODULE_NAME
};

// =============================================================================
// WAREHOUSE API - Uses centralized client
// =============================================================================
// Do not call fetch() directly for warehouse. Use columnMapperRequest().
import { columnMapperRequest, forceRefreshAuth, debugAuthDump } from "../../Common/warehouse.js";

// =============================================================================
// BOOTSTRAP CONFIG - NOW HANDLED GLOBALLY
// =============================================================================
// Config sync is handled by Common/bootstrap.js called from module-selector.
// SS_PF_Config values are already populated when this module loads.
// This module just reads from configState (populated by loadConfigurationValues).

// =============================================================================
// OFFICE-SAFE DIALOGS (window.alert not supported in Office Add-ins)
// =============================================================================

/**
 * Show a toast notification in the UI (Office-safe alternative to window.alert)
 */
function showToast(message, type = "info", duration = 4000) {
    // Remove existing toasts
    document.querySelectorAll(".pf-toast").forEach(t => t.remove());
    
    const toast = document.createElement("div");
    toast.className = `pf-toast pf-toast--${type}`;
    const iconSvg = type === "success" ? CHECK_CIRCLE_SVG : type === "error" ? X_CIRCLE_SVG : INFO_CIRCLE_SVG;
    toast.innerHTML = `
        <div class="pf-toast-content">
            <span class="pf-toast-icon" aria-hidden="true">${iconSvg}</span>
            <span class="pf-toast-message">${message.replace(/\n/g, "<br>")}</span>
        </div>
        <button class="pf-toast-close" onclick="this.parentElement.remove()">×</button>
    `;
    
    // Add styles if not already present
    if (!document.getElementById("pf-toast-styles")) {
        const style = document.createElement("style");
        style.id = "pf-toast-styles";
        style.textContent = `
            .pf-toast {
                position: fixed;
                top: 20px;
                left: 50%;
                transform: translateX(-50%);
                background: var(--glass-bg);
                backdrop-filter: var(--glass-blur);
                -webkit-backdrop-filter: var(--glass-blur);
                color: white;
                padding: 16px 20px;
                border-radius: 8px;
                box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                z-index: 10000;
                max-width: 90%;
                display: flex;
                align-items: flex-start;
                gap: 12px;
                animation: toastIn 0.3s ease;
            }
            .pf-toast--success { border-left: 4px solid #22c55e; }
            .pf-toast--error { border-left: 4px solid #ef4444; }
            .pf-toast--info { border-left: 4px solid #3b82f6; }
            .pf-toast-content { display: flex; align-items: flex-start; gap: 8px; flex: 1; }
            .pf-toast-icon .pf-icon { width: 18px; height: 18px; }
            .pf-toast-message { font-size: 14px; line-height: 1.4; }
            .pf-toast-close { background: none; border: none; color: #888; font-size: 20px; cursor: pointer; padding: 0; margin-left: 8px; }
            .pf-toast-close:hover { color: white; }
            @keyframes toastIn { from { opacity: 0; transform: translateX(-50%) translateY(-20px); } }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(toast);
    
    if (duration > 0) {
        setTimeout(() => toast.remove(), duration);
    }
    
    return toast;
}

/**
 * Show a simple finalize prompt centered in the side panel
 * Appears after file download - stays until user clicks
 */
function showSaveCompletePrompt() {
    // Remove any existing prompts
    document.querySelectorAll(".pf-save-prompt").forEach(p => p.remove());
    
    const prompt = document.createElement("div");
    prompt.className = "pf-save-prompt";
    prompt.innerHTML = `
        <div class="pf-save-prompt-content">
            <div class="pf-save-prompt-title">Good work!</div>
            <div class="pf-save-prompt-subtitle">Ready to finalize?</div>
            <button type="button" class="pf-save-prompt-btn">
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <polyline points="20 6 9 17 4 12"/>
                </svg>
                Finalize
            </button>
        </div>
    `;
    
    // Add styles
    if (!document.getElementById("pf-save-prompt-styles")) {
        const style = document.createElement("style");
        style.id = "pf-save-prompt-styles";
        style.textContent = `
            .pf-save-prompt {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: var(--glass-bg);
                backdrop-filter: var(--glass-blur);
                -webkit-backdrop-filter: var(--glass-blur);
                border: 1px solid rgba(99, 102, 241, 0.3);
                color: white;
                padding: 32px 40px;
                border-radius: 20px;
                box-shadow: 
                    0 24px 64px rgba(0, 0, 0, 0.5),
                    0 0 0 1px rgba(99, 102, 241, 0.2) inset,
                    0 0 60px rgba(99, 102, 241, 0.1);
                z-index: 10002;
                text-align: center;
                animation: pf-prompt-fade-in 0.3s ease;
            }
            @keyframes pf-prompt-fade-in {
                from { opacity: 0; transform: translate(-50%, -50%) scale(0.95); }
                to { opacity: 1; transform: translate(-50%, -50%) scale(1); }
            }
            .pf-save-prompt-content {
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 8px;
            }
            .pf-save-prompt-title {
                font-size: 20px;
                font-weight: 700;
                color: #fff;
            }
            .pf-save-prompt-subtitle {
                font-size: 14px;
                color: rgba(255, 255, 255, 0.6);
                margin-bottom: 12px;
            }
            .pf-save-prompt-btn {
                background: linear-gradient(145deg, #6366f1, #4f46e5);
                border: none;
                color: white;
                padding: 16px 40px;
                border-radius: 12px;
                font-size: 18px;
                font-weight: 700;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
                transition: all 0.2s ease;
                box-shadow: 0 8px 24px rgba(99, 102, 241, 0.4), 0 0 0 2px rgba(99, 102, 241, 0.2);
                margin-top: 8px;
            }
            .pf-save-prompt-btn:hover {
                transform: translateY(-3px);
                box-shadow: 0 12px 32px rgba(99, 102, 241, 0.5), 0 0 0 2px rgba(99, 102, 241, 0.3);
                background: linear-gradient(145deg, #7c3aed, #6366f1);
            }
            .pf-save-prompt-btn:active {
                transform: translateY(-1px);
            }
            .pf-save-prompt.closing {
                animation: pf-prompt-fade-out 0.2s ease forwards;
            }
            @keyframes pf-prompt-fade-out {
                to { opacity: 0; transform: translate(-50%, -50%) scale(0.95); }
            }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(prompt);
    
    // Handle "Finalize" button click
    const doneBtn = prompt.querySelector(".pf-save-prompt-btn");
    doneBtn.addEventListener("click", () => {
        prompt.classList.add("closing");
        setTimeout(() => {
            prompt.remove();
            // Now show the celebratory toast
            showArchiveSuccessToast();
        }, 200);
    });
}

/**
 * Show celebratory archive success toast with rotating messages
 * Displays for 5 seconds then redirects to Module Selector
 */
function showArchiveSuccessToast() {
    // Remove existing toasts
    document.querySelectorAll(".pf-toast, .pf-success-toast").forEach(t => t.remove());
    
    // Lucide icon SVGs with Prairie Forge purple
    const messages = [
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z"/></svg>`,
            title: "Done.",
            subtitle: "Efficiency unlocked."
        },
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"/><path d="m9 12 2 2 4-4"/></svg>`,
            title: "Locked in.",
            subtitle: "You're good to go."
        },
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="1"/><path d="M20.2 20.2c2.04-2.03.02-7.36-4.5-11.9-4.54-4.52-9.87-6.54-11.9-4.5-2.04 2.03-.02 7.36 4.5 11.9 4.54 4.52 9.87 6.54 11.9 4.5Z"/><path d="M15.7 15.7c4.52-4.54 6.54-9.87 4.5-11.9-2.03-2.04-7.36-.02-11.9 4.5-4.52 4.54-6.54 9.87-4.5 11.9 2.03 2.04 7.36.02 11.9-4.5Z"/></svg>`,
            title: "Stored.",
            subtitle: "Everything stays aligned."
        },
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>`,
            title: "Payroll archived.",
            subtitle: "Flow restored."
        }
    ];
    
    // Pick a random message
    const msg = messages[Math.floor(Math.random() * messages.length)];
    
    const toast = document.createElement("div");
    toast.className = "pf-success-toast";
    toast.innerHTML = `
        <div class="pf-success-toast-icon">${msg.icon}</div>
        <div class="pf-success-toast-text">
            <div class="pf-success-toast-title">${msg.title}</div>
            <div class="pf-success-toast-subtitle">${msg.subtitle}</div>
        </div>
        <div class="pf-success-toast-progress"></div>
    `;
    
    // Add styles if not already present
    if (!document.getElementById("pf-success-toast-styles")) {
        const style = document.createElement("style");
        style.id = "pf-success-toast-styles";
        style.textContent = `
            .pf-success-toast {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: var(--glass-bg);
                backdrop-filter: var(--glass-blur);
                -webkit-backdrop-filter: var(--glass-blur);
                border: 1px solid rgba(99, 102, 241, 0.3);
                color: white;
                padding: 36px 48px 24px;
                border-radius: 24px;
                box-shadow: 
                    0 32px 64px rgba(0, 0, 0, 0.5),
                    0 0 0 1px rgba(99, 102, 241, 0.2) inset,
                    0 0 80px rgba(99, 102, 241, 0.15);
                z-index: 10002;
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 16px;
                animation: pf-success-in 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
                overflow: hidden;
                text-align: center;
                min-width: 220px;
            }
            @keyframes pf-success-in {
                from { 
                    opacity: 0; 
                    transform: translate(-50%, -50%) scale(0.9);
                }
                to { 
                    opacity: 1; 
                    transform: translate(-50%, -50%) scale(1);
                }
            }
            @keyframes pf-success-out {
                from { 
                    opacity: 1; 
                    transform: translate(-50%, -50%) scale(1);
                }
                to { 
                    opacity: 0; 
                    transform: translate(-50%, -50%) scale(0.95) translateY(-20px);
                }
            }
            .pf-success-toast-icon {
                width: 64px;
                height: 64px;
                background: linear-gradient(145deg, #6366f1, #4f46e5);
                border-radius: 18px;
                display: flex;
                align-items: center;
                justify-content: center;
                color: white;
                box-shadow: 0 8px 24px rgba(99, 102, 241, 0.4);
                animation: pf-icon-pulse 2s ease-in-out infinite;
            }
            @keyframes pf-icon-pulse {
                0%, 100% { box-shadow: 0 8px 24px rgba(99, 102, 241, 0.4); }
                50% { box-shadow: 0 8px 32px rgba(99, 102, 241, 0.6); }
            }
            .pf-success-toast-text {
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 4px;
            }
            .pf-success-toast-title {
                font-size: 22px;
                font-weight: 700;
                color: #fff;
                letter-spacing: -0.5px;
            }
            .pf-success-toast-subtitle {
                font-size: 15px;
                color: rgba(255, 255, 255, 0.6);
                font-weight: 500;
            }
            .pf-success-toast-progress {
                position: absolute;
                bottom: 0;
                left: 0;
                height: 4px;
                background: linear-gradient(90deg, #6366f1, #a855f7, #6366f1);
                background-size: 200% 100%;
                animation: pf-progress-shrink 5s linear forwards, pf-progress-shimmer 1s linear infinite;
                border-radius: 0 0 24px 24px;
            }
            @keyframes pf-progress-shrink {
                from { width: 100%; }
                to { width: 0%; }
            }
            @keyframes pf-progress-shimmer {
                from { background-position: 200% 0; }
                to { background-position: -200% 0; }
            }
            .pf-success-toast.closing {
                animation: pf-success-out 0.3s ease forwards;
            }
            .pf-success-backdrop {
                position: fixed;
                inset: 0;
                background: rgba(0, 0, 0, 0.4);
                backdrop-filter: blur(4px);
                -webkit-backdrop-filter: blur(4px);
                z-index: 10001;
                animation: pf-confirm-fade-in 0.2s ease;
            }
        `;
        document.head.appendChild(style);
    }
    
    // Add backdrop
    const backdrop = document.createElement("div");
    backdrop.className = "pf-success-backdrop";
    document.body.appendChild(backdrop);
    document.body.appendChild(toast);
    
    // After 5 seconds, close and redirect to Module Selector
    setTimeout(() => {
        toast.classList.add("closing");
        backdrop.style.opacity = "0";
        backdrop.style.transition = "opacity 0.3s ease";
        
        setTimeout(() => {
            toast.remove();
            backdrop.remove();
            // Navigate to Module Selector homepage
            navigateToModuleSelector();
        }, 300);
    }, 5000);
}

/**
 * Show a confirmation dialog (Office-safe alternative to window.confirm)
 * Apple-inspired design with glassmorphism
 * Returns a Promise that resolves to true/false
 */
function showConfirm(message, options = {}) {
    const {
        title = "Confirm Action",
        confirmText = "Continue",
        cancelText = "Cancel",
        icon = FILE_TEXT_ICON_SVG,
        destructive = false
    } = options;
    
    return new Promise((resolve) => {
        // Remove existing dialogs
        document.querySelectorAll(".pf-confirm-overlay").forEach(d => d.remove());
        
        const overlay = document.createElement("div");
        overlay.className = "pf-confirm-overlay";
        overlay.innerHTML = `
            <div class="pf-confirm-dialog">
                <div class="pf-confirm-icon">${icon}</div>
                <div class="pf-confirm-title">${title}</div>
                <div class="pf-confirm-message">${message.replace(/\n/g, "<br>")}</div>
                <div class="pf-confirm-buttons">
                    <button class="pf-confirm-btn pf-confirm-btn--cancel">${cancelText}</button>
                    <button class="pf-confirm-btn pf-confirm-btn--ok ${destructive ? 'pf-confirm-btn--destructive' : ''}">${confirmText}</button>
                </div>
            </div>
        `;
        
        // Add styles if not already present
        if (!document.getElementById("pf-confirm-styles")) {
            const style = document.createElement("style");
            style.id = "pf-confirm-styles";
            style.textContent = `
                .pf-confirm-overlay {
                    position: fixed;
                    inset: 0;
                    background: rgba(0, 0, 0, 0.5);
                    backdrop-filter: blur(8px);
                    -webkit-backdrop-filter: blur(8px);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10001;
                    animation: pf-confirm-fade-in 0.2s ease;
                }
                @keyframes pf-confirm-fade-in {
                    from { opacity: 0; }
                    to { opacity: 1; }
                }
                @keyframes pf-confirm-scale-in {
                    from { opacity: 0; transform: scale(0.95) translateY(-10px); }
                    to { opacity: 1; transform: scale(1) translateY(0); }
                }
                .pf-confirm-dialog {
                    background: var(--glass-bg);
                    backdrop-filter: var(--glass-blur);
                    -webkit-backdrop-filter: var(--glass-blur);
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    color: white;
                    padding: 28px 32px;
                    border-radius: 20px;
                    max-width: 380px;
                    width: 90%;
                    box-shadow: 
                        0 24px 48px rgba(0, 0, 0, 0.4),
                        0 0 0 1px rgba(255, 255, 255, 0.05) inset,
                        0 1px 0 rgba(255, 255, 255, 0.1) inset;
                    text-align: center;
                    animation: pf-confirm-scale-in 0.25s cubic-bezier(0.34, 1.56, 0.64, 1);
                }
                .pf-confirm-icon {
                    font-size: 48px;
                    margin-bottom: 16px;
                    filter: drop-shadow(0 4px 8px rgba(0,0,0,0.3));
                }
                .pf-confirm-title {
                    font-size: 18px;
                    font-weight: 600;
                    color: #fff;
                    margin-bottom: 12px;
                    letter-spacing: -0.3px;
                }
                .pf-confirm-message {
                    font-size: 14px;
                    line-height: 1.6;
                    color: rgba(255, 255, 255, 0.7);
                    margin-bottom: 24px;
                    text-align: left;
                }
                .pf-confirm-buttons {
                    display: flex;
                    gap: 12px;
                    justify-content: center;
                }
                .pf-confirm-btn {
                    flex: 1;
                    padding: 12px 24px;
                    border-radius: 12px;
                    border: none;
                    cursor: pointer;
                    font-size: 15px;
                    font-weight: 600;
                    letter-spacing: -0.2px;
                    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
                }
                .pf-confirm-btn:active {
                    transform: scale(0.97);
                }
                .pf-confirm-btn--cancel {
                    background: rgba(255, 255, 255, 0.08);
                    color: rgba(255, 255, 255, 0.9);
                    border: 1px solid rgba(255, 255, 255, 0.1);
                }
                .pf-confirm-btn--cancel:hover {
                    background: rgba(255, 255, 255, 0.12);
                    border-color: rgba(255, 255, 255, 0.15);
                }
                .pf-confirm-btn--ok {
                    background: linear-gradient(145deg, #6366f1, #4f46e5);
                    color: white;
                    box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4);
                }
                .pf-confirm-btn--ok:hover {
                    background: linear-gradient(145deg, #818cf8, #6366f1);
                    box-shadow: 0 6px 16px rgba(99, 102, 241, 0.5);
                    transform: translateY(-1px);
                }
                .pf-confirm-btn--destructive {
                    background: linear-gradient(145deg, #ef4444, #dc2626);
                    box-shadow: 0 4px 12px rgba(239, 68, 68, 0.4);
                }
                .pf-confirm-btn--destructive:hover {
                    background: linear-gradient(145deg, #f87171, #ef4444);
                    box-shadow: 0 6px 16px rgba(239, 68, 68, 0.5);
                }
            `;
            document.head.appendChild(style);
        }
        
        document.body.appendChild(overlay);
        
        // Close on overlay click (outside dialog)
        overlay.addEventListener("click", (e) => {
            if (e.target === overlay) {
                overlay.remove();
                resolve(false);
            }
        });
        
        overlay.querySelector(".pf-confirm-btn--cancel").onclick = () => {
            overlay.remove();
            resolve(false);
        };
        overlay.querySelector(".pf-confirm-btn--ok").onclick = () => {
            overlay.remove();
            resolve(true);
        };
    });
}
const MODULE_CONFIG_SHEET = SHEET_NAMES.CONFIG || "SS_PF_Config";
const CONFIG_TABLE_CANDIDATES = ["SS_PF_Config"];
const DEFAULT_CONFIG_CATEGORY = "Run Settings";
const HERO_COPY =
    "Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel. Every run follows the same guidance so you stay audit-ready.";

// Filter out hidden steps from navigation
const WORKFLOW_STEPS = STEP_DETAILS
    .filter(step => !step.hidden)
    .map(({ id, title }) => ({ id, title }));

// SS_PF_Config structure: Category (0), Field (1), Value (2), Permanent (3)
// For module-prefix rows: Field=Prefix (e.g., "PR_"), Value=Module Key (e.g., "payroll-recorder")
// For Config rows: Field=Setting Name, Value=Setting Value, Permanent=Y/N flag
const CONFIG_COLUMNS = {
    TYPE: 0,      // Category column (A)
    FIELD: 1,     // Field name column (B)
    VALUE: 2,     // Value column (C)
    PERMANENT: 3, // Permanent column (D) - Y/N flag for archive persistence
    TITLE: -1     // Not used
};
const DEFAULT_CONFIG_TYPE = "Run Settings";
const DEFAULT_CONFIG_TITLE = "";
const DEFAULT_CONFIG_PERMANENT = "N";
const NOTE_PLACEHOLDER = "Enter notes here...";
const JE_TOTAL_DEBIT_FIELD = "PR_JE_Debit_Total";
const JE_TOTAL_CREDIT_FIELD = "PR_JE_Credit_Total";
const JE_DIFFERENCE_FIELD = "PR_JE_Difference";

// Step notes/sign-off fields - Pattern: PR_{Type}_{StepName}
const STEP_NOTES_FIELDS = {
    0: { note: "PR_Notes_Config", reviewer: "PR_Reviewer_Config", signOff: "PR_SignOff_Config" },
    1: { note: "PR_Notes_Import", reviewer: "PR_Reviewer_Import", signOff: "PR_SignOff_Import" },
    2: { note: "PR_Notes_Review", reviewer: "PR_Reviewer_Review", signOff: "PR_SignOff_Review" },
    3: { note: "PR_Notes_JE", reviewer: "PR_Reviewer_JE", signOff: "PR_SignOff_JE" },
    4: { note: "PR_Notes_Archive", reviewer: "PR_Reviewer_Archive", signOff: "PR_SignOff_Archive" }
};
const STEP_COMPLETE_FIELDS = {
    0: "PR_Complete_Config",
    1: "PR_Complete_Import",
    2: "PR_Complete_Review",
    3: "PR_Complete_JE",
    4: "PR_Complete_Archive"
};
// Step → Sheet mapping for tab activation on navigation
// ONE-WAY SYNC: Panel navigation activates Excel tabs (not vice versa)
const STEP_SHEET_MAP = {
    0: "PR_Homepage",            // Configuration
    1: "PR_Data_Clean",          // Upload & Validate
    2: "PR_Expense_Review",      // Expense Review
    3: "PR_JE_Draft",            // Journal Entry Prep
    4: "PR_Archive_Summary"      // Archive & Clear
};
// Config field names - Pattern: PR_{Descriptor}
const CONFIG_REVIEWER_FIELD = "PR_Reviewer";
const PAYROLL_PROVIDER_FIELD = "PR_Payroll_Provider";

const appState = {
    statusText: "",
    focusedIndex: 0,
    activeView: "home",
    activeStepId: null,
    stepStatuses: WORKFLOW_STEPS.reduce((map, step) => ({ ...map, [step.id]: "pending" }), {})
};

const configState = {
    loaded: false,
    values: {},
    permanents: {},
    overrides: {
        accountingPeriod: false,
        jeId: false
    }
};

const pendingWrites = new Map();
let resolvedConfigTableName = null;
// Payroll date - primary name first, legacy fallbacks for migration
const PAYROLL_DATE_ALIASES = [
    "PR_Payroll_Date",
    "Payroll Date (YYYY-MM-DD)", 
    "Payroll_Date", 
    "Payroll Date",
    "Payroll_Date_(YYYY-MM-DD)"
];

let pendingScrollIndex = null;

const validationState = {
    loading: false,
    lastError: null,
    prDataTotal: null,
    cleanTotal: null,
    reconDifference: null,
    bankAmount: "",
    bankDifference: null,
    plugEnabled: false
};

const expenseReviewState = {
    loading: false,
    lastError: null,
    periods: [],
    copilotResponse: "",
    // Data completeness check
    completenessCheck: {
        currentPeriod: null,   // { match: true/false, prDataClean: number, currentTotal: number }
        historicalPeriods: null // { match: true/false, archiveSum: number, periodsSum: number }
    },
    // Taxonomy-driven classification
    unclassifiedColumns: [],  // Advisory: columns not found in dictionary or dimensions
    measureColumns: [],       // Columns used as measures (from taxonomy)
    dimensionColumns: [],     // Columns used as dimensions (from taxonomy)
    // Unclassified diagnostic data
    unclassifiedTotals: {},   // { columnName: totalDollarAmount } for numeric unclassified columns
    sumUnclassifiedNumeric: 0, // Total of all unclassified numeric columns
    classificationDelta: null  // PR_Data_Clean total - Expense Review total (for diagnosis)
};

// =============================================================================
// ADA INSIGHTS STATE - Expense Review AI Assistant
// =============================================================================

const adaInsightsState = {
    loading: false,
    lastError: null,
    response: null,           // Parsed response from Ada
    contextPack: null,        // Last generated context pack (for debugging)
    lastRefresh: null,        // Timestamp of last insights generation
    collapsed: false          // Panel collapse state
};

const journalState = {
    debitTotal: null,
    creditTotal: null,
    difference: null,
    cleanTotal: null,              // PR_Data_Clean total for comparison
    loading: false,
    lastError: null,
    // Validation results (consistent with PTO module)
    validationRun: false,
    issues: [],                    // [{check: "name", passed: false, detail: "explanation"}]
    // Allocation-specific state (unified with measure universe)
    allocationLines: [],           // Aggregated allocation lines
    unmappedColumns: [],           // PF columns without GL mapping
    unmappedTotal: 0,              // Dollar amount in unmapped columns
    glMappings: null,              // Cached GL mappings from database
    lastGenerated: null            // Timestamp of last generation
};

/**
 * Run data completeness check for Payroll Expense Review
 * UNIFIED: Uses getPRDataCleanMeasureUniverse() as the single source of truth
 * Validates that Expense Review total matches Step 1's PR_Data_Clean total
 */
async function runPayrollCompletenessCheck() {
    console.log("Completeness Check - Starting...");
    if (!hasExcelRuntime()) {
        console.log("Completeness Check - Excel runtime not available");
        return;
    }
    
    // CRITICAL: Get measure universe (same as Step 1) as single source of truth
    const measureUniverse = await getPRDataCleanMeasureUniverse();
    const step1Total = measureUniverse.total;
    
    console.log("════════════════════════════════════════════════════════════════");
    console.log("COMPLETENESS CHECK: Step 1 Measure Universe");
    console.log("════════════════════════════════════════════════════════════════");
    console.log(`Step 1 PR_Data_Clean Total: $${step1Total.toLocaleString(undefined, { minimumFractionDigits: 2 })}`);
    console.log(`Included measures: ${measureUniverse.includedMeasureHeaders.length}`);
    console.log("════════════════════════════════════════════════════════════════\n");
    
    // Store in state for UI display
    expenseReviewState.measureUniverseTotal = step1Total;
    expenseReviewState.measureUniverseHeaders = measureUniverse.includedMeasureHeaders;
    
    try {
        await Excel.run(async (context) => {
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
            const archiveTotalsSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_TOTALS);
            
            cleanSheet.load("isNullObject");
            archiveSheet.load("isNullObject");
            archiveTotalsSheet.load("isNullObject");
            await context.sync();
            
            const results = {
                currentPeriod: null,
                historicalPeriods: null
            };
            
            // Check 1: Current period total from PR_Data_Clean
            // DIAGNOSTIC: Show three different totals to identify filtering issues
            if (!cleanSheet.isNullObject) {
                const cleanRange = cleanSheet.getUsedRangeOrNullObject();
                cleanRange.load("values");
                await context.sync();
                
                if (!cleanRange.isNullObject && cleanRange.values && cleanRange.values.length > 1) {
                    const rawHeaders = cleanRange.values[0] || [];
                    const headers = rawHeaders.map(h => String(h || "").trim());
                    const headersLower = headers.map(h => h.toLowerCase());
                    const dataRows = cleanRange.values.slice(1);
                    
                    console.log("Completeness Check - PR_Data_Clean headers:", headers);
                    
                    const taxonomy = expenseTaxonomyCache;
                    
                    // ═══════════════════════════════════════════════════════════════════
                    // DIAGNOSTIC: Calculate three totals to identify filtering issues
                    // ═══════════════════════════════════════════════════════════════════
                    
                    // Identify dimension columns (to exclude from numeric totals)
                    const dimensionIndices = new Set();
                    headers.forEach((header, idx) => {
                        const headerLower = header.toLowerCase();
                        if (taxonomy.loaded && taxonomy.dimensions && taxonomy.dimensions.has(headerLower)) {
                            dimensionIndices.add(idx);
                        }
                    });
                    
                    // TOTAL 1: Pure sheet-driven (all numeric columns except dimensions)
                    const numericColumns1 = [];
                    const columnTotals1 = {};
                    headers.forEach((header, idx) => {
                        if (dimensionIndices.has(idx)) return;
                        let colTotal = 0;
                        let hasNumeric = false;
                        for (const row of dataRows) {
                            const val = Number(row[idx]);
                            if (!isNaN(val) && val !== 0) {
                                hasNumeric = true;
                                colTotal += val;
                            }
                        }
                        if (hasNumeric) {
                            numericColumns1.push({ idx, header, total: colTotal });
                            columnTotals1[header] = colTotal;
                        }
                    });
                    const total1 = numericColumns1.reduce((sum, c) => sum + c.total, 0);
                    
                    // TOTAL 2: Dictionary-matched (headers that exist in dictionary, no include/side filtering)
                    const numericColumns2 = [];
                    const columnTotals2 = {};
                    headers.forEach((header, idx) => {
                        if (dimensionIndices.has(idx)) return;
                        const headerLower = header.toLowerCase();
                        const inDict = taxonomy.loaded && taxonomy.measures && taxonomy.measures[headerLower];
                        if (!inDict) return;
                        let colTotal = 0;
                        for (const row of dataRows) {
                            const val = Number(row[idx]);
                            if (!isNaN(val)) colTotal += val;
                        }
                        numericColumns2.push({ idx, header, total: colTotal, meta: taxonomy.measures[headerLower] });
                        columnTotals2[header] = colTotal;
                    });
                    const total2 = numericColumns2.reduce((sum, c) => sum + c.total, 0);
                    
                    // TOTAL 3: Current logic (with include/side filtering)
                    // Uses shouldIncludeInExpenseReview with full exclusion tracking
                    const measureColumnIndices = [];
                    const excludedColumns = []; // Track excluded columns with dollar amounts and reasons
                    const excludedByRule = { 
                        summaryExclusion: [], 
                        includeFalse: [], 
                        sideEE: [], 
                        sideNA: [],
                        dimension: []
                    };
                    const suspiciousNumeric = []; // Advisory: numeric columns that may not be dollar amounts
                    
                    headers.forEach((header, idx) => {
                        const headerLower = header.toLowerCase();
                        
                        // Track dimension exclusions
                        if (dimensionIndices.has(idx)) {
                            // Calculate total for dimension to show in diagnostics
                            let colTotal = 0;
                            for (const row of dataRows) {
                                const val = Number(row[idx]);
                                if (!isNaN(val)) colTotal += val;
                            }
                            if (Math.abs(colTotal) > 0) {
                                excludedByRule.dimension.push(header);
                                excludedColumns.push({ header, total: colTotal, reason: "dimension" });
                            }
                            return;
                        }
                        
                        // Check if it's a numeric column
                        let hasNumeric = false;
                        let sampleValues = [];
                        let colTotal = 0;
                        for (const row of dataRows) {
                            const val = Number(row[idx]);
                            if (!isNaN(val)) {
                                colTotal += val;
                                if (val !== 0) {
                                    hasNumeric = true;
                                    if (sampleValues.length < 10) sampleValues.push(val);
                                }
                            }
                        }
                        if (!hasNumeric) return;
                        
                        // PR_Data_Clean headers drive the measure list
                        // Dictionary only enriches metadata (bucket/side/order)
                        const meta = (taxonomy.loaded && taxonomy.measures) ? taxonomy.measures[headerLower] : null;
                        
                        // Use shouldIncludeInExpenseReview with header for summary check
                        const inclusionResult = shouldIncludeInExpenseReview(meta, headerLower);
                        
                        if (!inclusionResult.include) {
                            // Track exclusion with dollar amount and reason
                            excludedColumns.push({ header, total: colTotal, reason: inclusionResult.reason });
                            
                            // Categorize by rule for summary
                            if (inclusionResult.reason === "summary_exclusion") {
                                excludedByRule.summaryExclusion.push(header);
                            } else if (inclusionResult.reason === "include=false") {
                                excludedByRule.includeFalse.push(header);
                            } else if (inclusionResult.reason === "side='ee'") {
                                excludedByRule.sideEE.push(header);
                            } else if (inclusionResult.reason === "side='na'") {
                                excludedByRule.sideNA.push(header);
                            }
                            return;
                        }
                        
                        // Lightweight sanity guard: flag suspicious numeric columns
                        // Advisory only - still include them
                        if (!meta && !headerLower.includes("amount")) {
                            const allSmallIntegers = sampleValues.every(v => Number.isInteger(v) && Math.abs(v) < 1000);
                            if (allSmallIntegers && sampleValues.length > 0) {
                                suspiciousNumeric.push({ header, total: colTotal, note: "Included (numeric, verify classification)" });
                            }
                        }
                        
                        // Include this column
                        const sign = meta?.sign ?? 1;
                        measureColumnIndices.push({ idx, sign, header, meta });
                    });
                    
                    // Calculate total 3
                    let total3 = 0;
                    const columnTotals3 = {};
                    measureColumnIndices.forEach(({ idx, sign, header }) => {
                        let colTotal = 0;
                        for (const row of dataRows) {
                            const val = Number(row[idx]) || 0;
                            colTotal += val * sign;
                        }
                        columnTotals3[header] = colTotal;
                        total3 += colTotal;
                    });
                    
                    // Calculate delta explanation
                    const totalExcluded = excludedColumns.reduce((sum, c) => sum + c.total, 0);
                    const delta = total1 - total3;
                    
                    // ═══════════════════════════════════════════════════════════════════
                    // DIAGNOSTIC OUTPUT
                    // ═══════════════════════════════════════════════════════════════════
                    const formatDollar = (n) => `$${n.toLocaleString(undefined, { minimumFractionDigits: 2 })}`;
                    const topN = (obj, n) => Object.entries(obj).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1])).slice(0, n);
                    const topNArr = (arr, n) => arr.slice().sort((a, b) => Math.abs(b.total) - Math.abs(a.total)).slice(0, n);
                    
                    console.log("════════════════════════════════════════════════════════════════");
                    console.log("DIAGNOSTIC: PR_Data_Clean Total Analysis");
                    console.log("════════════════════════════════════════════════════════════════");
                    
                    console.log(`\n1. TOTAL_NUMERIC_SHEET (all numeric columns except dimensions):`);
                    console.log(`   Total: ${formatDollar(total1)} (${numericColumns1.length} columns)`);
                    console.log(`   Top 15 columns:`);
                    topN(columnTotals1, 15).forEach(([col, amt], i) => console.log(`      ${i+1}. ${col}: ${formatDollar(amt)}`));
                    
                    console.log(`\n2. TOTAL_DICTIONARY_MATCHED (headers in dictionary, no filtering):`);
                    console.log(`   Total: ${formatDollar(total2)} (${numericColumns2.length} columns)`);
                    console.log(`   Top 15 columns:`);
                    topN(columnTotals2, 15).forEach(([col, amt], i) => console.log(`      ${i+1}. ${col}: ${formatDollar(amt)}`));
                    
                    console.log(`\n3. TOTAL_CURRENT_LOGIC (with include/side filtering):`);
                    console.log(`   Total: ${formatDollar(total3)} (${measureColumnIndices.length} columns)`);
                    console.log(`   Top 15 columns:`);
                    topN(columnTotals3, 15).forEach(([col, amt], i) => console.log(`      ${i+1}. ${col}: ${formatDollar(amt)}`));
                    
                    console.log(`\n4. DELTA EXPLANATION:`);
                    console.log(`   TOTAL_NUMERIC_SHEET - TOTAL_CURRENT_LOGIC = ${formatDollar(delta)}`);
                    console.log(`   Sum of excluded columns = ${formatDollar(totalExcluded)}`);
                    const deltaReconciles = Math.abs(delta - totalExcluded) < 1;
                    console.log(`   Reconciles: ${deltaReconciles ? "YES" : "NO (investigate)"}`);
                    if (!deltaReconciles) {
                        console.log(`   Unexplained gap: ${formatDollar(delta - totalExcluded)}`);
                    }
                    
                    console.log(`\n5. TOP 10 EXCLUDED BY DOLLARS:`);
                    topNArr(excludedColumns, 10).forEach((col, i) => {
                        console.log(`      ${i+1}. ${col.header}: ${formatDollar(col.total)} (${col.reason})`);
                    });
                    
                    console.log(`\nEXCLUSION SUMMARY BY RULE:`);
                    console.log(`   summary_exclusion (Gross/Net Pay): ${excludedByRule.summaryExclusion.length} → ${excludedByRule.summaryExclusion.join(", ") || "(none)"}`);
                    console.log(`   expense_review_include=false: ${excludedByRule.includeFalse.length} → ${excludedByRule.includeFalse.join(", ") || "(none)"}`);
                    console.log(`   side='ee' (employee): ${excludedByRule.sideEE.length} → ${excludedByRule.sideEE.join(", ") || "(none)"}`);
                    console.log(`   side='na' (summary): ${excludedByRule.sideNA.length} → ${excludedByRule.sideNA.join(", ") || "(none)"}`);
                    console.log(`   dimension: ${excludedByRule.dimension.length} → ${excludedByRule.dimension.slice(0, 5).join(", ")}${excludedByRule.dimension.length > 5 ? "..." : ""}`);
                    
                    if (suspiciousNumeric.length > 0) {
                        console.log(`\nSUSPICIOUS NUMERIC (included, verify classification):`);
                        suspiciousNumeric.forEach((col, i) => {
                            console.log(`      ${i+1}. ${col.header}: ${formatDollar(col.total)} - ${col.note}`);
                        });
                    }
                    
                    console.log("════════════════════════════════════════════════════════════════\n");
                    
                    // Get current period total from state (Expense Review aggregation)
                    const currentPeriodTotal = expenseReviewState.periods?.[0]?.summary?.total || 0;
                    
                    // UNIFIED: Use Step 1 measure universe total as the authoritative value
                    const prDataCleanTotal = step1Total;
                    
                    console.log("\n════════════════════════════════════════════════════════════════");
                    console.log("RECONCILIATION CHECK");
                    console.log("════════════════════════════════════════════════════════════════");
                    console.log(`Step 1 Universe Total:   ${formatDollar(step1Total)}`);
                    console.log(`Expense Review Total:    ${formatDollar(currentPeriodTotal)}`);
                    const reconDelta = Math.abs(step1Total - currentPeriodTotal);
                    console.log(`Delta:                   ${formatDollar(reconDelta)}`);
                    const match = reconDelta < 1;
                    console.log(`Reconciled:              ${match ? "YES" : "NO - INVESTIGATE"}`);
                    console.log("════════════════════════════════════════════════════════════════\n");
                    
                    results.currentPeriod = {
                        match,
                        prDataClean: prDataCleanTotal,
                        currentTotal: currentPeriodTotal
                    };
                }
            }
            
            // Check 2: Historical periods - match by date against PR_Archive_Totals (preferred) or PR_Archive_Summary
            const historicalPeriods = (expenseReviewState.periods || []).slice(1, 6);
            console.log("Completeness Check - Looking for periods:", historicalPeriods.map(p => p.key || p.label));

            let archiveLookup = new Map();
            let archiveSource = null;

            if (!archiveTotalsSheet.isNullObject) {
                const archiveTotalsRange = archiveTotalsSheet.getUsedRangeOrNullObject();
                archiveTotalsRange.load("values");
                await context.sync();

                if (!archiveTotalsRange.isNullObject && archiveTotalsRange.values && archiveTotalsRange.values.length > 1) {
                    const archivePeriods = buildArchivePeriodsFromTotalsSheet(archiveTotalsRange.values);
                    archivePeriods.forEach(p => {
                        const normalizedKey = normalizeDateForLookup(p.key);
                        if (!normalizedKey) return;
                        archiveLookup.set(normalizedKey, Number(p.summary?.total) || 0);
                    });

                    if (archiveLookup.size > 0) {
                        archiveSource = "PR_Archive_Totals";
                        console.log("Completeness Check - Using PR_Archive_Totals lookup keys:", Array.from(archiveLookup.keys()));
                    }
                }
            }

            if (!archiveSource && !archiveSheet.isNullObject) {
                const archiveRange = archiveSheet.getUsedRangeOrNullObject();
                archiveRange.load("values");
                await context.sync();

                if (!archiveRange.isNullObject && archiveRange.values && archiveRange.values.length > 1) {
                    const headers = (archiveRange.values[0] || []).map(h => String(h || "").toLowerCase().trim());

                    // Find date column (pay period, payroll date, date, period)
                    const dateIdx = headers.findIndex(h => 
                        h.includes("pay period") || h.includes("payroll date") || 
                        h === "date" || h === "period" || h.includes("period")
                    );

                    // Find amount/total column
                    const amountIdx = headers.findIndex(h => h.includes("amount"));
                    const totalIdx = amountIdx >= 0 ? amountIdx : headers.findIndex(h => 
                        h === "total" || h === "all-in" || h === "allin" || 
                        h === "all-in total" || h === "total payroll" || h.includes("total")
                    );

                    console.log("Completeness Check - PR_Archive_Summary headers:", headers);
                    console.log("Completeness Check - Date column index:", dateIdx, "Total column index:", totalIdx);

                    if (totalIdx >= 0 && dateIdx >= 0) {
                        const dataRows = archiveRange.values.slice(1);
                        archiveLookup = new Map();

                        // Build a lookup map from archive: normalize dates to YYYY-MM-DD for matching
                        // SUM all rows for each date (archive may have multiple rows per pay period)
                        for (const row of dataRows) {
                            const rawDate = row[dateIdx];
                            const normalizedKey = normalizeDateForLookup(rawDate);
                            if (normalizedKey) {
                                const amount = Number(row[totalIdx]) || 0;
                                const existing = archiveLookup.get(normalizedKey) || 0;
                                archiveLookup.set(normalizedKey, existing + amount);
                            }
                        }

                        if (archiveLookup.size > 0) {
                            archiveSource = "PR_Archive_Summary";
                            console.log("Completeness Check - Using PR_Archive_Summary lookup keys:", Array.from(archiveLookup.keys()));
                            console.log("Completeness Check - PR_Archive_Summary lookup values:", Array.from(archiveLookup.entries()));
                        }
                    } else {
                        console.warn("Completeness Check - Missing date or total column in PR_Archive_Summary");
                        console.warn("  Date column index:", dateIdx, "Total column index:", totalIdx);
                    }
                }
            }

            if (archiveSource && historicalPeriods.length > 0) {
                // Match each historical period against archive
                let archiveSum = 0;
                let periodsSum = 0;
                let matchedCount = 0;
                const periodDetails = [];

                for (const period of historicalPeriods) {
                    const periodKey = period.key || period.label || "";
                    const normalizedPeriodKey = normalizeDateForLookup(periodKey);
                    const periodTotal = period.summary?.total || 0;
                    periodsSum += periodTotal;

                    const archiveTotal = archiveLookup.get(normalizedPeriodKey);
                    if (archiveTotal !== undefined) {
                        archiveSum += archiveTotal;
                        matchedCount++;
                        periodDetails.push({
                            period: periodKey,
                            calculated: periodTotal,
                            archive: archiveTotal,
                            match: Math.abs(periodTotal - archiveTotal) < 1
                        });
                    } else {
                        console.warn(`Completeness Check - Period ${periodKey} (normalized: ${normalizedPeriodKey}) not found in archive`);
                        periodDetails.push({
                            period: periodKey,
                            calculated: periodTotal,
                            archive: null,
                            match: false
                        });
                    }
                }

                console.log("Completeness Check - Period details:", periodDetails);
                console.log("Completeness Check - Matched", matchedCount, "of", historicalPeriods.length, "periods");
                console.log("Completeness Check - Archive sum:", archiveSum, "Periods sum:", periodsSum);

                const allMatched = matchedCount === historicalPeriods.length && historicalPeriods.length > 0;
                const totalsMatch = Math.abs(archiveSum - periodsSum) < 1;
                const match = allMatched && totalsMatch;

                results.historicalPeriods = {
                    match,
                    archiveSum,
                    periodsSum,
                    matchedCount,
                    totalPeriods: historicalPeriods.length,
                    details: periodDetails,
                    source: archiveSource
                };
            }
            
            // Update state with results
            expenseReviewState.completenessCheck = results;
            console.log("Completeness Check - Results:", JSON.stringify(results));
        });
        console.log("Completeness Check - Complete!");
    } catch (error) {
        console.error("Payroll completeness check failed:", error);
    }
}

/**
 * Render Data Completeness Check card for Expense Review step
 * Shows detailed comparison with difference amounts
 */
function renderPayrollCompletenessCard() {
    const check = expenseReviewState.completenessCheck || {};
    const hasRun = expenseReviewState.periods?.length > 0;
    
    // Helper to format currency
    const fmt = (val) => formatNumberDisplay(Math.round(val || 0));
    
    // Helper to format difference with sign
    const fmtDiff = (diff) => {
        const absDiff = Math.abs(diff);
        if (absDiff < 1) return "—";
        const sign = diff > 0 ? "+" : "-";
        return `${sign}$${Math.round(absDiff).toLocaleString()}`;
    };
    
    // Render a comparison row with source values and difference
    const renderComparisonRow = (label, sourceLabel, sourceVal, calcLabel, calcVal, isMatch, isPending) => {
        const diff = (sourceVal || 0) - (calcVal || 0);
        
        let statusIcon;
        let statusClass;
        if (isPending) {
            statusIcon = `<span class="pf-complete-status pf-complete-status--pending" aria-hidden="true">${INFO_CIRCLE_SVG}</span>`;
            statusClass = "pending";
        } else if (isMatch) {
            statusIcon = `<span class="pf-complete-status pf-complete-status--pass" aria-hidden="true">${CHECK_CIRCLE_SVG}</span>`;
            statusClass = "pass";
        } else {
            statusIcon = `<span class="pf-complete-status pf-complete-status--fail" aria-hidden="true">${X_CIRCLE_SVG}</span>`;
            statusClass = "fail";
        }
        
        const diffDisplay = isPending ? "" : `
            <div class="pf-complete-diff ${statusClass}">
                ${fmtDiff(diff)}
            </div>
        `;
        
        return `
            <div class="pf-complete-row ${statusClass}">
                <div class="pf-complete-header">
                    ${statusIcon}
                    <span class="pf-complete-label">${escapeHtml(label)}</span>
                </div>
                ${!isPending ? `
                <div class="pf-complete-values">
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${escapeHtml(sourceLabel)}:</span>
                        <span class="pf-complete-amount">${fmt(sourceVal)}</span>
                    </div>
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${escapeHtml(calcLabel)}:</span>
                        <span class="pf-complete-amount">${fmt(calcVal)}</span>
                    </div>
                </div>
                ${diffDisplay}
                ` : `
                <div class="pf-complete-values">
                    <span class="pf-complete-pending-text">Click Run/Refresh to check</span>
                </div>
                `}
            </div>
        `;
    };
    
    // Current period check
    const currentResult = check.currentPeriod;
    const currentPending = !hasRun || currentResult === null || currentResult === undefined;
    const currentRow = renderComparisonRow(
        "Current Period",
        "PR_Data_Clean Total",
        currentResult?.prDataClean,
        "Calculated Total",
        currentResult?.currentTotal,
        currentResult?.match,
        currentPending
    );
    
    // Historical periods check - with matched count info
    const histResult = check.historicalPeriods;
    const histPending = !hasRun || histResult === null || histResult === undefined;
    const matchedCount = histResult?.matchedCount || 0;
    const totalPeriods = histResult?.totalPeriods || 0;
    const archiveSourceLabel = histResult?.source === "PR_Archive_Totals" ? "PR_Archive_Totals (matched)" : "PR_Archive_Summary (matched)";
    const histLabel = totalPeriods > 0 
        ? `Historical Periods (${matchedCount}/${totalPeriods} matched)`
        : "Historical Periods";
    const histRow = renderComparisonRow(
        histLabel,
        archiveSourceLabel,
        histResult?.archiveSum,
        "Calculated Total",
        histResult?.periodsSum,
        histResult?.match,
        histPending
    );
    
    // Build period details if available (for debugging/expanded view)
    let periodDetailsHtml = "";
    if (!histPending && histResult?.details?.length > 0) {
        const detailRows = histResult.details.map(d => {
            const matchIcon = d.archive === null ? ALERT_TRIANGLE_SVG : (d.match ? CHECK_CIRCLE_SVG : X_CIRCLE_SVG);
            const archiveVal = d.archive !== null ? fmt(d.archive) : "Not found";
            return `
                <div class="pf-complete-detail-row">
                    <span class="pf-complete-detail-date">${escapeHtml(d.period)}</span>
                    <span class="pf-complete-detail-icon" aria-hidden="true">${matchIcon}</span>
                    <span class="pf-complete-detail-vals">${fmt(d.calculated)} vs ${archiveVal}</span>
                </div>
            `;
        }).join("");
        periodDetailsHtml = `
            <div class="pf-complete-details-section">
                <div class="pf-complete-details-header">Period-by-Period Match</div>
                ${detailRows}
            </div>
        `;
    }
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-completeness-card">
            <div class="pf-config-head">
                <h3>Data Completeness Check</h3>
                <p class="pf-config-subtext">Verify source data matches calculated totals</p>
            </div>
            <div class="pf-complete-container">
                ${currentRow}
                ${histRow}
                ${periodDetailsHtml}
            </div>
        </article>
    `;
}

// =============================================================================
// ADA INSIGHTS PANEL - Expense Review AI Analysis
// =============================================================================

/**
 * Ada Insights prompt template for Expense Review
 * Provides structured instructions for generating executive-ready analysis
 */
const ADA_EXPENSE_REVIEW_SYSTEM_PROMPT = `You are Ada, Prairie Forge's AI financial analyst. You're embedded in the Payroll Recorder module, helping accountants and CFOs review payroll expenses.

CRITICAL: You MUST use ONLY the data provided in the context pack. Never invent numbers.

OUTPUT FORMAT (always follow this structure):

## Executive Summary
- 3 bullet points summarizing the key findings
- Focus on what executives need to know

## Key Drivers
- Ranked list of top 5-8 factors driving the current period's results
- Include dollar amounts and percentages where relevant
- Reference specific measures/buckets by name

## Roster/Headcount Commentary
- Comment on new hires, departures, and department changes
- Use cautious language:
  • For missing employees: "not seen this period (may be no-hours)"
  • For new employees: "first seen this period (likely new hire)"
- Only name individuals if they appear in the roster delta lists or drive large variances

## Risks & Anomalies
- Flag any concerning patterns, mismatches, or areas needing attention
- Highlight unclassified amounts or reconciliation gaps
- Note any data quality issues

## Recommended Next Actions
- 3-6 specific, actionable items
- Prioritize by impact

STYLE RULES:
- Be concise but thorough
- Use plain text status labels (no emoji)
- Format currency as $X,XXX
- Format percentages as X.X%
- Use bullet points, not paragraphs
- Assume the reader is a finance professional`;

/**
 * Generate Ada Insights for Expense Review
 * Builds context pack and calls Ada API
 */
async function generateAdaInsights(userPrompt = null) {
    console.log("[AdaInsights] Generating insights...");
    
    // Update state
    adaInsightsState.loading = true;
    adaInsightsState.lastError = null;
    renderApp();
    
    try {
        // Build context pack
        const contextPack = await buildExpenseReviewContextPack();
        adaInsightsState.contextPack = contextPack;
        
        // Check for critical errors
        if (!contextPack.availability.has_pr_data_clean) {
            throw new Error("PR_Data_Clean is required. Run Create Matrix first.");
        }
        
        // Build the prompt
        const defaultPrompt = "Analyze this payroll period and provide an executive-ready summary. Focus on key drivers, headcount changes, and any risks or anomalies.";
        const prompt = userPrompt || defaultPrompt;
        
        // Format context for Ada
        const contextString = formatContextForAda(contextPack);
        
        // Call Ada API - uses database-configured prompt for payroll-recorder module
        const response = await callAdaApi({
            userPrompt: `${prompt}\n\n--- CONTEXT DATA (use ONLY this data) ---\n${contextString}`,
            contextPack: contextPack,
            functionContext: "analysis"
        });
        
        // Update state with response
        adaInsightsState.response = response;
        adaInsightsState.lastRefresh = new Date().toISOString();
        adaInsightsState.loading = false;
        
        console.log("[AdaInsights] Insights generated successfully");
        renderApp();
        
    } catch (error) {
        console.error("[AdaInsights] Error generating insights:", error);
        adaInsightsState.loading = false;
        adaInsightsState.lastError = error.message;
        renderApp();
    }
}

/**
 * Format context pack as structured text for Ada
 */
function formatContextForAda(contextPack) {
    const fmt = (val) => typeof val === 'number' ? `$${val.toLocaleString(undefined, { minimumFractionDigits: 2 })}` : val;
    const fmtPct = (val) => typeof val === 'number' ? `${(val * 100).toFixed(1)}%` : val;
    
    let text = "";
    
    // Period info
    text += `PERIOD:\n`;
    text += `  Current: ${contextPack.period.current_key || "Unknown"}\n`;
    text += `  Prior: ${contextPack.period.prior_key || "Not available"}\n`;
    text += `  Basis: ${contextPack.identity.basis_mode}\n\n`;
    
    // Totals
    text += `TOTALS:\n`;
    text += `  Expense Review Current: ${fmt(contextPack.totals.expense_review_total_current)}\n`;
    text += `  Expense Review Prior: ${fmt(contextPack.totals.expense_review_total_prior)}\n`;
    text += `  Period-over-Period Change: ${fmt(contextPack.totals.expense_review_total_current - contextPack.totals.expense_review_total_prior)}\n`;
    if (contextPack.totals.bank_statement_amount) {
        text += `  Bank Statement: ${fmt(contextPack.totals.bank_statement_amount)}\n`;
        text += `  Bank Delta: ${fmt(contextPack.totals.bank_delta)}\n`;
    }
    text += `\n`;
    
    // Buckets
    text += `EXPENSE BUCKETS (Current vs Prior):\n`;
    text += `  FIXED: ${fmt(contextPack.drivers.bucket_totals_current.FIXED)} vs ${fmt(contextPack.drivers.bucket_totals_prior.FIXED)} (Δ ${fmt(contextPack.drivers.bucket_deltas.FIXED)})\n`;
    text += `  VARIABLE: ${fmt(contextPack.drivers.bucket_totals_current.VARIABLE)} vs ${fmt(contextPack.drivers.bucket_totals_prior.VARIABLE)} (Δ ${fmt(contextPack.drivers.bucket_deltas.VARIABLE)})\n`;
    text += `  BURDEN: ${fmt(contextPack.drivers.bucket_totals_current.BURDEN)} vs ${fmt(contextPack.drivers.bucket_totals_prior.BURDEN)} (Δ ${fmt(contextPack.drivers.bucket_deltas.BURDEN)})\n`;
    text += `\n`;
    
    // Top measure deltas
    if (contextPack.drivers.top_measure_deltas?.length > 0) {
        text += `TOP MEASURE CHANGES (by dollar impact):\n`;
        contextPack.drivers.top_measure_deltas.slice(0, 10).forEach((m, i) => {
            text += `  ${i + 1}. ${m.pf_column_name}: ${fmt(m.current_amount)} vs ${fmt(m.prior_amount)} (Δ ${fmt(m.delta_amount)}) [${m.bucket}]\n`;
        });
        text += `\n`;
    }
    
    // Department deltas
    if (contextPack.drivers.department_deltas?.length > 0) {
        text += `DEPARTMENT CHANGES (by dollar impact):\n`;
        contextPack.drivers.department_deltas.slice(0, 8).forEach((d, i) => {
            text += `  ${i + 1}. ${d.department_name}: ${fmt(d.current_amount)} vs ${fmt(d.prior_amount)} (Δ ${fmt(d.delta_amount)})\n`;
        });
        text += `\n`;
    }
    
    // Roster context
    const roster = contextPack.roster_context;
    if (roster && !roster.error) {
        text += `ROSTER CHANGES:\n`;
        text += `  Join Key: ${roster.join_key_used}\n`;
        
        if (roster.roster_new_this_period?.length > 0) {
            text += `  New This Period (${roster.roster_new_this_period.length}):\n`;
            roster.roster_new_this_period.slice(0, 10).forEach(emp => {
                text += `    - ${emp.name} (${emp.department}) - ${emp.note}\n`;
            });
        }
        
        if (roster.roster_missing_this_period?.length > 0) {
            text += `  Not Seen This Period (${roster.roster_missing_this_period.length}):\n`;
            roster.roster_missing_this_period.slice(0, 10).forEach(emp => {
                text += `    - ${emp.name} (${emp.department}) - ${emp.note}\n`;
            });
        }
        
        if (roster.roster_reactivated?.length > 0) {
            text += `  Reactivations (${roster.roster_reactivated.length}):\n`;
            roster.roster_reactivated.forEach(emp => {
                text += `    - ${emp.name} (${emp.department}) - ${emp.note}\n`;
            });
        }
        
        if (roster.roster_department_changes?.length > 0) {
            text += `  Department Changes (${roster.roster_department_changes.length}):\n`;
            roster.roster_department_changes.forEach(emp => {
                text += `    - ${emp.name}: ${emp.previous_department} to ${emp.current_department}\n`;
            });
        }
        
        // Headcount bridge
        const depts = Object.keys(roster.headcount_delta_by_department || {});
        if (depts.length > 0) {
            text += `\n  HEADCOUNT BY DEPARTMENT:\n`;
            depts.forEach(dept => {
                const current = roster.headcount_by_department_current[dept] || 0;
                const prior = roster.headcount_by_department_prior[dept] || 0;
                const delta = roster.headcount_delta_by_department[dept] || 0;
                const newHires = roster.new_hires_by_department[dept] || 0;
                const missing = roster.missing_by_department[dept] || 0;
                text += `    ${dept}: ${current} (prior: ${prior}, Δ ${delta >= 0 ? '+' : ''}${delta})`;
                if (newHires > 0) text += ` [+${newHires} new]`;
                if (missing > 0) text += ` [-${missing} missing]`;
                text += `\n`;
            });
        }
        text += `\n`;
    } else if (roster?.error) {
        text += `ROSTER: ${roster.error}\n\n`;
    }
    
    // Data quality / exclusions
    if (contextPack.delta_breakdown.top_excluded_by_dollars?.length > 0) {
        text += `EXCLUDED FROM TOTALS (top by dollars):\n`;
        contextPack.delta_breakdown.top_excluded_by_dollars.slice(0, 5).forEach(e => {
            text += `  - ${e.header}: ${fmt(e.total)} (${e.reason})\n`;
        });
        text += `\n`;
    }
    
    if (contextPack.delta_breakdown.total_unclassified > 0) {
        text += `UNCLASSIFIED COLUMNS: ${fmt(contextPack.delta_breakdown.total_unclassified)} total\n`;
        if (contextPack.metadata.measures_missing_dictionary_metadata?.length > 0) {
            text += `  Columns: ${contextPack.metadata.measures_missing_dictionary_metadata.slice(0, 5).join(", ")}\n`;
        }
        text += `\n`;
    }
    
    // Availability
    text += `DATA AVAILABILITY:\n`;
    text += `  PR_Data_Clean: ${contextPack.availability.has_pr_data_clean ? "Yes" : "No"}\n`;
    text += `  Archive Totals: ${contextPack.availability.has_archive_totals ? "Yes" : "No"}\n`;
    text += `  Archive Summary: ${contextPack.availability.has_archive_summary ? "Yes" : "No"}\n`;
    text += `  Employee Roster: ${contextPack.availability.has_employee_roster ? "Yes" : "No"}\n`;
    text += `  Prior Period: ${contextPack.availability.has_prior_period ? "Yes" : "No"}\n`;
    
    if (contextPack.availability.error_messages?.length > 0) {
        text += `\nERRORS/WARNINGS:\n`;
        contextPack.availability.error_messages.forEach(msg => {
            text += `  - ${msg}\n`;
        });
    }
    
    return text;
}

/**
 * Call Ada API endpoint (Supabase copilot edge function)
 * Can be called with either:
 * - Object params: { systemPrompt, userPrompt, contextPack, functionContext }
 * - Positional params: (prompt, context, messageHistory) - for copilot.js compatibility
 */
async function callAdaApi(promptOrParams, context, messageHistory) {
    // Handle both calling conventions
    let systemPrompt, userPrompt, contextPack, functionContext;
    
    if (typeof promptOrParams === 'object' && promptOrParams !== null && !Array.isArray(promptOrParams) && promptOrParams.userPrompt !== undefined) {
        // Called with object params (original style)
        ({ systemPrompt, userPrompt, contextPack, functionContext } = promptOrParams);
    } else {
        // Called with positional params (copilot.js style)
        userPrompt = promptOrParams;
        contextPack = context;
        functionContext = "analysis";
    }
    
    // Supabase copilot endpoint
    const COPILOT_URL = "https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1/copilot";
    const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE";
    
    // Get customer ID from workbook config (secure identifier for logging/analytics)
    const customerId = getConfigValue("SS_Company_ID") || null;
    
    try {
        console.log("[AdaInsights] Calling copilot API...", { module: "payroll-recorder", function: functionContext || "analysis", customerId: customerId ? "set" : "not set" });
        
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
                prompt: userPrompt,
                context: contextPack,
                module: "payroll-recorder",
                function: functionContext || "analysis",
                customerId: customerId,
                // Only pass systemPrompt if we want to override the database config
                ...(systemPrompt ? { systemPrompt } : {})
            })
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error("[AdaInsights] API error:", response.status, errorText);
            throw new Error(`API request failed: ${response.status}`);
        }
        
        const data = await response.json();
        console.log("[AdaInsights] API response received:", data.usage || "no usage info");
        
        if (data.message || data.response) {
            return data.message || data.response;
        }
        
        // Fallback if no message in response
        console.warn("[AdaInsights] No message in API response, using local generation");
        return generateLocalInsightsSummary(params);
        
    } catch (error) {
        console.warn("[AdaInsights] API call failed, using local generation:", error);
        return generateLocalInsightsSummary(params);
    }
}

/**
 * Generate a local insights summary when API is unavailable
 * This provides a structured analysis based on the context pack
 */
function generateLocalInsightsSummary(params) {
    const contextPack = adaInsightsState.contextPack;
    if (!contextPack) return "Unable to generate insights - no context available.";
    
    const fmt = (val) => typeof val === 'number' ? `$${val.toLocaleString(undefined, { minimumFractionDigits: 0 })}` : val;
    
    let summary = "";
    
    // Executive Summary
    const currentTotal = contextPack.totals.expense_review_total_current || 0;
    const priorTotal = contextPack.totals.expense_review_total_prior || 0;
    const periodDelta = currentTotal - priorTotal;
    const periodPctChange = priorTotal ? ((periodDelta / priorTotal) * 100).toFixed(1) : 0;
    
    summary += `## Executive Summary\n\n`;
    summary += `- **Current period total: ${fmt(currentTotal)}** `;
    if (priorTotal > 0) {
        summary += `(${periodDelta >= 0 ? '+' : ''}${fmt(periodDelta)} / ${periodPctChange}% vs prior)\n`;
    } else {
        summary += `(no prior period for comparison)\n`;
    }
    
    // Check roster for headcount context
    const roster = contextPack.roster_context;
    const newCount = roster?.roster_new_this_period?.length || 0;
    const missingCount = roster?.roster_missing_this_period?.length || 0;
    if (newCount > 0 || missingCount > 0) {
        summary += `- Headcount changes: ${newCount} new employee${newCount !== 1 ? 's' : ''}, ${missingCount} not seen this period\n`;
    }
    
    // Check data quality
    const unclassifiedTotal = contextPack.delta_breakdown.total_unclassified || 0;
    if (unclassifiedTotal > 1000) {
        summary += `- ${fmt(unclassifiedTotal)} in unclassified amounts requires review\n`;
    } else {
        summary += `- All amounts properly classified\n`;
    }
    
    // Key Drivers
    summary += `\n## Key Drivers\n\n`;
    
    // Bucket analysis
    const buckets = contextPack.drivers.bucket_deltas;
    const sortedBuckets = Object.entries(buckets)
        .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
    
    sortedBuckets.forEach(([bucket, delta]) => {
        const current = contextPack.drivers.bucket_totals_current[bucket] || 0;
        if (current > 0) {
            summary += `- **${bucket}**: ${fmt(current)}`;
            if (Math.abs(delta) > 100) {
                summary += ` (${delta >= 0 ? '+' : ''}${fmt(delta)} vs prior)`;
            }
            summary += `\n`;
        }
    });
    
    // Top measure changes
    const topMeasures = (contextPack.drivers.top_measure_deltas || [])
        .filter(m => Math.abs(m.delta_amount) > 500)
        .slice(0, 5);
    
    if (topMeasures.length > 0) {
        summary += `\n**Top changing measures:**\n`;
        topMeasures.forEach(m => {
            summary += `- ${m.pf_column_name}: ${fmt(m.delta_amount)} change\n`;
        });
    }
    
    // Roster Commentary
    summary += `\n## Roster/Headcount Commentary\n\n`;
    
    if (roster && !roster.error) {
        if (newCount > 0) {
            summary += `**New this period (${newCount}):**\n`;
            roster.roster_new_this_period.slice(0, 5).forEach(emp => {
                summary += `- ${emp.name} (${emp.department}) - first seen this period (likely new hire)\n`;
            });
            if (newCount > 5) summary += `- ...and ${newCount - 5} more\n`;
        }
        
        if (missingCount > 0) {
            summary += `\n**Not seen this period (${missingCount}):**\n`;
            roster.roster_missing_this_period.slice(0, 5).forEach(emp => {
                summary += `- ${emp.name} (${emp.department}) - not seen this period (may be no-hours)\n`;
            });
            if (missingCount > 5) summary += `- ...and ${missingCount - 5} more\n`;
        }
        
        if (roster.roster_reactivated?.length > 0) {
            summary += `\n**Reactivations:**\n`;
            roster.roster_reactivated.forEach(emp => {
                summary += `- ${emp.name} - was terminated, now appearing in payroll\n`;
            });
        }
        
        if (newCount === 0 && missingCount === 0 && !roster.roster_reactivated?.length) {
            summary += `No significant headcount changes detected\n`;
        }
    } else {
        summary += `Roster data unavailable - ${roster?.error || "unknown error"}\n`;
    }
    
    // Risks & Anomalies
    summary += `\n## Risks & Anomalies\n\n`;
    
    let hasRisks = false;
    
    if (unclassifiedTotal > 1000) {
        summary += `- **Unclassified amounts**: ${fmt(unclassifiedTotal)} - review column classifications\n`;
        hasRisks = true;
    }
    
    const bankDelta = contextPack.totals.bank_delta;
    if (bankDelta && Math.abs(bankDelta) > 0.01) {
        summary += `- **Bank reconciliation gap**: ${fmt(bankDelta)}\n`;
        hasRisks = true;
    }
    
    if (Math.abs(periodPctChange) > 15) {
        summary += `- **Large period variance**: ${periodPctChange}% change vs prior period\n`;
        hasRisks = true;
    }
    
    if (!hasRisks) {
        summary += `No significant risks or anomalies detected\n`;
    }
    
    // Recommended Actions
    summary += `\n## Recommended Next Actions\n\n`;
    
    let actionNum = 1;
    
    if (unclassifiedTotal > 1000) {
        summary += `${actionNum++}. Review and classify the ${contextPack.metadata.measures_missing_dictionary_metadata?.length || 0} unclassified columns\n`;
    }
    
    if (missingCount > 0) {
        summary += `${actionNum++}. Verify status of ${missingCount} employees not seen this period\n`;
    }
    
    if (roster?.roster_reactivated?.length > 0) {
        summary += `${actionNum++}. Review ${roster.roster_reactivated.length} reactivation(s) - confirm employment status\n`;
    }
    
    if (Math.abs(periodPctChange) > 10 && priorTotal > 0) {
        summary += `${actionNum++}. Investigate ${Math.abs(periodPctChange).toFixed(1)}% period-over-period variance\n`;
    }
    
    if (actionNum === 1) {
        summary += `1. Review expense summary and proceed to Journal Entry Prep\n`;
        summary += `2. Archive current period data\n`;
    }
    
    summary += `${actionNum}. Export journal entry for upload to accounting system\n`;
    
    return summary;
}

/**
 * Render the Ada Insights panel for Expense Review
 */
function renderAdaInsightsPanel() {
    const { loading, lastError, response, lastRefresh, collapsed } = adaInsightsState;
    
    const collapseIcon = collapsed 
        ? `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m6 9 6 6 6-6"/></svg>`
        : `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m18 15-6-6-6 6"/></svg>`;
    
    const statusDot = loading 
        ? `<span class="pf-ada-insights-dot pf-ada-insights-dot--busy"></span>`
        : lastError 
            ? `<span class="pf-ada-insights-dot pf-ada-insights-dot--error"></span>`
            : response 
                ? `<span class="pf-ada-insights-dot pf-ada-insights-dot--ready"></span>`
                : `<span class="pf-ada-insights-dot"></span>`;
    
    const refreshLabel = lastRefresh 
        ? `Last updated: ${new Date(lastRefresh).toLocaleTimeString()}`
        : "Not yet generated";
    
    // Quick prompt buttons
    const quickPrompts = [
        { id: "main", label: "Generate Insights", prompt: null },
        { id: "changes", label: "Biggest Changes", prompt: "Explain the biggest changes vs last period - what's driving the variance?" },
        { id: "headcount", label: "Headcount Impact", prompt: "What changed in headcount that could explain the expense changes?" },
        { id: "unreconciled", label: "Unreconciled", prompt: "Explain any unreconciled totals or unclassified amounts" },
        { id: "gl", label: "GL Gaps", prompt: "Are there any GL mapping gaps or classification issues to address?" }
    ];
    
    const quickButtonsHtml = quickPrompts.map(p => `
        <button type="button" class="pf-ada-quick-btn ${p.id === 'main' ? 'pf-ada-quick-btn--primary' : ''}" 
                data-ada-prompt="${p.id}" title="${p.label}">
            ${p.label}
        </button>
    `).join('');
    
    // Response content
    let contentHtml = "";
    if (loading) {
        contentHtml = `
            <div class="pf-ada-insights-loading">
                <div class="pf-ada-insights-spinner"></div>
                <p>Analyzing payroll data...</p>
            </div>
        `;
    } else if (lastError) {
        contentHtml = `
            <div class="pf-ada-insights-error">
                <p>${escapeHtml(lastError)}</p>
            </div>
        `;
    } else if (response) {
        // Parse markdown-ish response into HTML
        const formattedResponse = formatAdaResponse(response);
        contentHtml = `
            <div class="pf-ada-insights-response">
                ${formattedResponse}
            </div>
        `;
    } else {
        contentHtml = `
            <div class="pf-ada-insights-empty">
                <p>Click <strong>Generate Insights</strong> to analyze this payroll period.</p>
                <p class="pf-ada-insights-hint">Ada will provide an executive-ready summary of key drivers, headcount changes, and recommended actions.</p>
            </div>
        `;
    }
    
    return `
        <article class="pf-step-card pf-ada-insights-panel ${collapsed ? 'pf-ada-insights-panel--collapsed' : ''}" id="ada-insights-panel">
            <header class="pf-ada-insights-header">
                <div class="pf-ada-insights-title">
                    ${statusDot}
                    <img class="pf-ada-insights-avatar" src="${BRANDING?.ADA_IMAGE_URL || ''}" alt="Ada" onerror="this.style.display='none'" />
                    <div class="pf-ada-insights-name">
                        <span class="pf-ada-title"><span class="pf-ada-title--ask">ask</span><span class="pf-ada-title--ada">ADA</span></span>
                        <span class="pf-ada-insights-refresh">${refreshLabel}</span>
                    </div>
                </div>
                <button type="button" class="pf-ada-insights-collapse" id="ada-insights-collapse" title="Collapse">
                    ${collapseIcon}
                </button>
            </header>
            
            <div class="pf-ada-insights-body ${collapsed ? 'hidden' : ''}">
                <div class="pf-ada-quick-prompts">
                    ${quickButtonsHtml}
                </div>
                
                <div class="pf-ada-insights-content" id="ada-insights-content">
                    ${contentHtml}
                </div>
            </div>
        </article>
    `;
}

/**
 * Format Ada's response (markdown-like) into HTML
 */
function formatAdaResponse(text) {
    if (!text) return "";
    
    // Escape HTML first
    let html = escapeHtml(text);
    
    // Headers (## -> h3, ### -> h4)
    html = html.replace(/^### (.+)$/gm, '<h4 class="pf-ada-h4">$1</h4>');
    html = html.replace(/^## (.+)$/gm, '<h3 class="pf-ada-h3">$1</h3>');
    
    // Bold (**text**)
    html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    
    // Lists (- item)
    html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul class="pf-ada-list">$&</ul>');
    
    // Numbered lists (1. item)
    html = html.replace(/^\d+\. (.+)$/gm, '<li>$1</li>');
    
    // Paragraphs (double newline)
    html = html.replace(/\n\n/g, '</p><p>');
    html = `<p>${html}</p>`;
    
    // Clean up empty paragraphs
    html = html.replace(/<p>\s*<\/p>/g, '');
    html = html.replace(/<p>\s*(<h[34])/g, '$1');
    html = html.replace(/(<\/h[34]>)\s*<\/p>/g, '$1');
    html = html.replace(/<p>\s*(<ul)/g, '$1');
    html = html.replace(/(<\/ul>)\s*<\/p>/g, '$1');
    
    return html;
}

/**
 * Get step-specific info panel configuration for Payroll Recorder
 */
function getStepInfoConfig(stepId) {
    switch (stepId) {
        case 0:
            return {
                title: "Configuration",
                content: `
                    <div class="pf-info-section">
                        <h4>What This Step Does</h4>
                        <p>Sets up the key parameters for your payroll review. Complete this before importing data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Key Fields</h4>
                        <ul>
                            <li><strong>Payroll Date</strong> — The period-end date for this payroll run</li>
                            <li><strong>Accounting Period</strong> — Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> — Reference number for your accounting system</li>
                            <li><strong>Provider Link</strong> — Quick access to your payroll provider portal</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>Tip</h4>
                        <p>The accounting period and JE ID auto-generate based on your payroll date, but you can override them if needed.</p>
                    </div>
                `
            };
        case 1:
            return {
                title: "Import Payroll Data",
                content: `
                    <div class="pf-info-section">
                        <h4>What This Step Does</h4>
                        <p>Upload your payroll export, map columns, create the data matrix, and reconcile to your bank statement.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Column Status Guide</h4>
                        <ul>
                            <li><strong>Saved</strong> — Previously confirmed mapping, reused automatically</li>
                            <li><strong>Auto</strong> — Matched from known patterns</li>
                            <li><strong>Review</strong> — Best guess, please verify</li>
                            <li><strong>Select</strong> — No match found, choose manually</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7); margin-top: 8px;">Nothing is saved until you click "Create Matrix".</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Bank Reconciliation</h4>
                        <p>After creating the matrix, compare your system total to the bank statement amount:</p>
                        <ul>
                            <li><strong>System Total</strong> — Sum of amount columns in PR_Data_Clean</li>
                            <li><strong>Bank Amount</strong> — Enter what actually left the bank</li>
                            <li><strong>Difference</strong> — Should be $0.00 (within $0.01 tolerance)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>Tip</h4>
                        <p>Column headers don't need to match exactly—the system is flexible with naming. Just make sure each field is present.</p>
                    </div>
                `
            };
        case 2:
            return {
                title: "Expense Review",
                content: `
                    <div class="pf-info-section">
                        <h4>What This Step Does</h4>
                        <p>Generates an executive-ready payroll expense summary for CFO review, with period comparisons and trend analysis.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>Data Sources</h4>
                        <ul>
                            <li><strong>PR_Data_Clean</strong> — Current period payroll data (cleaned and categorized)</li>
                            <li><strong>SS_Employee_Roster</strong> — Department assignments and employee details</li>
                            <li><strong>PR_Archive_Summary</strong> — Historical payroll data for trend analysis</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>How Amounts Are Calculated</h4>
                        <table style="width:100%; font-size: 11px; margin-top: 8px; border-collapse: collapse;">
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Fixed Salary</strong></td>
                                <td style="padding: 6px 0;">Regular wages, salaries, and base pay</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Variable Salary</strong></td>
                                <td style="padding: 6px 0;">Commissions, bonuses, overtime, and incentive pay</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Gross Pay</strong></td>
                                <td style="padding: 6px 0;">Fixed + Variable Salary</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Burden</strong></td>
                                <td style="padding: 6px 0;">Employer taxes (FICA, Medicare, FUTA, SUTA), health insurance, 401(k) match, and other employer-paid benefits</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>All-In Total</strong></td>
                                <td style="padding: 6px 0;">Gross Pay + Burden = Total cost to employer</td>
                            </tr>
                            <tr>
                                <td style="padding: 6px 0;"><strong>Burden Rate</strong></td>
                                <td style="padding: 6px 0;">Burden ÷ All-In Total (typically 10-18%)</td>
                            </tr>
                        </table>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>Report Sections</h4>
                        <ul>
                            <li><strong>Executive Summary</strong> — Current vs prior period comparison (frozen at top)</li>
                            <li><strong>Department Breakdown</strong> — Cost allocation by cost center</li>
                            <li><strong>Historical Context</strong> — Where current metrics fall within historical ranges</li>
                            <li><strong>Period Trends</strong> — 6-period trend chart for Total, Fixed, Variable, Burden, and Headcount</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>Historical Context Visualization</h4>
                        <p>The spectrum bars show where your current period falls relative to your historical min/max:</p>
                        <p style="font-family: Consolas, monospace; color: #6366f1; margin: 8px 0;">Current period is shown as a marker on a low-to-high range.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7);">Left = Low, Right = High.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>Review Tips</h4>
                        <ul>
                            <li>Compare <strong>Burden Rate</strong> — Should be consistent period-to-period (10-18% typical)</li>
                            <li>Watch <strong>Variable Salary</strong> spikes — May indicate commission/bonus timing</li>
                            <li>Verify <strong>Headcount changes</strong> — Should align with HR records</li>
                            <li>Flag variances <strong>> 10%</strong> from prior period for follow-up</li>
                        </ul>
                    </div>
                `
            };
        case 3:
            return {
                title: "Journal Entry",
                content: `
                    <div class="pf-info-section">
                        <h4>What This Step Does</h4>
                        <p>Generates a balanced journal entry from your payroll data, ready for upload to your accounting system.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>How the JE Works</h4>
                        <p>Maps payroll categories to GL accounts:</p>
                        <ul>
                            <li><strong>Expenses</strong>: Debits to departmental expense accounts</li>
                            <li><strong>Liabilities</strong>: Credits to payable accounts</li>
                            <li><strong>Cash</strong>: Credit to bank account</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>Validation Checks</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> — Entry must balance</li>
                            <li><strong>All accounts mapped</strong> — No unassigned categories</li>
                            <li><strong>Totals match</strong> — JE ties to PR_Data_Clean</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>Tip</h4>
                        <p>Review the allocation in PR_JE_Draft before exporting. Unmapped rows need GL account assignment.</p>
                    </div>
                `
            };
        case 4:
            return {
                title: "Archive & Clear",
                content: `
                    <div class="pf-info-section">
                        <h4>What This Step Does</h4>
                        <p>Creates a backup of your completed payroll run, then resets the workbook so you're ready for the next pay period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Step 1: Create Backup</h4>
                        <p>A new workbook opens containing all your payroll tabs. You'll choose where to save it on your computer or shared drive.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Tip: Use a consistent naming convention like "Payroll_Archive_2024-01-15"</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Step 2: Update History</h4>
                        <p>The current period's totals are saved to PR_Archive_Summary. This powers the trend charts and completeness checks for future periods.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Keeps 5 periods of history — oldest is removed automatically</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Step 3: Clear Working Data</h4>
                        <p>Data is cleared from the working sheets:</p>
                        <ul>
                            <li>PR_Data_Clean (processed payroll data)</li>
                            <li>PR_Expense_Review (summary & charts)</li>
                            <li>PR_JE_Draft (journal entry lines)</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Headers are preserved — only data rows are cleared</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Step 4: Reset for Next Period</h4>
                        <ul>
                            <li>Payroll Date, Accounting Period, JE ID cleared</li>
                            <li>All sign-offs and completion flags reset</li>
                            <li>Notes cleared (unless you locked them)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>Before You Archive</h4>
                        <ul>
                            <li>JE uploaded to your accounting system</li>
                            <li>All review steps signed off</li>
                            <li>Lock any notes you want to keep</li>
                        </ul>
                    </div>
                `
            };
        default:
            return {
                title: "Payroll Recorder",
                content: `
                    <div class="pf-info-section">
                        <h4>Welcome to Payroll Recorder</h4>
                        <p>This module helps you normalize payroll exports, enforce controls, and prep journal entries.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>Workflow Overview</h4>
                        <ol style="margin: 8px 0; padding-left: 20px;">
                            <li>Configure period settings</li>
                            <li>Import payroll data</li>
                            <li>Review headcount alignment</li>
                            <li>Validate against bank</li>
                            <li>Review expense summary</li>
                            <li>Generate journal entry</li>
                            <li>Archive and reset</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <p>Click a step card to get started, or tap the <strong>Info</strong> button on any step for detailed guidance.</p>
                    </div>
                `
            };
    }
}

initializeOffice(() => init());

/**
 * Ensure SS_PF_Config sheet and table exist with proper structure
 * Creates the config sheet and table if they don't exist
 */
async function ensureConfigSheet() {
    if (!hasExcelRuntime()) return;
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();
            
            // Check if SS_PF_Config sheet exists
            let configSheet = worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) {
                console.log("[Payroll] Creating SS_PF_Config sheet...");
                configSheet = worksheets.add("SS_PF_Config");
                
                // Add headers
                const headers = ["Category", "Field", "Value", "Permanent"];
                const headerRange = configSheet.getRange("A1:D1");
                headerRange.values = [headers];
                formatSheetHeaders(headerRange);
                
                // Add default module-prefix rows
                const defaultData = [
                    ["module-prefix", "PR_", "payroll-recorder", "Y"],
                    ["module-prefix", "PTO_", "pto-accrual", "Y"],
                    ["module-prefix", "SS_", "system", "Y"],
                    ["Run Settings", "SS_Company_Name", "", "Y"],
                    ["Run Settings", "SS_Company_ID", "", "Y"]
                ];
                const dataRange = configSheet.getRange(`A2:D${1 + defaultData.length}`);
                dataRange.values = defaultData;
                
                await context.sync();
                
                // Create the table
                const tableRange = configSheet.getRange(`A1:D${1 + defaultData.length}`);
                const table = configSheet.tables.add(tableRange, true);
                table.name = "SS_PF_Config";
                table.style = "TableStyleMedium2";
                
                // Auto-fit columns
                tableRange.format.autofitColumns();
                
                await context.sync();
                console.log("[Payroll] SS_PF_Config sheet and table created");
            } else {
                // Sheet exists - check if table exists
                const tables = context.workbook.tables;
                tables.load("items/name");
                await context.sync();
                
                const hasConfigTable = tables.items.some(t => t.name === "SS_PF_Config");
                
                if (!hasConfigTable) {
                    console.log("[Payroll] SS_PF_Config sheet exists but no table - creating table...");
                    
                    // Get used range to determine table extent
                    const usedRange = configSheet.getUsedRangeOrNullObject();
                    usedRange.load("address,rowCount");
                    await context.sync();
                    
                    if (!usedRange.isNullObject && usedRange.rowCount > 0) {
                        const table = configSheet.tables.add(usedRange, true);
                        table.name = "SS_PF_Config";
                        table.style = "TableStyleMedium2";
                        await context.sync();
                        console.log("[Payroll] SS_PF_Config table created from existing data");
                    } else {
                        // Empty sheet - add headers and create table
                        const headers = ["Category", "Field", "Value", "Permanent"];
                        const headerRange = configSheet.getRange("A1:D1");
                        headerRange.values = [headers];
                        formatSheetHeaders(headerRange);
                        
                        const defaultData = [
                            ["module-prefix", "PR_", "payroll-recorder", "Y"],
                            ["module-prefix", "PTO_", "pto-accrual", "Y"],
                            ["module-prefix", "SS_", "system", "Y"]
                        ];
                        const dataRange = configSheet.getRange(`A2:D${1 + defaultData.length}`);
                        dataRange.values = defaultData;
                        
                        await context.sync();
                        
                        const tableRange = configSheet.getRange(`A1:D${1 + defaultData.length}`);
                        const table = configSheet.tables.add(tableRange, true);
                        table.name = "SS_PF_Config";
                        table.style = "TableStyleMedium2";
                        
                        await context.sync();
                        console.log("[Payroll] SS_PF_Config table created with defaults");
                    }
                }
            }
        });
    } catch (error) {
        console.error("[Payroll] Error ensuring config sheet:", error);
    }
}

async function init() {
    try {
        // Initialize global context for Ada (homepage by default)
        window.PRAIRIE_FORGE_CONTEXT.step = null;
        window.PRAIRIE_FORGE_CONTEXT.stepName = "Homepage";
        
        await ensureConfigSheet();
        await ensureTabVisibility();
        await loadConfigurationValues();
        
        // Update company ID in context after config loads
        window.PRAIRIE_FORGE_CONTEXT.companyId = getConfigValue("SS_Company_ID") || null;
        
        // NOTE: Bootstrap config sync is now handled GLOBALLY by module-selector.
        // SS_PF_Config values are already populated when this module loads.
        // See Common/bootstrap.js for the source of truth.
        
        // Always start at homepage (no route restoration)
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        
        // Save initial route state
        saveRouteState(MODULE_KEY, buildRouteString(MODULE_KEY, "home", null), {
            activeView: "home",
            activeStepId: null,
            focusedIndex: 0
        });
        
        renderApp();
    } catch (error) {
        console.error("[Payroll] Module initialization failed:", error);
        throw error;
    }
}

async function ensureTabVisibility() {
    // Apply prefix-based tab visibility
    // Shows PR_* tabs, hides PTO_* and SS_* tabs
    try {
        await applyModuleTabVisibility(MODULE_KEY);
        console.log(`[Payroll] Tab visibility applied for ${MODULE_KEY}`);
    } catch (error) {
        console.warn("[Payroll] Could not apply tab visibility:", error);
    }
}

function renderApp() {
    const root = document.body;
    if (!root) return;
    
    // Remove seasonal elements before re-rendering to prevent layering/blurriness
    const seasonalSelectors = ['.pf-holiday-snow', '.pf-holiday-lights'];
    seasonalSelectors.forEach(selector => {
        const elements = root.querySelectorAll(selector);
        elements.forEach(el => el.remove());
    });
    
    const prevDisabled = appState.focusedIndex <= 0 ? "disabled" : "";
    const nextDisabled = appState.focusedIndex >= WORKFLOW_STEPS.length - 1 ? "disabled" : "";
    const isConfigView = appState.activeView === "config";
    const isStepView = appState.activeView === "step";
    const isHomeView = !isConfigView && !isStepView;
    const viewMarkup = isConfigView
        ? renderConfigView()
        : isStepView
            ? renderStepView(appState.activeStepId)
            : renderHomeView();
    root.innerHTML = `
        <div class="pf-root">
            ${renderBanner(prevDisabled, nextDisabled)}
            ${viewMarkup}
            ${renderFooter()}
        </div>
    `;
    
    // Mount info FAB with step-specific content (only on step/config views, not homepage)
    const infoFabElement = document.getElementById("pf-info-fab-payroll");
    if (isHomeView) {
        // Remove info fab on homepage
        if (infoFabElement) infoFabElement.remove();
    } else if (window.PrairieForge?.mountInfoFab) {
        const infoConfig = getStepInfoConfig(appState.activeStepId);
        PrairieForge.mountInfoFab({ 
            title: infoConfig.title, 
            content: infoConfig.content, 
            buttonId: "pf-info-fab-payroll" 
        });
    }
    
    // Mount Quick Access Modal to document.body (outside pf-root for proper z-index)
    mountQuickAccessModal();
    
    bindSharedInteractions();
    if (isConfigView) {
        bindConfigInteractions();
    } else if (isStepView) {
        // Bind step cards in sidebar (for navigating between steps)
        bindHomeInteractions();
        try {
            bindStepInteractions(appState.activeStepId);
        } catch (error) {
            console.warn("Payroll Recorder: failed to bind step interactions", error);
        }
    } else {
        bindHomeInteractions();
    }
    scrollFocusedIntoView();
    
    // Show/hide Ada FAB based on view
    if (isHomeView) {
        renderAdaFab();
    } else {
        removeAdaFab();
    }
}

function renderBanner(prevDisabled, nextDisabled) {
    return `
        <div class="pf-brand-float" aria-hidden="true">
            <span class="pf-brand-wave"></span>
        </div>
        <header class="pf-banner">
            <div class="pf-nav-bar">
                <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${prevDisabled}>
                    ${ARROW_LEFT_SVG}
                    <span class="sr-only">Previous step</span>
                </button>
                <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                    ${HOME_ICON_SVG}
                    <span class="sr-only">Module Home</span>
                </button>
                <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                    ${MODULES_ICON_SVG}
                    <span class="sr-only">Module Selector</span>
                </button>
                <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${nextDisabled}>
                    ${ARROW_RIGHT_SVG}
                    <span class="sr-only">Next step</span>
                </button>
                <div class="pf-nav-divider"></div>
                    <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                        ${MENU_ICON_SVG}
                        <span class="sr-only">Quick Access Menu</span>
                    </button>
            </div>
        </header>
    `;
}

/**
 * Mount Quick Access Modal directly to document.body (outside pf-root stacking context)
 * This ensures the modal appears above all other content regardless of z-index conflicts
 */
function mountQuickAccessModal() {
    // Remove existing modal if present
    const existing = document.getElementById("quick-access-modal");
    if (existing) existing.remove();
    
    const providerLink = getPayrollProviderLink();
    const hasProviderLink = !!providerLink;
    
    const modal = document.createElement("div");
    modal.id = "quick-access-modal";
    modal.className = "pf-quick-modal hidden";
    modal.innerHTML = `
        <div class="pf-quick-modal-backdrop" data-close></div>
        <div class="pf-quick-modal-card">
            <div class="pf-quick-modal-header">
                <h3 class="pf-quick-modal-title">Quick Access</h3>
                <button id="quick-access-close" class="pf-quick-modal-close pf-clickable" type="button" aria-label="Close">
                    ${X_ICON_SVG}
                </button>
            </div>
            <div class="pf-quick-modal-items">
                ${hasProviderLink ? `
                <a id="nav-provider-link" class="pf-quick-modal-item pf-clickable" href="${escapeHtml(providerLink)}" target="_blank" rel="noopener">
                    ${FILE_TEXT_ICON_SVG}
                    <span>Payroll Provider Report</span>
                </a>
                ` : ''}
                <button id="nav-accounting-software" class="pf-quick-modal-item pf-clickable" type="button">
                    ${UPLOAD_ICON_SVG}
                    <span>Accounting Software</span>
                </button>
                <button id="nav-employee-roster" class="pf-quick-modal-item pf-clickable" type="button">
                    ${USERS_ICON_SVG}
                    <span>Employee Roster</span>
                </button>
                <button id="nav-chart-of-accounts" class="pf-quick-modal-item pf-clickable" type="button">
                    ${BOOK_ICON_SVG}
                    <span>Chart of Accounts</span>
                </button>
                <button id="nav-config" class="pf-quick-modal-item pf-clickable" type="button">
                    ${SETTINGS_ICON_SVG}
                    <span>Configuration</span>
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
}

function renderHomeView() {
    return `
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">Payroll Recorder</h2>
            <p class="pf-hero-copy">${HERO_COPY}</p>
            <p class="pf-hero-hint">${escapeHtml(appState.statusText || "")}</p>
        </section>
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${WORKFLOW_STEPS.map((step, index) => renderStepCard(step, index)).join("")}
            </div>
        </section>
    `;
}

function renderConfigView() {
    if (!configState.loaded) {
        return `
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration…</p>
                </article>
            </section>
        `;
    }
    const stepFields = STEP_NOTES_FIELDS[0];
    const payrollDate = formatDateInput(getPayrollDateValue());
    const accountingPeriod = formatDateInput(getConfigValue("PR_Accounting_Period"));
    const jeId = getConfigValue("PR_Journal_Entry_ID");
    const accountingLink = getConfigValue("SS_Accounting_Software");
    const payrollLink = getPayrollProviderLink();
    const companyName = getConfigValue("SS_Company_Name");
    const companyId = getConfigValue("SS_Company_ID");
    const userName = getConfigValue(CONFIG_REVIEWER_FIELD) || getReviewerDefault();
    const notes = stepFields ? getConfigValue(stepFields.note) : "";
    const notesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const reviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const signOffDate = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const isStepComplete = Boolean(signOffDate || getConfigValue(STEP_COMPLETE_FIELDS[0]));

    return `
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every payroll run.</p>
            <p class="pf-hero-hint">${escapeHtml(appState.statusText || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Period Data</h3>
                    <p class="pf-config-subtext">Fields in this section may change each period.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Your Name (Used for sign-offs)</span>
                        <input type="text" id="config-user-name" value="${escapeHtml(userName)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Date</span>
                        <input type="date" id="config-payroll-date" value="${escapeHtml(payrollDate)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${escapeHtml(accountingPeriod)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-je-id" value="${escapeHtml(jeId)}" placeholder="PR-AUTO-YYYY-MM-DD">
                    </label>
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Static Data</h3>
                    <p class="pf-config-subtext">Fields rarely change but should be reviewed.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Company Name</span>
                        <input type="text" id="config-company-name" value="${escapeHtml(companyName)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Company ID <span class="pf-field-hint">(from Prairie Forge CRM)</span></span>
                        <input type="text" id="config-company-id" value="${escapeHtml(companyId)}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${escapeHtml(payrollLink)}" placeholder="https://…">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${escapeHtml(accountingLink)}" placeholder="https://…">
                    </label>
                </div>
            </article>
            ${
                stepFields
                    ? renderInlineNotes({
                          textareaId: "config-notes",
                          value: notes,
                          permanentId: "config-notes-permanent",
                          isPermanent: notesPermanent,
                          hintId: "",
                          saveButtonId: "config-notes-save"
                      })
                    : ""
            }
            ${
                stepFields
                    ? renderSignoff({
                          reviewerInputId: "config-reviewer-name",
                          reviewerValue: reviewer,
                          signoffInputId: "config-signoff-date",
                          signoffValue: signOffDate,
                          isComplete: isStepComplete,
                          saveButtonId: "config-signoff-save",
                          completeButtonId: "config-signoff-toggle",
                          prevButtonId: "config-signoff-prev",
                          nextButtonId: "config-signoff-next"
                      })
                    : ""
            }
        </section>
    `;
}

function renderProviderCard() {
    const providerValue = getPayrollProviderLink();
    const disabledAttr = providerValue ? "" : ' data-disabled="true" aria-disabled="true"';
    return `
        <article class="pf-step-card pf-cta-card pf-clickable" data-provider-card${disabledAttr}>
            <div>
                <p class="pf-cta-label">Payroll Provider</p>
                <p class="pf-cta-text">
                    ${
                        providerValue
                            ? escapeHtml(providerValue)
                            : "Add a Payroll Provider link in configuration to enable this shortcut."
                    }
                </p>
            </div>
            <span aria-hidden="true">↗</span>
        </article>
    `;
}

function renderQuickTipsCard() {
    return `
        <article class="pf-step-card pf-step-detail pf-quick-tips">
            <h3>Quick Tips</h3>
            <p>Helpful tips for importing payroll data will appear here soon.</p>
        </article>
    `;
}

// =============================================================================
// FILE UPLOAD & ADA ANALYSIS STATE
// =============================================================================
const uploadState = {
    file: null,           // Uploaded File object
    fileName: "",         // Display name
    headers: [],          // Extracted column headers
    rowCount: 0,          // Number of data rows
    parsedData: null,     // Full parsed data (2D array)
    analyzing: false,     // Is Ada analyzing?
    mappings: null,       // Suggested column mappings from Ada
    mappingSource: null,  // 'saved' | 'ada_suggested' | null
    error: null           // Error message if any
};

// =============================================================================
// DICTIONARY OPTIONS CACHE (fetched once per session)
// =============================================================================
const dictionaryCache = {
    amountOptions: null,      // string[] - from ada_payroll_column_dictionary.pf_column_name
    dimensionOptions: null,   // string[] - from ada_payroll_dimensions.normalized_dimension
    loading: false,
    loaded: false
};

/**
 * Fetch dictionary options from the database (once per session)
 * - Amount options from ada_payroll_column_dictionary
 * - Dimension options from ada_payroll_dimensions
 */
async function fetchDictionaryOptions() {
    if (dictionaryCache.loaded || dictionaryCache.loading) {
        return;
    }
    
    dictionaryCache.loading = true;
    console.log("[Dictionary] Fetching options from database...");
    
    try {
        // Do not call fetch() directly for warehouse. Use columnMapperRequest().
        const result = await columnMapperRequest("get_options", {
            module: MODULE_KEY
        }, "fetchDictionaryOptions");
        
        if (result.ok) {
            const data = result.data;
            dictionaryCache.amountOptions = data.amount_options || [];
            dictionaryCache.dimensionOptions = data.dimension_options || [];
            dictionaryCache.loaded = true;
            console.log(`[Dictionary] Loaded ${dictionaryCache.amountOptions.length} amounts, ${dictionaryCache.dimensionOptions.length} dimensions`);
        } else {
            console.warn("[Dictionary] Failed to fetch options, using fallbacks");
            useFallbackDictionaries();
        }
    } catch (error) {
        console.error("[Dictionary] Error fetching options:", error);
        useFallbackDictionaries();
    } finally {
        dictionaryCache.loading = false;
    }
}

/**
 * Fallback dictionaries if API fails
 */
// NOTE: BLOCKED_HEADER_PATTERNS and isBlockedHeader() have been removed.
// Column filtering is now database-driven via include_in_matrix field
// in ada_customer_column_mappings table.

function useFallbackDictionaries() {
    dictionaryCache.amountOptions = [
        "Wages_Salary_Amount", "Wages_Overtime_Amount", "Variable_Bonus_Amount", "Variable_Commission_Amount",
        "Wages_PTO_Amount", "Wages_Hourly_Amount", "Wages_Other_Amount",
        "Federal_Taxes_Employee_Amount", "State_Taxes_Employee_Amount",
        "FICA_Taxes_Employee_Amount", "FICA_Taxes_Employer_Amount",
        "401K_Employee_Amount", "401K_Employer_Amount",
        "Health_Employee_Amount", "Health_Employer_Amount",
        "Benefits_Employee_Amount", "Benefits_Employer_Amount",
        "Taxes_PEO_Employer_Amount", "Fees_PEO_Employer_Amount",
        "Workers_Comp_Employer_Amount", "Reimbursements_Amount"
    ];
    dictionaryCache.dimensionOptions = [
        "Employee_Name", "Employee_ID", "Department", "Department_Code", "Department_Name",
        "Location", "Cost_Center", "Job_Title", "Pay_Date", "Pay_Period_Start", "Pay_Period_End",
        "Check_Number", "Pay_Type", "Pay_Frequency"
    ];
    dictionaryCache.loaded = true;
}

// =============================================================================
// EXPENSE REVIEW TAXONOMY CACHE
// Fetched once per session, drives expense review column classification
// =============================================================================
const expenseTaxonomyCache = {
    measures: null,      // Record<string, { bucket, include, sign, displayOrder }>
    dimensions: null,    // string[] - headers that are dimensions (not summed)
    loading: false,
    loaded: false
};

/**
 * Fetch expense review taxonomy from the database (single source of truth)
 * 
 * DATA SOURCES:
 * - Measures: public.ada_payroll_column_dictionary (financial columns)
 *   - pf_column_name: Canonical header name (matches PR_Data_Clean headers)
 *   - expense_review_bucket: Classification (FIXED/VARIABLE/BURDEN/TAX/DEDUCTION/etc.)
 *   - expense_review_include: Whether to include in totals (default true)
 *   - default_sign: +1 for expenses, -1 for deductions (default +1)
 *   - display_order: Sort order in UI (default 100)
 * 
 * - Dimensions: public.ada_payroll_dimensions (grouping columns)
 *   - normalized_dimension: Canonical header name
 *   - semantic_group: Category (identity, location, time, etc.)
 * 
 * USAGE:
 * - For each header in PR_Data_Clean:
 *   1. If matches dimension → use for grouping (Employee, Department, etc.)
 *   2. If matches measure → aggregate by bucket with sign applied
 *   3. If matches neither → show as "Unclassified" advisory warning
 */
async function fetchExpenseTaxonomy() {
    if (expenseTaxonomyCache.loaded || expenseTaxonomyCache.loading) {
        return expenseTaxonomyCache;
    }
    
    expenseTaxonomyCache.loading = true;
    console.log("[ExpenseTaxonomy] Fetching taxonomy from database...");
    
    try {
        const result = await columnMapperRequest("get_expense_taxonomy", {
            module: MODULE_KEY
        }, "fetchExpenseTaxonomy");
        
        if (result.ok) {
            const data = result.data;
            
            // Debug: Log API response details
            console.log(`[ExpenseTaxonomy] API response - measureRowCount: ${data.debug?.measureRowCount || 'N/A'}, dimensionRowCount: ${data.debug?.dimensionRowCount || 'N/A'}`);
            
            // Normalize measure keys to lowercase for case-insensitive lookup
            // Match PR_Data_Clean headers to dictionary.pf_column_name with trim + case-insensitive
            // 
            // INCLUSION LOGIC (using `side` as PRIMARY signal):
            // - side = 'er' (employer) → include in expense review totals
            // - side = 'ee' (employee) → exclude from totals
            // - side = 'na' (not applicable) → exclude from totals
            // expense_review_bucket is only for grouping (FIXED/VARIABLE/BURDEN), not inclusion
            const rawMeasures = data.measures || {};
            expenseTaxonomyCache.measures = {};
            Object.entries(rawMeasures).forEach(([key, value]) => {
                // Trim and lowercase for case-insensitive matching
                const normalizedKey = String(key || "").trim().toLowerCase();
                expenseTaxonomyCache.measures[normalizedKey] = value;
            });
            
            // Normalize dimension keys to lowercase for case-insensitive lookup
            expenseTaxonomyCache.dimensions = new Set(
                (data.dimensions || []).map(d => String(d || "").trim().toLowerCase())
            );
            expenseTaxonomyCache.loaded = true;
            
            // Debug: Log loaded counts and check for specific columns
            console.log(`[ExpenseTaxonomy] Loaded ${Object.keys(expenseTaxonomyCache.measures).length} measures, ${expenseTaxonomyCache.dimensions.size} dimensions`);
            
            // Debug: Check if specific problematic columns exist in loaded dictionary
            const checkCols = ["401k_employer_amount", "fees_peo_employer_amount"];
            checkCols.forEach(col => {
                const found = expenseTaxonomyCache.measures[col];
                console.log(`[ExpenseTaxonomy] Dictionary contains '${col}': ${found ? 'YES' : 'NO'}${found ? ` (side=${found.side}, bucket=${found.bucket}, include=${found.include})` : ''}`);
            });
            
            // Debug: Log first 10 measure keys to verify loading
            const sampleKeys = Object.keys(expenseTaxonomyCache.measures).slice(0, 10);
            console.log(`[ExpenseTaxonomy] Sample measure keys: ${sampleKeys.join(', ')}`);
        } else {
            console.warn("[ExpenseTaxonomy] Failed to fetch taxonomy, using fallbacks");
            useFallbackExpenseTaxonomy();
        }
    } catch (error) {
        console.error("[ExpenseTaxonomy] Error fetching taxonomy:", error);
        useFallbackExpenseTaxonomy();
    } finally {
        expenseTaxonomyCache.loading = false;
    }
    
    return expenseTaxonomyCache;
}

/**
 * Fallback expense taxonomy if API fails
 */
function useFallbackExpenseTaxonomy() {
    // Default measure classifications based on common payroll terminology
    // NOTE: Keys are lowercase for case-insensitive lookup
    // 
    // INCLUSION LOGIC (using `side` as PRIMARY signal):
    // - side = 'er' (employer) → include in expense review totals
    // - side = 'ee' (employee) → exclude from totals
    // - side = 'na' (not applicable) → exclude from totals
    // expense_review_bucket is only for grouping (FIXED/VARIABLE/BURDEN), not inclusion
    expenseTaxonomyCache.measures = {
        // Fixed salary (employer-paid wages)
        "wages_salary_amount": { bucket: "FIXED", include: true, sign: 1, displayOrder: 10, side: "er" },
        "regular_pay_amount": { bucket: "FIXED", include: true, sign: 1, displayOrder: 10, side: "er" },
        "pto_amount": { bucket: "FIXED", include: true, sign: 1, displayOrder: 15, side: "er" },
        "sick_pay_amount": { bucket: "FIXED", include: true, sign: 1, displayOrder: 15, side: "er" },
        "holiday_pay_amount": { bucket: "FIXED", include: true, sign: 1, displayOrder: 15, side: "er" },
        // Variable (employer-paid)
        "overtime_amount": { bucket: "VARIABLE", include: true, sign: 1, displayOrder: 20, side: "er" },
        "bonus_amount": { bucket: "VARIABLE", include: true, sign: 1, displayOrder: 20, side: "er" },
        "commission_amount": { bucket: "VARIABLE", include: true, sign: 1, displayOrder: 20, side: "er" },
        // Burden / Employer Benefits (all employer-paid, especially for PEO scenarios)
        "employer_taxes_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 30, side: "er" },
        "fica_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 30, side: "er" },
        "medicare_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 30, side: "er" },
        "futa_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 30, side: "er" },
        "suta_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 30, side: "er" },
        "health_insurance_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 35, side: "er" },
        "401k_match_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 35, side: "er" },
        "401k_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 35, side: "er" },
        "benefits_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 35, side: "er" },
        "fees_peo_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 36, side: "er" },
        "workers_comp_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 37, side: "er" },
        "disability_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 38, side: "er" },
        "life_insurance_employer_amount": { bucket: "BURDEN", include: true, sign: 1, displayOrder: 39, side: "er" },
        // Employee deductions/taxes - NOT included in employer expense review (side='ee')
        "federal_withholding_amount": { bucket: "TAX", include: false, sign: -1, displayOrder: 40, side: "ee" },
        "state_withholding_amount": { bucket: "TAX", include: false, sign: -1, displayOrder: 40, side: "ee" },
        "fica_employee_amount": { bucket: "TAX", include: false, sign: -1, displayOrder: 40, side: "ee" },
        "medicare_employee_amount": { bucket: "TAX", include: false, sign: -1, displayOrder: 40, side: "ee" },
        "401k_employee_amount": { bucket: "DEDUCTION", include: false, sign: -1, displayOrder: 50, side: "ee" },
        // Summary columns - not included (side='na')
        "net_pay_amount": { bucket: "OTHER", include: false, sign: 1, displayOrder: 90, side: "na" },
        "gross_pay_amount": { bucket: "OTHER", include: false, sign: 1, displayOrder: 90, side: "na" }
    };
    
    // NOTE: Dimension keys are lowercase for case-insensitive lookup
    expenseTaxonomyCache.dimensions = new Set([
        "employee_name", "employee_id", "department", "department_code", "department_name",
        "location", "cost_center", "job_title", "pay_date", "payroll_date",
        "pay_period_start", "pay_period_end", "check_number", "pay_type", "pay_frequency"
    ]);
    
    expenseTaxonomyCache.loaded = true;
}

/**
 * Classify a canonical header using the expense taxonomy
 * Returns: { kind: "dimension" | "measure" | "unclassified", metadata: {...} | null }
 * 
 * NOTE: Lookups are case-insensitive (headers normalized to lowercase with trim)
 * A column is "classified" if it exists in the dictionary, even if include=false
 * The include flag only controls whether it contributes to Expense Review totals
 */
function classifyHeaderByTaxonomy(canonicalHeader) {
    const taxonomy = expenseTaxonomyCache;
    
    if (!taxonomy.loaded) {
        console.warn("[classifyHeaderByTaxonomy] Taxonomy not loaded yet");
        return { kind: "unclassified", metadata: null };
    }
    
    // Normalize header to lowercase with trim for case-insensitive lookup
    // This matches PR_Data_Clean headers to dictionary.pf_column_name
    const headerLower = String(canonicalHeader || "").trim().toLowerCase();
    
    // Check if it's a dimension (not summed)
    if (taxonomy.dimensions && taxonomy.dimensions.has(headerLower)) {
        return { kind: "dimension", metadata: null };
    }
    
    // Check if it's a measure (exists in dictionary)
    // A column is "classified" as a measure if it exists in the dictionary
    // The include flag only controls whether it contributes to totals
    if (taxonomy.measures && taxonomy.measures[headerLower]) {
        return { kind: "measure", metadata: taxonomy.measures[headerLower] };
    }
    
    // Debug: Log unclassified headers that contain "amount" (likely should be classified)
    if (headerLower.includes("amount") || headerLower.includes("employer") || headerLower.includes("fee")) {
        console.warn(`[classifyHeaderByTaxonomy] UNCLASSIFIED potential measure: '${canonicalHeader}' (normalized: '${headerLower}') - not found in dictionary`);
    }
    
    // Unclassified - show advisory warning
    return { kind: "unclassified", metadata: null };
}

// =============================================================================
// GL MAPPINGS - REMOVED (now using jeLoadGLMappings in JE V2 section)
// =============================================================================

/**
 * Hard exclusions for summary totals that should NEVER be counted as expenses.
 * These are surgical safeguards to prevent double-counting.
 * Keep this list minimal and explicit.
 */
const EXPENSE_REVIEW_SUMMARY_EXCLUSIONS = new Set([
    "gross_pay_amount",
    "net_pay_amount"
]);

/**
 * Display names for expense_bucket values in Expense Review
 */
const BUCKET_DISPLAY_NAMES = {
    FIXED: "Fixed Wages",
    VARIABLE: "Variable Compensation",
    BURDEN: "Employer Burden",
    BENEFITS: "Benefits & Retirement",
    // Legacy mappings for backward compatibility
    BENEFIT: "Benefits & Retirement",
    TAX: "Employer Burden",
    DEDUCTION: "Other Deductions",
    REIMBURSEMENT: "Reimbursements",
    OTHER: "Other Expenses",
    UNCLASSIFIED: "Unclassified"
};

/**
 * Normalize bucket names to canonical values
 * Maps database bucket variations to the expected bucket keys
 * 
 * Canonical buckets: FIXED, VARIABLE, BURDEN, BENEFIT, TAX, DEDUCTION, REIMBURSEMENT, OTHER
 */
function normalizeBucketName(bucket) {
    if (!bucket) return "OTHER";
    
    const upper = String(bucket).toUpperCase().trim();
    
    // Map variations to canonical names
    switch (upper) {
        // FIXED variations (wages, salaries)
        case "FIXED":
        case "BASE":
        case "SALARY":
        case "WAGES":
            return "FIXED";
        
        // VARIABLE variations (bonuses, commissions)
        case "VARIABLE":
        case "BONUS":
        case "COMMISSION":
            return "VARIABLE";
        
        // BURDEN variations (employer taxes, FICA, etc.)
        case "BURDEN":
        case "EMPLOYER_TAX":
        case "PAYROLL_TAX":
            return "BURDEN";
        
        // BENEFIT variations (401k, health insurance, etc.)
        case "BENEFIT":
        case "BENEFITS":  // Handle plural form
        case "EMPLOYER_BENEFIT":
            return "BENEFIT";
        
        // TAX variations
        case "TAX":
        case "TAXES":
            return "TAX";
        
        // DEDUCTION variations
        case "DEDUCTION":
        case "DEDUCTIONS":
            return "DEDUCTION";
        
        // REIMBURSEMENT variations
        case "REIMBURSEMENT":
        case "REIMBURSEMENTS":
        case "EXPENSE":
            return "REIMBURSEMENT";
        
        // Everything else
        default:
            return "OTHER";
    }
}

/**
 * Determine if a column should be included in Expense Review
 * 
 * NEW LOGIC (database-driven via customer mappings):
 * - If expense_bucket is FIXED, VARIABLE, BURDEN, or BENEFITS → INCLUDE
 * - If expense_bucket is EXCLUDE or NULL → EXCLUDE
 * - Summary exclusions (gross_pay, net_pay) always excluded as safeguard
 * 
 * @param {object|null} metadata - Dictionary metadata for the column (legacy, for sign only)
 * @param {string} headerLower - Lowercase header name for summary exclusion check
 * @param {object|null} customerMapping - Customer mapping with expense_bucket (preferred)
 * @returns {{ include: boolean, reason: string|null, bucket: string|null }}
 */
function shouldIncludeInExpenseReview(metadata, headerLower = "", customerMapping = null) {
    // Check summary exclusions first (highest priority safeguard)
    if (headerLower && EXPENSE_REVIEW_SUMMARY_EXCLUSIONS.has(headerLower)) {
        return { include: false, reason: "summary_exclusion", bucket: null };
    }
    
    // If customer mapping exists, use expense_bucket for inclusion decision
    if (customerMapping) {
        const bucket = (customerMapping.expense_bucket || "").toUpperCase();
        
        // EXCLUDE bucket means don't include
        if (bucket === "EXCLUDE") {
            return { include: false, reason: "expense_bucket=EXCLUDE", bucket: null };
        }
        
        // Valid buckets that should be included
        const validBuckets = ["FIXED", "VARIABLE", "BURDEN", "BENEFITS", "BENEFIT", "TAX", "REIMBURSEMENT", "OTHER"];
        if (validBuckets.includes(bucket)) {
            return { include: true, reason: null, bucket };
        }
        
        // No expense_bucket set - exclude by default (customer should configure)
        if (!bucket) {
            return { include: false, reason: "no_expense_bucket", bucket: null };
        }
        
        // Unknown bucket - include but flag
        return { include: true, reason: null, bucket: bucket || "UNCLASSIFIED" };
    }
    
    // LEGACY FALLBACK: No customer mapping - use dictionary metadata
    // This path is for backward compatibility during transition
    if (!metadata) {
        return { include: true, reason: null, bucket: "UNCLASSIFIED" };
    }
    
    // Check explicit include=false flag from dictionary
    if (metadata.include === false) {
        return { include: false, reason: "include=false", bucket: null };
    }
    
    // Legacy side-based logic (will be removed once all customers have mappings)
    const side = String(metadata.side || "").toLowerCase().trim();
    if (side === 'ee') {
        return { include: false, reason: "side='ee'", bucket: null };
    }
    if (side === 'na') {
        return { include: false, reason: "side='na'", bucket: null };
    }
    
    // Everything else is included with dictionary bucket
    const bucket = normalizeBucketName(metadata.bucket);
    return { include: true, reason: null, bucket };
}

/**
 * Bind file upload drag-and-drop and button interactions
 */
function bindFileUploadInteractions() {
    const dropzone = document.getElementById("upload-dropzone");
    const fileInput = document.getElementById("upload-file-input");
    const browseBtn = document.getElementById("upload-browse-btn");
    const clearBtn = document.getElementById("upload-clear-btn");
    const analyzeBtn = document.getElementById("ada-analyze-btn");
    const applyBtn = document.getElementById("mapping-apply-btn");
    
    // Dropzone events
    if (dropzone) {
        dropzone.addEventListener("dragover", (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add("pf-upload-zone--dragover");
        });
        
        dropzone.addEventListener("dragleave", (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove("pf-upload-zone--dragover");
        });
        
        dropzone.addEventListener("drop", (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove("pf-upload-zone--dragover");
            
            const files = e.dataTransfer?.files;
            if (files?.length > 0) {
                handleFileUpload(files[0]);
            }
        });
        
        dropzone.addEventListener("click", () => {
            fileInput?.click();
        });
    }
    
    // File input change
    fileInput?.addEventListener("change", (e) => {
        const file = e.target.files?.[0];
        if (file) {
            handleFileUpload(file);
        }
    });
    
    // Browse button
    browseBtn?.addEventListener("click", (e) => {
        e.stopPropagation();
        fileInput?.click();
    });
    
    // Clear file button
    clearBtn?.addEventListener("click", () => {
        resetUploadState();
        renderApp();
    });
    
    // Auto-map columns button
    analyzeBtn?.addEventListener("click", () => {
        triggerAdaColumnAnalysis();
    });
    
    // Apply mappings and create matrix button (on status card)
    applyBtn?.addEventListener("click", () => {
        createMatrix();
    });
    
    // Open mapping modal buttons (pills on status card)
    document.getElementById("mapping-open-modal-btn")?.addEventListener("click", () => openColumnMappingModal());
    document.getElementById("mapping-open-modal-btn-review")?.addEventListener("click", () => openColumnMappingModal());
    document.getElementById("mapping-open-modal-btn-unmapped")?.addEventListener("click", () => openColumnMappingModal());
    
    // Mapping dropdown changes (uses new contract: target, kind)
    document.querySelectorAll(".pf-mapping-select").forEach(select => {
        select.addEventListener("change", (e) => {
            const index = parseInt(e.target.dataset.index);
            const newTarget = e.target.value;
            const selectedOption = e.target.selectedOptions[0];
            const newKind = selectedOption?.dataset?.kind || "amount"; // Default to amount
            
            if (uploadState.mappings?.[index]) {
                const mapping = uploadState.mappings[index];
                
                if (newTarget) {
                    // User selected a mapping - target is CANONICAL
                    mapping.target = newTarget;
                    mapping.kind = newKind;
                    mapping.source = "user_manual";
                    mapping.confidence = 1.0;
                    mapping.manual_override = true;
                } else {
                    // User selected "Skip"
                    mapping.target = null;
                    mapping.kind = null;
                    mapping.source = "unmapped";
                    mapping.confidence = 0;
                }
                
                // Re-render to update status badges
                renderApp();
            }
        });
    });
    
    // Ambiguous mapping choice buttons
    document.querySelectorAll(".pf-ambiguous-btn").forEach(btn => {
        btn.addEventListener("click", (e) => {
            const index = parseInt(e.target.dataset.index);
            const choice = e.target.dataset.choice;
            const target = e.target.dataset.target;
            
            if (uploadState.mappings?.[index]) {
                const mapping = uploadState.mappings[index];
                
                if (choice === "amount" && mapping.amount_option) {
                    // target is CANONICAL
                    mapping.target = mapping.amount_option.target;
                    mapping.kind = "amount";
                    mapping.source = "amount";
                    mapping.confidence = mapping.amount_option.confidence;
                } else if (choice === "dimension" && mapping.dimension_option) {
                    // target is CANONICAL
                    mapping.target = mapping.dimension_option.target;
                    mapping.kind = "dimension";
                    mapping.source = "dimension";
                    mapping.confidence = mapping.dimension_option.confidence;
                } else if (choice === "skip") {
                    mapping.target = null;
                    mapping.kind = null;
                    mapping.source = "unmapped";
                    mapping.confidence = 0;
                }
                
                // Clear ambiguous options
                mapping.amount_option = null;
                mapping.dimension_option = null;
                mapping.manual_override = true;
                mapping.confirmed = true;
                
                // Re-render to show dropdown instead of buttons
                renderApp();
            }
        });
    });
}

/**
 * Reset upload state to initial values
 */
function resetUploadState() {
    uploadState.file = null;
    uploadState.fileName = "";
    uploadState.headers = [];
    uploadState.rowCount = 0;
    uploadState.parsedData = null;
    uploadState.analyzing = false;
    uploadState.mappings = null;
    uploadState.mappingSource = null;
    uploadState.error = null;
}

/**
 * Handle file upload - parse CSV or Excel file
 */
async function handleFileUpload(file) {
    const validTypes = [
        "text/csv",
        "application/vnd.ms-excel",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ];
    const validExtensions = [".csv", ".xlsx", ".xls"];
    
    const ext = file.name.toLowerCase().slice(file.name.lastIndexOf("."));
    if (!validExtensions.includes(ext) && !validTypes.includes(file.type)) {
        uploadState.error = "Please upload a CSV or Excel file (.csv, .xlsx, .xls)";
        renderApp();
        return;
    }
    
    uploadState.error = null;
    uploadState.fileName = file.name;
    uploadState.file = file;
    
    try {
        const data = await parsePayrollFile(file);
        if (!data || data.length < 2) {
            uploadState.error = "File appears empty or has no data rows.";
            renderApp();
            return;
        }
        
        // Filter out completely empty columns (where header and all values are empty)
        const rawHeaders = data[0].map(h => String(h || "").trim());
        const nonEmptyColumnIndices = [];
        
        for (let colIdx = 0; colIdx < rawHeaders.length; colIdx++) {
            const header = rawHeaders[colIdx];
            // Check if column has any non-empty values in data rows
            const hasData = data.slice(1).some(row => {
                const val = String(row[colIdx] || "").trim();
                return val !== "";
            });
            
            // Keep column if it has a header OR has data in any row
            if (header !== "" || hasData) {
                nonEmptyColumnIndices.push(colIdx);
            }
        }
        
        // Filter data to only include non-empty columns
        const filteredData = data.map(row => 
            nonEmptyColumnIndices.map(idx => row[idx])
        );
        
        uploadState.headers = filteredData[0].map(h => String(h || "").trim());
        uploadState.rowCount = filteredData.length - 1;
        uploadState.parsedData = filteredData;
        uploadState.mappings = null;
        uploadState.mappingSource = null;
        
        console.log(`[Upload] Parsed ${uploadState.headers.length} columns, ${uploadState.rowCount} rows`);
        console.log("[Upload] Headers:", uploadState.headers);
        
        renderApp();

        await triggerAdaColumnAnalysis();
        openColumnMappingModal();
    } catch (error) {
        console.error("[Upload] Parse error:", error);
        uploadState.error = `Failed to parse file: ${error.message}`;
        renderApp();
    }
}

/**
 * Parse a payroll file (CSV or Excel) using XLSX library
 * @param {File} file - The file to parse
 * @returns {Promise<Array<Array<any>>>} - 2D array of data
 */
async function parsePayrollFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: "array" });
                
                // Get first sheet
                const sheetName = workbook.SheetNames[0];
                if (!sheetName) {
                    reject(new Error("No sheets found in workbook"));
                    return;
                }
                
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { 
                    header: 1,
                    defval: "",
                    blankrows: false
                });
                
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => reject(new Error("Failed to read file"));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Trigger Ada to analyze column headers and suggest mappings
 * 
 * API contract returns:
 * {
 *   raw_header: string,
 *   kind: "amount" | "dimension" | "ambiguous" | null,
 *   target: string | null,     // PF canonical name (pf_column_name)
 *   source: "saved" | "amount" | "dimension" | "fuzzy" | "unmapped" | "ambiguous",
 *   confidence: number,
 *   gl_account?: string,
 *   gl_account_name?: string,
 *   amount_option?: { target, confidence },
 *   dimension_option?: { target, confidence }
 * }
 */
async function triggerAdaColumnAnalysis() {
    if (!uploadState.headers.length) {
        showToast("No file uploaded yet.", "error");
        return;
    }
    
    uploadState.analyzing = true;
    uploadState.error = null;
    renderApp();
    
    try {
        // Fetch dictionary options (for dropdown population)
        await fetchDictionaryOptions();
        
        const companyId = getConfigValue("SS_Company_ID");
        
        console.log("[Ada] Starting analysis with:", {
            headers: uploadState.headers,
            companyId: companyId,
            module: MODULE_KEY
        });
        
        // Debug toast
        const headerPreview = uploadState.headers.slice(0, 3).join(", ");
        const companyInfo = companyId ? companyId.substring(0, 8) + "..." : "NOT SET";
        showToast(`Analyzing: ${headerPreview}... (Company: ${companyInfo})`, "info");
        
        // NOTE: Header filtering now handled by include_in_matrix in database
        // All headers sent to API; filtering happens when building matrix
        
        // Do not call fetch() directly for warehouse. Use columnMapperRequest().
        const result = await columnMapperRequest("analyze", {
            headers: uploadState.headers,
            crm_company_id: companyId || null,
            module: MODULE_KEY
        }, "analyzeColumns");
        
        if (!result.ok) {
            throw new Error(`API error: ${result.status} - ${result.error}`);
        }
        
        // Extract data from warehouse response
        const apiData = result.data;
        
        console.log("[Ada] Column analysis complete:", apiData);
        
        // Validate new API contract - check for 'kind' and 'target' fields
        const hasNewContract = apiData.mappings && apiData.mappings.length > 0 && 
            Object.hasOwn(apiData.mappings[0], "kind");
        
        if (!hasNewContract && apiData.mappings && apiData.mappings.length > 0) {
            console.warn("[Ada] Received old API format - check Edge Function deployment");
            showToast("Mapping returned an unexpected format. Please retry or contact support.", "error");
            uploadState.analyzing = false;
            renderApp();
            return;
        }
        
        // Map API response to uploadState.mappings with UI state flags
        // CANONICAL: target is the ONLY PF field (pf_column_name)
        uploadState.mappings = (apiData.mappings || []).map(m => {
            // Handle include_in_matrix - convert string "false" to boolean false
            let includeInMatrix = m.include_in_matrix;
            if (includeInMatrix === "false" || includeInMatrix === false) {
                includeInMatrix = false;
            } else if (includeInMatrix === "true" || includeInMatrix === true) {
                includeInMatrix = true;
            } else {
                includeInMatrix = true; // Default to true for backwards compatibility
            }
            
            return {
                // Core mapping data from API
                raw_header: m.raw_header,
                kind: m.kind,                           // "amount" | "dimension" | "ambiguous" | null
                target: m.target,                       // PF canonical name (pf_column_name)
                source: m.source,                       // "saved" | "amount" | "dimension" | "fuzzy" | "ambiguous" | "unmapped"
                confidence: m.confidence || 0,
                gl_account: m.gl_account || null,
                gl_account_name: m.gl_account_name || null,
                
                // NEW: Database-driven filtering and bucket classification
                include_in_matrix: includeInMatrix,
                expense_bucket: m.expense_bucket || null,         // FIXED, VARIABLE, TAXES, BENEFITS, OTHER
                
                // For ambiguous rows, carry both options
                amount_option: m.amount_option || null,
                dimension_option: m.dimension_option || null,
                
                // UI state flags
                confirmed: false,                       // User clicked Apply or accepted
                manual_override: false                  // User changed the mapping
            };
        });
        
        uploadState.mappingSource = result.source || "dictionary";
        uploadState.analyzing = false;
        
        // Summary toast
        const matchCount = result.matched || 0;
        const unmappedCount = result.unmapped || 0;
        const ambiguousCount = result.ambiguous || 0;
        const needsReview = (result.fuzzy || 0) + ambiguousCount;
        
        if (needsReview > 0) {
            showToast(`Auto-mapped ${matchCount} of ${result.total} columns. ${needsReview} need review.`, "info");
        } else {
            showToast(`Auto-mapped ${matchCount} of ${result.total} columns. ${unmappedCount} unmapped.`, "success");
        }
        
        renderApp();
        
    } catch (error) {
        console.error("[Ada] Analysis failed:", error);
        uploadState.analyzing = false;
        
        // Fallback to local basic mapping if API fails
        uploadState.mappings = generateBasicMappings(uploadState.headers);
        uploadState.mappingSource = "local_fallback";
        uploadState.error = "Couldn't connect to mapping service. Using basic auto-detection.";
        
        renderApp();
    }
}

/**
 * Generate basic column mappings using simple keyword matching
 * Used as fallback when API is unavailable
 * Returns new contract shape: { kind, target, source, confidence, ... }
 */
function generateBasicMappings(headers) {
    // NOTE: Header filtering now handled by include_in_matrix in database
    // All headers processed; filtering happens when building matrix
    
    // Map keywords to { target, kind }
    const keywordMap = {
        // Dimensions
        "employee": { target: "Employee_Name", kind: "dimension" },
        "name": { target: "Employee_Name", kind: "dimension" },
        "emp name": { target: "Employee_Name", kind: "dimension" },
        "employee name": { target: "Employee_Name", kind: "dimension" },
        "worker": { target: "Employee_Name", kind: "dimension" },
        "department": { target: "Department", kind: "dimension" },
        "dept": { target: "Department", kind: "dimension" },
        "division": { target: "Department", kind: "dimension" },
        "cost center": { target: "Cost_Center", kind: "dimension" },
        "location": { target: "Location", kind: "dimension" },
        "pay date": { target: "Pay_Date", kind: "dimension" },
        "check date": { target: "Pay_Date", kind: "dimension" },
        // Amounts
        "regular": { target: "Wages_Salary_Amount", kind: "amount" },
        "regular wages": { target: "Wages_Salary_Amount", kind: "amount" },
        "regular pay": { target: "Wages_Salary_Amount", kind: "amount" },
        "base pay": { target: "Wages_Salary_Amount", kind: "amount" },
        "salary": { target: "Wages_Salary_Amount", kind: "amount" },
        "overtime": { target: "Wages_Overtime_Amount", kind: "amount" },
        "ot": { target: "Wages_Overtime_Amount", kind: "amount" },
        "overtime pay": { target: "Wages_Overtime_Amount", kind: "amount" },
        "ot earns": { target: "Wages_Overtime_Amount", kind: "amount" },
        "bonus": { target: "Variable_Bonus_Amount", kind: "amount" },
        "bonus earns": { target: "Variable_Bonus_Amount", kind: "amount" },
        "commission": { target: "Variable_Commission_Amount", kind: "amount" },
        "commission earns": { target: "Variable_Commission_Amount", kind: "amount" },
        "commissions": { target: "Variable_Commission_Amount", kind: "amount" },
        "federal": { target: "Federal_Withholding_Amount", kind: "amount" },
        "fed tax": { target: "Federal_Withholding_Amount", kind: "amount" },
        "federal tax": { target: "Federal_Withholding_Amount", kind: "amount" },
        "federal withholding": { target: "Federal_Withholding_Amount", kind: "amount" },
        "state": { target: "State_Withholding_Amount", kind: "amount" },
        "state tax": { target: "State_Withholding_Amount", kind: "amount" },
        "state withholding": { target: "State_Withholding_Amount", kind: "amount" },
        "fica": { target: "FICA_Employee_Amount", kind: "amount" },
        "social security": { target: "FICA_Employee_Amount", kind: "amount" },
        "ss": { target: "FICA_Employee_Amount", kind: "amount" },
        "medicare": { target: "Medicare_Employee_Amount", kind: "amount" },
        "401k": { target: "401K_Employee_Amount", kind: "amount" },
        "401(k)": { target: "401K_Employee_Amount", kind: "amount" },
        "retirement": { target: "401K_Employee_Amount", kind: "amount" },
        "health": { target: "Health_Insurance_Amount", kind: "amount" },
        "health insurance": { target: "Health_Insurance_Amount", kind: "amount" },
        "medical": { target: "Health_Insurance_Amount", kind: "amount" },
        "dental": { target: "Dental_Insurance_Amount", kind: "amount" },
        "vision": { target: "Vision_Insurance_Amount", kind: "amount" },
        "hsa": { target: "HSA_Employee_Amount", kind: "amount" },
        // More dimensions
        "employee id": { target: "Employee_ID", kind: "dimension" },
        "emp id": { target: "Employee_ID", kind: "dimension" },
        "job": { target: "Job_Title", kind: "dimension" },
        "job title": { target: "Job_Title", kind: "dimension" },
        "title": { target: "Job_Title", kind: "dimension" },
        "position": { target: "Job_Title", kind: "dimension" }
    };
    
    return headers.map(header => {
        const normalized = header.toLowerCase().trim();
        let match = null;
        let confidence = 0.5;
        
        // Check for exact match first
        if (keywordMap[normalized]) {
            match = keywordMap[normalized];
            confidence = 0.9;
        } else {
            // Check for partial matches
            for (const [keyword, entry] of Object.entries(keywordMap)) {
                if (normalized.includes(keyword) || keyword.includes(normalized)) {
                    match = entry;
                    confidence = 0.7;
                    break;
                }
            }
        }
        
        // Return new contract shape - target is CANONICAL
        return {
            raw_header: header,
            kind: match ? match.kind : null,
            target: match ? match.target : null,  // PF canonical name
            source: match ? "local_fallback" : "unmapped",
            confidence: match ? confidence : 0,
            gl_account: null,
            gl_account_name: null,
            amount_option: null,
            dimension_option: null,
            confirmed: false,
            manual_override: false
        };
    });
}

/**
 * Apply the confirmed mappings and import data to PR_Data
 * Uses new contract: target (PF canonical name), kind
 */
async function applyMappingsAndImport() {
    if (!uploadState.mappings || !uploadState.parsedData) {
        showToast("No data to import.", "error");
        return;
    }
    
    // Get valid mappings: must have target and kind (amount or dimension)
    const selectedMappings = uploadState.mappings
        .filter(m => m.target && (m.kind === "amount" || m.kind === "dimension"))
        .map(m => ({
            raw_header: m.raw_header,
            target: m.target,              // PF canonical name
            kind: m.kind,                  // "amount" or "dimension"
            sourceIndex: uploadState.headers.indexOf(m.raw_header),
            gl_account: m.gl_account,
            gl_account_name: m.gl_account_name
        }));
    
    if (selectedMappings.length === 0) {
        showToast("Please select at least one column mapping.", "error");
        return;
    }
    
    // Check for unresolved ambiguous mappings
    const ambiguousCount = uploadState.mappings.filter(m => m.kind === "ambiguous").length;
    if (ambiguousCount > 0) {
        showToast(`Please resolve ${ambiguousCount} ambiguous mapping(s) before importing.`, "error");
        return;
    }
    
    showToast("Importing data to PR_Data_Clean...", "info");
    
    try {
        // Save mappings to database if company ID is set
        const companyId = getConfigValue("SS_Company_ID");
        if (companyId) {
            await saveColumnMappings(companyId, selectedMappings);
        }
        
        // Transform and import data
        await importMappedDataToPRData(selectedMappings);
        
        // Clear upload state
        resetUploadState();
        
        showToast("Data imported successfully!", "success");
        renderApp();
        
    } catch (error) {
        console.error("[Import] Failed:", error);
        showToast(`Import failed: ${error.message}`, "error");
    }
}

/**
 * Save confirmed column mappings to the database
 * CANONICAL: target is the PF column name (pf_column_name)
 * 
 * Payload format:
 * {
 *   action: "save",
 *   crm_company_id: "<SS_Company_ID>",
 *   module: "payroll-recorder",
 *   mappings: [{ raw_header, target, kind }, ...]
 * }
 */
async function saveColumnMappings(companyId, mappings) {
    try {
        const payload = {
            action: "save",
            company_id: companyId,
            module: MODULE_KEY,
            mappings: mappings.map(m => ({
                raw_header: m.raw_header,
                target: m.target,      // PF canonical name (pf_column_name)
                kind: m.kind           // "amount" or "dimension"
            }))
        };
        
        console.log("[Mappings] Saving:", payload);
        
        // Do not call fetch() directly for warehouse. Use columnMapperRequest().
        const result = await columnMapperRequest("save", {
            crm_company_id: payload.company_id,
            module: payload.module,
            mappings: payload.mappings
        }, "saveMappings");
        
        if (!result.ok) {
            console.warn("[Mappings] Failed to save:", result.status, result.error);
        } else {
            console.log("[Mappings] Saved successfully:", result.data);
        }
    } catch (error) {
        console.warn("[Mappings] Save error (non-blocking):", error);
    }
}

// =============================================================================
// CREATE MATRIX - Generate PR_Data_Clean directly from uploaded data + mappings
// =============================================================================

/**
 * Required dimension fields for PR_Data_Clean
 * These must be mapped for the matrix to be created
 */
const REQUIRED_DIMENSION_FIELDS = [
    "Employee_Name"  // At minimum, we need to identify employees
];

/**
 * Canonical header order for PR_Data_Clean output
 * Dimensions first (in order), then amounts (alphabetical)
 */
const CANONICAL_DIMENSION_ORDER = [
    "Pay_Date",
    "Pay_Period_Start", 
    "Pay_Period_End",
    "Employee_Name",
    "Employee_ID",
    "Department",
    "Department_Code",
    "Department_Name",
    "Location",
    "Cost_Center",
    "Job_Title",
    "Check_Number",
    "Pay_Type",
    "Pay_Frequency"
];

/**
 * Create Matrix: Main entrypoint for generating PR_Data_Clean
 * Validates prerequisites, builds mapped data, writes to PR_Data_Clean
 */
async function createMatrix() {
    console.log("[CreateMatrix] Starting matrix creation...");
    
    // Validate prerequisites
    if (!uploadState.parsedData || uploadState.parsedData.length < 2) {
        showToast("Upload data first. No file has been uploaded.", "error");
        console.warn("[CreateMatrix] No parsed data available");
        return;
    }
    
    if (!uploadState.mappings || uploadState.mappings.length === 0) {
        showToast("Confirm mappings first.", "error");
        console.warn("[CreateMatrix] No mappings available");
        return;
    }
    
    // Get valid mappings: must have target, kind, and include_in_matrix !== false
    // Debug: Log include_in_matrix values to diagnose filtering
    console.log("[CreateMatrix] Mapping include_in_matrix values:", 
        uploadState.mappings.map(m => ({
            raw_header: m.raw_header,
            include_in_matrix: m.include_in_matrix,
            type: typeof m.include_in_matrix
        }))
    );
    
    const validMappings = uploadState.mappings
        .filter(m => {
            if (!m.target) return false;
            if (m.kind !== "amount" && m.kind !== "dimension") return false;
            
            // Check include_in_matrix - handle boolean, string, null, undefined
            // Exclude if explicitly set to false (boolean or string)
            if (m.include_in_matrix === false || m.include_in_matrix === "false") {
                console.log(`[CreateMatrix] Excluding ${m.raw_header}: include_in_matrix=${m.include_in_matrix} (${typeof m.include_in_matrix})`);
                return false;
            }
            return true;
        })
        .map(m => ({
            raw_header: m.raw_header,
            target: m.target,              // PF canonical name
            kind: m.kind,                  // "amount" or "dimension"
            sourceIndex: uploadState.headers.indexOf(m.raw_header),
            gl_account: m.gl_account,
            gl_account_name: m.gl_account_name,
            expense_bucket: m.expense_bucket || null  // Include expense_bucket for rate calculation
        }));
    
    // Log what was excluded
    const excludedMappings = uploadState.mappings.filter(m => 
        m.include_in_matrix === false || m.include_in_matrix === "false"
    );
    if (excludedMappings.length > 0) {
        console.log(`[CreateMatrix] Excluding ${excludedMappings.length} columns (include_in_matrix=false):`, 
            excludedMappings.map(m => m.raw_header));
    }
    
    if (validMappings.length === 0) {
        showToast("Please select at least one column mapping.", "error");
        console.warn("[CreateMatrix] No valid mappings selected");
        return;
    }
    
    // Check for unresolved ambiguous mappings
    const ambiguousCount = uploadState.mappings.filter(m => m.kind === "ambiguous").length;
    if (ambiguousCount > 0) {
        showToast(`Please resolve ${ambiguousCount} ambiguous mapping(s) before creating matrix.`, "error");
        console.warn("[CreateMatrix] Unresolved ambiguous mappings:", ambiguousCount);
        return;
    }
    
    // Check for required fields
    const missingRequired = checkRequiredFields(validMappings);
    if (missingRequired.length > 0) {
        showToast(`Mapping incomplete: missing required fields: ${missingRequired.join(", ")}`, "error");
        console.warn("[CreateMatrix] Missing required fields:", missingRequired);
        return;
    }
    
    showToast("Creating matrix...", "info", 10000);
    
    try {
        // Build the 2D output array
        const outputData = buildCleanDataset(validMappings);
        
        console.log("[CreateMatrix] Built output data:", {
            headers: outputData[0],
            rowCount: outputData.length - 1,
            columnCount: outputData[0].length
        });
        
        // Write to PR_Data_Clean
        await writeToDataClean(outputData);
        
        // CRITICAL: Invalidate measure universe cache since PR_Data_Clean changed
        invalidateMeasureUniverseCache();
        
        // Save mappings to database (non-blocking)
        const companyId = getConfigValue("SS_Company_ID");
        if (companyId) {
            await saveColumnMappings(companyId, validMappings);
        }

        // Refresh bank reconciliation, payroll coverage, and roster updates with new data
        console.log("[CreateMatrix] Refreshing validation checks...");
        await refreshBankReconciliation();
        await refreshPayrollCoverage();
        await computeRosterDeltas();

        // Advisory: update roster rates from the newly-built PR_Data_Clean (non-blocking)
        try {
            await updateRosterRatesFromPayrollAdvisory();
        } catch (e) {
            console.warn("[RosterRates] Advisory update failed (non-blocking):", e);
        }
        
        showToast(`Matrix created successfully! ${outputData.length - 1} rows written to PR_Data_Clean.`, "success");
        console.log("[CreateMatrix] Complete");
        
        renderApp();
        
    } catch (error) {
        console.error("[CreateMatrix] Failed:", error);
        showToast(`Failed to create matrix: ${error.message}`, "error");
    }
}

/**
 * Check if all required fields are mapped
 * Returns array of missing field names (empty if all present)
 */
function checkRequiredFields(mappings) {
    const mappedTargets = new Set(mappings.map(m => m.target));
    const missing = [];
    
    for (const required of REQUIRED_DIMENSION_FIELDS) {
        if (!mappedTargets.has(required)) {
            missing.push(required);
        }
    }
    
    return missing;
}

/**
 * Keywords to exclude from Employee Name values
 * Rows containing these in the Employee Name column are filtered out
 */
const EMPLOYEE_NAME_EXCLUSIONS = [
    "total",
    "totals", 
    "summary",
    "subtotal",
    "subtotals",
    "grand total",
    "grand totals"
];

/**
 * Check if a value should be excluded based on Employee Name exclusion keywords
 */
function isExcludedEmployeeName(value) {
    if (!value) return true; // Exclude empty/null values
    const normalized = String(value).toLowerCase().trim();
    if (!normalized) return true;
    return EMPLOYEE_NAME_EXCLUSIONS.some(keyword => normalized.includes(keyword));
}

/**
 * Build the clean dataset from raw data and mappings
 * Returns 2D array: [headers, ...dataRows]
 */
function buildCleanDataset(mappings) {
    const rawData = uploadState.parsedData;
    const rawHeaders = rawData[0];
    
    // Separate dimensions and amounts for proper ordering
    const dimensionMappings = mappings.filter(m => m.kind === "dimension");
    const amountMappings = mappings.filter(m => m.kind === "amount");
    
    // Sort dimensions by canonical order, amounts alphabetically
    const sortedDimensions = dimensionMappings.sort((a, b) => {
        const aIdx = CANONICAL_DIMENSION_ORDER.indexOf(a.target);
        const bIdx = CANONICAL_DIMENSION_ORDER.indexOf(b.target);
        // If not in canonical order, put at end alphabetically
        if (aIdx === -1 && bIdx === -1) return a.target.localeCompare(b.target);
        if (aIdx === -1) return 1;
        if (bIdx === -1) return -1;
        return aIdx - bIdx;
    });
    
    const sortedAmounts = amountMappings.sort((a, b) => a.target.localeCompare(b.target));
    
    // Combine: dimensions first, then amounts
    const orderedMappings = [...sortedDimensions, ...sortedAmounts];
    
    // Build canonical headers
    const canonicalHeaders = orderedMappings.map(m => m.target);
    
    // Build mapped column indexes
    const mappedIndexes = orderedMappings.map(m => m.sourceIndex);
    
    // Find Employee_Name source column index for filtering
    const employeeNameMapping = mappings.find(m => m.target === "Employee_Name");
    const employeeNameSourceIndex = employeeNameMapping ? employeeNameMapping.sourceIndex : -1;
    
    console.log("[BuildCleanDataset] Ordered mappings:", orderedMappings.map(m => ({
        target: m.target,
        kind: m.kind,
        sourceIndex: m.sourceIndex
    })));
    console.log("[BuildCleanDataset] Employee_Name source index:", employeeNameSourceIndex);
    
    // Build output array
    const output = [];
    output.push(canonicalHeaders);
    
    let skippedRows = 0;
    
    // Process each data row (skip header row)
    for (let i = 1; i < rawData.length; i++) {
        const rawRow = rawData[i];
        
        // Get Employee Name value for filtering
        const employeeName = employeeNameSourceIndex >= 0 ? rawRow[employeeNameSourceIndex] : rawRow[0];
        
        // Skip rows with excluded Employee Name values (Total, Totals, Summary, Subtotal, etc.)
        if (isExcludedEmployeeName(employeeName)) {
            skippedRows++;
            continue;
        }
        
        const outputRow = mappedIndexes.map(idx => {
            if (idx === -1 || idx === null || idx === undefined) return "";
            const value = rawRow[idx];
            // Clean up the value - handle numbers and strings appropriately
            if (value === null || value === undefined) return "";
            return value;
        });
        
        output.push(outputRow);
    }
    
    console.log("[BuildCleanDataset] Output stats:", {
        headerCount: canonicalHeaders.length,
        dataRowCount: output.length - 1,
        skippedRows: skippedRows,
        dimensions: sortedDimensions.length,
        amounts: sortedAmounts.length
    });
    
    return output;
}

/**
 * Write 2D array to PR_Data_Clean sheet
 * Creates sheet if not exists, clears and replaces if exists
 */
async function writeToDataClean(data) {
    if (!data || data.length === 0) {
        throw new Error("No data to write");
    }
    
    const sheetName = SHEET_NAMES.DATA_CLEAN;
    const rowCount = data.length;
    const colCount = data[0].length;
    
    console.log(`[WriteToDataClean] Writing ${rowCount} rows x ${colCount} columns to ${sheetName}`);
    
    await Excel.run(async (context) => {
        // Try to get existing sheet, or create new one
        let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();
        
        if (sheet.isNullObject) {
            // Create the sheet
            console.log(`[WriteToDataClean] Creating new sheet: ${sheetName}`);
            sheet = context.workbook.worksheets.add(sheetName);
            await context.sync();
        } else {
            // Clear existing used range
            console.log(`[WriteToDataClean] Clearing existing sheet: ${sheetName}`);
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("isNullObject");
            await context.sync();
            
            if (!usedRange.isNullObject) {
                usedRange.clear();
                await context.sync();
            }
        }
        
        // Write data starting at A1 (single batch write)
        const targetRange = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
        targetRange.values = data;
        
        // Format header row
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, colCount);
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#4C2FFF";  // Prairie Forge purple
        headerRange.format.font.color = "#FFFFFF";
        
        // Freeze top row
        sheet.freezePanes.freezeRows(1);
        
        // Auto-fit columns for readability
        targetRange.format.autofitColumns();
        
        // Activate the sheet to show user the result
        sheet.activate();
        
        await context.sync();
        
        console.log(`[WriteToDataClean] Successfully wrote ${rowCount} rows to ${sheetName}`);
    });
}

/**
 * Import the mapped data into PR_Data_Clean sheet
 * Uses target (PF canonical name) as the normalized column header
 * @deprecated Use createMatrix() instead - this function is legacy
 */
async function importMappedDataToPRData(mappings) {
    if (!uploadState.parsedData || uploadState.parsedData.length < 2) {
        throw new Error("No data to import");
    }
    
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
        sheet.load("isNullObject");
        await context.sync();
        
        if (sheet.isNullObject) {
            throw new Error("PR_Data_Clean sheet not found. Please create it first.");
        }
        
        // Clear existing data
        const usedRange = sheet.getUsedRangeOrNullObject();
        await context.sync();
        if (!usedRange.isNullObject) {
            usedRange.clear();
        }
        
        // Build the output data
        // Header row: use target (PF canonical names)
        // This is the key change - downstream formulas expect names like:
        // "Wages_Salary_Amount", "Employee_Name", "Department", etc.
        const headerRow = mappings.map(m => m.target);
        
        console.log("[Import] Normalized headers:", headerRow);
        
        // Data rows: extract only mapped columns
        const dataRows = [];
        for (let i = 1; i < uploadState.parsedData.length; i++) {
            const sourceRow = uploadState.parsedData[i];
            const targetRow = mappings.map(m => sourceRow[m.sourceIndex] ?? "");
            dataRows.push(targetRow);
        }
        
        // Combine header + data
        const allData = [headerRow, ...dataRows];
        
        // Write to sheet
        const targetRange = sheet.getRangeByIndexes(0, 0, allData.length, allData[0].length);
        targetRange.values = allData;
        
        // Format headers
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, headerRow.length);
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#4C2FFF";
        headerRange.format.font.color = "#FFFFFF";
        
        // Auto-fit columns
        targetRange.format.autofitColumns();
        
        await context.sync();
        console.log(`[Import] Wrote ${allData.length} rows to PR_Data_Clean`);
    });
}

/**
 * Render the file upload dropzone HTML
 */
function renderFileUploadZone() {
    const hasFile = uploadState.file || uploadState.headers.length > 0;
    const isAnalyzing = uploadState.analyzing;
    
    if (isAnalyzing) {
        return `
            <div class="pf-upload-zone pf-upload-zone--analyzing">
                <div class="pf-upload-spinner"></div>
                <p class="pf-upload-status">Analyzing your columns…</p>
            </div>
        `;
    }
    
    if (hasFile) {
        return `
            <div class="pf-upload-zone pf-upload-zone--ready">
                <div class="pf-upload-file-info">
                    <svg class="pf-upload-file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                        <line x1="16" y1="13" x2="8" y2="13"/>
                        <line x1="16" y1="17" x2="8" y2="17"/>
                    </svg>
                    <div class="pf-upload-file-details">
                        <span class="pf-upload-filename">${escapeHtml(uploadState.fileName)}</span>
                        <span class="pf-upload-meta">${uploadState.headers.length} columns • ${uploadState.rowCount.toLocaleString()} rows</span>
                    </div>
                    <button type="button" class="pf-upload-clear" id="upload-clear-btn" title="Remove file">×</button>
                </div>
            </div>
        `;
    }
    
    return `
        <div class="pf-upload-zone" id="upload-dropzone">
            <input type="file" id="upload-file-input" accept=".csv,.xlsx,.xls" hidden>
            <div class="pf-upload-content">
                <svg class="pf-upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="17 8 12 3 7 8"/>
                    <line x1="12" y1="3" x2="12" y2="15"/>
                </svg>
                <p class="pf-upload-text">Drop your payroll file here</p>
                <p class="pf-upload-hint">or <button type="button" class="pf-upload-browse" id="upload-browse-btn">browse</button> to upload</p>
                <p class="pf-upload-formats">Supports CSV, XLSX, XLS</p>
            </div>
        </div>
    `;
}

/**
 * Render column mapping preview after analysis
 * Uses new contract: kind, target, source, confidence
 * 
 * UI RULES:
 * - No amount/dimension labels visible to user (internal only)
 * - Status pills: Saved, Auto, Review, Select
 * - Info icon with contextual help
 * - Single action button: "Create Data Sheet"
 * 
 * NOW: Returns a compact status indicator with a button to open the full mapper modal.
 * This prevents the side panel from jumping around when mapping UI appears/changes.
 */
function renderColumnMappingPreview() {
    if (!uploadState.mappings || uploadState.mappings.length === 0) {
        return "";
    }

    // Filter out excluded columns for display counts
    const visibleMappings = uploadState.mappings.filter(m => 
        m.include_in_matrix !== false && m.include_in_matrix !== "false"
    );
    
    const total = visibleMappings.length;
    const mapped = visibleMappings.filter(m => m.target && m.kind !== "ambiguous").length;
    const needsReview = visibleMappings.filter(m => m.kind === "ambiguous" || (m.source === "fuzzy" && m.target)).length;
    const needsMapping = visibleMappings.filter(m => !m.target || m.source === "unmapped").length;

    const sourceLabel = uploadState.mappingSource === "saved"
        ? "Saved"
        : uploadState.mappingSource === "local_fallback"
            ? "Auto"
            : "Auto";

    let statusBadge = "";
    if (needsReview > 0 || needsMapping > 0) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--review" role="status"><span>Review</span></span>`;
    } else {
        statusBadge = `<span class="pf-status-badge pf-status-badge--ok" role="status"><span>Ready</span></span>`;
    }
    
    // Compact status card with CTA
    return `
        <article class="pf-step-card pf-step-detail pf-config-card pf-mapping-status-card">
            <div class="pf-config-head">
                <h3>Column Mapping ${statusBadge}</h3>
                <p class="pf-config-subtext">${sourceLabel} — ${mapped} of ${total} mapped</p>
            </div>
            <div class="pf-mapping-expand-bars">
                <button type="button" class="pf-expand-bar pf-expand-bar--success pf-clickable" id="mapping-open-modal-btn">
                    <svg class="pf-expand-bar-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                        <polyline points="6 9 12 15 18 9"/>
                    </svg>
                    <span class="pf-expand-bar-count">${mapped}</span>
                    <span class="pf-expand-bar-label">mapped</span>
                </button>
                ${needsReview > 0 ? `
                    <button type="button" class="pf-expand-bar pf-expand-bar--warning pf-clickable" id="mapping-open-modal-btn-review">
                        <svg class="pf-expand-bar-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                            <polyline points="6 9 12 15 18 9"/>
                        </svg>
                        <span class="pf-expand-bar-count">${needsReview}</span>
                        <span class="pf-expand-bar-label">to review</span>
                    </button>
                ` : ""}
                ${needsMapping > 0 ? `
                    <button type="button" class="pf-expand-bar pf-expand-bar--muted pf-clickable" id="mapping-open-modal-btn-unmapped">
                        <svg class="pf-expand-bar-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                            <polyline points="6 9 12 15 18 9"/>
                        </svg>
                        <span class="pf-expand-bar-count">${needsMapping}</span>
                        <span class="pf-expand-bar-label">to select</span>
                    </button>
                ` : ""}
            </div>
            <div class="pf-mapping-actions pf-mapping-actions--cta">
                <button type="button" class="pf-cta-button pf-clickable" id="mapping-apply-btn" title="Create the data matrix">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="20" height="20">
                        <rect x="3" y="3" width="18" height="18" rx="2"/>
                        <path d="M3 9h18"/>
                        <path d="M3 15h18"/>
                        <path d="M9 3v18"/>
                        <path d="M15 3v18"/>
                    </svg>
                    <span>Create Matrix</span>
                </button>
            </div>
        </article>
    `;
}

/**
 * Open the column mapping modal with full mapping UI
 */
function openColumnMappingModal() {
    // Remove existing modal if present
    const existing = document.getElementById("column-mapping-modal");
    if (existing) existing.remove();
    
    if (!uploadState.mappings || uploadState.mappings.length === 0) {
        showToast("No mappings to edit.", "info");
        return;
    }
    
    // Filter out columns that are excluded from matrix (include_in_matrix === false)
    // These are system columns like Gross_Pay that shouldn't be shown to users
    const visibleMappings = uploadState.mappings.filter(m => 
        m.include_in_matrix !== false && m.include_in_matrix !== "false"
    );
    
    const excludedCount = uploadState.mappings.length - visibleMappings.length;
    if (excludedCount > 0) {
        console.log(`[MappingUI] Hiding ${excludedCount} columns with include_in_matrix=false`);
    }
    
    const sourceLabel = uploadState.mappingSource === "saved" 
        ? "Using your saved mappings" 
        : "Auto-mapped columns";
    
    const mappingRows = visibleMappings.map((m, i) => {
        // Find the original index in uploadState.mappings for data-index
        const originalIndex = uploadState.mappings.indexOf(m);
        
        // Determine status pill - simplified, no jargon
        let statusPill = "";
        let rowClass = "pf-mapping-row";
        
        if (m.kind === "ambiguous") {
            statusPill = `<span class="pf-pill pf-pill--warning">Review</span>`;
            rowClass += " pf-mapping-row--ambiguous";
        } else if (m.source === "unmapped" || !m.target) {
            statusPill = `<span class="pf-pill pf-pill--neutral">Select</span>`;
            rowClass += " pf-mapping-row--unmapped";
        } else if (m.source === "fuzzy" || m.confidence < 0.9) {
            statusPill = `<span class="pf-pill pf-pill--warning">Review</span>`;
            rowClass += " pf-mapping-row--low-conf";
        } else if (m.source === "saved") {
            statusPill = `<span class="pf-pill pf-pill--success">Saved</span>`;
            rowClass += " pf-mapping-row--saved";
        } else {
            statusPill = `<span class="pf-pill pf-pill--success">Auto</span>`;
        }
        
        return `
            <div class="${rowClass}" data-index="${originalIndex}">
                <span class="pf-mapping-raw">${escapeHtml(m.raw_header)}</span>
                <span class="pf-mapping-arrow">to</span>
                ${m.kind === "ambiguous" ? renderAmbiguousOptions(m, originalIndex) : `
                    <select class="pf-mapping-select" data-index="${originalIndex}" data-raw="${escapeHtml(m.raw_header)}">
                        <option value="">Skip this column</option>
                        ${renderMappingOptions(m.target, m.kind)}
                    </select>
                `}
                ${statusPill}
            </div>
        `;
    }).join("");
    
    // Summary counts - use visible mappings for display
    const mapped = visibleMappings.filter(m => m.target && m.kind !== "ambiguous").length;
    const needsReview = visibleMappings.filter(m => m.kind === "ambiguous" || (m.source === "fuzzy" && m.target)).length;
    const needsMapping = visibleMappings.filter(m => !m.target || m.source === "unmapped").length;
    const total = visibleMappings.length;
    const isDirty = uploadState.mappings.some(m => m.manual_override);
    
    const modal = document.createElement("div");
    modal.id = "column-mapping-modal";
    modal.className = "pf-mapping-modal";
    modal.innerHTML = `
        <div class="pf-mapping-modal-backdrop" data-close></div>
        <div class="pf-mapping-modal-card">
            <div class="pf-mapping-modal-header">
                <div>
                    <h3 class="pf-mapping-modal-title">Column Mapping</h3>
                    <p class="pf-mapping-modal-subtext">${sourceLabel}. Review and adjust as needed.</p>
                </div>
                <button class="pf-coverage-modal-close pf-clickable" type="button" aria-label="Close" data-close>
                    ${X_ICON_SVG}
                </button>
            </div>
            <div class="pf-mapping-modal-summary">
                <span class="pf-summary-item">${mapped} of ${total} mapped</span>
                ${needsReview > 0 ? `<span class="pf-summary-item pf-summary--warning">${needsReview} to review</span>` : ''}
                ${needsMapping > 0 ? `<span class="pf-summary-item pf-summary--muted">${needsMapping} to select</span>` : ''}
            </div>
            <div class="pf-mapping-modal-body">
                <div class="pf-mapping-grid">
                    ${mappingRows}
                </div>
            </div>
            <div class="pf-mapping-modal-footer">
                <button type="button" class="pf-pill-btn pf-pill-btn--primary ${isDirty ? "pf-clickable" : ""}" id="modal-mapping-save-btn" ${isDirty ? "" : "disabled"}>
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                        <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                        <polyline points="17 21 17 13 7 13 7 21"/>
                        <polyline points="7 3 7 8 15 8"/>
                    </svg>
                    <span>Save</span>
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Bind close handlers
    modal.querySelectorAll("[data-close]").forEach(el => {
        el.addEventListener("click", () => closeColumnMappingModal());
    });
    
    // Bind Save button in modal
    modal.querySelector("#modal-mapping-save-btn")?.addEventListener("click", async () => {
        // Get company ID from config
        const companyId = configState.values?.SS_Company_ID || configState.values?.company_id;
        if (!companyId) {
            showToast("Company ID not found. Save mappings after configuration is complete.", "info");
            closeColumnMappingModal();
            return;
        }
        
        // Save the current mappings
        const mappingsToSave = uploadState.mappings.filter(m => m.target && m.kind);
        if (mappingsToSave.length > 0) {
            await saveColumnMappings(companyId, mappingsToSave);
            showToast(`Saved ${mappingsToSave.length} column mapping(s).`, "success");
        }
        closeColumnMappingModal();
    });
    
    // Bind mapping select changes
    modal.querySelectorAll(".pf-mapping-select").forEach(select => {
        select.addEventListener("change", (e) => {
            const index = parseInt(e.target.dataset.index, 10);
            const newTarget = e.target.value;
            const selectedOption = e.target.selectedOptions?.[0];
            const newKind = selectedOption?.dataset?.kind || null;
            if (uploadState.mappings[index]) {
                uploadState.mappings[index].target = newTarget || null;
                uploadState.mappings[index].kind = newTarget ? newKind : null;
                uploadState.mappings[index].manual_override = true;
                uploadState.mappings[index].source = newTarget ? "manual" : "unmapped";
                // Re-render modal content to update pills
                updateMappingModalSummary();
            }
        });
    });
    
    // Bind ambiguous option buttons
    modal.querySelectorAll(".pf-ambiguous-btn").forEach(btn => {
        btn.addEventListener("click", (e) => {
            const index = parseInt(btn.dataset.index, 10);
            const choice = btn.dataset.choice;
            const target = btn.dataset.target;
            
            if (uploadState.mappings[index]) {
                if (choice === "skip") {
                    uploadState.mappings[index].target = null;
                    uploadState.mappings[index].kind = null;
                    uploadState.mappings[index].source = "unmapped";
                } else if (choice === "amount") {
                    uploadState.mappings[index].target = target;
                    uploadState.mappings[index].kind = "amount";
                    uploadState.mappings[index].source = "manual";
                } else if (choice === "dimension") {
                    uploadState.mappings[index].target = target;
                    uploadState.mappings[index].kind = "dimension";
                    uploadState.mappings[index].source = "manual";
                }
                uploadState.mappings[index].manual_override = true;
                // Refresh modal
                closeColumnMappingModal();
                openColumnMappingModal();
            }
        });
    });
    
    // Close on escape key
    const escHandler = (e) => {
        if (e.key === "Escape") {
            closeColumnMappingModal();
            document.removeEventListener("keydown", escHandler);
        }
    };
    document.addEventListener("keydown", escHandler);
}

/**
 * Close the column mapping modal
 */
function closeColumnMappingModal() {
    const modal = document.getElementById("column-mapping-modal");
    if (modal) modal.remove();
    // Re-render the side panel to update status
    renderApp();
}

/**
 * Update the modal summary counts without full re-render
 */
function updateMappingModalSummary() {
    const modal = document.getElementById("column-mapping-modal");
    if (!modal) return;
    
    const mapped = uploadState.mappings.filter(m => m.target && m.kind !== "ambiguous").length;
    const needsReview = uploadState.mappings.filter(m => m.kind === "ambiguous" || (m.source === "fuzzy" && m.target)).length;
    const needsMapping = uploadState.mappings.filter(m => !m.target || m.source === "unmapped").length;
    const total = uploadState.mappings.length;
    
    const summaryEl = modal.querySelector(".pf-mapping-modal-summary");
    if (summaryEl) {
        summaryEl.innerHTML = `
            <span class="pf-summary-item">${mapped} of ${total} mapped</span>
            ${needsReview > 0 ? `<span class="pf-summary-item pf-summary--warning">${needsReview} to review</span>` : ''}
            ${needsMapping > 0 ? `<span class="pf-summary-item pf-summary--muted">${needsMapping} to select</span>` : ''}
        `;
    }

    const saveBtn = modal.querySelector("#modal-mapping-save-btn");
    if (saveBtn) {
        const isDirty = uploadState.mappings.some(m => m.manual_override);
        saveBtn.disabled = !isDirty;
        if (isDirty) {
            saveBtn.classList.add("pf-clickable");
        } else {
            saveBtn.classList.remove("pf-clickable");
        }
    }
}

/**
 * Render options for ambiguous mappings (multiple matches found)
 * UI does NOT expose amount/dimension terminology - just presents choices
 */
function renderAmbiguousOptions(mapping, index) {
    const option1 = mapping.amount_option?.target || "Option 1";
    const option2 = mapping.dimension_option?.target || "Option 2";
    
    return `
        <div class="pf-ambiguous-options">
            <button type="button" class="pf-ambiguous-btn pf-ambiguous-btn--primary" 
                    data-index="${index}" data-choice="amount" data-target="${escapeHtml(option1)}">
                ${escapeHtml(formatKeyLabel(option1))}
            </button>
            <button type="button" class="pf-ambiguous-btn pf-ambiguous-btn--secondary" 
                    data-index="${index}" data-choice="dimension" data-target="${escapeHtml(option2)}">
                ${escapeHtml(formatKeyLabel(option2))}
            </button>
            <button type="button" class="pf-ambiguous-btn pf-ambiguous-btn--skip" 
                    data-index="${index}" data-choice="skip">
                Skip
            </button>
        </div>
    `;
}

/**
 * Render options for the mapping dropdown
 * Uses dictionary options from ada_payroll_column_dictionary (amounts)
 * and ada_payroll_dimensions (dimensions)
 * 
 * IMPORTANT: filterKind determines which options to show:
 *   - "amount" → only amount options
 *   - "dimension" → only dimension options
 *   - null → show both (for ambiguous or user override)
 * 
 * @param {string} selectedTarget - Currently selected PF canonical name
 * @param {string} filterKind - Filter: "amount" | "dimension" | null
 */
function renderMappingOptions(selectedTarget, filterKind = null) {
    // Use cached dictionary options (source of truth)
    const amountOptions = dictionaryCache.amountOptions || [];
    const dimensionOptions = dictionaryCache.dimensionOptions || [];
    
    let html = "";
    
    // If selectedTarget exists and isn't in our lists, add it first (to preserve unknown values)
    const allTargets = new Set([...amountOptions, ...dimensionOptions]);
    if (selectedTarget && !allTargets.has(selectedTarget)) {
        const label = formatKeyLabel(selectedTarget);
        html += `<option value="${escapeHtml(selectedTarget)}" selected>${escapeHtml(label)}</option>`;
    }
    
    // STRICT FILTERING: Show only the options for the specified kind
    // No cross-contamination!
    
    if (filterKind === "amount") {
        // ONLY amounts
        html += `<optgroup label="Amounts">`;
        for (const target of amountOptions.sort()) {
            const label = formatKeyLabel(target);
            const selected = target === selectedTarget ? "selected" : "";
            html += `<option value="${escapeHtml(target)}" data-kind="amount" ${selected}>${escapeHtml(label)}</option>`;
        }
        html += `</optgroup>`;
    } else if (filterKind === "dimension") {
        // ONLY dimensions
        html += `<optgroup label="Dimensions">`;
        for (const target of dimensionOptions.sort()) {
            const label = formatKeyLabel(target);
            const selected = target === selectedTarget ? "selected" : "";
            html += `<option value="${escapeHtml(target)}" data-kind="dimension" ${selected}>${escapeHtml(label)}</option>`;
        }
        html += `</optgroup>`;
    } else {
        // Show both (for ambiguous or no filter)
        html += `<optgroup label="Amounts">`;
        for (const target of amountOptions.sort()) {
            const label = formatKeyLabel(target);
            const selected = target === selectedTarget ? "selected" : "";
            html += `<option value="${escapeHtml(target)}" data-kind="amount" ${selected}>${escapeHtml(label)}</option>`;
        }
        html += `</optgroup>`;
        
        html += `<optgroup label="Dimensions">`;
        for (const target of dimensionOptions.sort()) {
            const label = formatKeyLabel(target);
            const selected = target === selectedTarget ? "selected" : "";
            html += `<option value="${escapeHtml(target)}" data-kind="dimension" ${selected}>${escapeHtml(label)}</option>`;
        }
        html += `</optgroup>`;
    }
    
    return html;
}

/**
 * Format a PF canonical name (target) into a human-readable label
 * e.g., "Wages_Salary_Amount" → "Wages Salary Amount"
 */
function formatKeyLabel(key) {
    if (!key) return "";
    // Handle UPPERCASE keys: "REGULARPAY" → "Regular Pay"
    // Handle snake_case keys: "regular_pay" → "Regular Pay"
    return key
        .replace(/_/g, " ")
        .replace(/([a-z])([A-Z])/g, "$1 $2")  // camelCase split
        .replace(/([A-Z]+)/g, (match) => match.charAt(0) + match.slice(1).toLowerCase())
        .replace(/\b\w/g, c => c.toUpperCase())
        .trim();
}

function renderImportStep(detail) {
    const stepFields = getStepNoteFields(1);
    const notesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[1]));
    // Error display
    const errorHtml = uploadState.error 
        ? `<p class="pf-upload-error">${escapeHtml(uploadState.error)}</p>` 
        : "";
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">Upload & Validate Payroll Data</h2>
            <p class="pf-hero-copy">Upload your payroll export, create the data matrix, and verify coverage.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Import</h3>
                    <p class="pf-config-subtext">Drop your payroll file to auto-map columns.</p>
                </div>
                ${renderFileUploadZone()}
                ${errorHtml}
            </article>
            ${renderColumnMappingPreview()}
            
            <div class="pf-validation-section">
                <h3 class="pf-validation-header">Validation</h3>
                <p class="pf-validation-subtext">Advisory checks — review but not required to proceed.</p>
                ${renderBankReconciliationCard()}
                ${renderEmployeeCoverageCard()}
            </div>
            ${stepFields ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-1",
                    value: stepNotes || "",
                    permanentId: "step-notes-lock-1",
                    isPermanent: notesPermanent,
                    saveButtonId: "step-notes-save-1"
                })}
                ${renderSignoff({
                    reviewerInputId: "step-reviewer-1",
                    reviewerValue: stepReviewer,
                    signoffInputId: "step-signoff-1",
                    signoffValue: stepSignOff,
                    isComplete: stepComplete,
                    saveButtonId: "step-signoff-save-1",
                    completeButtonId: "step-signoff-toggle-1"
                })}
            ` : ""}
        </section>
    `;
}

/**
 * Render taxonomy classification advisory card
 * Shows which columns are being used for measures/dimensions and any unclassified columns
 */
function renderTaxonomyAdvisoryCard() {
    const { unclassifiedColumns, measureColumns, dimensionColumns, unclassifiedTotals, sumUnclassifiedNumeric, completenessCheck } = expenseReviewState;
    
    // Don't show if no data has been processed yet
    if (!measureColumns?.length && !dimensionColumns?.length && !unclassifiedColumns?.length) {
        return "";
    }
    
    const hasUnclassified = unclassifiedColumns?.length > 0;
    const statusBadge = hasUnclassified
        ? `<span class="pf-status-badge pf-status-badge--warning" role="status">${ALERT_TRIANGLE_SVG}<span>Review</span></span>`
        : `<span class="pf-status-badge pf-status-badge--success" role="status">${CHECK_CIRCLE_SVG}<span>OK</span></span>`;
    
    // Group measures by bucket
    const bucketGroups = {};
    (measureColumns || []).forEach(col => {
        const bucket = col.bucket || "OTHER";
        if (!bucketGroups[bucket]) bucketGroups[bucket] = [];
        bucketGroups[bucket].push(col.header);
    });
    
    const bucketSummary = Object.entries(bucketGroups)
        .map(([bucket, cols]) => `<strong>${bucket}</strong>: ${cols.length} columns`)
        .join(" · ");
    
    let unclassifiedHtml = "";
    if (hasUnclassified) {
        const displayCols = unclassifiedColumns.slice(0, 5);
        const moreCount = unclassifiedColumns.length - displayCols.length;
        
        // DIAGNOSTIC: Calculate delta and show top unclassified by dollars
        const expenseTotal = expenseReviewState.periods?.[0]?.summary?.total || 0;
        const prDataCleanTotal = completenessCheck?.currentPeriod?.prDataClean || 0;
        const delta = prDataCleanTotal - expenseTotal;
        
        // Get top 5 unclassified columns by dollar amount
        const sortedUnclassified = Object.entries(unclassifiedTotals || {})
            .filter(([_, amt]) => Math.abs(amt) > 0)
            .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
        
        const top5Unclassified = sortedUnclassified.slice(0, 5);
        const formatDollar = (num) => `$${Math.abs(num).toLocaleString(undefined, { minimumFractionDigits: 2 })}${num < 0 ? " (credit)" : ""}`;
        
        // Check if unclassified sum explains the delta
        const deltaExplained = sumUnclassifiedNumeric > 0 && Math.abs(delta - sumUnclassifiedNumeric) < 1;
        
        let diagnosticHtml = "";
        if (sumUnclassifiedNumeric > 0) {
            diagnosticHtml = `
                <div style="margin-top: 12px; padding: 10px; background: rgba(251, 191, 36, 0.1); border-radius: 6px; border-left: 3px solid #fbbf24;">
                    <p style="font-size: 12px; font-weight: 600; color: #fbbf24; margin-bottom: 6px;">
                        Diagnostic: Missing Amount Analysis
                    </p>
                    <div style="font-size: 11px; color: rgba(255,255,255,0.8); line-height: 1.6;">
                        <p><strong>Sum of unclassified numeric columns:</strong> ${formatDollar(sumUnclassifiedNumeric)}</p>
                        ${prDataCleanTotal > 0 ? `<p><strong>PR_Data_Clean Total:</strong> ${formatDollar(prDataCleanTotal)}</p>` : ""}
                        ${expenseTotal > 0 ? `<p><strong>Expense Review Total:</strong> ${formatDollar(expenseTotal)}</p>` : ""}
                        ${delta !== 0 ? `<p><strong>Delta (difference):</strong> ${formatDollar(delta)}</p>` : ""}
                        ${deltaExplained ? `<p style="color: #4ade80; font-weight: 600;">Unclassified columns explain the difference.</p>` : ""}
                    </div>
                    ${top5Unclassified.length > 0 ? `
                        <p style="font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.9); margin-top: 10px;">Top unclassified by dollars:</p>
                        <ol style="font-size: 11px; color: rgba(255,255,255,0.7); margin: 4px 0 0 16px; padding: 0;">
                            ${top5Unclassified.map(([col, amt]) => `<li style="margin-bottom: 2px;"><code style="background: rgba(255,255,255,0.1); padding: 1px 4px; border-radius: 3px;">${escapeHtml(col)}</code>: ${formatDollar(amt)}</li>`).join("")}
                        </ol>
                    ` : ""}
                </div>
            `;
        }
        
        unclassifiedHtml = `
            <div class="pf-coverage-hint" style="margin-top: 8px;">
                <p class="pf-metric-hint pf-metric-hint--warning">
                    <strong>Unclassified columns:</strong> ${displayCols.join(", ")}${moreCount > 0 ? ` (+${moreCount} more)` : ""}
                </p>
                <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 4px;">
                    These columns are not in the dictionary. They won't affect totals but may need classification.
                </p>
                ${diagnosticHtml}
            </div>
        `;
        }
        
        return `
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                <div>
                    <h3>Column Classification ${statusBadge}</h3>
                    <p class="pf-config-subtext">Taxonomy-driven expense categorization</p>
                </div>
                </div>
            <div style="padding: 12px 16px; font-size: 12px;">
                <p><strong>Dimensions:</strong> ${dimensionColumns?.length || 0} columns (grouping)</p>
                <p><strong>Measures:</strong> ${measureColumns?.length || 0} columns ${bucketSummary ? `(${bucketSummary})` : ""}</p>
                ${unclassifiedHtml}
                </div>
            </article>
    `;
}

/**
 * Debug panel showing UNIFIED Step 1 <-> Expense Review reconciliation
 */
function renderExpenseReviewDebugPanel() {
    // Get current period from state
    const currentPeriod = expenseReviewState.periods?.[0];
    const hasData = !!currentPeriod;
    
    if (!hasData) {
        return ""; // Don't show debug panel if no data
    }
    
    const fmt = (val) => `$${Math.round(val || 0).toLocaleString()}`;
    
    // Extract values that UI is using
    const uiFixed = currentPeriod.summary?.fixed || 0;
    const uiVariable = currentPeriod.summary?.variable || 0;
    const uiBurden = currentPeriod.summary?.burden || 0;
    const uiTotal = currentPeriod.summary?.total || (uiFixed + uiVariable + uiBurden);
    
    // Step 1 measure universe (single source of truth)
    const step1Total = expenseReviewState.measureUniverseTotal || 0;
    const step1Measures = expenseReviewState.measureUniverseHeaders?.length || 0;
    
    // Additional state values
    const measureCols = expenseReviewState.measureColumns || [];
    const measureCount = measureCols.length;
    const unclassifiedCount = expenseReviewState.unclassifiedColumns?.length || 0;
    
    // Delta calculation
    const delta = Math.abs(step1Total - uiTotal);
    const isReconciled = delta < 1;
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" style="border: 1px dashed rgba(234, 179, 8, 0.5); background: rgba(234, 179, 8, 0.05);">
            <div class="pf-config-head">
                <h3>Reconciliation: Step 1 vs Expense Review</h3>
                <p class="pf-config-subtext">Expense Review must match Step 1's PR_Data_Clean total</p>
            </div>
            <div style="font-family: monospace; font-size: 12px; line-height: 1.8;">
                <div><strong>Step 1 PR_Data_Clean (Measure Universe):</strong></div>
                <div style="padding-left: 16px;">
                    Total: <strong>${fmt(step1Total)}</strong><br>
                    Measures: ${step1Measures} columns
                </div>
                <br>
                <div><strong>Expense Review (UI Display):</strong></div>
                <div style="padding-left: 16px;">
                    FIXED: ${fmt(uiFixed)}<br>
                    VARIABLE: ${fmt(uiVariable)}<br>
                    BURDEN: ${fmt(uiBurden)}<br>
                    <strong>TOTAL: ${fmt(uiTotal)}</strong><br>
                    Measures: ${measureCount} columns${unclassifiedCount > 0 ? ` (${unclassifiedCount} unclassified)` : ""}
                </div>
                <br>
                <div style="padding: 8px; border-radius: 4px; background: ${isReconciled ? 'rgba(34, 197, 94, 0.1)' : 'rgba(239, 68, 68, 0.1)'}; color: ${isReconciled ? '#22c55e' : '#ef4444'}">
                    <strong>RECONCILIATION:</strong> 
                    ${isReconciled 
                        ? "Step 1 and Expense Review match" 
                        : `Delta: ${fmt(delta)} (investigate)`}
                </div>
            </div>
            <div style="margin-top: 12px;">
                <button type="button" class="pf-action-toggle pf-action-toggle--subtle" id="expense-trace-btn" title="Run full diagnostic trace">
                    Run Trace
                </button>
            </div>
        </article>
    `;
}

function renderExpenseReviewStep(detail) {
    const stepFields = getStepNoteFields(2);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[2]));
    const statusBanner = expenseReviewState.loading
        ? `<p class="pf-step-note">Preparing executive summary…</p>`
        : expenseReviewState.lastError
            ? `<p class="pf-step-note">${escapeHtml(expenseReviewState.lastError)}</p>`
            : "";
    
    // Ada will be available via floating button instead of embedded card
    
    // TASK 3: Debug output - show where the UI total comes from
    const debugTotalsMarkup = renderExpenseReviewDebugPanel();

    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${statusBanner}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Perform Analysis</h3>
                    <p class="pf-config-subtext">Populate Expense Review and perform review.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle" id="expense-run-btn" title="Run expense review analysis">${CALCULATOR_ICON_SVG}</button>`,
                        "Run"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle" id="expense-refresh-btn" title="Refresh expense data">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                </div>
            </article>
            ${/* HIDDEN: Column Classification - internal tool, not customer-facing
            ${renderTaxonomyAdvisoryCard()}
            */ ''}
            ${renderPayrollCompletenessCard()}
            ${/* HIDDEN: Debug reconciliation panel - development tool only
            ${debugTotalsMarkup}
            */ ''}
            ${
                stepFields
                    ? `
            <!-- Ada Assistant Card -->
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-ada-card" id="expense-ada-btn" title="Ask Ada about payroll data">
                    <img src="${BRANDING.ADA_IMAGE_URL}" alt="Ada" class="pf-ada-icon" onerror="this.style.display='none'" />
                    <div class="pf-ada-title">Ask Ada</div>
                    <div class="pf-ada-subtitle">Your smart assistant to help troubleshoot, review, and analyze</div>
                </div>
            </article>
            ${renderInlineNotes({
                textareaId: "step-notes-input",
                value: stepNotes,
                permanentId: "step-notes-permanent",
                isPermanent: isFieldPermanent(stepFields.note),
                saveButtonId: "step-notes-save-2"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-name",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-2",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-2",
                completeButtonId: "expense-signoff-toggle"
            })}
            `
                    : ""
            }
        </section>
    `;
}

function renderJournalStep(detail) {
    const stepFields = getStepNoteFields(3);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const notesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[3]));
    const statusNote = journalState.lastError
        ? `<p class="pf-step-note" style="color: #ef4444;">${escapeHtml(journalState.lastError)}</p>`
        : "";
    
    // Validation state (consistent with PTO module)
    const hasRun = journalState.validationRun;
    const issues = journalState.issues || [];
    
    // Unmapped columns (payroll-specific)
    const unmappedColumns = journalState.unmappedColumns || [];
    const unmappedTotal = journalState.unmappedTotal || 0;
    const hasUnmapped = unmappedColumns.length > 0;
    
    // Define check descriptions (consistent with PTO module)
    const checkDefinitions = [
        { key: "Debits = Credits", desc: "∑ Debits = ∑ Credits" },
        { key: "JE Matches Source Total", desc: "∑ JE expense = ∑ PR_Data_Clean" },
        { key: "All Columns Mapped", desc: "All column+dept combos have GL mappings" }
    ];
    
    // Helper to render check rows (consistent with PTO module)
    const renderCheckRow = (def) => {
        const issue = issues.find(i => i.check === def.key);
        const pending = !hasRun;
        let circleHtml;
        
        if (pending) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (issue?.passed) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(def.desc)}</span>
            </div>
        `;
    };
    
    const checkRows = checkDefinitions.map(def => renderCheckRow(def)).join("");
    
    // Build issues card if there are failures (consistent with PTO module)
    const failedIssues = issues.filter(i => !i.passed);
    let issuesCard = "";
    if (hasRun && failedIssues.length > 0) {
        issuesCard = `
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${failedIssues.map(i => `<li><strong>${escapeHtml(i.check)}:</strong> ${escapeHtml(i.detail)}</li>`).join("")}
                </ul>
            </article>
        `;
    }
    
    // Unmapped columns detail panel (payroll-specific)
    const fmt = (val) => formatNumberDisplay(Math.round(val || 0));
    const unmappedPanelHtml = hasUnmapped ? `
        <article class="pf-step-card pf-step-detail" style="border: 1px solid #f59e0b; background: rgba(245, 158, 11, 0.05);">
            <div class="pf-config-head">
                <h3 style="color: #f59e0b;">Unmapped GL Combinations</h3>
                <p class="pf-config-subtext">These column + department combinations need GL account mappings in <code>ada_customer_gl_mappings</code>.</p>
            </div>
            <div style="max-height: 200px; overflow-y: auto;">
                <table style="width: 100%; font-size: 12px; border-collapse: collapse;">
                    <thead>
                        <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);">
                            <th style="text-align: left; padding: 4px 8px;">PF Column Name</th>
                            <th style="text-align: left; padding: 4px 8px;">Department</th>
                            <th style="text-align: right; padding: 4px 8px;">Amount</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${unmappedColumns.slice(0, 15).map(col => `
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.05);">
                                <td style="padding: 4px 8px; font-family: monospace; font-size: 11px;">${escapeHtml(col.header)}</td>
                                <td style="padding: 4px 8px; font-size: 11px;">${escapeHtml(col.department || "(none)")}</td>
                                <td style="text-align: right; padding: 4px 8px;">${fmt(col.total)}</td>
                            </tr>
                        `).join("")}
                        ${unmappedColumns.length > 15 ? `
                            <tr><td colspan="3" style="padding: 4px 8px; color: #888;">... and ${unmappedColumns.length - 15} more</td></tr>
                        ` : ""}
                    </tbody>
                    <tfoot>
                        <tr style="border-top: 1px solid rgba(255,255,255,0.1); font-weight: bold;">
                            <td colspan="2" style="padding: 4px 8px;">Total Unmapped</td>
                            <td style="text-align: right; padding: 4px 8px;">${fmt(unmappedTotal)}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </article>
    ` : "";

    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">Generate journal entry for QuickBooks import.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Build balanced JE from PR_Data_Clean using your GL mappings.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate JE from PR_Data_Clean + GL mappings">${TABLE_ICON_SVG}</button>`,
                        "Generate"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${statusNote}
                <div class="pf-je-checks-container">
                    ${checkRows}
                </div>
            </article>
            ${issuesCard}
            ${unmappedPanelHtml}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head" style="position: relative;">
                    <div>
                        <h3>Export Journal Entry</h3>
                        <p class="pf-config-subtext">Download journal entry as CSV for QuickBooks import.</p>
                    </div>
                    <button type="button" class="pf-info-icon-btn" id="je-info-btn" aria-label="Export instructions" style="position: absolute; top: 0; right: 0;">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="12" y1="16" x2="12" y2="12"></line>
                            <line x1="12" y1="8" x2="12.01" y2="8"></line>
                        </svg>
                    </button>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${DOWNLOAD_ICON_SVG}</button>`,
                        "Export CSV"
                    )}
                </div>
            </article>
            
            <!-- Info Modal for JE Export Instructions -->
            <div id="je-info-modal" class="pf-modal-overlay" style="display: none;">
                <div class="pf-modal-content">
                    <div class="pf-modal-header">
                        <h3>📋 QuickBooks Import Instructions</h3>
                        <button type="button" class="pf-modal-close" id="je-info-close">&times;</button>
                    </div>
                    <div class="pf-modal-body">
                        <ol style="margin: 0 0 16px 20px; padding: 0; line-height: 1.8;">
                            <li style="margin-bottom: 12px;">
                                <strong>Assign bank feed transaction</strong><br>
                                <span style="color: #9ca3af;">Assign the bank feed transaction to uncategorized expense (no need to separate this out like in the past)</span>
                            </li>
                            <li style="margin-bottom: 12px;">
                                <strong>Export CSV file</strong><br>
                                <span style="color: #9ca3af;">Click the Export button above and save the .csv file to your desktop (or other temporary folder)</span>
                            </li>
                            <li style="margin-bottom: 12px;">
                                <strong>Upload to QuickBooks</strong><br>
                                <span style="color: #9ca3af;">Click the ⚙️ icon → Import Data → Journal Entry → Upload a file to import data. Map any fields that don't automap.</span>
                            </li>
                        </ol>
                        <hr style="border: none; border-top: 1px solid rgba(255,255,255,0.15); margin: 16px 0;">
                        <div style="font-size: 13px; color: #9ca3af; line-height: 1.6;">
                            <p style="margin: 0 0 12px 0;">
                                <strong style="color: #fbbf24;">💡 Note:</strong> The journal entry will also be booked to uncategorized expense. If everything goes as planned, uncategorized expense should be zero after recording the bank transaction and this journal entry.
                            </p>
                            <p style="margin: 0;">
                                If there are differences, they may indicate a fee that was charged but not presented in the payroll report. This should be recorded as a separate journal entry.
                            </p>
                        </div>
                    </div>
                </div>
            </div>
            
            ${stepFields ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-input",
                    value: stepNotes || "",
                    permanentId: "step-notes-permanent",
                    isPermanent: notesPermanent,
                    saveButtonId: "step-notes-save-3"
                })}
                ${renderSignoff({
                    reviewerInputId: "step-reviewer-name",
                    reviewerValue: stepReviewer,
                    signoffInputId: "step-signoff-3",
                    signoffValue: stepSignOff,
                    isComplete: stepComplete,
                    saveButtonId: "step-signoff-save-3",
                    completeButtonId: "step-signoff-toggle-3"
                })}
            `
                    : ""
            }
        </section>
    `;
}

function renderArchiveStep(detail) {
    const completionItems = WORKFLOW_STEPS.filter((step) => step.id !== 4).map((step) => ({
        id: step.id,
        title: step.title,
        complete: isStepCompleteFromConfig(step.id)
    }));
    const allComplete = completionItems.every((item) => item.complete);
    const incompleteCount = completionItems.filter(i => !i.complete).length;
    
    // Debug: Log completion status
    console.log("[Archive Step] Completion check:", completionItems.map(i => `Step ${i.id}: ${i.complete}`).join(", "));
    console.log("[Archive Step] All complete:", allComplete, "Incomplete count:", incompleteCount);
    const statusList = completionItems
        .map(
            (item) => `
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${item.complete ? "is-active" : ""}" aria-pressed="${item.complete}">
                        ${CHECK_ICON_SVG}
                    </span>
                    <div>
                        <h3>${escapeHtml(item.title)}</h3>
                        <p class="pf-config-subtext">${item.complete ? "Complete" : "Not complete"}</p>
                    </div>
                </div>
            </article>
        `
        )
        .join("");
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${statusList}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Create an archive of this module's sheets and clear work tabs.</p>
                </div>
                ${!allComplete ? `<p class="pf-step-note pf-step-note--info">Complete all ${incompleteCount} remaining step(s) before archiving. Click the checkmark on each step's sign-off section.</p>` : `<p class="pf-step-note" style="color: #22c55e;">All steps complete. Ready to archive!</p>`}
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn ${allComplete ? 'pf-cta-button' : ''}" id="archive-run-btn" ${allComplete ? "" : "disabled"} onclick="window.handleArchiveClick()" style="${allComplete ? '' : 'opacity: 0.5; cursor: not-allowed;'}">
                        ${allComplete ? 'Archive Now' : 'Archive (Complete Steps First)'}
                    </button>
                </div>
            </article>
        </section>
    `;
}
function renderStepView(stepId) {
    const detail =
        STEP_DETAILS.find((step) => step.id === stepId) || {
            id: stepId ?? "-",
            title: "Workflow Step",
            summary: "",
            description: "",
            checklist: []
        };
    if (stepId === 1) return renderImportStep(detail);
    if (stepId === 2) return renderExpenseReviewStep(detail);
    if (stepId === 3) return renderJournalStep(detail);
    if (stepId === 4) return renderArchiveStep(detail);
    const isStepOne = false; // Step 1 now has dedicated render
    const stepFields = getStepNoteFields(stepId);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepNotesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete =
        stepFields && STEP_COMPLETE_FIELDS[stepId]
            ? Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[stepId]))
            : Boolean(stepSignOff);
    const highlights = (detail.highlights || [])
        .map(
            (item) => `
            <div class="pf-step-highlight">
                <span class="pf-step-highlight-label">${escapeHtml(item.label)}</span>
                <span class="pf-step-highlight-detail">${escapeHtml(item.detail)}</span>
            </div>
        `
        )
        .join("");
    const checklist =
        (detail.checklist || [])
            .map((item) => `<li>${escapeHtml(item)}</li>`)
            .join("") || "";
    const descriptionText = isStepOne
        ? ""
        : detail.description || "Detailed guidance will appear here.";
    const actionButtons = [];
    if (!isStepOne && detail.ctaLabel) {
        actionButtons.push(
            `<button type="button" class="pf-pill-btn" id="step-action-btn">${escapeHtml(detail.ctaLabel)}</button>`
        );
    }
    if (!isStepOne) {
        actionButtons.push(
            `<button type="button" class="pf-pill-btn pf-pill-btn--ghost" id="step-back-btn">Back to Step List</button>`
        );
    }
    const actionSection = actionButtons.length
        ? `<div class="pf-pill-row pf-config-actions">${actionButtons.join("")}</div>`
        : "";
    const providerLink = getPayrollProviderLink();
    const providerSection = isStepOne
        ? `
            <div class="pf-link-card">
                <h3 class="pf-link-card__title">Payroll Reports</h3>
                <p class="pf-link-card__subtitle">Open your latest payroll export.</p>
                <div class="pf-link-list">
                    <a
                        class="pf-link-item"
                        id="pr-provider-link"
                        ${providerLink ? `href="${escapeHtml(providerLink)}" target="_blank" rel="noopener noreferrer"` : `aria-disabled="true"`}
                    >
                        <span class="pf-link-item__icon">${LINK_ICON_SVG}</span>
                        <span class="pf-link-item__body">
                            <span class="pf-link-item__title">Open Payroll Export</span>
                            <span class="pf-link-item__meta">${escapeHtml(
                                providerLink || "Add a provider link in Configuration"
                            )}</span>
                        </span>
                    </a>
                </div>
            </div>
        `
        : "";
    const quickTipsSection = "";
    const highlightSection =
        !isStepOne && highlights ? `<article class="pf-step-card pf-step-detail">${highlights}</article>` : "";
    const checklistSection =
        !isStepOne && checklist
            ? `<article class="pf-step-card pf-step-detail">
                            <h3 class="pf-step-subtitle">Checklist</h3>
                            <ul class="pf-step-checklist">${checklist}</ul>
                        </article>`
            : "";
    const descriptionSection =
        !isStepOne || descriptionText || actionSection
            ? `
            <article class="pf-step-card pf-step-detail">
                <p class="pf-step-title">${escapeHtml(descriptionText)}</p>
                ${!isStepOne && detail.statusHint ? `<p class="pf-step-note">${escapeHtml(detail.statusHint)}</p>` : ""}
                ${actionSection}
            </article>
        `
            : "";
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <p class="pf-hero-hint">${escapeHtml(appState.statusText || "")}</p>
        </section>
        <section class="pf-step-guide">
            ${providerSection}
            ${quickTipsSection}
            ${descriptionSection}
            ${highlightSection}
            ${checklistSection}
            ${
                stepFields
                    ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-input",
                    value: stepNotes,
                    permanentId: "step-notes-permanent",
                    isPermanent: stepNotesPermanent,
                    saveButtonId: "step-notes-save"
                })}
                ${renderSignoff({
                    reviewerInputId: "step-reviewer-name",
                    reviewerValue: stepReviewer,
                    signoffInputId: `step-signoff-${stepId}`,
                    signoffValue: stepSignOff,
                    isComplete: stepComplete,
                    saveButtonId: `step-signoff-save-${stepId}`,
                    completeButtonId: `step-signoff-toggle-${stepId}`,
                    subtext: "Ready to move on? Save and click Done when finished."
                })}
            `
                    : ""
            }
        </section>
    `;
}

function renderStepCard(step, index) {
    const isActive = appState.focusedIndex === index ? "pf-step-card--active" : "";
    const icon = getStepIconSvg(getStepType(step.id));
    return `
        <article class="pf-step-card pf-clickable ${isActive}" data-step-card data-step-index="${index}" data-step-id="${step.id}">
            <p class="pf-step-index">Step ${step.id}</p>
            <h3 class="pf-step-title">${icon ? `${icon}` : ""}${escapeHtml(step.title)}</h3>
        </article>
    `;
}

function renderFooter() {
    return `
        <footer class="pf-brand-footer">
            <div class="pf-brand-text">
                <div class="pf-brand-label">prairie.forge</div>
                <div class="pf-brand-meta">© Prairie Forge LLC, 2025. All rights reserved. Version ${VERSION}</div>
            </div>
        </footer>
    `;
}

function getStepType(stepId) {
    if (stepId === 0) return "config";
    if (stepId === 1) return "import";
    if (stepId === 2) return "review";
    if (stepId === 3) return "journal";
    if (stepId === 4) return "archive";
    return "";
}

function bindSharedInteractions() {
    document.getElementById("nav-home")?.addEventListener("click", () => {
        returnHome();
        document.getElementById("pf-hero")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("nav-selector")?.addEventListener("click", () => {
        window.location.href = "../module-selector/index.html";
    });
    document.getElementById("nav-prev")?.addEventListener("click", () => moveFocus(-1));
    document.getElementById("nav-next")?.addEventListener("click", () => moveFocus(1));
    
    // Quick Access Modal toggle
    const quickToggle = document.getElementById("nav-quick-toggle");
    const quickModal = document.getElementById("quick-access-modal");
    const quickClose = document.getElementById("quick-access-close");
    const quickBackdrop = quickModal?.querySelector(".pf-quick-modal-backdrop");
    
    const closeQuickAccess = () => {
        quickModal?.classList.add("hidden");
        quickToggle?.classList.remove("is-active");
    };
    
    quickToggle?.addEventListener("click", (e) => {
        e.stopPropagation();
        quickModal?.classList.toggle("hidden");
        quickToggle.classList.toggle("is-active");
    });
    
    // Close button
    quickClose?.addEventListener("click", closeQuickAccess);
    
    // Click backdrop to close
    quickBackdrop?.addEventListener("click", closeQuickAccess);
    
    // Quick Access - Payroll Provider Report (link handles itself)
    
    // Quick Access - Employee Roster
    document.getElementById("nav-employee-roster")?.addEventListener("click", async () => {
        closeQuickAccess();
        await showAndActivateSheet("SS_Employee_Roster");
    });
    
    // Quick Access - Chart of Accounts
    document.getElementById("nav-chart-of-accounts")?.addEventListener("click", async () => {
        closeQuickAccess();
        await showAndActivateSheet("SS_Chart_of_Accounts");
    });
    
    // Quick Access - Accounting Software
    document.getElementById("nav-accounting-software")?.addEventListener("click", async () => {
        closeQuickAccess();
        await openAccountingSoftware();
    });
    
    // Quick Access - Configuration
    document.getElementById("nav-config")?.addEventListener("click", async () => {
        closeQuickAccess();
        await showAndActivateSheet("SS_PF_Config");
    });
    
}

/**
 * Fetch configuration sheets (SS_* and any with "mapping" in name)
 */
async function getConfigurationSheets() {
    if (typeof Excel === "undefined") {
        return [];
    }
    try {
        return await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            const matches = worksheets.items.filter((sheet) => {
                const name = sheet.name || "";
                const upper = name.toUpperCase();
                return upper.startsWith("SS_") || upper.includes("MAPPING") || upper.includes("HOMEPAGE");
            });
            return matches
                .map((sheet) => ({
                    name: sheet.name,
                    visible: sheet.visibility === Excel.SheetVisibility.visible,
                    isHomepage: (sheet.name || "").toUpperCase().includes("HOMEPAGE")
                }))
                .sort((a, b) => {
                    if (a.isHomepage && !b.isHomepage) return 1;
                    if (!a.isHomepage && b.isHomepage) return -1;
                    return a.name.localeCompare(b.name);
                });
        });
    } catch (error) {
        console.error("[Config] Error reading configuration sheets:", error);
        return [];
    }
}

/**
 * Ensure config modal exists in DOM
 */
function ensureConfigModal() {
    if (document.getElementById("config-sheet-modal")) return;
    const modal = document.createElement("div");
    modal.id = "config-sheet-modal";
    modal.className = "pf-config-modal hidden";
    modal.innerHTML = `
        <div class="pf-config-modal-backdrop" data-close></div>
        <div class="pf-config-modal-card">
            <div class="pf-config-modal-head">
                <h3>Configuration Sheets</h3>
                <button type="button" class="pf-config-close" data-close aria-label="Close">×</button>
            </div>
            <div class="pf-config-modal-body">
                <p class="pf-config-hint">Choose a configuration or mapping sheet to unhide and open.</p>
                <div id="config-sheet-list" class="pf-config-sheet-list">Loading…</div>
            </div>
        </div>
    `;
    document.body.appendChild(modal);

    if (!document.getElementById("pf-config-modal-styles")) {
        const style = document.createElement("style");
        style.id = "pf-config-modal-styles";
        style.textContent = `
            .pf-config-modal { position: fixed; inset: 0; display: flex; align-items: center; justify-content: center; z-index: 10000; }
            .pf-config-modal.hidden { display: none; }
            .pf-config-modal-backdrop { position: absolute; inset: 0; background: rgba(0,0,0,0.6); }
            .pf-config-modal-card { position: relative; background: var(--glass-bg); backdrop-filter: var(--glass-blur); -webkit-backdrop-filter: var(--glass-blur); color: #f8fafc; border-radius: 12px; padding: 22px; width: min(440px, 90%); box-shadow: 0 20px 60px rgba(0,0,0,0.35); }
            .pf-config-modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
            .pf-config-close { background: transparent; border: none; color: #f8fafc; font-size: 20px; cursor: pointer; }
            .pf-config-hint { margin: 0 0 12px 0; color: #cbd5e1; font-size: 14px; }
            .pf-config-sheet-list { display: flex; flex-direction: column; gap: 10px; max-height: 260px; overflow-y: auto; }
            .pf-config-sheet { display: flex; justify-content: space-between; align-items: center; padding: 12px 14px; background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.18); border-radius: 10px; cursor: pointer; color: #e2e8f0; font-weight: 600; }
            .pf-config-sheet:hover { background: rgba(255,255,255,0.16); }
            .pf-config-pill { font-size: 12px; color: #c7d2fe; }
        `;
        document.head.appendChild(style);
    }
}

/**
 * Open config modal and populate with config/mapping sheets
 */
async function openConfigModal() {
    ensureConfigModal();
    const modal = document.getElementById("config-sheet-modal");
    const list = document.getElementById("config-sheet-list");
    if (!modal || !list) return;

    list.textContent = "Loading…";
    modal.classList.remove("hidden");

    const sheets = await getConfigurationSheets();
    if (!sheets.length) {
        list.textContent = "No configuration sheets found.";
    } else {
        list.innerHTML = "";
        sheets.forEach((sheet) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "pf-config-sheet";
            btn.innerHTML = `<span>${sheet.name}</span><span class="pf-config-pill">${sheet.visible ? "Visible" : "Hidden"}</span>`;
            btn.addEventListener("click", async () => {
                await openConfigSheet(sheet.name);
                modal.classList.add("hidden");
            });
            list.appendChild(btn);
        });
    }

    modal.querySelectorAll("[data-close]").forEach((el) =>
        el.addEventListener("click", () => modal.classList.add("hidden"))
    );
}

/**
 * Open a config sheet (unhide if needed)
 */
async function openConfigSheet(sheetName) {
    if (!sheetName || typeof Excel === "undefined") return;
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            let sheet = worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject,visibility");
            await context.sync();

            if (sheet.isNullObject) {
                sheet = worksheets.add(sheetName);
            }
            sheet.visibility = Excel.SheetVisibility.visible;
            await context.sync();

            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
            console.log(`[Config] Opened sheet: ${sheetName}`);
        });
    } catch (error) {
        console.error("[Config] Error opening sheet", sheetName, error);
    }
}

/**
 * Unhide system sheets (SS_* prefix) for configuration access
 */
async function unhideSystemSheets() {
    if (typeof Excel === "undefined") {
        console.log("Excel not available");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            let unhiddenCount = 0;
            worksheets.items.forEach((sheet) => {
                if (sheet.name.toUpperCase().startsWith("SS_")) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                    console.log(`[Config] Made visible: ${sheet.name}`);
                    unhiddenCount++;
                }
            });
            
            await context.sync();
            
            // Activate SS_PF_Config if it exists
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (!configSheet.isNullObject) {
                configSheet.activate();
                configSheet.getRange("A1").select();
                await context.sync();
            }
            
            console.log(`[Config] ${unhiddenCount} system sheets now visible`);
        });
    } catch (error) {
        console.error("[Config] Error unhiding system sheets:", error);
    }
}

/**
 * Open a reference data sheet (creates if doesn't exist)
 * Makes sheet visible first if it's hidden
 */
async function openReferenceSheet(sheetName) {
    if (!sheetName || typeof Excel === "undefined") {
        return;
    }
    
    const defaultHeaders = {
        "SS_Employee_Roster": ["Employee", "Department", "Pay_Rate", "Status", "Hire_Date"],
        "SS_Chart_of_Accounts": ["Account_Number", "Account_Name", "Type", "Category"]
    };
    
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject,visibility");
            await context.sync();
            
            if (sheet.isNullObject) {
                // Create the sheet with default headers
                sheet = context.workbook.worksheets.add(sheetName);
                const headers = defaultHeaders[sheetName] || ["Column1", "Column2"];
                const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = "#f0f0f0";
                headerRange.format.autofitColumns();
                await context.sync();
            } else {
                // Make sure sheet is visible before activating (it may be hidden by tab visibility)
                sheet.visibility = Excel.SheetVisibility.visible;
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
            console.log(`[Quick Access] Opened sheet: ${sheetName}`);
        });
    } catch (error) {
        console.error("Error opening reference sheet:", error);
    }
}

/**
 * Navigate directly to the PR_Expense_Mapping sheet
 */
async function navigateToExpenseMapping() {
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject("PR_Expense_Mapping");
            sheet.load("isNullObject,visibility");
            await context.sync();
            
            if (sheet.isNullObject) {
                // Create the sheet with default headers
                sheet = context.workbook.worksheets.add("PR_Expense_Mapping");
                const headers = ["Expense_Category", "GL_Account", "Description", "Active"];
                const headerRange = sheet.getRange("A1:D1");
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
            } else {
                // Make sure sheet is visible before activating
                sheet.visibility = Excel.SheetVisibility.visible;
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
            console.log("[Quick Access] Opened PR_Expense_Mapping");
        });
    } catch (error) {
        console.error("Error navigating to PR_Expense_Mapping:", error);
    }
}

function bindHomeInteractions() {
    document.querySelectorAll("[data-step-card]").forEach((card) => {
        const index = Number(card.getAttribute("data-step-index"));
        card.addEventListener("click", () => focusStep(index));
    });
}

function bindConfigInteractions() {
    // User Name field - writes to SS_PF_Config as PR_Reviewer_Name
    const userNameInput = document.getElementById("config-user-name");
    userNameInput?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        scheduleConfigWrite(CONFIG_REVIEWER_FIELD, value);
        // Also update the reviewer field in the signoff section if it's empty
        const reviewerInput = document.getElementById("config-reviewer-name");
        if (reviewerInput && !reviewerInput.value) {
            reviewerInput.value = value;
        }
    });

    // Initialize custom date picker for Payroll Date
    initDatePicker("config-payroll-date", {
        onChange: (value) => {
            // Always use the primary field name to avoid duplicate rows
            scheduleConfigWrite("PR_Payroll_Date", value);
            
            // Clear step completion when payroll date changes (user needs to re-confirm)
            clearStepCompletion(0);
            
        if (!value) return;
            
            // Always derive and update Accounting Period and JE ID when payroll date changes
            // This ensures they stay in sync even if user had previously set values
            const derivedPeriod = deriveAccountingPeriod(value);
            if (derivedPeriod) {
                const periodInput = document.getElementById("config-accounting-period");
                if (periodInput) periodInput.value = derivedPeriod;
                scheduleConfigWrite("PR_Accounting_Period", derivedPeriod);
                // Reset override flag so future date changes will also update
                configState.overrides.accountingPeriod = false;
            }
            
            const derivedJe = deriveJeId(value);
            if (derivedJe) {
                const jeInput = document.getElementById("config-je-id");
                if (jeInput) jeInput.value = derivedJe;
                scheduleConfigWrite("PR_Journal_Entry_ID", derivedJe);
                // Reset override flag so future date changes will also update
                configState.overrides.jeId = false;
            }
        }
    });

    const stepFields = getStepNoteFields(0);

    const periodInput = document.getElementById("config-accounting-period");
    periodInput?.addEventListener("change", (event) => {
        configState.overrides.accountingPeriod = Boolean(event.target.value);
        scheduleConfigWrite("PR_Accounting_Period", event.target.value || "");
        clearStepCompletion(0); // Clear "Done" when field changes
    });

    const jeInput = document.getElementById("config-je-id");
    jeInput?.addEventListener("change", (event) => {
        configState.overrides.jeId = Boolean(event.target.value);
        scheduleConfigWrite("PR_Journal_Entry_ID", event.target.value.trim());
        clearStepCompletion(0); // Clear "Done" when field changes
    });

    document.getElementById("config-company-name")?.addEventListener("change", (event) => {
        scheduleConfigWrite("SS_Company_Name", event.target.value.trim());
    });

    document.getElementById("config-company-id")?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        // Basic UUID validation
        const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
        if (value && !uuidRegex.test(value)) {
            showToast("Company ID should be a valid UUID from the CRM.", "error");
            return;
        }
        scheduleConfigWrite("SS_Company_ID", value);
        if (value) {
            showToast("Company ID saved. Ada will now remember your column mappings.", "success");
        }
    });

    document.getElementById("config-payroll-provider")?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        scheduleConfigWrite(PAYROLL_PROVIDER_FIELD, value);
    });

    document.getElementById("config-accounting-link")?.addEventListener("change", (event) => {
        scheduleConfigWrite("SS_Accounting_Software", event.target.value.trim());
    });

    const notesInput = document.getElementById("config-notes");
    notesInput?.addEventListener("input", (event) => {
        if (stepFields) {
            scheduleConfigWrite(stepFields.note, event.target.value, { debounceMs: 400 });
        }
    });

    if (stepFields) {
        const lockButton = document.getElementById("config-notes-permanent");
        if (lockButton) {
            lockButton.addEventListener("click", () => {
                const nextState = !lockButton.classList.contains("is-locked");
                updateLockButtonVisual(lockButton, nextState);
                setNotePermanent(stepFields.note, nextState);
            });
            updateLockButtonVisual(lockButton, isFieldPermanent(stepFields.note));
        }

        const notesSaveBtn = document.getElementById("config-notes-save");
        notesSaveBtn?.addEventListener("click", () => {
            if (!notesInput) return;
            scheduleConfigWrite(stepFields.note, notesInput.value);
            updateSaveButtonState(notesSaveBtn, true);
        });
    }

    const reviewerInput = document.getElementById("config-reviewer-name");
    reviewerInput?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        if (stepFields) {
            scheduleConfigWrite(stepFields.reviewer, value);
        }
        scheduleConfigWrite(CONFIG_REVIEWER_FIELD, value);
        const signoffInput = document.getElementById("config-signoff-date");
        if (value && signoffInput && !signoffInput.value) {
            const today = todayIso();
            signoffInput.value = today;
            if (stepFields) {
                scheduleConfigWrite(stepFields.signOff, today);
            }
        }
    });

    document.getElementById("config-signoff-date")?.addEventListener("change", (event) => {
        if (stepFields) {
            scheduleConfigWrite(stepFields.signOff, event.target.value || "");
        }
    });

    const signoffSaveBtn = document.getElementById("config-signoff-save");
    signoffSaveBtn?.addEventListener("click", () => {
        const reviewerValue = reviewerInput?.value?.trim() || "";
        const signoffInput = document.getElementById("config-signoff-date");
        const signoffValue = signoffInput?.value || "";
        if (stepFields) {
            scheduleConfigWrite(stepFields.reviewer, reviewerValue);
            scheduleConfigWrite(stepFields.signOff, signoffValue);
        }
        scheduleConfigWrite(CONFIG_REVIEWER_FIELD, reviewerValue);
        updateSaveButtonState(signoffSaveBtn, true);
    });

    initSaveTracking();

    if (stepFields) {
        // Calculate isStepComplete for Step 0
        const signOffValue = getConfigValue(stepFields.signOff);
        const completeValue = getConfigValue(STEP_COMPLETE_FIELDS[0]);
        const isStepComplete = Boolean(signOffValue || completeValue === "Y" || completeValue === true);
        console.log(`[Step 0] Binding signoff toggle. signOff="${signOffValue}", complete="${completeValue}", isComplete=${isStepComplete}`);
        
        bindSignoffToggle({
            buttonId: "config-signoff-toggle",
            inputId: "config-signoff-date",
            fieldName: stepFields.signOff,
            completeField: STEP_COMPLETE_FIELDS[0],
            initialActive: isStepComplete,
            stepId: 0, // Step 0 has no prerequisites
            onComplete: getStepCompleteHandler(0)
        });
        bindSignoffNavButtons("config-signoff-toggle-prev", "config-signoff-toggle-next");
    }
}

function bindStepInteractions(stepId) {
    // Debug: Log which step is being bound
    if (stepId === 4) {
        console.log("[Archive DEBUG] bindStepInteractions called for step 4");
    }
    
    document.getElementById("step-back-btn")?.addEventListener("click", () => {
        returnHome();
    });
    document.getElementById("step-action-btn")?.addEventListener("click", () => {
        const detail = STEP_DETAILS.find((step) => step.id === stepId);
        showToast(detail?.ctaLabel ? `${detail.ctaLabel} coming soon.` : "Step actions coming soon.", "info");
    });

    if (stepId === 1) {
        document.getElementById("import-open-data-btn")?.addEventListener("click", () => openDataSheet());
        document.getElementById("import-clear-btn")?.addEventListener("click", () => clearPrDataSheet());
        
        // File upload interactions
        bindFileUploadInteractions();
        
        // Bank reconciliation input (Step 1)
        document.getElementById("step1-bank-amount-input")?.addEventListener("blur", handleStep1BankAmountInput);
        document.getElementById("step1-bank-amount-input")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") handleStep1BankAmountInput(e);
        });
        
        // Bank reconciliation refresh button - preserve scroll position
        document.getElementById("bank-recon-refresh-btn")?.addEventListener("click", async () => {
            showToast("Refreshing bank reconciliation...", "info", 2000);
            const scrollY = window.scrollY;
            await refreshBankReconciliation();
            renderApp();
            // Restore scroll position after render
            requestAnimationFrame(() => window.scrollTo(0, scrollY));
        });
        
        // Payroll coverage refresh button - preserve scroll position
        document.getElementById("coverage-refresh-btn")?.addEventListener("click", async () => {
            showToast("Refreshing payroll coverage...", "info", 2000);
            const scrollY = window.scrollY;
            await refreshPayrollCoverage();
            renderApp();
            // Restore scroll position after render
            requestAnimationFrame(() => window.scrollTo(0, scrollY));
        });

        // Employee roster updates buttons
        document.getElementById("roster-apply-btn")?.addEventListener("click", async () => {
            await applyRosterUpdates();
        });
        document.getElementById("roster-refresh-btn")?.addEventListener("click", async () => {
            showToast("Refreshing roster analysis...", "info", 2000);
            await refreshRosterUpdates();
        });

        // Employee Coverage expandable bar - open reconciliation modal
        document.getElementById("employee-coverage-bar")?.addEventListener("click", (e) => {
            e.preventDefault();
            openEmployeeCoverageModal();
        });
        
        // Sign-off navigation buttons (Prev/Next at bottom of step)
        document.getElementById("step-signoff-toggle-1-prev")?.addEventListener("click", () => {
            moveFocus(-1); // Go to Config
        });
        document.getElementById("step-signoff-toggle-1-next")?.addEventListener("click", () => {
            moveFocus(1); // Go to Step 2
        });
    }
    if (stepId === 3) {
        document.getElementById("je-run-btn")?.addEventListener("click", () => runJournalSummary());
        document.getElementById("je-save-btn")?.addEventListener("click", () => saveJournalSummary());
        document.getElementById("je-create-btn")?.addEventListener("click", async () => {
            console.log("[JE] Generate button clicked");
            showToast("Starting journal entry generation...", "info", 2000);
            try {
                await createJournalEntryDraftV2();
            } catch (error) {
                console.error("[JE] Button click error:", error);
                showToast(`Error: ${error.message || "Journal entry generation failed"}`, "error", 8000);
            }
        });
        document.getElementById("je-export-btn")?.addEventListener("click", () => exportJournalDraft());
        
        // Info modal toggle
        document.getElementById("je-info-btn")?.addEventListener("click", () => {
            const modal = document.getElementById("je-info-modal");
            if (modal) modal.style.display = "flex";
        });
        
        // Info modal close button
        document.getElementById("je-info-close")?.addEventListener("click", () => {
            const modal = document.getElementById("je-info-modal");
            if (modal) modal.style.display = "none";
        });
        
        // Close modal when clicking overlay background
        document.getElementById("je-info-modal")?.addEventListener("click", (e) => {
            if (e.target.id === "je-info-modal") {
                e.target.style.display = "none";
            }
        });
        
        // Sign-off navigation buttons (Prev/Next at bottom of step)
        document.getElementById("step-signoff-toggle-3-prev")?.addEventListener("click", () => {
            moveFocus(-1); // Go to Step 2
        });
        document.getElementById("step-signoff-toggle-3-next")?.addEventListener("click", () => {
            moveFocus(1); // Go to Step 4
        });
    }
    if (stepId === 2) {
        const container = document.querySelector(".pf-step-guide");
        if (container) {
            // CoPilot API endpoint - Update this with your Supabase project URL
            const COPILOT_API_ENDPOINT = "https://your-project.supabase.co/functions/v1/copilot";
            
            // Bind CoPilot with full context provider for intelligent analysis
            bindCopilotCard(container, { 
                id: "expense-review-copilot",
                // Uncomment to enable real AI (requires Supabase Edge Function deployment)
                // apiEndpoint: COPILOT_API_ENDPOINT,
                contextProvider: createPayrollContextProvider(),
                systemPrompt: `You are Prairie Forge CoPilot, an expert financial analyst assistant for payroll expense review.

CONTEXT: You're embedded in the Payroll Recorder Excel add-in, helping accountants and CFOs review payroll data before journal entry export.

YOUR CAPABILITIES:
1. Analyze payroll expense data for accuracy and completeness
2. Identify trends, anomalies, and variances requiring attention
3. Prepare executive-ready insights and talking points
4. Validate journal entries before export to accounting system

COMMUNICATION STYLE:
- Be concise and actionable
- Use bullet points and tables for clarity
- Use plain text status labels (no emoji)
- Format currency as $X,XXX (no decimals for totals)
- Format percentages as X.X%
- Always end with 2-3 concrete next steps

ANALYSIS FOCUS:
- Period-over-period changes exceeding 10%
- Department cost anomalies vs historical norms
- Headcount vs payroll expense alignment
- Burden rate outliers (normal range: 15-35%)
- Missing or incomplete GL account mappings
- Data quality issues (blanks, duplicates, mismatches)

When asked about variances, explain the business drivers, not just the numbers.
When asked about readiness, be specific about what passes and what needs attention.`
            });
        }
        document.getElementById("expense-run-btn")?.addEventListener("click", () => {
            prepareExpenseReviewData();
        });
        document.getElementById("expense-refresh-btn")?.addEventListener("click", () => {
            prepareExpenseReviewData();
        });
        document.getElementById("expense-ada-btn")?.addEventListener("click", () => {
            // Import and call showAdaModal from homepage-sheet.js
            import("../../Common/homepage-sheet.js").then(module => {
                module.showAdaModal();
            });
        });
        
        // Debug trace button
        document.getElementById("expense-trace-btn")?.addEventListener("click", async () => {
            showToast("Running diagnostic trace...", "info", 3000);
            const trace = await traceExpenseReviewMeasurePipeline();
            console.log("[Trace] Result object:", trace);
            showToast("Trace complete - see console for full output", "success", 5000);
        });

        // Bind Ada copilot card
        // Note: bindCopilotCard expects a PARENT container where [data-copilot] element exists
        bindCopilotCard(document.body, {
            id: "payroll-copilot",
            contextProvider: createExcelContextProvider({
                dataClean: 'PR_Data_Clean',
                expenseReview: 'PR_Expense_Review',
                config: 'SS_PF_Config'
            }),
            onPrompt: callAdaApi
        });
        
        // Ada Insights panel interactions
        document.getElementById("ada-insights-collapse")?.addEventListener("click", () => {
            adaInsightsState.collapsed = !adaInsightsState.collapsed;
            renderApp();
        });
        
        // Ada quick prompt buttons
        document.querySelectorAll("[data-ada-prompt]").forEach(btn => {
            btn.addEventListener("click", async () => {
                const promptId = btn.dataset.adaPrompt;
                const quickPrompts = {
                    "main": null, // Default prompt
                    "changes": "Explain the biggest changes vs last period - what's driving the variance?",
                    "headcount": "What changed in headcount that could explain the expense changes?",
                    "unreconciled": "Explain any unreconciled totals or unclassified amounts",
                    "gl": "Are there any GL mapping gaps or classification issues to address?"
                };
                const prompt = quickPrompts[promptId];
                await generateAdaInsights(prompt);
            });
        });
        
        // Sign-off navigation buttons (Prev/Next at bottom of step)
        document.getElementById("expense-signoff-toggle-prev")?.addEventListener("click", () => {
            moveFocus(-1); // Go to Step 1
        });
        document.getElementById("expense-signoff-toggle-next")?.addEventListener("click", () => {
            moveFocus(1); // Go to Step 3
        });
    }

    const fields = getStepNoteFields(stepId);
    console.log(`[Step ${stepId}] Binding interactions, fields:`, fields);
    if (fields) {
        // Handle Step 1's different ID pattern
        const notesInputId = stepId === 1 ? "step-notes-1" : "step-notes-input";
        const notesInput = document.getElementById(notesInputId);
        console.log(`[Step ${stepId}] Notes input found:`, !!notesInput, `(id: ${notesInputId})`);
        // Get notes save button - Step 2 uses "step-notes-save-2", Step 3 uses "step-notes-save-3"
        const notesSaveBtn = stepId === 1
            ? document.getElementById("step-notes-save-1")
            : stepId === 2
                ? document.getElementById("step-notes-save-2")
                : stepId === 3
                    ? document.getElementById("step-notes-save-3")
                    : document.getElementById("step-notes-save");
        notesInput?.addEventListener("input", (event) => {
            scheduleConfigWrite(fields.note, event.target.value, { debounceMs: 400 });
            // Step 2 is deprecated - no special handling needed
        });
        notesSaveBtn?.addEventListener("click", () => {
            if (!notesInput) return;
            scheduleConfigWrite(fields.note, notesInput.value);
            updateSaveButtonState(notesSaveBtn, true);
        });
        // Handle Step 1's different ID pattern for reviewer
        const reviewerInputId = stepId === 1 ? "step-reviewer-1" : "step-reviewer-name";
        const reviewerInput = document.getElementById(reviewerInputId);
        reviewerInput?.addEventListener("change", (event) => {
            const value = event.target.value.trim();
            scheduleConfigWrite(fields.reviewer, value);
            // Get signoff input for auto-filling date when reviewer enters name
            const signoffInput = stepId === 1
                ? document.getElementById("step-signoff-1")
                : stepId === 2
                    ? document.getElementById("step-signoff-2")
                    : stepId === 3
                        ? document.getElementById("step-signoff-3")
                        : document.getElementById(`step-signoff-${stepId}`);
            if (value && signoffInput && !signoffInput.value) {
                const today = todayIso();
                signoffInput.value = today;
                scheduleConfigWrite(fields.signOff, today);
            }
        });
        // Signoff input ID - use consistent step-signoff-N pattern
        const signoffInputId = stepId === 1
            ? "step-signoff-1"
            : stepId === 2
                ? "step-signoff-2"
                : stepId === 3
                    ? "step-signoff-3"
                    : `step-signoff-${stepId}`;
        console.log(`[Step ${stepId}] Signoff input ID: ${signoffInputId}, found:`, !!document.getElementById(signoffInputId));
        document.getElementById(signoffInputId)?.addEventListener("change", (event) => {
            scheduleConfigWrite(fields.signOff, event.target.value || "");
        });
        // Handle Step 1's different lock button ID
        const lockButtonId = stepId === 1 ? "step-notes-lock-1" : "step-notes-permanent";
        const lockButton = document.getElementById(lockButtonId);
        if (lockButton) {
            lockButton.addEventListener("click", () => {
                const nextState = !lockButton.classList.contains("is-locked");
                updateLockButtonVisual(lockButton, nextState);
                setNotePermanent(fields.note, nextState);
                // Step 2 is deprecated
            });
            updateLockButtonVisual(lockButton, isFieldPermanent(fields.note));
        }
        // Get signoff save button - use consistent step-signoff-save-N pattern
        const signoffSaveBtn = stepId === 1
            ? document.getElementById("step-signoff-save-1")
            : stepId === 2
                ? document.getElementById("step-signoff-save-2")
                : stepId === 3
                    ? document.getElementById("step-signoff-save-3")
                    : document.getElementById(`step-signoff-save-${stepId}`);
        signoffSaveBtn?.addEventListener("click", () => {
            const reviewerValue = reviewerInput?.value?.trim() || "";
            const signoffValue = document.getElementById(signoffInputId)?.value || "";
            scheduleConfigWrite(fields.reviewer, reviewerValue);
            scheduleConfigWrite(fields.signOff, signoffValue);
            updateSaveButtonState(signoffSaveBtn, true);
        });
        initSaveTracking();
        const completeField = STEP_COMPLETE_FIELDS[stepId];
        const initialCompleteFlag = completeField ? Boolean(getConfigValue(completeField)) : false;
        const initialSignOff = getConfigValue(fields.signOff);
        // Map stepId to the actual button ID used in render functions
        const toggleButtonId = stepId === 1
            ? "step-signoff-toggle-1"
            : stepId === 2
                ? "expense-signoff-toggle"  // renderExpenseReviewStep uses this
                : stepId === 3
                    ? "step-signoff-toggle-3"  // We'll fix renderJournalStep to use this
                    : `step-signoff-toggle-${stepId}`;
        console.log(`[Step ${stepId}] Toggle button ID: ${toggleButtonId}, found:`, !!document.getElementById(toggleButtonId));
        bindSignoffToggle({
            buttonId: toggleButtonId,
            inputId: signoffInputId,
            fieldName: fields.signOff,
            completeField,
            requireNotesCheck: null, // Notes no longer required for any step
            initialActive: Boolean(initialSignOff || initialCompleteFlag),
            stepId, // Pass stepId for sequential validation
            onComplete: getStepCompleteHandler(stepId)
        });
        bindSignoffNavButtons(`${toggleButtonId}-prev`, `${toggleButtonId}-next`);
    }

    if (stepId === 4) {
        const archiveBtn = document.getElementById("archive-run-btn");
        if (archiveBtn) {
            // Remove any existing listeners by cloning
            const newBtn = archiveBtn.cloneNode(true);
            archiveBtn.parentNode.replaceChild(newBtn, archiveBtn);
            
            newBtn.onclick = async function() {
                showToast("Starting archive process...", "info", 2000);
                try {
                    await handleArchiveRun();
                } catch (error) {
                    showToast("Archive Error: " + error.message, "error", 8000);
                }
            };
        }
    }
}

function focusStep(index) {
    console.log(`[NAV DEBUG] focusStep(${index}) called`);
    if (Number.isNaN(index) || index < 0 || index >= WORKFLOW_STEPS.length) {
        console.log(`[NAV DEBUG] Invalid index, returning`);
        return;
    }
    const step = WORKFLOW_STEPS[index];
    if (!step) {
        console.log(`[NAV DEBUG] No step at index ${index}, returning`);
        return;
    }
    console.log(`[NAV DEBUG] Setting state: focusedIndex=${index}, activeStepId=${step.id}, title="${step.title}"`);
    pendingScrollIndex = index;
    const view = step.id === 0 ? "config" : "step";
    setState({ focusedIndex: index, activeView: view, activeStepId: step.id });
    
    // Update global context for Ada
    window.PRAIRIE_FORGE_CONTEXT.step = step.id;
    window.PRAIRIE_FORGE_CONTEXT.stepName = step.title;
    window.PRAIRIE_FORGE_CONTEXT.companyId = getConfigValue("SS_Company_ID") || null;
    
    // Activate the corresponding Excel sheet for this step
    const sheetName = STEP_SHEET_MAP[step.id];
    if (sheetName) {
        console.log("[NAV] Step→Sheet activation", {
            moduleKey: MODULE_KEY,
            stepIndex: index,
            stepId: step.id,
            targetSheetName: sheetName
        });
        showAndActivateSheet(sheetName)
            .then(() => {
                console.log("[NAV] Step→Sheet activation success", {
                    moduleKey: MODULE_KEY,
                    stepId: step.id,
                    targetSheetName: sheetName
                });
            })
            .catch((err) => {
                console.warn("[NAV] Step→Sheet activation failed", {
                    moduleKey: MODULE_KEY,
                    stepId: step.id,
                    targetSheetName: sheetName,
                    error: err?.message ?? String(err)
                });
            });
    }
}

function moveFocus(delta) {
    console.log(`[NAV DEBUG] moveFocus(${delta}) called`);
    console.log(`[NAV DEBUG] Current state: focusedIndex=${appState.focusedIndex}, activeView=${appState.activeView}, activeStepId=${appState.activeStepId}`);
    console.log(`[NAV DEBUG] WORKFLOW_STEPS count: ${WORKFLOW_STEPS.length}`);
    
    // When on home view, forward should go to step 0 (Config) first
    if (appState.activeView === "home" && delta > 0) {
        console.log(`[NAV DEBUG] On home view, going to step 0`);
        focusStep(0);
        window.scrollTo({ top: 0, behavior: "smooth" });
        return;
    }
    const next = appState.focusedIndex + delta;
    const clamped = Math.max(0, Math.min(WORKFLOW_STEPS.length - 1, next));
    console.log(`[NAV DEBUG] Calculated: next=${next}, clamped=${clamped}`);
    console.log(`[NAV DEBUG] Target step: id=${WORKFLOW_STEPS[clamped]?.id}, title="${WORKFLOW_STEPS[clamped]?.title}"`);
    focusStep(clamped);
    window.scrollTo({ top: 0, behavior: "smooth" });
}

function bindSignoffNavButtons(prevButtonId, nextButtonId) {
    document.getElementById(prevButtonId)?.addEventListener("click", () => moveFocus(-1));
    document.getElementById(nextButtonId)?.addEventListener("click", () => moveFocus(1));
}

function scrollFocusedIntoView() {
    if (appState.activeView !== "home") return;
    if (pendingScrollIndex === null) return;
    const card = document.querySelector(`[data-step-card][data-step-index="${pendingScrollIndex}"]`);
    pendingScrollIndex = null;
    card?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function openConfigView() {
    focusStep(0);
}

async function returnHome() {
    // Activate the module homepage sheet
    const homepageConfig = getHomepageConfig(MODULE_KEY);
    await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
    
    setState({ activeView: "home", activeStepId: null });
}

/**
 * Navigate to the Module Selector (used after archive completes)
 */
async function navigateToModuleSelector() {
    // Reset tab visibility before redirect to avoid Mac Excel tab accumulation
    try {
        await applyModuleTabVisibility("module-selector");
    } catch (e) {
        console.warn("[Payroll] Could not apply module-selector tab visibility before redirect:", e);
    }
    // Navigate to Module Selector page
    // Use relative path from current location (payroll-recorder/ -> module-selector/)
    window.location.href = "../module-selector/index.html";
}

function setState(partial) {
    Object.assign(appState, partial);
    
    // Persist route state for taskpane reload recovery
    const route = buildRouteString(MODULE_KEY, appState.activeView, appState.activeStepId);
    saveRouteState(MODULE_KEY, route, {
        activeView: appState.activeView,
        activeStepId: appState.activeStepId,
        focusedIndex: appState.focusedIndex
    });
    
    renderApp();
}

function getReviewerDefault() {
    // Fallback to legacy field if the new one is missing
    return getConfigValue(CONFIG_REVIEWER_FIELD) || getConfigValue("SS_Default_Reviewer") || "";
}

function updateActionToggleState(button, isActive) {
    if (!button) return;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
}

function markJeSaveState(isSaved) {
    const btn = document.getElementById("je-save-btn");
    if (!btn) return;
    btn.classList.toggle("is-saved", isSaved);
}

/**
 * Clear step completion status when user makes changes
 * Forces user to re-confirm "Done" after modifications
 * @param {number} stepId - The step ID to clear (0-6)
 */
function clearStepCompletion(stepId) {
    const fields = STEP_NOTES_FIELDS[stepId];
    const completeField = STEP_COMPLETE_FIELDS[stepId];
    
    if (!fields || !completeField) return;
    
    // Check if step was previously complete
    const signOffValue = getConfigValue(fields.signOff);
    const completeValue = getConfigValue(completeField);
    const wasComplete = Boolean(signOffValue) || completeValue === "Y" || completeValue === true;
    
    if (!wasComplete) return; // Nothing to clear
    
    console.log(`[Signoff] Clearing completion for step ${stepId} due to field change`);
    
    // Clear sign-off date and completion flag
    scheduleConfigWrite(fields.signOff, "");
    scheduleConfigWrite(completeField, "");
    
    // Update the button UI if visible
    const button = document.querySelector(`[id$="-signoff-toggle"], [id$="-signoff-toggle-${stepId}"]`);
    if (button) {
        button.classList.remove("is-active");
        button.setAttribute("aria-pressed", "false");
    }
    
    // Update the signoff date input if visible
    const signoffInput = document.querySelector(`[id^="config-signoff-"], [id^="step-signoff-"]`);
    if (signoffInput) {
        signoffInput.value = "";
    }
}

/**
 * Get current step completion status for sequential validation
 * @returns {Object} Map of step IDs to boolean completion status
 */
function getStepCompletionStatus() {
    const status = {};
    console.log("[Signoff] Checking step completion status...");
    Object.keys(STEP_NOTES_FIELDS).forEach(stepIdStr => {
        const stepId = parseInt(stepIdStr, 10);
        const fields = STEP_NOTES_FIELDS[stepId];
        if (!fields) {
            status[stepId] = false;
            return;
        }
        // A step is complete if it has a sign-off date OR is explicitly marked complete
        const signOffValue = getConfigValue(fields.signOff);
        const completeField = STEP_COMPLETE_FIELDS[stepId];
        const completeValue = getConfigValue(completeField);
        const isComplete = Boolean(signOffValue) || completeValue === "Y" || completeValue === true;
        status[stepId] = isComplete;
        console.log(`[Signoff] Step ${stepId}: signOff="${signOffValue}", complete="${completeValue}" → ${isComplete ? "COMPLETE" : "pending"}`);
    });
    console.log("[Signoff] Status summary:", status);
    return status;
}

function bindSignoffToggle({
    buttonId,
    inputId,
    fieldName,
    completeField,
    requireNotesCheck,
    onComplete,
    initialActive = false,
    stepId = null // NEW: Step ID for sequential validation
}) {
    const button = document.getElementById(buttonId);
    if (!button) {
        console.warn(`[Signoff] Button not found: ${buttonId}`);
        return;
    }
    const input = inputId ? document.getElementById(inputId) : null;
    const initial = initialActive || Boolean(input?.value);
    updateActionToggleState(button, initial);
    console.log(`[Signoff] Bound ${buttonId}, initial active: ${initial}, stepId: ${stepId}`);
    
    // Handle Done button click
    button.addEventListener("click", () => {
        console.log(`[Signoff] Done button clicked: ${buttonId}, stepId: ${stepId}`);
        
        // Check sequential completion if stepId is provided
        if (stepId !== null && stepId > 0) {
            const completionStatus = getStepCompletionStatus();
            const { canComplete, message } = canCompleteStep(stepId, completionStatus);
            
            // Only block if trying to COMPLETE (not uncomplete)
            const isCurrentlyActive = button.classList.contains("is-active");
            console.log(`[Signoff] canComplete: ${canComplete}, isCurrentlyActive: ${isCurrentlyActive}`);
            if (!isCurrentlyActive && !canComplete) {
                console.log(`[Signoff] BLOCKED: ${message}`);
                showBlockedToast(message);
                return;
            }
        }
        
        // Notes check removed - only sequential validation required
        const nextActive = !button.classList.contains("is-active");
        console.log(`[Signoff] ${buttonId} clicked, toggling to: ${nextActive}`);
        updateActionToggleState(button, nextActive);
        if (input) {
            input.value = nextActive ? todayIso() : "";
        }
        if (fieldName) {
            const dateValue = nextActive ? todayIso() : "";
            console.log(`[Signoff] Writing ${fieldName} = "${dateValue}"`);
            scheduleConfigWrite(fieldName, dateValue);
        }
        if (completeField) {
            const completeValue = nextActive ? "Y" : "";
            console.log(`[Signoff] Writing ${completeField} = "${completeValue}"`);
            scheduleConfigWrite(completeField, completeValue);
        }
        if (nextActive) {
            window.scrollTo({ top: 0, behavior: "smooth" });
        }
        if (nextActive && typeof onComplete === "function") {
            onComplete();
        }
    });
    
    // Handle manual date input change - sync the button state
    if (input) {
        input.addEventListener("change", () => {
            const hasDate = Boolean(input.value);
            const isCurrentlyActive = button.classList.contains("is-active");
            if (hasDate !== isCurrentlyActive) {
                console.log(`[Signoff] Date input changed, syncing button to: ${hasDate}`);
                updateActionToggleState(button, hasDate);
                if (fieldName) {
                    scheduleConfigWrite(fieldName, input.value || "");
                }
                if (completeField) {
                    scheduleConfigWrite(completeField, hasDate ? "Y" : "");
                }
            }
        });
    }
}

async function openConfigurationSheet() {
    if (!hasExcelRuntime()) {
        showToast("Open this module inside Excel to edit configuration settings.", "info");
        return;
    }
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MODULE_CONFIG_SHEET);
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
        setState({ statusText: `${MODULE_CONFIG_SHEET} opened.` });
    } catch (error) {
        console.error("Unable to open configuration sheet", error);
        showToast(`Unable to open ${MODULE_CONFIG_SHEET}. Confirm the sheet exists in this workbook.`, "error");
    }
}

async function openDataSheet() {
    if (!hasExcelRuntime()) {
        showToast("Open this module inside Excel to access the data sheet.", "info");
        return;
    }
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA_CLEAN);
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error("Unable to open PR_Data_Clean sheet", error);
        showToast(`Unable to open ${SHEET_NAMES.DATA_CLEAN}. Confirm the sheet exists in this workbook.`, "error");
    }
}

async function clearPrDataSheet() {
    if (!hasExcelRuntime()) {
        showToast("Open this module inside Excel to clear data.", "info");
        return;
    }
    const confirmed = await showConfirm(
        "All data in PR_Data_Clean will be permanently removed.\n\nThis action cannot be undone.",
        {
            title: "Clear Payroll Data",
            icon: TRASH_ICON_SVG,
            confirmText: "Clear Data",
            cancelText: "Keep Data",
            destructive: true
        }
    );
    if (!confirmed) return;
    
    try {
        await Excel.run(async (context) => {
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            cleanSheet.load("isNullObject");
            await context.sync();
            
            if (!cleanSheet.isNullObject) {
                const usedRange = cleanSheet.getUsedRangeOrNullObject();
                usedRange.load("isNullObject");
                await context.sync();
                
                if (!usedRange.isNullObject) {
                    // Clear all except header row (row 1)
                    const cleanDataRange = cleanSheet.getRange("A2:Z10000");
                    cleanDataRange.clear(Excel.ClearApplyTo.contents);
                    await context.sync();
                }
                
                cleanSheet.activate();
                cleanSheet.getRange("A1").select();
                await context.sync();
            }
        });
        
        // Reset workflow states so they'll rebuild when needed
        updateValidationState({
            prDataTotal: null,
            cleanTotal: null,
            reconDifference: null
        }, { rerender: false });
        
        showToast("PR_Data_Clean cleared successfully.", "success");
    } catch (error) {
        console.error("Unable to clear PR_Data_Clean sheet", error);
        showToast("Unable to clear PR_Data_Clean. Please try again.", "error");
    }
}

async function getConfigTable(context) {
    if (!CONFIG_TABLE_CANDIDATES.length) return null;
    if (resolvedConfigTableName) {
        const existing = context.workbook.tables.getItemOrNullObject(resolvedConfigTableName);
        existing.load("name");
        await context.sync();
        if (!existing.isNullObject) {
            return existing;
        }
        resolvedConfigTableName = null;
    }
    const tables = context.workbook.tables;
    tables.load("items/name");
    await context.sync();

    const foundTableNames = tables.items?.map((t) => t.name) || [];
    
    // Debug logging - visible in browser console (F12)
    console.log("[Payroll] Looking for config table:", CONFIG_TABLE_CANDIDATES);
    console.log("[Payroll] Found tables in workbook:", foundTableNames);

    const match = tables.items?.find((table) => CONFIG_TABLE_CANDIDATES.includes(table.name));
    if (!match) {
        console.warn("[Payroll] CONFIG TABLE NOT FOUND!");
        console.warn("[Payroll] Expected table named: SS_PF_Config");
        console.warn("[Payroll] Available tables:", foundTableNames);
        console.warn("[Payroll] To fix: Select your data in SS_PF_Config sheet -> Insert -> Table -> Name it 'SS_PF_Config'");
        return null;
    }
    console.log("[Payroll] Config table found:", match.name);
    resolvedConfigTableName = match.name;
    return context.workbook.tables.getItem(match.name);
}

async function loadConfigurationValues() {
    if (!hasExcelRuntime()) {
        configState.loaded = true;
        return;
    }
    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            if (!table) {
                console.warn("Payroll Recorder: SS_PF_Config table is missing.");
                configState.loaded = true;
                return;
            }
            const body = table.getDataBodyRange();
            body.load("values");
            await context.sync();
            const rows = body.values || [];
            const map = {};
            const permanents = {};
            rows.forEach((row) => {
                const field = normalizeFieldName(row[CONFIG_COLUMNS.FIELD]);
                if (!field) return;
                map[field] = row[CONFIG_COLUMNS.VALUE] ?? "";
                permanents[field] = row[CONFIG_COLUMNS.PERMANENT] ?? "";
            });
            configState.values = map;
            configState.permanents = permanents;
            // Check both new and legacy field names for overrides
            configState.overrides.accountingPeriod = Boolean(map.PR_Accounting_Period || map.Accounting_Period);
            configState.overrides.jeId = Boolean(map.PR_Journal_Entry_ID || map.Journal_Entry_ID);
            configState.loaded = true;
        });
    } catch (error) {
        console.warn("Payroll Recorder: unable to load PF_Config table.", error);
        configState.loaded = true;
    }
}

function getConfigValue(field) {
    return configState.values[field] ?? "";
}

function resolvePayrollDateFieldName() {
    const keys = Object.keys(configState.values || {});
    const match = PAYROLL_DATE_ALIASES.find((alias) => keys.includes(alias));
    return match || PAYROLL_DATE_ALIASES[0];
}

function getPayrollDateValue() {
    return getConfigValue(resolvePayrollDateFieldName());
}

function getPayrollProviderLink() {
    // Check module-specific field (PR_Payroll_Provider), then legacy fallback
    return (
        getConfigValue(PAYROLL_PROVIDER_FIELD) || 
        getConfigValue("Payroll_Provider_Link") || 
        ""
    ).trim();
}

function isFieldPermanent(field) {
    return parseBooleanFlag(configState.permanents[field]);
}

function isStepCompleteFromConfig(stepId) {
    const field = STEP_COMPLETE_FIELDS[stepId];
    if (!field) return false;
    return parseBooleanFlag(getConfigValue(field));
}

function setNotePermanent(field, isPermanent) {
    const normalizedField = normalizeFieldName(field);
    if (!normalizedField) return;
    configState.permanents[normalizedField] = isPermanent ? "Y" : "N";
    void writeConfigPermanent(normalizedField, isPermanent ? "Y" : "N");
}

function parseBooleanFlag(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "true" || normalized === "y" || normalized === "yes" || normalized === "1";
}

function normalizeFieldName(value) {
    return String(value ?? "").trim();
}

function isNoiseName(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    if (!normalized) return true;
    
    // Exact match terms (headers, placeholders)
    const exactMatchTerms = [
        "employee", "employee name", "name", "full name",
        "header", "column", "n/a", "none", "blank", "null", "undefined"
    ];
    if (exactMatchTerms.some(term => normalized === term || normalized === term.replace(/\s+/g, ""))) {
        return true;
    }
    
    // Contains-based filtering for totals and aggregate rows
    // These patterns indicate the row is a summary/total, not an individual employee
    const containsTerms = [
        "total",           // "~Report Totals", "Grand Total", "Subtotal"
        "subtotal",
        "summary",
        "grand total",
        "totals for department",  // "~Totals for DEPARTMENT : 10 - COGS Support"
        "department total",
        "dept total",
        "report total"
    ];
    if (containsTerms.some(term => normalized.includes(term))) {
        return true;
    }
    
    // Starts-with patterns (payroll systems often use ~ or * for totals)
    const startsWithPatterns = [
        "~",              // Common payroll export prefix for totals
        "*",              // Another common total indicator
        "---",            // Separator rows
        "==="
    ];
    if (startsWithPatterns.some(prefix => normalized.startsWith(prefix))) {
        return true;
    }
    
    return false;
}

function formatDateInput(value) {
    if (!value) return "";
    const parts = parseDateInput(value);
    if (!parts) return "";
    return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function deriveAccountingPeriod(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    // Validate year is reasonable (1900-2100)
    if (parts.year < 1900 || parts.year > 2100) {
        console.warn("deriveAccountingPeriod - Invalid year:", parts.year, "from input:", payrollDate);
        return "";
    }
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    // Format: "Dec 2025" - use full 4-digit year to prevent Excel date interpretation
    return `${monthNames[parts.month - 1]} ${parts.year}`;
}

function deriveJeId(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    // Validate year is reasonable (1900-2100)
    if (parts.year < 1900 || parts.year > 2100) {
        console.warn("deriveJeId - Invalid year:", parts.year, "from input:", payrollDate);
        return "";
    }
    return `PR-AUTO-${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function todayIso() {
    return formatDateFromDate(new Date());
}

function scheduleConfigWrite(fieldName, value, options = {}) {
    const normalizedField = normalizeFieldName(fieldName);
    configState.values[normalizedField] = value ?? "";
    const delay = options.debounceMs ?? 0;
    if (!delay) {
        const existing = pendingWrites.get(normalizedField);
        if (existing) clearTimeout(existing);
        pendingWrites.delete(normalizedField);
        void writeConfigValue(normalizedField, value ?? "");
        return;
    }
    if (pendingWrites.has(normalizedField)) {
        clearTimeout(pendingWrites.get(normalizedField));
    }
    const timer = setTimeout(() => {
        pendingWrites.delete(normalizedField);
        void writeConfigValue(normalizedField, value ?? "");
    }, delay);
    pendingWrites.set(normalizedField, timer);
}

// Fields that should be forced to Text format to prevent Excel auto-conversion
const TEXT_FORMAT_FIELDS = [
    "PR_Accounting_Period",
    "PTO_Accounting_Period",
    "Accounting_Period"
];

async function writeConfigValue(fieldName, value) {
    const normalizedField = normalizeFieldName(fieldName);
    configState.values[normalizedField] = value ?? "";
    console.log(`[Payroll] Writing config: ${normalizedField} = "${value}"`);
    if (!hasExcelRuntime()) {
        console.warn("[Payroll] Excel runtime not available - cannot write");
        return;
    }
    
    // Check if this field needs text formatting
    const forceTextFormat = TEXT_FORMAT_FIELDS.some(f => 
        normalizedField === f || normalizedField.toLowerCase() === f.toLowerCase()
    );
    
    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            if (!table) {
                console.error("[Payroll] ❌ Cannot write - config table not found");
                return;
            }
            const body = table.getDataBodyRange();
            const headerRange = table.getHeaderRowRange();
            body.load("values");
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0] || [];
            const rows = body.values || [];
            const columnCount = headers.length;
            console.log(`[Payroll] Table has ${rows.length} rows, ${columnCount} columns`);

            // Find ALL matching rows (to handle duplicates)
            const matchingIndices = [];
            rows.forEach((row, idx) => {
                if (normalizeFieldName(row[CONFIG_COLUMNS.FIELD]) === normalizedField) {
                    matchingIndices.push(idx);
                }
            });

            if (matchingIndices.length === 0) {
                // No existing row - add new one
                configState.permanents[normalizedField] = configState.permanents[normalizedField] ?? DEFAULT_CONFIG_PERMANENT;
                const newRow = new Array(columnCount).fill("");
                if (CONFIG_COLUMNS.TYPE >= 0 && CONFIG_COLUMNS.TYPE < columnCount) newRow[CONFIG_COLUMNS.TYPE] = DEFAULT_CONFIG_TYPE;
                if (CONFIG_COLUMNS.FIELD >= 0 && CONFIG_COLUMNS.FIELD < columnCount) newRow[CONFIG_COLUMNS.FIELD] = normalizedField;
                if (CONFIG_COLUMNS.VALUE >= 0 && CONFIG_COLUMNS.VALUE < columnCount) newRow[CONFIG_COLUMNS.VALUE] = value ?? "";
                if (CONFIG_COLUMNS.PERMANENT >= 0 && CONFIG_COLUMNS.PERMANENT < columnCount) newRow[CONFIG_COLUMNS.PERMANENT] = DEFAULT_CONFIG_PERMANENT;
                console.log(`[Payroll] Adding NEW row:`, newRow);
                table.rows.add(null, [newRow]);
                await context.sync();
                
                // Force text format for specific fields to prevent Excel date conversion
                if (forceTextFormat) {
                    // Get the newly added row (last row in table)
                    const tableRows = table.rows;
                    tableRows.load("count");
                    await context.sync();
                    const lastRowIdx = tableRows.count - 1;
                    const newRowRange = table.rows.getItemAt(lastRowIdx).getRange();
                    const valueCell = newRowRange.getCell(0, CONFIG_COLUMNS.VALUE);
                    valueCell.numberFormat = [["@"]]; // Text format
                    valueCell.values = [[value ?? ""]]; // Re-write to apply format
                    await context.sync();
                    console.log(`[Payroll] Applied text format to ${normalizedField}`);
                }
                
                console.log(`[Payroll] New row added for ${normalizedField}`);
            } else {
                // Update the first matching row
                const targetIndex = matchingIndices[0];
                console.log(`[Payroll] Updating existing row ${targetIndex} for ${normalizedField}`);
                const targetCell = body.getCell(targetIndex, CONFIG_COLUMNS.VALUE);
                
                // Force text format for specific fields
                if (forceTextFormat) {
                    targetCell.numberFormat = [["@"]]; // Text format
                }
                targetCell.values = [[value ?? ""]];
                await context.sync();
                console.log(`[Payroll] Updated ${normalizedField}`);

                // Delete duplicate rows (in reverse order to maintain indices)
                if (matchingIndices.length > 1) {
                    console.log(`[Payroll] Found ${matchingIndices.length - 1} duplicate rows for ${normalizedField}, removing...`);
                    const duplicateIndices = matchingIndices.slice(1).reverse();
                    for (const dupIdx of duplicateIndices) {
                        try {
                            table.rows.getItemAt(dupIdx).delete();
                        } catch (e) {
                            console.warn(`[Payroll] Could not delete duplicate row ${dupIdx}:`, e.message);
                        }
                    }
                    await context.sync();
                    console.log(`[Payroll] Removed duplicate rows for ${normalizedField}`);
                }
            }
        });
    } catch (error) {
        console.error(`[Payroll] ❌ Write failed for ${fieldName}:`, error);
    }
}

async function writeConfigPermanent(fieldName, marker) {
    const normalizedField = normalizeFieldName(fieldName);
    if (!normalizedField) return;
    if (!hasExcelRuntime()) return;
    // Store in local state
    configState.permanents[normalizedField] = marker;
    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            if (!table) {
                console.warn(`Payroll Recorder: unable to locate config table when toggling ${fieldName} permanent flag.`);
                return;
            }
            const body = table.getDataBodyRange();
            body.load("values");
            await context.sync();
            const rows = body.values || [];
            const targetIndex = rows.findIndex(
                (row) => normalizeFieldName(row[CONFIG_COLUMNS.FIELD]) === normalizedField
            );
            if (targetIndex === -1) return;
            body.getCell(targetIndex, CONFIG_COLUMNS.PERMANENT).values = [[marker]];
            await context.sync();
        });
    } catch (error) {
        console.warn(`Payroll Recorder: unable to update permanent flag for ${fieldName}`, error);
    }
}

function parseDateInput(value) {
    if (!value) return null;
    
    // Handle string value
    const strValue = String(value).trim();
    
    // Try YYYY-MM-DD format first
    const isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(strValue);
    if (isoMatch) {
        const year = Number(isoMatch[1]);
        const month = Number(isoMatch[2]);
        const day = Number(isoMatch[3]);
        if (year && month && day) return { year, month, day };
    }
    
    // Try MM/DD/YYYY format
    const usMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(strValue);
    if (usMatch) {
        const month = Number(usMatch[1]);
        const day = Number(usMatch[2]);
        const year = Number(usMatch[3]);
        if (year && month && day) return { year, month, day };
    }
    
    // Handle Excel serial date number
    // Use UTC to avoid timezone offset issues - Excel dates should be treated as UTC
    const numValue = Number(value);
    if (Number.isFinite(numValue) && numValue > 40000 && numValue < 60000) {
        // Excel serial date: days since Jan 1, 1900 (with 1900 leap year bug)
        // Convert to UTC timestamp to avoid local timezone shifting the date
        const utcDays = Math.floor(numValue - 25569); // Days since Unix epoch (Jan 1, 1970)
        const utcMs = utcDays * 86400 * 1000;
        const jsDate = new Date(utcMs);
        if (!isNaN(jsDate.getTime())) {
            // Use UTC methods to extract date components to prevent timezone shift
            const isoDate = `${jsDate.getUTCFullYear()}-${String(jsDate.getUTCMonth() + 1).padStart(2, "0")}-${String(jsDate.getUTCDate()).padStart(2, "0")}`;
            console.log("DEBUG parseDateInput - Converted Excel serial", numValue, "to", isoDate);
            return {
                year: jsDate.getUTCFullYear(),
                month: jsDate.getUTCMonth() + 1,
                day: jsDate.getUTCDate()
            };
        }
    }
    
    // Try parsing as Date string
    const dateObj = new Date(strValue);
    if (!isNaN(dateObj.getTime())) {
        return {
            year: dateObj.getFullYear(),
            month: dateObj.getMonth() + 1,
            day: dateObj.getDate()
        };
    }
    
    console.warn("DEBUG parseDateInput - Could not parse date value:", value);
    return null;
}

function formatDateFromDate(date) {
    // Use UTC methods if this date was derived from Excel serial number
    // to prevent timezone shift causing off-by-one day errors
    if (date._isUTC) {
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, "0");
        const day = String(date.getUTCDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    // For regular dates (like "today"), use local time
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

/**
 * Format date for roster display (MM/DD/YYYY format)
 * Handles YYYY-MM-DD strings, Date objects, and numeric serial numbers
 */
function formatDateForRoster(value) {
    if (!value) return "";
    
    // If it's already a date string in YYYY-MM-DD format, convert to MM/DD/YYYY
    if (typeof value === "string" && value.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const [yyyy, mm, dd] = value.split("-");
        return `${mm}/${dd}/${yyyy}`;
    }
    
    // If it's a Date object
    if (value instanceof Date && !isNaN(value.getTime())) {
        const mm = String(value.getMonth() + 1).padStart(2, "0");
        const dd = String(value.getDate()).padStart(2, "0");
        const yyyy = value.getFullYear();
        return `${mm}/${dd}/${yyyy}`;
    }
    
    // If it's a number (Excel serial), convert to date first
    if (typeof value === "number" && Number.isFinite(value)) {
        const ms = Math.round((value - 25569) * 86400 * 1000);
        const d = new Date(ms);
        if (Number.isFinite(d.getTime())) {
            const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
            const dd = String(d.getUTCDate()).padStart(2, "0");
            const yyyy = d.getUTCFullYear();
            return `${mm}/${dd}/${yyyy}`;
        }
    }
    
    // Return as-is if we can't parse
    return String(value).trim();
}

/**
 * Normalize a date value for lookup comparison
 * Handles Excel serial dates, Date objects, and various string formats
 * Returns YYYY-MM-DD string or null if unparseable
 */
function normalizeDateForLookup(value) {
    if (!value) return null;
    
    // If it's already a YYYY-MM-DD string, return it
    if (typeof value === "string") {
        const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
            return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
        }
    }
    
    // Try parsing with our date parser
    const parts = parseDateInput(value);
    if (parts) {
        return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
    }
    
    return null;
}

/**
 * Create a context provider for Prairie Forge CoPilot
 * Reads current payroll data to provide intelligent, contextual responses
 */
function createPayrollContextProvider() {
    return async () => {
        if (!hasExcelRuntime()) return null;
        
        try {
            return await Excel.run(async (context) => {
                const result = {
                    timestamp: new Date().toISOString(),
                    period: null,
                    summary: {},
                    departments: [],
                    recentPeriods: [],
                    dataQuality: {}
                };
                
                // Read config for period info
                const configTable = await getConfigTable(context);
                if (configTable) {
                    const configBody = configTable.getDataBodyRange();
                    configBody.load("values");
                    await context.sync();
                    
                    const configRows = configBody.values || [];
                    for (const row of configRows) {
                        const fieldName = String(row[CONFIG_COLUMNS.FIELD] || "").trim();
                        const fieldValue = row[CONFIG_COLUMNS.VALUE];
                        
                        if (fieldName.toLowerCase().includes("accounting") && fieldName.toLowerCase().includes("period")) {
                            result.period = String(fieldValue || "").trim();
                        }
                    }
                }
                
                // Read PR_Data_Clean for summary stats
                const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
                cleanSheet.load("isNullObject");
                await context.sync();
                
                if (!cleanSheet.isNullObject) {
                    const range = cleanSheet.getUsedRangeOrNullObject();
                    range.load("values");
                    await context.sync();
                    
                    if (!range.isNullObject && range.values?.length > 1) {
                        const headers = range.values[0].map(h => normalizeHeader(h));
                        const data = range.values.slice(1);
                        
                        // Find relevant columns
                        const amountIdx = headers.findIndex(h => h.includes("amount"));
                        const deptIdx = pickDepartmentIndex(headers);
                        const employeeIdx = headers.findIndex(h => h.includes("employee"));
                        
                        // Calculate totals
                        let totalAmount = 0;
                        const employeeSet = new Set();
                        const deptTotals = new Map();
                        
                        for (const row of data) {
                            const amount = Number(row[amountIdx]) || 0;
                            totalAmount += amount;
                            
                            if (employeeIdx >= 0) {
                                const emp = String(row[employeeIdx] || "").trim();
                                if (emp) employeeSet.add(emp);
                            }
                            
                            if (deptIdx >= 0) {
                                const dept = String(row[deptIdx] || "").trim();
                                if (dept) {
                                    deptTotals.set(dept, (deptTotals.get(dept) || 0) + amount);
                                }
                            }
                        }
                        
                        result.summary = {
                            total: totalAmount,
                            employeeCount: employeeSet.size,
                            avgPerEmployee: employeeSet.size ? totalAmount / employeeSet.size : 0,
                            rowCount: data.length
                        };
                        
                        // Department breakdown
                        result.departments = Array.from(deptTotals.entries())
                            .map(([name, total]) => ({
                                name,
                                total,
                                percentOfTotal: totalAmount ? (total / totalAmount) : 0
                            }))
                            .sort((a, b) => b.total - a.total);
                        
                        result.dataQuality.dataCleanReady = true;
                        result.dataQuality.rowCount = data.length;
                    }
                }
                
                // Read PR_Archive_Summary for trend data
                const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
                archiveSheet.load("isNullObject");
                await context.sync();
                
                if (!archiveSheet.isNullObject) {
                    const archiveRange = archiveSheet.getUsedRangeOrNullObject();
                    archiveRange.load("values");
                    await context.sync();
                    
                    if (!archiveRange.isNullObject && archiveRange.values?.length > 1) {
                        const headers = archiveRange.values[0].map(h => normalizeHeader(h));
                        const periodIdx = headers.findIndex(h => h.includes("period"));
                        const totalIdx = headers.findIndex(h => h.includes("total"));
                        
                        result.recentPeriods = archiveRange.values.slice(1, 6).map(row => ({
                            period: row[periodIdx] || "",
                            total: Number(row[totalIdx]) || 0
                        }));
                        
                        result.dataQuality.archiveAvailable = true;
                        result.dataQuality.periodsAvailable = result.recentPeriods.length;
                    }
                }
                
                // Check JE Draft status
                const jeSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.JE_DRAFT);
                jeSheet.load("isNullObject");
                await context.sync();
                
                if (!jeSheet.isNullObject) {
                    const jeRange = jeSheet.getUsedRangeOrNullObject();
                    jeRange.load("values");
                    await context.sync();
                    
                    if (!jeRange.isNullObject && jeRange.values?.length > 1) {
                        const headers = jeRange.values[0].map(h => normalizeHeader(h));
                        const debitIdx = headers.findIndex(h => h.includes("debit"));
                        const creditIdx = headers.findIndex(h => h.includes("credit"));
                        
                        let totalDebit = 0;
                        let totalCredit = 0;
                        
                        for (const row of jeRange.values.slice(1)) {
                            totalDebit += Number(row[debitIdx]) || 0;
                            totalCredit += Number(row[creditIdx]) || 0;
                        }
                        
                        result.journalEntry = {
                            totalDebit,
                            totalCredit,
                            difference: Math.abs(totalDebit - totalCredit),
                            isBalanced: Math.abs(totalDebit - totalCredit) < 0.01,
                            lineCount: jeRange.values.length - 1
                        };
                        
                        result.dataQuality.jeDraftReady = true;
                    }
                }
                
                console.log("CoPilot context gathered:", result);
                return result;
            });
        } catch (error) {
            console.warn("CoPilot context provider error:", error);
            return null;
        }
    };
}

// =============================================================================
// ADA INSIGHTS - EXPENSE REVIEW CONTEXT PACK BUILDER
// =============================================================================

/**
 * Build the comprehensive context pack for Ada Insights on Expense Review
 * This is the ONLY input Ada needs - all analysis must be grounded in this data
 * 
 * @returns {Promise<object>} - ExpenseReviewContextPack
 */
async function buildExpenseReviewContextPack() {
    console.log("[AdaInsights] Building context pack...");
    
    const contextPack = {
        // A) Identity + Period
        identity: {
            module: "payroll-recorder",
            timestamp: new Date().toISOString(),
            basis_mode: "CASH_OUTFLOW" // PEO default - payroll report equals cash outflow
        },
        period: {
            current_key: null,
            prior_key: null,
            periods_available: 0
        },
        
        // B) Totals + Reconciliation
        totals: {
            expense_review_total_current: 0,
            expense_review_total_prior: 0,
            pr_data_clean_total_numeric: 0,
            calculated_total_current_logic: 0,
            bank_statement_amount: null,
            bank_delta: null
        },
        
        // C) Delta breakdown (from diagnostic)
        delta_breakdown: {
            total_unclassified: 0,
            total_excluded_by_side_ee: 0,
            total_excluded_by_side_na: 0,
            total_excluded_by_include_false: 0,
            total_excluded_by_summary: 0,
            top_unclassified_by_dollars: [],
            top_excluded_by_dollars: []
        },
        
        // D) Drivers (bucket aggregates)
        drivers: {
            bucket_totals_current: { FIXED: 0, VARIABLE: 0, BURDEN: 0 },
            bucket_totals_prior: { FIXED: 0, VARIABLE: 0, BURDEN: 0 },
            bucket_deltas: { FIXED: 0, VARIABLE: 0, BURDEN: 0 },
            top_measure_deltas: [],
            department_deltas: []
        },
        
        // E) Metadata coverage
        metadata: {
            measures_in_sheet_count: 0,
            measures_with_dictionary_metadata_count: 0,
            measures_missing_dictionary_metadata: [],
            measures_with_blank_side: [],
            new_measures_this_period: []
        },
        
        // F) Roster context (computed separately)
        roster_context: null,
        
        // G) Data availability flags
        availability: {
            has_pr_data_clean: false,
            has_archive_totals: false,
            has_archive_summary: false,
            has_employee_roster: false,
            has_prior_period: false,
            error_messages: []
        }
    };
    
    if (!hasExcelRuntime()) {
        contextPack.availability.error_messages.push("Excel runtime is unavailable.");
        return contextPack;
    }
    
    try {
        await Excel.run(async (context) => {
            // Load sheets
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
            const archiveTotalsSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_TOTALS);
            
            cleanSheet.load("isNullObject");
            archiveSheet.load("isNullObject");
            archiveTotalsSheet.load("isNullObject");
            await context.sync();

            let archiveTotalsHasData = false;
            let usedArchiveTotalsForPrior = false;
            
            // --- PR_Data_Clean Analysis ---
            if (!cleanSheet.isNullObject) {
                contextPack.availability.has_pr_data_clean = true;
                
                const cleanRange = cleanSheet.getUsedRangeOrNullObject();
                cleanRange.load("values");
                await context.sync();
                
                if (!cleanRange.isNullObject && cleanRange.values && cleanRange.values.length > 1) {
                    const headers = cleanRange.values[0].map(h => String(h || "").trim());
                    const headersLower = headers.map(h => h.toLowerCase());
                    const dataRows = cleanRange.values.slice(1);
                    
                    // Find Pay_Date column for period key
                    const payDateIdx = headersLower.findIndex(h => 
                        h === "pay_date" || h === "payroll_date" || h.includes("pay_period")
                    );
                    
                    if (payDateIdx >= 0 && dataRows.length > 0) {
                        const payDateRaw = dataRows[0][payDateIdx];
                        contextPack.period.current_key = normalizePayDate(payDateRaw);
                    }
                    
                    // Use taxonomy for classification
                    const taxonomy = expenseTaxonomyCache;
                    const dimensionIndices = new Set();
                    
                    // Identify dimension columns
                    headers.forEach((header, idx) => {
                        const headerLower = header.toLowerCase();
                        if (taxonomy.loaded && taxonomy.dimensions && taxonomy.dimensions.has(headerLower)) {
                            dimensionIndices.add(idx);
                        }
                    });
                    
                    // Analyze each column
                    let totalNumeric = 0;
                    let measuresInSheet = 0;
                    let measuresWithMeta = 0;
                    const measureDeltas = [];
                    const excludedEE = [];
                    const excludedNA = [];
                    const excludedIncludeFalse = [];
                    const excludedSummary = [];
                    const unclassifiedCols = [];
                    const blankSideCols = [];
                    
                    headers.forEach((header, idx) => {
                        if (dimensionIndices.has(idx)) return;
                        
                        const headerLower = header.toLowerCase();
                        let colTotal = 0;
                        let hasNumeric = false;
                        
                        for (const row of dataRows) {
                            const val = Number(row[idx]);
                            if (!isNaN(val) && val !== 0) {
                                hasNumeric = true;
                                colTotal += val;
                            }
                        }
                        
                        if (!hasNumeric) return;
                        
                        totalNumeric += colTotal;
                        measuresInSheet++;
                        
                        const meta = (taxonomy.loaded && taxonomy.measures) ? taxonomy.measures[headerLower] : null;
                        
                        if (meta) {
                            measuresWithMeta++;
                            
                            // Track blank side
                            const side = String(meta.side || "").toLowerCase().trim();
                            if (!side) {
                                blankSideCols.push({ header, total: colTotal });
                            }
                            
                            // Check exclusion reason
                            if (EXPENSE_REVIEW_SUMMARY_EXCLUSIONS.has(headerLower)) {
                                excludedSummary.push({ header, total: colTotal, reason: "summary_exclusion" });
                            } else if (side === 'ee') {
                                excludedEE.push({ header, total: colTotal });
                            } else if (side === 'na') {
                                excludedNA.push({ header, total: colTotal });
                            } else if (meta.include === false) {
                                excludedIncludeFalse.push({ header, total: colTotal });
                            } else {
                                // Included measure - track for deltas
                                const bucket = (meta.bucket || "OTHER").toUpperCase();
                                measureDeltas.push({
                                    pf_column_name: header,
                                    bucket,
                                    side: meta.side || null,
                                    current_amount: colTotal * (meta.sign ?? 1),
                                    prior_amount: 0, // Will be filled from archive
                                    delta_amount: 0
                                });
                            }
                        } else {
                            // Check summary exclusions even without meta
                            if (EXPENSE_REVIEW_SUMMARY_EXCLUSIONS.has(headerLower)) {
                                excludedSummary.push({ header, total: colTotal, reason: "summary_exclusion" });
                            } else {
                                unclassifiedCols.push({ header, total: colTotal });
                            }
                        }
                    });
                    
                    contextPack.totals.pr_data_clean_total_numeric = totalNumeric;
                    contextPack.metadata.measures_in_sheet_count = measuresInSheet;
                    contextPack.metadata.measures_with_dictionary_metadata_count = measuresWithMeta;
                    contextPack.metadata.measures_missing_dictionary_metadata = unclassifiedCols.slice(0, 15).map(c => c.header);
                    contextPack.metadata.measures_with_blank_side = blankSideCols.slice(0, 10).map(c => c.header);
                    
                    // Delta breakdown
                    contextPack.delta_breakdown.total_excluded_by_side_ee = excludedEE.reduce((s, c) => s + c.total, 0);
                    contextPack.delta_breakdown.total_excluded_by_side_na = excludedNA.reduce((s, c) => s + c.total, 0);
                    contextPack.delta_breakdown.total_excluded_by_include_false = excludedIncludeFalse.reduce((s, c) => s + c.total, 0);
                    contextPack.delta_breakdown.total_excluded_by_summary = excludedSummary.reduce((s, c) => s + c.total, 0);
                    contextPack.delta_breakdown.total_unclassified = unclassifiedCols.reduce((s, c) => s + c.total, 0);
                    
                    // Top exclusions
                    const allExcluded = [
                        ...excludedEE.map(e => ({ ...e, reason: "side='ee'" })),
                        ...excludedNA.map(e => ({ ...e, reason: "side='na'" })),
                        ...excludedIncludeFalse.map(e => ({ ...e, reason: "include=false" })),
                        ...excludedSummary
                    ].sort((a, b) => Math.abs(b.total) - Math.abs(a.total));
                    contextPack.delta_breakdown.top_excluded_by_dollars = allExcluded.slice(0, 10);
                    
                    // Top unclassified
                    contextPack.delta_breakdown.top_unclassified_by_dollars = unclassifiedCols
                        .sort((a, b) => Math.abs(b.total) - Math.abs(a.total))
                        .slice(0, 10);
                    
                    // Store measure deltas for later enrichment
                    contextPack.drivers.top_measure_deltas = measureDeltas;
                }
            } else {
                contextPack.availability.error_messages.push("Run Create Matrix to generate PR_Data_Clean first.");
            }

            // --- PR_Archive_Totals Analysis (preferred) ---
            if (!archiveTotalsSheet.isNullObject) {
                contextPack.availability.has_archive_totals = true;

                const archiveTotalsRange = archiveTotalsSheet.getUsedRangeOrNullObject();
                archiveTotalsRange.load("values");
                await context.sync();

                archiveTotalsHasData = !archiveTotalsRange.isNullObject && archiveTotalsRange.values && archiveTotalsRange.values.length > 1;

                if (archiveTotalsHasData) {
                    const totalsPeriods = buildArchivePeriodsFromTotalsSheet(archiveTotalsRange.values);

                    if (totalsPeriods.length > 0) {
                        contextPack.period.periods_available = totalsPeriods.length;

                        const currentKey = contextPack.period.current_key;
                        const priorPeriod = totalsPeriods.find(p => p.key !== currentKey);

                        if (priorPeriod) {
                            contextPack.period.prior_key = priorPeriod.key;
                            contextPack.availability.has_prior_period = true;
                            usedArchiveTotalsForPrior = true;

                            contextPack.totals.expense_review_total_prior = priorPeriod.summary?.total || 0;
                            contextPack.drivers.bucket_totals_prior = {
                                FIXED: priorPeriod.summary?.fixed || 0,
                                VARIABLE: priorPeriod.summary?.variable || 0,
                                BURDEN: priorPeriod.summary?.burden || 0
                            };
                        }
                    }
                }
            }
            
            // --- PR_Archive_Summary Analysis ---
            if (!archiveSheet.isNullObject) {
                contextPack.availability.has_archive_summary = true;
                
                const archiveRange = archiveSheet.getUsedRangeOrNullObject();
                archiveRange.load("values");
                await context.sync();
                
                if (!archiveRange.isNullObject && archiveRange.values && archiveRange.values.length > 1) {
                    const headers = archiveRange.values[0].map(h => String(h || "").trim());
                    const headersLower = headers.map(h => h.toLowerCase());
                    const archiveData = archiveRange.values.slice(1);
                    
                    // Find date column
                    const dateIdx = headersLower.findIndex(h => 
                        h === "pay_date" || h === "payroll_date" || h.includes("pay_period")
                    );
                    
                    if (dateIdx >= 0) {
                        // Group by period key
                        const periodTotals = new Map();
                        const taxonomy = expenseTaxonomyCache;
                        
                        archiveData.forEach(row => {
                            const periodKey = normalizePayDate(row[dateIdx]);
                            if (!periodKey) return;
                            
                            if (!periodTotals.has(periodKey)) {
                                periodTotals.set(periodKey, {
                                    fixed: 0, variable: 0, burden: 0, total: 0,
                                    measureTotals: new Map()
                                });
                            }
                            
                            const period = periodTotals.get(periodKey);
                            
                            // Sum measure columns
                            headers.forEach((header, idx) => {
                                const headerLower = header.toLowerCase();
                                const meta = (taxonomy.loaded && taxonomy.measures) ? taxonomy.measures[headerLower] : null;
                                
                                if (!meta) return;
                                
                                const val = Number(row[idx]) || 0;
                                const signedVal = val * (meta.sign ?? 1);
                                
                                // Check inclusion
                                const inclusionResult = shouldIncludeInExpenseReview(meta, headerLower);
                                if (!inclusionResult.include) return;
                                
                                period.total += signedVal;
                                
                                const bucket = (meta.bucket || "OTHER").toUpperCase();
                                if (bucket === "FIXED") period.fixed += signedVal;
                                else if (bucket === "VARIABLE") period.variable += signedVal;
                                else period.burden += signedVal;
                                
                                // Track per-measure total for delta calculation
                                const currentMeasureTotal = period.measureTotals.get(header) || 0;
                                period.measureTotals.set(header, currentMeasureTotal + signedVal);
                            });
                        });
                        
                        // Sort periods descending
                        const sortedPeriods = Array.from(periodTotals.entries())
                            .sort((a, b) => b[0].localeCompare(a[0]));
                        
                        if (!contextPack.period.periods_available) {
                            contextPack.period.periods_available = sortedPeriods.length;
                        }

                        const currentKey = contextPack.period.current_key;
                        let priorKey = contextPack.period.prior_key;

                        if (!priorKey) {
                            const priorPeriod = sortedPeriods.find(([key]) => key !== currentKey);
                            if (priorPeriod) {
                                priorKey = priorPeriod[0];
                                contextPack.period.prior_key = priorKey;
                                contextPack.availability.has_prior_period = true;
                            }
                        }

                        const priorData = priorKey ? periodTotals.get(priorKey) : null;

                        if (priorData) {
                            if (!usedArchiveTotalsForPrior) {
                                contextPack.availability.has_prior_period = true;
                                contextPack.totals.expense_review_total_prior = priorData.total;
                                contextPack.drivers.bucket_totals_prior = {
                                    FIXED: priorData.fixed,
                                    VARIABLE: priorData.variable,
                                    BURDEN: priorData.burden
                                };
                            }

                            // Enrich measure deltas with prior amounts
                            contextPack.drivers.top_measure_deltas.forEach(measure => {
                                const priorAmount = priorData.measureTotals.get(measure.pf_column_name) || 0;
                                measure.prior_amount = priorAmount;
                                measure.delta_amount = measure.current_amount - priorAmount;
                            });
                        }
                    }
                }
            } else if (!archiveTotalsHasData) {
                contextPack.availability.error_messages.push("Archive data not available. Run archive for at least 1 period to enable comparisons.");
            }
        });
        
        // --- Use Expense Review State for Current Totals ---
        if (expenseReviewState.periods && expenseReviewState.periods.length > 0) {
            const currentPeriod = expenseReviewState.periods[0];
            contextPack.totals.expense_review_total_current = currentPeriod.summary?.total || 0;
            contextPack.totals.calculated_total_current_logic = currentPeriod.summary?.total || 0;
            
            contextPack.drivers.bucket_totals_current = {
                FIXED: currentPeriod.summary?.fixed || 0,
                VARIABLE: currentPeriod.summary?.variable || 0,
                BURDEN: currentPeriod.summary?.burden || 0
            };
            
            // Calculate bucket deltas
            contextPack.drivers.bucket_deltas = {
                FIXED: contextPack.drivers.bucket_totals_current.FIXED - contextPack.drivers.bucket_totals_prior.FIXED,
                VARIABLE: contextPack.drivers.bucket_totals_current.VARIABLE - contextPack.drivers.bucket_totals_prior.VARIABLE,
                BURDEN: contextPack.drivers.bucket_totals_current.BURDEN - contextPack.drivers.bucket_totals_prior.BURDEN
            };
            
            // Department deltas from expense review state
            if (currentPeriod.departments) {
                contextPack.drivers.department_deltas = currentPeriod.departments
                    .filter(d => !d.isTotal)
                    .map(d => ({
                        department_name: d.name,
                        current_amount: d.allIn || 0,
                        prior_amount: (d.allIn || 0) - (d.delta || 0),
                        delta_amount: d.delta || 0
                    }))
                    .sort((a, b) => Math.abs(b.delta_amount) - Math.abs(a.delta_amount))
                    .slice(0, 10);
            }
        }
        
        // Sort measure deltas by absolute delta
        contextPack.drivers.top_measure_deltas = contextPack.drivers.top_measure_deltas
            .sort((a, b) => Math.abs(b.delta_amount) - Math.abs(a.delta_amount))
            .slice(0, 15);
        
        // Bank reconciliation from state
        if (bankReconState.bankAmount) {
            contextPack.totals.bank_statement_amount = Number(bankReconState.bankAmount) || null;
            contextPack.totals.bank_delta = bankReconState.difference;
        }
        
        // Build roster context
        contextPack.roster_context = await buildRosterContext(contextPack.period.current_key, contextPack.period.prior_key);
        if (contextPack.roster_context) {
            contextPack.availability.has_employee_roster = !contextPack.roster_context.error;
            if (contextPack.roster_context.error) {
                contextPack.availability.error_messages.push(contextPack.roster_context.error);
            }
        }
        
        console.log("[AdaInsights] Context pack built:", contextPack);
        return contextPack;
        
    } catch (error) {
        console.error("[AdaInsights] Error building context pack:", error);
        contextPack.availability.error_messages.push(`Error: ${error.message}`);
        return contextPack;
    }
}

/**
 * Normalize a pay date value to YYYY-MM-DD format
 */
function normalizePayDate(dateValue) {
    if (!dateValue) return null;
    
    // If already a string in YYYY-MM-DD format
    if (typeof dateValue === 'string' && /^\d{4}-\d{2}-\d{2}/.test(dateValue)) {
        return dateValue.slice(0, 10);
    }
    
    // If Excel serial number
    if (typeof dateValue === 'number') {
        const date = new Date((dateValue - 25569) * 86400000);
        return date.toISOString().slice(0, 10);
    }
    
    // Try parsing as date
    try {
        const date = new Date(dateValue);
        if (!isNaN(date.getTime())) {
            return date.toISOString().slice(0, 10);
        }
    } catch (_e) {
        // Invalid date format - fall through to string conversion
    }
    
    return String(dateValue).trim();
}

/**
 * Build roster context with employee deltas for Ada Insights
 * Computes new hires, missing employees, reactivations, and department changes
 * 
 * @param {string} currentPeriodKey - Current period key (YYYY-MM-DD)
 * @param {string} priorPeriodKey - Prior period key (YYYY-MM-DD)
 * @returns {Promise<object>} - RosterContext
 */
async function buildRosterContext(currentPeriodKey, priorPeriodKey) {
    console.log("[AdaInsights] Building roster context...", { currentPeriodKey, priorPeriodKey });
    
    const rosterContext = {
        error: null,
        join_key_used: null,
        
        // Roster deltas
        roster_new_this_period: [],
        roster_missing_this_period: [],
        roster_reactivated: [],
        roster_department_changes: [],
        
        // Department headcount bridge
        headcount_by_department_current: {},
        headcount_by_department_prior: {},
        headcount_delta_by_department: {},
        new_hires_by_department: {},
        missing_by_department: {}
    };
    
    /**
     * Normalize employee name for matching
     * Handles extra spaces, casing, punctuation, and suffixes
     */
    function normalizeEmployeeName(name) {
        if (!name) return "";
        return String(name)
            .toLowerCase()
            .trim()
            .replace(/\s+/g, " ")           // Collapse multiple spaces
            .replace(/[.,'"]/g, "")          // Remove punctuation
            .replace(/\s+(jr|sr|ii|iii|iv)$/i, " $1");  // Normalize suffixes
    }
    
    if (!hasExcelRuntime()) {
        rosterContext.error = "Excel runtime unavailable";
        return rosterContext;
    }
    
    try {
        await Excel.run(async (context) => {
            // Load roster and PR_Data_Clean
            const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            
            rosterSheet.load("isNullObject");
            cleanSheet.load("isNullObject");
            await context.sync();
            
            if (rosterSheet.isNullObject) {
                rosterContext.error = "Employee roster not found. Create/update roster from payroll to enable headcount explanations.";
                return;
            }
            
            if (cleanSheet.isNullObject) {
                rosterContext.error = "PR_Data_Clean not found. Run Create Matrix first.";
                return;
            }
            
            // Load roster data
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            rosterRange.load("values");
            
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            cleanRange.load("values");
            
            await context.sync();
            
            if (rosterRange.isNullObject || !rosterRange.values || rosterRange.values.length < 2) {
                rosterContext.error = "Employee roster is empty";
                return;
            }
            
            // Parse roster headers and data
            const rosterHeaders = rosterRange.values[0].map(h => String(h || "").toLowerCase().trim());
            const rosterData = rosterRange.values.slice(1);
            
            const rosterIdxMap = {
                employee_id: rosterHeaders.findIndex(h => h === "employee_id" || h === "emp_id"),
                employee_name: rosterHeaders.findIndex(h => h === "employee_name" || h === "employee"),
                department: rosterHeaders.findIndex(h => h === "department_name" || h === "department"),
                status: rosterHeaders.findIndex(h => h === "employment_status" || h === "status"),
                first_seen: rosterHeaders.findIndex(h => h === "first_seen_pay_date"),
                last_seen: rosterHeaders.findIndex(h => h === "last_seen_pay_date"),
                termination_date: rosterHeaders.findIndex(h => h === "termination_effective_date"),
                is_manual: rosterHeaders.findIndex(h => h === "is_manually_managed")
            };
            
            // ============================================================================
            // Parse PR_Data_Clean FIRST - need this to determine correct join key
            // ============================================================================
            let cleanHeaders = [];
            let cleanData = [];
            let cleanIdxMap = { employee_id: -1, employee_name: -1, department: -1 };
            
            if (!cleanRange.isNullObject && cleanRange.values && cleanRange.values.length > 1) {
                cleanHeaders = cleanRange.values[0].map(h => String(h || "").toLowerCase().trim());
                cleanData = cleanRange.values.slice(1);
                
                cleanIdxMap = {
                    employee_id: cleanHeaders.findIndex(h => h === "employee_id" || h === "emp_id"),
                    employee_name: cleanHeaders.findIndex(h => h === "employee_name" || h === "employee"),
                    department: cleanHeaders.findIndex(h => h === "department_name" || h === "department")
                };
            }
            
            // ============================================================================
            // DETERMINE JOIN KEY TYPE - Check if BOTH datasets have Employee_ID with VALUES
            // Must happen BEFORE building rosterMap!
            // ============================================================================
            const rosterHasIdColumn = rosterIdxMap.employee_id >= 0;
            const cleanHasIdColumn = cleanIdxMap.employee_id >= 0;
            
            let rosterIdHasValues = false;
            let cleanIdHasValues = false;
            
            if (rosterHasIdColumn) {
                rosterIdHasValues = rosterData.some(row => {
                    const val = String(row[rosterIdxMap.employee_id] || "").trim();
                    // Check it's not empty AND not just a name (names contain spaces)
                    return val.length > 0 && !val.includes(" ");
                });
            }
            
            if (cleanHasIdColumn) {
                cleanIdHasValues = cleanData.some(row => {
                    const val = String(row[cleanIdxMap.employee_id] || "").trim();
                    return val.length > 0 && !val.includes(" ");
                });
            }
            
            // Only use Employee_ID if BOTH datasets have the column AND have actual ID values
            const useEmployeeId = rosterHasIdColumn && cleanHasIdColumn && rosterIdHasValues && cleanIdHasValues;
            const joinKeyType = useEmployeeId ? "Employee_ID" : "Employee_Name";
            rosterContext.join_key_used = joinKeyType;
            
            console.log(`[RosterContext] Join key determination:`);
            console.log(`  - Roster has Employee_ID column: ${rosterHasIdColumn}, with values: ${rosterIdHasValues}`);
            console.log(`  - Payroll has Employee_ID column: ${cleanHasIdColumn}, with values: ${cleanIdHasValues}`);
            console.log(`  - Using join key: ${joinKeyType}`);
            
            // ============================================================================
            // BUILD ROSTER MAP - Using the correctly determined join key type
            // ============================================================================
            const rosterMap = new Map(); // joinKey -> roster row data
            
            rosterData.forEach(row => {
                let joinKey;
                if (useEmployeeId) {
                    joinKey = String(row[rosterIdxMap.employee_id] || "").toLowerCase().trim();
                } else {
                    // Use Employee_Name - normalize for better matching
                    joinKey = normalizeEmployeeName(row[rosterIdxMap.employee_name]);
                }
                
                if (!joinKey) return;
                
                rosterMap.set(joinKey, {
                    employee_id: rosterIdxMap.employee_id >= 0 ? String(row[rosterIdxMap.employee_id] || "").trim() : null,
                    employee_name: rosterIdxMap.employee_name >= 0 ? String(row[rosterIdxMap.employee_name] || "").trim() : null,
                    department: rosterIdxMap.department >= 0 ? String(row[rosterIdxMap.department] || "").trim() : null,
                    status: rosterIdxMap.status >= 0 ? String(row[rosterIdxMap.status] || "").trim() : "Unknown",
                    first_seen: rosterIdxMap.first_seen >= 0 ? normalizePayDate(row[rosterIdxMap.first_seen]) : null,
                    last_seen: rosterIdxMap.last_seen >= 0 ? normalizePayDate(row[rosterIdxMap.last_seen]) : null,
                    termination_date: rosterIdxMap.termination_date >= 0 ? normalizePayDate(row[rosterIdxMap.termination_date]) : null,
                    is_manual: rosterIdxMap.is_manual >= 0 && String(row[rosterIdxMap.is_manual] || "").toLowerCase() === "true"
                });
            });
            
            console.log(`[RosterContext] Built rosterMap with ${rosterMap.size} entries using ${joinKeyType}`);
            
            // ============================================================================
            // BUILD CURRENT PAYROLL SET - Using the same join key type
            // ============================================================================
            if (cleanData.length > 0) {
                // Build current payroll set
                const currentPayrollSet = new Map(); // joinKey -> { name, department }
                
                cleanData.forEach(row => {
                    let joinKey;
                    if (useEmployeeId) {
                        joinKey = String(row[cleanIdxMap.employee_id] || "").toLowerCase().trim();
                    } else {
                        // Use Employee_Name - normalize for better matching
                        joinKey = normalizeEmployeeName(row[cleanIdxMap.employee_name]);
                    }
                    
                    if (!joinKey) return;
                    
                    if (!currentPayrollSet.has(joinKey)) {
                        currentPayrollSet.set(joinKey, {
                            name: cleanIdxMap.employee_name >= 0 ? String(row[cleanIdxMap.employee_name] || "").trim() : joinKey,
                            department: cleanIdxMap.department >= 0 ? String(row[cleanIdxMap.department] || "").trim() : "Unknown"
                        });
                    }
                });
                
                console.log(`[RosterContext] Built currentPayrollSet with ${currentPayrollSet.size} entries using ${joinKeyType}`);
                
                // Department headcount current
                currentPayrollSet.forEach(emp => {
                    const dept = emp.department || "Unknown";
                    rosterContext.headcount_by_department_current[dept] = 
                        (rosterContext.headcount_by_department_current[dept] || 0) + 1;
                });
                
                // Compare: find new employees and reactivations
                currentPayrollSet.forEach((emp, joinKey) => {
                    const rosterEntry = rosterMap.get(joinKey);
                    
                    if (!rosterEntry) {
                        // New employee - not in roster
                        rosterContext.roster_new_this_period.push({
                            key: joinKey,
                            name: emp.name,
                            department: emp.department,
                            note: "first seen this period (likely new hire)"
                        });
                        
                        rosterContext.new_hires_by_department[emp.department] = 
                            (rosterContext.new_hires_by_department[emp.department] || 0) + 1;
                    } else if (rosterEntry.status?.toLowerCase() === "terminated") {
                        // Reactivation - terminated but appearing in payroll
                        rosterContext.roster_reactivated.push({
                            key: joinKey,
                            name: rosterEntry.employee_name || emp.name,
                            department: emp.department,
                            termination_date: rosterEntry.termination_date,
                            is_manually_managed: rosterEntry.is_manual,
                            note: "reactivation detected (was terminated)"
                        });
                    } else if (rosterEntry.first_seen === currentPeriodKey) {
                        // New based on first_seen date
                        rosterContext.roster_new_this_period.push({
                            key: joinKey,
                            name: rosterEntry.employee_name || emp.name,
                            department: emp.department,
                            note: "first seen this period (new hire)"
                        });
                        
                        rosterContext.new_hires_by_department[emp.department] = 
                            (rosterContext.new_hires_by_department[emp.department] || 0) + 1;
                    } else if (rosterEntry.department && emp.department && 
                               rosterEntry.department.toLowerCase() !== emp.department.toLowerCase()) {
                        // Department change
                        rosterContext.roster_department_changes.push({
                            key: joinKey,
                            name: rosterEntry.employee_name || emp.name,
                            previous_department: rosterEntry.department,
                            current_department: emp.department,
                            note: "department changed"
                        });
                    }
                });
                
                // Compare: find missing employees (in roster but not in current payroll)
                rosterMap.forEach((rosterEntry, joinKey) => {
                    if (!currentPayrollSet.has(joinKey)) {
                        const status = rosterEntry.status?.toLowerCase();
                        
                        // Only flag active employees as missing
                        if (status === "active" || status === "unknown") {
                            rosterContext.roster_missing_this_period.push({
                                key: joinKey,
                                name: rosterEntry.employee_name,
                                department: rosterEntry.department,
                                last_seen_date: rosterEntry.last_seen,
                                note: "not seen this period (may be no-hours or termination)"
                            });
                            
                            const dept = rosterEntry.department || "Unknown";
                            rosterContext.missing_by_department[dept] = 
                                (rosterContext.missing_by_department[dept] || 0) + 1;
                        }
                    }
                });
                
                // Calculate prior period headcount (approximate from roster)
                // Use roster entries that were active as of prior period
                if (priorPeriodKey) {
                    rosterMap.forEach((entry) => {
                        const firstSeen = entry.first_seen;
                        const lastSeen = entry.last_seen;
                        
                        // Was this employee present in prior period?
                        const wasPresent = firstSeen && firstSeen <= priorPeriodKey &&
                            (!lastSeen || lastSeen >= priorPeriodKey);
                        
                        if (wasPresent) {
                            const dept = entry.department || "Unknown";
                            rosterContext.headcount_by_department_prior[dept] = 
                                (rosterContext.headcount_by_department_prior[dept] || 0) + 1;
                        }
                    });
                }
                
                // Calculate department headcount deltas
                const allDepts = new Set([
                    ...Object.keys(rosterContext.headcount_by_department_current),
                    ...Object.keys(rosterContext.headcount_by_department_prior)
                ]);
                
                allDepts.forEach(dept => {
                    const current = rosterContext.headcount_by_department_current[dept] || 0;
                    const prior = rosterContext.headcount_by_department_prior[dept] || 0;
                    rosterContext.headcount_delta_by_department[dept] = current - prior;
                });
            }
            
            // Cap lists to avoid huge context
            rosterContext.roster_new_this_period = rosterContext.roster_new_this_period.slice(0, 20);
            rosterContext.roster_missing_this_period = rosterContext.roster_missing_this_period.slice(0, 20);
            rosterContext.roster_reactivated = rosterContext.roster_reactivated.slice(0, 10);
            rosterContext.roster_department_changes = rosterContext.roster_department_changes.slice(0, 10);
        });
        
        console.log("[AdaInsights] Roster context built:", rosterContext);
        return rosterContext;
        
    } catch (error) {
        console.error("[AdaInsights] Error building roster context:", error);
        rosterContext.error = `Error: ${error.message}`;
        return rosterContext;
    }
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

/**
 * Render a button with a label underneath
 */
function renderLabeledButton(buttonHtml, label) {
    return `
        <div class="pf-labeled-button">
            ${buttonHtml}
            <span class="pf-button-label">${escapeHtml(label)}</span>
        </div>
    `;
}

function hasExcelRuntime() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

function getStepNoteFields(stepId) {
    return STEP_NOTES_FIELDS[stepId] || null;
}


function formatMetricValue(value) {
    if (value === null || value === undefined) return "---";
    if (typeof value === "number" && Number.isInteger(value)) return value.toString();
    return value;
}

function formatSignedValue(value) {
    if (value === null || value === undefined) return "---";
    if (typeof value !== "number" || Number.isNaN(value)) return "---";
    if (value === 0) return "0";
    return value > 0 ? `+${value}` : value.toString();
}

function formatCurrency(value) {
    if (value === null || value === undefined || Number.isNaN(value)) return "---";
    if (typeof value !== "number") return value;
    return value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatNumberDisplay(value) {
    if (value === null || value === undefined || value === "") return "---";
    const num = Number(value);
    const formatter = window.PrairieForge?.formatNumber;
    if (Number.isFinite(num)) {
        return formatter ? formatter(num) : num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(value);
}

function formatBankInput(value) {
    const numeric = parseBankAmount(value);
    if (!Number.isFinite(numeric)) return "";
    // Format as XXX,XXX.XX (no dollar sign per user preference)
    return numeric.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function buildCsv(rows = []) {
    return rows
        .map((row) =>
            (row || [])
                .map((cell) => {
                    if (cell == null) return "";
                    const str = String(cell);
                    if (/[",\n]/.test(str)) return `"${str.replace(/"/g, '""')}"`;
                    return str;
                })
                .join(",")
        )
        .join("\n");
}

function downloadCsv(filename, content) {
    const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function parseBankAmount(value) {
    if (typeof value === "number") return value;
    if (value == null) return NaN;
    const cleaned = String(value).replace(/[^0-9.-]/g, "");
    const parsed = Number.parseFloat(cleaned);
    return Number.isFinite(parsed) ? parsed : NaN;
}

// =============================================================================
// BANK RECONCILIATION SERVICE
// Provides system total computation and bank reconciliation state management
// =============================================================================

/**
 * Config field name for persisting bank amount
 */
const BANK_AMOUNT_CONFIG_FIELD = "PR_Bank_Amount";

/**
 * Tolerance for bank reconciliation (within this amount = reconciled)
 */
const BANK_RECONCILIATION_TOLERANCE = 0.01;

/**
 * Bank reconciliation state (also uses validationState for backward compat)
 */
const bankReconState = {
    systemTotal: null,      // Sum of amount columns from PR_Data_Clean
    bankAmount: null,       // User-entered bank statement amount
    difference: null,       // bankAmount - systemTotal
    isReconciled: false,    // Within tolerance
    loading: false,
    lastError: null,
    amountColumns: []       // List of columns used for system total
};

// =============================================================================
// PR_DATA_CLEAN MEASURE UNIVERSE - Single Source of Truth
// =============================================================================
// This is the AUTHORITATIVE source for what columns are included in totals.
// Both Step 1 (Bank Reconciliation) and Expense Review MUST use this.
// Dictionary metadata is for enrichment/bucketing only - NOT for inclusion/exclusion.

/**
 * Hard exclusions - summary columns that should NEVER be included
 * These are totals/summaries that would cause double-counting
 */
const MEASURE_UNIVERSE_EXCLUSIONS = new Set([
    "gross_pay_amount",
    "net_pay_amount",
    "total_deductions_amount",
    "total_earnings_amount"
]);

/**
 * Dimension patterns - columns that are for grouping, not summing
 */
const DIMENSION_PATTERNS = [
    "employee", "department", "location", "date", "period", 
    "name", "id", "code", "title", "check", "pay_type", "frequency"
];

/**
 * Cache for the measure universe (cleared when PR_Data_Clean changes)
 */
let measureUniverseCache = null;

/**
 * Get the PR_Data_Clean Measure Universe - SINGLE SOURCE OF TRUTH
 * 
 * This function defines EXACTLY which columns are included in totals.
 * Rules:
 * 1. Column header contains "amount" (case-insensitive)
 * 2. Column is NOT in MEASURE_UNIVERSE_EXCLUSIONS
 * 3. Column is NOT a dimension (doesn't match DIMENSION_PATTERNS)
 * 
 * Returns:
 * - sheetName: "PR_Data_Clean"
 * - usedRangeAddress: Excel range address
 * - allHeaders: all column headers
 * - dimensionHeaders: headers classified as dimensions
 * - includedMeasureHeaders: headers included in totals (THE AUTHORITATIVE LIST)
 * - excludedHeaders: headers excluded with reasons
 * - perColumnSums: { header: sum } for each included measure
 * - total: sum of all included measures
 * - error: null or error message
 */
async function getPRDataCleanMeasureUniverse() {
    console.log("[MeasureUniverse] Getting PR_Data_Clean measure universe...");
    
    // Return cache if available
    if (measureUniverseCache) {
        console.log("[MeasureUniverse] Returning cached universe");
        return measureUniverseCache;
    }
    
    const result = {
        sheetName: "PR_Data_Clean",
        usedRangeAddress: null,
        allHeaders: [],
        dimensionHeaders: [],
        includedMeasureHeaders: [],
        excludedHeaders: [],
        perColumnSums: {},
        total: 0,
        error: null
    };
    
    if (!hasExcelRuntime()) {
        result.error = "Excel runtime unavailable";
        return result;
    }
    
    try {
        await Excel.run(async (context) => {
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            cleanSheet.load("isNullObject");
            await context.sync();
            
            if (cleanSheet.isNullObject) {
                result.error = "Create Matrix first - PR_Data_Clean not found";
                return;
            }
            
            const usedRange = cleanSheet.getUsedRangeOrNullObject();
            usedRange.load(["address", "values", "rowCount", "columnCount"]);
            await context.sync();
            
            if (usedRange.isNullObject || usedRange.rowCount < 2) {
                result.error = "PR_Data_Clean is empty";
                return;
            }
            
            result.usedRangeAddress = usedRange.address;
            const values = usedRange.values;
            const headers = values[0].map(h => String(h || "").trim());
            result.allHeaders = [...headers];
            
            // Classify each column
            headers.forEach((header, idx) => {
                const headerLower = header.toLowerCase();
                
                // Check if it's a dimension
                const isDimension = DIMENSION_PATTERNS.some(p => headerLower.includes(p));
                if (isDimension) {
                    result.dimensionHeaders.push(header);
                    result.excludedHeaders.push({ header, reason: "dimension" });
                    return;
                }
                
                // Check if it contains "amount" (primary measure indicator)
                const hasAmount = headerLower.includes("amount");
                if (!hasAmount) {
                    result.excludedHeaders.push({ header, reason: "no_amount_keyword" });
                    return;
                }
                
                // Check hard exclusions (summary totals)
                if (MEASURE_UNIVERSE_EXCLUSIONS.has(headerLower)) {
                    result.excludedHeaders.push({ header, reason: "summary_exclusion" });
                    return;
                }
                
                // This column is INCLUDED - compute its sum
                let colSum = 0;
                for (let rowIdx = 1; rowIdx < values.length; rowIdx++) {
                    const cellValue = values[rowIdx][idx];
                    const numValue = Number(cellValue);
                    if (Number.isFinite(numValue)) {
                        colSum += numValue;
                    }
                }
                
                result.includedMeasureHeaders.push(header);
                result.perColumnSums[header] = colSum;
                result.total += colSum;
            });
        });
        
        // Debug output
        console.log("[MeasureUniverse] ═══════════════════════════════════════════════════════");
        console.log(`[MeasureUniverse] Total: $${result.total.toLocaleString(undefined, { minimumFractionDigits: 2 })}`);
        console.log(`[MeasureUniverse] Included measures: ${result.includedMeasureHeaders.length}`);
        console.log(`[MeasureUniverse] Dimensions: ${result.dimensionHeaders.length}`);
        console.log(`[MeasureUniverse] Excluded: ${result.excludedHeaders.length}`);
        
        // Top 10 by dollars
        const sortedBySum = Object.entries(result.perColumnSums)
            .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
        console.log("[MeasureUniverse] Top 10 included by dollars:");
        sortedBySum.slice(0, 10).forEach(([h, sum], i) => {
            console.log(`   ${i+1}. ${h}: $${sum.toLocaleString()}`);
        });
        
        console.log("[MeasureUniverse] ═══════════════════════════════════════════════════════");
        
        // Cache the result
        measureUniverseCache = result;
        
    } catch (error) {
        console.error("[MeasureUniverse] Error:", error);
        result.error = error.message;
    }
    
    return result;
}

/**
 * Invalidate the measure universe cache (call when PR_Data_Clean changes)
 */
function invalidateMeasureUniverseCache() {
    console.log("[MeasureUniverse] Cache invalidated");
    measureUniverseCache = null;
}

/**
 * Get system total from PR_Data_Clean by summing amount columns
 * NOW USES getPRDataCleanMeasureUniverse() as single source of truth
 */
async function computeSystemTotalFromPRDataClean() {
    const universe = await getPRDataCleanMeasureUniverse();
    
    if (universe.error) {
        return { total: null, error: universe.error, columns: [] };
    }
    
    return { 
        total: universe.total, 
        error: null, 
        columns: universe.includedMeasureHeaders 
    };
}

/**
 * Load bank amount from config storage
 */
async function loadBankAmountFromConfig() {
    try {
        const storedValue = getConfigValue(BANK_AMOUNT_CONFIG_FIELD);
        const parsed = parseBankAmount(storedValue);
        if (Number.isFinite(parsed)) {
            bankReconState.bankAmount = parsed;
            validationState.bankAmount = parsed;
            console.log("[BankRecon] Loaded bank amount from config:", parsed);
        }
    } catch (error) {
        console.warn("[BankRecon] Error loading bank amount from config:", error);
    }
}

/**
 * Save bank amount to config storage (persisted)
 */
function saveBankAmountToConfig(amount) {
    const numericAmount = parseBankAmount(amount);
    if (Number.isFinite(numericAmount)) {
        scheduleConfigWrite(BANK_AMOUNT_CONFIG_FIELD, numericAmount);
        console.log("[BankRecon] Saved bank amount to config:", numericAmount);
    }
}

/**
 * Update bank reconciliation state and compute difference
 */
function updateBankReconState(partial = {}) {
    Object.assign(bankReconState, partial);
    
    const systemTotal = Number(bankReconState.systemTotal);
    const bankAmount = Number(bankReconState.bankAmount);
    
    if (Number.isFinite(systemTotal) && Number.isFinite(bankAmount)) {
        bankReconState.difference = bankAmount - systemTotal;
        bankReconState.isReconciled = Math.abs(bankReconState.difference) <= BANK_RECONCILIATION_TOLERANCE;
    } else {
        bankReconState.difference = null;
        bankReconState.isReconciled = false;
    }
    
    // Also update validationState for backward compat
    validationState.cleanTotal = bankReconState.systemTotal;
    validationState.bankAmount = bankReconState.bankAmount;
    validationState.bankDifference = bankReconState.difference;
    
    console.log("[BankRecon] State updated:", {
        systemTotal: bankReconState.systemTotal,
        bankAmount: bankReconState.bankAmount,
        difference: bankReconState.difference,
        isReconciled: bankReconState.isReconciled
    });
}

/**
 * Refresh bank reconciliation data (compute system total, update state)
 */
async function refreshBankReconciliation() {
    bankReconState.loading = true;
    bankReconState.lastError = null;
    
    try {
        // Load bank amount from config if not already loaded
        if (bankReconState.bankAmount === null) {
            await loadBankAmountFromConfig();
        }
        
        // Compute system total from PR_Data_Clean
        const result = await computeSystemTotalFromPRDataClean();
        
        updateBankReconState({
            systemTotal: result.total,
            amountColumns: result.columns,
            lastError: result.error,
            loading: false
        });
        
        return result;
    } catch (error) {
        bankReconState.loading = false;
        bankReconState.lastError = error.message;
        console.error("[BankRecon] Refresh failed:", error);
        return { total: null, error: error.message, columns: [] };
    }
}

/**
 * Handle bank amount input change in Step 1
 */
function handleStep1BankAmountInput(event) {
    const inputEl = event?.target && event.target instanceof HTMLInputElement
        ? event.target
        : document.getElementById("step1-bank-amount-input");
    
    const numeric = parseBankAmount(inputEl?.value);
    const formatted = formatBankInput(numeric);
    
    if (inputEl) {
        inputEl.value = formatted;
    }
    
    // Update state and persist
    updateBankReconState({ bankAmount: numeric });
    saveBankAmountToConfig(numeric);
    
    // Update UI displays
    updateBankReconciliationUI();
}

/**
 * Update bank reconciliation UI elements in Step 1
 */
function updateBankReconciliationUI() {
    const systemTotalEl = document.getElementById("step1-system-total-value");
    const diffEl = document.getElementById("step1-bank-diff-value");
    const hintEl = document.getElementById("step1-bank-diff-hint");
    const statusEl = document.getElementById("step1-recon-status");
    
    if (systemTotalEl) {
        systemTotalEl.value = formatCurrency(bankReconState.systemTotal);
    }
    
    if (diffEl) {
        diffEl.value = bankReconState.difference != null 
            ? formatCurrency(bankReconState.difference) 
            : "---";
    }
    
    if (hintEl) {
        if (bankReconState.difference === null) {
            hintEl.textContent = "";
        } else if (bankReconState.isReconciled) {
            hintEl.textContent = "Reconciled";
            hintEl.className = "pf-metric-hint pf-metric-hint--success";
        } else {
            hintEl.textContent = `Difference of ${formatCurrency(Math.abs(bankReconState.difference))} exceeds tolerance.`;
            hintEl.className = "pf-metric-hint pf-metric-hint--warning";
        }
    }
    
    if (statusEl) {
        if (bankReconState.lastError) {
            statusEl.innerHTML = `<p class="pf-step-note">${escapeHtml(bankReconState.lastError)}</p>`;
        } else {
            statusEl.innerHTML = "";
        }
    }
}

/**
 * Render bank reconciliation card HTML for Step 1
 * Status badges include text labels for accessibility (not color-only)
 */
function renderBankReconciliationCard() {
    const hasSystemTotal = bankReconState.systemTotal !== null;
    const systemTotal = hasSystemTotal ? formatCurrency(bankReconState.systemTotal) : "---";
    const bankValue = formatBankInput(bankReconState.bankAmount);
    const diffValue = bankReconState.difference != null ? formatCurrency(bankReconState.difference) : "---";
    
    let hintText = "";
    let hintClass = "pf-metric-hint";
    let statusBadge = "";
    let statusAriaLabel = "";
    
    if (bankReconState.loading) {
        // Pending = validation is running
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status"><span>Loading</span></span>`;
        hintText = "Calculating system total...";
        statusAriaLabel = "Bank reconciliation loading";
    } else if (!hasSystemTotal && !bankReconState.lastError) {
        // Unavailable = cannot run due to missing prerequisites
        statusBadge = `<span class="pf-status-badge pf-status-badge--unavailable" role="status"><span>Pending</span></span>`;
        hintText = "Click 'Create Matrix' above to generate PR_Data_Clean, then the system total will appear here.";
        hintClass = "pf-metric-hint pf-metric-hint--info";
        statusAriaLabel = "Bank reconciliation pending - waiting for PR_Data_Clean";
    } else if (bankReconState.lastError) {
        // Unavailable = error state
        statusBadge = `<span class="pf-status-badge pf-status-badge--unavailable" role="status"><span>Unavailable</span></span>`;
        hintText = bankReconState.lastError;
        hintClass = "pf-metric-hint pf-metric-hint--error";
        statusAriaLabel = "Bank reconciliation unavailable: " + bankReconState.lastError;
    } else if (bankReconState.difference !== null) {
        if (bankReconState.isReconciled) {
            statusBadge = `<span class="pf-status-badge pf-status-badge--ok" role="status"><span>Reconciled</span></span>`;
            hintText = "System total matches bank statement within tolerance.";
            hintClass = "pf-metric-hint pf-metric-hint--success";
            statusAriaLabel = "Bank reconciliation passed";
        } else {
            statusBadge = `<span class="pf-status-badge pf-status-badge--review" role="status"><span>Review</span></span>`;
            hintText = `Difference of ${formatCurrency(Math.abs(bankReconState.difference))} exceeds tolerance.`;
            hintClass = "pf-metric-hint pf-metric-hint--warning";
            statusAriaLabel = "Bank reconciliation needs review - difference exceeds tolerance";
        }
    } else if (hasSystemTotal && !bankReconState.bankAmount) {
        // Ready to compare but no bank amount entered
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status"><span>Pending</span></span>`;
        hintText = "Enter your bank statement amount to compare.";
        hintClass = "pf-metric-hint pf-metric-hint--info";
        statusAriaLabel = "Bank reconciliation pending - enter bank amount";
    }
    
    // Show refresh button if we have system total (subtle placement)
    const refreshButton = hasSystemTotal ? `
        <button type="button" class="pf-action-toggle pf-action-toggle--subtle pf-clickable" id="bank-recon-refresh-btn" title="Refresh system total" style="margin-left: auto;">
            ${REFRESH_ICON_SVG}
        </button>
    ` : "";
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" id="bank-recon-card" aria-label="${statusAriaLabel}">
            <div class="pf-config-head" style="display: flex; align-items: center;">
                <div>
                    <h3>Bank Reconciliation ${statusBadge}</h3>
                    <p class="pf-config-subtext">Compare payroll total to the amount pulled from the bank.</p>
                </div>
                ${refreshButton}
            </div>
            <div id="step1-recon-status"></div>
            <div class="pf-config-grid pf-metric-grid">
                <label class="pf-config-field">
                    <span>System Total (PR_Data_Clean)</span>
                    <input id="step1-system-total-value" type="text" class="pf-readonly-input pf-metric-value" value="${systemTotal}" readonly aria-label="System total from PR_Data_Clean">
                </label>
                <label class="pf-config-field">
                    <span>Bank Statement Amount</span>
                    <input
                        type="text"
                        inputmode="decimal"
                        id="step1-bank-amount-input"
                        class="pf-metric-input"
                        value="${escapeHtml(bankValue)}"
                        placeholder="0.00"
                        aria-label="Enter bank statement amount"
                        ${!hasSystemTotal ? 'disabled' : ''}
                    >
                </label>
                <label class="pf-config-field">
                    <span>Difference</span>
                    <input id="step1-bank-diff-value" type="text" class="pf-readonly-input pf-metric-value" value="${diffValue}" readonly aria-label="Difference between system and bank totals">
                </label>
            </div>
            <p class="${hintClass}" id="step1-bank-diff-hint">${escapeHtml(hintText)}</p>
        </article>
    `;
}

// =============================================================================
// PAYROLL COVERAGE SERVICE
// Compares roster employees/departments vs PR_Data_Clean (advisory checks)
// =============================================================================

/**
 * Payroll coverage state - tracks employee and department coverage
 */
const payrollCoverageState = {
    loading: false,
    lastError: null,
    hasData: false,
    joinKeyUsed: null,  // "Employee_ID" | "Employee_Name" - documents which key was used
    employee: {
        rosterCount: 0,
        payrollCount: 0,
        missingFromPayroll: [],   // Roster employees not in payroll
        extraInPayroll: [],       // Payroll employees not in roster
        status: "pending"         // "ok" | "review" | "pending" | "unavailable"
    },
    department: {
        rosterDepts: [],
        payrollDepts: [],
        missingFromPayroll: [],   // Roster depts not in payroll
        extraInPayroll: [],       // Payroll depts not in roster
        status: "pending"
    }
};

/**
 * Maximum number of names/depts to show in the UI before "+X more"
 */
const COVERAGE_MAX_DISPLAY_ITEMS = 8;

/**
 * Normalize a value for join key comparison
 * - trim whitespace
 * - lowercase
 * - collapse multiple spaces
 * - handle null/undefined
 */
function normalizeJoinKey(value) {
    if (value === null || value === undefined) return "";
    return String(value)
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ")           // Collapse multiple spaces
        .replace(/[.,'"]/g, "")          // Remove punctuation (match normalizeEmployeeName)
        .replace(/\s+(jr|sr|ii|iii|iv)$/i, " $1");  // Normalize suffixes
}

/**
 * Check if a column has actual values (not just headers with empty data)
 * @param {Array<Array>} dataRows - Data rows (excluding header)
 * @param {number} colIdx - Column index to check
 * @returns {boolean} True if at least some rows have non-empty values
 */
function columnHasValues(dataRows, colIdx) {
    if (colIdx < 0 || !dataRows || dataRows.length === 0) return false;
    
    // Check if at least 10% of rows have values (or at least 1 row)
    const minRows = Math.max(1, Math.floor(dataRows.length * 0.1));
    let foundCount = 0;
    
    for (const row of dataRows) {
        const val = String(row[colIdx] || "").trim();
        // Must be non-empty AND not look like a name (IDs don't have spaces typically)
        if (val.length > 0 && !val.includes(" ")) {
            foundCount++;
            if (foundCount >= minRows) return true;
        }
    }
    return false;
}

/**
 * Find the best join key column index in headers
 * Priority: Employee_ID > Employee_Name
 * BUT only use Employee_ID if it actually has values!
 * @param {Array<string>} headers - Normalized header names
 * @param {Array<Array>} dataRows - Optional data rows to check for values
 * @returns {{ index: number, type: "Employee_ID" | "Employee_Name" | null }}
 */
function findEmployeeJoinKeyColumn(headers, dataRows = null) {
    // Priority 1: Employee_ID (but only if it has actual values)
    const idIdx = headers.findIndex(h => 
        h.includes("employee") && h.includes("id") && !h.includes("name")
    );
    
    // Fallback for "emp_id", "empid", etc.
    const altIdIdx = headers.findIndex(h => 
        (h.includes("emp") && h.includes("id")) || h === "empid" || h === "emp_id"
    );
    
    const employeeIdIdx = idIdx >= 0 ? idIdx : altIdIdx;
    
    // Only use Employee_ID if we have data rows AND the column has values
    if (employeeIdIdx >= 0) {
        if (!dataRows || columnHasValues(dataRows, employeeIdIdx)) {
            return { index: employeeIdIdx, type: "Employee_ID" };
        }
        console.log(`[JoinKey] Employee_ID column found at index ${employeeIdIdx} but has no values - falling back to Employee_Name`);
    }
    
    // Priority 2: Employee_Name
    const nameIdx = headers.findIndex(h => 
        h.includes("employee") && (h.includes("name") || !h.includes("id"))
    );
    if (nameIdx >= 0) {
        return { index: nameIdx, type: "Employee_Name" };
    }
    
    // Fallback: Any column with "employee"
    const fallbackIdx = headers.findIndex(h => h.includes("employee"));
    if (fallbackIdx >= 0) {
        return { index: fallbackIdx, type: "Employee_Name" };
    }
    
    return { index: -1, type: null };
}

/**
 * Compute payroll coverage by comparing roster vs PR_Data_Clean
 * Advisory only - does not block workflow
 * 
 * Join Key Priority:
 * 1. Employee_ID (if present in both datasets)
 * 2. Employee_Name (fallback)
 */
async function refreshPayrollCoverage() {
    if (!hasExcelRuntime()) {
        payrollCoverageState.lastError = "Excel runtime unavailable.";
        payrollCoverageState.hasData = false;
        payrollCoverageState.employee.status = "unavailable";
        payrollCoverageState.department.status = "unavailable";
        return;
    }
    
    payrollCoverageState.loading = true;
    payrollCoverageState.lastError = null;
    payrollCoverageState.employee.status = "pending";
    payrollCoverageState.department.status = "pending";
    
    try {
        const result = await Excel.run(async (context) => {
            // Get roster sheet
            const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
            rosterSheet.load("isNullObject");
            await context.sync();
            
            if (rosterSheet.isNullObject) {
                return { 
                    error: "Connect or create an employee roster (SS_Employee_Roster) to enable coverage checks.",
                    status: "unavailable"
                };
            }
            
            // Get PR_Data_Clean sheet
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            cleanSheet.load("isNullObject");
            await context.sync();
            
            if (cleanSheet.isNullObject) {
                return { 
                    error: "Run 'Create Matrix' to generate PR_Data_Clean first.",
                    status: "unavailable"
                };
            }
            
            // OPTIMIZATION: Read header rows first to identify required columns
            const rosterHeaderRange = rosterSheet.getRangeByIndexes(0, 0, 1, 20); // First row, up to 20 cols
            const cleanHeaderRange = cleanSheet.getRangeByIndexes(0, 0, 1, 50); // First row, up to 50 cols
            rosterHeaderRange.load("values");
            cleanHeaderRange.load("values");
            await context.sync();
            
            const rosterHeaders = (rosterHeaderRange.values[0] || []).map(h => normalizeHeader(String(h || "")));
            const cleanHeaders = (cleanHeaderRange.values[0] || []).map(h => normalizeHeader(String(h || "")));
            
            // Find potential join key column indices (will verify with data later)
            const rosterIdIdx = rosterHeaders.findIndex(h => h.includes("employee") && h.includes("id") && !h.includes("name"));
            const cleanIdIdx = cleanHeaders.findIndex(h => h.includes("employee") && h.includes("id") && !h.includes("name"));
            const rosterNameIdx = rosterHeaders.findIndex(h => h.includes("employee") && (h.includes("name") || !h.includes("id")));
            const cleanNameIdx = cleanHeaders.findIndex(h => h.includes("employee") && (h.includes("name") || !h.includes("id")));
            
            // We need at least Employee_Name in both
            if (rosterNameIdx < 0 || cleanNameIdx < 0) {
                const missingIn = [];
                if (rosterNameIdx < 0) missingIn.push("roster");
                if (cleanNameIdx < 0) missingIn.push("PR_Data_Clean");
                return { 
                    error: `Employee identifier column not found in ${missingIn.join(" and ")}. Check your column mappings.`,
                    status: "unavailable"
                };
            }
            
            // Find department column indices
            const rosterDeptIdx = pickDepartmentIndex(rosterHeaders);
            const cleanDeptIdx = pickDepartmentIndex(cleanHeaders);
            
            // Find status column in roster to filter out terminated employees
            const rosterStatusIdx = rosterHeaders.findIndex(h => 
                h === "employment_status" || h === "status"
            );
            
            // OPTIMIZATION: Read only required columns from both sheets
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            rosterRange.load(["values", "rowCount"]);
            cleanRange.load(["values", "rowCount"]);
            await context.sync();
            
            if (rosterRange.isNullObject || rosterRange.rowCount < 2) {
                return { error: "Employee roster is empty.", status: "unavailable" };
            }
            
            if (cleanRange.isNullObject || cleanRange.rowCount < 2) {
                return { error: "PR_Data_Clean is empty. Run 'Create Matrix' first.", status: "unavailable" };
            }
            
            const rosterValues = rosterRange.values || [];
            const cleanValues = cleanRange.values || [];
            const rosterDataRows = rosterValues.slice(1);
            const cleanDataRows = cleanValues.slice(1);
            
            // ============================================================================
            // DETERMINE JOIN KEY - Check if BOTH datasets have Employee_ID with VALUES
            // ============================================================================
            const rosterIdHasValues = rosterIdIdx >= 0 && columnHasValues(rosterDataRows, rosterIdIdx);
            const cleanIdHasValues = cleanIdIdx >= 0 && columnHasValues(cleanDataRows, cleanIdIdx);
            
            // Only use Employee_ID if BOTH datasets have the column AND have actual values
            const useEmployeeId = rosterIdHasValues && cleanIdHasValues;
            const effectiveJoinKey = useEmployeeId ? "Employee_ID" : "Employee_Name";
            const rosterKeyIdx = useEmployeeId ? rosterIdIdx : rosterNameIdx;
            const cleanKeyIdx = useEmployeeId ? cleanIdIdx : cleanNameIdx;
            
            console.log(`[PayrollCoverage] Join key determination:`);
            console.log(`  - Roster has Employee_ID column: ${rosterIdIdx >= 0}, with values: ${rosterIdHasValues}`);
            console.log(`  - Payroll has Employee_ID column: ${cleanIdIdx >= 0}, with values: ${cleanIdHasValues}`);
            console.log(`  - Using join key: ${effectiveJoinKey}`);
            
            // Build employee maps using the effective join key
            const rosterEmployees = new Map();
            const cleanEmployees = new Map();
            
            // Track terminated count for reporting
            let terminatedCount = 0;
            
            // Parse roster (skip header row) - only include active employees
            for (let i = 1; i < rosterValues.length; i++) {
                const row = rosterValues[i];
                const keyValue = row[rosterKeyIdx];
                const key = normalizeJoinKey(keyValue);
                
                if (!key || isNoiseName(key)) continue;
                
                // Skip terminated employees - they shouldn't be expected in payroll
                if (rosterStatusIdx >= 0) {
                    const status = String(row[rosterStatusIdx] || "").toLowerCase().trim();
                    if (status === "terminated" || status === "inactive" || status === "term") {
                        terminatedCount++;
                        continue;
                    }
                }
                
                if (!rosterEmployees.has(key)) {
                    const dept = rosterDeptIdx >= 0 ? normalizeJoinKey(row[rosterDeptIdx]) : "";
                    rosterEmployees.set(key, {
                        name: normalizeString(keyValue) || key,
                        department: dept
                    });
                }
            }
            
            // Parse PR_Data_Clean (skip header row)
            for (let i = 1; i < cleanValues.length; i++) {
                const row = cleanValues[i];
                const keyValue = row[cleanKeyIdx];
                const key = normalizeJoinKey(keyValue);
                
                if (!key || isNoiseName(key)) continue;
                
                if (!cleanEmployees.has(key)) {
                    const dept = cleanDeptIdx >= 0 ? normalizeJoinKey(row[cleanDeptIdx]) : "";
                    cleanEmployees.set(key, {
                        name: normalizeString(keyValue) || key,
                        department: dept
                    });
                }
            }
            
            if (rosterEmployees.size === 0) {
                return { error: `No employees found in roster using ${effectiveJoinKey}.`, status: "unavailable" };
            }
            
            if (cleanEmployees.size === 0) {
                return { error: `No employees found in PR_Data_Clean using ${effectiveJoinKey}. Check column mapping.`, status: "unavailable" };
            }
            
            // Compute employee coverage
            const missingFromPayroll = [];
            const extraInPayroll = [];
            
            // Employees in roster but NOT in payroll
            rosterEmployees.forEach((entry, key) => {
                if (!cleanEmployees.has(key)) {
                    missingFromPayroll.push({
                        name: entry.name,
                        department: entry.department || "—"
                    });
                }
            });
            
            // Employees in payroll but NOT in roster
            cleanEmployees.forEach((entry, key) => {
                if (!rosterEmployees.has(key)) {
                    extraInPayroll.push({
                        name: entry.name,
                        department: entry.department || "—"
                    });
                }
            });
            
            // Compute department coverage (normalize for comparison)
            const rosterDepts = new Set();
            const payrollDepts = new Set();
            
            rosterEmployees.forEach(entry => {
                if (entry.department) rosterDepts.add(entry.department);
            });
            
            cleanEmployees.forEach(entry => {
                if (entry.department) payrollDepts.add(entry.department);
            });
            
            const missingDeptsFromPayroll = [...rosterDepts].filter(d => !payrollDepts.has(d));
            const extraDeptsInPayroll = [...payrollDepts].filter(d => !rosterDepts.has(d));
            
            return {
                hasData: true,
                joinKeyUsed: effectiveJoinKey,
                employee: {
                    rosterCount: rosterEmployees.size,  // Active employees only
                    terminatedCount: terminatedCount,    // Excluded from comparison
                    payrollCount: cleanEmployees.size,
                    missingFromPayroll,
                    extraInPayroll,
                    status: (missingFromPayroll.length > 0 || extraInPayroll.length > 0) ? "review" : "ok"
                },
                department: {
                    rosterDepts: [...rosterDepts],
                    payrollDepts: [...payrollDepts],
                    missingFromPayroll: missingDeptsFromPayroll,
                    extraInPayroll: extraDeptsInPayroll,
                    status: (missingDeptsFromPayroll.length > 0 || extraDeptsInPayroll.length > 0) ? "review" : "ok"
                }
            };
        });
        
        if (result.error) {
            payrollCoverageState.lastError = result.error;
            payrollCoverageState.hasData = false;
            payrollCoverageState.joinKeyUsed = null;
            payrollCoverageState.employee.status = result.status || "unavailable";
            payrollCoverageState.department.status = result.status || "unavailable";
        } else {
            payrollCoverageState.hasData = true;
            payrollCoverageState.joinKeyUsed = result.joinKeyUsed;
            payrollCoverageState.employee = result.employee;
            payrollCoverageState.department = result.department;
            payrollCoverageState.lastError = null;
            
            console.log("[PayrollCoverage] Analysis complete:", {
                joinKey: result.joinKeyUsed,
                rosterEmployees: result.employee.rosterCount,
                payrollEmployees: result.employee.payrollCount,
                missingFromPayroll: result.employee.missingFromPayroll.length,
                extraInPayroll: result.employee.extraInPayroll.length,
                deptMismatches: result.department.missingFromPayroll.length + result.department.extraInPayroll.length
            });
        }
        
    } catch (error) {
        console.error("[PayrollCoverage] Error:", error);
        payrollCoverageState.lastError = `Coverage check failed: ${error.message}`;
        payrollCoverageState.hasData = false;
        payrollCoverageState.joinKeyUsed = null;
        payrollCoverageState.employee.status = "unavailable";
        payrollCoverageState.department.status = "unavailable";
    } finally {
        payrollCoverageState.loading = false;
    }
}

/**
 * Parse PR_Data_Clean values for coverage comparison
 * @deprecated - Now handled inline in refreshPayrollCoverage for better performance
 */
function parsePRDataCleanValues(values) {
    const result = {
        totalEmployees: 0,
        employeeMap: new Map()
    };
    
    if (!values || values.length < 2) return result;
    
    // First row is headers
    const headers = values[0].map(h => normalizeHeader(String(h || "")));
    
    // Use centralized join key finder
    const joinKeyInfo = findEmployeeJoinKeyColumn(headers);
    const actualEmployeeIdx = joinKeyInfo.index;
    const departmentIdx = pickDepartmentIndex(headers);
    
    if (actualEmployeeIdx === -1) {
        console.warn("[PayrollCoverage] No employee column found in PR_Data_Clean. Headers:", headers);
        return result;
    }
    
    const employeeSet = new Set();
    
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const employee = row[actualEmployeeIdx];
        const key = normalizeJoinKey(employee);
        
        if (!key || isNoiseName(key)) continue;
        
        if (!employeeSet.has(key)) {
            employeeSet.add(key);
            result.totalEmployees++;
            
            const department = departmentIdx >= 0 ? row[departmentIdx] : "";
            result.employeeMap.set(key, {
                name: normalizeString(employee) || key,
                department: normalizeJoinKey(department)
            });
        }
    }
    
    return result;
}

/**
 * Render Employee Coverage as an expandable bar
 * Uses rosterUpdateState as the single source of truth
 * Payroll = source of truth, Roster = local copy to keep in sync
 */
function renderEmployeeCoverageCard() {
    const state = rosterUpdateState;
    
    // Determine status
    let statusText = "";
    let statusClass = "";
    let statusBadge = "";
    
    if (state.loading) {
        statusText = "Checking...";
        statusClass = "pf-expand-bar--muted";
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status"><span>Loading</span></span>`;
    } else if (state.lastError) {
        statusText = "Unavailable";
        statusClass = "pf-expand-bar--muted";
        statusBadge = `<span class="pf-status-badge pf-status-badge--unavailable" role="status"><span>Unavailable</span></span>`;
    } else if (!state.hasData) {
        statusText = "Run Create Matrix";
        statusClass = "pf-expand-bar--muted";
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status"><span>Pending</span></span>`;
    } else {
        const newCount = (state.newHires || []).length;
        const termCount = (state.missingEmployees || []).length;
        const reactCount = (state.reactivations || []).length;
        const totalChanges = newCount + termCount + reactCount;
        
        if (totalChanges === 0) {
            statusText = "Roster in sync";
            statusClass = "pf-expand-bar--success";
            statusBadge = `<span class="pf-status-badge pf-status-badge--ok" role="status"><span>OK</span></span>`;
        } else {
            const parts = [];
            if (newCount > 0) parts.push(`+${newCount} new`);
            if (termCount > 0) parts.push(`−${termCount} termed`);
            if (reactCount > 0) parts.push(`${reactCount} reactivated`);
            statusText = parts.join(", ");
            statusClass = "pf-expand-bar--warning";
            statusBadge = `<span class="pf-status-badge pf-status-badge--review" role="status"><span>Review</span></span>`;
        }
    }
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" id="employee-coverage-card">
            <div class="pf-config-head">
                <h3>Employee Roster ${statusBadge}</h3>
                <p class="pf-config-subtext">Compare payroll employees to local roster. Click to review changes.</p>
            </div>
            <button type="button" class="pf-expand-bar ${statusClass} pf-clickable" id="employee-coverage-bar">
                <svg class="pf-expand-bar-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                    <polyline points="6 9 12 15 18 9"/>
                </svg>
                <span class="pf-expand-bar-label">Reconciliation</span>
                <span class="pf-expand-bar-count">${statusText}</span>
            </button>
        </article>
    `;
}

/**
 * Open the Employee Roster reconciliation modal
 * Shows payroll as source of truth, with changes needed to sync roster
 */
function openEmployeeCoverageModal() {
    const state = rosterUpdateState;
    
    // Remove existing modal
    const existing = document.getElementById("coverage-detail-modal");
    if (existing) existing.remove();
    
    // Build modal content
    let bodyContent = "";
    
    if (state.loading) {
        bodyContent = `<p class="pf-subsection-hint">Analyzing coverage...</p>`;
    } else if (state.lastError) {
        bodyContent = `<p class="pf-subsection-hint pf-subsection-hint--warn">${escapeHtml(state.lastError)}</p>`;
    } else if (!state.hasData) {
        bodyContent = `<p class="pf-subsection-hint">Run 'Create Matrix' first to compare payroll vs roster.</p>`;
    } else {
        bodyContent = renderRosterReconciliation();
    }
    
    const modal = document.createElement("div");
    modal.id = "coverage-detail-modal";
    modal.className = "pf-coverage-modal";
    modal.innerHTML = `
        <div class="pf-coverage-modal-backdrop" data-close></div>
        <div class="pf-coverage-modal-card">
            <div class="pf-coverage-modal-header">
                <h3 class="pf-coverage-modal-title">Employee Roster Reconciliation</h3>
                <button class="pf-coverage-modal-close pf-clickable" type="button" aria-label="Close" data-close>
                    ${X_ICON_SVG}
                </button>
            </div>
            <div class="pf-coverage-modal-body">
                ${bodyContent}
            </div>
            <div class="pf-modal-footer">
                ${state.hasData && ((state.newHires || []).length > 0 || (state.missingEmployees || []).length > 0 || (state.reactivations || []).length > 0)
                    ? `<button type="button" class="pf-pill-btn pf-pill-btn--primary" id="modal-roster-apply-btn">Push Selected to Roster</button>`
                    : `<button type="button" class="pf-pill-btn" data-close>Close</button>`}
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Bind close handlers
    modal.querySelectorAll("[data-close]").forEach(el => {
        el.addEventListener("click", () => modal.remove());
    });

    // Track checkbox selection into overrides
    modal.querySelectorAll(".pf-recon-checkbox").forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
            const action = checkbox.dataset.action;
            const key = checkbox.dataset.key;
            if (!action || !key) return;

            // Checked = take default suggested action, unless manual-managed reactivation (requires explicit override)
            if (checkbox.checked) {
                if (action === "reactivate") {
                    const emp = (rosterUpdateState.reactivations || []).find(r => String(r.key || r.name) === String(key));
                    if (emp?.isManuallyManaged) {
                        rosterUpdateState.overrides.set(key, { action: "reactivate" });
                    } else {
                        rosterUpdateState.overrides.delete(key);
                    }
                } else {
                    rosterUpdateState.overrides.delete(key);
                }
                return;
            }

            // Unchecked = opt out of the suggested action
            if (action === "new") {
                rosterUpdateState.overrides.set(key, { action: "ignore" });
            } else if (action === "terminate") {
                rosterUpdateState.overrides.set(key, { action: "keep_active" });
            } else if (action === "reactivate") {
                rosterUpdateState.overrides.set(key, { action: "keep_terminated" });
            }
        });
    });

    // Bind apply button
    const applyBtn = modal.querySelector("#modal-roster-apply-btn");
    if (applyBtn) {
        applyBtn.addEventListener("click", async () => {
            await applyRosterUpdates();
            modal.remove();
        });
    }
    
    // Close on escape
    const escHandler = (e) => {
        if (e.key === "Escape") {
            modal.remove();
            document.removeEventListener("keydown", escHandler);
        }
    };
    document.addEventListener("keydown", escHandler);
}

/**
 * Render the roster reconciliation content
 * Payroll = source of truth, shows what changes are needed to sync roster
 */
function renderRosterReconciliation() {
    const state = rosterUpdateState;
    const newHires = state.newHires || [];
    const missingEmployees = state.missingEmployees || [];
    const reactivations = state.reactivations || [];
    const stillActive = state.stillActive || 0;
    
    // Calculate totals
    const payrollCount = stillActive + newHires.length + reactivations.length;
    const rosterActiveCount = stillActive + missingEmployees.length;
    const netChange = newHires.length + reactivations.length - missingEmployees.length;
    
    let html = `<div class="pf-roster-reconciliation">`;
    
    // Top line: Payroll count (source of truth)
    html += `
        <div class="pf-recon-row pf-recon-row--header">
            <span class="pf-recon-label">Employees per Payroll Report</span>
            <span class="pf-recon-value">${payrollCount}</span>
        </div>
    `;
    
    // Breakdown: Changes needed
    if (newHires.length > 0 || missingEmployees.length > 0 || reactivations.length > 0) {
        html += `<div class="pf-recon-changes">`;
        
        // New hires (add to roster)
        newHires.forEach((hire, i) => {
            const key = hire.key || hire.name;
            html += `
                <div class="pf-recon-row pf-recon-row--add">
                    <label class="pf-recon-checkbox-label">
                        <input type="checkbox" class="pf-recon-checkbox" data-action="new" data-key="${escapeHtml(key)}" checked>
                        <span class="pf-recon-name">New: ${escapeHtml(hire.name || key)}</span>
                    </label>
                    <span class="pf-recon-value pf-recon-value--add">+1</span>
                </div>
            `;
        });
        
        // Reactivations
        reactivations.forEach((emp) => {
            const key = emp.key || emp.name;
            html += `
                <div class="pf-recon-row pf-recon-row--add">
                    <label class="pf-recon-checkbox-label">
                        <input type="checkbox" class="pf-recon-checkbox" data-action="reactivate" data-key="${escapeHtml(key)}" checked>
                        <span class="pf-recon-name">Reactivate: ${escapeHtml(emp.name || key)}</span>
                    </label>
                    <span class="pf-recon-value pf-recon-value--add">+1</span>
                </div>
            `;
        });
        
        // Missing from payroll (potential terminations)
        missingEmployees.forEach((emp) => {
            const key = emp.key || emp.name;
            html += `
                <div class="pf-recon-row pf-recon-row--subtract">
                    <label class="pf-recon-checkbox-label">
                        <input type="checkbox" class="pf-recon-checkbox" data-action="terminate" data-key="${escapeHtml(key)}" checked>
                        <span class="pf-recon-name">Termed: ${escapeHtml(emp.name || key)}</span>
                    </label>
                    <span class="pf-recon-value pf-recon-value--subtract">−1</span>
                </div>
            `;
        });
        
        html += `</div>`;
    } else {
        html += `
            <div class="pf-recon-row pf-recon-row--ok">
                <span class="pf-recon-label">No changes needed</span>
                <span class="pf-recon-value">—</span>
            </div>
        `;
    }
    
    // Bottom line: Roster total (should tie to SS_Employee_Roster active count)
    const expectedRosterTotal = stillActive + newHires.length + reactivations.length;
    html += `
        <div class="pf-recon-row pf-recon-row--footer">
            <span class="pf-recon-label">Total per Roster (after updates)</span>
            <span class="pf-recon-value">${expectedRosterTotal}</span>
        </div>
    `;
    
    // Matching info
    const joinKeyLabel = state.joinKeyUsed === "Employee_ID" ? "Employee ID" : "Employee Name";
    html += `
        <p class="pf-recon-hint">
            Matching by: <strong>${joinKeyLabel}</strong>
        </p>
    `;
    
    html += `</div>`;
    
    return html;
}

/**
 * Show coverage detail modal for a specific category
 */
function showCoverageDetailModal(category) {
    const coverageState = payrollCoverageState;
    const rosterState = rosterUpdateState;
    
    // Get data for the category
    const rosterActive = rosterState.stillActive || coverageState.employee?.rosterCount || 0;
    const payrollCount = coverageState.employee?.payrollCount || 0;
    const newHires = rosterState.newHires || [];
    const missingEmployees = rosterState.missingEmployees || [];
    const reactivations = rosterState.reactivations || [];
    const notInPayroll = coverageState.employee?.missingFromPayroll || [];
    const notInRoster = coverageState.employee?.extraInPayroll || [];
    
    let title = "";
    let content = "";
    
    switch (category) {
        case "summary":
            title = "Coverage Summary";
            content = renderCoverageSummaryDetail(rosterActive, payrollCount, newHires, missingEmployees, reactivations, notInPayroll, notInRoster);
            break;
        case "not-in-payroll":
            title = `${notInPayroll.length} Active Employees Not in Payroll`;
            content = renderEmployeeListDetail(notInPayroll, "These employees are marked active in the roster but were not found in the current payroll data.", true);
            break;
        case "reactivated":
            title = `${reactivations.length} Reactivated Employees`;
            content = renderEmployeeListDetail(reactivations, "These employees were previously terminated but have appeared in the current payroll. They may need their roster status updated.", true);
            break;
        case "missing":
            title = `${missingEmployees.length} Missing from Payroll`;
            content = renderEmployeeListDetail(missingEmployees, "These active roster employees were not seen in the current payroll period. This may indicate terminations, leave, or data issues.", true);
            break;
        case "new":
            title = `${newHires.length} New Employees`;
            content = renderEmployeeListDetail(newHires, "These employees appeared in payroll but are not yet in the roster. They will be added when you push updates.", true);
            break;
        case "payroll-only":
            title = `${notInRoster.length} In Payroll Only`;
            content = renderEmployeeListDetail(notInRoster, "These employees are in the payroll data but not found in the roster. Review to determine if they should be added.", true);
            break;
        default:
            return;
    }
    
    // Create and show modal
    showCoverageModal(title, content);
}

/**
 * Render the summary detail with headcount bridge calculation
 */
function renderCoverageSummaryDetail(rosterActive, payrollCount, newHires, missingEmployees, reactivations, notInPayroll, notInRoster) {
    const difference = payrollCount - rosterActive;
    const differenceText = difference === 0 ? "Match" : (difference > 0 ? `+${difference}` : `${difference}`);
    const differenceClass = difference === 0 ? "match" : "mismatch";
    
    return `
        <div class="pf-coverage-detail-summary">
            <h4>Headcount Bridge</h4>
            <div class="pf-headcount-bridge">
                <div class="pf-bridge-row pf-bridge-row--base">
                    <span class="pf-bridge-label">Active in Roster</span>
                    <span class="pf-bridge-value">${rosterActive}</span>
                </div>
                ${newHires.length > 0 ? `
                <div class="pf-bridge-row pf-bridge-row--add">
                    <span class="pf-bridge-label">+ New employees (in payroll, not in roster)</span>
                    <span class="pf-bridge-value">+${newHires.length}</span>
                </div>
                ` : ""}
                ${reactivations.length > 0 ? `
                <div class="pf-bridge-row pf-bridge-row--add">
                    <span class="pf-bridge-label">+ Reactivated (terminated but in payroll)</span>
                    <span class="pf-bridge-value">+${reactivations.length}</span>
                </div>
                ` : ""}
                ${missingEmployees.length > 0 ? `
                <div class="pf-bridge-row pf-bridge-row--subtract">
                    <span class="pf-bridge-label">− Missing (active but not in payroll)</span>
                    <span class="pf-bridge-value">−${missingEmployees.length}</span>
                </div>
                ` : ""}
                ${notInPayroll.length > 0 ? `
                <div class="pf-bridge-row pf-bridge-row--subtract">
                    <span class="pf-bridge-label">− Not in payroll (active roster, no match)</span>
                    <span class="pf-bridge-value">−${notInPayroll.length}</span>
                </div>
                ` : ""}
                <div class="pf-bridge-row pf-bridge-row--total">
                    <span class="pf-bridge-label">In Payroll</span>
                    <span class="pf-bridge-value">${payrollCount}</span>
                </div>
                <div class="pf-bridge-row pf-bridge-row--diff pf-bridge-row--${differenceClass}">
                    <span class="pf-bridge-label">Difference</span>
                    <span class="pf-bridge-value">${differenceText}</span>
                </div>
            </div>
            
            <h4 style="margin-top: 20px;">What to Review</h4>
            <ul class="pf-coverage-review-list">
                ${notInPayroll.length > 0 ? `<li><strong>${notInPayroll.length} not in payroll</strong> — Verify these active employees should have payroll entries</li>` : ""}
                ${reactivations.length > 0 ? `<li><strong>${reactivations.length} reactivated</strong> — Update roster status from Terminated to Active</li>` : ""}
                ${missingEmployees.length > 0 ? `<li><strong>${missingEmployees.length} missing</strong> — Check if these are terminations, leaves, or data issues</li>` : ""}
                ${newHires.length > 0 ? `<li><strong>${newHires.length} new</strong> — Add to roster via "Push Updates"</li>` : ""}
                ${notInRoster.length > 0 ? `<li><strong>${notInRoster.length} payroll only</strong> — Review if these should be added to roster</li>` : ""}
                ${notInPayroll.length === 0 && reactivations.length === 0 && missingEmployees.length === 0 && newHires.length === 0 && notInRoster.length === 0 ? `<li>✓ No issues found — roster and payroll are in sync</li>` : ""}
            </ul>
        </div>
    `;
}

/**
 * Render a list of employees with optional department
 */
function renderEmployeeListDetail(employees, description, showDepartment = false) {
    if (!employees || employees.length === 0) {
        return `<p>${description}</p><p><em>No employees in this category.</em></p>`;
    }
    
    let html = `<p style="margin-bottom: 16px;">${description}</p>`;
    html += `<div class="pf-employee-list">`;
    
    employees.forEach((emp, index) => {
        const name = emp.name || emp;
        const dept = showDepartment && emp.department ? emp.department : null;
        html += `
            <div class="pf-employee-list-item">
                <span class="pf-employee-index">${index + 1}.</span>
                <span class="pf-employee-name">${escapeHtml(name)}</span>
                ${dept ? `<span class="pf-employee-dept">${escapeHtml(dept)}</span>` : ""}
            </div>
        `;
    });
    
    html += `</div>`;
    return html;
}

/**
 * Show coverage detail modal
 */
function showCoverageModal(title, content) {
    // Remove existing modal if present
    const existing = document.getElementById("coverage-detail-modal");
    if (existing) existing.remove();
    
    const modal = document.createElement("div");
    modal.id = "coverage-detail-modal";
    modal.className = "pf-coverage-modal";
    modal.innerHTML = `
        <div class="pf-coverage-modal-backdrop" data-close></div>
        <div class="pf-coverage-modal-card">
            <div class="pf-coverage-modal-header">
                <h3 class="pf-coverage-modal-title">${escapeHtml(title)}</h3>
                <button class="pf-coverage-modal-close pf-clickable" type="button" aria-label="Close" data-close>
                    ${X_ICON_SVG}
                </button>
            </div>
            <div class="pf-coverage-modal-body">
                ${content}
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Bind close handlers
    modal.querySelectorAll("[data-close]").forEach(el => {
        el.addEventListener("click", () => modal.remove());
    });
    
    // Close on escape key
    const escHandler = (e) => {
        if (e.key === "Escape") {
            modal.remove();
            document.removeEventListener("keydown", escHandler);
        }
    };
    document.addEventListener("keydown", escHandler);
}

/**
 * Render the coverage check subsection content
 * @deprecated Use renderUnifiedCoverageContent() instead
 */
function renderCoverageSubsection() {
    const state = payrollCoverageState;
    
    if (state.loading) {
        return `<p class="pf-subsection-hint">Checking coverage...</p>`;
    }
    
    if (state.lastError) {
        return `<p class="pf-subsection-hint pf-subsection-hint--warn">${escapeHtml(state.lastError)}</p>`;
    }
    
    if (!state.hasData) {
        return `<p class="pf-subsection-hint">Run 'Create Matrix' to check coverage.</p>`;
    }
    
    const empMissing = state.employee.missingFromPayroll;
    const empExtra = state.employee.extraInPayroll;
    
    let html = `
        <div class="pf-coverage-stats">
            <span class="pf-coverage-stat">Roster: <strong>${state.employee.rosterCount}</strong></span>
            <span class="pf-coverage-stat">Payroll: <strong>${state.employee.payrollCount}</strong></span>
        </div>
    `;
    
    if (empMissing.length === 0 && empExtra.length === 0) {
        html += `
            <div class="pf-coverage-item pf-coverage-item--ok">
                <span class="pf-coverage-item-icon">✓</span>
                <span>All roster employees covered</span>
            </div>
        `;
    } else {
        if (empMissing.length > 0) {
            const displayNames = empMissing.slice(0, 3).map(e => e.name);
            const moreCount = empMissing.length - 3;
            html += `
                <div class="pf-coverage-item pf-coverage-item--warn">
                    <span class="pf-coverage-item-icon">⚠</span>
                    <div>
                        <strong>${empMissing.length} not in payroll</strong>
                        <div class="pf-coverage-item-names">${displayNames.join(", ")}${moreCount > 0 ? ` +${moreCount}` : ""}</div>
                    </div>
                </div>
            `;
        }
        
        if (empExtra.length > 0) {
            const displayNames = empExtra.slice(0, 3).map(e => e.name);
            const moreCount = empExtra.length - 3;
            html += `
                <div class="pf-coverage-item pf-coverage-item--info">
                    <span class="pf-coverage-item-icon">ℹ</span>
                    <div>
                        <strong>${empExtra.length} not in roster</strong>
                        <div class="pf-coverage-item-names">${displayNames.join(", ")}${moreCount > 0 ? ` +${moreCount}` : ""}</div>
                    </div>
                </div>
            `;
        }
    }
    
    return html;
}

/**
 * Render the roster updates subsection content
 */
function renderRosterUpdatesSubsection() {
    const state = rosterUpdateState;
    
    if (state.loading) {
        return `<p class="pf-subsection-hint">Analyzing changes...</p>`;
    }
    
    if (state.lastError) {
        return `<p class="pf-subsection-hint pf-subsection-hint--warn">${escapeHtml(state.lastError)}</p>`;
    }
    
    if (!state.hasData) {
        return `<p class="pf-subsection-hint">Run 'Create Matrix' to analyze roster.</p>`;
    }
    
    const { newHires, missingEmployees, reactivations, stillActive } = state;
    
    let html = `
        <div class="pf-coverage-stats">
            <span class="pf-coverage-stat" style="color: rgba(255,255,255,0.7);"><strong>${stillActive}</strong> active</span>
            <span class="pf-coverage-stat" style="color: #4ade80;"><strong>${newHires.length}</strong> new</span>
            ${reactivations.length > 0 ? `<span class="pf-coverage-stat" style="color: #60a5fa;"><strong>${reactivations.length}</strong> reactivated</span>` : ""}
            <span class="pf-coverage-stat" style="color: #fbbf24;"><strong>${missingEmployees.length}</strong> missing</span>
        </div>
    `;
    
    const hasChanges = newHires.length > 0 || missingEmployees.length > 0 || reactivations.length > 0;
    
    if (!hasChanges) {
        html += `
            <div class="pf-coverage-item pf-coverage-item--ok">
                <span class="pf-coverage-item-icon">✓</span>
                <span>Roster is up to date</span>
            </div>
        `;
    } else {
        if (newHires.length > 0) {
            const names = newHires.slice(0, 3).map(h => h.name);
            const more = newHires.length - 3;
            html += `
                <div class="pf-coverage-item pf-coverage-item--new">
                    <span class="pf-coverage-item-icon">+</span>
                    <div>
                        <strong>New:</strong> ${names.join(", ")}${more > 0 ? ` +${more}` : ""}
                    </div>
                </div>
            `;
        }
        
        if (reactivations.length > 0) {
            const names = reactivations.slice(0, 3).map(r => r.name);
            const more = reactivations.length - 3;
            html += `
                <div class="pf-coverage-item pf-coverage-item--reactivated">
                    <span class="pf-coverage-item-icon" aria-hidden="true">${ALERT_TRIANGLE_SVG}</span>
                    <div>
                        <strong>Reactivated:</strong> ${names.join(", ")}${more > 0 ? ` +${more}` : ""}
                    </div>
                </div>
            `;
        }
        
        if (missingEmployees.length > 0) {
            const names = missingEmployees.slice(0, 3).map(m => m.name);
            const more = missingEmployees.length - 3;
            html += `
                <div class="pf-coverage-item pf-coverage-item--missing">
                    <span class="pf-coverage-item-icon" aria-hidden="true">${ALERT_TRIANGLE_SVG}</span>
                    <div>
                        <strong>Missing:</strong> ${names.join(", ")}${more > 0 ? ` +${more}` : ""}
                    </div>
                </div>
            `;
        }
        
        html += `
            <div class="pf-roster-actions">
                <button type="button" class="pf-pill-btn pf-pill-btn--sm" id="roster-apply-btn" ${state.applyPending ? "disabled" : ""}>
                    Apply Updates
                </button>
                <button type="button" class="pf-action-toggle pf-action-toggle--subtle" id="roster-refresh-btn" title="Refresh">
                    ${REFRESH_ICON_SVG}
                </button>
            </div>
        `;
    }
    
    return html;
}

/**
 * Render the Payroll Coverage card for Step 1 Validation section
 * Status badges include text labels for accessibility (not color-only)
 * @deprecated Use renderEmployeeCoverageCard() instead - kept for reference
 */
function renderPayrollCoverageCard() {
    const state = payrollCoverageState;
    
    // Determine overall status with accessible text labels
    let statusBadge = "";
    let statusClass = "";
    let statusAriaLabel = "";
    
    if (state.loading) {
        // Pending = validation is running
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status" aria-label="Loading coverage check">${INFO_CIRCLE_SVG}<span>Loading</span></span>`;
        statusAriaLabel = "Coverage check in progress";
    } else if (state.lastError) {
        // Unavailable = cannot run due to missing prerequisites
        statusBadge = `<span class="pf-status-badge pf-status-badge--unavailable" role="status" aria-label="Coverage check unavailable">${ALERT_TRIANGLE_SVG}<span>Unavailable</span></span>`;
        statusClass = "pf-coverage-unavailable";
        statusAriaLabel = "Coverage check cannot run: " + state.lastError;
    } else if (!state.hasData) {
        // Pending = can run but hasn't completed
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status" aria-label="Coverage check pending"><span>Pending</span></span>`;
        statusAriaLabel = "Coverage check has not run yet";
    } else if (state.employee.status === "ok" && state.department.status === "ok") {
        statusBadge = `<span class="pf-status-badge pf-status-badge--ok" role="status" aria-label="Coverage check passed">${CHECK_CIRCLE_SVG}<span>OK</span></span>`;
        statusClass = "pf-coverage-ok";
        statusAriaLabel = "All employees covered";
    } else {
        statusBadge = `<span class="pf-status-badge pf-status-badge--review" role="status" aria-label="Coverage check needs review">${ALERT_TRIANGLE_SVG}<span>Review</span></span>`;
        statusClass = "pf-coverage-review";
        statusAriaLabel = "Coverage differences found - review recommended";
    }
    
    // Build employee coverage section
    let employeeSectionHtml = "";
    if (state.hasData) {
        const empMissing = state.employee.missingFromPayroll;
        const empExtra = state.employee.extraInPayroll;
        
        if (empMissing.length === 0 && empExtra.length === 0) {
            employeeSectionHtml = `
                <div class="pf-coverage-check pf-coverage-check--pass">
                    <span class="pf-coverage-icon" aria-hidden="true">${CHECK_CIRCLE_SVG}</span>
                    <span>Payroll covers all ${state.employee.rosterCount} roster employees.</span>
                </div>
            `;
        } else {
            let items = [];
            
            if (empMissing.length > 0) {
                const displayNames = empMissing.slice(0, COVERAGE_MAX_DISPLAY_ITEMS).map(e => e.name);
                const moreCount = empMissing.length - COVERAGE_MAX_DISPLAY_ITEMS;
                items.push(`
                    <div class="pf-coverage-check pf-coverage-check--warn">
                        <span class="pf-coverage-icon" aria-hidden="true">${ALERT_TRIANGLE_SVG}</span>
                        <div>
                            <strong>${empMissing.length} roster employee${empMissing.length > 1 ? 's' : ''} not in payroll</strong>
                            <span class="pf-coverage-hint">(possible incomplete coding / accrual risk)</span>
                            <div class="pf-coverage-names">${displayNames.join(", ")}${moreCount > 0 ? ` +${moreCount} more` : ""}</div>
                        </div>
                    </div>
                `);
            }
            
            if (empExtra.length > 0) {
                const displayNames = empExtra.slice(0, COVERAGE_MAX_DISPLAY_ITEMS).map(e => e.name);
                const moreCount = empExtra.length - COVERAGE_MAX_DISPLAY_ITEMS;
                items.push(`
                    <div class="pf-coverage-check pf-coverage-check--info">
                        <span class="pf-coverage-icon" aria-hidden="true">${INFO_CIRCLE_SVG}</span>
                        <div>
                            <strong>${empExtra.length} payroll employee${empExtra.length > 1 ? 's' : ''} not in roster</strong>
                            <span class="pf-coverage-hint">(may indicate new hires or contractors)</span>
                            <div class="pf-coverage-names">${displayNames.join(", ")}${moreCount > 0 ? ` +${moreCount} more` : ""}</div>
                        </div>
                    </div>
                `);
            }
            
            employeeSectionHtml = items.join("");
        }
    } else if (state.lastError) {
        employeeSectionHtml = `
            <div class="pf-coverage-check pf-coverage-check--error">
                <span class="pf-coverage-icon">!</span>
                <span>${escapeHtml(state.lastError)}</span>
            </div>
        `;
    } else {
        employeeSectionHtml = `
            <div class="pf-coverage-check pf-coverage-check--pending">
                <span class="pf-coverage-icon" aria-hidden="true">${INFO_CIRCLE_SVG}</span>
                <span>Run 'Create Matrix' to enable coverage checks.</span>
            </div>
        `;
    }
    
    // Build department coverage section (secondary, only show if issues exist)
    let departmentSectionHtml = "";
    if (state.hasData) {
        const deptMissing = state.department.missingFromPayroll;
        const deptExtra = state.department.extraInPayroll;
        
        if (deptMissing.length > 0 || deptExtra.length > 0) {
            let deptItems = [];
            
            if (deptMissing.length > 0) {
                deptItems.push(`
                    <div class="pf-coverage-dept-item">
                        <span>Roster depts not in payroll:</span>
                        <span class="pf-coverage-dept-list">${deptMissing.join(", ")}</span>
                    </div>
                `);
            }
            
            if (deptExtra.length > 0) {
                deptItems.push(`
                    <div class="pf-coverage-dept-item">
                        <span>Payroll depts not in roster:</span>
                        <span class="pf-coverage-dept-list">${deptExtra.join(", ")}</span>
                    </div>
                `);
            }
            
            departmentSectionHtml = `
                <div class="pf-coverage-dept-section">
                    <div class="pf-coverage-dept-header">
                        <span class="pf-coverage-icon" aria-hidden="true">${GRID_ICON_SVG}</span>
                        <strong>Department differences</strong>
                        <span class="pf-coverage-hint">(may indicate miscoding)</span>
                    </div>
                    ${deptItems.join("")}
                </div>
            `;
        }
    }
    
    // Refresh button (subtle when data is fresh)
    const refreshButton = `
        <button type="button" class="pf-action-toggle pf-action-toggle--subtle pf-clickable" id="coverage-refresh-btn" title="Refresh coverage check" style="margin-left: auto;">
            ${REFRESH_ICON_SVG}
        </button>
    `;
    
    // Show which join key is being used (transparency for users)
    const joinKeyInfo = state.joinKeyUsed 
        ? `<span class="pf-coverage-join-key" title="Matching employees using ${state.joinKeyUsed}">Matching by: ${state.joinKeyUsed.replace("_", " ")}</span>` 
        : "";
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card ${statusClass}" id="coverage-card" aria-label="${statusAriaLabel}">
            <div class="pf-config-head" style="display: flex; align-items: center;">
                <div>
                    <h3>Payroll Coverage ${statusBadge}</h3>
                    <p class="pf-config-subtext">Compare roster employees vs payroll data (advisory). ${joinKeyInfo}</p>
                </div>
                ${refreshButton}
            </div>
            <div class="pf-coverage-content">
                <div class="pf-coverage-summary">
                    ${state.hasData ? `
                        <span class="pf-coverage-stat">Roster: ${state.employee.rosterCount}</span>
                        <span class="pf-coverage-stat">Payroll: ${state.employee.payrollCount}</span>
                    ` : ""}
                </div>
                ${employeeSectionHtml}
                ${departmentSectionHtml}
            </div>
        </article>
    `;
}

// =============================================================================
// AUTO-MAINTAINED EMPLOYEE ROSTER SERVICE
// SS_Employee_Roster is updated automatically based on PR_Data_Clean
// =============================================================================

/**
 * Standard hours per semi-monthly pay period
 * Semi-monthly = 24 pay periods per year, assuming 1920 work hours/year → 80 hours/period
 */
const HOURS_PER_PAY_PERIOD = 80;

/**
 * Required columns for SS_Employee_Roster (schema evolution)
 * If any column is missing, it will be added automatically
 */
const ROSTER_REQUIRED_COLUMNS = [
    "Employee_ID",
    "Employee_Name",
    "Department_Name",
    "Employment_Status",        // "Active" | "Terminated" | "Unknown"
    "First_Seen_Pay_Date",
    "Last_Seen_Pay_Date",
    "Termination_Effective_Date",
    "Last_Terminated_Date",     // Preserved history when reactivated
    "Hourly_Rate",              // Computed from Fixed bucket / 80 hours
    "Hourly_Rate_Prismhr",      // Rate from PrismHR PTO Liability report (synced by pto-accrual)
    "Rate_Source",              // "Payroll" | "Manual" | description of rate origin
    "Rate_Updated_Date",        // Last time rate was auto-computed
    "Is_Manually_Managed",      // TRUE = don't auto-update this row
    "Notes"
];

const ROSTER_COLUMN_DEFAULTS = {
    Employment_Status: "Unknown",
    Is_Manually_Managed: "FALSE",
    First_Seen_Pay_Date: "",
    Last_Seen_Pay_Date: "",
    Termination_Effective_Date: "",
    Last_Terminated_Date: "",
    Hourly_Rate: "",
    Hourly_Rate_Prismhr: "",
    Rate_Source: "",
    Rate_Updated_Date: "",
    Notes: ""
};

/**
 * Roster update state - tracks detected changes
 */
const rosterUpdateState = {
    loading: false,
    lastError: null,
    hasData: false,
    periodKey: null,
    joinKeyUsed: null,  // "Employee_ID" | "Employee_Name"
    newHires: [],       // [{key, name, department}]
    missingEmployees: [], // [{key, name, department, lastSeenDate}] - candidates for termination
    reactivations: [],  // [{key, name, department, terminationDate, isManuallyManaged}] - terminated employees appearing in payroll
    stillActive: 0,     // count of employees found in both (already Active)
    status: "pending",  // "ok" | "review" | "pending" | "unavailable"
    applyPending: false,
    overrides: new Map() // joinKey -> { action: "keep_active" | "mark_terminated" | "add" | "ignore" | "reactivate" | "keep_terminated" }
};

/**
 * Ensure SS_Employee_Roster has all required columns (schema evolution)
 * Adds missing columns, never deletes existing ones
 */
/**
 * Backfill empty Employee_ID cells with Employee_Name values
 * Run once during roster schema validation
 * Fixes issue where customers matching by name had blank Employee_ID column
 */
async function backfillEmptyEmployeeIds() {
    if (!hasExcelRuntime()) return;
    
    try {
        await Excel.run(async (context) => {
            const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
            rosterSheet.load("isNullObject");
            await context.sync();
            
            if (rosterSheet.isNullObject) return;
            
            const usedRange = rosterSheet.getUsedRangeOrNullObject();
            usedRange.load("values, rowCount, columnCount");
            await context.sync();
            
            if (usedRange.isNullObject || !usedRange.values || usedRange.values.length < 2) return;
            
            const headers = usedRange.values[0].map(h => String(h || "").toLowerCase().trim());
            const idIdx = headers.findIndex(h => h === "employee_id");
            const nameIdx = headers.findIndex(h => h === "employee_name");
            
            if (idIdx < 0 || nameIdx < 0) return;
            
            let backfillCount = 0;
            const values = usedRange.values;
            
            for (let i = 1; i < values.length; i++) {
                const idValue = String(values[i][idIdx] || "").trim();
                const nameValue = String(values[i][nameIdx] || "").trim();
                
                if (!idValue && nameValue) {
                    values[i][idIdx] = nameValue;
                    backfillCount++;
                }
            }
            
            if (backfillCount > 0) {
                usedRange.values = values;
                await context.sync();
                console.log(`[RosterBackfill] Populated ${backfillCount} empty Employee_ID cells with Employee_Name`);
            }
        });
    } catch (error) {
        console.warn("[RosterBackfill] Error:", error.message);
    }
}

async function ensureRosterSchema() {
    if (!hasExcelRuntime()) return { ok: false, error: "Excel runtime unavailable" };
    
    try {
        return await Excel.run(async (context) => {
            let rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
            rosterSheet.load("isNullObject");
            await context.sync();
            
            // Create sheet if missing
            if (rosterSheet.isNullObject) {
                console.log("[RosterSchema] Creating SS_Employee_Roster sheet");
                rosterSheet = context.workbook.worksheets.add("SS_Employee_Roster");
                
                // Write header row with all required columns
                const headerRange = rosterSheet.getRangeByIndexes(0, 0, 1, ROSTER_REQUIRED_COLUMNS.length);
                headerRange.values = [ROSTER_REQUIRED_COLUMNS];
                headerRange.format.font.bold = true;
                await context.sync();
                
                return { ok: true, created: true, addedColumns: ROSTER_REQUIRED_COLUMNS };
            }
            
            // Sheet exists - check for missing columns
            const usedRange = rosterSheet.getUsedRangeOrNullObject();
            usedRange.load("values, columnCount");
            await context.sync();
            
            let existingHeaders = [];
            if (!usedRange.isNullObject && usedRange.values && usedRange.values.length > 0) {
                existingHeaders = usedRange.values[0].map(h => String(h || "").trim());
            }
            
            const existingHeadersLower = new Set(existingHeaders.map(h => h.toLowerCase()));
            const missingColumns = ROSTER_REQUIRED_COLUMNS.filter(col => 
                !existingHeadersLower.has(col.toLowerCase())
            );
            
            if (missingColumns.length > 0) {
                console.log("[RosterSchema] Adding missing columns:", missingColumns);
                
                // Add missing columns at the end
                const startCol = existingHeaders.length;
                const headerRange = rosterSheet.getRangeByIndexes(0, startCol, 1, missingColumns.length);
                headerRange.values = [missingColumns];
                headerRange.format.font.bold = true;
                
                // Backfill default values for existing rows
                const rowCount = usedRange.isNullObject ? 0 : (usedRange.values?.length || 1) - 1;
                if (rowCount > 0) {
                    const defaultValues = missingColumns.map(col => ROSTER_COLUMN_DEFAULTS[col] || "");
                    const fillData = Array(rowCount).fill(defaultValues);
                    const fillRange = rosterSheet.getRangeByIndexes(1, startCol, rowCount, missingColumns.length);
                    fillRange.values = fillData;
                }
                
                await context.sync();
                return { ok: true, created: false, addedColumns: missingColumns };
            }
            
            return { ok: true, created: false, addedColumns: [] };
        });
    } catch (error) {
        console.error("[RosterSchema] Error:", error);
        return { ok: false, error: error.message };
    }
    
    // Backfill any empty Employee_IDs with names (for name-matched rosters)
    await backfillEmptyEmployeeIds();
}

/**
 * Fetch customer column mappings with expense_bucket from database
 * Used for rate calculation and expense review grouping
 * 
 * @param {string} companyId - Company UUID
 * @param {string} module - Module name (e.g., "payroll-recorder")
 * @returns {Promise<Array<{raw_header: string, pf_column_name: string, expense_bucket: string, include_in_matrix: boolean}>>}
 */
async function getCustomerMappingsWithBuckets(companyId, module) {
    if (!companyId) {
        console.warn('[CustomerMappings] No companyId provided');
        return [];
    }
    
    const SUPABASE_URL = "https://jgciqwzwacaesqjaoadc.supabase.co";
    const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE";
    
    try {
        const response = await fetch(
            `${SUPABASE_URL}/rest/v1/ada_customer_column_mappings?company_id=eq.${encodeURIComponent(companyId)}&module=eq.${encodeURIComponent(module)}&select=raw_header,pf_column_name,expense_bucket,include_in_matrix`,
            {
                headers: {
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
                    'Content-Type': 'application/json'
                }
            }
        );
        
        if (!response.ok) {
            console.warn('[CustomerMappings] Failed to fetch:', response.status, response.statusText);
            return [];
        }
        
        const mappings = await response.json();
        console.log(`[CustomerMappings] Loaded ${mappings.length} mappings for company=${companyId}, module=${module}`);
        
        return mappings;
    } catch (error) {
        console.warn('[CustomerMappings] Error fetching mappings:', error.message);
        return [];
    }
}

/**
 * Calculate per-employee Fixed bucket totals from PR_Data_Clean
 * Uses customer mappings expense_bucket to identify FIXED columns
 * 
 * @returns {{ ok: boolean, rates: Map<joinKey, { fixedTotal: number, hourlyRate: number }>, error?: string }}
 */
async function calculateEmployeeRatesFromPayroll() {
    if (!hasExcelRuntime()) return { ok: false, error: "Excel runtime unavailable", rates: new Map() };
    
    // Get company ID for customer mappings lookup
    const companyId = getConfigValue("SS_Company_ID");
    const module = "payroll-recorder";
    
    // Fetch customer mappings with expense_bucket
    const customerMappings = await getCustomerMappingsWithBuckets(companyId, module);
    console.log(`[RateCalc] Loaded ${customerMappings.length} customer mappings`);
    
    try {
        return await Excel.run(async (context) => {
            const dataSheet = context.workbook.worksheets.getItemOrNullObject("PR_Data_Clean");
            dataSheet.load("isNullObject");
            await context.sync();
            
            if (dataSheet.isNullObject) {
                console.log("[RateCalc] PR_Data_Clean not found");
                return { ok: false, error: "PR_Data_Clean not found", rates: new Map() };
            }
            
            const dataRange = dataSheet.getUsedRangeOrNullObject();
            dataRange.load("values");
            await context.sync();
            
            if (dataRange.isNullObject || !dataRange.values || dataRange.values.length < 2) {
                console.log("[RateCalc] PR_Data_Clean is empty");
                return { ok: false, error: "PR_Data_Clean is empty", rates: new Map() };
            }
            
            const headers = dataRange.values[0].map(h => String(h || "").toLowerCase().trim());
            const dataRows = dataRange.values.slice(1);
            
            // Find employee join key column
            const joinKeyInfo = findEmployeeJoinKeyColumn(headers);
            if (!joinKeyInfo || joinKeyInfo.index < 0) {
                console.log("[RateCalc] No employee identifier column found");
                return { ok: false, error: "No employee identifier column", rates: new Map() };
            }
            
            // Load taxonomy for sign information (still useful for debit/credit)
            const taxonomy = expenseTaxonomyCache.loaded ? expenseTaxonomyCache : null;
            console.log(`[RateCalc] Taxonomy has ${Object.keys(taxonomy?.measures || {}).length} measures loaded`);
            console.log(`[RateCalc] PR_Data_Clean headers: ${headers.slice(0, 10).join(', ')}...`);
            
            // Build FIXED column list from customer mappings (expense_bucket = 'FIXED')
            const fixedColumnIndexes = [];
            const fixedColumnNames = [];
            
            // Debug: Log what we're trying to match
            console.log(`[RateCalc] Customer mappings pf_column_names:`, 
                customerMappings.map(m => ({ pf: m.pf_column_name, bucket: m.expense_bucket })));
            console.log(`[RateCalc] PR_Data_Clean headers to match:`, headers);

            headers.forEach((header, colIdx) => {
                const headerLower = header.toLowerCase().trim();
                
                // Find matching customer mapping by pf_column_name (primary) or raw_header (fallback)
                const mapping = customerMappings.find(m => 
                    m.pf_column_name?.toLowerCase() === headerLower ||
                    m.raw_header?.toLowerCase() === headerLower
                );
                
                if (!mapping) {
                    // Debug: Log unmatched headers that look like amounts
                    if (headerLower.includes('amount') || headerLower.includes('wages') || headerLower.includes('salary')) {
                        console.log(`[RateCalc] No mapping found for "${header}"`);
                    }
                    return;
                }
                
                const bucket = (mapping.expense_bucket || "").toUpperCase();
                
                if (bucket === "FIXED") {
                    // Get sign from taxonomy dictionary, default to 1
                    const dictEntry = taxonomy?.measures?.[headerLower];
                    const sign = dictEntry?.sign ?? 1;
                    
                    fixedColumnIndexes.push({ colIdx, sign });
                    fixedColumnNames.push(header);
                    console.log(`[RateCalc]   ${header}: bucket=FIXED, sign=${sign}`);
                } else if (bucket) {
                    console.log(`[RateCalc]   ${header}: bucket=${bucket} (not FIXED, skipping)`);
                }
            });

            console.log(`[RateCalc] Found ${fixedColumnIndexes.length} FIXED columns: ${fixedColumnNames.join(', ') || '(none)'}`);

            if (fixedColumnIndexes.length === 0) {
                console.warn("[RateCalc] No FIXED bucket columns found in customer mappings");
                console.log("[RateCalc] Customer mappings expense_buckets:", 
                    [...new Set(customerMappings.map(m => m.expense_bucket).filter(Boolean))]);
                // Don't fail - just return empty rates
                return { ok: true, rates: new Map(), warning: "No FIXED bucket columns in customer mappings" };
            }
            
            // Calculate Fixed total per employee
            const employeeRates = new Map();
            
            for (const row of dataRows) {
                const keyValue = row[joinKeyInfo.index];
                const key = normalizeJoinKey(keyValue);
                if (!key || isNoiseName(key)) continue;
                
                // Sum Fixed columns for this row
                let rowFixedTotal = 0;
                for (const { colIdx, sign } of fixedColumnIndexes) {
                    const val = Number(row[colIdx]) || 0;
                    rowFixedTotal += val * sign;
                }
                
                // Accumulate by employee (in case of multiple rows per employee)
                if (employeeRates.has(key)) {
                    const existing = employeeRates.get(key);
                    existing.fixedTotal += rowFixedTotal;
                } else {
                    employeeRates.set(key, { fixedTotal: rowFixedTotal });
                }
            }
            
            // Calculate hourly rate: Fixed / 80 hours
            for (const [key, data] of employeeRates) {
                if (data.fixedTotal > 0) {
                    data.hourlyRate = data.fixedTotal / HOURS_PER_PAY_PERIOD;
                } else {
                    data.hourlyRate = 0;
                }
            }
            
            console.log(`[RateCalc] Computed hourly rates for ${employeeRates.size} employees (Fixed / ${HOURS_PER_PAY_PERIOD} hrs)`);
            
            return { ok: true, rates: employeeRates };
        });
    } catch (error) {
        console.error("[RateCalc] Error:", error);
        return { ok: false, error: error.message, rates: new Map() };
    }
}

async function updateRosterRatesFromPayrollAdvisory() {
    if (!hasExcelRuntime()) return;

    // Ensure taxonomy is loaded before calculating rates
    await fetchExpenseTaxonomy();
    
    const schema = await ensureRosterSchema();
    if (!schema.ok) {
        console.warn("[RosterRates] Cannot update rates - roster schema unavailable:", schema.error);
        return;
    }

    const rateResult = await calculateEmployeeRatesFromPayroll();
    if (!rateResult.ok || !rateResult.rates || rateResult.rates.size === 0) {
        console.warn("[RosterRates] No rates computed:", rateResult.error || "no rows");
        return;
    }

    const currentResult = await extractCurrentEmployeeSet();
    const joinKeyUsed = currentResult.ok ? currentResult.joinKeyUsed : "Employee_ID";
    const payrollDateISO = normalizePeriodKey(getPayrollDateValue()) || todayIso();

    await Excel.run(async (context) => {
        const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
        rosterSheet.load("isNullObject");
        await context.sync();
        if (rosterSheet.isNullObject) {
            console.warn("[RosterRates] SS_Employee_Roster not found");
            return;
        }

        const rosterRange = rosterSheet.getUsedRangeOrNullObject();
        rosterRange.load("values,rowCount,columnCount");
        await context.sync();

        if (rosterRange.isNullObject || !rosterRange.values || rosterRange.values.length < 2) {
            console.warn("[RosterRates] SS_Employee_Roster is empty");
            return;
        }

        const headers = rosterRange.values[0].map((h) => String(h || "").trim());
        const headersLower = headers.map((h) => h.toLowerCase());

        const idx = {};
        ROSTER_REQUIRED_COLUMNS.forEach((col) => {
            idx[col] = headersLower.indexOf(col.toLowerCase());
        });

        const joinKeyIdx = joinKeyUsed === "Employee_ID" && idx.Employee_ID >= 0 ? idx.Employee_ID : idx.Employee_Name;
        if (joinKeyIdx < 0) {
            console.warn("[RosterRates] No join key column found on roster for", joinKeyUsed);
            return;
        }
        if (idx.Hourly_Rate < 0 || idx.Rate_Source < 0 || idx.Rate_Updated_Date < 0) {
            console.warn("[RosterRates] Missing rate columns on roster (should not happen after ensureRosterSchema)");
            return;
        }

        const dataRows = rosterRange.values.slice(1);
        const rosterMap = new Map();
        dataRows.forEach((row, i) => {
            const key = normalizeJoinKey(row[joinKeyIdx]);
            if (key) rosterMap.set(key, { rowIndex: i + 1, row });
        });

        let matched = 0;
        let updated = 0;
        const skippedNotFound = [];
        const skippedInvalid = [];

        for (const [key, rateData] of rateResult.rates.entries()) {
            const rosterRow = rosterMap.get(key);
            if (!rosterRow) {
                skippedNotFound.push(key);
                continue;
            }
            matched++;

            const isManuallyManaged = idx.Is_Manually_Managed >= 0
                ? String(rosterRow.row[idx.Is_Manually_Managed] || "").toUpperCase() === "TRUE"
                : false;
            if (isManuallyManaged) continue;

            const hourlyRate = Number(rateData?.hourlyRate);
            if (!Number.isFinite(hourlyRate) || hourlyRate <= 0) {
                skippedInvalid.push(key);
                continue;
            }

            const currentRate = Number(rosterRow.row[idx.Hourly_Rate]) || 0;
            const rateChanged = currentRate === 0 || Math.abs(hourlyRate - currentRate) / currentRate > 0.01;
            if (!rateChanged) continue;

            rosterSheet.getRangeByIndexes(rosterRow.rowIndex, idx.Hourly_Rate, 1, 1).values = [[Math.round(hourlyRate * 100) / 100]];
            rosterSheet.getRangeByIndexes(rosterRow.rowIndex, idx.Rate_Source, 1, 1).values = [[`Payroll (Fixed / ${HOURS_PER_PAY_PERIOD}hrs)`]];
            rosterSheet.getRangeByIndexes(rosterRow.rowIndex, idx.Rate_Updated_Date, 1, 1).values = [[payrollDateISO]];
            updated++;
        }

        await context.sync();
        console.log("[RosterRates] Update complete", {
            computedRates: rateResult.rates.size,
            matched,
            updated,
            skippedNotFound: skippedNotFound.length,
            skippedInvalid: skippedInvalid.length,
            payrollDateISO
        });
        if (skippedNotFound.length) console.warn("[RosterRates] Employees not found in roster (skipped):", skippedNotFound.slice(0, 25));
        if (skippedInvalid.length) console.warn("[RosterRates] Employees with invalid/unknown rate (skipped):", skippedInvalid.slice(0, 25));
    });
}

/**
 * Extract current employee set from PR_Data_Clean
 * Returns: { employees: Map<joinKey, {name, department}>, periodKey, joinKeyUsed }
 */
async function extractCurrentEmployeeSet() {
    if (!hasExcelRuntime()) return { ok: false, error: "Excel runtime unavailable" };
    
    try {
        return await Excel.run(async (context) => {
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            cleanSheet.load("isNullObject");
            await context.sync();
            
            if (cleanSheet.isNullObject) {
                return { ok: false, error: "PR_Data_Clean not found. Run Create Matrix first." };
            }
            
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            cleanRange.load("values");
            await context.sync();
            
            if (cleanRange.isNullObject || !cleanRange.values || cleanRange.values.length < 2) {
                return { ok: false, error: "PR_Data_Clean is empty." };
            }
            
            const headers = cleanRange.values[0].map(h => normalizeHeader(String(h || "")));
            const dataRows = cleanRange.values.slice(1);
            
            // Find join key column - pass dataRows so we can check if Employee_ID has actual values
            const joinKeyInfo = findEmployeeJoinKeyColumn(headers, dataRows);
            if (joinKeyInfo.index < 0) {
                return { ok: false, error: "No employee identifier column found in PR_Data_Clean." };
            }
            
            console.log(`[RosterUpdate] Using join key: ${joinKeyInfo.type} (column index ${joinKeyInfo.index})`);
            
            // Find department column
            const deptIdx = pickDepartmentIndex(headers);
            
            // Find period key column (Pay_Date, Payroll_Date, etc.)
            const periodIdx = headers.findIndex(h => 
                h === "payroll_date" || h === "pay_date" || 
                h.includes("payroll") && h.includes("date") ||
                h.includes("pay") && h.includes("date")
            );
            
            // Get period key from config or first data row
            let periodKey = getPayrollDateValue();
            if (!periodKey && periodIdx >= 0 && dataRows[0]) {
                const rawDate = dataRows[0][periodIdx];
                periodKey = normalizePeriodKey(rawDate);
            }
            if (!periodKey) {
                periodKey = new Date().toISOString().split("T")[0]; // fallback to today
            }
            
            // Build employee map
            const employees = new Map();
            for (const row of dataRows) {
                const keyValue = row[joinKeyInfo.index];
                const key = normalizeJoinKey(keyValue);
                
                if (!key || isNoiseName(key)) continue;
                
                if (!employees.has(key)) {
                    employees.set(key, {
                        name: normalizeString(keyValue) || key,
                        department: deptIdx >= 0 ? normalizeString(row[deptIdx]) : ""
                    });
                }
            }
            
            console.log(`[RosterUpdate] Extracted ${employees.size} employees from PR_Data_Clean, period: ${periodKey}`);
            
            return {
                ok: true,
                employees,
                periodKey,
                joinKeyUsed: joinKeyInfo.type
            };
        });
    } catch (error) {
        console.error("[RosterUpdate] Extract error:", error);
        return { ok: false, error: error.message };
    }
}

/**
 * Read current roster data
 * Returns: { employees: Map<joinKey, rowData>, headers, headerIndexes }
 */
async function readCurrentRoster() {
    if (!hasExcelRuntime()) return { ok: false, error: "Excel runtime unavailable" };
    
    try {
        return await Excel.run(async (context) => {
            const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
            rosterSheet.load("isNullObject");
            await context.sync();
            
            if (rosterSheet.isNullObject) {
                return { ok: true, employees: new Map(), headers: [], headerIndexes: {}, isEmpty: true };
            }
            
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            rosterRange.load("values");
            await context.sync();
            
            if (rosterRange.isNullObject || !rosterRange.values || rosterRange.values.length < 1) {
                return { ok: true, employees: new Map(), headers: [], headerIndexes: {}, isEmpty: true };
            }
            
            const headers = rosterRange.values[0].map(h => String(h || "").trim());
            const headersLower = headers.map(h => h.toLowerCase());
            
            // Build header index map
            const headerIndexes = {};
            ROSTER_REQUIRED_COLUMNS.forEach(col => {
                const idx = headersLower.indexOf(col.toLowerCase());
                headerIndexes[col] = idx;
            });
            
            const dataRows = rosterRange.values.slice(1);
            
            // Find join key column - prefer Employee_ID only if it has actual values
            let joinKeyIdx = headerIndexes.Employee_Name;
            let joinKeyType = "Employee_Name";
            
            if (headerIndexes.Employee_ID >= 0) {
                // Check if Employee_ID column actually has values
                if (columnHasValues(dataRows, headerIndexes.Employee_ID)) {
                    joinKeyIdx = headerIndexes.Employee_ID;
                    joinKeyType = "Employee_ID";
                } else {
                    console.log(`[RosterUpdate] Roster has Employee_ID column but no values - using Employee_Name`);
                }
            }
            
            if (joinKeyIdx < 0) {
                return { ok: false, error: "No employee identifier column in roster." };
            }
            
            console.log(`[RosterUpdate] Reading roster with join key: ${joinKeyType} (column index ${joinKeyIdx})`);
            
            // Build employee map with row data
            const employees = new Map();
            
            for (let i = 0; i < dataRows.length; i++) {
                const row = dataRows[i];
                const keyValue = row[joinKeyIdx];
                const key = normalizeJoinKey(keyValue);
                
                if (!key) continue;
                
                employees.set(key, {
                    rowIndex: i + 1, // 1-based (accounting for header)
                    employeeId: headerIndexes.Employee_ID >= 0 ? row[headerIndexes.Employee_ID] : "",
                    employeeName: headerIndexes.Employee_Name >= 0 ? row[headerIndexes.Employee_Name] : "",
                    department: headerIndexes.Department_Name >= 0 ? row[headerIndexes.Department_Name] : "",
                    status: headerIndexes.Employment_Status >= 0 ? row[headerIndexes.Employment_Status] : "Unknown",
                    firstSeen: headerIndexes.First_Seen_Pay_Date >= 0 ? row[headerIndexes.First_Seen_Pay_Date] : "",
                    lastSeen: headerIndexes.Last_Seen_Pay_Date >= 0 ? row[headerIndexes.Last_Seen_Pay_Date] : "",
                    terminationDate: headerIndexes.Termination_Effective_Date >= 0 ? row[headerIndexes.Termination_Effective_Date] : "",
                    lastTerminatedDate: headerIndexes.Last_Terminated_Date >= 0 ? row[headerIndexes.Last_Terminated_Date] : "",
                    isManuallyManaged: headerIndexes.Is_Manually_Managed >= 0 ? 
                        String(row[headerIndexes.Is_Manually_Managed]).toUpperCase() === "TRUE" : false,
                    notes: headerIndexes.Notes >= 0 ? row[headerIndexes.Notes] : "",
                    rawRow: row
                });
            }
            
            return {
                ok: true,
                employees,
                headers,
                headerIndexes,
                joinKeyType,
                isEmpty: employees.size === 0
            };
        });
    } catch (error) {
        console.error("[RosterUpdate] Read roster error:", error);
        return { ok: false, error: error.message };
    }
}

/**
 * Compute roster deltas between current payroll and existing roster
 */
async function computeRosterDeltas() {
    rosterUpdateState.loading = true;
    rosterUpdateState.lastError = null;
    
    try {
        // Ensure schema first
        const schemaResult = await ensureRosterSchema();
        if (!schemaResult.ok) {
            rosterUpdateState.lastError = schemaResult.error;
            rosterUpdateState.status = "unavailable";
            return;
        }
        
        // Extract current employee set from PR_Data_Clean
        const currentResult = await extractCurrentEmployeeSet();
        if (!currentResult.ok) {
            rosterUpdateState.lastError = currentResult.error;
            rosterUpdateState.status = "unavailable";
            return;
        }
        
        // Read existing roster
        const rosterResult = await readCurrentRoster();
        if (!rosterResult.ok) {
            rosterUpdateState.lastError = rosterResult.error;
            rosterUpdateState.status = "unavailable";
            return;
        }
        
        const currentEmployees = currentResult.employees;
        const rosterEmployees = rosterResult.employees;
        const periodKey = currentResult.periodKey;
        
        // Compute deltas
        const newHires = [];
        const missingEmployees = [];
        const reactivations = [];
        let stillActive = 0;
        
        // Find new hires (in payroll but not in roster) AND reactivations (terminated employees appearing in payroll)
        currentEmployees.forEach((data, key) => {
            if (!rosterEmployees.has(key)) {
                // New hire - not in roster at all
                newHires.push({
                    key,
                    name: data.name,
                    department: data.department
                });
            } else {
                // Employee exists in roster - check if they're terminated (potential reactivation)
                const rosterData = rosterEmployees.get(key);
                const status = String(rosterData.status || "").toLowerCase();
                
                if (status === "terminated") {
                    // Reactivation detected - terminated employee appearing in current payroll
                    reactivations.push({
                        key,
                        name: rosterData.employeeName || rosterData.employeeId || key,
                        department: data.department || rosterData.department,
                        terminationDate: rosterData.terminationDate,
                        isManuallyManaged: rosterData.isManuallyManaged
                    });
                } else {
                    // Still active (already Active/Unknown status)
                    stillActive++;
                }
            }
        });
        
        // Find missing employees (in roster as Active but not in current payroll)
        rosterEmployees.forEach((data, key) => {
            const status = String(data.status || "").toLowerCase();
            const isActive = status === "active" || status === "unknown" || status === "";
            
            if (isActive && !currentEmployees.has(key)) {
                missingEmployees.push({
                    key,
                    name: data.employeeName || data.employeeId || key,
                    department: data.department,
                    lastSeenDate: data.lastSeen,
                    isManuallyManaged: data.isManuallyManaged
                });
            }
        });
        
        // Update state
        rosterUpdateState.periodKey = periodKey;
        rosterUpdateState.joinKeyUsed = currentResult.joinKeyUsed;
        rosterUpdateState.newHires = newHires;
        rosterUpdateState.missingEmployees = missingEmployees;
        rosterUpdateState.reactivations = reactivations;
        rosterUpdateState.stillActive = stillActive;
        rosterUpdateState.hasData = true;
        
        // Determine status
        if (newHires.length === 0 && missingEmployees.length === 0 && reactivations.length === 0) {
            rosterUpdateState.status = "ok";
        } else {
            rosterUpdateState.status = "review";
        }
        
        console.log(`[RosterUpdate] Deltas computed: ${newHires.length} new, ${missingEmployees.length} missing, ${stillActive} still active`);
        
    } catch (error) {
        console.error("[RosterUpdate] Compute deltas error:", error);
        rosterUpdateState.lastError = error.message;
        rosterUpdateState.status = "unavailable";
    } finally {
        rosterUpdateState.loading = false;
    }
}

/**
 * Apply roster updates (add new hires, mark terminations)
 * Respects Is_Manually_Managed flag and user overrides
 */
async function applyRosterUpdates() {
    if (!hasExcelRuntime()) {
        showToast("Excel runtime unavailable", "error");
        return;
    }
    
    rosterUpdateState.applyPending = true;
    showToast("Applying roster updates...", "info", 2000);
    
    try {
        await Excel.run(async (context) => {
            const rosterSheet = context.workbook.worksheets.getItem("SS_Employee_Roster");
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            rosterRange.load("values, rowCount, columnCount");
            await context.sync();
            
            const headers = rosterRange.values[0].map(h => String(h || "").trim());
            const headersLower = headers.map(h => h.toLowerCase());
            
            // Build header index map
            const idx = {};
            ROSTER_REQUIRED_COLUMNS.forEach(col => {
                idx[col] = headersLower.indexOf(col.toLowerCase());
            });
            
            const periodKey = rosterUpdateState.periodKey;
            
            // Read current data
            const dataRows = rosterRange.values.slice(1);
            
            // ========================================================================
            // CRITICAL: Use the SAME join key that extractCurrentEmployeeSet() uses
            // This ensures roster and payroll are compared with matching keys
            // The joinKeyUsed was determined by computeRosterDeltas() based on whether
            // BOTH datasets have Employee_ID with actual values
            // ========================================================================
            const joinKeyUsed = rosterUpdateState.joinKeyUsed || "Employee_Name";
            let joinKeyIdx;
            if (joinKeyUsed === "Employee_ID" && idx.Employee_ID >= 0) {
                joinKeyIdx = idx.Employee_ID;
            } else {
                joinKeyIdx = idx.Employee_Name;
            }
            
            console.log(`[RosterUpdate] Using join key: ${joinKeyUsed} (roster column index: ${joinKeyIdx})`);
            
            // Build roster map using the SAME join key type as payroll
            const rosterMap = new Map();
            dataRows.forEach((row, i) => {
                const key = normalizeJoinKey(row[joinKeyIdx]);
                if (key) rosterMap.set(key, { rowIndex: i, row });
            });
            
            console.log(`[RosterUpdate] Built rosterMap with ${rosterMap.size} entries using ${joinKeyUsed}`);
            
            // Track updates
            const rowUpdates = []; // [{rowIndex, colIndex, value}]
            const newRows = [];
            let updatedCount = 0;
            let addedCount = 0;
            let terminatedCount = 0;
            let ratesUpdatedCount = 0;
            
            // Pre-calculate employee rates from current payroll
            const rateResult = await calculateEmployeeRatesFromPayroll();
            const employeeRates = rateResult.ok ? rateResult.rates : new Map();
            console.log(`[RosterUpdate] Rate calculation: ${rateResult.ok ? employeeRates.size + " employees" : rateResult.error}`);
            
            // 1. Process still-active employees (update Last_Seen_Pay_Date)
            const currentResult = await extractCurrentEmployeeSet();
            if (currentResult.ok) {
                currentResult.employees.forEach((empData, key) => {
                    const existing = rosterMap.get(key);
                    if (existing) {
                        const isManuallyManaged = idx.Is_Manually_Managed >= 0 && String(existing.row[idx.Is_Manually_Managed] || "").toUpperCase() === "TRUE";
                        // Always update Last_Seen_Pay_Date (even for manually managed)
                        if (idx.Last_Seen_Pay_Date >= 0) {
                            rowUpdates.push({
                                rowIndex: existing.rowIndex + 1,
                                colIndex: idx.Last_Seen_Pay_Date,
                                value: formatDateForRoster(periodKey)
                            });
                        }
                        
                        // Update status to Active
                        if (!isManuallyManaged && idx.Employment_Status >= 0) {
                            const currentStatus = String(existing.row[idx.Employment_Status] || "").toLowerCase();
                            if (currentStatus !== "active") {
                                rowUpdates.push({
                                    rowIndex: existing.rowIndex + 1,
                                    colIndex: idx.Employment_Status,
                                    value: "Active"
                                });
                            }
                        }
                        
                        // Update department if empty
                        if (!isManuallyManaged && idx.Department_Name >= 0 && empData.department) {
                            const currentDept = String(existing.row[idx.Department_Name] || "").trim();
                            if (!currentDept) {
                                rowUpdates.push({
                                    rowIndex: existing.rowIndex + 1,
                                    colIndex: idx.Department_Name,
                                    value: empData.department
                                });
                            }
                        }
                        
                        // Update hourly rate from payroll
                        if (!isManuallyManaged && idx.Hourly_Rate >= 0) {
                            const rateData = employeeRates.get(key);
                            if (rateData && rateData.hourlyRate > 0) {
                                const currentRate = Number(existing.row[idx.Hourly_Rate]) || 0;
                                // Only update if rate changed by more than 1% or was empty
                                const rateChanged = currentRate === 0 || 
                                    Math.abs(rateData.hourlyRate - currentRate) / currentRate > 0.01;
                                
                                if (rateChanged) {
                                    rowUpdates.push({
                                        rowIndex: existing.rowIndex + 1,
                                        colIndex: idx.Hourly_Rate,
                                        value: Math.round(rateData.hourlyRate * 100) / 100 // Round to cents
                                    });
                                    
                                    // Update source to indicate payroll-derived
                                    if (idx.Rate_Source >= 0) {
                                        rowUpdates.push({
                                            rowIndex: existing.rowIndex + 1,
                                            colIndex: idx.Rate_Source,
                                            value: `Payroll (Fixed $${rateData.fixedTotal.toFixed(0)} / ${HOURS_PER_PAY_PERIOD}hrs)`
                                        });
                                    }
                                    
                                    // Update timestamp
                                    if (idx.Rate_Updated_Date >= 0) {
                                        rowUpdates.push({
                                            rowIndex: existing.rowIndex + 1,
                                            colIndex: idx.Rate_Updated_Date,
                                            value: periodKey
                                        });
                                    }
                                    
                                    ratesUpdatedCount++;
                                }
                            }
                        }
                        
                        updatedCount++;
                    }
                });
            }
            
            // 2. Add new hires
            const { newHires, overrides } = rosterUpdateState;
            for (const hire of newHires) {
                const override = overrides.get(hire.key);
                if (override?.action === "ignore") continue;
                
                const newRow = Array(headers.length).fill("");
                // Only populate Employee_ID if we actually have an ID value
                // Don't overwrite with name - leave blank if no ID available (PTO sync will populate)
                if (idx.Employee_ID >= 0) {
                    if (joinKeyUsed === "Employee_ID") {
                        // We're matching by ID, so the key IS the employee ID
                        newRow[idx.Employee_ID] = hire.key;
                    } else if (hire.employeeId) {
                        // Use actual Employee_ID if available in the hire data
                        newRow[idx.Employee_ID] = hire.employeeId;
                    }
                    // Otherwise leave blank - will be populated by PTO sync
                }
                if (idx.Employee_Name >= 0) newRow[idx.Employee_Name] = hire.name;
                if (idx.Department_Name >= 0) newRow[idx.Department_Name] = hire.department;
                if (idx.Employment_Status >= 0) newRow[idx.Employment_Status] = "Active";
                if (idx.First_Seen_Pay_Date >= 0) newRow[idx.First_Seen_Pay_Date] = formatDateForRoster(periodKey);
                if (idx.Last_Seen_Pay_Date >= 0) newRow[idx.Last_Seen_Pay_Date] = formatDateForRoster(periodKey);
                if (idx.Is_Manually_Managed >= 0) newRow[idx.Is_Manually_Managed] = "FALSE";
                
                // Set initial hourly rate from payroll for new hires
                const rateData = employeeRates.get(hire.key);
                if (rateData && rateData.hourlyRate > 0) {
                    if (idx.Hourly_Rate >= 0) {
                        newRow[idx.Hourly_Rate] = Math.round(rateData.hourlyRate * 100) / 100;
                    }
                    if (idx.Rate_Source >= 0) {
                        newRow[idx.Rate_Source] = `Payroll (Fixed $${rateData.fixedTotal.toFixed(0)} / ${HOURS_PER_PAY_PERIOD}hrs)`;
                    }
                    if (idx.Rate_Updated_Date >= 0) {
                        newRow[idx.Rate_Updated_Date] = formatDateForRoster(periodKey);
                    }
                    ratesUpdatedCount++;
                }
                
                newRows.push(newRow);
                addedCount++;
            }
            
            // 3. Handle reactivations (terminated employees appearing in current payroll)
            const { reactivations } = rosterUpdateState;
            let reactivatedCount = 0;
            
            for (const reactivation of reactivations) {
                const override = overrides.get(reactivation.key);

                // If manually managed and no explicit override, require user review
                if (reactivation.isManuallyManaged && override?.action !== "reactivate") {
                    continue;
                }
                
                // If user explicitly wants to keep terminated, skip
                if (override?.action === "keep_terminated") {
                    continue;
                }
                
                // Default: reactivate (set to Active, preserve termination history)
                const existing = rosterMap.get(reactivation.key);
                if (existing) {
                    // Preserve termination history in Last_Terminated_Date before clearing
                    if (idx.Last_Terminated_Date >= 0 && reactivation.terminationDate) {
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Last_Terminated_Date,
                            value: formatDateForRoster(reactivation.terminationDate)
                        });
                    }
                    
                    // Clear current termination date
                    if (idx.Termination_Effective_Date >= 0) {
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Termination_Effective_Date,
                            value: ""
                        });
                    }
                    
                    // Set status to Active
                    if (idx.Employment_Status >= 0) {
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Employment_Status,
                            value: "Active"
                        });
                    }
                    
                    // Update Last_Seen_Pay_Date
                    if (idx.Last_Seen_Pay_Date >= 0) {
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Last_Seen_Pay_Date,
                            value: formatDateForRoster(periodKey)
                        });
                    }
                    
                    // Add note about reactivation
                    if (idx.Notes >= 0) {
                        const existingNotes = existing.row[idx.Notes] || "";
                        const reactivationNote = `Reactivated ${periodKey} (prev term: ${reactivation.terminationDate || "unknown"})`;
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Notes,
                            value: existingNotes ? `${existingNotes}; ${reactivationNote}` : reactivationNote
                        });
                    }
                    
                    reactivatedCount++;
                }
            }
            
            // 4. Mark terminated employees (if not manually managed and not overridden)
            const { missingEmployees } = rosterUpdateState;
            for (const missing of missingEmployees) {
                if (missing.isManuallyManaged) continue;
                const override = overrides.get(missing.key);
                if (override?.action === "keep_active") {
                    // User wants to keep this person active - set Is_Manually_Managed
                    const existing = rosterMap.get(missing.key);
                    if (existing && idx.Is_Manually_Managed >= 0) {
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Is_Manually_Managed,
                            value: "TRUE"
                        });
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Notes,
                            value: `Kept active by user override (${periodKey})`
                        });
                    }
                    continue;
                }
                
                // Default: mark as terminated
                const existing = rosterMap.get(missing.key);
                if (existing) {
                    if (idx.Employment_Status >= 0) {
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Employment_Status,
                            value: "Terminated"
                        });
                    }
                    if (idx.Termination_Effective_Date >= 0) {
                        rowUpdates.push({
                            rowIndex: existing.rowIndex + 1,
                            colIndex: idx.Termination_Effective_Date,
                            value: formatDateForRoster(periodKey)
                        });
                    }
                    terminatedCount++;
                }
            }
            
            // Apply row updates (batch by row for efficiency)
            const updatesByRow = new Map();
            rowUpdates.forEach(u => {
                if (!updatesByRow.has(u.rowIndex)) updatesByRow.set(u.rowIndex, []);
                updatesByRow.get(u.rowIndex).push(u);
            });
            
            for (const [rowIdx, updates] of updatesByRow) {
                for (const u of updates) {
                    const cell = rosterSheet.getRangeByIndexes(rowIdx, u.colIndex, 1, 1);
                    cell.values = [[u.value]];
                }
            }
            
            // Append new rows
            if (newRows.length > 0) {
                const lastRow = rosterRange.rowCount;
                const appendRange = rosterSheet.getRangeByIndexes(lastRow, 0, newRows.length, headers.length);
                appendRange.values = newRows;
            }
            
            // Apply date formatting to date columns
            const dateFormat = "yyyy-mm-dd";
            const dateColumns = [
                idx.First_Seen_Pay_Date,
                idx.Last_Seen_Pay_Date,
                idx.Termination_Effective_Date,
                idx.Last_Terminated_Date,
                idx.Rate_Updated_Date
            ].filter(i => i >= 0);
            
            for (const colIdx of dateColumns) {
                const dataRowCount = rosterRange.rowCount - 1;
                if (dataRowCount > 0) {
                    const colRange = rosterSheet.getRangeByIndexes(1, colIdx, dataRowCount, 1);
                    colRange.numberFormat = [[dateFormat]];
                }
            }
            
            await context.sync();
            
            // Log summary
            console.log(`[RosterUpdate] Applied: ${addedCount} added, ${reactivatedCount} reactivated, ${terminatedCount} terminated, ${updatedCount} updated, ${ratesUpdatedCount} rates updated`);
            
            // Clear state and refresh
            rosterUpdateState.newHires = [];
            rosterUpdateState.missingEmployees = [];
            rosterUpdateState.reactivations = [];
            rosterUpdateState.overrides.clear();
            rosterUpdateState.status = "ok";
            
            // Build toast message
            const parts = [];
            if (addedCount > 0) parts.push(`${addedCount} added`);
            if (reactivatedCount > 0) parts.push(`${reactivatedCount} reactivated`);
            if (terminatedCount > 0) parts.push(`${terminatedCount} terminated`);
            if (ratesUpdatedCount > 0) parts.push(`${ratesUpdatedCount} rates updated`);
            const toastMsg = parts.length > 0 ? `Roster updated: ${parts.join(", ")}` : "Roster is up to date";
            
            showToast(toastMsg, "success", 4000);
        });
        
    } catch (error) {
        console.error("[RosterUpdate] Apply error:", error);
        showToast("Failed to update roster: " + error.message, "error", 6000);
    } finally {
        rosterUpdateState.applyPending = false;
        renderApp();
    }
}

/**
 * Render Employee Roster Updates advisory card for Step 1
 */
function renderEmployeeRosterUpdatesCard() {
    const state = rosterUpdateState;
    
    // Determine status badge
    let statusBadge = "";
    if (state.loading) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status"><span>Loading</span></span>`;
    } else if (state.lastError) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--unavailable" role="status"><span>Unavailable</span></span>`;
    } else if (!state.hasData) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status"><span>Pending</span></span>`;
    } else if (state.status === "ok") {
        statusBadge = `<span class="pf-status-badge pf-status-badge--ok" role="status"><span>OK</span></span>`;
    } else {
        statusBadge = `<span class="pf-status-badge pf-status-badge--review" role="status"><span>Review</span></span>`;
    }
    
    // Build content
    let contentHtml = "";
    
    if (state.loading) {
        contentHtml = `<p style="color: rgba(255,255,255,0.6); font-size: 12px;">Analyzing roster changes...</p>`;
    } else if (state.lastError) {
        contentHtml = `<p class="pf-metric-hint pf-metric-hint--warning">${escapeHtml(state.lastError)}</p>`;
    } else if (!state.hasData) {
        contentHtml = `<p style="color: rgba(255,255,255,0.6); font-size: 12px;">Run Create Matrix to analyze roster changes.</p>`;
    } else {
        const { newHires, missingEmployees, reactivations, stillActive, periodKey, joinKeyUsed } = state;
        
        const periodInfo = periodKey ? `<span style="color: rgba(255,255,255,0.5); font-size: 11px;">Period: ${formatFriendlyPeriod(periodKey)}</span>` : "";
        const joinKeyInfo = joinKeyUsed ? `<span style="color: rgba(255,255,255,0.5); font-size: 11px; margin-left: 8px;">Matching by: ${joinKeyUsed}</span>` : "";
        
        let summaryHtml = `
            <div style="display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 8px; font-size: 12px;">
                <span><strong>${stillActive}</strong> still active</span>
                <span style="color: #4ade80;"><strong>${newHires.length}</strong> new</span>
                ${reactivations.length > 0 ? `<span style="color: #60a5fa;"><strong>${reactivations.length}</strong> reactivation${reactivations.length !== 1 ? "s" : ""}</span>` : ""}
                <span style="color: #fbbf24;"><strong>${missingEmployees.length}</strong> missing</span>
            </div>
            <div style="font-size: 11px; color: rgba(255,255,255,0.5); margin-bottom: 8px;">
                ${periodInfo} ${joinKeyInfo}
            </div>
        `;
        
        // New hires list
        let newHiresHtml = "";
        if (newHires.length > 0) {
            const displayCount = Math.min(newHires.length, 8);
            const moreCount = newHires.length - displayCount;
            const names = newHires.slice(0, displayCount).map(h => escapeHtml(h.name)).join(", ");
            newHiresHtml = `
                <div class="pf-coverage-check pf-coverage-check--pass" style="margin-top: 8px;">
                    <span class="pf-coverage-icon" style="color: #4ade80;" aria-hidden="true">${CHECK_CIRCLE_SVG}</span>
                    <span><strong>New employees:</strong> ${names}${moreCount > 0 ? ` (+${moreCount} more)` : ""}</span>
                </div>
            `;
        }
        
        // Reactivations list (terminated employees appearing again)
        let reactivationsHtml = "";
        if (reactivations.length > 0) {
            const displayCount = Math.min(reactivations.length, 8);
            const moreCount = reactivations.length - displayCount;
            const names = reactivations.slice(0, displayCount).map(r => escapeHtml(r.name)).join(", ");
            const manualCount = reactivations.filter(r => r.isManuallyManaged).length;
            const manualNote = manualCount > 0 ? ` <span style="color: #f59e0b;">(${manualCount} require review)</span>` : "";
            reactivationsHtml = `
                <div class="pf-coverage-check" style="margin-top: 8px; border-left: 3px solid #60a5fa; padding-left: 12px;">
                    <span class="pf-coverage-icon" style="color: #60a5fa;" aria-hidden="true">${ALERT_TRIANGLE_SVG}</span>
                    <span><strong>Reactivation detected:</strong> ${names}${moreCount > 0 ? ` (+${moreCount} more)` : ""}${manualNote}</span>
                </div>
                <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 4px; margin-left: 24px;">
                    Previously terminated employees appearing in current payroll. Termination history will be preserved.
                </p>
            `;
        }
        
        // Missing employees list (candidate terminations)
        let missingHtml = "";
        if (missingEmployees.length > 0) {
            const displayCount = Math.min(missingEmployees.length, 8);
            const moreCount = missingEmployees.length - displayCount;
            const names = missingEmployees.slice(0, displayCount).map(m => escapeHtml(m.name)).join(", ");
            const manualCount = missingEmployees.filter(m => m.isManuallyManaged).length;
            const manualNote = manualCount > 0 ? ` (${manualCount} manually managed)` : "";
            missingHtml = `
                <div class="pf-coverage-check pf-coverage-check--warn" style="margin-top: 8px;">
                    <span class="pf-coverage-icon" style="color: #fbbf24;" aria-hidden="true">${ALERT_TRIANGLE_SVG}</span>
                    <span><strong>Missing from payroll:</strong> ${names}${moreCount > 0 ? ` (+${moreCount} more)` : ""}${manualNote}</span>
                </div>
            `;
        }

        // Action buttons
        let actionsHtml = "";
        const hasChanges = newHires.length > 0 || missingEmployees.length > 0 || reactivations.length > 0;
        if (hasChanges) {
            actionsHtml = `
                <div class="pf-pill-row" style="margin-top: 12px; gap: 8px;">
                    <button type="button" class="pf-pill-btn pf-pill-btn--sm" id="roster-apply-btn" ${state.applyPending ? "disabled" : ""}>
                        Apply Suggested Updates
                    </button>
                    <button type="button" class="pf-action-toggle pf-action-toggle--subtle" id="roster-refresh-btn">
                        ${REFRESH_ICON_SVG}<span style="margin-left: 6px;">Refresh</span>
                    </button>
                </div>
                <p style="font-size: 11px; color: rgba(255,255,255,0.5); margin-top: 8px;">
                    New employees added as Active. Missing marked Terminated. Reactivations restore Active status (history preserved).
                    Manually managed rows require explicit review.
                </p>
            `;
        } else {
            actionsHtml = `
                <div style="margin-top: 8px;">
                    <button type="button" class="pf-action-toggle pf-action-toggle--subtle" id="roster-refresh-btn">
                        ↻ Refresh
                    </button>
                </div>
            `;
        }
        
        contentHtml = summaryHtml + newHiresHtml + reactivationsHtml + missingHtml + actionsHtml;
    }
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head">
                <div>
                    <h3>Employee Roster Updates ${statusBadge}</h3>
                    <p class="pf-config-subtext">Auto-maintain SS_Employee_Roster from payroll data (advisory)</p>
                </div>
            </div>
            <div style="padding: 12px 16px;">
                ${contentHtml}
            </div>
        </article>
    `;
}

/**
 * Refresh roster update analysis (preserves scroll position)
 */
async function refreshRosterUpdates() {
    const scrollY = window.scrollY;
    await computeRosterDeltas();
    renderApp();
    // Restore scroll position after render
    requestAnimationFrame(() => window.scrollTo(0, scrollY));
}

function normalizePeriodKey(value) {
    if (value instanceof Date) {
        return formatDateFromDate(value);
    }
    if (typeof value === "number" && !Number.isNaN(value)) {
        // Excel serial date - convert using UTC to avoid timezone shift
        const date = convertExcelDate(value);
        return date ? formatDateFromDate(date) : "";
    }
    const str = String(value ?? "").trim();
    if (!str) return "";
    
    // If already in ISO format (YYYY-MM-DD), return as-is - no conversion needed
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    
    // For other date formats, extract components directly to avoid timezone issues
    // Try to parse using local interpretation
    const parsed = new Date(str);
    if (!Number.isNaN(parsed.getTime())) {
        // Use local methods since the string was likely entered in local time
        const year = parsed.getFullYear();
        const month = String(parsed.getMonth() + 1).padStart(2, "0");
        const day = String(parsed.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    return str;
}

function convertExcelDate(serial) {
    if (!Number.isFinite(serial)) return null;
    const utcDays = Math.floor(serial - 25569);
    if (!Number.isFinite(utcDays)) return null;
    const utcValue = utcDays * 86400 * 1000;
    const date = new Date(utcValue);
    // Mark this date as UTC-derived so formatDateFromDate can use UTC methods
    date._isUTC = true;
    return date;
}

function formatFriendlyPeriod(key) {
    if (!key) return "";
    
    // If ISO format (YYYY-MM-DD), parse manually to avoid timezone shift
    if (/^\d{4}-\d{2}-\d{2}$/.test(key)) {
        const [year, month, day] = key.split("-").map(Number);
        const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        return `${monthNames[month - 1]} ${day}, ${year}`;
    }
    
    // For other formats, try parsing but use local interpretation
    const parsed = new Date(key);
    if (!Number.isNaN(parsed.getTime())) {
        return parsed.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
    }
    return key;
}

function toNumber(value) {
    if (value == null || value === "") return 0;
    const num = Number(value);
    return Number.isFinite(num) ? num : 0;
}

/**
 * LEGACY: Classify expense component by keywords (fallback only)
 */
function classifyExpenseComponentLegacy(label) {
    const normalized = normalizeString(label).toLowerCase();
    if (!normalized) return "VARIABLE";
    if (
        normalized.includes("burden") ||
        normalized.includes("tax") ||
        normalized.includes("benefit") ||
        normalized.includes("fica") ||
        normalized.includes("insurance") ||
        normalized.includes("health") ||
        normalized.includes("medicare")
    ) {
        return "BURDEN";
    }
    if (
        normalized.includes("bonus") ||
        normalized.includes("commission") ||
        normalized.includes("variable") ||
        normalized.includes("overtime") ||
        normalized.includes("per diem")
    ) {
        return "VARIABLE";
    }
    return "FIXED";
}

// =============================================================================
// DEEP DIAGNOSTIC: Expense Review Measure Pipeline Trace
// =============================================================================

/**
 * Trace the entire Expense Review measure pipeline to diagnose inclusion issues.
 * This function logs detailed diagnostics at every stage to identify where
 * the "collapse" happens (why only a subset of PR_Data_Clean is being recognized).
 * 
 * Call this from the console or add a debug button to trigger it.
 */
async function traceExpenseReviewMeasurePipeline() {
    console.log("\n");
    console.log("╔══════════════════════════════════════════════════════════════════════════╗");
    console.log("║  EXPENSE REVIEW MEASURE PIPELINE TRACE                                   ║");
    console.log("║  Deep diagnostic for PR_Data_Clean → Expense Review data flow           ║");
    console.log("╚══════════════════════════════════════════════════════════════════════════╝\n");
    
    const trace = {
        phase0: { sheetName: null, usedRange: null, rowCount: 0, colCount: 0, headers: [], sampleData: [] },
        phase1: {
            allHeaders: [],
            numericCandidateHeaders: [],
            dimensionHeaders: [],
            measureHeadersAfterDimRemoval: [],
            matchedDictionaryHeaders: [],
            unmatchedHeaders: [],
            dictionaryMetadata: {}
        },
        phase2: {
            perColumnDecisions: [],
            includedColumns: [],
            excludedColumns: [],
            excludeReasonCounts: {}
        },
        totals: {
            TOTAL_NUMERIC_SHEET: 0,
            TOTAL_AFTER_METADATA_JOIN: 0,
            TOTAL_AFTER_INCLUSION_RULES: 0,
            TOTAL_BURDEN_ONLY: 0,
            delta: 0,
            excludedSum: 0
        }
    };
    
    if (!hasExcelRuntime()) {
        console.error("[Trace] Excel runtime not available");
        return trace;
    }
    
    try {
        await Excel.run(async (context) => {
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 0: Confirm we're reading the correct sheet/range
            // ═══════════════════════════════════════════════════════════════════════
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 0: Sheet/Range Confirmation");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            cleanSheet.load("name, isNullObject");
            await context.sync();
            
            if (cleanSheet.isNullObject) {
                console.error("[Trace] ❌ PR_Data_Clean sheet NOT FOUND!");
                return;
            }
            
            trace.phase0.sheetName = cleanSheet.name;
            console.log(`[Trace] ✓ Sheet found: ${cleanSheet.name}`);
            
            const usedRange = cleanSheet.getUsedRangeOrNullObject();
            usedRange.load("address, rowCount, columnCount, values");
            await context.sync();
            
            if (usedRange.isNullObject) {
                console.error("[Trace] ❌ Used range is null - sheet is empty!");
                return;
            }
            
            trace.phase0.usedRange = usedRange.address;
            trace.phase0.rowCount = usedRange.rowCount;
            trace.phase0.colCount = usedRange.columnCount;
            
            console.log(`[Trace] ✓ usedRange: ${usedRange.address}`);
            console.log(`[Trace] ✓ rowCount: ${usedRange.rowCount}, colCount: ${usedRange.columnCount}`);
            
            const allValues = usedRange.values;
            const rawHeaders = allValues[0] || [];
            const headers = rawHeaders.map(h => String(h || "").trim());
            const headersLower = headers.map(h => h.toLowerCase());
            trace.phase0.headers = headers.slice(0, 30);
            
            console.log(`[Trace] Headers (first 30):`);
            headers.slice(0, 30).forEach((h, i) => console.log(`   ${i}: "${h}"`));
            
            const dataRows = allValues.slice(1);
            console.log(`[Trace] Data rows: ${dataRows.length}`);
            
            // Sample values from first 3 numeric columns
            const numericSamples = [];
            for (let col = 0; col < headers.length && numericSamples.length < 5; col++) {
                let hasNum = false;
                let samples = [];
                for (let row = 0; row < Math.min(5, dataRows.length); row++) {
                    const val = dataRows[row][col];
                    if (typeof val === 'number' || (!isNaN(Number(val)) && val !== '')) {
                        hasNum = true;
                        samples.push(val);
                    }
                }
                if (hasNum) {
                    numericSamples.push({ header: headers[col], samples });
                }
            }
            trace.phase0.sampleData = numericSamples;
            console.log(`[Trace] Sample numeric columns:`, numericSamples);
            
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 1A: Raw header inventory
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 1A: Raw Header Inventory");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            trace.phase1.allHeaders = [...headers];
            console.log(`[Trace] Total headers: ${headers.length}`);
            
            // Identify numeric candidate headers (columns with at least one numeric value)
            const numericCandidates = [];
            const columnSums = {};
            const columnNonZeroCounts = {};
            
            headers.forEach((header, colIdx) => {
                let sum = 0;
                let nonZeroCount = 0;
                let hasNumeric = false;
                
                for (let rowIdx = 0; rowIdx < Math.min(200, dataRows.length); rowIdx++) {
                    const val = dataRows[rowIdx][colIdx];
                    const num = Number(val);
                    if (!isNaN(num)) {
                        if (num !== 0) {
                            hasNumeric = true;
                            nonZeroCount++;
                        }
                        sum += num;
                    }
                }
                
                // Full sum for all rows
                let fullSum = 0;
                for (const row of dataRows) {
                    const val = Number(row[colIdx]);
                    if (!isNaN(val)) fullSum += val;
                }
                
                columnSums[header] = fullSum;
                columnNonZeroCounts[header] = nonZeroCount;
                
                if (hasNumeric) {
                    numericCandidates.push({ header, sum: fullSum, nonZeroCount });
                }
            });
            
            trace.phase1.numericCandidateHeaders = numericCandidates.map(c => c.header);
            console.log(`[Trace] Numeric candidate headers: ${numericCandidates.length} of ${headers.length}`);
            numericCandidates.forEach(c => console.log(`   ${c.header}: $${c.sum.toLocaleString()} (${c.nonZeroCount} non-zero)`));
            
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 1B: Dimension detection result
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 1B: Dimension Detection");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            const taxonomy = expenseTaxonomyCache;
            console.log(`[Trace] Taxonomy loaded: ${taxonomy.loaded}`);
            console.log(`[Trace] Dimensions in cache: ${taxonomy.dimensions?.size || 0}`);
            console.log(`[Trace] Measures in cache: ${Object.keys(taxonomy.measures || {}).length}`);
            
            if (taxonomy.dimensions) {
                console.log(`[Trace] Dimension keywords: ${[...taxonomy.dimensions].slice(0, 20).join(", ")}`);
            }
            
            const dimensionHeaders = [];
            const dimensionReasons = {};
            const measureCandidates = [];
            
            headers.forEach((header, idx) => {
                const headerLower = header.toLowerCase();
                
                // Check if it's a dimension
                let isDimension = false;
                let dimensionReason = null;
                
                if (taxonomy.loaded && taxonomy.dimensions && taxonomy.dimensions.has(headerLower)) {
                    isDimension = true;
                    dimensionReason = "matched ada_payroll_dimensions";
                }
                
                if (isDimension) {
                    dimensionHeaders.push(header);
                    dimensionReasons[header] = dimensionReason;
                } else {
                    measureCandidates.push(header);
                }
            });
            
            trace.phase1.dimensionHeaders = dimensionHeaders;
            trace.phase1.measureHeadersAfterDimRemoval = measureCandidates;
            
            console.log(`[Trace] Dimension headers detected: ${dimensionHeaders.length}`);
            dimensionHeaders.forEach(h => console.log(`   ${h} → ${dimensionReasons[h]}`));
            console.log(`[Trace] Measure candidates after dimension removal: ${measureCandidates.length}`);
            
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 1C: Dictionary enrichment join result
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 1C: Dictionary Enrichment Join");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            const matchedHeaders = [];
            const unmatchedHeaders = [];
            const metadataSnapshot = {};
            
            measureCandidates.forEach(header => {
                const headerLower = header.toLowerCase().trim();
                const meta = (taxonomy.loaded && taxonomy.measures) ? taxonomy.measures[headerLower] : null;
                
                if (meta) {
                    matchedHeaders.push(header);
                    metadataSnapshot[header] = {
                        pf_column_name: header,
                        data_type: meta.data_type || "number",
                        term_type: meta.term_type || null,
                        side: meta.side || null,
                        bucket: meta.bucket || null,
                        include: meta.include,
                        sign: meta.sign
                    };
                } else {
                    unmatchedHeaders.push(header);
                }
            });
            
            trace.phase1.matchedDictionaryHeaders = matchedHeaders;
            trace.phase1.unmatchedHeaders = unmatchedHeaders;
            trace.phase1.dictionaryMetadata = metadataSnapshot;
            
            console.log(`[Trace] Matched dictionary headers: ${matchedHeaders.length}`);
            matchedHeaders.forEach(h => {
                const m = metadataSnapshot[h];
                console.log(`   ${h}: side=${m.side}, bucket=${m.bucket}, include=${m.include}, sign=${m.sign}`);
            });
            
            console.log(`[Trace] Unmatched headers (NOT in dictionary): ${unmatchedHeaders.length}`);
            unmatchedHeaders.forEach(h => console.log(`   ${h}`));
            
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 1D: Inclusion filter result (the actual drop)
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 1D: Inclusion Filter Decisions");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            const perColumnDecisions = [];
            const includedColumns = [];
            const excludedColumns = [];
            const excludeReasonCounts = {};
            
            headers.forEach((header, colIdx) => {
                const headerLower = header.toLowerCase();
                const sum = columnSums[header] || 0;
                const nonZeroCount = columnNonZeroCounts[header] || 0;
                const isNumeric = numericCandidates.some(c => c.header === header);
                const isDimension = dimensionHeaders.includes(header);
                const meta = metadataSnapshot[header] || null;
                
                let includeDecision = "EXCLUDE";
                let excludeReason = "unknown";
                
                if (isDimension) {
                    excludeReason = "dimension";
                } else if (!isNumeric) {
                    excludeReason = "non_numeric";
                } else if (EXPENSE_REVIEW_SUMMARY_EXCLUSIONS.has(headerLower)) {
                    excludeReason = "summary_exclusion";
                } else if (meta) {
                    // Apply inclusion rules
                    const inclusionResult = shouldIncludeInExpenseReview(meta, headerLower);
                    if (inclusionResult.include) {
                        includeDecision = "INCLUDE";
                        excludeReason = null;
                    } else {
                        excludeReason = inclusionResult.reason || "unknown_meta_rule";
                    }
                } else {
                    // No dictionary metadata - use default inclusion (permissive)
                    // Check if it looks like an amount column
                    if (headerLower.includes("amount")) {
                        includeDecision = "INCLUDE";
                        excludeReason = null;
                    } else if (headerLower.includes("rate") || headerLower.includes("percent") || headerLower.includes("pct")) {
                        excludeReason = "rate_or_percent_column";
                    } else {
                        // Default permissive for numeric columns
                        includeDecision = "INCLUDE";
                        excludeReason = null;
                    }
                }
                
                const decision = {
                    header,
                    colIdx,
                    isNumeric,
                    isDimension,
                    matchedDictionary: !!meta,
                    side: meta?.side || null,
                    bucket: meta?.bucket || null,
                    includeFlag: meta?.include,
                    includeDecision,
                    excludeReason,
                    sum,
                    absSum: Math.abs(sum),
                    nonZeroCount
                };
                
                perColumnDecisions.push(decision);
                
                if (includeDecision === "INCLUDE") {
                    includedColumns.push(decision);
                } else {
                    excludedColumns.push(decision);
                    excludeReasonCounts[excludeReason] = (excludeReasonCounts[excludeReason] || 0) + 1;
                }
            });
            
            trace.phase2.perColumnDecisions = perColumnDecisions;
            trace.phase2.includedColumns = includedColumns;
            trace.phase2.excludedColumns = excludedColumns;
            trace.phase2.excludeReasonCounts = excludeReasonCounts;
            
            // Sort by absolute sum for reporting
            const topIncluded = [...includedColumns].sort((a, b) => b.absSum - a.absSum).slice(0, 15);
            const topExcluded = [...excludedColumns].sort((a, b) => b.absSum - a.absSum).slice(0, 15);
            
            console.log(`\n[Trace] INCLUDED columns: ${includedColumns.length}`);
            console.log(`[Trace] Top 15 INCLUDED by dollars:`);
            topIncluded.forEach((c, i) => {
                console.log(`   ${i+1}. ${c.header}: $${c.sum.toLocaleString()} (dict=${c.matchedDictionary}, side=${c.side}, bucket=${c.bucket})`);
            });
            
            console.log(`\n[Trace] EXCLUDED columns: ${excludedColumns.length}`);
            console.log(`[Trace] Exclusion reason breakdown:`);
            Object.entries(excludeReasonCounts).sort((a, b) => b[1] - a[1]).forEach(([reason, count]) => {
                console.log(`   ${reason}: ${count} columns`);
            });
            
            console.log(`\n[Trace] Top 15 EXCLUDED by dollars:`);
            topExcluded.forEach((c, i) => {
                console.log(`   ${i+1}. ${c.header}: $${c.sum.toLocaleString()} (reason=${c.excludeReason}, side=${c.side})`);
            });
            
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 2: Totals That Must Always Reconcile
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 2: Reconciling Totals");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            // TOTAL_NUMERIC_SHEET: Sum of all numeric columns excluding dimensions
            const totalNumericSheet = numericCandidates
                .filter(c => !dimensionHeaders.includes(c.header))
                .reduce((sum, c) => sum + c.sum, 0);
            
            // TOTAL_AFTER_METADATA_JOIN: Sum of columns that matched dictionary (no include/side filtering)
            const totalAfterMetadataJoin = numericCandidates
                .filter(c => matchedHeaders.includes(c.header))
                .reduce((sum, c) => sum + c.sum, 0);
            
            // TOTAL_AFTER_INCLUSION_RULES: Sum of included columns only
            const totalAfterInclusionRules = includedColumns.reduce((sum, c) => sum + c.sum, 0);
            
            // TOTAL_BURDEN_ONLY: Sum of columns bucketed as BURDEN
            const totalBurdenOnly = includedColumns
                .filter(c => c.bucket?.toUpperCase() === "BURDEN")
                .reduce((sum, c) => sum + c.sum, 0);
            
            // Calculate deltas
            const delta = totalNumericSheet - totalAfterInclusionRules;
            const excludedSum = excludedColumns.reduce((sum, c) => sum + c.sum, 0);
            
            trace.totals = {
                TOTAL_NUMERIC_SHEET: totalNumericSheet,
                TOTAL_AFTER_METADATA_JOIN: totalAfterMetadataJoin,
                TOTAL_AFTER_INCLUSION_RULES: totalAfterInclusionRules,
                TOTAL_BURDEN_ONLY: totalBurdenOnly,
                delta,
                excludedSum
            };
            
            const fmt = n => `$${n.toLocaleString(undefined, { minimumFractionDigits: 2 })}`;
            
            console.log(`\n[Trace] RECONCILING TOTALS:`);
            console.log(`   1. TOTAL_NUMERIC_SHEET (all numeric, excl dimensions): ${fmt(totalNumericSheet)}`);
            console.log(`   2. TOTAL_AFTER_METADATA_JOIN (matched dictionary):     ${fmt(totalAfterMetadataJoin)}`);
            console.log(`   3. TOTAL_AFTER_INCLUSION_RULES (final included):       ${fmt(totalAfterInclusionRules)}`);
            console.log(`   4. TOTAL_BURDEN_ONLY (bucket=BURDEN):                   ${fmt(totalBurdenOnly)}`);
            console.log(`\n[Trace] DELTA ANALYSIS:`);
            console.log(`   Delta (NUMERIC_SHEET - INCLUSION_RULES): ${fmt(delta)}`);
            console.log(`   Sum of excluded columns:                 ${fmt(excludedSum)}`);
            console.log(`   Reconciles: ${Math.abs(delta - excludedSum) < 1 ? "✅ YES" : "❌ NO (gap: " + fmt(delta - excludedSum) + ")"}`);
            
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 3: Hypothesis Tests
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 3: Hypothesis Tests");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            // Hypothesis A: Rate vs Amount confusion
            const rateColumns = headers.filter(h => h.toLowerCase().includes("rate"));
            const amountColumns = headers.filter(h => h.toLowerCase().includes("amount"));
            
            console.log(`\n[Hypothesis A] Rate vs Amount columns:`);
            console.log(`   *Rate* columns: ${rateColumns.length} → ${rateColumns.slice(0, 10).join(", ")}`);
            console.log(`   *Amount* columns: ${amountColumns.length} → ${amountColumns.slice(0, 10).join(", ")}`);
            
            const rateSum = rateColumns.reduce((sum, h) => sum + (columnSums[h] || 0), 0);
            const amountSum = amountColumns.reduce((sum, h) => sum + (columnSums[h] || 0), 0);
            console.log(`   Rate columns sum: ${fmt(rateSum)}`);
            console.log(`   Amount columns sum: ${fmt(amountSum)}`);
            
            const rateColumnsIncluded = includedColumns.filter(c => c.header.toLowerCase().includes("rate"));
            const amountColumnsIncluded = includedColumns.filter(c => c.header.toLowerCase().includes("amount"));
            console.log(`   Rate columns INCLUDED: ${rateColumnsIncluded.length}`);
            console.log(`   Amount columns INCLUDED: ${amountColumnsIncluded.length}`);
            
            // Hypothesis B: Dimension over-detection
            const suspectDimensions = dimensionHeaders.filter(h => 
                h.toLowerCase().includes("amount") || 
                h.toLowerCase().includes("pay") ||
                h.toLowerCase().includes("wage") ||
                h.toLowerCase().includes("salary")
            );
            console.log(`\n[Hypothesis B] Dimension over-detection:`);
            console.log(`   Suspect dimensions (should be measures): ${suspectDimensions.length}`);
            suspectDimensions.forEach(h => console.log(`   ⚠️ ${h} flagged as dimension but looks like a measure`));
            
            // Hypothesis C: Dictionary join mismatch
            const joinMatchRate = measureCandidates.length > 0 
                ? (matchedHeaders.length / measureCandidates.length * 100).toFixed(1)
                : 0;
            console.log(`\n[Hypothesis C] Dictionary join success rate: ${joinMatchRate}%`);
            console.log(`   ${matchedHeaders.length} matched, ${unmatchedHeaders.length} unmatched`);
            if (unmatchedHeaders.length > 0) {
                console.log(`   Sample unmatched (check normalization):`);
                unmatchedHeaders.slice(0, 10).forEach(h => {
                    const normalized = h.toLowerCase().trim();
                    console.log(`     "${h}" → normalized: "${normalized}"`);
                });
            }
            
            // Hypothesis D: Inclusion filter excluding too much
            const excludedByMissingMeta = excludedColumns.filter(c => !c.matchedDictionary && c.isNumeric && !c.isDimension);
            const excludedBySide = excludedColumns.filter(c => c.excludeReason?.includes("side"));
            console.log(`\n[Hypothesis D] Inclusion filter analysis:`);
            console.log(`   Excluded by missing metadata: ${excludedByMissingMeta.length} columns, ${fmt(excludedByMissingMeta.reduce((s, c) => s + c.sum, 0))}`);
            console.log(`   Excluded by side (ee/na): ${excludedBySide.length} columns, ${fmt(excludedBySide.reduce((s, c) => s + c.sum, 0))}`);
            
            // ═══════════════════════════════════════════════════════════════════════
            // PHASE 4: Export Debug Report to Hidden Sheet
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("PHASE 4: Export Debug Report");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            try {
                // Get or create debug sheet
                let debugSheet = context.workbook.worksheets.getItemOrNullObject("PF_Debug_ExpenseReview");
                debugSheet.load("isNullObject");
                await context.sync();
                
                if (!debugSheet.isNullObject) {
                    debugSheet.delete();
                    await context.sync();
                }
                
                debugSheet = context.workbook.worksheets.add("PF_Debug_ExpenseReview");
                debugSheet.visibility = Excel.SheetVisibility.hidden;
                
                // Build debug table
                const debugHeaders = [
                    "header", "isNumericCandidate", "isDimension", "matchedDictionary",
                    "side", "bucket", "includeFlag", "includeDecision", "excludeReason",
                    "sum", "absSum", "nonZeroCount"
                ];
                
                const debugData = perColumnDecisions
                    .sort((a, b) => b.absSum - a.absSum)
                    .map(d => [
                        d.header,
                        d.isNumeric ? "TRUE" : "FALSE",
                        d.isDimension ? "TRUE" : "FALSE",
                        d.matchedDictionary ? "TRUE" : "FALSE",
                        d.side || "",
                        d.bucket || "",
                        d.includeFlag === undefined ? "" : String(d.includeFlag),
                        d.includeDecision,
                        d.excludeReason || "",
                        d.sum,
                        d.absSum,
                        d.nonZeroCount
                    ]);
                
                const allDebugData = [debugHeaders, ...debugData];
                const debugRange = debugSheet.getRange(`A1:L${allDebugData.length}`);
                debugRange.values = allDebugData;
                
                // Format header row
                const headerRow = debugSheet.getRange("A1:L1");
                formatSheetHeaders(headerRow);
                
                // Add summary section
                const summaryStart = allDebugData.length + 3;
                const summaryRange = debugSheet.getRange(`A${summaryStart}:C${summaryStart + 10}`);
                summaryRange.values = [
                    ["SUMMARY", "", ""],
                    ["", "", ""],
                    ["TOTAL_NUMERIC_SHEET", totalNumericSheet, ""],
                    ["TOTAL_AFTER_METADATA_JOIN", totalAfterMetadataJoin, ""],
                    ["TOTAL_AFTER_INCLUSION_RULES", totalAfterInclusionRules, ""],
                    ["TOTAL_BURDEN_ONLY", totalBurdenOnly, ""],
                    ["", "", ""],
                    ["DELTA (excluded)", delta, ""],
                    ["SUM OF EXCLUDED COLUMNS", excludedSum, ""],
                    ["RECONCILES?", Math.abs(delta - excludedSum) < 1 ? "YES" : "NO", ""]
                ];
                
                await context.sync();
                console.log(`[Trace] ✓ Debug sheet created: PF_Debug_ExpenseReview (hidden)`);
                console.log(`[Trace] ✓ To view: Right-click sheet tabs → Unhide → PF_Debug_ExpenseReview`);
                
            } catch (sheetError) {
                console.warn("[Trace] Could not create debug sheet:", sheetError);
            }
            
            // ═══════════════════════════════════════════════════════════════════════
            // TASK 1 & 2: WAGES SPOTLIGHT - Focused verification of wage columns
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("WAGES SPOTLIGHT: Focused verification of wage/salary columns");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            // Task 1: Check specific wage columns
            const wagePatterns = ["wages", "salary", "regular_pay", "pay_amount", "hourly", "base_pay"];
            const wageHeaders = headers.filter(h => {
                const lower = h.toLowerCase();
                return wagePatterns.some(p => lower.includes(p));
            });
            
            console.log(`\n[WagesCheck] ALL headers (${headers.length} total):`);
            headers.forEach((h, i) => console.log(`   [${i}] ${h}`));
            
            // Specific checks
            const hasWagesSalary = headers.some(h => h.toLowerCase() === "wages_salary_amount");
            const hasRegularPay = headers.some(h => h.toLowerCase() === "regular_pay_amount");
            
            const wagesSalarySum = perColumnDecisions.find(c => c.header.toLowerCase() === "wages_salary_amount")?.sum || 0;
            const regularPaySum = perColumnDecisions.find(c => c.header.toLowerCase() === "regular_pay_amount")?.sum || 0;
            
            console.log(`\n[WagesCheck] hasWagesSalary=${hasWagesSalary}, sumWagesSalary=${fmt(wagesSalarySum)}`);
            console.log(`[WagesCheck] hasRegularPay=${hasRegularPay}, sumRegularPay=${fmt(regularPaySum)}`);
            
            // Find closest matches if not found
            if (!hasWagesSalary) {
                const closestMatches = headers
                    .filter(h => h.toLowerCase().includes("wage") || h.toLowerCase().includes("salary"))
                    .slice(0, 5);
                console.log(`[WagesCheck] "Wages_Salary_Amount" NOT FOUND. Closest matches: ${closestMatches.join(", ") || "(none)"}`);
            }
            if (!hasRegularPay) {
                const closestMatches = headers
                    .filter(h => h.toLowerCase().includes("regular") || h.toLowerCase().includes("pay"))
                    .slice(0, 5);
                console.log(`[WagesCheck] "Regular_Pay_Amount" NOT FOUND. Closest matches: ${closestMatches.join(", ") || "(none)"}`);
            }
            
            // Task 2: Spotlight on all wage-like columns
            console.log(`\n[WagesSpotlight] All wage-like columns (${wageHeaders.length} found):`);
            wageHeaders.forEach(h => {
                const decision = perColumnDecisions.find(c => c.header === h);
                if (decision) {
                    console.log(`   ${h}: sum=${fmt(decision.sum)}, isNumeric=${decision.isNumeric}, isDim=${decision.isDimension}, dictMatch=${decision.matchedDictionary}, side=${decision.side}, bucket=${decision.bucket}, include=${decision.includeDecision}, reason=${decision.excludeReason || "(none)"}`);
                    
                    // ALERT for high-dollar excluded wage columns
                    if (decision.sum > 10000 && decision.includeDecision !== "INCLUDE") {
                        console.error(`   ⚠️ [WagesSpotlight][ALERT] High-dollar wage column EXCLUDED: ${h} reason=${decision.excludeReason} meta={side:${decision.side}, bucket:${decision.bucket}, include:${decision.includeFlag}}`);
                    }
                }
            });
            
            // Check what's actually being included in each bucket (using normalized names)
            const fixedColumns = includedColumns.filter(c => normalizeBucketName(c.bucket) === "FIXED");
            const variableColumns = includedColumns.filter(c => normalizeBucketName(c.bucket) === "VARIABLE");
            const burdenColumns = includedColumns.filter(c => normalizeBucketName(c.bucket) === "BURDEN");
            const benefitColumns = includedColumns.filter(c => normalizeBucketName(c.bucket) === "BENEFIT");
            const otherColumns = includedColumns.filter(c => normalizeBucketName(c.bucket) === "OTHER");
            
            console.log(`\n[BucketBreakdown] Columns by bucket (INCLUDED only):`);
            console.log(`   FIXED (${fixedColumns.length}): ${fmt(fixedColumns.reduce((s, c) => s + c.sum, 0))} → ${fixedColumns.map(c => c.header).join(", ") || "(none)"}`);
            console.log(`   VARIABLE (${variableColumns.length}): ${fmt(variableColumns.reduce((s, c) => s + c.sum, 0))} → ${variableColumns.map(c => c.header).join(", ") || "(none)"}`);
            console.log(`   BURDEN (${burdenColumns.length}): ${fmt(burdenColumns.reduce((s, c) => s + c.sum, 0))} → ${burdenColumns.map(c => c.header).join(", ") || "(none)"}`);
            console.log(`   BENEFITS (${benefitColumns.length}): ${fmt(benefitColumns.reduce((s, c) => s + c.sum, 0))} → ${benefitColumns.map(c => c.header).join(", ") || "(none)"}`);
            console.log(`   OTHER/NULL (${otherColumns.length}): ${fmt(otherColumns.reduce((s, c) => s + c.sum, 0))} → ${otherColumns.map(c => c.header + ":" + c.bucket).join(", ") || "(none)"}`);
            
            // Store bucket breakdown for UI debug
            trace.bucketBreakdown = {
                FIXED: { count: fixedColumns.length, sum: fixedColumns.reduce((s, c) => s + c.sum, 0), columns: fixedColumns.map(c => c.header) },
                VARIABLE: { count: variableColumns.length, sum: variableColumns.reduce((s, c) => s + c.sum, 0), columns: variableColumns.map(c => c.header) },
                BURDEN: { count: burdenColumns.length, sum: burdenColumns.reduce((s, c) => s + c.sum, 0), columns: burdenColumns.map(c => c.header) },
                BENEFITS: { count: benefitColumns.length, sum: benefitColumns.reduce((s, c) => s + c.sum, 0), columns: benefitColumns.map(c => c.header) },
                OTHER: { count: otherColumns.length, sum: otherColumns.reduce((s, c) => s + c.sum, 0), columns: otherColumns.map(c => c.header) }
            };
            
            // ═══════════════════════════════════════════════════════════════════════
            // TASK 3: Compare with what parseExpenseRows actually produces
            // ═══════════════════════════════════════════════════════════════════════
            console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            console.log("TASK 3: UI Pipeline Verification (parseExpenseRows output)");
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            
            // Actually call parseExpenseRows to see what it produces
            const parseResult = parseExpenseRows(allValues);
            
            let uiFixed = 0, uiVariable = 0, uiBurden = 0;
            parseResult.forEach(row => {
                uiFixed += row.fixed || 0;
                uiVariable += row.variable || 0;
                uiBurden += row.burden || 0;
            });
            const uiTotal = uiFixed + uiVariable + uiBurden;
            
            console.log(`[UIVerify] parseExpenseRows produced ${parseResult.length} rows`);
            console.log(`[UIVerify] UI will show:`);
            console.log(`   FIXED: ${fmt(uiFixed)}`);
            console.log(`   VARIABLE: ${fmt(uiVariable)}`);
            console.log(`   BURDEN: ${fmt(uiBurden)}`);
            console.log(`   TOTAL: ${fmt(uiTotal)}`);
            
            console.log(`\n[UIVerify] COMPARISON:`);
            console.log(`   Trace TOTAL_AFTER_INCLUSION_RULES: ${fmt(totalAfterInclusionRules)}`);
            console.log(`   UI TOTAL (fixed+variable+burden):  ${fmt(uiTotal)}`);
            console.log(`   MATCH: ${Math.abs(totalAfterInclusionRules - uiTotal) < 1 ? "✅ YES" : "❌ NO - DISCREPANCY FOUND!"}`);
            
            if (Math.abs(totalAfterInclusionRules - uiTotal) >= 1) {
                const discrepancy = totalAfterInclusionRules - uiTotal;
                console.error(`\n[UIVerify] ⚠️ DISCREPANCY: ${fmt(discrepancy)}`);
                console.error(`[UIVerify] This means the trace computation and parseExpenseRows use DIFFERENT logic!`);
                
                // Deep dive: what columns does parseExpenseRows classify as measures?
                console.log(`\n[UIVerify] Investigating parseExpenseRows classification...`);
            }
            
            // Store for trace result
            trace.uiTotals = { fixed: uiFixed, variable: uiVariable, burden: uiBurden, total: uiTotal };
            trace.discrepancy = totalAfterInclusionRules - uiTotal;
            
            console.log("\n╔══════════════════════════════════════════════════════════════════════════╗");
            console.log("║  TRACE COMPLETE                                                          ║");
            console.log("╚══════════════════════════════════════════════════════════════════════════╝\n");
        });
        
        return trace;
        
    } catch (error) {
        console.error("[Trace] Error during trace:", error);
        return trace;
    }
}

// Make trace function available globally for console access
if (typeof window !== 'undefined') {
    window.traceExpenseReviewMeasurePipeline = traceExpenseReviewMeasurePipeline;
}

/**
 * Parse expense rows - UNIFIED WITH STEP 1 MEASURE UNIVERSE
 * 
 * CRITICAL: This function uses the SAME measure list as Step 1 (getPRDataCleanMeasureUniverse)
 * Dictionary metadata is applied AFTER for bucket labels/presentation ONLY
 * 
 * DATA SOURCE:
 * - Input: PR_Data_Clean values (current period) 
 * - measureUniverse: pre-computed from getPRDataCleanMeasureUniverse()
 * 
 * FLOW:
 * 1. Use measureUniverse.includedMeasureHeaders as THE authoritative measure list
 * 2. Look up dictionary metadata for each header (bucket, sign, label)
 * 3. If not in dictionary -> bucket = "UNCLASSIFIED" but STILL INCLUDED
 * 4. Aggregate by bucket for display, total must match universe.total
 */
function parseExpenseRows(values, measureUniverse = null) {
    if (!values || values.length < 2) return [];
    
    console.log("[parseExpenseRows] ═══════════════════════════════════════════════════════");
    console.log("[parseExpenseRows] Using UNIFIED measure universe from Step 1");
    
    // Get raw headers
    const rawHeaders = values[0] || [];
    const headers = rawHeaders.map((h) => String(h || "").trim());
    
    // If no universe provided, create a simple one from headers (fallback)
    if (!measureUniverse) {
        console.warn("[parseExpenseRows] No measure universe provided - using fallback");
        measureUniverse = {
            includedMeasureHeaders: headers.filter(h => 
                h.toLowerCase().includes("amount") && 
                !MEASURE_UNIVERSE_EXCLUSIONS.has(h.toLowerCase())
            ),
            perColumnSums: {},
            total: 0
        };
    }
    
    // Build header index map
    const headerIndexMap = {};
    headers.forEach((h, idx) => {
        headerIndexMap[h] = idx;
    });
    
    // Get customer mappings from state (loaded in prepareExpenseReviewData)
    const customerMappings = expenseReviewState.customerMappings || [];
    
    // Build measure metadata using customer mappings (preferred) or dictionary (fallback)
    const measureMetadata = {};
    const unclassifiedHeaders = [];
    
    measureUniverse.includedMeasureHeaders.forEach(header => {
        const headerLower = header.toLowerCase();
        
        // Priority 1: Customer mapping expense_bucket
        const customerMapping = customerMappings.find(m => 
            m.pf_column_name?.toLowerCase() === headerLower ||
            m.raw_header?.toLowerCase() === headerLower
        );
        
        if (customerMapping?.expense_bucket) {
            // Use customer mapping bucket (normalized to uppercase)
            const bucket = customerMapping.expense_bucket.toUpperCase();
            // Get sign from dictionary if available
            const dictEntry = expenseTaxonomyCache?.measures?.[headerLower];
            measureMetadata[header] = {
                bucket: bucket,
                sign: dictEntry?.sign ?? 1,
                label: dictEntry?.label || header,
                source: "customer_mapping"
            };
        } else {
            // Priority 2: Dictionary metadata (fallback)
            const dictEntry = expenseTaxonomyCache?.measures?.[headerLower];
            
            if (dictEntry) {
                measureMetadata[header] = {
                    bucket: normalizeBucketName(dictEntry.bucket),
                    sign: dictEntry.sign ?? 1,
                    label: dictEntry.label || header,
                    source: "dictionary"
                };
            } else {
                // No customer mapping or dictionary entry - UNCLASSIFIED
                measureMetadata[header] = {
                    bucket: "UNCLASSIFIED",
                    sign: 1,
                    label: header,
                    source: "none"
                };
                unclassifiedHeaders.push(header);
            }
        }
    });
    
    // Log classification summary
    const bucketCounts = {};
    const sourceCounts = { customer_mapping: 0, dictionary: 0, none: 0 };
    Object.values(measureMetadata).forEach(m => {
        bucketCounts[m.bucket] = (bucketCounts[m.bucket] || 0) + 1;
        sourceCounts[m.source] = (sourceCounts[m.source] || 0) + 1;
    });
    console.log("[parseExpenseRows] Bucket distribution:", bucketCounts);
    console.log("[parseExpenseRows] Classification sources:", sourceCounts);
    
    if (unclassifiedHeaders.length > 0) {
        console.warn("[parseExpenseRows] UNCLASSIFIED columns (included in totals):", unclassifiedHeaders);
    }
    
    // Store classification info in state for UI display
    expenseReviewState.unclassifiedColumns = unclassifiedHeaders;
    expenseReviewState.measureColumns = measureUniverse.includedMeasureHeaders.map(h => ({
        header: h,
        bucket: measureMetadata[h]?.bucket,
        source: measureMetadata[h]?.source,
        include: true // Always true - universe is authoritative
    }));
    expenseReviewState.measureUniverse = measureUniverse; // Store for UI reconciliation
    
    // Find key dimension columns (Employee, Department, Pay_Date)
    const employeeIdx = headers.findIndex(h => {
        const hl = h.toLowerCase();
        return hl === "employee_name" || hl === "employee_id" || (hl.includes("employee") && !hl.includes("amount"));
    });
    
    // Use pickDepartmentIndex for consistent department name preference
    // Priority: Department Description > Department Name > Department (non-code)
    const departmentIdx = pickDepartmentIndex(headers);
    
    // Check if data has Pay_Date column (archive has it, PR_Data_Clean doesn't)
    const payDateIdx = headers.findIndex(h => {
        const hl = h.toLowerCase();
        return hl === "pay_date" || hl === "payroll_date" || (hl.includes("payroll") && hl.includes("date"));
    });
    
    // Get period from config as fallback (for PR_Data_Clean which has no Pay_Date column)
    const rawConfigPeriod = getPayrollDateValue();
    const configPeriod = rawConfigPeriod ? normalizePeriodKey(rawConfigPeriod) : "";
    
    const hasPayDateColumn = payDateIdx >= 0;
    
    console.log("[parseExpenseRows] Column indexes:", {
        employee: employeeIdx,
        department: departmentIdx,
        payDate: payDateIdx,
        hasPayDateColumn
    });
    
    if (hasPayDateColumn) {
        console.log("[parseExpenseRows] Using Pay_Date from row data (archive mode)");
    } else {
        console.log("[parseExpenseRows] Using config Pay_Date (current period mode):", {
            raw: rawConfigPeriod,
            rawType: typeof rawConfigPeriod,
            normalized: configPeriod
        });
        if (!configPeriod) {
            console.error("[parseExpenseRows] No Pay_Date configured in SS_PF_Config!");
        }
    }
    
    // Build measure columns with indexes
    const measureColumns = measureUniverse.includedMeasureHeaders.map(header => ({
        header,
        index: headerIndexMap[header],
        metadata: measureMetadata[header]
    })).filter(c => c.index !== undefined);
    
    console.log(`[parseExpenseRows] Processing ${measureColumns.length} measure columns from universe`);
    
    // Process rows - read Pay_Date from row if available, otherwise use config
    const rows = [];
    const fallbackPeriod = configPeriod || new Date().toISOString().split("T")[0];
    let skippedCount = 0;
    const samplePayDates = [];
    
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        
        // Determine period for this row
        let period;
        if (hasPayDateColumn) {
            // Archive mode: read Pay_Date from row data
            const rawRowDate = row[payDateIdx];
            
            // Collect sample Pay_Date values for diagnostics (first 5)
            if (samplePayDates.length < 5) {
                samplePayDates.push({ row: i, raw: rawRowDate, type: typeof rawRowDate });
            }
            
            const parsedDate = parsePeriodDate(rawRowDate);
            period = parsedDate ? formatDateFromDate(parsedDate) : normalizePeriodKey(rawRowDate);
            if (!period) {
                skippedCount++;
                if (skippedCount <= 3) {
                    console.warn(`[parseExpenseRows] Row ${i} has invalid Pay_Date, skipping:`, rawRowDate, typeof rawRowDate);
                }
                continue;
            }
        } else {
            // Current period mode: use config Pay_Date for all rows
            period = fallbackPeriod;
        }
        
        const employee = employeeIdx >= 0 ? normalizeString(row[employeeIdx]) : "";
        const department = departmentIdx >= 0 ? normalizeString(row[departmentIdx]) || "Unassigned" : "Unassigned";
        
        // Aggregate measures by bucket (including UNCLASSIFIED)
        const bucketTotals = {
            FIXED: 0,
            VARIABLE: 0,
            BURDEN: 0,
            BENEFITS: 0,   // NEW: Customer mapping bucket
            BENEFIT: 0,    // Legacy dictionary bucket
            TAX: 0,
            DEDUCTION: 0,
            REIMBURSEMENT: 0,
            OTHER: 0,
            UNCLASSIFIED: 0  // Track unclassified separately
        };
        
        measureColumns.forEach(col => {
            const value = toNumber(row[col.index]);
            const sign = col.metadata?.sign ?? 1;
            const bucket = col.metadata?.bucket || "UNCLASSIFIED";
            bucketTotals[bucket] = (bucketTotals[bucket] || 0) + (value * sign);
        });
        
        // Map buckets to legacy fixed/variable/burden for backward compatibility
        // CRITICAL: Include UNCLASSIFIED in burden to ensure total reconciles
        // BENEFITS (new) and BENEFIT (legacy) both map to burden
        const fixed = (bucketTotals.FIXED || 0) + (bucketTotals.REIMBURSEMENT || 0);
        const variable = bucketTotals.VARIABLE || 0;
        const burden = (bucketTotals.BURDEN || 0) + 
                       (bucketTotals.BENEFITS || 0) +   // NEW: Customer mapping bucket
                       (bucketTotals.BENEFIT || 0) + 
                       (bucketTotals.TAX || 0) + 
                       (bucketTotals.OTHER || 0) + 
                       (bucketTotals.UNCLASSIFIED || 0);  // UNCLASSIFIED goes to burden
        
        // Skip rows with no meaningful amounts
        if (fixed === 0 && variable === 0 && burden === 0) continue;
        
        rows.push({
            period,
            employee,
            department: department || "Unassigned",
            fixed,
            variable,
            burden,
            bucketTotals
        });
    }
    
    // Calculate totals for verification
    let totalFixed = 0, totalVariable = 0, totalBurden = 0;
    rows.forEach(r => {
        totalFixed += r.fixed;
        totalVariable += r.variable;
        totalBurden += r.burden;
    });
    const expenseReviewTotal = totalFixed + totalVariable + totalBurden;
    
    // DIAGNOSTIC: Show distinct periods parsed
    const distinctPeriods = [...new Set(rows.map(r => r.period))];
    
    console.log("[parseExpenseRows] ───────────────────────────────────────────────────────");
    console.log(`[parseExpenseRows] Processed ${rows.length} rows from ${values.length - 1} input rows`);
    if (hasPayDateColumn) {
        console.log(`[parseExpenseRows] ARCHIVE MODE - Skipped ${skippedCount} rows with invalid Pay_Date`);
        console.log(`[parseExpenseRows] Sample Pay_Date values:`, samplePayDates);
        console.log(`[parseExpenseRows] Distinct periods found: ${distinctPeriods.length}`, distinctPeriods);
    } else {
        console.log(`[parseExpenseRows] CURRENT MODE - All rows assigned period: ${fallbackPeriod}`);
    }
    console.log(`[parseExpenseRows] FIXED:    $${totalFixed.toLocaleString(undefined, { minimumFractionDigits: 2 })}`);
    console.log(`[parseExpenseRows] VARIABLE: $${totalVariable.toLocaleString(undefined, { minimumFractionDigits: 2 })}`);
    console.log(`[parseExpenseRows] BURDEN:   $${totalBurden.toLocaleString(undefined, { minimumFractionDigits: 2 })}`);
    console.log(`[parseExpenseRows] TOTAL:    $${expenseReviewTotal.toLocaleString(undefined, { minimumFractionDigits: 2 })}`);
    console.log(`[parseExpenseRows] Universe: $${measureUniverse.total.toLocaleString(undefined, { minimumFractionDigits: 2 })}`);
    
    const delta = Math.abs(expenseReviewTotal - measureUniverse.total);
    if (delta > 1) {
        console.error(`[parseExpenseRows] ⚠️ RECONCILIATION FAILED! Delta: $${delta.toFixed(2)}`);
    } else {
        console.log(`[parseExpenseRows] ✅ RECONCILED with Step 1 universe (delta: $${delta.toFixed(2)})`);
    }
    console.log("[parseExpenseRows] ═══════════════════════════════════════════════════════");
    
    return rows;
}

function aggregateExpensePeriods(rows) {
    const map = new Map();
    rows.forEach((row) => {
        const key = row.period;
        if (!key) return;
        if (!map.has(key)) {
            map.set(key, {
                key,
                label: formatFriendlyPeriod(key),
                employees: new Set(),
                departments: new Map(),
                summary: { fixed: 0, variable: 0, burden: 0 }
            });
        }
        const bucket = map.get(key);
        bucket.employees.add(row.employee || `EMP-${bucket.employees.size + 1}`);
        const deptKey = row.department || "Unassigned";
        if (!bucket.departments.has(deptKey)) {
            bucket.departments.set(deptKey, {
                name: deptKey,
                fixed: 0,
                variable: 0,
                burden: 0,
                employees: new Set()
            });
        }
        const dept = bucket.departments.get(deptKey);
        dept.fixed += row.fixed;
        dept.variable += row.variable;
        dept.burden += row.burden;
        dept.employees.add(row.employee || `EMP-${dept.employees.size + 1}`);
        bucket.summary.fixed += row.fixed;
        bucket.summary.variable += row.variable;
        bucket.summary.burden += row.burden;
    });
    const result = [];
    map.forEach((bucket) => {
        const total = bucket.summary.fixed + bucket.summary.variable + bucket.summary.burden;
        const departments = Array.from(bucket.departments.values()).map((dept) => {
            const gross = dept.fixed + dept.variable;
            const allIn = gross + dept.burden;
            return {
                name: dept.name,
                fixed: dept.fixed,
                variable: dept.variable,
                gross,
                burden: dept.burden,
                allIn,
                percent: total ? allIn / total : 0,
                headcount: dept.employees.size,
                delta: 0
            };
        });
        departments.sort((a, b) => b.allIn - a.allIn);
        const summary = {
            employeeCount: bucket.employees.size,
            fixed: bucket.summary.fixed,
            variable: bucket.summary.variable,
            burden: bucket.summary.burden,
            total,
            burdenRate: total ? bucket.summary.burden / total : 0,
            delta: 0
        };
        const totalsRow = {
            name: "Totals",
            fixed: bucket.summary.fixed,
            variable: bucket.summary.variable,
            gross: bucket.summary.fixed + bucket.summary.variable,
            burden: bucket.summary.burden,
            allIn: total,
            percent: total ? 1 : 0,
            headcount: bucket.employees.size,
            delta: 0,
            isTotal: true
        };
        result.push({
            key: bucket.key,
            label: bucket.label,
            summary,
            departments,
            totalsRow
        });
    });
    return result.sort((a, b) => (a.key < b.key ? 1 : -1));
}

/**
 * Extract archive totals by dynamically recomputing from department summary rows.
 * 
 * IMPORTANT: PR_Archive_Summary does NOT persist pre-calculated bucket totals.
 * All totals are derived at read time from:
 * - Raw measure columns (copied from PR_Data_Clean)
 * - Department-level TOTAL rows (copied from PR_Expense_Review)
 * 
 * This function identifies the grand TOTAL row for each period (Row_Type = "TOTAL")
 * and extracts the department summary metrics (Fixed, Variable, Burden, Headcount).
 * 
 * @param {any[][]} archiveValues - Archive sheet values (including header)
 * @returns {Map<string, {fixed: number, variable: number, burden: number, total: number, headcount: number}>}
 */
function extractStoredArchiveTotals(archiveValues) {
    const storedTotals = new Map();
    
    if (!archiveValues || archiveValues.length < 2) {
        return storedTotals;
    }
    
    const headers = archiveValues[0];
    const headersLower = headers.map(h => String(h || "").toLowerCase().trim());
    const deptIdx = headersLower.findIndex(h => 
        h.includes("department") &&
        !h.includes("id") &&
        !h.includes("#") &&
        !h.includes("code") &&
        !h.includes("number")
    );
    
    // Find date column for period key
    const dateIdx = headersLower.findIndex(h => 
        h === "payroll_date" || h === "pay_date" ||
        (h.includes("payroll") && h.includes("date")) || 
        h.includes("pay period") || h === "date"
    );
    
    // Find Row_Type column to identify department summary rows
    const rowTypeIdx = headersLower.findIndex(h => 
        h === "row_type" || h === "rowtype"
    );
    
    // Fallback: Find employee column for backward compatibility with old archives
    const employeeIdx = headersLower.findIndex(h => 
        h.includes("employee") && !h.includes("amount")
    );
    
    // Normalize header for flexible matching
    const normalizeHeader = (h) => String(h || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    
    // Find department summary columns dynamically by normalized name
    const findHeaderIdx = (targetNames) => {
        for (const target of targetNames) {
            const normalized = normalizeHeader(target);
            const idx = headers.findIndex(h => normalizeHeader(h) === normalized);
            if (idx >= 0) return idx;
        }
        return -1;
    };
    
    const fixedSalaryIdx = findHeaderIdx(["Fixed Salary", "Fixed", "Base Salary"]);
    const variableSalaryIdx = findHeaderIdx(["Variable Salary", "Variable", "Variable Pay"]);
    const grossPayIdx = findHeaderIdx(["Gross Pay", "Gross", "Gross Wages"]);
    const burdenIdx = findHeaderIdx(["Burden", "Payroll Burden", "Employer Burden"]);
    const allInTotalIdx = findHeaderIdx(["All-In Total", "All In Total", "Total", "All-In"]);
    const percentOfTotalIdx = findHeaderIdx(["% of Total", "Percent of Total", "Percentage"]);
    const headcountIdx = findHeaderIdx(["Headcount", "Head Count", "Employee Count"]);
    
    // If no department summary columns exist, return empty (legacy archive data)
    if (fixedSalaryIdx < 0 && variableSalaryIdx < 0 && burdenIdx < 0 && allInTotalIdx < 0) {
        console.log("[Archive] No department summary columns found - using calculated values");
        return storedTotals;
    }
    
    console.log("[Archive] Found department summary columns - extracting TOTAL rows", {
        fixedSalaryIdx,
        variableSalaryIdx,
        grossPayIdx,
        burdenIdx,
        allInTotalIdx,
        percentOfTotalIdx,
        headcountIdx
    });
    
    // Extract stored totals by finding "TOTAL" rows in each period
    const dataRows = archiveValues.slice(1);
    const normalize = (s) => String(s || "").trim().toLowerCase();
    
    dataRows.forEach(row => {
        const rawDate = dateIdx >= 0 ? row[dateIdx] : "";
        const parsedDate = parsePeriodDate(rawDate);
        const periodKey = parsedDate 
            ? formatDateFromDate(parsedDate)
            : String(rawDate || "").trim();
        
        // Check if this is a GRAND TOTAL row (not department subtotal)
        // CRITICAL: Must distinguish "TOTAL" from "Sales & Marketing Total", etc.
        let isTotalRow = false;
        let rowTypeValue = "";
        let employeeValue = "";
        
        if (rowTypeIdx >= 0) {
            // Use Row_Type column if available (robust, explicit)
            // FIXED: Check for "TOTAL" (grand total), not "Department_Total" (dept subtotal)
            rowTypeValue = String(row[rowTypeIdx] || "").trim();
            const normRowType = normalize(rowTypeValue);
            isTotalRow = normRowType === "total" || normRowType === "grand_total" || normRowType === "grand total";
        }

        if (!isTotalRow && employeeIdx >= 0) {
            employeeValue = String(row[employeeIdx] || "").trim();
            const normEmployee = normalize(employeeValue);
            isTotalRow = normEmployee === "total" || normEmployee === "total total";
        }
        
        if (periodKey && isTotalRow && !storedTotals.has(periodKey)) {
            const fixed = fixedSalaryIdx >= 0 ? (Number(row[fixedSalaryIdx]) || 0) : 0;
            const variable = variableSalaryIdx >= 0 ? (Number(row[variableSalaryIdx]) || 0) : 0;
            const burden = burdenIdx >= 0 ? (Number(row[burdenIdx]) || 0) : 0;
            const total = allInTotalIdx >= 0 ? (Number(row[allInTotalIdx]) || 0) : 0;
            const headcount = headcountIdx >= 0 ? (Number(row[headcountIdx]) || 0) : 0;
            
            // Only store if we have meaningful data
            if (total > 0 || fixed > 0 || variable > 0 || burden > 0) {
                storedTotals.set(periodKey, {
                    fixed,
                    variable,
                    burden,
                    total,
                    headcount
                });
                
                // DIAGNOSTIC LOG: Confirm we're using the correct TOTAL row
                console.log(`[ArchiveTotals] Using TOTAL row for ${periodKey}:`, {
                    periodKey,
                    rowType: rowTypeValue || "(using fallback)",
                    employeeCell: employeeValue || row[employeeIdx],
                    deptCell: deptIdx >= 0 ? row[deptIdx] : "(no dept col)",
                    headcount,
                    fixed: `$${fixed.toLocaleString()}`,
                    variable: `$${variable.toLocaleString()}`,
                    burden: `$${burden.toLocaleString()}`,
                    allInTotal: `$${total.toLocaleString()}`
                });
            }
        }
    });
    
    return storedTotals;
}

function buildArchivePeriodsFromTotalsSheet(archiveTotalsValues) {
    if (!archiveTotalsValues || archiveTotalsValues.length < 2) {
        return [];
    }

    const headers = archiveTotalsValues[0] || [];
    const headersLower = headers.map(h => String(h || "").toLowerCase().trim());

    const dateIdx = headersLower.findIndex(h => 
        h === "payroll_date" || h === "pay_date" ||
        (h.includes("payroll") && h.includes("date")) ||
        h.includes("pay period") || h === "date"
    );

    const rowTypeIdx = headersLower.findIndex(h => h === "row_type" || h === "rowtype");
    const deptIdx = headersLower.findIndex(h => 
        h.includes("department") &&
        !h.includes("id") &&
        !h.includes("#") &&
        !h.includes("code") &&
        !h.includes("number")
    );

    const normalizeHeader = (h) => String(h || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const findHeaderIdx = (targetNames) => {
        for (const target of targetNames) {
            const normalized = normalizeHeader(target);
            const idx = headers.findIndex(h => normalizeHeader(h) === normalized);
            if (idx >= 0) return idx;
        }
        return -1;
    };

    const fixedSalaryIdx = findHeaderIdx(["Fixed Salary", "Fixed", "Base Salary"]);
    const variableSalaryIdx = findHeaderIdx(["Variable Salary", "Variable", "Variable Pay"]);
    const grossPayIdx = findHeaderIdx(["Gross Pay", "Gross", "Gross Wages"]);
    const burdenIdx = findHeaderIdx(["Burden", "Payroll Burden", "Employer Burden"]);
    const allInTotalIdx = findHeaderIdx(["All-In Total", "All In Total", "Total", "All-In"]);
    const percentOfTotalIdx = findHeaderIdx(["% of Total", "Percent of Total", "Percentage"]);
    const headcountIdx = findHeaderIdx(["Headcount", "Head Count", "Employee Count"]);

    const normalize = (s) => String(s || "").trim().toLowerCase();
    const periodMap = new Map();

    const dataRows = archiveTotalsValues.slice(1);
    dataRows.forEach((row) => {
        const rawDate = dateIdx >= 0 ? row[dateIdx] : "";
        const parsedDate = parsePeriodDate(rawDate);
        const periodKey = parsedDate ? formatDateFromDate(parsedDate) : String(rawDate || "").trim();
        if (!periodKey) return;

        if (!periodMap.has(periodKey)) {
            periodMap.set(periodKey, { totals: null, departments: [] });
        }
        const period = periodMap.get(periodKey);

        const deptName = deptIdx >= 0 ? String(row[deptIdx] || "").trim() : "";
        const rowTypeValue = rowTypeIdx >= 0 ? normalize(row[rowTypeIdx]) : "";

        const fixed = fixedSalaryIdx >= 0 ? (Number(row[fixedSalaryIdx]) || 0) : 0;
        const variable = variableSalaryIdx >= 0 ? (Number(row[variableSalaryIdx]) || 0) : 0;
        const gross = grossPayIdx >= 0 ? (Number(row[grossPayIdx]) || 0) : (fixed + variable);
        const burden = burdenIdx >= 0 ? (Number(row[burdenIdx]) || 0) : 0;
        const total = allInTotalIdx >= 0 ? (Number(row[allInTotalIdx]) || 0) : 0;
        const headcount = headcountIdx >= 0 ? (Number(row[headcountIdx]) || 0) : 0;
        const percent = percentOfTotalIdx >= 0 ? (Number(row[percentOfTotalIdx]) || 0) : 0;

        const isTotalRow = rowTypeValue === "total" || normalize(deptName) === "total";
        if (isTotalRow) {
            period.totals = { fixed, variable, burden, total, headcount };
            return;
        }

        if (!deptName) return;

        period.departments.push({
            name: deptName,
            fixed,
            variable,
            gross,
            burden,
            allIn: total,
            percent,
            headcount,
            delta: 0,
            isTotal: false
        });
    });

    const sortedKeys = Array.from(periodMap.keys()).sort((a, b) => (a < b ? 1 : -1));
    return sortedKeys.map((periodKey) => {
        const period = periodMap.get(periodKey);
        const departments = (period?.departments || []).slice().sort((a, b) => (b.allIn || 0) - (a.allIn || 0));
        const totals = period?.totals || departments.reduce((acc, d) => {
            acc.fixed += d.fixed || 0;
            acc.variable += d.variable || 0;
            acc.burden += d.burden || 0;
            acc.total += d.allIn || 0;
            return acc;
        }, { fixed: 0, variable: 0, burden: 0, total: 0, headcount: 0 });

        const totalAllIn = totals.total || 0;
        const normalizedDepartments = departments.map((d) => {
            if (d.percent && d.percent > 0) return d;
            return {
                ...d,
                percent: totalAllIn ? (d.allIn || 0) / totalAllIn : 0
            };
        });

        return {
            key: periodKey,
            label: formatFriendlyPeriod(periodKey),
            summary: {
                fixed: totals.fixed,
                variable: totals.variable,
                burden: totals.burden,
                total: totals.total,
                employeeCount: totals.headcount || 0,
                burdenRate: totals.total ? totals.burden / totals.total : 0,
                delta: 0
            },
            departments: normalizedDepartments,
            totalsRow: {
                name: "TOTAL",
                fixed: totals.fixed,
                variable: totals.variable,
                gross: totals.fixed + totals.variable,
                burden: totals.burden,
                allIn: totals.total,
                percent: totals.total ? 1 : 0,
                headcount: totals.headcount || 0,
                delta: 0,
                isTotal: true
            }
        };
    });
}

/**
 * Build expense review periods by combining current and archive data
 * 
 * PRIOR PERIOD SOURCE DOCUMENTATION:
 * - Current period data: PR_Data_Clean sheet (canonical pf_column_name headers)
 * - Prior period data: PR_Archive_Summary sheet (stores historical run totals)
 * - Period key: Payroll_Date column (YYYY-MM-DD format)
 * 
 * The archive sheet stores aggregated totals by period key, using the same
 * column structure as PR_Data_Clean. When archiving, the current period's
 * totals are appended to PR_Archive_Summary with the period key.
 * 
 * BUCKET TOTALS:
 * - Archive data includes _Archive_* columns with pre-calculated bucket totals
 * - These preserve the classification at the time of archiving
 * - If available, these are used instead of recalculating (for historical consistency)
 * 
 * Prior period comparisons are keyed by pf_column_name (canonical headers).
 */
 function buildExpenseReviewPeriods(cleanValues, archiveValues, measureUniverse = null, archiveTotalsValues = null) {
    console.log("[buildExpenseReviewPeriods] ═══════════════════════════════════════════════");
    console.log("[buildExpenseReviewPeriods] INPUT DATA:");
    console.log("  - cleanValues rows:", cleanValues?.length || 0);
    console.log("  - archiveValues rows:", archiveValues?.length || 0);
    console.log("  - archiveTotalsValues rows:", archiveTotalsValues?.length || 0);
    console.log("  - measureUniverse provided:", !!measureUniverse);
    
    if (archiveValues && archiveValues.length > 0) {
        console.log("  - Archive headers:", archiveValues[0]?.slice(0, 5), "...");
        console.log("  - Archive sample row 1:", archiveValues[1]?.slice(0, 5), "...");
    }
    
    // Pass measureUniverse to parseExpenseRows for unified totals
    console.log("[buildExpenseReviewPeriods] Parsing current period data...");
    const currentPeriods = aggregateExpensePeriods(parseExpenseRows(cleanValues, measureUniverse));
    
    let archivePeriods = [];
    if (archiveTotalsValues && archiveTotalsValues.length > 1) {
        console.log("[buildExpenseReviewPeriods] Building archive periods from PR_Archive_Totals...");
        archivePeriods = buildArchivePeriodsFromTotalsSheet(archiveTotalsValues);
    } else {
        console.log("[buildExpenseReviewPeriods] Building archive periods from PR_Archive_Summary totals...");
        const storedArchiveTotals = extractStoredArchiveTotals(archiveValues);
        if (storedArchiveTotals.size > 0) {
            archivePeriods = Array.from(storedArchiveTotals.entries()).map(([periodKey, stored]) => {
                console.log(`[Archive] Period ${periodKey}: $${stored.total.toLocaleString()}`);
                return {
                    key: periodKey,
                    label: formatFriendlyPeriod(periodKey),
                    summary: {
                        fixed: stored.fixed,
                        variable: stored.variable,
                        burden: stored.burden,
                        total: stored.total,
                        employeeCount: stored.headcount || 0,
                        burdenRate: stored.total ? stored.burden / stored.total : 0,
                        delta: 0
                    },
                    departments: [],
                    totalsRow: {
                        name: "TOTAL",
                        fixed: stored.fixed,
                        variable: stored.variable,
                        gross: stored.fixed + stored.variable,
                        burden: stored.burden,
                        allIn: stored.total,
                        percent: 1,
                        headcount: stored.headcount || 0,
                        delta: 0,
                        isTotal: true
                    }
                };
            });
        } else {
            console.warn("[Archive] No stored totals found - archive periods will be empty");
        }
    }
    
    console.log("[buildExpenseReviewPeriods] PARSED PERIODS:");
    console.log("  - currentPeriods:", currentPeriods.map(p => ({ key: p.key, employees: p.summary?.employeeCount, total: p.summary?.total })));
    console.log("  - archivePeriods:", archivePeriods.map(p => ({ key: p.key, employees: p.summary?.employeeCount, total: p.summary?.total })));
    
    const archiveMap = new Map(archivePeriods.map((period) => [period.key, period]));
    const combined = [];
    if (currentPeriods.length) {
        combined.push(currentPeriods[0]);
        archiveMap.delete(currentPeriods[0].key);
    }
    // Add all archive periods (deduplication happens below)
    archivePeriods.forEach((period) => {
        if (combined.length >= 6) return;
        if (!combined.some((existing) => existing.key === period.key)) {
            combined.push(period);
        }
    });
    
    console.log("[buildExpenseReviewPeriods] Combined before filter:", combined.map(p => ({ key: p.key, employees: p.summary?.employeeCount, total: p.summary?.total })));
    
    // Filter to only include periods that look like real pay periods:
    // - Must have at least 3 employees (filters out test data or partial entries)
    // - Must have meaningful total (> $1000) to filter out adjustment entries
    const minEmployeesForPayPeriod = 3;
    const minTotalForPayPeriod = 1000;
    
    const sorted = combined
        .filter((period) => {
            if (!period || !period.key) {
                console.log("buildExpenseReviewPeriods - EXCLUDED (no key):", period);
                return false;
            }
            const total = period.summary?.total || 
                ((period.summary?.fixed || 0) + (period.summary?.variable || 0) + (period.summary?.burden || 0));
            const employeeCount = period.summary?.employeeCount || 0;
            // Always include the current period (first one), filter others
            if (combined.indexOf(period) === 0) {
                console.log(`buildExpenseReviewPeriods - INCLUDED (current): ${period.key} - ${employeeCount} employees, $${total}`);
                return true;
            }
            const included = employeeCount >= minEmployeesForPayPeriod && total >= minTotalForPayPeriod;
            console.log(`  ${included ? "✓ INCLUDED" : "✗ EXCLUDED"}: ${period.key} - ${employeeCount} employees, $${total.toLocaleString()} (needs >=${minEmployeesForPayPeriod} emp, >=$${minTotalForPayPeriod})`);
            return included;
        })
        .sort((a, b) => (a.key < b.key ? 1 : -1))
        .slice(0, 6);
    
    console.log("[buildExpenseReviewPeriods] FINAL periods:", sorted.map(p => p.key));
    console.log("[buildExpenseReviewPeriods] ═══════════════════════════════════════════════");
    sorted.forEach((period, index) => {
        const previous = sorted[index + 1];
        const delta = previous ? period.summary.total - previous.summary.total : 0;
        period.summary.delta = delta;
        const previousDeptMap = new Map((previous?.departments || []).map((dept) => [dept.name, dept]));
        period.departments.forEach((dept) => {
            const prev = previousDeptMap.get(dept.name);
            dept.delta = prev ? dept.allIn - prev.allIn : 0;
        });
        period.totalsRow.delta = delta;
    });
    return sorted;
}

async function prepareExpenseReviewData() {
    if (!hasExcelRuntime()) {
        updateExpenseReviewState({
            loading: false,
            lastError: "Excel runtime is unavailable."
        });
        return;
    }
    updateExpenseReviewState({ loading: true, lastError: null });
    
    // Fetch customer mappings with expense_bucket for bucket classification
    const companyId = getConfigValue("SS_Company_ID");
    const module = "payroll-recorder";
    const customerMappings = await getCustomerMappingsWithBuckets(companyId, module);
    
    // Log bucket distribution
    const bucketCounts = {
        FIXED: customerMappings.filter(m => m.expense_bucket === 'FIXED').length,
        VARIABLE: customerMappings.filter(m => m.expense_bucket === 'VARIABLE').length,
        BURDEN: customerMappings.filter(m => m.expense_bucket === 'BURDEN').length,
        BENEFITS: customerMappings.filter(m => m.expense_bucket === 'BENEFITS').length,
        EXCLUDE: customerMappings.filter(m => m.expense_bucket === 'EXCLUDE').length
    };
    console.log(`[ExpenseReview] Loaded ${customerMappings.length} customer mappings`);
    console.log(`[ExpenseReview] Buckets: FIXED=${bucketCounts.FIXED}, VARIABLE=${bucketCounts.VARIABLE}, BURDEN=${bucketCounts.BURDEN}, BENEFITS=${bucketCounts.BENEFITS}, EXCLUDE=${bucketCounts.EXCLUDE}`);
    
    // Store customer mappings in state for use by other functions
    expenseReviewState.customerMappings = customerMappings;
    
    // Ensure expense taxonomy is loaded before processing (for sign info)
    try {
        await fetchExpenseTaxonomy();
    } catch (taxonomyError) {
        console.warn("[ExpenseReview] Failed to load taxonomy, using fallbacks:", taxonomyError);
    }
    
    // CRITICAL: Get measure universe FIRST to ensure totals match Step 1
    console.log("[ExpenseReview] Getting measure universe from Step 1...");
    const measureUniverse = await getPRDataCleanMeasureUniverse();
    if (measureUniverse.error) {
        console.warn("[ExpenseReview] Measure universe error:", measureUniverse.error);
    } else {
        console.log(`[ExpenseReview] Measure universe: ${measureUniverse.includedMeasureHeaders.length} measures, total=$${measureUniverse.total.toLocaleString()}`);
    }
    
    try {
        const periods = await Excel.run(async (context) => {
            // Check if required sheets exist
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
            const archiveTotalsSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_TOTALS);
            const reviewSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.EXPENSE_REVIEW);
            
            cleanSheet.load("isNullObject, name");
            archiveSheet.load("isNullObject, name");
            archiveTotalsSheet.load("isNullObject, name");
            reviewSheet.load("isNullObject, name");
            await context.sync();
            
            console.log("Expense Review - Sheet check:", {
                cleanSheet: cleanSheet.isNullObject ? "MISSING" : cleanSheet.name,
                archiveSheet: archiveSheet.isNullObject ? "MISSING" : archiveSheet.name,
                archiveTotalsSheet: archiveTotalsSheet.isNullObject ? "MISSING" : archiveTotalsSheet.name,
                reviewSheet: reviewSheet.isNullObject ? "MISSING" : reviewSheet.name
            });
            
            // Create missing sheets if needed
            if (reviewSheet.isNullObject) {
                console.log("Creating PR_Expense_Review sheet...");
                const newReviewSheet = context.workbook.worksheets.add(SHEET_NAMES.EXPENSE_REVIEW);
                await context.sync();
                // Re-get the sheet
                const createdSheet = context.workbook.worksheets.getItem(SHEET_NAMES.EXPENSE_REVIEW);
                
                // Get data even if cleanSheet is missing (will be empty)
                let cleanValues = [];
                let archiveValues = [];
                let archiveTotalsValues = [];
                
                if (!cleanSheet.isNullObject) {
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
                    cleanRange.load("values");
                    await context.sync();
                    cleanValues = cleanRange.isNullObject ? [] : cleanRange.values || [];
                }
                
                if (!archiveSheet.isNullObject) {
            const archiveRange = archiveSheet.getUsedRangeOrNullObject();
                    archiveRange.load("values");
                    await context.sync();
                    archiveValues = archiveRange.isNullObject ? [] : archiveRange.values || [];
                }

                if (!archiveTotalsSheet.isNullObject) {
            const archiveTotalsRange = archiveTotalsSheet.getUsedRangeOrNullObject();
                    archiveTotalsRange.load("values");
                    await context.sync();
                    archiveTotalsValues = archiveTotalsRange.isNullObject ? [] : archiveTotalsRange.values || [];
                }
                
                const periodData = buildExpenseReviewPeriods(cleanValues, archiveValues, measureUniverse, archiveTotalsValues);
                await writeExpenseReviewSheet(context, createdSheet, periodData);
                return periodData;
            }
            
            // Get data from existing sheets
            let cleanValues = [];
            let archiveValues = [];
            let archiveTotalsValues = [];
            
            if (!cleanSheet.isNullObject) {
                const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            cleanRange.load("values");
                await context.sync();
                cleanValues = cleanRange.isNullObject ? [] : cleanRange.values || [];
                console.log("Expense Review - PR_Data_Clean rows:", cleanValues.length);
            } else {
                console.warn("Expense Review - PR_Data_Clean sheet not found, using empty data");
            }
            
            if (!archiveSheet.isNullObject) {
                const archiveRange = archiveSheet.getUsedRangeOrNullObject();
            archiveRange.load("values");
            await context.sync();
                archiveValues = archiveRange.isNullObject ? [] : archiveRange.values || [];
                console.log("Expense Review - PR_Archive_Summary rows:", archiveValues.length);
            } else {
                console.warn("Expense Review - PR_Archive_Summary sheet not found, using empty data");
            }

            if (!archiveTotalsSheet.isNullObject) {
                const archiveTotalsRange = archiveTotalsSheet.getUsedRangeOrNullObject();
            archiveTotalsRange.load("values");
            await context.sync();
                archiveTotalsValues = archiveTotalsRange.isNullObject ? [] : archiveTotalsRange.values || [];
                console.log("Expense Review - PR_Archive_Totals rows:", archiveTotalsValues.length);
            } else {
                console.warn("Expense Review - PR_Archive_Totals sheet not found, using fallback archive data");
            }
            
            const periodData = buildExpenseReviewPeriods(cleanValues, archiveValues, measureUniverse, archiveTotalsValues);
            console.log("Expense Review - Periods built:", periodData.length);
            
            await writeExpenseReviewSheet(context, reviewSheet, periodData);
            return periodData;
        });
        updateExpenseReviewState({ loading: false, periods, lastError: null });
        
        // Run completeness check after data is loaded
        await runPayrollCompletenessCheck();
        renderApp(); // Re-render to show completeness check results
    } catch (error) {
        console.error("Expense Review: unable to build executive summary", error);
        console.error("Error details:", error.message, error.stack);
        updateExpenseReviewState({
            loading: false,
            lastError: `Unable to build the Expense Review: ${error.message || "Unknown error"}`,
            periods: []
        });
    }
}

async function writeExpenseReviewSheet(context, sheet, periods) {
    if (!sheet) {
        console.error("writeExpenseReviewSheet: sheet is null/undefined");
        return;
    }
    
    console.log("writeExpenseReviewSheet: Building executive dashboard with", periods.length, "periods");
    
    // Clear existing content and charts
    try {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("address");
        const charts = sheet.charts;
        charts.load("items");
        await context.sync();
        if (!usedRange.isNullObject) {
            usedRange.clear();
            await context.sync();
        }
        // Remove existing charts
        for (let i = charts.items.length - 1; i >= 0; i--) {
            charts.items[i].delete();
        }
        await context.sync();
    } catch (e) {
        console.warn("Could not clear sheet:", e);
    }
    
    // ═══════════════════════════════════════════════════════════════════
    // PREPARE DATA
    // ═══════════════════════════════════════════════════════════════════
    const current = periods[0] || {};
    // For "Same Period Prior Month", use periods[2] if available (skip most recent archived period)
    // This gives us the prior month period, not the most recent pay period
    const prior = periods.length > 2 ? periods[2] : (periods[1] || {});
    const currentSummary = current.summary || {};
    const priorSummary = prior.summary || {};
    
    // Get period from config
    // IMPORTANT: PR_Payroll_Date can be an Excel serial (number). Never pass numeric serials into new Date(serial)
    // because JS treats numbers as milliseconds since 1970, producing 12/31/1969 for small values.
    const rawPayrollDate = getPayrollDateValue();
    const normalizedPayrollDate = rawPayrollDate ? normalizePeriodKey(rawPayrollDate) : "";
    const configPeriod = getConfigValue("PR_Accounting_Period") || normalizedPayrollDate || "";
    
    // Key metrics
    const totalPayroll = Number(currentSummary.total) || 0;
    const priorTotal = Number(priorSummary.total) || 0;
    const periodChange = totalPayroll - priorTotal;
    const periodChangePct = priorTotal ? periodChange / priorTotal : 0;
    const employeeCount = Number(currentSummary.employeeCount) || 0;
    const priorEmployeeCount = Number(priorSummary.employeeCount) || 0;
    const headcountChange = employeeCount - priorEmployeeCount;
    const avgPerEmployee = employeeCount ? totalPayroll / employeeCount : 0;
    const priorAvgPerEmployee = priorEmployeeCount ? priorTotal / priorEmployeeCount : 0;
    const avgChange = avgPerEmployee - priorAvgPerEmployee;
    
    // Variable comp analysis - detect if this is a "variable pay" period
    // Look for commission, bonus patterns in period data
    const hasVariableComp = detectVariableCompPeriod(periods);
    const basePayOnly = detectBasePayOnlyPeriod(current, periods);
    
    const periodLabel = current.label || current.key || "Current Period";
    const generatedTimestamp = new Date().toLocaleString("en-US", { 
        month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit"
    });
    
    // Trend helpers
    const trendArrow = (val) => val > 0 ? "▲" : val < 0 ? "▼" : "—";
    
    // ═══════════════════════════════════════════════════════════════════
    // CALCULATE HISTORICAL RANGES FOR SPECTRUM VISUALIZATION
    // ═══════════════════════════════════════════════════════════════════
    const historicalTotals = periods.map(p => p.summary?.total || 0).filter(t => t > 0);
    const historicalAvgPerEmp = periods.map(p => {
        const s = p.summary || {};
        const emp = s.employeeCount || 0;
        return emp > 0 ? (s.total || 0) / emp : 0;
    }).filter(a => a > 0);
    const historicalChangePcts = periods.slice(0, -1).map((p, i) => {
        const curr = p.summary?.total || 0;
        const prev = periods[i + 1]?.summary?.total || 0;
        return prev > 0 ? (curr - prev) / prev : 0;
    });
    
    // Calculate ranges - includes current value to ensure the spectrum adjusts if current is outside historical range
    const calcRange = (arr, currentValue = null) => {
        // Include current value in range if provided (allows range to expand beyond historical)
        const values = currentValue !== null ? [...arr, currentValue] : arr;
        if (!values.length) return { min: 0, max: 0, avg: 0 };
        const min = Math.min(...values);
        const max = Math.max(...values);
        // Average is calculated from just the historical values (or all if no current provided)
        const avgBase = arr.length ? arr : values;
        const avg = avgBase.reduce((a, b) => a + b, 0) / avgBase.length;
        return { min, max, avg };
    };
    
    const payrollRange = calcRange(historicalTotals, totalPayroll);
    const avgEmpRange = calcRange(historicalAvgPerEmp, avgPerEmployee);
    const changePctRange = calcRange(historicalChangePcts);
    
    // Spectrum builder - creates a visual bar showing where current value falls
    // Uses Unicode block characters: ░ (light), ▒ (medium), ● (current position)
    const buildSpectrum = (current, min, max, width = 20) => {
        if (max <= min) return "░".repeat(width);
        const range = max - min;
        const position = Math.max(0, Math.min(1, (current - min) / range));
        const markerPos = Math.round(position * (width - 1));
        
        let bar = "";
        for (let i = 0; i < width; i++) {
            if (i === markerPos) {
                bar += "●";
            } else {
                bar += "░";
            }
        }
        return bar;
    };
    
    // ═══════════════════════════════════════════════════════════════════
    // EXTRACT FIXED/VARIABLE/BURDEN BREAKDOWNS
    // ═══════════════════════════════════════════════════════════════════
    const currentFixed = Number(currentSummary.fixed) || 0;
    const currentVariable = Number(currentSummary.variable) || 0;
    const currentBurden = Number(currentSummary.burden) || 0;
    const currentGross = currentFixed + currentVariable;
    const currentBurdenRate = totalPayroll ? currentBurden / totalPayroll : 0;
    
    const priorFixed = Number(priorSummary.fixed) || 0;
    const priorVariable = Number(priorSummary.variable) || 0;
    const priorBurden = Number(priorSummary.burden) || 0;
    const priorBurdenRate = priorTotal ? priorBurden / priorTotal : 0;
    
    // Calculate Sales & Marketing variable vs Other departments variable
    const departments = current.departments || [];
    const salesMarketingDepts = departments.filter(d => {
        const name = (d.name || "").toLowerCase();
        return name.includes("sales") || name.includes("marketing");
    });
    const otherDepts = departments.filter(d => {
        const name = (d.name || "").toLowerCase();
        return !name.includes("sales") && !name.includes("marketing");
    });
    
    const salesMarketingVariable = salesMarketingDepts.reduce((sum, d) => sum + (d.variable || 0), 0);
    const salesMarketingHeadcount = salesMarketingDepts.reduce((sum, d) => sum + (d.headcount || 0), 0);
    const otherVariable = otherDepts.reduce((sum, d) => sum + (d.variable || 0), 0);
    const otherHeadcount = otherDepts.reduce((sum, d) => sum + (d.headcount || 0), 0);
    
    const avgVariableSalesMarketing = salesMarketingHeadcount ? salesMarketingVariable / salesMarketingHeadcount : 0;
    const avgVariableOther = otherHeadcount ? otherVariable / otherHeadcount : 0;
    const avgFixedPerEmployee = employeeCount ? currentFixed / employeeCount : 0;
    
    // ═══════════════════════════════════════════════════════════════════
    // BUILD CLEAN DATA ARRAY - NEW LAYOUT
    // ═══════════════════════════════════════════════════════════════════
    const data = [];
    let rowIdx = 0;
    const rowMap = {};
    
    // ─── HEADER (Left justified) ───
    rowMap.headerStart = rowIdx;
    // Format period for the header. If we have an ISO date, render as a friendly long date.
    let formattedPeriod = configPeriod || periodLabel;
    if (/^\d{4}-\d{2}-\d{2}$/.test(formattedPeriod)) {
        const [year, month, day] = formattedPeriod.split("-").map(Number);
        const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        if (year && month && day) {
            formattedPeriod = `${monthNames[month - 1]} ${day}, ${year}`;
        }
    }
    
    data.push(["PAYROLL EXPENSE REVIEW"]); rowIdx++;
    data.push([`Period: ${formattedPeriod}`]); rowIdx++;
    data.push([`Generated: ${generatedTimestamp}`]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.headerEnd = rowIdx - 1;
    
    // ─── EXECUTIVE SUMMARY (Frozen Section) ───
    rowMap.execSummaryStart = rowIdx;
    data.push(["EXECUTIVE SUMMARY"]); rowIdx++;
    rowMap.execSummaryHeader = rowIdx - 1;
    data.push([""]); rowIdx++;
    
    // Column headers: Pay Date, Headcount, Fixed Salary, Variable Salary, Burden, Total Payroll, Burden Rate
    data.push(["", "Pay Date", "Headcount", "Fixed Salary", "Variable Salary", "Burden", "Total Payroll", "Burden Rate"]); rowIdx++;
    rowMap.execSummaryColHeaders = rowIdx - 1;
    
    // Current Pay Period row
    data.push(["Current Pay Period", current.label || current.key || "", employeeCount, currentFixed, currentVariable, currentBurden, totalPayroll, currentBurdenRate]); rowIdx++;
    rowMap.execSummaryCurrentRow = rowIdx - 1;
    
    // Same Period Prior Month row  
    data.push(["Same Period Prior Month", prior.label || prior.key || "", priorEmployeeCount, priorFixed, priorVariable, priorBurden, priorTotal, priorBurdenRate]); rowIdx++;
    rowMap.execSummaryPriorRow = rowIdx - 1;
    
    data.push([""]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.execSummaryEnd = rowIdx - 1;
    
    // ─── CURRENT PERIOD BREAKDOWN (DEPARTMENT) ───
    rowMap.deptBreakdownStart = rowIdx;
    data.push(["CURRENT PERIOD BREAKDOWN (DEPARTMENT)"]); rowIdx++;
    rowMap.deptBreakdownHeader = rowIdx - 1;
    data.push([""]); rowIdx++;
    data.push([`Payroll Date`, current.label || current.key || ""]); rowIdx++;
    data.push([""]); rowIdx++;
    
    // Column headers: Department, Fixed Salary, Variable Salary, Gross Pay, Burden, All-In Total, % of Total, Headcount
    data.push(["Department", "Fixed Salary", "Variable Salary", "Gross Pay", "Burden", "All-In Total", "% of Total", "Headcount"]); rowIdx++;
    rowMap.deptColHeaders = rowIdx - 1;
    
    // Department data rows
    const sortedDepts = [...departments].sort((a, b) => (b.allIn || 0) - (a.allIn || 0));
    rowMap.deptDataStart = rowIdx;
    sortedDepts.forEach((dept) => {
        data.push([
            dept.name || "",
            dept.fixed || 0,
            dept.variable || 0,
            dept.gross || 0,
            dept.burden || 0,
            dept.allIn || 0,
            dept.percent || 0,
            dept.headcount || 0
        ]); rowIdx++;
    });
    rowMap.deptDataEnd = rowIdx - 1;
    
    // TOTAL row
    if (current.totalsRow) {
        const totals = current.totalsRow;
        data.push([
            "TOTAL",
            totals.fixed || 0,
            totals.variable || 0,
            totals.gross || 0,
            totals.burden || 0,
            totals.allIn || 0,
            1,
            totals.headcount || 0
        ]); rowIdx++;
        rowMap.deptTotalsRow = rowIdx - 1;
    }
    
    data.push([""]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.deptBreakdownEnd = rowIdx - 1;
    
    // ─── HISTORICAL CONTEXT ───
    rowMap.historicalStart = rowIdx;
    data.push(["HISTORICAL CONTEXT"]); rowIdx++;
    rowMap.historicalHeader = rowIdx - 1;
    data.push([`Visual comparison of current period vs. historical range (${periods.length} periods). The dot (●) shows where you currently stand.`]); rowIdx++;
    data.push([""]); rowIdx++;
    
    // Format helpers for spectrum labels  
    const fmtK = (n) => `$${Math.round(n / 1000)}K`;
    const fmtPct = (n) => `${(n * 100).toFixed(1)}%`;
    
    // Column headers for historical context
    data.push(["", "Metric", "Low", "Range", "High", "", "Current", "Average"]); rowIdx++;
    rowMap.historicalColHeaders = rowIdx - 1;
    
    // Calculate additional historical ranges - include current values to allow range expansion
    // Don't filter zeros here - let calcRange handle it with current value inclusion
    const historicalFixed = periods.map(p => p.summary?.fixed || 0).filter(t => t > 0);
    const historicalVariable = periods.map(p => p.summary?.variable || 0); // Keep zeros for proper range
    const historicalBurdenRates = periods.map(p => {
        const s = p.summary || {};
        return s.total ? (s.burden || 0) / s.total : 0;
    }); // Keep zeros for proper range
    const historicalAvgFixed = periods.map(p => {
        const s = p.summary || {};
        const emp = s.employeeCount || 0;
        return emp > 0 ? (s.fixed || 0) / emp : 0;
    }).filter(a => a > 0);
    
    const fixedRange = calcRange(historicalFixed, currentFixed);
    const variableRange = calcRange(historicalVariable, currentVariable);
    const burdenRateRange = calcRange(historicalBurdenRates, currentBurdenRate);
    const avgFixedRange = calcRange(historicalAvgFixed, avgFixedPerEmployee);
    
    // Build spectrum rows
    rowMap.spectrumRows = [];
    
    // Total Payroll
    const payrollSpectrum = buildSpectrum(totalPayroll, payrollRange.min, payrollRange.max, 25);
    data.push(["", "Total Payroll", fmtK(payrollRange.min), payrollSpectrum, fmtK(payrollRange.max), "", fmtK(totalPayroll), fmtK(payrollRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    // Total Fixed Payroll
    const fixedSpectrum = buildSpectrum(currentFixed, fixedRange.min, fixedRange.max, 25);
    data.push(["", "Total Fixed Payroll", fmtK(fixedRange.min), fixedSpectrum, fmtK(fixedRange.max), "", fmtK(currentFixed), fmtK(fixedRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    // Total Variable Payroll
    const variableSpectrum = buildSpectrum(currentVariable, variableRange.min, variableRange.max, 25);
    data.push(["", "Total Variable Payroll", fmtK(variableRange.min), variableSpectrum, fmtK(variableRange.max), "", fmtK(currentVariable), fmtK(variableRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    data.push([""]); rowIdx++;
    
    // Avg Fixed Payroll per Employee
    const avgFixedSpectrum = buildSpectrum(avgFixedPerEmployee, avgFixedRange.min, avgFixedRange.max, 25);
    data.push(["", "Avg Fixed Payroll / Employee", fmtK(avgFixedRange.min), avgFixedSpectrum, fmtK(avgFixedRange.max), "", fmtK(avgFixedPerEmployee), fmtK(avgFixedRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    // Calculate historical ranges for Sales & Marketing variable
    const historicalAvgVarSM = periods.map(p => {
        const depts = p.departments || [];
        const smDepts = depts.filter(d => {
            const name = (d.name || "").toLowerCase();
            return name.includes("sales") || name.includes("marketing");
        });
        const smVar = smDepts.reduce((sum, d) => sum + (d.variable || 0), 0);
        const smHc = smDepts.reduce((sum, d) => sum + (d.headcount || 0), 0);
        return smHc > 0 ? smVar / smHc : 0;
    }); // Keep zeros for proper range - current value inclusion will handle expansion
    const avgVarSMRange = calcRange(historicalAvgVarSM, avgVariableSalesMarketing);
    
    // Calculate historical ranges for Other departments variable
    const historicalAvgVarOther = periods.map(p => {
        const depts = p.departments || [];
        const otherD = depts.filter(d => {
            const name = (d.name || "").toLowerCase();
            return !name.includes("sales") && !name.includes("marketing");
        });
        const otherV = otherD.reduce((sum, d) => sum + (d.variable || 0), 0);
        const otherH = otherD.reduce((sum, d) => sum + (d.headcount || 0), 0);
        return otherH > 0 ? otherV / otherH : 0;
    }); // Keep zeros for proper range
    const avgVarOtherRange = calcRange(historicalAvgVarOther, avgVariableOther);
    
    // Avg Variable Payroll per Sales & Marketing (with spectrum visualization)
    if (salesMarketingHeadcount > 0) {
        const avgVarSMSpectrum = buildSpectrum(avgVariableSalesMarketing, avgVarSMRange.min, avgVarSMRange.max, 25);
        data.push(["", "Avg Variable / Sales & Marketing", fmtK(avgVarSMRange.min), avgVarSMSpectrum, fmtK(avgVarSMRange.max), "", fmtK(avgVariableSalesMarketing), `${salesMarketingHeadcount} emp`]); rowIdx++;
        rowMap.spectrumRows.push(rowIdx - 1);
    }
    
    // Avg Variable Payroll per Other Departments (with spectrum visualization)
    if (otherHeadcount > 0) {
        const avgVarOtherSpectrum = buildSpectrum(avgVariableOther, avgVarOtherRange.min, avgVarOtherRange.max, 25);
        data.push(["", "Avg Variable / Other Depts", fmtK(avgVarOtherRange.min), avgVarOtherSpectrum, fmtK(avgVarOtherRange.max), "", fmtK(avgVariableOther), `${otherHeadcount} emp`]); rowIdx++;
        rowMap.spectrumRows.push(rowIdx - 1);
    }
    
    data.push([""]); rowIdx++;
    
    // Burden Rate (%)
    const burdenRateSpectrum = buildSpectrum(currentBurdenRate, burdenRateRange.min, burdenRateRange.max, 25);
    data.push(["", "Burden Rate (%)", fmtPct(burdenRateRange.min), burdenRateSpectrum, fmtPct(burdenRateRange.max), "", fmtPct(currentBurdenRate), fmtPct(burdenRateRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    data.push([""]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.historicalEnd = rowIdx - 1;
    
    // ─── PERIOD TRENDS ───
    rowMap.trendsStart = rowIdx;
    data.push(["PERIOD TRENDS"]); rowIdx++;
    rowMap.trendsHeader = rowIdx - 1;
    data.push([""]); rowIdx++;
    
    // Trend data table (will be used for chart)
    data.push(["Pay Period", "Total Payroll", "Fixed Payroll", "Variable Payroll", "Burden", "Headcount"]); rowIdx++;
    rowMap.trendColHeaders = rowIdx - 1;
    
    // Up to 6 periods in reverse chronological order (oldest first for chart)
    const trendPeriods = periods.slice(0, 6).reverse();
    rowMap.trendDataStart = rowIdx;
    trendPeriods.forEach((period) => {
        const s = period.summary || {};
        data.push([
            period.label || period.key || "",
            s.total || 0,
            s.fixed || 0,
            s.variable || 0,
            s.burden || 0,
            s.employeeCount || 0
        ]); rowIdx++;
    });
    rowMap.trendDataEnd = rowIdx - 1;
    
    data.push([""]); rowIdx++;
    rowMap.trendsEnd = rowIdx - 1;
    
    // Reserve space for charts (payroll chart + headcount chart)
    rowMap.chartStart = rowIdx;
    for (let i = 0; i < 15; i++) {
        data.push([""]); rowIdx++;
    }
    rowMap.payrollChartEnd = rowIdx - 1;
    
    // Space for headcount chart
    rowMap.headcountChartStart = rowIdx;
    for (let i = 0; i < 12; i++) {
        data.push([""]); rowIdx++;
    }
    rowMap.headcountChartEnd = rowIdx - 1;
    
    // ═══════════════════════════════════════════════════════════════════
    // WRITE DATA (10 columns to accommodate all data)
    // ═══════════════════════════════════════════════════════════════════
    console.log("writeExpenseReviewSheet: Writing", data.length, "rows");
    
    // Normalize all rows to 10 columns
    const normalizedData = data.map(row => {
        const r = Array.isArray(row) ? row : [""];
        while (r.length < 10) r.push("");
        return r.slice(0, 10);
    });
    
    try {
        const dataRange = sheet.getRangeByIndexes(0, 0, normalizedData.length, 10);
        dataRange.values = normalizedData;
        await context.sync();
    } catch (writeError) {
        console.error("writeExpenseReviewSheet: Write failed", writeError);
        throw writeError;
    }
    
    // ═══════════════════════════════════════════════════════════════════
    // APPLY FORMATTING
    // ═══════════════════════════════════════════════════════════════════
    try {
        // Column widths
        sheet.getRange("A:A").format.columnWidth = 200;   // Section headers / Department names
        sheet.getRange("B:B").format.columnWidth = 130;   // Fixed Salary / Metric names
        sheet.getRange("C:C").format.columnWidth = 100;   // Variable Salary / Low
        sheet.getRange("D:D").format.columnWidth = 200;   // Gross Pay / Spectrum (wider for dots)
        sheet.getRange("E:E").format.columnWidth = 100;   // Burden / High
        sheet.getRange("F:F").format.columnWidth = 100;   // All-In Total
        sheet.getRange("G:G").format.columnWidth = 100;   // % of Total / Current
        sheet.getRange("H:H").format.columnWidth = 100;   // Headcount / Average
        sheet.getRange("I:I").format.columnWidth = 80;
        sheet.getRange("J:J").format.columnWidth = 80;
        await context.sync();
        
        // ─── HEADER SECTION (Left justified) ───
        const titleCell = sheet.getRange("A1");
        titleCell.format.font.bold = true;
        titleCell.format.font.size = 22;
        titleCell.format.font.color = "#1e293b";
        
        sheet.getRange("A2").format.font.size = 11;
        sheet.getRange("A2").format.font.color = "#64748b";
        sheet.getRange("A3").format.font.size = 10;
        sheet.getRange("A3").format.font.color = "#94a3b8";
        await context.sync();
        
        // ─── EXECUTIVE SUMMARY SECTION ───
        const execHeader = sheet.getRange(`A${rowMap.execSummaryHeader + 1}`);
        execHeader.format.font.bold = true;
        execHeader.format.font.size = 14;
        execHeader.format.font.color = "#1e293b";
        
        // Column headers - dark background
        const execColHeaders = sheet.getRange(`A${rowMap.execSummaryColHeaders + 1}:H${rowMap.execSummaryColHeaders + 1}`);
        execColHeaders.format.font.bold = true;
        execColHeaders.format.font.size = 10;
        execColHeaders.format.fill.color = "#1e293b";
        execColHeaders.format.font.color = "#ffffff";
        
        // Current period row - light green background
        const currentRow = sheet.getRange(`A${rowMap.execSummaryCurrentRow + 1}:H${rowMap.execSummaryCurrentRow + 1}`);
        currentRow.format.fill.color = "#dcfce7";
        currentRow.format.font.bold = true;
        
        // Prior period row - light gray background
        const priorRow = sheet.getRange(`A${rowMap.execSummaryPriorRow + 1}:H${rowMap.execSummaryPriorRow + 1}`);
        priorRow.format.fill.color = "#f1f5f9";
        
        // Number formats for executive summary
        for (const rowNum of [rowMap.execSummaryCurrentRow + 1, rowMap.execSummaryPriorRow + 1]) {
            sheet.getRange(`C${rowNum}`).numberFormat = [["#,##0"]];       // Headcount
            sheet.getRange(`D${rowNum}`).numberFormat = [["$#,##0"]];      // Fixed
            sheet.getRange(`E${rowNum}`).numberFormat = [["$#,##0"]];      // Variable
            sheet.getRange(`F${rowNum}`).numberFormat = [["$#,##0"]];      // Burden
            sheet.getRange(`G${rowNum}`).numberFormat = [["$#,##0"]];      // Total
            sheet.getRange(`H${rowNum}`).numberFormat = [["0.00%"]];       // Burden Rate
        }
        await context.sync();
        
        // ─── DEPARTMENT BREAKDOWN SECTION ───
        const deptHeader = sheet.getRange(`A${rowMap.deptBreakdownHeader + 1}`);
        deptHeader.format.font.bold = true;
        deptHeader.format.font.size = 14;
        deptHeader.format.font.color = "#1e293b";
        
        // Column headers - dark background
        const deptColHeaders = sheet.getRange(`A${rowMap.deptColHeaders + 1}:H${rowMap.deptColHeaders + 1}`);
        deptColHeaders.format.font.bold = true;
        deptColHeaders.format.font.size = 10;
        deptColHeaders.format.fill.color = "#1e293b";
        deptColHeaders.format.font.color = "#ffffff";
        
        // Department data rows
        for (let i = rowMap.deptDataStart; i <= rowMap.deptDataEnd; i++) {
            const row = i + 1;
            sheet.getRange(`B${row}`).numberFormat = [["$#,##0"]];   // Fixed
            sheet.getRange(`C${row}`).numberFormat = [["$#,##0"]];   // Variable
            sheet.getRange(`D${row}`).numberFormat = [["$#,##0"]];   // Gross
            sheet.getRange(`E${row}`).numberFormat = [["$#,##0"]];   // Burden
            sheet.getRange(`F${row}`).numberFormat = [["$#,##0"]];   // All-In
            sheet.getRange(`G${row}`).numberFormat = [["0.00%"]];    // % of Total
            sheet.getRange(`H${row}`).numberFormat = [["#,##0"]];    // Headcount
            
            // Alternate row shading
            if ((i - rowMap.deptDataStart) % 2 === 1) {
                sheet.getRange(`A${row}:H${row}`).format.fill.color = "#f8fafc";
            }
        }
        
        // Totals row - dark background
        if (rowMap.deptTotalsRow) {
            const totalsRange = sheet.getRange(`A${rowMap.deptTotalsRow + 1}:H${rowMap.deptTotalsRow + 1}`);
            totalsRange.format.font.bold = true;
            totalsRange.format.fill.color = "#1e293b";
            totalsRange.format.font.color = "#ffffff";
            
            const row = rowMap.deptTotalsRow + 1;
            sheet.getRange(`B${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`C${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`D${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`E${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`F${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`G${row}`).numberFormat = [["0%"]];
            sheet.getRange(`H${row}`).numberFormat = [["#,##0"]];
        }
        await context.sync();
        
        // ─── HISTORICAL CONTEXT SECTION ───
        const histHeader = sheet.getRange(`A${rowMap.historicalHeader + 1}`);
        histHeader.format.font.bold = true;
        histHeader.format.font.size = 14;
        histHeader.format.font.color = "#1e293b";
        
        // Description text
        sheet.getRange(`A${rowMap.historicalHeader + 2}`).format.font.size = 10;
        sheet.getRange(`A${rowMap.historicalHeader + 2}`).format.font.color = "#64748b";
        sheet.getRange(`A${rowMap.historicalHeader + 2}`).format.font.italic = true;
        
        // Column headers - center Low, High, Current, Average
        const histColHeaders = sheet.getRange(`A${rowMap.historicalColHeaders + 1}:H${rowMap.historicalColHeaders + 1}`);
        histColHeaders.format.font.bold = true;
        histColHeaders.format.font.size = 10;
        histColHeaders.format.fill.color = "#e2e8f0";
        histColHeaders.format.font.color = "#334155";
        // Center the column headers for Low (C), High (E), Current (G), Average (H)
        sheet.getRange(`C${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        sheet.getRange(`E${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        sheet.getRange(`G${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        sheet.getRange(`H${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        
        // Format spectrum rows
        rowMap.spectrumRows.forEach(r => {
            // Spectrum bar - use monospace font for consistent width
            sheet.getRange(`D${r + 1}`).format.font.name = "Consolas";
            sheet.getRange(`D${r + 1}`).format.font.size = 14;
            sheet.getRange(`D${r + 1}`).format.font.color = "#6366f1";
            sheet.getRange(`D${r + 1}`).format.horizontalAlignment = "Center";
            
            // Metric label (left-aligned by default)
            sheet.getRange(`B${r + 1}`).format.font.color = "#334155";
            
            // Low (C) - centered
            sheet.getRange(`C${r + 1}`).format.font.color = "#94a3b8";
            sheet.getRange(`C${r + 1}`).format.horizontalAlignment = "Center";
            
            // High (E) - centered
            sheet.getRange(`E${r + 1}`).format.font.color = "#94a3b8";
            sheet.getRange(`E${r + 1}`).format.horizontalAlignment = "Center";
            
            // Current (G) - centered, bold
            sheet.getRange(`G${r + 1}`).format.font.bold = true;
            sheet.getRange(`G${r + 1}`).format.font.color = "#1e293b";
            sheet.getRange(`G${r + 1}`).format.horizontalAlignment = "Center";
            
            // Average (H) - centered
            sheet.getRange(`H${r + 1}`).format.font.color = "#94a3b8";
            sheet.getRange(`H${r + 1}`).format.horizontalAlignment = "Center";
        });
        await context.sync();
        
        // ─── PERIOD TRENDS SECTION ───
        const trendsHeader = sheet.getRange(`A${rowMap.trendsHeader + 1}`);
        trendsHeader.format.font.bold = true;
        trendsHeader.format.font.size = 14;
        trendsHeader.format.font.color = "#1e293b";
        
        // Trend table headers
        const trendColHeaders = sheet.getRange(`A${rowMap.trendColHeaders + 1}:F${rowMap.trendColHeaders + 1}`);
        trendColHeaders.format.font.bold = true;
        trendColHeaders.format.font.size = 10;
        trendColHeaders.format.fill.color = "#1e293b";
        trendColHeaders.format.font.color = "#ffffff";
        
        // Trend data number formats
        for (let i = rowMap.trendDataStart; i <= rowMap.trendDataEnd; i++) {
            const row = i + 1;
            sheet.getRange(`B${row}`).numberFormat = [["$#,##0"]];   // Total
            sheet.getRange(`C${row}`).numberFormat = [["$#,##0"]];   // Fixed
            sheet.getRange(`D${row}`).numberFormat = [["$#,##0"]];   // Variable
            sheet.getRange(`E${row}`).numberFormat = [["$#,##0"]];   // Burden
            sheet.getRange(`F${row}`).numberFormat = [["#,##0"]];    // Headcount
            
            if ((i - rowMap.trendDataStart) % 2 === 1) {
                sheet.getRange(`A${row}:F${row}`).format.fill.color = "#f8fafc";
            }
        }
        await context.sync();
        
        // ─── CREATE PAYROLL TRENDS CHART (without Headcount) ───
        if (trendPeriods.length >= 2) {
            try {
                // Chart data range - exclude Headcount column (A:E instead of A:F)
                const payrollChartRange = sheet.getRange(`A${rowMap.trendColHeaders + 1}:E${rowMap.trendDataEnd + 1}`);
                
                // Create the payroll chart
                const payrollChart = sheet.charts.add(
                    Excel.ChartType.lineMarkers,
                    payrollChartRange,
                    Excel.ChartSeriesBy.columns
                );
                
                // Position the chart below the trend data
                payrollChart.setPosition(`A${rowMap.chartStart + 1}`, `J${rowMap.payrollChartEnd + 1}`);
                payrollChart.title.text = "Payroll Expense Trends";
                payrollChart.title.format.font.size = 14;
                payrollChart.title.format.font.bold = true;
                
                // Configure legend
                payrollChart.legend.position = Excel.ChartLegendPosition.bottom;
                
                // Style the chart
                payrollChart.format.fill.setSolidColor("#ffffff");
                payrollChart.format.border.lineStyle = Excel.ChartLineStyle.continuous;
                payrollChart.format.border.color = "#e2e8f0";
                
                // Set X-axis to use text/category labels (not date axis)
                // This prevents Excel from interpolating dates between pay periods
                const categoryAxis = payrollChart.axes.getItem(Excel.ChartAxisType.category);
                categoryAxis.categoryType = Excel.ChartAxisCategoryType.textAxis;
                categoryAxis.setCategoryNames(sheet.getRange(`A${rowMap.trendDataStart + 1}:A${rowMap.trendDataEnd + 1}`));
                
                await context.sync();
                
                // Format series colors: Total=Blue, Fixed=Green, Variable=Orange, Burden=Purple
                const payrollSeries = payrollChart.series;
                payrollSeries.load("count");
                await context.sync();
                
                const payrollColors = ["#3b82f6", "#22c55e", "#f97316", "#8b5cf6"];
                for (let i = 0; i < Math.min(payrollSeries.count, payrollColors.length); i++) {
                    const s = payrollSeries.getItemAt(i);
                    s.format.line.color = payrollColors[i];
                    s.format.line.weight = 2;
                    s.markerStyle = Excel.ChartMarkerStyle.circle;
                    s.markerSize = 6;
                    s.markerBackgroundColor = payrollColors[i];
                }
                await context.sync();
                
                console.log("writeExpenseReviewSheet: Payroll chart created successfully");
            } catch (chartError) {
                console.warn("writeExpenseReviewSheet: Payroll chart creation failed (non-critical)", chartError);
            }
            
            // ─── CREATE HEADCOUNT CHART (separate scale) ───
            try {
                // For headcount, we need to create a chart from contiguous data
                // Use just columns A (Pay Period) and F (Headcount) by creating from full range
                // then removing unwanted series
                const headcountChartRange = sheet.getRange(`A${rowMap.trendColHeaders + 1}:F${rowMap.trendDataEnd + 1}`);
                
                // Create the headcount chart from full range
                const headcountChart = sheet.charts.add(
                    Excel.ChartType.lineMarkers,
                    headcountChartRange,
                    Excel.ChartSeriesBy.columns
                );
                
                // Position below the payroll chart
                headcountChart.setPosition(`A${rowMap.headcountChartStart + 1}`, `J${rowMap.headcountChartEnd + 1}`);
                headcountChart.title.text = "Headcount Trend";
                headcountChart.title.format.font.size = 12;
                headcountChart.title.format.font.bold = true;
                
                // Configure legend
                headcountChart.legend.visible = false;
                
                // Style the chart
                headcountChart.format.fill.setSolidColor("#ffffff");
                headcountChart.format.border.lineStyle = Excel.ChartLineStyle.continuous;
                headcountChart.format.border.color = "#e2e8f0";
                
                // Set X-axis to use text/category labels (not date axis)
                const headcountCategoryAxis = headcountChart.axes.getItem(Excel.ChartAxisType.category);
                headcountCategoryAxis.categoryType = Excel.ChartAxisCategoryType.textAxis;
                headcountCategoryAxis.setCategoryNames(sheet.getRange(`A${rowMap.trendDataStart + 1}:A${rowMap.trendDataEnd + 1}`));
                
                await context.sync();
                
                // Delete unwanted series (Total, Fixed, Variable, Burden) - keep only Headcount
                const hcSeries = headcountChart.series;
                hcSeries.load("count, items/name");
                await context.sync();
                
                // Delete series in reverse order (so indices don't shift)
                // Series order: Total Payroll, Fixed Payroll, Variable Payroll, Burden, Headcount
                // We want to keep only the last one (Headcount)
                for (let i = hcSeries.count - 2; i >= 0; i--) {
                    const s = hcSeries.getItemAt(i);
                    s.delete();
                }
                await context.sync();
                
                // Reload and format remaining series (Headcount)
                hcSeries.load("count");
                await context.sync();
                
                if (hcSeries.count > 0) {
                    const s = hcSeries.getItemAt(0);
                    s.format.line.color = "#64748b";
                    s.format.line.weight = 2.5;
                    s.markerStyle = Excel.ChartMarkerStyle.circle;
                    s.markerSize = 8;
                    s.markerBackgroundColor = "#64748b";
                }
                await context.sync();
                
                console.log("writeExpenseReviewSheet: Headcount chart created successfully");
            } catch (chartError) {
                console.warn("writeExpenseReviewSheet: Headcount chart creation failed (non-critical)", chartError);
            }
        }
        
        // ─── FINAL TOUCHES ───
        // Freeze first 11 rows (header + executive summary)
        sheet.freezePanes.freezeRows(rowMap.execSummaryEnd + 1);
        sheet.pageLayout.orientation = Excel.PageOrientation.landscape;
        sheet.getRange("A1").select();
        await context.sync();
        
        console.log("writeExpenseReviewSheet: Formatting applied successfully");
        
    } catch (formatError) {
        console.warn("writeExpenseReviewSheet: Formatting error (non-critical)", formatError);
    }
}

// Helper: Detect if this is a variable compensation period (has commissions/bonuses)
function detectVariableCompPeriod(periods) {
    if (!periods || !periods.length) return false;
    const current = periods[0];
    // Look for commission or bonus in the data categories
    const categories = current.summary?.categories || [];
    return categories.some(cat => {
        const name = (cat.name || "").toLowerCase();
        return name.includes("commission") || name.includes("bonus") || name.includes("variable");
    });
}

// Helper: Detect if this is likely a "base pay only" period
function detectBasePayOnlyPeriod(current, periods) {
    if (!current || periods.length < 2) return false;
    
    // Calculate average payroll across all periods
    const totals = periods.map(p => p.summary?.total || 0).filter(t => t > 0);
    if (totals.length < 2) return false;
    
    const avg = totals.reduce((a, b) => a + b, 0) / totals.length;
    const currentTotal = current.summary?.total || 0;
    
    // If current period is significantly below average (>10% lower), likely base-pay only
    const percentOfAvg = avg > 0 ? currentTotal / avg : 1;
    return percentOfAvg < 0.90;
}

async function activateWorksheet(name) {
    if (!hasExcelRuntime() || !name) return;
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(name);
            sheet.load("name");
            await context.sync();
            if (sheet.isNullObject) return;
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.warn(`Payroll Recorder: unable to activate worksheet ${name}`, error);
    }
}


function updateValidationState(partial = {}, { rerender = true } = {}) {
    Object.assign(validationState, partial);
    const prTotal = Number(validationState.prDataTotal);
    const cleanTotal = Number(validationState.cleanTotal);
    validationState.reconDifference =
        Number.isFinite(prTotal) && Number.isFinite(cleanTotal) ? prTotal - cleanTotal : null;
    const bankAmountNumber = parseBankAmount(validationState.bankAmount);
    validationState.bankDifference =
        Number.isFinite(cleanTotal) && !Number.isNaN(bankAmountNumber)
            ? cleanTotal - bankAmountNumber
            : null;
    validationState.plugEnabled =
        validationState.bankDifference != null && Math.abs(validationState.bankDifference) >= 0.5;
    if (rerender) {
        renderApp();
    } else {
        refreshValidationUiMetrics();
    }
}

function refreshValidationUiMetrics() {
    if (appState.activeStepId !== 3) return;
    const assignValue = (id, value) => {
        const element = document.getElementById(id);
        if (element) {
            element.value = value;
        }
    };
    assignValue("pr-data-total-value", formatCurrency(validationState.prDataTotal));
    assignValue("clean-total-value", formatCurrency(validationState.cleanTotal));
    assignValue("recon-diff-value", formatCurrency(validationState.reconDifference));
    assignValue("bank-clean-total-value", formatCurrency(validationState.cleanTotal));
    assignValue(
        "bank-diff-value",
        validationState.bankDifference != null ? formatCurrency(validationState.bankDifference) : "---"
    );
    const hint = document.getElementById("bank-diff-hint");
    if (hint) {
        hint.textContent =
            validationState.bankDifference == null
                ? ""
                : Math.abs(validationState.bankDifference) < 0.5
                    ? "Difference within acceptable tolerance."
                    : "Difference exceeds tolerance and should be resolved.";
    }
    const plugButton = document.getElementById("bank-plug-btn");
    if (plugButton) {
        plugButton.disabled = !validationState.plugEnabled;
    }
}

function updateExpenseReviewState(partial = {}, { rerender = true } = {}) {
    Object.assign(expenseReviewState, partial);
    if (rerender) {
        renderApp();
    }
}

async function prepareValidationData() {
    if (!hasExcelRuntime()) {
        updateValidationState({
            loading: false,
            lastError: "Excel runtime is unavailable.",
            prDataTotal: null,
            cleanTotal: null
        });
        return;
    }
    updateValidationState({ loading: true, lastError: null });
    try {
        // Read payroll date fresh from SS_PF_Config table
        let payrollDate = "";
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            console.log("DEBUG - Config table found:", !!table);
            if (table) {
                const body = table.getDataBodyRange();
                body.load("values");
                await context.sync();
                const rows = body.values || [];
                console.log("DEBUG - Config table rows:", rows.length);
                console.log("DEBUG - Looking for payroll date aliases:", PAYROLL_DATE_ALIASES);
                console.log("DEBUG - CONFIG_COLUMNS.FIELD:", CONFIG_COLUMNS.FIELD, "CONFIG_COLUMNS.VALUE:", CONFIG_COLUMNS.VALUE);
                
                // Look for payroll date field in config
                for (const row of rows) {
                    const fieldName = String(row[CONFIG_COLUMNS.FIELD] || "").trim();
                    const fieldValue = row[CONFIG_COLUMNS.VALUE];
                    
                    // Check if this is a payroll date field
                    const isMatch = PAYROLL_DATE_ALIASES.some(alias => 
                        fieldName === alias || 
                        normalizeFieldName(fieldName) === normalizeFieldName(alias)
                    );
                    
                    if (fieldName.toLowerCase().includes("payroll") || fieldName.toLowerCase().includes("date")) {
                        console.log("DEBUG - Potential date field:", fieldName, "=", fieldValue, "| isMatch:", isMatch);
                    }
                    
                    if (isMatch) {
                        const rawValue = row[CONFIG_COLUMNS.VALUE];
                        console.log("DEBUG - Found payroll date field!", fieldName, "raw value:", rawValue);
                        payrollDate = formatDateInput(rawValue) || "";
                        console.log("DEBUG - Formatted payroll date:", payrollDate);
                        break;
                    }
                }
                
                if (!payrollDate) {
                    console.warn("DEBUG - No payroll date found in config! Available fields:");
                    rows.forEach((row, i) => {
                        console.log(`  Row ${i}: Field="${row[CONFIG_COLUMNS.FIELD]}" Value="${row[CONFIG_COLUMNS.VALUE]}"`);
                    });
                }
            } else {
                console.warn("DEBUG - Config table not found!");
            }
        });
        console.log("DEBUG prepareValidationData - Final Payroll Date:", payrollDate || "(empty)");
        const result = await Excel.run(async (context) => {
            // NOTE: PR_Data sheet no longer exists - workflow goes directly to PR_Data_Clean
            const mappingSheet = context.workbook.worksheets.getItem(SHEET_NAMES.EXPENSE_MAPPING);
            const cleanSheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA_CLEAN);
            const mappingRange = mappingSheet.getUsedRangeOrNullObject();
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            mappingRange.load("values");
            cleanRange.load(["address", "rowCount", "values"]);
            await context.sync();
            const dataValues = cleanRange.isNullObject ? [] : cleanRange.values || [];
            const mappingValues = mappingRange.isNullObject ? [] : mappingRange.values || [];
            console.log("DEBUG prepareValidationData - PR_Data_Clean rows:", dataValues.length);
            console.log("DEBUG prepareValidationData - PR_Data_Clean headers:", dataValues[0]);
            console.log("DEBUG prepareValidationData - PR_Expense_Mapping rows:", mappingValues.length);
            const mappingHeaders = mappingValues[0]?.map((header) => normalizeHeader(header)) || [];
            const findHeaderIndex = (predicate) => mappingHeaders.findIndex(predicate);
            const mappingIdx = {
                category: findHeaderIndex((header) => header.includes("category")),
                accountNumber: findHeaderIndex(
                    (header) => header.includes("account") && (header.includes("number") || header.includes("#"))
                ),
                accountName: findHeaderIndex((header) => header.includes("account") && header.includes("name")),
                expenseReview: findHeaderIndex((header) => header.includes("expense") && header.includes("review"))
            };
            const mappingMap = new Map();
            mappingValues.slice(1).forEach((row) => {
                const category =
                    mappingIdx.category >= 0 ? normalizePayrollCategory(row[mappingIdx.category]) : "";
                if (!category) return;
                mappingMap.set(category, {
                    accountNumber: mappingIdx.accountNumber >= 0 ? row[mappingIdx.accountNumber] ?? "" : "",
                    accountName: mappingIdx.accountName >= 0 ? row[mappingIdx.accountName] ?? "" : "",
                    expenseReview: mappingIdx.expenseReview >= 0 ? row[mappingIdx.expenseReview] ?? "" : ""
                });
            });
            // Read existing headers from PR_Data_Clean
            const cleanHeaderRange = cleanSheet.getRangeByIndexes(0, 0, 1, 8);
            cleanHeaderRange.load("values");
            await context.sync();

            const existingHeaders = cleanHeaderRange.values[0] || [];
            const cleanHeadersNormalized = existingHeaders.map((h) => normalizeHeader(h));
            console.log("DEBUG prepareValidationData - PR_Data_Clean headers:", existingHeaders);
            console.log("DEBUG prepareValidationData - PR_Data_Clean normalized:", cleanHeadersNormalized);

            // Map output fields to column positions in PR_Data_Clean
            console.log("DEBUG - PR_Data_Clean headers:", existingHeaders);
            console.log("DEBUG - PR_Data_Clean normalized headers:", cleanHeadersNormalized);
            
            const payrollDateColIdx = cleanHeadersNormalized.findIndex(
                    (h) => (h.includes("payroll") || h.includes("period")) && h.includes("date")
            );
            console.log("DEBUG - payrollDate column index:", payrollDateColIdx);
            if (payrollDateColIdx === -1) {
                console.warn("DEBUG - No payroll date column found! Looking for header containing 'payroll'/'period' AND 'date'");
                cleanHeadersNormalized.forEach((h, i) => console.log(`  Col ${i}: "${h}"`));
            }
            
            const fieldMap = {
                payrollDate: payrollDateColIdx,
                employee: cleanHeadersNormalized.findIndex((h) => h.includes("employee")),
                department: pickDepartmentIndex(cleanHeadersNormalized),
                payrollCategory: cleanHeadersNormalized.findIndex((h) => h.includes("payroll") && h.includes("category")),
                accountNumber: cleanHeadersNormalized.findIndex((h) => h.includes("account") && (h.includes("number") || h.includes("#"))),
                accountName: cleanHeadersNormalized.findIndex((h) => h.includes("account") && h.includes("name")),
                amount: cleanHeadersNormalized.findIndex((h) => h.includes("amount")),
                expenseReview: cleanHeadersNormalized.findIndex((h) => h.includes("expense") && h.includes("review"))
            };
            console.log("DEBUG prepareValidationData - fieldMap:", fieldMap);

            const columnCount = existingHeaders.length;
            const cleanRows = [];
            let prDataTotal = 0;
            let cleanTotal = 0;
            if (dataValues.length >= 2) {
                const headerRow = dataValues[0];
                const headers = headerRow.map((header) => normalizeHeader(header));
                console.log("DEBUG prepareValidationData - Normalized headers:", headers);
                const employeeIdx = headers.findIndex((header) => header.includes("employee"));
                const departmentIdx = pickDepartmentIndex(headers);
                console.log("DEBUG prepareValidationData - Employee column index:", employeeIdx, "searching for 'employee' in:", headers[6]);
                console.log("DEBUG prepareValidationData - Department column index:", departmentIdx);
                const hasMappings = mappingMap.size > 0;
                const numericColumns = headers.reduce((list, header, index) => {
                    if (index === employeeIdx || index === departmentIdx) return list;
                    if (!header) return list;
                    if (header.includes("total") || header.includes("gross")) return list;
                    if (header.includes("date") || header.includes("period")) return list;
                    const normalizedCategory = normalizePayrollCategory(headerRow[index] || header);
                    if (hasMappings && !mappingMap.has(normalizedCategory)) return list;
                    list.push(index);
                    return list;
                }, []);
                console.log("DEBUG prepareValidationData - Numeric columns:", numericColumns.length, numericColumns);
                for (let i = 1; i < dataValues.length; i += 1) {
                    const row = dataValues[i];
                    const employee = employeeIdx >= 0 ? normalizeString(row[employeeIdx]) : "";
                    if (!employee || employee.toLowerCase().includes("total")) continue;
                    const department = departmentIdx >= 0 ? row[departmentIdx] || "" : "";
                    numericColumns.forEach((columnIndex) => {
                        const rawValue = row[columnIndex];
                        const amount = Number(rawValue);
                        if (!Number.isFinite(amount) || amount === 0) return;
                        prDataTotal += amount;
                        const payrollCategory = headerRow[columnIndex] || headers[columnIndex] || `Column ${columnIndex + 1}`;
                        const mapping = mappingMap.get(normalizePayrollCategory(payrollCategory)) || {};
                        cleanTotal += amount;

                        // Build row matching existing column positions
                        const newRow = new Array(columnCount).fill("");
                        // Write payroll date to the appropriate column
                        if (fieldMap.payrollDate >= 0) {
                            newRow[fieldMap.payrollDate] = payrollDate;
                        } else if (columnCount > 0) {
                            // Fallback to first column if no header match
                            newRow[0] = payrollDate;
                        }
                        // Log first row being built to verify payrollDate
                        if (cleanRows.length === 0) {
                            console.log("DEBUG - Building first PR_Data_Clean row:");
                            console.log("  payrollDate value:", payrollDate);
                            console.log("  fieldMap.payrollDate:", fieldMap.payrollDate);
                            console.log("  Writing to column index:", fieldMap.payrollDate >= 0 ? fieldMap.payrollDate : 0);
                        }
                        if (fieldMap.employee >= 0) newRow[fieldMap.employee] = employee;
                        if (fieldMap.department >= 0) newRow[fieldMap.department] = department;
                        if (fieldMap.payrollCategory >= 0) newRow[fieldMap.payrollCategory] = payrollCategory;
                        if (fieldMap.accountNumber >= 0) newRow[fieldMap.accountNumber] = mapping.accountNumber || "";
                        if (fieldMap.accountName >= 0) newRow[fieldMap.accountName] = mapping.accountName || "";
                        if (fieldMap.amount >= 0) newRow[fieldMap.amount] = amount;
                        if (fieldMap.expenseReview >= 0) newRow[fieldMap.expenseReview] = mapping.expenseReview || "";
                        cleanRows.push(newRow);
                    });
                }
            }
            console.log("DEBUG prepareValidationData - Clean rows generated:", cleanRows.length);
            console.log("DEBUG prepareValidationData - PR_Data_Clean total:", prDataTotal, "Clean total:", cleanTotal);
            console.log("DEBUG prepareValidationData - columnCount:", columnCount, "cleanRange.address:", cleanRange.address);
            // Clear only data rows (Row 2+), preserve headers
            if (!cleanRange.isNullObject && cleanRange.address) {
                console.log("DEBUG prepareValidationData - Clearing data rows...");
                const existingRowCount = Math.max(0, (cleanRange.rowCount || 0) - 1);
                const rowsToClear = Math.max(1, existingRowCount);
                const dataBodyRange = cleanSheet.getRangeByIndexes(1, 0, rowsToClear, columnCount);
                dataBodyRange.clear();
                await context.sync();
                console.log("DEBUG prepareValidationData - Data rows cleared");
            }
            // Write data starting at Row 2
            console.log("DEBUG prepareValidationData - About to write", cleanRows.length, "rows with", columnCount, "columns");
            if (cleanRows.length > 0) {
                const targetRange = cleanSheet.getRangeByIndexes(1, 0, cleanRows.length, columnCount);
                targetRange.values = cleanRows;
                console.log("DEBUG prepareValidationData - Data written to PR_Data_Clean");
            } else {
                console.log("DEBUG prepareValidationData - No rows to write!");
            }
            await context.sync();
            return { prDataTotal, cleanTotal };
        });
        updateValidationState({
            loading: false,
            lastError: null,
            prDataTotal: result.prDataTotal,
            cleanTotal: result.cleanTotal
        });
    } catch (error) {
        console.warn("Validate & Reconcile: unable to prepare PR_Data_Clean", error);
        updateValidationState({
            loading: false,
            prDataTotal: null,
            cleanTotal: null,
            lastError: "Unable to prepare reconciliation data. Try again."
        });
    }
}

function parseRosterValues(values) {
    const result = {
        activeCount: 0,
        departmentCount: 0,
        employeeMap: new Map()
    };
    if (!values || !values.length) return result;
    const { headers, dataStartIndex } = findHeaderRow(values, ["employee"]);
    if (!headers.length || dataStartIndex == null) return result;
    const employeeIdx = findEmployeeIndex(headers);
    const terminationIdx = headers.findIndex((header) => header.includes("termination"));
    const departmentIdx = pickDepartmentIndex(headers);
    if (employeeIdx === -1) return result;
    const activeSet = new Set();

    for (let i = dataStartIndex; i < values.length; i += 1) {
        const row = values[i];
        const employee = row[employeeIdx];
        const key = normalizeKey(employee);
        if (!key || isNoiseName(key)) continue;
        const termination = terminationIdx >= 0 ? row[terminationIdx] : "";
        const department = departmentIdx >= 0 ? row[departmentIdx] : "";
        const isActive = !normalizeString(termination);
        if (isActive && !activeSet.has(key)) {
            activeSet.add(key);
            result.activeCount += 1;
        }
        if (department) {
            result.departmentCount += 1;
        }
        if (!result.employeeMap.has(key)) {
            result.employeeMap.set(key, {
                name: normalizeString(employee) || key,
                department: normalizeString(department),
                termination: termination
            });
        }
    }
    return result;
}

function parsePayrollValues(values) {
    const result = {
        totalEmployees: 0,
        departmentCount: 0,
        employeeMap: new Map()
    };
    if (!values || !values.length) return result;
    const { headers, dataStartIndex } = findHeaderRow(values, ["employee"]);
    if (!headers.length || dataStartIndex == null) return result;
    const employeeIdx = findEmployeeIndex(headers);
    const departmentIdx = pickDepartmentIndex(headers);
    if (employeeIdx === -1) return result;
    const employeeSet = new Set();

    for (let i = dataStartIndex; i < values.length; i += 1) {
        const row = values[i];
        const employee = row[employeeIdx];
        const key = normalizeKey(employee);
        if (!key || isNoiseName(key)) continue;
        if (!employeeSet.has(key)) {
            employeeSet.add(key);
            result.totalEmployees += 1;
        }
        const department = departmentIdx >= 0 ? row[departmentIdx] : "";
        if (department) {
            result.departmentCount += 1;
        }
        if (!result.employeeMap.has(key)) {
            result.employeeMap.set(key, {
                name: normalizeString(employee) || key,
                department: normalizeString(department)
            });
        }
    }
    return result;
}

function normalizeHeader(value) {
    return normalizeString(value).toLowerCase();
}

function findEmployeeIndex(headers = []) {
    const preferred = headers.findIndex((header) => header.includes("employee") && header.includes("name"));
    if (preferred >= 0) return preferred;
    const fallback = headers.findIndex((header) => header.includes("employee"));
    return fallback;
}

function findHeaderRow(rows, requiredTokens = []) {
    let headerRow = [];
    let headerIndex = null;
    (rows || []).some((row, index) => {
        const normalized = (row || []).map(normalizeHeader);
        const hasRequired = requiredTokens.every((token) => normalized.some((cell) => cell.includes(token)));
        if (hasRequired) {
            headerRow = normalized;
            headerIndex = index;
            return true;
        }
        return false;
    });
    return {
        headers: headerRow,
        dataStartIndex: headerIndex != null ? headerIndex + 1 : null
    };
}

function normalizeString(value) {
    return value == null ? "" : String(value).trim();
}

/**
 * Normalize a value for use as a lookup key (lowercase, trimmed)
 */
function normalizeKey(value) {
    return normalizeString(value).toLowerCase();
}

function normalizePayrollCategory(value) {
    return normalizeString(value).toLowerCase();
}

function pickDepartmentIndex(headers = []) {
    const candidates = headers.map((h, idx) => ({ idx, value: normalizeHeader(h) }));
    
    // Priority 1: "Department Description" - the actual department name
    const description = candidates.find(({ value }) => 
        value.includes("department") && value.includes("description")
    );
    if (description) {
        console.log("DEBUG pickDepartmentIndex - Using 'Department Description' at index:", description.idx);
        return description.idx;
    }
    
    // Priority 2: "Department Name"
    const deptName = candidates.find(({ value }) => 
        value.includes("department") && value.includes("name")
    );
    if (deptName) {
        console.log("DEBUG pickDepartmentIndex - Using 'Department Name' at index:", deptName.idx);
        return deptName.idx;
    }
    
    // Priority 3: Department but NOT id/code/number (likely a name field)
    const nonId = candidates.find(({ value }) =>
        value.includes("department") && 
        !value.includes("id") && 
        !value.includes("#") && 
        !value.includes("code") &&
        !value.includes("number")
    );
    if (nonId) {
        console.log("DEBUG pickDepartmentIndex - Using non-ID department at index:", nonId.idx);
        return nonId.idx;
    }
    
    // Priority 4: Any department column as fallback
    const fallback = candidates.find(({ value }) => value.includes("department"));
    if (fallback) {
        console.log("DEBUG pickDepartmentIndex - Using fallback department at index:", fallback.idx);
    }
    return fallback ? fallback.idx : -1;
}

// Note: The following headcount-related functions were removed when Step 2 was deprecated:
// - showHeadcountModal
// - closeHeadcountModal
// - updateHeadcountSignoffState  
// - handleHeadcountSignoff
// - enforceHeadcountSkipNote
// - bindHeadcountNotesGuard
// 
// Payroll coverage is now handled in Step 1 via the Payroll Coverage card.

function handleBankAmountInput(event) {
    const inputEl = event?.target && event.target instanceof HTMLInputElement
        ? event.target
        : document.getElementById("bank-amount-input");
    const numeric = parseBankAmount(inputEl?.value);
    const formatted = formatBankInput(numeric);
    if (inputEl) {
        inputEl.value = formatted;
    }
    updateValidationState({ bankAmount: numeric }, { rerender: false });
}

function handlePlugDifference() {
    showToast("Difference resolution will be available soon.", "info");
}

/**
 * Get the completion handler for a step
 * All steps (0-5) advance to the next step when completed
 * Step 6 (Archive) has its own special flow with popup and return to home
 */
function getStepCompleteHandler(stepId) {
    // Step 4 (Archive) is handled separately with archive popup
    if (stepId === 4) return null;
    
    // All other steps advance to the next step
    return () => advanceToNextStep(stepId);
}

/**
 * Advance to the next step after completing the current one
 */
function advanceToNextStep(currentStepId) {
    const currentIndex = WORKFLOW_STEPS.findIndex((step) => step.id === currentStepId);
    if (currentIndex === -1) return;
    
    const nextIndex = currentIndex + 1;
    if (nextIndex < WORKFLOW_STEPS.length) {
        focusStep(nextIndex);
        // Scroll side panel to top
        scrollPanelToTop();
    }
}

/**
 * Scroll the side panel to the top
 */
function scrollPanelToTop() {
    // Try multiple selectors for the scrollable container
    const containers = [
        document.querySelector('.pf-root'),
        document.querySelector('.pf-step-guide'),
        document.body
    ];
    
    for (const container of containers) {
        if (container) {
            container.scrollTo({ top: 0, behavior: 'smooth' });
        }
    }
    
    // Also scroll window
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Legacy handlers (kept for backwards compatibility)
function handleValidationComplete() {
    advanceToNextStep(3);
}

function handleExpenseReviewComplete() {
    advanceToNextStep(4);
}

/**
 * Archive workflow - MUST run in order to prevent data loss:
 * 1. Archive payroll tabs to new workbook (user chooses save location)
 * 2. Update PR_Archive_Summary (replace oldest period with current)
 * 3. Clear working data from PR_Data_Clean, PR_Expense_Review, PR_JE_Draft
 * 4. Clear non-permanent step notes
 * 5. Reset non-permanent config values
 */
// Global handler for archive button (inline onclick)
window.handleArchiveClick = async function() {
    console.log("[Archive] handleArchiveClick triggered!");
    showToast("Starting archive process...", "info", 2000);
    try {
        await handleArchiveRun();
    } catch (error) {
        console.error("[Archive] handleArchiveClick error:", error);
        showToast("Archive failed: " + error.message, "error", 8000);
    }
};

async function handleArchiveRun() {
    console.log("[Archive] handleArchiveRun called");
    
    if (!hasExcelRuntime()) {
        showToast("Excel runtime is unavailable.", "error");
        return;
    }
    
    // Confirm before proceeding
    const confirmed = await showConfirm(
        "This will:\n\n" +
        "• Download an Excel archive file\n" +
        "• Update PR_Archive_Summary\n" +
        "• Clear working data from all sheets\n" +
        "• Reset non-permanent notes & config\n\n" +
        "Make sure you've completed all review steps.",
        {
            title: "Archive Payroll Run",
            icon: "📦",
            confirmText: "Archive Now",
            cancelText: "Not Yet"
        }
    );
    
    if (!confirmed) {
        console.log("[Archive] User cancelled");
        showToast("Archive cancelled", "info", 2000);
        return;
    }
    
    console.log("[Archive] User confirmed, starting archive process...");
    
    try {
        // ═══════════════════════════════════════════════════════════════════
        // STEP 1: Create archive copy of workbook
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 1: Creating archive workbook...");
        
        const archiveSuccess = await createArchiveWorkbook();
        if (!archiveSuccess) {
            console.log("[Archive] Archive cancelled or failed");
            return;
        }
        
        console.log("[Archive] Step 1 complete: Archive workbook created/user confirmed backup");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 2: Update PR_Archive_Summary with current period
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 2: Updating PR_Archive_Summary...");
        
        await updateArchiveSummary();
        
        console.log("[Archive] Step 2 complete: Archive summary updated");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 3: Clear working data from payroll sheets
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 3: Clearing working data...");
        
        await clearWorkingData();
        
        console.log("[Archive] Step 3 complete: Working data cleared");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 4: Clear non-permanent step notes
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 4: Clearing non-permanent notes...");
        
        await clearNonPermanentNotes();
        
        console.log("[Archive] Step 4 complete: Notes cleared");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 5: Reset non-permanent config values
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 5: Resetting config values...");
        
        await resetNonPermanentConfig();
        
        console.log("[Archive] Step 5 complete: Config reset");
        
        // ═══════════════════════════════════════════════════════════════════
        // COMPLETE
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Archive workflow complete!");
        
        // Reload config and re-render
        await loadConfigurationValues();
        renderApp();
        
        // Show "save complete" prompt - waits for user to confirm they've saved
        showSaveCompletePrompt();
        
    } catch (error) {
        console.error("[Archive] Error during archive:", error);
        showToast(
            "Archive Error: " + error.message,
            "error",
            10000
        );
    }
}

/**
 * Step 1: Export payroll data to Excel file for archiving
 * Downloads a proper .xlsx file with all sheets
 */
async function createArchiveWorkbook() {
    try {
        // Get current payroll date for filename
        const payrollDate = getPayrollDateValue() || new Date().toISOString().split("T")[0];
        const filename = `Payroll_Archive_${payrollDate}.xlsx`;
        
        console.log("[Archive] Creating Excel archive file...");
        
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const sourceSheets = workbook.worksheets;
            sourceSheets.load("items/name");
            await context.sync();
            
            // Sheets to archive (in order they'll appear in the file)
            const sheetsToArchive = [
                SHEET_NAMES.JE_DRAFT,
                SHEET_NAMES.DATA_CLEAN,
                SHEET_NAMES.EXPENSE_REVIEW
            ];
            
            // Create new workbook using SheetJS
            const newWorkbook = XLSX.utils.book_new();
            let sheetsAdded = 0;

            // Summary sheet (friendly metadata)
            const summaryRows = [];
            summaryRows.push(["Payroll Archive Summary"]);
            summaryRows.push(["Archived At", new Date().toISOString()]);
            summaryRows.push(["Payroll Date", payrollDate]);
            summaryRows.push(["Accounting Period", getConfigValue("PR_Accounting_Period") || ""]);
            summaryRows.push(["Journal Entry ID", getConfigValue("PR_Journal_Entry_ID") || ""]);
            summaryRows.push([]);
            // Get sign-off/notes from each step using STEP_NOTES_FIELDS
            const cfgFields = STEP_NOTES_FIELDS[0];
            const importFields = STEP_NOTES_FIELDS[1];
            const reviewFields = STEP_NOTES_FIELDS[2];
            const jeFields = STEP_NOTES_FIELDS[3];
            // Use main PR_Reviewer field (set in Config step) as the primary reviewer
            const mainReviewer = getConfigValue(CONFIG_REVIEWER_FIELD) || "";
            summaryRows.push(["Reviewer", mainReviewer]);
            summaryRows.push(["Config Sign-off", cfgFields ? getConfigValue(cfgFields.signOff) || "" : ""]);
            summaryRows.push(["Config Notes", cfgFields ? getConfigValue(cfgFields.note) || "" : ""]);
            summaryRows.push([]);
            summaryRows.push(["Import Sign-off", importFields ? getConfigValue(importFields.signOff) || "" : ""]);
            summaryRows.push(["Import Notes", importFields ? getConfigValue(importFields.note) || "" : ""]);
            summaryRows.push([]);
            summaryRows.push(["Expense Review Sign-off", reviewFields ? getConfigValue(reviewFields.signOff) || "" : ""]);
            summaryRows.push(["Expense Review Notes", reviewFields ? getConfigValue(reviewFields.note) || "" : ""]);
            summaryRows.push([]);
            summaryRows.push(["JE Sign-off", jeFields ? getConfigValue(jeFields.signOff) || "" : ""]);
            summaryRows.push(["JE Notes", jeFields ? getConfigValue(jeFields.note) || "" : ""]);
            summaryRows.push([]);
            summaryRows.push(["Generated By", "Payroll Recorder"]);
            summaryRows.push(["Archive File", filename]);

            const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
            
            // Format summary sheet with proper column widths
            setXlsxColumnWidths(summarySheet, [
                XLSX_COLUMN_WIDTHS.extraWide,  // Label column
                XLSX_COLUMN_WIDTHS.description  // Value column
            ]);
            
            XLSX.utils.book_append_sheet(newWorkbook, summarySheet, "Archive_Summary");
            sheetsAdded++;
            
            for (const sheetName of sheetsToArchive) {
                const sourceSheet = sourceSheets.items.find(s => s.name === sheetName);
                if (!sourceSheet) {
                    console.log(`[Archive] Sheet not found: ${sheetName}`);
                    continue;
                }
                
                const usedRange = sourceSheet.getUsedRangeOrNullObject();
                usedRange.load("values");
                await context.sync();
                
                if (!usedRange.isNullObject && usedRange.values && usedRange.values.length > 0) {
                    // Convert to SheetJS worksheet
                    const worksheet = XLSX.utils.aoa_to_sheet(usedRange.values);
                    
                    // Apply formatting based on sheet type
                    if (usedRange.values.length > 0) {
                        if (sheetName === SHEET_NAMES.EXPENSE_REVIEW) {
                            // PR_Expense_Review has a complex multi-section layout
                            // Use special formatter that scans cell content
                            formatExpenseReviewSheet(worksheet, usedRange.values);
                            console.log(`[Archive] Applied Expense Review formatting to ${sheetName}`);
                        } else {
                            // Standard sheets with headers in row 1
                            const headers = usedRange.values[0];
                            const rowCount = usedRange.values.length;
                            formatXlsxWorksheet(worksheet, headers, rowCount, {
                                autoFormat: true,
                                autoSize: true
                            });
                        }
                    }
                    
                    // Add to workbook
                    XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);
                    sheetsAdded++;
                    console.log(`[Archive] Added sheet: ${sheetName} (${usedRange.values.length} rows)`);
                }
            }
            
            if (sheetsAdded === 0) {
                showToast("No data to archive. Please complete the payroll workflow first.", "error");
                return false;
            }
            
            // Generate Excel file and download
            try {
                console.log(`[Archive] Writing ${sheetsAdded} sheets to Excel...`);
                const excelBuffer = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });
                console.log(`[Archive] Excel buffer created: ${excelBuffer.byteLength} bytes`);
                
                const blob = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                console.log(`[Archive] Blob created: ${blob.size} bytes`);
                
                downloadFile(blob, filename);
                
                console.log(`[Archive] Downloaded: ${filename} with ${sheetsAdded} sheets`);
                
                showToast(`📥 Archive downloaded: ${filename} (${sheetsAdded} sheets)`, "success", 5000);
                
                return true;
            } catch (downloadError) {
                console.error("[Archive] Download error:", downloadError);
                throw new Error(`Failed to write Excel file: ${downloadError.message}`);
            }
        });
        
    } catch (error) {
        console.error("[Archive] Error creating archive:", error);
        showToast("Archive Export Error: " + error.message, "error", 8000);
        
        // Ask if user wants to continue anyway
        return await showConfirm(
            "Archive download failed.\n\n" +
            "Do you want to continue with clearing the data?\n\n" +
            "Make sure you have saved a backup first!"
        );
    }
}

/**
 * Trigger browser download of a file
 */
function downloadFile(blob, filename) {
    try {
        if (!blob || blob.size === 0) {
            throw new Error("Blob is empty");
        }
        
        const url = URL.createObjectURL(blob);
        console.log(`[Download] URL created: ${url}`);
        
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        link.style.visibility = "hidden";
        
        document.body.appendChild(link);
        console.log("[Download] Link appended to DOM");
        
        link.click();
        console.log("[Download] Click triggered");
        
        document.body.removeChild(link);
        console.log("[Download] Link removed from DOM");
        
        // Clean up the URL object
        setTimeout(() => {
            URL.revokeObjectURL(url);
            console.log("[Download] URL revoked");
        }, 100);
        
    } catch (error) {
        console.error("[Download] Error:", error);
        throw error;
    }
}

/**
 * Parse a period date from various formats (ISO, Excel serial, MM/DD/YYYY, etc.)
 */
function parsePeriodDate(value) {
    if (!value) return null;
    
    // If it's already a Date
    if (value instanceof Date) {
        return isNaN(value.getTime()) ? null : value;
    }
    
    // If it's an Excel serial number (a number like 45678)
    if (typeof value === 'number' && value > 1000 && value < 100000) {
        const date = convertExcelDate(value);
        return date;
    }
    
    const str = String(value).trim();
    if (!str) return null;
    
    // ISO format: YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        const [y, m, d] = str.split('-').map(Number);
        return new Date(y, m - 1, d);
    }
    
    // MM/DD/YYYY format
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
        const [m, d, y] = str.split('/').map(Number);
        return new Date(y, m - 1, d);
    }
    
    // Try standard Date parsing as fallback
    const parsed = new Date(str);
    return isNaN(parsed.getTime()) ? null : parsed;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ARCHIVE SUMMARY HARDENING
// PR_Archive_Summary is a forward-compatible history store for Expense Review.
// Uses ada_payroll_column_dictionary for measure classification (not heuristics).
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Classify archive columns using taxonomy tables.
 * 
 * DATA SOURCES:
 * - Dimensions: public.ada_payroll_dimensions (normalized_dimension)
 * - Measures: public.ada_payroll_column_dictionary (pf_column_name, data_type='number')
 * 
 * Required dimension columns for archive:
 * - Pay_Date (or Payroll_Date) - period key
 * - Employee_Name - identity
 * - Department_Name - grouping
 * 
 * @param {string[]} headers - PR_Data_Clean headers (canonical pf_column_name)
 * @returns {{ dimensions: Set<string>, measures: Set<string>, unknown: Set<string>, dimensionIndices: number[], measureIndices: number[] }}
 */
function classifyArchiveColumns(headers) {
    const dimensions = new Set();
    const measures = new Set();
    const unknown = new Set();
    const dimensionIndices = [];
    const measureIndices = [];
    
    // Use cached taxonomy if available
    // NOTE: taxonomy keys are lowercase for case-insensitive lookup
    const taxonomy = expenseTaxonomyCache.loaded ? expenseTaxonomyCache : null;
    
    headers.forEach((header, idx) => {
        const normalizedHeader = String(header || "").trim();
        const headerLower = normalizedHeader.toLowerCase();
        
        // Check if it's a dimension (case-insensitive)
        if (taxonomy?.dimensions?.has(headerLower)) {
            dimensions.add(normalizedHeader);
            dimensionIndices.push(idx);
        }
        // Check if it's a measure in dictionary (case-insensitive)
        else if (taxonomy?.measures?.[headerLower]) {
            measures.add(normalizedHeader);
            measureIndices.push(idx);
        }
        // Fallback: treat as dimension if known dimension keywords
        else if (
            headerLower.includes("employee") || 
            headerLower.includes("department") ||
            headerLower.includes("location") ||
            headerLower.includes("date") ||
            headerLower.includes("period") ||
            headerLower.includes("cost_center") ||
            headerLower.includes("entity")
        ) {
            dimensions.add(normalizedHeader);
            dimensionIndices.push(idx);
            console.warn(`[Archive] Fallback dimension classification: ${normalizedHeader}`);
        }
        // Fallback: check if it looks numeric by examining first few data rows
        else {
            // Mark as unknown for now - will be handled by caller
            unknown.add(normalizedHeader);
        }
    });
    
    return { dimensions, measures, unknown, dimensionIndices, measureIndices };
}

/**
 * Ensure PR_Archive_Summary schema is forward-compatible.
 * Adds new measure columns from PR_Data_Clean that don't exist in archive.
 * Backfills 0 for existing rows in new columns.
 * 
 * SCHEMA EVOLUTION RULES:
 * - Never delete columns (preserve history)
 * - Add new measure columns at end
 * - Backfill 0 for existing rows
 * - Preserve dimension columns: Pay_Date, Employee_Name, Department_Name
 * 
 * @param {string[]} cleanHeaders - Headers from PR_Data_Clean
 * @param {string[]} archiveHeaders - Existing headers from PR_Archive_Summary
 * @param {any[][]} archiveData - Existing archive data rows
 * @returns {{ 
 *   mergedHeaders: string[], 
 *   updatedArchiveData: any[][], 
 *   addedColumns: string[],
 *   dimensionColumns: string[],
 *   measureColumns: string[]
 * }}
 */
function ensureArchiveSummarySchema(cleanHeaders, archiveHeaders, archiveData) {
    const addedColumns = [];
    const archiveHeaderSet = new Set(archiveHeaders.map(h => String(h || "").trim()));
    
    // Classify clean headers
    const classification = classifyArchiveColumns(cleanHeaders);
    
    // Required dimension columns (must exist)
    const requiredDimensions = ["Payroll_Date", "Pay_Date", "Employee_Name", "Department_Name"];
    const foundDimensions = requiredDimensions.filter(d => 
        cleanHeaders.some(h => String(h || "").trim().toLowerCase() === d.toLowerCase())
    );
    
    console.log(`[Archive Schema] Required dimensions found: ${foundDimensions.join(", ") || "NONE"}`);
    console.log(`[Archive Schema] Measure columns identified: ${classification.measures.size}`);
    console.log(`[Archive Schema] Unknown columns: ${Array.from(classification.unknown).join(", ") || "none"}`);
    
    // Determine columns to add (measures from clean data not in archive)
    const columnsToAdd = cleanHeaders.filter(h => {
        const header = String(h || "").trim();
        // Add if: (1) not already in archive, (2) is a measure OR in clean but not classified
        const notInArchive = !archiveHeaderSet.has(header);
        const isMeasure = classification.measures.has(header);
        const isUnknown = classification.unknown.has(header);
        return notInArchive && (isMeasure || isUnknown);
    });
    
    // Build merged headers: archive headers + new columns
    const mergedHeaders = [...archiveHeaders];
    columnsToAdd.forEach(col => {
        const colName = String(col || "").trim();
        mergedHeaders.push(colName);
        addedColumns.push(colName);
        console.log(`[Archive Schema] Adding new column: ${colName}`);
    });
    
    // Backfill 0 for new columns in existing archive data
    const updatedArchiveData = archiveData.map(row => {
        const extendedRow = [...row];
        // Pad with 0 for each new column
        for (let i = 0; i < addedColumns.length; i++) {
            extendedRow.push(0);
        }
        return extendedRow;
    });
    
    // Track final column classifications
    const dimensionColumns = mergedHeaders.filter(h => classification.dimensions.has(String(h || "").trim()));
    const measureColumns = mergedHeaders.filter(h => classification.measures.has(String(h || "").trim()));
    
    return {
        mergedHeaders,
        updatedArchiveData,
        addedColumns,
        dimensionColumns,
        measureColumns
    };
}

/**
 * Log archive validation summary (advisory, does not block).
 * 
 * @param {{
 *   periodsRetained: number,
 *   currentPeriodKey: string,
 *   rowsArchived: number,
 *   measureColumnCount: number,
 *   columnsAdded: string[],
 *   dimensionColumnsPresent: boolean
 * }} summary
 */
function logArchiveValidation(summary) {
    console.log("════════════════════════════════════════════════════════════");
    console.log("[Archive Validation Summary]");
    console.log(`  Periods retained: ${summary.periodsRetained}`);
    console.log(`  Current period key: ${summary.currentPeriodKey}`);
    console.log(`  Rows archived: ${summary.rowsArchived}`);
    console.log(`  Measure columns: ${summary.measureColumnCount}`);
    console.log(`  Columns added: ${summary.columnsAdded.length > 0 ? summary.columnsAdded.join(", ") : "none"}`);
    console.log(`  Required dimensions present: ${summary.dimensionColumnsPresent ? "YES" : "WARNING - MISSING"}`);
    console.log("════════════════════════════════════════════════════════════");
    
    // Advisory warnings
    if (!summary.dimensionColumnsPresent) {
        console.warn("[Archive] WARNING: Some required dimension columns may be missing (Pay_Date, Employee_Name, Department_Name)");
    }
    if (summary.columnsAdded.length > 0) {
        console.log(`[Archive] Schema evolved: ${summary.columnsAdded.length} new column(s) added and backfilled with 0`);
    }
}

// =============================================================================
// ARCHIVE BUCKET SUMMARY COLUMNS
// These columns store pre-aggregated bucket totals at archive time, preserving
// the classification as it was when the period was processed.
// =============================================================================
const ARCHIVE_SUMMARY_COLUMNS = [
    "_Archive_Fixed_Total",
    "_Archive_Variable_Total", 
    "_Archive_Burden_Total",
    "_Archive_Period_Total",
    "_Archive_Headcount"
];

/**
 * Calculate bucket totals for archiving from PR_Data_Clean
 * Uses current taxonomy classification - this is the "snapshot" of how
 * the period was classified at archive time.
 * 
 * @param {string[]} headers - PR_Data_Clean headers
 * @param {any[][]} dataRows - PR_Data_Clean data (excluding header)
 * @returns {{ fixed: number, variable: number, burden: number, total: number, headcount: number }}
 */
function calculateArchiveBucketTotals(headers, dataRows) {
    const taxonomy = expenseTaxonomyCache.loaded ? expenseTaxonomyCache : null;
    const headersLower = headers.map(h => String(h || "").toLowerCase().trim());
    
    let fixedTotal = 0;
    let variableTotal = 0;
    let burdenTotal = 0;
    const employees = new Set();
    
    // Find employee name column for headcount
    const employeeIdx = headersLower.findIndex(h => 
        h.includes("employee") && h.includes("name")
    );
    
    // Classify and sum each column
    headers.forEach((header, colIdx) => {
        const headerLower = String(header || "").toLowerCase().trim();
        
        // Skip if it's a dimension column
        if (taxonomy?.dimensions?.has(headerLower)) {
            return;
        }
        
        // Get dictionary metadata
        const dictEntry = taxonomy?.measures?.[headerLower];
        if (!dictEntry) {
            return; // Not in dictionary, skip
        }
        
        // Only include if side = 'er' (employer expense)
        const side = dictEntry.side || "er";
        if (side === "ee" || side === "na") {
            return; // Employee deduction or N/A, skip
        }
        
        // Get bucket classification
        const bucket = normalizeBucketName(dictEntry.bucket);
        const sign = dictEntry.sign ?? 1;
        
        // Sum this column across all rows
        let columnSum = 0;
        dataRows.forEach(row => {
            const val = Number(row[colIdx]) || 0;
            columnSum += val * sign;
        });
        
        // Add to appropriate bucket
        if (bucket === "FIXED") {
            fixedTotal += columnSum;
        } else if (bucket === "VARIABLE") {
            variableTotal += columnSum;
        } else if (bucket === "BURDEN" || bucket === "BENEFIT") {
            burdenTotal += columnSum;
        }
        // Other buckets (TAX, DEDUCTION, OTHER) are excluded from employer expense
    });
    
    // Count unique employees
    if (employeeIdx >= 0) {
        dataRows.forEach(row => {
            const empName = String(row[employeeIdx] || "").trim();
            if (empName) {
                employees.add(empName);
            }
        });
    }
    
    const total = fixedTotal + variableTotal + burdenTotal;
    
    console.log(`[Archive] Bucket totals calculated: FIXED=${fixedTotal.toLocaleString()}, VARIABLE=${variableTotal.toLocaleString()}, BURDEN=${burdenTotal.toLocaleString()}, TOTAL=${total.toLocaleString()}, HEADCOUNT=${employees.size}`);
    
    return {
        fixed: fixedTotal,
        variable: variableTotal,
        burden: burdenTotal,
        total,
        headcount: employees.size
    };
}

/**
 * Step 2: Update PR_Archive_Summary
 * - Copy all rows from PR_Data_Clean to PR_Archive_Summary
 * - Calculate and store bucket totals (FIXED/VARIABLE/BURDEN) for historical comparison
 * - Ensure schema evolution (add new measure columns, backfill 0)
 * - Remove oldest period's rows if more than 5 distinct periods exist
 * - Idempotent: replaces rows for current period key
 * 
 * PRIOR PERIOD SOURCE: PR_Archive_Summary sheet
 * PERIOD KEY: Payroll_Date (or Pay_Date) column, normalized to YYYY-MM-DD
 */
async function updateArchiveSummary() {
    await Excel.run(async (context) => {
        const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
        let archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
        let archiveTotalsSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_TOTALS);
        const reviewSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.EXPENSE_REVIEW);
        
        cleanSheet.load("isNullObject, name");
        archiveSheet.load("isNullObject, name");
        archiveTotalsSheet.load("isNullObject, name");
        reviewSheet.load("isNullObject, name");
        await context.sync();
        
        if (cleanSheet.isNullObject) {
            console.warn("[Archive] PR_Data_Clean not found - skipping");
            return;
        }
        
        // CREATE the archive sheet if it doesn't exist
        if (archiveSheet.isNullObject) {
            console.log("[Archive] PR_Archive_Summary not found - CREATING it...");
            archiveSheet = context.workbook.worksheets.add(SHEET_NAMES.ARCHIVE_SUMMARY);
            
            // Format header row
            const headerRange = archiveSheet.getRange("1:1");
            formatSheetHeaders(headerRange);
            
            await context.sync();
            console.log("[Archive] PR_Archive_Summary sheet created successfully");
        }

        if (archiveTotalsSheet.isNullObject) {
            console.log("[Archive] PR_Archive_Totals not found - CREATING it...");
            archiveTotalsSheet = context.workbook.worksheets.add(SHEET_NAMES.ARCHIVE_TOTALS);

            const totalsHeaderRange = archiveTotalsSheet.getRange("1:1");
            formatSheetHeaders(totalsHeaderRange);

            await context.sync();
            console.log("[Archive] PR_Archive_Totals sheet created successfully");
        }
        
        // Get current period data from PR_Data_Clean
        const cleanRange = cleanSheet.getUsedRangeOrNullObject();
        cleanRange.load("values");
        await context.sync();
        
        if (cleanRange.isNullObject || !cleanRange.values || cleanRange.values.length < 2) {
            console.warn("[Archive] PR_Data_Clean is empty - skipping archive summary update");
            return;
        }
        
        const cleanHeaders = cleanRange.values[0];
        const cleanData = cleanRange.values.slice(1);
        console.log(`[Archive] PR_Data_Clean has ${cleanData.length} rows to archive`);
        
        // Get current period date
        const rawCurrentDate = getPayrollDateValue();
        const parsedCurrentDate = parsePeriodDate(rawCurrentDate);
        const currentPeriodDate = parsedCurrentDate 
            ? formatDateFromDate(parsedCurrentDate)
            : String(rawCurrentDate || "").trim();
        
        console.log(`[Archive] Current period date (raw): ${rawCurrentDate}, normalized: ${currentPeriodDate}`);
        
        // Read department breakdown table from PR_Expense_Review to get dynamic headers
        let deptMetricHeaders = [];
        let deptTableData = null;
        
        if (!reviewSheet.isNullObject) {
            console.log("[Archive] Reading department table headers from PR_Expense_Review");
            
            // Try named range first
            let deptRange = null;
            try {
                deptRange = context.workbook.names.getItemOrNullObject("PR_Dept_Breakdown");
                deptRange.load("name, value");
                await context.sync();
                
                if (!deptRange.isNullObject) {
                    console.log("[Archive] Using named range: PR_Dept_Breakdown");
                    const rangeAddress = deptRange.value;
                    const actualRange = reviewSheet.getRange(rangeAddress);
                    actualRange.load("values");
                    await context.sync();
                    deptTableData = actualRange.values;
                }
            } catch (err) {
                console.warn("[Archive] Named range PR_Dept_Breakdown not found, falling back to scan");
            }
            
            // Fallback: scan for department table
            if (!deptTableData) {
                console.log("[Archive] Scanning PR_Expense_Review for department table");
                const reviewRange = reviewSheet.getUsedRangeOrNullObject();
                reviewRange.load("values");
                await context.sync();
                
                if (!reviewRange.isNullObject && reviewRange.values && reviewRange.values.length > 0) {
                    const reviewData = reviewRange.values;
                    
                    // Find section marker first for bounded search
                    let searchStartIdx = 0;
                    for (let i = 0; i < reviewData.length; i++) {
                        const cellValue = String(reviewData[i][0] || "").toUpperCase();
                        if (cellValue.includes("CURRENT PERIOD BREAKDOWN") && cellValue.includes("DEPARTMENT")) {
                            searchStartIdx = i + 1;
                            console.log(`[Archive] Found department section marker at row ${i}, starting search at row ${searchStartIdx}`);
                            break;
                        }
                    }
                    
                    // Find header row with multiple expected metrics
                    let deptHeaderRowIdx = -1;
                    for (let i = searchStartIdx; i < reviewData.length; i++) {
                        const row = reviewData[i];
                        const rowStr = row.map(c => String(c || "").toLowerCase()).join("|");
                        
                        // Must contain Department AND at least 2 of the expected metrics
                        const hasDept = rowStr.includes("department");
                        const hasGrossPay = rowStr.includes("gross") && rowStr.includes("pay");
                        const hasBurden = rowStr.includes("burden");
                        const hasAllIn = rowStr.includes("all") && rowStr.includes("in");
                        
                        const metricCount = [hasGrossPay, hasBurden, hasAllIn].filter(Boolean).length;
                        
                        if (hasDept && metricCount >= 2) {
                            deptHeaderRowIdx = i;
                            console.log(`[Archive] Found department header row at ${i} (matched ${metricCount} metrics)`);
                            break;
                        }
                    }
                    
                    if (deptHeaderRowIdx >= 0) {
                        // Extract table: header row + data rows until blank/TOTAL
                        const headerRow = reviewData[deptHeaderRowIdx];
                        const tableRows = [headerRow];
                        
                        for (let i = deptHeaderRowIdx + 1; i < reviewData.length; i++) {
                            const deptName = String(reviewData[i][0] || "").trim();
                            if (!deptName) break;
                            tableRows.push(reviewData[i]);
                            if (deptName === "TOTAL") break;
                        }
                        
                        deptTableData = tableRows;
                        console.log(`[Archive] Extracted department table: ${tableRows.length} rows (1 header + ${tableRows.length - 1} data)`);
                    } else {
                        console.warn("[Archive] ⚠ Could not find department table header in PR_Expense_Review");
                    }
                }
            }
            
            // Extract metric headers from dept table (exclude Department column)
            if (deptTableData && deptTableData.length > 0) {
                const deptHeaders = deptTableData[0].map(h => String(h || "").trim());
                const deptColIdx = deptHeaders.findIndex(h => 
                    h.toLowerCase().replace(/[^a-z]/g, "") === "department"
                );
                
                deptMetricHeaders = deptHeaders.filter((h, idx) => idx !== deptColIdx && h !== "");
                console.log(`[Archive] Extracted ${deptMetricHeaders.length} dept metric headers: ${deptMetricHeaders.join(", ")}`);
            }
        }
        
        // Get existing archive data
        const archiveRange = archiveSheet.getUsedRangeOrNullObject();
        archiveRange.load("values,rowCount,columnCount");
        await context.sync();
        
        // Initialize or expand archive headers
        let archiveHeaders = [];
        if (!archiveRange.isNullObject && archiveRange.values && archiveRange.values.length > 0) {
            archiveHeaders = archiveRange.values[0];
            
            // Check if we need to expand headers with new dept metrics or add Row_Type
            const existingHeaders = archiveHeaders.map(h => String(h || "").toLowerCase().trim());
            
            // Add Row_Type column if missing (for backward compatibility with old archives)
            const hasRowType = existingHeaders.some(h => h === "row_type" || h === "rowtype");
            if (!hasRowType) {
                console.log("[Archive] Adding Row_Type column to existing archive headers");
                // Insert Row_Type after Pay_Date (index 1)
                archiveHeaders = [archiveHeaders[0], "Row_Type", ...archiveHeaders.slice(1)];
            }
            
            const newMetrics = deptMetricHeaders.filter(metric => {
                const normalized = metric.toLowerCase().replace(/[^a-z0-9]/g, "");
                return !existingHeaders.some(existing => 
                    existing.replace(/[^a-z0-9]/g, "") === normalized
                );
            });
            
            if (newMetrics.length > 0) {
                console.log(`[Archive] Expanding archive headers with ${newMetrics.length} new metrics: ${newMetrics.join(", ")}`);
                archiveHeaders = [...archiveHeaders, ...newMetrics];
            }
        } else {
            console.log("[Archive] Initializing archive headers from PR_Data_Clean + dynamic dept metrics");
            archiveHeaders = ["Pay_Date", "Row_Type", ...cleanHeaders, ...deptMetricHeaders];
        }
        
        // Group existing archive data by period date
        const archiveHeadersLower = archiveHeaders.map(h => String(h || "").toLowerCase().trim());
        const archiveDateIdx = archiveHeadersLower.findIndex(h => 
            h === "pay_date" || h === "payroll_date" || h.includes("pay_period")
        );
        
        const periodMap = new Map();
        if (!archiveRange.isNullObject && archiveRange.values && archiveRange.values.length > 1) {
            const archiveData = archiveRange.values.slice(1);
            archiveData.forEach(row => {
                const rawDate = archiveDateIdx >= 0 ? row[archiveDateIdx] : "";
                const parsedDate = parsePeriodDate(rawDate);
                const periodKey = parsedDate ? formatDateFromDate(parsedDate) : String(rawDate || "").trim();
                
                if (periodKey) {
                    if (!periodMap.has(periodKey)) {
                        periodMap.set(periodKey, []);
                    }
                    periodMap.get(periodKey).push(row);
                }
            });
        }
        
        console.log(`[Archive] Found ${periodMap.size} existing periods in archive`);
        
        // Remove current period if it exists (idempotent replace)
        if (currentPeriodDate && periodMap.has(currentPeriodDate)) {
            console.log(`[Archive] Removing existing data for period: ${currentPeriodDate} (idempotent replace)`);
            periodMap.delete(currentPeriodDate);
        }
        
        // Build new archive data by inserting Pay_Date and Row_Type into each row
        const newArchiveData = cleanData.map(row => {
            return archiveHeaders.map((archiveHeader, colIdx) => {
                const archiveHeaderNorm = String(archiveHeader || "").trim();
                const archiveHeaderLower = archiveHeaderNorm.toLowerCase();
                
                // Insert current period date for Pay_Date column
                if (archiveHeaderLower === "pay_date" || archiveHeaderLower === "payroll_date" || archiveHeaderLower.includes("pay_period")) {
                    return currentPeriodDate;
                }
                
                // Insert row type for Row_Type column
                if (archiveHeaderLower === "row_type" || archiveHeaderLower === "rowtype") {
                    return "Employee";
                }
                
                // Find corresponding column in clean data
                const cleanIdx = cleanHeaders.findIndex(h => String(h || "").trim() === archiveHeaderNorm);
                if (cleanIdx >= 0 && cleanIdx < row.length) {
                    return row[cleanIdx];
                }
                
                // Column exists in archive but not in clean data
                return "";
            });
        });

        // PR_Archive_Summary stores employee-level detail only.
        // Department totals are stored in PR_Archive_Totals (fixed schema).
        console.log(`[Archive] Period ${currentPeriodDate}: ${newArchiveData.length} employee rows (department rows omitted from PR_Archive_Summary)`);

        // Add current period data (employee rows only)
        periodMap.set(currentPeriodDate || `period_${Date.now()}`, newArchiveData);
        
        // Keep only the most recent MAX_PERIODS (5) periods
        const MAX_PERIODS = 5;
        const sortedPeriods = Array.from(periodMap.keys()).sort((a, b) => {
            const dateA = new Date(a);
            const dateB = new Date(b);
            if (!isNaN(dateA) && !isNaN(dateB)) {
                return dateB - dateA; // Descending (newest first)
            }
            return b.localeCompare(a);
        });
        
        console.log(`[Archive] Total periods in archive: ${sortedPeriods.length}`);
        console.log(`[Archive] Periods (newest to oldest): ${sortedPeriods.join(", ")}`);
        
        if (sortedPeriods.length > MAX_PERIODS) {
            const periodsToRemove = sortedPeriods.slice(MAX_PERIODS);
            console.log(`[Archive] ⚠ Dropping ${periodsToRemove.length} oldest period(s): ${periodsToRemove.join(", ")}`);
            periodsToRemove.forEach(key => {
                const rowCount = periodMap.get(key)?.length || 0;
                console.log(`[Archive]   - Removing period ${key} (${rowCount} rows)`);
                periodMap.delete(key);
            });
        } else {
            console.log(`[Archive] ✓ Retaining all ${sortedPeriods.length} periods (under limit of ${MAX_PERIODS})`);
        }
        
        // Flatten all periods back into rows (newest periods first)
        const finalSortedPeriods = Array.from(periodMap.keys()).sort((a, b) => {
            const dateA = new Date(a);
            const dateB = new Date(b);
            if (!isNaN(dateA) && !isNaN(dateB)) {
                return dateB - dateA;
            }
            return b.localeCompare(a);
        });
        
        const allArchiveRows = [];
        finalSortedPeriods.forEach(periodKey => {
            const rows = periodMap.get(periodKey);
            allArchiveRows.push(...rows);
        });

        try {
            const totalsHeaders = [
                "Pay_Date",
                "Row_Type",
                "Department",
                "Fixed Salary",
                "Variable Salary",
                "Gross Pay",
                "Burden",
                "All-In Total",
                "% of Total",
                "Headcount"
            ];

            const normalizeTotalsHeader = (h) => String(h || "").toLowerCase().replace(/[^a-z0-9]/g, "");
            const findDeptHeaderIdx = (headers, targets) => {
                const normalizedHeaders = headers.map(normalizeTotalsHeader);
                for (const target of targets) {
                    const normTarget = normalizeTotalsHeader(target);
                    const idx = normalizedHeaders.findIndex(h => h === normTarget);
                    if (idx >= 0) return idx;
                }
                return -1;
            };

            let currentPeriodTotalsRows = [];

            if (deptTableData && deptTableData.length > 1) {
                const deptHeaders = (deptTableData[0] || []).map(h => String(h || "").trim());
                const deptColIdx = findDeptHeaderIdx(deptHeaders, ["Department"]);

                const fixedSalaryIdx = findDeptHeaderIdx(deptHeaders, ["Fixed Salary", "Fixed", "Base Salary"]);
                const variableSalaryIdx = findDeptHeaderIdx(deptHeaders, ["Variable Salary", "Variable", "Variable Pay"]);
                const grossPayIdx = findDeptHeaderIdx(deptHeaders, ["Gross Pay", "Gross", "Gross Wages"]);
                const burdenIdx = findDeptHeaderIdx(deptHeaders, ["Burden", "Payroll Burden", "Employer Burden"]);
                const allInTotalIdx = findDeptHeaderIdx(deptHeaders, ["All-In Total", "All In Total", "All-In", "Total"]);
                const percentOfTotalIdx = findDeptHeaderIdx(deptHeaders, ["% of Total", "Percent of Total", "Percentage"]);
                const headcountIdx = findDeptHeaderIdx(deptHeaders, ["Headcount", "Head Count", "Employee Count"]);

                let totalAllIn = 0;

                for (let i = 1; i < deptTableData.length; i++) {
                    const row = deptTableData[i] || [];
                    const deptName = deptColIdx >= 0 ? String(row[deptColIdx] || "").trim() : "";
                    if (!deptName) break;

                    const fixed = fixedSalaryIdx >= 0 ? (Number(row[fixedSalaryIdx]) || 0) : 0;
                    const variable = variableSalaryIdx >= 0 ? (Number(row[variableSalaryIdx]) || 0) : 0;
                    const gross = grossPayIdx >= 0 ? (Number(row[grossPayIdx]) || 0) : (fixed + variable);
                    const burden = burdenIdx >= 0 ? (Number(row[burdenIdx]) || 0) : 0;
                    const allIn = allInTotalIdx >= 0 ? (Number(row[allInTotalIdx]) || 0) : (gross + burden);
                    const percent = percentOfTotalIdx >= 0 ? (Number(row[percentOfTotalIdx]) || 0) : 0;
                    const headcount = headcountIdx >= 0 ? (Number(row[headcountIdx]) || 0) : 0;

                    if (deptName === "TOTAL") {
                        totalAllIn = allIn;
                    }

                    currentPeriodTotalsRows.push([
                        currentPeriodDate,
                        deptName === "TOTAL" ? "TOTAL" : "DEPT_TOTAL",
                        deptName,
                        fixed,
                        variable,
                        gross,
                        burden,
                        allIn,
                        percent,
                        headcount
                    ]);

                    if (deptName === "TOTAL") {
                        break;
                    }
                }

                if (totalAllIn > 0) {
                    currentPeriodTotalsRows = currentPeriodTotalsRows.map((r) => {
                        const rowType = String(r[1] || "").trim().toUpperCase();
                        if (rowType === "TOTAL") return r;
                        const existingPercent = Number(r[8]) || 0;
                        if (existingPercent > 0) return r;
                        const allIn = Number(r[7]) || 0;
                        return [...r.slice(0, 8), allIn / totalAllIn, ...r.slice(9)];
                    });
                }
            }

            if (currentPeriodTotalsRows.length === 0) {
                const bucketTotals = calculateArchiveBucketTotals(cleanHeaders, cleanData);
                currentPeriodTotalsRows = [[
                    currentPeriodDate,
                    "TOTAL",
                    "TOTAL",
                    bucketTotals.fixed || 0,
                    bucketTotals.variable || 0,
                    (bucketTotals.fixed || 0) + (bucketTotals.variable || 0),
                    bucketTotals.burden || 0,
                    bucketTotals.total || 0,
                    (bucketTotals.total || 0) ? 1 : 0,
                    bucketTotals.headcount || 0
                ]];
            }

            const totalsRange = archiveTotalsSheet.getUsedRangeOrNullObject();
            totalsRange.load("values,rowCount,columnCount,isNullObject");
            await context.sync();

            const existingTotalsValues = (!totalsRange.isNullObject && totalsRange.values) ? totalsRange.values : [];
            const existingTotalsHeaders = existingTotalsValues.length > 0 ? existingTotalsValues[0] : [];
            const existingTotalsHeadersNorm = existingTotalsHeaders.map(normalizeTotalsHeader);
            const totalsHeadersNorm = totalsHeaders.map(normalizeTotalsHeader);

            const totalsIdxMap = totalsHeadersNorm.map(norm => existingTotalsHeadersNorm.findIndex(h => h === norm));
            const normalizeExistingTotalsRow = (row) => {
                return totalsHeaders.map((_, targetIdx) => {
                    const srcIdx = totalsIdxMap[targetIdx];
                    if (srcIdx < 0) {
                        if (targetIdx <= 2) return "";
                        return 0;
                    }
                    return row[srcIdx];
                });
            };

            const totalsPeriodMap = new Map();
            if (existingTotalsValues.length > 1) {
                const existingRows = existingTotalsValues.slice(1);
                existingRows.forEach((row) => {
                    const normalizedRow = normalizeExistingTotalsRow(row);
                    const rawDate = normalizedRow[0];
                    const parsedDate = parsePeriodDate(rawDate);
                    const periodKey = parsedDate ? formatDateFromDate(parsedDate) : String(rawDate || "").trim();
                    if (!periodKey) return;

                    if (!totalsPeriodMap.has(periodKey)) {
                        totalsPeriodMap.set(periodKey, []);
                    }
                    totalsPeriodMap.get(periodKey).push(normalizedRow);
                });
            }

            if (currentPeriodDate && totalsPeriodMap.has(currentPeriodDate)) {
                console.log(`[Archive] Removing existing totals for period: ${currentPeriodDate} (idempotent replace)`);
                totalsPeriodMap.delete(currentPeriodDate);
            }

            totalsPeriodMap.set(currentPeriodDate || `period_${Date.now()}`, currentPeriodTotalsRows);

            const sortedTotalsPeriods = Array.from(totalsPeriodMap.keys()).sort((a, b) => {
                const dateA = new Date(a);
                const dateB = new Date(b);
                if (!isNaN(dateA) && !isNaN(dateB)) {
                    return dateB - dateA;
                }
                return b.localeCompare(a);
            });

            if (sortedTotalsPeriods.length > MAX_PERIODS) {
                const periodsToRemove = sortedTotalsPeriods.slice(MAX_PERIODS);
                console.log(`[Archive] ⚠ Dropping ${periodsToRemove.length} oldest totals period(s): ${periodsToRemove.join(", ")}`);
                periodsToRemove.forEach(key => totalsPeriodMap.delete(key));
            }

            const retainedTotalsPeriods = Array.from(totalsPeriodMap.keys()).sort((a, b) => {
                const dateA = new Date(a);
                const dateB = new Date(b);
                if (!isNaN(dateA) && !isNaN(dateB)) {
                    return dateB - dateA;
                }
                return b.localeCompare(a);
            });

            const allTotalsRows = [];
            retainedTotalsPeriods.forEach((periodKey) => {
                allTotalsRows.push(...(totalsPeriodMap.get(periodKey) || []));
            });

            if (!totalsRange.isNullObject) {
                totalsRange.clear(Excel.ClearApplyTo.contents);
            }

            const totalsHeaderRange = archiveTotalsSheet.getRangeByIndexes(0, 0, 1, totalsHeaders.length);
            totalsHeaderRange.values = [totalsHeaders];

            if (allTotalsRows.length > 0) {
                const totalsDataRange = archiveTotalsSheet.getRangeByIndexes(1, 0, allTotalsRows.length, totalsHeaders.length);
                totalsDataRange.values = allTotalsRows;
            }

            await context.sync();

            totalsHeaderRange.format.font.bold = true;
            totalsHeaderRange.format.font.size = 11;
            totalsHeaderRange.format.fill.color = "#1e293b";
            totalsHeaderRange.format.font.color = "#ffffff";
            totalsHeaderRange.format.horizontalAlignment = "Center";
            totalsHeaderRange.format.verticalAlignment = "Center";

            archiveTotalsSheet.freezePanes.freezeRows(1);

            if (allTotalsRows.length > 0) {
                const dataRowCount = allTotalsRows.length;

                archiveTotalsSheet.getRangeByIndexes(1, 0, dataRowCount, 1).numberFormat = [["mm/dd/yyyy"]];
                archiveTotalsSheet.getRangeByIndexes(1, 3, dataRowCount, 1).numberFormat = [["$#,##0"]];
                archiveTotalsSheet.getRangeByIndexes(1, 4, dataRowCount, 1).numberFormat = [["$#,##0"]];
                archiveTotalsSheet.getRangeByIndexes(1, 5, dataRowCount, 1).numberFormat = [["$#,##0"]];
                archiveTotalsSheet.getRangeByIndexes(1, 6, dataRowCount, 1).numberFormat = [["$#,##0"]];
                archiveTotalsSheet.getRangeByIndexes(1, 7, dataRowCount, 1).numberFormat = [["$#,##0"]];
                archiveTotalsSheet.getRangeByIndexes(1, 8, dataRowCount, 1).numberFormat = [["0.00%"]];
                archiveTotalsSheet.getRangeByIndexes(1, 9, dataRowCount, 1).numberFormat = [["#,##0"]];

                for (let rowIdx = 0; rowIdx < allTotalsRows.length; rowIdx++) {
                    const row = allTotalsRows[rowIdx] || [];
                    const rowType = String(row[1] || "").trim().toUpperCase();
                    const deptName = String(row[2] || "").trim().toUpperCase();
                    if (rowType === "TOTAL" || deptName === "TOTAL") {
                        const rowRange = archiveTotalsSheet.getRangeByIndexes(rowIdx + 1, 0, 1, totalsHeaders.length);
                        rowRange.format.fill.color = "#e2e8f0";
                        rowRange.format.font.bold = true;
                        rowRange.format.borders.getItem("EdgeTop").style = "Continuous";
                        rowRange.format.borders.getItem("EdgeTop").weight = "Medium";
                    }
                }
            }

            archiveTotalsSheet.getUsedRange().format.autofitColumns();
            await context.sync();

            console.log(`[Archive] ✅ Archive totals updated: ${allTotalsRows.length} rows across ${retainedTotalsPeriods.length} periods`);
        } catch (err) {
            console.warn("[Archive] ⚠ Failed to update PR_Archive_Totals:", err);
        }
        
        // Write archive data
        if (allArchiveRows.length > 0) {
            // Write headers first
            const headerRange = archiveSheet.getRangeByIndexes(0, 0, 1, archiveHeaders.length);
            headerRange.values = [archiveHeaders];
            
            // Write data rows
            const dataRange = archiveSheet.getRangeByIndexes(1, 0, allArchiveRows.length, archiveHeaders.length);
            dataRange.values = allArchiveRows;
            
            await context.sync();
            
            // ═══════════════════════════════════════════════════════════════════
            // FORMAT ARCHIVE SHEET FOR CUSTOMER READABILITY
            // ═══════════════════════════════════════════════════════════════════
            
            // Format header row
            headerRange.format.font.bold = true;
            headerRange.format.font.size = 11;
            headerRange.format.fill.color = "#1e293b";
            headerRange.format.font.color = "#ffffff";
            headerRange.format.horizontalAlignment = "Center";
            headerRange.format.verticalAlignment = "Center";
            
            // Freeze top row
            archiveSheet.freezePanes.freezeRows(1);
            
            // Find column indices for formatting
            const payDateColIdx = 0;
            const employeeColIdx = archiveHeaders.findIndex(h => {
                const hl = String(h || "").toLowerCase();
                return hl.includes("employee") && !hl.includes("amount");
            });
            
            // Find the 7 department summary column indices
            const fixedSalaryColIdx = archiveHeaders.indexOf("Fixed Salary");
            const variableSalaryColIdx = archiveHeaders.indexOf("Variable Salary");
            const grossPayColIdx = archiveHeaders.indexOf("Gross Pay");
            const burdenColIdx = archiveHeaders.indexOf("Burden");
            const allInTotalColIdx = archiveHeaders.indexOf("All-In Total");
            const percentOfTotalColIdx = archiveHeaders.indexOf("% of Total");
            const headcountColIdx = archiveHeaders.indexOf("Headcount");
            
            // Apply number formats to data rows
            for (let rowIdx = 1; rowIdx <= allArchiveRows.length; rowIdx++) {
                const row = allArchiveRows[rowIdx - 1];
                const employeeName = employeeColIdx >= 0 ? String(row[employeeColIdx] || "").trim() : "";
                const isDepartmentRow = employeeName.endsWith(" Total");
                
                // Format Pay_Date column as date
                if (payDateColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, payDateColIdx, 1, 1);
                    cell.numberFormat = [["mm/dd/yyyy"]];
                }
                
                // Format currency columns (only in department summary columns)
                if (fixedSalaryColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, fixedSalaryColIdx, 1, 1);
                    cell.numberFormat = [["$#,##0"]];
                }
                if (variableSalaryColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, variableSalaryColIdx, 1, 1);
                    cell.numberFormat = [["$#,##0"]];
                }
                if (grossPayColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, grossPayColIdx, 1, 1);
                    cell.numberFormat = [["$#,##0"]];
                }
                if (burdenColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, burdenColIdx, 1, 1);
                    cell.numberFormat = [["$#,##0"]];
                }
                if (allInTotalColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, allInTotalColIdx, 1, 1);
                    cell.numberFormat = [["$#,##0"]];
                }
                
                // Format percentage column
                if (percentOfTotalColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, percentOfTotalColIdx, 1, 1);
                    cell.numberFormat = [["0.00%"]];
                }
                
                // Format headcount column as number
                if (headcountColIdx >= 0) {
                    const cell = archiveSheet.getRangeByIndexes(rowIdx, headcountColIdx, 1, 1);
                    cell.numberFormat = [["#,##0"]];
                }
                
                // Highlight department summary rows with light background
                if (isDepartmentRow) {
                    const rowRange = archiveSheet.getRangeByIndexes(rowIdx, 0, 1, archiveHeaders.length);
                    rowRange.format.fill.color = "#f1f5f9";
                    rowRange.format.font.bold = true;
                    
                    // Special formatting for TOTAL row
                    if (employeeName === "TOTAL" || employeeName === "TOTAL Total") {
                        rowRange.format.fill.color = "#e2e8f0";
                        rowRange.format.font.color = "#1e293b";
                        rowRange.format.borders.getItem("EdgeTop").style = "Continuous";
                        rowRange.format.borders.getItem("EdgeTop").weight = "Medium";
                    }
                }
            }
            
            // Auto-fit columns
            archiveSheet.getUsedRange().format.autofitColumns();
            
            // Set minimum column widths for readability
            if (payDateColIdx >= 0) {
                archiveSheet.getRange(`${String.fromCharCode(65 + payDateColIdx)}:${String.fromCharCode(65 + payDateColIdx)}`).format.columnWidth = 85;
            }
            if (employeeColIdx >= 0) {
                archiveSheet.getRange(`${String.fromCharCode(65 + employeeColIdx)}:${String.fromCharCode(65 + employeeColIdx)}`).format.columnWidth = 200;
            }
            
            await context.sync();
            
            console.log(`[Archive] ✅ Archive summary updated and formatted: ${allArchiveRows.length} rows across ${sortedPeriods.length} periods`);
        } else {
            console.log("[Archive] No data to write to archive");
        }
        
        await context.sync();
    });
}

/**
 * Step 3: Clear working data from payroll sheets (data only, not headers)
 */
async function clearWorkingData() {
    const sheetsToClear = [
        SHEET_NAMES.DATA_CLEAN,
        SHEET_NAMES.EXPENSE_REVIEW,
        SHEET_NAMES.JE_DRAFT
    ];
    
    await Excel.run(async (context) => {
        for (const sheetName of sheetsToClear) {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            
            if (sheet.isNullObject) {
                console.log(`[Archive] Sheet not found: ${sheetName}`);
                continue;
            }
            
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount,columnCount,address");
            await context.sync();
            
            if (usedRange.isNullObject || usedRange.rowCount <= 1) {
                console.log(`[Archive] Sheet empty or headers only: ${sheetName}`);
                continue;
            }
            
            // Clear data rows (row 2 onwards), keep headers (row 1)
            const dataRange = sheet.getRange(`A2:${String.fromCharCode(64 + usedRange.columnCount)}${usedRange.rowCount}`);
            dataRange.clear(Excel.ClearApplyTo.contents);
            
            // Also clear any charts on expense review
            if (sheetName === SHEET_NAMES.EXPENSE_REVIEW) {
                const charts = sheet.charts;
                charts.load("items");
                await context.sync();
                for (let i = charts.items.length - 1; i >= 0; i--) {
                    charts.items[i].delete();
                }
            }
            
            await context.sync();
            console.log(`[Archive] Cleared data from: ${sheetName}`);
        }
    });
}

/**
 * Step 4: Clear non-permanent step notes from SS_PF_Config
 */
async function clearNonPermanentNotes() {
    await Excel.run(async (context) => {
        const table = await getConfigTable(context);
        if (!table) {
            console.warn("[Archive] Config table not found");
            return;
        }
        
        const body = table.getDataBodyRange();
        body.load("values,rowCount");
        await context.sync();
        
        const rows = body.values || [];
        let clearedCount = 0;
        
        // Find note fields to clear
        const noteFields = Object.values(STEP_NOTES_FIELDS).map(f => f.note);
        
        for (let i = 0; i < rows.length; i++) {
            const fieldName = String(rows[i][CONFIG_COLUMNS.FIELD] || "").trim();
            const permanentFlag = String(rows[i][CONFIG_COLUMNS.PERMANENT] || "").toUpperCase();
            
            // Check if this is a note field and not permanent
            if (noteFields.includes(fieldName) && permanentFlag !== "Y") {
                body.getCell(i, CONFIG_COLUMNS.VALUE).values = [[""]];
                clearedCount++;
            }
        }
        
        await context.sync();
        console.log(`[Archive] Cleared ${clearedCount} non-permanent notes`);
    });
}

/**
 * Step 5: Reset non-permanent config values (run-specific settings)
 */
async function resetNonPermanentConfig() {
    // Reset ALL fields where Permanent != "Y"
    // This clears period-specific data while preserving static configuration
    
    await Excel.run(async (context) => {
        const table = await getConfigTable(context);
        if (!table) {
            console.warn("[Archive] Config table not found");
            return;
        }
        
        const body = table.getDataBodyRange();
        body.load("values,rowCount");
        await context.sync();
        
        const rows = body.values || [];
        let resetCount = 0;
        const resetFields = [];
        const preservedFields = [];
        
        for (let i = 0; i < rows.length; i++) {
            const fieldName = String(rows[i][CONFIG_COLUMNS.FIELD] || "").trim();
            const currentValue = rows[i][CONFIG_COLUMNS.VALUE];
            const permanentFlag = String(rows[i][CONFIG_COLUMNS.PERMANENT] || "").toUpperCase().trim();
            
            // Clear any field that is NOT marked as permanent ("Y")
            // Empty/missing permanent flag = NOT permanent = should be cleared
            if (permanentFlag !== "Y") {
                // Only clear if there's actually a value
                if (currentValue !== "" && currentValue !== null && currentValue !== undefined) {
                    body.getCell(i, CONFIG_COLUMNS.VALUE).values = [[""]];
                    resetCount++;
                    resetFields.push(fieldName);
                }
            } else {
                preservedFields.push(fieldName);
            }
        }
        
        await context.sync();
        console.log(`[Archive] Reset ${resetCount} non-permanent config values:`, resetFields);
        console.log(`[Archive] Preserved ${preservedFields.length} permanent fields:`, preservedFields);
        
        // Clear local state for reset fields
        resetFields.forEach(fieldName => {
            const normalizedKey = normalizeFieldName(fieldName);
            Object.keys(configState.values).forEach(key => {
                if (normalizeFieldName(key) === normalizedKey) {
                    configState.values[key] = "";
                }
            });
        });
    });
}

async function runJournalSummary() {
    if (!hasExcelRuntime()) {
        showToast("Excel runtime is unavailable.", "error");
        return;
    }
    journalState.loading = true;
    journalState.lastError = null;
    journalState.validationRun = false;
    markJeSaveState(false);
    renderApp();
    try {
        const totals = await Excel.run(async (context) => {
            const jeSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.JE_DRAFT);
            jeSheet.load("isNullObject");
            await context.sync();
            
            if (jeSheet.isNullObject) {
                throw new Error(`${SHEET_NAMES.JE_DRAFT} sheet not found. Click "Generate" first to create the journal entry.`);
            }
            
            const jeRange = jeSheet.getUsedRangeOrNullObject();
            jeRange.load("values");
            await context.sync();
            
            const values = jeRange.isNullObject ? [] : jeRange.values || [];
            if (!values.length) {
                throw new Error(`${SHEET_NAMES.JE_DRAFT} is empty. Click "Generate" to create the journal entry.`);
            }
            const headers = (values[0] || []).map((h) => normalizeHeader(h));
            console.log("[JE-Validation] Headers:", headers);
            
            const debitIdx = headers.findIndex((h) => h.includes("debit"));
            const creditIdx = headers.findIndex((h) => h.includes("credit"));
            const acctNameIdx = headers.findIndex((h) => h.includes("account") && h.includes("name"));
            console.log("[JE-Validation] Column indexes: debit=", debitIdx, "credit=", creditIdx, "acctName=", acctNameIdx);
            
            if (debitIdx === -1 || creditIdx === -1) {
                throw new Error("Debit/Credit columns not found in JE Draft.");
            }
            let debitTotal = 0;
            let creditTotal = 0;
            let jeExpenseTotal = 0; // Sum of expense line amounts (not clearing account offset)
            
            values.slice(1).forEach((row, rowIdx) => {
                const debit = Number(row[debitIdx]) || 0;
                const credit = Number(row[creditIdx]) || 0;
                const acctName = acctNameIdx !== -1 ? String(row[acctNameIdx] || "").trim().toLowerCase() : "";
                
                debitTotal += debit;
                creditTotal += credit;
                
                // Sum expense lines only (not the offset account)
                // Offset account is "Uncategorized Expense" (or legacy "Payroll Clearing Account")
                const isOffsetLine = acctName.includes("uncategorized") || acctName.includes("clearing");
                if (acctName && !isOffsetLine) {
                    jeExpenseTotal += debit - credit;  // Net = debit minus credit per line
                }
                
                // Debug first few rows
                if (rowIdx < 3) {
                    console.log(`[JE-Validation] Row ${rowIdx}: debit=${debit}, credit=${credit}, acct="${acctName}", isOffset=${isOffsetLine}`);
                }
            });
            
            console.log("[JE-Validation] Totals: debit=", debitTotal, "credit=", creditTotal, "jeExpense=", jeExpenseTotal);
            
            // Also read PR_Data_Clean total for comparison (dynamic, schema-driven)
            // IMPORTANT: Use the existing measure universe as the single source of truth.
            const universe = await getPRDataCleanMeasureUniverse();
            const cleanTotal = Number.isFinite(Number(universe.total)) ? universe.total : 0;
            const cleanTotalsInfo = {
                includedColumns: Array.isArray(universe.includedMeasureHeaders) ? universe.includedMeasureHeaders : [],
                excludedColumns: Array.isArray(universe.excludedHeaders) ? universe.excludedHeaders : []
            };
            
            // Build validation issues array (consistent with PTO module)
            const difference = debitTotal - creditTotal;
            const issues = [];

            // Check 0: Payroll date configured (non-blocking)
            const payrollDateFieldValue = getConfigValue("PR_Payroll_Date");
            if (!String(payrollDateFieldValue || "").trim()) {
                issues.push({
                    check: "Payroll Date Configured",
                    passed: false,
                    detail: "Payroll date missing in SS_Config (PR_Payroll_Date)."
                });
            } else {
                issues.push({ check: "Payroll Date Configured", passed: true, detail: "" });
            }
            
            // Check 1: Debits = Credits (core JE balance check)
            if (Math.abs(difference) >= 0.01) {
                issues.push({
                    check: "Debits = Credits",
                    passed: false,
                    detail: difference > 0 
                        ? `Debits exceed credits by ${formatNumberDisplay(Math.abs(difference))}`
                        : `Credits exceed debits by ${formatNumberDisplay(Math.abs(difference))}`
                });
            } else {
                issues.push({ check: "Debits = Credits", passed: true, detail: "" });
            }
            
            // Check 2: JE Matches Source Total (expense lines should match PR_Data_Clean)
            const sourceDiff = Math.abs(jeExpenseTotal - cleanTotal);
            if (sourceDiff >= 0.01) {
                issues.push({
                    check: "JE Matches Source Total",
                    passed: false,
                    detail: `PR_Data_Clean total ${formatNumberDisplay(cleanTotal)} vs JE expense total ${formatNumberDisplay(jeExpenseTotal)} (Δ ${formatNumberDisplay(sourceDiff)})`
                });
            } else {
                issues.push({ check: "JE Matches Source Total", passed: true, detail: "" });
            }
            
            // Check 3: Unmapped Columns (warning if any)
            const unmappedCount = journalState.unmappedColumns?.length || 0;
            const unmappedTotal = journalState.unmappedTotal || 0;
            if (unmappedCount > 0) {
                issues.push({
                    check: "All Columns Mapped",
                    passed: false,
                    detail: `${unmappedCount} column+dept combinations (${formatNumberDisplay(Math.abs(unmappedTotal))}) need GL mappings`
                });
            } else {
                issues.push({ check: "All Columns Mapped", passed: true, detail: "" });
            }
            
            return { 
                debitTotal, 
                creditTotal, 
                difference,
                cleanTotal,
                cleanTotalsInfo,
                issues,
                validationRun: true
            };
        });
        Object.assign(journalState, totals, { lastError: null });
    } catch (error) {
        console.warn("JE summary:", error);
        journalState.lastError = error?.message || "Unable to calculate journal totals.";
        journalState.debitTotal = null;
        journalState.creditTotal = null;
        journalState.difference = null;
        journalState.cleanTotal = null;
        journalState.validationRun = false;
        journalState.issues = [];
    } finally {
        journalState.loading = false;
        renderApp();
    }
}

async function saveJournalSummary() {
    try {
        const debit = Number.isFinite(Number(journalState.debitTotal)) ? journalState.debitTotal : "";
        const credit = Number.isFinite(Number(journalState.creditTotal)) ? journalState.creditTotal : "";
        const diff = Number.isFinite(Number(journalState.difference)) ? journalState.difference : "";
        await Promise.all([
            writeConfigValue(JE_TOTAL_DEBIT_FIELD, String(debit)),
            writeConfigValue(JE_TOTAL_CREDIT_FIELD, String(credit)),
            writeConfigValue(JE_DIFFERENCE_FIELD, String(diff))
        ]);
        markJeSaveState(true);
    } catch (error) {
        console.error("JE save:", error);
    }
}

// =============================================================================
// OLD JE GENERATION REMOVED - Now using createJournalEntryDraftV2() 
// See "JOURNAL ENTRY GENERATION V2" section at end of file
// =============================================================================

async function exportJournalDraft() {
    if (!hasExcelRuntime()) {
        showToast("Excel runtime is unavailable.", "error");
        return;
    }
    try {
        const { rows } = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.JE_DRAFT);
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            const values = range.isNullObject ? [] : range.values || [];
            if (!values.length) {
                throw new Error(`${SHEET_NAMES.JE_DRAFT} is empty.`);
            }
            return { rows: values };
        });

        // Ensure Debit/Credit export is QBO-ready: 2 decimals, no currency symbols/commas
        const headers = (rows[0] || []).map((h) => String(h || "").trim().toLowerCase());
        const debitIdx = headers.findIndex((h) => h.includes("debit"));
        const creditIdx = headers.findIndex((h) => h.includes("credit"));
        const dateIdx = headers.findIndex((h) => h.includes("journaldate") || h.includes("txndate") || h === "date");
        
        const normalizedRows = rows.map((row, idx) => {
            if (idx === 0) return row;
            const next = [...(row || [])];
            const normalizeAmount = (v) => {
                if (v === null || v === "") return "";
                const n = Number(v);
                if (!Number.isFinite(n)) return "";
                return n.toFixed(2);
            };
            // Normalize date column - Excel may return serial numbers
            const normalizeDate = (v) => {
                if (v === null || v === "") return "";
                // If it's already a string in MM/DD/YYYY format, return as-is
                if (typeof v === "string" && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v.trim())) {
                    return v.trim();
                }
                // Handle Excel serial number
                if (typeof v === "number" && Number.isFinite(v)) {
                    const ms = Math.round((v - 25569) * 86400 * 1000);
                    const d = new Date(ms);
                    if (Number.isFinite(d.getTime())) {
                        const mm = String(d.getMonth() + 1).padStart(2, "0");
                        const dd = String(d.getDate()).padStart(2, "0");
                        const yyyy = d.getFullYear();
                        if (yyyy >= 1900 && yyyy <= 2100) {
                            return `${mm}/${dd}/${yyyy}`;
                        }
                    }
                }
                // Handle Date object
                if (v instanceof Date && Number.isFinite(v.getTime())) {
                    const mm = String(v.getMonth() + 1).padStart(2, "0");
                    const dd = String(v.getDate()).padStart(2, "0");
                    const yyyy = v.getFullYear();
                    return `${mm}/${dd}/${yyyy}`;
                }
                return String(v);
            };
            if (debitIdx >= 0) next[debitIdx] = normalizeAmount(next[debitIdx]);
            if (creditIdx >= 0) next[creditIdx] = normalizeAmount(next[creditIdx]);
            if (dateIdx >= 0) next[dateIdx] = normalizeDate(next[dateIdx]);
            return next;
        });

        const csv = buildCsv(normalizedRows);
        downloadCsv(`pr-je-draft-${todayIso()}.csv`, csv);
    } catch (error) {
        console.warn("JE export:", error);
        showToast("Unable to export the JE draft. Confirm the sheet has data.", "error");
    }
}

async function readConfigValueFromWorkbook(fieldName) {
    if (!hasExcelRuntime()) return getConfigValue(fieldName);
    try {
        const { value } = await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            if (!table) return { value: "" };
            const body = table.getDataBodyRange();
            body.load("values");
            await context.sync();
            const rows = body.values || [];
            const target = normalizeFieldName(fieldName);
            for (const row of rows) {
                const rowField = normalizeFieldName(row[CONFIG_COLUMNS.FIELD]);
                if (rowField === target) {
                    return { value: row[CONFIG_COLUMNS.VALUE] ?? "" };
                }
            }
            return { value: "" };
        });
        return value;
    } catch (error) {
        console.warn("readConfigValueFromWorkbook:", error);
        return getConfigValue(fieldName);
    }
}

function formatJournalDateForQbo(raw) {
    if (raw === null || raw === undefined) return { value: "", valid: false };
    if (raw === "") return { value: "", valid: false };

    // Excel serial number
    if (typeof raw === "number" && Number.isFinite(raw)) {
        // Excel serial 1 = 1900-01-01 (with 1900 leap year bug)
        // Use UTC to avoid timezone shift
        const ms = Math.round((raw - 25569) * 86400 * 1000);
        const d = new Date(ms);
        if (!Number.isFinite(d.getTime())) return { value: "", valid: false };
        
        // Use UTC methods to avoid timezone shift
        const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
        const dd = String(d.getUTCDate()).padStart(2, "0");
        const yyyy = d.getUTCFullYear();
        
        if (yyyy < 1900 || yyyy > 2100) return { value: "", valid: false };
        return { value: `${mm}/${dd}/${yyyy}`, valid: true };
    }

    if (raw instanceof Date) {
        const d = raw;
        if (!Number.isFinite(d.getTime())) return { value: "", valid: false };
        
        // Use UTC methods to avoid timezone shift
        const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
        const dd = String(d.getUTCDate()).padStart(2, "0");
        const yyyy = d.getUTCFullYear();
        
        if (yyyy < 1900 || yyyy > 2100) return { value: "", valid: false };
        return { value: `${mm}/${dd}/${yyyy}`, valid: true };
    }

    const s = String(raw).trim();
    if (!s) return { value: "", valid: false };

    // Prefer the existing strict parser if possible
    const parts = parseDateInput(s);
    if (parts && parts.year >= 1900 && parts.year <= 2100) {
        const mm = String(parts.month).padStart(2, "0");
        const dd = String(parts.day).padStart(2, "0");
        return { value: `${mm}/${dd}/${parts.year}`, valid: true };
    }

    // Fallback: attempt Date parse but reject epoch/invalid
    const d = new Date(s);
    if (!Number.isFinite(d.getTime())) return { value: "", valid: false };
    
    // Use UTC methods to avoid timezone shift
    const yyyy = d.getUTCFullYear();
    if (yyyy < 1900 || yyyy > 2100) return { value: "", valid: false };
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const dd = String(d.getUTCDate()).padStart(2, "0");
    return { value: `${mm}/${dd}/${yyyy}`, valid: true };
}

/**
 * Open the accounting software URL from SS_PF_Config
 * Reads SS_Accounting_Software field and opens in new window
 */
async function openAccountingSoftware() {
    // First check if we have the URL in cached config
    let accountingUrl = getConfigValue("SS_Accounting_Software");
    
    // If not in cache, try reading from Excel
    if (!accountingUrl && hasExcelRuntime()) {
        try {
            await Excel.run(async (context) => {
                const table = await getConfigTable(context);
                if (!table) return;
                
                const body = table.getDataBodyRange();
                body.load("values");
                await context.sync();
                
                const rows = body.values || [];
                for (const row of rows) {
                    const fieldName = String(row[CONFIG_COLUMNS.FIELD] || "").trim();
                    if (fieldName === "SS_Accounting_Software" || 
                        normalizeFieldName(fieldName) === normalizeFieldName("SS_Accounting_Software")) {
                        accountingUrl = String(row[CONFIG_COLUMNS.VALUE] || "").trim();
                        break;
                    }
                }
            });
        } catch (error) {
            console.warn("Error reading accounting software URL:", error);
        }
    }
    
    if (!accountingUrl) {
        showToast("No accounting software URL configured. Add SS_Accounting_Software to SS_PF_Config.", "info", 5000);
        return;
    }
    
    // Ensure URL has protocol
    if (!accountingUrl.startsWith("http://") && !accountingUrl.startsWith("https://")) {
        accountingUrl = "https://" + accountingUrl;
    }
    
    // Open in new window/tab
    window.open(accountingUrl, "_blank");
    showToast("Opening accounting software...", "success", 2000);
}

// =============================================================================
// JOURNAL ENTRY GENERATION V2 - CLEAN REBUILD
// =============================================================================
// Pipeline: PR_Data_Clean → ada_customer_gl_mappings → PR_JE_Draft
// Output: Expense allocation file for QuickBooks (PEO customer)
// =============================================================================

/**
 * Normalize key for matching (MANDATORY for all lookups)
 * - Trim, lowercase, collapse spaces, replace & with 'and'
 */
function jeNormalizeKey(value) {
    return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/\u00A0/g, " ")      // non-breaking space
        .replace(/\s+/g, " ")          // collapse multiple spaces
        .replace(/&/g, "and");         // ampersand to 'and'
}

/**
 * Fetch GL mappings from ada_customer_gl_mappings
 * Returns Map<normalizedKey, {gl_account, gl_account_name}>
 * Key format: "pf_column_name|department"
 */
async function jeLoadGLMappings(companyId) {
    const SUPABASE_URL = "https://jgciqwzwacaesqjaoadc.supabase.co";
    const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE";
    
    const url = `${SUPABASE_URL}/rest/v1/ada_customer_gl_mappings?company_id=eq.${encodeURIComponent(companyId)}&module=eq.payroll-recorder&select=pf_column_name,gl_account,gl_account_name,department`;
    
    console.log("[JE-V2] Fetching GL mappings:", url);
    
    const response = await fetch(url, {
        headers: {
            "apikey": SUPABASE_KEY,
            "Authorization": `Bearer ${SUPABASE_KEY}`,
            "Content-Type": "application/json"
        }
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`GL mappings fetch failed: ${response.status} - ${errorText}`);
    }
    
    const rows = await response.json();
    console.log(`[JE-V2] Fetched ${rows.length} GL mapping rows`);
    
    // Build lookup map: "pf_column_name|department" → {gl_account, gl_account_name}
    const mappings = new Map();
    
    for (const row of rows) {
        const pfKey = jeNormalizeKey(row.pf_column_name);
        const deptKey = jeNormalizeKey(row.department || "");
        const lookupKey = `${pfKey}|${deptKey}`;
        
        if (!mappings.has(lookupKey)) {
            mappings.set(lookupKey, {
                gl_account: row.gl_account,
                gl_account_name: row.gl_account_name || ""
            });
        }
    }
    
    console.log(`[JE-V2] Built ${mappings.size} unique GL lookup keys`);
    return mappings;
}

/**
 * Create Journal Entry Draft - CLEAN REBUILD
 * Reads PR_Data_Clean, aggregates by (pf_column_name, department), writes PR_JE_Draft
 */
async function createJournalEntryDraftV2() {
    console.log("[JE-V2] ═══════════════════════════════════════════════════════════════");
    console.log("[JE-V2] Starting Journal Entry Generation (QuickBooks Format)");
    console.log("[JE-V2] ═══════════════════════════════════════════════════════════════");
    
    if (typeof Excel === "undefined") {
        console.error("[JE-V2] Excel runtime not available");
        showToast("Excel runtime not available.", "error");
        return;
    }
    
    console.log("[JE-V2] Excel runtime available, proceeding...");
    showToast("Generating Journal Entry...", "info", 10000);
    
    try {
        // =====================================================================
        // 1. GET COMPANY ID AND JOURNAL CONFIG FROM SS_PF_Config
        // =====================================================================
        const companyId = getConfigValue("SS_Company_ID");
        const journalNo = getConfigValue("PR_Journal_Entry_ID") || "";
        const journalDateRaw = await readConfigValueFromWorkbook("PR_Payroll_Date");
        
        console.log(`[JE-V2] company_id: "${companyId}"`);
        console.log(`[JE-V2] JournalNo (PR_Journal_Entry_ID): "${journalNo}"`);
        console.log(`[JE-V2] JournalDate (PR_Payroll_Date): "${journalDateRaw}"`);
        
        if (!companyId) {
            throw new Error("SS_Company_ID not set in SS_PF_Config. Add a row with Field='SS_Company_ID' and your company UUID.");
        }
        
        if (!journalNo) {
            throw new Error("PR_Journal_Entry_ID not set. Please enter a Journal Entry ID in the Configuration step.");
        }
        
        // Format date for QuickBooks (MM/DD/YYYY)
        // If missing/invalid: write blank (non-blocking) and surface via validation UI.
        const formatted = formatJournalDateForQbo(journalDateRaw);
        const formattedDate = formatted.valid ? formatted.value : "";
        
        // =====================================================================
        // 2. LOAD GL MAPPINGS FROM DATABASE
        // =====================================================================
        const glMappings = await jeLoadGLMappings(companyId);
        console.log(`[JE-V2] GL mappings loaded: ${glMappings.size}`);
        
        if (glMappings.size === 0) {
            throw new Error(`No GL mappings found for company_id="${companyId}" and module="payroll-recorder". Populate ada_customer_gl_mappings first.`);
        }
        
        // =====================================================================
        // 3. READ PR_DATA_CLEAN, CHART OF ACCOUNTS, AND AGGREGATE
        // =====================================================================
        let aggregatedLines = [];
        let chartOfAccountsLookup = new Map();  // Account Number → Account Name
        
        await Excel.run(async (context) => {
            // =====================================================================
            // 3a. READ CHART OF ACCOUNTS FOR ACCOUNT NAME LOOKUP
            // =====================================================================
            try {
                const coaSheet = context.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");
                coaSheet.load("isNullObject");
                await context.sync();
                
                if (!coaSheet.isNullObject) {
                    const coaUsedRange = coaSheet.getUsedRangeOrNullObject();
                    coaUsedRange.load("isNullObject");
                    await context.sync();
                    
                    if (!coaUsedRange.isNullObject) {
                        coaUsedRange.load("values");
                        await context.sync();
                        
                        if (coaUsedRange.values && coaUsedRange.values.length > 1) {
                            const coaHeaders = coaUsedRange.values[0];
                            const coaRows = coaUsedRange.values.slice(1);
                            
                            console.log(`[JE-V2] COA Headers found: ${JSON.stringify(coaHeaders)}`);
                            
                            // Find account number column - be more flexible with header names
                            const acctNumIdx = coaHeaders.findIndex(h => {
                                const normalized = jeNormalizeKey(h);
                                return normalized === "account_number" || 
                                       normalized === "accountnumber" || 
                                       normalized === "account number" ||
                                       normalized === "account" ||
                                       normalized === "acct" ||
                                       normalized === "acctnum" ||
                                       normalized === "acct_num" ||
                                       normalized === "number" ||
                                       normalized === "gl_account" ||
                                       (normalized.includes("account") && normalized.includes("num"));
                            });
                            
                            // Find account name column - be more flexible with header names
                            const acctNameIdx = coaHeaders.findIndex(h => {
                                const normalized = jeNormalizeKey(h);
                                return normalized === "account_name" || 
                                       normalized === "accountname" ||
                                       normalized === "account name" ||
                                       normalized === "name" ||
                                       normalized === "acct_name" ||
                                       normalized === "acctname" ||
                                       normalized === "description" ||
                                       normalized === "account_description" ||
                                       (normalized.includes("account") && normalized.includes("name"));
                            });
                            
                            console.log(`[JE-V2] Account Number column index: ${acctNumIdx}`);
                            console.log(`[JE-V2] Account Name column index: ${acctNameIdx}`);
                            
                            if (acctNumIdx >= 0 && acctNameIdx >= 0) {
                                for (const row of coaRows) {
                                    const acctNumber = String(row[acctNumIdx] || "").trim();
                                    const acctName = String(row[acctNameIdx] || "").trim();
                                    if (acctNumber) {
                                        chartOfAccountsLookup.set(acctNumber, acctName);
                                    }
                                }
                                console.log(`[JE-V2] Chart of Accounts lookup built: ${chartOfAccountsLookup.size} accounts`);
                                // Log sample entries for debugging
                                console.log(`[JE-V2] Sample COA entries:`, Array.from(chartOfAccountsLookup.entries()).slice(0, 5));
                            } else {
                                console.warn(`[JE-V2] Could not find COA columns: acctNumIdx=${acctNumIdx}, acctNameIdx=${acctNameIdx}`);
                            }
                        }
                    }
                }
            } catch (coaError) {
                console.warn("[JE-V2] Error reading SS_Chart_of_Accounts (non-fatal):", coaError.message);
            }
            
            // =====================================================================
            // 3b. READ PR_DATA_CLEAN
            // =====================================================================
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject("PR_Data_Clean");
            cleanSheet.load("isNullObject");
            await context.sync();
            
            if (cleanSheet.isNullObject) {
                throw new Error("PR_Data_Clean sheet not found. Run Create Matrix first.");
            }
            
            const usedRange = cleanSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (usedRange.isNullObject) {
                throw new Error("PR_Data_Clean is empty.");
            }
            
            const allRows = usedRange.values || [];
            if (allRows.length < 2) {
                throw new Error("PR_Data_Clean has no data rows.");
            }
            
            const headers = allRows[0];
            const dataRows = allRows.slice(1);
            
            console.log(`[JE-V2] PR_Data_Clean: ${headers.length} columns, ${dataRows.length} rows`);
            
            // Identify Department column
            const deptIndex = headers.findIndex(h => 
                jeNormalizeKey(h) === "department_name" || jeNormalizeKey(h) === "department"
            );
            
            // Identify numeric columns (skip known dimension columns)
            const DIMENSION_COLUMNS = new Set([
                "pay_date", "pay_period_start", "pay_period_end",
                "employee_name", "employee_id", "department", "department_code", 
                "department_name", "location", "cost_center", "job_title",
                "check_number", "pay_type", "pay_frequency"
            ]);
            
            const numericColumns = [];
            for (let i = 0; i < headers.length; i++) {
                const headerNorm = jeNormalizeKey(headers[i]);
                if (DIMENSION_COLUMNS.has(headerNorm)) continue;
                
                const hasNumeric = dataRows.some(row => {
                    const val = row[i];
                    return val !== null && val !== "" && !isNaN(Number(val));
                });
                
                if (hasNumeric) {
                    numericColumns.push({ index: i, header: headers[i] });
                }
            }
            
            console.log(`[JE-V2] Numeric columns found: ${numericColumns.length}`);
            
            // Aggregate: Map<"pf_column_name|department", {amount, signedAmount}>
            const aggregation = new Map();
            
            for (const row of dataRows) {
                const department = deptIndex >= 0 ? String(row[deptIndex] || "") : "";
                
                for (const col of numericColumns) {
                    const rawValue = row[col.index];
                    const amount = Number(rawValue);
                    
                    if (isNaN(amount) || amount === 0) continue;
                    
                    const pfKey = jeNormalizeKey(col.header);
                    const deptKey = jeNormalizeKey(department);
                    const aggKey = `${pfKey}|${deptKey}`;
                    
                    const existing = aggregation.get(aggKey) || { 
                        pf_column_name: col.header, 
                        department: department, 
                        amount: 0  // Net signed amount
                    };
                    existing.amount += amount;  // Keep sign for debit/credit determination
                    aggregation.set(aggKey, existing);
                }
            }
            
            console.log(`[JE-V2] Aggregated lines: ${aggregation.size}`);
            
            // =====================================================================
            // 4. RESOLVE GL ACCOUNTS AND BUILD JE LINES
            // =====================================================================
            const jeLines = [];
            
            for (const [aggKey, data] of aggregation) {
                if (Math.abs(data.amount) < 0.01) continue;  // Skip zero amounts
                
                const pfNorm = jeNormalizeKey(data.pf_column_name);
                const deptNorm = jeNormalizeKey(data.department);
                
                // Try exact match first, then fallback without department
                let lookupKey = `${pfNorm}|${deptNorm}`;
                let glEntry = glMappings.get(lookupKey);
                
                if (!glEntry) {
                    lookupKey = `${pfNorm}|`;
                    glEntry = glMappings.get(lookupKey);
                }
                
                if (!glEntry) {
                    console.error(`[JE] MISSING GL MAPPING:`, {
                        pf_column_name: data.pf_column_name,
                        department: data.department,
                        hint: "Check ada_customer_gl_mappings has this pf_column_name"
                    });
                    throw new Error(
                        `Missing GL mapping for: pf_column_name="${data.pf_column_name}", ` +
                        `department="${data.department}". ` +
                        `Verify this pf_column_name exists in ada_payroll_column_dictionary.`
                    );
                }
                
                // Look up account name from SS_Chart_of_Accounts (priority over GL mapping)
                const accountName = chartOfAccountsLookup.get(glEntry.gl_account) || glEntry.gl_account_name || "";
                
                jeLines.push({
                    pf_column_name: data.pf_column_name,
                    department: data.department,
                    amount: data.amount,  // Signed amount for debit/credit determination
                    gl_account: glEntry.gl_account,
                    gl_account_name: accountName
                });
            }
            
            console.log(`[JE-V2] JE lines with GL accounts: ${jeLines.length}`);
            aggregatedLines = jeLines;
        });
        
        // =====================================================================
        // 5. BUILD QUICKBOOKS-FORMAT OUTPUT
        // =====================================================================
        if (aggregatedLines.length === 0) {
            throw new Error("No JE lines generated. Check PR_Data_Clean has numeric data.");
        }
        
        // Calculate totals for offset entry
        let totalDebits = 0;
        let totalCredits = 0;
        
        for (const line of aggregatedLines) {
            if (line.amount > 0) {
                totalDebits += line.amount;
            } else {
                totalCredits += Math.abs(line.amount);
            }
        }
        
        const offsetAmount = totalDebits - totalCredits;
        
        console.log("[JE-V2] ═══════════════════════════════════════════════════════════════");
        console.log("[JE-V2] QUICKBOOKS JE SUMMARY");
        console.log(`[JE-V2]   JournalNo: ${journalNo}`);
        console.log(`[JE-V2]   JournalDate: ${formattedDate}`);
        console.log(`[JE-V2]   Total Debits: ${totalDebits.toFixed(2)}`);
        console.log(`[JE-V2]   Total Credits: ${totalCredits.toFixed(2)}`);
        console.log(`[JE-V2]   Offset to Uncategorized Expense: ${offsetAmount.toFixed(2)}`);
        console.log("[JE-V2] ═══════════════════════════════════════════════════════════════");
        
        await Excel.run(async (context) => {
            // Get or create PR_JE_Draft sheet
            let draftSheet = context.workbook.worksheets.getItemOrNullObject("PR_JE_Draft");
            draftSheet.load("isNullObject");
            await context.sync();
            
            if (draftSheet.isNullObject) {
                draftSheet = context.workbook.worksheets.add("PR_JE_Draft");
                console.log("[JE-V2] Created PR_JE_Draft sheet");
            } else {
                // Make sure sheet is visible (it may be hidden by tab visibility)
                draftSheet.visibility = Excel.SheetVisibility.visible;
                const usedRange = draftSheet.getUsedRangeOrNullObject();
                usedRange.load("isNullObject");
                await context.sync();
                if (!usedRange.isNullObject) {
                    usedRange.clear();
                }
                console.log("[JE-V2] Cleared existing PR_JE_Draft");
            }
            
            // QuickBooks Headers (EXACT format requested)
            const headers = [
                "JournalNo",
                "JournalDate",
                "Account Name",
                "Debits",
                "Credits",
                "Description"
            ];
            
            // Build output rows
            const outputRows = [headers];
            
            for (const line of aggregatedLines) {
                const isDebit = line.amount > 0;
                const absAmount = Math.abs(line.amount);
                const description = `${line.department}${line.department ? " - " : ""}${line.pf_column_name}`;
                
                // Look up account name from COA and format as "AccountNumber AccountName" for QuickBooks
                const glAccountStr = String(line.gl_account).trim();
                const accountNameFromCOA = chartOfAccountsLookup.get(glAccountStr);
                
                // QuickBooks format: "52160 Support PEO:Support Onshore Labor:Support 401k Employer Contribution"
                let accountName;
                if (accountNameFromCOA) {
                    accountName = `${glAccountStr} ${accountNameFromCOA}`;
                    console.log(`[JE-V2] GL "${glAccountStr}" → "${accountName}"`);
                } else if (line.gl_account_name) {
                    // Fallback: use gl_account_name with account number prefix
                    accountName = `${glAccountStr} ${line.gl_account_name}`;
                    console.log(`[JE-V2] GL "${glAccountStr}" not in COA, using fallback: "${accountName}"`);
                } else {
                    // Last resort: just the account number
                    accountName = glAccountStr;
                    console.log(`[JE-V2] GL "${glAccountStr}" - no name found, using number only`);
                }
                
                outputRows.push([
                    journalNo,                              // JournalNo
                    formattedDate,                          // JournalDate
                    accountName,                            // Account Name (AccountNumber + AccountName)
                    isDebit ? absAmount : "",               // Debits (amount > 0)
                    isDebit ? "" : absAmount,               // Credits (ABS of amount < 0)
                    description                             // Description = Department + PF_column_name
                ]);
            }
            
            // Add Uncategorized Expense offset line to balance the JE
            if (Math.abs(offsetAmount) >= 0.01) {
                outputRows.push([
                    journalNo,
                    formattedDate,
                    "Uncategorized Expense",
                    offsetAmount < 0 ? Math.abs(offsetAmount) : "",  // Debit if offset is negative
                    offsetAmount > 0 ? offsetAmount : "",            // Credit if offset is positive
                    "Payroll Offset"
                ]);
            }
            
            // Write to sheet
            const outputRange = draftSheet.getRangeByIndexes(0, 0, outputRows.length, headers.length);
            outputRange.values = outputRows;
            
            // Format headers
            const headerRange = draftSheet.getRangeByIndexes(0, 0, 1, headers.length);
            formatSheetHeaders(headerRange);
            
            // Format numeric columns (Debits, Credits) for QBO: 2 decimals, no currency symbol, no commas
            const currencyFormat = "0.00";
            const dataRowCount = outputRows.length - 1;
            if (dataRowCount > 0) {
                draftSheet.getRangeByIndexes(1, 3, dataRowCount, 1).numberFormat = [[currencyFormat]];
                draftSheet.getRangeByIndexes(1, 4, dataRowCount, 1).numberFormat = [[currencyFormat]];
            }
            
            // Auto-fit columns
            draftSheet.getUsedRange().format.autofitColumns();
            
            // Freeze header row
            draftSheet.freezePanes.freezeRows(1);
            
            // Activate the sheet
            draftSheet.activate();
            
            await context.sync();
            console.log(`[JE-V2] Wrote ${outputRows.length - 1} rows to PR_JE_Draft (QuickBooks format)`);
        });
        
        // Update journal state for validation UI
        journalState.debitTotal = totalDebits + (offsetAmount < 0 ? Math.abs(offsetAmount) : 0);
        journalState.creditTotal = totalCredits + (offsetAmount > 0 ? offsetAmount : 0);
        journalState.difference = journalState.debitTotal - journalState.creditTotal;
        journalState.lastGenerated = new Date().toISOString();
        
        // Run validation checks after generation
        await runJournalSummary();
        
        // Refresh validation display
        renderApp();
        
        showToast(`Journal Entry created! ${aggregatedLines.length + 1} lines (including offset), balanced`, "success", 5000);
        
    } catch (error) {
        console.error("[JE-V2] ERROR:", error);
        showToast(`JE Error: ${error.message}`, "error", 10000);
    }
}
