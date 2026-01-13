# Installation Key Implementation - Discovery Answers

**Project**: Customer-Tai-Software (TaiTools)  
**Project ID**: jgciqwzwacaesqjaoadc  
**Current Installation Key**: `pf_install_9f3c2b1a_20251212`  
**Date**: January 12, 2026

---

## Section 1: Current Installation Key Usage

### Answer 1.1: Bootstrap Function Location

**File Path**: `/Users/d.paeth/Customer-Tai-Software/Common/bootstrap.js`

**Complete Function Code**:
```javascript
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
```

**Called From**:
- **File**: `/Users/d.paeth/Customer-Tai-Software/module-selector/selector.js`
- **Line**: 524
- **Context**: Called during add-in initialization, AFTER ensuring SS_PF_Config exists
```javascript
// Line 522-533 in selector.js
console.log("[Init] Running bootstrap config sync...");
try {
    const bootstrapResult = await bootstrapConfigSync();
    console.log("[Init] Bootstrap result:", bootstrapResult);
    if (!bootstrapResult.success) {
        console.warn("[Init] Bootstrap failed but continuing:", bootstrapResult.error);
    }
} catch (bootstrapError) {
    console.error("[Bootstrap error (non-fatal):", bootstrapError);
    // Continue anyway - manual config still works
}
```

---

### Answer 1.2: Installation Key Storage

**Current Storage Method**: **Hardcoded constant**

**Location**: `/Users/d.paeth/Customer-Tai-Software/Common/bootstrap.js`, Line 25

```javascript
const INSTALLATION_KEY = "pf_install_9f3c2b1a_20251212";
```

**How It's Used**:
1. **Passed to Edge Function** (Line 132 in bootstrap.js):
```javascript
const result = await columnMapperRequest("bootstrap", {
    installation_key: INSTALLATION_KEY
}, "bootstrap_fetch");
```

2. **Added to Response Data** (Line 146 in bootstrap.js):
```javascript
// Add installation_key to data so it gets written to SS_PF_Config
data.installation_key = INSTALLATION_KEY;
```

3. **Written to Excel** via `writeConfigValues()` function which maps it to `SS_Installation_Key` field in SS_PF_Config table

**Field Mapping** (Lines 28-35 in bootstrap.js):
```javascript
const FIELD_MAP = {
    "installation_key": "SS_Installation_Key",
    "company_id": "SS_Company_ID",
    "ss_company_name": "SS_Company_Name",
    "ss_accounting_software": "SS_Accounting_Software",
    "pto_payroll_provider": "PTO_Payroll_Provider",
    "pr_payroll_provider": "PR_Payroll_Provider"
};
```

**NOT stored in**:
- ❌ Environment variables
- ❌ Config files (.json, .env)
- ❌ Excel settings/custom properties (only written to SS_PF_Config table)
- ❌ OfficeRuntime.storage

---

### Answer 1.3: Edge Function Integration

**API Call Implementation**: Uses centralized `warehouseRequest()` wrapper

**File**: `/Users/d.paeth/Customer-Tai-Software/Common/warehouse.js`

**Edge Function URL**: 
```javascript
const WAREHOUSE_BASE_URL = "https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1";
// Full endpoint: https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1/column-mapper
```

**Request Function** (Lines 145-198 in warehouse.js):
```javascript
export async function warehouseRequest(endpoint, body, callsiteTag = "unknown") {
    const url = `${WAREHOUSE_BASE_URL}/${endpoint}`;
    const ctx = getWarehouseAuthContext();
    
    // Log request (no secrets)
    console.log(`[DW] callsite=${callsiteTag} url=${url} tokenSource=${ctx.tokenSource}`);
    
    // PREFLIGHT INVARIANT: Auth must be valid
    if (!ctx.isValid) {
        const errorMsg = `[DW:PREFLIGHT] ❌ Authorization missing!`;
        console.error(errorMsg);
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
            return { ok: false, status: 401, error: `Unauthorized: ${errorBody}` };
        }
        
        if (!response.ok) {
            const errorBody = await response.text();
            console.error(`[DW:ERROR] callsite=${callsiteTag} status=${response.status}`);
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
```

**Convenience Wrapper** (Lines 207-209):
```javascript
export async function columnMapperRequest(action, payload, callsiteTag) {
    return warehouseRequest("column-mapper", { action, ...payload }, callsiteTag || action);
}
```

**Request Parameters for Bootstrap**:
```javascript
{
    action: "bootstrap",
    installation_key: "pf_install_9f3c2b1a_20251212"
}
```

**Successful Response Structure**:
```javascript
{
    ok: true,
    status: 200,
    data: {
        success: true,
        company_id: "...",
        ss_company_name: "...",
        ss_accounting_software: "...",
        pto_payroll_provider: "...",
        pr_payroll_provider: "..."
    }
}
```

**Failed Response Structure**:
```javascript
{
    ok: false,
    status: 401|500|etc,
    error: "error message"
}
```

**Authentication**:
- Uses Supabase anon key stored in `warehouse.js` (Line 25)
- Token loaded from `OfficeRuntime.storage` if available, falls back to built-in anon key
- Auth refreshed before every bootstrap call via `forceRefreshAuth()`

---

### Answer 1.4: Config Writing to Excel

**Function**: `writeConfigValues()` in `/Users/d.paeth/Customer-Tai-Software/Common/bootstrap.js`

**Complete Function** (Lines 166-193):
```javascript
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
```

**Uses**: `saveConfigValue()` from `/Users/d.paeth/Customer-Tai-Software/Common/gateway.js`

**Target Table**: `SS_PF_Config` (Excel table)

**Columns Written**:
| API Response Key | SS_PF_Config Field | Example Value |
|-----------------|-------------------|---------------|
| installation_key | SS_Installation_Key | pf_install_9f3c2b1a_20251212 |
| company_id | SS_Company_ID | company_123 |
| ss_company_name | SS_Company_Name | Acme Corp |
| ss_accounting_software | SS_Accounting_Software | QuickBooks |
| pto_payroll_provider | PTO_Payroll_Provider | ADP |
| pr_payroll_provider | PR_Payroll_Provider | Paychex |

**Table Structure**: SS_PF_Config is a key-value table with columns:
- Field (key name)
- Value (field value)
- Type (optional)
- Title (optional)
- Permanent (optional)

---

## Section 2: Project Structure

### Answer 2.1: Source Code Organization

**Directory Structure**:
```
Customer-Tai-Software/
├── Common/                    # Shared utilities across modules
│   ├── bootstrap.js          # Bootstrap config sync (CRITICAL)
│   ├── warehouse.js          # API client for Supabase Edge Functions
│   ├── gateway.js            # Excel helpers, config read/write
│   ├── sheet-formatting.js   # Excel formatting utilities
│   └── [18 other utility files]
├── module-selector/          # Landing page / module launcher
│   ├── selector.js           # Main entry point, calls bootstrap
│   ├── index.html            # Task pane HTML
│   └── selector.css          # Styles
├── payroll-recorder/         # Payroll module
│   ├── src/
│   │   └── workflow.js       # Source code (bundled to app.bundle.js)
│   ├── app.bundle.js         # Built output
│   ├── app.bundle.js.map     # Source map
│   ├── index.html            # Module task pane
│   └── [other files]
├── pto-accrual/              # PTO module
│   ├── src/
│   │   └── index.js          # Source code (bundled to app.bundle.js)
│   ├── app.bundle.js         # Built output
│   ├── app.bundle.js.map     # Source map
│   ├── index.html            # Module task pane
│   └── [other files]
├── scripts/                  # Build scripts
│   ├── build-payroll.js      # esbuild config for payroll
│   └── build-pto.js          # esbuild config for PTO
├── assets/                   # Icons and images
├── docs/                     # Documentation
├── supabase/                 # Supabase Edge Functions
│   └── functions/
│       └── column-mapper/
│           └── index.ts      # Edge Function code
├── TaiTools_manifest.xml     # Office Add-in manifest
├── package.json              # Dependencies and build scripts
└── [other config files]
```

**Key Observations**:
- **Source files**: `src/` subdirectories in each module
- **Built files**: `app.bundle.js` in module root directories
- **Shared code**: `Common/` directory (not bundled, loaded directly)
- **Entry points**: `module-selector/selector.js` (first load), then module-specific bundles

---

### Answer 2.2: Entry Point

**Primary Entry Point**: `/Users/d.paeth/Customer-Tai-Software/module-selector/selector.js`

**Initialization Flow**:
1. Office.onReady fires
2. `init()` function called (Line 502)
3. `ensureConfigSheetAndTable()` - Creates SS_PF_Config if needed
4. **`bootstrapConfigSync()`** - Syncs config from Supabase
5. UI rendering (hero, modules, etc.)

**First 50 Lines of selector.js**:
```javascript
/**
 * Module Selector - Landing page for ForgeSuite
 * Allows users to choose between Payroll Recorder and PTO Accrual
 */

import { 
    initializeOffice, 
    hasExcelRuntime, 
    loadConfigFromTable,
    DEFAULT_COLUMN_ALIASES,
    clearColumnAliasCache
} from "../Common/gateway.js";
import { bootstrapConfigSync } from "../Common/bootstrap.js";
import { formatSheetHeaders } from "../Common/sheet-formatting.js";

/**
 * Module definitions
 */
const MODULES = [
    {
        id: "payroll-recorder",
        name: "Payroll Recorder",
        description: "Transform payroll data into journal entries",
        icon: "💰",
        available: true,
        url: "./payroll-recorder/index.html"
    },
    {
        id: "pto-accrual",
        name: "PTO Accrual",
        description: "Calculate PTO liability and generate journal entries",
        icon: "🏖️",
        available: true,
        url: "./pto-accrual/index.html"
    }
];

// DOM elements
const heroGreetingEl = document.getElementById("heroGreeting");
const moduleCountEl = document.getElementById("moduleCount");
const modulesContainerEl = document.getElementById("modulesContainer");

// Initialize when Office is ready
initializeOffice((info) => {
    console.log("Office initialized:", info);
    init();
});

async function init() {
    try {
        console.log("ForgeSuite init starting...");
        // ... (continues with bootstrap call)
```

---

### Answer 2.3: Technology Stack

**From package.json**:

**Language**: JavaScript (ES2019+, no TypeScript in client code)

**Framework**: Vanilla JavaScript (no React/Vue/Angular)

**Bundler**: esbuild v0.27.0

**Key Dependencies**:
```json
{
  "devDependencies": {
    "@eslint/js": "^9.39.1",
    "esbuild": "^0.27.0",
    "eslint": "^9.39.1",
    "globals": "^16.1.0"
  },
  "dependencies": {
    "xlsx": "^0.18.5"  // SheetJS for Excel file generation
  }
}
```

**Office.js**: Loaded via CDN in HTML files
```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
```

**No Frontend Framework**: Pure JavaScript with manual DOM manipulation

**Build System**: Custom esbuild scripts (not webpack/rollup/vite)

---

## Section 3: Build and Deployment

### Answer 3.1: Build Configuration

**Build Scripts**: Custom Node.js scripts using esbuild API

**Payroll Build**: `/Users/d.paeth/Customer-Tai-Software/scripts/build-payroll.js`
```javascript
const esbuild = require('esbuild');

await esbuild.build({
    entryPoints: ['payroll-recorder/src/workflow.js'],
    bundle: true,
    outfile: 'payroll-recorder/app.bundle.js',
    format: 'iife',
    platform: 'browser',
    target: ['es2019'],
    sourcemap: true,
    minify: true,
    banner: {
        js: '/* Prairie Forge Payroll Recorder */'
    },
    define: {
        __BUILD_COMMIT__: JSON.stringify(COMMIT_HASH)
    }
});
```

**PTO Build**: `/Users/d.paeth/Customer-Tai-Software/scripts/build-pto.js` (similar structure)

**Build Commands**:
```json
{
  "build:payroll": "node scripts/build-payroll.js",
  "build:pto": "node scripts/build-pto.js",
  "build:all": "npm run build:payroll && npm run build:pto"
}
```

**Output Directories**:
- Payroll: `payroll-recorder/app.bundle.js`
- PTO: `pto-accrual/app.bundle.js`

**HTML Files**: NOT bundled - referenced directly with cache-busting query params

**Deployment**: GitHub Pages
- Repository: `prairie-forge-ai/Customer-Tai-Software`
- URL: `https://prairie-forge-ai.github.io/Customer-Tai-Software/`

**Cache Busting**: Build scripts update version hashes in HTML files

---

### Answer 3.2: Manifest File

**File**: `/Users/d.paeth/Customer-Tai-Software/TaiTools_manifest.xml`

**Current Version**: 2.0.1.0

**Key URLs**:
```xml
<SourceLocation DefaultValue="https://prairie-forge-ai.github.io/Customer-Tai-Software/module-selector/index.html"/>
<IconUrl DefaultValue="https://prairie-forge-ai.github.io/Customer-Tai-Software/assets/icon-80.png"/>
<SupportUrl DefaultValue="https://prairieforge.ai/support"/>
```

**AppDomains**:
```xml
<AppDomains>
  <AppDomain>https://prairie-forge-ai.github.io</AppDomain>
  <AppDomain>https://assets.prairieforge.ai</AppDomain>
  <AppDomain>https://prairieforge.ai</AppDomain>
</AppDomains>
```

**Add-in ID**: `16b2f680-b42c-4f45-ab67-af2d2c3c9d15`

---

## Section 4: Excel Integration Details

### Answer 4.1: Excel Settings API Usage

**Current Usage**: ❌ **NOT USED**

**Search Results**: No usage of `context.workbook.settings` found in codebase

**Current Storage Methods**:
1. **SS_PF_Config Table** - Excel table for configuration (key-value pairs)
2. **OfficeRuntime.storage** - Used only for auth token caching
3. **In-memory cache** - Bootstrap results cached in `bootstrap.js`

**Opportunity**: Excel Settings API could be used for installation key storage

---

### Answer 4.2: Startup Sequence

**Complete Startup Flow**:

1. **Office.onReady** fires (in module-selector/selector.js)
   ```javascript
   initializeOffice((info) => {
       console.log("Office initialized:", info);
       init();
   });
   ```

2. **init()** function executes (Line 502)

3. **ensureConfigSheetAndTable()** - Creates SS_PF_Config if missing
   - Creates "SS_PF_Config" sheet
   - Creates table with columns: Field, Value, Type, Title, Permanent
   - Writes default values

4. **bootstrapConfigSync()** - CRITICAL STEP
   - Refreshes auth token
   - Calls Edge Function with installation_key
   - Validates response
   - Writes config to SS_PF_Config table
   - Caches result

5. **renderHero()** - Displays greeting

6. **renderModules()** - Shows module cards

7. **wireActions()** - Attaches event listeners

8. **initQuickAccess()** - Sets up quick access features

9. **applyModuleTabVisibility()** - Manages Excel sheet visibility

10. **activateHomepageWithRetry()** - Activates homepage sheet

11. **renderAdaFab()** - Shows AI assistant button

**Functions Called in Order**:
```javascript
async function init() {
    await ensureConfigSheetAndTable();
    await bootstrapConfigSync();
    renderHero();
    renderModules();
    wireActions();
    initQuickAccess();
    await applyModuleTabVisibility("module-selector");
    await activateHomepageWithRetry();
    renderAdaFab();
}
```

---

## Section 5: Current User Experience

### Answer 5.1: Module Selector

**What It Is**: Landing page / launcher for the add-in

**File**: `/Users/d.paeth/Customer-Tai-Software/module-selector/selector.js`

**Purpose**: 
- First page users see when opening TaiTools
- Displays available modules (Payroll Recorder, PTO Accrual)
- Handles global initialization including bootstrap

**Bootstrap Call Location** (Lines 522-533):
```javascript
console.log("[Init] Running bootstrap config sync...");
try {
    const bootstrapResult = await bootstrapConfigSync();
    console.log("[Init] Bootstrap result:", bootstrapResult);
    if (!bootstrapResult.success) {
        console.warn("[Init] Bootstrap failed but continuing:", bootstrapResult.error);
    }
} catch (bootstrapError) {
    console.error("[Init] Bootstrap error (non-fatal):", bootstrapError);
    // Continue anyway - manual config still works
}
```

**When It Happens**: 
- During `init()` function
- After `Office.onReady` fires
- AFTER ensuring SS_PF_Config exists
- BEFORE rendering UI

**Error Handling**: Non-fatal - UI still loads if bootstrap fails

---

### Answer 5.2: Error Handling

**Bootstrap Errors** (in bootstrap.js):
```javascript
// Stage 1: Fetch error
catch (fetchError) {
    console.error("[Bootstrap] ❌ Fetch failed:", fetchError.message);
    return { success: false, error: fetchError.message, stage: "fetch" };
}

// Stage 2: Validation error
catch (validationError) {
    console.error("[Bootstrap] ❌ Validation failed:", validationError.message);
    return { success: false, error: validationError.message, stage: "validate" };
}

// Stage 3: Write error
catch (writeError) {
    console.error("[Bootstrap] ❌ Write failed:", writeError.message);
    return { success: false, error: writeError.message, stage: "write" };
}
```

**Module Selector Handling** (in selector.js):
```javascript
try {
    const bootstrapResult = await bootstrapConfigSync();
    if (!bootstrapResult.success) {
        console.warn("[Init] Bootstrap failed but continuing:", bootstrapResult.error);
    }
} catch (bootstrapError) {
    console.error("[Init] Bootstrap error (non-fatal):", bootstrapError);
    // Continue anyway - manual config still works
}
```

**User Notification**: 
- ❌ No toast/alert shown to user for bootstrap failures
- ✅ Errors logged to console
- ✅ Add-in continues to function (manual config still works)

**SS_PF_Config Write Failures**:
- Individual field errors collected in `results.errors` array
- Logged to console but not blocking
- Partial success possible (some fields written, others skipped)

---

## Section 6: Supabase Integration

### Answer 6.1: Database Schema

**Table**: `ada_addin_installations`

**Expected Columns** (based on code):
- `installation_key` (primary key)
- `company_id`
- `ss_company_name`
- `ss_accounting_software`
- `pto_payroll_provider`
- `pr_payroll_provider`

**Note**: Actual schema should be confirmed in Supabase dashboard

**Related Tables** (from column-mapper function):
- `ada_customer_column_mappings` - Column mapping storage
- `ada_pf_columns` - Prairie Forge canonical columns
- `ada_pf_dimensions` - Dimension definitions

---

### Answer 6.2: Edge Function Details

**File**: `/Users/d.paeth/Customer-Tai-Software/supabase/functions/column-mapper/index.ts`

**Endpoint**: `https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1/column-mapper`

**Bootstrap Action**:
```typescript
interface ColumnMapperRequest {
  action: "analyze" | "save" | "get_options" | "get_expense_taxonomy" | "bootstrap" | "debug" | "get_dimensions";
  installation_key?: string;
  // ... other fields
}
```

**Expected Validation**:
1. Receives `{ action: "bootstrap", installation_key: "pf_install_..." }`
2. Queries `ada_addin_installations` table
3. Returns company config if installation_key exists
4. Returns error if not found or invalid

**Response Format**:
```typescript
{
  success: true,
  company_id: string,
  ss_company_name: string,
  ss_accounting_software: string,
  pto_payroll_provider: string,
  pr_payroll_provider: string
}
```

---

## Section 7: Testing Environment

### Answer 7.1: Development Setup

**No Dev Server**: Static files served from GitHub Pages

**Local Testing**:
1. Make code changes in `src/` files
2. Run build: `npm run build:payroll` or `npm run build:pto`
3. Commit and push to GitHub
4. GitHub Pages serves updated files
5. Sideload manifest in Excel

**Sideloading**:
- Use `TaiTools_manifest.xml`
- File > Options > Trust Center > Trusted Add-in Catalogs
- Or use Office Add-ins Developer Mode

**Build Commands**:
```bash
npm run build:payroll  # Build payroll module
npm run build:pto      # Build PTO module
npm run build:all      # Build both modules
```

---

### Answer 7.2: Existing Tests

**Test Directory**: `/Users/d.paeth/Customer-Tai-Software/tests/`

**Test Command**: `npm test` (runs `node --test "tests/**/*.mjs"`)

**Testing Framework**: Node.js built-in test runner (no Jest/Mocha)

**Test Files**: Would need to examine tests/ directory for specifics

---

## Section 8: Critical Implementation Notes

### Current Limitations

1. **Hardcoded Installation Key**: Must be changed in source code and rebuilt
2. **No Runtime Validation**: Installation key not validated until bootstrap runs
3. **No User Feedback**: Bootstrap failures are silent to end users
4. **GitHub Pages Deployment**: Requires commit/push for any changes
5. **No Environment Variables**: All config is hardcoded or in Excel

### Opportunities for Improvement

1. **Excel Settings API**: Store installation key persistently
2. **Manifest CustomProperties**: Pass installation key via manifest
3. **User Prompt**: Allow user to enter installation key on first run
4. **Validation UI**: Show clear error messages for invalid keys
5. **Admin Panel**: UI for managing installation settings

---

## Next Steps

With this discovery information, you can now create:
1. Precise implementation instructions for installation key validation
2. Code modifications that integrate with existing patterns
3. Build configuration updates
4. Testing procedures
5. Deployment workflow

---

**Document Generated**: January 12, 2026  
**For**: Prairie Forge / TaiTools Installation Key Implementation  
**Contact**: connect@prairieforge.ai
