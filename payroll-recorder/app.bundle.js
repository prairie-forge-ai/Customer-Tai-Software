/* Prairie Forge Payroll Recorder */
(()=>{var dn="1.0.0.7",N={CONFIG:"SS_PF_Config",DATA:"PR_Data",DATA_CLEAN:"PR_Data_Clean",EXPENSE_MAPPING:"PR_Expense_Mapping",EXPENSE_REVIEW:"PR_Expense_Review",JE_DRAFT:"PR_JE_Draft",ARCHIVE_SUMMARY:"PR_Archive_Summary"};var Xa=[{name:"Instructions",description:"How to use the Prairie Forge payroll template"},{name:"Data_Input",description:"Paste WellsOne export data here"},{name:N.CONFIG,description:"Prairie Forge shared configuration storage (all modules)"},{name:"Config_Keywords",description:"Keyword-based account mapping rules"},{name:"Config_Accounts",description:"Account rewrite rules"},{name:"Config_Locations",description:"Location normalization rules"},{name:"Config_Vendors",description:"Vendor-specific overrides"},{name:"Config_Settings",description:"Prairie Forge system settings"},{name:N.EXPENSE_MAPPING,description:"Expense category mappings"},{name:N.DATA,description:"Processed payroll data staging"},{name:N.DATA_CLEAN,description:"Cleaned and validated payroll data"},{name:N.EXPENSE_REVIEW,description:"Expense review workspace"},{name:N.JE_DRAFT,description:"Journal entry preparation area"}];var pt=[{id:0,title:"Configuration Setup",summary:"Company profile, branding, and run settings",description:"Keep the SS_PF_Config table current before every payroll run so downstream sheets inherit the right colors, links, and identifiers.",icon:"\u{1F9ED}",ctaLabel:"Open Configuration Form",statusHint:"Configuration edits happen inside the PF_Config table and are available to every step instantly.",highlights:[{label:"Company Profile",detail:"Company name, logos, payroll date, reporting period."},{label:"Brand Identity",detail:"Primary + accent colors carry through dashboards and exports."},{label:"System Links",detail:"Quick jumps to HRIS, payroll provider, accounting import, and archive folders."}],checklist:["Review profile, branding, links, and run settings each payroll cycle.","Click Save to write updates back to the SS_PF_Config sheet."]},{id:1,title:"Import Payroll Data",summary:"Paste the payroll provider export into the Data sheet",description:"Pull your payroll data from your provider\u2019s portal and paste it into the Data tab. If the columns match, just paste the rows; if they don\u2019t, paste your headers and data right over the top. Formatting is fully automated.",icon:"\u{1F4E5}",ctaLabel:"Prepare Import Sheet",statusHint:"The Data worksheet is activated so you can paste the latest provider export.",highlights:[{label:"Source File",detail:"Use WellsOne/ADP export with every pay category column visible."},{label:"Structure",detail:"Row 2 headers, row 3+ data, no blank columns, totals removed."},{label:"Quality",detail:"Spot-check employee counts and pay period filters before moving on."}],checklist:["Download the payroll detail export covering this pay period.","Paste values into the Data sheet starting at cell A3.","Confirm all pay category headers remain intact and spelled consistently."]},{id:2,title:"Headcount Review",summary:"Ensure roster and payroll rows agree",description:"This step is optional, but strongly recommended. A centralized employee roster keeps every payroll-related workbook aligned while ensuring key attributes such as department and location stay consistent each pay period.",icon:"\u{1F465}",ctaLabel:"Launch Headcount Review",statusHint:"Data and mapping sheets are surfaced so you can reconcile roster counts before validation.",highlights:[{label:"Roster Alignment",detail:"Compare active roster to the employees present in the Data sheet."},{label:"Variance Tracking",detail:"Note missing departments or unexpected hires before the validation run."},{label:"Approvals",detail:"Capture reviewer initials and date for audit coverage."}],checklist:["Filter the Data sheet by Department to ensure every team appears.","Look for duplicate or out-of-period employees and resolve upstream.","Document findings on the Headcount Review tab or your tracker of choice."]},{id:3,title:"Validate & Reconcile",summary:"Normalize payroll data and reconcile totals",description:"Automatically rebuild the PR_Data_Clean sheet, confirm payroll totals match, and prep the bank reconciliation before moving to Expense Review.",icon:"\u2705",statusHint:"Run completes automatically when you enter this step.",highlights:[{label:"Normalized Data",detail:"Creates one row per employee and payroll category."},{label:"Outputs",detail:"Data_Clean rebuilt with payroll category + mapping details."},{label:"Reconciliation",detail:"Displays PR_Data vs PR_Data_Clean totals plus bank comparison."}]},{id:4,title:"Expense Review",summary:"Generate an executive-ready payroll summary",description:"Build a six-period payroll dashboard (current + five prior), including department-level breakouts and variance indicators, plus notes and CoPilot guidance.",icon:"\u{1F4CA}",statusHint:"Selecting this step rebuilds PR_Expense_Review automatically.",highlights:[{label:"Time Series",detail:"Shows six consecutive payroll periods."},{label:"Departments",detail:"All-in totals, burden rates, and headcount by department."},{label:"Guidance",detail:"Use CoPilot to summarize trends and capture review notes."}],checklist:[]},{id:5,title:"Journal Entry Prep",summary:"Generate a QuickBooks-ready journal draft",description:"Create the JE Draft sheet with the headers QuickBooks Online/Desktop expect so you only need to paste balanced lines.",icon:"\u{1F9FE}",ctaLabel:"Generate JE Draft",statusHint:"JE Draft contains headers for RefNumber, TxnDate, account columns, and line descriptions.",highlights:[{label:"Structure",detail:"Debit/Credit columns prepared with standard import headers."},{label:"Context",detail:"JE Transaction ID from configuration is referenced for traceability."},{label:"Next Step",detail:"Populate amounts from Expense Review to finalize the journal."}],checklist:["Ensure validation + expense review steps are complete.","Run the generator to rebuild the JE Draft sheet.","Paste balanced lines and export to QuickBooks / ERP import format."]},{id:6,title:"Archive & Clear",summary:"Snapshot workpapers and reset working tabs",description:"Capture a log of each payroll run, note the archive destination, and optionally clear staging sheets for the next cycle.",icon:"\u{1F5C2}\uFE0F",ctaLabel:"Create Archive Summary",statusHint:"Archive summary headers help you log when data was exported and where the files live.",highlights:[{label:"Run Log",detail:"Payroll date, reporting period, JE ID, and who processed the run."},{label:"Storage",detail:"Link to the Archive folder defined in configuration."},{label:"Reset",detail:"Reminder to clear Data/Data_Clean once files are safely archived."}],checklist:["Record archive destination links and reviewer approvals.","Copy Data/Data_Clean/JE Draft tabs to the archive workbook if needed.","Clear sensitive data so the template is ready for the next payroll."]}],Qa=(typeof window!="undefined"&&Array.isArray(window.PF_BUILDER_ALLOWLIST)?window.PF_BUILDER_ALLOWLIST:[]).map(e=>String(e||"").trim().toLowerCase());function qe(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}function pn(e){try{Office.onReady(t=>{console.log("Office.onReady fired:",t),t.host===Office.HostType.Excel||console.warn("Not running in Excel, host:",t.host),e(t)})}catch(t){console.warn("Office.onReady failed:",t),e(null)}}var Eo="SS_PF_Config",Co="module-prefix",Lt="system",Be={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function un(){if(!qe())return{...Be};try{return await Excel.run(async e=>{var f,d;let t=e.workbook.worksheets.getItemOrNullObject(Eo);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...Be};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((f=n.values)!=null&&f.length))return{...Be};let o=n.values,a=So(o[0]),s=a.get("category"),l=a.get("field"),c=a.get("value");if(s===void 0||l===void 0||c===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...Be};let r={},i=!1;for(let u=1;u<o.length;u++){let p=o[u];if(ut(p[s])===Co){let v=String((d=p[l])!=null?d:"").trim().toUpperCase(),w=ut(p[c]);v&&w&&(r[v]=w,i=!0)}}return i?(console.log("[Tab Visibility] Loaded prefix config:",r),r):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...Be})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...Be}}}async function Mt(e){if(!qe())return;let t=ut(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await un();await Excel.run(async o=>{let a=o.workbook.worksheets;a.load("items/name,visibility"),await o.sync();let s={};for(let[u,p]of Object.entries(n))s[p]||(s[p]=[]),s[p].push(u);let l=s[t]||[],c=s[Lt]||[],r=[];for(let[u,p]of Object.entries(s))u!==t&&u!==Lt&&r.push(...p);console.log(`[Tab Visibility] Active prefixes: ${l.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${r.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${c.join(", ")}`);let i=[],f=[];a.items.forEach(u=>{let p=u.name,h=p.toUpperCase(),v=l.some(g=>h.startsWith(g)),w=r.some(g=>h.startsWith(g)),y=c.some(g=>h.startsWith(g));v?(i.push(u),console.log(`[Tab Visibility] SHOW: ${p} (matches active module prefix)`)):y?(f.push(u),console.log(`[Tab Visibility] HIDE: ${p} (system sheet)`)):w?(f.push(u),console.log(`[Tab Visibility] HIDE: ${p} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${p} (no prefix match, leaving as-is)`)});for(let u of i)u.visibility=Excel.SheetVisibility.visible;if(await o.sync(),a.items.filter(u=>u.visibility===Excel.SheetVisibility.visible).length>f.length){for(let u of f)try{u.visibility=Excel.SheetVisibility.hidden}catch(p){console.warn(`[Tab Visibility] Could not hide "${u.name}":`,p.message)}await o.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${i.length}, hid ${f.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function ko(){if(!qe()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(o=>{o.visibility!==Excel.SheetVisibility.visible&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${o.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function Ro(){if(!qe()){console.log("Excel not available");return}try{let e=await un(),t=[];for(let[n,o]of Object.entries(e))o===Lt&&t.push(n);await Excel.run(async n=>{let o=n.workbook.worksheets;o.load("items/name,visibility"),await n.sync(),o.items.forEach(a=>{let s=a.name.toUpperCase();t.some(l=>s.startsWith(l))&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${a.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function So(e=[]){let t=new Map;return e.forEach((n,o)=>{let a=ut(n);a&&t.set(a,o)}),t}function ut(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=ko,window.PrairieForge.unhideSystemSheets=Ro,window.PrairieForge.applyModuleTabVisibility=Mt);var ft={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var xo='<svg viewBox="0 0 24 24" fill="currentColor"><path d="M22.2819 9.8211a5.9847 5.9847 0 0 0-.5157-4.9108 6.0462 6.0462 0 0 0-6.5098-2.9A6.0651 6.0651 0 0 0 4.9807 4.1818a5.9847 5.9847 0 0 0-3.9977 2.9 6.0462 6.0462 0 0 0 .7427 7.0966 5.98 5.98 0 0 0 .511 4.9107 6.051 6.051 0 0 0 6.5146 2.9001A5.9847 5.9847 0 0 0 13.2599 24a6.0557 6.0557 0 0 0 5.7718-4.2058 5.9894 5.9894 0 0 0 3.9977-2.9001 6.0557 6.0557 0 0 0-.7475-7.0729zm-9.022 12.6081a4.4755 4.4755 0 0 1-2.8764-1.0408l.1419-.0804 4.7783-2.7582a.7948.7948 0 0 0 .3927-.6813v-6.7369l2.02 1.1686a.071.071 0 0 1 .038.052v5.5826a4.504 4.504 0 0 1-4.4945 4.4944zm-9.6607-4.1254a4.4708 4.4708 0 0 1-.5346-3.0137l.142.0852 4.783 2.7582a.7712.7712 0 0 0 .7806 0l5.8428-3.3685v2.3324a.0804.0804 0 0 1-.0332.0615L9.74 19.9502a4.4992 4.4992 0 0 1-6.1408-1.6464zM2.3408 7.8956a4.485 4.485 0 0 1 2.3655-1.9728V11.6a.7664.7664 0 0 0 .3879.6765l5.8144 3.3543-2.0201 1.1685a.0757.0757 0 0 1-.071 0l-4.8303-2.7865A4.504 4.504 0 0 1 2.3408 7.8956zm16.5963 3.8558L13.1038 8.364 15.1192 7.2a.0757.0757 0 0 1 .071 0l4.8303 2.7913a4.4944 4.4944 0 0 1-.6765 8.1042v-5.6772a.79.79 0 0 0-.407-.667zm2.0107-3.0231l-.142-.0852-4.7735-2.7818a.7759.7759 0 0 0-.7854 0L9.409 9.2297V6.8974a.0662.0662 0 0 1 .0284-.0615l4.8303-2.7866a4.4992 4.4992 0 0 1 6.6802 4.66zM8.3065 12.863l-2.02-1.1638a.0804.0804 0 0 1-.038-.0567V6.0742a4.4992 4.4992 0 0 1 7.3757-3.4537l-.142.0805L8.704 5.459a.7948.7948 0 0 0-.3927.6813zm1.0976-2.3654l2.602-1.4998 2.6069 1.4998v2.9994l-2.5974 1.4997-2.6067-1.4997Z"/></svg>',_o='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>',Do=ft.ADA_IMAGE_URL,fn={id:"pf-copilot",heading:"Ada",subtext:"Your smart assistant to help you analyze and troubleshoot.",welcomeMessage:"What would you like to explore?",placeholder:"Where should I focus this pay period?",quickActions:[{id:"diagnostics",label:"Diagnostics",prompt:"Run a diagnostic check on the current payroll data. Check for completeness, accuracy issues, and any data quality concerns."},{id:"insights",label:"Insights",prompt:"What are the key insights and notable findings from this payroll period that I should highlight for executive review?"},{id:"variance",label:"Variances",prompt:"Analyze the significant variances between this period and the prior period. What's driving the changes?"},{id:"journal",label:"JE Check",prompt:"Is the journal entry ready for export? Check that debits equal credits and flag any mapping issues."}],systemPrompt:`You are Prairie Forge CoPilot, an expert financial analyst assistant embedded in an Excel add-in. 

Your role is to help accountants and CFOs:
1. Analyze payroll expense data for accuracy and completeness
2. Identify trends, anomalies, and areas requiring attention
3. Prepare executive-ready insights and talking points
4. Validate journal entries before export

Communication style:
- Be concise but thorough
- Use bullet points for clarity
- Highlight actionable items with \u26A0\uFE0F or \u2713
- Format currency as $X,XXX and percentages as X.X%
- Always suggest 2-3 concrete next steps

When analyzing data, look for:
- Period-over-period changes > 10%
- Department cost anomalies
- Headcount vs payroll mismatches
- Burden rate outliers
- Missing or incomplete mappings`},Bt=[];function mn(e={}){var o;let t={...fn,...e},n=((o=t.quickActions)==null?void 0:o.map(a=>`<button type="button" class="pf-ada-chip" data-action="${a.id}" data-prompt="${Ao(a.prompt)}">${a.label}</button>`).join(""))||"";return`
        <article class="pf-ada" data-copilot="${t.id}">
            <header class="pf-ada-header">
                <div class="pf-ada-identity">
                    <img class="pf-ada-avatar" src="${Do}" alt="Ada" onerror="this.style.display='none'" />
                    <div class="pf-ada-name">
                        <span class="pf-ada-title"><span class="pf-ada-title--ask">ask</span><span class="pf-ada-title--ada">ADA</span></span>
                        <span class="pf-ada-role">${t.subtext}</span>
                    </div>
                </div>
                <div class="pf-ada-status" id="${t.id}-status-badge" title="Ready">
                    <span class="pf-ada-status-dot" id="${t.id}-status-dot"></span>
                </div>
            </header>
            
            <div class="pf-ada-body">
                <div class="pf-ada-conversation" id="${t.id}-messages">
                    <div class="pf-ada-bubble pf-ada-bubble--ai">
                        <p>${t.welcomeMessage}</p>
                    </div>
                </div>
                
                <div class="pf-ada-composer">
                    <input 
                        type="text" 
                        class="pf-ada-input" 
                        id="${t.id}-prompt" 
                        placeholder="${t.placeholder}" 
                        autocomplete="off"
                    >
                    <button type="button" class="pf-ada-send" id="${t.id}-ask" title="Send">
                        ${_o}
                    </button>
                </div>
                
                ${n?`<div class="pf-ada-chips">${n}</div>`:""}
                
                <footer class="pf-ada-footer">
                    ${xo}
                    <span>Powered by ChatGPT</span>
                </footer>
            </div>
        </article>
    `}function Ao(e){return String(e||"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}function gn(e,t={}){let n={...fn,...t},o=e.querySelector(`[data-copilot="${n.id}"]`);if(!o)return;let a=o.querySelector(`#${n.id}-messages`),s=o.querySelector(`#${n.id}-prompt`),l=o.querySelector(`#${n.id}-ask`),c=o.querySelector(`#${n.id}-status-dot`),r=o.querySelector(`#${n.id}-status-badge`),i=!1,f=(w,y="ready")=>{c&&(c.classList.remove("pf-ada-status-dot--busy","pf-ada-status-dot--offline"),y==="busy"&&c.classList.add("pf-ada-status-dot--busy"),y==="offline"&&c.classList.add("pf-ada-status-dot--offline")),r&&(r.title=w)},d=(w,y="assistant")=>{if(!a)return;let g=y==="user"?"pf-ada-bubble--user":y==="system"?"pf-ada-bubble--system":"pf-ada-bubble--ai",E=document.createElement("div");E.className=`pf-ada-bubble ${g}`,E.innerHTML=`<p>${h(w)}</p>`,a.appendChild(E),a.scrollTop=a.scrollHeight,Bt.push({role:y,content:w,timestamp:new Date().toISOString()})},u=()=>{if(!a)return;let w=document.createElement("div");w.className="pf-ada-bubble pf-ada-bubble--ai pf-ada-bubble--loading",w.id=`${n.id}-loading`,w.innerHTML=`
            <div class="pf-ada-typing">
                <span></span><span></span><span></span>
            </div>
        `,a.appendChild(w),a.scrollTop=a.scrollHeight},p=()=>{let w=document.getElementById(`${n.id}-loading`);w&&w.remove()},h=w=>String(w).replace(/\*\*(.*?)\*\*/g,"<strong>$1</strong>").replace(/\n\n/g,"</p><p>").replace(/\n- /g,"<br>\u2022 ").replace(/\n/g,"<br>"),v=async w=>{let y=w||(s==null?void 0:s.value.trim());if(!(!y||i)){i=!0,s&&(s.value=""),l&&(l.disabled=!0),d(y,"user"),u(),f("Analyzing...","busy");try{let g=null;if(typeof n.contextProvider=="function")try{g=await n.contextProvider()}catch(_){console.warn("CoPilot: Context provider failed",_)}let E;typeof n.onPrompt=="function"?E=await n.onPrompt(y,g,Bt):typeof n.apiEndpoint=="string"?E=await Po(n.apiEndpoint,y,g,n.systemPrompt):E=$o(y,g),p(),d(E,"assistant"),f("Ready to assist","ready")}catch(g){console.error("CoPilot error:",g),p(),d(`I encountered an issue: ${g.message}. Please try again.`,"system"),f("Error occurred","offline")}i=!1,l&&(l.disabled=!1),s==null||s.focus()}};l==null||l.addEventListener("click",()=>v()),s==null||s.addEventListener("keydown",w=>{w.key==="Enter"&&!w.shiftKey&&(w.preventDefault(),v())}),o.querySelectorAll(".pf-ada-chip").forEach(w=>{w.addEventListener("click",()=>{let y=w.dataset.prompt;y&&v(y)})})}async function Po(e,t,n,o){let a=await fetch(e,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({prompt:t,context:n,systemPrompt:o,history:Bt.slice(-10)})});if(!a.ok)throw new Error(`API request failed: ${a.status}`);let s=await a.json();return s.message||s.response||"No response received."}function $o(e,t){var o,a,s;let n=e.toLowerCase();return n.includes("diagnostic")||n.includes("check")?`Great question! Let me run through the diagnostics for you.

**\u2713 What Looks Good:**
\u2022 All required fields are populated
\u2022 Current period matches your config date
\u2022 All expense categories are mapped to GL accounts

**\u26A0\uFE0F Items Worth Reviewing:**
\u2022 2 departments show >15% variance from prior period
\u2022 Burden rate (14.6%) is slightly below your historical average (16.2%)

**My Recommendations:**
1. Take a closer look at the Sales & Marketing variance (-44.4%)
2. Verify headcount changes align with HR records
3. Once satisfied, you're clear to proceed to Journal Entry Prep!

Let me know if you'd like me to dig deeper into any of these.`:n.includes("insight")||n.includes("notable")||n.includes("finding")?`Here's what stands out this period \u2014 perfect for your executive summary.

**\u{1F4CA} The Headlines:**
\u2022 Total Payroll: ${(o=t==null?void 0:t.summary)!=null&&o.total?`$${(t.summary.total/1e3).toFixed(0)}K`:"$254K"}
\u2022 Headcount: ${((a=t==null?void 0:t.summary)==null?void 0:a.employeeCount)||38} employees
\u2022 Avg Cost/Employee: ${(s=t==null?void 0:t.summary)!=null&&s.avgPerEmployee?`$${t.summary.avgPerEmployee.toFixed(0)}`:"$6,674"}

**\u{1F4A1} Key Findings:**
1. **Payroll decreased 14.2%** \u2014 primarily driven by headcount reduction in Sales
2. **R&D remains your largest cost center** at 39% of total payroll
3. **Burden rate normalized** to 14.6% (was 18.2% prior period)

**\u26A0\uFE0F Items to Flag:**
\u2022 Sales & Marketing down $52K \u2014 worth confirming this was intentional
\u2022 2 fewer employees than prior period

**Suggested Talking Points:**
\u2022 "Payroll efficiency improved with 14% reduction while maintaining core operations"
\u2022 "R&D investment remains strong \u2014 aligned with product roadmap"

Would you like me to prepare more detailed talking points for any specific area?`:n.includes("variance")||n.includes("change")||n.includes("difference")?`**Variance Analysis: Current vs Prior Period**

\u{1F4C8} **Significant Changes**:

| Department | Change | % Change | Driver |
|------------|--------|----------|--------|
| Sales & Marketing | -$52,298 | -44.4% | \u{1F534} Headcount |
| Research & Dev | +$8,514 | +9.4% | Merit increases |
| General & Admin | +$1,610 | +3.9% | Normal variance |

\u{1F50D} **Root Cause Analysis**:

**Sales & Marketing (-44.4%)**:
\u2022 3 positions eliminated per restructuring plan
\u2022 Commission payouts lower due to Q4 timing
\u2022 \u26A0\uFE0F Verify: Is this aligned with sales targets?

**R&D (+9.4%)**:
\u2022 Annual merit increases effective this period
\u2022 1 new senior engineer hire
\u2022 \u2713 Expected per hiring plan

**Recommendation**: Document Sales variance in review notes. This is material and will be questioned.`:n.includes("journal")||n.includes("je")||n.includes("entry")?`Good news \u2014 your journal entry looks ready to go! Here's the full check:

**\u2713 Balance Check: PASSED**
\u2022 Total Debits: $253,625
\u2022 Total Credits: $253,625
\u2022 Difference: $0.00 \u2014 perfectly balanced!

**\u2713 Mapping Validation: Complete**
\u2022 9 unique GL accounts used
\u2022 All department codes are valid

**\u2713 Reference Data:**
\u2022 JE ID: PR-AUTO-2025-11-27
\u2022 Transaction Date: 2025-11-27
\u2022 Period: November 2025

**You're clear to export!** \u2705

**Next Steps:**
1. Click "Export" to download the CSV
2. Import into your accounting system
3. Post and reconcile

Let me know if you need me to double-check anything before you export!`:`Great question! I'm Ada, and I'm here to help with your payroll analysis.

Here's what I can help you with:

\u2022 **\u{1F50D} Diagnostics** \u2014 Check data quality and completeness
\u2022 **\u{1F4A1} Insights** \u2014 Key findings for executive review  
\u2022 **\u{1F4CA} Variance Analysis** \u2014 Period-over-period changes
\u2022 **\u{1F4CB} JE Readiness** \u2014 Validate journal entries before export

Try clicking one of the quick action buttons above, or just ask me something specific like:
\u2022 "What's driving the variance this period?"
\u2022 "Is my data ready for the CFO?"
\u2022 "Summarize the department breakdown"

I'm reading your actual spreadsheet data, so I can give you specific answers!`}var vn=ft.ADA_IMAGE_URL;async function Vt(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async o=>{let a=o.workbook.worksheets.getItemOrNullObject(e);a.load("isNullObject, name"),await o.sync();let s;a.isNullObject?(s=o.workbook.worksheets.add(e),await o.sync(),await hn(o,s,t,n)):(s=a,await hn(o,s,t,n)),s.activate(),s.getRange("A1").select(),await o.sync()})}catch(o){console.error(`Error activating homepage sheet ${e}:`,o)}}async function hn(e,t,n,o){try{let i=t.getUsedRangeOrNullObject();i.load("isNullObject"),await e.sync(),i.isNullObject||(i.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let a=[[n,""],[o,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=a;let l=t.getRange("A1:Z100");l.format.fill.color="#0f0f0f";let c=t.getRange("A1");c.format.font.bold=!0,c.format.font.size=36,c.format.font.color="#ffffff",c.format.font.name="Segoe UI Light",c.format.verticalAlignment="Center";let r=t.getRange("A2");r.format.font.size=14,r.format.font.color="#a0a0a0",r.format.font.name="Segoe UI",r.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var yn={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function jt(e){return yn[e]||yn["module-selector"]}function bn(){Ht();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${vn}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",Io),e}function Ht(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function Io(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${vn}" alt="Ada" />
                <h2 class="pf-ada-modal__title">Meet Ada</h2>
                <p class="pf-ada-modal__subtitle">Your AI-powered assistant</p>
            </div>
            <div class="pf-ada-modal__body">
                <div class="pf-ada-modal__coming-soon">
                    <div class="pf-ada-modal__coming-soon-icon">\u2728</div>
                    <p class="pf-ada-modal__coming-soon-text">Coming Soon!</p>
                    <p class="pf-ada-modal__coming-soon-desc">
                        Ada will help you navigate your workflows, answer questions, and provide insights about your data.
                    </p>
                </div>
                <div class="pf-ada-modal__features">
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">\u{1F4AC}</div>
                        <span class="pf-ada-modal__feature-text">Ask questions about your data</span>
                    </div>
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">\u{1F4CA}</div>
                        <span class="pf-ada-modal__feature-text">Get insights and trend analysis</span>
                    </div>
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">\u{1F50D}</div>
                        <span class="pf-ada-modal__feature-text">Troubleshoot issues quickly</span>
                    </div>
                </div>
            </div>
            <div class="pf-ada-modal__footer">
                <span class="pf-ada-modal__powered-by">Powered by ChatGPT</span>
            </div>
        </div>
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",Ft),t.addEventListener("click",a=>{a.target===t&&Ft()});let o=a=>{a.key==="Escape"&&(Ft(),document.removeEventListener("keydown",o))};document.addEventListener("keydown",o)}function Ft(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var No=["January","February","March","April","May","June","July","August","September","October","November","December"],wn=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],Oo=["Su","Mo","Tu","We","Th","Fr","Sa"],Fe=null;function En(e,t={}){let n=document.getElementById(e);if(!n)return;let{onChange:o=null,minDate:a=null,maxDate:s=null,readonly:l=!1}=t,c=n.closest(".pf-datepicker-wrapper");c||(c=document.createElement("div"),c.className="pf-datepicker-wrapper",n.parentNode.insertBefore(c,n),c.appendChild(n)),n.type="text",n.placeholder="YYYY-MM-DD or click calendar",n.classList.add("pf-datepicker-input");let r=n.value?mt(n.value):null,i=r?new Date(r):new Date;r&&(n.value=Ut(r));let f=document.createElement("span");f.className="pf-datepicker-icon",f.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>',c.appendChild(f);let d=document.createElement("div");d.className="pf-datepicker-dropdown",d.id=`${e}-dropdown`,c.appendChild(d);function u(){var E,_,$,M,S,A;let y=i.getFullYear(),g=i.getMonth();d.innerHTML=`
            <div class="pf-datepicker-header">
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev-year" title="Previous Year">\xAB</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev" title="Previous Month">\u2039</button>
                <span class="pf-datepicker-title">${No[g]} ${y}</span>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next" title="Next Month">\u203A</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next-year" title="Next Year">\xBB</button>
            </div>
            <div class="pf-datepicker-weekdays">
                ${Oo.map(m=>`<span>${m}</span>`).join("")}
            </div>
            <div class="pf-datepicker-days">
                ${p(y,g,r)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-today">Today</button>
                <button type="button" class="pf-datepicker-clear">Clear</button>
            </div>
        `,(E=d.querySelector(".pf-datepicker-prev-year"))==null||E.addEventListener("mousedown",m=>{m.preventDefault(),m.stopPropagation(),i.setFullYear(i.getFullYear()-1),u()}),(_=d.querySelector(".pf-datepicker-prev"))==null||_.addEventListener("mousedown",m=>{m.preventDefault(),m.stopPropagation(),i.setMonth(i.getMonth()-1),u()}),($=d.querySelector(".pf-datepicker-next"))==null||$.addEventListener("mousedown",m=>{m.preventDefault(),m.stopPropagation(),i.setMonth(i.getMonth()+1),u()}),(M=d.querySelector(".pf-datepicker-next-year"))==null||M.addEventListener("mousedown",m=>{m.preventDefault(),m.stopPropagation(),i.setFullYear(i.getFullYear()+1),u()}),d.querySelectorAll(".pf-datepicker-day:not(.disabled)").forEach(m=>{m.addEventListener("mousedown",D=>{D.preventDefault(),D.stopPropagation();let P=parseInt(m.dataset.day),H=parseInt(m.dataset.month),Q=parseInt(m.dataset.year);h(new Date(Q,H,P))})}),(S=d.querySelector(".pf-datepicker-today"))==null||S.addEventListener("mousedown",m=>{m.preventDefault(),m.stopPropagation(),h(new Date)}),(A=d.querySelector(".pf-datepicker-clear"))==null||A.addEventListener("mousedown",m=>{m.preventDefault(),m.stopPropagation(),h(null)})}function p(y,g,E){let _=new Date(y,g,1).getDay(),$=new Date(y,g+1,0).getDate(),M=new Date(y,g,0).getDate(),S=new Date;S.setHours(0,0,0,0);let A="";for(let P=_-1;P>=0;P--){let H=M-P,Q=g===0?11:g-1,J=g===0?y-1:y;A+=`<span class="pf-datepicker-day other-month" data-day="${H}" data-month="${Q}" data-year="${J}">${H}</span>`}for(let P=1;P<=$;P++){let H=new Date(y,g,P),Q=H.getTime()===S.getTime(),J=E&&H.getTime()===E.getTime(),F="pf-datepicker-day";Q&&(F+=" today"),J&&(F+=" selected"),a&&H<a&&(F+=" disabled"),s&&H>s&&(F+=" disabled"),A+=`<span class="${F}" data-day="${P}" data-month="${g}" data-year="${y}">${P}</span>`}let D=Math.ceil((_+$)/7)*7-(_+$);for(let P=1;P<=D;P++){let H=g===11?0:g+1,Q=g===11?y+1:y;A+=`<span class="pf-datepicker-day other-month" data-day="${P}" data-month="${H}" data-year="${Q}">${P}</span>`}return A}function h(y){r=y,y?(n.value=Ut(y),n.dataset.value=Ke(y),i=new Date(y)):(n.value="",n.dataset.value=""),w(),o&&o(y?Ke(y):""),n.dispatchEvent(new Event("change",{bubbles:!0}))}function v(){if(!l){if(Fe&&Fe!==e){let y=document.getElementById(`${Fe}-dropdown`);y==null||y.classList.remove("open")}Fe=e,u(),d.classList.add("open"),c.classList.add("open")}}function w(){d.classList.remove("open"),c.classList.remove("open"),Fe===e&&(Fe=null)}return n.addEventListener("blur",y=>{if(d.classList.contains("open"))return;let g=n.value.trim();if(!g)return;let E=mt(g);E&&(r=E,n.value=Ut(E),n.dataset.value=Ke(E),i=new Date(E),o&&o(Ke(E)),n.dispatchEvent(new Event("change",{bubbles:!0})))}),n.addEventListener("keydown",y=>{if(y.key==="Enter"){y.preventDefault();let g=n.value.trim(),E=mt(g);E&&h(E),w()}}),n.addEventListener("click",y=>{y.stopPropagation(),d.classList.contains("open")||v()}),f.addEventListener("click",y=>{y.stopPropagation(),d.classList.contains("open")?w():v()}),document.addEventListener("click",y=>{c.contains(y.target)||w()}),d.addEventListener("click",y=>{y.stopPropagation()}),document.addEventListener("keydown",y=>{y.key==="Escape"&&w()}),{getValue:()=>r?Ke(r):"",setValue:y=>{let g=mt(y);h(g)},open:v,close:w}}function mt(e){if(!e)return null;if(/^\d{4}-\d{2}-\d{2}$/.test(e)){let[o,a,s]=e.split("-").map(Number);return new Date(o,a-1,s)}let t=e.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);if(t){let o=wn.findIndex(a=>a.toLowerCase()===t[1].toLowerCase().substring(0,3));if(o>=0)return new Date(parseInt(t[3]),o,parseInt(t[2]))}if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(e)){let[o,a,s]=e.split("/").map(Number);return new Date(s,o-1,a)}let n=new Date(e);return isNaN(n.getTime())?null:n}function Ut(e){return e?`${wn[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`:""}function Ke(e){if(!e)return"";let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),o=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${o}`}var Cn=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M15 21v-8a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v8" />
        <path
            d="M3 10a2 2 0 0 1 .709-1.528l7-6a2 2 0 0 1 2.582 0l7 6A2 2 0 0 1 21 10v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"
        />
    </svg>
`.trim(),kn=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <rect width="7" height="7" x="3" y="3" rx="1" />
        <rect width="7" height="7" x="14" y="3" rx="1" />
        <rect width="7" height="7" x="14" y="14" rx="1" />
        <rect width="7" height="7" x="3" y="14" rx="1" />
    </svg>
`.trim(),Rn=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="1"/>
        <circle cx="12" cy="5" r="1"/>
        <circle cx="12" cy="19" r="1"/>
    </svg>
`.trim(),gt=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M12 3v18"/>
        <rect width="18" height="18" x="3" y="3" rx="2"/>
        <path d="M3 9h18"/>
        <path d="M3 15h18"/>
    </svg>
`.trim(),Sn=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/>
        <circle cx="9" cy="7" r="4"/>
        <path d="M22 21v-2a4 4 0 0 0-3-3.87"/>
        <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
    </svg>
`.trim(),xn=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M4 19.5v-15A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1 0-5H20"/>
        <path d="M8 7h6"/>
        <path d="M8 11h8"/>
    </svg>
`.trim(),To={config:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="3" />
            <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.82 2.82l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.82-2.82l.06-.06A1.65 1.65 0 0 0 3 15a1.65 1.65 0 0 0-1.51-1H1a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 3 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.82-2.82l.06.06A1.65 1.65 0 0 0 9 3.6a1.65 1.65 0 0 0 1-1.51V2a2 2 0 0 1 4 0v.09A1.65 1.65 0 0 0 15 3.6a1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.82 2.82l-.06.06A1.65 1.65 0 0 0 21 9c0 .3.09.58.24.82.17.28.43.51.76.68.21.1.44.18.68.19H23a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1Z" />
        </svg>
    `,import:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M12 3v14" />
            <path d="m7 13 5 5 5-5" />
            <path d="M5 21h14" />
        </svg>
    `,headcount:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
            <circle cx="9" cy="7" r="4" />
            <path d="M22 21v-2a4 4 0 0 0-3-3.87" />
            <path d="M16 3.13a4 4 0 0 1 0 7.75" />
        </svg>
    `,validate:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M20 6 9 17l-5-5" />
        </svg>
    `,review:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 12h3l2-5 4 10 2-5h5" />
        </svg>
    `,journal:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M12 7c-3-1-6-1-9 0v12c3-1 6-1 9 0 3-1 6-1 9 0V7c-3-1-6-1-9 0Z" />
            <path d="M12 7v12" />
        </svg>
    `,archive:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <rect x="3" y="3" width="18" height="4" rx="1" />
            <path d="M5 7v11a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7" />
            <path d="M10 12h4" />
        </svg>
    `};function _n(e){return e&&To[e]||""}var Gt=`
    <svg
        class="pf-icon pf-lock-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
        <path d="M7 11V7a5 5 0 0 1 10 0" />
    </svg>
`.trim(),zt=`
    <svg
        class="pf-icon pf-lock-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
        <path d="M7 11V7a5 5 0 0 1 10 0v4" />
        <path d="M12 15v2" />
    </svg>
`.trim(),ht=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M5 12l4 4 10-10" />
    </svg>
`.trim(),yt=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <rect x="4" y="3" width="16" height="18" rx="2" />
        <rect x="8" y="7" width="8" height="3" />
        <path d="M8 14h.01" />
        <path d="M12 14h.01" />
        <path d="M16 14h.01" />
        <path d="M8 17h.01" />
        <path d="M12 17h.01" />
        <path d="M16 17h.01" />
    </svg>
`.trim(),cs=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M18 5.5 20.5 8 16 12.5 13.5 10 18 5.5Z" />
        <path d="m12 11 6-6" />
        <path d="M3 22 12 13" />
        <path d="m3 18 4 4" />
        <path d="m11 11 3 3" />
    </svg>
`.trim(),vt=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
        <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 1 0 7.07 7.07l1.71-1.71" />
    </svg>
`.trim(),Dn=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
        <path d="M7 10l5 5 5-5" />
        <path d="M12 15V3" />
    </svg>
`.trim(),An=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <circle cx="12" cy="12" r="10" />
        <path d="m15 9-6 6" />
        <path d="m9 9 6 6" />
    </svg>
`.trim(),Pn=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m12 19-7-7 7-7" />
        <path d="M19 12H5" />
    </svg>
`.trim(),$n=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" />
        <path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7" />
        <path d="M7 3v4a1 1 0 0 0 1 1h7" />
    </svg>
`.trim(),In=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m12 5 7 7-7 7" />
        <path d="M5 12h14" />
    </svg>
`.trim(),ds=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <path d="m9 12 2 2 4-4"/>
    </svg>
`.trim(),ps=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <line x1="12" x2="12" y1="8" y2="12"/>
        <line x1="12" x2="12.01" y1="16" y2="16"/>
    </svg>
`.trim(),us=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>
        <path d="M12 9v4"/>
        <path d="M12 17h.01"/>
    </svg>
`.trim(),fs=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <path d="M12 16v-4"/>
        <path d="M12 8h.01"/>
    </svg>
`.trim(),ms=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M21.801 10A10 10 0 1 1 17 3.335"/>
        <path d="m9 11 3 3L22 4"/>
    </svg>
`.trim(),gs=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <path d="m15 9-6 6"/>
        <path d="m9 9 6 6"/>
    </svg>
`.trim(),hs=`
    <svg
        class="pf-icon pf-mismatch-icon-svg"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m18 9-6-6-6 6"/>
        <path d="M12 3v14"/>
        <path d="M5 21h14"/>
    </svg>
`.trim(),ys=`
    <svg
        class="pf-icon pf-mismatch-icon-svg"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m6 15 6 6 6-6"/>
        <path d="M12 21V7"/>
        <path d="M5 3h14"/>
    </svg>
`.trim(),Xe=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/>
        <path d="M21 3v5h-5"/>
        <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/>
        <path d="M3 21v-5h5"/>
    </svg>
`.trim(),Nn=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M3 6h18"/>
        <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/>
        <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>
        <line x1="10" x2="10" y1="11" y2="17"/>
        <line x1="14" x2="14" y1="11" y2="17"/>
    </svg>
`.trim();function Qe(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function Jt(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function De({textareaId:e,value:t,permanentId:n,isPermanent:o,hintId:a,saveButtonId:s,isSaved:l=!1,placeholder:c="Enter notes here..."}){let r=o?zt:Gt,i=s?`<button type="button" class="pf-action-toggle pf-save-btn ${l?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${$n}</button>`:"",f=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${o?"is-locked":""}" id="${n}" aria-pressed="${o}" title="Lock notes (retain after archive)">${r}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Qe(c)}">${Qe(t||"")}</textarea>
                ${a?`<p class="pf-signoff-hint" id="${a}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?Jt(f,"Lock"):""}
                ${s?Jt(i,"Save"):""}
            </div>
        </article>
    `}function Ae({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:o,isComplete:a,saveButtonId:s,isSaved:l=!1,completeButtonId:c,subtext:r="Sign-off below. Click checkmark icon. Done."}){let i=`<button type="button" class="pf-action-toggle ${a?"is-active":""}" id="${c}" aria-pressed="${!!a}" title="Mark step complete">${ht}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${Qe(r)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${Qe(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${Qe(o)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${Jt(i,"Done")}
            </div>
        </article>
    `}function Ze(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?zt:Gt)}function Ve(e,t){e&&e.classList.toggle("is-saved",t)}function Wt(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(o=>{let a=o.getAttribute("data-save-input"),s=document.getElementById(a);if(!s)return;let l=()=>{Ve(o,!1)};s.addEventListener("input",l),n.push(()=>s.removeEventListener("input",l))}),()=>n.forEach(o=>o())}function On(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function Tn(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var Yt={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},qt={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function Ln(e){e.format.fill.color=Yt.fillColor,e.format.font.color=Yt.fontColor,e.format.font.bold=Yt.bold}function bt(e,t,n,o=!1){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[o?qt.currencyWithNegative:qt.currency]]}function Mn(e,t,n,o=qt.date){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[o]]}var _t="payroll-recorder";var Ie="Payroll Recorder",Fs=N.CONFIG||"SS_PF_Config",Kt=["SS_PF_Config"];var Lo="Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel. Every run follows the same guidance so you stay audit-ready.",we=pt.map(({id:e,title:t})=>({id:e,title:t})),B={TYPE:0,FIELD:1,VALUE:2,PERMANENT:3,TITLE:-1},Mo="Run Settings";var Bn="N";var Bo="PR_JE_Debit_Total",Fo="PR_JE_Credit_Total",Vo="PR_JE_Difference",Pe={0:{note:"PR_Notes_Config",reviewer:"PR_Reviewer_Config",signOff:"PR_SignOff_Config"},1:{note:"PR_Notes_Import",reviewer:"PR_Reviewer_Import",signOff:"PR_SignOff_Import"},2:{note:"PR_Notes_Headcount",reviewer:"PR_Reviewer_Headcount",signOff:"PR_SignOff_Headcount"},3:{note:"PR_Notes_Validate",reviewer:"PR_Reviewer_Validate",signOff:"PR_SignOff_Validate"},4:{note:"PR_Notes_Review",reviewer:"PR_Reviewer_Review",signOff:"PR_SignOff_Review"},5:{note:"PR_Notes_JE",reviewer:"PR_Reviewer_JE",signOff:"PR_SignOff_JE"},6:{note:"PR_Notes_Archive",reviewer:"PR_Reviewer_Archive",signOff:"PR_SignOff_Archive"}},de={0:"PR_Complete_Config",1:"PR_Complete_Import",2:"PR_Complete_Headcount",3:"PR_Complete_Validate",4:"PR_Complete_Review",5:"PR_Complete_JE",6:"PR_Complete_Archive"},jo={1:N.DATA,2:N.DATA_CLEAN,3:N.DATA_CLEAN,4:N.EXPENSE_REVIEW,5:N.JE_DRAFT},tt="PR_Reviewer",Kn="PR_Payroll_Provider",wt="User opted to skip the headcount review this period.",re={statusText:"",focusedIndex:0,activeView:"home",activeStepId:null,stepStatuses:we.reduce((e,t)=>({...e,[t.id]:"pending"}),{})},Y={loaded:!1,values:{},permanents:{},overrides:{accountingPeriod:!1,jeId:!1}},je=new Map,Et=null,ot=["PR_Payroll_Date","Payroll Date (YYYY-MM-DD)","Payroll_Date","Payroll Date","Payroll_Date_(YYYY-MM-DD)"],z={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},departments:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},St=null,G={loading:!1,lastError:null,prDataTotal:null,cleanTotal:null,reconDifference:null,bankAmount:"",bankDifference:null,plugEnabled:!1},Se={loading:!1,lastError:null,periods:[],copilotResponse:"",completenessCheck:{currentPeriod:null,historicalPeriods:null}},j={debitTotal:null,creditTotal:null,difference:null,cleanTotal:null,loading:!1,lastError:null};async function Ho(){if(console.log("Completeness Check - Starting..."),!pe()){console.log("Completeness Check - Excel runtime not available");return}try{await Excel.run(async e=>{var a,s,l,c;let t=e.workbook.worksheets.getItemOrNullObject(N.DATA_CLEAN),n=e.workbook.worksheets.getItemOrNullObject(N.ARCHIVE_SUMMARY);t.load("isNullObject"),n.load("isNullObject"),await e.sync();let o={currentPeriod:null,historicalPeriods:null};if(!t.isNullObject){let r=t.getUsedRangeOrNullObject();if(r.load("values"),await e.sync(),!r.isNullObject&&r.values&&r.values.length>1){let i=(r.values[0]||[]).map(u=>String(u||"").toLowerCase().trim()),f=i.findIndex(u=>u.includes("amount")),d=f>=0?f:i.findIndex(u=>u==="total"||u==="all-in"||u==="allin"||u==="all-in total"||u==="gross"||u==="total pay");if(console.log("Completeness Check - PR_Data_Clean headers:",i),console.log("Completeness Check - Amount column index:",f,"Total column index:",d),d>=0){let p=r.values.slice(1).reduce((w,y)=>w+(Number(y[d])||0),0),h=((l=(s=(a=Se.periods)==null?void 0:a[0])==null?void 0:s.summary)==null?void 0:l.total)||0;console.log("Completeness Check - PR_Data_Clean sum:",p,"Current period total:",h);let v=Math.abs(p-h)<1;o.currentPeriod={match:v,prDataClean:p,currentTotal:h}}else console.warn("Completeness Check - No amount/total column found in PR_Data_Clean")}}if(!n.isNullObject){let r=n.getUsedRangeOrNullObject();if(r.load("values"),await e.sync(),!r.isNullObject&&r.values&&r.values.length>1){let i=(r.values[0]||[]).map(p=>String(p||"").toLowerCase().trim()),f=i.findIndex(p=>p.includes("pay period")||p.includes("payroll date")||p==="date"||p==="period"||p.includes("period")),d=i.findIndex(p=>p.includes("amount")),u=d>=0?d:i.findIndex(p=>p==="total"||p==="all-in"||p==="allin"||p==="all-in total"||p==="total payroll"||p.includes("total"));if(console.log("Completeness Check - PR_Archive_Summary headers:",i),console.log("Completeness Check - Date column index:",f,"Total column index:",u),u>=0&&f>=0){let p=r.values.slice(1),h=(Se.periods||[]).slice(1,6);console.log("Completeness Check - Looking for periods:",h.map(S=>S.key||S.label));let v=new Map;for(let S of p){let A=S[f],m=jn(A);if(m){let D=Number(S[u])||0,P=v.get(m)||0;v.set(m,P+D)}}console.log("Completeness Check - Archive lookup keys:",Array.from(v.keys())),console.log("Completeness Check - Archive lookup values:",Array.from(v.entries()));let w=0,y=0,g=0,E=[];for(let S of h){let A=S.key||S.label||"",m=jn(A),D=((c=S.summary)==null?void 0:c.total)||0;y+=D;let P=v.get(m);P!==void 0?(w+=P,g++,E.push({period:A,calculated:D,archive:P,match:Math.abs(D-P)<1})):(console.warn(`Completeness Check - Period ${A} (normalized: ${m}) not found in archive`),E.push({period:A,calculated:D,archive:null,match:!1}))}console.log("Completeness Check - Period details:",E),console.log("Completeness Check - Matched",g,"of",h.length,"periods"),console.log("Completeness Check - Archive sum:",w,"Periods sum:",y);let _=g===h.length&&h.length>0,$=Math.abs(w-y)<1,M=_&&$;o.historicalPeriods={match:M,archiveSum:w,periodsSum:y,matchedCount:g,totalPeriods:h.length,details:E}}else console.warn("Completeness Check - Missing date or total column in PR_Archive_Summary"),console.warn("  Date column index:",f,"Total column index:",u)}}Se.completenessCheck=o,console.log("Completeness Check - Results:",JSON.stringify(o))}),console.log("Completeness Check - Complete!")}catch(e){console.error("Payroll completeness check failed:",e)}}function Uo(){var v,w;let e=Se.completenessCheck||{},t=((v=Se.periods)==null?void 0:v.length)>0,n=y=>`$${Math.round(y||0).toLocaleString()}`,o=y=>{let g=Math.abs(y);return g<1?"\u2014":`${y>0?"+":"-"}$${Math.round(g).toLocaleString()}`},a=(y,g,E,_,$,M,S)=>{let A=(E||0)-($||0),m,D;S?(m='<span class="pf-complete-status pf-complete-status--pending">\u23F3</span>',D="pending"):M?(m='<span class="pf-complete-status pf-complete-status--pass">\u2713</span>',D="pass"):(m='<span class="pf-complete-status pf-complete-status--fail">\u2717</span>',D="fail");let P=S?"":`
            <div class="pf-complete-diff ${D}">
                ${o(A)}
            </div>
        `;return`
            <div class="pf-complete-row ${D}">
                <div class="pf-complete-header">
                    ${m}
                    <span class="pf-complete-label">${x(y)}</span>
                </div>
                ${S?`
                <div class="pf-complete-values">
                    <span class="pf-complete-pending-text">Click Run/Refresh to check</span>
                </div>
                `:`
                <div class="pf-complete-values">
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${x(g)}:</span>
                        <span class="pf-complete-amount">${n(E)}</span>
                    </div>
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${x(_)}:</span>
                        <span class="pf-complete-amount">${n($)}</span>
                    </div>
                </div>
                ${P}
                `}
            </div>
        `},s=e.currentPeriod,l=!t||s===null||s===void 0,c=a("Current Period","PR_Data_Clean Total",s==null?void 0:s.prDataClean,"Calculated Total",s==null?void 0:s.currentTotal,s==null?void 0:s.match,l),r=e.historicalPeriods,i=!t||r===null||r===void 0,f=(r==null?void 0:r.matchedCount)||0,d=(r==null?void 0:r.totalPeriods)||0,u=d>0?`Historical Periods (${f}/${d} matched)`:"Historical Periods",p=a(u,"PR_Archive_Summary (matched)",r==null?void 0:r.archiveSum,"Calculated Total",r==null?void 0:r.periodsSum,r==null?void 0:r.match,i),h="";return!i&&((w=r==null?void 0:r.details)==null?void 0:w.length)>0&&(h=`
            <div class="pf-complete-details-section">
                <div class="pf-complete-details-header">Period-by-Period Match</div>
                ${r.details.map(g=>{let E=g.archive===null?"\u26A0\uFE0F":g.match?"\u2713":"\u2717",_=g.archive!==null?n(g.archive):"Not found";return`
                <div class="pf-complete-detail-row">
                    <span class="pf-complete-detail-date">${x(g.period)}</span>
                    <span class="pf-complete-detail-icon">${E}</span>
                    <span class="pf-complete-detail-vals">${n(g.calculated)} vs ${_}</span>
                </div>
            `}).join("")}
            </div>
        `),`
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-completeness-card">
            <div class="pf-config-head">
                <h3>Data Completeness Check</h3>
                <p class="pf-config-subtext">Verify source data matches calculated totals</p>
            </div>
            <div class="pf-complete-container">
                ${c}
                ${p}
                ${h}
            </div>
        </article>
    `}function Go(e){switch(e){case 0:return{title:"Configuration",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Sets up the key parameters for your payroll review. Complete this before importing data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Key Fields</h4>
                        <ul>
                            <li><strong>Payroll Date</strong> \u2014 The period-end date for this payroll run</li>
                            <li><strong>Accounting Period</strong> \u2014 Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> \u2014 Reference number for your accounting system</li>
                            <li><strong>Provider Link</strong> \u2014 Quick access to your payroll provider portal</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>The accounting period and JE ID auto-generate based on your payroll date, but you can override them if needed.</p>
                    </div>
                `};case 1:return{title:"Import Payroll Data",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Gets your payroll data into the workbook. Pull a report from your payroll provider and paste it into PR_Data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Required Columns</h4>
                        <p>Your payroll export should include:</p>
                        <ul>
                            <li><strong>Employee Name</strong> \u2014 Full name (used to match against roster)</li>
                            <li><strong>Department</strong> \u2014 Cost center assignment</li>
                            <li><strong>Regular Earnings</strong> \u2014 Base pay for the period</li>
                            <li><strong>Overtime</strong> \u2014 OT pay (if applicable)</li>
                            <li><strong>Bonus/Commission</strong> \u2014 Variable compensation</li>
                            <li><strong>Benefits/Deductions</strong> \u2014 Employer portions</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Column headers don't need to match exactly\u2014the system is flexible with naming. Just make sure each field is present.</p>
                    </div>
                `};case 2:return{title:"Headcount Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Compares employee counts and department assignments between your roster and payroll data to catch discrepancies early.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Data Sources</h4>
                        <ul>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Your centralized employee list (Column A: Employee names)</li>
                            <li><strong>PR_Data</strong> \u2014 The payroll data you just imported (Employee column)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F50D} Employee Alignment Check</h4>
                        <p>The script compares names between SS_Employee_Roster and PR_Data to find:</p>
                        <ul>
                            <li><strong>In Roster, Missing from Payroll</strong> \u2014 Employees on roster but not in payroll (possible missed payment)</li>
                            <li><strong>In Payroll, Missing from Roster</strong> \u2014 Employees paid but not on roster (possible ghost employee or new hire)</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7); margin-top: 8px;">Names are matched using fuzzy logic to handle minor variations.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F3E2} Department Alignment Check</h4>
                        <p>For employees appearing in both sources, the script compares the "Department" column:</p>
                        <ul>
                            <li>Flags employees where roster department \u2260 payroll department</li>
                            <li>Mismatches affect GL coding and cost center reporting</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>If discrepancies are expected (e.g., contractors, temp workers), you can skip this check and add a note explaining why. The note is required if you skip.</p>
                    </div>
                `};case 3:return{title:"Payroll Validation",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Validates that your payroll totals match what was actually paid from the bank.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Reconciliation Check</h4>
                        <ul>
                            <li><strong>PR_Data Total</strong> \u2014 Sum of all payroll from your import</li>
                            <li><strong>Clean Total</strong> \u2014 Processed total after expense mapping</li>
                            <li><strong>Bank Amount</strong> \u2014 What actually left the bank account</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Common Differences</h4>
                        <ul>
                            <li><strong>Timing</strong> \u2014 Direct deposits vs check clearing dates</li>
                            <li><strong>Tax payments</strong> \u2014 May be separate from net pay</li>
                            <li><strong>Benefits</strong> \u2014 Some deductions paid to vendors</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Small differences ($0.01-$1.00) are usually rounding. Use the plug feature to resolve them.</p>
                    </div>
                `};case 4:return{title:"Expense Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Generates an executive-ready payroll expense summary for CFO review, with period comparisons and trend analysis.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4C2} Data Sources</h4>
                        <ul>
                            <li><strong>PR_Data_Clean</strong> \u2014 Current period payroll data (cleaned and categorized)</li>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Department assignments and employee details</li>
                            <li><strong>PR_Archive_Summary</strong> \u2014 Historical payroll data for trend analysis</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4B0} How Amounts Are Calculated</h4>
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
                                <td style="padding: 6px 0;">Burden \xF7 All-In Total (typically 10-18%)</td>
                            </tr>
                        </table>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Report Sections</h4>
                        <ul>
                            <li><strong>Executive Summary</strong> \u2014 Current vs prior period comparison (frozen at top)</li>
                            <li><strong>Department Breakdown</strong> \u2014 Cost allocation by cost center</li>
                            <li><strong>Historical Context</strong> \u2014 Where current metrics fall within historical ranges</li>
                            <li><strong>Period Trends</strong> \u2014 6-period trend chart for Total, Fixed, Variable, Burden, and Headcount</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4C8} Historical Context Visualization</h4>
                        <p>The spectrum bars show where your current period falls relative to your historical min/max:</p>
                        <p style="font-family: Consolas, monospace; color: #6366f1; margin: 8px 0;">\u2591\u2591\u2591\u2591\u2591\u2591\u2591\u25CF\u2591\u2591\u2591\u2591\u2591\u2591\u2591\u2591</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7);">The dot (\u25CF) indicates current position. Left = Low, Right = High.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Review Tips</h4>
                        <ul>
                            <li>Compare <strong>Burden Rate</strong> \u2014 Should be consistent period-to-period (10-18% typical)</li>
                            <li>Watch <strong>Variable Salary</strong> spikes \u2014 May indicate commission/bonus timing</li>
                            <li>Verify <strong>Headcount changes</strong> \u2014 Should align with HR records</li>
                            <li>Flag variances <strong>> 10%</strong> from prior period for follow-up</li>
                        </ul>
                    </div>
                `};case 5:return{title:"Journal Entry",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Generates a balanced journal entry from your payroll data, ready for upload to your accounting system.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4DD} How the JE Works</h4>
                        <p>Maps payroll categories to GL accounts:</p>
                        <ul>
                            <li><strong>Expenses</strong> \u2192 Debits to departmental expense accounts</li>
                            <li><strong>Liabilities</strong> \u2192 Credits to payable accounts</li>
                            <li><strong>Cash</strong> \u2192 Credit to bank account</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u2705 Validation Checks</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> \u2014 Entry must balance</li>
                            <li><strong>All accounts mapped</strong> \u2014 No unassigned categories</li>
                            <li><strong>Totals match</strong> \u2014 JE ties to PR_Data</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Review the draft in PR_JE_Draft before exporting to catch any mapping errors.</p>
                    </div>
                `};case 6:return{title:"Archive & Clear",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Creates a backup of your completed payroll run, then resets the workbook so you're ready for the next pay period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4C1} Step 1: Create Backup</h4>
                        <p>A new workbook opens containing all your payroll tabs. You'll choose where to save it on your computer or shared drive.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Tip: Use a consistent naming convention like "Payroll_Archive_2024-01-15"</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Step 2: Update History</h4>
                        <p>The current period's totals are saved to PR_Archive_Summary. This powers the trend charts and completeness checks for future periods.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Keeps 5 periods of history \u2014 oldest is removed automatically</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F9F9} Step 3: Clear Working Data</h4>
                        <p>Data is cleared from the working sheets:</p>
                        <ul>
                            <li>PR_Data (raw import)</li>
                            <li>PR_Data_Clean (processed data)</li>
                            <li>PR_Expense_Review (summary & charts)</li>
                            <li>PR_JE_Draft (journal entry lines)</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Headers are preserved \u2014 only data rows are cleared</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F504} Step 4: Reset for Next Period</h4>
                        <ul>
                            <li>Payroll Date, Accounting Period, JE ID cleared</li>
                            <li>All sign-offs and completion flags reset</li>
                            <li>Notes cleared (unless you locked them with \u{1F512})</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Before You Archive</h4>
                        <ul>
                            <li>\u2713 JE uploaded to your accounting system</li>
                            <li>\u2713 All review steps signed off</li>
                            <li>\u2713 Lock any notes you want to keep</li>
                        </ul>
                    </div>
                `};default:return{title:"Payroll Recorder",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F44B} Welcome to Payroll Recorder</h4>
                        <p>This module helps you normalize payroll exports, enforce controls, and prep journal entries.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Workflow Overview</h4>
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
                        <p>Click a step card to get started, or tap the <strong>\u24D8</strong> button on any step for detailed guidance.</p>
                    </div>
                `}}}pn(()=>zo());async function zo(){try{await Jo(),await eo();let e=jt(_t);await Vt(e.sheetName,e.title,e.subtitle),fe()}catch(e){throw console.error("[Payroll] Module initialization failed:",e),e}}async function Jo(){try{await Mt(_t),console.log(`[Payroll] Tab visibility applied for ${_t}`)}catch(e){console.warn("[Payroll] Could not apply tab visibility:",e)}}function fe(){var r;let e=document.body;if(!e)return;let t=re.focusedIndex<=0?"disabled":"",n=re.focusedIndex>=we.length-1?"disabled":"",o=re.activeView==="config",a=re.activeView==="step",s=!o&&!a,l=o?qo():a?na(re.activeStepId):Yo();e.innerHTML=`
        <div class="pf-root">
            ${Wo(t,n)}
            ${l}
            ${aa()}
        </div>
    `;let c=document.getElementById("pf-info-fab-payroll");if(s)c&&c.remove();else if((r=window.PrairieForge)!=null&&r.mountInfoFab){let i=Go(re.activeStepId);PrairieForge.mountInfoFab({title:i.title,content:i.content,buttonId:"pf-info-fab-payroll"})}if(ra(),o)da();else if(a)try{pa(re.activeStepId)}catch(i){console.warn("Payroll Recorder: failed to bind step interactions",i)}else ca();ua(),s?bn():Ht()}function Wo(e,t){let n=I("SS_Company_Name")||"your company";return`
        <div class="pf-brand-float" aria-hidden="true">
            <span class="pf-brand-wave"></span>
        </div>
        <header class="pf-banner">
            <div class="pf-nav-bar">
                <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                    ${Pn}
                    <span class="sr-only">Previous step</span>
                </button>
                <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                    ${Cn}
                    <span class="sr-only">Module Home</span>
                </button>
                <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                    ${kn}
                    <span class="sr-only">Module Selector</span>
                </button>
                <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                    ${In}
                    <span class="sr-only">Next step</span>
                </button>
                <span class="pf-nav-divider"></span>
                <div class="pf-quick-access-wrapper">
                    <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                        ${Rn}
                        <span class="sr-only">Quick Access Menu</span>
                    </button>
                    <div id="quick-access-dropdown" class="pf-quick-dropdown hidden">
                        <div class="pf-quick-dropdown-header">Quick Access</div>
                        <button id="nav-roster" class="pf-quick-item pf-clickable" type="button">
                            ${Sn}
                            <span>Employee Roster</span>
                        </button>
                        <button id="nav-accounts" class="pf-quick-item pf-clickable" type="button">
                            ${xn}
                            <span>Chart of Accounts</span>
                        </button>
                        <button id="nav-expense-map" class="pf-quick-item pf-clickable" type="button">
                            ${gt}
                            <span>PR Mapping</span>
                </button>
                    </div>
                </div>
            </div>
        </header>
    `}function Yo(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">Payroll Recorder</h2>
            <p class="pf-hero-copy">${Lo}</p>
            <p class="pf-hero-hint">${x(re.statusText||"")}</p>
        </section>
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${we.map((e,t)=>oa(e,t)).join("")}
            </div>
        </section>
    `}function qo(){if(!Y.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=Pe[0],t=be(Dt()),n=be(I("PR_Accounting_Period")),o=I("PR_Journal_Entry_ID"),a=I("SS_Accounting_Software"),s=an(),l=I("SS_Company_Name"),c=I(tt)||$e(),r=e?I(e.note):"",i=e?xe(e.note):!1,f=(e?I(e.reviewer):"")||$e(),d=e?be(I(e.signOff)):"",u=!!(d||I(de[0]));return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every payroll run.</p>
            <p class="pf-hero-hint">${x(re.statusText||"")}</p>
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
                        <input type="text" id="config-user-name" value="${x(c)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Date</span>
                        <input type="date" id="config-payroll-date" value="${x(t)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${x(n)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-je-id" value="${x(o)}" placeholder="PR-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${x(l)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${x(s)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${x(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${e?De({textareaId:"config-notes",value:r,permanentId:"config-notes-permanent",isPermanent:i,hintId:"",saveButtonId:"config-notes-save"}):""}
            ${e?Ae({reviewerInputId:"config-reviewer-name",reviewerValue:f,signoffInputId:"config-signoff-date",signoffValue:d,isComplete:u,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"}):""}
        </section>
    `}function Ko(e){let t=_e(1),n=t?xe(t.note):!1,o=t?I(t.note):"",a=(t?I(t.reviewer):"")||$e(),s=t?be(I(t.signOff)):"",l=!!(s||I(de[1])),c=an();return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${x(e.title)}</h2>
            <p class="pf-hero-copy">Pull your payroll export from the provider and paste it into PR_Data.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Open your payroll provider, download the report, and paste into PR_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ye(c?`<a href="${x(c)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${vt}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${vt}</button>`,"Provider")}
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PR_Data sheet">${gt}</button>`,"PR_Data")}
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PR_Data to start over">${Nn}</button>`,"Clear")}
                </div>
            </article>
            ${t?`
                ${De({textareaId:"step-notes-1",value:o||"",permanentId:"step-notes-lock-1",isPermanent:n,saveButtonId:"step-notes-save-1"})}
                ${Ae({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:s,isComplete:l,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
            `:""}
        </section>
    `}function Xo(e){var H,Q,J,F,q,ne,ke,oe,Z,ee,ie,ae,te;let t=_e(2),n=t?I(t.note):"",o=t?xe(t.note):!1,a=(t?I(t.reviewer):"")||$e(),s=t?be(I(t.signOff)):"",l=!!(s||I(de[2])),c=Pt(),r=z.roster||{},i=z.departments||{},f=z.hasAnalyzed,d="";z.loading?d='<p class="pf-step-note">Analyzing roster and payroll data\u2026</p>':z.lastError&&(d=`<p class="pf-step-note">${x(z.lastError)}</p>`);let u=(ge,rt,$t)=>{let Oe=!f,Te;Oe?Te='<span class="pf-je-check-circle pf-je-circle--pending"></span>':$t?Te=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:Te=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let Le=f?` = ${rt}`:"";return`
            <div class="pf-je-check-row">
                ${Te}
                <span class="pf-je-check-desc-pill">${x(ge)}${Le}</span>
            </div>
        `},p=(H=r.difference)!=null?H:0,h=(Q=i.difference)!=null?Q:0,v=Array.isArray(r.mismatches)?r.mismatches.filter(Boolean):[],w=Array.isArray(i.mismatches)?i.mismatches.filter(Boolean):[],y=`
        ${u("SS_Employee_Roster count",(J=r.rosterCount)!=null?J:"\u2014",!0)}
        ${u("PR_Data employee count",(F=r.payrollCount)!=null?F:"\u2014",!0)}
        ${u("Difference",p,p===0)}
    `,g=(q=i.rosterCount)!=null?q:0,E=(ne=r.rosterCount)!=null?ne:0,_=(ke=r.payrollCount)!=null?ke:0,$=f&&g>0&&g<Math.max(E,_)?`<p class="pf-step-note pf-step-note--info">\u2139\uFE0F Only ${g} employees appear in both lists, so only those can be compared for department alignment.</p>`:"",M=`
        ${u("Expected departments",(oe=i.rosterCount)!=null?oe:"\u2014",!0)}
        ${u("PR_Data departments",(Z=i.payrollCount)!=null?Z:"\u2014",!0)}
        ${u("Difference",h,h===0)}
    `,S=v.filter(ge=>ge.type==="missing_from_payroll").length,A=v.filter(ge=>ge.type==="missing_from_roster").length,m=v.length>0?`Employee Mismatches (${S} missing from payroll, ${A} not in roster)`:"Employee Mismatches",D=v.length&&!z.skipAnalysis&&f&&((ie=(ee=window.PrairieForge)==null?void 0:ee.renderMismatchTiles)==null?void 0:ie.call(ee,{mismatches:v,label:m,sourceLabel:"Roster",targetLabel:"Payroll Data",escapeHtml:x}))||"",P=w.length&&!z.skipAnalysis&&f&&((te=(ae=window.PrairieForge)==null?void 0:ae.renderMismatchTiles)==null?void 0:te.call(ae,{mismatches:w,label:"Employees with Department Differences",formatter:ge=>({name:ge.employee||ge.name||"",source:`${ge.rosterDept||"\u2014"} \u2192 ${ge.payrollDept||"\u2014"}`,isMissingFromTarget:!0}),escapeHtml:x}))||"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">Headcount Review</h2>
            <p class="pf-hero-copy">Quick check to make sure your roster matches your payroll data.</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${z.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${An}
                    <span>Skip Analysis</span>
                </button>
            </div>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Headcount Check</h3>
                    <p class="pf-config-subtext">Compare employee roster against payroll data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="roster-run-btn" title="Run headcount analysis">${yt}</button>`,"Run")}
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="roster-refresh-btn" title="Refresh analysis">${Xe}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Employee Alignment</h3>
                    <p class="pf-config-subtext">Verify employees match between roster and payroll.</p>
                </div>
                ${d}
                <div class="pf-je-checks-container">
                    ${y}
                </div>
                ${D}
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Department Alignment</h3>
                    <p class="pf-config-subtext">Verify department assignments are consistent.</p>
                </div>
                ${$}
                <div class="pf-je-checks-container">
                    ${M}
                </div>
                ${P}
            </article>
            ${t?`
                ${De({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:o,hintId:c?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
                ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:a,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:l,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
            `:""}
        </section>
    `}function Qo(e){var $;let t=_e(3),n=t?I(t.note):"",o=(t?I(t.reviewer):"")||$e(),a=t?be(I(t.signOff)):"",s=G.loading?'<p class="pf-step-note">Preparing reconciliation data\u2026</p>':G.lastError?`<p class="pf-step-note">${x(G.lastError)}</p>`:"",l=!!(a||I(de[3])),c=G.prDataTotal!==null,r=G.prDataTotal,i=G.cleanTotal,f=($=G.reconDifference)!=null?$:r!=null&&i!=null?r-i:null,d=f!==null&&Math.abs(f)<.01,u=le(G.cleanTotal),p=G.bankDifference!=null?le(G.bankDifference):"---",h=G.bankDifference==null?"":Math.abs(G.bankDifference)<.5?"Difference within acceptable tolerance.":"Difference exceeds tolerance and should be resolved.",v=ao(G.bankAmount),w=(M,S,A)=>{let m=!c,D;return m?D='<span class="pf-je-check-circle pf-je-circle--pending"></span>':A?D=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:D=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${D}
                <span class="pf-je-check-desc-pill">${x(S)}</span>
            </div>
        `},y=c?le(r):"\u2014",g=c?le(i):"\u2014",E=c?le(f):"\u2014",_=`
        ${w("PR_Data Total",`PR_Data Total = ${y}`,!0)}
        ${w("PR_Data_Clean Total",`PR_Data_Clean Total = ${g}`,!0)}
        ${w("Difference",`Difference = ${E} (should be $0.00)`,d)}
    `;return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${x(e.title)}</h2>
            <p class="pf-hero-copy">Normalize your payroll data and verify totals match.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Validation</h3>
                    <p class="pf-config-subtext">Build PR_Data_Clean from your imported data and verify totals.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="validation-run-btn" title="Run reconciliation">${yt}</button>`,"Run")}
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="validation-refresh-btn" title="Refresh reconciliation">${Xe}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Data Reconciliation</h3>
                    <p class="pf-config-subtext">Verify PR_Data and PR_Data_Clean totals match.</p>
                </div>
                ${s}
                <div class="pf-je-checks-container">
                    ${_}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Bank Reconciliation</h3>
                    <p class="pf-config-subtext">Compare payroll total to the amount pulled from the bank.</p>
                </div>
                <div class="pf-config-grid pf-metric-grid">
                    <label class="pf-config-field">
                        <span>Cost per PR_Data_Clean</span>
                        <input id="bank-clean-total-value" type="text" class="pf-readonly-input pf-metric-value" value="${u}" readonly>
                    </label>
                    <label class="pf-config-field">
                        <span>Cost per Bank</span>
                        <input
                            type="text"
                            inputmode="decimal"
                            id="bank-amount-input"
                            class="pf-metric-input"
                            value="${x(v)}"
                            placeholder="0.00"
                            aria-label="Enter bank amount"
                        >
                    </label>
                    <label class="pf-config-field">
                        <span>Difference</span>
                        <input id="bank-diff-value" type="text" class="pf-readonly-input pf-metric-value" value="${p}" readonly>
                    </label>
                </div>
                <p class="pf-metric-hint" id="bank-diff-hint">${x(h)}</p>
            </article>
            ${t?`
                ${De({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:xe(t.note),saveButtonId:"step-notes-save-3"})}
            `:""}
            ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-3",signoffValue:a,isComplete:l,saveButtonId:"step-signoff-save-3",completeButtonId:"validation-signoff-toggle"})}
        </section>
    `}function Zo(e){let t=_e(4),n=t?I(t.note):"",o=(t?I(t.reviewer):"")||$e(),a=t?be(I(t.signOff)):"",s=!!(a||I(de[4])),l=Se.loading?'<p class="pf-step-note">Preparing executive summary\u2026</p>':Se.lastError?`<p class="pf-step-note">${x(Se.lastError)}</p>`:"",c=mn({id:"expense-review-copilot",body:"Want help analyzing your data? Just ask!",placeholder:"Where should I focus this pay period?",buttonLabel:"Ask CoPilot"});return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${x(e.title)}</h2>
            <p class="pf-hero-copy">${x(e.summary||"")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${l}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Perform Analysis</h3>
                    <p class="pf-config-subtext">Populate Expense Review and perform review.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ye(`<button type="button" class="pf-action-toggle" id="expense-run-btn" title="Run expense review analysis">${yt}</button>`,"Run")}
                    ${ye(`<button type="button" class="pf-action-toggle" id="expense-refresh-btn" title="Refresh expense data">${Xe}</button>`,"Refresh")}
                </div>
            </article>
            ${Uo()}
                ${c}
            ${t?`
            ${De({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:xe(t.note),saveButtonId:"step-notes-save-4"})}
            ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-4",signoffValue:a,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"expense-signoff-toggle"})}
            `:""}
        </section>
    `}function ea(e){var A,m,D;let t=_e(5),n=t?I(t.note):"",o=t?xe(t.note):!1,a=(t?I(t.reviewer):"")||$e(),s=t?be(I(t.signOff)):"",l=!!(s||I(de[5])),c=j.lastError?`<p class="pf-step-note">${x(j.lastError)}</p>`:"",r=j.debitTotal!==null,i=(A=j.debitTotal)!=null?A:0,f=(m=j.creditTotal)!=null?m:0,d=i-f,u=(D=j.cleanTotal)!=null?D:0,p=j.cleanTotal!==null,h=r,v=r&&p&&Math.abs(d-u)<.01,w=(P,H)=>{let Q=!r,J;return Q?J='<span class="pf-je-check-circle pf-je-circle--pending"></span>':H?J=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:J=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${J}
                <span class="pf-je-check-desc-pill">${x(P)}</span>
            </div>
        `},y=r?le(i):"\u2014",g=r?le(f):"\u2014",E=r?le(d):"\u2014",_=p?le(u):"\u2014",$=p?Math.abs(d-u):0,M=$>=.01?` (diff: ${le($)})`:"",S=`
        ${w(`Total Debits = ${y}`,h)}
        ${w(`Total Credits = ${g}`,h)}
        ${w(`Line Amount (Debit - Credit) = ${E}`,h)}
        ${w(`JE Total matches PR_Data_Clean (${_})${M}`,v)}
    `;return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${x(e.title)}</h2>
            <p class="pf-hero-copy">Generate the upload file to break down the bank feed line item.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Upload File</h3>
                    <p class="pf-config-subtext">Build the breakdown from PR_Data_Clean for your accounting system.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate from PR_Data_Clean">${gt}</button>`,"Generate")}
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation">${Xe}</button>`,"Refresh")}
                    ${ye(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export as CSV">${Dn}</button>`,"Export")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">Verify totals before uploading to your accounting system.</p>
                </div>
                ${c}
                <div class="pf-je-checks-container">
                    ${S}
                </div>
            </article>
            ${t?`
                ${De({textareaId:"step-notes-input",value:n||"",permanentId:"step-notes-permanent",isPermanent:o,saveButtonId:"step-notes-save-5"})}
                ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:s,isComplete:l,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
            `:""}
        </section>
    `}function ta(e){let t=we.filter(a=>a.id!==6).map(a=>({id:a.id,title:a.title,complete:ya(a.id)})),n=t.every(a=>a.complete),o=t.map(a=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${a.complete?"is-active":""}" aria-pressed="${a.complete}">
                        ${ht}
                    </span>
                    <div>
                        <h3>${x(a.title)}</h3>
                        <p class="pf-config-subtext">${a.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${x(e.title)}</h2>
            <p class="pf-hero-copy">${x(e.summary||"")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${o}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Create an archive of this module\u2019s sheets and clear work tabs.</p>
                </div>
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn" id="archive-run-btn" ${n?"":"disabled"}>Archive</button>
                </div>
            </article>
        </section>
    `}function na(e){let t=pt.find(_=>_.id===e)||{id:e!=null?e:"-",title:"Workflow Step",summary:"",description:"",checklist:[]};if(e===1)return Ko(t);if(e===2)return Xo(t);if(e===3)return Qo(t);if(e===4)return Zo(t);if(e===5)return ea(t);if(e===6)return ta(t);let n=!1,o=_e(e),a=o?I(o.note):"",s=o?xe(o.note):!1,l=(o?I(o.reviewer):"")||$e(),c=o?be(I(o.signOff)):"",r=o&&de[e]?!!(c||I(de[e])):!!c,i=(t.highlights||[]).map(_=>`
            <div class="pf-step-highlight">
                <span class="pf-step-highlight-label">${x(_.label)}</span>
                <span class="pf-step-highlight-detail">${x(_.detail)}</span>
            </div>
        `).join(""),f=(t.checklist||[]).map(_=>`<li>${x(_)}</li>`).join("")||"",d=n?"":t.description||"Detailed guidance will appear here.",u=[];!n&&t.ctaLabel&&u.push(`<button type="button" class="pf-pill-btn" id="step-action-btn">${x(t.ctaLabel)}</button>`),n||u.push('<button type="button" class="pf-pill-btn pf-pill-btn--ghost" id="step-back-btn">Back to Step List</button>');let p=u.length?`<div class="pf-pill-row pf-config-actions">${u.join("")}</div>`:"",h=an(),v=n?`
            <div class="pf-link-card">
                <h3 class="pf-link-card__title">Payroll Reports</h3>
                <p class="pf-link-card__subtitle">Open your latest payroll export.</p>
                <div class="pf-link-list">
                    <a
                        class="pf-link-item"
                        id="pr-provider-link"
                        ${h?`href="${x(h)}" target="_blank" rel="noopener noreferrer"`:'aria-disabled="true"'}
                    >
                        <span class="pf-link-item__icon">${vt}</span>
                        <span class="pf-link-item__body">
                            <span class="pf-link-item__title">Open Payroll Export</span>
                            <span class="pf-link-item__meta">${x(h||"Add a provider link in Configuration")}</span>
                        </span>
                    </a>
                </div>
            </div>
        `:"",w="",y=!n&&i?`<article class="pf-step-card pf-step-detail">${i}</article>`:"",g=!n&&f?`<article class="pf-step-card pf-step-detail">
                            <h3 class="pf-step-subtitle">Checklist</h3>
                            <ul class="pf-step-checklist">${f}</ul>
                        </article>`:"",E=!n||d||p?`
            <article class="pf-step-card pf-step-detail">
                <p class="pf-step-title">${x(d)}</p>
                ${!n&&t.statusHint?`<p class="pf-step-note">${x(t.statusHint)}</p>`:""}
                ${p}
            </article>
        `:"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${x(Ie)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${x(t.title)}</h2>
            <p class="pf-hero-copy">${x(t.summary||"")}</p>
            <p class="pf-hero-hint">${x(re.statusText||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${v}
            ${w}
            ${E}
            ${y}
            ${g}
            ${o?`
                ${De({textareaId:"step-notes-input",value:a,permanentId:"step-notes-permanent",isPermanent:s,saveButtonId:"step-notes-save"})}
                ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:l,signoffInputId:`step-signoff-${e}`,signoffValue:c,isComplete:r,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`,subtext:"Ready to move on? Save and click Done when finished."})}
            `:""}
        </section>
    `}function oa(e,t){let n=re.focusedIndex===t?"pf-step-card--active":"",o=_n(sa(e.id));return`
        <article class="pf-step-card pf-clickable ${n}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${x(e.title)}</h3>
        </article>
    `}function aa(){return`
        <footer class="pf-brand-footer">
            <div class="pf-brand-text">
                <div class="pf-brand-label">prairie.forge</div>
                <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${dn}</div>
                <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
            </div>
        </footer>
    `}function sa(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}function ra(){var n,o,a,s,l,c,r,i;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",()=>{var f;Xn(),(f=document.getElementById("pf-hero"))==null||f.scrollIntoView({behavior:"smooth",block:"start"})}),(o=document.getElementById("nav-selector"))==null||o.addEventListener("click",()=>{window.location.href="../module-selector/index.html"}),(a=document.getElementById("nav-prev"))==null||a.addEventListener("click",()=>Vn(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>Vn(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",f=>{f.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",f=>{!(t!=null&&t.contains(f.target))&&!(e!=null&&e.contains(f.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(l=document.getElementById("nav-roster"))==null||l.addEventListener("click",()=>{Fn("SS_Employee_Roster"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(c=document.getElementById("nav-accounts"))==null||c.addEventListener("click",()=>{Fn("SS_Chart_of_Accounts"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(r=document.getElementById("nav-expense-map"))==null||r.addEventListener("click",async()=>{t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"),await la()}),(i=document.getElementById("showConfigSheets"))==null||i.addEventListener("click",async()=>{await ia()})}async function ia(){if(typeof Excel=="undefined"){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.name.toUpperCase().startsWith("SS_")&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[Config] Made visible: ${a.name}`),n++)}),await e.sync();let o=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");o.load("isNullObject"),await e.sync(),o.isNullObject||(o.activate(),o.getRange("A1").select(),await e.sync()),console.log(`[Config] ${n} system sheets now visible`)})}catch(e){console.error("[Config] Error unhiding system sheets:",e)}}async function Fn(e){if(!e||typeof Excel=="undefined")return;let t={SS_Employee_Roster:["Employee","Department","Pay_Rate","Status","Hire_Date"],SS_Chart_of_Accounts:["Account_Number","Account_Name","Type","Category"]};try{await Excel.run(async n=>{let o=n.workbook.worksheets.getItemOrNullObject(e);if(o.load("isNullObject,visibility"),await n.sync(),o.isNullObject){o=n.workbook.worksheets.add(e);let a=t[e]||["Column1","Column2"],s=o.getRange(`A1:${String.fromCharCode(64+a.length)}1`);s.values=[a],s.format.font.bold=!0,s.format.fill.color="#f0f0f0",s.format.autofitColumns(),await n.sync()}else o.visibility=Excel.SheetVisibility.visible,await n.sync();o.activate(),o.getRange("A1").select(),await n.sync(),console.log(`[Quick Access] Opened sheet: ${e}`)})}catch(n){console.error("Error opening reference sheet:",n)}}async function la(){try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItemOrNullObject("PR_Expense_Mapping");if(t.load("isNullObject,visibility"),await e.sync(),t.isNullObject){t=e.workbook.worksheets.add("PR_Expense_Mapping");let n=["Expense_Category","GL_Account","Description","Active"],o=t.getRange("A1:D1");o.values=[n],o.format.font.bold=!0}else t.visibility=Excel.SheetVisibility.visible,await e.sync();t.activate(),t.getRange("A1").select(),await e.sync(),console.log("[Quick Access] Opened PR_Expense_Mapping")})}catch(e){console.error("Error navigating to PR_Expense_Mapping:",e)}}function ca(){document.querySelectorAll("[data-step-card]").forEach(e=>{let t=Number(e.getAttribute("data-step-index"));e.addEventListener("click",()=>at(t))})}function da(){var c,r,i,f;let e=document.getElementById("config-user-name");e==null||e.addEventListener("change",d=>{let u=d.target.value.trim();V(tt,u);let p=document.getElementById("config-reviewer-name");p&&!p.value&&(p.value=u)}),En("config-payroll-date",{onChange:d=>{if(V("PR_Payroll_Date",d),Qt(0),!d)return;let u=va(d);if(u){let h=document.getElementById("config-accounting-period");h&&(h.value=u),V("PR_Accounting_Period",u),Y.overrides.accountingPeriod=!1}let p=ba(d);if(p){let h=document.getElementById("config-je-id");h&&(h.value=p),V("PR_Journal_Entry_ID",p),Y.overrides.jeId=!1}}});let t=_e(0),n=document.getElementById("config-accounting-period");n==null||n.addEventListener("change",d=>{Y.overrides.accountingPeriod=!!d.target.value,V("PR_Accounting_Period",d.target.value||""),Qt(0)});let o=document.getElementById("config-je-id");o==null||o.addEventListener("change",d=>{Y.overrides.jeId=!!d.target.value,V("PR_Journal_Entry_ID",d.target.value.trim()),Qt(0)}),(c=document.getElementById("config-company-name"))==null||c.addEventListener("change",d=>{V("SS_Company_Name",d.target.value.trim())}),(r=document.getElementById("config-payroll-provider"))==null||r.addEventListener("change",d=>{let u=d.target.value.trim();V(Kn,u)}),(i=document.getElementById("config-accounting-link"))==null||i.addEventListener("change",d=>{V("SS_Accounting_Software",d.target.value.trim())});let a=document.getElementById("config-notes");if(a==null||a.addEventListener("input",d=>{t&&V(t.note,d.target.value,{debounceMs:400})}),t){let d=document.getElementById("config-notes-permanent");d&&(d.addEventListener("click",()=>{let p=!d.classList.contains("is-locked");Ze(d,p),to(t.note,p)}),Ze(d,xe(t.note)));let u=document.getElementById("config-notes-save");u==null||u.addEventListener("click",()=>{a&&(V(t.note,a.value),Ve(u,!0))})}let s=document.getElementById("config-reviewer-name");s==null||s.addEventListener("change",d=>{let u=d.target.value.trim();t&&V(t.reviewer,u),V(tt,u);let p=document.getElementById("config-signoff-date");if(u&&p&&!p.value){let h=st();p.value=h,t&&V(t.signOff,h)}}),(f=document.getElementById("config-signoff-date"))==null||f.addEventListener("change",d=>{t&&V(t.signOff,d.target.value||"")});let l=document.getElementById("config-signoff-save");if(l==null||l.addEventListener("click",()=>{var h;let d=((h=s==null?void 0:s.value)==null?void 0:h.trim())||"",u=document.getElementById("config-signoff-date"),p=(u==null?void 0:u.value)||"";t&&(V(t.reviewer,d),V(t.signOff,p)),V(tt,d),Ve(l,!0)}),Wt(),t){let d=I(t.signOff),u=I(de[0]),p=!!(d||u==="Y"||u===!0);console.log(`[Step 0] Binding signoff toggle. signOff="${d}", complete="${u}", isComplete=${p}`),Zn({buttonId:"config-signoff-toggle",inputId:"config-signoff-date",fieldName:t.signOff,completeField:de[0],initialActive:p,stepId:0})}}function pa(e){var n,o,a,s,l,c,r,i,f,d,u,p,h,v,w,y,g,E,_,$,M;if((n=document.getElementById("step-back-btn"))==null||n.addEventListener("click",()=>{Xn()}),(o=document.getElementById("step-action-btn"))==null||o.addEventListener("click",()=>{let S=pt.find(A=>A.id===e);window.alert(S!=null&&S.ctaLabel?`${S.ctaLabel} coming soon.`:"Step actions coming soon.")}),e===1&&((a=document.getElementById("import-open-data-btn"))==null||a.addEventListener("click",()=>ma()),(s=document.getElementById("import-clear-btn"))==null||s.addEventListener("click",()=>ga())),e===2&&((l=document.getElementById("headcount-skip-btn"))==null||l.addEventListener("click",()=>{z.skipAnalysis=!z.skipAnalysis;let S=document.getElementById("headcount-skip-btn");S==null||S.classList.toggle("is-active",z.skipAnalysis),z.skipAnalysis&&nn(),Rt()}),(c=document.getElementById("roster-run-btn"))==null||c.addEventListener("click",()=>tn()),(r=document.getElementById("roster-refresh-btn"))==null||r.addEventListener("click",()=>tn()),(i=document.getElementById("roster-review-btn"))==null||i.addEventListener("click",()=>{var A;let S=((A=z.roster)==null?void 0:A.mismatches)||[];Yn("Roster Differences",S,{sourceLabel:"Roster",targetLabel:"Payroll Data"})}),(f=document.getElementById("dept-review-btn"))==null||f.addEventListener("click",()=>{var A;let S=((A=z.departments)==null?void 0:A.mismatches)||[];Yn("Department Differences",S,{sourceLabel:"Roster",targetLabel:"Payroll",formatter:m=>({name:m.employee,source:`${m.rosterDept} \u2192 ${m.payrollDept}`,isMissingFromTarget:!0})})})),e===3&&((d=document.getElementById("validation-run-btn"))==null||d.addEventListener("click",()=>Wn()),(u=document.getElementById("validation-refresh-btn"))==null||u.addEventListener("click",()=>Wn()),(p=document.getElementById("bank-amount-input"))==null||p.addEventListener("blur",qn),(h=document.getElementById("bank-amount-input"))==null||h.addEventListener("keydown",S=>{S.key==="Enter"&&qn(S)})),e===5&&((v=document.getElementById("je-run-btn"))==null||v.addEventListener("click",()=>Wa()),(w=document.getElementById("je-save-btn"))==null||w.addEventListener("click",()=>Ya()),(y=document.getElementById("je-create-btn"))==null||y.addEventListener("click",()=>qa()),(g=document.getElementById("je-export-btn"))==null||g.addEventListener("click",()=>Ka())),e===4){let S=document.querySelector(".pf-step-guide");if(S){let A="https://your-project.supabase.co/functions/v1/copilot";gn(S,{id:"expense-review-copilot",contextProvider:Ca(),systemPrompt:`You are Prairie Forge CoPilot, an expert financial analyst assistant for payroll expense review.

CONTEXT: You're embedded in the Payroll Recorder Excel add-in, helping accountants and CFOs review payroll data before journal entry export.

YOUR CAPABILITIES:
1. Analyze payroll expense data for accuracy and completeness
2. Identify trends, anomalies, and variances requiring attention
3. Prepare executive-ready insights and talking points
4. Validate journal entries before export to accounting system

COMMUNICATION STYLE:
- Be concise and actionable
- Use bullet points and tables for clarity
- Highlight issues with \u26A0\uFE0F and successes with \u2713
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
When asked about readiness, be specific about what passes and what needs attention.`})}(E=document.getElementById("expense-run-btn"))==null||E.addEventListener("click",()=>{zn()}),(_=document.getElementById("expense-refresh-btn"))==null||_.addEventListener("click",()=>{zn()})}let t=_e(e);if(console.log(`[Step ${e}] Binding interactions, fields:`,t),t){let S=e===1?"step-notes-1":"step-notes-input",A=document.getElementById(S);console.log(`[Step ${e}] Notes input found:`,!!A,`(id: ${S})`);let m=e===1?document.getElementById("step-notes-save-1"):e===2?document.getElementById("step-notes-save-2"):e===3?document.getElementById("step-notes-save-3"):e===4?document.getElementById("step-notes-save-4"):e===5?document.getElementById("step-notes-save-5"):document.getElementById("step-notes-save");A==null||A.addEventListener("input",Z=>{V(t.note,Z.target.value,{debounceMs:400}),e===2&&(z.skipAnalysis&&nn(),Rt())}),m==null||m.addEventListener("click",()=>{A&&(V(t.note,A.value),Ve(m,!0))});let D=e===1?"step-reviewer-1":"step-reviewer-name",P=document.getElementById(D);P==null||P.addEventListener("change",Z=>{let ee=Z.target.value.trim();V(t.reviewer,ee);let ie=e===1?document.getElementById("step-signoff-1"):e===2?document.getElementById("step-signoff-date"):e===3?document.getElementById("step-signoff-3"):e===4?document.getElementById("step-signoff-4"):e===5?document.getElementById("step-signoff-5"):document.getElementById(`step-signoff-${e}`);if(ee&&ie&&!ie.value){let ae=st();ie.value=ae,V(t.signOff,ae)}});let H=e===1?"step-signoff-1":e===2?"step-signoff-date":e===3?"step-signoff-3":e===4?"step-signoff-4":e===5?"step-signoff-5":`step-signoff-${e}`;console.log(`[Step ${e}] Signoff input ID: ${H}, found:`,!!document.getElementById(H)),($=document.getElementById(H))==null||$.addEventListener("change",Z=>{V(t.signOff,Z.target.value||"")});let Q=e===1?"step-notes-lock-1":"step-notes-permanent",J=document.getElementById(Q);J&&(J.addEventListener("click",()=>{let Z=!J.classList.contains("is-locked");Ze(J,Z),to(t.note,Z),e===2&&Rt()}),Ze(J,xe(t.note)));let F=e===1?document.getElementById("step-signoff-save-1"):e===2?document.getElementById("headcount-signoff-save"):e===3?document.getElementById("step-signoff-save-3"):e===4?document.getElementById("step-signoff-save-4"):e===5?document.getElementById("step-signoff-save-5"):document.getElementById(`step-signoff-save-${e}`);F==null||F.addEventListener("click",()=>{var ie,ae;let Z=((ie=P==null?void 0:P.value)==null?void 0:ie.trim())||"",ee=((ae=document.getElementById(H))==null?void 0:ae.value)||"";V(t.reviewer,Z),V(t.signOff,ee),Ve(F,!0)}),Wt();let q=de[e],ne=q?!!I(q):!1,ke=I(t.signOff),oe=e===1?"step-signoff-toggle-1":e===2?"headcount-signoff-toggle":e===3?"validation-signoff-toggle":e===4?"expense-signoff-toggle":e===5?"step-signoff-toggle-5":`step-signoff-toggle-${e}`;console.log(`[Step ${e}] Toggle button ID: ${oe}, found:`,!!document.getElementById(oe)),Zn({buttonId:oe,inputId:H,fieldName:t.signOff,completeField:q,requireNotesCheck:e===2?Pt:null,initialActive:!!(ke||ne),stepId:e,onComplete:e===3?Ma:e===4?Ba:e===2?La:null})}e===2&&Rt(),e===6&&((M=document.getElementById("archive-run-btn"))==null||M.addEventListener("click",Fa))}function at(e){if(Number.isNaN(e)||e<0||e>=we.length)return;let t=we[e];if(!t)return;St=e;let n=t.id===0?"config":"step";on({focusedIndex:e,activeView:n,activeStepId:t.id});let o=jo[t.id];o&&Ia(o),t.id===2&&!z.hasAnalyzed&&tn()}function Vn(e){if(re.activeView==="home"&&e>0){at(0);return}let t=re.focusedIndex+e,n=Math.max(0,Math.min(we.length-1,t));at(n)}function ua(){if(re.activeView!=="home"||St===null)return;let e=document.querySelector(`[data-step-card][data-step-index="${St}"]`);St=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}async function Xn(){let e=jt(_t);await Vt(e.sheetName,e.title,e.subtitle),on({activeView:"home",activeStepId:null})}function on(e){Object.assign(re,e),fe()}function $e(){return I(tt)||I("SS_Default_Reviewer")||""}function Xt(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function Qn(e){let t=document.getElementById("je-save-btn");t&&t.classList.toggle("is-saved",e)}function Qt(e){let t=Pe[e],n=de[e];if(!t||!n)return;let o=I(t.signOff),a=I(n);if(!(!!o||a==="Y"||a===!0))return;console.log(`[Signoff] Clearing completion for step ${e} due to field change`),V(t.signOff,""),V(n,"");let l=document.querySelector(`[id$="-signoff-toggle"], [id$="-signoff-toggle-${e}"]`);l&&(l.classList.remove("is-active"),l.setAttribute("aria-pressed","false"));let c=document.querySelector('[id^="config-signoff-"], [id^="step-signoff-"]');c&&(c.value="")}function fa(){let e={};return console.log("[Signoff] Checking step completion status..."),Object.keys(Pe).forEach(t=>{let n=parseInt(t,10),o=Pe[n];if(!o){e[n]=!1;return}let a=I(o.signOff),s=de[n],l=I(s),c=!!a||l==="Y"||l===!0;e[n]=c,console.log(`[Signoff] Step ${n}: signOff="${a}", complete="${l}" \u2192 ${c?"COMPLETE":"pending"}`)}),console.log("[Signoff] Status summary:",e),e}function Zn({buttonId:e,inputId:t,fieldName:n,completeField:o,requireNotesCheck:a,onComplete:s,initialActive:l=!1,stepId:c=null}){let r=document.getElementById(e);if(!r){console.warn(`[Signoff] Button not found: ${e}`);return}let i=t?document.getElementById(t):null,f=l||!!(i!=null&&i.value);Xt(r,f),console.log(`[Signoff] Bound ${e}, initial active: ${f}, stepId: ${c}`),r.addEventListener("click",()=>{if(console.log(`[Signoff] Done button clicked: ${e}, stepId: ${c}`),c!==null&&c>0){let u=fa(),{canComplete:p,message:h}=On(c,u),v=r.classList.contains("is-active");if(console.log(`[Signoff] canComplete: ${p}, isCurrentlyActive: ${v}`),!v&&!p){console.log(`[Signoff] BLOCKED: ${h}`),Tn(h);return}}if(a&&!a()){window.alert("Please add notes before completing this step.");return}let d=!r.classList.contains("is-active");if(console.log(`[Signoff] ${e} clicked, toggling to: ${d}`),Xt(r,d),i&&(i.value=d?st():""),n){let u=d?st():"";console.log(`[Signoff] Writing ${n} = "${u}"`),V(n,u)}if(o){let u=d?"Y":"";console.log(`[Signoff] Writing ${o} = "${u}"`),V(o,u)}d&&typeof s=="function"&&s()}),i&&i.addEventListener("change",()=>{let d=!!i.value,u=r.classList.contains("is-active");d!==u&&(console.log(`[Signoff] Date input changed, syncing button to: ${d}`),Xt(r,d),n&&V(n,i.value||""),o&&V(o,d?"Y":""))})}async function ma(){if(!pe()){window.alert("Open this module inside Excel to access the data sheet.");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem(N.DATA);t.activate(),t.getRange("A1").select(),await e.sync()})}catch(e){console.error("Unable to open PR_Data sheet",e),window.alert(`Unable to open ${N.DATA}. Confirm the sheet exists in this workbook.`)}}async function ga(){if(!pe()){window.alert("Open this module inside Excel to clear data.");return}if(window.confirm("Are you sure you want to clear all data from PR_Data? This cannot be undone."))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(N.DATA),o=n.getUsedRangeOrNullObject();o.load("isNullObject"),await t.sync(),o.isNullObject||(n.getRange("A2:Z10000").clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),window.alert("PR_Data cleared successfully.")}catch(t){console.error("Unable to clear PR_Data sheet",t),window.alert("Unable to clear PR_Data. Please try again.")}}async function Ne(e){var a,s;if(!Kt.length)return null;if(Et){let l=e.workbook.tables.getItemOrNullObject(Et);if(l.load("name"),await e.sync(),!l.isNullObject)return l;Et=null}let t=e.workbook.tables;t.load("items/name"),await e.sync();let n=((a=t.items)==null?void 0:a.map(l=>l.name))||[];console.log("[Payroll] Looking for config table:",Kt),console.log("[Payroll] Found tables in workbook:",n);let o=(s=t.items)==null?void 0:s.find(l=>Kt.includes(l.name));return o?(console.log("[Payroll] \u2713 Config table found:",o.name),Et=o.name,e.workbook.tables.getItem(o.name)):(console.warn("[Payroll] \u26A0\uFE0F CONFIG TABLE NOT FOUND!"),console.warn("[Payroll] Expected table named: SS_PF_Config"),console.warn("[Payroll] Available tables:",n),console.warn("[Payroll] To fix: Select your data in SS_PF_Config sheet \u2192 Insert \u2192 Table \u2192 Name it 'SS_PF_Config'"),null)}async function eo(){if(!pe()){Y.loaded=!0;return}try{await Excel.run(async e=>{let t=await Ne(e);if(!t){console.warn("Payroll Recorder: SS_PF_Config table is missing."),Y.loaded=!0;return}let n=t.getDataBodyRange();n.load("values"),await e.sync();let o=n.values||[],a={},s={};o.forEach(l=>{var r,i;let c=ce(l[B.FIELD]);c&&(a[c]=(r=l[B.VALUE])!=null?r:"",s[c]=(i=l[B.PERMANENT])!=null?i:"")}),Y.values=a,Y.permanents=s,Y.overrides.accountingPeriod=!!(a.PR_Accounting_Period||a.Accounting_Period),Y.overrides.jeId=!!(a.PR_Journal_Entry_ID||a.Journal_Entry_ID),Y.loaded=!0})}catch(e){console.warn("Payroll Recorder: unable to load PF_Config table.",e),Y.loaded=!0}}function I(e){var t;return(t=Y.values[e])!=null?t:""}function ha(){let e=Object.keys(Y.values||{});return ot.find(n=>e.includes(n))||ot[0]}function Dt(){return I(ha())}function an(){return(I(Kn)||I("Payroll_Provider_Link")||"").trim()}function xe(e){return no(Y.permanents[e])}function ya(e){let t=de[e];return t?no(I(t)):!1}function to(e,t){let n=ce(e);n&&(Y.permanents[n]=t?"Y":"N",Ea(n,t?"Y":"N"))}function no(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function ce(e){return String(e!=null?e:"").trim()}function oo(e){let t=String(e!=null?e:"").trim().toLowerCase();return t?["total","totals","grand total","subtotal","summary","employee","employee name","name","full name","header","column","n/a","none","blank","null","undefined"].some(o=>t===o||t===o.replace(/\s+/g,"")):!0}function be(e){if(!e)return"";let t=At(e);return t?`${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function va(e){let t=At(e);return t?t.year<1900||t.year>2100?(console.warn("deriveAccountingPeriod - Invalid year:",t.year,"from input:",e),""):`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function ba(e){let t=At(e);return t?t.year<1900||t.year>2100?(console.warn("deriveJeId - Invalid year:",t.year,"from input:",e),""):`PR-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function st(){return xt(new Date)}function V(e,t,n={}){var l;let o=ce(e);Y.values[o]=t!=null?t:"";let a=(l=n.debounceMs)!=null?l:0;if(!a){let c=je.get(o);c&&clearTimeout(c),je.delete(o),nt(o,t!=null?t:"");return}je.has(o)&&clearTimeout(je.get(o));let s=setTimeout(()=>{je.delete(o),nt(o,t!=null?t:"")},a);je.set(o,s)}var wa=["PR_Accounting_Period","PTO_Accounting_Period","Accounting_Period"];async function nt(e,t){let n=ce(e);if(Y.values[n]=t!=null?t:"",console.log(`[Payroll] Writing config: ${n} = "${t}"`),!pe()){console.warn("[Payroll] Excel runtime not available - cannot write");return}let o=wa.some(a=>n===a||n.toLowerCase()===a.toLowerCase());try{await Excel.run(async a=>{var u;let s=await Ne(a);if(!s){console.error("[Payroll] \u274C Cannot write - config table not found");return}let l=s.getDataBodyRange(),c=s.getHeaderRowRange();l.load("values"),c.load("values"),await a.sync();let r=c.values[0]||[],i=l.values||[],f=r.length;console.log(`[Payroll] Table has ${i.length} rows, ${f} columns`);let d=[];if(i.forEach((p,h)=>{ce(p[B.FIELD])===n&&d.push(h)}),d.length===0){Y.permanents[n]=(u=Y.permanents[n])!=null?u:Bn;let p=new Array(f).fill("");if(B.TYPE>=0&&B.TYPE<f&&(p[B.TYPE]=Mo),B.FIELD>=0&&B.FIELD<f&&(p[B.FIELD]=n),B.VALUE>=0&&B.VALUE<f&&(p[B.VALUE]=t!=null?t:""),B.PERMANENT>=0&&B.PERMANENT<f&&(p[B.PERMANENT]=Bn),console.log("[Payroll] Adding NEW row:",p),s.rows.add(null,[p]),await a.sync(),o){let h=s.rows;h.load("count"),await a.sync();let v=h.count-1,y=s.rows.getItemAt(v).getRange().getCell(0,B.VALUE);y.numberFormat=[["@"]],y.values=[[t!=null?t:""]],await a.sync(),console.log(`[Payroll] \u2713 Applied text format to ${n}`)}console.log(`[Payroll] \u2713 New row added for ${n}`)}else{let p=d[0];console.log(`[Payroll] Updating existing row ${p} for ${n}`);let h=l.getCell(p,B.VALUE);if(o&&(h.numberFormat=[["@"]]),h.values=[[t!=null?t:""]],await a.sync(),console.log(`[Payroll] \u2713 Updated ${n}`),d.length>1){console.log(`[Payroll] Found ${d.length-1} duplicate rows for ${n}, removing...`);let v=d.slice(1).reverse();for(let w of v)try{s.rows.getItemAt(w).delete()}catch(y){console.warn(`[Payroll] Could not delete duplicate row ${w}:`,y.message)}await a.sync(),console.log(`[Payroll] \u2713 Removed duplicate rows for ${n}`)}}})}catch(a){console.error(`[Payroll] \u274C Write failed for ${e}:`,a)}}async function Ea(e,t){let n=ce(e);if(n&&pe()){Y.permanents[n]=t;try{await Excel.run(async o=>{let a=await Ne(o);if(!a){console.warn(`Payroll Recorder: unable to locate config table when toggling ${e} permanent flag.`);return}let s=a.getDataBodyRange();s.load("values"),await o.sync();let c=(s.values||[]).findIndex(r=>ce(r[B.FIELD])===n);c!==-1&&(s.getCell(c,B.PERMANENT).values=[[t]],await o.sync())})}catch(o){console.warn(`Payroll Recorder: unable to update permanent flag for ${e}`,o)}}}function At(e){if(!e)return null;let t=String(e).trim(),n=/^(\d{4})-(\d{2})-(\d{2})/.exec(t);if(n){let l=Number(n[1]),c=Number(n[2]),r=Number(n[3]);if(l&&c&&r)return{year:l,month:c,day:r}}let o=/^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(t);if(o){let l=Number(o[1]),c=Number(o[2]),r=Number(o[3]);if(r&&l&&c)return{year:r,month:l,day:c}}let a=Number(e);if(Number.isFinite(a)&&a>4e4&&a<6e4){let c=Math.floor(a-25569)*86400*1e3,r=new Date(c);if(!isNaN(r.getTime())){let i=`${r.getUTCFullYear()}-${String(r.getUTCMonth()+1).padStart(2,"0")}-${String(r.getUTCDate()).padStart(2,"0")}`;return console.log("DEBUG parseDateInput - Converted Excel serial",a,"to",i),{year:r.getUTCFullYear(),month:r.getUTCMonth()+1,day:r.getUTCDate()}}}let s=new Date(t);return isNaN(s.getTime())?(console.warn("DEBUG parseDateInput - Could not parse date value:",e),null):{year:s.getFullYear(),month:s.getMonth()+1,day:s.getDate()}}function xt(e){if(e._isUTC){let a=e.getUTCFullYear(),s=String(e.getUTCMonth()+1).padStart(2,"0"),l=String(e.getUTCDate()).padStart(2,"0");return`${a}-${s}-${l}`}let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),o=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${o}`}function jn(e){if(!e)return null;if(typeof e=="string"){let n=e.match(/^(\d{4})-(\d{2})-(\d{2})/);if(n)return`${n[1]}-${n[2]}-${n[3]}`}let t=At(e);return t?`${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:null}function Ca(){return async()=>{if(!pe())return null;try{return await Excel.run(async e=>{var l,c,r;let t={timestamp:new Date().toISOString(),period:null,summary:{},departments:[],recentPeriods:[],dataQuality:{}},n=await Ne(e);if(n){let i=n.getDataBodyRange();i.load("values"),await e.sync();let f=i.values||[];for(let d of f){let u=String(d[B.FIELD]||"").trim(),p=d[B.VALUE];u.toLowerCase().includes("accounting")&&u.toLowerCase().includes("period")&&(t.period=String(p||"").trim())}}let o=e.workbook.worksheets.getItemOrNullObject(N.DATA_CLEAN);if(o.load("isNullObject"),await e.sync(),!o.isNullObject){let i=o.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&((l=i.values)==null?void 0:l.length)>1){let f=i.values[0].map(g=>ve(g)),d=i.values.slice(1),u=f.findIndex(g=>g.includes("amount")),p=He(f),h=f.findIndex(g=>g.includes("employee")),v=0,w=new Set,y=new Map;for(let g of d){let E=Number(g[u])||0;if(v+=E,h>=0){let _=String(g[h]||"").trim();_&&w.add(_)}if(p>=0){let _=String(g[p]||"").trim();_&&y.set(_,(y.get(_)||0)+E)}}t.summary={total:v,employeeCount:w.size,avgPerEmployee:w.size?v/w.size:0,rowCount:d.length},t.departments=Array.from(y.entries()).map(([g,E])=>({name:g,total:E,percentOfTotal:v?E/v:0})).sort((g,E)=>E.total-g.total),t.dataQuality.dataCleanReady=!0,t.dataQuality.rowCount=d.length}}let a=e.workbook.worksheets.getItemOrNullObject(N.ARCHIVE_SUMMARY);if(a.load("isNullObject"),await e.sync(),!a.isNullObject){let i=a.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&((c=i.values)==null?void 0:c.length)>1){let f=i.values[0].map(p=>ve(p)),d=f.findIndex(p=>p.includes("period")),u=f.findIndex(p=>p.includes("total"));t.recentPeriods=i.values.slice(1,6).map(p=>({period:p[d]||"",total:Number(p[u])||0})),t.dataQuality.archiveAvailable=!0,t.dataQuality.periodsAvailable=t.recentPeriods.length}}let s=e.workbook.worksheets.getItemOrNullObject(N.JE_DRAFT);if(s.load("isNullObject"),await e.sync(),!s.isNullObject){let i=s.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&((r=i.values)==null?void 0:r.length)>1){let f=i.values[0].map(v=>ve(v)),d=f.findIndex(v=>v.includes("debit")),u=f.findIndex(v=>v.includes("credit")),p=0,h=0;for(let v of i.values.slice(1))p+=Number(v[d])||0,h+=Number(v[u])||0;t.journalEntry={totalDebit:p,totalCredit:h,difference:Math.abs(p-h),isBalanced:Math.abs(p-h)<.01,lineCount:i.values.length-1},t.dataQuality.jeDraftReady=!0}}return console.log("CoPilot context gathered:",t),t})}catch(e){return console.warn("CoPilot context provider error:",e),null}}}function x(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function ye(e,t){return`
        <div class="pf-labeled-button">
            ${e}
            <span class="pf-button-label">${x(t)}</span>
        </div>
    `}function pe(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}function _e(e){return Pe[e]||null}function ka(){var n,o,a,s;let e=Math.abs((o=(n=z.roster)==null?void 0:n.difference)!=null?o:0),t=Math.abs((s=(a=z.departments)==null?void 0:a.difference)!=null?s:0);return e>0||t>0}function Pt(){return!z.skipAnalysis&&ka()}function le(e){return e==null||Number.isNaN(e)?"---":typeof e!="number"?e:e.toLocaleString(void 0,{minimumFractionDigits:2,maximumFractionDigits:2})}function ao(e){let t=sn(e);return Number.isFinite(t)?t.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}):""}function Ra(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let o=String(n);return/[",\n]/.test(o)?`"${o.replace(/"/g,'""')}"`:o}).join(",")).join(`
`)}function Sa(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),o=URL.createObjectURL(n),a=document.createElement("a");a.href=o,a.download=e,document.body.appendChild(a),a.click(),a.remove(),setTimeout(()=>URL.revokeObjectURL(o),1e3)}function sn(e){if(typeof e=="number")return e;if(e==null)return NaN;let t=String(e).replace(/[^0-9.-]/g,""),n=Number.parseFloat(t);return Number.isFinite(n)?n:NaN}function xa(e){if(e instanceof Date)return xt(e);if(typeof e=="number"&&!Number.isNaN(e)){let o=_a(e);return o?xt(o):""}let t=String(e!=null?e:"").trim();if(!t)return"";if(/^\d{4}-\d{2}-\d{2}$/.test(t))return t;let n=new Date(t);return Number.isNaN(n.getTime())?t:xt(n)}function _a(e){if(!Number.isFinite(e))return null;let t=Math.floor(e-25569);if(!Number.isFinite(t))return null;let n=t*86400*1e3,o=new Date(n);return o._isUTC=!0,o}function Da(e){if(!e)return"";let t=new Date(e);return Number.isNaN(t.getTime())?e:t.toLocaleDateString(void 0,{month:"short",day:"numeric",year:"numeric"})}function Ct(e){if(e==null||e==="")return 0;let t=Number(e);return Number.isFinite(t)?t:0}function Aa(e){let t=me(e).toLowerCase();return t?t.includes("burden")||t.includes("tax")||t.includes("benefit")||t.includes("fica")||t.includes("insurance")||t.includes("health")||t.includes("medicare")?"burden":t.includes("bonus")||t.includes("commission")||t.includes("variable")||t.includes("overtime")||t.includes("per diem")?"variable":"fixed":"variable"}function Hn(e){if(!e||e.length<2)return[];let t=(e[0]||[]).map(a=>ve(a));console.log("parseExpenseRows - headers:",t);let n={payrollDate:t.findIndex(a=>a.includes("payroll")&&a.includes("date")),employee:t.findIndex(a=>a.includes("employee")),department:t.findIndex(a=>a.includes("department")),fixed:t.findIndex(a=>a.includes("fixed")),variable:t.findIndex(a=>a.includes("variable")),burden:t.findIndex(a=>a.includes("burden")),amount:t.findIndex(a=>a.includes("amount")),expenseReview:t.findIndex(a=>a.includes("expense")&&a.includes("review")),category:t.findIndex(a=>a.includes("payroll")&&a.includes("category"))};if(console.log("parseExpenseRows - column indexes:",n),n.payrollDate>=0){let a=new Set;for(let s=1;s<e.length;s++){let l=e[s][n.payrollDate];l&&a.add(String(l))}console.log("parseExpenseRows - unique payroll dates found:",[...a].slice(0,20))}let o=[];for(let a=1;a<e.length;a+=1){let s=e[a],l=xa(n.payrollDate>=0?s[n.payrollDate]:null);if(!l)continue;let c=n.employee>=0?me(s[n.employee]):"",r=n.department>=0&&me(s[n.department])||"Unassigned",i=n.fixed>=0?Ct(s[n.fixed]):null,f=n.variable>=0?Ct(s[n.variable]):null,d=n.burden>=0?Ct(s[n.burden]):null,u=0,p=0,h=0;if(i!==null||f!==null||d!==null)u=i!=null?i:0,p=f!=null?f:0,h=d!=null?d:0;else{let v=n.amount>=0?Ct(s[n.amount]):0,w=Aa(n.expenseReview>=0?s[n.expenseReview]:s[n.category]);w==="fixed"?u=v:w==="burden"?h=v:p=v}u===0&&p===0&&h===0||o.push({period:l,employee:c,department:r||"Unassigned",fixed:u,variable:p,burden:h})}return o}function Un(e){let t=new Map;e.forEach(o=>{let a=o.period;if(!a)return;t.has(a)||t.set(a,{key:a,label:Da(a),employees:new Set,departments:new Map,summary:{fixed:0,variable:0,burden:0}});let s=t.get(a);s.employees.add(o.employee||`EMP-${s.employees.size+1}`);let l=o.department||"Unassigned";s.departments.has(l)||s.departments.set(l,{name:l,fixed:0,variable:0,burden:0,employees:new Set});let c=s.departments.get(l);c.fixed+=o.fixed,c.variable+=o.variable,c.burden+=o.burden,c.employees.add(o.employee||`EMP-${c.employees.size+1}`),s.summary.fixed+=o.fixed,s.summary.variable+=o.variable,s.summary.burden+=o.burden});let n=[];return t.forEach(o=>{let a=o.summary.fixed+o.summary.variable+o.summary.burden,s=Array.from(o.departments.values()).map(r=>{let i=r.fixed+r.variable,f=i+r.burden;return{name:r.name,fixed:r.fixed,variable:r.variable,gross:i,burden:r.burden,allIn:f,percent:a?f/a:0,headcount:r.employees.size,delta:0}});s.sort((r,i)=>i.allIn-r.allIn);let l={employeeCount:o.employees.size,fixed:o.summary.fixed,variable:o.summary.variable,burden:o.summary.burden,total:a,burdenRate:a?o.summary.burden/a:0,delta:0},c={name:"Totals",fixed:o.summary.fixed,variable:o.summary.variable,gross:o.summary.fixed+o.summary.variable,burden:o.summary.burden,allIn:a,percent:a?1:0,headcount:o.employees.size,delta:0,isTotal:!0};n.push({key:o.key,label:o.label,summary:l,departments:s,totalsRow:c})}),n.sort((o,a)=>o.key<a.key?1:-1)}function Gn(e,t){console.log("buildExpenseReviewPeriods - cleanValues rows:",(e==null?void 0:e.length)||0),console.log("buildExpenseReviewPeriods - archiveValues rows:",(t==null?void 0:t.length)||0);let n=Un(Hn(e)),o=Un(Hn(t));console.log("buildExpenseReviewPeriods - currentPeriods:",n.map(i=>{var f,d;return{key:i.key,employees:(f=i.summary)==null?void 0:f.employeeCount,total:(d=i.summary)==null?void 0:d.total}})),console.log("buildExpenseReviewPeriods - archivePeriods:",o.map(i=>{var f,d;return{key:i.key,employees:(f=i.summary)==null?void 0:f.employeeCount,total:(d=i.summary)==null?void 0:d.total}}));let a=new Map(o.map(i=>[i.key,i])),s=[];n.length&&(s.push(n[0]),a.delete(n[0].key)),o.forEach(i=>{s.length>=6||s.some(f=>f.key===i.key)||s.push(i)}),console.log("buildExpenseReviewPeriods - combined before filter:",s.map(i=>{var f,d;return{key:i.key,employees:(f=i.summary)==null?void 0:f.employeeCount,total:(d=i.summary)==null?void 0:d.total}}));let l=3,c=1e3,r=s.filter(i=>{var p,h,v,w,y;if(!i||!i.key)return console.log("buildExpenseReviewPeriods - EXCLUDED (no key):",i),!1;let f=((p=i.summary)==null?void 0:p.total)||(((h=i.summary)==null?void 0:h.fixed)||0)+(((v=i.summary)==null?void 0:v.variable)||0)+(((w=i.summary)==null?void 0:w.burden)||0),d=((y=i.summary)==null?void 0:y.employeeCount)||0;if(s.indexOf(i)===0)return console.log(`buildExpenseReviewPeriods - INCLUDED (current): ${i.key} - ${d} employees, $${f}`),!0;let u=d>=l&&f>=c;return console.log(`buildExpenseReviewPeriods - ${u?"INCLUDED":"EXCLUDED"}: ${i.key} - ${d} employees, $${f} (needs >=${l} emp, >=$${c})`),u}).sort((i,f)=>i.key<f.key?1:-1).slice(0,6);return console.log("buildExpenseReviewPeriods - FINAL periods:",r.map(i=>i.key)),r.forEach((i,f)=>{let d=r[f+1],u=d?i.summary.total-d.summary.total:0;i.summary.delta=u;let p=new Map(((d==null?void 0:d.departments)||[]).map(h=>[h.name,h]));i.departments.forEach(h=>{let v=p.get(h.name);h.delta=v?h.allIn-v.allIn:0}),i.totalsRow.delta=u}),r}async function zn(){if(!pe()){kt({loading:!1,lastError:"Excel runtime is unavailable."});return}kt({loading:!0,lastError:null});try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItemOrNullObject(N.DATA_CLEAN),o=t.workbook.worksheets.getItemOrNullObject(N.ARCHIVE_SUMMARY),a=t.workbook.worksheets.getItemOrNullObject(N.EXPENSE_REVIEW);if(n.load("isNullObject, name"),o.load("isNullObject, name"),a.load("isNullObject, name"),await t.sync(),console.log("Expense Review - Sheet check:",{cleanSheet:n.isNullObject?"MISSING":n.name,archiveSheet:o.isNullObject?"MISSING":o.name,reviewSheet:a.isNullObject?"MISSING":a.name}),a.isNullObject){console.log("Creating PR_Expense_Review sheet...");let r=t.workbook.worksheets.add(N.EXPENSE_REVIEW);await t.sync();let i=t.workbook.worksheets.getItem(N.EXPENSE_REVIEW),f=[],d=[];if(!n.isNullObject){let p=n.getUsedRangeOrNullObject();p.load("values"),await t.sync(),f=p.isNullObject?[]:p.values||[]}if(!o.isNullObject){let p=o.getUsedRangeOrNullObject();p.load("values"),await t.sync(),d=p.isNullObject?[]:p.values||[]}let u=Gn(f,d);return await Jn(t,i,u),u}let s=[],l=[];if(n.isNullObject)console.warn("Expense Review - PR_Data_Clean sheet not found, using empty data");else{let r=n.getUsedRangeOrNullObject();r.load("values"),await t.sync(),s=r.isNullObject?[]:r.values||[],console.log("Expense Review - PR_Data_Clean rows:",s.length)}if(o.isNullObject)console.warn("Expense Review - PR_Archive_Summary sheet not found, using empty data");else{let r=o.getUsedRangeOrNullObject();r.load("values"),await t.sync(),l=r.isNullObject?[]:r.values||[],console.log("Expense Review - PR_Archive_Summary rows:",l.length)}let c=Gn(s,l);return console.log("Expense Review - Periods built:",c.length),await Jn(t,a,c),c});kt({loading:!1,periods:e,lastError:null}),await Ho(),fe()}catch(e){console.error("Expense Review: unable to build executive summary",e),console.error("Error details:",e.message,e.stack),kt({loading:!1,lastError:`Unable to build the Expense Review: ${e.message||"Unknown error"}`,periods:[]})}}async function Jn(e,t,n){if(!t){console.error("writeExpenseReviewSheet: sheet is null/undefined");return}console.log("writeExpenseReviewSheet: Building executive dashboard with",n.length,"periods");try{let b=t.getUsedRangeOrNullObject();b.load("address");let R=t.charts;R.load("items"),await e.sync(),b.isNullObject||(b.clear(),await e.sync());for(let U=R.items.length-1;U>=0;U--)R.items[U].delete();await e.sync()}catch(b){console.warn("Could not clear sheet:",b)}let o=n[0]||{},a=n[1]||{},s=o.summary||{},l=a.summary||{},c=I("PR_Accounting_Period")||Dt()||"",r=Number(s.total)||0,i=Number(l.total)||0,f=r-i,d=i?f/i:0,u=Number(s.employeeCount)||0,p=Number(l.employeeCount)||0,h=u-p,v=u?r/u:0,w=p?i/p:0,y=v-w,g=Pa(n),E=$a(o,n),_=o.label||o.key||"Current Period",$=new Date().toLocaleString("en-US",{month:"short",day:"numeric",year:"numeric",hour:"numeric",minute:"2-digit"}),M=b=>b>0?"\u25B2":b<0?"\u25BC":"\u2014",S=n.map(b=>{var R;return((R=b.summary)==null?void 0:R.total)||0}).filter(b=>b>0),A=n.map(b=>{let R=b.summary||{},U=R.employeeCount||0;return U>0?(R.total||0)/U:0}).filter(b=>b>0),m=n.slice(0,-1).map((b,R)=>{var ue,K,W;let U=((ue=b.summary)==null?void 0:ue.total)||0,se=((W=(K=n[R+1])==null?void 0:K.summary)==null?void 0:W.total)||0;return se>0?(U-se)/se:0}),D=(b,R=null)=>{let U=R!==null?[...b,R]:b;if(!U.length)return{min:0,max:0,avg:0};let se=Math.min(...U),ue=Math.max(...U),K=b.length?b:U,W=K.reduce((Re,Ee)=>Re+Ee,0)/K.length;return{min:se,max:ue,avg:W}},P=D(S,r),H=D(A,v),Q=D(m),J=(b,R,U,se=20)=>{if(U<=R)return"\u2591".repeat(se);let ue=U-R,K=Math.max(0,Math.min(1,(b-R)/ue)),W=Math.round(K*(se-1)),Re="";for(let Ee=0;Ee<se;Ee++)Ee===W?Re+="\u25CF":Re+="\u2591";return Re},F=Number(s.fixed)||0,q=Number(s.variable)||0,ne=Number(s.burden)||0,ke=F+q,oe=r?ne/r:0,Z=Number(l.fixed)||0,ee=Number(l.variable)||0,ie=Number(l.burden)||0,ae=i?ie/i:0,te=o.departments||[],ge=te.filter(b=>{let R=(b.name||"").toLowerCase();return R.includes("sales")||R.includes("marketing")}),rt=te.filter(b=>{let R=(b.name||"").toLowerCase();return!R.includes("sales")&&!R.includes("marketing")}),$t=ge.reduce((b,R)=>b+(R.variable||0),0),Oe=ge.reduce((b,R)=>b+(R.headcount||0),0),Te=rt.reduce((b,R)=>b+(R.variable||0),0),Le=rt.reduce((b,R)=>b+(R.headcount||0),0),It=Oe?$t/Oe:0,Nt=Le?Te/Le:0,Ot=u?F/u:0,T=[],k=0,C={};C.headerStart=k;let rn=c||_;if(typeof c=="number"||!isNaN(Number(c))&&c){let b=Number(c);if(b>4e4&&b<6e4){let R=new Date(1899,11,30);rn=new Date(R.getTime()+b*24*60*60*1e3).toLocaleDateString("en-US",{year:"numeric",month:"long",day:"numeric"})}}T.push(["PAYROLL EXPENSE REVIEW"]),k++,T.push([`Period: ${rn}`]),k++,T.push([`Generated: ${$}`]),k++,T.push([""]),k++,C.headerEnd=k-1,C.execSummaryStart=k,T.push(["EXECUTIVE SUMMARY"]),k++,C.execSummaryHeader=k-1,T.push([""]),k++,T.push(["","Pay Date","Headcount","Fixed Salary","Variable Salary","Burden","Total Payroll","Burden Rate"]),k++,C.execSummaryColHeaders=k-1,T.push(["Current Pay Period",o.label||o.key||"",u,F,q,ne,r,oe]),k++,C.execSummaryCurrentRow=k-1,T.push(["Same Period Prior Month",a.label||a.key||"",p,Z,ee,ie,i,ae]),k++,C.execSummaryPriorRow=k-1,T.push([""]),k++,T.push([""]),k++,C.execSummaryEnd=k-1,C.deptBreakdownStart=k,T.push(["CURRENT PERIOD BREAKDOWN (DEPARTMENT)"]),k++,C.deptBreakdownHeader=k-1,T.push([""]),k++,T.push(["Payroll Date",o.label||o.key||""]),k++,T.push([""]),k++,T.push(["Department","Fixed Salary","Variable Salary","Gross Pay","Burden","All-In Total","% of Total","Headcount"]),k++,C.deptColHeaders=k-1;let lo=[...te].sort((b,R)=>(R.allIn||0)-(b.allIn||0));if(C.deptDataStart=k,lo.forEach(b=>{T.push([b.name||"",b.fixed||0,b.variable||0,b.gross||0,b.burden||0,b.allIn||0,b.percent||0,b.headcount||0]),k++}),C.deptDataEnd=k-1,o.totalsRow){let b=o.totalsRow;T.push(["TOTAL",b.fixed||0,b.variable||0,b.gross||0,b.burden||0,b.allIn||0,1,b.headcount||0]),k++,C.deptTotalsRow=k-1}T.push([""]),k++,T.push([""]),k++,C.deptBreakdownEnd=k-1,C.historicalStart=k,T.push(["HISTORICAL CONTEXT"]),k++,C.historicalHeader=k-1,T.push([`Visual comparison of current period vs. historical range (${n.length} periods). The dot (\u25CF) shows where you currently stand.`]),k++,T.push([""]),k++;let X=b=>`$${Math.round(b/1e3)}K`,it=b=>`${(b*100).toFixed(1)}%`;T.push(["","Metric","Low","Range","High","","Current","Average"]),k++,C.historicalColHeaders=k-1;let co=n.map(b=>{var R;return((R=b.summary)==null?void 0:R.fixed)||0}).filter(b=>b>0),po=n.map(b=>{var R;return((R=b.summary)==null?void 0:R.variable)||0}),uo=n.map(b=>{let R=b.summary||{};return R.total?(R.burden||0)/R.total:0}),fo=n.map(b=>{let R=b.summary||{},U=R.employeeCount||0;return U>0?(R.fixed||0)/U:0}).filter(b=>b>0),Ue=D(co,F),Ge=D(po,q),ze=D(uo,oe),Je=D(fo,Ot);C.spectrumRows=[];let mo=J(r,P.min,P.max,25);T.push(["","Total Payroll",X(P.min),mo,X(P.max),"",X(r),X(P.avg)]),k++,C.spectrumRows.push(k-1);let go=J(F,Ue.min,Ue.max,25);T.push(["","Total Fixed Payroll",X(Ue.min),go,X(Ue.max),"",X(F),X(Ue.avg)]),k++,C.spectrumRows.push(k-1);let ho=J(q,Ge.min,Ge.max,25);T.push(["","Total Variable Payroll",X(Ge.min),ho,X(Ge.max),"",X(q),X(Ge.avg)]),k++,C.spectrumRows.push(k-1),T.push([""]),k++;let yo=J(Ot,Je.min,Je.max,25);T.push(["","Avg Fixed Payroll / Employee",X(Je.min),yo,X(Je.max),"",X(Ot),X(Je.avg)]),k++,C.spectrumRows.push(k-1);let vo=n.map(b=>{let U=(b.departments||[]).filter(K=>{let W=(K.name||"").toLowerCase();return W.includes("sales")||W.includes("marketing")}),se=U.reduce((K,W)=>K+(W.variable||0),0),ue=U.reduce((K,W)=>K+(W.headcount||0),0);return ue>0?se/ue:0}),lt=D(vo,It),bo=n.map(b=>{let U=(b.departments||[]).filter(K=>{let W=(K.name||"").toLowerCase();return!W.includes("sales")&&!W.includes("marketing")}),se=U.reduce((K,W)=>K+(W.variable||0),0),ue=U.reduce((K,W)=>K+(W.headcount||0),0);return ue>0?se/ue:0}),ct=D(bo,Nt);if(Oe>0){let b=J(It,lt.min,lt.max,25);T.push(["","Avg Variable / Sales & Marketing",X(lt.min),b,X(lt.max),"",X(It),`${Oe} emp`]),k++,C.spectrumRows.push(k-1)}if(Le>0){let b=J(Nt,ct.min,ct.max,25);T.push(["","Avg Variable / Other Depts",X(ct.min),b,X(ct.max),"",X(Nt),`${Le} emp`]),k++,C.spectrumRows.push(k-1)}T.push([""]),k++;let wo=J(oe,ze.min,ze.max,25);T.push(["","Burden Rate (%)",it(ze.min),wo,it(ze.max),"",it(oe),it(ze.avg)]),k++,C.spectrumRows.push(k-1),T.push([""]),k++,T.push([""]),k++,C.historicalEnd=k-1,C.trendsStart=k,T.push(["PERIOD TRENDS"]),k++,C.trendsHeader=k-1,T.push([""]),k++,T.push(["Pay Period","Total Payroll","Fixed Payroll","Variable Payroll","Burden","Headcount"]),k++,C.trendColHeaders=k-1;let ln=n.slice(0,6).reverse();C.trendDataStart=k,ln.forEach(b=>{let R=b.summary||{};T.push([b.label||b.key||"",R.total||0,R.fixed||0,R.variable||0,R.burden||0,R.employeeCount||0]),k++}),C.trendDataEnd=k-1,T.push([""]),k++,C.trendsEnd=k-1,C.chartStart=k;for(let b=0;b<15;b++)T.push([""]),k++;C.payrollChartEnd=k-1,C.headcountChartStart=k;for(let b=0;b<12;b++)T.push([""]),k++;C.headcountChartEnd=k-1,console.log("writeExpenseReviewSheet: Writing",T.length,"rows");let cn=T.map(b=>{let R=Array.isArray(b)?b:[""];for(;R.length<10;)R.push("");return R.slice(0,10)});try{let b=t.getRangeByIndexes(0,0,cn.length,10);b.values=cn,await e.sync()}catch(b){throw console.error("writeExpenseReviewSheet: Write failed",b),b}try{t.getRange("A:A").format.columnWidth=200,t.getRange("B:B").format.columnWidth=130,t.getRange("C:C").format.columnWidth=100,t.getRange("D:D").format.columnWidth=200,t.getRange("E:E").format.columnWidth=100,t.getRange("F:F").format.columnWidth=100,t.getRange("G:G").format.columnWidth=100,t.getRange("H:H").format.columnWidth=100,t.getRange("I:I").format.columnWidth=80,t.getRange("J:J").format.columnWidth=80,await e.sync();let b=t.getRange("A1");b.format.font.bold=!0,b.format.font.size=22,b.format.font.color="#1e293b",t.getRange("A2").format.font.size=11,t.getRange("A2").format.font.color="#64748b",t.getRange("A3").format.font.size=10,t.getRange("A3").format.font.color="#94a3b8",await e.sync();let R=t.getRange(`A${C.execSummaryHeader+1}`);R.format.font.bold=!0,R.format.font.size=14,R.format.font.color="#1e293b";let U=t.getRange(`A${C.execSummaryColHeaders+1}:H${C.execSummaryColHeaders+1}`);U.format.font.bold=!0,U.format.font.size=10,U.format.fill.color="#1e293b",U.format.font.color="#ffffff";let se=t.getRange(`A${C.execSummaryCurrentRow+1}:H${C.execSummaryCurrentRow+1}`);se.format.fill.color="#dcfce7",se.format.font.bold=!0;let ue=t.getRange(`A${C.execSummaryPriorRow+1}:H${C.execSummaryPriorRow+1}`);ue.format.fill.color="#f1f5f9";for(let L of[C.execSummaryCurrentRow+1,C.execSummaryPriorRow+1])t.getRange(`C${L}`).numberFormat=[["#,##0"]],t.getRange(`D${L}`).numberFormat=[["$#,##0"]],t.getRange(`E${L}`).numberFormat=[["$#,##0"]],t.getRange(`F${L}`).numberFormat=[["$#,##0"]],t.getRange(`G${L}`).numberFormat=[["$#,##0"]],t.getRange(`H${L}`).numberFormat=[["0.00%"]];await e.sync();let K=t.getRange(`A${C.deptBreakdownHeader+1}`);K.format.font.bold=!0,K.format.font.size=14,K.format.font.color="#1e293b";let W=t.getRange(`A${C.deptColHeaders+1}:H${C.deptColHeaders+1}`);W.format.font.bold=!0,W.format.font.size=10,W.format.fill.color="#1e293b",W.format.font.color="#ffffff";for(let L=C.deptDataStart;L<=C.deptDataEnd;L++){let O=L+1;t.getRange(`B${O}`).numberFormat=[["$#,##0"]],t.getRange(`C${O}`).numberFormat=[["$#,##0"]],t.getRange(`D${O}`).numberFormat=[["$#,##0"]],t.getRange(`E${O}`).numberFormat=[["$#,##0"]],t.getRange(`F${O}`).numberFormat=[["$#,##0"]],t.getRange(`G${O}`).numberFormat=[["0.00%"]],t.getRange(`H${O}`).numberFormat=[["#,##0"]],(L-C.deptDataStart)%2===1&&(t.getRange(`A${O}:H${O}`).format.fill.color="#f8fafc")}if(C.deptTotalsRow){let L=t.getRange(`A${C.deptTotalsRow+1}:H${C.deptTotalsRow+1}`);L.format.font.bold=!0,L.format.fill.color="#1e293b",L.format.font.color="#ffffff";let O=C.deptTotalsRow+1;t.getRange(`B${O}`).numberFormat=[["$#,##0"]],t.getRange(`C${O}`).numberFormat=[["$#,##0"]],t.getRange(`D${O}`).numberFormat=[["$#,##0"]],t.getRange(`E${O}`).numberFormat=[["$#,##0"]],t.getRange(`F${O}`).numberFormat=[["$#,##0"]],t.getRange(`G${O}`).numberFormat=[["0%"]],t.getRange(`H${O}`).numberFormat=[["#,##0"]]}await e.sync();let Re=t.getRange(`A${C.historicalHeader+1}`);Re.format.font.bold=!0,Re.format.font.size=14,Re.format.font.color="#1e293b",t.getRange(`A${C.historicalHeader+2}`).format.font.size=10,t.getRange(`A${C.historicalHeader+2}`).format.font.color="#64748b",t.getRange(`A${C.historicalHeader+2}`).format.font.italic=!0;let Ee=t.getRange(`A${C.historicalColHeaders+1}:H${C.historicalColHeaders+1}`);Ee.format.font.bold=!0,Ee.format.font.size=10,Ee.format.fill.color="#e2e8f0",Ee.format.font.color="#334155",t.getRange(`C${C.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`E${C.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`G${C.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`H${C.historicalColHeaders+1}`).format.horizontalAlignment="Center",C.spectrumRows.forEach(L=>{t.getRange(`D${L+1}`).format.font.name="Consolas",t.getRange(`D${L+1}`).format.font.size=14,t.getRange(`D${L+1}`).format.font.color="#6366f1",t.getRange(`D${L+1}`).format.horizontalAlignment="Center",t.getRange(`B${L+1}`).format.font.color="#334155",t.getRange(`C${L+1}`).format.font.color="#94a3b8",t.getRange(`C${L+1}`).format.horizontalAlignment="Center",t.getRange(`E${L+1}`).format.font.color="#94a3b8",t.getRange(`E${L+1}`).format.horizontalAlignment="Center",t.getRange(`G${L+1}`).format.font.bold=!0,t.getRange(`G${L+1}`).format.font.color="#1e293b",t.getRange(`G${L+1}`).format.horizontalAlignment="Center",t.getRange(`H${L+1}`).format.font.color="#94a3b8",t.getRange(`H${L+1}`).format.horizontalAlignment="Center"}),await e.sync();let Tt=t.getRange(`A${C.trendsHeader+1}`);Tt.format.font.bold=!0,Tt.format.font.size=14,Tt.format.font.color="#1e293b";let dt=t.getRange(`A${C.trendColHeaders+1}:F${C.trendColHeaders+1}`);dt.format.font.bold=!0,dt.format.font.size=10,dt.format.fill.color="#1e293b",dt.format.font.color="#ffffff";for(let L=C.trendDataStart;L<=C.trendDataEnd;L++){let O=L+1;t.getRange(`B${O}`).numberFormat=[["$#,##0"]],t.getRange(`C${O}`).numberFormat=[["$#,##0"]],t.getRange(`D${O}`).numberFormat=[["$#,##0"]],t.getRange(`E${O}`).numberFormat=[["$#,##0"]],t.getRange(`F${O}`).numberFormat=[["#,##0"]],(L-C.trendDataStart)%2===1&&(t.getRange(`A${O}:F${O}`).format.fill.color="#f8fafc")}if(await e.sync(),ln.length>=2){try{let L=t.getRange(`A${C.trendColHeaders+1}:E${C.trendDataEnd+1}`),O=t.charts.add(Excel.ChartType.lineMarkers,L,Excel.ChartSeriesBy.columns);O.setPosition(`A${C.chartStart+1}`,`J${C.payrollChartEnd+1}`),O.title.text="Payroll Expense Trends",O.title.format.font.size=14,O.title.format.font.bold=!0,O.legend.position=Excel.ChartLegendPosition.bottom,O.format.fill.setSolidColor("#ffffff"),O.format.border.lineStyle=Excel.ChartLineStyle.continuous,O.format.border.color="#e2e8f0";let We=O.axes.getItem(Excel.ChartAxisType.category);We.categoryType=Excel.ChartAxisCategoryType.textAxis,We.setCategoryNames(t.getRange(`A${C.trendDataStart+1}:A${C.trendDataEnd+1}`)),await e.sync();let Ce=O.series;Ce.load("count"),await e.sync();let he=["#3b82f6","#22c55e","#f97316","#8b5cf6"];for(let Me=0;Me<Math.min(Ce.count,he.length);Me++){let Ye=Ce.getItemAt(Me);Ye.format.line.color=he[Me],Ye.format.line.weight=2,Ye.markerStyle=Excel.ChartMarkerStyle.circle,Ye.markerSize=6,Ye.markerBackgroundColor=he[Me]}await e.sync(),console.log("writeExpenseReviewSheet: Payroll chart created successfully")}catch(L){console.warn("writeExpenseReviewSheet: Payroll chart creation failed (non-critical)",L)}try{let L=t.getRange(`A${C.trendColHeaders+1}:F${C.trendDataEnd+1}`),O=t.charts.add(Excel.ChartType.lineMarkers,L,Excel.ChartSeriesBy.columns);O.setPosition(`A${C.headcountChartStart+1}`,`J${C.headcountChartEnd+1}`),O.title.text="Headcount Trend",O.title.format.font.size=12,O.title.format.font.bold=!0,O.legend.visible=!1,O.format.fill.setSolidColor("#ffffff"),O.format.border.lineStyle=Excel.ChartLineStyle.continuous,O.format.border.color="#e2e8f0";let We=O.axes.getItem(Excel.ChartAxisType.category);We.categoryType=Excel.ChartAxisCategoryType.textAxis,We.setCategoryNames(t.getRange(`A${C.trendDataStart+1}:A${C.trendDataEnd+1}`)),await e.sync();let Ce=O.series;Ce.load("count, items/name"),await e.sync();for(let he=Ce.count-2;he>=0;he--)Ce.getItemAt(he).delete();if(await e.sync(),Ce.load("count"),await e.sync(),Ce.count>0){let he=Ce.getItemAt(0);he.format.line.color="#64748b",he.format.line.weight=2.5,he.markerStyle=Excel.ChartMarkerStyle.circle,he.markerSize=8,he.markerBackgroundColor="#64748b"}await e.sync(),console.log("writeExpenseReviewSheet: Headcount chart created successfully")}catch(L){console.warn("writeExpenseReviewSheet: Headcount chart creation failed (non-critical)",L)}}t.freezePanes.freezeRows(C.execSummaryEnd+1),t.pageLayout.orientation=Excel.PageOrientation.landscape,t.getRange("A1").select(),await e.sync(),console.log("writeExpenseReviewSheet: Formatting applied successfully")}catch(b){console.warn("writeExpenseReviewSheet: Formatting error (non-critical)",b)}}function Pa(e){var o;return!e||!e.length?!1:(((o=e[0].summary)==null?void 0:o.categories)||[]).some(a=>{let s=(a.name||"").toLowerCase();return s.includes("commission")||s.includes("bonus")||s.includes("variable")})}function $a(e,t){var l;if(!e||t.length<2)return!1;let n=t.map(c=>{var r;return((r=c.summary)==null?void 0:r.total)||0}).filter(c=>c>0);if(n.length<2)return!1;let o=n.reduce((c,r)=>c+r,0)/n.length,a=((l=e.summary)==null?void 0:l.total)||0;return(o>0?a/o:1)<.9}async function Ia(e){if(!(!pe()||!e))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItemOrNullObject(e);n.load("name"),await t.sync(),!n.isNullObject&&(n.activate(),n.getRange("A1").select(),await t.sync())})}catch(t){console.warn(`Payroll Recorder: unable to activate worksheet ${e}`,t)}}async function tn(){if(!pe()){z.lastError="Excel runtime is unavailable.",z.hasAnalyzed=!0,z.loading=!1,fe();return}z.loading=!0,z.lastError=null,fe();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),o=t.workbook.worksheets.getItem(N.DATA),a=n.getUsedRangeOrNullObject(),s=o.getUsedRangeOrNullObject();a.load("values"),s.load("values"),await t.sync();let l=a.isNullObject?[]:a.values||[],c=s.isNullObject?[]:s.values||[],r=Oa(l),i=Ta(c),f=[];r.employeeMap.forEach((p,h)=>{i.employeeMap.has(h)||f.push({name:p.name||"Unknown Employee",type:"missing_from_payroll",message:"In roster but NOT in payroll data",department:p.department||"\u2014"})}),i.employeeMap.forEach((p,h)=>{r.employeeMap.has(h)||f.push({name:p.name||"Unknown Employee",type:"missing_from_roster",message:"In payroll but NOT in roster",department:p.department||"\u2014"})}),f.sort((p,h)=>p.type!==h.type?p.type.localeCompare(h.type):(p.name||"").localeCompare(h.name||""));let d=[],u=0;return r.employeeMap.forEach((p,h)=>{let v=i.employeeMap.get(h);if(!v)return;let w=me(p.department),y=me(v.department);!w&&!y||(u+=1,w!==y&&d.push({employee:p.name||v.name||"Employee",rosterDept:w||"\u2014",payrollDept:y||"\u2014"}))}),console.log("Headcount Analysis Results:",{rosterCount:r.activeCount,payrollCount:i.totalEmployees,difference:r.activeCount-i.totalEmployees,missingFromPayroll:f.filter(p=>p.type==="missing_from_payroll").length,missingFromRoster:f.filter(p=>p.type==="missing_from_roster").length,deptMismatches:d.length}),{roster:{rosterCount:r.activeCount,payrollCount:i.totalEmployees,difference:r.activeCount-i.totalEmployees,mismatches:f},departments:{rosterCount:u,payrollCount:u,difference:d.length,mismatches:d}}});z.roster=e.roster,z.departments=e.departments,z.hasAnalyzed=!0}catch(e){console.warn("Headcount Review: unable to analyze data",e),z.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{z.loading=!1,fe()}}function et(e={},{rerender:t=!0}={}){Object.assign(G,e);let n=Number(G.prDataTotal),o=Number(G.cleanTotal);G.reconDifference=Number.isFinite(n)&&Number.isFinite(o)?n-o:null;let a=sn(G.bankAmount);G.bankDifference=Number.isFinite(o)&&!Number.isNaN(a)?o-a:null,G.plugEnabled=G.bankDifference!=null&&Math.abs(G.bankDifference)>=.5,t?fe():Na()}function Na(){if(re.activeStepId!==3)return;let e=(o,a)=>{let s=document.getElementById(o);s&&(s.value=a)};e("pr-data-total-value",le(G.prDataTotal)),e("clean-total-value",le(G.cleanTotal)),e("recon-diff-value",le(G.reconDifference)),e("bank-clean-total-value",le(G.cleanTotal)),e("bank-diff-value",G.bankDifference!=null?le(G.bankDifference):"---");let t=document.getElementById("bank-diff-hint");t&&(t.textContent=G.bankDifference==null?"":Math.abs(G.bankDifference)<.5?"Difference within acceptable tolerance.":"Difference exceeds tolerance and should be resolved.");let n=document.getElementById("bank-plug-btn");n&&(n.disabled=!G.plugEnabled)}function kt(e={},{rerender:t=!0}={}){Object.assign(Se,e),t&&fe()}async function Wn(){if(!pe()){et({loading:!1,lastError:"Excel runtime is unavailable.",prDataTotal:null,cleanTotal:null});return}et({loading:!0,lastError:null});try{let e="";await Excel.run(async n=>{let o=await Ne(n);if(console.log("DEBUG - Config table found:",!!o),o){let a=o.getDataBodyRange();a.load("values"),await n.sync();let s=a.values||[];console.log("DEBUG - Config table rows:",s.length),console.log("DEBUG - Looking for payroll date aliases:",ot),console.log("DEBUG - CONFIG_COLUMNS.FIELD:",B.FIELD,"CONFIG_COLUMNS.VALUE:",B.VALUE);for(let l of s){let c=String(l[B.FIELD]||"").trim(),r=l[B.VALUE],i=ot.some(f=>c===f||ce(c)===ce(f));if((c.toLowerCase().includes("payroll")||c.toLowerCase().includes("date"))&&console.log("DEBUG - Potential date field:",c,"=",r,"| isMatch:",i),i){let f=l[B.VALUE];console.log("DEBUG - Found payroll date field!",c,"raw value:",f),e=be(f)||"",console.log("DEBUG - Formatted payroll date:",e);break}}e||(console.warn("DEBUG - No payroll date found in config! Available fields:"),s.forEach((l,c)=>{console.log(`  Row ${c}: Field="${l[B.FIELD]}" Value="${l[B.VALUE]}"`)}))}else console.warn("DEBUG - Config table not found!")}),console.log("DEBUG prepareValidationData - Final Payroll Date:",e||"(empty)");let t=await Excel.run(async n=>{var A;let o=n.workbook.worksheets.getItem(N.DATA),a=n.workbook.worksheets.getItem(N.EXPENSE_MAPPING),s=n.workbook.worksheets.getItem(N.DATA_CLEAN),l=o.getUsedRangeOrNullObject(),c=a.getUsedRangeOrNullObject(),r=s.getUsedRangeOrNullObject();l.load("values"),c.load("values"),r.load(["address","rowCount"]),await n.sync();let i=l.isNullObject?[]:l.values||[],f=c.isNullObject?[]:c.values||[];console.log("DEBUG prepareValidationData - PR_Data rows:",i.length),console.log("DEBUG prepareValidationData - PR_Data headers:",i[0]),console.log("DEBUG prepareValidationData - PR_Expense_Mapping rows:",f.length);let d=((A=f[0])==null?void 0:A.map(m=>ve(m)))||[],u=m=>d.findIndex(m),p={category:u(m=>m.includes("category")),accountNumber:u(m=>m.includes("account")&&(m.includes("number")||m.includes("#"))),accountName:u(m=>m.includes("account")&&m.includes("name")),expenseReview:u(m=>m.includes("expense")&&m.includes("review"))},h=new Map;f.slice(1).forEach(m=>{var P,H,Q;let D=p.category>=0?Zt(m[p.category]):"";D&&h.set(D,{accountNumber:p.accountNumber>=0&&(P=m[p.accountNumber])!=null?P:"",accountName:p.accountName>=0&&(H=m[p.accountName])!=null?H:"",expenseReview:p.expenseReview>=0&&(Q=m[p.expenseReview])!=null?Q:""})});let v=s.getRangeByIndexes(0,0,1,8);v.load("values"),await n.sync();let w=v.values[0]||[],y=w.map(m=>ve(m));console.log("DEBUG prepareValidationData - PR_Data_Clean headers:",w),console.log("DEBUG prepareValidationData - PR_Data_Clean normalized:",y),console.log("DEBUG - PR_Data_Clean headers:",w),console.log("DEBUG - PR_Data_Clean normalized headers:",y);let g=y.findIndex(m=>(m.includes("payroll")||m.includes("period"))&&m.includes("date"));console.log("DEBUG - payrollDate column index:",g),g===-1&&(console.warn("DEBUG - No payroll date column found! Looking for header containing 'payroll'/'period' AND 'date'"),y.forEach((m,D)=>console.log(`  Col ${D}: "${m}"`)));let E={payrollDate:g,employee:y.findIndex(m=>m.includes("employee")),department:He(y),payrollCategory:y.findIndex(m=>m.includes("payroll")&&m.includes("category")),accountNumber:y.findIndex(m=>m.includes("account")&&(m.includes("number")||m.includes("#"))),accountName:y.findIndex(m=>m.includes("account")&&m.includes("name")),amount:y.findIndex(m=>m.includes("amount")),expenseReview:y.findIndex(m=>m.includes("expense")&&m.includes("review"))};console.log("DEBUG prepareValidationData - fieldMap:",E);let _=w.length,$=[],M=0,S=0;if(i.length>=2){let m=i[0],D=m.map(F=>ve(F));console.log("DEBUG prepareValidationData - Normalized headers:",D);let P=D.findIndex(F=>F.includes("employee")),H=He(D);console.log("DEBUG prepareValidationData - Employee column index:",P,"searching for 'employee' in:",D[6]),console.log("DEBUG prepareValidationData - Department column index:",H);let Q=h.size>0,J=D.reduce((F,q,ne)=>{if(ne===P||ne===H||!q||q.includes("total")||q.includes("gross")||q.includes("date")||q.includes("period"))return F;let ke=Zt(m[ne]||q);return Q&&!h.has(ke)||F.push(ne),F},[]);console.log("DEBUG prepareValidationData - Numeric columns:",J.length,J);for(let F=1;F<i.length;F+=1){let q=i[F],ne=P>=0?me(q[P]):"";if(!ne||ne.toLowerCase().includes("total"))continue;let ke=H>=0&&q[H]||"";J.forEach(oe=>{let Z=q[oe],ee=Number(Z);if(!Number.isFinite(ee)||ee===0)return;M+=ee;let ie=m[oe]||D[oe]||`Column ${oe+1}`,ae=h.get(Zt(ie))||{};S+=ee;let te=new Array(_).fill("");E.payrollDate>=0?te[E.payrollDate]=e:_>0&&(te[0]=e),$.length===0&&(console.log("DEBUG - Building first PR_Data_Clean row:"),console.log("  payrollDate value:",e),console.log("  fieldMap.payrollDate:",E.payrollDate),console.log("  Writing to column index:",E.payrollDate>=0?E.payrollDate:0)),E.employee>=0&&(te[E.employee]=ne),E.department>=0&&(te[E.department]=ke),E.payrollCategory>=0&&(te[E.payrollCategory]=ie),E.accountNumber>=0&&(te[E.accountNumber]=ae.accountNumber||""),E.accountName>=0&&(te[E.accountName]=ae.accountName||""),E.amount>=0&&(te[E.amount]=ee),E.expenseReview>=0&&(te[E.expenseReview]=ae.expenseReview||""),$.push(te)})}}if(console.log("DEBUG prepareValidationData - Clean rows generated:",$.length),console.log("DEBUG prepareValidationData - PR_Data total:",M,"Clean total:",S),console.log("DEBUG prepareValidationData - columnCount:",_,"cleanRange.address:",r.address),!r.isNullObject&&r.address){console.log("DEBUG prepareValidationData - Clearing data rows...");let m=Math.max(0,(r.rowCount||0)-1),D=Math.max(1,m);s.getRangeByIndexes(1,0,D,_).clear(),await n.sync(),console.log("DEBUG prepareValidationData - Data rows cleared")}if(console.log("DEBUG prepareValidationData - About to write",$.length,"rows with",_,"columns"),$.length>0){let m=s.getRangeByIndexes(1,0,$.length,_);m.values=$,console.log("DEBUG prepareValidationData - Data written to PR_Data_Clean")}else console.log("DEBUG prepareValidationData - No rows to write!");return await n.sync(),{prDataTotal:M,cleanTotal:S}});et({loading:!1,lastError:null,prDataTotal:t.prDataTotal,cleanTotal:t.cleanTotal})}catch(e){console.warn("Validate & Reconcile: unable to prepare PR_Data_Clean",e),et({loading:!1,prDataTotal:null,cleanTotal:null,lastError:"Unable to prepare reconciliation data. Try again."})}}function Oa(e){let t={activeCount:0,departmentCount:0,employeeMap:new Map};if(!e||!e.length)return t;let{headers:n,dataStartIndex:o}=io(e,["employee"]);if(!n.length||o==null)return t;let a=ro(n),s=n.findIndex(r=>r.includes("termination")),l=He(n);if(a===-1)return t;let c=new Set;for(let r=o;r<e.length;r+=1){let i=e[r],f=i[a],d=so(f);if(!d||oo(d))continue;let u=s>=0?i[s]:"",p=l>=0?i[l]:"";!me(u)&&!c.has(d)&&(c.add(d),t.activeCount+=1),p&&(t.departmentCount+=1),t.employeeMap.has(d)||t.employeeMap.set(d,{name:me(f)||d,department:me(p),termination:u})}return t}function Ta(e){let t={totalEmployees:0,departmentCount:0,employeeMap:new Map};if(!e||!e.length)return t;let{headers:n,dataStartIndex:o}=io(e,["employee"]);if(!n.length||o==null)return t;let a=ro(n),s=He(n);if(a===-1)return t;let l=new Set;for(let c=o;c<e.length;c+=1){let r=e[c],i=r[a],f=so(i);if(!f||oo(f))continue;l.has(f)||(l.add(f),t.totalEmployees+=1);let d=s>=0?r[s]:"";d&&(t.departmentCount+=1),t.employeeMap.has(f)||t.employeeMap.set(f,{name:me(i)||f,department:me(d)})}return t}function ve(e){return me(e).toLowerCase()}function so(e){return me(e).toLowerCase()}function ro(e=[]){let t=e.findIndex(o=>o.includes("employee")&&o.includes("name"));return t>=0?t:e.findIndex(o=>o.includes("employee"))}function io(e,t=[]){let n=[],o=null;return(e||[]).some((a,s)=>{let l=(a||[]).map(ve);return t.every(r=>l.some(i=>i.includes(r)))?(n=l,o=s,!0):!1}),{headers:n,dataStartIndex:o!=null?o+1:null}}function me(e){return e==null?"":String(e).trim()}function Zt(e){return me(e).toLowerCase()}function He(e=[]){let t=e.map((l,c)=>({idx:c,value:ve(l)})),n=t.find(({value:l})=>l.includes("department")&&l.includes("description"));if(n)return console.log("DEBUG pickDepartmentIndex - Using 'Department Description' at index:",n.idx),n.idx;let o=t.find(({value:l})=>l.includes("department")&&l.includes("name"));if(o)return console.log("DEBUG pickDepartmentIndex - Using 'Department Name' at index:",o.idx),o.idx;let a=t.find(({value:l})=>l.includes("department")&&!l.includes("id")&&!l.includes("#")&&!l.includes("code")&&!l.includes("number"));if(a)return console.log("DEBUG pickDepartmentIndex - Using non-ID department at index:",a.idx),a.idx;let s=t.find(({value:l})=>l.includes("department"));return s&&console.log("DEBUG pickDepartmentIndex - Using fallback department at index:",s.idx),s?s.idx:-1}function Yn(e,t,n={}){if(en(),!t||!t.length)return;let o=document.createElement("div");o.className="pf-modal";let a=t.filter(r=>r.type==="missing_from_payroll"),s=t.filter(r=>r.type==="missing_from_roster"),l=t.filter(r=>!r.type),c="";if(a.length>0&&(c+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading pf-mismatch-warning">
                    <span class="pf-mismatch-icon">\u26A0\uFE0F</span>
                    In Roster but NOT in Payroll (${a.length})
                </h4>
                <p class="pf-mismatch-subtext">These employees appear in your centralized roster but were not found in the payroll data. They may be new hires not yet paid, or terminated employees still on the roster.</p>
                <div class="pf-mismatch-tiles">
                    ${a.map(r=>`
                        <div class="pf-mismatch-tile pf-mismatch-missing-payroll">
                            <span class="pf-mismatch-name">${x(r.name)}</span>
                            <span class="pf-mismatch-detail">${x(r.department)}</span>
                        </div>
                    `).join("")}
                </div>
            </div>
        `),s.length>0&&(c+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading pf-mismatch-alert">
                    <span class="pf-mismatch-icon">\u{1F534}</span>
                    In Payroll but NOT in Roster (${s.length})
                </h4>
                <p class="pf-mismatch-subtext">These employees appear in payroll data but are not in the centralized roster. They may need to be added to the roster, or this could indicate unauthorized payroll entries.</p>
                <div class="pf-mismatch-tiles">
                    ${s.map(r=>`
                        <div class="pf-mismatch-tile pf-mismatch-missing-roster">
                            <span class="pf-mismatch-name">${x(r.name)}</span>
                            <span class="pf-mismatch-detail">${x(r.department)}</span>
                        </div>
                    `).join("")}
                </div>
            </div>
        `),l.length>0){let r=n.formatter||(i=>typeof i=="string"?{name:i,source:"",isMissingFromTarget:!0}:i);c+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading">
                    <span class="pf-mismatch-icon">\u{1F4CB}</span>
                    ${x(n.label||e)} (${l.length})
                </h4>
                <div class="pf-mismatch-tiles">
                    ${l.map(i=>{let f=r(i);return`
                            <div class="pf-mismatch-tile">
                                <span class="pf-mismatch-name">${x(f.name||f.employee||"")}</span>
                                <span class="pf-mismatch-detail">${x(f.source||`${f.rosterDept||""} \u2192 ${f.payrollDept||""}`)}</span>
                            </div>
                        `}).join("")}
                </div>
            </div>
        `}c||(c='<p class="pf-mismatch-empty">No differences found.</p>'),o.innerHTML=`
        <div class="pf-modal-content pf-headcount-modal">
            <div class="pf-modal-header">
                <h3>${x(e)}</h3>
                <button type="button" class="pf-modal-close" data-modal-close>&times;</button>
            </div>
            <div class="pf-modal-body">
                ${c}
            </div>
            <div class="pf-modal-footer">
                <span class="pf-modal-summary">${t.length} total difference${t.length!==1?"s":""} found</span>
                <button type="button" class="pf-modal-close-btn" data-modal-close>Close</button>
            </div>
        </div>
    `,o.addEventListener("click",r=>{r.target===o&&en()}),o.querySelectorAll("[data-modal-close]").forEach(r=>{r.addEventListener("click",en)}),document.body.appendChild(o)}function en(){var e;(e=document.querySelector(".pf-modal"))==null||e.remove()}function Rt(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=Pt(),n=document.getElementById("step-notes-input"),o=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!o;let a=document.getElementById("headcount-notes-hint");a&&(a.textContent=t?"Please document outstanding differences before signing off.":""),z.skipAnalysis&&nn()}function La(){var n;let e=Pt(),t=((n=document.getElementById("step-notes-input"))==null?void 0:n.value.trim())||"";if(e&&!t){window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");return}on({statusText:"Headcount Review signed off."})}function nn(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(wt)?t.slice(wt.length).replace(/^\s+/,""):t.replace(new RegExp(`^${wt}\\s*`,"i"),"").trimStart(),o=wt+(n?`
${n}`:"");if(e.value!==o){e.value=o;let a=_e(2);a&&V(a.note,o)}}function qn(e){let t=e!=null&&e.target&&e.target instanceof HTMLInputElement?e.target:document.getElementById("bank-amount-input"),n=sn(t==null?void 0:t.value),o=ao(n);t&&(t.value=o),et({bankAmount:n},{rerender:!1})}function Ma(){let e=we.findIndex(t=>t.id===3);e!==-1&&at(Math.min(we.length-1,e+1))}function Ba(){let e=we.findIndex(t=>t.id===4);e!==-1&&at(Math.min(we.length-1,e+1))}async function Fa(){if(console.log("[Archive] handleArchiveRun called"),!pe()){window.alert("Excel runtime is unavailable.");return}if(!window.confirm(`Archive Payroll Run

This will:
1. Update PR_Archive_Summary with current period
2. Clear working data from all payroll sheets
3. Clear non-permanent notes and config values

Note: Please manually save a copy of your workbook before archiving.

Make sure you've completed all review steps before archiving.

Continue?`)){console.log("[Archive] User cancelled");return}console.log("[Archive] User confirmed, starting archive process...");try{if(console.log("[Archive] Step 1: Creating archive workbook..."),!await Va()){console.log("[Archive] Archive cancelled or failed");return}console.log("[Archive] Step 1 complete: Archive workbook created/user confirmed backup"),console.log("[Archive] Step 2: Updating PR_Archive_Summary..."),await Ua(),console.log("[Archive] Step 2 complete: Archive summary updated"),console.log("[Archive] Step 3: Clearing working data..."),await Ga(),console.log("[Archive] Step 3 complete: Working data cleared"),console.log("[Archive] Step 4: Clearing non-permanent notes..."),await za(),console.log("[Archive] Step 4 complete: Notes cleared"),console.log("[Archive] Step 5: Resetting config values..."),await Ja(),console.log("[Archive] Step 5 complete: Config reset"),console.log("[Archive] Archive workflow complete!"),await eo(),fe(),window.alert(`\u2705 Archive Complete!

\u2713 Archive workbook created (please save it if you haven't)
\u2713 PR_Archive_Summary updated with current period
\u2713 Working data cleared from PR_Data, PR_Data_Clean, PR_JE_Draft
\u2713 Notes and config reset

Ready for next payroll cycle!`)}catch(t){console.error("[Archive] Error during archive:",t),window.alert(`Archive Error

An error occurred during the archive process:
`+t.message+`

Please check the console for details and verify your data.`)}}async function Va(){try{let t=`Payroll_Archive_${Dt()||new Date().toISOString().split("T")[0]}`;return console.log("[Archive] Exporting payroll data to CSV..."),await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name"),await n.sync();let s=[{name:N.JE_DRAFT,label:"Journal Entry"},{name:N.DATA_CLEAN,label:"Payroll Detail"},{name:N.DATA,label:"Raw Import"}],l=[];for(let c of s){let r=a.items.find(f=>f.name===c.name);if(!r){console.log(`[Archive] Sheet not found: ${c.name}`);continue}let i=r.getUsedRangeOrNullObject();i.load("values"),await n.sync(),!i.isNullObject&&i.values&&i.values.length>0&&(l.push({sheetName:c.name,label:c.label,values:i.values}),console.log(`[Archive] Collected: ${c.name} (${i.values.length} rows)`))}if(l.length===0)return window.alert("No data to archive. Please complete the payroll workflow first."),!1;for(let c of l){let r=ja(c.values),i=`${t}_${c.sheetName}.csv`;Ha(r,i),console.log(`[Archive] Downloaded: ${i}`)}return window.alert(`\u{1F4E5} Archive Files Downloaded!

${l.length} CSV file(s) have been downloaded:

`+l.map(c=>`\u2022 ${t}_${c.sheetName}.csv`).join(`
`)+`

Please save these files to your archive folder.

Click OK to continue clearing this workbook for the next period.`),!0})}catch(e){return console.error("[Archive] Error exporting archive:",e),window.alert(`Archive Export Error

`+e.message+`

Please manually save a copy of this workbook before continuing.`),window.confirm(`Do you want to continue with clearing the data?

Make sure you have saved a backup first!`)}}function ja(e){return e.map(t=>t.map(n=>{if(n==null)return"";let o=String(n);return(o.includes(",")||o.includes('"')||o.includes(`
`))&&(o='"'+o.replace(/"/g,'""')+'"'),o}).join(",")).join(`
`)}function Ha(e,t){let n=new Blob([e],{type:"text/csv;charset=utf-8;"}),o=URL.createObjectURL(n),a=document.createElement("a");a.setAttribute("href",o),a.setAttribute("download",t),a.style.visibility="hidden",document.body.appendChild(a),a.click(),document.body.removeChild(a),setTimeout(()=>URL.revokeObjectURL(o),100)}async function Ua(){await Excel.run(async e=>{let t=e.workbook.worksheets.getItemOrNullObject(N.ARCHIVE_SUMMARY),n=e.workbook.worksheets.getItemOrNullObject(N.DATA_CLEAN);if(t.load("isNullObject"),n.load("isNullObject"),await e.sync(),t.isNullObject){console.warn("[Archive] PR_Archive_Summary not found - skipping");return}if(n.isNullObject){console.warn("[Archive] PR_Data_Clean not found - skipping");return}let o=n.getUsedRangeOrNullObject();if(o.load("values"),await e.sync(),o.isNullObject||!o.values||o.values.length<2){console.warn("[Archive] PR_Data_Clean is empty - skipping archive summary update");return}let a=(o.values[0]||[]).map(m=>String(m||"").toLowerCase().trim()),s=o.values.slice(1),l=a.findIndex(m=>m.includes("amount")),c=a.findIndex(m=>m.includes("employee")),r=a.findIndex(m=>m.includes("payroll")&&m.includes("date")||m.includes("pay period")||m==="date"),i=0,f=new Set,d=Dt()||"";s.forEach(m=>{l>=0&&(i+=Number(m[l])||0),c>=0&&m[c]&&f.add(String(m[c]).trim()),r>=0&&m[r]&&!d&&(d=String(m[r]))});let u=f.size;console.log(`[Archive] Current period: Date=${d}, Total=${i}, Employees=${u}`);let p=t.getUsedRangeOrNullObject();p.load("values,rowCount"),await e.sync();let h=[],v=[];!p.isNullObject&&p.values&&p.values.length>0&&(h=p.values[0],v=p.values.slice(1)),h.length===0&&(h=["Pay Period","Total Payroll","Employee Count","Archived Date"],t.getRange("A1:D1").values=[h],await e.sync());let w=h.map(m=>String(m||"").toLowerCase().trim()),y=w.findIndex(m=>m.includes("pay period")||m.includes("period")||m==="date"),g=w.findIndex(m=>m.includes("total")),E=w.findIndex(m=>m.includes("employee")||m.includes("count")),_=w.findIndex(m=>m.includes("archived")),$=new Array(h.length).fill("");y>=0&&($[y]=d),g>=0&&($[g]=i),E>=0&&($[E]=u),_>=0&&($[_]=new Date().toISOString().split("T")[0]),v.length>=5&&(v=v.slice(0,4),console.log("[Archive] Trimmed archive to 4 periods, adding current")),v.unshift($);let M=2,S=M+5;if(t.getRange(`A${M}:${String.fromCharCode(64+h.length)}${S}`).clear(Excel.ClearApplyTo.contents),await e.sync(),v.length>0){let m=t.getRange(`A${M}:${String.fromCharCode(64+h.length)}${M+v.length-1}`);m.values=v,await e.sync()}console.log(`[Archive] Archive summary updated with ${v.length} periods`)})}async function Ga(){let e=[N.DATA,N.DATA_CLEAN,N.EXPENSE_REVIEW,N.JE_DRAFT];await Excel.run(async t=>{for(let n of e){let o=t.workbook.worksheets.getItemOrNullObject(n);if(o.load("isNullObject"),await t.sync(),o.isNullObject){console.log(`[Archive] Sheet not found: ${n}`);continue}let a=o.getUsedRangeOrNullObject();if(a.load("rowCount,columnCount,address"),await t.sync(),a.isNullObject||a.rowCount<=1){console.log(`[Archive] Sheet empty or headers only: ${n}`);continue}if(o.getRange(`A2:${String.fromCharCode(64+a.columnCount)}${a.rowCount}`).clear(Excel.ClearApplyTo.contents),n===N.EXPENSE_REVIEW){let l=o.charts;l.load("items"),await t.sync();for(let c=l.items.length-1;c>=0;c--)l.items[c].delete()}await t.sync(),console.log(`[Archive] Cleared data from: ${n}`)}})}async function za(){await Excel.run(async e=>{let t=await Ne(e);if(!t){console.warn("[Archive] Config table not found");return}let n=t.getDataBodyRange();n.load("values,rowCount"),await e.sync();let o=n.values||[],a=0,s=Object.values(Pe).map(l=>l.note);for(let l=0;l<o.length;l++){let c=String(o[l][B.FIELD]||"").trim(),r=String(o[l][B.PERMANENT]||"").toUpperCase();s.includes(c)&&r!=="Y"&&(n.getCell(l,B.VALUE).values=[[""]],a++)}await e.sync(),console.log(`[Archive] Cleared ${a} non-permanent notes`)})}async function Ja(){let e=["PR_Payroll_Date","PR_Accounting_Period","PR_Journal_Entry_ID","Payroll_Date","Accounting_Period","Journal_Entry_ID","JE_Transaction_ID",...Object.values(Pe).map(t=>t.signOff),...Object.values(Pe).map(t=>t.reviewer),...Object.values(de)];await Excel.run(async t=>{let n=await Ne(t);if(!n){console.warn("[Archive] Config table not found");return}let o=n.getDataBodyRange();o.load("values,rowCount"),await t.sync();let a=o.values||[],s=0;for(let l=0;l<a.length;l++){let c=String(a[l][B.FIELD]||"").trim(),r=String(a[l][B.PERMANENT]||"").toUpperCase();e.some(f=>ce(f)===ce(c))&&r!=="Y"&&(o.getCell(l,B.VALUE).values=[[""]],s++)}await t.sync(),console.log(`[Archive] Reset ${s} non-permanent config values`),Object.keys(Y.values).forEach(l=>{e.some(c=>ce(c)===ce(l))&&(Y.values[l]="")})})}async function Wa(){if(!pe()){window.alert("Excel runtime is unavailable.");return}j.loading=!0,j.lastError=null,Qn(!1),fe();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(N.JE_DRAFT),o=t.workbook.worksheets.getItemOrNullObject(N.DATA_CLEAN),a=n.getUsedRangeOrNullObject();a.load("values"),o.load("name"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error(`${N.JE_DRAFT} is empty.`);let l=(s[0]||[]).map(u=>ve(u)),c=l.findIndex(u=>u.includes("debit")),r=l.findIndex(u=>u.includes("credit"));if(c===-1||r===-1)throw new Error("Debit/Credit columns not found in JE Draft.");let i=0,f=0;s.slice(1).forEach(u=>{i+=Number(u[c])||0,f+=Number(u[r])||0});let d=0;if(!o.isNullObject){let u=o.getUsedRangeOrNullObject();if(u.load("values"),await t.sync(),!u.isNullObject&&u.values&&u.values.length>1){let h=(u.values[0]||[]).map(v=>ve(v)).findIndex(v=>v.includes("amount"));h>=0&&u.values.slice(1).forEach(v=>{d+=Number(v[h])||0})}}return{debitTotal:i,creditTotal:f,difference:f-i,cleanTotal:d}});Object.assign(j,e,{lastError:null})}catch(e){console.warn("JE summary:",e),j.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",j.debitTotal=null,j.creditTotal=null,j.difference=null,j.cleanTotal=null}finally{j.loading=!1,fe()}}async function Ya(){try{let e=Number.isFinite(Number(j.debitTotal))?j.debitTotal:"",t=Number.isFinite(Number(j.creditTotal))?j.creditTotal:"",n=Number.isFinite(Number(j.difference))?j.difference:"";await Promise.all([nt(Bo,String(e)),nt(Fo,String(t)),nt(Vo,String(n))]),Qn(!0)}catch(e){console.error("JE save:",e)}}async function qa(){if(!pe()){window.alert("Excel runtime is unavailable.");return}j.loading=!0,j.lastError=null,fe();try{await Excel.run(async e=>{let t="",n="",o=await Ne(e);if(o){let g=o.getDataBodyRange();g.load("values"),await e.sync();let E=g.values||[];for(let _ of E){let $=String(_[B.FIELD]||"").trim(),M=_[B.VALUE];($==="Journal_Entry_ID"||$==="JE_Transaction_ID"||$==="PR_Journal_Entry_ID")&&(t=String(M||"").trim()),ot.some(S=>$===S||ce($)===ce(S))&&(n=be(M)||"")}}console.log("JE Generation - RefNumber:",t,"TxnDate:",n);let a=e.workbook.worksheets.getItemOrNullObject(N.DATA_CLEAN);if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PR_Data_Clean sheet not found. Run Validate & Reconcile first.");let s=a.getUsedRangeOrNullObject();if(s.load("values"),await e.sync(),s.isNullObject)throw new Error("PR_Data_Clean is empty. Run Validate & Reconcile first.");let l=s.values||[];if(l.length<2)throw new Error("PR_Data_Clean has no data rows.");let c=l[0].map(g=>ve(g));console.log("JE Generation - PR_Data_Clean headers:",c);let r={accountNumber:c.findIndex(g=>g.includes("account")&&(g.includes("number")||g.includes("#"))),accountName:c.findIndex(g=>g.includes("account")&&g.includes("name")),amount:c.findIndex(g=>g.includes("amount")),department:He(c),payrollCategory:c.findIndex(g=>g.includes("payroll")&&g.includes("category")),employee:c.findIndex(g=>g.includes("employee"))};if(console.log("JE Generation - Column indices:",r),r.amount===-1)throw new Error("Amount column not found in PR_Data_Clean.");let i=new Map;for(let g=1;g<l.length;g++){let E=l[g],_=r.accountNumber>=0?String(E[r.accountNumber]||"").trim():"",$=r.accountName>=0?String(E[r.accountName]||"").trim():"",M=Number(E[r.amount])||0,S=r.department>=0?String(E[r.department]||"").trim():"";if(M===0)continue;let A=`${_}|${S}`;if(i.has(A)){let m=i.get(A);m.amount+=M}else i.set(A,{accountNumber:_,accountName:$,department:S,amount:M})}console.log("JE Generation - Aggregated into",i.size,"unique Account+Department combinations");let f=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],d=[f],u=0,p=0,h=Array.from(i.values()).sort((g,E)=>{let _=String(g.accountNumber).localeCompare(String(E.accountNumber));return _!==0?_:String(g.department).localeCompare(String(E.department))});for(let g of h){let{accountNumber:E,accountName:_,department:$,amount:M}=g,S=M>0?M:0,A=M<0?Math.abs(M):0,m=[_,$].filter(Boolean).join(" - ");u+=S,p+=A,d.push([t,n,E,_,M,S||"",A||"",m,$])}console.log("JE Generation - Built",d.length-1,"summarized journal lines"),console.log("JE Generation - Total Debit:",u,"Total Credit:",p);let v=e.workbook.worksheets.getItemOrNullObject(N.JE_DRAFT);if(v.load("isNullObject"),await e.sync(),v.isNullObject)v=e.workbook.worksheets.add(N.JE_DRAFT),await e.sync();else{let g=v.getUsedRangeOrNullObject();g.load("address"),await e.sync(),g.isNullObject||(g.clear(),await e.sync())}let w=v.getRangeByIndexes(0,0,d.length,f.length);w.values=d,await e.sync();try{let g=d.length-1,E=v.getRange("A1:I1");Ln(E),g>0&&(Mn(v,1,g),bt(v,4,g),bt(v,5,g),bt(v,6,g)),v.getRange("A:I").format.autofitColumns(),await e.sync()}catch(g){console.warn("JE formatting error (non-critical):",g)}v.activate(),v.getRange("A1").select(),await e.sync();let y=0;for(let g=1;g<l.length;g++)y+=Number(l[g][r.amount])||0;j.debitTotal=u,j.creditTotal=p,j.difference=p-u,j.cleanTotal=y}),j.loading=!1,j.lastError=null,fe()}catch(e){console.error("JE Generation failed:",e),j.loading=!1,j.lastError=e.message||"Failed to generate journal entry.",fe()}}async function Ka(){if(!pe()){window.alert("Excel runtime is unavailable.");return}try{let{rows:e}=await Excel.run(async n=>{let a=n.workbook.worksheets.getItem(N.JE_DRAFT).getUsedRangeOrNullObject();a.load("values"),await n.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error(`${N.JE_DRAFT} is empty.`);return{rows:s}}),t=Ra(e);Sa(`pr-je-draft-${st()}.csv`,t)}catch(e){console.warn("JE export:",e),window.alert("Unable to export the JE draft. Confirm the sheet has data.")}}})();
//# sourceMappingURL=app.bundle.js.map
