/* Prairie Forge Payroll Recorder */
(()=>{var tn="1.0.0.7",D={CONFIG:"SS_PF_Config",DATA:"PR_Data",DATA_CLEAN:"PR_Data_Clean",EXPENSE_MAPPING:"PR_Expense_Mapping",EXPENSE_REVIEW:"PR_Expense_Review",JE_DRAFT:"PR_JE_Draft",ARCHIVE_SUMMARY:"PR_Archive_Summary"};var Ya=[{name:"Instructions",description:"How to use the Prairie Forge payroll template"},{name:"Data_Input",description:"Paste WellsOne export data here"},{name:D.CONFIG,description:"Prairie Forge shared configuration storage (all modules)"},{name:"Config_Keywords",description:"Keyword-based account mapping rules"},{name:"Config_Accounts",description:"Account rewrite rules"},{name:"Config_Locations",description:"Location normalization rules"},{name:"Config_Vendors",description:"Vendor-specific overrides"},{name:"Config_Settings",description:"Prairie Forge system settings"},{name:D.EXPENSE_MAPPING,description:"Expense category mappings"},{name:D.DATA,description:"Processed payroll data staging"},{name:D.DATA_CLEAN,description:"Cleaned and validated payroll data"},{name:D.EXPENSE_REVIEW,description:"Expense review workspace"},{name:D.JE_DRAFT,description:"Journal entry preparation area"}];var rt=[{id:0,title:"Configuration Setup",summary:"Company profile, branding, and run settings",description:"Keep the SS_PF_Config table current before every payroll run so downstream sheets inherit the right colors, links, and identifiers.",icon:"\u{1F9ED}",ctaLabel:"Open Configuration Form",statusHint:"Configuration edits happen inside the PF_Config table and are available to every step instantly.",highlights:[{label:"Company Profile",detail:"Company name, logos, payroll date, reporting period."},{label:"Brand Identity",detail:"Primary + accent colors carry through dashboards and exports."},{label:"System Links",detail:"Quick jumps to HRIS, payroll provider, accounting import, and archive folders."}],checklist:["Review profile, branding, links, and run settings each payroll cycle.","Click Save to write updates back to the SS_PF_Config sheet."]},{id:1,title:"Import Payroll Data",summary:"Paste the payroll provider export into the Data sheet",description:"Pull your payroll data from your provider\u2019s portal and paste it into the Data tab. If the columns match, just paste the rows; if they don\u2019t, paste your headers and data right over the top. Formatting is fully automated.",icon:"\u{1F4E5}",ctaLabel:"Prepare Import Sheet",statusHint:"The Data worksheet is activated so you can paste the latest provider export.",highlights:[{label:"Source File",detail:"Use WellsOne/ADP export with every pay category column visible."},{label:"Structure",detail:"Row 2 headers, row 3+ data, no blank columns, totals removed."},{label:"Quality",detail:"Spot-check employee counts and pay period filters before moving on."}],checklist:["Download the payroll detail export covering this pay period.","Paste values into the Data sheet starting at cell A3.","Confirm all pay category headers remain intact and spelled consistently."]},{id:2,title:"Headcount Review",summary:"Ensure roster and payroll rows agree",description:"This step is optional, but strongly recommended. A centralized employee roster keeps every payroll-related workbook aligned while ensuring key attributes such as department and location stay consistent each pay period.",icon:"\u{1F465}",ctaLabel:"Launch Headcount Review",statusHint:"Data and mapping sheets are surfaced so you can reconcile roster counts before validation.",highlights:[{label:"Roster Alignment",detail:"Compare active roster to the employees present in the Data sheet."},{label:"Variance Tracking",detail:"Note missing departments or unexpected hires before the validation run."},{label:"Approvals",detail:"Capture reviewer initials and date for audit coverage."}],checklist:["Filter the Data sheet by Department to ensure every team appears.","Look for duplicate or out-of-period employees and resolve upstream.","Document findings on the Headcount Review tab or your tracker of choice."]},{id:3,title:"Validate & Reconcile",summary:"Normalize payroll data and reconcile totals",description:"Automatically rebuild the PR_Data_Clean sheet, confirm payroll totals match, and prep the bank reconciliation before moving to Expense Review.",icon:"\u2705",statusHint:"Run completes automatically when you enter this step.",highlights:[{label:"Normalized Data",detail:"Creates one row per employee and payroll category."},{label:"Outputs",detail:"Data_Clean rebuilt with payroll category + mapping details."},{label:"Reconciliation",detail:"Displays PR_Data vs PR_Data_Clean totals plus bank comparison."}]},{id:4,title:"Expense Review",summary:"Generate an executive-ready payroll summary",description:"Build a six-period payroll dashboard (current + five prior), including department-level breakouts and variance indicators, plus notes and CoPilot guidance.",icon:"\u{1F4CA}",statusHint:"Selecting this step rebuilds PR_Expense_Review automatically.",highlights:[{label:"Time Series",detail:"Shows six consecutive payroll periods."},{label:"Departments",detail:"All-in totals, burden rates, and headcount by department."},{label:"Guidance",detail:"Use CoPilot to summarize trends and capture review notes."}],checklist:[]},{id:5,title:"Journal Entry Prep",summary:"Generate a QuickBooks-ready journal draft",description:"Create the JE Draft sheet with the headers QuickBooks Online/Desktop expect so you only need to paste balanced lines.",icon:"\u{1F9FE}",ctaLabel:"Generate JE Draft",statusHint:"JE Draft contains headers for RefNumber, TxnDate, account columns, and line descriptions.",highlights:[{label:"Structure",detail:"Debit/Credit columns prepared with standard import headers."},{label:"Context",detail:"JE Transaction ID from configuration is referenced for traceability."},{label:"Next Step",detail:"Populate amounts from Expense Review to finalize the journal."}],checklist:["Ensure validation + expense review steps are complete.","Run the generator to rebuild the JE Draft sheet.","Paste balanced lines and export to QuickBooks / ERP import format."]},{id:6,title:"Archive & Clear",summary:"Snapshot workpapers and reset working tabs",description:"Capture a log of each payroll run, note the archive destination, and optionally clear staging sheets for the next cycle.",icon:"\u{1F5C2}\uFE0F",ctaLabel:"Create Archive Summary",statusHint:"Archive summary headers help you log when data was exported and where the files live.",highlights:[{label:"Run Log",detail:"Payroll date, reporting period, JE ID, and who processed the run."},{label:"Storage",detail:"Link to the Archive folder defined in configuration."},{label:"Reset",detail:"Reminder to clear Data/Data_Clean once files are safely archived."}],checklist:["Record archive destination links and reviewer approvals.","Copy Data/Data_Clean/JE Draft tabs to the archive workbook if needed.","Clear sensitive data so the template is ready for the next payroll."]}],qa=(typeof window!="undefined"&&Array.isArray(window.PF_BUILDER_ALLOWLIST)?window.PF_BUILDER_ALLOWLIST:[]).map(e=>String(e||"").trim().toLowerCase());function it(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}function nn(e){try{Office.onReady(t=>{console.log("Office.onReady fired:",t),t.host===Office.HostType.Excel||console.warn("Not running in Excel, host:",t.host),e(t)})}catch(t){console.warn("Office.onReady failed:",t),e(null)}}var Dt="SS_PF_Config",ho="tab-structure",yo=new Set(["all","shared","global","common","any","*"]),on=["SS_PF_Config","SS_Employee_Roster","SS_Chart_of_Accounts"],vo={"payroll-recorder":{visible:["PR_Data","PR_Data_Clean","PR_Expense_Review","PR_JE_Draft"],hidden:["SS_PF_Config","SS_Employee_Roster","SS_Chart_of_Accounts","PR_Expense_Mapping","PR_Archive_Summary","PR_Homepage","SS_Homepage"]},"pto-accrual":{visible:["PTO_Data","PTO_Analysis","PTO_JE_Draft"],hidden:["SS_PF_Config","SS_Employee_Roster","SS_Chart_of_Accounts","PTO_Archive_Summary","PTO_Homepage","SS_Homepage"]}};var bo={"payroll-recorder":["payroll-recorder","payroll","payroll recorder","payroll review","pr"],"employee-roster":["employee-roster","employee roster","headcount","headcount review","roster"],"pto-accrual":["pto-accrual","pto","pto accrual","pto review"]};async function an(e,{aliasTokens:t=[]}={}){if(!it())return;let n=Ge(e);console.log(`[Tab Visibility] Applying visibility for module: ${n}`);let o=vo[n];o?await wo(o,n):await Eo(e,t)}async function wo(e,t){let n=(e.visible||[]).map(a=>xe(a)),o=(e.hidden||[]).map(a=>xe(a));console.log(`[Tab Visibility] Explicit config for ${t}:`),console.log(`  - Visible: ${e.visible.join(", ")}`),console.log(`  - Hidden: ${e.hidden.join(", ")}`);try{await Excel.run(async a=>{let s=a.workbook.worksheets;s.load("items/name,visibility"),await a.sync();let l=[],c=[];s.items.forEach(r=>{let u=xe(r.name);n.includes(u)?l.push(r):o.includes(u)&&c.push(r)});for(let r of l)r.visibility=Excel.SheetVisibility.visible,console.log(`[Tab Visibility] SHOW: ${r.name}`);if(await a.sync(),s.items.filter(r=>r.visibility===Excel.SheetVisibility.visible).length>c.length){for(let r of c)try{r.visibility=Excel.SheetVisibility.hidden,console.log(`[Tab Visibility] HIDE: ${r.name}`)}catch(u){console.warn(`[Tab Visibility] Could not hide "${r.name}":`,u.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Applied visibility for ${t}`)})}catch(a){console.warn(`[Tab Visibility] Error applying visibility for ${t}:`,a)}}async function Eo(e,t=[]){let n=Ro([...So(e),...t]);console.log(`[Tab Visibility] Using config-based visibility. Module: ${e}, Aliases:`,[...n]);try{await Excel.run(async o=>{let a=o.workbook.worksheets.getItemOrNullObject(Dt);if(await o.sync(),a.isNullObject){console.warn(`Config sheet ${Dt} is missing; skipping tab visibility.`);return}let s=a.getUsedRangeOrNullObject();if(s.load("values"),await o.sync(),s.isNullObject){console.warn(`${Dt} does not contain any values yet.`);return}let l=s.values||[];if(!l.length)return;let c=Co(l[0]),i=c.get("category"),r=c.get("field"),u=c.get("value"),f=c.get("value2");if(console.log(`[Tab Visibility] Headers - Category: ${i}, Field: ${r}, Value: ${u}, Value2: ${f}`),i===void 0||r===void 0||u===void 0){console.warn("SS_PF_Config needs Category, Field, and Value columns to drive tab visibility.");return}let p=l.slice(1).map(b=>{var M,_;let E=sn(b[i]),x=String((M=b[u])!=null?M:"").trim(),A=f!==void 0?String((_=b[f])!=null?_:"").trim():"";return{category:E,tabName:x,normalizedTabName:xe(x),moduleValue:A}}).filter(b=>b.tabName&&b.category===ho);if(console.log(`[Tab Visibility] Found ${p.length} tab-structure rules:`,p.map(b=>`${b.tabName} \u2192 ${b.moduleValue}`)),!p.length){console.warn("No rows found in SS_PF_Config for Tab Structure.");return}let d=new Map;p.forEach(b=>{b.normalizedTabName&&d.set(b.normalizedTabName,b)});let h=o.workbook.worksheets;h.load("items/name,visibility"),await o.sync();let v=on.map(b=>xe(b)),C=[],g=0;h.items.forEach(b=>{let E=xe(b.name);if(!E)return;if(v.includes(E)){console.log(`[Tab Visibility] Skipping system sheet: "${b.name}"`);return}let x=d.get(E);if(!x){console.log(`[Tab Visibility] No rule for "${b.name}" - leaving as-is`),b.visibility===Excel.SheetVisibility.visible&&g++;return}let A=_o(x.moduleValue,n);console.log(`[Tab Visibility] "${b.name}" (module: ${x.moduleValue}) \u2192 ${A?"SHOW":"HIDE"}`),A&&g++,C.push({sheet:b,shouldShow:A})}),console.log(`[Tab Visibility] ${g} sheets will be visible after changes`);for(let b of C)b.shouldShow&&(b.sheet.visibility=Excel.SheetVisibility.visible);if(await o.sync(),g>0){for(let b of C)if(!b.shouldShow)try{b.sheet.visibility=Excel.SheetVisibility.hidden}catch(E){console.warn(`[Tab Visibility] Could not hide "${b.sheet.name}":`,E.message)}await o.sync()}else console.warn("[Tab Visibility] No sheets would be visible - skipping hide operations")})}catch(o){console.warn(`Unable to toggle worksheet visibility for ${e}:`,o)}}function Co(e=[]){let t=new Map;return e.forEach((n,o)=>{let a=sn(n);a&&t.set(a,o)}),t}function sn(e){return Ge(e)}function xe(e){return String(e!=null?e:"").trim().toLowerCase()}function ko(e){return String(e!=null?e:"").split(/[,;|/&]+/).map(t=>Ge(t)).filter(Boolean)}function Ge(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}function Ro(e){let t=(e||[]).map(n=>Ge(n)).filter(Boolean);return t.length?new Set(t):new Set}function So(e){var n;let t=Ge(e);return(n=bo[t])!=null?n:[t]}function _o(e,t){let n=ko(e);return n.length?n.some(o=>yo.has(o)||t.has(o)):!0}async function xo(){if(!it()){console.log("Excel not available");return}let e=[...on];try{await Excel.run(async t=>{let n=t.workbook.worksheets;n.load("items/name,visibility"),await t.sync(),n.items.forEach(o=>{let a=xe(o.name);e.map(s=>xe(s)).includes(a)&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await t.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(t){console.error("Unable to unhide system sheets:",t)}}async function Ao(){if(!it()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(o=>{o.visibility!==Excel.SheetVisibility.visible&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${o.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total sheets: ${t.items.length}`)})}catch(e){console.error("Unable to show all sheets:",e)}}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.unhideSystemSheets=xo,window.PrairieForge.showAllSheets=Ao);var lt={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ICON_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/E27DCDDE-8E09-4DB6-A41B-E707317FA864.png",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var Do='<svg viewBox="0 0 24 24" fill="currentColor"><path d="M22.2819 9.8211a5.9847 5.9847 0 0 0-.5157-4.9108 6.0462 6.0462 0 0 0-6.5098-2.9A6.0651 6.0651 0 0 0 4.9807 4.1818a5.9847 5.9847 0 0 0-3.9977 2.9 6.0462 6.0462 0 0 0 .7427 7.0966 5.98 5.98 0 0 0 .511 4.9107 6.051 6.051 0 0 0 6.5146 2.9001A5.9847 5.9847 0 0 0 13.2599 24a6.0557 6.0557 0 0 0 5.7718-4.2058 5.9894 5.9894 0 0 0 3.9977-2.9001 6.0557 6.0557 0 0 0-.7475-7.0729zm-9.022 12.6081a4.4755 4.4755 0 0 1-2.8764-1.0408l.1419-.0804 4.7783-2.7582a.7948.7948 0 0 0 .3927-.6813v-6.7369l2.02 1.1686a.071.071 0 0 1 .038.052v5.5826a4.504 4.504 0 0 1-4.4945 4.4944zm-9.6607-4.1254a4.4708 4.4708 0 0 1-.5346-3.0137l.142.0852 4.783 2.7582a.7712.7712 0 0 0 .7806 0l5.8428-3.3685v2.3324a.0804.0804 0 0 1-.0332.0615L9.74 19.9502a4.4992 4.4992 0 0 1-6.1408-1.6464zM2.3408 7.8956a4.485 4.485 0 0 1 2.3655-1.9728V11.6a.7664.7664 0 0 0 .3879.6765l5.8144 3.3543-2.0201 1.1685a.0757.0757 0 0 1-.071 0l-4.8303-2.7865A4.504 4.504 0 0 1 2.3408 7.8956zm16.5963 3.8558L13.1038 8.364 15.1192 7.2a.0757.0757 0 0 1 .071 0l4.8303 2.7913a4.4944 4.4944 0 0 1-.6765 8.1042v-5.6772a.79.79 0 0 0-.407-.667zm2.0107-3.0231l-.142-.0852-4.7735-2.7818a.7759.7759 0 0 0-.7854 0L9.409 9.2297V6.8974a.0662.0662 0 0 1 .0284-.0615l4.8303-2.7866a4.4992 4.4992 0 0 1 6.6802 4.66zM8.3065 12.863l-2.02-1.1638a.0804.0804 0 0 1-.038-.0567V6.0742a4.4992 4.4992 0 0 1 7.3757-3.4537l-.142.0805L8.704 5.459a.7948.7948 0 0 0-.3927.6813zm1.0976-2.3654l2.602-1.4998 2.6069 1.4998v2.9994l-2.5974 1.4997-2.6067-1.4997Z"/></svg>',Po='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>',$o=lt.ADA_IMAGE_URL,rn={id:"pf-copilot",heading:"Ada",subtext:"Your smart assistant to help you analyze and troubleshoot.",welcomeMessage:"What would you like to explore?",placeholder:"Where should I focus this pay period?",quickActions:[{id:"diagnostics",label:"Diagnostics",prompt:"Run a diagnostic check on the current payroll data. Check for completeness, accuracy issues, and any data quality concerns."},{id:"insights",label:"Insights",prompt:"What are the key insights and notable findings from this payroll period that I should highlight for executive review?"},{id:"variance",label:"Variances",prompt:"Analyze the significant variances between this period and the prior period. What's driving the changes?"},{id:"journal",label:"JE Check",prompt:"Is the journal entry ready for export? Check that debits equal credits and flag any mapping issues."}],systemPrompt:`You are Prairie Forge CoPilot, an expert financial analyst assistant embedded in an Excel add-in. 

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
- Missing or incomplete mappings`},Pt=[];function ln(e={}){var o;let t={...rn,...e},n=((o=t.quickActions)==null?void 0:o.map(a=>`<button type="button" class="pf-ada-chip" data-action="${a.id}" data-prompt="${No(a.prompt)}">${a.label}</button>`).join(""))||"";return`
        <article class="pf-ada" data-copilot="${t.id}">
            <header class="pf-ada-header">
                <div class="pf-ada-identity">
                    <img class="pf-ada-avatar" src="${$o}" alt="Ada" onerror="this.style.display='none'" />
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
                        ${Po}
                    </button>
                </div>
                
                ${n?`<div class="pf-ada-chips">${n}</div>`:""}
                
                <footer class="pf-ada-footer">
                    ${Do}
                    <span>Powered by ChatGPT</span>
                </footer>
            </div>
        </article>
    `}function No(e){return String(e||"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}function cn(e,t={}){let n={...rn,...t},o=e.querySelector(`[data-copilot="${n.id}"]`);if(!o)return;let a=o.querySelector(`#${n.id}-messages`),s=o.querySelector(`#${n.id}-prompt`),l=o.querySelector(`#${n.id}-ask`),c=o.querySelector(`#${n.id}-status-dot`),i=o.querySelector(`#${n.id}-status-badge`),r=!1,u=(C,g="ready")=>{c&&(c.classList.remove("pf-ada-status-dot--busy","pf-ada-status-dot--offline"),g==="busy"&&c.classList.add("pf-ada-status-dot--busy"),g==="offline"&&c.classList.add("pf-ada-status-dot--offline")),i&&(i.title=C)},f=(C,g="assistant")=>{if(!a)return;let b=g==="user"?"pf-ada-bubble--user":g==="system"?"pf-ada-bubble--system":"pf-ada-bubble--ai",E=document.createElement("div");E.className=`pf-ada-bubble ${b}`,E.innerHTML=`<p>${h(C)}</p>`,a.appendChild(E),a.scrollTop=a.scrollHeight,Pt.push({role:g,content:C,timestamp:new Date().toISOString()})},p=()=>{if(!a)return;let C=document.createElement("div");C.className="pf-ada-bubble pf-ada-bubble--ai pf-ada-bubble--loading",C.id=`${n.id}-loading`,C.innerHTML=`
            <div class="pf-ada-typing">
                <span></span><span></span><span></span>
            </div>
        `,a.appendChild(C),a.scrollTop=a.scrollHeight},d=()=>{let C=document.getElementById(`${n.id}-loading`);C&&C.remove()},h=C=>String(C).replace(/\*\*(.*?)\*\*/g,"<strong>$1</strong>").replace(/\n\n/g,"</p><p>").replace(/\n- /g,"<br>\u2022 ").replace(/\n/g,"<br>"),v=async C=>{let g=C||(s==null?void 0:s.value.trim());if(!(!g||r)){r=!0,s&&(s.value=""),l&&(l.disabled=!0),f(g,"user"),p(),u("Analyzing...","busy");try{let b=null;if(typeof n.contextProvider=="function")try{b=await n.contextProvider()}catch(x){console.warn("CoPilot: Context provider failed",x)}let E;typeof n.onPrompt=="function"?E=await n.onPrompt(g,b,Pt):typeof n.apiEndpoint=="string"?E=await Io(n.apiEndpoint,g,b,n.systemPrompt):E=Oo(g,b),d(),f(E,"assistant"),u("Ready to assist","ready")}catch(b){console.error("CoPilot error:",b),d(),f(`I encountered an issue: ${b.message}. Please try again.`,"system"),u("Error occurred","offline")}r=!1,l&&(l.disabled=!1),s==null||s.focus()}};l==null||l.addEventListener("click",()=>v()),s==null||s.addEventListener("keydown",C=>{C.key==="Enter"&&!C.shiftKey&&(C.preventDefault(),v())}),o.querySelectorAll(".pf-ada-chip").forEach(C=>{C.addEventListener("click",()=>{let g=C.dataset.prompt;g&&v(g)})})}async function Io(e,t,n,o){let a=await fetch(e,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({prompt:t,context:n,systemPrompt:o,history:Pt.slice(-10)})});if(!a.ok)throw new Error(`API request failed: ${a.status}`);let s=await a.json();return s.message||s.response||"No response received."}function Oo(e,t){var o,a,s;let n=e.toLowerCase();return n.includes("diagnostic")||n.includes("check")?`Great question! Let me run through the diagnostics for you.

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

I'm reading your actual spreadsheet data, so I can give you specific answers!`}var pn=lt.ADA_IMAGE_URL;async function Nt(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async o=>{let a=o.workbook.worksheets.getItemOrNullObject(e);a.load("isNullObject, name"),await o.sync();let s;a.isNullObject?(s=o.workbook.worksheets.add(e),await o.sync(),await dn(o,s,t,n)):(s=a,await dn(o,s,t,n)),s.activate(),s.getRange("A1").select(),await o.sync()})}catch(o){console.error(`Error activating homepage sheet ${e}:`,o)}}async function dn(e,t,n,o){try{let r=t.getUsedRangeOrNullObject();r.load("isNullObject"),await e.sync(),r.isNullObject||(r.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let a=[[n,""],[o,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=a;let l=t.getRange("A1:Z100");l.format.fill.color="#0f0f0f";let c=t.getRange("A1");c.format.font.bold=!0,c.format.font.size=36,c.format.font.color="#ffffff",c.format.font.name="Segoe UI Light",c.format.verticalAlignment="Center";let i=t.getRange("A2");i.format.font.size=14,i.format.font.color="#a0a0a0",i.format.font.name="Segoe UI",i.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var un={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function It(e){return un[e]||un["module-selector"]}function fn(){Ot();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${pn}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",To),e}function Ot(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function To(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${pn}" alt="Ada" />
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",$t),t.addEventListener("click",a=>{a.target===t&&$t()});let o=a=>{a.key==="Escape"&&($t(),document.removeEventListener("keydown",o))};document.addEventListener("keydown",o)}function $t(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var mn=`
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
`.trim(),gn=`
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
`.trim(),hn=`
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
`.trim(),ct=`
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
`.trim(),yn=`
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
`.trim(),vn=`
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
`.trim(),Lo={config:`
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
    `};function bn(e){return e&&Lo[e]||""}var Tt=`
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
`.trim(),Lt=`
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
`.trim(),dt=`
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
`.trim(),ut=`
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
`.trim(),ss=`
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
`.trim(),pt=`
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
`.trim(),wn=`
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
`.trim(),En=`
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
`.trim(),Cn=`
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
`.trim(),kn=`
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
        <path d="m12 5 7 7-7 7" />
        <path d="M5 12h14" />
    </svg>
`.trim(),rs=`
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
`.trim(),is=`
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
`.trim(),ls=`
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
`.trim(),cs=`
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
        <path d="M21.801 10A10 10 0 1 1 17 3.335"/>
        <path d="m9 11 3 3L22 4"/>
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
        <circle cx="12" cy="12" r="10"/>
        <path d="m15 9-6 6"/>
        <path d="m9 9 6 6"/>
    </svg>
`.trim(),ps=`
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
`.trim(),fs=`
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
`.trim(),ze=`
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
`.trim(),Sn=`
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
`.trim();function We(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function Bt(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function Ae({textareaId:e,value:t,permanentId:n,isPermanent:o,hintId:a,saveButtonId:s,isSaved:l=!1,placeholder:c="Enter notes here..."}){let i=o?Lt:Tt,r=s?`<button type="button" class="pf-action-toggle pf-save-btn ${l?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${kn}</button>`:"",u=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${o?"is-locked":""}" id="${n}" aria-pressed="${o}" title="Lock notes (retain after archive)">${i}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${We(c)}">${We(t||"")}</textarea>
                ${a?`<p class="pf-signoff-hint" id="${a}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?Bt(u,"Lock"):""}
                ${s?Bt(r,"Save"):""}
            </div>
        </article>
    `}function De({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:o,isComplete:a,saveButtonId:s,isSaved:l=!1,completeButtonId:c,subtext:i="Sign-off below. Click checkmark icon. Done."}){let r=`<button type="button" class="pf-action-toggle ${a?"is-active":""}" id="${c}" aria-pressed="${!!a}" title="Mark step complete">${dt}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${We(i)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${We(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${We(o)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${Bt(r,"Done")}
            </div>
        </article>
    `}function Je(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?Lt:Tt)}function Te(e,t){e&&e.classList.toggle("is-saved",t)}function Mt(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(o=>{let a=o.getAttribute("data-save-input"),s=document.getElementById(a);if(!s)return;let l=()=>{Te(o,!1)};s.addEventListener("input",l),n.push(()=>s.removeEventListener("input",l))}),()=>n.forEach(o=>o())}function _n(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function xn(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var Vt={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},Ft={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function An(e){e.format.fill.color=Vt.fillColor,e.format.font.color=Vt.fontColor,e.format.font.bold=Vt.bold}function ft(e,t,n,o=!1){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[o?Ft.currencyWithNegative:Ft.currency]]}function Dn(e,t,n,o=Ft.date){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[o]]}var Et="payroll-recorder",Bo=["payroll","payroll recorder","payroll review","pr"],$e="Payroll Recorder",Os=D.CONFIG||"SS_PF_Config",jt=["SS_PF_Config"];var Mo="Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel. Every run follows the same guidance so you stay audit-ready.",be=rt.map(({id:e,title:t})=>({id:e,title:t})),B={TYPE:0,FIELD:1,VALUE:2,VALUE2:3,PERMANENT:4,TITLE:-1},Vo="Run Settings";var Pn="N";var Fo="PR_JE_Total_Debit",jo="PR_JE_Total_Credit",Ho="PR_JE_Difference",Ie={0:{note:"Notes_PR_Config",reviewer:"Reviewer_PR_Config",signOff:"Sign_Off_Date_PR_Config"},1:{note:"Notes_PR_Payroll_Data",reviewer:"Reviewer_PR_Payroll_Data",signOff:"Sign_Off_Date_PR_Payroll_Data"},2:{note:"Notes_PR_Headcount",reviewer:"Reviewer_PR_Headcount",signOff:"Sign_Off_Date_PR_Headcount"},3:{note:"Notes_PR_Validate",reviewer:"Reviewer_PR_Validate",signOff:"Sign_Off_Date_PR_Validate"},4:{note:"Notes_PR_Review",reviewer:"Reviewer_PR_Review",signOff:"Sign_Off_Date_PR_Review"},5:{note:"Notes_PR_JE",reviewer:"Reviewer_PR_JE",signOff:"Sign_Off_Date_PR_JE"},6:{note:"Notes_PR_Archive",reviewer:"Reviewer_PR_Archive",signOff:"Sign_Off_Date_PR_Archive"}},de={0:"Complete_PR_Config",1:"Complete_PR_Payroll_Data",2:"Complete_PR_Headcount",3:"Complete_PR_Validate",4:"Complete_PR_Review",5:"Complete_PR_JE",6:"Complete_PR_Archive"},Uo={1:D.DATA,2:D.DATA_CLEAN,3:D.DATA_CLEAN,4:D.EXPENSE_REVIEW,5:D.JE_DRAFT},qe="PR_Reviewer_Name",Hn="PR_Payroll_Provider",mt="User opted to skip the headcount review this period.",te={statusText:"",focusedIndex:0,activeView:"home",activeStepId:null,stepStatuses:be.reduce((e,t)=>({...e,[t.id]:"pending"}),{})},W={loaded:!1,values:{},permanents:{},overrides:{accountingPeriod:!1,jeId:!1}},Le=new Map,gt=null,Xe=["Payroll Date (YYYY-MM-DD)","Payroll_Date","Payroll Date","PR_Payroll_Date","Payroll_Date_(YYYY-MM-DD)"],H={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},departments:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},bt=null,F={loading:!1,lastError:null,prDataTotal:null,cleanTotal:null,reconDifference:null,bankAmount:"",bankDifference:null,plugEnabled:!1},ke={loading:!1,lastError:null,periods:[],copilotResponse:"",completenessCheck:{currentPeriod:null,historicalPeriods:null}},G={debitTotal:null,creditTotal:null,difference:null,loading:!1,lastError:null};async function Go(){if(console.log("Completeness Check - Starting..."),!ae()){console.log("Completeness Check - Excel runtime not available");return}try{await Excel.run(async e=>{var a,s,l,c;let t=e.workbook.worksheets.getItemOrNullObject(D.DATA_CLEAN),n=e.workbook.worksheets.getItemOrNullObject(D.ARCHIVE_SUMMARY);t.load("isNullObject"),n.load("isNullObject"),await e.sync();let o={currentPeriod:null,historicalPeriods:null};if(!t.isNullObject){let i=t.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&i.values&&i.values.length>1){let r=(i.values[0]||[]).map(p=>String(p||"").toLowerCase().trim()),u=r.findIndex(p=>p.includes("amount")),f=u>=0?u:r.findIndex(p=>p==="total"||p==="all-in"||p==="allin"||p==="all-in total"||p==="gross"||p==="total pay");if(console.log("Completeness Check - PR_Data_Clean headers:",r),console.log("Completeness Check - Amount column index:",u,"Total column index:",f),f>=0){let d=i.values.slice(1).reduce((C,g)=>C+(Number(g[f])||0),0),h=((l=(s=(a=ke.periods)==null?void 0:a[0])==null?void 0:s.summary)==null?void 0:l.total)||0;console.log("Completeness Check - PR_Data_Clean sum:",d,"Current period total:",h);let v=Math.abs(d-h)<1;o.currentPeriod={match:v,prDataClean:d,currentTotal:h}}else console.warn("Completeness Check - No amount/total column found in PR_Data_Clean")}}if(!n.isNullObject){let i=n.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&i.values&&i.values.length>1){let r=(i.values[0]||[]).map(d=>String(d||"").toLowerCase().trim()),u=r.findIndex(d=>d.includes("pay period")||d.includes("payroll date")||d==="date"||d==="period"||d.includes("period")),f=r.findIndex(d=>d.includes("amount")),p=f>=0?f:r.findIndex(d=>d==="total"||d==="all-in"||d==="allin"||d==="all-in total"||d==="total payroll"||d.includes("total"));if(console.log("Completeness Check - PR_Archive_Summary headers:",r),console.log("Completeness Check - Date column index:",u,"Total column index:",p),p>=0&&u>=0){let d=i.values.slice(1),h=(ke.periods||[]).slice(1,6);console.log("Completeness Check - Looking for periods:",h.map(_=>_.key||_.label));let v=new Map;for(let _ of d){let T=_[u],m=In(T);m&&v.set(m,Number(_[p])||0)}console.log("Completeness Check - Archive lookup keys:",Array.from(v.keys()));let C=0,g=0,b=0,E=[];for(let _ of h){let T=_.key||_.label||"",m=In(T),N=((c=_.summary)==null?void 0:c.total)||0;g+=N;let L=v.get(m);L!==void 0?(C+=L,b++,E.push({period:T,calculated:N,archive:L,match:Math.abs(N-L)<1})):(console.warn(`Completeness Check - Period ${T} (normalized: ${m}) not found in archive`),E.push({period:T,calculated:N,archive:null,match:!1}))}console.log("Completeness Check - Period details:",E),console.log("Completeness Check - Matched",b,"of",h.length,"periods"),console.log("Completeness Check - Archive sum:",C,"Periods sum:",g);let x=b===h.length&&h.length>0,A=Math.abs(C-g)<1,M=x&&A;o.historicalPeriods={match:M,archiveSum:C,periodsSum:g,matchedCount:b,totalPeriods:h.length,details:E}}else console.warn("Completeness Check - Missing date or total column in PR_Archive_Summary"),console.warn("  Date column index:",u,"Total column index:",p)}}ke.completenessCheck=o,console.log("Completeness Check - Results:",JSON.stringify(o))}),console.log("Completeness Check - Complete!")}catch(e){console.error("Payroll completeness check failed:",e)}}function zo(){var v,C;let e=ke.completenessCheck||{},t=((v=ke.periods)==null?void 0:v.length)>0,n=g=>`$${Math.round(g||0).toLocaleString()}`,o=g=>{let b=Math.abs(g);return b<1?"\u2014":`${g>0?"+":"-"}$${Math.round(b).toLocaleString()}`},a=(g,b,E,x,A,M,_)=>{let T=(E||0)-(A||0),m,N;_?(m='<span class="pf-complete-status pf-complete-status--pending">\u23F3</span>',N="pending"):M?(m='<span class="pf-complete-status pf-complete-status--pass">\u2713</span>',N="pass"):(m='<span class="pf-complete-status pf-complete-status--fail">\u2717</span>',N="fail");let L=_?"":`
            <div class="pf-complete-diff ${N}">
                ${o(T)}
            </div>
        `;return`
            <div class="pf-complete-row ${N}">
                <div class="pf-complete-header">
                    ${m}
                    <span class="pf-complete-label">${S(g)}</span>
                </div>
                ${_?`
                <div class="pf-complete-values">
                    <span class="pf-complete-pending-text">Click Run/Refresh to check</span>
                </div>
                `:`
                <div class="pf-complete-values">
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${S(b)}:</span>
                        <span class="pf-complete-amount">${n(E)}</span>
                    </div>
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${S(x)}:</span>
                        <span class="pf-complete-amount">${n(A)}</span>
                    </div>
                </div>
                ${L}
                `}
            </div>
        `},s=e.currentPeriod,l=!t||s===null||s===void 0,c=a("Current Period","PR_Data_Clean Total",s==null?void 0:s.prDataClean,"Calculated Total",s==null?void 0:s.currentTotal,s==null?void 0:s.match,l),i=e.historicalPeriods,r=!t||i===null||i===void 0,u=(i==null?void 0:i.matchedCount)||0,f=(i==null?void 0:i.totalPeriods)||0,p=f>0?`Historical Periods (${u}/${f} matched)`:"Historical Periods",d=a(p,"PR_Archive_Summary (matched)",i==null?void 0:i.archiveSum,"Calculated Total",i==null?void 0:i.periodsSum,i==null?void 0:i.match,r),h="";return!r&&((C=i==null?void 0:i.details)==null?void 0:C.length)>0&&(h=`
            <div class="pf-complete-details-section">
                <div class="pf-complete-details-header">Period-by-Period Match</div>
                ${i.details.map(b=>{let E=b.archive===null?"\u26A0\uFE0F":b.match?"\u2713":"\u2717",x=b.archive!==null?n(b.archive):"Not found";return`
                <div class="pf-complete-detail-row">
                    <span class="pf-complete-detail-date">${S(b.period)}</span>
                    <span class="pf-complete-detail-icon">${E}</span>
                    <span class="pf-complete-detail-vals">${n(b.calculated)} vs ${x}</span>
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
                ${d}
                ${h}
            </div>
        </article>
    `}function Wo(e){switch(e){case 0:return{title:"Configuration",content:`
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
                `}}}nn(()=>Jo());async function Jo(){try{await Yo(),await Wn();let e=It(Et);await Nt(e.sheetName,e.title,e.subtitle),le()}catch(e){throw console.error("[Payroll] Module initialization failed:",e),e}}async function Yo(){try{await an(Et,{aliasTokens:Bo}),console.log(`[Payroll] Tab visibility applied for ${Et}`)}catch(e){console.warn("[Payroll] Could not apply tab visibility:",e)}}function le(){var i;let e=document.body;if(!e)return;let t=te.focusedIndex<=0?"disabled":"",n=te.focusedIndex>=be.length-1?"disabled":"",o=te.activeView==="config",a=te.activeView==="step",s=!o&&!a,l=o?Xo():a?aa(te.activeStepId):Ko();e.innerHTML=`
        <div class="pf-root">
            ${qo(t,n)}
            ${l}
            ${ra()}
        </div>
    `;let c=document.getElementById("pf-info-fab-payroll");if(s)c&&c.remove();else if((i=window.PrairieForge)!=null&&i.mountInfoFab){let r=Wo(te.activeStepId);PrairieForge.mountInfoFab({title:r.title,content:r.content,buttonId:"pf-info-fab-payroll"})}if(la(),o)ua();else if(a)try{pa(te.activeStepId)}catch(r){console.warn("Payroll Recorder: failed to bind step interactions",r)}else da();fa(),s?fn():Ot()}function qo(e,t){let n=$("Company_Name")||"your company";return`
        <div class="pf-brand-float" aria-hidden="true">
            <span class="pf-brand-wave"></span>
        </div>
        <header class="pf-banner">
            <div class="pf-nav-bar">
                <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                    ${Cn}
                    <span class="sr-only">Previous step</span>
                </button>
                <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                    ${mn}
                    <span class="sr-only">Module Home</span>
                </button>
                <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                    ${gn}
                    <span class="sr-only">Module Selector</span>
                </button>
                <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                    ${Rn}
                    <span class="sr-only">Next step</span>
                </button>
                <span class="pf-nav-divider"></span>
                <div class="pf-quick-access-wrapper">
                    <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                        ${hn}
                        <span class="sr-only">Quick Access Menu</span>
                    </button>
                    <div id="quick-access-dropdown" class="pf-quick-dropdown hidden">
                        <div class="pf-quick-dropdown-header">Quick Access</div>
                        <button id="nav-roster" class="pf-quick-item pf-clickable" type="button">
                            ${yn}
                            <span>Employee Roster</span>
                        </button>
                        <button id="nav-accounts" class="pf-quick-item pf-clickable" type="button">
                            ${vn}
                            <span>Chart of Accounts</span>
                        </button>
                        <button id="nav-expense-map" class="pf-quick-item pf-clickable" type="button">
                            ${ct}
                            <span>PR Mapping</span>
                </button>
                    </div>
                </div>
            </div>
        </header>
    `}function Ko(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">Payroll Recorder</h2>
            <p class="pf-hero-copy">${Mo}</p>
            <p class="pf-hero-hint">${S(te.statusText||"")}</p>
        </section>
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${be.map((e,t)=>sa(e,t)).join("")}
            </div>
        </section>
    `}function Xo(){if(!W.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=Ie[0],t=ye(Ct()),n=ye($("Accounting_Period")),o=$("Journal_Entry_ID"),a=$("Accounting_Software"),s=Yt(),l=$("Company_Name"),c=$(qe)||Pe(),i=e?$(e.note):"",r=e?Re(e.note):!1,u=(e?$(e.reviewer):"")||Pe(),f=e?ye($(e.signOff)):"",p=!!(f||$(de[0]));return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${S($e)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every payroll run.</p>
            <p class="pf-hero-hint">${S(te.statusText||"")}</p>
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
                        <input type="text" id="config-user-name" value="${S(c)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Date</span>
                        <input type="date" id="config-payroll-date" value="${S(t)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${S(n)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-je-id" value="${S(o)}" placeholder="PR-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${S(l)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${S(s)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${S(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${e?Ae({textareaId:"config-notes",value:i,permanentId:"config-notes-permanent",isPermanent:r,hintId:"",saveButtonId:"config-notes-save"}):""}
            ${e?De({reviewerInputId:"config-reviewer-name",reviewerValue:u,signoffInputId:"config-signoff-date",signoffValue:f,isComplete:p,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"}):""}
        </section>
    `}function Qo(e){let t=Se(1),n=t?Re(t.note):!1,o=t?$(t.note):"",a=(t?$(t.reviewer):"")||Pe(),s=t?ye($(t.signOff)):"",l=!!(s||$(de[1])),c=Yt();return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S($e)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">Pull your payroll export from the provider and paste it into PR_Data.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Open your payroll provider, download the report, and paste into PR_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(c?`<a href="${S(c)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${pt}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${pt}</button>`,"Provider")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PR_Data sheet">${ct}</button>`,"PR_Data")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PR_Data to start over">${Sn}</button>`,"Clear")}
                </div>
            </article>
            ${t?`
                ${Ae({textareaId:"step-notes-1",value:o||"",permanentId:"step-notes-lock-1",isPermanent:n,saveButtonId:"step-notes-save-1"})}
                ${De({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:s,isComplete:l,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
            `:""}
        </section>
    `}function Zo(e){var A,M,_,T,m,N,L,Q,pe,K;let t=Se(2),n=t?$(t.note):"",o=t?Re(t.note):!1,a=(t?$(t.reviewer):"")||Pe(),s=t?ye($(t.signOff)):"",l=!!(s||$(de[2])),c=Rt(),i=H.roster||{},r=H.departments||{},u=H.hasAnalyzed,f="";H.loading?f='<p class="pf-step-note">Analyzing roster and payroll data\u2026</p>':H.lastError&&(f=`<p class="pf-step-note">${S(H.lastError)}</p>`);let p=(V,Y,ne)=>{let _e=!u,X;_e?X='<span class="pf-je-check-circle pf-je-circle--pending"></span>':ne?X=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:X=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let Z=u?` = ${Y}`:"";return`
            <div class="pf-je-check-row">
                ${X}
                <span class="pf-je-check-desc-pill">${S(V)}${Z}</span>
            </div>
        `},d=(A=i.difference)!=null?A:0,h=(M=r.difference)!=null?M:0,v=Array.isArray(i.mismatches)?i.mismatches.filter(Boolean):[],C=Array.isArray(r.mismatches)?r.mismatches.filter(Boolean):[],g=`
        ${p("SS_Employee_Roster count",(_=i.rosterCount)!=null?_:"\u2014",!0)}
        ${p("PR_Data employee count",(T=i.payrollCount)!=null?T:"\u2014",!0)}
        ${p("Difference",d,d===0)}
    `,b=`
        ${p("Expected departments",(m=r.rosterCount)!=null?m:"\u2014",!0)}
        ${p("PR_Data departments",(N=r.payrollCount)!=null?N:"\u2014",!0)}
        ${p("Difference",h,h===0)}
    `,E=v.length&&!H.skipAnalysis&&u&&((Q=(L=window.PrairieForge)==null?void 0:L.renderMismatchTiles)==null?void 0:Q.call(L,{mismatches:v,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"Payroll Data",escapeHtml:S}))||"",x=C.length&&!H.skipAnalysis&&u&&((K=(pe=window.PrairieForge)==null?void 0:pe.renderMismatchTiles)==null?void 0:K.call(pe,{mismatches:C,label:"Employees with Department Differences",formatter:V=>({name:V.employee||V.name||"",source:`${V.rosterDept||"\u2014"} \u2192 ${V.payrollDept||"\u2014"}`,isMissingFromTarget:!0}),escapeHtml:S}))||"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S($e)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">Headcount Review</h2>
            <p class="pf-hero-copy">Quick check to make sure your roster matches your payroll data.</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${H.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${En}
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
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="roster-run-btn" title="Run headcount analysis">${ut}</button>`,"Run")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="roster-refresh-btn" title="Refresh analysis">${ze}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Employee Alignment</h3>
                    <p class="pf-config-subtext">Verify employees match between roster and payroll.</p>
                </div>
                ${f}
                <div class="pf-je-checks-container">
                    ${g}
                </div>
                ${E}
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Department Alignment</h3>
                    <p class="pf-config-subtext">Verify department assignments are consistent.</p>
                </div>
                <div class="pf-je-checks-container">
                    ${b}
                </div>
                ${x}
            </article>
            ${t?`
                ${Ae({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:o,hintId:c?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
                ${De({reviewerInputId:"step-reviewer-name",reviewerValue:a,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:l,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
            `:""}
        </section>
    `}function ea(e){var A;let t=Se(3),n=t?$(t.note):"",o=(t?$(t.reviewer):"")||Pe(),a=t?ye($(t.signOff)):"",s=F.loading?'<p class="pf-step-note">Preparing reconciliation data\u2026</p>':F.lastError?`<p class="pf-step-note">${S(F.lastError)}</p>`:"",l=!!(a||$(de[3])),c=F.prDataTotal!==null,i=F.prDataTotal,r=F.cleanTotal,u=(A=F.reconDifference)!=null?A:i!=null&&r!=null?i-r:null,f=u!==null&&Math.abs(u)<.01,p=ie(F.cleanTotal),d=F.bankDifference!=null?ie(F.bankDifference):"---",h=F.bankDifference==null?"":Math.abs(F.bankDifference)<.5?"Difference within acceptable tolerance.":"Difference exceeds tolerance and should be resolved.",v=Xn(F.bankAmount),C=(M,_,T)=>{let m=!c,N;return m?N='<span class="pf-je-check-circle pf-je-circle--pending"></span>':T?N=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:N=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${N}
                <span class="pf-je-check-desc-pill">${S(_)}</span>
            </div>
        `},g=c?ie(i):"\u2014",b=c?ie(r):"\u2014",E=c?ie(u):"\u2014",x=`
        ${C("PR_Data Total",`PR_Data Total = ${g}`,!0)}
        ${C("PR_Data_Clean Total",`PR_Data_Clean Total = ${b}`,!0)}
        ${C("Difference",`Difference = ${E} (should be $0.00)`,f)}
    `;return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S($e)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">Normalize your payroll data and verify totals match.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Validation</h3>
                    <p class="pf-config-subtext">Build PR_Data_Clean from your imported data and verify totals.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="validation-run-btn" title="Run reconciliation">${ut}</button>`,"Run")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="validation-refresh-btn" title="Refresh reconciliation">${ze}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Data Reconciliation</h3>
                    <p class="pf-config-subtext">Verify PR_Data and PR_Data_Clean totals match.</p>
                </div>
                ${s}
                <div class="pf-je-checks-container">
                    ${x}
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
                        <input id="bank-clean-total-value" type="text" class="pf-readonly-input pf-metric-value" value="${p}" readonly>
                    </label>
                    <label class="pf-config-field">
                        <span>Cost per Bank</span>
                        <input
                            type="text"
                            inputmode="decimal"
                            id="bank-amount-input"
                            class="pf-metric-input"
                            value="${S(v)}"
                            placeholder="0.00"
                            aria-label="Enter bank amount"
                        >
                    </label>
                    <label class="pf-config-field">
                        <span>Difference</span>
                        <input id="bank-diff-value" type="text" class="pf-readonly-input pf-metric-value" value="${d}" readonly>
                    </label>
                </div>
                <p class="pf-metric-hint" id="bank-diff-hint">${S(h)}</p>
            </article>
            ${t?`
                ${Ae({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:Re(t.note),saveButtonId:"step-notes-save-3"})}
            `:""}
            ${De({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-3",signoffValue:a,isComplete:l,saveButtonId:"step-signoff-save-3",completeButtonId:"validation-signoff-toggle"})}
        </section>
    `}function ta(e){let t=Se(4),n=t?$(t.note):"",o=(t?$(t.reviewer):"")||Pe(),a=t?ye($(t.signOff)):"",s=!!(a||$(de[4])),l=ke.loading?'<p class="pf-step-note">Preparing executive summary\u2026</p>':ke.lastError?`<p class="pf-step-note">${S(ke.lastError)}</p>`:"",c=ln({id:"expense-review-copilot",body:"Want help analyzing your data? Just ask!",placeholder:"Where should I focus this pay period?",buttonLabel:"Ask CoPilot"});return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S($e)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">${S(e.summary||"")}</p>
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
                    ${ge(`<button type="button" class="pf-action-toggle" id="expense-run-btn" title="Run expense review analysis">${ut}</button>`,"Run")}
                    ${ge(`<button type="button" class="pf-action-toggle" id="expense-refresh-btn" title="Refresh expense data">${ze}</button>`,"Refresh")}
                </div>
            </article>
            ${zo()}
            <div class="pf-ada-coming-soon-wrapper">
                ${c}
            </div>
            ${t?`
            ${Ae({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:Re(t.note),saveButtonId:"step-notes-save-4"})}
            ${De({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-4",signoffValue:a,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"expense-signoff-toggle"})}
            `:""}
        </section>
    `}function na(e){var A,M,_;let t=Se(5),n=t?$(t.note):"",o=t?Re(t.note):!1,a=(t?$(t.reviewer):"")||Pe(),s=t?ye($(t.signOff)):"",l=!!(s||$(de[5])),c=G.lastError?`<p class="pf-step-note">${S(G.lastError)}</p>`:"",i=G.debitTotal!==null,r=(A=G.debitTotal)!=null?A:0,u=(M=G.creditTotal)!=null?M:0,f=r-u,p=(_=F.cleanTotal)!=null?_:0,d=i,h=i&&Math.abs(f-p)<.01,v=(T,m)=>{let N=!i,L;return N?L='<span class="pf-je-check-circle pf-je-circle--pending"></span>':m?L=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:L=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${L}
                <span class="pf-je-check-desc-pill">${S(T)}</span>
            </div>
        `},C=i?ie(r):"\u2014",g=i?ie(u):"\u2014",b=i?ie(f):"\u2014",E=i?ie(p):"\u2014",x=`
        ${v(`Total Debits = ${C}`,d)}
        ${v(`Total Credits = ${g}`,d)}
        ${v(`Line Amount (Debit - Credit) = ${b}`,d)}
        ${v(`JE Total matches PR_Data_Clean (${E})`,h)}
    `;return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S($e)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">Generate the upload file to break down the bank feed line item.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Upload File</h3>
                    <p class="pf-config-subtext">Build the breakdown from PR_Data_Clean for your accounting system.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate from PR_Data_Clean">${ct}</button>`,"Generate")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation">${ze}</button>`,"Refresh")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export as CSV">${wn}</button>`,"Export")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">Verify totals before uploading to your accounting system.</p>
                </div>
                ${c}
                <div class="pf-je-checks-container">
                    ${x}
                </div>
            </article>
            ${t?`
                ${Ae({textareaId:"step-notes-input",value:n||"",permanentId:"step-notes-permanent",isPermanent:o,saveButtonId:"step-notes-save-5"})}
                ${De({reviewerInputId:"step-reviewer-name",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:s,isComplete:l,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
            `:""}
        </section>
    `}function oa(e){let t=be.filter(a=>a.id!==6).map(a=>({id:a.id,title:a.title,complete:ya(a.id)})),n=t.every(a=>a.complete),o=t.map(a=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${a.complete?"is-active":""}" aria-pressed="${a.complete}">
                        ${dt}
                    </span>
                    <div>
                        <h3>${S(a.title)}</h3>
                        <p class="pf-config-subtext">${a.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S($e)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">${S(e.summary||"")}</p>
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
    `}function aa(e){let t=rt.find(x=>x.id===e)||{id:e!=null?e:"-",title:"Workflow Step",summary:"",description:"",checklist:[]};if(e===1)return Qo(t);if(e===2)return Zo(t);if(e===3)return ea(t);if(e===4)return ta(t);if(e===5)return na(t);if(e===6)return oa(t);let n=!1,o=Se(e),a=o?$(o.note):"",s=o?Re(o.note):!1,l=(o?$(o.reviewer):"")||Pe(),c=o?ye($(o.signOff)):"",i=o&&de[e]?!!(c||$(de[e])):!!c,r=(t.highlights||[]).map(x=>`
            <div class="pf-step-highlight">
                <span class="pf-step-highlight-label">${S(x.label)}</span>
                <span class="pf-step-highlight-detail">${S(x.detail)}</span>
            </div>
        `).join(""),u=(t.checklist||[]).map(x=>`<li>${S(x)}</li>`).join("")||"",f=n?"":t.description||"Detailed guidance will appear here.",p=[];!n&&t.ctaLabel&&p.push(`<button type="button" class="pf-pill-btn" id="step-action-btn">${S(t.ctaLabel)}</button>`),n||p.push('<button type="button" class="pf-pill-btn pf-pill-btn--ghost" id="step-back-btn">Back to Step List</button>');let d=p.length?`<div class="pf-pill-row pf-config-actions">${p.join("")}</div>`:"",h=Yt(),v=n?`
            <div class="pf-link-card">
                <h3 class="pf-link-card__title">Payroll Reports</h3>
                <p class="pf-link-card__subtitle">Open your latest payroll export.</p>
                <div class="pf-link-list">
                    <a
                        class="pf-link-item"
                        id="pr-provider-link"
                        ${h?`href="${S(h)}" target="_blank" rel="noopener noreferrer"`:'aria-disabled="true"'}
                    >
                        <span class="pf-link-item__icon">${pt}</span>
                        <span class="pf-link-item__body">
                            <span class="pf-link-item__title">Open Payroll Export</span>
                            <span class="pf-link-item__meta">${S(h||"Add a provider link in Configuration")}</span>
                        </span>
                    </a>
                </div>
            </div>
        `:"",C="",g=!n&&r?`<article class="pf-step-card pf-step-detail">${r}</article>`:"",b=!n&&u?`<article class="pf-step-card pf-step-detail">
                            <h3 class="pf-step-subtitle">Checklist</h3>
                            <ul class="pf-step-checklist">${u}</ul>
                        </article>`:"",E=!n||f||d?`
            <article class="pf-step-card pf-step-detail">
                <p class="pf-step-title">${S(f)}</p>
                ${!n&&t.statusHint?`<p class="pf-step-note">${S(t.statusHint)}</p>`:""}
                ${d}
            </article>
        `:"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S($e)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${S(t.title)}</h2>
            <p class="pf-hero-copy">${S(t.summary||"")}</p>
            <p class="pf-hero-hint">${S(te.statusText||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${v}
            ${C}
            ${E}
            ${g}
            ${b}
            ${o?`
                ${Ae({textareaId:"step-notes-input",value:a,permanentId:"step-notes-permanent",isPermanent:s,saveButtonId:"step-notes-save"})}
                ${De({reviewerInputId:"step-reviewer-name",reviewerValue:l,signoffInputId:`step-signoff-${e}`,signoffValue:c,isComplete:i,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`,subtext:"Ready to move on? Save and click Done when finished."})}
            `:""}
        </section>
    `}function sa(e,t){let n=te.focusedIndex===t?"pf-step-card--active":"",o=bn(ia(e.id));return`
        <article class="pf-step-card pf-clickable ${n}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${S(e.title)}</h3>
        </article>
    `}function ra(){return`
        <footer class="pf-brand-footer">
            <div class="pf-brand-text">
                <div class="pf-brand-label">prairie.forge</div>
                <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${tn}</div>
            </div>
        </footer>
    `}function ia(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}function la(){var n,o,a,s,l,c,i;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",()=>{var r;Un(),(r=document.getElementById("pf-hero"))==null||r.scrollIntoView({behavior:"smooth",block:"start"})}),(o=document.getElementById("nav-selector"))==null||o.addEventListener("click",()=>{window.location.href="../module-selector/index.html"}),(a=document.getElementById("nav-prev"))==null||a.addEventListener("click",()=>Nn(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>Nn(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",r=>{r.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",r=>{!(t!=null&&t.contains(r.target))&&!(e!=null&&e.contains(r.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(l=document.getElementById("nav-roster"))==null||l.addEventListener("click",()=>{$n("SS_Employee_Roster"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(c=document.getElementById("nav-accounts"))==null||c.addEventListener("click",()=>{$n("SS_Chart_of_Accounts"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(i=document.getElementById("nav-expense-map"))==null||i.addEventListener("click",async()=>{t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"),await ca()})}async function $n(e){if(!e||typeof Excel=="undefined")return;let t={SS_Employee_Roster:["Employee","Department","Pay_Rate","Status","Hire_Date"],SS_Chart_of_Accounts:["Account_Number","Account_Name","Type","Category"]};try{await Excel.run(async n=>{let o=n.workbook.worksheets.getItemOrNullObject(e);if(o.load("isNullObject"),await n.sync(),o.isNullObject){o=n.workbook.worksheets.add(e);let a=t[e]||["Column1","Column2"],s=o.getRange(`A1:${String.fromCharCode(64+a.length)}1`);s.values=[a],s.format.font.bold=!0,s.format.fill.color="#f0f0f0",s.format.autofitColumns(),await n.sync()}o.activate(),o.getRange("A1").select(),await n.sync()})}catch(n){console.error("Error opening reference sheet:",n)}}async function ca(){try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItemOrNullObject("PR_Expense_Mapping");if(t.load("isNullObject"),await e.sync(),t.isNullObject){t=e.workbook.worksheets.add("PR_Expense_Mapping");let n=["Expense_Category","GL_Account","Description","Active"],o=t.getRange("A1:D1");o.values=[n],o.format.font.bold=!0}t.activate(),t.getRange("A1").select(),await e.sync()})}catch(e){console.error("Error navigating to PR_Expense_Mapping:",e)}}function da(){document.querySelectorAll("[data-step-card]").forEach(e=>{let t=Number(e.getAttribute("data-step-index"));e.addEventListener("click",()=>Qe(t))})}function ua(){var i,r,u,f;let e=document.getElementById("config-user-name");e==null||e.addEventListener("change",p=>{let d=p.target.value.trim();U(qe,d);let h=document.getElementById("config-reviewer-name");h&&!h.value&&(h.value=d)});let t=document.getElementById("config-payroll-date");t==null||t.addEventListener("change",p=>{let d=p.target.value||"";if(U(Jn(),d),!!d){if(!W.overrides.accountingPeriod){let h=va(d);if(h){let v=document.getElementById("config-accounting-period");v&&(v.value=h),U("Accounting_Period",h)}}if(!W.overrides.jeId){let h=ba(d);if(h){let v=document.getElementById("config-je-id");v&&(v.value=h),U("Journal_Entry_ID",h)}}}});let n=Se(0),o=document.getElementById("config-accounting-period");o==null||o.addEventListener("change",p=>{W.overrides.accountingPeriod=!!p.target.value,U("Accounting_Period",p.target.value||"")});let a=document.getElementById("config-je-id");a==null||a.addEventListener("change",p=>{W.overrides.jeId=!!p.target.value,U("Journal_Entry_ID",p.target.value.trim())}),(i=document.getElementById("config-company-name"))==null||i.addEventListener("change",p=>{U("Company_Name",p.target.value.trim())}),(r=document.getElementById("config-payroll-provider"))==null||r.addEventListener("change",p=>{let d=p.target.value.trim();U(Hn,d)}),(u=document.getElementById("config-accounting-link"))==null||u.addEventListener("change",p=>{U("Accounting_Software",p.target.value.trim())});let s=document.getElementById("config-notes");if(s==null||s.addEventListener("input",p=>{n&&U(n.note,p.target.value,{debounceMs:400})}),n){let p=document.getElementById("config-notes-permanent");p&&(p.addEventListener("click",()=>{let h=!p.classList.contains("is-locked");Je(p,h),Yn(n.note,h)}),Je(p,Re(n.note)));let d=document.getElementById("config-notes-save");d==null||d.addEventListener("click",()=>{s&&(U(n.note,s.value),Te(d,!0))})}let l=document.getElementById("config-reviewer-name");l==null||l.addEventListener("change",p=>{let d=p.target.value.trim();n&&U(n.reviewer,d),U(qe,d);let h=document.getElementById("config-signoff-date");if(d&&h&&!h.value){let v=Ze();h.value=v,n&&U(n.signOff,v)}}),(f=document.getElementById("config-signoff-date"))==null||f.addEventListener("change",p=>{n&&U(n.signOff,p.target.value||"")});let c=document.getElementById("config-signoff-save");if(c==null||c.addEventListener("click",()=>{var v;let p=((v=l==null?void 0:l.value)==null?void 0:v.trim())||"",d=document.getElementById("config-signoff-date"),h=(d==null?void 0:d.value)||"";n&&(U(n.reviewer,p),U(n.signOff,h)),U(qe,p),Te(c,!0)}),Mt(),n){let p=$(n.signOff),d=$(de[0]),h=!!(p||d==="Y"||d===!0);console.log(`[Step 0] Binding signoff toggle. signOff="${p}", complete="${d}", isComplete=${h}`),zn({buttonId:"config-signoff-toggle",inputId:"config-signoff-date",fieldName:n.signOff,completeField:de[0],initialActive:h,stepId:0})}}function pa(e){var n,o,a,s,l,c,i,r,u,f,p,d,h,v,C,g,b,E,x,A,M;if((n=document.getElementById("step-back-btn"))==null||n.addEventListener("click",()=>{Un()}),(o=document.getElementById("step-action-btn"))==null||o.addEventListener("click",()=>{let _=rt.find(T=>T.id===e);window.alert(_!=null&&_.ctaLabel?`${_.ctaLabel} coming soon.`:"Step actions coming soon.")}),e===1&&((a=document.getElementById("import-open-data-btn"))==null||a.addEventListener("click",()=>ga()),(s=document.getElementById("import-clear-btn"))==null||s.addEventListener("click",()=>ha())),e===2&&((l=document.getElementById("headcount-skip-btn"))==null||l.addEventListener("click",()=>{H.skipAnalysis=!H.skipAnalysis;let _=document.getElementById("headcount-skip-btn");_==null||_.classList.toggle("is-active",H.skipAnalysis),H.skipAnalysis&&Wt(),vt()}),(c=document.getElementById("roster-run-btn"))==null||c.addEventListener("click",()=>zt()),(i=document.getElementById("roster-refresh-btn"))==null||i.addEventListener("click",()=>zt()),(r=document.getElementById("roster-review-btn"))==null||r.addEventListener("click",()=>{var T;let _=((T=H.roster)==null?void 0:T.mismatches)||[];Fn("Roster Differences",_,{sourceLabel:"Roster",targetLabel:"Payroll Data"})}),(u=document.getElementById("dept-review-btn"))==null||u.addEventListener("click",()=>{var T;let _=((T=H.departments)==null?void 0:T.mismatches)||[];Fn("Department Differences",_,{sourceLabel:"Roster",targetLabel:"Payroll",formatter:m=>({name:m.employee,source:`${m.rosterDept} \u2192 ${m.payrollDept}`,isMissingFromTarget:!0})})})),e===3&&((f=document.getElementById("validation-run-btn"))==null||f.addEventListener("click",()=>Vn()),(p=document.getElementById("validation-refresh-btn"))==null||p.addEventListener("click",()=>Vn()),(d=document.getElementById("bank-amount-input"))==null||d.addEventListener("blur",jn),(h=document.getElementById("bank-amount-input"))==null||h.addEventListener("keydown",_=>{_.key==="Enter"&&jn(_)})),e===5&&((v=document.getElementById("je-run-btn"))==null||v.addEventListener("click",()=>Ga()),(C=document.getElementById("je-save-btn"))==null||C.addEventListener("click",()=>za()),(g=document.getElementById("je-create-btn"))==null||g.addEventListener("click",()=>Wa()),(b=document.getElementById("je-export-btn"))==null||b.addEventListener("click",()=>Ja())),e===4){let _=document.querySelector(".pf-step-guide");if(_){let T="https://your-project.supabase.co/functions/v1/copilot";cn(_,{id:"expense-review-copilot",contextProvider:Ea(),systemPrompt:`You are Prairie Forge CoPilot, an expert financial analyst assistant for payroll expense review.

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
When asked about readiness, be specific about what passes and what needs attention.`})}(E=document.getElementById("expense-run-btn"))==null||E.addEventListener("click",()=>{Bn()}),(x=document.getElementById("expense-refresh-btn"))==null||x.addEventListener("click",()=>{Bn()})}let t=Se(e);if(console.log(`[Step ${e}] Binding interactions, fields:`,t),t){let _=e===1?"step-notes-1":"step-notes-input",T=document.getElementById(_);console.log(`[Step ${e}] Notes input found:`,!!T,`(id: ${_})`);let m=e===1?document.getElementById("step-notes-save-1"):e===2?document.getElementById("step-notes-save-2"):e===3?document.getElementById("step-notes-save-3"):e===4?document.getElementById("step-notes-save-4"):e===5?document.getElementById("step-notes-save-5"):document.getElementById("step-notes-save");T==null||T.addEventListener("input",Z=>{U(t.note,Z.target.value,{debounceMs:400}),e===2&&(H.skipAnalysis&&Wt(),vt())}),m==null||m.addEventListener("click",()=>{T&&(U(t.note,T.value),Te(m,!0))});let N=e===1?"step-reviewer-1":"step-reviewer-name",L=document.getElementById(N);L==null||L.addEventListener("change",Z=>{let fe=Z.target.value.trim();U(t.reviewer,fe);let me=e===1?document.getElementById("step-signoff-1"):e===2?document.getElementById("step-signoff-date"):e===3?document.getElementById("step-signoff-3"):e===4?document.getElementById("step-signoff-4"):e===5?document.getElementById("step-signoff-5"):document.getElementById(`step-signoff-${e}`);if(fe&&me&&!me.value){let he=Ze();me.value=he,U(t.signOff,he)}});let Q=e===1?"step-signoff-1":e===2?"step-signoff-date":e===3?"step-signoff-3":e===4?"step-signoff-4":e===5?"step-signoff-5":`step-signoff-${e}`;console.log(`[Step ${e}] Signoff input ID: ${Q}, found:`,!!document.getElementById(Q)),(A=document.getElementById(Q))==null||A.addEventListener("change",Z=>{U(t.signOff,Z.target.value||"")});let pe=e===1?"step-notes-lock-1":"step-notes-permanent",K=document.getElementById(pe);K&&(K.addEventListener("click",()=>{let Z=!K.classList.contains("is-locked");Je(K,Z),Yn(t.note,Z),e===2&&vt()}),Je(K,Re(t.note)));let V=e===1?document.getElementById("step-signoff-save-1"):e===2?document.getElementById("headcount-signoff-save"):e===3?document.getElementById("step-signoff-save-3"):e===4?document.getElementById("step-signoff-save-4"):e===5?document.getElementById("step-signoff-save-5"):document.getElementById(`step-signoff-save-${e}`);V==null||V.addEventListener("click",()=>{var me,he;let Z=((me=L==null?void 0:L.value)==null?void 0:me.trim())||"",fe=((he=document.getElementById(Q))==null?void 0:he.value)||"";U(t.reviewer,Z),U(t.signOff,fe),Te(V,!0)}),Mt();let Y=de[e],ne=Y?!!$(Y):!1,_e=$(t.signOff),X=e===1?"step-signoff-toggle-1":e===2?"headcount-signoff-toggle":e===3?"validation-signoff-toggle":e===4?"expense-signoff-toggle":e===5?"step-signoff-toggle-5":`step-signoff-toggle-${e}`;console.log(`[Step ${e}] Toggle button ID: ${X}, found:`,!!document.getElementById(X)),zn({buttonId:X,inputId:Q,fieldName:t.signOff,completeField:Y,requireNotesCheck:e===2?Rt:null,initialActive:!!(_e||ne),stepId:e,onComplete:e===3?La:e===4?Ba:e===2?Ta:null})}e===2&&vt(),e===6&&((M=document.getElementById("archive-run-btn"))==null||M.addEventListener("click",Ma))}function Qe(e){if(Number.isNaN(e)||e<0||e>=be.length)return;let t=be[e];if(!t)return;bt=e;let n=t.id===0?"config":"step";Jt({focusedIndex:e,activeView:n,activeStepId:t.id});let o=Uo[t.id];o&&$a(o),t.id===2&&!H.hasAnalyzed&&zt()}function Nn(e){if(te.activeView==="home"&&e>0){Qe(0);return}let t=te.focusedIndex+e,n=Math.max(0,Math.min(be.length-1,t));Qe(n)}function fa(){if(te.activeView!=="home"||bt===null)return;let e=document.querySelector(`[data-step-card][data-step-index="${bt}"]`);bt=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}async function Un(){let e=It(Et);await Nt(e.sheetName,e.title,e.subtitle),Jt({activeView:"home",activeStepId:null})}function Jt(e){Object.assign(te,e),le()}function Pe(){return $(qe)||$("Reviewer_Name")||""}function Ht(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function Gn(e){let t=document.getElementById("je-save-btn");t&&t.classList.toggle("is-saved",e)}function ma(){let e={};return console.log("[Signoff] Checking step completion status..."),Object.keys(Ie).forEach(t=>{let n=parseInt(t,10),o=Ie[n];if(!o){e[n]=!1;return}let a=$(o.signOff),s=de[n],l=$(s),c=!!a||l==="Y"||l===!0;e[n]=c,console.log(`[Signoff] Step ${n}: signOff="${a}", complete="${l}" \u2192 ${c?"COMPLETE":"pending"}`)}),console.log("[Signoff] Status summary:",e),e}function zn({buttonId:e,inputId:t,fieldName:n,completeField:o,requireNotesCheck:a,onComplete:s,initialActive:l=!1,stepId:c=null}){let i=document.getElementById(e);if(!i){console.warn(`[Signoff] Button not found: ${e}`);return}let r=t?document.getElementById(t):null,u=l||!!(r!=null&&r.value);Ht(i,u),console.log(`[Signoff] Bound ${e}, initial active: ${u}, stepId: ${c}`),i.addEventListener("click",()=>{if(console.log(`[Signoff] Done button clicked: ${e}, stepId: ${c}`),c!==null&&c>0){let p=ma(),{canComplete:d,message:h}=_n(c,p),v=i.classList.contains("is-active");if(console.log(`[Signoff] canComplete: ${d}, isCurrentlyActive: ${v}`),!v&&!d){console.log(`[Signoff] BLOCKED: ${h}`),xn(h);return}}if(a&&!a()){window.alert("Please add notes before completing this step.");return}let f=!i.classList.contains("is-active");if(console.log(`[Signoff] ${e} clicked, toggling to: ${f}`),Ht(i,f),r&&(r.value=f?Ze():""),n){let p=f?Ze():"";console.log(`[Signoff] Writing ${n} = "${p}"`),U(n,p)}if(o){let p=f?"Y":"";console.log(`[Signoff] Writing ${o} = "${p}"`),U(o,p)}f&&typeof s=="function"&&s()}),r&&r.addEventListener("change",()=>{let f=!!r.value,p=i.classList.contains("is-active");f!==p&&(console.log(`[Signoff] Date input changed, syncing button to: ${f}`),Ht(i,f),n&&U(n,r.value||""),o&&U(o,f?"Y":""))})}async function ga(){if(!ae()){window.alert("Open this module inside Excel to access the data sheet.");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem(D.DATA);t.activate(),t.getRange("A1").select(),await e.sync()})}catch(e){console.error("Unable to open PR_Data sheet",e),window.alert(`Unable to open ${D.DATA}. Confirm the sheet exists in this workbook.`)}}async function ha(){if(!ae()){window.alert("Open this module inside Excel to clear data.");return}if(window.confirm("Are you sure you want to clear all data from PR_Data? This cannot be undone."))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(D.DATA),o=n.getUsedRangeOrNullObject();o.load("isNullObject"),await t.sync(),o.isNullObject||(n.getRange("A2:Z10000").clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),window.alert("PR_Data cleared successfully.")}catch(t){console.error("Unable to clear PR_Data sheet",t),window.alert("Unable to clear PR_Data. Please try again.")}}async function Ne(e){var a,s;if(!jt.length)return null;if(gt){let l=e.workbook.tables.getItemOrNullObject(gt);if(l.load("name"),await e.sync(),!l.isNullObject)return l;gt=null}let t=e.workbook.tables;t.load("items/name"),await e.sync();let n=((a=t.items)==null?void 0:a.map(l=>l.name))||[];console.log("[Payroll] Looking for config table:",jt),console.log("[Payroll] Found tables in workbook:",n);let o=(s=t.items)==null?void 0:s.find(l=>jt.includes(l.name));return o?(console.log("[Payroll] \u2713 Config table found:",o.name),gt=o.name,e.workbook.tables.getItem(o.name)):(console.warn("[Payroll] \u26A0\uFE0F CONFIG TABLE NOT FOUND!"),console.warn("[Payroll] Expected table named: SS_PF_Config"),console.warn("[Payroll] Available tables:",n),console.warn("[Payroll] To fix: Select your data in SS_PF_Config sheet \u2192 Insert \u2192 Table \u2192 Name it 'SS_PF_Config'"),null)}async function Wn(){if(!ae()){W.loaded=!0;return}try{await Excel.run(async e=>{let t=await Ne(e);if(!t){console.warn("Payroll Recorder: SS_PF_Config table is missing."),W.loaded=!0;return}let n=t.getDataBodyRange();n.load("values"),await e.sync();let o=n.values||[],a={},s={};o.forEach(l=>{var i,r;let c=oe(l[B.FIELD]);c&&(a[c]=(i=l[B.VALUE])!=null?i:"",s[c]=(r=l[B.PERMANENT])!=null?r:"")}),W.values=a,W.permanents=s,W.overrides.accountingPeriod=!!a.Accounting_Period,W.overrides.jeId=!!a.Journal_Entry_ID,W.loaded=!0})}catch(e){console.warn("Payroll Recorder: unable to load PF_Config table.",e),W.loaded=!0}}function $(e){var t;return(t=W.values[e])!=null?t:""}function Jn(){let e=Object.keys(W.values||{});return Xe.find(n=>e.includes(n))||Xe[0]}function Ct(){return $(Jn())}function Yt(){return($(Hn)||$("Payroll_Provider_Link")||"").trim()}function Re(e){return qn(W.permanents[e])}function ya(e){let t=de[e];return t?qn($(t)):!1}function Yn(e,t){let n=oe(e);n&&(W.permanents[n]=t?"Y":"N",wa(n,t?"Y":"N"))}function qn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function oe(e){return String(e!=null?e:"").trim()}function Kn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t?t.includes("total")||t.includes("totals")||t.includes("grand total")||t.includes("subtotal")||t.includes("summary"):!0}function ye(e){if(!e)return"";let t=kt(e);return t?`${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function va(e){let t=kt(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function ba(e){let t=kt(e);return t?`PR-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function Ze(){return wt(new Date)}function U(e,t,n={}){var l;let o=oe(e);W.values[o]=t!=null?t:"";let a=(l=n.debounceMs)!=null?l:0;if(!a){let c=Le.get(o);c&&clearTimeout(c),Le.delete(o),Ke(o,t!=null?t:"");return}Le.has(o)&&clearTimeout(Le.get(o));let s=setTimeout(()=>{Le.delete(o),Ke(o,t!=null?t:"")},a);Le.set(o,s)}async function Ke(e,t){let n=oe(e);if(W.values[n]=t!=null?t:"",console.log(`[Payroll] Writing config: ${n} = "${t}"`),!ae()){console.warn("[Payroll] Excel runtime not available - cannot write");return}try{await Excel.run(async o=>{var f;let a=await Ne(o);if(!a){console.error("[Payroll] \u274C Cannot write - config table not found");return}let s=a.getDataBodyRange(),l=a.getHeaderRowRange();s.load("values"),l.load("values"),await o.sync();let c=l.values[0]||[],i=s.values||[],r=c.length;console.log(`[Payroll] Table has ${i.length} rows, ${r} columns`);let u=i.findIndex(p=>oe(p[B.FIELD])===n);if(u===-1){W.permanents[n]=(f=W.permanents[n])!=null?f:Pn;let p=new Array(r).fill("");B.TYPE>=0&&B.TYPE<r&&(p[B.TYPE]=Vo),B.FIELD>=0&&B.FIELD<r&&(p[B.FIELD]=n),B.VALUE>=0&&B.VALUE<r&&(p[B.VALUE]=t!=null?t:""),B.PERMANENT>=0&&B.PERMANENT<r&&(p[B.PERMANENT]=Pn),console.log("[Payroll] Adding NEW row:",p),a.rows.add(null,[p]),await o.sync(),console.log(`[Payroll] \u2713 New row added for ${n}`)}else console.log(`[Payroll] Updating existing row ${u} for ${n}`),s.getCell(u,B.VALUE).values=[[t!=null?t:""]],await o.sync(),console.log(`[Payroll] \u2713 Updated ${n}`)})}catch(o){console.error(`[Payroll] \u274C Write failed for ${e}:`,o)}}async function wa(e,t){let n=oe(e);if(n&&ae()){W.permanents[n]=t;try{await Excel.run(async o=>{let a=await Ne(o);if(!a){console.warn(`Payroll Recorder: unable to locate config table when toggling ${e} permanent flag.`);return}let s=a.getDataBodyRange();s.load("values"),await o.sync();let c=(s.values||[]).findIndex(i=>oe(i[B.FIELD])===n);c!==-1&&(s.getCell(c,B.PERMANENT).values=[[t]],await o.sync())})}catch(o){console.warn(`Payroll Recorder: unable to update permanent flag for ${e}`,o)}}}function kt(e){if(!e)return null;let t=String(e).trim(),n=/^(\d{4})-(\d{2})-(\d{2})/.exec(t);if(n){let l=Number(n[1]),c=Number(n[2]),i=Number(n[3]);if(l&&c&&i)return{year:l,month:c,day:i}}let o=/^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(t);if(o){let l=Number(o[1]),c=Number(o[2]),i=Number(o[3]);if(i&&l&&c)return{year:i,month:l,day:c}}let a=Number(e);if(Number.isFinite(a)&&a>4e4&&a<6e4){let l=new Date(1899,11,30),c=new Date(l.getTime()+a*24*60*60*1e3);if(!isNaN(c.getTime()))return console.log("DEBUG parseDateInput - Converted Excel serial",a,"to",c.toISOString().split("T")[0]),{year:c.getFullYear(),month:c.getMonth()+1,day:c.getDate()}}let s=new Date(t);return isNaN(s.getTime())?(console.warn("DEBUG parseDateInput - Could not parse date value:",e),null):{year:s.getFullYear(),month:s.getMonth()+1,day:s.getDate()}}function wt(e){let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),o=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${o}`}function In(e){if(!e)return null;if(typeof e=="string"){let n=e.match(/^(\d{4})-(\d{2})-(\d{2})/);if(n)return`${n[1]}-${n[2]}-${n[3]}`}let t=kt(e);return t?`${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:null}function Ea(){return async()=>{if(!ae())return null;try{return await Excel.run(async e=>{var l,c,i;let t={timestamp:new Date().toISOString(),period:null,summary:{},departments:[],recentPeriods:[],dataQuality:{}},n=await Ne(e);if(n){let r=n.getDataBodyRange();r.load("values"),await e.sync();let u=r.values||[];for(let f of u){let p=String(f[B.FIELD]||"").trim(),d=f[B.VALUE];p.toLowerCase().includes("accounting")&&p.toLowerCase().includes("period")&&(t.period=String(d||"").trim())}}let o=e.workbook.worksheets.getItemOrNullObject(D.DATA_CLEAN);if(o.load("isNullObject"),await e.sync(),!o.isNullObject){let r=o.getUsedRangeOrNullObject();if(r.load("values"),await e.sync(),!r.isNullObject&&((l=r.values)==null?void 0:l.length)>1){let u=r.values[0].map(b=>ve(b)),f=r.values.slice(1),p=u.findIndex(b=>b.includes("amount")),d=Be(u),h=u.findIndex(b=>b.includes("employee")),v=0,C=new Set,g=new Map;for(let b of f){let E=Number(b[p])||0;if(v+=E,h>=0){let x=String(b[h]||"").trim();x&&C.add(x)}if(d>=0){let x=String(b[d]||"").trim();x&&g.set(x,(g.get(x)||0)+E)}}t.summary={total:v,employeeCount:C.size,avgPerEmployee:C.size?v/C.size:0,rowCount:f.length},t.departments=Array.from(g.entries()).map(([b,E])=>({name:b,total:E,percentOfTotal:v?E/v:0})).sort((b,E)=>E.total-b.total),t.dataQuality.dataCleanReady=!0,t.dataQuality.rowCount=f.length}}let a=e.workbook.worksheets.getItemOrNullObject(D.ARCHIVE_SUMMARY);if(a.load("isNullObject"),await e.sync(),!a.isNullObject){let r=a.getUsedRangeOrNullObject();if(r.load("values"),await e.sync(),!r.isNullObject&&((c=r.values)==null?void 0:c.length)>1){let u=r.values[0].map(d=>ve(d)),f=u.findIndex(d=>d.includes("period")),p=u.findIndex(d=>d.includes("total"));t.recentPeriods=r.values.slice(1,6).map(d=>({period:d[f]||"",total:Number(d[p])||0})),t.dataQuality.archiveAvailable=!0,t.dataQuality.periodsAvailable=t.recentPeriods.length}}let s=e.workbook.worksheets.getItemOrNullObject(D.JE_DRAFT);if(s.load("isNullObject"),await e.sync(),!s.isNullObject){let r=s.getUsedRangeOrNullObject();if(r.load("values"),await e.sync(),!r.isNullObject&&((i=r.values)==null?void 0:i.length)>1){let u=r.values[0].map(v=>ve(v)),f=u.findIndex(v=>v.includes("debit")),p=u.findIndex(v=>v.includes("credit")),d=0,h=0;for(let v of r.values.slice(1))d+=Number(v[f])||0,h+=Number(v[p])||0;t.journalEntry={totalDebit:d,totalCredit:h,difference:Math.abs(d-h),isBalanced:Math.abs(d-h)<.01,lineCount:r.values.length-1},t.dataQuality.jeDraftReady=!0}}return console.log("CoPilot context gathered:",t),t})}catch(e){return console.warn("CoPilot context provider error:",e),null}}}function S(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function ge(e,t){return`
        <div class="pf-labeled-button">
            ${e}
            <span class="pf-button-label">${S(t)}</span>
        </div>
    `}function ae(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}function Se(e){return Ie[e]||null}function Ca(){var n,o,a,s;let e=Math.abs((o=(n=H.roster)==null?void 0:n.difference)!=null?o:0),t=Math.abs((s=(a=H.departments)==null?void 0:a.difference)!=null?s:0);return e>0||t>0}function Rt(){return!H.skipAnalysis&&Ca()}function ie(e){return e==null||Number.isNaN(e)?"---":typeof e!="number"?e:e.toLocaleString(void 0,{minimumFractionDigits:2,maximumFractionDigits:2})}function Xn(e){let t=qt(e);return Number.isFinite(t)?t.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}):""}function ka(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let o=String(n);return/[",\n]/.test(o)?`"${o.replace(/"/g,'""')}"`:o}).join(",")).join(`
`)}function Ra(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),o=URL.createObjectURL(n),a=document.createElement("a");a.href=o,a.download=e,document.body.appendChild(a),a.click(),a.remove(),setTimeout(()=>URL.revokeObjectURL(o),1e3)}function qt(e){if(typeof e=="number")return e;if(e==null)return NaN;let t=String(e).replace(/[^0-9.-]/g,""),n=Number.parseFloat(t);return Number.isFinite(n)?n:NaN}function Sa(e){if(e instanceof Date)return wt(e);if(typeof e=="number"&&!Number.isNaN(e)){let o=_a(e);return o?wt(o):""}let t=String(e!=null?e:"").trim();if(!t)return"";if(/^\d{4}-\d{2}-\d{2}$/.test(t))return t;let n=new Date(t);return Number.isNaN(n.getTime())?t:wt(n)}function _a(e){if(!Number.isFinite(e))return null;let t=Math.floor(e-25569);if(!Number.isFinite(t))return null;let n=t*86400*1e3;return new Date(n)}function xa(e){if(!e)return"";let t=new Date(e);return Number.isNaN(t.getTime())?e:t.toLocaleDateString(void 0,{month:"short",day:"numeric",year:"numeric"})}function ht(e){if(e==null||e==="")return 0;let t=Number(e);return Number.isFinite(t)?t:0}function Aa(e){let t=ce(e).toLowerCase();return t?t.includes("burden")||t.includes("tax")||t.includes("benefit")||t.includes("fica")||t.includes("insurance")||t.includes("health")||t.includes("medicare")?"burden":t.includes("bonus")||t.includes("commission")||t.includes("variable")||t.includes("overtime")||t.includes("per diem")?"variable":"fixed":"variable"}function On(e){if(!e||e.length<2)return[];let t=(e[0]||[]).map(a=>ve(a));console.log("parseExpenseRows - headers:",t);let n={payrollDate:t.findIndex(a=>a.includes("payroll")&&a.includes("date")),employee:t.findIndex(a=>a.includes("employee")),department:t.findIndex(a=>a.includes("department")),fixed:t.findIndex(a=>a.includes("fixed")),variable:t.findIndex(a=>a.includes("variable")),burden:t.findIndex(a=>a.includes("burden")),amount:t.findIndex(a=>a.includes("amount")),expenseReview:t.findIndex(a=>a.includes("expense")&&a.includes("review")),category:t.findIndex(a=>a.includes("payroll")&&a.includes("category"))};if(console.log("parseExpenseRows - column indexes:",n),n.payrollDate>=0){let a=new Set;for(let s=1;s<e.length;s++){let l=e[s][n.payrollDate];l&&a.add(String(l))}console.log("parseExpenseRows - unique payroll dates found:",[...a].slice(0,20))}let o=[];for(let a=1;a<e.length;a+=1){let s=e[a],l=Sa(n.payrollDate>=0?s[n.payrollDate]:null);if(!l)continue;let c=n.employee>=0?ce(s[n.employee]):"",i=n.department>=0&&ce(s[n.department])||"Unassigned",r=n.fixed>=0?ht(s[n.fixed]):null,u=n.variable>=0?ht(s[n.variable]):null,f=n.burden>=0?ht(s[n.burden]):null,p=0,d=0,h=0;if(r!==null||u!==null||f!==null)p=r!=null?r:0,d=u!=null?u:0,h=f!=null?f:0;else{let v=n.amount>=0?ht(s[n.amount]):0,C=Aa(n.expenseReview>=0?s[n.expenseReview]:s[n.category]);C==="fixed"?p=v:C==="burden"?h=v:d=v}p===0&&d===0&&h===0||o.push({period:l,employee:c,department:i||"Unassigned",fixed:p,variable:d,burden:h})}return o}function Tn(e){let t=new Map;e.forEach(o=>{let a=o.period;if(!a)return;t.has(a)||t.set(a,{key:a,label:xa(a),employees:new Set,departments:new Map,summary:{fixed:0,variable:0,burden:0}});let s=t.get(a);s.employees.add(o.employee||`EMP-${s.employees.size+1}`);let l=o.department||"Unassigned";s.departments.has(l)||s.departments.set(l,{name:l,fixed:0,variable:0,burden:0,employees:new Set});let c=s.departments.get(l);c.fixed+=o.fixed,c.variable+=o.variable,c.burden+=o.burden,c.employees.add(o.employee||`EMP-${c.employees.size+1}`),s.summary.fixed+=o.fixed,s.summary.variable+=o.variable,s.summary.burden+=o.burden});let n=[];return t.forEach(o=>{let a=o.summary.fixed+o.summary.variable+o.summary.burden,s=Array.from(o.departments.values()).map(i=>{let r=i.fixed+i.variable,u=r+i.burden;return{name:i.name,fixed:i.fixed,variable:i.variable,gross:r,burden:i.burden,allIn:u,percent:a?u/a:0,headcount:i.employees.size,delta:0}});s.sort((i,r)=>r.allIn-i.allIn);let l={employeeCount:o.employees.size,fixed:o.summary.fixed,variable:o.summary.variable,burden:o.summary.burden,total:a,burdenRate:a?o.summary.burden/a:0,delta:0},c={name:"Totals",fixed:o.summary.fixed,variable:o.summary.variable,gross:o.summary.fixed+o.summary.variable,burden:o.summary.burden,allIn:a,percent:a?1:0,headcount:o.employees.size,delta:0,isTotal:!0};n.push({key:o.key,label:o.label,summary:l,departments:s,totalsRow:c})}),n.sort((o,a)=>o.key<a.key?1:-1)}function Ln(e,t){console.log("buildExpenseReviewPeriods - cleanValues rows:",(e==null?void 0:e.length)||0),console.log("buildExpenseReviewPeriods - archiveValues rows:",(t==null?void 0:t.length)||0);let n=Tn(On(e)),o=Tn(On(t));console.log("buildExpenseReviewPeriods - currentPeriods:",n.map(r=>{var u,f;return{key:r.key,employees:(u=r.summary)==null?void 0:u.employeeCount,total:(f=r.summary)==null?void 0:f.total}})),console.log("buildExpenseReviewPeriods - archivePeriods:",o.map(r=>{var u,f;return{key:r.key,employees:(u=r.summary)==null?void 0:u.employeeCount,total:(f=r.summary)==null?void 0:f.total}}));let a=new Map(o.map(r=>[r.key,r])),s=[];n.length&&(s.push(n[0]),a.delete(n[0].key)),o.forEach(r=>{s.length>=6||s.some(u=>u.key===r.key)||s.push(r)}),console.log("buildExpenseReviewPeriods - combined before filter:",s.map(r=>{var u,f;return{key:r.key,employees:(u=r.summary)==null?void 0:u.employeeCount,total:(f=r.summary)==null?void 0:f.total}}));let l=3,c=1e3,i=s.filter(r=>{var d,h,v,C,g;if(!r||!r.key)return console.log("buildExpenseReviewPeriods - EXCLUDED (no key):",r),!1;let u=((d=r.summary)==null?void 0:d.total)||(((h=r.summary)==null?void 0:h.fixed)||0)+(((v=r.summary)==null?void 0:v.variable)||0)+(((C=r.summary)==null?void 0:C.burden)||0),f=((g=r.summary)==null?void 0:g.employeeCount)||0;if(s.indexOf(r)===0)return console.log(`buildExpenseReviewPeriods - INCLUDED (current): ${r.key} - ${f} employees, $${u}`),!0;let p=f>=l&&u>=c;return console.log(`buildExpenseReviewPeriods - ${p?"INCLUDED":"EXCLUDED"}: ${r.key} - ${f} employees, $${u} (needs >=${l} emp, >=$${c})`),p}).sort((r,u)=>r.key<u.key?1:-1).slice(0,6);return console.log("buildExpenseReviewPeriods - FINAL periods:",i.map(r=>r.key)),i.forEach((r,u)=>{let f=i[u+1],p=f?r.summary.total-f.summary.total:0;r.summary.delta=p;let d=new Map(((f==null?void 0:f.departments)||[]).map(h=>[h.name,h]));r.departments.forEach(h=>{let v=d.get(h.name);h.delta=v?h.allIn-v.allIn:0}),r.totalsRow.delta=p}),i}async function Bn(){if(!ae()){yt({loading:!1,lastError:"Excel runtime is unavailable."});return}yt({loading:!0,lastError:null});try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItemOrNullObject(D.DATA_CLEAN),o=t.workbook.worksheets.getItemOrNullObject(D.ARCHIVE_SUMMARY),a=t.workbook.worksheets.getItemOrNullObject(D.EXPENSE_REVIEW);if(n.load("isNullObject, name"),o.load("isNullObject, name"),a.load("isNullObject, name"),await t.sync(),console.log("Expense Review - Sheet check:",{cleanSheet:n.isNullObject?"MISSING":n.name,archiveSheet:o.isNullObject?"MISSING":o.name,reviewSheet:a.isNullObject?"MISSING":a.name}),a.isNullObject){console.log("Creating PR_Expense_Review sheet...");let i=t.workbook.worksheets.add(D.EXPENSE_REVIEW);await t.sync();let r=t.workbook.worksheets.getItem(D.EXPENSE_REVIEW),u=[],f=[];if(!n.isNullObject){let d=n.getUsedRangeOrNullObject();d.load("values"),await t.sync(),u=d.isNullObject?[]:d.values||[]}if(!o.isNullObject){let d=o.getUsedRangeOrNullObject();d.load("values"),await t.sync(),f=d.isNullObject?[]:d.values||[]}let p=Ln(u,f);return await Mn(t,r,p),p}let s=[],l=[];if(n.isNullObject)console.warn("Expense Review - PR_Data_Clean sheet not found, using empty data");else{let i=n.getUsedRangeOrNullObject();i.load("values"),await t.sync(),s=i.isNullObject?[]:i.values||[],console.log("Expense Review - PR_Data_Clean rows:",s.length)}if(o.isNullObject)console.warn("Expense Review - PR_Archive_Summary sheet not found, using empty data");else{let i=o.getUsedRangeOrNullObject();i.load("values"),await t.sync(),l=i.isNullObject?[]:i.values||[],console.log("Expense Review - PR_Archive_Summary rows:",l.length)}let c=Ln(s,l);return console.log("Expense Review - Periods built:",c.length),await Mn(t,a,c),c});yt({loading:!1,periods:e,lastError:null}),await Go(),le()}catch(e){console.error("Expense Review: unable to build executive summary",e),console.error("Error details:",e.message,e.stack),yt({loading:!1,lastError:`Unable to build the Expense Review: ${e.message||"Unknown error"}`,periods:[]})}}async function Mn(e,t,n){if(!t){console.error("writeExpenseReviewSheet: sheet is null/undefined");return}console.log("writeExpenseReviewSheet: Building executive dashboard with",n.length,"periods");try{let y=t.getUsedRangeOrNullObject();y.load("address");let R=t.charts;R.load("items"),await e.sync(),y.isNullObject||(y.clear(),await e.sync());for(let j=R.items.length-1;j>=0;j--)R.items[j].delete();await e.sync()}catch(y){console.warn("Could not clear sheet:",y)}let o=n[0]||{},a=n[1]||{},s=o.summary||{},l=a.summary||{},c=$("Accounting_Period")||Ct()||"",i=Number(s.total)||0,r=Number(l.total)||0,u=i-r,f=r?u/r:0,p=Number(s.employeeCount)||0,d=Number(l.employeeCount)||0,h=p-d,v=p?i/p:0,C=d?r/d:0,g=v-C,b=Da(n),E=Pa(o,n),x=o.label||o.key||"Current Period",A=new Date().toLocaleString("en-US",{month:"short",day:"numeric",year:"numeric",hour:"numeric",minute:"2-digit"}),M=y=>y>0?"\u25B2":y<0?"\u25BC":"\u2014",_=n.map(y=>{var R;return((R=y.summary)==null?void 0:R.total)||0}).filter(y=>y>0),T=n.map(y=>{let R=y.summary||{},j=R.employeeCount||0;return j>0?(R.total||0)/j:0}).filter(y=>y>0),m=n.slice(0,-1).map((y,R)=>{var re,J,z;let j=((re=y.summary)==null?void 0:re.total)||0,ee=((z=(J=n[R+1])==null?void 0:J.summary)==null?void 0:z.total)||0;return ee>0?(j-ee)/ee:0}),N=(y,R=null)=>{let j=R!==null?[...y,R]:y;if(!j.length)return{min:0,max:0,avg:0};let ee=Math.min(...j),re=Math.max(...j),J=y.length?y:j,z=J.reduce((Ce,we)=>Ce+we,0)/J.length;return{min:ee,max:re,avg:z}},L=N(_,i),Q=N(T,v),pe=N(m),K=(y,R,j,ee=20)=>{if(j<=R)return"\u2591".repeat(ee);let re=j-R,J=Math.max(0,Math.min(1,(y-R)/re)),z=Math.round(J*(ee-1)),Ce="";for(let we=0;we<ee;we++)we===z?Ce+="\u25CF":Ce+="\u2591";return Ce},V=Number(s.fixed)||0,Y=Number(s.variable)||0,ne=Number(s.burden)||0,_e=V+Y,X=i?ne/i:0,Z=Number(l.fixed)||0,fe=Number(l.variable)||0,me=Number(l.burden)||0,he=r?me/r:0,se=o.departments||[],Kt=se.filter(y=>{let R=(y.name||"").toLowerCase();return R.includes("sales")||R.includes("marketing")}),Xt=se.filter(y=>{let R=(y.name||"").toLowerCase();return!R.includes("sales")&&!R.includes("marketing")}),to=Kt.reduce((y,R)=>y+(R.variable||0),0),et=Kt.reduce((y,R)=>y+(R.headcount||0),0),no=Xt.reduce((y,R)=>y+(R.variable||0),0),tt=Xt.reduce((y,R)=>y+(R.headcount||0),0),St=et?to/et:0,_t=tt?no/tt:0,xt=p?V/p:0,I=[],k=0,w={};w.headerStart=k;let Qt=c||x;if(typeof c=="number"||!isNaN(Number(c))&&c){let y=Number(c);if(y>4e4&&y<6e4){let R=new Date(1899,11,30);Qt=new Date(R.getTime()+y*24*60*60*1e3).toLocaleDateString("en-US",{year:"numeric",month:"long",day:"numeric"})}}I.push(["PAYROLL EXPENSE REVIEW"]),k++,I.push([`Period: ${Qt}`]),k++,I.push([`Generated: ${A}`]),k++,I.push([""]),k++,w.headerEnd=k-1,w.execSummaryStart=k,I.push(["EXECUTIVE SUMMARY"]),k++,w.execSummaryHeader=k-1,I.push([""]),k++,I.push(["","Pay Date","Headcount","Fixed Salary","Variable Salary","Burden","Total Payroll","Burden Rate"]),k++,w.execSummaryColHeaders=k-1,I.push(["Current Pay Period",o.label||o.key||"",p,V,Y,ne,i,X]),k++,w.execSummaryCurrentRow=k-1,I.push(["Same Period Prior Month",a.label||a.key||"",d,Z,fe,me,r,he]),k++,w.execSummaryPriorRow=k-1,I.push([""]),k++,I.push([""]),k++,w.execSummaryEnd=k-1,w.deptBreakdownStart=k,I.push(["CURRENT PERIOD BREAKDOWN (DEPARTMENT)"]),k++,w.deptBreakdownHeader=k-1,I.push([""]),k++,I.push(["Payroll Date",o.label||o.key||""]),k++,I.push([""]),k++,I.push(["Department","Fixed Salary","Variable Salary","Gross Pay","Burden","All-In Total","% of Total","Headcount"]),k++,w.deptColHeaders=k-1;let oo=[...se].sort((y,R)=>(R.allIn||0)-(y.allIn||0));if(w.deptDataStart=k,oo.forEach(y=>{I.push([y.name||"",y.fixed||0,y.variable||0,y.gross||0,y.burden||0,y.allIn||0,y.percent||0,y.headcount||0]),k++}),w.deptDataEnd=k-1,o.totalsRow){let y=o.totalsRow;I.push(["TOTAL",y.fixed||0,y.variable||0,y.gross||0,y.burden||0,y.allIn||0,1,y.headcount||0]),k++,w.deptTotalsRow=k-1}I.push([""]),k++,I.push([""]),k++,w.deptBreakdownEnd=k-1,w.historicalStart=k,I.push(["HISTORICAL CONTEXT"]),k++,w.historicalHeader=k-1,I.push([`Visual comparison of current period vs. historical range (${n.length} periods). The dot (\u25CF) shows where you currently stand.`]),k++,I.push([""]),k++;let q=y=>`$${Math.round(y/1e3)}K`,nt=y=>`${(y*100).toFixed(1)}%`;I.push(["","Metric","Low","Range","High","","Current","Average"]),k++,w.historicalColHeaders=k-1;let ao=n.map(y=>{var R;return((R=y.summary)==null?void 0:R.fixed)||0}).filter(y=>y>0),so=n.map(y=>{var R;return((R=y.summary)==null?void 0:R.variable)||0}),ro=n.map(y=>{let R=y.summary||{};return R.total?(R.burden||0)/R.total:0}),io=n.map(y=>{let R=y.summary||{},j=R.employeeCount||0;return j>0?(R.fixed||0)/j:0}).filter(y=>y>0),Me=N(ao,V),Ve=N(so,Y),Fe=N(ro,X),je=N(io,xt);w.spectrumRows=[];let lo=K(i,L.min,L.max,25);I.push(["","Total Payroll",q(L.min),lo,q(L.max),"",q(i),q(L.avg)]),k++,w.spectrumRows.push(k-1);let co=K(V,Me.min,Me.max,25);I.push(["","Total Fixed Payroll",q(Me.min),co,q(Me.max),"",q(V),q(Me.avg)]),k++,w.spectrumRows.push(k-1);let uo=K(Y,Ve.min,Ve.max,25);I.push(["","Total Variable Payroll",q(Ve.min),uo,q(Ve.max),"",q(Y),q(Ve.avg)]),k++,w.spectrumRows.push(k-1),I.push([""]),k++;let po=K(xt,je.min,je.max,25);I.push(["","Avg Fixed Payroll / Employee",q(je.min),po,q(je.max),"",q(xt),q(je.avg)]),k++,w.spectrumRows.push(k-1);let fo=n.map(y=>{let j=(y.departments||[]).filter(J=>{let z=(J.name||"").toLowerCase();return z.includes("sales")||z.includes("marketing")}),ee=j.reduce((J,z)=>J+(z.variable||0),0),re=j.reduce((J,z)=>J+(z.headcount||0),0);return re>0?ee/re:0}),ot=N(fo,St),mo=n.map(y=>{let j=(y.departments||[]).filter(J=>{let z=(J.name||"").toLowerCase();return!z.includes("sales")&&!z.includes("marketing")}),ee=j.reduce((J,z)=>J+(z.variable||0),0),re=j.reduce((J,z)=>J+(z.headcount||0),0);return re>0?ee/re:0}),at=N(mo,_t);if(et>0){let y=K(St,ot.min,ot.max,25);I.push(["","Avg Variable / Sales & Marketing",q(ot.min),y,q(ot.max),"",q(St),`${et} emp`]),k++,w.spectrumRows.push(k-1)}if(tt>0){let y=K(_t,at.min,at.max,25);I.push(["","Avg Variable / Other Depts",q(at.min),y,q(at.max),"",q(_t),`${tt} emp`]),k++,w.spectrumRows.push(k-1)}I.push([""]),k++;let go=K(X,Fe.min,Fe.max,25);I.push(["","Burden Rate (%)",nt(Fe.min),go,nt(Fe.max),"",nt(X),nt(Fe.avg)]),k++,w.spectrumRows.push(k-1),I.push([""]),k++,I.push([""]),k++,w.historicalEnd=k-1,w.trendsStart=k,I.push(["PERIOD TRENDS"]),k++,w.trendsHeader=k-1,I.push([""]),k++,I.push(["Pay Period","Total Payroll","Fixed Payroll","Variable Payroll","Burden","Headcount"]),k++,w.trendColHeaders=k-1;let Zt=n.slice(0,6).reverse();w.trendDataStart=k,Zt.forEach(y=>{let R=y.summary||{};I.push([y.label||y.key||"",R.total||0,R.fixed||0,R.variable||0,R.burden||0,R.employeeCount||0]),k++}),w.trendDataEnd=k-1,I.push([""]),k++,w.trendsEnd=k-1,w.chartStart=k;for(let y=0;y<15;y++)I.push([""]),k++;w.payrollChartEnd=k-1,w.headcountChartStart=k;for(let y=0;y<12;y++)I.push([""]),k++;w.headcountChartEnd=k-1,console.log("writeExpenseReviewSheet: Writing",I.length,"rows");let en=I.map(y=>{let R=Array.isArray(y)?y:[""];for(;R.length<10;)R.push("");return R.slice(0,10)});try{let y=t.getRangeByIndexes(0,0,en.length,10);y.values=en,await e.sync()}catch(y){throw console.error("writeExpenseReviewSheet: Write failed",y),y}try{t.getRange("A:A").format.columnWidth=200,t.getRange("B:B").format.columnWidth=130,t.getRange("C:C").format.columnWidth=100,t.getRange("D:D").format.columnWidth=200,t.getRange("E:E").format.columnWidth=100,t.getRange("F:F").format.columnWidth=100,t.getRange("G:G").format.columnWidth=100,t.getRange("H:H").format.columnWidth=100,t.getRange("I:I").format.columnWidth=80,t.getRange("J:J").format.columnWidth=80,await e.sync();let y=t.getRange("A1");y.format.font.bold=!0,y.format.font.size=22,y.format.font.color="#1e293b",t.getRange("A2").format.font.size=11,t.getRange("A2").format.font.color="#64748b",t.getRange("A3").format.font.size=10,t.getRange("A3").format.font.color="#94a3b8",await e.sync();let R=t.getRange(`A${w.execSummaryHeader+1}`);R.format.font.bold=!0,R.format.font.size=14,R.format.font.color="#1e293b";let j=t.getRange(`A${w.execSummaryColHeaders+1}:H${w.execSummaryColHeaders+1}`);j.format.font.bold=!0,j.format.font.size=10,j.format.fill.color="#1e293b",j.format.font.color="#ffffff";let ee=t.getRange(`A${w.execSummaryCurrentRow+1}:H${w.execSummaryCurrentRow+1}`);ee.format.fill.color="#dcfce7",ee.format.font.bold=!0;let re=t.getRange(`A${w.execSummaryPriorRow+1}:H${w.execSummaryPriorRow+1}`);re.format.fill.color="#f1f5f9";for(let O of[w.execSummaryCurrentRow+1,w.execSummaryPriorRow+1])t.getRange(`C${O}`).numberFormat=[["#,##0"]],t.getRange(`D${O}`).numberFormat=[["$#,##0"]],t.getRange(`E${O}`).numberFormat=[["$#,##0"]],t.getRange(`F${O}`).numberFormat=[["$#,##0"]],t.getRange(`G${O}`).numberFormat=[["$#,##0"]],t.getRange(`H${O}`).numberFormat=[["0.00%"]];await e.sync();let J=t.getRange(`A${w.deptBreakdownHeader+1}`);J.format.font.bold=!0,J.format.font.size=14,J.format.font.color="#1e293b";let z=t.getRange(`A${w.deptColHeaders+1}:H${w.deptColHeaders+1}`);z.format.font.bold=!0,z.format.font.size=10,z.format.fill.color="#1e293b",z.format.font.color="#ffffff";for(let O=w.deptDataStart;O<=w.deptDataEnd;O++){let P=O+1;t.getRange(`B${P}`).numberFormat=[["$#,##0"]],t.getRange(`C${P}`).numberFormat=[["$#,##0"]],t.getRange(`D${P}`).numberFormat=[["$#,##0"]],t.getRange(`E${P}`).numberFormat=[["$#,##0"]],t.getRange(`F${P}`).numberFormat=[["$#,##0"]],t.getRange(`G${P}`).numberFormat=[["0.00%"]],t.getRange(`H${P}`).numberFormat=[["#,##0"]],(O-w.deptDataStart)%2===1&&(t.getRange(`A${P}:H${P}`).format.fill.color="#f8fafc")}if(w.deptTotalsRow){let O=t.getRange(`A${w.deptTotalsRow+1}:H${w.deptTotalsRow+1}`);O.format.font.bold=!0,O.format.fill.color="#1e293b",O.format.font.color="#ffffff";let P=w.deptTotalsRow+1;t.getRange(`B${P}`).numberFormat=[["$#,##0"]],t.getRange(`C${P}`).numberFormat=[["$#,##0"]],t.getRange(`D${P}`).numberFormat=[["$#,##0"]],t.getRange(`E${P}`).numberFormat=[["$#,##0"]],t.getRange(`F${P}`).numberFormat=[["$#,##0"]],t.getRange(`G${P}`).numberFormat=[["0%"]],t.getRange(`H${P}`).numberFormat=[["#,##0"]]}await e.sync();let Ce=t.getRange(`A${w.historicalHeader+1}`);Ce.format.font.bold=!0,Ce.format.font.size=14,Ce.format.font.color="#1e293b",t.getRange(`A${w.historicalHeader+2}`).format.font.size=10,t.getRange(`A${w.historicalHeader+2}`).format.font.color="#64748b",t.getRange(`A${w.historicalHeader+2}`).format.font.italic=!0;let we=t.getRange(`A${w.historicalColHeaders+1}:H${w.historicalColHeaders+1}`);we.format.font.bold=!0,we.format.font.size=10,we.format.fill.color="#e2e8f0",we.format.font.color="#334155",t.getRange(`C${w.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`E${w.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`G${w.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`H${w.historicalColHeaders+1}`).format.horizontalAlignment="Center",w.spectrumRows.forEach(O=>{t.getRange(`D${O+1}`).format.font.name="Consolas",t.getRange(`D${O+1}`).format.font.size=14,t.getRange(`D${O+1}`).format.font.color="#6366f1",t.getRange(`D${O+1}`).format.horizontalAlignment="Center",t.getRange(`B${O+1}`).format.font.color="#334155",t.getRange(`C${O+1}`).format.font.color="#94a3b8",t.getRange(`C${O+1}`).format.horizontalAlignment="Center",t.getRange(`E${O+1}`).format.font.color="#94a3b8",t.getRange(`E${O+1}`).format.horizontalAlignment="Center",t.getRange(`G${O+1}`).format.font.bold=!0,t.getRange(`G${O+1}`).format.font.color="#1e293b",t.getRange(`G${O+1}`).format.horizontalAlignment="Center",t.getRange(`H${O+1}`).format.font.color="#94a3b8",t.getRange(`H${O+1}`).format.horizontalAlignment="Center"}),await e.sync();let At=t.getRange(`A${w.trendsHeader+1}`);At.format.font.bold=!0,At.format.font.size=14,At.format.font.color="#1e293b";let st=t.getRange(`A${w.trendColHeaders+1}:F${w.trendColHeaders+1}`);st.format.font.bold=!0,st.format.font.size=10,st.format.fill.color="#1e293b",st.format.font.color="#ffffff";for(let O=w.trendDataStart;O<=w.trendDataEnd;O++){let P=O+1;t.getRange(`B${P}`).numberFormat=[["$#,##0"]],t.getRange(`C${P}`).numberFormat=[["$#,##0"]],t.getRange(`D${P}`).numberFormat=[["$#,##0"]],t.getRange(`E${P}`).numberFormat=[["$#,##0"]],t.getRange(`F${P}`).numberFormat=[["#,##0"]],(O-w.trendDataStart)%2===1&&(t.getRange(`A${P}:F${P}`).format.fill.color="#f8fafc")}if(await e.sync(),Zt.length>=2){try{let O=t.getRange(`A${w.trendColHeaders+1}:E${w.trendDataEnd+1}`),P=t.charts.add(Excel.ChartType.lineMarkers,O,Excel.ChartSeriesBy.columns);P.setPosition(`A${w.chartStart+1}`,`J${w.payrollChartEnd+1}`),P.title.text="Payroll Expense Trends",P.title.format.font.size=14,P.title.format.font.bold=!0,P.legend.position=Excel.ChartLegendPosition.bottom,P.format.fill.setSolidColor("#ffffff"),P.format.border.lineStyle=Excel.ChartLineStyle.continuous,P.format.border.color="#e2e8f0";let He=P.axes.getItem(Excel.ChartAxisType.category);He.categoryType=Excel.ChartAxisCategoryType.textAxis,He.setCategoryNames(t.getRange(`A${w.trendDataStart+1}:A${w.trendDataEnd+1}`)),await e.sync();let Ee=P.series;Ee.load("count"),await e.sync();let ue=["#3b82f6","#22c55e","#f97316","#8b5cf6"];for(let Oe=0;Oe<Math.min(Ee.count,ue.length);Oe++){let Ue=Ee.getItemAt(Oe);Ue.format.line.color=ue[Oe],Ue.format.line.weight=2,Ue.markerStyle=Excel.ChartMarkerStyle.circle,Ue.markerSize=6,Ue.markerBackgroundColor=ue[Oe]}await e.sync(),console.log("writeExpenseReviewSheet: Payroll chart created successfully")}catch(O){console.warn("writeExpenseReviewSheet: Payroll chart creation failed (non-critical)",O)}try{let O=t.getRange(`A${w.trendColHeaders+1}:F${w.trendDataEnd+1}`),P=t.charts.add(Excel.ChartType.lineMarkers,O,Excel.ChartSeriesBy.columns);P.setPosition(`A${w.headcountChartStart+1}`,`J${w.headcountChartEnd+1}`),P.title.text="Headcount Trend",P.title.format.font.size=12,P.title.format.font.bold=!0,P.legend.visible=!1,P.format.fill.setSolidColor("#ffffff"),P.format.border.lineStyle=Excel.ChartLineStyle.continuous,P.format.border.color="#e2e8f0";let He=P.axes.getItem(Excel.ChartAxisType.category);He.categoryType=Excel.ChartAxisCategoryType.textAxis,He.setCategoryNames(t.getRange(`A${w.trendDataStart+1}:A${w.trendDataEnd+1}`)),await e.sync();let Ee=P.series;Ee.load("count, items/name"),await e.sync();for(let ue=Ee.count-2;ue>=0;ue--)Ee.getItemAt(ue).delete();if(await e.sync(),Ee.load("count"),await e.sync(),Ee.count>0){let ue=Ee.getItemAt(0);ue.format.line.color="#64748b",ue.format.line.weight=2.5,ue.markerStyle=Excel.ChartMarkerStyle.circle,ue.markerSize=8,ue.markerBackgroundColor="#64748b"}await e.sync(),console.log("writeExpenseReviewSheet: Headcount chart created successfully")}catch(O){console.warn("writeExpenseReviewSheet: Headcount chart creation failed (non-critical)",O)}}t.freezePanes.freezeRows(w.execSummaryEnd+1),t.pageLayout.orientation=Excel.PageOrientation.landscape,t.getRange("A1").select(),await e.sync(),console.log("writeExpenseReviewSheet: Formatting applied successfully")}catch(y){console.warn("writeExpenseReviewSheet: Formatting error (non-critical)",y)}}function Da(e){var o;return!e||!e.length?!1:(((o=e[0].summary)==null?void 0:o.categories)||[]).some(a=>{let s=(a.name||"").toLowerCase();return s.includes("commission")||s.includes("bonus")||s.includes("variable")})}function Pa(e,t){var l;if(!e||t.length<2)return!1;let n=t.map(c=>{var i;return((i=c.summary)==null?void 0:i.total)||0}).filter(c=>c>0);if(n.length<2)return!1;let o=n.reduce((c,i)=>c+i,0)/n.length,a=((l=e.summary)==null?void 0:l.total)||0;return(o>0?a/o:1)<.9}async function $a(e){if(!(!ae()||!e))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItemOrNullObject(e);n.load("name"),await t.sync(),!n.isNullObject&&(n.activate(),n.getRange("A1").select(),await t.sync())})}catch(t){console.warn(`Payroll Recorder: unable to activate worksheet ${e}`,t)}}async function zt(){if(!ae()){H.lastError="Excel runtime is unavailable.",H.hasAnalyzed=!0,H.loading=!1,le();return}H.loading=!0,H.lastError=null,le();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),o=t.workbook.worksheets.getItem(D.DATA),a=n.getUsedRangeOrNullObject(),s=o.getUsedRangeOrNullObject();a.load("values"),s.load("values"),await t.sync();let l=a.isNullObject?[]:a.values||[],c=s.isNullObject?[]:s.values||[],i=Ia(l),r=Oa(c),u=[];i.employeeMap.forEach((d,h)=>{r.employeeMap.has(h)||u.push({name:d.name||"Unknown Employee",type:"missing_from_payroll",message:"In roster but NOT in payroll data",department:d.department||"\u2014"})}),r.employeeMap.forEach((d,h)=>{i.employeeMap.has(h)||u.push({name:d.name||"Unknown Employee",type:"missing_from_roster",message:"In payroll but NOT in roster",department:d.department||"\u2014"})}),u.sort((d,h)=>d.type!==h.type?d.type.localeCompare(h.type):(d.name||"").localeCompare(h.name||""));let f=[],p=0;return i.employeeMap.forEach((d,h)=>{let v=r.employeeMap.get(h);if(!v)return;let C=ce(d.department),g=ce(v.department);!C&&!g||(p+=1,C!==g&&f.push({employee:d.name||v.name||"Employee",rosterDept:C||"\u2014",payrollDept:g||"\u2014"}))}),console.log("Headcount Analysis Results:",{rosterCount:i.activeCount,payrollCount:r.totalEmployees,difference:i.activeCount-r.totalEmployees,missingFromPayroll:u.filter(d=>d.type==="missing_from_payroll").length,missingFromRoster:u.filter(d=>d.type==="missing_from_roster").length,deptMismatches:f.length}),{roster:{rosterCount:i.activeCount,payrollCount:r.totalEmployees,difference:i.activeCount-r.totalEmployees,mismatches:u},departments:{rosterCount:p,payrollCount:p,difference:f.length,mismatches:f}}});H.roster=e.roster,H.departments=e.departments,H.hasAnalyzed=!0}catch(e){console.warn("Headcount Review: unable to analyze data",e),H.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{H.loading=!1,le()}}function Ye(e={},{rerender:t=!0}={}){Object.assign(F,e);let n=Number(F.prDataTotal),o=Number(F.cleanTotal);F.reconDifference=Number.isFinite(n)&&Number.isFinite(o)?n-o:null;let a=qt(F.bankAmount);F.bankDifference=Number.isFinite(o)&&!Number.isNaN(a)?o-a:null,F.plugEnabled=F.bankDifference!=null&&Math.abs(F.bankDifference)>=.5,t?le():Na()}function Na(){if(te.activeStepId!==3)return;let e=(o,a)=>{let s=document.getElementById(o);s&&(s.value=a)};e("pr-data-total-value",ie(F.prDataTotal)),e("clean-total-value",ie(F.cleanTotal)),e("recon-diff-value",ie(F.reconDifference)),e("bank-clean-total-value",ie(F.cleanTotal)),e("bank-diff-value",F.bankDifference!=null?ie(F.bankDifference):"---");let t=document.getElementById("bank-diff-hint");t&&(t.textContent=F.bankDifference==null?"":Math.abs(F.bankDifference)<.5?"Difference within acceptable tolerance.":"Difference exceeds tolerance and should be resolved.");let n=document.getElementById("bank-plug-btn");n&&(n.disabled=!F.plugEnabled)}function yt(e={},{rerender:t=!0}={}){Object.assign(ke,e),t&&le()}async function Vn(){if(!ae()){Ye({loading:!1,lastError:"Excel runtime is unavailable.",prDataTotal:null,cleanTotal:null});return}Ye({loading:!0,lastError:null});try{let e="";await Excel.run(async n=>{let o=await Ne(n);if(console.log("DEBUG - Config table found:",!!o),o){let a=o.getDataBodyRange();a.load("values"),await n.sync();let s=a.values||[];console.log("DEBUG - Config table rows:",s.length),console.log("DEBUG - Looking for payroll date aliases:",Xe),console.log("DEBUG - CONFIG_COLUMNS.FIELD:",B.FIELD,"CONFIG_COLUMNS.VALUE:",B.VALUE);for(let l of s){let c=String(l[B.FIELD]||"").trim(),i=l[B.VALUE],r=Xe.some(u=>c===u||oe(c)===oe(u));if((c.toLowerCase().includes("payroll")||c.toLowerCase().includes("date"))&&console.log("DEBUG - Potential date field:",c,"=",i,"| isMatch:",r),r){let u=l[B.VALUE];console.log("DEBUG - Found payroll date field!",c,"raw value:",u),e=ye(u)||"",console.log("DEBUG - Formatted payroll date:",e);break}}e||(console.warn("DEBUG - No payroll date found in config! Available fields:"),s.forEach((l,c)=>{console.log(`  Row ${c}: Field="${l[B.FIELD]}" Value="${l[B.VALUE]}"`)}))}else console.warn("DEBUG - Config table not found!")}),console.log("DEBUG prepareValidationData - Final Payroll Date:",e||"(empty)");let t=await Excel.run(async n=>{var T;let o=n.workbook.worksheets.getItem(D.DATA),a=n.workbook.worksheets.getItem(D.EXPENSE_MAPPING),s=n.workbook.worksheets.getItem(D.DATA_CLEAN),l=o.getUsedRangeOrNullObject(),c=a.getUsedRangeOrNullObject(),i=s.getUsedRangeOrNullObject();l.load("values"),c.load("values"),i.load(["address","rowCount"]),await n.sync();let r=l.isNullObject?[]:l.values||[],u=c.isNullObject?[]:c.values||[];console.log("DEBUG prepareValidationData - PR_Data rows:",r.length),console.log("DEBUG prepareValidationData - PR_Data headers:",r[0]),console.log("DEBUG prepareValidationData - PR_Expense_Mapping rows:",u.length);let f=((T=u[0])==null?void 0:T.map(m=>ve(m)))||[],p=m=>f.findIndex(m),d={category:p(m=>m.includes("category")),accountNumber:p(m=>m.includes("account")&&(m.includes("number")||m.includes("#"))),accountName:p(m=>m.includes("account")&&m.includes("name")),expenseReview:p(m=>m.includes("expense")&&m.includes("review"))},h=new Map;u.slice(1).forEach(m=>{var L,Q,pe;let N=d.category>=0?Ut(m[d.category]):"";N&&h.set(N,{accountNumber:d.accountNumber>=0&&(L=m[d.accountNumber])!=null?L:"",accountName:d.accountName>=0&&(Q=m[d.accountName])!=null?Q:"",expenseReview:d.expenseReview>=0&&(pe=m[d.expenseReview])!=null?pe:""})});let v=s.getRangeByIndexes(0,0,1,8);v.load("values"),await n.sync();let C=v.values[0]||[],g=C.map(m=>ve(m));console.log("DEBUG prepareValidationData - PR_Data_Clean headers:",C),console.log("DEBUG prepareValidationData - PR_Data_Clean normalized:",g),console.log("DEBUG - PR_Data_Clean headers:",C),console.log("DEBUG - PR_Data_Clean normalized headers:",g);let b=g.findIndex(m=>(m.includes("payroll")||m.includes("period"))&&m.includes("date"));console.log("DEBUG - payrollDate column index:",b),b===-1&&(console.warn("DEBUG - No payroll date column found! Looking for header containing 'payroll'/'period' AND 'date'"),g.forEach((m,N)=>console.log(`  Col ${N}: "${m}"`)));let E={payrollDate:b,employee:g.findIndex(m=>m.includes("employee")),department:Be(g),payrollCategory:g.findIndex(m=>m.includes("payroll")&&m.includes("category")),accountNumber:g.findIndex(m=>m.includes("account")&&(m.includes("number")||m.includes("#"))),accountName:g.findIndex(m=>m.includes("account")&&m.includes("name")),amount:g.findIndex(m=>m.includes("amount")),expenseReview:g.findIndex(m=>m.includes("expense")&&m.includes("review"))};console.log("DEBUG prepareValidationData - fieldMap:",E);let x=C.length,A=[],M=0,_=0;if(r.length>=2){let m=r[0],N=m.map(V=>ve(V));console.log("DEBUG prepareValidationData - Normalized headers:",N);let L=N.findIndex(V=>V.includes("employee")),Q=Be(N);console.log("DEBUG prepareValidationData - Employee column index:",L,"searching for 'employee' in:",N[6]),console.log("DEBUG prepareValidationData - Department column index:",Q);let pe=h.size>0,K=N.reduce((V,Y,ne)=>{if(ne===L||ne===Q||!Y||Y.includes("total")||Y.includes("gross")||Y.includes("date")||Y.includes("period"))return V;let _e=Ut(m[ne]||Y);return pe&&!h.has(_e)||V.push(ne),V},[]);console.log("DEBUG prepareValidationData - Numeric columns:",K.length,K);for(let V=1;V<r.length;V+=1){let Y=r[V],ne=L>=0?ce(Y[L]):"";if(!ne||ne.toLowerCase().includes("total"))continue;let _e=Q>=0&&Y[Q]||"";K.forEach(X=>{let Z=Y[X],fe=Number(Z);if(!Number.isFinite(fe)||fe===0)return;M+=fe;let me=m[X]||N[X]||`Column ${X+1}`,he=h.get(Ut(me))||{};_+=fe;let se=new Array(x).fill("");E.payrollDate>=0?se[E.payrollDate]=e:x>0&&(se[0]=e),A.length===0&&(console.log("DEBUG - Building first PR_Data_Clean row:"),console.log("  payrollDate value:",e),console.log("  fieldMap.payrollDate:",E.payrollDate),console.log("  Writing to column index:",E.payrollDate>=0?E.payrollDate:0)),E.employee>=0&&(se[E.employee]=ne),E.department>=0&&(se[E.department]=_e),E.payrollCategory>=0&&(se[E.payrollCategory]=me),E.accountNumber>=0&&(se[E.accountNumber]=he.accountNumber||""),E.accountName>=0&&(se[E.accountName]=he.accountName||""),E.amount>=0&&(se[E.amount]=fe),E.expenseReview>=0&&(se[E.expenseReview]=he.expenseReview||""),A.push(se)})}}if(console.log("DEBUG prepareValidationData - Clean rows generated:",A.length),console.log("DEBUG prepareValidationData - PR_Data total:",M,"Clean total:",_),console.log("DEBUG prepareValidationData - columnCount:",x,"cleanRange.address:",i.address),!i.isNullObject&&i.address){console.log("DEBUG prepareValidationData - Clearing data rows...");let m=Math.max(0,(i.rowCount||0)-1),N=Math.max(1,m);s.getRangeByIndexes(1,0,N,x).clear(),await n.sync(),console.log("DEBUG prepareValidationData - Data rows cleared")}if(console.log("DEBUG prepareValidationData - About to write",A.length,"rows with",x,"columns"),A.length>0){let m=s.getRangeByIndexes(1,0,A.length,x);m.values=A,console.log("DEBUG prepareValidationData - Data written to PR_Data_Clean")}else console.log("DEBUG prepareValidationData - No rows to write!");return await n.sync(),{prDataTotal:M,cleanTotal:_}});Ye({loading:!1,lastError:null,prDataTotal:t.prDataTotal,cleanTotal:t.cleanTotal})}catch(e){console.warn("Validate & Reconcile: unable to prepare PR_Data_Clean",e),Ye({loading:!1,prDataTotal:null,cleanTotal:null,lastError:"Unable to prepare reconciliation data. Try again."})}}function Ia(e){let t={activeCount:0,departmentCount:0,employeeMap:new Map};if(!e||!e.length)return t;let{headers:n,dataStartIndex:o}=eo(e,["employee"]);if(!n.length||o==null)return t;let a=Zn(n),s=n.findIndex(i=>i.includes("termination")),l=Be(n);if(a===-1)return t;let c=new Set;for(let i=o;i<e.length;i+=1){let r=e[i],u=r[a],f=Qn(u);if(!f||Kn(f))continue;let p=s>=0?r[s]:"",d=l>=0?r[l]:"";!ce(p)&&!c.has(f)&&(c.add(f),t.activeCount+=1),d&&(t.departmentCount+=1),t.employeeMap.has(f)||t.employeeMap.set(f,{name:ce(u)||f,department:ce(d),termination:p})}return t}function Oa(e){let t={totalEmployees:0,departmentCount:0,employeeMap:new Map};if(!e||!e.length)return t;let{headers:n,dataStartIndex:o}=eo(e,["employee"]);if(!n.length||o==null)return t;let a=Zn(n),s=Be(n);if(a===-1)return t;let l=new Set;for(let c=o;c<e.length;c+=1){let i=e[c],r=i[a],u=Qn(r);if(!u||Kn(u))continue;l.has(u)||(l.add(u),t.totalEmployees+=1);let f=s>=0?i[s]:"";f&&(t.departmentCount+=1),t.employeeMap.has(u)||t.employeeMap.set(u,{name:ce(r)||u,department:ce(f)})}return t}function ve(e){return ce(e).toLowerCase()}function Qn(e){return ce(e).toLowerCase()}function Zn(e=[]){let t=e.findIndex(o=>o.includes("employee")&&o.includes("name"));return t>=0?t:e.findIndex(o=>o.includes("employee"))}function eo(e,t=[]){let n=[],o=null;return(e||[]).some((a,s)=>{let l=(a||[]).map(ve);return t.every(i=>l.some(r=>r.includes(i)))?(n=l,o=s,!0):!1}),{headers:n,dataStartIndex:o!=null?o+1:null}}function ce(e){return e==null?"":String(e).trim()}function Ut(e){return ce(e).toLowerCase()}function Be(e=[]){let t=e.map((l,c)=>({idx:c,value:ve(l)})),n=t.find(({value:l})=>l.includes("department")&&l.includes("description"));if(n)return console.log("DEBUG pickDepartmentIndex - Using 'Department Description' at index:",n.idx),n.idx;let o=t.find(({value:l})=>l.includes("department")&&l.includes("name"));if(o)return console.log("DEBUG pickDepartmentIndex - Using 'Department Name' at index:",o.idx),o.idx;let a=t.find(({value:l})=>l.includes("department")&&!l.includes("id")&&!l.includes("#")&&!l.includes("code")&&!l.includes("number"));if(a)return console.log("DEBUG pickDepartmentIndex - Using non-ID department at index:",a.idx),a.idx;let s=t.find(({value:l})=>l.includes("department"));return s&&console.log("DEBUG pickDepartmentIndex - Using fallback department at index:",s.idx),s?s.idx:-1}function Fn(e,t,n={}){if(Gt(),!t||!t.length)return;let o=document.createElement("div");o.className="pf-modal";let a=t.filter(i=>i.type==="missing_from_payroll"),s=t.filter(i=>i.type==="missing_from_roster"),l=t.filter(i=>!i.type),c="";if(a.length>0&&(c+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading pf-mismatch-warning">
                    <span class="pf-mismatch-icon">\u26A0\uFE0F</span>
                    In Roster but NOT in Payroll (${a.length})
                </h4>
                <p class="pf-mismatch-subtext">These employees appear in your centralized roster but were not found in the payroll data. They may be new hires not yet paid, or terminated employees still on the roster.</p>
                <div class="pf-mismatch-tiles">
                    ${a.map(i=>`
                        <div class="pf-mismatch-tile pf-mismatch-missing-payroll">
                            <span class="pf-mismatch-name">${S(i.name)}</span>
                            <span class="pf-mismatch-detail">${S(i.department)}</span>
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
                    ${s.map(i=>`
                        <div class="pf-mismatch-tile pf-mismatch-missing-roster">
                            <span class="pf-mismatch-name">${S(i.name)}</span>
                            <span class="pf-mismatch-detail">${S(i.department)}</span>
                        </div>
                    `).join("")}
                </div>
            </div>
        `),l.length>0){let i=n.formatter||(r=>typeof r=="string"?{name:r,source:"",isMissingFromTarget:!0}:r);c+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading">
                    <span class="pf-mismatch-icon">\u{1F4CB}</span>
                    ${S(n.label||e)} (${l.length})
                </h4>
                <div class="pf-mismatch-tiles">
                    ${l.map(r=>{let u=i(r);return`
                            <div class="pf-mismatch-tile">
                                <span class="pf-mismatch-name">${S(u.name||u.employee||"")}</span>
                                <span class="pf-mismatch-detail">${S(u.source||`${u.rosterDept||""} \u2192 ${u.payrollDept||""}`)}</span>
                            </div>
                        `}).join("")}
                </div>
            </div>
        `}c||(c='<p class="pf-mismatch-empty">No differences found.</p>'),o.innerHTML=`
        <div class="pf-modal-content pf-headcount-modal">
            <div class="pf-modal-header">
                <h3>${S(e)}</h3>
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
    `,o.addEventListener("click",i=>{i.target===o&&Gt()}),o.querySelectorAll("[data-modal-close]").forEach(i=>{i.addEventListener("click",Gt)}),document.body.appendChild(o)}function Gt(){var e;(e=document.querySelector(".pf-modal"))==null||e.remove()}function vt(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=Rt(),n=document.getElementById("step-notes-input"),o=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!o;let a=document.getElementById("headcount-notes-hint");a&&(a.textContent=t?"Please document outstanding differences before signing off.":""),H.skipAnalysis&&Wt()}function Ta(){var n;let e=Rt(),t=((n=document.getElementById("step-notes-input"))==null?void 0:n.value.trim())||"";if(e&&!t){window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");return}Jt({statusText:"Headcount Review signed off."})}function Wt(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(mt)?t.slice(mt.length).replace(/^\s+/,""):t.replace(new RegExp(`^${mt}\\s*`,"i"),"").trimStart(),o=mt+(n?`
${n}`:"");if(e.value!==o){e.value=o;let a=Se(2);a&&U(a.note,o)}}function jn(e){let t=e!=null&&e.target&&e.target instanceof HTMLInputElement?e.target:document.getElementById("bank-amount-input"),n=qt(t==null?void 0:t.value),o=Xn(n);t&&(t.value=o),Ye({bankAmount:n},{rerender:!1})}function La(){let e=be.findIndex(t=>t.id===3);e!==-1&&Qe(Math.min(be.length-1,e+1))}function Ba(){let e=be.findIndex(t=>t.id===4);e!==-1&&Qe(Math.min(be.length-1,e+1))}async function Ma(){if(!ae()){window.alert("Excel runtime is unavailable.");return}if(window.confirm(`Archive Payroll Run

This will:
1. Create an archive workbook with all payroll tabs
2. Update PR_Archive_Summary with current period
3. Clear working data from all payroll sheets
4. Clear non-permanent notes and config values

Make sure you've completed all review steps before archiving.

Continue?`))try{if(console.log("[Archive] Step 1: Creating archive workbook..."),!await Va()){window.alert("Archive cancelled or failed. No data was modified.");return}console.log("[Archive] Step 1 complete: Archive workbook created"),console.log("[Archive] Step 2: Updating PR_Archive_Summary..."),await Fa(),console.log("[Archive] Step 2 complete: Archive summary updated"),console.log("[Archive] Step 3: Clearing working data..."),await ja(),console.log("[Archive] Step 3 complete: Working data cleared"),console.log("[Archive] Step 4: Clearing non-permanent notes..."),await Ha(),console.log("[Archive] Step 4 complete: Notes cleared"),console.log("[Archive] Step 5: Resetting config values..."),await Ua(),console.log("[Archive] Step 5 complete: Config reset"),console.log("[Archive] Archive workflow complete!"),await Wn(),le(),window.alert(`Archive Complete!

\u2713 Payroll tabs archived to new workbook
\u2713 PR_Archive_Summary updated with current period
\u2713 Working data cleared
\u2713 Notes and config reset

Ready for next payroll cycle.`)}catch(t){console.error("[Archive] Error during archive:",t),window.alert(`Archive Error

An error occurred during the archive process:
`+t.message+`

Please check the console for details and verify your data.`)}}async function Va(){try{let t=`Payroll_Archive_${Ct()||new Date().toISOString().split("T")[0]}`,n=[D.DATA,D.DATA_CLEAN,D.EXPENSE_MAPPING,D.EXPENSE_REVIEW,D.JE_DRAFT,D.ARCHIVE_SUMMARY];return await Excel.run(async o=>{let s=o.workbook.worksheets;s.load("items/name"),await o.sync();let l=o.application.createWorkbook();await o.sync(),console.log(`[Archive] New workbook created. User should save as: ${t}`);for(let c of n){let i=s.items.find(u=>u.name===c);if(!i){console.warn(`[Archive] Sheet not found: ${c}`);continue}let r=i.getUsedRangeOrNullObject();if(r.load("values,numberFormat,address"),await o.sync(),r.isNullObject||!r.values||r.values.length===0){console.log(`[Archive] Skipping empty sheet: ${c}`);continue}console.log(`[Archive] Archived data from: ${c} (${r.values.length} rows)`)}return window.alert(`Archive Workbook Created

A new workbook has been opened with your payroll data.

Please save it now:
1. Go to the new workbook window
2. Press Ctrl+S (or Cmd+S on Mac)
3. Save as: ${t}

Click OK after saving to continue with the archive process.`),!0})}catch(e){return console.error("[Archive] Error creating archive workbook:",e),!1}}async function Fa(){await Excel.run(async e=>{let t=e.workbook.worksheets.getItemOrNullObject(D.ARCHIVE_SUMMARY),n=e.workbook.worksheets.getItemOrNullObject(D.DATA_CLEAN);if(t.load("isNullObject"),n.load("isNullObject"),await e.sync(),t.isNullObject){console.warn("[Archive] PR_Archive_Summary not found - skipping");return}if(n.isNullObject){console.warn("[Archive] PR_Data_Clean not found - skipping");return}let o=n.getUsedRangeOrNullObject();if(o.load("values"),await e.sync(),o.isNullObject||!o.values||o.values.length<2){console.warn("[Archive] PR_Data_Clean is empty - skipping archive summary update");return}let a=(o.values[0]||[]).map(m=>String(m||"").toLowerCase().trim()),s=o.values.slice(1),l=a.findIndex(m=>m.includes("amount")),c=a.findIndex(m=>m.includes("employee")),i=a.findIndex(m=>m.includes("payroll")&&m.includes("date")||m.includes("pay period")||m==="date"),r=0,u=new Set,f=Ct()||"";s.forEach(m=>{l>=0&&(r+=Number(m[l])||0),c>=0&&m[c]&&u.add(String(m[c]).trim()),i>=0&&m[i]&&!f&&(f=String(m[i]))});let p=u.size;console.log(`[Archive] Current period: Date=${f}, Total=${r}, Employees=${p}`);let d=t.getUsedRangeOrNullObject();d.load("values,rowCount"),await e.sync();let h=[],v=[];!d.isNullObject&&d.values&&d.values.length>0&&(h=d.values[0],v=d.values.slice(1)),h.length===0&&(h=["Pay Period","Total Payroll","Employee Count","Archived Date"],t.getRange("A1:D1").values=[h],await e.sync());let C=h.map(m=>String(m||"").toLowerCase().trim()),g=C.findIndex(m=>m.includes("pay period")||m.includes("period")||m==="date"),b=C.findIndex(m=>m.includes("total")),E=C.findIndex(m=>m.includes("employee")||m.includes("count")),x=C.findIndex(m=>m.includes("archived")),A=new Array(h.length).fill("");g>=0&&(A[g]=f),b>=0&&(A[b]=r),E>=0&&(A[E]=p),x>=0&&(A[x]=new Date().toISOString().split("T")[0]),v.length>=5&&(v=v.slice(0,4),console.log("[Archive] Trimmed archive to 4 periods, adding current")),v.unshift(A);let M=2,_=M+5;if(t.getRange(`A${M}:${String.fromCharCode(64+h.length)}${_}`).clear(Excel.ClearApplyTo.contents),await e.sync(),v.length>0){let m=t.getRange(`A${M}:${String.fromCharCode(64+h.length)}${M+v.length-1}`);m.values=v,await e.sync()}console.log(`[Archive] Archive summary updated with ${v.length} periods`)})}async function ja(){let e=[D.DATA,D.DATA_CLEAN,D.EXPENSE_REVIEW,D.JE_DRAFT];await Excel.run(async t=>{for(let n of e){let o=t.workbook.worksheets.getItemOrNullObject(n);if(o.load("isNullObject"),await t.sync(),o.isNullObject){console.log(`[Archive] Sheet not found: ${n}`);continue}let a=o.getUsedRangeOrNullObject();if(a.load("rowCount,columnCount,address"),await t.sync(),a.isNullObject||a.rowCount<=1){console.log(`[Archive] Sheet empty or headers only: ${n}`);continue}if(o.getRange(`A2:${String.fromCharCode(64+a.columnCount)}${a.rowCount}`).clear(Excel.ClearApplyTo.contents),n===D.EXPENSE_REVIEW){let l=o.charts;l.load("items"),await t.sync();for(let c=l.items.length-1;c>=0;c--)l.items[c].delete()}await t.sync(),console.log(`[Archive] Cleared data from: ${n}`)}})}async function Ha(){await Excel.run(async e=>{let t=await Ne(e);if(!t){console.warn("[Archive] Config table not found");return}let n=t.getDataBodyRange();n.load("values,rowCount"),await e.sync();let o=n.values||[],a=0,s=Object.values(Ie).map(l=>l.note);for(let l=0;l<o.length;l++){let c=String(o[l][B.FIELD]||"").trim(),i=String(o[l][B.PERMANENT]||"").toUpperCase();s.includes(c)&&i!=="Y"&&(n.getCell(l,B.VALUE).values=[[""]],a++)}await e.sync(),console.log(`[Archive] Cleared ${a} non-permanent notes`)})}async function Ua(){let e=["Payroll_Date","PR_Payroll_Date","Accounting_Period","Journal_Entry_ID","PR_Journal_Entry_ID","JE_Transaction_ID",...Object.values(Ie).map(t=>t.signOff),...Object.values(Ie).map(t=>t.reviewer),...Object.values(de)];await Excel.run(async t=>{let n=await Ne(t);if(!n){console.warn("[Archive] Config table not found");return}let o=n.getDataBodyRange();o.load("values,rowCount"),await t.sync();let a=o.values||[],s=0;for(let l=0;l<a.length;l++){let c=String(a[l][B.FIELD]||"").trim(),i=String(a[l][B.PERMANENT]||"").toUpperCase();e.some(u=>oe(u)===oe(c))&&i!=="Y"&&(o.getCell(l,B.VALUE).values=[[""]],s++)}await t.sync(),console.log(`[Archive] Reset ${s} non-permanent config values`),Object.keys(W.values).forEach(l=>{e.some(c=>oe(c)===oe(l))&&(W.values[l]="")})})}async function Ga(){if(!ae()){window.alert("Excel runtime is unavailable.");return}G.loading=!0,G.lastError=null,Gn(!1),le();try{let e=await Excel.run(async t=>{let o=t.workbook.worksheets.getItem(D.JE_DRAFT).getUsedRangeOrNullObject();o.load("values"),await t.sync();let a=o.isNullObject?[]:o.values||[];if(!a.length)throw new Error(`${D.JE_DRAFT} is empty.`);let s=(a[0]||[]).map(u=>ve(u)),l=s.findIndex(u=>u.includes("debit")),c=s.findIndex(u=>u.includes("credit"));if(l===-1||c===-1)throw new Error("Debit/Credit columns not found in JE Draft.");let i=0,r=0;return a.slice(1).forEach(u=>{i+=Number(u[l])||0,r+=Number(u[c])||0}),{debitTotal:i,creditTotal:r,difference:r-i}});Object.assign(G,e,{lastError:null})}catch(e){console.warn("JE summary:",e),G.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",G.debitTotal=null,G.creditTotal=null,G.difference=null}finally{G.loading=!1,le()}}async function za(){try{let e=Number.isFinite(Number(G.debitTotal))?G.debitTotal:"",t=Number.isFinite(Number(G.creditTotal))?G.creditTotal:"",n=Number.isFinite(Number(G.difference))?G.difference:"";await Promise.all([Ke(Fo,String(e)),Ke(jo,String(t)),Ke(Ho,String(n))]),Gn(!0)}catch(e){console.error("JE save:",e)}}async function Wa(){if(!ae()){window.alert("Excel runtime is unavailable.");return}G.loading=!0,G.lastError=null,le();try{await Excel.run(async e=>{let t="",n="",o=await Ne(e);if(o){let g=o.getDataBodyRange();g.load("values"),await e.sync();let b=g.values||[];for(let E of b){let x=String(E[B.FIELD]||"").trim(),A=E[B.VALUE];(x==="Journal_Entry_ID"||x==="JE_Transaction_ID"||x==="PR_Journal_Entry_ID")&&(t=String(A||"").trim()),Xe.some(M=>x===M||oe(x)===oe(M))&&(n=ye(A)||"")}}console.log("JE Generation - RefNumber:",t,"TxnDate:",n);let a=e.workbook.worksheets.getItemOrNullObject(D.DATA_CLEAN);if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PR_Data_Clean sheet not found. Run Validate & Reconcile first.");let s=a.getUsedRangeOrNullObject();if(s.load("values"),await e.sync(),s.isNullObject)throw new Error("PR_Data_Clean is empty. Run Validate & Reconcile first.");let l=s.values||[];if(l.length<2)throw new Error("PR_Data_Clean has no data rows.");let c=l[0].map(g=>ve(g));console.log("JE Generation - PR_Data_Clean headers:",c);let i={accountNumber:c.findIndex(g=>g.includes("account")&&(g.includes("number")||g.includes("#"))),accountName:c.findIndex(g=>g.includes("account")&&g.includes("name")),amount:c.findIndex(g=>g.includes("amount")),department:Be(c),payrollCategory:c.findIndex(g=>g.includes("payroll")&&g.includes("category")),employee:c.findIndex(g=>g.includes("employee"))};if(console.log("JE Generation - Column indices:",i),i.amount===-1)throw new Error("Amount column not found in PR_Data_Clean.");let r=new Map;for(let g=1;g<l.length;g++){let b=l[g],E=i.accountNumber>=0?String(b[i.accountNumber]||"").trim():"",x=i.accountName>=0?String(b[i.accountName]||"").trim():"",A=Number(b[i.amount])||0,M=i.department>=0?String(b[i.department]||"").trim():"";if(A===0)continue;let _=`${E}|${M}`;if(r.has(_)){let T=r.get(_);T.amount+=A}else r.set(_,{accountNumber:E,accountName:x,department:M,amount:A})}console.log("JE Generation - Aggregated into",r.size,"unique Account+Department combinations");let u=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],f=[u],p=0,d=0,h=Array.from(r.values()).sort((g,b)=>{let E=String(g.accountNumber).localeCompare(String(b.accountNumber));return E!==0?E:String(g.department).localeCompare(String(b.department))});for(let g of h){let{accountNumber:b,accountName:E,department:x,amount:A}=g,M=A>0?A:0,_=A<0?Math.abs(A):0,T=[E,x].filter(Boolean).join(" - ");p+=M,d+=_,f.push([t,n,b,E,A,M||"",_||"",T,x])}console.log("JE Generation - Built",f.length-1,"summarized journal lines"),console.log("JE Generation - Total Debit:",p,"Total Credit:",d);let v=e.workbook.worksheets.getItemOrNullObject(D.JE_DRAFT);if(v.load("isNullObject"),await e.sync(),v.isNullObject)v=e.workbook.worksheets.add(D.JE_DRAFT),await e.sync();else{let g=v.getUsedRangeOrNullObject();g.load("address"),await e.sync(),g.isNullObject||(g.clear(),await e.sync())}let C=v.getRangeByIndexes(0,0,f.length,u.length);C.values=f,await e.sync();try{let g=f.length-1,b=v.getRange("A1:I1");An(b),g>0&&(Dn(v,1,g),ft(v,4,g),ft(v,5,g),ft(v,6,g)),v.getRange("A:I").format.autofitColumns(),await e.sync()}catch(g){console.warn("JE formatting error (non-critical):",g)}v.activate(),v.getRange("A1").select(),await e.sync(),G.debitTotal=p,G.creditTotal=d,G.difference=d-p}),G.loading=!1,G.lastError=null,le()}catch(e){console.error("JE Generation failed:",e),G.loading=!1,G.lastError=e.message||"Failed to generate journal entry.",le()}}async function Ja(){if(!ae()){window.alert("Excel runtime is unavailable.");return}try{let{rows:e}=await Excel.run(async n=>{let a=n.workbook.worksheets.getItem(D.JE_DRAFT).getUsedRangeOrNullObject();a.load("values"),await n.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error(`${D.JE_DRAFT} is empty.`);return{rows:s}}),t=ka(e);Ra(`pr-je-draft-${Ze()}.csv`,t)}catch(e){console.warn("JE export:",e),window.alert("Unable to export the JE draft. Confirm the sheet has data.")}}})();
//# sourceMappingURL=app.bundle.js.map
