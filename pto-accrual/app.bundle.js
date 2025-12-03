/* Prairie Forge PTO Accrual */
(()=>{function K(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}var qe="SS_PF_Config";async function kt(e,t=[qe]){var o;let n=e.workbook.tables;n.load("items/name"),await e.sync();let a=(o=n.items)==null?void 0:o.find(s=>t.includes(s.name));return a?e.workbook.tables.getItem(a.name):(console.warn("Config table not found. Looking for:",t),null)}function Ot(e){let t=e.map(n=>String(n||"").trim().toLowerCase());return{field:t.findIndex(n=>n==="field"||n==="field name"||n==="setting"),value:t.findIndex(n=>n==="value"||n==="setting value"),type:t.findIndex(n=>n==="type"||n==="category"),title:t.findIndex(n=>n==="title"||n==="display name"),permanent:t.findIndex(n=>n==="permanent"||n==="persist")}}async function xt(e=[qe]){if(!K())return{};try{return await Excel.run(async t=>{let n=await kt(t,e);if(!n)return{};let a=n.getDataBodyRange(),o=n.getHeaderRowRange();a.load("values"),o.load("values"),await t.sync();let s=o.values[0]||[],i=Ot(s);if(i.field===-1||i.value===-1)return console.warn("Config table missing FIELD or VALUE columns. Headers:",s),{};let d={};return(a.values||[]).forEach(l=>{var f;let p=String(l[i.field]||"").trim();p&&(d[p]=(f=l[i.value])!=null?f:"")}),console.log("Configuration loaded:",Object.keys(d).length,"fields"),d})}catch(t){return console.error("Failed to load configuration:",t),{}}}async function _e(e,t,n=[qe]){if(!K())return!1;try{return await Excel.run(async a=>{let o=await kt(a,n);if(!o){console.warn("Config table not found for write");return}let s=o.getDataBodyRange(),i=o.getHeaderRowRange();s.load("values"),i.load("values"),await a.sync();let d=i.values[0]||[],r=Ot(d);if(r.field===-1||r.value===-1){console.error("Config table missing FIELD or VALUE columns");return}let p=(s.values||[]).findIndex(f=>String(f[r.field]||"").trim()===e);if(p>=0)s.getCell(p,r.value).values=[[t]];else{let f=new Array(d.length).fill("");r.type>=0&&(f[r.type]="Run Settings"),f[r.field]=e,f[r.value]=t,r.permanent>=0&&(f[r.permanent]="N"),r.title>=0&&(f[r.title]=""),o.rows.add(null,[f]),console.log("Added new config row:",e,"=",t)}await a.sync(),console.log("Saved config:",e,"=",t)}),!0}catch(a){return console.error("Failed to save config:",e,a),!1}}var On="SS_PF_Config",xn="module-prefix",Ye="system",ke={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function St(){if(!K())return{...ke};try{return await Excel.run(async e=>{var p,f;let t=e.workbook.worksheets.getItemOrNullObject(On);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...ke};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((p=n.values)!=null&&p.length))return{...ke};let a=n.values,o=Cn(a[0]),s=o.get("category"),i=o.get("field"),d=o.get("value");if(s===void 0||i===void 0||d===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...ke};let r={},l=!1;for(let u=1;u<a.length;u++){let c=a[u];if(je(c[s])===xn){let y=String((f=c[i])!=null?f:"").trim().toUpperCase(),b=je(c[d]);y&&b&&(r[y]=b,l=!0)}}return l?(console.log("[Tab Visibility] Loaded prefix config:",r),r):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...ke})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...ke}}}async function We(e){if(!K())return;let t=je(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await St();await Excel.run(async a=>{let o=a.workbook.worksheets;o.load("items/name,visibility"),await a.sync();let s={};for(let[u,c]of Object.entries(n))s[c]||(s[c]=[]),s[c].push(u);let i=s[t]||[],d=s[Ye]||[],r=[];for(let[u,c]of Object.entries(s))u!==t&&u!==Ye&&r.push(...c);console.log(`[Tab Visibility] Active prefixes: ${i.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${r.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${d.join(", ")}`);let l=[],p=[];o.items.forEach(u=>{let c=u.name,g=c.toUpperCase(),y=i.some(S=>g.startsWith(S)),b=r.some(S=>g.startsWith(S)),m=d.some(S=>g.startsWith(S));y?(l.push(u),console.log(`[Tab Visibility] SHOW: ${c} (matches active module prefix)`)):m?(p.push(u),console.log(`[Tab Visibility] HIDE: ${c} (system sheet)`)):b?(p.push(u),console.log(`[Tab Visibility] HIDE: ${c} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${c} (no prefix match, leaving as-is)`)});for(let u of l)u.visibility=Excel.SheetVisibility.visible;if(await a.sync(),o.items.filter(u=>u.visibility===Excel.SheetVisibility.visible).length>p.length){for(let u of p)try{u.visibility=Excel.SheetVisibility.hidden}catch(c){console.warn(`[Tab Visibility] Could not hide "${u.name}":`,c.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${l.length}, hid ${p.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function Sn(){if(!K()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.visibility!==Excel.SheetVisibility.visible&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${a.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function En(){if(!K()){console.log("Excel not available");return}try{let e=await St(),t=[];for(let[n,a]of Object.entries(e))a===Ye&&t.push(n);await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name,visibility"),await n.sync(),a.items.forEach(o=>{let s=o.name.toUpperCase();t.some(i=>s.startsWith(i))&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function Cn(e=[]){let t=new Map;return e.forEach((n,a)=>{let o=je(n);o&&t.set(o,a)}),t}function je(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=Sn,window.PrairieForge.unhideSystemSheets=En,window.PrairieForge.applyModuleTabVisibility=We);var Et={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var Tt=Et.ADA_IMAGE_URL;async function Le(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async a=>{let o=a.workbook.worksheets.getItemOrNullObject(e);o.load("isNullObject, name, visibility"),await a.sync();let s;o.isNullObject?(s=a.workbook.worksheets.add(e),await a.sync(),await Ct(a,s,t,n)):(s=o,s.visibility!==Excel.SheetVisibility.visible&&(s.visibility=Excel.SheetVisibility.visible,await a.sync()),await Ct(a,s,t,n)),s.activate(),s.getRange("A1").select(),await a.sync()})}catch(a){console.error(`Error activating homepage sheet ${e}:`,a)}}async function Ct(e,t,n,a){try{let l=t.getUsedRangeOrNullObject();l.load("isNullObject"),await e.sync(),l.isNullObject||(l.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let o=[[n,""],[a,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=o;let i=t.getRange("A1:Z100");i.format.fill.color="#0f0f0f";let d=t.getRange("A1");d.format.font.bold=!0,d.format.font.size=36,d.format.font.color="#ffffff",d.format.font.name="Segoe UI Light",d.format.verticalAlignment="Center";let r=t.getRange("A2");r.format.font.size=14,r.format.font.color="#a0a0a0",r.format.font.name="Segoe UI",r.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var _t={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function Be(e){return _t[e]||_t["module-selector"]}function Pt(){Qe();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${Tt}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",_n),e}function Qe(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function _n(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${Tt}" alt="Ada" />
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",Ke),t.addEventListener("click",o=>{o.target===t&&Ke()});let a=o=>{o.key==="Escape"&&(Ke(),document.removeEventListener("keydown",a))};document.addEventListener("keydown",a)}function Ke(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var Tn=["January","February","March","April","May","June","July","August","September","October","November","December"],Rt=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],Pn=["Su","Mo","Tu","We","Th","Fr","Sa"],Q=null,ae=null;function At(e,t={}){let n=document.getElementById(e);if(!n)return;let{onChange:a=null,minDate:o=null,maxDate:s=null,readonly:i=!1}=t,d=n.closest(".pf-datepicker-wrapper");d||(d=document.createElement("div"),d.className="pf-datepicker-wrapper",n.parentNode.insertBefore(d,n),d.appendChild(n)),n.type="text",n.placeholder="Select date...",n.classList.add("pf-datepicker-input"),n.readOnly=!0;let r=n.value?It(n.value):null;r&&(n.value=Ze(r),n.dataset.value=Te(r));let l=d.querySelector(".pf-datepicker-icon");l||(l=document.createElement("span"),l.className="pf-datepicker-icon",l.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>',d.appendChild(l));let p={inputId:e,input:n,selectedDate:r,viewDate:r?new Date(r):new Date,onChange:a,minDate:o,maxDate:s};function f(){i||(ae=p,In())}return n.addEventListener("click",f),l.addEventListener("click",f),{getValue:()=>p.selectedDate?Te(p.selectedDate):"",setValue:u=>{let c=It(u);p.selectedDate=c,p.viewDate=c?new Date(c):new Date,c?(n.value=Ze(c),n.dataset.value=Te(c)):(n.value="",n.dataset.value="")},open:f,close:Me}}function In(){ae&&(Q||(Q=document.createElement("div"),Q.className="pf-datepicker-modal",Q.id="pf-datepicker-modal",document.body.appendChild(Q)),Dt(),requestAnimationFrame(()=>{Q.classList.add("is-open")}),document.addEventListener("keydown",Nt))}function Me(){Q&&Q.classList.remove("is-open"),document.removeEventListener("keydown",Nt),ae=null}function Nt(e){e.key==="Escape"&&Me()}function Dt(){if(!Q||!ae)return;let{viewDate:e,selectedDate:t,minDate:n,maxDate:a}=ae,o=e.getFullYear(),s=e.getMonth();Q.innerHTML=`
        <div class="pf-datepicker-backdrop"></div>
        <div class="pf-datepicker-container">
            <div class="pf-datepicker-header">
                <div class="pf-datepicker-nav-group">
                    <button type="button" class="pf-datepicker-nav" data-action="prev-year" title="Previous Year">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="11 17 6 12 11 7"/><polyline points="18 17 13 12 18 7"/></svg>
                    </button>
                    <button type="button" class="pf-datepicker-nav" data-action="prev-month" title="Previous Month">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>
                    </button>
                </div>
                <span class="pf-datepicker-title">${Tn[s]} ${o}</span>
                <div class="pf-datepicker-nav-group">
                    <button type="button" class="pf-datepicker-nav" data-action="next-month" title="Next Month">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>
                    </button>
                    <button type="button" class="pf-datepicker-nav" data-action="next-year" title="Next Year">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="13 17 18 12 13 7"/><polyline points="6 17 11 12 6 7"/></svg>
                    </button>
                </div>
            </div>
            <div class="pf-datepicker-weekdays">
                ${Pn.map(i=>`<span>${i}</span>`).join("")}
            </div>
            <div class="pf-datepicker-days">
                ${Nn(o,s,t,n,a)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-btn pf-datepicker-today" data-action="today">Today</button>
                <button type="button" class="pf-datepicker-btn pf-datepicker-clear" data-action="clear">Clear</button>
            </div>
        </div>
    `,Rn()}function Rn(){var e;Q&&((e=Q.querySelector(".pf-datepicker-backdrop"))==null||e.addEventListener("click",Me),Q.querySelectorAll(".pf-datepicker-nav").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=t.dataset.action;An(a)})}),Q.querySelectorAll(".pf-datepicker-day:not(.disabled)").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=parseInt(t.dataset.day),o=parseInt(t.dataset.month),s=parseInt(t.dataset.year);Xe(new Date(s,o,a))})}),Q.querySelectorAll(".pf-datepicker-btn").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=t.dataset.action;a==="today"?Xe(new Date):a==="clear"&&Xe(null)})}))}function An(e){if(!ae)return;let t=ae.viewDate;switch(e){case"prev-year":t.setFullYear(t.getFullYear()-1);break;case"prev-month":t.setMonth(t.getMonth()-1);break;case"next-month":t.setMonth(t.getMonth()+1);break;case"next-year":t.setFullYear(t.getFullYear()+1);break}Dt()}function Xe(e){if(!ae)return;let{input:t,onChange:n}=ae;ae.selectedDate=e,e?(t.value=Ze(e),t.dataset.value=Te(e),ae.viewDate=new Date(e)):(t.value="",t.dataset.value=""),n&&n(e?Te(e):""),t.dispatchEvent(new Event("change",{bubbles:!0})),Me()}function Nn(e,t,n,a,o){let s=new Date(e,t,1).getDay(),i=new Date(e,t+1,0).getDate(),d=new Date(e,t,0).getDate(),r=new Date;r.setHours(0,0,0,0),n&&(n=new Date(n),n.setHours(0,0,0,0));let l="";for(let c=s-1;c>=0;c--){let g=d-c,y=t===0?11:t-1,b=t===0?e-1:e;l+=`<button type="button" class="pf-datepicker-day other-month" data-day="${g}" data-month="${y}" data-year="${b}">${g}</button>`}for(let c=1;c<=i;c++){let g=new Date(e,t,c);g.setHours(0,0,0,0);let y=g.getTime()===r.getTime(),b=n&&g.getTime()===n.getTime(),m="pf-datepicker-day";y&&(m+=" today"),b&&(m+=" selected");let S=!1;a&&g<a&&(S=!0),o&&g>o&&(S=!0),S&&(m+=" disabled"),l+=`<button type="button" class="${m}" data-day="${c}" data-month="${t}" data-year="${e}" ${S?"disabled":""}>${c}</button>`}let p=42,f=s+i,u=p-f;for(let c=1;c<=u;c++){let g=t===11?0:t+1,y=t===11?e+1:e;l+=`<button type="button" class="pf-datepicker-day other-month" data-day="${c}" data-month="${g}" data-year="${y}">${c}</button>`}return l}function It(e){if(!e)return null;if(/^\d{4}-\d{2}-\d{2}$/.test(e)){let[a,o,s]=e.split("-").map(Number);return new Date(a,o-1,s)}let t=e.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);if(t){let a=Rt.findIndex(o=>o.toLowerCase()===t[1].toLowerCase().substring(0,3));if(a>=0)return new Date(parseInt(t[3]),a,parseInt(t[2]))}if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(e)){let[a,o,s]=e.split("/").map(Number);return new Date(s,a-1,o)}let n=new Date(e);return isNaN(n.getTime())?null:n}function Ze(e){return e?`${Rt[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`:""}function Te(e){if(!e)return"";let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}var $t=`
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
`.trim(),jt=`
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
`.trim(),Lt=`
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
`.trim(),Ve=`
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
`.trim(),Ja=`
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
`.trim(),za=`
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
`.trim(),Dn={config:`
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
    `};function Bt(e){return e&&Dn[e]||""}var et=`
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
`.trim(),tt=`
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
`.trim(),Pe=`
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
`.trim(),He=`
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
`.trim(),qa=`
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
`.trim(),nt=`
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
`.trim(),Mt=`
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
`.trim(),Ya=`
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
        <path d="M17 8l-5-5-5 5" />
        <path d="M12 3v12" />
    </svg>
`.trim(),Vt=`
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
`.trim(),Ht=`
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
`.trim(),Ft=`
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
`.trim(),Ut=`
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
`.trim(),Wa=`
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
`.trim(),Ka=`
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
`.trim(),Qa=`
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
`.trim(),Xa=`
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
`.trim(),Za=`
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
`.trim(),eo=`
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
`.trim(),to=`
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
`.trim(),no=`
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
`.trim(),Ie=`
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
`.trim(),Gt=`
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
`.trim();function Re(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function G(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function pe({textareaId:e,value:t,permanentId:n,isPermanent:a,hintId:o,saveButtonId:s,isSaved:i=!1,placeholder:d="Enter notes here..."}){let r=a?tt:et,l=s?`<button type="button" class="pf-action-toggle pf-save-btn ${i?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Ft}</button>`:"",p=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${a?"is-locked":""}" id="${n}" aria-pressed="${a}" title="Lock notes (retain after archive)">${r}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Re(d)}">${Re(t||"")}</textarea>
                ${o?`<p class="pf-signoff-hint" id="${o}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?G(p,"Lock"):""}
                ${s?G(l,"Save"):""}
            </div>
        </article>
    `}function fe({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:a,isComplete:o,saveButtonId:s,isSaved:i=!1,completeButtonId:d,subtext:r="Sign-off below. Click checkmark icon. Done."}){let l=`<button type="button" class="pf-action-toggle ${o?"is-active":""}" id="${d}" aria-pressed="${!!o}" title="Mark step complete">${Pe}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${Re(r)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${Re(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${Re(a)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${G(l,"Done")}
            </div>
        </article>
    `}function at(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?tt:et)}function ce(e,t){e&&e.classList.toggle("is-saved",t)}function ot(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(a=>{let o=a.getAttribute("data-save-input"),s=document.getElementById(o);if(!s)return;let i=()=>{ce(a,!1)};s.addEventListener("input",i),n.push(()=>s.removeEventListener("input",i))}),()=>n.forEach(a=>a())}function Jt(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function zt(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var st={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},Fe={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function it(e){e.format.fill.color=st.fillColor,e.format.font.color=st.fontColor,e.format.font.bold=st.bold}function ge(e,t,n,a=!1){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a?Fe.currencyWithNegative:Fe.currency]]}function Oe(e,t,n){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[Fe.number]]}function qt(e,t,n,a=Fe.date){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a]]}var $n="d74b68e",Ne="pto-accrual";var he="PTO Accrual";function Z(e,t="info",n=4e3){document.querySelectorAll(".pf-toast").forEach(o=>o.remove());let a=document.createElement("div");if(a.className=`pf-toast pf-toast--${t}`,a.innerHTML=`
        <div class="pf-toast-content">
            <span class="pf-toast-icon">${t==="success"?"\u2705":t==="error"?"\u274C":"\u2139\uFE0F"}</span>
            <span class="pf-toast-message">${e.replace(/\n/g,"<br>")}</span>
        </div>
        <button class="pf-toast-close" onclick="this.parentElement.remove()">\xD7</button>
    `,!document.getElementById("pf-toast-styles")){let o=document.createElement("style");o.id="pf-toast-styles",o.textContent=`
            .pf-toast {
                position: fixed;
                top: 20px;
                left: 50%;
                transform: translateX(-50%);
                background: #1a1a2e;
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
            .pf-toast-icon { font-size: 18px; }
            .pf-toast-message { font-size: 14px; line-height: 1.4; }
            .pf-toast-close { background: none; border: none; color: #888; font-size: 20px; cursor: pointer; padding: 0; margin-left: 8px; }
            .pf-toast-close:hover { color: white; }
            @keyframes toastIn { from { opacity: 0; transform: translateX(-50%) translateY(-20px); } }
        `,document.head.appendChild(o)}return document.body.appendChild(a),n>0&&setTimeout(()=>a.remove(),n),a}function jn(e,t={}){let{title:n="Confirm Action",confirmText:a="Continue",cancelText:o="Cancel",icon:s="\u{1F4CB}",destructive:i=!1}=t;return new Promise(d=>{document.querySelectorAll(".pf-confirm-overlay").forEach(l=>l.remove());let r=document.createElement("div");if(r.className="pf-confirm-overlay",r.innerHTML=`
            <div class="pf-confirm-dialog">
                <div class="pf-confirm-icon">${s}</div>
                <div class="pf-confirm-title">${n}</div>
                <div class="pf-confirm-message">${e.replace(/\n/g,"<br>")}</div>
                <div class="pf-confirm-buttons">
                    <button class="pf-confirm-btn pf-confirm-btn--cancel">${o}</button>
                    <button class="pf-confirm-btn pf-confirm-btn--ok ${i?"pf-confirm-btn--destructive":""}">${a}</button>
                </div>
            </div>
        `,!document.getElementById("pf-confirm-styles")){let l=document.createElement("style");l.id="pf-confirm-styles",l.textContent=`
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
                    background: linear-gradient(145deg, rgba(30, 30, 50, 0.95), rgba(20, 20, 35, 0.98));
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
            `,document.head.appendChild(l)}document.body.appendChild(r),r.addEventListener("click",l=>{l.target===r&&(r.remove(),d(!1))}),r.querySelector(".pf-confirm-btn--cancel").onclick=()=>{r.remove(),d(!1)},r.querySelector(".pf-confirm-btn--ok").onclick=()=>{r.remove(),d(!0)}})}var Ln="Calculate your PTO liability, compare against last period, and generate a balanced journal entry\u2014all without leaving Excel.",Bn="../module-selector/index.html",Mn="pf-loader-overlay",me=["SS_PF_Config"],w={payrollProvider:"PTO_Payroll_Provider",payrollDate:"PTO_Analysis_Date",accountingPeriod:"PTO_Accounting_Period",journalEntryId:"PTO_Journal_Entry_ID",companyName:"SS_Company_Name",accountingSoftware:"SS_Accounting_Software",reviewerName:"PTO_Reviewer",validationDataBalance:"PTO_Validation_Data_Balance",validationCleanBalance:"PTO_Validation_Clean_Balance",validationDifference:"PTO_Validation_Difference",headcountRosterCount:"PTO_Headcount_Roster_Count",headcountPayrollCount:"PTO_Headcount_Payroll_Count",headcountDifference:"PTO_Headcount_Difference",journalDebitTotal:"PTO_JE_Debit_Total",journalCreditTotal:"PTO_JE_Credit_Total",journalDifference:"PTO_JE_Difference"},ye="User opted to skip the headcount review this period.",Ge={0:{note:"PTO_Notes_Config",reviewer:"PTO_Reviewer_Config",signOff:"PTO_SignOff_Config"},1:{note:"PTO_Notes_Import",reviewer:"PTO_Reviewer_Import",signOff:"PTO_SignOff_Import"},2:{note:"PTO_Notes_Headcount",reviewer:"PTO_Reviewer_Headcount",signOff:"PTO_SignOff_Headcount"},3:{note:"PTO_Notes_Validate",reviewer:"PTO_Reviewer_Validate",signOff:"PTO_SignOff_Validate"},4:{note:"PTO_Notes_Review",reviewer:"PTO_Reviewer_Review",signOff:"PTO_SignOff_Review"},5:{note:"PTO_Notes_JE",reviewer:"PTO_Reviewer_JE",signOff:"PTO_SignOff_JE"},6:{note:"PTO_Notes_Archive",reviewer:"PTO_Reviewer_Archive",signOff:"PTO_SignOff_Archive"}},ln={0:"PTO_Complete_Config",1:"PTO_Complete_Import",2:"PTO_Complete_Headcount",3:"PTO_Complete_Validate",4:"PTO_Complete_Review",5:"PTO_Complete_JE",6:"PTO_Complete_Archive"};var ne=[{id:0,title:"Configuration",summary:"Set the analysis date, accounting period, and review details for this run.",description:"Complete this step first to ensure all downstream calculations use the correct period settings.",actionLabel:"Configure Workbook",secondaryAction:{sheet:"SS_PF_Config",label:"Open Config Sheet"}},{id:1,title:"Import PTO Data",summary:"Pull your latest PTO export from payroll and paste it into PTO_Data.",description:"Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",actionLabel:"Import Sample Data",secondaryAction:{sheet:"PTO_Data",label:"Open Data Sheet"}},{id:2,title:"Headcount Review",summary:"Quick check to make sure your roster matches your PTO data.",description:"Compare employees in PTO_Data against your employee roster to catch any discrepancies.",actionLabel:"Open Headcount Review",secondaryAction:{sheet:"SS_Employee_Roster",label:"Open Sheet"}},{id:3,title:"Data Quality Review",summary:"Scan your PTO data for potential errors before crunching numbers.",description:"Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",actionLabel:"Click to Run Quality Check"},{id:4,title:"PTO Accrual Review",summary:"Review the calculated liability for each employee and compare to last period.",description:"The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",actionLabel:"Click to Perform Review"},{id:5,title:"Journal Entry Prep",summary:"Generate a balanced journal entry, run validation checks, and export when ready.",description:"Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",actionLabel:"Open Journal Draft",secondaryAction:{sheet:"PTO_JE_Draft",label:"Open Sheet"}},{id:6,title:"Archive & Reset",summary:"Save this period's results and prepare for the next cycle.",description:"Archive the current analysis so it becomes the 'prior period' for your next review.",actionLabel:"Archive Run"}],Vn={0:"PTO_Homepage",1:"PTO_Data",2:"PTO_Data",3:"PTO_Analysis",4:"PTO_Analysis",5:"PTO_JE_Draft"},Hn={PTO_Homepage:0,PTO_Data:1,PTO_Analysis:4,PTO_JE_Draft:5,PTO_Archive_Summary:6,SS_PF_Config:0,SS_Employee_Roster:2};var Fn=ne.reduce((e,t)=>(e[t.id]="pending",e),{}),D={activeView:"home",activeStepId:null,focusedIndex:0,stepStatuses:Fn},O={loaded:!1,steps:{},permanents:{},completes:{},values:{},overrides:{accountingPeriod:!1,journalId:!1}},Ae=null,rt=null,Ue=null,xe=new Map,A={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},J={debitTotal:null,creditTotal:null,difference:null,lineAmountSum:null,analysisChangeTotal:null,jeChangeTotal:null,loading:!1,lastError:null,validationRun:!1,issues:[]},Y={hasRun:!1,loading:!1,acknowledged:!1,balanceIssues:[],zeroBalances:[],accrualOutliers:[],totalIssues:0,totalEmployees:0},X={cleanDataReady:!1,employeeCount:0,lastRun:null,loading:!1,lastError:null,missingPayRates:[],missingDepartments:[],ignoredMissingPayRates:new Set,completenessCheck:{accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null}};async function Un(){var e;try{Ae=document.getElementById("app"),rt=document.getElementById("loading"),await zn(),await qn(),(e=window.PrairieForge)!=null&&e.loadSharedConfig&&await window.PrairieForge.loadSharedConfig();let t=Be(Ne);await Le(t.sheetName,t.title,t.subtitle),await Gn(),rt&&rt.remove(),Ae&&(Ae.hidden=!1),oe()}catch(t){throw console.error("[PTO] Module initialization failed:",t),t}}async function Gn(){if(se())try{await Excel.run(async e=>{e.workbook.worksheets.onActivated.add(Jn),await e.sync(),console.log("[PTO] Worksheet change listener registered")})}catch(e){console.warn("[PTO] Could not set up worksheet listener:",e)}}async function Jn(e){try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e.worksheetId);n.load("name"),await t.sync();let a=n.name,o=Hn[a];if(console.log(`[PTO] Tab changed to: ${a} \u2192 Step ${o}`),o!==void 0&&o!==D.activeStepId){let s=STEPS.findIndex(i=>i.id===o);if(s>=0){let i=o===0?"config":"step";D.activeView=i,D.activeStepId=o,D.focusedIndex=s,oe()}}})}catch(t){console.warn("[PTO] Error handling worksheet change:",t)}}async function zn(){try{await We(Ne),console.log(`[PTO] Tab visibility applied for ${Ne}`)}catch(e){console.warn("[PTO] Could not apply tab visibility:",e)}}async function qn(){var e;if(!K()){O.loaded=!0;return}try{let t=await xt(me),n={};(e=window.PrairieForge)!=null&&e.loadSharedConfig&&(await window.PrairieForge.loadSharedConfig(),window.PrairieForge._sharedConfigCache&&window.PrairieForge._sharedConfigCache.forEach((s,i)=>{n[i]=s}));let a={...t},o={SS_Default_Reviewer:w.reviewerName,Default_Reviewer:w.reviewerName,PTO_Reviewer:w.reviewerName,SS_Company_Name:w.companyName,Company_Name:w.companyName,SS_Payroll_Provider:w.payrollProvider,Payroll_Provider_Link:w.payrollProvider,SS_Accounting_Software:w.accountingSoftware,Accounting_Software_Link:w.accountingSoftware};Object.entries(o).forEach(([s,i])=>{n[s]&&!a[i]&&(a[i]=n[s])}),Object.entries(n).forEach(([s,i])=>{s.startsWith("PTO_")&&i&&(a[s]=i)}),O.permanents=await Yn(),O.values=a||{},O.overrides.accountingPeriod=!!(a!=null&&a[w.accountingPeriod]),O.overrides.journalId=!!(a!=null&&a[w.journalEntryId]),Object.entries(Ge).forEach(([s,i])=>{var d,r,l;O.steps[s]={notes:(d=a[i.note])!=null?d:"",reviewer:(r=a[i.reviewer])!=null?r:"",signOffDate:(l=a[i.signOff])!=null?l:""}}),O.completes=Object.entries(ln).reduce((s,[i,d])=>{var r;return s[i]=(r=a[d])!=null?r:"",s},{}),O.loaded=!0}catch(t){console.warn("PTO: unable to load configuration fields",t),O.loaded=!0}}async function Yn(){let e={};if(!K())return e;let t=new Map;Object.entries(Ge).forEach(([n,a])=>{a.note&&t.set(a.note.trim(),Number(n))});try{await Excel.run(async n=>{let a=n.workbook.tables.getItemOrNullObject(me[0]);if(await n.sync(),a.isNullObject)return;let o=a.getDataBodyRange(),s=a.getHeaderRowRange();o.load("values"),s.load("values"),await n.sync();let d=(s.values[0]||[]).map(l=>String(l||"").trim().toLowerCase()),r={field:d.findIndex(l=>l==="field"||l==="field name"||l==="setting"),permanent:d.findIndex(l=>l==="permanent"||l==="persist")};r.field===-1||r.permanent===-1||(o.values||[]).forEach(l=>{let p=String(l[r.field]||"").trim(),f=t.get(p);if(f==null)return;let u=xa(l[r.permanent]);e[f]=u})})}catch(n){console.warn("PTO: unable to load permanent flags",n)}return e}function oe(){var d;if(!Ae)return;let e=D.focusedIndex<=0?"disabled":"",t=D.focusedIndex>=ne.length-1?"disabled":"",n=D.activeView==="step"&&D.activeStepId!=null,o=D.activeView==="config"?cn():n?ta(D.activeStepId):`${Kn()}${Qn()}`;Ae.innerHTML=`
        <div class="pf-root">
            <div class="pf-brand-float" aria-hidden="true">
                <span class="pf-brand-wave"></span>
            </div>
            <header class="pf-banner">
                <div class="pf-nav-bar">
                    <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                        ${Ht}
                        <span class="sr-only">Previous step</span>
                    </button>
                    <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                        ${$t}
                        <span class="sr-only">Module Home</span>
                    </button>
                    <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                        ${jt}
                        <span class="sr-only">Module Selector</span>
                    </button>
                    <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                        ${Ut}
                        <span class="sr-only">Next step</span>
                    </button>
                    <span class="pf-nav-divider"></span>
                    <div class="pf-quick-access-wrapper">
                        <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                            ${Lt}
                            <span class="sr-only">Quick Access Menu</span>
                        </button>
                        <div id="quick-access-dropdown" class="pf-quick-dropdown hidden">
                            <div class="pf-quick-dropdown-header">Quick Access</div>
                            <button id="nav-config" class="pf-quick-item pf-clickable" type="button">
                                ${Ve}
                                <span>Configuration</span>
                            </button>
                        </div>
                    </div>
                </div>
            </header>
            ${o}
            <footer class="pf-brand-footer">
                <div class="pf-brand-text">
                    <div class="pf-brand-label">prairie.forge</div>
                    <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${$n}</div>
                    <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
                </div>
            </footer>
        </div>
    `;let s=D.activeView==="home"||D.activeView!=="step"&&D.activeView!=="config",i=document.getElementById("pf-info-fab-pto");if(s)i&&i.remove();else if((d=window.PrairieForge)!=null&&d.mountInfoFab){let r=Wn(D.activeStepId);PrairieForge.mountInfoFab({title:r.title,content:r.content,buttonId:"pf-info-fab-pto"})}na(),ia(),s?Pt():Qe()}function Wn(e){switch(e){case 0:return{title:"Configuration",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Sets up the key parameters for your PTO review. Complete this before importing data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Key Fields</h4>
                        <ul>
                            <li><strong>Analysis Date</strong> \u2014 The period-end date (e.g., 11/30/2024)</li>
                            <li><strong>Accounting Period</strong> \u2014 Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> \u2014 Reference number for your accounting system</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>The accounting period and JE ID auto-generate based on your analysis date, but you can override them if needed.</p>
                    </div>
                `};case 1:return{title:"Data Import",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Gets your PTO data into the workbook. Pull a report from your payroll provider and paste it into PTO_Data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Required Columns</h4>
                        <p>Your payroll export should include:</p>
                        <ul>
                            <li><strong>Employee Name</strong> \u2014 Full name (used to match against roster)</li>
                            <li><strong>Accrual Rate</strong> \u2014 Hours accrued per pay period</li>
                            <li><strong>Carry Over</strong> \u2014 Hours carried from prior year</li>
                            <li><strong>YTD Accrued</strong> \u2014 Total hours accrued this year</li>
                            <li><strong>YTD Used</strong> \u2014 Total hours used this year</li>
                            <li><strong>Balance</strong> \u2014 Current available hours</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Column headers don't need to match exactly\u2014the system is flexible with naming. Just make sure each field is present.</p>
                    </div>
                `};case 2:return{title:"Headcount Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Compares employee counts between your roster and PTO data to catch discrepancies early.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Data Sources</h4>
                        <ul>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Your centralized employee list</li>
                            <li><strong>PTO_Data</strong> \u2014 The data you just imported</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F50D} What to Look For</h4>
                        <ul>
                            <li><strong>In Roster, Not in PTO</strong> \u2014 May need to add PTO records</li>
                            <li><strong>In PTO, Not in Roster</strong> \u2014 Could be terminated employees</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>If discrepancies are expected (e.g., contractors without PTO), you can skip this check.</p>
                    </div>
                `};case 3:return{title:"Data Quality Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Scans your PTO data for anomalies that could cause problems later in the process.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Balance Issues (Critical)</h4>
                        <p>Flags when:</p>
                        <ul>
                            <li><strong>Negative Balance</strong> \u2014 Balance is less than zero</li>
                            <li><strong>Overdrawn</strong> \u2014 Used more than available (YTD Used > Carry Over + YTD Accrued)</li>
                        </ul>
                        <p class="pf-info-note">Usually indicates missing accrual entries or data errors in payroll.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} High Accrual Rates (Warning)</h4>
                        <p>Employees with Accrual Rate > 8 hours/period may have data entry errors.</p>
                        <p class="pf-info-note">Most bi-weekly accruals are 3-6 hours.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>You can acknowledge issues and proceed, but it's best to fix them in your source system first.</p>
                    </div>
                `};case 4:return{title:"PTO Accrual Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Calculates the PTO liability for each employee and compares it to last period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Data Sources</h4>
                        <ul>
                            <li><strong>PTO_Data</strong> \u2014 Your imported PTO balances</li>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Department assignments</li>
                            <li><strong>PR_Archive_Summary</strong> \u2014 Pay rates from payroll history</li>
                            <li><strong>PTO_Archive_Summary</strong> \u2014 Last period's liability (for comparison)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4B0} How Liability is Calculated</h4>
                        <div class="pf-info-formula">
                            Liability = Balance (hours) \xD7 Hourly Rate
                        </div>
                        <p class="pf-info-note">Hourly rate comes from Regular Earnings \xF7 80 hours in your payroll history.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4C8} How Change is Calculated</h4>
                        <div class="pf-info-formula">
                            Change = Current Liability \u2212 Prior Period Liability
                        </div>
                        <ul>
                            <li><span style="color: #30d158;">Positive</span> = Liability went up (book expense)</li>
                            <li><span style="color: #ff453a;">Negative</span> = Liability went down (reverse expense)</li>
                        </ul>
                    </div>
                `};case 5:return{title:"Journal Entry Prep",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Generates a balanced journal entry from your PTO analysis, ready for upload to your accounting system.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4DD} How the JE Works</h4>
                        <p>Groups the <strong>Change</strong> amounts by department:</p>
                        <ul>
                            <li><span style="color: #30d158;">Positive Change</span> \u2192 Debit expense account</li>
                            <li><span style="color: #ff453a;">Negative Change</span> \u2192 Credit expense account</li>
                        </ul>
                        <p>The offset always goes to <strong>21540</strong> (Accrued PTO liability).</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F3E2} Department \u2192 Account Mapping</h4>
                        <table style="width:100%; font-size: 12px; margin-top: 8px;">
                            <tr><td>General & Admin</td><td style="text-align:right">64110</td></tr>
                            <tr><td>R&D</td><td style="text-align:right">62110</td></tr>
                            <tr><td>Marketing</td><td style="text-align:right">61610</td></tr>
                            <tr><td>Sales & Marketing</td><td style="text-align:right">61110</td></tr>
                            <tr><td>COGS Onboarding</td><td style="text-align:right">53110</td></tr>
                            <tr><td>COGS Prof. Services</td><td style="text-align:right">56110</td></tr>
                            <tr><td>COGS Support</td><td style="text-align:right">52110</td></tr>
                            <tr><td>Client Success</td><td style="text-align:right">61811</td></tr>
                        </table>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u2705 Validation Checks</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> \u2014 Entry must balance</li>
                            <li><strong>Line Amounts = $0</strong> \u2014 Net change must be zero</li>
                            <li><strong>JE Matches Analysis</strong> \u2014 Totals tie back to your data</li>
                        </ul>
                    </div>
                `};case 6:return{title:"Archive & Reset",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Saves this period's results so they become the "prior period" for your next review.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4E6} What Gets Saved</h4>
                        <ul>
                            <li><strong>PTO_Archive_Summary</strong> \u2014 Employee name, liability amount, and analysis date</li>
                            <li>This data is used to calculate the "Change" column next period</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Important</h4>
                        <p>Only the <strong>most recent period</strong> is kept in the archive. Running archive again will overwrite the previous data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Make sure your JE has been uploaded to your accounting system before archiving.</p>
                    </div>
                `};default:return{title:"PTO Accrual",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F44B} Welcome to PTO Accrual</h4>
                        <p>This module helps you calculate PTO liabilities and generate journal entries each period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Workflow Overview</h4>
                        <ol style="margin: 8px 0; padding-left: 20px;">
                            <li>Configure your period settings</li>
                            <li>Import PTO data from payroll</li>
                            <li>Review headcount alignment</li>
                            <li>Check data quality</li>
                            <li>Review calculated liabilities</li>
                            <li>Generate and export journal entry</li>
                            <li>Archive for next period</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <p>Click a step card to get started, or tap the <strong>\u24D8</strong> button on any step for detailed guidance.</p>
                    </div>
                `}}}function Kn(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">PTO Accrual</h2>
            <p class="pf-hero-copy">${Ln}</p>
        </section>
    `}function Qn(){return`
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${ne.map((e,t)=>Xn(e,t)).join("")}
            </div>
        </section>
    `}function Xn(e,t){let n=D.stepStatuses[e.id]||"pending",a=D.activeView==="step"&&D.focusedIndex===t?"pf-step-card--active":"",o=Bt(va(e.id));return`
        <article class="pf-step-card pf-clickable ${a}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${e.title}</h3>
        </article>
    `}function Zn(e){let t=ne.filter(o=>o.id!==6).map(o=>({id:o.id,title:o.title,complete:ra(o.id)})),n=t.every(o=>o.complete),a=t.map(o=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${o.complete?"is-active":""}" aria-pressed="${o.complete}">
                        ${Pe}
                    </span>
                    <div>
                        <h3>${v(o.title)}</h3>
                        <p class="pf-config-subtext">${o.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${v(he)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${v(e.title)}</h2>
            <p class="pf-hero-copy">${v(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${a}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Only enabled when all steps above are complete.</p>
                </div>
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn" id="archive-run-btn" ${n?"":"disabled"}>Archive</button>
                </div>
            </article>
        </section>
    `}function cn(){if(!O.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=tn(le(w.payrollDate)),t=tn(le(w.accountingPeriod)),n=le(w.journalEntryId),a=le(w.accountingSoftware),o=le(w.payrollProvider),s=le(w.companyName),i=le(w.reviewerName),d=we(0),r=!!O.permanents[0],l=!!(vn(O.completes[0])||d.signOffDate),p=be(d==null?void 0:d.reviewer),f=(d==null?void 0:d.signOffDate)||"";return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${v(he)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every PTO run.</p>
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
                        <input type="text" id="config-user-name" value="${v(i)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${v(e)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${v(t)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${v(n)}" placeholder="PTO-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${v(s)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${v(o)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${v(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${pe({textareaId:"config-notes",value:d.notes||"",permanentId:"config-notes-lock",isPermanent:r,hintId:"",saveButtonId:"config-notes-save"})}
            ${fe({reviewerInputId:"config-reviewer",reviewerValue:p,signoffInputId:"config-signoff-date",signoffValue:f,isComplete:l,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"})}
        </section>
    `}function ea(e){let t=we(1),n=!!O.permanents[1],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(O.completes[1])||o),i=le(w.payrollProvider);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${v(he)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${v(e.title)}</h2>
            <p class="pf-hero-copy">${v(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export, then paste into PTO_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${G(i?`<a href="${v(i)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${nt}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${nt}</button>`,"Provider")}
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${Ve}</button>`,"PTO_Data")}
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${Gt}</button>`,"Clear")}
                </div>
            </article>
            ${pe({textareaId:"step-notes-1",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-1",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-1"})}
            ${fe({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
        </section>
    `}function ta(e){let t=ne.find(d=>d.id===e);if(!t)return"";if(e===0)return cn();if(e===1)return ea(t);if(e===2)return _a(t);if(e===3)return Pa(t);if(e===4)return Ia(t);if(e===5)return Ra(t);if(t.id===6)return Zn(t);let n=we(e),a=!!O.permanents[e],o=be(n==null?void 0:n.reviewer),s=(n==null?void 0:n.signOffDate)||"",i=!!(Se(O.completes[e])||s);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${v(he)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${v(t.title)}</h2>
            <p class="pf-hero-copy">${v(t.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${pe({textareaId:`step-notes-${e}`,value:(n==null?void 0:n.notes)||"",permanentId:`step-notes-lock-${e}`,isPermanent:a,hintId:"",saveButtonId:`step-notes-save-${e}`})}
            ${fe({reviewerInputId:`step-reviewer-${e}`,reviewerValue:o,signoffInputId:`step-signoff-${e}`,signoffValue:s,isComplete:i,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`})}
        </section>
    `}function na(){var n,a,o,s,i,d;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",async()=>{var l;let r=Be(Ne);await Le(r.sheetName,r.title,r.subtitle),$e({activeView:"home",activeStepId:null}),(l=document.getElementById("pf-hero"))==null||l.scrollIntoView({behavior:"smooth",block:"start"})}),(a=document.getElementById("nav-selector"))==null||a.addEventListener("click",()=>{window.location.href=Bn}),(o=document.getElementById("nav-prev"))==null||o.addEventListener("click",()=>Yt(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>Yt(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",r=>{r.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",r=>{!(t!=null&&t.contains(r.target))&&!(e!=null&&e.contains(r.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(i=document.getElementById("nav-config"))==null||i.addEventListener("click",async()=>{t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"),await Zt()}),(d=document.getElementById("showConfigSheets"))==null||d.addEventListener("click",async()=>{await Zt()}),document.querySelectorAll("[data-step-card]").forEach(r=>{let l=Number(r.getAttribute("data-step-index")),p=Number(r.getAttribute("data-step-id"));r.addEventListener("click",()=>De(l,p))}),D.activeView==="config"?oa():D.activeView==="step"&&D.activeStepId!=null&&aa(D.activeStepId)}function aa(e){var p,f,u,c,g,y,b,m,S,E,R,C,j,L,B,W,F;let t=e===2?document.getElementById("step-notes-input"):document.getElementById(`step-notes-${e}`),n=e===2?document.getElementById("step-reviewer-name"):document.getElementById(`step-reviewer-${e}`),a=e===2?document.getElementById("step-signoff-date"):document.getElementById(`step-signoff-${e}`),o=document.getElementById("step-back-btn"),s=e===2?document.getElementById("step-notes-lock-2"):document.getElementById(`step-notes-lock-${e}`),i=e===2?document.getElementById("step-notes-save-2"):document.getElementById(`step-notes-save-${e}`);i==null||i.addEventListener("click",async()=>{let T=(t==null?void 0:t.value)||"";await te(e,"notes",T),ce(i,!0)});let d=e===2?document.getElementById("headcount-signoff-save"):document.getElementById(`step-signoff-save-${e}`);d==null||d.addEventListener("click",async()=>{let T=(n==null?void 0:n.value)||"";await te(e,"reviewer",T),ce(d,!0)}),ot();let r=e===2?"headcount-signoff-toggle":`step-signoff-toggle-${e}`,l=e===2?"step-signoff-date":`step-signoff-${e}`;yn(e,{buttonId:r,inputId:l,canActivate:e===2?()=>{var P;return!bn()||((P=document.getElementById("step-notes-input"))==null?void 0:P.value.trim())||""?!0:(Z("Please enter a brief explanation of the headcount differences before completing this step.","info"),!1)}:null,onComplete:sa(e)}),o==null||o.addEventListener("click",async()=>{let T=Be(Ne);await Le(T.sheetName,T.title,T.subtitle),$e({activeView:"home",activeStepId:null})}),s==null||s.addEventListener("click",async()=>{let T=!s.classList.contains("is-locked");at(s,T),await mn(e,T)}),e===6&&((p=document.getElementById("archive-run-btn"))==null||p.addEventListener("click",()=>{})),e===1&&((f=document.getElementById("import-open-data-btn"))==null||f.addEventListener("click",()=>fn("PTO_Data")),(u=document.getElementById("import-clear-btn"))==null||u.addEventListener("click",()=>ga())),e===2&&((c=document.getElementById("headcount-skip-btn"))==null||c.addEventListener("click",()=>{A.skipAnalysis=!A.skipAnalysis;let T=document.getElementById("headcount-skip-btn");T==null||T.classList.toggle("is-active",A.skipAnalysis),A.skipAnalysis&&rn(),sn()}),(g=document.getElementById("headcount-run-btn"))==null||g.addEventListener("click",()=>ct()),(y=document.getElementById("headcount-refresh-btn"))==null||y.addEventListener("click",()=>ct()),ja(),A.skipAnalysis&&rn(),sn()),e===3&&((b=document.getElementById("quality-run-btn"))==null||b.addEventListener("click",()=>Kt()),(m=document.getElementById("quality-refresh-btn"))==null||m.addEventListener("click",()=>Kt()),(S=document.getElementById("quality-acknowledge-btn"))==null||S.addEventListener("click",()=>ca())),e===4&&((E=document.getElementById("analysis-refresh-btn"))==null||E.addEventListener("click",()=>Qt()),(R=document.getElementById("analysis-run-btn"))==null||R.addEventListener("click",()=>Qt()),(C=document.getElementById("payrate-save-btn"))==null||C.addEventListener("click",Wt),(j=document.getElementById("payrate-ignore-btn"))==null||j.addEventListener("click",la),(L=document.getElementById("payrate-input"))==null||L.addEventListener("keydown",T=>{T.key==="Enter"&&Wt()})),e===5&&((B=document.getElementById("je-create-btn"))==null||B.addEventListener("click",()=>pa()),(W=document.getElementById("je-run-btn"))==null||W.addEventListener("click",()=>pn()),(F=document.getElementById("je-export-btn"))==null||F.addEventListener("click",()=>fa()))}function oa(){var d,r,l,p,f;At("config-payroll-date",{onChange:u=>{if(re(w.payrollDate,u),!!u){if(!O.overrides.accountingPeriod){let c=ka(u);if(c){let g=document.getElementById("config-accounting-period");g&&(g.value=c),re(w.accountingPeriod,c)}}if(!O.overrides.journalId){let c=Oa(u);if(c){let g=document.getElementById("config-journal-id");g&&(g.value=c),re(w.journalEntryId,c)}}}}});let e=document.getElementById("config-accounting-period");e==null||e.addEventListener("change",u=>{O.overrides.accountingPeriod=!!u.target.value,re(w.accountingPeriod,u.target.value||"")});let t=document.getElementById("config-journal-id");t==null||t.addEventListener("change",u=>{O.overrides.journalId=!!u.target.value,re(w.journalEntryId,u.target.value.trim())}),(d=document.getElementById("config-company-name"))==null||d.addEventListener("change",u=>{re(w.companyName,u.target.value.trim())}),(r=document.getElementById("config-payroll-provider"))==null||r.addEventListener("change",u=>{re(w.payrollProvider,u.target.value.trim())}),(l=document.getElementById("config-accounting-link"))==null||l.addEventListener("change",u=>{re(w.accountingSoftware,u.target.value.trim())}),(p=document.getElementById("config-user-name"))==null||p.addEventListener("change",u=>{re(w.reviewerName,u.target.value.trim())});let n=document.getElementById("config-notes");n==null||n.addEventListener("input",u=>{te(0,"notes",u.target.value)});let a=document.getElementById("config-notes-lock");a==null||a.addEventListener("click",async()=>{let u=!a.classList.contains("is-locked");at(a,u),await mn(0,u)});let o=document.getElementById("config-notes-save");o==null||o.addEventListener("click",async()=>{n&&(await te(0,"notes",n.value),ce(o,!0))});let s=document.getElementById("config-reviewer");s==null||s.addEventListener("change",u=>{let c=u.target.value.trim();te(0,"reviewer",c);let g=document.getElementById("config-signoff-date");if(c&&g&&!g.value){let y=ut();g.value=y,te(0,"signOffDate",y),hn(0,!0)}}),(f=document.getElementById("config-signoff-date"))==null||f.addEventListener("change",u=>{te(0,"signOffDate",u.target.value||"")});let i=document.getElementById("config-signoff-save");i==null||i.addEventListener("click",async()=>{var g,y;let u=((g=s==null?void 0:s.value)==null?void 0:g.trim())||"",c=((y=document.getElementById("config-signoff-date"))==null?void 0:y.value)||"";await te(0,"reviewer",u),await te(0,"signOffDate",c),ce(i,!0)}),ot(),yn(0,{buttonId:"config-signoff-toggle",inputId:"config-signoff-date",onComplete:()=>{Ca(),dn(0),un()}})}function De(e,t=null){if(e<0||e>=ne.length)return;Ue=e;let n=t!=null?t:ne[e].id;$e({focusedIndex:e,activeView:n===0?"config":"step",activeStepId:n});let o=Vn[n];o&&fn(o),n===2&&!A.hasAnalyzed&&(wn(),ct())}function sa(e){return e===6?null:()=>dn(e)}function dn(e){let t=ne.findIndex(a=>a.id===e);if(t===-1)return;let n=t+1;n<ne.length&&(De(n,ne[n].id),un())}function un(){let e=[document.querySelector(".pf-root"),document.querySelector(".pf-step-guide"),document.body];for(let t of e)t&&t.scrollTo({top:0,behavior:"smooth"});window.scrollTo({top:0,behavior:"smooth"})}function Yt(e){let t=D.focusedIndex+e,n=Math.max(0,Math.min(ne.length-1,t));De(n,ne[n].id)}function ia(){if(Ue===null)return;let e=document.querySelector(`[data-step-index="${Ue}"]`);Ue=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}function ra(e){return vn(O.completes[e])}function $e(e){e.stepStatuses&&(D.stepStatuses={...D.stepStatuses,...e.stepStatuses}),Object.assign(D,{...e,stepStatuses:D.stepStatuses}),oe()}function se(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}async function Wt(){let e=document.getElementById("payrate-input");if(!e)return;let t=parseFloat(e.value),n=e.dataset.employee,a=parseInt(e.dataset.row,10);if(isNaN(t)||t<=0){Z("Please enter a valid pay rate greater than 0.","info");return}if(!n||isNaN(a)){console.error("Missing employee data on input");return}ee(!0,"Updating pay rate...");try{await Excel.run(async o=>{let s=o.workbook.worksheets.getItem("PTO_Analysis"),i=s.getCell(a-1,3);i.values=[[t]];let d=s.getCell(a-1,8);d.load("values"),await o.sync();let l=(Number(d.values[0][0])||0)*t,p=s.getCell(a-1,9);p.values=[[l]];let f=s.getCell(a-1,10);f.load("values"),await o.sync();let u=Number(f.values[0][0])||0,c=l-u,g=s.getCell(a-1,11);g.values=[[c]],await o.sync()}),X.missingPayRates=X.missingPayRates.filter(o=>o.name!==n),ee(!1),De(3,3)}catch(o){console.error("Failed to save pay rate:",o),Z(`Failed to save pay rate: ${o.message}`,"error"),ee(!1)}}function la(){let e=document.getElementById("payrate-input");if(!e)return;let t=e.dataset.employee;t&&(X.ignoredMissingPayRates.add(t),X.missingPayRates=X.missingPayRates.filter(n=>n.name!==t)),De(3,3)}async function Kt(){if(!se()){Z("Excel is not available. Open this module inside Excel to run quality check.","info");return}Y.loading=!0,ee(!0,"Analyzing data quality..."),ce(document.getElementById("quality-save-btn"),!1);try{await Excel.run(async t=>{var b;let a=t.workbook.worksheets.getItem("PTO_Data").getUsedRangeOrNullObject();a.load("values"),await t.sync();let o=a.isNullObject?[]:a.values||[];if(!o.length||o.length<2)throw new Error("PTO_Data is empty or has no data rows.");let s=(o[0]||[]).map(m=>z(m));console.log("[Data Quality] PTO_Data headers:",o[0]);let i=s.findIndex(m=>m==="employee name"||m==="employeename");i===-1&&(i=s.findIndex(m=>m.includes("employee")&&m.includes("name"))),i===-1&&(i=s.findIndex(m=>m==="name"||m.includes("name")&&!m.includes("company")&&!m.includes("form"))),console.log("[Data Quality] Employee name column index:",i,"Header:",(b=o[0])==null?void 0:b[i]);let d=$(s,["balance"]),r=$(s,["accrual rate","accrualrate"]),l=$(s,["carry over","carryover"]),p=$(s,["ytd accrued","ytdaccrued"]),f=$(s,["ytd used","ytdused"]),u=[],c=[],g=[],y=o.slice(1);y.forEach((m,S)=>{let E=S+2,R=i!==-1?String(m[i]||"").trim():`Row ${E}`;if(!R)return;let C=d!==-1&&Number(m[d])||0,j=r!==-1&&Number(m[r])||0,L=l!==-1&&Number(m[l])||0,B=p!==-1&&Number(m[p])||0,W=f!==-1&&Number(m[f])||0,F=L+B;C<0?u.push({name:R,issue:`Negative balance: ${C.toFixed(2)} hrs`,rowIndex:E}):W>F&&F>0&&u.push({name:R,issue:`Used ${W.toFixed(0)} hrs but only ${F.toFixed(0)} available`,rowIndex:E}),C===0&&(L>0||B>0)&&c.push({name:R,rowIndex:E}),j>8&&g.push({name:R,accrualRate:j,rowIndex:E})}),Y.balanceIssues=u,Y.zeroBalances=c,Y.accrualOutliers=g,Y.totalIssues=u.length,Y.totalEmployees=y.filter(m=>m.some(S=>S!==null&&S!=="")).length,Y.hasRun=!0});let e=Y.balanceIssues.length>0;$e({stepStatuses:{3:e?"blocked":"complete"}})}catch(e){console.error("Data quality check error:",e),Z(`Quality check failed: ${e.message}`,"error"),Y.hasRun=!1}finally{Y.loading=!1,ee(!1),oe()}}function ca(){Y.acknowledged=!0,$e({stepStatuses:{3:"complete"}}),oe()}async function da(){if(se())try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem("PTO_Data"),n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=t.getUsedRangeOrNullObject();if(a.load("values"),n.load("isNullObject"),await e.sync(),n.isNullObject){X.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let o=n.getUsedRangeOrNullObject();o.load("values"),await e.sync();let s=a.isNullObject?[]:a.values||[],i=o.isNullObject?[]:o.values||[];if(!s.length||!i.length){X.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let d=(p,f,u)=>{let c=(p[0]||[]).map(b=>z(b)),g=$(c,f);return g===-1?null:p.slice(1).reduce((b,m)=>b+(Number(m[g])||0),0)},r=[{key:"accrualRate",aliases:["accrual rate","accrualrate"]},{key:"carryOver",aliases:["carry over","carryover","carry_over"]},{key:"ytdAccrued",aliases:["ytd accrued","ytdaccrued","ytd_accrued"]},{key:"ytdUsed",aliases:["ytd used","ytdused","ytd_used"]},{key:"balance",aliases:["balance"]}],l={};for(let p of r){let f=d(s,p.aliases,"PTO_Data"),u=d(i,p.aliases,"PTO_Analysis");if(f===null||u===null)l[p.key]=null;else{let c=Math.abs(f-u)<.01;l[p.key]={match:c,ptoData:f,ptoAnalysis:u}}}X.completenessCheck=l})}catch(e){console.error("Completeness check failed:",e)}}async function Qt(){if(!se()){Z("Excel is not available. Open this module inside Excel to run analysis.","info");return}ee(!0,"Running analysis...");try{await wn(),await da(),X.cleanDataReady=!0,oe()}catch(e){console.error("Full analysis error:",e),Z(`Analysis failed: ${e.message}`,"error")}finally{ee(!1)}}async function pn(){if(!se()){Z("Excel is not available. Open this module inside Excel to run journal checks.","info");return}J.loading=!0,J.lastError=null,ce(document.getElementById("je-save-btn"),!1),oe();try{let e=await Excel.run(async t=>{let a=t.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();a.load("values");let o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis");o.load("isNullObject"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty. Generate the JE first.");let i=(s[0]||[]).map(E=>z(E)),d=$(i,["debit"]),r=$(i,["credit"]),l=$(i,["lineamount","line amount"]),p=$(i,["account number","accountnumber"]);if(d===-1||r===-1)throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");let f=0,u=0,c=0,g=0;s.slice(1).forEach(E=>{let R=Number(E[d])||0,C=Number(E[r])||0,j=l!==-1&&Number(E[l])||0,L=p!==-1?String(E[p]||"").trim():"";f+=R,u+=C,c+=j,L&&L!=="21540"&&(g+=j)});let y=0;if(!o.isNullObject){let E=o.getUsedRangeOrNullObject();E.load("values"),await t.sync();let R=E.isNullObject?[]:E.values||[];if(R.length>1){let C=(R[0]||[]).map(L=>z(L)),j=$(C,["change"]);j!==-1&&R.slice(1).forEach(L=>{y+=Number(L[j])||0})}}let b=f-u,m=[];Math.abs(b)>=.01?m.push({check:"Debits = Credits",passed:!1,detail:b>0?`Debits exceed credits by $${Math.abs(b).toLocaleString(void 0,{minimumFractionDigits:2})}`:`Credits exceed debits by $${Math.abs(b).toLocaleString(void 0,{minimumFractionDigits:2})}`}):m.push({check:"Debits = Credits",passed:!0,detail:""}),Math.abs(c)>=.01?m.push({check:"Line Amounts Sum to Zero",passed:!1,detail:`Line amounts sum to $${c.toLocaleString(void 0,{minimumFractionDigits:2})} (should be $0.00)`}):m.push({check:"Line Amounts Sum to Zero",passed:!0,detail:""});let S=Math.abs(g-y);return S>=.01?m.push({check:"JE Matches Analysis Total",passed:!1,detail:`JE expense total ($${g.toLocaleString(void 0,{minimumFractionDigits:2})}) differs from PTO_Analysis Change total ($${y.toLocaleString(void 0,{minimumFractionDigits:2})}) by $${S.toLocaleString(void 0,{minimumFractionDigits:2})}`}):m.push({check:"JE Matches Analysis Total",passed:!0,detail:""}),{debitTotal:f,creditTotal:u,difference:b,lineAmountSum:c,jeChangeTotal:g,analysisChangeTotal:y,issues:m,validationRun:!0}});Object.assign(J,e,{lastError:null})}catch(e){console.warn("PTO JE summary:",e),J.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",J.debitTotal=null,J.creditTotal=null,J.difference=null,J.lineAmountSum=null,J.jeChangeTotal=null,J.analysisChangeTotal=null,J.issues=[],J.validationRun=!1}finally{J.loading=!1,oe()}}var ua={"general & administrative":"64110","general and administrative":"64110","g&a":"64110","research & development":"62110","research and development":"62110","r&d":"62110",marketing:"61610","cogs onboarding":"53110","cogs prof. services":"56110","cogs professional services":"56110","sales & marketing":"61110","sales and marketing":"61110","cogs support":"52110","client success":"61811"},Xt="21540";async function pa(){if(!se()){Z("Excel is not available. Open this module inside Excel to create the journal entry.","info");return}ee(!0,"Creating PTO Journal Entry...");try{await Excel.run(async e=>{let t=[],n=e.workbook.tables.getItemOrNullObject(me[0]);if(n.load("isNullObject"),await e.sync(),n.isNullObject){let h=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");if(h.load("isNullObject"),await e.sync(),!h.isNullObject){let _=h.getUsedRangeOrNullObject();_.load("values"),await e.sync();let I=_.isNullObject?[]:_.values||[];t=I.length>1?I.slice(1):[]}}else{let h=n.getDataBodyRange();h.load("values"),await e.sync(),t=h.values||[]}let a=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis");if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PTO_Analysis sheet not found. Please ensure the worksheet exists.");let o=a.getUsedRangeOrNullObject();o.load("values");let s=e.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");s.load("isNullObject"),await e.sync();let i=[];if(!s.isNullObject){let h=s.getUsedRangeOrNullObject();h.load("values"),await e.sync(),i=h.isNullObject?[]:h.values||[]}let d=o.isNullObject?[]:o.values||[];if(!d.length||d.length<2)throw new Error("PTO_Analysis is empty or has no data rows. Run the analysis first (Step 4).");let r={};t.forEach(h=>{let _=String(h[1]||"").trim(),I=h[2];_&&(r[_]=I)}),(!r[w.journalEntryId]||!r[w.payrollDate])&&console.warn("[JE Draft] Missing config values - RefNumber:",r[w.journalEntryId],"TxnDate:",r[w.payrollDate]);let l=r[w.journalEntryId]||"",p=r[w.payrollDate]||"",f=r[w.accountingPeriod]||"",u="";if(p)try{let h;if(typeof p=="number"||/^\d{4,5}$/.test(String(p).trim())){let _=Number(p),I=new Date(1899,11,30);h=new Date(I.getTime()+_*24*60*60*1e3)}else h=new Date(p);if(!isNaN(h.getTime())&&h.getFullYear()>1970){let _=String(h.getMonth()+1).padStart(2,"0"),I=String(h.getDate()).padStart(2,"0"),N=h.getFullYear();u=`${_}/${I}/${N}`}else console.warn("[JE Draft] Date parsing resulted in invalid date:",p,"->",h),u=String(p)}catch(h){console.warn("[JE Draft] Could not parse TxnDate:",p,h),u=String(p)}let c=f?`${f} PTO Accrual`:"PTO Accrual",g={};if(i.length>1){let h=(i[0]||[]).map(N=>z(N)),_=$(h,["account number","accountnumber","account","acct"]),I=$(h,["account name","accountname","name","description"]);_!==-1&&I!==-1&&i.slice(1).forEach(N=>{let q=String(N[_]||"").trim(),de=String(N[I]||"").trim();q&&(g[q]=de)})}let y=(d[0]||[]).map(h=>z(h));console.log("[JE Draft] PTO_Analysis headers:",y),console.log("[JE Draft] PTO_Analysis row count:",d.length-1);let b=$(y,["department"]),m=$(y,["change"]);if(console.log("[JE Draft] Column indices - Department:",b,"Change:",m),b===-1||m===-1)throw new Error(`Could not find required columns in PTO_Analysis. Found headers: ${y.join(", ")}. Looking for "Department" (found: ${b!==-1}) and "Change" (found: ${m!==-1}).`);let S={},E=0,R=0,C=0;if(d.slice(1).forEach((h,_)=>{E++;let I=String(h[b]||"").trim(),N=h[m],q=Number(N)||0;if(_<3&&console.log(`[JE Draft] Row ${_+2}: Dept="${I}", Change raw="${N}", Change num=${q}`),!I){C++;return}if(q===0){R++;return}S[I]||(S[I]=0),S[I]+=q}),console.log(`[JE Draft] Data summary: ${E} rows, ${R} with zero change, ${C} missing dept`),console.log("[JE Draft] Department totals:",S),Object.keys(S).length===0){let h=`No journal entry lines could be created.

`;throw R===E?(h+=`All 'Change' amounts in PTO_Analysis are $0.00.

`,h+=`Common causes:
`,h+=`\u2022 Missing Pay Rate data (Liability = Balance \xD7 Pay Rate)
`,h+=`\u2022 No prior period data to compare against
`,h+=`\u2022 PTO Analysis hasn't been run yet

`,h+="Please verify Pay Rate values exist in PTO_Analysis."):C===E?(h+=`All rows are missing Department values.

`,h+="Please ensure the 'Department' column is populated in PTO_Analysis."):(h+=`Found ${E} rows but none had both a Department and non-zero Change amount.
`,h+=`\u2022 ${R} rows with zero change
`,h+=`\u2022 ${C} rows missing department`),new Error(h)}let L=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],B=[L],W=0,F=0;Object.entries(S).forEach(([h,_])=>{if(Math.abs(_)<.01)return;let I=h.toLowerCase().trim(),N=ua[I]||"",q=g[N]||"",de=_>0?Math.abs(_):0,k=_<0?Math.abs(_):0;W+=de,F+=k,B.push([l,u,N,q,_,de,k,c,h])});let T=W-F;if(Math.abs(T)>=.01){let h=T<0?Math.abs(T):0,_=T>0?Math.abs(T):0,I=g[Xt]||"Accrued PTO";B.push([l,u,Xt,I,-T,h,_,c,""])}let P=e.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");if(P.load("isNullObject"),await e.sync(),P.isNullObject)P=e.workbook.worksheets.add("PTO_JE_Draft");else{let h=P.getUsedRangeOrNullObject();h.load("isNullObject"),await e.sync(),h.isNullObject||h.clear()}if(B.length>0){let h=P.getRangeByIndexes(0,0,B.length,L.length);h.values=B;let _=P.getRangeByIndexes(0,0,1,L.length);it(_);let I=B.length-1;I>0&&(ge(P,4,I,!0),ge(P,5,I),ge(P,6,I)),h.format.autofitColumns()}await e.sync(),P.activate(),P.getRange("A1").select(),await e.sync()}),await pn()}catch(e){console.error("Create JE Draft error:",e),Z(`Unable to create Journal Entry: ${e.message}`,"error")}finally{ee(!1)}}async function fa(){if(!se()){Z("Excel is not available. Open this module inside Excel to export.","info");return}ee(!0,"Preparing JE CSV...");try{let{rows:e}=await Excel.run(async n=>{let o=n.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();o.load("values"),await n.sync();let s=o.isNullObject?[]:o.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty.");return{rows:s}}),t=Da(e);$a(`pto-je-draft-${ut()}.csv`,t)}catch(e){console.error("PTO JE export:",e),Z("Unable to export the JE draft. Confirm the sheet has data.","error")}finally{ee(!1)}}async function fn(e){if(!(!e||!se()))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e);n.activate(),n.getRange("A1").select(),await t.sync()})}catch(t){console.error(t)}}async function ga(){if(!(!se()||!await jn(`All data in PTO_Data will be permanently removed.

This action cannot be undone.`,{title:"Clear PTO Data",icon:"\u{1F5D1}\uFE0F",confirmText:"Clear Data",cancelText:"Keep Data",destructive:!0}))){ee(!0);try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("PTO_Data"),a=n.getUsedRangeOrNullObject();a.load("rowCount"),await t.sync(),!a.isNullObject&&a.rowCount>1&&(n.getRangeByIndexes(1,0,a.rowCount-1,20).clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),Z("PTO_Data cleared successfully. You can now paste new data.","success")}catch(t){console.error("Clear PTO_Data error:",t),Z(`Failed to clear PTO_Data: ${t.message}`,"error")}finally{ee(!1)}}}async function ma(){if(!se())return[];try{return await Excel.run(async e=>{let t=e.workbook.worksheets;return t.load("items/name,visibility"),await e.sync(),t.items.filter(a=>{let s=(a.name||"").toUpperCase();return s.startsWith("SS_")||s.includes("MAPPING")}).map(a=>({name:a.name,visible:a.visibility===Excel.SheetVisibility.visible}))})}catch(e){return console.error("[Config] Error reading configuration sheets:",e),[]}}function ha(){if(document.getElementById("config-sheet-modal"))return;let e=document.createElement("div");if(e.id="config-sheet-modal",e.className="pf-config-modal hidden",e.innerHTML=`
        <div class="pf-config-modal-backdrop" data-close></div>
        <div class="pf-config-modal-card">
            <div class="pf-config-modal-head">
                <h3>Configuration Sheets</h3>
                <button type="button" class="pf-config-close" data-close aria-label="Close">\xD7</button>
            </div>
            <div class="pf-config-modal-body">
                <p class="pf-config-hint">Choose a configuration or mapping sheet to unhide and open.</p>
                <div id="config-sheet-list" class="pf-config-sheet-list">Loading\u2026</div>
            </div>
        </div>
    `,document.body.appendChild(e),!document.getElementById("pf-config-modal-styles")){let t=document.createElement("style");t.id="pf-config-modal-styles",t.textContent=`
            .pf-config-modal { position: fixed; inset: 0; display: flex; align-items: center; justify-content: center; z-index: 10000; }
            .pf-config-modal.hidden { display: none; }
            .pf-config-modal-backdrop { position: absolute; inset: 0; background: rgba(0,0,0,0.6); }
            .pf-config-modal-card { position: relative; background: #0f172a; color: #e2e8f0; border-radius: 12px; padding: 20px; width: min(420px, 90%); box-shadow: 0 20px 60px rgba(0,0,0,0.35); }
            .pf-config-modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
            .pf-config-close { background: transparent; border: none; color: #e2e8f0; font-size: 20px; cursor: pointer; }
            .pf-config-hint { margin: 0 0 12px 0; color: #94a3b8; font-size: 14px; }
            .pf-config-sheet-list { display: flex; flex-direction: column; gap: 8px; max-height: 240px; overflow-y: auto; }
            .pf-config-sheet { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); border-radius: 8px; cursor: pointer; }
            .pf-config-sheet:hover { background: rgba(255,255,255,0.08); }
            .pf-config-pill { font-size: 12px; color: #a5b4fc; }
        `,document.head.appendChild(t)}}async function Zt(){ha();let e=document.getElementById("config-sheet-modal"),t=document.getElementById("config-sheet-list");if(!e||!t)return;t.textContent="Loading\u2026",e.classList.remove("hidden");let n=await ma();n.length?(t.innerHTML="",n.forEach(a=>{let o=document.createElement("button");o.type="button",o.className="pf-config-sheet",o.innerHTML=`<span>${a.name}</span><span class="pf-config-pill">${a.visible?"Visible":"Hidden"}</span>`,o.addEventListener("click",async()=>{await ya(a.name),e.classList.add("hidden")}),t.appendChild(o)})):t.textContent="No configuration sheets found.",e.querySelectorAll("[data-close]").forEach(a=>a.addEventListener("click",()=>e.classList.add("hidden")))}async function ya(e){if(!(!e||!se()))try{await Excel.run(async t=>{let n=t.workbook.worksheets,a=n.getItemOrNullObject(e);a.load("isNullObject,visibility"),await t.sync(),a.isNullObject&&(a=n.add(e)),a.visibility=Excel.SheetVisibility.visible,await t.sync(),a.activate(),a.getRange("A1").select(),await t.sync(),console.log(`[Config] Opened sheet: ${e}`)})}catch(t){console.error("[Config] Error opening sheet",e,t)}}function le(e){var n,a;let t=String(e!=null?e:"").trim();return(a=(n=O.values)==null?void 0:n[t])!=null?a:""}function be(e){var n;if(e)return e;let t=le(w.reviewerName);if(t)return t;if((n=window.PrairieForge)!=null&&n._sharedConfigCache){let a=window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer")||window.PrairieForge._sharedConfigCache.get("Default_Reviewer");if(a)return a}return""}function re(e,t,n={}){var i;let a=String(e!=null?e:"").trim();if(!a)return;O.values[a]=t!=null?t:"";let o=(i=n.debounceMs)!=null?i:0;if(!o){let d=xe.get(a);d&&clearTimeout(d),xe.delete(a),_e(a,t!=null?t:"",me);return}xe.has(a)&&clearTimeout(xe.get(a));let s=setTimeout(()=>{xe.delete(a),_e(a,t!=null?t:"",me)},o);xe.set(a,s)}function z(e){return String(e!=null?e:"").trim().toLowerCase()}function ee(e,t="Working..."){let n=document.getElementById(Mn);n&&(n.style.display="none")}function lt(){Un()}typeof Office!="undefined"&&Office.onReady?Office.onReady(()=>lt()).catch(()=>lt()):lt();function we(e){return O.steps[e]||{notes:"",reviewer:"",signOffDate:""}}function gn(e){return Ge[e]||{}}function va(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}async function te(e,t,n){let a=O.steps[e]||{notes:"",reviewer:"",signOffDate:""};a[t]=n,O.steps[e]=a;let o=gn(e),s=t==="notes"?o.note:t==="reviewer"?o.reviewer:o.signOff;if(s&&K())try{await _e(s,n,me)}catch(i){console.warn("PTO: unable to save field",s,i)}}async function mn(e,t){O.permanents[e]=t;let n=gn(e);if(n!=null&&n.note&&K())try{await Excel.run(async a=>{var u;let o=a.workbook.tables.getItemOrNullObject(me[0]);if(await a.sync(),o.isNullObject)return;let s=o.getDataBodyRange(),i=o.getHeaderRowRange();s.load("values"),i.load("values"),await a.sync();let d=i.values[0]||[],r=d.map(c=>String(c||"").trim().toLowerCase()),l={field:r.findIndex(c=>c==="field"||c==="field name"||c==="setting"),permanent:r.findIndex(c=>c==="permanent"||c==="persist"),value:r.findIndex(c=>c==="value"||c==="setting value"),type:r.findIndex(c=>c==="type"||c==="category"),title:r.findIndex(c=>c==="title"||c==="display name")};if(l.field===-1)return;let f=(s.values||[]).findIndex(c=>String(c[l.field]||"").trim()===n.note);if(f>=0)l.permanent>=0&&(s.getCell(f,l.permanent).values=[[t?"Y":"N"]]);else{let c=new Array(d.length).fill("");l.type>=0&&(c[l.type]="Other"),l.title>=0&&(c[l.title]=""),c[l.field]=n.note,l.permanent>=0&&(c[l.permanent]=t?"Y":"N"),l.value>=0&&(c[l.value]=((u=O.steps[e])==null?void 0:u.notes)||""),o.rows.add(null,[c])}await a.sync()})}catch(a){console.warn("PTO: unable to update permanent flag",a)}}async function hn(e,t){let n=ln[e];if(n&&(O.completes[e]=t?"Y":"",!!K()))try{await _e(n,t?"Y":"",me)}catch(a){console.warn("PTO: unable to save completion flag",n,a)}}function en(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function ba(){let e={};return Object.keys(Ge).forEach(t=>{var s;let n=parseInt(t,10),a=!!((s=O.steps[n])!=null&&s.signOffDate),o=!!O.completes[n];e[n]=a||o}),e}function yn(e,{buttonId:t,inputId:n,canActivate:a=null,onComplete:o=null}){var r;let s=document.getElementById(t);if(!s)return;let i=document.getElementById(n),d=!!((r=O.steps[e])!=null&&r.signOffDate)||!!O.completes[e];en(s,d),s.addEventListener("click",()=>{if(!s.classList.contains("is-active")&&e>0){let f=ba(),{canComplete:u,message:c}=Jt(e,f);if(!u){zt(c);return}}if(typeof a=="function"&&!a())return;let p=!s.classList.contains("is-active");en(s,p),i&&(i.value=p?ut():"",te(e,"signOffDate",i.value)),hn(e,p),p&&typeof o=="function"&&o()})}function v(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function wa(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function vn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function Se(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function dt(e){if(!e)return null;let t=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(e));if(!t)return null;let n=Number(t[1]),a=Number(t[2]),o=Number(t[3]);return!n||!a||!o?null:{year:n,month:a,day:o}}function tn(e){if(!e)return"";let t=dt(e);if(!t)return"";let{year:n,month:a,day:o}=t;return`${n}-${String(a).padStart(2,"0")}-${String(o).padStart(2,"0")}`}function ka(e){let t=dt(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function Oa(e){let t=dt(e);return t?`PTO-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function ut(){let e=new Date,t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}function xa(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="y"||t==="yes"||t==="true"||t==="t"||t==="1"}function Sa(e){if(e instanceof Date)return e.getTime();if(typeof e=="number"){let n=Ea(e);return n?n.getTime():null}let t=new Date(e);return Number.isNaN(t.getTime())?null:t.getTime()}function Ea(e){if(!Number.isFinite(e))return null;let t=new Date(Date.UTC(1899,11,30));return new Date(t.getTime()+e*24*60*60*1e3)}function Ca(){let e=n=>{var a,o;return((o=(a=document.getElementById(n))==null?void 0:a.value)==null?void 0:o.trim())||""};[{id:"config-payroll-date",field:w.payrollDate},{id:"config-accounting-period",field:w.accountingPeriod},{id:"config-journal-id",field:w.journalEntryId},{id:"config-company-name",field:w.companyName},{id:"config-payroll-provider",field:w.payrollProvider},{id:"config-accounting-link",field:w.accountingSoftware},{id:"config-user-name",field:w.reviewerName}].forEach(({id:n,field:a})=>{let o=e(n);a&&re(a,o)})}function $(e,t=[]){let n=t.map(a=>z(a));return e.findIndex(a=>n.some(o=>a.includes(o)))}function _a(e){var E,R,C,j,L,B,W,F,T;let t=we(2),n=(t==null?void 0:t.notes)||"",a=!!O.permanents[2],o=be(t==null?void 0:t.reviewer),s=(t==null?void 0:t.signOffDate)||"",i=!!(Se(O.completes[2])||s),d=A.roster||{},r=A.hasAnalyzed,l=(R=(E=A.roster)==null?void 0:E.difference)!=null?R:0,p=!A.skipAnalysis&&Math.abs(l)>0,f=(C=d.rosterCount)!=null?C:0,u=(j=d.payrollCount)!=null?j:0,c=(L=d.difference)!=null?L:u-f,g=Array.isArray(d.mismatches)?d.mismatches.filter(Boolean):[],y="";A.loading?y=((W=(B=window.PrairieForge)==null?void 0:B.renderStatusBanner)==null?void 0:W.call(B,{type:"info",message:"Analyzing headcount\u2026",escapeHtml:v}))||"":A.lastError&&(y=((T=(F=window.PrairieForge)==null?void 0:F.renderStatusBanner)==null?void 0:T.call(F,{type:"error",message:A.lastError,escapeHtml:v}))||"");let b=(P,h,_,I)=>{let N=!r,q;N?q='<span class="pf-je-check-circle pf-je-circle--pending"></span>':I?q=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:q=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let de=r?` = ${_}`:"";return`
            <div class="pf-je-check-row">
                ${q}
                <span class="pf-je-check-desc-pill">${v(P)}${de}</span>
            </div>
        `},m=`
        ${b("SS_Employee_Roster count","Active employees in roster",f,!0)}
        ${b("PTO_Data count","Unique employees in PTO data",u,!0)}
        ${b("Difference","Should be zero",c,c===0)}
    `,S=g.length&&!A.skipAnalysis&&r?window.PrairieForge.renderMismatchTiles({mismatches:g,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"PTO Data",escapeHtml:v}):"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${v(he)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${v(e.title)}</h2>
            <p class="pf-hero-copy">${v(e.summary||"")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${A.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${Vt}
                    <span>Skip Analysis</span>
                </button>
            </div>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Headcount Check</h3>
                    <p class="pf-config-subtext">Compare employee roster against PTO data to identify discrepancies.</p>
                </div>
                <div class="pf-signoff-action">
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${He}</button>`,"Run")}
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${Ie}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${y}
                <div class="pf-je-checks-container">
                    ${m}
                </div>
                ${S}
            </article>
            ${pe({textareaId:"step-notes-input",value:n,permanentId:"step-notes-lock-2",isPermanent:a,hintId:p?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
            ${fe({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:i,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
        </section>
    `}function Ta(){let e=X.completenessCheck||{},t=X.missingPayRates||[],n=[{key:"accrualRate",label:"Accrual Rate",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"carryOver",label:"Carry Over",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdAccrued",label:"YTD Accrued",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdUsed",label:"YTD Used",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"balance",label:"Balance",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"}],o=n.every(l=>e[l.key]!==null&&e[l.key]!==void 0)&&n.every(l=>{var p;return(p=e[l.key])==null?void 0:p.match}),s=t.length>0,i=l=>{let p=e[l.key],f=p==null,u;return f?u='<span class="pf-je-check-circle pf-je-circle--pending"></span>':p.match?u=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:u=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${u}
                <span class="pf-je-check-desc-pill">${v(l.label)}: ${v(l.desc)}</span>
            </div>
        `},d=n.map(l=>i(l)).join(""),r="";if(s){let l=t[0],p=t.length-1;r=`
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${v(l.name)}</strong> to calculate liability
                </p>
                <div class="pf-readiness-input-row">
                    <div class="pf-readiness-input-field">
                        <span class="pf-readiness-input-prefix">$</span>
                        <input type="number" 
                               id="payrate-input" 
                               class="pf-readiness-input" 
                               placeholder="0.00" 
                               step="0.01"
                               min="0"
                               data-employee="${wa(l.name)}"
                               data-row="${l.rowIndex}">
                    </div>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--secondary" id="payrate-ignore-btn">
                        Skip
                    </button>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--primary" id="payrate-save-btn">
                        Save
                    </button>
                </div>
                ${p>0?`<p class="pf-readiness-remaining">${p} more employee${p>1?"s":""} need pay rates</p>`:""}
            </div>
        `}return`
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-readiness-card">
            <div class="pf-config-head">
                <h3>Data Completeness</h3>
                <p class="pf-config-subtext">Quick check that all your data transferred correctly.</p>
            </div>
            <div class="pf-je-checks-container">
                ${d}
            </div>
            ${r}
        </article>
    `}function Pa(e){var c,g,y,b,m,S,E,R;let t=we(3),n=!!O.permanents[3],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(O.completes[3])||o),i=Y.hasRun,{balanceIssues:d,zeroBalances:r,accrualOutliers:l,totalEmployees:p}=Y,f="";if(Y.loading)f=((g=(c=window.PrairieForge)==null?void 0:c.renderStatusBanner)==null?void 0:g.call(c,{type:"info",message:"Analyzing data quality...",escapeHtml:v}))||"";else if(i){let C=d.length,j=l.length+r.length;C>0?f=((b=(y=window.PrairieForge)==null?void 0:y.renderStatusBanner)==null?void 0:b.call(y,{type:"error",title:`${C} Balance Issue${C>1?"s":""} Found`,message:"Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",escapeHtml:v}))||"":j>0?f=((S=(m=window.PrairieForge)==null?void 0:m.renderStatusBanner)==null?void 0:S.call(m,{type:"warning",title:"No Critical Issues",message:`${j} informational item${j>1?"s":""} to review (see below).`,escapeHtml:v}))||"":f=((R=(E=window.PrairieForge)==null?void 0:E.renderStatusBanner)==null?void 0:R.call(E,{type:"success",title:"Data Quality Passed",message:`${p} employee${p!==1?"s":""} checked \u2014 no anomalies found.`,escapeHtml:v}))||""}let u=[];return i&&d.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u26A0\uFE0F</span>
                    <span class="pf-quality-issue-title">Balance Issues (${d.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${d.slice(0,5).map(C=>`<li><strong>${v(C.name)}</strong>: ${v(C.issue)}</li>`).join("")}
                    ${d.length>5?`<li class="pf-quality-more">+${d.length-5} more</li>`:""}
                </ul>
            </div>
        `),i&&l.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--warning">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u{1F4CA}</span>
                    <span class="pf-quality-issue-title">High Accrual Rates (${l.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${l.slice(0,5).map(C=>`<li><strong>${v(C.name)}</strong>: ${C.accrualRate.toFixed(2)} hrs/period</li>`).join("")}
                    ${l.length>5?`<li class="pf-quality-more">+${l.length-5} more</li>`:""}
                </ul>
            </div>
        `),i&&r.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--info">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u2139\uFE0F</span>
                    <span class="pf-quality-issue-title">Zero Balances (${r.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${r.slice(0,5).map(C=>`<li><strong>${v(C.name)}</strong></li>`).join("")}
                    ${r.length>5?`<li class="pf-quality-more">+${r.length-5} more</li>`:""}
                </ul>
            </div>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${v(he)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${v(e.title)}</h2>
            <p class="pf-hero-copy">${v(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${f}
                <div class="pf-signoff-action">
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${He}</button>`,"Run")}
                </div>
            </article>
            ${u.length>0?`
                <article class="pf-step-card pf-step-detail">
                    <div class="pf-config-head">
                        <h3>Issues Found</h3>
                        <p class="pf-config-subtext">Fix issues in PTO_Data and re-run, or acknowledge to continue.</p>
                    </div>
                    <div class="pf-quality-issues-grid">
                        ${u.join("")}
                    </div>
                    <div class="pf-quality-actions-bar">
                        ${Y.acknowledged?'<p class="pf-quality-actions-hint"><span class="pf-acknowledged-badge">\u2713 Issues Acknowledged</span></p>':""}
                        <div class="pf-signoff-action">
                            ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${Ie}</button>`,"Refresh")}
                            ${Y.acknowledged?"":G(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${Pe}</button>`,"Continue")}
                        </div>
                    </div>
                </article>
            `:""}
            ${pe({textareaId:"step-notes-3",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-3",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-3"})}
            ${fe({reviewerInputId:"step-reviewer-3",reviewerValue:a,signoffInputId:"step-signoff-3",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-3",completeButtonId:"step-signoff-toggle-3"})}
        </section>
    `}function Ia(e){let t=we(4),n=!!O.permanents[4],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(O.completes[4])||o);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${v(he)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${v(e.title)}</h2>
            <p class="pf-hero-copy">${v(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Analysis</h3>
                    <p class="pf-config-subtext">Calculate liabilities and compare against last period.</p>
                </div>
                <div class="pf-signoff-action">
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${He}</button>`,"Run")}
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${Ie}</button>`,"Refresh")}
                </div>
            </article>
            ${Ta()}
            ${pe({textareaId:"step-notes-4",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-4",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-4"})}
            ${fe({reviewerInputId:"step-reviewer-4",reviewerValue:a,signoffInputId:"step-signoff-4",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"step-signoff-toggle-4"})}
        </section>
    `}function Ra(e){let t=we(5),n=!!O.permanents[5],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(O.completes[5])||o),i=J.lastError?`<p class="pf-step-note">${v(J.lastError)}</p>`:"",d=J.validationRun,r=J.issues||[],l=[{key:"Debits = Credits",desc:"\u2211 Debit column = \u2211 Credit column"},{key:"Line Amounts Sum to Zero",desc:"\u2211 Line Amount = $0.00"},{key:"JE Matches Analysis Total",desc:"\u2211 Expense line amounts = \u2211 PTO_Analysis Change"}],p=g=>{let y=r.find(S=>S.check===g.key),b=!d,m;return b?m='<span class="pf-je-check-circle pf-je-circle--pending"></span>':y!=null&&y.passed?m=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:m=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${m}
                <span class="pf-je-check-desc-pill">${v(g.desc)}</span>
            </div>
        `},f=l.map(g=>p(g)).join(""),u=r.filter(g=>!g.passed),c="";return d&&u.length>0&&(c=`
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>\u26A0\uFE0F Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${u.map(g=>`<li><strong>${v(g.check)}:</strong> ${v(g.detail)}</li>`).join("")}
                </ul>
            </article>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${v(he)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${v(e.title)}</h2>
            <p class="pf-hero-copy">${v(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${Ve}</button>`,"Generate")}
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${Ie}</button>`,"Refresh")}
                    ${G(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${Mt}</button>`,"Export")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${i}
                <div class="pf-je-checks-container">
                    ${f}
                </div>
            </article>
            ${c}
            ${pe({textareaId:"step-notes-5",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-5",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-5"})}
            ${fe({reviewerInputId:"step-reviewer-5",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
        </section>
    `}function Aa(){var t,n;return Math.abs((n=(t=A.roster)==null?void 0:t.difference)!=null?n:0)>0}function bn(){return!A.skipAnalysis&&Aa()}async function ct(){if(!K()){A.loading=!1,A.lastError="Excel runtime is unavailable.",oe();return}A.loading=!0,A.lastError=null,ce(document.getElementById("headcount-save-btn"),!1),oe();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),a=t.workbook.worksheets.getItem("PTO_Data"),o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.getUsedRangeOrNullObject(),i=a.getUsedRangeOrNullObject();s.load("values"),i.load("values"),o.load("isNullObject"),await t.sync();let d=null;o.isNullObject||(d=o.getUsedRangeOrNullObject(),d.load("values")),await t.sync();let r=s.isNullObject?[]:s.values||[],l=i.isNullObject?[]:i.values||[],p=d&&!d.isNullObject?d.values||[]:[],f=p.length?p:l;return Na(r,f)});A.roster=e.roster,A.hasAnalyzed=!0,A.lastError=null}catch(e){console.warn("PTO headcount: unable to analyze data",e),A.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{A.loading=!1,oe()}}function nn(e){if(!e)return!0;let t=e.toLowerCase().trim();return t?["total","subtotal","sum","count","grand","average","avg"].some(a=>t.includes(a)):!0}function Na(e,t){let n={rosterCount:0,payrollCount:0,difference:0,mismatches:[]};if(((e==null?void 0:e.length)||0)<2||((t==null?void 0:t.length)||0)<2)return console.warn("Headcount: insufficient data rows",{rosterRows:(e==null?void 0:e.length)||0,payrollRows:(t==null?void 0:t.length)||0}),{roster:n};let a=an(e),o=an(t),s=a.headers,i=o.headers,d={employee:on(s),termination:s.findIndex(c=>c.includes("termination"))},r={employee:on(i)};console.log("Headcount column detection:",{rosterEmployeeCol:d.employee,rosterTerminationCol:d.termination,payrollEmployeeCol:r.employee,rosterHeaders:s.slice(0,5),payrollHeaders:i.slice(0,5)});let l=new Set,p=new Set;for(let c=a.startIndex;c<e.length;c+=1){let g=e[c],y=d.employee>=0?ve(g[d.employee]):"";nn(y)||d.termination>=0&&ve(g[d.termination])||l.add(y.toLowerCase())}for(let c=o.startIndex;c<t.length;c+=1){let g=t[c],y=r.employee>=0?ve(g[r.employee]):"";nn(y)||p.add(y.toLowerCase())}n.rosterCount=l.size,n.payrollCount=p.size,n.difference=n.payrollCount-n.rosterCount,console.log("Headcount results:",{rosterCount:n.rosterCount,payrollCount:n.payrollCount,difference:n.difference});let f=[...l].filter(c=>!p.has(c)),u=[...p].filter(c=>!l.has(c));return n.mismatches=[...f.map(c=>`In roster, missing in PTO_Data: ${c}`),...u.map(c=>`In PTO_Data, missing in roster: ${c}`)],{roster:n}}function an(e){if(!Array.isArray(e)||!e.length)return{headers:[],startIndex:1};let t=e.findIndex((o=[])=>o.some(s=>ve(s).toLowerCase().includes("employee"))),n=t===-1?0:t;return{headers:(e[n]||[]).map(o=>ve(o).toLowerCase()),startIndex:n+1}}function on(e=[]){let t=-1,n=-1;return e.forEach((a,o)=>{let s=a.toLowerCase();if(!s.includes("employee"))return;let i=1;s.includes("name")?i=4:s.includes("id")?i=2:i=3,i>n&&(n=i,t=o)}),t}function ve(e){return e==null?"":String(e).trim()}async function wn(e=null){let t=async n=>{let a=n.workbook.worksheets.getItem("PTO_Data"),o=n.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster"),i=n.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary"),d=n.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary"),r=a.getUsedRangeOrNullObject();r.load("values"),o.load("isNullObject"),s.load("isNullObject"),i.load("isNullObject"),d.load("isNullObject"),await n.sync();let l=r.isNullObject?[]:r.values||[];if(!l.length)return;let p=(l[0]||[]).map(k=>z(k)),f=p.findIndex(k=>k.includes("employee")&&k.includes("name")),u=f>=0?f:0,c=$(p,["accrual rate"]),g=$(p,["carry over","carryover"]),y=p.findIndex(k=>k.includes("ytd")&&(k.includes("accrued")||k.includes("accrual"))),b=p.findIndex(k=>k.includes("ytd")&&k.includes("used")),m=$(p,["balance","current balance","pto balance"]);console.log("[PTO Analysis] PTO_Data headers:",p),console.log("[PTO Analysis] Column indices found:",{employee:u,accrualRate:c,carryOver:g,ytdAccrued:y,ytdUsed:b,balance:m}),b>=0?console.log(`[PTO Analysis] YTD Used column: "${p[b]}" at index ${b}`):console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:",p);let S=l.slice(1).map(k=>ve(k[u])).filter(k=>k&&!k.toLowerCase().includes("total")),E=new Map;l.slice(1).forEach(k=>{let U=z(k[u]);!U||U.includes("total")||E.set(U,k)});let R=new Map;if(s.isNullObject)console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");else{let k=s.getUsedRangeOrNullObject();k.load("values"),await n.sync();let U=k.isNullObject?[]:k.values||[];if(U.length){let M=(U[0]||[]).map(x=>z(x));console.log("[PTO Analysis] SS_Employee_Roster headers:",M);let V=M.findIndex(x=>x.includes("employee")&&x.includes("name"));V<0&&(V=M.findIndex(x=>x==="employee"||x==="name"||x==="full name"));let H=M.findIndex(x=>x.includes("department"));console.log(`[PTO Analysis] Roster column indices - Name: ${V}, Dept: ${H}`),V>=0&&H>=0?(U.slice(1).forEach(x=>{let ie=z(x[V]),ue=ve(x[H]);ie&&R.set(ie,ue)}),console.log(`[PTO Analysis] Built roster map with ${R.size} employees`)):console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster")}}let C=new Map;if(!i.isNullObject){let k=i.getUsedRangeOrNullObject();k.load("values"),await n.sync();let U=k.isNullObject?[]:k.values||[];if(U.length){let M=(U[0]||[]).map(H=>z(H)),V={payrollDate:$(M,["payroll date"]),employee:$(M,["employee"]),category:$(M,["payroll category","category"]),amount:$(M,["amount","gross salary","gross_salary","earnings"])};V.employee>=0&&V.category>=0&&V.amount>=0&&U.slice(1).forEach(H=>{let x=z(H[V.employee]);if(!x)return;let ie=z(H[V.category]);if(!ie.includes("regular")||!ie.includes("earn"))return;let ue=Number(H[V.amount])||0;if(!ue)return;let Ee=Sa(H[V.payrollDate]),Ce=C.get(x);(!Ce||Ee!=null&&Ee>Ce.timestamp)&&C.set(x,{payRate:ue/80,timestamp:Ee})})}}let j=new Map;if(!d.isNullObject){let k=d.getUsedRangeOrNullObject();k.load("values"),await n.sync();let U=k.isNullObject?[]:k.values||[];if(U.length>1){let M=(U[0]||[]).map(x=>z(x)),V=M.findIndex(x=>x.includes("employee")&&x.includes("name")),H=$(M,["liability amount","liability","accrued pto"]);V>=0&&H>=0&&U.slice(1).forEach(x=>{let ie=z(x[V]);if(!ie)return;let ue=Number(x[H])||0;j.set(ie,ue)})}}let L=le(w.payrollDate)||"",B=[],W=[],F=S.map((k,U)=>{var gt,mt,ht,yt,vt,bt,wt;let M=z(k),V=R.get(M)||"",H=(mt=(gt=C.get(M))==null?void 0:gt.payRate)!=null?mt:"",x=E.get(M),ie=x&&c>=0&&(ht=x[c])!=null?ht:"",ue=x&&g>=0&&(yt=x[g])!=null?yt:"",Ee=x&&y>=0&&(vt=x[y])!=null?vt:"",Ce=x&&b>=0&&(bt=x[b])!=null?bt:"";(M.includes("avalos")||M.includes("sarah"))&&console.log(`[PTO Debug] ${k}:`,{ytdUsedIdx:b,rawValue:x?x[b]:"no dataRow",ytdUsed:Ce,fullRow:x});let Je=x&&m>=0&&Number(x[m])||0,pt=U+2;!H&&typeof H!="number"&&B.push({name:k,rowIndex:pt}),V||W.push({name:k,rowIndex:pt});let ze=typeof H=="number"&&Je?Je*H:0,ft=(wt=j.get(M))!=null?wt:0,kn=(typeof ze=="number"?ze:0)-ft;return[L,k,V,H,ie,ue,Ee,Ce,Je,ze,ft,kn]});X.missingPayRates=B.filter(k=>!X.ignoredMissingPayRates.has(k.name)),X.missingDepartments=W,console.log(`[PTO Analysis] Data quality: ${B.length} missing pay rates, ${W.length} missing departments`);let T=[["Analysis Date","Employee Name","Department","Pay Rate","Accrual Rate","Carry Over","YTD Accrued","YTD Used","Balance","Liability Amount","Accrued PTO $ [Prior Period]","Change"],...F],P=o.isNullObject?n.workbook.worksheets.add("PTO_Analysis"):o,h=P.getUsedRangeOrNullObject();h.load("address"),await n.sync(),h.isNullObject||h.clear();let _=T[0].length,I=T.length,N=F.length,q=P.getRangeByIndexes(0,0,I,_);q.values=T;let de=P.getRangeByIndexes(0,0,1,_);it(de),N>0&&(qt(P,0,N),ge(P,3,N),Oe(P,4,N),Oe(P,5,N),Oe(P,6,N),Oe(P,7,N),Oe(P,8,N),ge(P,9,N),ge(P,10,N),ge(P,11,N,!0)),q.format.autofitColumns(),P.getRange("A1").select(),await n.sync()};K()&&(e?await t(e):await Excel.run(t))}function Da(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let a=String(n);return/[",\n]/.test(a)?`"${a.replace(/"/g,'""')}"`:a}).join(",")).join(`
`)}function $a(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),a=URL.createObjectURL(n),o=document.createElement("a");o.href=a,o.download=e,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(a),1e3)}function sn(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=bn(),n=document.getElementById("step-notes-input"),a=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!a;let o=document.getElementById("headcount-notes-hint");o&&(o.textContent=t?"Please document outstanding differences before signing off.":"")}function rn(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(ye)?t.slice(ye.length).replace(/^\s+/,""):t.replace(new RegExp(`^${ye}\\s*`,"i"),"").trimStart(),a=ye+(n?`
${n}`:"");e.value!==a&&(e.value=a),te(2,"notes",e.value)}function ja(){let e=document.getElementById("step-notes-input");e&&e.addEventListener("input",()=>{if(!A.skipAnalysis)return;let t=e.value||"";if(!t.startsWith(ye)){let n=t.replace(ye,"").trimStart();e.value=ye+(n?`
${n}`:"")}te(2,"notes",e.value)})}})();
//# sourceMappingURL=app.bundle.js.map
