/* Prairie Forge PTO Accrual */
(()=>{function W(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}var Ke="SS_PF_Config";async function Et(e,t=[Ke]){var o;let n=e.workbook.tables;n.load("items/name"),await e.sync();let a=(o=n.items)==null?void 0:o.find(s=>t.includes(s.name));return a?e.workbook.tables.getItem(a.name):(console.warn("Config table not found. Looking for:",t),null)}function Ct(e){let t=e.map(n=>String(n||"").trim().toLowerCase());return{field:t.findIndex(n=>n==="field"||n==="field name"||n==="setting"),value:t.findIndex(n=>n==="value"||n==="setting value"),type:t.findIndex(n=>n==="type"||n==="category"),title:t.findIndex(n=>n==="title"||n==="display name"),permanent:t.findIndex(n=>n==="permanent"||n==="persist")}}async function Qe(e=[Ke]){if(!W())return{};try{return await Excel.run(async t=>{let n=await Et(t,e);if(!n)return{};let a=n.getDataBodyRange(),o=n.getHeaderRowRange();a.load("values"),o.load("values"),await t.sync();let s=o.values[0]||[],l=Ct(s);if(l.field===-1||l.value===-1)return console.warn("Config table missing FIELD or VALUE columns. Headers:",s),{};let i={};return(a.values||[]).forEach(c=>{var f;let p=String(c[l.field]||"").trim();p&&(i[p]=(f=c[l.value])!=null?f:"")}),console.log("Configuration loaded:",Object.keys(i).length,"fields"),i})}catch(t){return console.error("Failed to load configuration:",t),{}}}async function Pe(e,t,n=[Ke]){if(!W())return!1;try{return await Excel.run(async a=>{let o=await Et(a,n);if(!o){console.warn("Config table not found for write");return}let s=o.getDataBodyRange(),l=o.getHeaderRowRange();s.load("values"),l.load("values"),await a.sync();let i=l.values[0]||[],r=Ct(i);if(r.field===-1||r.value===-1){console.error("Config table missing FIELD or VALUE columns");return}let p=(s.values||[]).findIndex(f=>String(f[r.field]||"").trim()===e);if(p>=0)s.getCell(p,r.value).values=[[t]];else{let f=new Array(i.length).fill("");r.type>=0&&(f[r.type]="Run Settings"),f[r.field]=e,f[r.value]=t,r.permanent>=0&&(f[r.permanent]="N"),r.title>=0&&(f[r.title]=""),o.rows.add(null,[f]),console.log("Added new config row:",e,"=",t)}await a.sync(),console.log("Saved config:",e,"=",t)}),!0}catch(a){return console.error("Failed to save config:",e,a),!1}}var Sn="SS_PF_Config",xn="module-prefix",Xe="system",ke={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function _t(){if(!W())return{...ke};try{return await Excel.run(async e=>{var p,f;let t=e.workbook.worksheets.getItemOrNullObject(Sn);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...ke};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((p=n.values)!=null&&p.length))return{...ke};let a=n.values,o=_n(a[0]),s=o.get("category"),l=o.get("field"),i=o.get("value");if(s===void 0||l===void 0||i===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...ke};let r={},c=!1;for(let u=1;u<a.length;u++){let d=a[u];if(je(d[s])===xn){let y=String((f=d[l])!=null?f:"").trim().toUpperCase(),w=je(d[i]);y&&w&&(r[y]=w,c=!0)}}return c?(console.log("[Tab Visibility] Loaded prefix config:",r),r):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...ke})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...ke}}}async function Ze(e){if(!W())return;let t=je(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await _t();await Excel.run(async a=>{let o=a.workbook.worksheets;o.load("items/name,visibility"),await a.sync();let s={};for(let[u,d]of Object.entries(n))s[d]||(s[d]=[]),s[d].push(u);let l=s[t]||[],i=s[Xe]||[],r=[];for(let[u,d]of Object.entries(s))u!==t&&u!==Xe&&r.push(...d);console.log(`[Tab Visibility] Active prefixes: ${l.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${r.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${i.join(", ")}`);let c=[],p=[];o.items.forEach(u=>{let d=u.name,g=d.toUpperCase(),y=l.some(E=>g.startsWith(E)),w=r.some(E=>g.startsWith(E)),h=i.some(E=>g.startsWith(E));y?(c.push(u),console.log(`[Tab Visibility] SHOW: ${d} (matches active module prefix)`)):h?(p.push(u),console.log(`[Tab Visibility] HIDE: ${d} (system sheet)`)):w?(p.push(u),console.log(`[Tab Visibility] HIDE: ${d} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${d} (no prefix match, leaving as-is)`)});for(let u of c)u.visibility=Excel.SheetVisibility.visible;if(await a.sync(),o.items.filter(u=>u.visibility===Excel.SheetVisibility.visible).length>p.length){for(let u of p)try{u.visibility=Excel.SheetVisibility.hidden}catch(d){console.warn(`[Tab Visibility] Could not hide "${u.name}":`,d.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${c.length}, hid ${p.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function En(){if(!W()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.visibility!==Excel.SheetVisibility.visible&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${a.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function Cn(){if(!W()){console.log("Excel not available");return}try{let e=await _t(),t=[];for(let[n,a]of Object.entries(e))a===Xe&&t.push(n);await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name,visibility"),await n.sync(),a.items.forEach(o=>{let s=o.name.toUpperCase();t.some(l=>s.startsWith(l))&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function _n(e=[]){let t=new Map;return e.forEach((n,a)=>{let o=je(n);o&&t.set(o,a)}),t}function je(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=En,window.PrairieForge.unhideSystemSheets=Cn,window.PrairieForge.applyModuleTabVisibility=Ze);var Pt={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var Rt=Pt.ADA_IMAGE_URL;async function Le(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async a=>{let o=a.workbook.worksheets.getItemOrNullObject(e);o.load("isNullObject, name, visibility"),await a.sync();let s;o.isNullObject?(s=a.workbook.worksheets.add(e),await a.sync(),await Tt(a,s,t,n)):(s=o,s.visibility!==Excel.SheetVisibility.visible&&(s.visibility=Excel.SheetVisibility.visible,await a.sync()),await Tt(a,s,t,n)),s.activate(),s.getRange("A1").select(),await a.sync()})}catch(a){console.error(`Error activating homepage sheet ${e}:`,a)}}async function Tt(e,t,n,a){try{let c=t.getUsedRangeOrNullObject();c.load("isNullObject"),await e.sync(),c.isNullObject||(c.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let o=[[n,""],[a,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=o;let l=t.getRange("A1:Z100");l.format.fill.color="#0f0f0f";let i=t.getRange("A1");i.format.font.bold=!0,i.format.font.size=36,i.format.font.color="#ffffff",i.format.font.name="Segoe UI Light",i.format.verticalAlignment="Center";let r=t.getRange("A2");r.format.font.size=14,r.format.font.color="#a0a0a0",r.format.font.name="Segoe UI",r.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var It={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function Be(e){return It[e]||It["module-selector"]}function At(){tt();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${Rt}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",Pn),e}function tt(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function Pn(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${Rt}" alt="Ada" />
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",et),t.addEventListener("click",o=>{o.target===t&&et()});let a=o=>{o.key==="Escape"&&(et(),document.removeEventListener("keydown",a))};document.addEventListener("keydown",a)}function et(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var Tn=["January","February","March","April","May","June","July","August","September","October","November","December"],Dt=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],In=["Su","Mo","Tu","We","Th","Fr","Sa"],X=null,se=null;function $t(e,t={}){let n=document.getElementById(e);if(!n)return;let{onChange:a=null,minDate:o=null,maxDate:s=null,readonly:l=!1}=t,i=n.closest(".pf-datepicker-wrapper");i||(i=document.createElement("div"),i.className="pf-datepicker-wrapper",n.parentNode.insertBefore(i,n),i.appendChild(n)),n.type="text",n.placeholder="Select date...",n.classList.add("pf-datepicker-input"),n.readOnly=!0;let r=n.value?Nt(n.value):null;r&&(n.value=at(r),n.dataset.value=Te(r));let c=i.querySelector(".pf-datepicker-icon");c||(c=document.createElement("span"),c.className="pf-datepicker-icon",c.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>',i.appendChild(c));let p={inputId:e,input:n,selectedDate:r,viewDate:r?new Date(r):new Date,onChange:a,minDate:o,maxDate:s};function f(){l||(se=p,Rn())}return n.addEventListener("click",f),c.addEventListener("click",f),{getValue:()=>p.selectedDate?Te(p.selectedDate):"",setValue:u=>{let d=Nt(u);p.selectedDate=d,p.viewDate=d?new Date(d):new Date,d?(n.value=at(d),n.dataset.value=Te(d)):(n.value="",n.dataset.value="")},open:f,close:Me}}function Rn(){se&&(X||(X=document.createElement("div"),X.className="pf-datepicker-modal",X.id="pf-datepicker-modal",document.body.appendChild(X)),Lt(),requestAnimationFrame(()=>{X.classList.add("is-open")}),document.addEventListener("keydown",jt))}function Me(){X&&X.classList.remove("is-open"),document.removeEventListener("keydown",jt),se=null}function jt(e){e.key==="Escape"&&Me()}function Lt(){if(!X||!se)return;let{viewDate:e,selectedDate:t,minDate:n,maxDate:a}=se,o=e.getFullYear(),s=e.getMonth();X.innerHTML=`
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
                ${In.map(l=>`<span>${l}</span>`).join("")}
            </div>
            <div class="pf-datepicker-days">
                ${Dn(o,s,t,n,a)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-btn pf-datepicker-today" data-action="today">Today</button>
                <button type="button" class="pf-datepicker-btn pf-datepicker-clear" data-action="clear">Clear</button>
            </div>
        </div>
    `,An()}function An(){var e;X&&((e=X.querySelector(".pf-datepicker-backdrop"))==null||e.addEventListener("click",Me),X.querySelectorAll(".pf-datepicker-nav").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=t.dataset.action;Nn(a)})}),X.querySelectorAll(".pf-datepicker-day:not(.disabled)").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=parseInt(t.dataset.day),o=parseInt(t.dataset.month),s=parseInt(t.dataset.year);nt(new Date(s,o,a))})}),X.querySelectorAll(".pf-datepicker-btn").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=t.dataset.action;a==="today"?nt(new Date):a==="clear"&&nt(null)})}))}function Nn(e){if(!se)return;let t=se.viewDate;switch(e){case"prev-year":t.setFullYear(t.getFullYear()-1);break;case"prev-month":t.setMonth(t.getMonth()-1);break;case"next-month":t.setMonth(t.getMonth()+1);break;case"next-year":t.setFullYear(t.getFullYear()+1);break}Lt()}function nt(e){if(!se)return;let{input:t,onChange:n}=se;se.selectedDate=e,e?(t.value=at(e),t.dataset.value=Te(e),se.viewDate=new Date(e)):(t.value="",t.dataset.value=""),n&&n(e?Te(e):""),t.dispatchEvent(new Event("change",{bubbles:!0})),Me()}function Dn(e,t,n,a,o){let s=new Date(e,t,1).getDay(),l=new Date(e,t+1,0).getDate(),i=new Date(e,t,0).getDate(),r=new Date;r.setHours(0,0,0,0),n&&(n=new Date(n),n.setHours(0,0,0,0));let c="";for(let d=s-1;d>=0;d--){let g=i-d,y=t===0?11:t-1,w=t===0?e-1:e;c+=`<button type="button" class="pf-datepicker-day other-month" data-day="${g}" data-month="${y}" data-year="${w}">${g}</button>`}for(let d=1;d<=l;d++){let g=new Date(e,t,d);g.setHours(0,0,0,0);let y=g.getTime()===r.getTime(),w=n&&g.getTime()===n.getTime(),h="pf-datepicker-day";y&&(h+=" today"),w&&(h+=" selected");let E=!1;a&&g<a&&(E=!0),o&&g>o&&(E=!0),E&&(h+=" disabled"),c+=`<button type="button" class="${h}" data-day="${d}" data-month="${t}" data-year="${e}" ${E?"disabled":""}>${d}</button>`}let p=42,f=s+l,u=p-f;for(let d=1;d<=u;d++){let g=t===11?0:t+1,y=t===11?e+1:e;c+=`<button type="button" class="pf-datepicker-day other-month" data-day="${d}" data-month="${g}" data-year="${y}">${d}</button>`}return c}function Nt(e){if(!e)return null;if(/^\d{4}-\d{2}-\d{2}$/.test(e)){let[a,o,s]=e.split("-").map(Number);return new Date(a,o-1,s)}let t=e.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);if(t){let a=Dt.findIndex(o=>o.toLowerCase()===t[1].toLowerCase().substring(0,3));if(a>=0)return new Date(parseInt(t[3]),a,parseInt(t[2]))}if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(e)){let[a,o,s]=e.split("/").map(Number);return new Date(s,a-1,o)}let n=new Date(e);return isNaN(n.getTime())?null:n}function at(e){return e?`${Dt[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`:""}function Te(e){if(!e)return"";let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}var Bt=`
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
`.trim(),Mt=`
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
`.trim(),Vt=`
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
`.trim(),oo=`
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
`.trim(),so=`
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
`.trim(),$n={config:`
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
    `};function Ht(e){return e&&$n[e]||""}var ot=`
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
`.trim(),st=`
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
`.trim(),io=`
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
`.trim(),it=`
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
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
        <path d="M7 10l5 5 5-5" />
        <path d="M12 15V3" />
    </svg>
`.trim(),Ut=`
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
        <circle cx="12" cy="12" r="10" />
        <path d="m15 9-6 6" />
        <path d="m9 9 6 6" />
    </svg>
`.trim(),Fe=`
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
`.trim(),Jt=`
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
`.trim(),Ue=`
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
`.trim(),ro=`
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
`.trim(),lo=`
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
`.trim(),co=`
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
`.trim(),uo=`
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
`.trim(),po=`
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
`.trim(),fo=`
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
`.trim(),go=`
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
`.trim(),mo=`
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
`.trim(),Re=`
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
`.trim(),zt=`
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
`.trim();function Ae(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function M(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function ge({textareaId:e,value:t,permanentId:n,isPermanent:a,hintId:o,saveButtonId:s,isSaved:l=!1,placeholder:i="Enter notes here..."}){let r=a?st:ot,c=s?`<button type="button" class="pf-action-toggle pf-save-btn ${l?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Jt}</button>`:"",p=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${a?"is-locked":""}" id="${n}" aria-pressed="${a}" title="Lock notes (retain after archive)">${r}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Ae(i)}">${Ae(t||"")}</textarea>
                ${o?`<p class="pf-signoff-hint" id="${o}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?M(p,"Lock"):""}
                ${s?M(c,"Save"):""}
            </div>
        </article>
    `}function me({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:a,isComplete:o,saveButtonId:s,isSaved:l=!1,completeButtonId:i,subtext:r="Sign-off below. Click checkmark icon. Done.",prevButtonId:c=null,nextButtonId:p=null}){let f=c||`${i}-prev`,u=p||`${i}-next`,d=`<button type="button" class="pf-action-toggle ${o?"is-active":""}" id="${i}" aria-pressed="${!!o}" title="Mark step complete">${Ie}</button>`,g=`<button type="button" class="pf-action-toggle pf-nav-toggle" id="${f}" title="Previous step">${Fe}</button>`,y=`<button type="button" class="pf-action-toggle pf-nav-toggle" id="${u}" title="Next step">${Ue}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${Ae(r)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${Ae(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${Ae(a)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${g?M(g,"Prev"):""}
                ${M(d,"Done")}
                ${y?M(y,"Next"):""}
            </div>
        </article>
    `}function rt(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?st:ot)}function de(e,t){e&&e.classList.toggle("is-saved",t)}function lt(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(a=>{let o=a.getAttribute("data-save-input"),s=document.getElementById(o);if(!s)return;let l=()=>{de(a,!1)};s.addEventListener("input",l),n.push(()=>s.removeEventListener("input",l))}),()=>n.forEach(a=>a())}function qt(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function Yt(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var ct={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},Ge={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function dt(e){e.format.fill.color=ct.fillColor,e.format.font.color=ct.fontColor,e.format.font.bold=ct.bold}function he(e,t,n,a=!1){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a?Ge.currencyWithNegative:Ge.currency]]}function Oe(e,t,n){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[Ge.number]]}function Wt(e,t,n,a=Ge.date){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a]]}var jn="cf5858f",De="pto-accrual";var ue="PTO Accrual";function z(e,t="info",n=4e3){document.querySelectorAll(".pf-toast").forEach(o=>o.remove());let a=document.createElement("div");if(a.className=`pf-toast pf-toast--${t}`,a.innerHTML=`
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
        `,document.head.appendChild(o)}return document.body.appendChild(a),n>0&&setTimeout(()=>a.remove(),n),a}function Ln(e,t={}){let{title:n="Confirm Action",confirmText:a="Continue",cancelText:o="Cancel",icon:s="\u{1F4CB}",destructive:l=!1}=t;return new Promise(i=>{document.querySelectorAll(".pf-confirm-overlay").forEach(c=>c.remove());let r=document.createElement("div");if(r.className="pf-confirm-overlay",r.innerHTML=`
            <div class="pf-confirm-dialog">
                <div class="pf-confirm-icon">${s}</div>
                <div class="pf-confirm-title">${n}</div>
                <div class="pf-confirm-message">${e.replace(/\n/g,"<br>")}</div>
                <div class="pf-confirm-buttons">
                    <button class="pf-confirm-btn pf-confirm-btn--cancel">${o}</button>
                    <button class="pf-confirm-btn pf-confirm-btn--ok ${l?"pf-confirm-btn--destructive":""}">${a}</button>
                </div>
            </div>
        `,!document.getElementById("pf-confirm-styles")){let c=document.createElement("style");c.id="pf-confirm-styles",c.textContent=`
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
            `,document.head.appendChild(c)}document.body.appendChild(r),r.addEventListener("click",c=>{c.target===r&&(r.remove(),i(!1))}),r.querySelector(".pf-confirm-btn--cancel").onclick=()=>{r.remove(),i(!1)},r.querySelector(".pf-confirm-btn--ok").onclick=()=>{r.remove(),i(!0)}})}var Bn="Calculate your PTO liability, compare against last period, and generate a balanced journal entry\u2014all without leaving Excel.",Mn="../module-selector/index.html",Vn="pf-loader-overlay",ce=["SS_PF_Config"],k={payrollProvider:"PTO_Payroll_Provider",payrollDate:"PTO_Analysis_Date",accountingPeriod:"PTO_Accounting_Period",journalEntryId:"PTO_Journal_Entry_ID",companyName:"SS_Company_Name",accountingSoftware:"SS_Accounting_Software",reviewerName:"PTO_Reviewer",validationDataBalance:"PTO_Validation_Data_Balance",validationCleanBalance:"PTO_Validation_Clean_Balance",validationDifference:"PTO_Validation_Difference",headcountRosterCount:"PTO_Headcount_Roster_Count",headcountPayrollCount:"PTO_Headcount_Payroll_Count",headcountDifference:"PTO_Headcount_Difference",journalDebitTotal:"PTO_JE_Debit_Total",journalCreditTotal:"PTO_JE_Credit_Total",journalDifference:"PTO_JE_Difference"},ye="User opted to skip the headcount review this period.",qe={0:{note:"PTO_Notes_Config",reviewer:"PTO_Reviewer_Config",signOff:"PTO_SignOff_Config"},1:{note:"PTO_Notes_Import",reviewer:"PTO_Reviewer_Import",signOff:"PTO_SignOff_Import"},2:{note:"PTO_Notes_Headcount",reviewer:"PTO_Reviewer_Headcount",signOff:"PTO_SignOff_Headcount"},3:{note:"PTO_Notes_Validate",reviewer:"PTO_Reviewer_Validate",signOff:"PTO_SignOff_Validate"},4:{note:"PTO_Notes_Review",reviewer:"PTO_Reviewer_Review",signOff:"PTO_SignOff_Review"},5:{note:"PTO_Notes_JE",reviewer:"PTO_Reviewer_JE",signOff:"PTO_SignOff_JE"},6:{note:"PTO_Notes_Archive",reviewer:"PTO_Reviewer_Archive",signOff:"PTO_SignOff_Archive"}},ln={0:"PTO_Complete_Config",1:"PTO_Complete_Import",2:"PTO_Complete_Headcount",3:"PTO_Complete_Validate",4:"PTO_Complete_Review",5:"PTO_Complete_JE",6:"PTO_Complete_Archive"};var ae=[{id:0,title:"Configuration",summary:"Set the analysis date, accounting period, and review details for this run.",description:"Complete this step first to ensure all downstream calculations use the correct period settings.",actionLabel:"Configure Workbook",secondaryAction:{sheet:"SS_PF_Config",label:"Open Config Sheet"}},{id:1,title:"Import PTO Data",summary:"Pull your latest PTO export from payroll and paste it into PTO_Data.",description:"Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",actionLabel:"Import Sample Data",secondaryAction:{sheet:"PTO_Data",label:"Open Data Sheet"}},{id:2,title:"Headcount Review",summary:"Quick check to make sure your roster matches your PTO data.",description:"Compare employees in PTO_Data against your employee roster to catch any discrepancies.",actionLabel:"Open Headcount Review",secondaryAction:{sheet:"SS_Employee_Roster",label:"Open Sheet"}},{id:3,title:"Data Quality Review",summary:"Scan your PTO data for potential errors before crunching numbers.",description:"Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",actionLabel:"Click to Run Quality Check"},{id:4,title:"PTO Accrual Review",summary:"Review the calculated liability for each employee and compare to last period.",description:"The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",actionLabel:"Click to Perform Review"},{id:5,title:"Journal Entry Prep",summary:"Generate a balanced journal entry, run validation checks, and export when ready.",description:"Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",actionLabel:"Open Journal Draft",secondaryAction:{sheet:"PTO_JE_Draft",label:"Open Sheet"}},{id:6,title:"Archive & Reset",summary:"Save this period's results and prepare for the next cycle.",description:"Archive the current analysis so it becomes the 'prior period' for your next review.",actionLabel:"Archive Run"}],Hn={0:"PTO_Homepage",1:"PTO_Data",2:"PTO_Data",3:"PTO_Analysis",4:"PTO_Analysis",5:"PTO_JE_Draft"},Fn={PTO_Homepage:0,PTO_Data:1,PTO_Analysis:4,PTO_JE_Draft:5,PTO_Archive_Summary:6,SS_PF_Config:0,SS_Employee_Roster:2};var Un=ae.reduce((e,t)=>(e[t.id]="pending",e),{}),R={activeView:"home",activeStepId:null,focusedIndex:0,stepStatuses:Un},O={loaded:!1,steps:{},permanents:{},completes:{},values:{},overrides:{accountingPeriod:!1,journalId:!1}},Ne=null,ut=null,Je=null,Se=new Map,A={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},J={debitTotal:null,creditTotal:null,difference:null,lineAmountSum:null,analysisChangeTotal:null,jeChangeTotal:null,loading:!1,lastError:null,validationRun:!1,issues:[]},K={hasRun:!1,loading:!1,acknowledged:!1,balanceIssues:[],zeroBalances:[],accrualOutliers:[],totalIssues:0,totalEmployees:0},ee={cleanDataReady:!1,employeeCount:0,lastRun:null,loading:!1,lastError:null,missingPayRates:[],missingDepartments:[],ignoredMissingPayRates:new Set,completenessCheck:{accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null}};async function Gn(){var e;try{Ne=document.getElementById("app"),ut=document.getElementById("loading"),await qn(),await Yn(),(e=window.PrairieForge)!=null&&e.loadSharedConfig&&await window.PrairieForge.loadSharedConfig();let t=Be(De);await Le(t.sheetName,t.title,t.subtitle),await Jn(),ut&&ut.remove(),Ne&&(Ne.hidden=!1),ie()}catch(t){throw console.error("[PTO] Module initialization failed:",t),t}}async function Jn(){if(oe())try{await Excel.run(async e=>{e.workbook.worksheets.onActivated.add(zn),await e.sync(),console.log("[PTO] Worksheet change listener registered")})}catch(e){console.warn("[PTO] Could not set up worksheet listener:",e)}}async function zn(e){try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e.worksheetId);n.load("name"),await t.sync();let a=n.name,o=Fn[a];if(console.log(`[PTO] Tab changed to: ${a} \u2192 Step ${o}`),o!==void 0&&o!==R.activeStepId){let s=STEPS.findIndex(l=>l.id===o);if(s>=0){let l=o===0?"config":"step";R.activeView=l,R.activeStepId=o,R.focusedIndex=s,ie()}}})}catch(t){console.warn("[PTO] Error handling worksheet change:",t)}}async function qn(){try{await Ze(De),console.log(`[PTO] Tab visibility applied for ${De}`)}catch(e){console.warn("[PTO] Could not apply tab visibility:",e)}}async function Yn(){var e;if(!W()){O.loaded=!0;return}try{let t=await Qe(ce),n={};(e=window.PrairieForge)!=null&&e.loadSharedConfig&&(await window.PrairieForge.loadSharedConfig(),window.PrairieForge._sharedConfigCache&&window.PrairieForge._sharedConfigCache.forEach((s,l)=>{n[l]=s}));let a={...t},o={SS_Default_Reviewer:k.reviewerName,Default_Reviewer:k.reviewerName,PTO_Reviewer:k.reviewerName,SS_Company_Name:k.companyName,Company_Name:k.companyName,SS_Payroll_Provider:k.payrollProvider,Payroll_Provider_Link:k.payrollProvider,SS_Accounting_Software:k.accountingSoftware,Accounting_Software_Link:k.accountingSoftware};Object.entries(o).forEach(([s,l])=>{n[s]&&!a[l]&&(a[l]=n[s])}),Object.entries(n).forEach(([s,l])=>{s.startsWith("PTO_")&&l&&(a[s]=l)}),O.permanents=await Wn(),O.values=a||{},O.overrides.accountingPeriod=!!(a!=null&&a[k.accountingPeriod]),O.overrides.journalId=!!(a!=null&&a[k.journalEntryId]),Object.entries(qe).forEach(([s,l])=>{var i,r,c;O.steps[s]={notes:(i=a[l.note])!=null?i:"",reviewer:(r=a[l.reviewer])!=null?r:"",signOffDate:(c=a[l.signOff])!=null?c:""}}),O.completes=Object.entries(ln).reduce((s,[l,i])=>{var r;return s[l]=(r=a[i])!=null?r:"",s},{}),O.loaded=!0}catch(t){console.warn("PTO: unable to load configuration fields",t),O.loaded=!0}}async function Wn(){let e={};if(!W())return e;let t=new Map;Object.entries(qe).forEach(([n,a])=>{a.note&&t.set(a.note.trim(),Number(n))});try{await Excel.run(async n=>{let a=n.workbook.tables.getItemOrNullObject(ce[0]);if(await n.sync(),a.isNullObject)return;let o=a.getDataBodyRange(),s=a.getHeaderRowRange();o.load("values"),s.load("values"),await n.sync();let i=(s.values[0]||[]).map(c=>String(c||"").trim().toLowerCase()),r={field:i.findIndex(c=>c==="field"||c==="field name"||c==="setting"),permanent:i.findIndex(c=>c==="permanent"||c==="persist")};r.field===-1||r.permanent===-1||(o.values||[]).forEach(c=>{let p=String(c[r.field]||"").trim(),f=t.get(p);if(f==null)return;let u=ja(c[r.permanent]);e[f]=u})})}catch(n){console.warn("PTO: unable to load permanent flags",n)}return e}function ie(){var i;if(!Ne)return;let e=R.focusedIndex<=0?"disabled":"",t=R.focusedIndex>=ae.length-1?"disabled":"",n=R.activeView==="step"&&R.activeStepId!=null,o=R.activeView==="config"?cn():n?na(R.activeStepId):`${Qn()}${Xn()}`;Ne.innerHTML=`
        <div class="pf-root">
            <div class="pf-brand-float" aria-hidden="true">
                <span class="pf-brand-wave"></span>
            </div>
            <header class="pf-banner">
                <div class="pf-nav-bar">
                    <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                        ${Fe}
                        <span class="sr-only">Previous step</span>
                    </button>
                    <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                        ${Bt}
                        <span class="sr-only">Module Home</span>
                    </button>
                    <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                        ${Mt}
                        <span class="sr-only">Module Selector</span>
                    </button>
                    <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                        ${Ue}
                        <span class="sr-only">Next step</span>
                    </button>
                    <span class="pf-nav-divider"></span>
                    <div class="pf-quick-access-wrapper">
                        <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                            ${Vt}
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
                    <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${jn}</div>
                </div>
            </footer>
        </div>
    `;let s=R.activeView==="home"||R.activeView!=="step"&&R.activeView!=="config",l=document.getElementById("pf-info-fab-pto");if(s)l&&l.remove();else if((i=window.PrairieForge)!=null&&i.mountInfoFab){let r=Kn(R.activeStepId);PrairieForge.mountInfoFab({title:r.title,content:r.content,buttonId:"pf-info-fab-pto"})}aa(),ra(),s?At():tt()}function Kn(e){switch(e){case 0:return{title:"Configuration",content:`
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
                `}}}function Qn(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">PTO Accrual</h2>
            <p class="pf-hero-copy">${Bn}</p>
        </section>
    `}function Xn(){return`
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${ae.map((e,t)=>Zn(e,t)).join("")}
            </div>
        </section>
    `}function Zn(e,t){let n=R.stepStatuses[e.id]||"pending",a=R.activeView==="step"&&R.focusedIndex===t?"pf-step-card--active":"",o=Ht(Ra(e.id));return`
        <article class="pf-step-card pf-clickable ${a}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${e.title}</h3>
        </article>
    `}function ea(e){let t=ae.filter(o=>o.id!==6).map(o=>({id:o.id,title:o.title,complete:la(o.id)})),n=t.every(o=>o.complete),a=t.map(o=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${o.complete?"is-active":""}" aria-pressed="${o.complete}">
                        ${Ie}
                    </span>
                    <div>
                        <h3>${b(o.title)}</h3>
                        <p class="pf-config-subtext">${o.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${b(e.title)}</h2>
            <p class="pf-hero-copy">${b(e.summary||"")}</p>
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
        `;let e=tn(te(k.payrollDate)),t=tn(te(k.accountingPeriod)),n=te(k.journalEntryId),a=te(k.accountingSoftware),o=te(k.payrollProvider),s=te(k.companyName),l=te(k.reviewerName),i=we(0),r=!!O.permanents[0],c=!!(bn(O.completes[0])||i.signOffDate),p=be(i==null?void 0:i.reviewer),f=(i==null?void 0:i.signOffDate)||"";return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${b(ue)} | Step 0</p>
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
                        <input type="text" id="config-user-name" value="${b(l)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${b(e)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${b(t)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${b(n)}" placeholder="PTO-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${b(s)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${b(o)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${b(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${ge({textareaId:"config-notes",value:i.notes||"",permanentId:"config-notes-lock",isPermanent:r,hintId:"",saveButtonId:"config-notes-save"})}
            ${me({reviewerInputId:"config-reviewer",reviewerValue:p,signoffInputId:"config-signoff-date",signoffValue:f,isComplete:c,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"})}
        </section>
    `}function ta(e){let t=we(1),n=!!O.permanents[1],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Ee(O.completes[1])||o),l=te(k.payrollProvider);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${b(e.title)}</h2>
            <p class="pf-hero-copy">${b(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export, then paste into PTO_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${M(l?`<a href="${b(l)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${it}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${it}</button>`,"Provider")}
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${Ve}</button>`,"PTO_Data")}
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${zt}</button>`,"Clear")}
                </div>
            </article>
            ${ge({textareaId:"step-notes-1",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-1",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-1"})}
            ${me({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
        </section>
    `}function na(e){let t=ae.find(i=>i.id===e);if(!t)return"";if(e===0)return cn();if(e===1)return ta(t);if(e===2)return Va(t);if(e===3)return Fa(t);if(e===4)return Ua(t);if(e===5)return Ga(t);if(t.id===6)return ea(t);let n=we(e),a=!!O.permanents[e],o=be(n==null?void 0:n.reviewer),s=(n==null?void 0:n.signOffDate)||"",l=!!(Ee(O.completes[e])||s);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(ue)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${b(t.title)}</h2>
            <p class="pf-hero-copy">${b(t.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${ge({textareaId:`step-notes-${e}`,value:(n==null?void 0:n.notes)||"",permanentId:`step-notes-lock-${e}`,isPermanent:a,hintId:"",saveButtonId:`step-notes-save-${e}`})}
            ${me({reviewerInputId:`step-reviewer-${e}`,reviewerValue:o,signoffInputId:`step-signoff-${e}`,signoffValue:s,isComplete:l,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`})}
        </section>
    `}function aa(){var n,a,o,s,l;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",async()=>{var r;let i=Be(De);await Le(i.sheetName,i.title,i.subtitle),xe({activeView:"home",activeStepId:null}),(r=document.getElementById("pf-hero"))==null||r.scrollIntoView({behavior:"smooth",block:"start"})}),(a=document.getElementById("nav-selector"))==null||a.addEventListener("click",()=>{window.location.href=Mn}),(o=document.getElementById("nav-prev"))==null||o.addEventListener("click",()=>ze(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>ze(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",i=>{i.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",i=>{!(t!=null&&t.contains(i.target))&&!(e!=null&&e.contains(i.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(l=document.getElementById("nav-config"))==null||l.addEventListener("click",async()=>{t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"),await Sa()}),document.querySelectorAll("[data-step-card]").forEach(i=>{let r=Number(i.getAttribute("data-step-index")),c=Number(i.getAttribute("data-step-id"));i.addEventListener("click",()=>$e(r,c))}),R.activeView==="config"?sa():R.activeView==="step"&&R.activeStepId!=null&&oa(R.activeStepId)}function oa(e){var u,d,g,y,w,h,E,C,I,_,j,L,B,Q,U,q,T,m;let t=e===2?document.getElementById("step-notes-input"):document.getElementById(`step-notes-${e}`),n=e===2?document.getElementById("step-reviewer-name"):document.getElementById(`step-reviewer-${e}`),a=e===2?document.getElementById("step-signoff-date"):document.getElementById(`step-signoff-${e}`),o=document.getElementById("step-back-btn"),s=e===2?document.getElementById("step-notes-lock-2"):document.getElementById(`step-notes-lock-${e}`),l=e===2?document.getElementById("step-notes-save-2"):document.getElementById(`step-notes-save-${e}`);l==null||l.addEventListener("click",async()=>{let v=(t==null?void 0:t.value)||"";await ne(e,"notes",v),de(l,!0)});let i=e===2?document.getElementById("headcount-signoff-save"):document.getElementById(`step-signoff-save-${e}`);i==null||i.addEventListener("click",async()=>{let v=(n==null?void 0:n.value)||"";await ne(e,"reviewer",v),de(i,!0)}),lt();let r=e===2?"headcount-signoff-toggle":`step-signoff-toggle-${e}`,c=`${r}-prev`,p=`${r}-next`,f=e===2?"step-signoff-date":`step-signoff-${e}`;vn(e,{buttonId:r,inputId:f,canActivate:e===2?()=>{var P;return!wn()||((P=document.getElementById("step-notes-input"))==null?void 0:P.value.trim())||""?!0:(z("Please enter a brief explanation of the headcount differences before completing this step.","info"),!1)}:null,onComplete:ia(e)}),pn(c,p),o==null||o.addEventListener("click",async()=>{let v=Be(De);await Le(v.sheetName,v.title,v.subtitle),xe({activeView:"home",activeStepId:null})}),s==null||s.addEventListener("click",async()=>{let v=!s.classList.contains("is-locked");rt(s,v),await hn(e,v)}),e===6&&((u=document.getElementById("archive-run-btn"))==null||u.addEventListener("click",()=>{ha()})),e===1&&((d=document.getElementById("import-open-data-btn"))==null||d.addEventListener("click",()=>gn("PTO_Data")),(g=document.getElementById("import-clear-btn"))==null||g.addEventListener("click",()=>wa())),e===2&&((y=document.getElementById("headcount-skip-btn"))==null||y.addEventListener("click",()=>{A.skipAnalysis=!A.skipAnalysis;let v=document.getElementById("headcount-skip-btn");v==null||v.classList.toggle("is-active",A.skipAnalysis),A.skipAnalysis&&rn(),sn()}),(w=document.getElementById("headcount-run-btn"))==null||w.addEventListener("click",()=>ft()),(h=document.getElementById("headcount-refresh-btn"))==null||h.addEventListener("click",()=>ft()),Wa(),A.skipAnalysis&&rn(),sn()),e===3&&((E=document.getElementById("quality-run-btn"))==null||E.addEventListener("click",()=>Qt()),(C=document.getElementById("quality-refresh-btn"))==null||C.addEventListener("click",()=>Qt()),(I=document.getElementById("quality-acknowledge-btn"))==null||I.addEventListener("click",()=>da())),e===4&&((_=document.getElementById("analysis-refresh-btn"))==null||_.addEventListener("click",()=>Xt()),(j=document.getElementById("analysis-run-btn"))==null||j.addEventListener("click",()=>Xt()),(L=document.getElementById("payrate-save-btn"))==null||L.addEventListener("click",Kt),(B=document.getElementById("payrate-ignore-btn"))==null||B.addEventListener("click",ca),(Q=document.getElementById("payrate-input"))==null||Q.addEventListener("keydown",v=>{v.key==="Enter"&&Kt()})),e===5&&((U=document.getElementById("je-create-btn"))==null||U.addEventListener("click",()=>fa()),(q=document.getElementById("je-run-btn"))==null||q.addEventListener("click",()=>fn()),(T=document.getElementById("je-export-btn"))==null||T.addEventListener("click",()=>ga()),(m=document.getElementById("je-upload-btn"))==null||m.addEventListener("click",()=>ma()))}function sa(){var i,r,c,p,f;$t("config-payroll-date",{onChange:u=>{if(le(k.payrollDate,u),!!u){if(O.overrides.accountingPeriod=!1,O.overrides.journalId=!1,!O.overrides.accountingPeriod){let d=Da(u);if(d){let g=document.getElementById("config-accounting-period");g&&(g.value=d),le(k.accountingPeriod,d)}}if(!O.overrides.journalId){let d=$a(u);if(d){let g=document.getElementById("config-journal-id");g&&(g.value=d),le(k.journalEntryId,d)}}}}});let e=document.getElementById("config-accounting-period");e==null||e.addEventListener("change",u=>{O.overrides.accountingPeriod=!!u.target.value,le(k.accountingPeriod,u.target.value||"")});let t=document.getElementById("config-journal-id");t==null||t.addEventListener("change",u=>{O.overrides.journalId=!!u.target.value,le(k.journalEntryId,u.target.value.trim())}),(i=document.getElementById("config-company-name"))==null||i.addEventListener("change",u=>{le(k.companyName,u.target.value.trim())}),(r=document.getElementById("config-payroll-provider"))==null||r.addEventListener("change",u=>{le(k.payrollProvider,u.target.value.trim())}),(c=document.getElementById("config-accounting-link"))==null||c.addEventListener("change",u=>{le(k.accountingSoftware,u.target.value.trim())}),(p=document.getElementById("config-user-name"))==null||p.addEventListener("change",u=>{le(k.reviewerName,u.target.value.trim())});let n=document.getElementById("config-notes");n==null||n.addEventListener("input",u=>{ne(0,"notes",u.target.value)});let a=document.getElementById("config-notes-lock");a==null||a.addEventListener("click",async()=>{let u=!a.classList.contains("is-locked");rt(a,u),await hn(0,u)});let o=document.getElementById("config-notes-save");o==null||o.addEventListener("click",async()=>{n&&(await ne(0,"notes",n.value),de(o,!0))});let s=document.getElementById("config-reviewer");s==null||s.addEventListener("change",u=>{let d=u.target.value.trim();ne(0,"reviewer",d);let g=document.getElementById("config-signoff-date");if(d&&g&&!g.value){let y=mt();g.value=y,ne(0,"signOffDate",y),yn(0,!0)}}),(f=document.getElementById("config-signoff-date"))==null||f.addEventListener("change",u=>{ne(0,"signOffDate",u.target.value||"")});let l=document.getElementById("config-signoff-save");l==null||l.addEventListener("click",async()=>{var g,y;let u=((g=s==null?void 0:s.value)==null?void 0:g.trim())||"",d=((y=document.getElementById("config-signoff-date"))==null?void 0:y.value)||"";await ne(0,"reviewer",u),await ne(0,"signOffDate",d),de(l,!0)}),lt(),vn(0,{buttonId:"config-signoff-toggle",inputId:"config-signoff-date",onComplete:()=>{Ma(),dn(0),un()}}),pn("config-signoff-toggle-prev","config-signoff-toggle-next")}function $e(e,t=null){if(e<0||e>=ae.length)return;Je=e;let n=t!=null?t:ae[e].id;xe({focusedIndex:e,activeView:n===0?"config":"step",activeStepId:n});let o=Hn[n];o&&gn(o),n===2&&!A.hasAnalyzed&&(kn(),ft())}function ia(e){return e===6?null:()=>dn(e)}function dn(e){let t=ae.findIndex(a=>a.id===e);if(t===-1)return;let n=t+1;n<ae.length&&($e(n,ae[n].id),un())}function un(){let e=[document.querySelector(".pf-root"),document.querySelector(".pf-step-guide"),document.body];for(let t of e)t&&t.scrollTo({top:0,behavior:"smooth"});window.scrollTo({top:0,behavior:"smooth"})}function ze(e){let t=R.focusedIndex+e,n=Math.max(0,Math.min(ae.length-1,t));$e(n,ae[n].id),window.scrollTo({top:0,behavior:"smooth"})}function pn(e,t){var n,a;(n=document.getElementById(e))==null||n.addEventListener("click",()=>ze(-1)),(a=document.getElementById(t))==null||a.addEventListener("click",()=>ze(1))}function ra(){if(Je===null)return;let e=document.querySelector(`[data-step-index="${Je}"]`);Je=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}function la(e){var a;let t=bn(O.completes[e]),n=!!((a=O.steps[e])!=null&&a.signOffDate);return t||n}function xe(e){e.stepStatuses&&(R.stepStatuses={...R.stepStatuses,...e.stepStatuses}),Object.assign(R,{...e,stepStatuses:R.stepStatuses}),ie()}function oe(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}async function Kt(){let e=document.getElementById("payrate-input");if(!e)return;let t=parseFloat(e.value),n=e.dataset.employee,a=parseInt(e.dataset.row,10);if(isNaN(t)||t<=0){z("Please enter a valid pay rate greater than 0.","info");return}if(!n||isNaN(a)){console.error("Missing employee data on input");return}Z(!0,"Updating pay rate...");try{await Excel.run(async o=>{let s=o.workbook.worksheets.getItem("PTO_Analysis"),l=s.getCell(a-1,3);l.values=[[t]];let i=s.getCell(a-1,8);i.load("values"),await o.sync();let c=(Number(i.values[0][0])||0)*t,p=s.getCell(a-1,9);p.values=[[c]];let f=s.getCell(a-1,10);f.load("values"),await o.sync();let u=Number(f.values[0][0])||0,d=c-u,g=s.getCell(a-1,11);g.values=[[d]],await o.sync()}),ee.missingPayRates=ee.missingPayRates.filter(o=>o.name!==n),Z(!1),$e(3,3)}catch(o){console.error("Failed to save pay rate:",o),z(`Failed to save pay rate: ${o.message}`,"error"),Z(!1)}}function ca(){let e=document.getElementById("payrate-input");if(!e)return;let t=e.dataset.employee;t&&(ee.ignoredMissingPayRates.add(t),ee.missingPayRates=ee.missingPayRates.filter(n=>n.name!==t)),$e(3,3)}async function Qt(){if(!oe()){z("Excel is not available. Open this module inside Excel to run quality check.","info");return}K.loading=!0,Z(!0,"Analyzing data quality..."),de(document.getElementById("quality-save-btn"),!1);try{await Excel.run(async t=>{var w;let a=t.workbook.worksheets.getItem("PTO_Data").getUsedRangeOrNullObject();a.load("values"),await t.sync();let o=a.isNullObject?[]:a.values||[];if(!o.length||o.length<2)throw new Error("PTO_Data is empty or has no data rows.");let s=(o[0]||[]).map(h=>D(h));console.log("[Data Quality] PTO_Data headers:",o[0]);let l=s.findIndex(h=>h==="employee name"||h==="employeename");l===-1&&(l=s.findIndex(h=>h.includes("employee")&&h.includes("name"))),l===-1&&(l=s.findIndex(h=>h==="name"||h.includes("name")&&!h.includes("company")&&!h.includes("form"))),console.log("[Data Quality] Employee name column index:",l,"Header:",(w=o[0])==null?void 0:w[l]);let i=$(s,["balance"]),r=$(s,["accrual rate","accrualrate"]),c=$(s,["carry over","carryover"]),p=$(s,["ytd accrued","ytdaccrued"]),f=$(s,["ytd used","ytdused"]),u=[],d=[],g=[],y=o.slice(1);y.forEach((h,E)=>{let C=E+2,I=l!==-1?String(h[l]||"").trim():`Row ${C}`;if(!I)return;let _=i!==-1&&Number(h[i])||0,j=r!==-1&&Number(h[r])||0,L=c!==-1&&Number(h[c])||0,B=p!==-1&&Number(h[p])||0,Q=f!==-1&&Number(h[f])||0,U=L+B;_<0?u.push({name:I,issue:`Negative balance: ${_.toFixed(2)} hrs`,rowIndex:C}):Q>U&&U>0&&u.push({name:I,issue:`Used ${Q.toFixed(0)} hrs but only ${U.toFixed(0)} available`,rowIndex:C}),_===0&&(L>0||B>0)&&d.push({name:I,rowIndex:C}),j>8&&g.push({name:I,accrualRate:j,rowIndex:C})}),K.balanceIssues=u,K.zeroBalances=d,K.accrualOutliers=g,K.totalIssues=u.length,K.totalEmployees=y.filter(h=>h.some(E=>E!==null&&E!=="")).length,K.hasRun=!0});let e=K.balanceIssues.length>0;xe({stepStatuses:{3:e?"blocked":"complete"}})}catch(e){console.error("Data quality check error:",e),z(`Quality check failed: ${e.message}`,"error"),K.hasRun=!1}finally{K.loading=!1,Z(!1),ie()}}function da(){K.acknowledged=!0,xe({stepStatuses:{3:"complete"}}),ie()}async function ua(){if(oe())try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem("PTO_Data"),n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=t.getUsedRangeOrNullObject();if(a.load("values"),n.load("isNullObject"),await e.sync(),n.isNullObject){ee.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let o=n.getUsedRangeOrNullObject();o.load("values"),await e.sync();let s=a.isNullObject?[]:a.values||[],l=o.isNullObject?[]:o.values||[];if(!s.length||!l.length){ee.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let i=(p,f,u)=>{let d=(p[0]||[]).map(w=>D(w)),g=$(d,f);return g===-1?null:p.slice(1).reduce((w,h)=>w+(Number(h[g])||0),0)},r=[{key:"accrualRate",aliases:["accrual rate","accrualrate"]},{key:"carryOver",aliases:["carry over","carryover","carry_over"]},{key:"ytdAccrued",aliases:["ytd accrued","ytdaccrued","ytd_accrued"]},{key:"ytdUsed",aliases:["ytd used","ytdused","ytd_used"]},{key:"balance",aliases:["balance"]}],c={};for(let p of r){let f=i(s,p.aliases,"PTO_Data"),u=i(l,p.aliases,"PTO_Analysis");if(f===null||u===null)c[p.key]=null;else{let d=Math.abs(f-u)<.01;c[p.key]={match:d,ptoData:f,ptoAnalysis:u}}}ee.completenessCheck=c})}catch(e){console.error("Completeness check failed:",e)}}async function Xt(){if(!oe()){z("Excel is not available. Open this module inside Excel to run analysis.","info");return}Z(!0,"Running analysis...");try{await kn(),await ua(),ee.cleanDataReady=!0,ie()}catch(e){console.error("Full analysis error:",e),z(`Analysis failed: ${e.message}`,"error")}finally{Z(!1)}}async function fn(){if(!oe()){z("Excel is not available. Open this module inside Excel to run journal checks.","info");return}J.loading=!0,J.lastError=null,de(document.getElementById("je-save-btn"),!1),ie();try{let e=await Excel.run(async t=>{let a=t.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();a.load("values");let o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis");o.load("isNullObject"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty. Generate the JE first.");let l=(s[0]||[]).map(C=>D(C)),i=$(l,["debit"]),r=$(l,["credit"]),c=$(l,["lineamount","line amount"]),p=$(l,["account number","accountnumber"]);if(i===-1||r===-1)throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");let f=0,u=0,d=0,g=0;s.slice(1).forEach(C=>{let I=Number(C[i])||0,_=Number(C[r])||0,j=c!==-1&&Number(C[c])||0,L=p!==-1?String(C[p]||"").trim():"";f+=I,u+=_,d+=j,L&&L!=="21540"&&(g+=j)});let y=0;if(!o.isNullObject){let C=o.getUsedRangeOrNullObject();C.load("values"),await t.sync();let I=C.isNullObject?[]:C.values||[];if(I.length>1){let _=(I[0]||[]).map(L=>D(L)),j=$(_,["change"]);j!==-1&&I.slice(1).forEach(L=>{y+=Number(L[j])||0})}}let w=f-u,h=[];Math.abs(w)>=.01?h.push({check:"Debits = Credits",passed:!1,detail:w>0?`Debits exceed credits by $${Math.abs(w).toLocaleString(void 0,{minimumFractionDigits:2})}`:`Credits exceed debits by $${Math.abs(w).toLocaleString(void 0,{minimumFractionDigits:2})}`}):h.push({check:"Debits = Credits",passed:!0,detail:""}),Math.abs(d)>=.01?h.push({check:"Line Amounts Sum to Zero",passed:!1,detail:`Line amounts sum to $${d.toLocaleString(void 0,{minimumFractionDigits:2})} (should be $0.00)`}):h.push({check:"Line Amounts Sum to Zero",passed:!0,detail:""});let E=Math.abs(g-y);return E>=.01?h.push({check:"JE Matches Analysis Total",passed:!1,detail:`JE expense total ($${g.toLocaleString(void 0,{minimumFractionDigits:2})}) differs from PTO_Analysis Change total ($${y.toLocaleString(void 0,{minimumFractionDigits:2})}) by $${E.toLocaleString(void 0,{minimumFractionDigits:2})}`}):h.push({check:"JE Matches Analysis Total",passed:!0,detail:""}),{debitTotal:f,creditTotal:u,difference:w,lineAmountSum:d,jeChangeTotal:g,analysisChangeTotal:y,issues:h,validationRun:!0}});Object.assign(J,e,{lastError:null})}catch(e){console.warn("PTO JE summary:",e),J.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",J.debitTotal=null,J.creditTotal=null,J.difference=null,J.lineAmountSum=null,J.jeChangeTotal=null,J.analysisChangeTotal=null,J.issues=[],J.validationRun=!1}finally{J.loading=!1,ie()}}var pa={"general & administrative":"64110","general and administrative":"64110","g&a":"64110","research & development":"62110","research and development":"62110","r&d":"62110",marketing:"61610","cogs onboarding":"53110","cogs prof. services":"56110","cogs professional services":"56110","sales & marketing":"61110","sales and marketing":"61110","cogs support":"52110","client success":"61811"},Zt="21540";async function fa(){if(!oe()){z("Excel is not available. Open this module inside Excel to create the journal entry.","info");return}Z(!0,"Creating PTO Journal Entry...");try{await Excel.run(async e=>{let t=[],n=e.workbook.tables.getItemOrNullObject(ce[0]);if(n.load("isNullObject"),await e.sync(),n.isNullObject){let m=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");if(m.load("isNullObject"),await e.sync(),!m.isNullObject){let v=m.getUsedRangeOrNullObject();v.load("values"),await e.sync();let P=v.isNullObject?[]:v.values||[];t=P.length>1?P.slice(1):[]}}else{let m=n.getDataBodyRange();m.load("values"),await e.sync(),t=m.values||[]}let a=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis");if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PTO_Analysis sheet not found. Please ensure the worksheet exists.");let o=a.getUsedRangeOrNullObject();o.load("values");let s=e.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");s.load("isNullObject"),await e.sync();let l=[];if(!s.isNullObject){let m=s.getUsedRangeOrNullObject();m.load("values"),await e.sync(),l=m.isNullObject?[]:m.values||[]}let i=o.isNullObject?[]:o.values||[];if(!i.length||i.length<2)throw new Error("PTO_Analysis is empty or has no data rows. Run the analysis first (Step 4).");let r={};t.forEach(m=>{let v=String(m[1]||"").trim(),P=m[2];v&&(r[v]=P)}),(!r[k.journalEntryId]||!r[k.payrollDate])&&console.warn("[JE Draft] Missing config values - RefNumber:",r[k.journalEntryId],"TxnDate:",r[k.payrollDate]);let c=r[k.journalEntryId]||"",p=r[k.payrollDate]||"",f=r[k.accountingPeriod]||"",u="";if(p)try{let m;if(typeof p=="number"||/^\d{4,5}$/.test(String(p).trim())){let v=Number(p),P=new Date(1899,11,30);m=new Date(P.getTime()+v*24*60*60*1e3)}else m=new Date(p);if(!isNaN(m.getTime())&&m.getFullYear()>1970){let v=String(m.getMonth()+1).padStart(2,"0"),P=String(m.getDate()).padStart(2,"0"),N=m.getFullYear();u=`${v}/${P}/${N}`}else console.warn("[JE Draft] Date parsing resulted in invalid date:",p,"->",m),u=String(p)}catch(m){console.warn("[JE Draft] Could not parse TxnDate:",p,m),u=String(p)}let d=f?`${f} PTO Accrual`:"PTO Accrual",g={};if(l.length>1){let m=(l[0]||[]).map(N=>D(N)),v=$(m,["account number","accountnumber","account","acct"]),P=$(m,["account name","accountname","name","description"]);v!==-1&&P!==-1&&l.slice(1).forEach(N=>{let Y=String(N[v]||"").trim(),pe=String(N[P]||"").trim();Y&&(g[Y]=pe)})}let y=(i[0]||[]).map(m=>D(m));console.log("[JE Draft] PTO_Analysis headers:",y),console.log("[JE Draft] PTO_Analysis row count:",i.length-1);let w=$(y,["department"]),h=$(y,["change"]);if(console.log("[JE Draft] Column indices - Department:",w,"Change:",h),w===-1||h===-1)throw new Error(`Could not find required columns in PTO_Analysis. Found headers: ${y.join(", ")}. Looking for "Department" (found: ${w!==-1}) and "Change" (found: ${h!==-1}).`);let E={},C=0,I=0,_=0;if(i.slice(1).forEach((m,v)=>{C++;let P=String(m[w]||"").trim(),N=m[h],Y=Number(N)||0;if(v<3&&console.log(`[JE Draft] Row ${v+2}: Dept="${P}", Change raw="${N}", Change num=${Y}`),!P){_++;return}if(Y===0){I++;return}E[P]||(E[P]=0),E[P]+=Y}),console.log(`[JE Draft] Data summary: ${C} rows, ${I} with zero change, ${_} missing dept`),console.log("[JE Draft] Department totals:",E),Object.keys(E).length===0){let m=`No journal entry lines could be created.

`;throw I===C?(m+=`All 'Change' amounts in PTO_Analysis are $0.00.

`,m+=`Common causes:
`,m+=`\u2022 Missing Pay Rate data (Liability = Balance \xD7 Pay Rate)
`,m+=`\u2022 No prior period data to compare against
`,m+=`\u2022 PTO Analysis hasn't been run yet

`,m+="Please verify Pay Rate values exist in PTO_Analysis."):_===C?(m+=`All rows are missing Department values.

`,m+="Please ensure the 'Department' column is populated in PTO_Analysis."):(m+=`Found ${C} rows but none had both a Department and non-zero Change amount.
`,m+=`\u2022 ${I} rows with zero change
`,m+=`\u2022 ${_} rows missing department`),new Error(m)}let L=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],B=[L],Q=0,U=0;Object.entries(E).forEach(([m,v])=>{if(Math.abs(v)<.01)return;let P=m.toLowerCase().trim(),N=pa[P]||"",Y=g[N]||"",pe=v>0?Math.abs(v):0,S=v<0?Math.abs(v):0;Q+=pe,U+=S,B.push([c,u,N,Y,v,pe,S,d,m])});let q=Q-U;if(Math.abs(q)>=.01){let m=q<0?Math.abs(q):0,v=q>0?Math.abs(q):0,P=g[Zt]||"Accrued PTO";B.push([c,u,Zt,P,-q,m,v,d,""])}let T=e.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");if(T.load("isNullObject"),await e.sync(),T.isNullObject)T=e.workbook.worksheets.add("PTO_JE_Draft");else{let m=T.getUsedRangeOrNullObject();m.load("isNullObject"),await e.sync(),m.isNullObject||m.clear()}if(B.length>0){let m=T.getRangeByIndexes(0,0,B.length,L.length);m.values=B;let v=T.getRangeByIndexes(0,0,1,L.length);dt(v);let P=B.length-1;P>0&&(he(T,4,P,!0),he(T,5,P),he(T,6,P)),m.format.autofitColumns()}await e.sync(),T.activate(),T.getRange("A1").select(),await e.sync()}),await fn()}catch(e){console.error("Create JE Draft error:",e),z(`Unable to create Journal Entry: ${e.message}`,"error")}finally{Z(!1)}}async function ga(){if(!oe()){z("Excel is not available. Open this module inside Excel to export.","info");return}Z(!0,"Preparing JE CSV...");try{let{rows:e}=await Excel.run(async n=>{let o=n.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();o.load("values"),await n.sync();let s=o.isNullObject?[]:o.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty.");return{rows:s}}),t=qa(e);Ya(`pto-je-draft-${mt()}.csv`,t)}catch(e){console.error("PTO JE export:",e),z("Unable to export the JE draft. Confirm the sheet has data.","error")}finally{Z(!1)}}async function ma(){let e=te(k.accountingSoftware)||te("SS_Accounting_Software");if(!e&&W())try{let t=await Qe(ce);e=t.SS_Accounting_Software||t.Accounting_Software||t[k.accountingSoftware]}catch(t){console.warn("Error reading accounting software URL:",t)}if(!e){z("No accounting software URL configured. Add SS_Accounting_Software to SS_PF_Config.","info",5e3);return}!e.startsWith("http://")&&!e.startsWith("https://")&&(e="https://"+e),window.open(e,"_blank"),z("Opening accounting software...","success",2e3)}async function ha(){if(!(!R.workbookReady||!oe())){Z(!0,"Archiving PTO outputs...");try{await Excel.run(async e=>{let t=await Ca(e),n=e.workbook.tables.getItemOrNullObject("PTOArchiveLog");n.load("isNullObject"),await e.sync(),n.isNullObject||n.rows.add(null,[[new Date().toISOString(),"Archived PTO run","Completed via module"]]),await Pa(e),await Ia(e,t);let a=["PTO_JE_Draft","PTO_Analysis","PTO_Data"];for(let o of a)await Ea(e,o);await _a(e),await e.sync()}),xe({stepStatuses:{1:"pending",2:"pending",3:"pending",4:"pending",5:"pending",6:"complete"}}),ya()}catch(e){console.error(e),z("Archive failed: "+e.message,"error")}finally{Z(!1)}}}function ya(){document.querySelectorAll(".pf-save-prompt").forEach(n=>n.remove());let e=document.createElement("div");if(e.className="pf-save-prompt",e.innerHTML=`
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
    `,!document.getElementById("pf-save-prompt-styles")){let n=document.createElement("style");n.id="pf-save-prompt-styles",n.textContent=`
            .pf-save-prompt {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: linear-gradient(145deg, rgba(30, 30, 50, 0.98), rgba(20, 20, 35, 0.99));
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
                padding: 12px 28px;
                border-radius: 12px;
                font-size: 15px;
                font-weight: 600;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 8px;
                transition: all 0.2s ease;
            }
            .pf-save-prompt-btn:hover {
                transform: translateY(-2px);
                box-shadow: 0 8px 20px rgba(99, 102, 241, 0.4);
            }
            .pf-save-prompt-btn:active {
                transform: translateY(0);
            }
            .pf-save-prompt.closing {
                animation: pf-prompt-fade-out 0.2s ease forwards;
            }
            @keyframes pf-prompt-fade-out {
                to { opacity: 0; transform: translate(-50%, -50%) scale(0.95); }
            }
        `,document.head.appendChild(n)}document.body.appendChild(e),e.querySelector(".pf-save-prompt-btn").addEventListener("click",()=>{e.classList.add("closing"),setTimeout(()=>{e.remove(),va()},200)})}function va(){document.querySelectorAll(".pf-toast, .pf-success-toast").forEach(o=>o.remove());let e=[{icon:'<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z"/></svg>',title:"Done.",subtitle:"Efficiency unlocked."},{icon:'<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"/><path d="m9 12 2 2 4-4"/></svg>',title:"Locked in.",subtitle:"You're good to go."},{icon:'<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="1"/><path d="M20.2 20.2c2.04-2.03.02-7.36-4.5-11.9-4.54-4.52-9.87-6.54-11.9-4.5-2.04 2.03-.02 7.36 4.5 11.9 4.54 4.52 9.87 6.54 11.9 4.5Z"/><path d="M15.7 15.7c4.52-4.54 6.54-9.87 4.5-11.9-2.03-2.04-7.36-.02-11.9 4.5-4.52 4.54-6.54 9.87-4.5 11.9 2.03 2.04 7.36.02 11.9-4.5Z"/></svg>',title:"Stored.",subtitle:"Everything stays aligned."},{icon:'<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>',title:"PTO archived.",subtitle:"Flow restored."}],t=e[Math.floor(Math.random()*e.length)],n=document.createElement("div");if(n.className="pf-success-toast",n.innerHTML=`
        <div class="pf-success-toast-icon">${t.icon}</div>
        <div class="pf-success-toast-text">
            <div class="pf-success-toast-title">${t.title}</div>
            <div class="pf-success-toast-subtitle">${t.subtitle}</div>
        </div>
        <div class="pf-success-toast-progress"></div>
    `,!document.getElementById("pf-success-toast-styles")){let o=document.createElement("style");o.id="pf-success-toast-styles",o.textContent=`
            .pf-success-toast {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: linear-gradient(145deg, rgba(30, 30, 50, 0.98), rgba(20, 20, 35, 0.99));
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
        `,document.head.appendChild(o)}let a=document.createElement("div");a.className="pf-success-backdrop",document.body.appendChild(a),document.body.appendChild(n),setTimeout(()=>{n.classList.add("closing"),a.style.opacity="0",a.style.transition="opacity 0.3s ease",setTimeout(()=>{n.remove(),a.remove(),ba()},300)},5e3)}function ba(){window.location.href="../module-selector/index.html"}async function gn(e){if(!(!e||!oe()))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e);n.activate(),n.getRange("A1").select(),await t.sync()})}catch(t){console.error(t)}}async function wa(){if(!(!oe()||!await Ln(`All data in PTO_Data will be permanently removed.

This action cannot be undone.`,{title:"Clear PTO Data",icon:"\u{1F5D1}\uFE0F",confirmText:"Clear Data",cancelText:"Keep Data",destructive:!0}))){Z(!0);try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("PTO_Data"),a=n.getUsedRangeOrNullObject();a.load("rowCount"),await t.sync(),!a.isNullObject&&a.rowCount>1&&(n.getRangeByIndexes(1,0,a.rowCount-1,20).clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),z("PTO_Data cleared successfully. You can now paste new data.","success")}catch(t){console.error("Clear PTO_Data error:",t),z(`Failed to clear PTO_Data: ${t.message}`,"error")}finally{Z(!1)}}}async function ka(){if(!oe())return[];try{return await Excel.run(async e=>{let t=e.workbook.worksheets;return t.load("items/name,visibility"),await e.sync(),t.items.filter(a=>{let s=(a.name||"").toUpperCase();return s.startsWith("SS_")||s.includes("MAPPING")||s.includes("HOMEPAGE")}).map(a=>({name:a.name,visible:a.visibility===Excel.SheetVisibility.visible,isHomepage:(a.name||"").toUpperCase().includes("HOMEPAGE")})).sort((a,o)=>a.isHomepage&&!o.isHomepage?1:!a.isHomepage&&o.isHomepage?-1:a.name.localeCompare(o.name))})}catch(e){return console.error("[Config] Error reading configuration sheets:",e),[]}}function Oa(){if(document.getElementById("config-sheet-modal"))return;let e=document.createElement("div");if(e.id="config-sheet-modal",e.className="pf-config-modal hidden",e.innerHTML=`
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
            .pf-config-modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
            .pf-config-close { background: transparent; border: none; color: #f8fafc; font-size: 20px; cursor: pointer; }
            .pf-config-hint { margin: 0 0 12px 0; color: #cbd5e1; font-size: 14px; }
            .pf-config-sheet-list { display: flex; flex-direction: column; gap: 10px; max-height: 260px; overflow-y: auto; }
            .pf-config-sheet { display: flex; justify-content: space-between; align-items: center; padding: 12px 14px; background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.18); border-radius: 10px; cursor: pointer; color: #e2e8f0; font-weight: 600; }
            .pf-config-sheet:hover { background: rgba(255,255,255,0.16); }
            .pf-config-pill { font-size: 12px; color: #c7d2fe; }
        `,document.head.appendChild(t)}}async function Sa(){Oa();let e=document.getElementById("config-sheet-modal"),t=document.getElementById("config-sheet-list");if(!e||!t)return;t.textContent="Loading\u2026",e.classList.remove("hidden");let n=await ka();n.length?(t.innerHTML="",n.forEach(a=>{let o=document.createElement("button");o.type="button",o.className="pf-config-sheet",o.innerHTML=`<span>${a.name}</span><span class="pf-config-pill">${a.visible?"Visible":"Hidden"}</span>`,o.addEventListener("click",async()=>{await xa(a.name),e.classList.add("hidden")}),t.appendChild(o)})):t.textContent="No configuration sheets found.",e.querySelectorAll("[data-close]").forEach(a=>a.addEventListener("click",()=>e.classList.add("hidden")))}async function xa(e){if(!(!e||!oe()))try{await Excel.run(async t=>{let n=t.workbook.worksheets,a=n.getItemOrNullObject(e);a.load("isNullObject,visibility"),await t.sync(),a.isNullObject&&(a=n.add(e)),a.visibility=Excel.SheetVisibility.visible,await t.sync(),a.activate(),a.getRange("A1").select(),await t.sync(),console.log(`[Config] Opened sheet: ${e}`)})}catch(t){console.error("[Config] Error opening sheet",e,t)}}async function Ea(e,t){let n=e.workbook.worksheets.getItem(t),a=n.getUsedRangeOrNullObject();if(a.load("rowCount"),await e.sync(),a.isNullObject||a.rowCount<=1)return;n.getRangeByIndexes(1,0,a.rowCount-1,a.columnCount).clear()}async function Ca(e){let t=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis");if(t.load("isNullObject"),await e.sync(),t.isNullObject)return[];let n=t.getUsedRangeOrNullObject();n.load("values"),await e.sync();let a=n.isNullObject?[]:n.values||[];return a.length<=1?[]:a.slice(1)}async function _a(e){let t=e.workbook.tables.getItemOrNullObject(ce[0]);if(await e.sync(),t.isNullObject)return;let n=t.getDataBodyRange(),a=t.getHeaderRowRange();n.load("values"),a.load("values"),await e.sync();let s=(a.values[0]||[]).map(i=>D(i)),l={field:s.findIndex(i=>i==="field"||i==="field name"||i==="setting"),permanent:s.findIndex(i=>i==="permanent"||i==="persist"),value:s.findIndex(i=>i==="value"||i==="setting value")};l.field===-1||l.value===-1||l.permanent===-1||(n.values||[]).forEach((i,r)=>{var f;let c=String((f=i[l.permanent])!=null?f:"").trim().toLowerCase();c!=="y"&&c!=="yes"&&c!=="true"&&c!=="t"&&c!=="1"&&(n.getCell(r,l.value).values=[[""]])})}async function Pa(e){let t=await Ta(e);if(!t.length)return;let n=[];for(let i of t)try{let r=e.workbook.worksheets.getItemOrNullObject(i);if(r.load("isNullObject"),await e.sync(),r.isNullObject)continue;let c=r.getUsedRangeOrNullObject();c.load("values"),await e.sync();let f=(c.isNullObject?[]:c.values||[]).map(u=>u.map(d=>`"${String(d!=null?d:"").replace(/"/g,'""')}"`).join(",")).join(`
`);n.push(`# Sheet: ${i}
${f}`)}catch(r){console.warn("PTO: unable to export sheet",i,r)}if(!n.length)return;let a=`${new Date().toISOString().slice(0,10)} - Tai Software PTO Accrual.xlsx`,o=new Blob([n.join(`

`)],{type:"text/csv"}),s=URL.createObjectURL(o),l=document.createElement("a");l.href=s,l.download=a,document.body.appendChild(l),l.click(),document.body.removeChild(l),URL.revokeObjectURL(s)}async function Ta(e){let t=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");if(t.load("isNullObject"),await e.sync(),t.isNullObject)return[];let n=t.getUsedRangeOrNullObject();n.load("values"),await e.sync();let a=n.isNullObject?[]:n.values||[];if(!a.length)return[];let o=(a[0]||[]).map(r=>D(r)),s={category:o.findIndex(r=>r==="category"),module:o.findIndex(r=>r==="module"),field:o.findIndex(r=>r==="field"),value:o.findIndex(r=>r==="value")};if(s.category===-1||s.field===-1)return[];let l=D(ue),i=a.slice(1).filter(r=>{let c=D(r[s.category]),p=s.module>=0?D(r[s.module]):"",f=s.value>=0?D(r[s.value]):"";return c==="tab-structure"&&(p===l||f==="pto-accrual")}).map(r=>{var c;return String((c=r[s.field])!=null?c:"").trim()}).filter(Boolean);return Array.from(new Set(i))}async function Ia(e,t=null){let n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=e.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary");if(n.load("isNullObject"),a.load("isNullObject"),await e.sync(),n.isNullObject||a.isNullObject)return;let o=t;if(!o||!o.length){let i=n.getUsedRangeOrNullObject();i.load("values"),await e.sync(),o=i.isNullObject?[]:i.values||[]}if(!o.length)return;let s=a.getUsedRangeOrNullObject();s.load("isNullObject"),await e.sync(),s.isNullObject||s.clear();let l=a.getRangeByIndexes(0,0,o.length,o[0].length);l.values=o,l.format.autofitColumns(),a.getRange("A1").select(),await e.sync()}function te(e){var n,a;let t=String(e!=null?e:"").trim();return(a=(n=O.values)==null?void 0:n[t])!=null?a:""}function be(e){var n;if(e)return e;let t=te(k.reviewerName);if(t)return t;if((n=window.PrairieForge)!=null&&n._sharedConfigCache){let a=window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer")||window.PrairieForge._sharedConfigCache.get("Default_Reviewer");if(a)return a}return""}function le(e,t,n={}){var l;let a=String(e!=null?e:"").trim();if(!a)return;O.values[a]=t!=null?t:"";let o=(l=n.debounceMs)!=null?l:0;if(!o){let i=Se.get(a);i&&clearTimeout(i),Se.delete(a),Pe(a,t!=null?t:"",ce);return}Se.has(a)&&clearTimeout(Se.get(a));let s=setTimeout(()=>{Se.delete(a),Pe(a,t!=null?t:"",ce)},o);Se.set(a,s)}function D(e){return String(e!=null?e:"").trim().toLowerCase()}function Z(e,t="Working..."){let n=document.getElementById(Vn);n&&(n.style.display="none")}function pt(){Gn()}typeof Office!="undefined"&&Office.onReady?Office.onReady(()=>pt()).catch(()=>pt()):pt();function we(e){return O.steps[e]||{notes:"",reviewer:"",signOffDate:""}}function mn(e){return qe[e]||{}}function Ra(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}async function ne(e,t,n){let a=O.steps[e]||{notes:"",reviewer:"",signOffDate:""};a[t]=n,O.steps[e]=a;let o=mn(e),s=t==="notes"?o.note:t==="reviewer"?o.reviewer:o.signOff;if(s&&W())try{await Pe(s,n,ce)}catch(l){console.warn("PTO: unable to save field",s,l)}}async function hn(e,t){O.permanents[e]=t;let n=mn(e);if(n!=null&&n.note&&W())try{await Excel.run(async a=>{var u;let o=a.workbook.tables.getItemOrNullObject(ce[0]);if(await a.sync(),o.isNullObject)return;let s=o.getDataBodyRange(),l=o.getHeaderRowRange();s.load("values"),l.load("values"),await a.sync();let i=l.values[0]||[],r=i.map(d=>String(d||"").trim().toLowerCase()),c={field:r.findIndex(d=>d==="field"||d==="field name"||d==="setting"),permanent:r.findIndex(d=>d==="permanent"||d==="persist"),value:r.findIndex(d=>d==="value"||d==="setting value"),type:r.findIndex(d=>d==="type"||d==="category"),title:r.findIndex(d=>d==="title"||d==="display name")};if(c.field===-1)return;let f=(s.values||[]).findIndex(d=>String(d[c.field]||"").trim()===n.note);if(f>=0)c.permanent>=0&&(s.getCell(f,c.permanent).values=[[t?"Y":"N"]]);else{let d=new Array(i.length).fill("");c.type>=0&&(d[c.type]="Other"),c.title>=0&&(d[c.title]=""),d[c.field]=n.note,c.permanent>=0&&(d[c.permanent]=t?"Y":"N"),c.value>=0&&(d[c.value]=((u=O.steps[e])==null?void 0:u.notes)||""),o.rows.add(null,[d])}await a.sync()})}catch(a){console.warn("PTO: unable to update permanent flag",a)}}async function yn(e,t){let n=ln[e];if(n&&(O.completes[e]=t?"Y":"",!!W()))try{await Pe(n,t?"Y":"",ce)}catch(a){console.warn("PTO: unable to save completion flag",n,a)}}function en(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function Aa(){let e={};return Object.keys(qe).forEach(t=>{var s;let n=parseInt(t,10),a=!!((s=O.steps[n])!=null&&s.signOffDate),o=!!O.completes[n];e[n]=a||o}),e}function vn(e,{buttonId:t,inputId:n,canActivate:a=null,onComplete:o=null}){var r;let s=document.getElementById(t);if(!s)return;let l=document.getElementById(n),i=!!((r=O.steps[e])!=null&&r.signOffDate)||!!O.completes[e];en(s,i),s.addEventListener("click",()=>{if(!s.classList.contains("is-active")&&e>0){let f=Aa(),{canComplete:u,message:d}=qt(e,f);if(!u){Yt(d);return}}if(typeof a=="function"&&!a())return;let p=!s.classList.contains("is-active");en(s,p),l&&(l.value=p?mt():"",ne(e,"signOffDate",l.value)),yn(e,p),p&&window.scrollTo({top:0,behavior:"smooth"}),p&&typeof o=="function"&&o()})}function b(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function Na(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function bn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function Ee(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function gt(e){if(!e)return null;let t=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(e));if(!t)return null;let n=Number(t[1]),a=Number(t[2]),o=Number(t[3]);return!n||!a||!o?null:{year:n,month:a,day:o}}function tn(e){if(!e)return"";let t=gt(e);if(!t)return"";let{year:n,month:a,day:o}=t;return`${n}-${String(a).padStart(2,"0")}-${String(o).padStart(2,"0")}`}function Da(e){let t=gt(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function $a(e){let t=gt(e);return t?`PTO-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function mt(){let e=new Date,t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}function ja(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="y"||t==="yes"||t==="true"||t==="t"||t==="1"}function La(e){if(e instanceof Date)return e.getTime();if(typeof e=="number"){let n=Ba(e);return n?n.getTime():null}let t=new Date(e);return Number.isNaN(t.getTime())?null:t.getTime()}function Ba(e){if(!Number.isFinite(e))return null;let t=new Date(Date.UTC(1899,11,30));return new Date(t.getTime()+e*24*60*60*1e3)}function Ma(){let e=n=>{var a,o;return((o=(a=document.getElementById(n))==null?void 0:a.value)==null?void 0:o.trim())||""};[{id:"config-payroll-date",field:k.payrollDate},{id:"config-accounting-period",field:k.accountingPeriod},{id:"config-journal-id",field:k.journalEntryId},{id:"config-company-name",field:k.companyName},{id:"config-payroll-provider",field:k.payrollProvider},{id:"config-accounting-link",field:k.accountingSoftware},{id:"config-user-name",field:k.reviewerName}].forEach(({id:n,field:a})=>{let o=e(n);a&&le(a,o)})}function $(e,t=[]){let n=t.map(a=>D(a));return e.findIndex(a=>n.some(o=>a.includes(o)))}function Va(e){var C,I,_,j,L,B,Q,U,q;let t=we(2),n=(t==null?void 0:t.notes)||"",a=!!O.permanents[2],o=be(t==null?void 0:t.reviewer),s=(t==null?void 0:t.signOffDate)||"",l=!!(Ee(O.completes[2])||s),i=A.roster||{},r=A.hasAnalyzed,c=(I=(C=A.roster)==null?void 0:C.difference)!=null?I:0,p=!A.skipAnalysis&&Math.abs(c)>0,f=(_=i.rosterCount)!=null?_:0,u=(j=i.payrollCount)!=null?j:0,d=(L=i.difference)!=null?L:u-f,g=Array.isArray(i.mismatches)?i.mismatches.filter(Boolean):[],y="";A.loading?y=((Q=(B=window.PrairieForge)==null?void 0:B.renderStatusBanner)==null?void 0:Q.call(B,{type:"info",message:"Analyzing headcount\u2026",escapeHtml:b}))||"":A.lastError&&(y=((q=(U=window.PrairieForge)==null?void 0:U.renderStatusBanner)==null?void 0:q.call(U,{type:"error",message:A.lastError,escapeHtml:b}))||"");let w=(T,m,v,P)=>{let N=!r,Y;N?Y='<span class="pf-je-check-circle pf-je-circle--pending"></span>':P?Y=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:Y=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let pe=r?` = ${v}`:"";return`
            <div class="pf-je-check-row">
                ${Y}
                <span class="pf-je-check-desc-pill">${b(T)}${pe}</span>
            </div>
        `},h=`
        ${w("SS_Employee_Roster count","Active employees in roster",f,!0)}
        ${w("PTO_Data count","Unique employees in PTO data",u,!0)}
        ${w("Difference","Should be zero",d,d===0)}
    `,E=g.length&&!A.skipAnalysis&&r?window.PrairieForge.renderMismatchTiles({mismatches:g,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"PTO Data",escapeHtml:b}):"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${b(e.title)}</h2>
            <p class="pf-hero-copy">${b(e.summary||"")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${A.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${Gt}
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
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${He}</button>`,"Run")}
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${Re}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${y}
                <div class="pf-je-checks-container">
                    ${h}
                </div>
                ${E}
            </article>
            ${ge({textareaId:"step-notes-input",value:n,permanentId:"step-notes-lock-2",isPermanent:a,hintId:p?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
            ${me({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:l,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
        </section>
    `}function Ha(){let e=ee.completenessCheck||{},t=ee.missingPayRates||[],n=[{key:"accrualRate",label:"Accrual Rate",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"carryOver",label:"Carry Over",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdAccrued",label:"YTD Accrued",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdUsed",label:"YTD Used",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"balance",label:"Balance",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"}],o=n.every(c=>e[c.key]!==null&&e[c.key]!==void 0)&&n.every(c=>{var p;return(p=e[c.key])==null?void 0:p.match}),s=t.length>0,l=c=>{let p=e[c.key],f=p==null,u;return f?u='<span class="pf-je-check-circle pf-je-circle--pending"></span>':p.match?u=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:u=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${u}
                <span class="pf-je-check-desc-pill">${b(c.label)}: ${b(c.desc)}</span>
            </div>
        `},i=n.map(c=>l(c)).join(""),r="";if(s){let c=t[0],p=t.length-1;r=`
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${b(c.name)}</strong> to calculate liability
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
                               data-employee="${Na(c.name)}"
                               data-row="${c.rowIndex}">
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
                ${i}
            </div>
            ${r}
        </article>
    `}function Fa(e){var d,g,y,w,h,E,C,I;let t=we(3),n=!!O.permanents[3],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Ee(O.completes[3])||o),l=K.hasRun,{balanceIssues:i,zeroBalances:r,accrualOutliers:c,totalEmployees:p}=K,f="";if(K.loading)f=((g=(d=window.PrairieForge)==null?void 0:d.renderStatusBanner)==null?void 0:g.call(d,{type:"info",message:"Analyzing data quality...",escapeHtml:b}))||"";else if(l){let _=i.length,j=c.length+r.length;_>0?f=((w=(y=window.PrairieForge)==null?void 0:y.renderStatusBanner)==null?void 0:w.call(y,{type:"error",title:`${_} Balance Issue${_>1?"s":""} Found`,message:"Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",escapeHtml:b}))||"":j>0?f=((E=(h=window.PrairieForge)==null?void 0:h.renderStatusBanner)==null?void 0:E.call(h,{type:"warning",title:"No Critical Issues",message:`${j} informational item${j>1?"s":""} to review (see below).`,escapeHtml:b}))||"":f=((I=(C=window.PrairieForge)==null?void 0:C.renderStatusBanner)==null?void 0:I.call(C,{type:"success",title:"Data Quality Passed",message:`${p} employee${p!==1?"s":""} checked \u2014 no anomalies found.`,escapeHtml:b}))||""}let u=[];return l&&i.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u26A0\uFE0F</span>
                    <span class="pf-quality-issue-title">Balance Issues (${i.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${i.slice(0,5).map(_=>`<li><strong>${b(_.name)}</strong>: ${b(_.issue)}</li>`).join("")}
                    ${i.length>5?`<li class="pf-quality-more">+${i.length-5} more</li>`:""}
                </ul>
            </div>
        `),l&&c.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--warning">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u{1F4CA}</span>
                    <span class="pf-quality-issue-title">High Accrual Rates (${c.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${c.slice(0,5).map(_=>`<li><strong>${b(_.name)}</strong>: ${_.accrualRate.toFixed(2)} hrs/period</li>`).join("")}
                    ${c.length>5?`<li class="pf-quality-more">+${c.length-5} more</li>`:""}
                </ul>
            </div>
        `),l&&r.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--info">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u2139\uFE0F</span>
                    <span class="pf-quality-issue-title">Zero Balances (${r.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${r.slice(0,5).map(_=>`<li><strong>${b(_.name)}</strong></li>`).join("")}
                    ${r.length>5?`<li class="pf-quality-more">+${r.length-5} more</li>`:""}
                </ul>
            </div>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${b(e.title)}</h2>
            <p class="pf-hero-copy">${b(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${f}
                <div class="pf-signoff-action">
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${He}</button>`,"Run")}
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
                        ${K.acknowledged?'<p class="pf-quality-actions-hint"><span class="pf-acknowledged-badge">\u2713 Issues Acknowledged</span></p>':""}
                        <div class="pf-signoff-action">
                            ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${Re}</button>`,"Refresh")}
                            ${K.acknowledged?"":M(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${Ie}</button>`,"Continue")}
                        </div>
                    </div>
                </article>
            `:""}
            ${ge({textareaId:"step-notes-3",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-3",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-3"})}
            ${me({reviewerInputId:"step-reviewer-3",reviewerValue:a,signoffInputId:"step-signoff-3",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-3",completeButtonId:"step-signoff-toggle-3"})}
        </section>
    `}function Ua(e){let t=we(4),n=!!O.permanents[4],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Ee(O.completes[4])||o);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${b(e.title)}</h2>
            <p class="pf-hero-copy">${b(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Analysis</h3>
                    <p class="pf-config-subtext">Calculate liabilities and compare against last period.</p>
                </div>
                <div class="pf-signoff-action">
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${He}</button>`,"Run")}
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${Re}</button>`,"Refresh")}
                </div>
            </article>
            ${Ha()}
            ${ge({textareaId:"step-notes-4",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-4",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-4"})}
            ${me({reviewerInputId:"step-reviewer-4",reviewerValue:a,signoffInputId:"step-signoff-4",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"step-signoff-toggle-4"})}
        </section>
    `}function Ga(e){let t=we(5),n=!!O.permanents[5],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Ee(O.completes[5])||o),l=J.lastError?`<p class="pf-step-note">${b(J.lastError)}</p>`:"",i=J.validationRun,r=J.issues||[],c=[{key:"Debits = Credits",desc:"\u2211 Debit column = \u2211 Credit column"},{key:"Line Amounts Sum to Zero",desc:"\u2211 Line Amount = $0.00"},{key:"JE Matches Analysis Total",desc:"\u2211 Expense line amounts = \u2211 PTO_Analysis Change"}],p=g=>{let y=r.find(E=>E.check===g.key),w=!i,h;return w?h='<span class="pf-je-check-circle pf-je-circle--pending"></span>':y!=null&&y.passed?h=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:h=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${h}
                <span class="pf-je-check-desc-pill">${b(g.desc)}</span>
            </div>
        `},f=c.map(g=>p(g)).join(""),u=r.filter(g=>!g.passed),d="";return i&&u.length>0&&(d=`
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>\u26A0\uFE0F Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${u.map(g=>`<li><strong>${b(g.check)}:</strong> ${b(g.detail)}</li>`).join("")}
                </ul>
            </article>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${b(e.title)}</h2>
            <p class="pf-hero-copy">${b(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${Ve}</button>`,"Generate")}
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${Re}</button>`,"Refresh")}
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${Ft}</button>`,"Export")}
                    ${M(`<button type="button" class="pf-action-toggle pf-clickable" id="je-upload-btn" title="Open accounting software upload">${Ut}</button>`,"Upload")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${l}
                <div class="pf-je-checks-container">
                    ${f}
                </div>
            </article>
            ${d}
            ${ge({textareaId:"step-notes-5",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-5",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-5"})}
            ${me({reviewerInputId:"step-reviewer-5",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
        </section>
    `}function Ja(){var t,n;return Math.abs((n=(t=A.roster)==null?void 0:t.difference)!=null?n:0)>0}function wn(){return!A.skipAnalysis&&Ja()}async function ft(){if(!W()){A.loading=!1,A.lastError="Excel runtime is unavailable.",ie();return}A.loading=!0,A.lastError=null,de(document.getElementById("headcount-save-btn"),!1),ie();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),a=t.workbook.worksheets.getItem("PTO_Data"),o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.getUsedRangeOrNullObject(),l=a.getUsedRangeOrNullObject();s.load("values"),l.load("values"),o.load("isNullObject"),await t.sync();let i=null;o.isNullObject||(i=o.getUsedRangeOrNullObject(),i.load("values")),await t.sync();let r=s.isNullObject?[]:s.values||[],c=l.isNullObject?[]:l.values||[],p=i&&!i.isNullObject?i.values||[]:[],f=p.length?p:c;return za(r,f)});A.roster=e.roster,A.hasAnalyzed=!0,A.lastError=null}catch(e){console.warn("PTO headcount: unable to analyze data",e),A.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{A.loading=!1,ie()}}function nn(e){if(!e)return!0;let t=e.toLowerCase().trim();return t?["total","subtotal","sum","count","grand","average","avg"].some(a=>t.includes(a)):!0}function za(e,t){let n={rosterCount:0,payrollCount:0,difference:0,mismatches:[]};if(((e==null?void 0:e.length)||0)<2||((t==null?void 0:t.length)||0)<2)return console.warn("Headcount: insufficient data rows",{rosterRows:(e==null?void 0:e.length)||0,payrollRows:(t==null?void 0:t.length)||0}),{roster:n};let a=an(e),o=an(t),s=a.headers,l=o.headers,i={employee:on(s),termination:s.findIndex(d=>d.includes("termination"))},r={employee:on(l)};console.log("Headcount column detection:",{rosterEmployeeCol:i.employee,rosterTerminationCol:i.termination,payrollEmployeeCol:r.employee,rosterHeaders:s.slice(0,5),payrollHeaders:l.slice(0,5)});let c=new Set,p=new Set;for(let d=a.startIndex;d<e.length;d+=1){let g=e[d],y=i.employee>=0?ve(g[i.employee]):"";nn(y)||i.termination>=0&&ve(g[i.termination])||c.add(y.toLowerCase())}for(let d=o.startIndex;d<t.length;d+=1){let g=t[d],y=r.employee>=0?ve(g[r.employee]):"";nn(y)||p.add(y.toLowerCase())}n.rosterCount=c.size,n.payrollCount=p.size,n.difference=n.payrollCount-n.rosterCount,console.log("Headcount results:",{rosterCount:n.rosterCount,payrollCount:n.payrollCount,difference:n.difference});let f=[...c].filter(d=>!p.has(d)),u=[...p].filter(d=>!c.has(d));return n.mismatches=[...f.map(d=>`In roster, missing in PTO_Data: ${d}`),...u.map(d=>`In PTO_Data, missing in roster: ${d}`)],{roster:n}}function an(e){if(!Array.isArray(e)||!e.length)return{headers:[],startIndex:1};let t=e.findIndex((o=[])=>o.some(s=>ve(s).toLowerCase().includes("employee"))),n=t===-1?0:t;return{headers:(e[n]||[]).map(o=>ve(o).toLowerCase()),startIndex:n+1}}function on(e=[]){let t=-1,n=-1;return e.forEach((a,o)=>{let s=a.toLowerCase();if(!s.includes("employee"))return;let l=1;s.includes("name")?l=4:s.includes("id")?l=2:l=3,l>n&&(n=l,t=o)}),t}function ve(e){return e==null?"":String(e).trim()}async function kn(e=null){let t=async n=>{let a=n.workbook.worksheets.getItem("PTO_Data"),o=n.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster"),l=n.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary"),i=n.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary"),r=a.getUsedRangeOrNullObject();r.load("values"),o.load("isNullObject"),s.load("isNullObject"),l.load("isNullObject"),i.load("isNullObject"),await n.sync();let c=r.isNullObject?[]:r.values||[];if(!c.length)return;let p=(c[0]||[]).map(S=>D(S)),f=p.findIndex(S=>S.includes("employee")&&S.includes("name")),u=f>=0?f:0,d=$(p,["accrual rate"]),g=$(p,["carry over","carryover"]),y=p.findIndex(S=>S.includes("ytd")&&(S.includes("accrued")||S.includes("accrual"))),w=p.findIndex(S=>S.includes("ytd")&&S.includes("used")),h=$(p,["balance","current balance","pto balance"]);console.log("[PTO Analysis] PTO_Data headers:",p),console.log("[PTO Analysis] Column indices found:",{employee:u,accrualRate:d,carryOver:g,ytdAccrued:y,ytdUsed:w,balance:h}),w>=0?console.log(`[PTO Analysis] YTD Used column: "${p[w]}" at index ${w}`):console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:",p);let E=c.slice(1).map(S=>ve(S[u])).filter(S=>S&&!S.toLowerCase().includes("total")),C=new Map;c.slice(1).forEach(S=>{let G=D(S[u]);!G||G.includes("total")||C.set(G,S)});let I=new Map;if(s.isNullObject)console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");else{let S=s.getUsedRangeOrNullObject();S.load("values"),await n.sync();let G=S.isNullObject?[]:S.values||[];if(G.length){let V=(G[0]||[]).map(x=>D(x));console.log("[PTO Analysis] SS_Employee_Roster headers:",V);let H=V.findIndex(x=>x.includes("employee")&&x.includes("name"));H<0&&(H=V.findIndex(x=>x==="employee"||x==="name"||x==="full name"));let F=V.findIndex(x=>x.includes("department"));console.log(`[PTO Analysis] Roster column indices - Name: ${H}, Dept: ${F}`),H>=0&&F>=0?(G.slice(1).forEach(x=>{let re=D(x[H]),fe=ve(x[F]);re&&I.set(re,fe)}),console.log(`[PTO Analysis] Built roster map with ${I.size} employees`)):console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster")}}let _=new Map;if(!l.isNullObject){let S=l.getUsedRangeOrNullObject();S.load("values"),await n.sync();let G=S.isNullObject?[]:S.values||[];if(G.length){let V=(G[0]||[]).map(F=>D(F)),H={payrollDate:$(V,["payroll date"]),employee:$(V,["employee"]),category:$(V,["payroll category","category"]),amount:$(V,["amount","gross salary","gross_salary","earnings"])};H.employee>=0&&H.category>=0&&H.amount>=0&&G.slice(1).forEach(F=>{let x=D(F[H.employee]);if(!x)return;let re=D(F[H.category]);if(!re.includes("regular")||!re.includes("earn"))return;let fe=Number(F[H.amount])||0;if(!fe)return;let Ce=La(F[H.payrollDate]),_e=_.get(x);(!_e||Ce!=null&&Ce>_e.timestamp)&&_.set(x,{payRate:fe/80,timestamp:Ce})})}}let j=new Map;if(!i.isNullObject){let S=i.getUsedRangeOrNullObject();S.load("values"),await n.sync();let G=S.isNullObject?[]:S.values||[];if(G.length>1){let V=(G[0]||[]).map(x=>D(x)),H=V.findIndex(x=>x.includes("employee")&&x.includes("name")),F=$(V,["liability amount","liability","accrued pto"]);H>=0&&F>=0&&G.slice(1).forEach(x=>{let re=D(x[H]);if(!re)return;let fe=Number(x[F])||0;j.set(re,fe)})}}let L=te(k.payrollDate)||"",B=[],Q=[],U=E.map((S,G)=>{var vt,bt,wt,kt,Ot,St,xt;let V=D(S),H=I.get(V)||"",F=(bt=(vt=_.get(V))==null?void 0:vt.payRate)!=null?bt:"",x=C.get(V),re=x&&d>=0&&(wt=x[d])!=null?wt:"",fe=x&&g>=0&&(kt=x[g])!=null?kt:"",Ce=x&&y>=0&&(Ot=x[y])!=null?Ot:"",_e=x&&w>=0&&(St=x[w])!=null?St:"";(V.includes("avalos")||V.includes("sarah"))&&console.log(`[PTO Debug] ${S}:`,{ytdUsedIdx:w,rawValue:x?x[w]:"no dataRow",ytdUsed:_e,fullRow:x});let Ye=x&&h>=0&&Number(x[h])||0,ht=G+2;!F&&typeof F!="number"&&B.push({name:S,rowIndex:ht}),H||Q.push({name:S,rowIndex:ht});let We=typeof F=="number"&&Ye?Ye*F:0,yt=(xt=j.get(V))!=null?xt:0,On=(typeof We=="number"?We:0)-yt;return[L,S,H,F,re,fe,Ce,_e,Ye,We,yt,On]});ee.missingPayRates=B.filter(S=>!ee.ignoredMissingPayRates.has(S.name)),ee.missingDepartments=Q,console.log(`[PTO Analysis] Data quality: ${B.length} missing pay rates, ${Q.length} missing departments`);let q=[["Analysis Date","Employee Name","Department","Pay Rate","Accrual Rate","Carry Over","YTD Accrued","YTD Used","Balance","Liability Amount","Accrued PTO $ [Prior Period]","Change"],...U],T=o.isNullObject?n.workbook.worksheets.add("PTO_Analysis"):o,m=T.getUsedRangeOrNullObject();m.load("address"),await n.sync(),m.isNullObject||m.clear();let v=q[0].length,P=q.length,N=U.length,Y=T.getRangeByIndexes(0,0,P,v);Y.values=q;let pe=T.getRangeByIndexes(0,0,1,v);dt(pe),N>0&&(Wt(T,0,N),he(T,3,N),Oe(T,4,N),Oe(T,5,N),Oe(T,6,N),Oe(T,7,N),Oe(T,8,N),he(T,9,N),he(T,10,N),he(T,11,N,!0)),Y.format.autofitColumns(),T.getRange("A1").select(),await n.sync()};W()&&(e?await t(e):await Excel.run(t))}function qa(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let a=String(n);return/[",\n]/.test(a)?`"${a.replace(/"/g,'""')}"`:a}).join(",")).join(`
`)}function Ya(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),a=URL.createObjectURL(n),o=document.createElement("a");o.href=a,o.download=e,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(a),1e3)}function sn(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=wn(),n=document.getElementById("step-notes-input"),a=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!a;let o=document.getElementById("headcount-notes-hint");o&&(o.textContent=t?"Please document outstanding differences before signing off.":"")}function rn(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(ye)?t.slice(ye.length).replace(/^\s+/,""):t.replace(new RegExp(`^${ye}\\s*`,"i"),"").trimStart(),a=ye+(n?`
${n}`:"");e.value!==a&&(e.value=a),ne(2,"notes",e.value)}function Wa(){let e=document.getElementById("step-notes-input");e&&e.addEventListener("input",()=>{if(!A.skipAnalysis)return;let t=e.value||"";if(!t.startsWith(ye)){let n=t.replace(ye,"").trimStart();e.value=ye+(n?`
${n}`:"")}ne(2,"notes",e.value)})}})();
//# sourceMappingURL=app.bundle.js.map
