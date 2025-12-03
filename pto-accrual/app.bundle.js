/* Prairie Forge PTO Accrual */
(()=>{function Y(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}var Ke="SS_PF_Config";async function Et(e,t=[Ke]){var o;let n=e.workbook.tables;n.load("items/name"),await e.sync();let a=(o=n.items)==null?void 0:o.find(s=>t.includes(s.name));return a?e.workbook.tables.getItem(a.name):(console.warn("Config table not found. Looking for:",t),null)}function Ct(e){let t=e.map(n=>String(n||"").trim().toLowerCase());return{field:t.findIndex(n=>n==="field"||n==="field name"||n==="setting"),value:t.findIndex(n=>n==="value"||n==="setting value"),type:t.findIndex(n=>n==="type"||n==="category"),title:t.findIndex(n=>n==="title"||n==="display name"),permanent:t.findIndex(n=>n==="permanent"||n==="persist")}}async function Qe(e=[Ke]){if(!Y())return{};try{return await Excel.run(async t=>{let n=await Et(t,e);if(!n)return{};let a=n.getDataBodyRange(),o=n.getHeaderRowRange();a.load("values"),o.load("values"),await t.sync();let s=o.values[0]||[],r=Ct(s);if(r.field===-1||r.value===-1)return console.warn("Config table missing FIELD or VALUE columns. Headers:",s),{};let i={};return(a.values||[]).forEach(d=>{var f;let p=String(d[r.field]||"").trim();p&&(i[p]=(f=d[r.value])!=null?f:"")}),console.log("Configuration loaded:",Object.keys(i).length,"fields"),i})}catch(t){return console.error("Failed to load configuration:",t),{}}}async function _e(e,t,n=[Ke]){if(!Y())return!1;try{return await Excel.run(async a=>{let o=await Et(a,n);if(!o){console.warn("Config table not found for write");return}let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let i=r.values[0]||[],c=Ct(i);if(c.field===-1||c.value===-1){console.error("Config table missing FIELD or VALUE columns");return}let p=(s.values||[]).findIndex(f=>String(f[c.field]||"").trim()===e);if(p>=0)s.getCell(p,c.value).values=[[t]];else{let f=new Array(i.length).fill("");c.type>=0&&(f[c.type]="Run Settings"),f[c.field]=e,f[c.value]=t,c.permanent>=0&&(f[c.permanent]="N"),c.title>=0&&(f[c.title]=""),o.rows.add(null,[f]),console.log("Added new config row:",e,"=",t)}await a.sync(),console.log("Saved config:",e,"=",t)}),!0}catch(a){return console.error("Failed to save config:",e,a),!1}}var Sn="SS_PF_Config",xn="module-prefix",Xe="system",ke={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function _t(){if(!Y())return{...ke};try{return await Excel.run(async e=>{var p,f;let t=e.workbook.worksheets.getItemOrNullObject(Sn);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...ke};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((p=n.values)!=null&&p.length))return{...ke};let a=n.values,o=_n(a[0]),s=o.get("category"),r=o.get("field"),i=o.get("value");if(s===void 0||r===void 0||i===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...ke};let c={},d=!1;for(let u=1;u<a.length;u++){let l=a[u];if(je(l[s])===xn){let y=String((f=l[r])!=null?f:"").trim().toUpperCase(),w=je(l[i]);y&&w&&(c[y]=w,d=!0)}}return d?(console.log("[Tab Visibility] Loaded prefix config:",c),c):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...ke})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...ke}}}async function Ze(e){if(!Y())return;let t=je(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await _t();await Excel.run(async a=>{let o=a.workbook.worksheets;o.load("items/name,visibility"),await a.sync();let s={};for(let[u,l]of Object.entries(n))s[l]||(s[l]=[]),s[l].push(u);let r=s[t]||[],i=s[Xe]||[],c=[];for(let[u,l]of Object.entries(s))u!==t&&u!==Xe&&c.push(...l);console.log(`[Tab Visibility] Active prefixes: ${r.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${c.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${i.join(", ")}`);let d=[],p=[];o.items.forEach(u=>{let l=u.name,g=l.toUpperCase(),y=r.some(E=>g.startsWith(E)),w=c.some(E=>g.startsWith(E)),h=i.some(E=>g.startsWith(E));y?(d.push(u),console.log(`[Tab Visibility] SHOW: ${l} (matches active module prefix)`)):h?(p.push(u),console.log(`[Tab Visibility] HIDE: ${l} (system sheet)`)):w?(p.push(u),console.log(`[Tab Visibility] HIDE: ${l} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${l} (no prefix match, leaving as-is)`)});for(let u of d)u.visibility=Excel.SheetVisibility.visible;if(await a.sync(),o.items.filter(u=>u.visibility===Excel.SheetVisibility.visible).length>p.length){for(let u of p)try{u.visibility=Excel.SheetVisibility.hidden}catch(l){console.warn(`[Tab Visibility] Could not hide "${u.name}":`,l.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${d.length}, hid ${p.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function En(){if(!Y()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.visibility!==Excel.SheetVisibility.visible&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${a.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function Cn(){if(!Y()){console.log("Excel not available");return}try{let e=await _t(),t=[];for(let[n,a]of Object.entries(e))a===Xe&&t.push(n);await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name,visibility"),await n.sync(),a.items.forEach(o=>{let s=o.name.toUpperCase();t.some(r=>s.startsWith(r))&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function _n(e=[]){let t=new Map;return e.forEach((n,a)=>{let o=je(n);o&&t.set(o,a)}),t}function je(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=En,window.PrairieForge.unhideSystemSheets=Cn,window.PrairieForge.applyModuleTabVisibility=Ze);var Pt={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var Rt=Pt.ADA_IMAGE_URL;async function Le(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async a=>{let o=a.workbook.worksheets.getItemOrNullObject(e);o.load("isNullObject, name, visibility"),await a.sync();let s;o.isNullObject?(s=a.workbook.worksheets.add(e),await a.sync(),await Tt(a,s,t,n)):(s=o,s.visibility!==Excel.SheetVisibility.visible&&(s.visibility=Excel.SheetVisibility.visible,await a.sync()),await Tt(a,s,t,n)),s.activate(),s.getRange("A1").select(),await a.sync()})}catch(a){console.error(`Error activating homepage sheet ${e}:`,a)}}async function Tt(e,t,n,a){try{let d=t.getUsedRangeOrNullObject();d.load("isNullObject"),await e.sync(),d.isNullObject||(d.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let o=[[n,""],[a,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=o;let r=t.getRange("A1:Z100");r.format.fill.color="#0f0f0f";let i=t.getRange("A1");i.format.font.bold=!0,i.format.font.size=36,i.format.font.color="#ffffff",i.format.font.name="Segoe UI Light",i.format.verticalAlignment="Center";let c=t.getRange("A2");c.format.font.size=14,c.format.font.color="#a0a0a0",c.format.font.name="Segoe UI",c.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var It={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function Be(e){return It[e]||It["module-selector"]}function At(){tt();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",et),t.addEventListener("click",o=>{o.target===t&&et()});let a=o=>{o.key==="Escape"&&(et(),document.removeEventListener("keydown",a))};document.addEventListener("keydown",a)}function et(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var Tn=["January","February","March","April","May","June","July","August","September","October","November","December"],Dt=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],In=["Su","Mo","Tu","We","Th","Fr","Sa"],X=null,oe=null;function $t(e,t={}){let n=document.getElementById(e);if(!n)return;let{onChange:a=null,minDate:o=null,maxDate:s=null,readonly:r=!1}=t,i=n.closest(".pf-datepicker-wrapper");i||(i=document.createElement("div"),i.className="pf-datepicker-wrapper",n.parentNode.insertBefore(i,n),i.appendChild(n)),n.type="text",n.placeholder="Select date...",n.classList.add("pf-datepicker-input"),n.readOnly=!0;let c=n.value?Nt(n.value):null;c&&(n.value=at(c),n.dataset.value=Pe(c));let d=i.querySelector(".pf-datepicker-icon");d||(d=document.createElement("span"),d.className="pf-datepicker-icon",d.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>',i.appendChild(d));let p={inputId:e,input:n,selectedDate:c,viewDate:c?new Date(c):new Date,onChange:a,minDate:o,maxDate:s};function f(){r||(oe=p,Rn())}return n.addEventListener("click",f),d.addEventListener("click",f),{getValue:()=>p.selectedDate?Pe(p.selectedDate):"",setValue:u=>{let l=Nt(u);p.selectedDate=l,p.viewDate=l?new Date(l):new Date,l?(n.value=at(l),n.dataset.value=Pe(l)):(n.value="",n.dataset.value="")},open:f,close:Me}}function Rn(){oe&&(X||(X=document.createElement("div"),X.className="pf-datepicker-modal",X.id="pf-datepicker-modal",document.body.appendChild(X)),Lt(),requestAnimationFrame(()=>{X.classList.add("is-open")}),document.addEventListener("keydown",jt))}function Me(){X&&X.classList.remove("is-open"),document.removeEventListener("keydown",jt),oe=null}function jt(e){e.key==="Escape"&&Me()}function Lt(){if(!X||!oe)return;let{viewDate:e,selectedDate:t,minDate:n,maxDate:a}=oe,o=e.getFullYear(),s=e.getMonth();X.innerHTML=`
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
                ${In.map(r=>`<span>${r}</span>`).join("")}
            </div>
            <div class="pf-datepicker-days">
                ${Dn(o,s,t,n,a)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-btn pf-datepicker-today" data-action="today">Today</button>
                <button type="button" class="pf-datepicker-btn pf-datepicker-clear" data-action="clear">Clear</button>
            </div>
        </div>
    `,An()}function An(){var e;X&&((e=X.querySelector(".pf-datepicker-backdrop"))==null||e.addEventListener("click",Me),X.querySelectorAll(".pf-datepicker-nav").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=t.dataset.action;Nn(a)})}),X.querySelectorAll(".pf-datepicker-day:not(.disabled)").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=parseInt(t.dataset.day),o=parseInt(t.dataset.month),s=parseInt(t.dataset.year);nt(new Date(s,o,a))})}),X.querySelectorAll(".pf-datepicker-btn").forEach(t=>{t.addEventListener("click",n=>{n.preventDefault(),n.stopPropagation();let a=t.dataset.action;a==="today"?nt(new Date):a==="clear"&&nt(null)})}))}function Nn(e){if(!oe)return;let t=oe.viewDate;switch(e){case"prev-year":t.setFullYear(t.getFullYear()-1);break;case"prev-month":t.setMonth(t.getMonth()-1);break;case"next-month":t.setMonth(t.getMonth()+1);break;case"next-year":t.setFullYear(t.getFullYear()+1);break}Lt()}function nt(e){if(!oe)return;let{input:t,onChange:n}=oe;oe.selectedDate=e,e?(t.value=at(e),t.dataset.value=Pe(e),oe.viewDate=new Date(e)):(t.value="",t.dataset.value=""),n&&n(e?Pe(e):""),t.dispatchEvent(new Event("change",{bubbles:!0})),Me()}function Dn(e,t,n,a,o){let s=new Date(e,t,1).getDay(),r=new Date(e,t+1,0).getDate(),i=new Date(e,t,0).getDate(),c=new Date;c.setHours(0,0,0,0),n&&(n=new Date(n),n.setHours(0,0,0,0));let d="";for(let l=s-1;l>=0;l--){let g=i-l,y=t===0?11:t-1,w=t===0?e-1:e;d+=`<button type="button" class="pf-datepicker-day other-month" data-day="${g}" data-month="${y}" data-year="${w}">${g}</button>`}for(let l=1;l<=r;l++){let g=new Date(e,t,l);g.setHours(0,0,0,0);let y=g.getTime()===c.getTime(),w=n&&g.getTime()===n.getTime(),h="pf-datepicker-day";y&&(h+=" today"),w&&(h+=" selected");let E=!1;a&&g<a&&(E=!0),o&&g>o&&(E=!0),E&&(h+=" disabled"),d+=`<button type="button" class="${h}" data-day="${l}" data-month="${t}" data-year="${e}" ${E?"disabled":""}>${l}</button>`}let p=42,f=s+r,u=p-f;for(let l=1;l<=u;l++){let g=t===11?0:t+1,y=t===11?e+1:e;d+=`<button type="button" class="pf-datepicker-day other-month" data-day="${l}" data-month="${g}" data-year="${y}">${l}</button>`}return d}function Nt(e){if(!e)return null;if(/^\d{4}-\d{2}-\d{2}$/.test(e)){let[a,o,s]=e.split("-").map(Number);return new Date(a,o-1,s)}let t=e.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);if(t){let a=Dt.findIndex(o=>o.toLowerCase()===t[1].toLowerCase().substring(0,3));if(a>=0)return new Date(parseInt(t[3]),a,parseInt(t[2]))}if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(e)){let[a,o,s]=e.split("/").map(Number);return new Date(s,a-1,o)}let n=new Date(e);return isNaN(n.getTime())?null:n}function at(e){return e?`${Dt[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`:""}function Pe(e){if(!e)return"";let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}var Bt=`
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
`.trim(),Ya=`
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
`.trim(),Wa=`
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
`.trim(),Te=`
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
`.trim(),Ka=`
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
        <circle cx="12" cy="12" r="10"/>
        <path d="m9 12 2 2 4-4"/>
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
        <line x1="12" x2="12" y1="8" y2="12"/>
        <line x1="12" x2="12.01" y1="16" y2="16"/>
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
        <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>
        <path d="M12 9v4"/>
        <path d="M12 17h.01"/>
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
        <path d="M12 16v-4"/>
        <path d="M12 8h.01"/>
    </svg>
`.trim(),to=`
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
`.trim(),no=`
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
`.trim(),ao=`
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
`.trim(),oo=`
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
`.trim();function Re(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function B(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function fe({textareaId:e,value:t,permanentId:n,isPermanent:a,hintId:o,saveButtonId:s,isSaved:r=!1,placeholder:i="Enter notes here..."}){let c=a?st:ot,d=s?`<button type="button" class="pf-action-toggle pf-save-btn ${r?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Jt}</button>`:"",p=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${a?"is-locked":""}" id="${n}" aria-pressed="${a}" title="Lock notes (retain after archive)">${c}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Re(i)}">${Re(t||"")}</textarea>
                ${o?`<p class="pf-signoff-hint" id="${o}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?B(p,"Lock"):""}
                ${s?B(d,"Save"):""}
            </div>
        </article>
    `}function ge({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:a,isComplete:o,saveButtonId:s,isSaved:r=!1,completeButtonId:i,subtext:c="Sign-off below. Click checkmark icon. Done.",prevButtonId:d=null,nextButtonId:p=null}){let f=d||`${i}-prev`,u=p||`${i}-next`,l=`<button type="button" class="pf-action-toggle ${o?"is-active":""}" id="${i}" aria-pressed="${!!o}" title="Mark step complete">${Te}</button>`,g=`<button type="button" class="pf-action-toggle pf-nav-toggle" id="${f}" title="Previous step">${Fe}</button>`,y=`<button type="button" class="pf-action-toggle pf-nav-toggle" id="${u}" title="Next step">${Ue}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${Re(c)}</p>
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
                ${g?B(g,"Prev"):""}
                ${B(l,"Done")}
                ${y?B(y,"Next"):""}
            </div>
        </article>
    `}function rt(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?st:ot)}function ce(e,t){e&&e.classList.toggle("is-saved",t)}function lt(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(a=>{let o=a.getAttribute("data-save-input"),s=document.getElementById(o);if(!s)return;let r=()=>{ce(a,!1)};s.addEventListener("input",r),n.push(()=>s.removeEventListener("input",r))}),()=>n.forEach(a=>a())}function qt(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function Yt(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var ct={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},Ge={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function dt(e){e.format.fill.color=ct.fillColor,e.format.font.color=ct.fontColor,e.format.font.bold=ct.bold}function me(e,t,n,a=!1){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a?Ge.currencyWithNegative:Ge.currency]]}function Oe(e,t,n){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[Ge.number]]}function Wt(e,t,n,a=Ge.date){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a]]}var jn="c2b49ed",Ne="pto-accrual";var he="PTO Accrual";function K(e,t="info",n=4e3){document.querySelectorAll(".pf-toast").forEach(o=>o.remove());let a=document.createElement("div");if(a.className=`pf-toast pf-toast--${t}`,a.innerHTML=`
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
        `,document.head.appendChild(o)}return document.body.appendChild(a),n>0&&setTimeout(()=>a.remove(),n),a}function Ln(e,t={}){let{title:n="Confirm Action",confirmText:a="Continue",cancelText:o="Cancel",icon:s="\u{1F4CB}",destructive:r=!1}=t;return new Promise(i=>{document.querySelectorAll(".pf-confirm-overlay").forEach(d=>d.remove());let c=document.createElement("div");if(c.className="pf-confirm-overlay",c.innerHTML=`
            <div class="pf-confirm-dialog">
                <div class="pf-confirm-icon">${s}</div>
                <div class="pf-confirm-title">${n}</div>
                <div class="pf-confirm-message">${e.replace(/\n/g,"<br>")}</div>
                <div class="pf-confirm-buttons">
                    <button class="pf-confirm-btn pf-confirm-btn--cancel">${o}</button>
                    <button class="pf-confirm-btn pf-confirm-btn--ok ${r?"pf-confirm-btn--destructive":""}">${a}</button>
                </div>
            </div>
        `,!document.getElementById("pf-confirm-styles")){let d=document.createElement("style");d.id="pf-confirm-styles",d.textContent=`
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
            `,document.head.appendChild(d)}document.body.appendChild(c),c.addEventListener("click",d=>{d.target===c&&(c.remove(),i(!1))}),c.querySelector(".pf-confirm-btn--cancel").onclick=()=>{c.remove(),i(!1)},c.querySelector(".pf-confirm-btn--ok").onclick=()=>{c.remove(),i(!0)}})}var Bn="Calculate your PTO liability, compare against last period, and generate a balanced journal entry\u2014all without leaving Excel.",Mn="../module-selector/index.html",Vn="pf-loader-overlay",de=["SS_PF_Config"],k={payrollProvider:"PTO_Payroll_Provider",payrollDate:"PTO_Analysis_Date",accountingPeriod:"PTO_Accounting_Period",journalEntryId:"PTO_Journal_Entry_ID",companyName:"SS_Company_Name",accountingSoftware:"SS_Accounting_Software",reviewerName:"PTO_Reviewer",validationDataBalance:"PTO_Validation_Data_Balance",validationCleanBalance:"PTO_Validation_Clean_Balance",validationDifference:"PTO_Validation_Difference",headcountRosterCount:"PTO_Headcount_Roster_Count",headcountPayrollCount:"PTO_Headcount_Payroll_Count",headcountDifference:"PTO_Headcount_Difference",journalDebitTotal:"PTO_JE_Debit_Total",journalCreditTotal:"PTO_JE_Credit_Total",journalDifference:"PTO_JE_Difference"},ye="User opted to skip the headcount review this period.",qe={0:{note:"PTO_Notes_Config",reviewer:"PTO_Reviewer_Config",signOff:"PTO_SignOff_Config"},1:{note:"PTO_Notes_Import",reviewer:"PTO_Reviewer_Import",signOff:"PTO_SignOff_Import"},2:{note:"PTO_Notes_Headcount",reviewer:"PTO_Reviewer_Headcount",signOff:"PTO_SignOff_Headcount"},3:{note:"PTO_Notes_Validate",reviewer:"PTO_Reviewer_Validate",signOff:"PTO_SignOff_Validate"},4:{note:"PTO_Notes_Review",reviewer:"PTO_Reviewer_Review",signOff:"PTO_SignOff_Review"},5:{note:"PTO_Notes_JE",reviewer:"PTO_Reviewer_JE",signOff:"PTO_SignOff_JE"},6:{note:"PTO_Notes_Archive",reviewer:"PTO_Reviewer_Archive",signOff:"PTO_SignOff_Archive"}},ln={0:"PTO_Complete_Config",1:"PTO_Complete_Import",2:"PTO_Complete_Headcount",3:"PTO_Complete_Validate",4:"PTO_Complete_Review",5:"PTO_Complete_JE",6:"PTO_Complete_Archive"};var ae=[{id:0,title:"Configuration",summary:"Set the analysis date, accounting period, and review details for this run.",description:"Complete this step first to ensure all downstream calculations use the correct period settings.",actionLabel:"Configure Workbook",secondaryAction:{sheet:"SS_PF_Config",label:"Open Config Sheet"}},{id:1,title:"Import PTO Data",summary:"Pull your latest PTO export from payroll and paste it into PTO_Data.",description:"Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",actionLabel:"Import Sample Data",secondaryAction:{sheet:"PTO_Data",label:"Open Data Sheet"}},{id:2,title:"Headcount Review",summary:"Quick check to make sure your roster matches your PTO data.",description:"Compare employees in PTO_Data against your employee roster to catch any discrepancies.",actionLabel:"Open Headcount Review",secondaryAction:{sheet:"SS_Employee_Roster",label:"Open Sheet"}},{id:3,title:"Data Quality Review",summary:"Scan your PTO data for potential errors before crunching numbers.",description:"Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",actionLabel:"Click to Run Quality Check"},{id:4,title:"PTO Accrual Review",summary:"Review the calculated liability for each employee and compare to last period.",description:"The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",actionLabel:"Click to Perform Review"},{id:5,title:"Journal Entry Prep",summary:"Generate a balanced journal entry, run validation checks, and export when ready.",description:"Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",actionLabel:"Open Journal Draft",secondaryAction:{sheet:"PTO_JE_Draft",label:"Open Sheet"}},{id:6,title:"Archive & Reset",summary:"Save this period's results and prepare for the next cycle.",description:"Archive the current analysis so it becomes the 'prior period' for your next review.",actionLabel:"Archive Run"}],Hn={0:"PTO_Homepage",1:"PTO_Data",2:"PTO_Data",3:"PTO_Analysis",4:"PTO_Analysis",5:"PTO_JE_Draft"},Fn={PTO_Homepage:0,PTO_Data:1,PTO_Analysis:4,PTO_JE_Draft:5,PTO_Archive_Summary:6,SS_PF_Config:0,SS_Employee_Roster:2};var Un=ae.reduce((e,t)=>(e[t.id]="pending",e),{}),N={activeView:"home",activeStepId:null,focusedIndex:0,stepStatuses:Un},S={loaded:!1,steps:{},permanents:{},completes:{},values:{},overrides:{accountingPeriod:!1,journalId:!1}},Ae=null,ut=null,Je=null,Se=new Map,R={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},G={debitTotal:null,creditTotal:null,difference:null,lineAmountSum:null,analysisChangeTotal:null,jeChangeTotal:null,loading:!1,lastError:null,validationRun:!1,issues:[]},W={hasRun:!1,loading:!1,acknowledged:!1,balanceIssues:[],zeroBalances:[],accrualOutliers:[],totalIssues:0,totalEmployees:0},Z={cleanDataReady:!1,employeeCount:0,lastRun:null,loading:!1,lastError:null,missingPayRates:[],missingDepartments:[],ignoredMissingPayRates:new Set,completenessCheck:{accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null}};async function Gn(){var e;try{Ae=document.getElementById("app"),ut=document.getElementById("loading"),await qn(),await Yn(),(e=window.PrairieForge)!=null&&e.loadSharedConfig&&await window.PrairieForge.loadSharedConfig();let t=Be(Ne);await Le(t.sheetName,t.title,t.subtitle),await Jn(),ut&&ut.remove(),Ae&&(Ae.hidden=!1),se()}catch(t){throw console.error("[PTO] Module initialization failed:",t),t}}async function Jn(){if(ie())try{await Excel.run(async e=>{e.workbook.worksheets.onActivated.add(zn),await e.sync(),console.log("[PTO] Worksheet change listener registered")})}catch(e){console.warn("[PTO] Could not set up worksheet listener:",e)}}async function zn(e){try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e.worksheetId);n.load("name"),await t.sync();let a=n.name,o=Fn[a];if(console.log(`[PTO] Tab changed to: ${a} \u2192 Step ${o}`),o!==void 0&&o!==N.activeStepId){let s=STEPS.findIndex(r=>r.id===o);if(s>=0){let r=o===0?"config":"step";N.activeView=r,N.activeStepId=o,N.focusedIndex=s,se()}}})}catch(t){console.warn("[PTO] Error handling worksheet change:",t)}}async function qn(){try{await Ze(Ne),console.log(`[PTO] Tab visibility applied for ${Ne}`)}catch(e){console.warn("[PTO] Could not apply tab visibility:",e)}}async function Yn(){var e;if(!Y()){S.loaded=!0;return}try{let t=await Qe(de),n={};(e=window.PrairieForge)!=null&&e.loadSharedConfig&&(await window.PrairieForge.loadSharedConfig(),window.PrairieForge._sharedConfigCache&&window.PrairieForge._sharedConfigCache.forEach((s,r)=>{n[r]=s}));let a={...t},o={SS_Default_Reviewer:k.reviewerName,Default_Reviewer:k.reviewerName,PTO_Reviewer:k.reviewerName,SS_Company_Name:k.companyName,Company_Name:k.companyName,SS_Payroll_Provider:k.payrollProvider,Payroll_Provider_Link:k.payrollProvider,SS_Accounting_Software:k.accountingSoftware,Accounting_Software_Link:k.accountingSoftware};Object.entries(o).forEach(([s,r])=>{n[s]&&!a[r]&&(a[r]=n[s])}),Object.entries(n).forEach(([s,r])=>{s.startsWith("PTO_")&&r&&(a[s]=r)}),S.permanents=await Wn(),S.values=a||{},S.overrides.accountingPeriod=!!(a!=null&&a[k.accountingPeriod]),S.overrides.journalId=!!(a!=null&&a[k.journalEntryId]),Object.entries(qe).forEach(([s,r])=>{var i,c,d;S.steps[s]={notes:(i=a[r.note])!=null?i:"",reviewer:(c=a[r.reviewer])!=null?c:"",signOffDate:(d=a[r.signOff])!=null?d:""}}),S.completes=Object.entries(ln).reduce((s,[r,i])=>{var c;return s[r]=(c=a[i])!=null?c:"",s},{}),S.loaded=!0}catch(t){console.warn("PTO: unable to load configuration fields",t),S.loaded=!0}}async function Wn(){let e={};if(!Y())return e;let t=new Map;Object.entries(qe).forEach(([n,a])=>{a.note&&t.set(a.note.trim(),Number(n))});try{await Excel.run(async n=>{let a=n.workbook.tables.getItemOrNullObject(de[0]);if(await n.sync(),a.isNullObject)return;let o=a.getDataBodyRange(),s=a.getHeaderRowRange();o.load("values"),s.load("values"),await n.sync();let i=(s.values[0]||[]).map(d=>String(d||"").trim().toLowerCase()),c={field:i.findIndex(d=>d==="field"||d==="field name"||d==="setting"),permanent:i.findIndex(d=>d==="permanent"||d==="persist")};c.field===-1||c.permanent===-1||(o.values||[]).forEach(d=>{let p=String(d[c.field]||"").trim(),f=t.get(p);if(f==null)return;let u=Ca(d[c.permanent]);e[f]=u})})}catch(n){console.warn("PTO: unable to load permanent flags",n)}return e}function se(){var i;if(!Ae)return;let e=N.focusedIndex<=0?"disabled":"",t=N.focusedIndex>=ae.length-1?"disabled":"",n=N.activeView==="step"&&N.activeStepId!=null,o=N.activeView==="config"?cn():n?na(N.activeStepId):`${Qn()}${Xn()}`;Ae.innerHTML=`
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
    `;let s=N.activeView==="home"||N.activeView!=="step"&&N.activeView!=="config",r=document.getElementById("pf-info-fab-pto");if(s)r&&r.remove();else if((i=window.PrairieForge)!=null&&i.mountInfoFab){let c=Kn(N.activeStepId);PrairieForge.mountInfoFab({title:c.title,content:c.content,buttonId:"pf-info-fab-pto"})}aa(),ra(),s?At():tt()}function Kn(e){switch(e){case 0:return{title:"Configuration",content:`
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
    `}function Zn(e,t){let n=N.stepStatuses[e.id]||"pending",a=N.activeView==="step"&&N.focusedIndex===t?"pf-step-card--active":"",o=Ht(ka(e.id));return`
        <article class="pf-step-card pf-clickable ${a}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${e.title}</h3>
        </article>
    `}function ea(e){let t=ae.filter(o=>o.id!==6).map(o=>({id:o.id,title:o.title,complete:la(o.id)})),n=t.every(o=>o.complete),a=t.map(o=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${o.complete?"is-active":""}" aria-pressed="${o.complete}">
                        ${Te}
                    </span>
                    <div>
                        <h3>${b(o.title)}</h3>
                        <p class="pf-config-subtext">${o.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(he)} | Step ${e.id}</p>
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
    `}function cn(){if(!S.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=tn(te(k.payrollDate)),t=tn(te(k.accountingPeriod)),n=te(k.journalEntryId),a=te(k.accountingSoftware),o=te(k.payrollProvider),s=te(k.companyName),r=te(k.reviewerName),i=we(0),c=!!S.permanents[0],d=!!(bn(S.completes[0])||i.signOffDate),p=be(i==null?void 0:i.reviewer),f=(i==null?void 0:i.signOffDate)||"";return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${b(he)} | Step 0</p>
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
                        <input type="text" id="config-user-name" value="${b(r)}" placeholder="Full name">
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
            ${fe({textareaId:"config-notes",value:i.notes||"",permanentId:"config-notes-lock",isPermanent:c,hintId:"",saveButtonId:"config-notes-save"})}
            ${ge({reviewerInputId:"config-reviewer",reviewerValue:p,signoffInputId:"config-signoff-date",signoffValue:f,isComplete:d,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"})}
        </section>
    `}function ta(e){let t=we(1),n=!!S.permanents[1],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(xe(S.completes[1])||o),r=te(k.payrollProvider);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(he)} | Step ${e.id}</p>
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
                    ${B(r?`<a href="${b(r)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${it}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${it}</button>`,"Provider")}
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${Ve}</button>`,"PTO_Data")}
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${zt}</button>`,"Clear")}
                </div>
            </article>
            ${fe({textareaId:"step-notes-1",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-1",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-1"})}
            ${ge({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
        </section>
    `}function na(e){let t=ae.find(i=>i.id===e);if(!t)return"";if(e===0)return cn();if(e===1)return ta(t);if(e===2)return Ia(t);if(e===3)return Aa(t);if(e===4)return Na(t);if(e===5)return Da(t);if(t.id===6)return ea(t);let n=we(e),a=!!S.permanents[e],o=be(n==null?void 0:n.reviewer),s=(n==null?void 0:n.signOffDate)||"",r=!!(xe(S.completes[e])||s);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(he)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${b(t.title)}</h2>
            <p class="pf-hero-copy">${b(t.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${fe({textareaId:`step-notes-${e}`,value:(n==null?void 0:n.notes)||"",permanentId:`step-notes-lock-${e}`,isPermanent:a,hintId:"",saveButtonId:`step-notes-save-${e}`})}
            ${ge({reviewerInputId:`step-reviewer-${e}`,reviewerValue:o,signoffInputId:`step-signoff-${e}`,signoffValue:s,isComplete:r,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`})}
        </section>
    `}function aa(){var n,a,o,s,r;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",async()=>{var c;let i=Be(Ne);await Le(i.sheetName,i.title,i.subtitle),$e({activeView:"home",activeStepId:null}),(c=document.getElementById("pf-hero"))==null||c.scrollIntoView({behavior:"smooth",block:"start"})}),(a=document.getElementById("nav-selector"))==null||a.addEventListener("click",()=>{window.location.href=Mn}),(o=document.getElementById("nav-prev"))==null||o.addEventListener("click",()=>ze(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>ze(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",i=>{i.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",i=>{!(t!=null&&t.contains(i.target))&&!(e!=null&&e.contains(i.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(r=document.getElementById("nav-config"))==null||r.addEventListener("click",async()=>{t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"),await ba()}),document.querySelectorAll("[data-step-card]").forEach(i=>{let c=Number(i.getAttribute("data-step-index")),d=Number(i.getAttribute("data-step-id"));i.addEventListener("click",()=>De(c,d))}),N.activeView==="config"?sa():N.activeView==="step"&&N.activeStepId!=null&&oa(N.activeStepId)}function oa(e){var u,l,g,y,w,h,E,C,I,_,$,j,L,Q,F,z,T,m;let t=e===2?document.getElementById("step-notes-input"):document.getElementById(`step-notes-${e}`),n=e===2?document.getElementById("step-reviewer-name"):document.getElementById(`step-reviewer-${e}`),a=e===2?document.getElementById("step-signoff-date"):document.getElementById(`step-signoff-${e}`),o=document.getElementById("step-back-btn"),s=e===2?document.getElementById("step-notes-lock-2"):document.getElementById(`step-notes-lock-${e}`),r=e===2?document.getElementById("step-notes-save-2"):document.getElementById(`step-notes-save-${e}`);r==null||r.addEventListener("click",async()=>{let v=(t==null?void 0:t.value)||"";await ne(e,"notes",v),ce(r,!0)});let i=e===2?document.getElementById("headcount-signoff-save"):document.getElementById(`step-signoff-save-${e}`);i==null||i.addEventListener("click",async()=>{let v=(n==null?void 0:n.value)||"";await ne(e,"reviewer",v),ce(i,!0)}),lt();let c=e===2?"headcount-signoff-toggle":`step-signoff-toggle-${e}`,d=`${c}-prev`,p=`${c}-next`,f=e===2?"step-signoff-date":`step-signoff-${e}`;vn(e,{buttonId:c,inputId:f,canActivate:e===2?()=>{var P;return!wn()||((P=document.getElementById("step-notes-input"))==null?void 0:P.value.trim())||""?!0:(K("Please enter a brief explanation of the headcount differences before completing this step.","info"),!1)}:null,onComplete:ia(e)}),pn(d,p),o==null||o.addEventListener("click",async()=>{let v=Be(Ne);await Le(v.sheetName,v.title,v.subtitle),$e({activeView:"home",activeStepId:null})}),s==null||s.addEventListener("click",async()=>{let v=!s.classList.contains("is-locked");rt(s,v),await hn(e,v)}),e===6&&((u=document.getElementById("archive-run-btn"))==null||u.addEventListener("click",()=>{})),e===1&&((l=document.getElementById("import-open-data-btn"))==null||l.addEventListener("click",()=>gn("PTO_Data")),(g=document.getElementById("import-clear-btn"))==null||g.addEventListener("click",()=>ha())),e===2&&((y=document.getElementById("headcount-skip-btn"))==null||y.addEventListener("click",()=>{R.skipAnalysis=!R.skipAnalysis;let v=document.getElementById("headcount-skip-btn");v==null||v.classList.toggle("is-active",R.skipAnalysis),R.skipAnalysis&&rn(),sn()}),(w=document.getElementById("headcount-run-btn"))==null||w.addEventListener("click",()=>ft()),(h=document.getElementById("headcount-refresh-btn"))==null||h.addEventListener("click",()=>ft()),Ma(),R.skipAnalysis&&rn(),sn()),e===3&&((E=document.getElementById("quality-run-btn"))==null||E.addEventListener("click",()=>Qt()),(C=document.getElementById("quality-refresh-btn"))==null||C.addEventListener("click",()=>Qt()),(I=document.getElementById("quality-acknowledge-btn"))==null||I.addEventListener("click",()=>da())),e===4&&((_=document.getElementById("analysis-refresh-btn"))==null||_.addEventListener("click",()=>Xt()),($=document.getElementById("analysis-run-btn"))==null||$.addEventListener("click",()=>Xt()),(j=document.getElementById("payrate-save-btn"))==null||j.addEventListener("click",Kt),(L=document.getElementById("payrate-ignore-btn"))==null||L.addEventListener("click",ca),(Q=document.getElementById("payrate-input"))==null||Q.addEventListener("keydown",v=>{v.key==="Enter"&&Kt()})),e===5&&((F=document.getElementById("je-create-btn"))==null||F.addEventListener("click",()=>fa()),(z=document.getElementById("je-run-btn"))==null||z.addEventListener("click",()=>fn()),(T=document.getElementById("je-export-btn"))==null||T.addEventListener("click",()=>ga()),(m=document.getElementById("je-upload-btn"))==null||m.addEventListener("click",()=>ma()))}function sa(){var i,c,d,p,f;$t("config-payroll-date",{onChange:u=>{if(le(k.payrollDate,u),!!u){if(!S.overrides.accountingPeriod){let l=xa(u);if(l){let g=document.getElementById("config-accounting-period");g&&(g.value=l),le(k.accountingPeriod,l)}}if(!S.overrides.journalId){let l=Ea(u);if(l){let g=document.getElementById("config-journal-id");g&&(g.value=l),le(k.journalEntryId,l)}}}}});let e=document.getElementById("config-accounting-period");e==null||e.addEventListener("change",u=>{S.overrides.accountingPeriod=!!u.target.value,le(k.accountingPeriod,u.target.value||"")});let t=document.getElementById("config-journal-id");t==null||t.addEventListener("change",u=>{S.overrides.journalId=!!u.target.value,le(k.journalEntryId,u.target.value.trim())}),(i=document.getElementById("config-company-name"))==null||i.addEventListener("change",u=>{le(k.companyName,u.target.value.trim())}),(c=document.getElementById("config-payroll-provider"))==null||c.addEventListener("change",u=>{le(k.payrollProvider,u.target.value.trim())}),(d=document.getElementById("config-accounting-link"))==null||d.addEventListener("change",u=>{le(k.accountingSoftware,u.target.value.trim())}),(p=document.getElementById("config-user-name"))==null||p.addEventListener("change",u=>{le(k.reviewerName,u.target.value.trim())});let n=document.getElementById("config-notes");n==null||n.addEventListener("input",u=>{ne(0,"notes",u.target.value)});let a=document.getElementById("config-notes-lock");a==null||a.addEventListener("click",async()=>{let u=!a.classList.contains("is-locked");rt(a,u),await hn(0,u)});let o=document.getElementById("config-notes-save");o==null||o.addEventListener("click",async()=>{n&&(await ne(0,"notes",n.value),ce(o,!0))});let s=document.getElementById("config-reviewer");s==null||s.addEventListener("change",u=>{let l=u.target.value.trim();ne(0,"reviewer",l);let g=document.getElementById("config-signoff-date");if(l&&g&&!g.value){let y=mt();g.value=y,ne(0,"signOffDate",y),yn(0,!0)}}),(f=document.getElementById("config-signoff-date"))==null||f.addEventListener("change",u=>{ne(0,"signOffDate",u.target.value||"")});let r=document.getElementById("config-signoff-save");r==null||r.addEventListener("click",async()=>{var g,y;let u=((g=s==null?void 0:s.value)==null?void 0:g.trim())||"",l=((y=document.getElementById("config-signoff-date"))==null?void 0:y.value)||"";await ne(0,"reviewer",u),await ne(0,"signOffDate",l),ce(r,!0)}),lt(),vn(0,{buttonId:"config-signoff-toggle",inputId:"config-signoff-date",onComplete:()=>{Ta(),dn(0),un()}}),pn("config-signoff-toggle-prev","config-signoff-toggle-next")}function De(e,t=null){if(e<0||e>=ae.length)return;Je=e;let n=t!=null?t:ae[e].id;$e({focusedIndex:e,activeView:n===0?"config":"step",activeStepId:n});let o=Hn[n];o&&gn(o),n===2&&!R.hasAnalyzed&&(kn(),ft())}function ia(e){return e===6?null:()=>dn(e)}function dn(e){let t=ae.findIndex(a=>a.id===e);if(t===-1)return;let n=t+1;n<ae.length&&(De(n,ae[n].id),un())}function un(){let e=[document.querySelector(".pf-root"),document.querySelector(".pf-step-guide"),document.body];for(let t of e)t&&t.scrollTo({top:0,behavior:"smooth"});window.scrollTo({top:0,behavior:"smooth"})}function ze(e){let t=N.focusedIndex+e,n=Math.max(0,Math.min(ae.length-1,t));De(n,ae[n].id),window.scrollTo({top:0,behavior:"smooth"})}function pn(e,t){var n,a;(n=document.getElementById(e))==null||n.addEventListener("click",()=>ze(-1)),(a=document.getElementById(t))==null||a.addEventListener("click",()=>ze(1))}function ra(){if(Je===null)return;let e=document.querySelector(`[data-step-index="${Je}"]`);Je=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}function la(e){return bn(S.completes[e])}function $e(e){e.stepStatuses&&(N.stepStatuses={...N.stepStatuses,...e.stepStatuses}),Object.assign(N,{...e,stepStatuses:N.stepStatuses}),se()}function ie(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}async function Kt(){let e=document.getElementById("payrate-input");if(!e)return;let t=parseFloat(e.value),n=e.dataset.employee,a=parseInt(e.dataset.row,10);if(isNaN(t)||t<=0){K("Please enter a valid pay rate greater than 0.","info");return}if(!n||isNaN(a)){console.error("Missing employee data on input");return}ee(!0,"Updating pay rate...");try{await Excel.run(async o=>{let s=o.workbook.worksheets.getItem("PTO_Analysis"),r=s.getCell(a-1,3);r.values=[[t]];let i=s.getCell(a-1,8);i.load("values"),await o.sync();let d=(Number(i.values[0][0])||0)*t,p=s.getCell(a-1,9);p.values=[[d]];let f=s.getCell(a-1,10);f.load("values"),await o.sync();let u=Number(f.values[0][0])||0,l=d-u,g=s.getCell(a-1,11);g.values=[[l]],await o.sync()}),Z.missingPayRates=Z.missingPayRates.filter(o=>o.name!==n),ee(!1),De(3,3)}catch(o){console.error("Failed to save pay rate:",o),K(`Failed to save pay rate: ${o.message}`,"error"),ee(!1)}}function ca(){let e=document.getElementById("payrate-input");if(!e)return;let t=e.dataset.employee;t&&(Z.ignoredMissingPayRates.add(t),Z.missingPayRates=Z.missingPayRates.filter(n=>n.name!==t)),De(3,3)}async function Qt(){if(!ie()){K("Excel is not available. Open this module inside Excel to run quality check.","info");return}W.loading=!0,ee(!0,"Analyzing data quality..."),ce(document.getElementById("quality-save-btn"),!1);try{await Excel.run(async t=>{var w;let a=t.workbook.worksheets.getItem("PTO_Data").getUsedRangeOrNullObject();a.load("values"),await t.sync();let o=a.isNullObject?[]:a.values||[];if(!o.length||o.length<2)throw new Error("PTO_Data is empty or has no data rows.");let s=(o[0]||[]).map(h=>J(h));console.log("[Data Quality] PTO_Data headers:",o[0]);let r=s.findIndex(h=>h==="employee name"||h==="employeename");r===-1&&(r=s.findIndex(h=>h.includes("employee")&&h.includes("name"))),r===-1&&(r=s.findIndex(h=>h==="name"||h.includes("name")&&!h.includes("company")&&!h.includes("form"))),console.log("[Data Quality] Employee name column index:",r,"Header:",(w=o[0])==null?void 0:w[r]);let i=D(s,["balance"]),c=D(s,["accrual rate","accrualrate"]),d=D(s,["carry over","carryover"]),p=D(s,["ytd accrued","ytdaccrued"]),f=D(s,["ytd used","ytdused"]),u=[],l=[],g=[],y=o.slice(1);y.forEach((h,E)=>{let C=E+2,I=r!==-1?String(h[r]||"").trim():`Row ${C}`;if(!I)return;let _=i!==-1&&Number(h[i])||0,$=c!==-1&&Number(h[c])||0,j=d!==-1&&Number(h[d])||0,L=p!==-1&&Number(h[p])||0,Q=f!==-1&&Number(h[f])||0,F=j+L;_<0?u.push({name:I,issue:`Negative balance: ${_.toFixed(2)} hrs`,rowIndex:C}):Q>F&&F>0&&u.push({name:I,issue:`Used ${Q.toFixed(0)} hrs but only ${F.toFixed(0)} available`,rowIndex:C}),_===0&&(j>0||L>0)&&l.push({name:I,rowIndex:C}),$>8&&g.push({name:I,accrualRate:$,rowIndex:C})}),W.balanceIssues=u,W.zeroBalances=l,W.accrualOutliers=g,W.totalIssues=u.length,W.totalEmployees=y.filter(h=>h.some(E=>E!==null&&E!=="")).length,W.hasRun=!0});let e=W.balanceIssues.length>0;$e({stepStatuses:{3:e?"blocked":"complete"}})}catch(e){console.error("Data quality check error:",e),K(`Quality check failed: ${e.message}`,"error"),W.hasRun=!1}finally{W.loading=!1,ee(!1),se()}}function da(){W.acknowledged=!0,$e({stepStatuses:{3:"complete"}}),se()}async function ua(){if(ie())try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem("PTO_Data"),n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=t.getUsedRangeOrNullObject();if(a.load("values"),n.load("isNullObject"),await e.sync(),n.isNullObject){Z.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let o=n.getUsedRangeOrNullObject();o.load("values"),await e.sync();let s=a.isNullObject?[]:a.values||[],r=o.isNullObject?[]:o.values||[];if(!s.length||!r.length){Z.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let i=(p,f,u)=>{let l=(p[0]||[]).map(w=>J(w)),g=D(l,f);return g===-1?null:p.slice(1).reduce((w,h)=>w+(Number(h[g])||0),0)},c=[{key:"accrualRate",aliases:["accrual rate","accrualrate"]},{key:"carryOver",aliases:["carry over","carryover","carry_over"]},{key:"ytdAccrued",aliases:["ytd accrued","ytdaccrued","ytd_accrued"]},{key:"ytdUsed",aliases:["ytd used","ytdused","ytd_used"]},{key:"balance",aliases:["balance"]}],d={};for(let p of c){let f=i(s,p.aliases,"PTO_Data"),u=i(r,p.aliases,"PTO_Analysis");if(f===null||u===null)d[p.key]=null;else{let l=Math.abs(f-u)<.01;d[p.key]={match:l,ptoData:f,ptoAnalysis:u}}}Z.completenessCheck=d})}catch(e){console.error("Completeness check failed:",e)}}async function Xt(){if(!ie()){K("Excel is not available. Open this module inside Excel to run analysis.","info");return}ee(!0,"Running analysis...");try{await kn(),await ua(),Z.cleanDataReady=!0,se()}catch(e){console.error("Full analysis error:",e),K(`Analysis failed: ${e.message}`,"error")}finally{ee(!1)}}async function fn(){if(!ie()){K("Excel is not available. Open this module inside Excel to run journal checks.","info");return}G.loading=!0,G.lastError=null,ce(document.getElementById("je-save-btn"),!1),se();try{let e=await Excel.run(async t=>{let a=t.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();a.load("values");let o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis");o.load("isNullObject"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty. Generate the JE first.");let r=(s[0]||[]).map(C=>J(C)),i=D(r,["debit"]),c=D(r,["credit"]),d=D(r,["lineamount","line amount"]),p=D(r,["account number","accountnumber"]);if(i===-1||c===-1)throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");let f=0,u=0,l=0,g=0;s.slice(1).forEach(C=>{let I=Number(C[i])||0,_=Number(C[c])||0,$=d!==-1&&Number(C[d])||0,j=p!==-1?String(C[p]||"").trim():"";f+=I,u+=_,l+=$,j&&j!=="21540"&&(g+=$)});let y=0;if(!o.isNullObject){let C=o.getUsedRangeOrNullObject();C.load("values"),await t.sync();let I=C.isNullObject?[]:C.values||[];if(I.length>1){let _=(I[0]||[]).map(j=>J(j)),$=D(_,["change"]);$!==-1&&I.slice(1).forEach(j=>{y+=Number(j[$])||0})}}let w=f-u,h=[];Math.abs(w)>=.01?h.push({check:"Debits = Credits",passed:!1,detail:w>0?`Debits exceed credits by $${Math.abs(w).toLocaleString(void 0,{minimumFractionDigits:2})}`:`Credits exceed debits by $${Math.abs(w).toLocaleString(void 0,{minimumFractionDigits:2})}`}):h.push({check:"Debits = Credits",passed:!0,detail:""}),Math.abs(l)>=.01?h.push({check:"Line Amounts Sum to Zero",passed:!1,detail:`Line amounts sum to $${l.toLocaleString(void 0,{minimumFractionDigits:2})} (should be $0.00)`}):h.push({check:"Line Amounts Sum to Zero",passed:!0,detail:""});let E=Math.abs(g-y);return E>=.01?h.push({check:"JE Matches Analysis Total",passed:!1,detail:`JE expense total ($${g.toLocaleString(void 0,{minimumFractionDigits:2})}) differs from PTO_Analysis Change total ($${y.toLocaleString(void 0,{minimumFractionDigits:2})}) by $${E.toLocaleString(void 0,{minimumFractionDigits:2})}`}):h.push({check:"JE Matches Analysis Total",passed:!0,detail:""}),{debitTotal:f,creditTotal:u,difference:w,lineAmountSum:l,jeChangeTotal:g,analysisChangeTotal:y,issues:h,validationRun:!0}});Object.assign(G,e,{lastError:null})}catch(e){console.warn("PTO JE summary:",e),G.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",G.debitTotal=null,G.creditTotal=null,G.difference=null,G.lineAmountSum=null,G.jeChangeTotal=null,G.analysisChangeTotal=null,G.issues=[],G.validationRun=!1}finally{G.loading=!1,se()}}var pa={"general & administrative":"64110","general and administrative":"64110","g&a":"64110","research & development":"62110","research and development":"62110","r&d":"62110",marketing:"61610","cogs onboarding":"53110","cogs prof. services":"56110","cogs professional services":"56110","sales & marketing":"61110","sales and marketing":"61110","cogs support":"52110","client success":"61811"},Zt="21540";async function fa(){if(!ie()){K("Excel is not available. Open this module inside Excel to create the journal entry.","info");return}ee(!0,"Creating PTO Journal Entry...");try{await Excel.run(async e=>{let t=[],n=e.workbook.tables.getItemOrNullObject(de[0]);if(n.load("isNullObject"),await e.sync(),n.isNullObject){let m=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");if(m.load("isNullObject"),await e.sync(),!m.isNullObject){let v=m.getUsedRangeOrNullObject();v.load("values"),await e.sync();let P=v.isNullObject?[]:v.values||[];t=P.length>1?P.slice(1):[]}}else{let m=n.getDataBodyRange();m.load("values"),await e.sync(),t=m.values||[]}let a=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis");if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PTO_Analysis sheet not found. Please ensure the worksheet exists.");let o=a.getUsedRangeOrNullObject();o.load("values");let s=e.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");s.load("isNullObject"),await e.sync();let r=[];if(!s.isNullObject){let m=s.getUsedRangeOrNullObject();m.load("values"),await e.sync(),r=m.isNullObject?[]:m.values||[]}let i=o.isNullObject?[]:o.values||[];if(!i.length||i.length<2)throw new Error("PTO_Analysis is empty or has no data rows. Run the analysis first (Step 4).");let c={};t.forEach(m=>{let v=String(m[1]||"").trim(),P=m[2];v&&(c[v]=P)}),(!c[k.journalEntryId]||!c[k.payrollDate])&&console.warn("[JE Draft] Missing config values - RefNumber:",c[k.journalEntryId],"TxnDate:",c[k.payrollDate]);let d=c[k.journalEntryId]||"",p=c[k.payrollDate]||"",f=c[k.accountingPeriod]||"",u="";if(p)try{let m;if(typeof p=="number"||/^\d{4,5}$/.test(String(p).trim())){let v=Number(p),P=new Date(1899,11,30);m=new Date(P.getTime()+v*24*60*60*1e3)}else m=new Date(p);if(!isNaN(m.getTime())&&m.getFullYear()>1970){let v=String(m.getMonth()+1).padStart(2,"0"),P=String(m.getDate()).padStart(2,"0"),A=m.getFullYear();u=`${v}/${P}/${A}`}else console.warn("[JE Draft] Date parsing resulted in invalid date:",p,"->",m),u=String(p)}catch(m){console.warn("[JE Draft] Could not parse TxnDate:",p,m),u=String(p)}let l=f?`${f} PTO Accrual`:"PTO Accrual",g={};if(r.length>1){let m=(r[0]||[]).map(A=>J(A)),v=D(m,["account number","accountnumber","account","acct"]),P=D(m,["account name","accountname","name","description"]);v!==-1&&P!==-1&&r.slice(1).forEach(A=>{let q=String(A[v]||"").trim(),ue=String(A[P]||"").trim();q&&(g[q]=ue)})}let y=(i[0]||[]).map(m=>J(m));console.log("[JE Draft] PTO_Analysis headers:",y),console.log("[JE Draft] PTO_Analysis row count:",i.length-1);let w=D(y,["department"]),h=D(y,["change"]);if(console.log("[JE Draft] Column indices - Department:",w,"Change:",h),w===-1||h===-1)throw new Error(`Could not find required columns in PTO_Analysis. Found headers: ${y.join(", ")}. Looking for "Department" (found: ${w!==-1}) and "Change" (found: ${h!==-1}).`);let E={},C=0,I=0,_=0;if(i.slice(1).forEach((m,v)=>{C++;let P=String(m[w]||"").trim(),A=m[h],q=Number(A)||0;if(v<3&&console.log(`[JE Draft] Row ${v+2}: Dept="${P}", Change raw="${A}", Change num=${q}`),!P){_++;return}if(q===0){I++;return}E[P]||(E[P]=0),E[P]+=q}),console.log(`[JE Draft] Data summary: ${C} rows, ${I} with zero change, ${_} missing dept`),console.log("[JE Draft] Department totals:",E),Object.keys(E).length===0){let m=`No journal entry lines could be created.

`;throw I===C?(m+=`All 'Change' amounts in PTO_Analysis are $0.00.

`,m+=`Common causes:
`,m+=`\u2022 Missing Pay Rate data (Liability = Balance \xD7 Pay Rate)
`,m+=`\u2022 No prior period data to compare against
`,m+=`\u2022 PTO Analysis hasn't been run yet

`,m+="Please verify Pay Rate values exist in PTO_Analysis."):_===C?(m+=`All rows are missing Department values.

`,m+="Please ensure the 'Department' column is populated in PTO_Analysis."):(m+=`Found ${C} rows but none had both a Department and non-zero Change amount.
`,m+=`\u2022 ${I} rows with zero change
`,m+=`\u2022 ${_} rows missing department`),new Error(m)}let j=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],L=[j],Q=0,F=0;Object.entries(E).forEach(([m,v])=>{if(Math.abs(v)<.01)return;let P=m.toLowerCase().trim(),A=pa[P]||"",q=g[A]||"",ue=v>0?Math.abs(v):0,O=v<0?Math.abs(v):0;Q+=ue,F+=O,L.push([d,u,A,q,v,ue,O,l,m])});let z=Q-F;if(Math.abs(z)>=.01){let m=z<0?Math.abs(z):0,v=z>0?Math.abs(z):0,P=g[Zt]||"Accrued PTO";L.push([d,u,Zt,P,-z,m,v,l,""])}let T=e.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");if(T.load("isNullObject"),await e.sync(),T.isNullObject)T=e.workbook.worksheets.add("PTO_JE_Draft");else{let m=T.getUsedRangeOrNullObject();m.load("isNullObject"),await e.sync(),m.isNullObject||m.clear()}if(L.length>0){let m=T.getRangeByIndexes(0,0,L.length,j.length);m.values=L;let v=T.getRangeByIndexes(0,0,1,j.length);dt(v);let P=L.length-1;P>0&&(me(T,4,P,!0),me(T,5,P),me(T,6,P)),m.format.autofitColumns()}await e.sync(),T.activate(),T.getRange("A1").select(),await e.sync()}),await fn()}catch(e){console.error("Create JE Draft error:",e),K(`Unable to create Journal Entry: ${e.message}`,"error")}finally{ee(!1)}}async function ga(){if(!ie()){K("Excel is not available. Open this module inside Excel to export.","info");return}ee(!0,"Preparing JE CSV...");try{let{rows:e}=await Excel.run(async n=>{let o=n.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();o.load("values"),await n.sync();let s=o.isNullObject?[]:o.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty.");return{rows:s}}),t=La(e);Ba(`pto-je-draft-${mt()}.csv`,t)}catch(e){console.error("PTO JE export:",e),K("Unable to export the JE draft. Confirm the sheet has data.","error")}finally{ee(!1)}}async function ma(){let e=te(k.accountingSoftware)||te("SS_Accounting_Software");if(!e&&Y())try{let t=await Qe(de);e=t.SS_Accounting_Software||t.Accounting_Software||t[k.accountingSoftware]}catch(t){console.warn("Error reading accounting software URL:",t)}if(!e){K("No accounting software URL configured. Add SS_Accounting_Software to SS_PF_Config.","info",5e3);return}!e.startsWith("http://")&&!e.startsWith("https://")&&(e="https://"+e),window.open(e,"_blank"),K("Opening accounting software...","success",2e3)}async function gn(e){if(!(!e||!ie()))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e);n.activate(),n.getRange("A1").select(),await t.sync()})}catch(t){console.error(t)}}async function ha(){if(!(!ie()||!await Ln(`All data in PTO_Data will be permanently removed.

This action cannot be undone.`,{title:"Clear PTO Data",icon:"\u{1F5D1}\uFE0F",confirmText:"Clear Data",cancelText:"Keep Data",destructive:!0}))){ee(!0);try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("PTO_Data"),a=n.getUsedRangeOrNullObject();a.load("rowCount"),await t.sync(),!a.isNullObject&&a.rowCount>1&&(n.getRangeByIndexes(1,0,a.rowCount-1,20).clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),K("PTO_Data cleared successfully. You can now paste new data.","success")}catch(t){console.error("Clear PTO_Data error:",t),K(`Failed to clear PTO_Data: ${t.message}`,"error")}finally{ee(!1)}}}async function ya(){if(!ie())return[];try{return await Excel.run(async e=>{let t=e.workbook.worksheets;return t.load("items/name,visibility"),await e.sync(),t.items.filter(a=>{let s=(a.name||"").toUpperCase();return s.startsWith("SS_")||s.includes("MAPPING")||s.includes("HOMEPAGE")}).map(a=>({name:a.name,visible:a.visibility===Excel.SheetVisibility.visible,isHomepage:(a.name||"").toUpperCase().includes("HOMEPAGE")})).sort((a,o)=>a.isHomepage&&!o.isHomepage?1:!a.isHomepage&&o.isHomepage?-1:a.name.localeCompare(o.name))})}catch(e){return console.error("[Config] Error reading configuration sheets:",e),[]}}function va(){if(document.getElementById("config-sheet-modal"))return;let e=document.createElement("div");if(e.id="config-sheet-modal",e.className="pf-config-modal hidden",e.innerHTML=`
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
        `,document.head.appendChild(t)}}async function ba(){va();let e=document.getElementById("config-sheet-modal"),t=document.getElementById("config-sheet-list");if(!e||!t)return;t.textContent="Loading\u2026",e.classList.remove("hidden");let n=await ya();n.length?(t.innerHTML="",n.forEach(a=>{let o=document.createElement("button");o.type="button",o.className="pf-config-sheet",o.innerHTML=`<span>${a.name}</span><span class="pf-config-pill">${a.visible?"Visible":"Hidden"}</span>`,o.addEventListener("click",async()=>{await wa(a.name),e.classList.add("hidden")}),t.appendChild(o)})):t.textContent="No configuration sheets found.",e.querySelectorAll("[data-close]").forEach(a=>a.addEventListener("click",()=>e.classList.add("hidden")))}async function wa(e){if(!(!e||!ie()))try{await Excel.run(async t=>{let n=t.workbook.worksheets,a=n.getItemOrNullObject(e);a.load("isNullObject,visibility"),await t.sync(),a.isNullObject&&(a=n.add(e)),a.visibility=Excel.SheetVisibility.visible,await t.sync(),a.activate(),a.getRange("A1").select(),await t.sync(),console.log(`[Config] Opened sheet: ${e}`)})}catch(t){console.error("[Config] Error opening sheet",e,t)}}function te(e){var n,a;let t=String(e!=null?e:"").trim();return(a=(n=S.values)==null?void 0:n[t])!=null?a:""}function be(e){var n;if(e)return e;let t=te(k.reviewerName);if(t)return t;if((n=window.PrairieForge)!=null&&n._sharedConfigCache){let a=window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer")||window.PrairieForge._sharedConfigCache.get("Default_Reviewer");if(a)return a}return""}function le(e,t,n={}){var r;let a=String(e!=null?e:"").trim();if(!a)return;S.values[a]=t!=null?t:"";let o=(r=n.debounceMs)!=null?r:0;if(!o){let i=Se.get(a);i&&clearTimeout(i),Se.delete(a),_e(a,t!=null?t:"",de);return}Se.has(a)&&clearTimeout(Se.get(a));let s=setTimeout(()=>{Se.delete(a),_e(a,t!=null?t:"",de)},o);Se.set(a,s)}function J(e){return String(e!=null?e:"").trim().toLowerCase()}function ee(e,t="Working..."){let n=document.getElementById(Vn);n&&(n.style.display="none")}function pt(){Gn()}typeof Office!="undefined"&&Office.onReady?Office.onReady(()=>pt()).catch(()=>pt()):pt();function we(e){return S.steps[e]||{notes:"",reviewer:"",signOffDate:""}}function mn(e){return qe[e]||{}}function ka(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}async function ne(e,t,n){let a=S.steps[e]||{notes:"",reviewer:"",signOffDate:""};a[t]=n,S.steps[e]=a;let o=mn(e),s=t==="notes"?o.note:t==="reviewer"?o.reviewer:o.signOff;if(s&&Y())try{await _e(s,n,de)}catch(r){console.warn("PTO: unable to save field",s,r)}}async function hn(e,t){S.permanents[e]=t;let n=mn(e);if(n!=null&&n.note&&Y())try{await Excel.run(async a=>{var u;let o=a.workbook.tables.getItemOrNullObject(de[0]);if(await a.sync(),o.isNullObject)return;let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let i=r.values[0]||[],c=i.map(l=>String(l||"").trim().toLowerCase()),d={field:c.findIndex(l=>l==="field"||l==="field name"||l==="setting"),permanent:c.findIndex(l=>l==="permanent"||l==="persist"),value:c.findIndex(l=>l==="value"||l==="setting value"),type:c.findIndex(l=>l==="type"||l==="category"),title:c.findIndex(l=>l==="title"||l==="display name")};if(d.field===-1)return;let f=(s.values||[]).findIndex(l=>String(l[d.field]||"").trim()===n.note);if(f>=0)d.permanent>=0&&(s.getCell(f,d.permanent).values=[[t?"Y":"N"]]);else{let l=new Array(i.length).fill("");d.type>=0&&(l[d.type]="Other"),d.title>=0&&(l[d.title]=""),l[d.field]=n.note,d.permanent>=0&&(l[d.permanent]=t?"Y":"N"),d.value>=0&&(l[d.value]=((u=S.steps[e])==null?void 0:u.notes)||""),o.rows.add(null,[l])}await a.sync()})}catch(a){console.warn("PTO: unable to update permanent flag",a)}}async function yn(e,t){let n=ln[e];if(n&&(S.completes[e]=t?"Y":"",!!Y()))try{await _e(n,t?"Y":"",de)}catch(a){console.warn("PTO: unable to save completion flag",n,a)}}function en(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function Oa(){let e={};return Object.keys(qe).forEach(t=>{var s;let n=parseInt(t,10),a=!!((s=S.steps[n])!=null&&s.signOffDate),o=!!S.completes[n];e[n]=a||o}),e}function vn(e,{buttonId:t,inputId:n,canActivate:a=null,onComplete:o=null}){var c;let s=document.getElementById(t);if(!s)return;let r=document.getElementById(n),i=!!((c=S.steps[e])!=null&&c.signOffDate)||!!S.completes[e];en(s,i),s.addEventListener("click",()=>{if(!s.classList.contains("is-active")&&e>0){let f=Oa(),{canComplete:u,message:l}=qt(e,f);if(!u){Yt(l);return}}if(typeof a=="function"&&!a())return;let p=!s.classList.contains("is-active");en(s,p),r&&(r.value=p?mt():"",ne(e,"signOffDate",r.value)),yn(e,p),p&&typeof o=="function"&&o()})}function b(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function Sa(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function bn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function xe(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function gt(e){if(!e)return null;let t=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(e));if(!t)return null;let n=Number(t[1]),a=Number(t[2]),o=Number(t[3]);return!n||!a||!o?null:{year:n,month:a,day:o}}function tn(e){if(!e)return"";let t=gt(e);if(!t)return"";let{year:n,month:a,day:o}=t;return`${n}-${String(a).padStart(2,"0")}-${String(o).padStart(2,"0")}`}function xa(e){let t=gt(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function Ea(e){let t=gt(e);return t?`PTO-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function mt(){let e=new Date,t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}function Ca(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="y"||t==="yes"||t==="true"||t==="t"||t==="1"}function _a(e){if(e instanceof Date)return e.getTime();if(typeof e=="number"){let n=Pa(e);return n?n.getTime():null}let t=new Date(e);return Number.isNaN(t.getTime())?null:t.getTime()}function Pa(e){if(!Number.isFinite(e))return null;let t=new Date(Date.UTC(1899,11,30));return new Date(t.getTime()+e*24*60*60*1e3)}function Ta(){let e=n=>{var a,o;return((o=(a=document.getElementById(n))==null?void 0:a.value)==null?void 0:o.trim())||""};[{id:"config-payroll-date",field:k.payrollDate},{id:"config-accounting-period",field:k.accountingPeriod},{id:"config-journal-id",field:k.journalEntryId},{id:"config-company-name",field:k.companyName},{id:"config-payroll-provider",field:k.payrollProvider},{id:"config-accounting-link",field:k.accountingSoftware},{id:"config-user-name",field:k.reviewerName}].forEach(({id:n,field:a})=>{let o=e(n);a&&le(a,o)})}function D(e,t=[]){let n=t.map(a=>J(a));return e.findIndex(a=>n.some(o=>a.includes(o)))}function Ia(e){var C,I,_,$,j,L,Q,F,z;let t=we(2),n=(t==null?void 0:t.notes)||"",a=!!S.permanents[2],o=be(t==null?void 0:t.reviewer),s=(t==null?void 0:t.signOffDate)||"",r=!!(xe(S.completes[2])||s),i=R.roster||{},c=R.hasAnalyzed,d=(I=(C=R.roster)==null?void 0:C.difference)!=null?I:0,p=!R.skipAnalysis&&Math.abs(d)>0,f=(_=i.rosterCount)!=null?_:0,u=($=i.payrollCount)!=null?$:0,l=(j=i.difference)!=null?j:u-f,g=Array.isArray(i.mismatches)?i.mismatches.filter(Boolean):[],y="";R.loading?y=((Q=(L=window.PrairieForge)==null?void 0:L.renderStatusBanner)==null?void 0:Q.call(L,{type:"info",message:"Analyzing headcount\u2026",escapeHtml:b}))||"":R.lastError&&(y=((z=(F=window.PrairieForge)==null?void 0:F.renderStatusBanner)==null?void 0:z.call(F,{type:"error",message:R.lastError,escapeHtml:b}))||"");let w=(T,m,v,P)=>{let A=!c,q;A?q='<span class="pf-je-check-circle pf-je-circle--pending"></span>':P?q=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:q=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let ue=c?` = ${v}`:"";return`
            <div class="pf-je-check-row">
                ${q}
                <span class="pf-je-check-desc-pill">${b(T)}${ue}</span>
            </div>
        `},h=`
        ${w("SS_Employee_Roster count","Active employees in roster",f,!0)}
        ${w("PTO_Data count","Unique employees in PTO data",u,!0)}
        ${w("Difference","Should be zero",l,l===0)}
    `,E=g.length&&!R.skipAnalysis&&c?window.PrairieForge.renderMismatchTiles({mismatches:g,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"PTO Data",escapeHtml:b}):"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(he)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${b(e.title)}</h2>
            <p class="pf-hero-copy">${b(e.summary||"")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${R.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
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
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${He}</button>`,"Run")}
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${Ie}</button>`,"Refresh")}
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
            ${fe({textareaId:"step-notes-input",value:n,permanentId:"step-notes-lock-2",isPermanent:a,hintId:p?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
            ${ge({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:r,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
        </section>
    `}function Ra(){let e=Z.completenessCheck||{},t=Z.missingPayRates||[],n=[{key:"accrualRate",label:"Accrual Rate",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"carryOver",label:"Carry Over",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdAccrued",label:"YTD Accrued",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdUsed",label:"YTD Used",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"balance",label:"Balance",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"}],o=n.every(d=>e[d.key]!==null&&e[d.key]!==void 0)&&n.every(d=>{var p;return(p=e[d.key])==null?void 0:p.match}),s=t.length>0,r=d=>{let p=e[d.key],f=p==null,u;return f?u='<span class="pf-je-check-circle pf-je-circle--pending"></span>':p.match?u=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:u=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${u}
                <span class="pf-je-check-desc-pill">${b(d.label)}: ${b(d.desc)}</span>
            </div>
        `},i=n.map(d=>r(d)).join(""),c="";if(s){let d=t[0],p=t.length-1;c=`
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${b(d.name)}</strong> to calculate liability
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
                               data-employee="${Sa(d.name)}"
                               data-row="${d.rowIndex}">
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
            ${c}
        </article>
    `}function Aa(e){var l,g,y,w,h,E,C,I;let t=we(3),n=!!S.permanents[3],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(xe(S.completes[3])||o),r=W.hasRun,{balanceIssues:i,zeroBalances:c,accrualOutliers:d,totalEmployees:p}=W,f="";if(W.loading)f=((g=(l=window.PrairieForge)==null?void 0:l.renderStatusBanner)==null?void 0:g.call(l,{type:"info",message:"Analyzing data quality...",escapeHtml:b}))||"";else if(r){let _=i.length,$=d.length+c.length;_>0?f=((w=(y=window.PrairieForge)==null?void 0:y.renderStatusBanner)==null?void 0:w.call(y,{type:"error",title:`${_} Balance Issue${_>1?"s":""} Found`,message:"Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",escapeHtml:b}))||"":$>0?f=((E=(h=window.PrairieForge)==null?void 0:h.renderStatusBanner)==null?void 0:E.call(h,{type:"warning",title:"No Critical Issues",message:`${$} informational item${$>1?"s":""} to review (see below).`,escapeHtml:b}))||"":f=((I=(C=window.PrairieForge)==null?void 0:C.renderStatusBanner)==null?void 0:I.call(C,{type:"success",title:"Data Quality Passed",message:`${p} employee${p!==1?"s":""} checked \u2014 no anomalies found.`,escapeHtml:b}))||""}let u=[];return r&&i.length>0&&u.push(`
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
        `),r&&d.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--warning">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u{1F4CA}</span>
                    <span class="pf-quality-issue-title">High Accrual Rates (${d.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${d.slice(0,5).map(_=>`<li><strong>${b(_.name)}</strong>: ${_.accrualRate.toFixed(2)} hrs/period</li>`).join("")}
                    ${d.length>5?`<li class="pf-quality-more">+${d.length-5} more</li>`:""}
                </ul>
            </div>
        `),r&&c.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--info">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u2139\uFE0F</span>
                    <span class="pf-quality-issue-title">Zero Balances (${c.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${c.slice(0,5).map(_=>`<li><strong>${b(_.name)}</strong></li>`).join("")}
                    ${c.length>5?`<li class="pf-quality-more">+${c.length-5} more</li>`:""}
                </ul>
            </div>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(he)} | Step ${e.id}</p>
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
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${He}</button>`,"Run")}
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
                        ${W.acknowledged?'<p class="pf-quality-actions-hint"><span class="pf-acknowledged-badge">\u2713 Issues Acknowledged</span></p>':""}
                        <div class="pf-signoff-action">
                            ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${Ie}</button>`,"Refresh")}
                            ${W.acknowledged?"":B(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${Te}</button>`,"Continue")}
                        </div>
                    </div>
                </article>
            `:""}
            ${fe({textareaId:"step-notes-3",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-3",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-3"})}
            ${ge({reviewerInputId:"step-reviewer-3",reviewerValue:a,signoffInputId:"step-signoff-3",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-3",completeButtonId:"step-signoff-toggle-3"})}
        </section>
    `}function Na(e){let t=we(4),n=!!S.permanents[4],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(xe(S.completes[4])||o);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${b(he)} | Step ${e.id}</p>
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
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${He}</button>`,"Run")}
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${Ie}</button>`,"Refresh")}
                </div>
            </article>
            ${Ra()}
            ${fe({textareaId:"step-notes-4",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-4",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-4"})}
            ${ge({reviewerInputId:"step-reviewer-4",reviewerValue:a,signoffInputId:"step-signoff-4",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"step-signoff-toggle-4"})}
        </section>
    `}function Da(e){let t=we(5),n=!!S.permanents[5],a=be(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(xe(S.completes[5])||o),r=G.lastError?`<p class="pf-step-note">${b(G.lastError)}</p>`:"",i=G.validationRun,c=G.issues||[],d=[{key:"Debits = Credits",desc:"\u2211 Debit column = \u2211 Credit column"},{key:"Line Amounts Sum to Zero",desc:"\u2211 Line Amount = $0.00"},{key:"JE Matches Analysis Total",desc:"\u2211 Expense line amounts = \u2211 PTO_Analysis Change"}],p=g=>{let y=c.find(E=>E.check===g.key),w=!i,h;return w?h='<span class="pf-je-check-circle pf-je-circle--pending"></span>':y!=null&&y.passed?h=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:h=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${h}
                <span class="pf-je-check-desc-pill">${b(g.desc)}</span>
            </div>
        `},f=d.map(g=>p(g)).join(""),u=c.filter(g=>!g.passed),l="";return i&&u.length>0&&(l=`
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
            <p class="pf-hero-copy">${b(he)} | Step ${e.id}</p>
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
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${Ve}</button>`,"Generate")}
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${Ie}</button>`,"Refresh")}
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${Ft}</button>`,"Export")}
                    ${B(`<button type="button" class="pf-action-toggle pf-clickable" id="je-upload-btn" title="Open accounting software upload">${Ut}</button>`,"Upload")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${r}
                <div class="pf-je-checks-container">
                    ${f}
                </div>
            </article>
            ${l}
            ${fe({textareaId:"step-notes-5",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-5",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-5"})}
            ${ge({reviewerInputId:"step-reviewer-5",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
        </section>
    `}function $a(){var t,n;return Math.abs((n=(t=R.roster)==null?void 0:t.difference)!=null?n:0)>0}function wn(){return!R.skipAnalysis&&$a()}async function ft(){if(!Y()){R.loading=!1,R.lastError="Excel runtime is unavailable.",se();return}R.loading=!0,R.lastError=null,ce(document.getElementById("headcount-save-btn"),!1),se();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),a=t.workbook.worksheets.getItem("PTO_Data"),o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.getUsedRangeOrNullObject(),r=a.getUsedRangeOrNullObject();s.load("values"),r.load("values"),o.load("isNullObject"),await t.sync();let i=null;o.isNullObject||(i=o.getUsedRangeOrNullObject(),i.load("values")),await t.sync();let c=s.isNullObject?[]:s.values||[],d=r.isNullObject?[]:r.values||[],p=i&&!i.isNullObject?i.values||[]:[],f=p.length?p:d;return ja(c,f)});R.roster=e.roster,R.hasAnalyzed=!0,R.lastError=null}catch(e){console.warn("PTO headcount: unable to analyze data",e),R.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{R.loading=!1,se()}}function nn(e){if(!e)return!0;let t=e.toLowerCase().trim();return t?["total","subtotal","sum","count","grand","average","avg"].some(a=>t.includes(a)):!0}function ja(e,t){let n={rosterCount:0,payrollCount:0,difference:0,mismatches:[]};if(((e==null?void 0:e.length)||0)<2||((t==null?void 0:t.length)||0)<2)return console.warn("Headcount: insufficient data rows",{rosterRows:(e==null?void 0:e.length)||0,payrollRows:(t==null?void 0:t.length)||0}),{roster:n};let a=an(e),o=an(t),s=a.headers,r=o.headers,i={employee:on(s),termination:s.findIndex(l=>l.includes("termination"))},c={employee:on(r)};console.log("Headcount column detection:",{rosterEmployeeCol:i.employee,rosterTerminationCol:i.termination,payrollEmployeeCol:c.employee,rosterHeaders:s.slice(0,5),payrollHeaders:r.slice(0,5)});let d=new Set,p=new Set;for(let l=a.startIndex;l<e.length;l+=1){let g=e[l],y=i.employee>=0?ve(g[i.employee]):"";nn(y)||i.termination>=0&&ve(g[i.termination])||d.add(y.toLowerCase())}for(let l=o.startIndex;l<t.length;l+=1){let g=t[l],y=c.employee>=0?ve(g[c.employee]):"";nn(y)||p.add(y.toLowerCase())}n.rosterCount=d.size,n.payrollCount=p.size,n.difference=n.payrollCount-n.rosterCount,console.log("Headcount results:",{rosterCount:n.rosterCount,payrollCount:n.payrollCount,difference:n.difference});let f=[...d].filter(l=>!p.has(l)),u=[...p].filter(l=>!d.has(l));return n.mismatches=[...f.map(l=>`In roster, missing in PTO_Data: ${l}`),...u.map(l=>`In PTO_Data, missing in roster: ${l}`)],{roster:n}}function an(e){if(!Array.isArray(e)||!e.length)return{headers:[],startIndex:1};let t=e.findIndex((o=[])=>o.some(s=>ve(s).toLowerCase().includes("employee"))),n=t===-1?0:t;return{headers:(e[n]||[]).map(o=>ve(o).toLowerCase()),startIndex:n+1}}function on(e=[]){let t=-1,n=-1;return e.forEach((a,o)=>{let s=a.toLowerCase();if(!s.includes("employee"))return;let r=1;s.includes("name")?r=4:s.includes("id")?r=2:r=3,r>n&&(n=r,t=o)}),t}function ve(e){return e==null?"":String(e).trim()}async function kn(e=null){let t=async n=>{let a=n.workbook.worksheets.getItem("PTO_Data"),o=n.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster"),r=n.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary"),i=n.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary"),c=a.getUsedRangeOrNullObject();c.load("values"),o.load("isNullObject"),s.load("isNullObject"),r.load("isNullObject"),i.load("isNullObject"),await n.sync();let d=c.isNullObject?[]:c.values||[];if(!d.length)return;let p=(d[0]||[]).map(O=>J(O)),f=p.findIndex(O=>O.includes("employee")&&O.includes("name")),u=f>=0?f:0,l=D(p,["accrual rate"]),g=D(p,["carry over","carryover"]),y=p.findIndex(O=>O.includes("ytd")&&(O.includes("accrued")||O.includes("accrual"))),w=p.findIndex(O=>O.includes("ytd")&&O.includes("used")),h=D(p,["balance","current balance","pto balance"]);console.log("[PTO Analysis] PTO_Data headers:",p),console.log("[PTO Analysis] Column indices found:",{employee:u,accrualRate:l,carryOver:g,ytdAccrued:y,ytdUsed:w,balance:h}),w>=0?console.log(`[PTO Analysis] YTD Used column: "${p[w]}" at index ${w}`):console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:",p);let E=d.slice(1).map(O=>ve(O[u])).filter(O=>O&&!O.toLowerCase().includes("total")),C=new Map;d.slice(1).forEach(O=>{let U=J(O[u]);!U||U.includes("total")||C.set(U,O)});let I=new Map;if(s.isNullObject)console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");else{let O=s.getUsedRangeOrNullObject();O.load("values"),await n.sync();let U=O.isNullObject?[]:O.values||[];if(U.length){let M=(U[0]||[]).map(x=>J(x));console.log("[PTO Analysis] SS_Employee_Roster headers:",M);let V=M.findIndex(x=>x.includes("employee")&&x.includes("name"));V<0&&(V=M.findIndex(x=>x==="employee"||x==="name"||x==="full name"));let H=M.findIndex(x=>x.includes("department"));console.log(`[PTO Analysis] Roster column indices - Name: ${V}, Dept: ${H}`),V>=0&&H>=0?(U.slice(1).forEach(x=>{let re=J(x[V]),pe=ve(x[H]);re&&I.set(re,pe)}),console.log(`[PTO Analysis] Built roster map with ${I.size} employees`)):console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster")}}let _=new Map;if(!r.isNullObject){let O=r.getUsedRangeOrNullObject();O.load("values"),await n.sync();let U=O.isNullObject?[]:O.values||[];if(U.length){let M=(U[0]||[]).map(H=>J(H)),V={payrollDate:D(M,["payroll date"]),employee:D(M,["employee"]),category:D(M,["payroll category","category"]),amount:D(M,["amount","gross salary","gross_salary","earnings"])};V.employee>=0&&V.category>=0&&V.amount>=0&&U.slice(1).forEach(H=>{let x=J(H[V.employee]);if(!x)return;let re=J(H[V.category]);if(!re.includes("regular")||!re.includes("earn"))return;let pe=Number(H[V.amount])||0;if(!pe)return;let Ee=_a(H[V.payrollDate]),Ce=_.get(x);(!Ce||Ee!=null&&Ee>Ce.timestamp)&&_.set(x,{payRate:pe/80,timestamp:Ee})})}}let $=new Map;if(!i.isNullObject){let O=i.getUsedRangeOrNullObject();O.load("values"),await n.sync();let U=O.isNullObject?[]:O.values||[];if(U.length>1){let M=(U[0]||[]).map(x=>J(x)),V=M.findIndex(x=>x.includes("employee")&&x.includes("name")),H=D(M,["liability amount","liability","accrued pto"]);V>=0&&H>=0&&U.slice(1).forEach(x=>{let re=J(x[V]);if(!re)return;let pe=Number(x[H])||0;$.set(re,pe)})}}let j=te(k.payrollDate)||"",L=[],Q=[],F=E.map((O,U)=>{var vt,bt,wt,kt,Ot,St,xt;let M=J(O),V=I.get(M)||"",H=(bt=(vt=_.get(M))==null?void 0:vt.payRate)!=null?bt:"",x=C.get(M),re=x&&l>=0&&(wt=x[l])!=null?wt:"",pe=x&&g>=0&&(kt=x[g])!=null?kt:"",Ee=x&&y>=0&&(Ot=x[y])!=null?Ot:"",Ce=x&&w>=0&&(St=x[w])!=null?St:"";(M.includes("avalos")||M.includes("sarah"))&&console.log(`[PTO Debug] ${O}:`,{ytdUsedIdx:w,rawValue:x?x[w]:"no dataRow",ytdUsed:Ce,fullRow:x});let Ye=x&&h>=0&&Number(x[h])||0,ht=U+2;!H&&typeof H!="number"&&L.push({name:O,rowIndex:ht}),V||Q.push({name:O,rowIndex:ht});let We=typeof H=="number"&&Ye?Ye*H:0,yt=(xt=$.get(M))!=null?xt:0,On=(typeof We=="number"?We:0)-yt;return[j,O,V,H,re,pe,Ee,Ce,Ye,We,yt,On]});Z.missingPayRates=L.filter(O=>!Z.ignoredMissingPayRates.has(O.name)),Z.missingDepartments=Q,console.log(`[PTO Analysis] Data quality: ${L.length} missing pay rates, ${Q.length} missing departments`);let z=[["Analysis Date","Employee Name","Department","Pay Rate","Accrual Rate","Carry Over","YTD Accrued","YTD Used","Balance","Liability Amount","Accrued PTO $ [Prior Period]","Change"],...F],T=o.isNullObject?n.workbook.worksheets.add("PTO_Analysis"):o,m=T.getUsedRangeOrNullObject();m.load("address"),await n.sync(),m.isNullObject||m.clear();let v=z[0].length,P=z.length,A=F.length,q=T.getRangeByIndexes(0,0,P,v);q.values=z;let ue=T.getRangeByIndexes(0,0,1,v);dt(ue),A>0&&(Wt(T,0,A),me(T,3,A),Oe(T,4,A),Oe(T,5,A),Oe(T,6,A),Oe(T,7,A),Oe(T,8,A),me(T,9,A),me(T,10,A),me(T,11,A,!0)),q.format.autofitColumns(),T.getRange("A1").select(),await n.sync()};Y()&&(e?await t(e):await Excel.run(t))}function La(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let a=String(n);return/[",\n]/.test(a)?`"${a.replace(/"/g,'""')}"`:a}).join(",")).join(`
`)}function Ba(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),a=URL.createObjectURL(n),o=document.createElement("a");o.href=a,o.download=e,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(a),1e3)}function sn(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=wn(),n=document.getElementById("step-notes-input"),a=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!a;let o=document.getElementById("headcount-notes-hint");o&&(o.textContent=t?"Please document outstanding differences before signing off.":"")}function rn(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(ye)?t.slice(ye.length).replace(/^\s+/,""):t.replace(new RegExp(`^${ye}\\s*`,"i"),"").trimStart(),a=ye+(n?`
${n}`:"");e.value!==a&&(e.value=a),ne(2,"notes",e.value)}function Ma(){let e=document.getElementById("step-notes-input");e&&e.addEventListener("input",()=>{if(!R.skipAnalysis)return;let t=e.value||"";if(!t.startsWith(ye)){let n=t.replace(ye,"").trimStart();e.value=ye+(n?`
${n}`:"")}ne(2,"notes",e.value)})}})();
//# sourceMappingURL=app.bundle.js.map
