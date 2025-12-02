/* Prairie Forge PTO Accrual */
(()=>{function q(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}var Ue="SS_PF_Config";async function vt(e,t=[Ue]){var o;let n=e.workbook.tables;n.load("items/name"),await e.sync();let a=(o=n.items)==null?void 0:o.find(s=>t.includes(s.name));return a?e.workbook.tables.getItem(a.name):(console.warn("Config table not found. Looking for:",t),null)}function bt(e){let t=e.map(n=>String(n||"").trim().toLowerCase());return{field:t.findIndex(n=>n==="field"||n==="field name"||n==="setting"),value:t.findIndex(n=>n==="value"||n==="setting value"),type:t.findIndex(n=>n==="type"||n==="category"),title:t.findIndex(n=>n==="title"||n==="display name"),permanent:t.findIndex(n=>n==="permanent"||n==="persist")}}async function wt(e=[Ue]){if(!q())return{};try{return await Excel.run(async t=>{let n=await vt(t,e);if(!n)return{};let a=n.getDataBodyRange(),o=n.getHeaderRowRange();a.load("values"),o.load("values"),await t.sync();let s=o.values[0]||[],r=bt(s);if(r.field===-1||r.value===-1)return console.warn("Config table missing FIELD or VALUE columns. Headers:",s),{};let l={};return(a.values||[]).forEach(i=>{var g;let f=String(i[r.field]||"").trim();f&&(l[f]=(g=i[r.value])!=null?g:"")}),console.log("Configuration loaded:",Object.keys(l).length,"fields"),l})}catch(t){return console.error("Failed to load configuration:",t),{}}}async function Ee(e,t,n=[Ue]){if(!q())return!1;try{return await Excel.run(async a=>{let o=await vt(a,n);if(!o){console.warn("Config table not found for write");return}let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let l=r.values[0]||[],c=bt(l);if(c.field===-1||c.value===-1){console.error("Config table missing FIELD or VALUE columns");return}let f=(s.values||[]).findIndex(g=>String(g[c.field]||"").trim()===e);if(f>=0)s.getCell(f,c.value).values=[[t]];else{let g=new Array(l.length).fill("");c.type>=0&&(g[c.type]="Run Settings"),g[c.field]=e,g[c.value]=t,c.permanent>=0&&(g[c.permanent]="N"),c.title>=0&&(g[c.title]=""),o.rows.add(null,[g]),console.log("Added new config row:",e,"=",t)}await a.sync(),console.log("Saved config:",e,"=",t)}),!0}catch(a){return console.error("Failed to save config:",e,a),!1}}var yn="SS_PF_Config",hn="module-prefix",Ge="system",he={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function Ot(){if(!q())return{...he};try{return await Excel.run(async e=>{var f,g;let t=e.workbook.worksheets.getItemOrNullObject(yn);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...he};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((f=n.values)!=null&&f.length))return{...he};let a=n.values,o=wn(a[0]),s=o.get("category"),r=o.get("field"),l=o.get("value");if(s===void 0||r===void 0||l===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...he};let c={},i=!1;for(let u=1;u<a.length;u++){let d=a[u];if(xe(d[s])===hn){let v=String((g=d[r])!=null?g:"").trim().toUpperCase(),b=xe(d[l]);v&&b&&(c[v]=b,i=!0)}}return i?(console.log("[Tab Visibility] Loaded prefix config:",c),c):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...he})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...he}}}async function Je(e){if(!q())return;let t=xe(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await Ot();await Excel.run(async a=>{let o=a.workbook.worksheets;o.load("items/name,visibility"),await a.sync();let s={};for(let[u,d]of Object.entries(n))s[d]||(s[d]=[]),s[d].push(u);let r=s[t]||[],l=s[Ge]||[],c=[];for(let[u,d]of Object.entries(s))u!==t&&u!==Ge&&c.push(...d);console.log(`[Tab Visibility] Active prefixes: ${r.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${c.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${l.join(", ")}`);let i=[],f=[];o.items.forEach(u=>{let d=u.name,y=d.toUpperCase(),v=r.some(k=>y.startsWith(k)),b=c.some(k=>y.startsWith(k)),p=l.some(k=>y.startsWith(k));v?(i.push(u),console.log(`[Tab Visibility] SHOW: ${d} (matches active module prefix)`)):p?(f.push(u),console.log(`[Tab Visibility] HIDE: ${d} (system sheet)`)):b?(f.push(u),console.log(`[Tab Visibility] HIDE: ${d} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${d} (no prefix match, leaving as-is)`)});for(let u of i)u.visibility=Excel.SheetVisibility.visible;if(await a.sync(),o.items.filter(u=>u.visibility===Excel.SheetVisibility.visible).length>f.length){for(let u of f)try{u.visibility=Excel.SheetVisibility.hidden}catch(d){console.warn(`[Tab Visibility] Could not hide "${u.name}":`,d.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${i.length}, hid ${f.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function vn(){if(!q()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.visibility!==Excel.SheetVisibility.visible&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${a.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function bn(){if(!q()){console.log("Excel not available");return}try{let e=await Ot(),t=[];for(let[n,a]of Object.entries(e))a===Ge&&t.push(n);await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name,visibility"),await n.sync(),a.items.forEach(o=>{let s=o.name.toUpperCase();t.some(r=>s.startsWith(r))&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function wn(e=[]){let t=new Map;return e.forEach((n,a)=>{let o=xe(n);o&&t.set(o,a)}),t}function xe(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=vn,window.PrairieForge.unhideSystemSheets=bn,window.PrairieForge.applyModuleTabVisibility=Je);var kt={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var _t=kt.ADA_IMAGE_URL;async function Ne(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async a=>{let o=a.workbook.worksheets.getItemOrNullObject(e);o.load("isNullObject, name"),await a.sync();let s;o.isNullObject?(s=a.workbook.worksheets.add(e),await a.sync(),await St(a,s,t,n)):(s=o,await St(a,s,t,n)),s.activate(),s.getRange("A1").select(),await a.sync()})}catch(a){console.error(`Error activating homepage sheet ${e}:`,a)}}async function St(e,t,n,a){try{let i=t.getUsedRangeOrNullObject();i.load("isNullObject"),await e.sync(),i.isNullObject||(i.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let o=[[n,""],[a,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=o;let r=t.getRange("A1:Z100");r.format.fill.color="#0f0f0f";let l=t.getRange("A1");l.format.font.bold=!0,l.format.font.size=36,l.format.font.color="#ffffff",l.format.font.name="Segoe UI Light",l.format.verticalAlignment="Center";let c=t.getRange("A2");c.format.font.size=14,c.format.font.color="#a0a0a0",c.format.font.name="Segoe UI",c.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var Et={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function Ae(e){return Et[e]||Et["module-selector"]}function Ct(){Ye();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${_t}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",On),e}function Ye(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function On(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${_t}" alt="Ada" />
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",qe),t.addEventListener("click",o=>{o.target===t&&qe()});let a=o=>{o.key==="Escape"&&(qe(),document.removeEventListener("keydown",a))};document.addEventListener("keydown",a)}function qe(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var kn=["January","February","March","April","May","June","July","August","September","October","November","December"],Rt=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],Sn=["Su","Mo","Tu","We","Th","Fr","Sa"],ve=null;function Tt(e,t={}){let n=document.getElementById(e);if(!n)return;let{onChange:a=null,minDate:o=null,maxDate:s=null,readonly:r=!1}=t,l=n.closest(".pf-datepicker-wrapper");l||(l=document.createElement("div"),l.className="pf-datepicker-wrapper",n.parentNode.insertBefore(l,n),l.appendChild(n)),n.type="text",n.readOnly=!0,n.classList.add("pf-datepicker-input");let c=n.value?Pt(n.value):null,i=c?new Date(c):new Date;c&&(n.value=It(c));let f=document.createElement("span");f.className="pf-datepicker-icon",f.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>',l.appendChild(f);let g=document.createElement("div");g.className="pf-datepicker-dropdown",g.id=`${e}-dropdown`,l.appendChild(g);function u(){var C,T,P,A,N,x;let p=i.getFullYear(),k=i.getMonth();g.innerHTML=`
            <div class="pf-datepicker-header">
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev-year" title="Previous Year">\xAB</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev" title="Previous Month">\u2039</button>
                <span class="pf-datepicker-title">${kn[k]} ${p}</span>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next" title="Next Month">\u203A</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next-year" title="Next Year">\xBB</button>
            </div>
            <div class="pf-datepicker-weekdays">
                ${Sn.map(m=>`<span>${m}</span>`).join("")}
            </div>
            <div class="pf-datepicker-days">
                ${d(p,k,c)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-today">Today</button>
                <button type="button" class="pf-datepicker-clear">Clear</button>
            </div>
        `,(C=g.querySelector(".pf-datepicker-prev-year"))==null||C.addEventListener("click",m=>{m.stopPropagation(),i.setFullYear(i.getFullYear()-1),u()}),(T=g.querySelector(".pf-datepicker-prev"))==null||T.addEventListener("click",m=>{m.stopPropagation(),i.setMonth(i.getMonth()-1),u()}),(P=g.querySelector(".pf-datepicker-next"))==null||P.addEventListener("click",m=>{m.stopPropagation(),i.setMonth(i.getMonth()+1),u()}),(A=g.querySelector(".pf-datepicker-next-year"))==null||A.addEventListener("click",m=>{m.stopPropagation(),i.setFullYear(i.getFullYear()+1),u()}),g.querySelectorAll(".pf-datepicker-day:not(.disabled)").forEach(m=>{m.addEventListener("click",O=>{O.stopPropagation();let h=parseInt(m.dataset.day),E=parseInt(m.dataset.month),V=parseInt(m.dataset.year);y(new Date(V,E,h))})}),(N=g.querySelector(".pf-datepicker-today"))==null||N.addEventListener("click",m=>{m.stopPropagation(),y(new Date)}),(x=g.querySelector(".pf-datepicker-clear"))==null||x.addEventListener("click",m=>{m.stopPropagation(),y(null)})}function d(p,k,C){let T=new Date(p,k,1).getDay(),P=new Date(p,k+1,0).getDate(),A=new Date(p,k,0).getDate(),N=new Date;N.setHours(0,0,0,0);let x="";for(let h=T-1;h>=0;h--){let E=A-h,V=k===0?11:k-1,W=k===0?p-1:p;x+=`<span class="pf-datepicker-day other-month" data-day="${E}" data-month="${V}" data-year="${W}">${E}</span>`}for(let h=1;h<=P;h++){let E=new Date(p,k,h),V=E.getTime()===N.getTime(),W=C&&E.getTime()===C.getTime(),Q="pf-datepicker-day";V&&(Q+=" today"),W&&(Q+=" selected"),o&&E<o&&(Q+=" disabled"),s&&E>s&&(Q+=" disabled"),x+=`<span class="${Q}" data-day="${h}" data-month="${k}" data-year="${p}">${h}</span>`}let O=Math.ceil((T+P)/7)*7-(T+P);for(let h=1;h<=O;h++){let E=k===11?0:k+1,V=k===11?p+1:p;x+=`<span class="pf-datepicker-day other-month" data-day="${h}" data-month="${E}" data-year="${V}">${h}</span>`}return x}function y(p){c=p,p?(n.value=It(p),n.dataset.value=We(p),i=new Date(p)):(n.value="",n.dataset.value=""),b(),a&&a(p?We(p):""),n.dispatchEvent(new Event("change",{bubbles:!0}))}function v(){if(!r){if(ve&&ve!==e){let p=document.getElementById(`${ve}-dropdown`);p==null||p.classList.remove("open")}ve=e,u(),g.classList.add("open"),l.classList.add("open")}}function b(){g.classList.remove("open"),l.classList.remove("open"),ve===e&&(ve=null)}return n.addEventListener("click",p=>{p.stopPropagation(),g.classList.contains("open")?b():v()}),f.addEventListener("click",p=>{p.stopPropagation(),g.classList.contains("open")?b():v()}),document.addEventListener("click",p=>{l.contains(p.target)||b()}),document.addEventListener("keydown",p=>{p.key==="Escape"&&b()}),{getValue:()=>c?We(c):"",setValue:p=>{let k=Pt(p);y(k)},open:v,close:b}}function Pt(e){if(!e)return null;if(/^\d{4}-\d{2}-\d{2}$/.test(e)){let[a,o,s]=e.split("-").map(Number);return new Date(a,o-1,s)}let t=e.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);if(t){let a=Rt.findIndex(o=>o.toLowerCase()===t[1].toLowerCase().substring(0,3));if(a>=0)return new Date(parseInt(t[3]),a,parseInt(t[2]))}if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(e)){let[a,o,s]=e.split("/").map(Number);return new Date(s,a-1,o)}let n=new Date(e);return isNaN(n.getTime())?null:n}function It(e){return e?`${Rt[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`:""}function We(e){if(!e)return"";let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}var xt=`
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
`.trim(),Nt=`
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
`.trim(),At=`
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
`.trim(),ze=`
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
`.trim(),Dt=`
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
`.trim(),$t=`
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
`.trim(),En={config:`
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
    `};function jt(e){return e&&En[e]||""}var Ke=`
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
`.trim(),Qe=`
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
`.trim(),_e=`
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
`.trim(),De=`
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
`.trim(),Ia=`
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
        <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
        <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 1 0 7.07 7.07l1.71-1.71" />
    </svg>
`.trim(),Lt=`
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
`.trim(),Bt=`
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
        <path d="m12 19-7-7 7-7" />
        <path d="M19 12H5" />
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
        <path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" />
        <path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7" />
        <path d="M7 3v4a1 1 0 0 0 1 1h7" />
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
        <path d="m12 5 7 7-7 7" />
        <path d="M5 12h14" />
    </svg>
`.trim(),Ra=`
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
`.trim(),Ta=`
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
`.trim(),xa=`
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
`.trim(),Na=`
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
`.trim(),Aa=`
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
`.trim(),Da=`
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
`.trim(),$a=`
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
`.trim(),ja=`
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
`.trim(),Ce=`
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
        <path d="M3 6h18"/>
        <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/>
        <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>
        <line x1="10" x2="10" y1="11" y2="17"/>
        <line x1="14" x2="14" y1="11" y2="17"/>
    </svg>
`.trim();function Pe(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function F(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function re({textareaId:e,value:t,permanentId:n,isPermanent:a,hintId:o,saveButtonId:s,isSaved:r=!1,placeholder:l="Enter notes here..."}){let c=a?Qe:Ke,i=s?`<button type="button" class="pf-action-toggle pf-save-btn ${r?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Vt}</button>`:"",f=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${a?"is-locked":""}" id="${n}" aria-pressed="${a}" title="Lock notes (retain after archive)">${c}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Pe(l)}">${Pe(t||"")}</textarea>
                ${o?`<p class="pf-signoff-hint" id="${o}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?F(f,"Lock"):""}
                ${s?F(i,"Save"):""}
            </div>
        </article>
    `}function le({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:a,isComplete:o,saveButtonId:s,isSaved:r=!1,completeButtonId:l,subtext:c="Sign-off below. Click checkmark icon. Done."}){let i=`<button type="button" class="pf-action-toggle ${o?"is-active":""}" id="${l}" aria-pressed="${!!o}" title="Mark step complete">${_e}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${Pe(c)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${Pe(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${Pe(a)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${F(i,"Done")}
            </div>
        </article>
    `}function Ze(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?Qe:Ke)}function oe(e,t){e&&e.classList.toggle("is-saved",t)}function et(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(a=>{let o=a.getAttribute("data-save-input"),s=document.getElementById(o);if(!s)return;let r=()=>{oe(a,!1)};s.addEventListener("input",r),n.push(()=>s.removeEventListener("input",r))}),()=>n.forEach(a=>a())}function Ut(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function Gt(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var tt={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},$e={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function nt(e){e.format.fill.color=tt.fillColor,e.format.font.color=tt.fontColor,e.format.font.bold=tt.bold}function ce(e,t,n,a=!1){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a?$e.currencyWithNegative:$e.currency]]}function be(e,t,n){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[$e.number]]}function Jt(e,t,n,a=$e.date){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a]]}var _n="1.1.0",Re="pto-accrual";var ue="PTO Accrual",Cn="Calculate your PTO liability, compare against last period, and generate a balanced journal entry\u2014all without leaving Excel.",Pn="../module-selector/index.html",In="pf-loader-overlay",de=["SS_PF_Config"],S={payrollProvider:"PTO_Payroll_Provider",payrollDate:"PTO_Analysis_Date",accountingPeriod:"PTO_Accounting_Period",journalEntryId:"PTO_Journal_Entry_ID",companyName:"SS_Company_Name",accountingSoftware:"SS_Accounting_Software",reviewerName:"PTO_Reviewer",validationDataBalance:"PTO_Validation_Data_Balance",validationCleanBalance:"PTO_Validation_Clean_Balance",validationDifference:"PTO_Validation_Difference",headcountRosterCount:"PTO_Headcount_Roster_Count",headcountPayrollCount:"PTO_Headcount_Payroll_Count",headcountDifference:"PTO_Headcount_Difference",journalDebitTotal:"PTO_JE_Debit_Total",journalCreditTotal:"PTO_JE_Credit_Total",journalDifference:"PTO_JE_Difference"},pe="User opted to skip the headcount review this period.",Be={0:{note:"PTO_Notes_Config",reviewer:"PTO_Reviewer_Config",signOff:"PTO_SignOff_Config"},1:{note:"PTO_Notes_Import",reviewer:"PTO_Reviewer_Import",signOff:"PTO_SignOff_Import"},2:{note:"PTO_Notes_Headcount",reviewer:"PTO_Reviewer_Headcount",signOff:"PTO_SignOff_Headcount"},3:{note:"PTO_Notes_Validate",reviewer:"PTO_Reviewer_Validate",signOff:"PTO_SignOff_Validate"},4:{note:"PTO_Notes_Review",reviewer:"PTO_Reviewer_Review",signOff:"PTO_SignOff_Review"},5:{note:"PTO_Notes_JE",reviewer:"PTO_Reviewer_JE",signOff:"PTO_SignOff_JE"},6:{note:"PTO_Notes_Archive",reviewer:"PTO_Reviewer_Archive",signOff:"PTO_SignOff_Archive"}},sn={0:"PTO_Complete_Config",1:"PTO_Complete_Import",2:"PTO_Complete_Headcount",3:"PTO_Complete_Validate",4:"PTO_Complete_Review",5:"PTO_Complete_JE",6:"PTO_Complete_Archive"};var se=[{id:0,title:"Configuration",summary:"Set the analysis date, accounting period, and review details for this run.",description:"Complete this step first to ensure all downstream calculations use the correct period settings.",actionLabel:"Configure Workbook",secondaryAction:{sheet:"SS_PF_Config",label:"Open Config Sheet"}},{id:1,title:"Import PTO Data",summary:"Pull your latest PTO export from payroll and paste it into PTO_Data.",description:"Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",actionLabel:"Import Sample Data",secondaryAction:{sheet:"PTO_Data",label:"Open Data Sheet"}},{id:2,title:"Headcount Review",summary:"Quick check to make sure your roster matches your PTO data.",description:"Compare employees in PTO_Data against your employee roster to catch any discrepancies.",actionLabel:"Open Headcount Review",secondaryAction:{sheet:"SS_Employee_Roster",label:"Open Sheet"}},{id:3,title:"Data Quality Review",summary:"Scan your PTO data for potential errors before crunching numbers.",description:"Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",actionLabel:"Click to Run Quality Check"},{id:4,title:"PTO Accrual Review",summary:"Review the calculated liability for each employee and compare to last period.",description:"The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",actionLabel:"Click to Perform Review"},{id:5,title:"Journal Entry Prep",summary:"Generate a balanced journal entry, run validation checks, and export when ready.",description:"Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",actionLabel:"Open Journal Draft",secondaryAction:{sheet:"PTO_JE_Draft",label:"Open Sheet"}},{id:6,title:"Archive & Reset",summary:"Save this period's results and prepare for the next cycle.",description:"Archive the current analysis so it becomes the 'prior period' for your next review.",actionLabel:"Archive Run"}];var Rn=se.reduce((e,t)=>(e[t.id]="pending",e),{}),j={activeView:"home",activeStepId:null,focusedIndex:0,stepStatuses:Rn},_={loaded:!1,steps:{},permanents:{},completes:{},values:{},overrides:{accountingPeriod:!1,journalId:!1}},Ie=null,at=null,je=null,we=new Map,D={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},U={debitTotal:null,creditTotal:null,difference:null,lineAmountSum:null,analysisChangeTotal:null,jeChangeTotal:null,loading:!1,lastError:null,validationRun:!1,issues:[]},J={hasRun:!1,loading:!1,acknowledged:!1,balanceIssues:[],zeroBalances:[],accrualOutliers:[],totalIssues:0,totalEmployees:0},Y={cleanDataReady:!1,employeeCount:0,lastRun:null,loading:!1,lastError:null,missingPayRates:[],missingDepartments:[],ignoredMissingPayRates:new Set,completenessCheck:{accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null}};async function Tn(){var e;try{Ie=document.getElementById("app"),at=document.getElementById("loading"),await xn(),await Nn(),(e=window.PrairieForge)!=null&&e.loadSharedConfig&&await window.PrairieForge.loadSharedConfig();let t=Ae(Re);await Ne(t.sheetName,t.title,t.subtitle),at&&at.remove(),Ie&&(Ie.hidden=!1),ne()}catch(t){throw console.error("[PTO] Module initialization failed:",t),t}}async function xn(){try{await Je(Re),console.log(`[PTO] Tab visibility applied for ${Re}`)}catch(e){console.warn("[PTO] Could not apply tab visibility:",e)}}async function Nn(){var e;if(!q()){_.loaded=!0;return}try{let t=await wt(de),n={};(e=window.PrairieForge)!=null&&e.loadSharedConfig&&(await window.PrairieForge.loadSharedConfig(),window.PrairieForge._sharedConfigCache&&window.PrairieForge._sharedConfigCache.forEach((s,r)=>{n[r]=s}));let a={...t},o={SS_Default_Reviewer:S.reviewerName,Default_Reviewer:S.reviewerName,PTO_Reviewer:S.reviewerName,SS_Company_Name:S.companyName,Company_Name:S.companyName,SS_Payroll_Provider:S.payrollProvider,Payroll_Provider_Link:S.payrollProvider,SS_Accounting_Software:S.accountingSoftware,Accounting_Software_Link:S.accountingSoftware};Object.entries(o).forEach(([s,r])=>{n[s]&&!a[r]&&(a[r]=n[s])}),Object.entries(n).forEach(([s,r])=>{s.startsWith("PTO_")&&r&&(a[s]=r)}),_.permanents=await An(),_.values=a||{},_.overrides.accountingPeriod=!!(a!=null&&a[S.accountingPeriod]),_.overrides.journalId=!!(a!=null&&a[S.journalEntryId]),Object.entries(Be).forEach(([s,r])=>{var l,c,i;_.steps[s]={notes:(l=a[r.note])!=null?l:"",reviewer:(c=a[r.reviewer])!=null?c:"",signOffDate:(i=a[r.signOff])!=null?i:""}}),_.completes=Object.entries(sn).reduce((s,[r,l])=>{var c;return s[r]=(c=a[l])!=null?c:"",s},{}),_.loaded=!0}catch(t){console.warn("PTO: unable to load configuration fields",t),_.loaded=!0}}async function An(){let e={};if(!q())return e;let t=new Map;Object.entries(Be).forEach(([n,a])=>{a.note&&t.set(a.note.trim(),Number(n))});try{await Excel.run(async n=>{let a=n.workbook.tables.getItemOrNullObject(de[0]);if(await n.sync(),a.isNullObject)return;let o=a.getDataBodyRange(),s=a.getHeaderRowRange();o.load("values"),s.load("values"),await n.sync();let l=(s.values[0]||[]).map(i=>String(i||"").trim().toLowerCase()),c={field:l.findIndex(i=>i==="field"||i==="field name"||i==="setting"),permanent:l.findIndex(i=>i==="permanent"||i==="persist")};c.field===-1||c.permanent===-1||(o.values||[]).forEach(i=>{let f=String(i[c.field]||"").trim(),g=t.get(f);if(g==null)return;let u=sa(i[c.permanent]);e[g]=u})})}catch(n){console.warn("PTO: unable to load permanent flags",n)}return e}function ne(){var l;if(!Ie)return;let e=j.focusedIndex<=0?"disabled":"",t=j.focusedIndex>=se.length-1?"disabled":"",n=j.activeView==="step"&&j.activeStepId!=null,o=j.activeView==="config"?rn():n?Vn(j.activeStepId):`${$n()}${jn()}`;Ie.innerHTML=`
        <div class="pf-root">
            <div class="pf-brand-float" aria-hidden="true">
                <span class="pf-brand-wave"></span>
            </div>
            <header class="pf-banner">
                <div class="pf-nav-bar">
                    <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                        ${Mt}
                        <span class="sr-only">Previous step</span>
                    </button>
                    <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                        ${xt}
                        <span class="sr-only">Module Home</span>
                    </button>
                    <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                        ${Nt}
                        <span class="sr-only">Module Selector</span>
                    </button>
                    <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                        ${Ht}
                        <span class="sr-only">Next step</span>
                    </button>
                    <span class="pf-nav-divider"></span>
                    <div class="pf-quick-access-wrapper">
                        <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                            ${At}
                            <span class="sr-only">Quick Access Menu</span>
                        </button>
                        <div id="quick-access-dropdown" class="pf-quick-dropdown hidden">
                            <div class="pf-quick-dropdown-header">Quick Access</div>
                            <button id="nav-roster" class="pf-quick-item pf-clickable" type="button">
                                ${Dt}
                                <span>Employee Roster</span>
                            </button>
                            <button id="nav-accounts" class="pf-quick-item pf-clickable" type="button">
                                ${$t}
                                <span>Chart of Accounts</span>
                            </button>
                        </div>
                    </div>
                </div>
            </header>
            ${o}
            <footer class="pf-brand-footer">
                <div class="pf-brand-text">
                    <div class="pf-brand-label">prairie.forge</div>
                    <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${_n}</div>
                    <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
                </div>
            </footer>
        </div>
    `;let s=j.activeView==="home"||j.activeView!=="step"&&j.activeView!=="config",r=document.getElementById("pf-info-fab-pto");if(s)r&&r.remove();else if((l=window.PrairieForge)!=null&&l.mountInfoFab){let c=Dn(j.activeStepId);PrairieForge.mountInfoFab({title:c.title,content:c.content,buttonId:"pf-info-fab-pto"})}Hn(),Gn(),s?Ct():Ye()}function Dn(e){switch(e){case 0:return{title:"Configuration",content:`
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
                `}}}function $n(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">PTO Accrual</h2>
            <p class="pf-hero-copy">${Cn}</p>
        </section>
    `}function jn(){return`
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${se.map((e,t)=>Ln(e,t)).join("")}
            </div>
        </section>
    `}function Ln(e,t){let n=j.stepStatuses[e.id]||"pending",a=j.activeView==="step"&&j.focusedIndex===t?"pf-step-card--active":"",o=jt(ea(e.id));return`
        <article class="pf-step-card pf-clickable ${a}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${e.title}</h3>
        </article>
    `}function Bn(e){let t=se.filter(o=>o.id!==6).map(o=>({id:o.id,title:o.title,complete:Jn(o.id)})),n=t.every(o=>o.complete),a=t.map(o=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${o.complete?"is-active":""}" aria-pressed="${o.complete}">
                        ${_e}
                    </span>
                    <div>
                        <h3>${w(o.title)}</h3>
                        <p class="pf-config-subtext">${o.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
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
    `}function rn(){if(!_.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=Zt(te(S.payrollDate)),t=Zt(te(S.accountingPeriod)),n=te(S.journalEntryId),a=te(S.accountingSoftware),o=te(S.payrollProvider),s=te(S.companyName),r=te(S.reviewerName),l=ye(0),c=!!_.permanents[0],i=!!(pn(_.completes[0])||l.signOffDate),f=me(l==null?void 0:l.reviewer),g=(l==null?void 0:l.signOffDate)||"";return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${w(ue)} | Step 0</p>
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
                        <input type="text" id="config-user-name" value="${w(r)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${w(e)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${w(t)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${w(n)}" placeholder="PTO-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${w(s)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${w(o)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${w(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${re({textareaId:"config-notes",value:l.notes||"",permanentId:"config-notes-lock",isPermanent:c,hintId:"",saveButtonId:"config-notes-save"})}
            ${le({reviewerInputId:"config-reviewer",reviewerValue:f,signoffInputId:"config-signoff-date",signoffValue:g,isComplete:i,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"})}
        </section>
    `}function Mn(e){let t=ye(1),n=!!_.permanents[1],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Oe(_.completes[1])||o),r=te(S.payrollProvider);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export, then paste into PTO_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${F(r?`<a href="${w(r)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${Xe}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${Xe}</button>`,"Provider")}
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${ze}</button>`,"PTO_Data")}
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${Ft}</button>`,"Clear")}
                </div>
            </article>
            ${re({textareaId:"step-notes-1",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-1",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-1"})}
            ${le({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
        </section>
    `}function Vn(e){let t=se.find(l=>l.id===e);if(!t)return"";if(e===0)return rn();if(e===1)return Mn(t);if(e===2)return ca(t);if(e===3)return ua(t);if(e===4)return fa(t);if(e===5)return pa(t);if(t.id===6)return Bn(t);let n=ye(e),a=!!_.permanents[e],o=me(n==null?void 0:n.reviewer),s=(n==null?void 0:n.signOffDate)||"",r=!!(Oe(_.completes[e])||s);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${w(t.title)}</h2>
            <p class="pf-hero-copy">${w(t.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${re({textareaId:`step-notes-${e}`,value:(n==null?void 0:n.notes)||"",permanentId:`step-notes-lock-${e}`,isPermanent:a,hintId:"",saveButtonId:`step-notes-save-${e}`})}
            ${le({reviewerInputId:`step-reviewer-${e}`,reviewerValue:o,signoffInputId:`step-signoff-${e}`,signoffValue:s,isComplete:r,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`})}
        </section>
    `}function Hn(){var n,a,o,s,r,l,c;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",async()=>{var f;let i=Ae(Re);await Ne(i.sheetName,i.title,i.subtitle),Te({activeView:"home",activeStepId:null}),(f=document.getElementById("pf-hero"))==null||f.scrollIntoView({behavior:"smooth",block:"start"})}),(a=document.getElementById("nav-selector"))==null||a.addEventListener("click",()=>{window.location.href=Pn}),(o=document.getElementById("nav-prev"))==null||o.addEventListener("click",()=>qt(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>qt(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",i=>{i.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",i=>{!(t!=null&&t.contains(i.target))&&!(e!=null&&e.contains(i.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(r=document.getElementById("nav-roster"))==null||r.addEventListener("click",()=>{Qt("SS_Employee_Roster"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(l=document.getElementById("nav-accounts"))==null||l.addEventListener("click",()=>{Qt("SS_Chart_of_Accounts"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(c=document.getElementById("showConfigSheets"))==null||c.addEventListener("click",async()=>{await Zn()}),document.querySelectorAll("[data-step-card]").forEach(i=>{let f=Number(i.getAttribute("data-step-index")),g=Number(i.getAttribute("data-step-id"));i.addEventListener("click",()=>Me(f,g))}),j.activeView==="config"?Un():j.activeView==="step"&&j.activeStepId!=null&&Fn(j.activeStepId)}function Fn(e){var f,g,u,d,y,v,b,p,k,C,T,P,A,N,x,m,O;let t=e===2?document.getElementById("step-notes-input"):document.getElementById(`step-notes-${e}`),n=e===2?document.getElementById("step-reviewer-name"):document.getElementById(`step-reviewer-${e}`),a=e===2?document.getElementById("step-signoff-date"):document.getElementById(`step-signoff-${e}`),o=document.getElementById("step-back-btn"),s=e===2?document.getElementById("step-notes-lock-2"):document.getElementById(`step-notes-lock-${e}`),r=e===2?document.getElementById("step-notes-save-2"):document.getElementById(`step-notes-save-${e}`);r==null||r.addEventListener("click",async()=>{let h=(t==null?void 0:t.value)||"";await X(e,"notes",h),oe(r,!0)});let l=e===2?document.getElementById("headcount-signoff-save"):document.getElementById(`step-signoff-save-${e}`);l==null||l.addEventListener("click",async()=>{let h=(n==null?void 0:n.value)||"";await X(e,"reviewer",h),oe(l,!0)}),et();let c=e===2?"headcount-signoff-toggle":`step-signoff-toggle-${e}`,i=e===2?"step-signoff-date":`step-signoff-${e}`;fn(e,{buttonId:c,inputId:i,canActivate:e===2?()=>{var E;return!lt()||((E=document.getElementById("step-notes-input"))==null?void 0:E.value.trim())||""?!0:(window.alert("Please enter a brief explanation of the headcount differences before completing this step."),!1)}:null,onComplete:e===2?ba:null}),o==null||o.addEventListener("click",async()=>{let h=Ae(Re);await Ne(h.sheetName,h.title,h.subtitle),Te({activeView:"home",activeStepId:null})}),s==null||s.addEventListener("click",async()=>{let h=!s.classList.contains("is-locked");Ze(s,h),await dn(e,h)}),e===6&&((f=document.getElementById("archive-run-btn"))==null||f.addEventListener("click",()=>{})),e===1&&((g=document.getElementById("import-open-data-btn"))==null||g.addEventListener("click",()=>Le("PTO_Data")),(u=document.getElementById("import-clear-btn"))==null||u.addEventListener("click",()=>Xn())),e===2&&((d=document.getElementById("headcount-skip-btn"))==null||d.addEventListener("click",()=>{D.skipAnalysis=!D.skipAnalysis;let h=document.getElementById("headcount-skip-btn");h==null||h.classList.toggle("is-active",D.skipAnalysis),D.skipAnalysis&&on(),an()}),(y=document.getElementById("headcount-run-btn"))==null||y.addEventListener("click",()=>st()),(v=document.getElementById("headcount-refresh-btn"))==null||v.addEventListener("click",()=>st()),va(),D.skipAnalysis&&on(),an()),e===3&&((b=document.getElementById("quality-run-btn"))==null||b.addEventListener("click",()=>Wt()),(p=document.getElementById("quality-refresh-btn"))==null||p.addEventListener("click",()=>Wt()),(k=document.getElementById("quality-acknowledge-btn"))==null||k.addEventListener("click",()=>Yn())),e===4&&((C=document.getElementById("analysis-refresh-btn"))==null||C.addEventListener("click",()=>zt()),(T=document.getElementById("analysis-run-btn"))==null||T.addEventListener("click",()=>zt()),(P=document.getElementById("payrate-save-btn"))==null||P.addEventListener("click",Yt),(A=document.getElementById("payrate-ignore-btn"))==null||A.addEventListener("click",qn),(N=document.getElementById("payrate-input"))==null||N.addEventListener("keydown",h=>{h.key==="Enter"&&Yt()})),e===5&&((x=document.getElementById("je-create-btn"))==null||x.addEventListener("click",()=>Kn()),(m=document.getElementById("je-run-btn"))==null||m.addEventListener("click",()=>ln()),(O=document.getElementById("je-export-btn"))==null||O.addEventListener("click",()=>Qn()))}function Un(){var l,c,i,f,g;Tt("config-payroll-date",{onChange:u=>{if(ee(S.payrollDate,u),!!u){if(!_.overrides.accountingPeriod){let d=aa(u);if(d){let y=document.getElementById("config-accounting-period");y&&(y.value=d),ee(S.accountingPeriod,d)}}if(!_.overrides.journalId){let d=oa(u);if(d){let y=document.getElementById("config-journal-id");y&&(y.value=d),ee(S.journalEntryId,d)}}}}});let e=document.getElementById("config-accounting-period");e==null||e.addEventListener("change",u=>{_.overrides.accountingPeriod=!!u.target.value,ee(S.accountingPeriod,u.target.value||"")});let t=document.getElementById("config-journal-id");t==null||t.addEventListener("change",u=>{_.overrides.journalId=!!u.target.value,ee(S.journalEntryId,u.target.value.trim())}),(l=document.getElementById("config-company-name"))==null||l.addEventListener("change",u=>{ee(S.companyName,u.target.value.trim())}),(c=document.getElementById("config-payroll-provider"))==null||c.addEventListener("change",u=>{ee(S.payrollProvider,u.target.value.trim())}),(i=document.getElementById("config-accounting-link"))==null||i.addEventListener("change",u=>{ee(S.accountingSoftware,u.target.value.trim())}),(f=document.getElementById("config-user-name"))==null||f.addEventListener("change",u=>{ee(S.reviewerName,u.target.value.trim())});let n=document.getElementById("config-notes");n==null||n.addEventListener("input",u=>{X(0,"notes",u.target.value)});let a=document.getElementById("config-notes-lock");a==null||a.addEventListener("click",async()=>{let u=!a.classList.contains("is-locked");Ze(a,u),await dn(0,u)});let o=document.getElementById("config-notes-save");o==null||o.addEventListener("click",async()=>{n&&(await X(0,"notes",n.value),oe(o,!0))});let s=document.getElementById("config-reviewer");s==null||s.addEventListener("change",u=>{let d=u.target.value.trim();X(0,"reviewer",d);let y=document.getElementById("config-signoff-date");if(d&&y&&!y.value){let v=rt();y.value=v,X(0,"signOffDate",v),un(0,!0)}}),(g=document.getElementById("config-signoff-date"))==null||g.addEventListener("change",u=>{X(0,"signOffDate",u.target.value||"")});let r=document.getElementById("config-signoff-save");r==null||r.addEventListener("click",async()=>{var y,v;let u=((y=s==null?void 0:s.value)==null?void 0:y.trim())||"",d=((v=document.getElementById("config-signoff-date"))==null?void 0:v.value)||"";await X(0,"reviewer",u),await X(0,"signOffDate",d),oe(r,!0)}),et(),fn(0,{buttonId:"config-signoff-toggle",inputId:"config-signoff-date",onComplete:la})}function Me(e,t=null){if(e<0||e>=se.length)return;je=e;let n=t!=null?t:se[e].id;Te({focusedIndex:e,activeView:n===0?"config":"step",activeStepId:n}),n===1&&Le("PTO_Data"),n===2&&!D.hasAnalyzed&&(gn(),st()),n===3&&Le("PTO_Data"),n===5&&Le("PTO_JE_Draft")}function qt(e){let t=j.focusedIndex+e,n=Math.max(0,Math.min(se.length-1,t));Me(n,se[n].id)}function Gn(){if(je===null)return;let e=document.querySelector(`[data-step-index="${je}"]`);je=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}function Jn(e){return pn(_.completes[e])}function Te(e){e.stepStatuses&&(j.stepStatuses={...j.stepStatuses,...e.stepStatuses}),Object.assign(j,{...e,stepStatuses:j.stepStatuses}),ne()}function ae(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}async function Yt(){let e=document.getElementById("payrate-input");if(!e)return;let t=parseFloat(e.value),n=e.dataset.employee,a=parseInt(e.dataset.row,10);if(isNaN(t)||t<=0){window.alert("Please enter a valid pay rate greater than 0.");return}if(!n||isNaN(a)){console.error("Missing employee data on input");return}K(!0,"Updating pay rate...");try{await Excel.run(async o=>{let s=o.workbook.worksheets.getItem("PTO_Analysis"),r=s.getCell(a-1,3);r.values=[[t]];let l=s.getCell(a-1,8);l.load("values"),await o.sync();let i=(Number(l.values[0][0])||0)*t,f=s.getCell(a-1,9);f.values=[[i]];let g=s.getCell(a-1,10);g.load("values"),await o.sync();let u=Number(g.values[0][0])||0,d=i-u,y=s.getCell(a-1,11);y.values=[[d]],await o.sync()}),Y.missingPayRates=Y.missingPayRates.filter(o=>o.name!==n),K(!1),Me(3,3)}catch(o){console.error("Failed to save pay rate:",o),window.alert(`Failed to save pay rate: ${o.message}`),K(!1)}}function qn(){let e=document.getElementById("payrate-input");if(!e)return;let t=e.dataset.employee;t&&(Y.ignoredMissingPayRates.add(t),Y.missingPayRates=Y.missingPayRates.filter(n=>n.name!==t)),Me(3,3)}async function Wt(){if(!ae()){window.alert("Excel is not available. Open this module inside Excel to run quality check.");return}J.loading=!0,K(!0,"Analyzing data quality..."),oe(document.getElementById("quality-save-btn"),!1);try{await Excel.run(async t=>{var b;let a=t.workbook.worksheets.getItem("PTO_Data").getUsedRangeOrNullObject();a.load("values"),await t.sync();let o=a.isNullObject?[]:a.values||[];if(!o.length||o.length<2)throw new Error("PTO_Data is empty or has no data rows.");let s=(o[0]||[]).map(p=>G(p));console.log("[Data Quality] PTO_Data headers:",o[0]);let r=s.findIndex(p=>p==="employee name"||p==="employeename");r===-1&&(r=s.findIndex(p=>p.includes("employee")&&p.includes("name"))),r===-1&&(r=s.findIndex(p=>p==="name"||p.includes("name")&&!p.includes("company")&&!p.includes("form"))),console.log("[Data Quality] Employee name column index:",r,"Header:",(b=o[0])==null?void 0:b[r]);let l=$(s,["balance"]),c=$(s,["accrual rate","accrualrate"]),i=$(s,["carry over","carryover"]),f=$(s,["ytd accrued","ytdaccrued"]),g=$(s,["ytd used","ytdused"]),u=[],d=[],y=[],v=o.slice(1);v.forEach((p,k)=>{let C=k+2,T=r!==-1?String(p[r]||"").trim():`Row ${C}`;if(!T)return;let P=l!==-1&&Number(p[l])||0,A=c!==-1&&Number(p[c])||0,N=i!==-1&&Number(p[i])||0,x=f!==-1&&Number(p[f])||0,m=g!==-1&&Number(p[g])||0,O=N+x;P<0?u.push({name:T,issue:`Negative balance: ${P.toFixed(2)} hrs`,rowIndex:C}):m>O&&O>0&&u.push({name:T,issue:`Used ${m.toFixed(0)} hrs but only ${O.toFixed(0)} available`,rowIndex:C}),P===0&&(N>0||x>0)&&d.push({name:T,rowIndex:C}),A>8&&y.push({name:T,accrualRate:A,rowIndex:C})}),J.balanceIssues=u,J.zeroBalances=d,J.accrualOutliers=y,J.totalIssues=u.length,J.totalEmployees=v.filter(p=>p.some(k=>k!==null&&k!=="")).length,J.hasRun=!0});let e=J.balanceIssues.length>0;Te({stepStatuses:{3:e?"blocked":"complete"}})}catch(e){console.error("Data quality check error:",e),window.alert(`Quality check failed: ${e.message}`),J.hasRun=!1}finally{J.loading=!1,K(!1),ne()}}function Yn(){J.acknowledged=!0,Te({stepStatuses:{3:"complete"}}),ne()}async function Wn(){if(ae())try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem("PTO_Data"),n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=t.getUsedRangeOrNullObject();if(a.load("values"),n.load("isNullObject"),await e.sync(),n.isNullObject){Y.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let o=n.getUsedRangeOrNullObject();o.load("values"),await e.sync();let s=a.isNullObject?[]:a.values||[],r=o.isNullObject?[]:o.values||[];if(!s.length||!r.length){Y.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let l=(f,g,u)=>{let d=(f[0]||[]).map(b=>G(b)),y=$(d,g);return y===-1?null:f.slice(1).reduce((b,p)=>b+(Number(p[y])||0),0)},c=[{key:"accrualRate",aliases:["accrual rate","accrualrate"]},{key:"carryOver",aliases:["carry over","carryover","carry_over"]},{key:"ytdAccrued",aliases:["ytd accrued","ytdaccrued","ytd_accrued"]},{key:"ytdUsed",aliases:["ytd used","ytdused","ytd_used"]},{key:"balance",aliases:["balance"]}],i={};for(let f of c){let g=l(s,f.aliases,"PTO_Data"),u=l(r,f.aliases,"PTO_Analysis");if(g===null||u===null)i[f.key]=null;else{let d=Math.abs(g-u)<.01;i[f.key]={match:d,ptoData:g,ptoAnalysis:u}}}Y.completenessCheck=i})}catch(e){console.error("Completeness check failed:",e)}}async function zt(){if(!ae()){window.alert("Excel is not available. Open this module inside Excel to run analysis.");return}K(!0,"Running analysis...");try{await gn(),await Wn(),Y.cleanDataReady=!0,ne()}catch(e){console.error("Full analysis error:",e),window.alert(`Analysis failed: ${e.message}`)}finally{K(!1)}}async function ln(){if(!ae()){window.alert("Excel is not available. Open this module inside Excel to run journal checks.");return}U.loading=!0,U.lastError=null,oe(document.getElementById("je-save-btn"),!1),ne();try{let e=await Excel.run(async t=>{let a=t.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();a.load("values");let o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis");o.load("isNullObject"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty. Generate the JE first.");let r=(s[0]||[]).map(C=>G(C)),l=$(r,["debit"]),c=$(r,["credit"]),i=$(r,["lineamount","line amount"]),f=$(r,["account number","accountnumber"]);if(l===-1||c===-1)throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");let g=0,u=0,d=0,y=0;s.slice(1).forEach(C=>{let T=Number(C[l])||0,P=Number(C[c])||0,A=i!==-1&&Number(C[i])||0,N=f!==-1?String(C[f]||"").trim():"";g+=T,u+=P,d+=A,N&&N!=="21540"&&(y+=A)});let v=0;if(!o.isNullObject){let C=o.getUsedRangeOrNullObject();C.load("values"),await t.sync();let T=C.isNullObject?[]:C.values||[];if(T.length>1){let P=(T[0]||[]).map(N=>G(N)),A=$(P,["change"]);A!==-1&&T.slice(1).forEach(N=>{v+=Number(N[A])||0})}}let b=g-u,p=[];Math.abs(b)>=.01?p.push({check:"Debits = Credits",passed:!1,detail:b>0?`Debits exceed credits by $${Math.abs(b).toLocaleString(void 0,{minimumFractionDigits:2})}`:`Credits exceed debits by $${Math.abs(b).toLocaleString(void 0,{minimumFractionDigits:2})}`}):p.push({check:"Debits = Credits",passed:!0,detail:""}),Math.abs(d)>=.01?p.push({check:"Line Amounts Sum to Zero",passed:!1,detail:`Line amounts sum to $${d.toLocaleString(void 0,{minimumFractionDigits:2})} (should be $0.00)`}):p.push({check:"Line Amounts Sum to Zero",passed:!0,detail:""});let k=Math.abs(y-v);return k>=.01?p.push({check:"JE Matches Analysis Total",passed:!1,detail:`JE expense total ($${y.toLocaleString(void 0,{minimumFractionDigits:2})}) differs from PTO_Analysis Change total ($${v.toLocaleString(void 0,{minimumFractionDigits:2})}) by $${k.toLocaleString(void 0,{minimumFractionDigits:2})}`}):p.push({check:"JE Matches Analysis Total",passed:!0,detail:""}),{debitTotal:g,creditTotal:u,difference:b,lineAmountSum:d,jeChangeTotal:y,analysisChangeTotal:v,issues:p,validationRun:!0}});Object.assign(U,e,{lastError:null})}catch(e){console.warn("PTO JE summary:",e),U.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",U.debitTotal=null,U.creditTotal=null,U.difference=null,U.lineAmountSum=null,U.jeChangeTotal=null,U.analysisChangeTotal=null,U.issues=[],U.validationRun=!1}finally{U.loading=!1,ne()}}var zn={"general & administrative":"64110","general and administrative":"64110","g&a":"64110","research & development":"62110","research and development":"62110","r&d":"62110",marketing:"61610","cogs onboarding":"53110","cogs prof. services":"56110","cogs professional services":"56110","sales & marketing":"61110","sales and marketing":"61110","cogs support":"52110","client success":"61811"},Kt="21540";async function Kn(){if(!ae()){window.alert("Excel is not available. Open this module inside Excel to create the journal entry.");return}K(!0,"Creating PTO Journal Entry...");try{await Excel.run(async e=>{let t=[],n=e.workbook.tables.getItemOrNullObject(de[0]);if(n.load("isNullObject"),await e.sync(),n.isNullObject){let m=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");if(m.load("isNullObject"),await e.sync(),!m.isNullObject){let O=m.getUsedRangeOrNullObject();O.load("values"),await e.sync();let h=O.isNullObject?[]:O.values||[];t=h.length>1?h.slice(1):[]}}else{let m=n.getDataBodyRange();m.load("values"),await e.sync(),t=m.values||[]}let a=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis");if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PTO_Analysis sheet not found. Please ensure the worksheet exists.");let o=a.getUsedRangeOrNullObject();o.load("values");let s=e.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");s.load("isNullObject"),await e.sync();let r=[];if(!s.isNullObject){let m=s.getUsedRangeOrNullObject();m.load("values"),await e.sync(),r=m.isNullObject?[]:m.values||[]}let l=o.isNullObject?[]:o.values||[];if(!l.length||l.length<2)throw new Error("PTO_Analysis is empty or has no data rows. Run the analysis first (Step 4).");let c={};t.forEach(m=>{let O=String(m[1]||"").trim(),h=m[2];O&&(c[O]=h)}),(!c[S.journalEntryId]||!c[S.payrollDate])&&console.warn("[JE Draft] Missing config values - RefNumber:",c[S.journalEntryId],"TxnDate:",c[S.payrollDate]);let i=c[S.journalEntryId]||"",f=c[S.payrollDate]||"",g=c[S.accountingPeriod]||"",u="";if(f)try{let m;if(typeof f=="number"||/^\d{4,5}$/.test(String(f).trim())){let O=Number(f),h=new Date(1899,11,30);m=new Date(h.getTime()+O*24*60*60*1e3)}else m=new Date(f);if(!isNaN(m.getTime())&&m.getFullYear()>1970){let O=String(m.getMonth()+1).padStart(2,"0"),h=String(m.getDate()).padStart(2,"0"),E=m.getFullYear();u=`${O}/${h}/${E}`}else console.warn("[JE Draft] Date parsing resulted in invalid date:",f,"->",m),u=String(f)}catch(m){console.warn("[JE Draft] Could not parse TxnDate:",f,m),u=String(f)}let d=g?`${g} PTO Accrual`:"PTO Accrual",y={};if(r.length>1){let m=(r[0]||[]).map(E=>G(E)),O=$(m,["account number","accountnumber","account","acct"]),h=$(m,["account name","accountname","name","description"]);O!==-1&&h!==-1&&r.slice(1).forEach(E=>{let V=String(E[O]||"").trim(),W=String(E[h]||"").trim();V&&(y[V]=W)})}let v=(l[0]||[]).map(m=>G(m));console.log("[JE Draft] PTO_Analysis headers:",v),console.log("[JE Draft] PTO_Analysis row count:",l.length-1);let b=$(v,["department"]),p=$(v,["change"]);if(console.log("[JE Draft] Column indices - Department:",b,"Change:",p),b===-1||p===-1)throw new Error(`Could not find required columns in PTO_Analysis. Found headers: ${v.join(", ")}. Looking for "Department" (found: ${b!==-1}) and "Change" (found: ${p!==-1}).`);let k={};l.slice(1).forEach(m=>{let O=String(m[b]||"").trim(),h=Number(m[p])||0;O&&h!==0&&(k[O]||(k[O]=0),k[O]+=h)});let C=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],T=[C],P=0,A=0;Object.entries(k).forEach(([m,O])=>{if(Math.abs(O)<.01)return;let h=m.toLowerCase().trim(),E=zn[h]||"",V=y[E]||"",W=O>0?Math.abs(O):0,Q=O<0?Math.abs(O):0;P+=W,A+=Q,T.push([i,u,E,V,O,W,Q,d,m])});let N=P-A;if(Math.abs(N)>=.01){let m=N<0?Math.abs(N):0,O=N>0?Math.abs(N):0,h=y[Kt]||"Accrued PTO";T.push([i,u,Kt,h,-N,m,O,d,""])}let x=e.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");if(x.load("isNullObject"),await e.sync(),x.isNullObject)x=e.workbook.worksheets.add("PTO_JE_Draft");else{let m=x.getUsedRangeOrNullObject();m.load("isNullObject"),await e.sync(),m.isNullObject||m.clear()}if(T.length>0){let m=x.getRangeByIndexes(0,0,T.length,C.length);m.values=T;let O=x.getRangeByIndexes(0,0,1,C.length);nt(O);let h=T.length-1;h>0&&(ce(x,4,h,!0),ce(x,5,h),ce(x,6,h)),m.format.autofitColumns()}await e.sync(),x.activate(),x.getRange("A1").select(),await e.sync()}),await ln()}catch(e){console.error("Create JE Draft error:",e),window.alert(`Unable to create Journal Entry: ${e.message}`)}finally{K(!1)}}async function Qn(){if(!ae()){window.alert("Excel is not available. Open this module inside Excel to export.");return}K(!0,"Preparing JE CSV...");try{let{rows:e}=await Excel.run(async n=>{let o=n.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();o.load("values"),await n.sync();let s=o.isNullObject?[]:o.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty.");return{rows:s}}),t=ya(e);ha(`pto-je-draft-${rt()}.csv`,t)}catch(e){console.error("PTO JE export:",e),window.alert("Unable to export the JE draft. Confirm the sheet has data.")}finally{K(!1)}}async function Le(e){if(!(!e||!ae()))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e);n.activate(),n.getRange("A1").select(),await t.sync()})}catch(t){console.error(t)}}async function Xn(){if(!(!ae()||!window.confirm("This will clear all data in PTO_Data. Are you sure?"))){K(!0);try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("PTO_Data"),a=n.getUsedRangeOrNullObject();a.load("rowCount"),await t.sync(),!a.isNullObject&&a.rowCount>1&&(n.getRangeByIndexes(1,0,a.rowCount-1,20).clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),window.alert("PTO_Data cleared successfully. You can now paste new data.")}catch(t){console.error("Clear PTO_Data error:",t),window.alert(`Failed to clear PTO_Data: ${t.message}`)}finally{K(!1)}}}async function Qt(e){if(!e||!ae())return;let t={SS_Employee_Roster:["Employee","Department","Pay_Rate","Status","Hire_Date"],SS_Chart_of_Accounts:["Account_Number","Account_Name","Type","Category"]};try{await Excel.run(async n=>{let a=n.workbook.worksheets.getItemOrNullObject(e);if(a.load("isNullObject"),await n.sync(),a.isNullObject){a=n.workbook.worksheets.add(e);let o=t[e]||["Column1","Column2"],s=a.getRange(`A1:${String.fromCharCode(64+o.length)}1`);s.values=[o],s.format.font.bold=!0,s.format.fill.color="#f0f0f0",s.format.autofitColumns(),await n.sync()}a.activate(),a.getRange("A1").select(),await n.sync()})}catch(n){console.error("Error opening reference sheet:",n)}}async function Zn(){if(!ae()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(o=>{o.name.toUpperCase().startsWith("SS_")&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Config] Made visible: ${o.name}`),n++)}),await e.sync();let a=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");a.load("isNullObject"),await e.sync(),a.isNullObject||(a.activate(),a.getRange("A1").select(),await e.sync()),console.log(`[Config] ${n} system sheets now visible`)})}catch(e){console.error("[Config] Error unhiding system sheets:",e)}}function te(e){var n,a;let t=String(e!=null?e:"").trim();return(a=(n=_.values)==null?void 0:n[t])!=null?a:""}function me(e){var n;if(e)return e;let t=te(S.reviewerName);if(t)return t;if((n=window.PrairieForge)!=null&&n._sharedConfigCache){let a=window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer")||window.PrairieForge._sharedConfigCache.get("Default_Reviewer");if(a)return a}return""}function ee(e,t,n={}){var r;let a=String(e!=null?e:"").trim();if(!a)return;_.values[a]=t!=null?t:"";let o=(r=n.debounceMs)!=null?r:0;if(!o){let l=we.get(a);l&&clearTimeout(l),we.delete(a),Ee(a,t!=null?t:"",de);return}we.has(a)&&clearTimeout(we.get(a));let s=setTimeout(()=>{we.delete(a),Ee(a,t!=null?t:"",de)},o);we.set(a,s)}function G(e){return String(e!=null?e:"").trim().toLowerCase()}function K(e,t="Working..."){let n=document.getElementById(In);n&&(n.style.display="none")}function ot(){Tn()}typeof Office!="undefined"&&Office.onReady?Office.onReady(()=>ot()).catch(()=>ot()):ot();function ye(e){return _.steps[e]||{notes:"",reviewer:"",signOffDate:""}}function cn(e){return Be[e]||{}}function ea(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}async function X(e,t,n){let a=_.steps[e]||{notes:"",reviewer:"",signOffDate:""};a[t]=n,_.steps[e]=a;let o=cn(e),s=t==="notes"?o.note:t==="reviewer"?o.reviewer:o.signOff;if(s&&q())try{await Ee(s,n,de)}catch(r){console.warn("PTO: unable to save field",s,r)}}async function dn(e,t){_.permanents[e]=t;let n=cn(e);if(n!=null&&n.note&&q())try{await Excel.run(async a=>{var u;let o=a.workbook.tables.getItemOrNullObject(de[0]);if(await a.sync(),o.isNullObject)return;let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let l=r.values[0]||[],c=l.map(d=>String(d||"").trim().toLowerCase()),i={field:c.findIndex(d=>d==="field"||d==="field name"||d==="setting"),permanent:c.findIndex(d=>d==="permanent"||d==="persist"),value:c.findIndex(d=>d==="value"||d==="setting value"),type:c.findIndex(d=>d==="type"||d==="category"),title:c.findIndex(d=>d==="title"||d==="display name")};if(i.field===-1)return;let g=(s.values||[]).findIndex(d=>String(d[i.field]||"").trim()===n.note);if(g>=0)i.permanent>=0&&(s.getCell(g,i.permanent).values=[[t?"Y":"N"]]);else{let d=new Array(l.length).fill("");i.type>=0&&(d[i.type]="Other"),i.title>=0&&(d[i.title]=""),d[i.field]=n.note,i.permanent>=0&&(d[i.permanent]=t?"Y":"N"),i.value>=0&&(d[i.value]=((u=_.steps[e])==null?void 0:u.notes)||""),o.rows.add(null,[d])}await a.sync()})}catch(a){console.warn("PTO: unable to update permanent flag",a)}}async function un(e,t){let n=sn[e];if(n&&(_.completes[e]=t?"Y":"",!!q()))try{await Ee(n,t?"Y":"",de)}catch(a){console.warn("PTO: unable to save completion flag",n,a)}}function Xt(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function ta(){let e={};return Object.keys(Be).forEach(t=>{var s;let n=parseInt(t,10),a=!!((s=_.steps[n])!=null&&s.signOffDate),o=!!_.completes[n];e[n]=a||o}),e}function fn(e,{buttonId:t,inputId:n,canActivate:a=null,onComplete:o=null}){var c;let s=document.getElementById(t);if(!s)return;let r=document.getElementById(n),l=!!((c=_.steps[e])!=null&&c.signOffDate)||!!_.completes[e];Xt(s,l),s.addEventListener("click",()=>{if(!s.classList.contains("is-active")&&e>0){let g=ta(),{canComplete:u,message:d}=Ut(e,g);if(!u){Gt(d);return}}if(typeof a=="function"&&!a())return;let f=!s.classList.contains("is-active");Xt(s,f),r&&(r.value=f?rt():"",X(e,"signOffDate",r.value)),un(e,f),f&&typeof o=="function"&&o()})}function w(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function na(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function pn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function Oe(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function it(e){if(!e)return null;let t=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(e));if(!t)return null;let n=Number(t[1]),a=Number(t[2]),o=Number(t[3]);return!n||!a||!o?null:{year:n,month:a,day:o}}function Zt(e){if(!e)return"";let t=it(e);if(!t)return"";let{year:n,month:a,day:o}=t;return`${n}-${String(a).padStart(2,"0")}-${String(o).padStart(2,"0")}`}function aa(e){let t=it(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function oa(e){let t=it(e);return t?`PTO-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function rt(){let e=new Date,t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}function sa(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="y"||t==="yes"||t==="true"||t==="t"||t==="1"}function ia(e){if(e instanceof Date)return e.getTime();if(typeof e=="number"){let n=ra(e);return n?n.getTime():null}let t=new Date(e);return Number.isNaN(t.getTime())?null:t.getTime()}function ra(e){if(!Number.isFinite(e))return null;let t=new Date(Date.UTC(1899,11,30));return new Date(t.getTime()+e*24*60*60*1e3)}function la(){let e=n=>{var a,o;return((o=(a=document.getElementById(n))==null?void 0:a.value)==null?void 0:o.trim())||""};[{id:"config-payroll-date",field:S.payrollDate},{id:"config-accounting-period",field:S.accountingPeriod},{id:"config-journal-id",field:S.journalEntryId},{id:"config-company-name",field:S.companyName},{id:"config-payroll-provider",field:S.payrollProvider},{id:"config-accounting-link",field:S.accountingSoftware},{id:"config-user-name",field:S.reviewerName}].forEach(({id:n,field:a})=>{let o=e(n);a&&ee(a,o)})}function $(e,t=[]){let n=t.map(a=>G(a));return e.findIndex(a=>n.some(o=>a.includes(o)))}function ca(e){var C,T,P,A,N,x,m,O,h;let t=ye(2),n=(t==null?void 0:t.notes)||"",a=!!_.permanents[2],o=me(t==null?void 0:t.reviewer),s=(t==null?void 0:t.signOffDate)||"",r=!!(Oe(_.completes[2])||s),l=D.roster||{},c=D.hasAnalyzed,i=(T=(C=D.roster)==null?void 0:C.difference)!=null?T:0,f=!D.skipAnalysis&&Math.abs(i)>0,g=(P=l.rosterCount)!=null?P:0,u=(A=l.payrollCount)!=null?A:0,d=(N=l.difference)!=null?N:u-g,y=Array.isArray(l.mismatches)?l.mismatches.filter(Boolean):[],v="";D.loading?v=((m=(x=window.PrairieForge)==null?void 0:x.renderStatusBanner)==null?void 0:m.call(x,{type:"info",message:"Analyzing headcount\u2026",escapeHtml:w}))||"":D.lastError&&(v=((h=(O=window.PrairieForge)==null?void 0:O.renderStatusBanner)==null?void 0:h.call(O,{type:"error",message:D.lastError,escapeHtml:w}))||"");let b=(E,V,W,Q)=>{let z=!c,fe;z?fe='<span class="pf-je-check-circle pf-je-circle--pending"></span>':Q?fe=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:fe=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let Ve=c?` = ${W}`:"";return`
            <div class="pf-je-check-row">
                ${fe}
                <span class="pf-je-check-desc-pill">${w(E)}${Ve}</span>
            </div>
        `},p=`
        ${b("SS_Employee_Roster count","Active employees in roster",g,!0)}
        ${b("PTO_Data count","Unique employees in PTO data",u,!0)}
        ${b("Difference","Should be zero",d,d===0)}
    `,k=y.length&&!D.skipAnalysis&&c?window.PrairieForge.renderMismatchTiles({mismatches:y,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"PTO Data",escapeHtml:w}):"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${D.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${Bt}
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
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${De}</button>`,"Run")}
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${Ce}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${v}
                <div class="pf-je-checks-container">
                    ${p}
                </div>
                ${k}
            </article>
            ${re({textareaId:"step-notes-input",value:n,permanentId:"step-notes-lock-2",isPermanent:a,hintId:f?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
            ${le({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:r,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
        </section>
    `}function da(){let e=Y.completenessCheck||{},t=Y.missingPayRates||[],n=[{key:"accrualRate",label:"Accrual Rate",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"carryOver",label:"Carry Over",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdAccrued",label:"YTD Accrued",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdUsed",label:"YTD Used",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"balance",label:"Balance",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"}],o=n.every(i=>e[i.key]!==null&&e[i.key]!==void 0)&&n.every(i=>{var f;return(f=e[i.key])==null?void 0:f.match}),s=t.length>0,r=i=>{let f=e[i.key],g=f==null,u;return g?u='<span class="pf-je-check-circle pf-je-circle--pending"></span>':f.match?u=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:u=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${u}
                <span class="pf-je-check-desc-pill">${w(i.label)}: ${w(i.desc)}</span>
            </div>
        `},l=n.map(i=>r(i)).join(""),c="";if(s){let i=t[0],f=t.length-1;c=`
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${w(i.name)}</strong> to calculate liability
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
                               data-employee="${na(i.name)}"
                               data-row="${i.rowIndex}">
                    </div>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--secondary" id="payrate-ignore-btn">
                        Skip
                    </button>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--primary" id="payrate-save-btn">
                        Save
                    </button>
                </div>
                ${f>0?`<p class="pf-readiness-remaining">${f} more employee${f>1?"s":""} need pay rates</p>`:""}
            </div>
        `}return`
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-readiness-card">
            <div class="pf-config-head">
                <h3>Data Completeness</h3>
                <p class="pf-config-subtext">Quick check that all your data transferred correctly.</p>
            </div>
            <div class="pf-je-checks-container">
                ${l}
            </div>
            ${c}
        </article>
    `}function ua(e){var d,y,v,b,p,k,C,T;let t=ye(3),n=!!_.permanents[3],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Oe(_.completes[3])||o),r=J.hasRun,{balanceIssues:l,zeroBalances:c,accrualOutliers:i,totalEmployees:f}=J,g="";if(J.loading)g=((y=(d=window.PrairieForge)==null?void 0:d.renderStatusBanner)==null?void 0:y.call(d,{type:"info",message:"Analyzing data quality...",escapeHtml:w}))||"";else if(r){let P=l.length,A=i.length+c.length;P>0?g=((b=(v=window.PrairieForge)==null?void 0:v.renderStatusBanner)==null?void 0:b.call(v,{type:"error",title:`${P} Balance Issue${P>1?"s":""} Found`,message:"Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",escapeHtml:w}))||"":A>0?g=((k=(p=window.PrairieForge)==null?void 0:p.renderStatusBanner)==null?void 0:k.call(p,{type:"warning",title:"No Critical Issues",message:`${A} informational item${A>1?"s":""} to review (see below).`,escapeHtml:w}))||"":g=((T=(C=window.PrairieForge)==null?void 0:C.renderStatusBanner)==null?void 0:T.call(C,{type:"success",title:"Data Quality Passed",message:`${f} employee${f!==1?"s":""} checked \u2014 no anomalies found.`,escapeHtml:w}))||""}let u=[];return r&&l.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u26A0\uFE0F</span>
                    <span class="pf-quality-issue-title">Balance Issues (${l.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${l.slice(0,5).map(P=>`<li><strong>${w(P.name)}</strong>: ${w(P.issue)}</li>`).join("")}
                    ${l.length>5?`<li class="pf-quality-more">+${l.length-5} more</li>`:""}
                </ul>
            </div>
        `),r&&i.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--warning">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u{1F4CA}</span>
                    <span class="pf-quality-issue-title">High Accrual Rates (${i.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${i.slice(0,5).map(P=>`<li><strong>${w(P.name)}</strong>: ${P.accrualRate.toFixed(2)} hrs/period</li>`).join("")}
                    ${i.length>5?`<li class="pf-quality-more">+${i.length-5} more</li>`:""}
                </ul>
            </div>
        `),r&&c.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--info">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u2139\uFE0F</span>
                    <span class="pf-quality-issue-title">Zero Balances (${c.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${c.slice(0,5).map(P=>`<li><strong>${w(P.name)}</strong></li>`).join("")}
                    ${c.length>5?`<li class="pf-quality-more">+${c.length-5} more</li>`:""}
                </ul>
            </div>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${g}
                <div class="pf-signoff-action">
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${De}</button>`,"Run")}
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
                        ${J.acknowledged?'<p class="pf-quality-actions-hint"><span class="pf-acknowledged-badge">\u2713 Issues Acknowledged</span></p>':""}
                        <div class="pf-signoff-action">
                            ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${Ce}</button>`,"Refresh")}
                            ${J.acknowledged?"":F(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${_e}</button>`,"Continue")}
                        </div>
                    </div>
                </article>
            `:""}
            ${re({textareaId:"step-notes-3",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-3",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-3"})}
            ${le({reviewerInputId:"step-reviewer-3",reviewerValue:a,signoffInputId:"step-signoff-3",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-3",completeButtonId:"step-signoff-toggle-3"})}
        </section>
    `}function fa(e){let t=ye(4),n=!!_.permanents[4],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Oe(_.completes[4])||o);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Analysis</h3>
                    <p class="pf-config-subtext">Calculate liabilities and compare against last period.</p>
                </div>
                <div class="pf-signoff-action">
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${De}</button>`,"Run")}
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${Ce}</button>`,"Refresh")}
                </div>
            </article>
            ${da()}
            ${re({textareaId:"step-notes-4",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-4",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-4"})}
            ${le({reviewerInputId:"step-reviewer-4",reviewerValue:a,signoffInputId:"step-signoff-4",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"step-signoff-toggle-4"})}
        </section>
    `}function pa(e){let t=ye(5),n=!!_.permanents[5],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Oe(_.completes[5])||o),r=U.lastError?`<p class="pf-step-note">${w(U.lastError)}</p>`:"",l=U.validationRun,c=U.issues||[],i=[{key:"Debits = Credits",desc:"\u2211 Debit column = \u2211 Credit column"},{key:"Line Amounts Sum to Zero",desc:"\u2211 Line Amount = $0.00"},{key:"JE Matches Analysis Total",desc:"\u2211 Expense line amounts = \u2211 PTO_Analysis Change"}],f=y=>{let v=c.find(k=>k.check===y.key),b=!l,p;return b?p='<span class="pf-je-check-circle pf-je-circle--pending"></span>':v!=null&&v.passed?p=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:p=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${p}
                <span class="pf-je-check-desc-pill">${w(y.desc)}</span>
            </div>
        `},g=i.map(y=>f(y)).join(""),u=c.filter(y=>!y.passed),d="";return l&&u.length>0&&(d=`
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>\u26A0\uFE0F Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${u.map(y=>`<li><strong>${w(y.check)}:</strong> ${w(y.detail)}</li>`).join("")}
                </ul>
            </article>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${ze}</button>`,"Generate")}
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${Ce}</button>`,"Refresh")}
                    ${F(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${Lt}</button>`,"Export")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${r}
                <div class="pf-je-checks-container">
                    ${g}
                </div>
            </article>
            ${d}
            ${re({textareaId:"step-notes-5",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-5",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-5"})}
            ${le({reviewerInputId:"step-reviewer-5",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
        </section>
    `}function ga(){var t,n;return Math.abs((n=(t=D.roster)==null?void 0:t.difference)!=null?n:0)>0}function lt(){return!D.skipAnalysis&&ga()}async function st(){if(!q()){D.loading=!1,D.lastError="Excel runtime is unavailable.",ne();return}D.loading=!0,D.lastError=null,oe(document.getElementById("headcount-save-btn"),!1),ne();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),a=t.workbook.worksheets.getItem("PTO_Data"),o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.getUsedRangeOrNullObject(),r=a.getUsedRangeOrNullObject();s.load("values"),r.load("values"),o.load("isNullObject"),await t.sync();let l=null;o.isNullObject||(l=o.getUsedRangeOrNullObject(),l.load("values")),await t.sync();let c=s.isNullObject?[]:s.values||[],i=r.isNullObject?[]:r.values||[],f=l&&!l.isNullObject?l.values||[]:[],g=f.length?f:i;return ma(c,g)});D.roster=e.roster,D.hasAnalyzed=!0,D.lastError=null}catch(e){console.warn("PTO headcount: unable to analyze data",e),D.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{D.loading=!1,ne()}}function en(e){if(!e)return!0;let t=e.toLowerCase().trim();return t?["total","subtotal","sum","count","grand","average","avg"].some(a=>t.includes(a)):!0}function ma(e,t){let n={rosterCount:0,payrollCount:0,difference:0,mismatches:[]};if(((e==null?void 0:e.length)||0)<2||((t==null?void 0:t.length)||0)<2)return console.warn("Headcount: insufficient data rows",{rosterRows:(e==null?void 0:e.length)||0,payrollRows:(t==null?void 0:t.length)||0}),{roster:n};let a=tn(e),o=tn(t),s=a.headers,r=o.headers,l={employee:nn(s),termination:s.findIndex(d=>d.includes("termination"))},c={employee:nn(r)};console.log("Headcount column detection:",{rosterEmployeeCol:l.employee,rosterTerminationCol:l.termination,payrollEmployeeCol:c.employee,rosterHeaders:s.slice(0,5),payrollHeaders:r.slice(0,5)});let i=new Set,f=new Set;for(let d=a.startIndex;d<e.length;d+=1){let y=e[d],v=l.employee>=0?ge(y[l.employee]):"";en(v)||l.termination>=0&&ge(y[l.termination])||i.add(v.toLowerCase())}for(let d=o.startIndex;d<t.length;d+=1){let y=t[d],v=c.employee>=0?ge(y[c.employee]):"";en(v)||f.add(v.toLowerCase())}n.rosterCount=i.size,n.payrollCount=f.size,n.difference=n.payrollCount-n.rosterCount,console.log("Headcount results:",{rosterCount:n.rosterCount,payrollCount:n.payrollCount,difference:n.difference});let g=[...i].filter(d=>!f.has(d)),u=[...f].filter(d=>!i.has(d));return n.mismatches=[...g.map(d=>`In roster, missing in PTO_Data: ${d}`),...u.map(d=>`In PTO_Data, missing in roster: ${d}`)],{roster:n}}function tn(e){if(!Array.isArray(e)||!e.length)return{headers:[],startIndex:1};let t=e.findIndex((o=[])=>o.some(s=>ge(s).toLowerCase().includes("employee"))),n=t===-1?0:t;return{headers:(e[n]||[]).map(o=>ge(o).toLowerCase()),startIndex:n+1}}function nn(e=[]){let t=-1,n=-1;return e.forEach((a,o)=>{let s=a.toLowerCase();if(!s.includes("employee"))return;let r=1;s.includes("name")?r=4:s.includes("id")?r=2:r=3,r>n&&(n=r,t=o)}),t}function ge(e){return e==null?"":String(e).trim()}async function gn(e=null){let t=async n=>{let a=n.workbook.worksheets.getItem("PTO_Data"),o=n.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster"),r=n.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary"),l=n.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary"),c=a.getUsedRangeOrNullObject();c.load("values"),o.load("isNullObject"),s.load("isNullObject"),r.load("isNullObject"),l.load("isNullObject"),await n.sync();let i=c.isNullObject?[]:c.values||[];if(!i.length)return;let f=(i[0]||[]).map(I=>G(I)),g=f.findIndex(I=>I.includes("employee")&&I.includes("name")),u=g>=0?g:0,d=$(f,["accrual rate"]),y=$(f,["carry over","carryover"]),v=f.findIndex(I=>I.includes("ytd")&&(I.includes("accrued")||I.includes("accrual"))),b=f.findIndex(I=>I.includes("ytd")&&I.includes("used")),p=$(f,["balance","current balance","pto balance"]);console.log("[PTO Analysis] PTO_Data headers:",f),console.log("[PTO Analysis] Column indices found:",{employee:u,accrualRate:d,carryOver:y,ytdAccrued:v,ytdUsed:b,balance:p}),b>=0?console.log(`[PTO Analysis] YTD Used column: "${f[b]}" at index ${b}`):console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:",f);let k=i.slice(1).map(I=>ge(I[u])).filter(I=>I&&!I.toLowerCase().includes("total")),C=new Map;i.slice(1).forEach(I=>{let H=G(I[u]);!H||H.includes("total")||C.set(H,I)});let T=new Map;if(s.isNullObject)console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");else{let I=s.getUsedRangeOrNullObject();I.load("values"),await n.sync();let H=I.isNullObject?[]:I.values||[];if(H.length){let L=(H[0]||[]).map(R=>G(R));console.log("[PTO Analysis] SS_Employee_Roster headers:",L);let B=L.findIndex(R=>R.includes("employee")&&R.includes("name"));B<0&&(B=L.findIndex(R=>R==="employee"||R==="name"||R==="full name"));let M=L.findIndex(R=>R.includes("department"));console.log(`[PTO Analysis] Roster column indices - Name: ${B}, Dept: ${M}`),B>=0&&M>=0?(H.slice(1).forEach(R=>{let Z=G(R[B]),ie=ge(R[M]);Z&&T.set(Z,ie)}),console.log(`[PTO Analysis] Built roster map with ${T.size} employees`)):console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster")}}let P=new Map;if(!r.isNullObject){let I=r.getUsedRangeOrNullObject();I.load("values"),await n.sync();let H=I.isNullObject?[]:I.values||[];if(H.length){let L=(H[0]||[]).map(M=>G(M)),B={payrollDate:$(L,["payroll date"]),employee:$(L,["employee"]),category:$(L,["payroll category","category"]),amount:$(L,["amount","gross salary","gross_salary","earnings"])};B.employee>=0&&B.category>=0&&B.amount>=0&&H.slice(1).forEach(M=>{let R=G(M[B.employee]);if(!R)return;let Z=G(M[B.category]);if(!Z.includes("regular")||!Z.includes("earn"))return;let ie=Number(M[B.amount])||0;if(!ie)return;let ke=ia(M[B.payrollDate]),Se=P.get(R);(!Se||ke!=null&&ke>Se.timestamp)&&P.set(R,{payRate:ie/80,timestamp:ke})})}}let A=new Map;if(!l.isNullObject){let I=l.getUsedRangeOrNullObject();I.load("values"),await n.sync();let H=I.isNullObject?[]:I.values||[];if(H.length>1){let L=(H[0]||[]).map(R=>G(R)),B=L.findIndex(R=>R.includes("employee")&&R.includes("name")),M=$(L,["liability amount","liability","accrued pto"]);B>=0&&M>=0&&H.slice(1).forEach(R=>{let Z=G(R[B]);if(!Z)return;let ie=Number(R[M])||0;A.set(Z,ie)})}}let N=te(S.payrollDate)||"",x=[],m=[],O=k.map((I,H)=>{var ut,ft,pt,gt,mt,yt,ht;let L=G(I),B=T.get(L)||"",M=(ft=(ut=P.get(L))==null?void 0:ut.payRate)!=null?ft:"",R=C.get(L),Z=R&&d>=0&&(pt=R[d])!=null?pt:"",ie=R&&y>=0&&(gt=R[y])!=null?gt:"",ke=R&&v>=0&&(mt=R[v])!=null?mt:"",Se=R&&b>=0&&(yt=R[b])!=null?yt:"";(L.includes("avalos")||L.includes("sarah"))&&console.log(`[PTO Debug] ${I}:`,{ytdUsedIdx:b,rawValue:R?R[b]:"no dataRow",ytdUsed:Se,fullRow:R});let He=R&&p>=0&&Number(R[p])||0,ct=H+2;!M&&typeof M!="number"&&x.push({name:I,rowIndex:ct}),B||m.push({name:I,rowIndex:ct});let Fe=typeof M=="number"&&He?He*M:0,dt=(ht=A.get(L))!=null?ht:0,mn=(typeof Fe=="number"?Fe:0)-dt;return[N,I,B,M,Z,ie,ke,Se,He,Fe,dt,mn]});Y.missingPayRates=x.filter(I=>!Y.ignoredMissingPayRates.has(I.name)),Y.missingDepartments=m,console.log(`[PTO Analysis] Data quality: ${x.length} missing pay rates, ${m.length} missing departments`);let h=[["Analysis Date","Employee Name","Department","Pay Rate","Accrual Rate","Carry Over","YTD Accrued","YTD Used","Balance","Liability Amount","Accrued PTO $ [Prior Period]","Change"],...O],E=o.isNullObject?n.workbook.worksheets.add("PTO_Analysis"):o,V=E.getUsedRangeOrNullObject();V.load("address"),await n.sync(),V.isNullObject||V.clear();let W=h[0].length,Q=h.length,z=O.length,fe=E.getRangeByIndexes(0,0,Q,W);fe.values=h;let Ve=E.getRangeByIndexes(0,0,1,W);nt(Ve),z>0&&(Jt(E,0,z),ce(E,3,z),be(E,4,z),be(E,5,z),be(E,6,z),be(E,7,z),be(E,8,z),ce(E,9,z),ce(E,10,z),ce(E,11,z,!0)),fe.format.autofitColumns(),E.getRange("A1").select(),await n.sync()};q()&&(e?await t(e):await Excel.run(t))}function ya(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let a=String(n);return/[",\n]/.test(a)?`"${a.replace(/"/g,'""')}"`:a}).join(",")).join(`
`)}function ha(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),a=URL.createObjectURL(n),o=document.createElement("a");o.href=a,o.download=e,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(a),1e3)}function an(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=lt(),n=document.getElementById("step-notes-input"),a=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!a;let o=document.getElementById("headcount-notes-hint");o&&(o.textContent=t?"Please document outstanding differences before signing off.":"")}function on(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(pe)?t.slice(pe.length).replace(/^\s+/,""):t.replace(new RegExp(`^${pe}\\s*`,"i"),"").trimStart(),a=pe+(n?`
${n}`:"");e.value!==a&&(e.value=a),X(2,"notes",e.value)}function va(){let e=document.getElementById("step-notes-input");e&&e.addEventListener("input",()=>{if(!D.skipAnalysis)return;let t=e.value||"";if(!t.startsWith(pe)){let n=t.replace(pe,"").trimStart();e.value=pe+(n?`
${n}`:"")}X(2,"notes",e.value)})}function ba(){var n;let e=lt(),t=((n=document.getElementById("step-notes-input"))==null?void 0:n.value.trim())||"";if(e&&!t){window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");return}}})();
//# sourceMappingURL=app.bundle.js.map
