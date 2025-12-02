/* Prairie Forge PTO Accrual */
(()=>{function K(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}var Je="SS_PF_Config";async function wt(e,t=[Je]){var o;let n=e.workbook.tables;n.load("items/name"),await e.sync();let a=(o=n.items)==null?void 0:o.find(s=>t.includes(s.name));return a?e.workbook.tables.getItem(a.name):(console.warn("Config table not found. Looking for:",t),null)}function Ot(e){let t=e.map(n=>String(n||"").trim().toLowerCase());return{field:t.findIndex(n=>n==="field"||n==="field name"||n==="setting"),value:t.findIndex(n=>n==="value"||n==="setting value"),type:t.findIndex(n=>n==="type"||n==="category"),title:t.findIndex(n=>n==="title"||n==="display name"),permanent:t.findIndex(n=>n==="permanent"||n==="persist")}}async function kt(e=[Je]){if(!K())return{};try{return await Excel.run(async t=>{let n=await wt(t,e);if(!n)return{};let a=n.getDataBodyRange(),o=n.getHeaderRowRange();a.load("values"),o.load("values"),await t.sync();let s=o.values[0]||[],r=Ot(s);if(r.field===-1||r.value===-1)return console.warn("Config table missing FIELD or VALUE columns. Headers:",s),{};let l={};return(a.values||[]).forEach(i=>{var g;let p=String(i[r.field]||"").trim();p&&(l[p]=(g=i[r.value])!=null?g:"")}),console.log("Configuration loaded:",Object.keys(l).length,"fields"),l})}catch(t){return console.error("Failed to load configuration:",t),{}}}async function _e(e,t,n=[Je]){if(!K())return!1;try{return await Excel.run(async a=>{let o=await wt(a,n);if(!o){console.warn("Config table not found for write");return}let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let l=r.values[0]||[],c=Ot(l);if(c.field===-1||c.value===-1){console.error("Config table missing FIELD or VALUE columns");return}let p=(s.values||[]).findIndex(g=>String(g[c.field]||"").trim()===e);if(p>=0)s.getCell(p,c.value).values=[[t]];else{let g=new Array(l.length).fill("");c.type>=0&&(g[c.type]="Run Settings"),g[c.field]=e,g[c.value]=t,c.permanent>=0&&(g[c.permanent]="N"),c.title>=0&&(g[c.title]=""),o.rows.add(null,[g]),console.log("Added new config row:",e,"=",t)}await a.sync(),console.log("Saved config:",e,"=",t)}),!0}catch(a){return console.error("Failed to save config:",e,a),!1}}var yn="SS_PF_Config",hn="module-prefix",Ye="system",ve={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function St(){if(!K())return{...ve};try{return await Excel.run(async e=>{var p,g;let t=e.workbook.worksheets.getItemOrNullObject(yn);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...ve};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((p=n.values)!=null&&p.length))return{...ve};let a=n.values,o=wn(a[0]),s=o.get("category"),r=o.get("field"),l=o.get("value");if(s===void 0||r===void 0||l===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...ve};let c={},i=!1;for(let u=1;u<a.length;u++){let d=a[u];if(Ae(d[s])===hn){let h=String((g=d[r])!=null?g:"").trim().toUpperCase(),w=Ae(d[l]);h&&w&&(c[h]=w,i=!0)}}return i?(console.log("[Tab Visibility] Loaded prefix config:",c),c):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...ve})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...ve}}}async function qe(e){if(!K())return;let t=Ae(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await St();await Excel.run(async a=>{let o=a.workbook.worksheets;o.load("items/name,visibility"),await a.sync();let s={};for(let[u,d]of Object.entries(n))s[d]||(s[d]=[]),s[d].push(u);let r=s[t]||[],l=s[Ye]||[],c=[];for(let[u,d]of Object.entries(s))u!==t&&u!==Ye&&c.push(...d);console.log(`[Tab Visibility] Active prefixes: ${r.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${c.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${l.join(", ")}`);let i=[],p=[];o.items.forEach(u=>{let d=u.name,y=d.toUpperCase(),h=r.some(v=>y.startsWith(v)),w=c.some(v=>y.startsWith(v)),f=l.some(v=>y.startsWith(v));h?(i.push(u),console.log(`[Tab Visibility] SHOW: ${d} (matches active module prefix)`)):f?(p.push(u),console.log(`[Tab Visibility] HIDE: ${d} (system sheet)`)):w?(p.push(u),console.log(`[Tab Visibility] HIDE: ${d} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${d} (no prefix match, leaving as-is)`)});for(let u of i)u.visibility=Excel.SheetVisibility.visible;if(await a.sync(),o.items.filter(u=>u.visibility===Excel.SheetVisibility.visible).length>p.length){for(let u of p)try{u.visibility=Excel.SheetVisibility.hidden}catch(d){console.warn(`[Tab Visibility] Could not hide "${u.name}":`,d.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${i.length}, hid ${p.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function vn(){if(!K()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.visibility!==Excel.SheetVisibility.visible&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${a.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function bn(){if(!K()){console.log("Excel not available");return}try{let e=await St(),t=[];for(let[n,a]of Object.entries(e))a===Ye&&t.push(n);await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name,visibility"),await n.sync(),a.items.forEach(o=>{let s=o.name.toUpperCase();t.some(r=>s.startsWith(r))&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function wn(e=[]){let t=new Map;return e.forEach((n,a)=>{let o=Ae(n);o&&t.set(o,a)}),t}function Ae(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=vn,window.PrairieForge.unhideSystemSheets=bn,window.PrairieForge.applyModuleTabVisibility=qe);var Et={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var Pt=Et.ADA_IMAGE_URL;async function De(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async a=>{let o=a.workbook.worksheets.getItemOrNullObject(e);o.load("isNullObject, name"),await a.sync();let s;o.isNullObject?(s=a.workbook.worksheets.add(e),await a.sync(),await _t(a,s,t,n)):(s=o,await _t(a,s,t,n)),s.activate(),s.getRange("A1").select(),await a.sync()})}catch(a){console.error(`Error activating homepage sheet ${e}:`,a)}}async function _t(e,t,n,a){try{let i=t.getUsedRangeOrNullObject();i.load("isNullObject"),await e.sync(),i.isNullObject||(i.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let o=[[n,""],[a,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=o;let r=t.getRange("A1:Z100");r.format.fill.color="#0f0f0f";let l=t.getRange("A1");l.format.font.bold=!0,l.format.font.size=36,l.format.font.color="#ffffff",l.format.font.name="Segoe UI Light",l.format.verticalAlignment="Center";let c=t.getRange("A2");c.format.font.size=14,c.format.font.color="#a0a0a0",c.format.font.name="Segoe UI",c.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var Ct={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function $e(e){return Ct[e]||Ct["module-selector"]}function It(){We();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${Pt}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",On),e}function We(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function On(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${Pt}" alt="Ada" />
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",ze),t.addEventListener("click",o=>{o.target===t&&ze()});let a=o=>{o.key==="Escape"&&(ze(),document.removeEventListener("keydown",a))};document.addEventListener("keydown",a)}function ze(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var kn=["January","February","March","April","May","June","July","August","September","October","November","December"],Rt=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],Sn=["Su","Mo","Tu","We","Th","Fr","Sa"],be=null;function Tt(e,t={}){let n=document.getElementById(e);if(!n)return;let{onChange:a=null,minDate:o=null,maxDate:s=null,readonly:r=!1}=t,l=n.closest(".pf-datepicker-wrapper");l||(l=document.createElement("div"),l.className="pf-datepicker-wrapper",n.parentNode.insertBefore(l,n),l.appendChild(n)),n.type="text",n.placeholder="YYYY-MM-DD or click calendar",n.classList.add("pf-datepicker-input");let c=n.value?je(n.value):null,i=c?new Date(c):new Date;c&&(n.value=Ke(c));let p=document.createElement("span");p.className="pf-datepicker-icon",p.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>',l.appendChild(p);let g=document.createElement("div");g.className="pf-datepicker-dropdown",g.id=`${e}-dropdown`,l.appendChild(g);function u(){var b,T,C,D,$,A;let f=i.getFullYear(),v=i.getMonth();g.innerHTML=`
            <div class="pf-datepicker-header">
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev-year" title="Previous Year">\xAB</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev" title="Previous Month">\u2039</button>
                <span class="pf-datepicker-title">${kn[v]} ${f}</span>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next" title="Next Month">\u203A</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next-year" title="Next Year">\xBB</button>
            </div>
            <div class="pf-datepicker-weekdays">
                ${Sn.map(P=>`<span>${P}</span>`).join("")}
            </div>
            <div class="pf-datepicker-days">
                ${d(f,v,c)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-today">Today</button>
                <button type="button" class="pf-datepicker-clear">Clear</button>
            </div>
        `,(b=g.querySelector(".pf-datepicker-prev-year"))==null||b.addEventListener("mousedown",P=>{P.preventDefault(),P.stopPropagation(),i.setFullYear(i.getFullYear()-1),u()}),(T=g.querySelector(".pf-datepicker-prev"))==null||T.addEventListener("mousedown",P=>{P.preventDefault(),P.stopPropagation(),i.setMonth(i.getMonth()-1),u()}),(C=g.querySelector(".pf-datepicker-next"))==null||C.addEventListener("mousedown",P=>{P.preventDefault(),P.stopPropagation(),i.setMonth(i.getMonth()+1),u()}),(D=g.querySelector(".pf-datepicker-next-year"))==null||D.addEventListener("mousedown",P=>{P.preventDefault(),P.stopPropagation(),i.setFullYear(i.getFullYear()+1),u()}),g.querySelectorAll(".pf-datepicker-day:not(.disabled)").forEach(P=>{P.addEventListener("mousedown",M=>{M.preventDefault(),M.stopPropagation();let k=parseInt(P.dataset.day),E=parseInt(P.dataset.month),m=parseInt(P.dataset.year);y(new Date(m,E,k))})}),($=g.querySelector(".pf-datepicker-today"))==null||$.addEventListener("mousedown",P=>{P.preventDefault(),P.stopPropagation(),y(new Date)}),(A=g.querySelector(".pf-datepicker-clear"))==null||A.addEventListener("mousedown",P=>{P.preventDefault(),P.stopPropagation(),y(null)})}function d(f,v,b){let T=new Date(f,v,1).getDay(),C=new Date(f,v+1,0).getDate(),D=new Date(f,v,0).getDate(),$=new Date;$.setHours(0,0,0,0);let A="";for(let k=T-1;k>=0;k--){let E=D-k,m=v===0?11:v-1,x=v===0?f-1:f;A+=`<span class="pf-datepicker-day other-month" data-day="${E}" data-month="${m}" data-year="${x}">${E}</span>`}for(let k=1;k<=C;k++){let E=new Date(f,v,k),m=E.getTime()===$.getTime(),x=b&&E.getTime()===b.getTime(),N="pf-datepicker-day";m&&(N+=" today"),x&&(N+=" selected"),o&&E<o&&(N+=" disabled"),s&&E>s&&(N+=" disabled"),A+=`<span class="${N}" data-day="${k}" data-month="${v}" data-year="${f}">${k}</span>`}let M=Math.ceil((T+C)/7)*7-(T+C);for(let k=1;k<=M;k++){let E=v===11?0:v+1,m=v===11?f+1:f;A+=`<span class="pf-datepicker-day other-month" data-day="${k}" data-month="${E}" data-year="${m}">${k}</span>`}return A}function y(f){c=f,f?(n.value=Ke(f),n.dataset.value=Ce(f),i=new Date(f)):(n.value="",n.dataset.value=""),w(),a&&a(f?Ce(f):""),n.dispatchEvent(new Event("change",{bubbles:!0}))}function h(){if(!r){if(be&&be!==e){let f=document.getElementById(`${be}-dropdown`);f==null||f.classList.remove("open")}be=e,u(),g.classList.add("open"),l.classList.add("open")}}function w(){g.classList.remove("open"),l.classList.remove("open"),be===e&&(be=null)}return n.addEventListener("blur",f=>{if(g.classList.contains("open"))return;let v=n.value.trim();if(!v)return;let b=je(v);b&&(c=b,n.value=Ke(b),n.dataset.value=Ce(b),i=new Date(b),a&&a(Ce(b)),n.dispatchEvent(new Event("change",{bubbles:!0})))}),n.addEventListener("keydown",f=>{if(f.key==="Enter"){f.preventDefault();let v=n.value.trim(),b=je(v);b&&y(b),w()}}),n.addEventListener("click",f=>{f.stopPropagation(),g.classList.contains("open")||h()}),p.addEventListener("click",f=>{f.stopPropagation(),g.classList.contains("open")?w():h()}),document.addEventListener("click",f=>{l.contains(f.target)||w()}),g.addEventListener("click",f=>{f.stopPropagation()}),document.addEventListener("keydown",f=>{f.key==="Escape"&&w()}),{getValue:()=>c?Ce(c):"",setValue:f=>{let v=je(f);y(v)},open:h,close:w}}function je(e){if(!e)return null;if(/^\d{4}-\d{2}-\d{2}$/.test(e)){let[a,o,s]=e.split("-").map(Number);return new Date(a,o-1,s)}let t=e.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);if(t){let a=Rt.findIndex(o=>o.toLowerCase()===t[1].toLowerCase().substring(0,3));if(a>=0)return new Date(parseInt(t[3]),a,parseInt(t[2]))}if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(e)){let[a,o,s]=e.split("/").map(Number);return new Date(s,a-1,o)}let n=new Date(e);return isNaN(n.getTime())?null:n}function Ke(e){return e?`${Rt[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`:""}function Ce(e){if(!e)return"";let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}var xt=`
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
`.trim(),Qe=`
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
    `};function jt(e){return e&&En[e]||""}var Xe=`
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
`.trim(),Ze=`
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
`.trim(),Le=`
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
`.trim(),et=`
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
`.trim();function Re(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function J(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function ce({textareaId:e,value:t,permanentId:n,isPermanent:a,hintId:o,saveButtonId:s,isSaved:r=!1,placeholder:l="Enter notes here..."}){let c=a?Ze:Xe,i=s?`<button type="button" class="pf-action-toggle pf-save-btn ${r?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Vt}</button>`:"",p=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${a?"is-locked":""}" id="${n}" aria-pressed="${a}" title="Lock notes (retain after archive)">${c}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Re(l)}">${Re(t||"")}</textarea>
                ${o?`<p class="pf-signoff-hint" id="${o}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?J(p,"Lock"):""}
                ${s?J(i,"Save"):""}
            </div>
        </article>
    `}function de({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:a,isComplete:o,saveButtonId:s,isSaved:r=!1,completeButtonId:l,subtext:c="Sign-off below. Click checkmark icon. Done."}){let i=`<button type="button" class="pf-action-toggle ${o?"is-active":""}" id="${l}" aria-pressed="${!!o}" title="Mark step complete">${Pe}</button>`;return`
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
                ${J(i,"Done")}
            </div>
        </article>
    `}function tt(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?Ze:Xe)}function se(e,t){e&&e.classList.toggle("is-saved",t)}function nt(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(a=>{let o=a.getAttribute("data-save-input"),s=document.getElementById(o);if(!s)return;let r=()=>{se(a,!1)};s.addEventListener("input",r),n.push(()=>s.removeEventListener("input",r))}),()=>n.forEach(a=>a())}function Ut(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function Gt(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var at={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},Be={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function ot(e){e.format.fill.color=at.fillColor,e.format.font.color=at.fontColor,e.format.font.bold=at.bold}function ue(e,t,n,a=!1){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a?Be.currencyWithNegative:Be.currency]]}function we(e,t,n){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[Be.number]]}function Jt(e,t,n,a=Be.date){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a]]}var _n="1.1.0",xe="pto-accrual";var fe="PTO Accrual",Cn="Calculate your PTO liability, compare against last period, and generate a balanced journal entry\u2014all without leaving Excel.",Pn="../module-selector/index.html",In="pf-loader-overlay",pe=["SS_PF_Config"],S={payrollProvider:"PTO_Payroll_Provider",payrollDate:"PTO_Analysis_Date",accountingPeriod:"PTO_Accounting_Period",journalEntryId:"PTO_Journal_Entry_ID",companyName:"SS_Company_Name",accountingSoftware:"SS_Accounting_Software",reviewerName:"PTO_Reviewer",validationDataBalance:"PTO_Validation_Data_Balance",validationCleanBalance:"PTO_Validation_Clean_Balance",validationDifference:"PTO_Validation_Difference",headcountRosterCount:"PTO_Headcount_Roster_Count",headcountPayrollCount:"PTO_Headcount_Payroll_Count",headcountDifference:"PTO_Headcount_Difference",journalDebitTotal:"PTO_JE_Debit_Total",journalCreditTotal:"PTO_JE_Credit_Total",journalDifference:"PTO_JE_Difference"},ge="User opted to skip the headcount review this period.",He={0:{note:"PTO_Notes_Config",reviewer:"PTO_Reviewer_Config",signOff:"PTO_SignOff_Config"},1:{note:"PTO_Notes_Import",reviewer:"PTO_Reviewer_Import",signOff:"PTO_SignOff_Import"},2:{note:"PTO_Notes_Headcount",reviewer:"PTO_Reviewer_Headcount",signOff:"PTO_SignOff_Headcount"},3:{note:"PTO_Notes_Validate",reviewer:"PTO_Reviewer_Validate",signOff:"PTO_SignOff_Validate"},4:{note:"PTO_Notes_Review",reviewer:"PTO_Reviewer_Review",signOff:"PTO_SignOff_Review"},5:{note:"PTO_Notes_JE",reviewer:"PTO_Reviewer_JE",signOff:"PTO_SignOff_JE"},6:{note:"PTO_Notes_Archive",reviewer:"PTO_Reviewer_Archive",signOff:"PTO_SignOff_Archive"}},sn={0:"PTO_Complete_Config",1:"PTO_Complete_Import",2:"PTO_Complete_Headcount",3:"PTO_Complete_Validate",4:"PTO_Complete_Review",5:"PTO_Complete_JE",6:"PTO_Complete_Archive"};var ie=[{id:0,title:"Configuration",summary:"Set the analysis date, accounting period, and review details for this run.",description:"Complete this step first to ensure all downstream calculations use the correct period settings.",actionLabel:"Configure Workbook",secondaryAction:{sheet:"SS_PF_Config",label:"Open Config Sheet"}},{id:1,title:"Import PTO Data",summary:"Pull your latest PTO export from payroll and paste it into PTO_Data.",description:"Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",actionLabel:"Import Sample Data",secondaryAction:{sheet:"PTO_Data",label:"Open Data Sheet"}},{id:2,title:"Headcount Review",summary:"Quick check to make sure your roster matches your PTO data.",description:"Compare employees in PTO_Data against your employee roster to catch any discrepancies.",actionLabel:"Open Headcount Review",secondaryAction:{sheet:"SS_Employee_Roster",label:"Open Sheet"}},{id:3,title:"Data Quality Review",summary:"Scan your PTO data for potential errors before crunching numbers.",description:"Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",actionLabel:"Click to Run Quality Check"},{id:4,title:"PTO Accrual Review",summary:"Review the calculated liability for each employee and compare to last period.",description:"The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",actionLabel:"Click to Perform Review"},{id:5,title:"Journal Entry Prep",summary:"Generate a balanced journal entry, run validation checks, and export when ready.",description:"Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",actionLabel:"Open Journal Draft",secondaryAction:{sheet:"PTO_JE_Draft",label:"Open Sheet"}},{id:6,title:"Archive & Reset",summary:"Save this period's results and prepare for the next cycle.",description:"Archive the current analysis so it becomes the 'prior period' for your next review.",actionLabel:"Archive Run"}];var Rn=ie.reduce((e,t)=>(e[t.id]="pending",e),{}),V={activeView:"home",activeStepId:null,focusedIndex:0,stepStatuses:Rn},I={loaded:!1,steps:{},permanents:{},completes:{},values:{},overrides:{accountingPeriod:!1,journalId:!1}},Te=null,st=null,Me=null,Oe=new Map,j={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},Y={debitTotal:null,creditTotal:null,difference:null,lineAmountSum:null,analysisChangeTotal:null,jeChangeTotal:null,loading:!1,lastError:null,validationRun:!1,issues:[]},W={hasRun:!1,loading:!1,acknowledged:!1,balanceIssues:[],zeroBalances:[],accrualOutliers:[],totalIssues:0,totalEmployees:0},Q={cleanDataReady:!1,employeeCount:0,lastRun:null,loading:!1,lastError:null,missingPayRates:[],missingDepartments:[],ignoredMissingPayRates:new Set,completenessCheck:{accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null}};async function Tn(){var e;try{Te=document.getElementById("app"),st=document.getElementById("loading"),await xn(),await Nn(),(e=window.PrairieForge)!=null&&e.loadSharedConfig&&await window.PrairieForge.loadSharedConfig();let t=$e(xe);await De(t.sheetName,t.title,t.subtitle),st&&st.remove(),Te&&(Te.hidden=!1),ae()}catch(t){throw console.error("[PTO] Module initialization failed:",t),t}}async function xn(){try{await qe(xe),console.log(`[PTO] Tab visibility applied for ${xe}`)}catch(e){console.warn("[PTO] Could not apply tab visibility:",e)}}async function Nn(){var e;if(!K()){I.loaded=!0;return}try{let t=await kt(pe),n={};(e=window.PrairieForge)!=null&&e.loadSharedConfig&&(await window.PrairieForge.loadSharedConfig(),window.PrairieForge._sharedConfigCache&&window.PrairieForge._sharedConfigCache.forEach((s,r)=>{n[r]=s}));let a={...t},o={SS_Default_Reviewer:S.reviewerName,Default_Reviewer:S.reviewerName,PTO_Reviewer:S.reviewerName,SS_Company_Name:S.companyName,Company_Name:S.companyName,SS_Payroll_Provider:S.payrollProvider,Payroll_Provider_Link:S.payrollProvider,SS_Accounting_Software:S.accountingSoftware,Accounting_Software_Link:S.accountingSoftware};Object.entries(o).forEach(([s,r])=>{n[s]&&!a[r]&&(a[r]=n[s])}),Object.entries(n).forEach(([s,r])=>{s.startsWith("PTO_")&&r&&(a[s]=r)}),I.permanents=await An(),I.values=a||{},I.overrides.accountingPeriod=!!(a!=null&&a[S.accountingPeriod]),I.overrides.journalId=!!(a!=null&&a[S.journalEntryId]),Object.entries(He).forEach(([s,r])=>{var l,c,i;I.steps[s]={notes:(l=a[r.note])!=null?l:"",reviewer:(c=a[r.reviewer])!=null?c:"",signOffDate:(i=a[r.signOff])!=null?i:""}}),I.completes=Object.entries(sn).reduce((s,[r,l])=>{var c;return s[r]=(c=a[l])!=null?c:"",s},{}),I.loaded=!0}catch(t){console.warn("PTO: unable to load configuration fields",t),I.loaded=!0}}async function An(){let e={};if(!K())return e;let t=new Map;Object.entries(He).forEach(([n,a])=>{a.note&&t.set(a.note.trim(),Number(n))});try{await Excel.run(async n=>{let a=n.workbook.tables.getItemOrNullObject(pe[0]);if(await n.sync(),a.isNullObject)return;let o=a.getDataBodyRange(),s=a.getHeaderRowRange();o.load("values"),s.load("values"),await n.sync();let l=(s.values[0]||[]).map(i=>String(i||"").trim().toLowerCase()),c={field:l.findIndex(i=>i==="field"||i==="field name"||i==="setting"),permanent:l.findIndex(i=>i==="permanent"||i==="persist")};c.field===-1||c.permanent===-1||(o.values||[]).forEach(i=>{let p=String(i[c.field]||"").trim(),g=t.get(p);if(g==null)return;let u=sa(i[c.permanent]);e[g]=u})})}catch(n){console.warn("PTO: unable to load permanent flags",n)}return e}function ae(){var l;if(!Te)return;let e=V.focusedIndex<=0?"disabled":"",t=V.focusedIndex>=ie.length-1?"disabled":"",n=V.activeView==="step"&&V.activeStepId!=null,o=V.activeView==="config"?rn():n?Vn(V.activeStepId):`${$n()}${jn()}`;Te.innerHTML=`
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
    `;let s=V.activeView==="home"||V.activeView!=="step"&&V.activeView!=="config",r=document.getElementById("pf-info-fab-pto");if(s)r&&r.remove();else if((l=window.PrairieForge)!=null&&l.mountInfoFab){let c=Dn(V.activeStepId);PrairieForge.mountInfoFab({title:c.title,content:c.content,buttonId:"pf-info-fab-pto"})}Hn(),Gn(),s?It():We()}function Dn(e){switch(e){case 0:return{title:"Configuration",content:`
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
                ${ie.map((e,t)=>Ln(e,t)).join("")}
            </div>
        </section>
    `}function Ln(e,t){let n=V.stepStatuses[e.id]||"pending",a=V.activeView==="step"&&V.focusedIndex===t?"pf-step-card--active":"",o=jt(ea(e.id));return`
        <article class="pf-step-card pf-clickable ${a}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${e.title}</h3>
        </article>
    `}function Bn(e){let t=ie.filter(o=>o.id!==6).map(o=>({id:o.id,title:o.title,complete:Jn(o.id)})),n=t.every(o=>o.complete),a=t.map(o=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${o.complete?"is-active":""}" aria-pressed="${o.complete}">
                        ${Pe}
                    </span>
                    <div>
                        <h3>${O(o.title)}</h3>
                        <p class="pf-config-subtext">${o.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${O(fe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${O(e.title)}</h2>
            <p class="pf-hero-copy">${O(e.summary||"")}</p>
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
    `}function rn(){if(!I.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=Zt(ne(S.payrollDate)),t=Zt(ne(S.accountingPeriod)),n=ne(S.journalEntryId),a=ne(S.accountingSoftware),o=ne(S.payrollProvider),s=ne(S.companyName),r=ne(S.reviewerName),l=he(0),c=!!I.permanents[0],i=!!(fn(I.completes[0])||l.signOffDate),p=ye(l==null?void 0:l.reviewer),g=(l==null?void 0:l.signOffDate)||"";return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${O(fe)} | Step 0</p>
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
                        <input type="text" id="config-user-name" value="${O(r)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${O(e)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${O(t)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${O(n)}" placeholder="PTO-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${O(s)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${O(o)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${O(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${ce({textareaId:"config-notes",value:l.notes||"",permanentId:"config-notes-lock",isPermanent:c,hintId:"",saveButtonId:"config-notes-save"})}
            ${de({reviewerInputId:"config-reviewer",reviewerValue:p,signoffInputId:"config-signoff-date",signoffValue:g,isComplete:i,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"})}
        </section>
    `}function Mn(e){let t=he(1),n=!!I.permanents[1],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(ke(I.completes[1])||o),r=ne(S.payrollProvider);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${O(fe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${O(e.title)}</h2>
            <p class="pf-hero-copy">${O(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export, then paste into PTO_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${J(r?`<a href="${O(r)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${et}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${et}</button>`,"Provider")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${Qe}</button>`,"PTO_Data")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${Ft}</button>`,"Clear")}
                </div>
            </article>
            ${ce({textareaId:"step-notes-1",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-1",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-1"})}
            ${de({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
        </section>
    `}function Vn(e){let t=ie.find(l=>l.id===e);if(!t)return"";if(e===0)return rn();if(e===1)return Mn(t);if(e===2)return ca(t);if(e===3)return ua(t);if(e===4)return pa(t);if(e===5)return fa(t);if(t.id===6)return Bn(t);let n=he(e),a=!!I.permanents[e],o=ye(n==null?void 0:n.reviewer),s=(n==null?void 0:n.signOffDate)||"",r=!!(ke(I.completes[e])||s);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${O(fe)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${O(t.title)}</h2>
            <p class="pf-hero-copy">${O(t.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${ce({textareaId:`step-notes-${e}`,value:(n==null?void 0:n.notes)||"",permanentId:`step-notes-lock-${e}`,isPermanent:a,hintId:"",saveButtonId:`step-notes-save-${e}`})}
            ${de({reviewerInputId:`step-reviewer-${e}`,reviewerValue:o,signoffInputId:`step-signoff-${e}`,signoffValue:s,isComplete:r,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`})}
        </section>
    `}function Hn(){var n,a,o,s,r,l,c;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",async()=>{var p;let i=$e(xe);await De(i.sheetName,i.title,i.subtitle),Ne({activeView:"home",activeStepId:null}),(p=document.getElementById("pf-hero"))==null||p.scrollIntoView({behavior:"smooth",block:"start"})}),(a=document.getElementById("nav-selector"))==null||a.addEventListener("click",()=>{window.location.href=Pn}),(o=document.getElementById("nav-prev"))==null||o.addEventListener("click",()=>Yt(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>Yt(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",i=>{i.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",i=>{!(t!=null&&t.contains(i.target))&&!(e!=null&&e.contains(i.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(r=document.getElementById("nav-roster"))==null||r.addEventListener("click",()=>{Qt("SS_Employee_Roster"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(l=document.getElementById("nav-accounts"))==null||l.addEventListener("click",()=>{Qt("SS_Chart_of_Accounts"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(c=document.getElementById("showConfigSheets"))==null||c.addEventListener("click",async()=>{await Zn()}),document.querySelectorAll("[data-step-card]").forEach(i=>{let p=Number(i.getAttribute("data-step-index")),g=Number(i.getAttribute("data-step-id"));i.addEventListener("click",()=>Fe(p,g))}),V.activeView==="config"?Un():V.activeView==="step"&&V.activeStepId!=null&&Fn(V.activeStepId)}function Fn(e){var p,g,u,d,y,h,w,f,v,b,T,C,D,$,A,P,M;let t=e===2?document.getElementById("step-notes-input"):document.getElementById(`step-notes-${e}`),n=e===2?document.getElementById("step-reviewer-name"):document.getElementById(`step-reviewer-${e}`),a=e===2?document.getElementById("step-signoff-date"):document.getElementById(`step-signoff-${e}`),o=document.getElementById("step-back-btn"),s=e===2?document.getElementById("step-notes-lock-2"):document.getElementById(`step-notes-lock-${e}`),r=e===2?document.getElementById("step-notes-save-2"):document.getElementById(`step-notes-save-${e}`);r==null||r.addEventListener("click",async()=>{let k=(t==null?void 0:t.value)||"";await Z(e,"notes",k),se(r,!0)});let l=e===2?document.getElementById("headcount-signoff-save"):document.getElementById(`step-signoff-save-${e}`);l==null||l.addEventListener("click",async()=>{let k=(n==null?void 0:n.value)||"";await Z(e,"reviewer",k),se(l,!0)}),nt();let c=e===2?"headcount-signoff-toggle":`step-signoff-toggle-${e}`,i=e===2?"step-signoff-date":`step-signoff-${e}`;pn(e,{buttonId:c,inputId:i,canActivate:e===2?()=>{var E;return!dt()||((E=document.getElementById("step-notes-input"))==null?void 0:E.value.trim())||""?!0:(window.alert("Please enter a brief explanation of the headcount differences before completing this step."),!1)}:null,onComplete:e===2?ba:null}),o==null||o.addEventListener("click",async()=>{let k=$e(xe);await De(k.sheetName,k.title,k.subtitle),Ne({activeView:"home",activeStepId:null})}),s==null||s.addEventListener("click",async()=>{let k=!s.classList.contains("is-locked");tt(s,k),await dn(e,k)}),e===6&&((p=document.getElementById("archive-run-btn"))==null||p.addEventListener("click",()=>{})),e===1&&((g=document.getElementById("import-open-data-btn"))==null||g.addEventListener("click",()=>Ve("PTO_Data")),(u=document.getElementById("import-clear-btn"))==null||u.addEventListener("click",()=>Xn())),e===2&&((d=document.getElementById("headcount-skip-btn"))==null||d.addEventListener("click",()=>{j.skipAnalysis=!j.skipAnalysis;let k=document.getElementById("headcount-skip-btn");k==null||k.classList.toggle("is-active",j.skipAnalysis),j.skipAnalysis&&on(),an()}),(y=document.getElementById("headcount-run-btn"))==null||y.addEventListener("click",()=>rt()),(h=document.getElementById("headcount-refresh-btn"))==null||h.addEventListener("click",()=>rt()),va(),j.skipAnalysis&&on(),an()),e===3&&((w=document.getElementById("quality-run-btn"))==null||w.addEventListener("click",()=>zt()),(f=document.getElementById("quality-refresh-btn"))==null||f.addEventListener("click",()=>zt()),(v=document.getElementById("quality-acknowledge-btn"))==null||v.addEventListener("click",()=>qn())),e===4&&((b=document.getElementById("analysis-refresh-btn"))==null||b.addEventListener("click",()=>Wt()),(T=document.getElementById("analysis-run-btn"))==null||T.addEventListener("click",()=>Wt()),(C=document.getElementById("payrate-save-btn"))==null||C.addEventListener("click",qt),(D=document.getElementById("payrate-ignore-btn"))==null||D.addEventListener("click",Yn),($=document.getElementById("payrate-input"))==null||$.addEventListener("keydown",k=>{k.key==="Enter"&&qt()})),e===5&&((A=document.getElementById("je-create-btn"))==null||A.addEventListener("click",()=>Kn()),(P=document.getElementById("je-run-btn"))==null||P.addEventListener("click",()=>ln()),(M=document.getElementById("je-export-btn"))==null||M.addEventListener("click",()=>Qn()))}function Un(){var l,c,i,p,g;Tt("config-payroll-date",{onChange:u=>{if(te(S.payrollDate,u),!!u){if(!I.overrides.accountingPeriod){let d=aa(u);if(d){let y=document.getElementById("config-accounting-period");y&&(y.value=d),te(S.accountingPeriod,d)}}if(!I.overrides.journalId){let d=oa(u);if(d){let y=document.getElementById("config-journal-id");y&&(y.value=d),te(S.journalEntryId,d)}}}}});let e=document.getElementById("config-accounting-period");e==null||e.addEventListener("change",u=>{I.overrides.accountingPeriod=!!u.target.value,te(S.accountingPeriod,u.target.value||"")});let t=document.getElementById("config-journal-id");t==null||t.addEventListener("change",u=>{I.overrides.journalId=!!u.target.value,te(S.journalEntryId,u.target.value.trim())}),(l=document.getElementById("config-company-name"))==null||l.addEventListener("change",u=>{te(S.companyName,u.target.value.trim())}),(c=document.getElementById("config-payroll-provider"))==null||c.addEventListener("change",u=>{te(S.payrollProvider,u.target.value.trim())}),(i=document.getElementById("config-accounting-link"))==null||i.addEventListener("change",u=>{te(S.accountingSoftware,u.target.value.trim())}),(p=document.getElementById("config-user-name"))==null||p.addEventListener("change",u=>{te(S.reviewerName,u.target.value.trim())});let n=document.getElementById("config-notes");n==null||n.addEventListener("input",u=>{Z(0,"notes",u.target.value)});let a=document.getElementById("config-notes-lock");a==null||a.addEventListener("click",async()=>{let u=!a.classList.contains("is-locked");tt(a,u),await dn(0,u)});let o=document.getElementById("config-notes-save");o==null||o.addEventListener("click",async()=>{n&&(await Z(0,"notes",n.value),se(o,!0))});let s=document.getElementById("config-reviewer");s==null||s.addEventListener("change",u=>{let d=u.target.value.trim();Z(0,"reviewer",d);let y=document.getElementById("config-signoff-date");if(d&&y&&!y.value){let h=ct();y.value=h,Z(0,"signOffDate",h),un(0,!0)}}),(g=document.getElementById("config-signoff-date"))==null||g.addEventListener("change",u=>{Z(0,"signOffDate",u.target.value||"")});let r=document.getElementById("config-signoff-save");r==null||r.addEventListener("click",async()=>{var y,h;let u=((y=s==null?void 0:s.value)==null?void 0:y.trim())||"",d=((h=document.getElementById("config-signoff-date"))==null?void 0:h.value)||"";await Z(0,"reviewer",u),await Z(0,"signOffDate",d),se(r,!0)}),nt(),pn(0,{buttonId:"config-signoff-toggle",inputId:"config-signoff-date",onComplete:la})}function Fe(e,t=null){if(e<0||e>=ie.length)return;Me=e;let n=t!=null?t:ie[e].id;Ne({focusedIndex:e,activeView:n===0?"config":"step",activeStepId:n}),n===1&&Ve("PTO_Data"),n===2&&!j.hasAnalyzed&&(gn(),rt()),n===3&&Ve("PTO_Data"),n===5&&Ve("PTO_JE_Draft")}function Yt(e){let t=V.focusedIndex+e,n=Math.max(0,Math.min(ie.length-1,t));Fe(n,ie[n].id)}function Gn(){if(Me===null)return;let e=document.querySelector(`[data-step-index="${Me}"]`);Me=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}function Jn(e){return fn(I.completes[e])}function Ne(e){e.stepStatuses&&(V.stepStatuses={...V.stepStatuses,...e.stepStatuses}),Object.assign(V,{...e,stepStatuses:V.stepStatuses}),ae()}function oe(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}async function qt(){let e=document.getElementById("payrate-input");if(!e)return;let t=parseFloat(e.value),n=e.dataset.employee,a=parseInt(e.dataset.row,10);if(isNaN(t)||t<=0){window.alert("Please enter a valid pay rate greater than 0.");return}if(!n||isNaN(a)){console.error("Missing employee data on input");return}X(!0,"Updating pay rate...");try{await Excel.run(async o=>{let s=o.workbook.worksheets.getItem("PTO_Analysis"),r=s.getCell(a-1,3);r.values=[[t]];let l=s.getCell(a-1,8);l.load("values"),await o.sync();let i=(Number(l.values[0][0])||0)*t,p=s.getCell(a-1,9);p.values=[[i]];let g=s.getCell(a-1,10);g.load("values"),await o.sync();let u=Number(g.values[0][0])||0,d=i-u,y=s.getCell(a-1,11);y.values=[[d]],await o.sync()}),Q.missingPayRates=Q.missingPayRates.filter(o=>o.name!==n),X(!1),Fe(3,3)}catch(o){console.error("Failed to save pay rate:",o),window.alert(`Failed to save pay rate: ${o.message}`),X(!1)}}function Yn(){let e=document.getElementById("payrate-input");if(!e)return;let t=e.dataset.employee;t&&(Q.ignoredMissingPayRates.add(t),Q.missingPayRates=Q.missingPayRates.filter(n=>n.name!==t)),Fe(3,3)}async function zt(){if(!oe()){window.alert("Excel is not available. Open this module inside Excel to run quality check.");return}W.loading=!0,X(!0,"Analyzing data quality..."),se(document.getElementById("quality-save-btn"),!1);try{await Excel.run(async t=>{var w;let a=t.workbook.worksheets.getItem("PTO_Data").getUsedRangeOrNullObject();a.load("values"),await t.sync();let o=a.isNullObject?[]:a.values||[];if(!o.length||o.length<2)throw new Error("PTO_Data is empty or has no data rows.");let s=(o[0]||[]).map(f=>q(f));console.log("[Data Quality] PTO_Data headers:",o[0]);let r=s.findIndex(f=>f==="employee name"||f==="employeename");r===-1&&(r=s.findIndex(f=>f.includes("employee")&&f.includes("name"))),r===-1&&(r=s.findIndex(f=>f==="name"||f.includes("name")&&!f.includes("company")&&!f.includes("form"))),console.log("[Data Quality] Employee name column index:",r,"Header:",(w=o[0])==null?void 0:w[r]);let l=B(s,["balance"]),c=B(s,["accrual rate","accrualrate"]),i=B(s,["carry over","carryover"]),p=B(s,["ytd accrued","ytdaccrued"]),g=B(s,["ytd used","ytdused"]),u=[],d=[],y=[],h=o.slice(1);h.forEach((f,v)=>{let b=v+2,T=r!==-1?String(f[r]||"").trim():`Row ${b}`;if(!T)return;let C=l!==-1&&Number(f[l])||0,D=c!==-1&&Number(f[c])||0,$=i!==-1&&Number(f[i])||0,A=p!==-1&&Number(f[p])||0,P=g!==-1&&Number(f[g])||0,M=$+A;C<0?u.push({name:T,issue:`Negative balance: ${C.toFixed(2)} hrs`,rowIndex:b}):P>M&&M>0&&u.push({name:T,issue:`Used ${P.toFixed(0)} hrs but only ${M.toFixed(0)} available`,rowIndex:b}),C===0&&($>0||A>0)&&d.push({name:T,rowIndex:b}),D>8&&y.push({name:T,accrualRate:D,rowIndex:b})}),W.balanceIssues=u,W.zeroBalances=d,W.accrualOutliers=y,W.totalIssues=u.length,W.totalEmployees=h.filter(f=>f.some(v=>v!==null&&v!=="")).length,W.hasRun=!0});let e=W.balanceIssues.length>0;Ne({stepStatuses:{3:e?"blocked":"complete"}})}catch(e){console.error("Data quality check error:",e),window.alert(`Quality check failed: ${e.message}`),W.hasRun=!1}finally{W.loading=!1,X(!1),ae()}}function qn(){W.acknowledged=!0,Ne({stepStatuses:{3:"complete"}}),ae()}async function zn(){if(oe())try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem("PTO_Data"),n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=t.getUsedRangeOrNullObject();if(a.load("values"),n.load("isNullObject"),await e.sync(),n.isNullObject){Q.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let o=n.getUsedRangeOrNullObject();o.load("values"),await e.sync();let s=a.isNullObject?[]:a.values||[],r=o.isNullObject?[]:o.values||[];if(!s.length||!r.length){Q.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let l=(p,g,u)=>{let d=(p[0]||[]).map(w=>q(w)),y=B(d,g);return y===-1?null:p.slice(1).reduce((w,f)=>w+(Number(f[y])||0),0)},c=[{key:"accrualRate",aliases:["accrual rate","accrualrate"]},{key:"carryOver",aliases:["carry over","carryover","carry_over"]},{key:"ytdAccrued",aliases:["ytd accrued","ytdaccrued","ytd_accrued"]},{key:"ytdUsed",aliases:["ytd used","ytdused","ytd_used"]},{key:"balance",aliases:["balance"]}],i={};for(let p of c){let g=l(s,p.aliases,"PTO_Data"),u=l(r,p.aliases,"PTO_Analysis");if(g===null||u===null)i[p.key]=null;else{let d=Math.abs(g-u)<.01;i[p.key]={match:d,ptoData:g,ptoAnalysis:u}}}Q.completenessCheck=i})}catch(e){console.error("Completeness check failed:",e)}}async function Wt(){if(!oe()){window.alert("Excel is not available. Open this module inside Excel to run analysis.");return}X(!0,"Running analysis...");try{await gn(),await zn(),Q.cleanDataReady=!0,ae()}catch(e){console.error("Full analysis error:",e),window.alert(`Analysis failed: ${e.message}`)}finally{X(!1)}}async function ln(){if(!oe()){window.alert("Excel is not available. Open this module inside Excel to run journal checks.");return}Y.loading=!0,Y.lastError=null,se(document.getElementById("je-save-btn"),!1),ae();try{let e=await Excel.run(async t=>{let a=t.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();a.load("values");let o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis");o.load("isNullObject"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty. Generate the JE first.");let r=(s[0]||[]).map(b=>q(b)),l=B(r,["debit"]),c=B(r,["credit"]),i=B(r,["lineamount","line amount"]),p=B(r,["account number","accountnumber"]);if(l===-1||c===-1)throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");let g=0,u=0,d=0,y=0;s.slice(1).forEach(b=>{let T=Number(b[l])||0,C=Number(b[c])||0,D=i!==-1&&Number(b[i])||0,$=p!==-1?String(b[p]||"").trim():"";g+=T,u+=C,d+=D,$&&$!=="21540"&&(y+=D)});let h=0;if(!o.isNullObject){let b=o.getUsedRangeOrNullObject();b.load("values"),await t.sync();let T=b.isNullObject?[]:b.values||[];if(T.length>1){let C=(T[0]||[]).map($=>q($)),D=B(C,["change"]);D!==-1&&T.slice(1).forEach($=>{h+=Number($[D])||0})}}let w=g-u,f=[];Math.abs(w)>=.01?f.push({check:"Debits = Credits",passed:!1,detail:w>0?`Debits exceed credits by $${Math.abs(w).toLocaleString(void 0,{minimumFractionDigits:2})}`:`Credits exceed debits by $${Math.abs(w).toLocaleString(void 0,{minimumFractionDigits:2})}`}):f.push({check:"Debits = Credits",passed:!0,detail:""}),Math.abs(d)>=.01?f.push({check:"Line Amounts Sum to Zero",passed:!1,detail:`Line amounts sum to $${d.toLocaleString(void 0,{minimumFractionDigits:2})} (should be $0.00)`}):f.push({check:"Line Amounts Sum to Zero",passed:!0,detail:""});let v=Math.abs(y-h);return v>=.01?f.push({check:"JE Matches Analysis Total",passed:!1,detail:`JE expense total ($${y.toLocaleString(void 0,{minimumFractionDigits:2})}) differs from PTO_Analysis Change total ($${h.toLocaleString(void 0,{minimumFractionDigits:2})}) by $${v.toLocaleString(void 0,{minimumFractionDigits:2})}`}):f.push({check:"JE Matches Analysis Total",passed:!0,detail:""}),{debitTotal:g,creditTotal:u,difference:w,lineAmountSum:d,jeChangeTotal:y,analysisChangeTotal:h,issues:f,validationRun:!0}});Object.assign(Y,e,{lastError:null})}catch(e){console.warn("PTO JE summary:",e),Y.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",Y.debitTotal=null,Y.creditTotal=null,Y.difference=null,Y.lineAmountSum=null,Y.jeChangeTotal=null,Y.analysisChangeTotal=null,Y.issues=[],Y.validationRun=!1}finally{Y.loading=!1,ae()}}var Wn={"general & administrative":"64110","general and administrative":"64110","g&a":"64110","research & development":"62110","research and development":"62110","r&d":"62110",marketing:"61610","cogs onboarding":"53110","cogs prof. services":"56110","cogs professional services":"56110","sales & marketing":"61110","sales and marketing":"61110","cogs support":"52110","client success":"61811"},Kt="21540";async function Kn(){if(!oe()){window.alert("Excel is not available. Open this module inside Excel to create the journal entry.");return}X(!0,"Creating PTO Journal Entry...");try{await Excel.run(async e=>{let t=[],n=e.workbook.tables.getItemOrNullObject(pe[0]);if(n.load("isNullObject"),await e.sync(),n.isNullObject){let m=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");if(m.load("isNullObject"),await e.sync(),!m.isNullObject){let x=m.getUsedRangeOrNullObject();x.load("values"),await e.sync();let N=x.isNullObject?[]:x.values||[];t=N.length>1?N.slice(1):[]}}else{let m=n.getDataBodyRange();m.load("values"),await e.sync(),t=m.values||[]}let a=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis");if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PTO_Analysis sheet not found. Please ensure the worksheet exists.");let o=a.getUsedRangeOrNullObject();o.load("values");let s=e.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");s.load("isNullObject"),await e.sync();let r=[];if(!s.isNullObject){let m=s.getUsedRangeOrNullObject();m.load("values"),await e.sync(),r=m.isNullObject?[]:m.values||[]}let l=o.isNullObject?[]:o.values||[];if(!l.length||l.length<2)throw new Error("PTO_Analysis is empty or has no data rows. Run the analysis first (Step 4).");let c={};t.forEach(m=>{let x=String(m[1]||"").trim(),N=m[2];x&&(c[x]=N)}),(!c[S.journalEntryId]||!c[S.payrollDate])&&console.warn("[JE Draft] Missing config values - RefNumber:",c[S.journalEntryId],"TxnDate:",c[S.payrollDate]);let i=c[S.journalEntryId]||"",p=c[S.payrollDate]||"",g=c[S.accountingPeriod]||"",u="";if(p)try{let m;if(typeof p=="number"||/^\d{4,5}$/.test(String(p).trim())){let x=Number(p),N=new Date(1899,11,30);m=new Date(N.getTime()+x*24*60*60*1e3)}else m=new Date(p);if(!isNaN(m.getTime())&&m.getFullYear()>1970){let x=String(m.getMonth()+1).padStart(2,"0"),N=String(m.getDate()).padStart(2,"0"),L=m.getFullYear();u=`${x}/${N}/${L}`}else console.warn("[JE Draft] Date parsing resulted in invalid date:",p,"->",m),u=String(p)}catch(m){console.warn("[JE Draft] Could not parse TxnDate:",p,m),u=String(p)}let d=g?`${g} PTO Accrual`:"PTO Accrual",y={};if(r.length>1){let m=(r[0]||[]).map(L=>q(L)),x=B(m,["account number","accountnumber","account","acct"]),N=B(m,["account name","accountname","name","description"]);x!==-1&&N!==-1&&r.slice(1).forEach(L=>{let z=String(L[x]||"").trim(),re=String(L[N]||"").trim();z&&(y[z]=re)})}let h=(l[0]||[]).map(m=>q(m));console.log("[JE Draft] PTO_Analysis headers:",h),console.log("[JE Draft] PTO_Analysis row count:",l.length-1);let w=B(h,["department"]),f=B(h,["change"]);if(console.log("[JE Draft] Column indices - Department:",w,"Change:",f),w===-1||f===-1)throw new Error(`Could not find required columns in PTO_Analysis. Found headers: ${h.join(", ")}. Looking for "Department" (found: ${w!==-1}) and "Change" (found: ${f!==-1}).`);let v={},b=0,T=0,C=0;if(l.slice(1).forEach((m,x)=>{b++;let N=String(m[w]||"").trim(),L=m[f],z=Number(L)||0;if(x<3&&console.log(`[JE Draft] Row ${x+2}: Dept="${N}", Change raw="${L}", Change num=${z}`),!N){C++;return}if(z===0){T++;return}v[N]||(v[N]=0),v[N]+=z}),console.log(`[JE Draft] Data summary: ${b} rows, ${T} with zero change, ${C} missing dept`),console.log("[JE Draft] Department totals:",v),Object.keys(v).length===0){let m=`No journal entry lines could be created.

`;throw T===b?(m+=`All 'Change' amounts in PTO_Analysis are $0.00.

`,m+=`Common causes:
`,m+=`\u2022 Missing Pay Rate data (Liability = Balance \xD7 Pay Rate)
`,m+=`\u2022 No prior period data to compare against
`,m+=`\u2022 PTO Analysis hasn't been run yet

`,m+="Please verify Pay Rate values exist in PTO_Analysis."):C===b?(m+=`All rows are missing Department values.

`,m+="Please ensure the 'Department' column is populated in PTO_Analysis."):(m+=`Found ${b} rows but none had both a Department and non-zero Change amount.
`,m+=`\u2022 ${T} rows with zero change
`,m+=`\u2022 ${C} rows missing department`),new Error(m)}let $=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],A=[$],P=0,M=0;Object.entries(v).forEach(([m,x])=>{if(Math.abs(x)<.01)return;let N=m.toLowerCase().trim(),L=Wn[N]||"",z=y[L]||"",re=x>0?Math.abs(x):0,_=x<0?Math.abs(x):0;P+=re,M+=_,A.push([i,u,L,z,x,re,_,d,m])});let k=P-M;if(Math.abs(k)>=.01){let m=k<0?Math.abs(k):0,x=k>0?Math.abs(k):0,N=y[Kt]||"Accrued PTO";A.push([i,u,Kt,N,-k,m,x,d,""])}let E=e.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");if(E.load("isNullObject"),await e.sync(),E.isNullObject)E=e.workbook.worksheets.add("PTO_JE_Draft");else{let m=E.getUsedRangeOrNullObject();m.load("isNullObject"),await e.sync(),m.isNullObject||m.clear()}if(A.length>0){let m=E.getRangeByIndexes(0,0,A.length,$.length);m.values=A;let x=E.getRangeByIndexes(0,0,1,$.length);ot(x);let N=A.length-1;N>0&&(ue(E,4,N,!0),ue(E,5,N),ue(E,6,N)),m.format.autofitColumns()}await e.sync(),E.activate(),E.getRange("A1").select(),await e.sync()}),await ln()}catch(e){console.error("Create JE Draft error:",e),window.alert(`Unable to create Journal Entry: ${e.message}`)}finally{X(!1)}}async function Qn(){if(!oe()){window.alert("Excel is not available. Open this module inside Excel to export.");return}X(!0,"Preparing JE CSV...");try{let{rows:e}=await Excel.run(async n=>{let o=n.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();o.load("values"),await n.sync();let s=o.isNullObject?[]:o.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty.");return{rows:s}}),t=ya(e);ha(`pto-je-draft-${ct()}.csv`,t)}catch(e){console.error("PTO JE export:",e),window.alert("Unable to export the JE draft. Confirm the sheet has data.")}finally{X(!1)}}async function Ve(e){if(!(!e||!oe()))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e);n.activate(),n.getRange("A1").select(),await t.sync()})}catch(t){console.error(t)}}async function Xn(){if(!(!oe()||!window.confirm("This will clear all data in PTO_Data. Are you sure?"))){X(!0);try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("PTO_Data"),a=n.getUsedRangeOrNullObject();a.load("rowCount"),await t.sync(),!a.isNullObject&&a.rowCount>1&&(n.getRangeByIndexes(1,0,a.rowCount-1,20).clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),window.alert("PTO_Data cleared successfully. You can now paste new data.")}catch(t){console.error("Clear PTO_Data error:",t),window.alert(`Failed to clear PTO_Data: ${t.message}`)}finally{X(!1)}}}async function Qt(e){if(!e||!oe())return;let t={SS_Employee_Roster:["Employee","Department","Pay_Rate","Status","Hire_Date"],SS_Chart_of_Accounts:["Account_Number","Account_Name","Type","Category"]};try{await Excel.run(async n=>{let a=n.workbook.worksheets.getItemOrNullObject(e);if(a.load("isNullObject"),await n.sync(),a.isNullObject){a=n.workbook.worksheets.add(e);let o=t[e]||["Column1","Column2"],s=a.getRange(`A1:${String.fromCharCode(64+o.length)}1`);s.values=[o],s.format.font.bold=!0,s.format.fill.color="#f0f0f0",s.format.autofitColumns(),await n.sync()}a.activate(),a.getRange("A1").select(),await n.sync()})}catch(n){console.error("Error opening reference sheet:",n)}}async function Zn(){if(!oe()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(o=>{o.name.toUpperCase().startsWith("SS_")&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Config] Made visible: ${o.name}`),n++)}),await e.sync();let a=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");a.load("isNullObject"),await e.sync(),a.isNullObject||(a.activate(),a.getRange("A1").select(),await e.sync()),console.log(`[Config] ${n} system sheets now visible`)})}catch(e){console.error("[Config] Error unhiding system sheets:",e)}}function ne(e){var n,a;let t=String(e!=null?e:"").trim();return(a=(n=I.values)==null?void 0:n[t])!=null?a:""}function ye(e){var n;if(e)return e;let t=ne(S.reviewerName);if(t)return t;if((n=window.PrairieForge)!=null&&n._sharedConfigCache){let a=window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer")||window.PrairieForge._sharedConfigCache.get("Default_Reviewer");if(a)return a}return""}function te(e,t,n={}){var r;let a=String(e!=null?e:"").trim();if(!a)return;I.values[a]=t!=null?t:"";let o=(r=n.debounceMs)!=null?r:0;if(!o){let l=Oe.get(a);l&&clearTimeout(l),Oe.delete(a),_e(a,t!=null?t:"",pe);return}Oe.has(a)&&clearTimeout(Oe.get(a));let s=setTimeout(()=>{Oe.delete(a),_e(a,t!=null?t:"",pe)},o);Oe.set(a,s)}function q(e){return String(e!=null?e:"").trim().toLowerCase()}function X(e,t="Working..."){let n=document.getElementById(In);n&&(n.style.display="none")}function it(){Tn()}typeof Office!="undefined"&&Office.onReady?Office.onReady(()=>it()).catch(()=>it()):it();function he(e){return I.steps[e]||{notes:"",reviewer:"",signOffDate:""}}function cn(e){return He[e]||{}}function ea(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}async function Z(e,t,n){let a=I.steps[e]||{notes:"",reviewer:"",signOffDate:""};a[t]=n,I.steps[e]=a;let o=cn(e),s=t==="notes"?o.note:t==="reviewer"?o.reviewer:o.signOff;if(s&&K())try{await _e(s,n,pe)}catch(r){console.warn("PTO: unable to save field",s,r)}}async function dn(e,t){I.permanents[e]=t;let n=cn(e);if(n!=null&&n.note&&K())try{await Excel.run(async a=>{var u;let o=a.workbook.tables.getItemOrNullObject(pe[0]);if(await a.sync(),o.isNullObject)return;let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let l=r.values[0]||[],c=l.map(d=>String(d||"").trim().toLowerCase()),i={field:c.findIndex(d=>d==="field"||d==="field name"||d==="setting"),permanent:c.findIndex(d=>d==="permanent"||d==="persist"),value:c.findIndex(d=>d==="value"||d==="setting value"),type:c.findIndex(d=>d==="type"||d==="category"),title:c.findIndex(d=>d==="title"||d==="display name")};if(i.field===-1)return;let g=(s.values||[]).findIndex(d=>String(d[i.field]||"").trim()===n.note);if(g>=0)i.permanent>=0&&(s.getCell(g,i.permanent).values=[[t?"Y":"N"]]);else{let d=new Array(l.length).fill("");i.type>=0&&(d[i.type]="Other"),i.title>=0&&(d[i.title]=""),d[i.field]=n.note,i.permanent>=0&&(d[i.permanent]=t?"Y":"N"),i.value>=0&&(d[i.value]=((u=I.steps[e])==null?void 0:u.notes)||""),o.rows.add(null,[d])}await a.sync()})}catch(a){console.warn("PTO: unable to update permanent flag",a)}}async function un(e,t){let n=sn[e];if(n&&(I.completes[e]=t?"Y":"",!!K()))try{await _e(n,t?"Y":"",pe)}catch(a){console.warn("PTO: unable to save completion flag",n,a)}}function Xt(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function ta(){let e={};return Object.keys(He).forEach(t=>{var s;let n=parseInt(t,10),a=!!((s=I.steps[n])!=null&&s.signOffDate),o=!!I.completes[n];e[n]=a||o}),e}function pn(e,{buttonId:t,inputId:n,canActivate:a=null,onComplete:o=null}){var c;let s=document.getElementById(t);if(!s)return;let r=document.getElementById(n),l=!!((c=I.steps[e])!=null&&c.signOffDate)||!!I.completes[e];Xt(s,l),s.addEventListener("click",()=>{if(!s.classList.contains("is-active")&&e>0){let g=ta(),{canComplete:u,message:d}=Ut(e,g);if(!u){Gt(d);return}}if(typeof a=="function"&&!a())return;let p=!s.classList.contains("is-active");Xt(s,p),r&&(r.value=p?ct():"",Z(e,"signOffDate",r.value)),un(e,p),p&&typeof o=="function"&&o()})}function O(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function na(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function fn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function ke(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function lt(e){if(!e)return null;let t=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(e));if(!t)return null;let n=Number(t[1]),a=Number(t[2]),o=Number(t[3]);return!n||!a||!o?null:{year:n,month:a,day:o}}function Zt(e){if(!e)return"";let t=lt(e);if(!t)return"";let{year:n,month:a,day:o}=t;return`${n}-${String(a).padStart(2,"0")}-${String(o).padStart(2,"0")}`}function aa(e){let t=lt(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function oa(e){let t=lt(e);return t?`PTO-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function ct(){let e=new Date,t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}function sa(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="y"||t==="yes"||t==="true"||t==="t"||t==="1"}function ia(e){if(e instanceof Date)return e.getTime();if(typeof e=="number"){let n=ra(e);return n?n.getTime():null}let t=new Date(e);return Number.isNaN(t.getTime())?null:t.getTime()}function ra(e){if(!Number.isFinite(e))return null;let t=new Date(Date.UTC(1899,11,30));return new Date(t.getTime()+e*24*60*60*1e3)}function la(){let e=n=>{var a,o;return((o=(a=document.getElementById(n))==null?void 0:a.value)==null?void 0:o.trim())||""};[{id:"config-payroll-date",field:S.payrollDate},{id:"config-accounting-period",field:S.accountingPeriod},{id:"config-journal-id",field:S.journalEntryId},{id:"config-company-name",field:S.companyName},{id:"config-payroll-provider",field:S.payrollProvider},{id:"config-accounting-link",field:S.accountingSoftware},{id:"config-user-name",field:S.reviewerName}].forEach(({id:n,field:a})=>{let o=e(n);a&&te(a,o)})}function B(e,t=[]){let n=t.map(a=>q(a));return e.findIndex(a=>n.some(o=>a.includes(o)))}function ca(e){var b,T,C,D,$,A,P,M,k;let t=he(2),n=(t==null?void 0:t.notes)||"",a=!!I.permanents[2],o=ye(t==null?void 0:t.reviewer),s=(t==null?void 0:t.signOffDate)||"",r=!!(ke(I.completes[2])||s),l=j.roster||{},c=j.hasAnalyzed,i=(T=(b=j.roster)==null?void 0:b.difference)!=null?T:0,p=!j.skipAnalysis&&Math.abs(i)>0,g=(C=l.rosterCount)!=null?C:0,u=(D=l.payrollCount)!=null?D:0,d=($=l.difference)!=null?$:u-g,y=Array.isArray(l.mismatches)?l.mismatches.filter(Boolean):[],h="";j.loading?h=((P=(A=window.PrairieForge)==null?void 0:A.renderStatusBanner)==null?void 0:P.call(A,{type:"info",message:"Analyzing headcount\u2026",escapeHtml:O}))||"":j.lastError&&(h=((k=(M=window.PrairieForge)==null?void 0:M.renderStatusBanner)==null?void 0:k.call(M,{type:"error",message:j.lastError,escapeHtml:O}))||"");let w=(E,m,x,N)=>{let L=!c,z;L?z='<span class="pf-je-check-circle pf-je-circle--pending"></span>':N?z=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:z=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let re=c?` = ${x}`:"";return`
            <div class="pf-je-check-row">
                ${z}
                <span class="pf-je-check-desc-pill">${O(E)}${re}</span>
            </div>
        `},f=`
        ${w("SS_Employee_Roster count","Active employees in roster",g,!0)}
        ${w("PTO_Data count","Unique employees in PTO data",u,!0)}
        ${w("Difference","Should be zero",d,d===0)}
    `,v=y.length&&!j.skipAnalysis&&c?window.PrairieForge.renderMismatchTiles({mismatches:y,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"PTO Data",escapeHtml:O}):"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${O(fe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${O(e.title)}</h2>
            <p class="pf-hero-copy">${O(e.summary||"")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${j.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
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
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${Le}</button>`,"Run")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${Ie}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${h}
                <div class="pf-je-checks-container">
                    ${f}
                </div>
                ${v}
            </article>
            ${ce({textareaId:"step-notes-input",value:n,permanentId:"step-notes-lock-2",isPermanent:a,hintId:p?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
            ${de({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:r,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
        </section>
    `}function da(){let e=Q.completenessCheck||{},t=Q.missingPayRates||[],n=[{key:"accrualRate",label:"Accrual Rate",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"carryOver",label:"Carry Over",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdAccrued",label:"YTD Accrued",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdUsed",label:"YTD Used",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"balance",label:"Balance",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"}],o=n.every(i=>e[i.key]!==null&&e[i.key]!==void 0)&&n.every(i=>{var p;return(p=e[i.key])==null?void 0:p.match}),s=t.length>0,r=i=>{let p=e[i.key],g=p==null,u;return g?u='<span class="pf-je-check-circle pf-je-circle--pending"></span>':p.match?u=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:u=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${u}
                <span class="pf-je-check-desc-pill">${O(i.label)}: ${O(i.desc)}</span>
            </div>
        `},l=n.map(i=>r(i)).join(""),c="";if(s){let i=t[0],p=t.length-1;c=`
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${O(i.name)}</strong> to calculate liability
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
                ${p>0?`<p class="pf-readiness-remaining">${p} more employee${p>1?"s":""} need pay rates</p>`:""}
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
    `}function ua(e){var d,y,h,w,f,v,b,T;let t=he(3),n=!!I.permanents[3],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(ke(I.completes[3])||o),r=W.hasRun,{balanceIssues:l,zeroBalances:c,accrualOutliers:i,totalEmployees:p}=W,g="";if(W.loading)g=((y=(d=window.PrairieForge)==null?void 0:d.renderStatusBanner)==null?void 0:y.call(d,{type:"info",message:"Analyzing data quality...",escapeHtml:O}))||"";else if(r){let C=l.length,D=i.length+c.length;C>0?g=((w=(h=window.PrairieForge)==null?void 0:h.renderStatusBanner)==null?void 0:w.call(h,{type:"error",title:`${C} Balance Issue${C>1?"s":""} Found`,message:"Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",escapeHtml:O}))||"":D>0?g=((v=(f=window.PrairieForge)==null?void 0:f.renderStatusBanner)==null?void 0:v.call(f,{type:"warning",title:"No Critical Issues",message:`${D} informational item${D>1?"s":""} to review (see below).`,escapeHtml:O}))||"":g=((T=(b=window.PrairieForge)==null?void 0:b.renderStatusBanner)==null?void 0:T.call(b,{type:"success",title:"Data Quality Passed",message:`${p} employee${p!==1?"s":""} checked \u2014 no anomalies found.`,escapeHtml:O}))||""}let u=[];return r&&l.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u26A0\uFE0F</span>
                    <span class="pf-quality-issue-title">Balance Issues (${l.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${l.slice(0,5).map(C=>`<li><strong>${O(C.name)}</strong>: ${O(C.issue)}</li>`).join("")}
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
                    ${i.slice(0,5).map(C=>`<li><strong>${O(C.name)}</strong>: ${C.accrualRate.toFixed(2)} hrs/period</li>`).join("")}
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
                    ${c.slice(0,5).map(C=>`<li><strong>${O(C.name)}</strong></li>`).join("")}
                    ${c.length>5?`<li class="pf-quality-more">+${c.length-5} more</li>`:""}
                </ul>
            </div>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${O(fe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${O(e.title)}</h2>
            <p class="pf-hero-copy">${O(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${g}
                <div class="pf-signoff-action">
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${Le}</button>`,"Run")}
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
                            ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${Ie}</button>`,"Refresh")}
                            ${W.acknowledged?"":J(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${Pe}</button>`,"Continue")}
                        </div>
                    </div>
                </article>
            `:""}
            ${ce({textareaId:"step-notes-3",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-3",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-3"})}
            ${de({reviewerInputId:"step-reviewer-3",reviewerValue:a,signoffInputId:"step-signoff-3",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-3",completeButtonId:"step-signoff-toggle-3"})}
        </section>
    `}function pa(e){let t=he(4),n=!!I.permanents[4],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(ke(I.completes[4])||o);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${O(fe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${O(e.title)}</h2>
            <p class="pf-hero-copy">${O(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Analysis</h3>
                    <p class="pf-config-subtext">Calculate liabilities and compare against last period.</p>
                </div>
                <div class="pf-signoff-action">
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${Le}</button>`,"Run")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${Ie}</button>`,"Refresh")}
                </div>
            </article>
            ${da()}
            ${ce({textareaId:"step-notes-4",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-4",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-4"})}
            ${de({reviewerInputId:"step-reviewer-4",reviewerValue:a,signoffInputId:"step-signoff-4",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"step-signoff-toggle-4"})}
        </section>
    `}function fa(e){let t=he(5),n=!!I.permanents[5],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(ke(I.completes[5])||o),r=Y.lastError?`<p class="pf-step-note">${O(Y.lastError)}</p>`:"",l=Y.validationRun,c=Y.issues||[],i=[{key:"Debits = Credits",desc:"\u2211 Debit column = \u2211 Credit column"},{key:"Line Amounts Sum to Zero",desc:"\u2211 Line Amount = $0.00"},{key:"JE Matches Analysis Total",desc:"\u2211 Expense line amounts = \u2211 PTO_Analysis Change"}],p=y=>{let h=c.find(v=>v.check===y.key),w=!l,f;return w?f='<span class="pf-je-check-circle pf-je-circle--pending"></span>':h!=null&&h.passed?f=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:f=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${f}
                <span class="pf-je-check-desc-pill">${O(y.desc)}</span>
            </div>
        `},g=i.map(y=>p(y)).join(""),u=c.filter(y=>!y.passed),d="";return l&&u.length>0&&(d=`
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>\u26A0\uFE0F Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${u.map(y=>`<li><strong>${O(y.check)}:</strong> ${O(y.detail)}</li>`).join("")}
                </ul>
            </article>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${O(fe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${O(e.title)}</h2>
            <p class="pf-hero-copy">${O(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${Qe}</button>`,"Generate")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${Ie}</button>`,"Refresh")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${Lt}</button>`,"Export")}
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
            ${ce({textareaId:"step-notes-5",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-5",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-5"})}
            ${de({reviewerInputId:"step-reviewer-5",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
        </section>
    `}function ga(){var t,n;return Math.abs((n=(t=j.roster)==null?void 0:t.difference)!=null?n:0)>0}function dt(){return!j.skipAnalysis&&ga()}async function rt(){if(!K()){j.loading=!1,j.lastError="Excel runtime is unavailable.",ae();return}j.loading=!0,j.lastError=null,se(document.getElementById("headcount-save-btn"),!1),ae();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),a=t.workbook.worksheets.getItem("PTO_Data"),o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.getUsedRangeOrNullObject(),r=a.getUsedRangeOrNullObject();s.load("values"),r.load("values"),o.load("isNullObject"),await t.sync();let l=null;o.isNullObject||(l=o.getUsedRangeOrNullObject(),l.load("values")),await t.sync();let c=s.isNullObject?[]:s.values||[],i=r.isNullObject?[]:r.values||[],p=l&&!l.isNullObject?l.values||[]:[],g=p.length?p:i;return ma(c,g)});j.roster=e.roster,j.hasAnalyzed=!0,j.lastError=null}catch(e){console.warn("PTO headcount: unable to analyze data",e),j.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{j.loading=!1,ae()}}function en(e){if(!e)return!0;let t=e.toLowerCase().trim();return t?["total","subtotal","sum","count","grand","average","avg"].some(a=>t.includes(a)):!0}function ma(e,t){let n={rosterCount:0,payrollCount:0,difference:0,mismatches:[]};if(((e==null?void 0:e.length)||0)<2||((t==null?void 0:t.length)||0)<2)return console.warn("Headcount: insufficient data rows",{rosterRows:(e==null?void 0:e.length)||0,payrollRows:(t==null?void 0:t.length)||0}),{roster:n};let a=tn(e),o=tn(t),s=a.headers,r=o.headers,l={employee:nn(s),termination:s.findIndex(d=>d.includes("termination"))},c={employee:nn(r)};console.log("Headcount column detection:",{rosterEmployeeCol:l.employee,rosterTerminationCol:l.termination,payrollEmployeeCol:c.employee,rosterHeaders:s.slice(0,5),payrollHeaders:r.slice(0,5)});let i=new Set,p=new Set;for(let d=a.startIndex;d<e.length;d+=1){let y=e[d],h=l.employee>=0?me(y[l.employee]):"";en(h)||l.termination>=0&&me(y[l.termination])||i.add(h.toLowerCase())}for(let d=o.startIndex;d<t.length;d+=1){let y=t[d],h=c.employee>=0?me(y[c.employee]):"";en(h)||p.add(h.toLowerCase())}n.rosterCount=i.size,n.payrollCount=p.size,n.difference=n.payrollCount-n.rosterCount,console.log("Headcount results:",{rosterCount:n.rosterCount,payrollCount:n.payrollCount,difference:n.difference});let g=[...i].filter(d=>!p.has(d)),u=[...p].filter(d=>!i.has(d));return n.mismatches=[...g.map(d=>`In roster, missing in PTO_Data: ${d}`),...u.map(d=>`In PTO_Data, missing in roster: ${d}`)],{roster:n}}function tn(e){if(!Array.isArray(e)||!e.length)return{headers:[],startIndex:1};let t=e.findIndex((o=[])=>o.some(s=>me(s).toLowerCase().includes("employee"))),n=t===-1?0:t;return{headers:(e[n]||[]).map(o=>me(o).toLowerCase()),startIndex:n+1}}function nn(e=[]){let t=-1,n=-1;return e.forEach((a,o)=>{let s=a.toLowerCase();if(!s.includes("employee"))return;let r=1;s.includes("name")?r=4:s.includes("id")?r=2:r=3,r>n&&(n=r,t=o)}),t}function me(e){return e==null?"":String(e).trim()}async function gn(e=null){let t=async n=>{let a=n.workbook.worksheets.getItem("PTO_Data"),o=n.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster"),r=n.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary"),l=n.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary"),c=a.getUsedRangeOrNullObject();c.load("values"),o.load("isNullObject"),s.load("isNullObject"),r.load("isNullObject"),l.load("isNullObject"),await n.sync();let i=c.isNullObject?[]:c.values||[];if(!i.length)return;let p=(i[0]||[]).map(_=>q(_)),g=p.findIndex(_=>_.includes("employee")&&_.includes("name")),u=g>=0?g:0,d=B(p,["accrual rate"]),y=B(p,["carry over","carryover"]),h=p.findIndex(_=>_.includes("ytd")&&(_.includes("accrued")||_.includes("accrual"))),w=p.findIndex(_=>_.includes("ytd")&&_.includes("used")),f=B(p,["balance","current balance","pto balance"]);console.log("[PTO Analysis] PTO_Data headers:",p),console.log("[PTO Analysis] Column indices found:",{employee:u,accrualRate:d,carryOver:y,ytdAccrued:h,ytdUsed:w,balance:f}),w>=0?console.log(`[PTO Analysis] YTD Used column: "${p[w]}" at index ${w}`):console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:",p);let v=i.slice(1).map(_=>me(_[u])).filter(_=>_&&!_.toLowerCase().includes("total")),b=new Map;i.slice(1).forEach(_=>{let G=q(_[u]);!G||G.includes("total")||b.set(G,_)});let T=new Map;if(s.isNullObject)console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");else{let _=s.getUsedRangeOrNullObject();_.load("values"),await n.sync();let G=_.isNullObject?[]:_.values||[];if(G.length){let H=(G[0]||[]).map(R=>q(R));console.log("[PTO Analysis] SS_Employee_Roster headers:",H);let F=H.findIndex(R=>R.includes("employee")&&R.includes("name"));F<0&&(F=H.findIndex(R=>R==="employee"||R==="name"||R==="full name"));let U=H.findIndex(R=>R.includes("department"));console.log(`[PTO Analysis] Roster column indices - Name: ${F}, Dept: ${U}`),F>=0&&U>=0?(G.slice(1).forEach(R=>{let ee=q(R[F]),le=me(R[U]);ee&&T.set(ee,le)}),console.log(`[PTO Analysis] Built roster map with ${T.size} employees`)):console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster")}}let C=new Map;if(!r.isNullObject){let _=r.getUsedRangeOrNullObject();_.load("values"),await n.sync();let G=_.isNullObject?[]:_.values||[];if(G.length){let H=(G[0]||[]).map(U=>q(U)),F={payrollDate:B(H,["payroll date"]),employee:B(H,["employee"]),category:B(H,["payroll category","category"]),amount:B(H,["amount","gross salary","gross_salary","earnings"])};F.employee>=0&&F.category>=0&&F.amount>=0&&G.slice(1).forEach(U=>{let R=q(U[F.employee]);if(!R)return;let ee=q(U[F.category]);if(!ee.includes("regular")||!ee.includes("earn"))return;let le=Number(U[F.amount])||0;if(!le)return;let Se=ia(U[F.payrollDate]),Ee=C.get(R);(!Ee||Se!=null&&Se>Ee.timestamp)&&C.set(R,{payRate:le/80,timestamp:Se})})}}let D=new Map;if(!l.isNullObject){let _=l.getUsedRangeOrNullObject();_.load("values"),await n.sync();let G=_.isNullObject?[]:_.values||[];if(G.length>1){let H=(G[0]||[]).map(R=>q(R)),F=H.findIndex(R=>R.includes("employee")&&R.includes("name")),U=B(H,["liability amount","liability","accrued pto"]);F>=0&&U>=0&&G.slice(1).forEach(R=>{let ee=q(R[F]);if(!ee)return;let le=Number(R[U])||0;D.set(ee,le)})}}let $=ne(S.payrollDate)||"",A=[],P=[],M=v.map((_,G)=>{var ft,gt,mt,yt,ht,vt,bt;let H=q(_),F=T.get(H)||"",U=(gt=(ft=C.get(H))==null?void 0:ft.payRate)!=null?gt:"",R=b.get(H),ee=R&&d>=0&&(mt=R[d])!=null?mt:"",le=R&&y>=0&&(yt=R[y])!=null?yt:"",Se=R&&h>=0&&(ht=R[h])!=null?ht:"",Ee=R&&w>=0&&(vt=R[w])!=null?vt:"";(H.includes("avalos")||H.includes("sarah"))&&console.log(`[PTO Debug] ${_}:`,{ytdUsedIdx:w,rawValue:R?R[w]:"no dataRow",ytdUsed:Ee,fullRow:R});let Ue=R&&f>=0&&Number(R[f])||0,ut=G+2;!U&&typeof U!="number"&&A.push({name:_,rowIndex:ut}),F||P.push({name:_,rowIndex:ut});let Ge=typeof U=="number"&&Ue?Ue*U:0,pt=(bt=D.get(H))!=null?bt:0,mn=(typeof Ge=="number"?Ge:0)-pt;return[$,_,F,U,ee,le,Se,Ee,Ue,Ge,pt,mn]});Q.missingPayRates=A.filter(_=>!Q.ignoredMissingPayRates.has(_.name)),Q.missingDepartments=P,console.log(`[PTO Analysis] Data quality: ${A.length} missing pay rates, ${P.length} missing departments`);let k=[["Analysis Date","Employee Name","Department","Pay Rate","Accrual Rate","Carry Over","YTD Accrued","YTD Used","Balance","Liability Amount","Accrued PTO $ [Prior Period]","Change"],...M],E=o.isNullObject?n.workbook.worksheets.add("PTO_Analysis"):o,m=E.getUsedRangeOrNullObject();m.load("address"),await n.sync(),m.isNullObject||m.clear();let x=k[0].length,N=k.length,L=M.length,z=E.getRangeByIndexes(0,0,N,x);z.values=k;let re=E.getRangeByIndexes(0,0,1,x);ot(re),L>0&&(Jt(E,0,L),ue(E,3,L),we(E,4,L),we(E,5,L),we(E,6,L),we(E,7,L),we(E,8,L),ue(E,9,L),ue(E,10,L),ue(E,11,L,!0)),z.format.autofitColumns(),E.getRange("A1").select(),await n.sync()};K()&&(e?await t(e):await Excel.run(t))}function ya(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let a=String(n);return/[",\n]/.test(a)?`"${a.replace(/"/g,'""')}"`:a}).join(",")).join(`
`)}function ha(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),a=URL.createObjectURL(n),o=document.createElement("a");o.href=a,o.download=e,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(a),1e3)}function an(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=dt(),n=document.getElementById("step-notes-input"),a=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!a;let o=document.getElementById("headcount-notes-hint");o&&(o.textContent=t?"Please document outstanding differences before signing off.":"")}function on(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(ge)?t.slice(ge.length).replace(/^\s+/,""):t.replace(new RegExp(`^${ge}\\s*`,"i"),"").trimStart(),a=ge+(n?`
${n}`:"");e.value!==a&&(e.value=a),Z(2,"notes",e.value)}function va(){let e=document.getElementById("step-notes-input");e&&e.addEventListener("input",()=>{if(!j.skipAnalysis)return;let t=e.value||"";if(!t.startsWith(ge)){let n=t.replace(ge,"").trimStart();e.value=ge+(n?`
${n}`:"")}Z(2,"notes",e.value)})}function ba(){var n;let e=dt(),t=((n=document.getElementById("step-notes-input"))==null?void 0:n.value.trim())||"";if(e&&!t){window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");return}}})();
//# sourceMappingURL=app.bundle.js.map
