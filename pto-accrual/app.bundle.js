/* Prairie Forge PTO Accrual */
(()=>{function K(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}var Je="SS_PF_Config";async function bt(e,t=[Je]){var o;let n=e.workbook.tables;n.load("items/name"),await e.sync();let a=(o=n.items)==null?void 0:o.find(s=>t.includes(s.name));return a?e.workbook.tables.getItem(a.name):(console.warn("Config table not found. Looking for:",t),null)}function wt(e){let t=e.map(n=>String(n||"").trim().toLowerCase());return{field:t.findIndex(n=>n==="field"||n==="field name"||n==="setting"),value:t.findIndex(n=>n==="value"||n==="setting value"),type:t.findIndex(n=>n==="type"||n==="category"),title:t.findIndex(n=>n==="title"||n==="display name"),permanent:t.findIndex(n=>n==="permanent"||n==="persist")}}async function kt(e=[Je]){if(!K())return{};try{return await Excel.run(async t=>{let n=await bt(t,e);if(!n)return{};let a=n.getDataBodyRange(),o=n.getHeaderRowRange();a.load("values"),o.load("values"),await t.sync();let s=o.values[0]||[],r=wt(s);if(r.field===-1||r.value===-1)return console.warn("Config table missing FIELD or VALUE columns. Headers:",s),{};let l={};return(a.values||[]).forEach(i=>{var g;let p=String(i[r.field]||"").trim();p&&(l[p]=(g=i[r.value])!=null?g:"")}),console.log("Configuration loaded:",Object.keys(l).length,"fields"),l})}catch(t){return console.error("Failed to load configuration:",t),{}}}async function Ce(e,t,n=[Je]){if(!K())return!1;try{return await Excel.run(async a=>{let o=await bt(a,n);if(!o){console.warn("Config table not found for write");return}let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let l=r.values[0]||[],c=wt(l);if(c.field===-1||c.value===-1){console.error("Config table missing FIELD or VALUE columns");return}let p=(s.values||[]).findIndex(g=>String(g[c.field]||"").trim()===e);if(p>=0)s.getCell(p,c.value).values=[[t]];else{let g=new Array(l.length).fill("");c.type>=0&&(g[c.type]="Run Settings"),g[c.field]=e,g[c.value]=t,c.permanent>=0&&(g[c.permanent]="N"),c.title>=0&&(g[c.title]=""),o.rows.add(null,[g]),console.log("Added new config row:",e,"=",t)}await a.sync(),console.log("Saved config:",e,"=",t)}),!0}catch(a){return console.error("Failed to save config:",e,a),!1}}var bn="SS_PF_Config",wn="module-prefix",ze="system",be={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function Ot(){if(!K())return{...be};try{return await Excel.run(async e=>{var p,g;let t=e.workbook.worksheets.getItemOrNullObject(bn);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...be};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((p=n.values)!=null&&p.length))return{...be};let a=n.values,o=Sn(a[0]),s=o.get("category"),r=o.get("field"),l=o.get("value");if(s===void 0||r===void 0||l===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...be};let c={},i=!1;for(let u=1;u<a.length;u++){let d=a[u];if($e(d[s])===wn){let y=String((g=d[r])!=null?g:"").trim().toUpperCase(),b=$e(d[l]);y&&b&&(c[y]=b,i=!0)}}return i?(console.log("[Tab Visibility] Loaded prefix config:",c),c):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...be})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...be}}}async function qe(e){if(!K())return;let t=$e(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await Ot();await Excel.run(async a=>{let o=a.workbook.worksheets;o.load("items/name,visibility"),await a.sync();let s={};for(let[u,d]of Object.entries(n))s[d]||(s[d]=[]),s[d].push(u);let r=s[t]||[],l=s[ze]||[],c=[];for(let[u,d]of Object.entries(s))u!==t&&u!==ze&&c.push(...d);console.log(`[Tab Visibility] Active prefixes: ${r.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${c.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${l.join(", ")}`);let i=[],p=[];o.items.forEach(u=>{let d=u.name,h=d.toUpperCase(),y=r.some(v=>h.startsWith(v)),b=c.some(v=>h.startsWith(v)),f=l.some(v=>h.startsWith(v));y?(i.push(u),console.log(`[Tab Visibility] SHOW: ${d} (matches active module prefix)`)):f?(p.push(u),console.log(`[Tab Visibility] HIDE: ${d} (system sheet)`)):b?(p.push(u),console.log(`[Tab Visibility] HIDE: ${d} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${d} (no prefix match, leaving as-is)`)});for(let u of i)u.visibility=Excel.SheetVisibility.visible;if(await a.sync(),o.items.filter(u=>u.visibility===Excel.SheetVisibility.visible).length>p.length){for(let u of p)try{u.visibility=Excel.SheetVisibility.hidden}catch(d){console.warn(`[Tab Visibility] Could not hide "${u.name}":`,d.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${i.length}, hid ${p.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function kn(){if(!K()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.visibility!==Excel.SheetVisibility.visible&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${a.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function On(){if(!K()){console.log("Excel not available");return}try{let e=await Ot(),t=[];for(let[n,a]of Object.entries(e))a===ze&&t.push(n);await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name,visibility"),await n.sync(),a.items.forEach(o=>{let s=o.name.toUpperCase();t.some(r=>s.startsWith(r))&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function Sn(e=[]){let t=new Map;return e.forEach((n,a)=>{let o=$e(n);o&&t.set(o,a)}),t}function $e(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=kn,window.PrairieForge.unhideSystemSheets=On,window.PrairieForge.applyModuleTabVisibility=qe);var St={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var Ct=St.ADA_IMAGE_URL;async function je(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async a=>{let o=a.workbook.worksheets.getItemOrNullObject(e);o.load("isNullObject, name"),await a.sync();let s;o.isNullObject?(s=a.workbook.worksheets.add(e),await a.sync(),await Et(a,s,t,n)):(s=o,await Et(a,s,t,n)),s.activate(),s.getRange("A1").select(),await a.sync()})}catch(a){console.error(`Error activating homepage sheet ${e}:`,a)}}async function Et(e,t,n,a){try{let i=t.getUsedRangeOrNullObject();i.load("isNullObject"),await e.sync(),i.isNullObject||(i.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let o=[[n,""],[a,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=o;let r=t.getRange("A1:Z100");r.format.fill.color="#0f0f0f";let l=t.getRange("A1");l.format.font.bold=!0,l.format.font.size=36,l.format.font.color="#ffffff",l.format.font.name="Segoe UI Light",l.format.verticalAlignment="Center";let c=t.getRange("A2");c.format.font.size=14,c.format.font.color="#a0a0a0",c.format.font.name="Segoe UI",c.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var xt={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function Le(e){return xt[e]||xt["module-selector"]}function _t(){We();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${Ct}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",En),e}function We(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function En(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${Ct}" alt="Ada" />
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",Ye),t.addEventListener("click",o=>{o.target===t&&Ye()});let a=o=>{o.key==="Escape"&&(Ye(),document.removeEventListener("keydown",a))};document.addEventListener("keydown",a)}function Ye(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var xn=["January","February","March","April","May","June","July","August","September","October","November","December"],Pt=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],Cn=["Su","Mo","Tu","We","Th","Fr","Sa"],we=null;function Tt(e,t={}){let n=document.getElementById(e);if(!n)return;let{onChange:a=null,minDate:o=null,maxDate:s=null,readonly:r=!1}=t,l=n.closest(".pf-datepicker-wrapper");l||(l=document.createElement("div"),l.className="pf-datepicker-wrapper",n.parentNode.insertBefore(l,n),l.appendChild(n)),n.type="text",n.placeholder="YYYY-MM-DD or click calendar",n.classList.add("pf-datepicker-input");let c=n.value?Be(n.value):null,i=c?new Date(c):new Date;c&&(n.value=Ke(c));let p=document.createElement("span");p.className="pf-datepicker-icon",p.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>',l.appendChild(p);let g=document.createElement("div");g.className="pf-datepicker-dropdown",g.id=`${e}-dropdown`,l.appendChild(g);function u(){var w,R,_,$,j,D,V;let f=i.getFullYear(),v=i.getMonth();g.innerHTML=`
            <div class="pf-datepicker-header">
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev-year" title="Previous Year">\xAB</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev" title="Previous Month">\u2039</button>
                <span class="pf-datepicker-title">${xn[v]} ${f}</span>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next" title="Next Month">\u203A</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next-year" title="Next Year">\xBB</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-close" title="Close">\xD7</button>
            </div>
            <div class="pf-datepicker-weekdays">
                ${Cn.map(S=>`<span>${S}</span>`).join("")}
            </div>
            <div class="pf-datepicker-days">
                ${d(f,v,c)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-today">Today</button>
                <button type="button" class="pf-datepicker-clear">Clear</button>
            </div>
        `,(w=g.querySelector(".pf-datepicker-prev-year"))==null||w.addEventListener("mousedown",S=>{S.preventDefault(),S.stopPropagation(),i.setFullYear(i.getFullYear()-1),u()}),(R=g.querySelector(".pf-datepicker-prev"))==null||R.addEventListener("mousedown",S=>{S.preventDefault(),S.stopPropagation(),i.setMonth(i.getMonth()-1),u()}),(_=g.querySelector(".pf-datepicker-next"))==null||_.addEventListener("mousedown",S=>{S.preventDefault(),S.stopPropagation(),i.setMonth(i.getMonth()+1),u()}),($=g.querySelector(".pf-datepicker-next-year"))==null||$.addEventListener("mousedown",S=>{S.preventDefault(),S.stopPropagation(),i.setFullYear(i.getFullYear()+1),u()}),(j=g.querySelector(".pf-datepicker-close"))==null||j.addEventListener("mousedown",S=>{S.preventDefault(),S.stopPropagation(),b()}),g.querySelectorAll(".pf-datepicker-day:not(.disabled)").forEach(S=>{S.addEventListener("mousedown",I=>{I.preventDefault(),I.stopPropagation();let O=parseInt(S.dataset.day),m=parseInt(S.dataset.month),P=parseInt(S.dataset.year);h(new Date(P,m,O))})}),(D=g.querySelector(".pf-datepicker-today"))==null||D.addEventListener("mousedown",S=>{S.preventDefault(),S.stopPropagation(),h(new Date)}),(V=g.querySelector(".pf-datepicker-clear"))==null||V.addEventListener("mousedown",S=>{S.preventDefault(),S.stopPropagation(),h(null)})}function d(f,v,w){let R=new Date(f,v,1).getDay(),_=new Date(f,v+1,0).getDate(),$=new Date(f,v,0).getDate(),j=new Date;j.setHours(0,0,0,0);let D="";for(let O=R-1;O>=0;O--){let m=$-O,P=v===0?11:v-1,A=v===0?f-1:f;D+=`<span class="pf-datepicker-day other-month" data-day="${m}" data-month="${P}" data-year="${A}">${m}</span>`}for(let O=1;O<=_;O++){let m=new Date(f,v,O),P=m.getTime()===j.getTime(),A=w&&m.getTime()===w.getTime(),N="pf-datepicker-day";P&&(N+=" today"),A&&(N+=" selected"),o&&m<o&&(N+=" disabled"),s&&m>s&&(N+=" disabled"),D+=`<span class="${N}" data-day="${O}" data-month="${v}" data-year="${f}">${O}</span>`}let V=42,S=R+_,I=V-S;for(let O=1;O<=I;O++){let m=v===11?0:v+1,P=v===11?f+1:f;D+=`<span class="pf-datepicker-day other-month" data-day="${O}" data-month="${m}" data-year="${P}">${O}</span>`}return D}function h(f){c=f,f?(n.value=Ke(f),n.dataset.value=_e(f),i=new Date(f)):(n.value="",n.dataset.value=""),b(),a&&a(f?_e(f):""),n.dispatchEvent(new Event("change",{bubbles:!0}))}function y(){if(!r){if(we&&we!==e){let f=document.getElementById(`${we}-dropdown`);f==null||f.classList.remove("open")}we=e,u(),g.classList.add("open"),l.classList.add("open")}}function b(){g.classList.remove("open"),l.classList.remove("open"),we===e&&(we=null)}return n.addEventListener("blur",f=>{if(g.classList.contains("open"))return;let v=n.value.trim();if(!v)return;let w=Be(v);w&&(c=w,n.value=Ke(w),n.dataset.value=_e(w),i=new Date(w),a&&a(_e(w)),n.dispatchEvent(new Event("change",{bubbles:!0})))}),n.addEventListener("keydown",f=>{if(f.key==="Enter"){f.preventDefault();let v=n.value.trim(),w=Be(v);w&&h(w),b()}}),n.addEventListener("click",f=>{f.stopPropagation(),g.classList.contains("open")||y()}),p.addEventListener("click",f=>{f.stopPropagation(),g.classList.contains("open")?b():y()}),document.addEventListener("click",f=>{l.contains(f.target)||b()}),g.addEventListener("click",f=>{f.stopPropagation()}),document.addEventListener("keydown",f=>{f.key==="Escape"&&b()}),{getValue:()=>c?_e(c):"",setValue:f=>{let v=Be(f);h(v)},open:y,close:b}}function Be(e){if(!e)return null;if(/^\d{4}-\d{2}-\d{2}$/.test(e)){let[a,o,s]=e.split("-").map(Number);return new Date(a,o-1,s)}let t=e.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);if(t){let a=Pt.findIndex(o=>o.toLowerCase()===t[1].toLowerCase().substring(0,3));if(a>=0)return new Date(parseInt(t[3]),a,parseInt(t[2]))}if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(e)){let[a,o,s]=e.split("/").map(Number);return new Date(s,a-1,o)}let n=new Date(e);return isNaN(n.getTime())?null:n}function Ke(e){return e?`${Pt[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`:""}function _e(e){if(!e)return"";let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}var It=`
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
`.trim(),Rt=`
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
        <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/>
        <circle cx="9" cy="7" r="4"/>
        <path d="M22 21v-2a4 4 0 0 0-3-3.87"/>
        <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
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
        <path d="M4 19.5v-15A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1 0-5H20"/>
        <path d="M8 7h6"/>
        <path d="M8 11h8"/>
    </svg>
`.trim(),_n={config:`
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
    `};function $t(e){return e&&_n[e]||""}var Xe=`
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
`.trim(),Me=`
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
`.trim(),ja=`
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
`.trim(),jt=`
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
        <circle cx="12" cy="12" r="10" />
        <path d="m15 9-6 6" />
        <path d="m9 9 6 6" />
    </svg>
`.trim(),Bt=`
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
        <path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" />
        <path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7" />
        <path d="M7 3v4a1 1 0 0 0 1 1h7" />
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
        <path d="m12 5 7 7-7 7" />
        <path d="M5 12h14" />
    </svg>
`.trim(),La=`
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
`.trim(),Ba=`
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
`.trim(),Ma=`
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
`.trim(),Va=`
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
`.trim(),Ha=`
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
`.trim(),Fa=`
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
`.trim(),Ua=`
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
`.trim(),Ga=`
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
        <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/>
        <path d="M21 3v5h-5"/>
        <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/>
        <path d="M3 21v-5h5"/>
    </svg>
`.trim(),Ht=`
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
`.trim();function Ie(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function J(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function de({textareaId:e,value:t,permanentId:n,isPermanent:a,hintId:o,saveButtonId:s,isSaved:r=!1,placeholder:l="Enter notes here..."}){let c=a?Ze:Xe,i=s?`<button type="button" class="pf-action-toggle pf-save-btn ${r?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Mt}</button>`:"",p=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${a?"is-locked":""}" id="${n}" aria-pressed="${a}" title="Lock notes (retain after archive)">${c}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Ie(l)}">${Ie(t||"")}</textarea>
                ${o?`<p class="pf-signoff-hint" id="${o}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?J(p,"Lock"):""}
                ${s?J(i,"Save"):""}
            </div>
        </article>
    `}function ue({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:a,isComplete:o,saveButtonId:s,isSaved:r=!1,completeButtonId:l,subtext:c="Sign-off below. Click checkmark icon. Done."}){let i=`<button type="button" class="pf-action-toggle ${o?"is-active":""}" id="${l}" aria-pressed="${!!o}" title="Mark step complete">${Pe}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${Ie(c)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${Ie(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${Ie(a)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${J(i,"Done")}
            </div>
        </article>
    `}function tt(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?Ze:Xe)}function re(e,t){e&&e.classList.toggle("is-saved",t)}function nt(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(a=>{let o=a.getAttribute("data-save-input"),s=document.getElementById(o);if(!s)return;let r=()=>{re(a,!1)};s.addEventListener("input",r),n.push(()=>s.removeEventListener("input",r))}),()=>n.forEach(a=>a())}function Ft(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function Ut(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var at={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},Ve={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function ot(e){e.format.fill.color=at.fillColor,e.format.font.color=at.fontColor,e.format.font.bold=at.bold}function pe(e,t,n,a=!1){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a?Ve.currencyWithNegative:Ve.currency]]}function ke(e,t,n){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[Ve.number]]}function Gt(e,t,n,a=Ve.date){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a]]}var Pn="1.1.0",Ae="pto-accrual";var ge="PTO Accrual";function X(e,t="info",n=4e3){document.querySelectorAll(".pf-toast").forEach(o=>o.remove());let a=document.createElement("div");if(a.className=`pf-toast pf-toast--${t}`,a.innerHTML=`
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
        `,document.head.appendChild(o)}return document.body.appendChild(a),n>0&&setTimeout(()=>a.remove(),n),a}function Tn(e,t={}){let{title:n="Confirm Action",confirmText:a="Continue",cancelText:o="Cancel",icon:s="\u{1F4CB}",destructive:r=!1}=t;return new Promise(l=>{document.querySelectorAll(".pf-confirm-overlay").forEach(i=>i.remove());let c=document.createElement("div");if(c.className="pf-confirm-overlay",c.innerHTML=`
            <div class="pf-confirm-dialog">
                <div class="pf-confirm-icon">${s}</div>
                <div class="pf-confirm-title">${n}</div>
                <div class="pf-confirm-message">${e.replace(/\n/g,"<br>")}</div>
                <div class="pf-confirm-buttons">
                    <button class="pf-confirm-btn pf-confirm-btn--cancel">${o}</button>
                    <button class="pf-confirm-btn pf-confirm-btn--ok ${r?"pf-confirm-btn--destructive":""}">${a}</button>
                </div>
            </div>
        `,!document.getElementById("pf-confirm-styles")){let i=document.createElement("style");i.id="pf-confirm-styles",i.textContent=`
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
            `,document.head.appendChild(i)}document.body.appendChild(c),c.addEventListener("click",i=>{i.target===c&&(c.remove(),l(!1))}),c.querySelector(".pf-confirm-btn--cancel").onclick=()=>{c.remove(),l(!1)},c.querySelector(".pf-confirm-btn--ok").onclick=()=>{c.remove(),l(!0)}})}var In="Calculate your PTO liability, compare against last period, and generate a balanced journal entry\u2014all without leaving Excel.",Rn="../module-selector/index.html",An="pf-loader-overlay",fe=["SS_PF_Config"],E={payrollProvider:"PTO_Payroll_Provider",payrollDate:"PTO_Analysis_Date",accountingPeriod:"PTO_Accounting_Period",journalEntryId:"PTO_Journal_Entry_ID",companyName:"SS_Company_Name",accountingSoftware:"SS_Accounting_Software",reviewerName:"PTO_Reviewer",validationDataBalance:"PTO_Validation_Data_Balance",validationCleanBalance:"PTO_Validation_Clean_Balance",validationDifference:"PTO_Validation_Difference",headcountRosterCount:"PTO_Headcount_Roster_Count",headcountPayrollCount:"PTO_Headcount_Payroll_Count",headcountDifference:"PTO_Headcount_Difference",journalDebitTotal:"PTO_JE_Debit_Total",journalCreditTotal:"PTO_JE_Credit_Total",journalDifference:"PTO_JE_Difference"},me="User opted to skip the headcount review this period.",Fe={0:{note:"PTO_Notes_Config",reviewer:"PTO_Reviewer_Config",signOff:"PTO_SignOff_Config"},1:{note:"PTO_Notes_Import",reviewer:"PTO_Reviewer_Import",signOff:"PTO_SignOff_Import"},2:{note:"PTO_Notes_Headcount",reviewer:"PTO_Reviewer_Headcount",signOff:"PTO_SignOff_Headcount"},3:{note:"PTO_Notes_Validate",reviewer:"PTO_Reviewer_Validate",signOff:"PTO_SignOff_Validate"},4:{note:"PTO_Notes_Review",reviewer:"PTO_Reviewer_Review",signOff:"PTO_SignOff_Review"},5:{note:"PTO_Notes_JE",reviewer:"PTO_Reviewer_JE",signOff:"PTO_SignOff_JE"},6:{note:"PTO_Notes_Archive",reviewer:"PTO_Reviewer_Archive",signOff:"PTO_SignOff_Archive"}},on={0:"PTO_Complete_Config",1:"PTO_Complete_Import",2:"PTO_Complete_Headcount",3:"PTO_Complete_Validate",4:"PTO_Complete_Review",5:"PTO_Complete_JE",6:"PTO_Complete_Archive"};var te=[{id:0,title:"Configuration",summary:"Set the analysis date, accounting period, and review details for this run.",description:"Complete this step first to ensure all downstream calculations use the correct period settings.",actionLabel:"Configure Workbook",secondaryAction:{sheet:"SS_PF_Config",label:"Open Config Sheet"}},{id:1,title:"Import PTO Data",summary:"Pull your latest PTO export from payroll and paste it into PTO_Data.",description:"Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",actionLabel:"Import Sample Data",secondaryAction:{sheet:"PTO_Data",label:"Open Data Sheet"}},{id:2,title:"Headcount Review",summary:"Quick check to make sure your roster matches your PTO data.",description:"Compare employees in PTO_Data against your employee roster to catch any discrepancies.",actionLabel:"Open Headcount Review",secondaryAction:{sheet:"SS_Employee_Roster",label:"Open Sheet"}},{id:3,title:"Data Quality Review",summary:"Scan your PTO data for potential errors before crunching numbers.",description:"Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",actionLabel:"Click to Run Quality Check"},{id:4,title:"PTO Accrual Review",summary:"Review the calculated liability for each employee and compare to last period.",description:"The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",actionLabel:"Click to Perform Review"},{id:5,title:"Journal Entry Prep",summary:"Generate a balanced journal entry, run validation checks, and export when ready.",description:"Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",actionLabel:"Open Journal Draft",secondaryAction:{sheet:"PTO_JE_Draft",label:"Open Sheet"}},{id:6,title:"Archive & Reset",summary:"Save this period's results and prepare for the next cycle.",description:"Archive the current analysis so it becomes the 'prior period' for your next review.",actionLabel:"Archive Run"}],Nn={0:"PTO_Homepage",1:"PTO_Data",2:"PTO_Data",3:"PTO_Analysis",4:"PTO_Analysis",5:"PTO_JE_Draft"},Dn={PTO_Homepage:0,PTO_Data:1,PTO_Analysis:4,PTO_JE_Draft:5,PTO_Archive_Summary:6,SS_PF_Config:0,SS_Employee_Roster:2};var $n=te.reduce((e,t)=>(e[t.id]="pending",e),{}),B={activeView:"home",activeStepId:null,focusedIndex:0,stepStatuses:$n},C={loaded:!1,steps:{},permanents:{},completes:{},values:{},overrides:{accountingPeriod:!1,journalId:!1}},Re=null,st=null,He=null,Oe=new Map,L={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},z={debitTotal:null,creditTotal:null,difference:null,lineAmountSum:null,analysisChangeTotal:null,jeChangeTotal:null,loading:!1,lastError:null,validationRun:!1,issues:[]},W={hasRun:!1,loading:!1,acknowledged:!1,balanceIssues:[],zeroBalances:[],accrualOutliers:[],totalIssues:0,totalEmployees:0},Q={cleanDataReady:!1,employeeCount:0,lastRun:null,loading:!1,lastError:null,missingPayRates:[],missingDepartments:[],ignoredMissingPayRates:new Set,completenessCheck:{accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null}};async function jn(){var e;try{Re=document.getElementById("app"),st=document.getElementById("loading"),await Mn(),await Vn(),(e=window.PrairieForge)!=null&&e.loadSharedConfig&&await window.PrairieForge.loadSharedConfig();let t=Le(Ae);await je(t.sheetName,t.title,t.subtitle),await Ln(),st&&st.remove(),Re&&(Re.hidden=!1),ne()}catch(t){throw console.error("[PTO] Module initialization failed:",t),t}}async function Ln(){if(ae())try{await Excel.run(async e=>{e.workbook.worksheets.onActivated.add(Bn),await e.sync(),console.log("[PTO] Worksheet change listener registered")})}catch(e){console.warn("[PTO] Could not set up worksheet listener:",e)}}async function Bn(e){try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e.worksheetId);n.load("name"),await t.sync();let a=n.name,o=Dn[a];if(console.log(`[PTO] Tab changed to: ${a} \u2192 Step ${o}`),o!==void 0&&o!==B.activeStepId){let s=STEPS.findIndex(r=>r.id===o);if(s>=0){let r=o===0?"config":"step";B.activeView=r,B.activeStepId=o,B.focusedIndex=s,ne()}}})}catch(t){console.warn("[PTO] Error handling worksheet change:",t)}}async function Mn(){try{await qe(Ae),console.log(`[PTO] Tab visibility applied for ${Ae}`)}catch(e){console.warn("[PTO] Could not apply tab visibility:",e)}}async function Vn(){var e;if(!K()){C.loaded=!0;return}try{let t=await kt(fe),n={};(e=window.PrairieForge)!=null&&e.loadSharedConfig&&(await window.PrairieForge.loadSharedConfig(),window.PrairieForge._sharedConfigCache&&window.PrairieForge._sharedConfigCache.forEach((s,r)=>{n[r]=s}));let a={...t},o={SS_Default_Reviewer:E.reviewerName,Default_Reviewer:E.reviewerName,PTO_Reviewer:E.reviewerName,SS_Company_Name:E.companyName,Company_Name:E.companyName,SS_Payroll_Provider:E.payrollProvider,Payroll_Provider_Link:E.payrollProvider,SS_Accounting_Software:E.accountingSoftware,Accounting_Software_Link:E.accountingSoftware};Object.entries(o).forEach(([s,r])=>{n[s]&&!a[r]&&(a[r]=n[s])}),Object.entries(n).forEach(([s,r])=>{s.startsWith("PTO_")&&r&&(a[s]=r)}),C.permanents=await Hn(),C.values=a||{},C.overrides.accountingPeriod=!!(a!=null&&a[E.accountingPeriod]),C.overrides.journalId=!!(a!=null&&a[E.journalEntryId]),Object.entries(Fe).forEach(([s,r])=>{var l,c,i;C.steps[s]={notes:(l=a[r.note])!=null?l:"",reviewer:(c=a[r.reviewer])!=null?c:"",signOffDate:(i=a[r.signOff])!=null?i:""}}),C.completes=Object.entries(on).reduce((s,[r,l])=>{var c;return s[r]=(c=a[l])!=null?c:"",s},{}),C.loaded=!0}catch(t){console.warn("PTO: unable to load configuration fields",t),C.loaded=!0}}async function Hn(){let e={};if(!K())return e;let t=new Map;Object.entries(Fe).forEach(([n,a])=>{a.note&&t.set(a.note.trim(),Number(n))});try{await Excel.run(async n=>{let a=n.workbook.tables.getItemOrNullObject(fe[0]);if(await n.sync(),a.isNullObject)return;let o=a.getDataBodyRange(),s=a.getHeaderRowRange();o.load("values"),s.load("values"),await n.sync();let l=(s.values[0]||[]).map(i=>String(i||"").trim().toLowerCase()),c={field:l.findIndex(i=>i==="field"||i==="field name"||i==="setting"),permanent:l.findIndex(i=>i==="permanent"||i==="persist")};c.field===-1||c.permanent===-1||(o.values||[]).forEach(i=>{let p=String(i[c.field]||"").trim(),g=t.get(p);if(g==null)return;let u=ga(i[c.permanent]);e[g]=u})})}catch(n){console.warn("PTO: unable to load permanent flags",n)}return e}function ne(){var l;if(!Re)return;let e=B.focusedIndex<=0?"disabled":"",t=B.focusedIndex>=te.length-1?"disabled":"",n=B.activeView==="step"&&B.activeStepId!=null,o=B.activeView==="config"?sn():n?Yn(B.activeStepId):`${Un()}${Gn()}`;Re.innerHTML=`
        <div class="pf-root">
            <div class="pf-brand-float" aria-hidden="true">
                <span class="pf-brand-wave"></span>
            </div>
            <header class="pf-banner">
                <div class="pf-nav-bar">
                    <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                        ${Bt}
                        <span class="sr-only">Previous step</span>
                    </button>
                    <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                        ${It}
                        <span class="sr-only">Module Home</span>
                    </button>
                    <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                        ${Rt}
                        <span class="sr-only">Module Selector</span>
                    </button>
                    <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                        ${Vt}
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
                                ${Nt}
                                <span>Employee Roster</span>
                            </button>
                            <button id="nav-accounts" class="pf-quick-item pf-clickable" type="button">
                                ${Dt}
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
                    <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${Pn}</div>
                    <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
                </div>
            </footer>
        </div>
    `;let s=B.activeView==="home"||B.activeView!=="step"&&B.activeView!=="config",r=document.getElementById("pf-info-fab-pto");if(s)r&&r.remove();else if((l=window.PrairieForge)!=null&&l.mountInfoFab){let c=Fn(B.activeStepId);PrairieForge.mountInfoFab({title:c.title,content:c.content,buttonId:"pf-info-fab-pto"})}Wn(),Zn(),s?_t():We()}function Fn(e){switch(e){case 0:return{title:"Configuration",content:`
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
                `}}}function Un(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">PTO Accrual</h2>
            <p class="pf-hero-copy">${In}</p>
        </section>
    `}function Gn(){return`
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${te.map((e,t)=>Jn(e,t)).join("")}
            </div>
        </section>
    `}function Jn(e,t){let n=B.stepStatuses[e.id]||"pending",a=B.activeView==="step"&&B.focusedIndex===t?"pf-step-card--active":"",o=$t(ca(e.id));return`
        <article class="pf-step-card pf-clickable ${a}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${e.title}</h3>
        </article>
    `}function zn(e){let t=te.filter(o=>o.id!==6).map(o=>({id:o.id,title:o.title,complete:ea(o.id)})),n=t.every(o=>o.complete),a=t.map(o=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${o.complete?"is-active":""}" aria-pressed="${o.complete}">
                        ${Pe}
                    </span>
                    <div>
                        <h3>${k(o.title)}</h3>
                        <p class="pf-config-subtext">${o.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${k(ge)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${k(e.title)}</h2>
            <p class="pf-hero-copy">${k(e.summary||"")}</p>
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
    `}function sn(){if(!C.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=Xt(ie(E.payrollDate)),t=Xt(ie(E.accountingPeriod)),n=ie(E.journalEntryId),a=ie(E.accountingSoftware),o=ie(E.payrollProvider),s=ie(E.companyName),r=ie(E.reviewerName),l=ve(0),c=!!C.permanents[0],i=!!(mn(C.completes[0])||l.signOffDate),p=ye(l==null?void 0:l.reviewer),g=(l==null?void 0:l.signOffDate)||"";return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${k(ge)} | Step 0</p>
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
                        <input type="text" id="config-user-name" value="${k(r)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${k(e)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${k(t)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${k(n)}" placeholder="PTO-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${k(s)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${k(o)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${k(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${de({textareaId:"config-notes",value:l.notes||"",permanentId:"config-notes-lock",isPermanent:c,hintId:"",saveButtonId:"config-notes-save"})}
            ${ue({reviewerInputId:"config-reviewer",reviewerValue:p,signoffInputId:"config-signoff-date",signoffValue:g,isComplete:i,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"})}
        </section>
    `}function qn(e){let t=ve(1),n=!!C.permanents[1],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(C.completes[1])||o),r=ie(E.payrollProvider);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${k(ge)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${k(e.title)}</h2>
            <p class="pf-hero-copy">${k(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export, then paste into PTO_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${J(r?`<a href="${k(r)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${et}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${et}</button>`,"Provider")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${Qe}</button>`,"PTO_Data")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${Ht}</button>`,"Clear")}
                </div>
            </article>
            ${de({textareaId:"step-notes-1",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-1",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-1"})}
            ${ue({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
        </section>
    `}function Yn(e){let t=te.find(l=>l.id===e);if(!t)return"";if(e===0)return sn();if(e===1)return qn(t);if(e===2)return va(t);if(e===3)return wa(t);if(e===4)return ka(t);if(e===5)return Oa(t);if(t.id===6)return zn(t);let n=ve(e),a=!!C.permanents[e],o=ye(n==null?void 0:n.reviewer),s=(n==null?void 0:n.signOffDate)||"",r=!!(Se(C.completes[e])||s);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${k(ge)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${k(t.title)}</h2>
            <p class="pf-hero-copy">${k(t.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${de({textareaId:`step-notes-${e}`,value:(n==null?void 0:n.notes)||"",permanentId:`step-notes-lock-${e}`,isPermanent:a,hintId:"",saveButtonId:`step-notes-save-${e}`})}
            ${ue({reviewerInputId:`step-reviewer-${e}`,reviewerValue:o,signoffInputId:`step-signoff-${e}`,signoffValue:s,isComplete:r,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`})}
        </section>
    `}function Wn(){var n,a,o,s,r,l,c;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",async()=>{var p;let i=Le(Ae);await je(i.sheetName,i.title,i.subtitle),De({activeView:"home",activeStepId:null}),(p=document.getElementById("pf-hero"))==null||p.scrollIntoView({behavior:"smooth",block:"start"})}),(a=document.getElementById("nav-selector"))==null||a.addEventListener("click",()=>{window.location.href=Rn}),(o=document.getElementById("nav-prev"))==null||o.addEventListener("click",()=>Jt(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>Jt(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",i=>{i.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",i=>{!(t!=null&&t.contains(i.target))&&!(e!=null&&e.contains(i.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(r=document.getElementById("nav-roster"))==null||r.addEventListener("click",()=>{Kt("SS_Employee_Roster"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(l=document.getElementById("nav-accounts"))==null||l.addEventListener("click",()=>{Kt("SS_Chart_of_Accounts"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(c=document.getElementById("showConfigSheets"))==null||c.addEventListener("click",async()=>{await la()}),document.querySelectorAll("[data-step-card]").forEach(i=>{let p=Number(i.getAttribute("data-step-index")),g=Number(i.getAttribute("data-step-id"));i.addEventListener("click",()=>Ne(p,g))}),B.activeView==="config"?Qn():B.activeView==="step"&&B.activeStepId!=null&&Kn(B.activeStepId)}function Kn(e){var p,g,u,d,h,y,b,f,v,w,R,_,$,j,D,V,S;let t=e===2?document.getElementById("step-notes-input"):document.getElementById(`step-notes-${e}`),n=e===2?document.getElementById("step-reviewer-name"):document.getElementById(`step-reviewer-${e}`),a=e===2?document.getElementById("step-signoff-date"):document.getElementById(`step-signoff-${e}`),o=document.getElementById("step-back-btn"),s=e===2?document.getElementById("step-notes-lock-2"):document.getElementById(`step-notes-lock-${e}`),r=e===2?document.getElementById("step-notes-save-2"):document.getElementById(`step-notes-save-${e}`);r==null||r.addEventListener("click",async()=>{let I=(t==null?void 0:t.value)||"";await ee(e,"notes",I),re(r,!0)});let l=e===2?document.getElementById("headcount-signoff-save"):document.getElementById(`step-signoff-save-${e}`);l==null||l.addEventListener("click",async()=>{let I=(n==null?void 0:n.value)||"";await ee(e,"reviewer",I),re(l,!0)}),nt();let c=e===2?"headcount-signoff-toggle":`step-signoff-toggle-${e}`,i=e===2?"step-signoff-date":`step-signoff-${e}`;gn(e,{buttonId:c,inputId:i,canActivate:e===2?()=>{var O;return!hn()||((O=document.getElementById("step-notes-input"))==null?void 0:O.value.trim())||""?!0:(X("Please enter a brief explanation of the headcount differences before completing this step.","info"),!1)}:null,onComplete:Xn(e)}),o==null||o.addEventListener("click",async()=>{let I=Le(Ae);await je(I.sheetName,I.title,I.subtitle),De({activeView:"home",activeStepId:null})}),s==null||s.addEventListener("click",async()=>{let I=!s.classList.contains("is-locked");tt(s,I),await pn(e,I)}),e===6&&((p=document.getElementById("archive-run-btn"))==null||p.addEventListener("click",()=>{})),e===1&&((g=document.getElementById("import-open-data-btn"))==null||g.addEventListener("click",()=>dn("PTO_Data")),(u=document.getElementById("import-clear-btn"))==null||u.addEventListener("click",()=>ra())),e===2&&((d=document.getElementById("headcount-skip-btn"))==null||d.addEventListener("click",()=>{L.skipAnalysis=!L.skipAnalysis;let I=document.getElementById("headcount-skip-btn");I==null||I.classList.toggle("is-active",L.skipAnalysis),L.skipAnalysis&&an(),nn()}),(h=document.getElementById("headcount-run-btn"))==null||h.addEventListener("click",()=>rt()),(y=document.getElementById("headcount-refresh-btn"))==null||y.addEventListener("click",()=>rt()),_a(),L.skipAnalysis&&an(),nn()),e===3&&((b=document.getElementById("quality-run-btn"))==null||b.addEventListener("click",()=>qt()),(f=document.getElementById("quality-refresh-btn"))==null||f.addEventListener("click",()=>qt()),(v=document.getElementById("quality-acknowledge-btn"))==null||v.addEventListener("click",()=>na())),e===4&&((w=document.getElementById("analysis-refresh-btn"))==null||w.addEventListener("click",()=>Yt()),(R=document.getElementById("analysis-run-btn"))==null||R.addEventListener("click",()=>Yt()),(_=document.getElementById("payrate-save-btn"))==null||_.addEventListener("click",zt),($=document.getElementById("payrate-ignore-btn"))==null||$.addEventListener("click",ta),(j=document.getElementById("payrate-input"))==null||j.addEventListener("keydown",I=>{I.key==="Enter"&&zt()})),e===5&&((D=document.getElementById("je-create-btn"))==null||D.addEventListener("click",()=>sa()),(V=document.getElementById("je-run-btn"))==null||V.addEventListener("click",()=>cn()),(S=document.getElementById("je-export-btn"))==null||S.addEventListener("click",()=>ia()))}function Qn(){var l,c,i,p,g;Tt("config-payroll-date",{onChange:u=>{if(se(E.payrollDate,u),!!u){if(!C.overrides.accountingPeriod){let d=pa(u);if(d){let h=document.getElementById("config-accounting-period");h&&(h.value=d),se(E.accountingPeriod,d)}}if(!C.overrides.journalId){let d=fa(u);if(d){let h=document.getElementById("config-journal-id");h&&(h.value=d),se(E.journalEntryId,d)}}}}});let e=document.getElementById("config-accounting-period");e==null||e.addEventListener("change",u=>{C.overrides.accountingPeriod=!!u.target.value,se(E.accountingPeriod,u.target.value||"")});let t=document.getElementById("config-journal-id");t==null||t.addEventListener("change",u=>{C.overrides.journalId=!!u.target.value,se(E.journalEntryId,u.target.value.trim())}),(l=document.getElementById("config-company-name"))==null||l.addEventListener("change",u=>{se(E.companyName,u.target.value.trim())}),(c=document.getElementById("config-payroll-provider"))==null||c.addEventListener("change",u=>{se(E.payrollProvider,u.target.value.trim())}),(i=document.getElementById("config-accounting-link"))==null||i.addEventListener("change",u=>{se(E.accountingSoftware,u.target.value.trim())}),(p=document.getElementById("config-user-name"))==null||p.addEventListener("change",u=>{se(E.reviewerName,u.target.value.trim())});let n=document.getElementById("config-notes");n==null||n.addEventListener("input",u=>{ee(0,"notes",u.target.value)});let a=document.getElementById("config-notes-lock");a==null||a.addEventListener("click",async()=>{let u=!a.classList.contains("is-locked");tt(a,u),await pn(0,u)});let o=document.getElementById("config-notes-save");o==null||o.addEventListener("click",async()=>{n&&(await ee(0,"notes",n.value),re(o,!0))});let s=document.getElementById("config-reviewer");s==null||s.addEventListener("change",u=>{let d=u.target.value.trim();ee(0,"reviewer",d);let h=document.getElementById("config-signoff-date");if(d&&h&&!h.value){let y=ct();h.value=y,ee(0,"signOffDate",y),fn(0,!0)}}),(g=document.getElementById("config-signoff-date"))==null||g.addEventListener("change",u=>{ee(0,"signOffDate",u.target.value||"")});let r=document.getElementById("config-signoff-save");r==null||r.addEventListener("click",async()=>{var h,y;let u=((h=s==null?void 0:s.value)==null?void 0:h.trim())||"",d=((y=document.getElementById("config-signoff-date"))==null?void 0:y.value)||"";await ee(0,"reviewer",u),await ee(0,"signOffDate",d),re(r,!0)}),nt(),gn(0,{buttonId:"config-signoff-toggle",inputId:"config-signoff-date",onComplete:()=>{ya(),rn(0),ln()}})}function Ne(e,t=null){if(e<0||e>=te.length)return;He=e;let n=t!=null?t:te[e].id;De({focusedIndex:e,activeView:n===0?"config":"step",activeStepId:n});let o=Nn[n];o&&dn(o),n===2&&!L.hasAnalyzed&&(yn(),rt())}function Xn(e){return e===6?null:()=>rn(e)}function rn(e){let t=te.findIndex(a=>a.id===e);if(t===-1)return;let n=t+1;n<te.length&&(Ne(n,te[n].id),ln())}function ln(){let e=[document.querySelector(".pf-root"),document.querySelector(".pf-step-guide"),document.body];for(let t of e)t&&t.scrollTo({top:0,behavior:"smooth"});window.scrollTo({top:0,behavior:"smooth"})}function Jt(e){let t=B.focusedIndex+e,n=Math.max(0,Math.min(te.length-1,t));Ne(n,te[n].id)}function Zn(){if(He===null)return;let e=document.querySelector(`[data-step-index="${He}"]`);He=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}function ea(e){return mn(C.completes[e])}function De(e){e.stepStatuses&&(B.stepStatuses={...B.stepStatuses,...e.stepStatuses}),Object.assign(B,{...e,stepStatuses:B.stepStatuses}),ne()}function ae(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}async function zt(){let e=document.getElementById("payrate-input");if(!e)return;let t=parseFloat(e.value),n=e.dataset.employee,a=parseInt(e.dataset.row,10);if(isNaN(t)||t<=0){X("Please enter a valid pay rate greater than 0.","info");return}if(!n||isNaN(a)){console.error("Missing employee data on input");return}Z(!0,"Updating pay rate...");try{await Excel.run(async o=>{let s=o.workbook.worksheets.getItem("PTO_Analysis"),r=s.getCell(a-1,3);r.values=[[t]];let l=s.getCell(a-1,8);l.load("values"),await o.sync();let i=(Number(l.values[0][0])||0)*t,p=s.getCell(a-1,9);p.values=[[i]];let g=s.getCell(a-1,10);g.load("values"),await o.sync();let u=Number(g.values[0][0])||0,d=i-u,h=s.getCell(a-1,11);h.values=[[d]],await o.sync()}),Q.missingPayRates=Q.missingPayRates.filter(o=>o.name!==n),Z(!1),Ne(3,3)}catch(o){console.error("Failed to save pay rate:",o),X(`Failed to save pay rate: ${o.message}`,"error"),Z(!1)}}function ta(){let e=document.getElementById("payrate-input");if(!e)return;let t=e.dataset.employee;t&&(Q.ignoredMissingPayRates.add(t),Q.missingPayRates=Q.missingPayRates.filter(n=>n.name!==t)),Ne(3,3)}async function qt(){if(!ae()){X("Excel is not available. Open this module inside Excel to run quality check.","info");return}W.loading=!0,Z(!0,"Analyzing data quality..."),re(document.getElementById("quality-save-btn"),!1);try{await Excel.run(async t=>{var b;let a=t.workbook.worksheets.getItem("PTO_Data").getUsedRangeOrNullObject();a.load("values"),await t.sync();let o=a.isNullObject?[]:a.values||[];if(!o.length||o.length<2)throw new Error("PTO_Data is empty or has no data rows.");let s=(o[0]||[]).map(f=>q(f));console.log("[Data Quality] PTO_Data headers:",o[0]);let r=s.findIndex(f=>f==="employee name"||f==="employeename");r===-1&&(r=s.findIndex(f=>f.includes("employee")&&f.includes("name"))),r===-1&&(r=s.findIndex(f=>f==="name"||f.includes("name")&&!f.includes("company")&&!f.includes("form"))),console.log("[Data Quality] Employee name column index:",r,"Header:",(b=o[0])==null?void 0:b[r]);let l=M(s,["balance"]),c=M(s,["accrual rate","accrualrate"]),i=M(s,["carry over","carryover"]),p=M(s,["ytd accrued","ytdaccrued"]),g=M(s,["ytd used","ytdused"]),u=[],d=[],h=[],y=o.slice(1);y.forEach((f,v)=>{let w=v+2,R=r!==-1?String(f[r]||"").trim():`Row ${w}`;if(!R)return;let _=l!==-1&&Number(f[l])||0,$=c!==-1&&Number(f[c])||0,j=i!==-1&&Number(f[i])||0,D=p!==-1&&Number(f[p])||0,V=g!==-1&&Number(f[g])||0,S=j+D;_<0?u.push({name:R,issue:`Negative balance: ${_.toFixed(2)} hrs`,rowIndex:w}):V>S&&S>0&&u.push({name:R,issue:`Used ${V.toFixed(0)} hrs but only ${S.toFixed(0)} available`,rowIndex:w}),_===0&&(j>0||D>0)&&d.push({name:R,rowIndex:w}),$>8&&h.push({name:R,accrualRate:$,rowIndex:w})}),W.balanceIssues=u,W.zeroBalances=d,W.accrualOutliers=h,W.totalIssues=u.length,W.totalEmployees=y.filter(f=>f.some(v=>v!==null&&v!=="")).length,W.hasRun=!0});let e=W.balanceIssues.length>0;De({stepStatuses:{3:e?"blocked":"complete"}})}catch(e){console.error("Data quality check error:",e),X(`Quality check failed: ${e.message}`,"error"),W.hasRun=!1}finally{W.loading=!1,Z(!1),ne()}}function na(){W.acknowledged=!0,De({stepStatuses:{3:"complete"}}),ne()}async function aa(){if(ae())try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem("PTO_Data"),n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=t.getUsedRangeOrNullObject();if(a.load("values"),n.load("isNullObject"),await e.sync(),n.isNullObject){Q.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let o=n.getUsedRangeOrNullObject();o.load("values"),await e.sync();let s=a.isNullObject?[]:a.values||[],r=o.isNullObject?[]:o.values||[];if(!s.length||!r.length){Q.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let l=(p,g,u)=>{let d=(p[0]||[]).map(b=>q(b)),h=M(d,g);return h===-1?null:p.slice(1).reduce((b,f)=>b+(Number(f[h])||0),0)},c=[{key:"accrualRate",aliases:["accrual rate","accrualrate"]},{key:"carryOver",aliases:["carry over","carryover","carry_over"]},{key:"ytdAccrued",aliases:["ytd accrued","ytdaccrued","ytd_accrued"]},{key:"ytdUsed",aliases:["ytd used","ytdused","ytd_used"]},{key:"balance",aliases:["balance"]}],i={};for(let p of c){let g=l(s,p.aliases,"PTO_Data"),u=l(r,p.aliases,"PTO_Analysis");if(g===null||u===null)i[p.key]=null;else{let d=Math.abs(g-u)<.01;i[p.key]={match:d,ptoData:g,ptoAnalysis:u}}}Q.completenessCheck=i})}catch(e){console.error("Completeness check failed:",e)}}async function Yt(){if(!ae()){X("Excel is not available. Open this module inside Excel to run analysis.","info");return}Z(!0,"Running analysis...");try{await yn(),await aa(),Q.cleanDataReady=!0,ne()}catch(e){console.error("Full analysis error:",e),X(`Analysis failed: ${e.message}`,"error")}finally{Z(!1)}}async function cn(){if(!ae()){X("Excel is not available. Open this module inside Excel to run journal checks.","info");return}z.loading=!0,z.lastError=null,re(document.getElementById("je-save-btn"),!1),ne();try{let e=await Excel.run(async t=>{let a=t.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();a.load("values");let o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis");o.load("isNullObject"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty. Generate the JE first.");let r=(s[0]||[]).map(w=>q(w)),l=M(r,["debit"]),c=M(r,["credit"]),i=M(r,["lineamount","line amount"]),p=M(r,["account number","accountnumber"]);if(l===-1||c===-1)throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");let g=0,u=0,d=0,h=0;s.slice(1).forEach(w=>{let R=Number(w[l])||0,_=Number(w[c])||0,$=i!==-1&&Number(w[i])||0,j=p!==-1?String(w[p]||"").trim():"";g+=R,u+=_,d+=$,j&&j!=="21540"&&(h+=$)});let y=0;if(!o.isNullObject){let w=o.getUsedRangeOrNullObject();w.load("values"),await t.sync();let R=w.isNullObject?[]:w.values||[];if(R.length>1){let _=(R[0]||[]).map(j=>q(j)),$=M(_,["change"]);$!==-1&&R.slice(1).forEach(j=>{y+=Number(j[$])||0})}}let b=g-u,f=[];Math.abs(b)>=.01?f.push({check:"Debits = Credits",passed:!1,detail:b>0?`Debits exceed credits by $${Math.abs(b).toLocaleString(void 0,{minimumFractionDigits:2})}`:`Credits exceed debits by $${Math.abs(b).toLocaleString(void 0,{minimumFractionDigits:2})}`}):f.push({check:"Debits = Credits",passed:!0,detail:""}),Math.abs(d)>=.01?f.push({check:"Line Amounts Sum to Zero",passed:!1,detail:`Line amounts sum to $${d.toLocaleString(void 0,{minimumFractionDigits:2})} (should be $0.00)`}):f.push({check:"Line Amounts Sum to Zero",passed:!0,detail:""});let v=Math.abs(h-y);return v>=.01?f.push({check:"JE Matches Analysis Total",passed:!1,detail:`JE expense total ($${h.toLocaleString(void 0,{minimumFractionDigits:2})}) differs from PTO_Analysis Change total ($${y.toLocaleString(void 0,{minimumFractionDigits:2})}) by $${v.toLocaleString(void 0,{minimumFractionDigits:2})}`}):f.push({check:"JE Matches Analysis Total",passed:!0,detail:""}),{debitTotal:g,creditTotal:u,difference:b,lineAmountSum:d,jeChangeTotal:h,analysisChangeTotal:y,issues:f,validationRun:!0}});Object.assign(z,e,{lastError:null})}catch(e){console.warn("PTO JE summary:",e),z.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",z.debitTotal=null,z.creditTotal=null,z.difference=null,z.lineAmountSum=null,z.jeChangeTotal=null,z.analysisChangeTotal=null,z.issues=[],z.validationRun=!1}finally{z.loading=!1,ne()}}var oa={"general & administrative":"64110","general and administrative":"64110","g&a":"64110","research & development":"62110","research and development":"62110","r&d":"62110",marketing:"61610","cogs onboarding":"53110","cogs prof. services":"56110","cogs professional services":"56110","sales & marketing":"61110","sales and marketing":"61110","cogs support":"52110","client success":"61811"},Wt="21540";async function sa(){if(!ae()){X("Excel is not available. Open this module inside Excel to create the journal entry.","info");return}Z(!0,"Creating PTO Journal Entry...");try{await Excel.run(async e=>{let t=[],n=e.workbook.tables.getItemOrNullObject(fe[0]);if(n.load("isNullObject"),await e.sync(),n.isNullObject){let m=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");if(m.load("isNullObject"),await e.sync(),!m.isNullObject){let P=m.getUsedRangeOrNullObject();P.load("values"),await e.sync();let A=P.isNullObject?[]:P.values||[];t=A.length>1?A.slice(1):[]}}else{let m=n.getDataBodyRange();m.load("values"),await e.sync(),t=m.values||[]}let a=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis");if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PTO_Analysis sheet not found. Please ensure the worksheet exists.");let o=a.getUsedRangeOrNullObject();o.load("values");let s=e.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");s.load("isNullObject"),await e.sync();let r=[];if(!s.isNullObject){let m=s.getUsedRangeOrNullObject();m.load("values"),await e.sync(),r=m.isNullObject?[]:m.values||[]}let l=o.isNullObject?[]:o.values||[];if(!l.length||l.length<2)throw new Error("PTO_Analysis is empty or has no data rows. Run the analysis first (Step 4).");let c={};t.forEach(m=>{let P=String(m[1]||"").trim(),A=m[2];P&&(c[P]=A)}),(!c[E.journalEntryId]||!c[E.payrollDate])&&console.warn("[JE Draft] Missing config values - RefNumber:",c[E.journalEntryId],"TxnDate:",c[E.payrollDate]);let i=c[E.journalEntryId]||"",p=c[E.payrollDate]||"",g=c[E.accountingPeriod]||"",u="";if(p)try{let m;if(typeof p=="number"||/^\d{4,5}$/.test(String(p).trim())){let P=Number(p),A=new Date(1899,11,30);m=new Date(A.getTime()+P*24*60*60*1e3)}else m=new Date(p);if(!isNaN(m.getTime())&&m.getFullYear()>1970){let P=String(m.getMonth()+1).padStart(2,"0"),A=String(m.getDate()).padStart(2,"0"),N=m.getFullYear();u=`${P}/${A}/${N}`}else console.warn("[JE Draft] Date parsing resulted in invalid date:",p,"->",m),u=String(p)}catch(m){console.warn("[JE Draft] Could not parse TxnDate:",p,m),u=String(p)}let d=g?`${g} PTO Accrual`:"PTO Accrual",h={};if(r.length>1){let m=(r[0]||[]).map(N=>q(N)),P=M(m,["account number","accountnumber","account","acct"]),A=M(m,["account name","accountname","name","description"]);P!==-1&&A!==-1&&r.slice(1).forEach(N=>{let Y=String(N[P]||"").trim(),le=String(N[A]||"").trim();Y&&(h[Y]=le)})}let y=(l[0]||[]).map(m=>q(m));console.log("[JE Draft] PTO_Analysis headers:",y),console.log("[JE Draft] PTO_Analysis row count:",l.length-1);let b=M(y,["department"]),f=M(y,["change"]);if(console.log("[JE Draft] Column indices - Department:",b,"Change:",f),b===-1||f===-1)throw new Error(`Could not find required columns in PTO_Analysis. Found headers: ${y.join(", ")}. Looking for "Department" (found: ${b!==-1}) and "Change" (found: ${f!==-1}).`);let v={},w=0,R=0,_=0;if(l.slice(1).forEach((m,P)=>{w++;let A=String(m[b]||"").trim(),N=m[f],Y=Number(N)||0;if(P<3&&console.log(`[JE Draft] Row ${P+2}: Dept="${A}", Change raw="${N}", Change num=${Y}`),!A){_++;return}if(Y===0){R++;return}v[A]||(v[A]=0),v[A]+=Y}),console.log(`[JE Draft] Data summary: ${w} rows, ${R} with zero change, ${_} missing dept`),console.log("[JE Draft] Department totals:",v),Object.keys(v).length===0){let m=`No journal entry lines could be created.

`;throw R===w?(m+=`All 'Change' amounts in PTO_Analysis are $0.00.

`,m+=`Common causes:
`,m+=`\u2022 Missing Pay Rate data (Liability = Balance \xD7 Pay Rate)
`,m+=`\u2022 No prior period data to compare against
`,m+=`\u2022 PTO Analysis hasn't been run yet

`,m+="Please verify Pay Rate values exist in PTO_Analysis."):_===w?(m+=`All rows are missing Department values.

`,m+="Please ensure the 'Department' column is populated in PTO_Analysis."):(m+=`Found ${w} rows but none had both a Department and non-zero Change amount.
`,m+=`\u2022 ${R} rows with zero change
`,m+=`\u2022 ${_} rows missing department`),new Error(m)}let j=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],D=[j],V=0,S=0;Object.entries(v).forEach(([m,P])=>{if(Math.abs(P)<.01)return;let A=m.toLowerCase().trim(),N=oa[A]||"",Y=h[N]||"",le=P>0?Math.abs(P):0,x=P<0?Math.abs(P):0;V+=le,S+=x,D.push([i,u,N,Y,P,le,x,d,m])});let I=V-S;if(Math.abs(I)>=.01){let m=I<0?Math.abs(I):0,P=I>0?Math.abs(I):0,A=h[Wt]||"Accrued PTO";D.push([i,u,Wt,A,-I,m,P,d,""])}let O=e.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");if(O.load("isNullObject"),await e.sync(),O.isNullObject)O=e.workbook.worksheets.add("PTO_JE_Draft");else{let m=O.getUsedRangeOrNullObject();m.load("isNullObject"),await e.sync(),m.isNullObject||m.clear()}if(D.length>0){let m=O.getRangeByIndexes(0,0,D.length,j.length);m.values=D;let P=O.getRangeByIndexes(0,0,1,j.length);ot(P);let A=D.length-1;A>0&&(pe(O,4,A,!0),pe(O,5,A),pe(O,6,A)),m.format.autofitColumns()}await e.sync(),O.activate(),O.getRange("A1").select(),await e.sync()}),await cn()}catch(e){console.error("Create JE Draft error:",e),X(`Unable to create Journal Entry: ${e.message}`,"error")}finally{Z(!1)}}async function ia(){if(!ae()){X("Excel is not available. Open this module inside Excel to export.","info");return}Z(!0,"Preparing JE CSV...");try{let{rows:e}=await Excel.run(async n=>{let o=n.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();o.load("values"),await n.sync();let s=o.isNullObject?[]:o.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty.");return{rows:s}}),t=xa(e);Ca(`pto-je-draft-${ct()}.csv`,t)}catch(e){console.error("PTO JE export:",e),X("Unable to export the JE draft. Confirm the sheet has data.","error")}finally{Z(!1)}}async function dn(e){if(!(!e||!ae()))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e);n.activate(),n.getRange("A1").select(),await t.sync()})}catch(t){console.error(t)}}async function ra(){if(!(!ae()||!await Tn(`All data in PTO_Data will be permanently removed.

This action cannot be undone.`,{title:"Clear PTO Data",icon:"\u{1F5D1}\uFE0F",confirmText:"Clear Data",cancelText:"Keep Data",destructive:!0}))){Z(!0);try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("PTO_Data"),a=n.getUsedRangeOrNullObject();a.load("rowCount"),await t.sync(),!a.isNullObject&&a.rowCount>1&&(n.getRangeByIndexes(1,0,a.rowCount-1,20).clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),X("PTO_Data cleared successfully. You can now paste new data.","success")}catch(t){console.error("Clear PTO_Data error:",t),X(`Failed to clear PTO_Data: ${t.message}`,"error")}finally{Z(!1)}}}async function Kt(e){if(!e||!ae())return;let t={SS_Employee_Roster:["Employee","Department","Pay_Rate","Status","Hire_Date"],SS_Chart_of_Accounts:["Account_Number","Account_Name","Type","Category"]};try{await Excel.run(async n=>{let a=n.workbook.worksheets.getItemOrNullObject(e);if(a.load("isNullObject"),await n.sync(),a.isNullObject){a=n.workbook.worksheets.add(e);let o=t[e]||["Column1","Column2"],s=a.getRange(`A1:${String.fromCharCode(64+o.length)}1`);s.values=[o],s.format.font.bold=!0,s.format.fill.color="#f0f0f0",s.format.autofitColumns(),await n.sync()}a.activate(),a.getRange("A1").select(),await n.sync()})}catch(n){console.error("Error opening reference sheet:",n)}}async function la(){if(!ae()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(o=>{o.name.toUpperCase().startsWith("SS_")&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Config] Made visible: ${o.name}`),n++)}),await e.sync();let a=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");a.load("isNullObject"),await e.sync(),a.isNullObject||(a.activate(),a.getRange("A1").select(),await e.sync()),console.log(`[Config] ${n} system sheets now visible`)})}catch(e){console.error("[Config] Error unhiding system sheets:",e)}}function ie(e){var n,a;let t=String(e!=null?e:"").trim();return(a=(n=C.values)==null?void 0:n[t])!=null?a:""}function ye(e){var n;if(e)return e;let t=ie(E.reviewerName);if(t)return t;if((n=window.PrairieForge)!=null&&n._sharedConfigCache){let a=window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer")||window.PrairieForge._sharedConfigCache.get("Default_Reviewer");if(a)return a}return""}function se(e,t,n={}){var r;let a=String(e!=null?e:"").trim();if(!a)return;C.values[a]=t!=null?t:"";let o=(r=n.debounceMs)!=null?r:0;if(!o){let l=Oe.get(a);l&&clearTimeout(l),Oe.delete(a),Ce(a,t!=null?t:"",fe);return}Oe.has(a)&&clearTimeout(Oe.get(a));let s=setTimeout(()=>{Oe.delete(a),Ce(a,t!=null?t:"",fe)},o);Oe.set(a,s)}function q(e){return String(e!=null?e:"").trim().toLowerCase()}function Z(e,t="Working..."){let n=document.getElementById(An);n&&(n.style.display="none")}function it(){jn()}typeof Office!="undefined"&&Office.onReady?Office.onReady(()=>it()).catch(()=>it()):it();function ve(e){return C.steps[e]||{notes:"",reviewer:"",signOffDate:""}}function un(e){return Fe[e]||{}}function ca(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}async function ee(e,t,n){let a=C.steps[e]||{notes:"",reviewer:"",signOffDate:""};a[t]=n,C.steps[e]=a;let o=un(e),s=t==="notes"?o.note:t==="reviewer"?o.reviewer:o.signOff;if(s&&K())try{await Ce(s,n,fe)}catch(r){console.warn("PTO: unable to save field",s,r)}}async function pn(e,t){C.permanents[e]=t;let n=un(e);if(n!=null&&n.note&&K())try{await Excel.run(async a=>{var u;let o=a.workbook.tables.getItemOrNullObject(fe[0]);if(await a.sync(),o.isNullObject)return;let s=o.getDataBodyRange(),r=o.getHeaderRowRange();s.load("values"),r.load("values"),await a.sync();let l=r.values[0]||[],c=l.map(d=>String(d||"").trim().toLowerCase()),i={field:c.findIndex(d=>d==="field"||d==="field name"||d==="setting"),permanent:c.findIndex(d=>d==="permanent"||d==="persist"),value:c.findIndex(d=>d==="value"||d==="setting value"),type:c.findIndex(d=>d==="type"||d==="category"),title:c.findIndex(d=>d==="title"||d==="display name")};if(i.field===-1)return;let g=(s.values||[]).findIndex(d=>String(d[i.field]||"").trim()===n.note);if(g>=0)i.permanent>=0&&(s.getCell(g,i.permanent).values=[[t?"Y":"N"]]);else{let d=new Array(l.length).fill("");i.type>=0&&(d[i.type]="Other"),i.title>=0&&(d[i.title]=""),d[i.field]=n.note,i.permanent>=0&&(d[i.permanent]=t?"Y":"N"),i.value>=0&&(d[i.value]=((u=C.steps[e])==null?void 0:u.notes)||""),o.rows.add(null,[d])}await a.sync()})}catch(a){console.warn("PTO: unable to update permanent flag",a)}}async function fn(e,t){let n=on[e];if(n&&(C.completes[e]=t?"Y":"",!!K()))try{await Ce(n,t?"Y":"",fe)}catch(a){console.warn("PTO: unable to save completion flag",n,a)}}function Qt(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function da(){let e={};return Object.keys(Fe).forEach(t=>{var s;let n=parseInt(t,10),a=!!((s=C.steps[n])!=null&&s.signOffDate),o=!!C.completes[n];e[n]=a||o}),e}function gn(e,{buttonId:t,inputId:n,canActivate:a=null,onComplete:o=null}){var c;let s=document.getElementById(t);if(!s)return;let r=document.getElementById(n),l=!!((c=C.steps[e])!=null&&c.signOffDate)||!!C.completes[e];Qt(s,l),s.addEventListener("click",()=>{if(!s.classList.contains("is-active")&&e>0){let g=da(),{canComplete:u,message:d}=Ft(e,g);if(!u){Ut(d);return}}if(typeof a=="function"&&!a())return;let p=!s.classList.contains("is-active");Qt(s,p),r&&(r.value=p?ct():"",ee(e,"signOffDate",r.value)),fn(e,p),p&&typeof o=="function"&&o()})}function k(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function ua(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function mn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function Se(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function lt(e){if(!e)return null;let t=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(e));if(!t)return null;let n=Number(t[1]),a=Number(t[2]),o=Number(t[3]);return!n||!a||!o?null:{year:n,month:a,day:o}}function Xt(e){if(!e)return"";let t=lt(e);if(!t)return"";let{year:n,month:a,day:o}=t;return`${n}-${String(a).padStart(2,"0")}-${String(o).padStart(2,"0")}`}function pa(e){let t=lt(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function fa(e){let t=lt(e);return t?`PTO-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function ct(){let e=new Date,t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}function ga(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="y"||t==="yes"||t==="true"||t==="t"||t==="1"}function ma(e){if(e instanceof Date)return e.getTime();if(typeof e=="number"){let n=ha(e);return n?n.getTime():null}let t=new Date(e);return Number.isNaN(t.getTime())?null:t.getTime()}function ha(e){if(!Number.isFinite(e))return null;let t=new Date(Date.UTC(1899,11,30));return new Date(t.getTime()+e*24*60*60*1e3)}function ya(){let e=n=>{var a,o;return((o=(a=document.getElementById(n))==null?void 0:a.value)==null?void 0:o.trim())||""};[{id:"config-payroll-date",field:E.payrollDate},{id:"config-accounting-period",field:E.accountingPeriod},{id:"config-journal-id",field:E.journalEntryId},{id:"config-company-name",field:E.companyName},{id:"config-payroll-provider",field:E.payrollProvider},{id:"config-accounting-link",field:E.accountingSoftware},{id:"config-user-name",field:E.reviewerName}].forEach(({id:n,field:a})=>{let o=e(n);a&&se(a,o)})}function M(e,t=[]){let n=t.map(a=>q(a));return e.findIndex(a=>n.some(o=>a.includes(o)))}function va(e){var w,R,_,$,j,D,V,S,I;let t=ve(2),n=(t==null?void 0:t.notes)||"",a=!!C.permanents[2],o=ye(t==null?void 0:t.reviewer),s=(t==null?void 0:t.signOffDate)||"",r=!!(Se(C.completes[2])||s),l=L.roster||{},c=L.hasAnalyzed,i=(R=(w=L.roster)==null?void 0:w.difference)!=null?R:0,p=!L.skipAnalysis&&Math.abs(i)>0,g=(_=l.rosterCount)!=null?_:0,u=($=l.payrollCount)!=null?$:0,d=(j=l.difference)!=null?j:u-g,h=Array.isArray(l.mismatches)?l.mismatches.filter(Boolean):[],y="";L.loading?y=((V=(D=window.PrairieForge)==null?void 0:D.renderStatusBanner)==null?void 0:V.call(D,{type:"info",message:"Analyzing headcount\u2026",escapeHtml:k}))||"":L.lastError&&(y=((I=(S=window.PrairieForge)==null?void 0:S.renderStatusBanner)==null?void 0:I.call(S,{type:"error",message:L.lastError,escapeHtml:k}))||"");let b=(O,m,P,A)=>{let N=!c,Y;N?Y='<span class="pf-je-check-circle pf-je-circle--pending"></span>':A?Y=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:Y=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let le=c?` = ${P}`:"";return`
            <div class="pf-je-check-row">
                ${Y}
                <span class="pf-je-check-desc-pill">${k(O)}${le}</span>
            </div>
        `},f=`
        ${b("SS_Employee_Roster count","Active employees in roster",g,!0)}
        ${b("PTO_Data count","Unique employees in PTO data",u,!0)}
        ${b("Difference","Should be zero",d,d===0)}
    `,v=h.length&&!L.skipAnalysis&&c?window.PrairieForge.renderMismatchTiles({mismatches:h,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"PTO Data",escapeHtml:k}):"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${k(ge)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${k(e.title)}</h2>
            <p class="pf-hero-copy">${k(e.summary||"")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${L.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${Lt}
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
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${Me}</button>`,"Run")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${Te}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${y}
                <div class="pf-je-checks-container">
                    ${f}
                </div>
                ${v}
            </article>
            ${de({textareaId:"step-notes-input",value:n,permanentId:"step-notes-lock-2",isPermanent:a,hintId:p?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
            ${ue({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:r,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
        </section>
    `}function ba(){let e=Q.completenessCheck||{},t=Q.missingPayRates||[],n=[{key:"accrualRate",label:"Accrual Rate",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"carryOver",label:"Carry Over",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdAccrued",label:"YTD Accrued",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdUsed",label:"YTD Used",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"balance",label:"Balance",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"}],o=n.every(i=>e[i.key]!==null&&e[i.key]!==void 0)&&n.every(i=>{var p;return(p=e[i.key])==null?void 0:p.match}),s=t.length>0,r=i=>{let p=e[i.key],g=p==null,u;return g?u='<span class="pf-je-check-circle pf-je-circle--pending"></span>':p.match?u=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:u=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${u}
                <span class="pf-je-check-desc-pill">${k(i.label)}: ${k(i.desc)}</span>
            </div>
        `},l=n.map(i=>r(i)).join(""),c="";if(s){let i=t[0],p=t.length-1;c=`
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${k(i.name)}</strong> to calculate liability
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
                               data-employee="${ua(i.name)}"
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
    `}function wa(e){var d,h,y,b,f,v,w,R;let t=ve(3),n=!!C.permanents[3],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(C.completes[3])||o),r=W.hasRun,{balanceIssues:l,zeroBalances:c,accrualOutliers:i,totalEmployees:p}=W,g="";if(W.loading)g=((h=(d=window.PrairieForge)==null?void 0:d.renderStatusBanner)==null?void 0:h.call(d,{type:"info",message:"Analyzing data quality...",escapeHtml:k}))||"";else if(r){let _=l.length,$=i.length+c.length;_>0?g=((b=(y=window.PrairieForge)==null?void 0:y.renderStatusBanner)==null?void 0:b.call(y,{type:"error",title:`${_} Balance Issue${_>1?"s":""} Found`,message:"Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",escapeHtml:k}))||"":$>0?g=((v=(f=window.PrairieForge)==null?void 0:f.renderStatusBanner)==null?void 0:v.call(f,{type:"warning",title:"No Critical Issues",message:`${$} informational item${$>1?"s":""} to review (see below).`,escapeHtml:k}))||"":g=((R=(w=window.PrairieForge)==null?void 0:w.renderStatusBanner)==null?void 0:R.call(w,{type:"success",title:"Data Quality Passed",message:`${p} employee${p!==1?"s":""} checked \u2014 no anomalies found.`,escapeHtml:k}))||""}let u=[];return r&&l.length>0&&u.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u26A0\uFE0F</span>
                    <span class="pf-quality-issue-title">Balance Issues (${l.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${l.slice(0,5).map(_=>`<li><strong>${k(_.name)}</strong>: ${k(_.issue)}</li>`).join("")}
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
                    ${i.slice(0,5).map(_=>`<li><strong>${k(_.name)}</strong>: ${_.accrualRate.toFixed(2)} hrs/period</li>`).join("")}
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
                    ${c.slice(0,5).map(_=>`<li><strong>${k(_.name)}</strong></li>`).join("")}
                    ${c.length>5?`<li class="pf-quality-more">+${c.length-5} more</li>`:""}
                </ul>
            </div>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${k(ge)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${k(e.title)}</h2>
            <p class="pf-hero-copy">${k(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${g}
                <div class="pf-signoff-action">
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${Me}</button>`,"Run")}
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
                            ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${Te}</button>`,"Refresh")}
                            ${W.acknowledged?"":J(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${Pe}</button>`,"Continue")}
                        </div>
                    </div>
                </article>
            `:""}
            ${de({textareaId:"step-notes-3",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-3",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-3"})}
            ${ue({reviewerInputId:"step-reviewer-3",reviewerValue:a,signoffInputId:"step-signoff-3",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-3",completeButtonId:"step-signoff-toggle-3"})}
        </section>
    `}function ka(e){let t=ve(4),n=!!C.permanents[4],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(C.completes[4])||o);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${k(ge)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${k(e.title)}</h2>
            <p class="pf-hero-copy">${k(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Analysis</h3>
                    <p class="pf-config-subtext">Calculate liabilities and compare against last period.</p>
                </div>
                <div class="pf-signoff-action">
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${Me}</button>`,"Run")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${Te}</button>`,"Refresh")}
                </div>
            </article>
            ${ba()}
            ${de({textareaId:"step-notes-4",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-4",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-4"})}
            ${ue({reviewerInputId:"step-reviewer-4",reviewerValue:a,signoffInputId:"step-signoff-4",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"step-signoff-toggle-4"})}
        </section>
    `}function Oa(e){let t=ve(5),n=!!C.permanents[5],a=ye(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(Se(C.completes[5])||o),r=z.lastError?`<p class="pf-step-note">${k(z.lastError)}</p>`:"",l=z.validationRun,c=z.issues||[],i=[{key:"Debits = Credits",desc:"\u2211 Debit column = \u2211 Credit column"},{key:"Line Amounts Sum to Zero",desc:"\u2211 Line Amount = $0.00"},{key:"JE Matches Analysis Total",desc:"\u2211 Expense line amounts = \u2211 PTO_Analysis Change"}],p=h=>{let y=c.find(v=>v.check===h.key),b=!l,f;return b?f='<span class="pf-je-check-circle pf-je-circle--pending"></span>':y!=null&&y.passed?f=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:f=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${f}
                <span class="pf-je-check-desc-pill">${k(h.desc)}</span>
            </div>
        `},g=i.map(h=>p(h)).join(""),u=c.filter(h=>!h.passed),d="";return l&&u.length>0&&(d=`
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>\u26A0\uFE0F Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${u.map(h=>`<li><strong>${k(h.check)}:</strong> ${k(h.detail)}</li>`).join("")}
                </ul>
            </article>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${k(ge)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${k(e.title)}</h2>
            <p class="pf-hero-copy">${k(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${Qe}</button>`,"Generate")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${Te}</button>`,"Refresh")}
                    ${J(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${jt}</button>`,"Export")}
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
            ${de({textareaId:"step-notes-5",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-5",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-5"})}
            ${ue({reviewerInputId:"step-reviewer-5",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
        </section>
    `}function Sa(){var t,n;return Math.abs((n=(t=L.roster)==null?void 0:t.difference)!=null?n:0)>0}function hn(){return!L.skipAnalysis&&Sa()}async function rt(){if(!K()){L.loading=!1,L.lastError="Excel runtime is unavailable.",ne();return}L.loading=!0,L.lastError=null,re(document.getElementById("headcount-save-btn"),!1),ne();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),a=t.workbook.worksheets.getItem("PTO_Data"),o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.getUsedRangeOrNullObject(),r=a.getUsedRangeOrNullObject();s.load("values"),r.load("values"),o.load("isNullObject"),await t.sync();let l=null;o.isNullObject||(l=o.getUsedRangeOrNullObject(),l.load("values")),await t.sync();let c=s.isNullObject?[]:s.values||[],i=r.isNullObject?[]:r.values||[],p=l&&!l.isNullObject?l.values||[]:[],g=p.length?p:i;return Ea(c,g)});L.roster=e.roster,L.hasAnalyzed=!0,L.lastError=null}catch(e){console.warn("PTO headcount: unable to analyze data",e),L.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{L.loading=!1,ne()}}function Zt(e){if(!e)return!0;let t=e.toLowerCase().trim();return t?["total","subtotal","sum","count","grand","average","avg"].some(a=>t.includes(a)):!0}function Ea(e,t){let n={rosterCount:0,payrollCount:0,difference:0,mismatches:[]};if(((e==null?void 0:e.length)||0)<2||((t==null?void 0:t.length)||0)<2)return console.warn("Headcount: insufficient data rows",{rosterRows:(e==null?void 0:e.length)||0,payrollRows:(t==null?void 0:t.length)||0}),{roster:n};let a=en(e),o=en(t),s=a.headers,r=o.headers,l={employee:tn(s),termination:s.findIndex(d=>d.includes("termination"))},c={employee:tn(r)};console.log("Headcount column detection:",{rosterEmployeeCol:l.employee,rosterTerminationCol:l.termination,payrollEmployeeCol:c.employee,rosterHeaders:s.slice(0,5),payrollHeaders:r.slice(0,5)});let i=new Set,p=new Set;for(let d=a.startIndex;d<e.length;d+=1){let h=e[d],y=l.employee>=0?he(h[l.employee]):"";Zt(y)||l.termination>=0&&he(h[l.termination])||i.add(y.toLowerCase())}for(let d=o.startIndex;d<t.length;d+=1){let h=t[d],y=c.employee>=0?he(h[c.employee]):"";Zt(y)||p.add(y.toLowerCase())}n.rosterCount=i.size,n.payrollCount=p.size,n.difference=n.payrollCount-n.rosterCount,console.log("Headcount results:",{rosterCount:n.rosterCount,payrollCount:n.payrollCount,difference:n.difference});let g=[...i].filter(d=>!p.has(d)),u=[...p].filter(d=>!i.has(d));return n.mismatches=[...g.map(d=>`In roster, missing in PTO_Data: ${d}`),...u.map(d=>`In PTO_Data, missing in roster: ${d}`)],{roster:n}}function en(e){if(!Array.isArray(e)||!e.length)return{headers:[],startIndex:1};let t=e.findIndex((o=[])=>o.some(s=>he(s).toLowerCase().includes("employee"))),n=t===-1?0:t;return{headers:(e[n]||[]).map(o=>he(o).toLowerCase()),startIndex:n+1}}function tn(e=[]){let t=-1,n=-1;return e.forEach((a,o)=>{let s=a.toLowerCase();if(!s.includes("employee"))return;let r=1;s.includes("name")?r=4:s.includes("id")?r=2:r=3,r>n&&(n=r,t=o)}),t}function he(e){return e==null?"":String(e).trim()}async function yn(e=null){let t=async n=>{let a=n.workbook.worksheets.getItem("PTO_Data"),o=n.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster"),r=n.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary"),l=n.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary"),c=a.getUsedRangeOrNullObject();c.load("values"),o.load("isNullObject"),s.load("isNullObject"),r.load("isNullObject"),l.load("isNullObject"),await n.sync();let i=c.isNullObject?[]:c.values||[];if(!i.length)return;let p=(i[0]||[]).map(x=>q(x)),g=p.findIndex(x=>x.includes("employee")&&x.includes("name")),u=g>=0?g:0,d=M(p,["accrual rate"]),h=M(p,["carry over","carryover"]),y=p.findIndex(x=>x.includes("ytd")&&(x.includes("accrued")||x.includes("accrual"))),b=p.findIndex(x=>x.includes("ytd")&&x.includes("used")),f=M(p,["balance","current balance","pto balance"]);console.log("[PTO Analysis] PTO_Data headers:",p),console.log("[PTO Analysis] Column indices found:",{employee:u,accrualRate:d,carryOver:h,ytdAccrued:y,ytdUsed:b,balance:f}),b>=0?console.log(`[PTO Analysis] YTD Used column: "${p[b]}" at index ${b}`):console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:",p);let v=i.slice(1).map(x=>he(x[u])).filter(x=>x&&!x.toLowerCase().includes("total")),w=new Map;i.slice(1).forEach(x=>{let G=q(x[u]);!G||G.includes("total")||w.set(G,x)});let R=new Map;if(s.isNullObject)console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");else{let x=s.getUsedRangeOrNullObject();x.load("values"),await n.sync();let G=x.isNullObject?[]:x.values||[];if(G.length){let H=(G[0]||[]).map(T=>q(T));console.log("[PTO Analysis] SS_Employee_Roster headers:",H);let F=H.findIndex(T=>T.includes("employee")&&T.includes("name"));F<0&&(F=H.findIndex(T=>T==="employee"||T==="name"||T==="full name"));let U=H.findIndex(T=>T.includes("department"));console.log(`[PTO Analysis] Roster column indices - Name: ${F}, Dept: ${U}`),F>=0&&U>=0?(G.slice(1).forEach(T=>{let oe=q(T[F]),ce=he(T[U]);oe&&R.set(oe,ce)}),console.log(`[PTO Analysis] Built roster map with ${R.size} employees`)):console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster")}}let _=new Map;if(!r.isNullObject){let x=r.getUsedRangeOrNullObject();x.load("values"),await n.sync();let G=x.isNullObject?[]:x.values||[];if(G.length){let H=(G[0]||[]).map(U=>q(U)),F={payrollDate:M(H,["payroll date"]),employee:M(H,["employee"]),category:M(H,["payroll category","category"]),amount:M(H,["amount","gross salary","gross_salary","earnings"])};F.employee>=0&&F.category>=0&&F.amount>=0&&G.slice(1).forEach(U=>{let T=q(U[F.employee]);if(!T)return;let oe=q(U[F.category]);if(!oe.includes("regular")||!oe.includes("earn"))return;let ce=Number(U[F.amount])||0;if(!ce)return;let Ee=ma(U[F.payrollDate]),xe=_.get(T);(!xe||Ee!=null&&Ee>xe.timestamp)&&_.set(T,{payRate:ce/80,timestamp:Ee})})}}let $=new Map;if(!l.isNullObject){let x=l.getUsedRangeOrNullObject();x.load("values"),await n.sync();let G=x.isNullObject?[]:x.values||[];if(G.length>1){let H=(G[0]||[]).map(T=>q(T)),F=H.findIndex(T=>T.includes("employee")&&T.includes("name")),U=M(H,["liability amount","liability","accrued pto"]);F>=0&&U>=0&&G.slice(1).forEach(T=>{let oe=q(T[F]);if(!oe)return;let ce=Number(T[U])||0;$.set(oe,ce)})}}let j=ie(E.payrollDate)||"",D=[],V=[],S=v.map((x,G)=>{var pt,ft,gt,mt,ht,yt,vt;let H=q(x),F=R.get(H)||"",U=(ft=(pt=_.get(H))==null?void 0:pt.payRate)!=null?ft:"",T=w.get(H),oe=T&&d>=0&&(gt=T[d])!=null?gt:"",ce=T&&h>=0&&(mt=T[h])!=null?mt:"",Ee=T&&y>=0&&(ht=T[y])!=null?ht:"",xe=T&&b>=0&&(yt=T[b])!=null?yt:"";(H.includes("avalos")||H.includes("sarah"))&&console.log(`[PTO Debug] ${x}:`,{ytdUsedIdx:b,rawValue:T?T[b]:"no dataRow",ytdUsed:xe,fullRow:T});let Ue=T&&f>=0&&Number(T[f])||0,dt=G+2;!U&&typeof U!="number"&&D.push({name:x,rowIndex:dt}),F||V.push({name:x,rowIndex:dt});let Ge=typeof U=="number"&&Ue?Ue*U:0,ut=(vt=$.get(H))!=null?vt:0,vn=(typeof Ge=="number"?Ge:0)-ut;return[j,x,F,U,oe,ce,Ee,xe,Ue,Ge,ut,vn]});Q.missingPayRates=D.filter(x=>!Q.ignoredMissingPayRates.has(x.name)),Q.missingDepartments=V,console.log(`[PTO Analysis] Data quality: ${D.length} missing pay rates, ${V.length} missing departments`);let I=[["Analysis Date","Employee Name","Department","Pay Rate","Accrual Rate","Carry Over","YTD Accrued","YTD Used","Balance","Liability Amount","Accrued PTO $ [Prior Period]","Change"],...S],O=o.isNullObject?n.workbook.worksheets.add("PTO_Analysis"):o,m=O.getUsedRangeOrNullObject();m.load("address"),await n.sync(),m.isNullObject||m.clear();let P=I[0].length,A=I.length,N=S.length,Y=O.getRangeByIndexes(0,0,A,P);Y.values=I;let le=O.getRangeByIndexes(0,0,1,P);ot(le),N>0&&(Gt(O,0,N),pe(O,3,N),ke(O,4,N),ke(O,5,N),ke(O,6,N),ke(O,7,N),ke(O,8,N),pe(O,9,N),pe(O,10,N),pe(O,11,N,!0)),Y.format.autofitColumns(),O.getRange("A1").select(),await n.sync()};K()&&(e?await t(e):await Excel.run(t))}function xa(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let a=String(n);return/[",\n]/.test(a)?`"${a.replace(/"/g,'""')}"`:a}).join(",")).join(`
`)}function Ca(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),a=URL.createObjectURL(n),o=document.createElement("a");o.href=a,o.download=e,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(a),1e3)}function nn(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=hn(),n=document.getElementById("step-notes-input"),a=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!a;let o=document.getElementById("headcount-notes-hint");o&&(o.textContent=t?"Please document outstanding differences before signing off.":"")}function an(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(me)?t.slice(me.length).replace(/^\s+/,""):t.replace(new RegExp(`^${me}\\s*`,"i"),"").trimStart(),a=me+(n?`
${n}`:"");e.value!==a&&(e.value=a),ee(2,"notes",e.value)}function _a(){let e=document.getElementById("step-notes-input");e&&e.addEventListener("input",()=>{if(!L.skipAnalysis)return;let t=e.value||"";if(!t.startsWith(me)){let n=t.replace(me,"").trimStart();e.value=me+(n?`
${n}`:"")}ee(2,"notes",e.value)})}})();
//# sourceMappingURL=app.bundle.js.map
