const STORAGE_KEY='workInbox_briefings_v1', TODAY_KEY='workInbox_today_v1', TICKS_KEY='workInbox_ticks_v1';

const SEED={"date":"Tuesday 2 June 2026","subtitle":"Clear backlog before annual reviews","context":"Back from leave. Key absences: James Salas Guillen (returns Thursday), Sarah Rowles (returns Friday). FlexPoints budget closes end of June - allocation still pending. Three P1 tickets unresolved from last week. DSE go-live prep needs sign-off from Simon before Friday.","prioritiesToday":[{"title":"DSE go-live sign-off -- get confirmation from Simon","source":"H&S ROADMAP 01/06","actions":["[01 Jun 2026] Raised with Simon at roadmap meeting. Awaiting written confirmation.","[TODO] Chase Simon at Wednesday 1-1 if not received before then."]},{"title":"FlexPoints allocation -- Chemistry and Holiday Records","source":"INBOX 2026-06-01 09:15","actions":["[01 Jun 2026] Chemistry confirmed. Holiday Records quote still outstanding from AG.","[AWAITING] Mike West at AG - quote for Holiday Records module.","[TODO] Confirm remaining balance and allocate once quote received."]},{"title":"P1 incident -- OSM data feed failure","source":"INCIDENT LOG","actions":["[28 May 2026] OSM feed failed overnight. Asta managing manually.","[MONITOR] No resolution from supplier yet. Escalate if not resolved by COB today."]}],"prioritiesWeek":[{"title":"SharePoint documentation -- written response and team alignment","source":"SK 1-1 08/06","actions":["[08 Jun 2026] Simon flagged guidance not updated when drive moved.","[TODO] Written response confirming current guidance.","[TODO] Schedule team alignment meeting."]},{"title":"HWP archived users -- one-off DSC upload cover","source":"JAMES HANDOVER","actions":["[TODO] One-off DSC upload cover while James on leave.","[MONITOR] Two open HWP tickets with Gail Miller."]}],"fyi":[{"title":"Vacancy alert email retest -- case 68388326","sub":"Conor to pick up on return from leave.","badge":"Parked","badgeType":"gray"},{"title":"Iris enhancements on hold","sub":"Pending H&S funding approval. New AE contact: Michael Hanson.","badge":"H&S Roadmap","badgeType":"gray"}],"calToday":[{"time":"09:00","title":"FA Team Daily Catchup","sub":"Teams - Michael, Asta, James"},{"time":"14:00","title":"H&S Roadmap","sub":"James Salas Guillen - Teams"}],"calTomorrow":[{"time":"09:00","title":"FA Team Daily Catchup","sub":"Teams"},{"time":"11:00","title":"1-1 Simon Burford","sub":"Teams"}],"absences":["James Salas Guillen - returns Thursday 5 June","Sarah Rowles - returns Friday 6 June"]};

function getStore(){try{return JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}')}catch(e){return{}}}
function saveStore(d){localStorage.setItem(STORAGE_KEY,JSON.stringify(d))}
function getTicks(){try{return JSON.parse(localStorage.getItem(TICKS_KEY)||'{}')}catch(e){return{}}}
function saveTicks(t){localStorage.setItem(TICKS_KEY,JSON.stringify(t));scheduleStateSync()}

let currentData=null, currentKey=null;

// --- cross-machine tick sync (via Cloudflare Worker) ---
const STATE_WRITER_URL='https://cc-tasks-writer.kevinlelitte.workers.dev';
const TICKS_URL='https://raw.githubusercontent.com/begb0037admin/work-inbox/main/data/ticks.json';
let stateSyncTimer=null, stateSyncReady=false;
function scheduleStateSync(){
  if(!stateSyncReady) return;
  clearTimeout(stateSyncTimer);
  stateSyncTimer=setTimeout(pushTicks,1500);
}
async function pushTicks(){
  try{
    const doc={ticks:getTicks(),updated_at:new Date().toISOString()};
    const res=await fetch(STATE_WRITER_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({target:'inbox-state',message:'tick sync',doc:doc})});
    const out=await res.json().catch(()=>({}));
    if(!res.ok||!out.ok) throw new Error(out.error||('HTTP '+res.status));
    console.log('Ticks synced to GitHub');
  }catch(e){
    console.warn('Tick sync failed',e);
    wiNotify('Ticks not synced - local only. '+(e.message||''));
  }
}
function wiNotify(msg){
  let el=document.getElementById('wiToast');
  if(!el){
    el=document.createElement('div');
    el.id='wiToast';
    el.style.cssText='position:fixed;bottom:18px;right:18px;max-width:340px;background:#1a2740;color:#fff;padding:12px 16px;border-radius:10px;font-size:13px;z-index:999;box-shadow:0 8px 24px rgba(0,0,0,.25);line-height:1.45';
    document.body.appendChild(el);
  }
  el.textContent=msg;
  el.style.display='block';
  clearTimeout(el._t);
  el._t=setTimeout(()=>{el.style.display='none'},7000);
}
async function loadRemoteTicks(){
  try{
    const res=await fetch(TICKS_URL+'?t='+Date.now());
    if(res.ok){
      const st=await res.json();
      if(st&&typeof st==='object'&&st.ticks) localStorage.setItem(TICKS_KEY,JSON.stringify(st.ticks));
    }
  }catch(e){console.warn('Remote ticks unavailable',e);}
  stateSyncReady=true;
}


function toggleImport(){
  document.getElementById('importPanel').classList.toggle('visible');
  document.getElementById('archivePanel').classList.remove('visible');
  document.getElementById('importError').style.display='none';
}
/* Show/Hide Done: showingDoneItems is the single source of truth for whether
   done items render hidden. It must ONLY ever change inside toggleShowDone(),
   fired only by the Show/Hide Done button's own onclick. Every render path
   (renderItems/renderPriorityCards) reads this same variable when building
   card HTML, so a full re-render triggered by anything else -- a tick, a
   drag/drop, an accidental micro-drag from a plain click on a draggable card
   (dragstart fires from HTML5 drag-and-drop on the slightest pointer
   movement, even during what feels like a simple click) -- always re-applies
   the current toggle state instead of silently reverting to "all visible".
   Previously this function mutated .card-hidden on the live DOM directly;
   that mutation was wiped out by the next renderBriefing() call from any
   other interaction, which is exactly the "click a card and hidden done
   items reappear" bug. Fixed 12 Aug 2026. */
let showingDoneItems=false;
function toggleShowDone(){
  const btn=document.getElementById('btn-show-done');
  showingDoneItems=!showingDoneItems;
  if(btn) btn.textContent=showingDoneItems?'Hide done':'Show done';
  if(window._wipData&&window._wipKey) renderBriefing(window._wipData,window._wipKey);
}
function toggleArchive(){
  const p=document.getElementById('archivePanel');
  p.classList.toggle('visible');
  document.getElementById('importPanel').classList.remove('visible');
  if(p.classList.contains('visible')) renderArchiveList();
}
function getArchiveData(){
  const store=getStore(), ticks=getTicks();
  return Object.keys(store).sort().reverse().map(k=>{
    const d=store[k];
    const dateStr=d.date||k.replace(/_/g,' ');
    return {key:k, data:d, dateStr:dateStr, ticks:ticks};
  });
}
function getSectionLabel(s){
  return {urgent:'Urgent',needs:'Needs',fyi:'FYI',low:'Low',priorities:'Priority'}[s]||s;
}
function getAllItemsForDay(data){
  const sections=['priorities','urgent','needs','fyi','low'];
  const result=[];
  sections.forEach(s=>{
    (data[s]||[]).forEach((item,i)=>{result.push({section:s,index:i,title:item.title||item.text||'(untitled)'});});
  });
  return result;
}
function renderArchiveList(){
  const entries=getArchiveData();
  const el=document.getElementById('archiveList');
  const actionsEl=document.getElementById('archiveActions');
  if(!entries.length){
    el.innerHTML='<div class="archive-no-items">No briefings stored yet.</div>';
    actionsEl.style.display='none';
    return;
  }
  actionsEl.style.display='flex';
  const ticks=getTicks();
  el.innerHTML=entries.map((e,di)=>{
    const items=getAllItemsForDay(e.data);
    const tickedCount=items.filter(it=>ticks[e.key+'_'+it.section+'_'+it.index]).length;
    const itemsHtml=items.map(it=>{
      const t=!!ticks[e.key+'_'+it.section+'_'+it.index];
      return `<div class="archive-item-row"><span class="${t?'archive-item-tick':'archive-item-untick'}">${t?'✓':'○'}</span><span class="archive-item-section">${getSectionLabel(it.section)}</span><span>${it.title}</span></div>`;
    }).join('');
    return `<div class="archive-day">
      <div class="archive-day-header" onclick="toggleArchiveDay(${di})">
        <div><div class="archive-day-date">${e.dateStr}</div><div class="archive-day-meta">${items.length} items · ${tickedCount} done</div></div>
        <span class="archive-day-arrow" id="arch-arrow-${di}">–</span>
      </div>
      <div class="archive-day-items" id="arch-items-${di}">${itemsHtml}</div>
    </div>`;
  }).join('');
}
function toggleArchiveDay(i){
  const el=document.getElementById('arch-items-'+i);
  const arrow=document.getElementById('arch-arrow-'+i);
  const open=el.classList.toggle('open');
  arrow.textContent=open?'▲':'–';
}
function exportArchiveMd(){
  const entries=getArchiveData();
  const ticks=getTicks();
  let md='# Inbox Briefing Archive\n\nExported: '+new Date().toLocaleString('en-GB')+'\n\n---\n\n';
  entries.forEach(e=>{
    md+='## '+e.dateStr+'\n\n';
    const sections=['priorities','urgent','needs','fyi','low'];
    const labels={priorities:'Priority Actions',urgent:'Urgent',needs:'Needs Response',fyi:'FYI',low:'Low Priority'};
    sections.forEach(s=>{
      const items=e.data[s]||[];
      if(!items.length) return;
      md+='### '+labels[s]+'\n\n';
      items.forEach((item,i)=>{
        const t=!!ticks[e.key+'_'+s+'_'+i];
        const title=item.title||item.text||'(untitled)';
        md+=(t?'- [x] ':'- [ ] ')+title+'\n';
        const sub=item.sub||item.notes||'';
        if(sub) md+='  > '+sub.replace(/<[^>]+>/g,'').replace(/\s+/g,' ').trim()+'\n';
      });
      md+='\n';
    });
    md+='---\n\n';
  });
  const blob=new Blob([md],{type:'text/markdown'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  a.download='inbox-archive-'+new Date().toISOString().slice(0,10)+'.md';
  a.click();
  URL.revokeObjectURL(url);
}
function purgeOldTicks(){
  const days=parseInt(document.getElementById('purgeDays').value)||30;
  if(!confirm('Purge all briefings and ticks older than '+days+' days? This cannot be undone.')) return;
  const cutoff=new Date();
  cutoff.setDate(cutoff.getDate()-days);
  const store=getStore(), ticks=getTicks();
  let purgedCount=0;
  Object.keys(store).forEach(k=>{
    const d=store[k];
    const dateStr=d.date||'';
    const parsed=new Date(dateStr);
    const age=isNaN(parsed)?null:parsed;
    if(!age||age<cutoff){
      delete store[k];
      Object.keys(ticks).filter(tk=>tk.startsWith(k+'_')).forEach(tk=>{delete ticks[tk];purgedCount++;});
    }
  });
  saveStore(store);
  saveTicks(ticks);
  renderArchiveList();
  alert('Purge complete.');
}
function loadFromImport(){
  const raw=document.getElementById('importInput').value.trim();
  const err=document.getElementById('importError'); err.style.display='none';
  let data; try{data=JSON.parse(raw)}catch(e){err.style.display='block';return;}
  if(!data.date){err.textContent='Missing "date" field.';err.style.display='block';return;}
  const key=data.date.replace(/[^a-zA-Z0-9]/g,'_');
  const store=getStore(); store[key]=data; saveStore(store);
  localStorage.setItem(TODAY_KEY,key);
  document.getElementById('importPanel').classList.remove('visible');
  document.getElementById('importInput').value='';
  renderBriefing(data,key);
}
function clearToday(){
  if(!confirm('Clear today and return to seed data?')) return;
  localStorage.removeItem(TODAY_KEY);
  renderBriefing(SEED,'seed');
}
// Tick storage key -- day-independent and position-independent for any id
// carrying a stable identity (the 'eid_'/'id_' prefix _priGetKey() below
// produces), falling back to the old calendar-day-scoped key only for the
// rare item with no stable id at all. Fixed 20 Aug 2026 as a prerequisite
// for rebuilding Phase 3.9 (server-side scroll-out persistence) -- without
// this, any item Phase 3.9 carries across a day boundary would resurrect
// as undone every day regardless of its real done state, exactly the "mark
// done, refresh, it comes back" incident from 17 Aug 2026 (see
// wi-tick-resurrection-incident-17aug.md in the `drew` repo). That incident
// fix shipped once already and worked correctly in production before being
// swept up in an unrelated same-night full revert; reused verbatim here
// rather than reinvented, since it was never the part that was wrong.
function _tickStorageKey(id){
  return (typeof id==='string'&&(id.indexOf('eid_')===0||id.indexOf('id_')===0)) ? id : (currentKey+'_'+id);
}
function toggleTick(id){
  const ticks=getTicks(), k=_tickStorageKey(id);
  ticks[k]=!ticks[k]; saveTicks(ticks);
  const cb=document.getElementById('cb_'+id);
  const item=document.getElementById('item_'+id);
  const prow=document.getElementById('prow_'+id);
  if(cb){
    if(cb.classList.contains('card-done-btn')) cb.classList.toggle('done',ticks[k]);
    else cb.classList.toggle('checked',ticks[k]);
  }
  if(item){
    const wrapper=item.closest('.card-link');
    if(ticks[k]){
      item.classList.add('done');
      const titleEl=item.querySelector('.card-ph-title');
      if(titleEl) titleEl.classList.add('done');
      if(!showingDoneItems) item.classList.add('card-hidden');
      else item.classList.remove('card-hidden');
      if(wrapper){
        wrapper.classList.add('done');
        if(!showingDoneItems) wrapper.classList.add('card-hidden');
        else wrapper.classList.remove('card-hidden');
      }
    } else {
      item.classList.remove('done','card-hidden');
      const titleEl=item.querySelector('.card-ph-title');
      if(titleEl) titleEl.classList.remove('done');
      if(wrapper) wrapper.classList.remove('done','card-hidden');
    }
  }
  if(prow){
    if(ticks[k]){ prow.classList.add('done'); prow.style.display=''; }
    else { prow.classList.remove('done'); prow.style.display=''; }
  }
}
function openEmail(entryId,ev){
  if(ev){ev.preventDefault();ev.stopPropagation();}
  window.location.href='openmail://'+entryId+'/';
}
function isTicked(id){if(!currentKey) return false; return !!getTicks()[_tickStorageKey(id)];}
function badge(text,type){return text?`<span class="badge badge-${type||'gray'}">${text}</span>`:''}

function escapeHtml(text){
  if(text===undefined||text===null) return '';
  return String(text)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function sanitizeSub(text){
  if(!text) return '';
  return text
    .replace(/<https?:\/\/[^>]*>/gi, '')
    .replace(/<https?:\/\/[^\s<>]*/gi, '')
    .replace(/<(?!\/?strong\b)[^>]*>/gi, '')
    .replace(/<(?!\/?strong\b)[^>]*$/gi, '')
    .replace(/\r\n/g,' ').replace(/\r/g,' ').replace(/\n/g,' ');
}

function renderItems(items,cls){
  if(!items||!items.length) return '<div class="no-items">None today.</div>';
  return items.map((item,i)=>{
    const id=cls+'_'+i, ticked=isTicked(id);
    const hasLink=item.entry_id&&item.entry_id.length>0;
    const dragAttrs=`draggable="true" ondragstart="emailCardDragStart(event,'${cls}',${i})" ondragend="emailCardDragEnd(event)"`;
    const hiddenCls=(ticked&&!showingDoneItems)?' card-hidden':'';
    const cardHtml=`<div class="card${ticked?' done':''}${hiddenCls}" id="item_${id}"><div class="cb-wrap"><div class="cb${ticked?' checked':''}" id="cb_${id}" onclick="toggleTick('${id}');event.stopPropagation()"></div></div><div class="card-accent ac-${cls==="urgent"?"r":cls==="needs"?"o":cls==="fyi"?"b":"g"}"></div><div class="card-body"><div class="card-title">${item.title} ${badge(item.badge,item.badgeType)}${(()=>{if(!item.received)return '';const c=new Date();c.setDate(c.getDate()-4);c.setHours(0,0,0,0);return new Date(item.received+'T12:00:00')>=c?badge('NEW','green'):'';})()}</div>${item.sub?`<div class="card-sub">${sanitizeSub(item.sub)}</div>`:''}</div><div class="card-date">${item.received||''}</div></div>`;
    return hasLink?`<a class="card-link${ticked?' done':''}${hiddenCls}" href="javascript:void(0)" onclick="openEmail('${item.entry_id}',event)" ${dragAttrs}>${cardHtml}</a>`:`<div ${dragAttrs}>${cardHtml}</div>`;
  }).join('');
}

function renderSidebarCal(items, containerId){
  const el=document.getElementById(containerId);
  if(!el) return;
  if(!items||!items.length){el.innerHTML='<div style="padding:4px 18px 8px;font-size:11px;color:rgba(255,255,255,0.3);font-style:italic">None</div>';return;}
  el.innerHTML=items.map((c,i)=>`<div class="cal-item${i===0?' active':''}">${c.time?`<div class="cal-time">${c.time}</div>`:''}<div class="cal-title">${c.title}</div>${c.sub?`<div class="cal-sub">${c.sub}</div>`:''}${c.alert?`<div class="cal-alert">⚠ ${c.alert}</div>`:''}</div>`).join('');
}

function renderMainCal(data){
  const el=document.getElementById('contextBar');
  if(!el) return;
  const now=new Date();
  const nowMins=now.getHours()*60+now.getMinutes();
  const todayDate=now.getDate(), todayMonth=now.getMonth(), todayYear=now.getFullYear();

  function parseTimeMins(t){
    if(!t) return -1;
    const p=t.split(':');
    return p.length<2?-1:parseInt(p[0])*60+parseInt(p[1]);
  }

  function renderBlock(items,headerHtml,isToday){
    if(!items||!items.length) return `<div class="main-cal-block"><div class="main-cal-block-header">${headerHtml}</div><div class="main-cal-none">No meetings</div></div>`;
    let nextFound=false;
    const rows=items.map(c=>{
      const mins=parseTimeMins(c.time);
      const isPast=isToday&&mins>=0&&mins<nowMins;
      const isNext=isToday&&!isPast&&!nextFound&&mins>=nowMins;
      if(isNext) nextFound=true;
      const cls=isPast?' past':isNext?' next':'';
      return `<div class="main-cal-item${cls}"><span class="main-cal-time">${escapeHtml(c.time||'')}</span><div><div class="main-cal-title">${escapeHtml(c.title)}</div>${c.sub?`<div class="main-cal-sub">${escapeHtml(c.sub)}</div>`:''}${c.summary?`<div class="main-cal-summary">${escapeHtml(c.summary)}</div>`:''}</div></div>`;
    }).join('');
    return `<div class="main-cal-block"><div class="main-cal-block-header">${headerHtml}</div>${rows}</div>`;
  }

  function renderMiniCal(){
    const monthName=now.toLocaleDateString('en-GB',{month:'long',year:'numeric'});
    const firstDay=new Date(todayYear,todayMonth,1);
    const daysInMonth=new Date(todayYear,todayMonth+1,0).getDate();
    let startDow=firstDay.getDay()-1; if(startDow<0) startDow=6;
    const tom=new Date(now); tom.setDate(tom.getDate()+1);
    const tomDate=tom.getDate();
    const hasTodayMtg=data.calToday&&data.calToday.length>0;
    const hasTomMtg=data.calTomorrow&&data.calTomorrow.length>0;
    const dayNames=['M','T','W','T','F','S','S'];
    let cells=dayNames.map(d=>`<div class="mini-cal-day-name">${d}</div>`).join('');
    for(let i=0;i<startDow;i++) cells+='<div class="mini-cal-day other-month"></div>';
    for(let d=1;d<=daysInMonth;d++){
      const isT=d===todayDate, isTom=d===tomDate;
      const hasMtg=(isT&&hasTodayMtg)||(isTom&&hasTomMtg);
      const cls='mini-cal-day'+(isT?' today':hasMtg?' has-meeting':'');
      cells+=`<div class="${cls}">${d}</div>`;
    }
    return `<div class="main-cal-block"><div class="main-cal-block-header">${monthName}</div><div class="mini-cal-grid">${cells}</div></div>`;
  }

  const todayHeader='Today &mdash; '+now.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long'});
  const tom=new Date(now); tom.setDate(tom.getDate()+1);
  const tomHeader='Tomorrow &mdash; '+tom.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long'});
  el.innerHTML=`<div class="main-cal-panel">${renderBlock(data.calToday,todayHeader,true)}${renderBlock(data.calTomorrow,tomHeader,false)}${renderMiniCal()}</div>`;
}

function togglePriCard(i){
  const body=document.getElementById('pribody_'+i);
  const arrow=document.getElementById('priarrow_'+i);
  const isOpen=body.style.display==='block';
  body.style.display=isOpen?'none':'block';
  arrow.textContent=isOpen?'–':'▲';
}

// Priority drag-and-drop helpers
let _priDragState=null,_priDragEl=null,_priDragDropped=false;
// Added 20 Aug 2026, drag-and-drop architecture rework: origin-position
// memory (for an O(1) revert-on-cancel) and the actual drop-target section
// (so priDragEnd can tell a same-zone reorder from a real cross-zone move
// without re-deriving it from localStorage).
let _priOriginParent=null,_priOriginNextSibling=null,_priDropTargetSec=null;
function _priGetLegacyTitleKey(p){return(p.title||p.text||p.subject||'').toLowerCase().replace(/[^a-z0-9]/g,'').substring(0,40)||'item';}
// Prefer a stable identifier (entry_id for email-sourced items, id for
// priorities items) over the item's display title -- fixed 12 Aug 2026,
// closing out the "cards vanish on move" bug this same file's HANDOVER.md
// entry flagged but didn't fix yet. Two genuinely different real items can
// share an identical title (a meeting reschedule notice re-using the
// original invite's subject line, a recurring standing-meeting subject,
// etc.) -- confirmed live against Kevin's own data/briefing.json: an
// "Incident Reporting PUG" email in Needs Response (entry_id
// ...67A8967C720000) and a separately-triaged "Incident Reporting PUG" item
// in FYI/Parked (a different entry_id) share exact title text but are two
// different emails. applyPriOverrides() below used to dedupe purely on this
// title text via _seen, which silently and permanently dropped the second
// one it encountered -- not just from its own section, from every section,
// since the drop happens before the override/section assignment ever runs.
// That's exactly "I moved a card and it vanished, not in the destination,
// not back in the source": if the dragged item's title collided with any
// item processed earlier in the pt/ptom/pw/fyi/urgent/needs merge order, no
// override could ever rescue it. Falls back to the old title-slug only when
// an item genuinely has no stable id (matches prior behaviour for that
// case; verified live that all six default arrays currently have 100%
// entry_id/id coverage, so this fallback is a safety net, not the common
// path).
function _priGetKey(p){
  if(p.entry_id) return 'eid_'+p.entry_id;
  if(p.entryId) return 'eid_'+p.entryId;
  if(p.id) return 'id_'+p.id;
  return _priGetLegacyTitleKey(p);
}
function _priGetOverrides(){try{return JSON.parse(localStorage.getItem('workInbox_priOverrides_v1')||'{}');}catch(e){return{};}}
function _priSetOverride(key,sec){const o=_priGetOverrides();o[key]=sec;localStorage.setItem('workInbox_priOverrides_v1',JSON.stringify(o));}
function _priGetOrder(){try{return JSON.parse(localStorage.getItem('workInbox_priOrder_v1')||'{}');}catch(e){return{};}}
function _priSetOrder(pt,ptom,pw,pfyi,ur,nr){localStorage.setItem('workInbox_priOrder_v1',JSON.stringify({pt,ptom:ptom||[],pw,pfyi:pfyi||[],ur:ur||[],nr:nr||[]}));}
function _getCustomPri(){try{return JSON.parse(localStorage.getItem('workInbox_customPri_v1')||'[]');}catch(e){return[];}}
function _saveCustomPri(arr){localStorage.setItem('workInbox_customPri_v1',JSON.stringify(arr));}
function _addEmailCardToPriority(item,cls,sec){const arr=_getCustomPri();const priKey=_priGetKey(item);if(arr.findIndex(x=>x._priKey===priKey)<0){arr.push({...item,_priKey:priKey,_dfSec:sec,_cls:cls});_saveCustomPri(arr);}_priSetOverride(priKey,sec);}

function applyPriOverrides(data){
  const all=[...(data.prioritiesToday||[]).map(p=>({...p,_dfSec:'pt'})),...(data.prioritiesTomorrow||[]).map(p=>({...p,_dfSec:'ptom'})),...(data.prioritiesWeek||[]).map(p=>({...p,_dfSec:'pw'})),...(data.fyi||[]).map(p=>({...p,text:p.title,_dfSec:'pfyi'})),...(data.urgent||[]).map(p=>({...p,text:p.title,_dfSec:'ur'})),...(data.needs||[]).map(p=>({...p,text:p.title,_dfSec:'nr'})),..._getCustomPri()];
  const ovr=_priGetOverrides(),ord=_priGetOrder(),secs={pt:[],ptom:[],pw:[],pfyi:[],ur:[],nr:[]};
  const validSecs=['pt','ptom','pw','pfyi','ur','nr'];
  const _seen=new Set();
  for(const item of all){
    const k=_priGetKey(item);
    if(_seen.has(k))continue;
    _seen.add(k);
    // Overrides saved by a drag before this fix are keyed by the old
    // title-only slug -- fall back to that so Kevin's existing manual
    // section placements keep applying even though new drags now save
    // under the new stable-id key.
    const legacyKey=_priGetLegacyTitleKey(item);
    const s=ovr[k]||ovr[legacyKey]||item._dfSec;
    (validSecs.includes(s)?secs[s]:secs.pw).push({...item,_priKey:k});
  }
  // Newest-first-insertion requirement, 20 Aug 2026 -- an item with no
  // recorded position in ord[s] (a fresh Outlook-pull arrival into this
  // section, or a section it's never been manually reordered within) used
  // to sort with om[key]??999, i.e. AFTER every item that does have a
  // recorded index -- new items landed at the BOTTOM. Kevin wants new
  // arrivals at the TOP instead, while every item that already has a real
  // recorded index keeps that exact relative order (om[key]??999 among
  // themselves is unaffected by this change -- only where the "no index"
  // bucket sorts relative to them changes, from after to before). Using -1
  // (always less than any real 0-based index) instead of 999 achieves this;
  // Array.prototype.sort is spec-guaranteed stable, so multiple new items
  // arriving in the same pull keep their relative merge order among
  // themselves at the top, they don't get shuffled.
  for(const s of validSecs){if(ord[s]&&ord[s].length){const om={};ord[s].forEach((k,i)=>om[k]=i);secs[s].sort((a,b)=>(om[a._priKey]??-1)-(om[b._priKey]??-1));}}
  return secs;
}

function priDragStart(e,sec,priKey){
  _priDragState={sec,priKey};
  _priDragEl=e.currentTarget;
  _priDragDropped=false;
  _priDropTargetSec=null;
  // Remember exactly where this card came from (parent zone + the sibling
  // it sat before) -- fixed 20 Aug 2026, drag-and-drop architecture rework.
  // priCardDragOver/priZoneDragOver live-move this real DOM node as a drag
  // preview on every dragover; if the drag ends without a real drop (a
  // cancelled drag, an accidental micro-drag from a plain click, dragging
  // outside the window), the ONLY thing that needs undoing is putting this
  // one node back -- a single insertBefore, not a full-board rebuild. See
  // priDragEnd() for where this is used.
  _priOriginParent=_priDragEl.parentElement;
  _priOriginNextSibling=_priDragEl.nextElementSibling;
  e.dataTransfer.effectAllowed='move';
  e.dataTransfer.setData('text/plain',priKey);
  // Consistent drag ghost across Chrome/Edge/Firefox -- fixed 12 Aug 2026
  // per wi-dragdrop-review-12aug.md Tier 1. Without setDragImage(), each
  // browser renders its own default ghost (a snapshot of the card taken at
  // dragstart) while priCardDragOver/priZoneDragOver simultaneously live-move
  // the *real* dragged node in the DOM for reorder preview -- so the user
  // sees a static browser ghost and a moving blue-bordered real card at the
  // same time, which is exactly the "card visually snaps around
  // unpredictably" complaint. Using an explicit offscreen clone, anchored to
  // the cursor's original grab point within the card, makes the ghost
  // identical across browsers and independent of the live reorder.
  try{
    const rect=_priDragEl.getBoundingClientRect();
    const ghost=_priDragEl.cloneNode(true);
    ghost.style.position='absolute';
    ghost.style.top='-9999px';
    ghost.style.left='-9999px';
    ghost.style.width=rect.width+'px';
    ghost.style.pointerEvents='none';
    ghost.style.margin='0';
    ghost.classList.remove('pri-dragging');
    document.body.appendChild(ghost);
    // Cleanup must run whether setDragImage succeeds or throws (e.g.
    // unsupported), otherwise a failed call leaks the offscreen clone --
    // Codex review finding, 12 Aug 2026: cleanup was previously registered
    // only after a successful call.
    try{
      e.dataTransfer.setDragImage(ghost,e.clientX-rect.left,e.clientY-rect.top);
    }finally{
      setTimeout(()=>{ghost.remove();},0);
    }
  }catch(err){/* setDragImage unsupported/element issue -- falls back to browser default ghost */}
  setTimeout(()=>{if(_priDragEl)_priDragEl.classList.add('pri-dragging');},0);
}
function priDragEnd(e){
  // Flush any pending, not-yet-applied reorder from the last dragover event
  // before reading DOM order for persistence -- otherwise a fast
  // drop/release right after the final dragover (before the next animation
  // frame fires) would persist the pre-preview position instead of what was
  // last previewed. Codex review finding, 12 Aug 2026.
  if(_priReorderRAF!=null){cancelAnimationFrame(_priReorderRAF);_priReorderRAF=null;}
  _priRunReorderFrame();
  _priPendingReorder=null;
  if(_priDragEl)_priDragEl.classList.remove('pri-dragging');
  document.querySelectorAll('.pri-drop-zone.pri-zone-active').forEach(el=>el.classList.remove('pri-zone-active'));

  // Drag-and-drop architecture rework, 20 Aug 2026 (see wi-dragdrop-review-
  // 12aug.md for the original review this closes out). priDragEnd used to
  // unconditionally call renderBriefing() -- a full innerHTML destroy/
  // rebuild of all six board sections -- on EVERY dragend, successful drop
  // or not. That was the structural root cause behind two bugs already
  // found and fixed the same day as the review (Show/Hide Done silently
  // resetting, cards vanishing on a title collision): any client-side UI
  // state not explicitly re-derived by the render functions gets silently
  // wiped by a full rebuild, and a plain click can trigger a dragstart/
  // dragend cycle without an intentional drag at all. Closing this
  // structurally (not just patching each symptom as it's found) means: no
  // full rebuild on ANY dragend any more. Two cases instead:
  if(!_priDragDropped){
    // No real drop -- put the dragged node back exactly where it started.
    // The live preview (priCardDragOver/priZoneDragOver) already moved the
    // real DOM node speculatively; since nothing was persisted, the only
    // correct recovery is undoing that one move. O(1), touches no other
    // card's DOM node, needs no data re-derivation.
    if(_priDragEl&&_priOriginParent){
      _priOriginParent.insertBefore(_priDragEl,_priOriginNextSibling);
    }
  }else{
    const fromSec=_priDragState&&_priDragState.sec;
    const toSec=_priDropTargetSec||fromSec;
    const crossZoneMove=!!(_priDragEl&&fromSec&&toSec&&fromSec!==toSec);
    // Newest-first-insertion requirement, 20 Aug 2026 -- a card entering a
    // DIFFERENT section for the first time (a cross-zone move -- "manual
    // move/re-categorization into that zone") lands at the TOP of the
    // destination zone, not wherever the live drag-preview
    // (priCardDragOver/priZoneDragOver) happened to leave it. Deliberately
    // scoped: this override runs ONLY when fromSec!==toSec, and it runs
    // BEFORE the sk[] DOM-order snapshot below so the forced top position
    // is what actually gets persisted to workInbox_priOrder_v1, not
    // overwritten by it. A same-zone reorder (fromSec===toSec, the far
    // more common case) never enters this branch at all -- the live
    // preview's exact drop position is captured completely unchanged,
    // per Kevin's hard constraint that manual drag-to-reorder within a
    // section must keep working exactly as before.
    if(crossZoneMove){
      const destZone=document.querySelector(`.pri-drop-zone[data-sec="${toSec}"]`);
      if(destZone) destZone.insertBefore(_priDragEl,destZone.firstElementChild);
    }
    const allSecs=['pt','ptom','pw','pfyi','ur','nr'];
    const sk={};
    allSecs.forEach(s=>{sk[s]=Array.from(document.querySelectorAll(`.pri-drop-zone[data-sec="${s}"] .card-ph`)).map(c=>c.dataset.prikey);});
    _priSetOrder(sk.pt,sk.ptom,sk.pw,sk.pfyi,sk.ur,sk.nr);

    if(crossZoneMove){
      // Real cross-zone move: the card's own markup depends on its section
      // (badge visibility, the section literal baked into its own drag
      // handlers, data-sec) so THIS ONE card needs fresh markup -- every
      // other card on the board is untouched. This is the only case that
      // still needs any markup regenerated; same-zone reorders need
      // nothing further; the live DOM already reflects the final order.
      const priKey=_priDragEl.dataset.prikey;
      const priorities=applyPriOverrides(window._wipData||{})[toSec]||[];
      const p=priorities.find(x=>x._priKey===priKey);
      if(p){
        const html=_priRenderOneCard(p,toSec);
        const tmp=document.createElement('div');
        tmp.innerHTML=html.trim();
        const newEl=tmp.firstElementChild;
        _priDragEl.replaceWith(newEl);
        _priDragEl=newEl;
      }
      _priUpdateZoneChrome(fromSec);
      _priUpdateZoneChrome(toSec);
    }
  }
  _runCardSearch();
  _priDragState=null;_priDragEl=null;_priDragDropped=false;
  _priOriginParent=null;_priOriginNextSibling=null;_priDropTargetSec=null;
}
let _emailDragData=null;
function emailCardDragStart(e,cls,idx){
  if(!window._wipData)return;
  const item=(window._wipData[cls]||[])[idx];
  if(!item)return;
  _emailDragData={item,cls,idx};
  e.dataTransfer.effectAllowed='move';
  e.dataTransfer.setData('text/plain','email:'+cls+'_'+idx);
}
function emailCardDragEnd(e){
  _emailDragData=null;
  document.querySelectorAll('.pri-drop-zone.pri-zone-active').forEach(el=>el.classList.remove('pri-zone-active'));
}
// --- rAF-batched reorder with hysteresis -- fixed 12 Aug 2026 per
// wi-dragdrop-review-12aug.md Tier 1. priCardDragOver/priZoneDragOver used
// to run a synchronous getBoundingClientRect() + DOM move on every single
// native dragover event, which browsers fire at high frequency while the
// pointer moves -- a per-event reflow+mutation that produces visible
// stutter on long lists (FYI/Parked has been measured elsewhere this
// session-cluster at up to ~290 displayed cards). Now each dragover just
// records the latest pointer/target and schedules a single
// requestAnimationFrame callback (a no-op if one is already pending) that
// performs at most one reorder decision + DOM mutation per frame. The
// midpoint-only before/after boundary check (e.clientY<rect.top+rect.height/2)
// also had no hysteresis, so hovering near a card's vertical centre could
// flip the insertion point back and forth every event -- a buffer band
// around the midpoint plus a per-target "last committed side" memory now
// only flips the decision once the pointer is clearly past the midpoint,
// not right at it.
//
// Pending state is a SINGLE directive (card-hover or zone-hover), not two
// independent ones -- Codex review finding, 12 Aug 2026: two independent
// pending records let one rAF callback apply a stale card-hover mutation
// followed by a newer zone-hover mutation in the same frame (two DOM
// mutations instead of one, and the stale one could momentarily win). Only
// the most recent dragover -- whichever type it was -- is kept.
let _priReorderRAF=null;
let _priPendingReorder=null; // {type:'card',target,zoneSec,clientY} | {type:'zone',zone} | null
const _priHysteresisFrac=0.15; // 15% of card height either side of midpoint
let _priLastBefore=new WeakMap(); // target el -> last committed before/after decision

function _priScheduleReorderFrame(){
  if(_priReorderRAF!=null)return;
  _priReorderRAF=requestAnimationFrame(_priRunReorderFrame);
}
function _priRunReorderFrame(){
  _priReorderRAF=null;
  const pending=_priPendingReorder;
  _priPendingReorder=null;
  if(!_priDragState||!_priDragEl||!pending)return;
  if(pending.type==='card'){
    const{target,zoneSec,clientY}=pending;
    if(target!==_priDragEl&&target.isConnected){
      const zone=document.querySelector(`.pri-drop-zone[data-sec="${zoneSec}"]`);
      if(zone){
        const r=target.getBoundingClientRect();
        const mid=r.top+r.height/2;
        const buffer=r.height*_priHysteresisFrac;
        let before;
        if(clientY<mid-buffer)before=true;
        else if(clientY>mid+buffer)before=false;
        else before=_priLastBefore.has(target)?_priLastBefore.get(target):(clientY<mid);
        _priLastBefore.set(target,before);
        zone.insertBefore(_priDragEl,before?target:target.nextSibling);
      }
    }
  }else if(pending.type==='zone'){
    const{zone}=pending;
    if(zone&&_priDragEl&&!zone.contains(_priDragEl))zone.appendChild(_priDragEl);
  }
}
function priCardDragOver(e,sec,priKey){
  if(!_priDragState&&!_emailDragData)return;
  e.preventDefault();e.stopPropagation();e.dataTransfer.dropEffect='move';
  if(_priDragState&&_priDragEl){
    const target=e.currentTarget;
    if(target===_priDragEl)return;
    _priPendingReorder={type:'card',target,zoneSec:sec,clientY:e.clientY};
    _priScheduleReorderFrame();
  }
}
function priCardDragLeave(e,priKey){
  // Clear a pending reorder if it targets the card being left -- otherwise
  // a deferred frame could still reorder against a card the pointer is no
  // longer hovering (Codex review, pass 1). Guard against the parent->child
  // dragleave/dragenter pair the browser fires when the pointer moves onto
  // a NESTED element inside the same card (title text, action buttons,
  // etc) -- relatedTarget still sits inside e.currentTarget in that case,
  // so it is not a real "left the card" event, and clearing on it would
  // wipe an otherwise-valid pending reorder before it gets a chance to
  // flush (Codex review, pass 3).
  if(e.currentTarget&&e.currentTarget.contains&&e.currentTarget.contains(e.relatedTarget))return;
  if(_priPendingReorder&&_priPendingReorder.type==='card'&&_priPendingReorder.target===e.currentTarget){
    _priPendingReorder=null;
  }
}
function priCardDrop(e,sec,priKey){
  e.preventDefault();e.stopPropagation();
  if(_emailDragData){const{item,cls}=_emailDragData;_addEmailCardToPriority(item,cls,sec);emailCardDragEnd(e);_priInsertCardIntoBoard(item,cls,sec);return;}
  if(!_priDragState)return;
  const{sec:fromSec,priKey:fromKey}=_priDragState;
  if(fromSec!==sec)_priSetOverride(fromKey,sec);
  _priDropTargetSec=sec;
  _priDragDropped=true;
}
function priZoneDragOver(e,sec){
  if(!_priDragState&&!_emailDragData)return;
  e.preventDefault();e.dataTransfer.dropEffect='move';
  const zone=document.querySelector(`.pri-drop-zone[data-sec="${sec}"]`);
  if(!zone)return;
  zone.classList.add('pri-zone-active');
  if(_priDragState&&_priDragEl){
    _priPendingReorder={type:'zone',zone};
    _priScheduleReorderFrame();
  }
}
function priZoneDragLeave(e,sec){
  const z=document.querySelector(`.pri-drop-zone[data-sec="${sec}"]`);
  if(z&&!z.contains(e.relatedTarget)){
    z.classList.remove('pri-zone-active');
    // Same staleness risk as priCardDragLeave above -- clear a pending
    // zone-append if the pointer has left the zone it targeted.
    if(_priPendingReorder&&_priPendingReorder.type==='zone'&&_priPendingReorder.zone===z){
      _priPendingReorder=null;
    }
  }
}
function priZoneDrop(e,sec){
  e.preventDefault();
  if(_emailDragData){const{item,cls}=_emailDragData;_addEmailCardToPriority(item,cls,sec);emailCardDragEnd(e);_priInsertCardIntoBoard(item,cls,sec);return;}
  if(!_priDragState)return;
  const{sec:fromSec,priKey:fromKey}=_priDragState;
  if(fromSec!==sec)_priSetOverride(fromKey,sec);
  _priDropTargetSec=sec;
  _priDragDropped=true;
}

// Staleness for "Priority actions - this week" (zone 'pw'), 21 Aug 2026 --
// Phase 2 item 3, work-inbox stability plan. This is the SAME definition
// used by command-centre's own lastActivityTs() fix, same day: genuine
// activity is a manually-logged/untagged action entry, or one tagged
// "(email: Kevin (sent to: ...)" -- Kevin's own sent reply. A routine
// auto-logged inbound email, tagged "(email: <sender> - <subject>)" by
// fetch_inbox.py Phase 3.5/3.6, does NOT reset the clock on its own. The
// pw zone's default contents are command-centre's own tier:'week' tasks,
// mirrored verbatim including their `actions[]` array (see fetch_inbox.py,
// the "Command Centre loaded" block) -- so this reuses the exact same
// action-log strings command-centre's fix reads, not a re-derived copy,
// which is what makes the two dashboards' definitions actually the same
// definition rather than two definitions that happen to agree today.
// Threshold (21 days) matches command-centre's own CC_STALE_DAYS.week.
// Scoped to 'pw' only, per Kevin's ask -- this is new aging visibility for
// "Priorities This Week" specifically, not a redesign of the other five
// board sections (todo/tomorrow/urgent/needs/fyi), which have no aging
// signal built for them here and are deliberately left untouched.
var WI_PW_STALE_DAYS=21;
var WI_MONTHS={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
function _priLastActivityTs(p){
  var best=0,earliest=Infinity,genuine=0;
  if(p.dateAdded){var dv=new Date(p.dateAdded+'T12:00:00').getTime();if(!isNaN(dv)&&dv>best)best=dv;}
  if(p.lastUpdated){var lv=new Date(p.lastUpdated).getTime();if(!isNaN(lv)&&lv>best)best=lv;}
  var acts=p.actions;
  if(acts){
    if(!Array.isArray(acts))acts=[acts];
    acts.forEach(function(a){
      var s=String(a);
      var m=/^\s*\[(\d{1,2})\s+([A-Za-z]{3})[A-Za-z]*\.?\s*(\d{4})?\]/.exec(s);
      if(!m)return;
      var mo=WI_MONTHS[m[2].toLowerCase()];
      if(mo===undefined)return;
      var yr=m[3]?parseInt(m[3],10):new Date().getFullYear();
      var v=new Date(yr,mo,parseInt(m[1],10)).getTime();
      if(isNaN(v))return;
      if(v<earliest)earliest=v;
      var hasEmailTag=/\(email:/i.test(s);
      var isKevinSent=/\(email:\s*Kevin\s*\(sent to:/i.test(s);
      if(hasEmailTag&&!isKevinSent)return;
      if(v>genuine)genuine=v;
    });
  }
  if(genuine>best)best=genuine;
  if(!best&&earliest!==Infinity)best=earliest;
  // Fallback for items dragged in from a raw email card (urgent/needs/fyi),
  // which carry no actions[] log at all -- only the underlying message's
  // own received_raw timestamp. Same "first-seen, never touched again"
  // semantics as the actions-log earliest-entry fallback above: a single
  // inbound receipt date is a genuine (if weak) aging anchor, since there
  // is by definition no second inbound touch here to falsely reset it.
  if(!best&&p.received_raw){var rv=new Date(p.received_raw).getTime();if(!isNaN(rv))best=rv;}
  return best;
}
function _priStaleDays(p,sec){
  if(sec!=='pw')return null; // scoped to "Priorities This Week" only, see comment above
  var ts=_priLastActivityTs(p);
  if(!ts)return null;
  var days=Math.floor((Date.now()-ts)/(24*3600*1000));
  return days>=WI_PW_STALE_DAYS?days:null;
}

// Single-card renderer -- extracted 20 Aug 2026 (drag-and-drop architecture
// rework) from what used to be inline in renderPriorityCards' .map(). Now
// the ONE place that knows how to render a priority card, used both by the
// full multi-card render below AND by priDragEnd's/`_priInsertCardIntoBoard`'s
// targeted single-card patch paths -- closing the "two different code paths
// producing the same visual card state" divergence risk the 12 Aug review
// flagged (point 5), not just the render/state coupling.
//
// DOM/tick id is now the card's own stable `priKey` (already 'eid_<id>' or
// 'id_<id>' from _priGetKey(), confirmed 100% coverage across all six
// source arrays) instead of a render-position index ('pri_'+sec+'_'+i).
// This is the second half of the tick-key stability fix (see
// _tickStorageKey() above): a positional id silently detaches from its
// card on ANY reorder, cross-zone move, or fresh pipeline run that
// reshuffles array order -- exactly the mechanism behind the 17 Aug tick-
// resurrection incident. Using the item's own identity means a card's
// done-state and DOM id follow the card itself, including across a
// cross-zone drag, not the slot it happens to occupy.
function _priRenderOneCard(p,sec){
  const _mo=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const _recentPfxs4=Array.from({length:4},(_,i)=>{const d=new Date();d.setDate(d.getDate()-i);return'['+String(d.getDate()).padStart(2,'0')+' '+_mo[d.getMonth()]+' '+d.getFullYear()+']';});
  const _cutoff=new Date();_cutoff.setDate(_cutoff.getDate()-4);_cutoff.setHours(0,0,0,0);
  const _mo2={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  function _firstActionDate(actions){if(!actions||!actions.length)return null;const m=actions[0].match(/^\[(\d{1,2}) (\w{3}) (\d{4})\]/);if(!m)return null;return new Date(parseInt(m[3]),_mo2[m[2]],parseInt(m[1]),12,0,0);}
  const priKey=p._priKey||_priGetKey(p);
  const id=priKey, ticked=isTicked(id);
  const titleText=(p.title||p.text||'(untitled)').replace(' -- ',' — ');
  const aiBadge=(p.badge&&sec!=='pfyi')?badge(p.badge,p.badgeType||'gray'):'';
  const createdDate=p.dateAdded?new Date(p.dateAdded+'T12:00:00'):_firstActionDate(p.actions);
  const newBadge=(!aiBadge&&createdDate&&createdDate>=_cutoff)?badge('NEW','green'):'';
  const updBadge=(!aiBadge&&!newBadge&&p.actions&&p.actions.some(a=>_recentPfxs4.some(pfx=>a.startsWith(pfx))))?badge('UPDATED','blue'):'';
  const theBadge=aiBadge||newBadge||updBadge;
  // "Priorities This Week" staleness badge, 21 Aug 2026 (Phase 2 item 3) --
  // additive, does not replace theBadge above. Only shown for the 'pw'
  // zone (see _priStaleDays' own scope guard) and never on an already-
  // ticked/done item, matching command-centre's own "!done" guard on its
  // stale badge.
  const staleDaysVal=(!ticked)?_priStaleDays(p,sec):null;
  const staleBadge=(staleDaysVal!==null)?`<span class="badge badge-red" title="In Priorities This Week but no genuine activity logged for ${staleDaysVal} days">${staleDaysVal}D QUIET</span>`:'';
  let subText='';
  if(p.actions&&p.actions.length){
    const todo=p.actions.find(a=>a.startsWith('[TODO]')||a.startsWith('[AWAITING]'));
    const latest=todo||p.actions[p.actions.length-1];
    if(latest) subText=latest.replace(/^\[[^\]]+\]\s*/,'');
  }
  const subLine=(p.source&&subText)?p.source+' · '+subText:(p.source||subText||p.ai_summary||p.sub||'');
  const emailBtn=(p.entry_id||p.entryId)?`<span class="card-icon" title="Open email" onclick="openEmail('${p.entry_id||p.entryId}',event)">&#9993;</span>`:'';
  const ccBtn=p.id?`<span class="card-icon-cc" title="Command Centre" onclick="window.open('https://cc.lelitte.co.uk/#${p.id}','_blank');event.stopPropagation()">CC&#8594;</span>`:'';
  const hiddenCls=(ticked&&!showingDoneItems)?' card-hidden':'';
  return `<div class="card-ph${ticked?' done':''}${hiddenCls}" id="item_${id}" data-prikey="${priKey}" data-sec="${sec}" draggable="true" ondragstart="priDragStart(event,'${sec}','${priKey}')" ondragend="priDragEnd(event)" ondragover="priCardDragOver(event,'${sec}','${priKey}')" ondragleave="priCardDragLeave(event,'${priKey}')" ondrop="priCardDrop(event,'${sec}','${priKey}')">
      <span class="card-drag" onclick="event.stopPropagation()">&#10783;</span>
      <button class="card-done-btn${ticked?' done':''}" id="cb_${id}" onclick="toggleTick('${id}');event.stopPropagation()" aria-label="Mark done"></button>
      <div class="card-ph-body">
        <div class="card-ph-title${ticked?' done':''}">${titleText}</div>
        ${subLine?`<div class="card-ph-sub">${sanitizeSub(subLine)}</div>`:''}
      </div>
      <div class="card-ph-actions">${theBadge}${staleBadge}${emailBtn}${ccBtn}</div>
    </div>`;
}
function _priZonePlaceholderHtml(sec){return sec==='pfyi'?'Drop items here to park':'Drop items here';}
function renderPriorityCards(priorities,key,sec){
  if(!priorities||!priorities.length) return `<div class="pri-zone-empty">${_priZonePlaceholderHtml(sec)}</div>`;
  return priorities.map(p=>_priRenderOneCard(p,sec)).join('');
}

// Targeted zone chrome patch (section header count + empty-zone placeholder)
// -- added 20 Aug 2026 alongside the drag-and-drop architecture rework, used
// after a single-card DOM patch instead of regenerating the header/zone
// markup for all six sections via a full renderBriefing().
function _priZoneCardCount(sec){
  const zone=document.querySelector(`.pri-drop-zone[data-sec="${sec}"]`);
  return zone?zone.querySelectorAll('.card-ph').length:0;
}
function _priUpdateZoneChrome(sec){
  const zone=document.querySelector(`.pri-drop-zone[data-sec="${sec}"]`);
  if(!zone)return;
  const count=zone.querySelectorAll('.card-ph').length;
  const wrap=zone.closest('[id^="sec-"]');
  const countEl=wrap?wrap.querySelector('.sec-count'):null;
  // Only overwrite the plain-number case -- if the header is showing the
  // "N threads (M messages)" server-side-dedup form (fyiRawCount), a
  // client-side count patch can't know the new raw count, so leave that
  // form alone rather than clobber it with a number that would be wrong.
  // A drag can only ever change the displayed thread count anyway (moving
  // a card doesn't change how many raw messages it collapsed from), so this
  // is a display-only edge case, not a correctness one.
  if(countEl&&!countEl.querySelector('.sec-count-raw')) countEl.textContent=String(count);
  let placeholder=zone.querySelector('.pri-zone-empty:not(.wi-search-no-match)');
  if(count===0&&!placeholder){
    placeholder=document.createElement('div');
    placeholder.className='pri-zone-empty';
    placeholder.textContent=_priZonePlaceholderHtml(sec);
    zone.appendChild(placeholder);
  }else if(count>0&&placeholder){
    placeholder.remove();
  }
}
// Inbox-card-dragged-onto-board insert patch -- replaces the previous
// unconditional renderBriefing() call in priCardDrop/priZoneDrop's
// _emailDragData branch. _addEmailCardToPriority() has already updated the
// underlying data (workInbox_customPri_v1 + the section override) by the
// time this runs; this only needs to add the ONE new card's markup to the
// target zone's DOM, same principle as the rest of this rework.
function _priInsertCardIntoBoard(item,cls,sec){
  const zone=document.querySelector(`.pri-drop-zone[data-sec="${sec}"]`);
  if(!zone)return;
  const priorities=applyPriOverrides(window._wipData||{})[sec]||[];
  const priKey=_priGetKey(item);
  const p=priorities.find(x=>x._priKey===priKey);
  if(!p)return; // defensive -- should always be found immediately after _addEmailCardToPriority
  const html=_priRenderOneCard(p,sec);
  const tmp=document.createElement('div');
  tmp.innerHTML=html.trim();
  // Newest-first-insertion requirement, 20 Aug 2026 -- an email card
  // dragged from the Inbox column onto the board is a NEW item entering
  // this zone, so it goes to the top (insertBefore the zone's current
  // first card), not the bottom (appendChild, the old behaviour).
  zone.insertBefore(tmp.firstElementChild,zone.firstElementChild);
  _priUpdateZoneChrome(sec);
  _runCardSearch();
}

// Staleness banner: fetch_inbox.py is scheduled Mon-Fri at 06:00, 09:00,
// 12:00, 15:00, 18:00. Rather than a blunt "older than N hours" check (which
// would falsely fire every weekend, since nothing runs Sat/Sun by design),
// compute the most recent run time that should already have happened as of
// right now, and compare that against data.refreshed_at. A 90-minute grace
// period covers the run's own execution time before flagging it as missed.
const SCHEDULE_RUN_HOURS=[6,9,12,15,18];
const SCHEDULE_GRACE_MINUTES=90;

function _mostRecentExpectedRun(now){
  for(let dayOffset=0; dayOffset<9; dayOffset++){
    const day=new Date(now.getFullYear(),now.getMonth(),now.getDate()-dayOffset);
    const dow=day.getDay(); // 0=Sun .. 6=Sat
    if(dow===0||dow===6) continue;
    for(let i=SCHEDULE_RUN_HOURS.length-1;i>=0;i--){
      const runTime=new Date(day.getFullYear(),day.getMonth(),day.getDate(),SCHEDULE_RUN_HOURS[i],0,0);
      const withGrace=new Date(runTime.getTime()+SCHEDULE_GRACE_MINUTES*60000);
      if(withGrace<=now) return runTime;
    }
  }
  return null;
}

function _parseRefreshedAt(str,refYear){
  if(!str) return null;
  const m=str.match(/(\d{1,2})\s+(\w+)\s*[·\-]\s*(\d{1,2}):(\d{2})/);
  if(!m) return null;
  const MONTHS=['january','february','march','april','may','june','july','august','september','october','november','december'];
  const monIdx=MONTHS.indexOf(m[2].toLowerCase());
  if(monIdx<0) return null;
  return new Date(refYear,monIdx,parseInt(m[1]),parseInt(m[3]),parseInt(m[4]),0);
}

function renderStaleBanner(data){
  let el=document.getElementById('staleBanner');
  if(!el){
    el=document.createElement('div');
    el.id='staleBanner';
    const headerDate=document.getElementById('headerDate');
    if(headerDate&&headerDate.parentNode) headerDate.parentNode.insertBefore(el,headerDate.nextSibling);
    else document.body.insertBefore(el,document.body.firstChild);
  }
  const now=new Date();
  const refreshed=_parseRefreshedAt(data.refreshed_at,now.getFullYear());
  const expected=_mostRecentExpectedRun(now);
  if(!refreshed||!expected||refreshed>=expected){
    el.style.display='none';
    el.innerHTML='';
    return;
  }
  const hoursBehind=Math.round((now-refreshed)/3600000);
  el.style.cssText='display:block;background:#a3271f;color:#fff;padding:10px 18px;border-radius:8px;margin:10px 0 4px;font-size:13px;font-weight:600;';
  el.innerHTML='&#9888; Data may be out of date &mdash; last refreshed '+escapeHtml(data.refreshed_at||'unknown')+
    ' ('+hoursBehind+'h ago). A refresh was expected by '+escapeHtml(expected.toLocaleString('en-GB',{weekday:'short',hour:'2-digit',minute:'2-digit'}))+
    '. Run "Run Inbox Briefing.bat" manually if this persists.';
}

function renderBriefing(data,key){
  currentData=data; currentKey=key;
  window._wipData=data; window._wipKey=key;
  document.getElementById('pageTitle').textContent=getGreeting();
  document.getElementById('headerDate').textContent=data.date+(data.subtitle?' · '+data.subtitle:'');
  renderStaleBanner(data);
  const stamp=document.getElementById('refresh-stamp');
  if(stamp&&data.refreshed_at) stamp.textContent='Last refreshed: '+data.refreshed_at;
  renderCalPanel(data);
  setupCtxTicker(data.context);
  const absEl=document.getElementById('absencesSidebar');
  if(absEl){
    if(data.absences&&data.absences.length){
      const fmtAbsence=a=>{
        let text=String(a).trim();
        if(text&&!/ - |returns|today|tomorrow|next week|date unknown/i.test(text)){
          text+=' - date unknown';
        }
        return text.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
      };
      absEl.innerHTML='<ul class="abs-list">'+data.absences.map(a=>`<li>${fmtAbsence(a)}</li>`).join('')+'</ul>';
    } else {
      absEl.innerHTML='<span style="font-size:11px;color:rgba(255,255,255,0.3);font-style:italic">None recorded</span>';
    }
  }
  const priSecs=applyPriOverrides(data);
  document.getElementById('inboxCol').innerHTML=`<div class="inbox-grid" id="inboxGrid">
    <div id="col-left">
      <div id="sec-urgent-wrap">
        ${_secHeadHtml('ur','dot-r','Urgent – action required today',priSecs.ur.length)}
        <div class="pri-drop-zone" data-sec="ur" ondragover="priZoneDragOver(event,'ur')" ondragleave="priZoneDragLeave(event,'ur')" ondrop="priZoneDrop(event,'ur')">${priSecs.ur.length?renderPriorityCards(priSecs.ur,key,'ur'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
      </div>
      <div id="sec-tomorrow-wrap" style="margin-top:18px">
        ${_secHeadHtml('ptom','dot-o','Priority actions – tomorrow',priSecs.ptom.length)}
        <div class="pri-drop-zone" data-sec="ptom" ondragover="priZoneDragOver(event,'ptom')" ondragleave="priZoneDragLeave(event,'ptom')" ondrop="priZoneDrop(event,'ptom')">${priSecs.ptom.length?renderPriorityCards(priSecs.ptom,key,'ptom'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
      </div>
      <div id="sec-week-wrap" style="margin-top:18px">
        ${_secHeadHtml('pw','dot-green','Priority actions – this week',priSecs.pw.length)}
        <div class="pri-drop-zone" data-sec="pw" ondragover="priZoneDragOver(event,'pw')" ondragleave="priZoneDragLeave(event,'pw')" ondrop="priZoneDrop(event,'pw')">${priSecs.pw.length?renderPriorityCards(priSecs.pw,key,'pw'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
      </div>
    </div>
    <div id="col-right">
      <div id="sec-today-wrap">
        ${_secHeadHtml('pt','dot-r','Priority actions – today',priSecs.pt.length)}
        <div class="pri-drop-zone" data-sec="pt" ondragover="priZoneDragOver(event,'pt')" ondragleave="priZoneDragLeave(event,'pt')" ondrop="priZoneDrop(event,'pt')">${priSecs.pt.length?renderPriorityCards(priSecs.pt,key,'pt'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
      </div>
      <div id="sec-needs-wrap" style="margin-top:18px">
        ${_secHeadHtml('nr','dot-o','Needs response – within 24–48 hrs',priSecs.nr.length)}
        <div class="pri-drop-zone" data-sec="nr" ondragover="priZoneDragOver(event,'nr')" ondragleave="priZoneDragLeave(event,'nr')" ondrop="priZoneDrop(event,'nr')">${priSecs.nr.length?renderPriorityCards(priSecs.nr,key,'nr'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
      </div>
      <div id="sec-parked-wrap" style="margin-top:18px">
        ${_secHeadHtml('pfyi','dot-g','FYI / Parked',priSecs.pfyi.length,typeof data.fyiRawCount==='number'?data.fyiRawCount:undefined)}
        <div class="pri-drop-zone" data-sec="pfyi" ondragover="priZoneDragOver(event,'pfyi')" ondragleave="priZoneDragLeave(event,'pfyi')" ondrop="priZoneDrop(event,'pfyi')">${priSecs.pfyi.length?renderPriorityCards(priSecs.pfyi,key,'pfyi'):'<div class="pri-zone-empty">Drop items here to park</div>'}</div>
      </div>
    </div>
  </div>`;
  ['pt','ptom','ur','pw','nr','pfyi'].forEach(sec=>applySecCollapse(sec,!!getCollapsedSecs()[sec]));
  _runCardSearch();
}

// Card search (Priorities board) — plain client-side substring filter across
// each card's full rendered text (subject/sender/summary), live match count,
// per-section "No matches" state. Called at the end of renderBriefing() so an
// active search term survives a full re-render (drag-drop, tick, priority
// overrides all rebuild #inboxGrid's innerHTML from scratch).
let _wiSearchTerm='';
function applyCardSearch(val){
  _wiSearchTerm=(val||'').trim().toLowerCase();
  const clearBtn=document.getElementById('wiSearchClear');
  if(clearBtn) clearBtn.style.display=_wiSearchTerm?'':'none';
  _runCardSearch();
}
function clearCardSearch(){
  const input=document.getElementById('wiSearchInput');
  if(input) input.value='';
  applyCardSearch('');
}
function _runCardSearch(){
  const zones=document.querySelectorAll('.pri-drop-zone');
  let totalCards=0, totalMatched=0;
  zones.forEach(zone=>{
    const cards=zone.querySelectorAll('.card-ph');
    if(!cards.length) return; // leave the existing "Drop items here" empty-zone state untouched
    let matched=0;
    cards.forEach(card=>{
      totalCards++;
      const isMatch=!_wiSearchTerm||card.textContent.toLowerCase().includes(_wiSearchTerm);
      card.style.display=isMatch?'':'none';
      if(isMatch){matched++;totalMatched++;}
    });
    let noMatchEl=zone.querySelector('.wi-search-no-match');
    if(_wiSearchTerm&&matched===0){
      if(!noMatchEl){
        noMatchEl=document.createElement('div');
        noMatchEl.className='pri-zone-empty wi-search-no-match';
        noMatchEl.textContent='No matches';
        zone.appendChild(noMatchEl);
      }
    } else if(noMatchEl){
      noMatchEl.remove();
    }
  });
  const countEl=document.getElementById('wiSearchCount');
  if(countEl) countEl.textContent=_wiSearchTerm?(totalMatched+' of '+totalCards+' match'+(totalCards===1?'':'es')):'';
}

// Collapse/expand per section (Urgent, Needs response, etc.) - state persists
// across reloads via localStorage, since these lists (Needs response
// especially) can grow long enough that always-expanded becomes unwieldy.
function getCollapsedSecs(){try{return JSON.parse(localStorage.getItem('workInbox_collapsedSecs_v1')||'{}');}catch(e){return{};}}
function setCollapsedSecs(o){localStorage.setItem('workInbox_collapsedSecs_v1',JSON.stringify(o));}
function toggleSecCollapse(sec){
  const o=getCollapsedSecs();
  o[sec]=!o[sec];
  setCollapsedSecs(o);
  applySecCollapse(sec,o[sec]);
}
function applySecCollapse(sec,collapsed){
  const zone=document.querySelector('.pri-drop-zone[data-sec="'+sec+'"]');
  const chev=document.getElementById('chev_'+sec);
  if(zone) zone.style.display=collapsed?'none':'';
  if(chev) chev.innerHTML=collapsed?'&#9656;':'&#9662;';
}
function _secHeadHtml(sec,dotClass,label,count,rawCount){
  // rawCount (optional): the true pre-dedup message count, when it differs
  // from the displayed `count`. Added 12 Aug 2026 so any dedup applied to a
  // board section is visible and labelled, not a silent reduction Kevin has
  // no way to see -- e.g. "18 threads (21 messages)" instead of just "18".
  // Server-side thread-collapse (fetch_inbox.py Phase 3.3c) already
  // resolves genuine duplicate "RE:"/"FW:" threads before this ever runs;
  // this label only fires when the two numbers genuinely differ, so a
  // section with no collapsing still shows a plain count as before.
  const countHtml=(typeof rawCount==='number'&&rawCount!==count)
    ? count+' threads <span class="sec-count-raw" style="font-weight:400;color:var(--text-muted)">('+rawCount+' messages)</span>'
    : String(count);
  return '<div class="sec-head" onclick="toggleSecCollapse(\''+sec+'\')" style="cursor:pointer;user-select:none">'
    +'<span class="sec-chev" id="chev_'+sec+'" style="display:inline-block;width:14px;color:var(--text-muted);font-size:11px">&#9662;</span>'
    +'<span class="sec-dot '+dotClass+'"></span><span class="sec-lbl">'+label+'</span><span class="sec-rule"></span><span class="sec-count">'+countHtml+'</span></div>';
}

function toggleSum(id,btn){const el=document.getElementById(id);const exp=el.classList.toggle('expanded');btn.textContent=exp?'Show less':'Show more';}

function renderCalPanel(data){
  const el=document.getElementById('calPanel');
  if(!el) return;
  const now=new Date();
  const nowMins=now.getHours()*60+now.getMinutes();
  const todayDate=now.getDate(), todayMonth=now.getMonth(), todayYear=now.getFullYear();
  function parseTimeMins(t){if(!t)return -1;const p=t.split(':');return p.length<2?-1:parseInt(p[0])*60+parseInt(p[1]);}
  // Same weekend-skipping semantics as the backend's next_workday() -- kept
  // in sync deliberately, not shared code, since this is a small pure date
  // helper duplicated across the Python/JS boundary.
  function nextWorkday(d){
    const n=new Date(d); n.setDate(n.getDate()+1);
    while(n.getDay()===0||n.getDay()===6) n.setDate(n.getDate()+1);
    return n;
  }
  function renderBlock(items,headerHtml,isToday,bodyId){
    if(!items||!items.length) return `<div class="main-cal-block"><div class="main-cal-block-header">${headerHtml}</div><div class="main-cal-none">No meetings</div></div>`;
    let nextFound=false;
    const rows=items.map((c,i)=>{
      const mins=parseTimeMins(c.time);
      const isPast=isToday&&mins>=0&&mins<nowMins;
      const isNext=isToday&&!isPast&&!nextFound&&mins>=nowMins;
      if(isNext) nextFound=true;
      const cls=isPast?' past':isNext?' next':'';
      const sumId=c.id?'sum_'+c.id:bodyId+i;
      // Only link to Command Centre when this meeting has a genuine matching
      // task id (c.ccTaskId, attached server-side in fetch_inbox.py via an
      // exact emailRef match -- see _match_cc_task_id() there). Command
      // Centre's own js/app.js reads window.location.hash on load and
      // scrolls to + highlights '#card-'+hash, so this deep-links straight
      // to the real item instead of just landing on the CC homepage --
      // Kevin's explicit ask, 10 Aug 2026: "it should high[light] the item
      // so i can drill dowwn into the email if required." No matching task
      // -> no CC link at all, rather than one that goes nowhere useful.
      const ccLink=c.ccTaskId?`<a class="summary-cc-link" href="https://cc.lelitte.co.uk/#${encodeURIComponent(c.ccTaskId)}" target="_blank">CC &#8594;</a>`:'';
      return `<div class="main-cal-item${cls}"><span class="main-cal-time">${escapeHtml(c.time||'')}</span><div style="flex:1;min-width:0"><div class="main-cal-title">${escapeHtml(c.title)}</div>${c.sub?`<div class="main-cal-sub">${escapeHtml(c.sub)}</div>`:''}${c.summary?`<div class="main-cal-summary-wrap"><div class="main-cal-summary-text" id="${sumId}">${escapeHtml(c.summary)}</div><div class="main-cal-summary-footer"><button class="summary-toggle" onclick="toggleSum('${sumId}',this)">Show more</button>${ccLink}</div></div>`:''}</div></div>`;
    }).join('');
    return `<div class="main-cal-block"><div class="main-cal-block-header">${headerHtml}</div><div class="cal-col-body" id="${bodyId}">${rows}</div></div>`;
  }
  // mtgDates: real Date objects for each of the 4 rolling day-view columns
  // that actually has at least one item -- used to mark "has-meeting" dots
  // across however many of the 4 displayed months those dates fall in.
  function renderMiniCal(monthOffset,mtgDates){
    const calDate=new Date(todayYear,todayMonth+(monthOffset||0),1);
    const calYear=calDate.getFullYear(), calMonth=calDate.getMonth();
    const monthName=calDate.toLocaleDateString('en-GB',{month:'long',year:'numeric'});
    const daysInMonth=new Date(calYear,calMonth+1,0).getDate();
    let startDow=calDate.getDay()-1; if(startDow<0) startDow=6;
    const dayNames=['M','T','W','T','F','S','S'];
    let cells=dayNames.map(d=>`<div class="mini-cal-day-name">${d}</div>`).join('');
    for(let i=0;i<startDow;i++) cells+='<div class="mini-cal-day other-month"></div>';
    for(let d=1;d<=daysInMonth;d++){
      const isT=(d===todayDate&&calMonth===todayMonth&&calYear===todayYear);
      const hasMtg=mtgDates.some(md=>md.getDate()===d&&md.getMonth()===calMonth&&md.getFullYear()===calYear);
      const cls='mini-cal-day'+(isT?' today':hasMtg?' has-meeting':'');
      cells+=`<div class="${cls}">${d}</div>`;
    }
    return `<div class="main-cal-block"><div class="main-cal-block-header">${monthName}</div><div class="mini-cal-grid">${cells}</div></div>`;
  }
  const todayHeader='Today &mdash; '+now.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long'});
  const naiveTom=new Date(now); naiveTom.setDate(naiveTom.getDate()+1);
  const skippedWeekend=naiveTom.getDay()===0||naiveTom.getDay()===6;
  const tom=nextWorkday(now);
  const tomHeader=(skippedWeekend?'Next Week':'Tomorrow')+' &mdash; '+tom.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long'});
  // Rolling 4-day window (today + next 3 working days) -- Kevin's explicit
  // request, 10 Aug 2026: "today, tomorrow, day after that, and day after
  // that... when tomorrow comes, it will drop and get Friday." Day2/Day3
  // just carry their own weekday name as the header, same as how Kevin
  // described them, rather than a "Today"/"Tomorrow"-style label.
  const day2=nextWorkday(tom);
  const day3=nextWorkday(day2);
  const day2Header=day2.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long'});
  const day3Header=day3.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long'});

  const mtgDates=[];
  if(data.calToday&&data.calToday.length) mtgDates.push(new Date(todayYear,todayMonth,todayDate));
  if(data.calTomorrow&&data.calTomorrow.length) mtgDates.push(tom);
  if(data.calDay2&&data.calDay2.length) mtgDates.push(day2);
  if(data.calDay3&&data.calDay3.length) mtgDates.push(day3);

  const daysRow=renderBlock(data.calToday,todayHeader,true,'calBodyToday')
    +renderBlock(data.calTomorrow,tomHeader,false,'calBodyTom')
    +renderBlock(data.calDay2,day2Header,false,'calBodyDay2')
    +renderBlock(data.calDay3,day3Header,false,'calBodyDay3');
  const monthsRow=renderMiniCal(0,mtgDates)+renderMiniCal(1,mtgDates)+renderMiniCal(2,mtgDates)+renderMiniCal(3,mtgDates);
  el.innerHTML=`<div class="main-cal-panel"><div class="main-cal-days-row">${daysRow}</div><div class="main-cal-months-row">${monthsRow}</div></div>`;
}

let _ctxSentences=[], _ctxIdx=0, _ctxTimer=null, _ctxPaused=false;
function setupCtxTicker(context){
  const el=document.getElementById('contextBar');
  if(!el||!context){if(el)el.innerHTML='';return;}
  _ctxSentences=context.split(/(?<=[.!?])\s+/).filter(s=>s.trim().length>4);
  if(!_ctxSentences.length){el.innerHTML='';return;}
  _ctxIdx=0;
  el.innerHTML=`<div class="ctx-strip" onmouseenter="_ctxPaused=true" onmouseleave="_ctxPaused=false" onclick="_jumpCtx(_ctxIdx+1)">
    <div class="ctx-label">Briefing context</div>
    <div class="ctx-text" id="ctxText"></div>
    <div class="ctx-dots" id="ctxDots"></div>
  </div>`;
  _renderCtx();
  if(_ctxTimer) clearInterval(_ctxTimer);
  _ctxTimer=setInterval(()=>{if(!_ctxPaused){_ctxIdx=(_ctxIdx+1)%_ctxSentences.length;_renderCtx();}},4500);
}
function _renderCtx(){
  const txt=document.getElementById('ctxText');
  const dots=document.getElementById('ctxDots');
  if(txt){txt.style.animation='none';txt.offsetHeight;txt.style.animation='ctxFlipIn .35s ease';txt.textContent=_ctxSentences[_ctxIdx];}
  if(dots) dots.innerHTML=_ctxSentences.map((_,i)=>`<div class="ctx-dot${i===_ctxIdx?' active':''}" onclick="event.stopPropagation();_jumpCtx(${i})"></div>`).join('');
}
function _jumpCtx(i){
  _ctxIdx=((i%_ctxSentences.length)+_ctxSentences.length)%_ctxSentences.length;
  _renderCtx();
}

function applyFilter(val){
  const map={today:'sec-today-wrap',tomorrow:'sec-tomorrow-wrap',week:'sec-week-wrap',parked:'sec-parked-wrap'};
  const grid=document.getElementById('inboxGrid');
  if(!grid) return;
  const sel=document.getElementById('tierSelect');
  if(sel) sel.value=val;
  const cl=document.getElementById('col-left'), cr=document.getElementById('col-right');
  if(val==='all'){
    Object.values(map).forEach(id=>{const el=document.getElementById(id);if(el)el.style.display='';});
    if(cl)cl.style.display=''; if(cr)cr.style.display='';
    grid.style.gridTemplateColumns='';
  } else {
    Object.entries(map).forEach(([tier,id])=>{
      const el=document.getElementById(id);
      if(el) el.style.display=(tier===val)?'':'none';
    });
    if(val==='today'||val==='tomorrow'){
      if(cl)cl.style.display=''; if(cr)cr.style.display='none';
    } else {
      if(cl)cl.style.display='none'; if(cr)cr.style.display='';
    }
    grid.style.gridTemplateColumns='1fr';
  }
}
function clearSel(){
  document.querySelectorAll('.ticker-stat.selected').forEach(el=>el.classList.remove('selected'));
}
function clickStat(tier){
  const stat=document.querySelector(`.ticker-stat[data-tier="${tier}"]`);
  const wasSelected=stat&&stat.classList.contains('selected');
  clearSel();
  if(!wasSelected){
    if(stat) stat.classList.add('selected');
    applyFilter(tier);
  } else {
    applyFilter('all');
  }
}

function getGreeting(){
  const hour=parseInt(new Date().toLocaleString('en-GB',{timeZone:'Europe/London',hour:'numeric',hour12:false}));
  if(hour>=5&&hour<12) return 'Good morning, Kevin';
  if(hour>=12&&hour<18) return 'Good afternoon, Kevin';
  return 'Good evening, Kevin';
}

const BRIEFING_API='https://github-proxy.lelitte.co.uk/work-inbox/data/briefing.json';

async function init(){
  const titleEl=document.getElementById('pageTitle');
  if(titleEl) titleEl.textContent=getGreeting();

  await loadRemoteTicks();
  loadDraftedReplies();

  let data=null;

  try{
    const res=await fetch(BRIEFING_API+'?t='+Date.now(),{cache:'no-store'});
    if(res.ok){
      data=await res.json();
      const key=data.date?data.date.replace(/[^a-zA-Z0-9]/g,'_'):'latest';
      const store=getStore(); store[key]=data; saveStore(store);
      localStorage.setItem(TODAY_KEY,key);
    }
  }catch(e){console.warn('GitHub fetch failed:',e);}

  if(!data){
    const todayKey=localStorage.getItem(TODAY_KEY);
    if(todayKey){const store=getStore(); if(store[todayKey]) data=store[todayKey];}
  }

  if(!data){
    document.getElementById('headerDate').textContent='No briefing available. Run fetch_inbox.py to generate one.';
    return;
  }

  const currentKey=data.date?data.date.replace(/[^a-zA-Z0-9]/g,'_'):'latest';
  const ticks=getTicks();
  const allItems=[...(data.urgent||[]),(data.needs||[]),(data.fyi||[]),(data.low||[])].flat();
  if(allItems.length>0){
    const sections=['urgent','needs','fyi','low'];
    let hiddenCount=0;
    sections.forEach(function(s){
      (data[s]||[]).forEach(function(_,i){
        if(ticks[currentKey+'_'+s+'_'+i]) hiddenCount++;
      });
    });
    if(hiddenCount===allItems.length){
      saveTicks({});
    }
  }

  renderBriefing(data, currentKey);
}

// Tabs -- added 10 Aug 2026 once Drafted Replies joined Calendar + Priorities
// on one long scroll and it got too crowded. Persisted per-browser via
// localStorage so a refresh doesn't silently reset which tab was open.
const ACTIVE_TAB_KEY='workInbox_activeTab_v1';
const VALID_TABS=['priorities','calendar','drafted'];

function switchTab(tab){
  if(VALID_TABS.indexOf(tab)===-1) tab='priorities';
  document.querySelectorAll('.tab-btn').forEach(function(btn){
    btn.classList.toggle('active', btn.getAttribute('data-tab')===tab);
  });
  document.querySelectorAll('.tab-content').forEach(function(el){el.classList.remove('active');});
  const contentEl=document.getElementById('tabContent'+tab.charAt(0).toUpperCase()+tab.slice(1));
  if(contentEl) contentEl.classList.add('active');
  localStorage.setItem(ACTIVE_TAB_KEY, tab);
}

function initTabs(){
  const saved=localStorage.getItem(ACTIVE_TAB_KEY);
  switchTab(VALID_TABS.indexOf(saved)!==-1 ? saved : 'priorities');
}
initTabs();

init();

// CC ticker — reads Command Centre tasks.json
async function loadCcTicker(){
  try{
    const res=await fetch('https://github-proxy.lelitte.co.uk/command-centre/data/tasks.json?t='+Date.now(),{cache:'no-store'});
    if(!res.ok) throw new Error('fetch failed');
    const d=await res.json();
    const tasks=Array.isArray(d)?d:(d.tasks||[]);
    const openTasks=tasks.filter(t=>!t.done);
    const now=new Date(); now.setHours(0,0,0,0);
    function ageDays(t){
      if(!t.dateAdded) return 0;
      const dd=new Date(t.dateAdded); dd.setHours(0,0,0,0);
      return Math.max(0,Math.round((now-dd)/86400000));
    }
    const todayTasks=openTasks.filter(t=>t.tier==='today');
    const ages=openTasks.map(t=>ageDays(t));
    function setEl(id,v){const el=document.getElementById(id);if(el)el.textContent=v;}
    setEl('cc-today-count',todayTasks.length);
    setEl('cc-tmrw-count',openTasks.filter(t=>t.tier==='tomorrow').length);
    setEl('cc-week-count',openTasks.filter(t=>t.tier==='week').length);
    setEl('cc-parked-count',openTasks.filter(t=>t.tier==='parked').length);
    const stalled=todayTasks.filter(t=>ageDays(t)>=5).length;
    const oldest=ages.length?Math.max(...ages):0;
    const avg=ages.length?Math.round(ages.reduce((a,b)=>a+b,0)/ages.length):0;
    const twoWeeks=openTasks.filter(t=>ageDays(t)>=14).length;
    setEl('cc-stalled',stalled||'—');
    setEl('cc-oldest',oldest?oldest+'d':'—');
    setEl('cc-avg',avg?avg+'d':'—');
    setEl('cc-twoweeks',twoWeeks||'—');
  }catch(e){
    console.warn('CC ticker fetch failed',e);
  }
}
loadCcTicker();
setInterval(loadCcTicker, 60000);

// Drafted Replies — agent-commons issue #3, item 4 (4B, dashboard-only review).
// Reads work-inbox's own data/drafted_replies.json, a mirror publish_drafted_replies.py
// keeps in sync from agent-commons/pending-email-drafts/drafts.json (agent-commons is
// private -- the browser never talks to it directly; this file is the public,
// already-redacted, already-safe copy). Nothing here writes to a mailbox or sends
// anything -- "Mark sent"/"Discard" are bookkeeping only, riding the exact same
// tick-sync mechanism (getTicks/saveTicks/pushTicks) already used for email cards,
// just under a 'draft_' key prefix so it doesn't collide with per-day briefing ticks.
const DRAFTED_REPLIES_API='https://github-proxy.lelitte.co.uk/work-inbox/data/drafted_replies.json';

function draftTickKey(sourceEntryId){return 'draft_'+sourceEntryId;}
function draftStatus(sourceEntryId){return getTicks()[draftTickKey(sourceEntryId)]||null;}
function setDraftStatus(sourceEntryId,status){
  const ticks=getTicks();
  ticks[draftTickKey(sourceEntryId)]=status;
  saveTicks(ticks);
}

function drTierBadge(tier){
  const label=tier==='senior-management'?'Senior mgmt':tier==='direct-report'?'Direct report':'Other';
  return `<span class="dr-tier-badge dr-tier-${escapeHtml(tier||'other')}">${label}</span>`;
}

function toggleDraftExpand(id,ev){
  if(ev){ev.preventDefault();ev.stopPropagation();}
  const el=document.getElementById('drtext_'+id);
  if(el) el.classList.toggle('expanded');
  const hint=document.getElementById('drhint_'+id);
  if(hint) hint.textContent=el&&el.classList.contains('expanded')?'Show less':'Show more';
}

async function copyDraftText(id,ev){
  if(ev){ev.preventDefault();ev.stopPropagation();}
  const el=document.getElementById('drtext_'+id);
  if(!el) return;
  try{
    await navigator.clipboard.writeText(el.dataset.raw||el.textContent);
    wiNotify('Draft copied to clipboard.');
  }catch(e){
    wiNotify('Copy failed — select and copy manually.');
  }
}

function markDraft(id,sourceEntryId,status,ev){
  if(ev){ev.preventDefault();ev.stopPropagation();}
  setDraftStatus(sourceEntryId,status);
  const card=document.getElementById('drcard_'+id);
  if(card) card.remove();
  refreshDraftedRepliesCount();
}

function _updateDraftedTabBadge(count){
  const tabBadge=document.getElementById('tabDraftedCount');
  if(!tabBadge) return;
  if(count>0){tabBadge.textContent=count;tabBadge.style.display='';}
  else{tabBadge.style.display='none';}
}

function refreshDraftedRepliesCount(){
  const remaining=document.querySelectorAll('.dr-card').length;
  const badgeEl=document.getElementById('drCountBadge');
  if(badgeEl) badgeEl.textContent=remaining;
  _updateDraftedTabBadge(remaining);
  const panel=document.getElementById('draftedRepliesPanel');
  if(panel&&remaining===0){
    const list=document.getElementById('drList');
    if(list) list.innerHTML='<div class="dr-empty">No drafts waiting for review.</div>';
  }
}

function renderDraftedReplies(payload){
  const panel=document.getElementById('draftedRepliesPanel');
  if(!panel) return;
  const entries=(payload&&payload.entries)||[];
  const pending=entries.filter(e=>!draftStatus(e.source_entry_id));

  if(pending.length===0){
    panel.innerHTML=`<div class="dr-header"><div class="dr-title">Drafted Replies</div></div><div class="dr-empty">No drafts waiting for review.</div>`;
    _updateDraftedTabBadge(0);
    return;
  }

  const cards=pending.map((e,i)=>{
    const id='dr'+i;
    const draftEsc=escapeHtml(e.draft_text||'');
    const hasSource=e.source_entry_id&&e.source_entry_id.length>0;
    const openLink=hasSource?`<a href="javascript:void(0)" class="dr-btn" onclick="openEmail('${e.source_entry_id}',event)">Open original</a>`:'';

    // Confidence -- Lauren's own design explicitly calls this "impossible to
    // miss, not a hover tooltip," so it renders as a visible badge + reason
    // line in every card that has one, never hidden behind an expand/hover.
    const conf=e.confidence;
    const confHtml=(conf&&conf.level)?`<div class="dr-confidence dr-confidence-${escapeHtml(conf.level)}">
        <span class="dr-confidence-label">${escapeHtml(conf.level)} confidence</span>
        ${conf.reason?`<span class="dr-confidence-reason">${escapeHtml(conf.reason)}</span>`:''}
      </div>`:'';

    // Inline flags -- facts/figures Lauren couldn't verify. Per her own
    // stated design this is "a hard requirement, not a nice-to-have" --
    // rendered as its own visible callout, not folded into the draft text.
    const flags=Array.isArray(e.inline_flags)?e.inline_flags:[];
    const flagsHtml=flags.length?`<div class="dr-flags"><div class="dr-flags-title">Needs your confirmation</div><ul class="dr-flags-list">${flags.map(f=>`<li>${escapeHtml(f)}</li>`).join('')}</ul></div>`:'';

    return `<div class="dr-card" id="drcard_${id}">
      <div class="dr-card-top">
        <div class="dr-subject">${escapeHtml(e.subject||'(no subject)')}</div>
        <div class="dr-meta">${drTierBadge(e.sender_tier)}<span class="dr-timestamp">${escapeHtml(e.drafted_at||'')}</span></div>
      </div>
      ${confHtml}
      <div class="dr-draft-text" id="drtext_${id}" data-raw="${draftEsc}" onclick="toggleDraftExpand('${id}',event)">${draftEsc}</div>
      <span class="dr-expand-hint" id="drhint_${id}" onclick="toggleDraftExpand('${id}',event)">Show more</span>
      ${flagsHtml}
      <div class="dr-actions">
        <button class="dr-btn dr-btn-primary" onclick="copyDraftText('${id}',event)">Copy to clipboard</button>
        ${openLink}
        <button class="dr-btn" onclick="markDraft('${id}','${e.source_entry_id}','sent',event)">Mark sent</button>
        <button class="dr-btn dr-btn-danger" onclick="markDraft('${id}','${e.source_entry_id}','discarded',event)">Discard</button>
      </div>
    </div>`;
  }).join('');

  panel.innerHTML=`<div class="dr-header"><div class="dr-title">Drafted Replies <span class="dr-count-badge" id="drCountBadge">${pending.length}</span></div></div><div class="dr-subtitle">Drafted by Lauren in Kevin's style — review, copy, and send from Outlook yourself. Nothing here sends anything automatically.</div><div id="drList">${cards}</div>`;
  _updateDraftedTabBadge(pending.length);
}

async function loadDraftedReplies(){
  try{
    const res=await fetch(DRAFTED_REPLIES_API+'?t='+Date.now(),{cache:'no-store'});
    if(!res.ok) throw new Error('fetch failed');
    const payload=await res.json();
    renderDraftedReplies(payload);
  }catch(e){
    console.warn('Drafted replies fetch failed',e);
    const panel=document.getElementById('draftedRepliesPanel');
    if(panel&&!panel.innerHTML) panel.innerHTML='<div class="dr-header"><div class="dr-title">Drafted Replies</div></div><div class="dr-empty">Unable to load right now.</div>';
  }
}
// First call happens from init(), after loadRemoteTicks() resolves -- calling
// it here unconditionally would race remote-tick loading and could briefly
// re-show a draft Kevin already marked sent/discarded on another machine.
setInterval(loadDraftedReplies, 60000);

/* CLOCK */
function updateWiClock(){
  var n=new Date();
  var time=n.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
  var date=n.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  var tel=document.getElementById('wi-clock-time');
  if(tel) tel.textContent=time;
  var del=document.getElementById('sidebarDate');
  if(del) del.textContent=date;
}
updateWiClock();
setInterval(updateWiClock,1000);
