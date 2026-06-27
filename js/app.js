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
let showingDoneItems=false;
function toggleShowDone(){
  const btn=document.getElementById('btn-show-done');
  showingDoneItems=!showingDoneItems;
  if(showingDoneItems){
    document.querySelectorAll('.card.done, .card-link.done, .pri-card.done').forEach(el=>el.classList.remove('card-hidden'));
    document.querySelectorAll('.priority-row.done').forEach(el=>el.style.display='flex');
    btn.textContent='Hide done';
  } else {
    document.querySelectorAll('.card.done, .card-link.done, .pri-card.done').forEach(el=>el.classList.add('card-hidden'));
    document.querySelectorAll('.priority-row.done').forEach(el=>el.style.display='none');
    btn.textContent='Show done';
  }
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
function toggleTick(id){
  const ticks=getTicks(), k=currentKey+'_'+id;
  ticks[k]=!ticks[k]; saveTicks(ticks);
  const cb=document.getElementById('cb_'+id);
  const item=document.getElementById('item_'+id);
  const prow=document.getElementById('prow_'+id);
  if(cb) cb.classList.toggle('checked',ticks[k]);
  if(item){
    const wrapper=item.closest('.card-link');
    if(ticks[k]){
      item.classList.add('done');
      if(!showingDoneItems) item.classList.add('card-hidden');
      else item.classList.remove('card-hidden');
      if(wrapper){
        wrapper.classList.add('done');
        if(!showingDoneItems) wrapper.classList.add('card-hidden');
        else wrapper.classList.remove('card-hidden');
      }
    } else {
      item.classList.remove('done','card-hidden');
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
function isTicked(id){if(!currentKey) return false; return !!getTicks()[currentKey+'_'+id];}
function badge(text,type){return text?`<span class="badge badge-${type||'gray'}">${text}</span>`:''}

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
    const cardHtml=`<div class="card${ticked?' done':''}" id="item_${id}"><div class="cb-wrap"><div class="cb${ticked?' checked':''}" id="cb_${id}" onclick="toggleTick('${id}');event.stopPropagation()"></div></div><div class="card-accent ac-${cls==="urgent"?"r":cls==="needs"?"o":cls==="fyi"?"b":"g"}"></div><div class="card-body"><div class="card-title">${item.title} ${badge(item.badge,item.badgeType)}${(()=>{if(!item.received)return '';const c=new Date();c.setDate(c.getDate()-4);c.setHours(0,0,0,0);return new Date(item.received+'T12:00:00')>=c?badge('NEW','green'):'';})()}</div>${item.sub?`<div class="card-sub">${sanitizeSub(item.sub)}</div>`:''}</div><div class="card-date">${item.received||''}</div></div>`;
    return hasLink?`<a class="card-link${ticked?' done':''}" href="javascript:void(0)" onclick="openEmail('${item.entry_id}',event)" ${dragAttrs}>${cardHtml}</a>`:`<div ${dragAttrs}>${cardHtml}</div>`;
  }).join('');
}

function renderSidebarCal(items, containerId){
  const el=document.getElementById(containerId);
  if(!el) return;
  if(!items||!items.length){el.innerHTML='<div style="padding:4px 18px 8px;font-size:11px;color:rgba(255,255,255,0.3);font-style:italic">None</div>';return;}
  el.innerHTML=items.map((c,i)=>`<div class="cal-item${i===0?' active':''}">${c.time?`<div class="cal-time">${c.time}</div>`:''}<div class="cal-title">${c.title}</div>${c.sub?`<div class="cal-sub">${c.sub}</div>`:''}${c.alert?`<div class="cal-alert">⚠ ${c.alert}</div>`:''}</div>`).join('');
}

function togglePriCard(i){
  const body=document.getElementById('pribody_'+i);
  const arrow=document.getElementById('priarrow_'+i);
  const isOpen=body.style.display==='block';
  body.style.display=isOpen?'none':'block';
  arrow.textContent=isOpen?'–':'▲';
}

// Priority drag-and-drop helpers
let _priDragState=null;
function _priGetKey(p){return(p.title||p.text||p.subject||'').toLowerCase().replace(/[^a-z0-9]/g,'').substring(0,40)||'item';}
function _priGetOverrides(){try{return JSON.parse(localStorage.getItem('workInbox_priOverrides_v1')||'{}');}catch(e){return{};}}
function _priSetOverride(key,sec){const o=_priGetOverrides();o[key]=sec;localStorage.setItem('workInbox_priOverrides_v1',JSON.stringify(o));}
function _priGetOrder(){try{return JSON.parse(localStorage.getItem('workInbox_priOrder_v1')||'{}');}catch(e){return{};}}
function _priSetOrder(pt,ptom,pw,pfyi,ur,nr){localStorage.setItem('workInbox_priOrder_v1',JSON.stringify({pt,ptom:ptom||[],pw,pfyi:pfyi||[],ur:ur||[],nr:nr||[]}));}
function _getCustomPri(){try{return JSON.parse(localStorage.getItem('workInbox_customPri_v1')||'[]');}catch(e){return[];}}
function _saveCustomPri(arr){localStorage.setItem('workInbox_customPri_v1',JSON.stringify(arr));}
function _addEmailCardToPriority(item,cls,sec){const arr=_getCustomPri();const priKey=_priGetKey(item);if(arr.findIndex(x=>x._priKey===priKey)<0){arr.push({...item,_priKey:priKey,_dfSec:sec,_cls:cls});_saveCustomPri(arr);}_priSetOverride(priKey,sec);}

function applyPriOverrides(data){
  const all=[...(data.prioritiesToday||[]).map(p=>({...p,_dfSec:'pt'})),...(data.prioritiesTomorrow||[]).map(p=>({...p,_dfSec:'ptom'})),...(data.prioritiesWeek||[]).map(p=>({...p,_dfSec:'pw'})),...(data.fyi||[]).map(p=>({...p,text:p.title,_dfSec:'pfyi'})),..._getCustomPri()];
  const ovr=_priGetOverrides(),ord=_priGetOrder(),secs={pt:[],ptom:[],pw:[],pfyi:[],ur:[],nr:[]};
  const validSecs=['pt','ptom','pw','pfyi','ur','nr'];
  const _seen=new Set();
  for(const item of all){const k=_priGetKey(item);if(_seen.has(k))continue;_seen.add(k);const s=ovr[k]||item._dfSec;(validSecs.includes(s)?secs[s]:secs.pw).push({...item,_priKey:k});}
  for(const s of validSecs){if(ord[s]&&ord[s].length){const om={};ord[s].forEach((k,i)=>om[k]=i);secs[s].sort((a,b)=>(om[a._priKey]??999)-(om[b._priKey]??999));}}
  return secs;
}

function priDragStart(e,sec,priKey){
  _priDragState={sec,priKey};
  e.dataTransfer.effectAllowed='move';
  e.dataTransfer.setData('text/plain',priKey);
  setTimeout(()=>{const el=document.querySelector(`.pri-card[data-prikey="${priKey}"]`);if(el)el.classList.add('pri-dragging');},0);
}
function priDragEnd(e){
  document.querySelectorAll('.pri-card.pri-dragging').forEach(el=>el.classList.remove('pri-dragging'));
  document.querySelectorAll('.pri-card.pri-drop-before,.pri-card.pri-drop-after').forEach(el=>el.classList.remove('pri-drop-before','pri-drop-after'));
  document.querySelectorAll('.pri-drop-zone.pri-zone-active').forEach(el=>el.classList.remove('pri-zone-active'));
  _priDragState=null;
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
function priCardDragOver(e,sec,priKey){
  if(!_priDragState&&!_emailDragData)return;
  e.preventDefault();e.stopPropagation();e.dataTransfer.dropEffect='move';
  if(_priDragState){document.querySelectorAll('.pri-card.pri-drop-before,.pri-card.pri-drop-after').forEach(el=>el.classList.remove('pri-drop-before','pri-drop-after'));const r=e.currentTarget.getBoundingClientRect();e.currentTarget.classList.add(e.clientY<r.top+r.height/2?'pri-drop-before':'pri-drop-after');}
}
function priCardDragLeave(e,priKey){
  if(e.relatedTarget&&e.currentTarget.contains(e.relatedTarget))return;
  e.currentTarget.classList.remove('pri-drop-before','pri-drop-after');
}
function priCardDrop(e,sec,priKey){
  e.preventDefault();e.stopPropagation();
  if(_emailDragData){const{item,cls}=_emailDragData;_addEmailCardToPriority(item,cls,sec);emailCardDragEnd(e);if(window._wipData&&window._wipKey)renderBriefing(window._wipData,window._wipKey);return;}
  if(!_priDragState)return;
  const{sec:fromSec,priKey:fromKey}=_priDragState;
  const before=e.currentTarget.classList.contains('pri-drop-before');
  const allSecs=['pt','ptom','pw','pfyi','ur','nr'];
  const sk={};
  allSecs.forEach(s=>{sk[s]=Array.from(document.querySelectorAll(`.pri-drop-zone[data-sec="${s}"] .pri-card`)).map(c=>c.dataset.prikey);});
  if(fromSec!==sec)_priSetOverride(fromKey,sec);
  allSecs.forEach(s=>{sk[s]=sk[s].filter(k=>k!==fromKey);});
  const ti=sk[sec].indexOf(priKey);
  sk[sec].splice(ti>=0?(before?ti:ti+1):sk[sec].length,0,fromKey);
  _priSetOrder(sk.pt,sk.ptom,sk.pw,sk.pfyi,sk.ur,sk.nr);
  priDragEnd(e);
  if(window._wipData&&window._wipKey)renderBriefing(window._wipData,window._wipKey);
}
function priZoneDragOver(e,sec){
  if(!_priDragState&&!_emailDragData)return;
  e.preventDefault();e.dataTransfer.dropEffect='move';
  document.querySelector(`.pri-drop-zone[data-sec="${sec}"]`)?.classList.add('pri-zone-active');
}
function priZoneDragLeave(e,sec){
  const z=document.querySelector(`.pri-drop-zone[data-sec="${sec}"]`);
  if(z&&!z.contains(e.relatedTarget))z.classList.remove('pri-zone-active');
}
function priZoneDrop(e,sec){
  e.preventDefault();
  if(_emailDragData){const{item,cls}=_emailDragData;_addEmailCardToPriority(item,cls,sec);emailCardDragEnd(e);if(window._wipData&&window._wipKey)renderBriefing(window._wipData,window._wipKey);return;}
  if(!_priDragState)return;
  const{sec:fromSec,priKey:fromKey}=_priDragState;
  if(fromSec!==sec)_priSetOverride(fromKey,sec);
  const allSecs=['pt','ptom','pw','pfyi','ur','nr'];
  const sk={};
  allSecs.forEach(s=>{sk[s]=Array.from(document.querySelectorAll(`.pri-drop-zone[data-sec="${s}"] .pri-card`)).map(c=>c.dataset.prikey);});
  allSecs.forEach(s=>{sk[s]=sk[s].filter(k=>k!==fromKey);});
  sk[sec].push(fromKey);
  _priSetOrder(sk.pt,sk.ptom,sk.pw,sk.pfyi,sk.ur,sk.nr);
  priDragEnd(e);
  if(window._wipData&&window._wipKey)renderBriefing(window._wipData,window._wipKey);
}

function renderPriorityCards(priorities,key,sec){
  if(!priorities||!priorities.length) return '<div class="pri-zone-empty">Drop items here</div>';
  const _mo=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  // Canvas text measurement for smart one-line vs two-line decision
  if(!renderPriorityCards._canvas){renderPriorityCards._canvas=document.createElement('canvas');renderPriorityCards._ctx=renderPriorityCards._canvas.getContext('2d');}
  const _mctx=renderPriorityCards._ctx;
  function _tw(t,f){_mctx.font=f;return _mctx.measureText(t).width;}
  const _titleFont='600 15px Inter,sans-serif';
  const _gridCellW=Math.max(280,(window.innerWidth-340-72-24)/2);
  const _cardInnerW=_gridCellW-32;
  const _fixedW=55;
  const _recentPfxs4=Array.from({length:4},(_,i)=>{const d=new Date();d.setDate(d.getDate()-i);return'['+String(d.getDate()).padStart(2,'0')+' '+_mo[d.getMonth()]+' '+d.getFullYear()+']';});
  const _cutoff=new Date();_cutoff.setDate(_cutoff.getDate()-4);_cutoff.setHours(0,0,0,0);
  const _mo2={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  function _firstActionDate(actions){if(!actions||!actions.length)return null;const m=actions[0].match(/^\[(\d{1,2}) (\w{3}) (\d{4})\]/);if(!m)return null;return new Date(parseInt(m[3]),_mo2[m[2]],parseInt(m[1]),12,0,0);}
  return priorities.map((p,i)=>{
    const priKey=p._priKey||_priGetKey(p);
    const id='pri_'+sec+'_'+i, ticked=isTicked(id);
    const _rawTitle=p.title||p.text||'(untitled)';
    const _titleParts=_rawTitle.split(' -- ');
    const titleText=_titleParts[0];
    const titleSub=_titleParts.length>1?_titleParts.slice(1).join(' -- '):'';
    const aiBadge=(p.badge&&sec!=='pfyi')?badge(p.badge,p.badgeType||'gray'):'';
    const createdDate=p.dateAdded?new Date(p.dateAdded+'T12:00:00'):_firstActionDate(p.actions);
    const newBadge=(!aiBadge&&createdDate&&createdDate>=_cutoff)?badge('NEW','green'):'';
    const updBadge=(!aiBadge&&!newBadge&&p.actions&&p.actions.some(a=>_recentPfxs4.some(pfx=>a.startsWith(pfx))))?badge('UPDATED','blue'):'';
    const theBadge=aiBadge||newBadge||updBadge;
    // Measure whether combined title fits on one line
    const _badgeW=theBadge?82:0;
    const _btnW=(p.entry_id||p.entryId)?118:0;
    const _ccBtnW=p.id?96:0;
    const _availW=_cardInnerW-_fixedW-_badgeW-_btnW-_ccBtnW;
    const _combined=titleSub?titleText+' – '+titleSub:titleText;
    const _oneLine=!titleSub||_tw(_combined,_titleFont)<=_availW;
    const summaryHtml=p.ai_summary?`<div class="pri-ai-summary">${sanitizeSub(p.ai_summary)}</div>`:'';
    const metaHtml=`<div class="pri-meta">
      ${p.source?`<span class="pri-meta-item"><span class="pri-meta-label">Source</span>${p.source}</span>`:''}
      ${p.dateAdded?`<span class="pri-meta-item"><span class="pri-meta-label">Added</span>${p.dateAdded}</span>`:''}
      ${p.notes?`<div class="pri-meta-notes"><span class="pri-meta-label">Notes</span><div class="pri-notes-text">${sanitizeSub(p.notes)}</div></div>`:''}
    </div>`;
    return `<div class="pri-card${ticked?' done':''}" id="item_${id}" data-prikey="${priKey}" data-sec="${sec}" draggable="true" ondragstart="priDragStart(event,'${sec}','${priKey}')" ondragend="priDragEnd(event)" ondragover="priCardDragOver(event,'${sec}','${priKey}')" ondragleave="priCardDragLeave(event,'${priKey}')" ondrop="priCardDrop(event,'${sec}','${priKey}')">
      <div class="pri-card-header" onclick="togglePriCard('${sec}_${i}')">
        <div class="pri-drag-handle" onclick="event.stopPropagation()">⣿⣿</div>
        <div class="cb-wrap"><div class="cb${ticked?' checked':''}" id="cb_${id}" onclick="toggleTick('${id}');event.stopPropagation()"></div></div>
        <div class="pri-card-title-wrap">${_oneLine?`<div class="pri-card-title">${_combined}</div>`:`<div class="pri-card-title-main">${titleText}</div><div class="pri-card-title-sub">${titleSub}</div>`}</div>
        ${theBadge}
        ${(p.entry_id||p.entryId)?`<button class="sg-btn" onclick="openEmail('${p.entry_id||p.entryId}',event)">Open email</button>`:''}
        ${p.id?`<button class="sg-btn sg-btn-cc" onclick="window.open('https://cc.lelitte.co.uk/#${p.id}','_blank');event.stopPropagation()">CC →</button>`:''}
        <span class="pri-arrow" id="priarrow_${sec}_${i}">–</span>
      </div>
      <div class="pri-card-body" id="pribody_${sec}_${i}">${summaryHtml}${metaHtml}</div>
    </div>`;
  }).join('');
}

function renderBriefing(data,key){
  currentData=data; currentKey=key;
  window._wipData=data; window._wipKey=key;

  // Header
  document.getElementById('pageTitle').textContent=getGreeting();
  document.getElementById('headerDate').textContent=data.date+(data.subtitle?' · '+data.subtitle:'');
  document.getElementById('sidebarDate').textContent=data.date||'';;
  const stamp=document.getElementById('refresh-stamp');
  if(stamp&&data.refreshed_at) stamp.textContent='Last refreshed: '+data.refreshed_at;

  // Sidebar calendar
  renderSidebarCal(data.calToday,'calTodaySidebar');
  renderSidebarCal(data.calTomorrow,'calTomorrowSidebar');

  // Absences
  const absEl=document.getElementById('absencesSidebar');
  if(data.absences&&data.absences.length){
    absEl.innerHTML='<ul class="abs-list">'+data.absences.map(a=>`<li>${a}</li>`).join('')+'</ul>';
  } else {
    absEl.innerHTML='<span style="font-size:11px;color:rgba(255,255,255,0.3);font-style:italic">None recorded</span>';
  }

  // Context bar
  const ctxEl=document.getElementById('contextBar');
  if(data.context){
    const sentences=data.context.split(/(?<=\.)\ +/).filter(s=>s.trim().length>0);
    const items=sentences.map(s=>`<div class="context-bar-item"><span class="context-bar-bullet">—</span><span>${s.trim()}</span></div>`).join('');
    ctxEl.innerHTML=`<div class="context-bar"><div class="context-bar-title">Context</div>${items}</div>`;
  } else { ctxEl.innerHTML=''; }

  // Main inbox — 2×2 grid
  const priSecs=applyPriOverrides(data);
  const left=`<div class="inbox-grid">
    <div class="inbox-grid-cell">
      <div class="sec-head"><span class="sec-dot dot-r"></span><span class="sec-lbl">Priority actions – today</span><span class="sec-rule"></span></div>
      <div class="pri-drop-zone" data-sec="pt" ondragover="priZoneDragOver(event,'pt')" ondragleave="priZoneDragLeave(event,'pt')" ondrop="priZoneDrop(event,'pt')">${priSecs.pt.length?renderPriorityCards(priSecs.pt,key,'pt'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
    </div>
    <div class="inbox-grid-cell">
      <div class="sec-head"><span class="sec-dot dot-o"></span><span class="sec-lbl">Priority actions – tomorrow</span><span class="sec-rule"></span></div>
      <div class="pri-drop-zone" data-sec="ptom" ondragover="priZoneDragOver(event,'ptom')" ondragleave="priZoneDragLeave(event,'ptom')" ondrop="priZoneDrop(event,'ptom')">${priSecs.ptom.length?renderPriorityCards(priSecs.ptom,key,'ptom'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
    </div>
    <div class="inbox-grid-cell">
      <div class="sec-head"><span class="sec-dot dot-green"></span><span class="sec-lbl">Priority actions – this week</span><span class="sec-rule"></span></div>
      <div class="pri-drop-zone" data-sec="pw" ondragover="priZoneDragOver(event,'pw')" ondragleave="priZoneDragLeave(event,'pw')" ondrop="priZoneDrop(event,'pw')">${priSecs.pw.length?renderPriorityCards(priSecs.pw,key,'pw'):'<div class="pri-zone-empty">Drop items here</div>'}</div>
    </div>
    <div class="inbox-grid-cell">
      <div class="sec-head"><span class="sec-dot dot-g"></span><span class="sec-lbl">FYI – no action needed</span><span class="sec-rule"></span></div>
      <div class="pri-drop-zone" data-sec="pfyi" ondragover="priZoneDragOver(event,'pfyi')" ondragleave="priZoneDragLeave(event,'pfyi')" ondrop="priZoneDrop(event,'pfyi')">${priSecs.pfyi.length?renderPriorityCards(priSecs.pfyi,key,'pfyi'):'<div class="pri-zone-empty">Drop items here to park</div>'}</div>
    </div>
  </div>`;
  document.getElementById('inboxCol').innerHTML=left;
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

init();

// Tasks widget
async function loadTasksWidget(){
  var el=document.getElementById('tasksWidget');
  var dot=document.getElementById('tasksLiveDot');
  if(!el)return;
  try{
    var res=await fetch('https://github-proxy.lelitte.co.uk/command-centre/data/tasks.json?t='+Date.now(),{cache:'no-store'});
    if(!res.ok)throw new Error('fetch failed');
    var data=await res.json();
    var tasks=Array.isArray(data)?data:(data.tasks||[]);
    var todayCount=tasks.filter(function(t){return t.tier==='today';}).length;
    var tomorrowCount=tasks.filter(function(t){return t.tier==='tomorrow';}).length;
    var weekCount=tasks.filter(function(t){return t.tier==='week';}).length;
    var todoCount=tasks.filter(function(t){return (t.actions||[]).some(function(a){return a.includes('[TODO]');});}).length;
    el.innerHTML=
      '<div class="tasks-widget-row"><span class="tasks-widget-name">Today</span><span class="tasks-widget-count">'+todayCount+'</span></div>'+
      '<div class="tasks-widget-row"><span class="tasks-widget-name">Tomorrow</span><span class="tasks-widget-count">'+tomorrowCount+'</span></div>'+
      '<div class="tasks-widget-row"><span class="tasks-widget-name">This week</span><span class="tasks-widget-count">'+weekCount+'</span></div>'+
      '<div class="tasks-widget-row"><span class="tasks-widget-name">Actions due</span><span class="tasks-widget-count">'+todoCount+'</span></div>'+
      '<a class="tasks-widget-link" href="https://cc.lelitte.co.uk/" target="_blank">→ Open command centre</a>';
    if(dot){dot.style.background='#4ade80';dot.classList.add('pulsing');setTimeout(function(){dot.classList.remove('pulsing');},700);}
  }catch(e){
    el.innerHTML='<div class="tasks-widget-unavailable">Tasks unavailable</div>';
    if(dot)dot.style.background='#f87171';
  }
}
loadTasksWidget();
setInterval(loadTasksWidget, 30000);