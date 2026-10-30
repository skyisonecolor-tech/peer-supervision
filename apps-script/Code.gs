/**
 * 2026-2학기 동료장학 수업공개 취합 — Google Apps Script 서버
 * 구글 시트 > 확장 프로그램 > Apps Script 에 붙여 넣고 웹 앱으로 배포합니다.
 * 코드를 바꾼 뒤에는 [배포 > 배포 관리 > 수정 > 새 버전]으로 다시 배포하세요(주소 유지).
 *
 * [관리자 비밀번호 설정] 이 코드에는 비밀번호가 들어 있지 않습니다.
 *   Apps Script 왼쪽 ⚙ 프로젝트 설정 > 스크립트 속성 > 속성 추가
 *   속성: ADMIN_PW   값: (사용할 관리자 비밀번호)
 */
function adminPw(){ return PropertiesService.getScriptProperties().getProperty('ADMIN_PW') || ''; }
// 비밀번호를 10번 틀리면 10분 동안 관리자 기능 잠금
function authOk(pw){
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pwFails')||0);
  if(fails>=10) throw new Error('비밀번호를 여러 번 틀려 10분 동안 잠겼습니다. 잠시 후 다시 시도해 주세요.');
  const real = adminPw();
  if(!real) throw new Error('서버에 관리자 비밀번호(스크립트 속성 ADMIN_PW)가 설정되지 않았습니다.');
  if(String(pw)===real){ cache.remove('pwFails'); return true; }
  cache.put('pwFails', String(fails+1), 600);
  return false;
}
const SHEETS = { peer:'동료장학', open:'전교사공개', obs:'수업참관', roster:'명단', settings:'설정' };
const COLS = ['key','교과','성명','공개날짜(동료장학)','교시','장소','과목','주제','최초제출','최종수정','교사수정횟수','관리자수정','이력(JSON)'];
const OCOLS = ['key','교과','성명','동료장학 참관(JSON)','전교사 참관(JSON)','최초제출','최종수정','교사수정횟수','관리자수정','이력(JSON)'];
const PEER_CAP = 4, PEER_CAP_SAME = 5, OPEN_CAP = 14;
const DEFAULT_SETTINGS = {year:2026, semester:2, timetable:[], ttSwap:[], inputState:'before', obsOpen:false, obsLock:false, openDate:'', excludeDates:[], refDates:[], peerObsCount:2, openObsCount:1};

function doGet(){ return out({ok:true, msg:'동료장학 취합 서버 동작 중'}); }

function doPost(e){
  let p;
  try{ p = JSON.parse(e.postData.contents); }catch(err){ return out({error:'잘못된 요청'}); }
  // 읽기는 잠금 없이 (여러 선생님이 동시에 새로고침해도 줄 서지 않도록) + 짧은 캐시
  if(p.action==='load'){ try{ return out(loadCached()); }catch(err){ return out({error:String(err.message||err)}); } }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try{
    const needAdmin = ['saveRoster','adminUpdate','adminDelete','setSetting','adminObsUpdate','adminLessonObs','reset'];
    if(p.action==='checkPw') return out({ok: authOk(p.pw)});
    if(needAdmin.indexOf(p.action)>=0 && !authOk(p.pw)) return out({error:'관리자 비밀번호 오류'});
    switch(p.action){
      case 'submit':         return out(submit(p));
      case 'reset':          return out(resetData(p.what||[]));
      case 'submitObs':      return out(submitObs(p));
      case 'adminUpdate':    return out(adminUpdate(p));
      case 'adminObsUpdate': return out(adminObsUpdate(p));
      case 'adminLessonObs': return out(adminLessonObs(p));
      case 'adminDelete':    deleteRec(p.type, p.key); return out({status:'ok'});
      case 'saveRoster':     saveRoster(p.roster); return out({status:'ok'});
      case 'setSetting':     setSetting(p.k, p.v); return out({status:'ok'});
    }
    return out({error:'알 수 없는 요청'});
  }catch(err){ return out({error:String(err.message||err)}); }
  finally{ clearLoadCache(); lock.releaseLock(); }
}

/* ---------- 읽기 캐시 (3초) ---------- */
function loadCached(){
  const c = CacheService.getScriptCache();
  const hit = c.get('loadAll'); if(hit) return JSON.parse(hit);
  const data = loadAll();
  try{ c.put('loadAll', JSON.stringify(data), 3); }catch(e){}   // 100KB 넘으면 캐시하지 않음
  return data;
}
function clearLoadCache(){ try{ CacheService.getScriptCache().remove('loadAll'); }catch(e){} }

/* ---------- 초기화 (지우기 전에 같은 시트 파일 안에 백업 탭 생성) ---------- */
function backupSheet(name){
  const sh = ss().getSheetByName(name); if(!sh || sh.getLastRow()<=1) return null;
  const nm = name+'_백업_'+Utilities.formatDate(new Date(),'Asia/Seoul','yyMMdd_HHmmss');
  sh.copyTo(ss()).setName(nm); return nm;
}
function resetData(what){
  const has = k=>what.indexOf(k)>=0; const backups = [];
  const clearType = (type)=>{ const b=backupSheet(SHEETS[type]); if(b) backups.push(b);
    const sh = sheet(SHEETS[type], colsOf(type)); sh.clearContents(); sh.appendRow(colsOf(type)); };
  if(has('peer')) clearType('peer');
  if(has('open')) clearType('open');
  if(has('obs'))  clearType('obs');
  if(has('roster')){ const b=backupSheet(SHEETS.roster); if(b) backups.push(b); saveRoster([]); }
  if(has('timetable')||has('dates')||has('phase')){ const b=backupSheet(SHEETS.settings); if(b) backups.push(b); }
  if(has('timetable')){ setSetting('timetable', []); setSetting('ttSwap', []); }
  if(has('dates')){ setSetting('openDate',''); setSetting('excludeDates',[]); setSetting('refDates',[]); }
  if(has('phase')){ setSetting('obsLock', false); setSetting('obsOpen', false); setSetting('inputState','before'); }
  return {status:'ok', backups:backups};
}

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name, header){
  let sh = ss().getSheetByName(name);
  if(!sh){ sh = ss().insertSheet(name); if(header){ sh.appendRow(header); sh.setFrozenRows(1); } }
  return sh;
}
function now(){ return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'); }
function normName(n){ return String(n||'').replace(/\s+/g,''); }
function keyOf(subj, name){ return String(subj||'').trim()+'|'+normName(name); }
function colsOf(type){ return type==='obs' ? OCOLS : COLS; }

/* ---------- 읽기/쓰기 ---------- */
function readType(type){
  const C = colsOf(type);
  const sh = sheet(SHEETS[type], C);
  const v = sh.getDataRange().getDisplayValues(); const box = {};
  for(let i=1;i<v.length;i++){
    const r=v[i]; if(!r[0]) continue;
    if(type==='obs'){
      box[r[0]] = {key:r[0], subj:r[1], name:r[2], peer:r[3]?JSON.parse(r[3]):[], open:r[4]?JSON.parse(r[4]):[],
        createdAt:r[5], updatedAt:r[6], revisions:Number(r[7]||0), adminEdited:r[8]==='Y', history:r[9]?JSON.parse(r[9]):[], _row:i+1};
    } else {
      box[r[0]] = {key:r[0], subj:r[1], name:r[2], date:r[3], period:r[4], place:r[5], course:r[6], topic:r[7],
        createdAt:r[8], updatedAt:r[9], revisions:Number(r[10]||0), adminEdited:r[11]==='Y', history:r[12]?JSON.parse(r[12]):[], _row:i+1};
    }
  }
  return box;
}
function writeRec(type, rec, row){
  const C = colsOf(type); const sh = sheet(SHEETS[type], C);
  const vals = type==='obs'
    ? [[rec.key, rec.subj, rec.name, JSON.stringify(rec.peer||[]), JSON.stringify(rec.open||[]), rec.createdAt, rec.updatedAt, rec.revisions||0, rec.adminEdited?'Y':'', JSON.stringify(rec.history||[])]]
    : [[rec.key, rec.subj, rec.name, rec.date||'', rec.period, rec.place, rec.course, rec.topic, rec.createdAt, rec.updatedAt, rec.revisions||0, rec.adminEdited?'Y':'', JSON.stringify(rec.history||[])]];
  const r = row || sh.getLastRow()+1;
  sh.getRange(r,1,1,C.length).setNumberFormat('@').setValues(vals);
}
function strip(r){ const o={}; Object.keys(r).forEach(k=>{ if(k!=='history' && k!=='_row') o[k]=r[k]; }); return o; }
function clean(box){ Object.keys(box).forEach(k=>delete box[k]._row); return box; }
function deleteRec(type, key){ const r = readType(type)[key]; if(r) sheet(SHEETS[type], colsOf(type)).deleteRow(r._row); }

/* ---------- 동료장학 / 전교사 공개 제출 ---------- */
function submit(p){
  const st = getSettings();
  if(st.obsOpen || st.inputState!=='open') throw new Error('지금은 계획 입력 기간이 아닙니다.');
  if(p.type==='peer'){ const ex = findExclude(p.rec.date, st.excludeDates); if(ex) throw new Error('수업공개를 할 수 없는 날짜입니다: '+ex.reason); }
  if(p.type==='open'){
    p.rec.date='';
    const rr = readRoster().find(r=>r.subj===p.rec.subj && normName(r.name)===normName(p.rec.name));
    if(!rr || !rr.openTarget) throw new Error('전교사 공개수업 대상이 아닙니다.');
  }
  const box = readType(p.type); const k = keyOf(p.rec.subj, p.rec.name); const old = box[k];
  if(old && !p.force){ delete old._row; return {status:'exists', existing:old}; }
  const t = now();
  const rec = Object.assign({}, p.rec, {key:k, createdAt: old?old.createdAt:t, updatedAt:t,
    revisions: old ? old.revisions+1 : 0, adminEdited: old?old.adminEdited:false,
    history: old ? old.history.concat([{at:t, by:'교사', before:strip(old)}]) : []});
  writeRec(p.type, rec, old?old._row:null);
  return {status:'ok', overwritten:!!old};
}
function adminUpdate(p){
  const box = readType(p.type); const old = box[p.key];
  if(!old) throw new Error('자료를 찾지 못했습니다.');
  const nk = keyOf(p.rec.subj, p.rec.name);
  if(nk!==p.key && box[nk]) throw new Error('같은 교과·성명의 자료가 이미 있습니다.');
  const t = now();
  const rec = Object.assign({}, old, p.rec, {key:nk, updatedAt:t, adminEdited:true,
    history: old.history.concat([{at:t, by:'관리자', before:strip(old)}])});
  writeRec(p.type, rec, old._row);
  return {status:'ok'};
}

/* ---------- 수업 참관 ---------- */
function lessonLabel(type, L, st){
  const d = type==='peer' ? L.date : st.openDate;
  return (d?d.slice(5).replace('-','/')+' ':'') + L.period + ' ' + L.subj + ' ' + L.name + '(' + L.course + ')';
}
function submitObs(p){
  const st = getSettings();
  if(!st.obsOpen || st.obsLock) throw new Error('지금은 수업 참관 신청 기간이 아닙니다.');
  const subj = p.rec.subj, name = normName(p.rec.name), me = keyOf(subj, name);
  if(!readRoster().some(r=>r.subj===subj && normName(r.name)===name)) throw new Error('이름이 정확하지 않습니다.');
  const selP = p.rec.peer||[], selO = p.rec.open||[];
  if(selP.length!==Number(st.peerObsCount) || selO.length!==Number(st.openObsCount))
    throw new Error('동료장학 '+st.peerObsCount+'개, 전교사 공개 '+st.openObsCount+'개를 정확히 선택해야 합니다.');
  if(new Set(selP).size!==selP.length || new Set(selO).size!==selO.length) throw new Error('같은 수업이 두 번 선택되었습니다.');
  const peer = readType('peer'), open = readType('open'), obs = readType('obs');
  const old = obs[me];
  // 시간 겹침
  const times = {};
  const addT = (t, label)=>{ if(times[t]) throw new Error('시간이 겹칩니다: '+label+' / '+times[t]); times[t]=label; };
  if(peer[me]) addT(peer[me].date+'#'+parseInt(peer[me].period), '본인 동료장학 수업');
  if(open[me] && st.openDate) addT(st.openDate+'#'+parseInt(open[me].period), '본인 전교사 공개 수업');
  const fails = [];
  selP.forEach(k=>{
    const L = peer[k]; if(!L) throw new Error('없어진 수업이 있습니다. 새로고침 후 다시 선택해 주세요.');
    const bz = ttBusyAt(st, subj, name, L.date, L.period); if(bz) throw new Error('본인 수업 시간입니다('+bz+'): '+lessonLabel('peer',L,st));
    if(k===me) throw new Error('본인 수업은 참관할 수 없습니다.');
    addT(L.date+'#'+parseInt(L.period), lessonLabel('peer',L,st));
    const n = Object.keys(obs).filter(ok=>ok!==me && obs[ok].peer.indexOf(k)>=0).length;
    const cap = (L.subj===subj) ? PEER_CAP_SAME : PEER_CAP;
    if(n>=cap) fails.push(lessonLabel('peer',L,st));
  });
  selO.forEach(k=>{
    const L = open[k]; if(!L) throw new Error('없어진 수업이 있습니다. 새로고침 후 다시 선택해 주세요.');
    const bz = ttBusyAt(st, subj, name, st.openDate, L.period); if(bz) throw new Error('본인 수업 시간입니다('+bz+'): '+lessonLabel('open',L,st));
    if(k===me) throw new Error('본인 수업은 참관할 수 없습니다.');
    if(st.openDate) addT(st.openDate+'#'+parseInt(L.period), lessonLabel('open',L,st));
    const n = Object.keys(obs).filter(ok=>ok!==me && obs[ok].open.indexOf(k)>=0).length;
    if(n>=OPEN_CAP) fails.push(lessonLabel('open',L,st));
  });
  if(fails.length) return {status:'full', full:fails};
  if(old && !p.force){ delete old._row; return {status:'exists', existing:old}; }
  const t = now();
  const rec = {key:me, subj:subj, name:name, peer:selP, open:selO, createdAt: old?old.createdAt:t, updatedAt:t,
    revisions: old?old.revisions+1:0, adminEdited: old?old.adminEdited:false,
    history: old ? old.history.concat([{at:t, by:'교사', before:strip(old)}]) : []};
  writeRec('obs', rec, old?old._row:null);
  return {status:'ok', overwritten:!!old};
}
function adminObsUpdate(p){
  const obs = readType('obs'); const old = obs[p.key]; const t = now();
  let rec;
  if(old) rec = Object.assign({}, old, {peer:p.rec.peer, open:p.rec.open, updatedAt:t, adminEdited:true,
    history: old.history.concat([{at:t, by:'관리자', before:strip(old)}])});
  else rec = {key:p.key, subj:p.rec.subj, name:p.rec.name, peer:p.rec.peer, open:p.rec.open, createdAt:t, updatedAt:t, revisions:0, adminEdited:true, history:[]};
  writeRec('obs', rec, old?old._row:null);
  return {status:'ok'};
}

// 수업 하나의 참관교사 목록을 관리자가 통째로 지정 (정원·횟수 제한 없음)
function adminLessonObs(p){
  const obs = readType('obs'); const t = now();
  const want = {}; (p.people||[]).forEach(x=>want[x.key]=x);
  const keys = Object.keys(obs).concat(Object.keys(want).filter(k=>!obs[k]));
  keys.forEach(k=>{
    const o = obs[k]; const has = !!(o && o[p.type].indexOf(p.lkey)>=0), should = !!want[k];
    if(has===should) return;
    if(o){
      const rec = Object.assign({}, o, {updatedAt:t, adminEdited:true, history:o.history.concat([{at:t, by:'관리자', before:strip(o)}])});
      rec[p.type] = should ? o[p.type].concat([p.lkey]) : o[p.type].filter(x=>x!==p.lkey);
      writeRec('obs', rec, o._row);
    } else {
      const w = want[k];
      writeRec('obs', {key:k, subj:w.subj, name:w.name, peer:p.type==='peer'?[p.lkey]:[], open:p.type==='open'?[p.lkey]:[], createdAt:t, updatedAt:t, revisions:0, adminEdited:true, history:[]}, null);
    }
  });
  return {status:'ok'};
}

/* ---------- 시간표 ---------- */
function weekdayOf(date, swaps){
  if(!date) return '';
  for(const x of (swaps||[])) if(x.date===date) return x.dow;
  const a = date.split('-').map(Number); return ['일','월','화','수','목','금','토'][new Date(a[0],a[1]-1,a[2]).getDay()];
}
function ttFind(tt, subj, name){
  const n = normName(name); const c = (tt||[]).filter(x=>x.n===n);
  if(!c.length) return null;
  return c.find(x=>x.s===subj) || (c.length===1?c[0]:null) || c.find(x=>subj.indexOf(x.s)>=0 || x.s.indexOf(subj.slice(0,2))>=0) || c[0];
}
function ttBusyAt(st, subj, name, date, period){
  const t = ttFind(st.timetable, subj, name); if(!t || !date) return null;
  const wd = weekdayOf(date, st.ttSwap), p = parseInt(period);
  return t.b.indexOf(wd+p)>=0 ? wd+'요일 '+p+'교시' : null;
}

/* ---------- 명단 ---------- */
function readRoster(){
  const v = sheet(SHEETS.roster, ['교과','교사','전교사 공개수업 대상']).getDataRange().getDisplayValues();
  return v.slice(1).filter(r=>r[0]).map(r=>({subj:r[0], name:r[1], openTarget:r[2]==='Y'}));
}
function saveRoster(list){
  const sh = sheet(SHEETS.roster, ['교과','교사','전교사 공개수업 대상']);
  sh.clearContents(); sh.appendRow(['교과','교사','전교사 공개수업 대상']);
  if(list.length) sh.getRange(2,1,list.length,3).setNumberFormat('@').setValues(list.map(r=>[r.subj, r.name, r.openTarget?'Y':'']));
}

/* ---------- 설정 (값은 JSON 문자열로 저장) ---------- */
function getSettings(){
  const v = sheet(SHEETS.settings, ['항목','값']).getDataRange().getValues();
  const o = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  v.slice(1).forEach(r=>{
    if(!r[0]) return; const val = r[1];
    if(val===true || val===false){ o[r[0]] = val; return; }
    try{ o[r[0]] = JSON.parse(val); }catch(e){ o[r[0]] = (val==='TRUE'||val==='true') ? true : val; }
  });
  return o;
}
function setSetting(k, val){
  const st = getSettings();
  if(k==='obsOpen' && val===true && st.inputState!=='closed') throw new Error('계획 입력을 먼저 마감해야 수업 참관을 시작할 수 있습니다.');
  if(k==='inputState' && val==='open' && st.obsOpen) throw new Error('수업 참관을 먼저 꺼야 입력을 다시 열 수 있습니다.');
  const sh = sheet(SHEETS.settings, ['항목','값']); const v = sh.getDataRange().getValues();
  const json = JSON.stringify(val);
  for(let i=1;i<v.length;i++) if(v[i][0]===k){ sh.getRange(i+1,2).setNumberFormat('@').setValue(json); return; }
  const r = sh.getLastRow()+1; sh.getRange(r,1).setValue(k); sh.getRange(r,2).setNumberFormat('@').setValue(json);
}
function findExclude(date, list){
  if(!date) return null;
  for(const x of (list||[])){ if(date>=x.start && date<=(x.end||x.start)) return x; }
  return null;
}

function loadAll(){
  return {roster:readRoster(), peer:clean(readType('peer')), open:clean(readType('open')), obs:clean(readType('obs')), settings:getSettings()};
}
