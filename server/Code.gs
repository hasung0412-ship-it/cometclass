/* 배곧해솔중학교 정보 학습 서버 (Google Apps Script)
   사용법: 스프레드시트 > 확장 프로그램 > Apps Script 에 전체 붙여넣기
   1) 함수 선택창에서 '초기설정' 실행 (처음 1회)  2) 배포 > 새 배포 > 웹 앱 */

const FREE = ['설명쓰기', '서술', '기타']; // 채점하지 않고 학생이 쓴 내용을 그대로 저장하는 유형
const ACT = { 개념: '개념익히기', 해석: '해석해보기', 작성: '응용해보기' };

const SERVER_VERSION='v91'; // 웹앱 주소를 브라우저로 열었을 때 'OK v41'가 보이면 최신 버전이 배포된 것입니다
function doGet() { return ContentService.createTextOutput('OK ' + SERVER_VERSION); }
function doPost(e) {
  let out;
  try { out = route(JSON.parse(e.postData.contents)); }
  catch (err) { out = { error: '서버 오류: ' + err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function route(d) {
  const A = d.action;
  if (A === 'login') return login(d);
  if (A === 'forgot') return forgot(d);
  if (A === 'years') return years();
  const u = session(d.token);
  if (!u) return { error: 'auth' };
  if (u.must && A !== 'changePw' && A !== 'me') return { error: '비밀번호를 먼저 바꿔 주세요.' };
  if (A === 'me') return u;
  if (A === 'changePw') return changePw(u, d);
  if (A === 'subjects') return { subjects: rowsC('과목').filter(r => subOK(u, r)).map(r => {
    const o = { name: String(r['과목']).replace(/[\u200B-\u200D\uFEFF]/g, '').trim(), icon: r['아이콘'] || '📘', ready: structure(r['과목']).length > 0, names: { 개념: r['개념이름'] || '', 해석: r['해석이름'] || '', 작성: r['작성이름'] || '' } };
    if (u.role === '관리자') o.why = why(r['과목']); // 관리자에게만 이유를 보여 줌
    return o;
  }) };
  if (['menu', 'lesson', 'submit', 'draft', 'sample'].includes(A) && !subAllowed(u, d.sub)) return { error: '이 과목은 입장할 수 없어요.' };
  if (A === 'toc') return toc(u);
  if (A === 'draft') return saveDraft(u, d);
  if (A === 'timer') return saveTimer(u, d);
  if (A === 'timerGet') return getTimer(u, d);
  if (A === 'info') return info();
  if (A === 'myGrades') return myGrades(u, d);
  if (A === 'menu') return menu(u, d);
  if (A === 'lesson') return lesson(u, d);
  if (A === 'submit') return submit(u, d);
  if (A === 'sample') return sampleCode(u, d);
  if (u.role !== '관리자') return { error: '권한이 없어요.' };
  if (A === 'classes') return classes(d.sub, d.year);
  if (A === 'classTable') return classTable(d.cls, d.sub, d.year);
  if (A === 'detail') return detail(d.id);
  if (A === 'resetPw') return resetPw(d.id);
  return { error: '알 수 없는 요청이에요.' };
}

/* ---------- 공통 도구 ---------- */
const sh = n => SpreadsheetApp.getActive().getSheetByName(n);
// 글자 비교: 보이지 않는 글자, 공백, 맥/윈도우의 한글 저장 방식 차이(NFC/NFD)를 무시하고 비교
const clean = x => String(x == null ? '' : x).normalize('NFC').replace(/[\s\u200B-\u200D\uFEFF]/g, '');
const same = (a, b) => clean(a) === clean(b);
const N = x => Number(String(x).replace(/[^\d.\-]/g, '')); // '1차시'처럼 글자가 섞여도 숫자만 읽기
const cache = () => CacheService.getScriptCache();
function rows(n) {
  const s = sh(n); if (!s || s.getLastRow() < 2) return [];
  const v = s.getDataRange().getValues(), h = v.shift();
  return v.filter(r => String(r[0]) !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}
function rowsC(n) { // 내용 시트는 30초간 캐시 (수정 후 최대 30초 뒤 반영)
  const c = cache(), k = 'r_' + n, h = c.get(k);
  if (h) return JSON.parse(h);
  const r = rows(n); try { c.put(k, JSON.stringify(r), 30); } catch (e) {}
  return r;
}
const setting = k => { const r = rowsC('설정').find(r => r['항목'] === k); return r ? r['값'] : ''; };
const hash = (id, pw) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'hs|' + id + '|' + pw));
/* ---- 계정 찾기 (연도별 계정) ----
   '연도' 열이 비어 있는 학생은 기준연도(설정 시트) 학생으로 봅니다. 관리자·test/spare 계정은 모든 연도 공통입니다.
   기준연도 학생의 내부 아이디는 학번 그대로, 다른 연도는 '2027_20101'처럼 연도가 붙어서 기록이 따로 저장됩니다. */
const BASE = () => String(setting('기준연도') || '2026').trim();
const effYearBy = (role, id, y) => ((role || '학생') === '관리자' || noPw(id) || isAfter(id)) ? '' : String(y || BASE()).trim();
const uidBy = (id, ey) => { id = String(id).trim(); return (!ey || ey === BASE()) ? id : ey + '_' + id; };
/* 로그인할 때만 쓰는 계정표 사본(20초 캐시): 많은 학생이 한꺼번에 로그인해도 시트를 매번 읽지 않아 훨씬 빨라요.
   비밀번호를 바꾸면 바로 비워지고, 시트를 직접 고친 경우는 최대 20초 뒤 반영돼요. */
function accValsC() {
  const c = cache(), h = c.get('acc_all');
  if (h) { try { return JSON.parse(h); } catch (e) {} }
  const v = sh('계정').getDataRange().getValues();
  try { c.put('acc_all', JSON.stringify(v), 20); } catch (e) {}
  return v;
}
function findAccBy(id, year, useCache) {
  const v = useCache ? accValsC() : sh('계정').getDataRange().getValues();
  const nrm = x => String(x).trim().replace(/^0+(?=.)/, ''); // 시트가 000412를 412로 바꿔 저장해도 찾을 수 있게
  let any = null;
  for (let i = 1; i < v.length; i++) {
    if (nrm(v[i][0]) !== nrm(id)) continue;
    const ey = effYearBy(v[i][3], v[i][0], v[i][6]);
    if (ey === String(year)) return { row: i + 1, v: v[i] };
    if (ey === '' && !any) any = { row: i + 1, v: v[i] };
  }
  return any;
}
function findAcc(uid) { const m = /^(\d{4})_(.+)$/.exec(String(uid)); return findAccBy(m ? m[2] : uid, m ? m[1] : BASE()); }
/* 로그인 표(토큰): 서명된 값이라 서버 캐시가 비워져도 유지됩니다 (12시간) */
// 테스트·예비 계정(test01, spare01 …)은 처음 로그인해도 비밀번호 변경 창을 띄우지 않음 (비밀번호는 계속 아이디와 같음)
const noPw = id => /^(test|spare)\d+$/i.test(String(id).trim());
const isAfter = id => /^after\d+$/i.test(String(id).trim()); // 방과후 계정(after01 …): 모든 연도 공통, 처음 비밀번호는 a1234(첫 로그인 때 변경)
const initPw = id => isAfter(id) ? 'a1234' : String(id).trim(); // 비밀번호 칸이 비어 있을 때(초기 상태)의 비밀번호
/* 방과후 과목: '과목' 시트의 '대상' 칸에 방과후 라고 적힌 과목은 방과후 계정(after..)만 입장합니다.
   방과후 계정은 방과후 과목만 보이고, 일반 학생에게는 방과후 과목이 보이지 않습니다. 관리자·test/spare 계정은 모두 볼 수 있습니다. */
const subOK = (u, row) => {
  if (u.role === '관리자' || /^(test|spare)\d+$/i.test(String(u.sid || u.id))) return true;
  const aft = isAfter(u.sid || u.id), tg = String(row['대상'] || '').trim() === '방과후';
  return aft === tg;
};
const subAllowed = (u, sub) => { const r = rowsC('과목').find(r => same(r['과목'], sub)); return !r || subOK(u, r); };
const nk = x => String(x).trim().replace(/^0+(?=.)/, '');
function secret() { const p = PropertiesService.getScriptProperties(); let s = p.getProperty('SECRET'); if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); } return s; }
const sign = x => Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(x, secret()));
function mkToken(id) { const p = Utilities.base64EncodeWebSafe(JSON.stringify({ i: id, e: Date.now() + 12 * 3600 * 1000 })); return p + '.' + sign(p); }
function session(t) {
  if (!t) return null;
  const [p, s] = String(t).split('.');
  if (!p || !s || sign(p) !== s) return null;
  let o; try { o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p)).getDataAsString()); } catch (e) { return null; }
  if (o.e < Date.now()) return null;
  const c = cache(), h = c.get('u_' + nk(o.i)); if (h) return JSON.parse(h);
  const a = findAcc(o.i); if (!a) return null;
  const u = { id: String(o.i), sid: String(a.v[0]).trim(), year: effYearBy(a.v[3], a.v[0], a.v[6]), name: a.v[1], cls: a.v[2], role: a.v[3] || '학생', must: !a.v[4] && !noPw(o.i) };
  c.put('u_' + nk(o.i), JSON.stringify(u), 600); return u;
}

/* ---------- 로그인 / 비밀번호 ---------- */
function login(d) {
  const id = String(d.id || '').trim(), yr = String(d.year || curYear()).trim(), a = findAccBy(id, yr, true), bad = { error: '학번 또는 비밀번호가 틀렸어요.' };
  if (!a) return bad;
  const h = a.v[4], init = !h;   // 비밀번호 칸이 비어 있으면 '초기 상태' = 비밀번호는 학번
  const ey = effYearBy(a.v[3], a.v[0], a.v[6]), uid = uidBy(id, ey);
  if (init ? String(d.pw) !== initPw(id) : hash(uid, d.pw) !== h) return bad;
  const u = { id: uid, sid: id, year: ey, name: a.v[1], cls: a.v[2], role: a.v[3] || '학생', must: init && !noPw(id) };
  cache().put('u_' + nk(uid), JSON.stringify(u), 600);
  return Object.assign({ token: mkToken(uid) }, u);
}
function changePw(u, d) {
  const pw = String(d.newPw || '');
  if (pw.length < 4) return { error: '비밀번호는 4자 이상이어야 해요.' };
  if (pw === u.id || pw === initPw(u.sid || u.id)) return { error: '처음 비밀번호와 다른 비밀번호를 써 주세요.' };
  const a = findAcc(u.id);
  sh('계정').getRange(a.row, 5).setValue(hash(u.id, pw)); cache().remove('acc_all');
  u.must = false; cache().put('u_' + nk(u.id), JSON.stringify(u), 600);
  return { ok: true };
}
function resetPw(id) {
  const a = findAcc(id); if (!a) return { error: '학생을 찾을 수 없어요.' };
  sh('계정').getRange(a.row, 5).clearContent(); cache().remove('u_' + nk(id)); cache().remove('acc_all');
  return { ok: true };
}
/* 학생이 '비밀번호를 잊어버렸어요'에서 학번 + 고유키로 직접 초기화 (관리자 계정은 불가) */
function forgot(d) {
  if (noPw(String(d.id || ''))) return { error: '이 계정은 비밀번호 초기화를 할 수 없어요. 선생님께 문의해 주세요.' }; // test/spare 계정은 차단
  const id = String(d.id || '').trim(), key = String(d.key || '').trim(), yr = String(d.year || curYear()).trim(), c = cache(), ck = 'f_' + id + '_' + yr, n = Number(c.get(ck) || 0);
  if (n >= 5) return { error: '시도 횟수를 넘었어요. 10분 뒤에 다시 해 보세요.' };
  const a = findAccBy(id, yr), k = a ? String(a.v[5]).trim() : '';
  if (!a || (a.v[3] || '학생') === '관리자' || !k || k.toLowerCase() !== key.toLowerCase()) {
    c.put(ck, String(n + 1), 600);
    return { error: '학번 또는 고유키가 맞지 않아요.' };
  }
  sh('계정').getRange(a.row, 5).clearContent(); cache().remove('acc_all');
  c.remove(ck); c.remove('u_' + nk(uidBy(id, effYearBy(a.v[3], a.v[0], a.v[6]))));
  return { ok: true };
}

/* ---------- 과목 / 차시 / 진도 ---------- */
function structure(sub) {
  const n = { 개념: {}, 해석: {}, 작성: {} };
  Object.keys(ACT).forEach(k => rowsC(ACT[k]).filter(r => same(r['과목'], sub)).forEach(r => { const c = N(r['차시']); n[k][c] = (n[k][c] || 0) + 1; }));
  return rowsC('차시').filter(l => same(l['과목'], sub)).map(l => {
    const no = N(l['차시']);
    return { no, title: l['제목'], t: { 개념: n.개념[no] ? 1 : 0, 해석: n.해석[no] || 0, 작성: n.작성[no] || 0 } };
  }).filter(l => l.t.개념 + l.t.해석 + l.t.작성 > 0).sort((a, b) => a.no - b.no);
}
const sliceProg = (m, sub) => { const o = {}, pre = sub + '|'; Object.keys(m).forEach(k => { if (k.indexOf(pre) === 0) o[k.slice(pre.length)] = m[k]; }); return o; };
function prog() { // 전체 학생 (관리자용)
  const m = {};
  rows('기록').forEach(r => {
    if (r['정답'] === true) { const id = String(r['학번']), k = r['과목'] + '|' + r['차시'] + '|' + r['활동']; (m[id] = m[id] || {})[k] = (m[id][k] || 0) + 1; }
  });
  return m;
}
function myRecs(id) { const s = sh('기록'), n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, 10).getValues().filter(r => String(r[0]) === String(id)); }
function progFor(id) { // 학생 한 명 (캐시)
  const c = cache(), k = 'p_' + id, h = c.get(k); if (h) return JSON.parse(h);
  const m = {}; myRecs(id).forEach(r => { if (r[6] === true) { const x = r[1] + '|' + r[2] + '|' + r[3]; m[x] = (m[x] || 0) + 1; } });
  c.put(k, JSON.stringify(m), 21600); return m;
}
function menu(u, d) {
  const L = structure(d.sub);
  if (!L.length) return { error: '아직 준비중입니다!' };
  const p = sliceProg(progFor(u.id), d.sub);
  return { lessons: L.map(l => Object.assign(l, { d: { 개념: p[l.no + '|개념'] || 0, 해석: p[l.no + '|해석'] || 0, 작성: p[l.no + '|작성'] || 0 } })) };
}
function lesson(u, d) {
  if (!structure(d.sub).length) return { error: '아직 준비중입니다!' };
  const no = Number(d.no), act = d.act;
  const R = rowsC(ACT[act]).filter(r => same(r['과목'], d.sub) && N(r['차시']) === no);
  const saved = {};
  myRecs(u.id).filter(r => same(r[1], d.sub) && Number(r[2]) === no && r[3] === act).forEach(r => saved[r[4]] = { answer: String(r[5]).replace(/^ (?=[=+\-@])/, ''), ok: r[6] === true, t: r[9] ? new Date(r[9]).getTime() : 0 });
  const drafts = {};
  myDrafts(u.id).filter(r => same(r[1], d.sub) && N(r[2]) === no && r[3] === act).forEach(r => drafts[r[4]] = { text: String(r[5]).replace(/^ (?=[=+\-@])/, ''), t: Number(r[6]) || 0 });
  let items;
  if (act === '개념') items = R.sort((a, b) => a['순서'] - b['순서']).map(r => ({ title: r['제목'], body: r['내용'] }));
  else if (act === '해석') items = R.sort((a, b) => a['문항'] - b['문항']).map(r => ({ q: r['문항'], text: r['문제'], code: r['코드'], type: String(r['유형'] || '').trim(), choices: r['유형'] === '객관식' ? String(r['보기']).split('|') : null }));
  else items = R.sort((a, b) => a['문항'] - b['문항']).map(r => ({ q: r['문항'], text: r['문제'], start: r['시작코드'], ins: String(r['테스트입력']).split(/\r?\n---\r?\n/), hint: String(r['힌트'] || ''), hasSample: String(r['예시코드'] || '').trim() !== '' }));
  return { items, saved, drafts };
}

function why(sub) { const c = n => sh(n) ? rowsC(n).filter(r => same(r['과목'], sub)).length : '(시트 이름을 못 찾음)'; return '차시 ' + c('차시') + '행 · 개념 ' + c('개념익히기') + ' · 해석 ' + c('해석해보기') + ' · 작성 ' + c('응용해보기'); }

/* ---------- 제출 저장 ---------- */
// 관리자·테스트·예비 계정은 어떤 활동 기록(제출, 임시저장, 집중 시간)도 서버에 저장하지 않음
const noLog = u => u.role === '관리자' || noPw(u.sid || u.id); // 관리자·테스트·예비 계정은 서버에 기록을 저장하지 않음 (방과후 계정은 저장됨: 컴퓨터실 PC를 껐다 켜도 남도록)
function submit(u, d) {
  const no = Number(d.no), act = d.act, q = Number(d.q || 0), sub = String(d.sub || '');
  let ok = true, ex = '', free = false;
  if (act === '해석') {
    const r = rowsC('해석해보기').find(r => same(r['과목'], sub) && N(r['차시']) === no && N(r['문항']) === q);
    if (!r) return { error: '문항을 찾을 수 없어요.' };
    free = FREE.includes(String(r['유형']).trim());
    ok = free ? !!String(d.answer).trim() : String(d.answer).trim().toLowerCase() === String(r['정답']).trim().toLowerCase();
    ex = r['해설'] || '';
  } else if (act === '작성') ok = !!d.ok;
  if (noLog(u)) return { correct: ok, explain: ex, free }; // 저장하지 않음
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sh('기록'), n = s.getLastRow(), now = new Date(), ans = String(d.answer || '').slice(0, 20000).replace(/^(?=[=+\-@])/, ' '); // 수식으로 해석되지 않게
    let row = 0;
    if (n > 1) { const j = s.getRange(2, 1, n - 1, 5).getValues().findIndex(r => String(r[0]) === u.id && same(r[1], sub) && Number(r[2]) === no && r[3] === act && Number(r[4]) === q); if (j >= 0) row = j + 2; }
    if (!row) s.appendRow([u.id, sub, no, act, q, ans, ok, 1, now, now]);
    else { const o = s.getRange(row, 6, 1, 5).getValues()[0]; s.getRange(row, 6, 1, 5).setValues([[ans, ok || o[1] === true, Number(o[2]) + 1, o[3], now]]); }
  } finally { lock.releaseLock(); }
  cache().remove('p_' + u.id);
  return { correct: ok, explain: ex, free };
}

/* ---------- 관리자 ---------- */
const students = () => rows('계정').filter(r => (r['역할'] || '학생') !== '관리자');
const effY = r => effYearBy(r['역할'], r['학번'], r['연도']);
const uidR = r => uidBy(r['학번'], effY(r));
function years() {
  const s = new Set(); rows('계정').forEach(r => { const y = effY(r); if (y) s.add(y); });
  const cur = curYear(); if (cur) s.add(cur); s.add(BASE());
  const L = [...s].sort().reverse(); return { years: L, cur: cur || L[0] };
}
function classes(sub, year) {
  year = String(year || curYear()).trim();
  const a = Number(setting('시작반')), b = Number(setting('끝반')), L = structure(sub), p = prog();
  const per = L.reduce((n, l) => n + l.t.개념 + l.t.해석 + l.t.작성, 0), out = [];
  for (let c = a; c <= b; c++) {
    const st = students().filter(r => Number(r['반']) === c && effY(r) === year);
    let done = 0;
    st.forEach(r => { const m = sliceProg(p[uidR(r)] || {}, sub); L.forEach(l => ['개념', '해석', '작성'].forEach(k => done += Math.min(l.t[k], m[l.no + '|' + k] || 0))); });
    out.push({ cls: c, count: st.length, pct: per && st.length ? Math.round(done / (per * st.length) * 100) : 0 });
  }
  return { classes: out };
}
function classTable(cls, sub, year) {
  year = String(year || curYear()).trim();
  const p = prog();
  return {
    lessons: structure(sub),
    students: students().filter(r => Number(r['반']) === Number(cls) && effY(r) === year).sort((x, y) => String(x['학번']).localeCompare(String(y['학번'])))
      .map(r => ({ id: uidR(r), sid: String(r['학번']), name: r['이름'], d: sliceProg(p[uidR(r)] || {}, sub) }))
  };
}
function detail(id) {
  const fs = new Set(rowsC('해석해보기').filter(x => FREE.includes(String(x['유형']).trim())).map(x => x['과목'] + '|' + x['차시'] + '|' + x['문항']));
  const a = findAcc(id), f = x => Utilities.formatDate(new Date(x), 'Asia/Seoul', 'MM-dd HH:mm');
  return {
    name: a ? a.v[1] : '', id, sid: a ? String(a.v[0]).trim() : id, times: timesOf(id),
    recs: rows('기록').filter(r => String(r['학번']) === String(id)).map(r => ({ sub: r['과목'], no: r['차시'], act: r['활동'], q: r['문항'], answer: r['답안'], ok: r['정답'] === true, free: r['활동'] === '해석' && fs.has(r['과목'] + '|' + r['차시'] + '|' + r['문항']), tries: r['시도'], first: f(r['최초제출']), last: f(r['최근제출']) }))
  };
}

/* ---------- 처음 1회: 시트 틀과 예시 만들기 ---------- */
const EX_CRIT = [ // 예시 평가기준 [수행평가명, 평가항목, 세분화항목, 점수들, 단계설명들(점수 순서대로, 비워도 됨)] (학기마다 시트에서 바꾸세요)
  ['디지털 데이터 표현하기', '진법 변환하기', '진법 변환 과정 제시하기', '6/5/4/3', []],
  ['디지털 데이터 표현하기', '진법 변환하기', '진법 변환 결과 제시하기', '4/3/2/1', []],
  ['디지털 데이터 표현하기', '비트 수 계산하기', '', '8/6/4/2', ["실생활의 데이터를 디지털 데이터로 표현하기 위해 필요한 비트 수를 논리적인 과정을 들어 정확히 제시함.", "실생활의 데이터를 디지털 데이터로 표현하기 위해 필요한 비트 수를 논리적인 과정을 들어 제시했으나, 비트 수를 계산하지 못함.", "실생활의 데이터를 디지털 데이터로 표현하기 위해 필요한 비트 수를 계산하였으나, 논리적인 과정을 제시하지 않음.", "실생활의 데이터를 디지털 데이터로 표현하기 위해 필요한 비트 수를 계산하지 못하고 논리적인 과정도 제시하지 않음."]],
  ['디지털 데이터 표현하기', '디지털 코드 변환하기', '', '6/4/2', []],
  ['디지털 데이터 표현하기', '문자 데이터 부호화하기', '', '6/4/2', []],
  ['프로그램 분석하기', '변숫값 분석하기', '', '7/5/3/1', []],
  ['프로그램 분석하기', '변숫값 추적하기', '', '8/7/6/5', []],
  ['프로그램 분석하기', '조건문 분석하기', '조건문 제시하기', '5/3/1', []],
  ['프로그램 분석하기', '조건문 분석하기', '조건 판단하기', '5/3/1', []],
  ['프로그램 분석하기', '리스트 분석하기', '', '5/3/1', []],
  ['프로그램 분석하기', '프로그램 출력 결과 제시하기', '', '5/4/3/2', []],
  ['데이터 분석하기', '데이터 특징 파악하기', '', '', []],
  ['데이터 분석하기', '데이터 구조화하기', '', '', []],
  ['데이터 분석하기', '데이터 의미 해석하기', '', '', []],
  ['데이터 분석하기', '데이터 활용 방안 제시하기', '', '', []]
];
const critRows = (yr, sm) => { // EX_CRIT → 시트 행 (한 행 = 한 단계. 첫 줄에만 이름·연도, 아래 단계 줄은 비워 둠)
  const out = [];
  EX_CRIT.forEach(c => {
    const L = String(c[3]).split('/').filter(x => x !== ''); if (!L.length) L.push('');
    L.forEach((sc, i) => out.push(i ? ['', '', '', '', '', Number(sc), (c[4] || [])[i] || ''] : [yr, sm, c[0], c[1], c[2], sc === '' ? '' : Number(sc), (c[4] || [])[0] || '']));
  });
  return out;
};
function 초기설정() {
  const ss = SpreadsheetApp.getActive();
  const T = {
    '설정': [['항목', '값'], ['학교명', '배곧해솔중학교'], ['시작반', 1], ['끝반', 7], ['기준연도', 2026], ['현재연도', ''], ['현재학기', ''], ['성적공개', '예']],
    '과목': [['과목', '아이콘', '개념이름', '해석이름', '작성이름', '대상'], ['파이썬', '🐍', '', '', ''], ['엔트리', '🧩', '', '', ''], ['데이터 분석', '📊', '', '데이터 해석하기', ''], ['파이썬 방과후', '🎓', '', '', '', '방과후']],
    '계정': [['학번', '이름', '반', '역할', '비밀번호(자동)', '고유키', '연도'], ['190028', '교사1', '', '관리자', '', '', ''], ['000412', '교사2', '', '관리자', '', '', ''], ['20101', '홍길동', 1, '학생', '', '', 2026]],
    '차시': [['과목', '차시', '제목'], ['파이썬', 1, '변수와 출력']],
    '개념익히기': [['과목', '차시', '순서', '제목', '내용'], ['파이썬', 1, 1, 'print() 함수', 'print()는 화면에 값을 보여 줍니다.\n```\nprint("안녕")\n```']],
    '해석해보기': [['과목', '차시', '문항', '문제', '코드', '유형', '보기', '정답', '해설'],
      ['파이썬', 1, 1, '실행 결과는?', 'a = 3\nb = 4\nprint(a + b)', '객관식', '7|34|a+b', '7', '3과 4를 더한 값이 출력됩니다.'],
      ['파이썬', 1, 2, '각 줄의 의미를 빈칸에 써 보세요.', 'for i in range(1, 11):  /* ? */\n    print(i)  /* ? */', '설명쓰기', '', '', '']],
    '응용해보기': [['과목', '차시', '문항', '문제', '시작코드', '테스트입력', '기대출력', '힌트', '예시코드', '예시비밀번호'], ['파이썬', 1, 1, '이름을 입력받아 "안녕, 이름!"을 출력하세요.', 'name = input()\n', '민수\n---\n지아', '안녕, 민수!\n---\n안녕, 지아!', 'input() 으로 입력을 받고, print() 로 출력해 보세요.\n---\n```\nname = input()\nprint("안녕, " + name + "!")\n```', 'name = input()\nprint("안녕, " + name + "!")', 'abc123']],
    '임시저장': [['학번', '과목', '차시', '활동', '문항', '내용', '저장시각']],
    '집중시간': [['키', '학번', '날짜', '집중(초)', '이탈(초)', '갱신']],
    '공지': [['날짜', '제목', '내용'], ['2026-10-23', '수업 안내', '오늘은 **변수**를 배워요.']],
    '용어사전': [['용어', '뜻', '예시', '과목'], ['변수', '값을 담아 두는 이름표가 붙은 상자', '`x = 3`', '파이썬']],
    '오류모음': [['오류', '뜻', '해결방법', '과목'], ['NameError', '정의하지 않은 이름을 사용했어요', '철자를 확인하고, 사용하기 전에 값을 넣었는지 확인해요.', '파이썬']],
    '치트시트': [['분류', '제목', '내용'], ['출력', 'print 사용법', '```\nprint("안녕")\n```']],
    '수행평가안내': [['분류', '제목', '내용'], ['1학기', '파이썬 코딩 수행평가', '**평가 기간** 11월 둘째 주\n**배점** 30점\n**평가 기준** 코드 작성(20), 발표(10)']],
    '평가기준': [['연도', '학기', '수행평가명', '평가항목', '세분화항목', '점수', '단계설명']].concat(critRows(2026, 1)),
    '등급기준': [['등급', '최소점수', '연도', '학기'], ['A', 90, 2026], ['B', 80, 2026], ['C', 70, 2026], ['D', 60, 2026], ['E', 0, 2026]],
    '수행평가성적': [['연도', '학번', '이름'].concat(EX_CRIT.map(c => c[2] || c[1]), ['비고']), [2026, '20101', '홍길동'].concat(EX_CRIT.map(c => { const l = String(c[3]).split('/').filter(x => x !== ''); return l.length ? Number(l[Math.min(1, l.length - 1)]) : ''; }), ['예시 행이에요. 지우고 쓰세요'])],
    '기록': [['학번', '과목', '차시', '활동', '문항', '답안', '정답', '시도', '최초제출', '최근제출']]
  };
  Object.keys(T).forEach(n => {
    let s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (n === '계정') s.getRange('A:A').setNumberFormat('@'), s.getRange('F:F').setNumberFormat('@'); // 학번 앞자리 0, 고유키 보존
    if (n === '기록' || n === '임시저장') s.getRange('A:A').setNumberFormat('@');
    if (n === '임시저장') s.getRange('F:F').setNumberFormat('@');
    if (n === '집중시간') s.getRange('A:C').setNumberFormat('@');
    if (n === '수행평가성적') s.getRange('B:B').setNumberFormat('@');
    if (s.getLastRow() === 0) { s.getRange(1, 1, T[n].length, T[n][0].length).setValues(T[n]); s.setFrozenRows(1); s.getRange(1, 1, 1, T[n][0].length).setFontWeight('bold').setBackground('#e8e0f7'); }
  });
}

/* 고유키가 비어 있는 학생 전원에게 6자리 키를 자동 부여 (직접 정해서 입력하셔도 됩니다) */
function 고유키일괄생성() {
  const s = sh('계정'), v = s.getDataRange().getValues(), ch = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  s.getRange('F:F').setNumberFormat('@');
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === '' || (v[i][3] || '학생') === '관리자' || String(v[i][5]).trim() !== '') continue;
    let k = ''; for (let j = 0; j < 6; j++) k += ch[Math.floor(Math.random() * ch.length)];
    s.getRange(i + 1, 6).setValue(k);
  }
}

/* 관리자 비밀번호를 원하는 값으로 직접 지정 (필요할 때만 실행) */
function 관리자비밀번호설정() {
  const ID = '1900028', PW = '000412';
  const a = findAcc(ID);
  if (!a) throw new Error('계정 시트에 ' + ID + ' 행이 없어요.');
  sh('계정').getRange(a.row, 5).setValue(hash(ID, PW)); cache().remove('acc_all');
}

/* ---------- 문제 확인용 (필요할 때만 실행) ---------- */
// 시트를 고친 직후 바로 반영하고 싶을 때
function 캐시비우기() {
  cache().removeAll(['r_설정', 'r_과목', 'r_차시', 'r_개념익히기', 'r_해석해보기', 'r_응용해보기', 'r_공지', 'r_용어사전', 'r_오류모음', 'r_치트시트', 'r_수행평가안내', 'acc_all', 'r_평가기준', 'r_등급기준']);
}
// 과목이 '준비중'으로 나오는 이유를 찾아 줍니다. 실행 후 아래쪽 '실행 로그'를 보세요.
function 점검() {
  캐시비우기();
  const subs = rows('과목').map(r => String(r['과목']).trim());
  Logger.log('과목 시트(글자 수): ' + JSON.stringify(subs.map(x => [x, x.length])));
  ['차시', '개념익히기', '해석해보기', '응용해보기'].forEach(n => {
    const s = sh(n);
    if (!s) { Logger.log('❌ 시트가 없어요: ' + n); return; }
    const h = s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0].map(x => String(x).trim());
    const R = rows(n);
    Logger.log('[' + n + '] 1행 제목=' + JSON.stringify(h) + ' / 데이터 행=' + R.length + (h.indexOf('과목') < 0 ? '  ❌ "과목" 제목 열이 없어요' : ''));
    subs.forEach(x => Logger.log('    ' + x + ' → ' + R.filter(r => same(r['과목'], x)).length + '행'));
    Logger.log('    이 시트의 과목 값(글자 수): ' + JSON.stringify([...new Set(R.map(r => String(r['과목'])))].map(x => [x, x.length])));
    const other = [...new Set(R.map(r => String(r['과목']).trim()))].filter(x => !subs.some(s => same(s, x)));
    if (other.length) Logger.log('    ⚠ 과목 시트에 없는 이름: ' + JSON.stringify(other));
    if (n === '응용해보기') Logger.log('    예시코드 열: ' + (h.indexOf('예시코드') < 0 ? '❌ 없음(열 제목을 확인하세요)' : '있음, 예시코드가 적힌 문제 ' + R.filter(r => String(r['예시코드'] || '').trim()).length + '개') + ' / 예시비밀번호 열: ' + (h.indexOf('예시비밀번호') < 0 ? '❌ 없음' : '있음') + ' / 힌트 열: ' + (h.indexOf('힌트') < 0 ? '없음' : '있음'));
    if (n === '해석해보기' || n === '응용해보기') {
      const seen = {}; R.forEach(r => { const k = r['과목'] + '|' + r['차시'] + '|' + r['문항']; seen[k] = (seen[k] || 0) + 1; });
      const bad = Object.keys(seen).filter(k => seen[k] > 1 || k.slice(-1) === '|');
      if (bad.length) Logger.log('    ⚠ 문항 번호가 비었거나 겹쳐요 (과목|차시|문항): ' + JSON.stringify(bad));
    }
  });
  Logger.log('준비 상태: ' + JSON.stringify(subs.map(x => [x, structure(x).length > 0])));
}

/* ---------- 입력 중인 답 임시저장 / 목차 ---------- */
function draftSheet() {
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName('임시저장');
  if (!s) {
    s = ss.insertSheet('임시저장'); s.getRange('A:A').setNumberFormat('@'); s.getRange('F:F').setNumberFormat('@');
    s.getRange(1, 1, 1, 7).setValues([['학번', '과목', '차시', '활동', '문항', '내용', '저장시각']]).setFontWeight('bold').setBackground('#e8e0f7'); s.setFrozenRows(1);
  }
  return s;
}
function myDrafts(id) { const s = sh('임시저장'); return (!s || s.getLastRow() < 2) ? [] : s.getRange(2, 1, s.getLastRow() - 1, 7).getValues().filter(r => String(r[0]) === String(id)); }
function saveDraft(u, d) {
  if (noLog(u)) return { ok: true }; // 저장하지 않음
  const sub = String(d.sub || ''), no = N(d.no), act = String(d.act || ''), q = Number(d.q || 0);
  if (!['개념', '해석', '작성'].includes(act)) return { error: '저장할 수 없는 활동이에요.' };
  const txt = String(d.text || '').slice(0, 20000).replace(/^(?=[=+\-@])/, ' ');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = draftSheet(), n = s.getLastRow(); let row = 0;
    if (n > 1) { const j = s.getRange(2, 1, n - 1, 5).getValues().findIndex(r => String(r[0]) === u.id && same(r[1], sub) && N(r[2]) === no && r[3] === act && Number(r[4]) === q); if (j >= 0) row = j + 2; }
    const t = Date.now();
    if (!row) s.appendRow([u.id, sub, no, act, q, txt, t]); else s.getRange(row, 6, 1, 2).setValues([[txt, t]]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}
/* 응용해보기 예시 코드: 비밀번호가 맞을 때만 서버가 코드를 보내 줍니다(화면에는 처음부터 들어 있지 않음). 관리자는 비밀번호 없이 열람 */
function sampleCode(u, d) {
  const sub = String(d.sub || ''), no = Number(d.no), q = Number(d.q || 0);
  const r = rowsC('응용해보기').find(r => same(r['과목'], sub) && N(r['차시']) === no && N(r['문항']) === q);
  if (!r || String(r['예시코드'] || '').trim() === '') return { error: '예시 코드가 없어요.' };
  if (u.role === '관리자') return { code: String(r['예시코드']) };
  const pw = String(r['예시비밀번호'] || '').trim().toLowerCase();
  if (!pw) return { error: '아직 비밀번호가 정해지지 않았어요. 선생님께 알려 주세요.' };
  const c = cache(), ck = 'sc_' + nk(u.id) + '_' + sub + '_' + no + '_' + q, n = Number(c.get(ck) || 0);
  if (n >= 3) return { error: 'locked', locked: true };
  if (String(d.pw || '').trim().toLowerCase() !== pw) { c.put(ck, String(n + 1), 1800); return n + 1 >= 3 ? { error: 'locked', locked: true } : { error: '비밀번호가 맞지 않아요. (' + (n + 1) + '/3)' }; } // 3번 틀리면 30분 동안 잠김
  c.remove(ck);
  return { code: String(r['예시코드']) };
}
function toc(u) { // 왼쪽 목차: 과목 > 차시 (진도 포함)
  const p = progFor(u.id);
  return { subjects: rowsC('과목').filter(r => subOK(u, r)).map(r => {
    const name = String(r['과목']).replace(/[\u200B-\u200D\uFEFF]/g, '').trim(), L = structure(r['과목']), m = sliceProg(p, name);
    const o = { name, icon: r['아이콘'] || '📘', ready: L.length > 0, names: { 개념: r['개념이름'] || '', 해석: r['해석이름'] || '', 작성: r['작성이름'] || '' }, lessons: L.map(l => Object.assign(l, { d: { 개념: m[l.no + '|개념'] || 0, 해석: m[l.no + '|해석'] || 0, 작성: m[l.no + '|작성'] || 0 } })) };
    if (u.role === '관리자') o.why = why(r['과목']);
    return o;
  }) };
}

/* 테스트(test01~05)·예비(spare01~05) 계정을 만들고 비밀번호를 a1234로 지정합니다. 한 번만 실행하세요 (다시 실행하면 같은 값으로 덮어씁니다).
   이 계정들은 모든 연도 공통이고, 로그인해도 비밀번호 변경 창이 뜨지 않습니다. 반 칸이 비어 있어서 학급 현황에는 나오지 않습니다. */
function 특수계정만들기() {
  const s = sh('계정'); s.getRange('A:A').setNumberFormat('@');
  const L = [];
  for (let i = 1; i <= 5; i++) { L.push(['test0' + i, '테스트' + i]); L.push(['spare0' + i, '예비' + i]); }
  L.forEach(([id, name]) => {
    const a = findAccBy(id, ''), row = [id, name, '', '학생', hash(id, 'a1234'), '', ''];
    if (a) s.getRange(a.row, 1, 1, 7).setValues([row]); else s.appendRow(row);
  });
}

/* 방과후 계정(after01~after20)을 만듭니다. 처음 비밀번호는 a1234이고, 처음 로그인하면 바꾸도록 안내됩니다.
   고유키(비밀번호 찾기용)도 함께 만들어 계정 시트 F열에 적어 줍니다. 이미 있는 계정은 이름·비밀번호·고유키를 건드리지 않으므로 여러 번 실행해도 안전합니다.
   (예전 버전으로 1234/a1234 비밀번호가 직접 저장된 계정은 초기 상태로 되돌려 줍니다.) 더 필요하면 아래 20을 바꾸세요. */
function 방과후계정만들기() {
  const COUNT = 20, s = sh('계정'), ch = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; s.getRange('A:A').setNumberFormat('@'); s.getRange('F:F').setNumberFormat('@');
  const key = () => { let k = ''; for (let j = 0; j < 6; j++) k += ch[Math.floor(Math.random() * ch.length)]; return k; };
  let made = 0, fixed = 0;
  for (let i = 1; i <= COUNT; i++) {
    const id = 'after' + (i < 10 ? '0' : '') + i, a = findAccBy(id, '');
    if (!a) { s.appendRow([id, '방과후' + i, '', '학생', '', key(), '']); made++; continue; }
    if (a.v[4] && (a.v[4] === hash(id, '1234') || a.v[4] === hash(id, 'a1234'))) { s.getRange(a.row, 5).clearContent(); fixed++; }
    if (!String(a.v[5]).trim()) s.getRange(a.row, 6).setValue(key());
  }
  캐시비우기();
  Logger.log('새로 만든 방과후 계정 ' + made + '개, 초기 상태로 되돌린 계정 ' + fixed + '개 (처음 비밀번호 a1234)');
}

const 알림 = m => { Logger.log(m); try { SpreadsheetApp.getActive().toast(m, '반 계정 만들기', 15); } catch (e) {} }; // 실행 로그 + 시트 오른쪽 아래 알림
/* 담당 반 학생 계정을 한꺼번에 만듭니다. 설정 시트의 '시작반' ~ '끝반', '학년'을 읽습니다('학년' 줄이 없으면 2로 보고 줄을 만들어 줍니다).
   반마다 30명(번호 01~30), 학번 = 학년 + 반 두 자리 + 번호 두 자리(예: 2학년 1반 1번 = 20101)로 만듭니다.
   연도는 2월부터 새 학년도로 적습니다(예: 2027년 2월에 실행 → 2027). 다른 기능은 3월부터 새 학년도로 바뀝니다.
   이름은 비워 두니 직접 채우세요. 비밀번호 칸은 비워 둬서 처음 비밀번호 = 학번(첫 로그인 때 변경), 고유키는 자동으로 만듭니다.
   같은 연도에 이미 있는 학번은 건드리지 않으므로 여러 번 실행해도 안전합니다. 실제로 없는 번호의 줄은 지워 주세요. */
function 반계정만들기() {
  const PER = 30, s = sh('계정'), ch = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', o = String(setting('현재연도') || '').trim(), t = kstNow();
  const yr = /^\d{4}$/.test(o) ? o : String(t.m >= 2 ? t.y : t.y - 1); // 수업 준비를 2월부터 하므로 2월이면 새 학년도 계정으로 만듦 (설정 '현재연도'가 있으면 그 값)
  const key = () => { let k = ''; for (let j = 0; j < 6; j++) k += ch[Math.floor(Math.random() * ch.length)]; return k; };
  const two = n => (n < 10 ? '0' : '') + n;
  캐시비우기();
  const a = Number(setting('시작반')), b = Number(setting('끝반'));
  if (!(a >= 1 && b >= a && b <= 99)) { 알림('❌ 설정 시트의 시작반·끝반을 숫자로 적어 주세요 (예: 1, 7)'); return; }
  let gr = String(setting('학년')).trim();
  if (!gr) { gr = '2'; sh('설정').appendRow(['학년', 2]); 캐시비우기(); }
  if (!/^[1-9]$/.test(gr)) { 알림('❌ 설정 시트의 학년을 숫자 하나로 적어 주세요 (예: 2)'); return; }
  const cls = []; for (let c = a; c <= b; c++) cls.push(c);
  const v = s.getDataRange().getValues(), have = new Set();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() && effYearBy(v[i][3], v[i][0], v[i][6]) === yr) have.add(nk(v[i][0]));
  const add = [];
  cls.forEach(c => { for (let n = 1; n <= PER; n++) { const id = gr + two(c) + two(n); if (!have.has(nk(id))) add.push([id, '', c, '학생', '', key(), Number(yr)]); } });
  if (!add.length) { 알림('추가할 계정이 없어요. ' + yr + '년 계정이 이미 모두 있어요.'); return; }
  let last = v.length; while (last > 1 && String(v[last - 1][0]).trim() === '') last--; // A열(학번)이 마지막으로 채워진 줄
  s.getRange('A:A').setNumberFormat('@'); s.getRange('F:F').setNumberFormat('@');
  s.getRange(last + 1, 1, add.length, 7).setValues(add);
  캐시비우기();
  알림(yr + '년 ' + gr + '학년 ' + a + '~' + b + '반: 새 계정 ' + add.length + '개를 만들었어요 (이미 있던 ' + (cls.length * PER - add.length) + '개는 건너뜀). 이름은 직접 채워 주세요.');
}

/* ---------- 수업 집중 시간 / 학습 도우미 ---------- */
function timerSheet() {
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName('집중시간');
  if (!s) { s = ss.insertSheet('집중시간'); s.getRange('A:C').setNumberFormat('@'); s.getRange(1, 1, 1, 6).setValues([['키', '학번', '날짜', '집중(초)', '이탈(초)', '갱신']]).setFontWeight('bold').setBackground('#e8e0f7'); s.setFrozenRows(1); }
  return s;
}
function saveTimer(u, d) {
  if (noLog(u)) return { ok: true };
  const date = String(d.date || '').slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { error: '날짜가 올바르지 않아요.' };
  const f = Math.max(0, Math.min(86400, Math.floor(Number(d.focus) || 0))), a = Math.max(0, Math.min(86400, Math.floor(Number(d.away) || 0))), st = Math.floor(Number(d.start) || 0), key = u.id + '|' + date + (st ? '|' + st : '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = timerSheet(), hit = s.getRange('A:A').createTextFinder(key).matchEntireCell(true).findNext();
    if (hit) { const r = hit.getRow(), o = s.getRange(r, 4, 1, 2).getValues()[0]; s.getRange(r, 4, 1, 3).setValues([[Math.max(Number(o[0]) || 0, f), Math.max(Number(o[1]) || 0, a), new Date()]]); }
    else s.appendRow([key, u.id, date, f, a, new Date(), st]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}
function getTimer(u, d) {
  const date = String(d.date || '').slice(0, 10), s = sh('집중시간');
  if (!s || s.getLastRow() < 2) return { focus: 0, away: 0 };
  const hit = s.getRange('A:A').createTextFinder(u.id + '|' + date).matchEntireCell(true).findNext();
  if (!hit) return { focus: 0, away: 0 };
  const o = s.getRange(hit.getRow(), 4, 1, 2).getValues()[0]; return { focus: Number(o[0]) || 0, away: Number(o[1]) || 0 };
}
const fd = x => { if (x === '' || x == null) return ''; const dt = x instanceof Date ? x : (/T\d\d:/.test(String(x)) ? new Date(x) : null); return dt ? Utilities.formatDate(dt, 'Asia/Seoul', 'yyyy-MM-dd') : String(x).trim(); };
function timesOf(id) { // 접속 구간(50분)마다 한 줄로 저장되므로 날짜별로 합쳐서 보여 줌
  const s = sh('집중시간'); if (!s || s.getLastRow() < 2) return [];
  const m = {};
  s.getRange(2, 2, s.getLastRow() - 1, 4).getValues().filter(r => String(r[0]) === String(id)).forEach(r => {
    const d = fd(r[1]), o = m[d] || (m[d] = { date: d, focus: 0, away: 0 }); o.focus += Number(r[2]) || 0; o.away += Number(r[3]) || 0;
  });
  return Object.keys(m).map(k => m[k]).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 14);
}
function info() {
  return {
    notices: rowsC('공지').map(r => ({ date: fd(r['날짜']), title: r['제목'] || '', body: r['내용'] || '' })).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 30),
    terms: rowsC('용어사전').map(r => ({ name: String(r['용어']), body: r['뜻'] || '', ex: r['예시'] || '', sub: r['과목'] || '' })),
    errors: rowsC('오류모음').map(r => ({ name: String(r['오류']), body: r['뜻'] || '', fix: r['해결방법'] || '', sub: r['과목'] || '' })),
    cheats: rowsC('치트시트').map(r => ({ cat: r['분류'] || '기타', title: r['제목'] || '', body: r['내용'] || '' })),
    assess: rowsC('수행평가안내').map(r => ({ cat: r['분류'] || '기타', title: r['제목'] || '', body: r['내용'] || '' })),
    cyear: curYear(), csem: curSem(),
    criteria: critList(curYear()).filter(c => c.sem <= curSem()).map(c => ({ sem: c.sem, area: c.area, elem: c.elem, item: c.item, levels: c.levels, lv: c.lv }))
  };
}

/* ---------- 수행평가: 평가기준 / 성적 ----------
   [평가기준] 영역 · 평가요소 · 항목 · 단계점수(6/5/4/3) · 기준설명 · 연도  (영역·평가요소·연도는 위 칸과 같으면 비워도 됨)
   [수행평가성적] 연도 · 학번 · 이름 · (항목 이름과 똑같은 제목의 열들) · 비고  → 칸에는 받은 점수를 적음
   [등급기준] 등급 · 최소점수 · 연도 */
/* 학년도·학기 자동 계산 (한국 시간): 3월부터 새 학년도. 3~7월=1학기, 8월~다음해 2월=2학기
   설정 시트에 '현재연도' / '현재학기' 칸에 숫자가 적혀 있으면 그 값을 우선합니다(비우면 자동). */
const kstNow = () => { const t = new Date(Date.now() + 9 * 3600 * 1000); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1 }; };
const curYear = () => { const o = String(setting('현재연도') || '').trim(); if (/^\d{4}$/.test(o)) return o; const n = kstNow(); return String(n.m >= 3 ? n.y : n.y - 1); };
const curSem = () => { const o = String(setting('현재학기') || '').trim(); if (/^[12]$/.test(o)) return Number(o); const m = kstNow().m; return (m >= 3 && m <= 7) ? 1 : 2; };
function critAll() { // 평가기준 시트: 한 행 = 한 단계(점수 + 설명). 같은 평가항목/세분화항목의 연속된 행을 한 항목으로 묶음
  const c = cache(), h = c.get('r_평가기준'); if (h) return JSON.parse(h);
  const s = sh('평가기준'), out = [];
  if (s && s.getLastRow() > 1) {
    const v = s.getDataRange().getValues(), H = v.shift().map(x => String(x).trim()), at = (r, k) => H.indexOf(k) < 0 ? '' : r[H.indexOf(k)];
    let ar = '', el = '', sb = '', yr = '', sm = 1, cur = null;
    v.forEach(r => {
      const a = String(at(r, '수행평가명')).trim(), e = String(at(r, '평가항목')).trim(), b = String(at(r, '세분화항목')).trim(), y = String(at(r, '연도')).trim();
      if (y) yr = y; const smm = String(at(r, '학기')).match(/[12]/); if (smm) sm = Number(smm[0]);
      if (a && a !== ar) { ar = a; if (!e) { el = ''; sb = ''; } }
      if (e) { el = e; sb = b; } else if (b) sb = b;
      if (!el && !sb) return;
      const m = String(at(r, '점수') == null ? '' : at(r, '점수')).match(/-?\d+(\.\d+)?/), sc = m ? Number(m[0]) : null, d = String(at(r, '단계설명') || '').trim();
      const key = (yr || '') + '|' + sm + '|' + ar + '|' + el + '|' + sb;
      if (!cur || cur.key !== key) { cur = { key: key, year: yr || curYear(), sem: sm, area: ar || '기타', elem: el, item: sb, name: sb || el, levels: [], lv: [], full: 0 }; out.push(cur); }
      if (sc != null) { cur.levels.push(sc); cur.lv.push({ s: sc, d: d }); cur.full = Math.max(cur.full, sc); }
    });
    out.forEach(o => { delete o.key; });
  }
  try { c.put('r_평가기준', JSON.stringify(out), 30); } catch (e) {}
  return out;
}
const critList = year => critAll().filter(c => c.year === String(year));
function gradeData(u) { // 한 학생의 성적(학기별 영역 점수·합계·등급). 이번 학년도만, 현재 학기까지만 보여 줌
  const yr = String(u.year || BASE()).trim(), sid = nk(u.sid || u.id), out = { name: u.name, year: yr, sems: [] };
  if (yr !== curYear()) { out.old = true; return out; }
  const all = critList(yr).filter(c => c.levels.length && c.sem <= curSem()), s = sh('수행평가성적');
  if (!all.length || !s || s.getLastRow() < 2) return out;
  const v = s.getDataRange().getValues(), H = v.shift().map(x => String(x).trim()), iy = H.indexOf('연도'), is = H.indexOf('학번'), im = H.indexOf('학기'), inote = H.indexOf('비고');
  if (is < 0) return out;
  const mine = v.filter(r => nk(r[is]) === sid && (String(iy < 0 ? '' : r[iy]).trim() || String(BASE())) === yr), R = x => Math.round(x * 100) / 100;
  const G0 = rows('등급기준').filter(x => String(x['연도'] || BASE()).trim() === yr && x['최소점수'] !== '');
  Array.from(new Set(all.map(c => c.sem))).sort().forEach(sm => {
    const C = all.filter(c => c.sem === sm), cell = r => String(im < 0 ? '' : r[im]).trim();
    const r = mine.find(r => cell(r) === String(sm)) || mine.find(r => cell(r) === ''); if (!r) return;
    const A = {}; let any = false, total = 0, max = 0;
    C.forEach(c => {
      const k = H.indexOf(c.name), raw = k < 0 ? '' : r[k], sc = (raw === '' || raw == null || isNaN(Number(raw))) ? null : Number(raw);
      if (sc != null) { any = true; total += sc; }
      max += c.full;
      const a = A[c.area] || (A[c.area] = { area: c.area, rows: [], sum: 0, max: 0 });
      a.rows.push({ elem: c.elem, item: c.item, score: sc, full: c.full }); a.sum += sc || 0; a.max += c.full;
    });
    if (!any) return;
    const G = G0.filter(x => { const g = String(x['학기'] || '').trim(); return !g || g === String(sm); }).sort((a, b) => Number(b['최소점수']) - Number(a['최소점수']));
    const hit = G.find(x => total >= Number(x['최소점수']));
    out.sems.push({ sem: sm, areas: Object.keys(A).map(k => Object.assign(A[k], { sum: R(A[k].sum) })), total: R(total), max: max, grade: hit ? String(hit['등급']) : '', note: inote < 0 ? '' : String(r[inote] || '') });
  });
  return out;
}
/* 내 성적 보기: 비밀번호를 한 번 더 확인해서 본인 성적만 보여 줌 */
function myGrades(u, d) {
  if (u.role === '관리자') return { error: '학생 계정으로 로그인했을 때만 볼 수 있어요.' };
  const c = cache(), ck = 'g_' + nk(u.id), n = Number(c.get(ck) || 0);
  if (n >= 5) return { error: '비밀번호를 여러 번 틀렸어요. 10분 뒤에 다시 해 보세요.' };
  const a = findAcc(u.id), h = a ? a.v[4] : '', pw = String(d.pw || '');
  const ok = a && (h ? hash(u.id, pw) === h : pw === initPw(u.sid));
  if (!ok) { c.put(ck, String(n + 1), 600); return { error: '비밀번호가 맞지 않아요.' }; }
  c.remove(ck);
  if (/^(아니오|n|no|false|닫힘)$/i.test(String(setting('성적공개') || '').trim())) return { error: '아직 성적이 공개되지 않았어요.' };
  const g = gradeData(u); if (g.old) return { error: '지난 학년도 성적은 볼 수 없어요.' };
  return g;
}
/* 평가기준에 새 항목을 넣은 뒤 한 번 실행하면, 수행평가성적 시트에 없는 항목 열을 '비고' 앞에 자동으로 만들어 줍니다 */
function 성적열추가() {
  캐시비우기();
  const s = sh('수행평가성적'); if (!s) { Logger.log('❌ 수행평가성적 시트가 없어요'); return; }
  const H = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(x => String(x).trim()), add = [];
  critList(curYear()).filter(c => c.levels.length).forEach(c => { if (H.indexOf(c.name) < 0 && add.indexOf(c.name) < 0) add.push(c.name); });
  if (!add.length) { Logger.log('추가할 열이 없어요. 이미 모두 있어요.'); return; }
  let at = H.indexOf('비고'); at = at < 0 ? H.length + 1 : at + 1; // 비고 열 앞(1부터 센 번호)
  s.insertColumnsBefore(at, add.length); s.getRange(1, at, 1, add.length).setValues([add]);
  Logger.log('추가한 열: ' + JSON.stringify(add));
}
/* ---------- 반별 채점표('1반 채점표' ~ '7반 채점표') ----------
   채점표만들기 : 계정 시트(반·학번·이름)와 평가기준 시트(현재 학기 항목·단계 점수)로 반마다 채점표 탭을 만들고,
                  수행평가성적 시트에 이미 있는 점수를 채워 넣습니다. 이미 있는 탭은 건드리지 않습니다.
                  평가기준이 바뀌어 다시 만들고 싶으면 먼저 '채점표반영'을 실행한 뒤 그 탭을 지우고 다시 실행하세요.
   onEdit       : 채점표에서 점수(또는 비고)를 고치면 수행평가성적 시트의 같은 칸이 바로 바뀝니다.
   채점표반영   : 모든 채점표를 수행평가성적 시트에 한꺼번에 옮깁니다. 채점표가 기준이라 채점표의 빈 칸은 빈 칸으로 바뀝니다.
   채점표 모양 : 2행=영역, 3행=항목 이름(수행평가성적 열 제목과 같음), 4행=배점, 5행부터 학생(A열 학번, B열 이름),
                  영역마다 소계 오른쪽에 '기본점수'·'만점' 체크칸, 학생 바로 아래 '반 평균' 줄(영역 소계·총점 평균). */
const isNum = x => x !== '' && x != null && typeof x !== 'boolean' && !isNaN(Number(x));
const SC_TAB = /^\d+\s*반\s*채점표$/;
const SC_TOP = 5; // 학생이 시작하는 줄
const AVG = '반 평균'; // 학생 아래 평균 줄의 A열 글자
const 학생수 = t => { const v = t.getRange(SC_TOP, 1, Math.max(1, t.getLastRow() - SC_TOP + 1), 1).getValues(); let n = 0; while (n < v.length && String(v[n][0]).trim() && String(v[n][0]).trim() !== AVG) n++; return n; };
const colA = n => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; }; // 3 → 'C'

function onOpen() {
  SpreadsheetApp.getUi().createMenu('채점표').addItem('채점표 만들기', '채점표만들기').addItem('수행평가성적에 전체 반영', '채점표반영').addItem('채점표 모양 맞추기 (기본점수·만점·반 평균)', '채점표정리').addItem('통계 만들기', '통계만들기')
    .addSeparator().addItem('학기 변경', '학기변경').addItem('연도 변경', '연도변경').addItem('반 계정 만들기', '반계정만들기').addToUi();
}

// 채점표 3행(항목)·2행(비고)에서 옮길 열 찾기 → [{c: 열 위치(0부터), name: 수행평가성적 열 이름}]
function 채점표열(r2, r3) {
  const out = [];
  r3.forEach((x, c) => { const n = String(x).trim(); if (c >= 2 && n && n !== '소계' && n !== 기본칸 && n !== 만점칸) out.push({ c: c, name: n }); });
  const ni = r2.map(x => String(x).trim()).indexOf('비고'); if (ni >= 0) out.push({ c: ni, name: '비고' });
  return out;
}
const 칸값 = (name, v) => name === '비고' ? String(v == null ? '' : v) : (isNum(v) ? Number(v) : '');

function 채점표만들기() {
  캐시비우기();
  const ss = SpreadsheetApp.getActive(), yr = curYear(), sm = curSem();
  const items = critList(yr).filter(c => c.levels.length && c.sem === sm);
  if (!items.length) { Logger.log('❌ ' + yr + '년 ' + sm + '학기 평가기준(단계 점수)이 없어요.'); return; }
  // 열 배치: A 학번, B 이름, 영역마다 항목들 + 소계, 마지막에 총점 · 비고 (영역마다 기본점수·만점 칸은 채점표정리가 소계 오른쪽에 넣음)
  const areas = [];
  items.forEach(c => { let a = areas.find(x => x.name === c.area); if (!a) areas.push(a = { name: c.area, items: [] }); a.items.push(c); });
  let k = 3; areas.forEach(a => { a.from = k; a.items.forEach(c => { c.col = k++; }); a.sub = k++; });
  const TOT = k, NOTE = k + 1, NC = NOTE;
  // 수행평가성적에 이미 있는 점수
  const g = sh('수행평가성적'), gv = g && g.getLastRow() > 1 ? g.getDataRange().getValues() : [[]];
  const GH = gv.shift().map(x => String(x).trim()), gy = GH.indexOf('연도'), gs = GH.indexOf('학번'), gn = GH.indexOf('비고'), old = {};
  if (gs >= 0) gv.forEach(r => { if ((String(gy < 0 ? '' : r[gy]).trim() || BASE()) === yr) old[nk(r[gs])] = r; });
  const st = students().filter(r => effY(r) === yr && Number(r['반']) > 0);
  const classes = [...new Set(st.map(r => Number(r['반'])))].sort((a, b) => a - b), made = [], skip = [];
  classes.forEach(cl => {
    const name = cl + '반 채점표'; if (ss.getSheetByName(name)) { skip.push(name); return; }
    const S = st.filter(r => Number(r['반']) === cl).sort((x, y) => String(x['학번']).localeCompare(String(y['학번'])));
    const n = S.length, last = SC_TOP + n - 1, t = ss.insertSheet(name, ss.getNumSheets());
    const top = [0, 1, 2, 3].map(() => new Array(NC).fill(''));
    top[0][0] = yr + '학년도 ' + sm + '학기 ' + String(S[0]['학번']).charAt(0) + '학년 ' + cl + '반 정보 수행평가 채점표';
    top[1][0] = '학번'; top[1][1] = '이름'; top[1][TOT - 1] = '총점'; top[1][NOTE - 1] = '비고';
    areas.forEach(a => { top[1][a.from - 1] = a.name; top[2][a.sub - 1] = '소계'; top[3][a.sub - 1] = '=SUM(' + colA(a.from) + '4:' + colA(a.sub - 1) + '4)'; });
    items.forEach(c => { top[2][c.col - 1] = c.name; top[3][c.col - 1] = c.full; });
    top[3][0] = '배점'; top[3][TOT - 1] = '=' + areas.map(a => colA(a.sub) + '4').join('+');
    t.getRange(1, 1, 4, NC).setValues(top);
    const body = S.map((r, i) => {
      const row = SC_TOP + i, o = old[nk(r['학번'])], w = new Array(NC).fill('');
      const rg = areas.map(a => colA(a.from) + row + ':' + colA(a.sub - 1) + row), all = rg.join(',');
      w[0] = String(r['학번']); w[1] = r['이름'];
      items.forEach(c => { const j = GH.indexOf(c.name); w[c.col - 1] = o && j >= 0 && isNum(o[j]) ? Number(o[j]) : ''; });
      areas.forEach((a, x) => { w[a.sub - 1] = '=IF(COUNT(' + rg[x] + ')=0,"",SUM(' + rg[x] + '))'; });
      w[TOT - 1] = '=IF(COUNT(' + all + ')=0,"",SUM(' + areas.map(a => colA(a.sub) + row).join(',') + '))';
      w[NOTE - 1] = o && gn >= 0 ? String(o[gn]).replace(/^(?=[=+\-@])/, ' ') : '';
      return w;
    });
    t.getRange(SC_TOP, 1, n, 1).setNumberFormat('@'); // 학번은 글자로
    t.getRange(SC_TOP, 1, n, NC).setValues(body);
    items.forEach(c => t.getRange(SC_TOP, c.col, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(c.levels.slice().sort((p, q) => q - p).map(String), true).setAllowInvalid(true).setHelpText('평가기준 단계 점수 중에서 선택하세요.').build()));
    // 서식
    t.getRange(1, 1, last, NC).setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    t.getRange(1, 1).setFontSize(13).setFontWeight('bold').setHorizontalAlignment('left');
    t.getRange(2, 1, 3, NC).setBackground('#e8e0f7').setFontWeight('bold').setWrap(true);
    const AC = ['#dce8f8', '#e2f2e2', '#fcebd9', '#f3e5f5'];
    areas.forEach((a, x) => t.getRange(2, a.from, 1, a.sub - a.from + 1).merge().setBackground(AC[x % AC.length]));
    [1, 2, TOT, NOTE].forEach(c => t.getRange(2, c, 2, 1).merge());
    t.getRange(4, 1, 1, 2).merge();
    areas.map(a => a.sub).concat([TOT]).forEach(c => t.getRange(SC_TOP, c, n, 1).setBackground('#f2f2f2').setFontWeight('bold'));
    t.getRange(SC_TOP, NOTE, n, 1).setHorizontalAlignment('left');
    t.getRange(2, 1, last - 1, NC).setBorder(true, true, true, true, true, true, '#bfbfbf', SpreadsheetApp.BorderStyle.SOLID);
    t.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A' + SC_TOP + '<>"",ISBLANK(' + colA(areas[0].from) + SC_TOP + '))').setBackground('#fff8d6')
        .setRanges(areas.map(a => t.getRange(SC_TOP, a.from, n, a.sub - a.from))).build()]); // 빈 점수 칸은 연노랑
    t.setFrozenRows(4); t.setFrozenColumns(2);
    t.setColumnWidths(1, NC, 78); t.setColumnWidth(NOTE, 160); t.setRowHeight(1, 34); t.setRowHeight(3, 58);
    if (t.getMaxColumns() > NC) t.deleteColumns(NC + 1, t.getMaxColumns() - NC);
    if (t.getMaxRows() > last + 4) t.deleteRows(last + 5, t.getMaxRows() - last - 4);
    made.push(name + '(' + n + '명)');
  });
  채점표정리(true); // 반 평균 줄 + 영역별 기본점수·만점 칸
  Logger.log('만든 채점표: ' + (made.length ? made.join(', ') : '없음') + (skip.length ? ' / 이미 있어서 건너뜀: ' + skip.join(', ') : ''));
}

// 수행평가성적 시트 읽기 (학번으로 줄 찾기, 없으면 새 줄)
function 성적시트(yr) {
  const s = sh('수행평가성적'); if (!s) return null;
  const H = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const iy = H.indexOf('연도'), is = H.indexOf('학번'), inm = H.indexOf('이름'); if (is < 0) return null;
  const data = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, H.length).getValues() : [];
  const find = sid => data.findIndex(r => nk(r[is]) === sid && (iy < 0 ? yr : (String(r[iy]).trim() || BASE())) === yr);
  const blank = (id, nm) => { const w = new Array(H.length).fill(''); if (iy >= 0) w[iy] = Number(yr); w[is] = String(id); if (inm >= 0) w[inm] = nm; return w; };
  return { s: s, H: H, data: data, find: find, blank: blank, inm: inm, is: is };
}

function 채점표반영(only, year) { // only·year: 학기변경/연도변경에서만 넘김 (메뉴에서 누르면 모든 채점표 → 올해)
  캐시비우기();
  const yr = year || curYear(), tabs = Array.isArray(only) ? only : SpreadsheetApp.getActive().getSheets().filter(s => SC_TAB.test(s.getName().trim()));
  if (!tabs.length) { Logger.log("❌ '1반 채점표' 같은 탭이 없어요. 먼저 '채점표만들기'를 실행하세요."); return; }
  const G = 성적시트(yr); if (!G) { Logger.log('❌ 수행평가성적 시트(학번 열)가 없어요'); return; }
  const unk = {}; let upd = 0, made = 0;
  tabs.forEach(t => {
    const v = t.getDataRange().getValues(), cols = 채점표열(v[1], v[2]);
    for (let i = SC_TOP - 1; i < v.length; i++) {
      const sid = nk(v[i][0]); if (!sid || sid === AVG) continue;
      let j = G.find(sid);
      if (j < 0) { if (cols.every(x => String(v[i][x.c]).trim() === '')) continue; G.data.push(G.blank(v[i][0], v[i][1])); j = G.data.length - 1; made++; } else upd++;
      const row = G.data[j];
      if (G.inm >= 0 && !String(row[G.inm]).trim()) row[G.inm] = v[i][1];
      cols.forEach(x => { const k = G.H.indexOf(x.name); if (k < 0) unk[x.name] = 1; else row[k] = 칸값(x.name, v[i][x.c]); });
    }
  });
  if (G.data.length) { G.s.getRange(1, G.is + 1, G.s.getMaxRows(), 1).setNumberFormat('@'); G.s.getRange(2, 1, G.data.length, G.H.length).setValues(G.data); }
  캐시비우기();
  Logger.log('완료: 갱신 ' + upd + '명, 새로 만든 학생 ' + made + '명 (' + yr + '년)');
  if (Object.keys(unk).length) Logger.log('⚠ 채점표에는 있는데 수행평가성적 시트에 같은 이름의 열이 없는 항목(성적열추가를 실행하세요): ' + JSON.stringify(Object.keys(unk)));
}

/* ---------- 학기 변경 / 연도 변경 (시트 메뉴 '채점표'에서 실행) ----------
   학기 변경 : 지금 채점표 점수를 수행평가성적에 저장 → 채점표 탭 이름을 '1반 채점표(2026-1학기)'처럼 바꿔 보관
               → 성적열추가 → 이번 학기 채점표 새로 만들기. (먼저 평가기준에 이번 학기 줄을 넣어 두세요)
   연도 변경 : 위와 같고, 새 학년도 계정(반 계정 만들기 + 이름 입력)과 평가기준·등급기준 새 연도 줄이 있는지도 확인합니다.
   학기는 8월부터 2학기, 학년도는 3월부터 새 해로 자동으로 바뀝니다. 미리 바꾸려면 설정 시트에 '현재학기' / '현재연도' 줄을 쓰세요.
   보관한 탭은 코드가 읽지 않으니 남겨 두거나 지워도 됩니다. */
const 채점표탭 = () => SpreadsheetApp.getActive().getSheets().filter(s => SC_TAB.test(s.getName().trim()));
const 탭학기 = t => { const m = /(\d{4})학년도\s*([12])학기/.exec(String(t.getRange(1, 1).getValue())); return m ? { y: m[1], s: Number(m[2]) } : null; }; // 채점표 1행 제목에서 읽음
function 채점표보관(tabs) { // 점수를 그 채점표의 연도 줄에 저장한 뒤 이름을 바꿔 보관
  const ss = SpreadsheetApp.getActive(), by = {};
  tabs.forEach(t => { const k = 탭학기(t), y = k ? k.y : curYear(); (by[y] = by[y] || []).push(t); });
  Object.keys(by).forEach(y => 채점표반영(by[y], y));
  tabs.forEach(t => { const k = 탭학기(t), base = t.getName().trim() + '(' + (k ? k.y + '-' + k.s + '학기' : '보관') + ')'; let n = base, i = 2; while (ss.getSheetByName(n)) n = base + i++; t.setName(n); });
}
function 바꾸기(kind) {
  const ui = SpreadsheetApp.getUi(); 캐시비우기();
  const yr = curYear(), sm = curSem(), tabs = 채점표탭(), what = yr + '년 ' + sm + '학기', warn = [];
  if (tabs.some(t => { const k = 탭학기(t); return k && k.y === yr && k.s === sm; })) {
    ui.alert(kind, '채점표가 이미 ' + what + ' 것이에요.\n학기는 8월부터 2학기, 학년도는 3월부터 새 해로 바뀌어요. 미리 바꾸려면 설정 시트에 현재학기 / 현재연도 줄을 넣어 주세요.', ui.ButtonSet.OK); return;
  }
  if (!critList(yr).some(c => c.levels.length && c.sem === sm)) { ui.alert(kind, what + ' 평가기준이 없어요. 평가기준 시트에 연도 ' + yr + ', 학기 ' + sm + '인 줄을 먼저 넣어 주세요.', ui.ButtonSet.OK); return; }
  if (kind === '연도 변경') {
    const st = students().filter(r => effY(r) === yr && Number(r['반']) > 0);
    if (!st.length) { ui.alert(kind, yr + '년 학생 계정이 없어요. 먼저 메뉴의 \'반 계정 만들기\'를 실행하고 이름을 넣어 주세요.', ui.ButtonSet.OK); return; }
    const noName = st.filter(r => !String(r['이름'] || '').trim()).length;
    if (noName) warn.push('이름이 빈 ' + yr + '년 계정 ' + noName + '개 (채점표에는 만들 때의 이름이 들어가요)');
    if (!rows('등급기준').some(x => String(x['연도'] || BASE()).trim() === yr)) warn.push('등급기준에 ' + yr + '년 줄이 없어요 (성적 등급이 안 나와요)');
  }
  const msg = (tabs.length ? '지금 채점표 ' + tabs.length + '개의 점수를 수행평가성적에 저장하고, 이름을 바꿔 보관해요.\n' : '') + what + ' 채점표를 새로 만들어요.' + (warn.length ? '\n\n⚠ ' + warn.join('\n⚠ ') : '') + '\n\n계속할까요?';
  if (ui.alert(kind, msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  채점표보관(tabs);
  성적열추가();
  채점표만들기();
  ui.alert(kind, '완료: ' + what + ' 채점표 ' + 채점표탭().length + '개를 만들었어요.' + (tabs.length ? ' 이전 채점표는 이름 끝에 (연도-학기)를 붙여 보관했어요.' : ''), ui.ButtonSet.OK);
}
function 학기변경() { 바꾸기('학기 변경'); }
function 연도변경() { 바꾸기('연도 변경'); }

/* ---------- 채점표 모양 맞추기: 영역별 기본점수·만점 칸 + 반 평균 줄 ----------
   기본점수: 영역마다 소계 오른쪽에 '기본점수' 체크칸이 있습니다. 체크하면 그 학생의 그 영역 항목이 평가기준의 가장 낮은 단계 점수로 바뀌고(이미 있던 점수도 덮어씀),
             체크를 풀면 그 영역 항목이 빈 칸이 됩니다. 수행평가성적에도 바로 저장됩니다.
   만점    : 기본점수 오른쪽의 '만점' 체크칸. 체크하면 그 영역 항목이 가장 높은 단계 점수로 바뀌고, 체크를 풀면 빈 칸이 됩니다.
             기본점수와 만점은 한쪽을 체크하면 다른 쪽 체크가 풀립니다.
   반 평균 : 학생 바로 아래 줄에 영역 소계·총점의 반 평균(소수 첫째 자리)을 수식으로 보여 줍니다. 점수가 빈 학생은 평균에서 빠집니다.
   새로 만드는 채점표에는 자동으로 들어가고, 이미 있는 채점표는 메뉴 '채점표 모양 맞추기'를 한 번 누르세요
   ('채점 완료' 열과 예전 버전의 오른쪽 끝 기본점수 칸은 지워 줍니다). 여러 번 눌러도 이미 있는 것은 다시 만들지 않습니다. */
const 기본칸 = '기본점수', 만점칸 = '만점';
const 영역시작 = (r2, c) => { while (c > 2 && !String(r2[c]).trim()) c--; return c; }; // c열이 속한 영역의 첫 열(0부터)
function 채점표정리(quiet) {
  let done = 0;
  채점표탭().forEach(t => {
    const n = 학생수(t); if (n < 1) return;
    const head = () => { const v = t.getRange(2, 1, 2, t.getLastColumn()).getValues(); return v.map(r => r.map(x => String(x).trim())); };
    let [r2, r3] = head(), changed = false;
    const dc = r2.indexOf('채점 완료'); if (dc >= 0) { t.deleteColumn(dc + 1); [r2, r3] = head(); changed = true; }
    const old = r2.indexOf(기본칸); // 예전 버전: 오른쪽 끝에 모아 둔 기본점수 칸
    if (old >= 0) { let k = 1; while (old + k < r2.length && !r2[old + k] && r3[old + k]) k++; t.deleteColumns(old + 1, k); [r2, r3] = head(); changed = true; }
    for (let c = r3.length - 1; c >= 2; c--) { // 오른쪽 영역부터 (열을 끼워 넣어도 왼쪽 위치가 안 바뀜)
      if (r3[c] !== '소계') continue;
      const add = [기본칸, 만점칸].filter((x, i) => r3[c + 1 + i] !== x); if (!add.length) continue; // v89 채점표는 만점 칸만 더함
      const a = 영역시작(r2, c), s = c + 4 - add.length, k = add.length; // s: 새 칸의 첫 열(1부터)
      t.insertColumnsAfter(s - 1, k);
      t.getRange(3, s, 1, k).setValues([add]);
      t.getRange(2, a + 1, 1, c - a + 3).breakApart().merge();
      t.getRange(3, s, 2, k).setBackground('#e8e0f7').setFontWeight('bold').setFontSize(9);
      t.getRange(SC_TOP, s, n, k).setBackground('#ffffff').setFontWeight('normal').insertCheckboxes();
      t.setColumnWidths(s, k, 58); changed = true;
    }
    [r2, r3] = head();
    const lc = r2.length, ar = SC_TOP + n, last = ar - 1;
    if (String(t.getRange(ar, 1).getValue()).trim() !== AVG) { // 반 평균 줄
      if (t.getMaxRows() < ar) t.insertRowsAfter(t.getMaxRows(), ar - t.getMaxRows());
      t.getRange(ar, 1).setValue(AVG);
      for (let c = 0; c < lc; c++) if (r3[c] === '소계' || r2[c] === '총점') { const L = colA(c + 1); t.getRange(ar, c + 1).setValue('=IFERROR(ROUND(AVERAGE(' + L + SC_TOP + ':' + L + last + '),1),"")'); }
      t.getRange(ar, 1, 1, lc).setBackground('#fff3e0').setFontWeight('bold').setHorizontalAlignment('center');
      changed = true;
    }
    if (changed) done++;
  });
  if (!quiet) { const m = done ? '채점표 ' + done + '개를 새 모양(영역별 기본점수·만점·반 평균)으로 맞췄어요.' : '모든 채점표가 이미 새 모양이에요.'; Logger.log(m); try { SpreadsheetApp.getActive().toast(m, '채점표', 10); } catch (e) {} }
}
// 고친 칸 중 기본점수·만점 체크칸이 있으면: 체크 → 그 영역 항목을 가장 낮은(만점은 가장 높은) 단계 점수로, 체크 해제 → 빈 칸으로. 바꾼 열 목록을 돌려줌
function 기본점수채우기(t, r0, r1, c0, c1, head, all) {
  const r2 = head[0].map(x => String(x).trim()), r3 = head[1].map(x => String(x).trim()), boxes = [];
  for (let c = c0 - 1; c <= c1 - 1; c++) if (r3[c] === 기본칸 || r3[c] === 만점칸) boxes.push({ c: c, from: 영역시작(r2, c), max: r3[c] === 만점칸 });
  if (!boxes.length) return [];
  const k = 탭학기(t), y = k ? k.y : curYear(), sm = k ? k.s : curSem(), lo = {}, hi = {};
  critAll().filter(c => c.year === y && c.sem === sm && c.levels.length).forEach(c => { lo[c.name] = Math.min.apply(null, c.levels); hi[c.name] = Math.max.apply(null, c.levels); });
  const v = t.getRange(r0, 1, r1 - r0 + 1, t.getLastColumn()).getValues(), used = [];
  boxes.forEach(b => {
    const items = all.filter(x => x.name !== '비고' && x.c >= b.from && x.c < b.c), pick = b.max ? hi : lo;
    const o = b.max ? b.c - 1 : b.c + 1, pair = r3[o] === (b.max ? 기본칸 : 만점칸); // 짝 체크칸(기본점수 ↔ 만점)
    v.forEach((r, i) => {
      if (!nk(r[0]) || nk(r[0]) === AVG || (r[b.c] !== true && r[b.c] !== false)) return;
      if (r[b.c] === true && pair && r[o] === true) t.getRange(r0 + i, o + 1).setValue(false);
      items.forEach(x => { if (r[b.c] === true && pick[x.name] == null) return; t.getRange(r0 + i, x.c + 1).setValue(r[b.c] === true ? pick[x.name] : ''); if (used.indexOf(x) < 0) used.push(x); });
    });
  });
  return used;
}

// 채점표에서 고친 칸만 수행평가성적에 바로 옮김 (간단한 트리거라 따로 설정할 필요 없음)
function onEdit(e) {
  const t = e.range.getSheet(); if (!SC_TAB.test(t.getName().trim())) return;
  const r0 = Math.max(e.range.getRow(), SC_TOP), r1 = e.range.getLastRow(), c0 = e.range.getColumn(), c1 = e.range.getLastColumn();
  if (r1 < r0) return;
  const lc = t.getLastColumn(), head = t.getRange(2, 1, 2, lc).getValues(), all = 채점표열(head[0], head[1]);
  const cols = all.filter(x => x.c + 1 >= c0 && x.c + 1 <= c1);
  기본점수채우기(t, r0, r1, c0, c1, head, all).forEach(x => { if (cols.indexOf(x) < 0) cols.push(x); });
  if (!cols.length) return;
  const v = t.getRange(r0, 1, r1 - r0 + 1, lc).getValues(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const G = 성적시트(curYear()); if (!G) return;
    v.forEach(r => {
      const sid = nk(r[0]); if (!sid || sid === AVG) return;
      let j = G.find(sid);
      if (j < 0) { const w = G.blank(r[0], r[1]); G.s.appendRow(w); G.data.push(w); j = G.data.length - 1; }
      cols.forEach(x => { const k = G.H.indexOf(x.name); if (k >= 0) G.s.getRange(j + 2, k + 1).setValue(칸값(x.name, r[x.c])); });
    });
  } finally { lock.releaseLock(); }
}

/* 성적 입력 후 실행하면 오타·빠진 열·범위를 벗어난 점수를 실행 로그에 알려 줍니다 */
function 성적점검() {
  캐시비우기();
  const yr = curYear(), C = critList(yr), s = sh('수행평가성적');
  Logger.log(yr + '년 평가기준: 항목 ' + C.length + '개 (단계점수가 있는 항목 ' + C.filter(c => c.levels.length).length + '개)');
  const seen = {}, dup = []; C.forEach(c => { const k = c.sem + '|' + c.name; if (seen[k]) dup.push(c.sem + '학기 ' + c.name); seen[k] = 1; });
  if (dup.length) Logger.log('⚠ 이름이 겹치는 항목(점수 열을 구분할 수 없어요): ' + JSON.stringify(dup));
  if (!s) { Logger.log('❌ 수행평가성적 시트가 없어요'); return; }
  const v = s.getDataRange().getValues(), H = v.shift().map(x => String(x).trim()), names = Array.from(new Set(C.filter(c => c.levels.length).map(c => c.name))), res = ['연도', '학기', '학번', '이름', '비고'];
  Logger.log('성적 시트에 없는 항목 열: ' + JSON.stringify(names.filter(n => H.indexOf(n) < 0)));
  Logger.log('평가기준에 없는 열(오타?): ' + JSON.stringify(H.filter(h => h && res.indexOf(h) < 0 && names.indexOf(h) < 0)));
  const iy = H.indexOf('연도'), is = H.indexOf('학번'), im = H.indexOf('학기'); let bad = 0, cnt = 0;
  v.forEach((r, i) => {
    if ((String(iy < 0 ? '' : r[iy]).trim() || String(BASE())) !== yr || String(r[is]) === '') return; cnt++;
    const rs = String(im < 0 ? '' : r[im]).trim(); C.filter(c => c.levels.length && (!rs || String(c.sem) === rs)).forEach(c => { const k = H.indexOf(c.name); if (k < 0 || r[k] === '') return; const x = Number(r[k]); if (isNaN(x) || x < 0 || x > c.full) { bad++; Logger.log('⚠ ' + (i + 2) + '행 ' + r[is] + ' · ' + c.name + ' 점수 ' + r[k] + ' (0~' + c.full + ' 범위 밖)'); } });
  });
  Logger.log('성적이 입력된 학생 ' + cnt + '명, 범위 밖 점수 ' + bad + '개');
}

/* ---------- 통계 탭 (시트 메뉴 '채점표 → 통계 만들기') ----------
   올해·현재 학기 성적으로 '통계' 탭을 새로 씁니다(누를 때마다 다시 계산).
   1) 반별 등급 분포 + 미채점 학생(번호 이름)  2) 영역별 반 평균·총 평균  3) 합계 점수 분포(5점 간격) 막대그래프
   점수 계산·등급은 학생 성적 화면(gradeData)과 같은 규칙: 숫자 칸만 합산, 등급기준의 최소점수 이상 중 가장 높은 등급.
   점수가 하나도 없는 학생은 미채점으로 보고 평균·그래프에서 뺍니다. */
function 통계만들기() {
  캐시비우기();
  const ss = SpreadsheetApp.getActive(), yr = curYear(), sm = curSem(), R1 = x => Math.round(x * 10) / 10;
  const C = critList(yr).filter(c => c.levels.length && c.sem === sm);
  if (!C.length) { Logger.log('❌ ' + yr + '년 ' + sm + '학기 평가기준(단계 점수)이 없어요.'); return; }
  const areas = [...new Set(C.map(c => c.area))], amax = areas.map(a => C.filter(c => c.area === a).reduce((s, c) => s + c.full, 0)), max = amax.reduce((s, x) => s + x, 0);
  const G = rows('등급기준').filter(x => String(x['연도'] || BASE()).trim() === yr && x['최소점수'] !== '' && (!String(x['학기'] || '').trim() || String(x['학기']).trim() === String(sm)))
    .sort((a, b) => Number(b['최소점수']) - Number(a['최소점수']));
  const grades = [...new Set(G.map(x => String(x['등급'])))].sort();
  // 수행평가성적: 올해 줄 중 이번 학기 줄(없으면 학기 칸이 빈 줄)
  const s = sh('수행평가성적'), v = s && s.getLastRow() > 1 ? s.getDataRange().getValues() : [[]];
  const H = v.shift().map(x => String(x).trim()), iy = H.indexOf('연도'), is = H.indexOf('학번'), im = H.indexOf('학기'), byId = {};
  if (is >= 0) v.forEach(r => {
    if ((String(iy < 0 ? '' : r[iy]).trim() || BASE()) !== yr) return;
    const k = nk(r[is]), cs = String(im < 0 ? '' : r[im]).trim();
    if (cs === String(sm)) byId[k] = r; else if (cs === '' && !byId[k]) byId[k] = r;
  });
  const st = students().filter(r => effY(r) === yr && Number(r['반']) > 0);
  const classes = [...new Set(st.map(r => Number(r['반'])))].sort((a, b) => a - b), done = [];
  const cls = classes.map(cl => {
    const o = { cl: cl, n: 0, g: {}, miss: [], asum: areas.map(() => 0), cnt: 0 };
    st.filter(r => Number(r['반']) === cl).sort((x, y) => String(x['학번']).localeCompare(String(y['학번']))).forEach(p => {
      o.n++;
      const r = byId[nk(p['학번'])], a = areas.map(() => 0); let any = false;
      if (r) C.forEach(c => { const k = H.indexOf(c.name), x = k < 0 ? '' : r[k]; if (isNum(x)) { any = true; a[areas.indexOf(c.area)] += Number(x); } });
      if (!any) { o.miss.push(Number(String(p['학번']).slice(-2)) + '번 ' + String(p['이름'] || '').trim()); return; }
      const tot = a.reduce((s, x) => s + x, 0), hit = G.find(x => tot >= Number(x['최소점수'])), gr = hit ? String(hit['등급']) : '';
      o.g[gr] = (o.g[gr] || 0) + 1; o.cnt++; a.forEach((x, i) => o.asum[i] += x); done.push(tot);
    });
    return o;
  });
  // 표 만들기 (모든 줄 너비를 W로 맞춰 한 번에 씀)
  const W = Math.max(grades.length + 3, areas.length + 2, 3), out = [], fmt = [];
  const row = (a, kind) => { while (a.length < W) a.push(''); out.push(a); fmt.push(kind || ''); };
  const none = cls.some(o => o.g['']);
  row([yr + '학년도 ' + sm + '학기 정보 수행평가 통계 (만든 시각: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + ')'], 'title'); row([]);
  row(['1. 반별 등급 분포 (명)'], 'sec');
  row(['반'].concat(grades, none ? ['등급없음'] : [], ['인원', '미채점 (번호 이름)']), 'head');
  cls.forEach(o => row([o.cl + '반'].concat(grades.map(g => o.g[g] || 0), none ? [o.g[''] || 0] : [], [o.n, o.miss.join(', ')])));
  row(['전체'].concat(grades.map(g => cls.reduce((s, o) => s + (o.g[g] || 0), 0)), none ? [cls.reduce((s, o) => s + (o.g[''] || 0), 0)] : [], [cls.reduce((s, o) => s + o.n, 0), cls.reduce((s, o) => s + o.miss.length, 0) + '명']), 'sum');
  row([]);
  row(['2. 수행평가 영역별 평균 (점) — 채점된 학생만'], 'sec');
  row(['반'].concat(areas.map((a, i) => a + ' (' + amax[i] + '점)'), ['합계 (' + max + '점)']), 'head');
  const avg = (sum, n) => n ? R1(sum / n) : '';
  cls.forEach(o => row([o.cl + '반 평균'].concat(o.asum.map(x => avg(x, o.cnt)), [avg(o.asum.reduce((s, x) => s + x, 0), o.cnt)])));
  const N0 = cls.reduce((s, o) => s + o.cnt, 0);
  row(['총 평균'].concat(areas.map((a, i) => avg(cls.reduce((s, o) => s + o.asum[i], 0), N0)), [avg(done.reduce((s, x) => s + x, 0), N0)]), 'sum');
  row([]);
  row(['3. 점수 분포 (합계 점수, 채점된 학생 전체)'], 'sec');
  row(['점수 구간', '학생 수'], 'head');
  const b0 = 5, chartTop = out.length + 1;
  if (done.length) {
    const lo = Math.floor(Math.min.apply(null, done) / b0) * b0, hi = Math.max(max, Math.max.apply(null, done)), last = Math.floor((hi - 1e-9) / b0) * b0;
    for (let x = lo; x <= last; x += b0) row([x + '~' + (x + b0 >= hi ? hi : x + b0 - 1), done.filter(t => t >= x && (x === last ? t <= hi : t < x + b0)).length]);
  } else row(['(채점된 학생 없음)', 0]);
  const chartEnd = out.length;
  // 시트에 쓰기
  const t = ss.getSheetByName('통계') || ss.insertSheet('통계', ss.getNumSheets());
  t.getCharts().forEach(c => t.removeChart(c)); t.clear();
  t.getRange(1, 1, out.length, W).setValues(out).setFontFamily('Gowun Dodum').setVerticalAlignment('middle');
  let hn = 0; // 표 너비 = 바로 위 제목 줄의 칸 수 (미채점 칸이 비어도 테두리를 맞춤)
  fmt.forEach((k, i) => {
    if (k === 'head') hn = out[i].filter(x => x !== '').length;
    const n = out[i].some(x => x !== '') ? hn : 0;
    if (k === 'title') t.getRange(i + 1, 1).setFontWeight('bold').setFontSize(13);
    if (k === 'sec') t.getRange(i + 1, 1).setFontWeight('bold').setBackground('#fff7d6');
    if (k === 'head' || k === 'sum' || k === '' && n) t.getRange(i + 1, 1, 1, n).setBorder(true, true, true, true, true, true, '#d0d0d0', null).setHorizontalAlignment('center');
    if (k === 'head') t.getRange(i + 1, 1, 1, n).setBackground('#eef4ff').setWrap(true);
    if (k === 'sum') t.getRange(i + 1, 1, 1, n).setBackground('#f3f3f3').setFontWeight('bold');
  });
  // 미채점 칸은 왼쪽 정렬·빨간 글씨
  const mc = grades.length + (none ? 3 : 2) + 1;
  t.getRange(5, mc, cls.length, 1).setHorizontalAlignment('left').setFontColor('#c5221f');
  t.setColumnWidth(1, 110); for (let c = 2; c <= W; c++) t.setColumnWidth(c, 90); t.setColumnWidth(mc, Math.max(260, t.getColumnWidth(mc)));
  t.setFrozenRows(1);
  t.insertChart(t.newChart().asColumnChart().addRange(t.getRange(chartTop, 1, chartEnd - chartTop + 1, 2))
    .setNumHeaders(1).setOption('title', '합계 점수 구간별 학생 수').setOption('legend', { position: 'none' }).setOption('colors', ['#a8c8f0'])
    .setOption('hAxis', { title: '점수 구간 (점)' }).setOption('vAxis', { title: '학생 수', format: '0' }).setOption('width', 520).setOption('height', 320)
    .setPosition(chartTop - 1, 4, 10, 0).build());
  const m = '통계 탭을 새로 만들었어요: ' + yr + '년 ' + sm + '학기, 채점된 학생 ' + N0 + '명 / 미채점 ' + cls.reduce((s, o) => s + o.miss.length, 0) + '명';
  Logger.log(m); try { ss.toast(m, '통계', 10); } catch (e) {}
}
