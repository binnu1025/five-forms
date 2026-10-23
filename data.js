/* 後台是 Google Sheet，中間隔一層 Apps Script。
   網頁只跟那個端點講話，從頭到尾碰不到試算表本身，
   所以 Sheet 可以只分享給會長與團長。

   在 app.js 之後載入，直接沿用它的 groups / flow / render* 等全域。
   讀不到就在選擇頁說清楚，不要退回任何假資料。 */

const API = 'https://script.google.com/macros/s/AKfycbwpJuvmtI_4jMG6NcEb_JtD15xg75Lzow_qQcA5yKu1gzBEWQ8fK4uDwLK3MtQzMBRDRg/exec';

// 目前是誰在用。接上 Google 登入後這裡換成 id_token，identify_() 已經準備好了。
const whoAmI = () => localStorage.getItem('five-gmail') || 'binnu1025@gmail.com';

let sheetReady = false;

/* Apps Script 偶爾會直接 Failed to fetch（實測 6 次有 1 次），
   失敗就安靜退回備份資料、寫入也會掉，使用者不會發現。所以重試。 */
async function withRetry(fn, tries = 3) {
  let last;
  for (let i = 0; i < tries; i++) {
    try { return await fn(); } catch (err) { last = err; await new Promise(r => setTimeout(r, 400 * (i + 1))); }
  }
  throw last;
}

async function apiGet() {
  return withRetry(async () => {
    const res = await fetch(`${API}?who=${encodeURIComponent(whoAmI())}&t=${Date.now()}`);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || '端點回報失敗');
    return json;
  });
}

/* 整張表換掉，不用去追哪一列改了哪一欄 */
async function apiReplace(sheet, rows) {
  if (!sheetReady) return;
  try {
    const out = await withRetry(async () => {
      const res = await fetch(API, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ sheet, who: whoAmI(), replace: true, rows }),
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    });
    if (!out.ok) toast('寫回試算表失敗：' + out.error);
  } catch (err) {
    toast('寫回試算表失敗，改動只留在這支手機上');
    console.warn('[五表] 寫回失敗：', err);
  }
}

/* ── Sheet 的列 ↔ app.js 的資料結構 ───────────────────────── */

const rowsToGroups = (roster, order) => {
  const byGroup = {};
  roster.filter(r => r['組別'] && r['組別'] !== '核心').forEach(r => {
    const g = byGroup[r['組別']] || (byGroup[r['組別']] = { lead: '', members: [] });
    if (r['角色'] === '組長') g.lead = r['姓名']; else g.members.push(r['姓名']);
  });
  const names = order.length ? order : Object.keys(byGroup);
  return names.filter(n => byGroup[n]).map(n => [n, byGroup[n].lead, byGroup[n].members]);
};

const rowsToFlow = rows => {
  const slots = [];
  rows.forEach(r => {
    const time = r['結束'] && r['結束'] !== r['開始'] ? `${r['開始']}–${r['結束']}` : String(r['開始'] || '');
    let slot = slots.find(s => s.time === time && s.name === r['全體流程']);
    if (!slot) slots.push(slot = { time, name: r['全體流程'] || '（未命名時段）', jobs: {} });
    if (r['組別'] && r['該組工作']) slot.jobs[r['組別']] = r['該組工作'];
  });
  return slots;
};

const progressToRows = () => progressRecords.map(t => ({
  活動ID: window.__eventId || 'A001', 會議: t.meeting, 組別: t.group, 負責人: t.owner,
  要完成事項: t.item, 目前進度: t.update, 下一步: t.next, 期限: t.due, 狀態: t.status,
}));
const materialsToRows = () => materials.map(m => ({
  活動ID: window.__eventId || 'A001', 組別: m[1], 物資: m[0], 數量: m[2], 攜帶者: m[3],
  借出照片: '', 已還: m[5] === 'done' ? 'TRUE' : 'FALSE', 已歸位: '', 歸還照片: '', 備註: m[6] || '',
}));
const financeToRows = () => budgets.map(b => ({
  活動ID: window.__eventId || 'A001', 類型: '成本', 組別: b[0], 項目: b[1],
  金額: b[2], 登錄人: b[3], 備註: b[4] || '',
}));

/* ── 畫面上那些寫死的數字 ─────────────────────────────────── */

function applyEvent(ev) {
  if (!ev || !ev.name) return;
  window.__eventId = ev.id || 'A001';
  window.__eventName = ev.name;
  window.__eventMeta = ev.meta || '';
  window.__eventDate = ev.date || '';
  document.title = `${ev.name}｜活動協作`;
  const btn = document.getElementById('eventButtonName');
  if (btn) btn.textContent = ev.name;
  document.querySelectorAll('.side-current b').forEach(el => { el.textContent = ev.name; });
  if (ev.meta) document.querySelectorAll('.side-current span').forEach(el => { el.textContent = ev.meta; });
}

/* 物資頁的提示與篩選數字原本寫死「4 件借出中／全部 6／借出中 4／已歸還 2」 */
function renderMaterialExtras() {
  const open = materials.filter(m => m[5] === 'open').length;
  const notice = document.querySelector('#materials .notice');
  if (notice) {
    notice.style.display = materials.length ? '' : 'none';
    const b = notice.querySelector('b');
    if (b) b.textContent = `${open} 件物資借出中`;
  }
  const pills = document.querySelectorAll('#materials .filter-pills button');
  if (pills[0]) pills[0].textContent = `全部 ${materials.length}`;
  if (pills[1]) pills[1].textContent = `借出中 ${open}`;
  if (pills[2]) pills[2].textContent = `已歸還 ${materials.length - open}`;
  // 頁籤上那顆紅點原本寫死 3
  const badge = document.querySelector('.category-tabs [data-page="materials"] i');
  if (badge) { badge.textContent = open; badge.style.display = open ? '' : 'none'; }
}

/* 會後精進原本是四筆寫死的示範，組名還是早就不存在的「引導組」 */
function renderReviews(rows) {
  reviews.splice(0, reviews.length, ...rows.map(r =>
    [r['組別'], [r['問題'], r['影響'], r['下次做法']].filter(Boolean).join('　→　')]));
  const el = document.getElementById('reviewList');
  if (!el) return;
  el.innerHTML = reviews.length
    ? reviews.map(r => `<article class="review-card"><h3>！${r[0]}</h3><p>${r[1]}</p></article>`).join('')
    : '<div class="progress-empty">這場活動還沒有會後精進紀錄</div>';
}

/* 歷史活動原本是三筆寫死的假活動，而且本次活動不在清單裡 */
function renderHistories(events, roster, currentId) {
  const past = events.filter(e => e['活動ID'] && e['活動ID'] !== currentId);
  histories.splice(0, histories.length, ...past.map(e => {
    const mine = roster.filter(r => r['活動ID'] === e['活動ID']);
    const gs = new Set(mine.map(r => r['組別']).filter(g => g && g !== '核心'));
    return {
      name: e['活動名稱'], date: e['日期'], place: e['地點'],
      groups: gs.size, people: mine.length, result: e['狀態'] || '已結案', notes: [],
    };
  }));
  renderHistory();
  renderEventMenu();
  const all = document.querySelector('[data-all-history]');
  if (all) all.style.display = histories.length ? '' : 'none';
  if (!histories.length && typeof eventHistoryOptions !== 'undefined' && eventHistoryOptions) {
    eventHistoryOptions.innerHTML = '<p class="dialog-note">目前只有這一場活動，還沒有已結案的紀錄。</p>';
  }
}

/* 指派／物資／預算三個對話框的「工作組」下拉是載入時就填好的，
   資料換掉之後沒人重填，所以一直停在示範的 10 組。 */
function refillGroupSelects() {
  ['assignGroup', 'materialGroup', 'budgetGroup'].forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    const keep = sel.value;
    sel.innerHTML = groups.map(g => `<option>${g[0]}</option>`).join('');
    if (groups.some(g => g[0] === keep)) sel.value = keep;
  });
}

/* 指派人員的名單原本來自寫死的 flocks（假人、假信箱），
   改成人才庫；端點只回姓名與雁群小組，信箱不會離開試算表。 */
function refillPersonSelect(pool) {
  const sel = document.getElementById('personSelect');
  if (!sel || !pool.length) return;
  const byFlock = {};
  pool.forEach(p => (byFlock[p['雁群小組'] || '未分組'] ||= []).push(p['姓名']));
  sel.innerHTML = '<option value="">請選擇雁群與人員</option>' +
    Object.entries(byFlock).map(([f, names]) =>
      `<optgroup label="${f}">${names.map(n => `<option>${n}</option>`).join('')}</optgroup>`).join('');
  const email = document.getElementById('personEmail');
  if (email) {
    const label = email.closest('label');
    if (label) label.style.display = 'none';   // 唯讀又帶不進來，在手機上只是佔位
  }
}

function renderHome() {
  document.querySelectorAll('#home .sk').forEach(el => el.classList.remove('sk', 'sk-s', 'sk-m', 'sk-l'));
  document.querySelectorAll('#home .sk-row').forEach(el => el.remove());
  const today = new Date().toISOString().slice(0, 10);
  const eventDay = window.__eventDate === today;
  const mine = groups.filter(g => g[1] === currentPerson || g[2].includes(currentPerson)).map(g => g[0]);
  const role = currentRole();
  const admin = !!role;

  /* 標題列：活動、哪一天、我是誰 */
  const when = document.getElementById('todayWhen');
  if (when) when.textContent = eventDay ? '今天就是活動當天'
    : window.__eventDate ? `${window.__eventMeta || window.__eventDate}${daysAway(window.__eventDate)}` : '';
  const title = document.getElementById('todayTitle');
  if (title) title.textContent = window.__eventName || '';
  const who = document.getElementById('todayRole');
  if (who) who.textContent = [currentPerson, role || mine.join('、') || '尚未分組'].filter(Boolean).join('・');

  /* 主區：輪到我做什麼 */
  const focusTitle = document.getElementById('focusTitle');
  const list = document.getElementById('focusList');
  const more = document.getElementById('focusMore');
  const myTasks = progressRecords.filter(t => t.status !== 'done' && (admin || mine.includes(t.group)));

  if (eventDay) {
    const slot = currentSlot();
    if (focusTitle) focusTitle.textContent = slot ? `現在・${slot.name}` : '活動流程';
    const jobs = slot ? Object.entries(slot.jobs || {}).filter(([g]) => admin || mine.includes(g)) : [];
    if (list) list.innerHTML = jobs.length
      ? jobs.map(([g, w]) => row(g, w, slot.time)).join('')
      : empty(slot ? '這個時段沒有你這一組的工作' : '細流裡還沒有時段');
    if (more) { more.hidden = false; more.textContent = '看全場細流'; more.onclick = () => go('run'); }
  } else {
    if (focusTitle) focusTitle.textContent = admin ? '還沒完成的工作'
      : mine.length ? `${mine.join('、')}還沒完成的` : '這場活動的工作';
    if (list) list.innerHTML = myTasks.length
      ? myTasks.slice(0, 4).map(t => row(t.group, t.item, dueLabel(t.due), t.status)).join('')
      : empty(mine.length || admin ? '目前沒有待辦，都完成了' : '你還沒有被指派到工作組');
    if (more) {
      more.hidden = myTasks.length <= 4;
      more.textContent = `看全部 ${myTasks.length} 項`;
      more.onclick = () => go('progress');
    }
  }

  /* 需要注意：卡住的、逾期的、沒還的 */
  const alerts = [];
  progressRecords.filter(t => t.status === 'blocked' && (admin || mine.includes(t.group)))
    .forEach(t => alerts.push(['卡住', `${t.group}・${t.item}`, 'progress']));
  progressRecords.filter(t => t.status !== 'done' && t.due && t.due < today && (admin || mine.includes(t.group)))
    .forEach(t => alerts.push(['已過期限', `${t.group}・${t.item}`, 'progress']));
  materials.filter(m => m[5] === 'open' && (admin || mine.includes(m[1])))
    .forEach(m => alerts.push(['尚未歸還', `${m[1]}・${m[0]} ${m[2]} 件`, 'materials']));
  const alertBlock = document.getElementById('attentionBlock');
  const alertList = document.getElementById('attentionList');
  if (alertBlock) alertBlock.hidden = !alerts.length;
  if (alertList) alertList.innerHTML = alerts.slice(0, 5).map(([tag, text, page]) =>
    `<button data-jump="${page}"><span>${tag}</span><b>${text}</b></button>`).join('');

  /* 各組進度：只有幹部需要看全場 */
  const teams = document.getElementById('teamsBlock');
  if (teams) {
    teams.hidden = !admin || !progressRecords.length;
    if (!teams.hidden) {
      const sub = document.getElementById('teamsSub');
      const done = progressRecords.filter(t => t.status === 'done').length;
      if (sub) sub.textContent = `${done}／${progressRecords.length} 項完成`;
      const el = document.getElementById('teamsList');
      if (el) el.innerHTML = groups.map(g => {
        const all = progressRecords.filter(t => t.group === g[0]);
        if (!all.length) return '';
        const ok = all.filter(t => t.status === 'done').length;
        const blocked = all.some(t => t.status === 'blocked');
        return `<div class="team-row ${blocked ? 'is-blocked' : ''}"><b>${g[0]}</b>
          <i><s style="width:${Math.round(ok / all.length * 100)}%"></s></i>
          <small>${ok}/${all.length}</small></div>`;
      }).join('');
    }
  }

  /* 快速入口：五表都在這裡，不要只有人力 */
  const jump = document.getElementById('homeJump');
  if (jump) {
    const openItems = materials.filter(m => m[5] === 'open').length;
    jump.innerHTML = [
      ['people', '人力表', `${groups.length} 組`],
      ['progress', '工作進度', progressRecords.length ? `${progressRecords.length} 項` : '尚未登錄'],
      ['run', '活動細流', flow.length ? `${flow.length} 個時段` : '尚未排定'],
      ['materials', '物資', materials.length ? `${openItems} 件未還` : '尚未登錄'],
      ['budget', '收支', budgets.length ? `${budgets.length} 筆` : '尚未登錄'],
      ['review', '會後精進', reviews.length ? `${reviews.length} 則` : '尚未記錄'],
    ].map(([p, name, sub]) => `<button data-jump="${p}"><b>${name}</b><span>${sub}</span></button>`).join('');
  }

  document.querySelectorAll('#home [data-jump]').forEach(b => b.onclick = () => go(b.dataset.jump));
}

function row(group, text, meta, status) {
  return `<article class="focus-row ${status || ''}"><div><small>${group}</small><b>${text}</b></div><time>${meta || ''}</time></article>`;
}
const empty = text => `<p class="focus-empty">${text}</p>`;

function daysAway(date) {
  const d = Math.round((new Date(date) - new Date(new Date().toISOString().slice(0, 10))) / 86400000);
  if (d === 0) return '';
  return d > 0 ? `・還有 ${d} 天` : `・已結束 ${-d} 天`;
}

function currentSlot() {
  const now = new Date();
  const mins = now.getHours() * 60 + now.getMinutes();
  const toMin = s => { const [h, m] = String(s).split(':').map(Number); return h * 60 + (m || 0); };
  return flow.find(s => {
    const [a, b] = String(s.time).split('–');
    return b ? mins >= toMin(a) && mins < toMin(b) : false;
  }) || flow[0];
}

function setMeetings(names) {
  const strip = document.getElementById('meetingDates');
  if (!strip || !names.length) return;
  strip.innerHTML = names.map((m, i) => `<button class="${i === 0 ? 'active' : ''}">${m}</button>`).join('');
  activeMeeting = names[0];
  strip.querySelectorAll('button').forEach(b => b.onclick = () => {
    strip.querySelectorAll('button').forEach(x => x.classList.toggle('active', x === b));
    activeMeeting = b.textContent; renderTasks();
  });
}

function notice(text) {
  const host = document.getElementById('home') || document.body;
  const el = document.createElement('div');
  el.className = 'live-notice bad';
  el.textContent = text;
  host.prepend(el);
}

/* ── 載入 ─────────────────────────────────────────────────── */

function applyAll(d) {
  if (d.groups && d.groups.length) { groups = d.groups; localStorage.removeItem('five-groups'); }
  if (d.core && d.core.length) {
    directors.splice(0, directors.length, ...d.core);
    currentPerson = d.core[0][0];
    localStorage.setItem('five-current-person', currentPerson);
  }
  if (d.flow && d.flow.length) flow.splice(0, flow.length, ...d.flow);
  if (d.progress) { progressRecords.splice(0, progressRecords.length, ...d.progress); localStorage.removeItem('five-progress-v2'); }
  if (d.materials) { materials.splice(0, materials.length, ...d.materials); localStorage.removeItem('five-materials'); }
  if (d.finance) { budgets.splice(0, budgets.length, ...d.finance); localStorage.removeItem('five-budgets'); }

  applyEvent(d.event);
  if (d.meetings) setMeetings(d.meetings);

  // 期限的快選鈕原本寫死 5/17、5/18、5/21、活動當天，改成這場活動自己的日期
  if (typeof dueQuickDays !== 'undefined') {
    const chips = (d.meetingRows || [])
      .filter(m => m['日期'])
      .map(m => [String(m['會議名稱']), String(m['日期']).slice(0, 10)]);
    if (d.event && d.event.date) chips.push(['活動當天', d.event.date]);
    const seen = new Set();
    const uniq = chips.filter(([, v]) => !seen.has(v) && seen.add(v));
    if (uniq.length) dueQuickDays.splice(0, dueQuickDays.length, ...uniq);
  }
  renderPeople(); renderIdentity(); renderTasks(); renderFlow(); renderMaterials(); renderBudget();
  renderMaterialExtras();
  refillGroupSelects();
  if (d.pool) refillPersonSelect(d.pool);
  if (d.reviews) renderReviews(d.reviews);
  if (d.events) renderHistories(d.events, d.roster || [], d.event && d.event.id);
  renderHome();
}

function fromSheet(json, eventId) {
  const t = json.data;
  const all = t['活動'] || [];
  const ev = all.find(e => e['活動ID'] === eventId) || all[0] || {};
  const only = rows => rows.filter(r => !r['活動ID'] || r['活動ID'] === ev['活動ID']);
  const order = String(ev['工作組'] || '').split(',').map(s => s.trim()).filter(Boolean);
  const roster = only(t['活動人力'] || []);
  return {
    event: {
      id: ev['活動ID'], name: ev['活動名稱'], date: ev['日期'], place: ev['地點'],
      meta: [ev['日期'], ev['地點']].filter(Boolean).join('・'),
    },
    core: roster.filter(r => r['組別'] === '核心').map(r => [r['姓名'], r['角色']]),
    groups: rowsToGroups(roster, order),
    meetings: only(t['會議'] || []).map(m => m['會議名稱']).filter(Boolean),
    meetingRows: only(t['會議'] || []).filter(m => m['會議名稱']),
    flow: rowsToFlow(only(t['活動細流'] || [])),
    progress: only(t['工作進度'] || []).map(r => ({
      meeting: r['會議'], group: r['組別'], owner: r['負責人'], item: r['要完成事項'],
      update: r['目前進度'], next: r['下一步'], due: r['期限'], status: r['狀態'] || 'todo',
    })),
    materials: only(t['物資借用'] || []).map(r => {
      const back = String(r['已還']).toUpperCase() === 'TRUE';
      return [r['物資'], r['組別'], Number(r['數量']) || 0, r['攜帶者'],
              back ? '已歸還' : '待歸還', back ? 'done' : 'open', ''];  // 第 7 格是照片，不是備註
    }),
    finance: only(t['收支'] || []).map(r => [r['組別'], r['項目'], Number(String(r['金額']).replace(/[^0-9.-]/g, '')) || 0, r['登錄人'], r['備註'] || '']),
    reviews: only(t['會後精進'] || []),
    events: t['活動'] || [],
    pool: t['人員庫'] || [],
    roster,
  };
}

let _raw = null;

function renderPicker(events, roster) {
  const list = document.getElementById('pickerList');
  if (!list) return;
  if (!events.length) {
    list.innerHTML = '<div class="picker-empty">試算表的「活動」分頁還沒有任何一場活動。</div>';
    return;
  }
  const live = events.filter(e => e['狀態'] !== '已結案');
  const past = events.filter(e => e['狀態'] === '已結案');
  const card = (e, now) => {
    const mine = roster.filter(r => r['活動ID'] === e['活動ID']);
    const gs = new Set(mine.map(r => r['組別']).filter(g => g && g !== '核心'));
    const bits = [e['日期'], e['地點'], gs.size ? `${gs.size} 組・${mine.length} 人` : ''].filter(Boolean);
    return `<button class="picker-card ${now ? 'now' : ''}" data-event="${e['活動ID']}">
      <b>${e['活動名稱'] || '（未命名活動）'}</b><small>${bits.join('・')}</small></button>`;
  };
  list.innerHTML =
    (live.length ? `<p class="picker-group">進行中</p>` + live.map(e => card(e, true)).join('') : '') +
    (past.length ? `<p class="picker-group">已結案</p>` + past.map(e => card(e, false)).join('') : '');
  list.querySelectorAll('[data-event]').forEach(b => b.onclick = () => enterEvent(b.dataset.event));
}

function enterEvent(id) {
  if (!_raw) return;
  applyAll(fromSheet(_raw, id));
  document.body.classList.remove('picking');
  go('home');
  window.scrollTo({ top: 0 });
}

function backToPicker() {
  document.body.classList.add('picking');
  window.scrollTo({ top: 0 });
}

async function load() {
  document.body.classList.add('picking');
  try {
    _raw = await apiGet();
    sheetReady = true;
    renderPicker(_raw.data['活動'] || [], _raw.data['活動人力'] || []);
    const origProgress = saveProgress, origMaterials = saveMaterials;
    window.saveProgress = function () { origProgress(); apiReplace('工作進度', progressToRows()); };
    window.saveMaterials = function () { origMaterials(); renderMaterialExtras(); apiReplace('物資借用', materialsToRows()); };
    ['budgetForm', 'deleteBudgetForm'].forEach(id => {
      const f = document.getElementById(id);
      if (f) f.addEventListener('submit', () => setTimeout(() => apiReplace('收支', financeToRows()), 0));
    });
    const back = document.querySelector('[data-current-event]');
    if (back) back.onclick = () => { hideSideMenu(); backToPicker(); };
    return;
  } catch (err) {
    console.warn('[五表] 讀不到試算表端點：', err);
  }
  const list = document.getElementById('pickerList');
  if (list) list.innerHTML = '<div class="picker-empty">連不上試算表，請稍後重新整理。<br>（網路不穩時偶爾會這樣，重試三次都失敗才會看到這句）</div>';
}

load();
