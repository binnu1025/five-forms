/* 後台是 Google Sheet，中間隔一層 Apps Script。
   網頁只跟那個端點講話，從頭到尾碰不到試算表本身，
   所以 Sheet 可以只分享給會長與團長。

   在 app.js 之後載入，直接沿用它的 groups / flow / render* 等全域。
   端點掛掉就退回 repo 裡的 data/current.json，再不行才留在示範資料。 */

const API = 'https://script.google.com/macros/s/AKfycbwpJuvmtI_4jMG6NcEb_JtD15xg75Lzow_qQcA5yKu1gzBEWQ8fK4uDwLK3MtQzMBRDRg/exec';
const FALLBACK = 'data/current.json';

// 目前是誰在用。接上 Google 登入後這裡換成 id_token，identify_() 已經準備好了。
const whoAmI = () => localStorage.getItem('five-gmail') || 'binnu1025@gmail.com';

let sheetReady = false;

async function apiGet() {
  const res = await fetch(`${API}?who=${encodeURIComponent(whoAmI())}&t=${Date.now()}`);
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || '端點回報失敗');
  return json;
}

/* 整張表換掉，不用去追哪一列改了哪一欄 */
async function apiReplace(sheet, rows) {
  if (!sheetReady) return;
  try {
    const res = await fetch(API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ sheet, who: whoAmI(), replace: true, rows }),
    });
    const out = await res.json();
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
  借出照片: '', 已還: m[5] ? 'TRUE' : 'FALSE', 已歸位: '', 歸還照片: '', 備註: m[4] || '',
}));
const financeToRows = () => budgets.map(b => ({
  活動ID: window.__eventId || 'A001', 類型: '成本', 組別: b[0], 項目: b[1],
  金額: b[2], 登錄人: b[3], 備註: b[4] || '',
}));

/* ── 畫面上那些寫死的數字 ─────────────────────────────────── */

function applyEvent(ev) {
  if (!ev || !ev.name) return;
  window.__eventId = ev.id || 'A001';
  window.__eventMeta = ev.meta || '';
  window.__eventDate = ev.date || '';
  document.title = `${ev.name}｜活動協作`;
  const btn = document.getElementById('eventButtonName');
  if (btn) btn.textContent = ev.name;
  document.querySelectorAll('.side-current b, .hero h1').forEach(el => { el.textContent = ev.name; });
  if (ev.meta) document.querySelectorAll('.side-current span, .hero p').forEach(el => { el.textContent = ev.meta; });
}

function renderHome() {
  const people = directors.length + groups.reduce((n, g) => n + (g[1] ? 1 : 0) + g[2].length, 0);
  const done = progressRecords.filter(t => t.status === 'done').length;
  const doing = progressRecords.filter(t => t.status === 'doing').length;
  const openItems = materials.filter(m => !m[5]).length;
  const ready = progressRecords.length ? Math.round(done / progressRecords.length * 100) : null;

  const stats = document.querySelectorAll('.hero-stats span b');
  if (stats[0]) stats[0].textContent = groups.length;
  if (stats[1]) stats[1].textContent = people;
  if (stats[2]) stats[2].textContent = ready === null ? '—' : ready + '%';

  const hello = document.querySelector('.welcome p');
  if (hello) hello.textContent = `你好，${currentPerson}`;
  const roleChip = document.querySelector('.welcome span');
  if (roleChip) roleChip.textContent = currentRole() || '組員';

  const cards = document.querySelectorAll('.action-grid button');
  const first = flow[0];
  if (cards[0] && first) {
    cards[0].querySelector('b').textContent = String(first.time).split('–')[0];
    cards[0].querySelector('span').textContent = '活動開始';
    cards[0].querySelector('strong').textContent = first.name;
  }
  if (cards[1]) { cards[1].querySelector('b').textContent = openItems; cards[1].style.display = openItems ? '' : 'none'; }
  if (cards[2]) { cards[2].querySelector('b').textContent = doing; cards[2].style.display = doing ? '' : 'none'; }

  const badge = document.querySelector('.category-tabs [data-page="materials"] i');
  if (badge) { badge.textContent = openItems; badge.style.display = openItems ? '' : 'none'; }

  document.querySelectorAll('.home-section header b').forEach(b => { b.textContent = ready === null ? '—' : ready + '%'; });
  const bar = document.querySelector('.readiness i');
  if (bar) bar.style.width = (ready || 0) + '%';

  const list = document.querySelector('.check-list');
  if (list) list.innerHTML = [
    [`已建立 ${groups.length} 個工作組`, groups.length > 0],
    ['各組組長已指派', groups.every(g => g[1])],
    [progressRecords.length ? `還有 ${doing} 項任務進行中` : '尚未登錄任何工作進度', false],
    [openItems ? `還有 ${openItems} 件物資待處理` : '尚未登錄任何物資', false],
  ].map(([t, ok]) => `<li class="${ok ? 'done' : ''}">${t}</li>`).join('');

  const runDate = document.querySelector('#run .section-title p');
  if (runDate && window.__eventMeta) runDate.textContent = window.__eventMeta;
  const nowCard = document.querySelector('#run .now-card');
  if (nowCard) {
    const today = new Date().toISOString().slice(0, 10);
    if (window.__eventDate === today && flow[0]) {
      nowCard.querySelector('b').textContent = flow[0].name;
      nowCard.querySelector('small').textContent = flow[0].time;
    } else nowCard.style.display = 'none';
  }

  const mine = groups.filter(g => g[1] === currentPerson || g[2].includes(currentPerson)).map(g => g[0]);
  const myWorkEl = document.getElementById('myWork');
  if (myWorkEl) {
    const rows = flow.filter(s => mine.some(g => s.jobs && s.jobs[g]));
    myWorkEl.innerHTML = rows.length
      ? rows.map((s, i) => {
          const g = mine.find(x => s.jobs[x]);
          return `<article class="${i === 0 ? 'now' : ''}"><time>${s.time}</time><div><small>${g}</small><b>${s.jobs[g]}</b></div></article>`;
        }).join('')
      : `<article><div><small>${mine[0] || '尚未分組'}</small><b>細流裡還沒有你這一組的工作</b></div></article>`;
  }
}

/* 物資頁的提示與篩選數字原本寫死「4 件借出中／全部 6／借出中 4／已歸還 2」 */
function renderMaterialExtras() {
  const open = materials.filter(m => !m[5]).length;
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
  if (!histories.length && eventHistoryOptions) {
    eventHistoryOptions.innerHTML = '<p class="dialog-note">目前只有這一場活動，還沒有已結案的紀錄。</p>';
  }
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
  if (d.reviews) renderReviews(d.reviews);
  if (d.events) renderHistories(d.events, d.roster || [], d.event && d.event.id);
  renderHome();
}

function fromSheet(json) {
  const t = json.data;
  const ev = (t['活動'] || [])[0] || {};
  const order = String(ev['工作組'] || '').split(',').map(s => s.trim()).filter(Boolean);
  const roster = t['活動人力'] || [];
  return {
    event: {
      id: ev['活動ID'], name: ev['活動名稱'], date: ev['日期'], place: ev['地點'],
      meta: [ev['日期'], ev['地點']].filter(Boolean).join('・'),
    },
    core: roster.filter(r => r['組別'] === '核心').map(r => [r['姓名'], r['角色']]),
    groups: rowsToGroups(roster, order),
    meetings: (t['會議'] || []).map(m => m['會議名稱']).filter(Boolean),
    meetingRows: (t['會議'] || []).filter(m => m['會議名稱']),
    flow: rowsToFlow(t['活動細流'] || []),
    progress: (t['工作進度'] || []).map(r => ({
      meeting: r['會議'], group: r['組別'], owner: r['負責人'], item: r['要完成事項'],
      update: r['目前進度'], next: r['下一步'], due: r['期限'], status: r['狀態'] || 'todo',
    })),
    materials: (t['物資借用'] || []).map(r => [r['物資'], r['組別'], r['數量'], r['攜帶者'], r['備註'] || '', String(r['已還']).toUpperCase() === 'TRUE', '']),
    finance: (t['收支'] || []).map(r => [r['組別'], r['項目'], Number(String(r['金額']).replace(/[^0-9.-]/g, '')) || 0, r['登錄人'], r['備註'] || '']),
    reviews: t['會後精進'] || [],
    events: t['活動'] || [],
    roster,
  };
}

async function load() {
  try {
    const json = await apiGet();
    applyAll(fromSheet(json));
    sheetReady = true;
    // 寫入攔截：資料一變就整張推回 Sheet
    const origProgress = saveProgress, origMaterials = saveMaterials;
    window.saveProgress = function () { origProgress(); apiReplace('工作進度', progressToRows()); };
    window.saveMaterials = function () { origMaterials(); apiReplace('物資借用', materialsToRows()); };
    const budgetObserver = () => apiReplace('收支', financeToRows());
    ['budgetForm', 'deleteBudgetForm'].forEach(id => {
      const f = document.getElementById(id);
      if (f) f.addEventListener('submit', () => setTimeout(budgetObserver, 0));
    });
    return;
  } catch (err) {
    console.warn('[五表] 讀不到試算表端點：', err);
  }
  try {
    const res = await fetch(FALLBACK, { cache: 'no-store' });
    applyAll(await res.json());
    notice('連不上試算表，目前顯示的是 repo 裡的備份資料，改動不會寫回去。');
  } catch (err) {
    notice('連不上試算表，也讀不到備份資料，目前顯示的是示範資料。');
    console.warn('[五表] 備份也讀不到：', err);
  }
}

load();
