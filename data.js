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
  // 新建立的活動還沒指派任何人，但「活動」分頁的工作組欄已經有清單，
  // 這些組要先長出來（組長空白），不然人力表會是空的、也沒地方指派。
  const names = order.length ? order : Object.keys(byGroup);
  Object.keys(byGroup).forEach(n => { if (!names.includes(n)) names.push(n); });
  return names.map(n => [n, (byGroup[n] || {}).lead || '', (byGroup[n] || {}).members || []]);
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
  活動ID: window.__eventId || 'A001', 會議: t.meeting || activeMeeting, 組別: t.group, 負責人: t.owner,
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
  const badge = document.getElementById('navMaterials');
  if (badge) { badge.textContent = open; badge.hidden = !open; }
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
  ['assignGroup', 'materialGroup', 'budgetGroup', 'flowGroup'].forEach(id => {
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

/* 前端該藏的東西。伺服器該擋的已經擋了，這裡是不要讓人看到做不到的按鈕。 */
function lockDownByRole() {
  const roles = (_raw && _raw.me && _raw.me.roles) || [];
  const admin = roles.some(r => ADMIN_ROLES.includes(r));

  // 組長只有會長／團長能換
  if (!admin) document.querySelectorAll('.change-leader').forEach(b => b.remove());

  // 組長只能調整自己那組：別組的「＋ 加入組員」拿掉
  if (!admin) {
    const leads = (_raw && _raw.me && _raw.me.leads) || [];
    document.querySelectorAll('[data-group]').forEach(b => {
      if (!leads.includes(b.dataset.group)) b.remove();
    });
    const addBtn = document.querySelector('#people .add');
    if (addBtn) addBtn.style.display = leads.length ? '' : 'none';
  }

  // 收支總額只有看得到全部的人能看；組長只看自己那組的明細，不該有總額
  const total = document.getElementById('budgetTotal');
  const totalBox = total && total.closest('.budget-total, .pledge-card, section, div');
  if (total && !(_raw && _raw.canSeeFinanceTotal)) {
    total.textContent = '—';
    const note = totalBox && totalBox.querySelector('small');
    if (note) note.textContent = '你看得到自己這一組的收支，總額僅核心幹部可見';
  }

  // 細流的「＋ 時段」是排流程，屬於幹部的事
  document.querySelectorAll('#run .add').forEach(b => { b.style.display = admin ? '' : 'none'; });

  // 細流：一次展開／收合全部時段，不用一個一個點
  const expandBtn = document.getElementById('expandAll');
  if (expandBtn) expandBtn.onclick = () => {
    const slots = [...document.querySelectorAll('#run .flow-slot')];
    const openAll = !slots.every(s => s.classList.contains('open'));
    slots.forEach(s => s.classList.toggle('open', openAll));
    expandBtn.textContent = openAll ? '全部收合' : '全部展開';
  };
  // 切換全場／我的／單組會重畫時段，展開狀態歸零，按鈕文字也要跟著回去
  document.querySelectorAll('#flowModes button').forEach(b => b.addEventListener('click', () => {
    if (expandBtn) expandBtn.textContent = '全部展開';
  }));
  const groupSel = document.getElementById('flowGroup');
  if (groupSel) groupSel.addEventListener('change', () => {
    if (expandBtn) expandBtn.textContent = '全部展開';
  });

  // 核心幹部不屬於任何工作組，「我的」對他們是死路，直接不顯示
  const myBtn = document.querySelector('#flowModes [data-mode="mine"]');
  if (myBtn) {
    const hasGroup = (window.__myGroups || []).length > 0;
    myBtn.style.display = hasGroup ? '' : 'none';
    if (!hasGroup && myBtn.classList.contains('active')) {
      const all = document.querySelector('#flowModes [data-mode="all"]');
      if (all) all.click();
    }
  }

  /* 細流頁的日期與「現在進行」卡原本寫死成 5 月 24 日、13:10 比賽＋輪打。
     這段在 v1.15 修過，v1.18 重寫 renderHome 時又被剪掉了。 */
  const runDate = document.querySelector('#run .section-title p');
  if (runDate && window.__eventMeta) runDate.textContent = window.__eventMeta;
  const nowCard = document.querySelector('#run .now-card');
  if (nowCard) {
    const today = new Date().toISOString().slice(0, 10);
    if (window.__eventDate === today && flow[0]) {
      nowCard.style.display = '';
      const b = nowCard.querySelector('b'), small = nowCard.querySelector('small');
      if (b) b.textContent = flow[0].name;
      if (small) small.textContent = flow[0].time;
    } else nowCard.style.display = 'none';   // 不是活動當天就別假裝有東西正在進行
  }
}

/* 工作進度改成一筆工作一列：不再按會議複製，「會議」欄變成「最後更新於」。
   選會議不是篩選，是換一個角度看——哪幾組這次還沒更新。 */
function renderTasksFlat() {
  const all = progressRecords;
  const shown = all.filter(t => progressFilter === 'all' || t.status === progressFilter);
  const counts = { doing: 0, done: 0, cancelled: 0 };
  all.forEach(t => { counts[t.status] = (counts[t.status] || 0) + 1; });

  const touched = new Set(all.filter(t => t.meeting === activeMeeting).map(t => t.group));
  const silent = groups.map(g => g[0]).filter(g => !touched.has(g));

  const summary = document.getElementById('meetingSummary');
  if (summary) summary.innerHTML =
    `<b>${activeMeeting || '本次會議'}</b><span>${all.length} 項工作</span>` +
    `<i>${counts.doing} 進行・${counts.done} 完成${counts.cancelled ? '・' + counts.cancelled + ' 取消' : ''}</i>` +
    (silent.length ? `<em class="silent">${silent.length} 組這次還沒更新：${silent.join('、')}</em>` : '');

  const list = document.getElementById('progressList');
  if (!list) return;
  list.innerHTML = shown.length ? shown.map(t => {
    const i = all.indexOf(t);
    const editable = canEditGroup(t.group);
    const stale = t.meeting !== activeMeeting;
    return `<article class="progress-row ${t.status === 'cancelled' ? 'is-cancelled' : ''} ${stale ? 'is-stale' : ''}" data-record="${i}">
      <div class="progress-main">
        <div class="progress-owner"><strong>${t.group}</strong><b>${t.owner || ''}</b></div>
        ${editable
          ? `<div class="progress-controls"><button class="change-leader" data-progress-edit="${i}">編輯</button><button class="progress-status ${t.status}" data-progress-status>${progressStates[t.status][0]}⌄</button></div>`
          : `<span class="progress-status ${t.status}">${progressStates[t.status][0]}</span>`}
        <h3>${t.item}</h3>
        <p><span>目前進度</span>${t.update || '（還沒填）'}</p>
        <p class="next-step"><span>下一步・${dueLabel(t.due)}</span>${t.next || ''}</p>
        ${stale ? `<p class="stale-note">最後更新於 ${t.meeting || '—'}，${activeMeeting} 還沒碰過</p>` : ''}
      </div></article>`;
  }).join('') : `<div class="progress-empty">這個篩選條件目前沒有工作</div>`;

  document.querySelectorAll('[data-progress-status]').forEach(b => b.onclick = () => {
    const rec = progressRecords[Number(b.closest('[data-record]').dataset.record)];
    const order = ['doing', 'done', 'cancelled'];
    rec.status = order[(order.indexOf(rec.status) + 1) % order.length];
    rec.meeting = activeMeeting;              // 動到就算這次更新過
    saveProgress(); renderTasks();
    toast(`${rec.group}已更新為${progressStates[rec.status][0]}`);
  });
  document.querySelectorAll('[data-progress-edit]').forEach(b =>
    b.onclick = () => openProgressEdit(Number(b.dataset.progressEdit)));
}

/* 狀態簡化成三種。試算表裡既有的 todo／blocked 一律當成進行中。 */
function normStatus(v) {
  const x = String(v || '').trim();
  if (x === 'done' || x === '已完成') return 'done';
  if (x === 'cancelled' || x === '取消') return 'cancelled';
  return 'doing';          // todo、blocked、空白、其他 → 進行中
}

function eventPhase() {
  const today = new Date().toISOString().slice(0, 10);
  if (!window.__eventDate) return 'before';
  if (window.__eventDate === today) return 'during';
  return window.__eventDate > today ? 'before' : 'after';
}

function renderHome() {
  document.querySelectorAll('#home .sk').forEach(el => el.classList.remove('sk', 'sk-s', 'sk-m', 'sk-l'));
  document.querySelectorAll('#home .sk-row').forEach(el => el.remove());

  const phase = eventPhase();
  const serverGroups = (_raw && _raw.me && _raw.me.groups) || [];
  const mine = serverGroups.filter(g => g && g !== '核心');
  const globalRoles = (_raw && _raw.me && _raw.me.roles) || [];
  const admin = globalRoles.some(r => ADMIN_ROLES.includes(r));

  /* 標題列 */
  const when = document.getElementById('todayWhen');
  if (when) when.textContent = {
    before: `${window.__eventMeta || ''}${daysAway(window.__eventDate)}`,
    during: '今天就是活動當天',
    after: `${window.__eventMeta || ''}${daysAway(window.__eventDate)}`,
  }[phase];
  const title = document.getElementById('todayTitle');
  if (title) title.textContent = window.__eventName || '';
  const who = document.getElementById('todayRole');
  if (who) who.textContent = currentPerson || '';

  /* ── 主區：依階段決定 ───────────────────────────── */
  const focusTitle = document.getElementById('focusTitle');
  const list = document.getElementById('focusList');
  const more = document.getElementById('focusMore');
  const visible = t => admin || mine.includes(t.group);
  const open = progressRecords.filter(t => t.status === 'doing' && visible(t));

  if (phase === 'during') {
    const slot = currentSlot();
    if (focusTitle) focusTitle.textContent = slot ? `現在・${slot.name}` : '活動流程';
    const jobs = slot ? Object.entries(slot.jobs || {}).filter(([g]) => admin || mine.includes(g)) : [];
    if (list) list.innerHTML = jobs.length
      ? jobs.map(([g, w]) => row(g, w, slot.time)).join('')
      : empty(mine.length ? '這個時段沒有你這一組的工作' : '細流裡還沒有時段');
    if (more) { more.hidden = false; more.textContent = '看全場細流'; more.onclick = () => go('run'); }

  } else if (phase === 'after') {
    const mineReview = reviews.filter(r => admin || mine.includes(r[0]));
    if (focusTitle) focusTitle.textContent = '活動結束了，記下下次可以更好的地方';
    if (list) list.innerHTML = mineReview.length
      ? mineReview.map(r => row(r[0], r[1], '')).join('')
      : empty(mine.length ? `${mine.join('、')}還沒寫會後精進` : '還沒有人寫會後精進');
    if (more) { more.hidden = false; more.textContent = '去寫'; more.onclick = () => go('review'); }

  } else {
    if (focusTitle) focusTitle.textContent = admin ? '還沒完成的工作'
      : mine.length ? `${mine.join('、')}還沒完成的` : '這場活動的工作';
    if (list) list.innerHTML = open.length
      ? open.slice(0, 4).map(t => row(t.group, t.item, dueLabel(t.due), t.status)).join('')
      : empty(mine.length || admin ? '目前沒有待辦，都完成了' : '你還沒有被指派到工作組');
    if (more) {
      more.hidden = open.length <= 4;
      more.textContent = `看全部 ${open.length} 項`;
      more.onclick = () => go('progress');
    }
  }

  /* ── 需要注意 ───────────────────────────────── */
  const today = new Date().toISOString().slice(0, 10);
  const alerts = [];
  progressRecords.filter(t => t.status === 'doing' && t.due && t.due < today && visible(t))
    .forEach(t => alerts.push(['已過期限', `${t.group}・${t.item}`, 'progress']));
  materials.filter(m => m[5] === 'open' && (admin || mine.includes(m[1])))
    .forEach(m => alerts.push(['尚未歸還', `${m[1]}・${m[0]} ${m[2]} 件`, 'materials']));
  const alertBlock = document.getElementById('attentionBlock');
  const alertList = document.getElementById('attentionList');
  if (alertBlock) alertBlock.hidden = !alerts.length;
  if (alertList) alertList.innerHTML = alerts.slice(0, 6).map(([tag, text, page]) =>
    `<button data-jump="${page}"><span>${tag}</span><b>${text}</b></button>`).join('');

  /* ── 各組進度：所有組都要出現，沒回報的才是重點 ───── */
  const teams = document.getElementById('teamsBlock');
  if (teams) {
    teams.hidden = !admin;
    if (admin) {
      const rows = groups.map(g => {
        const all = progressRecords.filter(t => t.group === g[0]);
        const done = all.filter(t => t.status === 'done').length;
        const fresh = all.some(t => t.meeting === activeMeeting);
        return { name: g[0], lead: g[1], total: all.length, done, fresh };
      });
      const silent = rows.filter(r => !r.total);
      const sub = document.getElementById('teamsSub');
      if (sub) sub.textContent = silent.length
        ? `${silent.length}／${groups.length} 組還沒回報`
        : `${rows.reduce((n, r) => n + r.done, 0)}／${progressRecords.length} 項完成`;

      const el = document.getElementById('teamsList');
      if (el) el.innerHTML = rows.map(r => {
        if (!r.total) return `<div class="team-row none"><b>${r.name}</b>
          <small>${r.lead ? r.lead + ' 尚未回報' : '還沒有組長'}</small></div>`;
        const pct = Math.round(r.done / r.total * 100);
        return `<div class="team-row ${pct === 100 ? 'is-done' : ''} ${r.fresh ? '' : 'stale'}"><b>${r.name}</b>
          <i><s style="width:${pct}%"></s></i>
          <small>${r.done}/${r.total}</small></div>`;
      }).join('');
    }
  }

  document.querySelectorAll('[data-jump]').forEach(b => b.onclick = () => {
    go(b.dataset.jump);
    window.scrollTo({ top: 0 });
  });
  const navBadge = document.getElementById('navMaterials');
  if (navBadge) {
    const openItems = materials.filter(m => m[5] === 'open').length;
    navBadge.textContent = openItems;
    navBadge.hidden = !openItems;
  }
  if (!window.__goWrapped) {
    const origGo = go;
    window.go = function (id) {
      origGo(id);
      document.querySelectorAll('.bottom-nav [data-jump]')
        .forEach(b => b.classList.toggle('active', b.dataset.jump === id));
    };
    window.__goWrapped = true;
  }
  const logo = document.querySelector('.logo');
  if (logo) logo.onclick = () => go('home');
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
  /* 一律替換，空的也要替換。
     之前寫成「有資料才換」，切到一場還沒排細流的新活動時，
     畫面會留著上一場（或 app.js 示範資料）的時段與幹部名單。 */
  groups = d.groups || [];
  localStorage.removeItem('five-groups');
  directors.splice(0, directors.length, ...(d.core || []));
  /* 目前是誰，要以伺服器認出的身分為準。
     之前寫成 d.core[0][0]（核心名單的第一個人），結果不管誰登入，
     畫面都把他當成會長——組長進來會看到幹部才該看的全場視圖。 */
  const meName = (_raw && _raw.me && _raw.me.name) || '';
  currentPerson = meName || (d.core && d.core.length ? d.core[0][0] : currentPerson);
  localStorage.setItem('five-current-person', currentPerson);
  flow.splice(0, flow.length, ...(d.flow || []));
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
  if (!window.__tasksPatched) { window.renderTasks = renderTasksFlat; window.__tasksPatched = true; }

  // 細流的「我的」檢視要知道我是哪幾組；原本 app.js 裡寫死成「環境組」
  window.__myGroups = ((_raw && _raw.me && _raw.me.groups) || []).filter(g => g && g !== '核心');

  renderPeople(); renderIdentity(); renderTasks(); renderFlow(); renderMaterials(); renderBudget();
  lockDownByRole();
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
      update: r['目前進度'], next: r['下一步'], due: r['期限'], status: normStatus(r['狀態']),
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

const ADMIN_ROLES = ['會長', '青年團團長', '活動組長', '副組長', '執行秘書', '台北團長', '新竹團長', '中壢團長'];
const isAdmin = () => (_raw && _raw.me.roles || []).some(r => ADMIN_ROLES.includes(r));

/* 組別一律用「組」結尾；自訂的也自動補上 */
const withGroupSuffix = n => {
  const s = String(n || '').trim();
  return (!s || s.endsWith('組')) ? s : s + '組';
};

/* 預設組別取自「設定」分頁；沒有就拿最近一場活動的工作組 */
function defaultGroups() {
  const fromSetting = (_raw.data['設定'] || []).find(r => r['項目'] === '預設組別');
  const list = String((fromSetting || {})['值'] || '').split(',').map(withGroupSuffix).filter(Boolean);
  if (list.length) return list;
  const last = (_raw.data['活動'] || []).slice(-1)[0] || {};
  return String(last['工作組'] || '').split(',').map(withGroupSuffix).filter(Boolean);
}

function nextEventId() {
  const ids = (_raw.data['活動'] || []).map(e => String(e['活動ID'] || ''));
  const nums = ids.map(x => /^A(\d+)$/.exec(x)).filter(Boolean).map(m => Number(m[1]));
  return 'A' + String((nums.length ? Math.max(...nums) : 0) + 1).padStart(3, '0');
}

let _picked = [];

function paintGroupPicks() {
  const box = document.getElementById('eventGroups');
  if (!box) return;
  box.innerHTML = _picked.map(([name, on]) =>
    `<button type="button" class="${on ? 'on' : ''}" data-g="${name}">${name}</button>`).join('');
  box.querySelectorAll('[data-g]').forEach(b => b.onclick = () => {
    const row = _picked.find(x => x[0] === b.dataset.g);
    row[1] = !row[1];
    paintGroupPicks();
  });
}

function addMeetingRow(name, date) {
  const box = document.getElementById('eventMeetings');
  if (!box) return;
  const row = document.createElement('div');
  row.className = 'meeting-row';
  row.innerHTML = `<input class="m-name" placeholder="會議名稱" value="${name || ''}">
    <input class="m-date" type="date" value="${date || ''}"><button type="button">×</button>`;
  row.querySelector('button').onclick = () => row.remove();
  box.appendChild(row);
}

function openEventDialog() {
  _picked = defaultGroups().map(n => [n, true]);
  paintGroupPicks();
  const box = document.getElementById('eventMeetings');
  if (box) box.innerHTML = '';
  addMeetingRow('會前會', '');
  eventForm.reset();
  paintGroupPicks();
  eventDialog.showModal();
}

async function createEvent(e) {
  e.preventDefault();
  const fd = new FormData(eventForm);
  const groupList = _picked.filter(x => x[1]).map(x => x[0]);
  if (!groupList.length) return toast('至少要留一個工作組');

  const id = nextEventId();
  const ev = {
    活動ID: id, 活動名稱: fd.get('name'), 日期: fd.get('date'),
    地點: fd.get('place') || '', 狀態: '規劃中', 工作組: groupList.join(','),
  };
  const meetings = [...document.querySelectorAll('#eventMeetings .meeting-row')]
    .map(r => ({ 活動ID: id, 會議名稱: r.querySelector('.m-name').value.trim(),
                 日期: r.querySelector('.m-date').value }))
    .filter(m => m['會議名稱']);

  const btn = document.getElementById('eventSubmitting') || eventForm.querySelector('.submit');
  btn.disabled = true; btn.textContent = '建立中…';
  try {
    const ok1 = await apiAppend('活動', [ev]);
    if (!ok1.ok) throw new Error(ok1.error);
    if (meetings.length) await apiAppend('會議', meetings);
    _raw = await apiGet();                       // 重新讀，讓選擇頁與新活動都是最新的
    renderPicker(_raw.data['活動'] || [], _raw.data['活動人力'] || []);
    eventDialog.close();
    enterEvent(id);
    toast(`${ev['活動名稱']} 已建立`);
  } catch (err) {
    toast('建立失敗：' + err.message);
    console.warn('[五表] 建立活動失敗：', err);
  } finally {
    btn.disabled = false; btn.textContent = '建立活動';
  }
}

/* 附加而不是整表覆蓋——建立活動不該動到既有的場次 */
async function apiAppend(sheet, rows) {
  return withRetry(async () => {
    const res = await fetch(API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ sheet, who: whoAmI(), rows }),
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    return res.json();
  });
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
    // 指派人員：原本只存在瀏覽器裡，現在要寫回試算表
    if (typeof assignForm !== 'undefined' && assignForm) {
      const origAssign = assignForm.onsubmit;
      assignForm.onsubmit = async e => {
        const fd = new FormData(assignForm);
        const row = {
          活動ID: window.__eventId || 'A001', Gmail: '',   // 信箱由伺服器從人員庫補
          姓名: fd.get('name'), 組別: fd.get('group'), 角色: fd.get('role'),
        };
        if (origAssign) origAssign.call(assignForm, e);
        const out = await apiAppend('活動人力', [row]);
        if (!out.ok) { toast(out.error); return; }
        _raw = await apiGet();
        applyAll(fromSheet(_raw, window.__eventId));
      };
    }

    const origProgress = saveProgress, origMaterials = saveMaterials;
    window.saveProgress = function () { origProgress(); apiReplace('工作進度', progressToRows()); };
    window.saveMaterials = function () { origMaterials(); renderMaterialExtras(); apiReplace('物資借用', materialsToRows()); };
    ['budgetForm', 'deleteBudgetForm'].forEach(id => {
      const f = document.getElementById(id);
      if (f) f.addEventListener('submit', () => setTimeout(() => apiReplace('收支', financeToRows()), 0));
    });
    const back = document.querySelector('[data-current-event]');
    if (back) back.onclick = () => { hideSideMenu(); backToPicker(); };
    const newBtn = document.getElementById('newEventBtn');
    if (newBtn) { newBtn.hidden = !isAdmin(); newBtn.onclick = openEventDialog; }
    if (typeof eventForm !== 'undefined' && eventForm) eventForm.onsubmit = createEvent;
    const addG = document.getElementById('addGroupBtn');
    if (addG) addG.onclick = () => {
      const inp = document.getElementById('newGroupName');
      const name = withGroupSuffix(inp.value);
      if (!name) return;
      if (!_picked.some(x => x[0] === name)) _picked.push([name, true]);
      inp.value = '';
      paintGroupPicks();
    };
    const addM = document.getElementById('addMeetingBtn');
    if (addM) addM.onclick = () => addMeetingRow('', '');
    return;
  } catch (err) {
    console.warn('[五表] 讀不到試算表端點：', err);
  }
  const list = document.getElementById('pickerList');
  if (list) list.innerHTML = '<div class="picker-empty">連不上試算表，請稍後重新整理。<br>（網路不穩時偶爾會這樣，重試三次都失敗才會看到這句）</div>';
}

load();
