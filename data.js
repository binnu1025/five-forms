/* 後台在這個 repo 裡：data/current.json。
   改那個檔、commit、push，網站就更新——不需要 Google、不需要任何授權。

   在 app.js 之後載入，所以直接沿用它的 groups / flow / directors / render* 等全域。
   讀不到就安靜留在示範資料，不讓頁面壞掉。 */

const DATA_URL = 'data/current.json?v=1.12';

function applyEvent(ev) {
  if (!ev || !ev.name) return;
  window.__eventMeta = ev.meta || '';
  window.__eventDate = ev.date || '';
  document.title = `${ev.name}｜活動協作`;
  const btn = document.getElementById('eventButtonName');
  if (btn) btn.textContent = ev.name;
  document.querySelectorAll('.side-current b').forEach(el => { el.textContent = ev.name; });
  document.querySelectorAll('.hero h1').forEach(el => { el.textContent = ev.name; });
  if (ev.meta) {
    document.querySelectorAll('.side-current span').forEach(el => { el.textContent = ev.meta; });
    document.querySelectorAll('.hero p').forEach(el => { el.textContent = ev.meta; });
  }
}

/* 首頁原本整片都是寫死的數字（10 工作組、48 位成員、68% 準備度、
   13:10 比賽與輪打…）。全部改成從資料算出來，沒有資料就不要假裝有。 */
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

  // 三張捷徑卡：沒有數字的就別擺一個 0 在那裡
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

  const readyNum = document.querySelectorAll('.home-section header b');
  readyNum.forEach(b => { b.textContent = ready === null ? '—' : ready + '%'; });
  const bar = document.querySelector('.readiness i');
  if (bar) bar.style.width = (ready || 0) + '%';

  const list = document.querySelector('.check-list');
  if (list) list.innerHTML = [
    [`已建立 ${groups.length} 個工作組`, groups.length > 0],
    [`各組組長已指派`, groups.every(g => g[1])],
    [progressRecords.length ? `還有 ${doing} 項任務進行中` : '尚未登錄任何工作進度', false],
    [openItems ? `還有 ${openItems} 件物資待處理` : '尚未登錄任何物資', false],
  ].map(([t, ok]) => `<li class="${ok ? 'done' : ''}">${t}</li>`).join('');

  // 細流頁的日期與「現在進行」卡也是寫死的（5 月 24 日、13:10 比賽＋輪打）
  const runDate = document.querySelector('#run .section-title p');
  if (runDate && window.__eventMeta) runDate.textContent = window.__eventMeta;
  const nowCard = document.querySelector('#run .now-card');
  if (nowCard) {
    const today = new Date().toISOString().slice(0, 10);
    if (window.__eventDate && window.__eventDate === today && flow[0]) {
      nowCard.querySelector('b').textContent = flow[0].name;
      nowCard.querySelector('small').textContent = flow[0].time;
    } else nowCard.style.display = 'none';   // 不是活動當天就別假裝有東西正在進行
  }

  // 「我的現場工作」原本也是寫死的，改成從細流撈出我這一組的時段
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

function liveNotice(text) {
  const host = document.getElementById('home') || document.body;
  const el = document.createElement('div');
  el.className = 'live-notice bad';
  el.textContent = text;
  host.prepend(el);
}

async function loadData() {
  let d;
  try {
    const res = await fetch(DATA_URL, { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    d = await res.json();
  } catch (err) {
    liveNotice('讀不到 data/current.json，目前顯示的是示範資料。');
    console.warn('[五表] 讀不到後台資料：', err);
    return;
  }

  if (Array.isArray(d.groups) && d.groups.length) {
    groups = d.groups;
    localStorage.removeItem('five-groups');      // 示範資料的殘留會蓋掉真實名單
  }
  if (Array.isArray(d.core) && d.core.length) {
    directors.splice(0, directors.length, ...d.core);
    currentPerson = d.core[0][0];
    localStorage.setItem('five-current-person', currentPerson);
  }
  if (Array.isArray(d.flow) && d.flow.length) flow.splice(0, flow.length, ...d.flow);

  // 討論紀錄的日期原本寫死 5/17、5/18、5/21、會前會
  if (Array.isArray(d.meetings) && d.meetings.length) {
    const strip = document.getElementById('meetingDates');
    if (strip) {
      strip.innerHTML = d.meetings.map((m, i) => `<button class="${i === 0 ? 'active' : ''}">${m}</button>`).join('');
      activeMeeting = d.meetings[0];
      strip.querySelectorAll('button').forEach(b => b.onclick = () => {
        strip.querySelectorAll('button').forEach(x => x.classList.toggle('active', x === b));
        activeMeeting = b.textContent; renderTasks();
      });
    }
  }

  if (Array.isArray(d.progress)) {
    progressRecords.splice(0, progressRecords.length, ...d.progress);
    localStorage.removeItem('five-progress-v2');
  }
  if (Array.isArray(d.materials)) {
    materials.splice(0, materials.length, ...d.materials);
    localStorage.removeItem('five-materials');
  }
  if (Array.isArray(d.finance)) {
    budgets.splice(0, budgets.length, ...d.finance);
    localStorage.removeItem('five-budgets');
  }

  applyEvent(d.event);
  renderPeople(); renderIdentity(); renderTasks(); renderFlow(); renderMaterials(); renderBudget();
  renderHome();
}

loadData();
