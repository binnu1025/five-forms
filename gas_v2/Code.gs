/**
 * 五表活動協作系統 — 唯讀端點（v2）
 *
 * 這一版只讀不寫。編輯都在 Google 試算表上做，網頁負責活動當天查看。
 * 所以沒有 doPost——以前那個「整張覆寫」把別場活動清掉的問題，
 * 在這個架構下不可能發生。
 *
 * 部署設定：
 *   執行身分：我（擁有者 cyp135.ticket@gmail.com）
 *   誰可以存取：任何人
 *
 * 「任何人」代表網址沒有門禁，所以這支腳本自己負責過濾：
 * 核心檔（收支）只回傳給核心幹部，人員庫的 Gmail 一律不外流。
 *
 * 用法：
 *   GET ?           → 活動清單＋我是誰
 *   GET ?event=A001 → 那一場的全部內容（依身分過濾）
 */

const PROP = PropertiesService.getScriptProperties();
const P_MASTER = 'MASTER_ID';

/** 看得到全部組別收支與總額的角色 */
const FULL_FINANCE_ROLES = ['會長', '青年團團長', '活動組長', '副組長', '財務長'];

/* ========== 讀取 ========== */

/** 一個檔案一次 batchGet 抓完所有分頁，不要一個分頁跑一趟。 */
function 讀整本_(id) {
  // 用 Sheets API 問分頁名稱就好，不要為了拿名字去 openById 整本打開——那很慢
  const meta = Sheets.Spreadsheets.get(id, { fields: 'sheets.properties.title' });
  const names = (meta.sheets || []).map(s => s.properties.title);
  const out = {};
  if (!names.length) return out;
  const res = Sheets.Spreadsheets.Values.batchGet(id, {
    ranges: names.map(n => "'" + n.replace(/'/g, "''") + "'"),
    valueRenderOption: 'FORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING',
  });
  (res.valueRanges || []).forEach((vr, i) => { out[names[i]] = vr.values || []; });
  return out;
}

/** 二維陣列 → 物件陣列，用第一列當欄位名。整列空白的跳過。 */
function 轉物件_(grid) {
  if (!grid || grid.length < 2) return [];
  const cols = grid[0].map(c => String(c).trim());
  return grid.slice(1)
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => {
      const o = {};
      cols.forEach((c, i) => { if (c) o[c] = r[i] === undefined ? '' : r[i]; });
      return o;
    });
}

/* ========== 身分 ========== */

/**
 * 認人。正式流程是前端送 Google 登入的 id_token 上來；
 * ?who= 是測試用的後門，正式上線前要拿掉。
 */
function 認身分_(params, master) {
  let email = '';
  if (params.id_token) {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(params.id_token),
      { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const info = JSON.parse(res.getContentText());
      if (String(info.email_verified) === 'true') email = String(info.email).toLowerCase();
    }
  }
  if (!email && params.who) email = String(params.who).toLowerCase();

  const 人員 = 轉物件_(master['人員庫']);
  const me = 人員.find(p => String(p['Gmail']).toLowerCase() === email);
  const roles = 轉物件_(master['權限'])
    .filter(r => String(r['Gmail']).toLowerCase() === email)
    .map(r => String(r['角色']));

  return {
    email: email,
    name: me ? String(me['姓名']) : '',
    roles: roles,
    known: !!me || roles.length > 0,
  };
}

/** 這個人在這一場活動裡是哪幾組、帶哪幾組。靠姓名比對，所以組員檔不需要存信箱。 */
function 這場的我_(who, 人力) {
  const mine = 人力.filter(r => String(r['姓名']).trim() === who.name);
  return {
    email: who.email,
    name: who.name,
    roles: who.roles.concat(mine.map(r => String(r['角色'])).filter(Boolean)),
    groups: mine.map(r => String(r['組別'])).filter(Boolean),
    leads: mine.filter(r => String(r['角色']) === '組長').map(r => String(r['組別'])),
    known: who.known,
  };
}

function 看得到錢_(me) {
  if (me.groups.indexOf('財務組') !== -1 || me.groups.indexOf('財務') !== -1) return true;
  return me.roles.some(r => FULL_FINANCE_ROLES.indexOf(r) !== -1);
}

/* ========== 整形 ========== */

/** 細流是寬表（一列一時段，一欄一組），轉成前端好用的形狀。 */
function 細流_(grid) {
  if (!grid || grid.length < 2) return [];
  const cols = grid[0].map(c => String(c).trim());
  const 固定 = ['開始', '結束', '全體流程'];
  const 組欄 = cols.map((c, i) => ({ 組別: c, i: i })).filter(x => x.組別 && 固定.indexOf(x.組別) === -1);

  return grid.slice(1)
    .filter(r => String(r[0] || '').trim() || String(r[2] || '').trim())
    .map(r => {
      const 工作 = {};
      組欄.forEach(x => {
        const v = String(r[x.i] === undefined ? '' : r[x.i]).trim();
        if (v) 工作[x.組別] = v;
      });
      return {
        開始: String(r[0] || ''), 結束: String(r[1] || ''),
        全體流程: String(r[2] || ''), 工作: 工作,
      };
    });
}

/** 每一組一個分頁，收成 { 組別: [工作, ...] }。 */
function 各組工作_(book, groups) {
  const out = {};
  groups.forEach(g => { if (book[g]) out[g] = 轉物件_(book[g]); });
  return out;
}

/** 組別順序以細流的欄位為準，沒有細流就看人力表出現的順序。 */
function 組別順序_(book) {
  const 細流格 = book['細流'];
  if (細流格 && 細流格.length) {
    const 固定 = ['開始', '結束', '全體流程'];
    const gs = 細流格[0].map(c => String(c).trim())
      .filter(c => c && 固定.indexOf(c) === -1);
    if (gs.length) return gs;
  }
  const seen = [];
  轉物件_(book['人力'] || []).forEach(r => {
    const g = String(r['組別'] || '').trim();
    if (g && seen.indexOf(g) === -1) seen.push(g);
  });
  return seen;
}

/* ========== 端點 ========== */

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const masterId = PROP.getProperty(P_MASTER);
    if (!masterId) return 回應_({ ok: false, error: '還沒建立總表，先執行 建立範本與總表()' });

    const master = 讀整本_(masterId);
    const who = 認身分_(params, master);
    const 活動 = 轉物件_(master['活動']);

    // 活動清單不外流檔案 ID——那是分享用的，網頁不需要知道
    const 清單 = 活動.map(a => ({
      活動ID: a['活動ID'], 活動名稱: a['活動名稱'], 日期: a['日期'],
      地點: a['地點'], 狀態: a['狀態'],
    }));

    if (!params.event) {
      return 回應_({ ok: true, me: who, events: 清單 });
    }

    const 這場 = 活動.find(a => String(a['活動ID']) === String(params.event));
    if (!這場) return 回應_({ ok: false, error: '找不到這場活動：' + params.event });

    const book = 讀整本_(String(這場['組員檔ID']));
    const 人力 = 轉物件_(book['人力'] || []);
    const me = 這場的我_(who, 人力);
    const groups = 組別順序_(book);

    const data = {
      活動: { 活動ID: 這場['活動ID'], 活動名稱: 這場['活動名稱'], 日期: 這場['日期'],
             地點: 這場['地點'], 狀態: 這場['狀態'], 組別: groups },
      會議: 轉物件_(book['會議'] || []),
      人力: 人力,
      細流: 細流_(book['細流']),
      物資: 轉物件_(book['物資'] || []),
      會後精進: 轉物件_(book['會後精進'] || []),
      工作: 各組工作_(book, groups),
      收支: null,
    };

    // 核心檔是另一個檔案，只有看得到錢的人才去讀它
    if (看得到錢_(me) && 這場['核心檔ID']) {
      const core = 讀整本_(String(這場['核心檔ID']));
      const rows = 轉物件_(core['收支'] || []);
      data.收支 = { 明細: rows, 總額: rows.reduce((s, r) => s + (Number(r['金額']) || 0), 0) };
    } else if (me.leads.length && 這場['核心檔ID']) {
      // 組長只看自己那組，不給總額
      const core = 讀整本_(String(這場['核心檔ID']));
      const rows = 轉物件_(core['收支'] || [])
        .filter(r => me.leads.indexOf(String(r['組別'])) !== -1);
      data.收支 = { 明細: rows, 總額: null };
    }

    return 回應_({ ok: true, me: me, events: 清單, data: data });
  } catch (err) {
    return 回應_({ ok: false, error: String(err) });
  }
}

function 回應_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
