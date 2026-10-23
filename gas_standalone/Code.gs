/**
 * 五表活動協作系統 — Sheet 中間層
 *
 * 貼進「五表後台」試算表的 Apps Script 編輯器，部署成網頁應用程式。
 * 網頁只跟這支腳本講話，從頭到尾碰不到試算表本身。
 *
 * 部署設定：
 *   執行身分：我（擁有者）        ← 這樣腳本才有權限讀你的 Sheet
 *   誰可以存取：任何人            ← 靜態網頁才呼叫得到
 *
 * 「任何人」代表網址本身沒有門禁，所以這支腳本自己負責過濾：
 * 人員庫、權限兩個分頁永遠不回傳，Gmail 欄位一律抹掉，收支依角色決定給不給。
 */

// 寫死試算表 ID，這樣獨立腳本（不是綁在試算表上的）也能用
const SHEET_ID = '1HY_snNwnirjY9DKZMtzA8RY1NF063nd25ymvIGmj8Bc';

// 網頁讀得到的分頁。人員庫與權限不在裡面，它們只在伺服器端用來判斷身分。
const READABLE = ['活動', '會議', '工作進度', '活動細流', '物資借用', '會後精進', '設定'];
// 網頁寫得進去的分頁，以及各自的欄位順序
const WRITABLE = {
  '工作進度': ['活動ID', '會議', '組別', '負責人', '要完成事項', '目前進度', '下一步', '期限', '狀態'],
  '物資借用': ['活動ID', '組別', '物資', '數量', '攜帶者', '借出照片', '已還', '已歸位', '歸還照片', '備註'],
  '收支':     ['活動ID', '類型', '組別', '項目', '金額', '登錄人', '備註'],
  '會後精進': ['活動ID', '組別', '問題', '影響', '下次做法', '記錄人'],
  // 下面這幾張只有核心角色能動（canEdit_ 會擋掉一般組員）
  '活動':     ['活動ID', '活動名稱', '日期', '地點', '狀態', '工作組'],
  '活動人力': ['活動ID', 'Gmail', '姓名', '組別', '角色'],
  '會議':     ['活動ID', '會議名稱', '日期'],
  '活動細流': ['活動ID', '開始', '結束', '全體流程', '組別', '該組工作'],
};
const ADMIN_ONLY = ['活動', '活動人力', '會議', '活動細流'];
// 這些角色看得到全部組別的收支與總額
const FULL_FINANCE_ROLES = ['會長', '青年團團長', '活動組長', '副組長', '財務長'];

function sheet_(name) {
  const s = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!s) throw new Error('找不到分頁：' + name);
  return s;
}

/* 原本每呼叫一次 rows_() 就往試算表跑一趟，一個請求要跑 13 趟（人員庫、
   活動人力還各讀兩次），這是 5 秒延遲的主因。改成一次 batchGet 把所有
   分頁抓回來，整個請求只跑一趟。

   _sheets 是這次執行的區域變數，每個請求都重新抓，不是快取。 */
let _sheets = null;

function allSheets_() {
  if (_sheets) return _sheets;
  _sheets = {};
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const names = ss.getSheets().map(sh => sh.getName());
  try {
    const res = Sheets.Spreadsheets.Values.batchGet(SHEET_ID, {
      ranges: names.map(n => "'" + n.replace(/'/g, "''") + "'"),
      valueRenderOption: 'FORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });
    (res.valueRanges || []).forEach((vr, i) => { _sheets[names[i]] = vr.values || []; });
  } catch (err) {
    // 進階服務不可用時退回逐張讀，慢但不會壞
    names.forEach(n => { _sheets[n] = ss.getSheetByName(n).getDataRange().getDisplayValues(); });
  }
  return _sheets;
}

function rows_(name) {
  const values = allSheets_()[name];
  if (!values) throw new Error('找不到分頁：' + name);
  if (values.length < 2) return [];
  const head = values[0].map(String);
  return values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null && c !== undefined))
    .map(r => {
      const o = {};
      head.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? '' : r[i]; });
      return o;
    });
}


/**
 * 認出呼叫者是誰。有 Google 登入就用 id_token（可信），
 * 還沒接登入時退回 who 參數（測試用，不可信，正式上線前要拿掉）。
 */
function identify_(params) {
  let email = '';
  if (params.id_token) {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(params.id_token),
      { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const info = JSON.parse(res.getContentText());
      if (info.email_verified === 'true' || info.email_verified === true) email = String(info.email).toLowerCase();
    }
  }
  if (!email && params.who) email = String(params.who).toLowerCase();  // 測試模式
  if (!email) return { email: '', name: '', roles: [], groups: [], known: false };

  const people = rows_('人員庫');
  const me = people.find(p => String(p['Gmail']).toLowerCase() === email);
  const roles = rows_('權限')
    .filter(r => String(r['Gmail']).toLowerCase() === email)
    .map(r => String(r['角色']));

  // 這個人在各活動裡擔任的角色與所屬工作組
  const assigns = rows_('活動人力').filter(a => String(a['Gmail']).toLowerCase() === email);
  assigns.forEach(a => { if (a['角色'] && roles.indexOf(String(a['角色'])) === -1) roles.push(String(a['角色'])); });

  return {
    email: email,
    name: me ? String(me['姓名']) : (assigns[0] ? String(assigns[0]['姓名']) : ''),
    roles: roles,
    groups: assigns.map(a => String(a['組別'])).filter(Boolean),
    leads: assigns.filter(a => String(a['角色']) === '組長').map(a => String(a['組別'])),
    known: !!me || assigns.length > 0,
  };
}

function canSeeAllFinance_(who) {
  // 財務組的人看得到全部，不管他在表上掛的是組長還是組員
  if (who.groups.indexOf('財務') !== -1 || who.groups.indexOf('財務組') !== -1) return true;
  return who.roles.some(r => FULL_FINANCE_ROLES.indexOf(r) !== -1);
}

/** 收支要在伺服器端就濾掉，不能整張送到手機再用前端隱藏 */
function filterFinance_(records, who) {
  if (canSeeAllFinance_(who)) return { rows: records, total: records.reduce((s, r) => s + (Number(r['金額']) || 0), 0) };
  if (who.leads.length) return { rows: records.filter(r => who.leads.indexOf(String(r['組別'])) !== -1), total: null };
  return { rows: [], total: null };
}

/** 活動人力要回傳，但 Gmail 不能外流 */
function roster_(who) {
  return rows_('活動人力').map(a => ({ 活動ID: a['活動ID'], 姓名: a['姓名'], 組別: a['組別'], 角色: a['角色'] }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    const who = identify_(p);
    if (!who.known) return json_({ ok: false, error: '不在名單中', me: { email: who.email } });

    const data = {};
    READABLE.forEach(name => { if (name !== '收支') data[name] = rows_(name); });
    data['活動人力'] = roster_(who);
    // 指派人員時要能從人才庫挑人，但信箱不能外流，所以只給姓名與雁群小組
    data['人員庫'] = rows_('人員庫')
      .filter(p => p['姓名'])
      .map(p => ({ 姓名: p['姓名'], 雁群小組: p['雁群小組'] }));
    const fin = filterFinance_(rows_('收支'), who);
    data['收支'] = fin.rows;

    return json_({
      ok: true,
      me: { email: who.email, name: who.name, roles: who.roles, groups: who.groups, leads: who.leads },
      canSeeFinanceTotal: canSeeAllFinance_(who),
      financeTotal: fin.total,
      data: data,
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * 新增或更新一列。
 * 前端用 text/plain 送 JSON，避免觸發 CORS 預檢（Apps Script 不處理 OPTIONS）。
 * body: { sheet, row: {...}, rowIndex?: number, who?, id_token? }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const who = identify_(body);
    if (!who.known) return json_({ ok: false, error: '不在名單中' });

    const name = body.sheet;
    const cols = WRITABLE[name];
    if (!cols) return json_({ ok: false, error: '這張表不開放寫入：' + name });

    if (ADMIN_ONLY.indexOf(name) !== -1 && !isAdmin_(who)) {
      return json_({ ok: false, error: '只有核心幹部能改「' + name + '」' });
    }

    const sh = sheet_(name);
    const incoming = body.rows || [body.row];   // rows：一次寫多列

    // replace：把舊資料整個換掉，不要一直往下疊
    if (body.replace) {
      const last = sh.getLastRow();
      if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
    }

    for (const row of incoming) {
      const group = String(row['組別'] || '');
      if (!body.replace && !canEdit_(who, group)) {
        return json_({ ok: false, error: '你沒有「' + group + '」的填寫權限' });
      }
      if (name === '收支' && !canSeeAllFinance_(who) && who.leads.indexOf(group) === -1) {
        return json_({ ok: false, error: '你不能登錄這一組的收支' });
      }
    }

    const lines = incoming.map(row => cols.map(c => row[c] !== undefined ? row[c] : ''));
    if (body.rowIndex && lines.length === 1) sh.getRange(body.rowIndex, 1, 1, cols.length).setValues([lines[0]]);
    else if (lines.length) sh.getRange(sh.getLastRow() + 1, 1, lines.length, cols.length).setValues(lines);

    return json_({ ok: true, written: lines.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function isAdmin_(who) {
  return who.roles.some(r => ['會長', '青年團團長', '活動組長', '副組長', '執行秘書', '台北團長', '新竹團長', '中壢團長'].indexOf(r) !== -1);
}

function canEdit_(who, group) {
  if (who.roles.some(r => ['會長', '青年團團長', '活動組長', '副組長'].indexOf(r) !== -1)) return true;
  return who.groups.indexOf(group) !== -1;
}
