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

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// 網頁讀得到的分頁。人員庫與權限不在裡面，它們只在伺服器端用來判斷身分。
const READABLE = ['活動', '會議', '工作進度', '活動細流', '物資借用', '會後精進', '設定'];
// 網頁寫得進去的分頁，以及各自的欄位順序
const WRITABLE = {
  '工作進度': ['活動ID', '會議', '組別', '負責人', '要完成事項', '目前進度', '下一步', '期限', '狀態'],
  '物資借用': ['活動ID', '組別', '物資', '數量', '攜帶者', '借出照片', '已還', '已歸位', '歸還照片', '備註'],
  '收支':     ['活動ID', '類型', '組別', '項目', '金額', '登錄人', '備註'],
  '會後精進': ['活動ID', '組別', '問題', '影響', '下次做法', '記錄人'],
};
// 這些角色看得到全部組別的收支與總額
const FULL_FINANCE_ROLES = ['會長', '青年團團長', '活動組長', '副組長', '財務長'];

function sheet_(name) {
  const s = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!s) throw new Error('找不到分頁：' + name);
  return s;
}

function rows_(name) {
  const values = sheet_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(String);
  return values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => {
      const o = {};
      head.forEach((h, i) => { if (h) o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i], 'Asia/Taipei', 'yyyy-MM-dd') : r[i]; });
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

    const group = String(body.row['組別'] || '');
    if (!canEdit_(who, group)) return json_({ ok: false, error: '你沒有「' + group + '」的填寫權限' });
    if (name === '收支' && !canSeeAllFinance_(who) && who.leads.indexOf(group) === -1) {
      return json_({ ok: false, error: '你不能登錄這一組的收支' });
    }

    const sh = sheet_(name);
    const line = cols.map(c => body.row[c] !== undefined ? body.row[c] : '');
    if (body.rowIndex) sh.getRange(body.rowIndex, 1, 1, cols.length).setValues([line]);
    else sh.appendRow(line);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function canEdit_(who, group) {
  if (who.roles.some(r => ['會長', '青年團團長', '活動組長', '副組長'].indexOf(r) !== -1)) return true;
  return who.groups.indexOf(group) !== -1;
}
