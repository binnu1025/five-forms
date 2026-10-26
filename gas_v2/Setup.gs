/**
 * 五表活動協作系統 — 建置與開新活動
 *
 * 架構（2026-10-26 改）：編輯都在 Google 試算表，網頁只負責活動當天查看。
 *
 *   總表（一份，永久）         活動清單、人員庫、權限
 *     ├── 範本｜組員            複製給每一場活動，分享給全部工作人員
 *     └── 範本｜核心            複製給每一場活動，只分享給核心幹部
 *
 * 一場活動 = 兩個檔案。金額只能靠拆檔案隔離，因為試算表的保護範圍
 * 只能擋編輯、擋不住查看。
 *
 * 檔案 ID 存在指令碼屬性裡，不寫死在程式碼（這個 repo 是公開的）。
 *
 * 使用順序：
 *   1. 執行一次 建立範本與總表()
 *   2. 之後每場活動執行 開新活動('2026.01 新春團拜', '2026-01-18', '北三教室')
 */

// PROP 與 P_MASTER 宣告在 Code.gs——Apps Script 的 .gs 共用同一個全域範圍，
// 在這裡再宣告一次會變成重複宣告的錯誤。
const P_TPL_TEAM = 'TEMPLATE_TEAM_ID';
const P_TPL_CORE = 'TEMPLATE_CORE_ID';

/** 公版組別。開新活動時可以傳別的進來覆蓋。 */
const 預設組別 = ['主持組', '報到組', '隊輔組', '計分組', '攝影組', '音控組', '環境組',
                 '交通組', '秩序組', '活動組', '財務組', '機動組', '社群組'];
const 狀態清單 = ['進行中', '已完成', '取消'];
const 角色清單 = ['組長', '組員'];

/* ========== 分頁定義 ========== */

const 總表分頁 = {
  '活動':   ['活動ID', '活動名稱', '日期', '地點', '狀態', '組員檔ID', '核心檔ID'],
  '人員庫': ['姓名', '雁群小組', 'Gmail'],
  '權限':   ['Gmail', '角色'],
};

const 組員檔分頁 = {
  '細流':     ['開始', '結束', '全體流程'],        // 每一組會各加一欄
  '人力':     ['組別', '角色', '姓名'],
  '物資':     ['組別', '物資', '數量', '攜帶者', '已還', '借出照片', '歸還照片', '備註'],
  '會議':     ['會議名稱', '日期'],
  '會後精進': ['組別', '下次可以更好的地方', '記錄人'],
  '清單':     ['組別', '狀態', '角色', '人員'],    // 給下拉選單用，會隱藏
};

/** 每一組的工作分頁長這樣。開新活動時依組別數量複製。 */
const 組工作欄位 = ['會議', '要完成事項', '負責人', '目前進度', '下一步', '期限', '狀態'];
const 組範本分頁名 = '＿組範本';

const 核心檔分頁 = {
  '收支': ['類型', '組別', '項目', '金額', '登錄人', '備註'],
  '清單': ['組別', '類型'],
};

/* ========== 建置 ========== */

/** 跑一次就好。建立總表與兩份範本，把 ID 記進指令碼屬性。 */
function 建立範本與總表() {
  const 已有 = PROP.getProperty(P_MASTER);
  if (已有) {
    throw new Error('總表已經建過了（' + 已有 + '）。要重建請先執行 清掉記住的檔案ID()。');
  }

  const master = SpreadsheetApp.create('五表｜總表');
  建分頁_(master, 總表分頁);
  master.getSheetByName('人員庫').getRange('A2:A').setNumberFormat('@');  // 姓名當文字，避免被當日期
  寫入_(master.getSheetByName('權限'), [['', '']]);                       // 留一列空的當範例

  const team = SpreadsheetApp.create('五表｜範本｜組員');
  建分頁_(team, 組員檔分頁);
  填清單_(team.getSheetByName('清單'));
  建組範本_(team);
  設組員檔驗證_(team);
  team.getSheetByName('清單').hideSheet();

  const core = SpreadsheetApp.create('五表｜範本｜核心');
  建分頁_(core, 核心檔分頁);
  const coreList = core.getSheetByName('清單');
  寫入_(coreList, 預設組別.map((g, i) => [g, ['成本', '收入'][i] || '']));
  設核心檔驗證_(core);
  coreList.hideSheet();

  PROP.setProperties({
    [P_MASTER]: master.getId(),
    [P_TPL_TEAM]: team.getId(),
    [P_TPL_CORE]: core.getId(),
  });

  const 訊息 = [
    '建好了，三個檔案都在這個帳號的雲端硬碟根目錄：',
    '  總表　　　 ' + master.getUrl(),
    '  範本｜組員 ' + team.getUrl(),
    '  範本｜核心 ' + core.getUrl(),
    '',
    '接下來：把人員庫填一填，然後執行 開新活動(...)。',
  ].join('\n');
  Logger.log(訊息);
  return 訊息;
}

/** 重建用。只是忘掉 ID，不會刪掉雲端硬碟上的檔案。 */
function 清掉記住的檔案ID() {
  PROP.deleteProperty(P_MASTER);
  PROP.deleteProperty(P_TPL_TEAM);
  PROP.deleteProperty(P_TPL_CORE);
  return '忘掉了。舊檔案還在雲端硬碟，要的話自己刪。';
}

function 建分頁_(ss, 定義) {
  Object.keys(定義).forEach((name, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(name) : ss.insertSheet(name);
    標題列_(sh, 定義[name]);
  });
}

/** 欄寬。autoResize 只會縮到剛好塞下標題那兩三個字，實際要打字完全不夠用。 */
const 欄寬_ = {
  '開始': 90, '結束': 90, '全體流程': 220,
  '組別': 110, '角色': 90, '姓名': 110,
  '物資': 180, '數量': 70, '攜帶者': 110, '已還': 80,
  '借出照片': 200, '歸還照片': 200, '備註': 240,
  '會議名稱': 160, '日期': 110,
  '下次可以更好的地方': 360, '記錄人': 110,
  '會議': 120, '要完成事項': 260, '負責人': 110,
  '目前進度': 300, '下一步': 240, '期限': 110, '狀態': 100,
  '類型': 90, '項目': 220, '金額': 110, '登錄人': 110,
  '活動ID': 90, '活動名稱': 240, '地點': 140, '組員檔ID': 130, '核心檔ID': 130,
  'Gmail': 220, '雁群小組': 130, '人員': 120, '值': 200, '說明': 260,
};
const 預設欄寬 = 140;   // 細流的各組欄位、以及任何沒列到的欄位

/** 依標題文字決定每一欄多寬；太寬的欄位改成自動換行，不然整列會被撐很高。 */
function 套欄寬_(sh) {
  const n = sh.getLastColumn();
  if (!n) return;
  const 標題 = sh.getRange(1, 1, 1, n).getDisplayValues()[0];
  標題.forEach((c, i) => {
    const w = 欄寬_[String(c).trim()] || 預設欄寬;
    sh.setColumnWidth(i + 1, w);
    const 策略 = w >= 200 ? SpreadsheetApp.WrapStrategy.WRAP : SpreadsheetApp.WrapStrategy.CLIP;
    sh.getRange(1, i + 1, sh.getMaxRows(), 1).setWrapStrategy(策略);
  });
  sh.getRange(1, 1, 1, n).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);  // 標題列不換行
}

/** 標題列：填字、粗體、凍結、保護起來不讓人改欄位名稱。 */
function 標題列_(sh, cols) {
  sh.getRange(1, 1, 1, cols.length).setValues([cols])
    .setFontWeight('bold').setBackground('#2f4f3f').setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, cols.length).protect()
    .setDescription('欄位名稱是程式讀取的依據，不要改')
    .setWarningOnly(true);
  if (sh.getMaxColumns() > cols.length) {
    sh.deleteColumns(cols.length + 1, sh.getMaxColumns() - cols.length);
  }
  套欄寬_(sh);
}

function 寫入_(sh, rows) {
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function 填清單_(sh) {
  const n = Math.max(預設組別.length, 狀態清單.length, 角色清單.length);
  const rows = [];
  for (let i = 0; i < n; i++) {
    rows.push([預設組別[i] || '', 狀態清單[i] || '', 角色清單[i] || '', '']);
  }
  寫入_(sh, rows);
}

/** 隱藏的原型分頁。開新活動時一組複製一份。 */
function 建組範本_(ss) {
  const sh = ss.insertSheet(組範本分頁名);
  標題列_(sh, 組工作欄位);
  const 清單 = ss.getSheetByName('清單');
  驗證_(sh, 'G', 清單.getRange('B2:B'));                       // 狀態
  驗證_(sh, 'C', 清單.getRange('D2:D'));                       // 負責人
  驗證_(sh, 'A', ss.getSheetByName('會議').getRange('A2:A'));  // 會議
  sh.getRange('F2:F').setNumberFormat('yyyy-mm-dd');           // 期限
  sh.hideSheet();
}

function 設組員檔驗證_(ss) {
  const 清單 = ss.getSheetByName('清單');
  const 物資 = ss.getSheetByName('物資');
  驗證_(物資, 'A', 清單.getRange('A2:A'));
  物資.getRange('E2:E').insertCheckboxes();                    // 已還
  const 人力 = ss.getSheetByName('人力');
  驗證_(人力, 'A', 清單.getRange('A2:A'));
  驗證_(人力, 'B', 清單.getRange('C2:C'));
  驗證_(人力, 'C', 清單.getRange('D2:D'));
  驗證_(ss.getSheetByName('會後精進'), 'A', 清單.getRange('A2:A'));
  ss.getSheetByName('會議').getRange('B2:B').setNumberFormat('yyyy-mm-dd');
}

function 設核心檔驗證_(ss) {
  const 清單 = ss.getSheetByName('清單');
  const 收支 = ss.getSheetByName('收支');
  驗證_(收支, 'A', 清單.getRange('B2:B'));                      // 類型
  驗證_(收支, 'B', 清單.getRange('A2:A'));                      // 組別
  收支.getRange('D2:D').setNumberFormat('#,##0');               // 金額
}

/** 整欄（第 2 列以下）套下拉選單。 */
function 驗證_(sh, 欄, 來源) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(來源, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(欄 + '2:' + 欄).setDataValidation(rule);
}

/* ========== 開新活動 ========== */

/**
 * 複製兩份範本，建好每一組的分頁，寫進總表。
 *
 *   開新活動('2026.01 新春團拜', '2026-01-18', '北三教室')
 *   開新活動('...', '...', '...', ['主持組', '報到組'])   ← 只要這兩組
 */
function 開新活動(名稱, 日期, 地點, 組別) {
  if (!名稱 || !日期) throw new Error('名稱跟日期是必填的');
  const groups = (組別 && 組別.length) ? 組別 : 預設組別;

  const masterId = 必須有_(P_MASTER, '總表');
  const master = SpreadsheetApp.openById(masterId);
  const 活動表 = master.getSheetByName('活動');
  const id = 下一個活動ID_(活動表);

  const team = 複製範本_(P_TPL_TEAM, 名稱 + '｜組員');
  const core = 複製範本_(P_TPL_CORE, 名稱 + '｜核心');

  建每組分頁_(team, groups);
  加細流組欄_(team, groups);
  寫組別清單_(team.getSheetByName('清單'), groups);
  寫組別清單_(core.getSheetByName('清單'), groups);
  帶入人員_(team, master);

  活動表.appendRow([id, 名稱, 日期, 地點 || '', '規劃中', team.getId(), core.getId()]);

  const 訊息 = [
    名稱 + ' 建好了（' + id + '）',
    '  組員檔（分享給全部工作人員）' + team.getUrl(),
    '  核心檔（只分享給核心幹部）　' + core.getUrl(),
    '',
    '分享是手動的——程式不自動加人，免得加錯。',
  ].join('\n');
  Logger.log(訊息);
  return 訊息;
}

function 必須有_(key, 名字) {
  const v = PROP.getProperty(key);
  if (!v) throw new Error('還沒建立' + 名字 + '，先執行 建立範本與總表()');
  return v;
}

function 複製範本_(key, 檔名) {
  return SpreadsheetApp.open(DriveApp.getFileById(必須有_(key, '範本')).makeCopy(檔名));
}

function 下一個活動ID_(sh) {
  const ids = sh.getLastRow() < 2 ? []
    : sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(r => String(r[0]));
  const nums = ids.map(x => /^A(\d+)$/.exec(x)).filter(Boolean).map(m => Number(m[1]));
  return 'A' + String((nums.length ? Math.max.apply(null, nums) : 0) + 1).padStart(3, '0');
}

/** 一組一個分頁，從隱藏的原型複製，所以下拉選單跟著帶過去。 */
function 建每組分頁_(ss, groups) {
  const proto = ss.getSheetByName(組範本分頁名);
  groups.forEach((g, i) => {
    const sh = proto.copyTo(ss).setName(g);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
    sh.showSheet();
  });
  ss.deleteSheet(proto);
}

/** 細流是寬表：一列一個時段，一欄一組。當天看板就長這樣。 */
function 加細流組欄_(ss, groups) {
  const sh = ss.getSheetByName('細流');
  const 起 = sh.getLastColumn() + 1;
  sh.insertColumnsAfter(sh.getLastColumn(), groups.length);
  sh.getRange(1, 起, 1, groups.length).setValues([groups])
    .setFontWeight('bold').setBackground('#2f4f3f').setFontColor('#ffffff');
  sh.getRange('A2:B').setNumberFormat('HH:mm');
  套欄寬_(sh);
}

function 寫組別清單_(sh, groups) {
  sh.getRange('A2:A').clearContent();
  sh.getRange(2, 1, groups.length, 1).setValues(groups.map(g => [g]));
}

/** 人員庫的姓名帶進組員檔的清單，讓「負責人」有得選。信箱不跟著過去。 */
function 帶入人員_(team, master) {
  const 人員 = master.getSheetByName('人員庫');
  if (人員.getLastRow() < 2) return;
  const names = 人員.getRange(2, 1, 人員.getLastRow() - 1, 1).getDisplayValues()
    .map(r => String(r[0]).trim()).filter(Boolean);
  if (!names.length) return;
  const 清單 = team.getSheetByName('清單');
  清單.getRange(2, 4, names.length, 1).setValues(names.map(n => [n]));
}

/* ========== 總表上的選單 ========== */

/**
 * 跑一次就好。在總表裝一個選單，以後開新活動直接在試算表上點，
 * 不用再回到 Apps Script 編輯器。
 */
function 安裝選單() {
  const id = 必須有_(P_MASTER, '總表');
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '選單_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('選單_').forSpreadsheet(id).onOpen().create();
  return '裝好了。重新整理總表就會看到「五表」選單。';
}

/**
 * 範本已經建好之後才改欄寬的話跑這個。順便把選單裝起來，
 * 這樣只要在編輯器按一次就好。
 */
function 完成設定() {
  const 修 = [];
  [P_TPL_TEAM, P_TPL_CORE].forEach(key => {
    const id = PROP.getProperty(key);
    if (!id) return;
    const ss = SpreadsheetApp.openById(id);
    ss.getSheets().forEach(sh => 套欄寬_(sh));
    修.push(ss.getName());
  });
  const master = PROP.getProperty(P_MASTER);
  if (master) {
    const ss = SpreadsheetApp.openById(master);
    ss.getSheets().forEach(sh => 套欄寬_(sh));
    修.push(ss.getName());
  }
  return [安裝選單(), '欄寬已整理：' + 修.join('、')].join(String.fromCharCode(10));
}

function 選單_() {
  SpreadsheetApp.getUi().createMenu('五表')
    .addItem('開新活動…', '開新活動對話框_')
    .addItem('目前設定', '目前設定_')
    .addToUi();
}

function 開新活動對話框_() {
  const ui = SpreadsheetApp.getUi();
  const 問 = (題目, 說明) => {
    const r = ui.prompt('開新活動', 題目 + (說明 ? '\n' + 說明 : ''), ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return null;
    return r.getResponseText().trim();
  };

  const 名稱 = 問('活動名稱', '例如：2026.01 新春團拜');
  if (名稱 === null) return;
  if (!名稱) return ui.alert('活動名稱不能空白');

  const 日期 = 問('日期', '格式 2026-01-18');
  if (日期 === null) return;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(日期)) return ui.alert('日期要寫成 2026-01-18 這種格式');

  const 地點 = 問('地點', '可以留空');
  if (地點 === null) return;

  const 組別字串 = 問('要哪幾組？', '用逗號分隔。留空就用公版十三組：\n' + 預設組別.join('、'));
  if (組別字串 === null) return;
  const 組別 = 組別字串
    ? 組別字串.split(/[,，、]/).map(x => x.trim()).filter(Boolean).map(加組字_)
    : null;

  ui.alert(開新活動(名稱, 日期, 地點, 組別));
}

/** 組別一律以「組」結尾，自訂的也補上——跟網頁那邊的規則一致 */
function 加組字_(n) {
  const s = String(n || '').trim();
  return (!s || s.slice(-1) === '組') ? s : s + '組';
}

function 目前設定_() {
  const 行 = [];
  [['總表', P_MASTER], ['範本｜組員', P_TPL_TEAM], ['範本｜核心', P_TPL_CORE]].forEach(([名, key]) => {
    const id = PROP.getProperty(key);
    行.push(名 + '：' + (id ? DriveApp.getFileById(id).getUrl() : '還沒建立'));
  });
  const 活動 = SpreadsheetApp.openById(必須有_(P_MASTER, '總表')).getSheetByName('活動');
  行.push('', '已經建了 ' + Math.max(0, 活動.getLastRow() - 1) + ' 場活動');
  SpreadsheetApp.getUi().alert('五表｜目前設定', 行.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
