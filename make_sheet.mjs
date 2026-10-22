// 產生 Google Sheet 後台的骨架檔。用法：npm i exceljs && node make_sheet.mjs
//
// 如果同目錄有 roster.local.json（真實名單，刻意不進版控），人員庫與權限會用它；
// 沒有的話就用下面的示範資料。真實信箱不要寫進這個檔案，repo 是公開的。
import ExcelJS from 'exceljs';
import { readFileSync, existsSync } from 'node:fs';

const demoFlocks = [
  { 小組: '晨曦雁群', 成員: [['緯宸','weichen@example.com'],['綺晨','qichen@example.com'],['雅方','yafang@example.com']] },
  { 小組: '和風雁群', 成員: [['倚涵','yihan@example.com'],['俊良','junliang@example.com']] },
];
const roster = existsSync('roster.local.json')
  ? JSON.parse(readFileSync('roster.local.json', 'utf8'))
  : { 雁群小組: demoFlocks, 權限: [] };
const real = existsSync('roster.local.json');

const people = roster.雁群小組.flatMap(({ 小組, 成員 }) => 成員.map(([n, m]) => [n, m, 小組, '在職']));
const groups = ['報到組','引導組','交通組','財務組','計分組','主持組','環境組','攝影組','醫療組','佈置組'];
const slots = [
  ['12:00','12:15','前置作業'],['12:15','12:30','全體會前會'],['12:30','13:05','報到與入場'],
  ['13:05','13:10','運動員宣誓'],['13:10','15:20','比賽與輪打'],['15:20','15:40','集合與頒獎'],['15:40','16:10','善後與歸還'],
];

const wb = new ExcelJS.Workbook();
wb.creator = '五表活動協作系統';

const tabs = [
  ['人員庫',   ['姓名','Gmail','雁群小組','狀態'], people],
  ['權限',     ['Gmail','姓名','角色','備註'], roster.權限 ?? []],
  ['活動',     ['活動ID','活動名稱','日期','地點','狀態','工作組'],
      [['A001','（範例）請改成你的活動','2027-05-24','','規劃中', groups.join(',')]]],
  ['活動人力', ['活動ID','Gmail','姓名','組別','角色'], []],
  ['會議',     ['活動ID','會議名稱','日期'], [['A001','第一次會前會','']]],
  ['工作進度', ['活動ID','會議','組別','負責人','要完成事項','目前進度','下一步','期限','狀態'], []],
  ['活動細流', ['活動ID','開始','結束','全體流程','組別','該組工作'], []],
  ['物資借用', ['活動ID','組別','物資','數量','攜帶者','借出照片','已還','已歸位','歸還照片','備註'], []],
  ['收支',     ['活動ID','類型','組別','項目','金額','登錄人','備註'], []],
  ['會後精進', ['活動ID','組別','問題','影響','下次做法','記錄人'], []],
  ['細流範本', ['開始','結束','全體流程','組別','該組工作'],
      slots.flatMap(([s, e, name]) => groups.map(g => [s, e, name, g, '']))],
  ['設定',     ['項目','值','說明'], [
      ['預設組別', groups.join(','), '建立活動時預設勾選的工作組'],
      ['Drive 照片資料夾ID', '', '物資照片上傳的目的地，填資料夾網址最後那一段'],
      ['可看收支的角色', '會長,青年團團長,活動組長,副組長,財務長', '看得到全部組別與總額'],
  ]],
];

for (const [name, headers, rows] of tabs) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.addRow(headers);
  rows.forEach(r => ws.addRow(r));
  const head = ws.getRow(1);
  head.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  head.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF245A45' } };
  head.alignment = { vertical: 'middle', horizontal: 'center' };
  head.height = 24;
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
  headers.forEach((h, i) => { ws.getColumn(i + 1).width = Math.max(14, h.length * 2.4 + 8); });
  if (name === '人員庫') ws.getColumn(2).width = 34;
  if (name === '權限') ws.getColumn(1).width = 34;
  if (name === '設定') ws.getColumn(1).width = 24;
}

// 檔案在 Excel 裡開著會寫不進去，這時可以傳一個別的路徑進來
const out = process.argv[2] || 'F:/codex/五表/outputs/five-forms/五表後台.xlsx';
await wb.xlsx.writeFile(out);
console.log(`名單來源：${real ? 'roster.local.json（真實名單）' : '示範資料'}`);
console.log('寫出：' + out);
for (const [name, headers, rows] of tabs) console.log(`  ${name}：${headers.length} 欄・${rows.length} 列`);
