import ExcelJS from 'exceljs';

const directors = [['緯宸','會長'],['倚涵','青年團團長'],['智鴻','活動組長'],['偉晴','副組長']];
const flocks = [
  ['晨曦雁群',[['緯宸','weichen@gmail.com'],['綺晨','qichen@gmail.com'],['雅方','yafang@gmail.com'],['筱穎','xiaoying@gmail.com'],['冠宇','guanyu@gmail.com'],['怡君','yijun@gmail.com'],['安琪','anqi@gmail.com']]],
  ['和風雁群',[['倚涵','yihan@gmail.com'],['俊良','junliang@gmail.com'],['子牟','zimou@gmail.com'],['惠萱','huixuan@gmail.com'],['明珠','mingzhu@gmail.com'],['品妤','pinyu@gmail.com'],['雨柔','yurou@gmail.com']]],
  ['遠山雁群',[['智鴻','zhihong@gmail.com'],['志仲','zhizhong@gmail.com'],['興隆','xinglong@gmail.com'],['振軒','zhenxuan@gmail.com'],['敬傑','jingjie@gmail.com'],['尚恩','shangen@gmail.com'],['思齊','siqi@gmail.com']]],
  ['星海雁群',[['偉晴','weiqing@gmail.com'],['馥瑜','fuyu@gmail.com'],['承恩','chengen@gmail.com'],['緯聖','weisheng@gmail.com'],['智為','zhiwei@gmail.com'],['秉芸','bingyun@gmail.com'],['沛筠','peiyun@gmail.com'],['若晴','ruoqing@gmail.com']]],
];
const groups = [['報到組','綺晨',['雅方','筱穎','佳蓉','婷羽']],['引導組','俊良',['子牟','惠萱','明珠','子倫','婉玲']],['交通組','志仲',['興隆','振軒','俊廷']],['財務組','馥瑜',['宛蓉','詩涵']],['計分組','敬傑',['尚恩','心如','凱倫','珮羽']],['主持組','承恩',['語晴','品妤']],['環境組','緯聖',['聖洲','靜雯','柏豪','心怡','思羽']],['攝影組','智為',['冠宇','哲豪','佩雯']],['醫療組','秉芸',['怡君','恩慧']],['佈置組','沛筠',['千淳','若晴','世傑','美玲']]];
const slots = [
  ['12:00','12:15','前置作業'],['12:15','12:30','全體會前會'],['12:30','13:05','報到與入場'],
  ['13:05','13:10','運動員宣誓'],['13:10','15:20','比賽與輪打'],['15:20','15:40','集合與頒獎'],['15:40','16:10','善後與歸還'],
];

// 人員庫：把所有出現過的人收齊，雁群名單有 Gmail 的直接帶入，其餘留空待補
const mail = new Map(), flockOf = new Map();
for (const [f, people] of flocks) for (const [n, m] of people) { mail.set(n, m); flockOf.set(n, f); }
const everyone = [];
const add = n => { if (!everyone.some(p => p[0] === n)) everyone.push([n, mail.get(n) || '', flockOf.get(n) || '', '在職']); };
for (const [n] of directors) add(n);
for (const [, leader, members] of groups) { add(leader); members.forEach(add); }
for (const [, people] of flocks) for (const [n] of people) add(n);

const wb = new ExcelJS.Workbook();
wb.creator = '五表活動協作系統';

const tabs = [
  ['人員庫',   ['姓名','Gmail','雁群','狀態'], everyone],
  ['權限',     ['Gmail','姓名','角色','備註'], [
      [mail.get('緯宸'), '緯宸', '會長', '固定，不隨活動改變'],
      [mail.get('倚涵'), '倚涵', '青年團團長', '每年改選，換人時改這一列'],
  ]],
  ['活動',     ['活動ID','活動名稱','日期','地點','狀態','工作組'], [
      ['A001','2027 春季運動會','2027-05-24','臺北活動中心','進行中', groups.map(g => g[0]).join(',')],
  ]],
  ['活動人力', ['活動ID','Gmail','姓名','組別','角色'], [
      ...directors.filter(d => d[1] !== '會長' && d[1] !== '青年團團長').map(([n, r]) => ['A001', mail.get(n) || '', n, '核心', r]),
      ...groups.flatMap(([g, leader, members]) => [
        ['A001', mail.get(leader) || '', leader, g, '組長'],
        ...members.map(m => ['A001', mail.get(m) || '', m, g, '組員']),
      ]),
  ]],
  ['會議',     ['活動ID','會議名稱','日期'], [
      ['A001','第一次會前會','2027-05-17'],['A001','第二次會前會','2027-05-18'],
      ['A001','第三次會前會','2027-05-21'],['A001','全體會前會','2027-05-23'],
  ]],
  ['工作進度', ['活動ID','會議','組別','負責人','要完成事項','目前進度','下一步','期限','狀態'], []],
  ['活動細流', ['活動ID','開始','結束','全體流程','組別','該組工作'], []],
  ['物資借用', ['活動ID','組別','物資','數量','攜帶者','借出照片','已還','已歸位','歸還照片','備註'], []],
  ['收支',     ['活動ID','類型','組別','項目','金額','登錄人','備註'], [
      ['A001','成本','報到組','名牌與簽到用品',1800,'綺晨',''],
      ['A001','收入','財務組','報名費',12000,'馥瑜','60 人 × 200'],
  ]],
  ['會後精進', ['活動ID','組別','問題','影響','下次做法','記錄人'], []],
  ['細流範本', ['開始','結束','全體流程','組別','該組工作'],
      slots.flatMap(([s, e, name]) => groups.map(g => [s, e, name, g[0], '']))],
  ['設定',     ['項目','值','說明'], [
      ['預設組別', groups.map(g => g[0]).join(','), '建立活動時預設勾選的工作組'],
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
  headers.forEach((h, i) => { ws.getColumn(i + 1).width = Math.max(12, h.length * 2.4 + 6); });
  ws.getColumn(1).width = name === '設定' ? 22 : 14;
}

const out = 'F:/codex/五表/outputs/five-forms/五表後台.xlsx';
await wb.xlsx.writeFile(out);
console.log('寫出：' + out);
for (const [name, headers, rows] of tabs) console.log(`  ${name}：${headers.length} 欄・${rows.length} 列`);
