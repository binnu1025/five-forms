import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
const out='F:/codex/五表/outputs/five-forms'; await fs.mkdir(out,{recursive:true});
const wb=Workbook.create();
const people=wb.worksheets.add('人力表'), flow=wb.worksheets.add('細流'), tasks=wb.worksheets.add('各組工作進度'), items=wb.worksheets.add('物資借用表'), reviews=wb.worksheets.add('會後精進紀錄');
const groups=[['報到組','綺晨','雅方、筱穎、佳蓉、婷羽'],['引導組','俊良','子牟、惠萱、明珠、子倫、婉玲'],['交通組','志仲','興隆、振軒、俊廷'],['財務組','馥瑜','宛蓉、詩涵'],['計分組','敬傑','尚恩、心如、凱倫、珮羽'],['主持組','承恩','語晴、品妤'],['環境組','緯聖','聖洲、靜雯、柏豪、心怡、思羽'],['攝影組','智為','冠宇、哲豪、佩雯'],['醫療組','秉芸','怡君、恩慧'],['佈置組','沛筠','千淳、若晴、世傑、美玲']];
people.getRange('A1:D1').values=[['角色','組別','姓名','Gmail']];
let pr=[['會長','核心','緯宸','president@gmail.com'],['青年團團長','核心','倚涵','leader@gmail.com'],['活動組長','核心','智鴻','event@gmail.com'],['副組長','核心','偉晴','deputy@gmail.com']];
for(const [g,l,ms] of groups){pr.push(['組長',g,l,`${l}@gmail.com`]); for(const m of ms.split('、'))pr.push(['組員',g,m,`${m}@gmail.com`]);}
people.getRange('A2').write(pr);
const slots=[['12:00','12:15','前置作業','擺放報到桌','貼引導圖與動線線','確認停車動線','準備零用金','準備計分紙','試音','垃圾桶就位','拍攝前置','醫療箱就位','拱門與指示牌'],['12:15','12:30','全體會前會','名單點交','引導組動線確認','確認分組路線','預算與收據提醒','分數規則確認','串場順序確認','清潔區域分工','鏡位分工','緊急流程確認','安全巡檢'],['12:30','13:05','報到與入場','核對名單與名牌','引導到隊伍區','維持車流與停車','處理現場繳費','發放計分表','暖場廣播','入口與場邊巡視','報到畫面與合照','注意不適參加者','補強指示牌'],['13:05','13:10','運動員宣誓','暫停報到','集合隊伍','維持場外動線','備查','備查','引導宣誓','維持動線','特寫與全景','備查','備查'],['13:10','15:20','比賽與輪打','遲到人員補報到','分流與換場','場外交通與接送','費用登記','即時計分與核對','場次播報','每30分鐘巡場','各場動態拍攝','受傷處置與紀錄','器材補位'],['15:20','15:40','集合與頒獎','協助集合','引導得獎者','維持出口車流','核對獎品','確認名次','主持頒獎','清空頒獎區','頒獎與大合照','備查','獎品上台'],['15:40','16:10','善後與歸還','報到桌撤收','出口動線引導','車輛離場引導','收據整理','計分資料封存','失物招領廣播','垃圾分類與場復','備份照片','醫療箱點交','物資點收']];
flow.getRange('A1:M1').values=[['開始','結束','全體流程',...groups.map(g=>g[0])]]; flow.getRange('A2').write(slots);
tasks.getRange('A1:G1').values=[['組別','任務','負責人','截止日','狀態','完成條件','備註']]; let tr=[]; for(const [g,l] of groups){tr.push([g,`完成${g}會前分工`,l,new Date('2027-05-21'),'進行中','組員已確認任務',''],[g,`確認${g}當天物資`,l,new Date('2027-05-23'),'未開始','數量、攜帶者、照片齊全','']);} tasks.getRange('A2').write(tr); tasks.getRange(`D2:D${tr.length+1}`).setNumberFormat('yyyy-mm-dd');
items.getRange('A1:J1').values=[['組別','物資','數量','攜帶者','借出照片','已還','已歸位','歸還照片','狀態','備註']]; items.getRange('A2').write([['計分組','計分紙',40,'敬傑','待上傳',false,false,'','待歸還','比賽結束後回收'],['攝影組','腳架',2,'智為','待上傳',true,false,'待上傳','待歸位',''],['環境組','水桶',4,'緯聖','待上傳',false,false,'','待歸還',''],['環境組','球具回收桶',2,'靜雯','待上傳',true,true,'待上傳','已完成',''],['醫療組','醫療箱',1,'秉芸','待上傳',false,false,'','活動中','善後點交'],['佈置組','打卡版',4,'沛筠','待上傳',true,true,'待上傳','已完成','']]);
reviews.getRange('A1:E1').values=[['組別','問題','影響','下次做法','記錄人']]; reviews.getRange('A2').write([['報到組','新成員不熟工作','開場時分工混亂','會前辦一次組內演練','綺晨'],['引導組','報到後動線不清','人潮堵塞','入口增設大型引導圖','俊良'],['環境組','巡場頻率不足','場邊垃圾堆積','每30分鐘固定巡場','緯聖']]);
for(const s of [people,flow,tasks,items,reviews]){s.showGridLines=false;s.freezePanes.freezeRows(1);const used=s.getUsedRange();used.format.font={name:'Arial',size:10,color:'#17231D'};used.format.verticalAlignment='center';used.format.rowHeight=24;used.format.borders={preset:'inside',style:'thin',color:'#E1E7E2'};const head=s.getRangeByIndexes(0,0,1,used.columnCount);head.format={fill:'#245A45',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center'};used.format.autofitColumns();}
people.tabColor='#245A45';flow.tabColor='#3C785D';items.tabColor='#D77B3F';
wb.recalculate();
const check=await wb.inspect({kind:'table',sheetId:'細流',range:'A1:M8',include:'values,formulas',tableMaxRows:8,tableMaxCols:13}); console.log(check.ndjson);
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:100},summary:'formula errors'}); console.log(errors.ndjson);
for(const s of ['人力表','細流','各組工作進度','物資借用表','會後精進紀錄']){const p=await wb.render({sheetName:s,autoCrop:'all',scale:1,format:'png'});await fs.writeFile(`${out}/${s}.png`,new Uint8Array(await p.arrayBuffer()));}
const file=await SpreadsheetFile.exportXlsx(wb); await file.save(`${out}/2027春季運動會_五表模擬資料.xlsx`);
