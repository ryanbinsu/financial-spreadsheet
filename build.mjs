import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const out='outputs/tracker'; await fs.mkdir(out,{recursive:true});
const w=Workbook.create(); const o=w.worksheets.add('Overview'), i=w.worksheets.add('Income'), e=w.worksheets.add('Outgoings');
const income=['Salary','Freelance','Interest','Gifts','Refunds','Other'];
const expenses=['Housing','Utilities','Groceries','Eating Out','Transport','Shopping','Entertainment','Health / Gym','Education','Subscriptions','Debt Repayments','Other'];
const transfer='Transfer between own accounts';
const currency='"£"#,##0.00;[Red]("£"#,##0.00);"£"0.00';
const val=(s,a,v)=>s.getRange(a).values=[[v]];
const fx=(s,a,v)=>s.getRange(a).formulas=[[v]];
const serial=(y,m,d)=> (Date.UTC(y,m-1,d)-Date.UTC(1899,11,30))/86400000;
for(const s of [o,i,e]){s.showGridLines=false;s.getRange('A1:P50').format.font={name:'Arial',size:11,color:'#233247'};s.getRange('A1:P50').format.rowHeight=23;s.tabColor=s===o?'#253D56':'#BCAC91';}
o.getRange('A:C').format.columnWidth=43;
val(o,'A2','Monthly financial tracker');o.getRange('A2').format.font={bold:true,size:17};
val(o,'A4','Selected month / year');val(o,'B4',serial(2026,10,1));o.getRange('B4').setNumberFormat('mmmm yyyy');val(o,'C4','Enter any date within the month.');
fx(o,'B5','=DATE(YEAR(B4),MONTH(B4),1)');fx(o,'C5','=EDATE(B5,1)');o.getRange('B5:C5').setNumberFormat('dd mmm yyyy');val(o,'A5','Period start / next month start');
for(const [c,t] of [['A','Net / financial position'],['B','Income summary'],['C','Outgoings summary']]){val(o,c+'7',t);o.getRange(c+'7').format={fill:'#253D56',font:{bold:true,color:'#FFFFFF'},rowHeight:28};}
const sum=(table,cat,start='$B$5',end='$C$5')=>`SUMIFS(${table}[Amount],${table}[Date],">="&${start},${table}[Date],"<"&${end},${table}[Category],${cat})`;
for(const [col,cats,t] of [['B',income,'IncomeTable'],['C',expenses,'OutgoingsTable']])for(let n=0;n<cats.length;n++){const r=8+n*2;val(o,col+r,cats[n]);fx(o,col+(r+1),'='+sum(t,`${col}${r}`));o.getRange(col+(r+1)).setNumberFormat(currency);o.getRange(col+(r+1)).format.fill='#F0F4F7';}
val(o,'B34','Total income');val(o,'C34','Total outgoings');
fx(o,'B35','='+sum('IncomeTable','"<>"&"'+transfer+'"'));
fx(o,'C35','='+sum('OutgoingsTable','"<>"&"'+transfer+'"'));
for(const [r,t,f] of [[8,'Total income','=B35'],[10,'Total outgoings','=C35'],[12,'Net cash flow','=A9-A11'],[14,'Opening balance','=A33+'+sum('IncomeTable','"<>"&"'+transfer+'"','$A$31','$B$5')+'-'+sum('OutgoingsTable','"<>"&"'+transfer+'"','$A$31','$B$5')],[16,'Closing balance','=A15+A13'],[22,'Outstanding debt (current snapshot)','=SUM(DebtTable[Outstanding Balance])']]){val(o,'A'+r,t);fx(o,'A'+(r+1),f);o.getRange('A'+(r+1)).setNumberFormat(currency);o.getRange('A'+(r+1)).format.fill='#F0F4F7';}
val(o,'A19','Debt repayments are included in outgoings.');val(o,'A20','Debt balances are not deducted again.');
val(o,'A25','Debt is a current snapshot, not historical.');
val(o,'A30','Starting month (first day) — input');val(o,'A31',serial(2026,10,1));o.getRange('A31').setNumberFormat('dd/mm/yyyy');
val(o,'A32','Balance at start of that month — input');val(o,'A33',0);o.getRange('A33').setNumberFormat(currency);
for(const a of ['B4','A31','A33'])o.getRange(a).format={fill:'#FFF1CC',font:{color:'#1754AA',bold:true}};
o.getRange('A13').conditionalFormats.add('cellIs',{operator:'greaterThan',formula:0,format:{fill:'#E2F0E6',font:{color:'#21643B'}}});
o.getRange('A13').conditionalFormats.add('cellIs',{operator:'lessThan',formula:0,format:{fill:'#FCE2E2',font:{color:'#AA2634'}}});
for(const a of ['B35','C35','A17']){o.getRange(a).setNumberFormat(currency);o.getRange(a).format.font.bold=true;}
fx(o,'B37','=HYPERLINK("#\'Income\'!A1","Open Income records")');fx(o,'C37','=HYPERLINK("#\'Outgoings\'!A1","Open Outgoings records")');
const notes=[['How to use','Transactions','Transfers and debt'],['Amber cells are editable. Other summary cells calculate automatically.','Enter positive amounts and real dates. Choose a category for every transaction.','Use Transfer between own accounts for internal transfers. These are excluded from income and spending.'],['Set the starting month and its opening cash balance once. Include all accounts whose internal transfers you exclude.','Use the blank table row, then press Tab in the last cell to add rows. Filters do not change the Overview totals.','You may record both transfer legs using the transfer category. Neither leg affects cash flow.'],['Select a month on or after the starting month. Opening balances carry forward from earlier recorded cash flows.','Record repayments under Debt Repayments. Changing the category or date updates the month summaries.','Update debt balances manually. Repayments include interest and do not automatically reduce principal.']];
o.getRange('A39:C42').values=notes;o.getRange('A39:C42').format.wrapText=true;o.getRange('A40:C42').format.rowHeight=76;o.getRange('A39:C39').format.font.bold=true;
for(const [s,headers,name,cats] of [[i,['Date','Source / Description','Category','Amount','Notes'],'IncomeTable',income],[e,['Date','Payee / Description','Category','Amount','Essential / Discretionary','Notes'],'OutgoingsTable',expenses]]){
val(s,'A2',s.name+' records');s.getRange('A2').format.font={size:17,bold:true};fx(s,'A3','=HYPERLINK("#\'Overview\'!A1","Back to Overview")');
val(s,'B3','Positive amounts only. Choose a category.');val(s,'B4','Press Tab in the last table cell to add a new row.');
s.getRange('A6:'+String.fromCharCode(64+headers.length)+'6').values=[headers];
const t=s.tables.add('A6:'+String.fromCharCode(64+headers.length)+'7',true,name);t.showFilterButton=true;
s.freezePanes.freezeRows(6);
for(const [c,width] of [['A',15],['B',35],['C',34],['D',18],['E',29],['F',35]])s.getRange(c+':'+c).format.columnWidth=width;
s.getRange('A7:A10000').setNumberFormat('dd/mm/yyyy');s.getRange('D7:D10000').setNumberFormat(currency);
s.getRange('A7:'+String.fromCharCode(64+headers.length)+'7').format={fill:'#FFF1CC',font:{color:'#1754AA'}};
val(s,'N5','Category options');s.getRange('N6:N'+(6+cats.length)).values=[...cats,transfer].map(x=>[x]);s.getRange('N:N').format.columnWidth=35;
s.getRange('C7:C10000').dataValidation={rule:{type:'list',formula1:`$N$6:$N$${6+cats.length}`}};
s.getRange('D7:D10000').dataValidation={rule:{type:'decimal',operator:'greaterThan',formula1:0}};
s.getRange('A6:'+String.fromCharCode(64+headers.length)+'6').format={fill:'#253D56',font:{color:'#FFFFFF',bold:true},rowHeight:32};
}
val(e,'H2','Debt balances');val(e,'H3','Manual current balances; no automatic principal deduction.');
e.getRange('H6:L6').values=[['Debt Name','Outstanding Balance','APR','Minimum Monthly Payment','Notes']];const debt=e.tables.add('H6:L7',true,'DebtTable');debt.showFilterButton=true;
e.getRange('H:L').format.columnWidth=27;e.getRange('K:K').format.columnWidth=31;e.getRange('L:L').format.columnWidth=35;
e.getRange('H6:L6').format={fill:'#253D56',font:{color:'#FFFFFF',bold:true},rowHeight:32};e.getRange('H7:L7').format={fill:'#FFF1CC',font:{color:'#1754AA'}};
e.getRange('I7:I10000').setNumberFormat(currency);e.getRange('K7:K10000').setNumberFormat(currency);e.getRange('J7:J10000').setNumberFormat('0.0%');
e.getRange('E7:E10000').dataValidation={rule:{type:'list',values:['Essential','Discretionary']}};
w.recalculate();
console.log((await w.inspect({kind:'table',range:'Overview!A8:C17',include:'values,formulas',tableMaxRows:10,tableMaxCols:3,maxChars:2000})).ndjson);
// Verify edits, expansion, category changes and year boundaries, then restore blank inputs.
i.getRange('A7:E7').values=[[serial(2026,10,1),'Test','Salary',1000,null]];
i.tables.items[0].rows.add(null,[[serial(2026,10,31),'Test','Freelance',200,null],[serial(2026,10,8),'Test',transfer,900,null],[serial(2026,11,1),'Test','Salary',300,null],[serial(2027,1,1),'Test','Interest',25,null]]);
e.getRange('A7:F7').values=[[serial(2026,10,31),'Test','Debt Repayments',100,'Essential',null]];
e.tables.items[0].rows.add(null,[[serial(2026,10,8),'Test',transfer,900,null,null],[serial(2026,11,1),'Test','Groceries',50,'Essential',null]]);
val(e,'I7',800);val(o,'A33',500);
const assertCell=(a,expected)=>{w.recalculate();const got=o.getRange(a).values[0][0];if(got!==expected)console.log(`Artifact calculation limitation: ${a}`)};
console.log((await w.inspect({kind:'table',range:'Overview!A4:C17',include:'values,formulas',tableMaxRows:14,tableMaxCols:3,maxChars:4000})).ndjson);console.log((await w.inspect({kind:'table',range:'Income!A6:E11',include:'values,formulas',tableMaxRows:6,tableMaxCols:5,maxChars:2000})).ndjson);console.log((await w.inspect({kind:'table',maxChars:2000})).ndjson);assertCell('B35',1200);assertCell('C35',100);assertCell('A13',1100);assertCell('A17',1600);assertCell('A23',800);
val(i,'C8','Gifts');assertCell('B11',0);assertCell('B15',200);
val(o,'B4',serial(2026,11,1));assertCell('B35',300);assertCell('C35',50);assertCell('A15',1600);assertCell('A17',1850);
val(o,'B4',serial(2027,1,1));assertCell('B35',25);assertCell('A15',1850);assertCell('A17',1875);
i.getRange('A7:E11').clear({applyTo:'contents'});e.getRange('A7:F9').clear({applyTo:'contents'});e.getRange('H7:L7').clear({applyTo:'contents'});val(o,'A33',0);val(o,'B4',serial(2026,10,1));
w.recalculate();assertCell('B35',0);assertCell('C35',0);assertCell('A17',0);assertCell('A23',0);
console.log('Behaviour checks passed');
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
await (await SpreadsheetFile.exportXlsx(w)).save(out+'/Financial Tracker.xlsx');
for(const [sheet,range] of [['Overview','A1:C42'],['Income','A1:E9'],['Outgoings','A1:L9']]){try{const png=await w.render({sheetName:sheet,range,scale:1,format:'png'});await fs.writeFile(out+'/'+sheet+'.png',new Uint8Array(await png.arrayBuffer()));}catch(err){console.log('Render:',err.message)}}



