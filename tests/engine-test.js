/* Extracts the engine out of index.html and exercises it in node. */
const fs = require('fs');
const html = fs.readFileSync('/home/user/Trust-me-/index.html', 'utf8');
let js = html.match(/<script>\n([\s\S]*?)<\/script>/)[1];
js = js.slice(0, js.lastIndexOf('/* ====', js.indexOf('   11. Wiring')));

const store = {};
global.window = { localStorage:{ getItem:k=>k in store?store[k]:null,
  setItem:(k,v)=>{store[k]=String(v)}, removeItem:k=>{delete store[k]} } };
function fake(){ return { textContent:'', innerHTML:'', value:'', checked:false, className:'',
  style:{}, dataset:{}, classList:{add(){},remove(){},contains:()=>false},
  querySelectorAll:()=>[], querySelector:()=>fake(), appendChild(){}, click(){}, addEventListener(){} }; }
global.document = { getElementById:()=>fake(), querySelectorAll:()=>[], querySelector:()=>fake(),
  createElement:()=>fake(), addEventListener(){} };

const E = new Function(js + `return { CONFIG, STATE, parseCSV, parseNum, mapHeaders, resolveBrand,
  normName, ingestOrg, ingestStats, commitStats, buildBoard, computePool, computePayouts,
  currentBoard, attendanceFor, payWithMetrics, projectMetrics, solvePath, suggestFor,
  ignoreName, reprocess, orderAgents, rand, pct, num, reconcile, orgIndex };`)();

let pass=0, fail=0;
const ok=(n,c,x)=>{ c?(pass++,console.log('  PASS  '+n)):(fail++,console.log('  FAIL  '+n+(x!==undefined?'  →  '+x:''))); };
const near=(a,b,t)=>Math.abs(a-b)<=(t===undefined?0.01:t);
const head=t=>console.log('\n=== '+t+' ===');

const ORG   = fs.readFileSync('/home/user/Trust-me-/organogram.csv','utf8');
const STATS = fs.readFileSync('/home/user/Trust-me-/sample-real-format.csv','utf8');

/* ---------------------------------------------------------- 1. blank cols */
head('1. Real export shape');
const p = E.parseCSV(STATS);
ok('delimiter is ";"', p.delim===';', p.delim);
ok('4 empty trailing columns dropped', p.dropped===4, p.dropped);
ok('11 real headings kept', p.headers.length===11, p.headers.length);
ok('no blank heading survives', p.headers.every(h=>h.trim()!==''));
const m = E.mapHeaders(p.headers);
ok('contacts <- Total Contacts (FA)', m.contacts==='Total Contacts (FA)', m.contacts);
ok('contacts NOT Inbound Calls', m.contacts!=='Inbound Calls');
ok('sales <- Gross Sales', m.sales==='Gross Sales', m.sales);
ok('sales NOT Gross Closings', m.sales!=='Gross Closings');
ok('closing <- Gross Closings', m.closing==='Gross Closings', m.closing);
ok('leads <- Leads', m.leads==='Leads', m.leads);
ok('eff <- Gross Eff (FA)', m.eff==='Gross Eff (FA)', m.eff);
ok('agent <- Operator Name', m.name==='Operator Name', m.name);
ok('no brand column found (correct)', !m.brand, m.brand);
ok('no QA column found (correct)', !m.qa, m.qa);

head('2. SA number parsing');
ok('"R 1 072 000,00"', E.parseNum('R 1 072 000,00')===1072000);
ok('"32,5"', E.parseNum('32,5')===32.5);
ok('"31,6%"', E.parseNum('31,6%')===31.6);
ok('"1.234,56"', E.parseNum('1.234,56')===1234.56);

head('3. Brand aliases (their own codes)');
ok('BIB -> Budget', E.resolveBrand('BIB')==='BUDGET', E.resolveBrand('BIB'));
ok('A&G -> AG', E.resolveBrand('A&G')==='AG');
ok('FFW -> FFW', E.resolveBrand('FFW')==='FFW');
ok('Budget Insurance', E.resolveBrand('Budget Insurance')==='BUDGET');
ok('Auto & General', E.resolveBrand('Auto & General')==='AG');
ok('First for Women', E.resolveBrand('First for Women')==='FFW');

/* --------------------------------------------------------- 4. organogram */
head('4. Organogram');
ok('loads', E.ingestOrg(ORG)===true);
ok('27 people', E.STATE.org.length===27, E.STATE.org.length);
const byBrand={}; E.STATE.org.forEach(x=>byBrand[x.brand]=(byBrand[x.brand]||0)+1);
ok('7 A&G / 10 Budget / 10 FFW', byBrand.AG===7&&byBrand.BUDGET===10&&byBrand.FFW===10, JSON.stringify(byBrand));
ok('everyone has a brand', E.STATE.org.every(x=>x.brand));
const sio = E.STATE.org.find(x=>x.salesName.indexOf('Siobhan')===0);
ok('two name columns kept apart', sio.rawName==='Sioban Nuwenhuis' && sio.salesName==='Siobhan Nieuwenhuis - JHB Agent',
   sio.rawName+' / '+sio.salesName);

/* --------------------------------------------------------- 5. name match */
head('5. Name matching — exact, never guessing');
E.ingestStats(STATS, '2026-08-20', false);
const rep = E.STATE.report;
ok('27 organogram agents matched', rep.loaded===27, rep.loaded);
ok('ZERO names need matching for the 27', E.STATE.unknown.length===6, E.STATE.unknown.length+' unknown');
ok('the 6 unknown are all Kimisha’s team',
   E.STATE.unknown.every(n=>['Bradwin Chetty','Inayith Naidoo','Makaylan Naicker','Lenisha Ebrahim','Pam Bulose','Ronnel Naidoo'].indexOf(n)>-1),
   E.STATE.unknown.join(', '));
ok('4 "Total" rows skipped', rep.skipped.filter(s=>s.name==='Total').length===4,
   rep.skipped.filter(s=>s.name==='Total').length);
ok('"Sioban Nuwenhuis" resolved via Raw Data Name',
   E.STATE.days['2026-08-20'].rows.some(r=>r.key===E.normName('Sioban Nuwenhuis')));
ok('no suggestion for an unrelated name', (()=>{
  const s = E.suggestFor('Sithabile Mthethwa');
  return !s || s.rawName==='Sithabile Mthethwa';
})(), (E.suggestFor('Sithabile Mthethwa')||{}).salesName);
ok('no suggestion for Kimisha’s Bradwin Chetty', !E.suggestFor('Bradwin Chetty'),
   (E.suggestFor('Bradwin Chetty')||{}).salesName);

/* ------------------------------------------------- 6. exclusion is structural */
head('6. Off-organogram staff excluded from the MATHS');
let board = E.currentBoard();
ok('board has exactly 27 agents', board.agents.length===27, board.agents.length);
ok('no Kimisha agent on the board',
   !board.agents.some(a=>['Bradwin Chetty','Pam Bulose','Ronnel Naidoo'].indexOf(a.name)>-1));
const B = board.brands;
ok('Budget headcount 10 (not 16)', B.BUDGET.headcount===10, B.BUDGET.headcount);
ok('A&G headcount 7', B.AG.headcount===7, B.AG.headcount);
ok('FFW headcount 10', B.FFW.headcount===10, B.FFW.headcount);
ok('Dial Direct headcount 0', B.DIALDIRECT.headcount===0, B.DIALDIRECT.headcount);
/* the average must come from organogram members only */
const budgetOrg = E.STATE.org.filter(x=>x.brand==='BUDGET').map(x=>E.normName(x.rawName));
const budgetRows = E.STATE.days['2026-08-20'].rows.filter(r=>budgetOrg.indexOf(r.key)>-1);
const expAvg = budgetRows.reduce((s,r)=>s+r.contacts,0)/budgetRows.length;
ok('Budget avg contacts from organogram members only', near(B.BUDGET.avgContacts, expAvg, 0.01),
   B.BUDGET.avgContacts.toFixed(2)+' vs '+expAvg.toFixed(2));
ok('gate = 90% of that average', near(B.BUDGET.gate, expAvg*0.9, 0.01), B.BUDGET.gate.toFixed(1));
ok('each brand has its own average', new Set([B.BUDGET.avgContacts,B.AG.avgContacts,B.FFW.avgContacts]).size===3);

/* ---------------------------------------------------------- 7. reconcile */
head('7. Column reconciliation on the real file');
ok('conversion reconciles on every row', rep.recon.conv.bad===0, rep.recon.conv.bad+'/'+rep.recon.n);
ok('CLOSING reconciles on every row', rep.recon.closing.bad===0, rep.recon.closing.bad+'/'+rep.recon.n);
ok('effectiveness reconciles on every row', rep.recon.eff.bad===0, rep.recon.eff.bad+'/'+rep.recon.n);
const jp = board.agents.find(a=>a.name==='Jermaine Pillay');
ok('Jermaine Pillay 285/116/53', jp.contacts===285&&jp.leads===116&&jp.sales===53);
ok('  conv 40,7%', near(jp.conversion*100, 40.7, .05), (jp.conversion*100).toFixed(1));
ok('  closing 45,7%', near(jp.closing*100, 45.7, .05), (jp.closing*100).toFixed(1));
ok('  effectiveness 18,6%', near(jp.effectiveness, 18.6, .05), jp.effectiveness.toFixed(1));

/* --------------------------------------------------------------- 8. gates */
head('8. Two gates');
const q = board.agents.filter(a=>a.qualified);
ok('some qualify', q.length>0, q.length);
ok('every qualifier cleared BOTH gates', q.every(a=>a.passContacts&&a.passEff));
ok('every non-qualifier earns exactly R0', board.agents.filter(a=>!a.qualified).every(a=>a.pay===0));
ok('Jermaine (18,6% vs 31,6%) earns R0', jp.pay===0, jp.pay);
ok('he is told how many sales short', jp.salesShort===Math.ceil(0.316*285-53), jp.salesShort);
const oneGate = board.agents.find(a=>a.passContacts&&!a.passEff);
ok('clearing only the contacts gate still pays R0', !oneGate || oneGate.pay===0);

/* ------------------------------------------------------------ 9. payouts */
head('9. Pool and budget guard');
const total = board.agents.reduce((s,a)=>s+a.pay,0);
ok('pool at base on day one', board.poolInfo.pool===450000, board.poolInfo.pool);
ok('total payout <= live pool', total<=board.poolInfo.pool+0.01, E.rand(total));
ok('Super Club 15/10/8k', (()=>{
  const s=q.slice().sort((a,b)=>b.points-a.points);
  return s[0].superClub===15000 && s[1].superClub===10000 && s[2].superClub===8000;
})());
const champs = board.agents.filter(a=>a.champion>0);
ok('one champion per earning brand with a qualifier', champs.length===new Set(champs.map(c=>c.brand)).size, champs.length);
ok('nobody over the individual cap', board.agents.every(a=>a.pay<=E.CONFIG.pool.individualCap+0.01));

head('10. Ordering — earners first, then closest to qualifying');
const ranked = board.ranked;
const firstZero = ranked.findIndex(a=>a.pay===0);
ok('every earner sits above every R0', ranked.slice(0,firstZero).every(a=>a.pay>0));
ok('earners descend by rand', (()=>{ for(let i=1;i<firstZero;i++) if(ranked[i].pay>ranked[i-1].pay) return false; return true; })());
ok('R0 group descends by closeness', (()=>{
  const z=ranked.slice(firstZero);
  for(let i=1;i<z.length;i++) if(z[i].closeness>z[i-1].closeness+1e-9) return false; return true;
})());
ok('top of the R0 group is the nearest miss', ranked[firstZero].closeness>=ranked[ranked.length-1].closeness);

/* -------------------------------------------------------- 11. ignore flow */
head('11. Ignore');
E.ignoreName('Bradwin Chetty');
ok('added to the ignore list', E.CONFIG.ignoredNames.indexOf(E.normName('Bradwin Chetty'))>-1);
ok('dropped from the unknown list', E.STATE.unknown.indexOf('Bradwin Chetty')===-1);
E.ingestStats(STATS, '2026-08-21', false);
ok('not asked about again on re-upload', E.STATE.unknown.indexOf('Bradwin Chetty')===-1, E.STATE.unknown.join(','));
ok('counted as ignored, not unknown', E.STATE.report.ignored===1, E.STATE.report.ignored);
ok('still never reaches the board', !E.currentBoard().agents.some(a=>a.name==='Bradwin Chetty'));
delete E.STATE.days['2026-08-21'];

/* ----------------------------------------------------------- 12. absence */
head('12. Absence and run rate');
const full = E.STATE.days['2026-08-20'].rows;
E.STATE.days['2026-08-21'] = { date:'2026-08-21',
  rows: full.filter(r => r.key !== E.normName('Jermaine Pillay')) };   // absent on day 2
E.STATE.activeDate = '2026-08-21';
const att = E.attendanceFor(E.normName('Jermaine Pillay'));
ok('2 uploaded days', att.totalDays===2, att.totalDays);
ok('present 1', att.present===1, att.present);
ok('absent 1', att.absent===1, att.absent);
ok('contacts run rate 285/day', near(att.rrContacts,285,.01), att.rrContacts);
ok('missed 285 contacts', near(att.missedContacts,285,.01), att.missedContacts);
ok('missed 53 sales', near(att.missedSales,53,.01), att.missedSales);
const never = E.attendanceFor('nobody at all');
ok('never-present gives no run rate', never.present===0 && never.rrContacts===0);
const b2 = E.currentBoard();
ok('absent agent excluded from the brand average', (()=>{
  const rows = E.STATE.days['2026-08-21'].rows.filter(r=>budgetOrg.indexOf(r.key)>-1);
  return near(b2.brands.BUDGET.avgContacts, rows.reduce((s,r)=>s+r.contacts,0)/rows.length, 0.01);
})(), b2.brands.BUDGET.avgContacts.toFixed(1));
ok('absence does not deflate the gate below the present-only average',
   b2.brands.BUDGET.present === b2.brands.BUDGET.agents.filter(a=>a.contacts>0).length);

/* --------------------------------------------------- 13. everyone appears */
head('13. Everyone on the organogram appears from day one');
E.STATE.days = {}; E.STATE.activeDate = null;
const b3 = E.currentBoard();
ok('27 agents with no stats at all', b3.agents.length===27, b3.agents.length);
ok('all show zero and no data', b3.agents.every(a=>a.contacts===0 && !a.hasData));
ok('nobody earns', b3.agents.every(a=>a.pay===0));
E.ingestStats(STATS, '2026-08-20', false);

/* ------------------------------------------------------ 14. path + rules */
head('14. Path loader and editable rules');
board = E.currentBoard();
const learner = board.agents.find(a=>a.name==='Jermaine Pillay');
const solved = E.solvePath(learner, board, 20000);
ok('a path is found or a ceiling reported', typeof solved.best.pay==='number');
ok('projection clears the contacts gate', solved.best.m.contacts>=Math.ceil(board.brands.BUDGET.gate));
ok('sales never exceed leads', solved.best.m.sales<=solved.best.m.leads);
ok('closing capped at 95%', solved.best.m.closing<=0.95+1e-9);
E.CONFIG.brands.BUDGET.effTarget = 20;
ok('lowering the Budget target adds qualifiers',
   E.currentBoard().brands.BUDGET.qualifiers >= B.BUDGET.qualifiers);
ok('it does not touch FFW', E.currentBoard().brands.FFW.qualifiers===B.FFW.qualifiers);
E.CONFIG.brands.BUDGET.effTarget = 31.6;
E.CONFIG.pool.base = 900000;
ok('base change flows through (cap lifts above base)', E.computePool('2026-08-20').pool===900000, E.computePool('2026-08-20').pool);
E.CONFIG.pool.base = 450000;
ok('budget guard holds after every rules change',
   E.currentBoard().agents.reduce((s,a)=>s+a.pay,0) <= E.currentBoard().poolInfo.pool+0.01);

console.log('\n'+'='.repeat(52)+'\n  '+pass+' passed, '+fail+' failed\n'+'='.repeat(52));
process.exit(fail?1:0);
