const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const fs = require('fs');
const tsv = t => t.split('\n').map(l=>l.split(';').join('\t')).join('\n');
const ORG   = tsv(fs.readFileSync('/home/user/Trust-me-/organogram.csv','utf8'));
const STATS = tsv(fs.readFileSync('/home/user/Trust-me-/sample-real-format.csv','utf8'));
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport:{width:1500,height:1100} });
  const errs=[], reqs=[];
  p.on('console', m => { if(m.type()==='error') errs.push(m.text()); });
  p.on('pageerror', e => errs.push('PAGEERROR: '+e.message));
  p.on('request', r => { if(!r.url().startsWith('file://') && !r.url().startsWith('data:')) reqs.push(r.url()); });
  await p.goto('file:///home/user/Trust-me-/index.html');
  await p.waitForTimeout(400);

  const paste = async text => p.evaluate(t => {
    const dt = new DataTransfer(); dt.setData('text', t);
    document.body.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, text);

  console.log('--- paste the organogram straight in, nothing tidied ---');
  await paste(ORG); await p.waitForTimeout(500);
  console.log('step 1        :', await p.textContent('#step1sub'));

  console.log('\n--- paste the raw export, blank columns and Total rows and all ---');
  await paste(STATS); await p.waitForTimeout(800);
  console.log('mapper opened :', await p.isVisible('#mapModal.on') ? 'YES - BAD' : 'no');
  console.log('step 2        :', await p.textContent('#step2sub'));
  console.log('report        :', (await p.textContent('#reportBox')).replace(/\s+/g,' ').trim().slice(0,215));
  console.log('board rows    :', await p.$$eval('#boardBody tr', r=>r.length));
  console.log('names badge   :', (await p.textContent('#namesBadge')).trim());
  console.log('top earner    :', await p.textContent('#stTop'), '-', await p.textContent('#stTopName'));
  console.log('allocated     :', await p.textContent('#stAlloc'));

  console.log('\n--- headerless paste must be refused ---');
  await paste(STATS.split('\n').slice(1).join('\n')); await p.waitForTimeout(500);
  console.log('toast         :', (await p.textContent('#toast')).trim());
  console.log('board intact  :', await p.$$eval('#boardBody tr', r=>r.length), 'rows still there');

  console.log('\n--- typing into a field is not hijacked ---');
  await p.click('#search'); await p.fill('#search','Jermaine'); await p.waitForTimeout(300);
  console.log('search works  :', await p.$$eval('#boardBody tr', r=>r.length), 'row(s) shown');
  await p.fill('#search','');

  console.log('\n--- the visible Paste button ---');
  await p.click('#btnPasteStats'); await p.waitForTimeout(200);
  console.log('panel opens   :', await p.isVisible('#pastePanel'));
  await p.fill('#pasteBox', STATS);
  await p.click('#btnPasteLoad'); await p.waitForTimeout(700);
  console.log('loaded via box:', await p.$$eval('#boardBody tr', r=>r.length), 'rows · panel closed:', !(await p.isVisible('#pastePanel')));

  await p.screenshot({ path:'shot-paste.png' });
  console.log('\nexternal reqs :', reqs.join(', ') || 'none');
  console.log('console errors:', errs.length ? errs.join(' | ') : 'none');
  await b.close();
})();
