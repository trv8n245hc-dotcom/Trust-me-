/* Reproduces the OneDrive / SharePoint preview: the page runs in a sandboxed
   iframe with no allow-same-origin, where touching localStorage THROWS. */
const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const fs = require('fs');
(async () => {
  fs.writeFileSync('/home/user/Trust-me-/__sandbox_probe.html',
    '<body style="margin:0"><iframe sandbox="allow-scripts" src="index.html" ' +
    'style="width:1400px;height:1000px;border:0"></iframe>');
  const b = await chromium.launch();
  const p = await b.newPage({ viewport:{width:1440,height:1000} });
  const errs = [];
  p.on('pageerror', e => errs.push('PAGEERROR: ' + e.message));
  p.on('console', m => { if(m.type()==='error') errs.push(m.text()); });
  await p.goto('file:///home/user/Trust-me-/__sandbox_probe.html');
  await p.waitForTimeout(1500);
  const fr = p.frames().find(f => f !== p.mainFrame());
  if(!fr){ console.log('no frame'); await b.close(); return; }
  console.log('storage blocked :', await fr.evaluate(() => { try{ window.localStorage.getItem('x'); return 'NO'; }catch(e){ return 'yes — ' + e.name; } }));
  console.log('page drew       :', await fr.textContent('#potValue'));
  console.log('banner shown    :', (await fr.textContent('#storeBanner')).replace(/\s+/g,' ').trim().slice(0,110));
  console.log('empty state ok  :', (await fr.textContent('#boardEmpty')).replace(/\s+/g,' ').trim().slice(0,60));
  console.log('tabs present    :', await fr.$$eval('.tab', t=>t.length));
  console.log('console errors  :', errs.length ? errs.join(' | ') : 'none');
  await p.screenshot({ path:'shot-sandbox.png' });
  await b.close();
  fs.unlinkSync('/home/user/Trust-me-/__sandbox_probe.html');
})();
