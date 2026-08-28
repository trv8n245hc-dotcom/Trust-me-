const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const fs = require('fs');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport:{width:1500,height:1150} });
  const errs=[], reqs=[];
  p.on('console', m => { if(m.type()==='error') errs.push(m.text()); });
  p.on('pageerror', e => errs.push('PAGEERROR: '+e.message));
  p.on('request', r => { if(!r.url().startsWith('file://') && !r.url().startsWith('data:')) reqs.push(r.url()); });
  await p.goto('file:///home/user/Trust-me-/index.html');
  await p.waitForTimeout(400);

  console.log('empty state    :', (await p.textContent('#boardEmpty')).replace(/\s+/g,' ').trim().slice(0,90));
  console.log('step 1         :', await p.textContent('#step1sub'));
  console.log('step 2         :', await p.textContent('#step2sub'));
  console.log('sample banner  :', (await p.textContent('#sampleBanner')).trim() || '(none — correct, does not auto-load)');

  // upload the two real files through the real inputs
  await p.setInputFiles('#fileOrg', '/home/user/Trust-me-/organogram.csv');
  await p.waitForTimeout(400);
  console.log('\nafter organogram:', await p.textContent('#step1sub'));
  await p.setInputFiles('#fileStats', '/home/user/Trust-me-/sample-real-format.csv');
  await p.waitForTimeout(700);

  console.log('mapper opened? :', await p.isVisible('#mapModal.on') ? 'YES - BAD' : 'no (correct)');
  console.log('report         :', (await p.textContent('#reportBox')).replace(/\s+/g,' ').trim().slice(0,230));
  console.log('board rows     :', await p.$$eval('#boardBody tr', r=>r.length));
  console.log('names badge    :', (await p.textContent('#namesBadge')).trim());
  console.log('pool           :', await p.textContent('#potValue'));
  console.log('top earner     :', await p.textContent('#stTop'), '-', await p.textContent('#stTopName'));

  const rows = await p.$$eval('#boardBody tr', rs => rs.map(r => {
    const td = r.querySelectorAll('td');
    return { name: td[1].querySelector('.agent').textContent.trim(),
             earn: td[2].textContent.trim(), brand: td[3].textContent.trim() };
  }));
  console.log('\ntop 5 by earnings:');
  rows.slice(0,5).forEach((r,i)=>console.log('  '+(i+1)+'. '+r.name.padEnd(26)+r.earn.padStart(11)+'  '+r.brand));
  console.log('bottom 3:');
  rows.slice(-3).forEach(r=>console.log('     '+r.name.padEnd(26)+r.earn.padStart(11)+'  '+r.brand));
  const names = rows.map(r=>r.name);
  console.log('\nKimisha team on board?', names.some(n=>['Bradwin Chetty','Pam Bulose','Ronnel Naidoo'].includes(n)) ? 'YES - BAD' : 'no (correct)');
  console.log('Siobhan shown by official name?', names.includes('Siobhan Nieuwenhuis - JHB Agent') ? 'yes' : 'NO - BAD');
  console.log('sorted highest first?', rows[0].earn === rows.map(r=>r.earn).sort((a,b)=>parseFloat(b.replace(/\D/g,''))-parseFloat(a.replace(/\D/g,'')))[0] ? 'yes' : 'check');

  // names tab
  await p.click('.tab[data-view="names"]'); await p.waitForTimeout(300);
  console.log('\nunmatched listed:', await p.$$eval('#namesList .nrow', r=>r.length));
  console.log('first unmatched :', (await p.textContent('#namesList .nrow')).replace(/\s+/g,' ').trim().slice(0,110));
  await p.click('#namesList .nrow button[data-ignore]'); await p.waitForTimeout(400);
  console.log('after Ignore    :', await p.$$eval('#namesList .nrow', r=>r.length), 'left ·',
              await p.$$eval('#ignoredList .nrow', r=>r.length), 'ignored');
  await p.click('#btnIgnoreAll'); await p.waitForTimeout(400);
  console.log('after Ignore all:', (await p.textContent('#namesList')).replace(/\s+/g,' ').trim().slice(0,60));
  console.log('badge cleared   :', (await p.textContent('#namesBadge')).trim() === '' ? 'yes' : 'no');

  // agent card
  await p.click('.tab[data-view="arena"]'); await p.waitForTimeout(200);
  await p.click('#boardBody tr:first-child .agent'); await p.waitForTimeout(400);
  console.log('\ncard open      :', await p.isVisible('#cardModal.on'));
  const card = (await p.textContent('#cardSheet')).replace(/\s+/g,' ').trim();
  console.log('card           :', card.slice(0,200));
  console.log('has absence    :', /Days present/.test(card) ? 'yes' : 'NO');
  await p.click('#cardSheet .x'); await p.waitForTimeout(250);
  console.log('card closes (X):', await p.isVisible('#cardModal.on') ? 'NO - BAD' : 'yes');
  await p.click('#boardBody tr:first-child .agent'); await p.waitForTimeout(300);
  await p.keyboard.press('Escape'); await p.waitForTimeout(250);
  console.log('card closes (Esc):', await p.isVisible('#cardModal.on') ? 'NO - BAD' : 'yes');

  await p.screenshot({ path:'shot-final.png' });

  // brands
  await p.click('.tab[data-view="brands"]'); await p.waitForTimeout(300);
  console.log('\nbrand cards    :', await p.$$eval('#brandCards .bcard', c=>c.length));
  console.log('budget card    :', (await p.textContent('#brandCards .bcard')).replace(/\s+/g,' ').trim().slice(0,170));

  // my path
  await p.click('.tab[data-view="path"]'); await p.waitForTimeout(400);
  console.log('\ncoach          :', (await p.textContent('.readout')).replace(/\s+/g,' ').trim().slice(0,170));
  await p.fill('#pathTarget','25000'); await p.click('#btnLoadPath'); await p.waitForTimeout(800);
  console.log('projected      :', await p.textContent('.bignum'));
  console.log('note           :', (await p.textContent('#view-path .pot-sub')).replace(/\s+/g,' ').trim().slice(0,150));

  // persistence
  await p.reload(); await p.waitForTimeout(600);
  console.log('\nafter reload   : org', await p.textContent('#step1sub'), '| rows', await p.$$eval('#boardBody tr', r=>r.length));

  // mobile
  await p.setViewportSize({width:390,height:844}); await p.waitForTimeout(300);
  console.log('mobile overflow:', await p.evaluate(()=>document.documentElement.scrollWidth-document.documentElement.clientWidth), 'px');

  console.log('\nexternal reqs  :', reqs.join(', ') || 'none');
  console.log('console errors :', errs.length ? errs.join(' | ') : 'none');

  await b.close();
})();
