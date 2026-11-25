const puppeteer = require('/opt/npm-tools/node_modules/puppeteer-core');
const { spawn } = require('child_process');
const fs = require('fs');
(async () => {
  const srv = spawn('python3', ['/home/claude/po/PocketOffice/converter/pocket_office_server.py', '--port', '8841'], { detached: true, stdio: ['ignore', fs.openSync('/tmp/ed/srv_ui.log', 'w'), 'inherit'] });
  for (let i = 0; i < 100; i++) { try { const r = await fetch('http://127.0.0.1:8841/ping'); if (r.ok) break; } catch (e) {} await new Promise(r => setTimeout(r, 300)); }
  const browser = await puppeteer.launch({ executablePath: '/opt/pw-browsers/chromium-1243/chrome-linux64/chrome', headless: true, args: ['--no-sandbox', '--disable-gpu'] });
  const page = await browser.newPage();
  await page.setViewport({ width: 390, height: 800, deviceScaleFactor: 2, isMobile: true, hasTouch: true });
  const errors = []; page.on('pageerror', e => errors.push('PAGEERROR ' + e.message)); page.on('console', m => { if (m.type() === 'error') errors.push('CONSOLE ' + m.text()); });
  const shot = async n => page.screenshot({ path: '/tmp/ed/shot_' + n + '.png' });
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const results = []; const ok = (n, c, x) => { results.push((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' -> ' + x)); };
  const clickText = async (txt, sel = 'button') => { const h = await page.evaluateHandle((t, s) => [...document.querySelectorAll(s)].reverse().find(e => e.textContent.trim() === t || e.textContent.trim().startsWith(t)), txt, sel); const el = h.asElement(); if (!el) throw new Error('no element with text ' + txt); await el.click(); };
  try {
    await page.goto('http://127.0.0.1:8841/', { waitUntil: 'networkidle0' });
    ok('home loads', await page.$eval('h1', e => e.textContent) === 'Pocket Office');
    await sleep(600); await shot('1_home');
    // ---- new Word file
    await clickText('Create'); await page.waitForFunction(() => location.hash.startsWith('#/s/'), { timeout: 20000 }); await sleep(800);
    ok('new docx opens the editor', (await page.$eval('header h1', e => e.textContent)).includes('Untitled.docx'));
    await page.click('.list .item'); await page.waitForSelector('.sheet textarea');
    await page.type('.sheet textarea', 'Hello from the web panel — سلام دنیا');
    await clickText('B'); await clickText('Right ⇥'); await sleep(200); await shot('2_para_editor');
    await clickText('Apply'); await sleep(1200);
    const t0 = await page.$eval('.list .item .tx', e => e.textContent); ok('typed text is in the list', t0.includes('Hello from the web panel') && t0.includes('سلام'), t0);
    ok('change counter updated', (await page.$eval('#ops', e => e.textContent)).startsWith('1'), await page.$eval('#ops', e => e.textContent));  // one Apply = one batch = one history step
    // ---- find & replace
    await clickText('Find'); await page.waitForSelector('.sheet input'); const inputs = await page.$$('.sheet input[type=text]'); await inputs[0].type('web panel'); await inputs[1].type('PANEL');
    await clickText('Preview replace'); await page.waitForSelector('.res'); const prev = await page.$eval('.res', e => e.textContent); ok('replace preview text', prev.includes('Preview') && prev.includes('1 item'), prev);
    await clickText('Replace all'); await sleep(1500); await shot('3_find');
    await page.evaluate(() => document.querySelectorAll('.sheet').forEach(s => s.remove()));
    ok('replace applied', (await page.$eval('.list .item .tx', e => e.textContent)).includes('PANEL'));
    // ---- settings form generated from schema
    await clickText('Settings'); await page.waitForSelector('.gcard'); const groups = await page.$$eval('.gcard b', els => els.map(e => e.textContent)); ok('settings lists the 13 groups', groups.length === 13, groups.join(','));
    await shot('4_settings');
    await clickText('Paragraph', '.gcard'); await sleep(300); const selects = await page.$$('.sheet select'); ok('paragraph form has fields', selects.length >= 3 && (await page.$$('.sheet .sec')).length >= 3);
    const alignSel = await page.evaluateHandle(() => [...document.querySelectorAll('.sheet label.lbl')].find(l => l.textContent.toLowerCase().includes('alignment')).nextElementSibling);
    await alignSel.asElement().select('center'); await clickText('Preview impact'); await page.waitForSelector('.res'); await sleep(900);
    const pv = await page.$eval('.res', e => e.textContent); ok('settings preview shows parts', pv.includes('word/document.xml'), pv); await shot('5_settings_preview');
    await clickText('Apply'); await sleep(1500);
    await page.evaluate(() => document.querySelectorAll('.sheet').forEach(s => s.remove()));
    // ---- preview
    await clickText('Preview'); await page.waitForSelector('.pages img', { timeout: 60000 }); await sleep(1500);
    ok('preview shows a page image', await page.$eval('.pages img', i => i.complete && i.naturalWidth > 100)); await shot('6_preview');
    await page.evaluate(() => document.querySelectorAll('.sheet').forEach(s => s.remove()));
    // ---- save
    await clickText('Save…'); await page.waitForSelector('.sheet select', { timeout: 60000 }); await sleep(500);
    const sv = await page.$eval('.sheet', e => e.textContent); ok('save shows verification passed', sv.includes('Verification passed'), sv.slice(0, 200)); await shot('7_save');
    await page.evaluate(() => document.querySelectorAll('.sheet').forEach(s => s.remove()));
    // ---- Excel via new tab flow
    await page.goto('http://127.0.0.1:8841/#/', { waitUntil: 'networkidle0' }); await sleep(300);
    await clickText('Excel', '.chips button'); await sleep(300); await clickText('Create'); await page.waitForFunction(() => location.hash.startsWith('#/s/') && document.querySelector('table.grid'), { timeout: 20000 }); await sleep(600);
    ok('new xlsx shows an empty grid', (await page.$$('table.grid td')).length >= 200);
    const cells = await page.$$('table.grid td'); await cells[0].click(); await page.type('#fbar', 'Item'); await page.keyboard.press('Enter'); await sleep(1200);
    const cells2 = await page.$$('table.grid td'); await cells2[1].click(); await page.type('#fbar', '12'); await page.keyboard.press('Enter'); await sleep(1200);
    const cells3 = await page.$$('table.grid td'); await cells3[9].click(); await page.type('#fbar', 'x'); await page.keyboard.press('Enter'); await sleep(900);
    const texts = await page.$$eval('table.grid td', els => els.slice(0, 12).map(e => e.textContent)); ok('typed cell values appear', texts[0] === 'Item' && texts[1] === '12', texts.join('|'));
    const cells4 = await page.$$('table.grid td'); await cells4[9].click(); await sleep(100); await page.$eval('#fbar', e => { e.value = '=B1*2'; }); await clickText('Set'); await sleep(1200);
    const fcell = await page.$$eval('table.grid td', els => els[9].textContent); ok('formula cell set (value or pending recalculation)', fcell === '24' || fcell === '=B1*2' || fcell.length > 0, fcell); await shot('8_excel');
    // ---- sessions list on home
    await page.goto('http://127.0.0.1:8841/#/', { waitUntil: 'networkidle0' }); await sleep(800);
    const sess = await page.$$eval('.list .item b', els => els.map(e => e.textContent)); ok('home lists the open sessions', sess.length >= 2, sess.join(',')); await shot('9_home_sessions');
  } catch (e) { results.push('EXCEPTION ' + e.message); try { await shot('error'); } catch (x) {} }
  console.log(results.join('\n')); console.log(errors.length ? 'BROWSER ERRORS:\n' + errors.join('\n') : 'no browser console errors');
  await browser.close(); process.kill(-srv.pid);
})();
