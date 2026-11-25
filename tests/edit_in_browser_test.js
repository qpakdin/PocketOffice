// Simulates the APK's "Edit a copy in the browser": raw upload to /edit/open (exactly what EditSession.java sends), then the browser opens /panel/#/s/<sid>.
const puppeteer = require('/opt/npm-tools/node_modules/puppeteer-core');
const { spawn } = require('child_process'); const fs = require('fs');
(async () => {
  const srv = spawn('python3', ['/home/claude/po/PocketOffice/converter/pocket_office_server.py', '--port', '8842'], { detached: true, stdio: ['ignore', fs.openSync('/tmp/ed/srv_eib.log', 'w'), 'inherit'] });
  for (let i = 0; i < 100; i++) { try { const r = await fetch('http://127.0.0.1:8842/ping'); if (r.ok) break; } catch (e) {} await new Promise(r => setTimeout(r, 300)); }
  const res = []; const ok = (n, c, x) => res.push((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' -> ' + x));
  try {
    const files = [['src.docx', 'docx'], ['x.xlsx', 'xlsx'], ['p.pptx', 'pptx']];
    const browser = await puppeteer.launch({ executablePath: '/opt/pw-browsers/chromium-1243/chrome-linux64/chrome', headless: true, args: ['--no-sandbox', '--disable-gpu'] });
    for (const [f, ext] of files) {
      const data = fs.readFileSync('/tmp/ed/' + f);
      const r = await fetch('http://127.0.0.1:8842/edit/open?ext=' + ext + '&name=' + encodeURIComponent(f), { method: 'POST', body: data, headers: { 'Content-Type': 'application/octet-stream' } });
      const j = await r.json(); ok('upload ' + f, j.ok && /^[a-f0-9]+$/.test(j.sid), JSON.stringify(j).slice(0, 200));
      const page = await browser.newPage(); await page.setViewport({ width: 390, height: 800, isMobile: true, hasTouch: true });
      const errs = []; page.on('pageerror', e => errs.push(e.message)); page.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });
      await page.goto('http://127.0.0.1:8842/panel/#/s/' + j.sid, { waitUntil: 'networkidle0' }); await new Promise(r => setTimeout(r, 1200));
      const title = await page.$eval('header h1', e => e.textContent).catch(() => '');
      ok('panel opens session for ' + f, title.includes(f), title);
      ok('no browser errors for ' + f, errs.length === 0, errs.join(' | '));
      await page.close();
    }
    // a protected / unsupported format must be refused with a readable message
    const bad = await fetch('http://127.0.0.1:8842/edit/open?ext=doc&name=a.doc', { method: 'POST', body: Buffer.from('x') }); const bj = await bad.json();
    ok('unsupported format refused with message', bj.ok === false && /Editing is available/.test(bj.error), JSON.stringify(bj).slice(0, 150));
    // the original upload is never modified: closing removes the session
    await browser.close();
  } catch (e) { res.push('EXCEPTION ' + e.message); }
  console.log(res.join('\n')); process.kill(-srv.pid);
})();
