/* Pocket Office web panel: editor for Word / Excel / PowerPoint copies. Every browser tab is its own window (its own session). */
(function () {
  'use strict';
  var app = document.getElementById('app');
  var S = null;                       // current editor state
  var FONTS = null;

  // ------------------------------------------------------------------ helpers
  function h(tag, a) {
    var e = document.createElement(tag);
    if (a) for (var k in a) {
      if (k === 'class') e.className = a[k]; else if (k === 'text') e.textContent = a[k]; else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), a[k]);
      else if (k === 'html') e.innerHTML = a[k]; else if (a[k] !== false && a[k] !== null && a[k] !== undefined) e.setAttribute(k, a[k] === true ? '' : a[k]);
    }
    function add(c) {
      if (c === null || c === undefined || c === false) return;
      if (Array.isArray(c)) { c.forEach(add); return; }
      e.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
    }
    for (var i = 2; i < arguments.length; i++) add(arguments[i]);
    return e;
  }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  var toastT;
  function toast(msg) { var t = document.getElementById('toast'); t.textContent = msg; t.className = 'show'; clearTimeout(toastT); toastT = setTimeout(function () { t.className = ''; }, 2600); }
  function busy(msg) { var d = h('div', { class: 'sheet' }, h('div', { class: 'box' }, h('h3', {}, h('span', { text: msg || 'Working…' }), h('i', { class: 'spin', style: 'border-color:#1f5fbf55;border-top-color:#1f5fbf' })))); document.body.appendChild(d); return function () { d.remove(); }; }
  function modal(title, body, buttons) {
    var close = function () { m.remove(); };
    var m = h('div', { class: 'sheet', onclick: function (e) { if (e.target === m) close(); } },
      h('div', { class: 'box' }, h('h3', {}, h('span', { text: title }), h('button', { text: '✕', onclick: close, 'aria-label': 'Close' })), body, buttons ? h('div', { class: 'row', style: 'margin-top:14px;justify-content:flex-end' }, buttons(close)) : null));
    document.body.appendChild(m); return { close: close, el: m };
  }
  function fail(e) { modal('Not done', h('div', { text: (e && e.message ? e.message : String(e)) + '\n\nNothing was changed.', style: 'white-space:pre-wrap' })); }
  function api(method, path, body, raw) {
    var opt = { method: method, headers: {} };
    if (body !== undefined && body !== null) { if (body instanceof Blob || body instanceof ArrayBuffer) { opt.body = body; } else { opt.body = JSON.stringify(body); opt.headers['Content-Type'] = 'application/json'; } }
    return fetch(path, opt).then(function (r) {
      if (raw) return r;
      return r.text().then(function (t) {
        var j; try { j = JSON.parse(t); } catch (e) { throw new Error('The server answered with an unexpected reply (HTTP ' + r.status + '): ' + t.slice(0, 200)); }
        if (!j.ok) throw new Error(j.error || 'The server refused the request.');
        return j;
      });
    });
  }
  function q(o) { return Object.keys(o).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(o[k]); }).join('&'); }
  function ext(name) { var m = /\.([A-Za-z0-9]{1,6})$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
  function cp(text, i) { return Array.from(text.slice(0, i)).length; }   // UTF-16 index -> code point index

  // ------------------------------------------------------------------ routing
  function route() {
    var m = /^#\/s\/([a-f0-9]+)/.exec(location.hash);
    if (m) editor(m[1]); else home();
  }
  window.addEventListener('hashchange', route);

  // ------------------------------------------------------------------ home
  function home() {
    S = null;
    app.innerHTML = '';
    app.appendChild(h('header', { class: 'top' }, h('h1', { text: 'Pocket Office' })));
    var main = h('main'); app.appendChild(main);
    var info = h('div', { class: 'mut', text: 'Connecting to the converter…' });
    api('GET', '/ping', null, true).then(function (r) { return r.text(); }).then(function (t) { info.textContent = t.replace(/\|/g, '·'); }).catch(function () { info.textContent = 'The converter did not answer.'; });

    // new file
    var cat = null, grp = 'Word', fmt = 'docx';
    var chips = h('div', { class: 'chips' }), sel = h('select'), name = h('input', { type: 'text', value: 'Untitled', 'aria-label': 'File name' }), note = h('div', { class: 'hint' });
    function fillFormats() {
      sel.innerHTML = ''; (cat ? cat.groups[grp] : []).forEach(function (e) { sel.appendChild(h('option', { value: e, text: '.' + e + '  —  ' + (cat.labels[e] || '') })); });
      sel.value = fmt = (cat && cat.groups[grp].indexOf(fmt) >= 0) ? fmt : (cat ? cat.groups[grp][0] : '');
      var nat = cat && cat.edit_as[sel.value]; note.textContent = nat ? 'Edited as .' + nat + ' and converted to .' + sel.value + ' by LibreOffice when you save.' : '';
    }
    function fillChips() { chips.innerHTML = ''; ['Word', 'Excel', 'PowerPoint'].forEach(function (g) { chips.appendChild(h('button', { text: g, class: g === grp ? 'on' : '', onclick: function () { grp = g; fillChips(); fillFormats(); } })); }); }
    sel.addEventListener('change', function () { fmt = sel.value; fillFormats(); });
    api('GET', '/edit/catalog').then(function (c) { cat = c; fillChips(); fillFormats(); }).catch(function (e) { note.style.whiteSpace = 'pre-wrap'; note.style.color = '#ff8a80'; note.style.fontWeight = '600'; note.textContent = e.message; });
    fillChips();
    main.appendChild(h('div', { class: 'card' }, h('h2', { text: 'New file' }), h('div', { class: 'mut', text: 'A valid empty file in the format you choose. Type into it, import content, then save.' }), chips, sel, note,
      h('div', { class: 'row', style: 'margin-top:8px' }, name, h('button', { class: 'pri', text: 'Create', onclick: function () {
        var done = busy('Creating…');
        api('POST', '/edit/new?' + q({ ext: sel.value, name: (name.value.trim() || 'Untitled') + '.' + sel.value }), {}).then(function (j) { done(); S0 = j; location.hash = '#/s/' + j.sid; }).catch(function (e) { done(); fail(e); });
      } }))));
    // open
    var file = h('input', { type: 'file', accept: '.docx,.docm,.dotx,.dotm,.xlsx,.xlsm,.xltx,.xltm,.pptx,.pptm,.ppsx,.ppsm,.potx,.potm' });
    main.appendChild(h('div', { class: 'card' }, h('h2', { text: 'Edit a file' }), h('div', { class: 'mut', text: 'Word, Excel and PowerPoint (OOXML). The original is only read; you edit a copy and save it as a new file.' }),
      h('div', { class: 'row', style: 'margin-top:8px' }, file, h('button', { class: 'pri', text: 'Open for editing', onclick: function () {
        var f = file.files[0]; if (!f) { toast('Choose a file first'); return; }
        var done = busy('Analyzing the file…');
        api('POST', '/edit/open?' + q({ ext: ext(f.name), name: f.name }), f).then(function (j) { done(); S0 = j; location.hash = '#/s/' + j.sid; }).catch(function (e) { done(); fail(e); });
      } }))));
    // sessions
    var box = h('div', { class: 'list' });
    function load() {
      api('GET', '/edit/sessions').then(function (j) {
        box.innerHTML = '';
        if (!j.sessions.length) box.appendChild(h('div', { class: 'mut', text: 'No open sessions.' }));
        j.sessions.forEach(function (s) {
          box.appendChild(h('div', { class: 'item' }, h('div', { class: 'row' }, h('div', {}, h('b', { text: s.name }), h('div', { class: 'mut', text: s.kind.toUpperCase() + ' · ' + s.changes + ' change(s)' + (s.new ? ' · new' : '') })),
            h('a', { href: '#/s/' + s.sid, target: '_blank' }, h('button', { text: 'Open in new tab' })), h('a', { href: '#/s/' + s.sid }, h('button', { class: 'pri', text: 'Open' })),
            h('button', { class: 'bad', text: 'Close', onclick: function () { api('POST', '/edit/close?sid=' + s.sid, {}).then(load); } }))));
        });
      }).catch(function () { });
    }
    load();
    main.appendChild(h('div', { class: 'card' }, h('h2', { text: 'Open sessions' }), h('div', { class: 'mut', text: 'Every browser tab is an independent window: open two sessions in two tabs (or split screen) and copy text between them.' }), box));
    main.appendChild(info);
  }
  var S0 = null;

  // ------------------------------------------------------------------ editor
  function editor(sid) {
    app.innerHTML = '';
    var done = busy('Loading…');
    var p = (S0 && S0.sid === sid) ? Promise.resolve(S0) : api('GET', '/edit/open_payload?sid=' + sid);
    var first = S0 && S0.sid === sid; S0 = null;
    p.then(function (j) {
      done(); S = { sid: sid, kind: j.kind, name: j.name, meta: j.meta, blocks: j.blocks || [], tables: j.tables || [], page: j.page || [], sheets: j.sheets || [], ops: j.ops || 0, report: j.report,
                    sheet: 0, r0: 1, c0: 1, sel: null, shown: 150, origPages: 0 };
      draw();
      if (first && j.report && !j.meta.new) showReport(true);
      else if (first && j.meta.new) toast('New empty ' + (j.meta.save_ext) + ' — start typing');
    }).catch(function (e) { done(); app.innerHTML = ''; app.appendChild(h('main', {}, h('div', { class: 'card' }, h('h2', { text: 'Session not found' }), h('div', { text: e.message }), h('p', {}, h('a', { href: '#/' }, h('button', { class: 'pri', text: 'Home' })))))); });
  }

  function draw() {
    app.innerHTML = '';
    var kindName = { docx: 'Word', xlsx: 'Excel', pptx: 'PowerPoint' }[S.kind];
    app.appendChild(h('header', { class: 'top' },
      h('button', { text: '← Home', onclick: function () { location.hash = '#/'; } }),
      h('h1', {}, S.name, h('span', { class: 'badge', text: kindName + ' · copy' })),
      h('span', { class: 'badge', id: 'ops', text: S.ops + ' change(s)' }),
      h('button', { text: '↶ Undo', onclick: undo }), h('button', { text: 'Find', onclick: findPanel }), h('button', { text: 'Settings', onclick: function () { settingsPanel(S.kind === 'xlsx' && S.sel ? { group: 'format', values: { sheets: S.sheets[S.sheet], range: S.sel.ref } } : null); } }),
      h('button', { text: 'Preview', onclick: preview }), h('button', { class: 'pri', text: 'Save…', onclick: savePanel }), h('button', { text: 'Report', onclick: function () { showReport(false); } }),
      h('a', { href: location.pathname + '#/', target: '_blank' }, h('button', { text: '＋ Tab' }))));
    var main = h('main', { id: 'main' }); app.appendChild(main);
    if (S.kind === 'xlsx') gridView(main); else listView(main);
  }
  function refreshOps() { var o = document.getElementById('ops'); if (o) o.textContent = S.ops + ' change(s)'; }
  function mergeBlocks(j) {
    if (j.blocks) { if (j.structure_changed || j.dry === false || j.blocks.length > 50) S.blocks = j.blocks; else j.blocks.forEach(function (b) { if (S.blocks[b.i]) S.blocks[b.i] = b; }); }
    if (j.tables) S.tables = j.tables; if (j.page) S.page = j.page; if (typeof j.ops === 'number') S.ops = j.ops;
  }

  // ---- paragraph list (Word / PowerPoint)
  function listView(main) {
    var list = h('div', { class: 'card list' });
    main.appendChild(h('div', { class: 'mut', style: 'margin:2px 4px 8px', text: 'Tap a paragraph to edit it. ⋯ = more actions. Locked paragraphs (pictures, fields, shapes …) are kept exactly as they are.' }));
    main.appendChild(list);
    renderList(list);
  }
  function renderList(list) {
    list.innerHTML = '';
    S.blocks.slice(0, S.shown).forEach(function (b) {
      var head = b.where ? b.where : (b.kind === 'cell' ? 'Table ' + (b.t + 1) + ' · row ' + (b.r + 1) + ' · col ' + (b.c + 1) : b.kind === 'textbox' ? 'Text box' : '¶ ' + (b.i + 1));
      if (b.style && b.style !== 'null') head += ' · ' + b.style;
      var st = 'text-align:' + (b.align === 'center' ? 'center' : b.align === 'right' ? 'right' : b.align === 'justify' ? 'justify' : 'left') + ';' + (b.bold ? 'font-weight:700;' : '') + (b.italic ? 'font-style:italic;' : '') + (b.underline ? 'text-decoration:underline;' : '');
      var tx = h('div', { class: 'tx' + (b.text ? '' : ' empty'), dir: 'auto', style: b.locked ? '' : st, text: b.text || '(empty paragraph)' });
      var item = h('div', { class: 'item' + (b.locked ? ' locked' : ''), onclick: function () { if (b.locked) lockedInfo(b); else paraEditor(b); } },
        h('div', { class: 'hd' }, h('span', { text: head + (b.locked ? '  🔒 ' + b.reason : '') }), h('span', { class: 'sp' }), h('button', { class: 'menu', text: '⋯', onclick: function (e) { e.stopPropagation(); rowMenu(b); } })), tx);
      list.appendChild(item);
    });
    if (S.blocks.length > S.shown) list.appendChild(h('div', { class: 'item', style: 'text-align:center;color:#1f5fbf', text: 'Show more (' + (S.blocks.length - S.shown) + ' more paragraphs)', onclick: function () { S.shown += 300; renderList(list); } }));
    if (!S.blocks.length) list.appendChild(h('div', { class: 'item mut', text: 'This file has no editable text paragraphs.' }));
  }
  function lockedInfo(b) { modal('Locked paragraph', h('div', { text: 'This paragraph contains ' + b.reason + '.\n\nIt is kept exactly as it is, so nothing in it can be damaged. You can still insert a paragraph next to it (⋯).', style: 'white-space:pre-wrap' })); }
  function send(ops, okMsg) {
    var d = busy('Applying…');
    return api('POST', '/edit/apply?sid=' + S.sid, { ops: ops }).then(function (j) { d(); mergeBlocks(j); refreshOps(); draw2(); if (okMsg) toast(okMsg); return j; }).catch(function (e) { d(); fail(e); throw e; });
  }
  function draw2() { var l = document.querySelector('.list'); if (l && S.kind !== 'xlsx') renderList(l); if (S.kind === 'xlsx') loadGrid(); }
  function rowMenu(b) {
    var items = [];
    if (b.kind !== 'textbox') items.push(['Insert a paragraph after this one', function () { send([{ op: 'insert_paragraph_after', p: b.i }]); }]);
    if (!b.locked && b.kind !== 'textbox') items.push(['Delete this paragraph', function () { if (confirm('Delete this paragraph?')) send([{ op: 'delete_paragraph', p: b.i }]); }]);
    if (b.kind === 'cell') {
      items.push(['Insert table row above', function () { send([{ op: 'insert_row', t: b.t, r: b.r, where: 'above' }]); }]);
      items.push(['Insert table row below', function () { send([{ op: 'insert_row', t: b.t, r: b.r, where: 'below' }]); }]);
      items.push(['Delete this table row', function () { if (confirm('Delete row ' + (b.r + 1) + ' of table ' + (b.t + 1) + '?')) send([{ op: 'delete_row', t: b.t, r: b.r }]); }]);
    }
    items.push(['Copy text', function () { (navigator.clipboard ? navigator.clipboard.writeText(b.text) : Promise.reject()).then(function () { toast('Copied'); }).catch(function () { prompt('Copy this text:', b.text); }); }]);
    var m = modal('¶ ' + (b.i + 1), h('div', {}, items.map(function (it) { return h('button', { style: 'display:block;width:100%;text-align:left;margin:6px 0', text: it[0], onclick: function () { m.close(); it[1](); } }); })));
  }

  // ---- paragraph editor
  function paraEditor(b) {
    var t = {}; function touch(k) { return function () { t[k] = true; }; }
    var ta = h('textarea', { dir: 'auto', rows: 4 }); ta.value = b.text;
    var font = h('input', { type: 'text', list: 'fontlist', placeholder: 'Font' + (b.font ? ' (now: ' + b.font + ')' : ''), oninput: touch('font') });
    var dl = h('datalist', { id: 'fontlist' }); fonts().then(function (fl) { fl.slice(0, 2500).forEach(function (f) { dl.appendChild(h('option', { value: f })); }); });
    var size = h('input', { type: 'number', step: '0.5', min: '1', placeholder: b.size ? 'Size pt (now ' + b.size + ')' : 'Size pt', oninput: touch('size') });
    function tog(label, key, on) { var bt = h('button', { text: label, class: on ? 'on' : '', style: 'min-width:44px', onclick: function () { t[key] = true; bt.classList.toggle('on'); } }); return bt; }
    var B = tog('B', 'bold', b.bold), I = tog('I', 'italic', b.italic), U = tog('U', 'underline', b.underline);
    var color = h('input', { type: 'color', value: '#' + (b.color || '000000'), style: 'width:56px;padding:2px', oninput: touch('color') });
    var al = null; var alBtns = ['left', 'center', 'right', 'justify'].map(function (a) { var bt = h('button', { text: { left: '⇤ Left', center: '⇔ Center', right: 'Right ⇥', justify: '☰ Justify' }[a], class: b.align === a ? 'on' : '', onclick: function () { al = a; t.para = true; alBtns.forEach(function (x) { x.classList.remove('on'); }); bt.classList.add('on'); } }); return bt; });
    var rtl = h('button', { text: b.rtl ? 'Right-to-left ✓' : 'Left-to-right', class: b.rtl ? 'on' : '', onclick: function () { t.rtl = true; t.para = true; rtl.classList.toggle('on'); rtl.textContent = rtl.classList.contains('on') ? 'Right-to-left ✓' : 'Left-to-right'; } });
    var line = h('select', { onchange: touch('para') }, h('option', { value: '', text: 'Line spacing: keep' }), ['1', '1.15', '1.5', '2', '2.5', '3'].map(function (v) { return h('option', { value: v, text: 'Line spacing: ' + v }); }));
    function num(ph) { return h('input', { type: 'number', step: '0.1', placeholder: ph, oninput: touch('para') }); }
    var bef = num('Before (pt)'), aft = num('After (pt)'), left = num('Indent (cm)'), first = num('1st line (cm)');
    var body = h('div', {}, ta, h('div', { class: 'hint', text: 'Select part of the text to format only that part; with nothing selected the whole paragraph is formatted.' }),
      h('label', { class: 'lbl', text: 'Font' }), font, dl, h('div', { class: 'row', style: 'margin-top:8px' }, size, B, I, U, color),
      h('label', { class: 'lbl', text: 'Paragraph' }), h('div', { class: 'chips' }, alBtns, rtl), line, h('div', { class: 'row', style: 'margin-top:8px' }, bef, aft, left, first));
    modal('Edit ' + (b.where || ('¶ ' + (b.i + 1))), body, function (close) {
      return [h('button', { text: 'Cancel', onclick: close }), h('button', { class: 'pri', text: 'Apply', onclick: function () {
        var ops = [], text = ta.value;
        if (text !== b.text) ops.push({ op: 'set_text', p: b.i, text: text });
        if (t.font || t.size || t.bold || t.italic || t.underline || t.color) {
          var f = { op: 'format', p: b.i }, a = ta.selectionStart, z = ta.selectionEnd;
          if (z > a && !(a === 0 && z === text.length)) { f.start = cp(text, a); f.end = cp(text, z); }
          if (t.font && font.value.trim()) f.font = font.value.trim(); if (t.size && size.value) f.size = parseFloat(size.value);
          if (t.bold) f.bold = B.classList.contains('on'); if (t.italic) f.italic = I.classList.contains('on'); if (t.underline) f.underline = U.classList.contains('on');
          if (t.color) f.color = color.value.slice(1).toUpperCase();
          if (Object.keys(f).length > 2) ops.push(f);
        }
        if (t.para) {
          var p = { op: 'para', p: b.i }; if (al) p.align = al; if (t.rtl) p.rtl = rtl.classList.contains('on'); if (line.value) p.line = parseFloat(line.value);
          if (bef.value) p.before = parseFloat(bef.value); if (aft.value) p.after = parseFloat(aft.value); if (left.value) p.left = parseFloat(left.value); if (first.value) p.first = parseFloat(first.value);
          if (Object.keys(p).length > 2) ops.push(p);
        }
        if (!ops.length) { close(); return; }
        send(ops).then(close).catch(function () { });
      } })];
    });
  }
  function fonts() {
    if (FONTS) return Promise.resolve(FONTS);
    var own = []; S.blocks.forEach(function (b) { if (b.font && own.indexOf(b.font) < 0) own.push(b.font); });
    return api('GET', '/fonts', null, true).then(function (r) { return r.text(); }).then(function (t) {
      var seen = {}, out = own.slice(); own.forEach(function (x) { seen[x.toLowerCase()] = 1; });
      t.split('\n').forEach(function (l) { l = l.trim(); if (!l || l[0] === '#') return; var i = l.indexOf('  |  '); if (i > 0) l = l.slice(0, i); var tc = l.replace(/(^|[ -])([a-z])/g, function (m, a, c) { return a + c.toUpperCase(); }); if (!seen[tc.toLowerCase()]) { seen[tc.toLowerCase()] = 1; out.push(tc); } });
      return (FONTS = out);
    }).catch(function () { return (FONTS = own); });
  }

  // ---- Excel grid
  function gridView(main) {
    main.appendChild(h('div', { class: 'tabs', id: 'tabs' }));
    main.appendChild(h('div', { class: 'card', style: 'padding:10px' }, h('div', { class: 'row' }, h('b', { id: 'cref', text: '—', style: 'flex:0 0 70px' }),
      h('input', { type: 'text', id: 'fbar', placeholder: 'Select a cell, then type a value or a formula (=SUM(A1:A5))', dir: 'auto', onkeydown: function (e) { if (e.key === 'Enter') setCell(); } }), h('button', { class: 'pri', text: 'Set', onclick: setCell }),
      h('button', { text: 'Clear', onclick: function () { document.getElementById('fbar').value = ''; setCell(); } })),
      h('div', { class: 'row', style: 'margin-top:8px' }, h('button', { text: '⇡ 20', onclick: function () { S.r0 = Math.max(1, S.r0 - 20); loadGrid(); } }), h('button', { text: '⇣ 20', onclick: function () { S.r0 += 20; loadGrid(); } }),
        h('button', { text: '⇠ 6', onclick: function () { S.c0 = Math.max(1, S.c0 - 6); loadGrid(); } }), h('button', { text: '⇢ 6', onclick: function () { S.c0 += 6; loadGrid(); } }),
        h('input', { type: 'text', id: 'goto', placeholder: 'Go to cell (e.g. D40)', style: 'min-width:140px', onkeydown: function (e) { if (e.key === 'Enter') gotoCell(this.value); } }))));
    main.appendChild(h('div', { class: 'gridwrap', id: 'gw' }));
    main.appendChild(h('div', { class: 'hint', id: 'gnote' }));
    loadGrid();
  }
  function colName(n) { var s = ''; while (n) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function gotoCell(v) { var m = /^([A-Za-z]{1,3})(\d{1,7})$/.exec((v || '').trim()); if (!m) { toast('Example: D40'); return; } var c = 0; m[1].toUpperCase().split('').forEach(function (ch) { c = c * 26 + ch.charCodeAt(0) - 64; }); S.c0 = Math.max(1, c - 1); S.r0 = Math.max(1, parseInt(m[2], 10) - 3); S.sel = { ref: m[1].toUpperCase() + m[2], c: c, r: parseInt(m[2], 10) }; loadGrid(); }
  function loadGrid() {
    api('GET', '/edit/grid?' + q({ sid: S.sid, sheet: S.sheet, r0: S.r0, c0: S.c0, rows: 30, cols: 8 })).then(function (g) {
      S.sheets = g.sheets; S.protectedSheet = g.protected;
      var tabs = document.getElementById('tabs'); if (tabs) { tabs.innerHTML = ''; g.sheets.forEach(function (n, i) { tabs.appendChild(h('button', { text: n, class: i === S.sheet ? 'on' : '', onclick: function () { S.sheet = i; S.r0 = 1; S.c0 = 1; S.sel = null; loadGrid(); } })); }); }
      var t = h('table', { class: 'grid' }), tr = h('tr', {}, h('th', { class: 'rh', text: '' }));
      for (var c = 0; c < 8; c++) tr.appendChild(h('th', { text: colName(S.c0 + c) })); t.appendChild(tr);
      g.rows.forEach(function (row, ri) {
        var r = S.r0 + ri, rr = h('tr', {}, h('th', { class: 'rh', text: String(r) }));
        row.forEach(function (cell, ci) {
          var c = S.c0 + ci, ref = colName(c) + r, txt = cell ? cell[0] : '', num = txt !== '' && !isNaN(Number(txt));
          var td = h('td', { class: (S.sel && S.sel.ref === ref ? 'sel ' : '') + (cell && cell[1] ? 'fm ' : '') + (num ? 'num' : ''), dir: 'auto', text: (txt === '' && cell && cell[1]) ? cell[1] : txt, title: cell && cell[1] ? cell[1] : '',
            onclick: function () { S.sel = { ref: ref, c: c, r: r, f: cell && cell[1] ? cell[1] : null, v: txt }; document.getElementById('cref').textContent = ref; document.getElementById('fbar').value = S.sel.f || txt; var o = document.querySelector('td.sel'); if (o) o.classList.remove('sel'); td.classList.add('sel'); document.getElementById('fbar').focus(); } });
          rr.appendChild(td);
        }); t.appendChild(rr);
      });
      var gw = document.getElementById('gw'); gw.innerHTML = ''; gw.appendChild(t);
      document.getElementById('gnote').textContent = 'Used area: ' + (g.max_row ? 'A1:' + colName(g.max_col) + g.max_row : 'empty') + (g.protected ? '  ·  this sheet is protected: editing is refused' : '') + '  ·  to format cells use Settings → Cell format (the selected cell is pre-filled).';
      if (S.sel) { document.getElementById('cref').textContent = S.sel.ref; }
    }).catch(fail);
  }
  function setCell() {
    if (!S.sel) { toast('Select a cell first'); return; }
    send([{ op: 'set_cell', sheet: S.sheet, ref: S.sel.ref, value: document.getElementById('fbar').value }], S.sel.ref + ' updated').then(function (j) { if (j.formula_results_cleared) toast('Formula results will be recalculated when the file is opened'); }).catch(function () { });
  }

  // ------------------------------------------------------------------ undo / report / preview
  function undo() { var d = busy('Undoing…'); api('POST', '/edit/undo?sid=' + S.sid, {}).then(function (j) { d(); S.blocks = j.blocks || S.blocks; S.ops = j.ops; S.tables = j.tables || S.tables; S.page = j.page || S.page; refreshOps(); draw2(); toast('Undone'); }).catch(function (e) { d(); toast(e.message); }); }
  function reportBody(rep) {
    var wrap = h('div', { class: 'rep' }); var st = rep.stats || {};
    if (Object.keys(st).length) wrap.appendChild(h('div', { class: 'mut', text: Object.keys(st).map(function (k) { return k + ': ' + st[k]; }).join('   ') }));
    (rep.items || []).forEach(function (it) { wrap.appendChild(h('div', { class: 'it' }, h('span', { class: 'lv ' + it.level, text: it.level }), it.text)); });
    wrap.appendChild(h('div', { class: 'mut', style: 'margin-top:8px', text: 'This report is also written to the converter log (Termux).' })); return wrap;
  }
  function showReport(start) { var blocked = S.report && S.report.blocked; modal(blocked ? 'Editing is not possible' : 'Preflight report', reportBody(S.report || { items: [] }), function (close) { return [h('button', { class: 'pri', text: start && !blocked ? 'Start editing a copy' : 'OK', onclick: close })]; }); }
  function preview() {
    var d = busy('Rendering with the real layout engine…');
    api('GET', '/edit/render?sid=' + S.sid).then(function (r) {
      d(); var dpi = 90, box = h('div', { class: 'pages' });
      function fill() { box.innerHTML = ''; for (var n = 1; n <= r.pages; n++) box.appendChild(h('img', { loading: 'lazy', alt: 'page ' + n, src: '/edit/page?' + q({ sid: S.sid, n: n, dpi: dpi, k: r.key }) })); }
      fill(); var bar = h('div', { class: 'row' }, h('span', { class: 'mut', text: r.pages + ' page(s)' }), h('button', { text: 'Sharper', onclick: function () { dpi = Math.min(180, dpi + 40); fill(); } }), h('button', { text: 'Lighter', onclick: function () { dpi = Math.max(50, dpi - 30); fill(); } }));
      modal('Preview (real layout)', h('div', {}, bar, box));
    }).catch(function (e) { d(); fail(e); });
  }

  // ------------------------------------------------------------------ find & replace
  function findPanel() {
    var qI = h('input', { type: 'text', placeholder: 'Find', dir: 'auto' }), rI = h('input', { type: 'text', placeholder: 'Replace with (leave empty to delete)', dir: 'auto' });
    function chk(l) { var c = h('input', { type: 'checkbox' }); return { c: c, el: h('label', { class: 'chk' }, c, l) }; }
    var cs = chk('Match case'), ww = chk('Whole word'), rx = chk('Regular expression'); var out = h('div', { class: 'list', style: 'margin-top:10px' }), resBox = h('div');
    function opts() { return { q: qI.value, case: cs.c.checked ? 1 : 0, word: ww.c.checked ? 1 : 0, regex: rx.c.checked ? 1 : 0 }; }
    function doFind() {
      if (!qI.value) { toast('Type something to find'); return; } var d = busy('Searching…');
      api('GET', '/edit/find?' + q(Object.assign({ sid: S.sid }, opts()))).then(function (j) {
        d(); out.innerHTML = ''; out.appendChild(h('div', { class: 'mut', style: 'padding:6px 2px', text: j.matches + ' match(es) in ' + j.paragraphs + ' ' + (S.kind === 'xlsx' ? 'cell(s)' : 'paragraph(s)') + (j.truncated ? ' (first ' + j.paragraphs + ' shown)' : '') }));
        j.results.forEach(function (r) {
          out.appendChild(h('div', { class: 'item' + (r.locked ? ' locked' : ''), onclick: function () {
            if (S.kind === 'xlsx') { S.sheet = r.sheet; gotoCell(r.ref); document.querySelectorAll('.sheet').forEach(function (s) { s.remove(); }); }
            else { var b = S.blocks[r.i]; document.querySelectorAll('.sheet').forEach(function (s) { s.remove(); }); if (b && !b.locked) paraEditor(b); else if (b) lockedInfo(b); } } },
            h('div', { class: 'hd', text: (S.kind === 'xlsx' ? r.sheet_name + '!' + r.ref : (r.where || '¶ ' + (r.i + 1))) + (r.count > 1 ? '  ×' + r.count : '') + (r.locked ? '  🔒' : '') }), h('div', { class: 'tx', dir: 'auto', text: r.snippet })));
        });
      }).catch(function (e) { d(); fail(e); });
    }
    function replace(dry) {
      if (!qI.value) { toast('Type something to find'); return; }
      var rule = [{ find: qI.value, replace: rI.value, match_case: cs.c.checked, whole_word: ww.c.checked, regex: rx.c.checked }];
      runSettings('text', { replacements: JSON.stringify(rule) }, null, dry, resBox, function () { if (!dry) { resBox.insertBefore(h('div', { class: 'mut', text: 'Done. Use ↶ Undo to take it back.' }), resBox.firstChild); } });
    }
    modal('Find & Replace', h('div', {}, qI, rI, h('div', { class: 'row', style: 'margin-top:4px' }, cs.el, ww.el, rx.el), h('div', { class: 'row', style: 'margin-top:8px' }, h('button', { class: 'pri', text: 'Find all', onclick: doFind }), h('button', { text: 'Preview replace', onclick: function () { replace(true); } }), h('button', { class: 'ok', text: 'Replace all', onclick: function () { replace(false); } })), resBox, out));
  }

  // ------------------------------------------------------------------ settings (forms generated from the schema)
  function settingsPanel(pre) {
    var d = busy('Loading settings…');
    api('GET', '/edit/schema?sid=' + S.sid).then(function (j) {
      d(); var body = h('div'); var m = modal('Settings', body);
      function groups() {
        body.innerHTML = ''; body.appendChild(h('div', { class: 'mut', text: 'Choose what to change. Every setting is previewed first; only the fields you fill in are applied; the original file is never touched.' }));
        j.groups.forEach(function (g) { body.appendChild(h('div', { class: 'gcard', onclick: function () { form(g); } }, h('b', { text: g.label }), h('span', { text: g.description }))); });
      }
      function form(g) {
        body.innerHTML = ''; var inputs = {};
        body.appendChild(h('button', { text: '← All settings', onclick: groups })); body.appendChild(h('h3', { style: 'margin-top:10px' }, h('span', { text: g.label })));
        var scope = null;
        if (g.uses_scope) {
          scope = {}; var parts = ['body', 'headers', 'footers', 'footnotes', 'endnotes']; var pw = h('div', { class: 'chips' });
          parts.forEach(function (p) { var c = h('input', { type: 'checkbox' }); c.checked = p === 'body'; scope[p] = c; pw.appendChild(h('label', { class: 'chk' }, c, p)); });
          scope.styles = h('input', { type: 'text', placeholder: 'only paragraphs with these styles (comma separated, empty = all)' }); scope.cover = h('input', { type: 'checkbox' });
          body.appendChild(h('div', { class: 'sec' }, h('h4', { text: 'Apply to' }), pw, scope.styles, h('label', { class: 'chk' }, scope.cover, 'include the cover page')));
        }
        g.sections.forEach(function (sec) {
          var s = h('div', { class: 'sec' }, h('h4', { text: sec.title })); body.appendChild(s);
          sec.fields.forEach(function (f) { s.appendChild(h('label', { class: 'lbl', text: f.label + (f.unit ? ' (' + f.unit + ')' : '') })); var inp = field(f); inputs[f.name] = { f: f, el: inp }; s.appendChild(inp); if (f.hint) s.appendChild(h('div', { class: 'hint', text: f.hint })); });
        });
        var res = h('div');
        function values() {
          var v = {}; Object.keys(inputs).forEach(function (k) {
            var f = inputs[k].f, el = inputs[k].el, x;
            if (f.kind === 'bool') { x = el.value === '' ? undefined : el.value === '1'; } else x = el.value.trim() === '' ? undefined : el.value.trim();
            if (x !== undefined) v[k] = x;
          }); return v;
        }
        function sc() { if (!scope) return null; var parts = ['body', 'headers', 'footers', 'footnotes', 'endnotes'].filter(function (p) { return scope[p].checked; }); return { parts: parts, styles: scope.styles.value.trim() || null, include_cover: scope.cover.checked }; }
        if (pre && pre.group === g.key) Object.keys(pre.values).forEach(function (k) { if (inputs[k]) inputs[k].el.value = pre.values[k]; });
        body.appendChild(res);
        body.appendChild(h('div', { class: 'row', style: 'margin-top:14px' }, h('button', { text: 'Preview impact', onclick: function () { runSettings(g.key, values(), sc(), true, res); } }), h('button', { class: 'pri', text: 'Apply', onclick: function () { runSettings(g.key, values(), sc(), false, res); } })));
      }
      var preGroup = pre && pre.group ? j.groups.filter(function (x) { return x.key === pre.group; })[0] : null;
      if (preGroup) form(preGroup); else groups();
    }).catch(function (e) { d(); fail(e); });
  }
  function field(f) {
    var e;
    if (f.kind === 'bool') e = h('select', {}, h('option', { value: '', text: '(unchanged)' }), h('option', { value: '1', text: 'Yes' }), h('option', { value: '0', text: 'No' }));
    else if (f.kind === 'choice') { e = h('select', {}, h('option', { value: '', text: '(unchanged)' })); f.choices.forEach(function (c) { e.appendChild(h('option', { value: c, text: c })); }); }
    else if (f.kind === 'json' || f.kind === 'list') e = h('textarea', { rows: 3, dir: 'auto', placeholder: f.kind === 'json' ? '[...]' : 'one per line or comma separated' });
    else if (f.kind === 'color') { e = h('input', { type: 'text', placeholder: 'RRGGBB, e.g. 1F3864 (or auto)' }); }
    else if (f.kind === 'int' || f.kind === 'float') e = h('input', { type: 'text', inputmode: 'decimal', placeholder: f.kind === 'int' ? 'whole number' : 'number' });
    else if (f.kind === 'length') e = h('input', { type: 'text', placeholder: f.unit ? 'e.g. 2 ' + f.unit + ' (or 12pt, 0.5in)' : 'e.g. 12pt, 2cm, 0.5in' });
    else e = h('input', { type: 'text', dir: 'auto' });
    return e;
  }
  function runSettings(group, values, scope, dry, resBox, after) {
    if (!Object.keys(values).length) { toast('Fill in at least one setting'); return; }
    var d = busy(dry ? 'Checking what would change…' : 'Applying…');
    api('POST', '/edit/settings?sid=' + S.sid, { group: group, values: values, scope: scope, dry: dry }).then(function (j) {
      d(); var r = j.result; resBox.innerHTML = '';
      var box = h('div', { class: 'res' }); resBox.appendChild(box);
      box.appendChild(h('div', { html: '<b>' + (dry ? 'Preview — nothing was changed yet' : (r.nothing_changed ? 'Nothing needed changing' : 'Applied')) + '</b>' }));
      var ch = Object.keys(r.changes || {}); if (ch.length) box.appendChild(h('div', { text: ch.map(function (k) { return k + ': ' + r.changes[k]; }).join('   ') })); else if (r.nothing_changed) box.appendChild(h('div', { class: 'mut', text: 'The file already has these settings, or no matching text was found.' }));
      if (r.paragraphs_text_changed) { box.appendChild(h('div', { text: r.paragraphs_text_changed + ' item(s) change their text' })); (r.samples || []).forEach(function (s) { box.appendChild(h('div', { class: 's', dir: 'auto', text: '“' + s.old + '” → “' + s.new + '”' })); }); }
      (r.warnings || []).forEach(function (w) { box.appendChild(h('div', { class: 'w', text: '⚠ ' + w })); });
      if (r.parts_changed && r.parts_changed.length) box.appendChild(h('div', { class: 's', text: 'Parts changed: ' + r.parts_changed.join(', ') + (r.parts_added && r.parts_added.length ? ' · added: ' + r.parts_added.join(', ') : '') + (r.parts_removed && r.parts_removed.length ? ' · removed: ' + r.parts_removed.join(', ') : '') }));
      if (!dry) { mergeBlocks(Object.assign({ dry: false }, j)); refreshOps(); draw2(); toast(r.nothing_changed ? 'Nothing to change' : 'Applied'); }
      if (after) after(j);
    }).catch(function (e) { d(); fail(e); });
  }

  // ------------------------------------------------------------------ save (verify first, then download)
  function savePanel() {
    var d = busy('Verifying the copy before it can be saved…');
    api('GET', '/edit/verify?' + q({ sid: S.sid, orig_pages: S.origPages })).then(function (j) {
      d(); var v = j.verify, body = h('div');
      body.appendChild(h('div', { class: 'res', style: 'margin:0 0 10px' }, h('b', { text: v.ok ? '✔ Verification passed' : '✖ VERIFICATION FAILED — saving is not recommended' })));
      v.checks.forEach(function (c) { body.appendChild(h('div', { class: 'it', style: 'padding:4px 0' }, h('span', { class: 'lv ' + (c.ok ? 'OK' : 'BLOCK'), text: c.ok ? 'OK' : 'FAIL' }), h('b', { text: c.name }), c.detail ? h('div', { class: 'mut', style: 'margin-inline-start:60px', text: c.detail }) : null)); });
      if (v.changes && v.changes.length) { body.appendChild(h('h4', { text: 'Changed items', style: 'margin:12px 0 4px' })); v.changes.slice(0, 10).forEach(function (c) { body.appendChild(h('div', { class: 's mut', dir: 'auto', text: '“' + c.old + '” → “' + c.new + '”' })); }); if (v.changes.length > 10) body.appendChild(h('div', { class: 'mut', text: '… and ' + (v.changes.length - 10) + ' more' })); }
      var native = S.meta.ext, wanted = S.meta.save_ext, opts = { docx: ['docx', 'pdf', 'odt', 'rtf', 'doc'], xlsx: ['xlsx', 'pdf', 'ods', 'xls'], pptx: ['pptx', 'pdf', 'odp', 'ppt'] }[S.kind].slice();
      [wanted, native].forEach(function (x) { if (opts.indexOf(x) < 0) opts.unshift(x); });
      var sel = h('select', {}); opts.forEach(function (o) { sel.appendChild(h('option', { value: o, text: '.' + o + (o === native ? '  (same format as the copy)' : o === 'pdf' ? '  (PDF)' : '  (converted by LibreOffice)') })); }); sel.value = wanted;
      body.appendChild(h('label', { class: 'lbl', text: 'Save as format' })); body.appendChild(sel);
      body.appendChild(h('div', { class: 'hint', text: 'A new file is downloaded by the browser (look in your Downloads). Your original file is not changed. Other formats than the copy\'s own are converted by LibreOffice and may differ in details.' }));
      modal('Save a copy', body, function (close) { return [h('button', { text: 'Cancel', onclick: close }), h('button', { class: v.ok ? 'ok' : 'bad', text: v.ok ? 'Download' : 'Download anyway', onclick: function () {
        var a = h('a', { href: '/edit/export?' + q({ sid: S.sid, as: sel.value }) }); document.body.appendChild(a); a.click(); a.remove(); close(); toast('Downloading…'); } })]; });
    }).catch(function (e) { d(); fail(e); });
  }

  route();
})();
