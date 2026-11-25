/* Pocket Office viewer: renders DOCX / XLSX / PPTX inside WebView and exports pages. */
(function () {
  'use strict';
  var B = window.Android || { onReady: function () {}, onPage: function () {}, onError: function () {}, progress: function () {},
    savePage: function () {}, exportDone: function () {}, exportFailed: function () {}, saveText: function () {}, copyText: function () {} };
  var root = document.getElementById('root');
  var msg = document.getElementById('msg');
  var type = new URLSearchParams(location.search).get('type');
  var pages = [];          // {el, w, h, y (slice offset, optional)}
  var cancelFlag = false;

  function fail(e) {
    var m = (e && e.message) ? e.message : String(e);
    msg.style.display = 'block';
    msg.textContent = 'Cannot display this file: ' + m;
    try { B.onError(m); } catch (x) {}
  }
  function need(name, ok) { if (!ok) throw new Error('library "' + name + '" is missing from the app build'); }
  function setViewport(w) {
    var vp = document.getElementById('vp');
    vp.setAttribute('content', 'width=' + Math.max(200, Math.min(10000, Math.ceil(w))) + ', user-scalable=yes, minimum-scale=0.05, maximum-scale=8');
  }
  function absTop(el) { return el.getBoundingClientRect().top + (window.pageYOffset || 0); }
  function el(tag, cls, parent) { var e = document.createElement(tag); if (cls) e.className = cls; if (parent) parent.appendChild(e); return e; }
  function nextFrame() { return new Promise(function (r) { setTimeout(r, 0); }); }

  /* ------------------------------------------------------------------ DOCX */
  async function renderDocx(buf) {
    need('docx-preview', window.docx && window.docx.renderAsync); need('jszip', window.JSZip);
    var host = el('div', null, root);
    await window.docx.renderAsync(buf, host, null, {
      className: 'docx', inWrapper: true, ignoreWidth: false, ignoreHeight: false, ignoreFonts: false,
      breakPages: true, ignoreLastRenderedPageBreak: false, experimental: true, trimXmlDeclaration: true,
      useBase64URL: true, renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true,
      renderAltChunks: true
    });
    var list = host.querySelectorAll('section.docx');
    var maxW = 0;
    if (!list.length) list = host.querySelectorAll('.docx-wrapper');
    for (var i = 0; i < list.length; i++) {
      var s = list[i];
      pages.push({ el: s, w: s.offsetWidth, h: s.offsetHeight });
      maxW = Math.max(maxW, s.offsetWidth);
    }
    setViewport(maxW + 16);
  }

  /* ------------------------------------------------------------------ colour helpers */
  function hexToRgb(h) { h = h.replace('#', ''); return [parseInt(h.substr(0, 2), 16), parseInt(h.substr(2, 2), 16), parseInt(h.substr(4, 2), 16)]; }
  function rgbToHex(r) { function c(v) { v = Math.max(0, Math.min(255, Math.round(v))); return (v < 16 ? '0' : '') + v.toString(16); } return '#' + c(r[0]) + c(r[1]) + c(r[2]); }
  function rgbToHsl(r) {
    var R = r[0] / 255, G = r[1] / 255, Bb = r[2] / 255, mx = Math.max(R, G, Bb), mn = Math.min(R, G, Bb), h = 0, s = 0, l = (mx + mn) / 2;
    if (mx !== mn) {
      var d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
      if (mx === R) h = (G - Bb) / d + (G < Bb ? 6 : 0); else if (mx === G) h = (Bb - R) / d + 2; else h = (R - G) / d + 4;
      h /= 6;
    }
    return [h, s, l];
  }
  function hslToRgb(a) {
    var h = a[0], s = a[1], l = a[2];
    function f(p, q, t) { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; }
    if (s === 0) return [l * 255, l * 255, l * 255];
    var q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    return [f(p, q, h + 1 / 3) * 255, f(p, q, h) * 255, f(p, q, h - 1 / 3) * 255];
  }
  function lumMod(hex, mod, off) { var hsl = rgbToHsl(hexToRgb(hex)); hsl[2] = Math.max(0, Math.min(1, hsl[2] * mod + off)); return rgbToHex(hslToRgb(hsl)); }
  function excelTint(hex, t) {
    if (!t) return hex; var hsl = rgbToHsl(hexToRgb(hex));
    hsl[2] = t < 0 ? hsl[2] * (1 + t) : hsl[2] * (1 - t) + t;
    return rgbToHex(hslToRgb(hsl));
  }

  /* ------------------------------------------------------------------ XLSX */
  var XL_INDEXED = ['#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF'];
  var XL_THEME = null;

  async function loadXlTheme(buf) {
    try {
      var zip = await window.JSZip.loadAsync(buf);
      var f = zip.file('xl/theme/theme1.xml'); if (!f) return;
      var doc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
      var names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      XL_THEME = names.map(function (n) {
        var node = doc.getElementsByTagName('a:' + n)[0]; if (!node) return '#000000';
        var c = node.firstElementChild; if (!c) return '#000000';
        var v = c.getAttribute('lastClr') || c.getAttribute('val') || '000000';
        return '#' + v;
      });
    } catch (e) { XL_THEME = null; }
  }
  function xlColor(c) {
    if (!c) return null;
    if (c.argb) { var s = String(c.argb); return '#' + s.substr(s.length - 6); }
    if (c.theme != null && XL_THEME) { var b = XL_THEME[c.theme]; return b ? excelTint(b, c.tint || 0) : null; }
    if (c.indexed != null) return XL_INDEXED[c.indexed] || null;
    return null;
  }
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function trimNum(v) { if (Number.isInteger(v)) return String(v); var s = String(parseFloat(v.toPrecision(10))); return s; }
  function fmtNum(v, f) {
    if (!f || f === 'General') return trimNum(v);
    var fmt = f.split(';')[0];
    var pct = /%/.test(fmt);
    var dm = fmt.match(/\.([0#]+)/); var dec = dm ? dm[1].length : 0;
    var comma = /#,##0|0,000|#,#/.test(fmt);
    var x = pct ? v * 100 : v;
    var s = Math.abs(x).toFixed(dec);
    if (dm && dm[1].indexOf('#') >= 0) s = String(parseFloat(s));
    if (comma) { var p = s.split('.'); p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ','); s = p.join('.'); }
    var pre = '', suf = '', m;
    if ((m = fmt.match(/^"([^"]*)"/))) pre = m[1]; else if ((m = fmt.match(/^\[\$([^\]\-]*)/))) pre = m[1]; else if (/^\$/.test(fmt)) pre = '$';
    if ((m = fmt.match(/"([^"]*)"\s*$/)) && !/^"/.test(fmt.replace(m[0], ''))) suf = m[1];
    if (pct) suf = '%' + suf;
    return (x < 0 ? '-' : '') + pre + s + suf;
  }
  function fmtDate(d, f) {
    f = f && f !== 'General' ? f.split(';')[0] : 'yyyy-mm-dd';
    f = f.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '');
    var hasTime = /[hs]/i.test(f), hasDate = /[yd]/i.test(f) || /m/i.test(f) && !hasTime;
    var Y = d.getUTCFullYear(), Mo = d.getUTCMonth(), D = d.getUTCDate(), H = d.getUTCHours(), Mi = d.getUTCMinutes(), S = d.getUTCSeconds();
    var ampm = /AM\/PM|A\/P/i.test(f), H12 = H % 12 || 12;
    var out = f.replace(/yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|AM\/PM|A\/P/gi, function (tok, off, str) {
      var t = tok.toLowerCase();
      if (t === 'yyyy') return String(Y); if (t === 'yy') return String(Y).slice(-2);
      if (t === 'mmmm') return MONTHS[Mo]; if (t === 'mmm') return MONTHS[Mo].substr(0, 3);
      if (t === 'mm' || t === 'm') {
        var before = str.substr(0, off).toLowerCase(), after = str.substr(off + tok.length).toLowerCase();
        var isMin = /h[^a-z]*$/.test(before) || /^[^a-z]*s/.test(after);
        var val = isMin ? Mi : Mo + 1; return t === 'mm' ? pad2(val) : String(val);
      }
      if (t === 'dd') return pad2(D); if (t === 'd') return String(D);
      if (t === 'hh') return pad2(ampm ? H12 : H); if (t === 'h') return String(ampm ? H12 : H);
      if (t === 'ss') return pad2(S); if (t === 's') return String(S);
      if (t === 'am/pm' || t === 'a/p') return H < 12 ? 'AM' : 'PM';
      return tok;
    });
    return out;
  }
  function xlText(cell) {
    var v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && !(v instanceof Date)) {
      if (v.richText) return v.richText.map(function (r) { return r.text; }).join('');
      if (v.formula !== undefined || v.sharedFormula !== undefined) v = v.result;
      else if (v.hyperlink) return v.text && v.text.richText ? v.text.richText.map(function (r) { return r.text; }).join('') : (v.text || v.hyperlink);
      else if (v.error) return v.error;
      else if (v.text) return String(v.text);
      else return '';
    }
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmtDate(v, cell.numFmt);
    if (typeof v === 'number') return fmtNum(v, cell.numFmt);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'object') { if (v.richText) return v.richText.map(function (r) { return r.text; }).join(''); if (v.error) return v.error; return ''; }
    return String(v);
  }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function addr(s) { var m = /^([A-Z]+)(\d+)$/.exec(s); return m ? { c: colNum(m[1]), r: parseInt(m[2], 10) } : { c: 1, r: 1 }; }
  var BORDER = { thin: '1px solid', hair: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double', mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', slantDashDot: '1px dashed' };

  function renderSheet(ws) {
    var wrap = el('div', 'sheet', root);
    el('div', 'sheetTitle', wrap).textContent = ws.name;
    var maxR = Math.min(ws.rowCount || 0, 3000), maxC = Math.min(ws.columnCount || 0, 150);
    var truncated = (ws.rowCount > maxR) || (ws.columnCount > maxC);
    var rtl = false; try { rtl = !!(ws.views && ws.views[0] && ws.views[0].rightToLeft); } catch (e) {}
    var grid = true; try { grid = !(ws.views && ws.views[0] && ws.views[0].showGridLines === false); } catch (e) {}
    var table = el('table', null, wrap); if (rtl) table.setAttribute('dir', 'rtl');
    var merges = {}, covered = {};
    var ml = (ws.model && ws.model.merges) || [];
    ml.forEach(function (m) {
      var p = m.split(':'), a = addr(p[0]), b = addr(p[1] || p[0]);
      merges[a.r + ',' + a.c] = { rs: b.r - a.r + 1, cs: b.c - a.c + 1 };
      for (var r = a.r; r <= b.r; r++) for (var c = a.c; c <= b.c; c++) if (r !== a.r || c !== a.c) covered[r + ',' + c] = true;
    });
    var cg = el('colgroup', null, table), totalW = 0, hiddenCols = {};
    for (var c = 1; c <= maxC; c++) {
      var col = ws.getColumn(c), cw = col.width ? Math.round(col.width * 7 + 5) : 64;
      if (col.hidden) { hiddenCols[c] = true; continue; }
      var ce = el('col', null, cg); ce.style.width = cw + 'px'; totalW += cw;
    }
    table.style.width = totalW + 'px';
    var defH = (ws.properties && ws.properties.defaultRowHeight) || 15;
    for (var r = 1; r <= maxR; r++) {
      var row = ws.getRow(r);
      if (row.hidden) continue;
      var tr = el('tr', null, table);
      tr.style.height = Math.round((row.height || defH) * 96 / 72) + 'px';
      for (var cc = 1; cc <= maxC; cc++) {
        if (hiddenCols[cc] || covered[r + ',' + cc]) continue;
        var cell = row.getCell(cc), td = el('td', null, tr), st = td.style, mg = merges[r + ',' + cc];
        if (mg) { if (mg.cs > 1) td.colSpan = mg.cs; if (mg.rs > 1) td.rowSpan = mg.rs; }
        var font = cell.font || {}, fill = cell.fill, al = cell.alignment || {}, bd = cell.border || {};
        if (font.name) st.fontFamily = '"' + font.name + '", Calibri, Arial, sans-serif';
        st.fontSize = (font.size || 11) + 'pt';
        if (font.bold) st.fontWeight = 'bold'; if (font.italic) st.fontStyle = 'italic';
        var deco = []; if (font.underline) deco.push('underline'); if (font.strike) deco.push('line-through'); if (deco.length) st.textDecoration = deco.join(' ');
        var fc = xlColor(font.color); st.color = fc || '#000';
        if (fill && fill.type === 'pattern' && fill.pattern === 'solid') { var bg = xlColor(fill.fgColor); if (bg) st.backgroundColor = bg; }
        var val = cell.value, isNum = typeof val === 'number' || val instanceof Date || (val && typeof val === 'object' && typeof val.result === 'number');
        var h = al.horizontal; if (!h || h === 'general') h = isNum ? 'right' : (typeof val === 'boolean' ? 'center' : 'left');
        if (rtl && !al.horizontal) h = isNum ? 'right' : 'right';
        st.textAlign = h === 'centerContinuous' ? 'center' : (h === 'fill' ? 'left' : h);
        st.verticalAlign = al.vertical === 'top' ? 'top' : (al.vertical === 'middle' || al.vertical === 'center' ? 'middle' : 'bottom');
        if (al.wrapText) { st.whiteSpace = 'pre-wrap'; st.wordBreak = 'break-word'; } else { st.whiteSpace = 'nowrap'; }
        if (al.indent) st.paddingLeft = (3 + al.indent * 9) + 'px';
        var sides = ['top', 'right', 'bottom', 'left'];
        sides.forEach(function (sd) {
          var b = bd[sd], css = 'border' + sd.charAt(0).toUpperCase() + sd.slice(1);
          if (b && b.style) st[css] = (BORDER[b.style] || '1px solid') + ' ' + (xlColor(b.color) || '#000');
          else if (grid) st[css] = '1px solid #e1e4ea';
        });
        td.textContent = xlText(cell);
      }
    }
    if (truncated) { var n = el('div', null, wrap); n.style.cssText = 'font:12px sans-serif;color:#a33;padding:6px 10px'; n.textContent = 'Large sheet: only the first ' + maxR + ' rows and ' + maxC + ' columns are shown.'; }
    return wrap;
  }

  async function renderXlsx(buf) {
    need('exceljs', window.ExcelJS); need('jszip', window.JSZip);
    var wb = new window.ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    await loadXlTheme(buf);
    var wraps = [];
    wb.eachSheet(function (ws) { if (ws.state === 'hidden' || ws.state === 'veryHidden') return; wraps.push(renderSheet(ws)); });
    if (!wraps.length) throw new Error('workbook has no visible sheets');
    var maxW = 0;
    wraps.forEach(function (w) { maxW = Math.max(maxW, w.offsetWidth); });
    // split every sheet into row-aligned pages
    wraps.forEach(function (w) {
      var wr = w.getBoundingClientRect(), trs = w.querySelectorAll('tr'), start = 0, total = w.offsetHeight, MAXH = 1400;
      var wpx = Math.max(w.offsetWidth, w.scrollWidth);
      for (var i = 0; i < trs.length; i++) {
        var tr = trs[i].getBoundingClientRect(), top = tr.top - wr.top, bottom = tr.bottom - wr.top;
        if (bottom - start > MAXH && top > start + 20) { pages.push({ el: w, y: start, w: wpx, h: Math.round(top - start), slice: true }); start = Math.round(top); }
      }
      pages.push({ el: w, y: start, w: wpx, h: Math.max(20, total - start), slice: true });
    });
    setViewport(Math.min(maxW + 32, 1100));
  }

  /* ------------------------------------------------------------------ PPTX */
  var EMU = 9525;
  function pxE(v) { return (parseInt(v, 10) || 0) / EMU; }
  function kids(node, name) { var out = []; if (!node) return out; for (var c = node.firstElementChild; c; c = c.nextElementSibling) if (c.nodeName === name) out.push(c); return out; }
  function kid(node, name) { return kids(node, name)[0] || null; }
  function path(node, names) { var n = node; for (var i = 0; i < names.length && n; i++) n = kid(n, names[i]); return n; }
  function attr(n, a, d) { if (!n) return d; var v = n.getAttribute(a); return v === null ? d : v; }
  function dirName(p) { var i = p.lastIndexOf('/'); return i < 0 ? '' : p.substr(0, i + 1); }
  function resolvePath(base, target) {
    if (target.charAt(0) === '/') return target.substr(1);
    var parts = (dirName(base) + target).split('/'), out = [];
    parts.forEach(function (p) { if (p === '..') out.pop(); else if (p !== '.' && p !== '') out.push(p); });
    return out.join('/');
  }
  var P = { zip: null, theme: {}, clrMap: { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }, media: {}, defText: null, W: 960, H: 540 };

  async function ptxt(p) { var f = P.zip.file(p); return f ? await f.async('string') : null; }
  async function pxml(p) { var t = await ptxt(p); return t ? new DOMParser().parseFromString(t, 'application/xml') : null; }
  async function prels(partPath) {
    var rp = dirName(partPath) + '_rels/' + partPath.substr(partPath.lastIndexOf('/') + 1) + '.rels', d = await pxml(rp), map = {};
    if (!d) return map;
    var rs = d.getElementsByTagName('Relationship');
    for (var i = 0; i < rs.length; i++) map[rs[i].getAttribute('Id')] = { type: (rs[i].getAttribute('Type') || '').split('/').pop(), target: resolvePath(partPath, rs[i].getAttribute('Target') || '') };
    return map;
  }
  async function pmedia(p) {
    if (P.media[p] !== undefined) return P.media[p];
    var f = P.zip.file(p), url = null;
    if (f) {
      var ext = p.split('.').pop().toLowerCase(), mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp' }[ext];
      if (mime) url = 'data:' + mime + ';base64,' + await f.async('base64');
    }
    P.media[p] = url; return url;
  }
  function schemeColor(val) {
    var n = P.clrMap[val] || val;
    return P.theme[n] || null;
  }
  // colour node = element that contains srgbClr/schemeClr/... (e.g. a:solidFill)
  function colorOf(node) {
    if (!node) return null;
    var c = kid(node, 'a:srgbClr') || kid(node, 'a:schemeClr') || kid(node, 'a:sysClr') || kid(node, 'a:prstClr');
    if (!c) return null;
    var hex = null;
    if (c.nodeName === 'a:srgbClr') hex = '#' + c.getAttribute('val');
    else if (c.nodeName === 'a:schemeClr') hex = schemeColor(c.getAttribute('val'));
    else if (c.nodeName === 'a:sysClr') hex = '#' + (c.getAttribute('lastClr') || '000000');
    else if (c.nodeName === 'a:prstClr') { var nm = c.getAttribute('val'); hex = nm === 'white' ? '#ffffff' : (nm === 'black' ? '#000000' : null); }
    if (!hex) return null;
    var lm = kid(c, 'a:lumMod'), lo = kid(c, 'a:lumOff'), tint = kid(c, 'a:tint'), shade = kid(c, 'a:shade'), al = kid(c, 'a:alpha');
    if (lm || lo) hex = lumMod(hex, lm ? parseInt(lm.getAttribute('val'), 10) / 100000 : 1, lo ? parseInt(lo.getAttribute('val'), 10) / 100000 : 0);
    if (tint) { var t = parseInt(tint.getAttribute('val'), 10) / 100000, rgb = hexToRgb(hex); hex = rgbToHex(rgb.map(function (v) { return v * t + 255 * (1 - t); })); }
    if (shade) { var s = parseInt(shade.getAttribute('val'), 10) / 100000; hex = rgbToHex(hexToRgb(hex).map(function (v) { return v * s; })); }
    if (al) { var a = parseInt(al.getAttribute('val'), 10) / 100000, r = hexToRgb(hex); return 'rgba(' + r[0] + ',' + r[1] + ',' + r[2] + ',' + a + ')'; }
    return hex;
  }
  async function fillCss(spPr, style) {
    // returns {css:'...'} style properties for background
    if (spPr) {
      if (kid(spPr, 'a:noFill')) return null;
      var sf = kid(spPr, 'a:solidFill'); if (sf) { var c = colorOf(sf); return c ? { backgroundColor: c } : null; }
      var gf = kid(spPr, 'a:gradFill');
      if (gf) {
        var stops = kids(path(gf, ['a:gsLst']), 'a:gs').map(function (g) { return { pos: parseInt(g.getAttribute('pos'), 10) / 1000, c: colorOf(g) || '#888' }; });
        var lin = kid(gf, 'a:lin'), ang = lin ? parseInt(lin.getAttribute('ang'), 10) / 60000 + 90 : 180;
        if (stops.length) return { background: 'linear-gradient(' + ang + 'deg,' + stops.map(function (s) { return s.c + ' ' + s.pos + '%'; }).join(',') + ')' };
      }
      var bf = kid(spPr, 'a:blipFill');
      if (bf) { var u = await blipUrl(bf, spPr); if (u) return { backgroundImage: 'url(' + u + ')', backgroundSize: '100% 100%' }; }
    }
    if (style) { var fr = kid(style, 'a:fillRef'); if (fr && fr.getAttribute('idx') !== '0') { var cc = colorOf(fr); if (cc) return { backgroundColor: cc }; } }
    return null;
  }
  var CUR_RELS = null; // relationships of the part currently being rendered
  async function blipUrl(bf) {
    var blip = kid(bf, 'a:blip'); if (!blip || !CUR_RELS) return null;
    var rid = blip.getAttribute('r:embed'), rel = CUR_RELS[rid]; if (!rel) return null;
    return await pmedia(rel.target);
  }
  function lineCss(spPr, style) {
    var ln = spPr ? kid(spPr, 'a:ln') : null, w = null, c = null, dash = null, none = false;
    if (ln) {
      if (kid(ln, 'a:noFill')) none = true;
      if (ln.getAttribute('w')) w = pxE(ln.getAttribute('w'));
      var sf = kid(ln, 'a:solidFill'); if (sf) c = colorOf(sf);
      var pd = kid(ln, 'a:prstDash'); if (pd) dash = pd.getAttribute('val');
    }
    if (none) return null;
    if (!c && style) { var lr = kid(style, 'a:lnRef'); if (lr && lr.getAttribute('idx') !== '0') { c = colorOf(lr); if (w === null) w = 1; } }
    if (!c) return null;
    return { w: Math.max(0.75, w === null ? 1 : w), c: c, dash: dash };
  }
  function xfrmOf(spPr) {
    var x = spPr ? kid(spPr, 'a:xfrm') : null; if (!x) return null;
    var off = kid(x, 'a:off'), ext = kid(x, 'a:ext'); if (!off || !ext) return null;
    var co = kid(x, 'a:chOff'), ce = kid(x, 'a:chExt');
    return { x: pxE(off.getAttribute('x')), y: pxE(off.getAttribute('y')), w: pxE(ext.getAttribute('cx')), h: pxE(ext.getAttribute('cy')),
      rot: parseInt(x.getAttribute('rot') || '0', 10) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1',
      cx: co ? pxE(co.getAttribute('x')) : 0, cy: co ? pxE(co.getAttribute('y')) : 0, cw: ce ? pxE(ce.getAttribute('cx')) : 0, ch: ce ? pxE(ce.getAttribute('cy')) : 0 };
  }
  function phOf(sp) { var nv = kid(sp, 'p:nvSpPr') || kid(sp, 'p:nvPicPr') || kid(sp, 'p:nvGraphicFramePr'); var ph = path(nv, ['p:nvPr', 'p:ph']); return ph ? { type: ph.getAttribute('type'), idx: ph.getAttribute('idx') } : null; }
  function normType(t) { if (!t || t === 'body' || t === 'subTitle' || t === 'obj') return 'body'; if (t === 'ctrTitle' || t === 'title') return 'title'; return t; }
  function findPh(tree, ph) {
    if (!tree || !ph) return null;
    var shapes = kids(tree, 'p:sp').concat(kids(tree, 'p:pic'), kids(tree, 'p:graphicFrame'));
    var byIdx = null, byType = null;
    shapes.forEach(function (s) { var p = phOf(s); if (!p) return; if (ph.idx && p.idx === ph.idx && !byIdx) byIdx = s; if (normType(p.type) === normType(ph.type) && !byType) byType = s; });
    return byIdx || byType;
  }
  function spPrOf(sp) { return kid(sp, 'p:spPr'); }

  // ---- text
  function lvlProps(containers, lvl) {
    var out = {}, tag = 'a:lvl' + (lvl + 1) + 'pPr';
    containers.forEach(function (c) {
      if (!c) return; var p = kid(c, tag); if (!p) return;
      ['algn', 'marL', 'indent', 'rtl'].forEach(function (k) { if (out[k] === undefined && p.getAttribute(k) !== null) out[k] = p.getAttribute(k); });
      if (out.bu === undefined) { if (kid(p, 'a:buNone')) out.bu = { none: true }; else if (kid(p, 'a:buChar')) out.bu = { ch: kid(p, 'a:buChar').getAttribute('char') }; else if (kid(p, 'a:buAutoNum')) out.bu = { auto: kid(p, 'a:buAutoNum').getAttribute('type') }; }
      var d = kid(p, 'a:defRPr');
      if (d) {
        ['sz', 'b', 'i', 'u'].forEach(function (k) { if (out[k] === undefined && d.getAttribute(k) !== null) out[k] = d.getAttribute(k); });
        if (out.color === undefined) { var cl = colorOf(kid(d, 'a:solidFill')); if (cl) out.color = cl; }
        if (out.font === undefined) { var lt = kid(d, 'a:latin'); if (lt) out.font = lt.getAttribute('typeface'); }
      }
      var ls = path(p, ['a:lnSpc', 'a:spcPct']); if (ls && out.lnSpc === undefined) out.lnSpc = parseInt(ls.getAttribute('val'), 10) / 100000;
      var sb = path(p, ['a:spcBef', 'a:spcPts']); if (sb && out.spcBef === undefined) out.spcBef = parseInt(sb.getAttribute('val'), 10) / 100;
      var sa = path(p, ['a:spcAft', 'a:spcPts']); if (sa && out.spcAft === undefined) out.spcAft = parseInt(sa.getAttribute('val'), 10) / 100;
    });
    return out;
  }
  function fontName(f) { if (!f) return null; if (f === '+mj-lt') return P.theme.majorFont; if (f === '+mn-lt') return P.theme.minorFont; if (f.charAt(0) === '+') return null; return f; }
  var BULLETS = { '•': '•', 'Ø': '➢', 'ü': '✔', 'q': '❑', 'v': '❖', '§': '▪', '–': '–', '-': '-', 'o': 'o' };

  function renderText(txBody, box, ctx) {
    // ctx: {kind:'title'|'body'|'other', containers:[...lstStyle nodes in priority order], fontScale}
    var bodyPr = kid(txBody, 'a:bodyPr'), anchor = attr(bodyPr, 'anchor', ctx.anchor || 't');
    var wrap = attr(bodyPr, 'wrap', 'square');
    var holder = el('div', null, box);
    var hs = holder.style;
    hs.position = 'absolute'; hs.left = '0'; hs.top = '0'; hs.width = '100%'; hs.height = '100%'; hs.display = 'flex'; hs.flexDirection = 'column';
    hs.justifyContent = anchor === 'ctr' ? 'center' : (anchor === 'b' ? 'flex-end' : 'flex-start'); hs.overflow = 'hidden';
    hs.paddingLeft = pxE(attr(bodyPr, 'lIns', 91440)) + 'px'; hs.paddingRight = pxE(attr(bodyPr, 'rIns', 91440)) + 'px';
    hs.paddingTop = pxE(attr(bodyPr, 'tIns', 45720)) + 'px'; hs.paddingBottom = pxE(attr(bodyPr, 'bIns', 45720)) + 'px';
    var scale = 1, na = bodyPr ? kid(bodyPr, 'a:normAutofit') : null;
    if (na && na.getAttribute('fontScale')) scale = parseInt(na.getAttribute('fontScale'), 10) / 100000;
    var own = kid(txBody, 'a:lstStyle');
    var containers = [own].concat(ctx.containers || []);
    var num = {};
    kids(txBody, 'a:p').forEach(function (p) {
      var pPr = kid(p, 'a:pPr'), lvl = parseInt(attr(pPr, 'lvl', '0'), 10) || 0, lp = lvlProps(containers, lvl);
      ['algn', 'marL', 'indent', 'rtl'].forEach(function (k) { if (pPr && pPr.getAttribute(k) !== null) lp[k] = pPr.getAttribute(k); });
      if (pPr) {
        if (kid(pPr, 'a:buNone')) lp.bu = { none: true }; else if (kid(pPr, 'a:buChar')) lp.bu = { ch: kid(pPr, 'a:buChar').getAttribute('char') }; else if (kid(pPr, 'a:buAutoNum')) lp.bu = { auto: kid(pPr, 'a:buAutoNum').getAttribute('type') };
        var ls = path(pPr, ['a:lnSpc', 'a:spcPct']); if (ls) lp.lnSpc = parseInt(ls.getAttribute('val'), 10) / 100000;
        var sb = path(pPr, ['a:spcBef', 'a:spcPts']); if (sb) lp.spcBef = parseInt(sb.getAttribute('val'), 10) / 100;
        var sa = path(pPr, ['a:spcAft', 'a:spcPts']); if (sa) lp.spcAft = parseInt(sa.getAttribute('val'), 10) / 100;
      }
      var div = el('div', null, holder), ds = div.style;
      ds.margin = '0'; ds.wordBreak = 'break-word'; ds.whiteSpace = wrap === 'none' ? 'pre' : 'pre-wrap';
      var marL = pxE(lp.marL || 0), ind = pxE(lp.indent || 0);
      var algn = lp.algn || 'l'; ds.textAlign = { l: 'left', r: 'right', ctr: 'center', just: 'justify', dist: 'justify' }[algn] || 'left';
      if (lp.rtl === '1') { ds.direction = 'rtl'; }
      if (lp.lnSpc) ds.lineHeight = String(lp.lnSpc * 1.2);
      if (lp.spcBef) ds.marginTop = (lp.spcBef * 96 / 72) + 'px'; if (lp.spcAft) ds.marginBottom = (lp.spcAft * 96 / 72) + 'px';
      var runs = [], maxSz = 0;
      var inner = el('span', null, null);
      for (var n = p.firstElementChild; n; n = n.nextElementSibling) {
        if (n.nodeName === 'a:r' || n.nodeName === 'a:fld') {
          var t = kid(n, 'a:t'); var rp = kid(n, 'a:rPr'), s = el('span', null, inner), ss = s.style;
          s.textContent = t ? t.textContent : '';
          var sz = rp && rp.getAttribute('sz') ? parseInt(rp.getAttribute('sz'), 10) : (lp.sz ? parseInt(lp.sz, 10) : 1800);
          var px = sz / 100 * 96 / 72 * scale; ss.fontSize = px + 'px'; maxSz = Math.max(maxSz, px);
          var b = rp && rp.getAttribute('b') !== null ? rp.getAttribute('b') : lp.b, it = rp && rp.getAttribute('i') !== null ? rp.getAttribute('i') : lp.i, u = rp && rp.getAttribute('u') !== null ? rp.getAttribute('u') : lp.u;
          if (b === '1') ss.fontWeight = 'bold'; if (it === '1') ss.fontStyle = 'italic';
          var dec = []; if (u && u !== 'none') dec.push('underline'); if (rp && rp.getAttribute('strike') && rp.getAttribute('strike') !== 'noStrike') dec.push('line-through'); if (dec.length) ss.textDecoration = dec.join(' ');
          var col = rp ? colorOf(kid(rp, 'a:solidFill')) : null; ss.color = col || lp.color || ctx.defColor || '#000000';
          var lat = rp ? kid(rp, 'a:latin') : null, fam = fontName(lat ? lat.getAttribute('typeface') : null) || fontName(lp.font) || P.theme.minorFont;
          var cs = rp ? kid(rp, 'a:cs') : null; if (cs && cs.getAttribute('typeface') && !/^\+/.test(cs.getAttribute('typeface'))) fam = (fam ? '"' + fam + '",' : '') + '"' + cs.getAttribute('typeface') + '"';
          else if (fam) fam = '"' + fam + '"';
          ss.fontFamily = (fam ? fam + ',' : '') + 'Arial, sans-serif';
          if (rp && rp.getAttribute('baseline') && parseInt(rp.getAttribute('baseline'), 10) > 0) { ss.verticalAlign = 'super'; ss.fontSize = (px * 0.7) + 'px'; }
          var hl = rp ? colorOf(kid(rp, 'a:highlight')) : null; if (hl) ss.backgroundColor = hl;
        } else if (n.nodeName === 'a:br') { inner.appendChild(document.createElement('br')); }
      }
      if (!inner.childNodes.length) { var sp = el('span', null, inner); sp.innerHTML = '&nbsp;'; var epr = kid(p, 'a:endParaRPr'); var esz = epr && epr.getAttribute('sz') ? parseInt(epr.getAttribute('sz'), 10) : (lp.sz ? parseInt(lp.sz, 10) : 1800); maxSz = esz / 100 * 96 / 72 * scale; sp.style.fontSize = maxSz + 'px'; }
      var bu = lp.bu, hasBu = bu && !bu.none && inner.textContent.replace(/\u00a0/g, '').trim() !== '';
      if (hasBu) {
        var label;
        if (bu.ch) label = BULLETS[bu.ch] || bu.ch;
        else { num[lvl] = (num[lvl] || 0) + 1; label = num[lvl] + '.'; }
        ds.display = 'flex'; ds.paddingLeft = Math.max(0, marL + ind) + 'px';
        var bl = el('span', null, div); bl.textContent = label; bl.style.cssText = 'flex:none;width:' + Math.max(12, -ind) + 'px;font-size:' + maxSz + 'px;color:' + (lp.color || ctx.defColor || '#000');
        var tx = el('span', null, div); tx.style.flex = '1'; tx.style.minWidth = '0'; while (inner.firstChild) tx.appendChild(inner.firstChild);
      } else {
        if (!hasBu) num[lvl] = 0;
        ds.paddingLeft = marL + 'px'; ds.textIndent = ind + 'px';
        while (inner.firstChild) div.appendChild(inner.firstChild);
      }
    });
  }

  function svgEl(tag) { return document.createElementNS('http://www.w3.org/2000/svg', tag); }
  function custGeomPath(cg, w, h) {
    var pl = path(cg, ['a:pathLst']); var p = pl ? kid(pl, 'a:path') : null; if (!p) return null;
    var pw = parseInt(p.getAttribute('w') || '0', 10) || 1, ph = parseInt(p.getAttribute('h') || '0', 10) || 1, d = [];
    function pt(n) { return (parseInt(n.getAttribute('x'), 10) * w / pw) + ',' + (parseInt(n.getAttribute('y'), 10) * h / ph); }
    for (var c = p.firstElementChild; c; c = c.nextElementSibling) {
      var pts = kids(c, 'a:pt');
      if (c.nodeName === 'a:moveTo' && pts[0]) d.push('M' + pt(pts[0]));
      else if (c.nodeName === 'a:lnTo' && pts[0]) d.push('L' + pt(pts[0]));
      else if (c.nodeName === 'a:cubicBezTo' && pts.length >= 3) d.push('C' + pt(pts[0]) + ' ' + pt(pts[1]) + ' ' + pt(pts[2]));
      else if (c.nodeName === 'a:quadBezTo' && pts.length >= 2) d.push('Q' + pt(pts[0]) + ' ' + pt(pts[1]));
      else if (c.nodeName === 'a:close') d.push('Z');
    }
    return d.join(' ');
  }
  function dashArray(d, w) { if (!d || d === 'solid') return null; var m = { dot: [1, 2], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3] }[d] || [4, 3]; return m.map(function (v) { return v * w; }).join(','); }

  async function renderShape(sp, parent, tf, layers) {
    var spPr = spPrOf(sp), ph = phOf(sp), xf = xfrmOf(spPr);
    var inheritedTx = null;
    if (!xf && ph) {
      var lay = layers.layout ? findPh(layers.layout, ph) : null, mas = layers.master ? findPh(layers.master, ph) : null;
      xf = xfrmOf(lay ? spPrOf(lay) : null) || xfrmOf(mas ? spPrOf(mas) : null);
    }
    if (!xf) return;
    var x = (xf.x - tf.cx) * tf.sx + tf.ox, y = (xf.y - tf.cy) * tf.sy + tf.oy, w = xf.w * tf.sx, h = xf.h * tf.sy;
    var box = el('div', null, parent), bs = box.style;
    bs.position = 'absolute'; bs.left = x + 'px'; bs.top = y + 'px'; bs.width = w + 'px'; bs.height = h + 'px';
    var tr = []; if (xf.rot) tr.push('rotate(' + xf.rot + 'deg)'); if (xf.flipH) tr.push('scaleX(-1)'); if (xf.flipV) tr.push('scaleY(-1)'); if (tr.length) bs.transform = tr.join(' ');
    var style = kid(sp, 'p:style'), geom = spPr ? kid(spPr, 'a:prstGeom') : null, cust = spPr ? kid(spPr, 'a:custGeom') : null, prst = geom ? geom.getAttribute('prst') : 'rect';
    var fill = await fillCss(spPr, style), ln = lineCss(spPr, style);
    if (!fill && ph && !spPr) fill = null;
    var isLine = /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(prst) || sp.nodeName === 'p:cxnSp';
    if (isLine) {
      var svg = svgEl('svg'); svg.setAttribute('width', w); svg.setAttribute('height', h); svg.style.cssText = 'position:absolute;left:0;top:0;overflow:visible';
      var l = svgEl('line'); l.setAttribute('x1', 0); l.setAttribute('y1', 0); l.setAttribute('x2', w); l.setAttribute('y2', h);
      var lc = ln || { w: 1, c: '#000000' }; l.setAttribute('stroke', lc.c); l.setAttribute('stroke-width', lc.w); var da = dashArray(lc.dash, lc.w); if (da) l.setAttribute('stroke-dasharray', da);
      svg.appendChild(l); box.appendChild(svg);
    } else if (cust || ['triangle', 'rtTriangle', 'diamond', 'hexagon', 'pentagon', 'octagon', 'parallelogram', 'trapezoid', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'chevron', 'homePlate', 'star5'].indexOf(prst) >= 0) {
      var s2 = svgEl('svg'); s2.setAttribute('width', w); s2.setAttribute('height', h); s2.style.cssText = 'position:absolute;left:0;top:0;overflow:visible';
      var shape;
      if (cust) { shape = svgEl('path'); var dd = custGeomPath(cust, w, h); if (dd) shape.setAttribute('d', dd); }
      else {
        shape = svgEl('polygon');
        var pts = {
          triangle: [[w / 2, 0], [w, h], [0, h]], rtTriangle: [[0, 0], [0, h], [w, h]], diamond: [[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]],
          hexagon: [[w * .25, 0], [w * .75, 0], [w, h / 2], [w * .75, h], [w * .25, h], [0, h / 2]], pentagon: [[w / 2, 0], [w, h * .38], [w * .81, h], [w * .19, h], [0, h * .38]],
          octagon: [[w * .29, 0], [w * .71, 0], [w, h * .29], [w, h * .71], [w * .71, h], [w * .29, h], [0, h * .71], [0, h * .29]], parallelogram: [[w * .25, 0], [w, 0], [w * .75, h], [0, h]],
          trapezoid: [[w * .2, 0], [w * .8, 0], [w, h], [0, h]], rightArrow: [[0, h * .25], [w * .6, h * .25], [w * .6, 0], [w, h / 2], [w * .6, h], [w * .6, h * .75], [0, h * .75]],
          leftArrow: [[w, h * .25], [w * .4, h * .25], [w * .4, 0], [0, h / 2], [w * .4, h], [w * .4, h * .75], [w, h * .75]], upArrow: [[w * .25, h], [w * .25, h * .4], [0, h * .4], [w / 2, 0], [w, h * .4], [w * .75, h * .4], [w * .75, h]],
          downArrow: [[w * .25, 0], [w * .25, h * .6], [0, h * .6], [w / 2, h], [w, h * .6], [w * .75, h * .6], [w * .75, 0]], chevron: [[0, 0], [w * .7, 0], [w, h / 2], [w * .7, h], [0, h], [w * .3, h / 2]],
          homePlate: [[0, 0], [w * .75, 0], [w, h / 2], [w * .75, h], [0, h]], star5: [[w * .5, 0], [w * .62, h * .36], [w, h * .38], [w * .69, h * .6], [w * .81, h], [w * .5, h * .77], [w * .19, h], [w * .31, h * .6], [0, h * .38], [w * .38, h * .36]]
        }[prst];
        shape.setAttribute('points', pts.map(function (p) { return p[0] + ',' + p[1]; }).join(' '));
      }
      var fc = fill && (fill.backgroundColor || null);
      shape.setAttribute('fill', fc || 'none'); if (ln) { shape.setAttribute('stroke', ln.c); shape.setAttribute('stroke-width', ln.w); } else shape.setAttribute('stroke', 'none');
      if (fill && fill.background) { var gid = 'g' + Math.random().toString(36).slice(2); var m = /linear-gradient\(([\d.\-]+)deg,(.*)\)/.exec(fill.background); if (m) { var defs = svgEl('defs'), lg = svgEl('linearGradient'); lg.setAttribute('id', gid); var a = parseFloat(m[1]) - 90; lg.setAttribute('gradientTransform', 'rotate(' + a + ' .5 .5)'); (m[2].match(/(rgba?\([^)]*\)|#[0-9a-fA-F]+) [\d.]+%/g) || []).forEach(function (s) { var q = s.split(' '), st = svgEl('stop'); st.setAttribute('offset', q[1]); st.setAttribute('stop-color', q[0]); lg.appendChild(st); }); defs.appendChild(lg); s2.appendChild(defs); shape.setAttribute('fill', 'url(#' + gid + ')'); } }
      s2.appendChild(shape); box.appendChild(s2);
    } else {
      if (fill) for (var k in fill) bs[k] = fill[k];
      if (ln) { bs.border = ln.w + 'px ' + (ln.dash && ln.dash !== 'solid' ? 'dashed' : 'solid') + ' ' + ln.c; }
      if (prst === 'ellipse') bs.borderRadius = '50%';
      else if (prst === 'roundRect') { var av = path(geom, ['a:avLst']); var gd = av ? kid(av, 'a:gd') : null; var adj = gd ? parseInt((gd.getAttribute('fmla') || '').split(' ')[1], 10) : 16667; bs.borderRadius = (Math.min(w, h) * adj / 100000) + 'px'; }
    }
    var tx = kid(sp, 'p:txBody');
    if (tx && (tx.textContent || '').trim() !== '') {
      var kind = ph ? (normType(ph.type) === 'title' ? 'title' : (normType(ph.type) === 'body' ? 'body' : 'other')) : 'other';
      var lay2 = ph && layers.layout ? findPh(layers.layout, ph) : null, mas2 = ph && layers.master ? findPh(layers.master, ph) : null;
      var cont = [];
      if (lay2) cont.push(path(lay2, ['p:txBody', 'a:lstStyle']));
      if (mas2) cont.push(path(mas2, ['p:txBody', 'a:lstStyle']));
      if (ph) cont.push(layers.txStyles ? kid(layers.txStyles, kind === 'title' ? 'p:titleStyle' : (kind === 'body' ? 'p:bodyStyle' : 'p:otherStyle')) : null);
      else cont.push(layers.txStyles ? kid(layers.txStyles, 'p:otherStyle') : null);
      cont.push(P.defText);
      var anchor = null; if (lay2) anchor = attr(path(lay2, ['p:txBody', 'a:bodyPr']), 'anchor', null); if (!anchor && mas2) anchor = attr(path(mas2, ['p:txBody', 'a:bodyPr']), 'anchor', null);
      var defColor = null;
      var fr = style ? kid(style, 'a:fontRef') : null; if (fr) defColor = colorOf(fr);
      renderText(tx, box, { containers: cont, kind: kind, anchor: anchor, defColor: defColor || (fill && fill.backgroundColor && !ph ? null : null) });
    }
  }

  async function renderPic(pic, parent, tf) {
    var spPr = spPrOf(pic), xf = xfrmOf(spPr); if (!xf) return;
    var bf = kid(pic, 'p:blipFill'), url = bf ? await blipUrl(bf) : null; if (!url) return;
    var x = (xf.x - tf.cx) * tf.sx + tf.ox, y = (xf.y - tf.cy) * tf.sy + tf.oy, w = xf.w * tf.sx, h = xf.h * tf.sy;
    var box = el('div', null, parent), bs = box.style; bs.position = 'absolute'; bs.left = x + 'px'; bs.top = y + 'px'; bs.width = w + 'px'; bs.height = h + 'px'; bs.overflow = 'hidden';
    var tr = []; if (xf.rot) tr.push('rotate(' + xf.rot + 'deg)'); if (xf.flipH) tr.push('scaleX(-1)'); if (xf.flipV) tr.push('scaleY(-1)'); if (tr.length) bs.transform = tr.join(' ');
    var sr = kid(bf, 'a:srcRect'), l = sr ? parseInt(sr.getAttribute('l') || '0', 10) / 100000 : 0, t = sr ? parseInt(sr.getAttribute('t') || '0', 10) / 100000 : 0, r = sr ? parseInt(sr.getAttribute('r') || '0', 10) / 100000 : 0, b = sr ? parseInt(sr.getAttribute('b') || '0', 10) / 100000 : 0;
    var iw = w / (1 - l - r), ih = h / (1 - t - b);
    var img = el('img', null, box); img.src = url; img.style.cssText = 'position:absolute;max-width:none;width:' + iw + 'px;height:' + ih + 'px;left:' + (-l * iw) + 'px;top:' + (-t * ih) + 'px';
  }

  function renderTable(gf, parent, tf, layers) {
    var xe = kid(gf, 'p:xfrm'); if (!xe) return;
    var off = kid(xe, 'a:off'), ext = kid(xe, 'a:ext'); if (!off || !ext) return;
    var x = (pxE(off.getAttribute('x')) - tf.cx) * tf.sx + tf.ox, y = (pxE(off.getAttribute('y')) - tf.cy) * tf.sy + tf.oy, w = pxE(ext.getAttribute('cx')) * tf.sx, h = pxE(ext.getAttribute('cy')) * tf.sy;
    var gd = path(gf, ['a:graphic', 'a:graphicData']), tbl = gd ? kid(gd, 'a:tbl') : null;
    var box = el('div', null, parent), bs = box.style; bs.position = 'absolute'; bs.left = x + 'px'; bs.top = y + 'px'; bs.width = w + 'px'; bs.height = h + 'px';
    if (!tbl) { bs.border = '1px dashed #999'; bs.color = '#666'; bs.font = '14px sans-serif'; bs.padding = '8px'; bs.background = 'rgba(255,255,255,.6)'; box.textContent = '[Chart / diagram – not rendered]'; return; }
    var pr = kid(tbl, 'a:tblPr'), firstRow = pr && pr.getAttribute('firstRow') === '1', band = pr && pr.getAttribute('bandRow') === '1';
    var table = el('table', null, box), ts = table.style; ts.borderCollapse = 'collapse'; ts.tableLayout = 'fixed'; ts.position = 'absolute'; ts.left = '0'; ts.top = '0';
    var cols = kids(kid(tbl, 'a:tblGrid'), 'a:gridCol').map(function (g) { return pxE(g.getAttribute('w')) * tf.sx; });
    ts.width = cols.reduce(function (a, b) { return a + b; }, 0) + 'px';
    var cg = el('colgroup', null, table); cols.forEach(function (cw) { el('col', null, cg).style.width = cw + 'px'; });
    var acc1 = P.theme.accent1 || '#4472C4';
    kids(tbl, 'a:tr').forEach(function (trn, ri) {
      var tr = el('tr', null, table); tr.style.height = (pxE(trn.getAttribute('h')) * tf.sy) + 'px';
      kids(trn, 'a:tc').forEach(function (tc) {
        if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') return;
        var td = el('td', null, tr), st = td.style, tp = kid(tc, 'a:tcPr');
        if (tc.getAttribute('gridSpan')) td.colSpan = parseInt(tc.getAttribute('gridSpan'), 10); if (tc.getAttribute('rowSpan')) td.rowSpan = parseInt(tc.getAttribute('rowSpan'), 10);
        var fillc = tp ? colorOf(kid(tp, 'a:solidFill')) : null, hdr = firstRow && ri === 0;
        if (!fillc && !(tp && kid(tp, 'a:noFill'))) fillc = hdr ? acc1 : (band ? ((ri % 2 === (firstRow ? 1 : 0)) ? lumMod(acc1, 0.2, 0.8) : lumMod(acc1, 0.4, 0.6)) : null);
        if (fillc) st.backgroundColor = fillc;
        st.border = '1px solid ' + (hdr || fillc ? '#ffffff' : '#bbbbbb'); st.verticalAlign = attr(tp, 'anchor', 't') === 'ctr' ? 'middle' : 'top'; st.overflow = 'hidden'; st.padding = '0'; st.position = 'relative';
        var cell = el('div', null, td); cell.style.cssText = 'position:relative;width:100%;min-height:' + Math.max(10, pxE(trn.getAttribute('h')) * tf.sy - 2) + 'px';
        var cellBox = el('div', null, cell); cellBox.style.cssText = 'padding:3px 6px';
        var txb = kid(tc, 'a:txBody');
        if (txb) {
          var tmp = el('div', null, null); tmp.style.position = 'relative';
          renderText(txb, tmp, { containers: [P.defText], kind: 'other', anchor: 't', defColor: hdr ? '#ffffff' : '#000000' });
          var h2 = tmp.firstChild; if (h2) { h2.style.position = 'static'; h2.style.height = 'auto'; h2.style.width = 'auto'; h2.style.padding = '0'; if (hdr) h2.style.fontWeight = 'bold'; cellBox.appendChild(h2); }
        }
      });
    });
  }

  async function renderTree(tree, parent, tf, layers, skipPlaceholders) {
    if (!tree) return;
    for (var n = tree.firstElementChild; n; n = n.nextElementSibling) {
      try {
        if (n.nodeName === 'p:sp' || n.nodeName === 'p:cxnSp') { if (skipPlaceholders && phOf(n)) continue; await renderShape(n, parent, tf, layers); }
        else if (n.nodeName === 'p:pic') { if (skipPlaceholders && phOf(n)) continue; await renderPic(n, parent, tf); }
        else if (n.nodeName === 'p:graphicFrame') { renderTable(n, parent, tf, layers); }
        else if (n.nodeName === 'p:grpSp') {
          var xf = xfrmOf(kid(n, 'p:grpSpPr')); if (!xf) continue;
          var sx = xf.cw ? xf.w / xf.cw : 1, sy = xf.ch ? xf.h / xf.ch : 1;
          var ntf = { sx: tf.sx * sx, sy: tf.sy * sy, cx: xf.cx, cy: xf.cy, ox: (xf.x - tf.cx) * tf.sx + tf.ox, oy: (xf.y - tf.cy) * tf.sy + tf.oy };
          await renderTree(n, parent, ntf, layers, skipPlaceholders);
        }
      } catch (e) { /* skip a broken shape, keep the rest */ }
    }
  }
  async function bgOf(cSld) {
    var bg = kid(cSld, 'p:bg'); if (!bg) return null;
    var pr = kid(bg, 'p:bgPr'); if (pr) return await fillCss(pr, null);
    var ref = kid(bg, 'p:bgRef'); if (ref) { var c = colorOf(ref); return c ? { backgroundColor: c } : null; }
    return null;
  }

  async function renderPptx(buf) {
    need('jszip', window.JSZip);
    P.zip = await window.JSZip.loadAsync(buf); P.media = {};
    var pres = await pxml('ppt/presentation.xml'); if (!pres) throw new Error('not a valid .pptx file');
    var sz = pres.getElementsByTagName('p:sldSz')[0]; P.W = sz ? pxE(sz.getAttribute('cx')) : 960; P.H = sz ? pxE(sz.getAttribute('cy')) : 540;
    P.defText = pres.getElementsByTagName('p:defaultTextStyle')[0] || null;
    var presRels = await prels('ppt/presentation.xml'), ids = pres.getElementsByTagName('p:sldId'), slidePaths = [];
    for (var i = 0; i < ids.length; i++) { var rel = presRels[ids[i].getAttribute('r:id')]; if (rel) slidePaths.push(rel.target); }
    if (!slidePaths.length) throw new Error('presentation has no slides');
    var cache = {};
    for (var si = 0; si < slidePaths.length; si++) {
      var sp = slidePaths[si], sdoc = await pxml(sp); if (!sdoc) continue;
      var srels = await prels(sp), layoutRel = null; for (var k in srels) if (srels[k].type === 'slideLayout') layoutRel = srels[k];
      var layoutPath = layoutRel ? layoutRel.target : null, L = null, M = null;
      if (layoutPath) {
        if (!cache[layoutPath]) {
          var ldoc = await pxml(layoutPath), lrels = await prels(layoutPath), mrel = null; for (var k2 in lrels) if (lrels[k2].type === 'slideMaster') mrel = lrels[k2];
          var mpath = mrel ? mrel.target : null, mdoc = null, mrels = {};
          if (mpath) {
            if (!cache[mpath]) {
              mdoc = await pxml(mpath); mrels = await prels(mpath);
              var th = null; for (var k3 in mrels) if (mrels[k3].type === 'theme') th = mrels[k3];
              cache[mpath] = { doc: mdoc, rels: mrels, theme: th ? await pxml(th.target) : null };
            }
          }
          cache[layoutPath] = { doc: ldoc, rels: lrels, master: mpath ? cache[mpath] : null };
        }
        L = cache[layoutPath]; M = L.master;
      }
      // theme + colour map (per master)
      P.theme = {}; P.clrMap = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
      if (M && M.theme) {
        var names = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
        names.forEach(function (nm) { var nd = M.theme.getElementsByTagName('a:' + nm)[0], c = nd ? nd.firstElementChild : null; if (c) P.theme[nm] = '#' + (c.getAttribute('lastClr') || c.getAttribute('val') || '000000'); });
        var mj = M.theme.getElementsByTagName('a:majorFont')[0], mn = M.theme.getElementsByTagName('a:minorFont')[0];
        P.theme.majorFont = mj && kid(mj, 'a:latin') ? kid(mj, 'a:latin').getAttribute('typeface') : null; P.theme.minorFont = mn && kid(mn, 'a:latin') ? kid(mn, 'a:latin').getAttribute('typeface') : null;
      }
      if (M && M.doc) { var cm = M.doc.getElementsByTagName('p:clrMap')[0]; if (cm) for (var ai = 0; ai < cm.attributes.length; ai++) P.clrMap[cm.attributes[ai].name] = cm.attributes[ai].value; }
      var slide = el('div', 'slide', root), ss = slide.style; ss.width = P.W + 'px'; ss.height = P.H + 'px';
      var sCSld = sdoc.getElementsByTagName('p:cSld')[0], lCSld = L && L.doc ? L.doc.getElementsByTagName('p:cSld')[0] : null, mCSld = M && M.doc ? M.doc.getElementsByTagName('p:cSld')[0] : null;
      CUR_RELS = srels; var bg = await bgOf(sCSld);
      if (!bg && lCSld) { CUR_RELS = L.rels; bg = await bgOf(lCSld); }
      if (!bg && mCSld) { CUR_RELS = M.rels; bg = await bgOf(mCSld); }
      if (bg) for (var bk in bg) ss[bk] = bg[bk];
      var txStyles = M && M.doc ? M.doc.getElementsByTagName('p:txStyles')[0] : null;
      var layersLM = { layout: lCSld ? kid(lCSld, 'p:spTree') : null, master: mCSld ? kid(mCSld, 'p:spTree') : null, txStyles: txStyles };
      var idTf = { sx: 1, sy: 1, cx: 0, cy: 0, ox: 0, oy: 0 };
      var showMaster = !(L && L.doc && L.doc.documentElement.getAttribute('showMasterSp') === '0');
      if (showMaster && mCSld) { CUR_RELS = M.rels; await renderTree(kid(mCSld, 'p:spTree'), slide, idTf, { layout: null, master: null, txStyles: txStyles }, true); }
      if (lCSld) { CUR_RELS = L.rels; await renderTree(kid(lCSld, 'p:spTree'), slide, idTf, { layout: null, master: layersLM.master, txStyles: txStyles }, true); }
      CUR_RELS = srels; await renderTree(kid(sCSld, 'p:spTree'), slide, idTf, layersLM, false);
      pages.push({ el: slide, w: P.W, h: P.H });
      if (si % 5 === 4) await nextFrame();
    }
    setViewport(P.W + 16);
  }

  /* ------------------------------------------------------------------ text + navigation */
  function extractText() {
    var parts = [];
    if (type === 'xlsx') {
      var sheets = root.querySelectorAll('.sheet');
      for (var i = 0; i < sheets.length; i++) {
        var t = sheets[i].querySelector('.sheetTitle'); parts.push('=== ' + (t ? t.textContent : 'Sheet ' + (i + 1)) + ' ===');
        var rows = sheets[i].querySelectorAll('tr');
        for (var r = 0; r < rows.length; r++) { var cells = rows[r].children, line = []; for (var c = 0; c < cells.length; c++) line.push(cells[c].textContent.replace(/[\t\r\n]+/g, ' ')); parts.push(line.join('\t').replace(/\t+$/, '')); }
        parts.push('');
      }
    } else if (type === 'pptx') {
      pages.forEach(function (p, i) { parts.push('--- Slide ' + (i + 1) + ' ---'); parts.push((p.el.innerText || '').trim()); parts.push(''); });
    } else {
      pages.forEach(function (p, i) { if (i) parts.push(''); parts.push((p.el.innerText || '').replace(/\u00a0/g, ' ')); });
    }
    return parts.join('\n');
  }
  window.getAllText = function () { return extractText(); };
  window.goPage = function (n) { var p = pages[n - 1]; if (!p) return; var e = p.el; if (p.slice && p.y) { window.scrollTo(0, absTop(e) + p.y - 8); } else { e.scrollIntoView({ block: 'start' }); } };
  var lastReported = -1, ticking = false;
  function currentPage() {
    var vv = window.visualViewport, vy = (vv ? vv.pageTop : (window.pageYOffset || 0)) + (vv ? vv.height : window.innerHeight) * 0.35;
    var found = 0;
    for (var i = 0; i < pages.length; i++) { var p = pages[i], top = absTop(p.el) + (p.slice ? p.y : 0); if (top <= vy) found = i; else break; }
    return found;
  }
  function onScroll() { if (ticking) return; ticking = true; requestAnimationFrame(function () { ticking = false; var c = currentPage(); if (c !== lastReported) { lastReported = c; try { B.onPage(c + 1); } catch (e) {} } }); }

  /* ------------------------------------------------------------------ export */
  window.cancelExport = function () { cancelFlag = true; };
  window.exportText = function () { try { B.saveText(extractText()); } catch (e) { try { B.exportFailed(String(e)); } catch (x) {} } };
  window.copyAllText = function () { try { B.copyText(extractText()); } catch (e) { try { B.exportFailed(String(e)); } catch (x) {} } };
  window.exportPages = async function (reqScale) {
    cancelFlag = false;
    try {
      need('html2canvas', window.html2canvas);
      if (document.fonts && document.fonts.ready) await document.fonts.ready;
      window.scrollTo(0, 0);
      var n = pages.length;
      for (var i = 0; i < n; i++) {
        if (cancelFlag) { B.exportFailed('cancelled'); return; }
        B.progress(i + 1, n);
        await nextFrame();
        var p = pages[i], w = Math.max(1, Math.round(p.w)), h = Math.max(1, Math.round(p.h));
        var sc = Math.min(reqScale, 4096 / Math.max(w, h), Math.sqrt(16000000 / (w * h)));
        var opts = { scale: sc, backgroundColor: '#ffffff', useCORS: true, allowTaint: true, logging: false, scrollX: 0, scrollY: 0,
          windowWidth: Math.max(window.innerWidth, document.documentElement.scrollWidth, w + 40), windowHeight: Math.max(window.innerHeight, h + 40) };
        if (p.slice) { opts.x = 0; opts.y = p.y; opts.width = w; opts.height = h; }
        var canvas = await window.html2canvas(p.el, opts);
        var url = canvas.toDataURL('image/png');
        canvas.width = 1; canvas.height = 1;
        B.savePage(i, w, h, url.substring(url.indexOf(',') + 1));
      }
      B.exportDone();
    } catch (e) { try { B.exportFailed((e && e.message) ? e.message : String(e)); } catch (x) {} }
  };

  /* ------------------------------------------------------------------ boot */
  window.addEventListener('scroll', onScroll, { passive: true });
  if (window.visualViewport) window.visualViewport.addEventListener('scroll', onScroll);
  (async function () {
    try {
      if (!type) throw new Error('unknown file type');
      var r = await fetch('/doc'); if (!r.ok) throw new Error('cannot read file (' + r.status + ')');
      var buf = await r.arrayBuffer();
      if (type === 'docx') await renderDocx(buf); else if (type === 'xlsx') await renderXlsx(buf); else if (type === 'pptx') await renderPptx(buf); else throw new Error('unsupported type ' + type);
      msg.style.display = 'none';
      await nextFrame();
      B.onReady(pages.length, type);
      onScroll();
    } catch (e) { fail(e); }
  })();
})();
