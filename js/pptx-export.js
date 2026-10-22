'use strict';
// PowerPoint（.pptx）への書き出し。PowerPoint で開いて編集できる形にする：
//   テキスト → テキストボックス（装飾・箇条書き・字下げ・揃え・行間）
//   ペンの線 → フリーフォームの線（筆圧による太さの変化は一定の太さになる）
//   直線・矢印 → コネクタ、四角・楕円 → 図形、画像 → 画像（SVG などは PNG にして埋め込む）
//   表 → 表、ページの背景 → スライドの背景
// 色は書き出すテーマ（白地・黒板）の実際の色にする。外部ライブラリは使わない（zip は無圧縮で作る）。
(function (B) {
  const EMU = 6350;              // 論理座標 1px = 6350 EMU（1920px ＝ 13.333 インチ）
  const SW = B.W * EMU, SH = B.H * EMU;
  const e = (px) => Math.round(px * EMU);
  const xml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  // ---------------- zip（無圧縮） ----------------
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(u8) { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }

  function zip(files) {
    const enc = new TextEncoder();
    const parts = [], central = [];
    let offset = 0;
    for (const [name, data] of files) {
      const body = typeof data === 'string' ? enc.encode(data) : data;
      const nm = enc.encode(name);
      const crc = crc32(body);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, body.length, true); lh.setUint32(22, body.length, true); lh.setUint16(26, nm.length, true);
      parts.push(new Uint8Array(lh.buffer), nm, body);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, body.length, true); ch.setUint32(24, body.length, true); ch.setUint16(28, nm.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), nm);
      offset += 30 + nm.length + body.length;
    }
    const csize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, csize, true); end.setUint32(16, offset, true);
    return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }

  // ---------------- 色 ----------------
  // テーマの CSS 変数から実際の色を読む（意味上の色名 → #rrggbb）
  function makePalette(theme) {
    const probe = document.createElement('div');
    probe.className = 'bpage theme-' + theme;
    probe.style.cssText = 'position:absolute;left:-99999px;top:0';
    document.body.appendChild(probe);
    const cs = getComputedStyle(probe);
    const v = (n) => cs.getPropertyValue(n).trim();
    const pal = { bg: v('--bg'), hlOpacity: parseFloat(v('--hl-opacity')) || 0.4, c: {}, hl: {}, hlbg: {} };
    for (const n of ['default', 'red', 'blue', 'green', 'gray']) pal.c[n] = v('--c-' + n);
    for (const n of ['yellow', 'pink', 'blue', 'green']) { pal.hl[n] = v('--hl-' + n); pal.hlbg[n] = v('--hlbg-' + n); }
    probe.remove();
    return pal;
  }

  function toHex(css) {
    const c = document.createElement('canvas').getContext('2d');
    c.fillStyle = '#000';
    c.fillStyle = css;
    const s = c.fillStyle;
    if (s[0] === '#') return { hex: s.slice(1).toUpperCase(), alpha: 1 };
    const m = /rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/.exec(s);
    if (!m) return { hex: '000000', alpha: 1 };
    return { hex: [m[1], m[2], m[3]].map((x) => Math.round(+x).toString(16).padStart(2, '0')).join('').toUpperCase(), alpha: m[4] != null ? +m[4] : 1 };
  }

  // 半透明の色を背景に重ねた不透明色（PowerPoint の蛍光ペン［文字の強調］は不透明のため）
  function over(fg, bg) {
    const a = fg.alpha;
    const f = [0, 2, 4].map((i) => parseInt(fg.hex.substr(i, 2), 16));
    const b = [0, 2, 4].map((i) => parseInt(bg.hex.substr(i, 2), 16));
    return f.map((x, i) => Math.round(x * a + b[i] * (1 - a)).toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  // ---------------- 文字 ----------------
  function colorEl(hex, alpha) {
    return `<a:solidFill><a:srgbClr val="${hex}">${alpha != null && alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ''}</a:srgbClr></a:solidFill>`;
  }

  // 板書の HTML（rich.js の決まった形）→ <a:p> の並び
  function paragraphs(html, fs, ctx, opts) {
    const flat = B.rich.parseHtml(html);
    const lines = B.rich.lineList(flat);
    const out = [];
    const defLh = opts.lh || 1.4;
    for (const ln of lines) {
      const la = B.rich.lineAttrsOf(flat, ln);
      const em = (v) => e((v || 0) * fs);
      const ppr = [];
      const attrs = [];
      if (la.pl) attrs.push(`marL="${em(la.pl)}"`);
      if (la.lv) attrs.push(`lvl="${la.lv}"`);
      if (la.ti) attrs.push(`indent="${em(la.ti)}"`);
      const al = la.ta || opts.align;
      if (al) attrs.push(`algn="${al === 'center' ? 'ctr' : al === 'right' ? 'r' : 'l'}"`);
      ppr.push(`<a:lnSpc><a:spcPct val="${Math.round((defLh / 1.2) * 100000)}"/></a:lnSpc>`);
      if (la.mt) ppr.push(`<a:spcBef><a:spcPts val="${Math.round(la.mt * fs / 2 * 100)}"/></a:spcBef>`);
      let prefix = '';
      const AUTO = { '1.': 'arabicPeriod', '1)': 'arabicParenR', '(1)': 'arabicParenBoth', 'i.': 'romanLcPeriod', 'i)': 'romanLcParenR', 'I.': 'romanUcPeriod', 'a.': 'alphaLcPeriod', 'a)': 'alphaLcParenR', 'A.': 'alphaUcPeriod', 'A)': 'alphaUcParenR' };
      if (la.ol && AUTO[la.ol]) {
        if (la.bc) ppr.push(`<a:buClr><a:srgbClr val="${ctx.color(la.bc).hex}"/></a:buClr>`);
        ppr.push(`<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${AUTO[la.ol]}"/>`);
      } else if (la.bu && la.bu.length === 1) {
        if (la.bc) ppr.push(`<a:buClr><a:srgbClr val="${ctx.color(la.bc).hex}"/></a:buClr>`);
        if (la.bs) ppr.push(`<a:buSzPct val="${Math.round(la.bs * 100000)}"/>`);
        ppr.push(`<a:buFont typeface="Arial"/><a:buChar char="${xml(la.bu)}"/>`);
      } else {
        if (la.bu) prefix = la.bu + '\t';
        ppr.push('<a:buNone/>');
      }
      let runs = '';
      const chars = flat.slice(ln.start, ln.end);
      if (prefix) runs += run(prefix, chars[0] ? chars[0].a : {}, fs, ctx, opts);
      for (let i = 0; i < chars.length;) {
        let j = i + 1;
        while (j < chars.length && sameRun(chars[i].a, chars[j].a)) j++;
        runs += run(chars.slice(i, j).map((c) => c.ch).join(''), chars[i].a, fs, ctx, opts);
        i = j;
      }
      const endSz = Math.round(fs / 2 * 100);
      out.push(`<a:p><a:pPr ${attrs.join(' ')}>${ppr.join('')}</a:pPr>${runs}<a:endParaRPr lang="ja-JP" sz="${endSz}"/></a:p>`);
    }
    return out.join('');
  }

  const RUN_KEYS = ['z', 'f', 'x', 'h', 'c', 'b', 'i', 'u', 'k', 's'];
  const sameRun = (a, b) => RUN_KEYS.every((k) => (a[k] || null) === (b[k] || null));

  function run(text, a, fs, ctx, opts) {
    const px = fs * (a.z || 1);
    const at = [`lang="ja-JP"`, `sz="${Math.max(100, Math.round(px / 2 * 100))}"`];
    if (a.b) at.push('b="1"');
    if (a.i) at.push('i="1"');
    let uFill = '';
    if (a.u) {
      const [st, col] = a.u === true ? ['sng', null] : String(a.u).split(':');
      at.push(`u="${st === 'dbl' ? 'dbl' : 'sng'}"`);
      if (col) uFill = `<a:uFill>${colorEl(ctx.color(col).hex)}</a:uFill>`;
    }
    if (a.k) at.push('strike="sngStrike"');
    if (a.s === 'sup') at.push('baseline="30000"');
    else if (a.s === 'sub') at.push('baseline="-25000"');
    const col = ctx.color(a.c || 'default');
    let inner = colorEl(col.hex);
    if (a.h) inner += `<a:highlight><a:srgbClr val="${ctx.highlight(a.h)}"/></a:highlight>`;
    inner += uFill;
    const face = a.f || (opts.font === 'ja' ? ctx.jaFont : null);
    if (face) inner += `<a:latin typeface="${xml(face)}"/><a:ea typeface="${xml(face)}"/><a:cs typeface="${xml(face)}"/>`;
    if (a.x) ctx.warn('box');
    return `<a:r><a:rPr ${at.join(' ')}>${inner}</a:rPr><a:t>${xml(text)}</a:t></a:r>`;
  }

  // ---------------- 図形 ----------------
  function nv(tag, id, name, extra) {
    return `<p:nv${tag}Pr><p:cNvPr id="${id}" name="${xml(name)}"/><p:cNv${tag}Pr${extra || ''}/><p:nvPr/></p:nv${tag}Pr>`;
  }

  function lineEl(wPx, color, opts) {
    const o = opts || {};
    return `<a:ln w="${e(wPx)}" cap="rnd">${colorEl(color.hex, color.alpha)}${o.dash ? '' : '<a:prstDash val="solid"/>'}<a:round/>${o.tail ? '<a:tailEnd type="triangle" w="med" len="med"/>' : ''}</a:ln>`;
  }

  function textBox(o, id, ctx) {
    const m = ctx.measure(o);
    const w = o.w || m.w;
    return `<p:sp>${nv('Sp', id, 'テキスト ' + id, ' txBox="1"')}<p:spPr><a:xfrm><a:off x="${e(o.x)}" y="${e(o.y)}"/><a:ext cx="${e(w)}" cy="${e(m.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
      + `<p:txBody><a:bodyPr wrap="${o.w ? 'square' : 'none'}" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"${o.cols ? ` numCol="${o.cols}" spcCol="${e(o.colGap || 0)}"` : ''}><a:noAutofit/></a:bodyPr><a:lstStyle/>`
      + paragraphs(o.html, o.fs, ctx, { lh: o.lh, align: o.align, font: o.font }) + '</p:txBody></p:sp>';
  }

  function strokeShape(o, id, ctx) {
    const pts = o.pts;
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const p of pts) { x0 = Math.min(x0, p[0]); y0 = Math.min(y0, p[1]); x1 = Math.max(x1, p[0]); y1 = Math.max(y1, p[1]); }
    const w = Math.max(1, x1 - x0), h = Math.max(1, y1 - y0);
    const marker = o.kind === 'marker';
    const col = marker ? Object.assign({}, ctx.color(o.color, 'hl'), { alpha: ctx.pal.hlOpacity }) : ctx.color(o.color);
    const P = (p) => `<a:pt x="${e(p[0] - x0)}" y="${e(p[1] - y0)}"/>`;
    let path = `<a:moveTo>${P(pts[0])}</a:moveTo>`;
    if (pts.length === 1) path += `<a:lnTo>${P([pts[0][0] + 0.5, pts[0][1]])}</a:lnTo>`;
    for (let i = 1; i < pts.length; i++) path += `<a:lnTo>${P(pts[i])}</a:lnTo>`;
    const width = o.pr ? o.size * 0.9 : o.size;
    return `<p:sp>${nv('Sp', id, (marker ? '蛍光ペン ' : 'ペン ') + id)}<p:spPr><a:xfrm><a:off x="${e(x0)}" y="${e(y0)}"/><a:ext cx="${e(w)}" cy="${e(h)}"/></a:xfrm>`
      + `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${e(w)}" h="${e(h)}" fill="none">${path}</a:path></a:pathLst></a:custGeom>`
      + `<a:noFill/>${lineEl(width, col)}</p:spPr></p:sp>`;
  }

  function shapeObj(o, id, ctx) {
    const col = ctx.color(o.color);
    if (o.shape === 'line' || o.shape === 'arrow') {
      const x = Math.min(o.x1, o.x2), y = Math.min(o.y1, o.y2);
      const flipH = o.x2 < o.x1, flipV = o.y2 < o.y1;
      return `<p:cxnSp>${nv('CxnSp', id, (o.shape === 'arrow' ? '矢印 ' : '直線 ') + id)}<p:spPr><a:xfrm${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}><a:off x="${e(x)}" y="${e(y)}"/><a:ext cx="${e(Math.abs(o.x2 - o.x1))}" cy="${e(Math.abs(o.y2 - o.y1))}"/></a:xfrm>`
        + `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${lineEl(o.width, col, { tail: o.shape === 'arrow' })}</p:spPr></p:cxnSp>`;
    }
    const x = Math.min(o.x1, o.x2), y = Math.min(o.y1, o.y2);
    return `<p:sp>${nv('Sp', id, (o.shape === 'rect' ? '四角形 ' : '楕円 ') + id)}<p:spPr><a:xfrm><a:off x="${e(x)}" y="${e(y)}"/><a:ext cx="${e(Math.abs(o.x2 - o.x1))}" cy="${e(Math.abs(o.y2 - o.y1))}"/></a:xfrm>`
      + `<a:prstGeom prst="${o.shape === 'rect' ? 'rect' : 'ellipse'}"><a:avLst/></a:prstGeom><a:noFill/>${lineEl(o.width, col)}</p:spPr></p:sp>`;
  }

  async function imageObj(o, id, ctx) {
    const media = await ctx.addMedia(o.src, o.w, o.h);
    if (!media) return '';
    return `<p:pic>${nv('Pic', id, '画像 ' + id, '')}<p:blipFill><a:blip r:embed="${media}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
      + `<p:spPr><a:xfrm><a:off x="${e(o.x)}" y="${e(o.y)}"/><a:ext cx="${e(o.w)}" cy="${e(o.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  }

  function tableObj(o, id, ctx) {
    const border = ctx.color(o.color || 'default');
    const bw = e(o.bw || 2);
    const ln = (n) => `<a:${n} w="${bw}" cap="flat" cmpd="sng" algn="ctr">${colorEl(border.hex)}<a:prstDash val="solid"/><a:round/></a:${n}>`;
    const total = o.cols.reduce((a, b) => a + b, 0);
    const m = ctx.measureEl(o);
    let rows = '';
    o.cells.forEach((row, r) => {
      rows += `<a:tr h="${e(m.rows[r] || o.rows[r])}">`;
      row.forEach((cell) => {
        const body = paragraphs(cell.html || '', o.fs, ctx, { lh: o.lh, align: cell.al, font: o.font }) || '<a:p><a:endParaRPr lang="ja-JP"/></a:p>';
        const anchor = cell.va === 'middle' ? ' anchor="ctr"' : cell.va === 'bottom' ? ' anchor="b"' : '';
        rows += `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${body}</a:txBody><a:tcPr marL="${e(12)}" marR="${e(12)}" marT="${e(6)}" marB="${e(6)}"${anchor}>${ln('lnL')}${ln('lnR')}${ln('lnT')}${ln('lnB')}${cell.fill ? colorEl(ctx.color(cell.fill).hex) : '<a:noFill/>'}</a:tcPr></a:tc>`;
      });
      rows += '</a:tr>';
    });
    return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="表 ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
      + `<p:xfrm><a:off x="${e(o.x)}" y="${e(o.y)}"/><a:ext cx="${e(total)}" cy="${e(m.h)}"/></p:xfrm>`
      + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>${o.cols.map((w) => `<a:gridCol w="${e(w)}"/>`).join('')}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
  }

  // ---------------- 画像の取り込み ----------------
  function dataUrlBytes(src) {
    const m = /^data:([^;,]+)(;base64)?,(.*)$/.exec(src);
    if (!m) return null;
    const bin = m[2] ? atob(m[3]) : decodeURIComponent(m[3]);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return { type: m[1], bytes: u8 };
  }

  // PNG / JPEG / GIF はそのまま、SVG などは PNG に描き直す
  async function toMedia(src, w, h) {
    const d = dataUrlBytes(src);
    if (!d) return null;
    if (d.type === 'image/png') return { ext: 'png', bytes: d.bytes };
    if (d.type === 'image/jpeg') return { ext: 'jpeg', bytes: d.bytes };
    if (d.type === 'image/gif') return { ext: 'gif', bytes: d.bytes };
    const img = new Image();
    img.src = src;
    await img.decode();
    const scale = d.type === 'image/svg+xml' ? 2 : 1;
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round((d.type === 'image/svg+xml' ? w : img.naturalWidth) * scale));
    c.height = Math.max(1, Math.round((d.type === 'image/svg+xml' ? h : img.naturalHeight) * scale));
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
    return { ext: 'png', bytes: new Uint8Array(await blob.arrayBuffer()) };
  }

  // ---------------- 寸法を測る ----------------
  function makeMeasurer(theme) {
    const host = document.createElement('div');
    host.style.cssText = 'position:absolute;left:-100000px;top:0;width:1920px;height:1080px;visibility:hidden';
    document.body.appendChild(host);
    const fake = { on: () => () => {}, doc: { theme }, page: { objects: [] }, pageById: () => null };
    const view = new B.View(fake, document, host, {});
    const put = (o) => { const el = view.create(o); view.root.appendChild(el); return el; };
    return {
      measure(o) { const el = put(o); const r = { w: el.offsetWidth, h: el.offsetHeight }; el.remove(); return r; },
      measureEl(o) {
        const el = put(o);
        const rows = Array.from(el.querySelectorAll('tr')).map((tr) => tr.offsetHeight);
        const r = { h: el.offsetHeight, rows };
        el.remove();
        return r;
      },
      done() { view.destroy(); host.remove(); },
    };
  }

  // ---------------- パッケージの部品 ----------------
  const THEME_XML = (pal, jaFont) => HEAD + `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Bansho"><a:themeElements>
<a:clrScheme name="Bansho"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F3A2E"/></a:dk2><a:lt2><a:srgbClr val="F3F3EC"/></a:lt2>
<a:accent1><a:srgbClr val="1D4ED8"/></a:accent1><a:accent2><a:srgbClr val="D62828"/></a:accent2><a:accent3><a:srgbClr val="15803D"/></a:accent3><a:accent4><a:srgbClr val="FFE100"/></a:accent4><a:accent5><a:srgbClr val="FF6FBF"/></a:accent5><a:accent6><a:srgbClr val="4CC3FF"/></a:accent6>
<a:hlink><a:srgbClr val="1D4ED8"/></a:hlink><a:folHlink><a:srgbClr val="7030A0"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Bansho"><a:majorFont><a:latin typeface="Segoe UI"/><a:ea typeface="${jaFont}"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface="${jaFont}"/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

  const lvlStyle = (tag) => `<p:${tag}><a:lvl1pPr><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:${tag}>`;
  const MASTER_XML = (pal) => HEAD + `<p:sldMaster ${NS}><p:cSld><p:bg><p:bgPr>${colorEl(pal.bgHex)}<a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>`
    + `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`
    + `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles>${lvlStyle('titleStyle')}${lvlStyle('bodyStyle')}${lvlStyle('otherStyle')}</p:txStyles></p:sldMaster>`;
  const LAYOUT_XML = HEAD + `<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="白紙"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  const REL = (items) => HEAD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`;
  const T = (n) => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + n;

  // ---------------- 本体 ----------------
  // doc → { blob, warnings }。theme: 書き出すテーマ（'white' など）
  B.exportPptx = async (doc, theme) => {
    const pal0 = makePalette(theme);
    const pal = Object.assign({}, pal0, { bgHex: toHex(pal0.bg).hex });
    const warnCount = new Map();
    const warn = (k) => warnCount.set(k, (warnCount.get(k) || 0) + 1);
    const measurer = makeMeasurer(theme);
    const files = [];
    const media = [];
    const jaFont = '游ゴシック';
    try {
      const slides = [];
      for (let si = 0; si < doc.pages.length; si++) {
        const page = doc.pages[si];
        const rels = [['rId1', T('slideLayout'), '../slideLayouts/slideLayout1.xml']];
        const ctx = {
          pal, warn, jaFont,
          measure: (o) => measurer.measure(o),
          measureEl: (o) => measurer.measureEl(o),
          color(c, kind) {
            if (!c) c = 'default';
            if (/^#[0-9a-f]{6}$/i.test(c)) return { hex: c.slice(1).toUpperCase(), alpha: 1 };
            return toHex((kind === 'hl' ? pal.hl : pal.c)[c] || pal.c.default);
          },
          highlight(h) { return over(toHex(pal.hlbg[h] || pal.hlbg.yellow), { hex: pal.bgHex }); },
          async addMedia(src, w, h) {
            const m = await toMedia(src, w, h).catch(() => null);
            if (!m) { warn('image'); return null; }
            const name = `image${media.length + 1}.${m.ext}`;
            media.push([`ppt/media/${name}`, m.bytes]);
            const rid = 'rId' + (rels.length + 1);
            rels.push([rid, T('image'), `../media/${name}`]);
            return rid;
          },
        };
        let id = 2;
        let tree = '';
        const all = (page.bg && page.bg.objs ? page.bg.objs : []).concat(page.objects);
        for (const o of all) {
          if (o.type === 'text') tree += textBox(o, id++, ctx);
          else if (o.type === 'stroke') tree += strokeShape(o, id++, ctx);
          else if (o.type === 'shape') tree += shapeObj(o, id++, ctx);
          else if (o.type === 'image') tree += await imageObj(o, id++, ctx);
          else if (o.type === 'table') tree += tableObj(o, id++, ctx);
        }
        const bgHex = page.bg && page.bg.color ? page.bg.color.slice(1).toUpperCase() : pal.bgHex;
        const slide = HEAD + `<p:sld ${NS}><p:cSld><p:bg><p:bgPr>${colorEl(bgHex)}<a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${tree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
        slides.push([slide, rels]);
      }

      const n = slides.length;
      const ct = HEAD + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>`
        + `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`
        + `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`
        + `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`
        + `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
        + `<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>`
        + `<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>`
        + `<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`
        + `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`
        + `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`
        + slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')
        + `</Types>`;
      files.push(['[Content_Types].xml', ct]);
      files.push(['_rels/.rels', REL([['rId1', T('officeDocument'), 'ppt/presentation.xml'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'], ['rId3', T('extended-properties'), 'docProps/app.xml']])]);
      const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
      files.push(['docProps/core.xml', HEAD + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>板書</dc:title><dc:creator>デジタル板書</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`]);
      files.push(['docProps/app.xml', HEAD + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>デジタル板書</Application><Slides>${n}</Slides></Properties>`]);
      files.push(['ppt/presentation.xml', HEAD + `<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 10}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${SW}" cy="${SH}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:lvl1pPr><a:defRPr lang="ja-JP"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>`]);
      files.push(['ppt/_rels/presentation.xml.rels', REL([['rId1', T('slideMaster'), 'slideMasters/slideMaster1.xml'], ['rId2', T('theme'), 'theme/theme1.xml'], ['rId3', T('presProps'), 'presProps.xml'], ['rId4', T('viewProps'), 'viewProps.xml'], ['rId5', T('tableStyles'), 'tableStyles.xml']].concat(slides.map((_, i) => [`rId${i + 10}`, T('slide'), `slides/slide${i + 1}.xml`])))]);
      files.push(['ppt/presProps.xml', HEAD + `<p:presentationPr ${NS}/>`]);
      files.push(['ppt/viewProps.xml', HEAD + `<p:viewPr ${NS}><p:normalViewPr><p:restoredLeft sz="15000"/><p:restoredTop sz="94000"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`]);
      files.push(['ppt/tableStyles.xml', HEAD + `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`]);
      files.push(['ppt/theme/theme1.xml', THEME_XML(pal, jaFont)]);
      files.push(['ppt/slideMasters/slideMaster1.xml', MASTER_XML(pal)]);
      files.push(['ppt/slideMasters/_rels/slideMaster1.xml.rels', REL([['rId1', T('slideLayout'), '../slideLayouts/slideLayout1.xml'], ['rId2', T('theme'), '../theme/theme1.xml']])]);
      files.push(['ppt/slideLayouts/slideLayout1.xml', LAYOUT_XML]);
      files.push(['ppt/slideLayouts/_rels/slideLayout1.xml.rels', REL([['rId1', T('slideMaster'), '../slideMasters/slideMaster1.xml']])]);
      slides.forEach(([x, rels], i) => {
        files.push([`ppt/slides/slide${i + 1}.xml`, x]);
        files.push([`ppt/slides/_rels/slide${i + 1}.xml.rels`, REL(rels)]);
      });
      files.push(...media);
    } finally {
      measurer.done();
    }
    const LABELS = { box: '四角囲み（PowerPoint にない装飾のため、囲みなしで書き出し）', image: '書き出せなかった画像' };
    return { blob: zip(files), warnings: [...warnCount.entries()].map(([k, v]) => `${LABELS[k] || k}：${v} か所`) };
  };
})(window.B);
