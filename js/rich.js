'use strict';
// テキストの書式処理。
// contenteditable の DOM を「1文字ごとの書式つき文字列」に分解し、書式を変えてから
// 決まった形の HTML に組み直す。こうすると、範囲の一部だけ装飾を外す・重なった装飾を
// 整理する、といった操作を確実に行え、文字を編集しても装飾は文字についていく。
//
// 書式キー（外側から内側の順に入れ子にする）
//   z: 文字の大きさ（ボックスの基準サイズに対する倍率。1 なら指定なし）
//   x: 四角囲み(true)  h: マーカー色  c: 文字色（意味上の色名、または PowerPoint から読み込んだ '#rrggbb'）
//   b: 太字  i: 斜体  k: 取り消し線  s: 'sub' | 'sup'
//   f: フォント名（PowerPoint から読み込んだ特殊なフォント。普段の入力では使わない）
//   u: 下線。true（普通の下線）または 'sng' / 'dbl'（二重線）に ':色' を付けた文字列（例 'dbl:red', 'sng:#ff0000'）
// 行の書式（各文字に持たせ、行の先頭の文字の値を使う。KEYS には含めない）
//   pl: 左の余白（em）  ti: 1 行目の字下げ（em、負ならぶら下げ）  mt: 段落の前の間隔（em）  ta: 'center' | 'right'
//   bu: 箇条書きの記号（文字列）  bw: 記号の幅（em）  bc: 記号の色（意味上の色名か '#rrggbb'）  bs: 記号の大きさ（倍率）
//   空の行の書式は、その行の前の改行文字に持たせる（行頭の文字がないため）
//   行の書式がある文字列は、各行を <div style="..."> にして出力する（ない場合は <br> で区切る）
(function (B) {
  const KEYS = ['z', 'f', 'x', 'h', 'c', 'b', 'i', 'u', 'k', 's'];
  const NONE = Object.freeze({});
  const BLOCK = new Set(['DIV', 'P', 'LI', 'UL', 'OL', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE']);

  // 文字が選ばれていない（カーソルだけの）ときは、ブラウザの「次に打つ文字の書式」を使う。
  // 文字色・マーカー・囲みは execCommand に目印の値を渡し、読み取り時に意味上の値へ戻す。
  const COLOR_SENT = { default: '#000001', red: '#000002', blue: '#000003', green: '#000004', gray: '#000005' };
  const HL_SENT = { none: 'rgb(0, 0, 9)', yellow: 'rgb(0, 0, 5)', pink: 'rgb(0, 0, 6)', blue: 'rgb(0, 0, 7)', green: 'rgb(0, 0, 8)' };
  const BOX_ON = 'bansho-box', BOX_OFF = 'bansho-nobox';
  const invert = (o) => Object.fromEntries(Object.entries(o).map(([k, v]) => [v, k]));
  const COLOR_BY_SENT = invert(COLOR_SENT);
  const HL_BY_SENT = invert(HL_SENT);

  // ブラウザが書式を style 属性で書き出した場合に備え、テーマの実際の色 → 意味上の色名の対応を作る
  let palette = null;
  function paletteMap() {
    if (palette) return palette;
    palette = {};
    const probe = document.createElement('div');
    probe.style.cssText = 'position:absolute;left:-99999px;top:0;width:10px;height:10px';
    document.body.appendChild(probe);
    for (const theme of ['white', 'board']) {
      probe.className = 'theme-' + theme;
      for (const name of ['default', 'red', 'blue', 'green', 'gray']) {
        const s = document.createElement('span');
        s.style.color = `var(--c-${name})`;
        probe.appendChild(s);
        palette[getComputedStyle(s).color] = name;
        s.remove();
      }
    }
    probe.remove();
    return palette;
  }

  function hexToRgb(h) {
    const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(h || '');
    return m ? `rgb(${parseInt(m[1], 16)}, ${parseInt(m[2], 16)}, ${parseInt(m[3], 16)})` : null;
  }

  function colorName(v) {
    if (!v) return undefined;
    v = v.trim().toLowerCase();
    if (COLOR_BY_SENT[v]) return COLOR_BY_SENT[v];
    const rgb = v[0] === '#' ? hexToRgb(v) : v;
    for (const [hex, name] of Object.entries(COLOR_SENT)) if (hexToRgb(hex) === rgb) return name;
    return paletteMap()[rgb] || rgbToHex(rgb);
  }

  function rgbToHex(rgb) {
    const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/.exec(rgb || '');
    if (!m || (m[4] !== undefined && Number(m[4]) === 0)) return undefined;
    return '#' + [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('');
  }

  const isHex = (v) => typeof v === 'string' && /^#[0-9a-f]{6}$/i.test(v);
  const SEM = /^(red|blue|green|gray)$/;
  B.isColorValue = (v) => isHex(v) || SEM.test(v || '');
  B.isBullet = (v) => typeof v === 'string' && v.length > 0 && v.length <= 6 && !/[<>&"]/.test(v);
  B.isFontName = (v) => typeof v === 'string' && /^[^'"<>;\\{}()]{1,80}$/.test(v);

  // 大きさの倍率は小数3桁に丸める（同じ大きさの文字が1つの span にまとまるように）
  const ratio = (v) => {
    const r = Math.round(v * 1000) / 1000;
    return r > 0 && r !== 1 ? r : null;
  };
  let basePx = 0; // parse 中のボックスの基準サイズ（px）。ブラウザが px で書き出した大きさを倍率に直すのに使う

  const set = (a, k, v) => (a[k] === v ? a : Object.assign({}, a, { [k]: v }));
  const cval = (name) => (name === 'default' ? null : name);

  // 要素が付け加える書式
  function attrsOf(el, a) {
    switch (el.tagName) {
      case 'B': case 'STRONG': a = set(a, 'b', true); break;
      case 'I': case 'EM': a = set(a, 'i', true); break;
      case 'U': case 'INS': {
        const dbl = el.classList.contains('ul-dbl') || (el.style && el.style.textDecorationStyle === 'double');
        let col = null;
        for (const cls of el.classList) { const m = /^ul-(red|blue|green|gray)$/.exec(cls); if (m) col = m[1]; }
        const sc = el.style && el.style.textDecorationColor;
        if (!col && sc) col = colorName(sc) || null;
        if (col === 'default') col = null;
        a = set(a, 'u', !dbl && !col ? true : `${dbl ? 'dbl' : 'sng'}${col ? ':' + col : ''}`);
        break;
      }
      case 'S': case 'STRIKE': case 'DEL': a = set(a, 'k', true); break;
      case 'SUB': a = set(a, 's', 'sub'); break;
      case 'SUP': a = set(a, 's', 'sup'); break;
      case 'FONT': {
        const c = colorName(el.getAttribute('color'));
        if (c) a = set(a, 'c', cval(c));
        const f = el.getAttribute('face');
        if (f === BOX_ON) a = set(a, 'x', true);
        else if (f === BOX_OFF) a = set(a, 'x', null);
        break;
      }
    }
    if (el.classList && el.classList.length) {
      for (const cls of el.classList) {
        let m;
        if ((m = /^c-([a-z]+)$/.exec(cls))) a = set(a, 'c', cval(m[1]));
        else if ((m = /^hl-([a-z]+)$/.exec(cls))) a = set(a, 'h', m[1]);
        else if (cls === 'deco-box') a = set(a, 'x', true);
      }
    }
    const st = el.style;
    if (st && st.length) {
      const fw = st.fontWeight;
      if (fw) a = set(a, 'b', fw === 'bold' || fw === 'bolder' || parseInt(fw, 10) >= 600 ? true : null);
      if (st.fontStyle) a = set(a, 'i', st.fontStyle === 'italic' || st.fontStyle === 'oblique' ? true : null);
      const td = st.textDecorationLine || st.textDecoration;
      if (td && td.includes('underline') && el.tagName !== 'U') a = set(a, 'u', true);
      if (td && td.includes('line-through')) a = set(a, 'k', true);
      if (st.verticalAlign === 'sub') a = set(a, 's', 'sub');
      else if (st.verticalAlign === 'super') a = set(a, 's', 'sup');
      else if (st.fontSize) {
        // 上付き・下付きの縮小（vertical-align と一緒に書かれる）は大きさとして扱わない
        let m;
        if ((m = /^([\d.]+)em$/.exec(st.fontSize))) a = set(a, 'z', ratio((a.z || 1) * parseFloat(m[1])));
        else if ((m = /^([\d.]+)px$/.exec(st.fontSize)) && basePx) a = set(a, 'z', ratio(parseFloat(m[1]) / basePx));
      }
      if (st.color) { const c = colorName(st.color); if (c) a = set(a, 'c', cval(c)); }
      if (st.backgroundColor) {
        const h = HL_BY_SENT[st.backgroundColor.replace(/\s+/g, ' ')];
        if (h) a = set(a, 'h', h === 'none' ? null : h);
      }
      if (st.fontFamily) {
        if (st.fontFamily.includes(BOX_OFF)) a = set(a, 'x', null);
        else if (st.fontFamily.includes(BOX_ON)) a = set(a, 'x', true);
        else if (el.tagName === 'SPAN') {
          const name = st.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
          if (name && !/^bansho|^(sans-serif|serif|inherit)$/i.test(name) && B.isFontName(name)) a = set(a, 'f', name);
        }
      }
      if (BLOCK.has(el.tagName)) {
        const em = (v) => { const m = /^(-?[\d.]+)em$/.exec(v || ''); return m ? Number(m[1]) : null; };
        if (st.paddingLeft) a = set(a, 'pl', em(st.paddingLeft));
        if (st.textIndent) a = set(a, 'ti', em(st.textIndent));
        if (st.marginTop) a = set(a, 'mt', em(st.marginTop));
        if (st.textAlign === 'center' || st.textAlign === 'right') a = set(a, 'ta', st.textAlign);
        else if (st.textAlign) a = set(a, 'ta', null);
        const bw = em(st.getPropertyValue('--bw').trim());
        if (bw != null) a = set(a, 'bw', bw);
        const bc = st.getPropertyValue('--bc').trim();
        if (bc) { const m = /^var\(--c-([a-z]+)\)$/.exec(bc); a = set(a, 'bc', m ? m[1] : isHex(bc) ? bc.toLowerCase() : null); }
        const bs = parseFloat(st.getPropertyValue('--bs'));
        if (bs > 0) a = set(a, 'bs', bs);
      }
      if (BLOCK.has(el.tagName) && el.hasAttribute('data-bu')) {
        a = set(a, 'bu', el.getAttribute('data-bu') || null);
      }
    }
    return a;
  }

  function hasContent(n) {
    if (n.nodeType === 3) return /[^​﻿]/.test(n.data) ? 'content' : 'none';
    if (n.nodeType !== 1) return 'none';
    if (BLOCK.has(n.tagName)) return 'block';
    if (n.tagName === 'BR') return 'content';
    for (const c of n.childNodes) {
      const k = hasContent(c);
      if (k !== 'none') return k;
    }
    return 'none';
  }

  // 行末に置かれただけの <br>（空行を保つための詰め物）は改行として数えない
  function isPlaceholderBr(br, root) {
    let n = br;
    for (;;) {
      if (n.nextSibling) {
        n = n.nextSibling;
        const k = hasContent(n);
        if (k === 'content') return false;
        if (k === 'block') return true;
      } else {
        n = n.parentNode;
        if (!n || n === root || BLOCK.has(n.tagName)) return true;
      }
    }
  }

  // root の内容を [{ch, a}] に分解する。points（DOM 上の位置）があれば文字位置に換算して返す。
  function parse(root, points) {
    let host = root.parentElement;
    while (host && !host.style.fontSize) host = host.parentElement;
    basePx = (host && parseFloat(host.style.fontSize)) || 0;
    const flat = [];
    points = points || [];
    const offs = points.map(() => null);
    let lineHasContent = false, blockEnded = false;
    const newline = (la) => { flat.push({ ch: '\n', a: la || NONE }); lineHasContent = false; };
    const mark = (node, idx) => {
      for (let k = 0; k < points.length; k++) {
        if (offs[k] === null && points[k].node === node && points[k].offset === idx) offs[k] = flat.length;
      }
    };
    const walk = (node, a) => {
      const kids = node.childNodes;
      for (let i = 0; i < kids.length; i++) {
        mark(node, i);
        const ch = kids[i];
        if (ch.nodeType === 3) {
          const t = ch.data;
          if (blockEnded && /[^​﻿]/.test(t)) { newline(); blockEnded = false; }
          for (let k = 0; k < t.length; k++) {
            mark(ch, k);
            const c = t[k];
            if (c === '\n') newline();
            else if (c === '​' || c === '﻿') continue;
            else { flat.push({ ch: c === ' ' ? ' ' : c, a }); lineHasContent = true; }
          }
          mark(ch, t.length);
        } else if (ch.nodeType === 1) {
          const tag = ch.tagName;
          if (tag === 'BR') {
            if (blockEnded) { newline(); blockEnded = false; }
            if (!isPlaceholderBr(ch, root)) newline();
          } else if (BLOCK.has(tag)) {
            const la = attrsOf(ch, a);
            if (blockEnded || lineHasContent) newline(la);
            blockEnded = false;
            walk(ch, la);
            blockEnded = true;
          } else {
            walk(ch, attrsOf(ch, a));
          }
        }
      }
      mark(node, kids.length);
    };
    walk(root, NONE);
    for (let k = 0; k < offs.length; k++) if (offs[k] === null) offs[k] = flat.length;
    return { flat, offs };
  }

  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  function wrap(key, v, inner) {
    switch (key) {
      case 'z': return `<span style="font-size:${v}em">${inner}</span>`;
      case 'f': return B.isFontName(v) ? `<span style="font-family:'${v}'">${inner}</span>` : inner;
      case 'k': return `<s>${inner}</s>`;
      case 'x': return `<span class="deco-box">${inner}</span>`;
      case 'h': return `<span class="hl-${v}">${inner}</span>`;
      case 'c': return isHex(v) ? `<span style="color:${v.toLowerCase()}">${inner}</span>` : `<span class="c-${v}">${inner}</span>`;
      case 'u': {
        if (v === true) return `<u>${inner}</u>`;
        const [style, col] = String(v).split(':');
        const cls = [style === 'dbl' ? 'ul-dbl' : '', col && !isHex(col) ? 'ul-' + col : ''].filter(Boolean).join(' ');
        return `<u${cls ? ` class="${cls}"` : ''}${isHex(col) ? ` style="text-decoration-color:${col.toLowerCase()}"` : ''}>${inner}</u>`;
      }
      case 's': return `<${v}>${inner}</${v}>`;
      default: return `<${key}>${inner}</${key}>`;
    }
  }

  function renderRuns(chars, level) {
    if (!chars.length) return '';
    if (level === KEYS.length) return esc(chars.map((c) => c.ch).join(''));
    const key = KEYS[level];
    let out = '';
    for (let i = 0; i < chars.length;) {
      const v = chars[i].a[key] || null;
      let j = i + 1;
      while (j < chars.length && (chars[j].a[key] || null) === v) j++;
      const inner = renderRuns(chars.slice(i, j), level + 1);
      out += v ? wrap(key, v, inner) : inner;
      i = j;
    }
    return out;
  }

  // 決まった形の HTML：行は <br> で区切り、装飾は KEYS の順に入れ子にする
  function render(flat) {
    const lines = [[]], nls = [null];
    for (const c of flat) {
      if (c.ch === '\n') { lines.push([]); nls.push(c); }
      else lines[lines.length - 1].push(c);
    }
    const lineAttrs = (i) => (lines[i].length ? lines[i][0].a : nls[i] ? nls[i].a : NONE);
    const lineStyle = (i) => {
      const a = lineAttrs(i);
      const css = [];
      if (a.pl) css.push(`padding-left:${a.pl}em`);
      if (a.ti) css.push(`text-indent:${a.ti}em`);
      if (a.mt) css.push(`margin-top:${a.mt}em`);
      if (a.ta) css.push(`text-align:${a.ta}`);
      if (a.bu && a.bw) css.push(`--bw:${a.bw}em`);
      if (a.bu && a.bc) css.push(`--bc:${isHex(a.bc) ? a.bc : `var(--c-${a.bc})`}`);
      if (a.bu && a.bs) css.push(`--bs:${a.bs}`);
      return css.join(';');
    };
    const bullet = (i) => { const b = lineAttrs(i).bu; return b && B.isBullet(b) ? ` data-bu="${esc(b)}"` : ''; };
    if (lines.some((l, i) => lineStyle(i) || bullet(i))) {
      return lines.map((l, i) => {
        const st = lineStyle(i);
        return `<div${st ? ` style="${st}"` : ''}${bullet(i)}>${l.length ? renderRuns(l, 0) : '<br>'}</div>`;
      }).join('');
    }
    let html = lines.map((l) => renderRuns(l, 0)).join('<br>');
    if (lines.length > 1 && !lines[lines.length - 1].length) html += '<br>';
    return html;
  }

  // 決まった形の HTML 上で、文字位置 off に当たる DOM 位置
  function pointAt(root, off) {
    const kidsArr = Array.from(root.childNodes);
    if (kidsArr.length && kidsArr.every((n) => n.nodeName === 'DIV')) {
      // 行ごとの <div>：行の境目を 1 文字（改行）として数える
      let count = 0;
      for (let i = 0; i < kidsArr.length; i++) {
        if (i > 0) count += 1;
        const texts = [];
        const collect = (n) => { for (const c of n.childNodes) { if (c.nodeType === 3) texts.push(c); else if (c.nodeType === 1) collect(c); } };
        collect(kidsArr[i]);
        const len = texts.reduce((a, t) => a + t.data.length, 0);
        if (off <= count + len) {
          let k = off - count;
          for (const t of texts) { if (k <= t.data.length) return { node: t, offset: k }; k -= t.data.length; }
          return { node: kidsArr[i], offset: 0 };
        }
        count += len;
      }
      const last = kidsArr[kidsArr.length - 1];
      return { node: last, offset: last.childNodes.length };
    }
    let count = 0, result = null;
    const walk = (node) => {
      for (const ch of Array.from(node.childNodes)) {
        if (result) return;
        if (ch.nodeType === 3) {
          const len = ch.data.length;
          if (off <= count + len) { result = { node: ch, offset: off - count }; return; }
          count += len;
        } else if (ch.nodeName === 'BR') {
          if (off === count) { result = { node, offset: Array.prototype.indexOf.call(node.childNodes, ch) }; return; }
          count += 1;
        } else if (ch.nodeType === 1) {
          walk(ch);
        }
      }
    };
    walk(root);
    return result || { node: root, offset: root.childNodes.length };
  }

  const inRange = (flat, s, e) => flat.slice(s, e).filter((c) => c.ch !== '\n');

  // [s, e) の書式 key を value にする。全体が既に value なら外す（トグル）。
  // force を true にするとトグルせずに value を設定する（文字色の「既定色」など）。
  function toggle(flat, s, e, key, value, force) {
    const chars = inRange(flat, s, e);
    if (!chars.length) return false;
    const all = chars.every((c) => (c.a[key] || null) === value);
    const nv = force || !all ? value : null;
    for (const c of chars) c.a = set(c.a, key, nv);
    return true;
  }

  // [s, e) に共通の書式（ばらばらの key は undefined）
  function common(flat, s, e) {
    let chars = inRange(flat, s, e);
    if (!chars.length) {
      const prev = s > 0 ? flat[s - 1] : null;
      chars = prev && prev.ch !== '\n' ? [prev] : [{ a: NONE }];
    }
    const out = {};
    for (const k of KEYS) {
      const v = chars[0].a[k] || null;
      out[k] = chars.every((c) => (c.a[k] || null) === v) ? v : undefined;
    }
    return out;
  }

  function parseHtml(html) {
    const div = document.createElement('div');
    div.innerHTML = html;
    return parse(div).flat;
  }

  // 行の一覧：{ start, end（文字の範囲）, nl（行の前の改行文字の位置。先頭行は -1） }
  function lineList(flat) {
    const out = [];
    let start = 0, nl = -1;
    for (let i = 0; i <= flat.length; i++) {
      if (i === flat.length || flat[i].ch === '\n') {
        out.push({ start, end: i, nl });
        nl = i;
        start = i + 1;
      }
    }
    return out;
  }

  const LINE_KEYS = ['pl', 'ti', 'mt', 'ta', 'bu', 'bw', 'bc', 'bs'];
  function lineAttrsOf(flat, ln) {
    if (ln.end > ln.start) return flat[ln.start].a;
    return ln.nl >= 0 ? flat[ln.nl].a : NONE;
  }
  // 行の書式を変える（行の全文字と、行の前の改行文字に同じ値を持たせる）
  function setLineAttrs(flat, ln, attrs) {
    const apply = (c) => {
      const a = Object.assign({}, c.a);
      for (const k of LINE_KEYS) if (k in attrs) { if (attrs[k] == null) delete a[k]; else a[k] = attrs[k]; }
      c.a = a;
    };
    for (let i = ln.start; i < ln.end; i++) apply(flat[i]);
    if (ln.nl >= 0) apply(flat[ln.nl]);
  }

  B.rich = {
    KEYS,
    ratio,
    lineList,
    lineAttrsOf,
    setLineAttrs,
    parse,
    render,
    pointAt,
    toggle,
    common,
    parseHtml,
    canonical: (root) => render(parse(root).flat),
    canonicalHtml: (html) => render(parseHtml(html)),
    isEmpty: (html) => !parseHtml(html).some((c) => c.ch.trim() !== ''),
    // カーソルだけのときに「次に打つ文字の書式」を設定する
    typingStyle(key, value, cur) {
      document.execCommand('styleWithCSS', false, false);
      switch (key) {
        case 'b': document.execCommand('bold'); break;
        case 'i': document.execCommand('italic'); break;
        case 'u': document.execCommand('underline'); break;
        case 'k': document.execCommand('strikeThrough'); break;
        case 's': document.execCommand(value === 'sub' ? 'subscript' : 'superscript'); break;
        case 'c': document.execCommand('foreColor', false, COLOR_SENT[value] || COLOR_SENT.default); break;
        case 'h': document.execCommand('hiliteColor', false, HL_SENT[cur.h === value ? 'none' : value]); break;
        case 'x': document.execCommand('fontName', false, cur.x ? BOX_OFF : BOX_ON); break;
      }
    },
  };
})(window.B);
