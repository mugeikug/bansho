'use strict';
// PowerPoint（.pptx）の読み込み。
// スライドの要素を板書のオブジェクトに変換する：
//   文字（テキストボックス・プレースホルダー・図形内の文字・表のセル）→ 編集できるテキスト
//   線・矢印、塗りなしの四角・楕円 → 板書の図形
//   インク（PowerPoint で手書きした線）→ 板書のペンの線（消しゴムで消せる）
//   それ以外の図形（塗りのある図形、ブロック矢印、フリーフォームなど）→ 見た目を保った SVG 画像
//   画像 → 画像（トリミングを反映）
//   グラフ・動画・EMF 画像など扱えないものは読み飛ばし、件数を報告する
// 色の扱いは2通り（opts.mode）：
//   'keep'  … 元の色のまま。スライドの背景色・マスターのデザインも背景として読み込む
//   'board' … 板書の意味上の色（既定色・赤・青・緑）に寄せる。背景・デザインは読まない（白地／黒板の切り替えで色が意味を保つ）
// 外部ライブラリは使わない（zip の展開はブラウザ標準の DecompressionStream）。
(function (B) {
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const EMU_PT = 12700;

  // ---------------- zip ----------------

  async function inflateRaw(u8) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([u8]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function unzip(buf) {
    const dv = new DataView(buf);
    const u8 = new Uint8Array(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('PowerPoint ファイル（zip）として読めません。');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const entries = new Map();
    const dec = new TextDecoder();
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const local = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nameLen));
      entries.set(name, { method, csize, local });
      p += 46 + nameLen + extraLen + commentLen;
    }
    return {
      has: (name) => entries.has(name),
      async bytes(name) {
        const e = entries.get(name);
        if (!e) return null;
        const ln = dv.getUint16(e.local + 26, true), le = dv.getUint16(e.local + 28, true);
        const start = e.local + 30 + ln + le;
        const data = u8.subarray(start, start + e.csize);
        if (e.method === 0) return data;
        if (e.method === 8) return inflateRaw(data);
        throw new Error('未対応の圧縮形式です。');
      },
    };
  }

  // ---------------- パッケージ（XML・リレーションシップ） ----------------

  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg && seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  }

  class Pkg {
    constructor(zip) {
      this.zip = zip;
      this.cache = new Map();
    }

    async xml(path) {
      if (this.cache.has(path)) return this.cache.get(path);
      const b = await this.zip.bytes(path);
      const doc = b ? new DOMParser().parseFromString(new TextDecoder().decode(b), 'application/xml') : null;
      const root = doc && !doc.getElementsByTagName('parsererror').length ? doc.documentElement : null;
      this.cache.set(path, root);
      return root;
    }

    async rels(path) {
      const i = path.lastIndexOf('/');
      const rp = path.slice(0, i) + '/_rels/' + path.slice(i + 1) + '.rels';
      const root = await this.xml(rp);
      const map = new Map();
      if (!root) return map;
      for (const r of root.children) {
        const type = (r.getAttribute('Type') || '').split('/').pop();
        const ext = r.getAttribute('TargetMode') === 'External';
        map.set(r.getAttribute('Id'), { type, target: ext ? r.getAttribute('Target') : resolvePath(path, r.getAttribute('Target')), external: ext });
      }
      return map;
    }
  }

  // ---------------- XML の小道具 ----------------

  const kids = (el, name) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
  const kid = (el, name) => (el ? Array.from(el.children).find((c) => c.localName === name) || null : null);
  const path = (el, ...names) => names.reduce((e, n) => kid(e, n), el);
  const num = (v, d) => (v == null || v === '' ? d : Number(v));
  const rAttr = (el, name) => (el ? el.getAttributeNS(NS_R, name) || el.getAttribute('r:' + name) : null);
  const descendant = (el, name) => (el ? el.getElementsByTagNameNS('*', name)[0] || null : null);

  // ---------------- 色 ----------------

  const PRESET_COLORS = {
    black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00',
    cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080',
    darkBlue: '00008B', darkRed: '8B0000', darkGreen: '006400', lightGray: 'D3D3D3', navy: '000080',
  };

  function hexToRgb(h) { const n = parseInt(h, 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; }
  function rgbToHex(r, g, b) { return [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase(); }

  function rgbToHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (mx + mn) / 2;
    if (mx !== mn) {
      const d = mx - mn;
      s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
      if (mx === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (mx === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h *= 60;
    }
    return [h, s, l];
  }

  function hslToRgb(h, s, l) {
    const f = (n) => {
      const k = (n + h / 30) % 12;
      const a = s * Math.min(l, 1 - l);
      return 255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1)));
    };
    return [f(0), f(8), f(4)];
  }

  // 色要素（srgbClr / schemeClr など）を含む要素から {hex, alpha} を求める
  function readColor(el, ctx) {
    if (!el) return null;
    const c = Array.from(el.children).find((x) => /^(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(x.localName));
    if (!c) return null;
    let hex = '000000';
    switch (c.localName) {
      case 'srgbClr': hex = c.getAttribute('val') || '000000'; break;
      case 'sysClr': hex = c.getAttribute('lastClr') || (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000'); break;
      case 'prstClr': hex = PRESET_COLORS[c.getAttribute('val')] || '000000'; break;
      case 'scrgbClr': hex = rgbToHex(num(c.getAttribute('r'), 0) / 100000 * 255, num(c.getAttribute('g'), 0) / 100000 * 255, num(c.getAttribute('b'), 0) / 100000 * 255); break;
      case 'hslClr': hex = rgbToHex(...hslToRgb(num(c.getAttribute('hue'), 0) / 60000, num(c.getAttribute('sat'), 0) / 100000, num(c.getAttribute('lum'), 0) / 100000)); break;
      case 'schemeClr': {
        let v = c.getAttribute('val');
        if (v === 'phClr') { if (ctx.phClr) return applyMods(ctx.phClr, c); v = 'tx1'; }
        if (ctx.clrMap[v]) v = ctx.clrMap[v];
        hex = ctx.theme.colors[v] || '000000';
        break;
      }
    }
    return applyMods({ hex, alpha: 1 }, c);
  }

  function applyMods(base, c) {
    let [r, g, b] = hexToRgb(base.hex);
    let alpha = base.alpha;
    for (const m of c.children) {
      const v = num(m.getAttribute('val'), 100000) / 100000;
      switch (m.localName) {
        case 'alpha': alpha = v; break;
        case 'lumMod': case 'lumOff': case 'satMod': {
          let [h, s, l] = rgbToHsl(r, g, b);
          if (m.localName === 'lumMod') l *= v;
          else if (m.localName === 'lumOff') l += v;
          else s *= v;
          [r, g, b] = hslToRgb(h, Math.max(0, Math.min(1, s)), Math.max(0, Math.min(1, l)));
          break;
        }
        case 'tint': r += (255 - r) * (1 - v); g += (255 - g) * (1 - v); b += (255 - b) * (1 - v); break;
        case 'shade': r *= v; g *= v; b *= v; break;
      }
    }
    return { hex: rgbToHex(r, g, b), alpha };
  }

  // 実際の色 → 板書の意味上の色（既定色・赤・青・緑）
  function semanticColor(col) {
    if (!col) return 'default';
    const [h, s, l] = rgbToHsl(...hexToRgb(col.hex));
    if (s < 0.25 && l >= 0.45 && l <= 0.85) return 'gray';
    if (s < 0.25 || l < 0.15 || l > 0.85) return 'default';
    if (h >= 75 && h < 165) return 'green';
    if (h >= 165 && h < 280) return 'blue';
    return 'red'; // 赤・橙・茶・黄・紫・桃
  }

  function markerColor(col) {
    const [h, s] = rgbToHsl(...hexToRgb(col.hex));
    if (s < 0.2) return 'yellow';
    if (h >= 40 && h < 75) return 'yellow';
    if (h >= 75 && h < 165) return 'green';
    if (h >= 165 && h < 260) return 'blue';
    if (h >= 260 || h < 20) return 'pink';
    return 'yellow';
  }

  const cssColor = (col) => `#${col.hex}`;

  // ---------------- テーマ ----------------

  async function parseTheme(pkg, themePath) {
    const t = { colors: {}, lnWidths: [9525, 25400, 38100] };
    const root = themePath ? await pkg.xml(themePath) : null;
    const scheme = descendant(root, 'clrScheme');
    if (scheme) {
      for (const c of scheme.children) {
        const x = c.firstElementChild;
        if (!x) continue;
        t.colors[c.localName] = x.localName === 'sysClr' ? (x.getAttribute('lastClr') || '000000') : (x.getAttribute('val') || '000000');
      }
    }
    const lnLst = descendant(root, 'lnStyleLst');
    if (lnLst) t.lnWidths = kids(lnLst, 'ln').map((l) => num(l.getAttribute('w'), 9525));
    const fontOf = (n) => { const e = path(descendant(root, n), 'latin'); return e ? e.getAttribute('typeface') : ''; };
    t.majorLatin = fontOf('majorFont');
    t.minorLatin = fontOf('minorFont');
    return t;
  }

  // ---------------- 図形の形（SVG パス） ----------------

  const P = (pts) => 'M' + pts.map((p) => `${r2(p[0])} ${r2(p[1])}`).join('L') + 'Z';
  const r2 = (n) => Math.round(n * 100) / 100;

  function adj(av, name, d) {
    const g = kids(av, 'gd').find((x) => x.getAttribute('name') === name);
    if (!g) return d;
    const m = /val\s+(-?\d+)/.exec(g.getAttribute('fmla') || '');
    return m ? Number(m[1]) : d;
  }

  function ellipseD(w, h) {
    return `M0 ${r2(h / 2)}A${r2(w / 2)} ${r2(h / 2)} 0 1 0 ${r2(w)} ${r2(h / 2)}A${r2(w / 2)} ${r2(h / 2)} 0 1 0 0 ${r2(h / 2)}Z`;
  }

  // 右向き矢印の点列を、向きに合わせて変換する
  function arrowPts(w, h, a1, a2, dir) {
    const horiz = dir === 'right' || dir === 'left';
    const L = horiz ? w : h, T = horiz ? h : w;
    const ss = Math.min(w, h);
    const y1 = T / 2 - (T * a1) / 200000;
    const hx = L - (ss * a2) / 100000;
    let pts = [[0, y1], [hx, y1], [hx, 0], [L, T / 2], [hx, T], [hx, T - y1], [0, T - y1]];
    if (dir === 'left') pts = pts.map(([x, y]) => [L - x, y]);
    if (dir === 'down') pts = pts.map(([x, y]) => [y, x]);
    if (dir === 'up') pts = pts.map(([x, y]) => [y, L - x]);
    return pts;
  }

  // プリセット形状 → [{d, fill(bool)}]。未対応は null
  function presetPaths(prst, w, h, av) {
    const ss = Math.min(w, h);
    const A = (name, d) => adj(av, name, d);
    switch (prst) {
      case 'rect': case 'flowChartProcess': case 'textBox':
        return [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]) }];
      case 'roundRect': case 'flowChartAlternateProcess': {
        const r = Math.min(ss * A('adj', 16667) / 100000, ss / 2);
        return [{ d: `M${r2(r)} 0H${r2(w - r)}A${r2(r)} ${r2(r)} 0 0 1 ${r2(w)} ${r2(r)}V${r2(h - r)}A${r2(r)} ${r2(r)} 0 0 1 ${r2(w - r)} ${r2(h)}H${r2(r)}A${r2(r)} ${r2(r)} 0 0 1 0 ${r2(h - r)}V${r2(r)}A${r2(r)} ${r2(r)} 0 0 1 ${r2(r)} 0Z` }];
      }
      case 'flowChartTerminator': {
        const r = h / 2;
        return [{ d: `M${r2(r)} 0H${r2(w - r)}A${r2(r)} ${r2(r)} 0 0 1 ${r2(w - r)} ${r2(h)}H${r2(r)}A${r2(r)} ${r2(r)} 0 0 1 ${r2(r)} 0Z` }];
      }
      case 'ellipse': case 'flowChartConnector': case 'donut':
        return [{ d: ellipseD(w, h) }];
      case 'triangle': { const x = w * A('adj', 50000) / 100000; return [{ d: P([[x, 0], [w, h], [0, h]]) }]; }
      case 'rtTriangle': return [{ d: P([[0, 0], [w, h], [0, h]]) }];
      case 'diamond': case 'flowChartDecision': return [{ d: P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]) }];
      case 'parallelogram': case 'flowChartInputOutput': { const o = Math.min(w, ss * A('adj', 25000) / 100000); return [{ d: P([[o, 0], [w, 0], [w - o, h], [0, h]]) }]; }
      case 'trapezoid': { const o = Math.min(w / 2, ss * A('adj', 25000) / 100000); return [{ d: P([[o, 0], [w - o, 0], [w, h], [0, h]]) }]; }
      case 'pentagon': return [{ d: P([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]) }];
      case 'hexagon': { const o = Math.min(w / 2, ss * A('adj', 25000) / 100000); return [{ d: P([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]) }]; }
      case 'octagon': { const o = Math.min(w / 2, h / 2, ss * A('adj', 29289) / 100000); return [{ d: P([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]) }]; }
      case 'star5': {
        const pts = [];
        for (let i = 0; i < 10; i++) {
          const ang = -Math.PI / 2 + (i * Math.PI) / 5;
          const rr = i % 2 ? 0.382 : 1;
          pts.push([w / 2 + (w / 2) * rr * Math.cos(ang), h * 0.53 + (h * 0.53) * rr * Math.sin(ang)]);
        }
        return [{ d: P(pts) }];
      }
      case 'rightArrow': case 'leftArrow': case 'upArrow': case 'downArrow':
        return [{ d: P(arrowPts(w, h, A('adj1', 50000), A('adj2', 50000), prst.replace('Arrow', ''))) }];
      case 'leftRightArrow': {
        const y1 = h / 2 - (h * A('adj1', 50000)) / 200000, hx = (ss * A('adj2', 50000)) / 100000;
        return [{ d: P([[0, h / 2], [hx, 0], [hx, y1], [w - hx, y1], [w - hx, 0], [w, h / 2], [w - hx, h], [w - hx, h - y1], [hx, h - y1], [hx, h]]) }];
      }
      case 'chevron': { const o = Math.min(w, ss * A('adj', 50000) / 100000); return [{ d: P([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]]) }]; }
      case 'homePlate': { const o = Math.min(w, ss * A('adj', 50000) / 100000); return [{ d: P([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]]) }]; }
      case 'plus': case 'mathPlus': {
        const o = ss * A('adj', 25000) / 100000;
        return [{ d: P([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h], [o, h], [o, h - o], [0, h - o], [0, o], [o, o]].map(([x, y]) => [x, y])) }];
      }
      case 'frame': {
        const o = ss * A('adj1', 12500) / 100000;
        return [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]) + P([[o, o], [o, h - o], [w - o, h - o], [w - o, o]]) }];
      }
      case 'wedgeRectCallout': case 'wedgeRoundRectCallout': {
        const tx = w / 2 + (w * A('adj1', -20833)) / 100000, ty = h / 2 + (h * A('adj2', 62500)) / 100000;
        const b0 = Math.max(0, Math.min(w * 0.84, tx - w * 0.08)), b1 = b0 + w * 0.16;
        if (ty > h) return [{ d: P([[0, 0], [w, 0], [w, h], [b1, h], [tx, ty], [b0, h], [0, h]]) }];
        if (ty < 0) return [{ d: P([[0, 0], [b0, 0], [tx, ty], [b1, 0], [w, 0], [w, h], [0, h]]) }];
        return [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]) }];
      }
      case 'leftBracket': case 'rightBracket': case 'bracketPair': {
        const r = Math.min(h / 2, ss * A('adj', 8333) / 100000);
        const left = (x0) => `M${r2(x0 + Math.min(w, r * 1.2))} 0Q${r2(x0)} 0 ${r2(x0)} ${r2(r)}V${r2(h - r)}Q${r2(x0)} ${r2(h)} ${r2(x0 + Math.min(w, r * 1.2))} ${r2(h)}`;
        const right = (x1) => `M${r2(x1 - Math.min(w, r * 1.2))} 0Q${r2(x1)} 0 ${r2(x1)} ${r2(r)}V${r2(h - r)}Q${r2(x1)} ${r2(h)} ${r2(x1 - Math.min(w, r * 1.2))} ${r2(h)}`;
        if (prst === 'leftBracket') return [{ d: `M${r2(w)} 0Q0 0 0 ${r2(r)}V${r2(h - r)}Q0 ${r2(h)} ${r2(w)} ${r2(h)}`, fill: false }];
        if (prst === 'rightBracket') return [{ d: `M0 0Q${r2(w)} 0 ${r2(w)} ${r2(r)}V${r2(h - r)}Q${r2(w)} ${r2(h)} 0 ${r2(h)}`, fill: false }];
        return [{ d: left(0) + right(w), fill: false }];
      }
      case 'leftBrace': case 'rightBrace': case 'bracePair': {
        const brace = (x0, x1) => { // x0: 外側（つる）, x1: 中央の突起側
          const m = (x0 + x1) / 2;
          return `M${r2(x0)} 0Q${r2(m)} 0 ${r2(m)} ${r2(h * 0.12)}V${r2(h * 0.38)}Q${r2(m)} ${r2(h / 2)} ${r2(x1)} ${r2(h / 2)}Q${r2(m)} ${r2(h / 2)} ${r2(m)} ${r2(h * 0.62)}V${r2(h * 0.88)}Q${r2(m)} ${r2(h)} ${r2(x0)} ${r2(h)}`;
        };
        if (prst === 'leftBrace') return [{ d: brace(w, 0), fill: false }];
        if (prst === 'rightBrace') return [{ d: brace(0, w), fill: false }];
        const bw = Math.min(w / 2, h * 0.1);
        return [{ d: brace(bw, 0) + brace(w - bw, w), fill: false }];
      }
      case 'line': case 'straightConnector1':
        return [{ d: `M0 0L${r2(w)} ${r2(h)}`, fill: false }];
      case 'bentConnector2':
        return [{ d: `M0 0H${r2(w)}V${r2(h)}`, fill: false }];
      case 'bentConnector3': {
        const x = (w * A('adj1', 50000)) / 100000;
        return [{ d: `M0 0H${r2(x)}V${r2(h)}H${r2(w)}`, fill: false }];
      }
      case 'curvedConnector3':
        return [{ d: `M0 0C${r2(w / 2)} 0 ${r2(w / 2)} ${r2(h)} ${r2(w)} ${r2(h)}`, fill: false }];
      case 'arc': {
        const a0 = A('adj1', 16200000) / 60000 * Math.PI / 180, a1 = A('adj2', 0) / 60000 * Math.PI / 180;
        const rx = w / 2, ry = h / 2;
        const sx = rx + rx * Math.cos(a0), sy = ry + ry * Math.sin(a0);
        const ex = rx + rx * Math.cos(a1), ey = ry + ry * Math.sin(a1);
        let sw = a1 - a0;
        if (sw < 0) sw += Math.PI * 2;
        return [{ d: `M${r2(sx)} ${r2(sy)}A${r2(rx)} ${r2(ry)} 0 ${sw > Math.PI ? 1 : 0} 1 ${r2(ex)} ${r2(ey)}`, fill: false }];
      }
    }
    return null;
  }

  // フリーフォーム（custGeom）→ [{d, fill, stroke}]
  function customPaths(cg, w, h) {
    const out = [];
    for (const p of kids(kid(cg, 'pathLst'), 'path')) {
      const pw = num(p.getAttribute('w'), 0) || w, ph = num(p.getAttribute('h'), 0) || h;
      const sx = w / pw, sy = h / ph;
      const pt = (el) => [num(el.getAttribute('x'), 0) * sx, num(el.getAttribute('y'), 0) * sy];
      let d = '';
      let cur = [0, 0];
      for (const c of p.children) {
        const pts = kids(c, 'pt').map(pt);
        switch (c.localName) {
          case 'moveTo': cur = pts[0]; d += `M${r2(cur[0])} ${r2(cur[1])}`; break;
          case 'lnTo': cur = pts[0]; d += `L${r2(cur[0])} ${r2(cur[1])}`; break;
          case 'cubicBezTo': d += 'C' + pts.map((q) => `${r2(q[0])} ${r2(q[1])}`).join(' '); cur = pts[2]; break;
          case 'quadBezTo': d += 'Q' + pts.map((q) => `${r2(q[0])} ${r2(q[1])}`).join(' '); cur = pts[1]; break;
          case 'arcTo': {
            const wR = num(c.getAttribute('wR'), 0) * sx, hR = num(c.getAttribute('hR'), 0) * sy;
            const st = num(c.getAttribute('stAng'), 0) / 60000 * Math.PI / 180;
            const sw = num(c.getAttribute('swAng'), 0) / 60000 * Math.PI / 180;
            const cx = cur[0] - wR * Math.cos(st), cy = cur[1] - hR * Math.sin(st);
            const ex = cx + wR * Math.cos(st + sw), ey = cy + hR * Math.sin(st + sw);
            d += `A${r2(wR)} ${r2(hR)} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${r2(ex)} ${r2(ey)}`;
            cur = [ex, ey];
            break;
          }
          case 'close': d += 'Z'; break;
        }
      }
      if (d) out.push({ d, fill: p.getAttribute('fill') !== 'none', stroke: p.getAttribute('stroke') !== '0' && p.getAttribute('stroke') !== 'false' });
    }
    return out;
  }

  // ---------------- 塗り・線 ----------------

  function readFill(spPr, style, ctx) {
    if (spPr) {
      for (const c of spPr.children) {
        switch (c.localName) {
          case 'noFill': return null;
          case 'solidFill': return readColor(c, ctx);
          case 'gradFill': {
            const gs = path(c, 'gsLst', 'gs');
            return gs ? readColor(gs, ctx) : null;
          }
          case 'pattFill': return readColor(kid(c, 'fgClr'), ctx);
          case 'blipFill': return { blip: c };
          case 'grpFill': return ctx.groupFill || null;
        }
      }
    }
    const ref = kid(style, 'fillRef');
    if (ref && num(ref.getAttribute('idx'), 0) > 0) return readColor(ref, ctx);
    return null;
  }

  function readLine(spPr, style, ctx) {
    const ln = kid(spPr, 'ln');
    const ref = kid(style, 'lnRef');
    const refIdx = ref ? num(ref.getAttribute('idx'), 0) : 0;
    let col = null;
    if (ln) {
      if (kid(ln, 'noFill')) return null;
      const sf = kid(ln, 'solidFill') || path(ln, 'gradFill', 'gsLst', 'gs');
      if (sf) col = readColor(sf, ctx);
    }
    if (!col) {
      if (refIdx > 0) col = readColor(ref, ctx);
      else return null;
    }
    let w = ln && ln.getAttribute('w') != null ? num(ln.getAttribute('w'), 9525) : (refIdx > 0 ? ctx.theme.lnWidths[refIdx - 1] || 9525 : 9525);
    const dash = path(ln, 'prstDash') ? path(ln, 'prstDash').getAttribute('val') : 'solid';
    const end = (n) => { const e = kid(ln, n); const t = e && e.getAttribute('type'); return t && t !== 'none' ? t : null; };
    return { color: col, w: Math.max(1, w * ctx.s), dash, head: end('headEnd'), tail: end('tailEnd') };
  }

  function dashArray(dash, w) {
    const k = Math.max(1, w);
    switch (dash) {
      case 'dash': case 'sysDash': return `${4 * k} ${3 * k}`;
      case 'dot': case 'sysDot': return `${k} ${k * 1.5}`;
      case 'lgDash': return `${8 * k} ${3 * k}`;
      case 'dashDot': case 'sysDashDot': return `${4 * k} ${3 * k} ${k} ${3 * k}`;
      case 'lgDashDot': return `${8 * k} ${3 * k} ${k} ${3 * k}`;
      case 'lgDashDotDot': case 'sysDashDotDot': return `${8 * k} ${3 * k} ${k} ${3 * k} ${k} ${3 * k}`;
    }
    return '';
  }

  // 線・コネクタの両端の向き（矢じり用）。最後の区間の向きに合わせる。
  function connectorEnds(prst, w, h) {
    switch (prst) {
      case 'line': case 'straightConnector1': return { head: [w, h, 0, 0], tail: [0, 0, w, h] };
      case 'bentConnector2': return { head: [w, 0, 0, 0], tail: [w, 0, w, h] };
      case 'bentConnector3': case 'curvedConnector3': return { head: [1, 0, 0, 0], tail: [w - 1, h, w, h] };
    }
    return null;
  }

  // ---------------- SVG 画像にする ----------------

  function svgDataUrl(svg) {
    return 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)));
  }

  // 線の端の矢じり（三角形）
  function arrowHead(x, y, ux, uy, w, color) {
    const L = Math.max(8, w * 4), hw = L * 0.5;
    const bx = x - ux * L, by = y - uy * L;
    return `<path d="M${r2(x)} ${r2(y)}L${r2(bx - uy * hw)} ${r2(by + ux * hw)}L${r2(bx + uy * hw)} ${r2(by - ux * hw)}Z" fill="${color}"/>`;
  }

  // box: {x, y, w, h}（px）、paths: ローカル座標のパス
  function shapeImage(box, paths, fill, line, rot, flipH, flipV, ends) {
    const pad = (line ? line.w / 2 + (line.head || line.tail ? line.w * 4 : 0) : 0) + 2;
    const rad = (rot * Math.PI) / 180;
    const cos = Math.abs(Math.cos(rad)), sin = Math.abs(Math.sin(rad));
    const W = box.w * cos + box.h * sin + pad * 2, H = box.w * sin + box.h * cos + pad * 2;
    let body = '';
    const st = line ? `stroke="${cssColor(line.color)}" stroke-opacity="${line.color.alpha}" stroke-width="${r2(line.w)}" stroke-linejoin="round" stroke-linecap="${line.dash === 'solid' ? 'round' : 'butt'}"${line.dash !== 'solid' ? ` stroke-dasharray="${dashArray(line.dash, line.w)}"` : ''}` : 'stroke="none"';
    for (const p of paths) {
      const f = fill && !fill.blip && p.fill !== false ? `fill="${cssColor(fill)}" fill-opacity="${fill.alpha}"` : 'fill="none"';
      body += `<path d="${p.d}" ${f} ${p.stroke === false ? 'stroke="none"' : st}/>`;
    }
    // ends: { head: [fromX, fromY, toX, toY], tail: [...] }（矢じりは to の位置に from から向けて描く）
    if (line && ends) {
      const col = cssColor(line.color);
      for (const k of ['head', 'tail']) {
        if (!line[k] || !ends[k]) continue;
        const [fx, fy, tx, ty] = ends[k];
        const L = Math.hypot(tx - fx, ty - fy) || 1;
        body += arrowHead(tx, ty, (tx - fx) / L, (ty - fy) / L, line.w, col);
      }
    }
    const tf = `translate(${r2(W / 2)} ${r2(H / 2)}) rotate(${rot}) scale(${flipH ? -1 : 1} ${flipV ? -1 : 1}) translate(${r2(-box.w / 2)} ${r2(-box.h / 2)})`;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${r2(W)}" height="${r2(H)}" viewBox="0 0 ${r2(W)} ${r2(H)}"><g transform="${tf}">${body}</g></svg>`;
    const cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    return { id: B.uid('i'), type: 'image', x: r2(cx - W / 2), y: r2(cy - H / 2), w: r2(W), h: r2(H), src: svgDataUrl(svg) };
  }

  // ---------------- 文字 ----------------

  // 段落・文字の書式は、図形自身 → レイアウト → マスターのプレースホルダー → マスターの既定スタイル の順に探す
  function styleChain(ctx, own, ph) {
    const chain = [own || null];
    if (ph) {
      if (ph.layout) chain.push(path(ph.layout, 'txBody', 'lstStyle'));
      if (ph.master) chain.push(path(ph.master, 'txBody', 'lstStyle'));
      const t = ph.type;
      const tx = ctx.txStyles;
      chain.push(t === 'title' || t === 'ctrTitle' ? tx.title : (t === 'body' || t === 'subTitle' || t === 'obj' || !t) ? tx.body : tx.other);
    } else {
      chain.push(ctx.txStyles.other);
    }
    chain.push(ctx.defaultTextStyle);
    // 先頭（図形自身の lstStyle）は無くても位置を保つ（findR が「先頭の次」に図形スタイルの文字色を挟むため）
    return [chain[0]].concat(chain.slice(1).filter(Boolean));
  }

  function lvlPPr(lst, lvl) { return kid(lst, `lvl${lvl + 1}pPr`); }

  // pPr / defRPr の属性・子要素を、段落自身 → 継承チェーンの順に探す
  function findP(pPr, chain, lvl, fn) {
    if (pPr) { const v = fn(pPr); if (v !== undefined) return v; }
    for (const lst of chain) {
      const p = lvlPPr(lst, lvl);
      if (p) { const v = fn(p); if (v !== undefined) return v; }
    }
    return undefined;
  }

  function findR(rPr, pPr, chain, lvl, fn, fontRefColor) {
    if (rPr) { const v = fn(rPr); if (v !== undefined) return v; }
    if (pPr) { const d = kid(pPr, 'defRPr'); if (d) { const v = fn(d); if (v !== undefined) return v; } }
    // 図形自身の lstStyle だけ先に見て、その次に図形スタイルの文字色（fontRef）
    for (let i = 0; i < chain.length; i++) {
      if (i === 1 && fontRefColor !== undefined) { const v = fn(null, fontRefColor); if (v !== undefined) return v; }
      const p = lvlPPr(chain[i], lvl);
      const d = kid(p, 'defRPr');
      if (d) { const v = fn(d); if (v !== undefined) return v; }
    }
    return undefined;
  }

  const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  function roman(n) { let s = ''; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; }
  function autoNum(type, n) {
    const alpha = (k) => String.fromCharCode(97 + ((k - 1) % 26));
    const m = /^(arabic|alphaLc|alphaUc|romanLc|romanUc|circleNum)/.exec(type || '') || [, 'arabic'];
    let core = String(n);
    if (m[1] === 'alphaLc') core = alpha(n);
    else if (m[1] === 'alphaUc') core = alpha(n).toUpperCase();
    else if (m[1] === 'romanLc') core = roman(n);
    else if (m[1] === 'romanUc') core = roman(n).toUpperCase();
    else if (m[1] === 'circleNum') return n <= 20 ? String.fromCharCode(0x2460 + n - 1) : `(${n})`;
    if (/ParenBoth$/.test(type)) return `(${core})`;
    if (/ParenR$/.test(type)) return `${core})`;
    if (/Plain$/.test(type)) return core;
    return core + '.';
  }

  // Wingdings / Symbol フォントの記号（箇条書きに使われる）→ Unicode
  const WINGDINGS = { l: '●', m: '○', n: '▪', o: '□', p: '□', q: '❑', r: '❒', s: '◆', t: '◆', u: '◆', v: '❖', w: '◆', x: '⌧', '§': '▪', 'Ø': '➢', 'ü': '✓', 'û': '✗', 'à': '→', 'è': '➔', F: '☞', J: '☺', '¨': '□' };
  const SYMBOL = { '·': '•', 'Þ': '⇒', '®': '→', '¾': '—', '-': '−' };
  function bulletChar(ch, font) {
    if (!ch) return '•';
    let c = ch;
    const code = c.charCodeAt(0);
    if (code >= 0xf000 && code <= 0xf0ff) c = String.fromCharCode(code - 0xf000); // 記号フォントの私用領域
    if (/wingdings/i.test(font || '') || c !== ch) return WINGDINGS[c] || '■';
    if (/^symbol$/i.test(font || '')) return SYMBOL[c] || '•';
    return ch;
  }

  // よくあるフォント（板書の既定のフォントで代わりに表示してよいもの）
  const COMMON_FONT = /^(yu ?gothic|游ゴシック|游明朝|yu ?mincho|meiryo|メイリオ|ms ?p?gothic|ｍｓ ?ｐ?ゴシック|ms ?p?mincho|ｍｓ ?ｐ?明朝|calibri|calibri light|arial|segoe ui|aptos|helvetica|times new roman|century|century gothic|cambria|georgia|verdana|tahoma|consolas|courier new|palatino linotype|book antiqua|garamond|hiragino|bizud|biz ud|noto sans|noto serif|游ゴシック light|游ゴシック medium|yu gothic light|yu gothic medium|yu gothic ui)/i;
  function isCommonFont(name) { return COMMON_FONT.test((name || '').trim()) || /^(symbol|wingdings)/i.test(name || ''); }

  // 欧文にも和文フォントを使っているか（游ゴシックなど）
  const JA_FONT = /[^\x00-\x7f]|yu ?gothic|yugoth|yu ?mincho|meiryo|ms ?p?gothic|ms ?p?mincho|hiragino|noto (sans|serif) (cjk|jp)|biz ?ud/i;
  function isJaFont(name, ctx) {
    if (name === '+mn-lt' || name === '+mn-ea' || !name) name = ctx.theme.minorLatin;
    else if (name === '+mj-lt' || name === '+mj-ea') name = ctx.theme.majorLatin;
    return JA_FONT.test(name || '');
  }

  // 文字列の幅（px）。箇条書きの記号のあとの空白の数を決めるのに使う
  let measureCtx = null;
  function textWidth(str, px, ja) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    measureCtx.font = `${px}px ${ja ? '"BanshoJa"' : '"BanshoLatin", "BanshoJa"'}, "Yu Gothic", sans-serif`;
    return measureCtx.measureText(str).width;
  }

  // txBody → { flat, sizes（文字ごとの板書の pt）, jaFlags, align, lnSpc, lnRed }
  // 各文字の ln に、その行の左余白・字下げ・段落前の間隔（px）を持たせる（makeText で em に直す）
  function readText(txBody, ctx, own, ph, opts) {
    const chain = styleChain(ctx, own, ph);
    const bodyPr = kid(txBody, 'bodyPr');
    const auto = path(bodyPr, 'normAutofit');
    const fontScale = auto ? num(auto.getAttribute('fontScale'), 100000) / 100000 : 1;
    const lnRed = auto ? num(auto.getAttribute('lnSpcReduction'), 0) / 100000 : 0;
    const k = ctx.s * EMU_PT / 2; // 実際の pt → 板書の pt
    const flat = [], sizes = [], jaFlags = [];
    let align = null, lnSpc = null;
    const counters = [];
    const paras = kids(txBody, 'p');
    paras.forEach((p, pi) => {
      const pPr = kid(p, 'pPr');
      const lvl = num(pPr && pPr.getAttribute('lvl'), 0);
      const P = (fn) => findP(pPr, chain, lvl, fn);
      if (pi === 0) {
        const al = P((e) => e.getAttribute('algn') || undefined);
        align = al === 'ctr' ? 'center' : al === 'r' ? 'right' : null;
        const sp = P((e) => kid(e, 'lnSpc') || undefined);
        if (sp && path(sp, 'spcPct')) lnSpc = { pct: num(path(sp, 'spcPct').getAttribute('val'), 100000) / 100000 };
        else if (sp && path(sp, 'spcPts')) lnSpc = { pt: num(path(sp, 'spcPts').getAttribute('val'), 0) / 100 * k };
      }
      const attr = (n) => P((e) => (e.getAttribute(n) != null ? e.getAttribute(n) : undefined));
      const al = P((e) => e.getAttribute('algn') || undefined);
      const ta = al === 'ctr' ? 'center' : al === 'r' ? 'right' : null;
      const marL = opts.noBullets ? 0 : num(attr('marL'), 0) * ctx.s * ctx.group.sx;
      const indent = opts.noBullets ? 0 : num(attr('indent'), 0) * ctx.s * ctx.group.sx;

      const runs = Array.from(p.children).filter((c) => c.localName === 'r' || c.localName === 'br' || c.localName === 'fld');
      const hasText = runs.some((r) => r.localName !== 'br' && (kid(r, 't') || {}).textContent);
      const firstR = runs.find((r) => r.localName !== 'br');
      const at0 = attrsOf(kid(firstR, 'rPr') || kid(p, 'endParaRPr'), pPr, chain, lvl, ctx, fontScale, opts);

      let mt = 0;
      if (pi > 0) {
        flat.push({ ch: '\n', a: {} }); sizes.push(0); jaFlags.push(false);
        const sb = P((e) => kid(e, 'spcBef') || undefined);
        if (sb && path(sb, 'spcPts')) mt = num(path(sb, 'spcPts').getAttribute('val'), 0) / 100 * k * 2;
        else if (sb && path(sb, 'spcPct')) mt = num(path(sb, 'spcPct').getAttribute('val'), 0) / 100000 * at0.pt * 2 * 1.2;
        mt *= 1 - lnRed;
      }
      let ln = { pl: marL, ti: indent, mt, ta, lv: Math.min(4, lvl) };
      const push = (ch, a, pt, ja) => { flat.push({ ch, a, ln }); sizes.push(pt); jaFlags.push(ja); };

      // 箇条書きの記号：行の書式（bu）にする。ぶら下げ字下げの幅（-indent）を記号の幅にして、本文を marL から始める
      if (hasText && !opts.noBullets) {
        const bu = P((e) => {
          for (const c of e.children) {
            if (c.localName === 'buNone') return { none: true };
            if (c.localName === 'buChar') return { ch: c.getAttribute('char') };
            if (c.localName === 'buAutoNum') return { auto: c.getAttribute('type'), start: num(c.getAttribute('startAt'), 1) };
            if (c.localName === 'buBlip') return { ch: '•' };
          }
          return undefined;
        });
        counters.length = lvl + 1;
        if (bu && !bu.none) {
          let text;
          const OL = { arabicPeriod: '1.', arabicParenR: '1)', arabicParenBoth: '(1)', romanLcPeriod: 'i.', romanLcParenR: 'i)', romanUcPeriod: 'I.', alphaLcPeriod: 'a.', alphaLcParenR: 'a)', alphaUcPeriod: 'A.', alphaUcParenR: 'A)' };
          if (bu.auto) {
            counters[lvl] = (counters[lvl] || bu.start - 1) + 1;
            text = autoNum(bu.auto, counters[lvl]);
            // 板書の番号付きリストで表せる形式なら、番号を自動で振る形にする（1 から始まるもののみ）
            if (OL[bu.auto] && bu.start === 1 && lvl <= 4) ln.ol = OL[bu.auto];
          } else {
            const font = P((e) => { const bf = kid(e, 'buFont'); return bf ? bf.getAttribute('typeface') : undefined; });
            text = bulletChar(bu.ch, font);
          }
          const px = at0.pt * 2;
          const own = textWidth(text + ' ', px, at0.ja) || px;
          ln.bu = text.slice(0, 6);
          ln.bw = indent < 0 ? Math.max(-indent, 1) : own;
          const buClr = P((e) => { const b = kid(e, 'buClr'); return b ? readColor(b, ctx) : undefined; });
          ln.bc = buClr ? ctx.colorOf(buClr) : at0.a.c || null;
          const bs = P((e) => { const b = kid(e, 'buSzPct'); return b ? num(b.getAttribute('val'), 100000) / 100000 : undefined; });
          if (bs && Math.abs(bs - 1) > 0.05) ln.bs = Math.round(Math.max(0.3, Math.min(2, bs)) * 100) / 100;
        }
      }

      for (const r of runs) {
        if (r.localName === 'br') {
          flat.push({ ch: '\n', a: {} }); sizes.push(0); jaFlags.push(false);
          ln = { pl: marL, ti: 0, mt: 0, ta, lv: Math.min(4, lvl) }; // 段落内の改行のあとは本文の開始位置から
          continue;
        }
        const at = attrsOf(kid(r, 'rPr'), pPr, chain, lvl, ctx, fontScale, opts);
        const text = (kid(r, 't') || {}).textContent || '';
        for (const ch of text) {
          if (ch === '\n') { flat.push({ ch: '\n', a: {} }); sizes.push(0); jaFlags.push(false); ln = { pl: marL, ti: 0, mt: 0, ta, lv: Math.min(4, lvl) }; continue; }
          push(ch, at.a, at.pt, at.ja);
        }
      }
    });
    return { flat, sizes, jaFlags, align, lnSpc, lnRed };
  }

  function attrsOf(rPr, pPr, chain, lvl, ctx, fontScale, opts) {
    const R = (fn) => findR(rPr, pPr, chain, lvl, fn, opts.fontRefColor);
    const sz = R((e) => (e && e.getAttribute('sz') != null ? num(e.getAttribute('sz'), 1800) : undefined));
    const bool = (n) => R((e) => (e && e.getAttribute(n) != null ? e.getAttribute(n) === '1' || e.getAttribute(n) === 'true' : undefined));
    const bold = bool('b'), ital = bool('i');
    const und = R((e) => (e && e.getAttribute('u') != null ? e.getAttribute('u') : undefined));
    const base = rPr ? num(rPr.getAttribute('baseline'), 0) : 0;
    const col = R((e, frc) => {
      if (frc !== undefined) return frc || undefined;
      const sf = kid(e, 'solidFill');
      return sf ? readColor(sf, ctx) : undefined;
    });
    const latin = R((e) => { const l = kid(e, 'latin'); return l ? l.getAttribute('typeface') : undefined; });
    const hl = rPr && kid(rPr, 'highlight') ? readColor(kid(rPr, 'highlight'), ctx) : null;
    const a = {};
    if (bold || opts.bold) a.b = true;
    if (ital) a.i = true;
    if (und && und !== 'none') {
      const uf = rPr && path(rPr, 'uFill', 'solidFill');
      const uc = uf ? ctx.colorOf(readColor(uf, ctx)) : null;
      const style = /^dbl/.test(und) ? 'dbl' : 'sng';
      a.u = style === 'sng' && !uc ? true : `${style}${uc ? ':' + uc : ''}`;
    }
    const strike = R((e) => (e && e.getAttribute('strike') != null ? e.getAttribute('strike') !== 'noStrike' : undefined));
    if (strike) a.k = true;
    if (base > 0) a.s = 'sup';
    else if (base < 0) a.s = 'sub';
    // 特殊なフォント（梵字・記号など）は、PC に入っていればそのフォントで表示する
    const ea = R((e) => { const l = kid(e, 'ea'); return l ? l.getAttribute('typeface') : undefined; });
    const sym = rPr && kid(rPr, 'sym') ? kid(rPr, 'sym').getAttribute('typeface') : null;
    const special = [sym, latin, ea].find((n) => n && !/^\+m[nj]-/.test(n) && !isCommonFont(n));
    if (special) a.f = special;
    const c = col ? ctx.colorOf(col) : null;
    if (c) a.c = c;
    if (hl) a.h = markerColor(hl);
    const pt = ((sz || 1800) / 100) * fontScale * (ctx.s * EMU_PT / 2);
    return { a, pt, ja: isJaFont(latin, ctx) };
  }

  // 文字列 → 板書のテキストオブジェクト（位置は後で実際の大きさを測って決める）
  function makeText(txBody, ctx, box, own, ph, bodyPrChain, opts) {
    const t = readText(txBody, ctx, own, ph, opts || {});
    if (!t.flat.some((c) => c.ch.trim())) return null;
    // 基準の大きさ：最も多い文字の大きさ
    const count = new Map();
    t.sizes.forEach((v) => { if (v) count.set(v, (count.get(v) || 0) + 1); });
    const basePt = [...count.entries()].sort((a, b) => b[1] - a[1])[0][0];
    const basePx = basePt * 2;
    const em = (px) => Math.round((px / basePx) * 1000) / 1000;
    // タブ：PowerPoint の既定のタブ位置（テキストボックスの左端から 1 インチごと）まで空白で詰める。
    // 板書のタブは字下げのあとから数えるため、字下げのある行では位置がずれてしまうので、ここで空白に置き換える。
    {
      const stop = 914400 * ctx.s;
      let x = null;
      for (let i = 0; i < t.flat.length; i++) {
        const c = t.flat[i];
        if (c.ch === '\n') { x = null; continue; }
        const px = t.sizes[i] * 2;
        if (x === null) x = c.ln ? (c.ln.bu ? c.ln.pl : c.ln.pl + c.ln.ti) : 0;
        if (c.ch === '\t') {
          const next = (Math.floor((x + 0.5) / stop) + 1) * stop;
          const sw = textWidth(' ', px, t.jaFlags[i]) || px * 0.3;
          const n = Math.max(1, Math.round((next - x) / sw));
          const size = t.sizes[i], ja = t.jaFlags[i];
          const spaces = Array.from({ length: n }, () => ({ ch: ' ', a: c.a, ln: c.ln }));
          t.flat.splice(i, 1, ...spaces);
          t.sizes.splice(i, 1, ...spaces.map(() => size));
          t.jaFlags.splice(i, 1, ...spaces.map(() => ja));
          i += n - 1;
          x = next;
        } else {
          x += textWidth(c.ch, px, t.jaFlags[i]);
        }
      }
    }
    const aligns = new Set(t.flat.filter((c) => c.ln).map((c) => c.ln.ta || null));
    const oneAlign = aligns.size <= 1 ? [...aligns][0] || null : undefined;
    t.flat.forEach((c, i) => {
      if (c.ch === '\n') return;
      const extra = {};
      if (oneAlign === undefined && c.ln && c.ln.ta) extra.ta = c.ln.ta;
      if (t.sizes[i] !== basePt) extra.z = B.rich.ratio(t.sizes[i] / basePt);
      if (c.ln) {
        if (Math.abs(c.ln.pl) > 0.5) extra.pl = em(c.ln.pl);
        if (Math.abs(c.ln.ti) > 0.5) extra.ti = em(c.ln.ti);
        if (c.ln.mt > 0.5) extra.mt = em(c.ln.mt);
        if (c.ln.bu || (c.ln.pl && c.ln.lv)) extra.lv = c.ln.lv;
        if (c.ln.bu) {
          if (c.ln.ol) extra.ol = c.ln.ol;
          else extra.bu = c.ln.bu;
          extra.bw = em(c.ln.bw);
          if (c.ln.bc) extra.bc = c.ln.bc;
          if (c.ln.bs) extra.bs = c.ln.bs;
        }
      }
      c.a = Object.assign({}, c.a, extra);
    });
    const bp = (name, d) => {
      for (const b of bodyPrChain) if (b && b.getAttribute(name) != null) return b.getAttribute(name);
      return d;
    };
    const lIns = num(bp('lIns', 91440), 91440) * ctx.s, rIns = num(bp('rIns', 91440), 91440) * ctx.s;
    const tIns = num(bp('tIns', 45720), 45720) * ctx.s, bIns = num(bp('bIns', 45720), 45720) * ctx.s;
    const wrapNone = bp('wrap', 'square') === 'none';
    const anchor = bp('anchor', 't');
    if (bp('vert', 'horz') !== 'horz') ctx.warn('vert');
    const obj = {
      id: B.uid('t'),
      type: 'text',
      x: Math.round(box.x + lIns),
      y: Math.round(box.y + tIns),
      w: wrapNone ? null : Math.max(40, Math.round(box.w - lIns - rIns)),
      fs: Math.round(basePx * 10) / 10,
      html: B.rich.render(t.flat),
    };
    if (oneAlign) obj.align = oneAlign;
    const numCol = num(bp('numCol', 1), 1);
    if (numCol >= 2 && numCol <= 4) {
      obj.cols = numCol;
      obj.colGap = Math.round(num(bp('spcCol', 0), 0) * ctx.s);
    }
    // 行間：PowerPoint の「1行」はおよそ文字の大きさの 1.2 倍
    let lh = 1.2;
    if (t.lnSpc && t.lnSpc.pct) lh = 1.2 * t.lnSpc.pct;
    else if (t.lnSpc && t.lnSpc.pt) lh = t.lnSpc.pt / basePt;
    lh *= 1 - t.lnRed;
    obj.lh = Math.round(Math.max(0.8, lh) * 100) / 100;
    // 欧文も和文フォント（游ゴシックなど）で書かれていれば、空白で揃えた位置がずれないよう同じにする
    if (t.jaFlags.filter(Boolean).length * 2 > t.flat.filter((c) => c.ch !== '\n').length) obj.font = 'ja';
    // 実際の高さを測って縦位置・中央揃えを合わせるための情報
    obj._place = { box: { x: box.x + lIns, y: box.y + tIns, w: box.w - lIns - rIns, h: box.h - tIns - bIns }, anchor, wrapNone };
    return obj;
  }

  // ---------------- インク（手書き）----------------
  // InkML の trace は「明示値 / 1 階差分（'）/ 2 階差分（"）」で圧縮されている
  function decodeTrace(text, n) {
    const pts = [];
    const val = new Array(n).fill(0), vel = new Array(n).fill(0), mode = new Array(n).fill('!');
    for (const raw of text.split(',')) {
      const toks = raw.match(/[!'"]?\s*(?:-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?|[?*])/g);
      if (!toks) continue;
      for (let c = 0; c < Math.min(n, toks.length); c++) {
        let t = toks[c].trim();
        if (/^[!'"]/.test(t)) { mode[c] = t[0]; t = t.slice(1).trim(); }
        if (t === '?' || t === '*') continue;
        const x = parseFloat(t);
        if (mode[c] === '!') { vel[c] = x - val[c]; val[c] = x; }
        else if (mode[c] === "'") { vel[c] = x; val[c] += x; }
        else { vel[c] += x; val[c] += vel[c]; }
      }
      pts.push(val.slice());
    }
    return pts;
  }

  async function readInk(cp, ctx, out) {
    const rid = rAttr(cp, 'id');
    const rel = rid && ctx.rels.get(rid);
    const box = xfrmBox(kid(cp, 'xfrm'), ctx);
    const root = rel ? await ctx.pkg.xml(rel.target) : null;
    if (!root || !box) return false;
    const all = (name) => Array.from(root.getElementsByTagNameNS('*', name));
    const xmlId = (e) => e.getAttribute('xml:id') || e.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'id');
    const chans = all('traceFormat')[0] ? Array.from(all('traceFormat')[0].children).filter((c) => c.localName === 'channel') : [];
    const names = chans.map((c) => c.getAttribute('name'));
    const ix = names.indexOf('X'), iy = names.indexOf('Y'), iF = names.indexOf('F');
    if (ix < 0 || iy < 0) return false;
    const fMax = iF >= 0 ? num(chans[iF].getAttribute('max'), 32767) || 32767 : 0;
    // 座標の単位（1 単位 = 1/res cm）。既定は 1/1000 cm
    let res = 1000;
    for (const cp2 of all('channelProperty')) {
      if (cp2.getAttribute('channel') === 'X' && cp2.getAttribute('name') === 'resolution') res = num(cp2.getAttribute('value'), 1000) || 1000;
    }
    const emuPerUnit = 360000 / res;
    const brushes = {};
    for (const b of all('brush')) {
      const props = {};
      for (const p of Array.from(b.children)) if (p.localName === 'brushProperty') props[p.getAttribute('name')] = p.getAttribute('value');
      brushes[xmlId(b)] = props;
    }
    const traces = all('trace').map((tr) => ({ pts: decodeTrace(tr.textContent || '', names.length), brush: brushes[(tr.getAttribute('brushRef') || '').replace('#', '')] || {} }));
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const t of traces) for (const p of t.pts) { x0 = Math.min(x0, p[ix]); x1 = Math.max(x1, p[ix]); y0 = Math.min(y0, p[iy]); y1 = Math.max(y1, p[iy]); }
    if (!isFinite(x0)) return false;
    // インクの外接矩形の中心を、スライド上の配置（xfrm）の中心に合わせる
    const sx = emuPerUnit * ctx.s * ctx.group.sx, sy = emuPerUnit * ctx.s * ctx.group.sy;
    const cx = box.x + box.w / 2 - ((x0 + x1) / 2) * sx, cy = box.y + box.h / 2 - ((y0 + y1) / 2) * sy;
    for (const t of traces) {
      if (!t.pts.length) continue;
      const b = t.brush;
      const marker = b.rasterOp === 'maskPen' || num(b.transparency, 0) > 0;
      const wpx = num(b.width, 0.05) * 360000 * ctx.s * ctx.group.sx; // 幅は cm
      const col = { hex: (b.color || '#000000').replace('#', '').toUpperCase(), alpha: 1 };
      const pts = [];
      for (const p of t.pts) {
        const x = Math.round((cx + p[ix] * sx) * 10) / 10, y = Math.round((cy + p[iy] * sy) * 10) / 10;
        const last = pts[pts.length - 1];
        if (last && Math.hypot(x - last[0], y - last[1]) < 1) continue;
        const pr = fMax ? Math.round((0.2 + 0.6 * Math.min(1, Math.max(0, p[iF] / fMax))) * 100) / 100 : 0.5;
        pts.push([x, y, pr]);
      }
      out.push({
        id: B.uid('s'),
        type: 'stroke',
        kind: marker ? 'marker' : 'pen',
        color: marker ? (ctx.keep ? '#' + col.hex.toLowerCase() : markerColor(col)) : ctx.colorOf(col) || 'default',
        size: Math.max(2, Math.round(wpx * 10) / 10),
        pr: !!fMax && !marker,
        pts,
      });
    }
    return true;
  }

  // ---------------- 図形の読み込み ----------------

  function xfrmBox(xfrm, ctx) {
    const off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext');
    if (!off || !ext) return null;
    const x = num(off.getAttribute('x'), 0), y = num(off.getAttribute('y'), 0);
    const cx = num(ext.getAttribute('cx'), 0), cy = num(ext.getAttribute('cy'), 0);
    const g = ctx.group;
    const gx = g.ax + g.sx * x, gy = g.ay + g.sy * y;
    return {
      x: ctx.ox + gx * ctx.s, y: ctx.oy + gy * ctx.s, w: cx * g.sx * ctx.s, h: cy * g.sy * ctx.s,
      rot: num(xfrm.getAttribute('rot'), 0) / 60000,
      flipH: xfrm.getAttribute('flipH') === '1', flipV: xfrm.getAttribute('flipV') === '1',
    };
  }

  // プレースホルダーの継承元（レイアウト・マスター上の同じプレースホルダー）
  function findPh(spTree, ph) {
    if (!spTree) return null;
    const type = ph.getAttribute('type') || 'body';
    const idx = ph.getAttribute('idx');
    const all = [];
    for (const sp of spTree.children) {
      const p = path(sp, 'nvSpPr', 'nvPr', 'ph') || path(sp, 'nvPicPr', 'nvPr', 'ph');
      if (p) all.push([sp, p]);
    }
    const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
    if (idx != null) {
      const m = all.find(([, p]) => p.getAttribute('idx') === idx);
      if (m) return m[0];
    }
    const m = all.find(([, p]) => (p.getAttribute('type') || 'body') === type) || all.find(([, p]) => norm(p.getAttribute('type') || 'body') === norm(type));
    return m ? m[0] : null;
  }

  async function readSp(sp, ctx, out) {
    const nv = kid(sp, 'nvSpPr') || kid(sp, 'nvCxnSpPr');
    if (path(nv, 'cNvPr') && path(nv, 'cNvPr').getAttribute('hidden') === '1') return;
    const phEl = path(nv, 'nvPr', 'ph');
    if (phEl && ctx.skipPh) return; // マスター・レイアウトのデザインを読むときはプレースホルダーを除く
    const phType = phEl ? phEl.getAttribute('type') || 'body' : null;
    if (phEl && ['dt', 'ftr', 'sldNum', 'hdr'].includes(phType)) return; // 日付・フッター・スライド番号は読まない
    const ph = phEl ? {
      type: phType,
      layout: findPh(ctx.layoutTree, phEl),
      master: null,
    } : null;
    if (ph) ph.master = findPh(ctx.masterTree, ph.layout ? path(ph.layout, 'nvSpPr', 'nvPr', 'ph') || phEl : phEl);

    const spPr = kid(sp, 'spPr');
    const xf = kid(spPr, 'xfrm') || (ph && path(ph.layout, 'spPr', 'xfrm')) || (ph && path(ph.master, 'spPr', 'xfrm'));
    const box = xfrmBox(xf, ctx);
    if (!box) return;
    const style = kid(sp, 'style');
    const fill = readFill(spPr, style, ctx);
    const line = readLine(spPr, style, ctx);
    const geom = kid(spPr, 'prstGeom'), cust = kid(spPr, 'custGeom');
    const prst = geom ? geom.getAttribute('prst') : cust ? null : 'rect';
    const av = kid(geom, 'avLst');
    const isLine = prst === 'line' || prst === 'straightConnector1';

    // 塗りに画像（数式の代替画像など）
    if (fill && fill.blip) {
      const pic = await pictureObj(fill.blip, box, ctx);
      if (pic) out.push(pic);
    } else if (isLine && line && !box.rot) {
      // 直線・矢印：板書の図形にする
      let x1 = box.x, y1 = box.y, x2 = box.x + box.w, y2 = box.y + box.h;
      if (box.flipH) [x1, x2] = [x2, x1];
      if (box.flipV) [y1, y2] = [y2, y1];
      const plain = line.dash === 'solid';
      if (plain && !(line.head && line.tail)) {
        if (line.head) [x1, y1, x2, y2] = [x2, y2, x1, y1];
        out.push({ id: B.uid('g'), type: 'shape', shape: line.head || line.tail ? 'arrow' : 'line', x1: r2(x1), y1: r2(y1), x2: r2(x2), y2: r2(y2), color: ctx.colorOf(line.color) || 'default', width: r2(line.w) });
      } else {
        const paths = [{ d: `M0 0L${r2(box.w)} ${r2(box.h)}`, fill: false }];
        out.push(shapeImage(box, paths, null, line, 0, box.flipH, box.flipV, connectorEnds('line', box.w, box.h)));
      }
    } else if ((prst === 'rect' || prst === 'ellipse') && !fill && line && line.dash === 'solid' && !box.rot) {
      out.push({ id: B.uid('g'), type: 'shape', shape: prst, x1: r2(box.x), y1: r2(box.y), x2: r2(box.x + box.w), y2: r2(box.y + box.h), color: ctx.colorOf(line.color) || 'default', width: r2(line.w) });
    } else if (fill || line) {
      let paths = cust ? customPaths(cust, box.w, box.h) : presetPaths(prst, box.w, box.h, av);
      if (!paths) { ctx.warn('geom'); paths = presetPaths('rect', box.w, box.h, null); }
      out.push(shapeImage(box, paths, fill, line, box.rot, box.flipH, box.flipV, connectorEnds(prst, box.w, box.h)));
    }

    // 文字
    const tx = kid(sp, 'txBody');
    if (tx) {
      const fontRef = kid(style, 'fontRef');
      const fontRefColor = fontRef ? readColor(fontRef, ctx) : undefined;
      const bodyPrs = [kid(tx, 'bodyPr'), ph && path(ph.layout, 'txBody', 'bodyPr'), ph && path(ph.master, 'txBody', 'bodyPr')];
      const t = makeText(tx, ctx, box, kid(tx, 'lstStyle'), ph, bodyPrs, { fontRefColor });
      if (t) out.push(t);
    }
  }

  async function pictureObj(blipFill, box, ctx) {
    const blip = kid(blipFill, 'blip');
    const rid = rAttr(blip, 'embed');
    const rel = rid && ctx.rels.get(rid);
    if (!rel || rel.external) { ctx.warn('link'); return null; }
    const ext = (rel.target.split('.').pop() || '').toLowerCase();
    const types = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
    if (!types[ext]) { ctx.warn(ext === 'emf' || ext === 'wmf' ? 'emf' : 'imgfmt'); return null; }
    const bytes = await ctx.pkg.zip.bytes(rel.target);
    if (!bytes) return null;
    let blob = new Blob([bytes], { type: types[ext] });
    // トリミング（srcRect：各辺を 1/1000 % 単位で切り取る）
    const sr = kid(blipFill, 'srcRect');
    if (sr && ext !== 'svg') {
      const l = num(sr.getAttribute('l'), 0) / 100000, t = num(sr.getAttribute('t'), 0) / 100000;
      const r = num(sr.getAttribute('r'), 0) / 100000, b = num(sr.getAttribute('b'), 0) / 100000;
      if (l || t || r || b) blob = await crop(blob, l, t, r, b);
    }
    if (box.rot) ctx.warn('rotPic');
    let img;
    try { img = await B.loadImage(blob); } catch (e) { ctx.warn('imgfmt'); return null; }
    return { id: B.uid('i'), type: 'image', x: r2(box.x), y: r2(box.y), w: r2(box.w), h: r2(box.h), src: img.src };
  }

  async function crop(blob, l, t, r, b) {
    const bmp = await createImageBitmap(blob);
    const sx = Math.max(0, bmp.width * l), sy = Math.max(0, bmp.height * t);
    const sw = Math.max(1, bmp.width * (1 - l - r)), sh = Math.max(1, bmp.height * (1 - t - b));
    const c = document.createElement('canvas');
    c.width = Math.round(sw);
    c.height = Math.round(sh);
    c.getContext('2d').drawImage(bmp, sx, sy, sw, sh, 0, 0, c.width, c.height);
    return await new Promise((res) => c.toBlob(res, blob.type === 'image/jpeg' ? 'image/jpeg' : 'image/png', 0.92));
  }

  async function readPic(pic, ctx, out) {
    const nv = kid(pic, 'nvPicPr');
    if (path(nv, 'cNvPr') && path(nv, 'cNvPr').getAttribute('hidden') === '1') return;
    if (ctx.skipPh && path(nv, 'nvPr', 'ph')) return;
    if (path(nv, 'nvPr', 'videoFile') || path(nv, 'nvPr', 'audioFile')) { ctx.warn('media'); return; }
    const phEl = path(nv, 'nvPr', 'ph');
    const spPr = kid(pic, 'spPr');
    let xf = kid(spPr, 'xfrm');
    if (!xf && phEl) {
      const l = findPh(ctx.layoutTree, phEl);
      xf = path(l, 'spPr', 'xfrm');
    }
    const box = xfrmBox(xf, ctx);
    if (!box) return;
    const o = await pictureObj(kid(pic, 'blipFill'), box, ctx);
    if (o) out.push(o);
  }

  async function readGraphicFrame(gf, ctx, out) {
    const gd = path(gf, 'graphic', 'graphicData');
    const uri = gd ? gd.getAttribute('uri') || '' : '';
    const box = xfrmBox(kid(gf, 'xfrm'), ctx);
    if (!box) return;
    if (/table$/.test(uri)) {
      readTable(kid(gd, 'tbl'), box, ctx, out);
    } else if (/diagram$/.test(uri)) {
      // SmartArt：PowerPoint が保存している描画結果（drawing パート）を図形として読む
      const ids = kid(gd, 'relIds');
      const dm = ids && ctx.rels.get(rAttr(ids, 'dm'));
      const dmRoot = dm ? await ctx.pkg.xml(dm.target) : null;
      const ext = descendant(dmRoot, 'dataModelExt');
      const rel = ext && ctx.rels.get(ext.getAttribute('relId'));
      const drawing = rel ? await ctx.pkg.xml(rel.target) : null;
      const tree = descendant(drawing, 'spTree');
      if (tree) {
        const saved = ctx.rels;
        ctx.rels = await ctx.pkg.rels(rel.target);
        await readTree(tree, ctx, out);
        ctx.rels = saved;
      } else {
        ctx.warn('smartart');
      }
    } else if (/chart/.test(uri)) {
      ctx.warn('chart');
    } else if (/ole$/.test(uri)) {
      // 埋め込みオブジェクト：代わりの画像があればそれを使う
      const pic = descendant(gd, 'pic');
      if (pic) await readPic(pic, ctx, out);
      else ctx.warn('ole');
    } else {
      ctx.warn('other');
    }
  }

  // 表：板書の表オブジェクトにする（セルの文字はそれぞれ編集できるテキスト）
  function readTable(tbl, box, ctx, out) {
    if (!tbl) return;
    const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => Math.max(20, Math.round(num(c.getAttribute('w'), 0) * ctx.s * ctx.group.sx)));
    const trs = kids(tbl, 'tr');
    if (!cols.length || !trs.length) return;
    const firstRow = path(tbl, 'tblPr') && path(tbl, 'tblPr').getAttribute('firstRow') === '1';
    const rows = trs.map((tr) => Math.max(10, Math.round(num(tr.getAttribute('h'), 0) * ctx.s * ctx.group.sy)));
    const cells = trs.map(() => cols.map(() => ({ html: '' })));
    const texts = [];
    let border = null;
    trs.forEach((tr, ri) => {
      let ci = 0;
      for (const tc of kids(tr, 'tc')) {
        if (ci >= cols.length) break;
        const span = num(tc.getAttribute('gridSpan'), 1);
        const merged = tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1';
        const tcPr = kid(tc, 'tcPr');
        const cell = cells[ri][ci];
        if (!merged) {
          const tb = kid(tc, 'txBody');
          if (tb) {
            const w = cols.slice(ci, ci + span).reduce((a, b) => a + b, 0);
            const t = makeText(tb, ctx, { x: 0, y: 0, w, h: rows[ri] }, kid(tb, 'lstStyle'), null, [], { bold: firstRow && ri === 0, noBullets: true });
            if (t) { texts.push([cell, t]); if (t.align) cell.al = t.align; }
          }
          const anchor = tcPr && tcPr.getAttribute('anchor');
          if (anchor === 'ctr') cell.va = 'middle';
          else if (anchor === 'b') cell.va = 'bottom';
          if (ctx.keep) {
            const sf = kid(tcPr, 'solidFill');
            const col = sf ? readColor(sf, ctx) : null;
            if (col) cell.fill = '#' + col.hex.toLowerCase();
          }
        }
        if (!border) {
          for (const n of ['lnB', 'lnT', 'lnL', 'lnR']) {
            const ln = kid(tcPr, n);
            if (ln && !kid(ln, 'noFill')) {
              const col = readColor(kid(ln, 'solidFill'), ctx);
              border = { color: col ? ctx.colorOf(col) : null, w: Math.max(1, num(ln.getAttribute('w'), 12700) * ctx.s) };
              break;
            }
          }
        }
        ci += span;
      }
    });
    // 表全体の文字の大きさ：最も多いセルの大きさ。違うセルは倍率で合わせる
    const sizes = new Map();
    for (const [, t] of texts) sizes.set(t.fs, (sizes.get(t.fs) || 0) + 1);
    const fs = texts.length ? [...sizes.entries()].sort((a, b) => b[1] - a[1])[0][0] : 36;
    for (const [cell, t] of texts) {
      let html = t.html;
      if (t.fs !== fs) {
        const flat = B.rich.parseHtml(html);
        for (const ch of flat) if (ch.ch !== '\n') ch.a = Object.assign({}, ch.a, { z: B.rich.ratio((ch.a.z || 1) * t.fs / fs) });
        html = B.rich.render(flat);
      }
      cell.html = html;
    }
    const obj = { id: B.uid('b'), type: 'table', x: Math.round(box.x), y: Math.round(box.y), fs, cols, rows, cells };
    const first = texts[0] && texts[0][1];
    if (first && first.lh) obj.lh = first.lh;
    if (texts.length && texts.filter(([, t]) => t.font === 'ja').length * 2 >= texts.length) obj.font = 'ja';
    obj.color = (border && border.color) || (ctx.keep ? '#808080' : 'default');
    obj.bw = Math.round(((border && border.w) || 2) * 10) / 10;
    out.push(obj);
  }

  async function readTree(tree, ctx, out) {
    for (let el of Array.from(tree.children)) {
      if (el.localName === 'AlternateContent') {
        // 数式などを含む新しい形式（Choice）は扱えないので、代わりの表現（Fallback）を使う
        const choice = kid(el, 'Choice');
        const useFallback = !choice || choice.getElementsByTagNameNS('*', 'oMath').length || choice.getElementsByTagNameNS('*', 'oMathPara').length || (!kid(choice, 'contentPart') && /p14|a14|v|w14/.test(choice.getAttribute('Requires') || '') && false);
        if (useFallback && choice && (choice.getElementsByTagNameNS('*', 'oMath').length)) ctx.warn('math');
        const pick = useFallback ? kid(el, 'Fallback') : choice;
        if (!pick) continue;
        for (const c of Array.from(pick.children)) await readOne(c, ctx, out);
        continue;
      }
      await readOne(el, ctx, out);
    }
  }

  async function readOne(el, ctx, out) {
    switch (el.localName) {
      case 'sp': case 'cxnSp': await readSp(el, ctx, out); break;
      case 'pic': await readPic(el, ctx, out); break;
      case 'graphicFrame': await readGraphicFrame(el, ctx, out); break;
      case 'grpSp': {
        const xf = path(el, 'grpSpPr', 'xfrm');
        const off = kid(xf, 'off'), ext = kid(xf, 'ext'), chOff = kid(xf, 'chOff'), chExt = kid(xf, 'chExt');
        const saved = ctx.group, savedFill = ctx.groupFill;
        if (off && ext && chOff && chExt) {
          const sx = num(chExt.getAttribute('cx'), 0) ? num(ext.getAttribute('cx'), 0) / num(chExt.getAttribute('cx'), 1) : 1;
          const sy = num(chExt.getAttribute('cy'), 0) ? num(ext.getAttribute('cy'), 0) / num(chExt.getAttribute('cy'), 1) : 1;
          const ax = num(off.getAttribute('x'), 0) - num(chOff.getAttribute('x'), 0) * sx;
          const ay = num(off.getAttribute('y'), 0) - num(chOff.getAttribute('y'), 0) * sy;
          // 親のグループ変換と合成
          ctx.group = { ax: saved.ax + saved.sx * ax, ay: saved.ay + saved.sy * ay, sx: saved.sx * sx, sy: saved.sy * sy };
        }
        const gf = readFill(kid(el, 'grpSpPr'), null, ctx);
        if (gf) ctx.groupFill = gf;
        await readTree(el, ctx, out);
        ctx.group = saved;
        ctx.groupFill = savedFill;
        break;
      }
      case 'contentPart': if (!(await readInk(el, ctx, out))) ctx.warn('ink'); break;
    }
  }

  // ---------------- 文字の大きさを測って配置を確定する ----------------

  function makeMeasurer() {
    const host = document.createElement('div');
    host.style.cssText = 'position:absolute;left:-100000px;top:0;width:1920px;height:1080px;visibility:hidden';
    document.body.appendChild(host);
    const fake = { on: () => () => {}, doc: { theme: 'white' }, page: { objects: [] }, pageById: () => null };
    const view = new B.View(fake, document, host, {});
    return {
      measure(obj) {
        const el = view.create(obj);
        view.root.appendChild(el);
        const r = { w: el.offsetWidth, h: el.offsetHeight };
        el.remove();
        return r;
      },
      done() { view.destroy(); host.remove(); },
    };
  }

  function placeText(o, m) {
    const p = o._place;
    delete o._place;
    const sz = m.measure(o);
    if (p.anchor === 'ctr') o.y = Math.round(p.box.y + (p.box.h - sz.h) / 2);
    else if (p.anchor === 'b') o.y = Math.round(p.box.y + p.box.h - sz.h);
    else o.y = Math.round(p.box.y);
    if (p.wrapNone && o.align) {
      if (o.align === 'center') o.x = Math.round(p.box.x + (p.box.w - sz.w) / 2);
      else if (o.align === 'right') o.x = Math.round(p.box.x + p.box.w - sz.w);
    }
    return sz;
  }

  // ---------------- 本体 ----------------

  const WARN_LABELS = {
    chart: 'グラフ', smartart: 'SmartArt', ole: '埋め込みオブジェクト', media: '動画・音声',
    emf: 'EMF/WMF 形式の画像（PowerPoint で PNG に変換すれば読めます）', imgfmt: '対応していない形式の画像',
    link: 'リンクされた画像', geom: '未対応の形の図形（四角形で代用）', rotPic: '回転した画像（回転なしで配置）',
    math: '数式（PowerPoint が保存した画像で配置）', vert: '縦書きの文字（横書きで配置）', ink: '読めなかったインク（手書き）',
    other: 'その他の要素',
  };

  // スライドの背景（色・画像）と、マスター・レイアウトのデザイン（プレースホルダー以外の図形）
  async function pageBackground(parts, ctx) {
    const bg = { color: null, objs: [] };
    const [slide, layout, master] = parts;
    const src = parts.find((p) => p.root && path(p.root, 'cSld', 'bg'));
    if (src) {
      const c = Object.assign({}, ctx, { rels: src.rels });
      const el = path(src.root, 'cSld', 'bg');
      const bgPr = kid(el, 'bgPr'), bgRef = kid(el, 'bgRef');
      const fill = bgPr ? readFill(bgPr, null, c) : bgRef ? readColor(bgRef, c) : null;
      if (fill && fill.blip) {
        const pic = await pictureObj(fill.blip, { x: ctx.ox, y: ctx.oy, w: B.W - ctx.ox * 2, h: B.H - ctx.oy * 2 }, c);
        if (pic) bg.objs.push(pic);
      } else if (fill && !/^F[A-F]F[A-F]F[A-F]$/i.test(fill.hex)) {
        bg.color = '#' + fill.hex.toLowerCase();
      }
    }
    const show = (p) => !p || !p.root || p.root.getAttribute('showMasterSp') !== '0';
    if (show(slide)) {
      const deco = [];
      if (master && show(layout)) await readTree(path(master.root, 'cSld', 'spTree'), Object.assign({}, ctx, { rels: master.rels, skipPh: true, group: { ax: 0, ay: 0, sx: 1, sy: 1 } }), deco);
      if (layout) await readTree(path(layout.root, 'cSld', 'spTree'), Object.assign({}, ctx, { rels: layout.rels, skipPh: true, group: { ax: 0, ay: 0, sx: 1, sy: 1 } }), deco);
      bg.objs.push(...deco);
    }
    return bg.color || bg.objs.length ? bg : null;
  }

  // 読み込みの前に、デザインのある（背景色やマスターの装飾がある）プレゼンテーションかを調べる
  B.probePptx = async (file) => {
    const pkg = new Pkg(unzip(await file.arrayBuffer()));
    const pres = await pkg.xml('ppt/presentation.xml');
    if (!pres) throw new Error('PowerPoint のプレゼンテーション（.pptx）ではありません。');
    const presRels = await pkg.rels('ppt/presentation.xml');
    const slides = kids(kid(pres, 'sldIdLst'), 'sldId');
    const seen = new Set();
    let designed = false;
    const hasDeco = (root) => {
      const tree = path(root, 'cSld', 'spTree');
      if (!tree) return false;
      return Array.from(tree.children).some((c) => /^(sp|pic|grpSp|cxnSp)$/.test(c.localName) && !c.getElementsByTagNameNS('*', 'ph').length);
    };
    const darkBg = (root) => {
      const bg = path(root, 'cSld', 'bg');
      if (!bg) return false;
      const hex = (descendant(bg, 'srgbClr') || { getAttribute: () => 'FFFFFF' }).getAttribute('val') || 'FFFFFF';
      return !!descendant(bg, 'blipFill') || !!descendant(bg, 'gradFill') || !/^F[A-F]F[A-F]F[A-F]$/i.test(hex) || (!!descendant(bg, 'schemeClr') && !/^(bg1|lt1)$/.test(descendant(bg, 'schemeClr').getAttribute('val')));
    };
    for (const sid of slides) {
      const rel = presRels.get(rAttr(sid, 'id'));
      if (!rel) continue;
      const sroot = await pkg.xml(rel.target);
      if (sroot && darkBg(sroot)) designed = true;
      const lrel = [...(await pkg.rels(rel.target)).values()].find((r) => r.type === 'slideLayout');
      if (!lrel || seen.has(lrel.target)) continue;
      seen.add(lrel.target);
      const lroot = await pkg.xml(lrel.target);
      const mrel = [...(await pkg.rels(lrel.target)).values()].find((r) => r.type === 'slideMaster');
      const mroot = mrel ? await pkg.xml(mrel.target) : null;
      if ((lroot && (darkBg(lroot) || hasDeco(lroot))) || (mroot && (darkBg(mroot) || hasDeco(mroot)))) designed = true;
    }
    return { slides: slides.length, designed };
  };

  // file（.pptx の File / Blob）→ { doc, warnings: [文字列] }
  // opts.mode: 'keep'（元の色・デザインのまま）| 'board'（板書の色に合わせる）
  B.importPptx = async (file, onProgress, opts) => {
    const keep = !!(opts && opts.mode === 'keep');
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    const zip = unzip(await file.arrayBuffer());
    const pkg = new Pkg(zip);
    const presPath = 'ppt/presentation.xml';
    const pres = await pkg.xml(presPath);
    if (!pres) throw new Error('PowerPoint のプレゼンテーション（.pptx）ではありません。');
    const presRels = await pkg.rels(presPath);
    const sz = kid(pres, 'sldSz');
    const cx = num(sz && sz.getAttribute('cx'), 12192000), cy = num(sz && sz.getAttribute('cy'), 6858000);
    // 高さを 1080 に合わせ、横に余りが出れば中央に寄せる（4:3 など）
    const s = Math.min(B.W / cx, B.H / cy);
    const ox = (B.W - cx * s) / 2, oy = (B.H - cy * s) / 2;

    const warnCount = new Map();
    const warn = (k) => warnCount.set(k, (warnCount.get(k) || 0) + 1);
    const defaultTextStyle = kid(pres, 'defaultTextStyle');
    const slideIds = kids(kid(pres, 'sldIdLst'), 'sldId');
    const pages = [];
    const measurer = makeMeasurer();
    try {
      for (let i = 0; i < slideIds.length; i++) {
        if (onProgress) onProgress(i + 1, slideIds.length);
        const rel = presRels.get(rAttr(slideIds[i], 'id'));
        if (!rel) continue;
        const slidePath = rel.target;
        const slide = await pkg.xml(slidePath);
        if (!slide) continue;
        const rels = await pkg.rels(slidePath);
        const layoutRel = [...rels.values()].find((r) => r.type === 'slideLayout');
        const layoutPath = layoutRel && layoutRel.target;
        const layout = layoutPath ? await pkg.xml(layoutPath) : null;
        const lrels = layoutPath ? await pkg.rels(layoutPath) : new Map();
        const masterRel = [...lrels.values()].find((r) => r.type === 'slideMaster');
        const masterPath = masterRel && masterRel.target;
        const master = masterPath ? await pkg.xml(masterPath) : null;
        const mrels = masterPath ? await pkg.rels(masterPath) : new Map();
        const themeRel = [...mrels.values()].find((r) => r.type === 'theme');
        const theme = await parseTheme(pkg, themeRel && themeRel.target);
        const cm = kid(master, 'clrMap');
        const clrMap = {};
        if (cm) for (const a of cm.attributes) clrMap[a.name] = a.value;
        const txs = kid(master, 'txStyles');
        const ctx = {
          pkg, rels, theme, clrMap, s, ox, oy, warn, keep,
          // 色の値：keep なら '#rrggbb'、board なら意味上の色名（既定色は null）
          colorOf: (col) => {
            if (!col) return null;
            if (keep) return '#' + col.hex.toLowerCase();
            const c = semanticColor(col);
            return c === 'default' ? null : c;
          },
          group: { ax: 0, ay: 0, sx: 1, sy: 1 },
          layoutTree: path(layout, 'cSld', 'spTree'),
          masterTree: path(master, 'cSld', 'spTree'),
          txStyles: { title: kid(txs, 'titleStyle'), body: kid(txs, 'bodyStyle'), other: kid(txs, 'otherStyle') },
          defaultTextStyle,
        };
        const raw = [];
        await readTree(path(slide, 'cSld', 'spTree'), ctx, raw);
        const finish = (list) => {
          const objs = [];
          for (const o of list) {
            if (o.type === 'text') placeText(o, measurer);
            objs.push(o);
          }
          return objs;
        };
        const page = { id: B.uid('p'), objects: finish(raw) };
        if (keep) {
          const bg = await pageBackground([{ root: slide, rels }, layout && { root: layout, rels: lrels }, master && { root: master, rels: mrels }], ctx);
          if (bg) { bg.objs = finish(bg.objs); page.bg = bg; }
        }
        pages.push(page);
      }
    } finally {
      measurer.done();
    }
    if (!pages.length) pages.push(B.newPage());
    const warnings = [...warnCount.entries()].map(([k, n]) => `${WARN_LABELS[k] || k}：${n} 件`);
    return { doc: { theme: 'white', pages }, warnings };
  };
})(window.B);
