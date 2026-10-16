'use strict';
// 図形の描画と、オブジェクト共通の外接矩形・当たり判定・変形。
(function (B) {
  const f1 = (n) => Math.round(n * 10) / 10;

  function arrowHead(o) {
    const L = Math.hypot(o.x2 - o.x1, o.y2 - o.y1);
    const head = Math.min(L * 0.6, Math.max(18, o.width * 4.2));
    return { L, head, ux: L ? (o.x2 - o.x1) / L : 1, uy: L ? (o.y2 - o.y1) / L : 0 };
  }

  // 図形の SVG（中身だけ）。色は意味上の色名から CSS 変数で決まる。
  function shapeMarkup(o) {
    const col = B.colorCss(o.color, 'c');
    const st = `stroke:${col};stroke-width:${o.width};fill:none;stroke-linecap:round;stroke-linejoin:round`;
    const { x1, y1, x2, y2 } = o;
    switch (o.shape) {
      case 'line':
        return `<line x1="${f1(x1)}" y1="${f1(y1)}" x2="${f1(x2)}" y2="${f1(y2)}" style="${st}"/>`;
      case 'arrow': {
        const { L, head, ux, uy } = arrowHead(o);
        if (L < 1) return '';
        const hw = head * 0.5;
        const bx = x2 - ux * head, by = y2 - uy * head;
        const ex = x2 - ux * head * 0.7, ey = y2 - uy * head * 0.7;
        const pts = [[x2, y2], [bx - uy * hw, by + ux * hw], [bx + uy * hw, by - ux * hw]].map((p) => f1(p[0]) + ',' + f1(p[1])).join(' ');
        return `<line x1="${f1(x1)}" y1="${f1(y1)}" x2="${f1(ex)}" y2="${f1(ey)}" style="${st}"/>` +
          `<polygon points="${pts}" style="fill:${col};stroke:${col};stroke-width:${Math.min(o.width, 4)};stroke-linejoin:round"/>`;
      }
      case 'rect':
        return `<rect x="${f1(Math.min(x1, x2))}" y="${f1(Math.min(y1, y2))}" width="${f1(Math.abs(x2 - x1))}" height="${f1(Math.abs(y2 - y1))}" style="${st}"/>`;
      case 'ellipse':
        return `<ellipse cx="${f1((x1 + x2) / 2)}" cy="${f1((y1 + y2) / 2)}" rx="${f1(Math.abs(x2 - x1) / 2)}" ry="${f1(Math.abs(y2 - y1) / 2)}" style="${st}"/>`;
    }
    return '';
  }

  // Shift キーでの制約：線・矢印は 45° 刻み（水平・垂直を含む）、四角・楕円は正方形・正円
  function constrain(shape, x1, y1, x2, y2) {
    const dx = x2 - x1, dy = y2 - y1;
    if (shape === 'line' || shape === 'arrow') {
      const step = Math.PI / 4;
      const a = Math.round(Math.atan2(dy, dx) / step) * step;
      const cos = Math.round(Math.cos(a) * 1e6) / 1e6, sin = Math.round(Math.sin(a) * 1e6) / 1e6;
      const len = dx * cos + dy * sin;
      return { x2: x1 + cos * len, y2: y1 + sin * len };
    }
    const d = Math.max(Math.abs(dx), Math.abs(dy));
    return { x2: x1 + (dx < 0 ? -d : d), y2: y1 + (dy < 0 ? -d : d) };
  }

  function shapeSegs(o) {
    const { x1, y1, x2, y2 } = o;
    if (o.shape === 'line' || o.shape === 'arrow') return [[x1, y1, x2, y2]];
    if (o.shape === 'rect') return [[x1, y1, x2, y1], [x2, y1, x2, y2], [x2, y2, x1, y2], [x1, y2, x1, y1]];
    const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, rx = Math.abs(x2 - x1) / 2, ry = Math.abs(y2 - y1) / 2;
    const segs = [];
    const N = 48;
    for (let i = 0; i < N; i++) {
      const a0 = (i / N) * 2 * Math.PI, a1 = ((i + 1) / N) * 2 * Math.PI;
      segs.push([cx + rx * Math.cos(a0), cy + ry * Math.sin(a0), cx + rx * Math.cos(a1), cy + ry * Math.sin(a1)]);
    }
    return segs;
  }

  function distSeg(px, py, ax, ay, bx, by) {
    const dx = bx - ax, dy = by - ay;
    const l2 = dx * dx + dy * dy;
    let t = l2 ? ((px - ax) * dx + (py - ay) * dy) / l2 : 0;
    t = Math.max(0, Math.min(1, t));
    return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
  }

  // textRect(obj) はテキストの実際の大きさ（DOM から測る）を返す関数
  function bbox(o, textRect) {
    if (o.type === 'stroke') {
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      for (const p of o.pts) {
        if (p[0] < x0) x0 = p[0];
        if (p[0] > x1) x1 = p[0];
        if (p[1] < y0) y0 = p[1];
        if (p[1] > y1) y1 = p[1];
      }
      const r = (o.pr ? o.size * 0.85 : o.size / 2);
      return { x0: x0 - r, y0: y0 - r, x1: x1 + r, y1: y1 + r };
    }
    if (o.type === 'shape') {
      let r = o.width / 2;
      if (o.shape === 'arrow') r = Math.max(r, arrowHead(o).head * 0.5);
      return { x0: Math.min(o.x1, o.x2) - r, y0: Math.min(o.y1, o.y2) - r, x1: Math.max(o.x1, o.x2) + r, y1: Math.max(o.y1, o.y2) + r };
    }
    if (o.type === 'text' || o.type === 'table') return textRect(o);
    if (o.type === 'image') return { x0: o.x, y0: o.y, x1: o.x + o.w, y1: o.y + o.h };
    return { x0: 0, y0: 0, x1: 0, y1: 0 };
  }

  // 選択ツールでのクリック判定（四角・楕円は内側も含む）
  function hitSelect(o, x, y, tol, textRect) {
    if (o.type === 'stroke') return B.ink.hit(o, x, y, tol);
    if (o.type === 'shape') {
      if (o.shape === 'rect' || o.shape === 'ellipse') {
        const b = bbox(o, textRect);
        return x >= b.x0 - tol && x <= b.x1 + tol && y >= b.y0 - tol && y <= b.y1 + tol;
      }
      return shapeSegs(o).some((s) => distSeg(x, y, s[0], s[1], s[2], s[3]) <= tol + o.width / 2);
    }
    if (o.type === 'text' || o.type === 'image' || o.type === 'table') {
      const b = bbox(o, textRect);
      return x >= b.x0 && x <= b.x1 && y >= b.y0 && y <= b.y1;
    }
    return false;
  }

  // 消しゴムの当たり判定（線と図形の線だけ）
  function hitErase(o, x, y, r) {
    if (o.type === 'stroke') return B.ink.hit(o, x, y, r);
    if (o.type === 'shape') return shapeSegs(o).some((s) => distSeg(x, y, s[0], s[1], s[2], s[3]) <= r + o.width / 2);
    return false;
  }

  // 平行移動 (dx, dy) と、(ox, oy) を中心とする拡大縮小 s を施した複製を返す
  function transform(o, t) {
    const s = t.s || 1, ox = t.ox || 0, oy = t.oy || 0, dx = t.dx || 0, dy = t.dy || 0;
    const fx = (x) => f1(ox + (x - ox) * s + dx);
    const fy = (y) => f1(oy + (y - oy) * s + dy);
    const c = B.cloneObj(o);
    if (o.type === 'stroke') {
      c.pts = o.pts.map((p) => [fx(p[0]), fy(p[1]), p[2]]);
      if (s !== 1) c.size = f1(o.size * s);
    } else if (o.type === 'shape') {
      c.x1 = fx(o.x1); c.y1 = fy(o.y1); c.x2 = fx(o.x2); c.y2 = fy(o.y2);
    } else if (o.type === 'table') {
      c.x = Math.round(fx(o.x)); c.y = Math.round(fy(o.y));
      if (s !== 1) {
        c.fs = f1(o.fs * s);
        c.cols = o.cols.map((w) => Math.round(w * s));
        c.rows = o.rows.map((h) => Math.round(h * s));
      }
    } else if (o.type === 'image') {
      c.x = fx(o.x); c.y = fy(o.y);
      c.w = f1(o.w * s); c.h = f1(o.h * s);
    } else if (o.type === 'text') {
      c.x = Math.round(fx(o.x)); c.y = Math.round(fy(o.y));
      if (s !== 1) {
        c.fs = f1(o.fs * s);
        if (o.w) c.w = Math.round(o.w * s);
      }
    }
    return c;
  }

  B.geom = { shapeMarkup, constrain, bbox, hitSelect, hitErase, transform };
})(window.B);
