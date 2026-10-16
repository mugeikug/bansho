'use strict';
// 手書き線の輪郭生成（筆圧で太さが変わる塗りつぶしパス）と当たり判定。
(function (B) {
  const f1 = (n) => Math.round(n * 10) / 10;

  // [1,2,1] の移動平均を2回。端点は動かさないので書き終わりの位置がずれない。
  function smooth(pts) {
    let a = pts;
    for (let pass = 0; pass < 2; pass++) {
      if (a.length < 3) break;
      const b = [a[0]];
      for (let i = 1; i < a.length - 1; i++) {
        b.push([
          (a[i - 1][0] + 2 * a[i][0] + a[i + 1][0]) / 4,
          (a[i - 1][1] + 2 * a[i][1] + a[i + 1][1]) / 4,
          (a[i - 1][2] + 2 * a[i][2] + a[i + 1][2]) / 4,
        ]);
      }
      b.push(a[a.length - 1]);
      a = b;
    }
    return a;
  }

  // 筆圧 → 太さの倍率。筆圧 0.5 で 1.0 倍。
  const pressureScale = (p) => Math.max(0.3, Math.min(1.7, 0.3 + 1.4 * p));

  function circle(x, y, r) {
    return `M${f1(x - r)} ${f1(y)}a${f1(r)} ${f1(r)} 0 1 0 ${f1(2 * r)} 0a${f1(r)} ${f1(r)} 0 1 0 ${f1(-2 * r)} 0Z`;
  }

  // 点列 [[x, y, pressure], ...] から SVG パスの d を作る
  function outline(pts, size, usePressure) {
    if (!pts.length) return '';
    const s = smooth(pts);
    const n = s.length;
    const w = s.map((p) => (usePressure ? size * pressureScale(p[2]) : size) / 2);
    if (n === 1) return circle(s[0][0], s[0][1], w[0]);

    const left = [], right = [];
    let nx = 0, ny = 0;
    for (let i = 0; i < n; i++) {
      const a = s[Math.max(0, i - 1)], b = s[Math.min(n - 1, i + 1)];
      let dx = b[0] - a[0], dy = b[1] - a[1];
      const len = Math.hypot(dx, dy);
      if (len > 1e-6) { nx = -dy / len; ny = dx / len; }
      left.push([s[i][0] + nx * w[i], s[i][1] + ny * w[i]]);
      right.push([s[i][0] - nx * w[i], s[i][1] - ny * w[i]]);
    }
    if (nx === 0 && ny === 0) return circle(s[0][0], s[0][1], Math.max(...w));

    const out = [];
    const side = (arr) => {
      for (let i = 1; i < arr.length - 1; i++) {
        const mx = (arr[i][0] + arr[i + 1][0]) / 2, my = (arr[i][1] + arr[i + 1][1]) / 2;
        out.push(`Q${f1(arr[i][0])} ${f1(arr[i][1])} ${f1(mx)} ${f1(my)}`);
      }
      const last = arr[arr.length - 1];
      out.push(`L${f1(last[0])} ${f1(last[1])}`);
    };
    out.push(`M${f1(left[0][0])} ${f1(left[0][1])}`);
    side(left);
    const re = w[n - 1];
    out.push(`A${f1(re)} ${f1(re)} 0 0 0 ${f1(right[n - 1][0])} ${f1(right[n - 1][1])}`);
    side(right.slice().reverse());
    const rs = w[0];
    out.push(`A${f1(rs)} ${f1(rs)} 0 0 0 ${f1(left[0][0])} ${f1(left[0][1])}Z`);
    return out.join('');
  }

  function bbox(obj) {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const p of obj.pts) {
      if (p[0] < x0) x0 = p[0];
      if (p[0] > x1) x1 = p[0];
      if (p[1] < y0) y0 = p[1];
      if (p[1] > y1) y1 = p[1];
    }
    const pad = obj.size;
    return { x0: x0 - pad, y0: y0 - pad, x1: x1 + pad, y1: y1 + pad };
  }

  function distSeg(px, py, ax, ay, bx, by) {
    const dx = bx - ax, dy = by - ay;
    const l2 = dx * dx + dy * dy;
    let t = l2 ? ((px - ax) * dx + (py - ay) * dy) / l2 : 0;
    t = Math.max(0, Math.min(1, t));
    return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
  }

  // 点の間隔が step 以下になるよう補間する（筆圧も補間）
  function densify(pts, step) {
    const out = [pts[0]];
    for (let i = 1; i < pts.length; i++) {
      const a = pts[i - 1], b = pts[i];
      const n = Math.ceil(Math.hypot(b[0] - a[0], b[1] - a[1]) / step);
      for (let k = 1; k < n; k++) {
        const t = k / n;
        out.push([f1(a[0] + (b[0] - a[0]) * t), f1(a[1] + (b[1] - a[1]) * t), Math.round((a[2] + (b[2] - a[2]) * t) * 100) / 100]);
      }
      out.push(b);
    }
    return out;
  }

  // 部分消しゴム：(x, y) 半径 r に掛かる部分を取り除き、残った部分の点列の配列を返す
  function split(obj, x, y, r) {
    const half = (obj.pr ? obj.size * 0.8 : obj.size) / 2;
    const lim = r + half;
    const pts = densify(obj.pts, Math.max(1, r / 3));
    const pieces = [];
    let cur = [];
    for (const p of pts) {
      if (Math.hypot(p[0] - x, p[1] - y) <= lim) {
        if (cur.length) pieces.push(cur);
        cur = [];
      } else {
        cur.push(p);
      }
    }
    if (cur.length) pieces.push(cur);
    // ごく短い切れ端は捨てる
    return pieces.filter((pc) => pc.length >= 2 || (pts.length === 1 && pc.length === 1));
  }

  B.ink = {
    outline,
    split,
    // 描画中は obj._d を null にして呼び出す → 最初のビューだけが計算する
    pathOf(obj) {
      if (!obj._d) obj._d = outline(obj.pts, obj.size, !!obj.pr);
      return obj._d;
    },
    invalidate(obj) { obj._d = null; obj._bb = null; },
    // 点 (x, y)・半径 r の消しゴムが線に触れているか
    hit(obj, x, y, r) {
      const bb = obj._bb || (obj._bb = bbox(obj));
      if (x < bb.x0 - r || x > bb.x1 + r || y < bb.y0 - r || y > bb.y1 + r) return false;
      const half = (obj.pr ? obj.size * 0.8 : obj.size) / 2;
      const lim = r + half;
      const p = obj.pts;
      if (p.length === 1) return Math.hypot(x - p[0][0], y - p[0][1]) <= lim;
      for (let i = 1; i < p.length; i++) {
        if (distSeg(x, y, p[i - 1][0], p[i - 1][1], p[i][0], p[i][1]) <= lim) return true;
      }
      return false;
    },
  };
})(window.B);
