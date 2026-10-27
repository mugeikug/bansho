'use strict';
// 編集画面のページ領域：選択・文字・ペン・消しゴム・図形の入力処理。
(function (B) {
  B.PEN_SIZES = [4, 7, 12];
  B.SHAPE_SIZES = [3, 6, 10];
  B.MARKER_SIZES = [22, 40];
  // 文字の大きさは pt で扱う（1pt ＝ 論理座標 2px。PowerPoint の 16:9 スライドと同じ基準）
  B.PT_STEPS = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96, 120];
  B.PT_MIN = 6;
  B.PT_MAX = 200;
  B.pxToPt = (px) => Math.round(px / 2 * 10) / 10;
  B.ERASER_SIZES = [10, 18, 32]; // 消しゴムの半径（論理座標）
  B.CLIP_PREFIX = 'BANSHO-CLIP:'; // 板書内のコピーをクリップボードに載せるときの目印
  B.INDENT_STEP = 1.5;            // 字下げ 1 段（em）
  B.BULLET_W = 1.0;               // 箇条書きの記号の幅（em）
  B.NUMBER_W = 1.6;               // 番号の幅（em）
  B.BULLETS = ['•', '–', '▪', '◦', '›'];   // 段ごとの記号
  B.SHAPES = ['line', 'arrow', 'rect', 'ellipse'];

  const r1 = (n) => Math.round(n * 10) / 10;
  function pressureOf(e) {
    if (e.pointerType !== 'pen') return 0.5;
    return Math.round(Math.max(0.05, e.pressure || 0) * 100) / 100;
  }
  const union = (a, b) => (a ? { x0: Math.min(a.x0, b.x0), y0: Math.min(a.y0, b.y0), x1: Math.max(a.x1, b.x1), y1: Math.max(a.y1, b.y1) } : b);

  class Editor {
    constructor(store, history, stage, state) {
      this.store = store;
      this.history = history;
      this.stage = stage;
      this.state = state;
      this.wrap = document.createElement('div');
      this.wrap.id = 'pagewrap';
      stage.appendChild(this.wrap);
      this.view = new B.View(store, document, this.wrap, { editor: true });
      this.selLayer = document.createElement('div');
      this.selLayer.className = 'sel-layer';
      this.wrap.appendChild(this.selLayer);
      this.cursor = document.createElement('div');
      this.cursor.className = 'eraser-cursor';
      this.wrap.appendChild(this.cursor);
      this.scale = 1;
      this.active = null;        // ドラッグ中の操作
      this.selection = new Set(); // 選択中のオブジェクト id
      this.ts = null;            // テキスト編集のセッション（文字単位の Undo を持つ）
      this.composing = false;
      this.clip = null;
      this.pasteCount = 0;
      this.editingId = null;
      this.onEditingChange = null;
      this.onSelectionChange = null;
      this.requestTool = null;   // ツールバー側でツールを切り替える関数

      new ResizeObserver(() => this.fit()).observe(stage);
      this.fit();
      this.bindPointer();
      this.bindText();
      store.on((e) => this.onStore(e));
      this.setTool(state.tool);
    }

    fit() {
      const W = this.stage.clientWidth, H = this.stage.clientHeight, pad = 16;
      const s = Math.max(0.05, Math.min((W - pad * 2) / B.W, (H - pad * 2) / B.H));
      this.scale = s;
      this.wrap.style.transform = `translate(${(W - B.W * s) / 2}px, ${(H - B.H * s) / 2}px) scale(${s})`;
      this.wrap.style.setProperty('--inv', String(1 / s));
    }

    toLogical(e) {
      const r = this.wrap.getBoundingClientRect();
      return { x: (e.clientX - r.left) / this.scale, y: (e.clientY - r.top) / this.scale };
    }

    setTool(tool) {
      this.state.tool = tool;
      this.wrap.className = 'tool-' + tool + (B.SHAPES.includes(tool) ? ' tool-shape' : '');
      if (tool !== 'text') this.endEditing();
      if (tool !== 'select') this.setSelection([]);
      this.cursor.style.display = 'none';
    }

    onStore(e) {
      if (e.type === 'current' || e.type === 'doc') {
        this.setSelection([]);
      } else if (this.selection.size && (e.type === 'obj-update' || e.type === 'obj-remove' || e.type === 'page-reset')) {
        if (e.type !== 'obj-update') {
          const ids = new Set(this.store.page.objects.map((o) => o.id));
          for (const id of Array.from(this.selection)) if (!ids.has(id)) this.selection.delete(id);
          if (this.onSelectionChange) this.onSelectionChange();
        }
        this.scheduleSel();
      }
    }

    // ---------------- ポインター入力 ----------------

    bindPointer() {
      const w = this.wrap;
      w.addEventListener('pointerdown', (e) => this.down(e));
      w.addEventListener('pointermove', (e) => this.move(e));
      w.addEventListener('pointerup', (e) => this.up(e));
      w.addEventListener('pointercancel', (e) => this.up(e));
      w.addEventListener('pointerleave', () => { if (!this.active) this.cursor.style.display = 'none'; });
      w.addEventListener('contextmenu', (e) => e.preventDefault());
      w.addEventListener('dblclick', (e) => this.dblclick(e));
    }

    down(e) {
      // パームリジェクション：タッチでは描かない（ペンとマウスのみ）
      if (e.pointerType === 'touch') { e.preventDefault(); return; }
      if (this.active) return;

      const t = e.target;
      const ch = t.closest && t.closest('.col-handle');
      if (ch) { this.colHandleDown(e, ch); return; }
      const handle = t.closest && t.closest('.handle');
      if (handle) { this.handleDown(e, handle); return; }
      const sh = t.closest && t.closest('.sel-handle');
      if (sh) { this.resizeDown(e, sh); return; }

      // ペンの消しゴム側（buttons=32）とサイドボタン（buttons=2）は消しゴム
      const penEraser = e.pointerType === 'pen' && ((e.buttons & 32) || (e.buttons & 2) || e.button === 5 || e.button === 2);
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      const tool = penEraser ? 'eraser' : this.state.tool;

      if (tool === 'text') { this.textDown(e); return; }

      e.preventDefault();
      this.endEditing();
      this.wrap.setPointerCapture(e.pointerId);
      const p = this.toLogical(e);
      const page = this.store.page;
      const base = { id: e.pointerId, tool, page, start: p };

      if (tool === 'select') { this.selectDown(e, p, base); return; }

      this.history.begin(page);
      if (tool === 'eraser') {
        this.active = Object.assign(base, { last: p });
        this.showEraser(p);
        this.eraseAlong(page, p, p);
      } else if (B.SHAPES.includes(tool)) {
        const obj = {
          id: B.uid('g'),
          type: 'shape',
          shape: tool,
          x1: r1(p.x), y1: r1(p.y), x2: r1(p.x), y2: r1(p.y),
          color: this.state.penColor,
          width: B.SHAPE_SIZES[this.state.penSize],
        };
        this.active = Object.assign(base, { obj });
        this.store.addObject(page, obj);
      } else {
        const marker = tool === 'marker';
        const st = this.state;
        const obj = {
          id: B.uid('s'),
          type: 'stroke',
          kind: marker ? 'marker' : 'pen',
          color: marker ? st.markerColor : st.penColor,
          size: marker ? B.MARKER_SIZES[st.markerSize] : B.PEN_SIZES[st.penSize],
          pr: !marker && e.pointerType === 'pen',
          pts: [[r1(p.x), r1(p.y), pressureOf(e)]],
        };
        this.active = Object.assign(base, { tool: marker ? 'marker' : 'pen', obj });
        this.store.addObject(page, obj);
      }
    }

    move(e) {
      const a = this.active;
      if (!a) {
        if (this.state.tool === 'eraser' && e.pointerType !== 'touch') this.showEraser(this.toLogical(e));
        return;
      }
      if (e.pointerId !== a.id) return;
      const events = e.getCoalescedEvents ? e.getCoalescedEvents() : [];
      const list = events.length ? events : [e];
      const p = this.toLogical(e);

      switch (a.tool) {
        case 'eraser':
          for (const ev of list) {
            const q = this.toLogical(ev);
            this.eraseAlong(a.page, a.last, q);
            a.last = q;
          }
          this.showEraser(a.last);
          break;
        case 'pen':
        case 'marker': {
          const pts = a.obj.pts;
          let added = false;
          for (const ev of list) {
            const q = this.toLogical(ev);
            const last = pts[pts.length - 1];
            if (Math.hypot(q.x - last[0], q.y - last[1]) < 1.2) continue;
            pts.push([r1(q.x), r1(q.y), pressureOf(ev)]);
            added = true;
          }
          if (added) {
            B.ink.invalidate(a.obj);
            this.store.updateObject(a.page, a.obj, { live: true });
          }
          break;
        }
        case 'move': {
          const dx = p.x - a.start.x, dy = p.y - a.start.y;
          if (!a.moved && Math.hypot(dx, dy) * this.scale < 3) return;
          a.moved = true;
          for (const o of a.origs) this.store.replaceObject(a.page, B.geom.transform(o, { dx, dy }));
          break;
        }
        case 'resize': {
          const vx = a.c0.x - a.anchor.x, vy = a.c0.y - a.anchor.y;
          const s = Math.max(0.1, ((p.x - a.anchor.x) * vx + (p.y - a.anchor.y) * vy) / (vx * vx + vy * vy || 1));
          for (const o of a.origs) this.store.replaceObject(a.page, B.geom.transform(o, { s, ox: a.anchor.x, oy: a.anchor.y }));
          break;
        }
        case 'marquee':
          this.drawMarquee(a.start, p);
          break;
        default: { // 図形
          let x2 = p.x, y2 = p.y;
          if (e.shiftKey) ({ x2, y2 } = B.geom.constrain(a.obj.shape, a.obj.x1, a.obj.y1, x2, y2));
          a.obj.x2 = r1(x2);
          a.obj.y2 = r1(y2);
          this.store.updateObject(a.page, a.obj, { live: true });
        }
      }
    }

    up(e) {
      const a = this.active;
      if (!a || e.pointerId !== a.id) return;
      this.active = null;
      switch (a.tool) {
        case 'eraser':
          if (this.state.tool !== 'eraser' || e.pointerType === 'touch') this.cursor.style.display = 'none';
          this.history.commit();
          break;
        case 'pen':
        case 'marker':
          B.ink.invalidate(a.obj);
          this.store.updateObject(a.page, a.obj, {});
          this.history.commit();
          break;
        case 'move':
          // 複数選択中に1つをクリックしただけなら、それだけを選択し直す
          if (!a.moved && !a.shift && this.selection.size > 1) this.setSelection([a.clickId]);
          this.history.commit();
          break;
        case 'resize':
          this.history.commit();
          break;
        case 'marquee':
          this.finishMarquee(a, this.toLogical(e));
          break;
        default: {
          const o = a.obj;
          if (Math.hypot(o.x2 - o.x1, o.y2 - o.y1) < 5) this.store.removeObject(a.page, o.id);
          this.history.commit();
        }
      }
    }

    eraserR() { return B.ERASER_SIZES[this.state.eraserSize] || B.ERASER_SIZES[1]; }

    showEraser(p) {
      const r = this.eraserR();
      const c = this.cursor.style;
      c.display = 'block';
      c.left = (p.x - r) + 'px';
      c.top = (p.y - r) + 'px';
      c.width = c.height = (r * 2) + 'px';
      this.cursor.classList.toggle('partial', this.state.eraserMode === 'partial');
    }

    // 前回位置から今回位置までを半径ぶんの間隔でなぞる。
    // 「線ごと」なら触れた線・図形を丸ごと消し、「部分」なら手書きの線の触れた部分だけを消す
    // （図形は部分的に消せないので丸ごと消す）。
    eraseAlong(page, from, to) {
      const r = this.eraserR();
      const partial = this.state.eraserMode === 'partial';
      const dist = Math.hypot(to.x - from.x, to.y - from.y);
      const steps = Math.max(1, Math.ceil(dist / (r * (partial ? 0.4 : 0.7))));
      for (let i = 0; i <= steps; i++) {
        const t = i / steps;
        const x = from.x + (to.x - from.x) * t, y = from.y + (to.y - from.y) * t;
        const hits = page.objects.filter((o) => (o.type === 'stroke' || o.type === 'shape') && B.geom.hitErase(o, x, y, r));
        for (const o of hits) {
          if (partial && o.type === 'stroke') {
            const pieces = B.ink.split(o, x, y, r).map((pts) => Object.assign(B.cloneObj(o), { id: B.uid('s'), pts }));
            this.store.replaceWith(page, o.id, pieces);
          } else {
            this.store.removeObject(page, o.id);
          }
        }
      }
    }

    // 画像を挿入する（at を中心に。省略時はページ中央）。大きい画像はページの 7 割に収める。
    insertImage(img, at) {
      const k = Math.min(1, (B.W * 0.7) / img.w, (B.H * 0.7) / img.h);
      const w = Math.round(img.w * k), h = Math.round(img.h * k);
      const cx = at ? at.x : B.W / 2, cy = at ? at.y : B.H / 2;
      const obj = {
        id: B.uid('i'),
        type: 'image',
        x: Math.round(Math.max(0, Math.min(B.W - w, cx - w / 2))),
        y: Math.round(Math.max(0, Math.min(B.H - h, cy - h / 2))),
        w, h,
        src: img.src,
      };
      this.endEditing();
      const page = this.store.page;
      this.history.begin(page);
      this.store.addObject(page, obj);
      this.history.commit();
      if (this.requestTool) this.requestTool('select');
      this.setSelection([obj.id]);
    }

    // 文字列を新しいテキストボックスとして置く（クリップボードの文字の貼り付け）
    insertText(text) {
      const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const lines = text.replace(/\r\n?/g, '\n').split('\n');
      const html = B.rich.canonicalHtml(lines.map(esc).join('<br>'));
      if (B.rich.isEmpty(html)) return;
      const obj = { id: B.uid('t'), type: 'text', x: 160, y: 160, w: null, fs: this.state.textPt * 2, html };
      this.endEditing();
      const page = this.store.page;
      this.history.begin(page);
      this.store.addObject(page, obj);
      this.history.commit();
      if (this.requestTool) this.requestTool('select');
      this.setSelection([obj.id]);
    }

    // ---------------- 選択 ----------------

    textRect(o) {
      const el = this.view.elOf(o.id);
      if (!el) return { x0: o.x, y0: o.y, x1: o.x + 10, y1: o.y + 10 };
      return { x0: el.offsetLeft, y0: el.offsetTop, x1: el.offsetLeft + el.offsetWidth, y1: el.offsetTop + el.offsetHeight };
    }

    hitAt(p) {
      const tol = 8 / this.scale;
      const objs = this.store.page.objects;
      const tr = (o) => this.textRect(o);
      for (let i = objs.length - 1; i >= 0; i--) {
        if (B.geom.hitSelect(objs[i], p.x, p.y, tol, tr)) return objs[i];
      }
      return null;
    }

    selectedObjs() {
      return this.store.page.objects.filter((o) => this.selection.has(o.id));
    }

    setSelection(ids) {
      if (!ids.length && !this.selection.size) return;
      this.selection = new Set(ids);
      this.renderSel();
      if (this.onSelectionChange) this.onSelectionChange();
    }

    selectDown(e, p, base) {
      const hit = this.hitAt(p);
      if (hit) {
        if (e.shiftKey) {
          if (this.selection.has(hit.id)) {
            this.selection.delete(hit.id);
            this.setSelection(Array.from(this.selection));
            return;
          }
          this.setSelection([...this.selection, hit.id]);
        } else if (!this.selection.has(hit.id)) {
          this.setSelection([hit.id]);
        }
        this.history.begin(base.page);
        this.active = Object.assign(base, { tool: 'move', origs: this.selectedObjs(), moved: false, clickId: hit.id, shift: e.shiftKey });
      } else {
        if (!e.shiftKey) this.setSelection([]);
        this.active = Object.assign(base, { tool: 'marquee', baseSel: Array.from(this.selection) });
      }
    }

    drawMarquee(a, b) {
      let m = this.selLayer.querySelector('.marquee');
      if (!m) {
        m = document.createElement('div');
        m.className = 'marquee';
        this.selLayer.appendChild(m);
      }
      Object.assign(m.style, {
        left: Math.min(a.x, b.x) + 'px', top: Math.min(a.y, b.y) + 'px',
        width: Math.abs(b.x - a.x) + 'px', height: Math.abs(b.y - a.y) + 'px',
      });
    }

    finishMarquee(a, p) {
      const r = { x0: Math.min(a.start.x, p.x), y0: Math.min(a.start.y, p.y), x1: Math.max(a.start.x, p.x), y1: Math.max(a.start.y, p.y) };
      const ids = new Set(a.baseSel);
      if ((r.x1 - r.x0) * this.scale > 3 || (r.y1 - r.y0) * this.scale > 3) {
        for (const o of this.store.page.objects) {
          const b = B.geom.bbox(o, (t) => this.textRect(t));
          if (b.x0 >= r.x0 && b.x1 <= r.x1 && b.y0 >= r.y0 && b.y1 <= r.y1) ids.add(o.id);
        }
      }
      this.selection = new Set(); // 強制的に描き直す
      this.setSelection(Array.from(ids));
      this.renderSel();
    }

    groupBox() {
      let g = null;
      for (const o of this.selectedObjs()) g = union(g, B.geom.bbox(o, (t) => this.textRect(t)));
      return g;
    }

    scheduleSel() {
      if (this.selPending) return;
      this.selPending = true;
      requestAnimationFrame(() => { this.selPending = false; this.renderSel(); });
    }

    renderSel() {
      const L = this.selLayer;
      L.textContent = '';
      const objs = this.selectedObjs();
      if (!objs.length) return;
      const box = (b, cls) => {
        const d = document.createElement('div');
        d.className = cls;
        Object.assign(d.style, { left: b.x0 + 'px', top: b.y0 + 'px', width: (b.x1 - b.x0) + 'px', height: (b.y1 - b.y0) + 'px' });
        L.appendChild(d);
        return d;
      };
      let g = null;
      for (const o of objs) {
        const b = B.geom.bbox(o, (t) => this.textRect(t));
        if (objs.length > 1) box(b, 'sel-box');
        g = union(g, b);
      }
      const gb = box(g, 'sel-group');
      for (const c of ['nw', 'ne', 'sw', 'se']) {
        const h = document.createElement('div');
        h.className = 'sel-handle ' + c;
        h.dataset.corner = c;
        h.title = 'ドラッグで拡大・縮小';
        gb.appendChild(h);
      }
    }

    resizeDown(e, h) {
      e.preventDefault();
      e.stopPropagation();
      const g = this.groupBox();
      if (!g) return;
      const c = h.dataset.corner;
      const page = this.store.page;
      this.wrap.setPointerCapture(e.pointerId);
      this.history.begin(page);
      this.active = {
        id: e.pointerId, tool: 'resize', page,
        anchor: { x: c.includes('w') ? g.x1 : g.x0, y: c.includes('n') ? g.y1 : g.y0 },
        c0: { x: c.includes('w') ? g.x0 : g.x1, y: c.includes('n') ? g.y0 : g.y1 },
        origs: this.selectedObjs(),
      };
    }

    // 選択中のオブジェクトを変更する共通処理（複製して差し替え、Undo に1回分として積む）
    modifySelection(fn) {
      const page = this.store.page;
      const objs = this.selectedObjs();
      if (!objs.length) return false;
      this.history.begin(page);
      for (const o of objs) {
        const c = fn(o);
        if (c) this.store.replaceObject(page, c);
      }
      this.history.commit();
      return true;
    }

    deleteSelection() {
      const page = this.store.page;
      const ids = Array.from(this.selection);
      if (!ids.length) return;
      this.history.begin(page);
      for (const id of ids) this.store.removeObject(page, id);
      this.history.commit();
      this.setSelection([]);
    }

    copySelection() {
      const objs = this.selectedObjs();
      if (!objs.length) return false;
      this.clip = objs.map(B.cloneObj);
      this.pasteCount = 0;
      return true;
    }

    paste() {
      if (!this.clip || !this.clip.length) return false;
      this.pasteCount++;
      const d = 30 * this.pasteCount;
      const page = this.store.page;
      this.history.begin(page);
      const ids = [];
      for (const o of this.clip) {
        const c = B.geom.transform(o, { dx: d, dy: d });
        c.id = B.uid(o.type[0]);
        this.store.addObject(page, c);
        ids.push(c.id);
      }
      this.history.commit();
      this.setSelection(ids);
      return true;
    }

    selectAll() { this.setSelection(this.store.page.objects.map((o) => o.id)); }

    nudge(dx, dy) { this.modifySelection((o) => B.geom.transform(o, { dx, dy })); }

    // 選択中の線・図形の色や太さを変える
    restyleSelection(color, sizeIdx) {
      return this.modifySelection((o) => {
        if (o.type === 'stroke' && o.kind === 'pen') {
          const c = B.cloneObj(o);
          if (color) c.color = color;
          if (sizeIdx != null) c.size = B.PEN_SIZES[sizeIdx];
          return c;
        }
        if (o.type === 'shape') {
          const c = B.cloneObj(o);
          if (color) c.color = color;
          if (sizeIdx != null) c.width = B.SHAPE_SIZES[sizeIdx];
          return c;
        }
        return null;
      });
    }

    // 文字を持つオブジェクト（テキストボックスと表）
    selectedTexts() { return this.selectedObjs().filter((o) => o.type === 'text' || o.type === 'table'); }

    // 選択中のテキストボックス全体に装飾をかける／外す
    formatSelectedTexts(key, value) {
      const texts = this.selectedTexts();
      if (!texts.length) return false;
      const htmlsOf = (o) => (o.type === 'table' ? o.cells.flat().map((c) => c.html) : [o.html]);
      const flats = texts.flatMap((o) => htmlsOf(o).map((h) => B.rich.parseHtml(h)));
      let nv = key === 'c' ? (value === 'default' ? null : value) : value;
      if (key !== 'c' && flats.every((f) => !f.length || B.rich.common(f, 0, f.length)[key] === value)) nv = null;
      const conv = (h) => { const f = B.rich.parseHtml(h); B.rich.toggle(f, 0, f.length, key, nv, true); return B.rich.render(f); };
      const ids = new Set(texts.map((o) => o.id));
      return this.modifySelection((o) => {
        if (!ids.has(o.id)) return null;
        const c = B.cloneObj(o);
        if (c.type === 'table') for (const row of c.cells) for (const cell of row) cell.html = conv(cell.html);
        else c.html = conv(c.html);
        return c;
      });
    }

    dblclick(e) {
      if (this.state.tool !== 'select') return;
      const hit = this.hitAt(this.toLogical(e));
      if (!hit || (hit.type !== 'text' && hit.type !== 'table')) return;
      if (this.requestTool) this.requestTool('text');
      const at = document.elementFromPoint(e.clientX, e.clientY);
      const c = (at && at.closest && at.closest('.content')) || this.view.elOf(hit.id).querySelector('.content');
      c.focus();
      const r = document.caretRangeFromPoint && document.caretRangeFromPoint(e.clientX, e.clientY);
      if (r && c.contains(r.startContainer)) {
        const sel = getSelection();
        sel.removeAllRanges();
        sel.addRange(r);
      }
    }

    // ---------------- テキスト ----------------

    textDown(e) {
      // 既存のテキストをクリック → ブラウザ標準の動作でカーソルを置く
      if (e.target.closest && e.target.closest('.obj .content')) return;
      if (e.target.closest && e.target.closest('.col-handle')) return;
      e.preventDefault(); // 今の編集中ボックスからフォーカスを移さない（下で新しいボックスに移す）
      this.endEditing();
      const p = this.toLogical(e);
      const fs = this.state.textPt * 2;
      const obj = {
        id: B.uid('t'),
        type: 'text',
        x: Math.round(p.x),
        y: Math.round(p.y - fs * 0.7),
        w: null,
        fs,
        html: '',
      };
      const page = this.store.page;
      this.history.begin(page);
      this.store.addObject(page, obj);
      const el = this.view.elOf(obj.id);
      if (el) el.querySelector('.content').focus();
    }

    // 文字の入力欄（テキストボックスの中身、または表のセル）
    contentOf(node) {
      const c = node && node.closest && node.closest('.content');
      if (c) return this.view.root.contains(c) ? c : null;
      const box = node && node.closest && node.closest('.text');
      return box ? box.querySelector('.content') : null;
    }

    objOf(node) {
      const box = node && node.closest && node.closest('.obj');
      if (!box) return null;
      return this.store.findObject(this.store.page, box.dataset.id);
    }

    cellOf(content) {
      const td = content && content.closest('td');
      return td ? [Number(td.dataset.r), Number(td.dataset.c)] : null;
    }

    bindText() {
      const root = this.view.root;
      root.addEventListener('input', (e) => {
        const c = this.contentOf(e.target);
        const obj = this.objOf(e.target);
        if (!c || !obj || !this.ts || this.ts.content !== c) return;
        this.setEditingHtml(B.rich.canonical(c));
        if (!e.isComposing) this.scheduleSnap();
      });
      root.addEventListener('compositionstart', () => { this.composing = true; });
      root.addEventListener('compositionend', () => { this.composing = false; this.scheduleSnap(); });
      root.addEventListener('beforeinput', (e) => {
        if (!this.contentOf(e.target)) return;
        // 元に戻す・書式はアプリ側で扱う（ブラウザ標準の処理だと装飾の整合が崩れる）
        if (e.inputType === 'historyUndo') { e.preventDefault(); this.textUndo(); }
        else if (e.inputType === 'historyRedo') { e.preventDefault(); this.textRedo(); }
        else if (e.inputType.startsWith('format')) e.preventDefault();
      });
      root.addEventListener('focusin', (e) => {
        const c = this.contentOf(e.target);
        const obj = this.objOf(e.target);
        if (c && obj) this.startTextSession(obj, c);
      });
      root.addEventListener('focusout', (e) => {
        // ウィンドウ自体がフォーカスを失っただけ（提示画面をクリックした等）なら編集継続
        if (document.activeElement === e.target) return;
        // 文字サイズの入力欄へ移っただけなら編集を続ける（入力後に戻ってくる）
        if (e.relatedTarget && e.relatedTarget.closest && e.relatedTarget.closest('.keep-text-focus')) return;
        this.endTextSession();
      });
      // 貼り付けは書式なしのテキストのみ
      root.addEventListener('paste', (e) => {
        if (!this.contentOf(e.target)) return;
        e.preventDefault();
        const text = e.clipboardData.getData('text/plain');
        // 板書内でコピーした図形などのデータは文字として貼らない
        if (text && !text.startsWith(B.CLIP_PREFIX)) document.execCommand('insertText', false, text);
      });
      root.addEventListener('drop', (e) => { if (this.contentOf(e.target)) e.preventDefault(); });
    }

    startTextSession(obj, content) {
      if (this.ts && this.ts.id === obj.id && this.ts.content === content) return;
      if (this.ts) this.endTextSession();
      const page = this.store.page;
      this.history.begin(page);
      // 確定済みのオブジェクトは書き換えない：複製に差し替えてから編集する
      this.store.replaceObject(page, B.cloneObj(obj), { source: 'editor' });
      this.ts = { id: obj.id, pageId: page.id, orig: obj, content, cell: obj.type === 'table' ? this.cellOf(content) : null, stack: [], idx: -1, timer: 0 };
      this.pushSnap();
      this.editingId = obj.id;
      if (this.onEditingChange) this.onEditingChange();
    }

    endTextSession() {
      const ts = this.ts;
      if (!ts) return;
      this.ts = null;
      clearTimeout(ts.timer);
      const page = this.store.pageById(ts.pageId);
      const obj = page && this.store.findObject(page, ts.id);
      if (obj && ts.cell) {
        // 表のセル：中身を決まった形に整えるだけ（空でも表は消さない）
        const c = ts.content;
        const cell = obj.cells[ts.cell[0]] && obj.cells[ts.cell[0]][ts.cell[1]];
        if (cell && c.isConnected) {
          const html = B.rich.isEmpty(B.rich.canonical(c)) ? '' : B.rich.canonical(c);
          if (c.innerHTML !== html) c.innerHTML = html;
          if (cell.html !== html) {
            cell.html = html;
            this.store.updateObject(page, obj, { source: 'editor', cell: ts.cell });
          }
        }
      } else if (obj) {
        const el = this.view.elOf(obj.id);
        const c = el && el.querySelector('.content');
        const html = c ? B.rich.canonical(c) : obj.html;
        if (B.rich.isEmpty(html)) {
          this.store.removeObject(page, obj.id);
        } else {
          if (c && c.innerHTML !== html) c.innerHTML = html;
          if (obj.html !== html) {
            obj.html = html;
            this.store.updateObject(page, obj, { source: 'editor' });
          }
        }
      }
      // 何も変えずに編集を終えたときは、編集用の複製を元のオブジェクトに戻す（空の「元に戻す」を作らない）
      if (page && ts.orig) {
        const i = page.objects.findIndex((o) => o.id === ts.id);
        if (i >= 0 && JSON.stringify(B.cloneObj(page.objects[i])) === JSON.stringify(B.cloneObj(ts.orig))) page.objects[i] = ts.orig;
      }
      this.history.commit();
      this.editingId = null;
      this.composing = false;
      if (this.onEditingChange) this.onEditingChange();
    }

    endEditing() {
      const a = document.activeElement;
      if (a && a.isContentEditable && this.view.root.contains(a)) a.blur();
      this.endTextSession();
    }

    editingObj() {
      if (!this.ts) return null;
      const page = this.store.pageById(this.ts.pageId);
      return page ? this.store.findObject(page, this.ts.id) : null;
    }

    // ---- テキスト編集中の Undo（文字単位） ----

    selPoints(c) {
      const sel = getSelection();
      if (!sel.rangeCount || !c.contains(sel.anchorNode)) return [];
      const r = sel.getRangeAt(0);
      return [{ node: r.startContainer, offset: r.startOffset }, { node: r.endContainer, offset: r.endOffset }];
    }

    capture() {
      const c = this.ts.content;
      const pts = this.selPoints(c);
      const { flat, offs } = B.rich.parse(c, pts);
      return { html: B.rich.render(flat), sel: pts.length ? offs : null };
    }

    pushSnap() {
      const ts = this.ts;
      if (!ts) return;
      clearTimeout(ts.timer);
      ts.timer = 0;
      const snap = this.capture();
      if (ts.idx >= 0 && ts.stack[ts.idx].html === snap.html) { ts.stack[ts.idx].sel = snap.sel; return; }
      ts.stack = ts.stack.slice(0, ts.idx + 1);
      ts.stack.push(snap);
      ts.idx = ts.stack.length - 1;
    }

    scheduleSnap() {
      const ts = this.ts;
      if (!ts) return;
      clearTimeout(ts.timer);
      ts.timer = setTimeout(() => this.pushSnap(), 600);
    }

    textCanUndo() {
      const ts = this.ts;
      if (!ts) return false;
      if (ts.timer) this.pushSnap();
      return ts.idx > 0;
    }

    textUndo() {
      const ts = this.ts;
      if (!ts || this.composing) return;
      if (ts.timer) this.pushSnap();
      if (ts.idx <= 0) return;
      ts.idx--;
      this.restoreSnap(ts.stack[ts.idx]);
    }

    textRedo() {
      const ts = this.ts;
      if (!ts || this.composing) return;
      if (ts.timer) this.pushSnap();
      if (ts.idx >= ts.stack.length - 1) return;
      ts.idx++;
      this.restoreSnap(ts.stack[ts.idx]);
    }

    restoreSnap(snap) {
      const c = this.ts.content;
      c.innerHTML = snap.html;
      const n = B.rich.parse(c).flat.length;
      const [s, e] = snap.sel || [n, n];
      this.setSel(c, s, e);
      this.setEditingHtml(snap.html);
    }

    // 文字位置で選択範囲を設定する。pointAt は決まった形の HTML（行は <br>）を前提にするので、
    // ブラウザが Enter で作った <div> の行などが残っていれば先に整える（IME 変換中には呼ばないこと）。
    setSel(c, s, e) {
      const html = B.rich.canonical(c);
      if (c.innerHTML !== html) c.innerHTML = html;
      const a = B.rich.pointAt(c, s), b = B.rich.pointAt(c, e);
      const r = document.createRange();
      r.setStart(a.node, a.offset);
      r.setEnd(b.node, b.offset);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
    }

    // ---- 文字の装飾 ----
    // key: x 囲み / h マーカー / c 文字色 / b 太字 / i 斜体 / u 下線 / s 上付き・下付き
    formatText(key, value) {
      const ts = this.ts;
      if (!ts || this.composing) return false;
      const c = ts.content;
      const pts = this.selPoints(c);
      if (!pts.length) return false;
      const { flat, offs } = B.rich.parse(c, pts);
      const s = Math.min(offs[0], offs[1]), e = Math.max(offs[0], offs[1]);
      if (!flat.slice(s, e).some((ch) => ch.ch !== '\n')) {
        // 文字を選んでいないとき：これから打つ文字の書式にする
        B.rich.typingStyle(key, value, B.rich.common(flat, s, e));
        return true;
      }
      if (ts.timer) this.pushSnap();
      if (key === 'c') B.rich.toggle(flat, s, e, 'c', value === 'default' ? null : value, true);
      else if (key === 'z') B.rich.toggle(flat, s, e, 'z', value, true);
      else B.rich.toggle(flat, s, e, key, value);
      const html = B.rich.render(flat);
      c.innerHTML = html;
      this.setSel(c, s, e);
      this.setEditingHtml(html);
      this.pushSnap();
      return true;
    }

    // ツールバーの表示用：選択範囲の書式
    formatState() {
      const ts = this.ts;
      if (!ts) {
        const texts = this.selectedTexts();
        if (!texts.length) return null;
        const f = B.rich.parseHtml(texts[0].html);
        return B.rich.common(f, 0, f.length);
      }
      const c = ts.content;
      const pts = this.selPoints(c);
      if (!pts.length) return null;
      const { flat, offs } = B.rich.parse(c, pts);
      const s = Math.min(offs[0], offs[1]), e = Math.max(offs[0], offs[1]);
      const st = B.rich.common(flat, s, e);
      if (s === e) {
        // カーソルだけのときは「次に打つ文字の書式」も反映する
        const q = (cmd) => { try { return document.queryCommandState(cmd); } catch (err) { return false; } };
        st.b = q('bold') || null;
        st.i = q('italic') || null;
        st.u = q('underline') || null;
        st.k = q('strikeThrough') || null;
        st.s = q('subscript') ? 'sub' : q('superscript') ? 'sup' : null;
      }
      return st;
    }

    // ---- 箇条書き・字下げ・タブ ----
    // 字下げは 1 段 1.5em。箇条書きの記号は行の書式（bu）で、記号の幅（bw）だけぶら下げる。

    // 選択範囲が含む行に fn(行の書式) → 新しい行の書式 を適用して描き直す
    changeLines(fn) {
      const ts = this.ts;
      if (!ts || this.composing) return false;
      const c = ts.content;
      const pts = this.selPoints(c);
      if (!pts.length) return false;
      const { flat, offs } = B.rich.parse(c, pts);
      const s = Math.min(offs[0], offs[1]), e = Math.max(offs[0], offs[1]);
      const lines = B.rich.lineList(flat).filter((ln) => ln.end >= s && ln.start <= e);
      if (ts.timer) this.pushSnap();
      let changed = false;
      for (const ln of lines) {
        const next = fn(B.rich.lineAttrsOf(flat, ln), lines, flat);
        if (next) { B.rich.setLineAttrs(flat, ln, next); changed = true; }
      }
      if (!changed) return false;
      const html = B.rich.render(flat);
      c.innerHTML = html;
      this.setSel(c, s, e);
      this.setEditingHtml(html);
      this.pushSnap();
      return true;
    }

    // dir: +1 で 1 段深く、-1 で浅く
    // 行の段（0 から）。箇条書き・番号の記号の幅を除いた左の余白から求める
    lineLevel(a) {
      if (a.lv != null) return a.lv;
      const bw = a.bu || a.ol ? a.bw || B.BULLET_W : 0;
      return Math.max(0, Math.round(((a.pl || 0) - bw) / B.INDENT_STEP));
    }

    // 段 level・記号の幅 bw のときの行の書式
    levelAttrs(level, bw) {
      const r = { pl: Math.round((level * B.INDENT_STEP + bw) * 1000) / 1000 || null, lv: level };
      r.ti = bw ? -bw : null;
      return r;
    }

    // dir: +1 で 1 段深く、-1 で浅く
    indentLines(dir) {
      return this.changeLines((a) => {
        const list = a.bu || a.ol;
        const bw = list ? a.bw || B.BULLET_W : 0;
        const level = this.lineLevel(a);
        const nl = Math.max(0, Math.min(4, level + dir));
        if (nl === level && !(dir < 0 && (a.pl || 0) - bw > 0.01)) return null;
        const r = this.levelAttrs(nl, bw);
        if (!list) { r.ti = a.ti || null; if (!nl) r.lv = null; }
        if (a.bu && B.BULLETS.includes(a.bu)) r.bu = B.BULLETS[nl % B.BULLETS.length];
        return r;
      });
    }

    // 箇条書きのオン・オフ（選んだ行がすべて箇条書きなら外す）
    toggleBullets() {
      let allOn = null;
      return this.changeLines((a, lines, flat) => {
        if (allOn === null) allOn = lines.every((ln) => B.rich.lineAttrsOf(flat, ln).bu);
        if (allOn) return this.listOff(a);
        if (a.bu) return null;
        const level = this.lineLevel(a);
        return Object.assign({ bu: B.BULLETS[level % B.BULLETS.length], ol: null, bc: null, bs: null, bw: B.BULLET_W }, this.levelAttrs(level, B.BULLET_W));
      });
    }

    // 番号付きリストのオン・オフ。fmt: '1.' 'a)' 'i.' など（選んだ行がすべてその形式なら外す）
    toggleNumbering(fmt) {
      let allOn = null;
      return this.changeLines((a, lines, flat) => {
        if (allOn === null) allOn = lines.every((ln) => B.rich.lineAttrsOf(flat, ln).ol === fmt);
        if (allOn) return this.listOff(a);
        const level = this.lineLevel(a);
        return Object.assign({ ol: fmt, bu: null, bc: null, bs: null, bw: B.NUMBER_W }, this.levelAttrs(level, B.NUMBER_W));
      });
    }

    listOff(a) {
      const level = this.lineLevel(a);
      return { bu: null, ol: null, bw: null, bc: null, bs: null, ti: null, pl: Math.round(level * B.INDENT_STEP * 1000) / 1000 || null, lv: level || null };
    }

    // 行の揃え（'left' | 'center' | 'right'）。文字入力中は選んだ行、選択ツールではボックス全体
    alignLines(ta) {
      if (this.ts) return this.changeLines(() => ({ ta }));
      return this.formatSelectedLines(() => ({ ta }));
    }

    // 行の書式（段落の前の間隔など）を、文字入力中は選んだ行に、選択ツールではボックス全体に
    setLineAttrs(attrs) {
      if (this.ts) return this.changeLines(() => attrs);
      return this.formatSelectedLines(() => attrs);
    }

    // 番号付きリストの開始番号：選んだ範囲の最初の番号付きの行に設定する
    setOlStart(n) {
      let done = false;
      return this.changeLines((a) => {
        if (done || !a.ol) return null;
        done = true;
        return { os: n === 1 ? null : n };
      });
    }

    // 選択中のテキストボックス・表のすべての行に行の書式を適用する
    formatSelectedLines(fn) {
      const conv = (html) => {
        const flat = B.rich.parseHtml(html);
        for (const ln of B.rich.lineList(flat)) {
          const next = fn(B.rich.lineAttrsOf(flat, ln));
          if (next) B.rich.setLineAttrs(flat, ln, next);
        }
        return B.rich.render(flat);
      };
      const ids = new Set(this.selectedTexts().map((o) => o.id));
      if (!ids.size) return false;
      return this.modifySelection((o) => {
        if (!ids.has(o.id)) return null;
        const c = B.cloneObj(o);
        if (c.type === 'table') for (const row of c.cells) for (const cell of row) cell.html = conv(cell.html);
        else { c.html = conv(c.html); delete c.align; }
        return c;
      });
    }

    // ツールバー表示用：選択範囲の最初の行の書式
    lineState() {
      const ts = this.ts;
      if (!ts) {
        const t = this.selectedTexts()[0];
        if (!t || t.type !== 'text') return null;
        const flat = B.rich.parseHtml(t.html);
        return Object.assign({ ta: t.align || null }, B.rich.lineAttrsOf(flat, B.rich.lineList(flat)[0]));
      }
      const pts = this.selPoints(ts.content);
      if (!pts.length) return null;
      const { flat, offs } = B.rich.parse(ts.content, pts);
      const s0 = Math.min(offs[0], offs[1]);
      const ln = B.rich.lineList(flat).find((l) => l.end >= s0) || B.rich.lineList(flat)[0];
      const obj = this.editingObj();
      return Object.assign({ ta: obj && obj.align || null }, B.rich.lineAttrsOf(flat, ln));
    }

    // Tab：行頭（箇条書き・字下げのある行）や複数行の選択では字下げ、それ以外はタブ文字
    handleTab(shift) {
      const ts = this.ts;
      if (!ts || this.composing) return;
      if (ts.cell) { this.moveCell(shift ? -1 : 1); return; }
      const pts = this.selPoints(ts.content);
      if (!pts.length) return;
      const { flat, offs } = B.rich.parse(ts.content, pts);
      const s = Math.min(offs[0], offs[1]), e = Math.max(offs[0], offs[1]);
      const lines = B.rich.lineList(flat).filter((ln) => ln.end >= s && ln.start <= e);
      const ln = lines[0];
      const a = B.rich.lineAttrsOf(flat, ln);
      const atStart = s === ln.start;
      if (shift) { this.indentLines(-1); return; }
      if (lines.length > 1 || (atStart && (a.bu || a.ol || (a.pl || 0) > 0))) { this.indentLines(1); return; }
      document.execCommand('insertText', false, '\t');
    }

    // 編集中の文字（テキストボックス、または表のセル）の HTML を更新して他の画面に伝える
    setEditingHtml(html) {
      const obj = this.editingObj();
      if (!obj) return;
      if (this.ts.cell) {
        const [r, c] = this.ts.cell;
        obj.cells[r][c].html = html;
        this.store.updateObject(this.store.page, obj, { source: 'editor', cell: this.ts.cell });
      } else {
        obj.html = html;
        this.store.updateObject(this.store.page, obj, { source: 'editor' });
      }
    }

    // ---- 文字の大きさ（pt） ----
    // pt は PowerPoint の 16:9 スライドと同じ基準（論理座標 1920px 幅 ＝ 960pt）なので 1pt ＝ 2px。
    // 範囲を選んでいればその文字だけ（ボックスの基準サイズに対する倍率で持つ）、
    // 選んでいなければボックス全体の基準サイズを変える。

    selRange() {
      const ts = this.ts;
      if (!ts) return null;
      const pts = this.selPoints(ts.content);
      if (!pts.length) return null;
      const { flat, offs } = B.rich.parse(ts.content, pts);
      const s = Math.min(offs[0], offs[1]), e = Math.max(offs[0], offs[1]);
      return { flat, s, e, empty: !flat.slice(s, e).some((ch) => ch.ch !== '\n') };
    }

    // ツールバーに表示する大きさ（pt）。範囲内で大きさがばらばらなら null。
    fontPt() {
      const obj = this.editingObj();
      if (obj) {
        const r = this.selRange();
        if (!r || r.empty) return B.pxToPt(obj.fs);
        const z = B.rich.common(r.flat, r.s, r.e).z;
        return z === undefined ? null : B.pxToPt(obj.fs * (z || 1));
      }
      const t = this.selectedTexts();
      if (t.length) return t.every((o) => o.fs === t[0].fs) ? B.pxToPt(t[0].fs) : null;
      return this.state.textPt;
    }

    setFontPt(pt) {
      pt = Math.max(B.PT_MIN, Math.min(B.PT_MAX, Math.round(pt * 2) / 2));
      const obj = this.editingObj();
      if (obj) {
        const r = this.selRange();
        if (r && !r.empty) {
          this.formatText('z', B.rich.ratio(pt * 2 / obj.fs));
        } else {
          obj.fs = pt * 2;
          this.store.updateObject(this.store.page, obj, { source: 'editor' });
          this.state.textPt = pt;
        }
        return;
      }
      this.state.textPt = pt;
      this.modifySelection((o) => {
        if (o.type !== 'text') return null;
        const c = B.cloneObj(o);
        c.fs = pt * 2;
        return c;
      });
    }

    // 1段階大きく（dir=1）／小さく（dir=-1）
    stepFontPt(dir) {
      const cur = this.fontPt() || this.state.textPt;
      const steps = B.PT_STEPS;
      let next;
      if (dir > 0) next = steps.find((v) => v > cur + 0.01) || Math.min(B.PT_MAX, cur + 8);
      else next = steps.slice().reverse().find((v) => v < cur - 0.01) || Math.max(B.PT_MIN, cur - 2);
      this.setFontPt(next);
    }

    // サイズ入力欄にフォーカスを移す前に、編集中の選択範囲を覚えておく
    rememberRange() {
      const r = this.selRange();
      this.savedRange = r ? { id: this.ts.id, s: r.s, e: r.e } : null;
    }

    // 覚えておいた選択範囲を編集中のボックスに戻す
    restoreRange() {
      const ts = this.ts, r = this.savedRange;
      if (!ts) return false;
      ts.content.focus();
      if (r && r.id === ts.id) this.setSel(ts.content, r.s, r.e);
      return true;
    }

    // ---------------- 表 ----------------

    // 新しい表を置いて、左上のセルを編集状態にする
    insertTable(rows, cols) {
      this.endEditing();
      const fs = this.state.textPt * 2;
      const cw = Math.max(120, Math.min(320, Math.floor(1500 / cols)));
      const rh = Math.round(fs * 1.4 + 16);
      const obj = {
        id: B.uid('b'),
        type: 'table',
        x: Math.round((B.W - cw * cols) / 2),
        y: Math.round(Math.max(40, (B.H - rh * rows) / 2)),
        fs,
        cols: new Array(cols).fill(cw),
        rows: new Array(rows).fill(rh),
        cells: Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({ html: '' }))),
      };
      const page = this.store.page;
      this.history.begin(page);
      this.store.addObject(page, obj);
      this.history.commit();
      if (this.requestTool) this.requestTool('text');
      this.focusCell(obj.id, 0, 0);
    }

    focusCell(id, r, c, selectAll) {
      const el = this.view.elOf(id);
      const td = el && el.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
      const content = td && td.querySelector('.content');
      if (!content) return;
      content.focus();
      const sel = getSelection();
      const range = document.createRange();
      range.selectNodeContents(content);
      if (!selectAll) range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    }

    // Tab / Shift+Tab：次・前のセルへ（最後のセルで Tab を押すと行を足す）
    moveCell(dir) {
      const ts = this.ts;
      const obj = this.editingObj();
      if (!ts || !obj || !ts.cell) return;
      const R = obj.rows.length, C = obj.cols.length;
      let i = ts.cell[0] * C + ts.cell[1] + dir;
      if (i < 0) return;
      if (i >= R * C) {
        this.tableOp('rowBelow');
        return;
      }
      this.focusCell(obj.id, Math.floor(i / C), i % C, true);
    }

    // 行・列の追加と削除。編集中のセルを基準にする
    tableOp(op) {
      const ts = this.ts;
      const obj = this.editingObj();
      if (!ts || !obj || !ts.cell) return;
      let [r, c] = ts.cell;
      const id = obj.id;
      const page = this.store.page;
      this.endEditing();
      const cur = this.store.findObject(page, id);
      if (!cur) return;
      const t = B.cloneObj(cur);
      const R = t.rows.length, C = t.cols.length;
      const emptyRow = () => Array.from({ length: t.cols.length }, () => ({ html: '' }));
      switch (op) {
        case 'rowAbove': t.rows.splice(r, 0, t.rows[r]); t.cells.splice(r, 0, emptyRow()); break;
        case 'rowBelow': t.rows.splice(r + 1, 0, t.rows[r]); t.cells.splice(r + 1, 0, emptyRow()); r += 1; break;
        case 'colLeft': t.cols.splice(c, 0, t.cols[c]); t.cells.forEach((row) => row.splice(c, 0, { html: '' })); break;
        case 'colRight': t.cols.splice(c + 1, 0, t.cols[c]); t.cells.forEach((row) => row.splice(c + 1, 0, { html: '' })); c += 1; break;
        case 'rowDelete':
          if (R <= 1) { op = 'delete'; break; }
          t.rows.splice(r, 1); t.cells.splice(r, 1); r = Math.min(r, R - 2); break;
        case 'colDelete':
          if (C <= 1) { op = 'delete'; break; }
          t.cols.splice(c, 1); t.cells.forEach((row) => row.splice(c, 1)); c = Math.min(c, C - 2); break;
      }
      this.history.begin(page);
      if (op === 'delete') this.store.removeObject(page, id);
      else this.store.replaceObject(page, t, { structure: true });
      this.history.commit();
      if (op !== 'delete') this.focusCell(id, r, c);
    }

    // 罫線：scope（'cell' | 'row' | 'col' | 'table'）の範囲の edges（'top' 'bottom' 'left' 'right' 'outer' 'inner' 'all'）を style にする。
    // 隣のセルと共有する辺は、両方のセルに同じ値を入れる（重なった線の見た目を揃えるため）
    tableBorder(scope, edges, style) {
      const ts = this.ts;
      const obj = this.editingObj();
      if (!ts || !obj || !ts.cell) return;
      const [r, c] = ts.cell;
      const id = obj.id;
      const page = this.store.page;
      this.endEditing();
      const cur = this.store.findObject(page, id);
      if (!cur) return;
      const t = B.cloneObj(cur);
      const R = t.rows.length, C = t.cols.length;
      const [r0, r1] = scope === 'row' || scope === 'cell' ? [r, r] : [0, R - 1];
      const [c0, c1] = scope === 'col' || scope === 'cell' ? [c, c] : [0, C - 1];
      const v = style === 'thin' ? null : style;
      const put = (rr, cc, side) => {
        if (rr < 0 || cc < 0 || rr >= R || cc >= C) return;
        const cell = t.cells[rr][cc];
        if (v) cell[side] = v; else delete cell[side];
      };
      const edge = (rr, cc, side) => {
        put(rr, cc, side);
        if (side === 'bt') put(rr - 1, cc, 'bb');
        if (side === 'bb') put(rr + 1, cc, 'bt');
        if (side === 'bl') put(rr, cc - 1, 'br');
        if (side === 'br') put(rr, cc + 1, 'bl');
      };
      const outer = edges === 'outer' || edges === 'all';
      const inner = edges === 'inner' || edges === 'all';
      for (let rr = r0; rr <= r1; rr++) {
        for (let cc = c0; cc <= c1; cc++) {
          if ((edges === 'top' || outer) && rr === r0) edge(rr, cc, 'bt');
          if ((edges === 'bottom' || outer) && rr === r1) edge(rr, cc, 'bb');
          if ((edges === 'left' || outer) && cc === c0) edge(rr, cc, 'bl');
          if ((edges === 'right' || outer) && cc === c1) edge(rr, cc, 'br');
          if (inner && rr < r1) edge(rr, cc, 'bb');
          if (inner && cc < c1) edge(rr, cc, 'br');
        }
      }
      this.history.begin(page);
      this.store.replaceObject(page, t, { structure: true });
      this.history.commit();
      this.focusCell(id, r, c);
    }

    // 列の境目のつまみで幅を変える
    colHandleDown(e, h) {
      e.preventDefault();
      e.stopPropagation();
      const box = h.closest('.obj');
      const page = this.store.page;
      const editing = this.ts && this.ts.id === box.dataset.id;
      if (!editing) {
        this.history.begin(page);
        const o = this.store.findObject(page, box.dataset.id);
        this.store.replaceObject(page, B.cloneObj(o), {});
      }
      const obj = this.store.findObject(page, box.dataset.id);
      const i = Number(h.dataset.i);
      const start = this.toLogical(e).x, w0 = obj.cols[i];
      h.setPointerCapture(e.pointerId);
      const move = (ev) => {
        obj.cols[i] = Math.max(40, Math.round(w0 + this.toLogical(ev).x - start));
        this.store.updateObject(page, obj, { source: 'editor' });
      };
      const up = () => {
        h.removeEventListener('pointermove', move);
        h.removeEventListener('pointerup', up);
        h.removeEventListener('pointercancel', up);
        if (!editing) this.history.commit();
      };
      h.addEventListener('pointermove', move);
      h.addEventListener('pointerup', up);
      h.addEventListener('pointercancel', up);
    }

    // テキストボックスの移動・幅変更・削除ハンドル（編集中のみ表示）
    handleDown(e, handle) {
      e.preventDefault(); // 編集中のフォーカスを保つ
      e.stopPropagation();
      const obj = this.objOf(handle);
      if (!obj) return;
      const page = this.store.page;
      if (handle.classList.contains('h-del')) {
        this.store.removeObject(page, obj.id);
        this.endTextSession();
        return;
      }
      const start = this.toLogical(e);
      const orig = { x: obj.x, y: obj.y, w: obj.w || this.view.elOf(obj.id).offsetWidth };
      const mode = handle.classList.contains('h-move') ? 'move' : 'width';
      handle.setPointerCapture(e.pointerId);
      const onMove = (ev) => {
        const p = this.toLogical(ev);
        const o = this.editingObj() || obj;
        if (mode === 'move') {
          o.x = Math.round(orig.x + p.x - start.x);
          o.y = Math.round(orig.y + p.y - start.y);
        } else {
          o.w = Math.max(60, Math.round(orig.w + p.x - start.x));
        }
        this.store.updateObject(page, o, { source: 'editor' });
      };
      const onUp = () => {
        handle.removeEventListener('pointermove', onMove);
        handle.removeEventListener('pointerup', onUp);
        handle.removeEventListener('pointercancel', onUp);
      };
      handle.addEventListener('pointermove', onMove);
      handle.addEventListener('pointerup', onUp);
      handle.addEventListener('pointercancel', onUp);
    }
  }

  B.Editor = Editor;
})(window.B);
