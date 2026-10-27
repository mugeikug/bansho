'use strict';
// ページの描画。編集画面・提示画面・サムネイル・印刷のすべてがこのクラスを使う。
// doc を差し替えれば別ウィンドウ（提示画面）の DOM にも描画できる。
(function (B) {
  const SVGNS = 'http://www.w3.org/2000/svg';

  class View {
    // opts.pageId: 指定すればそのページに固定（サムネイル・印刷）。省略すると編集中のページに追随。
    // opts.editor: 編集画面用（テキストを編集可能にし、操作用ハンドルを付ける）
    // opts.theme: テーマの上書き（印刷用）
    constructor(store, doc, host, opts) {
      this.store = store;
      this.doc = doc;
      this.opts = opts || {};
      this.root = doc.createElement('div');
      host.appendChild(this.root);
      this.els = new Map();
      this.pageId = null;
      this.applyTheme();
      this.showPage(this.target());
      this.off = store.on((e) => this.handle(e));
    }

    destroy() {
      this.off();
      this.root.remove();
      this.els.clear();
    }

    target() {
      return this.opts.pageId ? this.store.pageById(this.opts.pageId) : this.store.page;
    }

    applyTheme() {
      const theme = this.opts.theme || this.store.doc.theme;
      this.root.className = 'bpage theme-' + theme + (this.opts.editor ? ' editor-view' : '');
    }

    showPage(page) {
      this.root.textContent = '';
      this.els.clear();
      this.bgEl = null;
      this.pageId = page ? page.id : null;
      if (!page) return;
      if (page.bg) this.renderBg(page.bg);
      for (const obj of page.objects) this.root.appendChild(this.create(obj));
    }

    // ページの背景（PowerPoint から読み込んだスライドの背景色・デザイン）。選択や消しゴムの対象にしない
    renderBg(bg) {
      const el = this.doc.createElement('div');
      el.className = 'page-bg';
      if (bg.color) el.style.backgroundColor = bg.color;
      for (const o of bg.objs || []) {
        el.appendChild(this.create(o));
        this.els.delete(o.id);
      }
      this.root.appendChild(el);
      this.bgEl = el;
    }

    elOf(id) { return this.els.get(id) || null; }

    // 表示するページを固定する（null で編集中のページに追随）。提示画面の固定モード用。
    setPinned(pageId) {
      this.opts.pageId = pageId || null;
      const p = this.target();
      if (!p || p.id !== this.pageId) this.showPage(p);
    }

    handle(e) {
      switch (e.type) {
        case 'doc':
          this.applyTheme();
          this.showPage(this.target());
          break;
        case 'theme':
          this.applyTheme();
          break;
        case 'current':
          if (!this.opts.pageId) {
            const p = this.store.page;
            if (!p || p.id !== this.pageId) this.showPage(p);
          }
          break;
        case 'obj-add':
          if (e.page.id === this.pageId) {
            // index があればその重なり順の位置に入れる（root の子はページのオブジェクトと同じ順に並んでいる）
            const ref = e.index != null ? this.root.children[e.index + (this.bgEl ? 1 : 0)] || null : null;
            this.root.insertBefore(this.create(e.obj), ref);
          }
          break;
        case 'obj-update':
          if (e.page.id === this.pageId) this.update(e.obj, e.info);
          break;
        case 'page-reset':
          if (e.page.id === this.pageId) this.showPage(e.page);
          break;
        case 'obj-remove':
          if (e.page.id === this.pageId) {
            const el = this.els.get(e.id);
            if (el) { el.remove(); this.els.delete(e.id); }
          }
          break;
      }
    }

    create(obj) {
      const d = this.doc;
      let el;
      if (obj.type === 'stroke') {
        el = d.createElementNS(SVGNS, 'svg');
        el.setAttribute('class', 'obj stroke' + (obj.kind === 'marker' ? ' marker' : ''));
        el.setAttribute('viewBox', `0 0 ${B.W} ${B.H}`);
        el.setAttribute('width', B.W);
        el.setAttribute('height', B.H);
        const path = d.createElementNS(SVGNS, 'path');
        path.style.fill = B.colorCss(obj.color, obj.kind === 'marker' ? 'hl' : 'c');
        path.setAttribute('d', B.ink.pathOf(obj));
        el.appendChild(path);
      } else if (obj.type === 'table') {
        el = d.createElement('div');
        el.className = 'obj table';
        this.fillTable(el, obj);
      } else if (obj.type === 'image') {
        el = d.createElement('img');
        el.className = 'obj image';
        el.draggable = false;
        el.alt = '';
        el.src = obj.src;
        this.layoutImage(el, obj);
      } else if (obj.type === 'shape') {
        el = d.createElementNS(SVGNS, 'svg');
        el.setAttribute('class', 'obj shape');
        el.setAttribute('viewBox', `0 0 ${B.W} ${B.H}`);
        el.setAttribute('width', B.W);
        el.setAttribute('height', B.H);
        el.innerHTML = B.geom.shapeMarkup(obj);
      } else if (obj.type === 'text') {
        el = d.createElement('div');
        el.className = 'obj text';
        const c = d.createElement('div');
        c.className = 'content';
        c.innerHTML = obj.html;
        el.appendChild(c);
        if (this.opts.editor) {
          c.contentEditable = 'true';
          c.spellcheck = false;
          for (const [cls, title] of [['h-move', 'ドラッグで移動'], ['h-width', 'ドラッグで幅を変更'], ['h-del', 'このテキストを削除']]) {
            const h = d.createElement('div');
            h.className = 'handle ' + cls;
            h.title = title;
            if (cls === 'h-del') h.textContent = '×';
            el.appendChild(h);
          }
        }
        this.layoutText(el, obj);
      } else {
        el = d.createElement('div');
      }
      el.dataset.id = obj.id;
      this.els.set(obj.id, el);
      return el;
    }

    // 表：<table> を組み立てる。編集画面ではセルを編集可能にし、移動・削除・列幅のつまみを付ける
    fillTable(el, obj) {
      const d = this.doc;
      el.textContent = '';
      const tbl = d.createElement('table');
      tbl.className = 'tbl';
      const cg = d.createElement('colgroup');
      for (let i = 0; i < obj.cols.length; i++) cg.appendChild(d.createElement('col'));
      tbl.appendChild(cg);
      const tb = d.createElement('tbody');
      obj.cells.forEach((row, r) => {
        const tr = d.createElement('tr');
        row.forEach((cell, c) => {
          const td = d.createElement('td');
          td.dataset.r = r;
          td.dataset.c = c;
          if (cell.fill) td.style.backgroundColor = B.colorCss(cell.fill, 'c');
          if (cell.al) td.style.textAlign = cell.al;
          if (cell.va) td.style.verticalAlign = cell.va;
          const bw = obj.bw || 2;
          if (cell.bt) td.style.borderTop = B.borderCss(cell.bt, bw);
          if (cell.bb) td.style.borderBottom = B.borderCss(cell.bb, bw);
          if (cell.bl) td.style.borderLeft = B.borderCss(cell.bl, bw);
          if (cell.br) td.style.borderRight = B.borderCss(cell.br, bw);
          const content = d.createElement('div');
          content.className = 'content';
          content.innerHTML = cell.html;
          if (this.opts.editor) { content.contentEditable = 'true'; content.spellcheck = false; }
          td.appendChild(content);
          tr.appendChild(td);
        });
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);
      el.appendChild(tbl);
      if (this.opts.editor) {
        for (const [cls, title] of [['h-move', 'ドラッグで移動'], ['h-del', 'この表を削除']]) {
          const h = d.createElement('div');
          h.className = 'handle ' + cls;
          h.title = title;
          if (cls === 'h-del') h.textContent = '×';
          el.appendChild(h);
        }
        obj.cols.forEach((w, i) => {
          const h = d.createElement('div');
          h.className = 'col-handle';
          h.dataset.i = i;
          h.title = 'ドラッグで列の幅を変更';
          el.appendChild(h);
        });
      }
      this.layoutTable(el, obj);
    }

    layoutTable(el, obj) {
      const s = el.style;
      s.left = obj.x + 'px';
      s.top = obj.y + 'px';
      s.fontSize = obj.fs + 'px';
      s.lineHeight = obj.lh ? String(obj.lh) : '';
      el.classList.toggle('f-ja', obj.font === 'ja');
      s.setProperty('--tb', B.colorCss(obj.color || 'default', 'c'));
      s.setProperty('--tbw', (obj.bw || 2) + 'px');
      const tbl = el.querySelector('table');
      const total = obj.cols.reduce((a, b) => a + b, 0);
      tbl.style.width = total + 'px';
      tbl.querySelectorAll('col').forEach((col, i) => { col.style.width = (obj.cols[i] || 0) + 'px'; });
      tbl.querySelectorAll('tr').forEach((tr, i) => { tr.style.height = (obj.rows[i] || 0) + 'px'; });
      let x = 0;
      el.querySelectorAll('.col-handle').forEach((h, i) => { x += obj.cols[i] || 0; h.style.left = x + 'px'; });
    }

    layoutImage(el, obj) {
      Object.assign(el.style, { left: obj.x + 'px', top: obj.y + 'px', width: obj.w + 'px', height: obj.h + 'px' });
    }

    layoutText(el, obj) {
      const s = el.style;
      s.left = obj.x + 'px';
      s.top = obj.y + 'px';
      s.fontSize = obj.fs + 'px';
      el.classList.toggle('f-ja', obj.font === 'ja');
      s.textAlign = obj.align || '';     // 'center' | 'right'（PowerPoint から読み込んだ文字など）
      s.lineHeight = obj.lh ? String(obj.lh) : '';
      s.columnCount = obj.cols ? String(obj.cols) : '';     // 段組み（PowerPoint の「段」）
      s.columnGap = obj.cols && obj.colGap != null ? obj.colGap + 'px' : '';
      if (obj.w) {
        s.width = obj.w + 'px';
        s.maxWidth = '';
      } else {
        s.width = 'max-content';
        s.maxWidth = Math.max(120, B.W - obj.x - 24) + 'px';
      }
    }

    update(obj, info) {
      const el = this.els.get(obj.id);
      if (!el) return;
      if (obj.type === 'stroke') {
        el.firstChild.setAttribute('d', B.ink.pathOf(obj));
      } else if (obj.type === 'shape') {
        el.innerHTML = B.geom.shapeMarkup(obj);
      } else if (obj.type === 'table') {
        const editorTyping = this.opts.editor && info.source === 'editor';
        if (info.cell && !info.structure) {
          // セルの文字の入力：そのセルだけ書き換える（編集画面は DOM がすでに最新）
          if (!editorTyping) {
            const td = el.querySelector(`td[data-r="${info.cell[0]}"][data-c="${info.cell[1]}"] .content`);
            const cell = obj.cells[info.cell[0]] && obj.cells[info.cell[0]][info.cell[1]];
            if (td && cell && td.innerHTML !== cell.html) td.innerHTML = cell.html;
          }
          this.layoutTable(el, obj);
        } else if (editorTyping) {
          this.layoutTable(el, obj); // 編集中は作り直さない（入力欄のフォーカスを保つ）
        } else {
          const n = this.create(obj);
          el.replaceWith(n);
        }
      } else if (obj.type === 'image') {
        this.layoutImage(el, obj);
        if (el.getAttribute('src') !== obj.src) el.src = obj.src;
      } else if (obj.type === 'text') {
        this.layoutText(el, obj);
        if (!(this.opts.editor && info.source === 'editor')) {
          const c = el.querySelector('.content');
          if (c.innerHTML !== obj.html) c.innerHTML = obj.html;
        }
      }
    }
  }
  B.View = View;
})(window.B);
