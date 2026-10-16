'use strict';
// データモデルと変更通知。
// 座標はすべて論理座標（1920×1080）。色は意味上の色名（default / red / ...）で持つ。
window.B = window.B || {};
(function (B) {
  B.W = 1920;
  B.H = 1080;
  B.FORMAT = 'bansho-board';
  B.VERSION = 1;

  let seq = 0;
  B.uid = (prefix) => (prefix || 'o') + Date.now().toString(36) + (seq++).toString(36);

  B.newPage = () => ({ id: B.uid('p'), objects: [] });

  // 色の値 → CSS。意味上の色名はテーマの変数に、'#rrggbb'（PowerPoint から読み込んだ色）はそのまま。
  B.colorCss = (c, kind) => (/^#[0-9a-f]{6}$/i.test(c || '') ? c : `var(--${kind || 'c'}-${c || 'default'})`);
  const COLOR_OK = /^(default|red|blue|green|gray|yellow|pink|#[0-9a-f]{6})$/i;
  B.THEMES = ['white', 'board', 'black'];
  B.THEME_LABELS = { white: '白地', board: '黒板（緑）', black: '黒板（黒）' };

  class Store {
    constructor() {
      this.listeners = [];
      this.doc = { theme: 'white', pages: [B.newPage()] };
      this.current = 0;
      this.dirty = false;
    }

    on(fn) {
      this.listeners.push(fn);
      return () => { this.listeners = this.listeners.filter((f) => f !== fn); };
    }

    emit(evt) {
      for (const fn of this.listeners.slice()) {
        try { fn(evt); } catch (err) { console.error(err); }
      }
    }

    get page() { return this.doc.pages[this.current]; }

    pageById(id) { return this.doc.pages.find((p) => p.id === id) || null; }

    setDirty(v) {
      if (this.dirty === v) return;
      this.dirty = v;
      this.emit({ type: 'dirty' });
    }

    // 文書全体の差し替え（新規・開く）
    load(doc) {
      this.doc = doc;
      this.current = 0;
      this.emit({ type: 'doc' });
      this.emit({ type: 'pages' });
      this.emit({ type: 'current' });
      this.setDirty(false);
    }

    setCurrent(i) {
      i = Math.max(0, Math.min(this.doc.pages.length - 1, i));
      if (i === this.current) return;
      this.current = i;
      this.emit({ type: 'current' });
    }

    addPage(afterIndex) {
      const p = B.newPage();
      this.doc.pages.splice(afterIndex + 1, 0, p);
      this.emit({ type: 'pages' });
      this.setDirty(true);
      return afterIndex + 1;
    }

    deletePage(i) {
      if (this.doc.pages.length <= 1) return false;
      this.doc.pages.splice(i, 1);
      if (this.current >= this.doc.pages.length) this.current = this.doc.pages.length - 1;
      else if (this.current > i) this.current--;
      this.emit({ type: 'pages' });
      this.emit({ type: 'current' });
      this.setDirty(true);
      return true;
    }

    addObject(page, obj) {
      page.objects.push(obj);
      this.emit({ type: 'obj-add', page, obj });
      this.setDirty(true);
    }

    // info.source === 'editor' は「編集画面の DOM が既に最新」の意味（テキスト入力中）
    updateObject(page, obj, info) {
      this.emit({ type: 'obj-update', page, obj, info: info || {} });
      this.setDirty(true);
    }

    removeObject(page, id) {
      const i = page.objects.findIndex((o) => o.id === id);
      if (i < 0) return;
      page.objects.splice(i, 1);
      this.emit({ type: 'obj-remove', page, id });
      this.setDirty(true);
    }

    // 同じ id の別オブジェクトに差し替える（確定済みのオブジェクトは書き換えず複製して差し替える＝Undo のため）
    replaceObject(page, obj, info) {
      const i = page.objects.findIndex((o) => o.id === obj.id);
      if (i < 0) return;
      page.objects[i] = obj;
      this.emit({ type: 'obj-update', page, obj, info: info || {} });
      this.setDirty(true);
    }

    findObject(page, id) { return page.objects.find((o) => o.id === id) || null; }

    // 1つのオブジェクトを、同じ重なり順の位置で複数のオブジェクトに置き換える（部分消しゴム）
    replaceWith(page, id, objs) {
      const i = page.objects.findIndex((o) => o.id === id);
      if (i < 0) return;
      page.objects.splice(i, 1, ...objs);
      this.emit({ type: 'obj-remove', page, id });
      objs.forEach((obj, k) => this.emit({ type: 'obj-add', page, obj, index: i + k }));
      this.setDirty(true);
    }

    insertPage(index, page) {
      this.doc.pages.splice(index, 0, page);
      this.emit({ type: 'pages' });
      this.setDirty(true);
    }

    // from 番目のページを to 番目（移動後の位置）へ。表示中のページは同じページのまま。
    movePage(from, to) {
      const pages = this.doc.pages;
      if (from === to || from < 0 || from >= pages.length) return;
      const cur = this.page;
      const [p] = pages.splice(from, 1);
      pages.splice(Math.max(0, Math.min(pages.length, to)), 0, p);
      this.current = pages.indexOf(cur);
      this.emit({ type: 'pages' });
      this.emit({ type: 'current' });
      this.setDirty(true);
    }

    setTheme(theme) {
      if (this.doc.theme === theme) return;
      this.doc.theme = theme;
      this.emit({ type: 'theme' });
      this.setDirty(true);
    }

    // Undo / Redo 用
    resetObjects(page, objects) {
      page.objects = objects;
      this.emit({ type: 'page-reset', page });
      this.setDirty(true);
    }

    resetPages(pages, current) {
      this.doc.pages = pages;
      this.current = Math.max(0, Math.min(pages.length - 1, current));
      this.emit({ type: 'pages' });
      this.emit({ type: 'current' });
      this.setDirty(true);
    }
  }
  B.Store = Store;

  // ---- 元に戻す・やり直し ----
  // ページ内オブジェクト配列の浅いコピーを前後で記録する。
  // 確定済みのオブジェクトは決して書き換えない（変更時は複製して replaceObject）ので、浅いコピーで足りる。
  class History {
    constructor(store) {
      this.store = store;
      this.undoStack = [];
      this.redoStack = [];
      this.pending = null;
      this.onChange = null;
      store.on((e) => { if (e.type === 'doc') this.clear(); });
    }

    clear() {
      this.undoStack = [];
      this.redoStack = [];
      this.pending = null;
      this.changed();
    }

    changed() { if (this.onChange) this.onChange(); }

    // 操作の開始。同じページで既に開始していればそのまま（テキスト編集中のサイズ変更など）。
    begin(page) {
      if (this.pending) {
        if (this.pending.pageId === page.id) return;
        this.commit();
      }
      this.pending = { type: 'objs', pageId: page.id, before: page.objects.slice() };
    }

    commit() {
      const p = this.pending;
      if (!p) return;
      this.pending = null;
      const page = this.store.pageById(p.pageId);
      if (!page) return;
      const after = page.objects.slice();
      if (sameArray(p.before, after)) return;
      p.after = after;
      this.push(p);
    }

    // ページの追加・削除など、ページ構成を変える操作
    pages(fn) {
      this.commit();
      const s = this.store;
      const before = { pages: s.doc.pages.slice(), cur: s.current };
      const r = fn();
      const after = { pages: s.doc.pages.slice(), cur: s.current };
      if (!sameArray(before.pages, after.pages)) this.push({ type: 'pages', before, after });
      return r;
    }

    push(entry) {
      this.undoStack.push(entry);
      if (this.undoStack.length > 300) this.undoStack.shift();
      this.redoStack = [];
      this.changed();
    }

    canUndo() { return this.undoStack.length > 0 || !!this.pending; }
    canRedo() { return this.redoStack.length > 0; }

    undo() {
      this.commit();
      const e = this.undoStack.pop();
      if (!e) return false;
      this.apply(e, 'before');
      this.redoStack.push(e);
      this.changed();
      return true;
    }

    redo() {
      this.commit();
      const e = this.redoStack.pop();
      if (!e) return false;
      this.apply(e, 'after');
      this.undoStack.push(e);
      this.changed();
      return true;
    }

    apply(e, which) {
      const s = this.store;
      if (e.type === 'objs') {
        const page = s.pageById(e.pageId);
        if (!page) return;
        s.resetObjects(page, e[which].slice());
        s.setCurrent(s.doc.pages.indexOf(page)); // 変化したページを表示する
      } else {
        s.resetPages(e[which].pages.slice(), e[which].cur);
      }
    }
  }
  B.History = History;

  function sameArray(a, b) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }

  // 保存用プロパティだけを持つ複製（_ で始まるキャッシュは捨てる）
  B.cloneObj = (o) => JSON.parse(JSON.stringify(cleanObj(o)));

  // ---- 保存形式 ----

  function cleanObj(o) {
    const out = {};
    for (const k of Object.keys(o)) if (k[0] !== '_') out[k] = o[k];
    return out;
  }

  B.serialize = (doc) => JSON.stringify({
    format: B.FORMAT,
    version: B.VERSION,
    theme: doc.theme,
    pages: doc.pages.map((p) => Object.assign({ id: p.id, objects: p.objects.map(cleanObj) }, p.bg ? { bg: { color: p.bg.color || null, objs: (p.bg.objs || []).map(cleanObj) } } : {})),
  });

  B.deserialize = (text) => {
    let data;
    try { data = JSON.parse(text); } catch (e) { throw new Error('ファイルの形式が正しくありません（JSONとして読めません）。'); }
    if (!data || data.format !== B.FORMAT || !Array.isArray(data.pages)) {
      throw new Error('板書ファイルではありません。');
    }
    if (typeof data.version !== 'number' || data.version > B.VERSION) {
      throw new Error('このファイルは新しいバージョンのアプリで保存されています（形式 v' + data.version + '）。');
    }
    const pages = data.pages.map((p) => {
      const page = {
        id: typeof p.id === 'string' ? p.id : B.uid('p'),
        objects: (Array.isArray(p.objects) ? p.objects : []).map(normalizeObj).filter(Boolean),
      };
      if (p.bg && typeof p.bg === 'object') {
        page.bg = {
          color: /^#[0-9a-f]{6}$/i.test(p.bg.color || '') ? p.bg.color : null,
          objs: (Array.isArray(p.bg.objs) ? p.bg.objs : []).map(normalizeObj).filter(Boolean),
        };
      }
      return page;
    });
    if (!pages.length) pages.push(B.newPage());
    return { theme: B.THEMES.includes(data.theme) ? data.theme : 'white', pages };
  };

  B.normalizeObj = (o) => normalizeObj(o);

  function normalizeObj(o) {
    if (!o || typeof o !== 'object') return null;
    if (typeof o.id !== 'string') o.id = B.uid();
    if (o.type === 'stroke') {
      if (!Array.isArray(o.pts) || !o.pts.length) return null;
      if (!COLOR_OK.test(o.color || '')) o.color = o.kind === 'marker' ? 'yellow' : 'default';
      return o;
    }
    if (o.type === 'text') {
      o.html = B.sanitizeHtml(String(o.html || ''));
      if (o.align !== 'center' && o.align !== 'right') delete o.align;
      if (!(typeof o.lh === 'number' && o.lh > 0.5 && o.lh < 4)) delete o.lh;
      if (o.font !== 'ja') delete o.font;
      if (!(Number.isInteger(o.cols) && o.cols >= 2 && o.cols <= 4)) { delete o.cols; delete o.colGap; }
      else if (!(typeof o.colGap === 'number' && o.colGap >= 0 && o.colGap < 400)) delete o.colGap;
      return o;
    }
    if (o.type === 'table') {
      const okNum = (v) => typeof v === 'number' && v >= 0 && v < 5000;
      if (typeof o.x !== 'number' || typeof o.y !== 'number' || !okNum(o.fs)) return null;
      if (!Array.isArray(o.cols) || !o.cols.length || o.cols.length > 40 || !o.cols.every(okNum)) return null;
      if (!Array.isArray(o.rows) || !o.rows.length || o.rows.length > 200 || !o.rows.every(okNum)) return null;
      if (!Array.isArray(o.cells) || o.cells.length !== o.rows.length) return null;
      o.cells = o.cells.map((row) => o.cols.map((_, c) => {
        const cell = (Array.isArray(row) && row[c]) || {};
        const out = { html: B.sanitizeHtml(String(cell.html || '')) };
        if (/^#[0-9a-f]{6}$/i.test(cell.fill || '')) out.fill = cell.fill;
        if (cell.al === 'center' || cell.al === 'right') out.al = cell.al;
        if (cell.va === 'middle' || cell.va === 'bottom') out.va = cell.va;
        return out;
      }));
      if (o.color != null && !COLOR_OK.test(o.color)) o.color = 'default';
      if (!(typeof o.lh === 'number' && o.lh > 0.5 && o.lh < 4)) delete o.lh;
      if (o.font !== 'ja') delete o.font;
      return o;
    }
    if (o.type === 'image') {
      if (typeof o.src !== 'string' || !/^data:image\/(png|jpeg|gif|webp|svg\+xml|bmp);/.test(o.src)) return null;
      for (const k of ['x', 'y', 'w', 'h']) if (typeof o[k] !== 'number') return null;
      return o;
    }
    if (o.type === 'shape') {
      if (!['line', 'arrow', 'rect', 'ellipse'].includes(o.shape)) return null;
      for (const k of ['x1', 'y1', 'x2', 'y2', 'width']) if (typeof o[k] !== 'number') return null;
      if (!COLOR_OK.test(o.color || '')) o.color = 'default';
      return o;
    }
    return null; // 未知の種類は読み飛ばす
  }

  // ---- テキストの HTML を許可したタグだけに制限 ----

  const ALLOWED = new Set(['B', 'STRONG', 'U', 'I', 'EM', 'S', 'SUB', 'SUP', 'SPAN', 'BR', 'DIV', 'P']);
  const DROP = new Set(['SCRIPT', 'STYLE', 'IFRAME', 'OBJECT', 'EMBED', 'TEMPLATE', 'svg', 'SVG', 'MATH']);
  const CLASS_OK = /^(c|hl|deco|ul)-[a-z0-9]+$/;
  // 要素ごとに許す style（rich.js が出力するものだけ）
  const NUM_EM = /^-?\d+(\.\d+)?em$/;
  const HEX = /^#[0-9a-f]{6}$/i;
  const STYLE_RULES = {
    SPAN: { 'font-size': NUM_EM, color: HEX, 'font-family': /^'[^'"<>;\\{}()]{1,80}'$/ },
    U: { 'text-decoration-color': HEX },
    DIV: { 'padding-left': NUM_EM, 'text-indent': NUM_EM, 'margin-top': NUM_EM, 'text-align': /^(left|center|right)$/, '--bw': NUM_EM, '--bs': /^\d+(\.\d+)?$/, '--bc': /^(#[0-9a-f]{6}|var\(--c-(default|red|blue|green|gray)\))$/i },
  };
  function cleanStyle(tag, value) {
    const rules = STYLE_RULES[tag];
    if (!rules) return '';
    const out = [];
    for (const decl of value.split(';')) {
      if (!decl.trim()) continue;
      const i = decl.indexOf(':');
      const prop = decl.slice(0, i).trim().toLowerCase(), val = decl.slice(i + 1).trim();
      if (i < 0 || !rules[prop] || !rules[prop].test(val)) return '';
      out.push(`${prop}:${val}`);
    }
    return out.join(';');
  }

  B.sanitizeHtml = (html) => {
    const tpl = document.createElement('template');
    tpl.innerHTML = html;
    clean(tpl.content);
    return tpl.innerHTML;
  };

  function clean(node) {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) continue;
      if (child.nodeType !== 1) { child.remove(); continue; }
      if (DROP.has(child.tagName)) { child.remove(); continue; }
      clean(child);
      if (!ALLOWED.has(child.tagName)) {
        child.replaceWith(...Array.from(child.childNodes));
        continue;
      }
      for (const a of Array.from(child.attributes)) {
        if (a.name === 'style') {
          const st = cleanStyle(child.tagName, a.value);
          if (st) child.setAttribute('style', st);
          else child.removeAttribute('style');
          continue;
        }
        // 箇条書きの記号（短い文字列だけ）
        if (a.name === 'data-bu' && child.tagName === 'DIV' && a.value.length <= 6 && !/[<>&"]/.test(a.value)) continue;
        if (a.name !== 'class') child.removeAttribute(a.name);
      }
      if (child.hasAttribute('class')) {
        const cls = child.className.split(/\s+/).filter((c) => CLASS_OK.test(c));
        if (cls.length) child.className = cls.join(' ');
        else child.removeAttribute('class');
      }
    }
  }
})(window.B);
