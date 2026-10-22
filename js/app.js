'use strict';
// 編集画面の組み立て：ツールバー、ページ一覧、ショートカット、提示画面の制御。
(function (B) {
  const $ = (s) => document.querySelector(s);
  const $$ = (s) => Array.from(document.querySelectorAll(s));
  const PEN_COLORS = [['default', '既定色'], ['red', '赤'], ['blue', '青'], ['green', '緑']];
  const MARKER_COLORS = [['yellow', '黄'], ['pink', 'ピンク'], ['blue', '水色'], ['green', '緑']];
  const TEXT_COLORS = PEN_COLORS.concat([['gray', 'グレー']]);
  const TOOL_KEYS = { v: 'select', t: 'text', p: 'pen', h: 'marker', e: 'eraser', l: 'line', a: 'arrow', r: 'rect', o: 'ellipse' };

  const state = {
    tool: 'pen',
    penColor: 'default',
    penSize: 1,
    markerColor: 'yellow',
    markerSize: 0,
    textPt: 24,
    textMarker: 'yellow', // Ctrl+H でかけるマーカーの色（最後に使った色）
    eraserMode: 'stroke', // 'stroke'（線ごと）| 'partial'（部分）
    olFormat: '1.',       // 番号付きリストの形式（最後に選んだもの）
    eraserSize: 1,
  };

  const store = new B.Store();
  const history = new B.History(store);
  const editor = new B.Editor(store, history, $('#stage'), state);
  const presenter = new B.Presenter(store);
  const files = new B.FileIO(store, toast);
  const backup = new B.Backup(store, files);
  files.history = history;
  files.beforeChange = () => editor.endEditing();
  B.app = { store, history, editor, presenter, files, state, backup };

  // ---------------- ツール ----------------

  function setTool(t) {
    editor.setTool(t);
    for (const b of $$('#tools button')) b.classList.toggle('on', b.dataset.tool === t);
    const penLike = t === 'pen' || t === 'marker' || t === 'select' || B.SHAPES.includes(t);
    $('#pen-opts').style.display = penLike ? '' : 'none';
    $('#text-opts').style.display = t === 'text' || t === 'select' ? '' : 'none';
    $('#eraser-opts').style.display = t === 'eraser' ? '' : 'none';
    $('#no-opts').style.display = 'none';
    renderPenOpts();
    renderEraserOpts();
    updateFormatState();
  }
  editor.requestTool = setTool;

  function renderPenOpts() {
    const marker = state.tool === 'marker';
    const colors = $('#colors');
    colors.className = 'theme-' + store.doc.theme;
    colors.textContent = '';
    for (const [key, label] of marker ? MARKER_COLORS : PEN_COLORS) {
      const b = document.createElement('button');
      b.className = 'swatch' + ((marker ? state.markerColor : state.penColor) === key ? ' on' : '');
      b.style.background = `var(--${marker ? 'hl' : 'c'}-${key})`;
      b.title = label + (state.tool === 'select' ? '（選択中の線・図形の色を変える）' : '');
      b.onclick = () => {
        if (marker) state.markerColor = key;
        else state.penColor = key;
        if (state.tool === 'select') editor.restyleSelection(key, null);
        renderPenOpts();
      };
      colors.appendChild(b);
    }
    const sizes = $('#sizes');
    sizes.textContent = '';
    const n = marker ? B.MARKER_SIZES.length : B.PEN_SIZES.length;
    for (let i = 0; i < n; i++) {
      const b = document.createElement('button');
      b.className = 'size-btn' + ((marker ? state.markerSize : state.penSize) === i ? ' on' : '');
      b.title = (['細', '中', '太'][i] || '') + (state.tool === 'select' ? '（選択中の線・図形の太さを変える）' : '');
      const dot = document.createElement('i');
      const px = marker ? 8 + i * 8 : 3 + i * 4;
      dot.style.width = dot.style.height = px + 'px';
      b.appendChild(dot);
      b.onclick = () => {
        if (marker) state.markerSize = i;
        else state.penSize = i;
        if (state.tool === 'select') editor.restyleSelection(null, i);
        renderPenOpts();
      };
      sizes.appendChild(b);
    }
  }

  // ---------------- 消しゴム ----------------

  function renderEraserOpts() {
    for (const b of $$('[data-eraser-mode]')) b.classList.toggle('on', b.dataset.eraserMode === state.eraserMode);
    const box = $('#eraser-sizes');
    box.textContent = '';
    B.ERASER_SIZES.forEach((r, i) => {
      const b = document.createElement('button');
      b.className = 'size-btn' + (state.eraserSize === i ? ' on' : '');
      b.title = '消しゴムの大きさ：' + ['小', '中', '大'][i];
      const dot = document.createElement('i');
      dot.style.width = dot.style.height = (6 + i * 6) + 'px';
      dot.style.background = 'transparent';
      dot.style.border = '1.5px solid currentColor';
      b.appendChild(dot);
      b.onclick = () => { state.eraserSize = i; renderEraserOpts(); };
      box.appendChild(b);
    });
  }
  for (const b of $$('[data-eraser-mode]')) {
    b.onclick = () => { state.eraserMode = b.dataset.eraserMode; renderEraserOpts(); };
  }

  // ---------------- テーマ ----------------

  const themeSelect = $('#theme-select');
  themeSelect.addEventListener('change', () => {
    store.setTheme(themeSelect.value);
    themeSelect.blur();
  });
  function onThemeChanged() {
    themeSelect.value = store.doc.theme;
    renderPenOpts();
    buildTextOpts();
    updateFormatState();
  }

  // ---------------- 画像 ----------------

  const imageFile = $('#image-file');
  imageFile.addEventListener('change', async () => {
    const list = Array.from(imageFile.files || []);
    imageFile.value = '';
    await insertImages(list);
  });

  async function insertImages(blobs, at) {
    let n = 0;
    for (const blob of blobs) {
      try {
        const img = await B.loadImage(blob);
        const p = at ? { x: at.x + n * 40, y: at.y + n * 40 } : n ? { x: B.W / 2 + n * 40, y: B.H / 2 + n * 40 } : null;
        editor.insertImage(img, p);
        n++;
      } catch (err) {
        toast((blob.name || '画像') + ' を読み込めませんでした');
      }
    }
  }

  // ---------------- 表 ----------------

  // ツールバーのボタンの下に出す小さなメニュー（押してもフォーカス・選択範囲を奪わない）
  function openPopup(menu, anchor) {
    if (!menu.hidden) { menu.hidden = true; return; }
    for (const m of $$('.popup')) m.hidden = true;
    const r = anchor.getBoundingClientRect();
    menu.style.left = Math.min(r.left, window.innerWidth - 260) + 'px';
    menu.style.top = (r.bottom + 4) + 'px';
    menu.hidden = false;
  }
  for (const m of $$('.popup')) {
    m.addEventListener('mousedown', (e) => e.preventDefault());
    m.addEventListener('pointerdown', (e) => e.preventDefault());
  }
  document.addEventListener('pointerdown', (e) => {
    for (const m of $$('.popup')) {
      if (!m.hidden && !m.contains(e.target) && !e.target.closest('#btn-table, #btn-tableops, #fs-drop, [data-menu]')) m.hidden = true;
    }
  }, true);

  // 表の大きさを選ぶマス目（8 列 × 8 行）
  (() => {
    const grid = $('#table-grid');
    const label = $('#table-size');
    for (let r = 1; r <= 8; r++) {
      for (let c = 1; c <= 8; c++) {
        const i = document.createElement('i');
        i.dataset.r = r;
        i.dataset.c = c;
        grid.appendChild(i);
      }
    }
    const mark = (R, C) => {
      for (const i of grid.children) i.classList.toggle('on', Number(i.dataset.r) <= R && Number(i.dataset.c) <= C);
      label.textContent = R ? `${R} 行 × ${C} 列` : '表の大きさを選ぶ';
    };
    grid.addEventListener('pointerover', (e) => { if (e.target.dataset.r) mark(Number(e.target.dataset.r), Number(e.target.dataset.c)); });
    grid.addEventListener('pointerleave', () => mark(0, 0));
    grid.addEventListener('click', (e) => {
      if (!e.target.dataset.r) return;
      $('#table-picker').hidden = true;
      mark(0, 0);
      editor.insertTable(Number(e.target.dataset.r), Number(e.target.dataset.c));
    });
  })();

  $('#table-ops').addEventListener('click', (e) => {
    const b = e.target.closest('[data-op]');
    if (!b) return;
    $('#table-ops').hidden = true;
    if (b.dataset.op === 'delete' && !confirm('表全体を削除しますか？（元に戻す で取り消せます）')) return;
    editor.tableOp(b.dataset.op);
  });

  // ---------------- コピー・貼り付け ----------------
  // 板書内のコピーもシステムのクリップボードに載せる（後から別のものをコピーしたら、そちらが優先される）。

  const busyInText = () => !!editor.ts || document.activeElement === fsInput;
  let lastClipText = '';

  document.addEventListener('copy', (e) => {
    if (busyInText() || !editor.copySelection()) return;
    lastClipText = B.CLIP_PREFIX + JSON.stringify(editor.clip);
    e.clipboardData.setData('text/plain', lastClipText);
    e.preventDefault();
  });
  document.addEventListener('cut', (e) => {
    if (busyInText() || !editor.copySelection()) return;
    lastClipText = B.CLIP_PREFIX + JSON.stringify(editor.clip);
    e.clipboardData.setData('text/plain', lastClipText);
    e.preventDefault();
    editor.deleteSelection();
  });
  document.addEventListener('paste', (e) => {
    if (busyInText() || $('dialog[open]')) return;
    const dt = e.clipboardData;
    const images = Array.from(dt.files || []).filter((f) => f.type.startsWith('image/'));
    e.preventDefault();
    if (images.length) { insertImages(images); return; }
    const text = dt.getData('text/plain');
    if (text.startsWith(B.CLIP_PREFIX)) {
      if (text !== lastClipText) {
        // 別のウィンドウ（別の板書）でコピーしたもの
        try {
          const objs = JSON.parse(text.slice(B.CLIP_PREFIX.length)).map(B.normalizeObj).filter(Boolean);
          editor.clip = objs;
          editor.pasteCount = 0;
          lastClipText = text;
        } catch (err) { return; }
      }
      setTool('select');
      editor.paste();
    } else if (text.trim()) {
      editor.insertText(text);
    } else if (editor.clip) {
      setTool('select');
      editor.paste();
    }
  });

  // ---------------- ドラッグ＆ドロップ（画像・.board ファイル） ----------------

  const dropHint = $('#drop-hint');
  let dragDepth = 0;
  const hasFiles = (e) => Array.from(e.dataTransfer ? e.dataTransfer.types : []).includes('Files');
  document.addEventListener('dragenter', (e) => {
    if (!hasFiles(e)) return;
    dragDepth++;
    dropHint.hidden = false;
  });
  document.addEventListener('dragleave', (e) => {
    if (!hasFiles(e)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) dropHint.hidden = true;
  });
  document.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  document.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth = 0;
    dropHint.hidden = true;
    const list = Array.from(e.dataTransfer.files);
    const doc = list.find((f) => /\.(board|pptx)$/i.test(f.name));
    if (doc) { files.openFile(doc); return; }
    const images = list.filter((f) => f.type.startsWith('image/'));
    if (!images.length) { toast('画像・.board・.pptx ファイルをドロップしてください'); return; }
    // ページ上に落としたらその位置に置く
    const r = editor.wrap.getBoundingClientRect();
    const inside = e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom;
    insertImages(images, inside ? editor.toLogical(e) : null);
  });

  // ---------------- 文字の装飾 ----------------

  function buildTextOpts() {
    const tc = $('#text-colors');
    tc.className = 'theme-' + store.doc.theme;
    tc.textContent = '';
    TEXT_COLORS.forEach(([key, label], i) => {
      const b = document.createElement('button');
      b.className = 'fmt tcolor';
      b.dataset.fmt = 'c';
      b.dataset.val = key;
      b.title = `文字色：${label} (Alt+${i + 1})`;
      b.innerHTML = `<span style="color:var(--c-${key})">A</span><i style="background:var(--c-${key})"></i>`;
      tc.appendChild(b);
    });
    // マーカー：最後に使った色のボタン＋▾（色を選ぶ）
    const tm = $('#text-markers');
    tm.className = 'split theme-' + store.doc.theme;
    tm.innerHTML = `<button class="fmt tmarker main" data-fmt="h" data-val="${state.textMarker}" title="マーカー (Ctrl+H)"><span style="background:var(--hl-${state.textMarker})">ab</span></button><button class="drop" data-menu="marker-menu" title="マーカーの色">▾</button>`;
    buildMenus();
  }

  // ---- 文字の大きさ（pt） ----

  const fsInput = $('#fs-input');
  const fsMenu = $('#fs-menu');

  // ▾ のメニュー（下線の色・枠の色・マーカーの色・番号の形式）
  const COLOR_CHOICES = [['', '文字と同じ'], ['default', '既定色'], ['red', '赤'], ['blue', '青'], ['green', '緑'], ['gray', 'グレー']];
  function buildMenus() {
    const theme = 'theme-' + store.doc.theme;
    const sw = (c) => (c ? `<span class="sw" style="background:var(--c-${c})"></span>` : '<span class="sw" style="border:2px solid currentColor;box-sizing:border-box"></span>');
    const um = $('#ul-menu');
    um.className = 'popup menu-colors ' + theme;
    um.innerHTML = '<div class="mtitle">一重線</div><div class="row">' + COLOR_CHOICES.map(([c, l]) => `<button data-ul="sng:${c}" title="${l}">${sw(c)}</button>`).join('') + '</div>'
      + '<div class="mtitle">二重線</div><div class="row">' + COLOR_CHOICES.map(([c, l]) => `<button data-ul="dbl:${c}" title="${l}">${sw(c)}</button>`).join('') + '</div>';
    const bm = $('#box-menu');
    bm.className = 'popup menu-colors ' + theme;
    bm.innerHTML = '<div class="mtitle">枠の色</div><div class="row">' + COLOR_CHOICES.map(([c, l]) => `<button data-box="${c}" title="${l}">${sw(c)}</button>`).join('') + '</div>';
    const mm = $('#marker-menu');
    mm.className = 'popup menu-colors ' + theme;
    mm.innerHTML = '<div class="mtitle">マーカーの色</div><div class="row">' + MARKER_COLORS.map(([c, l]) => `<button data-marker="${c}" title="${l}"><span style="background:var(--hl-${c});padding:1px 5px">ab</span></button>`).join('') + '</div>';
    const om = $('#ol-menu');
    const sample = { '1.': '1. 2. 3.', '1)': '1) 2) 3)', '(1)': '(1) (2) (3)', 'i.': 'i. ii. iii.', 'i)': 'i) ii) iii)', 'I.': 'I. II. III.', 'a.': 'a. b. c.', 'a)': 'a) b) c)', 'A.': 'A. B. C.', 'A)': 'A) B) C)' };
    om.innerHTML = B.OL_FORMATS.map((f) => `<button data-ol="${f}">${sample[f]}</button>`).join('');
  }
  document.addEventListener('click', (e) => {
    const d = e.target.closest('[data-menu]');
    if (d) { openPopup($('#' + d.dataset.menu), d); return; }
    const pick = e.target.closest('[data-ul], [data-box], [data-marker], [data-ol], [data-align]');
    if (!pick || !pick.closest('.popup')) return;
    pick.closest('.popup').hidden = true;
    if (pick.dataset.ul != null) {
      const [st, c] = pick.dataset.ul.split(':');
      applyFormat('u', st === 'sng' && !c ? true : st + (c ? ':' + c : ''));
    } else if (pick.dataset.box != null) {
      applyFormat('x', pick.dataset.box || true);
    } else if (pick.dataset.marker) {
      state.textMarker = pick.dataset.marker;
      buildTextOpts();
      applyFormat('h', pick.dataset.marker);
    } else if (pick.dataset.ol) {
      state.olFormat = pick.dataset.ol;
      editor.toggleNumbering(pick.dataset.ol);
      updateFormatState();
    } else if (pick.dataset.align) {
      editor.alignLines(pick.dataset.align);
      updateFormatState();
    }
  });

  function renderTextSizes() {
    if (document.activeElement === fsInput) return; // 入力中は書き換えない
    const pt = editor.fontPt();
    fsInput.value = pt == null ? '' : String(pt);
  }

  function applyFontPt(pt) {
    if (!(pt > 0)) { renderTextSizes(); return; }
    if (editor.ts) editor.restoreRange();
    editor.setFontPt(pt);
    renderTextSizes();
    updateFormatState();
  }

  function stepFont(dir) {
    editor.stepFontPt(dir);
    renderTextSizes();
  }

  // 入力欄をクリックした時点で、編集中の文字の選択範囲を覚えておく（入力欄にフォーカスが移ると選択が見えなくなるため）
  fsInput.addEventListener('mousedown', () => { if (editor.ts && document.activeElement !== fsInput) editor.rememberRange(); });
  fsInput.addEventListener('focus', () => { if (editor.ts && !editor.savedRange) editor.rememberRange(); fsInput.select(); });
  fsInput.addEventListener('keydown', (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    e.stopPropagation();
    if (e.key === 'Enter') {
      e.preventDefault();
      const pt = parseFloat(fsInput.value.replace(/[^\d.]/g, ''));
      fsInput.blur();
      applyFontPt(pt);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      fsInput.blur();
      if (editor.ts) editor.restoreRange();
      renderTextSizes();
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      e.preventDefault();
      if (editor.ts) editor.restoreRange();
      stepFont(e.key === 'ArrowUp' ? 1 : -1);
      if (!editor.ts) fsInput.focus();
    }
  });
  fsInput.addEventListener('blur', () => {
    // 入力欄から編集中の文字以外へフォーカスが移ったら、文字の編集も終える
    setTimeout(() => {
      const a = document.activeElement;
      if (editor.ts && a !== fsInput && !editor.ts.content.contains(a)) editor.endTextSession();
      editor.savedRange = null;
      renderTextSizes();
    }, 0);
  });

  $('#fs-drop').addEventListener('click', () => {
    if (!fsMenu.hidden) { fsMenu.hidden = true; return; }
    if (editor.ts) editor.rememberRange();
    const cur = editor.fontPt();
    fsMenu.textContent = '';
    for (const v of B.PT_STEPS) {
      const b = document.createElement('button');
      b.textContent = v;
      if (v === cur) b.className = 'on';
      b.onclick = () => { fsMenu.hidden = true; applyFontPt(v); };
      fsMenu.appendChild(b);
    }
    const r = $('.fontsize').getBoundingClientRect();
    fsMenu.style.left = r.left + 'px';
    fsMenu.style.top = (r.bottom + 4) + 'px';
    fsMenu.style.minWidth = r.width + 'px';
    fsMenu.hidden = false;
    const on = fsMenu.querySelector('.on');
    if (on) on.scrollIntoView({ block: 'center' });
  });
  // 候補を押しても編集中の文字からフォーカスを奪わない
  fsMenu.addEventListener('mousedown', (e) => e.preventDefault());
  fsMenu.addEventListener('pointerdown', (e) => e.preventDefault());
  document.addEventListener('pointerdown', (e) => {
    if (!fsMenu.hidden && !fsMenu.contains(e.target) && !e.target.closest('#fs-drop')) fsMenu.hidden = true;
  }, true);

  function applyFormat(key, val) {
    if (val === undefined) val = true;
    if (key === 'h' && val !== state.textMarker) { state.textMarker = val; buildTextOpts(); }
    if (editor.ts) editor.formatText(key, val);
    else if (state.tool === 'select') editor.formatSelectedTexts(key, val);
    updateFormatState();
  }

  $('#text-opts').addEventListener('click', (e) => {
    const b = e.target.closest('[data-fmt]');
    if (b) applyFormat(b.dataset.fmt, b.dataset.val || true);
  });

  function updateFormatState() {
    const st = editor.formatState();
    for (const b of $$('#text-opts [data-fmt]')) {
      const k = b.dataset.fmt, v = b.dataset.val;
      let on = false;
      if (st) {
        if (k === 'c') on = st.c !== undefined && st.c === (v === 'default' ? null : v);
        else if (k === 'h' || k === 's') on = st[k] === v;
        else on = !!st[k]; // 下線は色・二重線つきの値もある
      }
      b.classList.toggle('on', on);
    }
    const ls = editor.lineState();
    $('#btn-bullets').classList.toggle('on', !!(ls && ls.bu && !ls.ol));
    $('#btn-numbering').classList.toggle('on', !!(ls && ls.ol));
    const ta = (ls && ls.ta) || 'left';
    const icon = $(`#align-menu [data-align="${ta}"] svg`);
    if (icon) $('#btn-align').innerHTML = icon.outerHTML;
    for (const b of $$('#align-menu [data-align]')) b.classList.toggle('on', b.dataset.align === ta);
    for (const b of $$('#ol-menu [data-ol]')) b.classList.toggle('on', !!ls && ls.ol === b.dataset.ol);
  }
  document.addEventListener('selectionchange', () => { if (editor.ts) { updateFormatState(); renderTextSizes(); } });
  editor.onEditingChange = () => {
    renderTextSizes();
    updateFormatState();
    if (history.onChange) history.onChange();
    $('#btn-tableops').hidden = !(editor.ts && editor.ts.cell);
  };
  editor.onSelectionChange = () => { renderTextSizes(); updateFormatState(); };

  // ツールバーのボタンを押しても編集中のテキストからフォーカス・選択範囲を奪わない
  $('#toolbar').addEventListener('mousedown', (e) => { if (e.target.closest('button')) e.preventDefault(); });
  $('#toolbar').addEventListener('pointerdown', (e) => { if (e.target.closest('button')) e.preventDefault(); });

  for (const b of $$('#tools button')) b.onclick = () => setTool(b.dataset.tool);

  // ---------------- コマンド ----------------

  const commands = {
    new: () => files.newDoc(),
    open: () => files.open(),
    save: () => files.save(false),
    saveas: () => files.save(true),
    print: () => files.print(),
    pptx: () => files.exportPptx(),
    undo: () => undo(),
    redo: () => redo(),
    present: () => togglePresenter(),
    fontup: () => stepFont(1),
    fontdown: () => stepFont(-1),
    pin: () => togglePin(),
    black: () => blank('black'),
    white: () => blank('white'),
    addpage: () => addPage(),
    duppage: () => duplicatePage(),
    delpage: () => deletePage(),
    image: () => imageFile.click(),
    table: () => openPopup($('#table-picker'), $('#btn-table')),
    bullets: () => { if (editor.toggleBullets()) updateFormatState(); },
    numbering: () => { if (editor.toggleNumbering(state.olFormat)) updateFormatState(); },
    indent: () => editor.indentLines(1),
    outdent: () => editor.indentLines(-1),
    tableops: () => openPopup($('#table-ops'), $('#btn-tableops')),
    help: () => showHelp(),
  };
  for (const b of $$('[data-cmd]')) {
    b.addEventListener('click', () => commands[b.dataset.cmd]());
  }

  function undo() {
    if (editor.active) return;
    if (editor.ts) editor.textUndo();
    else history.undo();
    updateFormatState();
  }

  function redo() {
    if (editor.active) return;
    if (editor.ts) editor.textRedo();
    else history.redo();
    updateFormatState();
  }

  history.onChange = () => {
    $('#btn-undo').disabled = !history.canUndo() && !editor.ts;
    $('#btn-redo').disabled = !history.canRedo() && !editor.ts;
  };

  // ---------------- ページ ----------------

  function gotoPage(i) {
    editor.endEditing();
    store.setCurrent(i);
  }

  function addPage() {
    editor.endEditing();
    history.pages(() => store.setCurrent(store.addPage(store.current)));
  }

  // ページを複製して直後に入れる（中身はすべて新しい id の複製）
  function duplicatePage() {
    editor.endEditing();
    const src = store.page;
    const copy = { id: B.uid('p'), objects: src.objects.map((o) => Object.assign(B.cloneObj(o), { id: B.uid(o.type[0]) })) };
    if (src.bg) copy.bg = JSON.parse(JSON.stringify(src.bg)); // 背景は変更されないので中身の id はそのままでよい
    history.pages(() => {
      store.insertPage(store.current + 1, copy);
      store.setCurrent(store.current + 1);
    });
    toast('ページを複製しました');
  }

  function movePage(from, to) {
    if (to < 0 || to >= store.doc.pages.length || from === to) return;
    editor.endEditing();
    history.pages(() => store.movePage(from, to));
  }

  function deletePage() {
    if (store.doc.pages.length <= 1) { toast('最後の1ページは削除できません'); return; }
    const p = store.page;
    if (p.objects.length && !confirm(`${store.current + 1} ページ目を削除しますか？\n（元に戻す で取り消せます）`)) return;
    editor.endEditing();
    history.pages(() => store.deletePage(store.current));
  }

  let thumbViews = [];
  function renderThumbs() {
    for (const v of thumbViews) v.destroy();
    thumbViews = [];
    const box = $('#thumbs');
    box.textContent = '';
    store.doc.pages.forEach((p, i) => {
      const t = document.createElement('div');
      t.className = 'thumb';
      t.innerHTML = `<div class="num">${i + 1}</div><div class="frame"></div>`;
      bindThumbDrag(t, i);
      box.appendChild(t);
      const v = new B.View(store, document, t.querySelector('.frame'), { pageId: p.id });
      v.root.style.transform = `scale(${154 / B.W})`;
      thumbViews.push(v);
    });
    markCurrent();
  }

  // サムネイルをクリックで移動、ドラッグで並べ替え
  function bindThumbDrag(t, i) {
    const box = $('#thumbs');
    let start = null, dragging = false, line = null, target = i;
    t.addEventListener('pointerdown', (e) => {
      if (e.button !== 0 && e.pointerType === 'mouse') return;
      start = { x: e.clientX, y: e.clientY };
      dragging = false;
      t.setPointerCapture(e.pointerId);
    });
    t.addEventListener('pointermove', (e) => {
      if (!start) return;
      if (!dragging) {
        if (Math.hypot(e.clientX - start.x, e.clientY - start.y) < 8) return;
        dragging = true;
        t.classList.add('dragging');
        line = document.createElement('div');
        line.className = 'drop-line';
        box.appendChild(line);
      }
      // 端に近づいたらスクロール
      const br = box.getBoundingClientRect();
      if (e.clientY < br.top + 30) box.scrollTop -= 12;
      else if (e.clientY > br.bottom - 30) box.scrollTop += 12;
      // 挿入位置（k 番目のサムネイルの前）
      const thumbs = $$('.thumb');
      let k = thumbs.length;
      for (let j = 0; j < thumbs.length; j++) {
        const r = thumbs[j].getBoundingClientRect();
        if (e.clientY < r.top + r.height / 2) { k = j; break; }
      }
      target = k;
      const ref = thumbs[Math.min(k, thumbs.length - 1)];
      const y = k < thumbs.length ? ref.offsetTop - 5 : ref.offsetTop + ref.offsetHeight + 2;
      line.style.top = y + 'px';
    });
    const end = (e) => {
      if (!start) return;
      start = null;
      if (line) line.remove();
      line = null;
      t.classList.remove('dragging');
      if (!dragging) { if (e.type === 'pointerup') gotoPage(i); return; }
      dragging = false;
      const to = target > i ? target - 1 : target;
      if (to !== i) movePage(i, to);
    };
    t.addEventListener('pointerup', end);
    t.addEventListener('pointercancel', end);
  }

  // 提示画面に出ているページ
  function presentedIndex() {
    if (!presenter.isOpen()) return -1;
    if (presenter.pinnedId) return store.doc.pages.findIndex((p) => p.id === presenter.pinnedId);
    return store.current;
  }

  function markCurrent() {
    const thumbs = $$('.thumb');
    const pi = presentedIndex();
    thumbs.forEach((t, i) => {
      t.classList.toggle('current', i === store.current);
      t.classList.toggle('presented', i === pi);
      t.title = i === pi ? '提示画面に表示中' : '';
    });
    const cur = thumbs[store.current];
    if (cur) cur.scrollIntoView({ block: 'nearest' });
    $('#page-no').textContent = `ページ ${store.current + 1} / ${store.doc.pages.length}`;
  }

  store.on((e) => {
    if (e.type === 'pages' || e.type === 'doc') {
      if (presenter.pinnedId && !store.pageById(presenter.pinnedId)) presenter.setPinned(null);
      renderThumbs();
    } else if (e.type === 'current') {
      markCurrent();
      renderTextSizes();
    } else if (e.type === 'dirty') {
      updateTitle();
    } else if (e.type === 'theme') {
      onThemeChanged();
    }
    if (e.type === 'doc') onThemeChanged();
  });

  // ---------------- 提示画面 ----------------

  async function togglePresenter() {
    if (presenter.isOpen()) { presenter.win.focus(); return; }
    const ok = await presenter.open();
    if (!ok) {
      alert('提示画面を開けませんでした。\nブラウザがポップアップをブロックした場合は、アドレスバー右端の表示から「常に許可」を選んでから、もう一度押してください。');
    }
    updatePresentState();
  }

  function togglePin() {
    presenter.setPinned(presenter.pinnedId ? null : store.page.id);
    toast(presenter.pinnedId ? `提示画面を ${store.current + 1} ページに固定しました` : '固定を解除しました（編集中のページに追随します）');
  }

  function blank(mode) {
    presenter.toggleBlank(mode);
    if (!presenter.isOpen()) toast('提示画面が開いていません');
  }

  function updatePresentState() {
    const open = presenter.isOpen();
    $('#btn-present').classList.toggle('on', open);
    $('#btn-present').textContent = open ? '提示中' : '提示画面';
    $('#btn-pin').classList.toggle('on', !!presenter.pinnedId);
    $('#btn-black').classList.toggle('on', presenter.blank === 'black');
    $('#btn-white').classList.toggle('on', presenter.blank === 'white');
    const s = $('#present-state');
    let msg = '提示画面：閉じています';
    if (open) {
      msg = '● 提示画面に表示中';
      if (presenter.pinnedId) {
        const i = store.doc.pages.findIndex((p) => p.id === presenter.pinnedId);
        msg += `（${i + 1} ページに固定）`;
      }
      if (presenter.blank) msg += presenter.blank === 'black' ? '・暗転中' : '・白画面中';
    }
    s.textContent = msg;
    s.classList.toggle('on', open);
    s.classList.toggle('warn', open && (!!presenter.pinnedId || !!presenter.blank));
    markCurrent();
  }
  presenter.onChange = updatePresentState;
  // 提示画面が閉じられたことを検知する
  setInterval(() => {
    const open = presenter.isOpen();
    if (open !== $('#btn-present').classList.contains('on')) updatePresentState();
  }, 1000);

  // ---------------- ショートカット ----------------

  function onKey(e) {
    if (e.isComposing || e.keyCode === 229) return; // IME 変換中は一切触らない
    if ($('dialog[open]')) return; // ダイアログ表示中はショートカットを無効に
    if (document.activeElement === fsInput) return; // 文字サイズの入力欄は自分でキーを処理する
    const inText = !!editor.ts && !!(document.activeElement && document.activeElement.isContentEditable);
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    const lower = k.length === 1 ? k.toLowerCase() : k;
    const textTarget = inText || (state.tool === 'select' && editor.selectedTexts().length > 0);
    const stop = () => { e.preventDefault(); e.stopPropagation(); };

    if (ctrl && !e.altKey) {
      if (lower === 's') { stop(); files.save(e.shiftKey); return; }
      if (lower === 'o') { stop(); files.open(); return; }
      if (lower === 'p') { stop(); files.print(); return; }
      if (lower === 'm') { stop(); addPage(); return; }
      if (lower === 'z') { stop(); if (e.shiftKey) redo(); else undo(); return; }
      if (lower === 'y') { stop(); redo(); return; }
      if (inText && e.shiftKey && lower === 'l') { stop(); editor.toggleBullets(); return; }
      if (textTarget || state.tool === 'text') {
        if (k === '>' || k === ']') { stop(); stepFont(1); return; }
        if (k === '<' || k === '[') { stop(); stepFont(-1); return; }
      }
      if (textTarget) {
        const fmt = {
          b: ['b'], i: ['i'], u: ['u'], k: ['x'], h: ['h', state.textMarker],
          ...(e.shiftKey ? { x: ['k'] } : {}),
          '=': ['s', 'sub'], ',': ['s', 'sub'], '+': ['s', 'sup'], '.': ['s', 'sup'],
        }[lower];
        if (fmt) { stop(); applyFormat(fmt[0], fmt[1]); return; }
      }
      if (inText) return; // 文字入力中の Ctrl+C / V / A などはブラウザ標準
      // Ctrl+C / X / V は copy・cut・paste イベントで処理する（画像の貼り付けのため既定動作を止めない）
      if (lower === 'd' && e.shiftKey) { stop(); duplicatePage(); return; }
      if (lower === 'd') { stop(); if (editor.copySelection()) editor.paste(); return; }
      if (e.shiftKey && k === 'ArrowUp') { stop(); movePage(store.current, store.current - 1); return; }
      if (e.shiftKey && k === 'ArrowDown') { stop(); movePage(store.current, store.current + 1); return; }
      if (lower === 'a') { stop(); setTool('select'); editor.selectAll(); return; }
      return;
    }
    if (e.altKey && !ctrl && /^Digit[1-5]$/.test(e.code)) {
      if (textTarget) { stop(); applyFormat('c', TEXT_COLORS[Number(e.code[5]) - 1][0]); }
      return;
    }
    if (k === 'Tab' && inText && !ctrl && !e.altKey) { stop(); editor.handleTab(e.shiftKey); return; }
    if (k === 'Escape') {
      if (inText) { stop(); editor.endEditing(); }
      else editor.setSelection([]);
      return;
    }
    if (k === 'F5') { stop(); togglePresenter(); return; }
    if (k === 'F1') { stop(); showHelp(); return; }
    if (inText) return; // 文字入力中は1文字のショートカットを無効にする

    if ((k === 'Delete' || k === 'Backspace') && editor.selection.size) { stop(); editor.deleteSelection(); return; }
    const arrows = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    if (arrows[k] && editor.selection.size) {
      stop();
      const d = e.shiftKey ? 10 : 1;
      editor.nudge(arrows[k][0] * d, arrows[k][1] * d);
      return;
    }
    if (k === 'PageDown' || k === 'ArrowDown' || k === 'ArrowRight') { stop(); gotoPage(store.current + 1); return; }
    if (k === 'PageUp' || k === 'ArrowUp' || k === 'ArrowLeft') { stop(); gotoPage(store.current - 1); return; }
    if (k === 'Home') { stop(); gotoPage(0); return; }
    if (k === 'End') { stop(); gotoPage(store.doc.pages.length - 1); return; }
    if (k === '?') { stop(); showHelp(); return; }
    if (e.altKey || e.shiftKey || ctrl) return;
    if (lower === 'b' || lower === '.') { stop(); blank('black'); return; }
    if (lower === 'w') { stop(); blank('white'); return; }
    if (lower === 'f') { stop(); togglePin(); return; }
    if (lower === 'i') { stop(); imageFile.click(); return; }
    const tool = TOOL_KEYS[lower];
    if (tool) { stop(); setTool(tool); }
  }
  document.addEventListener('keydown', onKey);
  presenter.onKey = (e) => {
    // 提示画面側では、ページ移動（クリッカー）と暗転・白画面・固定だけを受け付ける
    const pageKeys = ['PageDown', 'PageUp', 'ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight', 'Home', 'End'];
    if (pageKeys.includes(e.key)) {
      e.preventDefault();
      if (/Down|Right/.test(e.key)) gotoPage(store.current + 1);
      else if (/Up|Left/.test(e.key)) gotoPage(store.current - 1);
      else gotoPage(e.key === 'Home' ? 0 : store.doc.pages.length - 1);
    } else if (!e.ctrlKey && !e.altKey && !e.metaKey) {
      const lower = e.key.toLowerCase();
      if (lower === 'b' || lower === '.') { e.preventDefault(); blank('black'); }
      else if (lower === 'w') { e.preventDefault(); blank('white'); }
      else if (lower === 'f') { e.preventDefault(); togglePin(); }
    }
  };

  // ---------------- その他 ----------------

  function updateTitle() {
    document.title = (store.dirty ? '● ' : '') + files.name + ' - デジタル板書';
  }
  files.onName = updateTitle;

  window.addEventListener('beforeunload', (e) => {
    if (store.dirty) { e.preventDefault(); e.returnValue = ''; }
  });

  let toastTimer = 0;
  function toast(msg, sticky) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.remove('hide');
    clearTimeout(toastTimer);
    if (!sticky) toastTimer = setTimeout(() => t.classList.add('hide'), 3000);
  }

  // ---------------- ショートカット一覧 ----------------

  const HELP = [
    ['ツール', [
      ['V', '選択'], ['T', '文字'], ['P', 'ペン'], ['H', '蛍光ペン'], ['E', '消しゴム'],
      ['L / A / R / O', '直線 / 矢印 / 四角 / 楕円（描くとき Shift で水平・垂直・正方形・正円）'],
      ['I', '画像を挿入（Ctrl+V やドラッグ＆ドロップでも可）'],
      ['ペンの消しゴム側・サイドボタン', '消しゴム'],
    ]],
    ['文字（入力中は範囲に、選択ツールではボックス全体に）', [
      ['Ctrl+B / I / U', '太字 / 斜体 / 下線'],
      ['Ctrl+= または Ctrl+,', '下付き'], ['Ctrl+Shift+= または Ctrl+.', '上付き'],
      ['Ctrl+H', 'マーカー（最後に使った色）'], ['Ctrl+K', '四角囲み'],
      ['Alt+1〜5', '文字色：既定色 / 赤 / 青 / 緑 / グレー'],
      ['Ctrl+Shift+X', '取り消し線'],
      ['Ctrl+Shift+> / <（Ctrl+] / [）', '文字を1段階大きく / 小さく'],
      ['Tab / Shift+Tab', '行頭（箇条書き・字下げのある行）や複数行の選択では字下げを深く / 浅く。それ以外はタブ文字。表では次 / 前のセルへ'],
      ['Ctrl+Shift+L', '箇条書きのオン・オフ'],
      ['（ツールバー）', '番号付きリスト・行の揃え・下線の色と二重線・枠の色は ▾ から選ぶ'],
      ['Esc', '文字の入力を終える'],
    ]],
    ['編集', [
      ['Ctrl+Z / Ctrl+Y', '元に戻す / やり直し（文字入力中は文字単位）'],
      ['Ctrl+C / X / V', 'コピー / 切り取り / 貼り付け（画像・文字の貼り付けも可）'],
      ['Ctrl+D', '選択したものを複製'], ['Ctrl+A', 'ページ内をすべて選択'],
      ['Delete', '選択したものを削除'], ['← ↑ → ↓', '選択したものを移動（Shift で大きく）'],
    ]],
    ['ページ', [
      ['PageDown・↓・→ / PageUp・↑・←', '次 / 前のページ（何も選択していないとき）'],
      ['Home / End', '最初 / 最後のページ'],
      ['Ctrl+M', 'ページを追加'], ['Ctrl+Shift+D', 'ページを複製'],
      ['Ctrl+Shift+↑ / ↓', 'ページを前 / 後ろへ移動（サムネイルのドラッグでも可）'],
    ]],
    ['提示画面', [
      ['F5', '提示画面を開く'], ['F', 'このページに固定 / 解除'],
      ['B（または .） / W', '暗転 / 白画面（もう一度で解除）'],
      ['提示画面で PageDown など', 'ページ移動（クリッカー）。B / W / F も効きます'],
    ]],
    ['ファイル', [
      ['Ctrl+S / Ctrl+Shift+S', '上書き保存 / 名前を付けて保存'], ['Ctrl+O', '開く'],
      ['Ctrl+P', 'PDF出力'], ['F1 または ?', 'この一覧'],
    ]],
  ];

  function showHelp() {
    const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const keys = (t) => esc(t).split(/( \/ | または |、|（|）)/).filter((part) => part !== '').map((part) =>
      /^( \/ | または |、|（|）)$/.test(part) || /[ぁ-んァ-ン一-龥]/.test(part) ? part : `<kbd>${part}</kbd>`).join('');
    $('#help-body').innerHTML = '<table>' + HELP.map(([sec, rows]) =>
      `<tr><th colspan="2">${esc(sec)}</th></tr>` + rows.map(([k, d]) => `<tr><td>${keys(k)}</td><td>${esc(d)}</td></tr>`).join('')).join('') + '</table>';
    const dlg = $('#help-dialog');
    $('#help-close').onclick = () => dlg.close();
    dlg.showModal();
  }

  // ---------------- 自動バックアップからの復元 ----------------

  async function offerRestore() {
    const list = await backup.others();
    if (!list.length) return;
    const rec = list[0];
    const t = new Date(rec.time);
    const pad = (n) => String(n).padStart(2, '0');
    const when = `${t.getFullYear()}/${t.getMonth() + 1}/${t.getDate()} ${pad(t.getHours())}:${pad(t.getMinutes())}`;
    $('#restore-info').textContent = `「${rec.name}」 ${rec.pages} ページ（${when} 時点）` + (list.length > 1 ? `　ほかに ${list.length - 1} 件` : '');
    const dlg = $('#restore-dialog');
    const close = () => dlg.close();
    $('#restore-ok').onclick = () => {
      close();
      if (store.dirty && !confirm('今の内容は破棄されます。復元しますか？')) return;
      try { backup.restore(rec); toast('復元しました。必要なら「別名保存」で保存してください'); } catch (err) { alert('復元できませんでした。\n' + err.message); }
    };
    $('#restore-discard').onclick = () => {
      if (!confirm('残っている作業内容を破棄します。よろしいですか？')) return;
      close();
      for (const r of list) backup.remove(r.id);
    };
    $('#restore-later').onclick = close;
    dlg.showModal();
  }

  // ---------------- アプリとして使うとき（インストール版） ----------------

  // オフラインで起動できるよう、アプリのファイルを保存しておく（http(s) で開いたときだけ。file:// では使わない）
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('サービスワーカーを登録できませんでした', e));
  }

  // .board / .pptx をダブルクリック・「送る」などで開いたとき、そのファイルを受け取る
  // （.board は上書き保存先としても使えるファイルハンドルで届く）
  if ('launchQueue' in window) {
    window.launchQueue.setConsumer((params) => openLaunched(params));
  }
  async function openLaunched(params) {
    const handle = (params.files || [])[0]; // 1 つのウィンドウでは 1 ファイルだけ開く
    if (!handle) return;
    try {
      const file = await handle.getFile();
      await files.openFile(file, /\.board$/i.test(file.name) ? handle : null);
    } catch (e) {
      alert('ファイルを開けませんでした。\n' + (e && e.message ? e.message : e));
    }
  }
  B.app.openLaunched = openLaunched;

  buildTextOpts();
  renderThumbs();
  setTool(state.tool);
  renderTextSizes();
  updatePresentState();
  updateTitle();
  history.onChange();
  onThemeChanged();
  offerRestore();
})(window.B);
