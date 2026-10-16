'use strict';
// 保存・開く・PDF出力（印刷）
(function (B) {
  const EXT = '.board';
  const PICKER_TYPES = [{ description: '板書ファイル', accept: { 'application/json': [EXT] } }];
  const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const OPEN_TYPES = [{ description: '板書ファイル・PowerPoint', accept: { 'application/json': [EXT], [PPTX_MIME]: ['.pptx'] } }];
  const isPptx = (name) => /\.pptx$/i.test(name || '');
  const hasFSA = () => typeof window.showSaveFilePicker === 'function';

  class FileIO {
    constructor(store, toast) {
      this.store = store;
      this.toast = toast;
      this.handle = null;   // File System Access API のハンドル（上書き保存用）
      this.name = '無題';
      this.onName = null;
      this.history = null;      // ページ追加の Undo 用（app.js が設定）
      this.beforeChange = null; // 文書を差し替える前に呼ぶ（編集中の文字を確定する）
      this.printViews = [];
      window.addEventListener('beforeprint', () => this.buildPrint());
      window.addEventListener('afterprint', () => this.clearPrint());
    }

    setName(n) {
      this.name = n.replace(/\.board$/i, '');
      if (this.onName) this.onName();
    }

    async save(saveAs) {
      const text = B.serialize(this.store.doc);
      try {
        if (!saveAs && this.handle) {
          await this.write(this.handle, text);
        } else if (hasFSA()) {
          const h = await window.showSaveFilePicker({ suggestedName: this.name + EXT, types: PICKER_TYPES });
          await this.write(h, text);
          this.handle = h;
          this.setName(h.name);
        } else {
          // File System Access API が使えないときはダウンロードで保存
          const a = document.createElement('a');
          a.href = URL.createObjectURL(new Blob([text], { type: 'application/json' }));
          a.download = this.name + EXT;
          a.click();
          setTimeout(() => URL.revokeObjectURL(a.href), 10000);
        }
      } catch (err) {
        if (err && err.name === 'AbortError') return false;
        alert('保存できませんでした。\n' + (err && err.message ? err.message : err));
        return false;
      }
      this.store.setDirty(false);
      this.toast('保存しました：' + this.name + EXT);
      return true;
    }

    async write(handle, text) {
      if (handle.queryPermission && (await handle.queryPermission({ mode: 'readwrite' })) !== 'granted') {
        if ((await handle.requestPermission({ mode: 'readwrite' })) !== 'granted') {
          throw new Error('ファイルへの書き込みが許可されませんでした。');
        }
      }
      const w = await handle.createWritable();
      await w.write(text);
      await w.close();
    }

    // 開く（.board と .pptx）。PowerPoint は読み込み方を選んでから変換する
    async open() {
      try {
        let file, handle = null;
        if (typeof window.showOpenFilePicker === 'function') {
          [handle] = await window.showOpenFilePicker({ types: OPEN_TYPES });
          file = await handle.getFile();
        } else {
          file = await pickFile();
          if (!file) return;
        }
        await this.openFile(file, handle);
      } catch (err) {
        if (err && err.name === 'AbortError') return;
        alert('開けませんでした。\n' + (err && err.message ? err.message : err));
      }
    }

    // ファイルを開く（ドラッグ＆ドロップ・関連付けからも）。handle があれば上書き保存先にする
    async openFile(file, handle) {
      try {
        if (isPptx(file.name)) { await this.importPptx(file); return; }
        if (this.store.dirty && !confirm('保存していない変更があります。破棄して「' + file.name + '」を開きますか？')) return;
        await this.load(file, handle || null);
      } catch (err) {
        alert('開けませんでした。\n' + (err && err.message ? err.message : err));
      }
    }

    async load(file, handle) {
      const doc = B.deserialize(await file.text());
      if (this.beforeChange) this.beforeChange();
      this.handle = handle;
      this.setName(file.name);
      this.store.load(doc);
      this.toast('開きました：' + file.name);
    }

    // ---- PowerPoint の読み込み ----
    async importPptx(file) {
      const probe = await B.probePptx(file);
      const s = this.store;
      const hasContent = s.doc.pages.length > 1 || s.doc.pages.some((p) => p.objects.length);
      const choice = await askPptx(file.name, probe, hasContent);
      if (!choice) return;
      if (choice.where === 'new' && s.dirty && !confirm('保存していない変更があります。破棄して「' + file.name + '」を開きますか？')) return;
      document.body.classList.add('busy');
      try {
        const r = await B.importPptx(file, (i, n) => this.toast(`PowerPoint を読み込み中… ${i} / ${n}`, true), { mode: choice.mode });
        if (this.beforeChange) this.beforeChange();
        if (choice.where === 'new') {
          r.doc.theme = s.doc.theme;
          this.handle = null;
          this.setName(file.name.replace(/\.pptx$/i, ''));
          s.load(r.doc);
          s.setDirty(true); // まだ .board として保存していない
        } else {
          const at = choice.where === 'after' ? s.current + 1 : s.doc.pages.length;
          const run = () => {
            r.doc.pages.forEach((p, i) => s.insertPage(at + i, p));
            s.setCurrent(at);
          };
          if (this.history) this.history.pages(run); else run();
        }
        const msg = `PowerPoint を読み込みました：${file.name}（${r.doc.pages.length} ページ）`;
        if (r.warnings.length) alert(msg + '\n\n次の要素は読み込めなかったか、近い形に置き換えました：\n・' + r.warnings.join('\n・'));
        else this.toast(msg);
      } finally {
        document.body.classList.remove('busy');
      }
    }

    newDoc() {
      if (this.store.dirty && !confirm('保存していない変更があります。破棄して新規作成しますか？')) return;
      this.handle = null;
      this.setName('無題');
      this.store.load({ theme: this.store.doc.theme, pages: [B.newPage()] });
    }

    // 1ページ = 16:9 の用紙1枚。@page で用紙サイズと余白 0 を指定している（app.css）。
    // 印刷画面の送信先にプリンター（Adobe PDF など）を選ぶと 16:9 にならないので、先に案内を出す
    // 白地以外のテーマのときは、白地で出すか画面と同じテーマで出すかも選ぶ（案内を省略していても出す）
    print() {
      const a = document.activeElement;
      if (a && a.blur) a.blur();
      let skip = false;
      try { skip = localStorage.getItem('bansho.printGuide.skip') === '1'; } catch (e) { /* noop */ }
      const themed = this.store.doc.theme !== 'white';
      this.printTheme = 'white';
      if (skip && !themed) { window.print(); return; }
      const dlg = document.getElementById('print-guide');
      dlg.querySelector('.theme-choice').hidden = !themed;
      for (const el of dlg.querySelectorAll('.guide')) el.hidden = skip;
      const same = dlg.querySelector('input[name=print-theme][value=same]');
      dlg.querySelector('input[name=print-theme][value=white]').checked = true; // 毎回「白地」から（配布用が基本）
      if (themed) same.parentElement.lastChild.textContent = '画面と同じ（' + B.THEME_LABELS[this.store.doc.theme] + '）';
      const ok = document.getElementById('print-guide-ok');
      const cancel = document.getElementById('print-guide-cancel');
      const done = (go) => {
        ok.onclick = cancel.onclick = null;
        if (go && themed && same.checked) this.printTheme = this.store.doc.theme;
        if (go && document.getElementById('print-guide-skip').checked) {
          try { localStorage.setItem('bansho.printGuide.skip', '1'); } catch (e) { /* noop */ }
        }
        dlg.close();
        if (go) setTimeout(() => window.print(), 50);
      };
      ok.onclick = () => done(true);
      cancel.onclick = () => done(false);
      dlg.showModal();
      ok.focus();
    }

    buildPrint() {
      this.clearPrint();
      const root = document.getElementById('print-root');
      for (const p of this.store.doc.pages) {
        const sheet = document.createElement('div');
        sheet.className = 'print-page';
        root.appendChild(sheet);
        this.printViews.push(new B.View(this.store, document, sheet, { pageId: p.id, theme: this.printTheme || 'white' }));
      }
    }

    clearPrint() {
      for (const v of this.printViews) v.destroy();
      this.printViews = [];
      document.getElementById('print-root').textContent = '';
    }
  }

  // 画像ファイル（Blob）を読み込み、データURLと大きさを返す。大きすぎる画像は長辺 2400px に縮める。
  B.loadImage = async (blob) => {
    const src = await new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      r.readAsDataURL(blob);
    });
    const img = await new Promise((resolve, reject) => {
      const im = new Image();
      im.onload = () => resolve(im);
      im.onerror = () => reject(new Error('画像を読み込めませんでした。'));
      im.src = src;
    });
    let w = img.naturalWidth || 400, h = img.naturalHeight || 300;
    const MAX = 2400;
    if (blob.type !== 'image/svg+xml' && Math.max(w, h) > MAX) {
      const k = MAX / Math.max(w, h);
      w = Math.round(w * k);
      h = Math.round(h * k);
      const c = document.createElement('canvas');
      c.width = w;
      c.height = h;
      c.getContext('2d').drawImage(img, 0, 0, w, h);
      const type = blob.type === 'image/jpeg' ? 'image/jpeg' : 'image/png';
      return { src: c.toDataURL(type, 0.9), w, h };
    }
    return { src, w, h };
  };

  // PowerPoint の読み込み方を選ぶダイアログ → { mode: 'keep'|'board', where: 'new'|'append'|'after' } / null
  function askPptx(name, probe, hasContent) {
    const dlg = document.getElementById('pptx-dialog');
    document.getElementById('pptx-name').textContent = `${name}（${probe.slides} 枚のスライド）`;
    dlg.querySelector(`input[name=pptx-mode][value=${probe.designed ? 'keep' : 'board'}]`).checked = true;
    document.getElementById('pptx-hint').textContent = probe.designed
      ? 'このプレゼンテーションには背景色やデザインがあるので「元の色・デザインのまま」をおすすめします。'
      : '白い背景のシンプルなスライドなので「板書の色に合わせる」をおすすめします。';
    const where = document.getElementById('pptx-where');
    where.hidden = !hasContent;
    dlg.querySelector('input[name=pptx-where][value=new]').checked = true;
    return new Promise((resolve) => {
      const done = (v) => { dlg.close(); resolve(v); };
      document.getElementById('pptx-cancel').onclick = () => done(null);
      document.getElementById('pptx-ok').onclick = () => done({
        mode: dlg.querySelector('input[name=pptx-mode]:checked').value,
        where: hasContent ? dlg.querySelector('input[name=pptx-where]:checked').value : 'new',
      });
      dlg.oncancel = () => resolve(null);
      dlg.showModal();
      document.getElementById('pptx-ok').focus();
    });
  }

  function pickFile() {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = EXT + ',application/json,.pptx';
      input.onchange = () => resolve(input.files[0] || null);
      input.click();
    });
  }

  B.FileIO = FileIO;
})(window.B);
