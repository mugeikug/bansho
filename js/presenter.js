'use strict';
// 提示画面（プロジェクター用ウィンドウ）。
// window.open('') で開いた about:blank は開いた側と同じオリジンになるので、
// file:// から起動しても編集画面のスクリプトから直接 DOM を組み立てて描画できる。
(function (B) {
  const NAME = 'bansho-presenter-' + Date.now().toString(36);

  class Presenter {
    constructor(store) {
      this.store = store;
      this.win = null;
      this.view = null;
      this.onKey = null; // 提示画面側で押されたキー（クリッカーの PageDown など）を編集画面へ渡す
      this.pinnedId = null; // 固定モード：表示を固定しているページの id
      this.blank = null;    // 暗転：'black' | 'white' | null
      this.onChange = null;
      window.addEventListener('pagehide', () => this.close());
    }

    isOpen() { return !!(this.win && !this.win.closed && this.view); }

    async open() {
      if (this.isOpen()) { this.win.focus(); return true; }

      // 拡張ディスプレイがあればそちらに開く（「ウィンドウ管理」の許可が必要）
      let features = 'popup=yes,width=1280,height=720';
      try {
        if ('getScreenDetails' in window) {
          const sd = await window.getScreenDetails();
          const other = sd.screens.find((s) => s !== sd.currentScreen);
          if (other) {
            features = `popup=yes,left=${other.availLeft},top=${other.availTop},width=${other.availWidth},height=${other.availHeight}`;
          }
        }
      } catch (e) { /* 許可されなければ既定の位置に開く */ }

      const w = window.open('', NAME, features);
      if (!w) return false;
      this.win = w;
      this.build(w);
      return true;
    }

    // 固定モード：提示画面の表示を今のページに固定する／解除して編集中のページに追随させる
    setPinned(pageId) {
      this.pinnedId = pageId || null;
      if (this.view) this.view.setPinned(this.pinnedId);
      this.changed();
    }

    // 暗転・白画面（PowerPoint の B / W キー相当）。同じものをもう一度押すと解除。
    toggleBlank(mode) {
      this.blank = this.blank === mode ? null : mode;
      this.applyBlank();
      this.changed();
    }

    applyBlank() {
      if (!this.isOpen()) return;
      const el = this.win.document.getElementById('blank');
      if (!el) return;
      el.style.display = this.blank ? 'block' : 'none';
      el.style.background = this.blank === 'white' ? '#fff' : '#000';
    }

    changed() { if (this.onChange) this.onChange(); }

    close() {
      if (this.view) { this.view.destroy(); this.view = null; }
      if (this.win && !this.win.closed) this.win.close();
      this.win = null;
    }

    build(w) {
      const cssUrl = new URL('css/page.css', location.href).href;
      const d = w.document;
      d.open();
      d.write(`<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>提示画面 - デジタル板書</title>
<link rel="stylesheet" href="${cssUrl}">
<style>
html,body{margin:0;height:100%;background:#000;overflow:hidden;cursor:none;user-select:none}
#stage{position:fixed;inset:0;overflow:hidden}
#hint{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:rgba(0,0,0,.75);color:#fff;
 font:16px "Yu Gothic UI","Segoe UI",sans-serif;padding:10px 18px;border-radius:8px;transition:opacity .6s;pointer-events:none}
#hint.hide{opacity:0}
#blank{position:fixed;inset:0;display:none;z-index:5}
</style></head><body><div id="stage"></div>
<div id="blank"></div>
<div id="hint">ダブルクリック（または F11）で全画面表示</div>
<script>
document.addEventListener('dblclick', function () {
  if (document.fullscreenElement) document.exitFullscreen();
  else document.documentElement.requestFullscreen().catch(function () {});
});
setTimeout(function () { document.getElementById('hint').classList.add('hide'); }, 3500);
</script></body></html>`);
      d.close();

      const stage = d.getElementById('stage');
      this.view = new B.View(this.store, d, stage, { pageId: this.pinnedId });
      this.applyBlank();
      const fit = () => {
        const W = w.innerWidth, H = w.innerHeight;
        const s = Math.min(W / B.W, H / B.H);
        const r = this.view.root.style;
        r.transform = `translate(${(W - B.W * s) / 2}px, ${(H - B.H * s) / 2}px) scale(${s})`;
      };
      fit();
      w.addEventListener('resize', fit);
      w.addEventListener('keydown', (e) => { if (this.onKey) this.onKey(e); });
      w.addEventListener('pagehide', () => {
        if (this.win !== w) return;
        if (this.view) { this.view.destroy(); this.view = null; }
        this.win = null;
        this.changed();
      });
    }
  }
  B.Presenter = Presenter;
})(window.B);
