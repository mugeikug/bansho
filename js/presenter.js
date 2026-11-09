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
        if (!B.host && 'getScreenDetails' in window) {
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

    closeAndNotify() {
      this.close();
      this.changed();
    }

    menuOpen(w) { return w.document.getElementById('menu').style.display === 'block'; }
    hideMenu(w) { w.document.getElementById('menu').style.display = 'none'; }

    // 右クリックのメニュー（全画面表示の切り替え・閉じる）と、動かしたときだけ見えるマウスポインター
    setupControls(w) {
      const d = w.document;
      const menu = d.getElementById('menu');
      const full = d.getElementById('m-full');
      d.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        full.textContent = d.fullscreenElement ? '全画面表示を終了（Esc）' : '全画面表示';
        menu.style.display = 'block';
        const x = Math.min(e.clientX, w.innerWidth - menu.offsetWidth - 4);
        const y = Math.min(e.clientY, w.innerHeight - menu.offsetHeight - 4);
        menu.style.left = Math.max(0, x) + 'px';
        menu.style.top = Math.max(0, y) + 'px';
      });
      d.addEventListener('mousedown', (e) => { if (e.button === 0 && !menu.contains(e.target)) this.hideMenu(w); });
      full.addEventListener('click', () => {
        this.hideMenu(w);
        if (d.fullscreenElement) d.exitFullscreen().catch(() => {});
        else d.documentElement.requestFullscreen().catch(() => {});
      });
      d.getElementById('m-close').addEventListener('click', () => this.closeAndNotify());
      let timer = 0;
      d.addEventListener('mousemove', () => {
        d.body.classList.remove('idle');
        w.clearTimeout(timer);
        timer = w.setTimeout(() => { if (!this.menuOpen(w)) d.body.classList.add('idle'); }, 2000);
      });
    }

    build(w) {
      const cssUrl = new URL('css/page.css', location.href).href;
      const d = w.document;
      d.open();
      d.write(`<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>提示画面 - デジタル板書</title>
<link rel="stylesheet" href="${cssUrl}">
<style>
html,body{margin:0;height:100%;background:#000;overflow:hidden;user-select:none}
body.idle{cursor:none}
#stage{position:fixed;inset:0;overflow:hidden}
#hint{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:rgba(0,0,0,.75);color:#fff;
 font:16px "Yu Gothic UI","Segoe UI",sans-serif;padding:10px 18px;border-radius:8px;transition:opacity .6s;pointer-events:none}
#hint.hide{opacity:0}
#blank{position:fixed;inset:0;display:none;z-index:5}
#menu{position:fixed;display:none;z-index:10;min-width:200px;padding:4px 0;background:#fff;color:#222;border-radius:8px;
 box-shadow:0 4px 16px rgba(0,0,0,.4);font:16px "Yu Gothic UI","Segoe UI",sans-serif;cursor:default}
#menu button{display:block;width:100%;border:0;background:none;padding:9px 20px;text-align:left;font:inherit;color:inherit;cursor:pointer}
#menu button:hover{background:#e8eef8}
</style></head><body class="idle"><div id="stage"></div>
<div id="blank"></div>
<div id="menu"><button id="m-full"></button><button id="m-close">提示画面を閉じる（Esc）</button></div>
<div id="hint">ダブルクリック（または F11）で全画面表示　右クリック・Esc で閉じる</div>
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
      this.setupControls(w);
      w.addEventListener('keydown', (e) => {
        // Esc：メニューを閉じる → 全画面表示を終える → 提示画面を閉じる
        if (e.key === 'Escape') {
          e.preventDefault();
          if (this.menuOpen(w)) this.hideMenu(w);
          else if (w.document.fullscreenElement) w.document.exitFullscreen().catch(() => {});
          else this.closeAndNotify();
          return;
        }
        if (this.onKey) this.onKey(e);
      });
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
