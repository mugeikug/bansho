'use strict';
// デスクトップ版（WebView2 の exe）で動いているときの、exe とのやり取り。
// ブラウザ版では何もしない（B.host は未定義のまま）。
(function (B) {
  const wv = window.chrome && window.chrome.webview;
  if (!wv) return;
  let seq = 0;
  const pending = new Map();

  wv.addEventListener('message', (e) => {
    const m = e.data || {};
    if (m.id && pending.has(m.id)) {
      pending.get(m.id)(m);
      pending.delete(m.id);
    } else if (m.cmd === 'openFile' && B.host.onOpen) {
      B.host.onOpen(m);
    }
  });

  const toBytes = (b64) => {
    const bin = atob(b64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  };

  B.host = {
    // exe に頼んで結果を待つ
    call(cmd, args) {
      return new Promise((resolve) => {
        const id = ++seq;
        pending.set(id, resolve);
        wv.postMessage(Object.assign({ id, cmd }, args || {}));
      });
    },
    send(cmd, args) { wv.postMessage(Object.assign({ cmd }, args || {})); },
    // exe から届いたファイル → File（_hostPath に保存先のパスを持たせる）
    toFile(m) {
      const f = new File([toBytes(m.b64)], m.name);
      f._hostPath = m.path;
      return f;
    },
    async blobToB64(blob) {
      const buf = new Uint8Array(await blob.arrayBuffer());
      let bin = '';
      for (let i = 0; i < buf.length; i += 0x8000) bin += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
      return btoa(bin);
    },
    onOpen: null,
  };
})(window.B);
