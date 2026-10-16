'use strict';
// 自動バックアップ：未保存の変更があれば一定間隔でブラウザ内（IndexedDB）に保存する。
// ブラウザが落ちたり保存せずに閉じたりしても、次に開いたときに復元できる。
// タブ（起動）ごとに別の記録を持ち、保存済み（未保存の変更なし）になったら自分の記録を消す。
(function (B) {
  const DB = 'bansho';
  const STORE = 'backups';
  const INTERVAL = 15000;
  const KEEP_DAYS = 14;

  function openDb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function tx(db, mode, fn) {
    return new Promise((resolve, reject) => {
      const t = db.transaction(STORE, mode);
      const st = t.objectStore(STORE);
      const r = fn(st);
      t.oncomplete = () => resolve(r && 'result' in r ? r.result : undefined);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  }

  class Backup {
    constructor(store, files) {
      this.store = store;
      this.files = files;
      this.id = 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      this.changed = false;
      this.dbp = null;
      this.lastTime = 0;
      this.onSaved = null;
      store.on((e) => {
        if (e.type === 'dirty') {
          if (!store.dirty) this.remove(); // 保存した・開き直した
        } else if (e.type !== 'current') {
          this.changed = true;
        }
      });
      setInterval(() => this.tick(), INTERVAL);
      // タブを閉じる・切り替える直前にも書いておく
      document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') this.tick(); });
    }

    db() {
      if (!this.dbp) this.dbp = openDb().catch((e) => { console.warn('IndexedDB を使えません', e); return null; });
      return this.dbp;
    }

    async tick() {
      if (!this.changed || !this.store.dirty) return;
      this.changed = false;
      const db = await this.db();
      if (!db) return;
      const d = this.store.doc;
      const rec = {
        id: this.id,
        name: this.files.name,
        time: Date.now(),
        pages: d.pages.length,
        json: B.serialize(d),
      };
      try {
        await tx(db, 'readwrite', (st) => st.put(rec));
        this.lastTime = rec.time;
        if (this.onSaved) this.onSaved(rec.time);
      } catch (e) {
        console.warn('自動バックアップに失敗しました', e);
        this.changed = true;
      }
    }

    async remove(id) {
      const db = await this.db();
      if (!db) return;
      try { await tx(db, 'readwrite', (st) => st.delete(id || this.id)); } catch (e) { /* noop */ }
    }

    // ほかの（以前の）起動で残ったバックアップ。新しい順。古すぎるものは消す。
    async others() {
      const db = await this.db();
      if (!db) return [];
      let all = [];
      try { all = await tx(db, 'readonly', (st) => st.getAll()); } catch (e) { return []; }
      const limit = Date.now() - KEEP_DAYS * 86400000;
      const out = [];
      for (const r of all || []) {
        if (r.id === this.id) continue;
        if (r.time < limit) this.remove(r.id);
        else out.push(r);
      }
      return out.sort((a, b) => b.time - a.time);
    }

    // 復元：その内容を開き、未保存の状態にする（以後はこの起動のバックアップとして保存される）
    restore(rec) {
      const doc = B.deserialize(rec.json);
      this.files.handle = null;
      this.files.setName(rec.name || '無題');
      this.store.load(doc);
      this.store.setDirty(true);
      this.changed = true;
      this.remove(rec.id);
      this.tick();
    }
  }
  B.Backup = Backup;
})(window.B);
