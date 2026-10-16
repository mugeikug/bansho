// 公開用：sw.js の版を上げ、アプリのファイルがすべてキャッシュ対象に入っているか確かめてから GitHub に push する。
// 使い方（Bansho フォルダで）: node tools/deploy.js "変更の説明"
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const root = path.resolve(__dirname, '..');
const swPath = path.join(root, 'sw.js');
let sw = fs.readFileSync(swPath, 'utf8');

// 1) キャッシュ対象の確認（js / css / icons の取りこぼし防止）
const listed = new Set([...sw.matchAll(/'\.\/([^']+)'/g)].map((m) => m[1]));
const need = [];
for (const dir of ['js', 'css', 'icons']) {
  for (const f of fs.readdirSync(path.join(root, dir))) need.push(`${dir}/${f}`);
}
const missing = need.filter((f) => !listed.has(f));
if (missing.length) {
  console.error('sw.js の FILES に入っていないファイルがあります:\n  ' + missing.join('\n  '));
  process.exit(1);
}

// 2) 版を上げる（日付 + 通し番号）
const today = new Date().toISOString().slice(0, 10);
const m = /const VERSION = '([^']+)';/.exec(sw);
let n = 1;
if (m && m[1].startsWith(today + '.')) n = Number(m[1].split('.')[1]) + 1;
const version = `${today}.${n}`;
sw = sw.replace(/const VERSION = '[^']+';/, `const VERSION = '${version}';`);
fs.writeFileSync(swPath, sw);

// 3) commit & push
const msg = process.argv[2] || '更新';
const git = (...a) => execFileSync('git', a, { cwd: root, stdio: 'inherit' });
git('add', '-A');
git('commit', '-m', `${msg}（版 ${version}）\n\nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`);
git('push');
console.log('公開しました。版 ' + version);
