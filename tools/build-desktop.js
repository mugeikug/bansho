// デスクトップ版のビルド：exe を Release でビルドし、インストーラーを作って dist/ に置く。
// 使い方（Bansho フォルダで）: node tools/build-desktop.js
// Inno Setup 6 (ISCC.exe) が要る。Dropbox 内に直接出力すると ISCC のリソース書き込みが失敗するので、一時フォルダに作ってから写す。
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const root = path.resolve(__dirname, '..');
const desk = path.join(root, 'desktop');
const version = /<Version>([^<]+)<\/Version>/.exec(fs.readFileSync(path.join(desk, 'Bansho.csproj'), 'utf8'))[1];

const iscc = [
  path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Inno Setup 6', 'ISCC.exe'),
  'C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe',
  'C:\\Program Files\\Inno Setup 6\\ISCC.exe',
].find((p) => fs.existsSync(p));
if (!iscc) { console.error('Inno Setup 6 (ISCC.exe) が見つかりません。'); process.exit(1); }

console.log(`== exe をビルド (v${version})`);
execFileSync('dotnet', ['build', '-c', 'Release', '--nologo', '-v', 'q'], { cwd: desk, stdio: 'inherit' });
const src = path.join(desk, 'bin', 'Release', 'net48');

console.log('== インストーラーを作成');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bansho-setup-'));
execFileSync(iscc, ['/Q', `/DAppVersion=${version}`, `/DSourceDir=${src}`, `/O${tmp}`, 'Bansho.iss'], { cwd: path.join(desk, 'installer'), stdio: 'inherit' });

const name = `Bansho-Setup-${version}.exe`;
const dist = path.join(root, 'dist');
fs.mkdirSync(dist, { recursive: true });
fs.copyFileSync(path.join(tmp, name), path.join(dist, name));
console.log(`== できました: dist\\${name}`);

// 受け取る方向けの手順書（desktop/installer/guide.html → PDF）。日本語のパスは Edge に渡せないので一時フォルダで作る
const edge = ['C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].find((p) => fs.existsSync(p));
if (edge) {
  fs.copyFileSync(path.join(desk, 'installer', 'guide.html'), path.join(tmp, 'guide.html'));
  execFileSync(edge, ['--headless', '--disable-gpu', '--no-pdf-header-footer', `--user-data-dir=${path.join(tmp, 'prof')}`,
    `--print-to-pdf=${path.join(tmp, 'guide.pdf')}`, 'file:///' + path.join(tmp, 'guide.html').replace(/\\/g, '/')], { stdio: 'ignore' });
  // Edge は書き終える前に戻ることがあるので、PDF ができるまで待つ
  for (let i = 0; i < 60 && !fs.existsSync(path.join(tmp, 'guide.pdf')); i++) Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 500);
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 1000);
  const pdfName = 'デジタル板書_インストールと使い始め.pdf';
  fs.copyFileSync(path.join(tmp, 'guide.pdf'), path.join(dist, pdfName));
  console.log(`== 手順書: dist\\${pdfName}`);
} else {
  console.log('(Edge が見つからないので手順書の PDF は作りませんでした)');
}
fs.rmSync(tmp, { recursive: true, force: true });
