using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace Bansho
{
    // 板書の画面（app フォルダの index.html）を WebView2 で表示する本体ウィンドウ。
    // 画面側（js/host.js）とはメッセージでやり取りし、ファイルの読み書き・PDF の保存・提示画面を引き受ける。
    public sealed class MainForm : Form
    {
        private const string Host = "app.bansho.example";
        private readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private readonly List<PresenterForm> presenters = new List<PresenterForm>();
        private string pendingFile;
        private bool dirty;
        private bool closeConfirmed;

        public MainForm(string file)
        {
            pendingFile = file;
            Text = "デジタル板書";
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ClientSize = new Size(1400, 880);
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            Controls.Add(web);
            Load += async (s, e) => await InitAsync();
            FormClosing += OnFormClosing;
        }

        private static string AppDir => Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "app");

        // 自動テスト用：環境変数 BANSHO_AUTOSAVE_DIR があれば、保存のダイアログを出さずにそのフォルダへ保存する
        private static readonly string AutoSaveDir = Environment.GetEnvironmentVariable("BANSHO_AUTOSAVE_DIR");

        private async Task InitAsync()
        {
            var udf = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Bansho", "WebView2");
            var env = await CoreWebView2Environment.CreateAsync(null, udf);
            await web.EnsureCoreWebView2Async(env);
            var cw = web.CoreWebView2;
            cw.SetVirtualHostNameToFolderMapping(Host, AppDir, CoreWebView2HostResourceAccessKind.Allow);
            cw.Settings.IsStatusBarEnabled = false;
            cw.Settings.IsZoomControlEnabled = false;
            cw.WebMessageReceived += OnMessage;
            cw.DocumentTitleChanged += (s, e) => Text = cw.DocumentTitle;
            cw.NewWindowRequested += OnNewWindow;
            cw.Navigate($"https://{Host}/index.html");
        }

        // ---------------- 画面側からのメッセージ ----------------

        private void Reply(object id, Dictionary<string, object> data)
        {
            data["id"] = id;
            web.CoreWebView2.PostWebMessageAsJson(json.Serialize(data));
        }

        private async void OnMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            Dictionary<string, object> m;
            try { m = json.Deserialize<Dictionary<string, object>>(e.WebMessageAsJson); }
            catch { return; }
            var cmd = m.TryGetValue("cmd", out var c) ? c as string : null;
            m.TryGetValue("id", out var id);
            try
            {
                switch (cmd)
                {
                    case "ready":
                        if (pendingFile != null) { SendOpen(pendingFile); pendingFile = null; }
                        break;
                    case "dirty":
                        dirty = m.TryGetValue("value", out var v) && v is bool b && b;
                        break;
                    case "open":
                        Reply(id, OpenDialog());
                        break;
                    case "save":
                        Reply(id, Save(m));
                        break;
                    case "saveBytes":
                        Reply(id, SaveBytes(m));
                        break;
                    case "pdf":
                        Reply(id, await SavePdfAsync(m));
                        break;
                    case "savedForClose":
                        if (m.TryGetValue("ok", out var ok) && ok is bool okb && okb) { closeConfirmed = true; Close(); }
                        break;
                }
            }
            catch (Exception ex)
            {
                if (id != null) Reply(id, new Dictionary<string, object> { ["ok"] = false, ["error"] = ex.Message });
            }
        }

        // 起動時に渡されたファイル（関連付け・送る）を画面側に渡す
        private void SendOpen(string path)
        {
            try
            {
                var data = File.ReadAllBytes(path);
                web.CoreWebView2.PostWebMessageAsJson(json.Serialize(new Dictionary<string, object>
                {
                    ["cmd"] = "openFile", ["path"] = path, ["name"] = Path.GetFileName(path), ["b64"] = Convert.ToBase64String(data),
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "ファイルを開けませんでした。\n" + ex.Message, "デジタル板書", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private Dictionary<string, object> OpenDialog()
        {
            using (var d = new OpenFileDialog { Filter = "板書・PowerPoint (*.board;*.pptx)|*.board;*.pptx|板書ファイル (*.board)|*.board|PowerPoint (*.pptx)|*.pptx", Title = "開く" })
            {
                if (d.ShowDialog(this) != DialogResult.OK) return new Dictionary<string, object> { ["ok"] = false };
                return new Dictionary<string, object>
                {
                    ["ok"] = true, ["path"] = d.FileName, ["name"] = Path.GetFileName(d.FileName),
                    ["b64"] = Convert.ToBase64String(File.ReadAllBytes(d.FileName)),
                };
            }
        }

        // 保存先を決める（上書き保存なら path をそのまま、そうでなければダイアログ）
        private string AskSavePath(string suggested, string filter, string ext)
        {
            if (AutoSaveDir != null) return Path.Combine(AutoSaveDir, suggested);
            using (var d = new SaveFileDialog { Filter = filter, FileName = suggested, DefaultExt = ext, AddExtension = true, OverwritePrompt = true, Title = "保存" })
                return d.ShowDialog(this) == DialogResult.OK ? d.FileName : null;
        }

        private Dictionary<string, object> Save(Dictionary<string, object> m)
        {
            var path = m.TryGetValue("path", out var p) ? p as string : null;
            var saveAs = m.TryGetValue("saveAs", out var sa) && sa is bool sab && sab;
            if (saveAs || string.IsNullOrEmpty(path))
                path = AskSavePath(m["name"] as string, "板書ファイル (*.board)|*.board", "board");
            if (path == null) return new Dictionary<string, object> { ["ok"] = false, ["cancel"] = true };
            WriteAtomic(path, new UTF8Encoding(false).GetBytes(m["text"] as string));
            return new Dictionary<string, object> { ["ok"] = true, ["path"] = path, ["name"] = Path.GetFileName(path) };
        }

        private Dictionary<string, object> SaveBytes(Dictionary<string, object> m)
        {
            var name = m["name"] as string;
            var path = AskSavePath(name, "PowerPoint プレゼンテーション (*.pptx)|*.pptx", "pptx");
            if (path == null) return new Dictionary<string, object> { ["ok"] = false, ["cancel"] = true };
            WriteAtomic(path, Convert.FromBase64String(m["b64"] as string));
            return new Dictionary<string, object> { ["ok"] = true, ["path"] = path, ["name"] = Path.GetFileName(path) };
        }

        // 一時ファイルに書いてから置き換える（書き込み途中で失敗しても元のファイルを壊さない）
        private static void WriteAtomic(string path, byte[] data)
        {
            var tmp = path + ".tmp";
            File.WriteAllBytes(tmp, data);
            if (File.Exists(path)) File.Replace(tmp, path, null);
            else File.Move(tmp, path);
        }

        // PDF：画面側が印刷用のページを組み立てたあとで呼ばれる。1 ページ = 16:9（20 × 11.25 インチ）、余白なし
        private async Task<Dictionary<string, object>> SavePdfAsync(Dictionary<string, object> m)
        {
            var path = AskSavePath(m["name"] as string, "PDF (*.pdf)|*.pdf", "pdf");
            if (path == null) return new Dictionary<string, object> { ["ok"] = false, ["cancel"] = true };
            var s = web.CoreWebView2.Environment.CreatePrintSettings();
            s.PageWidth = 20;
            s.PageHeight = 11.25;
            s.MarginTop = s.MarginBottom = s.MarginLeft = s.MarginRight = 0;
            s.ShouldPrintBackgrounds = true;
            s.ShouldPrintHeaderAndFooter = false;
            s.ScaleFactor = 1;
            var ok = await web.CoreWebView2.PrintToPdfAsync(path, s);
            return new Dictionary<string, object> { ["ok"] = ok, ["path"] = path, ["name"] = Path.GetFileName(path) };
        }

        // ---------------- 提示画面 ----------------
        // 画面側の window.open を、この exe の別ウィンドウとして開く（開いた側から直接描画できる関係を保つ）
        private async void OnNewWindow(object sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            var deferral = e.GetDeferral();
            try
            {
                var f = new PresenterForm(Icon);
                presenters.Add(f);
                f.FormClosed += (s, a) => presenters.Remove(f);
                f.Show();
                await f.Web.EnsureCoreWebView2Async(web.CoreWebView2.Environment);
                f.Attach();
                e.NewWindow = f.Web.CoreWebView2;
                e.Handled = true;
            }
            finally
            {
                deferral.Complete();
            }
        }

        // ---------------- 閉じるとき ----------------
        private async void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (closeConfirmed || !dirty || AutoSaveDir != null)
            {
                foreach (var p in presenters.ToArray()) p.Close();
                return;
            }
            e.Cancel = true;
            var r = MessageBox.Show(this, "保存していない変更があります。保存してから閉じますか？", "デジタル板書",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (r == DialogResult.Cancel) return;
            if (r == DialogResult.No) { closeConfirmed = true; BeginInvoke(new Action(Close)); return; }
            // 保存してから閉じる（保存が終わると画面側から savedForClose が届く）
            await web.CoreWebView2.ExecuteScriptAsync("B.app.files.save(false).then((ok) => chrome.webview.postMessage({ cmd: 'savedForClose', ok }))");
        }
    }
}
