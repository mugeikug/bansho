using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace Bansho
{
    // 提示画面のウィンドウ。編集画面と別のディスプレイがあれば、そこに枠なしの全画面で開く。
    // ディスプレイが 1 つなら普通のウィンドウで開き、ダブルクリック（画面側の全画面要求）で全画面にする。
    // 最前面にするのは編集画面と別のディスプレイにあるときだけ（同じディスプレイで最前面だと編集画面が操作できなくなる）。
    public sealed class PresenterForm : Form
    {
        public WebView2 Web { get; } = new WebView2 { Dock = DockStyle.Fill };
        private readonly Form main;
        private Rectangle normalBounds;
        private bool full;

        public PresenterForm(Form main)
        {
            this.main = main;
            var icon = main.Icon;
            Text = "提示画面 - デジタル板書";
            Icon = icon;
            BackColor = Color.Black;
            Controls.Add(Web);
            StartPosition = FormStartPosition.Manual;
            var mainScreen = Screen.FromControl(main);
            var other = Screen.AllScreens.FirstOrDefault(s => s.DeviceName != mainScreen.DeviceName);
            if (other != null)
            {
                Bounds = other.Bounds;
                SetFull(true);
            }
            else
            {
                var wa = Screen.PrimaryScreen.WorkingArea;
                Bounds = new Rectangle(wa.X + wa.Width / 8, wa.Y + wa.Height / 8, wa.Width * 3 / 4, wa.Height * 3 / 4);
            }
            KeyPreview = true;
            // 全画面のままディスプレイを移ったとき（Win+Shift+矢印など）も、編集画面と同じディスプレイなら最前面をやめる
            LocationChanged += (s, e) => { if (full) TopMost = OnOtherScreen; };
        }

        public void Attach()
        {
            var cw = Web.CoreWebView2;
            cw.Settings.AreDefaultContextMenusEnabled = false;
            cw.Settings.IsStatusBarEnabled = false;
            // 画面側の全画面の要求（ダブルクリック・F11）に合わせて、ウィンドウを全画面にする／戻す
            cw.ContainsFullScreenElementChanged += (s, e) => SetFull(cw.ContainsFullScreenElement || OnOtherScreen && full);
            cw.WindowCloseRequested += (s, e) => Close();
        }

        private bool OnOtherScreen => Screen.FromRectangle(Bounds).DeviceName != Screen.FromControl(main).DeviceName;

        private void SetFull(bool on)
        {
            if (on == full && FormBorderStyle == FormBorderStyle.None) return;
            if (on)
            {
                if (!full) normalBounds = Bounds;
                var scr = Screen.FromRectangle(Bounds);
                FormBorderStyle = FormBorderStyle.None;
                Bounds = scr.Bounds;
                TopMost = OnOtherScreen;
            }
            else
            {
                FormBorderStyle = FormBorderStyle.Sizable;
                TopMost = false;
                if (!normalBounds.IsEmpty) Bounds = normalBounds;
            }
            full = on;
        }
    }
}
