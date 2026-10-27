using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace Bansho
{
    // 提示画面のウィンドウ。拡張ディスプレイがあれば、そこに枠なしの全画面で開く。
    // ディスプレイが 1 つなら普通のウィンドウで開き、ダブルクリック（画面側の全画面要求）で全画面にする。
    public sealed class PresenterForm : Form
    {
        public WebView2 Web { get; } = new WebView2 { Dock = DockStyle.Fill };
        private Rectangle normalBounds;
        private bool full;

        public PresenterForm(Icon icon)
        {
            Text = "提示画面 - デジタル板書";
            Icon = icon;
            BackColor = Color.Black;
            Controls.Add(Web);
            StartPosition = FormStartPosition.Manual;
            var other = Screen.AllScreens.FirstOrDefault(s => !s.Primary);
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
        }

        public void Attach()
        {
            var cw = Web.CoreWebView2;
            cw.Settings.AreDefaultContextMenusEnabled = false;
            cw.Settings.IsStatusBarEnabled = false;
            // 画面側の全画面の要求（ダブルクリック・F11）に合わせて、ウィンドウを全画面にする／戻す
            cw.ContainsFullScreenElementChanged += (s, e) => SetFull(cw.ContainsFullScreenElement || Screen.AllScreens.Length > 1 && full);
            cw.WindowCloseRequested += (s, e) => Close();
        }

        private void SetFull(bool on)
        {
            if (on == full && FormBorderStyle == FormBorderStyle.None) return;
            if (on)
            {
                if (!full) normalBounds = Bounds;
                var scr = Screen.FromRectangle(Bounds);
                FormBorderStyle = FormBorderStyle.None;
                Bounds = scr.Bounds;
                TopMost = Screen.AllScreens.Length > 1;
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
