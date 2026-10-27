using System;
using System.Windows.Forms;

namespace Bansho
{
    internal static class Program
    {
        // 引数：開く .board / .pptx のパス（ファイルの関連付け・「送る」から渡される）
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(args.Length > 0 ? args[0] : null));
        }
    }
}
