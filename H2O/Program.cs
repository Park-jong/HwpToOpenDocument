using System;
using System.Drawing;
using System.IO;
using System.Security;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public class Program
    {
        [STAThread]
        public static void Main()
        {
            Application.SetCompatibleTextRenderingDefault(false);
            Application.EnableVisualStyles();
            Application.Run(new MainForm());
        }
    }
}