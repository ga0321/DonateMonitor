using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DonateMonitor
{
    internal static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Global.InitSettings();
            Global.LoadSettings();
            Application.Run(new Startup());
            if (!Global._bExit && Global.IsEnableAnyService())
                Application.Run(new Monitor());
        }
    }
}
