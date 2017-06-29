using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ImportXlsToDataTable
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            //log for net
            log4net.Config.XmlConfigurator.Configure();
        }
    }
}
