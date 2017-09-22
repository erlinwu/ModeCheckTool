using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using System.Configuration;
using System.Data;

using MahApps.Metro.Controls;

namespace WpfTools
{
    class Program : Application
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            //用户验证操作Start
            //......
            //用户验证操作Start

            //方法1           
            WpfTools.App app2 = new WpfTools.App();
            app2.InitializeComponent();
            MainWindow windows = new MainWindow();
            app2.MainWindow = windows;
            app2.Run();

            //方法2
            //App app3 = new App();
            //app3.InitializeComponent();
            //app3.StartupUri = new Uri("MainWindow.xaml", UriKind.Relative);
            //app3.Run();
        }
    }
}
