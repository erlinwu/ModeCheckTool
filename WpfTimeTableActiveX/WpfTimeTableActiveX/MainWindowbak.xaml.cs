using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfTimeTableActiveX
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        #region 界面事件
        //界面初始化
        private void Window_Initialized(object sender, EventArgs e)
        {
            //读取时间表记录
            Timetable_reflash();
        }


        //新增时间表记录
        public void button_timetable_new_Click(object sender, RoutedEventArgs e)
        {
            TimeTableEdit wTimeTableEdit = new TimeTableEdit();
            wTimeTableEdit.Title = "时间表新建窗口";
            wTimeTableEdit.ShowDialog();//模式，弹出！  
            //isw.Show()//无模式，弹出！  
        }
        #endregion

        #region 界面事件
        public void Timetable_reflash()
        {

        }
        #endregion
    }
}
