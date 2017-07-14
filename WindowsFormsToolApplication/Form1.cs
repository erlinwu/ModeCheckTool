using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

using System.Xml.Linq;

using System.IO;
using NPOI.HSSF.UserModel;//2007office
using NPOI.XSSF.UserModel;//xlsx
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.HSSF.Util;

//sqlitebulk 非官方的封装
using DB.SQLITE.SQLiteBulkInsert;
using System.Data.SQLite;
//ArrayList
using System.Collections;

//json
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;




namespace ImportXlsToDataTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //数据导入
        DataTable dataTableTypeList = null;//类型表导入
        DataTable dataTableDeviceList = null;//设备列表导入
        DataTable dataTableModeList = null;//模式表导入

        //database setting
        public string HOST = "";
        public string DBNAME = "";
        public string USER = "";
        public string PASSWORD = "";
        public string DBTYPE = "";
        DBLib.DBLib _dbcon;//数据库连接

        //bulk 操作用
        SQLiteBulkInsert TARGET;//
        SQLiteConnection sqliteBlukCon;//

        #region 配置文件 configfile相关
        //modeinfo表字段配置
        List<string> modeinfo_cname_list = new List<string>();
        //modeinfo表字段配置
        List<string> modeattr_cname_list = new List<string>();
        //站点名称配置列表
        List<string> stationinfo_cname_list = new List<string>();
        List<string> stationinfo_desc_list = new List<string>();
        //模式对比 dpl的模板路径
        string MODECHECK_SETFILEPATH;
        //dp的类名
        string MODESET_TYPENAME="";

        //模式对比的设备类型和设备标准值的设定，转化为json来处理
        JArray MODECHECK_JSONSET = new JArray();//
        //模式表 校验的时候，校验结果显示位置生成用关键字
        List<string> modetable_check_keywords_list = new List<string>();
        //模式配置文件 模式号的搜索关键字
        List<string> modetable_keywords_list = new List<string>();


        #endregion

        #region 界面 事件区域
        //窗体初始化
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                //读取配置文件
                this.IniConfig();
                //初始化数据库连接
                this.IniDB();

            }
            catch (Exception ex)
            {
                ShowInfo("窗体初始化错误：" + ex.Message);
            }
        }


        //选择excel导入数据库
        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = openExcelDialog();
                if (filepath!="")
                {
                    //InitializeWorkbook(@filepath);//
                    dataTableTypeList = ExcelToDataTable(@filepath, true);
                    //datatable 批量导入sqlite
                    if (dataTableTypeList != null)
                    {
                        insertDB_sqlitebulk_ex(dataTableTypeList, "devicetype");
                    }
                        
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("读取设备类表配置文件异常：" + ex.Message);
            }

            
        }
        //模式表配置导入
        private void buttonImpMode_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = openExcelDialog();
                if (filepath != "")
                {
                    //InitializeWorkbook(@filepath);//
                    dataTableTypeList = ExcelToDataTable_modeinfo(@filepath);//模式的配置表相对不规范，另外写方法导入
                    //datatable 批量导入sqlite
                    if (dataTableTypeList != null)
                    { 
                        insertDB_sqlitebulk_ex(dataTableTypeList, "modeinfo");
                    }
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("读取模式表配置文件异常：" + ex.Message);
            }
        }
        //设备清单导入
        private void buttonImpDevList_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = openExcelDialog();
                if (filepath != "")
                {
                    
                    //InitializeWorkbook(@filepath);//
                    dataTableTypeList = ExcelToDataTable(@filepath, true,2,2, 3);//
                    //datatable 批量导入sqlite
                    if (dataTableTypeList != null)
                    {
                        insertDB_sqlitebulk_ex(dataTableTypeList, "devicelist");
                    }
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("读取设备清单文件异常：" + ex.Message);
            }
        }

        private void buttonScreenReflash_Click(object sender, EventArgs e)
        {
            if (buttonScreenReflash.Text == "停止刷新")
            {
                buttonScreenReflash.Text = "开始刷新";
            }
            else
            {
                buttonScreenReflash.Text = "停止刷新";
            }
        }

        //清空日志显示
        private void buttonScreenClear_Click(object sender, EventArgs e)
        {
            richTextBoxMain.Clear();//
        }

        //1.生成界面 选择站点的combobox刷新列表
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //添加站点列表
                comboBox_stationlist.Text = "";
                comboBox_stationlist.Items.Clear();
                if (stationinfo_cname_list==null||stationinfo_cname_list.Count == 0)
                {
                    ShowInfo("站点配置信息未获取到，请重新启动。");
                    return;
                }
                else
                {
                    for (int i = 0; i < stationinfo_cname_list.Count; i++)
                    {
                        comboBox_stationlist.Items.Add(stationinfo_desc_list[i] + stationinfo_cname_list[i]);
                    }
                    comboBox_stationlist.SelectedIndex = 0;
                }


                //添加系统列表
                comboBox_syslist.Text = "";
                comboBox_syslist.Items.Clear();
                //查询数据库系统列表
                List<string> station_syslist =new List<string>();
                string sql_getpagename = "select distinct pagename from modeinfo where stationname='" + stationinfo_cname_list[comboBox_stationlist.SelectedIndex] + "';";
                station_syslist = getDataList(sql_getpagename);
                
                if (station_syslist==null || station_syslist.Count == 0)
                {
                    ShowInfo("模式子系统配置信息未获取到");
                }
                else
                {
                    for (int i = 0; i < station_syslist.Count; i++)
                    {
                        comboBox_syslist.Items.Add(stationinfo_desc_list[i] + station_syslist[i]);
                    }
                    comboBox_syslist.SelectedIndex = 0;
                }
                

            }
            catch (Exception ex)
            {
                ShowInfo("站点相关配置信息获取时出错："+ex.Message);
            }
        }

        //站点模式子系统列表刷新
        private void comboBox_stationlist_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                comboBox_syslist.Text = "";
                comboBox_syslist.Items.Clear();
                //查询数据库系统列表
                List<string> station_syslist = new List<string>();
                string sql_getpagename = "select distinct pagename from modeinfo where stationname='" + stationinfo_cname_list[comboBox_stationlist.SelectedIndex] + "';";
                station_syslist = getDataList(sql_getpagename);

                if (station_syslist == null || station_syslist.Count == 0)
                {
                    ShowInfo("模式子系统配置信息未获取到");
                }
                else
                {
                    for (int i = 0; i < station_syslist.Count; i++)
                    {
                        comboBox_syslist.Items.Add(station_syslist[i]);
                    }
                    comboBox_syslist.SelectedIndex = 0;
                }


            }
            catch (Exception ex)
            {
                ShowInfo("模式相关子系统配置信息获取出错：" + ex.Message);
            }
        }

        //清空模式数据库配置表
        private void button1_Click(object sender, EventArgs e)
        {
            string SqlString;
            SqlString = "delete from modeinfo;";
            using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
            {
               int r_Result = _dbcon.ExcuteSql(SqlString);

                if (r_Result != -1)
                {
                    ShowInfo("删除记录数" + r_Result.ToString());
                }
                else
                {
                    ShowInfo("删除记录出错");
                }
            }

        }
        //清空设备类型配置表
        private void button2_Click(object sender, EventArgs e)
        {
            string SqlString;
            SqlString = "delete from devicetype;";
            using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
            {
                int r_Result = _dbcon.ExcuteSql(SqlString);

                if (r_Result != -1)
                {
                    ShowInfo("删除记录数" + r_Result.ToString());
                }
                else
                {
                    ShowInfo("删除记录出错");
                }
            }
        }
        //清空设备清单配置表
        private void button3_Click(object sender, EventArgs e)
        {
            string SqlString;
            SqlString = "delete from devicelist;";
            using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
            {
                int r_Result = _dbcon.ExcuteSql(SqlString);

                if (r_Result != -1)
                {
                    ShowInfo("删除记录数" + r_Result.ToString());
                }
                else
                {
                    ShowInfo("删除记录出错");
                }
            }
        }

        //生成模式配置dpl文件
        private void buttonCreateModeDPL_Click(object sender, EventArgs e)
        {
            try
            {
                string stationName;
                string filename="";
                IWorkbook workbook = null;
                ISheet sheet = null;
                stationName = stationinfo_cname_list[comboBox_stationlist.SelectedIndex];
                string time_excel_t="";//系统时间格式化 "dd.MM.yyyy HH.mm.ss.fff" //23.06.2017 06:19:10.833

                //查询站点的模式页面配置
                //1. MXLRMODX_SET_1.PANELNAME 获取
                string sql_getpagename = "select distinct pagename from modeinfo where stationname='"+ stationName + "';";
                List<string> pagename_list = new List<string>();
                pagename_list = getDataList(sql_getpagename);

                if (pagename_list != null)
                {
                    //打开模板 根据DpName 自动生成点表
                    workbook = get_Workbook(MODECHECK_SETFILEPATH, ref filename);
                    //定位 # Datapoint/DpId 生成 dpnames
                    if (workbook.NumberOfSheets == 0)
                    {
                        ShowInfo("模式对比配置文件生成模板没有sheet页，无法自动生成");
                    }
                    sheet = workbook.GetSheetAt(0);
                    int rownum, columnnum;
                    if (get_Cell_RowandColumn(sheet , "# Datapoint/DpId" , out rownum , out columnnum))//取出对应dpname设置的单元格位置 rownum/columnnum（第一个）
                    {
                        rownum = rownum + 2;//下移2行开始配置dpnames
                        sheet =npoi_InsertRows(sheet, rownum, pagename_list.Count, sheet.GetRow(rownum));//插入空行 dpnames
                        if (sheet == null) return;
                        for (int i = 0; i < pagename_list.Count; i++)
                        {
                            IRow t_row = sheet.GetRow(rownum + i);
                            if (t_row == null) t_row = sheet.CreateRow(rownum + i); //行没有内容的时候，创建
                            t_row.CreateCell(0);
                            t_row.GetCell(0).SetCellValue("M"+stationName+"MODX_SET_" + (i + 1).ToString());
                            t_row.CreateCell(1);
                            t_row.GetCell(1).SetCellValue(MODESET_TYPENAME);
                            t_row.CreateCell(2);
                            t_row.GetCell(2).SetCellValue(i.ToString());
                        }
                    }



                    //写入配置到 workbook内
                    int rownum_dpvalue = -1;
                    int columnnum_dpvalue = -1;
                    if (get_Cell_RowandColumn(sheet, "# DpValue", out rownum_dpvalue, out columnnum_dpvalue))//取出对应DpValue设置的单元格位置 rownum/columnnum（第一个）
                    {
                        rownum_dpvalue = rownum_dpvalue + 2;//下移2行开始配置dpnames
                        sheet = npoi_InsertRows(sheet, rownum_dpvalue, pagename_list.Count * 5, sheet.GetRow(rownum_dpvalue));//插入空行 DpValue
                        if (sheet == null) return;
                        time_excel_t = DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff");//23.06.2017 06:19:10.833
                    }

                    //# DpValue
                    for (int i = 0; i < pagename_list.Count; i++)
                    {
                        //每个系统下面的模式名称获取
                        //2 MXLRMODX_SET_1.MODENUM 
                        string sql_getmodename = "select distinct modename from modeinfo where stationname='" + stationName + "' and pagename='" + pagename_list[i] + "';";
                        List<string> modename_list = new List<string>();
                        modename_list = getDataList(sql_getmodename);
                        //3 MXLRMODX_SET_1.DPNAMES 
                        //JNL DXT 78个dpnames
                        string sql_getdpnames = "select distinct t1.deviceid, t2.devindex ,t2.devid from modeinfo t1 left join devicelist t2 on t1.deviceid=t2.SBDM and t1.stationname=t2.stationid " +
                                                " where stationname='" + stationName + "' and pagename='" + pagename_list[i] + "';";
                        List<string> dpname_list = new List<string>();//dpnames列表
                        List<string> devid_list = new List<string>();//设备的类型列表 和配置文件比较 转换成类型编号（转换不出来的置0）
                        getDataList_sp1(sql_getdpnames, ref dpname_list, ref devid_list);

                        //4 MXLRMODX_SET_1.MODEVALUE
                        List<string> modevalue_list = new List<string>();//每种模式下面的标准值 列表
                        modevalue_list = modevalue_checkout(modename_list, stationName, pagename_list[i]);//根据站点、系统号、模式编号3个关键字查询 得到设备运行标准状态列表
                        if (modevalue_list == null) break;//返回配置空 跳出本次循环

                        //5 MXLRMODX_SET_1.CONTROLTYPE
                        List<string> controltype_list = new List<string>();//设备类型 列表
                        controltype_list = controltype_checkout(devid_list);//根据devid （TVF TDZ UFO FF1 FF2...）以及配置文件 转换出类型的编号
                        if (controltype_list == null) break;//返回配置空 跳出本次循环

                        for (int j = 0; j < modeattr_cname_list.Count; j++)//暂时5个属性配置
                        {

                            IRow t_row = sheet.GetRow(rownum_dpvalue + i * 5 + j);
                            if (t_row == null) t_row = sheet.CreateRow(rownum_dpvalue + i * 5 + j); //行没有内容的时候，创建
                            t_row.CreateCell(0);
                            t_row.GetCell(0).SetCellValue("UI (1)/0");
                            t_row.CreateCell(1);
                            t_row.GetCell(1).SetCellValue("M" + stationName + "MODX_SET_" + (i + 1).ToString() + "." + modeattr_cname_list[j].ToString());//MXLRMODX_SET_1
                            t_row.CreateCell(2);
                            t_row.GetCell(2).SetCellValue(MODESET_TYPENAME);
                            //具体配置信息输入
                            t_row.CreateCell(3);
                            switch (j)
                            {
                                case 0://系统名称 例如：DXT KTXT_XI SDT等等 
                                    t_row.GetCell(3).SetCellValue(pagename_list[i].ToString());
                                    break;
                                case 1://模式列表 例如：1100, "1101", "1102", "1103", "7101", "7102", "7103", "7104", "1124"
                                    t_row.GetCell(3).SetCellValue(string.Join(",", modename_list.ToArray()));//逗号分割
                                    break;
                                case 2://设备dpnames 例如：MXLRSDTB_TVFB001, "MXLRSDTB_TVFB002", "MXLRSDTB_TVFA001", "MXLRSDTB_TVFA002", "MXLRSDTB_TDZB001"
                                    t_row.GetCell(3).SetCellValue(string.Join(",", dpname_list.ToArray()));//逗号分割
                                    break;
                                case 3://设备在各个模式下的标准值 
                                    t_row.GetCell(3).SetCellValue(string.Join(",", modevalue_list.ToArray()));//逗号分割
                                    break;
                                case 4://设备类型 
                                    t_row.GetCell(3).SetCellValue(string.Join("", controltype_list.ToArray()));//
                                    break;
                                default:
                                    break;
                            }

                            t_row.CreateCell(4);
                            t_row.GetCell(4).SetCellValue("0x8300000000000101");
                            t_row.CreateCell(5);
                            t_row.GetCell(5).SetCellValue(time_excel_t);
                        }
                    }
                       
                    //另存excel配置文件
                    saveExcelDialog(workbook, stationName+"_MODECHECK_SET");
                    ShowInfo(stationName+"站的模式对比配置文件保存完毕。");
                }


            }
            catch (Exception ex)
            {

                ShowInfo("生成模式配置文件dpl时出错：" + ex.Message);

            }

        }
        //生成模式配置的画面panel文件
        private void buttonCreateModePanel_Click(object sender, EventArgs e)
        {

        }

        //校对模式表的图纸编号是否正确 以及设备类型提示
        private void buttonCheckModeTable_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = openExcelDialog();
                if (filepath != "")
                {
                    string stationName;
                    IWorkbook workbook = null;
                    workbook = ExcelToWorkBook_modecheck(@filepath,out stationName);
                    
                    if (workbook != null)
                    {
                        //另存excel校对文件
                        saveExcelDialog(workbook, stationName + "_MODECHECK_RESULT");
                        ShowInfo("模式表excel校准文件生成并保存完毕。");
                    }
                    workbook = null;
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("模式表校对时出错：" + ex.Message);
            }
        }

        #endregion 界面 事件区域


        #region excel 处理方法

        /// <summary>  
        /// 将设备类表和列表excel导入到datatable  
        /// </summary>  
        /// <param name="filePath">excel路径</param>  
        /// <param name="isColumnName">表格是否是否有列名</param> 
        /// <param name="numColumnName">列名所在的行（默认2）可选参数</param> 
        /// <param name="typeOfExcel">导入配置的类型 1：设备类表/2：设备列表（默认1）可选参数</param> 
        /// <param name="dataColumnStart">可选参数 数据开始的行数（默认3，设备类表从第三行开始有数据）可选参数</param> 
        /// <returns>返回datatable</returns>
        public DataTable ExcelToDataTable(string filePath, bool isColumnName,int numColumnName=2,int typeOfExcel=1,int dataColumnStart=3)//
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;

            string SYSID="";//系统编号
            string DEVID = "";//系统编号
            string STATIONID = "";//车站缩写
            string MODESYSID = "";//模式下的系统ID

            int extNum = 2;//额外添加的列 可能用文件名和sheet名称构成
            string filename;//文件名

            string tablename;//初始化bulk的列头用的表名

            bool isFirstFlag = true;//首次运行加载列头信息
            try
            {
                using (fs = File.OpenRead(filePath))//方法二 using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    //扩展名获取
                    string extension = System.IO.Path.GetExtension(filePath);//扩展名
                    // 2007版本xlsx处理
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // xls处理
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    filename = System.IO.Path.GetFileName(filePath);

                    if (workbook != null)
                    {
                        //文件名称需要处理，把系统ID取出来，设备类表取的是第一个下划线前面的英文，统一处理成大写，保存为SYSID
                        int i_t;
                        switch (typeOfExcel)
                        {
                            case 1://设备类表
                                i_t = filename.IndexOf("_");
                                //文件名一定不可以写错
                                SYSID = filename.Substring(0, i_t).ToUpper().Trim();//BAS_DeviceTypeList.xlsx PSD_DeviceTypeList.xlsx
                                extNum = 2;
                                break;
                            case 2://设备列表
                                //站和系统编号 sheet里面包含 不另外获取

                                extNum = 0;//站和系统编号 sheet里面包含
                                break;
                           
                            default:
                                break;
                        }

                        
                        //无sheet excel 跳出
                        if (workbook.NumberOfSheets == 0) {
                            this.ShowInfo("文件"+ filename + "无sheet页。。。");
                            return null;
                            
                        };

                        ///循环读取sheet，对sheet名称需要处理（无汉字，3位，小写会统一处理成大写），保存为sheetName_t
                        dataTable = new DataTable();//就初始化一个datatable，类表和设备表的数据量不是很大
                        for (int i1=0; i1 < workbook.NumberOfSheets; i1++)
                        {
                            sheet = null;//清空
                            sheet = workbook.GetSheetAt(i1);
                            
                            if (sheet != null)
                            {
                                
                                string sheetName_t =sheet.SheetName.ToUpper();//sheet名称 处理成设备类型
                                //汉字无法作为字段名称，需要删除
                                // 在 ASCII码表中，英文的范围是0-127，而汉字则是大于127
                                bool isChs = false;
                                for (int i = 0; i < sheetName_t.Length; i++)
                                {
                                    if ((int)sheetName_t[i] > 127)
                                    {
                                        isChs = true;
                                        break;
                                    }
                                        
                                }
                                if (isChs) continue;//中文sheet名称无法处理
                                bool isVaild = false;//命名非法标志
                                
                                switch (typeOfExcel)
                                {
                                    case 1://设备类表
                                        if(sheetName_t.Length!=3)//3位设备类型
                                            isVaild = true;
                                        DEVID = sheetName_t;
                                        tablename = "devicetype";
                                        break;
                                    case 2://设备列表
                                        if (sheetName_t.Length != 3)//3位站名
                                            isVaild = true;
                                        tablename = "devicelist";
                                        break;
                                    
                                    default:
                                        tablename = "";
                                        break;
                                }
                                if (isVaild) continue;//sheet命名非法，无法处理

                                int rowCount = sheet.LastRowNum;//总行数  
                                if (rowCount > 0)
                                {
                                    IRow firstRow = sheet.GetRow(numColumnName-1);//取表头所在的行
                                    int cellCount = firstRow.LastCellNum;//表头行的列数


                                    //构建datatable的列的头 只需要运行一次
                                    if (isFirstFlag) {
                                        isFirstFlag = false;
                                        if (isColumnName)//是sheet内否有表头字段名称
                                        {
                                            //表头的特殊处理
                                            //1.模式表 需要增加文件名作为站号，sheet名作为模式号
                                            //2.设备类表 文件名作为系统号，sheet名作为设备类ID
                                            //3.设备列表 文件名里面有站号和系统号（先取系统号ID），sheet名作为站号    
                                            switch (typeOfExcel)
                                            {
                                                case 1://设备类表
                                                    column = null;
                                                    column = new DataColumn("SYSID");
                                                    dataTable.Columns.Add(column);
                                                    column = null;
                                                    column = new DataColumn("DEVID");
                                                    dataTable.Columns.Add(column);
                                                    break;
                                                case 2://设备列表
                                                
                                                    break;
                                                
                                                default:
                                                    break;
                                            }
                                            for (int i = firstRow.FirstCellNum; i <= cellCount; i++)
                                            {
                                                cell = firstRow.GetCell(i);
                                                if (cell != null)
                                                {
                                                    object str_cvalue = npoi_celldeal(cell);
                                                    if (str_cvalue.ToString() != "")
                                                    {
                                                        column = new DataColumn(str_cvalue.ToString().Trim());
                                                        dataTable.Columns.Add(column);
                                                    }
                                                    else//设备清单里面 第二行是序号，字段是空的，但是后面的列还是要的。所以不跳出，continue。
                                                    {
                                                        continue;
                                                    }
                                                }
                                            }
                                            //生成bulk 参数对照表
                                            addBulkParameters(tablename, dataTable);
                                        }
                                        else//没有表头就自定义 暂时不会有类似情况出现，因为后面需要表头去对应数据表插入数据
                                        {
                                            this.ShowInfo("表的字段信息不在sheet内，无法处理");//没有表头无法处理
                                            return null;
                                        }
                                    }

                                    //dataTable填充行数据 
                                    startRow = dataColumnStart-1;//数据开始的行编号初始化
                                    //问题：1 rowCount的获取不一定对，有些行没有任何内容，也会计算进入，需要添加row.Cell.count来判断
                                    //
                                    for (int i = startRow; i <=rowCount; i++)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null&&row.Cells.Count ==0) continue;//必须加cell的count判断，否则有空记录出现

                                        dataRow = dataTable.NewRow();//dataTable中创建新行
                                        bool filter_flag = false;
                                        //设备类表和模式表特殊表头处理
                                        switch (typeOfExcel)
                                        {
                                            case 1://设备类表
                                                dataRow[0] = SYSID;
                                                dataRow[1] = DEVID;
                                                //设备类表的特殊处理 必须包含字段Type、Index否则过滤掉
                                                //后续可以考虑 关键字段的判断可以做成配置
                                                if (npoi_celldeal(row.GetCell(0)).ToString() == ""|| npoi_celldeal(row.GetCell(1)).ToString() == "")
                                                {
                                                    filter_flag = true;
                                                }

                                                break;
                                            case 2://设备列表
                                                //设备清单 必须有设备编号、车站标识、系统标识
                                                if (npoi_celldeal(row.GetCell(2)).ToString() == "" || npoi_celldeal(row.GetCell(4)).ToString() == "" || npoi_celldeal(row.GetCell(5)).ToString() == "")
                                                {
                                                    filter_flag = true;
                                                }
                                                break;
                                            
                                            default:
                                                break;
                                        }
                                        if (filter_flag) continue;//过滤掉一些不符合条件的行数据
                                        //
                                        try
                                        {
                                            if (row.FirstCellNum == -1) continue;
                                            //row.FirstCellNum第一个不是空的cell
                                            int i_ts = 0;
                                            //extNum为设备类的特殊处理 站点编号在文件头上，类别ID在sheet的NAME上。
                                            for (int j =0+ extNum; j < cellCount + extNum; j++)
                                            {
                                                //设备清单的特殊处理 第二列 序号里面的数据不用
                                                if (typeOfExcel == 2 && j - extNum == 1) {
                                                    i_ts = i_ts + 1;
                                                    continue;
                                                }
                                                
                                                //有些cell有公式
                                                cell = row.GetCell(j - extNum);//cell数据获取
                                                //有些cell前后有多余的空格
                                                dataRow[j- i_ts] = npoi_celldeal(cell).ToString().Trim();
                                                
                                            }
                                        }
                                        catch (Exception ex) {
                                            this.ShowInfo("sheet" + sheetName_t+"处理过程出错。"+ex.Message);
                                        }
                                        
                                        //处理一行添加一行数据
                                        dataTable.Rows.Add(dataRow);//dataTable数据添加一行
                                    }
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                this.ShowInfo("导入文件" + filePath + "出错。" + ex.Message);
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }
        
        //模式对比表导入处理
        public DataTable ExcelToDataTable_modeinfo(string filePath)//
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;

            string STATIONID = "";//车站缩写
            string PAGENAME = "";//画面的名称

            string filename;//文件名

            int MOSHI_RNUM=-1;//每一页里面的 模式 2个字所在的行和列的编号
            int MOSHI_CNUM=-1;
 
            try
            {
                using (fs = File.OpenRead(filePath))//方法二 using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    //扩展名获取
                    string extension = System.IO.Path.GetExtension(filePath);//扩展名
                    // 2007版本xlsx处理
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // xls处理
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    filename = System.IO.Path.GetFileName(filePath);

                    if (workbook != null)
                    {
                        //文件名称需要处理，把系统ID取出来，设备类表取的是第一个下划线前面的英文，统一处理成大写，保存为SYSID
                        int i_t;
                        //文件名一定不可以写错
                        i_t = filename.IndexOf("_");
                        STATIONID = filename.Substring(0, i_t).ToUpper();//JNL_MODE.xlsx

                        if (workbook.NumberOfSheets == 0)
                        {
                            this.ShowInfo("文件" + filename + "无sheet页。。。");
                        };

                        ///循环读取sheet，对sheet名称需要处理（无汉字，小写会统一处理成大写），保存为sheetName_t
                        dataTable = new DataTable();//就初始化一个datatable，类表和设备表的数据量不是很大
                                                    //生成bulk 参数对照表
                        foreach (string s in modeinfo_cname_list)
                        {
                            dataTable.Columns.Add(s);
                        }
                        addBulkParameters("modeinfo", dataTable);

                        for (int i1 = 0; i1 < workbook.NumberOfSheets; i1++)
                        {
                            sheet = null;//清空
                            sheet = workbook.GetSheetAt(i1);

                            if (sheet != null)
                            {

                                string sheetName_t = sheet.SheetName.ToUpper().Trim();//sheet名称 处理成设备类型
                                //汉字无法作为字段名称，需要删除
                                //在 ASCII码表中，英文的范围是0-127，而汉字则是大于127
                                //模式表 里面有的sheet有中文括号 暂时不判断了
                                //bool isChs = false;
                                //for (int i = 0; i < sheetName_t.Length; i++)
                                //{
                                //    if ((int)sheetName_t[i] > 127)
                                //    {
                                //        isChs = true;
                                //        break;
                                //    }

                                //}
                                //if (isChs) continue;//

                                //sheet名称就是模式page的名称的开头
                                PAGENAME = sheetName_t;
                                
                                int rowCount = sheet.LastRowNum;//总行数  
                                if (rowCount > 0)
                                {
                                    MOSHI_RNUM = -1;
                                    MOSHI_CNUM = -1;

                                    //遍历取出模式的行号和列号
                                    foreach (IRow r in sheet)
                                    {
                                        foreach (ICell c in r)
                                        {
                                            object str_cvalue = npoi_celldeal(c);
                                            string s_str = str_cvalue.ToString();
                                            //excel表格中 模式号单元格里面的内容（模式号）不一定连续，有可能换行，空格，因此不能直接匹配模式号3个字。
                                            bool r_flag = true;
                                            for(int i=0;i< modetable_keywords_list.Count; i++)
                                            {
                                                if(s_str.IndexOf(modetable_keywords_list[i]) < 0)
                                                {
                                                    r_flag = false;
                                                    break;
                                                }

                                            }
                                            //匹配关键字成功
                                            if (r_flag) {
                                                int rowSpan_in = 0;
                                                int columnSpan_in = 0;
                                                NPOI.ExcelExtension.IsMergeCell(sheet, c.RowIndex, c.ColumnIndex,out rowSpan_in,out columnSpan_in);
                                                MOSHI_RNUM = c.RowIndex+ rowSpan_in-1;
                                                MOSHI_CNUM = c.ColumnIndex+ columnSpan_in-1;
                                                break;
                                            }
                                        }
                                        if (MOSHI_RNUM != -1 && MOSHI_CNUM != -1) break;
                                    }
                                    //
                                    if(MOSHI_RNUM>=0&& MOSHI_CNUM >= 0)//取出模式号
                                    {
                                        List<string> modeid_list = new List<string>();
                                        List<string> devid_list  = new List<string>();
                                        //根据模式的行列 取出所有的模式号
                                        for (int i = MOSHI_RNUM+1; i < rowCount; i++)//总行数以内
                                        {
                                            //20170712 需要关注一下是否是合并单元格 否则会出现遗漏的情况
                                            int rowSpan;
                                            int columnSpan;
                                            bool IsMerge= NPOI.ExcelExtension.IsMergeCell(sheet.GetRow(i).GetCell(MOSHI_CNUM), out rowSpan, out columnSpan);

                                            //
                                            string str_cvalue = sheet.GetRow(i).GetCell(MOSHI_CNUM).ToString();
                                            int out_i;
                                            
                                            if (str_cvalue!=""&& int.TryParse(str_cvalue, out out_i))
                                            {
                                                modeid_list.Add(str_cvalue.Trim());
                                            }
                                            else if (IsMerge)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                break;
                                            }


                                        }
                                        //取出模式对应的所有设备编号 取到空结束
                                        //20170710 出现结尾有空格的情况，需要考虑剔除首尾空格 
                                        for (int i = MOSHI_CNUM+1; i < sheet.GetRow(MOSHI_RNUM).LastCellNum; i++)
                                        {
                                            string str_devid = sheet.GetRow(MOSHI_RNUM).GetCell(i).ToString();
                                            //20170712 考虑有可能2列合并的可能
                                            int rowSpan;
                                            int columnSpan;
                                            bool IsMerge = NPOI.ExcelExtension.IsMergeCell(sheet.GetRow(MOSHI_RNUM).GetCell(i), out rowSpan, out columnSpan);


                                            if (str_devid != "")
                                            {
                                                devid_list.Add(str_devid.Trim());
                                            }
                                            else if (IsMerge)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        //取出所有的状态填充datatable
                                        int rowAdd = 0;//合并行所带来的行偏移
                                        for (int i = 0; i < modeid_list.Count; i++)
                                        {
                                            int colAdd = 0;//合并列所带来的列偏移
                                            int rowSpan=1;
                                            int columnSpan=1;
                                            for (int j = 0; j < devid_list.Count; j++)
                                            {
                                                //
                                                dataRow = dataTable.NewRow();//dataTable中创建新行
                                                dataRow[0] = STATIONID;
                                                dataRow[1] = PAGENAME;
                                                dataRow[2] = modeid_list[i];
                                                dataRow[3] = devid_list[j];
                                               
                                                //格式处理
                                                //模式对比 有可能比对值会有小写 处理掉
                                                dataRow[4] = npoi_celldeal(sheet.GetRow(MOSHI_RNUM + 1 + i+ rowAdd).GetCell(MOSHI_CNUM + j + 1+ colAdd)).ToString().ToUpper();

                                                //20170712 合并单元格处理 （处理有可能带来的行以及列的合并）
                                                if (NPOI.ExcelExtension.IsMergeCell(sheet.GetRow(MOSHI_RNUM + 1 + i).GetCell(MOSHI_CNUM + j + 1), out rowSpan, out columnSpan))
                                                {
                                                    colAdd = colAdd + columnSpan - 1;//如果是合并的单元格，列的
                                                }

                                                //处理一行添加一行数据
                                                dataTable.Rows.Add(dataRow);//dataTable数据添加一行
                                            }
                                            //处理行的偏移量
                                            rowSpan = 1;
                                            columnSpan = 1;
                                            if (NPOI.ExcelExtension.IsMergeCell(sheet.GetRow(MOSHI_RNUM + 1 + i+ rowAdd).GetCell(MOSHI_CNUM), out rowSpan, out columnSpan))
                                            {
                                                rowAdd = rowAdd + rowSpan - 1;
                                            }
                                        }
                                    }     
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                this.ShowInfo("导入文件" + filePath + "出错。" + ex.Message);
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        //模式表校对 生成workbook返回输出excel
        public IWorkbook ExcelToWorkBook_modecheck(string filePath,out string STATIONID)//
        {
            
            FileStream fs = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;


            STATIONID = "";//车站缩写
            string PAGENAME = "";//画面的名称

            string filename;//文件名

            int MOSHI_RNUM = -1;//每一页里面的 模式 2个字所在的行和列的编号
            int MOSHI_CNUM = -1;

            int MSDB_RNUM = -1;//模式对比位置
            int MSDB_CNUM = -1;


            try
            {
                using (fs = File.OpenRead(filePath))//方法二 using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    //扩展名获取
                    string extension = System.IO.Path.GetExtension(filePath);//扩展名
                    // 2007版本xlsx处理
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // xls处理
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    filename = System.IO.Path.GetFileName(filePath);

                    if (workbook != null)
                    {
                        //校对文字白色底色
                        ICellStyle s_white = workbook.CreateCellStyle();
                        s_white.FillForegroundColor = HSSFColor.White.Index;
                        s_white.FillPattern = FillPattern.SolidForeground;
                        //校对文字红色底色
                        ICellStyle s_red = workbook.CreateCellStyle();
                        s_red.FillForegroundColor = HSSFColor.Red.Index;
                        s_red.FillPattern = FillPattern.SolidForeground;

                        //文件名称需要处理，把系统ID取出来，设备类表取的是第一个下划线前面的英文，统一处理成大写，保存为SYSID
                        int i_t;
                        //文件名一定不可以写错
                        i_t = filename.IndexOf("_");
                        STATIONID = filename.Substring(0, i_t).ToUpper();//JNL_MODE.xlsx

                        if (workbook.NumberOfSheets == 0)
                        {
                            this.ShowInfo("文件" + filename + "无sheet页。。。");
                        };


                        for (int i1 = 0; i1 < workbook.NumberOfSheets; i1++)
                        {
                            sheet = null;//清空
                            sheet = workbook.GetSheetAt(i1);

                            if (sheet != null)
                            {
                                string sheetName_t = sheet.SheetName.ToUpper().Trim();//sheet名称 处理成设备类型
                               
                                //sheet名称就是模式page的名称的开头
                                PAGENAME = sheetName_t;

                                int rowCount = sheet.LastRowNum;//总行数  
                                if (rowCount > 0)
                                {
                                    MOSHI_RNUM = -1;
                                    MOSHI_CNUM = -1;

                                    //遍历取出“模式号”这个单元格的所在位置
                                    foreach (IRow r in sheet)
                                    {
                                        foreach (ICell c in r)
                                        {
                                            object str_cvalue = npoi_celldeal(c);
                                            string s_str = str_cvalue.ToString();
                                            //excel表格中 模式号单元格里面的内容（模式号）不一定连续，有可能换行，空格
                                            bool r_flag = true;
                                            for (int i = 0; i < modetable_keywords_list.Count; i++)
                                            {
                                                if (s_str.IndexOf(modetable_keywords_list[i]) < 0)
                                                {
                                                    r_flag = false;
                                                    break;
                                                }
                                                
                                            }
                                            //匹配关键字成功
                                            if (r_flag)
                                            {
                                                int rowSpan_in = 0;
                                                int columnSpan_in = 0;
                                                NPOI.ExcelExtension.IsMergeCell(sheet, c.RowIndex, c.ColumnIndex, out rowSpan_in, out columnSpan_in);
                                                MOSHI_RNUM = c.RowIndex + rowSpan_in - 1;
                                                MOSHI_CNUM = c.ColumnIndex + columnSpan_in - 1;
                                                break;
                                            }
                                        }
                                        //
                                        if (MOSHI_RNUM != -1 && MOSHI_CNUM != -1) break;
                                        
                                    }
                                    
                                    //
                                    MSDB_RNUM = -1;
                                    MSDB_CNUM = -1;
                                    //遍历取出“模式对比”这个单元格的所在位置
                                    foreach (IRow r in sheet)
                                    {
                                        foreach (ICell c in r)
                                        {
                                            object str_cvalue = npoi_celldeal(c);
                                            string s_str = str_cvalue.ToString();
                                            //
                                            bool r_flag = true;
                                            for (int i = 0; i < modetable_check_keywords_list.Count; i++)
                                            {
                                                if (s_str.IndexOf(modetable_check_keywords_list[i]) < 0)
                                                {
                                                    r_flag = false;
                                                    break;
                                                }

                                            }
                                            //匹配关键字成功
                                            if (r_flag)
                                            {
                                                int rowSpan_in = 0;
                                                int columnSpan_in = 0;
                                                NPOI.ExcelExtension.IsMergeCell(sheet, c.RowIndex, c.ColumnIndex, out rowSpan_in, out columnSpan_in);
                                                MSDB_RNUM = c.RowIndex + rowSpan_in - 1;
                                                MSDB_CNUM = c.ColumnIndex + columnSpan_in - 1;
                                                break;
                                            }
                                        }
                                        //
                                        if (MSDB_RNUM != -1 && MSDB_CNUM != -1) break;

                                    }
                                    //
                                    if (MOSHI_RNUM >= 0 && MOSHI_CNUM >= 0)//取出模式号
                                    {
                                        
                                        List<string> devid_list = new List<string>();//设备图纸编号列表

                                        //取出模式对应的所有设备编号 取到空结束
                                        //20170710 出现结尾有空格的情况，需要考虑剔除首尾空格 
                                        for (int i = MOSHI_CNUM + 1; i < sheet.GetRow(MOSHI_RNUM).LastCellNum; i++)
                                        {
                                            string str_devid = sheet.GetRow(MOSHI_RNUM).GetCell(i).ToString();
                                            //20170712 考虑有可能2列合并的可能
                                            int rowSpan;
                                            int columnSpan;
                                            bool IsMerge = NPOI.ExcelExtension.IsMergeCell(sheet.GetRow(MOSHI_RNUM).GetCell(i), out rowSpan, out columnSpan);


                                            if (str_devid != "")
                                            {
                                                devid_list.Add(str_devid.Trim());
                                            }
                                            else if (IsMerge)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }

                                        //根据取出来的设备编号，去匹配设备清单，抽取出 设备类型 和 设备tagnames清单
                                        List<string> devclass_list   = new List<string>();
                                        List<string> devtagname_list = new List<string>();
                                        //
                                        string str_tmp="";
                                        List<string> tmp_list = new List<string>();
                                        foreach (string s in devid_list)
                                        {
                                            tmp_list.Add('"' + s + '"');
                                        }
                                        str_tmp = string.Join(",", tmp_list);
                                        string sql_getdata = "select devindex, devid, sbdm from devicelist where stationid = '" + STATIONID + "' and sbdm in(" + str_tmp + ");";

                                        using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
                                        {
                                            DataSet ds_Result = _dbcon.GetData(sql_getdata);
                                            if (ds_Result == null|| ds_Result.Tables.Count == 0|| ds_Result.Tables[0].Rows.Count == 0)
                                            {
                                                ShowInfo("校对模式表文件，查询设备列表时没有记录。");
                                                //return null;
                                            }
                                            //处理ds到 设备类型 和 tagnames清单 2个list
                                            //for (int i = 0; i < ds_Result.Tables[0].Rows.Count; i++)
                                            //{
                                            //    devclass_list.Add(ds_Result.Tables[0].Rows[i][1].ToString());
                                            //    devtagname_list.Add(ds_Result.Tables[0].Rows[i][0].ToString());
                                            //}
                                            //
                                            string s1, s2;
                                            foreach(string devid in devid_list)
                                            {
                                                s1 = "";
                                                s2 = "";
                                                if (ds_Result != null && ds_Result.Tables.Count > 0 && ds_Result.Tables[0].Rows.Count > 0)
                                                {
                                                    for (int i = 0; i < ds_Result.Tables[0].Rows.Count; i++)
                                                    {
                                                        if (devid == ds_Result.Tables[0].Rows[i][2].ToString())
                                                        {
                                                            s1 = ds_Result.Tables[0].Rows[i][0].ToString();
                                                            s2 = ds_Result.Tables[0].Rows[i][1].ToString();
                                                        }
                                                    }
                                                }
                                                    
                                                devclass_list.Add(s1);
                                                devtagname_list.Add(s2);
                                            }
                                        }

                                        //取出所有的状态填充datatable
                                        sheet.CreateRow(MSDB_RNUM + 1);
                                        sheet.CreateRow(MSDB_RNUM + 2);
                                        //说明文字
                                        sheet.GetRow(MSDB_RNUM + 1).CreateCell(MSDB_CNUM + 1);
                                        sheet.GetRow(MSDB_RNUM + 1).GetCell(MSDB_CNUM + 1).SetCellValue("设备编号");
                                        sheet.GetRow(MSDB_RNUM + 1).GetCell(MSDB_CNUM + 1).CellStyle = s_white;
                                        sheet.GetRow(MSDB_RNUM + 2).CreateCell(MSDB_CNUM + 1);
                                        sheet.GetRow(MSDB_RNUM + 2).GetCell(MSDB_CNUM + 1).SetCellValue("设备标识");
                                        sheet.GetRow(MSDB_RNUM + 2).GetCell(MSDB_CNUM + 1).CellStyle = s_white;
                                        for (int i = 0; i < devid_list.Count; i++)
                                        {
                                            //class
                                            sheet.GetRow(MSDB_RNUM + 1).CreateCell(MSDB_CNUM + 2 + i);//未匹配出来的红色标记
                                            if (devclass_list[i]=="") { 
                                                sheet.GetRow(MSDB_RNUM + 1).GetCell(MSDB_CNUM + 2 + i).SetCellValue("无");
                                                sheet.GetRow(MSDB_RNUM + 1).GetCell(MSDB_CNUM + 2 + i).CellStyle = s_red;
                                            }
                                            else
                                            {
                                                sheet.GetRow(MSDB_RNUM + 1).GetCell(MSDB_CNUM + 2 + i).SetCellValue(devclass_list[i]);
                                                sheet.GetRow(MSDB_RNUM + 1).GetCell(MSDB_CNUM + 2 + i).CellStyle = s_white;
                                            }
                                            
                                            //tagname
                                            sheet.GetRow(MSDB_RNUM + 2).CreateCell(MSDB_CNUM + 2 + i);//未匹配出来的红色标记
                                            if (devclass_list[i] == "")
                                            {
                                                sheet.GetRow(MSDB_RNUM + 2).GetCell(MSDB_CNUM + 2 + i).SetCellValue("无");
                                                sheet.GetRow(MSDB_RNUM + 2).GetCell(MSDB_CNUM + 2 + i).CellStyle = s_red;
                                            }
                                            else
                                            {
                                                sheet.GetRow(MSDB_RNUM + 2).GetCell(MSDB_CNUM + 2 + i).SetCellValue(devtagname_list[i]);
                                                sheet.GetRow(MSDB_RNUM + 2).GetCell(MSDB_CNUM + 2 + i).CellStyle = s_white;
                                            }
                                            
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                }
                return workbook;
            }
            catch (Exception ex)
            {
                this.ShowInfo("校对模式表文件：" + filePath + "出错。" + ex.Message);
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }


        //npoi cell 格式处理 
        public object npoi_celldeal(ICell cell)
        {
            object ret_cell = "";
            try
            {
                if (cell == null)
                {
                    ret_cell = "";
                }
                else
                {
                    //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,) 
                    //单元格数据出来里，这部分后续要关注 
                    switch (cell.CellType)
                    {
                        case CellType.Blank:
                            ret_cell = "";
                            break;
                        case CellType.Numeric:
                            short format = cell.CellStyle.DataFormat;
                            //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                            if (format == 14 || format == 31 || format == 57 || format == 58)
                                ret_cell = cell.DateCellValue;
                            else
                                ret_cell = cell.NumericCellValue;
                            break;
                        case CellType.String:
                            ret_cell = cell.StringCellValue;
                            break;
                        case CellType.Formula://公式处理，暂时可以用
                            ret_cell = cell.StringCellValue;
                            break;
                    }
                }

            }
            catch(Exception ex)
            {
                return "error cell";
            }
            return ret_cell;
            
           
        }

        //根据文件路径 获取IWorkbook 返回文件名
        public IWorkbook get_Workbook(string filePath, ref string filename)
        {
            filename = "";
            FileStream fs = null;
            IWorkbook workbook = null;

            try
            {
                using (fs = File.OpenRead(filePath))//方法二 using (fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    //扩展名获取
                    string extension = System.IO.Path.GetExtension(filePath);//扩展名
                    // 2007版本xlsx处理
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // xls处理
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    filename = System.IO.Path.GetFileName(filePath);

                    if (workbook != null)
                    {
                        return workbook;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch(Exception ex)
            {
                ShowInfo("通过路径获取excel文件出错：" + ex.Message);
                return null;
            }
                        
        }

        //根据单元格内容 返回第一个符合要求的行，列
        public bool get_Cell_RowandColumn(ISheet sheet,string cellvalue,out int rownum,out int columnnum)
        {
            rownum = 0;
            columnnum = 0;
            try
            {
                ///遍历取出模式
                foreach (IRow r in sheet)
                {
                    foreach (ICell c in r)
                    {
                        object str_cvalue = npoi_celldeal(c);
                        int i_m = -1;
                        string s_str = str_cvalue.ToString();
                        i_m = s_str.IndexOf(cellvalue);
                        if (i_m >= 0)
                        {
                            rownum    = c.RowIndex;
                            columnnum = c.ColumnIndex;
                            break;
                        }
                    }

                }
                return true;
            }
            catch(Exception ex)
            {
                ShowInfo("获取sheet中的单元格 " + cellvalue + "时出错："+ex.Message);
                return false;
            }
        }

        //插入多行
        //startrownum 起始行的索引
        private ISheet npoi_InsertRows(ISheet sheet, int startrownum, int insertrownum, IRow sourcerow)
        {
            try
            {
                IRow targetRow = null;
                targetRow=sheet.CreateRow(startrownum + 1);//一定要新建一行，防止原始有内容的行被覆盖
                targetRow.CreateCell(0).SetCellValue("new");

                for (int i = startrownum; i < startrownum + insertrownum-1 ; i++)
                { 
                    ICell sourceCell = null;
                    ICell targetCell = null;

                    //不能使用 CreateRow 插入不了，会覆盖原来的行
                    targetRow = sheet.CopyRow(startrownum, i + 1);//第一次copy 会覆盖有内容的行
                    //
                    
                    //设置格式
                    for (int m = sourcerow.FirstCellNum; m < sourcerow.LastCellNum; m++)
                    {
                        sourceCell = sourcerow.GetCell(m);
                        if (sourceCell == null)
                            continue;
                        targetCell = targetRow.CreateCell(m);
                        //编码的继承 encode 无
                        targetCell.CellStyle = sourceCell.CellStyle;
                        targetCell.SetCellType(sourceCell.CellType);
                    }
                }

                return sheet;
            }
            catch(Exception ex)
            {
                ShowInfo("excel插入行时出错："+ex.Message);
                return null;
            }
        }


        #endregion excel 处理方法


        #region 公共方法
        //刷新消息显示信息
        public void ShowInfo(string Message)
        {
            this.Invoke(new Action<string>(this._WriteLog), new object[1] { Message });
        }

        //刷新界面信息
        public void _WriteLog(string Message)
        {
            if (buttonScreenReflash.Text == "停止刷新")
            {
                if (richTextBoxMain.Lines.Count() > 1000)
                {
                    richTextBoxMain.Clear();
                }
                richTextBoxMain.Text = richTextBoxMain.Text.Insert(0, "[" + System.DateTime.Now + "]:" + Message + "\n");
                //richTextBoxMain.AppendText("[" + System.DateTime.Now + "]:" + Message + "\n");
            }
        }

        //配置文件读取
        public void IniConfig()
        {
            string sPath = Application.StartupPath + "\\configfile.xml";//配置文件地址 \\和//都可以加在配置文件前面
            try
            {
                if (!File.Exists(sPath))
                {
                    this.ShowInfo("配置文件丢失，程序法无正常开启");
                }
                this.ShowInfo("读取配置文件");
                //读取配置文件
                XDocument xdConfigFile = XDocument.Load(sPath);
                XElement xeRoot = xdConfigFile.Descendants("root").First();

                //获取数据库配置信息
                XElement xedatabaseRoot = xeRoot.Descendants("dbinfo").First();
                //由于是sqlite数据库 文件型数据库 需要全路径
                HOST = System.Environment.CurrentDirectory+xedatabaseRoot.Attribute("HOST").Value.ToString();
                DBNAME = xedatabaseRoot.Attribute("DBNAME").Value.ToString();
                USER = xedatabaseRoot.Attribute("USER").Value.ToString();
                PASSWORD = xedatabaseRoot.Attribute("PASSWORD").Value.ToString();
                DBTYPE = xedatabaseRoot.Attribute("DBTYPE").Value.ToString();

                //模式配置的dp 类名称 EMCS_MODECHECK_SETTING
                XElement xeModeTypeNameRoot = xeRoot.Descendants("modecheck_otherset").First();
                MODESET_TYPENAME= xeModeTypeNameRoot.Attribute("typename").Value.ToString();

                //读取模式配置表的字段名称
                XElement xeModeinfoRoot = xeRoot.Descendants("modeinfo_table").First();
                if (xeModeinfoRoot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModeinfoRoot.Descendants("cname"))
                    {
                        modeinfo_cname_list.Add(xePort.Attribute("cname").Value.ToString());
                    }
                }
                if (modeinfo_cname_list.Count == 0)
                {
                    this.ShowInfo("模式信息配置表字段相关配置丢失，请重新修改配置文件");
                }

                //读取模式配置表的字段名称
                XElement xeStationinfoRoot = xeRoot.Descendants("stationinfo").First();
                if (xeStationinfoRoot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeStationinfoRoot.Descendants("cname"))
                    {
                        stationinfo_cname_list.Add(xePort.Attribute("cname").Value.ToString());
                        stationinfo_desc_list.Add(xePort.Attribute("desc").Value.ToString());
                    }
                }
                if (stationinfo_cname_list.Count == 0)
                {
                    this.ShowInfo("站点信息配置表字段相关配置丢失，请重新修改配置文件");
                }

                //
                //获取模式检查表的配置文件dpl的生成模板
                XElement xeModecheckSetDPLRoot = xeRoot.Descendants("modecheckset").First(); 
                MODECHECK_SETFILEPATH = System.Environment.CurrentDirectory + xeModecheckSetDPLRoot.Attribute("filepath").Value.ToString();
                if(MODECHECK_SETFILEPATH=="")
                    this.ShowInfo("模式配置模板文件未获取，无法生成模式配置dpl文件。请配置文件中的路径是否正确。");

                //获取模式对比的类型和标准值的配置文件
                XElement xeModecheckJsonsetRoot = xeRoot.Descendants("modecheck_jsonset").First();
                if (xeModecheckJsonsetRoot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModecheckJsonsetRoot.Descendants("c"))
                    {
                        JArray jsonset_arr=new JArray();
                        JObject jsonset_d;
                        foreach (XElement xeD in xePort.Descendants("d"))
                        {
                            jsonset_d = new JObject(
                                new JProperty("cname", xeD.Attribute("cname").Value.ToString()),
                                new JProperty("cid", xeD.Attribute("cid").Value.ToString()),
                                new JProperty("desc",xeD.Attribute("desc").Value.ToString())
                                );
                            jsonset_arr.Add(jsonset_d);
                        };

                        JObject jsonset_t = new JObject(
                                new JProperty("cname", xePort.Attribute("cname").Value.ToString()),
                                new JProperty("cid", xePort.Attribute("cid").Value.ToString()),
                                new JProperty("desc", xePort.Attribute("desc").Value.ToString()),
                                new JProperty("rows", jsonset_arr)
                                );
                        MODECHECK_JSONSET.Add(jsonset_t);
                    }
                }

                //模式配置dp的属性字段
                XElement xeModeAttroot = xeRoot.Descendants("modecheck_attr").First();
                if (xeModeAttroot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModeAttroot.Descendants("cname"))
                    {
                        modeattr_cname_list.Add(xePort.Attribute("cname").Value.ToString());
                    }
                }
                if (modeattr_cname_list.Count == 0)
                {
                    this.ShowInfo("模式配置dp的属性字段丢失，请重新修改配置文件");
                }

                //模式表excel 文件 搜索模式号的关键字
                XElement xeModeKeyWordroot = xeRoot.Descendants("modetable_msh").First();
                if (xeModeKeyWordroot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModeKeyWordroot.Descendants("cname"))
                    {
                        modetable_keywords_list.Add(xePort.Attribute("cname").Value.ToString());
                    }
                }
                if (modetable_keywords_list.Count == 0)
                {
                    this.ShowInfo("模式表excel文件中，搜索关键字的配置信息丢失。请查看并修改配置文件。");
                }

                //
                XElement xeModeTableCheckKeyWordroot = xeRoot.Descendants("modetable_check").First();
                if (xeModeTableCheckKeyWordroot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModeTableCheckKeyWordroot.Descendants("cname"))
                    {
                        modetable_check_keywords_list.Add(xePort.Attribute("cname").Value.ToString());
                    }
                }
                if (modetable_check_keywords_list.Count == 0)
                {

                    this.ShowInfo("模式表校对用配置modetable_check丢失。请查看并修改配置文件。");
                }



                //var jsonset = JsonConvert.SerializeObject(MODECHECK_JSONSET);
                //this.ShowInfo("模式配置文件接送对象：" + jsonset);
                this.ShowInfo("读取配置文件成功");
            }
            catch (Exception ex)
            {
                this.ShowInfo("读取配置文件异常：" + ex.Message);

            }
        }

        //获取单列记录
        public List<string> getDataList(string sql_getpagename)
        {
            try
            {
                List<string> pagename_list = new List<string>();
                using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
                {
                    DataSet ds_Result = _dbcon.GetData(sql_getpagename);

                    if (ds_Result != null)
                    {
                        if (ds_Result != null && ds_Result.Tables.Count != 0 && ds_Result.Tables[0].Rows.Count != 0)
                        {
                            for (int i = 0; i < ds_Result.Tables[0].Rows.Count; i++)
                            {
                                pagename_list.Add(ds_Result.Tables[0].Rows[i][0].ToString());
                            }
                            //页面配置数量大于0
                            if (pagename_list.Count > 0)
                            {
                                return pagename_list;
                            }
                            else
                            {
                                return null;
                            }
                        }
                        else
                        {
                            ShowInfo("查询语句未查到数据："+ sql_getpagename);
                            return null;
                        }
                    }
                    else
                    {
                        ShowInfo("查询语句执行出错："+ sql_getpagename);
                        return null;
                    }
                }
            }
            catch(Exception ex)
            {
                ShowInfo("单条记录查询出错..."+ex.Message);
                return null;
            }
        }

        //模式对比 获取dpnames和设备类型
        public int getDataList_sp1(string getDataList,ref List<string> ret_dpnames_list,ref List<string> ret_devid_list)
        {
            int ret = -1;//执行出错返回-1

            try
            {
                
                using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
                {
                    DataSet ds_Result = _dbcon.GetData(getDataList);

                    if (ds_Result != null)
                    {
                        if (ds_Result != null && ds_Result.Tables.Count != 0 && ds_Result.Tables[0].Rows.Count != 0)
                        {
                            for (int i = 0; i < ds_Result.Tables[0].Rows.Count; i++)
                            {
                                ret_dpnames_list.Add(ds_Result.Tables[0].Rows[i][1].ToString());
                                ret_devid_list.Add(ds_Result.Tables[0].Rows[i][2].ToString());
                            }
                            //页面配置数量大于0
                            if (ret_dpnames_list.Count > 0)
                            {
                                return 1;
                            }
                            else
                            {
                                return -1;
                            }
                        }
                        else
                        {
                            ShowInfo("查询语句未查到数据：" + getDataList);
                            return -1;
                        }
                    }
                    else
                    {
                        ShowInfo("查询语句执行出错：" + getDataList);
                        return -1;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowInfo("单条记录查询出错..."+ex.Message);
                return -1;
            }

            
        }

        //根据站点、系统号、模式编号3个关键字查询 得到设备运行标准状态列表
        public List<string> modevalue_checkout(List<string> modename_list,string stationName,string pagename)
        {
            List<string> ret_list = new List<string>();

            try
            {

                using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
                {

                    foreach(string mn in modename_list)
                    {
                        string getDataList= "select t1.standardvalue,t2.devid from modeinfo t1 left join devicelist t2 on t1.deviceid=t2.SBDM and t1.stationname=t2.stationid " +
                                                " where stationname='" + stationName + "' and pagename='" + pagename + "' and modename='"+ mn +"';";
                        
                        DataSet ds_Result = _dbcon.GetData(getDataList);

                        if (ds_Result != null)
                        {
                            if (ds_Result != null && ds_Result.Tables.Count != 0 && ds_Result.Tables[0].Rows.Count != 0)
                            {
                                List<string> ret_t = new List<string>();
                                for (int i = 0; i < ds_Result.Tables[0].Rows.Count; i++)
                                {
                                    string devid = ds_Result.Tables[0].Rows[i][1].ToString();//DDT UOF...
                                    string standardvelue = ds_Result.Tables[0].Rows[i][0].ToString();// X O - 空字符串都有 需要转码
                                    //
                                    bool flag_nocid = true;
                                    foreach (JObject jo in MODECHECK_JSONSET)//一个配置为一个json object
                                    {
                                        //SelectToken 方法使用
                                        if ((string)jo.SelectToken("cname") == devid)
                                        {
                                            IEnumerable<JToken> stv = jo.SelectTokens("$..rows[?(@.cname=='"+ standardvelue + "')]");
                                            if (stv.Count() == 0) break;

                                            ret_t.Add(stv.ToList()[0]["cid"].ToString());
                                            flag_nocid = false;
                                            break;//获取到cname
                                        }
                                    }
                                    if (flag_nocid)//未获取到canme 置零
                                    {
                                        ret_t.Add("0");
                                    }
                                    //
                                    
                                }
                                ret_list.Add(string.Join("", ret_t));


                            }
                            else
                            {
                                ShowInfo("查询语句未查到数据：" + getDataList);
                                return null;
                            }
                        }
                        else
                        {
                            ShowInfo("查询语句执行出错：" + getDataList);
                            return null;
                        }
                    }
                    //数量大于0
                    if (ret_list.Count > 0)
                    {
                        return ret_list;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowInfo("获取标准状态值列表出错：" + ex.Message);
                return null;
            }


        }

        //保存excel窗口
        //filename 默认文件名
        public void saveExcelDialog(IWorkbook workbook,string filename)
        {
            try
            {
                SaveFileDialog dialog = new SaveFileDialog();
                // 默认扩展名  
                dialog.DefaultExt = ".xls";
                // 默认文件名  
                dialog.FileName = filename;

                dialog.Title = "请选择文件夹";
                dialog.Filter = "所有文件(*.xls,*.xlsx,*.txt)|*.xls;*.xlsx;*.txt";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    System.IO.MemoryStream streamMemory = new System.IO.MemoryStream();
                    workbook.Write(streamMemory);
                    byte[] data = streamMemory.ToArray();

                    char[] datachar = new ASCIIEncoding().GetChars(data);
                    //char[] datachar = Encoding.UTF8.GetChars(data);
                    //char[] datachar = Encoding.Unicode.GetChars(data);

                    string strFileName = dialog.FileName;
                    // 实例化一个文件流  
                    FileStream streamFile = new FileStream(strFileName, FileMode.Create);
                    
                    // 开始写入  
                    streamFile.Write(data, 0, data.Length);

                    // 清空缓冲区 写入文件
                    streamFile.Flush();
                    //关闭流  
                    streamFile.Close();
                    workbook = null;
                    streamMemory.Close();
                    streamMemory.Dispose();

                    ////第二种方法 直接转码保存txt文件失败，无法实现
                    ////var utf8WithBom = new System.Text.UTF8Encoding(true);  // 用true来指定包含bom
                    //var utf8WithoutBom = new System.Text.UTF8Encoding(false);

                    //StreamWriter streamfile2 = new StreamWriter(strFileName, false, utf8WithoutBom);
                    //streamfile2.Write(datachar);
                    //streamfile2.Flush();
                    //streamfile2.Close();
                    //workbook = null;
                    //streamMemory.Close();
                    //streamMemory.Dispose();

                }

            }
            catch (Exception ex)
            {
                ShowInfo("到处excel文件出错："+ex.Message);
            }
        }

        //打开excel窗口
        public string openExcelDialog()
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = false;//多文件
                dialog.Title = "请选择文件夹";
                dialog.Filter = "所有文件(*.xls,*.xlsx)|*.xls;*.xlsx";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return dialog.FileName;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        //根据devid （TVF TDZ UFO FF1 FF2...）以及配置文件 转换出类型的编号
        public List<string> controltype_checkout(List<string> devid_list)
        {
            List<string> ret_devnu = new List<string>();
            try
            {
                foreach (string d in devid_list)//TVF FF1 FF2 UOF
                {
                    bool flag_nocid = true;
                    foreach (JObject jo in MODECHECK_JSONSET)//一个配置为一个json object
                    {
                        //SelectToken 方法使用
                        if ((string)jo.SelectToken("cname") == d)
                        {
                            ret_devnu.Add((string)jo.SelectToken("cid"));
                            flag_nocid = false;
                            break;//获取到cname
                        }
                    }
                    if (flag_nocid)//未获取到canme 置零
                    {
                        ret_devnu.Add("0");
                    }
                }
                if(devid_list.Count== ret_devnu.Count)
                {
                    return ret_devnu;//获取配置数量正确返回结果
                }
                else
                {
                    return null;
                }
            }
            catch(Exception ex)
            {
                ShowInfo("转换设备类型编号出错:" + ex.Message);
                return null;
            }
            
        }

        #endregion 公共方法


        #region  数据库相关

        //sqlite不支持bulk方式 需要自己处理，用事务提交

        /// <summary>
        /// 初始化数据库连接
        /// </summary>
        public void IniDB()
        {
            try
            {
                //暂时用bulk sqlite类操作数据库
                //_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE));
                //int a = _dbcon.trycon();
                //if (a > 0)
                //{
                //    this.ShowInfo("数据库初始化中，测试打开数据库成功");
                //}
                //else
                //{
                //    this.ShowInfo("数据库初始化中，测试打开数据库失败");
                //}
                //
                string constr="";
                constr = "Data Source=" + HOST + ";Version=3;Pooling=False;Max Pool Size=100;";//连接密码 Password="";设置连接池 Pooling=False;Max Pool Size=100;//只读连接Read Only=false";
                sqliteBlukCon = new SQLiteConnection(constr);
                try
                {
                    if (sqliteBlukCon.State != ConnectionState.Open)
                    {
                        sqliteBlukCon.Open();
                    }
                    this.ShowInfo("数据库初始化中，测试打开数据库成功");
                }
                catch
                {
                    this.ShowInfo("数据库初始化中，测试打开数据库失败"); ;
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("数据库初始化出错"+ex.Message);
            }
        }

        //批量数据导入sqlite 自己封装的bulk类（sqlite不支持 SqlBulkCopy 操作）
        public void insertDB_sqlitebulk_ex(DataTable dt,string tablename)
        {
            try { 
            //用自定义sqlite类操作
            SQLiteBulkInsert target;
            //Close connect to verify records were not inserted
            sqliteBlukCon.Close();
            //target = new SQLiteBulkInsert(sqliteBlukCon, tablename);
            target = TARGET;
            sqliteBlukCon.Open();
            target.CommitMax = 10000;//1w条数据 限制 默认值1w

            //Insert less records than commitmax
            foreach(DataRow r in dt.Rows)
            {
                ArrayList ListTmp = new ArrayList();
                //List<string> list = new List<string>();
                for (int y = 0; y < dt.Columns.Count; y++)
                {
                    ListTmp.Add(r[y].ToString());
                    
                }
                //object[] o_t = new object[] { ListTmp };
                target.Insert(ListTmp.ToArray());
            };

                
            //事务方式提交
            target.Flush();
            //操作记录数 有的话统计一下
            //
            ShowInfo("执行表" + tablename + "的bulk insert操作完成。");
            }
            catch(Exception ex)
            {
                ShowInfo("执行表"+ tablename + "的bulk insert操作出错：" + ex.Message);

            }

        }

        private void addBulkParameters(string tablename ,DataTable dt)
        {
            try
            {
                TARGET = null;//需要重新设置
                TARGET = new SQLiteBulkInsert(sqliteBlukCon, tablename);
                //根据表头字段 循环插入 类型全部是string 简单。否则特殊处理过于复杂
                foreach (DataColumn c in dt.Columns)
                {
                    TARGET.AddParameter(c.ColumnName, DbType.String);
                }
            }
            catch (Exception ex)
            {
                ShowInfo("始化表"+ tablename + "的数据库操作TARGET出错：" + ex.Message);
            }
            
                
        }





        #endregion  数据库相关

        
    }
}
