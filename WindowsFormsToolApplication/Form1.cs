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

using System.Xml;
using System.Xml.Linq;

using System.IO;
using NPOI.HSSF.UserModel;//2007office
using NPOI.XSSF.UserModel;//xlsx
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.HSSF.Util;
using NPOI;

//sqlitebulk 非官方的封装
using DB.SQLITE.SQLiteBulkInsert;
using System.Data.SQLite;
//ArrayList
using System.Collections;

//json
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

//
using System.Web;

//xml帮助类
using XMLHelper;
using System.Net;
using FileHelper;

//http request 帮助类
using HttpRequestHelper;

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
        //modeshape表字段配置
        List<string> modeshape_table = new List<string>();
        //modeshape表字段类型
        List<string> modeshape_type = new List<string>();

        //modeinfo表字段配置
        List<string> modeattr_cname_list = new List<string>();
        //站点名称配置列表
        List<string> stationinfo_cname_list = new List<string>();
        List<string> stationinfo_desc_list = new List<string>();
        //模式对比 dpl的模板路径
        string MODECHECK_SETFILEPATH;
        //模式对比 panel画面的模板路径
        string MODECHECK_PANELEXAMPLEPATH;
        //panel文件导入导出的路径
        string PANEL_INPUTFILEPATH;
        string PANEL_OUTPUTFILEPATH;
        string PANEL_BACKGROUND_COLOR;//panel背景颜色设置

        //dp的类名
        string MODESET_TYPENAME="";

        //模式对比的设备类型和设备标准值的设定，转化为json来处理
        JArray MODECHECK_JSONSET = new JArray();//
        //模式表 校验的时候，校验结果显示位置生成用关键字
        List<string> modetable_check_keywords_list = new List<string>();
        //模式配置文件 模式号的搜索关键字
        List<string> modetable_keywords_list = new List<string>();

        //模式对比画面名称解析 json
        JArray MODECHECK_PANELINFO = new JArray();
        //模式对比画面 矩形框的颜色信息 json
        JArray MODECHECK_RECTCOLOR = new JArray();//
        //http服务访问地址
        string httpurl_addGraphicPosition;
        string httpurl_getGraphicPositionByParams;

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


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                //清空数据库所有配置 20170829 暂时不清空 调试用
                //this.button1_Click(sender, e);
                //this.button2_Click(sender, e);
                //this.button3_Click(sender, e);
            }
            catch (Exception ex)
            {
                ShowInfo(ex.Message);
            }
        }



        //设备类别表导入数据库
        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                ////清空一下表数据，以免重复
                button2_Click(sender, e);
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

                    this.ShowInfo("设备类表导入结束");
                }
                else
                {
                    this.ShowInfo("文件路径不能为空");
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("读取设备类表配置文件异常：" + ex.Message);
            }

            
        }
        //模式表配置导入
        public void buttonImpMode_Click(object sender, EventArgs e)
        {
            try
            {
                //清空一下表数据，以免重复
                this.button1_Click(sender, e);
                //
                string filepath = openExcelDialog();
                if (filepath != "")
                {
                    DataTable dataTableShapeList = new DataTable();
                    //InitializeWorkbook(@filepath);//
                    dataTableModeList = ExcelToDataTable_modeinfo(@filepath,out dataTableShapeList);//模式的配置表相对不规范，另外写方法导入
                    //datatable 批量导入sqlite
                    if (dataTableModeList != null)
                    { 
                        insertDB_sqlitebulk_ex(dataTableModeList, "modeinfo");
                    }
                    //模式对比画面 宽高信息列表获取
                    if (dataTableShapeList != null)
                    {
                        addBulkParameters("modeshape",dataTableShapeList);
                        insertDB_sqlitebulk_ex(dataTableShapeList, "modeshape");
                    }
                    this.ShowInfo("模式配置表导入结束");

                }
                else
                {
                    this.ShowInfo("文件路径不能为空");
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
                //清空一下表数据，以免重复
                //button3_Click(sender, e);
                //
                string filepath = openExcelDialog();
                if (filepath != "")
                {

                    //InitializeWorkbook(@filepath);//
                    dataTableDeviceList = ExcelToDataTable(@filepath, true,2,2, 3);//
                    //datatable 批量导入sqlite
                    if (dataTableDeviceList != null)
                    {
                        insertDB_sqlitebulk_ex(dataTableDeviceList, "devicelist");
                    }
                    this.ShowInfo("设备清单列表导入结束");
                }
                else
                {
                    this.ShowInfo("文件路径不能为空");
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("读取设备清单文件异常：" + ex.Message);
            }
        }

        //提示窗口刷新 停止 切换
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
                if (tabControl1.SelectedIndex == 1)//只刷新选择导出tab的时候
                {
                    //添加站点列表
                    comboBox_stationlist.Text = "";
                    comboBox_stationlist.Items.Clear();
                    if (stationinfo_cname_list == null || stationinfo_cname_list.Count == 0)
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
                            comboBox_syslist.Items.Add(stationinfo_desc_list[i] + station_syslist[i]);
                        }
                        comboBox_syslist.SelectedIndex = 0;
                    }
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
        public void button1_Click(object sender, EventArgs e)
        {
            string SqlString;
            SqlString = "delete from modeinfo;delete from modeshape;";
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

        //生成模式配置dpl文件(20170726 配置写入画面 不写wincc数据库)
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
                        string sql_getdpnames = "select distinct t1.deviceid, t2.devindex ,t2.devid from modeinfo t1 left join devicelist t2 on t1.deviceid=t2.SBDM and t1.stationname=t2.stationid " +
                                                " where stationname='" + stationName + "' and pagename='" + pagename_list[i] + "' and (t2.sysid='DXTB' or t2.sysid='XXTB' or t2.sysid='SDTB' or t2.sysid is null);";
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
                    ShowInfo(stationName+"站的模式对比配置文件生成。");
                }


            }
            catch (Exception ex)
            {

                ShowInfo("生成模式配置文件dpl时出错：" + ex.Message);

            }

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
                    workbook = ExcelToWorkBook_modecheck(@filepath, out stationName);

                    if (workbook != null)
                    {
                        //另存excel校对文件
                        saveExcelDialog(workbook, stationName + "_MODECHECK_RESULT");
                        ShowInfo("模式表excel校准文件生成");
                    }
                    workbook = null;
                }
            }
            catch (Exception ex)
            {
                this.ShowInfo("模式表校对时出错：" + ex.Message);
            }
        }

        //生成模式配置的画面panel文件
        private void buttonCreateModePanel_Click(object sender, EventArgs e)
        {
            //获取站点和系统编号
            string stationName="";
            string modesysnum="";
            string stationShowName="";
            try {
                
                stationName = stationinfo_cname_list[comboBox_stationlist.SelectedIndex];
                stationShowName = stationinfo_desc_list[comboBox_stationlist.SelectedIndex];

                
                if(stationName==""|| comboBox_syslist.SelectedItem == null)
                {
                    MessageBox.Show("站点名称和系统名称不能未空", "信息提示框", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
                modesysnum = comboBox_syslist.SelectedItem.ToString();
                //从modeshape中抽取数据到json
                JArray shape_jsonobj = new JArray();
                shape_jsonobj = get_shape_jsonobj(stationName, modesysnum);
                if (shape_jsonobj == null)
                {
                    ShowInfo("获取模式对比画面的形状大小信息时没有返回结果，无法生成模式对比画面。");
                    return;
                }
                //从modeinfo中抽取数据到json
                JArray modeinfo_jsonobj = new JArray();
                modeinfo_jsonobj = get_modeinfo_jsonobj(stationName, modesysnum);
                if (modeinfo_jsonobj == null)
                {
                    ShowInfo("获取模式对比画面的数据信息时没有返回结果，无法生成模式对比画面。");
                    return;
                }


                //创建模式画面
                CreateModeCheckPanel(shape_jsonobj, modeinfo_jsonobj, stationShowName, stationName, modesysnum);

            }
            catch(Exception ex)
            {
                ShowInfo("生成"+ stationShowName + "的"+ modesysnum + "模式对比画面时出错：" + ex.Message);
            }
        }

        //获取模式对比画面的 数据信息
        public JArray get_modeinfo_jsonobj(string stationName, string modesysnum)
        {
            JArray ret_jsonobj = new JArray();
            string sql_str = "select distinct t1.deviceid,t2.devindex ,t2.devid,t1.width from modeinfo t1 "
                + "left join devicelist t2 on t1.deviceid=t2.SBDM and t1.stationname=t2.stationid where stationname='"
                + stationName + "' and pagename='" + modesysnum + "' and width<>'' and (t2.sysid='DXTB' or t2.sysid='XXTB' or t2.sysid='SDTB' or t2.sysid is null); ";
            try
            {
                using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
                {
                    DataSet ds_Result = _dbcon.GetData(sql_str);

                    if (ds_Result != null)
                    {
                        if (ds_Result != null && ds_Result.Tables.Count != 0 && ds_Result.Tables[0].Rows.Count != 0)
                        {
                            for (int i = 0; i < ds_Result.Tables[0].Rows.Count; i++)
                            {
                                JObject jsonset_t = new JObject(
                                                                new JProperty("deviceid",      ds_Result.Tables[0].Rows[i][0].ToString()),
                                                                new JProperty("devindex",      ds_Result.Tables[0].Rows[i][1].ToString()),
                                                                new JProperty("devid",         ds_Result.Tables[0].Rows[i][2].ToString()),
                                                                new JProperty("width",         ds_Result.Tables[0].Rows[i][3].ToString())
                                                                );
                                //
                                ret_jsonobj.Add(jsonset_t);

                            }

                            return ret_jsonobj;

                        }
                        else
                        {
                            ShowInfo("查询语句未查到数据：" + sql_str);
                            return null;
                        }
                    }
                    else
                    {
                        ShowInfo("查询语句执行出错：" + sql_str);
                        return null;
                    }
                }

            }
            catch (Exception ex)
            {
                ShowInfo("获取模式对比画面形状大小信息时出错：" + ex.Message);
                return null;
            }

        }

        //获取模式对比画面的 按钮和框的尺寸信息
        public JArray get_shape_jsonobj(string stationName, string modesysnum)
        {
            JArray ret_jsonobj = new JArray();
            string sql_str = "select modename,btnwidth,btnheight,kwidth,kheight from modeshape where stationname='"+ stationName + "' and pagename='"+ modesysnum + "'; ";
            try
            {
                using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
                {
                    DataSet ds_Result = _dbcon.GetData(sql_str);

                    if (ds_Result != null)
                    {
                        if (ds_Result != null && ds_Result.Tables.Count != 0 && ds_Result.Tables[0].Rows.Count != 0)
                        {
                            for (int i = 0; i < ds_Result.Tables[0].Rows.Count; i++)
                            {
                                JObject jsonset_t = new JObject(new JProperty("modename",  ds_Result.Tables[0].Rows[i][0].ToString()),
                                                                new JProperty("btnwidth",  ds_Result.Tables[0].Rows[i][1].ToString()),
                                                                new JProperty("btnheight", ds_Result.Tables[0].Rows[i][2].ToString()),
                                                                new JProperty("kwidth",    ds_Result.Tables[0].Rows[i][3].ToString()),
                                                                new JProperty("kheight",   ds_Result.Tables[0].Rows[i][4].ToString())
                                                                );
                                //
                                ret_jsonobj.Add(jsonset_t);
                                
                            }

                            return ret_jsonobj;

                        }
                        else
                        {
                            ShowInfo("查询语句未查到数据：" + sql_str);
                            return null;
                        }
                    }
                    else
                    {
                        ShowInfo("查询语句执行出错：" + sql_str);
                        return null;
                    }
                }
               
            }
            catch(Exception ex)
            {
                ShowInfo("获取模式对比画面形状大小信息时出错：" + ex.Message);
                return null;
            }
            
        }

        public void CreateModeCheckPanel(JArray shape_jsonobj, JArray modeinfo_jsonobj,string stationShowName,string stationName,string modesysnum )
        {
            string sPath = MODECHECK_PANELEXAMPLEPATH;//panel模板页面路径
            try
            {
                if (!File.Exists(sPath))
                {
                    this.ShowInfo("模式对比画面模板文件不存在，无法生成画面。");
                    return;
                }
                

                //读取模板文件
                XmlDocument xdPanelExample = new XmlDocument();  //实例化一个XmlDocument
                xdPanelExample.Load(sPath);
                //shapes节点
                XmlNode ShapesNode = xdPanelExample.SelectSingleNode("//panel/shapes");

                //背景图片选择 
                XmlNode xmlnode_panelbackground = xdPanelExample.SelectSingleNode("//panel/properties/prop[@name='Image']/prop");
                xmlnode_panelbackground.InnerText = "background/"+ stationName + "_"+modesysnum.Replace("_","-") + ".png";//路径保存(重要，中画线要转下划线)

                //当前模式文字 仅设置坐标即可
                XmlNodeList example_text_nowmode = xdPanelExample.SelectNodes("/panel/shapes/shape[@Name='PRIMITIVE_TEXT_NOWMODE']");
                //模式对比行文字 ToolTipText文字 例如：学林路,XLR,DXT,DXTB_MODA001
                //DXTB_MODA001 需要根据配置和逻辑拼接出来。
                XmlNodeList example_text_modeshow = xdPanelExample.SelectNodes("/panel/shapes/shape[@Name='PRIMITIVE_TEXT_MODESHOW']");

                //2个文字显示的坐标位置
                List<string> text_modenow_xy=new List<string>();
                List<string> text_modeshow_xy = new List<string>();

                #region //1.获取模式切换按钮模板
                XmlNodeList example_button = xdPanelExample.SelectNodes("/panel/shapes/shape[@Name='PUSH_BUTTON_EXAMPLE_MODESEL']");
                double btnwidth=0;//模式按钮宽高
                double btnheight=0;

                // page_modeinfo 例如 <prop name="en_US.utf8">学林路,XLR,DXT,DXTB_MODA001</prop>
                List<string> pagename_list = new List<string>();
                foreach(var it in comboBox_syslist.Items)
                {
                    pagename_list.Add(it.ToString());
                }
                string mode_sysname;
                string mode_sysside;
                //模式控制A,B端的输出点，（A,B最后会同步，这里自动生成都是A端的点,例如：DXTB_MODA001
                string mode_control_name = get_mode_controlname(pagename_list, modesysnum,out mode_sysname,out mode_sysside);
                //例子 学林路,XLR,DXT,DXTB_MODA001,小系统,II端 
                string page_modeinfo = stationShowName + "," + stationName + "," + modesysnum + "," + mode_control_name+"," + mode_sysname + ","+ mode_sysside;//页面大部分脚本显示公用信息
                string[] location_arr_btnfirst= { "0","0"} ;

                //Location Size Text 需要设置
                //button的名称 btn_mode_1 按照顺序排下去
                for (int i=0;i< shape_jsonobj.Count;i++)
                {
                    XmlNode example_button_t = example_button[0].CloneNode(true);
                    //SelectToken 方法使用
                    string modename = (string)shape_jsonobj[i].SelectToken("modename");
                    btnwidth = Math.Round((float)shape_jsonobj[i].SelectToken("btnwidth"), 0);  //size wincc oa 不允许有小数
                    btnheight = Math.Round((float)shape_jsonobj[i].SelectToken("btnheight")/(0.75*20), 0);//size wincc oa 不允许有小数
                    //Text
                    XmlNode xmlnode_t =example_button_t.SelectSingleNode("./properties/prop[@name='Text']");
                    foreach(XmlNode xn in xmlnode_t)
                    {
                        xn.InnerText = modename;
                    }

                    //Size 例如 <prop name="Size">52 39</prop>
                    xmlnode_t = example_button_t.SelectSingleNode("./properties/prop[@name='Size']");
                    xmlnode_t.InnerText = btnwidth.ToString() + " " + btnheight.ToString();

                    //Location 例如 <prop name="Location">185 376</prop>  (x y轴位置）
                    xmlnode_t = example_button_t.SelectSingleNode("./properties/prop[@name='Location']");
                    string[] location_arr = xmlnode_t.InnerText.Split(' ');
                    string x = location_arr[0];
                    string y = (Convert.ToDouble(location_arr[1])+ btnheight*i).ToString();
                    xmlnode_t.InnerText = x + " " + y;

                    //修改btn名称
                    XmlElement name_element = (XmlElement)example_button_t;
                    name_element.SetAttribute("Name", "btn_mode_" + i.ToString());


                    ShapesNode.AppendChild(example_button_t);
                    //获取第一个button的坐标 赋值给rec矩形框
                    if (i==0)
                    {
                        location_arr_btnfirst[0] = x;
                        location_arr_btnfirst[1] = y;
                    }

                    //调整模式状态和模式对比 2个文字显示的坐标
                    if(i== shape_jsonobj.Count - 1)
                    {
                        //当前模式状态的显示文字坐标 
                        XmlNode example_text_nownode_t = example_text_nowmode[0];
                        //Location 例如  <prop name="Location">210.1590909090908 748</prop>  (x y轴位置）
                        xmlnode_t = example_text_nownode_t.SelectSingleNode("./properties/prop[@name='Location']");
                        x = (Convert.ToDouble(location_arr[0])+ btnwidth*0.5).ToString();
                        y = (Convert.ToDouble(location_arr[1]) + btnheight * (shape_jsonobj.Count+0.5)).ToString();
                        xmlnode_t.InnerText = x + " " + y;
                        text_modenow_xy.Add(x);
                        text_modenow_xy.Add(y);


                        //模式对比文字坐标 
                        XmlNode example_text_modeshow_t = example_text_modeshow[0];
                        xmlnode_t = example_text_modeshow_t.SelectSingleNode("./properties/prop[@name='Location']");
                        x = (Convert.ToDouble(location_arr[0]) + btnwidth * 0.5).ToString();
                        y = (Convert.ToDouble(location_arr[1]) + btnheight * (shape_jsonobj.Count+1.5)).ToString();
                        xmlnode_t.InnerText = x + " " + y;
                        text_modeshow_xy.Add(x);
                        text_modeshow_xy.Add(y);
                        //ToolTipText
                        xmlnode_t = example_text_modeshow_t.SelectSingleNode("./properties/prop[@name='ToolTipText']");
                        foreach (XmlNode xn in xmlnode_t)
                        {
                            xn.InnerText = page_modeinfo;
                        }

                    }
                }
                #endregion

          
                #region//2.获取矩形框模板
                XmlNodeList example_rect = xdPanelExample.SelectNodes("/panel/shapes/shape[@Name='RECTANGLE_EXAMPLE_MODESELECTED']");

                for (int i = 0; i < shape_jsonobj.Count; i++)
                {
                    XmlNode example_rec_t = example_rect[0].CloneNode(true);
                    //获取模式号 判断矩形框对应的颜色
                    string modename = (string)shape_jsonobj[i].SelectToken("modename");//SelectToken 方法使用
                    string sysfirstnum = modename.Substring(0,1);
                    //颜色获取
                    IEnumerable<JToken> stv = MODECHECK_RECTCOLOR.SelectTokens("$.[?(@.sysfirstnum == '"+ sysfirstnum + "')].cname");
                    if (stv.Count() == 0) continue;
                    
                    //ForeColor 例如 <prop name="ForeColor">{0,162,0}</prop>默认绿色
                    XmlNode xmlnode_t = example_rec_t.SelectSingleNode("./properties/prop[@name='ForeColor']");
                    xmlnode_t.InnerText = stv.ToList()[0].ToString();

                    double kwidth = Math.Round((float)shape_jsonobj[i].SelectToken("kwidth"), 0);  //size wincc oa 不允许有小数
                    double kheight = Math.Round((float)shape_jsonobj[i].SelectToken("kheight") / (0.75 * 20), 0);//size wincc oa 不允许有小数

                    //Size 例如 <prop name="Size">1317.060606060606 40</prop>
                    xmlnode_t = example_rec_t.SelectSingleNode("./properties/prop[@name='Size']");
                    xmlnode_t.InnerText = kwidth.ToString() + " " + kheight.ToString();

                    //ToolTipText
                    xmlnode_t = example_rec_t.SelectSingleNode("./properties/prop[@name='ToolTipText']");
                    foreach (XmlNode xn in xmlnode_t)
                    {
                        xn.InnerText = modename;
                    }

                    //Location 例如 <prop name="Location">185 371</prop>  (x y轴位置）
                    //位置和第一个button相同
                    xmlnode_t = example_rec_t.SelectSingleNode("./properties/prop[@name='Location']");
                    string x = (Convert.ToDouble(location_arr_btnfirst[0])-1).ToString();//矩形框 需要缩进一个像素
                    string y = (Convert.ToDouble(location_arr_btnfirst[1]) + kheight * i).ToString();
                    xmlnode_t.InnerText = x + " " + y;

                    //修改rec名称
                    XmlElement name_element = (XmlElement)example_rec_t;
                    name_element.SetAttribute("Name", "rec_mode_" + i.ToString());

                    ShapesNode.AppendChild(example_rec_t);
                }
                #endregion


                #region//3.当前设备运行状态
                double ref_nowstatus_x = btnwidth*0.5+ Convert.ToDouble(text_modenow_xy[0])+22;//修正22 状态方块自己的宽度
                double ref_nowstatus_y = btnheight*0.5 + Convert.ToDouble(text_modenow_xy[1]);
                XmlNodeList example_ref_nowstatus = xdPanelExample.SelectNodes("/panel/shapes/reference[@Name='PANEL_REF_EXAMPLE_MODESTATUS']");
                for (int i = 0; i < modeinfo_jsonobj.Count; i++)
                {
                    XmlNode example_ref_nowstatus_t = example_ref_nowstatus[0].CloneNode(true);
                    //获取配置
                    string devindex = (string)modeinfo_jsonobj[i].SelectToken("devindex");
                    string devid = (string)modeinfo_jsonobj[i].SelectToken("devid");
                    string width = (string)modeinfo_jsonobj[i].SelectToken("width");
                    string deviceid = (string)modeinfo_jsonobj[i].SelectToken("deviceid"); 

                    //信息提示
                    if (devindex == "")
                    {
                        ShowInfo("模式对比画面，当前状态。第"+(i+1).ToString()+"列："+ deviceid + "没有匹配到对应的设备代码。");
                    }

                    //dollarParameters 配置
                    XmlNode xmlnode_t = example_ref_nowstatus_t.SelectSingleNode("./properties/prop[@name='dollarParameters']/prop/prop[@name='Value']");
                    xmlnode_t.InnerText = devindex;

                    //Location 坐标
                    xmlnode_t = example_ref_nowstatus_t.SelectSingleNode("./properties/prop[@name='Location']");
                    ref_nowstatus_x = ref_nowstatus_x + (Convert.ToDouble(width) * 0.5);  //x轴移动半格
                    string x = ref_nowstatus_x.ToString();
                    ref_nowstatus_x = ref_nowstatus_x + (Convert.ToDouble(width) * 0.5);  //x轴移动半格
                    string y = ref_nowstatus_y.ToString(); 
                    xmlnode_t.InnerText = x + " " + y;

                    //修改名称
                    XmlElement name_element = (XmlElement)example_ref_nowstatus_t;
                    name_element.SetAttribute("Name", "ref_modestatus_" + i.ToString());

                    //panel路径获取
                    IEnumerable<JToken> stv = MODECHECK_JSONSET.SelectTokens("$.[?(@.cname == '" + devid + "')].filename");
                    if (stv.Count() == 0) continue;

                    //FileName <prop name="FileName">objects/emcs/MOD_TVF.XML</prop>
                    xmlnode_t = example_ref_nowstatus_t.SelectSingleNode("./properties/prop[@name='FileName']");
                    xmlnode_t.InnerText = stv.ToList()[0].ToString();

                    
                    ShapesNode.AppendChild(example_ref_nowstatus_t);
                }
                #endregion


                #region//4.模式对比结果圆圈
                double ref_checkstatus_x = btnwidth * 0.5 + Convert.ToDouble(text_modeshow_xy[0])+16+8;//修正16 圆圈的直径+半径8 可能和属性Geometry有关
                double ref_checkstatus_y = btnheight* 0.5 + Convert.ToDouble(text_modeshow_xy[1]);//
                XmlNodeList example_ref_checkresult = xdPanelExample.SelectNodes("/panel/shapes/reference[@Name='PANEL_REF_EXAMPLE_MODECHECK']");
                for (int i = 0; i < modeinfo_jsonobj.Count; i++)
                {
                    XmlNode example_ref_checkresult_t = example_ref_checkresult[0].CloneNode(true);
                    //获取配置
                    string devindex = (string)modeinfo_jsonobj[i].SelectToken("devindex");
                    string devid = (string)modeinfo_jsonobj[i].SelectToken("devid");
                    string width = (string)modeinfo_jsonobj[i].SelectToken("width");

                    //Location 坐标
                    XmlNode xmlnode_t = example_ref_checkresult_t.SelectSingleNode("./properties/prop[@name='Location']");
                    ref_checkstatus_x = ref_checkstatus_x + (Convert.ToDouble(width) * 0.5);  //x轴移动半格
                    string x = ref_checkstatus_x.ToString();
                    ref_checkstatus_x = ref_checkstatus_x + (Convert.ToDouble(width) * 0.5);  //x轴移动半格
                    string y = ref_checkstatus_y.ToString();
                    xmlnode_t.InnerText = x + " " + y;

                    //修改名称 modecheck_signalshow_1 必须从1开始 否则wincc oa脚本无法关联上去。
                    XmlElement name_element = (XmlElement)example_ref_checkresult_t;
                    name_element.SetAttribute("Name", "modecheck_signalshow_" + (i+1).ToString());

                    ShapesNode.AppendChild(example_ref_checkresult_t);
                }
                #endregion


                #region//5.模式配置导入画面脚本 （不利用wincc数据库进行保存）
                XmlNode example_script = xdPanelExample.SelectSingleNode("//panel/events/script[@name='Initialize']");
                string script_code_t = example_script.InnerText;
                string script_code=get_panel_set(script_code_t);
                if (script_code == "")
                {
                    ShowInfo("生成画面时脚本文件解析出错，请检查模板文件");
                    return;
                }
                 
                //example_script.InnerText = script_code;
                example_script.InnerXml = "<![CDATA[" + script_code + "\r\n]]>";
                #endregion

                #region//6.模式画面切换按钮的自动生成
                XmlNodeList example_btn_modechange = xdPanelExample.SelectNodes("/panel/shapes/shape[@Name='PUSH_BUTTON_MODECHANGE']");
                XmlNodeList example_rec_modechange = xdPanelExample.SelectNodes("/panel/shapes/shape[@Name='RECTANGLE_MODECHANGE']");
                XmlNodeList example_text_modechange = xdPanelExample.SelectNodes("/panel/shapes/shape[@Name='PRIMITIVE_TEXT_MODECHANGE']");//图片按钮上的模式名称的text
                //按钮的大小，从模板文件获取，不固定死
                //按钮部分的自动排列生成
                //1.修改btn的Text；2.修改btn的tooltip作为脚本的文件路径；3.位置x,y的修改。
                int btn_rowsnum_modechange =2;//模式按钮的行数，一般应该都是2行。
                int btn_changerow_num =Convert.ToInt32(Math.Ceiling((double)(pagename_list.Count) / (double)btn_rowsnum_modechange));//换行的序号，需要向上取整
                //获取初始的xy值  btn 注意说明：图片的y值需要在按钮y值上减1
                XmlNode xmlnode_btn_t = example_btn_modechange[0].SelectSingleNode("./properties/prop[@name='Location']");
                string[] location_arr_t = xmlnode_btn_t.InnerText.Split(' ');
                int btn_x_t= Convert.ToInt32(location_arr_t[0]);
                int btn_y_t= Convert.ToInt32(location_arr_t[1]);
                //获取按钮大小size
                xmlnode_btn_t = example_btn_modechange[0].SelectSingleNode("./properties/prop[@name='Size']");
                string[] size_arr_t = xmlnode_btn_t.InnerText.Split(' ');
                int btn_w_t = Convert.ToInt32(size_arr_t[0]);
                int btn_h_t = Convert.ToInt32(size_arr_t[1]);

                for (int i = 0; i < pagename_list.Count; i++)
                {
                    //按钮的文字需要特殊处理 DXT-大系统 / SDT - 隧道系统
                    string btn_name;
                    if (pagename_list[i] == "DXT")
                    {
                        btn_name = "大系统";
                    }
                    else if(pagename_list[i] == "SDT")
                    {
                        btn_name = "隧道系统";
                    }
                    else
                    {
                        btn_name = pagename_list[i];
                    };
                    //x y轴的参数
                    int x_t;
                    int y_t;
                    if (i < btn_changerow_num)
                    {
                        x_t = btn_x_t+i* btn_w_t;
                        y_t = btn_y_t; 
                    }
                    else
                    {
                        x_t = btn_x_t +(i-btn_changerow_num) * btn_w_t;
                        y_t = btn_y_t+ btn_h_t;
                    }

                    if (pagename_list[i]== modesysnum)//当前页 用图片rectangle 和 text 作为按钮，
                    {
                        //当前按钮只有一个 不用克隆，直接修改模板元素属性
                        //图片按钮没有脚本，只需要改变位置 显示的文字
                        //图片按钮部分
                        XmlNode example_rec_modechange_t = example_rec_modechange[0];
                        //Location 图片的宽高，都需要-1
                        XmlNode xmlnode_t = example_rec_modechange_t.SelectSingleNode("./properties/prop[@name='Location']");
                        xmlnode_t.InnerText = (x_t-1).ToString() + " " + (y_t-1).ToString();
                        //图片文字部分
                        XmlNode example_text_modechange_t = example_text_modechange[0];
                        //文字偏移量 预估写死
                        xmlnode_t = example_text_modechange_t.SelectSingleNode("./properties/prop[@name='Location']");
                        xmlnode_t.InnerText = (x_t +44 ).ToString() + " " + (y_t + 4).ToString();
                        //图片按钮的文字显示
                        xmlnode_t = example_text_modechange_t.SelectSingleNode("./properties/prop[@name='Text']");
                        foreach (XmlNode xn in xmlnode_t)
                        {
                            xn.InnerText = btn_name;
                        }
                    }
                    else//其他直接用按钮
                    {
                        //其他页按钮多个，需要动态生成，模板用克隆方式
                        XmlNode example_btn_modechange_t = example_btn_modechange[0].CloneNode(true);
                        //Text 
                        XmlNode xmlnode_t = example_btn_modechange_t.SelectSingleNode("./properties/prop[@name='Text']");
                        foreach (XmlNode xn in xmlnode_t)
                        {
                            xn.InnerText = btn_name;
                        }
                        //ToolTipText 
                        xmlnode_t = example_btn_modechange_t.SelectSingleNode("./properties/prop[@name='ToolTipText']");
                        foreach (XmlNode xn in xmlnode_t)
                        {
                            xn.InnerText = pagename_list[i];
                        }
                        //Location
                        xmlnode_t = example_btn_modechange_t.SelectSingleNode("./properties/prop[@name='Location']");
                        xmlnode_t.InnerText = x_t.ToString() + " " + y_t.ToString();

                        //修改btn名称 添加到xml节点里
                        XmlElement name_element = (XmlElement)example_btn_modechange_t;
                        name_element.SetAttribute("Name", "btn_modechange_" + i.ToString());

                        ShapesNode.AppendChild(example_btn_modechange_t);
                    }
                }  
                #endregion


                //删除不需要的模板节点(动态生成的画面元素 基本都要把模板删除）
                ShapesNode.RemoveChild(example_button[0]);//模式按钮的模板
                ShapesNode.RemoveChild(example_rect[0]);//模式切换矩形框的模板
                ShapesNode.RemoveChild(example_ref_nowstatus[0]);//当前设备状态的模板
                ShapesNode.RemoveChild(example_ref_checkresult[0]);//模式对比结果的元素模板
                ShapesNode.RemoveChild(example_btn_modechange[0]);//模式页面切换的按钮模板
                
                //任何shape对象的TabOrder唯一 serialId唯一
                int TabOrder = 1;
                //shpaes下面所有的ref和shape进行TabOrder 编号
                XmlNodeList shapes_list = ShapesNode.SelectNodes("./*");
                foreach(XmlNode xn in shapes_list)
                {
                    XmlNode xmlnode_t = xn.SelectSingleNode("./properties/prop[@name='TabOrder']");
                    xmlnode_t.InnerText = TabOrder.ToString();
                    TabOrder = TabOrder + 1;
                }

                //所有引用panel的referenceId唯一
                int referenceId = 1;
                //shpaes下面所有的ref和shape进行TabOrder 编号
                XmlNodeList ref_list = ShapesNode.SelectNodes("./reference");
                foreach (XmlNode xn in ref_list)
                {
                    XmlElement xn_element = (XmlElement)xn;
                    xn_element.SetAttribute("referenceId", referenceId.ToString());
                    referenceId = referenceId + 1;
                }

                
                //保存xml文件 XLR_EMCS_MOD_DXT.xml
                string filename = stationName+"_EMCS_MOD_" + modesysnum;
                saveXMLDialog(xdPanelExample, filename);

            }
            catch (Exception ex)
            {
                this.ShowInfo("生成模式对比画面时出错：" + ex.Message);

            }
        }

        public string get_panel_set(string code_string)
        {
            try
            {
                
                string stationName;
                stationName = stationinfo_cname_list[comboBox_stationlist.SelectedIndex];//站点缩写
                string pagename = comboBox_syslist.SelectedItem.ToString();//panel页面名称

                //生成配置清单
                //1
                string modelist;//模式列表
                string sql_getmodename = "select distinct modename from modeinfo where stationname='" + stationName + "' and pagename='" + pagename + "';";
                List<string> modename_list = new List<string>();
                modename_list = getDataList(sql_getmodename);
                modelist= (string.Join("&quot;, &quot;", modename_list.ToArray()));
                
                //2
                string dpnamelist;//dpname list
                string controltype;//设备类型 list
                //
                string sql_getdpnames = "select distinct t1.deviceid, t2.devindex ,t2.devid from modeinfo t1 left join devicelist t2 on t1.deviceid=t2.SBDM and t1.stationname=t2.stationid " +
                                            " where stationname='" + stationName + "' and pagename='" + pagename + "' and (t2.sysid='DXTB' or t2.sysid='XXTB' or t2.sysid='SDTB' or t2.sysid is null);";
                List<string> dpname_list = new List<string>();//dpnames列表
                List<string> devid_list = new List<string>();//设备的类型列表 和配置文件比较 转换成类型编号（转换不出来的置0）
                getDataList_sp1(sql_getdpnames, ref dpname_list, ref devid_list);
                //
                dpnamelist  = (string.Join("&quot;, &quot;", dpname_list.ToArray()));
                //设备类型
                List<string> controltype_list = new List<string>();//设备类型 列表
                controltype_list = controltype_checkout(devid_list);//根据devid （TVF TDZ UFO FF1 FF2...）以及配置文件 转换出类型的编号
                controltype = (string.Join("", controltype_list.ToArray()));

                //3
                string modevaluelist;//模式标准值清单
                List<string> modevalue_list = new List<string>();//每种模式下面的标准值 列表
                modevalue_list = modevalue_checkout(modename_list, stationName, pagename);//根据站点、系统号、模式编号3个关键字查询 得到设备运行标准状态列表
                modevaluelist = (string.Join("&quot;, &quot;", modevalue_list.ToArray()));                                                                                  //

                //替换对应的脚本内容
                //dyn_string modelist = makeDynString("@modelist@");
                //dyn_string dpnamelist = makeDynString("@dpnamelist@");
                //dyn_string modevaluelist = makeDynString("@modevaluelist@");
                //string controltype = "@controltype@";
                string code_ret= code_string.Replace("@modelist@", modelist);
                code_ret = code_ret.Replace("@dpnamelist@", dpnamelist);
                code_ret = code_ret.Replace("@modevaluelist@", modevaluelist);
                code_ret = code_ret.Replace("@controltype@", controltype);
                return code_ret;
            }
            catch (Exception ex)
            {
                
                ShowInfo("生成模式配置文件dpl时出错：" + ex.Message);
                return "";
            }

        }


        //根据模式分页信息 拼接 列入：XXTB_MODA001   DXTB_MODA001等等
        public string get_mode_controlname(List<string> pagename_list, string modesysnum,out string mode_sysname_out,out string mode_sysside_out)
        {
            mode_sysname_out = "";
            mode_sysside_out = " ";
            try
            {
                
                string modesysnum_t=Regex.Replace(modesysnum, @"\d", "");//系统名称数字过滤
                string ret_dpname="";
                int i_dxt = 1, i_sdt = 1, i_xxta=1,i_xxtb = 1;
                int end_num = 1;
                IEnumerable<JToken> stv = null;
                IEnumerable<JToken> modesysname = null;
                IEnumerable<JToken> modesysside = null;
                for (int i = 0; i < pagename_list.Count; i++)
                {
                    stv = MODECHECK_PANELINFO.SelectTokens("$.[?(@.cname == '" + modesysnum_t + "')].sysid");
                    modesysname = MODECHECK_PANELINFO.SelectTokens("$.[?(@.cname == '" + modesysnum_t + "')].modesysname");
                    modesysside = MODECHECK_PANELINFO.SelectTokens("$.[?(@.cname == '" + modesysnum_t + "')].modesysside");
                    //这部分后续优化掉
                    if (pagename_list[i].IndexOf("_XII") >= 0)
                    {        
                        end_num = i_xxtb;
                        i_xxtb++;
                    }
                    else if (pagename_list[i].IndexOf("_XI") >= 0)
                    {
                        
                        end_num = i_xxta;
                        i_xxta++;
                    }
                    else if (pagename_list[i].IndexOf("SDT") >= 0)
                    {
                        end_num = i_sdt;
                        i_sdt++;
                    }
                    else if (pagename_list[i].IndexOf("DXT") >= 0)
                    {
                        end_num = i_dxt;
                        i_dxt++;
                    }
                    else
                    {
                        //do nothing
                    }


                if (modesysnum == pagename_list[i])
                    {
                        if (stv.Count() > 0)
                        {
                            ret_dpname = stv.ToList()[0].ToString();
                        }
                        if (modesysname.Count() > 0)//中文系统缩写获取
                        {
                            mode_sysname_out = modesysname.ToList()[0].ToString();
                        }
                        if (modesysside.Count() > 0)//系统AB端的获取 大系统和隧道系统没有，直接空格
                        {
                            mode_sysside_out = modesysside.ToList()[0].ToString();
                        }
                        
                        break;
                    }
                    
                    
                }

                //
                return ret_dpname + end_num.ToString("000"); ;
            }
            catch(Exception ex)
            {
                ShowInfo("生成模式对比画面时，页面对应模式点名获取错误：" + ex.Message);
                return "";
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
                                SYSID = filename.Split('_')[1];//ZKR_EMCS_DeviceList.xlsx
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
                                                        column.DataType = typeof(string);
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
                                                //设备清单的特殊处理 
                                                //1.第二列 序号里面的数据不用
                                                //2.beizhu3这一列的数据，用系统名称替换
                                                if (typeOfExcel == 2  ) {
                                                    if(j - extNum == 1)
                                                    {
                                                        i_ts = i_ts + 1;
                                                        continue;
                                                    }
                                                    if (j - extNum == 18)
                                                    {
                                                        dataRow[j - i_ts] = SYSID;
                                                        continue;
                                                    }  
                                                }

                                                //有些cell有公式
                                                cell = row.GetCell(j - extNum);//cell数据获取
                                                //有些cell前后有多余的空格
                                                dataRow[j- i_ts] =  npoi_celldeal(cell).ToString().Trim().ToUpper();
                                                
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
        public DataTable ExcelToDataTable_modeinfo(string filePath,out DataTable dtshape)//
        {

            dtshape = new DataTable();

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
                            dataTable.Columns.Add(s,typeof(string));
                        }
                        addBulkParameters("modeinfo", dataTable);
                        //形状table初始化column 和 类型
                        for (int i= 0; i < modeshape_table.Count; i++)
                        {
                            dtshape.Columns.Add(modeshape_table[i], GetTypeByString(modeshape_type[i]));
                        }
                        

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
                                            if (r_flag)
                                            {
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
                                        //模式切换按钮形状
                                        List<float> btnwidth_list = new List<float>();
                                        List<float> btnheight_list = new List<float>();
                                        //模式切换框的宽度
                                        float kwidth = 0;

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
                                            
                                            if (str_cvalue!="")
                                            {
                                                if (int.TryParse(str_cvalue, out out_i))
                                                {
                                                    modeid_list.Add(str_cvalue.Trim());
                                                    //模式单元格的宽高获取
                                                    float w_t = 0;
                                                    float h_t = 0;
                                                    ExcelExtension.IsMergeCellShape(sheet.GetRow(i).GetCell(MOSHI_CNUM), out h_t, out w_t);
                                                    btnwidth_list.Add(w_t);
                                                    btnheight_list.Add(h_t);
                                                }
                                                else
                                                {
                                                    ShowInfo("sheet"+ PAGENAME +"的第"+i.ToString()+"行的模式号不是纯数字");
                                                }
                                                

                                            }
                                            else if (IsMerge)
                                            {
                                                
                                                continue;
                                            }
                                            else
                                            {
                                                //无内容，模式读完
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

                                        float w_t1 = 0;//模式对比值的单元格的宽高获取
                                        float h_t1 = 0;
                                        for (int i = 0; i < modeid_list.Count; i++)
                                        {
                                            int colAdd = 0;//合并列所带来的列偏移 (20170724 列的合并比较复杂，暂时也没有这种情况出现，不做处理。)
                                            int rowSpan=1;
                                            int columnSpan=1;

                                            for (int j = 0; j < devid_list.Count; j++)
                                            {
                                                //模式对比值的datatable生成
                                                dataRow = dataTable.NewRow();//dataTable中创建新行
                                                dataRow[0] = STATIONID;
                                                dataRow[1] = PAGENAME;
                                                dataRow[2] = modeid_list[i];
                                                dataRow[3] = devid_list[j];
                                                
                                                //格式处理
                                                //模式对比 有可能比对值会有小写 处理掉
                                                dataRow[4] = npoi_celldeal(sheet.GetRow(MOSHI_RNUM + 1 + i+ rowAdd).GetCell(MOSHI_CNUM + j + 1)).ToString().ToUpper();
                                                //模式标准值的单元格宽度
                                                if (i == 0)
                                                {
                                                    //字符获取尝试 32.04左右的字符-像素比例处理
                                                    dataRow[5] = (sheet.GetColumnWidth(MOSHI_CNUM + j + 1 )/32).ToString();
                                                    //像素取值有一些问题，会有拉伸情况 20170724
                                                    //dataRow[5] = sheet.GetColumnWidthInPixels(MOSHI_CNUM + j + 1 ).ToString(); 
                                                }
                                                else
                                                {
                                                    dataRow[5] = "";
                                                }

                                                //处理一行添加一行数据
                                                dataTable.Rows.Add(dataRow);//dataTable数据添加一行
                                                
                                                //获取数值单元格的宽度 仅第一次轮询
                                                if (i == 0)
                                                {
                                                    NPOI.ExcelExtension.IsMergeCellShape(sheet.GetRow(MOSHI_RNUM + 1 + i).GetCell(MOSHI_CNUM + j + 1), out h_t1, out w_t1);
                                                    kwidth= kwidth+w_t1;
                                                }
                                            }
                                            //处理行的偏移量
                                            rowSpan = 1;
                                            columnSpan = 1;
                                            if (NPOI.ExcelExtension.IsMergeCell(sheet.GetRow(MOSHI_RNUM + 1 + i+ rowAdd).GetCell(MOSHI_CNUM), out rowSpan, out columnSpan))
                                            {
                                                rowAdd = rowAdd + rowSpan - 1;
                                            }
                                            //
                                            dataRow = dtshape.NewRow();
                                            dataRow[0] = STATIONID;
                                            dataRow[1] = PAGENAME;
                                            dataRow[2] = modeid_list[i];
                                            dataRow[3] = btnwidth_list[i];
                                            dataRow[4] = btnheight_list[i];
                                            dataRow[5] = kwidth+ btnwidth_list[i];
                                            dataRow[6] = btnheight_list[i];
                                            dtshape.Rows.Add(dataRow);
                                        }
                                    }     
                                }
                            }
                            //
                            

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
                                        string sql_getdata = "select devindex, devid, sbdm from devicelist t2 where stationid = '" + STATIONID + "' and sbdm in(" + str_tmp + ") and (t2.sysid='DXTB' or t2.sysid='XXTB' or t2.sysid='SDTB' or t2.sysid is null); ";

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
                XElement xeModeShapeRoot = xeRoot.Descendants("modeshape_table").First();
                if (xeModeShapeRoot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModeShapeRoot.Descendants("cname"))
                    {
                        modeshape_table.Add(xePort.Attribute("cname").Value.ToString());
                        modeshape_type.Add(xePort.Attribute("type").Value.ToString());
                    }
                }
                if (modeshape_table.Count == 0)
                {
                    this.ShowInfo("模式形状大小配置表字段相关配置丢失，请重新修改配置文件");
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
                    this.ShowInfo("模式配置模板文件路径为空，无法生成模式配置dpl文件。");

                //获取模式检查表的配置文件dpl的生成模板
                XElement xeModecheckPanelExamplePath = xeRoot.Descendants("modecheck_examplepanel").First();
                MODECHECK_PANELEXAMPLEPATH = System.Environment.CurrentDirectory + xeModecheckPanelExamplePath.Attribute("filepath").Value.ToString();
                if (MODECHECK_PANELEXAMPLEPATH == "")
                    this.ShowInfo("模式对比画面模板文件路径为空，无法生成模式对比画面。");

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
                                new JProperty("filename", xePort.Attribute("filename").Value.ToString()),
                                new JProperty("rows", jsonset_arr)
                                );
                        MODECHECK_JSONSET.Add(jsonset_t);
                    }
                }

                //获取模式对比导出画面的一些相关信息
                XElement xeModecheckPanelInfoRoot = xeRoot.Descendants("modetable_sysid").First();
                if (xeModecheckPanelInfoRoot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModecheckPanelInfoRoot.Descendants("cname"))
                    {
                        JObject jsonset_t = new JObject(
                                new JProperty("cname", xePort.Attribute("cname").Value.ToString()),
                                new JProperty("sysid", xePort.Attribute("sysid").Value.ToString()),
                                new JProperty("desc", xePort.Attribute("desc").Value.ToString()),
                                new JProperty("modesysname", xePort.Attribute("modesysname").Value.ToString()),
                                new JProperty("modesysside", xePort.Attribute("modesysside").Value.ToString())
                                );
                        MODECHECK_PANELINFO.Add(jsonset_t);
                    }
                }

                //获取模式对比导出画面的一些相关信息
                XElement xeModecheckRectangleColorRoot = xeRoot.Descendants("modetable_color").First();
                if (xeModecheckRectangleColorRoot.Elements().Count() > 0)
                {
                    foreach (XElement xePort in xeModecheckRectangleColorRoot.Descendants("cname"))
                    {
                        JObject jsonset_t = new JObject(
                                new JProperty("cname", xePort.Attribute("cname").Value.ToString()),
                                new JProperty("sysfirstnum", xePort.Attribute("sysfirstnum").Value.ToString()),
                                new JProperty("desc", xePort.Attribute("desc").Value.ToString())
                                );
                        MODECHECK_RECTCOLOR.Add(jsonset_t);
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

                //获取模式检查表的配置文件dpl的生成模板
                XElement xePanelRoot = xeRoot.Descendants("panelfile").First();
                PANEL_INPUTFILEPATH = System.Environment.CurrentDirectory + xePanelRoot.Attribute("input_filepath").Value.ToString();
                PANEL_OUTPUTFILEPATH = System.Environment.CurrentDirectory + xePanelRoot.Attribute("output_filepath").Value.ToString();
                PANEL_BACKGROUND_COLOR = xePanelRoot.Attribute("background_color").Value.ToString();
                //访问http服务地址 添加画面信息；获取画面信息
                httpurl_addGraphicPosition= xePanelRoot.Attribute("httpurl_addGraphicPosition").Value.ToString();
                httpurl_getGraphicPositionByParams = xePanelRoot.Attribute("httpurl_getGraphicPositionByParams").Value.ToString();


                if (PANEL_INPUTFILEPATH == "" || PANEL_OUTPUTFILEPATH == "")
                    this.ShowInfo("PANEL文件导入导出路径配置丢失，无法生成画面文件。");


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
            //int ret = -1;//执行出错返回-1

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
                                                " where stationname='" + stationName + "' and pagename='" + pagename + "' and modename='"+ mn + "' and (t2.sysid='DXTB' or t2.sysid='XXTB' or t2.sysid='SDTB' or t2.sysid is null)";
                        
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
                    ShowInfo(filename + "文件保存完毕");
                }
                
            }
            catch (Exception ex)
            {
                ShowInfo(filename+"excel文件保存出错：" +ex.Message);
            }
        }

        //保存xml窗口
        //filename 默认文件名
        public void saveXMLDialog(XmlDocument xmldoc, string filename)
        {
            try
            { 
                SaveFileDialog dialog = new SaveFileDialog();
                // 默认扩展名  
                dialog.DefaultExt = ".xml";
                // 默认文件名  
                dialog.FileName = filename;

                dialog.Title = "请选择文件夹";
                dialog.Filter = "所有文件(*.xml)|*.xml";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    xmldoc.Save(dialog.FileName);
                }
                ShowInfo(filename+"xml文件生成完毕");

            }
            catch (Exception ex)
            {
                ShowInfo(filename+"保存xml文件出错：" + ex.Message);
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

        //类型转换 需要持续扩展
        public static Type GetTypeByString(string type)
        {
            switch (type.ToLower())
            {
                case "bool":
                    return Type.GetType("System.Boolean", true, true);
                case "byte":
                    return Type.GetType("System.Byte", true, true);
                case "sbyte":
                    return Type.GetType("System.SByte", true, true);
                case "char":
                    return Type.GetType("System.Char", true, true);
                case "decimal":
                    return Type.GetType("System.Decimal", true, true);
                case "double":
                    return Type.GetType("System.Double", true, true);
                case "float":
                    return Type.GetType("System.Single", true, true);
                case "int":
                    return Type.GetType("System.Int32", true, true);
                case "uint":
                    return Type.GetType("System.UInt32", true, true);
                case "long":
                    return Type.GetType("System.Int64", true, true);
                case "ulong":
                    return Type.GetType("System.UInt64", true, true);
                case "object":
                    return Type.GetType("System.Object", true, true);
                case "short":
                    return Type.GetType("System.Int16", true, true);
                case "ushort":
                    return Type.GetType("System.UInt16", true, true);
                case "string":
                    return Type.GetType("System.String", true, true);
                case "date":
                case "datetime":
                    return Type.GetType("System.DateTime", true, true);
                case "guid":
                    return Type.GetType("System.Guid", true, true);
                default:
                    return Type.GetType(type, true, true);
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
            try
            {
                if (dt.Rows.Count == 0)
                {
                    ShowInfo("没有解析出数据，无法导入数据库。");
                    return;
                }
                
                //用自定义sqlite类操作
                SQLiteBulkInsert target;
                //Close connect to verify records were not inserted
                sqliteBlukCon.Close();
                //target = new SQLiteBulkInsert(sqliteBlukCon, tablename);
                target = TARGET;
                sqliteBlukCon.Open();
                target.CommitMax = 10000;//1w条数据 限制 默认值1w



                //Insert less records than commitmax
                foreach (DataRow r in dt.Rows)
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
            catch (Exception ex)
            {
                ShowInfo("执行表" + tablename + "的bulk insert操作出错：" + ex.Message);

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
                    //TARGET.AddParameter(c.ColumnName, DbType.String);//全部是string的写法
                    TARGET.InitialTypetoDbType();//初始化类型字典
                    TARGET.AddParameter(c.ColumnName, TARGET.typeMap[c.DataType]);
                    
                }
            }
            catch (Exception ex)
            {
                ShowInfo("始化表"+ tablename + "的数据库操作TARGET出错：" + ex.Message);
            }
            
                
        }



        #endregion  数据库相关

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //生成panel画面
        private void button_createpanel_Click(object sender, EventArgs e)
        {
            try
            {

                string stationName = stationinfo_cname_list[comboBox_stationlist.SelectedIndex];
                string stationShowName = stationinfo_desc_list[comboBox_stationlist.SelectedIndex];

                if (stationName == "" )
                {
                    MessageBox.Show("请选择站点进行画面生成", "信息提示框", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }

                //从panelfile导入xml文件 循环进行处理  PANEL_OUTPUTFILEPATH
                DirectoryInfo spath = new DirectoryInfo(PANEL_INPUTFILEPATH); 
                List<FileInformation> xmlfilesname = FileHelper.DirectoryAllFiles.GetAllFiles(spath, ".xml");

                if (xmlfilesname != null && xmlfilesname.Count > 0)
                {
                    //获取taglist全部信息 不进行重复抽取数据的操作
                    JArray ret_jsonobj = new JArray();
                    string sqlstr = "select DEVINDEX,SBDM,beizhu3 as SYSID,DEVID as CLASSID from devicelist where stationid='" + stationName + "' and devindex<>''  and devindex<>'' and sysid<>'' and classid<>'';";
                    ret_jsonobj = get_jsonobj_bysqlstr(sqlstr);
                    if (ret_jsonobj == null)
                    {
                        ShowInfo(stationName + "的数据未导入，无法自动生成相关画面。");
                        return;
                    }
                    foreach (FileInformation fileinfo in xmlfilesname)
                    {
                        CreatePanel(fileinfo.FilePath, stationName, stationShowName, ret_jsonobj);
                    }

                }
                else
                {
                    ShowInfo("提示：路径"+PANEL_INPUTFILEPATH+"下没有任何xml画面文件。");
                }
            }
            catch (Exception ex)
            {
                ShowInfo("生成panel画面文件时出错：" + ex.Message);
                
            }

        }

        //获取全部的devicelist的数据 转成jsonobject
        //通用方法 传递查询语句，返回对象属性根据字段名称来生成
        public JArray get_jsonobj_bysqlstr(string sql_str)
        {
            JArray ret_jsonobj = new JArray();
            try
            {
                using (_dbcon = new DBLib.DBLib(HOST, USER, PASSWORD, DBNAME, int.Parse(DBTYPE)))
                {
                    DataSet ds_Result = _dbcon.GetData(sql_str);

                    if (ds_Result != null)
                    {
                        if (ds_Result != null && ds_Result.Tables.Count != 0 && ds_Result.Tables[0].Rows.Count != 0)
                        {
                            foreach (DataRow mDr in ds_Result.Tables[0].Rows)
                            {
                                JObject jsonset_t = new JObject();
                                foreach (DataColumn mDc in ds_Result.Tables[0].Columns)
                                {
                                    jsonset_t.Add(new JProperty(mDc.ColumnName, mDr[mDc].ToString()));
                                }   
                                //json数组结果集
                                ret_jsonobj.Add(jsonset_t);
                            }
                            return ret_jsonobj;

                        }
                        else
                        {
                            ShowInfo("查询语句未查到数据：" + sql_str);
                            return null;
                        }
                    }
                    else
                    {
                        ShowInfo("查询语句执行出错：" + sql_str);
                        return null;
                    }
                }

            }
            catch (Exception ex)
            {
                ShowInfo("获取数据信息时出错：" + ex.Message);
                return null;
            }

        }

        

        //生成画面文件 自动保存到给定路径下
        public void CreatePanel(string sPath, string stationName,string stationShowName, JArray taginfo_jsonobj)
        {
            string output_filename="";//不含后缀文件名
            try
            {
                
                if (!File.Exists(sPath))
                {
                    this.ShowInfo("文件:" + sPath + "已丢失，不存在。");
                    return;
                }
                output_filename = Path.GetFileNameWithoutExtension(sPath);//不含后缀文件名

                //获取上次保存的画面信息
                string urlstr_t= httpurl_getGraphicPositionByParams;
                HttpRequestEx httpobjr_t = new HttpRequestEx();
                string retr_t = string.Empty;
                retr_t = httpobjr_t.HttpRequest_Call(urlstr_t, "graphicName=" + output_filename);//同步方式调用
                //返回画面信息
                //{"rows":[{"devIndex":"MXLR_TEST","graphicName":"xlr_panel_test","pid":2,
                //"position":"位置1","positionCode":"位置2","sbdm":"设备代码","updateTime":"2017-08-31 14:05:51"}],"total":1}
                JObject retobj=new JObject();
                JArray json_panelinfo_get=new JArray();
                if (retr_t != "")
                {
                    retobj = (JObject)JsonConvert.DeserializeObject(retr_t);//返回的画面信息对象
                    json_panelinfo_get = JArray.Parse(retobj.SelectToken("rows").ToString());//画面信息取出
                }
                

                JArray json_ary = new JArray();//上传平台的画面信息json队列

                //读取xml文件
                XmlDocument xdPanelExample = new XmlDocument();  //实例化一个XmlDocument
                xdPanelExample.Load(sPath);
                //临时xml文件
                XmlDocument xdPanel_t = new XmlDocument();
                xdPanel_t.Load(sPath);

                //shapes节点
                XmlNode ShapesNode = xdPanelExample.SelectSingleNode("//panel/shapes");

                //背景图片 设置 暂时空白 如果没有该节点 需要自动生成（后续考虑）
                XmlNode xmlnode_panelbackground = xdPanelExample.SelectSingleNode("//panel/properties/prop[@name='Image']/prop");
                if (xmlnode_panelbackground != null)
                {
                    xmlnode_panelbackground.InnerText = "background/" + stationName + "_*.png";
                }

                //背景颜色设置  <prop name="BackColor">{38,46,60}</prop>
                xmlnode_panelbackground = xdPanelExample.SelectSingleNode("//panel/properties/prop[@name='BackColor']");
                if (xmlnode_panelbackground != null)
                {
                    xmlnode_panelbackground.InnerText = PANEL_BACKGROUND_COLOR;

                }
                int referenceId = 1;//保证唯一
                int taborder = 1000;//画面元素tab顺序 任何shape对象的TabOrder唯一 serialId唯一
                string tagname = "";//图元绑定的tagname
                string classid = "";//图元的路径
                string sysid = "";//设备所属系统id
                string Geometry = "1 0 0 1 0 0";//相对位置
                //最大最小x y
                int max_x = 1900;
                int max_y = 855;

                //在上次画面生成时未出现的设备代码
                List<string> sSBDM_new = new List<string>();
                //没有对应点表信息的
                List<string> sSBDM_nosetting = new List<string>();
                //获取所有 PRIMITIVE_TEXT
                XmlNodeList text_replace_list = xdPanelExample.SelectNodes("/panel/shapes/shape[@shapeType='PRIMITIVE_TEXT']");

                //遍历获取能够对应点表的图元
                XmlNodeList refer_old_t = xdPanelExample.SelectNodes("/panel/shapes/reference");
                //删除所有的reference图元
                foreach (XmlNode refer_t in refer_old_t)
                {

                    ShapesNode.RemoveChild(refer_t);
                }

                foreach (XmlNode text_label in text_replace_list)
                {
                    
                    //获取 PRIMITIVE_TEXT xmlnode （正常监控画面应该都有）
                    XmlNode xmlnode_t = text_label.SelectSingleNode("./properties/prop[@name='Text']/prop[@name='zh_CN.utf8']");
                    if (xmlnode_t != null)
                    {

                        
                        //设备编号获取 图纸有可能有大小写混合的情况
                        //20170926 设备代码panel上面可能有空格，判断的时候需要剔除空格
                        string device_code = xmlnode_t.InnerText.Trim().ToUpper();
                        
                        //根据设备代码 获取对应的 tagname（点表名） 和 filename（设备类型）
                        tagname = "";
                        classid = "";
                        sysid = "";
                        foreach (var r in taginfo_jsonobj)
                        {
                            if (device_code == (string)r.SelectToken("SBDM"))
                            {
                                tagname = (string)r.SelectToken("DEVINDEX");
                                classid = (string)r.SelectToken("CLASSID");//根据大类来生成图元的路径
                                sysid = (string)r.SelectToken("SYSID");
                                break;
                            }
                            
                        }
                        if(tagname=="" || classid == "" || sysid == "")
                        {
                            sSBDM_nosetting.Add(device_code);
                            continue;//未匹配到结果
                        }
                        //Location和Geometry先获取界面上已经存在的图元 在临时对象上完成
                        XmlNode refrence_t = xdPanel_t.SelectSingleNode("/panel/shapes/reference/properties/prop[@name='dollarParameters']/prop[@name='dollarParameter'][prop='"+tagname+"']/prop");
                        XmlNode parentnode_t;//已经存在的图元属性元素节点
                        int i_exsitRefer = 0;
                        //Location获取
                        if (refrence_t != null)
                        {
                            //已经完成的图元的位置
                            parentnode_t = refrence_t.ParentNode.ParentNode.ParentNode;
                            xmlnode_t = parentnode_t.SelectSingleNode("./prop[@name='Location']");
                            i_exsitRefer = 1;

                        }
                        else
                        {
                            //text位置获取
                            xmlnode_t = text_label.SelectSingleNode("./properties/prop[@name='Location']");
                          
                        }

                        string[] location_arr;
                        string x = "0";
                        string y = "0";

                        if (xmlnode_t != null)
                        {
                            location_arr = xmlnode_t.InnerText.Split(' ');
                            x = location_arr[0].ToString();
                            y = location_arr[1].ToString();
                        }

                        //x,y的偏移范围界定 画面尺寸大部分1920*875
                        //画面的尺寸判断可选，每个站有2-3副画面会非常大。
                        if (checkBox_panelsize.Checked)
                        {
                            if (Convert.ToDouble(x) < 0)
                            {
                                x = "0";
                            }
                            else if (Convert.ToDouble(x) > max_x)
                            {

                                x = max_x.ToString();
                            }
                            if (Convert.ToDouble(y) < 0)
                            {
                                y = "0";
                            }
                            else if (Convert.ToDouble(y) > max_y)
                            {
                                y = max_y.ToString();
                            }
                        }
                        


                        //Geometry获取
                        if (refrence_t != null)
                        {
                            //已经完成的图元的位置
                            parentnode_t = refrence_t.ParentNode.ParentNode.ParentNode;
                            xmlnode_t = parentnode_t.SelectSingleNode("./prop[@name='Geometry']");
                            i_exsitRefer = 1;

                        }
                        else
                        {
                            //
                            xmlnode_t = text_label.SelectSingleNode("./properties/prop[@name='Geometry']");
                            
                        }
                        //部分text没有Geometry 需要初始化初始值一次
                        Geometry = "1 0 0 1 0 0";
                        if (xmlnode_t != null) Geometry = xmlnode_t.InnerText;

                        //上传http接口的panel信息
                        json_ary.Add(new JObject(new JProperty("devIndex", tagname),
                                                 new JProperty("position", x + " " + y),
                                                 new JProperty("positionCode", Geometry),
                                                 new JProperty("sbdm", device_code),
                                                 new JProperty("graphicName", output_filename)
                                                 )
                                    );

                        //画面信息对比 根据sbdm进行对比：位置/有无
                        if (json_panelinfo_get != null)//接口获取画面的信息不为空
                        {
                            int sFlag = 0;
                            string sDate_t = "";
                            for (int i = 0; i < json_panelinfo_get.Count; i++)
                            {
                                sDate_t = (string)json_panelinfo_get[i].SelectToken("updateTime");//获取最近一次画面更新时间
                                //SelectToken 方法使用
                                if (device_code== (string)json_panelinfo_get[i].SelectToken("sbdm"))//记录匹配情况下，对比位置
                                {
                                    if((string)json_panelinfo_get[i].SelectToken("position")== x + " " + y&& (string)json_panelinfo_get[i].SelectToken("positionCode")== Geometry)
                                    {
                                        ((JObject)json_panelinfo_get[i])["pid"] = "exist_ok";
                                    }
                                    else
                                    {
                                        ((JObject)json_panelinfo_get[i])["pid"] = "exist_positiondiff";
                                    }
                                    sFlag = 1;//匹配到记录 不是新增


                                }
    
                            }
                            if (sFlag == 0)//未匹配到的设备代码
                            {
                                sSBDM_new.Add(device_code);//新增设备
                            }
                            
                        }

                   

                        //根据设备编号获取对应图元信息
                        //1.referenceId / Name  /TabOrder 编号都要唯一
                        //2.FileName 根据类型生成
                        //3.Location 用x,y生成。 Geometry / 固定格式： 1 0 0 1 0 0
                        //4.dollarParameters - dollarParameter - Value 获取 tagname

                        string xpath_t = "/panel/shapes/reference[@referenceId='" + referenceId.ToString() + "']/properties/";
                        XMLHelper.XMLHelper.Set(xdPanelExample, xpath_t + "prop[@name='FileName']", "objects/"+ sysid + "/"+ classid + ".xml");//objects/EMCS/DDT.xml
                        XMLHelper.XMLHelper.Set(xdPanelExample, xpath_t + "prop[@name='Location']", (Convert.ToDouble(x) + 0).ToString() + " "+ (Convert.ToDouble(y) + 0).ToString());
                        XMLHelper.XMLHelper.Set(xdPanelExample, xpath_t + "prop[@name='Geometry']", Geometry);//
                        XMLHelper.XMLHelper.Set(xdPanelExample, xpath_t + "prop[@name='TabOrder']", taborder.ToString());
                        //
                        XMLHelper.XMLHelper.Set(xdPanelExample, xpath_t + "prop[@name='dollarParameters']/prop[@name='dollarParameter']/prop[@name='Dollar']", "$dp");
                        XMLHelper.XMLHelper.Set(xdPanelExample, xpath_t + "prop[@name='dollarParameters']/prop[@name='dollarParameter']/prop[@name='Value']", tagname);
                        //设置reference 图元属性 
                        XmlNode refnode_t = xdPanelExample.SelectSingleNode("/panel/shapes/reference[@referenceId='" + referenceId.ToString() + "']");//第一个
                        if (refnode_t != null)
                        {
                            XmlElement refnode_element = (XmlElement)refnode_t;
                            refnode_element.SetAttribute("Name", "PANEL_REF"+ referenceId.ToString());
                            refnode_element.SetAttribute("parentSerial", "-1"); 
                        }

                        //id自增
                        referenceId++;
                        taborder++;
                    }


                }
                //画面所含设备变化提示通知
                if (json_panelinfo_get != null)
                {
                    string sMsg = "";
                    string sDate_t = "";
                    List<string> sSBDM_t1 = new List<string>();//位置发生变化
                    List<string> sSBDM_t2 = new List<string>();//信息一致
                    List<string> sSBDM_t3 = new List<string>();//较上次减少设备
                    for (int i = 0; i < json_panelinfo_get.Count; i++)
                    {                                                                     
                        if ((string)json_panelinfo_get[i].SelectToken("pid")== "exist_positiondiff")//记录匹配情况下，对比位置
                        {
                            sDate_t = (string)json_panelinfo_get[i].SelectToken("updateTime");//获取最近一次画面更新时间
                            sSBDM_t1.Add((string)json_panelinfo_get[i].SelectToken("sbdm"));
                        }
                        else if((string)json_panelinfo_get[i].SelectToken("pid") == "exist_ok")
                        {
                            sDate_t = (string)json_panelinfo_get[i].SelectToken("updateTime");
                            sSBDM_t2.Add((string)json_panelinfo_get[i].SelectToken("sbdm"));
                        }
                        else
                        {
                            sDate_t = (string)json_panelinfo_get[i].SelectToken("updateTime");
                            sSBDM_t3.Add((string)json_panelinfo_get[i].SelectToken("sbdm"));
                        }

                    }
                    if (sSBDM_t1.Count>0)
                    {
                        sMsg = "提示：画面" + output_filename + "中的设备代码" + string.Join(",", sSBDM_t1.ToArray()) + "的位置与" + sDate_t + "的画面生成操作不同。";
                        ShowInfo(sMsg);
                    }
                    if (sSBDM_t2.Count > 0)
                    {
                        sMsg = "提示：画面" + output_filename + "中的设备代码" + string.Join(",", sSBDM_t2.ToArray()) + "与" + sDate_t + "的画面生成操作中的信息一致。";
                        ShowInfo(sMsg);
                        
                    }
                    if (sSBDM_t3.Count > 0)
                    {
                        sMsg = "提示：画面" + output_filename + "与" + sDate_t + "生成时设备代码：" + string.Join(",", sSBDM_t3.ToArray()) + "在本次生成时未出现。";
                        ShowInfo(sMsg);
                    }
                    if (sSBDM_new.Count > 0)
                    {
                        sMsg = "提示：画面" + output_filename + "本次生成新增设备代码：" + string.Join(",", sSBDM_new.ToArray()) ;
                        ShowInfo(sMsg);
                    }
                }
                if (sSBDM_nosetting.Count > 0)
                {
                    ShowInfo("画面" + output_filename + "本次生成中设备代码：" + string.Join(",", sSBDM_nosetting.ToArray()) + "未获取到对应的点表信息。");
                }

                //画面记录信息
                if(json_ary.Count>0)
                {
                    string urlstr = httpurl_addGraphicPosition;
                    HttpRequestEx httpobj = new HttpRequestEx();
                    string ret = string.Empty;
                    //ret = httpobj.HttpRequest_Call(urlstr, json_ary.ToString());//同步方式调用
                    httpobj.HttpRequest_AsyncCall(urlstr, "posList=" + json_ary.ToString());//异步方式调用
                    //HttpWeb_SavePanelInfo(urlstr, json_ary);
                    ShowInfo("画面"+ output_filename + "信息保存操作，调用服务（" + urlstr + "）返回：" + ret);
                }

                //保存xml文件
                sPath = sPath.Replace(PANEL_INPUTFILEPATH, PANEL_OUTPUTFILEPATH);//可能有嵌套文件夹的情况产生
                string sPath_t = Path.GetDirectoryName(sPath);
                if (!Directory.Exists(sPath_t))
                {
                    //
                    Directory.CreateDirectory(sPath_t);

                }
                
                xdPanelExample.Save(sPath);//
                
            }
   
            catch (Exception ex)
            {
                this.ShowInfo("生成监视画面"+ output_filename + "时出错：" + ex.Message);

            }
        }

        public string HttpWeb_SavePanelInfo(string urlstr, JArray json_ary) {

            try
            {

                string send_jsonstr = json_ary.ToString();
                HttpWebRequest request = WebRequest.Create(urlstr) as HttpWebRequest;//
                request.Timeout = 3111;
                request.Method = "post";
                request.KeepAlive = true;
                request.AllowAutoRedirect = false;
                request.ContentType = "application/x-www-form-urlencoded";
                byte[] postdatabtyes = Encoding.UTF8.GetBytes("posList="+ send_jsonstr);
                request.ContentLength = postdatabtyes.Length;
                Stream requeststream = request.GetRequestStream();
                requeststream.Write(postdatabtyes, 0, postdatabtyes.Length);
                requeststream.Close();
                string resp;
            
                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    resp = sr.ReadToEnd();
                }
                return resp;
            }
            catch(Exception ex)
            {
                ShowInfo("调用HttpWeb服务("+ urlstr + ")出错:" + ex.Message);
                return "";
            }
            
        }

        //调用java http接口 后续使用httpclient调用
        private void button4_Click(object sender, EventArgs e)
        {

            string urlstr = "http://128.64.90.37:8084/autotools/dataAccess/getGraphicPositionByParams/";
            try
            {

                string send_jsonstr = "xlr_panel_test";
                HttpWebRequest request = WebRequest.Create(urlstr) as HttpWebRequest;//
                request.Method = "post";
                request.KeepAlive = true;
                request.AllowAutoRedirect = false;
                request.ContentType = "application/x-www-form-urlencoded";
                byte[] postdatabtyes = Encoding.UTF8.GetBytes("graphicName=" + send_jsonstr);
                request.ContentLength = postdatabtyes.Length;
                Stream requeststream = request.GetRequestStream();
                requeststream.Write(postdatabtyes, 0, postdatabtyes.Length);
                requeststream.Close();
                string resp;

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    resp = sr.ReadToEnd();
                }
                JObject retobj =(JObject)JsonConvert.DeserializeObject(resp);
                JArray jet = JArray.Parse(retobj.SelectToken("rows").ToString());

            }
            catch (Exception ex)
            {
                ShowInfo("调用HttpWeb服务(" + urlstr + ")出错:" + ex.Message);
                
            }
            
        }

        public void button5_Click(object sender, EventArgs e)
        {
            JArray json_ary = new JArray();
            json_ary.Add(new JObject(new JProperty("devIndex", "MXLR_TEST"),
                                                 new JProperty("position","位置1"),
                                                 new JProperty("positionCode", "位置2"),
                                                 new JProperty("sbdm", "设备代码"),
                                                 new JProperty("graphicName", "xlr_panel_test")
                                                 )
                                    );
            //添加用户信息
            string urlstr = httpurl_addGraphicPosition;
            HttpRequestEx r = new HttpRequestEx();
            string ret= r.HttpRequest_Call(urlstr, json_ary.ToString());
            HttpWeb_SavePanelInfo(urlstr, json_ary);
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }
    }
}


