using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Linq;
using System.Timers;
using IBM.Data.DB2;//DLL方式连接DB2
using System.Data.SQLite;
using MySql.Data;
using MySql.Data.MySqlClient;

/*
 * 数据库操作类
 * 事务方式处理sql语句
 * 以dataset方式返回数据
 * 调用存储过程
 */
namespace DBLib
{
    public class DBLib : IDisposable
    {
        //数据库类型 20170705 增加sqlite处理
        public enum DataBaseType { Oracle = 0, SqlServer, Access,Excel,DB2,IBMDB2DLL,sqllite,mysql, none };
        //数据类型
        public enum DataType { Int = 0, Float, String, DateTime };



        #region 内部变量
        private SqlConnection sqlcon = null;//Sql数据类
        private OleDbConnection oledbcon = null;//Oledb连接类
        private DB2Connection DB2DLLcon = null;//DLL方式连接DB2
        private SQLiteConnection sqlitecon=null;//sqlite 连接类
        private MySqlConnection mysql = null;//ado.net mysql from Nuget

        private DataBaseType dbtype;
        private string host = null;//数据库地址
        private string database = null;//数据库名称
        private string user = null;//数据库用户
        private string password = null;//数据库密码

        //2013-05-31 添加持续连接标记位 允许使用持续连接进行数据库操作
        private bool isStanding = false;//默认采用短连接方式
        private object dbLocker = new object();// 数据库锁 用于建立连接时进行同步
        #endregion

        #region 属性
        public SqlConnection SqlCon
        {
            get
            {
                return sqlcon;
            }
        }


        public OleDbConnection OledbBCon
        {
            get
            {
                return oledbcon;
            }
        }

        public DataBaseType DBType
        {
            get
            {
                return dbtype;
            }
        }

        public string Host
        {
            get
            {
                return host;
            }
        }

        public string DataBase
        {
            get
            {
                return database;
            }
        }

        public string User
        {
            get
            {
                return user;
            }
        }

        public string PassWord
        {
            get
            {
                return password;
            }
        }
        #endregion

        

        #region 构造函数
        public DBLib() { }
        public DBLib(string s_ConnectionString, DataBaseType dbtype)
        {
            this.dbtype = dbtype;
            InitDBLib(s_ConnectionString, dbtype);           
        }

        /// <summary>
        /// DBLib构造函数
        /// </summary>
        /// <param name="s_Host">数据库主机地址</param>
        /// <param name="s_User">数据库用户名</param>
        /// <param name="s_PassWord">数据库密码</param>
        /// <param name="s_dbname">数据库名</param>
        /// <param name="dbt_type">数据库类型</param>
        public DBLib(string s_Host, string s_User, string s_PassWord, string s_dbname, DataBaseType dbt_type)
        {
            string s_constr = "";
            if (dbt_type == DataBaseType.SqlServer)
            {
                s_constr = "Data Source=" + s_Host + ";Initial Catalog=" + s_dbname + ";User ID=" + s_User + ";Password=" + s_PassWord;
            }
            else if (dbt_type == DataBaseType.Oracle)
            {
                s_constr = "Provider=msdaora;Data Source=" + s_dbname + ";User Id=" + s_User + ";Password=" + s_PassWord;
            }
            else if (dbt_type == DataBaseType.Access)
            {
                s_constr = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + s_dbname;
            }
            else if (dbt_type == DataBaseType.DB2)
            {
                s_constr = "Provider=IBMDADB2;Database="+s_dbname+";Hostname="+s_Host+";Protocol=TCPIP; Port=50000;Uid="+s_User+";Pwd="+s_PassWord+";";
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                s_constr = "Persist Security Info=True;Server=" + s_Host + ";Database=" + s_dbname + ";User ID=" + s_User + ";Password=" + s_PassWord + ";CurrentSchema=ADMINISTRATOR;";
            }
            else if (dbtype == DataBaseType.sqllite)
            {
                //s_Host数据文件路径
                s_constr = "Data Source=" + s_Host + ";Version=3;";//连接密码 Password="";设置连接池 Pooling=False;Max Pool Size=100;//只读连接Read Only=false";
            }
            else
            {
                s_constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + s_dbname + "';Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            }

            InitDBLib(s_constr, dbt_type);
        }


        /// <summary>
        /// DBLib构造函数
        /// </summary>
        /// <param name="s_Host">数据库主机地址</param>
        /// <param name="s_User">数据库用户名</param>
        /// <param name="s_PassWord">数据库密码</param>
        /// <param name="s_dbname">数据库名</param>
        /// <param name="flag">数据库类型标识 0-oracle 1-sql</param>
        public DBLib(string s_Host, string s_User, string s_PassWord, string s_dbname, int flag)
        {
            host = s_Host;
            user = s_User;
            password = s_PassWord;
            database = s_dbname;
            if (flag == 0)
            {
                dbtype = DataBaseType.Oracle;
            }
            else if (flag == 1)
            {
                dbtype = DataBaseType.SqlServer;
            }
            else if (flag == 2)
            {
                dbtype = DataBaseType.Access;
            }
            else if (flag == 3)
            {
                dbtype = DataBaseType.Excel;
            }
            else if (flag == 4)
            {
                dbtype = DataBaseType.DB2;
            }
            else if (flag == 5)
            {
                dbtype = DataBaseType.IBMDB2DLL;
            }
            else if (flag == 6)
            {
                dbtype = DataBaseType.sqllite;
            }
            string s_constr = "";
            if (dbtype == DataBaseType.SqlServer)
            {
                s_constr = "Data Source=" + s_Host + ";Initial Catalog=" + s_dbname + ";User ID=" + s_User + ";Password=" + s_PassWord;
                sqlcon = new SqlConnection(s_constr);
            }
            else if (dbtype == DataBaseType.Oracle)
            {
                s_constr = "Provider=msdaora;Data Source=" + s_dbname + ";User Id=" + s_User + ";Password=" + s_PassWord;
                oledbcon = new OleDbConnection(s_constr);
            }
            else if (dbtype == DataBaseType.Access)
            {
                s_constr = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + s_dbname;
                oledbcon = new OleDbConnection(s_constr);
            }
            else if (dbtype == DataBaseType.DB2)
            {
                s_constr = "Provider=IBMDADB2;Database=" + s_dbname + ";Hostname=" + s_Host + ";Protocol=TCPIP; Port=50000;Uid=" + s_User + ";Pwd=" + s_PassWord + ";";
                oledbcon = new OleDbConnection(s_constr);
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                s_constr = s_constr = "Persist Security Info=True;Server=" + s_Host + ";Database=" + s_dbname + ";User ID=" + s_User + ";Password=" + s_PassWord + ";CurrentSchema=ADMINISTRATOR;";
                DB2DLLcon = new IBM.Data.DB2.DB2Connection(s_constr);
            }
            else if (dbtype == DataBaseType.sqllite)//sqlite 20170705
            {
                s_constr = "Data Source=" + s_Host + ";Version=3;";//连接密码 Password="";设置连接池 Pooling=False;Max Pool Size=100;//只读连接Read Only=false";
                sqlitecon = new SQLiteConnection(s_constr);
            }
            else
            {
                s_constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + s_dbname + "';Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
                oledbcon = new OleDbConnection(s_constr);
            }
        }

        /// <summary>
        /// DBLib构造函数
        /// </summary>
        /// <param name="s_Host">数据库主机地址</param>
        /// <param name="s_User">数据库用户名</param>
        /// <param name="s_PassWord">数据库密码</param>
        /// <param name="s_dbname">数据库名</param>
        /// <param name="flag">数据库类型标识 0-oracle 1-sql</param>
        /// <param name="b_Standing">持续连接标记位</param>
        public DBLib(string s_Host, string s_User, string s_PassWord, string s_dbname, int flag,bool b_Standing)
        {
            host = s_Host;
            user = s_User;
            password = s_PassWord;
            database = s_dbname;
            if (flag == 0)
            {
                dbtype = DataBaseType.Oracle;
            }
            else if (flag == 1)
            {
                dbtype = DataBaseType.SqlServer;
            }
            else if (flag == 2)
            {
                dbtype = DataBaseType.Access;
            }
            else if (flag == 3)
            {
                dbtype = DataBaseType.Excel;
            }
            else if (flag == 4)
            {
                dbtype = DataBaseType.DB2;
            }
            else if (flag == 5)
            {
                dbtype = DataBaseType.IBMDB2DLL;
            }
            else if (flag == 6)
            {
                dbtype = DataBaseType.sqllite;
            }
            string s_constr = "";
            if (dbtype == DataBaseType.SqlServer)
            {
                s_constr = "Data Source=" + s_Host + ";Initial Catalog=" + s_dbname + ";User ID=" + s_User + ";Password=" + s_PassWord;
                sqlcon = new SqlConnection(s_constr);
            }
            else if (dbtype == DataBaseType.Oracle)
            {
                s_constr = "Provider=msdaora;Data Source=" + s_dbname + ";User Id=" + s_User + ";Password=" + s_PassWord;
                oledbcon = new OleDbConnection(s_constr);
            }
            else if (dbtype == DataBaseType.Access)
            {
                s_constr = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + s_dbname;
                oledbcon = new OleDbConnection(s_constr);
            }
            else if (dbtype == DataBaseType.DB2)
            {
                s_constr = "Provider=IBMDADB2;Database=" + s_dbname + ";Hostname=" + s_Host + ";Protocol=TCPIP; Port=50000;Uid=" + s_User + ";Pwd=" + s_PassWord + ";";
                oledbcon = new OleDbConnection(s_constr);
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                s_constr = s_constr = "Persist Security Info=True;Server=" + s_Host + ";Database=" + s_dbname + ";User ID=" + s_User + ";Password=" + s_PassWord + ";CurrentSchema=ADMINISTRATOR;";
                DB2DLLcon = new IBM.Data.DB2.DB2Connection(s_constr);
            }
            else if (dbtype == DataBaseType.sqllite)//sqlite 20170705
            {
                s_constr = "Data Source=" + s_Host + ";Version=3;";//连接密码 Password="";设置连接池 Pooling=False;Max Pool Size=100;//只读连接Read Only=false";
                sqlitecon = new SQLiteConnection(s_constr);
            }
            else
            {
                s_constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + s_dbname + "';Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
                oledbcon = new OleDbConnection(s_constr);
            }

            isStanding = b_Standing;
        }

        /// <summary>
        /// 使用长连接时定时调用trycon方式确保连接正常
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void tmConCheck_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (isStanding)
            {
                int i_DBState = this.trycon();
            }
        }

        private void InitDBLib(string s_ConnectionString, DataBaseType dbtype)
        {
            if (string.IsNullOrEmpty(s_ConnectionString))
                return;

            if (dbtype == DataBaseType.SqlServer)
            {
                sqlcon = new SqlConnection(s_ConnectionString);
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                DB2DLLcon = new IBM.Data.DB2.DB2Connection(s_ConnectionString);
            }
            else if (dbtype == DataBaseType.sqllite)//sqlite 20170705
            {
                sqlitecon = new SQLiteConnection(s_ConnectionString);
            }
            else
            {
                oledbcon = new OleDbConnection(s_ConnectionString);
            }

        }


        #endregion

        #region 数据库操作
        /// <summary>
        /// 测试数据库连接
        /// </summary>
        /// <returns></returns>
        public int trycon()// 测试数据库连接
        {
            int i_result = 0;
            if (dbtype == DataBaseType.SqlServer)
            {
                try
                {
                    if (sqlcon.State != ConnectionState.Open)
                    {
                        sqlcon.Open();
                    }
                    i_result = 1;
                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        sqlcon.Close();
                    }
                }
                return i_result;
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                try
                {
                    if (DB2DLLcon.State != ConnectionState.Open)
                    {
                        DB2DLLcon.Open();
                    }
                    i_result = 1;
                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        DB2DLLcon.Close();
                    }

                }
                return i_result;
                
            
            }
            else if (dbtype == DataBaseType.sqllite)//sqlite 20170705
            {
                try
                {
                    if (sqlitecon.State != ConnectionState.Open)
                    {
                        sqlitecon.Open();
                    }
                    i_result = 1;
                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    
                    if (!isStanding)
                    {
                        sqlitecon.Close();
                    }

                }
                return i_result;
            }
            else
            {
                try
                {
                    if (oledbcon.State != ConnectionState.Open)
                    {
                        oledbcon.Open();
                    }
                    i_result = 1;
                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        oledbcon.Close();
                    }
                    
                }
                return i_result;
            }
        }

        /// <summary>
        /// 执行Sql语句
        /// </summary>
        /// <param name="sqls">sql语句数组</param>
        /// <returns></returns>
        public int ExcuteSqls(string[] sqls)// 以事务方式处理sql语句
        {
            int i_result = 0;
            if (dbtype == DataBaseType.SqlServer)
            {
                SqlTransaction sc_sqltran = null;
                try
                {
                    if (sqlcon.State != ConnectionState.Open)
                    {
                        sqlcon.Open();
                    }
                    SqlCommand sc_command = new SqlCommand();
                    sc_command.CommandType = System.Data.CommandType.Text;
                    sc_command.Connection = sqlcon;
                    sc_sqltran = sqlcon.BeginTransaction();
                    sc_command.Transaction = sc_sqltran;
                    foreach (string s_sql in sqls)
                    {
                        sc_command.CommandText = s_sql;
                        i_result += sc_command.ExecuteNonQuery();
                    }
                    sc_sqltran.Commit();
                }
                catch
                {

                    //if (sc_sqltran != null)
                    //    sc_sqltran.Rollback();

                    try
                    {
                        sc_sqltran.Rollback();
                    }
                    catch (SqlException ex)
                    {
                        if (sc_sqltran.Connection != null)
                        {
                            //Console.WriteLine("回滚失败! 异常类型: " + ex.GetType());
                            throw new Exception("回滚失败，异常信息: " + ex.Message);
                            
                        }
                    } 
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        sqlcon.Close();
                    }
                }
                return i_result;
            }
            else if (dbtype == DataBaseType.IBMDB2DLL) //IBM DLL方式
            {
                DB2Transaction myDB2Command_tran = null;
                try
                {
                    if (DB2DLLcon.State != ConnectionState.Open)
                    {
                        DB2DLLcon.Open();
                    }
                    
                    DB2Command myDB2Command = new DB2Command();
                    myDB2Command.CommandType = System.Data.CommandType.Text;
                    myDB2Command.Connection = DB2DLLcon;
                    myDB2Command_tran = DB2DLLcon.BeginTransaction();
                    myDB2Command.Transaction = myDB2Command_tran;

                    foreach (string s_sql in sqls)
                    {
                        myDB2Command.CommandText = s_sql;
                        i_result += myDB2Command.ExecuteNonQuery();
                    }
                    myDB2Command_tran.Commit();
                }
                catch
                {
                   
                    try
                    {
                        myDB2Command_tran.Rollback();
                    }
                    catch (SqlException ex)
                    {
                        if (myDB2Command_tran.Connection != null)
                        {
                            //Console.WriteLine("回滚失败! 异常类型: " + ex.GetType());
                            throw new Exception("回滚失败，异常信息: " + ex.Message);
                            
                        }
                    }
                    i_result = -1;
                }
                finally
                {
                    if (!isStanding)
                    {
                        DB2DLLcon.Close();
                    }
                    
                }
                return i_result;
            }
            else if (dbtype == DataBaseType.sqllite) //sqlite 事务
            {
                SQLiteTransaction sqliteCommand_tran = null;
                try
                {
                    if (sqlitecon.State != ConnectionState.Open)
                    {
                        sqlitecon.Open();
                    }

                    SQLiteCommand sqliteCommand = new SQLiteCommand();
                    sqliteCommand.CommandType = System.Data.CommandType.Text;
                    sqliteCommand.Connection = sqlitecon;
                    sqliteCommand_tran = sqlitecon.BeginTransaction();
                    sqliteCommand.Transaction = sqliteCommand_tran;

                    foreach (string s_sql in sqls)
                    {
                        sqliteCommand.CommandText = s_sql;
                        i_result += sqliteCommand.ExecuteNonQuery();
                    }
                    sqliteCommand_tran.Commit();
                }
                catch
                {

                    try
                    {
                        sqliteCommand_tran.Rollback();
                    }
                    catch (SqlException ex)
                    {
                        if (sqliteCommand_tran.Connection != null)
                        {
                            //Console.WriteLine("回滚失败! 异常类型: " + ex.GetType());
                            throw new Exception("回滚失败，异常信息: " + ex.Message);
                            
                        }
                    }
                    i_result = -1;
                }
                finally
                {
                    if (!isStanding)
                    {
                        sqlitecon.Close();
                    }
                }
                return i_result;
            }
            else
            {
                OleDbTransaction ole_tran = null;
                try
                {
                    if (oledbcon.State != ConnectionState.Open)
                    {
                        oledbcon.Open();
                    }
                    OleDbCommand ole_command = new OleDbCommand();
                    ole_command.CommandType = System.Data.CommandType.Text;
                    ole_command.Connection = oledbcon;
                    ole_tran = oledbcon.BeginTransaction();
                    ole_command.Transaction = ole_tran;
                    foreach (string s_sql in sqls)
                    {
                        ole_command.CommandText = s_sql;
                        i_result += ole_command.ExecuteNonQuery();
                    }
                    ole_tran.Commit();
                }
                catch
                {
                    //if (ole_tran != null)
                    //    ole_tran.Rollback();

                    try
                    {
                        ole_tran.Rollback();
                    }
                    catch (SqlException ex)
                    {
                        if (ole_tran.Connection != null)
                        {
                            //Console.WriteLine("回滚失败! 异常类型: " + ex.GetType());
                            throw new Exception("回滚失败，异常信息: " + ex.Message);
                            
                        }
                    } 
                    i_result = -1;
                }
                finally
                {

                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        oledbcon.Close();
                    }
                    
                }
                return i_result;
            }
        }

        public int[] ExcuteSqls(string[] sqls, int packsize)//分包处理大批量数据
        {
            try
            {
                List<int> l_result = new List<int>();//执行结果
                List<string> l_subsqls = new List<string>();
                int count = 0;
                while (count < sqls.Length)
                {
                    l_subsqls.Add(sqls[count]);
                    count = count + 1;
                    if (l_subsqls.Count == packsize)
                    {
                        int int_subresult = this.ExcuteSqls(l_subsqls.ToArray());
                        l_result.Add(int_subresult);
                        l_subsqls.Clear();
                    }
                }
                if (l_subsqls.Count > 0)
                {
                    int int_subresult = this.ExcuteSqls(l_subsqls.ToArray());
                    l_result.Add(int_subresult);
                    l_subsqls.Clear();
                }
                return l_result.ToArray();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 执行Sql语句
        /// </summary>
        /// <param name="sqls">sql语句</param>
        /// <returns></returns>
        public int ExcuteSql(string s_sql)// 处理单条sql语句
        {
            int i_result = 0;
            if (dbtype == DataBaseType.SqlServer)
            {
                try
                {
                    
                    if (sqlcon.State != ConnectionState.Open)
                    {
                        sqlcon.Open();
                    }
                    SqlCommand sc_command = new SqlCommand();
                    sc_command.CommandType = System.Data.CommandType.Text;
                    sc_command.Connection = sqlcon;
                    sc_command.CommandText = s_sql;
                    i_result = sc_command.ExecuteNonQuery();
                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        sqlcon.Close();
                    }
                }
                return i_result;
            }
            else if (dbtype == DataBaseType.IBMDB2DLL) //增加DB2单条执行的方式
            {
                try
                {
                    if (DB2DLLcon.State != ConnectionState.Open)
                    {
                        DB2DLLcon.Open();
                    }
                    
                    DB2Command myDB2Command = new DB2Command();
                    myDB2Command.CommandType = System.Data.CommandType.Text;
                    myDB2Command.Connection = DB2DLLcon;
                    myDB2Command.CommandText = s_sql;
                    i_result = myDB2Command.ExecuteNonQuery();

                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    if (!isStanding)
                    {
                        DB2DLLcon.Close();
                    }
                    
                }
                return i_result;
            }
            else if (dbtype == DataBaseType.sqllite) //sqlite单条提交执行
            {
                try
                {
                    if (sqlitecon.State != ConnectionState.Open)
                    {
                        sqlitecon.Open();
                    }

                    SQLiteCommand sqliteCommand = new SQLiteCommand();
                    sqliteCommand.CommandType = System.Data.CommandType.Text;
                    sqliteCommand.Connection = sqlitecon;
                    sqliteCommand.CommandText = s_sql;
                    i_result = sqliteCommand.ExecuteNonQuery();

                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    if (!isStanding)
                    {
                        sqlitecon.Close();
                    }

                }
                return i_result;
            }
            else
            {
                try
                {
                    if (oledbcon.State != ConnectionState.Open)
                    {
                        oledbcon.Open();
                    }
                    OleDbCommand ole_command = new OleDbCommand();
                    ole_command.CommandType = System.Data.CommandType.Text;
                    ole_command.Connection = oledbcon;

                    ole_command.CommandText = s_sql;
                    i_result = ole_command.ExecuteNonQuery();
                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        oledbcon.Close();
                    }
                }
                return i_result;
            }
        }


        /// <summary>
        /// 获取数据集
        /// </summary>
        /// <param name="s_sql">查询sql语句</param>
        /// <returns></returns>
        public DataSet GetData(string s_sql)// 获取数据集
        {
            DataSet ds = new DataSet();
            if (dbtype == DataBaseType.SqlServer)
            {
                try
                {
                    if (sqlcon.State != ConnectionState.Open)
                    {
                        sqlcon.Open();
                    }
                    SqlDataAdapter sda = new SqlDataAdapter(s_sql, sqlcon);
                    sda.Fill(ds, "datatable");
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        sqlcon.Close();
                    }
                }
                return ds;
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                try
                {
                    if (DB2DLLcon.State != ConnectionState.Open)
                    {
                        DB2DLLcon.Open();
                    }
                    
                    DB2DataAdapter DB2DLLDA = new DB2DataAdapter(s_sql, DB2DLLcon);
                    DB2DLLDA.Fill(ds, "datatable");
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    if (!isStanding)
                    {
                        DB2DLLcon.Close();
                    }
                    
                }
                return ds;
            }
            else if (dbtype == DataBaseType.sqllite)//sqlite查询
            {
                try
                {
                    if (sqlitecon.State != ConnectionState.Open)
                    {
                        sqlitecon.Open();
                    }

                    SQLiteDataAdapter sda = new SQLiteDataAdapter(s_sql, sqlitecon);
                    sda.Fill(ds, "datatable");
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    if (!isStanding)
                    {
                        sqlitecon.Close();
                    }

                }
                return ds;
            }
            else
            {
                try
                {
                    if (oledbcon.State != ConnectionState.Open)
                    {
                        oledbcon.Open();
                    }
                    OleDbDataAdapter oda = new OleDbDataAdapter(s_sql, oledbcon);
                    oda.Fill(ds, "datatable");
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        oledbcon.Close();
                    }
                }
                return ds;
            }
        }

        /// <summary>
        /// 获取数据集结构
        /// </summary>
        /// <param name="s_sql">查询sql语句</param>
        /// <returns></returns>
        public DataSet GetDataTableMapped(string s_sql)// 获取数据集结构
        {
            DataSet ds = new DataSet();
            if (dbtype == DataBaseType.SqlServer)
            {
                try
                {
                    if (sqlcon.State != ConnectionState.Open)
                    {
                        sqlcon.Open();
                    }
                    SqlDataAdapter sda = new SqlDataAdapter(s_sql, sqlcon);
                    sda.FillSchema(ds, SchemaType.Mapped);
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        sqlcon.Close();
                    }
                }
                return ds;
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                try
                {
                    if (DB2DLLcon.State != ConnectionState.Open)
                    {
                        DB2DLLcon.Open();
                    }
                    
                    DB2DataAdapter DB2DLLDA = new DB2DataAdapter(s_sql, DB2DLLcon);
                    DB2DLLDA.FillSchema(ds, SchemaType.Mapped);
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    if (!isStanding)
                    {
                        DB2DLLcon.Close();
                    }
                    
                }
                return ds;
            }
            else if (dbtype == DataBaseType.sqllite)
            {
                try
                {
                    if (sqlitecon.State != ConnectionState.Open)
                    {
                        sqlitecon.Open();
                    }

                    SQLiteDataAdapter sda = new SQLiteDataAdapter(s_sql, sqlitecon);
                    sda.FillSchema(ds, SchemaType.Mapped);
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    if (!isStanding)
                    {
                        sqlitecon.Close();
                    }

                }
                return ds;
            }
            else
            {
                try
                {
                    if (oledbcon.State != ConnectionState.Open)
                    {
                        oledbcon.Open();
                    }
                    OleDbDataAdapter oda = new OleDbDataAdapter(s_sql, oledbcon);
                    oda.FillSchema(ds, SchemaType.Mapped);
                }
                catch
                {
                    ds = null;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        oledbcon.Close();
                    }
                }
                return ds;
            }
        }

        /// <summary>
        /// 执行存储过程
        /// </summary>
        /// <param name="s_proname">存储过程名</param>
        /// <param name="o_params">存储过程参数</param>
        /// <returns></returns>
        public int ExcutePro(string s_proname, object[] o_params) //执行存储过程
        {
            int i_result = 0;
            if (dbtype == DataBaseType.SqlServer)
            {
                try
                {
                    if (sqlcon.State != ConnectionState.Open)
                    {
                        sqlcon.Open();
                    }
                    SqlCommand sc_command = new SqlCommand();
                    sc_command.CommandType = System.Data.CommandType.StoredProcedure;
                    sc_command.CommandText = s_proname;
                    sc_command.Connection = sqlcon;
                    foreach (object o_param in o_params)
                    {
                        sc_command.Parameters.Add((SqlParameter)o_param);
                    }
                    i_result = sc_command.ExecuteNonQuery();

                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        sqlcon.Close();
                    }
                }
                return i_result;
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                try
                {
                    DB2DLLcon.Open();
                    DB2Command myDB2Command = new DB2Command();
                    myDB2Command.CommandType = System.Data.CommandType.StoredProcedure;
                    myDB2Command.CommandText = s_proname;
                    myDB2Command.Connection = DB2DLLcon;
                    foreach (object o_param in o_params)
                    {
                        myDB2Command.Parameters.Add((DB2Parameter)o_param);
                    }
                    i_result = myDB2Command.ExecuteNonQuery();

                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    DB2DLLcon.Close();
                }
                return i_result;
            }
            else if (dbtype == DataBaseType.sqllite)
            {
                try
                {
                    sqlitecon.Open();
                    SQLiteCommand sqliteCommand = new SQLiteCommand();
                    sqliteCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqliteCommand.CommandText = s_proname;
                    sqliteCommand.Connection = sqlitecon;
                    foreach (object o_param in o_params)
                    {
                        sqliteCommand.Parameters.Add((SQLiteParameter)o_param);
                    }
                    i_result = sqliteCommand.ExecuteNonQuery();

                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    sqlitecon.Close();
                }
                return i_result;
            }
            else
            {
                try
                {
                    if (oledbcon.State != ConnectionState.Open)
                    {
                        oledbcon.Open();
                    }
                    OleDbCommand ole_command = new OleDbCommand();
                    ole_command.CommandType = System.Data.CommandType.StoredProcedure;
                    ole_command.CommandText = s_proname;
                    foreach (object o_param in o_params)
                    {
                        ole_command.Parameters.Add((OleDbParameter)o_param);
                    }
                    ole_command.Connection = oledbcon;
                    i_result = ole_command.ExecuteNonQuery();
                }
                catch
                {
                    i_result = -1;
                }
                finally
                {
                    //2013-05-31 持续连接的话 连接后不要关闭
                    if (!isStanding)
                    {
                        oledbcon.Close();
                    }
                }
                return i_result;
            }
        }

        /// <summary>
        /// 销毁
        /// </summary>
        /// 使用using后 可能暂时不需要这块内容
        public void Dispose()
        {
            if (dbtype == DataBaseType.SqlServer)
            {
                sqlcon.Close();
                sqlcon.Dispose();
                sqlcon = null;
            }
            else if (dbtype == DataBaseType.IBMDB2DLL)
            {
                DB2DLLcon.Close();
                DB2DLLcon.Dispose();
                DB2DLLcon=null;
            }
            else if (dbtype == DataBaseType.sqllite)
            {
                sqlitecon.Close();
                sqlitecon.Dispose();
                sqlitecon = null;
            }
            else
            {
                oledbcon.Close();
                oledbcon.Dispose();
                oledbcon = null;
            }
        }

        //执行SQL语句文件
        public int[] DoSQLFile(string filename, int packsize)
        {
            List<int> i_l_result = new List<int>();
            string tmp_sql = "";
            List<string> sqls = new List<string>();
            StreamReader sr_reader = null;
            try
            {
                sr_reader = File.OpenText(filename);
                tmp_sql = sr_reader.ReadLine();
                while (tmp_sql != null)
                {
                    if (sqls.Count == packsize)
                    {
                        int i_dosql = this.ExcuteSqls(sqls.ToArray());
                        i_l_result.Add(i_dosql);
                        sqls.Clear();
                    }
                    sqls.Add(tmp_sql);
                    tmp_sql = sr_reader.ReadLine();
                }
                if (sqls.Count > 0)
                {
                    int i_dosql = this.ExcuteSqls(sqls.ToArray());
                    i_l_result.Add(i_dosql);
                }
                return i_l_result.ToArray();
            }
            catch
            {
                return null;
            }
            finally
            {
                if (sr_reader != null)
                {
                    sr_reader.Close();
                }
            }

        }

        //生成导出语句文件
        public int ExportSQLFile(DBLib sDataBase, string sTableName, string[] sColName, DBLib eDataBase, string eTableName, string[] eColName, DataType[] ColType, int transcount, string[] addcol, DataType[] adddt, string[] addvalues, string filename)
        {

            StreamWriter sw = null;
            try
            {
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }
                sw = File.CreateText(filename);
                DBLib sDB = sDataBase;
                DBLib eDB = eDataBase;
                DataSet sDS = null;
                //string[] inisertSql = null;
                string tmpsql = "";
                tmpsql = "select ";
                //int count = 0;
                for (int i = 0; i < sColName.Length; i++)
                {
                    tmpsql += sColName[i] + ",";
                }
                tmpsql = tmpsql.Substring(0, tmpsql.Length - 1) + " from " + sTableName;
                sDS = sDB.GetData(tmpsql);
                if (sDS != null && sDS.Tables[0].Rows.Count > 0)
                {




                    //inisertSql = new string[sDS.Tables[0].Rows.Count];
                    string sql_tmp = "";
                    for (int i = 0; i < sDS.Tables[0].Rows.Count; i++)
                    {
                        sql_tmp = "insert into " + eTableName + "(";
                        for (int j = 0; j < eColName.Length; j++)
                        {
                            sql_tmp += eColName[j] + ",";
                        }

                        if (addcol != null)
                        {
                            for (int j = 0; j < addcol.Length; j++)
                            {
                                sql_tmp += addcol[j] + ",";
                            }
                        }


                        sql_tmp = sql_tmp.Substring(0, sql_tmp.Length - 1) + ") values(";
                        for (int j = 0; j < ColType.Length; j++)
                        {
                            if (ColType[j] == DataType.String)
                            {
                                sql_tmp += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString().Trim() + "',";
                            }
                            else if (ColType[j] == DataType.Int)
                            {
                                sql_tmp += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString().Trim() + "',";
                            }
                            else if (ColType[j] == DataType.Float)
                            {
                                sql_tmp += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString().Trim() + "',";
                            }
                            else if (ColType[j] == DataType.DateTime)
                            {
                                if (eDB.DBType == DataBaseType.Oracle)
                                {
                                    sql_tmp += OracleToDate(sDS.Tables[0].Rows[i].ItemArray[j].ToString().Trim()) + ",";

                                }
                                else
                                {
                                    sql_tmp += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString().Trim() + "',";
                                }


                                //if (eDB.DBType == DataBaseType.SqlServer)
                                //{
                                //    inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                                //}
                                //else if (eDB.DBType == DataBaseType.Oracle)
                                //{
                                //    inisertSql[i] += OracleToDate(sDS.Tables[0].Rows[i].ItemArray[j].ToString()) + ",";
                                //}
                                //else
                                //{
                                //    return -4;
                                //}
                            }
                            else
                            {
                                return -3;
                            }

                        }

                        if (adddt != null)
                        {
                            for (int j = 0; j < adddt.Length; j++)
                            {
                                if (adddt[j] == DataType.String)
                                {
                                    sql_tmp += "'" + addvalues[j] + "',";
                                }
                                else if (adddt[j] == DataType.Int)
                                {
                                    sql_tmp += "'" + addvalues[j] + "',";
                                }
                                else if (adddt[j] == DataType.Float)
                                {
                                    sql_tmp += "'" + addvalues[j] + "',";
                                }
                                else if (adddt[j] == DataType.DateTime)
                                {
                                    if (eDB.DBType == DataBaseType.Oracle)
                                    {
                                        sql_tmp += OracleToDate(addvalues[j]) + ",";

                                    }
                                    else
                                    {
                                        sql_tmp += "'" + addvalues[j] + "',";
                                    }


                                    //if (eDB.DBType == DataBaseType.SqlServer)
                                    //{
                                    //    inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                                    //}
                                    //else if (eDB.DBType == DataBaseType.Oracle)
                                    //{
                                    //    inisertSql[i] += OracleToDate(sDS.Tables[0].Rows[i].ItemArray[j].ToString()) + ",";
                                    //}
                                    //else
                                    //{
                                    //    return -4;
                                    //}
                                }
                                else
                                {
                                    return -3;
                                }

                            }
                        }


                        sql_tmp = sql_tmp.Substring(0, sql_tmp.Length - 1) + ")";
                        sw.WriteLine(sql_tmp);
                    }
                    sw.Close();

                    return 1;
                }
                else
                {
                    return -9;
                }

                //    if (inisertSql.Length > transcount)
                //    {
                //        int g = inisertSql.Length / transcount;
                //        for (int i = 0; i < g; i++)
                //        {
                //            string[] ns = new string[transcount];
                //            for (int j = 0; j < transcount; j++)
                //            {
                //                ns[j] = inisertSql[i * transcount + j];
                //            }
                //            int a = eDB.ExcuteSqls(ns);
                //            if (a > 0)
                //            {
                //                count = count + a;
                //                continue;
                //            }
                //            else
                //            {
                //                return a;
                //            }
                //        }

                //        int b = inisertSql.Length % transcount;
                //        string[] nsy = new string[b];
                //        for (int i = 0; i < b; i++)
                //        {
                //            nsy[i] = inisertSql[g * transcount + i];
                //        }
                //        int c = eDB.ExcuteSqls(nsy);
                //        if (c > 0)
                //        {
                //            count = count + c;
                //        }
                //        else
                //        {
                //            return c;
                //        }
                //        return count;
                //    }
                //    else
                //    {
                //        int a = eDB.ExcuteSqls(inisertSql);
                //        return a;
                //    }

                //}
                //else
                //{
                //    return -2;
                //}


            }
            catch
            {
                sw.Close();
                return -1;

            }
        }


        //表导入操作（未测试）
        public int ExportDB(DBLib sDataBase, string sTableName, string[] sColName, DBLib eDataBase, string eTableName, string[] eColName, DataType[] ColType, int transcount, string[] addcol, DataType[] adddt, string[] addvalues)
        {
            try
            {
                DBLib sDB = sDataBase;
                DBLib eDB = eDataBase;
                DataSet sDS = null;
                string[] inisertSql = null;
                string tmpsql = "";
                tmpsql = "select ";
                int count = 0;
                for (int i = 0; i < sColName.Length; i++)
                {
                    tmpsql += sColName[i] + ",";
                }
                tmpsql = tmpsql.Substring(0, tmpsql.Length - 1) + " from " + sTableName;
                sDS = sDB.GetData(tmpsql);
                if (sDS != null && sDS.Tables[0].Rows.Count > 0)
                {
                    inisertSql = new string[sDS.Tables[0].Rows.Count];

                    for (int i = 0; i < inisertSql.Length; i++)
                    {
                        inisertSql[i] = "insert into " + eTableName + "(";
                        for (int j = 0; j < eColName.Length; j++)
                        {
                            inisertSql[i] += eColName[j] + ",";
                        }

                        for (int j = 0; j < addcol.Length; j++)
                        {
                            inisertSql[i] += addcol[j] + ",";
                        }


                        inisertSql[i] = inisertSql[i].Substring(0, inisertSql[i].Length - 1) + ") values(";
                        for (int j = 0; j < ColType.Length; j++)
                        {
                            if (ColType[j] == DataType.String)
                            {
                                inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                            }
                            else if (ColType[j] == DataType.Int)
                            {
                                inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                            }
                            else if (ColType[j] == DataType.Float)
                            {
                                inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                            }
                            else if (ColType[j] == DataType.DateTime)
                            {
                                if (eDB.DBType == DataBaseType.Oracle)
                                {
                                    inisertSql[i] += OracleToDate(sDS.Tables[0].Rows[i].ItemArray[j].ToString()) + ",";

                                }
                                else
                                {
                                    inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                                }


                                //if (eDB.DBType == DataBaseType.SqlServer)
                                //{
                                //    inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                                //}
                                //else if (eDB.DBType == DataBaseType.Oracle)
                                //{
                                //    inisertSql[i] += OracleToDate(sDS.Tables[0].Rows[i].ItemArray[j].ToString()) + ",";
                                //}
                                //else
                                //{
                                //    return -4;
                                //}
                            }
                            else
                            {
                                return -3;
                            }

                        }

                        for (int j = 0; j < adddt.Length; j++)
                        {
                            if (adddt[j] == DataType.String)
                            {
                                inisertSql[i] += "'" + addvalues[j] + "',";
                            }
                            else if (adddt[j] == DataType.Int)
                            {
                                inisertSql[i] += "'" + addvalues[j] + "',";
                            }
                            else if (adddt[j] == DataType.Float)
                            {
                                inisertSql[i] += "'" + addvalues[j] + "',";
                            }
                            else if (adddt[j] == DataType.DateTime)
                            {
                                if (eDB.DBType == DataBaseType.Oracle)
                                {
                                    inisertSql[i] += OracleToDate(addvalues[j]) + ",";

                                }
                                else
                                {
                                    inisertSql[i] += "'" + addvalues[j] + "',";
                                }


                                //if (eDB.DBType == DataBaseType.SqlServer)
                                //{
                                //    inisertSql[i] += "'" + sDS.Tables[0].Rows[i].ItemArray[j].ToString() + "',";
                                //}
                                //else if (eDB.DBType == DataBaseType.Oracle)
                                //{
                                //    inisertSql[i] += OracleToDate(sDS.Tables[0].Rows[i].ItemArray[j].ToString()) + ",";
                                //}
                                //else
                                //{
                                //    return -4;
                                //}
                            }
                            else
                            {
                                return -3;
                            }

                        }


                        inisertSql[i] = inisertSql[i].Substring(0, inisertSql[i].Length - 1) + ")";
                    }



                    if (inisertSql.Length > transcount)
                    {
                        int g = inisertSql.Length / transcount;
                        for (int i = 0; i < g; i++)
                        {
                            string[] ns = new string[transcount];
                            for (int j = 0; j < transcount; j++)
                            {
                                ns[j] = inisertSql[i * transcount + j];
                            }
                            int a = eDB.ExcuteSqls(ns);
                            if (a > 0)
                            {
                                count = count + a;
                                continue;
                            }
                            else
                            {
                                return a;
                            }
                        }

                        int b = inisertSql.Length % transcount;
                        string[] nsy = new string[b];
                        for (int i = 0; i < b; i++)
                        {
                            nsy[i] = inisertSql[g * transcount + i];
                        }
                        int c = eDB.ExcuteSqls(nsy);
                        if (c > 0)
                        {
                            count = count + c;
                        }
                        else
                        {
                            return c;
                        }
                        return count;
                    }
                    else
                    {
                        int a = eDB.ExcuteSqls(inisertSql);
                        return a;
                    }

                }
                else
                {
                    return -2;
                }

            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// 转换oracle时间格式
        /// </summary>
        /// <param name="datetime"></param>
        /// <returns></returns>
        private string OracleToDate(string datetime)
        {
            return "to_date('" + datetime + "','yyyy-MM-dd hh24:mi:ss')";
        }


        /// <summary>
        /// 获取表名
        /// </summary>
        /// <param name="dl"></param>
        /// <returns></returns>
        public string[] GetTableName()
        {
            try
            {
                List<string> l = new List<string>();
                if (dbtype == DBLib.DataBaseType.Access)
                {
                    oledbcon.Open();
                    DataTable dt = oledbcon.GetSchema("Tables");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        l.Add(dt.Rows[i][2].ToString());
                    }
                    return l.ToArray();
                }
                else if (dbtype == DBLib.DataBaseType.SqlServer)
                {
                    string s_sql = "select name from sysobjects where xtype='U' order by name";
                    DataSet ds = this.GetData(s_sql);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        l.Add(ds.Tables[0].Rows[i][0].ToString());
                    }
                    return l.ToArray();
                }
                else if (dbtype == DBLib.DataBaseType.IBMDB2DLL )
                {
                    string s_sql = "select table_name from user_tables order by table_name";
                    DataSet ds = this.GetData(s_sql);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        l.Add(ds.Tables[0].Rows[i][0].ToString());
                    }
                    return l.ToArray();
                }
                else if (dbtype == DBLib.DataBaseType.sqllite)//获取数据库所有表名
                {
                    string s_sql = "select name from sqlite_master where type='table' order by name";
                    DataSet ds = this.GetData(s_sql);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        l.Add(ds.Tables[0].Rows[i][0].ToString());
                    }
                    return l.ToArray();
                }
                else
                {
                    string s_sql = "select table_name from user_tables order by table_name";
                    DataSet ds = this.GetData(s_sql);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        l.Add(ds.Tables[0].Rows[i][0].ToString());
                    }
                    return l.ToArray();
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                if (oledbcon != null)
                {
                    oledbcon.Close();
                }
                if (sqlcon != null)
                {
                    sqlcon.Close();
                }
            }
        }

        /// <summary>
        /// 获取表中的列名
        /// </summary>
        /// <param name="tablename"></param>
        /// <returns></returns>
        public string[] GetColName(string tablename)
        {
            try
            {
                List<string> colname = new List<string>();
                if (dbtype == DBLib.DataBaseType.Access)//access 获取列名
                {
                    this.OledbBCon.Open();
                    DataTable colTbl = this.OledbBCon.GetSchema("columns", new string[] { null, null, tablename });
                    this.OledbBCon.Close();
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            colname.Add(colTbl.Rows[i]["COLUMN_NAME"].ToString());
                        }
                        return colname.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (dbtype == DBLib.DataBaseType.SqlServer)
                {
                    string s_sql = "select a.name,b.name from syscolumns a,systypes b where a.id=(select max(id) from sysobjects where xtype='u' and name='" + tablename + "')and a.xtype =b. xtype and a.xusertype=b.xusertype  order by a.name";
                    DataTable colTbl = this.GetData(s_sql).Tables[0];
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            colname.Add(colTbl.Rows[i][0].ToString());
                        }
                        return colname.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (dbtype == DBLib.DataBaseType.IBMDB2DLL )//获取表中的所有列明
                {
                    string s_sql = "select COLUMN_NAME,DATA_TYPE from user_tab_columns where table_name ='" + tablename + "' order by column_name";
                    DataTable colTbl = this.GetData(s_sql).Tables[0];
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            colname.Add(colTbl.Rows[i][0].ToString());
                        }
                        return colname.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (dbtype == DBLib.DataBaseType.sqllite)//获取表中的所有列明
                {
                    string s_sql = "pragma table_info([" + tablename + "])";
                    DataTable colTbl = this.GetData(s_sql).Tables[0];
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            colname.Add(colTbl.Rows[i][0].ToString());
                        }
                        return colname.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    string s_sql = "select COLUMN_NAME,DATA_TYPE from user_tab_columns where table_name ='" + tablename + "' order by column_name";
                    DataTable colTbl = this.GetData(s_sql).Tables[0];
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            colname.Add(colTbl.Rows[i][0].ToString());
                        }
                        return colname.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                if (oledbcon != null)
                {
                    oledbcon.Close();
                }
                if (sqlcon != null)
                {
                    sqlcon.Close();
                }
            }
        }

        /// <summary>
        /// 获取列类型
        /// </summary>
        /// <returns></returns>
        public DBLib.DataType[] GetColType(string tablename)
        {
            try
            {
                List<DataType> coltype = new List<DataType>();
                if (dbtype == DBLib.DataBaseType.Access)//access 获取列类型
                {
                    this.OledbBCon.Open();
                    DataTable colTbl = this.OledbBCon.GetSchema("columns", new string[] { null, null, tablename });
                    this.OledbBCon.Close();
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {

                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            string tmptype = colTbl.Rows[i]["DATA_TYPE"].ToString();
                            switch (tmptype)
                            {
                                case "7":
                                    coltype.Add(DBLib.DataType.DateTime);
                                    break;

                                case "3":
                                    coltype.Add(DBLib.DataType.Float);
                                    break;

                                case "130":
                                    coltype.Add(DBLib.DataType.String);
                                    break;

                                default:
                                    coltype.Add(DBLib.DataType.String);
                                    break;
                            }
                        }

                        return coltype.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (dbtype == DBLib.DataBaseType.SqlServer)
                {
                    string s_sql = "select a.name,b.name from syscolumns a,systypes b where a.id=(select max(id) from sysobjects where xtype='u' and name='" + tablename + "')and a.xtype =b. xtype order by a.name";
                    DataTable colTbl = this.GetData(s_sql).Tables[0];
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            string tmptype = colTbl.Rows[i][1].ToString();
                            switch (tmptype)
                            {
                                case "datetime":
                                    coltype.Add(DBLib.DataType.DateTime);
                                    break;

                                case "float":
                                    coltype.Add(DBLib.DataType.Float);
                                    break;

                                case "int":
                                    coltype.Add(DBLib.DataType.Int);
                                    break;

                                default:
                                    coltype.Add(DBLib.DataType.String);
                                    break;
                            }

                        }
                        return coltype.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    string s_sql = "select COLUMN_NAME,DATA_TYPE from user_tab_columns where table_name ='" + tablename + "' order by column_name";
                    DataTable colTbl = this.GetData(s_sql).Tables[0];
                    if (colTbl != null && colTbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < colTbl.Rows.Count; i++)
                        {
                            string tmptype = colTbl.Rows[i][1].ToString();
                            switch (tmptype)
                            {
                                case "DATE":
                                    coltype.Add(DBLib.DataType.DateTime);
                                    break;

                                case "NUMBER":
                                    coltype.Add(DBLib.DataType.Float);
                                    break;

                                case "FLOAT":
                                    coltype.Add(DBLib.DataType.Float);
                                    break;

                                default:
                                    coltype.Add(DBLib.DataType.String);
                                    break;
                            }
                        }
                        return coltype.ToArray();
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                if (oledbcon != null)
                {
                    oledbcon.Close();
                }
                if (sqlcon != null)
                {
                    sqlcon.Close();
                }
            }
        }


        #endregion

        #region 扩展功能 暂时没用 需要测试

        /// <summary>
        /// 根据属性名称获取属性值
        /// </summary>
        /// <param name="SourceObj"></param>
        /// <param name="fieldname"></param>
        /// <returns></returns>
        public object GetPropValue(object SourceObj, string fieldname)
        {
            Type t = SourceObj.GetType();
            IEnumerable<System.Reflection.PropertyInfo> fields = from f in t.GetProperties()
                                                                 where f.Name.ToLower() == fieldname.ToLower()
                                                                 select f;
            object result = null;
            if (fields.Count() > 0)
            {
                result = fields.First().GetValue(SourceObj, null);
            }
            return result;
        }


        /// <summary>
        /// 获取数据集中的列信息
        /// </summary>
        /// <param name="s_sql">查询sql语句</param>
        /// <returns></returns>
        public ColInFo[] GetColInfo(string s_sql)// 获取数据集
        {
            try
            {
                List<ColInFo> lst_ci = new List<ColInFo>();
                DataSet ds = new DataSet();
                string str_fields = s_sql.Substring(s_sql.IndexOf("select ") + 7, s_sql.IndexOf("from") - 7);
                string[] l_fields = str_fields.Split(',');
                for (int i = 0; i < l_fields.Length; i++)
                {
                    l_fields[i] = l_fields[i].Substring(0, l_fields[i].IndexOf(" as"));
                }
                if (dbtype == DataBaseType.SqlServer)
                {
                    try
                    {
                        sqlcon.Open();
                        SqlDataAdapter sda = new SqlDataAdapter(s_sql, sqlcon);
                        sda.FillSchema(ds, SchemaType.Source);
                        foreach (DataColumn dc in ds.Tables[0].Columns)
                        {
                            ColInFo tmp_cf = new ColInFo();
                            tmp_cf.ColName = dc.ColumnName;
                            tmp_cf.ColType = dc.DataType.ToString();
                            tmp_cf.SqlName = l_fields[dc.Table.Columns.IndexOf(dc)];
                            lst_ci.Add(tmp_cf);
                        }
                    }
                    catch
                    {
                        lst_ci = null;
                    }
                    finally
                    {
                        sqlcon.Close();
                    }
                    return lst_ci.ToArray();
                }
                else if (dbtype == DataBaseType.IBMDB2DLL )
                {
                    try
                    {
                        DB2DLLcon.Open();
                        DB2DataAdapter db2dllda = new DB2DataAdapter(s_sql, DB2DLLcon);
                        db2dllda.FillSchema(ds, SchemaType.Source);                        
                        foreach (DataColumn dc in ds.Tables[0].Columns)
                        {
                            ColInFo tmp_cf = new ColInFo();
                            tmp_cf.ColName = dc.ColumnName;
                            tmp_cf.ColType = dc.DataType.ToString();
                            tmp_cf.SqlName = l_fields[dc.Table.Columns.IndexOf(dc)];
                            lst_ci.Add(tmp_cf);
                        }
                    }
                    catch
                    {
                        lst_ci = null;
                    }
                    finally
                    {
                        DB2DLLcon.Close();
                    }
                    return lst_ci.ToArray();
                }
                else if (dbtype == DataBaseType.sqllite)//数据表里面的列的配置获取 暂时用不到
                {
                    try
                    {
                        sqlitecon.Open();
                        SQLiteDataAdapter sda = new SQLiteDataAdapter(s_sql, sqlitecon);
                        sda.FillSchema(ds, SchemaType.Source);//??可能有问题
                        foreach (DataColumn dc in ds.Tables[0].Columns)
                        {
                            ColInFo tmp_cf = new ColInFo();
                            tmp_cf.ColName = dc.ColumnName;
                            tmp_cf.ColType = dc.DataType.ToString();
                            tmp_cf.SqlName = l_fields[dc.Table.Columns.IndexOf(dc)];
                            lst_ci.Add(tmp_cf);
                        }
                    }
                    catch
                    {
                        lst_ci = null;
                    }
                    finally
                    {
                        sqlitecon.Close();
                    }
                    return lst_ci.ToArray();
                }
                else
                {
                    try
                    {
                        oledbcon.Open();
                        OleDbDataAdapter oda = new OleDbDataAdapter(s_sql, oledbcon);
                        oda.FillSchema(ds, SchemaType.Source);
                        foreach (DataColumn dc in ds.Tables[0].Columns)
                        {
                            ColInFo tmp_cf = new ColInFo();
                            tmp_cf.ColName = dc.ColumnName;
                            tmp_cf.ColType = dc.DataType.ToString();
                            tmp_cf.SqlName = l_fields[dc.Table.Columns.IndexOf(dc)];
                            lst_ci.Add(tmp_cf);
                        }
                    }
                    catch
                    {
                        lst_ci = null;
                    }
                    finally
                    {
                        oledbcon.Close();
                    }
                    return lst_ci.ToArray();
                }
            }
            catch
            {
                return null;
            }
        }


        #endregion
    }

    #region 扩展类
    /// <summary>
    /// 数据对象类
    /// </summary>
    public class DataClass
    {
        public static int C_PROPNUM = 50;
        #region 属性（可扩展）
        public string prop00
        {
            set
            {
                keyvalue[0] = value;
            }

            get
            {
                return keyvalue[0];
            }
        }
        public string prop01
        {
            set
            {
                keyvalue[1] = value;
            }
            get
            {
                return keyvalue[1];
            }
        }
        public string prop02
        {
            set
            {
                keyvalue[2] = value;
            }
            get
            {
                return keyvalue[2];
            }
        }
        public string prop03
        {
            set
            {
                keyvalue[3] = value;
            }
            get
            {
                return keyvalue[3];
            }
        }
        public string prop04
        {
            set
            {
                keyvalue[4] = value;
            }
            get
            {
                return keyvalue[4];
            }
        }
        public string prop05
        {
            set
            {
                keyvalue[5] = value;
            }
            get
            {
                return keyvalue[5];
            }
        }
        public string prop06
        {
            set
            {
                keyvalue[6] = value;
            }
            get
            {
                return keyvalue[6];
            }
        }
        public string prop07
        {
            set
            {
                keyvalue[7] = value;
            }
            get
            {
                return keyvalue[7];
            }
        }
        public string prop08
        {
            set
            {
                keyvalue[8] = value;
            }
            get
            {
                return keyvalue[8];
            }
        }
        public string prop09
        {
            set
            {
                keyvalue[9] = value;
            }
            get
            {
                return keyvalue[9];
            }
        }
        public string prop10
        {
            set
            {
                keyvalue[10] = value;
            }
            get
            {
                return keyvalue[10];
            }
        }
        public string prop11
        {
            set
            {
                keyvalue[11] = value;
            }
            get
            {
                return keyvalue[11];
            }
        }
        public string prop12
        {
            set
            {
                keyvalue[12] = value;
            }
            get
            {
                return keyvalue[12];
            }
        }
        public string prop13
        {
            set
            {
                keyvalue[13] = value;
            }
            get
            {
                return keyvalue[13];
            }
        }
        public string prop14
        {
            set
            {
                keyvalue[14] = value;
            }
            get
            {
                return keyvalue[14];
            }
        }
        public string prop15
        {
            set
            {
                keyvalue[15] = value;
            }
            get
            {
                return keyvalue[15];
            }
        }
        public string prop16
        {
            set
            {
                keyvalue[16] = value;
            }
            get
            {
                return keyvalue[16];
            }
        }
        public string prop17
        {
            set
            {
                keyvalue[17] = value;
            }
            get
            {
                return keyvalue[17];
            }
        }
        public string prop18
        {
            set
            {
                keyvalue[18] = value;
            }
            get
            {
                return keyvalue[18];
            }
        }
        public string prop19
        {
            set
            {
                keyvalue[19] = value;
            }
            get
            {
                return keyvalue[19];
            }
        }
        public string prop20
        {
            set
            {
                keyvalue[20] = value;
            }
            get
            {
                return keyvalue[20];
            }
        }
        public string prop21
        {
            set
            {
                keyvalue[21] = value;
            }
            get
            {
                return keyvalue[21];
            }
        }
        public string prop22
        {
            set
            {
                keyvalue[22] = value;
            }
            get
            {
                return keyvalue[22];
            }
        }
        public string prop23
        {
            set
            {
                keyvalue[23] = value;
            }
            get
            {
                return keyvalue[23];
            }
        }
        public string prop24
        {
            set
            {
                keyvalue[24] = value;
            }
            get
            {
                return keyvalue[24];
            }
        }
        public string prop25
        {
            set
            {
                keyvalue[25] = value;
            }
            get
            {
                return keyvalue[25];
            }
        }
        public string prop26
        {
            set
            {
                keyvalue[26] = value;
            }
            get
            {
                return keyvalue[26];
            }
        }
        public string prop27
        {
            set
            {
                keyvalue[27] = value;
            }
            get
            {
                return keyvalue[27];
            }
        }
        public string prop28
        {
            set
            {
                keyvalue[28] = value;
            }
            get
            {
                return keyvalue[28];
            }
        }
        public string prop29
        {
            set
            {
                keyvalue[29] = value;
            }
            get
            {
                return keyvalue[29];
            }
        }
        public string prop30
        {
            set
            {
                keyvalue[30] = value;
            }
            get
            {
                return keyvalue[30];
            }
        }
        public string prop31
        {
            set
            {
                keyvalue[31] = value;
            }
            get
            {
                return keyvalue[31];
            }
        }
        public string prop32
        {
            set
            {
                keyvalue[32] = value;
            }
            get
            {
                return keyvalue[32];
            }
        }
        public string prop33
        {
            set
            {
                keyvalue[33] = value;
            }
            get
            {
                return keyvalue[33];
            }
        }
        public string prop34
        {
            set
            {
                keyvalue[34] = value;
            }
            get
            {
                return keyvalue[34];
            }
        }
        public string prop35
        {
            set
            {
                keyvalue[35] = value;
            }
            get
            {
                return keyvalue[35];
            }
        }
        public string prop36
        {
            set
            {
                keyvalue[36] = value;
            }
            get
            {
                return keyvalue[36];
            }
        }
        public string prop37
        {
            set
            {
                keyvalue[37] = value;
            }
            get
            {
                return keyvalue[37];
            }
        }
        public string prop38
        {
            set
            {
                keyvalue[38] = value;
            }
            get
            {
                return keyvalue[38];
            }
        }
        public string prop39
        {
            set
            {
                keyvalue[39] = value;
            }
            get
            {
                return keyvalue[39];
            }
        }
        public string prop40
        {
            set
            {
                keyvalue[40] = value;
            }
            get
            {
                return keyvalue[40];
            }
        }
        public string prop41
        {
            set
            {
                keyvalue[41] = value;
            }
            get
            {
                return keyvalue[41];
            }
        }
        public string prop42
        {
            set
            {
                keyvalue[42] = value;
            }
            get
            {
                return keyvalue[42];
            }
        }
        public string prop43
        {
            set
            {
                keyvalue[43] = value;
            }
            get
            {
                return keyvalue[43];
            }
        }
        public string prop44
        {
            set
            {
                keyvalue[44] = value;
            }
            get
            {
                return keyvalue[44];
            }
        }
        public string prop45
        {
            set
            {
                keyvalue[45] = value;
            }
            get
            {
                return keyvalue[45];
            }
        }
        public string prop46
        {
            set
            {
                keyvalue[46] = value;
            }
            get
            {
                return keyvalue[46];
            }
        }
        public string prop47
        {
            set
            {
                keyvalue[47] = value;
            }
            get
            {
                return keyvalue[47];
            }
        }
        public string prop48
        {
            set
            {
                keyvalue[48] = value;
            }
            get
            {
                return keyvalue[48];
            }
        }
        public string prop49
        {
            set
            {
                keyvalue[49] = value;
            }
            get
            {
                return keyvalue[49];
            }
        }
        #endregion

        public string[] propkeys = new string[C_PROPNUM];//属性名称索引
        public string[] keyvalue = new string[C_PROPNUM];//属性值数组

        /// <summary>
        /// 构造函数 初始化类属性的索引
        /// </summary>
        public DataClass()
        {
            for (int i = 0; i < propkeys.Length; i++)
            {
                propkeys[i] = "prop" + i.ToString().PadLeft(2, '0');
            }
        }
    }

    /// <summary>
    /// 列信息类
    /// </summary>
    public class ColInFo
    {
        public string ColName { get; set; }
        public string ColType { get; set; }
        public string SqlName { get; set; }

    }

    #endregion
}
