using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SQLite;
using System.Data;

namespace DB.SQLITE.SQLiteBulkInsert
{
    public class SQLiteBulkInsert
    {
        private SQLiteConnection m_dbCon;
        private SQLiteCommand m_cmd;
        private SQLiteTransaction m_trans;

        private Dictionary<string ,SQLiteParameter> m_parameters = new Dictionary<string, SQLiteParameter>();//？？

        private uint m_counter = 0;

        private string m_beginInsertText;

        public readonly Dictionary<Type, DbType> typeMap = new Dictionary<Type, DbType>();

        
        public void InitialTypetoDbType()
        {
            InitialTypetoDbTypeDict();
        }

        //初始化 类型翻译字典
        private void InitialTypetoDbTypeDict()
        {
            typeMap[typeof(byte)] = DbType.Byte;
            typeMap[typeof(sbyte)] = DbType.SByte;
            typeMap[typeof(short)] = DbType.Int16;
            typeMap[typeof(ushort)] = DbType.UInt16;
            typeMap[typeof(int)] = DbType.Int32;
            typeMap[typeof(uint)] = DbType.UInt32;
            typeMap[typeof(long)] = DbType.Int64;
            typeMap[typeof(ulong)] = DbType.UInt64;
            typeMap[typeof(float)] = DbType.Single;
            typeMap[typeof(double)] = DbType.Double;
            typeMap[typeof(decimal)] = DbType.Decimal;
            typeMap[typeof(bool)] = DbType.Boolean;
            typeMap[typeof(string)] = DbType.String;
            typeMap[typeof(char)] = DbType.StringFixedLength;
            typeMap[typeof(Guid)] = DbType.Guid;
            typeMap[typeof(DateTime)] = DbType.DateTime;
            typeMap[typeof(DateTimeOffset)] = DbType.DateTimeOffset;
            typeMap[typeof(byte[])] = DbType.Binary;
            typeMap[typeof(byte?)] = DbType.Byte;
            typeMap[typeof(sbyte?)] = DbType.SByte;
            typeMap[typeof(short?)] = DbType.Int16;
            typeMap[typeof(ushort?)] = DbType.UInt16;
            typeMap[typeof(int?)] = DbType.Int32;
            typeMap[typeof(uint?)] = DbType.UInt32;
            typeMap[typeof(long?)] = DbType.Int64;
            typeMap[typeof(ulong?)] = DbType.UInt64;
            typeMap[typeof(float?)] = DbType.Single;
            typeMap[typeof(double?)] = DbType.Double;
            typeMap[typeof(decimal?)] = DbType.Decimal;
            typeMap[typeof(bool?)] = DbType.Boolean;
            typeMap[typeof(char?)] = DbType.StringFixedLength;
            typeMap[typeof(Guid?)] = DbType.Guid;
            typeMap[typeof(DateTime?)] = DbType.DateTime;
            typeMap[typeof(DateTimeOffset?)] = DbType.DateTimeOffset;
            typeMap[typeof(System.Data.Linq.Binary)] = DbType.Binary;
        }

        public SQLiteBulkInsert(SQLiteConnection dbConnection, string tableName)
        {
            m_dbCon = dbConnection;
            m_tableName = tableName;

            StringBuilder query = new StringBuilder(255);
            query.Append("INSERT INTO ["); query.Append(tableName); query.Append("] (");
            m_beginInsertText = query.ToString();
        }

        private bool m_allowBulkInsert = true;
        public bool AllowBulkInsert { get { return m_allowBulkInsert; } set { m_allowBulkInsert = value; } }

        public string CommandText
        {
            get
            {
                if (m_parameters.Count < 1)
                    throw new SQLiteException("You must add at least one parameter.");

                StringBuilder sb = new StringBuilder(255);
                sb.Append(m_beginInsertText);

                foreach (string param in m_parameters.Keys)
                {
                    sb.Append('[');
                    sb.Append(param);
                    sb.Append(']');
                    sb.Append(", ");
                }
                sb.Remove(sb.Length - 2, 2);

                sb.Append(") VALUES (");

                foreach (string param in m_parameters.Keys)
                {
                    sb.Append(m_paramDelim);
                    sb.Append(param);
                    sb.Append(", ");
                }
                sb.Remove(sb.Length - 2, 2);

                sb.Append(")");

                return sb.ToString();
            }
        }

        private uint m_commitMax = 10000;
        public uint CommitMax { get { return m_commitMax; } set { m_commitMax = value; } }

        private string m_tableName;
        public string TableName { get { return m_tableName; } }

        private string m_paramDelim = ":";
        public string ParamDelimiter { get { return m_paramDelim; } }

        public void AddParameter(string name, DbType dbType)
        {
            SQLiteParameter param = new SQLiteParameter(m_paramDelim + name, dbType);
            m_parameters.Add(name, param);
        }

        public void Flush()
        {
            try
            {
                if (m_trans != null)
                    m_trans.Commit();
            }
            catch (Exception ex) { throw new Exception("Could not commit transaction. See InnerException for more details", ex); }
            finally
            {
                if (m_trans != null)
                    m_trans.Dispose();

                m_trans = null;
                m_counter = 0;
            }
        }

        public void Insert(object[] paramValues)
        {
            try
            {
                if (paramValues.Length != m_parameters.Count)
                    throw new Exception("The values array count must be equal to the count of the number of parameters.");

                m_counter++;

                if (m_counter == 1)
                {
                    if (m_allowBulkInsert)
                        m_trans = m_dbCon.BeginTransaction();

                    m_cmd = m_dbCon.CreateCommand();
                    foreach (SQLiteParameter par in m_parameters.Values)
                        m_cmd.Parameters.Add(par);

                    m_cmd.CommandText = this.CommandText;
                }

                int i = 0;
                foreach (SQLiteParameter par in m_parameters.Values)
                {
                    par.Value = paramValues[i];
                    i++;
                }
                //执行出错 会导致连接卡死（可能是sqlite的连接独占原因）
                m_cmd.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                if (m_trans != null)
                    m_trans.Dispose();

                m_trans = null;
                m_counter = 0;
                throw new Exception("Could not do ExecuteNonQuery. See InnerException for more details:"+ex.Message);
            }

            if (m_counter == m_commitMax)
            {
                try
                {
                    if (m_trans != null)
                        m_trans.Commit();
                }
                catch (Exception ex) { }
                finally
                {
                    if (m_trans != null)
                    {
                        m_trans.Dispose();
                        m_trans = null;
                    }

                    m_counter = 0;
                }
            }
        }
    }
}