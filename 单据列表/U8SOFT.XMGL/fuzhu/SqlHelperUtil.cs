using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Data.SqlClient;
using System.Data;
using UFIDA.U8.UAP.Common;

namespace UFIDA.U8.UAP.Plugin.SalesVoucher
{
    public class SqlHelperUtil
    {
        public static void SqlHelpQueryOneLine(string connectionString, CommandType commandType, string commandText, IDictionary<string, object> result)
        {
            SqlDataReader reader = SqlHelper.ExecuteReader(connectionString, commandType, commandText);
            try
            {
                if (reader.Read())
                {
                    if (reader.FieldCount != result.Count)
                    {
                        throw new Exception("查询结果列数目与需要获取的结果列数目不一致!");
                    }
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        result[reader.GetName(i)] = reader.GetValue(i);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (reader != null && !reader.IsClosed)
                    reader.Close();
            }
        }

        /// <summary>
        /// 根据提供的键值对从数据库中取出内容更新内存中的数据表
        /// </summary>
        /// <param name="table">需要更新的数据表</param>
        /// <param name="connStr">数据库链接串</param>
        /// <param name="Keys">键值对</param>
        public static void UpdateDataTableByKey(DataTable table, string connStr, Hashtable Keys)
        {
            StringBuilder selectPart = new StringBuilder("SELECT ");
            for (int i = 0; i < table.Columns.Count; i++)
            {
                if (i != table.Columns.Count - 1)
                {
                    selectPart.Append(table.Columns[i].ColumnName).Append(",");
                }
                else
                {
                    selectPart.Append(table.Columns[i].ColumnName);
                }
            }

            string fromPart = " FROM " + table.TableName;

            StringBuilder wherePart = new StringBuilder(" WHERE ");
            int index = 0;
            foreach (object key in Keys.Keys)
            {
                index++;
                if (index != Keys.Count)
                {
                    wherePart.Append(key.ToString()).Append(" = ").Append(Keys[key].ToString()).Append(",");
                }
                else
                {
                    wherePart.Append(key.ToString()).Append(" = ").Append(Keys[key].ToString());
                }
            }

            string sqlScript = selectPart.Append(fromPart).Append(wherePart.ToString()).ToString();

            DataSet result = SqlHelper.ExecuteDataset(connStr, CommandType.Text, sqlScript);

            if (result.Tables.Contains(table.TableName))
            {
                table = result.Tables[table.TableName].Copy();
            }
        }

        #region 注释
        //public static void UpdateDataTableByForeignKey(DataTable table, string connStr, Hashtable foreignKeys)
        //{
        //    StringBuilder selectPart = new StringBuilder("SELECT ");
        //    for (int i = 0; i < table.Columns.Count; i++)
        //    {
        //        if (i != table.Columns.Count - 1)
        //        {
        //            selectPart.Append(table.Columns[i].ColumnName).Append(",");
        //        }
        //        else
        //        {
        //            selectPart.Append(table.Columns[i].ColumnName);
        //        }
        //    }

        //    string fromPart = " FROM " + table.TableName;

        //    StringBuilder wherePart = new StringBuilder(" WHERE ");
        //    int index = 0;
        //    foreach (object key in foreignKeys.Keys)
        //    {
        //        index++;
        //        if (index != foreignKeys.Count)
        //        {
        //            wherePart.Append(key.ToString()).Append(" = ").Append(foreignKeys[key].ToString()).Append(",");
        //        }
        //        else
        //        {
        //            wherePart.Append(key.ToString()).Append(" = ").Append(foreignKeys[key].ToString());
        //        }
        //    }

        //    string sqlScript = selectPart.Append(fromPart).Append(wherePart.ToString()).ToString();

        //    DataSet result = SqlHelper.ExecuteDataset(connStr, CommandType.Text, sqlScript);

        //    if (result.Tables.Contains(table.TableName))
        //    {
        //        table = result.Tables[table.TableName].Copy();
        //    }
        //}
        #endregion
    }
}
