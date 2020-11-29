//using System;
//using System.Collections.Generic;
//using System.Text;
//using System.Collections;
//using System.Data;
//using UFSoft.U8.Business.Interface;
//using UFSoft.U8.AppSvr.DAC.DODataTypeConvert;
//using U8Login;
//using UFIDA.U8.UAP.Common;
//using System.Xml;
//using System.IO;
//using System.Windows.Forms;
//using System.Reflection;

//namespace UFIDA.U8.UAP.Plugin.SalesVoucher
//{
//    public class BusinessInterfaceUtil
//    {
//        /// <summary>
//        /// 执行U8业务API
//        /// </summary>
//        /// <param name="login">U8登录对象</param>
//        /// <param name="regUnitName">注册单元名称，该值可以在 U8安装目录\UAP\BusinessInterface\Framework\UFSoft.U8.Business.Interface.config文件中的APIRepository节点的子节点Product 节点的name属性值</param>
//        /// <param name="adapterName">适配器名称 根据注册单元名称找到一个对应API配置文件的存放路径，在配置文件中的Adapter节点的Name属性会找到相应值</param>
//        /// <param name="apiName">业务API名称 API配置文件中根据具体业务的需要找到API 节点的name属性的值</param>
//        /// <param name="version">业务API版本号 该值可以在 U8安装目录\UAP\BusinessInterface\Framework\UFSoft.U8.Business.Interface.config文件中的APIRepository节点的子节点Product 节点的version属性值</param>
//        /// <param name="parameters">业务API方法的参数集合,Hashtable中key表示方法参数名字，value表示对应参数的值</param>
//        /// <returns>业务API方法的返回值</returns>
//        public static object ExcutingAPI(LoginInfo login, string regUnitName, string adapterName, string apiName, string version, Hashtable parameters)
//        {
//            if (regUnitName == null || regUnitName == ""
//                || adapterName == null || adapterName == ""
//                || apiName == null || apiName == ""
//                || version == null || version == "")
//                throw new ArgumentNullException("传入方法GetExcutingAPI参数有空值");

//            if (parameters == null)
//                throw new ArgumentNullException("parameters");

//            UFSoft.U8.Business.Interface.BizAdapterService service = new UFSoft.U8.Business.Interface.BizAdapterService();
//            //if (Directory.GetCurrentDirectory() != Application.StartupPath)
//            //    Directory.SetCurrentDirectory(Application.StartupPath);
//            service.ConfigFile = login.AppPath + @"\BusinessInterface\Framework\UFSoft.U8.Business.Interface.config";
//            IBizAPI resultAPI = service.GetBizAPI2(regUnitName, adapterName, apiName, version);
//            foreach (object key in parameters.Keys)
//            {
//                resultAPI.Parameters[key.ToString()].SetValue(parameters[key]);
//            }
//            ADODB.Connection conn = new ADODB.ConnectionClass();
//            try
//            {
//                conn.Open(login.UFDataOleDbConStr, login.DbUser, login.DbPwd, 0);
//                object u8login = login.U8Login as object;
//                object adoconn = (object)conn;
//                object retObj = resultAPI.Execute(u8login, adoconn, true);
//                object[] keys = new object[parameters.Keys.Count];
//                parameters.Keys.CopyTo(keys, 0);
//                for (int i = 0; i < keys.Length; i++)
//                {
//                    string key = keys[i].ToString();
//                    parameters[key] = resultAPI.Parameters[key];
//                }
//                return retObj;
//            }
//            catch (Exception ex)
//            {
//                throw new Exception(string.Format("调用 regUnitName is {0}, adapterName is {1}, apiName is {2}, version is {3} 出错. \n错误信息：{4}", regUnitName, adapterName, apiName, version, ex.Message), ex);
//            }
//            finally
//            {
//                if (conn != null)
//                {
//                    try
//                    {
//                        conn.Close();
//                    }
//                    catch
//                    {
//                    }
//                }
//            }
//        }

//        public static MSXML2.DOMDocument ConvertDataSetToDOM(DataSet dataset, string datatableName)
//        {
//            MSXML2.DOMDocument resultDocument = new MSXML2.DOMDocumentClass();
//            ConvertDataSet_Interface converter = new ConvertDataSet();
//            DOCollection recColl = converter.ToRecordset(dataset);
//            ADODB.Recordset rescordSet = recColl.Item(datatableName) as ADODB.Recordset;
//            rescordSet.Save((object)resultDocument, ADODB.PersistFormatEnum.adPersistXML);
//            return resultDocument;
//        }

//        public static void ConvertDOMToDataSet(ref DataSet dataset, string datatableName, MSXML2.DOMDocument dom)
//        {
//            ConvertRecordset_Interface converter = new ConvertRecordset();
//            ADODB.Recordset recordset = new ADODB.RecordsetClass();
//            recordset.Open(dom, Type.Missing, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, -1);
//            DOCollection colletion = new DOCollection();
//            colletion.Add(recordset, datatableName);
//            DataSet tempDS = converter.ToDataSet(colletion);
//            DataTable tempDT = new DataTable();
//            string tablename = tempDS.Tables[0].TableName;
//            tempDT = tempDS.Tables[datatableName].Copy();
//            dataset.Merge(tempDT);
//        }

//        //private static string editPropColumn = "editprop";

//        //private static string UAPRuntime_RowState = "UAPRuntime_RowState";

//        //public static MSXML2.DOMDocument ConvertDataSetToDOM(DataSet dataset, string datatableName)
//        //{
//        //    DataTable currentTable = dataset.Tables[datatableName];
//        //    if (!currentTable.Columns.Contains(editPropColumn))
//        //    {
//        //        DataColumn editPropCol = new DataColumn(editPropColumn, Type.GetType("System.String"));
//        //        currentTable.Columns.Add(editPropCol);
//        //    }
//        //    else
//        //    {
//        //        throw new Exception("数据集中已经包含了列 " + editPropColumn + " 无法将DataSet中的表 " + datatableName +" 转化为DOMDocument" );
//        //    }

//        //    if (currentTable.Columns.Contains(UAPRuntime_RowState))
//        //    {
//        //        for (int i = 0; i < currentTable.Rows.Count; i++)
//        //        {
//        //            DataRow oneRow = currentTable.Rows[i];
//        //            string rowState = oneRow[UAPRuntime_RowState].ToString();
//        //            switch (rowState)
//        //            {
//        //                //未更改
//        //                case "0":
//        //                    oneRow[editPropColumn] = "";
//        //                    break;
//        //                //新增
//        //                case "1":
//        //                    oneRow[editPropColumn] = "A";
//        //                    break;
//        //                case "2":
//        //                    oneRow[editPropColumn] = "D";
//        //                    break;
//        //                case "3":
//        //                    oneRow[editPropColumn] = "M";
//        //                    break;
//        //                default:
//        //                    break;
//        //            }
//        //        }
//        //    }
//        //    MSXML2.DOMDocument resultDocument = new MSXML2.DOMDocumentClass();
//        //    ConvertDataSet_Interface converter = new ConvertDataSet();
//        //    DOCollection recColl = converter.ToRecordset(dataset);
//        //    ADODB.Recordset rescordSet = recColl.Item(datatableName) as ADODB.Recordset;
//        //    rescordSet.Save((object)resultDocument, ADODB.PersistFormatEnum.adPersistXML);
//        //    return resultDocument;
//        //}

//        //public static void ConvertDOMToDataSet(ref DataSet dataset, string datatableName, MSXML2.DOMDocument dom)
//        //{
//        //    ConvertRecordset_Interface converter = new ConvertRecordset();
//        //    ADODB.Recordset recordset = new ADODB.RecordsetClass();
//        //    recordset.Open(dom, Type.Missing, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, -1);
//        //    DOCollection colletion = new DOCollection();
//        //    colletion.Add(recordset, datatableName);
//        //    DataSet tempDS = converter.ToDataSet(colletion);
//        //    DataTable tempDT = new DataTable();
//        //    string tablename = tempDS.Tables[0].TableName;
//        //    tempDT = tempDS.Tables[datatableName].Copy();

//        //    if (tempDT.Columns.Contains(editPropColumn))
//        //    {
//        //        if (tempDT.Columns.Contains(UAPRuntime_RowState))
//        //        {
//        //            for (int i = 0; i < tempDT.Rows.Count; i++)
//        //            {
//        //                DataRow oneRow = tempDT.Rows[i];
//        //                string rowState = oneRow[editPropColumn].ToString();
//        //                switch (rowState)
//        //                {
//        //                    //未更改
//        //                    case "":
//        //                        oneRow[UAPRuntime_RowState] = 0;
//        //                        break;
//        //                    //新增
//        //                    case "A":
//        //                        oneRow[UAPRuntime_RowState] = 1;
//        //                        break;
//        //                    case "D":
//        //                        oneRow[UAPRuntime_RowState] = 2;
//        //                        break;
//        //                    case "M":
//        //                        oneRow[UAPRuntime_RowState] = 3;
//        //                        break;
//        //                    default:
//        //                        break;
//        //                }
//        //            }
//        //        }
//        //        tempDT.Columns.Remove(editPropColumn);
//        //    }

//        //    dataset.Merge(tempDT);
//        //}
//    }
//}
