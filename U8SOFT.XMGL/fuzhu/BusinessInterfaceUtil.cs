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
//        /// ִ��U8ҵ��API
//        /// </summary>
//        /// <param name="login">U8��¼����</param>
//        /// <param name="regUnitName">ע�ᵥԪ���ƣ���ֵ������ U8��װĿ¼\UAP\BusinessInterface\Framework\UFSoft.U8.Business.Interface.config�ļ��е�APIRepository�ڵ���ӽڵ�Product �ڵ��name����ֵ</param>
//        /// <param name="adapterName">���������� ����ע�ᵥԪ�����ҵ�һ����ӦAPI�����ļ��Ĵ��·�����������ļ��е�Adapter�ڵ��Name���Ի��ҵ���Ӧֵ</param>
//        /// <param name="apiName">ҵ��API���� API�����ļ��и��ݾ���ҵ�����Ҫ�ҵ�API �ڵ��name���Ե�ֵ</param>
//        /// <param name="version">ҵ��API�汾�� ��ֵ������ U8��װĿ¼\UAP\BusinessInterface\Framework\UFSoft.U8.Business.Interface.config�ļ��е�APIRepository�ڵ���ӽڵ�Product �ڵ��version����ֵ</param>
//        /// <param name="parameters">ҵ��API�����Ĳ�������,Hashtable��key��ʾ�����������֣�value��ʾ��Ӧ������ֵ</param>
//        /// <returns>ҵ��API�����ķ���ֵ</returns>
//        public static object ExcutingAPI(LoginInfo login, string regUnitName, string adapterName, string apiName, string version, Hashtable parameters)
//        {
//            if (regUnitName == null || regUnitName == ""
//                || adapterName == null || adapterName == ""
//                || apiName == null || apiName == ""
//                || version == null || version == "")
//                throw new ArgumentNullException("���뷽��GetExcutingAPI�����п�ֵ");

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
//                throw new Exception(string.Format("���� regUnitName is {0}, adapterName is {1}, apiName is {2}, version is {3} ����. \n������Ϣ��{4}", regUnitName, adapterName, apiName, version, ex.Message), ex);
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
//        //        throw new Exception("���ݼ����Ѿ��������� " + editPropColumn + " �޷���DataSet�еı� " + datatableName +" ת��ΪDOMDocument" );
//        //    }

//        //    if (currentTable.Columns.Contains(UAPRuntime_RowState))
//        //    {
//        //        for (int i = 0; i < currentTable.Rows.Count; i++)
//        //        {
//        //            DataRow oneRow = currentTable.Rows[i];
//        //            string rowState = oneRow[UAPRuntime_RowState].ToString();
//        //            switch (rowState)
//        //            {
//        //                //δ����
//        //                case "0":
//        //                    oneRow[editPropColumn] = "";
//        //                    break;
//        //                //����
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
//        //                    //δ����
//        //                    case "":
//        //                        oneRow[UAPRuntime_RowState] = 0;
//        //                        break;
//        //                    //����
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
