using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UFIDA.U8.UAP.UI.Runtime.Model;
using UFIDA.U8.UAP.UI.Runtime.Common;
using System.Windows.Forms;
using System.Data;
using System.Windows;
//using SendMsg;
//using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
//using U8SOFT.XMRZ;

namespace U8SOFT.XMRZ
{
    class CancelVoucherButton : IButtonEventHandler
    {

        #region IButtonEventHandler 成员

        public string Excute(VoucherProxy ReceiptObject, string PreExcuteResult)
        {
            //using (SqlConnection conn = new SqlConnection(ReceiptObject.LoginInfo.UFDataSqlConStr))
            //{
            //    conn.Open();
            //    SqlTransaction tran = conn.BeginTransaction(System.Data.IsolationLevel.RepeatableRead);
            //    try
            //    {
            //        #region  审核之后不能作废(检查)
            //        string queryAuditSql = string.Format("select cAuditor from LK_XM_LX where LK1_0007_E001_PK = {0}", ReceiptObject.CurrentPKValue);
            //        SqlCommand queryAuditerCommand = new SqlCommand(queryAuditSql, conn, tran);
            //        object Auditer = queryAuditerCommand.ExecuteScalar();
            //        if (Auditer != System.DBNull.Value
            //            && Auditer.ToString() != "")
            //            throw new Exception("当前单据已被审核。");
            //        #endregion
            //        string queryIsCancelSql = string.Format("select cCanceler from LK_XM_LX where LK1_0007_E001_PK = {0}", ReceiptObject.CurrentPKValue);
            //        SqlCommand queryCommand = new SqlCommand(queryIsCancelSql, conn, tran);
            //        object cIsCancel = queryCommand.ExecuteScalar();
            //        if (cIsCancel != System.DBNull.Value
            //            && cIsCancel.ToString() != "")
            //            throw new Exception("当前单据已被作废。");
            //        string currentTime = DateTime.Now.ToString();
            //        string updateIsCancelSql = string.Format("update LK_XM_LX set cCanceler='{0}',cCancelTime='{1}' where LK1_0007_E001_PK='{2}'"
            //            , ReceiptObject.LoginInfo.UserName
            //            , currentTime
            //            , ReceiptObject.CurrentPKValue);
            //        SqlCommand updateCommand = new SqlCommand(updateIsCancelSql, conn, tran);
            //        int resultCount = updateCommand.ExecuteNonQuery();
            //        tran.Commit();
            //        ReceiptObject.Businesses["LK1_0007_E001"].Cells["cCanceler"].Value = ReceiptObject.LoginInfo.UserName;
            //        ReceiptObject.Businesses["LK1_0007_E001"].Cells["cCancelTime"].Value = currentTime;
            //        MessageBox.Show("成功作废");
            //    }
            //    catch (Exception ex)
            //    {
            //        tran.Rollback();
            //        tran.Dispose();
            //        MessageBox.Show(ex.Message);
            //        return MakeExcuteState(false, ex.Message);
            //    }
            //    finally
            //    {
            //        if (conn != null && conn.State != System.Data.ConnectionState.Closed)
            //            conn.Close();
            //    }
            //}
            //return this.MakeExcuteState(true, "");

            return null;



           
            //throw new NotImplementedException();
        }

        public string Excuted(VoucherProxy ReceiptObject, string PreExcuteResult)
        {
            return null;
        }

        public string Excuting(VoucherProxy ReceiptObject)
        {
            return null;
        }

        #endregion
   

        private byte[] GetImageBytes(Image image)
        {
            MemoryStream mstream = new MemoryStream();
            image.Save(mstream, System.Drawing.Imaging.ImageFormat.Jpeg);
            byte[] byteData = new Byte[mstream.Length];
            mstream.Position = 0;
            mstream.Read(byteData, 0, byteData.Length);
            mstream.Close();
            return byteData;
        }

        /// <summary>
        /// 创建一个系统函数的执行结
        /// </summary>
        /// <param name="success">执行成功与否标志</param>
        /// <param name="errinfo">如果执行错误，该参数为其其错误描述信息</param>
        /// <returns>执行结果的xml描述</returns>
        private string MakeExcuteState(bool success, string errinfo)
        {
            if (success == true)
                return "<result><system result=\"true\"/></result>";
            else
                return "<result><system result=\"false\" errinfo=\"" + errinfo + "\"/></result>";
        }
        

    }
}
    

