using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UFIDA.U8.UAP.UI.Runtime.Model;
using UFIDA.U8.UAP.UI.Runtime.Common;
using System.Windows.Forms;
using System.Data;
//using SendMsg;
//using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using fuzhu;
//using U8SOFT.XMRZ;

namespace U8SOFT.XMRZ
{
    class SaveVoucherButton : IButtonEventHandler
    {

        #region IButtonEventHandler 成员

        public string Excute(VoucherProxy ReceiptObject, string PreExcuteResult)
        {

            return null;



           
            //throw new NotImplementedException();
        }
        //保存后执行
        public string Excuted(VoucherProxy ReceiptObject, string PreExcuteResult)
        {


            //DataSet ds = ReceiptObject.GetData(false, false);
            //Business dt = ReceiptObject.Businesses["LK1_0007_E002"];
            //string cNo = DbHelper.GetDbString(dt.Rows[0].Cells["cNo"].Value);
            //DbHelper.conStr = ReceiptObject.LoginInfo.UFDataSqlConStr;

            ////不为空，保存，否则不保存
            //if (xmrz.currentControl.Image != null)
            //{
            //    byte[] imageBytes = DbHelper.GetImageBytes(xmrz.currentControl.Image);

            //    SqlTransaction tran = DbHelper.BeginTrans();
            //    try
            //    {
            //        string sql = "delete from LK1_0003_PIC where cno = @cno";
            //        DbHelper.ExecuteSqlWithTrans(sql, new SqlParameter[]{
            //                    new SqlParameter("@cno",cNo)}, tran);


            //        sql = "insert into LK1_0003_PIC(cno,pic) values(@cno,@pic)";
            //        DbHelper.ExecuteSqlWithTrans(sql, new SqlParameter[]{
            //                    new SqlParameter("@cno",cNo),
            //                    new SqlParameter("@pic",imageBytes)}, tran);

            //        DbHelper.CommitTransAndCloseConnection(tran);

            //        return null;
            //    }
            //    catch (Exception ex)
            //    {
            //        DbHelper.RollbackAndCloseConnection(tran);
            //        return "<result><system result=\"true\" errinfo=\"" + ex.Message + "\"/></result>";
            //        //CommonHelper.MsgError(ex.Message);
            //    }
            //}
            //else
            //{
            //    string sql = "delete from LK1_0003_PIC where cno = @cno";
            //    DbHelper.ExecuteNonQuery(sql, new SqlParameter[]{
            //                    new SqlParameter("@cno",cNo)});
                return null;
            //}
           
        }

        public string Excuting(VoucherProxy ReceiptObject)
        {
            DataSet ds = ReceiptObject.GetData(false, false);
            Business dt = ReceiptObject.Businesses["LK1_0007_E005"];
            Business dt1 = ReceiptObject.Businesses["LK1_0007_E001"];
            string sXzr = "/";

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sXzr =sXzr+ DbHelper.GetDbString(dt.Rows[i].Cells["xmzy"].Value) + "/";

                
                }
            
            }
            sXzr = sXzr + DbHelper.GetDbString(dt1.Rows[0].Cells["fzr"].Value) + "/";

            dt1.Rows[0].Cells["xzr"].Value = sXzr;
            return null;
        }

        #endregion
        #region 自定义参数
      
    
     

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
        #endregion

    }
}
    

