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
            //Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
            //string cNo = DbHelper.GetDbString(dt.Rows[0].Cells["cNo"].Value);
            //DbHelper.conStr = ReceiptObject.LoginInfo.UFDataSqlConStr;


            //try
            //{
            //    SqlParameter[] param = new SqlParameter[]{ 
            //                             new SqlParameter("@cItemName",cNo),
            //                             new SqlParameter("@ccode","FN")
            //            };


            //    DbHelper.ExecuteNonQuery("zdy_lk_sp_Item_Insert", param, CommandType.StoredProcedure);



            return null;
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    //DbHelper.RollbackAndCloseConnection(tran);
            //    return "<result><system result=\"true\" errinfo=\"" + ex.Message + "\"/></result>";

            //}
            //}
           
        }

        public string Excuting(VoucherProxy ReceiptObject)
        {
            DataSet ds = ReceiptObject.GetData(false, false);
            Business dt = ReceiptObject.Businesses["LK1_0007_E005"];
            Business dt1 = ReceiptObject.Businesses["LK1_0007_E001"];
            string sXzr = "/";

            string cFzr = DbHelper.GetDbString(dt1.Rows[0].Cells["fzr"].Value);
            string cStatus = DbHelper.GetDbString(dt1.Rows[0].Cells["xmzt"].Value);
            string id = dt1.Rows[0].Cells["LK1_0007_E001_PK"].Value;
            string cNo = dt1.Rows[0].Cells["cNo"].Value;


            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sXzr = sXzr + DbHelper.GetDbString(dt.Rows[i].Cells["xmzy"].Value) + "/";



                }

            }
            //加小方  20180326

            sXzr = sXzr + DbHelper.GetDbString(dt1.Rows[0].Cells["xmgly"].Value) + "/";
            sXzr = sXzr + DbHelper.GetDbString(dt1.Rows[0].Cells["fzr"].Value) + "/";


            dt1.Rows[0].Cells["xzr"].Value = sXzr;


            //负责人不为空
            if (!string.IsNullOrEmpty(cFzr))
            {
                if (cStatus == "未分配" || string.IsNullOrEmpty(cStatus))
                {

                    dt1.Rows[0].Cells["xmzt"].Value = "分配完成";
                    //发送通知
                    string cMsg = string.Format("LK1_0007	{0}	STEFLK1_0007		a:{1}	y:{2}", id, canshu.acc, canshu.ztYear);
                    string cmemo = "项目立项书编码：" + cNo + "已分配给您！";

                    string sql = string.Format(@"INSERT INTO UFSystem..ua_message (cmsgid,nmsgtype,cmsgtitle,cmsgcontent,csender,creceiver,dsend,nvalidday,bhasread,nurgent, account,[year],cmsgpara)
                        VALUES(newid(),2000,'{0}','{1}','{2}','{3}',getdate(),4,0,0,'{4}','{5}','{6}')",
                     cmemo, cmemo, canshu.u8Login.cUserId, cFzr, canshu.acc, canshu.ztYear, cMsg);
                    DbHelper.ExecuteNonQuery(sql);
                }

                

            }
            else
            {

                dt1.Rows[0].Cells["xmzt"].Value = "未分配";
            }
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
    

