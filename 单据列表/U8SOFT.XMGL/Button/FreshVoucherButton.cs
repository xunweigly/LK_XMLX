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
    class FreshVoucherButton : IButtonEventHandler
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
            //ReceiptObject.VoucherState = VoucherStateEnum.Edit;

            //DbHelper.conStr = ReceiptObject.LoginInfo.UFDataSqlConStr;
            //DataSet ds = ReceiptObject.GetData(false, false);

            ////byte[] imageBytes = DbHelper.GetImageBytes(xmrz.currentControl.Image);

            //Business dt = ReceiptObject.Businesses["LK1_0007_E002"];
            //string cNo = DbHelper.GetDbString(dt.Rows[0].Cells["cNo"].Value);


            //string sql = "select pic from LK1_0003_PIC where cno = '" + cNo + "'";
            //DataTable dtpic = DbHelper.ExecuteTable(sql);
            //if (dtpic.Rows.Count > 0)
            //{
            //    //读取图片
            //    if (DBNull.Value != dtpic.Rows[0]["pic"])
            //    {
            //        byte[] buffByte = (byte[])dtpic.Rows[0]["pic"];
            //        MemoryStream buf = new MemoryStream(buffByte);
            //        Image image = System.Drawing.Image.FromStream(buf);


            //        xmrz.currentControl.Image = image;

            //    }
            //    else
            //    {

            //        xmrz.currentControl.Image = null;
            //    }
            //}
            //else
            //{

            //    xmrz.currentControl.Image = null;
            //}
            return null;
           
        }

        public string Excuting(VoucherProxy ReceiptObject)
        {
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
    

