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

            //Business subEntity = ReceiptObject.Businesses["U8CUSTDEF_0018_E002"];//子表实体 U8CUSTDEF_0004_E002是子表实体编号，可以在UAP中看到

            //增加大量数据时，给DataSet赋值 然后调用 ReceiptObject.Refresh方法，值刷新一次界面。
            DataSet ds = ReceiptObject.GetData(false, false);
            Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
            string cFzr = DbHelper.GetDbString(dt.Rows[0].Cells["fzr"].Value);

            if (canshu.userName != "001" && canshu.userName != "demo" && canshu.cQx != "1" && canshu.userName != cFzr)
            {
                ReceiptObject.Businesses["LK1_0007_E002"].Columns["lxtfj"].Visible = false;
            }
            else
            {
                ReceiptObject.Businesses["LK1_0007_E002"].Columns["lxtfj"].Visible = true;
            }


            //DataTable dtSub = ds.Tables["sml_fwds"];
            //DataTable dt = subEntity.GetCheckedDataTabe();
            ////Business subEntity = ReceiptObject.Businesses["U8CUSTDEF_0004_E002"];//子表实体 U8CUSTDEF_0004_E002是子表实体编号，可以在UAP中看到
            //if (dt.Rows.Count > 0)
            //{
            //    string rowKey = subEntity.AddRow(); //rowKey是行 主键值
            //    subEntity.Rows[rowKey].Cells["xsxh"].Value = dt.Rows[0]["xsxh"].ToString();
            //    subEntity.Rows[rowKey].Cells["gzxx"].Value = dt.Rows[0]["gzxx"].ToString();
            //    subEntity.Rows[rowKey].Cells["Barcode"].Value = dt.Rows[0]["Barcode"].ToString();

            //    //for (int i = 0; i < 10; i++)
            //    //{

            //    //}
            //    //DataRow dr = dtSub.NewRow();
            //    //dr["U8CUSTDEF_0018_E001_PK"] = ReceiptObject.CurrentPKValue;
            //    ////dr["U8CUSTDEF_0018_E001_PK"] = dtSub.Rows[0]["U8CUSTDEF_0018_E001_PK"].ToString();  //必须给子表的外键 赋值，子表的外键值=主表的主键值
            //    //dr["xsxh"] = dt.Rows[0]["xsxh"].ToString();
            //    //dr[""] = dt.Rows[0]["gzxx"].ToString();
            //    //dr["Barcode"] = dt.Rows[0]["Barcode"].ToString();

            //    //dtSub.Rows.Add(dr);
            //}
            
            
            
            
            
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
    

