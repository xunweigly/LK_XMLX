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
    class DelVoucherButton : IButtonEventHandler
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

           
          //  DataSet ds = ReceiptObject.GetData(false, false);
          //  Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
          //  string cNo = DbHelper.GetDbString(dt.Rows[0].Cells["cNo"].Value);

          //  string sql = "delete from fitemss97 from fitemss97,fitemss97class where   fitemss97.cItemCcode = fitemss97class.cItemCcode and  fitemss97class.citemcname ='" + cNo + "'";

          //  DbHelper.ExecuteNonQuery(sql);
          //  sql = "delete from fitemss97class where citemcname ='" + cNo + "'";
          //DbHelper.ExecuteNonQuery(sql);



            
                return null;
            //}
           
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
    

