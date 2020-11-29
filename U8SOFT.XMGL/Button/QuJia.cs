using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UFIDA.U8.UAP.UI.Runtime.Model;
using UFIDA.U8.UAP.UI.Runtime.Common;
using System.Windows.Forms;
using System.Data;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using fuzhu;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Drawing;
//using UFIDA.U8.UAP.UI.Runtime.Proxy;


namespace U8SOFT.XMRZ
{
    class QuJia: IButtonEventHandler
    {

        #region IButtonEventHandler 成员

        public string Excute(VoucherProxy ReceiptObject, string PreExcuteResult)
        {

            DbHelper.conStr = ReceiptObject.LoginInfo.UFDataSqlConStr;
            //DataSet ds = ReceiptObject.GetData(false, false);

            ////byte[] imageBytes = DbHelper.GetImageBytes(xmrz.currentControl.Image);

            Business dt = ReceiptObject.Businesses["LK1_0007_E001"];

            string cNo = dt.Rows[0].Cells["cNO"].Value;



             string sql = @" update bom set  bom.iunitcost=b.baojia1, bom.iprice =b.baojia1*bom.iquantity,zsunitcost = b.baojia2,fdunitcost = b.baojia3,
bom.zsprice =b.baojia2*bom.zsqty,
bom.fdprice =b.baojia3*bom.fdqty   from LK1_XM_BOM bom,zdy_lk_xunjia b,
LK_XM_LX lx
where  bom.LK1_0007_E001_PK = lx.LK1_0007_E001_PK
and b.xmbm= lx.cNo and b.cinvcode =bom.cinvcode and b.cinvstd = bom.cinvstd and b.cinvname=bom.cinvname
  and lx.cno =  '" + cNo + "'";
//            string sql = @"update bom set  bom.iunitcost=b.CMEMO1, bom.iprice =b.CMEMO1*bom.iquantity,zsunitcost = b.cmemo2,fdunitcost = b.cmemo3,
//bom.zsprice =b.CMEMO2*bom.zsqty,
//bom.fdprice =b.CMEMO3*bom.fdqty   from LK1_XM_BOM bom,U8CUSTDEF_0015_E001 a ,U8CUSTDEF_0015_E002 b,
//LK_XM_LX lx,LK_XM_lxs lxs
//where a.U8CUSTDEF_0015_E001_PK =b.U8CUSTDEF_0015_E001_PK
//and bom.LK1_0007_E001_PK = lx.LK1_0007_E001_PK
//and a.citemname= lx.cNo and b.cinvcode =bom.cinvcode and b.cinvstd = bom.cinvstd and lxs.cNo = a.xmlx
//and lx.LK1_0007_E001_PK = lxs.LK1_0007_E001_PK and bom.LK1_0007_E002_PK = lxs.LK1_0007_E002_PK
//and b.cinvname=bom.cinvname  and lx.cno = '" + cNo + "'";
            int i = DbHelper.ExecuteNonQuery(sql);
            if (i > 0)
            {
                MessageBox.Show("取价完成");

                ReceiptObject.RefreshVoucherButton();
            }


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

        //自定义参数
        public static string StringNull(object obj)
        {
            if (obj == null)
                return null;
            //return DBNull.Value;
            else if (obj == DBNull.Value)
                return null;
            else
                return obj.ToString();
        }

        public static int IntNull(object obj)
        {
            if (obj == null)
                return 0;
            //return DBNull.Value;
            else if (obj == DBNull.Value)
                return 0;
            else
                return Convert.ToInt16(obj);
        }
    }

}
