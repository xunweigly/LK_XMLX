using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using UFIDA.U8.UAP.UI.Runtime.Model;
using UFIDA.U8.UAP.UI.Runtime.Common;
using UFIDA.U8.Pub.FileManager;
using UFIDA.U8.UAP.Common;
using System.Data;
using fuzhu;
using System.IO;
using System.Drawing;
//using UFIDA.U8.Framework.Lib.Context;

namespace U8SOFT.XMRZ
{
    public class xmrz : ReceiptPluginBase
    {
        public static UserControl2 currentControl = null;
        
        //[IsImplementAttribute(true)]
        //public override System.Windows.Forms.Control CreateControl(BusinessProxy businessObject, VoucherProxy voucherObject, string ID)
        //{


        //    Control resultControl = null;
        //    if (currentControl != null)
        //        currentControl.Dispose();
        //    currentControl = null;
        //    //控件id
        //    if (ID == "a339024b-1425-49e5-ae7b-725a7dedf965")
        //    {
        //        //使用usercontrol2  自定义控件
        //        UserControl2 tempControl = new UserControl2(businessObject, voucherObject);
        //        resultControl = tempControl;
        //        currentControl = tempControl;
        //        tempControl.Dock = DockStyle.Fill;
        //    }
        //    return resultControl;
        //}


        [IsImplement(true)]
        public override IButtonEventHandler GetButtonEventHandler(UFIDA.U8.UAP.UI.Runtime.Common.VoucherButtonArgs ButtonArgs, VoucherProxy voucherObject)
        {


            switch (ButtonArgs.ButtonKey)
            {
                //case "zhantie":
                //    //图片粘贴功能是新增加的功能。          
                //    return new CancelVoucherButton();
                case "btnSaveVoucher":
                    return new SaveVoucherButton();
                case "btnRefresh":
                    return new FreshVoucherButton();
                case "btnFirst":
                    return new FreshVoucherButton();
                case "btnLast":
                    return new FreshVoucherButton();
                case "btnPrecede":
                    return new FreshVoucherButton();
                case "btnNext":
                    return new FreshVoucherButton();
                case "btnExport":
                    return new WordPrint();
                case "btnPrint":
                    return new WordPrint2();
                case "btnXMZJ":
                    return new WordPrintXMZJ();
                //case "btnSubmit":
                //    return new XunJia();
                //case "采购询价":
                //    return new XunJia();
                //case "取价":
                //    return new QuJia();
                case "btnAddVoucher":
                    return new AddVoucherButton();
                case "btnDelVoucher":
                    return new DelVoucherButton();

                //case "dayin3":
                //    return new Dayin();
                    
            }
            return null;
        }

        /// <summary>
        /// 单据加载前
        /// </summary>
        /// <param name="login"></param>
        /// <param name="Cardnumber"></param>
        /// <param name="ds"></param>
        /// <param name="state"></param>
        /// <param name="loadingParas"></param>
        [IsImplement(true)]
        public override void ReceiptLoading(U8Login.clsLogin login, string Cardnumber, DataSet ds, VoucherStateEnum state, ReceiptLoadingParas loadingParas)
        {
            //设置查询权限
            DbHelper.conStr = login.UFDataConnstringForNet;
            canshu.userName = login.cUserId;
            canshu.ztYear = login.cBeginYear;

            string sql = "select cSysUserName from UA_User where cSysUserName is not null and  cUser_Id='" + canshu.userName + "'";
            DataTable dt = DbHelper.ExecuteTable(sql);
            if (dt.Rows.Count > 0)
            {
                canshu.cQx = DbHelper.GetDbString(dt.Rows[0]["cSysUserName"]);
            }
            else
            {
                canshu.cQx = "0";
            
            }

            if (canshu.userName != "001" && canshu.userName != "demo" && canshu.cQx !="1")
                //&& canshu.userName != "056"，20180302 取消
            {
                //loadingParas.DefaultCondition = "<filteritems><table name='LK_XM_LX'><Item key='cNo' operator1='=' val1='123' /></table></filteritems>";
                loadingParas.DefaultCondition = "<filteritems><table name='LK_XM_LX'><Item key='xzr' operator1='like' val1='%/" + canshu.userName + "/%' /></table></filteritems>";
                //Voucher
            }
        }

        /// <summary>
        /// 修改参照
        /// </summary>
        /// <param name="para">修改参照</param>
        /// <param name="businessObject">所属业务对象</param>
        /// <param name="voucherObject">所属表单对象</param>
        /// </summary>
        [IsImplement(true)]
        public override bool ReferOpening(ReferOpenEventArgs para, BusinessProxy businessObject, VoucherProxy voucherObject)
        {
            Business dt = voucherObject.Businesses["LK1_0007_E003"];
            if (para.Column.ColumnName == "cinvcode")
            {
                string fileterSql = " #FN[dedate] is null ";

                //MessageBox.Show(para.RefService.FilterSQL);
                para.RefService.FilterSQL = fileterSql;
                //MessageBox.Show(para.RefService.FilterSQL);
            }


            return true;
        }


        [IsImplement(true)]
        public override void ReceiptLoaded(VoucherProxy ReceiptObject)
        {
           
            DataSet ds = ReceiptObject.GetData(false, false);
            canshu.serverName = ReceiptObject.LoginInfo.AppServer;

            //U8Login.clsLogin u8Login = new U8Login.clsLogin();
            canshu.u8Login = ReceiptObject.VBLoginObject;

            // 除了管理员和负责人外，其他人看不到路线图
            Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
            string cFzr = DbHelper.GetDbString(dt.Rows[0].Cells["fzr"].Value);
            string cGly = DbHelper.GetDbString(dt.Rows[0].Cells["xmgly"].Value);
            if (canshu.userName != "001" && canshu.userName != "demo" && canshu.cQx != "1" && canshu.userName != cFzr && canshu.userName != cGly)
            {

                ReceiptObject.Businesses["LK1_0007_E002"].Columns["lxtfj"].Visible = false;
            }
            else
            {
                ReceiptObject.Businesses["LK1_0007_E002"].Columns["lxtfj"].Visible = true;
            }

            //UFIDA.U8.UAP.Common.LoginInfo login = new UFIDA.U8.UAP.Common.LoginInfo(ListService.VbLogin);
            //byte[] imageBytes = DbHelper.GetImageBytes(xmrz.currentControl.Image);

            //Business dt = ReceiptObject.Businesses["LK1_0007_E002"];
            //string cNo = DbHelper.GetDbString(dt.Rows[0].Cells["cNo"].Value);


            //string sql = "select pic from LK1_0003_PIC where cno = '"+cNo+"'";
            //DataTable dtpic = DbHelper.ExecuteTable(sql);
            //if (dtpic.Rows.Count > 0)
            //{
            //    //读取图片
            //    if (DBNull.Value != dtpic.Rows[0]["pic"])
            //    {
            //        byte[]  buffByte = (byte[])dtpic.Rows[0]["pic"];
            //        MemoryStream buf = new MemoryStream(buffByte);
            //        Image image = System.Drawing.Image.FromStream(buf);


            //        xmrz.currentControl.Image = image;

            //    }
            //    else
            //    {

            //        xmrz.currentControl.Image = null;
            //    }
            //}

        }

         //<summary>
         //鼠标左键双击数据单元格的接口
         //</summary>
         //<param name="para">鼠标双击信息</param>
         //<param name="businessObject">所属业务对象</param>
         //<param name="voucherObject">所属表单对象</param>
         //</summary>
         [IsImplement(true)]
        public override void CellDoubleClick(UFIDA.U8.UAP.UI.Runtime.Common.CellDoubleClickEventArgs para, BusinessProxy businessObject, VoucherProxy voucherObject)
        {
        //    try
        //          {
        //             

        //    FileManagerClient client = new FileManagerClient();
        //    client.FileOperator = "manager";
        //    client.OperatorPassWord = "manager";
        //    client.HostUrl = "PC201612191131";
        //    client.Port = 80;
        //    client.ProtocolType = "HTTP";
        //    client.IsWeb = true;
        //    client.ReadFile("39ba4ab3-df57-4788-8e72-4e69acccd398.txt", @"C:\1.txt");
        //    //client.AddFile(@"C:\2.TXT", "somebody", 0x2800000, "001", "001", 0x7d5, true);
        //}
        //catch (Exception exception)
        //{
        //    MessageBox.Show(exception.ToString());
        //}
    


            //  //判断为表体部分的第一子表的单元格数据发生变化
            //if (businessObject.EntityID == "LK1_0007_E002")
            //{
            //    string row = para.RowKey;
            //    string col = para.ColumnName;
            //    string sv = businessObject.Rows[row].Cells[col].Value;
            //    MessageBox.Show(sv);
            //    Clipboard.SetData(DataFormats.Text, sv);
            //}
            //MessageBox.Show("eerror");
             //throw new Exception("The method or operation is not implemented.");
        }
        /// <summary>
        /// 值更新之后的接口，对值的后续处理（如对其他Cell值的变更）在这里进行
        /// </summary>
        /// <param name="para">Cell的值变动参数</param>
        /// <param name="businessObject">所属业务对象</param>
        /// <param name="voucherObject">所属表单对象</param>
        [IsImplement(true)]
        public override void CellChanged(UFIDA.U8.UAP.UI.Runtime.Common.CellChangeEventArgs para, BusinessProxy businessObject, VoucherProxy voucherObject)
        {




            //判断为表体部分的第一子表的单元格数据发生变化
            if (businessObject.EntityID == "LK1_0007_E003")
            {
                //根据para中的ColumnName属性判断当前发生变化的为哪一列。
                //方法参数para中的其他属性介绍：
                //para.NewValue 当前单元格发生改变后的值
                //para.OldValue 当前单元格发生改变前的值
                //para.RefReturnData 当前单元格所在的列如果为“基础资料”字段，选中参照返回的行数据。
                //para.RowKey 当前单元格所在的行的主键值，内存中DataTable的主键值，不是数据库中的主键值。
                switch (para.ColumnName)
                {
                    //输入维修费，计算保内金额
                    case "cinvcode":
                        cCinvocdeCellChanged(para, businessObject);
                        break;
                    case "iunitcost":
                        cJGCellChanged(para, businessObject);
                        break;
                    case "iquantity":
                        cJGCellChanged(para, businessObject);
                        break;
                    case "zsunitcost":
                        cJGCellChanged(para, businessObject);
                        break;
                    case "zsqty":
                        cJGCellChanged(para, businessObject);
                        break;
                    case "fdunitcost":
                        cJGCellChanged(para, businessObject);
                        break;
                    case "fdqty":
                        cJGCellChanged(para, businessObject);
                        break;
                    ////输入折扣价，根据折扣价计算折扣和折扣金额
                    //case "cRealPrice":
                    //    cRealPriceCellChanged(para, businessObject);
                    //break;
                    default:
                        break;
                }
            }
            else if (businessObject.EntityID == "LK1_0007_E005")
            {
                //根据para中的ColumnName属性判断当前发生变化的为哪一列。
                //方法参数para中的其他属性介绍：
                //para.NewValue 当前单元格发生改变后的值
                //para.OldValue 当前单元格发生改变前的值
                //para.RefReturnData 当前单元格所在的列如果为“基础资料”字段，选中参照返回的行数据。
                //para.RowKey 当前单元格所在的行的主键值，内存中DataTable的主键值，不是数据库中的主键值。
                switch (para.ColumnName)
                {
                    //输入维修费，计算保内金额
                    case "xmzy":
                        cRYCellChanged(para, businessObject);
                        break;
                }
            }
            else if (businessObject.EntityID == "LK1_0007_E001")
            {
                //根据para中的ColumnName属性判断当前发生变化的为哪一列。
                //方法参数para中的其他属性介绍：
                //para.NewValue 当前单元格发生改变后的值
                //para.OldValue 当前单元格发生改变前的值
                //para.RefReturnData 当前单元格所在的列如果为“基础资料”字段，选中参照返回的行数据。
                //para.RowKey 当前单元格所在的行的主键值，内存中DataTable的主键值，不是数据库中的主键值。
                switch (para.ColumnName)
                {
                    //输入维修费，计算保内金额
                    case "xmbm":
                        cXMChanged(para, businessObject);
                        break;
                    case "xmrq":
                        cXMChanged(para, businessObject);
                        break;
                    default:
                        break;
                }
            }
        }

        #region 私有方法
        /// <summary>
        /// 输入存货编码，读取存货名称、规格、计量单位
        /// </summary>
        /// <param name="para"></param>
        /// <param name="businessObject"></param>
        private void cRYCellChanged(UFIDA.U8.UAP.UI.Runtime.Common.CellChangeEventArgs para, BusinessProxy businessObject)
        {
            string cPersoncode;

            cPersoncode = businessObject.Rows[para.RowKey].Cells["xmzy"].Value;
           string sql = @"select cpsn_name from hr_hi_person where   cpsn_num = '" + cPersoncode + "'";
            DataTable dtpic = DbHelper.ExecuteTable(sql);
            if (dtpic.Rows.Count > 0)
            {
                businessObject.Rows[para.RowKey].Cells["zyxm"].Value = DbHelper.GetDbString(dtpic.Rows[0]["cpsn_name"]);
            }

         

        }

        private void cXMChanged(UFIDA.U8.UAP.UI.Runtime.Common.CellChangeEventArgs para, BusinessProxy businessObject)
        {
          string xmbm   = businessObject.Rows[para.RowKey].Cells["xmbm"].Value;
string dDate =  businessObject.Rows[para.RowKey].Cells["xmrq"].Value;
string bm = xmbm + "-" + Convert.ToDateTime(dDate).ToString("yyyyMMdd");
string sql = "select count(1),max(cno) from lk_xm_lx where cNo like '" + bm + "%'";
DataTable dt =DbHelper.ExecuteTable(sql);
if (dt.Rows[0][0].ToString() =="0")
{
    businessObject.Rows[para.RowKey].Cells["cNo"].Value = xmbm + "-" + Convert.ToDateTime(dDate).ToString("yyyyMMdd") + "01";
}
else
{
    string bm2 = dt.Rows[0][1].ToString();
    businessObject.Rows[para.RowKey].Cells["cNo"].Value = bm + (Convert.ToInt32(bm2.Substring(bm2.Length - 2, 2))+1).ToString("00"); 
}



        }

            /// <summary>
        /// 输入人员,跳出组员
        /// </summary>
        /// <param name="para"></param>
        /// <param name="businessObject"></param>
        private void cCinvocdeCellChanged(UFIDA.U8.UAP.UI.Runtime.Common.CellChangeEventArgs para, BusinessProxy businessObject)
        {
            string cInvcode;

            cInvcode = businessObject.Rows[para.RowKey].Cells["cinvcode"].Value;
           string sql = @"select a.cInvCode,a.cInvName,a.cinvaddcode,  a.cInvStd,b.cComUnitName from inventory a,ComputationUnit b where isnull(a.dedate,'')='' and   a.cComUnitCode = b.cComUnitcode and  cinvcode = '" + cInvcode + "'";
            DataTable dtpic = DbHelper.ExecuteTable(sql);
            if (dtpic.Rows.Count > 0)
            {
                businessObject.Rows[para.RowKey].Cells["cinvname"].Value = DbHelper.GetDbString(dtpic.Rows[0]["cinvname"]);
                businessObject.Rows[para.RowKey].Cells["cinvaddcode"].Value = DbHelper.GetDbString(dtpic.Rows[0]["cinvaddcode"]);
                businessObject.Rows[para.RowKey].Cells["cinvstd"].Value = DbHelper.GetDbString(dtpic.Rows[0]["cinvstd"]);
                businessObject.Rows[para.RowKey].Cells["jiliangdw"].Value = DbHelper.GetDbString(dtpic.Rows[0]["cComUnitName"]);

            }

         

        }

        private void cJGCellChanged(UFIDA.U8.UAP.UI.Runtime.Common.CellChangeEventArgs para, BusinessProxy businessObject)
        {
            decimal dDj = 0;
            decimal dSl = 0;
            decimal dzsDj = 0;
            decimal dzsSl = 0;
            decimal dfdDj = 0;
            decimal dfdSl = 0;


            dDj = DbHelper.GetDbdecimal(businessObject.Rows[para.RowKey].Cells["iunitcost"].Value);
            dSl = DbHelper.GetDbdecimal(businessObject.Rows[para.RowKey].Cells["iquantity"].Value);
            dzsDj = DbHelper.GetDbdecimal(businessObject.Rows[para.RowKey].Cells["zsunitcost"].Value);
            dzsSl = DbHelper.GetDbdecimal(businessObject.Rows[para.RowKey].Cells["zsqty"].Value);
            dfdDj = DbHelper.GetDbdecimal(businessObject.Rows[para.RowKey].Cells["fdunitcost"].Value);
            dfdSl = DbHelper.GetDbdecimal(businessObject.Rows[para.RowKey].Cells["fdqty"].Value);


            businessObject.Rows[para.RowKey].Cells["iprice"].Value = decimal.Parse(string.Format("{0:#,###.00}", dDj * dSl)).ToString();
            businessObject.Rows[para.RowKey].Cells["zsprice"].Value = decimal.Parse(string.Format("{0:#,###.00}", dzsDj * dzsSl)).ToString();
            businessObject.Rows[para.RowKey].Cells["fdprice"].Value = decimal.Parse(string.Format("{0:#,###.00}", dfdDj * dfdSl)).ToString();

        }
        #endregion
    }
}
