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
using UFIDA.U8.Pub.FileManager;


namespace U8SOFT.XMRZ
{
    class WordPrintXMZJ:IButtonEventHandler
    {

        #region IButtonEventHandler 成员

        public string Excute(VoucherProxy ReceiptObject, string PreExcuteResult)
        {
             DataSet ds = ReceiptObject.GetData(false, false);
            Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
            string cFzr = DbHelper.GetDbString(dt.Rows[0].Cells["fzr"].Value);

            if (canshu.userName != "001" && canshu.userName != "demo" && canshu.userName != "902")
            {
                CommonHelper.MsgError("没有权限，无法打印");
                return MakeExcuteState(false, "没有权限");

            }

           //获得项目编码

            string  xmbm = dt.Rows[0].Cells["cno"].Value+"-1";
            string cNo = dt.Rows[0].Cells["cno"].Value;
            //ReceiptObject.ModifyVoucherButton();
            ////获取数据库连接字符串
            ////MessageBox.Show(ReceiptObject.LoginInfo.UFDataSqlConStr);
            ////ReceiptObject.Businesses["uap001_0001_E001"]
            ////读取数据，通过读取数据表的数据
            //string scinvcode = Convert.ToDateTime(ReceiptObject.Businesses["U8CUSTDEF_0013_E001"].Rows[0].Cells["scrq"].Value).ToString("yyyy-MM-dd"); 
            //MessageBox.Show(scinvcode);
            ////DataTable dt = ReceiptObject.Businesses["U8CUSTDEF_0013_E002"];
           //检验单打印参数，是否含英文、有序号、路径
            //int xuhaohang = 0;
            int PK5 = 0;
            string fuJian="";
            //int xuhaohangno = 5;
            //int hyw = 1;
            string mblj =null;
            //数据表 表头、表体
            //Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
            Business dt2 = ReceiptObject.Businesses["LK1_0007_E005"];  //组员
            DataTable dt3 = new DataTable();
            Business dt5 = ReceiptObject.Businesses["LK1_0007_E002"];  //路线

            DataTable dtlxs = new DataTable();      //立项书
            DataTable dtlld = new DataTable();      // 领料单
            DataTable dtzjt = new DataTable();      //中间体
            DataTable dtyltk = new DataTable();   //原料退库
            DataTable dt4 = new DataTable();    //其他费用 
          //先判断有没有选择工艺路线
            if (dt5.Rows.Count ==0)
            {
                CommonHelper.MsgError("没有路线，无法选择");
                return MakeExcuteState(false, "没有路线");
              
                //return null;
            
            }
            else if (dt5.Rows.Count > 0)
            {
               
                for (int i = 0; i < dt5.Rows.Count ; i++)
                {
                    string cCheck = DbHelper.GetDbString(dt5.Rows[i].Cells["bcheck"].Value);
                    if (cCheck == "True")
                    {
                        PK5 = DbHelper.GetDbInt(dt5.Rows[i].Cells["LK1_0007_E002_PK"].Value);
                        fuJian =DbHelper.GetDbString(dt5.Rows[i].Cells["lxtfj_UAPRuntime_FileID"].Value);
                    }

                }
                if (PK5 == 0)
                {
                    CommonHelper.MsgError("没有选择路线，无法生成");

                    return MakeExcuteState(false, "没有路线");
                
                }

                dt3 = DbHelper.Execute("select * from LK1_XM_BOM where LK1_0007_E002_PK =" + PK5.ToString() + "").Tables[0];
               
                ////立项书
                //dtlxs = DbHelper.Execute("select cinvcode,cinvname,cinvstd,jiliangdw as jldw, isnull(iquantity,0)+isnull(zsqty,0)+isnull(fdqty,0)  as iQuantity from LK1_XM_BOM where LK1_0007_E002_PK =" + PK5.ToString() + "").Tables[0];
                dtlxs = DbHelper.Execute(@"  SELECT  a.cCode,a.dDate,b.cinvcode,c.InvName as cinvname,c.InvStd as cinvstd,b.fQuantity,c.ComUnitName as jldw,b.fMoney,d.chdefine4,a.cDefine14 FROM   PU_AppVouch a,   PU_AppVouchs b,dbo.v_bas_inventory c ,dbo.PU_AppVouch_extradefine d WHERE a.id = b.id 
     AND b.cinvcode  = c.InvCode AND a.id = d.id AND d.chdefine4 = '是' and b.citemcode = '"+xmbm+"'").Tables[0];
                // 领料单
                dtlld = DbHelper.Execute(@"select c.ccode,c.ddate, a.cInvCode,b.InvName as cinvname,b.InvStd as cinvstd,b.ComUnitName as jldw,a.iQuantity from rdrecords11 a,v_bas_inventory  b,rdrecord11 c
where a.cInvCode= b.InvCode and a.id = c.id and a.citemcode = '"+xmbm+"'").Tables[0];
                //中间体
                dtzjt = DbHelper.Execute(@"select c.ccode,c.ddate, a.cInvCode,b.InvName as cinvname ,b.InvStd as cinvstd,b.ComUnitName  as jldw,a.iQuantity from rdrecords10 a,v_bas_inventory  b,rdrecord10 c
where a.cInvCode= b.InvCode and isnull(a.cdefine33,'')<>'' and a.id = c.id and a.cdefine33<>'产品' and a.cdefine33 <>'材料'  and a.citemcode = '" + xmbm + "'").Tables[0];
                //原料退库
                dtyltk = DbHelper.Execute(@"select c.ccode,c.ddate, a.cInvCode,b.InvName as cinvname,b.InvStd as cinvstd,b.ComUnitName  as jldw,a.iQuantity from rdrecords10 a,v_bas_inventory  b,rdrecord11 c
where a.cInvCode= b.InvCode and a.id =c.id  and a.cdefine33 ='材料'  and a.citemcode = '" + xmbm + "'").Tables[0];
                //其他费用
                dt4 = DbHelper.Execute(@"     SELECT  a.cPOID,a.dPODate,b.cinvcode,c.InvName as cinvname,c.InvStd as cinvstd,b.iQuantity,c.ComUnitName,b.iSum FROM   dbo.PO_Pomain a,   dbo.PO_Podetails b,dbo.v_bas_inventory c WHERE a.POID = b.POID 
     AND b.cinvcode  = c.InvCode and (B.citemcode = '" + cNo + "-2" + "' OR  B.citemcode = '" + cNo + "-3" + "')").Tables[0]; 


            }
            string ls = Environment.CurrentDirectory;
            //string[] fileInfo = Directory.GetFiles(ls+"\\标签格式","*.rdlc");


            //求合计金额,通过书签写。  

            string sXshj = DbHelper.GetDbdecimal(dt3.Compute("sum(iprice)", "")).ToString("0.00");
            string sZshj = DbHelper.GetDbdecimal(dt3.Compute("sum(zsprice)", "")).ToString("0.00");

            string sFdhj = DbHelper.GetDbdecimal(dt3.Compute("sum(fdprice)", "")).ToString("0.00");


            //string ls = Environment.CurrentDirectory;
            //string[] fileInfo = Directory.GetFiles(ls+"\\标签格式","*.rdlc");
            mblj = ls + @"\uap\runtime\项目总结模板.docx";
            //word 打印
            ////word文档，word进程
            Word._Application wApp = new Word.Application();
            Word._Document wDoc = new Word.Document();

            object templatefile = mblj;
            object missing = System.Reflection.Missing.Value;
            wDoc = wApp.Documents.Add(ref templatefile); //在现有进程内打开文档
            wDoc.Activate(); //当前文档置前

            wApp.Visible = true;//显示word

            foreach (Word.Bookmark bm in wDoc.Bookmarks)
            {

                string sbm = bm.Name;

                if (sbm == "xshj" && sXshj != "0.00")
                {
                    bm.Range.Text = sXshj;
                }
                else if (sbm == "zshj" && sZshj != "0.00")
                {

                    bm.Range.Text = sZshj;
                }
                else if (sbm == "fdhj" && sFdhj != "0.00")
                {

                    bm.Range.Text = sFdhj;
                }


                //在主表中的栏目
                if (dt.Columns[sbm] != null)
                {
                    if (dt.Rows[0].Cells[sbm].Value != "")
                    {
                        if (sbm == "xmrq" || sbm == "denddate" || sbm=="sjwcdate")
                        {

                            bm.Range.Text = DbHelper.GetDbDate(dt.Rows[0].Cells[sbm].Value).ToString("yyyy-MM-dd");
                            
                        }
                        else if (sbm == "quantity")
                        {
                            bm.Range.Text = DbHelper.GetDbdecimal(dt.Rows[0].Cells[sbm].Value).ToString("0.00") +
                                DbHelper.GetDbString(dt.Rows[0].Cells["jldw"].Value);
                        } 
                        else if (sbm == "mubiaosl")
                        {
                            bm.Range.Text = DbHelper.GetDbdecimal(dt.Rows[0].Cells[sbm].Value).ToString("0.00") +
                                DbHelper.GetDbString(dt.Rows[0].Cells["wljldw"].Value);
                        } 

                        else if (sbm == "xmys")
                        {
                            if (dt.Rows[0].Cells[sbm].Value != "")
                                bm.Range.Text = DbHelper.GetDbdecimal(dt.Rows[0].Cells[sbm].Value).ToString("0.00");
                        }

                           
                        else
                        {
                            bm.Range.Text = DbHelper.GetDbString(dt.Rows[0].Cells[sbm].Value);
                        }
                    }

                }

 //实际周期  sjzhouqi 
                else if (sbm == "sjzhouqi" )
                {
                    if (dt.Rows[0].Cells["sjwcdate"].Value != "" && dt.Rows[0].Cells["xmrq"].Value != "")
                    {
                        TimeSpan ts =DbHelper.GetDbDate(dt.Rows[0].Cells["sjwcdate"].Value) - DbHelper.GetDbDate(dt.Rows[0].Cells["xmrq"].Value);
                        bm.Range.Text = ts.Days.ToString()+"天";
                    }
                }
                    //，实际项目成本 sjxmcb  ，请购单+项目立项书+其他费用
                else if (sbm == "sjxmcb")
                {
                    string sql = string.Format(@" select SUM(ISNULL(b.iprice,0))+SUM(ISNULL(b.zsprice,0))+SUM(ISNULL(b.fdprice,0)) zje   FROM dbo.LK_XM_LX a,dbo.LK1_XM_BOM b WHERE a.LK1_0007_E001_PK = b.LK1_0007_E001_PK
AND a.cNo = '{0}'", cNo);
                    decimal dLxsje =DbHelper.GetDbdecimal(DbHelper.ExecuteScalar(sql));
                    sql = string.Format(@" select SUM(ISNULL(fMoney,0)) qgdje   FROM PU_AppVouchs,pu_appvouch WHERE  citemcode =  '{0}'", xmbm);
                    decimal dQgdje = DbHelper.GetDbdecimal(DbHelper.ExecuteScalar(sql));
                    sql = string.Format(@" SELECT SUM(ISNULL(isum,0))  fyje  FROM po_podetails WHERE  citemcode = '{0}' or citemcode = '{1}'", cNo + "-2", cNo + "-3");
                    decimal dFyje = DbHelper.GetDbdecimal(DbHelper.ExecuteScalar(sql));
                    bm.Range.Text = (dLxsje + dQgdje + dFyje).ToString("0.00");

                   
                }
                //项目协助人
                else if (sbm == "xmzzr" && dt2.Rows.Count > 0)
                {
                    string sXmz = "";
                    for (int i = 0; i < dt2.Rows.Count - 1; i++)
                    {
                        sXmz = DbHelper.GetDbString(dt2.Rows[i].Cells["zyxm"].Value) + "/";

                    }
                    //最后一个没有"/"
                    sXmz += DbHelper.GetDbString(dt2.Rows[dt2.Rows.Count - 1].Cells["zyxm"].Value);
                    bm.Range.Text = sXmz;
                }
                else if (sbm == "ycl" && dt3.Rows.Count > 0)
                {
                    #region 表格循环，使用光标向后移动实现。多行使用插入行。
                    //选择到标签使用
                    object what = Word.WdGoToItem.wdGoToBookmark;
                    object bookname = sbm;
                    wDoc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref bookname);
                    //向右移动使用参数
                    object dummy = System.Reflection.Missing.Value;
                    object count = 1;
                    object Unit = Word.WdUnits.wdCharacter;
                    //MessageBox.Show(dts.Rows.Count.ToString());


                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {


                        //有序号行，先写序号
                        wDoc.ActiveWindow.Selection.TypeText((i + 1).ToString());
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dt3.Rows[i]["cinvname"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dt3.Rows[i]["cinvcode"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        //wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dt3.Rows[i]["jiliangdw"]));
                        //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        //判断是否为空
                        if (DbHelper.GetDbdecimal(dt3.Rows[i]["iquantity"]) != 0m)
                        {

                            wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dt3.Rows[i]["iquantity"]).ToString("0.00") + DbHelper.GetDbString(dt3.Rows[i]["jiliangdw"]));
                        }
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        if (DbHelper.GetDbdecimal(dt3.Rows[i]["iquantity"]) != 0m)
                        {
                            wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dt3.Rows[i]["iprice"]).ToString("0.00"));
                        }

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        if (DbHelper.GetDbdecimal(dt3.Rows[i]["zsqty"]) != 0m)
                        {
                            wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dt3.Rows[i]["zsqty"]).ToString("0.00") + DbHelper.GetDbString(dt3.Rows[i]["jiliangdw"]));
                        }
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);

                        if (DbHelper.GetDbdecimal(dt3.Rows[i]["zsqty"]) != 0m)
                        {
                            wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dt3.Rows[i]["zsprice"]).ToString("0.00"));
                        }
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);

                        if (DbHelper.GetDbdecimal(dt3.Rows[i]["fdqty"]) != 0m)
                        {
                            wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dt3.Rows[i]["fdqty"]).ToString("0.00") + DbHelper.GetDbString(dt3.Rows[i]["jiliangdw"]));
                        }
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        if (DbHelper.GetDbdecimal(dt3.Rows[i]["fdqty"]) != 0m)
                        {
                            wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dt3.Rows[i]["fdprice"]).ToString("0.00"));
                        }
                        if (i < dt3.Rows.Count - 1)
                            wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                        //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);


                    }
                    #endregion
                }
                    //追加物料清单
                else if (sbm == "zhuijia" && dtlxs.Rows.Count > 0)
                {
                    #region 表格循环，使用光标向后移动实现。多行使用插入行。
                    //选择到标签使用
                    object what = Word.WdGoToItem.wdGoToBookmark;
                    object bookname = sbm;
                    wDoc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref bookname);
                    //向右移动使用参数
                    object dummy = System.Reflection.Missing.Value;
                    object count = 1;
                    object Unit = Word.WdUnits.wdCharacter;
                    //MessageBox.Show(dts.Rows.Count.ToString());


                    for (int i = 0; i < dtlxs.Rows.Count; i++)
                    {


                        //有序号行，先写序号
                        wDoc.ActiveWindow.Selection.TypeText((i + 1).ToString());
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlxs.Rows[i]["ccode"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbDate(dtlxs.Rows[i]["ddate"]).ToString("yyyy-MM-dd"));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlxs.Rows[i]["cinvcode"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlxs.Rows[i]["cinvname"]));
                        
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlxs.Rows[i]["cinvstd"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);

                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dtlxs.Rows[i]["fQuantity"]).ToString("0.00"));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlxs.Rows[i]["jldw"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dtlxs.Rows[i]["fmoney"]).ToString("0.00"));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlxs.Rows[i]["chdefine4"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlxs.Rows[i]["cdefine14"]));

                            if (i < dtlxs.Rows.Count - 1)
                            wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                        //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);


                    }
                    wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                    wDoc.ActiveWindow.Selection.TypeText("合计");
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                    wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dtlxs.Compute("sum(fmoney)","")).ToString("0.00"));
                    #endregion
                }
                    //中间体清单 zjtqd

                else if (sbm == "zjtqd" && dtzjt.Rows.Count > 0)
                {
                    #region 表格循环，使用光标向后移动实现。多行使用插入行。
                    //选择到标签使用
                    object what = Word.WdGoToItem.wdGoToBookmark;
                    object bookname = sbm;
                    wDoc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref bookname);
                    //向右移动使用参数
                    object dummy = System.Reflection.Missing.Value;
                    object count = 1;
                    object Unit = Word.WdUnits.wdCharacter;
                    //MessageBox.Show(dts.Rows.Count.ToString());


                    for (int i = 0; i < dtzjt.Rows.Count; i++)
                    {


                        //有序号行，先写序号
                        wDoc.ActiveWindow.Selection.TypeText((i + 1).ToString());
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtzjt.Rows[i]["ccode"]));
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbDate(dtzjt.Rows[i]["ddate"]).ToString("yyyy-MM-dd"));
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtzjt.Rows[i]["cinvcode"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtzjt.Rows[i]["cinvname"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtzjt.Rows[i]["cinvstd"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);

                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dtzjt.Rows[i]["iquantity"]).ToString("0.00") );


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtzjt.Rows[i]["jldw"]));

                        if (i < dtzjt.Rows.Count - 1)
                            wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                        //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);


                    }
                    #endregion
                }
                    //领料清单 llqd
                else if (sbm == "llqd" && dtlld.Rows.Count > 0)
                {
                     #region 表格循环，使用光标向后移动实现。多行使用插入行。
                    //选择到标签使用
                    object what = Word.WdGoToItem.wdGoToBookmark;
                    object bookname = sbm;
                    wDoc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref bookname);
                    //向右移动使用参数
                    object dummy = System.Reflection.Missing.Value;
                    object count = 1;
                    object Unit = Word.WdUnits.wdCharacter;
                    //MessageBox.Show(dts.Rows.Count.ToString());


                    for (int i = 0; i < dtlld.Rows.Count; i++)
                    {


                        //有序号行，先写序号
                        wDoc.ActiveWindow.Selection.TypeText((i + 1).ToString());

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlld.Rows[i]["ccode"]));
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbDate(dtlld.Rows[i]["ddate"]).ToString("yyyy-MM-dd"));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlld.Rows[i]["cinvcode"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlld.Rows[i]["cinvname"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlld.Rows[i]["cinvstd"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);

                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dtlld.Rows[i]["iquantity"]).ToString("0.00"));
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtlld.Rows[i]["jldw"]));
                        if (i < dtlld.Rows.Count - 1)
                            wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                        //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);


                    }
                    #endregion
                  }
                    //退料清单 tlqd
                else if (sbm == "tlqd" && dtyltk.Rows.Count > 0)
                {
                    #region 表格循环，使用光标向后移动实现。多行使用插入行。
                    //选择到标签使用
                    object what = Word.WdGoToItem.wdGoToBookmark;
                    object bookname = sbm;
                    wDoc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref bookname);
                    //向右移动使用参数
                    object dummy = System.Reflection.Missing.Value;
                    object count = 1;
                    object Unit = Word.WdUnits.wdCharacter;
                    //MessageBox.Show(dts.Rows.Count.ToString());


                    for (int i = 0; i < dtyltk.Rows.Count; i++)
                    {


                        //有序号行，先写序号
                        wDoc.ActiveWindow.Selection.TypeText((i + 1).ToString());

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtyltk.Rows[i]["ccode"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtyltk.Rows[i]["ddate"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtyltk.Rows[i]["cinvcode"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtyltk.Rows[i]["cinvname"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dtyltk.Rows[i]["cinvstd"]));


                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);

                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dtyltk.Rows[i]["iquantity"]).ToString("0.00") + DbHelper.GetDbString(dtyltk.Rows[i]["jldw"]));

                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);

                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dtyltk.Rows[i]["jldw"]).ToString("0.00") + DbHelper.GetDbString(dtyltk.Rows[i]["jldw"]));

                        if (i < dtyltk.Rows.Count - 1)
                            wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                        //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);


                    }
                    #endregion
                }
                else if (sbm == "qtfy" && dt4.Rows.Count > 0)
                {
                    #region 表格循环，使用光标向后移动实现。多行使用插入行。
                    //选择到标签使用
                    object what = Word.WdGoToItem.wdGoToBookmark;
                    object bookname = sbm;
                    wDoc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref bookname);
                    //向右移动使用参数
                    object dummy = System.Reflection.Missing.Value;
                    object count = 1;
                    object Unit = Word.WdUnits.wdCharacter;
                    //MessageBox.Show(dts.Rows.Count.ToString());
                    for (int i = 0; i < dt4.Rows.Count; i++)
                    {

                      //有序号行，先写序号
                        wDoc.ActiveWindow.Selection.TypeText((i + 1).ToString());
                        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(DbHelper.GetDbString(dt4.Rows[i]["cpoid"])));


                      wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                      wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbDate(dt4.Rows[i]["dpodate"]).ToString("yyyy-MM-dd"));

                      wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                      wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(DbHelper.GetDbString(dt4.Rows[i]["cinvcode"])));
                      wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                      wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(DbHelper.GetDbString(dt4.Rows[i]["cinvname"])));
                      wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                      wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(DbHelper.GetDbString(dt4.Rows[i]["cinvstd"])));
                      wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                      wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(DbHelper.GetDbString(dt4.Rows[i]["iquantity"])));
                      wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                      wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(DbHelper.GetDbString(dt4.Rows[i]["jldw"])));
                      wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                      wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(DbHelper.GetDbdecimal(dt4.Rows[i]["isum"]).ToString("0.00"))); 

                        if (i < dt4.Rows.Count - 1)
                                wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                            //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);


                  }
                    #endregion
                }
                else if (sbm == "pic")
                {
                    bm.Select();
                    Selection sel = wDoc.ActiveWindow.Selection;

                    try
                    {

                        string FileName = ls + @"\tempcode.bmp";//图片所在路径
                        FileManagerClient client = new FileManagerClient();
                        client.FileOperator = "manager";
                        client.OperatorPassWord = "manager";
                        client.HostUrl = canshu.serverName;
                        client.Port = 80;
                        client.ProtocolType = "HTTP";
                        client.IsWeb = true;
                        client.ReadFile(fuJian, FileName);
                        sel.InlineShapes.AddPicture(FileName);

                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show(exception.ToString());
                    }







                }
                

            }

//------------------------使用Printout方法进行打印------------------------------
            object background = false; //这个很重要，否则关闭的时候会提示请等待Word打印完毕后再退出，加上这个后可以使Word所有

//页面发送完毕到打印机后才执行下一步操作
            //wDoc.SaveAs("项目立项书");
            //wDoc.PrintOut(ref background); //打印
            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            //wDoc.Close(ref saveOption); //关闭当前文档，如果有多个模版文件进行操作，则执行完这一步后接着执行打开Word文档的方法即可
            saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            //wApp.Quit(ref saveOption); //关闭Word进程
            //MessageBox.Show("打印结束");

            //wDoc = wApp.Documents.Open("检验单打印");



            return MakeExcuteState(true, "没有路线");
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
