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
    class WordPrint2:IButtonEventHandler
    {

        #region IButtonEventHandler 成员

        public string Excute(VoucherProxy ReceiptObject, string PreExcuteResult)
        {
             DataSet ds = ReceiptObject.GetData(false, false);
            Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
            string cFzr = DbHelper.GetDbString(dt.Rows[0].Cells["fzr"].Value);

            if (canshu.userName != "001" && canshu.userName != "demo" && canshu.userName != cFzr && canshu.cQx != "1")
            {
                CommonHelper.MsgError("没有权限，无法打印");
                return MakeExcuteState(false, "没有权限");

            }

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
            //Business dt3 = ReceiptObject.Businesses["LK1_0007_E003"];  //原材料 
            //Business dt4 = ReceiptObject.Businesses["LK1_0007_E004"];  //其他费用
            Business dt5 = ReceiptObject.Businesses["LK1_0007_E002"];  //路线
          ////先判断有没有选择工艺路线
          //  if (dt5.Rows.Count ==0)
          //  {
          //      CommonHelper.MsgError("没有路线，无法选择");
          //      return MakeExcuteState(false, "没有路线");
              
          //      //return null;
            
          //  }
          //  else 
                if (dt5.Rows.Count > 0)
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

              
            }
                dt3 = DbHelper.Execute("select * from LK1_XM_BOM where LK1_0007_E002_PK =" + PK5.ToString() + "").Tables[0];

            string ls = Environment.CurrentDirectory;
            //string[] fileInfo = Directory.GetFiles(ls+"\\标签格式","*.rdlc");


            //求合计金额,通过书签写。  

            string sXshj = DbHelper.GetDbdecimal(dt3.Compute("sum(iprice)", "")).ToString("0.00");
            string sZshj = DbHelper.GetDbdecimal(dt3.Compute("sum(zsprice)", "")).ToString("0.00");

            string sFdhj = DbHelper.GetDbdecimal(dt3.Compute("sum(fdprice)", "")).ToString("0.00");


            //string ls = Environment.CurrentDirectory;
            //string[] fileInfo = Directory.GetFiles(ls+"\\标签格式","*.rdlc");
            mblj = ls + @"\uap\runtime\项目立项书模板.docx";
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
                        if (sbm == "xmrq" || sbm == "denddate")
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
                //else if (sbm == "qtfy" && dt4.Rows.Count > 0)
                //{
                //    #region 表格循环，使用光标向后移动实现。多行使用插入行。
                //    //选择到标签使用
                //    object what = Word.WdGoToItem.wdGoToBookmark;
                //    object bookname = sbm;
                //    wDoc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref bookname);
                //    //向右移动使用参数
                //    object dummy = System.Reflection.Missing.Value;
                //    object count = 1;
                //    object Unit = Word.WdUnits.wdCharacter;
                //    //MessageBox.Show(dts.Rows.Count.ToString());
                //    for (int i = 0; i < dt4.Rows.Count; i++)
                //    {

                //        //有序号行，先写序号
                //        wDoc.ActiveWindow.Selection.TypeText((i + 1).ToString());
                //        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                //        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbString(dt4.Rows[i].Cells["fymc"].Value));


                //        wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                //        wDoc.ActiveWindow.Selection.TypeText(DbHelper.GetDbdecimal(dt4.Rows[i].Cells["ifyprice"].Value).ToString("0.00"));

                //          if (i < dt4.Rows.Count - 1)
                //                wDoc.ActiveWindow.Selection.InsertRowsBelow(1);
                //            //wDoc.ActiveWindow.Selection.MoveRight(ref Unit, ref count, ref dummy);
                        

                //    }
                //    #endregion
                //}
                else if (sbm == "pic")
                {
                    if (!string.IsNullOrEmpty(fuJian))
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
                

            }

//------------------------使用Printout方法进行打印------------------------------
            object background = false; //这个很重要，否则关闭的时候会提示请等待Word打印完毕后再退出，加上这个后可以使Word所有

//页面发送完毕到打印机后才执行下一步操作
            //wDoc.SaveAs("项目立项书");
            wDoc.PrintOut(ref background); //打印
            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            wDoc.Close(ref saveOption); //关闭当前文档，如果有多个模版文件进行操作，则执行完这一步后接着执行打开Word文档的方法即可
            saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            wApp.Quit(ref saveOption); //关闭Word进程
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
