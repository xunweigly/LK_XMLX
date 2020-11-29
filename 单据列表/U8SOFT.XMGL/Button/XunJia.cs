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
using System.Xml;

//using UFIDA.U8.Audit.Interface;
//using UFIDA.U8.Audit.BusinessInfo;
//using UFIDA.U8.Audit.BusinessService;

using UFIDA.U8.Audit.ServiceProxy;
//using UFSoft.U8.Framework.Login.UI;
using UFSoft.U8.Framework.LoginContext;



namespace U8SOFT.XMRZ
{
    class XunJia:IButtonEventHandler
    {

        #region IButtonEventHandler 成员
       //private const string SubId = "DP";		//需要根据各业务子系统进行替换

        public string Excute(VoucherProxy ReceiptObject, string PreExcuteResult)
        {

            DataSet ds = ReceiptObject.GetData(false, false);
            Business dt = ReceiptObject.Businesses["LK1_0007_E001"];
            //Business dt2 = ReceiptObject.Businesses["LK1_0002_E002"];
            //if (dt2==null)
            string cNo = DbHelper.GetDbString(dt.Rows[0].Cells["cNo"].Value);
            DbHelper.conStr = ReceiptObject.LoginInfo.UFDataSqlConStr;


            //SqlTransaction tran = DbHelper.BeginTrans();
            try
            {
                SqlParameter[] param = new SqlParameter[]{ 
                                         new SqlParameter("@cno",cNo),
                                         new SqlParameter("@error",SqlDbType.NVarChar,100)
          };
                param[1].Direction = ParameterDirection.Output;

                DbHelper.ExecuteNonQuery("zdy_sp_lk_xunjia", param, CommandType.StoredProcedure);

                string sRe = DbHelper.GetDbString(param[1].Value);
                if (string.IsNullOrEmpty(sRe) == false)
                {
                    MessageBox.Show(sRe);

                    return  null;

                }
                MessageBox.Show("生成询价单完成");


       

	
//            //创建审批服务的客户端代理
//        AuditServiceProxy auditSvc = new AuditServiceProxy();

//        //构造Login的 CalledContext对象
//        CalledContext calledCtx = new CalledContext();
//        calledCtx.subId = "SA";
//        calledCtx.TaskID=canshu.u8Login.get_TaskId();
//        calledCtx.token = canshu.u8Login.userToken;
//        //业务对象标识
//        string bizObjectId = "UAPFORM.U8CUSTDEF_0015";
//        //业务事件标识  
//        string bizEventId = "U8CUSTDEF_0015.Commit";
//        //单据号
//        string voucherId = "41";
//        if (bizEventId == string.Empty || bizObjectId == string.Empty)
//    {
//    MessageBox.Show("请选择选择业务对象或业务事件!");
//    return null ;
//    }
//             bool bControled=true;
//             string errMsg="";
//             bool ret = auditSvc.SubmitApplicationMessage(bizObjectId, bizEventId, voucherId, calledCtx, ref bControled, ref errMsg);
//if (ret == true && bControled)
//    MessageBox.Show("提交成功");
//else
//    MessageBox.Show("提交失败，失败原因：" + errMsg);
     
        
  

                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //DbHelper.RollbackAndCloseConnection(tran);
                return "<result><system result=\"true\" errinfo=\"" + ex.Message + "\"/></result>";

            }



            //return null;
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


        ///// <summary>
        ///// 通过存储过程获取即将保存的单据可以使用的主键值
        ///// </summary>
        ///// <param name="conn"></param>
        ///// <param name="tran"></param>
        ///// <param name="uglogin">u8登陆对象</param>
        ///// <param name="amount">子表行数</param>
        ///// <param name="ifatherid">得到主表主键值</param>
        ///// <param name="iChildid">得到所有子表记录可以使用的主键的最大值</param>
        //private void GetNewSAVoucherID(SqlConnection conn, SqlTransaction tran, U8Login.clsLogin uglogin, int amount, ref int ifatherid, ref int iChildid)
        //{
        //    //获取remoteid
        //    SqlCommand queryRemoteid = new SqlCommand("select cvalue from accinformation where cname='cid'", conn, tran);
        //    string remoteid = queryRemoteid.ExecuteScalar().ToString();
        //    //调用存储过程sp_GetID
        //    SqlCommand queryID = new SqlCommand("sp_GetID", conn, tran);
        //    queryID.CommandType = System.Data.CommandType.StoredProcedure;
        //    SqlParameter remotingIDPara = new SqlParameter("@RemoteId", System.Data.SqlDbType.NVarChar, 2);
        //    remotingIDPara.Value = remoteid;
        //    SqlParameter cAcc_IdPara = new SqlParameter("@cAcc_Id", System.Data.SqlDbType.NVarChar, 3);
        //    cAcc_IdPara.Value = uglogin.get_cAcc_Id();
        //    SqlParameter cVouchTypePara = new SqlParameter("@cVouchType ", System.Data.SqlDbType.NVarChar, 50);
        //    cVouchTypePara.Value = "Somain";
        //    SqlParameter iAmountPara = new SqlParameter("@iAmount", System.Data.SqlDbType.Int);
        //    iAmountPara.Value = amount;
        //    SqlParameter iFatherIdPara = new SqlParameter("@iFatherId", System.Data.SqlDbType.Int);
        //    iFatherIdPara.Direction = System.Data.ParameterDirection.Output;
        //    SqlParameter iChildIdPara = new SqlParameter("@iChildId", System.Data.SqlDbType.Int);
        //    iChildIdPara.Direction = System.Data.ParameterDirection.Output;
        //    queryID.Parameters.Add(remotingIDPara);
        //    queryID.Parameters.Add(cAcc_IdPara);
        //    queryID.Parameters.Add(cVouchTypePara);
        //    queryID.Parameters.Add(iAmountPara);
        //    queryID.Parameters.Add(iFatherIdPara);
        //    queryID.Parameters.Add(iChildIdPara);
        //    queryID.ExecuteNonQuery();
        //    ifatherid = int.Parse(iFatherIdPara.Value.ToString());
        //    iChildid = int.Parse(iChildIdPara.Value.ToString());
        //}

        ///// <summary>
        ///// 获取即将生成的销售订单的订单流水号
        ///// </summary>
        ///// <param name="u8Login">U8登陆对象</param>
        ///// <returns>订单流水号</returns>
        //private string getNewSAVourchCode(U8Login.clsLogin u8Login)
        //{
        //    try
        //    {

        //        UFBillComponent.clsBillComponent CodeServer = new UFBillComponent.clsBillComponent();
        //        XmlDocument xmldom = new XmlDocument();
        //        string cardnum = "17";
        //        if (CodeServer.InitBill(u8Login.UfDbName, ref cardnum) == false)
        //        {
        //            throw new Exception("初始化单据号码失败！");
        //        }
        //        xmldom.LoadXml(CodeServer.GetBillFormat());
        //        if (xmldom.SelectSingleNode("//单据编号").Attributes["允许手工修改"].Value.ToLower() == "true")
        //        {
        //            return "";
        //        }
        //        if (xmldom.SelectNodes("//单据编号/前缀[@对象类型!=5]") != null)
        //        {
        //            XmlNodeList xmlist = xmldom.SelectNodes("//单据编号/前缀[@对象类型!=5]");
        //            for (int i = 0; i < xmlist.Count; i++)
        //            {
        //                if (xmlist[i].Attributes["对象类型"].Value.ToString() == "7")
        //                {
        //                    xmlist[i].Attributes["种子"].Value = u8Login.CurDate.ToString();
        //                }
        //                else if (xmlist[i].Attributes["对象类型"].Value.ToString() != "3")
        //                {
        //                    xmlist[i].Attributes["种子"].Value = "";
        //                }
        //            }
        //        }
        //        return CodeServer.GetNumber(xmldom.OuterXml, true);
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}

    }

}
