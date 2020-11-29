using System;
using System.Collections.Generic;
using System.Text;
using UFIDA.U8.UAP.UI.Runtime.List.UI;
using UFIDA.U8.UAP.UI.Runtime.List;
using UFIDA.U8.UAP.UI.Runtime.List.Metas;
using System.Data;
using UFGeneralFilter;
using Infragistics.Win.UltraWinGrid;
using UFIDA.U8.UAP.UI.Runtime.List.Filter;
using UFIDA.U8.UAP.UI.Runtime.Common;

using System.Windows.Forms;
using fuzhu;


namespace U8SOFT.XMRZ
{
    public class ListImpl : UFIDA.U8.UAP.UI.Runtime.List.UI.BaseUIEventHandler
    {
        /// <summary>
        /// 处理界面相关事件
        /// </summary>
        /// <param name="sender">消息触发者</param>
        /// <param name="eventType">事件类型</param>
        /// <param name="args">与事件有关的参数信息</param>
        public override void ProcessEvent(UIEventTypeEnum eventType, object sender, object args)
        {
            //使用默认的处理方式
            switch (eventType)
            {
                //新增按钮处理
                case UIEventTypeEnum.CUSTOMBUTTONCLICK:
                    {
                        //获得传入的当前点击的按钮参数
                        ToolbarItemMeta item = args as ToolbarItemMeta;
                        string buttonID = item.Id;
                        //根据按钮的ID来区分不同的新增按钮
                        if (buttonID == "btnListTest")
                        {
                          MessageBox.Show("123");
                            //XXXXX按钮的处理过程
                        }
                        else if (buttonID == "YYYYY")
                        {
                            //YYYYY按钮的处理过程
                        }

                        break;
                    }
                //已有功能按钮的处理
                default:
                    {
                        base.ProcessEvent(eventType, sender, args);
                        break;
                    }
            }
        }

        //#region 已有功能按钮的重写
        ///// <summary>
        ///// 处理导出事件函数
        ///// </summary>
        ///// <param name="sender">消息触发者</param>
        ///// <param name="args">导出事件对应的参数</param>
        //protected override void ProcessExportEvent(object sender, object args)
        //{

        //}
        ///// <summary>
        ///// 处理预览事件函数
        ///// </summary>
        ///// <param name="sender">消息触发者</param>
        ///// <param name="args">预览事件对应的参数</param>
        //protected override void ProcessPreviewEvent(object sender, object args)
        //{

        //}
        ///// <summary>
        ///// 处理打印事件函数
        ///// </summary>
        ///// <param name="sender">消息触发者</param>
        ///// <param name="args">打印事件对应的参数</param>
        //protected override void ProcessPrintEvent(object sender, object args)
        //{

        //}

        ///// <summary>
        ///// 过滤处理事件
        ///// </summary>
        ///// <param name="sender">消息触发者</param>
        ///// <param name="args">过滤处理相关的参数</param>
        //protected override void ProcessFilterEvent(object sender, object args)
        //{

        //}
        ///// <summary>
        ///// 处理列显示设置信息
        ///// </summary>
        ///// <param name="sender">消息触发者</param>
        ///// <param name="args">列显示设置相关参数</param>
        //protected override void ProcessColumnSetEvent(object sender, object args)
        //{

        //}
        ///// <summary>
        ///// 处理双击事件
        ///// </summary>
        ///// <param name="sender">消息触发者</param>
        ///// <param name="args">参数</param>
        //protected override void ProcessDBClickEvent(object sender, object args)
        //{

        //}

        ///// <summary>
        ///// 单据复制
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="args"></param>
        //protected override void ProcessVoucherCopy(object sender, object args)
        //{

        //}

        ///// <summary>
        ///// 处理列表汇总事件
        ///// </summary>
        ///// <param name="sender">消息触发者</param>
        ///// <param name="args">参数信息</param>
        //protected override void ProcessVoucherSumEvent(object sender, object args)
        //{

        //}
        ///// <summary>
        ///// 处理刷新按钮事件
        ///// </summary>
        ///// <param name="sender">消息发送者</param>
        ///// <param name="args">消息参数</param>
        //protected override void ProcessRefreshEvent(object sender, object args)
        //{

        //}
        //#endregion

        /// <summary>
        /// 过滤窗口弹出时的处理，可以动态设置缺省过滤条件
        /// </summary>
        /// <param name="listService">列表服务</param>
        /// <param name="filterSrv">过滤组件</param>
        /// 

         //[IsImplement(true)]
        public override void ReceiptListFiltering(UFGeneralListService listService, FilterSrvClass filterSrv)
        {
           
            U8FilterService filter = listService.FilterService as U8FilterService;
           
            U8Login.clsLogin login = (U8Login.clsLogin)listService.VbLogin;
            DbHelper.conStr = login.UFDataConnstringForNet;
            if (filter != null)
            {
                 string  userId = login.cUserId;

                 string cQx;
                 string sql = "select cSysUserName from UA_User where cSysUserName is not null and  cUser_Id='" + userId + "'";
                 DataTable dt = DbHelper.ExecuteTable(sql);
                 if (dt.Rows.Count > 0)
                 {
                     cQx = DbHelper.GetDbString(dt.Rows[0]["cSysUserName"]);
                 }
                 else
                 {
                     cQx = "0";

                 }


                 if (userId != "001" && userId != "demo"  && cQx != "1" && cQx != "2")
                {
                    //&& userId != "056"，20180302撤销
                    //<Item key='xzr' operator1='like' val1='%/" + canshu.userName + "/%' /></table></filteritems>
                    filter.DefaultFilterCondition = "<filter behidden=\"false\"><filteritems></filteritems><defaultcondition condition=\"#FN[xzr]  like '%/" + userId + "/%' \" /></filter>";
                    //Voucher
                }
            //filter.DefaultFilterCondition = "<filter behidden=\"false\"><filteritems></filteritems><defaultcondition condition=\"#FN[cNo] = N'123'\" /></filter>";
            }

        }

        /// <summary>
        /// 根据条件过滤源单据列表数据，可以重写缺省的过滤算法
        /// </summary>
        /// <param name="listService">列表服务</param>
        /// <param name="dataSet">查询后的数据(如果是分页的，只包含一页的数据)(该DataSet中只应该包含一个DataTable)</param>
        /// <param name="dataRowCountDataSet">查询到的数据总行数(该DataSet中只包含一个DataTable，并且DataTable中只有一行一列)</param>
        //public virtual void ReceiptListFilter(UFGeneralListService listService, SQLScript script, string queryBOid, bool needPaginate, int currentPageIndex, int pageSize, out DataSet dataSet, out DataSet dataRowCountDataSet)
        //{
        //    dataSet = null;
        //    dataRowCountDataSet = null;
        //}

        /// <summary>
        /// 根据条件过滤源单据列表数据，可以重写缺省的过滤算法
        /// </summary>
        /// <param name="filterArgs">过滤参数</param>
        //public override void ReceiptListFilter(FilterPluginArgs filterArgs)
        //{

        //}

        /// <summary>
        /// 过滤操作执行后的事件
        /// </summary>
        /// <param name="listService">列表服务</param>
        /// <param name="dataSet">查询后的数据(如果是分页的，只包含一页的数据)(该DataSet中只应该包含一个DataTable)</param>
        /// <param name="dataRowCountDataSet">查询到的数据总行数(该DataSet中只包含一个DataTable，并且DataTable中只有一行一列)</param>
        public override void ReceiptListFiltered(UFGeneralListService listService, DataSet dataSet, DataSet dataRowCountDataSet)
        {


        }

        /// <summary>
        /// 源单据列表过滤结果填充之前的处理
        /// </summary>
        /// <param name="listService">列表服务</param>
        /// <param name="dataSet">查询后的数据(如果是分页的，只包含一页的数据)(该DataSet中只应该包含一个DataTable)</param>
        //public override void ReceiptListFilling(UFGeneralListService listService, DataSet dataSet)
        //{

        
        //}

        ///// <summary>
        ///// 源单据列表过滤结果填充之后的处理
        ///// </summary>
        ///// <param name="listService">列表服务</param>
        ///// <param name="dataSet">查询后的数据(如果是分页的，只包含一页的数据)(该DataSet中只应该包含一个DataTable)</param>
        //public override void ReceiptListFilled(UFGeneralListService listService, DataSet dataSet)
        //{
        //    //通过父类访问该单据列表对应的列表服务变量，即该列表对应的模型
        //    UFGeneralListService service = base.ListService;

        //    //得到模型中被选中的数据行
        //    //List<object> selectedBusinee = service.GetSelectedListData();
        //    //selectedBusinee.
        //    ////获得选中数据行的所有主键值
        //    //IList<object> PKValues = this.GetPKValues(selectedBusinee);
        //    //if (PKValues.Count < 1)
        //    //{
        //    //    System.Windows.Forms.MessageBox.Show("请选择需要审核的单据", " 提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        //    //    return;
        //    //}
        //    ////通过模型中的U8登录对象得到在.Net中可以访问的登录对象
        //    //UFIDA.U8.UAP.Common.LoginInfo login = new UFIDA.U8.UAP.Common.LoginInfo(service.VbLogin);
        //    ////以下为根据登录对象获得相应数据库的事务，并在事务范围内对模型中的数据更新持久化到数据库中

        //}

        ///// <summary>
        ///// 源单据被选择事件，可以重写缺省的选择算法
        ///// </summary>
        ///// <param name="listService">列表服务</param>
        ///// <param name="sender">触发对象</param>
        ///// <param name="e">事件</param>
        //public override void ReceiptChecking(UFGeneralListService listService, object sender, CellEventArgs e)
        //{
        //}

        /// <summary>
        /// 源单据被选择事件，可以重写缺省的选择算法
        /// </summary>
        /// <param name="listService">列表服务</param>
        /// <param name="sender">触发对象</param>
        /// <param name="e">事件</param>
        /// <returns></returns>
        public override DataSet ReceiptCheck(UFGeneralListService listService, object sender, string voucherId)
        {
            return null;
        }

        /// <summary>
        /// 源单据被选择事件，可以重写缺省的选择算法
        /// </summary>
        /// <param name="listService">列表服务</param>
        /// <param name="sender">触发对象</param>
        /// <param name="e">事件</param>
        public override void ReceiptChecked(UFGeneralListService listService, object sender, CellEventArgs e)
        {

        }

    }
}
