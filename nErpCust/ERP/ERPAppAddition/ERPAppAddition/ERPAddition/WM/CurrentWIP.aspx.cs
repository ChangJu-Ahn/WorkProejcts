using System;
using System.Web.UI.WebControls;
using System.Web;

namespace ERPAppAddition.ERPAddition.WM
{
    public partial class CurrentWIP : System.Web.UI.Page
    {
        //public bool lBLoadFlag = false;
        #region "Functions"
        private void Initialize()
        {
            if (Request.QueryString["Menu"] == "false")
            {
                //SiteMapPath1.Visible = false;
                Table1.Visible = true;
            }
            if (Request.QueryString["LotID"] == "false")
            {
                tdLotID1.Visible = false;
                tdLotID2.Visible = false;
                tdCustLotID1.Visible = false;
                tdCustLotID2.Visible = false;
            }
        }
        private void SetDefaultValue()
        {
            try
            {
                GV.gOraCon.Open();
                GF.SetMatID2Ctrl(GV.gOraCon, mccMatID);
                GF.SetOper2Ctrl(GV.gOraCon, mccOper);
                GF.SetCreateCode2Ctrl(GV.gOraCon, mccCreateCode);
                GV.gOraCon.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Second exception caught.", e);
            }
        }


        private bool GetDefaultValue(ref string strMatID, ref string strOper, ref string strCreateCode, ref string strLotID, ref string strCustLotID, 
                                     ref string strMinDT, ref string strBackDT)
        {
            if (mccMatID.Count > 0) strMatID = "MAT_ID IN (" + mccMatID.SQLText.Trim() + ") \n";
            if (mccOper.Count > 0) strOper = "OPER IN (" + mccOper.SQLText.Trim() + ") \n";
            if (mccCreateCode.Count > 0) strCreateCode = "CREATE_CODE IN (" + mccCreateCode.SQLText.Trim() + ") \n";
            if (txtLotID.Text.Trim().Length > 0) strLotID = "LOT_ID LIKE '" + txtLotID.Text.Trim() + "%' \n";
            if (txtCustLotID.Text.Trim().Length > 0) strCustLotID = "LOT_CMF_2 LIKE '" + txtCustLotID.Text.Trim() + "%' \n";
            if (txtBackDT.Text.Trim().Length > 0)
            {
                DateTimeOffset dto = DateTimeOffset.Now;
                strBackDT = txtBackDT.Text.Trim();
                if (Convert.ToDouble(strBackDT) >= Convert.ToDouble(dto.ToString("yyyyMMddHHmmss")))
                    strBackDT = "";
                else
                    strMinDT = GF.AddDays(strBackDT, -365);
            }
            //.ToString("yyyyMdhmmss")
            return true;
        }

        private string MakeQuery(string strMatID, string strOper, string strCreateCode, string strLotID, string strCustLotID, string strMinDT, string strBackDT)
        {
            string strSQL;

            if (rdoViewType.SelectedValue == "B") // 공정별
                strSQL = "SELECT MAT_ID, GET_OPER_DESC(WIP.OPER) OPER_DESC, \n";
            else
                strSQL = "SELECT MAT_ID, LOT_DESC, \n"; //CASE WHEN WIP.CREATE_CODE = 'BNS' THEN 'B' ELSE 'N' END LOT_CATGRY, \n";

            //strSQL = strSQL + "         (SELECT SEQ_NUM FROM MWIPFLWOPR WHERE FACTORY = WIP.FACTORY AND FLOW = WIP.FLOW AND OPER = WIP.OPER) SEQ_NUM, \n";
            strSQL = strSQL + " OPER, CREATE_CODE, SUM(QTY_1) QTY_1 \n";
            if (strBackDT.Length == 0)
            {
                strSQL = strSQL + "  FROM MWIPLOTSTS WIP \n";
                strSQL = strSQL + " WHERE FACTORY = 'DISPLAY' AND LOT_DEL_FLAG = ' ' \n"; //  --AND OPER NOT LIKE 'R%' \n";
            }
            else
            {
                strSQL = strSQL + "  FROM MWIPLOTHIS WIP \n";
                strSQL = strSQL + " WHERE FACTORY  = 'DISPLAY' AND TRAN_TIME >= '" + strMinDT + "' AND TRAN_TIME < '" + strBackDT + "' \n"; //--AND TRAN_TIME >= TO_CHAR (TO_DATE ('" + strBackDT + "', 'YYYYMMDDHH24MISS') - 10000,  'YYYYMMDDHH24MISS') \n";
                strSQL = strSQL + "   AND (HIST_DEL_FLAG = ' ' OR (HIST_DEL_FLAG = 'Y' AND HIST_DEL_TIME >= '" + strBackDT + "')) AND QTY_1 > 0 AND LOT_DEL_FLAG = ' ' \n";
                //strSQL = strSQL + "   AND (TRAN_TIME, HIST_SEQ) IN (SELECT MAX(TRAN_TIME), MAX(HIST_SEQ) FROM MWIPLOTHIS WHERE LOT_ID = WIP.LOT_ID AND TRAN_TIME >= '" + strMinDT + "' AND TRAN_TIME < '" + strBackDT + "' AND FACTORY = 'DISPLAY' \n";
                //strSQL = strSQL + "                                 AND (HIST_DEL_FLAG = ' ' OR (HIST_DEL_FLAG = 'Y' AND HIST_DEL_TIME >= '" + strBackDT + "')) AND QTY_1 > 0 AND LOT_DEL_FLAG = ' ') \n";
                strSQL = strSQL + "   AND (LOT_ID, HIST_SEQ) IN (SELECT LOT_ID, MAX(HIST_SEQ) FROM MWIPLOTHIS WHERE LOT_ID = WIP.LOT_ID AND TRAN_TIME >= '" + strMinDT + "' AND TRAN_TIME < '" + strBackDT + "' AND FACTORY = 'DISPLAY' \n";
                strSQL = strSQL + "                                 AND (HIST_DEL_FLAG = ' ' OR (HIST_DEL_FLAG = 'Y' AND HIST_DEL_TIME >= '" + strBackDT + "')) AND QTY_1 > 0 AND LOT_DEL_FLAG = ' ' GROUP BY LOT_ID) \n";
                strSQL = strSQL + "     AND LOT_ID NOT IN (SELECT LOT_ID FROM MWIPLOTSHP WHERE LOT_ID = WIP.LOT_ID AND TRAN_TIME >= '" + strMinDT + "' AND TRAN_TIME < '" + strBackDT + "' AND HIST_DEL_FLAG = ' ') \n";
                strSQL = strSQL + "     AND LOT_ID NOT IN (SELECT LOT_ID FROM MWIPLOTSTS WHERE LOT_ID = WIP.LOT_ID AND LOT_DEL_FLAG = 'Y' AND LAST_TRAN_TIME >= '" + strMinDT + "' AND LAST_TRAN_TIME < '" + strBackDT + "') \n";
            }

            if (strMatID.Length > 0) strSQL = strSQL + " AND " + strMatID;
            if (strOper.Length > 0) strSQL = strSQL + " AND " + strOper;
            if (strCreateCode.Length > 0) strSQL = strSQL + " AND " + strCreateCode;
            if (strLotID.Length > 0) strSQL = strSQL + " AND " + strLotID;
            if (strCustLotID.Length > 0) strSQL = strSQL + " AND " + strCustLotID;
            //if (chkRetrieved.Checked == false) strSQL = strSQL + " AND LOT_CMF_3 = ' ' ";
            if (chkHolded.Checked == false) strSQL = strSQL + " AND HOLD_FLAG = ' ' ";

            if (rdoViewType.SelectedValue == "B")
                strSQL = strSQL + " GROUP BY MAT_ID, OPER, CREATE_CODE \n";
            else
                strSQL = strSQL + " GROUP BY MAT_ID, OPER, LOT_DESC, CREATE_CODE \n";
            //if (multichk_yn.Checked || strMatID == "%")
            //    strSQL = strSQL + "   ORDER BY 3,1 \n";
            //else
            strSQL = strSQL + " ORDER BY OPER \n";

            //string strTranTime = "20140508070000";
            //strSQL = "   SELECT MAT_ID, (SELECT OPER_DESC FROM MWIPOPRDEF WHERE FACTORY = 'DISPLAY' AND OPER = WIP.OPER) OPER_DESC, OPER, SUM(QTY_1) QTY_1 \n";
            //strSQL = strSQL + "        FROM MWIPLOTHIS WIP \n";
            //strSQL = strSQL + "       WHERE FACTORY  = 'DISPLAY'  AND HOLD_FLAG = ' ' AND OPER_IN_TIME < '" + strBackDT + "' --AND OPER_IN_TIME >= TO_CHAR (TO_DATE ('" + strBackDT + "', 'YYYYMMDDHH24MISS') - 10000,  'YYYYMMDDHH24MISS') \n";
            //strSQL = strSQL + "         AND (TRAN_TIME, HIST_SEQ) IN (SELECT MAX(TRAN_TIME), MAX(HIST_SEQ) FROM MWIPLOTHIS WHERE LOT_ID = WIP.LOT_ID AND FACTORY = WIP.FACTORY AND MAT_ID = WIP.MAT_ID AND TRAN_TIME < '" + strBackDT + "' AND (HIST_DEL_FLAG = ' ' OR HIST_DEL_TIME > '" + strBackDT + "')) \n";
            //strSQL = strSQL + "         AND LOT_ID NOT IN (SELECT LOT_ID FROM MWIPLOTSHP WHERE LOT_ID = WIP.LOT_ID AND TRAN_TIME < '" + strBackDT + "' AND HIST_DEL_FLAG = ' ') \n";
            //strSQL = strSQL + "         AND LOT_ID NOT IN (SELECT LOT_ID FROM MWIPLOTSTS WHERE LOT_ID = WIP.LOT_ID AND FACTORY = WIP.FACTORY AND  LOT_DEL_FLAG = 'Y' AND LAST_TRAN_TIME < '" + strBackDT + "') \n";
            //strSQL = strSQL + "         AND LOT_ID NOT IN (SELECT LOT_ID FROM MWIPLOTSTS WHERE LOT_ID = WIP.LOT_ID AND FACTORY = WIP.FACTORY AND MAT_ID = WIP.MAT_ID AND LAST_TRAN_CODE IN ('TERMINATE') AND LOT_DEL_FLAG = 'Y') \n";
            //strSQL = strSQL + "    GROUP BY MAT_ID, OPER ";

            return strSQL;
        }
        #endregion

        #region "Events"
        //protected void InitCreateCode(object sender, EventArgs e)
        //{
        //    mccCreateCode.lStrArrText = new string[2] { "ENG", "PROD" };
        //    //mccCreateCode.InitDefault(arrStrItems);
        //}
        //protected void Page_PreInit(Object sender, EventArgs e)
        //{
        //    if (Request.QueryString["Menu"] == "false") this.MasterPageFile = "../Site2.Master";
            
        //}
        protected void Page_Load(object sender, EventArgs e)
        {
            Initialize();
            SetDefaultValue();
            WebSiteCount();
            //if (GV.gStrPageTitle != Page.Title)
            //if (!Page.IsPostBack)
            //{

            //    string[] arrStrItems = new string[2] { "ENG", "PROD" };
            //    mccCreateCode.InitDefault(arrStrItems);
            //    //GV.gStrPageTitle = Page.Title;
            //}
        }
        protected void query_Click(object sender, EventArgs e)
        {
            string strSQL = "", strMatID = "", strOper = "", strCreateCode = "", strLotID = "", strCustLotID = "", strMinDT = "", strBackDT = "";

            if (GetDefaultValue(ref strMatID, ref strOper, ref strCreateCode, ref strLotID, ref strCustLotID, ref strMinDT, ref strBackDT) == false) return;
            if ((strSQL = MakeQuery(strMatID, strOper, strCreateCode, strLotID, strCustLotID, strMinDT, strBackDT).Trim()) == "") return;
            if (strSQL.Length < 1 || strSQL == "") return;

            WIPStockModule dtWIP = new WIPStockModule();
            if (rdoViewType.SelectedValue == "B")
                GF.CreateReport(dtWIP, 0, strSQL, ReportViewer1, "WM.CurrentWIP.rdlc", "dsCurrentWIP", "REPORT_" + this.Title + DateTime.Now.ToString());
            else
                GF.CreateReport(dtWIP, 0, strSQL, ReportViewer1, "WM.CurrentWIPByType.rdlc", "dsCurrentWIP", "REPORT_" + this.Title + DateTime.Now.ToString());
        }
        #endregion

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = "NEPES&NEPES_DISPLAY";
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }
    }
}