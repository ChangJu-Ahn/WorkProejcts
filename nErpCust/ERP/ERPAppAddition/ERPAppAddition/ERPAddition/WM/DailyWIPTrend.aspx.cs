using System;
using System.Web.UI.WebControls;
using System.Web;

namespace ERPAppAddition.ERPAddition.WM
{
    public partial class DailyWIPTrend : System.Web.UI.Page
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
            Init_ChkMore();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = "NEPES&NEPES_DISPLAY";
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void SetDefaultValue()
        {
            try
            {
                GC.ProdType ptValue = GC.ProdType.ALL;
                switch (rdoProdType.SelectedValue)
                {
                    case "DDI": ptValue = GC.ProdType.DDI; break;
                    case "WLP": ptValue = GC.ProdType.WLP; break;
                    case "FOWLP": ptValue = GC.ProdType.FOWLP; break;
                }

                GV.gOraCo2.Open();
                GF.SetPartID2CtrlEx(GV.gOraCo2, mccPartID, ptValue);
                GF.SetOper2CtrlEx(GV.gOraCo2, mccOper, ptValue);
                GF.SetCreateCode2CtrlEx(GV.gOraCo2, mccCreateCode, ptValue);
                GV.gOraCo2.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Second exception caught.", e);
            }
        }
        private bool GetDefaultValue(ref string strMatID, ref string strOper, ref string strCreateCode, ref string strFromDT, ref string strToDT, ref string strStdTM)
        {
            if (txtFromDT.Text.Trim().Length == 0) return false;
            if (txtToDT.Text.Trim().Length == 0) return false;
            strFromDT = txtFromDT.Text.Replace("-", "");
            strToDT = txtToDT.Text.Replace("-", "");
            strStdTM = ddlStdTM.SelectedValue;

            DateTimeOffset dto = DateTimeOffset.Now;
            if (Convert.ToDouble(strToDT) >= Convert.ToDouble(dto.ToString("yyyyMMdd")))
            {
                if (Convert.ToDecimal(strStdTM) > Convert.ToDecimal(dto.ToString("HH")))
                    strToDT = dto.AddDays(-1).ToString("yyyyMMdd");
                else
                    strToDT = dto.ToString("yyyyMMdd");
            }

            if (mccPartID.Count > 0) strMatID = "PART IN (" + mccPartID.SQLText.Trim() + ") \n";
            if (mccOper.Count > 0) strOper = "OPERATION IN (" + mccOper.SQLText.Trim() + ") \n";
            if (mccCreateCode.Count > 0) strCreateCode = "CREATE_CODE IN (" + mccCreateCode.SQLText.Trim() + ") \n";

            return true;
        }
        private string MakeQuery(string strMatID, string strOper, string strCreateCode, string strFromDT, string strToDT, string strStdTM)
        {
            string strSQL, strSQLCond0 = "", strSQLCond1 = "", strSQLCond2 = "", strSQLCond3 = "", strProdType = rdoProdType.SelectedValue; ;

            strSQL = "WITH WIP AS ( \n";
            strSQL = strSQL + "SELECT SUBSTR(POINT_TIME, 0, 8) WORK_DATE, OPERATION OPER, QTY_UNIT_1, CREATE_CODE C_CODE, \n";
            switch (strProdType)
            {
                case "ALL":
                    strSQL = strSQL + "       CASE WHEN OPERATION IN ('402T', '4050', '4100') THEN DECODE(INSTR(ROUTESET, 'PR'), 0, SUM(QTY_1)) ELSE SUM(QTY_1) END QTY_1 \n";
                    strSQLCond0 = " AND STATUS <> 99 AND PART_TYPE IN ('D', 'P') \n";
                    strSQLCond0 = strSQLCond0 + " AND NOT (CREATE_CODE = 'TCP' AND OPERATION = '9000' AND CUSTOMER IN ('HIMAX', 'HIMAX_SEMI', 'VALIDITY')) \n";
                    strSQLCond0 = strSQLCond0 + " AND ROUTESET NOT IN ('PR(NI)-T', 'PR(P2NI)-T') \n";
                    strSQLCond0 = strSQLCond0 + " AND CREATE_CODE <> 'RCP-S' -- AND CREATE_CODE LIKE 'RCP%' \n";
                    strSQLCond1 = "          WHERE W.OPER NOT IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90', 'G9000', 'I8500', 'I9000', 'HE900', 'J9000') \n";
                    strSQLCond2 = "          WHERE W.OPER IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90', 'G9000', 'I8500', 'I9000', 'HE900', 'J9000') \n";
                    //strSQLCond3 = "AND OUC.CREATE_CODE = DECODE(OUC.PLANT, 'P02', W.CC_GROUP, 'ALL')";
                    break;
                case "DDI":
                case "WLP":
                    strSQL = strSQL + "       CASE WHEN OPERATION IN ('402T', '4050', '4100') THEN DECODE(INSTR(ROUTESET, 'PR'), 0, SUM(QTY_1)) ELSE SUM(QTY_1) END QTY_1 \n"; //, SUM(QTY_1) QTY_1 \n";
                    //strCCGroup = "DECODE(SUBSTR(CREATE_CODE, 0, 3), 'WSL', 'WLP', 'WLC', 'WLP', 'DDI') CC_GROUP";
                    //strSQL = strSQL + strCCGroup + ", \n";
                    strSQLCond0 = " AND STATUS <> 99 AND PART_TYPE IN ('D', 'P') \n";
                    if (strProdType == "DDI")
                        strSQLCond0 = strSQLCond0 + "   AND NOT (CREATE_CODE = 'TCP' AND OPERATION = '9000' AND CUSTOMER IN ('HIMAX', 'HIMAX_SEMI', 'VALIDITY')) \n";
                    else
                        strSQLCond0 = strSQLCond0 + "   AND ROUTESET NOT IN ('PR(NI)-T', 'PR(P2NI)-T') \n";
                    strSQLCond1 = "          WHERE W.OPER NOT IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90') \n";
                    strSQLCond2 = "          WHERE W.OPER IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90') \n";
                    //strSQL = strSQL + "  FROM FWIP F LEFT JOIN OPRUNTCST OUC ON OUC.PLANT IN ('P01', 'P02', 'P09') AND F.CREATE_CODE = OUC.CREATE_CODE AND F.OPER = OUC.OPER --BETWEEN YYYYMMDD AND EXPIRY_DATE \n";
                    strSQLCond3 = "AND OUC.PLANT IN ('P01', 'P02', 'P09')";
                    break;
                //case "WLP":
                //    //strCCGroup = "DECODE(SUBSTR(CREATE_CODE, 0, 3), 'WLC', 'WLP', 'DDI') CC_GROUP";
                //    //strSQL = strSQL + strCCGroup + ", SUM(QTY_1) QTY_1 \n";
                //    strSQLCond0 = " AND PART_TYPE IN ('D', 'P') AND QTY_UNIT_1 = 'PCS' \n";
                //    strSQLCond1 = "          WHERE W.OPER NOT IN ('8900', '9000') \n";
                //    strSQLCond2 = "          WHERE W.OPER IN ('8900', '9000') \n";
                //    strSQLCond3 = "AND OUC.PLANT = 'P02' AND OUC.CREATE_CODE = W.CC_GROUP";
                //    break;
                //case "P09":
                //    //strCCGroup = "DECODE(SUBSTR(CREATE_CODE, 0, 3), 'WSL', 'WLP', 'DDI') CC_GROUP";
                //    //strSQL = strSQL + strCCGroup + ", SUM(QTY_1) QTY_1 \n";
                //    strSQLCond0 = " AND PART_TYPE IN ('D', 'P') AND PROD_TYPE ='C' AND QTY_UNIT_1 = 'SLS' \n";
                //    strSQLCond1 = "          WHERE W.OPER NOT IN ('FS40', 'FS90') \n";
                //    strSQLCond2 = "          WHERE W.OPER IN ('FS40', 'FS90') \n";
                //    strSQLCond3 = "AND OUC.PLANT = 'P09'";
                //    break;
                case "FOWLP":
                    //strCCGroup = "'FOWLP' CC_GROUP";
                    //strSQL = strSQL + strCCGroup + ", SUM(QTY_1) QTY_1 \n";
                    strSQL = strSQL + "       SUM(QTY_1) QTY_1 \n";
                    strSQLCond0 = " AND STATUS <> 99 AND PART_TYPE IN ('D', 'P') AND CREATE_CODE LIKE 'RCP%' AND CREATE_CODE <> 'RCP-S' \n";
                    strSQLCond1 = "          WHERE W.OPER NOT IN ('G9000', 'I8500', 'I9000', 'HE900', 'J9000') \n";
                    strSQLCond2 = "          WHERE W.OPER IN ('G9000', 'I8500', 'I9000', 'HE900', 'J9000') \n";
                    strSQLCond3 = "AND OUC.PLANT = 'P12'";
                    break;
            }

            strSQL = strSQL + "  FROM POINTWIP RW  \n";
            strSQL = strSQL + "  LEFT JOIN SYSCODEDATA SC ON RW.ACCOUNT_CODE = SC.SYSCODE_NAME AND SC.PLANT = RW.PLANT AND SYSTABLE_NAME = 'SCRAP_REASON' AND SYSCODE_GROUP = 'FREE' \n";
            strSQL = strSQL + " WHERE RW.PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND PART_TYPE IN ('D', 'P') AND SC.SYSCODE_NAME IS NULL \n";
            strSQL = strSQL + "   AND POINT_TIME BETWEEN '" + strFromDT + strStdTM + "' AND '" + strToDT + strStdTM + "' AND SUBSTR(POINT_TIME, 9, 2) = '" + strStdTM + "' \n";
            strSQL = strSQL + strSQLCond0;

            if (strMatID.Length > 0) strSQL = strSQL + " AND " + strMatID;
            if (strOper.Length > 0) strSQL = strSQL + " AND " + strOper;
            if (strCreateCode.Length > 0) strSQL = strSQL + " AND " + strCreateCode;
            if (chkHolded.Checked == false) strSQL = strSQL + " AND HOLD = 'N' ";

            strSQL = strSQL + " GROUP BY SUBSTR(POINT_TIME, 0, 8), OPERATION, CREATE_CODE, ROUTESET, QTY_UNIT_1 \n";
            strSQL = strSQL + "), FWIP AS ( \n";
            strSQL = strSQL + "SELECT OUI.*, W.* \n";
            strSQL = strSQL + "  FROM WIP W INNER JOIN OPRUNTINF OUI ON W.OPER BETWEEN OUI.OPER_IN AND OUI.OPER_OUT";
            if (strProdType == "ALL")
            {
                strSQL = strSQL + " AND W.C_CODE LIKE CASE WHEN OUI.PLANT = 'P12' THEN '%' ELSE OUI.GROUP5 END \n";
            }
            else
            {
                strSQL = strSQL + " AND GROUP1 = '" + strProdType + "'";
                if (strProdType != "FOWLP") strSQL = strSQL + " AND W.C_CODE = OUI.GROUP5 \n";
            }
            strSQL = strSQL + "), WIPINV AS ( \n";
            strSQL = strSQL + "SELECT WORK_DATE, W.GROUP1, SUM(W.QTY_1) WIP_QTY, SUM(W.QTY_1 * OUC.UNIT_COST) WIP_COST, 0 INV_QTY, 0 INV_COST  \n";
            strSQL = strSQL + "  FROM FWIP W LEFT JOIN OPRUNTCST OUC ON W.OPER = OUC.OPER " + strSQLCond3 + " AND W.CREATE_CODE = OUC.CREATE_CODE --AND W.WORK_DATE BETWEEN YYYYMMDD AND EXPIRY_DATE \n";
            strSQL = strSQL + strSQLCond1;
            strSQL = strSQL + " GROUP BY WORK_DATE, W.GROUP1 \n";
            strSQL = strSQL + " UNION ALL \n";
            strSQL = strSQL + "SELECT WORK_DATE, W.GROUP1, 0, 0, SUM(W.QTY_1), SUM(W.QTY_1 * OUC.UNIT_COST) \n";
            strSQL = strSQL + "  FROM FWIP W LEFT JOIN OPRUNTCST OUC ON W.OPER = OUC.OPER " + strSQLCond3 + " AND W.CREATE_CODE = OUC.CREATE_CODE --AND W.WORK_DATE BETWEEN YYYYMMDD AND EXPIRY_DATE \n";
            if (strProdType == "ALL")
                //strSQL = strSQL + "                                                 AND OUC.PLANT LIKE DECODE(W.QTY_UNIT_1, 'PCS', 'P02', 'P0%') \n";
                strSQL = strSQL + "                                                 AND OUC.UNIT =  W.QTY_UNIT_1 \n";
            strSQL = strSQL + strSQLCond2;
            strSQL = strSQL + " GROUP BY WORK_DATE, W.GROUP1) \n";
            strSQL = strSQL + "SELECT DECODE(WORK_DATE, '" + strToDT + "', 'L', 'N') RPT_TYPE, TO_DATE(WORK_DATE) WORK_DATE, GROUP1, SUM(WIP_QTY) WIP_QTY, NVL(SUM(WIP_COST), 0) WIP_COST, SUM(INV_QTY) INV_QTY, SUM(INV_COST) INV_COST \n";
            strSQL = strSQL + "  FROM WIPINV \n";
            strSQL = strSQL + " GROUP BY WORK_DATE, GROUP1 \n";
            strSQL = strSQL + " ORDER BY WORK_DATE, GROUP1 \n";

            return strSQL;
        }
        private void Init_ChkMore()
        {
            if (chkMore.Checked)
            {
                trInquiry2.Visible = true;
                rdoProdType.Items[0].Enabled = false;
                rdoProdType.AutoPostBack = true;
                if (rdoProdType.SelectedIndex == 0)
                {
                    rdoProdType.SelectedIndex = 1;
                    ReportViewer1.Reset();
                }
            }
            else
            {
                trInquiry2.Visible = false;
                rdoProdType.Items[0].Enabled = true;
                rdoProdType.AutoPostBack = false;
            }
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
        protected void InitTxtWorkDate(object sender, EventArgs e)
        {
            TextBox txtTemp = (TextBox)sender;
            if (txtTemp.ID == "txtFromDT")
                txtTemp.Text = DateTime.Now.ToString("yyyy-MM-01");
            else
                txtTemp.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            Initialize();
            WebSiteCount(); 
            if (chkMore.Checked) SetDefaultValue();
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
            string strSQL = "", strMatID = "", strOper = "", strCreateCode = "", strFromDT = "", strToDT = "", strStdTM = "";

            if (GetDefaultValue(ref strMatID, ref strOper, ref strCreateCode, ref strFromDT, ref strToDT, ref strStdTM) == false) return;
            if ((strSQL = MakeQuery(strMatID, strOper, strCreateCode, strFromDT, strToDT, strStdTM).Trim()) == "") return;
            if (strSQL.Length < 1 || strSQL == "") return;

            WIPStockModule dtWIP = new WIPStockModule();
            GF.CreateReport2(dtWIP, 3, strSQL, ReportViewer1, "WM.DailyWIPTrend.rdlc", "dsDailyWIPTrend", "REPORT_" + this.Title + DateTime.Now.ToString());
        }
        protected void rdoProdType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
        }
        protected void chkHolded_CheckedChanged(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
        }
        //protected void chkMore_CheckedChanged(object sender, EventArgs e)
        //{
        //    if(chkMore.Checked && rdoProdType.SelectedIndex == 0) ReportViewer1.Reset();
        //}
        #endregion
    }
}