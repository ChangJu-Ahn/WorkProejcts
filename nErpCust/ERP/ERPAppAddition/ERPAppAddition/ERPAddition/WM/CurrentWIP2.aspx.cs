using System;
using System.Web.UI.WebControls;
using System.Web;

namespace ERPAppAddition.ERPAddition.WM
{
    public partial class CurrentWIP2 : System.Web.UI.Page
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
                GC.ProdType ptValue = GC.ProdType.DDI;
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
        private bool GetDefaultValue(ref string strMatID, ref string strOper, ref string strCreateCode, ref string strLotID, ref string strCustLotID, ref string strBackDT, ref string strCrntDT)
        {
            if (mccPartID.Count > 0) strMatID = "PART IN (" + mccPartID.SQLText.Trim() + ") \n";
            if (mccOper.Count > 0) strOper = "OPERATION IN (" + mccOper.SQLText.Trim() + ") \n";
            if (mccCreateCode.Count > 0) strCreateCode = "CREATE_CODE IN (" + mccCreateCode.SQLText.Trim() + ") \n";
            if (txtLotID.Text.Trim().Length > 0) strLotID = "LOT_NUMBER LIKE '" + txtLotID.Text.Trim() + "%' \n";
            if (txtCustLotID.Text.Trim().Length > 0) strCustLotID = "UPLEVEL_LOT LIKE '" + txtCustLotID.Text.Trim() + "%' \n";

            DateTimeOffset dto = DateTimeOffset.Now;
            if (txtBackDT.Text.Trim().Length > 0)
            {
                strBackDT = txtBackDT.Text.Trim() + ddlBackTM.SelectedValue;
                if (Convert.ToDouble(strBackDT) >= Convert.ToDouble(dto.ToString("yyyyMMddHH")))
                    strBackDT = "";
                strCrntDT = strBackDT.Substring(0, 8);
            }
            else
                strCrntDT = dto.ToString("yyyyMMddHHmmss");

            //.ToString("yyyyMdhmmss")
            return true;
        }
        private string MakeQuery(string strMatID, string strOper, string strCreateCode, string strLotID, string strCustLotID, string strBackDT, string strCrntDT)
        {
            string strSQL, strSQLWithFrom = "", strProdType = rdoProdType.SelectedValue;

            if (strBackDT.Length == 0)
            {
                strSQLWithFrom = " FROM REPORTWIP RW \n";
                strSQLWithFrom = strSQLWithFrom + "  LEFT JOIN SYSCODEDATA SC ON RW.ACCOUNT_CODE = SC.SYSCODE_NAME AND SC.PLANT = RW.PLANT AND SYSTABLE_NAME = 'SCRAP_REASON' AND SYSCODE_GROUP = 'FREE' \n";
                strSQLWithFrom = strSQLWithFrom + " WHERE RW.PLANT = 'CCUBEDIGITAL' AND STATUS <> 99  AND PART_TYPE IN ('D', 'P') AND SC.SYSCODE_NAME IS NULL \n";
            }
            else
            {
                strSQLWithFrom = " FROM POINTWIP RW \n";
                strSQLWithFrom = strSQLWithFrom + "  LEFT JOIN SYSCODEDATA SC ON RW.ACCOUNT_CODE = SC.SYSCODE_NAME AND SC.PLANT = RW.PLANT AND SYSTABLE_NAME = 'SCRAP_REASON' AND SYSCODE_GROUP = 'FREE' \n";
                strSQLWithFrom = strSQLWithFrom + " WHERE RW.PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND PART_TYPE IN ('D', 'P') AND SC.SYSCODE_NAME IS NULL AND POINT_TIME = '" + strBackDT + "' \n";
            }

            if (strMatID.Length > 0) strSQLWithFrom = strSQLWithFrom + " AND " + strMatID;
            if (strOper.Length > 0) strSQLWithFrom = strSQLWithFrom + " AND " + strOper;
            if (strCreateCode.Length > 0) strSQLWithFrom = strSQLWithFrom + " AND " + strCreateCode;
            if (strLotID.Length > 0) strSQLWithFrom = strSQLWithFrom + " AND " + strLotID;
            if (strCustLotID.Length > 0) strSQLWithFrom = strSQLWithFrom + " AND " + strCustLotID;
            if (chkHolded.Checked == false) strSQLWithFrom = strSQLWithFrom + " AND HOLD = 'N' ";

            strSQL = "WITH WIP AS (\n";
            strSQL = strSQL + "SELECT PART MAT_ID, OPERATION OPER,  \n";

            strSQL = strSQL + "CASE WHEN CREATE_CODE = 'WSLD'  AND ROUTESET LIKE '%PR%' THEN 'WSLD-PR'  \n";
            strSQL = strSQL + "     WHEN ROUTESET LIKE '%WLP%' AND ROUTESET LIKE '%PR%' THEN 'WSLD-PR'  \n";
            strSQL = strSQL + "     ELSE CREATE_CODE END C_CODE, \n";
            strSQL = strSQL + "CUSTOMER, \n";

            switch (strProdType)
            {
                case "DDI":
                case "WLP":
                    strSQL = strSQL + "      SUM(QTY_1) AS QTY_1 \n"; //, SUM(QTY_1) QTY_1 \n";
                    if (strProdType == "DDI")
                        strSQLWithFrom = strSQLWithFrom + "   AND NOT (CREATE_CODE = 'TCP' AND OPERATION = '9000' AND CUSTOMER IN ('HIMAX', 'HIMAX_SEMI', 'VALIDITY')) \n";
                    else
                        strSQLWithFrom = strSQLWithFrom + "   AND ROUTESET NOT IN ('PR(NI)-T', 'PR(P2NI)-T') \n";
                    strSQL = strSQL + strSQLWithFrom + " GROUP BY PART, OPERATION, CREATE_CODE, ROUTESET, CUSTOMER \n";
                    strSQL = strSQL + "), FWIP AS ( \n";
                    strSQL = strSQL + "SELECT OUI.*, W.* \n";
                    strSQL = strSQL + "  FROM WIP W INNER JOIN OPRUNTINF OUI ON CASE WHEN W.C_CODE = 'WSLD' AND W.CUSTOMER IN('SILICON MITUS','DONGJINTECH','DONGBU HITEK','MAGNACHIP') THEN 'WSLD-CUS'  \n";
                    strSQL = strSQL + "                                              ELSE W.C_CODE \n";
                    strSQL = strSQL + "                                               END = OUI.GROUP5  \n";
                    strSQL = strSQL + "                                               AND GROUP1 = '" + strProdType + "' AND W.OPER BETWEEN OUI.OPER_IN AND OUI.OPER_OUT) \n";
                    strSQL = strSQL + "SELECT 'A' RPT_TYPE, F.SEQ, F.GROUP1, F.GROUP2, F.GROUP3, F.GROUP4, F.GROUP5 CREATE_CODE, '' MAT_ID, SUM(F.QTY_1) QTY_1, NVL(OUC.UNIT_COST, 0) UNIT_COST, \n";
                    strSQL = strSQL + "       F.OPER, (SELECT LONG_DESC FROM OPERATION@CCUBE WHERE OPERATION = F.OPER AND PLANT = 'CCUBEDIGITAL') OPER_DESC, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER NOT IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(F.QTY_1) ELSE 0 END WIP_QTY, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER NOT IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(F.QTY_1) * OUC.UNIT_COST ELSE 0 END WIP_COST, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(F.QTY_1) ELSE 0 END INV_QTY, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER IN ('4020', '5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(F.QTY_1) * OUC.UNIT_COST ELSE 0 END INV_COST, OUC.UNIT \n";
                    strSQL = strSQL + "  FROM FWIP F LEFT JOIN OPRUNTCST OUC ON OUC.PLANT IN ('P01', 'P02', 'P09') AND F.CREATE_CODE = OUC.CREATE_CODE AND F.OPER = OUC.OPER --BETWEEN YYYYMMDD AND EXPIRY_DATE \n";
                    strSQL = strSQL + " GROUP BY F.GROUP1, F.GROUP2, F.GROUP3, F.GROUP4, F.GROUP5, F.SEQ, F.OPER, OUC.UNIT_COST, OUC.UNIT \n";
                    break;

                //    strSQLSelect = " CASE WHEN W.OPER NOT IN ('5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(W.QTY_1) ELSE 0 END, CASE WHEN W.OPER NOT IN ('5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(W.QTY_1) * OUC.UNIT_COST ELSE 0 END, \n";
                //    strSQLSelect = strSQLSelect + " CASE WHEN W.OPER IN ('5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(W.QTY_1) ELSE 0 END, CASE WHEN W.OPER IN ('5000', '8900', '9000', 'FS40', 'FS90') THEN SUM(W.QTY_1) * OUC.UNIT_COST ELSE 0 END \n";
                //    strSQLCond = "AND OUC.PLANT = 'P01'";
                //    break;
                //case "WLP":
                //    strSQL = strSQL + "       SUM(QTY_1) QTY_1 \n";
                //    strSQL = strSQL + strSQLWithFrom;

                //    strSQLSelect = " CASE WHEN W.OPER NOT IN ('8900', '9000') THEN SUM(W.QTY_1) ELSE 0 END, CASE WHEN W.OPER NOT IN ('8900', '9000') THEN SUM(W.QTY_1) * OUC.UNIT_COST ELSE 0 END, \n";
                //    strSQLSelect = strSQLSelect + " CASE WHEN W.OPER IN ('8900', '9000') THEN SUM(W.QTY_1) ELSE 0 END, CASE WHEN W.OPER IN ('8900', '9000') THEN SUM(W.QTY_1) * OUC.UNIT_COST ELSE 0 END \n";
                //    strSQLCond = "AND OUC.PLANT = 'P02' AND OUC.CREATE_CODE = DECODE(W.CREATE_CODE, 'WLCSP', 'WLP', 'DDI')";
                //    break;
                //case "FOWLP":
                //    strSQL = strSQL + "       SUM(QTY_1) QTY_1 \n";
                //    strSQL = strSQL + strSQLWithFrom;

                //    strSQLSelect = " CASE WHEN W.OPER NOT IN ('FS40', 'FS90') THEN SUM(W.QTY_1) ELSE 0 END, CASE WHEN W.OPER NOT IN ('FS40', 'FS90') THEN SUM(W.QTY_1) * OUC.UNIT_COST ELSE 0 END, \n";
                //    strSQLSelect = strSQLSelect + " CASE WHEN W.OPER IN ('FS40', 'FS90') THEN SUM(W.QTY_1) ELSE 0 END, CASE WHEN W.OPER IN ('FS40', 'FS90') THEN SUM(W.QTY_1) * OUC.UNIT_COST ELSE 0 END \n";
                //    strSQLCond = "AND OUC.PLANT = 'P09'";
                //    break;
                case "FOWLP":
                    strSQL = strSQL + "       SUM(QTY_1) QTY_1 \n";
                    strSQL = strSQL + strSQLWithFrom + " AND CREATE_CODE LIKE 'RCP%' AND CREATE_CODE <> 'RCP-S' \n";
                    strSQL = strSQL + " GROUP BY PART, OPERATION, CREATE_CODE, ROUTESET,CUSTOMER \n";
                    strSQL = strSQL + "), FWIP AS ( \n";
                    strSQL = strSQL + "SELECT OUI.*, W.* \n";
                    strSQL = strSQL + "  FROM WIP W INNER JOIN OPRUNTINF OUI ON GROUP1 = '" + strProdType + "' AND W.OPER BETWEEN OUI.OPER_IN AND OUI.OPER_OUT) \n";
                    strSQL = strSQL + "SELECT 'A' RPT_TYPE, F.SEQ, F.GROUP1, F.GROUP2, F.GROUP3, F.GROUP4, F.GROUP5 CREATE_CODE, '' MAT_ID, SUM(F.QTY_1) QTY_1, NVL(OUC.UNIT_COST, 0) UNIT_COST, \n";
                    strSQL = strSQL + "       F.OPER, (SELECT LONG_DESC FROM OPERATION@CCUBE WHERE OPERATION = F.OPER AND PLANT = 'CCUBEDIGITAL') OPER_DESC, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER NOT IN ('G9000', 'I8500', 'I9000', 'HE900', 'J9000') THEN SUM(F.QTY_1) ELSE 0 END WIP_QTY, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER NOT IN ('G9000', 'I8500', 'I9000', 'HE900', 'J9000') THEN SUM(F.QTY_1) * OUC.UNIT_COST ELSE 0 END WIP_COST, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER IN ('G9000', 'I8500', 'I9000', 'HE900', 'J9000') THEN SUM(F.QTY_1) ELSE 0 END INV_QTY, \n";
                    strSQL = strSQL + "       CASE WHEN F.OPER IN ('G9000', 'I8500', 'I9000', 'HE900', 'J9000') THEN SUM(F.QTY_1) * OUC.UNIT_COST ELSE 0 END INV_COST, OUC.UNIT \n";
                    strSQL = strSQL + "  FROM FWIP F LEFT JOIN OPRUNTCST OUC ON OUC.PLANT = 'P12' AND F.CREATE_CODE = OUC.CREATE_CODE AND F.OPER = OUC.OPER --BETWEEN YYYYMMDD AND EXPIRY_DATE \n";
                    strSQL = strSQL + " GROUP BY F.GROUP1, F.GROUP2, F.GROUP3, F.GROUP4, F.GROUP5, F.SEQ, F.OPER, OUC.UNIT_COST, OUC.UNIT \n";
                    break;
            }

            //strSQL = strSQL + "GROUP BY PART, OPERATION, CREATE_CODE, ROUTESET) \n";
            strSQL = strSQL + " UNION ALL \n";
            strSQL = strSQL + "SELECT 'B', 0, '', '', '', '', F.GROUP5, MAT_ID, QTY_1, 0 UNIT_COST, F.OPER, \n";
            strSQL = strSQL + "       (SELECT LONG_DESC FROM OPERATION@CCUBE WHERE OPERATION = F.OPER AND PLANT = 'CCUBEDIGITAL') OPER_DESC, \n";
            strSQL = strSQL + "       0 WIP_QTY, 0 WIP_COST, 0 INV_QTY, 0 INV_COST, '' \n";
            strSQL = strSQL + "  FROM FWIP F \n";
            strSQL = strSQL + " ORDER BY 1, 2, OPER \n";
            //strSQL = strSQL + "SELECT 'B', '', W.OPER_DESC, W.OPER, OUC.CREATE_CODE, SUM(W.QTY_1), NVL(OUC.UNIT_COST, 0), \n";
            //strSQL = strSQL + strSQLSelect;
            //strSQL = strSQL + "  FROM WIP W LEFT JOIN OPRUNTCST OUC ON W.OPER = OUC.OPER " + strSQLCond + " AND '" + strCrntDT + "' BETWEEN YYYYMMDD AND EXPIRY_DATE \n";
            //strSQL = strSQL + " GROUP BY W.OPER_DESC, W.OPER, OUC.CREATE_CODE, OUC.UNIT_COST \n";
            //strSQL = strSQL + " ORDER BY RPT_TYPE, OPER \n";

            return strSQL;
        }
        private void Init_ChkMore()
        {
            if (chkMore.Checked)
            {
                trInquiry2.Visible = rdoProdType.AutoPostBack = true;
            }
            else
            {
                trInquiry2.Visible = rdoProdType.AutoPostBack = false;
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
            string strSQL = "", strMatID = "", strOper = "", strCreateCode = "", strLotID = "", strCustLotID = "", strBackDT = "", strCrntDT = "", strFName = "";

            if (GetDefaultValue(ref strMatID, ref strOper, ref strCreateCode, ref strLotID, ref strCustLotID, ref strBackDT, ref strCrntDT) == false) return;
            if ((strSQL = MakeQuery(strMatID, strOper, strCreateCode, strLotID, strCustLotID, strBackDT, strCrntDT).Trim()) == "") return;
            if (strSQL.Length < 1 || strSQL == "") return;

            WIPStockModule dtWIP = new WIPStockModule();
            strFName = "REPORT_" + this.Title + "_" + rdoProdType.SelectedValue + "_" + ((strBackDT == "") ? strCrntDT : strBackDT);
            GF.CreateReport2(dtWIP, 2, strSQL, ReportViewer1, "WM.CurrentWIPCost.rdlc", "dsCurrentWIPCost", strFName);
        }
        protected void rdoProdType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
        }

        protected void chkHolded_CheckedChanged(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
        }
        #endregion
    }
}