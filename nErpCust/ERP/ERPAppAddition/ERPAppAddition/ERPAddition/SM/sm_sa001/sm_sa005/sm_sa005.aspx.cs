using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

namespace ERPAppAddition.ERPAddition.SM.sm_sa001.sm_sa005
{
    public partial class sm_sa005 : System.Web.UI.Page
    {

        sa_fun fun = new sa_fun();

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string userid;
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;

                SetComboBoxScrap();
            }

            ShowTr();
        }

        private void SetComboBoxScrap()
        {
            string sSql = "SELECT GROUP1_DESC AS ITEM_NM, GROUP1_CODE AS ITEM_CD   FROM DBO.SA_SYS_CODE WHERE 1=1 GROUP BY GROUP1_DESC, GROUP1_CODE";
            SetcmbVelue(sSql, ddl_ScrapList);
        }

        private void SetcmbVelue(string sSql, DropDownList cmb)
        {
            DataTable UNIT = fun.getData(sSql);
            if (UNIT.Rows.Count > 0)
            {

                DataRow dr = UNIT.NewRow();

                UNIT.Rows.InsertAt(dr, 0);

                cmb.DataTextField = "ITEM_NM";
                cmb.DataValueField = "ITEM_CD";
                cmb.DataSource = UNIT;
                
            }
            else
            {
                DataTable dt = new DataTable();
                cmb.DataSource = dt;
            }
            cmb.DataBind();
        }

       

        protected void ddl_ScrapList_SelectedIndexChanged(object sender, EventArgs e)
        {
            initiControl();
            SetcmbVelue(GetLotNo(), ddl_R_Doc_No);

            string sSql = "";
            sSql = "SELECT UD_REFERENCE AS ITEM_CD , UD_REFERENCE AS ITEM_NM FROM B_USER_DEFINED_MINOR WHERE UD_MAJOR_CD = 'SA003'";
            SetcmbVelue(sSql, ddl_PlantCD);

            sSql = "SELECT UD_MINOR_CD AS ITEM_CD , UD_MINOR_NM AS ITEM_NM FROM B_USER_DEFINED_MINOR WHERE UD_MAJOR_CD = 'SA003'";
            SetcmbVelue(sSql, ddl_WeightEqp);
        }

        private void initiControl()
        {
            txtEtchWeight.Text = "";
            txtEtchWeightKey.Text = "";
            txtEtchWUnit.Text = "";
            txtPlatWeight1.Text = "";
            txtPlatWeight2.Text = "";
            txtPlatWgtKey1.Text = "";
            txtPlatWgtKey2.Text = "";
            txtPlatWUnit1.Text = "";
            txtPlatWUnit2.Text = "";


            btnEtchSave.Enabled = true;
            btnGetEtchWeight.Enabled = true;
            btnPlatSave1.Enabled = true;
            btnGetPlatWeight1.Enabled = true;
            btnGetPlatWeight2.Enabled = true;
            btnPlatSave2.Enabled = true;

            DataTable dt = new DataTable();
            ddl_PlantCD.DataSource = dt;
            ddl_PlantCD.DataBind();
            ddl_WeightEqp.DataSource = dt;
            ddl_WeightEqp.DataBind();
        }

        private void ShowTr()
        {
            if (ddl_ScrapList.SelectedValue == "S02")
            {
                
                PLAT.Attributes.Add("style", "display:none");
                ETCH.Attributes.Clear();
            }
            else if (ddl_ScrapList.SelectedValue == "S01")
            {
                PLAT.Attributes.Clear();
                ETCH.Attributes.Add("style", "display:none");
            }
        }

        protected void btnGetEtchWeight_Click(object sender, EventArgs e)
        {
            if (ChkGetWeight())
            {
                DataTable dt = GetWeight();
                if (dt.Rows.Count > 0)
                {
                    txtEtchWeight.Text = dt.Rows[0]["WEIGHT_VALUE"].ToString();
                    txtEtchWUnit.Text = dt.Rows[0]["WEIGHT_UNIT"].ToString();
                    txtEtchWeightKey.Text = dt.Rows[0]["WKEY"].ToString();
                    //ETCH.Attributes.Clear();
                }
            }

        }

        private DataTable GetWeight()
        {

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(" SELECT TOP 1 ");
            sb.AppendLine("  TRANS_TIME+'|'+ PLANT_CD+'|'+EQUIPMENT_ID AS WKEY");
            sb.AppendLine("  , PLANT_CD ");
            sb.AppendLine("  , EQUIPMENT_ID ");
            sb.AppendLine("  , WEIGHT_VALUE ");
            sb.AppendLine("  , WEIGHT_UNIT ");
            sb.AppendLine("  , WEIGHT_TIME ");
            sb.AppendLine(" FROM OUT_MAT_WEIGHT ");
            sb.AppendLine("  WHERE 1=1 ");
            sb.AppendLine("    AND STUFF(STUFF(STUFF(TRANS_TIME,13,0,':'),11,0,':'),9,0,' ')  > DATEADD(MINUTE, -30, GETDATE()) ");
            sb.AppendLine("    AND PLANT_CD = '"+ddl_PlantCD.Text+"'");
            sb.AppendLine("    AND EQUIPMENT_ID = '"+ddl_WeightEqp.Text+"'");
            sb.AppendLine("  ORDER BY TRANS_TIME DESC ");

            DataTable dt = fun.getData(sb.ToString());


            return dt;
        }

        private string GetLotNo()
        {

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(" SELECT ");
            sb.AppendLine("  AA.LOT_NO AS ITEM_CD ");
            sb.AppendLine("  , AA.LOT_NO +' ['+BB.R_RMK+']' AS ITEM_NM");
            sb.AppendLine(" FROM OUT_MAT_DRAIN AA ");
            sb.AppendLine(" INNER JOIN OUT_MAT_HIS BB ");
            sb.AppendLine("    ON AA.LOT_NO = BB.LOT_NO ");
            sb.AppendLine("  WHERE 1=1 ");
            sb.AppendLine("    AND AA.OUT_YN = 'N' ");
            sb.AppendLine("    AND AA.STATE_FLAG <> 'D'");
            sb.AppendLine("    AND DRAIN_MAT = '" + ddl_ScrapList.Text + "'");
           

            if(ddl_ScrapList.SelectedValue == "S02")
            {
                sb.AppendLine("    AND NOT EXISTS (SELECT * FROM OUT_LOT_AU_WEIGHT CC WHERE CC.DRAIN_MAT = 'S02' AND CC.LOT_NO = BB.LOT_NO )");
            }
            else if
                (ddl_ScrapList.SelectedValue == "S01")
            {
                sb.AppendLine("    AND NOT EXISTS (SELECT * FROM OUT_LOT_AU_WEIGHT CC WHERE CC.DRAIN_MAT = 'S01' AND CC.LOT_NO = BB.LOT_NO AND CC.WEIGHT_KEY2 IS NOT NULL )");
            }

            sb.AppendLine("  GROUP BY AA.LOT_NO, BB.R_RMK ");

            return sb.ToString();
        }

        protected void btnGetPlatWeight1_Click(object sender, EventArgs e)
        {

            if (ChkGetWeight())
            {
                DataTable dt = GetWeight();
                if (dt.Rows.Count > 0)
                {
                    txtPlatWeight1.Text = dt.Rows[0]["WEIGHT_VALUE"].ToString();
                    txtPlatWUnit1.Text = dt.Rows[0]["WEIGHT_UNIT"].ToString();

                    txtPlatWgtKey1.Text = dt.Rows[0]["WKEY"].ToString();
                    //ETCH.Attributes.Clear();
                }
            }
        }

        protected void btnGetPlatWeight2_Click(object sender, EventArgs e)
        {
            if (ChkGetWeight())
            {
                DataTable dt = GetWeight();
                if (dt.Rows.Count > 0)
                {
                    txtPlatWeight2.Text = dt.Rows[0]["WEIGHT_VALUE"].ToString();
                    txtPlatWUnit2.Text = dt.Rows[0]["WEIGHT_UNIT"].ToString();

                    txtPlatWgtKey2.Text = dt.Rows[0]["WKEY"].ToString();
                    //ETCH.Attributes.Clear();
                }
            }
        }

        private bool ChkGetWeight()
        {
            bool result = true;
            
            if(ddl_PlantCD.Text == "")
            {
                MessageBox.ShowMessage("회수 공장을 선택해 주세요! ", this.Page);
                result = false;
            }
            else if (ddl_WeightEqp.Text == "")
            {
                MessageBox.ShowMessage("측정 저울을 선택해 주세요! ", this.Page);
                result = false;
            }

            return result;

        }
                
        public int QueryExecute(string sql)
        {
            int value;

            sql_conn.Open();
            sql_cmd = sql_conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = sql;

            try
            {
                value = sql_cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
                value = -1;
            }
            sql_conn.Close();
            return value;
        }

        protected void btnEtchSave_Click(object sender, EventArgs e)
        {
            if (ddl_ScrapList.SelectedValue == "S02")
            {
                if (ddl_R_Doc_No.Text == "")
                {
                    MessageBox.ShowMessage("LotNo 를 선택하세요. ", this.Page);
                    return;
                }
                if (ddl_PlantCD.Text == "")
                {
                    MessageBox.ShowMessage("회수 공장을 선택하세요. ", this.Page);
                    return;
                }
                if (ddl_WeightEqp.Text == "")
                {
                    MessageBox.ShowMessage("측정 정울을 선택하세요. ", this.Page);
                    return;
                }
                if (txtEtchWeight.Text == "" || txtEtchWUnit.Text == "" || txtEtchWeightKey.Text == "")
                {
                    MessageBox.ShowMessage("측정값에 문제가 있습니다.. ", this.Page);
                    return;
                }


                StringBuilder sbSQL = new StringBuilder();

                sbSQL.AppendLine(" INSERT INTO OUT_LOT_AU_WEIGHT ");
                sbSQL.AppendLine(" ( ");
                sbSQL.AppendLine(" 	DRAIN_MAT ");
                sbSQL.AppendLine(" 	, LOT_NO ");
                sbSQL.AppendLine(" 	, WEIGHT_KEY1 ");
                sbSQL.AppendLine(" 	, WEIGHT_VALUE1 ");
                sbSQL.AppendLine(" 	, AU_QTY ");
                sbSQL.AppendLine(" 	, WEIGHT_UNIT ");
                sbSQL.AppendLine(" 	, INSRT_USER_ID ");
                sbSQL.AppendLine(" 	, INSRT_DT ");
                sbSQL.AppendLine(" 	, UPDT_USER_ID ");
                sbSQL.AppendLine(" 	, UPDT_DT ");
                sbSQL.AppendLine(" ) ");
                sbSQL.AppendLine(" VALUES ");
                sbSQL.AppendLine(" ( ");
                sbSQL.AppendLine(" 	 '" + ddl_ScrapList.SelectedValue + "'");
                sbSQL.AppendLine(" 	 ,'" + ddl_R_Doc_No.SelectedValue + "'");
                sbSQL.AppendLine(" 	 ,'" + txtEtchWeightKey.Text + "'");
                sbSQL.AppendLine(" 	 ," + txtEtchWeight.Text + "");
                sbSQL.AppendLine(" 	 ," + txtEtchWeight.Text + "");
                sbSQL.AppendLine(" 	 ,'" + txtEtchWUnit.Text + "'");
                sbSQL.AppendLine(" 	 ,'" + Session["User"].ToString() + "'");
                sbSQL.AppendLine(" 	, GETDATE() ");
                sbSQL.AppendLine(" 	 ,'" + Session["User"].ToString() + "'");
                sbSQL.AppendLine(" 	, GETDATE() "); 
                sbSQL.AppendLine(" ) ");

                if (QueryExecute(sbSQL.ToString()) < 0)
                {
                    MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                    return;
                }
                MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);

                SetLotSearch(ddl_ScrapList.SelectedValue, ddl_R_Doc_No.SelectedValue);

            }
            
        }

        private void SetLotSearch(string sMat, string sLot)
        {
            StringBuilder sbSQL = new StringBuilder();

            if(sMat == "S02")
            {
                sbSQL.AppendLine(" SELECT   ");
                sbSQL.AppendLine("  DRAIN_MAT  ");
                sbSQL.AppendLine("  , LOT_NO  ");
                sbSQL.AppendLine("  , WEIGHT_KEY1  ");
                sbSQL.AppendLine("  , WEIGHT_VALUE1  ");
                sbSQL.AppendLine("  , AU_QTY  ");
                sbSQL.AppendLine("  , WEIGHT_UNIT  ");
                sbSQL.AppendLine(" FROM OUT_LOT_AU_WEIGHT  ");
                sbSQL.AppendLine(" WHERE 1=1  ");
                sbSQL.AppendLine("  AND LOT_NO = '" + sLot + "'");
                sbSQL.AppendLine("  AND DRAIN_MAT = '" + sMat + "'");

                DataTable dt = fun.getData(sbSQL.ToString());

                if(dt.Rows.Count > 0)
                {
                    btnEtchSave.Enabled = false;
                    btnGetEtchWeight.Enabled = false;
                    txtEtchWeight.Text = dt.Rows[0]["WEIGHT_VALUE1"].ToString();
                    txtEtchWUnit.Text = dt.Rows[0]["WEIGHT_UNIT"].ToString();
                    txtEtchWeightKey.Text = dt.Rows[0]["WEIGHT_KEY1"].ToString();
                }
               
            }
            else if(sMat == "S01")
            {
                sbSQL.AppendLine(" SELECT   ");
                sbSQL.AppendLine("  DRAIN_MAT  ");
                sbSQL.AppendLine("  , LOT_NO  ");
                sbSQL.AppendLine("  , ISNULL(WEIGHT_KEY1, '') AS WEIGHT_KEY1  ");
                sbSQL.AppendLine("  , WEIGHT_VALUE1  ");
                sbSQL.AppendLine("  , ISNULL(WEIGHT_KEY2, '') AS WEIGHT_KEY2 ");
                sbSQL.AppendLine("  , WEIGHT_VALUE2  ");
                sbSQL.AppendLine("  , AU_QTY  ");
                sbSQL.AppendLine("  , WEIGHT_UNIT  ");
                sbSQL.AppendLine(" FROM OUT_LOT_AU_WEIGHT  ");
                sbSQL.AppendLine(" WHERE 1=1  ");
                sbSQL.AppendLine("  AND LOT_NO = '" + sLot + "'");
                sbSQL.AppendLine("  AND DRAIN_MAT = '" + sMat + "'");

                DataTable dt = fun.getData(sbSQL.ToString());

                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["WEIGHT_KEY1"].ToString() != "")
                    {
                        btnPlatSave1.Enabled = false;
                        btnGetPlatWeight1.Enabled = false;
                        txtPlatWeight1.Text = dt.Rows[0]["WEIGHT_VALUE1"].ToString();
                        txtPlatWUnit1.Text = dt.Rows[0]["WEIGHT_UNIT"].ToString();
                        txtPlatWgtKey1.Text = dt.Rows[0]["WEIGHT_KEY1"].ToString();
                    }
                    else
                    {
                        btnPlatSave1.Enabled = true;
                        btnGetPlatWeight1.Enabled = true;
                        txtPlatWeight1.Text = "";
                        txtPlatWUnit1.Text = "";
                        txtPlatWgtKey1.Text = "";

                    }
                    if (dt.Rows[0]["WEIGHT_KEY2"].ToString() != "")
                    {
                        btnPlatSave2.Enabled = false;
                        btnGetPlatWeight2.Enabled = false;
                        txtPlatWeight2.Text = dt.Rows[0]["WEIGHT_VALUE2"].ToString();
                        txtPlatWUnit2.Text = dt.Rows[0]["WEIGHT_UNIT"].ToString();
                        txtPlatWgtKey2.Text = dt.Rows[0]["WEIGHT_KEY2"].ToString();
                    }
                    else
                    {
                        btnPlatSave2.Enabled = true;
                        btnGetPlatWeight2.Enabled = true;
                        txtPlatWeight2.Text = "";
                        txtPlatWUnit2.Text = "";
                        txtPlatWgtKey2.Text = "";
                    }
                }
                
            }
        }

        protected void ddl_R_Doc_No_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(ddl_R_Doc_No.SelectedValue != "")
            {
                SetLotSearch(ddl_ScrapList.SelectedValue, ddl_R_Doc_No.SelectedValue);
            }
            else
            {
                initiControl();
            }
        }

        protected void btnPlatSave1_Click(object sender, EventArgs e)
        {
            if (ddl_ScrapList.SelectedValue == "S01")
            {
                if (ddl_R_Doc_No.Text == "")
                {
                    MessageBox.ShowMessage("LotNo 를 선택하세요. ", this.Page);
                    return;
                }
                if (ddl_PlantCD.Text == "")
                {
                    MessageBox.ShowMessage("회수 공장을 선택하세요. ", this.Page);
                    return;
                }
                if (ddl_WeightEqp.Text == "")
                {
                    MessageBox.ShowMessage("측정 정울을 선택하세요. ", this.Page);
                    return;
                }
                if (txtPlatWeight1.Text == "" || txtPlatWUnit1.Text == "" || txtPlatWgtKey1.Text == "")
                {
                    MessageBox.ShowMessage("측정값에 문제가 있습니다.. ", this.Page);
                    return;
                }


                StringBuilder sbSQL = new StringBuilder();

                sbSQL.AppendLine(" INSERT INTO OUT_LOT_AU_WEIGHT ");
                sbSQL.AppendLine(" ( ");
                sbSQL.AppendLine(" 	DRAIN_MAT ");
                sbSQL.AppendLine(" 	, LOT_NO ");
                sbSQL.AppendLine(" 	, WEIGHT_KEY1 ");
                sbSQL.AppendLine(" 	, WEIGHT_VALUE1 ");
                sbSQL.AppendLine(" 	, WEIGHT_UNIT ");
                sbSQL.AppendLine(" 	, INSRT_USER_ID ");
                sbSQL.AppendLine(" 	, INSRT_DT ");
                sbSQL.AppendLine(" 	, UPDT_USER_ID ");
                sbSQL.AppendLine(" 	, UPDT_DT ");
                sbSQL.AppendLine(" ) ");
                sbSQL.AppendLine(" VALUES ");
                sbSQL.AppendLine(" ( ");
                sbSQL.AppendLine(" 	 '" + ddl_ScrapList.SelectedValue + "'");
                sbSQL.AppendLine(" 	 ,'" + ddl_R_Doc_No.SelectedValue + "'");
                sbSQL.AppendLine(" 	 ,'" + txtPlatWgtKey1.Text + "'");
                sbSQL.AppendLine(" 	 ," + txtPlatWeight1.Text + "");
                sbSQL.AppendLine(" 	 ,'" + txtPlatWUnit1.Text + "'");
                sbSQL.AppendLine(" 	 ,'" + Session["User"].ToString() + "'");
                sbSQL.AppendLine(" 	, GETDATE() ");
                sbSQL.AppendLine(" 	 ,'" + Session["User"].ToString() + "'");
                sbSQL.AppendLine(" 	, GETDATE() ");
                sbSQL.AppendLine(" ) ");

                if (QueryExecute(sbSQL.ToString()) < 0)
                {
                    MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                    return;
                }
                MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);

                SetLotSearch(ddl_ScrapList.SelectedValue, ddl_R_Doc_No.SelectedValue);

            }
        }

        protected void btnPlatSave2_Click(object sender, EventArgs e)
        {
            if (ddl_ScrapList.SelectedValue == "S01")
            {
                if (txtPlatWeight1.Text == "")
                {
                    MessageBox.ShowMessage("작업전 중량 정보가 없습니다. ", this.Page);
                    return;
                }
                if (ddl_R_Doc_No.Text == "")
                {
                    MessageBox.ShowMessage("LotNo 를 선택하세요. ", this.Page);
                    return;
                }
                if (ddl_PlantCD.Text == "")
                {
                    MessageBox.ShowMessage("회수 공장을 선택하세요. ", this.Page);
                    return;
                }
                if (ddl_WeightEqp.Text == "")
                {
                    MessageBox.ShowMessage("측정 정울을 선택하세요. ", this.Page);
                    return;
                }
                if (txtPlatWeight2.Text == "" || txtPlatWUnit2.Text == "" || txtPlatWgtKey2.Text == "")
                {
                    MessageBox.ShowMessage("측정값에 문제가 있습니다. ", this.Page);
                    return;
                }

                if(txtPlatWgtKey1.Text == txtPlatWgtKey2.Text)
                {
                    MessageBox.ShowMessage("작업 전/후 측정 정보가 같습니다. ", this.Page);
                    txtPlatWgtKey2.Text = "";
                    txtPlatWUnit2.Text = "";
                    txtPlatWeight2.Text = "";
                    return;
                }


                StringBuilder sbSQL = new StringBuilder();

                sbSQL.AppendLine(" UPDATE OUT_LOT_AU_WEIGHT ");
                sbSQL.AppendLine(" SET WEIGHT_KEY2 = '" + txtPlatWgtKey2.Text + "'");
                sbSQL.AppendLine(" 	 , WEIGHT_VALUE2 = '" + txtPlatWeight2.Text + "'");
                sbSQL.AppendLine(" 	, AU_QTY = " + txtPlatWeight2.Text + " - WEIGHT_VALUE1" );
                sbSQL.AppendLine(" 	, WEIGHT_UNIT = '" + txtPlatWUnit2.Text + "'");
                sbSQL.AppendLine(" 	, UPDT_USER_ID =  '" + Session["User"].ToString() + "'");
                sbSQL.AppendLine(" 	, UPDT_DT = GETDATE() ");
                sbSQL.AppendLine(" WHERE 1=1 ");
                sbSQL.AppendLine("  AND  DRAIN_MAT = '" + ddl_ScrapList.SelectedValue + "'");
                sbSQL.AppendLine("  AND  LOT_NO = '" + ddl_R_Doc_No.SelectedValue + "'");
              

                if (QueryExecute(sbSQL.ToString()) < 0)
                {
                    MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                    return;
                }
                MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);

                SetLotSearch(ddl_ScrapList.SelectedValue, ddl_R_Doc_No.SelectedValue);

            }
        }

    }
}