using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.CM.CM_C2001;

namespace ERPAppAddition.ERPAddition.CM.CM_C2001
{
    public partial class CM_C2001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        DataSet dt = new DataSet();
        cls_prod_qty_month cls_dbexe = new cls_prod_qty_month();
        string ls_fr_dt, ls_to_dt ;
        int value;
        string ls_report_nm, ls_sql, ls_ddl_sql, ls_cost_cd, ls_yyyymm, ls_cost_cd_gp, ls_item_cd_gp, ls_weight;
        string sql , now_month, before_month;
        Decimal ld_weight;

        protected void Page_Load(object sender, EventArgs e)
        {
            Panel1.Visible = false;
            Panel2.Visible = false;
            //tb_fr_dt.Text = DateTime.Now.ToString("yyyyMM");
            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void db_exec(string type, string ls_sql)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = ls_sql;


            try
            {
                //숫자체크면 
                if (type == "check")
                {
                    value = cmd.ExecuteNonQuery();
                }
                else
                {
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void insert_prod_qty()
        {
            // 생산일보데이타 저장
            // might 디비에서 그룹웨어와 동일한 생산일보 데이타를 가져온다. (한달치)
            ls_fr_dt = tb_fr_dt.Text + "01";
            ls_to_dt = DateTime.ParseExact(ls_fr_dt, "yyyyMMdd", null).AddMonths(1).AddDays(-1).ToString("yyyyMMdd");

            DataTable dt1 = cls_dbexe.fetch(ls_fr_dt, ls_to_dt);
            if (dt1.Rows.Count > 0)
            {
                DataRow row = dt1.Rows[0];
                string ls_yyyymm, ls_m_gubun, ls_m_gubun_detail, ls_m_gubun_detail_nm, ls_seq;
                decimal ld_ddi_val, ld_wlp_val, ld_p_test_val, ld_cog_val, ld_cof_val, ld_f_test_val, ld_wlcsp_val, ld_twelve_val, ld_sangpum_val;
                DateTime ldt_isrt_dt;


                ls_yyyymm = row["yyyymm"].ToString();
                ls_m_gubun = row["m_gubun"].ToString();
                ls_m_gubun_detail = row["m_gubun_detail"].ToString();
                ls_m_gubun_detail_nm = row["m_gubun_detail_nm"].ToString();
                ld_ddi_val = Convert.ToDecimal(row["ddi_val"].ToString());
                ld_wlp_val = Convert.ToDecimal(row["wlp_val"].ToString());
                ld_p_test_val = Convert.ToDecimal(row["p_test_val"].ToString());
                ld_cog_val = Convert.ToDecimal(row["cog_val"].ToString());
                ld_cof_val = Convert.ToDecimal(row["cof_val"].ToString());
                ld_f_test_val = Convert.ToDecimal(row["f_test_val"].ToString());
                ld_wlcsp_val = Convert.ToDecimal(row["wlcsp_val"].ToString());
                ld_twelve_val = Convert.ToDecimal(row["twelve_val"].ToString());
                ld_sangpum_val = Convert.ToDecimal(row["sangpum_val"].ToString());
                ldt_isrt_dt = Convert.ToDateTime(row["isrt_dt"].ToString());
                ls_seq = row["seq"].ToString();

                ls_sql = " insert into FourM_R1002_NEPES " +
                         " Values('" + ls_yyyymm + "','" + ls_m_gubun + "','" + ls_m_gubun_detail + "','" + ls_m_gubun_detail_nm + "','" + ld_ddi_val + "','" + ld_wlp_val + "', " +
                         " '" + ld_p_test_val + "','" + ld_cog_val + "','" + ld_cof_val + "','" + ld_f_test_val + "','" + ld_wlcsp_val + "','" + ld_twelve_val + "','" + ld_sangpum_val + "',getdate(), '" + ls_seq + "' ) ";
                db_exec("nocheck",ls_sql);
               
            }
            else
            {
                return;
            }   
        }
        
        private void insert_data()
        {

            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "usp_fourm_sum";
            SqlParameter param1 = new SqlParameter("@yyyymm", SqlDbType.VarChar, 6);
            param1.Value = tb_fr_dt.Text;
            cmd.Parameters.Add(param1);

            try
            {
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }

            //생산일보데이타 insert
            insert_prod_qty();
             
        }

        protected void btn_request_Click(object sender, EventArgs e)
        {
            insert_data();
            //QueryCreator("new");    
            btn_old_request__Click(null, null);
            MessageBox.ShowMessage("계산완료 되었습니다", this.Page);
        }

        private void QueryCreator(string type)
        {
            if ((tb_fr_dt.Text == "") || (tb_fr_dt.Text == null))
            {
                string script = "alert(\"조회년월을 선택하여 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                tb_fr_dt.Focus();

            }
            else
            {
                ReportViewer1.Reset();

                //if (type == "new")
                //{
                //    // 신규 데이타 저장
                //    insert_data();
                //}
                if (RadioButtonList1.SelectedValue == "view1")
                {

                    ls_sql =
                    " WITH DATAHOUSE AS ( " +
                    " SELECT CASE WHEN M_GUBUN = 'QTY' THEN '수량'  " +
                    "             WHEN M_GUBUN = 'AMT' THEN '금액' " +
                    "             WHEN M_GUBUN = 'Panguan' THEN '판매비' " +
                    "             ELSE M_GUBUN END M_GUBUN_NM " +
                    "       ,YYYYMM,M_GUBUN,M_GUBUN_DETAIL,M_GUBUN_DETAIL_NM,DDI_VAL " +
                    "       ,WLP_VAL,P_TEST_VAL,COG_VAL,COF_VAL,F_TEST_VAL,WLCSP_VAL,TWELVE_VAL,SANGPUM_VAL,seq " +
                    " FROM FourM_R1002_NEPES		  " +
                    " WHERE YYYYMM = '" + tb_fr_dt.Text + "' " +
                    " ) " +
                    " SELECT *  " +
                    "   FROM DATAHOUSE " +
                    " UNION ALL " +
                    " SELECT A.*  " +
                    " FROM ( " +
                    " 		SELECT M_GUBUN_NM, YYYYMM, M_GUBUN,'' M_GUBUN_DETAIL ,'합계' M_GUBUN_DETAIL_NM " +
                    " 			 , ISNULL(SUM(ddi_val),0) ddi_val, ISNULL(SUM(wlp_val),0) wlp_val, ISNULL(SUM(p_test_val),0) p_test_val, ISNULL(SUM(cog_val),0) cog_val, ISNULL(SUM(cof_val),0) cof_val " +
                    " 			 , ISNULL(SUM(f_test_val),0) f_test_val, ISNULL(SUM(wlcsp_val),0) wlcsp_val, ISNULL(SUM(TWELVE_VAL),0) TWELVE_VAL " +
                    " 			 , ISNULL(SUM(sangpum_val),0)  sangpum_val " +
                    " 			 , CONVERT(INT,SEQ)  + '9' SEQ " +
                    " 		 FROM DATAHOUSE    " +
                    " 		WHERE SEQ <> '10'  " +
                    " 		GROUP BY  M_GUBUN_NM, YYYYMM, M_GUBUN,SEQ " +
                    " 		) A " +
                    " ORDER BY SEQ, M_GUBUN,M_GUBUN_DETAIL ";
                    //ds_cm_c2001 dt1 = new ds_cm_c2001();
                    dt = new ds_cm_c2001();
                    //ls_sql = "select * from FourM_R1002_NEPES where yyyymm = '" + tb_fr_dt.Text +"' ";
                    ls_report_nm = "rp_cm_c2001.rdlc";
                }

                if (RadioButtonList1.SelectedValue == "view2")
                {

                    if (DropDownList1.SelectedValue == "A")
                    {
                        ls_sql = "SELECT * FROM FourM_R2001_NEPES WHERE YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_a();
                        ls_report_nm = "rp_cm_c2001_a.rdlc";
                    }
                    if (DropDownList1.SelectedValue == "B")
                    {
                        ls_sql = "SELECT * FROM FourM_R2002_NEPES WHERE YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_b();
                        ls_report_nm = "rp_cm_c2001_b.rdlc";
                    }
                    if (DropDownList1.SelectedValue == "C")
                    {
                        ls_sql = "SELECT * FROM FourM_R2003_NEPES WHERE YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_c();
                        ls_report_nm = "rp_cm_c2001_c.rdlc";
                    }
                    if (DropDownList1.SelectedValue == "D")
                    {
                        ls_sql = "SELECT * FROM FourM_R2004_NEPES WHERE YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_d();
                        ls_report_nm = "rp_cm_c2001_d.rdlc";
                    }
                    if (DropDownList1.SelectedValue == "E")
                    {
                        ls_sql = "SELECT * FROM FourM_R2005_NEPES WHERE YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_e();
                        ls_report_nm = "rp_cm_c2001_e.rdlc";
                    }

                    if (DropDownList1.SelectedValue == "F")
                    {
                        ls_fr_dt = tb_fr_dt.Text + "01";
                        ls_to_dt = DateTime.ParseExact(ls_fr_dt, "yyyyMMdd", null).AddMonths(1).AddDays(-1).ToString("yyyyMMdd");
                        //기존자료 삭제하고
                        //db_exec("check", ls_sql);
                        //존재여부를 확인해서 
                        db_exec("check", "SELECT count(*) FROM FourM_R1002_NEPES	 where M_gubun = 'QTY' AND M_GUBUN_DETAIL = '1-PROD' AND  YYYYMM = '" + tb_fr_dt.Text + "'");
                        //다시 생산일보 데이타 가져오고
                        if (value < 1)
                        {
                            insert_prod_qty();
                        }
                        //cls_dbexe.fetch(ls_fr_dt, ls_to_dt);


                        ls_sql = "SELECT * FROM FourM_R1002_NEPES	 where M_gubun = 'QTY' AND M_GUBUN_DETAIL = '1-PROD' AND  YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_f();
                        ls_report_nm = "rp_cm_c2001_f.rdlc";
                    }
                }
                if (RadioButtonList1.SelectedValue == "view3")
                {
                    if (DropDownList1.SelectedValue == "A")
                    {
                        ls_sql = "select A.yyyymm, A.cost_cd, B.COST_NM cost_cd_nm, A.cost_gp_cd,C.UD_MINOR_NM cost_gp_cd_nm " +
                                 "  from FourM_B1001_NEPES A INNER JOIN B_COST_CENTER B ON A.cost_cd = B.COST_CD " +
                                 "                           INNER JOIN B_USER_DEFINED_MINOR C ON A.cost_gp_cd = C.UD_MINOR_CD AND C.UD_MAJOR_CD = 'A0002' " +
                                 " where YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_sys_b1();
                        ls_report_nm = "rp_cm_c2001_sys_b1.rdlc";
                    }

                    if (DropDownList1.SelectedValue == "B")
                    {
                        ls_sql = "select A.yyyymm, A.cost_gp_cd, B.UD_MINOR_NM COST_GP_CD_NM, A.item_gp_cd, C.UD_MINOR_NM ITEM_GP_CD_NM, A.weight " +
                                 "  from FourM_B1002_NEPES A INNER JOIN B_USER_DEFINED_MINOR B ON A.cost_gp_cd = B.UD_MINOR_CD AND B.UD_MAJOR_CD = 'A0002' " +
                                 "                           INNER JOIN B_USER_DEFINED_MINOR C ON A.item_gp_cd = C.UD_MINOR_CD AND C.UD_MAJOR_CD = 'A0003'" +
                                 " where YYYYMM = '" + tb_fr_dt.Text + "'";
                        dt = new ds_cm_c2001_sys_b2();
                        ls_report_nm = "rp_cm_c2001_sys_b2.rdlc";
                    }

                    if (DropDownList1.SelectedValue == "C")
                    {
                        ls_sql = "SELECT A.ITEM_NM , A.PROC_TYPE, A.ROUTE_SET , A.INCH, B.COST_ITEM_GP " +
                                 "  FROM FourM_B1003_NEPES A INNER JOIN FourM_B1004_NEPES B " +
                                 "                     ON A.PROC_TYPE = B.PROC_TYPE AND A.ROUTE_SET = B.ROUTE_SET AND A.INCH = B.INCH ";
                        dt = new ds_cm_c2001_sys_b3();
                        ls_report_nm = "rp_cm_c2001_sys_b3.rdlc";
                    }
                    if (DropDownList1.SelectedValue == "D")
                    {
                        ls_sql = " SELECT PROC_TYPE,ROUTE_SET,INCH,COST_ITEM_GP FROM FourM_B1004_NEPES " ;
                               
                        dt = new ds_cm_c2001_sys_b4();
                        ls_report_nm = "rp_cm_c2001_sys_b4.rdlc";
                    }
                    if (DropDownList1.SelectedValue == "E")
                    {
                        ls_sql = " SELECT DISTINCT B.PROC_TYPE,B.ROUTE_SET, B.INCH, A.COST_ITEM_GP " +
                                 "   FROM FourM_R2005_NEPES A LEFT JOIN (SELECT a.ITEM_NM,  a.PROC_TYPE, a.ROUTE_SET, a.INCH, b.COST_ITEM_GP  " +
                                 "                                       FROM FourM_B1003_NEPES a left join  FourM_B1004_NEPES b on a.PROC_TYPE = b.PROC_TYPE and a.ROUTE_SET = b.ROUTE_SET and a.INCH = b.INCH" +
                                 "                                      ) B ON A.item_nm = B.ITEM_NM" +
                                 "  WHERE A.COST_ITEM_GP IS NULL AND A.YYYYMM =  '" + tb_fr_dt.Text + "' " +
                                 "  ORDER BY A.COST_ITEM_GP ";

                        dt = new ds_cm_c2001_sys_b4();
                        ls_report_nm = "rp_cm_c2001_sys_b4.rdlc";
                    }
                }

                ReportCreator(dt, ls_sql, ReportViewer1, ls_report_nm, "DataSet1");
            }
        }

        private void ReportCreator(DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandTimeout = 500;
            cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;
            try
            {
               
                
                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);
                string title;
                if (RadioButtonList1.SelectedValue == "view1")
                {
                    title = "_집계_";
                }
                else
                {
                    title ="_" +  DropDownList1.SelectedItem.Text + "_";
                }
                _reportViewer.LocalReport.DisplayName = "REPORT_" + tb_fr_dt.Text + "_4M데이타" + title + DateTime.Now.ToShortDateString();

                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];

                _reportViewer.LocalReport.DataSources.Add(rds);

                _reportViewer.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
        //기존자료조회용
        protected void btn_old_request__Click(object sender, EventArgs e)
        {
            //UpdatePanel2.Triggers.cont
            QueryCreator("old"); 
        }

        // 조회항목선택
        protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
        {

            ReportViewer1.Reset();

            DropDownList1.Items.Clear();
            if (RadioButtonList1.SelectedValue == "view2")
            {
                ls_ddl_sql = "select UD_MINOR_CD, UD_MINOR_NM from B_USER_DEFINED_MINOR where UD_MAJOR_CD = 'A0006' order by UD_MINOR_CD";
                Panel1.Visible = false;
                Panel2.Visible = false;
                lblMessage.Text = "";
            }

            if (RadioButtonList1.SelectedValue == "view3")
            {
                Panel1.Visible = true;
                Panel2.Visible = false;
                ls_ddl_sql = "select UD_MINOR_CD, UD_MINOR_NM from B_USER_DEFINED_MINOR where UD_MAJOR_CD = 'A0007' order by UD_MINOR_CD ";
                lblMessage.Text = "";
            }

            if (RadioButtonList1.SelectedValue == "view1")
            {
                Panel1.Visible = false;
                Panel2.Visible = false;
                lblMessage.Text = "";
                ls_ddl_sql = "";
                return;
            }

            conn.Open();
            
            SqlCommand cmd2 = new SqlCommand(ls_ddl_sql, conn);
            dr = cmd2.ExecuteReader();
            if (DropDownList1.Items.Count < 2)
            {
                DropDownList1.DataSource = dr;
                DropDownList1.DataValueField = "UD_MINOR_CD";
                DropDownList1.DataTextField = "UD_MINOR_NM";
                DropDownList1.DataBind();
            }
            dr.Close();
            conn.Close();
        }

       
        // 업데이트 클릭
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
                string FilePath = Server.MapPath(FolderPath + FileName);
                FileUpload1.SaveAs(FilePath);
                GetExcelSheets(FilePath, Extension, "Yes");
            }
        }
        // 엑셀sheet 받아오기
        private void GetExcelSheets(string FilePath, string Extension, string isHDR)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".xls": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                             .ConnectionString;
                    break;
                case ".xlsx": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                             .ConnectionString;
                    break;
            }

            //Get the Sheets in Excel WorkBoo
            conStr = String.Format(conStr, FilePath, isHDR);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            cmdExcel.Connection = connExcel;
            connExcel.Open();

            //Bind the Sheets to DropDownList
            ddlSheets.Items.Clear();
            ddlSheets.Items.Add(new ListItem("--Select Sheet--", ""));
            ddlSheets.DataSource = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            ddlSheets.DataTextField = "TABLE_NAME";
            ddlSheets.DataValueField = "TABLE_NAME";
            ddlSheets.DataBind();

            DataTable dtCSV = new DataTable();
            OleDbCommand cmdSelect = new OleDbCommand(@"SELECT * FROM [" + ddlSheets.SelectedValue + "]", connExcel);
            OleDbDataAdapter daCSV = new OleDbDataAdapter();
           // OleDbCommand DBCommand;
            IDataReader DBReader;
            DBReader = cmdSelect.ExecuteReader();
            GridView1.DataSource = DBReader;
            GridView1.DataBind();
            //viewContent.Text = sheetname; //DBReader["date"].ToString();
            DBReader.Close();
            //daCSV.SelectCommand = cmdSelect;
            //daCSV.Fill(dtCSV);
            connExcel.Close();
            //txtTable.Text = "";
            HiddenField_fileName.Value = Path.GetFileName(FilePath);
            //FilePath = Path.GetFileName(FilePath);
            Panel2.Visible = true;
            Panel1.Visible = false;
        }

        // 엑셀저장 버튼 클릭시
        protected void btnSave_Click(object sender, EventArgs e)
        {
            //if (tb_fr_dt.Text == "" || tb_fr_dt.Text == null)
            //{
            //    string script = "alert(\"당월을 선택하여 주세요.\");";
            //    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            //    tb_fr_dt.Focus();
            //}
            //else
            //{
                // 코스트센터 그룹등록
                if (DropDownList1.SelectedValue == "A")
                {
                    ls_yyyymm = GridView1.Rows[0].Cells[2].Text.Trim();
                    ////선택한 월의 데이타가 있으면 지우고 입력한다.
                    sql = "select count(yyyymm) from FourM_B1001_NEPES where yyyymm = '" + ls_yyyymm + "' ";
                    if (QueryExecute(sql, "check") > 0)
                    {   //삭제진행
                        sql = " delete FourM_B1001_NEPES  where yyyymm = '" + ls_yyyymm + "' ";
                        QueryExecute(sql, "");
                    }
                    // 데이타 insert 쿼리

                    for (int i = 0; i < GridView1.Rows.Count; i++)
                    {
                        //기본양식 컬럼0-년도, 1 코스트센터, 3-코스트센터그룹 앞에 no, 추가했으니 + 1씩해줘야함.
                        ls_yyyymm = GridView1.Rows[i].Cells[1].Text.Trim();
                        if (ls_yyyymm == null || ls_yyyymm == "")
                        {
                            string script = "alert(\" " + Convert.ToString(i + 1) + " 번째 년월이 잘못되었습니다. 이전까지 저장하고 종료합니다. \");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            return;
                        }
                        ls_cost_cd = GridView1.Rows[i].Cells[2].Text.Trim();
                        if (ls_cost_cd == null || ls_cost_cd == "")
                        {
                            string script = "alert(\" " + Convert.ToString(i + 1) + " 번째 코스트센터코드가 잘못되었습니다. 이전까지 저장하고 종료합니다. \");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            return;
                        }
                        ls_cost_cd_gp = GridView1.Rows[i].Cells[3].Text.Trim();
                        if (ls_cost_cd_gp == null || ls_cost_cd_gp == "")
                        {
                            string script = "alert(\" " + Convert.ToString(i + 1) + " 번째 코스트센터 그룹이 잘못되었습니다. 이전까지 저장하고 종료합니다. \");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            return;
                        }

                        sql = "insert into FourM_B1001_NEPES values('" + ls_yyyymm + "' ,'" + ls_cost_cd + "','" + ls_cost_cd_gp + "'  ) ";

                        int j = i;
                        if (j + 1 == GridView1.Rows.Count || ls_yyyymm == "&nbsp;") 
                        {
                            string script = "alert(\" " + Convert.ToString(i) + "개가 저장되었습니다.\");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);                            
                            tb_fr_dt.Text = GridView1.Rows[0].Cells[2].Text.Trim();
                            btn_old_request__Click(null, null);
                            Panel1.Visible = true;
                            Panel2.Visible = false;
                            GridView1.DataSource = null;
                            GridView1.DataBind();
                            return;
                        }
                        else
                        {
                            if (QueryExecute(sql, "") < 1)
                            {
                                string script = "alert(\" 코스트센터 " + ls_cost_cd + " 저장시 오류가 발생하였습니다.\");";
                                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                                return;
                            }
                        }
                    }
                }

                if (DropDownList1.SelectedValue == "B")
                {
                    ls_yyyymm = GridView1.Rows[0].Cells[2].Text.Trim();
                    sql = "select count(yyyymm) from FourM_B1002_NEPES where yyyymm = '" + ls_yyyymm + "' ";
                    if (QueryExecute(sql, "check") > 0)
                    {   //삭제진행
                        sql = " delete FourM_B1002_NEPES  where yyyymm = '" + ls_yyyymm + "' ";
                        QueryExecute(sql, "");
                    }

                    for (int i = 0; i < GridView1.Rows.Count; i++)
                    {
                        //기본양식 컬럼0-년도, 1 코스트센터그룹, 3-품목그룹, 5-가중치(decimal) 앞에 삭제버튼을 추가했으니 + 1씩해줘야함.
                        ls_yyyymm = GridView1.Rows[i].Cells[1].Text.Trim();

                        if (ls_yyyymm == null || ls_yyyymm == "")
                        {
                            string script = "alert(\" " + Convert.ToString(i+1) + " 번째 년월이 잘못되었습니다. 이전까지 저장하고 종료합니다. \");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            return;
                        }
                        ls_cost_cd_gp = GridView1.Rows[i].Cells[2].Text.Trim();
                        if (ls_cost_cd_gp == null || ls_cost_cd_gp == "")
                        {
                            string script = "alert(\" " + Convert.ToString(i + 1) + " 번째 코스트센터코드가  잘못되었습니다. 이전까지 저장하고 종료합니다. \");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            return;
                        }
                        ls_item_cd_gp = GridView1.Rows[i].Cells[4].Text.Trim();
                        if (ls_item_cd_gp == null || ls_item_cd_gp == "")
                        {
                            string script = "alert(\" " + Convert.ToString(i + 1) + " 번째 품목그룹이 잘못되었습니다. 이전까지 저장하고 종료합니다. \");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            return;
                        }
                        ls_weight = GridView1.Rows[i].Cells[6].Text.Trim();
                        if (ls_weight == null || ls_weight == "")
                        {
                            string script = "alert(\" " + Convert.ToString(i + 1) + " 번째 가중치가 잘못되었습니다. 이전까지 저장하고 종료합니다. \");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            return;
                        }

                        if (ls_weight != "&nbsp;")
                            ld_weight = Convert.ToDecimal(ls_weight);

                        sql = "insert into FourM_B1002_NEPES values('" + ls_yyyymm + "' ,'" + ls_cost_cd_gp + "','" + ls_item_cd_gp + "','" + ld_weight + "'  ) ";

                        int j = i;
                        if ( j + 1  == GridView1.Rows.Count  || ls_yyyymm == "&nbsp;" ) //header 기본 -1에 row수가 0부터 시작하기에 또 -1
                        {
                            string script = "alert(\" " + Convert.ToString(j) + "개가 저장되었습니다.\");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                           
                            tb_fr_dt.Text = GridView1.Rows[0].Cells[2].Text.Trim();
                            btn_old_request__Click(null, null);
                            Panel1.Visible = true;
                            Panel2.Visible = false;
                            GridView1.DataSource = null;
                            GridView1.DataBind();
                            return;

                        }
                        else
                        {
                            if (QueryExecute(sql, "") < 1)
                            {
                                string script = "alert(\" 품목별 가중치  " + ls_cost_cd + " 저장시 오류가 발생하였습니다.\");";
                                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                                return;
                            }
                            
                        }
                    }
                }

                if (DropDownList1.SelectedValue == "E")
                {
                    //sql = "select count(*) from FourM_B1004_NEPES ";
                    //if (QueryExecute(sql, "check") > 0)
                    //{   //삭제진행
                    //    sql = " delete FourM_B1004_NEPES ";
                    //    QueryExecute(sql, "");
                    //}

                    for (int i = 0; i < GridView1.Rows.Count; i++)
                    {
                        //기본양식 컬럼1 -프로세스타입,  2-라우트셋, 3 인치, 4-품목그룹
                        string ls_proc_type = GridView1.Rows[i].Cells[1].Text.Trim();
                        string ls_route_set =  GridView1.Rows[i].Cells[2].Text.Trim();
                        string ls_inch = GridView1.Rows[i].Cells[3].Text.Trim();
                        string ls_cost_item_gp = GridView1.Rows[i].Cells[4].Text.Trim();
                        // 기 등록된 내용이 있으면 삭제
                        sql = "delete FourM_B1004_NEPES where proc_type = '" + ls_proc_type + "' and route_set = '" + ls_route_set + "' and inch = '" + ls_inch + "' and cost_item_gp = '" + ls_cost_item_gp + "' ";
                        QueryExecute(sql, "");

                        sql = "insert into FourM_B1004_NEPES values('" + ls_proc_type + "','" + ls_route_set + "','" + ls_inch + "' ,'" + ls_cost_item_gp + "'  ) ";
                        if (QueryExecute(sql, "") < 1)
                        {
                            string script = "alert(\" 품목그룹 " + ls_cost_item_gp + " 저장시 오류가 발생하였습니다.\");";
                            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                            GridView1.DataSource = null;
                            GridView1.DataBind();
                            return;
                        }

                    }
                    
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert(\" 저장되었습니다.\");", true);
                    GridView1.DataSource = null;
                    GridView1.DataBind();
                    lbl_bas_info_title.Text = "프로세스별 품목그룹";
                    btnMoveDataNextMonth.Visible = false;
                    Panel1.Visible = true;
                    Panel2.Visible = false;
                }
            //}

        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            Panel1.Visible = true;
            Panel2.Visible = false;
            GridView1.DataSource = null;
            GridView1.DataBind();
        }
        //전월복사 버튼 클릭시..
        protected void btnMoveDataNextMonth_Click(object sender, EventArgs e)
        {
            if (tb_fr_dt.Text == "" || tb_fr_dt.Text == null)
            {
                string script = "alert(\"당월을 선택하여 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                tb_fr_dt.Focus();
            }
            else
            {
                now_month = tb_fr_dt.Text + "01" ;
                
                //전월생성
                before_month = DateTime.ParseExact(now_month, "yyyyMMdd", null).AddMonths(0).AddDays(-1).ToString("yyyyMM");

                // 코스트센터 그룹 전월카피
                if (DropDownList1.SelectedValue == "A")
                {
                    ////선택한 월의 데이타가 있으면 지우고 입력한다.
                    sql = "select count(yyyymm) from FourM_B1001_NEPES where yyyymm = '" + tb_fr_dt.Text + "' ";
                    if (QueryExecute(sql, "check") > 0)
                    {   //삭제진행
                        sql = " delete FourM_B1001_NEPES  where yyyymm = '" + tb_fr_dt.Text + "' ";
                        QueryExecute(sql, "");
                    }
                    // 데이타 insert 쿼리
                    sql = " insert into FourM_B1001_NEPES select '" + tb_fr_dt.Text + "', cost_cd, cost_gp_cd from FourM_B1001_NEPES where yyyymm =  '" + before_month + "' ";                    
                }

                // 코스트센터 그룹별 품목별 가중치 등록
                if (DropDownList1.SelectedValue == "B")
                {
                    ////선택한 월의 데이타가 있으면 지우고 입력한다.
                    sql = "select count(yyyymm) from FourM_B1002_NEPES where yyyymm = '" + tb_fr_dt.Text + "' ";
                    if (QueryExecute(sql, "check") > 0)
                    {   //삭제진행
                        sql = " delete FourM_B1002_NEPES  where yyyymm = '" + tb_fr_dt.Text + "' ";
                        QueryExecute(sql, "");
                    }
                    // 데이타 insert 쿼리
                    sql = " insert into FourM_B1002_NEPES select '" + tb_fr_dt.Text + "', cost_gp_cd, item_gp_cd, weight from FourM_B1002_NEPES where yyyymm =  '" + before_month + "' ";
                }

                if (QueryExecute(sql, "") > 0)
                {
                    string script = "alert(\"복사되었습니다.\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    btn_old_request__Click(null, null);
                }
                else
                {
                    string script = "alert(\"데이타 복사시 문제가 발생하였습니다. 관리자에게 문의해 부세요.\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                }
            }
            Panel1.Visible = true;
            Panel2.Visible = false;

        }

        private int QueryExecute(string sql, string wk_type)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                //삭제시 기존 권한아이디에 프로그램이 연결되었는지 확인하기 위함.
                if (wk_type == "check")
                    value = Convert.ToInt32(cmd.ExecuteScalar());
                else
                    value = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }

            conn.Close();
            return value;
        }

        protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (RadioButtonList1.SelectedValue == "view3")
            {
                if (DropDownList1.SelectedValue == "A")
                {
                    lbl_bas_info_title.Text = "CostCenter 그룹";
                    btnMoveDataNextMonth.Visible = true;
                    Panel1.Visible = true;
                    Panel2.Visible = false;
                }

                if (DropDownList1.SelectedValue == "B")
                {
                    lbl_bas_info_title.Text = "품목그룹별 가중치";
                    btnMoveDataNextMonth.Visible = true;
                    Panel1.Visible = true;
                    Panel2.Visible = false;
                }

                if (DropDownList1.SelectedValue == "C")
                {
                    lbl_bas_info_title.Text = "디바이스별 품목그룹";
                    btnMoveDataNextMonth.Visible = false;
                    Panel1.Visible = true;
                    Panel2.Visible = false;
                }

                if (DropDownList1.SelectedValue == "D")
                {
                    lbl_bas_info_title.Text = "프로세스별 품목그룹";
                    btnMoveDataNextMonth.Visible = false;
                    Panel1.Visible = false;
                    Panel2.Visible = false;
                }

                if (DropDownList1.SelectedValue == "E")
                {
                    lbl_bas_info_title.Text = "프로세스별 품목그룹";
                    btnMoveDataNextMonth.Visible = false;
                    Panel1.Visible = true;
                    Panel2.Visible = false;
                }

                ReportViewer1.Reset();
                
            }

        }

        protected void btn_grid_delete_Click(object sender, EventArgs e)
        {
            Button btn_show_paper_cl = sender as Button;

            GridViewRow row = btn_show_paper_cl.NamingContainer as GridViewRow;
            int rowindex = row.RowIndex;
            GridView1.Rows[rowindex].Visible = false;
            
            Panel1.Visible = false;
            Panel2.Visible = true;
        } 

       
    }
}