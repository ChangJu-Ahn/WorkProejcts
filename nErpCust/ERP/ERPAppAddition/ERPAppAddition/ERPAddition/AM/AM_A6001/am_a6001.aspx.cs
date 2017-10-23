using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.PM.p1401ma6_nepes;
using FarPoint.Web.Spread;

namespace ERPAppAddition.ERPAddition.AM.AM_A6001
{
    public partial class am_a6001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_test1"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        SqlDataAdapter sqlAdapter1;
        DataSet ds = new DataSet();
        string ls_fr_dt, ls_to_dt, ls_biz_area_cd = "", ls_ctrl_cd = "";
        //string ls_tbl_id, ls_data_colm_id, ls_data_coml_nm, ls_major_cd, ls_key_colm_id2;
        string ls_msg_cd = "", ls_sp_id = "110";
        string ls_biz_area_cd_sql, ls_ctrl_cd_sql, ls_sql;
        int value;
        //string insert_yn = "N";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ReportViewer1.Reset();
                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        public DataTable QueryExeuteDT(string sql)
        {

            ds_am_a6001 ds = new ds_am_a6001();

            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;
            DataTable dt = new DataTable();

            try
            {
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                conn.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }

            return ds.Tables[0];
            //return ds.Tables[0];
        }


        protected void rbl_view_type_SelectedIndexChanged1(object sender, EventArgs e)
        {

            if (rbl_view_type.SelectedValue == "A") //코스트센터 연결 선택
            {
                Panel1.Visible = false; //등록 Pannel
                Panel_costset.Visible = true; //분류 Pannel

            

            }
            if (rbl_view_type.SelectedValue == "B") //등록 선택
            {
                Panel1.Visible = true; //등록 Pannel
                Panel_costset.Visible = false; //분류 Pannel
                btn_view.Visible = false; //조회버튼
                btn_save.Visible = true; //저장버튼
                btn_delete0.Visible = true; //삭제버튼

            }
            if (rbl_view_type.SelectedValue == "C") //조회 선택 
            {
                Panel1.Visible = true; //등록 Pannel
                Panel_costset.Visible = false; //분류 Pannel
                btn_save.Visible = false;//저장버튼
                btn_delete0.Visible = false; //삭제 버튼
                Label6.Visible = false; //금액 label
                tb_amt.Visible = false; //금액 Textbox
                btn_view.Visible = true; //조회버튼
                Panel_report.Visible = true;
            }

        }

      
        public int QueryExecute(SqlConnection connection, string sql, string wk_type)
        {

            connection.Open();
            cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {

                if (wk_type == "check")
                    value = Convert.ToInt32(cmd.ExecuteScalar());
                else
                    value = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                value = -1;
            }

            connection.Close();
            return value;
        }

        public DataTable QueryExeuteDT(SqlConnection connection, string sql)
        {
            ds_am_a6001 ds = new ds_am_a6001();

            connection.Open();
            cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                connection.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    connection.Close();
            }

            return ds.Tables[0];
        }

        //***************************대분류-소분류 품목그룹 연결화면- 조회버튼******************************

        protected void btn_itemgp_costset_Click(object sender, EventArgs e)
        {
            string sql;
            lsb_l_costset.Items.Clear(); //내용지우기
            lsb_r_costset.Items.Clear(); //내용지우기

            if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("대분류 품목그룹을선택해주세요.", this.Page);
            }
            else
            {
                //왼쪽 미등록소분류품목 가져오기
                sql = "select DISTINCT COST_NM,COST_CD " +
                      "from B_COST_CENTER where COST_NM NOT IN (SELECT COST_NM  FROM AM_A2001_COST where COST_GP_NM = '" + ddl_itemgp.SelectedValue + "' ) ";
                
                DataTable dt = QueryExeuteDT(conn, sql);
               if (dt.Rows.Count > 0)
                {

                    dt.Columns.Add("DisplayField", typeof(string), "COST_NM + ' [' + COST_CD + ']'");
                    lsb_l_costset.DataSource = dt;
                    lsb_l_costset.DataTextField = "DisplayField";
                  //  lsb_l_costset.DataTextField = "COST_NM";
                    lsb_l_costset.DataValueField = "COST_NM";
                    lsb_l_costset.DataBind();

                }
                //오른쪽 등록되어있는 소분류품복 가져오기
               sql = "select COST_NM,COST_CD from AM_A2001_COST where COST_GP_NM = '" + ddl_itemgp.SelectedValue + "'";
               dt = QueryExeuteDT(conn, sql);
               if (dt.Rows.Count > 0)
                 {
                    dt.Columns.Add("DisplayField", typeof(string), "COST_NM + ' [' + COST_CD + ']'");
                    lsb_r_costset.DataSource = dt;
                    lsb_r_costset.DataTextField = "DisplayField";
                    //  lsb_l_costset.DataTextField = "COST_NM";
                    lsb_r_costset.DataValueField = "COST_NM";
                    lsb_r_costset.DataBind();
                }
            }
        }

        protected void btn_move_right_Click1(object sender, EventArgs e)
        {

            if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을선택해주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_l_costset.Items.Count; i++)
                {
                    if (this.lsb_l_costset.Items[i].Selected)
                    {
                        this.lsb_r_costset.Items.Add(this.lsb_l_costset.Items[i]);

                        //this.lsb_r_costset.Items.Add(this.lsb_l_costset.Items[i]);
                        //System.Web.UI.WebControls.ListItem listItem = this.lsb_l_costset.Items[i];
                        string code = lsb_l_costset.Items[i].Text;
                        string[] codea = code.Split('[');
                        string name = codea[0].Trim();
                        string codeb = codea[1].Replace("]","").Trim();
                        //// $Name ($Code)
                        //string displayText = listItem.Text;
                        //string displayCode = "(" + listItem.Value + ")";
                        //int codeStartIndex = displayText.LastIndexOf(displayCode);

                        //string name = displayText.Substring(0, codeStartIndex - 1);
                        //string code = listItem.Value;

                        
                        //선택된품목그룹에왼쪽리스트박스내용을insert한다. 
                        string sql = "insert into AM_A2001_COST " +
                                     "values ('" + ddl_itemgp.SelectedValue + "', '" + codeb + "','" + name + "','unierp', getdate(), 'unierp', getdate()) ";

                        if (QueryExecute(conn, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타저장에실패했습니다.", this.Page);
                        this.lsb_l_costset.Items.Remove(this.lsb_l_costset.Items[i]);
                        i--;
                    }
                }

                btn_itemgp_costset_Click(null, null);
            }
        }
        protected void btn_move_left_Click1(object sender, EventArgs e)
        {
            if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을선택해주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_r_costset.Items.Count; i++)
                {
                    if (this.lsb_r_costset.Items[i].Selected)
                    {
                        this.lsb_l_costset.Items.Add(this.lsb_r_costset.Items[i]);
                       
                        string code = lsb_l_costset.Items[i].Text;
                        string[] codea = code.Split('(');
                        string name = codea[0].Trim();
                        string codeb = codea[1].Replace(")","").Trim();
                        //선택된품목그룹에왼쪽리스트박스내용을insert한다. 
                        string sql = "delete AM_A2001_COST " +
                                     "where cost_gp_nm = '" + ddl_itemgp.SelectedValue + "' and cost_nm =  '" + name + "' and and cost_cd =  '" + codeb + "'";
                       
                        if (QueryExecute(conn, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타저장에실패했습니다.", this.Page);

                        this.lsb_r_costset.Items.Remove(this.lsb_r_costset.Items[i]);
                        i--;

                    }
                }
                btn_itemgp_costset_Click(null, null);
            }
        }

        protected void btn_save_Click1(object sender, EventArgs e)
        {

            if (tb_yyyymm1 == null || tb_yyyymm1.Text.Equals(""))
            {
                MessageBox.ShowMessage("년월을 입력하세요.", this.Page);

                return;
            }
            if (ddl_item_gp.Text.Equals("-선택안됨-"))
            {
                MessageBox.ShowMessage("품목그룹 입력하세요.", this.Page);

                return;
            }
            if (ddl_acct.Text.Equals("-선택안됨-"))
            {
                MessageBox.ShowMessage("비용계정을 입력하세요.", this.Page);

                return;
            }
            if (tb_amt == null || tb_yyyymm1.Text.Equals(""))
            {
                MessageBox.ShowMessage("금액을 입력하세요.", this.Page);

                return;
            }
            conn.Open();

            string queryStr = "insert into am_a2001_test2(yyyymm,item_gp,acct_cd,amt,isrt_id,isrt_dt,updt_id,updt_dt)";
            queryStr += " values('" + tb_yyyymm1.Text + "', '" + ddl_item_gp.Text + "','" + ddl_acct.SelectedValue + "','" + tb_amt.Text + "',";
            queryStr += "'unierp','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','unierp','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";



            SqlCommand sComm = new SqlCommand(queryStr, conn);
            MessageBox.ShowMessage("저장되었습니다.", this.Page);

            sComm.ExecuteNonQuery();
            conn.Close();
        }

        //FpSpread1.SaveChanges();
        ////FpSpread1.Reset();
        //MessageBox.ShowMessage("저장되었습니다.", this.Page);
        //btn_view_Click_Click(this, new EventArgs());


        public int QueryExecute(string sql, string wk_type)
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
                value = -1;
            }

            conn.Close();
            return value;
        }


       protected void btn_delete0_Click(object sender, EventArgs e) //삭제버튼 클릭
        {
            conn.Open();

            string queryStr = "Delete from am_a2001_test2 where yyyymm='" + tb_yyyymm1.Text + "' and item_gp = '" + ddl_item_gp.Text + "' and acct_cd= '" + ddl_acct.SelectedValue + "'";

            SqlCommand sComm = new SqlCommand(queryStr, conn);
            MessageBox.ShowMessage("삭제되었습니다.", this.Page);
            sComm.ExecuteNonQuery();
            conn.Close();
        }

       

        protected void btn_view_Click(object sender, EventArgs e) //조회버튼 클릭
     
        {
        ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "dbo.usp_am2001_view";
            cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@yyyymm", SqlDbType.VarChar, 12);
            SqlParameter param2 = new SqlParameter("@item_gp", SqlDbType.VarChar, 20);
            SqlParameter param3 = new SqlParameter("@acct_cd", SqlDbType.VarChar, 2);
           

            param1.Value = tb_yyyymm1.Text;
            param2.Value = ddl_item_gp.Text;
            param3.Value = ddl_acct.Text;
          

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
          
            

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                
                ReportViewer1.LocalReport.ReportPath = Server.MapPath("am_a6001.rdlc");
                ReportViewer1.LocalReport.DisplayName = "구매집계예산등록" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                ReportViewer1.LocalReport.Refresh();

                //UpdatePanel1.Update();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
 

        //    System.Collections.IEnumerator enu = FpSpread1.ActiveSheetView.SelectionModel.GetEnumerator();
        //    FarPoint.Web.Spread.Model.CellRange cr;

        //    while (enu.MoveNext())
        //    {
        //        cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
        //        int a = FpSpread1.Sheets[0].ActiveRow;
        //        //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
        //        for (int i = 0; i < cr.RowCount; i++)
        //        {
        //            string yyyymm = FpSpread1.Sheets[0].Cells[FpSpread1.Sheets[0].ActiveRow, 0].Text;
        //            string plant_cd = FpSpread1.Sheets[0].Cells[FpSpread1.Sheets[0].ActiveRow, 1].Text;
        //            string acct_cd = FpSpread1.Sheets[0].Cells[FpSpread1.Sheets[0].ActiveRow, 3].Text;


        //            string sql = "delete AM_A2001_COST ";
        //            sql = sql + " where yyyymm  ='" + Convert.ToString(yyyymm) + "' and plant_cd = '" + plant_cd + "'  and acct_cd = '" + acct_cd + "' ";

        //            if (QueryExecute(sql, "") > 0)
        //                FpSpread1.Sheets[0].Rows.Remove(FpSpread1.Sheets[0].ActiveRow, 1);

        //            tb_yyyymm.Text = Convert.ToString(yyyymm);
        //        }
        //    }

        //    MessageBox.ShowMessage("삭제되었습니다.", this.Page);
        //    btn_view_Click_Click(null, null);
        //}

        //protected void FpSpread1_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        //{

        //    int colcnt;
        //    int i;
        //    string cg_yyyymm, cg_item_gp, cg_acct_cd,cg_acct_nm,cg_amt;
        //    string yyyymm, item_gp, acct_cd,acct_nm,amt;
        //    int r = (int)e.CommandArgument;
        //    colcnt = e.EditValues.Count - 1;



        //    for (i = 0; i <= colcnt; i++)
        //    {
        //        if (!object.ReferenceEquals(e.EditValues[i], FarPoint.Web.Spread.FpSpread.Unchanged))
        //        {
        //            string sql;

        //            //업데이트시
        //            if (FpSpread1.ActiveSheetView.ActiveRow == r && FpSpread1.Sheets[0].Cells[r, 0].Value != null)
        //            {
        //                /*기존값 가져오기*/
        //                yyyymm = FpSpread1.Sheets[0].Cells[r, 0].Value.ToString();
        //                item_gp = FpSpread1.Sheets[0].Cells[r, 1].Value.ToString();
        //                acct_cd = FpSpread1.Sheets[0].Cells[r, 2].Value.ToString();
        //                acct_nm = FpSpread1.Sheets[0].Cells[r, 3].Value.ToString();
        //                amt = FpSpread1.Sheets[0].Cells[r, 4].Value.ToString();


        //                /*변경된값 가져오기*/

        //                if (i == 0)
        //                    cg_yyyymm = e.EditValues[0].ToString();
        //                else
        //                    cg_yyyymm = yyyymm;

        //                if (i == 1)
        //                    cg_item_gp = e.EditValues[1].ToString();
        //                else
        //                    cg_item_gp = item_gp;

        //                if (i == 2)
        //                    cg_acct_cd = e.EditValues[3].ToString();
        //                else
        //                    cg_acct_cd = acct_cd;

        //                if (i == 3)
        //                    cg_acct_nm = e.EditValues[4].ToString();
        //                else
        //                    cg_acct_nm = acct_nm;
        //                if (i == 4)
        //                    cg_amt = e.EditValues[5].ToString();
        //                else
        //                    cg_amt = amt;

        //                sql = "update am_a2001_test2 ";
        //                sql = sql + "set yyyymm = '" + cg_yyyymm + "',item_gp = '" + cg_item_gp + "',acct_cd = '" + cg_acct_cd + "', acct_nm = '" + cg_acct_nm + "',amt = " + Convert.ToDecimal(cg_amt) + ", updt_dt =  getdate()";
        //                sql = sql + " where yyyymm='" + tb_yyyymm1.Text + "' and item_gp = '" + ddl_item_gp.Text + "' and acct_cd= '" + ddl_acct.SelectedValue + "'";
        //                QueryExecute(sql, "");
        //            }
        //            else
        //            {
        //                //r = r + 1;
        //                //int j = FpSpread1.Sheets[0].ColumnCount;
        //                yyyymm = e.EditValues[0].ToString();//FpSpread1.Sheets[0].Cells[r, 0].Text;
        //                item_gp = e.EditValues[1].ToString();
        //                acct_cd = e.EditValues[3].ToString();
        //                amt = e.EditValues[5].ToString();
        //                if (yyyymm == null || yyyymm == "")
        //                    MessageBox.ShowMessage("년월을 입력해주세요.", this.Page);
        //                else if (item_gp == null || item_gp == "")
        //                    MessageBox.ShowMessage("품목그룹 선택해주세요.", this.Page);
        //                else if (acct_cd == null || acct_cd == "")
        //                    MessageBox.ShowMessage("비용계정코드를 선택해주세요.", this.Page);
        //                else if (amt == null || amt == "")
        //                    MessageBox.ShowMessage("금액을 입력해주세요.", this.Page);
        //                else
        //                {
        //                    sql = "insert into am_a2001_test2 ";
        //                    sql = sql + "values('" + Convert.ToString(yyyymm) + "','" + item_gp + "','" + acct_cd + "'," + Convert.ToDecimal(amt) + ", 'unierp', getdate(), 'unierp', getdate())";
        //                    QueryExecute(sql, "");
        //                    tb_yyyymm.Text = Convert.ToString(yyyymm);
        //                }
        //            }
        //        }
        //    }
        //}

        //protected void btn_view1_Click(object sender, EventArgs e)
        //{

        //    string sql;
        //    string yyyymm = tb_yyyymm1.Text;
        //    if (yyyymm == null || yyyymm == "" || yyyymm == "%")
        //    {
        //        MessageBox.ShowMessage("년월을 입력해주세요.", this.Page);
        //        FpSpread1.Focus();
        //        return;
        //    }

        //    sql = "select yyyymm, item_gp, acct_cd, " +
        //          "       case when acct_cd = '1' then '수선유지비'  " +
        //          "            when acct_cd = '2' then '소모품비'  " +
        //          "            when acct_cd = '3' then '지급수수료' end acct_nm," +
        //          "amt " +
        //          " from am_a2001_test2 " +
        //          " where yyyymm = '" + tb_yyyymm.Text + "' " +
        //          "   and item_gp like '" + ddl_item_gp.Text + "' " +
        //          "   and acct_cd like '" + ddl_acct.SelectedValue + "' ";

        //    sqlAdapter1 = new SqlDataAdapter(sql, conn);

        //    sqlAdapter1.Fill(ds, "ds");

        //    FpSpread1.DataSource = ds;
        //    FpSpread1.DataBind();



    }

















}