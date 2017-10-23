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

namespace ERPAppAddition.ERPAddition.AM.AM_A2001
{
    public partial class AM_A2001 : System.Web.UI.Page
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

        protected void btn_retrive_Click(object sender, EventArgs e)
        {
            string sql;
            string yyyymm = tb_yyyymm.Text;
            if (yyyymm == null || yyyymm == "" || yyyymm == "%")
            {
                MessageBox.ShowMessage("년월을 입력해주세요.", this.Page);
                FpSpread1.Focus();
                return;
            }

            sql = "select a.yyyymm, a.plant_cd, b.plant_nm, a.acct_cd,   " +
                  "       case when a.acct_cd = '430027' then '수선유지비'  " + //기존: 1
                  "            when a.acct_cd = '430037' then '소모품비'  " +   //기존: 2
                  "            when a.acct_cd = '430037' then '지급수수료' end acct_nm " + //기존: 3
                  "      , a.amt " +
                  " from am_a2001 a inner join b_plant b on a.plant_cd = b.plant_cd " +
                  " where a.yyyymm = '" + tb_yyyymm.Text + "' " +
                  "   and a.plant_cd like '" + ddl_plant.SelectedValue + "' " +
                  "   and a.acct_cd like '" + ddl_acct_cd.SelectedValue + "' ";

            sqlAdapter1 = new SqlDataAdapter(sql, conn);

            sqlAdapter1.Fill(ds, "ds");

            FpSpread1.DataSource = ds;
            FpSpread1.DataBind();



        }

        protected void btn_save_Click(object sender, EventArgs e)
        {
            FpSpread1.SaveChanges();
            //FpSpread1.Reset();
            MessageBox.ShowMessage("저장되었습니다.", this.Page);
            btn_retrive_Click(this, new EventArgs());

        }
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

        protected void btn_rowadd_Click(object sender, EventArgs e) //Row추가
        {
            if (tb_rowcnt.Text == null || tb_rowcnt.Text == "")
            {
                MessageBox.ShowMessage("추가할 Row수를 입력해주세요.", this.Page);
                tb_rowcnt.Focus();
                return;
            }
            else
            {
                FpSpread1.Sheets[0].AddRows(FpSpread1.Sheets[0].RowCount, Convert.ToInt16(tb_rowcnt.Text)); //행 만들기
              
                for (int i = 0; i <= FpSpread1.Rows.Count - 1;   i++) //for문 추가
                {
                    string yyyymm = tb_yyyymm.Text;
                    string plant = ddl_plant.SelectedValue;
                    string plant_nm = ddl_plant.Items[ddl_plant.SelectedIndex].Text;
                    string acct_cd = ddl_acct_cd.SelectedValue;
                    string acct_nm = ddl_acct_cd.Items[ddl_acct_cd.SelectedIndex].Text;

                  // int i = FpSpread1.Sheets[0].RowCount - 1;

                    FpSpread1.Sheets[0].Cells[i, 0].Text = yyyymm;
                    FpSpread1.Sheets[0].Cells[i, 1].Text = plant;
                    FpSpread1.Sheets[0].Cells[i, 2].Text = plant_nm;
                    FpSpread1.Sheets[0].Cells[i, 3].Text = acct_cd;
                    FpSpread1.Sheets[0].Cells[i, 4].Text = acct_nm;
                }

            }
        }




        protected void btn_insert_Click(object sender, EventArgs e) //입력버튼
        {
            //추가(2014.03.02)

            string yyyymm = tb_yyyymm.Text;
            string plant = ddl_plant.SelectedValue;
            string plant_nm = ddl_plant.Items[ddl_plant.SelectedIndex].Text; //추가: 선택된 공장코드에 해당하는 공장명 받아오기
            string acct_cd = ddl_acct_cd.SelectedValue;
            string acct_nm = ddl_acct_cd.Items[ddl_acct_cd.SelectedIndex].Text;//추가: 선택된 계정코드에 해당하는 계정코드 받아오기
            string amt = "0"; // 최초 금액은 0으로 setup


            if (yyyymm == null || yyyymm == "")
                MessageBox.ShowMessage("년월을 입력해주세요.", this.Page);
            else if (plant == null || plant == "")
                MessageBox.ShowMessage("공장코드를 입력해주세요.", this.Page);
            else if (acct_cd == null || acct_cd == "")
                MessageBox.ShowMessage("비용계정코드를 입력해주세요.", this.Page);
            else if (amt == null || amt == "")
                MessageBox.ShowMessage("금액을 입력해주세요.", this.Page);

            else
            {
                int rowIndex = FpSpread1.Sheets[0].RowCount;

                FpSpread1.Sheets[0].AddRows(FpSpread1.Sheets[0].RowCount, 1);

                FpSpread1.Sheets[0].Cells[rowIndex, 0].Text = yyyymm;
                FpSpread1.Sheets[0].Cells[rowIndex, 1].Text = plant;
                FpSpread1.Sheets[0].Cells[rowIndex, 2].Text = plant_nm;
                FpSpread1.Sheets[0].Cells[rowIndex, 3].Text = acct_cd;
                FpSpread1.Sheets[0].Cells[rowIndex, 4].Text = acct_nm;
                FpSpread1.Sheets[0].Cells[rowIndex, 5].Text = amt;

                string sql = "insert into am_a2001 ";
                sql = sql + "values('" + Convert.ToString(yyyymm) + "','" + plant + "','" + acct_cd + "'," + Convert.ToDecimal(amt) + ", 'unierp', getdate(), 'unierp', getdate())";
                QueryExecute(sql, "");
                tb_yyyymm.Text = Convert.ToString(yyyymm);
            }


        }

        protected void btn_delete_Click(object sender, EventArgs e)
        {
            System.Collections.IEnumerator enu = FpSpread1.ActiveSheetView.SelectionModel.GetEnumerator();
            FarPoint.Web.Spread.Model.CellRange cr;

            while (enu.MoveNext())
            {
                cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                int a = FpSpread1.Sheets[0].ActiveRow;
                //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
                for (int i = 0; i < cr.RowCount; i++)
                {
                    string yyyymm = FpSpread1.Sheets[0].Cells[FpSpread1.Sheets[0].ActiveRow, 0].Text;
                    string plant = FpSpread1.Sheets[0].Cells[FpSpread1.Sheets[0].ActiveRow, 1].Text;
                    string acct_cd = FpSpread1.Sheets[0].Cells[FpSpread1.Sheets[0].ActiveRow, 3].Text;


                    string sql = "delete am_a2001 ";
                    sql = sql + " where yyyymm  ='" + Convert.ToString(yyyymm) + "' and plant_cd = '" + plant + "'  and acct_cd = '" + acct_cd + "' ";

                    if (QueryExecute(sql, "") > 0)
                        FpSpread1.Sheets[0].Rows.Remove(FpSpread1.Sheets[0].ActiveRow, 1);

                    tb_yyyymm.Text = Convert.ToString(yyyymm);
                }
            }

            MessageBox.ShowMessage("삭제되었습니다.", this.Page);
            btn_retrive_Click(null, null);
        }

        protected void FpSpread1_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {

            int colcnt;
            int i;
            string cg_yyyymm, cg_plant_cd, cg_acct_cd, cg_amt;
            //string yyyymm, plant_cd, acct_cd, amt;
            int r = (int)e.CommandArgument;
            colcnt = e.EditValues.Count - 1;


            for (i = 0; i <= colcnt; i++)
            {
                if (!object.ReferenceEquals(e.EditValues[i], FarPoint.Web.Spread.FpSpread.Unchanged))
                {
                    string sql;

                    /*기존값 가져오기*/
                    string yyyymm = FpSpread1.Sheets[0].Cells[r, 0].Value.ToString();
                    string plant_cd = FpSpread1.Sheets[0].Cells[r, 1].Value.ToString();
                    string acct_cd = FpSpread1.Sheets[0].Cells[r, 3].Value.ToString();
                    string amt = FpSpread1.Sheets[0].Cells[r, 5].Value.ToString();

                    /*변경된값 가져오기*/

                    if (i == 0)
                        cg_yyyymm = e.EditValues[0].ToString();
                    else
                        cg_yyyymm = yyyymm;

                    if (i == 1)
                        cg_plant_cd = e.EditValues[1].ToString();
                    else
                        cg_plant_cd = plant_cd;

                    if (i == 3)
                        cg_acct_cd = e.EditValues[3].ToString();
                    else
                        cg_acct_cd = acct_cd;

                    if (i == 5)
                        cg_amt = e.EditValues[5].ToString();
                    else
                        cg_amt = amt;

                    sql = "update am_a2001 ";
                    sql = sql + "set yyyymm = '" + cg_yyyymm + "',plant_cd = '" + cg_plant_cd + "',acct_cd = '" + cg_acct_cd + "', amt = " + Convert.ToDecimal(cg_amt) + ", updt_dt =  getdate()";
                    sql = sql + " where yyyymm = '" + yyyymm + "' and  plant_cd = '" + plant_cd + "' and acct_cd = '" + acct_cd + "' ";
                    QueryExecute(sql, "");
                }
            }
        }
  
        protected void FpSpread1_ActiveRowChanged(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {

        }


    }
}