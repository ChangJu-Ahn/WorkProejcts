using System;
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

namespace ERPAppAddition.ERPAddition.POPUP
{
    public partial class pop_item_cd : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["erp_db"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        protected void Page_Load(object sender, EventArgs e)
        {
            //page load시 공장정보를 보여준다.
            tb_plant_cd.Text = Request.QueryString["ControlVal"].ToString();
            string sql_plant_nm = "SELECT PLANT_NM FROM B_PLANT WHERE VALID_TO_DT >= GETDATE() and PLANT_CD = '" + tb_plant_cd.Text + "'";

            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql_plant_nm;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                tb_plant_nm.Text = dr[0].ToString();
            }

            dr.Close();
            conn.Close();

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string item_cd, item_nm, acct_cd, plant_cd;
            item_cd = tb_pop_item_cd.Text.Trim();
            item_nm = tb_pop_item_nm.Text.Trim();
            acct_cd = dl_pop_item_acct.Text.Trim();
            plant_cd = tb_plant_cd.Text.Trim();

            //string sql = "select top 10 emp_no, name from haa010t ";
            string var1 = "";
            var1 = var1 + "SELECT b.item_cd, " + "\n";
            var1 = var1 + "       b.item_nm, " + "\n";
            var1 = var1 + "       b.spec, " + "\n";
            var1 = var1 + "       b.basic_unit, " + "\n";
            var1 = var1 + "       dbo.Ufn_getcodename('P1001', a.item_acct)   item_acct_nm " + "\n";
            var1 = var1 + "FROM   b_item_by_plant a, " + "\n";
            var1 = var1 + "       b_item b, " + "\n";
            var1 = var1 + "       b_item_acct_inf c " + "\n";
            var1 = var1 + "WHERE  a.item_cd = b.item_cd " + "\n";
            var1 = var1 + "       AND a.item_acct = c.item_acct " + "\n";
            var1 = var1 + "       AND a.item_acct = '" + acct_cd + "' " + "\n";
            var1 = var1 + "       AND a.plant_cd = '" + plant_cd + "' " + "\n";
            var1 = var1 + "       AND a.item_cd >= '" + item_cd + "' " + "\n";
            var1 = var1 + "       AND b.item_nm LIKE '" + '%' + item_nm + '%' + "' " + "\n";
            var1 = var1 + "       AND b.item_nm >= '' " + "\n";
            var1 = var1 + "       AND ( b.item_class >= '' " + "\n";
            var1 = var1 + "             AND b.item_class <= 'zzzzzzzzzzzz' " + "\n";
            var1 = var1 + "              OR b.item_class IS NULL ) " + "\n";
            var1 = var1 + "       AND c.item_acct_group >= '' " + "\n";
            var1 = var1 + "       AND c.item_acct_group <= 'zz' " + "\n";
            var1 = var1 + "       AND a.procur_type >= '' " + "\n";
            var1 = var1 + "       AND a.procur_type <= 'zz' " + "\n";
            var1 = var1 + "       AND a.prod_env >= '' " + "\n";
            var1 = var1 + "       AND a.prod_env <= 'zz' " + "\n";
            var1 = var1 + "       AND a.valid_to_dt >= Getdate() " + "\n";
            var1 = var1 + "       AND b.spec LIKE '%%' " + "\n";
            var1 = var1 + "       AND a.tracking_flg LIKE '%' " + "\n";
            var1 = var1 + "ORDER  BY a.item_cd, " + "\n";
            var1 = var1 + "          b.item_nm ";

            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = var1;
            dr = cmd.ExecuteReader();
            pop_gridview1.DataSource = dr;
            pop_gridview1.DataBind();
            dr.Close();
            conn.Close();
        }

    }
}