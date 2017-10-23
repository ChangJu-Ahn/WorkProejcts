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


namespace ERPAppAddition.ERPAddition.AM.AM_A1001
{
    public partial class am_a1001 : System.Web.UI.Page
    {

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_display"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        
        string ls_fr_dt, ls_to_dt, ls_biz_area_cd = "", ls_ctrl_cd = "";
        //string ls_tbl_id, ls_data_colm_id, ls_data_coml_nm, ls_major_cd, ls_key_colm_id2;
        string ls_msg_cd="", ls_sp_id="110";
        string ls_biz_area_cd_sql, ls_ctrl_cd_sql, ls_sql;
        
        protected void Page_Load(object sender, EventArgs e)
        {
            ls_biz_area_cd_sql = "SELECT BIZ_AREA_CD , BIZ_AREA_NM FROM B_BIZ_AREA union all SELECT '', '전체' ORDER BY BIZ_AREA_CD ";
            ls_ctrl_cd_sql = "SELECT CTRL_CD ,  CTRL_NM  FROM A_CTRL_ITEM ORDER BY CTRL_NM";
            string db_name = String.Empty;
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["db"] != null && Request.QueryString["db"].ToString() != "")
                {
                    db_name = Request.QueryString["db"].ToString();
                    if (db_name.Length > 0)
                    {
                        conn = new SqlConnection(ConfigurationManager.ConnectionStrings[db_name].ConnectionString);
                    }
                    
                }
            }
            
            
            ls_sql = "";
            ds_am_a1001 dt1 = new ds_am_a1001();
            ReportCreator(dt1, ls_sql, ReportViewer1, "rv_am_a1001.rdlc", "DataSet1");

            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void bt_retrive_Click(object sender, EventArgs e)
        {
            ls_fr_dt = tb_fr_dt.Text.Trim();
            ls_to_dt = tb_to_dt.Text.Trim();
            ls_biz_area_cd = ddl_biz_area.Text.Trim();
                       
            ls_ctrl_cd = ddl_ctrl_value.Text.Trim();
            
            // 선택된 관리항목코드로 테이블 , 칼럼id등을 가져온다. 
            //string chk_table_sql = "Select TBL_ID,DATA_COLM_ID,DATA_COLM_NM, MAJOR_CD, key_colm_id2 From  A_CTRL_ITEM Where  CTRL_CD = '"+ls_ctrl_cd+"'";
            //conn.Open();
            //cmd = conn.CreateCommand();
            //cmd.CommandType = CommandType.Text;
            //cmd.CommandText = chk_table_sql;
            //dr = cmd.ExecuteReader();

            //while (dr.Read())
            //{
            //    ls_tbl_id = dr[0].ToString(); //테이블아이디
            //    ls_data_colm_id = dr[1].ToString(); //칼럼아이디
            //    ls_data_coml_nm = dr[2].ToString(); //칼럼명
            //    ls_major_cd = dr[3].ToString(); //major_cd
            //    ls_key_colm_id2 = dr[4].ToString(); //key_colm_id2
            //}

            //dr.Close();
            //conn.Close();
            ////관리항목명데이타를 가져오는 쿼리 만들기
            //string ctrl_val_sql ;
            //if (string.Compare(ls_tbl_id, "B_MINOR") == 0)
            //{
            //    ctrl_val_sql = "select " + ls_data_coml_nm + " from " + ls_tbl_id + " where major_cd = '" + ls_major_cd + "' and " + ls_data_colm_id + " =  ";
            //}
            //else
            //{
            //    if (string.Compare(ls_tbl_id, "B_BIZ_PARTNER") == 0)
            //    {
            //        ctrl_val_sql = "select " + ls_data_coml_nm + " from " + ls_tbl_id + " where '" + ls_key_colm_id2 + "' and  " + ls_data_colm_id + " =  ";
            //    }
            //    else
            //    {
            //        ctrl_val_sql = "select " + ls_data_coml_nm + " from " + ls_tbl_id + " where '" + ls_data_colm_id + "' =  ";
            //    }
                
            //}
                
           
            //손익계산 프로시져 실행
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "usp_a_pl_ctrl_val";
            SqlParameter param1 = new SqlParameter("@this_from_mnth", SqlDbType.VarChar, 6);
            SqlParameter param2 = new SqlParameter("@this_to_mnth", SqlDbType.VarChar, 6);
            SqlParameter param3 = new SqlParameter("@class_type", SqlDbType.VarChar, 6);
            SqlParameter param4 = new SqlParameter("@biz_area_cd", SqlDbType.VarChar, 10);
            SqlParameter param5 = new SqlParameter("@hq_brch_fg", SqlDbType.VarChar, 6);
            SqlParameter param6 = new SqlParameter("@zero_fg", SqlDbType.VarChar, 6);
            SqlParameter param7 = new SqlParameter("@ctrl_cd", SqlDbType.VarChar, 6);
            SqlParameter param8 = new SqlParameter("@usr_id", SqlDbType.VarChar, 20);
            SqlParameter param9 = new SqlParameter("@msg_cd", SqlDbType.VarChar, 6);
            SqlParameter param10 = new SqlParameter("@sp_id", SqlDbType.VarChar, 13);
            param10.Direction = ParameterDirection.Output;
            param1.Value = ls_fr_dt;
            param2.Value = ls_to_dt;
            param3.Value = "PL";
            param4.Value = ls_biz_area_cd;
            param5.Value = "N";
            param6.Value = "N";
            param7.Value = ls_ctrl_cd;
            param8.Value = "unierp2";
            param9.Value = ls_msg_cd;
            param10.Value = ls_sp_id;
            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
            cmd.Parameters.Add(param4);
            cmd.Parameters.Add(param5);
            cmd.Parameters.Add(param6);
            cmd.Parameters.Add(param7);
            cmd.Parameters.Add(param8);
            cmd.Parameters.Add(param9);
            cmd.Parameters.Add(param10);
            
            cmd.ExecuteNonQuery();
            ls_sp_id = cmd.Parameters["@sp_id"].Value.ToString();
            //dr = cmd.ExecuteReader();

            ////sp_id값 리턴받기
            //if (dr.Read())
            //{
            //    ls_sp_id = dr["@sp_id"].ToString();
            //}
            //dr.Close();
            conn.Close();
            // sp_id로 입력된 관리항목별 손익계산서 값을 가져오기

            ls_sql = "select class_nm , this_lamt, this_ramt, CASE WHEN A.CTRL_VAL = '' THEN '집계금액' ELSE ISNULL(a.ctrl_val,'집계금액') END  ctrl_val  from a_pl_mcs_ctrl_val a where sp_id = '" + ls_sp_id + "' order by class_cd, ctrl_val";
            ds_am_a1001 dt1 = new ds_am_a1001();
            ReportViewer1.Reset();
            ReportCreator(dt1, ls_sql, ReportViewer1, "rv_am_a1001.rdlc", "DataSet1");

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
                // 사업장 드랍다운리스트 내용을 보여준다.
                SqlCommand cmd2 = new SqlCommand(ls_biz_area_cd_sql, conn);

                dr = cmd2.ExecuteReader();
                if (ddl_biz_area.Items.Count < 2)
                {
                    ddl_biz_area.DataSource = dr;
                    ddl_biz_area.DataValueField = "BIZ_AREA_CD";
                    ddl_biz_area.DataTextField = "BIZ_AREA_NM";
                    ddl_biz_area.DataBind();
                }
                dr.Close();
                SqlCommand cmd3 = new SqlCommand(ls_ctrl_cd_sql, conn);

                dr = cmd3.ExecuteReader();
                if (ddl_ctrl_value.Items.Count < 2)
                {
                    ddl_ctrl_value.DataSource = dr;
                    ddl_ctrl_value.DataValueField = "CTRL_CD";
                    ddl_ctrl_value.DataTextField = "CTRL_NM";
                    ddl_ctrl_value.DataBind();
                }
                dr.Close();
                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                _reportViewer.LocalReport.DisplayName = "REPORT_관리항목_" + ls_ctrl_cd + "_손익계산서_" + DateTime.Now.ToShortDateString();
                
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

        
    }
}