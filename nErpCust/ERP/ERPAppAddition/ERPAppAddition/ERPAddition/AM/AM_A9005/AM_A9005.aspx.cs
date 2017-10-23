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
namespace ERPAppAddition.ERPAddition.AM.AM_A9005
{
    public partial class AM_A9005 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlCommand cmd1 = new SqlCommand();
        SqlDataReader dReader_select;
        string id = "";
        string ls_biz_area_cd_sql;
        string ls_month_sql;
        string ls_date;
        SqlDataReader dr;
        SqlDataReader dr1;
        SqlDataReader dr5;


        //private void open_serch()
        //{
        //    string Select_Qurey5 = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)
        //    string Select_Qurey6 = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)

        //    Select_Qurey5 = "USP_COST_REPORT_1'" + txt_yyyy.Text + "' ";
        //    Select_Qurey6 = "USP_COST_REPORT_TOTAL'" + txt_yyyy.Text + "' ";

        //    DataSet_AM_A9005 dt1 = new DataSet_AM_A9005();
        //    DataSet_AM_A9005_1 dt2 = new DataSet_AM_A9005_1();
        //    //DataSet_AM_A9005_1 dt2 = new DataSet_AM_A9005_1();


        //    ReportViewer1.Reset();
        //    //conn.Close();

        //    ReportCreator(dt1, Select_Qurey5, dt2, Select_Qurey6, ReportViewer1, "Report_AM_9005.rdlc", "DataSet1", "DataSet2");
        //}


        protected void Page_Load(object sender, EventArgs e)
        {
            Panel_bas_info.Visible = true;
            lb_yyyy.Visible = true;
            //txt_yyyy.Visible = true;
            //lb_mm.Visible = true;
            //txt_mm.Visible = true;
            Select_Button.Visible = true;

            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandTimeout = 500;
            cmd.CommandType = CommandType.Text;

            cmd1 = conn.CreateCommand();

            cmd1.CommandTimeout = 500;
            cmd1.CommandType = CommandType.Text;


            ls_biz_area_cd_sql = "SELECT '%'BIZ_AREA_CD, '전체' BIZ_AREA_NM union all SELECT 'P01' BIZ_AREA_CD , '1공장' BIZ_AREA_NM  union all SELECT 'P02' BIZ_AREA_CD, '2공장' BIZ_AREA_NM union all SELECT 'P09' BIZ_AREA_CD, 'PKG' BIZ_AREA_NM ";

            ls_date = "select  CONVERT(CHAR(6),GETDATE(),112)";


            ls_month_sql = "select distinct(yyyymm) yyyymm from day_cost_report_01 order by yyyymm desc";


            //txt_yyyy.Text = DateTime.Now.ToString("yyyy") + DateTime.Now.ToString("MM");

            // 사업장 드랍다운리스트 내용을 보여준다.
            SqlCommand cmd2 = new SqlCommand(ls_biz_area_cd_sql, conn);

            SqlCommand cmd5 = new SqlCommand(ls_month_sql, conn);

            //SqlCommand cmd3 = new SqlCommand(ls_date, conn);


            //dr = cmd2.ExecuteReader();



            dr = cmd5.ExecuteReader();

            if (ddl_month.Items.Count < 2)
            {
                ddl_month.DataSource = dr;
                ddl_month.DataValueField = "yyyymm";
                ddl_month.DataTextField = "yyyymm";
                ddl_month.DataBind();


            }

      


            //dr = cmd5.ExecuteReader();


            //dr = cmd3.ExecuteReader();

            //if (ddl_biz_area.Items.Count < 2)
            //{
            //    ddl_biz_area.DataSource = dr;
            //    ddl_biz_area.DataValueField = "BIZ_AREA_CD";
            //    ddl_biz_area.DataTextField = "BIZ_AREA_NM";
            //    ddl_biz_area.DataBind();
            //}
            
            dr.Close();

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


            // txt_yyyy.Text = ls_date;



            if (Request.QueryString["id"] != null)
            {
                id = Request.QueryString["id"];
            }
            else
                id = "";



            //open_serch();
            WebSiteCount();

        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }      

        private void ReportCreator(DataSet _dataSet, string _Query, DataSet _dataSet1, string _Query1, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName, string _ReportDataSourceName1)
        {

            //conn.Open();
            //cmd = conn.CreateCommand();
            //cmd.CommandTimeout = 500;
            //cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;

            DataSet ds1 = _dataSet1;

            try
            {
                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();

                cmd1.CommandText = _Query1;
                dr = cmd1.ExecuteReader();
                ds1.Tables[1].Load(dr);
                dr.Close();


                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                _reportViewer.LocalReport.EnableHyperlinks = true;

                _reportViewer.LocalReport.DisplayName = "REPORT_" + DateTime.Now.ToShortDateString();

                ReportDataSource rds = new ReportDataSource();
                ReportDataSource rds1 = new ReportDataSource();

                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                _reportViewer.LocalReport.DataSources.Add(rds);

                rds1.Name = _ReportDataSourceName1;
                rds1.Value = ds1.Tables[1];
                _reportViewer.LocalReport.DataSources.Add(rds1);

                _reportViewer.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }




        //    conn.Open();
        //    cmd = conn.CreateCommand();
        //    cmd.CommandTimeout = 500;
        //    cmd.CommandType = CommandType.Text;

        //    DataSet ds = _dataSet;
        //    DataSet ds1 = _dataSet;
        //    try

        //    {


        //        cmd.CommandText = _Query;
        //        dReader_select = cmd.ExecuteReader();
        //        ds.Tables[0].Load(dReader_select);
        //        dReader_select.Close();
        //        _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);


        //        cmd.CommandText = _Query1;
        //        dReader_select = cmd.ExecuteReader();
        //        ds1.Tables[0].Load(dReader_select);
        //        dReader_select.Close();

        //        _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);



        //        if (rbl_view_type.SelectedValue == "A")
        //        {
        //            _reportViewer.LocalReport.DisplayName = "전체_" + DateTime.Now.ToShortDateString();
        //        }
        //        else
        //        {
        //            _reportViewer.LocalReport.DisplayName = "세부_" + DateTime.Now.ToShortDateString();
        //        }

        //        ReportDataSource rds = new ReportDataSource();
        //        rds.Name = _ReportDataSourceName;
        //        rds.Value = ds.Tables[0];
        //        _reportViewer.LocalReport.DataSources.Add(rds);

        //        //_reportViewer.LocalReport.Refresh();


        //        ReportDataSource rds1 = new ReportDataSource();
        //        rds1.Name = _ReportDataSourceName1;
        //        rds1.Value = ds1.Tables[0];
        //        _reportViewer.LocalReport.DataSources.Add(rds1);


        //        _reportViewer.LocalReport.Refresh();
        //    }
        //    catch { }
        //    finally
        //    {
        //        if (conn.State == ConnectionState.Open)
        //            conn.Close();
        //    }
        //}


        //protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)

        //{
        //    //switch (rbl_view_type.SelectedValue)
        //    //{
        //    ////    case "A":
        //    //Panel_bas_info.Visible = true;
        //    //lb_yyyy.Visible = true;
        //    //txt_yyyy.Visible = true;
        //    //lb_mm.Visible = true;
        //    //txt_mm.Visible = true;
        //    //Select_Button.Visible = true;
        //    //        break;

        //    //    case "B":
        //    //        Panel_bas_info.Visible = true;
        //    //        lb_yyyy.Visible = true;
        //    //        txt_yyyy.Visible = true;
        //    //        lb_mm.Visible = false;
        //    //        txt_mm.Visible = false;
        //    //        Select_Button.Visible = true;
        //    //        break;

        //    //    default:
        //    //        Panel_bas_info.Visible = false;
        //    //        lb_yyyy.Visible = false;
        //    //        txt_yyyy.Visible = false;
        //    //        lb_mm.Visible = false;
        //    //        txt_mm.Visible = false;
        //    //        Select_Button.Visible = false;
        //    //        break;
        //    //}
        //}


        protected void Load_btn_Click(object sender, EventArgs e)
        {


            ReportViewer1.Reset();
            string Select_Qurey = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)
            string Select_Qurey1 = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)
            string Select_Qurey3 = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)
            string Select_Qurey4 = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)
            string Report = string.Empty;




            //if (rbl_view_type.SelectedValue == "A")
            //{
            //    if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
            //    {
            //        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('[날짜]를 입력하세요.')", true);
            //        return;
            //    }

            //if (txt_mm == null || txt_mm.Text.Equals(""))
            //{
            //    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('[날짜]를 입력하세요.')", true);
            //    return;
            //}


            //DataSet_AM_A9003TableAdapters.USP_DIS_AR_VIEWTableAdapter adapter = new DataSet_AM_A9003TableAdapters.USP_DIS_AR_VIEWTableAdapter();
            //DataSet_AM_A9003.USP_DIS_AR_VIEWDataTable dt1 = adapter.GetData(Convert.ToDateTime(txt_yyyy.Text), Convert.ToDateTime(txt_mm.Text));

            if (rbl_view_type.SelectedValue == "A")
            {

                string ls_month;

                ls_month = ddl_month.Text;


                Select_Qurey = "USP_COST_REPORT_1'" + ls_month + "' ";
                Select_Qurey1 = "USP_COST_REPORT_TOTAL'" + ls_month + "' ";

                DataSet_AM_A9005 dt1 = new DataSet_AM_A9005();
                DataSet_AM_A9005_1 dt2 = new DataSet_AM_A9005_1();
                //DataSet_AM_A9005_1 dt2 = new DataSet_AM_A9005_1();


                ReportViewer1.Reset();
                //conn.Close();

                ReportCreator(dt1, Select_Qurey, dt2, Select_Qurey1, ReportViewer1, "Report_AM_9005.rdlc", "DataSet1", "DataSet2");
            }

            if (rbl_view_type.SelectedValue == "B")

            {
                string ls_month;

                ls_month = ddl_month.Text;

                Select_Qurey = "USP_COST_REPORT_1'" +  ls_month + "' ";
                Select_Qurey1 = "USP_COST_REPORT_3'" + ls_month + "' ";

                DataSet_AM_A9005 dt3 = new DataSet_AM_A9005();
                DataSet_AM_A9005_1 dt4 = new DataSet_AM_A9005_1();


                ReportViewer1.Reset();
                //conn.Close();

                ReportCreator(dt3, Select_Qurey, dt4, Select_Qurey1, ReportViewer1, "Report_AM_9005_1.rdlc", "DataSet1", "DataSet2");

            }


            //Report = "Report_AM_9004_1.rdlc";
            //Select_Qurey1 = "USP_M_REPORT_F_1'" + txt_yyyy.Text + "' ";
            //Report = "Report_AM_9004_1.rdlc";



            //        }

            //        if (rbl_view_type.SelectedValue == "B")
            //        {
            //            if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
            //            {
            //                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('[날짜]을 입력하세요.')", true);
            //                return;
            //            }

            //            Select_Qurey = "USP_M_REPORT_F'" + txt_yyyy.Text + "' ";
            //            Report = "Report_AM_9004_1.rdlc";
            //        }

            //        cmd.Connection = conn;
            //        cmd.CommandText = Select_Qurey;
            //        //cmd.CommandText = Select_Qurey1;

            //        //dReader_select = cmd.ExecuteReader();


            //        //cmd.Connection = conn;
            //        cmd.CommandText = Select_Qurey1;

            //        dReader_select = cmd.ExecuteReader();




            //        if (dReader_select.Read())
            //        {

            //            DataSet_AM_A9004 dt1 = new DataSet_AM_A9004();
            //            DataSet_AM_A9004_1 dt2 = new DataSet_AM_A9004_1();


            //            ReportViewer1.Reset();
            //            //ReportViewer2.Reset();

            //            conn.Close();

            //            //ReportCreator(dt1, Select_Qurey, ReportViewer1, Report, "DataSet1");

            //            //ReportCreator(dt1, Select_Qurey, dt2, Select_Qurey1, ReportViewer1, "Report_AM_9004_1.rdlc", "DataSet1", "DataSet2");

            //            //ReportCreator(dt2, Select_Qurey1, ReportViewer1, Report, "DataSet2");
            //        }
            //        else
            //        {
            //            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 날짜의 데이터가 없습니다.')", true);
            //        }
            //        conn.Close();
            //    }
            //}
        }
    }
}