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
using ERPAppAddition.ERPAddition.AM.AM_A9002;

namespace ERPAppAddition.ERPAddition.AM.AM_A9002
{
    public partial class AM_A9002 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_amc"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dReader_select;
        string id = "";


        protected void Page_Load(object sender, EventArgs e)
        {
                if (Request.QueryString["id"] != null)
                {
                    id = Request.QueryString["id"];
                }
                else
                    id = "";

                WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
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
                dReader_select = cmd.ExecuteReader();
                ds.Tables[0].Load(dReader_select);
                dReader_select.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                if (rbl_view_type.SelectedValue == "A")
                {
                    _reportViewer.LocalReport.DisplayName = "REPORT_AMC일일실적조회_" + DateTime.Now.ToShortDateString();
                }
                else
                {
                    _reportViewer.LocalReport.DisplayName = "REPORT_AMC월별실적조회_" + DateTime.Now.ToShortDateString();
                }

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


        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(rbl_view_type.SelectedValue)
            {
                case "A" :
                Panel_bas_info.Visible = true;
                lb_yyyy.Visible = true;
                txt_yyyy.Visible = true;
                lb_mm.Visible = true;
                txt_mm.Visible = true;
                Select_Button.Visible = true;
                unit.Visible = true;
                break;

                case "B" :
                Panel_bas_info.Visible = true;
                lb_yyyy.Visible = true;
                txt_yyyy.Visible = true;
                lb_mm.Visible = false;
                txt_mm.Visible = false;
                Select_Button.Visible = true;
                unit.Visible = true;
                break;

                default :
                Panel_bas_info.Visible = false;
                lb_yyyy.Visible = false;
                txt_yyyy.Visible = false;
                lb_mm.Visible = false;
                txt_mm.Visible = false;
                Select_Button.Visible = false;
                unit.Visible = false;
                break;
            }
        }


        protected void Load_btn_Click(object sender, EventArgs e)
        {
            conn.Open();

            ReportViewer1.Reset();
            string Select_Qurey = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)
            string Report = string.Empty;
            

            if (rbl_view_type.SelectedValue == "A")
           {
               if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
                {
                   ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('[년]을 입력하세요.')", true);
                   return;
                }

               if (txt_mm == null || txt_mm.Text.Equals(""))
               {
                   ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('[월]을 입력하세요.')", true);
                   return;
               }

                   if (unit.SelectedValue == "WON")
                   {
                       Select_Qurey = "SELECT * FROM A_DAILY_AMT WHERE YYYY ='" + txt_yyyy.Text + "' AND MM ='" + txt_mm.Text + "'  ";
                       Report = "RP_AM_A9002_DAY.rdlc";
                   }
                   else
                   {
                       Select_Qurey = "SELECT * FROM A_DAILY_AMT WHERE YYYY ='" + txt_yyyy.Text + "' AND MM ='" + txt_mm.Text + "'  ";
                       Report = "RP_AM_A9002_DAY_x0.0001.rdlc";
                   }


           }

            if (rbl_view_type.SelectedValue == "B")
           {
               if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
               {
                   ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('[년]을 입력하세요.')", true);
                   return;
               }

                       if (unit.SelectedValue == "WON")
                       {
                           Select_Qurey = "USP_A_DAILY_AMT_MONSUM'" + txt_yyyy.Text + "' ";
                           Report = "RP_AM_A9002_MON.rdlc";
                       }
                       else
                       {
                           Select_Qurey = "USP_A_DAILY_AMT_MONSUM'" + txt_yyyy.Text + "' ";
                           Report = "RP_AM_A9002_MON_x0.0001.rdlc";
                       }
           }

            cmd.Connection = conn;
            cmd.CommandText = Select_Qurey;
            //cmd.CommandTimeout = 0;

            dReader_select = cmd.ExecuteReader();


            if (dReader_select.Read())
            {
                DataSet_AM_A9002 dt1 = new DataSet_AM_A9002();
                ReportViewer1.Reset();
                conn.Close();

                ReportCreator(dt1, Select_Qurey, ReportViewer1, Report, "DataSet1");
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 날짜의 데이터가 없습니다.')", true);
            }
            conn.Close();
        }
    }
}