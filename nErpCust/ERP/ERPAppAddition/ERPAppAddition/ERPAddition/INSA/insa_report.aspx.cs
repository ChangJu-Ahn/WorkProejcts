using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Services;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Reporting.WebForms;

namespace ERPAppAddition.ERPAddition.INSA
{
    public partial class insa_report : System.Web.UI.Page
    {

        string strPlant = "";
        string strGubun = "";
        string strOSFlag = "";
        string strTotal = "";
        public string strDay = "";

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_test1"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString["id"] != null)
            {
                Session["id"] = Request.QueryString["id"];
                Session.Timeout = 30;
            }
            Init_control();

            if (!IsPostBack)
            {              
                Init_Setting();
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

        private void Init_control()
        {
           
            if (rbl_view_type.SelectedValue == "A")
            {
                chk_standard.Visible = false;
                search_table.Visible = true;
                btn_hidden.Style["visibility"] = "hidden";
            }
            else if (rbl_view_type.SelectedValue == "B")
            {
                ReportViewer1.Reset();
                ReportViewer2.Reset();

                string id = "";

                if (Session["id"] != null)
                {
                     id = Session["id"].ToString();
                }

                if (id == "20130727" || id == "20090619")
                {
                    chk_standard.Visible = true;
                    search_table.Visible = false;
                }
                else
                {
                    search_table.Visible = false;
                    Response.Write("<script>alert('권한 오류')</script>");
                }

            }
        }

        private void Init_Setting()
        {
            
            if (Request.QueryString["plant"] != null)
                strPlant = Request.QueryString["plant"];

            if (Request.QueryString["gubun"] != null)
                strGubun = Request.QueryString["gubun"];

            if (Request.QueryString["os"] != null)
                strOSFlag = Request.QueryString["os"];

            if (Request.QueryString["total"] != null)
                strTotal = Request.QueryString["total"];

            if (Request.QueryString["day"] != null)
                strDay = Request.QueryString["day"];

            if (strOSFlag != "" && strTotal != "")
                Click_Result();

        }

        public void Click_Result()
        {
            string strId = "";

            if (Session["id"] == null)
            {
                MessageBox.ShowMessage("세션이 끊어졌습니다. 재접속 해주세요", this.Page);
            }
            else
            {
                strId = Session["id"].ToString();

                string sql = "EXEC USP_NEPES_INSA '" + strPlant + "','" + strGubun + "','" + strOSFlag + "','" + strTotal + "'";
                string sql1 = "EXEC USP_NEPES_INSA_TOTAL";

                DataSet_insa_detail dt2 = new DataSet_insa_detail();
                DataSet_insa1 dt1 = new DataSet_insa1();

                ReportViewer1.Reset();
                ReportViewer2.Reset();

                tb_yyyymm.Text = strDay;

                ReportCreator(dt1, sql1, ReportViewer1, "insa.rdlc", "DataSet1");


                if (strTotal == "Y")
                {
                    ReportCreator(dt2, sql, ReportViewer2, "insa_detail.rdlc", "DataSet2");
                }
                else
                {
                    ReportCreator(dt2, sql, ReportViewer2, "DataSet_insa_detail_add.rdlc", "DataSet2");
                }
                
                tb_yyyymm.Text = strDay;
            }
        }

        protected void bt_retrieve_Click(object sender, EventArgs e)
        {
            string sql = "EXEC USP_NEPES_INSA_TOTAL";

            DataSet_insa1 dt1 = new DataSet_insa1();

            ReportViewer1.Reset();
            ReportViewer2.Reset();

            ReportCreator(dt1, sql, ReportViewer1, "insa.rdlc", "DataSet1");

        }

        protected void bt_insert_click(object sender, EventArgs e)
        {

            conn.Open();
            cmd = conn.CreateCommand();

            try
            {
                string selecedItems = Request.Form[list_add.UniqueID];
                list_add.Items.Clear();

                if (selecedItems != null)
                {
                    string delSql = "DELETE FROM INSA_AUTHORITY WHERE EMP_NO = '" + combobox.Value.ToString() + "' ";

                    cmd.CommandText = delSql;
                    cmd.ExecuteNonQuery();

                    foreach (string item in selecedItems.Split(','))
                    {
                        string sql = "insert into INSA_AUTHORITY values('" + combobox.Value.ToString() + "','" + item + "')";

                        cmd.CommandText = sql;
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string delSql = "DELETE FROM INSA_AUTHORITY WHERE EMP_NO = '" + combobox.Value.ToString() + "' ";

                    cmd.CommandText = delSql;
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
                MessageBox.ShowMessage(message.Replace("'", "").Replace("\r", "").Replace("\n", ""), this.Page);
            }
            finally
            {
                MessageBox.ShowMessage("완료 되었습니다.", this.Page);
                if (conn.State == ConnectionState.Open)
                    conn.Close();
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

                _reportViewer.LocalReport.EnableHyperlinks = true;

                _reportViewer.LocalReport.DisplayName = "REPORT_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];

                ReportParameter day = new ReportParameter("strYYY", "20160314");
                ReportParameter id = new ReportParameter("strId", "20130727");
                _reportViewer.LocalReport.SetParameters(new ReportParameter[] { day,id });

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

        protected void btn_hidden_Click(object sender, EventArgs e)
        {
            conn.Open();
            try
            {

                string sql = "select bu_name,bu_code from INSADB.INBUS.DBO.c_busor WHERE bu_edate = '' AND BU_CODE IN(SELECT DEPT_CD FROM INSA_AUTHORITY WHERE EMP_NO = '" + combobox.Value.ToString() + "')";

                SqlCommand mySqlCommand = new SqlCommand(sql, conn);
                SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(mySqlCommand);

                DataSet myDataSet = new DataSet();
                mySqlDataAdapter.Fill(myDataSet);

                list_add.DataSource = myDataSet;
                list_add.DataTextField = "bu_name";
                list_add.DataValueField = "bu_code";

                list_add.DataBind();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
                MessageBox.ShowMessage(message.Replace("'", "").Replace("\r", "").Replace("\n", ""), this.Page);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
    }
}