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

namespace ERPAppAddition.ERPAddition.CM.CM_C1001
{
    public partial class CM_C1001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

        //SqlCommand cmd = new SqlCommand();
        //SqlDataReader dr;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //FillCostCenterInGrid();
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

        DataSet GetData(String queryString)
        {

            // Retrieve the connection string stored in the Web.config file.

            String connectionString = ConfigurationManager.ConnectionStrings["nepes"].ConnectionString;      

            DataSet ds = new DataSet();

            try
            {
                // Connect to the database and run the query.
                SqlConnection connection = new SqlConnection(connectionString);    
                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);

                // Fill the DataSet.
                adapter.Fill(ds);

            }
            catch (Exception ex)
            {
               
                // The connection failed. Display an error message.
                //Message.Text = "Unable to connect to the database.";

            }

            return ds;

        }
        
        private void FillCostCenterInGrid()
        {
             String queryString =
                    " select A.COST_CD, A.COST_NM, B.cost_gp_cd,C.UD_MINOR_NM cost_gp_nm " +
                    "   from B_COST_CENTER A inner JOIN crkd B ON A.COST_CD = B.cost_cd AND B.yyyymm LIKE '"+tb_yyyymm.Text+"' " +
                    "                        LEFT OUTER JOIN B_USER_DEFINED_MINOR C ON B.cost_gp_cd = C.UD_MINOR_CD AND C.UD_MAJOR_CD = 'A0002'";
             
             //ds_cm_c1001 ds = new ds_cm_c1001();
             DataSet ds = GetData(queryString);
             if (ds.Tables.Count > 0)
             {
                 GridView1.DataSource = ds;
                 GridView1.DataBind();
             }
             else
             {
                 //Message.Text = "Unable to connect to the database.";
             }  
            
        }

        protected void btn_retrive_Click(object sender, EventArgs e)
        {
            FillCostCenterInGrid();
        }
        private void DrawGridView()
        {
            if (ViewState["gridview1"] != null)
            {
                int nNumber = (int)ViewState["gridview1"];
                GridView gv = new GridView();
                gv = GridView1;
                Panel1.Controls.Add(gv);
            }
        }

        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
           // GridView1.EditIndex = -1;

           // FillCostCenterInGrid();
        }

        protected void GridView1_DataBound(object sender, EventArgs e)
        {

        }
    }
}