using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;

namespace ERPAppAddition.ERPAddition.OM.OM_O1001
{
    public partial class pop_om_o2001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        protected void Page_Load(object sender, EventArgs e)
        {
            GetPatList("", "");
        }

        private void GetPatList(string fromDate, string toDate)
        {


            string strSQL = "SELECT apply_no,invent_kr_nm,apply_dt FROM O_PATENT";

            strSQL += " where 1 = 1";

            if (!fromDate.Equals(""))
            {
                strSQL += " and CONVERT(CHAR(10), apply_dt, 23) >= '" + fromDate + "'";
            }

            if (!toDate.Equals(""))
            {
                strSQL += " and CONVERT(CHAR(10), apply_dt, 23) <= '" + toDate + "'";
            }

            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSQL;
            //SqlDataAdapter objAdapter = new SqlDataAdapter(strSQL, conn);
            ds_om_o1001 ds = new ds_om_o1001();
            dr = cmd.ExecuteReader();

            
            //DataSet objDataSet = new DataSet("O_PATENT");
            //objAdapter.Fill(objDataSet, "O_PATENT");

            ds.Tables[0].Load(dr);

           SerGridView.DataSource = ds.Tables[0]; //objDataSet.Tables["O_PATENT"];
            SerGridView.DataBind();
            conn.Close();

        }

           protected void btnSearch2_Click(object sender, EventArgs e)//조회화면에서 날짜 조정 후 확인 버튼 클릭
        {
            GetPatList(txtFromDate.Text, txtToDate.Text);
        }

          

        protected void SerGridView_SelectedIndexChanged1(object sender, EventArgs e)//선택된 값이 출원번호 텍스트박스로 이동
        {
            if (Page.PreviousPage != null)
            {
                TextBox SourceTextBox = (TextBox)Page.PreviousPage.FindControl("TxtApply_no");

                if (SourceTextBox != null)
                {
                    SourceTextBox.Text = SerGridView.Rows[SerGridView.SelectedIndex].Cells[0].ToString();

                }


            }
        }

      


    }
}