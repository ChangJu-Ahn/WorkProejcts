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

namespace ERPAppAddition.ERPAddition.MM.MM_MRP
{
    public partial class MM_MRP_LEADTIME : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        string userid, db_name;      
       
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["db"] != null && Request.QueryString["db"].ToString() != "")
                {
                    db_name = Request.QueryString["db"].ToString();
                    if (db_name.Length > 0)
                    {
                        conn = new SqlConnection(ConfigurationManager.ConnectionStrings[db_name].ConnectionString);
                        // conn = new SqlConnection(ConfigurationManager.ConnectionStrings[userid].ConnectionString);

                    }
                    userid = Request.QueryString["userid"];

                    Session["DBNM"] = Request.QueryString["db"].ToString();
                    Session["User"] = Request.QueryString["userid"];
                }

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

        protected void bt_item_cd_Click(object sender, EventArgs e) //품목 조회 버튼 클릭
        {
            pop_gridview1.DataSourceID = "";
            pop_gridview1.DataSource = null;
            pop_gridview1.DataBind();
        }

        //팝업창에서 조회버튼 
        protected void bt_retrive_Click(object sender, EventArgs e)
        {
            pop_gridview1.DataSource = "";
            pop_gridview1.DataSource = SqlDataSource2;
            pop_gridview1.DataBind();
            pop_gridview1.Visible = true;
            pop_gridview1.SelectedIndex = -1;
            ModalPopupExtender1.Show();
        }
        // 팝업창-취소버튼클릭시 
        protected void bt_cancel_Click(object sender, EventArgs e)
        {
            //기존 보여졌던 데이타들을 안보이게 초기화
            pop_gridview1.DataSource = dr;
            pop_gridview1.DataBind();
            pop_gridview1.SelectedIndex = -1;
            ModalPopupExtender1.Hide();
        }
        protected void pop_gridview1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            //pageallow를 계속적으로 수행하기 위해 아래 코드가 필요
            pop_gridview1.PageIndex = e.NewPageIndex;
            pop_gridview1.DataBind();
            //새페이지를 눌렀을경우 gridview가 사라지기에 다시 조회하도록 조회버튼 호출
            bt_retrive_Click(this, e);
        }

        protected void pop_gridview1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //GridViewRow row = pop_gridview1.SelectedRow;
            //ls_item_cd = row.Cells[1].Text;
            //ls_item_nm = row.Cells[2].Text;


            //btn_pop_ok.Enabled = true;
        }
        // ok 버튼을 클릭하면 부모창에 값을 전달한다.
        protected void btn_pop_ok_Click(object sender, EventArgs e)
        {

            int i_chk_rowcnt = pop_gridview1.Rows.Count;
            string ls_chk_selectrowindex = pop_gridview1.SelectedIndex.ToString();

            if (ls_chk_selectrowindex != "-1")
            {
                GridViewRow row = pop_gridview1.SelectedRow;
                tb_item_cd.Text = row.Cells[1].Text;
                tb_item_nm.Text = row.Cells[2].Text;
            }
            pop_gridview1.DataSource = dr;
            pop_gridview1.DataBind();

        }
        
      

        protected void bt_save_Click(object sender, System.EventArgs e)//저장버튼
        {

            if (tb_item_cd .Text== null || tb_item_cd.Text.Equals(""))
            {
                MessageBox.ShowMessage("품목코드를 입력하세요.", this.Page);

                return;
            }


            if (tb_bizpartner.Text == null || tb_bizpartner.Text.Equals(""))
            {
                MessageBox.ShowMessage("제조사를 입력하세요.", this.Page);

                return;
            }
              conn.Open();

              string queryStr = "insert into m_mrp_leadtime(ITEM_CD,LEADTIME,MOQ,BIZPARTNER,INSRT_DT,INSRT_USR_ID,UPDT_DT,UPDT_USR_ID)";
                       queryStr +=" values('" + tb_item_cd.Text + "', '" + tb_leadtime.Text + "','" + tb_moq.Text + "','" + tb_bizpartner.Text + "',";
                       queryStr += "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "')";
            
            SqlCommand sComm = new SqlCommand(queryStr, conn);
            MessageBox.ShowMessage("저장되었습니다.", this.Page);

            sComm.ExecuteNonQuery();
            conn.Close();

        }

        protected void bt_del_Click(object sender, System.EventArgs e) //삭제 버튼
        {
            conn.Open();

            string queryStr = "Delete from m_mrp_leadtime where ITEM_CD='" + tb_item_cd.Text + "' and BIZPARTNER = '" + tb_bizpartner.Text + "'";

            SqlCommand sComm = new SqlCommand(queryStr, conn);
            MessageBox.ShowMessage("삭제되었습니다.", this.Page);
            sComm.ExecuteNonQuery();
            conn.Close();
        }

        protected void bt_change_Click(object sender, System.EventArgs e) //수정 버튼
        {
            conn.Open();
          
            string queryStr = "update m_mrp_leadtime set ITEM_CD = '" + tb_item_cd.Text + "',LEADTIME='" + tb_leadtime.Text + "'";
            queryStr += "MOQ = '" + tb_moq.Text + "',BIZPARTNER = '" + tb_bizpartner.Text + "'";
            queryStr += " ,updt_dt='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',updt_user_id ='" + Session["User"].ToString() + "' where ITEM_CD='" + tb_item_cd.Text + "' and BIZPARTNER = '" + tb_bizpartner.Text + "'";

            SqlCommand sComm = new SqlCommand(queryStr, conn);

            MessageBox.ShowMessage("수정되었습니다.", this.Page);
            sComm.ExecuteNonQuery();
            conn.Close();

        }

        protected void bt_refresh_Click(object sender, System.EventArgs e) //재작성 버튼
        {
            tb_item_cd.Text = string.Empty;
            bt_item_cd.Text = string.Empty;
            tb_leadtime.Text = string.Empty;
            tb_moq.Text = string.Empty;
            tb_bizpartner.Text = string.Empty;


            bt_save.Enabled = true;
        }

        protected void bt_retrieve_Click(object sender, System.EventArgs e) //조회 버튼
        {
            string queryStr = "SELECT ITEM_CD,LEADTIME,MOQ,BIZPARTNER FROM m_mrp_leadtime  where ITEM_CD='" + tb_item_cd_0.Text + "' and BIZPARTNER = '" + tb_bizpartner_0.Text + "'";

            conn.Open();
            SqlCommand sComm = new SqlCommand(queryStr, conn);

            SqlDataReader dReader = sComm.ExecuteReader();
            string temp1, temp2, temp3, temp4;


            if (dReader.Read())
            {
                temp1 = "" + dReader[0].ToString();
                temp2 = "" + dReader[1].ToString();
                temp3 = "" + dReader[2].ToString();
                temp4 = "" + dReader[3].ToString();

                tb_item_cd.Text = temp1;
                tb_leadtime.Text = temp2;
                tb_moq.Text = temp3;
                tb_bizpartner.Text = temp4;             


                bt_save.Enabled = false;
            }
            conn.Close();
        }




    }
}