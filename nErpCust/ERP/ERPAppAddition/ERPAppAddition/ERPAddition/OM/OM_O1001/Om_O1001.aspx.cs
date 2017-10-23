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
using ERPAppAddition.ERPAddition.OM.OM_O1001;



namespace ERPAppAddition.ERPAddition.OM.OM_O1001
{
    
    public partial class Om_O1001 : System.Web.UI.Page    
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
                //else
                //{
                //  string script = "alert(\"프로그램 호출이 잘못되었습니다. 관리자에게 연락해주세요.\");";
                //ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                //}

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
        protected void BtnSearch_Click(object sender, EventArgs e)//출원번호 찾기 클릭
        {
            Response.Write("<script>window.open('pop_om_o2001.aspx','','top=100,left=100,width=800,height=600')</script>");

        }

        protected void BtnLoad_Click(object sender, EventArgs e)//조회 
        {
            if (TxtApplyNo == null || TxtApplyNo.Text.Equals(""))
            {
                MessageBox.ShowMessage("출원번호를 입력하세요.", this.Page);

                return;
            }

            string queryStr = "SELECT * FROM O_PATENT where apply_no='" + TxtApplyNo.Text + "'";

            conn.Open();
            SqlCommand sComm = new SqlCommand(queryStr, conn);

            SqlDataReader dReader = sComm.ExecuteReader();
            string temp1, temp2, temp3, temp4, temp5, temp6, temp7, temp8, temp9, temp10, temp11, temp12, temp13, temp14, temp15, temp16, temp17, temp18, temp19, temp20, temp21, temp22, temp23, temp24;


            if (dReader.Read())
            {
                temp1 = "" + dReader[0].ToString();
                temp2 = "" + dReader[1].ToString();
                temp3 = "" + dReader[2].ToString();
                temp4 = "" + dReader[3].ToString();
                temp5 = "" + dReader[4].ToString();
                temp6 = "" + dReader[5].ToString();
                temp7 = "" + dReader[6].ToString();
                temp8 = "" + dReader[7].ToString();
                temp9 = "" + dReader[8].ToString();
                temp10 = "" + dReader[9].ToString();
                temp11 = "" + dReader[10].ToString();
                temp12 = "" + dReader[11].ToString();
                temp13 = "" + dReader[12].ToString();
                temp14 = "" + dReader[13].ToString();
                temp15 = "" + dReader[14].ToString();
                temp16 = "" + dReader[15].ToString();
                temp17 = "" + dReader[16].ToString();
                temp18 = "" + dReader[17].ToString();
                temp19 = "" + dReader[18].ToString();
                temp20 = "" + dReader[19].ToString();
                temp21 = "" + dReader[20].ToString();
                temp22 = "" + dReader[21].ToString();
                temp23 = "" + dReader[22].ToString();
                temp24 = "" + dReader[23].ToString();

                TxtApplyNo.Text = temp1;
                DropDownDept_Cd.Text = temp2;
                DropDownCountry.Text = temp3;
                DropDownRegion.Text = temp4;
                DropDownStatus.Text = temp5;
                DropDownType_Cd1.Text = temp6;
                DropDownType_Cd2.Text = temp7;
                TxtPriorityDT.Text = temp8;
                TxtInvent_Kr_Nm.Text = temp9;
                TxtInvent_En_Nm.Text = temp10;
                TxtApply_DT.Text = temp11;
                TxtOpen_No.Text = temp12;
                TxtOpen_DT.Text = temp13;
                TxtRegistNo.Text = temp14;
                TxtRegist_DT.Text = temp15;
                TxtApply_Comp.Text = temp16;
                TxtInvent_Nm.Text = temp17;
                TxtE_Contry_Nm.Text = temp18;
                TxtPriorityCd.Text = temp19;
                TxtPct_Apply_No.Text = temp20;
                TxtSubstitute_Nm.Text = temp21;
                TxtAsset_DT.Text = temp22;
                TxtExp_DT.Text = temp23;
                TxtRemark.Text = temp24;

                BtnSave.Enabled = false;
            }
            conn.Close();


        }

        protected void BtnSave_Click(object sender, EventArgs e)//저장
        {
          
            
                if (TxtApplyNo == null || TxtApplyNo.Text.Equals(""))
                {
                    MessageBox.ShowMessage("출원번호를 입력하세요.", this.Page);

                    return;
                }


                conn.Open();

                string queryStr = "insert into O_PATENT(apply_no,dept_cd,country_cd,region_nm,status_cd,type_cd1,type_cd2,f_yyyy,invent_kr_nm,invent_en_nm,apply_dt,open_no,open_dt,regist_no,regist_dt,apply_comp,invent_nm,e_contry_nm,priority_no,pct_apply_no,substitute_nm,asset_dt,exp_dt,remark,isrt_dt,isrt_user_id,updt_dt,updt_user_id)  values('" + TxtApplyNo.Text + "', '" + DropDownDept_Cd.Text + "','" + DropDownCountry.Text + "','" + DropDownRegion.Text + "',";
                queryStr += "'" + DropDownStatus.Text + "','" + DropDownType_Cd1.Text + "','" + DropDownType_Cd2.Text + "','" + TxtPriorityDT.Text + "','" + TxtInvent_Kr_Nm.Text + "',";
                queryStr += "'" + TxtInvent_En_Nm.Text + "','" + TxtApply_DT.Text + "','" + TxtOpen_No.Text + "','" + TxtOpen_DT.Text + "','" + TxtRegistNo.Text + "','" + TxtRegist_DT.Text + "',";
                queryStr += "'" + TxtApply_Comp.Text + "','" + TxtInvent_Nm.Text + "','" + TxtE_Contry_Nm.Text + "','" + TxtPriorityCd.Text + "','" + TxtPct_Apply_No.Text + "','" + TxtSubstitute_Nm.Text + "',";
                queryStr += "'" + TxtAsset_DT.Text + "','" + TxtExp_DT.Text + "','" + TxtRemark.Text + "',  '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "')";



                SqlCommand sComm = new SqlCommand(queryStr, conn);
                MessageBox.ShowMessage("저장되었습니다.", this.Page);

                sComm.ExecuteNonQuery();
                conn.Close();
            }
       

        protected void BtnDel_Click1(object sender, EventArgs e)//삭제
        {

            conn.Open();

            string queryStr = "Delete from O_PATENT where apply_no='" + TxtApplyNo.Text + "'";
            
            SqlCommand sComm = new SqlCommand(queryStr, conn);
            MessageBox.ShowMessage("삭제되었습니다.", this.Page);
            sComm.ExecuteNonQuery();
            conn.Close();
        }

        protected void BtnChange_Click(object sender, EventArgs e) //수정
        {

            conn.Open();
            string a = TxtOpen_DT.Text;
            string queryStr = "update O_PATENT set dept_cd = '" + DropDownDept_Cd.Text + "',country_cd='" + DropDownCountry.Text + "',";
            queryStr += "region_nm = '" + DropDownRegion.Text + "',status_cd = '" + DropDownStatus.Text + "',type_cd1 = '" + DropDownType_Cd1.Text + "',";
            queryStr += "type_cd2 ='" + DropDownType_Cd2.Text + "',f_yyyy = '" + TxtPriorityDT.Text + "',invent_kr_nm = '" + TxtInvent_Kr_Nm.Text + "',";
            queryStr += "invent_en_nm = '" + TxtInvent_En_Nm.Text + "',apply_dt = '" + TxtApply_DT.Text + "',open_no =  '" + TxtOpen_No.Text + "',";
            queryStr += "open_dt = '" + TxtOpen_DT.Text + "',regist_no = '" + TxtRegistNo.Text + "',regist_dt = '" + TxtRegist_DT.Text + "',";
            queryStr += "apply_comp = '" + TxtApply_Comp.Text + "',invent_nm = '" + TxtInvent_Nm.Text + "',e_contry_nm = '" + TxtE_Contry_Nm.Text + "',";
            queryStr += "priority_no = '" + TxtPriorityCd.Text + "',pct_apply_no = '" + TxtPct_Apply_No.Text + "',substitute_nm = '" + TxtSubstitute_Nm.Text + "',";
            queryStr += "asset_dt = '" + TxtAsset_DT.Text + "',exp_dt = '" + TxtExp_DT.Text + "',remark = '" + TxtRemark.Text + "',updt_dt='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+"',updt_user_id ='"+ Session["User"].ToString() +"' where apply_no='" + TxtApplyNo.Text + "'";
            
            SqlCommand sComm = new SqlCommand(queryStr, conn);

            MessageBox.ShowMessage("수정되었습니다.", this.Page);
            sComm.ExecuteNonQuery();
            conn.Close();

          
        }

        protected void BtnClear_Click(object sender, EventArgs e) //재작성
        {
            TxtApplyNo.Text = string.Empty;
            TxtPriorityDT.Text = string.Empty;
            TxtInvent_Kr_Nm.Text = string.Empty;
            TxtInvent_En_Nm.Text = string.Empty;
            TxtApply_DT.Text = string.Empty;
            TxtOpen_No.Text = string.Empty;
            TxtOpen_DT.Text = string.Empty;
            TxtRegistNo.Text = string.Empty;
            TxtRegist_DT.Text = string.Empty;
            TxtApply_Comp.Text = string.Empty;
            TxtInvent_Nm.Text = string.Empty;
            TxtE_Contry_Nm.Text = string.Empty;
            TxtPriorityCd.Text = string.Empty;
            TxtPct_Apply_No.Text = string.Empty;
            TxtSubstitute_Nm.Text = string.Empty;
            TxtAsset_DT.Text = string.Empty;
            TxtExp_DT.Text = string.Empty;
            TxtRemark.Text = string.Empty;
            
            BtnSave.Enabled = true;
        }


        public string sql { get; set; }
    }
}