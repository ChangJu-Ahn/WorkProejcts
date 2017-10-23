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
using System.Drawing;
using System.Web.Helpers;


namespace ERPAppAddition.ERPAddition.SM.sm_sb001
{
    public partial class sm_sb001_A02 : System.Web.UI.Page
    {
        string userid;
        string process_instance_oid;
        int value;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["NEPES_MAIL"].ConnectionString);
        SqlConnection sql_conn_NEPES = new SqlConnection(ConfigurationManager.ConnectionStrings["NEPES"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        SqlCommand sql_cmd2 = new SqlCommand();
        SqlCommand sql_cmd_chk = new SqlCommand();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "pop";
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;

                if (Request.QueryString["process_instance_oid"] == null || Request.QueryString["process_instance_oid"] == "")
                    process_instance_oid = "397833"; // "358751";  // 미대상 359174  운영 397833   개발용  2738
                else
                    process_instance_oid = Request.QueryString["process_instance_oid"];
                Session["process_instance_oid"] = process_instance_oid;

                WebSiteCount();
            }

            search();            
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn_NEPES.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void search()
        {
            DataSet ds = new DataSet();
            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = getSQL();

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds, "DataSet1");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
                sql_conn.Close();

                if (ds.Tables["DataSet1"].Rows.Count <= 0)
                {
                    DataRow dr = ds.Tables["DataSet1"].NewRow();
                    ds.Tables["DataSet1"].Rows.InsertAt(dr, 0);


                    //dt.Rows.Add(new object[] { "" });
                    //GridView1.Columns[6].Visible = false;

                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }
                else
                {
                    DOC_NO.Text = ds.Tables["DataSet1"].Rows[0]["DOC_NO"].ToString();
                    SUBJ.Text = ds.Tables["DataSet1"].Rows[0]["SUBJ"].ToString();
                    OUT_DT_ITEM.Text = ds.Tables["DataSet1"].Rows[0]["EXPORT_DATE"].ToString();
                    STS.Text = ds.Tables["DataSet1"].Rows[0]["STS"].ToString();
                    FROM_FD_YN.Text = ds.Tables["DataSet1"].Rows[0]["FROM_FD_YN"].ToString();
                    USR_NM.Text = ds.Tables["DataSet1"].Rows[0]["USR_NM"].ToString();
                    USR_DUTY.Text = ds.Tables["DataSet1"].Rows[0]["USR_DUTY"].ToString();
                    APPR_DEPT.Text = ds.Tables["DataSet1"].Rows[0]["DEPT_NM"].ToString();
                    OUT_DT.Text = ds.Tables["DataSet1"].Rows[0]["OUT_DT"].ToString();
                    OUT_COF_USER.Text = ds.Tables["DataSet1"].Rows[0]["OUT_COF_USER"].ToString();


                    // 개발 ip : http://192.168.10.113   운영은 :  http://mail.nepes.co.kr
                    if (ds.Tables["DataSet1"].Rows[0]["IMAGE1_PATH"].ToString() != "")
                    {
                        ImageButton1.ImageUrl = "http://mail.nepes.co.kr" + ds.Tables["DataSet1"].Rows[0]["IMAGE1_PATH"].ToString();
                        ImageButton1.Width = 263;
                        ImageButton1.Height = 140;         
                    }
                    else
                    {
                        ImageButton1.Visible = false;
                    }
                    if (ds.Tables["DataSet1"].Rows[0]["IMAGE2_PATH"].ToString() != "")
                    {
                        ImageButton2.ImageUrl = "http://mail.nepes.co.kr" + ds.Tables["DataSet1"].Rows[0]["IMAGE2_PATH"].ToString();
                        ImageButton2.Width = 263;
                        ImageButton2.Height = 140;         
                    }
                    else
                    {
                        ImageButton2.Visible = false;
                    }
                    if (ds.Tables["DataSet1"].Rows[0]["IMAGE3_PATH"].ToString() != "")
                    {
                        ImageButton3.ImageUrl = "http://mail.nepes.co.kr" + ds.Tables["DataSet1"].Rows[0]["IMAGE3_PATH"].ToString();
                        ImageButton3.Width = 263;
                        ImageButton3.Height = 140;         
                    }
                    else
                    {
                        ImageButton3.Visible = false;
                    }
                }
                GridView1.DataSource = ds.Tables["DataSet1"];
                GridView1.DataBind();

                if (ds.Tables["DataSet1"].Rows[0]["FROM_FD_YN"].ToString() == "미대상" || ds.Tables["DataSet1"].Rows[0]["OUT_DT"].ToString() == "")
                {
                    FpSpread1.Visible = false;
                    table2.Visible = false;
                    pnl2.Visible = false;
                }
                else
                {
                    FpSpread1.Visible = true;
                    table2.Visible = true;
                    pnl2.Visible = true;


                    // 프로시져 실행: 기본데이타 생성
                    sql_conn.Open();
                    sql_cmd2 = sql_conn.CreateCommand();
                    sql_cmd2.CommandType = CommandType.Text;
                    sql_cmd2.CommandText = getDtl();

                    DataTable dt2 = new DataTable();
                    try
                    {
                        SqlDataAdapter da2 = new SqlDataAdapter(sql_cmd2);
                        da2.Fill(ds, "DataSet2");

                    }
                    catch (Exception ex)
                    {
                        if (sql_conn.State == ConnectionState.Open)
                            sql_conn.Close();
                    }
                    sql_conn.Close();

                    gridCryOut();
                    FpSpread1.DataSource = ds.Tables["DataSet2"];
                    FpSpread1.DataBind();
                }
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();

                if (sql_conn_NEPES.State == ConnectionState.Open)
                    sql_conn_NEPES.Close();
            }
        }
        private string getSQL()
        {
            string strProcOid = Session["process_instance_oid"].ToString();
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("USP_DIST_12_HED  '" + strProcOid + "' \n");
            return sbSQL.ToString();

        }

        private string getDtl()
        {
            string strProcOid = Session["process_instance_oid"].ToString();
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("USP_DIST_13_DTL  '" + strProcOid + "' \n");
            return sbSQL.ToString();
        }

        protected void btn_chkYES_Click(object sender, EventArgs e)
        {
            if(STS.Text == "승인대기")
            {
                MessageBox.ShowMessage("반출상태: [승인대기]시 반출할 수 없습니다.", this.Page);
                return;
            }

            StringBuilder sbSQL = new StringBuilder();
            
            string PROCOID = Session["process_instance_oid"].ToString(); //PROCESS_INSTANCE_OID
            string USERID = Session["User"].ToString();

            sbSQL.Append("UPDATE TB_CERTLOGTRAN_INFO \n");
            sbSQL.Append("SET OUT_DT = GETDATE()     \n");
            sbSQL.Append("   ,OUT_COF_USER ='" + USERID + "' \n");
            sbSQL.Append("   ,UPDT_USER_ID ='" + USERID + "' \n");
            sbSQL.Append("   ,UPDT_DT = GETDATE()  \n");
            sbSQL.Append("   ,PROGEM_ID ='sm_sb001_A02'  \n");
            sbSQL.Append("WHERE  PROCESS_INSTANCE_OID = '" + PROCOID + "' \n");
            
            string setSQL = sbSQL.ToString();
            if (OUT_DT.Text == "")
            {
                if (QueryExecute(setSQL) < 0)
                {
                    MessageBox.ShowMessage("반출정보에 오류가 있습니다.", this.Page);
                    return;
                }
                MessageBox.ShowMessage("반출정보가 확정되었습니다. .", this.Page);

                GridView1.DataSource = null;
                GridView1.DataBind();

                //REFLASH
                search();
            }
            else
            {
                MessageBox.ShowMessage("이미 반출확정 정보가 있습니다. .", this.Page);
            }
        }


        public int QueryExecute(string sql)
        {
            sql_conn.Open();
            sql_cmd = sql_conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = sql;

            try
            {
                value = sql_cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
                value = -1;
            }
            sql_conn.Close();
            return value;
        }

        protected void btn_chkCNL_Click(object sender, EventArgs e)
        {
            StringBuilder sbSQL = new StringBuilder();
            string PROCOID = Session["process_instance_oid"].ToString(); //PROCESS_INSTANCE_OID
            string USERID = Session["User"].ToString();

            sbSQL.Append("SELECT IN_DT_01 FROM TB_CERTLOGTRAN_INFO \n");
            sbSQL.Append("  WHERE PROCESS_INSTANCE_OID = '" + PROCOID + "' \n");
            sbSQL.Append("  AND IN_DT_01 IS NOT NULL                       \n");
            string setSQL = sbSQL.ToString();

            DataSet ds_chk = new DataSet();
            // 프로시져 실행: 기본데이타 생성
            sql_conn.Open();
            sql_cmd_chk = sql_conn.CreateCommand();
            sql_cmd_chk.CommandType = CommandType.Text;
            sql_cmd_chk.CommandText = setSQL;

            DataTable dt = new DataTable();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sql_cmd_chk);
                da.Fill(ds_chk, "DataSet1");
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }
            sql_conn.Close();
            

            if (ds_chk.Tables["DataSet1"].Rows.Count > 0)
            {
                MessageBox.ShowMessage("이미 반입정보가 있습니다. 반입정보 취소후 삭제가 가능합니다.", this.Page);
                return;
            }
            else
            {
                sbSQL.Clear();//초기화
                sbSQL.Append("UPDATE TB_CERTLOGTRAN_INFO \n");
                sbSQL.Append("SET OUT_DT = NULL       \n");
                sbSQL.Append("   ,OUT_COF_USER = NULL \n");                
                sbSQL.Append("   ,UPDT_USER_ID ='" + USERID + "' \n");
                sbSQL.Append("   ,UPDT_DT = GETDATE()  \n");
                sbSQL.Append("   ,PROGEM_ID ='sm_sb001_A02'  \n");
                sbSQL.Append("WHERE PROCESS_INSTANCE_OID = '" + PROCOID + "' \n");                

                string setSQL_Del = sbSQL.ToString();

                if (QueryExecute(setSQL_Del) < 0)
                {
                    MessageBox.ShowMessage("반출정보에 오류가 있습니다.", this.Page);
                    return;
                }
                MessageBox.ShowMessage("반출확정 정보가 취소되었습니다.", this.Page);

                GridView1.DataSource = null;
                GridView1.DataBind();

                //REFLASH
                search();
            }
        }

        private void gridCryOut()
        {   
            FpSpread1.Columns.Count = 21;
            FpSpread1.Rows.Count = 10;
            FpSpread1.ActiveSheetView.AutoPostBack = true;


            FpSpread1.Columns[0].DataField = "GOODS_NM";
            FpSpread1.Columns[1].DataField = "AMOUNT";
            FpSpread1.Columns[2].DataField = "UNIT";
            FpSpread1.Columns[3].DataField = "OUT_DT";
            FpSpread1.Columns[4].DataField = "IN_DT_01";
            FpSpread1.Columns[5].DataField = "IN_DT_AMT_01";
            FpSpread1.Columns[6].DataField = "IN_DT_USER_01";
            FpSpread1.Columns[7].DataField = "IN_DT_02";
            FpSpread1.Columns[8].DataField = "IN_DT_AMT_02";
            FpSpread1.Columns[9].DataField = "IN_DT_USER_02";
            FpSpread1.Columns[10].DataField = "IN_DT_03";
            FpSpread1.Columns[11].DataField = "IN_DT_AMT_03";
            FpSpread1.Columns[12].DataField = "IN_DT_USER_03";
            FpSpread1.Columns[13].DataField = "IN_DT_04";
            FpSpread1.Columns[14].DataField = "IN_DT_AMT_04";
            FpSpread1.Columns[15].DataField = "IN_DT_USER_04";
            FpSpread1.Columns[16].DataField = "IN_DT_05";
            FpSpread1.Columns[17].DataField = "IN_DT_AMT_05";
            FpSpread1.Columns[18].DataField = "IN_DT_USER_05";
            FpSpread1.Columns[19].DataField = "PROCESS_INSTANCE_OID";
            FarPoint.Web.Spread.TextCellType tct = new FarPoint.Web.Spread.TextCellType();
            FpSpread1.Columns[19].CellType = tct;
            FpSpread1.Columns[20].DataField = "GOODS_INDEX";            

            FpSpread1.Columns[0].Label = "품  명";
            FpSpread1.Columns[1].Label = "수  량";
            FpSpread1.Columns[2].Label = "단  위";
            FpSpread1.Columns[3].Label = "반출일";            
            FpSpread1.Columns[4].Label = "1차반입일";
            FpSpread1.Columns[5].Label = "반입수량";
            FpSpread1.Columns[6].Label = "반입자";
            FpSpread1.Columns[7].Label = "2차반입일";            
            FpSpread1.Columns[8].Label = "반입수량";
            FpSpread1.Columns[9].Label = "반입자";
            FpSpread1.Columns[10].Label = "3차반입일";
            FpSpread1.Columns[11].Label = "반입수량";
            FpSpread1.Columns[12].Label = "반입자";
            FpSpread1.Columns[13].Label = "4차반입일";
            FpSpread1.Columns[14].Label = "반입수량";
            FpSpread1.Columns[15].Label = "반입자";
            FpSpread1.Columns[16].Label = "5차반입일";            
            FpSpread1.Columns[17].Label = "반입수량";
            FpSpread1.Columns[18].Label = "반입자";
            FpSpread1.Columns[19].Label = "OID";
            FpSpread1.Columns[20].Label = "OID_INDEX";

            FpSpread1.Columns[0].Width = 80;
            FpSpread1.Columns[1].Width = 40;
            FpSpread1.Columns[2].Width = 30;
            FpSpread1.Columns[3].Width = 120;
            FpSpread1.Columns[4].Width = 120;
            FpSpread1.Columns[5].Width = 40;
            FpSpread1.Columns[6].Width = 60;
            FpSpread1.Columns[7].Width = 120;
            FpSpread1.Columns[8].Width = 40;
            FpSpread1.Columns[9].Width = 60;
            FpSpread1.Columns[10].Width = 120;
            FpSpread1.Columns[11].Width = 40;
            FpSpread1.Columns[12].Width = 60;
            FpSpread1.Columns[13].Width = 120;
            FpSpread1.Columns[14].Width = 40;
            FpSpread1.Columns[15].Width = 50;
            FpSpread1.Columns[16].Width = 120;
            FpSpread1.Columns[17].Width = 40;
            FpSpread1.Columns[18].Width = 60;
            FpSpread1.Columns[19].Width = 40;
            FpSpread1.Columns[20].Width = 60;

            for (int c = 0; c < FpSpread1.Columns.Count; c++)
            {
                SetSpreadColumnLock(c);
            }            
        }
        private void SetSpreadColumnLock(int column)
        {
            FpSpread1.ActiveSheetView.Protect = true;
            FpSpread1.ActiveSheetView.LockForeColor = Color.Black;
            FpSpread1.ActiveSheetView.Columns[column].Font.Name = "돋움체";
            FpSpread1.ActiveSheetView.Columns[column].Font.Size = 9;
            FpSpread1.ActiveSheetView.Columns[column].VerticalAlign = VerticalAlign.Middle;
            FpSpread1.ActiveSheetView.Columns[column].HorizontalAlign = HorizontalAlign.Center;            
        }

        protected void FpSpread1_CellClick(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            string row = e.CommandArgument.ToString();
            //{X=2,Y=3}
            row = row.Replace("{X=", "");
            row = row.Replace("Y=", "");
            string[] arryRow = row.Split(',');
            int rr = Convert.ToInt32(arryRow[0]);

            txt_item.Text = FpSpread1.Sheets[0].Cells[rr, 0].Text;
            labIndex.Text = FpSpread1.Sheets[0].Cells[rr, 20].Text;

            //DDL_PROC_TextChanged(null, null);
            /*선택된 row 색 칠하기*/
            for (int c = 0; c < FpSpread1.Rows.Count; c++)
            {
                FpSpread1.Rows[c].BackColor = Color.Empty;
            }
            FpSpread1.Rows[rr].BackColor = Color.LightPink;


            crtOutAMT.Focus();
        }


        protected void btn_cryYES_Click(object sender, EventArgs e)
        {
            if (labIndex.Text.Trim() != "" && crtOutAMT.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                string PROCOID = Session["process_instance_oid"].ToString(); //PROCESS_INSTANCE_OID
                string stIndex = labIndex.Text.Trim(); //index                         
                string AMT = crtOutAMT.Text.Trim(); //AMT
                string USERID = Session["User"].ToString(); //USERID

                sbSQL.Append("USP_DIST_14_UDT  '" + PROCOID + "','" + stIndex + "','" + AMT + "','" + USERID + "' \n");
                string setSQL = sbSQL.ToString();

                if (QueryExecute(setSQL) < 0)
                {
                    MessageBox.ShowMessage("반입정보 등록에 오류가 있습니다.", this.Page);
                    return;
                }
                MessageBox.ShowMessage("반입정보가 등록되었습니다.", this.Page);

                GridView1.DataSource = null;
                GridView1.DataBind();

                //REFLASH
                search();
                //txt_item.Text = "";  //초기화
                crtOutAMT.Text = ""; //초기화
            }else
            {
                MessageBox.ShowMessage("반입정보 : 수량 / 품명선택을 확인하세요", this.Page);
            }

            
        }

        protected void btn_cryCNL_Click(object sender, EventArgs e)
        {
            if (labIndex.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                string PROCOID = Session["process_instance_oid"].ToString(); //PROCESS_INSTANCE_OID
                string DOC_NUM = DOC_NO.Text.Trim(); //반출번호 
                string stIndex = labIndex.Text.Trim(); //index         

                sbSQL.Append("USP_DIST_14_DEL  '" + PROCOID + "','" + stIndex  + "' \n");
                string setSQL = sbSQL.ToString();

                if (QueryExecute(setSQL) < 0)
                {
                    MessageBox.ShowMessage("반입정보 취소에 오류가 있습니다.", this.Page);
                    return;
                }
                MessageBox.ShowMessage("반입정보 마지막 차수가 취소되었습니다.", this.Page);

                GridView1.DataSource = null;
                GridView1.DataBind();

                //REFLASH
                search();
                //txt_item.Text = "";  //초기화
                crtOutAMT.Text = ""; //초기화
            }
            else
            {
                MessageBox.ShowMessage("반입정보 : 품명선택을 확인하세요", this.Page);
            }
        }

        protected void btn_exit_Click(object sender, EventArgs e)
        {
            //Response.Write("<script>window.opener.location.reload();window.close();</script>");
            //Response.Write("<script language='javascript'>window.close();window.opener.location.href = 'sm_sb001_A01.aspx'</script>");
            Response.Write("<script language='javascript'>window.close();opener.document.getElementById('btn_SEARCH').click();</script>");
        }

        protected void ImageButton1_Click(object sender, ImageClickEventArgs e)
        {            
            ImageButton button1 = (ImageButton)sender;
            string url = button1.ImageUrl.ToString();

            Response.Write("<script>window.name='popup'</script>");
            Response.Write("<script>window.open('" + url + "','','resizable=yes, scrollbars=yes,  top=10,left=10,width=1024,height=768')</script>");            
        }

        protected void btnAdd_Click(object sender, EventArgs e)
        {
            int crtOutCt = 1;
            if (crtOutAMT.Text != "")
            {
                crtOutCt = Convert.ToInt32(crtOutAMT.Text) + 1;
                crtOutAMT.Text = crtOutCt.ToString();
            }
            else
            {
                crtOutAMT.Text = crtOutCt.ToString();
            }
            
        }

        protected void btnMin_Click(object sender, EventArgs e)
        {
            int crtOutCt = 0;
            if (crtOutAMT.Text != "" && crtOutAMT.Text != "0")
            {            
                crtOutCt = Convert.ToInt32(crtOutAMT.Text) - 1;
                crtOutAMT.Text = crtOutCt.ToString();
            }
            else
            {
                crtOutAMT.Text = crtOutCt.ToString();
            }
        }

        protected void btn_cryAllYes_Click(object sender, EventArgs e)
        {
            StringBuilder sbSQL = new StringBuilder();
            string PROCOID = Session["process_instance_oid"].ToString(); //PROCESS_INSTANCE_OID                
            string USERID = Session["User"].ToString(); //USERID

            sbSQL.Append("USP_DIST_14_UDTALL  '" + PROCOID + "','" + USERID + "' \n");
            string setSQL = sbSQL.ToString();

            if (QueryExecute(setSQL) < 0)
            {
                MessageBox.ShowMessage("반입정보 등록에 오류가 있습니다.", this.Page);
                return;
            }
            MessageBox.ShowMessage("모든 반입품목의 정보가 등록되었습니다.", this.Page);

            GridView1.DataSource = null;
            GridView1.DataBind();

            //REFLASH
            search();
            //txt_item.Text = "";  //초기화
            crtOutAMT.Text = ""; //초기화
        }
    }
}
