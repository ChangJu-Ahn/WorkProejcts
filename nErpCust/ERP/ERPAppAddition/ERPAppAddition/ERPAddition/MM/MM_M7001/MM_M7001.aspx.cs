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

namespace ERPAppAddition.ERPAddition.MM.MM_M7001
{
    public partial class MM_M7001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        string sql;

        protected void Page_Load(object sender, EventArgs e)
        {
            //ReportViewer1.Reset();
            ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", "default_textbox();", true);
            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void bt_retrieve_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();

            if (str_fr_dt.Text == "" || str_to_dt.Text == "" || str_fr_dt.Text == "형식 : YYYY-MM-DD" || str_to_dt.Text == "형식 : YYYY-MM-DD")
            {
                //ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", "fn_Inform('조회일을 선택해주세요.');", true);                
            }
            else
            {
                sql = "  SELECT " +
                             "A.MVMT_RCPT_NO," +
                             "A.GM_NO," +
                             "A.GM_SEQ_NO," +
                             "A.PLANT_CD," +
                             "C.PLANT_NM," +
                             "A.ITEM_CD," +
                             "B.ITEM_NM," +
                             "DBO.UFN_GET_DTLITEMNM (A.PO_NO, A.ITEM_CD, A.TRACKING_NO) AS DESCRIPTION," +
                             "B.SPEC," +
                             "CONVERT(VARCHAR,G.DLVY_DT,120) AS DLVY_DT," +
                             "A.MVMT_DT," +
                             "A.MVMT_UNIT," +
                             "A.MVMT_QTY," +
                             "A.MVMT_BASE_UNIT," +
                             "A.MVMT_BASE_QTY," +
                             "A.IV_QTY," +
                             "A.MVMT_CUR," +
                             "A.MVMT_DOC_AMT," +
                             "A.MVMT_LOC_AMT," +
                             "A.MVMT_SL_CD," +
                             "E.SL_NM," +
                             "A.IO_TYPE_CD," +
                             "F.IO_TYPE_NM," +
                             "A.PO_NO," +
                             "A.PO_SEQ_NO," +
                             "A.PUR_GRP," +
                             "H.PUR_GRP_NM," +
                             "A.BP_CD," +
                             "D.BP_NM," +
                             "F.RET_FLG," +
                             "K.MINOR_NM" +
                        " FROM M_PUR_GOODS_MVMT A" +
                             " INNER JOIN B_ITEM B" +
                              "  ON A.ITEM_CD = B.ITEM_CD" +
                             " INNER JOIN B_PLANT C" +
                              "  ON A.PLANT_CD = C.PLANT_CD" +
                             " INNER JOIN B_BIZ_PARTNER D" +
                              "  ON A.BP_CD = D.BP_CD AND D.BP_TYPE <> 'C'" +
                             " INNER JOIN B_STORAGE_LOCATION E" +
                              "  ON A.MVMT_SL_CD = E.SL_CD" +
                             " INNER JOIN M_MVMT_TYPE F" +
                              "  ON A.IO_TYPE_CD = F.IO_TYPE_CD AND F.RCPT_FLG <> 'Y'" +
                             " LEFT OUTER JOIN M_PUR_ORD_DTL G" +
                              "  ON A.PO_NO = G.PO_NO AND A.PO_SEQ_NO = G.PO_SEQ_NO" +
                             " INNER JOIN B_PUR_GRP H" +
                               " ON A.PUR_GRP = H.PUR_GRP" +
                             " INNER JOIN B_MINOR K ON  A.RET_TYPE =  K.MINOR_CD  AND K.MAJOR_CD    = 'B9017'" +
                       " WHERE     A.ITEM_CD >= ''" +
                        "     AND A.ITEM_CD <= 'ZZZZZZZZZ'" +
                          "   AND A.BP_CD >= ''" +
                          "   AND A.BP_CD <= 'ZZZZZZZZZ'" +
                            " AND A.MVMT_DT >= '" + str_fr_dt.Text + "'" +
                            " AND A.MVMT_DT <= '" + str_to_dt.Text + "'" +
                    //" AND A.MVMT_DT >= '2013-07-01'" +
                    //" AND A.MVMT_DT <= '2014-10-01'" +
                             " AND A.MVMT_SL_CD >= ''" +
                             " AND A.MVMT_SL_CD <= 'ZZZZZZZZZ'" +
                             " AND A.IO_TYPE_CD >= ''" +
                             " AND A.IO_TYPE_CD <= 'ZZZZZZZZZ'" +
                             "AND A.MVMT_RCPT_NO like '" + txt_no.Text.Trim() + "%'" +
                    //"AND A.MVMT_RCPT_NO = 'EX20130731000003'" +
                    " ORDER BY A.MVMT_RCPT_NO ASC," +
                             "A.GM_NO ASC," +
                             "A.GM_SEQ_NO ASC," +
                             "A.PLANT_CD ASC," +
                             "C.PLANT_NM ASC," +
                             "A.ITEM_CD ASC," +
                             "DBO.UFN_GET_DTLITEMNM (A.PO_NO, A.ITEM_CD, A.TRACKING_NO) ASC," +
                             "B.SPEC ASC," +
                             "G.DLVY_DT ASC," +
                             "A.MVMT_DT ASC," +
                             "A.MVMT_UNIT ASC," +
                             "A.MVMT_QTY ASC," +
                             "A.MVMT_BASE_UNIT ASC," +
                             "A.MVMT_BASE_QTY ASC," +
                             "A.IV_QTY ASC," +
                             "A.MVMT_CUR ASC," +
                             "A.MVMT_DOC_AMT ASC," +
                             "A.MVMT_LOC_AMT ASC," +
                             "A.MVMT_SL_CD ASC," +
                             "E.SL_NM ASC," +
                             "A.IO_TYPE_CD ASC," +
                             "F.IO_TYPE_NM ASC," +
                             "A.PO_NO ASC," +
                             "A.PO_SEQ_NO ASC," +
                             "A.PUR_GRP ASC," +
                             "H.PUR_GRP_NM ASC," +
                             "A.BP_CD ASC," +
                             "D.BP_NM ASC," +
                             "F.RET_FLG ASC";

                ds_mm_m7001 dt1 = new ds_mm_m7001();
                ReportViewer1.Reset();
                ReportCreator(dt1, sql, ReportViewer1, "rp_mm_m7001.rdlc", "DataSet1");
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

                //_reportViewer.LocalReport.DisplayName = "REPORT_" + dl_plant_cd.Text.Trim() + "_" + str_fr_dt.Text + "_" + str_to_dt.Text + "_" + RadioButtonList1.SelectedItem.Text + "_" + DateTime.Now.ToShortDateString();
                _reportViewer.LocalReport.DisplayName = "REPORT_" + txt_no.Text.Trim() + "_" + str_fr_dt.Text + "_" + str_to_dt.Text + "_" + DateTime.Now.ToShortDateString();
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