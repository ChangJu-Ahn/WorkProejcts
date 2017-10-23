using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.IO;

namespace ERPAppAddition.ERPAddition.MM.MM_M5001
{
    public partial class mm_m5001_image_upload : System.Web.UI.Page
    {


        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btn_img_upload_Click(object sender, EventArgs e)
        {
          
            //SqlCommand cmd_erp = new SqlCommand();

            string fileName = fuControl.PostedFile.FileName;
            fileName = System.IO.Path.GetFileName(fileName);
            fileName = Server.HtmlEncode(fileName);

            int fileSize = fuControl.PostedFile.ContentLength;

            Stream stream = fuControl.FileContent;

            byte[] image = new byte[fuControl.PostedFile.ContentLength];
            stream.Read(image, 0, fuControl.PostedFile.ContentLength);
            stream.Close();

            SqlConnection conn_erp =  new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

            SqlCommand scmd = new SqlCommand("dbo.usp_po_insertimage", conn_erp);
            scmd.CommandType = CommandType.StoredProcedure;

            SqlParameter spIdx = new SqlParameter("@idx", SqlDbType.Int);
            spIdx.Direction = ParameterDirection.Output;

            SqlParameter spFilename = new SqlParameter("@FileName", SqlDbType.NVarChar, 250);
            spFilename.Value = fileName;

            SqlParameter spImage = new SqlParameter("@Image", SqlDbType.Image);
            spImage.Value = image;

            scmd.Parameters.Add(spIdx);
            scmd.Parameters.Add(spFilename);
            scmd.Parameters.Add(spImage);

            conn_erp.Open();
            scmd.ExecuteNonQuery();
            conn_erp.Close();

            int newIdx = (int)spIdx.Value;

            Response.Write("새로운 IDX는 " + newIdx.ToString() + "입니다");

            conn_erp.Dispose();
            scmd.Dispose();

            Response.Redirect(Request.ServerVariables["SCRIPT_NAME"]);
        }
    }
}
