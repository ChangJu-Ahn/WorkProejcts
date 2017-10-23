using System;
using System.Data;
using System.Text;
//using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;

	public class cls_useSiteCount : Page
	{
        private string _strDB = string.Empty;
        private string _strMenu = string.Empty;

        protected void OnLoad()
        {
            _strDB = "";
        }

        #region 생성자 선언
        public cls_useSiteCount (string strDB, string strMenu)
        {
            _strDB = strDB;
            _strMenu = strMenu;
            
        }

        public cls_useSiteCount()
        {

        }
        #endregion

        #region get, set 프로퍼티
        public string strDB
        {
            set { _strDB = value; }
            get { return _strDB; }
        }
        public string strMenu
        {
            set { _strMenu = value; }
            get { return _strMenu; }
        }
        #endregion

        #region 메뉴 Count 메서드
        public void SetMenuCountInsert()
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

            try
            {
                string sql = string.Empty;
                string[] Menu_Id = _strMenu.Split('/');
                DateTime nowDate = DateTime.Now;

                sql = "INSERT INTO dbo.ERP_SITE_USE_COUNT_TBL (DataBase_Name, SiteFullMenu_Id, Menu_Id,insert_Date) VALUES (@DBNAME, @FULLMENU_ID, @MENU_ID, @DATE)";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.AddWithValue("@DBNAME", _strDB);
                cmd.Parameters.AddWithValue("@FULLMENU_ID", _strMenu);
                cmd.Parameters.AddWithValue("@MENU_ID", Menu_Id[Menu_Id.Length-1]);
                cmd.Parameters.AddWithValue("@DATE", nowDate);

                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch 
            {

            }
            finally
            {
                con.Close();
            }
        }
        #endregion
    }
