using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Data.Odbc;
using Oracle.DataAccess.Client;
using System.Data.SqlClient;
using System.Drawing;

namespace SRL.UserControls
{
    public partial class MultiCheckCombo : System.Web.UI.UserControl
    {
        //public event EventHandler CommitTransaction;
        public string[] InitTextList;// = new string[];

        protected void Page_Load(object sender, EventArgs e)
        {
            //CommitTransaction += new System.EventHandler(Sub_CommitTransaction);
            TextBox1 = new System.Web.UI.WebControls.TextBox();
        }
        public int RepeatColumns
        {
            set
            {
                CheckBoxList1.RepeatColumns = value;
            }
        }
        public bool AutoPostBack
        {
            set
            {
                CheckBoxList1.AutoPostBack = value;
                //TextBox1.AutoPostBack = value;
                //opupControlExtender1.OnHide = "";
            }
        }
        /// <summary>
        /// Set the Width of the CheckBoxList
        /// </summary>
        public int Width_CheckListBox
        {
            set
            {
                //int iRptCol = (CheckBoxList1.RepeatColumns > 1) ? CheckBoxList1.RepeatColumns : 1;
                CheckBoxList1.Width = value;
                Panel1.Width = value + 20;
            }
        }
        public int Height_Panel
        {
            set
            {
                Panel1.Height = value;
            }
        }
        /// <summary>
        /// Set the Width of the Combo
        /// </summary>
        public int Width_TextBox
        {
            set { TextBox1.Width = value; }
            get { return (Int32)TextBox1.Width.Value; }
        }
        public int Height_TextBox
        {
            set { TextBox1.Height = value; }
            get { return (Int32)TextBox1.Height.Value; }
        }
        public string InitText
        {
            set { InitTextList = value.Split('\\'); }
            //get { return lStrInitText; }
        }
        public bool Enabled
        {
            set { TextBox1.Enabled = value; }
        }
        /// <summary>
        /// Set the CheckBoxList font Size
        /// </summary>
        public FontUnit fontSize_CheckBoxList
        {
            set { CheckBoxList1.Font.Size = value; }
            get { return CheckBoxList1.Font.Size; }
        }
        /// <summary>
        /// Set the ComboBox font Size
        /// </summary>
        public FontUnit fontSize_TextBox
        {
            set { TextBox1.Font.Size = value; }            
        }
        /// <summary>
        /// Add Items to the CheckBoxList.
        /// </summary>
        /// <param name="array">ArrayList to be added to the CheckBoxList</param>
        public void AddItems(ArrayList array)
        {
            for (int i = 0; i < array.Count; i++)
            {
                CheckBoxList1.Items.Add(array[i].ToString());
            }
        }
        /// <summary>
        /// Add Items to the CheckBoxList
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="nombreCampoTexto">Field Name of the OdbcDataReader to Show in the CheckBoxList</param>
        /// <param name="nombreCampoValor">Value Field of the OdbcDataReader to be added to each Field Name (it can be the same string of the textField)</param>

        //public void AddItems(OdbcDataReader dr, string textField, string valueField)
        //{
        //    ClearAll();

        //    int i = 0;
        //    while (dr.Read())
        //    {
        //        CheckBoxList1.Items.Add(dr[textField].ToString());
        //        CheckBoxList1.Items[i].Value = i.ToString();
        //        i++;
        //    }
        //}
        public void AddItems(SqlDataReader dr, string textField, string valueField)
        {
            ClearAll();
            AddDefault();

            int i = 2;
            while (dr.Read())
            {
                CheckBoxList1.Items.Add(dr[textField].ToString());
                CheckBoxList1.Items[i].Value = dr[valueField].ToString();
                i++;
            }
        }
        public void AddItems(OracleDataReader dr, string textField, string valueField)
        {
            ClearAll();
            AddDefault();

            int i = 2;
            while (dr.Read())
            {
                CheckBoxList1.Items.Add(dr[textField].ToString());
                CheckBoxList1.Items[i].Value = dr[valueField].ToString(); // i.ToString();
                i++;
            }
        }
        public void AddItems(OracleDataReader dr, string textField, string valueField, string[] arrStrItems)
        {
            ClearAll();
            AddDefault();

            int i = 2;
            while (dr.Read())
            {
                CheckBoxList1.Items.Add(dr[textField].ToString());
                CheckBoxList1.Items[i].Value = dr[valueField].ToString(); // i.ToString();
                i++;
            }
            if (!Page.IsPostBack) AdjustText(arrStrItems);
        }
        /// <summary>
        /// Uncheck of the Items of the CheckBox
        /// </summary>
        public void UnSelectAllItems()
        {
            for (int i = 0; i < CheckBoxList1.Items.Count; i++)
            {
                CheckBoxList1.Items[i].Selected = false;
            }
        }
        protected void AdjustText(string[] arrStrItems)
        {
            if (arrStrItems == null) return;
            //if (this.Count > 0) return;
            for (int i = 0; this.Count != arrStrItems.Length; i++) //i < CheckBoxList1.Items.Count; i++)
            {
                foreach (string strValue in arrStrItems)
                {
                    if (CheckBoxList1.Items[i].Value == strValue)
                    {
                        if (hihSelectedText.Value.Contains(strValue) == false)
                        {
                            CheckBoxList1.Items[i].Selected = true;
                            TextBox1.Text = (TextBox1.Text != "") ? TextBox1.Text + ", " + strValue : strValue;
                            hihSelectedText.Value = "'" + TextBox1.Text.Replace(", ", "', '") + "'";
                            hihSelectedCount.Value = (this.Count + 1).ToString();
                        }
                        break;
                    }
                }
            }
            //Page.ClientScript.RegisterStartupScript(this.GetType(), "111SelectItem", "CheckItem()", true);
        }
        /// <summary>
        /// Delete all the Items of the CheckBox;
        /// </summary>
        public void ClearAll()
        {
            TextBox1.Text = "";
            CheckBoxList1.Items.Clear();
        }
        public void AddDefault()
        {
            //ListItem liAll = new ListItem("전체선택");
            //liAll.Attributes["style"] = "color:blue";
            //CheckBoxList1.Items.Add(liAll);
            CheckBoxList1.Items.Add("전체선택");
            CheckBoxList1.Items[0].Attributes["style"] = "color:blue";
            CheckBoxList1.Items[0].Value = "CHECK_ALL"; // i.ToString();
            CheckBoxList1.Items.Add("선택취소");
            CheckBoxList1.Items[1].Value = "UNCHECK_ALL"; // i.ToString();
            CheckBoxList1.Items[1].Attributes["style"] = "color:red";
        }
        /// <summary>
        /// Get or Set the Text shown in the Combo
        /// </summary>
        public string Text
        {
            get { return hihSelectedText.Value.Replace("'", "") ; }
            set { TextBox1.Text = value; }
        }
        public string SQLText
        {
            get { return hihSelectedText.Value; }
        }
        public int Count
        {            
            get { return hihSelectedCount.Value == "" ? 0 : Convert.ToUInt16(hihSelectedCount.Value); }
        }
        protected void CheckBoxList1_Init(object sender, EventArgs e)
        {
            CheckBoxList1.Attributes.Add("OnClick",
                                String.Format("javascript:CheckItem(this, '{0}', '{1}', '{2}');",
                                TextBox1.ClientID, hihSelectedText.ClientID, hihSelectedCount.ClientID));
            //hihSelectedCount.Value = "0";
        }
        //private void Sub_CommitTransaction(object sender, System.EventArgs e)
        //{
        //    // Code for Commit Activity goes here.
        //    Response.Write("Transaction Commited");
        //}
        public void ClearSQLText() // 실제 SQLText를 초기화해주는 메소드가 없어 생성 (2015-07-24 안창주)
        {
            hihSelectedText.Value = "";
        }

        public Color BackColor// 텍스트박스의 바탕색을 지정 (2016-07-11 안창주)
        {
            set { TextBox1.BackColor = value; }
        }
    }
}