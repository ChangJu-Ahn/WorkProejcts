using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

 public class MessageBox
{
    public static void ShowMessage(string MessageText, Page MyPage)
    {
        MyPage.ClientScript.RegisterStartupScript(MyPage.GetType(),
            "MessageBox", "alert('" + MessageText.Replace("'", "\'") + "');", true);
    }
}

