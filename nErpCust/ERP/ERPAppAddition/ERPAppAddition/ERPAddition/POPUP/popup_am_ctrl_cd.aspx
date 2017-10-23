

<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="popup_am_ctrl_cd.aspx.cs" Inherits="ERPAppAddition.ERPAddition.POPUP.popup_am_ctrl_cd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>팝업-관리항목선택</title>
        <style type="text/css">
        .style1
        {
            font-family: Arial; text-align: left; width: 250px;
        }
        .style2
        {
            font-family:Arial ; 
            background-color: #CCCCCC; font-size: small; text-align: center; width: 100px; height: 25px;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server" style = "width: 300; height:500;" >
    <div>
        <table style="border: thin solid #000080">
            <tr>
                <td class="style2">
                    관리항목&nbsp;
                </td>
                <td class="style1">
                    <asp:TextBox ID="tb_crtl_cd" runat="server" Width="50px" ></asp:TextBox>
                    <asp:TextBox ID="tb_ctrl_nm" runat="server" Width="100px" ></asp:TextBox>
                </td>
            </tr>
        </table>
        <asp:Button ID="Button1" runat="server" Text="조회" />
    </div>
    </form>
</body>
</html>
