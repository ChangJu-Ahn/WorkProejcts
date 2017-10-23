<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A4001.WebForm1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:Label ID="Label1" runat="server" Text="지급일(결의일)"></asp:Label>
&nbsp;<asp:TextBox ID="gl_dt" runat="server"></asp:TextBox>
    <cc1:CalendarExtender ID="gl_dt_CalendarExtender" runat="server" Enabled="True" 
        TargetControlID="gl_dt">
    </cc1:CalendarExtender>
    <br />
&nbsp;
    <br />
&nbsp;&nbsp;
    <asp:Button ID="load_btn" runat="server" Height="20px" Text="조 회" 
        Width="59px" />
    </form>
</body>
</html>
