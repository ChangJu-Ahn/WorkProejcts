<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_A4002.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A4002.AM_A4002" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>은행이체리스트(LED 직원)</title>
    <style type="text/css">

        .title
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align: left;
            font-weight:bold;
            background-color:#EAEAEA;
            color : Blue;                        
            vertical-align : middle;
            display: table-cell;
            line-height: 25px;
            height: 25px;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div dir="ltr">
    
        <asp:Label ID="Label2" runat="server" Text="은행이체리스트(LED 직원)" CssClass="title" Width="100%"></asp:Label>
    
        <br />
    
        <asp:Panel ID="Panel1" runat="server" BackColor="White" Height="36px" 
            Width="992px" BorderColor="White" BorderWidth="2px">
            <asp:Label ID="Label1" runat="server" 
                style="font-family: 돋움; font-weight: 700; font-size: small" 
                Text="지급일(결의일)"></asp:Label>
            &nbsp;&nbsp;&nbsp;<asp:DropDownList ID="ddlDATE" runat="server" AutoPostBack="true" Height="23px" style="margin-top: 0px" Width="152px">
            </asp:DropDownList>
            &nbsp;
            <asp:Button ID="Load_btn" runat="server" BackColor="#FFFFCC" Font-Bold="True" 
            Font-Size="Small" Height="26px" onclick="Load_btn_Click" Text="조 회" 
            Width="54px" />
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
        </asp:Panel>
    
&nbsp;<br />
        <br />
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana" 
            Font-Size="8pt" InteractiveDeviceInfos="(컬렉션)" WaitMessageFont-Names="Verdana" 
            WaitMessageFont-Size="14pt" Width="605px" SizeToReportContent="True">
            <LocalReport ReportPath="">
            </LocalReport>
        </rsweb:ReportViewer>
    
    </div>
    </form>
</body>
</html>
