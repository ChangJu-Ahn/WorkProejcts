<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_A5001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A5001.WebForm1" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>감가상각자산별조회</title>
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
    
        <asp:Label ID="Label2" runat="server" Text="감가상각자산별조회" CssClass="title" Width="100%"></asp:Label>
    
        <br />
    
        
           <asp:Button ID="Load_btn" runat="server" BackColor="#FFFFCC" Font-Bold="True" 
            Font-Size="Small" Height="26px" onclick="Load_btn_Click" Text="조 회" 
            Width="54px" Visible="False" />
       
    <asp:ScriptManager ID="ScriptManager1" runat="server">
                            </asp:ScriptManager>
        <br />
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="100%" 
            AsyncRendering="False" Height="600px" SizeToReportContent="True">
        </rsweb:ReportViewer>
    
    </div>
    </form>
</body>
</html>
