<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_s7001_sub.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_s7001.sm_s7001_sub" %>
<%@ Register Assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457"
    Namespace="FarPoint.Web.Spread" TagPrefix="FarPoint" %>


<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>음성 매출현황</title>
    <style type="text/css">
        .style12
        {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size:smaller;
        }
        .style1
        {
            width: 400px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
        .spread
        {
            width: 120px;
            
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size:smaller;
        }
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
        .auto-style3 {
            height: 28px;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table><tr><td class="auto-style3">    
    </div>
    <div>    
     <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
<script type="text/javascript">
    

    Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(beginReq);
    Sys.WebForms.PageRequestManager.getInstance().add_endRequest(endReq);
    function beginReq(sender, args) {
        //show the Popup
        $find(ModalProgress).show()
    }
    function endReq(sender, args) {
        //hide the Popup
        $find(ModalProgress).hide();
    }
        </script>
        <asp:Panel ID="Panel_default" runat="server">
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="95%" AsyncRendering="False" ShowZoomControl="False" SizeToReportContent="True">
        </rsweb:ReportViewer>
        </asp:Panel>        
    </div>
    </form>
</body>
</html>
