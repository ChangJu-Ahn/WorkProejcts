<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CM_C3001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.CM.CM_C3001.CM_C3001" %>


<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>원가 재공 비교 (MES VS ERP)</title>
    <style type="text/css">
        .style1
        {
            font-family: Arial;
            text-align: left;
            width: 300px;
        }
        .style2
        {
            font-family: Arial;
            background-color: #CCCCCC;
            font-size: small;
            text-align: center;
            width: 100px;
            height: 25px;
        }
        .style3
        {
            background-color: #FFFFCC;
        }
        .style4
        {
            font-family: "맑은 고딕";
            font-size: small;
        }
        .style6
        {
            font-size: small;
        }
    </style>
    <script type="text/javascript">
        function confirm_user() {
            if (confirm("Are you sure you want to go home ?") == true)
                return true;
            else
                return false;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table style="border: thin solid #000080; width:100%;" >
            <tr>
                <td class="style2">
                    <asp:Label ID="Label1" runat="server" Text="조회월"></asp:Label>
                </td>
                <td class="style1">
                   <table><tr>                   
                   <td><asp:TextBox ID="tb_fr_dt" runat="server" BackColor="#FFFFCC" Width="100px" 
                        MaxLength="6" CssClass="style3"></asp:TextBox>
                    <cc1:CalendarExtender ID="tb_fr_dt_CalendarExtender" runat="server" Enabled="True"
                        TargetControlID="tb_fr_dt" TodaysDateFormat="YYYY.MM" Format="yyyyMM">
                    </cc1:CalendarExtender>       </td>
                    <td><asp:Label ID="Label5" runat="server" Text="서버:" CssClass="style4"></asp:Label></td>
                   <td>
                       <asp:RadioButtonList ID="rbl_server" runat="server"   
                        RepeatDirection="Horizontal" CssClass="style4" 
                           onselectedindexchanged="rbl_server_SelectedIndexChanged">
                        <asp:ListItem>ERP</asp:ListItem>
                        <asp:ListItem Value="COST" Selected="True">원가</asp:ListItem>
                    </asp:RadioButtonList>   </td>
                   </tr></table>
                    
                             
                </td>
                <td class="style2">
                    <asp:Label ID="Label2" runat="server" Text="공장"></asp:Label>
                </td>
                <td class="style1">
                   <asp:DropDownList ID="ddl_plant_cd" runat="server" AppendDataBoundItems="True" Height="25px"
                        Width="170px" DataSourceID="SqlDataSource1" DataTextField="PLANT_NM" 
                        DataValueField="PLANT_CD" AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                        SelectCommand="SELECT PLANT_CD, PLANT_NM FROM B_PLANT WHERE (VALID_TO_DT &gt;= GETDATE())
UNION ALL
SELECT '%', '=== 전 체 ==='
ORDER BY 1">
                    </asp:SqlDataSource>
                </td>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="Label3" runat="server" Text="조회구분"></asp:Label>
                </td>
                <td class="style1">
                    <asp:RadioButtonList ID="RadioButtonList1" runat="server" 
                        RepeatDirection="Horizontal" CssClass="style4" AutoPostBack="True" 
                        TabIndex="10" 
                        onselectedindexchanged="RadioButtonList1_SelectedIndexChanged">
                        <asp:ListItem Selected="True" Value="view1">차이분조회</asp:ListItem>
                        <asp:ListItem Value="view2">공정별조회</asp:ListItem>
                        <asp:ListItem Value="view3">작지별조회</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td class="style2">
                    <asp:Label ID="Label4" runat="server" Text=""></asp:Label>
                </td>
                <td class="style1">
                    
                    <asp:Button ID="btn_request" runat="server" Text="조회" Width="100px" 
                        onclick="btn_request_Click" />
                </td>
            </tr>
        </table>
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <script type="text/javascript">
            var ModalProgress = '<%= ModalProgress.ClientID %>';

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

        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="btn_request">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel> 
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" /><br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
       <cc1:ModalPopupExtender ID="ModalProgress" runat="server" 
            PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="723px" 
            Height="600px" SizeToReportContent="True" AsyncRendering="False">
        </rsweb:ReportViewer>
    </div>
    </form>
</body>
</html>
