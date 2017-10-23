<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="am_a1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A1001.am_a1001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>손익계산서조회(관리항목별)</title>
    <style type="text/css">
        .style1
        {
            font-family: Arial;
            text-align: left;
            width: 250px;
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
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="border: thin solid #000080">
            <tr>
                <td class="style2">
                    회계일&nbsp;
                </td>
                <td class="style1">
                    <asp:TextBox ID="tb_fr_dt" runat="server" BackColor="#FFFF99" Width="100px" MaxLength="6"></asp:TextBox>
                    <cc1:CalendarExtender ID="tb_fr_dt_CalendarExtender" runat="server" Enabled="True"
                        TargetControlID="tb_fr_dt" TodaysDateFormat="YYYY.MM" Format="yyyyMM">
                    </cc1:CalendarExtender>
                    ~<asp:TextBox ID="tb_to_dt" runat="server" BackColor="#FFFF99" Width="100px" MaxLength="6"></asp:TextBox>
                    <cc1:CalendarExtender ID="tb_to_dt_CalendarExtender" runat="server" Enabled="True"
                        TargetControlID="tb_to_dt" Format="yyyyMM">
                    </cc1:CalendarExtender>
                </td>
                <td class="style2">
                    &nbsp; 사업장
                </td>
                <td class="style1">
                    <asp:DropDownList ID="ddl_biz_area" runat="server" >
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style2">
                    관리항목&nbsp;
                </td>
                <td class="style1">
                    <asp:DropDownList ID="ddl_ctrl_value" runat="server" >
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
    </div>
    &nbsp;<asp:Button ID="bt_retrive" runat="server" Text="조회" 
        Width="100px" onclick="bt_retrive_Click" /><asp:ScriptManager
        ID="ScriptManager1" runat="server">
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
                <asp:AsyncPostBackTrigger ControlID="bt_retrive">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel> 
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                <br /><br /><br /><br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
       <cc1:ModalPopupExtender ID="ModalProgress" runat="server" 
            PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        
    <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="600px" 
        AsyncRendering="False" SizeToReportContent="True">
    </rsweb:ReportViewer>
    </form>
</body>
</html>
