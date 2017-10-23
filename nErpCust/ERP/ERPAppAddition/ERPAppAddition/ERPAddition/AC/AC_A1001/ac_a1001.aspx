<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ac_a1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AC.AC_A1001.ac_a1001" %>
<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>프로젝트별 공사손익조회</title>
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
        .auto-style1 {
            width: 116px;
        }
        .auto-style2 {
            width: 133px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
        }
        .auto-style3 {
            height: 28px;
        }
        .auto-style4 {
            width: 100%;
            height: 28px;
        }
        .auto-style5 {
            width: 240px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
        }
        .textBox {            
            text-align: center;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label2" runat="server" Text="프로젝트별 공사손익조회" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>        
        
    </div>
    <div>
    <table style="border: thin solid #000080; ">
        <tr>
            <td class="auto-style2" >
                <strong>조회년월</strong>
            </td>
            <td class="auto-style5">
                <asp:TextBox class="textBox" ID="txtFrom" runat="server" Width="100px" MaxLength="6"></asp:TextBox>
                ~<asp:TextBox class="textBox" ID="txtTo" runat="server" Width="100px" MaxLength="6"></asp:TextBox>
            </td>
            <td class="auto-style1">
        <asp:Button ID="Button1" runat="server" Text="조회" Width="120px" 
            onclick="bt_retrieve_Click" />
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
                <asp:AsyncPostBackTrigger ControlID="Button1">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>         
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                <br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
       <cc1:ModalPopupExtender ID="ModalProgress" runat="server" 
            PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        <asp:Panel ID="Panel_Default_Btn" runat="server">
        </asp:Panel>        
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="934px" AsyncRendering="False"
            Height="430px" SizeToReportContent="True" WaitControlDisplayAfter="600000" >
        </rsweb:ReportViewer>
    </div>
    </form>
</body>
</html>
