<%@ Page Title="반도체 재공현황(상세)" Language="C#" AutoEventWireup="true" CodeBehind="CurrentWIP2.aspx.cs" Inherits="ERPAppAddition.ERPAddition.WM.CurrentWIP2" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolKit" %>
<%@ Register Src="~/Controls/MultiCheckCombo.ascx" TagName="MultiCheckCombo" TagPrefix="mcc" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; IE=EmulateIE7 ;charset=utf-8" />
    <link href="../../../Styles/Site_display.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="~/Scripts/jquery-2.1.1.min.js"></script>
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.9.1.js"></script>
    <title></title>
    <style type="text/css">
        .style0 {
            text-align: center;
            background-color: #99CCFF;
            font-weight: bold;
        }

        .style1 {
            text-align: center;
            background-color: #FF9A31;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <form id="form" runat="server">
        <ajaxToolKit:ToolkitScriptManager runat="Server" ID="ToolkitScriptManager1" />
        <asp:Timer ID="Timer1" runat="server" Enabled="False"></asp:Timer>
        <% //if (Request.QueryString["Menu"] != "false") Response.Write("<asp:SiteMapPath ID=\"SiteMapPath1\" runat=\"server\" />"); %>
        <asp:Table ID="Table1" runat="server" Visible="False">
            <asp:TableRow runat="server">
                <asp:TableHeaderCell runat="server">
                    <asp:Image ID="Image4" runat="server" ImageUrl="~/img/folder.gif" />
                </asp:TableHeaderCell>
                <asp:TableHeaderCell runat="server" Width="100%">
                    <asp:Label ID="Label1" runat="server" CssClass="title" Width="100%">반도체 재공현황(상세)</asp:Label>
                </asp:TableHeaderCell>
                <asp:TableCell runat="server"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <table>
            <tr>
                <td>
                    <table style="border: thin double #000080;">
                        <tr>
                            <td class="style0">구분</td>
                            <td>
                                <asp:RadioButtonList ID="rdoProdType" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" OnSelectedIndexChanged="rdoProdType_SelectedIndexChanged">
                                    <asp:ListItem Value="DDI" Selected="True" >DDI</asp:ListItem>
                                    <asp:ListItem Value="WLP" >WLP</asp:ListItem>
                                    <asp:ListItem Value="FOWLP" >FOWLP</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td>
                                <asp:CheckBox ID="chkHolded" runat="server" Text="보류 Lot" OnCheckedChanged="chkHolded_CheckedChanged" AutoPostBack="True" Checked="True" />
                            </td>
                            <td class="style0">기간</td>
                            <td>
                                <asp:TextBox ID="txtBackDT" runat="server" size="7"></asp:TextBox>
                                <ajaxToolKit:CalendarExtender ID="txtBackDT_CalendarExtender" runat="server" Format="yyyyMMdd" TargetControlID="txtBackDT"></ajaxToolKit:CalendarExtender>
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlBackTM" runat="server">
                                    <asp:ListItem Selected="True" Value="07">07시</asp:ListItem>
                                    <asp:ListItem Value="15">15시</asp:ListItem>
                                    <asp:ListItem Value="23">23시</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <td class="style1" colspan="2">
                                <asp:CheckBox ID="chkMore" runat="server" Text="More..." AutoPostBack="True" />
                            </td>
                            <td id="tdLotID1" runat="server" class="style0">LOT ID</td>
                            <td id="tdLotID2" runat="server">
                                <asp:TextBox ID="txtLotID" runat="server"></asp:TextBox>
                                <%--<asp:CheckBox ID="chkRetrieved" runat="server" Text="회수 Lot" Checked="true" />--%>
                            </td>
                            <td id="tdCustLotID1" runat="server" class="style0">고객 LOT</td>
                            <td id="tdCustLotID2" runat="server">
                                <asp:TextBox ID="txtCustLotID" runat="server"></asp:TextBox></td>
                        </tr>
                        <tr id="trInquiry2" runat="server">
                            <td class="style0">제품</td>
                            <td colspan="3"><%--<asp:DropDownList ID="operList" runat="server" AppendDataBoundItems="True"></asp:DropDownList>--%>
                                <mcc:MultiCheckCombo ID="mccPartID" runat="server" Width_CheckListBox="400" Width_TextBox="310" />
                            </td>
                            <td class="style0">공정</td>
                            <td colspan="5">
                                <mcc:MultiCheckCombo ID="mccOper" runat="server" Height_Panel="500" RepeatColumns="3" Width_CheckListBox="700" Width_TextBox="310" />
                            </td>
                            <td class="style0">생성코드</td>
                            <td>
                                <mcc:MultiCheckCombo ID="mccCreateCode" runat="server" Height_Panel="180" RepeatColumns="1" Width_CheckListBox="200" Width_TextBox="220" />
                            </td>
                        </tr>
                    </table>
                </td>
                <td>
                    <asp:Button runat="server" ID="query" Text="조회" OnClick="query_Click" Width="120px" />
                    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                        <Triggers>
                            <asp:AsyncPostBackTrigger ControlID="query"></asp:AsyncPostBackTrigger>
                        </Triggers>
                    </asp:UpdatePanel>
                </td>
            </tr>
        </table>
        <asp:UpdateProgress ID="UpdateProgress1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" CssClass="updateProgress" ImageUrl="~/img/loading9_mod.gif" />
                <br />
                <br />
                <br />
                <br />
                <asp:Image ID="Image2_1" runat="server" ImageUrl="~/img/ajax-loader.gif" />
            </ProgressTemplate>
        </asp:UpdateProgress>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" SizeToReportContent="True" Height="" Width=""></rsweb:ReportViewer>
    </form>
</body>
</html>
