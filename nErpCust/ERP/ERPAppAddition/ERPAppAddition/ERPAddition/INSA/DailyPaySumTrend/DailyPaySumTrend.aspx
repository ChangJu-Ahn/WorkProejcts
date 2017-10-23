﻿<%@ Page Title="Daily Report_ 인건비(Trend)" Language="C#" AutoEventWireup="true" CodeBehind="DailyPaySumTrend.aspx.cs" Inherits="ERPAppAddition.ERPAddition.INSA.DailyPaySumTrend.DailyPaySumTrend" %>
<%@ Register Src="~/Controls/MultiCheckCombo.ascx" TagName="MultiCheckCombo" TagPrefix="mcc" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; IE=EmulateIE7 ;charset=utf-8" />
    <link href="../../../Styles/Site_display.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="~/Scripts/jquery-2.1.1.min.js"></script>
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.9.1.js"></script>
    <title></title>
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.9.1.js"></script>
    <style type="text/css">
        .style0 {
            background-color: #99CCFF;
            font-weight: bold;
            text-align: center;
            width: 40px;
        }

        .auto-style1 {
            width: 361px;
        }
    </style>
</head>
<body>
    <form id="form" runat="server">
        <ajaxToolkit:ToolkitScriptManager runat="server" ID="ToolkitScriptManager1" />
        <asp:Table ID="Table1" runat="server" Visible="False">
            <asp:TableRow runat="server">
                <asp:TableHeaderCell runat="server">
                    <asp:Image ID="Image4" runat="server" ImageUrl="~/img/folder.gif" />
                </asp:TableHeaderCell>
                <asp:TableHeaderCell runat="server" Width="100%">
                    <asp:Label ID="Label1" runat="server" CssClass="title" Width="100%">O/S 인건비현황(Daily Trend)</asp:Label>
                </asp:TableHeaderCell>
                <asp:TableCell runat="server"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <table>
            <tr>
                <td>
                    <table style="border: thin double #000080">
                        <tr>
                            <td class="style0">사업부</td>
                            <td>
                                <asp:DropDownList ID="dr_dept" runat="server" AppendDataBoundItems="True" DataSourceID="SqlDataSource1"
                                    DataTextField="TEXT" DataValueField="DEPT" AutoPostBack="True" Width="70px" OnTextChanged="Reprt_mcc_Reset">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:MES_NDMES_MESMGR %>" ProviderName="<%$ ConnectionStrings:MES_NDMES_MESMGR.ProviderName %>"
                                    SelectCommand="SELECT 'ALL' AS TEXT, 'ALL' AS dept FROM SYS.&quot;DUAL&quot; UNION ALL SELECT 'Semi' AS TEXT, 'Semi' AS dept FROM SYS.&quot;DUAL&quot; DUAL_2 UNION ALL SELECT 'Display' AS TEXT, 'Display' AS dept FROM SYS.&quot;DUAL&quot; DUAL_1"></asp:SqlDataSource>
                            </td>
                            <td class="style0" id="td_id" runat="server">실적</td>
                            <td>
                                <asp:DropDownList runat="server" ID="drp_Prod" Visible="false" AutoPostBack="true" OnSelectedIndexChanged="Reprt_Reset">
                                    <asp:ListItem Text="없음" Selected="True" Value="A" />
                                    <asp:ListItem Text="BUMP" Value="B" />
                                    <asp:ListItem Text="P-TEST" Value="C" />
                                    <asp:ListItem Text="TAB" Value="D" />
                                    <asp:ListItem Text="COG(WLCSP)" Value="E" />
                                    <asp:ListItem Text="12BUMP" Value="F" />
                                    <asp:ListItem Text="FOWLP" Value="G" />
                                </asp:DropDownList>
                            </td>
                            <td>
                                <asp:DropDownList runat="server" ID="drp_Prod_sub" Visible="false" AutoPostBack="true" OnSelectedIndexChanged="Reprt_Reset">
                                    <asp:ListItem Text="ALL" Selected="True" Value="ALL" />
                                    <asp:ListItem Text="DDI" Value="DDI" />
                                    <asp:ListItem Text="WLP" Value="WLP" />
                                </asp:DropDownList>
                            </td>
                            <td class="style0">그룹</td>
                            <td>
                                <mcc:MultiCheckCombo ID="mcc_dr_Area" runat="server" Height_Panel="210" RepeatColumns="1" Width_CheckListBox="118" Width_TextBox="136" />
                            </td>
                            <td class="style0">파트</td>
                            <td>
                                <mcc:MultiCheckCombo ID="mcc_dr_Part" runat="server" Height_Panel="210" RepeatColumns="1" Width_CheckListBox="118" Width_TextBox="136" />
                            </td>
                            <td class="style0">공정</td>
                            <td>
                                <mcc:MultiCheckCombo ID="mcc_dr_Oprgrp" runat="server" Height_Panel="210" RepeatColumns="1" Width_CheckListBox="160" Width_TextBox="178" />
                            </td>
                            <td class="style0">기간</td>
                            <td>
                                <asp:TextBox ID="txtCrntDT" runat="server" BackColor="#FFFFCC" MaxLength="10" Width="55" OnInit="InitTxtWorkDate"></asp:TextBox>
                                <ajaxToolkit:CalendarExtender ID="txtCrntDT_CalendarExtender" runat="server" Format="yyyy-MM" TargetControlID="txtCrntDT"></ajaxToolkit:CalendarExtender>
                            </td>
                        </tr>
                    </table>
                </td>
                <td>
                    <asp:Button runat="server" ID="query" Text="조회" OnClick="query_Click" Width="80px" />
                    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                        <Triggers>
                            <asp:AsyncPostBackTrigger ControlID="query"></asp:AsyncPostBackTrigger>
                        </Triggers>
                    </asp:UpdatePanel>
                </td>
            </tr>
        </table>
        <asp:Label ID="viewContent" runat="server"></asp:Label>
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
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" SizeToReportContent="True" Width="" Height=""></rsweb:ReportViewer>
    </form>
</body>
</html>