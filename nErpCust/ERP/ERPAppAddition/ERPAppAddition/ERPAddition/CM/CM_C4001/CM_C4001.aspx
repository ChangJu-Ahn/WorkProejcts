<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CM_C4001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.CM.CM_C4001.CM_C4001" %>

<%@ Register Src="~/Controls/MultiCheckCombo.ascx" TagName="MultiCheckCombo" TagPrefix="mcc" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>BOM대비사용량현황(nepes)</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="../../../Styles/Site_display.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .style0 {
            background-color: #99CCFF;
            font-weight: bold;
            text-align: center;
            width: 60px;
            white-space: nowrap;
        }

        .required {
            BACKGROUND-COLOR: #ffffb4;
        }

        .BasicTb {
            width: auto;
        }

        .updateProgress {
        }
    </style>
    <script type="text/javascript">

        function Val_Check() {
            var plant = document.getElementById("ddl_Plant").value;
            var txtFrom = document.getElementById("txtdate_From").value;
            var txtTo = document.getElementById("txtdate_To").value;

            if (!txtValidCheck(plant)) {
                alert("공장코드는 필수입니다.");
                return false;
            }
            else if (!txtValidCheck(txtFrom) || !txtValidCheck(txtTo)) {
                alert("조회기간은 필수입니다. (ex: 201605 ~ 201605)");
                return false;
            }
            else if (txtFrom > txtTo) {
                alert("날짜조건이 잘못되었습니다.");
                return false;
            }

            return true;
        }

        function txtValidCheck(id) {
            if (id.length == 0)
                return false;
            else
                return true;
        }

    </script>
</head>
<body>
    <form id="frm1" runat="server">
        <cc1:ToolkitScriptManager runat="Server" ID="ToolkitScriptManager1" />
        <div style="height: 11px;">
        </div>
        <asp:Table ID="Table1" runat="server">
            <asp:TableRow runat="server">
                <asp:TableHeaderCell runat="server">
                    <asp:Image ID="Image4" runat="server" ImageUrl="~/img/folder.gif" />
                </asp:TableHeaderCell>
                <asp:TableHeaderCell runat="server" Width="100%">
                    <asp:Label ID="Label1" runat="server" CssClass="title" Width="100%">BOM대비사용량현황(nepes)</asp:Label>
                </asp:TableHeaderCell>
                <asp:TableCell runat="server"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <table>
            <tr>
                <td>
                    <table style="border: thin double #000080;" class="BasicTb">
                        <tr>
                            <td class="style0">공장</td>
                            <td>
                                <asp:DropDownList ID="ddl_Plant" runat="server" BackColor="Yellow">
                                </asp:DropDownList>
                            </td>
                            <td class="style0">품목계정</td>
                            <td>
                                <mcc:MultiCheckCombo ID="mcc_Acct" runat="server" Width_CheckListBox="232" Width_TextBox="250" RepeatColumns="1" Height_Panel="210" />
                            </td>
                            <td class="style0">조회기간</td>
                            <td>

                                <asp:TextBox ID="txtdate_From" runat="server" BackColor="Yellow" MaxLength="7" Width="50px" CssClass="required"></asp:TextBox>
                                <cc1:CalendarExtender ID="cal_From" runat="server" Enabled="True"
                                    Format="yyyyMM" TargetControlID="txtdate_From">
                                </cc1:CalendarExtender>
                            </td>
                            <td style="font-size: medium; font-weight: 700;">~ </td>
                            <td>
                                <asp:TextBox ID="txtdate_To" runat="server" BackColor="Yellow" MaxLength="7" Width="50px" CssClass="required"></asp:TextBox>
                                <cc1:CalendarExtender ID="cal_to" runat="server" Enabled="True"
                                    Format="yyyyMM" TargetControlID="txtdate_To">
                                </cc1:CalendarExtender>
                            </td>
                        </tr>
                    </table>
                </td>
                <td>
                    <asp:Button runat="server" ID="query" Text="조회" OnClick="btnSelect_Click" OnClientClick="return Val_Check()" Width="80px" />
                    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                        <Triggers>
                            <asp:AsyncPostBackTrigger ControlID="query"></asp:AsyncPostBackTrigger>
                        </Triggers>
                    </asp:UpdatePanel>
                </td>
            </tr>
        </table>
        <div style="height: 10px;">
        </div>
        <asp:UpdateProgress ID="UpdateProgress1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <div style="padding-left: 230px; padding-top: 50px;">
                    <asp:Image ID="Image3_1" runat="server" CssClass="updateProgress" ImageUrl="~/img/loading_spinner.gif" Height="173px" Width="230px" />
                </div>
            </ProgressTemplate>
        </asp:UpdateProgress>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" SizeToReportContent="True" Width="" Height=""></rsweb:ReportViewer>
    </form>
</body>
</html>
