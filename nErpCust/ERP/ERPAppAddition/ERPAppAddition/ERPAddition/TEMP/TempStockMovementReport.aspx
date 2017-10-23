<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="TempStockMovementReport.aspx.cs" Inherits="ERPAppAddition.ERPAddition.TEMP.TempStockMovementReport" %>

<%--<!DOCTYPE html >--%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>음성 재고이동 확인레포트(임시)</title>
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

        .BasicTb {
            width: auto;
            border: thin double #000080;
        }
    </style>
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
                    <asp:Label ID="Label1" runat="server" CssClass="title" Width="100%">음성 재고이동 확인레포트(임시)</asp:Label>
                </asp:TableHeaderCell>
                <asp:TableCell runat="server"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <table>
            <tr>
                <td>
                    <table style="border: thin double #000080;" class="BasicTb">
                        <tr>
                            <td class="style0">품목코드</td>
                            <td>
                                <asp:TextBox ID="txtItem" runat="server"></asp:TextBox>
                            </td>
                            <td class="style0">조회기간</td>
                            <td>
                                <%--<input type="date" id="txtdate_From1" name="txtdate_From1" runat="server" style="background-color:yellow;"/>--%>    <%--this is code which use for html5 calendar --%>
                                <asp:TextBox ID="txtdate_From" name="txtdate_From" runat="server" BackColor="Yellow" MaxLength="9" Width="63px" CssClass="required"></asp:TextBox>
                                <cc1:CalendarExtender ID="cal_From" runat="server" Enabled="True" Format="yyyyMMdd" TargetControlID="txtdate_From">
                                </cc1:CalendarExtender>
                            </td>
                            <td style="font-size: medium; font-weight: 700;">~ </td>
                            <td>
                                <%--<input type="date" id="txtdate_To1" name="txtdate_To1" runat="server" style="background-color:yellow;"/>--%>     <%--this is code which use specific html5 calendar--%>
                                <asp:TextBox ID="txtdate_To" name="txtdate_To" runat="server" BackColor="Yellow" MaxLength="9" Width="63px" CssClass="required"></asp:TextBox>
                                <cc1:CalendarExtender ID="cal_to" runat="server" Enabled="True" Format="yyyyMMdd" TargetControlID="txtdate_To">
                                </cc1:CalendarExtender>
                            </td>
                        </tr>
                    </table>
                </td>
                <td>
                    <asp:Button runat="server" ID="query" Text="조회" OnClick="btnSelect_Click" OnClientClick="return fn_Val_Check()" Width="80px" />
                    <asp:Button runat="server" ID="excelDown" Text="내려받기" OnClick="btnExcelDown_Click" OnClientClick="return fn_ValueGridCheck()" Width="80px" />
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
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="query"></asp:AsyncPostBackTrigger>
            </Triggers>
            <ContentTemplate>
                <div style="padding-left: 10px;">
                    <%--<asp:GridView runat="server" ID="dgList" Font-Size="Small" BorderStyle="Inset" BorderColor="Black" HeaderStyle-BackColor="LightSteelBlue" HeaderStyle-Font-Bold="true">--%>
                    <asp:GridView runat="server" ID="dgList" Font-Size="Small">
                    </asp:GridView>
            </ContentTemplate>
        </asp:UpdatePanel>
    </form>
    <%--This is the starting point of the javascript.--%>
    <script type="text/javascript" language="javascript">
        function fn_Val_Check() {
            var txtFromDt = document.getElementById("txtdate_From").value;
            var txtToDt = document.getElementById("txtdate_To").value;

            if (!fn_TxtValidCheck(txtFromDt) || !fn_TxtValidCheck(txtToDt)) {
                alert("조회기간은 필수입니다. (ex: 20160501 ~ 20160505)");
                return false;
            }
            else if (txtFromDt > txtToDt) {
                alert("날짜조건이 잘못되었습니다.");
                return false;
            }

            return true;
        }

        function fn_TxtValidCheck(id) {
            if (id.length == 0) return false;
            else return true;
        }

        function fn_OutputAlert(content) {
            alert("아래 내용을 관리자에게 문의하세요. \n * 내용 : [" + content + "]");
            return;
        }

        function fn_ValueGridCheck() {
            var tempGrid = document.getElementById("dgList");

            if (tempGrid == null) {
                alert("조회가 먼저 되어야 합니다.");
                return false;
            }
        }

        function fn_alert(text) {
            alert(text);
            alert("재검색을 위해 화면이 새로고침 됩니다.");
            window.location.reload(true);
        }
    </script>
</body>
</html>
