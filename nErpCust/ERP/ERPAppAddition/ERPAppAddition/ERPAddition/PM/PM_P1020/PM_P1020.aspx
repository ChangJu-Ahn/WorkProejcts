<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PM_P1020.aspx.cs" Inherits="ERPAppAddition.ERPAddition.PM.PM_P1020.PM_P1020" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>자품목 투입 정보</title>

    <style type="text/css">
        .style12 {
            width: 60px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            font-size: 10pt;
            text-align: center;
        }

        .style13 {
            width: 350px;
            font-family: 굴림체;
            font-size: 10pt;
        }

        .modalBackground {
            background-color: #CCCCFF;
            filter: alpha(opacity=40);
            opacity: 0.5;
        }

        .modalBackground2 {
            background-color: Gray;
            filter: alpha(opacity=50);
            opacity: 0.5;
        }

        .updateProgress {
            background-color: #ffffff;
            position: absolute;
            width: 180px;
            height: 65px;
        }

        .fixedheadercell {
            FONT-WEIGHT: bold;
            FONT-SIZE: 10pt;
            WIDTH: 200px;
            COLOR: white;
            FONT-FAMILY: Arial;
            BACKGROUND-COLOR: darkblue;
        }

        .fixedheadertable {
            left: 0px;
            position: relative;
            top: 0px;
            padding-right: 2px;
            padding-left: 2px;
            padding-bottom: 2px;
            padding-top: 2px;
        }

        .gridcell {
            WIDTH: 200px;
        }

        .div_center {
            width: 500px; /* 폭이나 높이가 일정해야 합니다. */
            height: 600px; /* 폭이나 높이가 일정해야 합니다. */
            position: absolute;
            top: 50%; /* 화면의 중앙에 위치 */
            left: 50%; /* 화면의 중앙에 위치 */
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="0">
        </asp:ScriptManager>
        <div>
            <table>
                <tr>
                    <td>
                        <asp:Image ID="Image1" runat="server" ImageUrl="~/Img/folder.gif" />
                    </td>
                    <td style="width: 100%; font-weight: 600; font-size: small">
                        <asp:Label ID="Label3" runat="server" Text="자품목 미투입정보 조회" CssClass="title" Width="100%"></asp:Label>
                    </td>
                </tr>
            </table>
        </div>
        <div>
            <table style="border: thin solid #000080">
                <tr>
                    <td class="style12">공&nbsp;&nbsp;장
                    </td>
                    <td>
                        <asp:DropDownList ID="dl_plant_cd" runat="server" AppendDataBoundItems="True" Height="25px" BackColor="Yellow"
                            Width="120px" DataSourceID="SqlDataSource1" DataTextField="PLANT_NM" DataValueField="PLANT_CD" AutoPostBack="True">
                            <asp:ListItem Selected="True"></asp:ListItem>
                        </asp:DropDownList>
                        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                            SelectCommand="SELECT [PLANT_CD], [PLANT_NM] FROM [B_PLANT]"></asp:SqlDataSource>
                    </td>
                    <td>
                        <asp:Button ID="bt_retrieve" runat="server" OnClick="bt_retrieve_Click" Text="조회"
                            Width="70px" />
                    </td>
                </tr>
            </table>
            <div style="font-weight: 600; font-size: small; padding-top:1em;">
                <asp:Label ID="Label2" runat="server" Text="설명 : BOM정보 중 자품목투입정보에 연결되지 않은 정보들을 출력합니다." CssClass="title" Width="100%"></asp:Label>
            </div>
        </div>
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <ContentTemplate>
                <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
                    <ProgressTemplate>
                        <asp:Image ID="Image3_1" runat="server" CssClass="updateProgress" ImageUrl="~/Img/loading9_mod.gif" />
                        <br />
                        <br />
                        <br />
                        <br />
                        <asp:Image ID="Image2_1" runat="server" ImageUrl="~/Img/ajax-loader.gif" />
                    </ProgressTemplate>
                </asp:UpdateProgress>
                <br />
                <br />
                <asp:Label Font-Bold="true" Font-Size="Small" runat="server" ID="lblCnt"></asp:Label>
            </ContentTemplate>
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="bt_retrieve"></asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="869px" AsyncRendering="False"
            Height="390px" SizeToReportContent="True">
        </rsweb:ReportViewer>
    </form>
</body>
</html>
