<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_AC1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_AC1001.AM_AC1001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    
    <title>일월자금실적_조회(Display)</title>

    <style type="text/css">
        .title {
            font-family: 돋음;
            font-size: 10pt;
            text-align: left;
            font-weight: bold;
            background-color: #EAEAEA;
            color: Blue;
            vertical-align: middle;
            display: table-cell;
            line-height: 25px;
            height: 25px;
        }

        .style2 {
            width: 60px;
            font-family: 돋음;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }


        .auto-style5 {
            height: 5px;
        }

        .auto-style6 {
            height: 7px;
        }

        .stytitle {
            width: 70px;
            font-family: 돋음;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }

    </style>

    <script language="javascript">

        function checkFourNumber(text) {
            if (text.value.length != 4) {
                alert("년도는 4자리만 입력가능합니다.");
                document.getElementById('txt_yyyy').value = "";
            }
        }

        function checkTooNumber(text) {
            if (text.value.length != 2) {
                alert("월 은 2자리만 입력가능합니다.");
                document.getElementById('txt_mm').value = "";
            }
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table>
                <tr>
                    <td>
                        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
                    </td>
                    <td style="width: 100%;">
                        <asp:Label ID="Label1" runat="server" Text="일월자금실적_조회(Display)" CssClass="title" Width="100%"></asp:Label>
                    </td>
                </tr>
            </table>
        </div>
        <div>
            <table style="border: thin solid #000080; height: 31px;">
                <tr>
                    <td class="style2">
                        <asp:Label ID="Label17" runat="server" Text="조회구분" BackColor="#99CCFF"
                            Font-Bold="True" Style="text-align: center; font-size: small"></asp:Label>
                    </td>
                    <td class="auto-style6">
                        <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small"
                            RepeatDirection="Horizontal"
                            OnSelectedIndexChanged="rbl_view_type_SelectedIndexChanged"
                            AutoPostBack="True" Width="345px" Style="margin-left: 0px; font-weight: 700;"
                            BackColor="White" Height="21px">
                            <asp:ListItem Value="A">일자금 실적</asp:ListItem>
                            <asp:ListItem Value="B">월자금 실적</asp:ListItem>
                            <asp:ListItem Value="C">계획대비 실적분석</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
            </table>
        </div>

        <asp:ScriptManager ID="scriptmanager1" runat="server"></asp:ScriptManager>

        <div>
            <table style="border: thin solid #000050; height: 31px;" runat="server" id="table" visible="false">
                <tr>
                    <td class="stytitle">
                        <asp:Label ID="ld_yyyy" runat="server" Font-Bold="True" Style="text-align: center; font-size: small" Text="[년](ex:1989)" Visible="False"></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox ID="txt_yyyy" runat="server" Style="background-color: #FFFF99" Visible="False" Width="60px" onchange ="checkFourNumber(this)"></asp:TextBox>
                    </td>
                    <td class="stytitle" id="td">
                        <asp:Label ID="ld_mm" runat="server" Font-Bold="True" Style="text-align: center; font-size: small" Text="[월]<br>(ex:03)" Visible="False"></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox ID="txt_mm" runat="server" Style="background-color: #FFFF99" Width="60px" Visible="False" onchange ="checkTooNumber(this)"></asp:TextBox>
                    </td>

                    <td class="auto-style1">
                        <asp:Button ID="Select_Button" runat="server" BackColor="#FFFFCC" Font-Bold="True" Font-Size="Small" Height="26px" OnClick="btn_Select_Click" Text="조 회" Width="54px" Visible="False" />
                    </td>
                </tr>
            </table>
        </div>
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="Select_Button"></asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" Height="75px" Width="179px" />
                <br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
                <br />
            </ProgressTemplate>
        </asp:UpdateProgress>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="869px" AsyncRendering="False"
            Height="1300px" SizeToReportContent="True">
        </rsweb:ReportViewer>

    </form>
</body>
</html>
