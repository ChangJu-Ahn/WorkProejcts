<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_A9002.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A9002.AM_A9002" %>

<%@ Register Assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457"
    Namespace="FarPoint.Web.Spread" TagPrefix="FarPoint" %>


<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .title {
            font-family: 굴림체;
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
            width: 118px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }


        .auto-style2 {
            width: 25px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 6px;
        }

        .auto-style3 {
            height: 6px;
        }
        .auto-style4 {
            width: 25px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 5px;
        }
        .auto-style5 {
            height: 5px;
        }
        .updateProgress {}
    </style>
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
                        <asp:Label ID="Label1" runat="server" Text="일일운용자금실적 조회(AMC)" CssClass="title" Width="100%"></asp:Label>
                    </td>
                </tr>
            </table>
            <table style="border: thin solid #000080; height: 31px;">
                <td class="style2">
                    <asp:Label ID="Label17" runat="server" Text="조회구분" BackColor="#99CCFF"
                        Font-Bold="True" Style="text-align: center; font-size: small"></asp:Label>
                </td>

                <td class="style3">
                    <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small"
                        RepeatDirection="Horizontal"
                        OnSelectedIndexChanged="rbl_view_type_SelectedIndexChanged"
                        AutoPostBack="True" Width="234px" Style="margin-left: 0px; font-weight: 700;"
                        BackColor="White" Height="21px">
                        <asp:ListItem Value="A">일일 실적조회</asp:ListItem>
                        <asp:ListItem Value="B">월 실적조회</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </table>
        </div>

        

<asp:ScriptManager ID ="scriptmanager1" runat ="server"></asp:ScriptManager>

        <asp:Panel ID="Panel_bas_info" runat="server" Visible="False" Width="317px">

            <table style="border: thin solid #000080; height: 31px;">
                <tr>
                    <td class ="style2">
                        <asp:Label ID="lb_yyyy" runat="server" Font-Bold="True" Style="text-align: center; font-size: small" Text="[년](ex:2015)" Visible="False"></asp:Label>
                    </td>
                    <td class="auto-style5">
                        <asp:TextBox ID="txt_yyyy" runat="server" Style="background-color: #FFFF99" Visible="False" Width="80px"></asp:TextBox>
                    </td>
                    <td class ="style2">
                        <asp:Label ID="lb_mm" runat="server" Font-Bold="True" Style="text-align: center; font-size: small" Text="[월](ex:02)"  Visible="False"></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="txt_mm" runat="server" Style="background-color: #FFFF99" Width="80px"  Visible="False"></asp:TextBox>
                    </td>
                     <td>
                        <asp:DropDownList ID="unit" runat="server" Style="background-color: #FFFF99" Width="90px"  Visible="False">
                            <asp:ListItem Value="WON">"원" 단위</asp:ListItem>
                            <asp:ListItem Value="thousand">"천" 단위</asp:ListItem>
                         </asp:DropDownList>
                    </td>
                    <td class="auto-style5">
                            <asp:Button ID="Select_Button" runat="server" BackColor="#FFFFCC" Font-Bold="True" Font-Size="Small" Height="26px" OnClick="Load_btn_Click" Text="조 회" Width="54px" Visible="False"/>
                    </td>
                </tr>
            </table>
        </asp:Panel>

         <br /> <br /> 
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="Select_Button">
                </asp:AsyncPostBackTrigger>
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
         Height="390px" SizeToReportContent="True">
        </rsweb:ReportViewer>

    </form>
</body>
</html>
