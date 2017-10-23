<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_s5001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_s5001.sm_s50011" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>미수금거래처(NEPES)</title>
    <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css" />
    <script src="//code.jquery.com/jquery-1.10.2.js"></script>
    <script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>

    <style type="text/css">
        .style12 {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            font-size: 10pt;
            text-align: center;
        }

        .style13 {
            width: 380px;
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

        .ModalWindow {
            border: solid1px#c0c0c0;
            background: #f0f0f0;
            padding: 0px10px10px10px;
            position: absolute;
            top: -1000px;
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
            top: 74%; /* 화면의 중앙에 위치 */
            left: 46%; /* 화면의 중앙에 위치 */
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */
        }

        .auto-style1 {
            width: 406px;
        }

        .auto-style2 {
            width: 406px;
            font-family: 굴림체;
            font-size: 10pt;
        }
    </style>
</head>
<body>
    <%--<script src="http://code.jquery.com/jquery-latest.js"></script>--%>
    <form id="form1" runat="server">
        <div>
            <div>
                <table style="border: thin solid #000080">
                    <tr>
                        <td class="style12">기준년도(YYYY)</td>
                        <td>
                            <asp:TextBox ID="tb_yyyy" runat="server" Style="background-color: #FFFFCC"></asp:TextBox>
                        </td>
                        <td class="style12">수 금 처</td>
                        <td class="auto-style2">
                            <asp:TextBox ID="tb_bp_cd" runat="server"></asp:TextBox>
                            <asp:Button ID="bt_bp_cd" runat="server" Text=".."
                                Style="height: 21px" OnClick="bt_bp_cd_Click1" />
                            <asp:TextBox ID="tb_bp_nm" runat="server"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="style12">구&nbsp;&nbsp;&nbsp; 분</td>
                        <td class="style13">
                            <asp:DropDownList ID="DropDownList1" runat="server" Height="22px" Width="109px">
                                <asp:ListItem Selected="True" Value="Value_Detail">상세조회</asp:ListItem>
                                <asp:ListItem Value="Value_Sum">집계조회</asp:ListItem>
                            </asp:DropDownList>
                        </td>
                        <td>
                            <asp:Button ID="bt_retrieve" runat="server" ClientIDMode="Static" Text="조회" Width="120px" OnClick="bt_retrieve_Click" Style="font-weight: 700; background-color: #FFFFCC" BorderStyle="Solid" />
                        </td>
                    </tr>
                </table>

                <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="0"></asp:ScriptManager>
                
                <script type="text/javascript">
                    var ModalProgress = '<%= ModalProgress.ClientID %>';

                    Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(beginReq);
                    Sys.WebForms.PageRequestManager.getInstance().add_endRequest(endReq);

                    function beginReq(sender, args) {
                        $find(ModalProgress).show()
                    }

                    function endReq(sender, args) {
                        $find(ModalProgress).hide();
                    }

                    $(document).keypress(function (e) {
                        if (e.keyCode == 13) {
                            $('#bt_retrieve').click();
                            e.preventDefault();
                        }
                    });
                </script>

                <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="bt_retrieve"></asp:AsyncPostBackTrigger></Triggers>
                </asp:UpdatePanel>
                <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
                    <ProgressTemplate>
                        <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                        <br />
                        <br />
                        <br />
                        <br />
                        <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
                    </ProgressTemplate>
                </asp:UpdateProgress>
                <cc1:ModalPopupExtender ID="ModalProgress" runat="server"
                    PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
                </cc1:ModalPopupExtender>
            </div>
            <div>
                <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="1493px" Height="770px">
                </rsweb:ReportViewer>
            </div>

        </div>
    </form>
</body>
</html>

