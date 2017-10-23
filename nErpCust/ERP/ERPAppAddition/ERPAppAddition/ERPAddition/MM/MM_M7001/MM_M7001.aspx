<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_M7001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_M7001.MM_M7001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <style type="text/css">
        .style12
        {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            font-size: 10pt;
            text-align: center;
        }
        
        .updateProgress
        {
            background-color: #ffffff;
            position: absolute;
            width: 180px;
            height: 65px;
        }
        
        .div_center
        {
            width: 500px; /* 폭이나 높이가 일정해야 합니다. */
            height: 600px; /* 폭이나 높이가 일정해야 합니다. */
            position: absolute;
            top: 50%; /* 화면의 중앙에 위치 */
            left: 50%; /* 화면의 중앙에 위치 */
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */
        }
        .style13
        {
            width: 304px;
        }
        .title
        {
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
        .t
        {
            font-family: "돋움";
            font-size: 10pt;
            color: #c0c0c0;
            border: 1px solid #999999;
        }
        
        .t1
        {
            font-family: "돋움";
            font-size: 10pt;
            color: #000000;
            font-weight: bold;
            border: 1px solid #999999;
        }
    </style>
    <script language="javascript" type="text/javascript">

        function fn_Inform(field) {

            if (document.getElementById('str_fr_dt').value == document.getElementById('str_fr_dt').defaultValue
                || document.getElementById('str_fr_dt').value == '' ||
                document.getElementById('str_to_dt').value == document.getElementById('str_to_dt').defaultValue
                || document.getElementById('str_to_dt').value == '') {
                alert("조회일을 입력해주세요.");
            }
        }

        function default_textbox() {
            document.getElementById('str_fr_dt').value = "형식 : YYYY-MM-DD";
            document.getElementById('str_to_dt').value = "형식 : YYYY-MM-DD"

            document.getElementById('str_fr_dt').className = "t";
            document.getElementById('str_to_dt').className = "t";
        }

        function clearField(field) {
            if (field.value == field.defaultValue) {
                field.value = '';
            }
        }

        function checkField1(field) {
            if (field.value == '') {
                document.getElementById('str_fr_dt').className = "t1";
            }
        }

        function checkField2(field) {
            if (field.value == '') {
                document.getElementById('str_to_dt').className = "t1";
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
                    <asp:Label ID="Label4" runat="server" Text="예외입고/반품등록 일괄 조회" CssClass="title" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    <div>
        <table style="border: thin solid #000080; width: 910px;">
            <tr>
                <td class="style12">
                    입출고번호
                </td>
                <td class="style13">
                    <asp:TextBox ID="txt_no" runat="server" Height="16px" Width="216px"></asp:TextBox>
                </td>
                <td class="style12">
                    조 회 일
                </td>
                <td>
                    <asp:TextBox ID="str_fr_dt" value = "형식 : YYYY-MM-DD" runat="server" BackColor="#FFFFCC" MaxLength="12"
                        Width="130px" onfocus="clearField(this);" onblur="checkField1(this);" onkeyPress="checkField1(this);"></asp:TextBox>
                    <cc1:CalendarExtender ID="str_fr_dt_CalendarExtender" runat="server" Enabled="True"
                        Format="yyyy-MM-dd" TargetControlID="str_fr_dt">
                    </cc1:CalendarExtender>
                    &nbsp;~&nbsp;
                    <asp:TextBox ID="str_to_dt" value = "형식 : YYYY-MM-DD" runat="server" BackColor="#FFFFCC" MaxLength="12"
                        Width="130px" onfocus="clearField(this);" onblur="checkField2(this);" onkeyPress="checkField2(this);"></asp:TextBox>
                    <cc1:CalendarExtender ID="str_to_dt_CalendarExtender" runat="server" Enabled="True"
                        Format="yyyy-MM-dd" TargetControlID="str_to_dt">
                    </cc1:CalendarExtender>
                </td>
            </tr>
        </table>
        <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="0">
        </asp:ScriptManager>
        <asp:Button ID="bt_retrieve" runat="server" OnClick="bt_retrieve_Click" Text="조회" OnClientClick="fn_Inform(this);"
            Width="100px" />
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
                <asp:AsyncPostBackTrigger ControlID="bt_retrieve"></asp:AsyncPostBackTrigger>
            </Triggers>
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
        <cc1:ModalPopupExtender ID="ModalProgress" runat="server" PopupControlID="UpdateProg1"
            TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
    </div>
    <div>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="934px" AsyncRendering="False"
            Height="430px" SizeToReportContent="True">
        </rsweb:ReportViewer>
    </div>
    </form>
</body>
</html>
