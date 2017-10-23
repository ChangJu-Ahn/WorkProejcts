<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="cm_c7001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.CM.CM_C7001.cm_c7001" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>원가계산현황</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="../../../Styles/Site_display.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css" />
    <script type="text/javascript" src="//code.jquery.com/jquery-1.10.2.js"></script>
    <script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
    <style type="text/css">
        .BasicTb {
            width: auto;
            border: thin double #000080;
        }

        td.tilte {
            background-color: #99CCFF;
            font-weight: bold;
            text-align: center;
            width: 70px;
            white-space: nowrap;
        }

        .ui-progressbar {
            position: relative;
        }
        .auto-style1 {
            height: 24px;
        }
    </style>
    <script type="text/javascript">

        function PopOpenUpdate(pr_No) {
            alert(pr_No);
            return;
        }

        function OutputAlert(content) {
            alert("아래 내용을 관리자에게 문의하세요. \n * 내용 : [" + content + "]");
            return;
        }

        function fn_GetPartner() {
            var PopWidth = 635;
            var PopHeight = 520;
            var dbGubun = document.getElementById("hdndbnm").value;
            var PopNodeUrl = "Pop_MM_MM002.aspx?dbName=" + dbGubun;
            var PopFont = "FONT-FAMILY: '맑은고딕';font-size:15px;";
            var PopParams = new Array(); //별도의 넘길 값은 없으나 형식에 맞추기 위해 배열객체만 선언

            var Retval = window.showModalDialog(PopNodeUrl, PopParams, PopFont + "dialogHeight:" + PopHeight + "px;dialogWidth:" + PopWidth + "px;resizable:no;status:no;help:no;scroll:no;location:no");

            if (Retval != null) {
                document.getElementById("txtPartner").innerText = Retval.toString();
                document.getElementById("hdnPartner").value = Retval.toString();
            }

        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <cc1:ToolkitScriptManager runat="Server" ID="ToolkitScriptManager1" />
        <div style="height: 11px;">
        </div>
        <table>
            <tr align="left">
                <th colspan="2" class="auto-style1">
                <asp:Image runat="server" ImageUrl="~/img/folder.gif" /> 원가 계산 현황</th>
            </tr>
            <tr>
                
               
            </tr>
        </table>
        <div>
            <table class="BasicTb">
                <tr>
                   
                    <td class="tilte">년월(YYYYMM)</td>
                    <td>
                        <asp:TextBox ID="tb_yyyymm" runat="server" Width="150px"></asp:TextBox>
                    </td>
                    <td class="tilte">단계</td>
                    <td>
                        <asp:TextBox ID="tb_workstep" runat="server" MaxLength="10" Width="68px" CssClass="required"></asp:TextBox>
                    </td>
                    <td style="font-size: medium; font-weight: 700;">&nbsp;</td>
                  
                    <td>
                        <asp:Button runat="server" ID="query" Text="조회"  Width="80px" />
                    </td>
                </tr>
            </table>
        </div>
        <div style="height: 10px;">
        </div>
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="query"></asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
        <asp:UpdateProgress ID="UpdateProgress1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <div style="padding-left: 230px; padding-top: 50px;">
                    <asp:Image ID="Image3_1" runat="server" CssClass="updateProgress" ImageUrl="~/img/loading_spinner.gif" Height="173px" Width="230px" />
                </div>
            </ProgressTemplate>
        </asp:UpdateProgress>
        <div>
            <rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana"
                Font-Size="8pt" WaitMessageFont-Names="Verdana"
                WaitMessageFont-Size="14pt" Width="605px" SizeToReportContent="True" PageCountMode="Actual">
            </rsweb:ReportViewer>
        </div>
    </form>
</body>
</html>