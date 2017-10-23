﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_s8001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_s8001.sm_s8001" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>비가동 Loss 상각비</title>
    <style type="text/css">
        .style12
        {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size:smaller;
        }
        .style1
        {
            width: 400px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
        .spread
        {
            width: 120px;
            
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size:smaller;
        }
       .title
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align: left;
            font-weight:bold;
            background-color:#EAEAEA;
            color : Blue;                        
            vertical-align : middle;
            display: table-cell;
            line-height: 25px;
            height: 25px;
        }
        .auto-style1 {
            width: 116px;
        }
        .auto-style2 {
            width: 133px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
        }
        .auto-style2_1 {            
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
        }
        .auto-style3 {
            height: 28px;
        }
        .auto-style4 {
            width: 100%;
            height: 28px;
        }
        .auto-style5 {
            width: 167px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label2" runat="server" Text="비가동 Loss 상각비" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>        
        
    </div>
    <div>
    <table style="border: thin solid #000080; ">
        <tr>            
            <td class="auto-style2" >
                <asp:CheckBox ID="chk_date" runat="server" Checked="True" OnCheckedChanged="chk_date_CheckedChanged" AutoPostBack="true"/>
                <strong>조회년월일</strong>
            </td>
            <td class="auto-style5">
                <asp:TextBox ID="str_date" runat="server" BackColor="#FFFFCC" MaxLength="8" 
                    Width="130px"></asp:TextBox>
                <cc1:CalendarExtender ID="str_date_CalendarExtender" runat="server" 
                    Enabled="True" Format="yyyyMMdd" TargetControlID="str_date">
                </cc1:CalendarExtender>
            </td>
            <td class="auto-style2" >
                <asp:CheckBox ID="chk_yyyymm" runat="server" AutoPostBack="true" OnCheckedChanged="chk_yyyymm_CheckedChanged"/>
                <strong>조회년월</strong>
            </td>
            <td class="auto-style5">
                
                <asp:DropDownList ID="tb_yyyymm" runat="server" Height="22px" Width="159px" Enabled="False" >
                </asp:DropDownList>
                
            </td>
            <td class="auto-style1">
        <asp:Button ID="Button1" runat="server" Text="조회" Width="120px" 
            onclick="Button1_Click" />
            </td>                      
        </tr>
    </table>
     <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
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
                <asp:AsyncPostBackTrigger ControlID="Button1">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel> 
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                <br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
       <cc1:ModalPopupExtender ID="ModalProgress" runat="server" 
            PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        <asp:Panel ID="Panel_Default_Btn" runat="server">
        </asp:Panel>        
        <asp:Panel ID="Panel_default" runat="server">
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="95%" 
            AsyncRendering="False" Height="600px">
        </rsweb:ReportViewer>
        </asp:Panel>        
    </div>
    </form>
</body>
</html>