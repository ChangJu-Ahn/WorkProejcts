<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="OM_O2001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.OM.OM_O2001.OM_O2001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>특허리스트조회</title>
    <style type="text/css">
        .dt
        {   font-family: 굴림체;
            font-size:10pt;
            text-align: center;
            margin-left: 0px;
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
        .gridstyle
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align: center;
        }
        .style5
        {
            height: 30px;
            width: 63px;
        }
        .style9
        {
            height: 30px;
            width: 173px;
        }
        .style14
        {
            height: 30px;
        }
        .style15
        {
            height: 30px;
            width: 352px;
        }
        .style19
        {
            width: 352px;
            height: 53px;
        }
        .style20
        {
            width: 63px;
            height: 53px;
        }
        .style21
        {
            width: 173px;
            height: 53px;
        }
        .style22
        {
            height: 30px;
            width: 61px;
        }
        .style23
        {
            width: 61px;
            height: 53px;
        }
        </style>
</head>
<body>
    
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <div>
    <table><tr><td>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td  style="width:100%;"><asp:Label ID="Label2" runat="server" Text="특허리스트" CssClass="title" Width="100%"></asp:Label></td></tr></table>
        
    </div>
    <div>
        <asp:Panel ID="Panel_Header" runat="server">
        <table style="height: 67px"><tr>
            <td class="style22"><asp:Label ID="Label4" runat="server" Text="출원번호" CssClass="dt"></asp:Label>
            </td>
            <td class="style15">
                <asp:TextBox ID="tb_apply_no" runat="server"></asp:TextBox>
            </td>
            <td class="style5">
                <asp:Label ID="Label8" runat="server" CssClass="dt" Text="출 원 인"></asp:Label>
            </td>
            <td class="style9">
                &nbsp;<asp:TextBox ID="tb_apply_comp" runat="server"></asp:TextBox>
            </td>
           
            <td class="style14">
                <asp:Button ID="btn_exe" runat="server" Text="조회"  CssClass="dt" Width="100px" 
                    onclick="btn_exe_Click" /></td>
            </tr>
            <tr>
                <td class="style23">
                    <asp:Label ID="Label7" runat="server" CssClass="dt" Text="출 원 일"></asp:Label>
                </td>
                <td class="style19">
                    <asp:TextBox ID="tb_fr_dt" runat="server" CssClass="dt" MaxLength="8"></asp:TextBox>
                    ~<asp:TextBox ID="tb_to_dt" runat="server" CssClass="dt" MaxLength="8"></asp:TextBox>
                    <cc1:CalendarExtender ID="tb_to_dt_CalendarExtender" runat="server" 
                        Enabled="True" Format="yyyyMMdd" TargetControlID="tb_to_dt">
                    </cc1:CalendarExtender>
                </td>
                <td class="style20">
                    &nbsp;</td>
                <td class="style21">
                    &nbsp;</td>
                </tr>
            </table>
          
            <asp:Panel ID="Panel1" runat="server">
            <asp:Label ID="Label3" runat="server" BackColor="#999999" Height="5px" 
                            style="margin-top: 0px; margin-bottom: 0px" Width="100%"></asp:Label>
            </asp:Panel>
            <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="600px" 
                SizeToReportContent="True" Width="100%" AsyncRendering="False" 
                style="margin-top: 0px"></rsweb:ReportViewer>
            </asp:Panel>
            </div>
    </form>
</body>
</html>
