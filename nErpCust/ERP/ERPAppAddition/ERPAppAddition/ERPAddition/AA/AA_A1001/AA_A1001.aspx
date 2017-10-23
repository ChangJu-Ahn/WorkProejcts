<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AA_A1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AA.AA_A1001.AA_A1001" %>
<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
        <title>감가상각비조회(월별)</title>
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
        .auto-style3 {
            height: 28px;
        }
        .auto-style4 {
            width: 100%;
            height: 28px;
        }
                
        .auto-style8 {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            height: 26px;
        }
        .auto-style9 {
            width: 859px;
            height: 26px;
        }
        
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label2" runat="server" Text="감가상각비조회(월별)" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>   
            
          <table  style="border: thin solid #000080">
            <tr>
                <td class="auto-style8">
                    기&nbsp;&nbsp;&nbsp;&nbsp;간
                </td>
                <td class="auto-style9">
                    &nbsp;<asp:TextBox ID="fr_yyyymmdd" runat="server" BackColor="#FFFFCC" MaxLength="6" Width="130px"></asp:TextBox>
                <cc1:CalendarExtender ID="str_fr_dt_CalendarExtender" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="fr_yyyymmdd">
                </cc1:CalendarExtender>
                <cc1:CalendarExtender ID="CalendarExtender1" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="fr_yyyymmdd">
                </cc1:CalendarExtender>
                    <asp:Button ID="btn_view" runat="server" Text="조회" Height="25px" OnClick="btn_view_Click" style="font-weight: 700; background-color: #FFFF99" Width="69px" />
                </td>
            </tr>

        </table>
   
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
   <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                <br /><br /><br /><br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="934px" AsyncRendering="False"
            Height="430px" SizeToReportContent="True" WaitControlDisplayAfter="600000" >
        </rsweb:ReportViewer>
        
        
        
         </div>
    </form>
</body>
</html>
