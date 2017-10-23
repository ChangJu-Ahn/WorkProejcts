<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="e_edu_gw.aspx.cs" Inherits="ERPAppAddition.ERPAddition.EDU.Edu_GW.e_edu_gw" %>
<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
        <title>교육비용관리 조회</title>
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
        .auto-style5 {
            width: 1090px;
        }
        
        .auto-style8 {
            width: 141px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            height: 26px;
        }
        .auto-style9 {
            width: 1090px;
            height: 26px;
        }
        
        .auto-style10 {
            width: 141px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
        }
        
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label2" runat="server" Text="교육비용관리 조회" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>   
                

          <table  style="border: thin solid #000080">
            <tr>
                <td class="auto-style8">
                    기&nbsp;&nbsp;&nbsp;&nbsp; 간
                </td>
                <td class="auto-style9">
                    <asp:TextBox ID="FR_YYYYMMDD" runat="server" BackColor="#FFFFCC" MaxLength="14" Width="130px"></asp:TextBox>
                <cc1:CalendarExtender ID="CalendarExtender1" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="FR_YYYYMMDD">
                </cc1:CalendarExtender>
                    ~&nbsp;<asp:TextBox ID="TO_YYYYMMDD" runat="server" BackColor="#FFFFCC" MaxLength="14" Width="130px"></asp:TextBox>
                <cc1:CalendarExtender ID="TO_YYYYMMDD_CalendarExtender" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="TO_YYYYMMDD">
                </cc1:CalendarExtender>
                </td>
              
            </tr>
                <td class="auto-style10">
                    사&nbsp;&nbsp;&nbsp;&nbsp; 번</td>
                <td class="auto-style5">
                    <asp:TextBox ID="SNO" runat="server"></asp:TextBox>
                    </td>
                    
                <tr>
                <td class="auto-style8">
                    교육구분</td>
                <td class="auto-style9">
                    <asp:DropDownList ID="EDU_TYPE" runat="server">
                        <asp:ListItem Value="%">-전체조회-</asp:ListItem>
                        <asp:ListItem Value="A1">A1(집합교육)</asp:ListItem>
                        <asp:ListItem Value="A3">A3(전파교육)</asp:ListItem>
                        <asp:ListItem Value="A5">A5(사내강사)</asp:ListItem>
                        <asp:ListItem Value="A7">A7(현업주관교육)</asp:ListItem>
                        <asp:ListItem Value="A9">A9(사내어학교육)</asp:ListItem>
                        <asp:ListItem Value="B1">B1(개별교육)</asp:ListItem>
                        <asp:ListItem Value="B3">B3(단체교육)</asp:ListItem>
                        <asp:ListItem Value="C1">C1(e-Learning)</asp:ListItem>
                        <asp:ListItem Value="C3">C3(독서통신) </asp:ListItem>
                        <asp:ListItem Value="D1">D1(TPM)</asp:ListItem>
                        <asp:ListItem Value="D3">D3(환경안전교육)</asp:ListItem>
                        <asp:ListItem Value="D5">D5(대학원)</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="Button1" runat="server" Text="조 회" BorderColor="Black" Height="25px" style="font-weight: 700; background-color: #FFFFCC; font-size: x-small;" Width="117px" OnClick="Button1_Click" />
                </td>
                          </tr>
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
