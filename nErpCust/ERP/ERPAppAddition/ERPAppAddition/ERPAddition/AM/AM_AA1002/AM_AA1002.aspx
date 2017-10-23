<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_AA1002.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_AA1002.AM_AA1002" %>

<%@ Register Assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457"
    Namespace="FarPoint.Web.Spread" TagPrefix="FarPoint" %>


<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
     <style type="text/css">

        
        .title
        {
            font-family: 굴림체;
            font-size:10pt;DKS
            text-align: left;
            font-weight:bold;
            background-color:#EAEAEA;
            color : Blue;                        
            vertical-align : middle;
            display: table-cell;
            line-height: 25px;
            height: 25px;
        }
               .style2
        {
            width: 118px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }  
             .style1
        {
            width: 118px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }  
        
        
               .auto-style14 {
             width: 54px;
             height: 7px;
         }

               .auto-style15 {
             height: 7px;
         }

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
       <asp:Label ID="Label1" runat="server" Text="일일운용자금실적 조회(NEPES)" CssClass="title" Width="100%"></asp:Label>
      </td>
            </tr>
        </table>
          <table style="border: thin solid #000080; height: 31px;">
      
                 <td class="style2">  
                     <asp:Label ID="Label17" runat="server" Text="조회구분" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
                <td class="style3">
                <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small" 
                    RepeatDirection="Horizontal" 
                    AutoPostBack="True" Width="368px" style="margin-left: 0px; font-weight: 700;" 
                    BackColor="White" Height="16px" OnSelectedIndexChanged="rbl_view_type_SelectedIndexChanged">
                    <asp:ListItem Value="A">일일 실적조회</asp:ListItem>
                    <asp:ListItem Value="B">월 실적조회</asp:ListItem>
                    <asp:ListItem Value="C">사업부 실적조회</asp:ListItem>
                </asp:RadioButtonList>                
                       
            </td>
                <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
    </table> 
        <asp:Panel ID="Panel2" runat="server">
            <table style="border: thin solid #000080; height: 31px;">
                <tr>
                    <td class="style2">
                        <asp:Label ID="lb_yyyy" runat="server" Text="년도" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                    </td >
                    <td class="auto-style15">
                        <asp:TextBox ID="cmb_yyyy" runat="server" MaxLength="2" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
                    </td>
                    <td class="style2">
                        <asp:Label ID="lb_mm" runat="server" Text="월" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                    </td >
                    <td class="auto-style15">
                        <asp:TextBox ID="txt_mm" runat="server" style="background-color: #FFFF99" Width="80px" MaxLength="2"></asp:TextBox>
                    </td>
                   
                   
                    <td class="auto-style14">
                        <asp:Button ID="Load_btn" runat="server" BackColor="#FFFFCC" Font-Bold="True" Font-Size="Small" Height="26px" Text="조 회" Width="54px" OnClick="Load_btn_Click1" />
                    </td>
                    
            </table>
</asp:Panel>
         

<rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana" 
            Font-Size="8pt" InteractiveDeviceInfos="(컬렉션)" WaitMessageFont-Names="Verdana" 
            WaitMessageFont-Size="14pt" Width="991px" SizeToReportContent="True">
            <LocalReport ReportPath="">
            </LocalReport>
        </rsweb:ReportViewer>

    </div>
    </form>
</body>
</html>
