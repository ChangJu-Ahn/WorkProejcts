
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_s4002.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_s4002.sm_s4002" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>  FCST 관리</title>
    <style type="text/css">

        .style12
        {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size:smaller;
            table-layout:fixed;
             
        }
        .style1
        {
            
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            table-layout:fixed;
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
        .default_font_size
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align:center;
        }
        .default_font_background
        {
            font-family: 굴림체;
            font-size:10pt;
            
        }
        .style13
        {
            height: 29px;
        }
    </style>
  
</head>
<body>
    <form id="form2" runat="server">
    <div>
        <table>
            <tr>
                <td>
                    <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
                </td>
                <td style="width:100%;">
                    <asp:Label ID="Label2" runat="server" CssClass="title" Text="FCST 관리" 
                        Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    <asp:Panel ID="Panel_menu1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <table style=" table-layout:fixed; border: thin solid #000080; width:100%;">
            <tr>
                <td class="style13">
                    <table>
                        <tr>
                           
                            <td class="style12">
                                <strong>기준년월</strong>
                            </td>
                            <td class="style1">
                                <asp:TextBox ID="tb_bas_yyyymm" runat="server" MaxLength="6" Width="94px" 
                                    BackColor="#FFFFCC"></asp:TextBox>
                            </td>
                            
                            <td class="style12">
                                <strong>버젼선택</strong>
                            </td>
                            <td class="style1">
                                <asp:DropDownList ID="ddl_version" runat="server" BackColor="#FFFFCC">
                                    <asp:ListItem>-선택안함-</asp:ListItem>
                                    <asp:ListItem>R0</asp:ListItem>
                                    <asp:ListItem>R1</asp:ListItem>
                                    <asp:ListItem>R2</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                          
                            
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Button ID="btn_select" runat="server" Text="조회" 
                        Width="100px" onclick="btn_select_Click" />
                    &nbsp;</td>
            </tr>
        </table>
    </asp:Panel>
       <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="100%" 
            AsyncRendering="False" Height="600px" >
        </rsweb:ReportViewer>
    
    </form>
       
     
   
</body>
</html>
