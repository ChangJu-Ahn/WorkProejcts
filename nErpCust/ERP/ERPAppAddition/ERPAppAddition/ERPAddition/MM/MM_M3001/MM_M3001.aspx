﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_M3001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_M3001.MM_M3001" %>

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
        
         .style55
        {
            width: 106px;
            font-family: 굴림체;
            font-size: smaller;
            text-align: center;
            background-color: Silver;
            height: 22px;
            font-weight: 700;
        }
        .style56
        {
            height: 22px;
        }
               
        .style57
        {
            width: 103px;
            font-family: 굴림체;
            font-size: smaller;
            text-align: center;
            background-color: Silver;
            height: 22px;
        }
               
        .style58
        {
            width: 103px;
            font-family: 굴림체;
            font-size: smaller;
            text-align: center;
            background-color: Silver;
            height: 22px;
            font-weight: 700;
        }
        .style59
        {
            width: 242px;
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
       <asp:Label ID="Label1" runat="server" Text="구매 MRP FCST 관리" CssClass="title" Width="100%"></asp:Label>
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
                    onselectedindexchanged="rbl_view_type_SelectedIndexChanged" 
                    AutoPostBack="True" Width="306px" style="margin-left: 0px; font-weight: 700;" 
                    BackColor="White" Height="16px">
                    <asp:ListItem Value="A">FCST 등록</asp:ListItem>
                    <asp:ListItem Value="B">FCST 조회</asp:ListItem>
                </asp:RadioButtonList>                
                        
            </td>
          
     
    </table> 
 <asp:Panel ID="panel_upload" runat="server" Visible="False" BorderStyle="Groove" 
            BorderColor="White" Width="99%">
            <table style="width: 99%">
                <tr >
                 
                     
                   <td class="style58"> 
                        <asp:Label ID="Label13" runat="server" Font-Size="Small" 
                            Text="Version선택" BackColor="Silver" style="font-weight: 700"></asp:Label></td>
                            <td class="style25">
                        <asp:DropDownList ID="list_regist_version" runat="server" BackColor="#FFFFCC" 
                             Width="90px" >
                            <asp:ListItem>-선택안함-</asp:ListItem>
                            <asp:ListItem>R0</asp:ListItem>
                            <asp:ListItem>R1</asp:ListItem>
                            <asp:ListItem>R2</asp:ListItem>
                        </asp:DropDownList></td>
                        <td class=style55>날짜선택</td>
                        <td class="style53">
                           <asp:Label ID="Label15" runat="server" Font-Size="Small" 
                                Text="*년: "></asp:Label>
                            <asp:TextBox ID="txt_regist_date_yyyy" runat="server" BackColor="#FFFFCC" 
                                Height="16px" Width="57px"></asp:TextBox>
                              <asp:Label ID="Label16" runat="server" Font-Size="Small" 
                                Text="*월: "></asp:Label>
                            <asp:DropDownList ID="txt_regist_date_mm" runat="server" BackColor="#FFFFCC" 
                                style="text-align: center">
                                <asp:ListItem>-선택안함-</asp:ListItem>
                                <asp:ListItem Value="01"></asp:ListItem>
                                <asp:ListItem Value="02"></asp:ListItem>
                                <asp:ListItem Value="03"></asp:ListItem>
                                <asp:ListItem Value="04"></asp:ListItem>
                                <asp:ListItem Value="05"></asp:ListItem>
                                <asp:ListItem Value="06"></asp:ListItem>
                                <asp:ListItem Value="07"></asp:ListItem>
                                <asp:ListItem Value="08"></asp:ListItem>
                                <asp:ListItem Value="09"></asp:ListItem>
                                <asp:ListItem Value="10"></asp:ListItem>
                                <asp:ListItem Value="11"></asp:ListItem>
                                <asp:ListItem Value="12"></asp:ListItem>
                            </asp:DropDownList>
                            </td>
                           
                            <td class=style55>Excel선택</td>
                            <td class="style54">
                            <asp:FileUpload ID="FileUpload1" runat="server" BackColor="#FFFFCC" />
                            &nbsp;<asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" 
                                    Text="Upload" ToolTip="엑셀자료를 가져와 화면에 보여준다." Width="87px" Height="21px" />
                            </td>
                            <td class=style57 style="border-style: none; font-weight: 700;">Sheet선택</td>
                            <td class="style56">
                            <asp:DropDownList ID="ddlSheets" runat="server" AutoPostBack="True" 
                                BackColor="#FFFFCC">
                            </asp:DropDownList>
                            <asp:Button ID="btnSave" runat="server" OnClick="btnSave_Click" 
                                OnClientClick="return confirm('당월데이타 존재시 삭제후 저장됩니다. 진행하시겠습니까?');" Text="Save" 
                                    Width="80px" />
                            <asp:Button ID="btnCancel0" runat="server" OnClick="btnCancel_Click" 
                                Text="Cancel" Width="80px" /></td>
                       
                        
                    
                        </tr>
                      
                 
            </table>
             
        </asp:Panel>
         <asp:Panel ID="Panel1" runat="server">
                            <td class="style17">
                                <asp:HiddenField ID="HiddenField_filePath" runat="server" />
                                <asp:HiddenField ID="HiddenField_extension" runat="server" />
                                <asp:HiddenField ID="HiddenField_fileName" runat="server" />
                            </td>
                        </asp:Panel>  
                       
                       
                        <asp:Panel ID="Panel_select" runat="server" Visible="False" BorderStyle="Groove" 
            BorderColor="White" Width="99%">
            <table style="width: 97%">
                <tr >
                   <td class="style58"> 
                        <asp:Label ID="Label5" runat="server" Font-Size="Small" 
                            Text="Version선택" BackColor="Silver" style="font-weight: 700"></asp:Label></td>
                            <td class="style25">
                                <asp:DropDownList ID="list_select_version" runat="server" BackColor="#FFFFCC" 
                                    style="text-align: center">
                                    <asp:ListItem>-선택안함-</asp:ListItem>
                                    <asp:ListItem>R0</asp:ListItem>
                                    <asp:ListItem>R1</asp:ListItem>
                                    <asp:ListItem>R2</asp:ListItem>
                                </asp:DropDownList>
                    </td>
                        <td class=style55>날짜선택</td>
                        <td class="style59">
                            <asp:Label ID="Label6" runat="server" Font-Size="Small" 
                                Text="*년: "></asp:Label>
                            <asp:TextBox ID="txt_select_date_yyyy" runat="server" BackColor="#FFFFCC" 
                                Height="16px" Width="57px"></asp:TextBox>  
                              &nbsp;   <asp:Label ID="Label7" runat="server" Font-Size="Small" 
                                Text="*월: "></asp:Label>
                            <asp:DropDownList ID="txt_select_date_mm" runat="server" BackColor="#FFFFCC" 
                                style="text-align: center">
                                <asp:ListItem>-선택안함-</asp:ListItem>
                                <asp:ListItem Value="01"></asp:ListItem>
                                <asp:ListItem Value="02"></asp:ListItem>
                                <asp:ListItem Value="03"></asp:ListItem>
                                <asp:ListItem Value="04"></asp:ListItem>
                                <asp:ListItem Value="05"></asp:ListItem>
                                <asp:ListItem Value="06"></asp:ListItem>
                                <asp:ListItem Value="07"></asp:ListItem>
                                <asp:ListItem Value="08"></asp:ListItem>
                                <asp:ListItem Value="09"></asp:ListItem>
                                <asp:ListItem Value="10"></asp:ListItem>
                                <asp:ListItem Value="11"></asp:ListItem>
                                <asp:ListItem Value="12"></asp:ListItem>
                            </asp:DropDownList>
                            </td>
                           
                           
                                                     <td class="style56">
                                
                                <asp:Button ID="btn_select" runat="server" Text="조회" 
                                    Width="100px" onclick="btn_select_Click" />
                    </td>
                       
                        
                    
                        </tr>
                      
                 
            </table>
             
        </asp:Panel>
           <asp:Panel ID="Panel_regist_excel_grid" runat="server" Visible="False">
           <asp:GridView ID="grid_regist_excel" runat="server"  CellPadding="4" 
                            GridLines="None" 
                  >
                                <AlternatingRowStyle BackColor="White" />
                                <EditRowStyle BackColor="#2461BF" />
                                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#EFF3FB" />
                                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                <SortedAscendingCellStyle BackColor="#F5F7FB" />
                                <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                                <SortedDescendingCellStyle BackColor="#E9EBEF" />
                                <SortedDescendingHeaderStyle BackColor="#4870BE" />
                            </asp:GridView>
           
                            
           
                        </asp:Panel> 
                        <asp:ScriptManager ID="ScriptManager1" runat="server">
                            </asp:ScriptManager>            
        <asp:Panel ID="Panel_select_excel_qty_grid" runat="server">
     
              <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="100%" 
            AsyncRendering="False" Height="600px" SizeToReportContent="True">
                </rsweb:ReportViewer>
                

                
            </asp:Panel>

    </div>
    </form>
</body>
</html>
