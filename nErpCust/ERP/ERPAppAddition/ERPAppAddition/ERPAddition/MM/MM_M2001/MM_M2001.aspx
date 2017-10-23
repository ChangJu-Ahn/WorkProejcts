.<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_M2001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_M2001.MM_M2001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>MRP 관리</title>
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
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table><tr><td>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td  style="width:100%;"><asp:Label ID="Label2" runat="server" Text="MRP 관리" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>    
    </div>
    <asp:Panel ID="Panel_menu1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <script type="text/javascript">
         function submitComment() {
         var oForm = document.commentToComment; // 보내기 위한 데이터가 존재하는 폼
        oForm.submit();  //다른 웹 페이지로 post 데이터를 보낸다.
        self.close(); // 현재 창을 닫는다.
    }
</script>
    <table style=" table-layout:fixed; border: thin solid #000080; width:100%;">
            <tr>
            <td>
            <table><tr>
            <td class="style12" >
                    <strong>작업선택</strong>
                </td>
                <td class="style1">
                    <asp:RadioButtonList ID="rbtn_work_type" runat="server"  
                        CssClass="default_font_size" RepeatDirection="Horizontal" 
                        AutoPostBack="True" 
                        onselectedindexchanged="rbtn_work_type_SelectedIndexChanged">
                        <asp:ListItem Selected="True" Value="create">생성</asp:ListItem>
                        <asp:ListItem Value="view">조회</asp:ListItem>
                    </asp:RadioButtonList>
                </td>  
                <td class="style12" >
                    <strong>기준년월</strong>
                </td>
                <td class="style1">
                    <asp:TextBox ID="tb_bas_yyyymm" runat="server" MaxLength="6"></asp:TextBox>
                    <cc1:CalendarExtender ID="tb_bas_yyyymm_CalendarExtender" runat="server" 
                        Enabled="True" TargetControlID="tb_bas_yyyymm"  Format="yyyyMM"  >
                    </cc1:CalendarExtender>
                </td>      
                <td class="style12" >
                    <strong>버젼선택</strong>
                </td>
                <td class="style1">
                    <asp:DropDownList ID="ddl_version" runat="server">
                        <asp:ListItem>-선택안함-</asp:ListItem>
                        <asp:ListItem>R0</asp:ListItem>
                        <asp:ListItem>R1</asp:ListItem>
                        <asp:ListItem>R2</asp:ListItem>
                    </asp:DropDownList>
                </td> 
                <td class="style12" >
                    <strong>작성일자</strong>
                </td>
                <td class="style1">
                    <asp:TextBox ID="tb_work_yyyymmdd" runat="server" MaxLength="8"></asp:TextBox>
                    <cc1:CalendarExtender ID="tb_work_yyyymmdd_CalendarExtender" runat="server" 
                        Enabled="True" TargetControlID="tb_work_yyyymmdd"  Format="yyyyMMdd" >
                    </cc1:CalendarExtender>
                    <asp:DropDownList ID="ddl_work_yyyymmdd" runat="server" Visible="False">
                    </asp:DropDownList>
                </td>     
               </tr>
               </table>       
             </td>   
            </tr>
            <tr><td>
                <asp:Button ID="btn_exe" runat="server" Text="생성" Width="100px" 
                    onclick="btn_exe_Click"/> &nbsp;&nbsp;
                <asp:Button ID="btn_save" runat="server" Text="저장"  Width="100px" 
                    onclick="btn_save_Click" />  
                <asp:Button ID="btn_view" runat="server" Text="조회"  Width="100px" 
                    onclick="btn_view_Click" /> &nbsp;&nbsp;
                <asp:Button ID="Btn_fcst_view" runat="server" Text="FCST 조회"  Width="100px" 
                    BackColor="#FFFF99" onclick="Btn_fcst_view_Click" 
                     />   
             </td></tr>
        </table>
        
    </asp:Panel>
    <asp:Panel ID="Panel_body" runat="server">
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="" 
            SizeToReportContent="True" Width="">
        </rsweb:ReportViewer>
    </asp:Panel>
    </form>
</body>
</html>
