<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="am_a6001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A6001.am_a6001" %>

<%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %>

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
               
               .style3
        {
            width: 1438px;
            height: 7px;
        }
                       
        .label {
                border-top :0px dashed #Orange;
                border-bottom:0px dashed #Orange;
                background-color:white;
                font-weight: bold;
                font-size:smaller;
                }


       
        .dt
        {
            font-size: small;
            text-align: center;
        }
       
        .style13
        {
            width: 163px;
        }
        .style14
        {
            width: 185px;
        }
        .style15
        {
            width: 56px;
        }
        .style16
        {
            width: 123px;
        }
        .style17
        {
            width: 54px;
        }
       
        .style18
        {
            width: 162px;
        }
       
        .style19
        {
            width: 28px;
        }
        .style20
        {
            width: 129px;
        }
       
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
                    <asp:Label ID="Label4" runat="server" Text="구매비용계정예산등록 관리" CssClass="title" Width="100%"></asp:Label>
    
    </div>
    
    <table style="border: thin solid #000080; height: 31px;">
      
                 <td class="style2">  
                     <asp:Label ID="Label17" runat="server" Text="조회구분" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
                <td class="style3">
                <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small" 
                    RepeatDirection="Horizontal"                      
                    AutoPostBack="True" Width="164px" style="margin-left: 0px; font-weight: 700;" 
                    BackColor="White" Height="16px" 
                        onselectedindexchanged="rbl_view_type_SelectedIndexChanged1">
                    <asp:ListItem Value="A">분류</asp:ListItem>                    
                    <asp:ListItem Value="B">등록</asp:ListItem>
                    <asp:ListItem Value="C">조회</asp:ListItem>
                    </asp:RadioButtonList>                
                        
            </td>
          
     
    </table> 
 

         <asp:Panel ID="Panel1" runat="server" Visible="false" >
        <table>
        
        </table>
  
        <asp:Panel ID="Panel_Header" runat="server">
        <table>
     <div>
     <table style="border: thin solid #000080; height: 67px;"><tr>
       
            <td class="style22"><asp:Label ID="Label3" runat="server" Text="년월(YYYYMM)" CssClass="dt"></asp:Label>
            </td>
            <td class="style14">
                <asp:TextBox ID="tb_yyyymm1" runat="server" MaxLength="6"></asp:TextBox>
            </td>
            <td class="style17">
                <asp:Label ID="Label8" runat="server" CssClass="dt" Text="품목그룹"></asp:Label>
            </td>
            <td class="style18">
                <asp:DropDownList ID="ddl_item_gp" runat="server" 
                    DataSourceID="SqlDataSource1" DataTextField="UD_MINOR_NM" 
                    DataValueField="UD_MINOR_NM">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                    ConnectionString="<%$ ConnectionStrings:nepes_test1 %>" 
                    SelectCommand="SELECT UD_MINOR_NM FROM [B_USER_DEFINED_MINOR] where UD_MAJOR_CD='M0003' union all select '-선택안됨-' order by 1 "></asp:SqlDataSource>
     <td class="style15">
                <asp:Label ID="Label5" runat="server" CssClass="dt" Text="비용계정"></asp:Label>
                </td>
            <td class="style20">
                <asp:DropDownList ID="ddl_acct" runat="server">
                    <asp:ListItem Value="%">-선택안됨-</asp:ListItem>
                    <asp:ListItem Value="1">수선유지비</asp:ListItem>
                    <asp:ListItem Value="2">소모품비</asp:ListItem>
                    <asp:ListItem Value="3">지급수수료</asp:ListItem>
                </asp:DropDownList>
            </td>
           <td class="style19">
                <asp:Label ID="Label6" runat="server" CssClass="dt" Text="금액"></asp:Label>
                </td>
            <td class="style16">
                <asp:TextBox ID="tb_amt" runat="server"></asp:TextBox>
            </td> 
            <td> 
            <asp:Button ID="btn_view" runat="server" Text="조회" Width="100px" 
                    onclick="btn_view_Click" />
                <asp:Button ID="btn_save" runat="server" CommandName="save" Text="저장" 
                    Width="100px" onclick="btn_save_Click1" />
                <asp:Button ID="btn_delete0" runat="server" Text="삭제" Width="100px" 
                    onclick="btn_delete0_Click" />
                    
                </td>    
            </div>
              </table>
            </asp:Panel>

             

             
    <br />
    </asp:Panel>  
      

  <asp:Panel ID="Panel_costset" runat="server" Visible="false">        <table>
        <tr><td>
            <asp:Label ID="Label1" runat="server" Text="코스트센터" CssClass="style13"></asp:Label>
            </td><td></td><td>
                <asp:Label ID="Label2" runat="server" Text="대분류품목그룹" CssClass="style13"></asp:Label>
                <asp:DropDownList ID="ddl_itemgp" runat="server" 
                DataSourceID="SqlDataSource1_ddl_itemgp" DataTextField="UD_MINOR_NM" 
                DataValueField="UD_MINOR_NM">
            </asp:DropDownList>
            <asp:Button ID="btn_itemgp_costset" runat="server" Text="조회" 
                 Width="100px" onclick="btn_itemgp_costset_Click" />
                <br />
                <asp:SqlDataSource ID="SqlDataSource1_ddl_itemgp" runat="server" 
                    ConnectionString="<%$ ConnectionStrings:NEPES_TEST1ConnectionString %>" 
                    SelectCommand="SELECT UD_MINOR_NM FROM [B_USER_DEFINED_MINOR] where UD_MAJOR_CD='M0003' union all select '-선택안됨-' order by 1 "></asp:SqlDataSource>
        </td></tr>
        <tr>
        <td><asp:ListBox ID="lsb_l_costset" runat="server" Height="450px" Width="430px" 
                BackColor="#F6F6F6" SelectionMode="Multiple"></asp:ListBox></td>
        <td><asp:Button ID="btn_move_right" runat="server" Text=">" 
                onclick="btn_move_right_Click1"  />
                    <br />
                    <br />
                    <asp:Button ID="btn_move_left" runat="server" Text="<" 
                onclick="btn_move_left_Click1"  /></td>
        <td>
            <asp:ListBox ID="lsb_r_costset" runat="server" Height="450px" Width="430px" 
                BackColor="#F6F6F6" SelectionMode="Multiple"></asp:ListBox>
        </td>
        </tr></table>
            
        </asp:Panel>
          <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
         <asp:Panel ID="Panel_report" runat="server" Visible="false">  
            <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="600px" 
                SizeToReportContent="True" Width="100%" AsyncRendering="False" 
                style="margin-top: 0px"></rsweb:ReportViewer>
                
 </asp:Panel>
  
    </form>
</body>
</html>