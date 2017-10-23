<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="E_N_WAY_ADMIN.aspx.cs" Inherits="ERPAppAddition.ERPAddition.EDU.Nepesway.E_N_WAY_ADMIN" %>
<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
        <title>NEPES WAY 실적조회</title>
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
        
        .auto-style11 {
            width: 859px;
        }
        
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label2" runat="server" Text="nepesway 실적조회" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>   
            <table style="border: thin solid #000080; height: 31px;">
      
                 <td class="auto-style8">  
                     <asp:Label ID="Label17" runat="server" Text="조회구분" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
                <td class="auto-style11">
                <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small" 
                    RepeatDirection="Horizontal" 
                    AutoPostBack="True" Width="306px" style="margin-left: 0px; font-weight: 700;" 
                    BackColor="White" Height="16px">
                    <asp:ListItem Value="A" Selected="True">음악교실</asp:ListItem>
                    <asp:ListItem Value="B">i훈련</asp:ListItem>
                    <asp:ListItem Value="C">마법노트</asp:ListItem>
                </asp:RadioButtonList>                
                        
            </td>
          
     
    </table>    
          <table  style="border: thin solid #000080">
            <tr>
                <td class="auto-style8">
                    기&nbsp;&nbsp;&nbsp;&nbsp;간
                </td>
                <td class="auto-style9">
                    &nbsp;<asp:TextBox ID="fr_yyyymmdd" runat="server" BackColor="#FFFFCC" MaxLength="12" Width="130px"></asp:TextBox>
                <cc1:CalendarExtender ID="str_fr_dt_CalendarExtender" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="fr_yyyymmdd">
                </cc1:CalendarExtender>
                <cc1:CalendarExtender ID="CalendarExtender1" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="fr_yyyymmdd">
                </cc1:CalendarExtender>
                    ~                 
                    <asp:TextBox ID="to_yyyymmdd" runat="server" BackColor="#FFFFCC" MaxLength="12" Width="130px"></asp:TextBox>
                <cc1:CalendarExtender ID="to_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="to_yyyymmdd">
                </cc1:CalendarExtender>
                </td>
              
            </tr>
                <td class="style12">
                   소&nbsp;&nbsp;&nbsp;&nbsp;속</td>
                <td class="auto-style11">
                    <asp:Label ID="lbl_busor1" runat="server" Text="사업부 : "></asp:Label>
                    <asp:DropDownList ID="ddl_busor1" runat="server" DataSourceID="SqlDataSource1" DataTextField="BUSOR_1" DataValueField="BUSOR_1">
                    </asp:DropDownList>&nbsp;&nbsp;
                    &nbsp;&nbsp;
                     &nbsp;&nbsp;
                    &nbsp;&nbsp;
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>" 
                        SelectCommand=
                       "SELECT DISTINCT dbo.BU_SPLIT(bu_name,1,'/')BUSOR_1 
                        FROM INSADB.INBUS.DBO.c_busor 
                        WHERE BU_EDATE='' --AND dbo.BU_SPLIT(bu_name,1,'/')!=''
                        AND dbo.BU_SPLIT(bu_name,1,'/')!=''
                        UNION ALL
                        SELECT '%' 
                        ORDER BY 1
                        "></asp:SqlDataSource>                  
                    </td>

          
            
                <tr>
                <td class="auto-style8">
                    사&nbsp;&nbsp;&nbsp;&nbsp;번
                </td>
                <td class="auto-style9">
                    <asp:TextBox ID="tb_sno" runat="server"></asp:TextBox>
                     &nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btn_view" runat="server" Text="조회" Height="25px" OnClick="btn_view_Click" style="font-weight: 700; background-color: #FFFF99" Width="69px" />
                </td>
                      <tr>
                
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
