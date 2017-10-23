<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_s4001_select.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_s4001.sm_s4001_select" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
            <table>
                <tr >
                    <td >
                        <asp:Label ID="Label3" runat="server" Text="*Version: " Font-Size="Small"></asp:Label>

                        <asp:DropDownList ID="list_select_version" runat="server">
                            <asp:ListItem>-선택안함-</asp:ListItem>
                            <asp:ListItem>R0</asp:ListItem>
                            <asp:ListItem>R1</asp:ListItem>
                            <asp:ListItem>R2</asp:ListItem>
                        </asp:DropDownList>
                        &nbsp;&nbsp;&nbsp;
                        <asp:Label ID="Label10" runat="server" Font-Size="Small" Text="*년: "></asp:Label>

                        <asp:TextBox ID="txt_select_date_yyyy" runat="server" Height="16px" 
                            Width="57px"></asp:TextBox>
                        
                        &nbsp;&nbsp;&nbsp;<asp:Label ID="Label14" runat="server" Font-Size="Small" 
                            Text="*월: "></asp:Label>
                        &nbsp;<asp:DropDownList ID="txt_select_date_mm" runat="server">
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
                        <asp:Label ID="Label11" runat="server" Font-Size="Small" Text="*품목대분류: "></asp:Label>
                        <asp:DropDownList ID="ddl_itemgp_select" runat="server"  AutoPostBack="true" 
                            DataSourceID="SqlDataSource1_ddl_itemgp" DataTextField="ITEM_GROUP" 
                            DataValueField="ITEM_GROUP" >
                        </asp:DropDownList>
                        &nbsp;&nbsp;<asp:Label ID="Label12" runat="server" Font-Size="Small" 
                            Text="*품목소분류: "></asp:Label>
                        <asp:DropDownList ID="ddl_itemgp_select_amt" runat="server" 
                            DataTextField="ITEM_AMT_GROUP" DataValueField="ITEM_AMT_GROUP">
                        </asp:DropDownList>
                        &nbsp;<asp:Button ID="btn_select" runat="server" 
                         Text="조회" Width="100px"  />
                        
                        </table>
                        <asp:RadioButtonList ID="rdl_qty_amt" runat="server" Font-Size="Small" 
                    Visible="false" RepeatDirection="Horizontal" AutoPostBack="True" 
                     
            BorderColor="#0066CC" BorderStyle="Double" >
            <asp:ListItem Value="A" Selected="True">수량</asp:ListItem>
            <asp:ListItem Value="B">금액</asp:ListItem>
        </asp:RadioButtonList>
           
                            <asp:ScriptManager ID="ScriptManager1" runat="server">
                            </asp:ScriptManager>
           
                        <asp:SqlDataSource ID="SqlDataSource1_ddl_itemgp" runat="server" 
                ConnectionString="<%$ ConnectionStrings:unierpsemi_ccube %>" 
                ProviderName="<%$ ConnectionStrings:unierpsemi_ccube.ProviderName %>" 
                SelectCommand="select item_group from t_device_group union all select '-선택안됨-' from dual where rownum &lt; 2 order by 1">
            </asp:SqlDataSource>
                      <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="600px" 
                SizeToReportContent="True" Width="100%" AsyncRendering="False" 
                style="margin-top: 0px"></rsweb:ReportViewer>
    
    </div>
    </form>
</body>
</html>
