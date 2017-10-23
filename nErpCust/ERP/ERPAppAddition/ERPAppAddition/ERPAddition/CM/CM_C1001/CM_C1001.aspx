<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CM_C1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.CM.CM_C1001.CM_C1001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
     <style type="text/css">
     .style3
        {
            width: 150px;
            text-align: center;
            background-color:#6699FF;
        }
         .style4
         {
             color: #FFFFFF;
         }
         </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table>
        <tr>
        <td class="style3" rowspan="1" style="border-width: 0px">
            <strong>
            <asp:Label ID="Label3" runat="server" Text="4M계산 년월" CssClass="style4"></asp:Label>
            </strong>
        </td><td>
            <asp:TextBox ID="tb_yyyymm" runat="server"></asp:TextBox>
         </td>
         <td>
             <asp:Button ID="btn_retrive" runat="server" Text="조회" 
                 onclick="btn_retrive_Click" />
             <asp:FileUpload ID="FileUpload1" runat="server" />
            </td>
         </tr></table>
    </div>
    <div>
        <asp:Panel ID="Panel1" runat="server" Width="559px">
            <asp:GridView ID="GridView1" runat="server" 
    AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" 
    GridLines="None">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:TemplateField HeaderText="NO"></asp:TemplateField>
                    <asp:TemplateField HeaderText="코스트센터" SortExpression="cost_cd">
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="코스트센터명" SortExpression="cost_cd_nm">
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="코스트그룹" SortExpression="cost_gp_cd">
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="코스트그룹명" SortExpression="cost_gp_nm">
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="수정"></asp:TemplateField>
                    <asp:TemplateField HeaderText="삭제"></asp:TemplateField>
                </Columns>
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
        
        
    </div>
    </form>
</body>
</html>
