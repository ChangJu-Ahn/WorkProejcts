<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="pop_item_cd.aspx.cs" Inherits="ERPAppAddition.ERPAddition.POPUP.pop_item_cd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>품목코드선택</title>
    <script type="text/javascript" src="../pop.js"></script>
    <style type="text/css">
        .style12
        {
            width: 110px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
        }
        .style1
        {
            width: 77px;
        }
        .style2
        {
            width: 160px;
        }
        .style3
        {
            margin-left: 0px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <table style="width:650px;">
            <tr>
            <td class="style12">공장</td>
            <td class="style2">
                <asp:TextBox ID="tb_plant_nm" Width="150px" runat="server" ReadOnly="True"></asp:TextBox></td>
            <td>
                <asp:TextBox ID="tb_plant_cd" runat="server" Visible="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style12"> 품목</td>
                <td class="style2">
                    <asp:TextBox ID="tb_pop_item_cd" runat="server" Width="150px"></asp:TextBox>
                </td>
                <td>
                    <asp:TextBox ID="tb_pop_item_nm" runat="server" CssClass="style3" Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style12">
                    품목계정</td>
                <td class="style2">
                    <asp:DropDownList ID="dl_pop_item_acct" runat="server" 
                        DataSourceID="SqlDataSource1" DataTextField="MINOR_NM" 
                        DataValueField="MINOR_CD" Height="22px" Width="150px">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:erp_db %>" 
                        SelectCommand="SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE (MAJOR_CD = 'p1001')">
                    </asp:SqlDataSource>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
        <asp:Button ID="Button1" runat="server" Text="조회" onclick="Button1_Click" 
            ViewStateMode="Enabled" Width="100px" />
        <asp:Button ID="Button2" runat="server" Text="취소" Width="100px" />
        <div style="overflow: auto; width: 700px; height: 350px">
            <asp:GridView ID="pop_gridview1" runat="server" BorderStyle="Solid" BorderWidth="1px"
                BackColor="White" BorderColor="#E7E7FF" CellPadding="3" HorizontalAlign="Left"
                AutoGenerateColumns="False" AutoGenerateSelectButton="True" 
                ShowHeaderWhenEmpty="True" Font-Size="Small"
                >
                <AlternatingRowStyle BackColor="#F7F7F7" />
                <Columns>
                    <asp:BoundField DataField="item_cd" HeaderText="품목코드">
                        <ItemStyle Width="100px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="item_nm" HeaderText="품목명">
                        <ItemStyle Width="220px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="spec" HeaderText="규격">
                        <ItemStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="basic_unit" HeaderText="품목단위">
                        <ItemStyle Width="60px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="item_acct_nm" HeaderText="품목계정">
                        <ItemStyle Width="100px" />
                    </asp:BoundField>
                </Columns>
                <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
                <PagerSettings PageButtonCount="100" />
                <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
                <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
                <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
                <SortedAscendingCellStyle BackColor="#F4F4FD" />
                <SortedAscendingHeaderStyle BackColor="#5A4C9D" />
                <SortedDescendingCellStyle BackColor="#D8D8F0" />
                <SortedDescendingHeaderStyle BackColor="#3E3277" />
            </asp:GridView>
        </div>
    
    </div>
    </form>
</body>
</html>
