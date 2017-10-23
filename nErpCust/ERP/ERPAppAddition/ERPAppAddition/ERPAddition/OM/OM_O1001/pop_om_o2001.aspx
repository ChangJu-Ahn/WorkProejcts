<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="pop_om_o2001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.OM.OM_O1001.pop_om_o2001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
        function AddData(apply_no) {
            opener.document.getElementById("TxtApplyNo").value = apply_no;
            window.close();
        }
    </script>
</head>
<body>

    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <div>
        <asp:Label ID="Label1" runat="server" Text="출원일"></asp:Label>
        <asp:TextBox ID="txtFromDate" runat="server" CssClass="dt" MaxLength="8"></asp:TextBox>
        <cc1:CalendarExtender ID="txtFromDate_CalendarExtender" runat="server" 
            TargetControlID="txtFromDate" Format="yyyyMMdd">
        </cc1:CalendarExtender>
        &nbsp;~
        <asp:TextBox ID="txtToDate" runat="server" CssClass="dt" MaxLength="8"></asp:TextBox>
        <cc1:CalendarExtender ID="txtToDate_CalendarExtender" runat="server" 
            Enabled="True" TargetControlID="txtToDate" Format="yyyyMMdd">
        </cc1:CalendarExtender>
        <asp:Button ID="btnSearch2" runat="server" onclick="btnSearch2_Click" 
            style="height: 21px; width: 40px" Text="검색" />
    </div>
    <div>
    
        <asp:GridView ID="SerGridView" runat="server" 
            AllowSorting="True" TextMode="MultiLine" AutoGenerateColumns="False" BackColor="White" 
            BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="4" 
            ForeColor="Black" GridLines="Horizontal" 
            onselectedindexchanged="SerGridView_SelectedIndexChanged1">
            <Columns>
                <asp:BoundField DataField="apply_no" HeaderText="출원번호" />
                <asp:BoundField DataField="invent_kr_nm" HeaderText="명칭(국문)" />
                <asp:BoundField DataField="apply_dt" HeaderText="출원일" />
                <asp:TemplateField HeaderText="선택">
                <ItemTemplate>
                <input id="btnSearch3" type="Button" value="선택" onclick="AddData('<%#DataBinder.Eval(Container.DataItem, "apply_no")%>')" />
                </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#CCCC99" ForeColor="Black" />
            <HeaderStyle BackColor="#333333" Font-Bold="True" ForeColor="White" />
            <PagerSettings Visible="False" />
            <PagerStyle BackColor="White" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CC3333" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F7F7F7" />
            <SortedAscendingHeaderStyle BackColor="#4B4B4B" />
            <SortedDescendingCellStyle BackColor="#E5E5E5" />
            <SortedDescendingHeaderStyle BackColor="#242121" />
        </asp:GridView>
        <br/>
    </div>
    </form>
</body>
</html>
