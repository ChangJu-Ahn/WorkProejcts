<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sb001_A02.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sb001.sm_sb001_A02"  MaintainScrollPositionOnPostback="true" %>

<%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<script language="javascript">
    function Exit() {
        if (self.screenTop > 9000) {
            //alert('닫힘');
            // 브라우저 닫힘
        } else {
            if (document.readyState == "complete") {
                //alert('닫기버튼');
                opener.document.getElementById('btn_SEARCH').click();
                // 새로고침
            } else if (document.readyState == "loading") {
                //alert('안에버튼');
                // 다른 사이트로 이동
            }
        }
    }
</script>

<script language="javascript" event="onunload" for="window">
    Exit();
</script>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .auto-style1 {
            width: 196px;
        }
        .auto-style2 {
            height: 25px;
        }
        .auto-style3 {
            width: 196px;
            height: 25px;
        }
        .auto-style4 {
            width: 81px;
        }
        .auto-style5 {
            width: 120px;
        }
        .auto-style6 {
            width: 95px;
        }
        .auto-style7 {
            width: 90px;
        }

        input {
                  border: 1px solid #bcbcbc;
                  border-radius: 0px;
                  -webkit-appearance: none;
                  height: 21px;
                  -webkit-box-sizing: border-box;
                  -moz-box-sizing: border-box;
                  box-sizing: border-box;
                }   
    </style>
</head>
<body>  
            <form id="form1" runat="server">
                <asp:Panel ID="Panel1" runat="server" ScrollBars="Vertical" Height="600">
    <div>
            <table style="width:820px; border: thin solid #000080; ">
                <tr>
                    <td colspan ="6" style="text-align: right"><asp:Button ID="btn_exit" runat="server" Text="닫    기" OnClick="btn_exit_Click" Width="120px" Visible="False"/></td>
                </tr>
            </tr>
            <tr>
                <td>문서번호&nbsp;:</td>
                <td>
                    <asp:TextBox ID="DOC_NO" runat="server" ReadOnly="True"></asp:TextBox>
                </td>
                <td>제&nbsp;&nbsp;&nbsp;&nbsp;목&nbsp;:</td>
                <td colspan="3">
                    <asp:TextBox ID="SUBJ" runat="server" ReadOnly="True" Width="460px"></asp:TextBox>
                </td>
                <tr>
                    <td>반출일자&nbsp;:</td>
                    <td>
                        <asp:TextBox ID="OUT_DT_ITEM" runat="server" ReadOnly="True"></asp:TextBox>
                    </td>
                    <td>반출상태&nbsp;:</td>
                    <td>
                        <asp:TextBox ID="STS" runat="server" ReadOnly="True"></asp:TextBox>
                    </td>
                    <td>반입상태&nbsp;:</td>
                    <td class="auto-style1">
                        <asp:TextBox ID="FROM_FD_YN" runat="server" ReadOnly="True"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style2">반&nbsp;출&nbsp;자&nbsp;:</td>
                    <td class="auto-style2">
                        <asp:TextBox ID="USR_NM" runat="server" ReadOnly="True"></asp:TextBox>
                    </td>
                    <td class="auto-style2">직&nbsp;&nbsp;&nbsp;&nbsp;급&nbsp;:</td>
                    <td class="auto-style2">
                        <asp:TextBox ID="USR_DUTY" runat="server" ReadOnly="True"></asp:TextBox>
                    </td>
                    <td class="auto-style2">소&nbsp;&nbsp;&nbsp;&nbsp;속&nbsp;:</td>
                    <td class="auto-style3">
                        <asp:TextBox ID="APPR_DEPT" runat="server" ReadOnly="True"></asp:TextBox>
                    </td>
                </tr>
            </tr>
        </table>
    </div>
        <div>
            <table>
                <tr>
                    <td>
                        반출내역&nbsp;:&nbsp; </td>
                    <td>                        
                        <asp:Button ID="btn_chkYES" runat="server" Text="확  인" OnClick="btn_chkYES_Click" Width="120px" />                        
                        </td>
                    <td>                        
                        <asp:Button ID="btn_chkCNL" runat="server" Text="취  소" OnClick="btn_chkCNL_Click" Width="120px" />
                        </td>
                        <td>&nbsp; &nbsp;확인일시&nbsp;:&nbsp;</td>
                        <td>
                            <asp:TextBox ID="OUT_DT" runat="server" ReadOnly="True"></asp:TextBox>
                        </td>
                        <td>
                            <asp:TextBox ID="OUT_COF_USER" runat="server" ReadOnly="True" Width="120px"></asp:TextBox>
                        </td>                                       
                </tr>                
            </table>
             <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width ="820px" CellPadding="4" Font-Bold="False" Font-Italic="False" Font-Overline="False" ForeColor="#333333" GridLines="None">
                 <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                 <Columns>
                     <asp:BoundField DataField="GOODS_NM" HeaderStyle-BackColor="#99CCFF" HeaderText="품    명" ItemStyle-HorizontalAlign="Center">
                     <HeaderStyle BackColor="#99CCFF" />
                     <ItemStyle HorizontalAlign="Center" />
                     </asp:BoundField>
                     <asp:BoundField DataField="AMOUNT" HeaderStyle-BackColor="#99CCFF" HeaderText="수    량" ItemStyle-HorizontalAlign="Center">
                     <HeaderStyle BackColor="#99CCFF" />
                     <ItemStyle HorizontalAlign="Center" />
                     </asp:BoundField>
                     <asp:BoundField DataField="UNIT" HeaderStyle-BackColor="#99CCFF" HeaderText="단    위" ItemStyle-HorizontalAlign="Center">
                     <HeaderStyle BackColor="#99CCFF" />
                     <ItemStyle HorizontalAlign="Center" />
                     </asp:BoundField>
                     <asp:BoundField DataField="RECARRY_DATE" HeaderStyle-BackColor="#99CCFF" HeaderText="반입예정일" ItemStyle-HorizontalAlign="Center">
                     <HeaderStyle BackColor="#99CCFF" />
                     <ItemStyle HorizontalAlign="Center" />
                     </asp:BoundField>
                     <asp:BoundField DataField="PROCESS_INSTANCE_OID" HeaderStyle-BackColor="#99CCFF" HeaderText="OID" ItemStyle-HorizontalAlign="Center" Visible="False">
                     <HeaderStyle BackColor="#99CCFF" />
                     <ItemStyle HorizontalAlign="Center" />
                     </asp:BoundField>
                     <asp:BoundField DataField="GOODS_INDEX" HeaderStyle-BackColor="#99CCFF" HeaderText="GINDEX" ItemStyle-HorizontalAlign="Center" Visible="False">
                     <HeaderStyle BackColor="#99CCFF"/>
                     <ItemStyle HorizontalAlign="Center" />
                     </asp:BoundField>
                 </Columns>
                 <EditRowStyle BackColor="#999999" />
                 <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                 <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                 <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                 <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                 <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                 <SortedAscendingCellStyle BackColor="#E9E7E2" />
                 <SortedAscendingHeaderStyle BackColor="#506C8C" />
                 <SortedDescendingCellStyle BackColor="#FFFDF8" />
                 <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
            </asp:GridView>
        </div>
        <div>            
            <table>                
                <tr>
                    <td>&nbsp;</td>
                </tr>
            </table>
            <table style="width:820px; border: thin solid #000080; " id="table2" runat="server">
            <tr>
                <td class="auto-style4">반입수량:
                </td>
                <td class="auto-style5">
                    <asp:TextBox ID="crtOutAMT" runat="server" Width="100" TextMode="Number"></asp:TextBox>
                </td>
                <td>
                    <asp:Button ID="btnAdd" runat="server" Text="▲" Width="20px" OnClick="btnAdd_Click" />
                    <asp:Button ID="btnMin" runat="server" Text="▼" Width="20px" OnClick="btnMin_Click" />
                </td>
                <td class="auto-style7">                        
                        <asp:Button ID="btn_cryYES" runat="server" Text="확  인" OnClick="btn_cryYES_Click" Width="88px" />
                    </td>
                    <td class="auto-style7">                        
                        <asp:Button ID="btn_cryCNL" runat="server" Text="취  소" OnClick="btn_cryCNL_Click" Width="88px" />
                    </td>
                <td class="auto-style7">                        
                        <asp:Button ID="btn_cryAllYes" runat="server" Text="전체확인" OnClick="btn_cryAllYes_Click" Width="90px"
                            OnClientClick="return confirm('반입품목 전체를 확정 하시겠습니까?');"/>                    
                    </td>
                <td class="auto-style6">반입품명&nbsp;:&nbsp;</td>
                <td class="auto-style5">
                    <asp:TextBox ID="txt_item" runat="server" Width="215px" ReadOnly="True" style="margin-left: 0px"></asp:TextBox>
                    <asp:Label ID="labIndex" runat="server" Visible="False"></asp:Label>
                </td>
            </tr>
        </table>
            <div style="overflow:scroll; width:820px;  padding:0; background-color:white;" id="pnl2" runat="server">
                <FarPoint:FpSpread ID="FpSpread1" runat="server" BorderColor="Black" BorderStyle="Solid"
            BorderWidth="1px" Height="400px" CssClass="default_font_size" Width ="1600px"
                ActiveSheetViewIndex="0" HorizontalScrollBarPolicy="Always" 
                VerticalScrollBarPolicy="Always"
                 currentPageIndex="0" DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;" EnableClientScript="False" CommandBarOnBottom="False" OnCellClick="FpSpread1_CellClick" ClientAutoCalculation="True">
            <Tab TabControlPolicy="Never" />
            <CommandBar BackColor="Control" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight"
                ButtonShadowColor="ControlDark" Background-Enable="False" Visible="False">
                <Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif" />
            </CommandBar>
            <Pager PageCount="1000" Align="Left" Mode="Both" Position="TopBottom" />
            <HierBar Font-Bold="False" Font-Italic="False" Font-Overline="False" 
                Font-Strikeout="False" Font-Underline="False" />
            <Sheets>
                <FarPoint:SheetView DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;34&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;34&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;False&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;34&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;ActiveSkin class=&quot;FarPoint.Web.Spread.SheetSkin&quot;&gt;&lt;Name&gt;aaa&lt;/Name&gt;&lt;BackColor&gt;Empty&lt;/BackColor&gt;&lt;CellBackColor&gt;Empty&lt;/CellBackColor&gt;&lt;CellForeColor&gt;Empty&lt;/CellForeColor&gt;&lt;CellSpacing&gt;0&lt;/CellSpacing&gt;&lt;GridLines&gt;Both&lt;/GridLines&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;HeaderBackColor&gt;Empty&lt;/HeaderBackColor&gt;&lt;HeaderForeColor&gt;Empty&lt;/HeaderForeColor&gt;&lt;FlatColumnHeader&gt;False&lt;/FlatColumnHeader&gt;&lt;FooterBackColor&gt;Empty&lt;/FooterBackColor&gt;&lt;FooterForeColor&gt;Empty&lt;/FooterForeColor&gt;&lt;FlatColumnFooter&gt;False&lt;/FlatColumnFooter&gt;&lt;FlatRowHeader&gt;False&lt;/FlatRowHeader&gt;&lt;HeaderFontBold&gt;False&lt;/HeaderFontBold&gt;&lt;FooterFontBold&gt;False&lt;/FooterFontBold&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionForeColor&gt;Empty&lt;/SelectionForeColor&gt;&lt;EvenRowBackColor&gt;Empty&lt;/EvenRowBackColor&gt;&lt;OddRowBackColor&gt;Empty&lt;/OddRowBackColor&gt;&lt;ShowColumnHeader&gt;True&lt;/ShowColumnHeader&gt;&lt;ShowColumnFooter&gt;False&lt;/ShowColumnFooter&gt;&lt;ShowRowHeader&gt;True&lt;/ShowRowHeader&gt;&lt;ColumnHeaderBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/ColumnHeaderBackground&gt;&lt;SheetCornerBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/SheetCornerBackground&gt;&lt;HeaderGrayAreaColor&gt;#7999c2&lt;/HeaderGrayAreaColor&gt;&lt;/ActiveSkin&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;80&quot; orientation=&quot;Horizontal&quot; count=&quot;34&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;80&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;0&quot;&gt;&lt;Size&gt;220&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;1&quot;&gt;&lt;Size&gt;85&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;2&quot;&gt;&lt;Size&gt;78&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;4&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;5&quot;&gt;&lt;Size&gt;65&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;6&quot;&gt;&lt;Size&gt;30&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;8&quot;&gt;&lt;Size&gt;180&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;9&quot;&gt;&lt;Size&gt;250&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;10&quot;&gt;&lt;Size&gt;55&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;11&quot;&gt;&lt;Size&gt;30&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;12&quot;&gt;&lt;Size&gt;150&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;13&quot;&gt;&lt;Size&gt;180&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;14&quot;&gt;&lt;Size&gt;100&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;15&quot;&gt;&lt;Size&gt;30&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;17&quot;&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;18&quot;&gt;&lt;Size&gt;150&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;19&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;20&quot;&gt;&lt;Size&gt;180&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;21&quot;&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;22&quot;&gt;&lt;Size&gt;70&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;23&quot;&gt;&lt;Size&gt;100&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;24&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;25&quot;&gt;&lt;Size&gt;90&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;26&quot;&gt;&lt;Size&gt;100&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;27&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;28&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;150&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;29&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;30&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;31&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;32&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;33&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot; /&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;34&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;34&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot; /&gt;&lt;ColumnStyles&gt;&lt;ColumnStyle Index=&quot;0&quot;&gt;&lt;Font&gt;&lt;Name&gt;Arial&lt;/Name&gt;&lt;Names&gt;&lt;Name&gt;Arial&lt;/Name&gt;&lt;/Names&gt;&lt;Size&gt;10pt&lt;/Size&gt;&lt;Bold&gt;False&lt;/Bold&gt;&lt;Italic&gt;False&lt;/Italic&gt;&lt;Overline&gt;False&lt;/Overline&gt;&lt;Strikeout&gt;False&lt;/Strikeout&gt;&lt;Underline&gt;False&lt;/Underline&gt;&lt;/Font&gt;&lt;GdiCharSet&gt;254&lt;/GdiCharSet&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;1&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;2&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;4&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;5&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;6&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;8&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;9&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;10&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;12&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;13&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;14&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;/ColumnStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;34&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot; /&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;false&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;AutoPostBack&gt;True&lt;/AutoPostBack&gt;&lt;DataAutoCellTypes&gt;False&lt;/DataAutoCellTypes&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;PageSize&gt;20&lt;/PageSize&gt;&lt;OperationMode&gt;RowMode&lt;/OperationMode&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;ColumnCount&gt;34&lt;/ColumnCount&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;DefaultColumnWidth&gt;80&lt;/DefaultColumnWidth&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;Value type=&quot;System.String&quot; whitespace=&quot;&quot; /&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;" 
                    PageSize="20" SheetName="Sheet1" AutoPostBack="True" OperationMode="RowMode" DataSourceID="(없음)" DefaultColumnWidth="80">
                </FarPoint:SheetView>
            </Sheets>
            <TitleInfo BackColor="#E7EFF7" Font-Size="X-Large" ForeColor="" 
                HorizontalAlign="Center" VerticalAlign="NotSet" Font-Bold="False" 
                Font-Italic="False" Font-Overline="False" Font-Strikeout="False" 
                Font-Underline="False" Text="">
            </TitleInfo>
        </FarPoint:FpSpread>  
                </div>
            <div>
                <table cellpadding="5" cellspacing="0" border="1" align="left" style="border-collapse:collapse; border:1px gray solid;">
                    <thead>
                        <tr>
                            <td width="263px" style="border: 1px gray solid; text-align: center;">이미지1</td>
                            <td width="263px" style="border:1px gray solid; text-align: center;">이미지2</td>
                            <td  width="263px" style="border:1px gray solid; text-align: center;">이미지3</td>
                        </tr>
                    </thead>
                    <thead>
                        <tr>
                            <td width="263px" style="border: 1px gray solid; text-align: center;">                                
                                <asp:ImageButton ID="ImageButton1" runat="server" OnClick="ImageButton1_Click" />
                            </td>
                            <td width="263px" style="border:1px gray solid; text-align: center;">
                                <asp:ImageButton ID="ImageButton2" runat="server" OnClick="ImageButton1_Click" />
                            </td>
                            <td  width="263px" style="border:1px gray solid; text-align: center;">
                                <asp:ImageButton ID="ImageButton3" runat="server" OnClick="ImageButton1_Click" />
                            </td>
                        </tr>
                    </thead>
                </table>
            </div>
        </div>  
        </asp:Panel>
        </form>
</html>
