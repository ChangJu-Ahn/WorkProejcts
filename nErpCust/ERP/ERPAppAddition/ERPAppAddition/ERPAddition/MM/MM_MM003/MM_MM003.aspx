<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_MM003.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_MM003.MM_MM003" %>
<%@ Register Assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" Namespace="FarPoint.Web.Spread" TagPrefix="FarPoint" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .BasicTb {
            width: auto;
            border: thin double #000080;
        }

        td.tilte {
            background-color: #99CCFF;
            font-weight: bold;
            text-align: center;
            width: 70px;
            white-space: nowrap;
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
            }
         .text-right {
            text-align: right;
        }
        .ui-progressbar {
            position: relative;
        }
        .spread {}
        .TD5
        {
            PADDING-RIGHT: 5px;
            WIDTH: 14%;
            BACKGROUND-COLOR: #e7e5ce;
            TEXT-ALIGN: right
        }
        .TD6
        {
            PADDING-LEFT: 5px;
            WIDTH: 36%;
            BACKGROUND-COLOR: #f5f5f5
        }
        .auto-style2 {
            width: 17px;
        }
    </style>
</head>
<body>
     <script language="javascript" type="text/javascript">

         function btnPrevclick() {
             FpSpread1.Prev();
         }

         function btnNextclick() {
             FpSpread1.Next();
         }

                       
         function fn_GetPartner(itemCD, itemNM) {
             var PopWidth = 635;
             var PopHeight = 520;

             var dbGubun = document.getElementById("hdnDB").value;
             var PopNodeUrl = "../../BP/Pop_Cost.aspx?dbName=" + dbGubun;
             var PopFont = "FONT-FAMILY: '맑은고딕';font-size:15px;";
             var PopParams = new Array(); //별도의 넘길 값은 없으나 형식에 맞추기 위해 배열객체만 선언

             var search = "&search=" + itemCD + "," + itemNM + "";
             PopNodeUrl += search;
             PopNodeUrl = encodeURI(PopNodeUrl);
             var Retval = window.showModalDialog(PopNodeUrl, PopParams, PopFont + "dialogHeight:" + PopHeight + "px;dialogWidth:" + PopWidth + "px;resizable:no;status:no;help:no;scroll:no;location:no");
                             
             return Retval
         }
         function fn_GetItem(columns, sql) {
             var PopWidth = 635;
             var PopHeight = 520;

             var dbGubun = document.getElementById("hdnDB").value;
             var PopNodeUrl = "../../BP/Pop_Item.aspx?dbName=" + dbGubun;
             var PopFont = "FONT-FAMILY: '맑은고딕';font-size:15px;";
             var PopParams = new Array(); //별도의 넘길 값은 없으나 형식에 맞추기 위해 배열객체만 선언


             var columnsList = "&columns=" + columns;
             var sqlQuery = "&sql=" + sql;
             //var search = "&search=" + itemCD + ";모델명," + itemNM + ";비고";


             PopNodeUrl += columnsList + sqlQuery;

             PopNodeUrl = encodeURI(PopNodeUrl);

             var Retval = window.showModalDialog(PopNodeUrl, PopParams, PopFont + "dialogHeight:" + PopHeight + "px;dialogWidth:" + PopWidth + "px;resizable:no;status:no;help:no;scroll:no;location:no");


             return Retval
         }

         function ClickBtnBP() {
             var BP_CD = document.getElementById('<%=txtBPCD.ClientID %>').value;
                var BP_NM = document.getElementById('<%=txtBPNM.ClientID %>').value;


                var Retval = fn_GetPartner(BP_CD, BP_NM)

                if (Retval != null) {

                    var Item_cd = Retval.split(";")[0];
                    var Item_nm = Retval.split(";")[1];

                    document.getElementById('<%=txtBPCD.ClientID %>').value = Item_cd;
                    document.getElementById('<%=txtBPNM.ClientID %>').value = Item_nm;
                }
            }

            function ClickBtnITEM() {

                var plant = document.getElementById('<%=txtPLANT_CD.ClientID %>').value;

                var col = "모델명,비고";
                
                var where = "";
                if (plant != "") {
                    where = " AND AA.PLANT_CD ='" + plant + "'";
                }

                var sql = "SELECT ITEM_CD, ITEM_NM FROM(SELECT AA.ITEM_CD, ITEM_NM FROM B_ITEM_BY_PLANT AA INNER JOIN B_ITEM BB ON AA.ITEM_CD =bb.ITEM_CD WHERE 1=1 " + where + " ) AA";
                var Retval = fn_GetItem(col, sql);

                if (Retval != null) {

                    var Item_cd = Retval.split(";")[0];
                    var Item_nm = Retval.split(";")[1];

                    document.getElementById('<%=txtITEMCD.ClientID %>').value = Item_cd;
                    document.getElementById('<%=txtITEMNM.ClientID %>').value = Item_nm;
                }
            }

         function ClickBtnPLANT() {

             
             var col = "공장,공장명";

             var sql = "SELECT ITEM_CD, ITEM_NM FROM(SELECT PLANT_CD AS ITEM_CD, PLANT_NM AS ITEM_NM FROM B_PLANT WHERE VALID_TO_DT > GETDATE())AA";
             var Retval = fn_GetItem(col, sql);

             if (Retval != null) {

                 var Item_cd = Retval.split(";")[0];
                 var Item_nm = Retval.split(";")[1];

                 document.getElementById('<%=txtPLANT_CD.ClientID %>').value = Item_cd;
                    document.getElementById('<%=txtPLANT_NM.ClientID %>').value = Item_nm;
                }
         }

         function ClickBtnPURGRP() {


             var col = "구매그룹,구매그룹명";

             var sql = " SELECT ITEM_CD, ITEM_NM FROM(SELECT PUR_GRP AS ITEM_CD, PUR_GRP_NM AS ITEM_NM FROM B_PUR_GRP)AA ";
             var Retval = fn_GetItem(col, sql);

             if (Retval != null) {

                 var Item_cd = Retval.split(";")[0];
                 var Item_nm = Retval.split(";")[1];

                 document.getElementById('<%=txtPUR_GRP.ClientID %>').value = Item_cd;
                 document.getElementById('<%=txtPUR_GRP_NM.ClientID %>').value = Item_nm;
             }
         }

        </script>

    <form id="form1" runat="server">
    <div>
        <table style="width: 100%;">
        
            <td class="auto-style2" >

               <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />

            </td>
            <td>
                <asp:Label ID="Label6" runat="server" Text="구매요청마감" CssClass="title"></asp:Label>
            </td>
            <td>
                <asp:HiddenField ID="hdnID" runat="server" />
            </td>
            <td>
                <asp:HiddenField ID="hdnDB" runat="server" />
            </td>
            <td class="text-right" >
                
                <asp:Button ID="btnSearch" runat="server" OnClick="bntSearch_Click" Text="조회" Height="26px" Width="70px" /> 
                <asp:Label ID="Label5" runat="server" Text=" "  Width="3px"></asp:Label>
                <asp:Button ID="btnSave" runat="server" Text="적용" Height="26px" Width="70px" OnClick="btnSave_Click" /> 
            </td>
        </tr>
    </table>  
    </div>
        <div>
            <table style="width: 100%; border:solid; border-color:black" cellSpacing="0" rules="none" >
                <tr>
                    <td class="TD5">
                        구매그룹
                    </td>
                   <td class="TD6">
                       <asp:TextBox ID="txtPUR_GRP" runat="server" Width="122px" ClientIDMode="Static"></asp:TextBox>
                    <asp:Button ID="btnBPU_GRP" runat="server" Text=".." Width="18px" OnClientClick="ClickBtnPURGRP()"/>
                    <asp:TextBox ID="txtPUR_GRP_NM" runat="server" BackColor="Silver" Width="167px" ClientIDMode="Static"></asp:TextBox>
                   </td>
                    <td class ="TD5">공장</td>
                    <td class =" TD6">
                        <asp:TextBox ID="txtPLANT_CD" runat="server" Width="122px" ClientIDMode="Static"></asp:TextBox>
                    <asp:Button ID="btnPLANT" runat="server" Text=".." Width="18px" OnClientClick="ClickBtnPLANT()"/>
                    <asp:TextBox ID="txtPLANT_NM" runat="server" BackColor="Silver" Width="167px" ClientIDMode="Static"></asp:TextBox>
                    </td>
                </tr>
              <tr>
                  <td class ="TD5">
                    공급처</td>
                <td class ="TD6">
                    <asp:TextBox ID="txtBPCD" runat="server" Width="122px" ClientIDMode="Static" ></asp:TextBox>
                    <asp:Button ID="btnBP" runat="server" Text=".." Width="18px" OnClientClick="ClickBtnBP()" />
                    <asp:TextBox ID="txtBPNM" runat="server" BackColor="Silver" Width="167px"></asp:TextBox>
                </td>
                <td class ="TD5">
                    발주예정일</td>
                <td class ="TD6">

                    <asp:TextBox ID="txtFrREQ_DT" runat="server" Width="122px" ></asp:TextBox>
                    <cc1:CalendarExtender ID="yyyymmdd_CalendarExtender0" runat="server" Enabled="True" Format="yyyyMMdd" TargetControlID="txtFrREQ_DT">
                    </cc1:CalendarExtender>
                    <asp:Label ID="lblFT" runat="server" Text=" ~ "></asp:Label>
                    <asp:TextBox ID="txtToREQ_DT" runat="server" Width="122px"></asp:TextBox>
                    <cc1:CalendarExtender ID="yyyymmdd_CalendarExtender1" runat="server" Enabled="True" Format="yyyyMMdd" TargetControlID="txtToREQ_DT">
                    </cc1:CalendarExtender>

                </td>
                
            </tr>
            <tr>
                <td class ="TD5">
                    필요일
                </td>
                <td class ="TD6">
                    <asp:TextBox ID="txtFrDLVY_DT" runat="server" Width="122px"></asp:TextBox>
                    <cc1:CalendarExtender ID="CalendarExtender1" runat="server" Enabled="True" Format="yyyyMMdd" TargetControlID="txtFrDLVY_DT">
                    </cc1:CalendarExtender>
                    <asp:Label ID="Label1" runat="server" Text=" ~ "></asp:Label>
                    <asp:TextBox ID="txtToDLVY_DT" runat="server" Width="122px"></asp:TextBox>
                    <cc1:CalendarExtender ID="CalendarExtender2" runat="server" Enabled="True" Format="yyyyMMdd" TargetControlID="txtToDLVY_DT">
                    </cc1:CalendarExtender>
                </td>
                <td class ="TD5">
                    품목명
                </td>
                <td class ="TD6">
                    <asp:TextBox ID="txtITEMCD" runat="server" Width="122px" ClientIDMode="Static"></asp:TextBox>
                    <asp:Button ID="btnITEM" runat="server" Text=".." Width="18px" OnClientClick="ClickBtnITEM()"/>
                    <asp:TextBox ID="txtITEMNM" runat="server" BackColor="Silver" Width="167px" ClientIDMode="Static"></asp:TextBox>
                </td>
                
            </tr>
            
            </table>
        </div>
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="True">       
    </asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                <ContentTemplate>
      <div>
      </div>
                    <asp:Button ID="btnPrev" runat="server" ClientIDMode="Static" Height="25px" OnClientClick="btnPrevclick()" Text="&lt;&lt;" />
                    <asp:Label ID="Label2" runat="server" Text=" " Width="2px"></asp:Label>
                    <asp:Button ID="btnNext" runat="server" ClientIDMode="Static" Height="25px" OnClientClick="btnNextclick()" Text="&gt;&gt;" />
                    <asp:Label ID="Label3" runat="server" Text=" " Width="5px"></asp:Label>
                    <asp:Label ID="Label4" runat="server" Text=" " Width="3px"></asp:Label>
                    <asp:Button ID="btnExport" runat="server" Height="25px" OnClick="btnExport_Click" Text="Excel" />
                 
                    <div style="height:650px">
                        <FarPoint:FpSpread ID="FpSpread1" runat="server" ActiveSheetViewIndex="0" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" ClientIDMode="AutoID" CssClass="sheet" DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;" Height="98%" Width="100%" ClientAutoCalculation="True">
                            <CommandBar BackColor="Control" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight" ButtonShadowColor="ControlDark" Visible="False">
                                <Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif" />
                            </CommandBar>
                            <Pager Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False" Font-Underline="False" />
                            <HierBar Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False" Font-Underline="False" />
                            <Sheets>
                                <FarPoint:SheetView DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;False&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot; /&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;4&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;4&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot; /&gt;&lt;CellStyles&gt;&lt;CellStyle Row=&quot;0&quot; Column=&quot;0&quot;&gt;&lt;n&gt;&lt;n&gt;바탕&lt;/n&gt;&lt;ns&gt;&lt;n&gt;바탕&lt;/n&gt;&lt;/ns&gt;&lt;s&gt;Smaller&lt;/s&gt;&lt;/n&gt;&lt;GdiCharSet&gt;254&lt;/GdiCharSet&gt;&lt;/CellStyle&gt;&lt;/CellStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;4&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot; /&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;false&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;DataAutoCellTypes&gt;False&lt;/DataAutoCellTypes&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;AllowUserFormulas&gt;False&lt;/AllowUserFormulas&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;SelectionModel class=&quot;FarPoint.Web.Spread.Model.DefaultSheetSelectionModel&quot;&gt;&lt;CellRange Row=&quot;0&quot; Column=&quot;0&quot; RowCount=&quot;1&quot; ColumnCount=&quot;1&quot; /&gt;&lt;/SelectionModel&gt;&lt;ActiveRow&gt;0&lt;/ActiveRow&gt;&lt;ActiveColumn&gt;0&lt;/ActiveColumn&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;" SheetName="Sheet1" AutoGenerateColumns="False" DataAutoCellTypes="False" AllowUserFormulas="False">
                                </FarPoint:SheetView>
                            </Sheets>
                            <TitleInfo BackColor="#E7EFF7" font-bold="False" font-italic="False" font-overline="False" Font-Size="X-Large" font-strikeout="False" font-underline="False" ForeColor="" HorizontalAlign="Center" VerticalAlign="NotSet">
                            </TitleInfo>
                        </FarPoint:FpSpread>
                    </div>
                         </ContentTemplate>
               
                <Triggers>
                    
                    <asp:AsyncPostBackTrigger ControlID="btnPrev" />
                    <asp:AsyncPostBackTrigger ControlID="btnNext" />
                   <asp:PostBackTrigger ControlID="btnExport" />
                </Triggers>
               
                </asp:UpdatePanel>
    </form>
</body>
</html>
