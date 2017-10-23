<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_MA1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_MA1001.MM_MA1001" %>
<%@ Register Assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457"
    Namespace="FarPoint.Web.Spread" TagPrefix="FarPoint" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>MRO 품목 관리(NEPES)</title> <link rel="stylesheet" type="text/css" href="../../../Style.css" />
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
        
        .spread
        {
            width: 120px;
            
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
       
        .style14
        {
            background: url('../../../r-center.png') repeat-x center top; /* Load Center Graphic */
            height: 18px; /* Set height of center graphic */
            margin-right: 25px; /* Set width of right graphic */;
            margin-left: 25px;
            width: 284px;
        }
        .label {
                border-top :0px dashed #Orange;
                border-bottom:0px dashed #Orange;
                background-color:white;
                font-weight: bold;
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


       
        .auto-style1 {
            width: 351px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
        .auto-style3 {
            width: 144px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
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

       
        .auto-style4 {
            background-color: #FFFFCC;
        }

       
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <table>
            <tr>
                <td>
       <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
                </td>
                <td style="width: 100%;">
                     <asp:Label ID="Label1" runat="server" Text="MRO 품목 관리(NEPES)" CssClass="title" Width="100%"></asp:Label>
                </td>
                <td>
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
                    <asp:ListItem Value="A">신규 등록</asp:ListItem>
                    <asp:ListItem Value="B">수정/삭제</asp:ListItem>
                </asp:RadioButtonList>                
                        
            </td>
          
     
    </table> 
    <asp:Panel ID="panel_upload" runat="server" Visible="False" BorderStyle="Groove" 
            BorderColor="White" Width="99%">
            <table style="width: 99%">
                <tr >
                 
                     
            
                           
                     <td class=style12>공장선택</td>
                        <td class="style53">
                         
                            <asp:DropDownList ID="ddl_item_cd" runat="server" style="background-color: #FFFFCC">
                                <asp:ListItem Value="SA-MRO">오창1공장</asp:ListItem>
                                <asp:ListItem Value="SB-MRO">오창2공장</asp:ListItem>
                                <asp:ListItem Value="SC-MRO">12인치공장</asp:ListItem>
                                <asp:ListItem Value="EM-MRO">음성공장</asp:ListItem>
                                <asp:ListItem Value="HM-MRO">본부</asp:ListItem>
                            </asp:DropDownList>
                            </td>

                        
                       
                           
                           
                            <td class=style12>Excel선택</td>
                            <td class="style54">
                            <asp:FileUpload ID="FileUpload1" runat="server" BackColor="#FFFFCC" />
                            &nbsp;<asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" 
                                    Text="Upload" ToolTip="엑셀자료를 가져와 화면에 보여준다." Width="87px" Height="21px" />
                            </td>
                            <td class=style12 style="border-style: none; font-weight: 700;">Sheet선택</td>
                            <td class="style56">
                            <asp:DropDownList ID="ddlSheets" runat="server" AutoPostBack="True" 
                                BackColor="#FFFFCC">
                            </asp:DropDownList>
                            <asp:Button ID="btnSave" runat="server" OnClick="btnSave_Click" 
                                OnClientClick="return confirm('이미 등록되어있는 MRO품목코가 있을 경우 저장되지 않습니다. 진행하시겠습니까?');" Text="Save" 
                                    Width="80px" />
                            <asp:Button ID="btnCancel0" runat="server"  
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
         <asp:Panel ID="Panel_del" runat="server" Visible="False">
    <table style="border: thin solid #000080; width: 100%;">
        <tr>
            <td class="style12" >
                <strong>공장</strong></td>
            <td class="auto-style1">               
                <asp:DropDownList ID="ddl_item_cd2" runat="server" style="background-color: #FFFFCC">
                    <asp:ListItem Value="SA-MRO">오창1공장</asp:ListItem>
                    <asp:ListItem Value="SB-MRO">오창2공장</asp:ListItem>
                    <asp:ListItem Value="SC-MRO">12인치공장</asp:ListItem>
                    <asp:ListItem Value="EM-MRO">음성공장</asp:ListItem>
                    <asp:ListItem Value="HM-MRO">본부</asp:ListItem>
                </asp:DropDownList>
            </td>
            
            <td class="auto-style3" >
                KEP품목코드</td>
            <td class="style1">
                
                <asp:TextBox ID="txt_mro_item_cd" runat="server" CssClass="auto-style4"></asp:TextBox>
                                        <asp:Button ID="btn_search" runat="server"  Text="조회" Width="100px" OnClick="btn_search_Click" />
            </td>
        </tr>
    </table>
              </asp:Panel>
    </div>
    <div>
                               <asp:Panel ID="Panel_menu" runat="server" Visible="False">
                            <table>
                                <tr>
                              
                                        <asp:Button ID="btn_down_line_add" runat="server" Text="추가(1줄)" Width="100px" 
                                            Visible="False" />
                                    </td>
                                   
                                   
                                    <td>
                                        <asp:Button ID="btn_save" runat="server" Text="저장" Width="100px" OnClick="btn_save_Click"  />
                                    </td>
                                    <td>
                                        <asp:Button ID="btn_del0" runat="server" OnClick="btn_del_Click" Text="삭제" Width="100px" />
                                        </td>
                                    
                                </tr>
                            </table>
                        </asp:Panel>
    </div>
    <div>
          <asp:Panel ID="panel_spread" runat="server" Visible="False">
        <FarPoint:FpSpread ID="FpSpread1" runat="server" BorderColor="Black" BorderStyle="Solid"
            BorderWidth="1px" Height="500px" Width="100%" ActiveSheetViewIndex="0" 
            
            
            DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;" 
            CommandBarOnBottom="False" 
            onactiverowchanged="FpSpread1_ActiveRowChanged" 
            onupdatecommand="FpSpread1_UpdateCommand" 
            WaitMessage="잠시만 기다리십시요" CssClass="label" currentPageIndex="0" >
            <CommandBar BackColor="Control" ButtonType="LinkButton" Visible="False">
<Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif"></Background>
            </CommandBar>
            <Pager Font-Bold="False" Font-Italic="False" Font-Overline="False" 
                Font-Strikeout="False" Font-Underline="False" />
            <HierBar Font-Bold="False" Font-Italic="False" Font-Overline="False" 
                Font-Strikeout="False" Font-Underline="False" />
            <Sheets>
                <FarPoint:SheetView SheetName="Sheet1" 
                    DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;0&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;10&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;KEP품목코드&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;대분류&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;중분류&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;품목명&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;4&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;규격&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;5&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;단위&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;6&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;단가&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;7&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;MOQ&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;8&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;조달일&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;9&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;사용여부&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;0&quot; columns=&quot;10&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;DataKeyField class=&quot;System.String[]&quot; assembly=&quot;mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089&quot; encoded=&quot;true&quot;&gt;AAEAAAD/////AQAAAAAAAAARAQAAAAAAAAAL&lt;/DataKeyField&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;Columns&gt;&lt;Column index=&quot;0&quot;&gt;&lt;ColumnInfo&gt;&lt;ColumnName&gt;yyyymm&lt;/ColumnName&gt;&lt;/ColumnInfo&gt;&lt;/Column&gt;&lt;Column index=&quot;2&quot;&gt;&lt;ColumnInfo&gt;&lt;ColumnName&gt;plant_cd&lt;/ColumnName&gt;&lt;/ColumnInfo&gt;&lt;/Column&gt;&lt;Column index=&quot;3&quot;&gt;&lt;ColumnInfo&gt;&lt;ColumnName&gt;acct_cd&lt;/ColumnName&gt;&lt;/ColumnInfo&gt;&lt;/Column&gt;&lt;Column index=&quot;5&quot;&gt;&lt;ColumnInfo&gt;&lt;ColumnName&gt;amt&lt;/ColumnName&gt;&lt;/ColumnInfo&gt;&lt;/Column&gt;&lt;Column index=&quot;6&quot;&gt;&lt;ColumnInfo&gt;&lt;ColumnName&gt;isrt_id&lt;/ColumnName&gt;&lt;/ColumnInfo&gt;&lt;/Column&gt;&lt;/Columns&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;10&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;10&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;Item index=&quot;0&quot;&gt;&lt;Size&gt;99&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;3&quot;&gt;&lt;Size&gt;179&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;4&quot;&gt;&lt;Size&gt;147&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;5&quot;&gt;&lt;Size&gt;77&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;6&quot;&gt;&lt;Size&gt;98&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;7&quot;&gt;&lt;Size&gt;76&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;8&quot;&gt;&lt;Size&gt;77&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;0&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot; /&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;10&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;0&quot; Columns=&quot;10&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot; /&gt;&lt;ColumnStyles&gt;&lt;ColumnStyle Index=&quot;0&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;1&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;2&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;3&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;4&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;5&quot;&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;6&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;7&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;8&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;9&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;/ColumnStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;10&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot; /&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chbg.gif&lt;/BackgroundImageUrl&gt;&lt;SelectedBackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/SelectedBackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;false&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;DataAutoCellTypes&gt;False&lt;/DataAutoCellTypes&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;PageSize&gt;500&lt;/PageSize&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;ColumnCount&gt;10&lt;/ColumnCount&gt;&lt;RowCount&gt;0&lt;/RowCount&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;" 
                    PageSize="500" DataAutoCellTypes="False" DataSourceID="" 
                    EditTemplateColumnCount="2" GridLineColor="#D0D7E5" 
                    GroupBarText="Drag a column to group by that column." 
                    SelectionBackColor="#EAECF5">
                </FarPoint:SheetView>
            </Sheets>

<TitleInfo BackColor="#E7EFF7" ForeColor="" HorizontalAlign="Center" VerticalAlign="NotSet" 
                Font-Size="X-Large" Font-Bold="False" Font-Italic="False" Font-Overline="False" 
                Font-Strikeout="False" Font-Underline="False"></TitleInfo>
        </FarPoint:FpSpread>
        <asp:SqlDataSource ID="SqlDataSource1_am_a2001" runat="server" 
            ConnectionString="<%$ ConnectionStrings:nepes_test1 %>" 
            SelectCommand="SELECT * FROM [am_a2001]"></asp:SqlDataSource>
        </asp:Panel>
    </div>
    </form>
</body>
</html>

