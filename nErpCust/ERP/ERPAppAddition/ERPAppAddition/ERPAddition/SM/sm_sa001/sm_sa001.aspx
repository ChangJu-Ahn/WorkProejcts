<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sa001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sa001.sm_sa001" %>

<%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
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
        .default_font_size
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align:center;
        }
                .tbl_list{border:1px solid #e8e9ea; border-collapse:collapse; background-color:#fff; font-size: 10pt;
            margin-right: 0px;
        }

        .auto-style45 {
            width: 90px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
        }
                            
        .auto-style36 {
            font-size: small;
            width: 126px;
            vertical-align : top;
            height: 22px;
        }
        .auto-style37 {
            font-size: small;
            width: 157px;
            vertical-align : top;
            height: 22px;
        }
        .auto-style38 {
            font-size: small;
            width: 67px;
            vertical-align : top;
            height: 22px;
        }
        .auto-style39 {
            font-size: small;
            width: 169px;
            vertical-align : top;
            height: 22px;
        }
            .auto-style47 {
            font-size : small;
            height: 15px;
        }
                    
        .auto-style33 {
            font-size: small;
            width: 126px;
            vertical-align : top;
            height: 21px;
        }
        .auto-style34 {
            font-size: small;
            width: 157px;
            vertical-align : top;
            height: 21px;
        }
        .auto-style35 {
            font-size: small;
            width: 100px;
            vertical-align : top;
            height: 21px;
        }
        
        .auto-style44 {
            font-size : small;
            width: 169px;
            height: 21px;
        }
        
        .auto-style48 {
            font-size : small;
            height: 22px;
            width: 157px;
        }
        .auto-style49 {
            font-size : small;
            width: 157px;
            height: 21px;
        }
        .auto-style51 {
            font-size : small;
            color : black;
            width: 83px;
            height: 21px;
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
            .auto-style101 {
            width: 1151px;
        }            
            .auto-style102 {
            width: 90px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
            height: 15px;
        }
            .auto-style103 {
            height: 31px;
        }
            </style>

</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager2" runat="server" EnablePageMethods="true">
    </asp:ScriptManager>
    <div>
    <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label6" runat="server" Text="Scrap 반출정보등록" CssClass=title Width="757%"></asp:Label>
        </td></tr></table>        
        
    </div>
    <div>
                <table style="border: thin solid #000080; width: 100%">
                    <tr>
                        <td class="auto-style101">
                            <table>                                
                                <tr>
                                    <td class="auto-style45">
                                        발생일:
                                    </td>
                                    <td class="style1">
                                        <asp:TextBox ID="tb_fr_yyyymmdd" runat="server" MaxLength="8" Width="65px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_fr_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_fr_yyyymmdd">
                                        </cc1:CalendarExtender>
                                        <asp:Label ID="Label1" runat="server" Text="~"></asp:Label>
                                        <asp:TextBox ID="tb_to_yyyymmdd" runat="server" MaxLength="8" Width="65px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_to_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_to_yyyymmdd">
                                        </cc1:CalendarExtender>
                                    </td>
                                    <td></td>
                                    <td>
                                        <asp:CheckBox ID="CheckBox2" runat="server" Text="반출대상" BackColor="#FF9933" Font-Size="Small" OnCheckedChanged="CheckBox1_CheckedChanged" Font-Bold="True" Checked="True"/>
                                        <asp:CheckBox ID="CheckBox1" runat="server" Text="전체조회" AutoPostBack="True" BackColor="#FF9933" Font-Size="Small" OnCheckedChanged="CheckBox1_CheckedChanged" Font-Bold="True"/>
                                    </td>
                                    <td colspan ="2">
                                        <asp:Button ID="btn_mighty_retrieve0" runat="server" OnClick="btn_mighty_retrieve_Click1" Text="조회" Width="80px" />
                                        <asp:Button ID="btn_mighty_save" runat="server" Width="80px" Text="저장" OnClick="btn_mighty_save_Click" />
                                    </td>
                                    <td>
                                        <asp:Label ID="Label5" runat="server" Text="입력일 :" Visible="False"></asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox ID="TXT_INPUTDT" runat="server" style="text-align: center" Width="138px" TextMode="Number" Enabled="False" MaxLength="8" BackColor="#CCCCCC" Visible="False"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style102">
                                        <asp:Label ID="Label2" runat="server" Text="공   장 : "></asp:Label>
                                    </td>
                                    <td class="auto-style47">
                                        <asp:DropDownList ID="SDDL_FAC" runat="server" Width="160px" AutoPostBack="True" OnSelectedIndexChanged="SDDL_FAC_SelectedIndexChanged">
                                        </asp:DropDownList>
                                    </td>                                    
                                    <td class="auto-style102">
                                        <asp:Label ID="Label11" runat="server" Text="Scrap종류:"></asp:Label>
                                    </td>
                                    <td class="auto-style47">
                                        <asp:DropDownList ID="DDL_MAT" runat="server" Height="20px" style="margin-top: 0px" Width="170px" AutoPostBack="True" OnSelectedIndexChanged="DDL_MAT_SelectedIndexChanged">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style102">                                        
                                        <asp:Label ID="Label4" runat="server" Text="발생장비 :"></asp:Label>
                                    </td>
                                    <td class="auto-style47">
                                        <asp:DropDownList ID="DDL_MACH" runat="server" Height="20px" style="margin-top: 0px" Width="105px">
                                        </asp:DropDownList>
                                    </td>
                                    <td>                                        
                                    </td>
                                    <td class="auto-style47">
                                    </td>
                                </tr>
                            </table>
                            <table><tr><td><h4 style="margin-bottom: 6px; height: 14px;">▼<strong>입력 내용 선택</strong></td></h4></tr>
                            </table>
                            
                            <table class="tbl_list" dir="ltr">
                                <tr>
                                    <td class="auto-style36">
                                        반출일(년월일시분) :</td>
                                    <td class="auto-style37">
                                        <asp:TextBox ID="TXT_RDT" runat="server" style="text-align: center" MaxLength="12" TextMode="Number" BackColor="#FFFFCC"></asp:TextBox>
                                    </td>                                    
                                    <td class="auto-style35">
                                        <asp:Label ID="Label10" runat="server" Text="반출처 : "></asp:Label>
                                    </td>
                                    <td class="auto-style48">
                                        <asp:DropDownList ID="DDL_CUST" runat="server" style="margin-top: 0px" Width="147px" BackColor="#FFFFCC">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style38">
                                        &nbsp;</td>
                                    <td class="auto-style39">
                                        &nbsp;</td>                                    
                                </tr>
                                <tr>
                                    <td class="auto-style33">
                                        <asp:Label ID="Label9" runat="server" Text="Scrap수량"></asp:Label>
                                    &nbsp;:
                                    </td>
                                    <td class="auto-style34">
                                        <asp:TextBox ID="TXT_SCRQTY" runat="server" style="text-align: center" Width="84px" Wrap="False" BackColor="#FFFFCC"
                                            onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46)) event.returnValue=false;"></asp:TextBox>
                                        <asp:DropDownList ID="DDL_UNIT" runat="server" Height="20px" style="margin-top: 0px" Width="55px" BackColor="#FFFFCC">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style35">
                                        <asp:Label ID="Label8" runat="server" Text="Scrap반출번호:"></asp:Label>
                                    </td>
                                    <td class="auto-style44">
                                        <asp:TextBox ID="TXT_DOCNO" runat="server" style="text-align: center" Width="137px" BackColor="#FFFFCC"></asp:TextBox>
                                    </td>
                                    <td class="auto-style51">
                                        <asp:Label ID="Label3" runat="server" Text="Au 농도: "></asp:Label>
                                    </td>
                                    <td class="auto-style49" colspan ="">
                                        <asp:TextBox ID="TXT_AUQTY" runat="server" style="text-align: right" Width="84px" Wrap="False" BackColor="#FFFFCC" EnableTheming="True"
                                            onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46)) event.returnValue=false;"></asp:TextBox>
                                        <asp:DropDownList ID="DDL_AUUNIT" runat="server" Height="20px" style="margin-top: 0px" Width="55px" BackColor="#FFFFCC">
                                        </asp:DropDownList>
                                    </td>                                    
                                    <td class="auto-style51">
                                        <asp:Label ID="Label7" runat="server" Text="비    고:"></asp:Label>
                                    </td>
                                    <td class="auto-style49" colspan ="">
                                        <asp:TextBox ID="TXT_RMK" runat="server" style="text-align: center"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan ="2" class="auto-style103" >
                                        <asp:CheckBox ID="CHK_ALL" runat="server" Text=" 확 인 " BackColor="#FF9933" Font-Size="Small" OnCheckedChanged="CHK_ALL_CheckedChanged" Font-Bold="True" AutoPostBack="True"/>
                                        <asp:CheckBox ID="CHK_CNALL" runat="server" Text="취소확인" BackColor="#FF9933" Font-Size="Small" OnCheckedChanged="CHK_CNALL_CheckedChanged" Font-Bold="True" AutoPostBack="True"/>
                                        
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <asp:Button ID="BUT_COMF3" runat="server" OnClick="BUT_COMF3_Click" Text="Excel 내려받기" Width="100px" Height="23px" />
                                        
                                    </td>
                                    
                                    <td colspan ="6" class="auto-style103" >                                        
                                        <asp:Button ID="Button1" runat="server" Text="선택수량 확인" Width="102px" OnClick="Button1_Click" Height="23px" />                                        
                                        <asp:Label ID="Label12" runat="server" Text="선택수량:"></asp:Label>
                                        <asp:TextBox ID="txt_sum" runat="server" Width="100px" style="text-align: right"></asp:TextBox>
                                        <asp:Label ID="Label13" runat="server" Text="취소선택수량:"></asp:Label>
                                        <asp:TextBox ID="txt_c_sum" runat="server" Width="100px" style="text-align: right"></asp:TextBox>
                                        <asp:Label ID="Label14" runat="server" Text="조회수량합계:"></asp:Label>
                                        <asp:TextBox ID="txt_all_sum" runat="server" Width="100px" style="text-align: right"></asp:TextBox>
                                    </td>                                    
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>                
            
    
                <FarPoint:FpSpread ID="FpSpread_new_data" runat="server" BorderColor="Black" BorderStyle="Solid"
                    BorderWidth="1px" Height="480px" Width="100%" CommandBarOnBottom="False" CssClass="default_font_size"
                    ActiveSheetViewIndex="0" DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;" WaitMessage="조회중입니다." HorizontalScrollBarPolicy="Always"
                    VerticalScrollBarPolicy="Always" currentPageIndex="0" ClientAutoCalculation="True"   >
                    <CommandBar BackColor="Control" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight"
                        ButtonShadowColor="ControlDark" UseSheetSkin="False" Visible="False">
                        <Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif" />
                    </CommandBar>
                    <Pager PageCount="1000" Align="Left" Mode="Both" Position="TopBottom" />
                    <HierBar Align="Left" ShowParentRow="False" ShowWholePath="False" />
                    <Sheets>
                        <FarPoint:SheetView DataSourceID="(없음)" DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;29&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;선택&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;취소선택&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;29&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;False&lt;/AutoGenerateColumns&gt;&lt;DataKeyField class=&quot;System.String[]&quot; assembly=&quot;mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089&quot; encoded=&quot;true&quot;&gt;AAEAAAD/////AQAAAAAAAAARAQAAAAAAAAAL&lt;/DataKeyField&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;RowIndexes class=&quot;System.Collections.ArrayList&quot; assembly=&quot;mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089&quot; encoded=&quot;true&quot;&gt;AAEAAAD/////AQAAAAAAAAAEAQAAABxTeXN0ZW0uQ29sbGVjdGlvbnMuQXJyYXlMaXN0AwAAAAZfaXRlbXMFX3NpemUIX3ZlcnNpb24FAAAICAkCAAAAAwAAACICAAAQAgAAAAQAAAAICAAAAAAICAEAAAAICAIAAAAKCw==&lt;/RowIndexes&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;1&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;1&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;2&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;29&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;ActiveSkin class=&quot;FarPoint.Web.Spread.SheetSkin&quot;&gt;&lt;Name&gt;Classic&lt;/Name&gt;&lt;BackColor&gt;Empty&lt;/BackColor&gt;&lt;CellBackColor&gt;Empty&lt;/CellBackColor&gt;&lt;CellForeColor&gt;Empty&lt;/CellForeColor&gt;&lt;CellSpacing&gt;0&lt;/CellSpacing&gt;&lt;GridLines&gt;Both&lt;/GridLines&gt;&lt;GridLineColor&gt;LightGray&lt;/GridLineColor&gt;&lt;HeaderBackColor&gt;Control&lt;/HeaderBackColor&gt;&lt;HeaderForeColor&gt;Empty&lt;/HeaderForeColor&gt;&lt;FlatColumnHeader&gt;False&lt;/FlatColumnHeader&gt;&lt;FooterBackColor&gt;Empty&lt;/FooterBackColor&gt;&lt;FooterForeColor&gt;Empty&lt;/FooterForeColor&gt;&lt;FlatColumnFooter&gt;False&lt;/FlatColumnFooter&gt;&lt;FlatRowHeader&gt;False&lt;/FlatRowHeader&gt;&lt;HeaderFontBold&gt;False&lt;/HeaderFontBold&gt;&lt;FooterFontBold&gt;False&lt;/FooterFontBold&gt;&lt;SelectionBackColor&gt;LightBlue&lt;/SelectionBackColor&gt;&lt;SelectionForeColor&gt;Empty&lt;/SelectionForeColor&gt;&lt;EvenRowBackColor&gt;Empty&lt;/EvenRowBackColor&gt;&lt;OddRowBackColor&gt;Empty&lt;/OddRowBackColor&gt;&lt;ShowColumnHeader&gt;True&lt;/ShowColumnHeader&gt;&lt;ShowColumnFooter&gt;False&lt;/ShowColumnFooter&gt;&lt;ShowRowHeader&gt;True&lt;/ShowRowHeader&gt;&lt;ColumnHeaderBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/ColumnHeaderBackground&gt;&lt;SheetCornerBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/SheetCornerBackground&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;/ActiveSkin&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;29&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;Item index=&quot;0&quot;&gt;&lt;Size&gt;35&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;1&quot;&gt;&lt;Size&gt;35&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;2&quot;&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;3&quot;&gt;&lt;Size&gt;100&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;4&quot;&gt;&lt;Size&gt;70&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;6&quot;&gt;&lt;Size&gt;70&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;9&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;10&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;11&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;12&quot;&gt;&lt;Size&gt;45&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;13&quot;&gt;&lt;Size&gt;150&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;14&quot;&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;15&quot;&gt;&lt;Tag type=&quot;System.String&quot;&gt;seq&lt;/Tag&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;16&quot;&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;17&quot;&gt;&lt;Size&gt;95&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;18&quot;&gt;&lt;Size&gt;70&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;20&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;21&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;22&quot;&gt;&lt;Size&gt;38&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;24&quot;&gt;&lt;Size&gt;50&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;25&quot;&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;26&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;27&quot;&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;28&quot;&gt;&lt;Size&gt;250&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;29&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;29&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;/DefaultStyle&gt;&lt;ColumnStyles&gt;&lt;ColumnStyle Index=&quot;0&quot;&gt;&lt;CellType class=&quot;FarPoint.Web.Spread.CheckBoxCellType&quot; /&gt;&lt;Locked&gt;False&lt;/Locked&gt;&lt;TabStop&gt;True&lt;/TabStop&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;1&quot;&gt;&lt;CellType class=&quot;FarPoint.Web.Spread.CheckBoxCellType&quot; /&gt;&lt;Locked&gt;False&lt;/Locked&gt;&lt;TabStop&gt;True&lt;/TabStop&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;2&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;3&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;4&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;6&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;8&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;9&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;10&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;11&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;12&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;13&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;14&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;15&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;16&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;17&quot;&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;18&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;20&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;21&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;22&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;24&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;25&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;26&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;27&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;28&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;/ColumnStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;29&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;true&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;DataAutoCellTypes&gt;False&lt;/DataAutoCellTypes&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;PageSize&gt;20&lt;/PageSize&gt;&lt;AllowPage&gt;False&lt;/AllowPage&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;ColumnCount&gt;29&lt;/ColumnCount&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;ActiveRow&gt;0&lt;/ActiveRow&gt;&lt;ActiveColumn&gt;19&lt;/ActiveColumn&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;"
                            SheetName="Sheet1" PageSize="20" AutoGenerateColumns="False" AllowVirtualScrollPaging="True" DataAutoCellTypes="False" AllowPage="False" GridLineColor="#D0D7E5" SelectionBackColor="#EAECF5">
                        </FarPoint:SheetView>
                    </Sheets>
                    <TitleInfo BackColor="#E7EFF7" Font-Size="X-Large" ForeColor="" HorizontalAlign="Center"
                        VerticalAlign="NotSet" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                        Font-Strikeout="False" Font-Underline="False">
                    </TitleInfo>
                </FarPoint:FpSpread>

       </div>       
    </form>
</body>
</html>
