<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sa003.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sa001.sm_sa003" %>

<%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .auto-style99 {
            width: 150x;            
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
        }

        
        .auto-style45 {
            width: 65px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
        }
            .style1
        {
            
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            table-layout:fixed;
        }
                .tbl_list{border:1px solid #e8e9ea; border-collapse:collapse; background-color:#fff; font-size: 10pt;
            margin-right: 0px;
        }

        .auto-style38 {
            font-size: small;
            width: 60px;
            vertical-align : top;
            height: 23px;
        }
        .auto-style39 {
            font-size: small;
            width: 70px;
            vertical-align : top;
            height: 23px;
        }
            .auto-style47 {
            font-size : small;
            height: 23px;
            width: 100px;
        }
                    
        .auto-style33 {
            font-size: small;
            width: 70px;
            vertical-align : top;
            height: 23px;
        }
        .auto-style34 {
            font-size: small;
            width: 140px;
            vertical-align : top;
            height: 23px;
        }
                
        .auto-style30 {
            font-size : small;
            color : black;
            width: 82px;
            height: 23px;
            text-align: center;
        }
        .auto-style43 {
            font-size : small;
            width: 150px;
            height: 23px;
        }
        
        .default_font_size
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align:center;
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
                
        .auto-style49 {
            font-size : small;
            width: 150px;
            height: 23px;
        }
                
        .auto-style50 {
            width: 160px;
            height: 25px;
        }
                
        .auto-style51 {
            height: 25px;
        }
                
        .auto-style100 {
            font-size : small;
            color : black;
            width: 81px;
            height: 23px;
        }
                        
        .auto-style104 {
            width: 154px;
            height: 25px;
        }
        .auto-style106 {
            font-size : small;
            color : black;
            width: 120px;
            height: 23px;
        }
        .auto-style107 {
            font-size : small;
            color : black;
            width: 52px;
            height: 23px;
        }
                
        .auto-style108 {
            font-size : small;
            color : black;
            width: 55px;
            height: 23px;
        }
        .auto-style110 {
            font-size : small;
            color : black;
            width: 50px;
            height: 23px;
        }
                
        .auto-style111 {
            width: 156px;
            height: 25px;
        }
                
        .auto-style112 {
            font-size: small;
            font-weight: bold;
        }
                
        .auto-style113 {
            width: 1094px;
        }
        .auto-style114 {
            width: 80px;
            height: 23px;
            text-align: right;
        }
        .auto-style119 {
            width: 250px;
            height: 23px;
        }
                
        .auto-style115 {
            font-size : small;
            color : black;
            width: 50px;
            height: 23px;
        }
                
        .auto-style120 {
            font-size : small;
            width: 148px;
            height: 23px;
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
        <td class="auto-style4"><asp:Label ID="Label6" runat="server" Text="사급자재 입고 등록 및 지급, 발주 등록" CssClass=title Width="757%"></asp:Label>
        </td></tr></table>
    </div>    
        <div>
             <table style="width: 1100px">
                    <tr>
                        <td class="auto-style113">
                            <table>
                                <tr>
                                    <td class="auto-style45">
                                        <strong>반 출 일</strong>
                                    </td>
                                    <td class="style1">
                                        <asp:TextBox ID="tb_fr_yyyymmdd" runat="server" MaxLength="8" Width="68px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_fr_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_fr_yyyymmdd">
                                        </cc1:CalendarExtender>
                                        <asp:Label ID="Label1" runat="server" Text="~"></asp:Label>
                                        <asp:TextBox ID="tb_to_yyyymmdd" runat="server" MaxLength="8" Width="68px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_to_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_to_yyyymmdd">
                                        </cc1:CalendarExtender>
                                    </td>
                                    <td>
                                        <asp:RadioButton ID="rdo01" runat="server" Checked="True" CssClass="auto-style112" GroupName="rdo01" Text="입고대상" />
                                        <asp:RadioButton ID="rdo02" runat="server" CssClass="auto-style112" GroupName="rdo01" Text="지급대상" />
                                        <asp:RadioButton ID="rdo03" runat="server" CssClass="auto-style112" GroupName="rdo01" Text="전체조회" />
                                    </td>
                                    <td>
                                        <asp:Button ID="btn_mighty_retrieve0" runat="server" OnClick="btn_mighty_retrieve0_Click" Text="조회" Width="100px" />
                                        <asp:Button ID="btn_mighty_save" runat="server" Width="100px" Text="저장" OnClick="btn_mighty_save_Click" />
                                        <asp:Button ID="btn_mighty_excel" runat="server" Width="100px" Text="Excel 내려받기" OnClick="btn_mighty_excel_Click" />
                                    </td>
                                    <td>
                                        <strong>&nbsp&nbsp&nbsp Lot No.</strong>
                                    </td>
                                    <td>
                                        <asp:TextBox ID="LOTNOBOX" runat="server" style="text-align: center" Enabled="False" Width="200px"></asp:TextBox>
                                    </td>
                                </tr>
                            </table>
                           <table><tr><td class="auto-style99">▼사급자재 입고</td></tr></table>                            
                            <table style="border: thin solid #000080; width: 100%">
                                <tr>                                    
                                    <td class="auto-style38">
                                        <asp:Label ID="Label4" runat="server" Text="입 고 일:"></asp:Label>
                                    </td>
                                    <td class="auto-style39">
                                        <asp:TextBox ID="TXT_INDT" runat="server" style="text-align: center" MaxLength="8" TextMode="Number" AutoPostBack="True" BackColor="#FFFFCC" Width="118px"></asp:TextBox>
                                    </td>
                                    <td class="auto-style110">
                                        <asp:Label ID="Label10" runat="server" Text="수 량:"></asp:Label>
                                    </td>
                                    <td class="auto-style120">
                                        <asp:TextBox ID="TXT_INQTY" runat="server" style="text-align: right" Width="75px" Wrap="False" BackColor="#FFFFCC" EnableTheming="False"
                                            onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46)) event.returnValue=false;"></asp:TextBox>
                                        <asp:DropDownList ID="DDL_INUNIT" runat="server" Height="20px" style="margin-top: 0px" Width="55px" BackColor="#FFFFCC">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style115">
                                        <asp:Label ID="Label2" runat="server" Text="입고자:"></asp:Label>
                                    </td>
                                    <td class="auto-style47">
                                        <asp:DropDownList ID="DDL_INCUST" runat="server" Height="20px" style="margin-top: 0px" Width="147px" BackColor="#FFFFCC">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style114">
                                        ※<asp:Label ID="Label14" runat="server" Text="Au농도:" style="font-size: small"></asp:Label></td>
                                    <td class="auto-style119" >

                                        <asp:TextBox ID="TXT_AUQTY" runat="server" style="text-align: right" Width="84px" Wrap="False" BackColor="#FFCCCC" EnableTheming="False"
                                            onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46)) event.returnValue=false;"></asp:TextBox>
                                        <asp:DropDownList ID="DDL_AUUNIT" runat="server" Height="20px" style="margin-top: 0px" Width="55px" BackColor="#FFCCCC">
                                        </asp:DropDownList>

                                    </td>                                    
                                </tr>
                                </table>
                            
                            <table><tr><td class="auto-style99">▼사급자재 지급</td></tr></table>
                            <table style="border: thin solid #000080; width: 100%">
                                <tr>
                                    <td class="auto-style33">
                                        <asp:Label ID="Label7" runat="server" Text="지 급 일:"></asp:Label>
                                    </td>
                                    <td class="auto-style34">
                                        <asp:TextBox ID="TXT_SENDDT" runat="server" style="text-align: center" Width="116px" TextMode="Number" BackColor="#CCFFCC" MaxLength="8" AutoPostBack="True" OnTextChanged="TXT_SENDDT_TextChanged"></asp:TextBox>
                                    </td>                                    
                                    
                                    <td class="auto-style108">
                                        <asp:Label ID="Label12" runat="server" Text="지 급 처:"></asp:Label>
                                    </td>
                                    <td class="auto-style49">
                                        <asp:DropDownList ID="DDL_SENDCUST" runat="server" Height="20px" style="margin-top: 0px" Width="147px" BackColor="#CCFFCC">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style30">
                                        <asp:Label ID="Label8" runat="server" Text="지 급 량:" style="text-align: right"></asp:Label>
                                    </td>
                                    <td class="auto-style43">
                                        <asp:TextBox ID="TXT_SENDQTY" runat="server" style="text-align: right" Width="93px" BackColor="#CCFFCC" AutoPostBack="True" OnTextChanged="TXT_SENDQTY_TextChanged" EnableTheming="False"
                                            onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46)) event.returnValue=false;"></asp:TextBox>
                                        <asp:TextBox ID="TXT_SENDQTYUNIT" runat="server" style="text-align: center" Width="33px" TextMode="Number" Wrap="False" BackColor="#CCFFCC" Enabled="False"></asp:TextBox>
                                        </td> 
                                     <td class="auto-style106">
                                        <asp:Label ID="Label11" runat="server" Text="사급자재 반출번호:"></asp:Label>
                                    </td>
                                    <td class="auto-style34">
                                        <asp:TextBox ID="TXT_SENDRMK" runat="server" style="text-align: center" Width="138px" BackColor="#CCFFCC" MaxLength="8" AutoPostBack="True" OnTextChanged="TXT_SENDDT_TextChanged"></asp:TextBox>
                                    </td>
                                    <td></td>
                                </tr>
                            </table>
                            <table><tr><td class="auto-style99">▼원자재 임가공</td></tr></table>
                            <table style="border: thin solid #000080; width: 100%">
                                <tr>
                                    <td class="auto-style100">
                                        <asp:Label ID="Label3" runat="server" Text="원자재 종류:"></asp:Label>
                                    </td>
                                    <td class="auto-style50">
                                        <asp:DropDownList ID="DDL_SENDMAT" runat="server" Height="20px" style="margin-top: 0px" Width="147px" BackColor="#CCFFCC">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style107">
                                        <asp:Label ID="Label5" runat="server" Text="수     율:"></asp:Label>
                                    </td>
                                    <td class="auto-style111">
                                        <asp:TextBox ID="TXT_SENDYIELD" runat="server" style="text-align: right" BackColor="#CCFFCC" AutoPostBack="True" OnTextChanged="TXT_SENDYIELD_TextChanged" Width="100px"></asp:TextBox>
                                        [%]</td>
                                    <td class="auto-style30">
                                        <asp:Label ID="Label9" runat="server" Text="임가공사급량:"></asp:Label>
                                    </td>
                                    <td class="auto-style104">
                                        <asp:TextBox ID="TXT_SENDOUTQTY" runat="server" style="text-align: right" Width="84px" Wrap="False" BackColor="#CCFFCC" EnableTheming="False"
                                            onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46)) event.returnValue=false;"></asp:TextBox>
                                        <asp:TextBox ID="TXT_SENDOUTUNIT" runat="server" style="text-align: center" Width="46px" TextMode="Number" Wrap="False" BackColor="#CCFFCC" Enabled="False"></asp:TextBox>
                                        </td>
                                    <td class="auto-style106">
                                        <asp:Label ID="Label13" runat="server" Text="임가공 원자재량:"></asp:Label>
                                    </td>
                                    <td class="auto-style51">
                                        <asp:TextBox ID="TXT_MATE_QTY" runat="server" style="text-align: center" Width="138px" BackColor="#CCFFCC"
                                            onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46)) event.returnValue=false;"></asp:TextBox>
                                        </td>
                                    <td></td>
                                </tr>
                                </table>
                        </td>
                    </tr>
                </table>                
        </div>
        <div style="overflow:scroll; width:1100px;  padding:0; background-color:white;">
        <FarPoint:FpSpread ID="FpSpread1" runat="server" BorderColor="Black" BorderStyle="Solid"
            BorderWidth="1px" Height="400px" CssClass="default_font_size" Width ="2800px"
                ActiveSheetViewIndex="0" HorizontalScrollBarPolicy="Always" 
                VerticalScrollBarPolicy="Always"
                 currentPageIndex="0" OnCellClick="FpSpread1_CellClick" DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;" EnableClientScript="False" CommandBarOnBottom="False">
            <Tab TabControlPolicy="Never" />
            <CommandBar BackColor="Control" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight"
                ButtonShadowColor="ControlDark" Background-Enable="False" Visible="False">
                <Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif" />
            </CommandBar>
            <Pager PageCount="1000" Align="Left" Mode="Both" Position="TopBottom" />
            <HierBar Font-Bold="False" Font-Italic="False" Font-Overline="False" 
                Font-Strikeout="False" Font-Underline="False" />
            <Sheets>
                <FarPoint:SheetView DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;34&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;34&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;False&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;34&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;ActiveSkin class=&quot;FarPoint.Web.Spread.SheetSkin&quot;&gt;&lt;Name&gt;Classic&lt;/Name&gt;&lt;BackColor&gt;Empty&lt;/BackColor&gt;&lt;CellBackColor&gt;Empty&lt;/CellBackColor&gt;&lt;CellForeColor&gt;Empty&lt;/CellForeColor&gt;&lt;CellSpacing&gt;0&lt;/CellSpacing&gt;&lt;GridLines&gt;Both&lt;/GridLines&gt;&lt;GridLineColor&gt;LightGray&lt;/GridLineColor&gt;&lt;HeaderBackColor&gt;Control&lt;/HeaderBackColor&gt;&lt;HeaderForeColor&gt;Empty&lt;/HeaderForeColor&gt;&lt;FlatColumnHeader&gt;False&lt;/FlatColumnHeader&gt;&lt;FooterBackColor&gt;Empty&lt;/FooterBackColor&gt;&lt;FooterForeColor&gt;Empty&lt;/FooterForeColor&gt;&lt;FlatColumnFooter&gt;False&lt;/FlatColumnFooter&gt;&lt;FlatRowHeader&gt;False&lt;/FlatRowHeader&gt;&lt;HeaderFontBold&gt;False&lt;/HeaderFontBold&gt;&lt;FooterFontBold&gt;False&lt;/FooterFontBold&gt;&lt;SelectionBackColor&gt;LightBlue&lt;/SelectionBackColor&gt;&lt;SelectionForeColor&gt;Empty&lt;/SelectionForeColor&gt;&lt;EvenRowBackColor&gt;Empty&lt;/EvenRowBackColor&gt;&lt;OddRowBackColor&gt;Empty&lt;/OddRowBackColor&gt;&lt;ShowColumnHeader&gt;True&lt;/ShowColumnHeader&gt;&lt;ShowColumnFooter&gt;False&lt;/ShowColumnFooter&gt;&lt;ShowRowHeader&gt;True&lt;/ShowRowHeader&gt;&lt;ColumnHeaderBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/ColumnHeaderBackground&gt;&lt;SheetCornerBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/SheetCornerBackground&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;/ActiveSkin&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;80&quot; orientation=&quot;Horizontal&quot; count=&quot;34&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;80&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;0&quot;&gt;&lt;Size&gt;220&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;1&quot;&gt;&lt;Size&gt;85&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;2&quot;&gt;&lt;Size&gt;78&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;4&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;5&quot;&gt;&lt;Size&gt;65&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;6&quot;&gt;&lt;Size&gt;30&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;8&quot;&gt;&lt;Size&gt;180&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;9&quot;&gt;&lt;Size&gt;250&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;10&quot;&gt;&lt;Size&gt;55&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;11&quot;&gt;&lt;Size&gt;30&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;12&quot;&gt;&lt;Size&gt;150&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;13&quot;&gt;&lt;Size&gt;180&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;14&quot;&gt;&lt;Size&gt;100&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;15&quot;&gt;&lt;Size&gt;30&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;17&quot;&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;18&quot;&gt;&lt;Size&gt;150&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;19&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;20&quot;&gt;&lt;Size&gt;180&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;21&quot;&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;22&quot;&gt;&lt;Size&gt;70&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;23&quot;&gt;&lt;Size&gt;100&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;24&quot;&gt;&lt;Size&gt;60&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;25&quot;&gt;&lt;Size&gt;90&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;26&quot;&gt;&lt;Size&gt;100&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;27&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;28&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;150&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;29&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;30&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;31&quot;&gt;&lt;Size&gt;120&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;32&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;Item index=&quot;33&quot;&gt;&lt;Visible&gt;True&lt;/Visible&gt;&lt;Size&gt;200&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;34&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;34&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot; /&gt;&lt;ColumnStyles&gt;&lt;ColumnStyle Index=&quot;0&quot;&gt;&lt;Font&gt;&lt;Name&gt;Arial&lt;/Name&gt;&lt;Names&gt;&lt;Name&gt;Arial&lt;/Name&gt;&lt;/Names&gt;&lt;Size&gt;10pt&lt;/Size&gt;&lt;Bold&gt;False&lt;/Bold&gt;&lt;Italic&gt;False&lt;/Italic&gt;&lt;Overline&gt;False&lt;/Overline&gt;&lt;Strikeout&gt;False&lt;/Strikeout&gt;&lt;Underline&gt;False&lt;/Underline&gt;&lt;/Font&gt;&lt;GdiCharSet&gt;254&lt;/GdiCharSet&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;VerticalAlign&gt;Middle&lt;/VerticalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;1&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;2&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;4&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;5&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;6&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;8&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;9&quot;&gt;&lt;HorizontalAlign&gt;Left&lt;/HorizontalAlign&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;10&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;12&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;13&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;14&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;/ColumnStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;34&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;false&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;AutoPostBack&gt;True&lt;/AutoPostBack&gt;&lt;DataAutoCellTypes&gt;False&lt;/DataAutoCellTypes&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;PageSize&gt;20&lt;/PageSize&gt;&lt;OperationMode&gt;RowMode&lt;/OperationMode&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;ColumnCount&gt;34&lt;/ColumnCount&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;DefaultColumnWidth&gt;80&lt;/DefaultColumnWidth&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;Value type=&quot;System.String&quot; whitespace=&quot;&quot; /&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;" 
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
    </form>
</body>
</html>
