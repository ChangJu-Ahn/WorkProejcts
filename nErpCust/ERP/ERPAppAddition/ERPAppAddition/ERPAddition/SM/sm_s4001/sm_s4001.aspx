<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_s4001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_s4001.sm_s4001" %>

<%@ Register Assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457"
    Namespace="FarPoint.Web.Spread" TagPrefix="FarPoint" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>FCST</title>
    <style type="text/css">
        .title
        {
            font-family: 굴림체;
            font-size: 10pt;
            text-align: left;
            font-weight: bold;
            background-color: #EAEAEA;
            color: Blue;
            vertical-align: middle;
            display: table-cell;
            line-height: 25px;
            height: 25px;
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
        
        .style3
        {
            width: 1438px;
            height: 7px;
        }
        
        .style17
        {
            height: 45px;
            width: 30px;
        }
        .style25
        {
            width: 100px;
            height: 22px;
        }
        
        .style53
        {
            width: 230px;
            height: 22px;
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
        
        .style57
        {
            width: 81px;
            font-family: 굴림체;
            font-size: smaller;
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
            width: 450px;
            height: 22px;
        }
        
        .style60
        {
            width: 74px;
            height: 22px;
        }
        .style61
        {
            width: 216px;
            height: 22px;
        }
        .style62
        {
            width: 320px;
            height: 22px;
        }
        .style63
        {
            width: 84px;
            font-family: 굴림체;
            font-size: smaller;
            text-align: center;
            background-color: Silver;
            height: 22px;
            font-weight: 700;
        }
        .style67
        {
            width: 81px;
            font-family: 굴림체;
            font-size: smaller;
            text-align: center;
            background-color: Silver;
            height: 22px;
            font-weight: 700;
        }
        
        .style68
        {
            width: 330px;
            height: 22px;
        }
        .style69
        {
            width: 87px;
            font-family: 굴림체;
            font-size: smaller;
            text-align: center;
            background-color: Silver;
            height: 22px;
            font-weight: 700;
        }
    </style>
    <%-- <script  type="text/javascript" id="FpSpread_amt_Script4">
		function FpSpread_amt_EditStopped(event){
		    var spread = document.all("FpSpread_amt");
		    spread.SetValue(spread.ActiveRow, 0, "수정", true);
		}
	</script>   --%>
    <script type="text/javascript" id="FpSpread_amt_Script3">
        function FpSpread_amt_DataChanged(event) {
            var spread = document.all("FpSpread_amt");
            spread.SetValue(spread.ActiveRow, 0, "수정", true);
        }
    </script>
</head>
<body>
    <form id="form1" runat="server" enctype="multipart/form-data">
    <div>
        <table>
            <tr>
                <td>
                    <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
                </td>
                <td style="width: 100%;">
                    <asp:Label ID="Label4" runat="server" Text="영업FCST 관리" CssClass="title" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table style="border: thin solid #000080; height: 31px;">
            <td class="style2">
                <asp:Label ID="Label17" runat="server" Text="조회구분" BackColor="#99CCFF" Font-Bold="True"
                    Style="text-align: center; font-size: small"></asp:Label>
            </td>
            <td class="style3">
                <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small" RepeatDirection="Horizontal"
                    OnSelectedIndexChanged="rbl_view_type_SelectedIndexChanged" AutoPostBack="True"
                    Width="417px" Style="margin-left: 0px; font-weight: 700;" BackColor="White" Height="16px">
                    <asp:ListItem Value="A">기준정보등록</asp:ListItem>
                    <asp:ListItem Value="B">FCST / 단가등록</asp:ListItem>
                    <asp:ListItem Value="C">FCST / 예상매출조회</asp:ListItem>
                </asp:RadioButtonList>
            </td>
        </table>
        <%-- <asp:Panel ID="Panel_bas_info" runat="server" Visible="False" 
                  Width="37%">
                   <asp:RadioButtonList ID="rbl_bas_type" runat="server" AutoPostBack="True" 
                       BackColor="White" Font-Size="Small" 
                       onselectedindexchanged="rbl_bas_type_SelectedIndexChanged" 
                       RepeatDirection="Horizontal" style="font-weight: 700; margin-left: 0px; background-color: #FFFFFF;" 
                       Visible="false" Width="396px" BorderColor="#999999" >
                       <asp:ListItem Selected="True" Value="A">소분류 품목그룹 등록</asp:ListItem>
                       <asp:ListItem Value="B">대분류-소분류 품목그룹 연결</asp:ListItem>
                   </asp:RadioButtonList>
               </asp:Panel></td>--%>
        <asp:Panel ID="Panel_qty_amt" runat="server" Visible="False">
            <asp:RadioButtonList ID="rdl_qty_amt" runat="server" Font-Size="Small" RepeatDirection="Horizontal"
                AutoPostBack="True" OnSelectedIndexChanged="rbl_qty_amt_type_SelectedIndexChanged"
                Style="font-weight: 700; text-align: left; background-color: #FFFFFF;" BackColor="White"
                Width="197px" BorderColor="#999999">
                <asp:ListItem Value="A" Selected="True">수량</asp:ListItem>
                <asp:ListItem Value="B">단가</asp:ListItem> 
                <%--<asp:ListItem Value="B">예상매출</asp:ListItem>--%>                 
            </asp:RadioButtonList>
        </asp:Panel>
        <asp:Panel ID="panel_upload" runat="server" Visible="False" BorderStyle="Groove"
            BorderColor="White" Width="99%">
            <table style="width: 93%; margin-bottom: 0px;">
                <tr>
                    <td class="style69">
                        <asp:Label ID="Label13" runat="server" Font-Size="Small" text-align="center" Text="Version선택"
                            BackColor="Silver" Style="font-weight: 700"></asp:Label>
                    </td>
                    <td class="style60">
                        <asp:DropDownList ID="list_regist_version" runat="server" BackColor="#FFFFCC" Width="86px"
                            Height="21px">
                            <asp:ListItem>-선택안함-</asp:ListItem>
                            <asp:ListItem>R0</asp:ListItem>
                            <asp:ListItem>R1</asp:ListItem>
                            <asp:ListItem>R2</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td class="style63">
                        &nbsp;날짜선택
                    </td>
                    <td class="style61">
                        <asp:Label ID="Label15" runat="server" Font-Size="Small" Text="*년: "></asp:Label>
                        <asp:TextBox ID="txt_regist_date_yyyy" runat="server" BackColor="#FFFFCC" Height="16px"
                            Width="48px"></asp:TextBox>
                        <asp:Label ID="Label16" runat="server" Font-Size="Small" Text="*월: "></asp:Label>
                        <asp:DropDownList ID="txt_regist_date_mm" runat="server" BackColor="#FFFFCC" Style="text-align: center">
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
                    </td>
                    <td class="style67">
                        &nbsp;Excel선택
                    </td>
                    <td class="style68">
                        <asp:FileUpload ID="FileUpload1" runat="server" BackColor="#FFFFCC" Width="224px" />
                        &nbsp;<asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="Upload"
                            ToolTip="엑셀자료를 가져와 화면에 보여준다." Width="87px" Height="21px" />
                    </td>
                    <td class="style57" style="border-style: none; font-weight: 700;">
                        &nbsp;Sheet선택
                    </td>
                    <td class="style62">
                        <asp:DropDownList ID="ddlSheets" runat="server" BackColor="#FFFFCC" AutoPostBack="true"
                            OnSelectedIndexChanged="ddlSheets_OnSelectedIndexChanged" Height="20px">
                        </asp:DropDownList>
                        <asp:Button ID="btnSave" runat="server" OnClick="btnSave_Click" OnClientClick="return confirm('당월데이타 존재시 삭제후 저장됩니다. 진행하시겠습니까?');"
                            Text="Save" Width="80px" />
                        <asp:Button ID="btnCancel0" runat="server" OnClick="btnCancel_Click" Text="Cancel"
                            Width="80px" />
                    </td>
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
        <asp:Panel ID="Panel_select" runat="server" Visible="False" BorderStyle="Groove"
            BorderColor="White" Width="99%">
            <table style="width: 100%" rules="all" border="1">
                <tr>
                    <td class="style58">
                        <asp:Label ID="Label5" runat="server" Font-Size="Small" text-align="center" Text="Version선택"
                            BackColor="Silver" Style="font-weight: 700"></asp:Label>
                    </td>
                    <td class="style25">
                        <asp:DropDownList ID="list_select_version" runat="server" BackColor="#FFFFCC" Style="text-align: center">
                            <asp:ListItem>-선택안함-</asp:ListItem>
                            <asp:ListItem>R0</asp:ListItem>
                            <asp:ListItem>R1</asp:ListItem>
                            <asp:ListItem>R2</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td class="style55">
                        날짜선택
                    </td>
                    <td class="style53">
                        <asp:Label ID="Label6" runat="server" Font-Size="Small" Text="*년: "></asp:Label>
                        <asp:TextBox ID="txt_select_date_yyyy" runat="server" BackColor="#FFFFCC" Height="16px"
                            Width="57px"></asp:TextBox>
                        <asp:Label ID="Label7" runat="server" Font-Size="Small" Text="*월: "></asp:Label>
                        <asp:DropDownList ID="txt_select_date_mm" runat="server" BackColor="#FFFFCC" Style="text-align: center">
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
                    </td>
                    <td class="style55">
                        분류선택
                    </td>
                    <td class="style59">
                        <asp:Label ID="Label3" runat="server" Text="*대분류" Style="font-size: small"></asp:Label>
                        <asp:DropDownList ID="ddl_itemgp_select" runat="server" AutoPostBack="true" BackColor="#FFFFCC"
                            DataSourceID="SqlDataSource1_ddl_itemgp" DataTextField="L_ITEM_AMT_GROUP" DataValueField="L_ITEM_AMT_GROUP"
                            OnSelectedIndexChanged="ddl_itemgp_select_SelectedIndexChanged" Style="text-align: left">
                        </asp:DropDownList>
                        &nbsp;<asp:Label ID="Label19" runat="server" Text="*중분류" Style="font-size: small"></asp:Label>
                        <asp:DropDownList ID="ddl_itemgp_select_amt" runat="server" AutoPostBack="true" BackColor="#FFFFCC"
                            DataTextField="M_ITEM_AMT_GROUP" DataValueField="M_ITEM_AMT_GROUP" OnSelectedIndexChanged="ddl_itemgp_select_amt_SelectedIndexChanged">
                        </asp:DropDownList>
                        <asp:Label ID="Label20" runat="server" Style="font-size: small" Text="*소분류"></asp:Label>
                        <asp:DropDownList ID="ddl_itemgp_select_s_amt" runat="server" BackColor="#FFFFCC" AutoPostBack = "true"
                            DataTextField="S_ITEM_AMT_GROUP" DataValueField="S_ITEM_AMT_GROUP" OnSelectedIndexChanged ="ddl_exchange_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                    <td class="style55">
                        환율
                    </td>
                    <td>
                        <asp:DropDownList ID="txt_to_currency1" runat="server" BackColor="#FFFFCC" Width="90px">
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem>KRW</asp:ListItem> 
                            <asp:ListItem>EUR</asp:ListItem>
                            <asp:ListItem>RUB</asp:ListItem>
                            <asp:ListItem>CNY</asp:ListItem>
                            <asp:ListItem>GBP</asp:ListItem>
                            <asp:ListItem>JPY</asp:ListItem>
                            <asp:ListItem>SGD</asp:ListItem>
                            <asp:ListItem>CHF</asp:ListItem>
                            <asp:ListItem>HKD</asp:ListItem>
                            <asp:ListItem>USD</asp:ListItem>
                        </asp:DropDownList>
                        <%--<asp:TextBox ID="txt_to_currency" runat="server" BackColor="#FFFFCC" Width="90px"></asp:TextBox>--%>
                        <asp:TextBox ID="txt_exchange" runat="server" BackColor="#FFFFCC" Width="100px"></asp:TextBox>
                    </td>
                </tr>
            </table>
            <asp:DropDownList ID="list_select" runat="server" BackColor="#FFFFCC" Height="20px">
                <asp:ListItem>-선택안함-</asp:ListItem>
                <asp:ListItem>대분류</asp:ListItem>
                <asp:ListItem>고객사</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btn_select" runat="server" OnClick="btn_select_Click" Text="조회" Width="100px" />
        </asp:Panel>
        <asp:Panel ID="Panel_regist_excel_grid" runat="server" Visible="False">
            <asp:GridView ID="grid_regist_excel" runat="server" CellPadding="4" GridLines="None">
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
        <asp:Panel ID="Panel_select_excel_qty_grid" runat="server">
            <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="100%" AsyncRendering="False"
                Height="600px" SizeToReportContent="True">
            </rsweb:ReportViewer>
        </asp:Panel>
        <asp:Panel ID="Panel_select_excel_amt_grid" runat="server">
            <rsweb:ReportViewer ID="ReportViewer2" runat="server" Width="100%" AsyncRendering="False"
                Height="600px" SizeToReportContent="True">
            </rsweb:ReportViewer>
        </asp:Panel>
        <%--  </ContentTemplate>
                <Triggers>
                    <asp:AsyncPostBackTrigger ControlID="btn_device_amt_view" />
                </Triggers>
       </asp:UpdatePanel>--%>
        <script type="text/javascript">
            var ModalProgress = '<%= ModalProgress.ClientID %>';

            Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(beginReq);
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(endReq);
            function beginReq(sender, args) {
                //show the Popup
                $find(ModalProgress).show()
            }
            function endReq(sender, args) {
                //hide the Popup
                $find(ModalProgress).hide();
            }
        </script>
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="Btn_select"></asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                <br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
        <cc1:ModalPopupExtender ID="ModalProgress" runat="server" PopupControlID="UpdateProg1"
            TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        <asp:Panel ID="Panel_Spread_bas" runat="server" Visible="False">
            <asp:Panel ID="Panel_Spread_Btn" runat="server" Visible="False">
                <asp:TextBox ID="tb_rowcnt" runat="server" BorderStyle="Outset" Width="30px" BackColor="#FFFFCC">1</asp:TextBox>
                <asp:Button ID="btn_Add" runat="server" Text="Row추가" Width="100px" OnClick="btn_Add_Click" />
                <asp:Button ID="btn_Delete" runat="server" Text="삭제" Width="100px" OnClick="btn_Delete_Click" />
                <asp:Button ID="btn_save" runat="server" Text="저장" Width="100px" OnClick="btn_save_Click" />
                <asp:Button ID="btn_exe" runat="server" Text="조회" Width="100px" OnClick="btn_exe_Click" />
            </asp:Panel>
            <FarPoint:FpSpread ID="FpSpread1_ITEMGR" runat="server" BorderColor="Black" BorderStyle="Solid"
                BorderWidth="1px" Height="500px" Width="681px" ActiveSheetViewIndex="0" CommandBarOnBottom="False"
                DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;"
                OnUpdateCommand="FpSpread1_ITEMGR_UpdateCommand" CssClass="spread" currentPageIndex="0">
                <CommandBar BackColor="Control" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight"
                    ButtonShadowColor="ControlDark" ButtonType="LinkButton" ShowPDFButton="True"
                    Theme="Office2007" Visible="False">
                    <Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif" />
                </CommandBar>
                <Pager Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False"
                    Font-Underline="False" />
                <HierBar Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False"
                    Font-Underline="False" />
                <Sheets>
                    <FarPoint:SheetView DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;0&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;대분류&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;중분류&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;소분류&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;비고&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;0&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;ActiveSkin class=&quot;FarPoint.Web.Spread.SheetSkin&quot;&gt;&lt;Name&gt;Classic&lt;/Name&gt;&lt;BackColor&gt;Empty&lt;/BackColor&gt;&lt;CellBackColor&gt;Empty&lt;/CellBackColor&gt;&lt;CellForeColor&gt;Empty&lt;/CellForeColor&gt;&lt;CellSpacing&gt;0&lt;/CellSpacing&gt;&lt;GridLines&gt;Both&lt;/GridLines&gt;&lt;GridLineColor&gt;LightGray&lt;/GridLineColor&gt;&lt;HeaderBackColor&gt;Control&lt;/HeaderBackColor&gt;&lt;HeaderForeColor&gt;Empty&lt;/HeaderForeColor&gt;&lt;FlatColumnHeader&gt;False&lt;/FlatColumnHeader&gt;&lt;FooterBackColor&gt;Empty&lt;/FooterBackColor&gt;&lt;FooterForeColor&gt;Empty&lt;/FooterForeColor&gt;&lt;FlatColumnFooter&gt;False&lt;/FlatColumnFooter&gt;&lt;FlatRowHeader&gt;False&lt;/FlatRowHeader&gt;&lt;HeaderFontBold&gt;False&lt;/HeaderFontBold&gt;&lt;FooterFontBold&gt;False&lt;/FooterFontBold&gt;&lt;SelectionBackColor&gt;LightBlue&lt;/SelectionBackColor&gt;&lt;SelectionForeColor&gt;Empty&lt;/SelectionForeColor&gt;&lt;EvenRowBackColor&gt;Empty&lt;/EvenRowBackColor&gt;&lt;OddRowBackColor&gt;Empty&lt;/OddRowBackColor&gt;&lt;ShowColumnHeader&gt;True&lt;/ShowColumnHeader&gt;&lt;ShowColumnFooter&gt;False&lt;/ShowColumnFooter&gt;&lt;ShowRowHeader&gt;True&lt;/ShowRowHeader&gt;&lt;ColumnHeaderBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/ColumnHeaderBackground&gt;&lt;SheetCornerBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/SheetCornerBackground&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;/ActiveSkin&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;Item index=&quot;3&quot;&gt;&lt;Size&gt;105&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;0&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;4&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;0&quot; Columns=&quot;4&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot; /&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;4&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;false&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;PageSize&gt;50&lt;/PageSize&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;RowCount&gt;0&lt;/RowCount&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;"
                        SheetName="Sheet1" PageSize="50">
                    </FarPoint:SheetView>
                </Sheets>
                <TitleInfo BackColor="#E7EFF7" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                    Font-Size="X-Large" Font-Strikeout="False" Font-Underline="False" ForeColor=""
                    HorizontalAlign="Center" VerticalAlign="NotSet">
                </TitleInfo>
            </FarPoint:FpSpread>
        </asp:Panel>
        <%-- <asp:Panel ID="Panel_routeset" runat="server" Visible="False">
        <table>
        <tr><td>
            <asp:Label ID="Label1" runat="server" Text="소분류품목그룹" CssClass="style13"></asp:Label>
            </td><td></td><td>
                <asp:Label ID="Label2" runat="server" Text="대분류품목그룹" CssClass="style13"></asp:Label>
                <asp:DropDownList ID="ddl_itemgp" runat="server" 
                DataSourceID="SqlDataSource1_ddl_itemgp" DataTextField="ITEM_GROUP" 
                DataValueField="ITEM_GROUP" >
            </asp:DropDownList>--%>
        <asp:SqlDataSource ID="SqlDataSource1_ddl_itemgp" runat="server" ConnectionString="<%$ ConnectionStrings:MES_CCUBE_UNIERP %>"
            ProviderName="<%$ ConnectionStrings:MES_CCUBE_UNIERP.ProviderName %>" SelectCommand="select DISTINCT L_ITEM_AMT_GROUP from T_DEVICE_AMT_GROUP_ADD union all select '-선택안됨-' from dual where rownum < 2 order by 1">
        </asp:SqlDataSource>
        <%--<asp:Button ID="btn_exe_itemgp_routeset" runat="server" Text="조회" 
                Width="100px" />
        </td></tr>
        <tr>
        <td><asp:ListBox ID="lsb_l_routeset" runat="server" Height="450px" Width="430px" 
                BackColor="#F6F6F6" SelectionMode="Multiple"></asp:ListBox></td>
        <td><asp:Button ID="btn_move_right" runat="server" Text=">" OnClick="btn_move_right_Click" />
                    <br />
                    <br />
                    <asp:Button ID="btn_move_left" runat="server" Text="<" OnClick="btn_move_left_Click" /></td>
        <td>
            <asp:ListBox ID="lsb_r_routeset" runat="server" Height="450px" Width="430px" 
                BackColor="#F6F6F6" SelectionMode="Multiple"></asp:ListBox>
        </td>
        </tr></table>
            
        </asp:Panel>--%>
    </div>
    </form>
</body>
</html>
