<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_s3001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_s3001.sm_s3001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>인보이스관리</title>
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
        .spread
        {
            width: 120px;
            
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
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
        .default_font_size
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align:center;
        }
        .default_font_background
        {
            font-family: 굴림체;
            font-size:10pt;
            
        }
    </style>
    <script type="text/javascript">
        function myCheckFunction() {
            var spread = document.getElementById("FpSpread_new_data");
            var index = event.srcElement.id.toString();
            var value = event.srcElement.checked;
            var splitstr = index.split(",");
            var rows = spread.GetRowCount();
            if (splitstr[2] == "ch") {
                for (var i = 0; i < rows; i++) {
                    spread.SetValue(i, splitstr[1], value);

                }

            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager2" runat="server">
    </asp:ScriptManager>
    <div>
    <table><tr><td>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td  style="width:100%;"><asp:Label ID="Label2" runat="server" Text="Invoice관리" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>       
        
    </div>
    <asp:Panel ID="Panel_menu1" runat="server">
    <table style=" table-layout:fixed; border: thin solid #000080; width:100%;">
            <tr>
            <td>
            <table><tr>
                <td class="style12" >
                    <strong>작업구분</strong>
                </td>
                <td class="style1">
                    <asp:RadioButtonList ID="rbtnl_chk_process" runat="server" CssClass="default_font_size"
                        RepeatDirection="Horizontal" 
                        onselectedindexchanged="rbtnl_chk_process_SelectedIndexChanged" 
                        AutoPostBack="True" >
                        <asp:ListItem Value="new" Selected="True">생성</asp:ListItem>
                        <asp:ListItem Value="view">수정</asp:ListItem>
                    </asp:RadioButtonList>
                </td>      
               </tr></table>       
             </td>   
            </tr>
        </table>
        
    </asp:Panel>
    <asp:Panel ID="Panel_menu2" runat="server">
   
    <table style=" table-layout:fixed; border: thin solid #000080; width:100%; ">
            <tr><td>
            <table><tr>
                <td class="style12">
                    <strong>인보이스번호</strong>
                </td>
                <td >
                    <asp:TextBox ID="tb_invoice_no" runat="server"></asp:TextBox>
                </td>
                <td>
                    <asp:UpdatePanel ID="UpdatePanel3" runat="server" UpdateMode="Conditional">
                        <ContentTemplate>
                            <asp:Button ID="btn_pop_invoice" runat="server" Text="..." Width="20px" OnClick="btn_pop_invoice_Click" />
                        </ContentTemplate>
                    <Triggers>
                   <asp:PostBackTrigger ControlID = "btn_pop_invoice" />
                   </Triggers>

                    </asp:UpdatePanel>
                </td>
                <td>
                    <asp:Button ID="btn_retrieve" runat="server" Text="조회" Width = "100px" 
                        onclick="btn_retrieve_Click" />
                </td>
                <td>
                    <asp:Button ID="btn_copy" runat="server" Text="Copy" Width = "100px" 
                        onclick="btn_copy_Click" />
                </td>      
                <td>
                    <asp:Button ID="btn_update" runat="server" Text="저장" Width = "100px" 
                        onclick="btn_update_Click" style="height: 21px" />
                </td>      
                <td>
                    <asp:Button ID="btn_delete" runat="server" Text="삭제" Width = "100px" /></td>
                <td>
                    <asp:Button ID="btn_preview" runat="server" Text="미리보기" Width = "100px" 
                        onclick="btn_preview_Click" /></td>
                <td>
                <table>
                <tr>
                <td>
                <asp:CheckBox ID="cb_view_lot" runat="server"  CssClass="default_font_size" 
                        Text="출력LOT보기"  /></td>
                          <td>
                              <asp:Label ID="Label35" runat="server" Text="*출력단가선택:"  CssClass="default_font_size"></asp:Label>
                              <asp:CheckBox ID="cb_price1" runat="server"   CssClass="default_font_size" 
                                  Text="단가" />
                              <asp:CheckBox ID="cb_price2" runat="server"   CssClass="default_font_size" 
                                  Text="원자재" />
                              <asp:CheckBox ID="cb_price3" runat="server"   CssClass="default_font_size" 
                                  Text="3자국" />
                </td>

                <tr>
                <td>
                 <asp:CheckBox ID="cb_view_pono" runat="server"  CssClass="default_font_size" 
                        Text="PO보기"  />
                </td>
                </tr></table>
                    
                </td>
               </tr></table>       
             </td>    
            </tr>
        </table>
        
    </asp:Panel>   
    
    
    <asp:Panel ID="Panel_body1" runat="server" Visible = "false">
        <table style="width:100%;">
           <tr><td>
               <asp:Panel ID="Panel_new_invoice_no" runat="server">
               <table><tr><td><asp:Label ID="Label32" runat="server" Text="신규인보이스번호"  CssClass="default_font_size"></asp:Label></td>
               <td><asp:TextBox ID="tb_new_invoice_no" runat="server"></asp:TextBox></td>
               </tr></table>
               </asp:Panel>
               
          </td>
           </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label3" runat="server" Text="1. 유무상구분" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rbtnl_pay_type" runat="server" CssClass="default_font_size"
                                    RepeatDirection="Horizontal">
                                    <asp:ListItem Selected="True" Value="COMMERCIAL">Commercial</asp:ListItem>
                                    <asp:ListItem Value="NONCOM">Non-commercial</asp:ListItem>
                                    <asp:ListItem Value="PROFOMA">Proforma</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td style=" height: 25px;">
                                <asp:Label ID="Label4" runat="server" Text="2. 대상구분" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rbtnl_target_type" runat="server" CssClass="default_font_size"
                                    RepeatDirection="Horizontal">
                                    <asp:ListItem Selected="True" Value="C">고객용</asp:ListItem>
                                    <asp:ListItem Value="B">면허용</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td style="background-color:Yellow; height: 25px;">
                                <asp:Label ID="Label31" runat="server" Text="확정여부" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" CssClass="default_font_size"
                                    RepeatDirection="Horizontal">
                                    <asp:ListItem Selected="True" Value="Y">확정</asp:ListItem>
                                    <asp:ListItem Value="N">미확정</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                 <ContentTemplate>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label5" runat="server" Text="3. 발신인정보" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                
                                <table>
                                    <tr>
                                        <td style="height: 25px;">
                                            <asp:TextBox ID="tb_ship_fr_cust_nm" runat="server" Width="80px"></asp:TextBox>
                                            <asp:Button ID="btn_pop_ship_fr_cust_cd" runat="server" Text="찾기" OnClick="btn_pop_ship_fr_cust_cd_Click" />
                                            <asp:CheckBox ID="cb_nepes_addr" runat="server" AutoPostBack="True" Text="nepes"   
                                                CssClass="default_font_size" 
                                                oncheckedchanged="cb_nepes_addr_CheckedChanged"/>
                                        </td>
                                        <td>
                                            <asp:Label ID="Label6" runat="server" Text="영문주소:" CssClass="default_font_size"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:TextBox ID="tb_ship_fr_add" runat="server" Width="370px" 
                                                BackColor="#CCCCCC" TextMode="MultiLine"></asp:TextBox>
                                        </td>
                                        <td>
                                            <asp:Label ID="Label7" runat="server" Text="Tel:" CssClass="default_font_size"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:TextBox ID="tb_ship_fr_tel" runat="server" Width="100px" 
                                                BackColor="#CCCCCC"></asp:TextBox>
                                        </td>
                                        <td>
                                            <asp:Label ID="Label8" runat="server" Text="Fax:" CssClass="default_font_size"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:TextBox ID="tb_ship_fr_fax" runat="server" Width="100px" 
                                                BackColor="#CCCCCC"></asp:TextBox>
                                        </td>                                        
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    </ContentTemplate>
                   <Triggers>
                   <asp:PostBackTrigger ControlID = "btn_pop_ship_fr_cust_cd" />
                   </Triggers>
                    </asp:UpdatePanel>
                </td>
            </tr>
            <tr>
                <td>
                <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                 <ContentTemplate>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label9" runat="server" Text="4. 수취인정보" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <table>
                                    <tr>
                                        <td style="height: 25px;">
                                            <asp:TextBox ID="tb_bill_to_cust_nm" runat="server"></asp:TextBox>
                                            <asp:Button ID="btn_pop_bill_cust_cd"
                                                runat="server" Text="찾기" onclick="btn_pop_bill_cust_cd_Click" />
                                                
                                        </td>
                                        <td>
                                <asp:Label ID="Label10" runat="server" Text="영문주소:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_bill_to_add" runat="server" Width="370px" 
                                    BackColor="#CCCCCC" TextMode="MultiLine"></asp:TextBox>
                            </td>
                            <td>
                                <asp:Label ID="Label11" runat="server" Text="Tel:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_bill_to_tel" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                            </td>
                            <td>
                                <asp:Label ID="Label12" runat="server" Text="Fax:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_bill_to_fax" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                            </td>
                            <td>
                                <asp:Label ID="Label33" runat="server" Text="수취인:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_bill_to_name" runat="server" Width="100px"></asp:TextBox>
                            </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                     </ContentTemplate>
                   <Triggers><asp:PostBackTrigger ControlID = "btn_pop_bill_cust_cd" /></Triggers>
                    </asp:UpdatePanel>
                </td>
            </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label13" runat="server" Text="5. 실물수령인" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <table>
                                    <tr>
                                        <td style="height: 25px;">
                                            <asp:TextBox ID="tb_ship_to_cust_nm" runat="server"></asp:TextBox>
                                            <asp:Button ID="btn_pop_ship_to_cust_cd"
                                                runat="server" Text="찾기" onclick="btn_pop_ship_to_cust_cd_Click" />
                                                 
                                        </td>
                                        <td>
                                <asp:Label ID="Label14" runat="server" Text="영문주소:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_ship_to_add" runat="server" Width="370px"
                                    BackColor="#CCCCCC" TextMode="MultiLine"></asp:TextBox>
                            </td>
                            <td>
                                <asp:Label ID="Label15" runat="server" Text="Tel:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_ship_to_tel" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                            </td>
                            <td>
                                <asp:Label ID="Label16" runat="server" Text="Fax:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_ship_to_fax" runat="server" Width="100px" ReadOnly="True" BackColor="#CCCCCC"></asp:TextBox>
                            </td>
                            <td>
                                <asp:Label ID="Label34" runat="server" Text="수령인:" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_ship_to_name" runat="server" Width="100px" ></asp:TextBox>
                            </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label17" runat="server" Text="6. 출발지" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_port_of_loading" runat="server" BackColor="#CCCCCC"
                                    Width="100px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label18" runat="server" Text="7. 도착지" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_final_destination" runat="server" BackColor="#CCCCCC"
                                    Width="100px"></asp:TextBox>
                            </td>
                            <td>
                                <asp:Button ID="btn_so_no" runat="server" Text="수주적용" />
                            </td>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label19" runat="server" Text="8. 운송업체" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_carrier" runat="server" Width="100px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label20" runat="server" Text="9. 발송일" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_board_on_about" runat="server" BackColor="#CCCCCC"
                                    Width="100px" MaxLength="8"></asp:TextBox>
                                <cc1:CalendarExtender ID="tb_board_on_about_CalendarExtender" runat="server" 
                                    Enabled="True" Format="yyyyMMdd" TargetControlID="tb_board_on_about">
                                </cc1:CalendarExtender>
                            </td>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label21" runat="server" Text="10. 발행일" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_invoice_dt" runat="server" BackColor="#CCCCCC"
                                    Width="100px" MaxLength="8" CssClass="default_font_size"></asp:TextBox>
                                <cc1:CalendarExtender ID="tb_invoice_dt_CalendarExtender" runat="server" 
                                    Enabled="True" Format="yyyyMMdd" TargetControlID="tb_invoice_dt">
                                </cc1:CalendarExtender>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label22" runat="server" Text="11. REMARK" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_remark" runat="server" Width="500px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label23" runat="server" Text="11-1.운임조건" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_remark_incoterms" runat="server" Width="100px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px; width: 120px;">
                                <asp:Label ID="Label24" runat="server" Text="11-2.유무상조건" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_remark_pay_type" runat="server" Width="100px"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 120px;">
                                <asp:Label ID="Label25" runat="server" Text="12.전체박스수량" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_total_box_cnt" runat="server" Width="100px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px;">
                                <asp:Label ID="Label26" runat="server" Text="13.HTS CODE" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_hts_code" runat="server" Width="100px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px;">
                                <asp:Label ID="Label27" runat="server" Text="14.원산지" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_country_of_org" runat="server" Width="100px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px;">
                                <asp:Label ID="Label28" runat="server" Text="15.NET-Weight" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_net_weight" runat="server" Width="100px"></asp:TextBox>
                            </td>
                            <td style=" height: 25px;">
                                <asp:Label ID="Label29" runat="server" Text="16.Gross-Weight" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tb_gross_weight" runat="server" Width="100px"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td style=" height: 25px; width: 100px;">
                                <asp:Label ID="Label30" runat="server" Text="17. 은행정보" CssClass="default_font_size"></asp:Label>
                            </td>
                            <td>
                                <table>
                                    <tr>
                                        <td style="height: 25px;">
                                            <asp:TextBox ID="tb_bank_info" runat="server"></asp:TextBox><asp:Button ID="btn_pop_bank_info"
                                                runat="server" Text="찾기" />
                                            <asp:CheckBox ID="cb_use_nepes_add" runat="server" CssClass="default_font_size" 
                                                AutoPostBack="True" oncheckedchanged="cb_use_nepes_add_CheckedChanged" 
                                                Text="신한BK서초남지점" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td>
                                <asp:TextBox ID="tb_bank_name" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                                <asp:TextBox ID="tb_bank_addr" runat="server" Width="300px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                                <asp:TextBox ID="tb_bank_branch" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                                <asp:TextBox ID="tb_bank_swiftcode" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                                <asp:TextBox ID="tb_bank_acct_no" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                                <asp:TextBox ID="tb_bank_accountee" runat="server" Width="100px" 
                                    BackColor="#CCCCCC"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Panel ID="Panel_view_data" runat="server">
        <div>
            <asp:Button ID="btn_spread_view_data_update" runat="server" Text="수정" 
                onclick="btn_spread_view_data_update_Click" Width="100px" /> 
            <asp:Button ID="btn_spread_view_data_delete" runat="server" Text="삭제" 
                Width="100px" onclick="btn_spread_view_data_delete_Click" /></div>
        <FarPoint:FpSpread ID="FpSpread_view_data" runat="server" BorderColor="Black" BorderStyle="Solid"
            BorderWidth="1px" Height="400px" Width="100%" CssClass="default_font_size" 
                onupdatecommand="FpSpread_view_data_UpdateCommand" 
                ActiveSheetViewIndex="0" HorizontalScrollBarPolicy="Always" 
                VerticalScrollBarPolicy="Always"  
                
                
                DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;" 
                >
            <CommandBar BackColor="Control" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight"
                ButtonShadowColor="ControlDark">
                <Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif" />
            </CommandBar>
            <Pager Font-Bold="False" Font-Italic="False" Font-Overline="False" 
                Font-Strikeout="False" Font-Underline="False" />
            <HierBar Font-Bold="False" Font-Italic="False" Font-Overline="False" 
                Font-Strikeout="False" Font-Underline="False" />
            <Sheets>
                <FarPoint:SheetView DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;21&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;21&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;21&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;ActiveSkin class=&quot;FarPoint.Web.Spread.SheetSkin&quot;&gt;&lt;Name&gt;Classic&lt;/Name&gt;&lt;BackColor&gt;Empty&lt;/BackColor&gt;&lt;CellBackColor&gt;Empty&lt;/CellBackColor&gt;&lt;CellForeColor&gt;Empty&lt;/CellForeColor&gt;&lt;CellSpacing&gt;0&lt;/CellSpacing&gt;&lt;GridLines&gt;Both&lt;/GridLines&gt;&lt;GridLineColor&gt;LightGray&lt;/GridLineColor&gt;&lt;HeaderBackColor&gt;Control&lt;/HeaderBackColor&gt;&lt;HeaderForeColor&gt;Empty&lt;/HeaderForeColor&gt;&lt;FlatColumnHeader&gt;False&lt;/FlatColumnHeader&gt;&lt;FooterBackColor&gt;Empty&lt;/FooterBackColor&gt;&lt;FooterForeColor&gt;Empty&lt;/FooterForeColor&gt;&lt;FlatColumnFooter&gt;False&lt;/FlatColumnFooter&gt;&lt;FlatRowHeader&gt;False&lt;/FlatRowHeader&gt;&lt;HeaderFontBold&gt;False&lt;/HeaderFontBold&gt;&lt;FooterFontBold&gt;False&lt;/FooterFontBold&gt;&lt;SelectionBackColor&gt;LightBlue&lt;/SelectionBackColor&gt;&lt;SelectionForeColor&gt;Empty&lt;/SelectionForeColor&gt;&lt;EvenRowBackColor&gt;Empty&lt;/EvenRowBackColor&gt;&lt;OddRowBackColor&gt;Empty&lt;/OddRowBackColor&gt;&lt;ShowColumnHeader&gt;True&lt;/ShowColumnHeader&gt;&lt;ShowColumnFooter&gt;False&lt;/ShowColumnFooter&gt;&lt;ShowRowHeader&gt;True&lt;/ShowRowHeader&gt;&lt;ColumnHeaderBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/ColumnHeaderBackground&gt;&lt;SheetCornerBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/SheetCornerBackground&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;/ActiveSkin&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;21&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;Item index=&quot;0&quot;&gt;&lt;Size&gt;91&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;21&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;21&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot; /&gt;&lt;ColumnStyles&gt;&lt;ColumnStyle Index=&quot;0&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;1&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;2&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;3&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;4&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;5&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;6&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;7&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;8&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;9&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;10&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;11&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;14&quot;&gt;&lt;CellType class=&quot;FarPoint.Web.Spread.DoubleCellType&quot;&gt;&lt;Size&gt;20&lt;/Size&gt;&lt;ErrorMsg&gt;Double: (ex, 1234.56)&lt;/ErrorMsg&gt;&lt;AllowWrap&gt;False&lt;/AllowWrap&gt;&lt;IsDateFormat&gt;False&lt;/IsDateFormat&gt;&lt;GeneralCellType /&gt;&lt;DecimalDigits&gt;4&lt;/DecimalDigits&gt;&lt;DoubleCellType /&gt;&lt;/CellType&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;17&quot;&gt;&lt;CellType class=&quot;FarPoint.Web.Spread.DoubleCellType&quot;&gt;&lt;Size&gt;20&lt;/Size&gt;&lt;ErrorMsg&gt;Double: (ex, 1234.56)&lt;/ErrorMsg&gt;&lt;AllowWrap&gt;False&lt;/AllowWrap&gt;&lt;IsDateFormat&gt;False&lt;/IsDateFormat&gt;&lt;GeneralCellType /&gt;&lt;DecimalDigits&gt;4&lt;/DecimalDigits&gt;&lt;DoubleCellType /&gt;&lt;/CellType&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;20&quot;&gt;&lt;CellType class=&quot;FarPoint.Web.Spread.DoubleCellType&quot;&gt;&lt;Size&gt;20&lt;/Size&gt;&lt;ErrorMsg&gt;Double: (ex, 1234.56)&lt;/ErrorMsg&gt;&lt;AllowWrap&gt;False&lt;/AllowWrap&gt;&lt;IsDateFormat&gt;False&lt;/IsDateFormat&gt;&lt;GeneralCellType /&gt;&lt;DecimalDigits&gt;4&lt;/DecimalDigits&gt;&lt;DoubleCellType /&gt;&lt;/CellType&gt;&lt;/ColumnStyle&gt;&lt;/ColumnStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;21&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;false&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;PageSize&gt;100&lt;/PageSize&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnCount&gt;21&lt;/ColumnCount&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;ActiveRow&gt;0&lt;/ActiveRow&gt;&lt;ActiveColumn&gt;10&lt;/ActiveColumn&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;" 
                    PageSize="100" SheetName="Sheet1">
                </FarPoint:SheetView>
            </Sheets>
            <TitleInfo BackColor="#E7EFF7" Font-Size="X-Large" ForeColor="" 
                HorizontalAlign="Center" VerticalAlign="NotSet" Font-Bold="False" 
                Font-Italic="False" Font-Overline="False" Font-Strikeout="False" 
                Font-Underline="False">
            </TitleInfo>
        </FarPoint:FpSpread>
        </asp:Panel>
        
    </asp:Panel>
    <asp:Panel ID="Panel_body2" runat="server" Visible="false">
        <asp:UpdatePanel ID="UpdatePanel_body2" runat="server" UpdateMode="Conditional">
            <ContentTemplate>
                <table style="border: thin solid #000080; width: 100%">
                    <tr>
                        <td>
                            <table>
                                <tr>
                                    <td class="style12">
                                        <strong>조회기간</strong>
                                    </td>
                                    <td class="style1">
                                        <asp:TextBox ID="tb_fr_yyyymmdd" runat="server" MaxLength="8" Width="70px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_fr_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_fr_yyyymmdd">
                                        </cc1:CalendarExtender>
                                        <asp:Label ID="Label1" runat="server" Text="~"></asp:Label>
                                        <asp:TextBox ID="tb_to_yyyymmdd" runat="server" MaxLength="8" Width="70px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_to_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_to_yyyymmdd">
                                        </cc1:CalendarExtender>
                                    </td>
                                    <td class="style12">
                                        거래처
                                    </td>
                                    <td class="style1">
                                        <asp:DropDownList ID="ddl_cust_cd" runat="server">
                                            <asp:ListItem Value="%">전체</asp:ListItem>
                                        </asp:DropDownList>
                                    </td>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        <asp:Button ID="btn_mighty_retrieve" runat="server" Text="조회" 
                                            OnClick="btn_mighty_retrieve_Click" Width="100px" />
                                    </td>
                                    <td>
                                        <asp:Button ID="btn_mighty_save" runat="server" Width="100px" Text="저장" OnClick="btn_mighty_save_Click" />
                                        <asp:Button ID="btn_select_all" runat="server" OnClick="btn_select_all_Click" 
                                            Text="전체선택" Visible="False" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <FarPoint:FpSpread ID="FpSpread_new_data" runat="server" BorderColor="Black" BorderStyle="Solid"
                    BorderWidth="1px" Height="400px" Width="100%" CommandBarOnBottom="False" CssClass="default_font_size"
                    ActiveSheetViewIndex="0" DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Spread /&gt;"
                    OnUpdateCommand="FpSpread_new_data_UpdateCommand" WaitMessage="조회중입니다." HorizontalScrollBarPolicy="Always"
                    VerticalScrollBarPolicy="Always" 
                    oncellclick="FpSpread_new_data_CellClick"   >
                    <CommandBar BackColor="Control" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight"
                        ButtonShadowColor="ControlDark">
                        <Background BackgroundImageUrl="SPREADCLIENTPATH:/img/cbbg.gif" />
                    </CommandBar>
                    <Pager Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False"
                        Font-Underline="False" PageCount="1000" />
                    <HierBar Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False"
                        Font-Underline="False" />
                    <Sheets>
                        <FarPoint:SheetView DataSourceID="(없음)" DesignString="&lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;&lt;Sheet&gt;&lt;Data&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;22&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.String&quot;&gt;선택&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;3&quot; columns=&quot;22&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;False&lt;/AutoGenerateColumns&gt;&lt;DataKeyField class=&quot;System.String[]&quot; assembly=&quot;mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089&quot; encoded=&quot;true&quot;&gt;AAEAAAD/////AQAAAAAAAAARAQAAAAAAAAAL&lt;/DataKeyField&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;SheetName&gt;Sheet1&lt;/SheetName&gt;&lt;RowIndexes class=&quot;System.Collections.ArrayList&quot; assembly=&quot;mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089&quot; encoded=&quot;true&quot;&gt;AAEAAAD/////AQAAAAAAAAAEAQAAABxTeXN0ZW0uQ29sbGVjdGlvbnMuQXJyYXlMaXN0AwAAAAZfaXRlbXMFX3NpemUIX3ZlcnNpb24FAAAICAkCAAAAAwAAACICAAAQAgAAAAQAAAAICAAAAAAICAEAAAAICAIAAAAKCw==&lt;/RowIndexes&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.String&quot; whitespace=&quot;&quot; /&gt;&lt;/Cell&gt;&lt;Cell row=&quot;1&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.String&quot; whitespace=&quot;&quot; /&gt;&lt;/Cell&gt;&lt;Cell row=&quot;2&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.String&quot; whitespace=&quot;&quot; /&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;1&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;22&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;/ColumnFooter&gt;&lt;/Data&gt;&lt;Presentation&gt;&lt;ActiveSkin class=&quot;FarPoint.Web.Spread.SheetSkin&quot;&gt;&lt;Name&gt;Classic&lt;/Name&gt;&lt;BackColor&gt;Empty&lt;/BackColor&gt;&lt;CellBackColor&gt;Empty&lt;/CellBackColor&gt;&lt;CellForeColor&gt;Empty&lt;/CellForeColor&gt;&lt;CellSpacing&gt;0&lt;/CellSpacing&gt;&lt;GridLines&gt;Both&lt;/GridLines&gt;&lt;GridLineColor&gt;LightGray&lt;/GridLineColor&gt;&lt;HeaderBackColor&gt;Control&lt;/HeaderBackColor&gt;&lt;HeaderForeColor&gt;Empty&lt;/HeaderForeColor&gt;&lt;FlatColumnHeader&gt;False&lt;/FlatColumnHeader&gt;&lt;FooterBackColor&gt;Empty&lt;/FooterBackColor&gt;&lt;FooterForeColor&gt;Empty&lt;/FooterForeColor&gt;&lt;FlatColumnFooter&gt;False&lt;/FlatColumnFooter&gt;&lt;FlatRowHeader&gt;False&lt;/FlatRowHeader&gt;&lt;HeaderFontBold&gt;False&lt;/HeaderFontBold&gt;&lt;FooterFontBold&gt;False&lt;/FooterFontBold&gt;&lt;SelectionBackColor&gt;LightBlue&lt;/SelectionBackColor&gt;&lt;SelectionForeColor&gt;Empty&lt;/SelectionForeColor&gt;&lt;EvenRowBackColor&gt;Empty&lt;/EvenRowBackColor&gt;&lt;OddRowBackColor&gt;Empty&lt;/OddRowBackColor&gt;&lt;ShowColumnHeader&gt;True&lt;/ShowColumnHeader&gt;&lt;ShowColumnFooter&gt;False&lt;/ShowColumnFooter&gt;&lt;ShowRowHeader&gt;True&lt;/ShowRowHeader&gt;&lt;ColumnHeaderBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/ColumnHeaderBackground&gt;&lt;SheetCornerBackground class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/SheetCornerBackground&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;/ActiveSkin&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;AxisModels&gt;&lt;Column class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;22&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;Item index=&quot;0&quot;&gt;&lt;Size&gt;36&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/Column&gt;&lt;RowHeaderColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;40&quot; orientation=&quot;Horizontal&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;Size&gt;40&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/RowHeaderColumn&gt;&lt;ColumnHeaderRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnHeaderRow&gt;&lt;ColumnFooterRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; defaultSize=&quot;22&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;Size&gt;22&lt;/Size&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/ColumnFooterRow&gt;&lt;/AxisModels&gt;&lt;StyleModels&gt;&lt;RowHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;RowHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/RowHeader&gt;&lt;ColumnHeader class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;22&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnHeaderDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ColumnStyles&gt;&lt;ColumnStyle Index=&quot;0&quot;&gt;&lt;CellType class=&quot;FarPoint.Web.Spread.CheckBoxCellType&quot; /&gt;&lt;/ColumnStyle&gt;&lt;/ColumnStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnHeader&gt;&lt;DataArea class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;3&quot; Columns=&quot;22&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;DataAreaDefault&quot;&gt;&lt;HorizontalAlign&gt;Center&lt;/HorizontalAlign&gt;&lt;/DefaultStyle&gt;&lt;ColumnStyles&gt;&lt;ColumnStyle Index=&quot;0&quot;&gt;&lt;CellType class=&quot;FarPoint.Web.Spread.CheckBoxCellType&quot; /&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;1&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;2&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;3&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;4&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;5&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;6&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;7&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;8&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;9&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;10&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;11&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;12&quot;&gt;&lt;Locked&gt;True&lt;/Locked&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;15&quot;&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;18&quot;&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;ColumnStyle Index=&quot;21&quot;&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/ColumnStyle&gt;&lt;/ColumnStyles&gt;&lt;ConditionalFormatCollections /&gt;&lt;/DataArea&gt;&lt;SheetCorner class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;1&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/SheetCorner&gt;&lt;ColumnFooter class=&quot;FarPoint.Web.Spread.Model.DefaultSheetStyleModel&quot; Rows=&quot;1&quot; Columns=&quot;22&quot;&gt;&lt;AltRowCount&gt;2&lt;/AltRowCount&gt;&lt;DefaultStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;ColumnFooterDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;/DefaultStyle&gt;&lt;ConditionalFormatCollections /&gt;&lt;/ColumnFooter&gt;&lt;/StyleModels&gt;&lt;MessageRowStyle class=&quot;FarPoint.Web.Spread.Appearance&quot;&gt;&lt;BackColor&gt;LightYellow&lt;/BackColor&gt;&lt;ForeColor&gt;Red&lt;/ForeColor&gt;&lt;/MessageRowStyle&gt;&lt;SheetCornerStyle class=&quot;FarPoint.Web.Spread.NamedStyle&quot; Parent=&quot;CornerDefault&quot;&gt;&lt;BackColor&gt;Control&lt;/BackColor&gt;&lt;Border class=&quot;FarPoint.Web.Spread.Border&quot; Size=&quot;1&quot; Style=&quot;Solid&quot;&gt;&lt;Bottom Color=&quot;ControlDark&quot; /&gt;&lt;Left Size=&quot;0&quot; /&gt;&lt;Right Color=&quot;ControlDark&quot; /&gt;&lt;Top Size=&quot;0&quot; /&gt;&lt;/Border&gt;&lt;Background class=&quot;FarPoint.Web.Spread.Background&quot;&gt;&lt;BackgroundImageUrl&gt;SPREADCLIENTPATH:/img/chm.png&lt;/BackgroundImageUrl&gt;&lt;/Background&gt;&lt;/SheetCornerStyle&gt;&lt;AllowLoadOnDemand&gt;false&lt;/AllowLoadOnDemand&gt;&lt;LoadRowIncrement &gt;10&lt;/LoadRowIncrement &gt;&lt;LoadInitRowCount &gt;30&lt;/LoadInitRowCount &gt;&lt;AllowVirtualScrollPaging&gt;false&lt;/AllowVirtualScrollPaging&gt;&lt;TopRow&gt;0&lt;/TopRow&gt;&lt;PreviewRowStyle class=&quot;FarPoint.Web.Spread.PreviewRowInfo&quot; /&gt;&lt;/Presentation&gt;&lt;Settings&gt;&lt;Name&gt;Sheet1&lt;/Name&gt;&lt;Categories&gt;&lt;Appearance&gt;&lt;GridLineColor&gt;#d0d7e5&lt;/GridLineColor&gt;&lt;HeaderGrayAreaColor&gt;Control&lt;/HeaderGrayAreaColor&gt;&lt;SelectionBackColor&gt;#eaecf5&lt;/SelectionBackColor&gt;&lt;SelectionBorder class=&quot;FarPoint.Web.Spread.Border&quot; /&gt;&lt;/Appearance&gt;&lt;Behavior&gt;&lt;EditTemplateColumnCount&gt;2&lt;/EditTemplateColumnCount&gt;&lt;PageSize&gt;200&lt;/PageSize&gt;&lt;GroupBarText&gt;Drag a column to group by that column.&lt;/GroupBarText&gt;&lt;/Behavior&gt;&lt;Layout&gt;&lt;RowHeaderColumnCount&gt;1&lt;/RowHeaderColumnCount&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnCount&gt;22&lt;/ColumnCount&gt;&lt;/Layout&gt;&lt;/Categories&gt;&lt;ActiveRow&gt;0&lt;/ActiveRow&gt;&lt;ActiveColumn&gt;21&lt;/ActiveColumn&gt;&lt;ColumnHeaderRowCount&gt;1&lt;/ColumnHeaderRowCount&gt;&lt;ColumnFooterRowCount&gt;1&lt;/ColumnFooterRowCount&gt;&lt;PrintInfo&gt;&lt;Header /&gt;&lt;Footer /&gt;&lt;ZoomFactor&gt;0&lt;/ZoomFactor&gt;&lt;FirstPageNumber&gt;1&lt;/FirstPageNumber&gt;&lt;Orientation&gt;Auto&lt;/Orientation&gt;&lt;PrintType&gt;All&lt;/PrintType&gt;&lt;PageOrder&gt;Auto&lt;/PageOrder&gt;&lt;BestFitCols&gt;False&lt;/BestFitCols&gt;&lt;BestFitRows&gt;False&lt;/BestFitRows&gt;&lt;PageStart&gt;-1&lt;/PageStart&gt;&lt;PageEnd&gt;-1&lt;/PageEnd&gt;&lt;ColStart&gt;-1&lt;/ColStart&gt;&lt;ColEnd&gt;-1&lt;/ColEnd&gt;&lt;RowStart&gt;-1&lt;/RowStart&gt;&lt;RowEnd&gt;-1&lt;/RowEnd&gt;&lt;ShowBorder&gt;True&lt;/ShowBorder&gt;&lt;ShowGrid&gt;True&lt;/ShowGrid&gt;&lt;ShowColor&gt;True&lt;/ShowColor&gt;&lt;ShowColumnHeader&gt;Inherit&lt;/ShowColumnHeader&gt;&lt;ShowRowHeader&gt;Inherit&lt;/ShowRowHeader&gt;&lt;ShowColumnFooter&gt;Inherit&lt;/ShowColumnFooter&gt;&lt;ShowColumnFooterEachPage&gt;True&lt;/ShowColumnFooterEachPage&gt;&lt;ShowTitle&gt;True&lt;/ShowTitle&gt;&lt;ShowSubtitle&gt;True&lt;/ShowSubtitle&gt;&lt;UseMax&gt;True&lt;/UseMax&gt;&lt;UseSmartPrint&gt;False&lt;/UseSmartPrint&gt;&lt;Opacity&gt;255&lt;/Opacity&gt;&lt;PrintNotes&gt;None&lt;/PrintNotes&gt;&lt;Centering&gt;None&lt;/Centering&gt;&lt;RepeatColStart&gt;-1&lt;/RepeatColStart&gt;&lt;RepeatColEnd&gt;-1&lt;/RepeatColEnd&gt;&lt;RepeatRowStart&gt;-1&lt;/RepeatRowStart&gt;&lt;RepeatRowEnd&gt;-1&lt;/RepeatRowEnd&gt;&lt;SmartPrintPagesTall&gt;1&lt;/SmartPrintPagesTall&gt;&lt;SmartPrintPagesWide&gt;1&lt;/SmartPrintPagesWide&gt;&lt;HeaderHeight&gt;-1&lt;/HeaderHeight&gt;&lt;FooterHeight&gt;-1&lt;/FooterHeight&gt;&lt;/PrintInfo&gt;&lt;TitleInfo class=&quot;FarPoint.Web.Spread.TitleInfo&quot;&gt;&lt;Style class=&quot;FarPoint.Web.Spread.StyleInfo&quot;&gt;&lt;BackColor&gt;#e7eff7&lt;/BackColor&gt;&lt;HorizontalAlign&gt;Right&lt;/HorizontalAlign&gt;&lt;/Style&gt;&lt;/TitleInfo&gt;&lt;LayoutTemplate class=&quot;FarPoint.Web.Spread.LayoutTemplate&quot;&gt;&lt;Layout&gt;&lt;ColumnCount&gt;4&lt;/ColumnCount&gt;&lt;RowCount&gt;1&lt;/RowCount&gt;&lt;/Layout&gt;&lt;Data&gt;&lt;LayoutData class=&quot;FarPoint.Web.Spread.Model.DefaultSheetDataModel&quot; rows=&quot;1&quot; columns=&quot;4&quot;&gt;&lt;AutoCalculation&gt;True&lt;/AutoCalculation&gt;&lt;AutoGenerateColumns&gt;True&lt;/AutoGenerateColumns&gt;&lt;ReferenceStyle&gt;A1&lt;/ReferenceStyle&gt;&lt;Iteration&gt;False&lt;/Iteration&gt;&lt;MaximumIterations&gt;1&lt;/MaximumIterations&gt;&lt;MaximumChange&gt;0.001&lt;/MaximumChange&gt;&lt;Cells&gt;&lt;Cell row=&quot;0&quot; column=&quot;0&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;0&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;1&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;1&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;2&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;2&lt;/Data&gt;&lt;/Cell&gt;&lt;Cell row=&quot;0&quot; column=&quot;3&quot;&gt;&lt;Data type=&quot;System.Int32&quot;&gt;3&lt;/Data&gt;&lt;/Cell&gt;&lt;/Cells&gt;&lt;/LayoutData&gt;&lt;/Data&gt;&lt;AxisModels&gt;&lt;LayoutColumn class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Horizontal&quot; count=&quot;4&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot;&gt;&lt;SortIndicator&gt;Ascending&lt;/SortIndicator&gt;&lt;/Item&gt;&lt;/Items&gt;&lt;/LayoutColumn&gt;&lt;LayoutRow class=&quot;FarPoint.Web.Spread.Model.DefaultSheetAxisModel&quot; orientation=&quot;Vertical&quot; count=&quot;1&quot;&gt;&lt;Items&gt;&lt;Item index=&quot;-1&quot; /&gt;&lt;/Items&gt;&lt;/LayoutRow&gt;&lt;/AxisModels&gt;&lt;/LayoutTemplate&gt;&lt;LayoutMode&gt;CellLayoutMode&lt;/LayoutMode&gt;&lt;CurrentPageIndex type=&quot;System.Int32&quot;&gt;0&lt;/CurrentPageIndex&gt;&lt;/Settings&gt;&lt;/Sheet&gt;"
                            SheetName="Sheet1" PageSize="200" AutoGenerateColumns="False">
                        </FarPoint:SheetView>
                    </Sheets>
                    <TitleInfo BackColor="#E7EFF7" Font-Size="X-Large" ForeColor="" HorizontalAlign="Center"
                        VerticalAlign="NotSet" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                        Font-Strikeout="False" Font-Underline="False">
                    </TitleInfo>
                </FarPoint:FpSpread>

                 
            </ContentTemplate>
        </asp:UpdatePanel>
    </asp:Panel>
    <asp:Panel ID="Panel_Report" runat="server">
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="" 
            SizeToReportContent="True" Width="">
        </rsweb:ReportViewer>
    </asp:Panel>
    <asp:HiddenField ID="hf_tb_ship_fr_cust_cd" runat="server" />
    <asp:HiddenField ID="hf_tb_bill_to_cust_cd" runat="server" />
    <asp:HiddenField ID="hf_tb_ship_to_cust_cd" runat="server" />
    <asp:HiddenField ID="hf_tb_bank_cd" runat="server" />
    <asp:HiddenField ID="HiddenField1_5" runat="server" />
    <asp:HiddenField ID="HiddenField1_6" runat="server" />
    <asp:HiddenField ID="HiddenField2_1" runat="server" />
    <asp:HiddenField ID="HiddenField2_2" runat="server" />
    <asp:HiddenField ID="HiddenField2_3" runat="server" />
    <asp:HiddenField ID="HiddenField2_4" runat="server" />
    <asp:HiddenField ID="HiddenField2_5" runat="server" />
    <asp:HiddenField ID="HiddenField2_6" runat="server" />
    <asp:HiddenField ID="HiddenField3_1" runat="server" />
    <asp:HiddenField ID="HiddenField3_2" runat="server" />
    <asp:HiddenField ID="HiddenField3_3" runat="server" />
    <asp:HiddenField ID="HiddenField3_4" runat="server" />
    <asp:HiddenField ID="HiddenField3_5" runat="server" />
    <asp:HiddenField ID="HiddenField3_6" runat="server" />
    </form>
</body>
</html>
