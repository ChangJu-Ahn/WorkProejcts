<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_AA1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_AA1001.AM_AA1001" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %><%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head runat="server"><meta http-equiv="Content-Type" content="text/html; charset=utf-8"/><title></title><style type="text/css">
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
         .style1
        {
            width: 118px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }  
        
        
        .table {
            border-collapse: collapse;
            border: 1px solid black;
        }
        
        .tbl_list{border:1px solid #e8e9ea; border-collapse:collapse; background-color:#fff; font-size: 10pt;}

        .tbl_list th{background:#113971; color:#fff;  font-weight:bold; text-align:center;}

        .tbl_list td{color:#787878;border-left:1px solid #e8e9ea; text-align:center;}

        .tbl_list th, .tbl_list td{padding:10px 10px 10px 01px; line-height:1.5em; border:1px solid #e8e9ea; font-size: 10pt;}

        .tbl_list .align_l{padding-left:15px }

        .tbl_list .subject{ height:20px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; line-height:1.5em; text-align:left;}

        .tbl_list .cont_text{padding:15px; text-align:left;}
             
               .auto-style6 {
            width: 12px;
            text-align: center;
        }
             
             
             
               .auto-style9 {
            width: 12px;
            height: 41px;
        }
                             
               .style2 {
            font-weight: 300;
            text-align: center;
        }
             
                 .auto-style11 {
            width: 300px;
        }

               .auto-style12 {
            width: 218px;
        }

               .auto-style14 {
            width: 193px;
            height: 11px;
        }
        .auto-style15 {
            width: 193px;
        }

               .auto-style16 {
            width: 130px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            
        }
       
        }

               </style></head><body><form id="form1" runat="server">
     <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager> 
    <div>
    
      <table>
            <tr>
                <td>
       <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
       </td>
                <td style="width: 100%;">
       <asp:Label ID="Label1" runat="server" Text="일일운용자금 등록(NEPES)" CssClass="title" Width="100%"></asp:Label>
      </td>
            </tr>
        </table>
        <table style="border: thin solid #000080; height: 31px;">
            <td class="style1">
                <asp:Label ID="Label314" runat="server" Text="조회구분" BackColor="#99CCFF" Font-Bold="True"
                    Style="text-align: center; font-size: small"></asp:Label>
            </td>
            <td class="auto-style11">
                <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small" RepeatDirection="Horizontal"
                    OnSelectedIndexChanged="rbl_view_type_SelectedIndexChanged" AutoPostBack="True"
                    Width="417px" Style="margin-left: 0px; font-weight: 700;" BackColor="White" Height="16px">
                    <asp:ListItem Value="A">일일실적등록</asp:ListItem>
                    <asp:ListItem Value="B">경상/계열사 수입등록</asp:ListItem>
                    
                </asp:RadioButtonList>
            </td>
        </table>
       <asp:Panel ID="Pane_excel" runat="server" Visible="False">
       <table style="border: thin solid #000080; height: 31px;">
                
           <td class="style1">  
                     <asp:Label ID="lb_yyyy" runat="server" Text="년도" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                    
            </td >  
                <td class="style3">
                    <asp:TextBox ID="txt_yyyy" runat="server" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
                        
            </td>
            <td class="style1">  
                     <asp:Label ID="lb_mm" runat="server" Text="월" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
             <td class="style3">
                    <asp:TextBox ID="txt_mm" runat="server" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
                        
            </td>
      <td class="style1">  
                     <asp:Label ID="lb_dd" runat="server" Text="일" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
             <td class="style3">
                    <asp:TextBox ID="txt_dd" runat="server" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
                        
            </td>

             <td class="auto-style16">  
                     <asp:Label ID="Label315" runat="server" Text="EXCEL선택" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
                       
            </td >
       
           <td class="auto-style12">
                        <asp:FileUpload ID="FileUpload1" runat="server" BackColor="#FFFFCC" Width="224px" />
                       
                    </td>
               <td>      <asp:DropDownList ID="ddlSheets" runat="server" AutoPostBack="True" BackColor="#FFFFCC">
            </asp:DropDownList>             </td >
            <td class="style56">
                                
                                <asp:Button ID="btnUpload" runat="server" Text="Upload"
                            ToolTip="엑셀자료를 가져와 화면에 보여준다." Width="87px" Height="21px" OnClick="btnUpload_Click" />
                    </td>

            <td class="style56">
                                
                                <asp:Button ID="btn_select" runat="server" Text="조회" 
                                    Width="100px" OnClick="btn_select_Click"   />
                    </td>
                        <td class="style56">
                                
                                <asp:Button ID="btn_save" runat="server" Text="저장" 
                                    Width="100px" OnClick="btn_save_Click"  />
                    </td>
                      
                       
    </table>
       </div>
         </asp:Panel>
         <asp:Panel ID="Panel_insert" runat="server" Visible="False">
       <table style="border: thin solid #000080; height: 31px;">
                
           <tr>
               <td class="style1">
                   <asp:Label ID="Label48" runat="server" BackColor="#99CCFF" Font-Bold="True" style="text-align: center; font-size: small" Text="년도"></asp:Label>
               </td>
               <td class="style3">
                   <asp:TextBox ID="txt_yyyy1" runat="server" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
               </td>
               <td class="style1">
                   <asp:Label ID="Label49" runat="server" BackColor="#99CCFF" Font-Bold="True" style="text-align: center; font-size: small" Text="월"></asp:Label>
               </td>
               <td class="style3">
                   <asp:TextBox ID="txt_mm1" runat="server" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
               </td>
               
               <td class="style3">
                   <asp:Button ID="btn_save2" runat="server" Text="저장" Width="100px" OnClick="btn_save2_Click" />
               </td>
                            
               <td class="style56">
                   <asp:Button ID="btn_select2" runat="server" Text="조회" Width="100px" OnClick="btn_select2_Click" />
               </td>
               <td class="style56">
                   <asp:Button ID="btn_delete2" runat="server"  Text="삭제" Width="100px" OnClick="btn_delete2_Click" />
               </td>
           </tr>
                      
                       
    </table>
      
         </asp:Panel>  </div>
       <asp:Panel ID="Panel1" runat="server">
                            <td class="style17">
                                <asp:HiddenField ID="HiddenField_filePath" runat="server" />
                                <asp:HiddenField ID="HiddenField_extension" runat="server" />
                                <asp:HiddenField ID="HiddenField_fileName" runat="server" />
                            </td>
                        </asp:Panel>  
        <asp:Panel ID="Panel_regist_excel_grid" runat="server" Visible="False" >
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
             <br />
             </asp:Panel>
        <asp:Panel ID="Panel_select_excel_qty_grid" runat="server">
     
              <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="100%" 
            AsyncRendering="False" Height="600px" SizeToReportContent="True">
                </rsweb:ReportViewer>
                

                
            </asp:Panel>
      <asp:Panel ID="Panel_amt" runat="server" Visible="False">
        <div id="div_down_spread">
           
            <table  class="tbl_list">
                <tr>
                    <th class="auto-style6"><asp:Label ID="lb_hd_no" runat="server" Text="NO" Width="50px"></asp:Label></th>
                    <th class="auto-style15"><asp:Label ID="lb_hd_list" runat="server" Text="항목" Width="100px"></asp:Label></th>
                    <th class="auto-style17"><asp:Label ID="lb_hd_amt" runat="server" Text="금액" Width="100px"></asp:Label></th>
                   
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="lb_no1" runat="server" Text="1." Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="lb_list1" runat="server" Text="경상수입" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17">&nbsp;</td>
                  
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="lb_no2" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="lb_list2" runat="server" Text="NEPES_LED" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt1" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="lb_no3" runat="server" Text=" " Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="lb_list3" runat="server" Text="Rigmah" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt2" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
           
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label2" runat="server" Text=" " Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label3" runat="server" Text="기타" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt3" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
               <tr>
                    <td><asp:Label ID="Label4" runat="server" Text="2." Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label5" runat="server" Text="계열사대여" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17">&nbsp;</td>
                  
                </tr>
                <tr>
                    <td class="auto-style9"><asp:Label ID="Label6" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label7" runat="server" Text="외환은행" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt4" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label8" runat="server" Text=" " Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label9" runat="server" Text="수출입은행" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt5" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
               <tr>
                    <td class="auto-style6"><asp:Label ID="Label10" runat="server" Text=" " Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label11" runat="server" Text="증자" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt6" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label12" runat="server" Text=" " Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label13" runat="server" Text="씨티은행" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt7" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label14" runat="server" Text=" " Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label15" runat="server" Text="하나은행" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt8" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label16" runat="server" Text=" " Width="50px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label17" runat="server" Text="산업은행" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt9" runat="server" style="text-align: center"></asp:TextBox></td>
                  
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label18" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label19" runat="server" Text="영엽외수입" Width="120px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt10" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
               <tr>
                    <td class="auto-style6"><asp:Label ID="Label20" runat="server" Text="3" Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label21" runat="server" Text="고정자산의 매각(SEMI)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt11" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label22" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label23" runat="server" Text="고정자산의 매각(NM_BU)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt12" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
              
              <tr>
                    <td class="auto-style6"><asp:Label ID="Label24" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label25" runat="server" Text="고정자산의 매각(SOLVR)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt13" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label26" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label27" runat="server" Text="고정자산의 매각(공통비)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt14" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label30" runat="server" Text="4" Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label31" runat="server" Text="대여금 회수(SEMI)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt16" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                  <tr>
                    <td class="auto-style6"><asp:Label ID="Label32" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label33" runat="server" Text="대여금 회수(NM BU)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt17" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label34" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label35" runat="server" Text="대여금 회수(SOLVR)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt18" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label36" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label37" runat="server" Text="대여금 회수(공통비)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt19" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                                <tr>
                    <td class="auto-style6"><asp:Label ID="Label40" runat="server" Text="5" Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label41" runat="server" Text="장기금융상품의 회수(SEMI)" Width="170px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt21" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6"><asp:Label ID="Label42" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label43" runat="server" Text="장기금융상품의 회수(NM BU)" Width="200px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt22" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label28" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label29" runat="server" Text="장기금융상품의 회수(SOLVR)" Width="200px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt23" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label38" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label39" runat="server" Text="장기금융상품의 회수(공통비)" Width="200px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt24" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label44" runat="server" Text="6" Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label45" runat="server" Text="투자유가증권/출자금 회수(SEMI)" Width="200px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt26" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                </tr>
                
                 <td class="auto-style6">
                     <asp:Label ID="Label46" runat="server" CssClass="style2" Height="16px" Text=" " Width="50px"></asp:Label>
                </td>
                <td class="auto-style14">
                    <asp:Label ID="Label47" runat="server" CssClass="style2" Text="투자유가증권/출자금 회수(NM BU)" Width="210px"></asp:Label>
                </td>
                <td class="auto-style17">
                    <asp:TextBox ID="tb_amt27" runat="server" style="text-align: center"></asp:TextBox>
                </td>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label50" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label51" runat="server" Text="투자유가증권/출자금 회수(SOLVR)" Width="210px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt28" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                 <tr>
                    <td class="auto-style6"><asp:Label ID="Label52" runat="server" Text=" " Width="50px" CssClass="style2" Height="16px"></asp:Label></td>
                    <td class="auto-style14"><asp:Label ID="Label53" runat="server" Text="투자유가증권/출자금 회수(공통비)" Width="210px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style17"><asp:TextBox ID="tb_amt29" runat="server" style="text-align: center"></asp:TextBox>
                    </td>
                </tr>
                </tr>
                 </table>
        </div>
        
       </asp:Panel>

   
    </form>
</body>
</html>
