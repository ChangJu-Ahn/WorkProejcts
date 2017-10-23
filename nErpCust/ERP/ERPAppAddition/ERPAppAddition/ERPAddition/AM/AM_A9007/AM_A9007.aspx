
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_A9007.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A9007.AM_A9007" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>현장재고 등록</title>
    <style type="text/css">

        
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
        .auto-style1 {
            width: 126px;
            font-family: 굴림체;
            font-size: smaller;
            text-align: center;
            background-color: Silver;
            height: 22px;
            font-weight: 700;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
 
   <table>
            <tr>
                <td>
       <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
       </td>
                <td style="width: 100%;">
       <asp:Label ID="Label1" runat="server" Text="현장재고 등록 및 재고금액 조회" CssClass="title" Width="100%"></asp:Label>
      </td>
            </tr>
        </table>
           </div>
        <table style="border: thin solid #000080; height: 31px;">
      
                 <td class="style2">  
                     <asp:Label ID="Label17" runat="server" Text="구분" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
                <td class="style3">
                <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small" 
                    RepeatDirection="Horizontal" 
                    onselectedindexchanged="rbl_view_type_SelectedIndexChanged" 
                    AutoPostBack="True" Width="983px" style="margin-left: 0px; font-weight: 700;" 
                    BackColor="White" Height="16px">
                    <asp:ListItem Value="A">재고 등록</asp:ListItem>
                    <asp:ListItem Value="B">네패스 재고금액 조회</asp:ListItem>
                    <asp:ListItem Value="C">음성 재고금액 조회</asp:ListItem>
                    <asp:ListItem Value="D">디스플레이 재고금액 조회</asp:ListItem>
                    <asp:ListItem Value="E">EM상품매출 조회</asp:ListItem>
                    <asp:ListItem Value="F">반도체 상품재고금액 조회</asp:ListItem>
                </asp:RadioButtonList>                
                        
            </td>
          
     
    </table> 
 <asp:Panel ID="panel_upload" runat="server" Visible="False" BorderStyle="Groove" 
            BorderColor="White" Width="99%">
            <table style="width: 99%">
                <tr >
                 
                     
                   <%--<td class="style58">--%> 
                   <%--     <asp:Label ID="Label13" runat="server" Font-Size="Small" 
                            Text="계획 버전" BackColor="Silver" style="font-weight: 700"></asp:Label></td>
                            <td class="style25">--%>
                        <%--<asp:DropDownList ID="ddl_plan_version" runat="server" BackColor="#FFFFCC" 
                             Width="90px" >
                            <asp:ListItem>-선택안함-</asp:ListItem>
                            <asp:ListItem Value="001">001(경영계획)</asp:ListItem>
                            <asp:ListItem Value="002">002(수정계획)</asp:ListItem>
                            <asp:ListItem Value="003">003(당월계획)</asp:ListItem>
                            <asp:ListItem Value="004">004(당월예산)</asp:ListItem>
                        </asp:DropDownList></td>--%>
                    <%-- <td class=style55>사업장 선택</td>
                        <td class="style53">--%>
                         
                 <%--      <asp:DropDownList ID="ddl_biz_cd" runat="server" 
                        DataSourceID="SqlDataSource1" DataTextField="biz_area_nm" 
                        DataValueField="biz_area_cd" Height="25px" Width="200px" AutoPostBack="True" style="background-color: #FFFFCC">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:nepes %>" SelectCommand="    SELECT BIZ_AREA_CD , BIZ_AREA_NM 
                                                                                             FROM B_BIZ_AREA 
                                                                                             union all SELECT '', '전체' ORDER BY BIZ_AREA_CD">
                      </asp:SqlDataSource>
                            </td>--%>

                        <td class=auto-style1>날짜선택<br /> (주별 날짜선택)</td>
                          <td class="style1">
                          <asp:DropDownList ID="txt_regist_date_yyyymm" runat="server" BackColor="Yellow" OnSelectedIndexChanged="txt_regist_date_yyyymm_SelectedIndexChanged" >
                          </asp:DropDownList>
                         </td>

                   <%--     <td class="style53">
                            <asp:TextBox ID="txt_regist_date_yyyymm" runat="server" BackColor="#FFFFCC" 
                                Height="16px" Width="75px"></asp:TextBox>
                            </td>--%>
                           
                            <td class=style55>Excel선택</td>
                            <td class="style54">
                            <asp:FileUpload ID="FileUpload1" runat="server" BackColor="#FFFFCC" />
                            &nbsp;<asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" 
                                    Text="Upload" ToolTip="엑셀자료를 가져와 화면에 보여준다." Width="87px" Height="21px" />
                            </td>
                            <td class=style57 style="border-style: none; font-weight: 700;">Sheet선택</td>
                            <td class="style56">
                            <asp:DropDownList ID="ddlSheets" runat="server" AutoPostBack="True" 
                                BackColor="#FFFFCC">
                            </asp:DropDownList>
                            <asp:Button ID="btnSave" runat="server" OnClick="btnSave_Click" 
                                OnClientClick="return confirm('당월데이타 존재시 삭제후 저장됩니다. 진행하시겠습니까?');" Text="네패스저장" 
                                    Width="80px" />
                              <asp:Button ID="Button1" runat="server" OnClick="btnemSave_Click" 
                                OnClientClick="return confirm('당월데이타 존재시 삭제후 저장됩니다. 진행하시겠습니까?');" Text="음성저장" 
                                    Width="80px" />
                                <asp:Button ID="Button2" runat="server" OnClick="btndsSave_Click" 
                                OnClientClick="return confirm('당월데이타 존재시 삭제후 저장됩니다. 진행하시겠습니까?');" Text="Display저장" 
                                    Width="80px" />
                                 <asp:Button ID="Button4" runat="server" OnClick="btnemgoodsSave_Click" 
                                OnClientClick="return confirm('당월데이타 존재시 삭제후 저장됩니다. 진행하시겠습니까?');" Text="EM상품저장" 
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
            <table style="width: 97%">
                <tr >
                  <%-- <td class="style58"> 
                        <asp:Label ID="Label5" runat="server" Font-Size="Small" 
                            Text="계획 버전" BackColor="Silver" style="font-weight: 700"></asp:Label></td>
                            <td class="style25">
                                <asp:DropDownList ID="ddl_select_version" runat="server" BackColor="#FFFFCC" 
                                    style="text-align: center">
                                   <asp:ListItem Value="0">-선택안함-</asp:ListItem>
                            <asp:ListItem Value="001">001(경영계획)</asp:ListItem>
                            <asp:ListItem Value="002">002(수정계획)</asp:ListItem>
                            <asp:ListItem Value="003">003(당월계획)</asp:ListItem>
                            <asp:ListItem Value="004">004(당월예산)</asp:ListItem>
                                </asp:DropDownList>
                    </td>--%>
                    <%-- <td class="style58"> 
                        <asp:Label ID="Label2" runat="server" Font-Size="Small" 
                            Text="사업장선택" BackColor="Silver" style="font-weight: 700"></asp:Label></td>
                            <td class="style25">
                                    <asp:DropDownList ID="ddl_select_biz" runat="server" 
                        DataSourceID="SqlDataSource1" DataTextField="biz_area_nm" 
                        DataValueField="biz_area_cd" Height="25px" Width="200px" AutoPostBack="True" style="background-color: #FFFFCC">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:nepes %>" SelectCommand="    SELECT BIZ_AREA_CD , BIZ_AREA_NM 
                                                                                             FROM B_BIZ_AREA 
                                                                                             union all SELECT '', '전체' ORDER BY BIZ_AREA_CD">
                      </asp:SqlDataSource>
                    </td>--%>
                        <td class=style55>날짜선택</td>

                          <td class="style1">
                          <asp:DropDownList ID="txt_select_date_yyyymm" runat="server" BackColor="Yellow" >
                          </asp:DropDownList>

                   <%--     <td class="style59">
                            <asp:TextBox ID="txt_select_date_yyyymm" runat="server" BackColor="#FFFFCC" 
                                Height="16px" Width="88px"></asp:TextBox>  
                              &nbsp;   
                            </td>
                           --%>
                           
                                <td class="style56">
                                
                                <asp:Button ID="btn_select" runat="server" Text="조회" 
                                    Width="100px" onclick="btn_select_Click" style="height: 21px" />

                                  
                                <asp:Button ID="btn_exec" runat="server" Text="실행" 
                                    Width="100px" onclick="btn_exec_Click" style="height: 21px" />

                               </td>


                    
                        </tr>
                      
                 
            </table>
             
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
        <asp:Panel ID="Panel_select_excel_qty_grid" runat="server">
              <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="100%" AsyncRendering="False" Height="600px" SizeToReportContent="True">
              </rsweb:ReportViewer>
        </asp:Panel>

                <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="btn_select">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
        <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" Height="75px" Width="179px" />
                <br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
         <br />
         </ProgressTemplate>
         </asp:UpdateProgress>
    </div>
    </form>
</body>
</html>
