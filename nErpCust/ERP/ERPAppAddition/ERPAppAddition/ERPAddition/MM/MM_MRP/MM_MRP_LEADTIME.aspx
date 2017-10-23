<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_MRP_LEADTIME.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_MRP.MM_MRP_LEADTIME" %>
<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>MRP Leadtime 관리</title>
</head>
      <style type="text/css">
        .style12 {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
        }

        .style1 {
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            table-layout: fixed;
        }

        .spread {
            width: 120px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
        }

        .title {
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

        .default_font_size {
            font-family: 굴림체;
            font-size: 10pt;
            text-align: center;
        }

        .default_font_background {
            font-family: 굴림체;
            font-size: 10pt;
        }

          .auto-style2 {
              width: 320px;
          }
          
          .auto-style7 {
              width: 120px;
              background-color: #99CCFF;
              font-weight: bold;
              font-family: 굴림체;
              text-align: center;
              font-size: smaller;
              table-layout: fixed;
              height: 23px;
          }
          .auto-style8 {
              width: 320px;
              height: 23px;
          }

          .auto-style9 {
              width: 394px;
          }

          .auto-style10 {
              width: 120px;
              background-color: #99CCFF;
              font-weight: bold;
              font-family: 굴림체;
              text-align: center;
              font-size: smaller;
              table-layout: fixed;
              height: 25px;
          }
          .auto-style11 {
              width: 394px;
              height: 25px;
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
                        <asp:Label ID="Label3" runat="server" Text="MRP LeadTime 관리" CssClass="title" Width="100%"></asp:Label>
                    </td>
                </tr>
            </table>
             <asp:Panel ID="Panel_view" runat="server">          
<table  style="border: thin solid #000080">
            <tr>
                              
                 <td class="style12">
                <asp:Label ID="Label5" runat="server" Text="품&nbsp;&nbsp;&nbsp;&nbsp;목"></asp:Label>
                </td>
                <td class="auto-style2">
                 <asp:TextBox ID="tb_item_cd_0" runat="server"></asp:TextBox>
                    <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="0">
                    </asp:ScriptManager>
                    </td>
                     </tr> 
            <tr>
            <td class="auto-style7">
                제 조 사</td><td class="auto-style8">
                    <asp:TextBox ID="tb_bizpartner_0" runat="server" Width="404px"></asp:TextBox>
                </td>
                <td>
                    <asp:Button ID="Button1" runat="server" Text="찾기" />
                </td>
                <td>
                    <asp:Button ID="bt_retrieve" runat="server" style="font-weight: 700; background-color: #FFFF99" Text="조회" Width="100px" OnClick="bt_retrieve_Click" />
                </td>
                <td>
                    <asp:Button ID="bt_save" runat="server" OnClick="bt_save_Click" style="font-weight: 700; background-color: #FFFF99" Text="저장" Width="100px" />
                </td>
                <td>
                    <asp:Button ID="bt_change" runat="server" OnClick="bt_change_Click" style="font-weight: 700; background-color: #FFFF99" Text="수정" Width="100px" />
                </td>
                <td>
                    <asp:Button ID="bt_del" runat="server" OnClick="bt_del_Click" style="font-weight: 700; background-color: #FFFF99" Text="삭제" Width="100px" />
                </td>
                <td>
                    <asp:Button ID="bt_refresh" runat="server" style="font-weight: 700; background-color: #FFFF99" Text="재작성" Width="100px" OnClick="bt_refresh_Click" />
                </td>
             </tr>
           
           </table>
                       </asp:Panel>   
             
     <table  style="border: thin solid #000080">
            <tr>
                 <td class="auto-style10">
                <asp:Label ID="Label4" runat="server" Text="품&nbsp;&nbsp;&nbsp;&nbsp;목"></asp:Label>
                </td>
                <td class="auto-style11">
                 <asp:TextBox ID="tb_item_cd" runat="server" MaxLength="30"></asp:TextBox>
                    <asp:Button ID="bt_item_cd" runat="server"  Text=".." 
                        style="width: 22px" />
                    <asp:TextBox ID="tb_item_nm" runat="server"></asp:TextBox></td>
                     </tr> <tr>  
            <td class="style12">
                LeadTime</td><td class="auto-style9">
                    <asp:TextBox ID="tb_leadtime" runat="server"></asp:TextBox>
                </td>
            </tr>
              <tr>
            <td class="style12">
                MOQ</td><td class="auto-style9">
                      <asp:TextBox ID="tb_moq" runat="server" MaxLength="10"></asp:TextBox>
                  </td>
             </tr>
            <tr>
            <td class="style12">
                제 조 사</td><td class="auto-style9">
                    <asp:TextBox ID="tb_bizpartner" runat="server" Width="404px" MaxLength="100"></asp:TextBox>
                </td>
             </tr>           
           </table>
             
              
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
                <asp:AsyncPostBackTrigger ControlID="bt_retrieve">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel> 
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                <br /><br /><br /><br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
       <cc1:ModalPopupExtender ID="ModalProgress" runat="server" 
            PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
    </div>
      <div>
    </div>
      <cc1:ModalPopupExtender ID="ModalPopupExtender1" runat="server" BackgroundCssClass="modalBackground2"
            PopupControlID="Panel1" TargetControlID="bt_item_cd" >
        </cc1:ModalPopupExtender>
      
        <div class="div_center">
          
            
        <asp:Panel ID="Panel1" runat="server" BorderStyle="Solid" Height="500px" Width="600px"
            BackColor="#CCCCFF">
            <table>
                <tr>
                    <td>
                        <asp:Label ID="Label1" runat="server" Text="*제품 : " ForeColor="Black"></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tb_pop_item_cd" runat="server" Width="100px"></asp:TextBox>
                    </td>
                    <td>
                        <asp:TextBox ID="tb_pop_item_nm" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label2" runat="server" Text="*계정 : " ForeColor="Black"></asp:Label>
                    </td>
                    <td>
                        <asp:DropDownList ID="dl_pop_item_acct" runat="server" Width="150px" DataSourceID="SqlDataSource2_item_acct"
                            DataTextField="MINOR_NM" DataValueField="MINOR_CD">
                        </asp:DropDownList>
                        <asp:SqlDataSource ID="SqlDataSource2_item_acct" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                            SelectCommand="select MINOR_CD, MINOR_NM from B_MINOR where MAJOR_CD = 'p1001'">
                        </asp:SqlDataSource>
                    </td>
                </tr>
            </table>
            <table>
                    <tr>
                        <td>
                            <asp:Button ID="pop_bt_retrive" runat="server" Text="조회" OnClick="bt_retrive_Click"
                                Width="100px" />
                        </td>
                        <td style="width: 400px; text-align: right;">
                            <asp:Button ID="bt_cancel" runat="server" Text="취소" Width="100px" OnClick="bt_cancel_Click" />
                        </td>
                        <td style="width: 100px; text-align: right;">
                            <asp:Button ID="btn_pop_ok" runat="server" Text="OK" Width="100px" 
                                OnClick="btn_pop_ok_Click" />
                        </td>
                    </tr>
                </table>
            <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional" 
                ClientIDMode="AutoID">
            <ContentTemplate>
                
            <asp:GridView ID="pop_gridview1" runat="server" AllowPaging="True" 
                AutoGenerateColumns="False" AutoGenerateSelectButton="True" CellPadding="4" 
                ForeColor="#333333" GridLines="None" 
                onpageindexchanging="pop_gridview1_PageIndexChanging" 
                    onselectedindexchanged="pop_gridview1_SelectedIndexChanged" PageSize="15" 
                    Width="600px" Font-Size="Small">

                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:BoundField DataField="item_cd" HeaderText="품목코드" />
                    <asp:BoundField DataField="item_nm" HeaderText="품목명" />
                    <asp:BoundField DataField="spec" HeaderText="규격" />
                    <asp:BoundField DataField="basic_unit" HeaderText="기준단위" />
                    <asp:BoundField DataField="item_acct" HeaderText="품목계정" />
                </Columns>
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
             </ContentTemplate>               
            </asp:UpdatePanel>    
        </asp:Panel>    
            
        </div>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
            ConnectionString="<%$ ConnectionStrings:nepes %>" SelectCommand="SELECT b.item_cd, 
       b.item_nm, 
       b.spec, 
       b.basic_unit, 
       dbo.Ufn_getcodename('P1001', a.item_acct)   item_acct 
FROM   b_item_by_plant a, 
       b_item b, 
       b_item_acct_inf c 
WHERE  a.item_cd = b.item_cd 
       AND a.item_acct = c.item_acct 
       AND a.item_acct = @item_acct
       AND a.item_cd &gt;= @item_cd
       AND b.item_nm LIKE @item_nm
       AND b.item_nm &gt;= '' 
       AND ( b.item_class &gt;= '' 
             AND b.item_class &lt;= 'zzzzzzzzzzzz' 
              OR b.item_class IS NULL ) 
       AND c.item_acct_group &gt;= '' 
       AND c.item_acct_group &lt;= 'zz' 
       AND a.procur_type &gt;= '' 
       AND a.procur_type &lt;= 'zz' 
       AND a.prod_env &gt;= '' 
       AND a.prod_env &lt;= 'zz' 
       AND a.valid_to_dt &gt;= Getdate() 
       AND b.spec LIKE '%%' 
       AND a.tracking_flg LIKE '%' 
ORDER  BY a.item_cd, 
          b.item_nm ">
            <SelectParameters>
                <asp:ControlParameter ControlID="dl_pop_item_acct" Name="item_acct" 
                    PropertyName="SelectedValue" />

                <asp:ControlParameter ControlID="tb_pop_item_cd" DefaultValue="%" 
                    Name="item_cd" PropertyName="Text" />
                <asp:ControlParameter ControlID="tb_pop_item_nm" DefaultValue="%" 
                    Name="item_nm" PropertyName="Text" />
            </SelectParameters>
        </asp:SqlDataSource>
                
    </div>
    </form>
</body>
</html>

