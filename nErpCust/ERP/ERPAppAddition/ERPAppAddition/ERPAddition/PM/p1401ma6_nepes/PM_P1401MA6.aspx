<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PM_P1401MA6.aspx.cs" Inherits="ERPAppAddition.ERPAddition.PM.p1401ma6_nepes.PM_P1401MA6" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>생산기준정보조회</title>
   
    <style type="text/css">
        .style12
        {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            font-size:10pt;
            text-align: center;
        }
        .modalBackground
        {
            background-color: #CCCCFF;
            filter: alpha(opacity=40);
            opacity: 0.5;
        }
        .modalBackground2
        {
            background-color: Gray;
            filter: alpha(opacity=50);
            opacity: 0.5;
        }      
        
        .updateProgress
        {
           
            background-color:#ffffff;
            position: absolute;
            width :180px;
            height: 65px;
        }
        .ModalWindow
        {
            border: solid1px#c0c0c0;
            background: #f0f0f0;
            padding: 0px10px10px10px;
            position: absolute;
            top: -1000px;
        }
       
         .fixedheadercell
        {
            FONT-WEIGHT: bold; 
            FONT-SIZE: 10pt; 
            WIDTH: 200px; 
            COLOR: white; 
            FONT-FAMILY: Arial; 
            BACKGROUND-COLOR: darkblue;
        }

        .fixedheadertable
        {
            left: 0px;
            position: relative;
            top: 0px;
            padding-right: 2px;
            padding-left: 2px;
            padding-bottom: 2px;
            padding-top: 2px;
        }

        .gridcell
        {
            WIDTH: 200px;
        }
        
        .div_center
        {
            width: 500px; /* 폭이나 높이가 일정해야 합니다. */ 
            height: 600px; /* 폭이나 높이가 일정해야 합니다. */ 
            position: absolute; 
            top: 50%; /* 화면의 중앙에 위치 */ 
            left: 50%; /* 화면의 중앙에 위치 */ 
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */ 
            
        }

        .style16
        {
            width: 509px;
        }
        .style17
        {
            width: 509px;
            font-family: 굴림체;
            font-size: 10pt;
        }

        .style18
        {
            width: 469px;
        }
        .style19
        {
            width: 469px;
            font-family: 굴림체;
            font-size: 10pt;
        }

        </style>
    
   
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table  style="border: thin solid #000080">
            <tr>
                <td class="style12">
                    공&nbsp;&nbsp;&nbsp;&nbsp;장
                </td>
                <td class="style16">
                    <asp:DropDownList ID="dl_plant_cd" runat="server" AppendDataBoundItems="True" Height="25px"
                        Width="170px" DataSourceID="SqlDataSource1" DataTextField="PLANT_NM" DataValueField="PLANT_CD">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                        SelectCommand="SELECT PLANT_CD, PLANT_NM FROM B_PLANT WHERE (VALID_TO_DT &gt; GETDATE())">
                    </asp:SqlDataSource>
                </td>
                <td class="style12">
                    품&nbsp;&nbsp;&nbsp;&nbsp;목
                </td>
                <td class="style18">
                    <asp:TextBox ID="tb_item_cd" runat="server"></asp:TextBox>
                    <asp:Button ID="bt_item_cd" runat="server"  Text=".." 
                        onclick="bt_item_cd_Click" />
                    <asp:TextBox ID="tb_item_nm" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style12">
                <asp:Label ID="lb_yyyymm" runat="server" Text="기 준 일"></asp:Label>
                </td>
                <td class="style17">
                    <asp:TextBox ID="tb_yyyymm" runat="server" MaxLength="8" ></asp:TextBox>
                &nbsp;(ex)20140101 / 미입력시 현재날짜 기준으로 조회)</td>
                <td class="style12">
                    &nbsp;</td>
                <td class="style19">
                    &nbsp;</td>
            </tr>
        </table>
        <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="0">
        </asp:ScriptManager>
        <asp:Button ID="bt_retrieve" runat="server" OnClick="bt_retrieve_Click" Text="조회"
            Width="100px" />
        
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
        
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="934px" AsyncRendering="False"
            Height="430px" SizeToReportContent="True" KeepSessionAlive ="true">
            <ServerReport Timeout="-1" />
        </rsweb:ReportViewer>
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
       AND a.plant_cd = @plant_cd
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
                <asp:ControlParameter ControlID="dl_plant_cd" Name="plant_cd" 
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
