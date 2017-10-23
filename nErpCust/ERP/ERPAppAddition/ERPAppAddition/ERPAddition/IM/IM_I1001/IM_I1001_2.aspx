<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeBehind="IM_I1001_2.aspx.cs" Inherits="ERPAppAddition.ERPAddition.IM.IM_I1001.IM_I1002" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Aging Stock조회_NEW</title>
    <style type="text/css">
        .style12
        {
            width: 80px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 맑은고딕;
            font-size:10pt;
            text-align: center;
        }
        .style13
        {
            width: 380px;
            font-family: 맑은고딕;
            font-size:10pt;
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
            border: 1px#c0c0c0;
            background: #f0f0f0;
            padding: 10px;
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
        .div_center
        {
            width: 607px; /* 폭이나 높이가 일정해야 합니다. */ 
            height: 508px; /* 폭이나 높이가 일정해야 합니다. */ 
            position: absolute; 
            top: 50%; /* 화면의 중앙에 위치 */ 
            left: 50%; /* 화면의 중앙에 위치 */ 
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */
        }

        .style15
        {
            font-size: small;
        }

        .auto-style1 {
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 맑은 고딕;
            font-size: 10pt;
            text-align: center;
        }
        .auto-style3 {
            font-family: 맑은 고딕;
            font-size: 10pt;
        }

         .grpCell1{
             background-color:LightGreen;
         }
        .grpCell2{
            background-color:skyblue;
        }
        .grpCell3{
            background-color:orange;
        }
        .grpCell4{
            background-color:OrangeRed;
        }
        .grpCellW{
            background-color:white;
        }
        

        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table>
            <tr>
                <td>
                    <table  style="border: thin solid #000080">
                            <tr>
                                <td class="auto-style1">
                                    공장
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddl_plant_cd" runat="server" AppendDataBoundItems="True" Height="25px"
                                        Width="170px" DataSourceID="SqlDataSource1" DataTextField="PLANT_NM" 
                                        DataValueField="PLANT_CD" AutoPostBack="True">
                                    </asp:DropDownList><asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                                        SelectCommand="SELECT PLANT_CD, PLANT_NM FROM B_PLANT WHERE (VALID_TO_DT &gt; GETDATE())"></asp:SqlDataSource>
                                </td>
               
                                <td class="auto-style1">
                                <asp:Label ID="Label4" runat="server" Text="품목"></asp:Label>
                                </td>
                                <td class="auto-style3">
                                 <asp:TextBox ID="tb_item_cd" runat="server"></asp:TextBox>
                                    </td>                               
                            </tr>
                            </table>
                </td>
                <td>
                    <asp:Button ID="bt_retrieve" runat="server" OnClick="bt_retrieve_Click" Text="조회" Width="120px" Height="30"/>
                </td>
                <td>
                    <asp:Button ID="btn_excel" runat="server" OnClick="Excel_Click" Text="Excel" Width="120px" Height="30"/>
                </td>
            </tr>
        </table> 
    </div>
    <div>
       <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="0">
        </asp:ScriptManager>
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
            <ContentTemplate>
                     <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" OnRowDataBound="GridView1_RowDataBound" OnRowCreated="GridView1_RowCreated" Width ="2500px">
                      <Columns>
                          <asp:BoundField DataField="ITEM_CD" HeaderText="품목코드"  ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="150px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="ITEM_NM" HeaderText="품목명"  ItemStyle-HorizontalAlign ="Left" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="400px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="ORDER_UNIT_PUR" HeaderText="단위"  ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Smaller"/>
                          <asp:BoundField DataField="AVG_DAYS" HeaderText="평균일수"  ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="110px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Smaller"/>
                          <asp:BoundField DataField="INV_QTY" HeaderText="Total Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="MOVING_AVG_PRC" HeaderText="단가" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="DIFF_AA" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="DIFF_AA_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="DIFF_A" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="DIFF_A_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="DIFF_B" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="DIFF_B_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_C" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_C_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_D" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_D_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_E" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_E_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_F" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_F_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_G" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_G_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_H" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_H_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_I" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_I_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_J" HeaderText="Qty" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="DIFF_J_AMT" HeaderText="Total" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                    </Columns>
                    </asp:GridView>
                </ContentTemplate>
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
    </form>
</body>
</html>
