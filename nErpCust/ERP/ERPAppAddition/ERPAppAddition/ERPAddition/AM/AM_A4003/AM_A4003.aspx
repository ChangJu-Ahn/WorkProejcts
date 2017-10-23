<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeBehind="AM_A4003.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A4003.AM_A4003" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>은행이체리스트_a5410oa1_ko441</title>
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

        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table>
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
                <asp:AsyncPostBackTrigger ControlID="GridView1">
                </asp:AsyncPostBackTrigger>
            </Triggers>
            <ContentTemplate>
                     <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width ="1660px">
                      <Columns>
                          <asp:BoundField DataField="CT" HeaderText="의뢰인번호"  ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="BANK_CD" HeaderText="은행코드"  ItemStyle-HorizontalAlign ="Left" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="80px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="MINOR_NM" HeaderText="은행명"  ItemStyle-HorizontalAlign ="Left" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="200px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="BANK_ACCT_NO" HeaderText="입금계좌번호"  ItemStyle-HorizontalAlign ="Left" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="200px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>
                          <asp:BoundField DataField="AMT" HeaderText="이체의뢰금액" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign ="Right" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="140px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>                          
                          <asp:BoundField DataField="TXT1" HeaderText="입지코드" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>
                          <asp:BoundField DataField="TXT2" HeaderText="적요" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False"/>                          
                          <asp:BoundField DataField="DEAL_BP_CD" HeaderText="입력예금주명" ItemStyle-HorizontalAlign ="Left" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="240px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small" HtmlEncode ="False" Visible ="False"/>                          
                          <asp:BoundField DataField="BP_NM" HeaderText="원장예금주명" ItemStyle-HorizontalAlign ="Left" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="340px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>                          
                          <asp:BoundField DataField="TXT3" HeaderText="Qty"  ItemStyle-HorizontalAlign ="Left" HeaderStyle-BackColor="#99CCFF" ItemStyle-Width="100px" ItemStyle-Font-Names="Arial" ItemStyle-Font-Size="Small"/>                                                    
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
