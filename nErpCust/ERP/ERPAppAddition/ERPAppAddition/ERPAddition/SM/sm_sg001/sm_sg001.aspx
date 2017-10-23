<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sg001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sg001.sm_sg001" %>
<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>통관 invoice 진행현황 조회</title>
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
        .style13
        {
            width: 380px;
            font-family: 굴림체;
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
            top: 74%; /* 화면의 중앙에 위치 */ 
            left: 46%; /* 화면의 중앙에 위치 */ 
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */ 
            
        }

         .auto-style1 {
             text-align: center;
         }

        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <div>
     <table  style="border: thin solid #000080">
            <tr>
                <td class="style12">
                    Invoice No
                </td>
                <td>
                    <asp:TextBox ID="tb_invoice" runat="server" style="background-color: #FFFFCC; text-align: center;"></asp:TextBox>
                </td>
                 
               
                     <tr>
            
                 <td class="style12">
                     년월일
                </td>
                <td class="auto-style1">
                   <asp:TextBox ID="tb_fr_yyyymmdd" runat="server" BackColor="#FFFFCC" MaxLength="12" Width="130px" style="text-align: center"></asp:TextBox>
                <cc1:CalendarExtender ID="str_fr_dt_CalendarExtender" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="tb_fr_yyyymmdd">
                </cc1:CalendarExtender>
                ~<asp:TextBox ID="tb_to_yyyymmdd" runat="server" BackColor="#FFFFCC" MaxLength="12" Width="130px" style="text-align: center"></asp:TextBox>
                <cc1:CalendarExtender ID="str_to_dt_CalendarExtender" runat="server" Enabled="True"
                    Format="yyyyMMdd" TargetControlID="tb_to_yyyymmdd">
                </cc1:CalendarExtender></td>
                <td>
                   <asp:Button ID="bt_retrieve" runat="server"  Text="조회"
            Width="100px" onclick="bt_retrieve_Click" />
                </td>
            </tr>
              
           </table>
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
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" AsyncRendering="False" 
            Height="687px" Width="1271px">
        </rsweb:ReportViewer>
    </div>
                
    </div>
    </form>
</body>
</html>
