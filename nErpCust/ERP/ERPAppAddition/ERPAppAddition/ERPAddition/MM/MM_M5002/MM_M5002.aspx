<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_M5002.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_M5002.MM_M5002" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>자동발주서 생성(Display)</title>
    <style type="text/css">
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
            background-color: #FFFFFF;
            font-weight: 700;
        }
        .default_font_background
        {
            font-family: 굴림체;
            font-size:10pt;
            
        }
        .style13
        {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
            height: 26px;
        }
              
       
        .style17
        {
            width: 270px;
        }
      
       
        .style19
        {
            width: 167px;
        }
      
       
        .style20
        {
            font-size: small;
        }
      
       
        .style27
        {
            width: 495px;
        }
              
       
        .style30
        {
            width: 735px;
        }
      
       
        .style31
        {
            width: 110px;
        }
      
       
        .style32
        {
            font-size: small;
            color: #FF0000;
        }
      
       
        .style33
        {
            height: 25px;
        }
        .style34
        {
            width: 202px;
            height: 25px;
        }
      
       
        .style35
        {
            width: 516px;
        }
      
       
        .style36
        {
            width: 202px;
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
    <form id="form1" runat="server" enctype="multipart/form-data">
    <asp:ScriptManager ID="ScriptManager2" runat="server">
    </asp:ScriptManager>
    <div>
    <table><tr><td>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td  style="width:100%;"><asp:Label ID="Label2" runat="server" Text="자동발주서 생성(Display)" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>       
        
    </div>
    
   
   
    <table style=" table-layout:fixed; border: thin solid #000080; width:100%; ">
            <td class="style17">
            <table><tr>
                <td class="style13">
                    <strong>발주번호</strong></td>
                <td >
                    <asp:TextBox ID="tb_po_no" runat="server" Width="148px"></asp:TextBox>
                </td> </table>       
            
            <td class="style19" >
                    <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                        <ContentTemplate>
                            <asp:Button ID="btn_pop_po" runat="server" Text="..." Width="20px" 
                                onclick="btn_pop_po_Click1" />
                            <asp:Button ID="btn_retrieve" runat="server"  Text="미리보기" 
                                Width="89px" onclick="btn_retrieve_Click1" />
                        </ContentTemplate>
                    <Triggers>
                   <asp:PostBackTrigger ControlID = "btn_pop_po" />
                   </Triggers>

                    </asp:UpdatePanel>
                </td>
               <td>
                    <asp:RadioButtonList ID="rbt_select" runat="server" CssClass="default_font_size"
                        RepeatDirection="Horizontal" 
                        AutoPostBack="True" 
                        onselectedindexchanged="rbt_select_SelectedIndexChanged"  >
                        <asp:ListItem Value="keyin">추가입력</asp:ListItem>
                                                
                    </asp:RadioButtonList>
                </td>
        
            
                
               </tr></table>       
            
         <asp:Panel ID="Panel_keyin" runat="server" Visible="False">
               <table><tr><td><asp:Label ID="lb_warranty" runat="server" Text="* Warranty : "  CssClass="default_font_size"></asp:Label></td>
               <td class="style30"><asp:TextBox ID="tb_keyin_warranty" runat="server" Width="620px"> </asp:TextBox>
                   <span class="style20">(미입력시 공란)</span></td> 
               </tr>
               <tr><td><asp:Label ID="lb_remark" runat="server" Text="* Remark : "  CssClass="default_font_size"></asp:Label></td>
               <td class="style30"><asp:TextBox ID="tb_keyin_remark" runat="server" Width="620px"></asp:TextBox> 
                   <span class="style20">(미입력시 공란)</span></td>
                   <td>
                       <asp:Button ID="btn_keyin_save" runat="server" Text="저장" 
                            Width="45px" onclick="btn_keyin_save_Click" style="font-weight: 700" />
                   </td>
                   <td>
                       <asp:Button ID="btn_update" runat="server" onclick="btn_update_Click" 
                           style="font-weight: 700" Text="수정" Width="45px" />
                   </td>
               </tr></table>
               </asp:Panel> 
              
      
    
    
    
    
    <asp:Panel ID="Panel_Report" runat="server">
       <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="" 
            SizeToReportContent="True" Width="">

        </rsweb:ReportViewer> 
    </asp:Panel>
 
    </form>
</body>
</html>

