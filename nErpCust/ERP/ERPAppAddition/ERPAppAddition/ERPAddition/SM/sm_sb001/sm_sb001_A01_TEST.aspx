<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sb001_A01_TEST.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sb001.sm_sb001_A01_TEST" %>
<%@ Register assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457" namespace="FarPoint.Web.Spread" tagprefix="FarPoint" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .default_font_size
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align:center;
        }
        .auto-style45 {
            width: 75px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
        }
                            
            .auto-style47 {
            font-size : small;
            height: 21px;
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
            }
            .auto-style101 {
            width: 1151px;
        }            
            .auto-style102 {
            width: 100px;
        }
            #Text1 {
            width: 94px;
        }
            </style>

</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager2" runat="server" EnablePageMethods="true">
    </asp:ScriptManager>
    <div>
    <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label6" runat="server" Text="반출정보관리(경비실)" CssClass=title Width="757%"></asp:Label>
        </td></tr></table>        
        
    </div>
    <div>
                <table style="border: thin solid #000080; width: 950px">
                    <tr>
                        <td class="auto-style101">
                            <table>                                
                                <tr>
                                    <td class="auto-style45">
                                        <asp:Label ID="Label2" runat="server" Text="공   장 : "></asp:Label>
                                    </td>
                                    <td class="auto-style47">
                                        <asp:DropDownList ID="DDL_PLANT" runat="server" Width="110px" Height="24px">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style45">
                                        반출일자:
                                    </td>
                                    <td class="style1">
                                        <asp:TextBox ID="tb_fr_yyyymmdd" runat="server" MaxLength="8" Width="65px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_fr_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_fr_yyyymmdd">
                                        </cc1:CalendarExtender>
                                        <asp:Label ID="Label1" runat="server" Text="~"></asp:Label>
                                        <asp:TextBox ID="tb_to_yyyymmdd" runat="server" MaxLength="8" Width="65px"></asp:TextBox>
                                        <cc1:CalendarExtender ID="tb_to_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                                            Format="yyyyMMdd" TargetControlID="tb_to_yyyymmdd">
                                        </cc1:CalendarExtender>
                                    </td>                                                                        
                                  <td class="auto-style45">
                                        문서번호:
                                    </td>
                                    <td>
                                        <asp:TextBox ID="DOC_NO" runat="server" Width="138px"></asp:TextBox>                                        
                                    </td>
                                    
                                    <td class="auto-style45">
                                        반 출 자:
                                    </td>
                                    <td class="auto-style102">
                                        <asp:TextBox ID="CREATOR" runat="server" Width="100px"></asp:TextBox>
                                    </td>
                                    <td>
                                        <asp:Button ID="btn_SEARCH" runat="server" Text="조회" Width="80px" OnClick="btn_select" />                                        
                                    </td>
                                </tr>                                
                            </table>                            
                        </td>
                    </tr>
                </table>                
            
    
       </div>       
        <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width ="950px">
              <Columns>
                <asp:BoundField DataField="FAC" HeaderText="공장구분"  ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" />
                <asp:BoundField DataField="DOC_NO" HeaderText="문서번호" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF"/>
                <asp:BoundField DataField="APDATE" HeaderText="반출일자" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF"/>
                <asp:BoundField DataField="CREATOR" HeaderText="반출자" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF"/>                
                  <asp:BoundField DataField="STS" HeaderText="반출상태" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF"/>                
                  <asp:BoundField DataField="RESTS" HeaderText="반입상태" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF"/>
                <asp:TemplateField HeaderText="상세화면" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">
                    <ItemTemplate>
                        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Click" Width ="80px"/>
                    </ItemTemplate>
               </asp:TemplateField>
                  
                  <asp:BoundField DataField="PROCESS_INSTANCE_OID" HeaderText="ID" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF" Visible="true"/>                
            </Columns>
        </asp:GridView>
    </div>        
    </form>
</body>
</html>
