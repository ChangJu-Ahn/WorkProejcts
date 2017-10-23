<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_A3001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A3001.AM_A30011" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>ERP 결재(결의전표)</title>
    <style type="text/css">
        .dt
        {   font-family: 굴림체;
            font-size:10pt;
            text-align: center;
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
        .gridstyle
        {
            font-family: 굴림체;
            font-size:10pt;
            text-align: center;
        }
        .style1
        {
            font-family: 굴림체;
            font-size: 10pt;
            text-align: right;
        }
   </style>
   <script type="text/javascript">
       function CheckOtherIsCheckedByGVID(spanChk) {
           var IsChecked = spanChk.checked;
           if (IsChecked) {
               spanChk.parentElement.parentElement.style.backgroundColor = '#228b22';
               spanChk.parentElement.parentElement.style.color = 'white';
           }

           var CurrentRdbID = spanChk.id;
           var Chk = spanChk;
           Parent = document.getElementById('GridView1');
           var items = Parent.getElementsByTagName('input');

           for (i = 0; i < items.length; i++) {
               if (items[i].id != CurrentRdbID && items[i].type == "radio") {
                   if (items[i].checked) {
                       items[i].checked = false;
                       items[i].parentElement.parentElement.style.backgroundColor = 'white';
                       items[i].parentElement.parentElement.style.color = 'black';
                   }
               }
           }
       }
       function CheckOtherIsCheckedByGVIDMore(spanChk) {
           var IsChecked = spanChk.checked;
           if (IsChecked) {
               spanChk.parentElement.parentElement.style.backgroundColor = '#228b22';
               spanChk.parentElement.parentElement.style.color = 'white';
           }
           var CurrentRdbID = spanChk.id;
           var Chk = spanChk;
           Parent = document.getElementById('GridView1');
           for (i = 0; i < Parent.rows.length; i++) {
               var tr = Parent.rows[i];
               var td = tr.childNodes[0];
               var item = td.firstChild;
               if (item.id != CurrentRdbID && item.type == "radio") {
                   if (item.checked) {
                       item.checked = false;
                       item.parentElement.parentElement.style.backgroundColor = 'white';
                       item.parentElement.parentElement.style.color = 'black';
                   }
               }
           }
       }
 
   </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <div>
    <table><tr><td>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td  style="width:100%;"><asp:Label ID="Label2" runat="server" Text="ERP 결재(결의전표)" CssClass=title Width="100%"></asp:Label></td></tr></table>
        
        
    </div>
    <div>
        <asp:Panel ID="Panel_Header" runat="server">
        <table><tr><td>
            <asp:Label ID="Label1" runat="server" Text="결의일자" CssClass="dt"></asp:Label></td>
            <td>
                <asp:TextBox ID="tb_fr_dt" runat="server" CssClass=dt></asp:TextBox> 
                <cc1:CalendarExtender ID="tb_fr_dt_CalendarExtender" runat="server" 
                    Enabled="True" TargetControlID="tb_fr_dt" Format="yyyy-MM-dd">
                </cc1:CalendarExtender>
                ~ 
                <asp:TextBox ID="tb_to_dt" runat="server"  CssClass=dt></asp:TextBox>
                <cc1:CalendarExtender ID="tb_to_dt_CalendarExtender" runat="server" 
                    Enabled="True" Format="yyyy-MM-dd"  TargetControlID="tb_to_dt" >
                </cc1:CalendarExtender>
            </td>
            <td>
                <asp:Button ID="btn_exe" runat="server" Text="조회"  CssClass=dt Width="100px" 
                    onclick="btn_exe_Click" /></td>
                <td>
                    <asp:Button ID="btn_gw" runat="server" Text="상신" CssClass=dt Width="100px" 
                        onclick="btn_gw_Click" /></td>

            </tr></table>
            <table style="width: 100%;">
                <tr>
                    <td>
                        <asp:Label ID="Label3" runat="server" BackColor="#999999" Height="5px" Width="100%"></asp:Label>
                        <br />
                    </td>
                </tr>
            </table>
       
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" 
                CssClass=gridstyle CellPadding="4" ForeColor="#333333" GridLines="None" 
                DataKeyNames="temp_gl_no"  >
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <Columns>
                    <asp:TemplateField HeaderText="구분">
                        <ItemTemplate>
                            <asp:RadioButton ID="rbtn_check" runat="server"  onclick="javascript:CheckOtherIsCheckedByGVID(this);"   />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="결의전표일">
                        <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                        <ItemTemplate>
                            <asp:Label ID="lbl_gl_dt" runat="server" Text='<%# Bind("temp_gl_dt") %>' Width="100px"></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                    </asp:TemplateField>                    
                    <asp:TemplateField HeaderText="결의번호">
                        <HeaderStyle HorizontalAlign="Center" Width="150px" Wrap="False" />
                        <ItemTemplate>
                            <asp:Label ID="lbl_gl_no" runat="server" Text='<%# Bind("temp_gl_no") %>' Width="150px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="부서명">
                        <HeaderStyle HorizontalAlign="Center" Width="150px" Wrap="False" />
                        <ItemTemplate>
                            <asp:Label ID="lbl_dept_nm" runat="server" Text='<%# Bind("dept_nm") %>' Width="150px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="금액">
                        <HeaderStyle HorizontalAlign="Center" Width="150px" Wrap="False" />
                        <ItemTemplate>
                            <asp:Label ID="lbl_gl_amt" runat="server" Text='<%# Bind("dr_amt") %>' Width="150px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="금액(자국)">
                        <HeaderStyle HorizontalAlign="Center" Wrap="False" Width="150px" />
                        <ItemTemplate>
                            <asp:Label ID="lbl_gl_loc_amt" runat="server" Text='<%# Bind("dr_loc_amt") %>' Width="150px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="uniERP등록자">
                        <HeaderStyle HorizontalAlign="Center" Width="150px" Wrap="False" />
                        <ItemTemplate>
                            <asp:Label ID="lbl_gl_userid" runat="server" Text='<%# Bind("insrt_user_id") %>' Width="150px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <EditRowStyle BackColor="#999999" />
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                <SortedAscendingCellStyle BackColor="#E9E7E2" />
                <SortedAscendingHeaderStyle BackColor="#506C8C" />
                <SortedDescendingCellStyle BackColor="#FFFDF8" />
                <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
            </asp:GridView>
            </ContentTemplate>
          </asp:UpdatePanel>
        </asp:Panel>
    </div>
    </form>
</body>
</html>
