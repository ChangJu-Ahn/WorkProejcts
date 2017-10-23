<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sb001_A03.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sb001.sm_sb001_A03" EnableEventValidation="False" %>
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
            width: 90px;
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
            .auto-style103 {
            width: 90px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
            height: 21px;
        }
        .auto-style104 {
            height: 21px;
        }
        .auto-style105 {
            width: 100px;
            height: 21px;
        }
        .list_comment {font-size:9pt; font-family:'Bodoni MT Poster'; width:100px}
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
        <td class="auto-style4"><asp:Label ID="Label6" runat="server" Text="반출정보조회" CssClass=title Width="757%"></asp:Label>
        </td></tr></table>        
        
    </div>
    <div>
                <table style="border: thin solid #000080; width: 1000px">
                    <tr>
                        <td class="auto-style101">
                            <table>                                
                                <tr>
                                    <td class="auto-style45">
                                        <asp:Label ID="Label2" runat="server" Text="공   장: "></asp:Label>
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
                                        <asp:TextBox ID="DOC_NO" runat="server" Width="120px"></asp:TextBox>                                        
                                    </td>
                                    
                                    <td class="auto-style45">
                                        반 출 자:</td>
                                    <td class="auto-style102">          
                                        <asp:TextBox ID="CREATOR" runat="server" Width="100px"></asp:TextBox>
                                    </td>
                                    <td>
                                        <asp:Button ID="btn_SEARCH" runat="server" Text="조회" Width="80px" OnClick="btn_select" />                                        
                                    </td>
                                </tr>
                                <tr>
                                    <td class="auto-style103">
                                        <asp:Label ID="Label3" runat="server" Text="반출상태: "></asp:Label>
                                    </td>
                                    <td class="auto-style47">
                                        <asp:DropDownList ID="DDL_STS" runat="server" Width="110px" Height="24px">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style103">
                                        반입상태:
                                    </td>
                                    <td class="auto-style104">
                                        <asp:DropDownList ID="DDL_RESTS" runat="server" Width="120px" Height="24px">
                                        </asp:DropDownList>
                                    </td>                                                                        
                                  <td class="auto-style103">
                                        품 목:
                                    </td>
                                    <td class="auto-style104"><asp:TextBox ID="DDL_GOODS_NM" runat="server" Width="120px"></asp:TextBox></td>
                                    
                                    <td class="auto-style103">
                                        미완료사유:</td>
                                    <td class="auto-style105">
                                        <asp:DropDownList ID="DDL_INCOMP" runat="server" Width="105px" Height="24px">
                                        </asp:DropDownList>
                                    </td>
                                    <td class="auto-style104">                                                                               
                                      <asp:Button ID="btn_Excel" runat="server" Text="Excel" Width="80px" OnClick="Excel_Click"/>                                                                                
                                    </td>                                    
                                </tr>
                            </table>                            
                        </td>
                    </tr>
                </table>                
            
    
       </div>       
        <div>
               <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" ForeColor="#333333" Width="2000px" CssClass="list_comment" >
                      <AlternatingRowStyle BackColor="White"/>
                      <Columns>
                        <asp:BoundField DataField="FAC" HeaderText="공장구분"  ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">
                            <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center" Wrap="true" CssClass="list_comment" VerticalAlign="Middle"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="APDATE" HeaderText="반출일자" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">
                            <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center" Width="85px" CssClass="list_comment"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DOC_NO" HeaderText="문서번호" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                            <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center" Width="80px" Wrap="true" CssClass="list_comment" ></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="SUBJ" HeaderText="제목" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">
                            <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="320px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                        <asp:BoundField DataField="CREATOR" HeaderText="반출자" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF"> 
                            <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center" Width="70px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="APPR_DEPT" HeaderText="기안부서" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="95px" Wrap="true" CssClass="list_comment" ></ItemStyle>
                          </asp:BoundField>                          
                          <asp:BoundField DataField="STS" HeaderText="반출상태" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="70px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="OUT_DT" HeaderText="반출확인일" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="85px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="OUT_COF_USER" HeaderText="반출확인자" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">  
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="70px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="RESTS" HeaderText="반입상태" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="70px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="GOODS_NM" HeaderText="품목" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Left" Width="200px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="AMOUNT" HeaderText="수량" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="60px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="UNIT" HeaderText="단위" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="40px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="FROM_FD_RAD" HeaderText="반출사유" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="70px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="RECARRY_DATE" HeaderText="반입예정일" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="85px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="IN_DT_AMT" HeaderText="반입수량" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="60px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="IN_DT" HeaderText="반입일" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="85px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="IN_USER" HeaderText="반입자" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="85px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="INCOMPLETE_NM" HeaderText="미완료사유" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="90px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="REMARK" HeaderText="상세정보" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="350px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="PROCESS_INSTANCE_OID" HeaderText="문서ID" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="100px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                          <asp:BoundField DataField="PARENT_INSTANCE_OID" HeaderText="모문서ID" ItemStyle-HorizontalAlign ="Center" HeaderStyle-BackColor="#99CCFF">                
                              <HeaderStyle BackColor="#99CCFF"></HeaderStyle>
                              <ItemStyle HorizontalAlign="Center" Width="100px"  Wrap="true" CssClass="list_comment"></ItemStyle>
                          </asp:BoundField>
                    </Columns>
                      <EditRowStyle BackColor="#7C6F57" />
                      <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
                      <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="BLACK" />
                      <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
                      <RowStyle BackColor="#E3EAEB" />
                      <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                      <SortedAscendingCellStyle BackColor="#F8FAFA" />
                      <SortedAscendingHeaderStyle BackColor="#246B61" />
                      <SortedDescendingCellStyle BackColor="#D4DFE1" />
                      <SortedDescendingHeaderStyle BackColor="#15524A" />
                </asp:GridView>    
                
    </div>        
    </form>
</body>
</html>
