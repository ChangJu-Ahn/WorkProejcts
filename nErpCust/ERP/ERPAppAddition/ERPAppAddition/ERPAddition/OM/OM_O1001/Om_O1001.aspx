<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Om_O1001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.OM.OM_O1001.Om_O1001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>전사특허관리등록</title>
  
    <style type="text/css">       
        .style2
        {
            width: 87%;
            float: left;
            height: 695px;
        }
        .style5
        {
        }
        .style11
        {
            height: 61px;
            width: 497px;
        }
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
        .style27
        {
            width: 331px;
            height: 26px;
        }
        .style47
        {
            width: 331px;
            text-align: left;
        }
        .style48
        {
            width: 331px;
            height: 26px;
            text-align: left;
        }
        .style51
        {
            width: 451px;
        }
        .style52
        {
            width: 451px;
            height: 26px;
        }
    </style>
</head>
<body>
     
 
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <div>
        <table>
            <tr>
                <td>
                   <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
                </td>
                <td style="width: 100%;">
                    <asp:Label ID="Label2" runat="server" Text="전사특허관리" CssClass="title" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    <td>
        <asp:Panel ID="Panel_header" runat="server" Height="66px">
            <asp:Label ID="Label1" runat="server" Text=" 출원번호 :"></asp:Label>
            <asp:TextBox ID="TxtApplyNo" runat="server" BackColor="#FFFFCC"></asp:TextBox>
            <asp:Button ID="BtnSearch" runat="server" OnClick="BtnSearch_Click" Text="찾기" Width="45px" />
            &nbsp;
            <asp:Button ID="BtnLoad" runat="server" OnClick="BtnLoad_Click" Text="조회" Height="21px"
                Width="100px" />
            <asp:Button ID="BtnSave" runat="server" Height="21px" Text="저장" Width="100px" OnClick="BtnSave_Click" />
            <asp:Button ID="BtnChange" runat="server" Height="21px" OnClick="BtnChange_Click"
                Text="수정" Width="100px" />
            <asp:Button ID="BtnDel" runat="server" Height="21px" Style="margin-top: 3px" Text="삭제"
                Width="100px" OnClick="BtnDel_Click1" />
            &nbsp;<asp:Button ID="BtnClear" runat="server" onclick="BtnClear_Click" Text="재작성" />
            </td>
            

            <br />
            <br />
            
        </asp:Panel>
       <asp:Label ID="Label3" runat="server" Text="발명의 명칭(국문)"></asp:Label>
                    <asp:TextBox ID="TxtInvent_Kr_Nm" runat="server" Height="25px" 
        Width="730px"></asp:TextBox>
                    <br />
                    <asp:Label ID="Label4" runat="server" Text="발명의 명칭(영문)"></asp:Label>
                    <asp:TextBox ID="TxtInvent_En_Nm" runat="server" Height="25px" 
                        TextMode="MultiLine" Width="730px"></asp:TextBox>
        <asp:Panel ID="Panel_body" runat="server">
            <table class="style2">
                
                <tr>
                    
                                <td class="style51">
                                <asp:Label ID="Label5" runat="server" Text="관리부서" Width="65px"></asp:Label>
                                    <asp:DropDownList ID="DropDownDept_Cd" runat="server" DataSourceID="SqlDataSource2"
                                        DataTextField="UD_MINOR_NM" DataValueField="UD_MINOR_CD">
                                    </asp:DropDownList>
                                    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                                        SelectCommand="SELECT  UD_MINOR_CD, UD_MINOR_NM FROM B_USER_DEFINED_MINOR WHERE (UD_MAJOR_CD = 'OP001') ORDER BY UD_MINOR_CD">
                                    </asp:SqlDataSource>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                </td>
                                <td class="style47">
                                     <asp:Label ID="Label7" runat="server" Text="공개번호" Width="65px"></asp:Label>
                                    <asp:TextBox ID="TxtOpen_No" runat="server" height="19px" width="139px"></asp:TextBox>
                                    &nbsp;&nbsp;
                                </td>
                    </tr>
                    <tr>
                        <td class="style51">
                            <asp:Label ID="Label6" runat="server" Text="해외대표" Width="65px"></asp:Label> 
                            <asp:DropDownList ID="DropDownRegion" runat="server">
                                <asp:ListItem>해외대표</asp:ListItem>
                                <asp:ListItem>N/A</asp:ListItem>
                            </asp:DropDownList>
                            &nbsp;&nbsp;&nbsp;
                        </td>
                        <td class="style47">
                            <asp:Label ID="Label8" runat="server" Text="공 개 일" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtOpen_DT" runat="server" MaxLength="8" height="19px" 
                                width="139px"></asp:TextBox>
                            <cc1:CalendarExtender ID="TxtOpen_DT_CalendarExtender" runat="server" Enabled="True"
                                TargetControlID="TxtOpen_DT" Format="yyyyMMdd"></cc1:CalendarExtender>
                        </td>
                    </tr>
                    <tr>
                        <td class="style51">
                            <asp:Label ID="Label9" runat="server" Text="상    태" Width="65px"></asp:Label>
                            <asp:DropDownList ID="DropDownStatus" runat="server" DataSourceID="SqlDataSource3"
                                DataTextField="UD_MINOR_NM" DataValueField="UD_MINOR_CD">
                            </asp:DropDownList>
                            <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                                SelectCommand="SELECT  UD_MINOR_CD, UD_MINOR_NM FROM B_USER_DEFINED_MINOR WHERE (UD_MAJOR_CD = 'OP003') ORDER BY UD_MINOR_CD">
                            </asp:SqlDataSource>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </td>
                        <td class="style47">
                            <asp:Label ID="Label10" runat="server" Text="등록번호" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtRegistNo" runat="server" Width="139px" Height="19px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="style52">
                            <asp:Label ID="Label11" runat="server" Text="구    분" Width="65px"></asp:Label>
                            <asp:DropDownList ID="DropDownType_Cd1" runat="server" DataSourceID="SqlDataSource7"
                                DataTextField="UD_MINOR_NM" DataValueField="UD_MINOR_CD">
                            </asp:DropDownList>
                            <asp:SqlDataSource ID="SqlDataSource7" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                                SelectCommand="SELECT UD_MINOR_CD, UD_MINOR_NM FROM B_USER_DEFINED_MINOR WHERE (UD_MAJOR_CD = 'OP004')">
                            </asp:SqlDataSource>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </td>
                        <td class="style48">
                            <asp:Label ID="Label12" runat="server" Text="등 록 일" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtRegist_DT" runat="server" MaxLength="8" height="19px" 
                                width="139px"></asp:TextBox>
                            <cc1:CalendarExtender ID="TxtRegist_DT_CalendarExtender" runat="server" Enabled="True"
                                Format="yyyyMMdd" TargetControlID="TxtRegist_DT"></cc1:CalendarExtender>
                        </td>
                    </tr>
                    <tr>
                        <td class="style51">
                            <asp:Label ID="Label13" runat="server" Text="분    류" Width="65px"></asp:Label>
                            <asp:DropDownList ID="DropDownType_Cd2" runat="server"
                                DataSourceID="SqlDataSource1" DataTextField="UD_MINOR_NM" DataValueField="UD_MINOR_CD">
                            </asp:DropDownList>
                            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                                SelectCommand="SELECT UD_MINOR_CD, UD_MINOR_NM FROM B_USER_DEFINED_MINOR WHERE (UD_MAJOR_CD = 'OP005') ORDER BY UD_MINOR_CD">
                            </asp:SqlDataSource>
                        </td>
                        <td class="style47">
                            <asp:Label ID="Label14" runat="server" Text="출 원 일" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtApply_DT" runat="server"  MaxLength="8" height="19px" 
                                width="139px"></asp:TextBox>
                            <cc1:CalendarExtender ID="TxtApply_DT_CalendarExtender" runat="server" Enabled="True"
                                Format="yyyyMMdd" TargetControlID="TxtApply_DT"></cc1:CalendarExtender>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td class="style52">
                             <asp:Label ID="Label15" runat="server" Text="우 선 권" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtPriorityCd" runat="server" Height="19px" Width="186px"></asp:TextBox>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </td>
                        <td class="style27">
                            <asp:Label ID="Label16" runat="server" Text="출 원 국" Width="65px"></asp:Label>
                            <asp:DropDownList ID="DropDownCountry" runat="server" DataSourceID="SqlDataSource6"
                                DataTextField="UD_MINOR_NM" DataValueField="UD_MINOR_CD">
                            </asp:DropDownList>
                            <asp:SqlDataSource ID="SqlDataSource6" runat="server" ConnectionString="<%$ ConnectionStrings:nepes %>"
                                SelectCommand="SELECT UD_MAJOR_CD,UD_MINOR_CD, UD_MINOR_NM FROM B_USER_DEFINED_MINOR WHERE (UD_MAJOR_CD = 'OP002') ORDER BY UD_MINOR_CD">
                            </asp:SqlDataSource>
                        </td>
                    </tr>
                    <tr>
                        <td class="style51">
                            <asp:Label ID="Label17" runat="server" Text="우 선 일" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtPriorityDT" runat="server" Height="19px" Width="186px"></asp:TextBox>
                            &nbsp;&nbsp;
                        </td>
                        <td class="style47">
                            <asp:Label ID="Label18" runat="server" Text="진 입 국" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtE_Contry_Nm" runat="server" height="19px" width="139px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="style51">
                            <asp:Label ID="Label19" runat="server" Text="대 리 인" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtSubstitute_Nm" runat="server" Height="19px" Width="306px"></asp:TextBox>
                        </td>
                        <td class="style47">
                            <asp:Label ID="Label20" runat="server" Text="발 명 자" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtInvent_Nm" runat="server" Height="19px" Width="185px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="style51">
                            <asp:Label ID="Label21" runat="server" Text="출 원 인" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtApply_Comp" runat="server" Height="19px" Width="306px"></asp:TextBox>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </td>
                        <td class="style47">
                            <asp:Label ID="Label22" runat="server" Text="연차료납기일"></asp:Label>
                            <asp:TextBox ID="TxtAsset_DT" runat="server"  MaxLength="8" height="19px" 
                                style="margin-left: 6px" width="139px"></asp:TextBox>
                            <cc1:CalendarExtender ID="TxtAsset_DT_CalendarExtender" runat="server" Enabled="True"
                                Format="yyyyMMdd" TargetControlID="TxtAsset_DT"></cc1:CalendarExtender>
                        </td>
                    </tr>
                    <tr>
                        <td class="style52">
                            <asp:Label ID="Label23" runat="server" Text="국제출원번호"></asp:Label>
                            <asp:TextBox ID="TxtPct_Apply_No" runat="server" Height="19px" Width="268px"></asp:TextBox>
                        </td>
                        <td class="style48" >
                             <asp:Label ID="Label24" runat="server" Text="기간만료일"></asp:Label>
                            <asp:TextBox ID="TxtExp_DT" runat="server"  MaxLength="8" height="19px" 
                                 width="131px"></asp:TextBox>
                            <cc1:CalendarExtender ID="TxtExp_DT_CalendarExtender" runat="server" Enabled="True"
                                Format="yyyyMMdd" TargetControlID="TxtExp_DT"></cc1:CalendarExtender>
                        </td>
                    </tr>
                    <tr>
                        <td class="style5" colspan="2" dir="ltr" rowspan="1">
                            <asp:Label ID="Label25" runat="server" Text="비    고" Width="65px"></asp:Label>
                            <asp:TextBox ID="TxtRemark" runat="server" Height="51px" Width="326px" 
                                ReadOnly="False" TextMode="MultiLine"></asp:TextBox>
                        
                    </td>
                </tr>
            </table>
        </asp:Panel>
         </div>
      
    </div>
    </form>
</body>
</html>
