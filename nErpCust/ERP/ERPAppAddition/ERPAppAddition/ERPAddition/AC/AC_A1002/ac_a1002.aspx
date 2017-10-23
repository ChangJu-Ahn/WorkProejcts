<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ac_a1002.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AC.AC_A1002.ac_a1002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>공사손익 프로젝트별 관리</title>
    <style type="text/css">
        .style12
        {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size:smaller;
        }
        .style1
        {
            width: 400px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
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
        .auto-style1 {
            width: 116px;
        }
        .auto-style2 {
            width: 133px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
        }
        .auto-style22 {
            width: 150px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: right;
            font-size: smaller;
        }
        .auto-style3 {
            height: 28px;
        }
        .auto-style4 {
            width: 100%;
            height: 28px;
        }
        .auto-style5 {
            width: 50px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
        .auto-style55 {
            width: 580px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
        .auto-style7 {
            width: 126px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
        }
        .auto-style8 {
            width: 233px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
        }
         .auto-table {            
            width: 133px;
            border:0;
            background-color: #99CCFF;
            border:0;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
        }
        .auto-style9 {
            width: 170px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
        }
        .auto-style10 {
            width: 133px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            height: 20px;
        }
        .auto-style11 {            
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            font-size: smaller;
        }
        .auto-style56 {
            width: 133px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            height: 23px;
        }
        .auto-style57 {
            width: 170px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            height: 23px;
        }
        .auto-style58 {
            width: 50px;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            height: 23px;
        }
        .auto-style59 {
            width: 150px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: right;
            font-size: smaller;
            height: 23px;
        }
        .auto-style60 {
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            font-size: smaller;
            height: 23px;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table><tr><td class="auto-style3">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
        </td>
        <td class="auto-style4"><asp:Label ID="Label2" runat="server" Text="공사손익 프로젝트별 관리" CssClass=title Width="100%"></asp:Label>
        </td></tr></table>        
        
    </div>
    <div>
    <table style="border: thin solid #000080; ">
        <tr>
            <td class="auto-style2" >
                프로젝트</td>
            <td class="auto-style55">
                <asp:DropDownList ID="DDL_PROJ" runat="server" Height="22px" Width="450px" OnSelectedIndexChanged="DDL_PROJ_SelectedIndexChanged" AutoPostBack="True" >
                </asp:DropDownList>
                <asp:TextBox ID="TXT_PROJ_NM" runat="server" Width="80px" style="text-align: center" Enabled="False" Wrap="False"></asp:TextBox>
                </td>
            <td class="auto-style1">
        <asp:Button ID="btn_search" runat="server" Text="조회" Width="120px" 
            onclick="bt_search_Click" />
            </td>
        <td class="auto-style1">
        <asp:Button ID="btn_save" runat="server" Text="저장" Width="120px" 
            onclick="bt_save_Click" />
            </td>
        </tr>
        </table>
        <table>
        <tr>
            <td class="auto-table" >
                공사진행일</td>
            <td class="auto-style7">
                <asp:TextBox ID="TXT_PROJ_FROM" runat="server" Width="100" MaxLength="8" TextMode="Number"></asp:TextBox>~
                </td>
            <td><asp:TextBox ID="TXT_PROJ_TO" runat="server" Width="100" MaxLength="8" TextMode="Number"></asp:TextBox>
            </td>
                </tr>
            </table>
        <table>
        <tr>
            <td class="auto-style2" >
                완료여부</td>
            <td class="auto-style8" colspan ="2">
                <asp:RadioButton ID="COMP_YES" runat="server" Text ="YES" GroupName="FLAG" />
                <asp:RadioButton ID="COMP_N0" runat="server" Text ="NO" Checked="True" GroupName="FLAG" />
            </td>            
        </tr>
            </table>
        <table>
        <tr>
            <td class="auto-style2" >
                수주금액</td>
            <td class="auto-style9">
                <asp:TextBox ID="TXT_CONT_AMT" runat="server" Width="159px" style="text-align: right" onkeyPress="if ((event.keyCode < 48) || (event.keyCode > 57))  event.returnValue=false;" ></asp:TextBox></td>        
            <td class="auto-style5">원</td>    
            <td class="auto-style22" >
                수주금액
                비고</td>
            <td class="auto-style11">
                <asp:TextBox ID="TXT_CONT_NOTE" runat="server" Width="450px" style="text-align: left"></asp:TextBox></td>        
            
        </tr>
            </table>
        <table>
        <tr>
            <td class="auto-style2" >
                실행원가</td>
            <td class="auto-style9">
                <asp:TextBox ID="TXT_EXE_COST_AMT" runat="server" Width="159px" style="text-align: right" onkeyPress="if ((event.keyCode < 48) || (event.keyCode > 57))  event.returnValue=false;" ></asp:TextBox></td>        
            <td colspan="3" class ="auto-style5">원</td>                
        </tr>
            </table>
        <table>
        <tr>
            <td class="auto-style10" >
                실행율</td>
            <td class="auto-style9">
                <asp:TextBox ID="TXT_ACT_RATE" runat="server" Width="159px" style="text-align: right" onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46))  event.returnValue=false;" MaxLength="5" ></asp:TextBox></td>        
            <td class="auto-style5">%</td>    
            <td class="auto-style22" >
                실행율
                비고</td>
            <td class="auto-style11">
                <asp:TextBox ID="TXT_ACT_NOTE" runat="server" Width="450px" style="text-align: left"></asp:TextBox></td>                    
        </tr>        
            </table>
        <table>
        <tr>
            <td class="auto-style56" >
                <strong>고용/산재요율</strong>
            </td>
            <td class="auto-style57">
                <asp:TextBox ID="TXT_EMP_RATE" runat="server" Width="159px" style="text-align: right" onkeyPress="if (((event.keyCode < 48) || (event.keyCode > 57)) && (event.keyCode != 46))  event.returnValue=false;" MaxLength="5" ></asp:TextBox></td>        
            <td class="auto-style58">%</td>    
            <td class="auto-style59" >
                &nbsp;고용/산재요율
                비고</td>
            <td class="auto-style60">
                <asp:TextBox ID="TXT_EMP_NOTE" runat="server" Width="450px" style="text-align: left"></asp:TextBox></td>                    
        </tr>
        </table>
               <table>
        <tr>
            <td class="auto-style56" >
                <strong>담당자</strong></td>
            <td class="auto-style57">
                <asp:TextBox ID="TXT_MANAGER" runat="server" Width="159px" style="text-align: right"  MaxLength="5" ></asp:TextBox></td>        
            <td class="auto-style58">&nbsp;</td>    
            <td class="auto-style59" >
                &nbsp;변경내역</td>
            <td class="auto-style60">
                <asp:TextBox ID="TXT_REVISION" runat="server" Width="450px" style="text-align: left"></asp:TextBox></td>                    
        </tr>
        </table>      
     <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>            
<script type="text/javascript">   

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
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" />
                <br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
       <cc1:ModalPopupExtender ID="ModalProgress" runat="server" 
            PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        <asp:Panel ID="Panel_Default_Btn" runat="server">
        </asp:Panel>                
    </div>
    </form>
</body>
</html>
