<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_A9001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_A9001.AM_A9001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>일일자금일보 실적등록(KO441)</title>

    <style type="text/css">
                .tbl_list{border:1px solid #e8e9ea; border-collapse:collapse; background-color:#fff; font-size: 10pt;
            margin-right: 0px;
        }

        .tbl_list th{background:#113971; color:#fff;  font-weight:bold; text-align:center;}

        .tbl_list td{color:#787878;border-left:1px solid #e8e9ea; text-align:center;}

        .tbl_list th, .tbl_list td{padding:10px 10px 10px 01px; line-height:1.5em; border:1px solid #e8e9ea; font-size: 10pt;}

        .tbl_list .align_l{padding-left:15px }

        .tbl_list .subject{ height:20px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; line-height:1.5em; text-align:left;}

        .tbl_list .cont_text{padding:15px; text-align:left;}

                 .style1
        {
            width: 118px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }  

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
            width: 390px; /* 폭이나 높이가 일정해야 합니다. */ 
            height: 795px; /* 폭이나 높이가 일정해야 합니다. */ 
            position: absolute; 
            top: 123%; /* 화면의 중앙에 위치 */ 
            left: 50%; /* 화면의 중앙에 위치 */ 
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */ 
            
        }

        .auto-style10 {
            width: 60px;
            font-family: 굴림체;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }
        .style2 {
            color : blueviolet;
            font-weight : 500;
        }
        .auto-style15 {
            font-size : small;
            height: 23px;
            width: 28px;
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
        .auto-style17 {
            font-size: small;
            width: 28px;
        }

        .auto-style30 {
            font-size : small;
            color : black;
        }
        .auto-style31 {
            width: 160px;
        }
        .auto-style32 {
            font-size : small;
            Height: 16px;
            Width: 160px;
        }
        .auto-style38 {
            font-size : small;
            color : black;
            width: 210px;
            background-color: #99CCFF;
        }
        .auto-style40 {
            font-size : small;
            Height: 16px;
            Width: 160px;
            background-color: #99CCFF;
        }
        .auto-style41 {
            font-size : small;
            color : black;
            background-color: #99CCFF;
        }
        .auto-style42 {
            font-size: small;
            width: 28px;
            background-color: #99CCFF;
        }
        </style>
</head>
<body>

    <script lang ="JavaScript">
<!--
    function currency(obj) {
        if (event.keyCode >= 48 && event.keyCode <= 57) {

        } else {
            event.returnValue = false
        }
    }
    function com(obj) {
        obj.value = unComma(obj.value);
        obj.value = Comma(obj.value);
    }
    function Comma(input) {

        var inputString = new String;
        var outputString = new String;
        var counter = 0;
        var decimalPoint = 0;
        var end = 0;
        var modval = 0;

        inputString = input.toString();
        outputString = '';
        decimalPoint = inputString.indexOf('.', 1);

        if (decimalPoint == -1) {
            end = inputString.length - (inputString.charAt(0) == '0' ? 1 : 0);
            for (counter = 1; counter <= inputString.length; counter++) {
                var modval = counter - Math.floor(counter / 3) * 3;
                outputString = (modval == 0 && counter < end ? ',' : '') + inputString.charAt(inputString.length - counter) + outputString;
            }
        }
        else {
            end = decimalPoint - (inputString.charAt(0) == '-' ? 1 : 0);
            for (counter = 1; counter <= decimalPoint ; counter++) {
                outputString = (counter == 0 && counter < end ? ',' : '') + inputString.charAt(decimalPoint - counter) + outputString;
            }
            for (counter = decimalPoint; counter < decimalPoint + 3; counter++) {
                outputString += inputString.charAt(counter);
            }
        }
        return (outputString);
    }

    /* -------------------------------------------------------------------------- */
    /* 기능 : 숫자에서 Comma 제거                                                 */
    /* 파라메터 설명 :                                                            */
    /*        -  input : 입력값                                                   */
    /* -------------------------------------------------------------------------- */
    function unComma(input) {
        var inputString = new String;
        var outputString = new String;
        var outputNumber = new Number;
        var counter = 0;
        if (input == '') {
            return 0
        }
        inputString = input;
        outputString = '';
        for (counter = 0; counter < inputString.length; counter++) {
            outputString += (inputString.charAt(counter) != ',' ? inputString.charAt(counter) : '');
        }
        outputNumber = parseFloat(outputString);
        return (outputNumber);
    }

    //-->
</script>


    <form id="form1" runat="server">  
    <div>
    <table>
        <tr>
            <td>
                <asp:Image ID="Image1" runat="server" ImageUrl="~/Img/folder.gif" />
            </td>
            <td style="width:100%;">
            <b><asp:Label ID="Label3" runat="server" Text="&nbsp일일자금일보 실적등록(AMC)" CssClass="title" Width="100%"></asp:Label></b>
            </td>
        </tr>
    </table>
    </div>
    <div>
           <table style="border: thin solid #000080; height: 43px;">
                 <td class="auto-style10">  
                     <asp:Label ID="lb_yyyy" runat="server" Text="년도(ex:1989)" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>  
            </td >
                <td class="style3">
                    <asp:TextBox ID="txt_yyyy" runat="server" style="background-color: #FFFF99" Width="100px"></asp:TextBox>
            </td>
            <td class="auto-style10">  
                     <asp:Label ID="lb_mm" runat="server" Text="월(ex:03)" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
             <td class="style3">
                    <asp:TextBox ID="txt_mm" runat="server" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
                        
            </td>
             <td class="auto-style10">  
                     <asp:Label ID="lb_dd" runat="server" Text="일(ex:07)" BackColor="#99CCFF" 
                         Font-Bold="True" style="text-align: center; font-size: small"></asp:Label>
                       
            </td >
             <td class="style3">
                    <asp:TextBox ID="txt_dd" runat="server" style="background-color: #FFFF99" Width="80px"></asp:TextBox>
                        
            </td>
           <td class="style56">
               <asp:Button ID="btn_select" runat="server" Text="조회" 
                                    Width="100px" OnClick="btn_select_Click"  />
                    </td>
                        <td class="style56">
                                
                                <asp:Button ID="btn_save" runat="server" Text="저장" 
                                    Width="100px" OnClick ="btn_save_Click"  />
                    </td>
                        <td class="style56">
                                
                                <asp:Button ID="btn_reselect" runat="server" Text="수정" 
                                    Width="100px" OnClick="btn_reselect_Click" />
                    </td>
<%--            <td class="auto-style1">
                                
                                <asp:Button ID="btn_delete" runat="server" Text="삭제" 
                                    Width="100px" OnClick="btn_delete_Click"  />
                    </td>         삭제할 경우 이력을 확인할 수 없어 수정버튼으로 대체--%>

    </table>
        </div>
        <asp:ScriptManager ID ="scriptmanager1" runat ="server"></asp:ScriptManager>
        <asp:UpdatePanel ID="Panel_List_View" runat="server" UpdateMode="Conditional">
            <ContentTemplate>
        <div id="div_down_spread">
            <table  class="tbl_list">
                <tr>
                    <th class="auto-style17"><asp:Label ID="lb_hd_no" runat="server" Text="NO" Font-Size="Small" Height="18px"></asp:Label></th>
                    <th colspan="1" class="auto-style30"><asp:Label ID="lb_hd_list" runat="server" Text="항목" Width="107px" Font-Size="Small" Height="16px"></asp:Label></th>
                    <th class="auto-style31">&nbsp;<asp:Label ID="lb_hd_amt" runat="server" Text="금액" Font-Size="Small"></asp:Label></th>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no1" runat="server" Text="1." CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list1" runat="server" Text="매출대전회수" Width="80px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt1" runat="server" style="text-align: center"  onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no2" runat="server" Text="2."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list2" runat="server" Text="부가세 환급" Width="80px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt2" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no3" runat="server" Text="3."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list3" runat="server" Text="관세 환급" Width="80px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt3" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no4" runat="server" Text="4."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list4" runat="server" Text="임대보증금/임대료 입금" Width="159px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt4" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style15"><asp:Label ID="lb_no5" runat="server" Text="5."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list5" runat="server" Text="수입이자 입" Width="80px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt5" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no6" runat="server" Text="6."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list6" runat="server" Text="기   타" Width="80px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt6" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no7" runat="server" Text="7."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style38"><asp:Label ID="lb_list7" runat="server" Text="영업활동상의 자금수입 (합계)" Width="203px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt7" runat="server" style="text-align: center"  ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no8" runat="server" Text="8."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list8" runat="server" Text="원자재 매입대금 지급" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt8" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no9" runat="server" Text="9."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list9" runat="server" Text="급여와 상여" Width="80px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt9" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no10" runat="server" Text="10."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list10" runat="server" Text="퇴직금의 지급" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt10" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no11" runat="server" Text="11."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list11" runat="server" Text="원천제세 납부" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt11" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no12" runat="server" Text="12."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list12" runat="server" Text="법정복리비 납부" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt12" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no13" runat="server" Text="13."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list13" runat="server" Text="인건비의 지급 (계산)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt13" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no14" runat="server" Text="14."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list14" runat="server" Text="경 비" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt14" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no15" runat="server" Text="15."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list15" runat="server" Text="임대보증금/임대료 지급" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt15" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no16" runat="server" Text="16."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list16" runat="server" Text="부가세 납부" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt16" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no17" runat="server" Text="17."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list17" runat="server" Text="지급이자" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt17" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no18" runat="server" Text="18."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list18" runat="server" Text="기 타" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt18" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no19" runat="server" Text="19."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list19" runat="server" Text="영업활동상의 자금지출 (계산)" Width="177px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt19" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no20" runat="server" Text="20."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list20" runat="server" Text="영업상의 순 자금유입 (계산)" Width="178px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt20" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no21" runat="server" Text="21."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list21" runat="server" Text="고정자산 매각" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt21" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no22" runat="server" Text="22."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list22" runat="server" Text="대여금 회수" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt22" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no23" runat="server" Text="23."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list23" runat="server" Text="장기금융상품의 회수" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt23" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no24" runat="server" Text="24."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list24" runat="server" Text="투자유가증권/출자금 회수" Width="180px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt24" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no25" runat="server" Text="25."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list25" runat="server" Text="투자활동의 Cash in (계산)" Width="179px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt25" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no26" runat="server" Text="26."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list26" runat="server" Text="토지/건물/구축물 취득" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt26" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no27" runat="server" Text="27."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list27" runat="server" Text="기계장치 취득" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt27" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no28" runat="server" Text="28."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list28" runat="server" Text="차량 및 공기구 취득" Width="180px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt28" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no29" runat="server" Text="29."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list29" runat="server" Text="고정자산의 취득 (계산)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt29" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no30" runat="server" Text="30."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list30" runat="server" Text="대여금 지급" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt30" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no31" runat="server" Text="31."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list31" runat="server" Text="장기금융상품의 매입" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt31" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no32" runat="server" Text="32."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list32" runat="server" Text="투자유가증권/출자금 지급" Width="180px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt32" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no33" runat="server" Text="33."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list33" runat="server" Text="투자활동의 Cash out (계산)" Width="189px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt33" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no34" runat="server" Text="34."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list34" runat="server" Text="투자활동의 자금흐름 (계산)" Width="173px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt34" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no35" runat="server" Text="35."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list35" runat="server" Text="신규 차입" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt35" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no36" runat="server" Text="36."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list36" runat="server" Text="USANCE 차입" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt36" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no37" runat="server" Text="37."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list37" runat="server" Text="증자 등" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt37" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no38" runat="server" Text="38."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list38" runat="server" Text="재무활동의 Cash in (계산)" Width="168px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt38" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no39" runat="server" Text="39."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list39" runat="server" Text="차입금의 상환" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt39" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no40" runat="server" Text="40."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list40" runat="server" Text="USANCE 상환" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt40" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style17"><asp:Label ID="lb_no41" runat="server" Text="41."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style30"><asp:Label ID="lb_list41" runat="server" Text="배당금/자기주식 취득 외" Width="190px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style32"><asp:TextBox ID="tb_amt41" runat="server" style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no42" runat="server" Text="42."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list42" runat="server" Text="재무활동의 Cash Out (계산)" Width="188px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt42" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no43" runat="server" Text="43."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list43" runat="server" Text="재무활동의 현금 흐름 (계산)" Width="170px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt43" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no44" runat="server" Text="44."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list44" runat="server" Text="전월이월 자금 (계산)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt44" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no45" runat="server" Text="45."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list45" runat="server" Text="자금의 과부족 (계산)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt45" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" ></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="auto-style42"><asp:Label ID="lb_no46" runat="server" Text="46."   CssClass="style2"></asp:Label></td>
                    <td class="auto-style41"><asp:Label ID="lb_list46" runat="server" Text="차월이월 자금 (계산)" Width="150px" CssClass="style2"></asp:Label></td>
                    <td class="auto-style40"><asp:TextBox ID="tb_amt46" runat="server" style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)"></asp:TextBox></td>
                </tr>

            </ContentTemplate>
                
    </asp:UpdatePanel>

    </form>
</body>
</html>

