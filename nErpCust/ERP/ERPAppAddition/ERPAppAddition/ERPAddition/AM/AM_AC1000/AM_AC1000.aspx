<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AM_AC1000.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AM.AM_AC1000.AM_AC1000" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>일월자금실적_등록(Display)</title>

    <style type="text/css">
        .tbl_list {
            font-family: 돋음;
            border: 3px solid #e8e9ea;
            border-color: black;
            border-collapse: collapse;
            background-color: #fff;
            font-size: 10pt;
            margin-right: 0px;
        }

        .tbl_list_1 {
            font-family: 돋음;
            border-top: 3px solid #e8e9ea;
            border-right: 3px solid #e8e9ea;
            border-left: 3px solid #e8e9ea;
            border-color: black;
            border-collapse: collapse;
            background-color: #fff;
            font-size: 10pt;
            margin-right: 0px;
        }

        .style22 {
            width: 60px;
            font-family: 돋음;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }

        .tbl_list_2 {
            font-family: 돋음;
            border-bottom: 3px solid #e8e9ea;
            border-right: 3px solid #e8e9ea;
            border-left: 3px solid #e8e9ea;
            border-color: black;
            border-collapse: collapse;
            background-color: #fff;
            font-size: 10pt;
            margin-right: 0px;
        }

            .tbl_list_1 th {
                background: Orange;
                color: black;
                font-weight: bold;
                font-family: 돋음;
                text-align: center;
                border-color: gray;
                border: 1px solid #e8e9ea;
                padding: 5px 10px 10px 01px;
                line-height: 1.5em;
            }

             .tbl_list_2 td {
                font-family: 돋음;
                padding: 5px 10px 10px 01px;
                line-height: 1.5em;
                border: 1px solid #e8e9ea;
                font-size: 10pt;
                border-color: gray;
            }

            .tbl_list .align_l {
                padding-left: 15px;
            }

            .tbl_list .subject {
                font-family: 돋음;
                height: 20px;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                line-height: 1.5em;
                text-align: left;
            }

            .tbl_list .cont_text {
                padding: 15px;
                text-align: left;
            }

        .stytitle {
            width: 70px;
            font-family: 돋음;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: Silver;
            height: 7px;
        }

         .stytitle_1 {
            width: 70px;
            font-family: 돋음;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }

        

        .stytitle2 {
            width: 120px;
            background-color: Silver;
            font-weight: bold;
            font-family: 돋음;
            font-size: 10pt;
            text-align: center;
        }

        .modalBackground {
            background-color: #CCCCFF;
            filter: alpha(opacity=40);
            opacity: 0.5;
        }

        .modalBackground2 {
            background-color: Gray;
            filter: alpha(opacity=50);
            opacity: 0.5;
        }

        .updateProgress {
            background-color: #ffffff;
            position: absolute;
            width: 150px;
            height: 65px;
        }

        .fixedheadercell {
            FONT-WEIGHT: bold;
            FONT-SIZE: 10pt;
            WIDTH: 200px;
            COLOR: white;
            FONT-FAMILY: Arial;
            BACKGROUND-COLOR: darkblue;
        }

        .fixedheadertable {
            left: 0px;
            position: relative;
            top: 0px;
            padding-right: 2px;
            padding-left: 2px;
            padding-bottom: 2px;
            padding-top: 2px;
        }

        .gridcell {
            WIDTH: 200px;
        }

        .div_center {
            width: 390px; /* 폭이나 높이가 일정해야 합니다. */
            height: 795px; /* 폭이나 높이가 일정해야 합니다. */
            position: absolute;
            top: 123%; /* 화면의 중앙에 위치 */
            left: 50%; /* 화면의 중앙에 위치 */
            margin: -43px 0 0 -120px; /* 높이의 절반과 너비의 절반 만큼 margin 을 이용하여 조절 해 줍니다. */
        }

        .auto-stytitle0 {
            width: 60px;
            font-family: 돋음;
            font-size: smaller;
            font-weight: 700;
            text-align: center;
            background-color: #99CCFF;
            height: 7px;
        }

        .style2 {
            color: #0C274C;
            font-weight: 500;
            font-weight: bold;
            text-align: center;
            font-size: 9pt;
        }

        .title {
            font-family: 돋음;
            font-size: 10pt;
            text-align: left;
            font-weight: bold;
            background-color: #EAEAEA;
            color: Blue;
            vertical-align: middle;
            display: table-cell;
            line-height: 25px;
            height: 25px;
        }

    </style>

    <script lang="javascript">
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

        function scrollX() {
            document.all.mainDisplay.scrollLeft = document.all.bottomLine.scrollLeft;
            document.all.top_title.scrollLeft = document.all.bottomLine.scrollLeft;
        }

        function scrollX_day() {
            document.all.mainDisplay_1.scrollLeft = document.all.bottomLine.scrollLeft;
            document.all.top_title_1.scrollLeft = document.all.bottomLine.scrollLeft;
        }

        function checkFourNumber(text) {
            if (text.value.length != 4) {
                alert("년도는 4자리만 입력가능합니다.");
                document.getElementById('txt_yyyy').value = "";
            }
        }

        function checkTooNumber_m(text) {
            if (text.value.length != 2) {
                alert("월 은 2자리만 입력가능합니다.");
                document.getElementById('txt_mm').value = "";
            }
        }

        function checkTooNumber_d(text) {
            if (text.value.length != 2) {
                alert("일 은 2자리만 입력가능합니다.");
                document.getElementById('txt_dd').value = "";
            }
        }

    </script>

</head>

<body>
    <form id="form1" runat="server">
        <div>
            <table>
                <tr>
                    <td>
                        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
                    </td>
                    <td style="width: 100%;">
                        <asp:Label ID="Label1" runat="server" Text="일월자금실적_등록(Display)" CssClass="title" Width="100%"></asp:Label>
                    </td>
                </tr>
            </table>
            <table style="border: thin solid #000050; height: 31px;">
                <tr>
                    <td class ="style22">
                        <asp:Label ID="Label17" runat="server" Text="등록구분" BackColor="#99CCFF" 
                            Font-Bold="True" Style="text-align: center; font-size: small" Width="65px"></asp:Label>
                    </td>
                    <td>
                        <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small"
                            RepeatDirection="Horizontal"
                            OnSelectedIndexChanged="rbl_view_type_SelectedIndexChanged"
                            AutoPostBack="True" Width="220px" Style="margin-left: 0px; font-weight: 700;"
                            BackColor="White" Height="21px">
                            <asp:ListItem Value="A">일 자금실적</asp:ListItem>
                            <asp:ListItem Value="B">월 자금계획</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
            </table>
        </div>

        <asp:ScriptManager ID="scriptmanager1" runat="server"></asp:ScriptManager>

        <div>
            <table style="border: thin solid #000050; height: 31px;" runat="server" id="table" visible="false">
                <tr>
                    <td class="stytitle_1">
                        <asp:Label ID="ld_yyyy" runat="server" Font-Bold="True" Style="text-align: center; font-size: small" Text="[년](ex:1989)" Visible="False"></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox ID="txt_yyyy" runat="server" Style="background-color: #FFFF99" Visible="False" Width="60px" onchange ="checkFourNumber(this)"></asp:TextBox>
                    </td>
                    <td class="stytitle_1">
                        <asp:Label ID="ld_mm" runat="server" Font-Bold="True" Style="text-align: center; font-size: small" Text="[월]<br>(ex:03)" Visible="False"></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox ID="txt_mm" runat="server" Style="background-color: #FFFF99" Width="60px" Visible="False" onchange ="checkTooNumber_m(this)"></asp:TextBox>
                    </td>
                    <td class="stytitle_1" runat="server" id="td">
                        <asp:Label ID="ld_dd" runat="server" Font-Bold="True" Style="text-align: center; font-size: small" Text="[일]<br>(ex:07)" Visible="False"></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox ID="txt_dd" runat="server" Style="background-color: #FFFF99" Width="60px" Visible="False" onchange ="checkTooNumber_d(this)"></asp:TextBox>
                    </td>
<%--                    <td>
                        <asp:CheckBox ID ="ckb_hld" runat = "server" Text ="공휴일체크" Font-Bold="True" Font-Size="Small"></asp:CheckBox>
                    </td>--%>
                    <td class="auto-style1">
                        <asp:Button ID="Select_Button" runat="server" BackColor="#FFFFCC" Font-Bold="True" Font-Size="Small" Height="26px" OnClick="btn_Select_Click" Text="조 회" Width="54px" Visible="False" />
                    </td>
                    <td class="auto-style1">
                        <asp:Button ID="Save_Button" runat="server" BackColor="#FFFFCC" Font-Bold="True" Font-Size="Small" Height="26px" OnClick="btn_Save_Click" Text="저 장" Width="54px" Visible="False" />
                    </td>
                    <td class="auto-style1">
                        <asp:Button ID="Update_Button" runat="server" BackColor="#FFFFCC" Font-Bold="True" Font-Size="Small" Height="26px" OnClick="btn_Update_Click" Text="수 정" Width="54px" Visible="False" />
                    </td>
                </tr>
            </table>
        </div>
        <%--입력 숨김창 div--%>


        <asp:UpdatePanel ID="Panel_List_View" runat="server" UpdateMode="Conditional">
            <ContentTemplate>

                <div id="div_day_spread" runat="server" visible="false">
                        <tr>
                            <div id="top_title_1" style="overflow: hidden; width: 850px">
                               <table class="tbl_list_1">
                                  <thead>
                                      <tr>
                                          <th style="width: 19px;">&nbsp;<asp:Label ID="lb_hd_no" runat="server" Height="15px" Text="NO" Width="16px"></asp:Label>
                                          </th>
                                          <th style="width: 152px;">
                                              <asp:Label ID="lb_hd_list" runat="server" Height="15px" Text="항  목"></asp:Label>
                                          </th>
                                          <th style="width: 115px;">&nbsp;<asp:Label ID="lb_hd_amt" runat="server" Height="15px" Text="금  액"></asp:Label>
                                          </th>
                                      </tr>
                                  </table>
                               </table>
                             </div>
                        </tr>
                        <div id="mainDisplay_1" style="width: 355px; height: 570px; overflow-y: scroll; overflow-x: hidden" >
                          <table class="tbl_list_2">
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no1" runat="server" Text="1." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list1" runat="server" Text="매출대전회수" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt1" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no2" runat="server" Text="2." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list2" runat="server" Text="부가세 환급" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt2" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no3" runat="server" Text="3." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list3" runat="server" Text="관세 환급" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt3" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no4" runat="server" Text="4." CssClass="style2"></asp:Label>.</td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list4" runat="server" Text="임대보증금/임대료 입금" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt4" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no5" runat="server" Text="5." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list5" runat="server" Text="수입이자 입" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt5" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no06" runat="server" Text="6." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list6" runat="server" Text="불용자재 매각" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt6" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no7" runat="server" Text="7." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list7" runat="server" Text="기   타" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt7" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no8" runat="server" Text="8." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list8" runat="server" Text="영업활동상의 자금수입" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt8" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no9" runat="server" Text="9." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list9" runat="server" Text="원자재 매입대금 지급" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt9" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no10" runat="server" Text="10." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list10" runat="server" Text="급여와 상여" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt10" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no11" runat="server" Text="11." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list11" runat="server" Text="퇴직금의 지급" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt11" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no12" runat="server" Text="12." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list12" runat="server" Text="원천제세 납부" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt12" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no13" runat="server" Text="13." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list13" runat="server" Text="법정복리비 납부" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt13" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no14" runat="server" Text="14." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list14" runat="server" Text="인건비의 지급" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt14" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no15" runat="server" Text="15." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list15" runat="server" Text="경 비" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt15" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no16" runat="server" Text="16." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list16" runat="server" Text="O/S(인건비,제경비)" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt16" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no17" runat="server" Text="17." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list17" runat="server" Text="외주가공비" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt17" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no18" runat="server" Text="18." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list18" runat="server" Text="식대/통근/임차료/전력" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt18" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no19" runat="server" Text="19." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list19" runat="server" Text="지급이자" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt19" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no20" runat="server" Text="20." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list20" runat="server" Text="임차보증금" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt20" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no21" runat="server" Text="21." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list21" runat="server" Text="기 타" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt21" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no22" runat="server" Text="22." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list22" runat="server" Text="영업활동상의 자금지출" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt22" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no23" runat="server" Text="23." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list23" runat="server" Text="영업상의 순 자금유입" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt23" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no24" runat="server" Text="24." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list24" runat="server" Text="고정자산 매각" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt24" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no25" runat="server" Text="25." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list25" runat="server" Text="대여금 회수" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt25" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no26" runat="server" Text="26." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list26" runat="server" Text="장기금융상품의 회수" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt26" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no27" runat="server" Text="27." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list27" runat="server" Text="투자유가증권/출자금 회수" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt27" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no28" runat="server" Text="28." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list28" runat="server" Text="투자활동의 Cash in" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt28" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no29" runat="server" Text="29." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list29" runat="server" Text="토지/건물/구축물 취득" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt29" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no30" runat="server" Text="30." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list30" runat="server" Text="기계장치 취득" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt30" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no31" runat="server" Text="31." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list31" runat="server" Text="차량 및 공기구 취득" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt31" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no32" runat="server" Text="32." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list32" runat="server" Text="고정자산의 취득" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt32" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no33" runat="server" Text="33." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list33" runat="server" Text="대여금 지급" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt33" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no34" runat="server" Text="34." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list34" runat="server" Text="장기금융상품의 매입" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt34" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no35" runat="server" Text="35." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list35" runat="server" Text="투자유가증권/출자금 지급" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt35" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no36" runat="server" Text="36." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list36" runat="server" Text="투자활동의 Cash out" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt36" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no37" runat="server" Text="37." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list37" runat="server" Text="투자활동의 자금흐름" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt37" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no38" runat="server" Text="38." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list38" runat="server" Text="신규 차입" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt38" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no39" runat="server" Text="39." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list39" runat="server" Text="USANCE 차입" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt39" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no40" runat="server" Text="40." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list40" runat="server" Text="증자 등" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt40" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no41" runat="server" Text="41." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list41" runat="server" Text="재무활동의 Cash in" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt41" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no42" runat="server" Text="42." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list42" runat="server" Text="차입금의 상환" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt42" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no43" runat="server" Text="43." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list43" runat="server" Text="USANCE 상환" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt43" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Silver;">&nbsp<asp:Label ID="lb_no44" runat="server" Text="44." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Silver">&nbsp<asp:Label ID="lb_list44" runat="server" Text="배당금/자기주식 취득 외" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt44" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no45" runat="server" Text="45." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list45" runat="server" Text="재무활동의 Cash Out" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt45" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: Khaki;">&nbsp<asp:Label ID="lb_no46" runat="server" Text="46." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: Khaki">&nbsp<asp:Label ID="lb_list46" runat="server" Text="재무활동의 현금 흐름" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt46" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: #CCFF66;">&nbsp<asp:Label ID="lb_no47" runat="server" Text="47." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: #CCFF66">&nbsp<asp:Label ID="lb_list47" runat="server" Text="전월이월 자금" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: #CCFF66">&nbsp<asp:TextBox ID="tb_amt47" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: #CCFF66;">&nbsp<asp:Label ID="lb_no48" runat="server" Text="48." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: #CCFF66">&nbsp<asp:Label ID="lb_list48" runat="server" Text="자금의 과부족" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: #CCFF66">&nbsp<asp:TextBox ID="tb_amt48" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="background-color: #CCFF66;">&nbsp<asp:Label ID="lb_no49" runat="server" Text="49." CssClass="style2"></asp:Label></td>
                                    <td class="auto-style2" style="background-color: #CCFF66">&nbsp<asp:Label ID="lb_list49" runat="server" Text="차월이월 자금" CssClass="style2"></asp:Label></td>
                                    <td style="background-color: #CCFF66">&nbsp<asp:TextBox ID="tb_amt49" runat="server" Style="text-align: center" ReadOnly="True" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                </tr>
                             </table>
                           </div>
                </div>
                <%--일 자금실적 입력 항목--%>

                <div id="div_month_spread" runat="server" visible="false">
                    <tr>
                        <div id="top_title" style="overflow: hidden; width: 515px">
                            <table class="tbl_list_1">
                                <thead>
                                    <tr>
                                        <th style="width : 81px;">&nbsp;<asp:Label ID="lb_hd_no_2" runat="server" Height="15px" Text="항 목"></asp:Label>
                                        </th>
                                        <th style="width : 176px;">
                                            <asp:Label ID="lb_hd_list_2" runat="server" Height="15px" Text="세 목"></asp:Label>
                                        </th>
                                        <th style="width : 129px;">&nbsp;<asp:Label ID="lb_hd_amt_2" runat="server" Height="15px" Text="계획 금액"></asp:Label>
                                        </th>
                                        <th style="width : 129px;">&nbsp;<asp:Label ID="Label2" runat="server" Height="15px" Text="실적 금액"></asp:Label>
                                        </th>
                                    </tr>
                                </thead>
                            </table>
                        </div>
                    </tr>
                    <div id="mainDisplay" style="width: 545px; height: 570px; overflow-y: scroll; overflow-x: hidden" >
                        <table class="tbl_list_2">
                            <tr>
                                <td></td>
                                <td colspan="2" style="background-color: Silver">&nbsp;<asp:Label ID="lb_plan_1" runat="server" CssClass="style2" Text="매출채권"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_1" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_RSLT_1" runat="server" onKeyPress="currency(this)" onKeyup="com(this)" Style="text-align: center" value="0" Width="106px"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="text-align: center" class="auto-style16">&nbsp;</td>
                                <td style="text-align: center; background-color: Silver;"><asp:Label ID="lb_plan_2" runat="server" CssClass="style2" Text="기타"></asp:Label></td>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_2" runat="server" Text="제세환급외" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_2" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_2" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: Khaki">&nbsp<asp:Label ID="lb_plan_3" runat="server" CssClass="style2" Text="영업현금유입"></asp:Label></td>
                                <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt_plan_3" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: Khaki">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_3" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16">&nbsp;</td>
                                <td class="auto-style2" colspan="2" style="background-color: Silver">&nbsp;<asp:Label ID="lb_plan_4" runat="server" CssClass="style2" Text="매입채무"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_4" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_4" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td colspan="3" style="background-color: Khaki">&nbsp;<asp:Label ID="lb_plan_5" runat="server" CssClass="style2" Text="미지급금"></asp:Label></td>
                                <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt_plan_5" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: Khaki">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_5" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="text-align: center" rowspan="6" class="auto-style16">&nbsp;</td>
                                <td rowspan="6" style="text-align: center">&nbsp;</td>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_6" runat="server" Text="O/S 인건비/제비용" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_6" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_6" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_7" runat="server" Text="외주가공비/용역비" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_7" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_7" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_8" runat="server" Text="소모/수선/수수료/운반/통관" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_8" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_8" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_9" runat="server" Text="식대/통근/전력/광열/임차" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_9" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_9" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_10" runat="server" Text="경상연구 재료비" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_10" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_10" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_11" runat="server" Text="기타경비" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_11" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_11" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="text-align: center" class="auto-style16">&nbsp;</td>
                                <td style="text-align: center; background-color: Silver;"><asp:Label ID="lb_plan_11" runat="server" CssClass="style2" Text="인건비"></asp:Label></td>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_12" runat="server" Text="자사인건비/4대보험/원천세" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_12" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_12" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="text-align: center" class="auto-style16">&nbsp;</td>
                                <td style="text-align: center; background-color: Silver;">&nbsp<asp:Label ID="lb_plan_13" runat="server" CssClass="style2" Text="지급이자"></asp:Label></td>
                                <td class="auto-style14" style="background-color: Silver">&nbsp<asp:Label ID="lb_list_plan_13" runat="server" Text="지급이자" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_13" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_13" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: Khaki">&nbsp<asp:Label ID="lb_plan_14" runat="server" CssClass="style2" Text="영업현금유출"></asp:Label></td>
                                <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt_plan_14" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: Khaki">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_14" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: Khaki">&nbsp<asp:Label ID="lb_plan_15" runat="server" CssClass="style2" Text="영업현금수지"></asp:Label></td>
                                <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt_plan_15" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: Khaki">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_15" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16">&nbsp</td>
                                <td colspan="2" style="background-color: Silver">&nbsp;<asp:Label ID="lb_plan_16" runat="server" CssClass="style2" Text="고정자산"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_16" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_16" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="text-align: center" class="auto-style16">&nbsp;</td>
                                <td style="background-color: Silver;"><asp:Label ID="lb_plan_17" runat="server" CssClass="style2" Text="기타"></asp:Label></td>
                                <td class="auto-style14" style="background-color: Silver"><asp:Label ID="lb_list_plan_17" runat="server" Text="기타영업외유출" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_17" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_17" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="text-align: center" class="auto-style16">&nbsp;</td>
                                <td style="background-color: Silver;"><asp:Label ID="lb_plan_18" runat="server" CssClass="style2" Text="기타"></asp:Label></td>
                                <td class="auto-style14" style="background-color: Silver"><asp:Label ID="lb_list_plan_18" runat="server" Text="기타영업외유입" CssClass="style2"></asp:Label></td>
                                <td style="background-color: Silver">&nbsp<asp:TextBox ID="tb_amt_plan_18" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                                <td style="background-color: Silver">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_18" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: Khaki">&nbsp<asp:Label ID="lb_plan_19" runat="server" CssClass="style2" Text="영업외현금유출"></asp:Label></td>
                                <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt_plan_19" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: Khaki">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_19" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: Khaki">&nbsp<asp:Label ID="lb_plan_20" runat="server" CssClass="style2" Text="당월총현금수지"></asp:Label></td>
                                <td style="background-color: Khaki">&nbsp<asp:TextBox ID="tb_amt_plan_20" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: Khaki">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_20" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: #CCFF66">&nbsp<asp:Label ID="lb_plan_21" runat="server" CssClass="style2" Text="기초현금"></asp:Label></td>
                                <td style="background-color: #CCFF66">&nbsp<asp:TextBox ID="tb_amt_plan_21" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: #CCFF66">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_21" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: #CCFF66">&nbsp<asp:Label ID="lb_plan_22" runat="server" CssClass="style2" Text="기말현금"></asp:Label>
                                </td>
                                <td style="background-color: #CCFF66">&nbsp<asp:TextBox ID="tb_amt_plan_22" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: #CCFF66">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_22" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="auto-style16" colspan="3" style="background-color: #CCFF66">&nbsp<asp:Label ID="lb_plan_23" runat="server" CssClass="style2" Text="가용현금"></asp:Label></td>
                                <td style="background-color: #CCFF66">&nbsp<asp:TextBox ID="tb_amt_plan_23" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                                <td style="background-color: #CCFF66">&nbsp;<asp:TextBox ID="tb_amt_plan_RSLT_23" runat="server" Style="text-align: center" onKeyPress="currency(this)" onKeyup="com(this)" Width="106px" ReadOnly="true" value="0"></asp:TextBox></td>
                            </tr>
                        </table>
                    </div>
                </div>
                <%--월 자금계획 입력 항목--%>
            </ContentTemplate>

        </asp:UpdatePanel>


    </form>
</body>
</html>
