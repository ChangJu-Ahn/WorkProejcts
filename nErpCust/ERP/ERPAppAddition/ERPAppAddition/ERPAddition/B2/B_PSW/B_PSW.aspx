<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="B_PSW.aspx.cs" Inherits="ERPAppAddition.ERPAddition.B2.B_PSW.B_PSW" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <style type="text/css">
        .auto-style45 {
            width: 90px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
        }

        .style1 {
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            table-layout: fixed;
        }

        .auto-style102 {
            width: 90px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            text-align: center;
            font-size: smaller;
            table-layout: fixed;
            height: 15px;
        }

        .auto-style47 {
            font-size: small;
            height: 15px;
        }

        .auto-style104 {
            font-weight: bold;
            font-family: 굴림체;
            text-align: left;
            table-layout: fixed;
            height: 25px;
        }
    </style>
    <script type="text/javascript">
        function uPageClose() {
            alert("사용자 등록 mail로 임시 비밀번호가 발급 되었습니다. Mail로 발송된 임시 비밀번호로 로그인 바랍니다. [Mail발송 대기시간 약 5분]");

            window.opener = 'nothing';
            window.open('', '_parent', '');
            window.close();
        }
    </script>
</head>
<body onblur="window.focus()">
    <form id="form1" runat="server">
        <div>

            <table>
                <tr>
                    <td class="auto-style45">계열사</td>
                    <td class="auto-style104">
                        <asp:DropDownList ID="ddl_fac" runat="server" Width="250px" AutoPostBack="false">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style102">ID :</td>
                    <td class="auto-style47">
                        <asp:TextBox ID="txt_id" runat="server" Width="250px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style102">E-Mail : </td>
                    <td class="auto-style47">
                        <asp:TextBox ID="txt_mail" runat="server" Width="250px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="text-align: center">
                        <asp:Button ID="send" runat="server" Width="170px" Text="임시 비밀번호 발급" OnClick="send_Click" Height="30px" />
                        <%--<asp:Button ID="end"  runat="server" Width="170px" Text="창닫기" OnClick="exit_Click" Height="30px"/>--%>
                        <asp:Button ID="end" runat="server" Width="170px" Text="창닫기" OnClientClick="window.close(); return false;" Height="30px" Style="margin-bottom: 0px" />
                    </td>
                </tr>
            </table>

            <table>
                <tr>
                    <td>1. 입력한 id 와 E-mail 주소를 확인합니다.<br />
                        2. 해당 E-mail 주소로 임시 비밀번호가 발급됩니다.<br />
                        3. 임시비밀번호로 로그인후 변경합니다.<br />
                        <br />
                        ※계열사는 정보시스템그룹으로 문의 바랍니다.<br />
                    </td>
                </tr>
            </table>

        </div>
    </form>
</body>
</html>
