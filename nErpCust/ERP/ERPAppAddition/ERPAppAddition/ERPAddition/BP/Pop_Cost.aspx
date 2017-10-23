<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Pop_Cost.aspx.cs" Inherits="ERPAppAddition.ERPAddition.BP.Pop_Cost" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>물류비용조회(거래처조회)</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="../../../Styles/Site_display.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css" />
    <script type="text/javascript" src="//code.jquery.com/jquery-1.10.2.js"></script>
    <script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
    <link rel="stylesheet" href="/resources/demos/style.css" />
    <style type="text/css">
        .BasicTb {
            width: auto;
            border: thin double #000080;
        }

        td.tilte {
            background-color: #99CCFF;
            font-weight: bold;
            text-align: center;
            width: 60px;
            white-space: nowrap;
        }

        .ui-progressbar {
            position: relative;
        }
    </style>
    <script type="text/javascript">

        window.onbeforeunload = function () {
            closePopUp();
        }


        function OutputAlert(content) {
            alert("아래 내용을 관리자에게 문의하세요. \n * 내용 : [" + content + "]");
            return;
        }

        function PopDateDeliver(BP_CD, BP_NM) {

           
            if (BP_CD != null) {
                window.returnValue = BP_CD + ";" + BP_NM;
                document.getElementById('txtPartnerCD').value = BP_CD;
            }

            self.close();
        }

        
        function closePopUp() {

            var item = document.getElementById('txtPartnerCD').value;

            if (item == "") {
                window.returnValue = ";";
            }
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <cc1:ToolkitScriptManager runat="Server" ID="ToolkitScriptManager1" />
        <div style="height: 20px;">
        </div>
        <div style="padding-left: 10px;">
            <table class="BasicTb">
                <tr>
                    <td class="tilte">거래처코드</td>
                    <td>
                        <asp:TextBox runat="server" ID="txtPartnerCD" Width="100"></asp:TextBox>
                    </td>
                    <td class="tilte">거래처명</td>
                    <td>
                        <asp:TextBox runat="server" ID="txtPartnerNm" Width="100"></asp:TextBox>
                    </td>
                    <td class="tilte">거래구분</td>
                    <td>
                        <asp:RadioButton runat="server" id ="rdoAll" Text="전체" GroupName="rdoType" Checked="true" />
                        <asp:RadioButton runat="server" id="rdoCS" Text="매출/매입" GroupName="rdoType" />
                        <asp:RadioButton runat="server" id="rdoS" Text="매입" GroupName="rdoType" />
                        <asp:RadioButton runat="server" id="rdoC" Text="매출" GroupName="rdoType" />
                    </td>
                    <td>
                        <asp:Button runat="server" ID="query" Text="조회" OnClick="btnSelect_Click" Width="50px" />
                    </td>
                </tr>
            </table>
        </div>
        <div style="height: 10px;">
        </div>
        <asp:UpdateProgress ID="UpdateProgress1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <div style="padding-left: 230px; padding-top: 50px;">
                    <asp:Image ID="Image3_1" runat="server" CssClass="updateProgress" ImageUrl="~/img/loading_spinner.gif" Height="173px" Width="230px" />
                </div>
            </ProgressTemplate>
        </asp:UpdateProgress>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <div style="margin-bottom: 10px; font-weight: 700; padding-left: 10px;">
                    <span>(총 조회건수 :
                    <asp:Label ID="lblListCnt" runat="server" />
                        건)</span>
                </div>
                <div style="padding-left: 10px;">
                    <asp:GridView runat="server" CssClass="BasicTb" ID="dgList" OnRowDataBound="dgList_RowDataBound" AutoGenerateColumns="false"
                        AllowPaging="true" PageSize="20" OnPageIndexChanging="dgList_PageIndexChanging">
                        <HeaderStyle HorizontalAlign="Center" />
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="120px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    거래처코드
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="330px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    거래처명
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="150px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    사업자등록번호
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
            </ContentTemplate>
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="query"></asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
    </form>
</body>
</html>
