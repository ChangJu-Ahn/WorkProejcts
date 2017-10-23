<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MM_MM001.aspx.cs" EnableEventValidation="false" Inherits="ERPAppAddition.ERPAddition.MM.MM_MM001.MM_MM001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>구매요청 연동데이터 상태조회</title>
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
            width: 70px;
            white-space: nowrap;
        }

        .ui-progressbar {
            position: relative;
        }
    </style>
    <script type="text/javascript">

        function PopOpenUpdate(pr_No) {
            alert(pr_No);

            return;
        }

        function OutputAlert(content) {
            alert("아래 내용을 관리자에게 문의하세요. \n * 내용 : [" + content + "]");
            return;
        }

    </script>
</head>
<body>
    <form id="frm1" runat="server">
        <cc1:ToolkitScriptManager runat="Server" ID="ToolkitScriptManager1" />
        <div style="height: 11px;">
        </div>
        <table>
            <tr>
                <th>
                    <asp:Image runat="server" ImageUrl="~/img/folder.gif" /></th>
                <th class="title">구매요청 연동데이터 상태조회</th>
            </tr>
            <tr>
                <th colspan="2">
                    <span style="color: blue">(현재 접속중인 ERP는 [&nbsp;<asp:Label runat="server" ID="lblerpName"></asp:Label>&nbsp;] 입니다)</span>
                </th>
            </tr>
        </table>
        <div>
            <table class="BasicTb">
                <tr>
                    <td class="tilte">공장</td>
                    <td>
                        <asp:DropDownList ID="ddl_Plant" runat="server" BackColor="Yellow">
                        </asp:DropDownList>
                    </td>
                    <td class="tilte">구매요청자</td>
                    <td>
                        <asp:TextBox runat="server" ID="txtReqPrsn" Width="60"></asp:TextBox>
                    </td>
                    <td class="tilte">문서번호</td>
                    <td>
                        <asp:TextBox runat="server" ID="txtDocument" Width="240"></asp:TextBox>
                    </td>
                    <td class="tilte">조회기간</td>
                    <td>
                        <asp:TextBox ID="txtdate_From" runat="server" MaxLength="10" Width="68px" CssClass="required"></asp:TextBox>
                        <cc1:CalendarExtender ID="cal_From" runat="server" Enabled="True"
                            Format="yyyy-MM-dd" TargetControlID="txtdate_From">
                        </cc1:CalendarExtender>
                    </td>
                    <td style="font-size: medium; font-weight: 700;">~</td>
                    <td>
                        <asp:TextBox ID="txtdate_To" runat="server" MaxLength="10" Width="68px" CssClass="required"></asp:TextBox>
                        <cc1:CalendarExtender ID="cal_to" runat="server" Enabled="True"
                            Format="yyyy-MM-dd" TargetControlID="txtdate_To">
                        </cc1:CalendarExtender>
                    </td>
                    <td class="tilte">MRO여부</td>
                    <td>
                        <asp:DropDownList ID="ddlMROFlag" runat="server">
                            <asp:ListItem Value="N" Selected="True">N</asp:ListItem>
                            <asp:ListItem Value="Y">Y</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td class="tilte">상태</td>
                    <td>
                        <asp:DropDownList ID="ddlFlagType" runat="server">
                            <asp:ListItem Value="A" Selected="True">전체</asp:ListItem>
                            <asp:ListItem Value="Y">완료</asp:ListItem>
                            <asp:ListItem Value="E">오류</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td>
                        <asp:Button runat="server" ID="query" Text="조회" OnClick="btnSelect_Click" Width="50px" />
                        <%--<asp:Button runat="server" ID="excel" Text="엑셀변환" OnClick="btnExcel_Click" Width="60px" />--%>
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
                <div style="margin-bottom: 10px; font-weight: 700;">
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
                                <HeaderStyle Width="50px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    수정
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="250px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    전자결재문서번호
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="130px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    구매요청번호
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="70px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    공장코드
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    품목코드
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>

                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="250px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    품목이름
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>

                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="90px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    수량
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="50px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    단위
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    구매요청일
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    입고희망일
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="80px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    구매요청자
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="70px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    처리상태
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="250px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    오류내용
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
