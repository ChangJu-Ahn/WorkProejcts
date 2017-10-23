<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AhnTest.aspx.cs" Inherits="ERPAppAddition.ERPAddition.AhnTest" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Label runat="server" ID="lblCategory">카테고리</asp:Label><asp:DropDownList runat="server" ID="ddlList" />
        </div>
        <div>
            <asp:Label runat="server" ID="lblSubject">제목</asp:Label><asp:TextBox runat="server" ID="txtContent"></asp:TextBox>
        </div>
        <div>
            <asp:Button runat="server" ID="btnSearch" Text="조회" OnClick="btnSearch_Click" />
        </div>
        <div>
            <asp:GridView runat="server" ID="gridView" OnRowDataBound="gridView_ItemDataBound" AllowPaging="true" PageSize="20" OnPageIndexChanging="gridView_PageIndexChanging" AutoGenerateColumns="false">
                <HeaderStyle HorizontalAlign="Center" />
                <Columns>
                    <asp:TemplateField>
                        <HeaderStyle Width="150px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                        <HeaderTemplate>
                           문서키
                        </HeaderTemplate>
                    </asp:TemplateField>
                </Columns>
                <Columns>
                    <asp:TemplateField>
                        <HeaderStyle Width="200px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                        <HeaderTemplate>
                            문서번호
                        </HeaderTemplate>
                    </asp:TemplateField>
                </Columns>
                <Columns>
                    <asp:TemplateField>
                        <HeaderStyle Width="400px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                        <HeaderTemplate>
                            제목
                        </HeaderTemplate>
                    </asp:TemplateField>
                </Columns>
                <Columns>
                    <asp:TemplateField>
                        <HeaderStyle Width="200px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                        <HeaderTemplate>
                            작성일
                        </HeaderTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </div>
    </form>
<%--    <script type="text/javascript">
        function on_view1(c_no, dsm_yn) {
            var url;
            dsm_yn = escape(dsm_yn);
            //url = "da072e.asp?control_no=" + c_no + "&dsm_yn=" + dsm_yn + "&dsm_del=Y";
            url = "http://ekp.nepes.co.kr/decision/da072e.asp?control_no=" + c_no + "&dsm_yn=" + dsm_yn + "&dsm_del=Y";
            gb11_win = window.open(url, "gb11_win", "width=670,height=567,top=220,left=100,scrollbars=1");
        }
    </script>--%>
</body>
</html>
