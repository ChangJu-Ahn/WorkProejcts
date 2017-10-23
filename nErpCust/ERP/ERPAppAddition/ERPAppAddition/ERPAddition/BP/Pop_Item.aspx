<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Pop_Item.aspx.cs" Inherits="ERPAppAddition.ERPAddition.BP.Pop_Item" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="../../../Styles/Site_display.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css" />
    <script type="text/javascript" src="//code.jquery.com/jquery-1.10.2.js"></script>
    <script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
    <link rel="stylesheet" href="/resources/demos/style.css" />
    <style type="text/css">
        .BasicTb {
            width: 100%;
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
        .auto-style1 {
            height: 23px;
        }
    </style>
    <script type="text/javascript">

        var Item_cd = "";
        var Item_nm = "";
        var txt_Item_cd = "";
        var txt_Item_nm = "";

        //IitiPopUp();

        window.onload = function () {
            IitiPopUp();
        }
        window.onbeforeunload = function () {
            closePopUp();
        }

        
        function IitiPopUp()
        {
            
            //window.Item_cd = window.dialogArguments["Item_cd"];
            //window.Item_nm = window.dialogArguments["Item_nm"];
            //window.txt_Item_cd = window.dialogArguments["txt_Item_cd"];
            //window.txt_Item_nm = window.dialogArguments["txt_Item_nm"];

            //document.getElementById("lblitem_cd").Text = window.txt_Item_cd;
            //var aa = document.getElementById("txtItem_cd")
            // aa.Text = window.Item_cd;
            //document.getElementById('lblItem_nm').Text = window.txt_Item_nm;
            //document.getElementById('txtItem_nm').Text = window.Item_nm;
            //alert(window.txt_Item_cd);
            //alert(aa);

            //var text1 = window.txt_Item_cd;
            //var text2 = window.txt_Item_nm;

            //document.getElementById('lblitem_cd').innerHTML = text1;
            //document.getElementById('lblItem_nm').innerHTML = text2;
            
        }

        function closePopUp() {

            var item = document.getElementById('txtItem_cd').value;

            if (item == "") {
                window.returnValue = ";";
            }
        }

        function OutputAlert(content) {
            alert("아래 내용을 관리자에게 문의하세요. \n * 내용 : [" + content + "]");
            return;
        }

        function PopDateDeliver(BP_CD, BP_NM) {

           
            if (BP_CD != null) {
                window.returnValue = BP_CD + ";" + BP_NM;

                document.getElementById('txtItem_cd').value = BP_CD;
            }



            self.close();
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
                    <td class="tilte" style="height: 23px">
                        <asp:Label ID="lblitem_cd" runat="server" Text=""></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox runat="server" ID="txtItem_cd" Width="90%" Text=""></asp:TextBox>
                    </td>
                    <td class="tilte" style="height: 23px">
                        <asp:Label ID="lblItem_nm" runat="server" Text=""></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox runat="server" ID="txtItem_nm" Width="90%"  Text=""></asp:TextBox>
                    </td>
                    <td style="width:70px">
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
                    <asp:GridView runat="server" CssClass="BasicTb" ID="dgList" OnRowDataBound="dgList_RowDataBound" AutoGenerateColumns="False"
                        AllowPaging="True" PageSize="20" OnPageIndexChanging="dgList_PageIndexChanging">
                        <HeaderStyle HorizontalAlign="Center" />
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="330px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    ITEM_CD
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="150px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    ITEM_NM
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
