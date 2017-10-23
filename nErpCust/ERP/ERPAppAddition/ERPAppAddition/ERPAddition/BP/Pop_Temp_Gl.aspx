<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Pop_Temp_Gl.aspx.cs" Inherits="ERPAppAddition.ERPAddition.BP.Pop_Temp_Gl" %>

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
            width: 600px;
            border: thin double #000080;
        }

        td.tilte {
            background-color: #99CCFF;
            font-weight: bold;
            text-align: center;
            width: 70px;
            white-space: nowrap;
        }
        td.txtRight
        {
            text-align: right;
        }

        .ui-progressbar {
            position: relative;
        }
        .auto-style2 {
            width: 194px;
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


        function IitiPopUp() {

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

            var tmp_gl = document.getElementById('txtTempGL_NO').value;

            if(tmp_gl == "")
            {
                window.returnValue = "";
            }
        }

        function OutputAlert(content) {
            alert("아래 내용을 관리자에게 문의하세요. \n * 내용 : [" + content + "]");
            return;
        }

        function PopDateDeliver(Temp_GL) {


            if (Temp_GL != null) {
                window.returnValue = Temp_GL;

                document.getElementById('txtTempGL_NO').value = Temp_GL;
            }

            self.close();
        }

        function ClickBtnDept() {

            var col = "부서코드,부서명"

            var d = new Date();

            var YYYYMMDD = d.getFullYear() + "-" + (d.getMonth()+1) + "-" + d.getDate();


            var sql = " SELECT * FROM (SELECT RTRIM(DEPT_CD) AS ITEM_CD ,DEPT_NM AS ITEM_NM FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT = ( SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= '"+YYYYMMDD+"' )  ) AA"

            var item_cd = document.getElementById('txtDept_cd').value;

            var Retval = fn_GetItem(col, sql, item_cd)

            if (Retval != null) {

                var Item_cd = Retval.split(";")[0];
                var Item_nm = Retval.split(";")[1];

                document.getElementById('txtDept_cd').value = Item_cd;
                    document.getElementById('txtDetp_nm').value = Item_nm;
            }

            return false;
        }

        function ClickBtnBizArea() {

            var col = "사업장코드,사업자명"

            var sql = " SELECT * FROM (SELECT DISTINCT BIZ_AREA_CD AS ITEM_CD,BIZ_AREA_NM AS ITEM_NM FROM   B_BIZ_AREA WHERE  BIZ_AREA_CD>= '' ) AA"

            var item = document.getElementById('txtBizArea_CD').value;
            var Retval = fn_GetItem(col, sql, item)

            if (Retval != null) {

                var Item_cd = Retval.split(";")[0];
                var Item_nm = Retval.split(";")[1];

                document.getElementById('txtBizArea_CD').value = Item_cd;
                document.getElementById('txtBizArea_NM').value = Item_nm;
            }

            return false;
        }

        function ClickBtnUser() {

            var col = "작성자코드,작성자명"

            var sql = " SELECT * FROM (SELECT DISTINCT USR_ID AS ITEM_CD ,USR_NM AS ITEM_NM FROM Z_USR_MAST_REC WHERE USR_ID >= '' ) AA"

            var item_cd = document.getElementById('txtUser_ID').value;
            var Retval = fn_GetItem(col, sql, item_cd)

            if (Retval != null) {

                var Item_cd = Retval.split(";")[0];
                var Item_nm = Retval.split(";")[1];

                document.getElementById('txtUser_ID').value = Item_cd;
                document.getElementById('txtUser_NM').value = Item_nm;
            }

            return false;
        }

        function fn_GetItem(columns, sql, item_cd) {
            var PopWidth = 535;
            var PopHeight = 520;

            var dbGubun = document.getElementById('hdfDbName').value;
            var PopNodeUrl = "../BP/Pop_Item.aspx?dbName=" + dbGubun;
            var PopFont = "FONT-FAMILY: '맑은고딕';font-size:15px;";
            var PopParams = new Array(); //별도의 넘길 값은 없으나 형식에 맞추기 위해 배열객체만 선언


            var columnsList = "&columns=" + columns;
            var sqlQuery = "&sql=" + sql;
            var item = "&itemcd=" + item_cd ;


            PopNodeUrl += columnsList + sqlQuery + item;

            PopNodeUrl = encodeURI(PopNodeUrl);

            var Retval = self.showModalDialog(PopNodeUrl, PopParams, PopFont + "dialogHeight:" + PopHeight + "px;dialogWidth:" + PopWidth + "px;resizable:no;status:no;help:no;scroll:no;location:no");


            return Retval
        }
       

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <cc1:ToolkitScriptManager runat="Server" ID="ToolkitScriptManager1" />
        <asp:HiddenField ID="hdfDbName" runat="server" />
        <div style="height: 20px;">
                        
        </div>
        <div style="padding-left: 10px;">
            <table style="width: 600px;">
                <tr>
                    <td class="tilte" style="height: 23px">
                        <asp:Label ID="Label6" runat="server" Text="전표번호"></asp:Label>
                    </td>
                    <td class="auto-style2">
                        <asp:TextBox runat="server" ID="txtTempGL_NO" Width="164px"  Text=""></asp:TextBox>
                    </td>
                     <td class="txtRight"> 
                        <asp:Button runat="server" ID="query" Text="조회" OnClick="btnSelect_Click" Width="50px" /> 
                    </td>
                </tr>
            </table>
            <table class="BasicTb">
                <tr>
                    <td class="tilte" style="height: 23px">
                        <asp:Label ID="lblDT" runat="server" Text="결의일자"></asp:Label>
                    </td>
                    <td class="auto-style2">
                        <asp:TextBox ID="tb_fr_yyyymmdd" runat="server" MaxLength="8" Width="71px"></asp:TextBox>
                        <cc1:CalendarExtender ID="tb_fr_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                            Format="yyyyMMdd" TargetControlID="tb_fr_yyyymmdd">
                        </cc1:CalendarExtender>
                        <asp:Label ID="Label1" runat="server" Text=" ~ "></asp:Label>
                        <asp:TextBox ID="tb_to_yyyymmdd" runat="server" MaxLength="8" Width="71px"></asp:TextBox>
                        <cc1:CalendarExtender ID="tb_to_yyyymmdd_CalendarExtender" runat="server" Enabled="True"
                            Format="yyyyMMdd" TargetControlID="tb_to_yyyymmdd">
                        </cc1:CalendarExtender>
                    </td>
                    <td class="tilte" style="height: 23px">
                        <asp:Label ID="lblItem_nm" runat="server" Text="부서코드"></asp:Label>
                    </td>
                    <td class="auto-style2">
                        <asp:TextBox runat="server" ID="txtDept_cd" Width="60px"  Text=""></asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text=".." Width="16px" OnClientClick="return ClickBtnDept()" />
                        <asp:TextBox runat="server" ID="txtDetp_nm" Width="80px"  Text=""></asp:TextBox>
                    </td>
                  </tr>
                <tr>
                    <td class="tilte" style="height: 23px">
                        <asp:Label ID="Label2" runat="server" Text="사업장"></asp:Label>
                    </td>
                    <td class="auto-style2">
                        <asp:TextBox runat="server" ID="txtBizArea_CD" Width="60px"  Text=""></asp:TextBox>
                        <asp:Button ID="btnBizArea" runat="server" Text=".." Width="16px"  OnClientClick="return ClickBtnBizArea()"/>
                        <asp:TextBox runat="server" ID="txtBizArea_NM" Width="80px"  Text=""></asp:TextBox>
                    </td>
                    <td class="tilte" style="height: 23px">
                         <asp:Label ID="Label3" runat="server" Text="작성자"></asp:Label>
                    </td>
                     <td>
                        <asp:TextBox runat="server" ID="txtUser_ID" Width="60px"  Text=""></asp:TextBox>
                        <asp:Button ID="Button3" runat="server" Text=".." Width="16px" OnClientClick="return ClickBtnUser()" />
                        <asp:TextBox runat="server" ID="txtUser_NM" Width="80px"  Text=""></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="tilte" style="height: 23px">
                        <asp:Label ID="Label4" runat="server" Text="승인상태"></asp:Label>
                    </td>
                    <td class="auto-style2">

                        <asp:RadioButton ID="rdoConf_Y" runat="server" Text="승인" Checked="True" />
                        <asp:RadioButton ID="rdoConf_N" runat="server" Text="미승인" />
                    </td>
                    <td  class="tilte" style="height: 23px">
                        <asp:Label ID="Label5" runat="server" Text="참조번호"></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox runat="server" ID="txtTempGl_Desc" Width="160px"  Text=""></asp:TextBox>
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
                <div style="margin-bottom: 10px; font-weight: 600; padding-left: 10px;">
                    <span>(총 조회건수 :
                    <asp:Label ID="lblListCnt" runat="server" />
                        건)</span>
                </div>
                <div style="padding-left: 10px; width:100% ; height:100%; overflow:scroll;">
                    <asp:GridView runat="server" CssClass="BasicTb" ID="dgList" OnRowDataBound="dgList_RowDataBound" AutoGenerateColumns="False"
                        AllowPaging="True" PageSize="20" OnPageIndexChanging="dgList_PageIndexChanging">
                        <HeaderStyle HorizontalAlign="Center" />
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="150px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    결의번호
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    부서명
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    결의일자
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    참조번호
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    작성자명
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    비고
                                </HeaderTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle Width="100px" HorizontalAlign="Center" VerticalAlign="Middle" ForeColor="White" BackColor="#003399" />
                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="false" />
                                <HeaderTemplate>
                                    금액
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