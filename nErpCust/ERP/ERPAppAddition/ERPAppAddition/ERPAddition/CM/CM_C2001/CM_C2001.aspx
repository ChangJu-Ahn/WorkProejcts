<%@ Page Language="C#" AutoEventWireup="TRUE" CodeBehind="CM_C2001.aspx.cs" Inherits="ERPAppAddition.ERPAddition.CM.CM_C2001.CM_C2001"  EnableEventValidation="true"  %>


<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>4M 데이타 조회</title>
    <style type="text/css">
        .style1
        {
            font-family: Arial;
            text-align: left;
            width: 300px;
        }
        .style2
        {
            font-family: Arial;
            background-color: #CCCCCC;
            font-size: small;
            text-align: center;
            width: 100px;
            height: 25px;
        }
        .style3
        {
            background-color: #FFFFCC;
        }
        .style4
        {
            font-family: "맑은 고딕";
            font-size: small;
        }
        .style6
        {
            font-size: small;
        }
    </style>
    <script type="text/javascript">
    function confirm_user()
    {
      if (confirm("Are you sure you want to go home ?")==true)
          return true;
      else
          return false;
  }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table style="border: thin solid #000080; width:100%;" >
            <tr>
                <td class="style2">
                    <asp:Label ID="Label1" runat="server" Text="조회월"></asp:Label>
                </td>
                <td class="style1">
                    <asp:TextBox ID="tb_fr_dt" runat="server" BackColor="#FFFFCC" Width="100px" 
                        MaxLength="6" CssClass="style3"></asp:TextBox>
                    <cc1:CalendarExtender ID="tb_fr_dt_CalendarExtender" runat="server" Enabled="True"
                        TargetControlID="tb_fr_dt" TodaysDateFormat="YYYY.MM" Format="yyyyMM">
                    </cc1:CalendarExtender>                   
                </td>
                <td class="style2">
                    <asp:Label ID="Label2" runat="server" Text="집계항목"></asp:Label>
                </td>
                <td class="style1">
                   
                </td>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="Label3" runat="server" Text="조회구분"></asp:Label>
                </td>
                <td class="style1">
                    <asp:RadioButtonList ID="RadioButtonList1" runat="server" 
                        RepeatDirection="Horizontal" CssClass="style4" AutoPostBack="True" 
                        onselectedindexchanged="RadioButtonList1_SelectedIndexChanged" 
                        TabIndex="10">
                        <asp:ListItem Selected="True" Value="view1">집계조회</asp:ListItem>
                        <asp:ListItem Value="view2">상세조회</asp:ListItem>
                        <asp:ListItem Value="view3">기준정보조회</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td class="style2">
                    <asp:Label ID="Label4" runat="server" Text="상세구분"></asp:Label>
                </td>
                <td class="style1">
                    <asp:DropDownList ID="DropDownList1" runat="server" TabIndex="20" 
                        onselectedindexchanged="DropDownList1_SelectedIndexChanged" 
                        AutoPostBack="True">
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
        <table style="width:100%;">
        <tr>
        <td >
            <asp:Button ID="btn_request" runat="server" Text="자료생성" 
                onclick="btn_request_Click" Width="100px" TabIndex="30" 
                onclientclick="return confirm('기존 데이타 존재시 삭제후 생성됩니다. 진행하시겠습니까?');" 
                ToolTip="조회월 자료를 삭제후 재생성한다." /></td>
                <td style="width:20px;"></td>

            <td>
                <asp:Button ID="btn_old_request" runat="server" Text="조회" 
                onclick="btn_old_request__Click" Width="100px" TabIndex="5" 
                    ToolTip="조회월 조회구분, 상세구분 내용을 조회한다." />
                </td>
                <td>
                <asp:Panel ID="Panel1" runat="server">
                    <asp:Label ID="lbl_bas_info_title" runat="server" Text="CostCenter 그룹 등록" CssClass="style6"></asp:Label>
                <asp:FileUpload ID="FileUpload1" runat="server" ToolTip="Import할 자료를 찾는다" />
                <asp:Button ID="btnUpload" runat="server" Text="Upload" 
                    onclick="btnUpload_Click" ToolTip="엑셀자료를 가져와 화면에 보여준다." />
                <asp:Button ID="btnMoveDataNextMonth" runat="server" Text="전월복사" 
                        onclick="btnMoveDataNextMonth_Click" 
                        onclientclick="return confirm('당월데이타 존재시 삭제후 복사됩니다. 진행하시겠습니까?');" 
                        ToolTip="전월 기준정보자료를 년월만 수정후 복사한다" />
                </asp:Panel>
                    
                <asp:Panel ID="Panel2" runat="server">
                    <asp:HiddenField ID="HiddenField_fileName" runat="server" />
                    <asp:Label ID="Label5" runat="server" Text="Sheet선택: "></asp:Label>
                 <asp:DropDownList ID="ddlSheets" runat="server">
                </asp:DropDownList>
                <asp:Button ID="btnSave" runat="server" Text="Save" onclick="btnSave_Click" onclientclick="return confirm('당월데이타 존재시 삭제후 저장됩니다. 진행하시겠습니까?');" />
                <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="btnCancel_Click" />
                </asp:Panel>
               
            </td>            
            </tr>
            <tr><td colspan="3" style="text-align:center; color : Green;">
                <asp:GridView ID="GridView1" runat="server">
                    <Columns>
                        <asp:TemplateField HeaderText="NO">
                        <ItemTemplate>
                                    <asp:Label ID="Label5" runat="server" Text="<%# Container.DataItemIndex + 1 %>" CssClass="style1"></asp:Label>
                                </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
                <asp:Label ID="lblMessage" runat="server" Text=""></asp:Label></td></tr>
            </table>
    </div>
    <div>
    
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <script type="text/javascript">
            var ModalProgress = '<%= ModalProgress.ClientID %>';

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
         <asp:UpdatePanel ID="UpdatePanel3" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="btnMoveDataNextMonth">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="btn_old_request">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel>
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="btn_request">
                </asp:AsyncPostBackTrigger>
            </Triggers>
        </asp:UpdatePanel> 
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" ImageUrl="~/img/loading9_mod.gif" CssClass="updateProgress" /><br />
                <asp:Image ImageUrl="~/img/ajax-loader.gif" ID="Image2_1" runat="server" />
            </ProgressTemplate>
        </asp:UpdateProgress>
       <cc1:ModalPopupExtender ID="ModalProgress" runat="server" 
            PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="723px" 
            Height="600px" SizeToReportContent="True">
        </rsweb:ReportViewer>
    </div>
    </form>
</body>
</html>
