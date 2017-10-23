<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sa005.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sa001.sm_sa005.sm_sa005" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .TD5
        {
            PADDING-RIGHT: 5px;
            WIDTH: 14%;
            BACKGROUND-COLOR: #e7e5ce;
            TEXT-ALIGN: right
        }
        .TD6
        {
            PADDING-LEFT: 5px;
            WIDTH: 36%;
            BACKGROUND-COLOR: #f5f5f5
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
     </style>
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
                   <asp:Label ID="Label1" runat="server" Text="Au 회수 중량 등록" CssClass="title" Width="100%"></asp:Label>
                  </td>
            </tr>
        </table>
        <table  style="width: 100%;">
            <tr >
                <td style="text-align: right;">
                    &nbsp;</td>
            </tr>
        </table>
        <table WIDTH ="100%"; border="0" cellSpacing="0" rules="none">
            <tr>
                <td class="TD5">
                    Scrap 종류
                </td>
                <td class="TD6">

                    <asp:DropDownList ID="ddl_ScrapList" runat="server" Height="22px" Width="215px" AutoPostBack="True" OnSelectedIndexChanged="ddl_ScrapList_SelectedIndexChanged">
                    </asp:DropDownList>

                </td>
                <td class="TD5">
                    LOT NO
                </td>
                <td class="TD6">
                     <asp:DropDownList ID="ddl_R_Doc_No" runat="server" Height="19px" Width="215px" AutoPostBack="True" OnSelectedIndexChanged="ddl_R_Doc_No_SelectedIndexChanged">
                    </asp:DropDownList>
                    
                </td>
            </tr>
            <tr>
                <td class="TD5">
                    회수 공장
                </td>
                <td class="TD6">

                    <asp:DropDownList ID="ddl_PlantCD" runat="server" Height="22px" Width="215px">
                    </asp:DropDownList>

                </td>
                <td class="TD5">
                    측정 저울</td>
                <td class="TD6">
                    
                    <asp:DropDownList ID="ddl_WeightEqp" runat="server" Height="22px" Width="215px">
                    </asp:DropDownList>

                </td>
            </tr>
            <tr id="ETCH" runat="server" style="display:none">
                <td class="TD5">
                    중량
                </td>
                <td class="TD6">
                    <asp:TextBox ID="txtEtchWeight" runat="server" Width="120px" ReadOnly="True"></asp:TextBox>
                    <asp:TextBox ID="txtEtchWUnit" runat="server" Width="24px" ReadOnly="True"></asp:TextBox>
                    <asp:Button ID="btnGetEtchWeight" runat="server" Text="가져오기" OnClick="btnGetEtchWeight_Click" />
                    <asp:TextBox ID="txtEtchWeightKey" runat="server" Width="100px" ReadOnly="True" Visible="False"></asp:TextBox>
                    <asp:Button ID="btnEtchSave" runat="server" Text="저장" OnClick="btnEtchSave_Click" />
                </td>
                <td class="TD5">
                </td>
                <td class="TD6">
                     
                </td>
            </tr>
            <tr id="PLAT" runat="server" style="display:none">
                <td class="TD5">
                    작업전 중량
                </td>
                <td class="TD6">
                        

                    <asp:TextBox ID="txtPlatWeight1" runat="server" Width="120px" ReadOnly="True"></asp:TextBox>
                    <asp:TextBox ID="txtPlatWUnit1" runat="server" Width="24px" ReadOnly="True"></asp:TextBox>
                    <asp:Button ID="btnGetPlatWeight1" runat="server" Text="가져오기" OnClick="btnGetPlatWeight1_Click" />
                    <asp:TextBox ID="txtPlatWgtKey1" runat="server" Width="100px" ReadOnly="True" Visible="False"></asp:TextBox>    

                    <asp:Button ID="btnPlatSave1" runat="server" Text="저장" OnClick="btnPlatSave1_Click" />

                </td>
                <td class="TD5">
                    작업후 중량
                </td>
                <td class="TD6">
                     <asp:TextBox ID="txtPlatWeight2" runat="server" Width="120px" ReadOnly="True"></asp:TextBox>
                     <asp:TextBox ID="txtPlatWUnit2" runat="server" Width="24px" ReadOnly="True"></asp:TextBox>   
                    <asp:Button ID="btnGetPlatWeight2" runat="server" Text="가져오기" OnClick="btnGetPlatWeight2_Click" />
                    <asp:TextBox ID="txtPlatWgtKey2" runat="server" Width="100px" ReadOnly="True" Visible="False"></asp:TextBox>
                    <asp:Button ID="btnPlatSave2" runat="server" Text="저장" OnClick="btnPlatSave2_Click" />
                </td>
            </tr>

        </table>
    </div>
    </form>
</body>
</html>
