<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="MultiCheckCombo.ascx.cs" Inherits="SRL.UserControls.MultiCheckCombo" Debug="true" ViewStateMode="Inherit" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<script type = "text/javascript"> 
//Script para incluir en el ComboBox1 cada item chekeado del CheckBoxList1Materiales
    function CheckItem(cblCtrl, tbCtrl, hihCtrl, hihCtrl2) {
        var vItems = cblCtrl.getElementsByTagName("input");
        //var arrayOfCheckBoxLabels = cblCtrl.getElementsByTagName("label");
        var i, iCnt = 0, strChecked = "", vDelim = /\, /g; //, TxtBox = "";

        if (vItems[0].checked) {    // Check All Items
            vItems[0].checked = false;
            for (i = 2; i < vItems.length; i++)
                vItems[i].checked = true;
        }
        else if (vItems[1].checked) {   // Uncheck All Items
            for (i = 1; i < vItems.length; i++)
                vItems[i].checked = false;
        }

        for (i = 2; i < vItems.length; i++) {
            if (vItems[i].checked) {
                strChecked = strChecked + ", " + vItems[i].value; // arrayOfCheckBoxValues[i].innerText;
                iCnt++;
            }
        }

        if (strChecked.length > 0) strChecked = strChecked.substring(2, strChecked.length); //sacar la primer 'coma'

        document.getElementById(tbCtrl).value = strChecked;
        document.getElementById(hihCtrl).value = (strChecked.length > 0) ? "'" + strChecked.replace(vDelim, "', '") + "'" : "";
        document.getElementById(hihCtrl2).value = iCnt;
        document.getElementById(tbCtrl).focus();
    }

//function SetFocus(tbCtrl) {
//    var tbTemp = tbCtrl;
//    var input = tbTemp.get_inputDomElement();
//    input.focus();
//}
</script>

<asp:TextBox ID="TextBox1" runat="server" ReadOnly="true" Width="350" ></asp:TextBox>
<ajaxToolkit:PopupControlExtender ID="PopupControlExtender1" runat="server" TargetControlID="TextBox1" PopupControlID="Panel1" Position="Bottom" ></ajaxToolkit:PopupControlExtender>
<input type="hidden" name="hihSelectedText" ID="hihSelectedText" runat="server"/>
<input type="hidden" name="hihSelectedCount" ID="hihSelectedCount" runat="server" value="0"/>
<asp:Panel ID="Panel1" runat="server" ScrollBars="Vertical" Width="350" Height="400" BackColor="AliceBlue" BorderColor="Gray" BorderWidth="1">    
    <asp:CheckBoxList ID="CheckBoxList1" runat="server" OnClick="CheckItem(this)" RepeatColumns="2" OnInit="CheckBoxList1_Init" RepeatDirection="Horizontal"></asp:CheckBoxList>    
</asp:Panel>
