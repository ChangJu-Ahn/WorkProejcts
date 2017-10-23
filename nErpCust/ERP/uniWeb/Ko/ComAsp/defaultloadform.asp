<!-- #Include file="../inc/CommResponse.inc" -->
<html>
<head>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/DefaultValue.vbs"></script>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/EventPopup.vbs"></script>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>
<script language="vbscript">
Sub Form_Load()
End Sub
Sub Window_OnLoad()
End Sub
Sub setclipboard(key)
	Dim clipboard
	Set clipboard = CreateObject("uni2kcm.SaveFile")
	clipboard.SetClipBoardData(key)
	Set clipboard = Nothing
	close()
End Sub
top.document.title = "Default Value"
</script>
</head>

<BODY TABINDEX="-1">
<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr>
		<td width="100%" valign="top">
<%
Const KEY_NAME = "DEFAULT_VALUE"
Const KEY_CNT = "KEY_CNT"
Const KEY_ITEM = "TXT"

Dim nCnt 
	
nCnt = 0

Response.Write "<table bgcolor=""#D1E8F9"" border=""1"" cellspacing=""0"" cellpadding=""3"" bordercolorlight=""#E2F9FA"" bordercolordark=""#C0D7E8"" width=""100%"">" & chr(13)
For Each ItemKey in Request.Cookies(KEY_NAME)
	If ItemKey <> KEY_CNT Then
		nCnt = nCnt + 1
		Response.Write "<TR>" & chr(10)
		Response.Write "	<TD>" & chr(10)
		Response.Write " &nbsp;&nbsp;&nbsp; <A HREF=""#"" OnClick='vbscript:setclipboard(""" & replace(Request.Cookies(KEY_NAME)(ItemKey),"""","""""") &""")'>" & Request.Cookies(KEY_NAME)(ItemKey) & "</a>" & chr(10)
		Response.Write "	</TD>" & chr(10)
		Response.Write "</TR>" & chr(10)
	End If 
Next
Response.Write "</table>" & chr(13)
%>
		</TD>
    </TR>
    <TR>
        <TD>
            <TABLE width="100%">
                <TR>
                    <TD align="left">&nbsp;&nbsp;
						<IMG SRC="../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="close()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Cancel.gif',1)"></IMG>
                    </TD>
                </TR>
            </TABLE>
		</td>
	</tr>
</table>
<%
If nCnt = 0 Then %>
<script language="vbscript">
	Msgbox "저장된 값이 없습니다.", vbInformation
	close()
</script>
<%
End If 
%>
</body>
</html>



