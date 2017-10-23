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
top.document.title = "Default Value"

Dim element_offset
element_offset= 1


' 이 화면은 DD 대상 화면이 압니다.
'========================================================================================
' Function Name : Document_onKeyDown
' Function Desc : hand all event of key down
'========================================================================================
Function Document_onKeyDown()
	Dim KeyCode
	
	On Error Resume Next
	
	KeyCode = window.event.keycode
	Select Case KeyCode	
		Case 13
               Call CheckCheckBox()
               Exit Function
		Case 27   'ESC
		       Self.Close
               Exit Function
     End Select		
     
End Function

Sub CheckCheckBox()
	Dim s, ret

	s = get_sequence_checkbox()

	If s = "" Then
		Msgbox "선택된 값이 없습니다.", vbInformation
		Exit Sub
	End If

	If Msgbox("선택한 항목을 저장 할까요?", vbQuestion + vbYESNO) = vbNo Then
		Exit Sub
	End If

	Dim strVal
	Dim nCnt

	nCnt = document.frm1.cnt.value

	strVal = "DefaultSave.asp" & "?cnt=" & nCnt
	For i = 0 To nCnt - 1
		If nCnt = 1 Then
			If document.frm1.cb.Checked = True Then
				strVal = strVal & "&cb" & i & "=" & document.frm1.cb.value
			Else
				strVal = strVal & "&cb" & i & "="
			End If
		Else
			If document.frm1.cb(i).Checked = True Then
				strVal = strVal & "&cb" & i & "=" & document.frm1.cb(i).value
			Else
				strVal = strVal & "&cb" & i & "="
			End If
		End If
	Next

	Call RunMyBizASP(MyBizASP, strVal)

End Sub

Sub RunMyBizASP(objIFrame, strURL)
	objIFrame.location.href = GetUserPath & strURL
End Sub

</script>
</head>

<BODY TABINDEX="-1">
<table border="0" cellspacing="0" cellpadding="0" width="100%"><tr><td valign="top">
<form name="frm1" target="MyBizASP" method="post" action="DefaultSave.asp">
<input type="hidden" name="cnt" value="<%=Request("cnt")%>">
<TABLE CELLSPACING=0 WIDTH="100%">
    <TR>
        <TD WIDTH="100%" VALIGN="TOP">
<%
	Dim i

	Response.Write "<table bgcolor=""#D1E8F9"" border=""1"" cellspacing=""0"" cellpadding=""1"" bordercolorlight=""#E2F9FA"" bordercolordark=""#C0D7E8"" width=""100%"">" & chr(13)
	For i = 1 to Request("cnt")
		Response.Write " <tr> " & chr(13)
		Response.Write " 	<TD NOWRAP><label for=""C" & i & """><input id=""C" & i & """ CLASS=""CHECK"" type=checkbox name=cb value='" & request.querystring("txt" & i) & "'> &nbsp;&nbsp;" & request.querystring("txt" & i) & "</label></td> " & chr(13)
		Response.Write " </tr>" & chr(13)
	Next 
	Response.Write "</table>" & chr(13)
%>
		</TD>
    </TR>
    <TR>
        <TD VALIGN="TOP">
            <TABLE width="100%">
                <TR>
                    <TD align="left">&nbsp;&nbsp;
						<IMG SRC="../image/ok_d.gif" style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="vbscript:CheckCheckBox()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../image/cancel_d.gif" style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="close()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
                    </TD>
                </TR>
            </TABLE>
        <TD>
    </TR>   
	<TR HEIGHT="0">
		<TD HEIGHT="0"><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT="0" FRAMEBORDER="0" SCROLLING="NO" noresize framespacing="0"></IFRAME></TD>
	</TR>
</TABLE>
</td></tr></table>
</form>
</body>
</html>
