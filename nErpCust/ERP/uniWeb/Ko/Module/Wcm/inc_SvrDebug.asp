<% session.CodePage=949 %>
<%
Dim gblnDebug

gblnDebug = TRUE

Sub PrintLog(Byval pDesc)
	If gblnDebug Then
		'Response.Write pDesc & "<br>" & vbCrLf
	End If
End Sub


' 메세지 출력 
Sub PrintMesg(Byval strMesg)
%>
<meta http-equiv="Content-Type" CONTENT="text/html; charset=<%=LANGUAGE_METATAG%>"><% ' 클라이언트에게 돌려줄 문자셋입니다. %>
<body>
<form name=a><textarea name=txtMesg style="display: none"><%=strMesg%> </textarea></form>
<script language=javascript>
alert(a.txtMesg.value);
</script>
</body>
<%
End Sub

' 메세지 출력 
Sub PrintMesg2(Byval strMesg)
%>
<form name=a><textarea name=txtMesg style="display: none"><%=strMesg%> </textarea></form>
<script language=javascript>
alert(a.txtMesg.value);
</script>
<%
End Sub

' 환경변수 출력 
Public Sub PrintEnviroment()
	Dim item
%>
<table>
<%	
	For Each item In Request.ServerVariables 
%>
<tr><td bgcolor=yellow><%=item%></td><td><%=Request.ServerVariables(item)%></td></tr>
<%	
	Next
%>
</table>
<%	
	Response.End
End Sub

' 폼 출력 
Public Sub PrintForm()
	Dim item
%>
<table>
<%	
	For Each item In Request.Form
%>
<tr><td bgcolor=yellow><%=item%></td><td><%=Request.Form(item)%></td></tr>
<%	
	Next
%>
</table>
<%	
	Response.End
End Sub

' Get양식 출력 
Public Sub PrintGet()
	Dim item
%>
<table>
<%	
	For Each item In Request.QueryString
%>
<tr><td bgcolor=yellow><%=item%></td><td><%=Request.QueryString(item)%></td></tr>
<%	
	Next
%>
</table>
<%	
	Response.End
End Sub

' 레코드셋 출력 
Public Sub PrintRs(Byref oRs)
	Dim i
%>
<table border=1 cellspacing=0 cellpadding=0>
<tr bgcolor=navy>
<% For i = 0 to oRs.Fields.Count-1 %>
<th><font color=white><%=oRs.Fields(i).name%></font></th>
<% Next %>
</tr>
<%	Do until oRs.EOF %>
<tr>
<% For i = 0 to oRs.Fields.Count-1 %>
	<td bgcolor=yellow><%=oRs(i)%></td>
<% Next %>	
</tr>
<%		oRs.MoveNext
	Loop
%>
</table>
<%	oRs.Close 
	Set oRs = Nothing	
	Response.End
End Sub

Sub DebugXML(Byval pstrSQL)
	Dim itop 
	itop = int(rnd()*100)*2
	Response.Write "<table style='position: absolute; float: left; left: 0; top: " & itop & "; z-index: 10' width=100% border=1 bgcolor=lightyellow><tr><td align=left><font color=red>Make SQL:</font></td></tr><tr><td align=left>"
	pstrSQL = Replace(pstrSQL, "<" , "&lt;")
	pstrSQL = Replace(pstrSQL, ">" , "&gt;")
	pstrSQL = Replace(pstrSQL, vbTab , "&nbsp;&nbsp;")
	pstrSQL = Replace(pstrSQL, vbCrLf , "<br>")
	Response.Write pstrSQL
	Response.Write "</td></tr></table>"
	Response.End 
End Sub

Private Function RemoveHostName(pURL)
	Dim iPos
	iPos = Instr(1, UCase(pURL), UCase(Request.ServerVariables("SERVER_NAME")))
	If iPos > 0 Then ' URL에 도메인명이 존재 
		RemoveHostName = Mid(pURL, iPos+Len(Request.ServerVariables("SERVER_NAME")))
	Else
		RemoveHostName = pURL
	End If
End Function

' 메세지 출력후 URL점프 
Sub GoToURLWithMesg(pMesg, pURL) 
%>
<meta http-equiv="Content-Type" CONTENT="text/html; charset=<%=LANGUAGE_METATAG%>"><% ' 클라이언트에게 돌려줄 문자셋입니다. %>
<script language=javascript>
function goToURL() {
alert(a.txtMesg.value);
location.replace("<%=pURL%>");
}
</script>
<body onload="javascsript:goToURL()">
<form name=a><textarea name=txtMesg style="display: none"><%=pMesg%>
</textarea></form>
</body>
<%
	Response.End
End Sub

' 메세지 출력후 URL점프 
Sub GoToBackWithMesg(pMesg) 
%>
<meta http-equiv="Content-Type" CONTENT="text/html; charset=<%=LANGUAGE_METATAG%>"><% ' 클라이언트에게 돌려줄 문자셋입니다. %>
<script language=javascript>
function goToURL() {
alert(a.txtMesg.value);
history.back();
}
</script>
<body onload="javascsript:goToURL()">
<form name=a><textarea name=txtMesg style="display: none"><%=pMesg%></textarea></form>
</body>
<%
	Response.End
End Sub

Sub GoToURLWithProgress(pURL)
	If Response.Buffer Then Response.Clear 
%>
<HTML>
<HEAD>
<TITLE>Wait for moment...... </TITLE>
<meta http-equiv="Content-Type" CONTENT="text/html; charset=<%=LANGUAGE_METATAG%>"><% ' 클라이언트에게 돌려줄 문자셋입니다. %>
<link REL="stylesheet" TYPE="text/css" HREF="/common/css/popup-<%=TYPE_LANGUAGE%>.css">
<link REL="stylesheet" TYPE="text/css" HREF="/common/css/default-<%=TYPE_LANGUAGE%>.css">
<script language=javascript>
	function window_onload() {
		location.replace("<%=pURL%>");
	}
</script>
</HEAD>
<BODY id="Body" bottommargin=0 rightmargin=0 leftmargin=0 topmargin=10 onload="javascript:window_onload()">
<div id="divMesg" name="divMesg">
<table align=center width=330 height=50 border=0 cellspacing=0 cellpadding=0 align=center>
<tr>
	<td><b>Please Waitting...</b></td>
</tr>
<tr>
	<td><img src="/common/images/etc/mini_progress.gif" width=330></td>
</tr>
</table>
</div>
</BODY>
</HTML>
<%
	Response.Flush 
	Response.End 
End Sub

Sub PrintSession
	Dim Item
%>
<table border=1 cellspacing=0 cellpadding=0>
<%	For Each Item In Session.Contents  %>
<tr>
	<td bgcolor=yellow><%=Item%></td><td><%=Session.Contents(Item)%></td>
</tr>
<% Next %>
</table>
<%
End Sub

Sub CloseAfterAlert()
%>
<meta http-equiv="Content-Type" CONTENT="text/html; charset=<%=LANGUAGE_METATAG%>"><% ' 클라이언트에게 돌려줄 문자셋입니다. %>
<script language=javascript>
function goToURL() {
alert(a.txtMesg.value);
self.close()
}
</script>
<body onload="javascsript:goToURL()">
<form name=a><textarea name=txtMesg style="display: none"><%=pMesg%>
</textarea></form>
</body>
<%
End Sub

' 폼 출력 
Public Sub PrintArray(pArr)
	Dim i, iLen
	iLen = Ubound(pArr)
%>
<table>
<%	
	For i= 1 To iLen
%>
<tr><td bgcolor=yellow><%=i%></td><td><%=pArr(i)%></td></tr>
<%	
	Next
%>
</table>
<%	
	Response.End
End Sub


Sub DebugSql(Byval pstrSQL)
	Dim itop 
	
	itop = int(rnd()*100)*2
	Response.Write "<table style='position: absolute; float: left; left: 0; top: " & itop & "; z-index: 3' width=100% border=1 bgcolor=lightyellow><tr><td align=left><font color=red>Make SQL:</font></td></tr><tr><td align=left>"
	
	if pstrSQL <> "" then	
		pstrSQL = Replace(pstrSQL, "<" , "&lt;")
		pstrSQL = Replace(pstrSQL, ">" , "&gt;")
		
'		pstrSQL = Replace(pstrSQL, vbTab , "&nbsp;&nbsp;")

		pstrSQL = "<pre>" & pstrSQL & "</pre>"
	else
		pstrSQL = "DebugSql에 넘어온 값이 없습니다."
	end if
	Response.Write pstrSQL	
	
	Response.Write "</td></tr></table>"
	Response.End 
End Sub


Function GetASPError()
	Dim objASPError
	Dim sMsg, blnErrorWritten
	
	Set objASPError = Server.GetLastError
	
	'If LCase(Response.ContentType) = "text/xml" Then sMsg = "<![CDATA["
'	sMsg = sMsg &  Response.ContentType & vbCrlf
	sMsg = sMsg &  objASPError.Category
	If objASPError.ASPCode > "" Then sMsg = sMsg &  ", " & objASPError.ASPCode
	sMsg = sMsg &  " (0x" & Hex(objASPError.Number) & ")" & vbCrLf

	sMsg = sMsg &  objASPError.Description & vbCrLf

	If objASPError.ASPDescription > "" Then sMsg = sMsg &  objASPError.ASPDescription & vbCrLf

	blnErrorWritten = False

	' Only show the Source if it is available and the request is from the same machine as IIS
	If objASPError.Source > "" Then
	  strServername = LCase(Request.ServerVariables("SERVER_NAME"))
	  strServerIP = Request.ServerVariables("LOCAL_ADDR")
	  strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
	  If (strServername = "localhost" Or strServerIP = strRemoteIP) And objASPError.File <> "?" Then
	    sMsg = sMsg &  objASPError.File 
	    If objASPError.Line > 0 Then sMsg = sMsg &  ", line " & objASPError.Line
	    If objASPError.Column > 0 Then sMsg = sMsg &  ", column " & objASPError.Column
	    sMsg = sMsg &  vbCrLf
	    sMsg = sMsg &  Server.HTMLEncode(objASPError.Source) & vbCrLf
	    If objASPError.Column > 0 Then sMsg = sMsg &  String((objASPError.Column - 1), "-") & "^" & vbCrLf
	    blnErrorWritten = True
	  End If
	End If

	If Not blnErrorWritten And objASPError.File <> "?" Then
	  sMsg = sMsg &  objASPError.File
	  If objASPError.Line > 0 Then sMsg = sMsg &  ", line " & objASPError.Line
	  If objASPError.Column > 0 Then sMsg = sMsg &  ", column " & objASPError.Column
	  sMsg = sMsg &  vbCrLf
	End If
	
	'If LCase(Response.ContentType) = "text/xml" Then sMsg = sMsg & "]]>"
	GetASPError = sMsg
End Function

Sub SaveErrorLog(pErr)
	if pErr.Number <> 0 Then
		GoToBackWithMesg pErr.Description
	Else
		GoToBackWithMesg GetASPError
	End If
End Sub
%>

