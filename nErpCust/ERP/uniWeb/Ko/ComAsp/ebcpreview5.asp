
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--
'**********************************************************************************************
'*  1. Module Name          : common
'*  2. Function Name        : Easybase web preview 
'*  3. Program ID           : 
'*  4. Program Name         :
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/01/3
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Suk-min
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE>PreView</TITLE>

<%
	Dim EBWidth, EBHeight, dialogWidth, dialogHeight, workFlag
	
	EBWidth  = Request("EBWidth")
	EBHeight = Request("EBHeight")
	
	dialogWidth  = Request("dialogWidth")
	dialogHeight = Request("dialogHeight")
	workFlag	 = Request("workFlag")
	
	Response.Buffer = True
	Response.Write "<span id=loading style=""position:absolute; left:" & (dialogWidth / 2 - 110) & "px;"
	Response.write "top:" & (dialogHeight / 2 - 40) & "px; width:146px; height:67px;"">"
	Response.Write "<table width=220 height=30 border=2 cellpadding=1 cellspacing=1 bordercolor=#CCCCCC bordercolorlight=#CCCCCC bgcolor=buttonface bordercolordark=#000000 vspace=0 hspace=0>"
	Response.Write "<tr bgcolor=#CED3E7>"
	Response.Write "<td colspan=2 bgcolor=#FFFFFF> <font size=2><img src=""../image/net.gif"" width=32 height=31 vspace=0 hspace=0 align=absmiddle></font>"
    Response.Write "<font face=""돋움"" size=2><b>&nbsp;&nbsp;작업을 처리중입니다...</b></font></td></tr></table></span>"
	Response.Flush
%>

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/operation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Dim arrParnet, arrParam
Dim uname
Dim pw
Dim filename
Dim condvar
Dim strUrl
Dim ebform
Dim IFlag
Dim work
	
    arrParent = window.dialogArguments    
    arrParam   = arrParent(0)
    
    uname     = arrparam(0)
    pw        = arrparam(1)
    filename  = arrparam(2)
    condvar   = arrparam(3)
    strUrl    = arrparam(4)
    work      = arrparam(5)
    ebForm    = ""    '2004-05-11 APPLET 문자열 지움 이진수 
    IFrag     = 0
	


Sub checkLoading()
	if IFrag = 0 then
		IFrag = 1
	else
		if UCASE(work) ="PRINT" then
		   self.close
		end if

		document.all("loading").style.visibility="hidden"
		IFrag = 0

	end if
End Sub

Sub getEBR()

	EBAction.id.value       = uname     
	EBAction.pw.value       = pw
	EBAction.doc.value      = filename
	EBAction.runvar.value   = condvar

    If UCASE(work) ="PRINT" Then
	   EBAction.form.value   = "PRESENTER"
    Else
    <%   
         If UCASE(workFlag) = "VIEW" Then
            Response.Write "EBAction.w.value = """ & EBWidth  & """" & vbCrLf
            Response.Write "EBAction.h.value = """ & EBHeight & """" & vbCrLf
         End If    
    %>
	   EBAction.form.value   = "ACTIVEX"
    End If

	EBAction.action         = strUrl
	EBAction.submit
	
	
	
End Sub
</SCRIPT>
</head>

<body onload="getEBR()">
<FORM NAME="EBAction" TARGET="MyBizASP" METHOD="POST"> 
<TABLE width=100%  height=100% cellspacing=0 cellpadding=0 >
 	<tr height=100% >
		<TD><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=auto framespacing=0 marginwidth=0 marginheight=0 onload="checkLoading()"></IFRAME></td>
	</tr>
	<tr height=* >
		<td >
			<input type="hidden" name="id" > 
			<input type="hidden" name="pw" >  
			<input type="hidden" name="doc" > 
			<%If UCASE(workFlag) = "VIEW" Then%>
			<input type="hidden" name="w"> 
			<input type="hidden" name="h"> 
			<%End If%>
			<input type="hidden" name="form" VALUE=ACTIVEX> 
			<input type="hidden" name="runvar" > 
		</td>
	</tr>
</table>
</FORM>
</BODY>
</HTML>
