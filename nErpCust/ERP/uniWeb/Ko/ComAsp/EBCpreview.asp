
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
	Dim EBWidth, EBHeight
	
	EBWidth  = Request("EBWidth")
	EBHeight = Request("EBHeight")
	
	Response.Buffer = True
	Response.Write "<span id=loading style=""position:absolute; left:" & (EBWidth / 2 - 110) & "px;"
	Response.write "top:" & (EBHeight / 2 - 40) & "px; width:146px; height:67px;"">"
	Response.Write "<table width=220 height=30 border=2 cellpadding=1 cellspacing=1 bordercolor=#CCCCCC bordercolorlight=#CCCCCC bgcolor=buttonface bordercolordark=#000000 vspace=0 hspace=0>"
	Response.Write "<tr bgcolor=#CED3E7>"
	Response.Write "<td colspan=2 bgcolor=#FFFFFF> <font size=2><img src=""../image/net.gif"" width=32 height=31 vspace=0 hspace=0 align=absmiddle></font>"
    Response.Write "<font face=""Arial"" size=2><b>&nbsp;&nbsp;작업을 처리중입니다...</b></font></td></tr></table></span>"
	Response.Flush
%>

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/operation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Dim arrParnet, arrParam
Dim uname
Dim dbname
Dim filename
Dim condvar
Dim ebdate
Dim strUrl
Dim Fixed
Dim IFlag
	
    arrParent = window.dialogArguments    
    arrParam   = arrParent(0)
    
    uname     = arrparam(0)
	dbname    = arrparam(1)
	filename  = arrparam(2)
	condvar   = arrparam(3)
	ebdate    = arrparam(4)	
	strUrl    = arrparam(5)
	Fixed     = arrParam(6)
	IFrag     = 0
	


sub checkLoading()
	if IFrag = 0 then
		IFrag = 1
	else
		document.all("loading").style.visibility="hidden"
		IFrag = 0
	end if
end sub

sub getEBR()
	
	
	EBAction.uname.value    = uname
	EBAction.dbname.value   = dbname
	EBAction.filename.value = filename
	EBAction.condvar.value  = condvar
	EBAction.date.value     = ebdate
	EBAction.fixed.value    = fixed
	
	EBAction.action         = strUrl
	EBAction.submit

end sub
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
			<input type="hidden" name="uname" > 
			<input type="hidden" name="dbname" >  
			<input type="hidden" name="filename" > 
			<input type="hidden" name="condvar" > 
			<input type="hidden" name="date" >    
			<input type="hidden" name="fixed" >    
		</td>
	</tr>
</table>
</FORM>
</BODY>
</HTML>
