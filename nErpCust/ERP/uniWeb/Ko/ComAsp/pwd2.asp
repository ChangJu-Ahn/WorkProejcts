<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<!-- #Include file="../inc/incSvrMain.asp"  -->
<%

    Dim iStrFlag
    Dim RegConfig
    Dim lResult
    Dim StrEisUrl
    Dim strApp
    Dim iTemp
    
    
    Call LoadBasisGlobalInf()
    
    iStrFlag = Request.QueryString("txtFlag")
    iTemp = Trim(Request.QueryString("initp"))
    
    If iTemp = "" Then
       iTemp= "unilogin.asp"        
    End If
    
    strApp = Trim(Request.QueryString("strApp"))
    
    If strApp <> "" Then
       StrEisUrl = GetGlobalInf2(NodeNm3,"gEWareURL")
    End If   
	
%>

<HTML>
<HEAD>
<TITLE><%=Request.Cookies("unierp")("gLogoName")%></TITLE>


<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">
<SCRIPT LANGUAGE="vbscript"   SRC="../inc/common.vbs"></SCRIPT>

<Script Language="VBScript">

Option Explicit

Sub Window_OnLoad()
<%If iStrFlag = "L" Then%>
     vtxtOld.focus
<%Else%>
     vtxtNew.focus
<%End If%>


End Sub

'========================================================================================
Function UNIMsgBox(pVal, pType, pTitle)
	MsgBox pVal, pType, pTitle
End Function

Sub HandleError(ByVal pData)
    If InStr(pData,"210112") > 0 Then
       vtxtOld.select
       vtxtOld.focus
    End If        
End Sub	

Sub doSubmit()
    If window.event.keyCode = 13 Then
       Call CheckVal()
    End If
End Sub

Sub CheckVal()
    Dim IntRetCD
    Dim strVal


<% If iStrFlag = "L" Then %>
	If Trim(vtxtOld.value) = "" Then
       Call  MsgBox (vtxtOld.alt & "은(는) 입력 필수 항목입니다." ,vbExclamation,"<%=Request.Cookies("unierp")("gLogoName")%>")
       vtxtOld.focus
       Exit Sub	
	End If
	frm1.txtOld.value = LCase(MXD(vtxtOld.value))
<% End If %>
   		
	If Len(vtxtNew.value) < 6 Then
        Call  MsgBox ("신 비밀번호는 6자리 이상이어야 합니다." ,vbExclamation,"<%=Request.Cookies("unierp")("gLogoName")%>")
		vtxtNew.select
		Exit Sub
	End If	    
		
	If vtxtNew.value <> vtxtRe.value Then
        Call  MsgBox ("신 비밀번호와 재확인 비밀번호가 맞지 않습니다." ,vbExclamation,"<%=Request.Cookies("unierp")("gLogoName")%>")
		vtxtNew.select
		Exit Sub
	End If
	
	frm1.txtNew.value = LCase(MXD(vtxtNew.value))
	frm1.txtRe.value  = LCase(MXD(vtxtRe.value))
	
    frm1.txtFlag.value = "<%=iStrFlag%>"
    frm1.action = "./pwdbiz.asp"
    frm1.submit
	
End Sub

Sub SaveOK()
    
	If "<%=strApp%>" ="EIS" then 
       Top.location.href = "<%=StrEisUrl%>" & "?UID=" & "<%=gUsrId%>" & "&gCon=" & "<%=gADODBConnString%>"
	Else
       Call  MsgBox ("다시 로그인을 하십시요." ,vbExclamation,"<%=Request.Cookies("unierp")("gLogoName")%>")
       If "<%=Request("SAPP")%>" = "EIS" Then
          Top.location.href = "../uniEISLogin.asp"
       Else   
          Top.location.href = "../<%=iTemp%>"
       End If   
	End If
End Sub

Sub onInit()

    Dim objConn
    
    On Error Resume Next

    Set objConn = CreateObject("uniConnector.cGlobal")
    
    If Err.number = 0 Then
       objConn.CheckURL("<%=Request.Cookies("unierp")("gURLLangUserID")%>")
       Call objConn.ExitProcess("A")
       Set objConn = Nothing
    End If
    
    top.location.href = "../<%=iTemp%>"
End Sub

</Script>

<body  BGCOLOR=#ffffff scroll=no>
<table width=100% height=100% border=0>
<tr align=center>
<td>&nbsp;</td>
<td>

<table cellpadding=0 cellspacing=0 width=500 border=0>

	<tr><td width = 100% height = 8  colspan=4></td></tr>
	
	<tr bgcolor = #B8BFAD><td width=100% height=1  colspan=4></td></tr>
	<tr><td width = 100% height = 1  colspan=4></td></tr>
	
	<tr bgcolor = #e6e6e6>
		<td height=20></td>
		<td colspan=3 class=tdclass05 style="background-color:#E6E6E6;font-weight:bold;"><center>비밀번호 변경</center></td>
	</tr>
	
	<tr><td width = 100% height = 1  colspan=4></td></tr>

	<tr bgcolor = #B8BFAD><td width=100% height=1  colspan=4></td></tr>



<%
If iStrFlag = "L" Then                 'The valid period of password has been expired.  Please register a new password. 
%>

	<tr bgcolor = #F3F5E7><td width=100% height=10  colspan=4></td></tr>
	<tr bgcolor= #F3F5E7>
		<td width = 5% height=25></td>
		<td width = 45% class=tdclass05 style="background-color:#F3F5E7;font-weight:bold;color:#666668;"><img src= '../image/login/logButton.gif'>&nbsp;구 비밀번호</td>
		<td width = 5% style="background-color:#F3F5E7;font-weight:bold;color:#666668;">:</td>
		<td width = 65% class=tdclass05 style="background-color:#F3F5E7;font-weight:bold;color:#666668;"><INPUT TYPE=PASSWORD id=vtxtOld name=vtxtOld ALT="구 비밀번호" style="BACKGROUND-COLOR: #ffffb4"></td>
	</tr>

<%
Else
%>


	<tr><td width = 100% height = 1  colspan=4></td></tr>
	<tr bgcolor = #e6e6e6>
		<td height=20></td>
		<td colspan=3 class=tdclass05 style="background-color:#E6E6E6;font-weight:bold;"><center>해당 ID로 최초 사용입니다.신규 비밀번호를 입력하십시오.</center></td>
	</tr>
	<tr><td width = 100% height = 1  colspan=4></td></tr>
	<tr bgcolor = #B8BFAD><td width=100% height=1  colspan=4></td></tr>
	<tr bgcolor = #F3F5E7><td width=100% height=10  colspan=4></td></tr>


<%
End If
%>

	<tr bgcolor= #F3F5E7>
		<td width = 5% height=25></td>
		<td width = 45% class=tdclass05 style="background-color:#F3F5E7;font-weight:bold;color:#666668;"><img src= '../image/login/logButton.gif'>&nbsp;신 비밀번호</td>
		<td width = 5% style="background-color:#F3F5E7;font-weight:bold;color:#666668;">:</td>
		<td width = 65% class=tdclass05 style="background-color:#F3F5E7;font-weight:bold;color:#666668;"><INPUT TYPE=PASSWORD id=vtxtNew name=vtxtNew onkeydown="doSubmit()" ALT="신 비밀번호" style="BACKGROUND-COLOR: #ffffb4" ></td>
	</tr>
	<tr bgcolor= #F3F5E7>
		<td width = 5% height=25></td>
		<td width = 45% class=tdclass05 style="background-color:#F3F5E7;font-weight:bold;color:#666668;"><img src= '../image/login/logButton.gif'>&nbsp;재확인 비밀번호</td>
		<td width = 5% style="background-color:#F3F5E7;font-weight:bold;color:#666668;">:</td>
		<td width = 65% class=tdclass05 style="background-color:#F3F5E7;font-weight:bold;color:#666668;"><INPUT TYPE=PASSWORD id=vtxtRe name=vtxtRe onkeydown="doSubmit()" ALT="재확인 비밀번호" style="BACKGROUND-COLOR: #ffffb4"></td>
	</tr>

    <tr bgcolor = #F3F5E7><td width=100% height=10  colspan=4></td></tr>

	<tr bgcolor= #F3F5E7>
		<td colspan=21 class=tdclass05 style="background-color:#F3F5E7;font-weight:bold;color:#666668;"><center>6문자 이상 비밀번호를 입력 하십시요.</center></td>
	</tr>

	<tr bgcolor = #F3F5E7><td width=100% height=10  colspan=4></td></tr>

	<tr bgcolor = #dd9944><td width=100% height=1  colspan=4></td></tr>
	<tr bgcolor = #F3F5E7><td width=100% height=3  colspan=4></td></tr>
	<tr bgcolor= #F3F5E7>
		<td align=center colspan=5>
		    <table>
		    <tr>
		    <td>
            <table class=btnclass02 cellpadding=0 cellspacing=0>
                <tr><td class=btntd02l><img src='../image/login/buttonleft.gif'></td>
	                <td class=btntd02><div align=center><a onclick="CheckVal()" class=btn02>반영</a></div></td>
                    <td class=btntd02r><img src='../image/login/buttonright.gif'></td>
                </tr>
            </table>
		    </td>
		    <td>
            <table class=btnclass02 cellpadding=0 cellspacing=0>
                <tr><td class=btntd02l><img src='../image/login/buttonleft.gif'></td>
	                <td class=btntd02><div align=center><a onclick="onInit()" class=btn02>로그인 화면으로</a></div></td>
                    <td class=btntd02r><img src='../image/login/buttonright.gif'></td>
                </tr>
            </table>
		    </td>
		    </tr>
		    </table>
            
		</td>
	<tr bgcolor = #F3F5E7><td width=100% height=2  colspan=4></td></tr>
	<tr bgcolor = #dd9944><td width=100% height=1  colspan=4></td></tr>
		
		</td>
	</tr>


</table>

</td>
<td>&nbsp;</td>
</tr>
<TR>
	<TD WIDTH=100% HEIGHT=1 COLSPAN=10><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
</TR>
</table>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<INPUT TYPE=HIDDEN   id= txtNew  name= txtNew >
<INPUT TYPE=HIDDEN   id= txtRe   name= txtRe >
<INPUT TYPE=HIDDEN   id= txtOld  name= txtOld >
<input type=HIDDEN   id= txtFlag name= txtFlag>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>

</BODY>
</HTML>