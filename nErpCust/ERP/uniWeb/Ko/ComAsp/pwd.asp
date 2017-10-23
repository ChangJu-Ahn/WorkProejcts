<%@ LANGUAGE="VBSCRIPT" %>
<%
Dim iStrFlag
iStrFlag  = Request("txtFlag")

%>

<HTML>
<HEAD>

<!-- #Include file="../inc/incSvrMain.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/IncCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/IncCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/IncCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit

Dim arrParent
		
Dim iStrFlag


On Error Resume Next

iStrFlag  = "<%=iStrFlag%>"
		
If iStrFlag = "L" Then
	top.document.title = "비밀번호 수정"
Else
	top.document.title = "비밀번호 입력"
End If
		
Self.Returnvalue = Array("CANCEL")

'=================================================================================================
Function OKClick()
    Dim strVal
    
	If CheckVal() = False Then
	   Exit Function
	End If	

    strVal = "PWDBiz.asp?txtFlag=" & iStrFlag

	If iStrFlag = "L" Then
		strVal = strVal & "&txtOld=" & LCase(MXD(Escape(Trim(txtOld.value))))
	End If
		
    strVal = strVal & "&txtNew=" & LCase(MXD(Escape(Trim(txtNew.value))))
    strVal = strVal & "&txtRe="  & LCase(MXD(Escape(Trim(txtRe.value))))
    strVal = strVal & "&txtUsrID=" & "<%=Request("txtUsr")%>" 
	    
	If "<%=Request("skipSave")%>" = "1" Then
		If CheckVal() = False Then
		   Exit Function
		End If	
		Call SaveOk()
	Else
		Call LayerShowHide(1)
		Call RunMyBizASP(MyBizASP, strVal)			'☜: 비지니스 ASP 를 가동 
	End If
End Function

Sub SaveOk()
	Self.Returnvalue = Array("OK", LCase(MXD(Escape(Trim(txtNew.value)))))
	Self.Close()
End Sub
 
Function CheckVal()
    Dim IntRetCD

	CheckVal = False

<%
If iStrFlag = "L" Then                 'The valid period of password has been expired.  Please register a new password. 
%>
	If Trim(txtOld.value) = "" Then
       Call  MsgBox (txtOld.alt & "은(는) 입력 필수 항목입니다." ,vbExclamation,"<%=Request.Cookies("unierp")("gLogoName")%>")
       txtOld.focus
       Exit Function	
	End If

<%
End If
%>
   		
	If Len(txtNew.value) < 6 Then
        Call  MsgBox ("비밀번호는 6자리 이상이어야 합니다." ,vbExclamation,"<%=Request.Cookies("unierp")("gLogoName")%>")
		txtNew.select
		Exit Function
	End If	    
		
	If txtNew.value <> txtRe.value Then
        Call  MsgBox ("비밀번호가 맞지 않습니다." ,vbExclamation,"<%=Request.Cookies("unierp")("gLogoName")%>")
		txtNew.select
		Exit Function
	End If
	
	CheckVal = True
End Function
	
'=================================================================================================
Function CancelClick()
	Self.Close()
End Function
	
'=================================================================================================
Sub Form_Load()


	If iStrFlag = "L" Then
		txtOld.focus
	Else
		txtNew.focus
	End If
		
End Sub

'=================================================================================================
Sub Window_onLoad()
    Call Form_Load()    
End Sub


'=================================================================================================
Sub RunMyBizASP(objIFrame, strURL)
	Call BtnDisabled(True)
	objIFrame.location.href = GetUserPath & strURL

End Sub

'=================================================================================================
Function GetUserPath()
	If gURLPath = "" or isEmpty(gURLPath) Then
		Dim strLoc, iPos , iLoc, strPath
		strLoc = window.location.href
                iLoc = inStr(1, strLoc, "?")
            
                If iLoc > 0 Then
                   strLoc = Left(strLoc, iLoc - 1)
                End If
		
		iLoc = 1: iPos = 0
		Do Until iLoc <= 0						
			iLoc = inStr(iPos+1, strLoc, "/")
			If iLoc <> 0 Then iPos = iLoc
		Loop	
		gURLPath = Left(strLoc, iPos)
	End If
	GetUserPath = gURLPath
End Function

Function document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function

Sub TextKeypress(pos)
	If window.event.keyCode = 13 Then
		Select Case pos
			Case 3
				Call OKClick()
		End Select
	End If
End sub
	
Sub HandleError(ByVal pData)
    If InStr(pData,"210112") > 0 Then
       txtOld.select
       txtOld.focus
    End If        
End Sub	
</SCRIPT>
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=10 CLASS="basicTB">
	<TR>
		<TD HEIGHT=*>
			<FIELDSET>
			<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
<%
If iStrFlag = "L" Then                 'The valid period of password has been expired.  Please register a new password. 
%>
				<TR>
					<TD CLASS="TD5" STYLE="WIDTH:40%">구 비밀번호</TD>
					<TD CLASS="TD6" STYLE="WIDTH:60%"><INPUT TYPE="PassWord" Name="txtOld" SIZE=15 MAXLENGTH=10 tag="12" ALT="구 비밀번호" onkeypress="TextKeyPress 1" style="BACKGROUND-COLOR: #ffffb4"></TD>
				</TR>		
<%
End If
%>
				<TR>
					<TD CLASS="TD5" STYLE="WIDTH:40%">신 비밀번호</TD>
					<TD CLASS="TD6" STYLE="WIDTH:60%"><INPUT TYPE="PassWord" NAME="txtNew" SIZE=15 MAXLENGTH=10 tag="12" ALT="신 비밀번호" onkeypress="TextKeyPress 2" style="BACKGROUND-COLOR: #ffffb4"></TD>
				</TR>		
				<TR>
					<TD CLASS="TD5" STYLE="WIDTH:40%">비밀번호 확인</TD>
					<TD CLASS="TD6" STYLE="WIDTH:60%"><INPUT TYPE="PassWord" NAME="txtRe" SIZE=15 MAXLENGTH=10 tag="12" ALT="비밀번호 확인" onkeypress="TextKeyPress 3" style="BACKGROUND-COLOR: #ffffb4"></TD>
				</TR>		
			</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=100% ALIGN=RIGHT>
						<IMG SRC="../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=1><IFRAME NAME="MyBizASP" SRC="PWDBiz.asp" WIDTH=100% HEIGHT=1 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm" tabindex=-1></iframe>
</DIV>
</BODY>
</HTML>
