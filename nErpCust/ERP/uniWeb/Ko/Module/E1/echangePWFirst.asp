<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<TITLE><%=gLogoName%>-패스워드 변경</TITLE>
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>

<%
Dim strPass
Dim Enc1, Enc2
Dim UID
Dim password    ' 현재 

	UID = gUsrId
    Password =  gPws


    Set Enc2 = Server.CreateObject("EDCodeCom.EDCodeObj.1")

    if err.number <> 0 then
       Call ServerMesgBox("CreateObject Error In EDCodeCom" ,vbInformation,I_MKSCRIPT)
       response.end
    end if
    
    strPass = Enc2.Decode(password)
	Set Enc2 = Nothing

%>

<script language=VBscript>
Dim CFlag : CFlag = True

Function Document_onClick()
Dim Evobj

Set Evobj = window.event.srcElement

    If IsNull(Evobj.id) Then
        CFlag = True
        Exit Function
    Else
        If UCase(Evobj.id) = "BUTTON1" Then
            CFlag = False
        Else
            CFlag = True
        End If
    End IF
    Set Evobj = nothing
Document_onClick = True
End Function
Sub Window_OnLoad()

    eForm.txtPassword2.focus()
    Call LockField(Document)	

end sub
Sub Window_onUnLoad()
    If CFlag Then
        call cmdExit_Click()
    End If
End Sub
Function cmdExit_Click()

    Self.Returnvalue = "C"
	window.self.close

End Function

function chkThisForm()
        if eForm.txtPassword2.value = "" then
            Call DisplayMsgBox("210110","X","X","X")
            eForm.txtPassword2.focus()
        elseif eForm.txtPassword3.value = "" then
            Call DisplayMsgBox("210111","X","X","X")
            eForm.txtPassword3.focus()
        elseif eForm.txtPassword2.value <> eForm.txtPassword3.value then
            Call DisplayMsgBox("210112","X","X","X")
            eForm.txtPassword3.focus()
        else
            eForm.submit()
        end if
End Function
</script>

</head>
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 

<body>
<form method="post" name="eForm" action="echangePW_ok.asp" target=formmenu>
	<TABLE cellSpacing=1 cellPadding=1 width="380" border=0 class="TABLETEST_BGCOLOR">
		<tr>
		  <td class=TDFAMILY_TITLE colspan=2>패스워드 변경</TD>
	   </tr>    
		<tr>
		  <td class=TDFAMILY_TITLE>ID</td>
		  <TD class=TDFAMILY><input maxLength="13" name="txtUID" value="<%=UID%>" size="13" type="text" tag="24"></TD>
		</tr>
		<tr>
		  <td class=TDFAMILY_TITLE>변경 비밀번호</td>
		  <td class=TDFAMILY><input maxLength="10" name="txtPassword2" size="10" type="password" tag="22"></td>
		</tr>
		<tr>
		  <td class=TDFAMILY_TITLE>비밀번호 확인</td>
		  <td class=TDFAMILY><input maxLength="10" name="txtPassword3" size="10" type="password" tag="22"></TD>
	   </tr>    
		<tr>
		  <td class=TDFAMILY_TITLE colspan=2><INPUT id=button1 onclick='VBScript:Call chkThisForm()' type=button value='수정' name=button1><INPUT id=button2 onclick='vbscript:call cmdExit_Click()' type=button value='취소' name=button2></TD>
	   </tr>    
	</TABLE>
</form>
<IFRAME NAME="formmenu"  BORDER=0 WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no framespacing =0></IFRAME>
</body>
</html>
