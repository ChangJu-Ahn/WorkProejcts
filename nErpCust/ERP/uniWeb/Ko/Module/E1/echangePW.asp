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
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

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
<SCRIPT LANGUAGE="VBScript">

</SCRIPT>

<%


    Dim strPass
    Dim Enc1, Enc2
    Dim UID
    Dim password    ' 현재 

	UID = gUsrId

'-------------------------------------------------------------------------------------
    Set Enc2 = Server.CreateObject("EDCodeCom.EDCodeObj.1")
    Call SubOpenDB(lgObjConn)															'☜: Make a DB Connection
    lgStrSQL = "Select password  from E11002T"
    lgStrSQL = lgStrSQL & " Where UID =  " & FilterVar(UID , "''", "S") & ""

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then

		password = lgObjRs("password")
		strPass = Enc2.Decode(password)

	End IF
    Call SubCloseDB(lgObjConn)															'☜: Close DB Connection

    if err.number <> 0 then
       Call ServerMesgBox("CreateObject Error In EDCodeCom" ,vbInformation,I_MKSCRIPT)
       response.end
    end if
    
	Set Enc2 = Nothing
'-------------------------------------------------------------------------------------
%>

<SCRIPT LANGUAGE=VBSCRIPT>

Sub Window_onLoad()

    eForm.txtPassword.focus()
    Call LockField(Document)	

end sub
Sub Window_unLoad()
end sub

Function cmdExit_Click()

    Self.Returnvalue = "C"
	window.self.close

End Function


function chkThisForm()

dim lgStrSQL

    if "<%=Password%>" <> "" then
        if eForm.txtPassword.value = "" then
            Call DisplayMsgBox("210109","X","X","X")
            eForm.txtPassword.focus()
        elseif  eForm.txtPassword.value <> "<%=strPass%>" then
            Call DisplayMsgBox("210106","X","X","X")
            eForm.txtPassword.focus()
        elseif eForm.txtPassword2.value = "" then
            Call DisplayMsgBox("210110","X","X","X")
           eForm.txtPassword2.focus()
        elseif eForm.txtPassword3.value = "" then
            Call DisplayMsgBox("210111","X","X","X")
            eForm.txtPassword3.focus()
        elseif eForm.txtPassword2.value <> eForm.txtPassword3.value then
            eForm.txtPassword3.value = ""
            Call DisplayMsgBox("210111","X","X","X")
            eForm.txtPassword3.focus()
        else
            eForm.submit()
        end if
    else
        eForm.submit()
    end if

End Function

</SCRIPT>
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 

</HEAD>
<BODY>
<center>
<form method="post" name="eForm" action="echangePW_ok.asp" target=formmenu>
	<TABLE cellSpacing=1 cellPadding=1 width="385" border=0  >
		<tr height=5><td></TD></tr>  	
		<tr>
		<td  valign="center" align="center">	   
	<TABLE cellSpacing=1 cellPadding=1 width="380" border=2 class="TABLETEST_BGCOLOR" valign="center" align="center">
		<tr>
		  <td class=TDFAMILY_TITLE colspan=2>패스워드 변경</TD>
	   </tr>    
		<tr>
		  <td class=TDFAMILY_TITLE >ID/성명</td>
		  <TD class=TDFAMILY><input maxLength="13" name="txtUID" value="<%=gUsrId%>" size="13" type="text" tag="24">&nbsp;/&nbsp;<%=gUsrNm%></TD>
		</tr>
		<tr>
		  <td class=TDFAMILY_TITLE >현재 비밀번호</td>
		  <td class=TDFAMILY><input maxLength="10" id=txtPassword value='' name="txtPassword" size="10" type="password" tag="22"></td>
		</tr>
		<tr>
		  <td class=TDFAMILY_TITLE >변경 비밀번호</td>
		  <td class=TDFAMILY><input maxLength="10" name="txtPassword2" size="10" type="password"  tag="22"></td>
		</tr>
		<tr>
		  <td class=TDFAMILY_TITLE >비밀번호 확인</td>
		  <td class=TDFAMILY><input maxLength="10" name="txtPassword3" size="10" type="password"  tag="22"></TD>
	   </tr>    
		<tr>
		  <td class=TDFAMILY_TITLE colspan=2><INPUT id=button1 onclick='VBScript:Call chkThisForm()' type=button value='수정' name=button1><INPUT id=button2 onclick='vbscript:call cmdExit_Click()' type=button value='취소' name=button2></TD>
	   </tr>    
	</TABLE>
	</td>
	   </tr> 	
	</TABLE>	
</form>
<IFRAME NAME="formmenu"  BORDER=0 WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no framespacing =0></IFRAME>
</body>
</html>
