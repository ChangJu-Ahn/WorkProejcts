<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<HTML>
<HEAD>

<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->

<TITLE><%=gLogoName%>-패스워드 변경</TITLE>

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

</SCRIPT>

<%
    Dim strPass
    Dim UID
    Dim password    ' 현재 

	UID = gUsrId

    Call SubOpenDB(lgObjConn)															'☜: Make a DB Connection
    lgStrSQL = "Select password  from E11002T"
    lgStrSQL = lgStrSQL & " Where UID =  " & FilterVar(UID , "''", "S") & ""

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then

		password = lgObjRs("password")
		strPass = password
		
	End IF

    Call SubCloseDB(lgObjConn)															'☜: Close DB Connection
   
%>

<SCRIPT LANGUAGE=VBSCRIPT>

Sub Window_onLoad()

    eForm.txtPassword.focus()
    Call LockField(Document)	

end sub
Sub Window_unLoad()
end sub

Sub SaveOK()
End Sub

Function cmdExit_Click()

    Self.Returnvalue = "C"
	window.self.close

End Function


function chkThisForm()

   Dim lgStrSQL

   if eForm.txtPassword.value = "" then
      Call DisplayMsgBox("210109","X","X","X")
      eForm.txtPassword.focus()
   elseif  ConnectorControl.xCVTG(eForm.txtPassword.value) <> "<%=strPass%>" then
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
      eForm.txtPassword3.value = ConnectorControl.xCVTG(eForm.txtPassword3.value)
      eForm.action = "echangePW_ok.asp"
      eForm.submit()
   end if
   
   

End Function

</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>
<body leftmargin="0" topmargin="0" marginwidth="0">
<form method="post" name="eForm" target=formmenu>
  <TABLE cellSpacing=0 cellPadding=0 width="400" border=0>
	<tr> 
	  <td width="10" height="5"></td>
	  <td></td>
	  <td width="10"></td>
	</tr>
	<tr> 
	  <td></td>
	  <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
	      <tr> 
	        <td width="30" height="30" align="center" bgcolor="#FFFFFF"><img src="../../CShared/ESSimage/title_icon.gif"></td>
	        <td bgcolor="#FFFFFF" class="contitle">패스워드 변경</td>
	      </tr>
	    </table></td>
	  <td></td>
	</tr>
	<tr> 
	  <td></td>
	  <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD">
			<tr> 
			<td class=ctrow01 >ID/성명</td>
			<TD class=ctrow02><input class="form02" maxLength="13" name="txtUID" value="<%=gUsrId%>" size="13" type="text" tag="24">&nbsp;/&nbsp;<%=gUsrNm%></TD>
			</tr>
			<tr>
			  <td class=ctrow01 >현재 비밀번호</td>
			  <td class=ctrow02><input class="form01" maxLength="10" id=txtPassword value='' name="txtPassword" size="10" type="password"></td>
			</tr>
			<tr>
			  <td class=ctrow01 >변경 비밀번호</td>
			  <td class=ctrow02><input class="form01" maxLength="10" name="txtPassword2" size="10" type="password" ></td>
			</tr>
			<tr>
			  <td class=ctrow01 >비밀번호 확인</td>
			  <td class=ctrow02><input class="form01" maxLength="10" name="txtPassword3" size="10" type="password" ></TD>
		   </tr>    
		  </TABLE>
	  </td>
	</tr> 	
	<tr> 
	  <td height="10"></td>
	  <td></td>
	  <td></td>
	</tr>
	<tr>
	  <td height="35" background="../../CShared/ESSimage/popup_bg_01.gif"></td>
	  <td align="center" valign="bottom" background="../../CShared/ESSimage/popup_bg_01.gif">
		<INPUT type=image id=button1 SRC="../ESSimage/button_06.gif" onclick='VBScript:Call chkThisForm()' value='수정' name=button1 onMouseOver="javascript:this.src='../ESSimage/button_r_06.gif';" onMouseOut="javascript:this.src='../ESSimage/button_06.gif';">
		<INPUT type=image id=button2 SRC="../ESSimage/button_03.gif" onclick='vbscript:call cmdExit_Click()' value='취소' name=button2 onMouseOver="javascript:this.src='../ESSimage/button_r_03.gif';" onMouseOut="javascript:this.src='../ESSimage/button_03.gif';"></TD>
	  <td background="../../CShared/ESSimage/popup_bg_01.gif"></td>
	</tr>
  </TABLE>	
</form>
<IFRAME NAME="formmenu"  BORDER=0 WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no framespacing =0></IFRAME>
</body>
</html>
