<html>
<head><title>인쇄 범위</title>

<!-- #Include file="../inc/IncServer.asp" -->
<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<%'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================%>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Dim objParent

Set objParent = window.dialogArguments

Sub rdoClick()
    with frm1
        .rdodirect1.checked = false
        .rdodirect2.checked = true
        if Trim(.txtFrom.value) = "" and Trim(.txtTo.value) = "" then
            .txtFrom.value = "1"
            .txtTo.value = objParent.All.nTotalPageCount.value
            if .txtTo.value = "" then
               .txtTo.value = "1"
            end if
        end if
    end with
End Sub

Sub rdoClick1()

    with frm1
        .rdodirect2.checked = false
        .rdodirect1.checked = true
        .txtFrom.value = ""
        .txtTo.value = ""
    end with
End Sub

Function CheckNumeric(ByVal strNum) 
  Dim Ret
  Dim intlen, intCnt, intAsc

  intlen = len(strNum)

  for intCnt = 1 to intlen

      intAsc = asc(mid(strNum, intCnt, 1))

	  if intAsc < 48 or intAsc > 57  then
         CheckNumeric = 1
		 Exit function
	  end if
  next

End Function


Function btnCancel_Click()
	Dim arrRet

	ReDim arrRet(3)
	
	arrRet(0) = "CLOSE"
	self.Returnvalue = arrRet
	self.close
End Function

Function btnPrint_Click()
'On error resume next
	Dim arrRet
	Dim IntRetCD
	Dim TotalPageCnt
	Dim iFromValue
	Dim iToValue
	ReDim arrRet(3)
	
	if Trim(objParent.All.nTotalPageCount.value) = "" then
		TotalPageCnt = 1
	Else
		TotalPageCnt = CInt(objParent.All.nTotalPageCount.value) 
	End if
        
        
	If frm1.all.checked = True Then
		arrRet(0) = "ALL"
		arrRet(1) = ""
		arrRet(2) = ""
	Else
		'김승진 2002.11.19 인쇄범위 지정이 될때만 아래로직을 수행 
		iFromValue = Trim(frm1.txtFrom.value)
		iToValue = Trim(frm1.txtTo.value)
		
		if len(iFromValue) <= 0 then
		  iFromValue = 1
		end if
		
		if len(iToValue) <= 0 then
		  iToValue = iFromValue
		end if
		
		If CheckNumeric(iFromValue) = 1 then 		  	
	            IntRetCD = DisplayMsgBox("900034","X","X","X") 	
	            frm1.txtFrom.focus
	            Exit Function
	     End If
	        
		If CInt(iFromValue) < 1 then		  	
	            IntRetCD = DisplayMsgBox("900034","X","X","X") 	
	            frm1.txtFrom.focus
	        Exit Function
	     End If
	
	    If CheckNumeric(iToValue)=1 then		  	
		IntRetCD = DisplayMsgBox("900034","X","X","X") 	
	        frm1.txtTo.focus
	        Exit Function
	    End If
	        
		If CInt(iFromValue) > CInt(iToValue) Then
			IntRetCD = DisplayMsgBox("900036","X","X","X") 	
                        frm1.txtFrom.focus
			Exit Function
		End If

		If TotalPageCnt < CInt(iToValue) Then
			IntRetCD = DisplayMsgBox("900037","X","X","X") 	
                        frm1.txtTo.focus
			Exit Function
		End If
                
		arrRet(0) = "RANGE"
		arrRet(1) = iFromValue
		arrRet(2) = iToValue
	End If

	self.Returnvalue = arrRet
	self.close
End Function

Sub form_load()
End sub
</script>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</head>
<body>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<table width="98%" height="90%" cellspacing="0" cellpadding="0" border="0" align="center" valign="middle">
	<tr>
		<td width="100%" align="center">
			<fieldset>
				<table width="100%" cellspacing="0" cellpadding="1">
					<tr>
						<td class="td5" align="middle">인쇄 범위</td>
					</tr>
				</table>
			</fieldset>
			<fieldset>
				<table width="100%" cellspacing="0" cellpadding="1">
					<tr>
						<td class="td6">
							<input type="radio" id="all" name="rdodirect1" class="radio" checked onclick="rdoclick1"><label for="all">모두</label>&nbsp;
						</td>
					</tr>
					<tr>    <div>
						<td class="td6">
							<input type="radio" id="range" name="rdodirect2" class="radio" onclick="rdoclick"><label for="range">페이지 지정</label> &nbsp;
							<input type="text" name="txtFrom" size="3" MAXLENGTH=4 onclick="rdoclick"> ~ <input type="text" name="txtTo" size="3" MAXLENGTH=4 onclick="rdoclick">
						</td></div>
					</tr>
				</table>
			</fieldset>
		</td>
	</tr>
	<tr>
		<td width="100%" align="center">
			<table cellspacing="10">
			 <tr>
				<td><input type="Button" value="인쇄" name="btnPrint" onclick="btnPrint_Click" class="NormalButton"></td>
				<td><input type="Button" value="취소" name="btnCancel" onclick="btnCancel_Click" class="NormalButton"></td>
			 </tr>
			</table>
		</td>
	</tr>
</table>
</FORM>
</body>
</html>
