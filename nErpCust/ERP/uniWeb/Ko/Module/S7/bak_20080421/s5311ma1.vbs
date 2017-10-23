'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_ID = "s5311mb1.asp"					
Const BIZ_PGM_BILLQRY_ID = "s5311mb2.asp"
Const BIZ_PGM_JUMP_ID = "s5312ma1"

'========================================


Dim gSelframeFlg				
Dim gblnWinEvent				
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
	
'세금계산서 관리방법 
Dim gtxtCreatedMeth
Dim gtxtHistoryflag
Dim	gBlnTaxbillnoMgmtMeth

' 부가세 유형 
Dim arrCollectVatType

'========================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE							
	lgBlnFlgChgValue = False							
	lgBlnFlawChgFlg = False
	lgIntGrpCount = 0									
		
	gblnWinEvent = False
				
End Function

'========================================
Sub SetDefaultVal()
	frm1.txtIssueDt.text = EndDate
	frm1.txtLocCurrency.value = parent.gCurrency 
	frm1.btnPosting.disabled = True
	lgBlnFlgChgValue = False
	'세금계산서 관리방법 설정 
	if gBlnTaxbillnoMgmtMeth then
		frm1.txtMinor_cd.value = gtxtCreatedMeth
		frm1.txtReference.value = gtxtHistoryflag
	end if
	
	Select Case gtxtCreatedMeth
	Case "A"
		Call ggoOper.SetReqAttr(window.document.frm1.txtTaxbillDocNo, "Q")
		window.document.frm1.btnTaxBillDocNo.disabled = True
	Case "M"
		Call ggoOper.SetReqAttr(window.document.frm1.txtTaxbillDocNo, "N")	
		window.document.frm1.btnTaxBillDocNo.disabled = True
	Case "P"
		Call ggoOper.SetReqAttr(window.document.frm1.txtTaxbillDocNo, "Q")
		window.document.frm1.btnTaxBillDocNo.disabled = False 	
	Case "X"
		Call ggoOper.SetReqAttr(window.document.frm1.txtTaxbillDocNo, "D")	
		window.document.frm1.btnTaxBillDocNo.disabled = False 					
	End Select
		
	frm1.txtTaxbillNo.focus
	Set gActiveElement = document.activeElement 
End Sub

'=========================================
Sub SetTaxBillNoMgmtMeth()
	Dim iCodeArr, iTypeArr
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Err.Clear
	 
	gBlnTaxbillnoMgmtMeth = FALSE	 
	
	Call CommonQueryRs(" MINOR_CD, REFERENCE ", " B_CONFIGURATION ", _
	                   " MAJOR_CD = " & FilterVar("S5001", "''", "S") & " AND SEQ_NO = " & FilterVar("1", "''", "S") & "  AND REFERENCE IS NOT NULL ",_
	                     lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	iCodeArr = Split(lgF0, Chr(11))
	iTypeArr = Split(lgF1, Chr(11))

	IF UBound(iCodeArr) = 1 THEN
		gtxtCreatedMeth = iCodeArr(0) '세금계산서 관리방법 
		gtxtHistoryflag = iTypeArr(0) 'Reference
		
		gBlnTaxbillnoMgmtMeth = TRUE
	ELSE '세금계산서 번호 관리방법이 없거나 두개이상 
		gBlnTaxbillnoMgmtMeth = FALSE
	END IF
End Sub	


'==========================================
Function OpenTaxbillNoPop()
	Dim iCalledAspName
	Dim strRet
	
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s5311pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5311pa1", "x")
		gblnWinEvent = False
		exit Function
	end if
	
	strRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	frm1.txtTaxbillNo.focus
			
	If strRet = "" Then
		If Err.Number <> 0 Then
		 Err.Clear 
		End If
		Exit Function
	Else
		Call SetTaxbillNoPop(strRet)
	End If	
End Function

'==========================================
Function OpenBillRef()
	Dim iCalledAspName
	Dim strRet
	
	On Error Resume Next

	if Not gBlnTaxbillnoMgmtMeth then
		Call DisplayMsgBox("205626", "x", "x", "x")
		'세금계산서 방법이 설정되지 않았거나 2개이상 설정되었습니다.
		Exit function
	End If 

	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 
				
	If gblnWinEvent = True Then Exit Function
		
	gblnWinEvent = True
				
	'20021228 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s5111ra3")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5111ra3", "x")
		gblnWinEvent = False
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If strRet(0) = "" Then
		If Err.Number <> 0 Then
		 Err.Clear 
		End If
		Exit Function
	Else
		Call SetBillRef(strRet)
	End If	
End Function	

'==========================================
Function OpenTaxBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If frm1.txtTaxBizAreaCd.readOnly = True Then Exit Function

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
	OpenTaxBizArea = FALSE

	arrParam(0) = "세금신고사업장"					
	arrParam(1) = "B_TAX_BIZ_AREA"						
	arrParam(2) = Trim(frm1.txtTaxBizAreaCd.value)		    
	arrParam(3) = ""				                    
	arrParam(4) = ""									
	arrParam(5) = "세금신고사업장"							

	arrField(0) = "TAX_BIZ_AREA_CD"								
	arrField(1) = "TAX_BIZ_AREA_NM"								

	arrHeader(0) = "세금신고사업장"							
	arrHeader(1) = "세금신고사업장명"							

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	frm1.txtTaxBizAreaCd.focus
	
	If arrRet(0) <> "" Then OpenTaxBizArea = SetTaxBizArea(arrRet)
End Function

'==========================================
Function OpenVatType()
	If frm1.txtVatType.readOnly = True Then
		gblnWinEvent = False
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"	' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtVatType.value)				' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
					& " And Config.MINOR_CD = Minor.MINOR_CD" _
					& " And Config.SEQ_NO = 1"		' Where Condition
	arrParam(5) = "VAT유형"						' TextBox 명칭 
		
    arrField(0) = "Minor.MINOR_CD"					' Field명(0)
    arrField(1) = "Minor.MINOR_NM"					' Field명(1)
    arrField(2) = "Config.REFERENCE"				' Field명(2)
	    	    
    arrHeader(0) = "VAT유형"						' Header명(0)
    arrHeader(1) = "VAT유형명"					' Header명(1)
	arrHeader(2) = "VAT율"					' Header명(2)

	arrParam(0) = arrParam(5)							' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False

	frm1.txtVatType.focus
	If arrRet(0) <> "" Then	Call SetVatType(arrRet)
End Function

'==========================================
Function OpenTaxNo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "세금계산서번호"											' 팝업 명칭 
	arrParam(1) = "S_TAX_DOC_NO"												
	arrParam(2) = Trim(frm1.txtTaxBillNo.value)									
	arrParam(3) = ""															
	
	if frm1.txtIssueDt.text = "" then
		arrParam(4) = " NOT EXISTS(SELECT TAX_DOC_NO FROM S_TAX_BILL_HDR WHERE TAX_DOC_NO = S_TAX_DOC_NO.TAX_DOC_NO) " _
		              & " AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  and USED_FLAG=" & FilterVar("C", "''", "S") & "  and convert(char(10),expiry_date,112) >= convert(char(10),getdate(),112)"	
	else
		arrParam(4) = " NOT EXISTS(SELECT TAX_DOC_NO FROM S_TAX_BILL_HDR WHERE TAX_DOC_NO = S_TAX_DOC_NO.TAX_DOC_NO)  " _
		              & " AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  and USED_FLAG=" & FilterVar("C", "''", "S") & "  and convert(char(10),expiry_date,112) >= "&UniConvDateToYYYYMMDD(frm1.txtIssueDt.text,parent.gDateFormat,"")					
	end if

	arrParam(5) = "세금계산서번호"											
					
	arrField(0) = "ED25" & parent.gColSep & "TAX_DOC_NO"					
	arrField(1) = "DD15" & parent.gColSep & "case when CONVERT(char(11),expiry_date,112)=" & FilterVar("29991231", "''", "S") & " then '' else CONVERT(char(11),expiry_date) end"
	arrField(2) = "ED15" & parent.gColSep & "TAX_BOOK_NO"
	arrField(3) = "ED5" & parent.gColSep & "TAX_BOOK_SEQ"
					    
	arrHeader(0) = "세금계산서번호"			
	arrHeader(1) = "유효일"					
	arrHeader(2) = "책번호(권)"				' Header명(2)%>
	arrHeader(3) = "책번호(호)"				' Header명(3)%>

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	frm1.txtTaxbillDocNo.focus
	
	If arrRet(0) <> "" Then	frm1.txtTaxBillDocNo.value = arrRet(0)
	
End Function

'====================================================
Sub GetTaxBizArea(Byval strFlag)
	On Error Resume Next

	Dim strSelectList, strFromList, strWhereList
	Dim strBilltoParty, strSalesGrp, strTaxBizArea
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp
	
	'세금신고 사업장 Edting시 유효값 Check 및 사업장 명 Fetch
	If strFlag = "NM" Then
		strTaxBizArea = frm1.txtTaxBizAreaCd.value
	Else
		strBilltoParty = frm1.txtBillToPartyCd.value
		strSalesGrp = frm1.txtSalesGrpCd.value
		'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 
		If Len(strBillToParty) > 0 And Len(strSalesGrp) > 0	Then strFlag = "*"
	End if
	
	strSelectList = " * "
	strFromList = " dbo.ufn_s_GetTaxBizArea ( " & FilterVar(strBilltoParty, "''", "S") & ",  " & FilterVar(strSalesGrp, "''", "S") & ",  " & FilterVar(strTaxBizArea, "''", "S") & ",  " & FilterVar(strFlag, "''", "S") & ") "
	strWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		arrTaxBizArea(0) = arrTemp(1)
		arrTaxBizArea(1) = arrTemp(2)
		Call SetTaxBizArea(arrTaxBizArea)
	Else
		If Err.number <> 0 Then	Err.Clear 

		' 세금 신고 사업장을 Editing한 경우 
		If strFlag = "NM" Then
			If Not OpenTaxBizArea() Then
				frm1.txtTaxBizAreaCd.value = ""
				frm1.txtTaxBizAreaNm.value = ""
			End if
		End if
	End if
End Sub

'==========================================
Function SetTaxbillNoPop(arrRet)
	frm1.txtTaxbillNo.Value = arrRet
End Function

'==========================================
Function SetBillRef(strRet)
	Call ggoOper.ClearField(Document, "A")
	Call InitVariables
	Call SetRadio
	Call SetDefaultVal

	frm1.txtHBillNo.value = strRet(0)
	frm1.txtHQueryMode.value = strRet(1)
	Dim strVal
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	strVal = BIZ_PGM_BILLQRY_ID & "?txtBillNo=" & Trim(frm1.txtHBillNo.value)
	strVal = strVal & "&txtQueryMode=" & Trim(frm1.txtHQueryMode.value)		    

	Call RunMyBizASP(MyBizASP, strVal)					
	
	frm1.txtTaxbillNo1.focus
	lgBlnFlgChgValue = True
End Function

'==========================================
Function SetTaxBizArea(arrRet)
	frm1.txtTaxBizAreaCd.Value = arrRet(0)
	frm1.txtTaxBizAreaNm.Value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'==========================================
Function SetVatType(arrRet)
	frm1.txtVatType.value = arrRet(0)
	frm1.txtVatTypeNm.value = arrRet(1)
	frm1.txtVatRate.Text = UNIConvNumPCToCompanyByCurrency(arrRet(2), parent.gCurrency, parent.ggExchRateNo, "X" , "X")
		
	lgBlnFlgChgValue = True
End Function

'========================================
Function FncPosting()
	If Trim(frm1.txtTaxbillNo1.value) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")	 '⊙: "Will you destory previous data" %>
		Exit Function
	End If

	Dim strVal

				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If


    strVal = BIZ_PGM_ID & "?txtMode=" & "POST"								
	strVal = strVal & "&txtTaxBillNo=" & Trim(frm1.txtTaxBillNo1.value)		

	Call RunMyBizASP(MyBizASP, strVal)
End Function	

'========================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877
	Dim strTemp, arrVal

	Select Case Kubun
		
	Case 1
		WriteCookie CookieSplit , frm1.txtHTaxBillNo.value 
	Case 0
		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function				
		frm1.txtTaxBillNo.value =  strTemp
			
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If
		
		Call MainQuery()
		WriteCookie CookieSplit , ""
	Case Else
		Exit Function
	End Select
	 		
End Function
	
'==========================================
Function JumpChgCheck()

	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(1)
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function

'========================================
Function SetRadio()
	Dim blnOldFlag

	blnOldFlag = lgBlnFlgChgValue

	frm1.rdoTaxBillType2.checked = True
	frm1.rdoVATCalcType2.checked = True

	lgBlnFlgChgValue = blnOldFlag
End Function

'============================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim Answer
	
	If lgBlnFlgChgValue = True Then Answer = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")
	If Answer = VBNO Then Exit Function

	If lgBlnFlgChgValue = False Then Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
	If Answer = VBNO Then Exit Function

	BtnSpreadCheck = True

End Function

'====================================================
Sub CurFormatNumericOCX()

	With frm1

		'공급가액 
		ggoOper.FormatFieldByObjectOfCur .txtBillAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'VAT금액 
		ggoOper.FormatFieldByObjectOfCur .txtVATAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	
	End With

End Sub

'==========================================
Function ProtectBody()

    On Error Resume Next
    
	Dim elmCnt, strTagName

	frm1.btnTaxBillDocNo.disabled = True

	For elmCnt = 1 to frm1.length - 1
		If Left(frm1.elements(elmCnt).getAttribute("tag"),1) = "2" Then
			Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "Q")
		End If

		If Err.number <> 0 Then	Err.Clear
	Next

End Function

'==========================================
Function ReleaseBody()

    On Error Resume Next
    
	Dim elmCnt, strTagName

	For elmCnt = 1 to frm1.length - 1
		Select Case Left(frm1.elements(elmCnt).getAttribute("tag"),2)
		Case "21","25"
			Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "D")
		Case "22","23"
			Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "N")
		End Select

		If Err.number <> 0 Then	Err.Clear
	Next

End Function
	
'========================================
Sub Form_Load()
		
	Call LoadInfTB19029														
	Call AppendNumberPlace("6", "4", "0")
		
	Call FormatField()
	Call LockFieldInit("L")
	
	Call SetTaxBillNoMgmtMeth
	Call SetDefaultVal()
				
	Call InitVariables		
	Call ChkTaxbillMgmtMeth()		
	Call CookiePage (0)
							
End Sub

'=========================================
Sub FormatField()
    With frm1
        ' 날짜 OCX Foramt 설정 
		Call FormatDATEField(.txtIssueDt)			
		' 숫자 OCX Foramt 설정 
		Call FormatDoubleSingleField(.txtBillAmt)
		Call FormatDoubleSingleField(.txtVATAmt)
		Call FormatDoubleSingleField(.txtBillLocAmt)
		Call FormatDoubleSingleField(.txtVATLocAmt)
		Call FormatDoubleSingleField(.txtVATRate)
    End With
End Sub

'===========================================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' 날짜 OCX
        Call LockObjectField(.txtIssueDt,"R")    
		Call LockObjectField(.txtBillAmt,"P")    
		Call LockObjectField(.txtVATAmt,"P")
		Call LockObjectField(.txtBillLocAmt,"P")
		Call LockObjectField(.txtVATLocAmt,"P")
		Call LockObjectField(.txtVATRate,"P")   

    End With
End Sub

	
'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
	
'========================================
Sub btnPosting_OnClick()
	If frm1.btnPosting.disabled <> True Then
		If BtnSpreadCheck = False Then Exit Sub
		Call FncPosting()
	End If
End Sub

'========================================
Sub txtIssueDt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================
Sub txtIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtIssueDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoTaxBillType1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoTaxBillType2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoVATCalcType1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoVATCalcType2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoVATIncFlag1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoVATIncFlag2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoVATCalcType1_OnClick()
	lgBlnFlgChgValue = True

	Call ggoOper.SetReqAttr(frm1.rdoVatIncFlag1, "Q")
	Call ggoOper.SetReqAttr(frm1.rdoVatIncFlag2, "Q")
	frm1.txtVatCalcType.value = "1"
End Sub

'========================================
Sub rdoVATCalcType2_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtVatCalcType.value = "2"
'		Call ggoOper.SetReqAttr(frm1.txtVatType, "N")
	Call ggoOper.SetReqAttr(frm1.rdoVatIncFlag1, "N")
	Call ggoOper.SetReqAttr(frm1.rdoVatIncFlag2, "N")
End Sub

'========================================
Sub rdoVatIncFlag1_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtVatIncFlag.value = "1"
End Sub

'========================================
Sub rdoVatIncFlag2_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtVatIncFlag.value = "2"
End Sub

'========================================
Sub txtVatType_OnChange()
	Dim VatType, VatTypeNm, VatRate

	VatType = Trim(frm1.txtVatType.value)
		
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

	frm1.txtVatTypeNm.value = VatTypeNm
	frm1.txtVatRate.text = UNIConvNumPCToCompanyByCurrency(VatRate, parent.gCurrency, parent.ggExchRateNo, "X" , "X")
End Sub

'==========================================
function txtTaxBizAreaCd_OnChange()
	If Trim(frm1.txtTaxBizAreaCd.value) = "" Then
		frm1.txtTaxBizAreaNm.value = ""
	Else
		Call GetTaxBizArea("NM")
		txtTaxBizAreaCd_OnChange=false
			If frm1.txtTaxBizAreaCd.value <> "" Then
				if frm1.rdoVATCalcType1.checked Then
					If frm1.rdoVATCalcType1.disabled = false Then 'kek추가 
						frm1.rdoVATCalcType1.focus 
					End if
				Else
					If frm1.rdoVATCalcType2.disabled = false Then 'kek추가				
						frm1.rdoVATCalcType2.focus 
					End if
				End if
			End if
	
	End if
End function

'========================================
Function ChkTaxbillMgmtMeth()
	If gBlnTaxbillnoMgmtMeth Then
		Call SetToolbar("1110000000001111")
		ChkTaxbillMgmtMeth = True
	ELSE
	    Call DisplayMsgBox("205626", "x", "x", "x")
	    '세금계산서 방법이 설정되지 않았거나 2개이상 설정되었습니다.
		Call SetToolbar("1100000000001111")					
		ChkTaxbillMgmtMeth = False
	End if
End Function

'========================================
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))
	
	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub

'========================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)  
		If arrCollectVatType(iCnt, 0) = UCase(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
End Sub

'========================================
Function FncQuery()
	Dim IntRetCD
		
	FncQuery = False											

	Err.Clear													
		
	If Not chkFieldByCell(frm1.txtTaxbillNo,"A",gPageNo) Then Exit Function

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	Call InitVariables
	
	Call DbQuery()											
		
	FncQuery = True											
End Function
	
'========================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False                                                     

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")					
	Call ggoOper.LockField(Document, "N")					
	Call InitVariables										
	Call SetRadio
	Call ReleaseBody()
	Call SetDefaultVal
	Call SetToolbar("1110000000001111")
		
	lgBlnFlgChgValue = False
	FncNew = True		
		
End Function
	
'========================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False									
		
	If lgIntFlgMode <> parent.OPMD_UMODE Then						
		Call DisplayMsgBox("900002","x","x","x")
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	Call DbDelete						

	FncDelete = True					
End Function

'========================================
Function FncSave()
	Dim IntRetCD
		
	FncSave = False												
		
	Err.Clear													
		
	If lgBlnFlgChgValue = False Then							
	    IntRetCD = DisplayMsgBox("900001","x","x","x")			
	    Exit Function
	End If	
	
	If Not chkFieldByCell(frm1.txtIssueDt, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtTaxBizAreaCd, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtVatType, "A", "1") Then Exit Function   
    
    If Not ChkFieldLengthByCell(frm1.txtRemark, "A", "1") Then Exit Function        
    		
	Call DbSave											
		
	FncSave = True										
End Function

'========================================
Function FncPrint()
	Call parent.FncPrint()														'☜: Protect system from crashing%>
End Function

'========================================
Function FncPrev() 
    Dim strVal
	    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002","x","x","x")  '☜ 바뀐부분 
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    End If

				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If


	frm1.txtPrevNext.value = "PREV"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							
    strVal = strVal & "&txtWarrentNo=" & Trim(frm1.txtWarrentNo.value)				
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		
	         
	Call RunMyBizASP(MyBizASP, strVal)
End Function

'========================================
Function FncNext() 
    Dim strVal
	    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002","x","x","x")  '☜ 바뀐부분 
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    End If

				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If


	frm1.txtPrevNext.value = "NEXT"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							
    strVal = strVal & "&txtWarrentNo=" & Trim(frm1.txtWarrentNo.value)				
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		
	         
	Call RunMyBizASP(MyBizASP, strVal)
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLE)
End Function

'========================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, True)
End Function

'========================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================
Function DbQuery()
	Err.Clear

	DbQuery = False														

	Dim strVal

				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							
	strVal = strVal & "&txtTaxbillNo=" & Trim(frm1.txtTaxbillNo.value)		
	strVal = strVal & "&empty=empty"
		
	Call RunMyBizASP(MyBizASP, strVal)
		
	DbQuery = True															
End Function

'========================================
Function DbSave()
	Err.Clear
		
	DbSave = False														
		
	If frm1.chkBillNoFlg.checked = True Then
		frm1.txtBillNoFlg.value = "Y"
	Else
		frm1.txtBillNoFlg.value = "N"
	End If				
		
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If


	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = parent.gUsrID			
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	DbSave = True														
End Function
	
'========================================
Function DbDelete()
	Err.Clear
		
	DbDelete = False														
		
	Dim strVal
		
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							
	strVal = strVal & "&txtTaxbillNo=" & Trim(frm1.txtTaxbillNo1.value)		
		
	Call RunMyBizASP(MyBizASP, strVal)
		
	DbDelete = True															
End Function

'========================================
Function DbQueryOk()	
	lgIntFlgMode = parent.OPMD_UMODE
	lgBlnFlgChgValue = False
	frm1.txtLocCurrency.value = parent.gCurrency

	'Call ggoOper.LockField(Document, "Q")

	If frm1.rdoPostFlg1.checked = True Then
		Call ProtectBody()
	ElseIf frm1.rdoPostFlg2.checked = True Then
		Call ReleaseBody()
			
		Call ggoOper.SetReqAttr(frm1.txtTaxbillNo1, "Q")
		Call ggoOper.SetReqAttr(frm1.chkBillNoFlg, "Q")
		Call ggoOper.SetReqAttr(frm1.txtTaxbillDocNo, "Q")
		
		window.document.frm1.btnTaxBillDocNo.disabled = True

		If frm1.rdoVATCalcType1.checked Then
			ggoOper.SetReqAttr frm1.rdoVATIncFlag1, "Q"
			ggoOper.SetReqAttr frm1.rdoVATIncFlag2, "Q"
			frm1.txtVatCalcType.value = "1"
		else
			frm1.txtVatCalcType.value = "2"
		end if

		If frm1.rdoVATIncflag1.checked Then
			frm1.txtVatIncFlag.value = "1"
		else
			frm1.txtVatIncFlag.value = "2"
		end if

		'부가세포함여부(내역이 등록된 경우 Protect)
		If UNICDbl(frm1.txtBillAmt.Text) > 0 Then
			ggoOper.SetReqAttr frm1.rdoVATCalcType1, "Q"
			ggoOper.SetReqAttr frm1.rdoVATCalcType2, "Q"
		 	ggoOper.SetReqAttr frm1.txtVATType, "Q"
		 	ggoOper.SetReqAttr frm1.rdoVATIncFlag1, "Q"
		 	ggoOper.SetReqAttr frm1.rdoVATIncFlag2, "Q"
		 	ggoOper.SetReqAttr frm1.txtIssueDt, "Q"
		 End if
	End If	  

	' 세금계산서번호 관리방법 Check		
	if gBlnTaxbillnoMgmtMeth then
		Call SetToolbar("11111000000211111")
	else
		Call SetToolbar("1101100000001111")
	end if

	lgBlnFlgChgValue = False
		
	frm1.txtTaxbillNo.focus	
End Function

'========================================
Function BillQueryOk()
	frm1.chkBillNoFlg.checked = False 
		
	If frm1.rdoVATCalcType1.checked Then
		ggoOper.SetReqAttr frm1.rdoVATIncFlag1, "Q"
		ggoOper.SetReqAttr frm1.rdoVATIncFlag2, "Q"
		frm1.txtVatCalcType.value = "1"
	else
		frm1.txtVatCalcType.value = "2"
	end if
		
	If frm1.rdoVATIncFlag1.checked Then
		frm1.txtVatIncflag.value = "1"
	else
		frm1.txtVatIncFlag.value = "2"
	end if

	Call SetToolbar("1110100000001111")										
End Function
	
'========================================
Function DbSaveOk()														
	Call InitVariables
	Call MainQuery()
End Function
	
'========================================
Function DbDeleteOk()													
	Call MainNew()
	' 세금계산서 관리방법 Check
	if Not gBlnTaxbillnoMgmtMeth then
		Call SetToolbar("1100000000001111")
	end if
End Function

'========================================
Function PostOk()													
	lgBlnFlgChgValue = False
	Call MainQuery()
End Function
