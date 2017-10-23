' External ASP File
'========================================
Const BIZ_PGM_ID = "s5111mb1.asp"
Const BIZ_BillDtl_JUMP_ID = "s5112ma1"
Const BIZ_BillCollect_JUMP_ID = "s5114ma1"

' Constant variables 
'========================================
Const TAB1 = 1                  '☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

Const SoNoHdr = "SoNoHdr"
Const LCNoHdr = "LCNoHdr"
Const PostFlag = "PostFlag"

' Common variables 
'========================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status


Dim arrCollectVatType

'세금계산서 관리방법 
Dim gtxtCreatedMeth
Dim gtxtHistoryflag
Dim gBlnTaxbillnoMgmtMeth

Dim IsOpenPop      ' Popup
Dim gSelframeFlg

'========================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'=========================================
Sub SetDefaultVal()
	With frm1

		.txtConBillNo.focus
		Set gActiveElement = document.activeElement 

		.txtBillCommand.value = "NEW"
		.txtBillDt.Text = EndDate
		.txtSalesGrpCd.value = Parent.gSalesGrp
		.txtToBizAreaCd.value = Parent.gSalesGrp
		.txtRadioFlag.value = "N"
		.txtXchgRate.Text = 1
		.rdoPostFlagN.checked = True
		.btnPostFlag.disabled = True
		.btnGLView.disabled = True
		.btnPreRcptView.disabled = True
		.btnPostFlag.value = "확정"
		.chkTaxNo.checked = False
		.chkSoNo.disabled = False
		.txtLocCur.value = Parent.gCurrency
		.btnBillTaxNo.disabled = True
		lgBlnFlgChgValue = False  
		
	End With
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
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1) '~~~ 첫번째 Tab
	gSelframeFlg = TAB1
End Function

'==========================================
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2) '~~~ 두번째 Tab
	gSelframeFlg = TAB2
End Function

'==========================================
Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3) '~~~ 두번째 Tab
	gSelframeFlg = TAB3
End Function

'==========================================
Function OpenSORef()
	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call DisplayMsgBox("204150", "X", "X", "X")
		Exit Function
	End IF
	
	iCalledAspName = AskPRAspName("s3111ba1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3111ba1", "X")
		Exit Function
	End If
		
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	' 공통 문제로 Popup에서 Error 발생 
	If Err.Number <> 0 Then Err.Clear 

	If arrRet(0) <> "" Then
		frm1.txtRefFlag.value = "S"
		Call SetToolBar("11101000000111")
		Call RefFlagPR()
		Call SetSORef(arrRet)
	End If 
End Function

'==========================================
Function OpenLCRef()
	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call DisplayMsgBox("205150", "X", "X", "X")
		Exit Function
	End IF

	iCalledAspName = AskPRAspName("s3111ba2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3111ba2", "X")
		lblnWinEvent = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	' 공통 문제로 Popup에서 Error 발생 
	If Err.Number <> 0 Then Err.Clear 

	If arrRet(0) <> "" Then
		frm1.txtRefFlag.value = "L"
		Call SetToolBar("11101000000111")
		Call RefFlagPR()
		Call SetLCRef(arrRet)
	End If 
End Function

'==========================================
Function OpenDNRef()
	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call DisplayMsgBox("205153", "X", "X", "X")
		Exit Function
	End IF

	iCalledAspName = AskPRAspName("s3111ba3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3111ba3", "X")
		lblnWinEvent = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	' 공통 문제로 Popup에서 Error 발생 
	If Err.Number <> 0 Then Err.Clear 

	' Popup에서 Cancel한 경우 
	If UBOUND(arrRet) = 0 Then	Exit Function
		 
	frm1.txtRefFlag.value = "D"
	Call SetToolBar("11101000000111")
	Call RefFlagPR()
		 
	' 정상출고인 경우 
	If Trim(arrRet(0)) <> "" Then
		Call SetSORef(arrRet)
	Else
		'예외출고인 경우 
		Call GetDnInfo(arrRet(3))
	End If
End Function

'==========================================
Function OpenConBillNo()
	On Error Resume Next
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent

	iCalledAspName = AskPRAspName("s5111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111pa1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	  
	strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=N", Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	frm1.txtConBillNo.focus
	If strRet <> "" Then  frm1.txtConBillNo.value = strRet  

End Function

'==========================================
Function OpenBillHdr(ByVal iBillHdr)
	 Dim arrRet
	 Dim arrParam(5), arrField(6), arrHeader(6)

	 OpenBillHdr = False
	 If IsOpenPop = True Then Exit Function
	 IsOpenPop = True

	 Select Case iBillHdr
	 CASE 0
		If frm1.txtBillTypeCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "매출채권형태"
		arrParam(1) = "S_BILL_TYPE_CONFIG"
		arrParam(2) = Trim(frm1.txtBillTypeCd.value)
		arrParam(3) = ""
		arrParam(4) = "EXCEPT_FLAG <> " & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXPORT_FLAG <> " & FilterVar("Y", "''", "S") & "  AND AS_FLAG = " & FilterVar("N", "''", "S") & " "
		arrParam(5) = "매출채권형태"
	 
		arrField(0) = "BILL_TYPE"
		arrField(1) = "BILL_TYPE_NM"

		arrHeader(0) = "매출채권형태"
		arrHeader(1) = "매출채권형태명"
		
		frm1.txtBillTypeCd.focus
	      
	 Case 1
		If frm1.txtPayerCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"  ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtPayerCd.value)        ' Code Condition
		arrParam(3) = ""              ' Name Cindition
					arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SPA", "''", "S") & " " _
					& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE LIKE " & FilterVar("C%", "''", "S") & " " _
					& "AND PARTNER_FTN.BP_CD =  " & FilterVar(frm1.txtSoldtoPartyCd.value, "''", "S") & " " ' Where Condition
		arrParam(5) = "수금처"      ' TextBox 명칭 
	  
	    arrField(0) = "PARTNER.BP_CD"     ' Field명(0)
	    arrField(1) = "PARTNER.BP_NM"     ' Field명(1)
		   
		arrHeader(0) = "수금처"      ' Header명(0)
		arrHeader(1) = "수금처명"     ' Header명(1)

		frm1.txtPayerCd.focus
	 Case 2            
		If frm1.txtToBizAreaCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_SALES_GRP"      ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtToBizAreaCd.value) ' Code Condition
		arrParam(3) = ""        ' Name Cindition
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "     ' Where Condition
		arrParam(5) = "수금영업그룹"    ' TextBox 명칭 
		   
		arrField(0) = "SALES_GRP"      ' Field명(0)
		arrField(1) = "SALES_GRP_NM"     ' Field명(1)
		      
		arrHeader(0) = "영업그룹"     ' Header명(0)
		arrHeader(1) = "영업그룹명"     ' Header명(1)

		frm1.txtToBizAreaCd.focus
	 Case 3
		If frm1.txtTaxBizAreaCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_TAX_BIZ_AREA"      ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtTaxBizAreaCd.value) ' Code Condition
		arrParam(3) = ""        ' Name Cindition
		arrParam(4) = ""        ' Where Condition
		arrParam(5) = "세금신고사업장"    ' TextBox 명칭 
		   
		arrField(0) = "TAX_BIZ_AREA_CD"      ' Field명(0)
		arrField(1) = "TAX_BIZ_AREA_NM"      ' Field명(1)
		      
		arrHeader(0) = "세금신고사업장"    ' Header명(0)
		arrHeader(1) = "세금신고사업장명"   ' Header명(1)

		frm1.txtTaxBizAreaCd.focus
	 Case 4
		If frm1.txtPayTypeCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If Trim(frm1.txtPayTermsCd.value) = "" Then
			Call DisplayMsgBox("205151", "X", "X", "X")
			'Msgbox "수주참조 또는 LC참조를 먼저 하세요"
			frm1.txtPayTermsCd.focus
			IsOpenPop = False   
			Exit Function
		End IF

		arrParam(0) = "입금유형"     ' 팝업 명칭 
		arrParam(1) = "B_MINOR,B_CONFIGURATION," _
					& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & ""_
    			    & "And MINOR_CD= " & FilterVar(frm1.txtPayTermsCd.value, "''", "S") & " And SEQ_NO>=2)C" ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtPayTypeCd.value)  ' Code Condition
		arrParam(3) = ""        ' Name Cindition
		arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
					& "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("R", "''", "S") & " )" ' Where Condition
		arrParam(5) = "입금유형"     ' TextBox 명칭 
	  
		arrField(0) = "B_MINOR.MINOR_CD"    ' Field명(0)
		arrField(1) = "B_MINOR.MINOR_NM"    ' Field명(1)
		   
		arrHeader(0) = "입금유형"     ' Header명(0)
		arrHeader(1) = "입금유형명"     ' Header명(1)

		frm1.txtPayTypeCd.focus
	 Case 5            

		If Trim(frm1.txtRefFlag.value) = "" Then
			Call DisplayMsgBox("205151", "X", "X", "X")
			'Msgbox "수주참조 또는 LC참조를 먼저 하세요"
			IsOpenPop = False   
			Exit Function
		End If
	  
		If frm1.txtPayTermsCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
		End If

		Select Case frm1.txtRefFlag.value
		Case "S"
				arrParam(4) = "CON.MINOR_CD = MINOR.MINOR_CD" _
							& " AND CON.MAJOR_CD = MINOR.MAJOR_CD AND CON.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "" _
							& " AND CON.REFERENCE = " & FilterVar("N", "''", "S") & " " _
							& " AND CON.SEQ_NO = 1"      ' Where Condition
		Case "D"
				arrParam(4) = "CON.MINOR_CD = MINOR.MINOR_CD" _
							& " AND CON.MAJOR_CD = MINOR.MAJOR_CD AND CON.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "" _
							& " AND CON.REFERENCE <> " & FilterVar("M", "''", "S") & " " _
							& " AND CON.SEQ_NO = 1"      ' Where Condition
		Case "L"
				arrParam(4) = "CON.MINOR_CD = MINOR.MINOR_CD" _
						    & " AND CON.MAJOR_CD = MINOR.MAJOR_CD AND CON.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "" _
							& " AND CON.REFERENCE = " & FilterVar("L", "''", "S") & " " _
							& " AND CON.SEQ_NO = 1"      ' Where Condition
		Case Else
				arrParam(4) = "CON.MINOR_CD = MINOR.MINOR_CD" _
							& " AND CON.MAJOR_CD = MINOR.MAJOR_CD AND CON.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "" _
							& " AND CON.REFERENCE <> " & FilterVar("M", "''", "S") & " " _
							& " AND CON.SEQ_NO = 1"      ' Where Condition
		End Select

		arrParam(1) = "B_MINOR MINOR, "_
					& "B_CONFIGURATION CON"      ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtPayTermsCd.value) ' Code Condition
		arrParam(3) = ""        ' Name Cindition
		arrParam(5) = "결제방법"     ' TextBox 명칭 
	  
	    arrField(0) = "CON.MINOR_CD"     ' Field명(0)
	    arrField(1) = "MINOR.MINOR_NM"     ' Field명(1)
	     
	    arrHeader(0) = "결제방법"     ' Header명(0)
	    arrHeader(1) = "결제방법명"     ' Header명(1)
	    
	    frm1.txtPayTermsCd.focus

	 Case 6            
		If frm1.txtBeneficiaryCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_BIZ_PARTNER"     ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtBeneficiaryCd.value) ' Code Condition
		arrParam(3) = ""        ' Name Cindition
		arrParam(4) = ""        ' Where Condition
		arrParam(5) = "양수자"      ' TextBox 명칭 
	  
		arrField(0) = "BP_CD"       ' Field명(0)
		arrField(1) = "BP_NM"       ' Field명(1)
		   
		arrHeader(0) = "양수자"      ' Header명(0)
		arrHeader(1) = "양수자명"     ' Header명(1)
		
		frm1.txtBeneficiaryCd.focus

	 Case 7            
		If frm1.txtApplicantCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_BIZ_PARTNER"     ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtApplicantCd.value) ' Code Condition
		arrParam(3) = ""        ' Name Cindition
		arrParam(4) = ""        ' Where Condition
		arrParam(5) = "양도자"      ' TextBox 명칭 
	  
		arrField(0) = "BP_CD"       ' Field명(0)
		arrField(1) = "BP_NM"       ' Field명(1)

		arrHeader(0) = "양도자"      ' Header명(0)
		arrHeader(1) = "양도자명"     ' Header명(1)

		frm1.txtApplicantCd.focus
	 Case 8
		If frm1.txtBillToPartyCd.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"  ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtBillToPartyCd.value)       ' Code Condition
		arrParam(3) = ""        ' Name Cindition
		arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SBI", "''", "S") & " " _
					& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE LIKE " & FilterVar("C%", "''", "S") & " " _
					& "AND PARTNER_FTN.BP_CD =  " & FilterVar(frm1.txtSoldtoPartyCd.value, "''", "S") & " " ' Where Condition
		arrParam(5) = "발행처"      ' TextBox 명칭 
		 
		arrField(0) = "PARTNER.BP_CD"     ' Field명(0)
		arrField(1) = "PARTNER.BP_NM"     ' Field명(1)

		arrHeader(0) = "발행처"      ' Header명(0)
		arrHeader(1) = "발행처명"     ' Header명(1)
		
		frm1.txtBillToPartyCd.focus

	 Case 10
		If frm1.txtVatType.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config" ' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtVatType.value)    ' Code Condition
		arrParam(3) = ""          ' Name Cindition
		arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
					& " And Config.MINOR_CD = Minor.MINOR_CD" _
					& " And Config.SEQ_NO = 1"    ' Where Condition
		arrParam(5) = "VAT유형"        ' TextBox 명칭 
	  
		arrField(0) = "Minor.MINOR_CD"       ' Field명(0)
		arrField(1) = "Minor.MINOR_NM"       ' Field명(1)
		arrField(2) = "Config.REFERENCE"      ' Field명(2)
		        
		arrHeader(0) = "VAT유형"       ' Header명(0)
		arrHeader(1) = "VAT유형명"       ' Header명(1)
		arrHeader(2) = "VAT율"        ' Header명(2)
		
		frm1.txtVatType.focus
		
	 End Select

	 arrParam(0) = arrParam(5)       ' 팝업 명칭 

	 arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	 IsOpenPop = False

	 If arrRet(0) = "" Then
		Exit Function
	 Else
		Call SetBillHdr(arrRet,iBillHdr)
		OpenBillHdr = True
	 End If 
End Function

'==========================================
Function OpenTaxNo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "세금계산서번호"    
	arrParam(1) = "S_TAX_DOC_NO"     
	arrParam(2) = Trim(frm1.txtTaxBillNo.value)  
	arrParam(3) = ""        
	 
	if frm1.txtBillDt.text = "" then
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  and USED_FLAG=" & FilterVar("C", "''", "S") & "  and convert(char(10),expiry_date,112) >= convert(char(10),getdate(),112)" 
	else
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  and USED_FLAG=" & FilterVar("C", "''", "S") & "  and convert(char(10),expiry_date,112) >= "&UniConvDateToYYYYMMDD(frm1.txtBillDt.text,Parent.gDateFormat,"")     
	end if
	 
	arrParam(5) = "세금계산서번호"    
	     
	arrField(0) = "ED25" & Parent.gColSep & "TAX_DOC_NO"     
	arrField(1) = "DD15" & Parent.gColSep & "expiry_date"
	arrField(2) = "ED15" & Parent.gColSep & "TAX_BOOK_NO"
	arrField(3) = "ED5" & Parent.gColSep & "TAX_BOOK_SEQ"
	         
	arrHeader(0) = "세금계산서번호"
	arrHeader(1) = "유효일"
	arrHeader(2) = "책번호(권)"
	arrHeader(3) = "책번호(호)"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtTaxBillNo.value = arrRet(0)
		lgBlnFlgChgValue = True
	End If 
 
End Function

'=========================================
Function SetSORef(Byval arrRet)

	Call SetRadio

	frm1.txtSoNo.value = arrRet(0)
	frm1.txtBillTypeCd.value = arrRet(1)
	frm1.txtBillTypeNm.value = arrRet(2)

	Call SOHdrQuery()

	frm1.txtBillNo.focus

	lgBlnFlgChgValue = true

End Function

'=========================================
Function SetDnRef(Byval arrRet)

	Call SetRadio

	With frm1
 
		.txtDocCur1.value   = arrRet(4)
		.txtDocCur2.value   = arrRet(4)

		Call CurFormatNumericOCX

		.txtVatType.value   = Trim(arrRet(10))
		.txtVatTypeNm.value   = Trim(arrRet(11))
		.txtVATRate.text   = UNIFormatNumber(arrRet(12), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

		.txtBillToPartyCd.value  = Trim(arrRet(1))
		.txtBillToPartyNm.value  = Trim(arrRet(2))
		.txtSoldtoPartyCd.value  = Trim(arrRet(1))
		.txtSoldtoPartyNm.value  = Trim(arrRet(2))
		.txtPayerCd.value   = Trim(arrRet(1))
		.txtPayerNm.value   = Trim(arrRet(2))

		.txtLocCur.value   = UCase(Parent.gCurrency)

		.txtSalesGrpCd.value  = Trim(arrRet(21))
		.txtSalesGrpNm.value  = Trim(arrRet(22))
		.txtToBizAreaCd.value  = Trim(arrRet(21))
		.txtToBizAreaNm.value  = Trim(arrRet(22))
	 
		.txtPaytermsTxt.value  = Trim(arrRet(23))
	 
		.txtPayTypeCd.value   = Trim(arrRet(5))
		.txtPayTypeNm.value   = Trim(arrRet(6))
		.txtPayTermsCd.value  = Trim(arrRet(7))
		.txtPayTermsNm.value  = Trim(arrRet(8))

		.txtBillTypeCd.value  = Trim(arrRet(15))
		.txtBillTypeNm.value  = Trim(arrRet(16))
	 
		.txtTaxBizAreaCd.value = Trim(arrRet(18))
		.txtTaxBizAreaNm.value = Trim(arrRet(19))
		
		'VAT포함여부 
		If Trim(arrRet(13)) = "2" Then
			.rdoVATIncFlag2.checked = True
			.txtVatIncFlag.value = "2"
		Else
			.rdoVATIncFlag1.checked = True
			.txtVatIncFlag.value = "1"
		End If

		If arrRet(9) = "0" Then
			.txtPayDur.Text  = ""
		Else
			.txtPayDur.Text  = Trim(arrRet(9))
		End If

		'반품 여부 
		.txtRetItemFlag.value = Trim(arrRet(20))

		'약정회전일 
		.txtCreditRotDay.value = Trim(arrRet(3))
	 
		Call CalcPlanIncomeDt()
		
		If UCase(Trim(.txtDocCur1.value)) = UCase(Parent.gCurrency) Then
			ggoOper.SetReqAttr .txtXchgRate, "Q"
		Else
			ggoOper.SetReqAttr .txtXchgRate, "N"
		End If

		.txtBillCommand.value = ""
	End With
 
	frm1.txtBillNo.focus
	lgBlnFlgChgValue = true
End Function

'=========================================
Function GetDNInfo(Byval strDnNo)

	Dim strSelectList, strFromList, strWhereList
	Dim strRs, arrDnInfo
 
	If Trim(strDnNo) = "" Then Exit Function
 
	strSelectList = " * "
	strFromList = " dbo.ufn_s_GetDnInfo ( " & FilterVar(strDnNo, "''", "S") & ", Default, Default) "
	strWhereList = ""
 
	Err.Clear

	'출하정보 Fetch
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrDnInfo = Split(strRs, Chr(11))
		Call SetDnRef(arrDnInfo)
		Exit Function
	Else
		If Err.number <> 0 Then
			Err.Clear 
			Exit Function
		End If
	End if
End Function

'=========================================
Sub GetTaxBizArea(Byval strFlag)

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
		If Len(strBillToParty) > 0 And Len(strSalesGrp) > 0 Then strFlag = "*"
	End if
 
	strSelectList = " * "
	strFromList = " dbo.ufn_s_GetTaxBizArea ( " & FilterVar(strBilltoParty, "''", "S") & ",  " & FilterVar(strSalesGrp, "''", "S") & ",  " & FilterVar(strTaxBizArea, "''", "S") & ",  " & FilterVar(strFlag, "''", "S") & ") "
	strWhereList = ""
 
	Err.Clear
	
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		arrTaxBizArea(0) = arrTemp(1)
		
		arrTaxBizArea(1) = arrTemp(2)
		Call SetBillHdr(arrTaxBizArea, 3)
	Else
	 	' 세금 신고 사업장을 Editing한 경우 
		If strFlag = "NM" Then
			If Not OpenBillHdr(3) Then
				frm1.txtTaxBizAreaCd.value = ""
				frm1.txtTaxBizAreaNm.value = ""
			End if
		End if
	End if
End Sub

'=========================================
Function GetCodeName(ByVal strArg1, ByVal strArg2, ByVal strArg3, ByVal strArg4, ByVal intArg5, ByVal strFlag, ByVal intFlag)

	Dim strSelectList, strFromList, strWhereList
	Dim strRs
	Dim arrRs(3), arrTemp
 
	GetCodeName = False
 
	strSelectList = " * "
	strFromList = " dbo.ufn_s_GetCodeName (" & strArg1 & ", " & strArg2 & ", " & strArg3 & ", " & strArg4 & ", " & intArg5 & ", " & strFlag & ") "
	strWhereList = ""
 
	Err.Clear
 
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		arrRs(0) = arrTemp(1)
		arrRs(1) = arrTemp(2)
		arrRs(2) = arrTemp(3)
		GetCodeName = SetBillHdr(arrRs, intFlag)
	Else
	 
		' 관련 Popup Display
		GetCodeName = OpenBillHdr(intFlag)
	End if
End Function

'=========================================
Function SetLCRef(Byval arrRet)

	Call SetRadio

	With frm1
		.txtLCNo.value = arrRet(0)
		.txtBillTypeCd.value = arrRet(1)
		.txtBillTypeNm.value = arrRet(2)
		.txtSoNo.value = arrRet(3)      
	End With

	If Trim(frm1.txtSoNo.value) <> "" Then 
		Call SOHdrQuery()
	Else
		Call LCHdrQuery()
	End If

	frm1.txtBillNo.focus
	
	lgBlnFlgChgValue = true

End Function

'=========================================
Function SetBillHdr(Byval arrRet,ByVal iBillHdr)

 SetBillHdr = False
 
 Select Case iBillHdr
 Case 0            ' 매출형태 
  frm1.txtBillTypeCd.value = arrRet(0)
  frm1.txtBillTypeNm.value = arrRet(1)
 Case 1            ' 수금처 
  frm1.txtPayerCd.value = arrRet(0)
  frm1.txtPayerNm.value = arrRet(1)
 Case 2            ' 수금사업장 
  frm1.txtToBizAreaCd.value = arrRet(0)
  frm1.txtToBizAreaNm.value = arrRet(1)
 Case 3            ' 세금신고사업장 
  frm1.txtTaxBizAreaCd.value = arrRet(0)
  frm1.txtTaxBizAreaNm.value = arrRet(1)
 Case 4            ' 입금유형 
  frm1.txtPayTypeCd.value = arrRet(0)
  frm1.txtPayTypeNm.value = arrRet(1)
 Case 5            ' 결제방법 
  frm1.txtPayTermsCd.value = arrRet(0)
  frm1.txtPayTermsNm.value = arrRet(1)
 Case 6            ' 양수자 
  frm1.txtBeneficiaryCd.value = arrRet(0)
  frm1.txtBeneficiaryNm.value = arrRet(1)
 Case 7            ' 양도자 
  frm1.txtApplicantCd.value = arrRet(0)
  frm1.txtApplicantNm.value = arrRet(1)
 Case 8            ' 발행처 
  frm1.txtBillToPartyCd.value = arrRet(0)
  frm1.txtBillToPartyNm.value = arrRet(1)
  Call GetTaxBizArea("BP")
 Case 10            ' VAT
  frm1.txtVatType.value = arrRet(0)
  frm1.txtVatTypeNm.value = arrRet(1)
  frm1.txtVatRate.Text = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
 End Select

 SetBillHdr = True
 lgBlnFlgChgValue = true
End Function

'=========================================
Sub PostFlagProtect()
	On Error Resume Next
	    
	Dim elmCnt

	For elmCnt = 1 to frm1.length - 1
		If Left(frm1.elements(elmCnt).getAttribute("tag"),1) = "2" Then
		  Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "Q")
		End If

		If Err.number <> 0 Then Err.Clear
	Next

End Sub

'=========================================
Sub PostFlagRelease()

    On Error Resume Next
    
 Dim elmCnt

 For elmCnt = 1 to frm1.length - 1

  Select Case Left(frm1.elements(elmCnt).getAttribute("tag"),2)
  Case "21","25"
   Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "D")
  Case "22","23"
   Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "N")
  End Select

  If Err.number <> 0 Then Err.Clear
 Next
 
End Sub

'=========================================
Sub RefFlagPR()

	On Error Resume Next
    
	If frm1.txtRefFlag.value = "D" Then
		Call ggoOper.SetReqAttr(frm1.txtBillTypeCd, "N")
		Call ggoOper.SetReqAttr(frm1.txtPayTermsCd, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtBillTypeCd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtPayTermsCd, "Q")
	End If
End Sub

'=========================================
Function SOHdrQuery() 
    
    Err.Clear                                                               
    SOHdrQuery = False                                                      
    
	If LayerShowHide(1) = False Then
		Exit Function 
    End If

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & SoNoHdr        
    strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)    
    
	Call RunMyBizASP(MyBizASP, strVal)          
 
    SOHdrQuery = True                                                       

End Function

'=========================================
Function SOHdrQueryOK() 

	If Trim(frm1.txtRefFlag.value) = "L" Then 
		Call LCHdrQuery()
	End If
    
End Function

'=========================================
Function LCHdrQuery() 
    
    Err.Clear                                                               
    LCHdrQuery = False                                                      
    
	If LayerShowHide(1) = False Then
		Exit Function 
    End If

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & LCNoHdr        
    strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)    
    
	Call RunMyBizASP(MyBizASP, strVal)          
 
    LCHdrQuery = True                                                       

End Function

'=========================================
Function rdoPostFlagFun(ExIm) 

 With frm1    

  Select Case ExIm
  Case "Export"

   If .txtRadioFlag.value = .rdoPostFlagY.value Then
    .rdoPostFlagY.checked = True
   ElseIf .txtRadioFlag.value = .rdoPostFlagN.value Then
    .rdoPostFlagN.checked = True
   End If

  Case "Import"

   If .rdoPostFlagY.checked = True Then
    .txtRadioFlag.value = .rdoPostFlagY.value
   ElseIf .rdoPostFlagN.checked = True Then
    .txtRadioFlag.value = .rdoPostFlagN.value
   End If

  End Select

 End With

End Function

'=========================================
Sub CalcPlanIncomeDt()
    Err.Clear                 

 If Trim(frm1.txtBillDt.Text) = "" Then Exit Sub
 If UNICDbl(frm1.txtCreditRotDay.value) = 0 Then
  frm1.txtPlanIncomeDt.Text = ""
 Else
  frm1.txtPlanIncomeDt.Text = UNIDateAdd("d", frm1.txtCreditRotDay.value, Trim(frm1.txtBillDt.Text), Parent.gDateFormat)
 End If
 
End Sub 

'=========================================
Function CookiePage(ByVal Kubun)
	On Error Resume Next

	Const CookieSplit = 4877      'Cookie Split String : CookiePage Function Use
		 
	Dim strTemp, arrVal

	If Kubun = 1 Then
		WriteCookie CookieSplit , frm1.txtHBillNo.value
	ElseIf Kubun = 0 Then
		
		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
		arrVal = Split(strTemp, Parent.gRowSep)
	
		If arrVal(0) = "" Then Exit Function
		frm1.txtConBillNo.value =  arrVal(0)
	
		Call DbQuery()
	
		WriteCookie CookieSplit , ""
	End If
End Function

'=========================================
Function JumpChgCheck(strJump)
	Dim IntRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(1)
	Call PgmJump(strJump)

End Function

'=========================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim Answer
	'변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 
	If lgBlnFlgChgValue = True Then Answer = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X") '데이타가 변경되었습니다. 계속 하시겠습니까?
	If Answer = VBNO Then Exit Function

	'변경이 없을때 작업진행여부 체크 
	If lgBlnFlgChgValue = False Then Answer = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X") '작업을 수행하시겠습니까?
	If Answer = VBNO Then Exit Function

	BtnSpreadCheck = True

End Function

'=========================================
Sub CurFormatNumericOCX()

	With frm1
		'매출채권금액 
		ggoOper.FormatFieldByObjectOfCur .txtBillAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		'VAT금액 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		'적립금액 
		ggoOper.FormatFieldByObjectOfCur .txtDepositAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		'총매출채권금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotBillAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		'총수금액 
		ggoOper.FormatFieldByObjectOfCur .txtIncomeAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		'통관FOB금액 
		ggoOper.FormatFieldByObjectOfCur .txtAcceptFobAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		
		
		'환율 
		
		ggoOper.FormatFieldByObjectOfCur .txtXchgRate, .txtDocCur1.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		
		
		
	End With

End Sub

'========================================
Function SetRadio()
	Dim blnOldFlag
	blnOldFlag = lgBlnFlgChgValue
	frm1.rdoVATCalcType1.checked = True
	frm1.rdoVATIncFlag1.checked = True

	Call ggoOper.SetReqAttr(window.document.frm1.txtVatType, "D")
	Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag1, "Q")
	Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncFlag2, "Q")
	lgBlnFlgChgValue = blnOldFlag
End Function

Sub LockFieldInit()
    Call FormatDATEField(frm1.txtBillDt)
    Call LockObjectField(frm1.txtBillDt,"R")

    Call FormatDATEField(frm1.txtPlanIncomeDt)
    Call LockObjectField(frm1.txtPlanIncomeDt,"O")

    Call FormatDoubleSingleField(frm1.txtXchgRate)
    Call LockObjectField(frm1.txtXchgRate,"R")

    Call FormatDoubleSingleField(frm1.txtVatRate)
    Call LockObjectField(frm1.txtVatRate,"P")

    Call FormatDoubleSingleField(frm1.txtPayDur)

    Call FormatDoubleSingleField(frm1.txtBillAmt)
    Call LockObjectField(frm1.txtBillAmt ,"P")

    Call FormatDoubleSingleField(frm1.txtBillAmt)
    Call LockObjectField(frm1.txtBillAmt ,"P")

    Call FormatDoubleSingleField(frm1.txtBillAmtLoc)
    Call LockObjectField(frm1.txtBillAmtLoc,"P")

    Call FormatDoubleSingleField(frm1.txtVatAmt)
    Call LockObjectField(frm1.txtVatAmt,"P")

    Call FormatDoubleSingleField(frm1.txtVatLocAmt)
    Call LockObjectField(frm1.txtVatLocAmt,"P")

    Call FormatDoubleSingleField(frm1.txtDepositAmt)
    Call LockObjectField(frm1.txtDepositAmt,"P")

    Call FormatDoubleSingleField(frm1.txtDepositAmtLoc)
    Call LockObjectField(frm1.txtDepositAmtLoc,"P")

    Call FormatDoubleSingleField(frm1.txtTotBillAmt)
    Call LockObjectField(frm1.txtTotBillAmt,"P")

    Call FormatDoubleSingleField(frm1.txtTotBillAmtLoc)
    Call LockObjectField(frm1.txtTotBillAmtLoc,"P")

    Call FormatDoubleSingleField(frm1.txtIncomeAmt)
    Call LockObjectField(frm1.txtIncomeAmt,"P")

    Call FormatDoubleSingleField(frm1.txtIncomeLocAmt)
    Call LockObjectField(frm1.txtIncomeLocAmt,"P")

    Call FormatDoubleSingleField(frm1.txtAcceptFobAmt)
    Call LockObjectField(frm1.txtAcceptFobAmt,"P")
End Sub

Sub LockFieldAll()
    Call LockObjectField(frm1.txtBillDt,"P")
    Call LockObjectField(frm1.txtPlanIncomeDt,"P")
    Call LockObjectField(frm1.txtXchgRate,"P")
    Call LockObjectField(frm1.txtPayDur,"P")

	Call ggoOper.SetReqAttr(frm1.chkSoNo, "Q")
	Call ggoOper.SetReqAttr(frm1.chkTaxNo, "Q")

	Call LockHTMLField(frm1.txtBillNo, "P")	
	Call LockHTMLField(frm1.rdoVATCalcType1, "P")	
	Call LockHTMLField(frm1.rdoVATCalcType2, "P")	
	Call LockHTMLField(frm1.txtVatType, "P")	
	Call LockHTMLField(frm1.rdoVatIncflag1, "P")	
	Call LockHTMLField(frm1.rdoVatIncflag2, "P")	

	Call LockHTMLField(frm1.txtBillTypeCd, "P")	
	Call LockHTMLField(frm1.txtBillToPartyCd, "P")	
	Call LockHTMLField(frm1.txtPayerCd, "P")	
	Call LockHTMLField(frm1.txtPayTermsCd, "P")	

	Call LockHTMLField(frm1.txtToBizAreaCd , "P")	
	Call LockHTMLField(frm1.txtTaxBizAreaCd , "P")	

	Call LockHTMLField(frm1.txtBeneficiaryCd, "P")	
	Call LockHTMLField(frm1.txtApplicantCd, "P")	
	Call LockHTMLField(frm1.txtPayTypeCd, "P")	
	Call LockHTMLField(frm1.txtVatType, "P")	
	Call LockHTMLField(frm1.txtPaytermsTxt, "P")	
	Call LockHTMLField(frm1.txtRemark, "P")	
End Sub

Sub LockFieldQuery()
	If frm1.txtRefFlag.value = "D" Then
		Call LockHTMLField(frm1.txtPayTermsCd, "R")	

		If UNICDbl(frm1.txtSts.value) < 3 Then
			Call LockHTMLField(frm1.txtBillTypeCd, "P")	
		Else
			Call LockHTMLField(frm1.txtBillTypeCd, "R")	
		End If
	Else
		Call LockHTMLField(frm1.txtBillTypeCd, "P")	
		Call LockHTMLField(frm1.txtPayTermsCd, "P")	
	End If

	Call LockHTMLField(frm1.txtBillNo, "P")	
	Call LockHTMLField(frm1.chkSoNo, "P")
		
	If UCase(Trim(frm1.txtDocCur1.value)) = UCase(Parent.gCurrency) Then
	    Call LockObjectField(frm1.txtXchgRate,"P")
	Else
	    Call LockObjectField(frm1.txtXchgRate,"R")
	End If
		
	If frm1.rdoVATCalcType1.checked Then
		Call LockHTMLField(frm1.txtVatType, "O")	
		Call LockHTMLField(frm1.rdoVatIncflag1, "P")	
		Call LockHTMLField(frm1.rdoVatIncflag2, "P")	
	Else
		Call LockHTMLField(frm1.txtVatType, "R")	
		
		If UNICDbl(frm1.txtSts.value) < 3 Then
			Call LockHTMLField(frm1.rdoVATCalcType1, "P")	
			Call LockHTMLField(frm1.rdoVATCalcType2, "P")	
		Else		
			Call LockHTMLField(frm1.rdoVatIncflag1, "O")	
			Call LockHTMLField(frm1.rdoVatIncflag2, "O")	
		End If
	End If

	Call LockHTMLField(frm1.txtBillToPartyCd, "R")	
	Call LockHTMLField(frm1.txtPayerCd, "R")	
	Call LockHTMLField(frm1.txtTaxBizAreaCd, "R")	

    Call LockObjectField(frm1.txtBillDt,"R")
    Call LockObjectField(frm1.txtPlanIncomeDt,"O")

	Call LockHTMLField(frm1.txtBeneficiaryCd, "O")	
	Call LockHTMLField(frm1.txtApplicantCd , "O")	
	Call LockHTMLField(frm1.txtPayTypeCd , "O")	
	Call LockHTMLField(frm1.txtPaytermsTxt  , "O")	
	Call LockHTMLField(frm1.txtRemark  , "O")	
End Sub

Sub LockFieldNew()
	Call LockHTMLField(frm1.chkSoNo, "O")
	Call LockHTMLField(frm1.chkTaxNo, "O")

	Call LockHTMLField(frm1.txtBillNo, "O")	
	Call LockHTMLField(frm1.txtBillTypeCd, "R")	
	Call LockHTMLField(frm1.txtBillToPartyCd, "R")	
	Call LockHTMLField(frm1.txtPayerCd, "R")	
	Call LockHTMLField(frm1.txtPayTermsCd, "R")	
	Call LockHTMLField(frm1.txtToBizAreaCd , "R")	
	Call LockHTMLField(frm1.txtTaxBizAreaCd , "R")	
	Call LockHTMLField(frm1.txtBeneficiaryCd, "O")	
	Call LockHTMLField(frm1.txtApplicantCd, "O")	
	Call LockHTMLField(frm1.txtPayTypeCd, "O")	
	Call LockHTMLField(frm1.txtPaytermsTxt, "O")	
	Call LockHTMLField(frm1.txtRemark, "O")	
End Sub

Function CheckField()
	
	CheckField = False

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		If Not chkFieldByCell(frm1.txtBillTypeCd, "A", "1")	Then Exit Function
		If Not chkFieldByCell(frm1.txtBillToPartyCd, "A", "1") Then Exit Function
		If Not chkFieldByCell(frm1.txtVatType, "A", "1") Then Exit Function
		If Not chkFieldByCell(frm1.txtPayTermsCd, "A", "1") Then Exit Function

		If Not ChkFieldLengthByCell(frm1.txtBillTypeCd, "A", "1") Then Exit Function        
		If Not ChkFieldLengthByCell(frm1.txtBillToPartyCd, "A", "1") Then Exit Function
		If Not ChkFieldLengthByCell(frm1.txtVatType, "A", "1") Then Exit Function
		If Not ChkFieldLengthByCell(frm1.txtPayTermsCd, "A", "1") Then Exit Function        
	End If

    If Not chkFieldByCell(frm1.txtBillDt , "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtPayerCd, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtToBizAreaCd, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtXchgRate, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtTaxBizAreaCd , "A", "1") Then Exit Function

    If Not ChkFieldLengthByCell(frm1.txtPaytermsTxt, "A", "1") Then Exit Function
    If Not ChkFieldLengthByCell(frm1.txtRemark, "A", "1") Then Exit Function

	CheckField = True

End Function
'=========================================
Sub Form_Load()

	Call LoadInfTB19029                                                     
	Call AppendNumberPlace("6","3","0")
'	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
'	Call ggoOper.LockField(Document, "N")                                   

	Call LockFieldInit
	Call SetTaxBillNoMgmtMeth
	Call SetDefaultVal
	
	Call ggoOper.SetReqAttr(window.document.frm1.txtVatType, "D")
	Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag1, "Q")
	Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncFlag2, "Q")

	Call ggoOper.SetReqAttr(window.document.frm1.rdoPostFlagY, "Q")
	Call ggoOper.SetReqAttr(window.document.frm1.rdoPostFlagN, "Q")

	Call SetToolBar("11100000000011")          
	Call InitVariables                                                      
	 
	Call CookiePage(0)
	Call ChangeTabs(TAB1)             'Because Textbox OCX Formatfield Display
	gIsTab     = "Y" : gTabMaxCnt = 3
End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================
Sub btnPostFlag_OnClick()
     
     If frm1.chkTaxNo.checked and Not gBlnTaxbillnoMgmtMeth Then
          Call DisplayMsgBox("205626", "x", "x", "x")
          '세금계산서 방법이 설정되지 않았거나 2개이상 설정되었습니다.
          Exit Sub
     End If       
 
     If BtnSpreadCheck = False Then Exit Sub
 
     Dim strVal

     frm1.txtInsrtUserId.value = Parent.gUsrID 
   
     If LayerShowHide(1) = False Then
          Exit Sub
     End If

     strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag        
     strVal = strVal & "&txtBillNo=" & Trim(frm1.txtBillNo.value)    '☜: 조회 조건 데이타 
     strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)
     strVal = strVal & "&txtChangeOrgId=" & Parent.gChangeOrgId

     Call RunMyBizASP(MyBizASP, strVal)           
 
End Sub

'==========================================
Sub btnGLView_OnClick()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent
 	
	If Trim(frm1.txtGLNo.value) <> "" Then
		 arrParam(0) = Trim(frm1.txtGLNo.value) '회계전표번호 
		 arrParam(1) = Trim(frm1.txtBillNo.value) 'Reference번호 
		 
		 if arrParam(0) = "" THEN Exit Sub
		 
		 iCalledAspName = AskPRAspName("a5120ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If

		 arrRet = window.showModalDialog(iCalledAspName , Array(window.parent,arrParam), _
		      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		      
	ElseIf Trim(frm1.txtTempGLNo.value) <> "" Then
	     arrParam(0) = Trim(frm1.txtTempGLNo.value) '결의전표번호 
	     arrParam(1) = Trim(frm1.txtBillNo.value) 'Reference번호 
	 
	     if arrParam(0) = "" THEN Exit Sub
	     
	     iCalledAspName = AskPRAspName("a5130ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If
		 
	     arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else 
	     Call DisplayMsgBox("205154", "X", "X", "X")
	End If 
	     lblnWinEvent = False
End Sub

'==========================================
Sub btnPreRcptView_OnClick()
     Dim arrRet
     Dim arrParam(4)
	 Dim iCalledAspName
	 Dim IntRetCD
	 Dim lblnWinEvent
 
     arrParam(0) = Trim(frm1.txtBillDt.Text)    '매출채권일 
     arrParam(1) = Trim(frm1.txtSoldToPartyCd.value)  '주문처 
     arrParam(2) = Trim(frm1.txtSoldToPartyNm.value)  '주문처 
     arrParam(3) = Trim(frm1.txtDocCur1.value)   '화폐 
     arrParam(4) = ""         '선수금번호 
 
     iCalledAspName = AskPRAspName("s5111ra7")
		 
	 If Trim(iCalledAspName) = "" Then
	      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111ra7", "X")
		  lblnWinEvent = False
		       Exit Sub
		  End If
     arrRet = window.showModalDialog(iCalledAspName & "?txtFlag=BH&txtCurrency=" & frm1.txtDocCur1.value, Array(window.parent,arrParam), _
     "dialogWidth=860px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
     lblnWinEvent = False	
End Sub

'========================================
Sub chkSoNo_OnClick()
     lgBlnFlgChgValue = True
End Sub

'========================================
Sub rdoVATCalcType1_OnClick()
     lgBlnFlgChgValue = True
     frm1.txtVatCalcType.value = "1"
 
     Call ggoOper.SetReqAttr(window.document.frm1.txtVatType, "D")

     ' 부가세 계산방법이 개별인 경우에는 부가세 포함여부 사용자가 수정 불가 
     Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag1, "Q")
     Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncFlag2, "Q")
     frm1.rdoVatIncFlag1.checked = True
End Sub

'========================================
Sub rdoVATCalcType2_OnClick()
     lgBlnFlgChgValue = True
     frm1.txtVatCalcType.value = "2"

     Call ggoOper.SetReqAttr(window.document.frm1.txtVatType, "N")

     ' 부가세 계산방법이 통합인 경우에는 부가세 포함여부 사용자가 변경 가능 
    Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag1, "N")
    Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag2, "N")
End Sub

'부가세 포함여부 
'========================================
Sub rdoVatIncFlag1_OnClick()
     lgBlnFlgChgValue = True 
End Sub

'부가세포함여부 
'========================================
Sub rdoVatIncFlag2_OnClick()
 lgBlnFlgChgValue = True
End Sub

'매출채권일 
'========================================
Sub txtBillDt_Change()
 If Trim(frm1.txtBillDt.Text) <> "" And Trim(frm1.txtBillCommand.value) = "" Then Call CalcPlanIncomeDt()
 lgBlnFlgChgValue = True
End Sub

'수금예정일 
'========================================
Sub txtPlanIncomeDt_Change()
 lgBlnFlgChgValue = True
End Sub

'환율 
'========================================
Sub txtXchgRate_Change()
 lgBlnFlgChgValue = True
End Sub

'결제기간 
'========================================
Sub txtPayDur_Change()
 lgBlnFlgChgValue = True
End Sub

'==========================================
Sub txtBillDt_DblClick(Button)
	If Trim(frm1.txtBillToPartyCd.value) = "" Then
		Call DisplayMsgBox("205151", "X", "X", "X")
		'  Msgbox "수주참조 또는 LC참조를 먼저 하세요"
		Exit Sub
	End If

	If Button = 1 Then
		frm1.txtBillDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillDt.Focus
	End If

End Sub

'==========================================
Sub txtPlanIncomeDt_DblClick(Button)

	If Trim(frm1.txtBillToPartyCd.value) = "" Then
		Call DisplayMsgBox("205151", "X", "X", "X")
		'  Msgbox "수주참조 또는 LC참조를 먼저 하세요"
		Exit Sub
	End If
	 
	If Button = 1 Then
		frm1.txtPlanIncomeDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPlanIncomeDt.Focus
	End If

End Sub

'==========================================
Function txtBillTypeCd_OnChange()
	If Trim(frm1.txtBillCommand.value) = "" Then
		Dim strCode
		 
		strCode = Trim(frm1.txtBillTypeCd.value)
		If strCode <> "" Then
			strCode = " " & FilterVar(strCode, "''", "S") & ""
			If Not GetCodeName(strCode, "" & FilterVar("N", "''", "S") & " ", "" & FilterVar("N", "''", "S") & " ", "default", "default", "" & FilterVar("BT", "''", "S") & "", 0) Then
				txtBillTypeCd_OnChange = False
				frm1.txtBillTypeCd.value = ""
				frm1.txtBillTypeNm.value = ""
			Else
				txtBillTypeCd_OnChange = True
			End If
		Else
			frm1.txtBillTypeNm.value = ""
		End If
	End If 
End Function

'==========================================
Sub txtBillToPartyCd_OnChange()
	If Trim(frm1.txtBillCommand.value) = "" Then
		If Trim(frm1.txtBillToPartyCd.value) = "" Then
			'frm1.txtBillToPartyNm.value = ""
		Else
			' 발행처 변경시 관련 세금신고 사업장 Fetch
			Call GetTaxBizArea("BP")
		End if
	End If 
End Sub

'==========================================
Function txtPayTermsCd_OnChange()
	'---입금유형 
	frm1.txtPayTypeCd.value = ""
	frm1.txtPayTypeNm.value = ""

	'---결제기간 
	frm1.txtPayDur.text = 0
End Function

'==========================================
Function txtVatType_OnChange()
	If Trim(frm1.txtBillCommand.value) = "" Then
		Dim strCode
		 
		strCode = Trim(frm1.txtVatType.value)
		If strCode <> "" Then
			strCode = " " & FilterVar(strCode, "''", "S") & ""
			If Not GetCodeName("" & FilterVar("B9001", "''", "S") & "", strCode, "default", "default", 1, "" & FilterVar("CF", "''", "S") & "", 10) Then
				txtVatType_OnChange = False
				frm1.txtVatType.value = ""
				frm1.txtVatTypeNm.value = ""
				frm1.txtVatRate.Text = "0"
			Else
				txtVatType_OnChange = True
			End If
		Else
			frm1.txtVatType.value = ""
			frm1.txtVatTypeNm.value = ""
			frm1.txtVatRate.Text = "0"
		End If
	End If 
End Function

'==========================================
Function txtTaxBizAreaCd_OnChange()
	If Trim(frm1.txtBillCommand.value) = "" Then
		If Trim(frm1.txtTaxBizAreaCd.value) = "" Then
			frm1.txtTaxBizAreaNm.value = ""
		Else
			Call GetTaxBizArea("NM")
			txtTaxBizAreaCd_OnChange = False
			If frm1.txtTaxBizAreaCd.value <> "" Then frm1.txtPaytermsTxt.focus
		End if
	End If
End Function

' 세금계산서 자동발행 여부에 따라 관련입력항목 Change
'==========================================
Sub chkTaxNo_OnClick()
	lgBlnFlgChgValue = True
	' 세금계산서 관리방법 Check 
	if frm1.chkTaxNo.checked Then
		Call chkTaxBillNoCreatedMeth
	else
		frm1.txtTaxBillNo.value = ""
		Call ggoOper.SetReqAttr(window.document.frm1.txtTaxBillNo, "Q")
		window.document.frm1.btnBillTaxNo.disabled = True
	end if
End Sub

' 세금계산서 자동발행 여부에 따라 관련입력항목 Change
'==========================================
Sub chkTaxBillNoCreatedMeth()
	lgBlnFlgChgValue = True
	' 세금계산서 관리방법 Check 
	if frm1.chkTaxNo.checked Then
		if gBlnTaxbillnoMgmtMeth Then
			Select Case gtxtCreatedMeth
			Case "A"
			Call ggoOper.SetReqAttr(window.document.frm1.txtTaxBillNo, "Q")
			window.document.frm1.btnBillTaxNo.disabled = True
			Case "M"
			Call ggoOper.SetReqAttr(window.document.frm1.txtTaxBillNo, "N") 
			window.document.frm1.btnBillTaxNo.disabled = True
			Case "P"
			Call ggoOper.SetReqAttr(window.document.frm1.txtTaxBillNo, "Q") 
			window.document.frm1.btnBillTaxNo.disabled = False 
			Case "X"
			Call ggoOper.SetReqAttr(window.document.frm1.txtTaxBillNo, "D") 
			window.document.frm1.btnBillTaxNo.disabled = False   
			End Select 
		Else
			Call ggoOper.SetReqAttr(window.document.frm1.txtTaxBillNo, "Q")
			window.document.frm1.btnBillTaxNo.disabled = True
		end if
	else
		' 저장된 데이터에 대한 처리 
		Call ggoOper.SetReqAttr(window.document.frm1.txtTaxBillNo, "Q")
		window.document.frm1.btnBillTaxNo.disabled = True
	end if
End Sub

'========================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If Not chkFieldByCell(frm1.txtConBillNo, "A", 1) Then Exit Function 

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Call ggoOper.ClearField(Document, "2")          
			Exit Function
		End If
     End If
    
    Call InitVariables               
	Call DbQuery
       
    FncQuery = True                
        
End Function

'========================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                      
    Call SetToolBar("11100000000011")           

	With frm1
		'입금액 
		.txtIncomeAmt.Text = 0
		'환율 
		.txtXchgRate.Text = 1
		'매출채권금액 
		.txtBillAmt.Text = 0
		'부가세율 
		.txtVatRate.Text = 0
		'매출채권VAT금액 
		.txtVatAmt.Text = 0
		'통관FOB금액 
		.txtAcceptFobAmt.Text = 0
		'결제기간 
		.txtPayDur.Text = 0
	End With

	Call LockFieldNew
	Call LockFieldInit
    Call SetDefaultVal
'	Call PostFlagRelease()
'   Call ggoOper.LockField(Document, "N")                                       
    Call SetRadio
    Call InitVariables               
 
    FncNew = True                

End Function

'========================================
Function FncDelete() 
    FncDelete = False               
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then Exit Function

    Call DbDelete
    
    FncDelete = True               
End Function

'========================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               

	If ValidDateCheck(frm1.txtBillDt, frm1.txtPlanIncomeDt) = False Then Exit Function

	Call rdoPostFlagFun("Import")
	    
	If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If

	If CheckField = False Then Exit Function

'	If Not chkField(Document, "2") Then          
'		If gPageNo > 0 Then
'			gSelframeFlg = gPageNo
'		End If
'		Exit Function
'	End If 


    Call DbSave
    
    FncSave = True                                                          
    
End Function

'========================================
Function FncPrint() 
    Call FncPrint()
End Function

'========================================
Function FncPrev() 
    Dim strVal
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    End If

	 If   LayerShowHide(1) = False Then Exit Function 

	 frm1.txtPrevNext.value = "P"

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       
    strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtBillNo.value)   
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  
         
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================
Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If
 
	frm1.txtPrevNext.value = "N"

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       
    strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtBillNo.value)   
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================
Function FncExcel() 
	Call FncExport(Parent.C_SINGLE)
End Function

'========================================
Function FncFind()
	On Error Resume Next                                                          
    Err.Clear                                                                     

    FncFind = False                                                               
     
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                               
    
    If Err.number = 0 Then	 
       FncFind = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'========================================
Function DbDelete() 
    Err.Clear                                                               
    
    DbDelete = False              

   
	If LayerShowHide(1) = False Then
		Exit Function 
	End If
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003       
    strVal = strVal & "&txtBillNo=" & Trim(frm1.txtBillNo.value)   '☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)          
 
    DbDelete = True                                                         

End Function

'========================================
Function DbDeleteOk()
     Call Initvariables()
     Call MainNew()
End Function

'========================================
Function DbQuery() 
    
    Err.Clear                                                               

    DbQuery = False                                                         
   
    If LayerShowHide(1) = False Then
         Exit Function 
    End If

    Dim strVal
    
    frm1.txtPrevNext.value = "Q"
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       
    strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtConBillNo.value)  
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  
    
    Call RunMyBizASP(MyBizASP, strVal)          
 
    DbQuery = True                                                          

End Function

'========================================
Function DbQueryOk()

	On Error Resume Next

	lgIntFlgMode = Parent.OPMD_UMODE
	    
	Call SetToolBar("11111000110111")          

	Call rdoPostFlagFun("Export")

	'수금예정일 추가 
	if UniConvDateToYYYYMMDD(frm1.txtPlanIncomeDt.Text,Parent.gDateFormat,"-") = "2999-12-31" then
		frm1.txtPlanIncomeDt.Text = ""    
	end if

	If frm1.txtchkTaxNo.value = "Y" Then
		frm1.chkTaxNo.checked = True
		Call chkTaxNo_OnClick()
	End If
 
	If UNICDbl(frm1.txtSts.value) < 3 Then
		frm1.btnPostFlag.disabled = False
	Else
		frm1.btnPostFlag.disabled = True
	End If

	frm1.txtBillCommand.value = ""

	If frm1.txtRadioFlag.value = "Y" Then
		Call LockFieldAll()
'		Call PostFlagProtect()
	ElseIf frm1.txtRadioFlag.value = "N" Then
		Call LockFieldQuery()
		Call chkTaxBillNoCreatedMeth 
	End If

	lgBlnFlgChgValue = False

End Function

'========================================
Function DbSave() 

    Err.Clear                

	DbSave = False               

   
	If LayerShowHide(1) = False Then
		Exit Function 
    End If

    Dim strVal

 With frm1
  .txtMode.value = Parent.UID_M0002
  .txtFlgMode.value = lgIntFlgMode
  .txtInsrtUserId.value = Parent.gUsrID 
  .txtUpdtUserId.value = Parent.gUsrID

  If .chkSoNo.checked = True Then
   .txtChkSoNo.value = "Y"
  Else
   .txtChkSoNo.value = "N"
  End If

  If .chkTaxNo.checked = True Then
   .txtchkTaxNo.value = "Y"
  Else
   .txtchkTaxNo.value = "N"
  End If

  ' vat적용기준 
  If .rdoVatCalcType1.checked Then
   .txtVatCalcType.value = "1"
  Else
   .txtVatCalcType.value = "2"
  End If

  ' vat포함여부 
  If .rdoVatIncFlag1.checked = True Then
   .txtVatIncFlag.value = "1"
  Elseif .rdoVatIncFlag2.checked = True Then
   .txtVatIncFlag.value = "2"
  End If

  Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
 
 End With
 
    DbSave = True                                                           

End Function

'========================================
Function DbSaveOk()
    Call InitVariables
    Call MainQuery()
End Function

