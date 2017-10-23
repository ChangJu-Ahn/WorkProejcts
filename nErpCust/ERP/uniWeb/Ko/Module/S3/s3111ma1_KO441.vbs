'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const DNCheck = "DNCheck"

Const BIZ_PGM_ID = "s3111mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_OnLine_ID = "s3111ab1.asp"											'☆: OnLine ADO 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "s3112ma1_ko441"

Const TAB1 = 1																	'☜: Tab의 위치 
Const TAB2 = 2

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim IsOpenPop						' Popup
Dim gSelframeFlg
Dim lsClickCfmYes
Dim lsClickCfmNo
Dim PrevRadioFlag					'☜: Radio Button의 이전값 
Dim PrevRadioType					'☜: Radio Button의 이전값 
Dim PrevRadioDnParcel
Dim arrCollectVatType

'==========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False
	lsClickCfmYes = False
	lsClickCfmNo = False

End Sub

'==========================================================================================================
Sub SetDefaultVal()

On Error Resume Next

	With frm1
		.txtConSo_no.focus
		.rdoCfm_flag2.checked = True
		.rdoPrice_flag1.checked = True		
		.txtSo_dt.text = EndDate		
		.txtCust_po_dt.text = EndDate		
		.txtValid_dt.text = EndDate		
		.txtContract_dt.text = EndDate
		.txtRadioFlag.value = .rdoCfm_flag2.value 
		.txtRadioType.value = .rdoPrice_flag1.value 
		.txtSales_Grp.value = parent.gSalesGrp
		.txtTo_Biz_Grp.value = parent.gSalesGrp
		.txtBeneficiary.value = parent.gCompany
		.txtBeneficiary_nm.value = parent.gCompanyNm 
		.txtDoc_cur.value = parent.gCurrency
		.txtXchg_rate.Text = 0
		.btnDNCheck.disabled = True
		.btnConfirm.disabled = True
		.btnConfirm.value = "확정처리"

		.txtSales_Grp.value = parent.gSalesGrp

	End With
	
	PrevRadioFlag = frm1.txtRadioFlag.value
	PrevRadioType = frm1.txtRadioType.value

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
	End If
	lgBlnFlgChgValue = False

End Sub

'==========================================================================================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

End Function

'==========================================================================================================
Function ClickTab2()

	If SoTypeExportCheck = False Then Exit Function		
	
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	
	gSelframeFlg = TAB2
										
	If Trim(frm1.txtValid_dt.text) = "" Then frm1.txtValid_dt.text = Trim(frm1.txtReq_dlvy_dt.text)

End Function


'==========================================================================================================
Function OpenSORef()
	Dim iCalledAspName
	Dim strRet
		
	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 
				
	If UCase(Trim(frm1.txtSoTypeRetItemFlag.value)) <> "Y" Then
		Call DisplayMsgBox("203155", "x", "x", "x")
		Exit Function
	End If	  

	If IsOpenPop = True Then Exit Function
		
	iCalledAspName = AskPRAspName("s3112ra5")	
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra5", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
				
	strRet = window.showModalDialog(iCalledAspName, Array(Window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
				
	If strRet = "" Then
		Exit Function
	Else
		Call SetSORef(strRet)
	End If	
End Function	


'==========================================================================================================
Function OpenPrjRef()
	Dim iCalledAspName
	Dim strRet
	Dim arrParam(2)
		
	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 
				
	If IsOpenPop = True Then Exit Function
		
	iCalledAspName = AskPRAspName("s3111ra9_KO441")	
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111ra9_KO441", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True

	arrParam(0) = ""
				
	strRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
				
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetPrjRef(strRet)
	End If	
	
End Function	


'==========================================================================================================
Function OpenSoNo(strSoNo)
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
			
	iCalledAspName = AskPRAspName("s3111pa1_KO441")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1_KO441", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True

	strRet = window.showModalDialog(iCalledAspName,Array(Window.parent, "SO_REG"), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		strSoNo.value = strRet
	End If
	
	frm1.txtConSo_no.focus 	

End Function

'==========================================================================================================
Function OpenRequried(ByVal iRequried)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If lsClickCfmYes = True Then Exit Function

	IsOpenPop = True

	Select Case iRequried
	Case 0												
		If lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		If UCase(frm1.txtSo_Type.className) = parent.UCN_PROTECTED Then 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "수주형태"					
		arrParam(1) = "S_SO_TYPE_CONFIG"				
		arrParam(2) = Trim(frm1.txtSo_Type.value)		
		arrParam(3) = Trim(frm1.txtSo_TypeNm.value)		
		arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  and SO_MGMT_FLAG <> " & FilterVar("N", "''", "S") & "  and STO_FLAG = " & FilterVar("N", "''", "S") & " "				
		arrParam(5) = "수주형태"		
	
		arrField(0) = "SO_TYPE"			
	    arrField(1) = "SO_TYPE_NM"		
	    arrField(2) = "EXPORT_FLAG"		
	    arrField(3) = "RET_ITEM_FLAG"	
	    arrField(4) = "AUTO_DN_FLAG"
	    arrField(5) = "CI_FLAG"	
			    
	    arrHeader(0) = "수주형태"					
	    arrHeader(1) = "수주형태명"					
	    arrHeader(2) = "수출여부"					
	    arrHeader(3) = "반품여부"					
	    arrHeader(4) = "자동출하생성여부"		
	    arrHeader(5) = "통관여부"			
	    
		frm1.txtSo_Type.focus 
	Case 1												
		If lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		If UCase(frm1.txtSales_Grp.className) = parent.UCN_PROTECTED Then 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "영업그룹"					
		arrParam(1) = "B_SALES_GRP"						
		arrParam(2) = Trim(frm1.txtSales_Grp.value)		
		arrParam(3) = Trim(frm1.txtSales_GrpNm.value)	
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
		arrParam(5) = "영업그룹"					
		
	    arrField(0) = "SALES_GRP"						
	    arrField(1) = "SALES_GRP_NM"					
	    
	    arrHeader(0) = "영업그룹"					
	    arrHeader(1) = "영업그룹명"					
		
		frm1.txtSales_Grp.focus 
	Case 2												

		If UCase(frm1.txtIncoTerms.className) = parent.UCN_PROTECTED Then 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "가격조건"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtIncoTerms.value)		
		arrParam(3) = Trim(frm1.txtIncoTerms_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9006", "''", "S") & ""				
		arrParam(5) = "가격조건"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "가격조건"					
	    arrHeader(1) = "가격조건명"					
		
		frm1.txtIncoTerms.focus 
	Case 3												

		If UCase(frm1.txtTo_Biz_Grp.className) = parent.UCN_PROTECTED Then 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "수금그룹"					
		arrParam(1) = "B_SALES_GRP"						
		arrParam(2) = Trim(frm1.txtTo_Biz_Grp.value)	
		arrParam(3) = Trim(frm1.txtTo_Biz_GrpNm.value)	
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
		arrParam(5) = "수금그룹"					
		
	    arrField(0) = "SALES_GRP"						
	    arrField(1) = "SALES_GRP_NM"					
	    
	    arrHeader(0) = "수금그룹"					
	    arrHeader(1) = "수금그룹명"					
	    
	    frm1.txtTo_Biz_Grp.focus 
	End Select

	arrParam(3) = ""			'☜: [Condition Name Delete]
    
	If iRequried = 0 Then 
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	End If
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetRequried(arrRet,iRequried)
	End If	
End Function

'==========================================================================================================
Function OpenBp(ByVal iRequried)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If lsClickCfmYes = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_BIZ_PARTNER"							
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "							
		
	arrField(0) = "BP_CD"									
	arrField(1) = "BP_NM"									
	arrField(2) = "BP_RGST_NO"
	
	arrHeader(2) = "사업자등록번호"
	    
	Select Case iRequried
	Case 0	
		If frm1.txtSold_to_party.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If
													
		If lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		If UCase(frm1.txtSold_to_party.className) = parent.UCN_PROTECTED Then 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(2) = Trim(frm1.txtSold_to_party.value)		
		arrParam(3) = Trim(frm1.txtSold_to_partyNm.value)	
		arrParam(5) = "주문처"							
		
	    arrHeader(0) = "주문처"							
	    arrHeader(1) = "주문처명"						
		
		frm1.txtSold_to_party.focus 
	Case 1												
		If frm1.txtManufacturer.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(2) = Trim(frm1.txtManufacturer.value)		
		arrParam(3) = Trim(frm1.txtManufacturer_nm.value)	
		arrParam(5) = "제조자"							
		
	    arrHeader(0) = "제조자"							
	    arrHeader(1) = "제조자명"						
		
		frm1.txtManufacturer.focus 
	Case 2												
		If frm1.txtAgent.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(2) = Trim(frm1.txtAgent.value)				
		arrParam(3) = Trim(frm1.txtAgent_nm.value)			
		arrParam(5) = "대행자"							
		
	    arrHeader(0) = "대행자"							
	    arrHeader(1) = "대행자명"						
		
		frm1.txtAgent.focus 
	Case 3												
		If frm1.txtBeneficiary.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(2) = Trim(frm1.txtBeneficiary.value)		
		arrParam(3) = Trim(frm1.txtBeneficiary_nm.value)	
		arrParam(5) = "수출자"							
		
	    arrHeader(0) = "수출자"							
	    arrHeader(1) = "수출자명"						
		
		frm1.txtBeneficiary.focus 
	End Select

	arrParam(0) = arrParam(5)								
	arrParam(3) = ""			'☜: [Condition Name Delete]
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBp(arrRet,iRequried)
	End If	
End Function


'==========================================================================================================
Function OpenOption(ByVal iOption)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iOption
	Case 0												
		If frm1.txtShip_to_party.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If Trim(frm1.txtSold_to_party.value) = "" Then
			Call DisplayMsgBox("203150","X","X","X")
			'MsgBox "주문처를 먼저 입력하세요!"
			frm1.txtSold_to_party.focus 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "납품처"												
		arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
		arrParam(2) = Trim(frm1.txtShip_to_party.value)						
		arrParam(3) = Trim(frm1.txtShip_to_partyNm.value)						
		arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SSH", "''", "S") & " " _
						& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE <= " & FilterVar("CS", "''", "S") & " " _
						& "AND PARTNER_FTN.BP_CD= " & FilterVar(frm1.txtSold_to_party.value, "''", "S") 	
		arrParam(5) = "납품처"							
		
	    arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"			
	    arrField(1) = "PARTNER.BP_NM"						
	    arrField(2) = "PARTNER_FTN.BP_CD"
	    arrField(3) = "PARTNER.BP_RGST_NO"	    					
	    arrField(4) = "PARTNER_FTN.PARTNER_FTN"				
	    	    
	    arrHeader(0) = "납품처"							
	    arrHeader(1) = "납품처명"						
	    arrHeader(2) = "거래처코드"						
	    arrHeader(3) = "사업자등록번호"						
	    arrHeader(4) = "거래처타입"						

		frm1.txtShip_to_party.focus 
	Case 1												
		If frm1.txtBill_to_party.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True or lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		If Trim(frm1.txtSold_to_party.value) = "" Then
			Call DisplayMsgBox("203150","X","X","X")
			'MsgBox "주문처를 먼저 입력하세요!"
			frm1.txtSold_to_party.focus 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "발행처"												
		arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
		arrParam(2) = Trim(frm1.txtBill_to_party.value)							
		arrParam(3) = Trim(frm1.txtBill_to_partyNm.value)						
		arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SBI", "''", "S") & " " _
						& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE <= " & FilterVar("CS", "''", "S") & " " _
						& "AND PARTNER_FTN.BP_CD= " & FilterVar(frm1.txtSold_to_party.value, "''", "S") 	
		arrParam(5) = "발행처"							
		
	    arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"			
	    arrField(1) = "PARTNER.BP_NM"						
	    arrField(2) = "PARTNER_FTN.BP_CD"		
	    arrField(3) = "PARTNER.BP_RGST_NO"			
	    arrField(4) = "PARTNER_FTN.PARTNER_FTN"					    
	    
	    arrHeader(0) = "발행처"							
	    arrHeader(1) = "발행처명"						
	    arrHeader(2) = "거래처코드"						
	    arrHeader(3) = "사업자등록번호"						
	    arrHeader(4) = "거래처타입"						
		
		frm1.txtBill_to_party.focus 
	Case 2												
		If frm1.txtDoc_cur.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "화폐"						
		arrParam(1) = "B_CURRENCY"						
		arrParam(2) = Trim(frm1.txtDoc_cur.value)		
		arrParam(3) = ""								
		arrParam(4) = ""								
		arrParam(5) = "화폐"						
		
	    arrField(0) = "CURRENCY"						
	    arrField(1) = "CURRENCY_DESC"					
	    
	    arrHeader(0) = "화폐"						
	    arrHeader(1) = "화폐명"						
		
		frm1.txtDoc_cur.focus 
	Case 3												
		If frm1.txtDeal_Type.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True or lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "판매유형"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtDeal_Type.value)		
		arrParam(3) = Trim(frm1.txtDeal_Type_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & ""				
		arrParam(5) = "판매유형"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "판매유형"					
	    arrHeader(1) = "판매유형명"					
		
		frm1.txtDeal_Type.focus 
	Case 4												
		If frm1.txtTrans_Meth.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "운송방법"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtTrans_Meth.value)	
		arrParam(3) = Trim(frm1.txtTrans_Meth_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9009", "''", "S") & ""				
		arrParam(5) = "운송방법"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "운송방법"					
	    arrHeader(1) = "운송방법명"					
		
		frm1.txtTrans_Meth.focus 
	Case 5												
		If frm1.txtPay_terms.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True or lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "결제방법"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtPay_terms.value)		
		arrParam(3) = Trim(frm1.txtPay_terms_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & ""				
		arrParam(5) = "결제방법"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "결제방법"					
	    arrHeader(1) = "결제방법명"					
		
		frm1.txtPay_terms.focus 
	Case 6												
		If frm1.txtVat_Type.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True or lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "VAT유형"								
		arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"	
		arrParam(2) = Trim(frm1.txtVat_Type.value)				
		arrParam(3) = ""										
		arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
						& " And Config.MINOR_CD = Minor.MINOR_CD" _
						& " And Config.SEQ_NO = 1"				
		arrParam(5) = "VAT유형"								
		
	    arrField(0) = "Minor.MINOR_CD"							
	    arrField(1) = "Minor.MINOR_NM"							
	    arrField(2) = "Config.REFERENCE"						
	    	    
	    arrHeader(0) = "VAT유형"								
	    arrHeader(1) = "VAT유형명"							
		arrHeader(2) = "VAT율"							
		
		frm1.txtVat_Type.focus 
	Case 7												
		If frm1.txtSending_Bank.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "송금은행"						
		arrParam(1) = "B_BANK"								
		arrParam(2) = Trim(frm1.txtSending_Bank.value)		
		arrParam(3) = Trim(frm1.txtSending_Bank_nm.value)	
		arrParam(4) = ""									
		arrParam(5) = "송금은행"						
		
	    arrField(0) = "BANK_CD"								
	    arrField(1) = "BANK_FULL_NM"						
	    arrField(2) = "BANK_NM"							
	    
	    arrHeader(0) = "송금은행"						
	    arrHeader(1) = "송금은행전명"					
	    arrHeader(2) = "송금은행명"					
		
		frm1.txtSending_Bank.focus 
	Case 8												
		If frm1.txtPack_cond.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "포장조건"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtPack_cond.value)		
		arrParam(3) = Trim(frm1.txtPack_cond_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9007", "''", "S") & ""				
		arrParam(5) = "포장조건"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "포장조건"					
	    arrHeader(1) = "포장조건명"					
		
		frm1.txtPack_cond.focus 
	Case 9												
		If frm1.txtInspect_meth.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "검사방법"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtInspect_meth.value)	
		arrParam(3) = Trim(frm1.txtInspect_meth_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9008", "''", "S") & ""				
		arrParam(5) = "검사방법"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "검사방법"					
	    arrHeader(1) = "검사방법명"					
		
		frm1.txtInspect_meth.focus 
	Case 10												
		If frm1.txtPay_type.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True or lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		If Trim(frm1.txtPay_terms.value) = "" Then
			Call DisplayMsgBox("205152", "X", "결제방법", "X")
			'MsgBox "결제방법을 먼저 입력하세요!"
			frm1.txtPay_terms.focus
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "입금유형"					
		arrParam(1) = "B_MINOR,B_CONFIGURATION," _
		& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & ""_
			& "And MINOR_CD=" & FilterVar(frm1.txtPay_terms.value, "''", "S") & " And SEQ_NO>=2)C"
		arrParam(2) = Trim(frm1.txtPay_type.value)		
		arrParam(3) = Trim(frm1.txtPay_type_nm.value)	
		arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
					& "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("R", "''", "S") & " )"	
		arrParam(5) = "입금유형"					
		
	    arrField(0) = "B_MINOR.MINOR_CD"				
	    arrField(1) = "B_MINOR.MINOR_NM"				
	    
	    arrHeader(0) = "입금유형"					
	    arrHeader(1) = "입금유형명"					
		
		frm1.txtPay_type.focus 
	Case 11												
		If frm1.txtPayer.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If Trim(frm1.txtSold_to_party.value) = "" Then
			Call DisplayMsgBox("203150", "X", "X", "X")
			'MsgBox "주문처를 먼저 입력하세요!"
			frm1.txtSold_to_party.focus 
			IsOpenPop = False			
			Exit Function
		End IF

		arrParam(0) = "수금처"												
		arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
		arrParam(2) = Trim(frm1.txtPayer.value)									
		arrParam(3) = Trim(frm1.txtPayerNm.value)								
		arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SPA", "''", "S") & " " _
						& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE <= " & FilterVar("CS", "''", "S") & " " _
						& "AND PARTNER_FTN.BP_CD=" & FilterVar(frm1.txtSold_to_party.value, "''", "S")  	
		arrParam(5) = "수금처"							
		
	    arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"			
	    arrField(1) = "PARTNER.BP_NM"						
	    arrField(2) = "PARTNER_FTN.BP_CD"			
	    arrField(3) = "PARTNER.BP_RGST_NO"		
	    arrField(4) = "PARTNER_FTN.PARTNER_FTN"					    
	    
	    arrHeader(0) = "수금처"							
	    arrHeader(1) = "수금처명"						
	    arrHeader(2) = "거래처"							
	    arrHeader(3) = "사업자등록번호"						
	    arrHeader(4) = "거래처타입"						
		
		frm1.txtPayer.focus 
	Case 12
		If frm1.txtVat_Inc_Flag.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True or lsClickCfmNo = True Then 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "VAT포함구분"					
		arrParam(1) = "B_MINOR"								
		arrParam(2) = Trim(frm1.txtVat_Inc_Flag.value)		
		arrParam(3) = Trim(frm1.txtVat_Inc_Flag_Nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("S4035", "''", "S") & ""					
		arrParam(5) = "VAT포함구분"					
		
	    arrField(0) = "MINOR_CD"							
	    arrField(1) = "MINOR_NM"							
	    	    
	    arrHeader(0) = "VAT포함구분"					
	    arrHeader(1) = "VAT포함구분명"				
		
		frm1.txtVat_Inc_Flag.focus 
	End Select

	arrParam(3) = ""			'☜: [Condition Name Delete]

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOption(arrRet,iOption)
	End If	
End Function

'==========================================================================================================
Function OpenMinorCd(Byval iOpenMinor)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lsClickCfmYes = True or lsClickCfmNo = True Then 
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_MINOR"								
		
    arrField(0) = "MINOR_CD"							
    arrField(1) = "MINOR_NM"							
	    	    
	Select Case iOpenMinor

	Case 1
		arrParam(2) = Trim(frm1.txtDischge_port_Cd.value)	
		arrParam(3) = Trim(frm1.txtDischge_port_Nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9092", "''", "S") & ""					
		arrParam(5) = "도착항"							
		    	    
	    arrHeader(0) = "도착항"							
	    arrHeader(1) = "도착항명"						
		
		frm1.txtDischge_port_Cd.focus 
	
	Case 2
		arrParam(2) = Trim(frm1.txtLoading_port_Cd.value)	
		arrParam(3) = Trim(frm1.txtLoading_port_Nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9092", "''", "S") & ""					
		arrParam(5) = "선적항"							
		    	    
	    arrHeader(0) = "선적항"							
	    arrHeader(1) = "선적항명"						
		
		frm1.txtLoading_port_Cd.focus 
	
	Case 3
		arrParam(2) = Trim(frm1.txtOrigin.value)			
		arrParam(3) = Trim(frm1.txtOriginNm.value)			
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9094", "''", "S") & ""					
		arrParam(5) = "원산지"							
		    	    
	    arrHeader(0) = "원산지"							
	    arrHeader(1) = "원산지명"						
		
		frm1.txtOrigin.focus 
	End Select

	arrParam(0) = arrParam(5)								
	arrParam(3) = ""			

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCd(iOpenMinor, arrRet)
	End If
End Function

'==========================================================================================================
Function SetSORef(strRet)
	Call ggoOper.ClearField(Document, "1")								 '⊙: Clear Condition  Field 
	Call InitVariables													 '⊙: Initializes local global variables 
	Call SetDefaultVal

	frm1.txtHSONo.value = strRet

	Dim strVal

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode=" & "RETURNSOQUERY"				
    strVal = strVal & "&txtConSo_no=" & Trim(frm1.txtHSONo.value)	

	Call RunMyBizASP(MyBizASP, strVal)								

	lgBlnFlgChgValue = True
End Function

'==========================================================================================================
Function SetPrjRef(strRet)
	Call ggoOper.ClearField(Document, "1")								 '⊙: Clear Condition  Field 
	Call InitVariables													 '⊙: Initializes local global variables 
	Call SetDefaultVal
	
	frm1.txtProjectCd.value = strRet(0)
	
	Dim strVal

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode=" & "PROJECTQUERY"				
    strVal = strVal & "&txtProjectCd=" & Trim(frm1.txtProjectCd.value)	

	Call RunMyBizASP(MyBizASP, strVal)								

	lgBlnFlgChgValue = True
End Function

'==========================================================================================================
Function SetMinorCd(Byval iSetMinor, arrRet)
	Select Case iSetMinor
		
		Case 1
			frm1.txtDischge_port_Cd.Value = arrRet(0)
			frm1.txtDischge_port_Nm.Value = arrRet(1)
		
		Case 2
			frm1.txtLoading_port_Cd.Value = arrRet(0)
			frm1.txtLoading_port_Nm.Value = arrRet(1)
		
		Case 3
			frm1.txtOrigin.Value = arrRet(0)
			frm1.txtOriginNm.Value = arrRet(1)
	End Select

	lgBlnFlgChgValue = True
End Function

'==========================================================================================================
Function SetRequried(Byval arrRet,ByVal iRequried)

If arrRet(0) <> "" Then 

	Select Case iRequried
	Case 0												
		frm1.txtSo_Type.value = arrRet(0)
		frm1.txtSo_TypeNm.value = arrRet(1)
		frm1.txtSoTypeExportFlag.value = arrRet(2)
		frm1.txtSoTypeRetItemFlag.value = arrRet(3)
		frm1.txtSoTypeCiFlag.value = arrRet(4)
		Call BizSoTypeExpChange
	Case 1												
		frm1.txtSales_Grp.value = arrRet(0)
		frm1.txtSales_GrpNm.value = arrRet(1)
	Case 2											
		frm1.txtIncoTerms.value = arrRet(0)
		frm1.txtIncoTerms_nm.value = arrRet(1)
	Case 3 
		frm1.txtTo_Biz_Grp.value = arrRet(0)
		frm1.txtTo_Biz_GrpNm.value = arrRet(1)
	End Select

	lgBlnFlgChgValue = True

End If

End Function


'==========================================================================================================
Function SetBp(Byval arrRet,ByVal iRequried)

If arrRet(0) <> "" Then 

	Select Case iRequried
	Case 0												
		frm1.txtSold_to_party.value = arrRet(0)
		frm1.txtSold_to_partyNm.value = arrRet(1)
		Call SoldToPartyLookUp()
	Case 1												
		frm1.txtManufacturer.value = arrRet(0)
		frm1.txtManufacturer_nm.value = arrRet(1)
	Case 2												
		frm1.txtAgent.value = arrRet(0)
		frm1.txtAgent_nm.value = arrRet(1)
	Case 3												
		frm1.txtBeneficiary.value = arrRet(0)
		frm1.txtBeneficiary_nm.value = arrRet(1)
	End Select

	lgBlnFlgChgValue = True

End If

End Function

'==========================================================================================================
Function SetOption(Byval arrRet,ByVal iOption)

If arrRet(0) <> "" Then 
	Select Case iOption
	Case 0												
		frm1.txtShip_to_party.value = arrRet(0)
		frm1.txtShip_to_partyNm.value = arrRet(1)
	Case 1												
		frm1.txtBill_to_party.value = arrRet(0)
		frm1.txtBill_to_partyNm.value = arrRet(1)
	Case 2												
		frm1.txtDoc_cur.value = arrRet(0)
		Call txtDoc_cur_OnChange
	Case 3												
		frm1.txtDeal_Type.value = arrRet(0)
		frm1.txtDeal_Type_nm.value = arrRet(1)
	Case 4												
		frm1.txtTrans_Meth.value = arrRet(0)
		frm1.txtTrans_Meth_nm.value = arrRet(1)
	Case 5											
		If frm1.txtPay_terms.value <> arrRet(0) Then Call txtPay_terms_OnChange()
		frm1.txtPay_terms.value = arrRet(0)
		frm1.txtPay_terms_nm.value = arrRet(1)
	Case 6											
		frm1.txtVat_Type.value = arrRet(0)
		frm1.txtVatTypeNm.value = arrRet(1)
		'frm1.txtVat_rate.text = arrRet(2)
		frm1.txtVat_rate.text = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	Case 7												
		frm1.txtSending_Bank.value = arrRet(0)
		frm1.txtSending_Bank_nm.value = arrRet(2)
	Case 8												
		frm1.txtPack_cond.value = arrRet(0)
		frm1.txtPack_cond_nm.value = arrRet(1)
	Case 9											
		frm1.txtInspect_meth.value = arrRet(0)
		frm1.txtInspect_meth_nm.value = arrRet(1)
	Case 10
		frm1.txtPay_type.value = arrRet(0)
		frm1.txtPay_Type_nm.value = arrRet(1)
	Case 11
		frm1.txtPayer.value = arrRet(0)				
		frm1.txtPayerNm.value = arrRet(1)
	Case 12
		frm1.txtVat_Inc_Flag.value = arrRet(0)		
		frm1.txtVat_Inc_Flag_Nm.value = arrRet(1)

	End Select

	lgBlnFlgChgValue = True

End If

End Function


'==========================================================================================================
Sub SoldToPartyLookUp()

    Err.Clear                                                               
    
	If LayerShowHide(1) = False Then
		Exit Sub
	End If
	    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUp"								
    strVal = strVal & "&txtSold_to_party=" & Trim(frm1.txtSold_to_party.value)
    
	Call RunMyBizASP(MyBizASP, strVal)											
	
End Sub


'==========================================================================================================
Sub UnLockColor_CfmNo()	
	With frm1
		Call ggoOper.SetReqAttr(.txtTo_Biz_Grp, "D")
		Call ggoOper.SetReqAttr(.txtDoc_cur, "N")
		Call ggoOper.SetReqAttr(.txtVat_Inc_Flag, "D")
		Call ggoOper.SetReqAttr(.txtVat_Type, "D")
		Call ggoOper.SetReqAttr(.txtTrans_Meth, "D")
		Call ggoOper.SetReqAttr(.txtPay_terms, "N")
		Call ggoOper.SetReqAttr(.txtCust_po_no, "D")
		Call ggoOper.SetReqAttr(.txtCust_po_dt, "D")
		Call ggoOper.SetReqAttr(.txtReq_dlvy_dt, "D")
		Call ggoOper.SetReqAttr(.txtPay_dur, "D")
		Call ggoOper.SetReqAttr(.txt_Payterms_txt, "D")
		Call ggoOper.SetReqAttr(.txtRemark, "D")
		Call ggoOper.SetReqAttr(.txtBill_to_party, "D")
		Call ggoOper.SetReqAttr(.txtPayer, "D")
	End With
End Sub

'==========================================================================================================
Sub LockColor_CfmYes()

	With frm1
		Call ggoOper.SetReqAttr(frm1.txtPay_terms, "Q")
		Call ggoOper.SetReqAttr(frm1.txtTo_Biz_Grp, "Q")
		Call ggoOper.SetReqAttr(frm1.txtDoc_cur, "Q")
		Call ggoOper.SetReqAttr(frm1.txtVat_Inc_Flag, "Q")
		Call ggoOper.SetReqAttr(frm1.txtVat_Type, "Q")
		Call ggoOper.SetReqAttr(frm1.txtTrans_Meth, "Q")
		Call ggoOper.SetReqAttr(frm1.txtCust_po_dt, "Q")
		Call ggoOper.SetReqAttr(frm1.txtReq_dlvy_dt, "Q")
		Call ggoOper.SetReqAttr(frm1.txtPay_dur, "Q")
		Call ggoOper.SetReqAttr(frm1.txt_Payterms_txt, "Q")
		Call ggoOper.SetReqAttr(frm1.txtRemark, "Q")
		Call ggoOper.SetReqAttr(frm1.txtXchg_rate, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBill_to_party, "Q")
		Call ggoOper.SetReqAttr(frm1.txtPayer, "Q")

	End With

End Sub

'==========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877				
	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtConSo_no.value & parent.gRowSep & frm1.txtRadioType.value

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtConSo_no.value =  arrVal(0)
		
'		frm1.txtConSo_no.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If
		
		Call MainQuery()
					
		WriteCookie CookieSplit , ""
		
	End If

End Function

'==========================================================================================================
Function LocValidDateCheck()

	LocValidDateCheck = False

	With frm1
		
		If .txtReq_dlvy_dt.text <> "" Then
			If UniConvDateToYYYYMMDD(.txtSo_dt.Text,parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtReq_dlvy_dt.Text,parent.gDateFormat,"") Then
				Call DisplayMsgBox("970022","X",.txtReq_dlvy_dt.Alt,.txtSo_dt.Alt)			
				Call ClickTab1()
				.txtSo_dt.focus
				Exit Function
				
			End If
		End If
		
		If .txtCust_po_dt.text <> "" and .txtReq_dlvy_dt.text <> "" Then
			If UniConvDateToYYYYMMDD(.txtCust_po_dt.Text,parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtReq_dlvy_dt.Text,parent.gDateFormat,"") Then
				Call DisplayMsgBox("970022","X",.txtReq_dlvy_dt.Alt,.txtCust_po_dt.Alt)
				Call ClickTab1()
				.txtCust_po_dt.focus
				Exit Function
				
			End If
		End If
		
	End With

	LocValidDateCheck = True

End Function

'==========================================================================================================
Function JumpChgCheck()

	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(1)
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function


'==========================================================================================================
Function SoTypeExportCheck()

	SoTypeExportCheck = False

	With frm1
		

		If UCase(Trim(.txtSoTypeExportFlag.value)) = "Y" Or UCase(Trim(.txtSoTypeCiFlag.value)) = "Y" Then		
			
		ElseIf UCase(Trim(.txtSoTypeExportFlag.value)) = "" Or UCase(Trim(.txtSoTypeCiFlag.value)) = "" Then
		
			MsgBox "수주형태를 먼저 선택하세요", vbInformation, parent.gLogoName
			.txtSo_Type.focus
			Exit Function
		Else
		
			MsgBox "수주형태:" & Space(1) & Trim(.txtSo_TypeNm.value) & "의 수출/통관여부는" & Space(1) & UCase(Trim(.txtSoTypeExportFlag.value)) & "/" & UCase(Trim(.txtSoTypeCiFlag.value)) & Space(1) & "입니다" & vbCrlf & vbCrlf _
				& "수주형태가 수출 또는 통관인 경우만 무역정보 탭으로 이동할 수 있습니다", vbInformation, parent.gLogoName
			.txtSo_Type.focus
			Exit Function	
					
		End If		

	End With

	SoTypeExportCheck = True

End Function


'==========================================================================================================
Function SoTypeExpRequiredChg()
	With frm1
		Call ggoOper.SetReqAttr(frm1.txtBeneficiary, "N")

		Call ggoOper.SetReqAttr(frm1.txtContract_dt, "N")

		Call ggoOper.SetReqAttr(frm1.txtValid_dt, "N")

		Call ggoOper.SetReqAttr(frm1.txtIncoTerms, "N")
	End With
End Function

'==========================================================================================================
Function SoTypeExpDefaultChg()
	With frm1
		
		Call ggoOper.SetReqAttr(frm1.txtBeneficiary, "D")
	
		Call ggoOper.SetReqAttr(frm1.txtContract_dt, "D")
		
		Call ggoOper.SetReqAttr(frm1.txtValid_dt, "D")
		
		Call ggoOper.SetReqAttr(frm1.txtIncoTerms, "D")
	End With
End Function


'==========================================================================================================
Function BizSoTypeExpChange()

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Dim strVal

	strVal = ""

	strVal = BIZ_PGM_ID & "?txtMode=" & "SoTypeExp"							'☜: 비지니스 처리 ASP의 상태 %>
	strVal = strVal & "&txtSo_Type=" & Trim(frm1.txtSo_Type.value)			'☆: 조회 조건 데이타 %>

	Call RunMyBizASP(MyBizASP, strVal)										

End Function

'==========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    
    If GetSetupMod(parent.gSetupMod,"y") = "Y" Then
		txtOpenPrjRef.style.display = ""
    End If
    
	Call AppendNumberPlace("6","3","0")
	Call AppendNumberPlace("7","3","2")
	
    Call FormatField()
    Call ggoOper.LockField(Document, "N")
	Call LockFieldInit("L")    
    
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    Call SetToolbar("11101000000011")										'⊙: 버튼 툴바 제어 
	Call CookiePage(0)
	Call ChangeTabs(TAB1)

	Call ggoOper.SetReqAttr(frm1.txtXchg_rate, "Q")

	gIsTab     = "Y"	:	gTabMaxCnt = 2
	
End Sub

'=========================================
Sub FormatField()
    With frm1
        ' 날짜 OCX Foramt 설정 
		Call FormatDATEField(.txtSo_dt)	
		Call FormatDATEField(.txtReq_dlvy_dt)		'납기일 
		Call FormatDATEField(.txtCust_po_dt)
		Call FormatDATEField(.txtContract_dt)
		Call FormatDATEField(.txtValid_dt)
		Call FormatDATEField(.txtship_dt)
		' 숫자 OCX Foramt 설정 
		Call FormatDoubleSingleField(.txtPay_dur)
		Call FormatDoubleSingleField(.txtNet_amt)
		Call FormatDoubleSingleField(.txtXchg_rate)
		Call FormatDoubleSingleField(.txtVat_rate)
		Call FormatDoubleSingleField(.txtVat_amt)
		Call FormatDoubleSingleField(.txtNet_Amt_Loc)
    End With
End Sub

'===========================================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' 날짜 OCX
        Call LockObjectField(.txtSo_dt,"R")    
		Call LockObjectField(.txtReq_dlvy_dt,"O")    
		Call LockObjectField(.txtCust_po_dt,"O")
		Call LockObjectField(.txtContract_dt,"O")
		Call LockObjectField(.txtValid_dt,"O")
		Call LockObjectField(.txtship_dt,"O")   
		Call LockObjectField(.txtPay_dur,"O")
		Call LockObjectField(.txtNet_amt,"P")
		Call LockObjectField(.txtXchg_rate,"O")
		Call LockObjectField(.txtVat_rate,"P")
		Call LockObjectField(.txtVat_amt,"P")
		Call LockObjectField(.txtNet_Amt_Loc,"P")       
		
		If pvFlag = "N" Then
			Call LockHTMLField(.txtSoNo, "O")	
        End If

    End With
End Sub

'==========================================================================================================
Function fncSoTypeExpChange()

	With frm1
		'20030122 inkuk
		If UCase(Trim(frm1.txtRetItemFlag.value)) <> UCase(Trim(frm1.txtSoTypeRetItemFlag.value)) AND checkSoDtlExist = False Then
			Call DisplayMsgBox("203244", "X", "X", "X")
		End If		
		
		If frm1.rdoCfm_flag1.checked = True Then
		
			Call ggoOper.SetReqAttr(frm1.txtBeneficiary, "Q")
			
			Call ggoOper.SetReqAttr(frm1.txtContract_dt, "Q")
		
			Call ggoOper.SetReqAttr(frm1.txtValid_dt, "Q")
			
			Call ggoOper.SetReqAttr(frm1.txtIncoTerms, "Q")
			Exit Function
		End If
        
		If UCase(Trim(frm1.txtSoTypeExportFlag.value)) = "Y" Or UCase(Trim(.txtSoTypeCiFlag.value)) = "Y" Then
			Call SoTypeExpRequiredChg()
		Else
			Call SoTypeExpDefaultChg()
		End If		
		
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			If Len(.txtHDlvyLt.value) Then
				.txtReq_dlvy_dt.text = UNIDateAdd("d", UNICDbl(.txtHDlvyLt.value), .txtSo_dt.text, parent.gDateFormat)
			End If
		End If

	End With

End Function

'==========================================================================================================
Function checkSoDtlExist()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim iCodeArr

	Err.Clear

	Call CommonQueryRs(" count(*) ", " S_SO_DTL ", " SO_NO =  " & FilterVar(frm1.txtSoNo.value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	iCodeArr = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Function
	End If

	IF iCodeArr(0) = 0 Then
		checkSoDtlExist = True
	Else
		checkSoDtlExist = False
	End If
End Function


'==========================================================================================================
Function RunAutoDN()

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Dim strVal

	strVal = ""

	strVal = BIZ_PGM_ID & "?txtMode=" & DNCheck								
	strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)			
	strVal = strVal & "&txtInsrtUserId=" & Trim(parent.gUsrID)
	strVal = strVal & "&RdoDnReq=" & Trim(frm1.RdoDnReq.value)
	
	Call RunMyBizASP(MyBizASP, strVal)										

End Function

'==========================================================================================================
Sub CurFormatNumericOCX()
	With frm1

		'개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtNet_amt, .txtDoc_cur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXchg_rate, .txtDoc_cur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		
	End With
End Sub


'==========================================================================================================
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub


'==========================================================================================================
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


'==========================================================================================================
Sub SetVatType()
	Dim VatType, VatTypeNm, VatRate

	VatType = frm1.txtVat_Type.value
	
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)
	
	frm1.txtVatTypeNm.value = VatTypeNm
	frm1.txtVat_rate.text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)	
End Sub



'==========================================================================================================
Function ObjectChange()
    On Error Resume Next
    
    Dim i, strTagName
	Dim strScript
            
	strScript = ""

'	strScript = strScript & "</" & "Script" & ">" & vbCrLf
	strScript = strScript & "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf   

    For i = 0 To Document.All.Length - 1
        strTagName = ""
        
        strTagName = UCase(Document.All(i).tagName)
        
        If Err.Number <> 0 Then
            Err.Clear
        Else
            Select Case strTagName
                Case "OBJECT"
                    If Document.All(i).Title = "FPDATETIME" Or Document.All(i).Title = "FPDOUBLESINGLE" Then
						strScript = strScript & "Sub " & Document.All(i).Name & "_Change()" & vbCrLf
						strScript = strScript & "lgBlnFlgChgValue = True" & vbCrLf	
						strScript = strScript & "End Sub" & vbCrLf
                    End If
            End Select
        End If
    Next

	strScript = strScript & "</" & "Script" & ">" & vbCrLf
'	strScript = strScript & "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf    

	msgbox strScript

End Function

'==========================================================================================================
Sub btnDNCheck_OnClick()

	Dim Answer	
	If lgBlnFlgChgValue = True Then Answer = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
	If Answer = VBNO Then Exit Sub
	
	If lgBlnFlgChgValue = False Then Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")	
	If Answer = VBNO Then Exit Sub	
	
	Call RunAutoDN

End Sub


'==========================================================================================================
Sub btnConfirm_OnClick()

	Dim Answer
	
	If lgBlnFlgChgValue = True Then Answer = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
	If Answer = VBNO Then Exit Sub

	If lgBlnFlgChgValue = False Then Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")	
	If Answer = VBNO Then Exit Sub

	If LayerShowHide(1) = False Then
		Exit Sub
	End If

	If Trim(frm1.RdoConfirm.value) = "Y" Then						'확정처리시 여신한도 체크 
		Call CheckCreditlimitSvr
	Else
	    Call ConfirmSO()	
	End If

End Sub

'==========================================================================================================
Function CheckCreditlimitSvr()

	Dim iStrVal
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	    
	iStrVal = BIZ_PGM_ID & "?txtMode=" & "CheckCreditlimit"      
	iStrVal = iStrVal & "&txtCaller=" & "SC"								'확정처리시 
	iStrVal = iStrVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)
	iStrVal = iStrVal & "&txtTotalAmt=" & "0"
	
	Call RunMyBizASP(MyBizASP, iStrVal)	

End Function

'==========================================================================================================
Sub ConfirmSO()
	
	Dim iStrVal
	
	iStrVal = ""
	iStrVal = BIZ_PGM_ID & "?txtMode=" & "btnCONFIRM"					
	iStrVal = iStrVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)	
	iStrVal = iStrVal & "&txtInsrtUserId=" & Trim(parent.gUsrID)	
	iStrVal = iStrVal & "&RdoConfirm=" & Trim(frm1.RdoConfirm.value)	

	Call RunMyBizASP(MyBizASP, iStrVal)    

End Sub

Sub rdoDnparcel_flag1_OnClick()
	frm1.txtRadioDnParcel.value = frm1.rdoDnparcel_flag1.value 
	Call RadioChange(0)
End Sub

Sub rdoDnparcel_flag2_OnClick()
	frm1.txtRadioDnParcel.value = frm1.rdoDnparcel_flag2.value 
	Call RadioChange(0)
End Sub

Sub rdoCfm_flag1_OnClick()
	frm1.txtRadioFlag.value = frm1.rdoCfm_flag1.value
	Call RadioChange(1)
End Sub

Sub rdoCfm_flag2_OnClick()
	frm1.txtRadioFlag.value = frm1.rdoCfm_flag2.value
	Call RadioChange(1)
End Sub

Function rdoPrice_flag1_OnClick()
	frm1.txtRadioType.value = frm1.rdoPrice_flag1.value 
	Call RadioChange(2) 
End Function

Function rdoPrice_flag2_OnClick()
	frm1.txtRadioType.value = frm1.rdoPrice_flag2.value
	Call RadioChange(2) 
End Function

Function RadioChange(ByVal Rval)

	Select Case Rval
	Case 0

	Case 1
		If frm1.txtRadioFlag.value <> PrevRadioFlag Then
			lgBlnFlgChgValue = True
		Else
			lgBlnFlgChgValue = False
		End IF
	Case 2
		If frm1.txtRadioType.value <> PrevRadioType Then
			lgBlnFlgChgValue = True
		Else
			lgBlnFlgChgValue = False		
		End IF
	End Select

End Function

'==========================================================================================================

Sub txtSo_dt_Change()
	lgBlnFlgChgValue = True
	frm1.txtReq_dlvy_dt.text = UNIDateAdd("d", UNICDbl(frm1.txtHDlvyLt.value), frm1.txtSo_dt.text, parent.gDateFormat)
End Sub

Sub txtCust_po_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtReq_dlvy_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPay_dur_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtVat_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtNet_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtXchg_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtVat_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtNet_Amt_Loc_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtContract_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtValid_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtship_dt_Change()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================================
Sub txtSo_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtSo_dt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtSo_dt.Focus
	End If
End Sub
Sub txtCust_po_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtCust_po_dt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtCust_po_dt.Focus
	End If
End Sub
Sub txtReq_dlvy_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReq_dlvy_dt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtReq_dlvy_dt.Focus
	End If
End Sub
Sub txtContract_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtContract_dt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtContract_dt.Focus
	End If
End Sub
Sub txtValid_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtValid_dt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtValid_dt.Focus
	End If
End Sub
Sub txtship_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtship_dt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtship_dt.Focus
	End If
End Sub

'==========================================================================================================
Sub txtSold_to_party_OnChange()
	If gLookUpEnable = False Then Exit Sub
	If Trim(frm1.txtSold_to_party.value) <> "" Then Call SoldToPartyLookUp()
End Sub

'==========================================================================================================
Sub txtPay_terms_OnChange()

	'---입금유형 
	frm1.txtPay_type.value = ""
	frm1.txtPay_Type_nm.value = ""

	'---결제기간 
	frm1.txtPay_dur.text = 0

End Sub

'==========================================================================================================
Sub txtDoc_cur_OnChange()	
	Call CurrencyOnChange
End Sub

'==========================================================================================================
Function CurrencyOnChange()

	If UCase(Trim(frm1.txtDoc_cur.value)) = UCase(parent.gCurrency) Then
		frm1.txtXchg_rate.Text = 0		
		Call ggoOper.SetReqAttr(frm1.txtXchg_rate, "Q")
	Else
		frm1.txtXchg_rate.Text = 0
		Call ggoOper.SetReqAttr(frm1.txtXchg_rate, "N")		
	End If
	Call CurFormatNumericOCX()

End Function


'==========================================================================================================
Sub txtSo_Type_OnChange()
	If gLookUpEnable = False Then Exit Sub
	frm1.txtSoTypeExportFlag.value = ""
	frm1.txtSoTypeRetItemFlag.value = ""
	frm1.txtSoTypeCiFlag.value = ""
	If Len(Trim(frm1.txtSo_Type.value)) > 0 Then BizSoTypeExpChange

End Sub

'==========================================================================================================
Sub txtVat_Type_OnChange()
	Call SetVatType()
End Sub

'==========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	If Not chkFieldByCell(frm1.txtConSo_no,"A",gPageNo) Then Exit Function

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
									
    Call InitVariables															
    Call SetDefaultVal   
	Call LockFieldInit("L") 
	Call UnLockColor_CfmNo()
    Call DbQuery															
       
    FncQuery = True																
        
End Function


'==========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False  

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If    

    Call ggoOper.ClearField(Document, "A")                                          
    Call LockFieldInit("N")                              
    Call SetToolbar("11101000000011")
    Call SetDefaultVal

	Call UnLockColor_CfmNo()

	Call ggoOper.SetReqAttr(frm1.txtSoNo, "D")			'옵션 
	Call ggoOper.SetReqAttr(frm1.txtSo_Type, "N")		'필수 
	Call ggoOper.SetReqAttr(frm1.txtSales_Grp, "N")
	Call ggoOper.SetReqAttr(frm1.txtSold_to_party, "N")
	Call ggoOper.SetReqAttr(frm1.txtShip_to_party, "N")
	Call ggoOper.SetReqAttr(frm1.txtXchg_rate, "Q")		
	Call ggoOper.SetReqAttr(frm1.txtDeal_Type, "N")

	frm1.txtSo_dt.text = EndDate
	frm1.txtCust_po_dt.text = EndDate
	frm1.txtReq_dlvy_dt.text = EndDate

	frm1.txtPay_dur.text = 0
	frm1.txtVat_rate.text = 0
	frm1.txtNet_amt.text = 0
	frm1.txtXchg_rate.text = 0
	frm1.txtVat_amt.text = 0
	frm1.txtVat_rate.text = 0
	frm1.txtNet_Amt_Loc.text = 0

    Call InitVariables														

    FncNew = True															

End Function

'==========================================================================================================
Function FncDelete() 
    
    Dim IntRetCD
    
    FncDelete = False												
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
   
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
    Call DbDelete															
    
    FncDelete = True                                                        
    
End Function

'==========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                     
    
    Err.Clear    

	If LocValidDateCheck = False Then	
		Exit Function							
	End if
	
    If lgBlnFlgChgValue = False Then		
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
	'---------------------------------------------------------------------
	' 필수입력체크 및 입력값 길이 체크 
    '-------------------------------------------------------------------
    If Not chkFieldByCell(frm1.txtSo_Type, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtSold_to_party, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtSales_Grp, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtShip_to_party, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtDeal_Type, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtPay_terms, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtDoc_cur, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtIncoTerms, "A", "2") Then Exit Function
    If Not chkFieldByCell(frm1.txtBeneficiary, "A", "2") Then Exit Function
    
    If Not ChkFieldLengthByCell(frm1.txtCust_po_no, "A", "1") Then Exit Function        
    If Not ChkFieldLengthByCell(frm1.txt_Payterms_txt, "A", "1") Then Exit Function
    If Not ChkFieldLengthByCell(frm1.txtRemark, "A", "1") Then Exit Function
    If Not ChkFieldLengthByCell(frm1.txtShip_dt_txt, "A", "2") Then Exit Function        
    If Not ChkFieldLengthByCell(frm1.txtDischge_city, "A", "2") Then Exit Function
    
	If UCase(Trim(frm1.txtRetItemFlag.value)) <> UCase(Trim(frm1.txtSoTypeRetItemFlag.value)) AND checkSoDtlExist = False Then
		Call DisplayMsgBox("203244", "X", "X", "X")
		Exit Function
	End If	
	
    Call DbSave				                                               
    
    FncSave = True                                                         
    
End Function

'==========================================================================================================
Function FncCopy() 

	Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE								
        
    Call ggoOper.ClearField(Document, "1")                          
    Call LockFieldInit("N")						
	Call UnLockColor_CfmNo()
    Call InitVariables												
    Call SetToolbar("11101000000011")

	frm1.txtNet_amt.Text = ""		
	frm1.txtVat_amt.Text = ""		
	frm1.txtNet_Amt_Loc.Text = ""
		
	frm1.rdoCfm_flag2.checked = True
	frm1.txtRadioFlag.value = frm1.rdoCfm_flag2.value 
	
	If frm1.rdoCfm_flag1.checked = True Then
		Call rdoCfm_flag1_OnClick()
	ElseIf frm1.rdoCfm_flag2.checked = True Then
		Call rdoCfm_flag2_OnClick()
		frm1.btnConfirm.value = "확정처리"
	End If
	
	If frm1.rdoPrice_flag1.checked = True Then
		Call rdoPrice_flag1_OnClick()
	ElseIf frm1.rdoPrice_flag2.checked = True Then
		Call rdoPrice_flag2_OnClick()
	End If	

	frm1.btnDNCheck.disabled = True
	frm1.txtSoNo.value = ""
	lgBlnFlgChgValue = True
	
	Call CurrencyOnChange
	
	If UCase(Trim(frm1.txtSoTypeExportFlag.value)) = "Y" Or UCase(Trim(.txtSoTypeCiFlag.value)) = "Y" Then
		Call SoTypeExpRequiredChg()
	Else
		Call SoTypeExpDefaultChg()
	End If
	
End Function

'==========================================================================================================
Function FncCancel() 
    On Error Resume Next                                                  
End Function


'==========================================================================================================
Function FncInsertRow() 
     On Error Resume Next                                                   
End Function


'==========================================================================================================
Function FncDeleteRow() 
    On Error Resume Next                                                    
End Function


'==========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function


'==========================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011", "X", "X", "X")  '☜ 바뀐부분 
		'Call MsgBox("이전 데이타가 없습니다..", vbInformation)
		Exit Function
    End If

End Function


'==========================================================================================================
Function FncNext() 
    On Error Resume Next                                                    

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")  '☜ 바뀐부분 
		'Call MsgBox("이전 데이타가 없습니다..", vbInformation)
		Exit Function
    End If

End Function


'==========================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLE)
End Function


'==========================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, True)
End Function


'==========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function


'==========================================================================================================
Function DbDelete() 
    Err.Clear                                                               
    
    DbDelete = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							
    strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)		
    
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbDelete = True                                                         

End Function

'==========================================================================================================
Function DbDeleteOk()														
	Call MainNew()
End Function

'==========================================================================================================
Function DbQuery() 
    
    Err.Clear                                                               
    
    DbQuery = False                                                         

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtConSo_no=" & Trim(frm1.txtConSo_no.value)		
    
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbQuery = True                                                          

End Function

'==========================================================================================================
Function DbQueryOk()														
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE											
	
	Call ggoOper.SetReqAttr(frm1.txtSoNo, "Q")
	
	Call ggoOper.SetReqAttr(frm1.txtSo_Type, "N")
	
	Call ggoOper.SetReqAttr(frm1.txtSales_Grp, "N")
	
	If UNICDbl(frm1.txtNet_amt.text) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtSold_to_party, "Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtSold_to_party, "N")
	End If
	
	Call ggoOper.SetReqAttr(frm1.txtShip_to_party, "N")

	If frm1.rdoCfm_flag1.checked = True Then
		Call ggoOper.LockField(Document, "Q")
		Call LockColor_CfmYes()
		lsClickCfmYes = True
	ElseIf frm1.rdoCfm_flag2.checked = True Then
	End If

	If frm1.rdoCfm_flag1.checked = True Then
		Call SetToolbar("11100000001111")
	Else
		If ChkSoByInf Then
			Call LockField("INF", True)
		Else
			Call LockField("INF", False)
		End If
		Call SetToolbar("11111000001111")
	End If

	PrevRadioFlag = frm1.txtRadioFlag.value
	PrevRadioType = frm1.txtRadioType.value
	
	Call SetbtnConfirmButton()
	
	lgBlnFlgChgValue = False

	Call BizSoTypeExpChange()

End Function

'==========================================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Dim strVal

	With frm1
		.txtMode.value = parent.UID_M0002											
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = parent.gUsrID 
		.txtUpdtUserId.value = parent.gUsrID

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           
    
End Function

'==========================================================================================================
Function DbSaveOk()												

    Call InitVariables
    Call MainQuery()

End Function

'==========================================================================================================
Function SetbtnConfirmButton()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	iStrSelectList	= " COUNT(*) "
	iStrFromList	= " S_SO_DTL "
	iStrWhereList	= " SO_NO =  " & FilterVar(frm1.txtSoNo.value, "''", "S") & ""

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, parent.gColSep)
		If CLng(iArrRs(1)) > 0 Then
			frm1.btnConfirm.disabled = False
		Else
			frm1.btnConfirm.disabled = True
		End If
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If	
End Function

'==========================================================================================================
Function ChkSoByInf()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	ChkSoByInf = False
	
	If Trim(frm1.txtCust_po_no.value) = "" Then Exit Function
	
	iStrSelectList = " COUNT(*) "
	iStrFromList  = " S_SO_DTL "
	iStrWhereList = " INF_NO IS NOT NULL AND SO_NO =  " & FilterVar(frm1.txtSONo.value, "''", "S") & ""

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, parent.gColSep)
		If CLng(iArrRs(1)) > 0 Then
			ChkSoByInf = True
		End If
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If

End Function

' 필드 속성 변경 
Sub LockField(ByVal pvStrCaller, ByVal pvBlnLock)
	If pvStrCaller = "INF" Then
		With frm1
			If pvBlnLock Then
				Call ggoOper.SetReqAttr(.txtCust_po_no ,"Q")
				Call ggoOper.SetReqAttr(.txtCust_po_dt,"Q")
			Else
				Call ggoOper.SetReqAttr(.txtCust_po_no ,"D")
				Call ggoOper.SetReqAttr(.txtCust_po_dt,"D")
			End If
		End With
	End If
End Sub
