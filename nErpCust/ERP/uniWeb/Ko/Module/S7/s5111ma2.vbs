  
' External ASP File
'========================================
Const BIZ_PGM_ID = "s5111mb2.asp"            
Const BIZ_BillDtl_JUMP_ID = "s5112ma2"
Const BIZ_BillCollect_JUMP_ID = "s5114ma1"

' Constant variables 
'========================================
Const PostFlag = "PostFlag"

Const TAB1 = 1                  '☜: Tab의 위치 
Const TAB2 = 2

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

		.txtBillDt.Text = EndDate
		.txtSalesGrpCd.value = Parent.gSalesGrp
		.txtToBizAreaCd.value = Parent.gSalesGrp
		.btnPostFlag.disabled = True
		.btnGLView.disabled = True
		.btnPostFlag.value = "확정"
		.btnPreRcptView.disabled = True
		.rdoVatCalcType1.checked = True
		.chkTaxNo.checked = False
		.txtLocCur.value = Parent.gCurrency
		.btnBillTaxNo.disabled = True 

		.rdoVATCalcType1.checked = True
		.rdoVATIncFlag1.checked = True

		Call ggoOper.SetReqAttr(.txtVatType, "D")

		lgBlnFlgChgValue = False
		Call CurrencyOnChange()
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
Function OpenConBillNo()
	On Error Resume Next
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
		   
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s5111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111pa1", "X")
		lblnWinEvent = False
		Exit Function
	End If
		 
	strRet = window.showModalDialog(iCalledAspName& "?txtExceptFlag=Y", Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtConBillNo.focus
	If strRet <> "" Then frm1.txtConBillNo.value = strRet  
	
End Function
   
'==========================================
Function OpenBillRef()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
  
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 

	iCalledAspName = AskPRAspName("s5111ra2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111ra2", "X")
		Exit Function
	End If
	    
	strRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If strRet(0) <> "" Then	Call SetBillRef(strRet)
End Function 

'==========================================
Function OpenBillHdr(ByVal iBillHdr)

	On Error Resume Next

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	OpenBillHdr = False
			 
	Select Case iBillHdr
		Case 1  ' 수금처 
			If frm1.txtPayerCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			If Trim(frm1.txtSoldToPartyCd.value) = "" Then
				Call DisplayMsgBox("203150","X","X","X")
				'MsgBox "주문처를 먼저 입력하세요!"
				frm1.txtSoldToPartyCd.focus 
				IsOpenPop = False   
				Exit Function
			End IF
			        
			          
			arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"  
			arrParam(2) = Trim(frm1.txtPayerCd.value)        
			arrParam(3) = ""        
			arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SPA", "''", "S") & " " _
			& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " ) " _ 
			& "AND PARTNER_FTN.BP_CD =  " & FilterVar(frm1.txtSoldtoPartyCd.value, "''", "S") & " "
			      
			arrParam(5) = "수금처"      
			  
			arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"  
			arrField(1) = "PARTNER.BP_NM" 
			arrField(2) = "PARTNER.BP_RGST_NO"
    		            
			arrHeader(0) = "수금처"      
			arrHeader(1) = "수금처명"     
			arrHeader(2) = "사업자등록번호"     

			frm1.txtPayerCd.focus

		Case 2  ' 수금영업그룹 
			If frm1.txtToBizAreaCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(1) = "B_SALES_GRP"      
			arrParam(2) = Trim(frm1.txtToBizAreaCd.value) 
			arrParam(3) = ""       
			arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "     
			arrParam(5) = "수금영업그룹"    
			     
			arrField(0) = "SALES_GRP"      
			arrField(1) = "SALES_GRP_NM"     
			        
			arrHeader(0) = "영업그룹"     
			arrHeader(1) = "영업그룹명"     

			frm1.txtToBizAreaCd.focus
			
		Case 3  ' 신고사업장 
			If frm1.txtTaxBizAreaCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(1) = "B_TAX_BIZ_AREA"      
			arrParam(2) = Trim(frm1.txtTaxBizAreaCd.value) 
			arrParam(3) = ""        
			arrParam(4) = ""        
			arrParam(5) = "세금신고사업장"    
			     
			arrField(0) = "TAX_BIZ_AREA_CD"      
			arrField(1) = "TAX_BIZ_AREA_NM"      
			        
			arrHeader(0) = "세금신고사업장"    
			arrHeader(1) = "세금신고사업장명"   

			frm1.txtTaxBizAreaCd.focus
			
		Case 4  ' 입금유형 
			If frm1.txtPayTypeCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			If Trim(frm1.txtPayTermsCd.value) = "" Then
				Call DisplayMsgBox("205152", "X", "결제방법", "X")
				'MsgBox "결제방법을 먼저 입력하세요!"
				frm1.txtPayTermsCd.focus
				IsOpenPop = False   
				Exit Function
			End IF

			arrParam(0) = "입금유형"												' 팝업 명칭 
			arrParam(1) = "B_MINOR,B_CONFIGURATION," _
							& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & ""_
							& "And MINOR_CD= " & FilterVar(frm1.txtPayTermsCd.value, "''", "S") & " And SEQ_NO>=2)C"
			arrParam(2) = Trim(frm1.txtPayTypeCd.value)
			arrParam(3) = ""
			arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
							& "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("R", "''", "S") & " )"
			arrParam(5) = "입금유형"
			  
			arrField(0) = "B_MINOR.MINOR_CD"
			arrField(1) = "B_MINOR.MINOR_NM"
			     
			arrHeader(0) = "입금유형"
			arrHeader(1) = "입금유형명"

			frm1.txtPayTypeCd.focus
			
		Case 5  ' 결제방법 
			If frm1.txtPayTermsCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(1) = "B_MINOR MINOR, B_CONFIGURATION CON"      
			arrParam(2) = Trim(frm1.txtPayTermsCd.value) 
			arrParam(3) = ""        
			arrParam(4) = "CON.MINOR_CD = MINOR.MINOR_CD" _
							& " AND CON.MAJOR_CD = MINOR.MAJOR_CD AND CON.MAJOR_CD = " & FilterVar("B9004", "''", "S") & "" _
							& " AND CON.REFERENCE = " & FilterVar("N", "''", "S") & " " _
							& " AND CON.SEQ_NO = 1"      
			arrParam(5) = "결제방법"     
			  
			arrField(0) = "CON.MINOR_CD"     
			arrField(1) = "MINOR.MINOR_NM"     
			     
			arrHeader(0) = "결제방법"     
			arrHeader(1) = "결제방법명"
			
			frm1.txtPayTermsCd.focus     

		Case 6  ' 발행처 
			If frm1.txtBillToPartyCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If
			  
			If Trim(frm1.txtSoldToPartyCd.value) = "" Then
				Call DisplayMsgBox("203150","X","X","X")
				'MsgBox "주문처를 먼저 입력하세요!"
				frm1.txtSoldToPartyCd.focus 
				IsOpenPop = False   
				Exit Function
			End IF
			  
			arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"  
			arrParam(2) = Trim(frm1.txtBillToPartyCd.value)       
			arrParam(3) = ""           
			arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SBI", "''", "S") & " " _
							& "AND PARTNER.BP_CD=PARTNER_FTN.PARTNER_BP_CD AND PARTNER.BP_TYPE in (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " ) " _
							& "AND PARTNER_FTN.BP_CD =  " & FilterVar(frm1.txtSoldtoPartyCd.value, "''", "S") & " " ' Where Condition

			arrParam(5) = "발행처"            
			   
			arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"        
			arrField(1) = "PARTNER.BP_NM"           
			arrField(2) = "PARTNER.BP_RGST_NO"

			arrHeader(0) = "발행처"            
			arrHeader(1) = "발행처명"           
			arrHeader(2) = "사업자등록번호"     
			 
			frm1.txtBillToPartyCd.focus

		Case 7  ' VAT
			If frm1.txtVatType.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config" 
			arrParam(2) = Trim(frm1.txtVatType.value)    
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
			
			frm1.txtVatType.focus

		Case 8  ' 영업그룹 
			If frm1.txtSalesGrpCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(1) = "B_SALES_GRP"      
			arrParam(2) = Trim(frm1.txtSalesGrpCd.value)  
			arrParam(3) = ""       
			arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "     
			arrParam(5) = "영업그룹"     
			  
			arrField(0) = "SALES_GRP"      
			arrField(1) = "SALES_GRP_NM"     
			     
			arrHeader(0) = "영업그룹"     
			arrHeader(1) = "영업그룹명"   
			
			frm1.txtSalesGrpCd.focus  

		Case 9  ' 매출채권형태 
			If frm1.txtBillTypeCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(1) = "S_BILL_TYPE_CONFIG"    
			arrParam(2) = Trim(frm1.txtBillTypeCd.value) 
			arrParam(3) = ""        
			arrParam(5) = "매출채권형태"    
				 
			arrField(0) = "BILL_TYPE"      
			arrField(1) = "BILL_TYPE_NM"     

			arrHeader(0) = "매출채권형태"     
			arrHeader(1) = "매출채권형태명"
			
			If Trim(frm1.txtHExportFlag.value) = "" Then
				arrParam(4) = "EXCEPT_FLAG = " & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND AS_FLAG = " & FilterVar("N", "''", "S") & " "  
			Else
				arrParam(4) = "EXCEPT_FLAG = " & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND AS_FLAG = " & FilterVar("N", "''", "S") & "  AND EXPORT_FLAG = " _
								& " " & FilterVar(frm1.txtHExportFlag.value, "''", "S") & " " 
			End If
			
			frm1.txtBillTypeCd.focus

		Case 10  ' 화폐구분 
			If frm1.txtDocCur1.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If

			arrParam(1) = "B_CURRENCY"      
			arrParam(2) = Trim(frm1.txtDocCur1.value)  
			arrParam(3) = ""        
			arrParam(4) = ""        
			arrParam(5) = "화폐"      
			 
			arrField(0) = "CURRENCY"      
			arrField(1) = "CURRENCY_DESC"     
			    
			arrHeader(0) = "화폐"      
			arrHeader(1) = "화폐명"      

			frm1.txtDocCur1.focus
			
		Case 11  ' 주문처 
			If frm1.txtSoldtoPartyCd.readOnly = True Then
				IsOpenPop = False
				Exit Function
			End If
			           
			arrParam(1) = "B_BIZ_PARTNER"            
			arrParam(2) = Trim(frm1.txtSoldtoPartyCd.value)        
			arrParam(3) = ""               
			arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "      
			 
			arrParam(5) = "주문처"      
			   
			arrField(0) = "BP_CD"       
			arrField(1) = "BP_NM"       
			arrField(2) = "BP_RGST_NO"
			      
			arrHeader(0) = "주문처"      
			arrHeader(1) = "주문처명"
			arrHeader(2) = "사업자등록번호"     
			
			frm1.txtSoldtoPartyCd.focus
	  
	End Select
	 
	arrParam(0) = arrParam(5)       ' 팝업 명칭 

	If Err.number <> 0 Then
		MsgBox err.Description, vbInformation,Parent.gLogoName
		IsOpenPop = False
		Exit Function
	End If
	    
	Select Case iBillHdr
	Case 1, 6, 11
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select 

	IsOpenPop = False

	If arrRet(0) <> "" Then	OpenBillHdr = SetBillHdr(arrRet,iBillHdr)
End Function

'===========================================================================
Function OpenTaxNo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "세금계산서번호"    ' 팝업 명칭 
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
	arrField(1) = "DD15" & Parent.gColSep & "CONVERT(char(11),expiry_date)" 
	arrField(2) = "convert(char(15),TAX_BOOK_NO)"
	arrField(3) = "convert(char(15),TAX_BOOK_SEQ)"
         
	arrHeader(0) = "세금계산서번호"   
	arrHeader(1) = "유효일"     
	arrHeader(2) = "책번호(권)"
	arrHeader(3) = "책번호(호)"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
									 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtTaxBillNo.focus
	
	If arrRet(0) <> "" Then
		frm1.txtTaxBillNo.value = arrRet(0)
		lgBlnFlgChgValue = True
	End If 
 
End Function

'====================================================
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
		If strRs <> "" Then
			arrTemp = Split(strRs, Chr(11))
			arrTaxBizArea(0) = arrTemp(1)
			arrTaxBizArea(1) = arrTemp(2)
			
			Call SetBillHdr(arrTaxBizArea, 3)
		End If
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

'====================================================
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

'====================================================
Function SetBillRef(strRet)
	Call ggoOper.ClearField(Document, "1")         
	Call InitVariables              

	frm1.txtRefBillNo.value = Trim(strRet(0))
	frm1.txtHExceptFlg.value = strRet(1)

	frm1.chkRefBillNoFlg.checked = True
	 
	Dim strVal
	   
	If LayerShowHide(1) = False Then Exit Function 

	' strRet(2) - 수출여부(Y:수출) 
	If strRet(2) = "Y" Then
		strVal = BIZ_PGM_ID & "?txtMode=" & "BLQuery"       
		strVal = strVal & "&txtBillNo=" & Trim(frm1.txtRefBillNo.value)
		strVal = strVal & "&txtExceptFlg=" & Trim(frm1.txtHExceptFlg.value)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & "BillQuery"       
		strVal = strVal & "&txtBillNo=" & Trim(frm1.txtRefBillNo.value)
		strVal = strVal & "&txtExceptFlg=" & Trim(frm1.txtHExceptFlg.value)
	End If
	    
	Call RunMyBizASP(MyBizASP, strVal)          

	frm1.txtBillNo.focus
	
	lgBlnFlgChgValue = True

End Function

'====================================================
Function SetBillHdr(Byval arrRet,ByVal iBillHdr)

	SetBillHdr = False
 
	Select Case iBillHdr
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
		Case 6            ' 발행처 
			frm1.txtBillToPartyCd.value = arrRet(0)
			frm1.txtBillToPartyNm.value = arrRet(1)
			'발행처 관련된 세금신고사업장 Fetch
			Call GetTaxBizArea("BP")
		Case 7            ' VAT
			frm1.txtVatType.value = arrRet(0)
			frm1.txtVatTypeNm.value = arrRet(1)
			frm1.txtVatRate.Text = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		Case 8            ' 영업그룹 
			frm1.txtSalesGrpCd.value = arrRet(0)
			frm1.txtSalesGrpNm.value = arrRet(1)
				  
			If Trim(frm1.txtToBizAreaCd.value) = "" Then
				frm1.txtToBizAreaCd.value = arrRet(0)
				frm1.txtToBizAreaNm.value = arrRet(1)
			End if
				  
			'해당 영업그룹 관련 세금신고사업장 Fetch 
			Call GetTaxBizArea("BA")
		Case 9            ' 매출채권형태 
			frm1.txtBillTypeCd.value = arrRet(0)
			frm1.txtBillTypeNm.value = arrRet(1)
		Case 10            ' 화폐 
			frm1.txtDocCur1.value = arrRet(0)
			Call CurrencyOnChange
		Case 11            ' 주문처 
			frm1.txtSoldtoPartyCd.value = arrRet(0)
			frm1.txtSoldtoPartyNm.value = arrRet(1)
			Call SoldToPartyLookUp()
	End Select

	SetBillHdr = True
	lgBlnFlgChgValue = True

End Function

'========================================
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

'========================================
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

'========================================
 Sub ProtectBillRefTag()
	With frm1
		Call ggoOper.SetReqAttr(.txtSoldtoPartyCd, "Q")
		Call ggoOper.SetReqAttr(.txtBillToPartyCd, "Q")
		Call ggoOper.SetReqAttr(.txtDocCur1, "Q")
	End With
 End Sub 

'========================================
 Sub ProtectVattypeTag()
	With frm1
		ggoOper.SetReqAttr frm1.txtVatType, "D"
		ggoOper.SetReqAttr frm1.rdoVatIncflag1, "Q"
		ggoOper.SetReqAttr frm1.rdoVatIncflag2, "Q"
	End With
 End Sub 

'========================================
 Sub ReleaseBillRefTag()
	With frm1
		Call ggoOper.SetReqAttr(.txtSoldtoPartyCd, "N")
		Call ggoOper.SetReqAttr(.txtBillToPartyCd, "N")
		Call ggoOper.SetReqAttr(.txtSalesGrpCd, "N")
		Call ggoOper.SetReqAttr(.txtToBizAreaCd, "N")
	End With   
 End Sub

'===========================================================================
Sub CalcPlanIncomeDt()
	Err.Clear
	If Trim(frm1.txtBillDt.Text) = "" Then Exit Sub

	If UNICDbl(frm1.txtCreditRotDay.value) = 0 Then
		frm1.txtPlanIncomeDt.Text = ""
	Else
		frm1.txtPlanIncomeDt.Text = UNIDateAdd("d", frm1.txtCreditRotDay.value, Trim(frm1.txtBillDt.Text), Parent.gDateFormat)
	End If
End Sub 

'===========================================================================
Sub SoldToPartyLookUp()

	Err.Clear
	If LayerShowHide(1) = False Then Exit Sub 

	Dim strVal
	    
	strVal = BIZ_PGM_ID & "?txtMode=" & "BillLookUp"       
	strVal = strVal & "&txtSoldtoPartyCd=" & Trim(frm1.txtSoldtoPartyCd.value)
	strVal = strVal & "&txtBillDt=" & Trim(frm1.txtBillDt.Text)
	    
	Call RunMyBizASP(MyBizASP, strVal) 
	 
	Call txtVatType_OnChange()
End Sub

'========================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next
	Const CookieSplit = 4877
	Dim strTemp, arrVal

	If Kubun = 1 Then
		WriteCookie CookieSplit , frm1.txtHBillNo.value & Parent.gRowSep & frm1.txtHRefFlag.value 
	ElseIf Kubun = 0 Then
		strTemp = ReadCookie(CookieSplit)
		
		If strTemp = "" then Exit Function
		arrVal = Split(strTemp, Parent.gRowSep)
		
		If arrVal(0) = "" Then Exit Function
		frm1.txtConBillNo.value =  arrVal(0)
		
		WriteCookie CookieSplit , ""
		Call DbQuery()
	End If

End Function

'===========================================================================
Function JumpChgCheck(strJump)

	Dim IntRetCD

	'************ 싱글인 경우 **************
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(1)

	Select Case strJump
		Case BIZ_BillDtl_JUMP_ID 
			Call PgmJump(BIZ_BillDtl_JUMP_ID)
		Case BIZ_BillCollect_JUMP_ID
			Call PgmJump(BIZ_BillCollect_JUMP_ID)
	End Select

End Function

'====================================================
Sub CurFormatNumericOCX()

 With frm1
  '매출채권금액 
  ggoOper.FormatFieldByObjectOfCur .txtBillAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  'VAT금액 
  ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  '적립금액 
  ggoOper.FormatFieldByObjectOfCur .txtDepositAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  '매출채권총금액 
  ggoOper.FormatFieldByObjectOfCur .txtTotBillAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  '총수금액 
  ggoOper.FormatFieldByObjectOfCur .txtIncomeAmt, .txtDocCur1.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
  '환율 
  ggoOper.FormatFieldByObjectOfCur .txtXchgRate, .txtDocCur1.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec

 End With

End Sub

'====================================================
Function BtnSpreadCheck()
 BtnSpreadCheck = False

 Dim Answer
 If lgBlnFlgChgValue = True Then Answer = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
 If Answer = VBNO Then Exit Function

 If lgBlnFlgChgValue = False Then Answer = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")
 If Answer = VBNO Then Exit Function

 BtnSpreadCheck = True

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
End Sub

Sub LockFieldAll()
    Call LockObjectField(frm1.txtBillDt,"P")
    Call LockObjectField(frm1.txtPlanIncomeDt,"P")
    Call LockObjectField(frm1.txtXchgRate,"P")
    Call LockObjectField(frm1.txtPayDur,"P")

	Call LockHTMLField(frm1.chkRefBillNoFlg, "P")
	Call LockHTMLField(frm1.chkTaxNo, "P")

	Call LockHTMLField(frm1.txtBillNo, "P")	
	Call LockHTMLField(frm1.rdoVATCalcType1, "P")	
	Call LockHTMLField(frm1.rdoVATCalcType2, "P")	
	Call LockHTMLField(frm1.txtVatType, "P")	
	Call LockHTMLField(frm1.rdoVatIncflag1, "P")	
	Call LockHTMLField(frm1.rdoVatIncflag2, "P")	

	Call LockHTMLField(frm1.txtBillTypeCd, "P")	
	Call LockHTMLField(frm1.txtBillToPartyCd, "P")	
	Call LockHTMLField(frm1.txtSoldtoPartyCd, "P")	
	Call LockHTMLField(frm1.txtPayerCd, "P")	
	Call LockHTMLField(frm1.txtPayTermsCd, "P")
	Call LockHTMLField(frm1.txtSalesGrpCd, "P")
	Call LockHTMLField(frm1.txtDocCur1, "P")

	Call LockHTMLField(frm1.txtToBizAreaCd , "P")	
	Call LockHTMLField(frm1.txtTaxBizAreaCd , "P")	

	Call LockHTMLField(frm1.txtPayTypeCd, "P")	
	Call LockHTMLField(frm1.txtVatType, "P")	
	Call LockHTMLField(frm1.txtPaytermsTxt, "P")	
	Call LockHTMLField(frm1.txtRemark, "P")	
End Sub

Sub LockFieldQuery()
	If UNICDbl(frm1.txtSts.value) < 3 Then
		Call LockHTMLField(frm1.txtDocCur1, "P")	
		Call LockHTMLField(frm1.txtBillToPartyCd, "P")	
		Call LockHTMLField(frm1.txtSoldtoPartyCd, "P")	
	Else
		If frm1.txtHRefFlag.value = "B" Then
			Call LockHTMLField(frm1.txtDocCur1, "P")	
			Call LockHTMLField(frm1.txtBillToPartyCd, "P")	
			Call LockHTMLField(frm1.txtSoldtoPartyCd, "P")	
		Else
			Call LockHTMLField(frm1.txtDocCur1, "R")	
			Call LockHTMLField(frm1.txtBillToPartyCd, "R")	
			Call LockHTMLField(frm1.txtSoldtoPartyCd, "R")	
		End If
	End If

	Call LockHTMLField(frm1.txtSalesGrpCd, "R")	
	Call LockHTMLField(frm1.txtToBizAreaCd, "R")	
	Call LockHTMLField(frm1.txtBillTypeCd, "P")	
	Call LockHTMLField(frm1.txtPayerCd, "P")	
	Call LockHTMLField(frm1.txtTaxBizAreaCd, "P")	
	Call LockHTMLField(frm1.txtPayTermsCd, "R")	
	Call LockHTMLField(frm1.txtBillNo, "P")	
	Call LockHTMLField(frm1.chkRefBillNoFlg, "P")
		
	If UCase(Trim(frm1.txtDocCur1.value)) = UCase(Parent.gCurrency) Then
	    Call LockObjectField(frm1.txtXchgRate,"P")
	Else
	    Call LockObjectField(frm1.txtXchgRate,"R")
	End If
	
	If UNICDbl(frm1.txtSts.value) < 3 Then
		Call LockHTMLField(frm1.rdoVatIncflag1, "P")	
		Call LockHTMLField(frm1.rdoVatIncflag2, "P")	
		Call LockHTMLField(frm1.rdoVATCalcType1, "P")	
		Call LockHTMLField(frm1.rdoVATCalcType2, "P")	
	Else		
		Call LockHTMLField(frm1.rdoVatIncflag1, "O")	
		Call LockHTMLField(frm1.rdoVatIncflag2, "O")	
		Call LockHTMLField(frm1.rdoVATCalcType1, "O")	
		Call LockHTMLField(frm1.rdoVATCalcType2, "O")	
	End If

	If frm1.rdoVATCalcType1.checked Then
		Call LockHTMLField(frm1.txtVatType, "O")	
	Else
		Call LockHTMLField(frm1.txtVatType, "R")	
	End If

	Call LockHTMLField(frm1.txtPayerCd, "R")	
	Call LockHTMLField(frm1.txtTaxBizAreaCd, "R")	

    Call LockObjectField(frm1.txtBillDt,"R")
    Call LockObjectField(frm1.txtPlanIncomeDt,"O")

	Call LockHTMLField(frm1.txtPayTypeCd , "O")	
	Call LockHTMLField(frm1.txtPaytermsTxt  , "O")	
	Call LockHTMLField(frm1.txtRemark  , "O")	
End Sub

Sub LockFieldNew()


	Call LockHTMLField(frm1.chkRefBillNoFlg, "O")
	Call LockHTMLField(frm1.chkTaxNo, "O")
	Call LockHTMLField(frm1.txtBillNo, "O")	

	Call LockHTMLField(frm1.txtBillTypeCd, "R")	
	Call LockHTMLField(frm1.txtSoldtoPartyCd, "R")	
	Call LockHTMLField(frm1.txtBillToPartyCd, "R")	
	Call LockHTMLField(frm1.txtPayerCd, "R")	
	Call LockHTMLField(frm1.txtDocCur1, "R")
	Call LockHTMLField(frm1.txtPayTermsCd, "R")	
	Call LockHTMLField(frm1.txtToBizAreaCd , "R")	
	Call LockHTMLField(frm1.txtTaxBizAreaCd , "R")	
	Call LockHTMLField(frm1.txtPayTypeCd, "O")	
	Call LockHTMLField(frm1.txtPaytermsTxt, "O")	
	Call LockHTMLField(frm1.txtRemark, "O")	
	Call LockHTMLField(frm1.txtVatType, "O")	
	Call LockHTMLField(frm1.rdoVatIncflag1, "O")	
	Call LockHTMLField(frm1.rdoVatIncflag2, "O")	
	Call LockHTMLField(frm1.rdoVATCalcType1, "O")	
	Call LockHTMLField(frm1.rdoVATCalcType2, "O")	
End Sub

Function CheckField()
	CheckField = False

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		If Not chkFieldByCell(frm1.txtBillTypeCd, "A", "1") Then Exit Function
		If Not chkFieldByCell(frm1.txtBillToPartyCd, "A", "1") Then Exit Function
		If Not chkFieldByCell(frm1.txtSoldtoPartyCd, "A", "1") Then Exit Function
		If Not chkFieldByCell(frm1.txtSalesGrpCd, "A", "1") Then Exit Function
		If Not chkFieldByCell(frm1.txtVatType, "A", "1") Then Exit Function
		If Not chkFieldByCell(frm1.txtPayTermsCd, "A", "1") Then Exit Function
	    If Not chkFieldByCell(frm1.txtPayerCd, "A", "1") Then Exit Function
	    If Not chkFieldByCell(frm1.txtDocCur1, "A", "1") Then Exit Function
	    If Not chkFieldByCell(frm1.txtTaxBizAreaCd, "A", "1") Then Exit Function
 	End If

    If Not chkFieldByCell(frm1.txtBillDt , "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtToBizAreaCd, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtXchgRate, "A", "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtTaxBizAreaCd , "A", "1") Then Exit Function

    If Not ChkFieldLengthByCell(frm1.txtPaytermsTxt, "A", "1") Then Exit Function
    If Not ChkFieldLengthByCell(frm1.txtRemark, "A", "1") Then Exit Function

	CheckField = True
End Function

'==========================================
Sub Form_Load()
	Call LoadInfTB19029                                                     
	Call AppendNumberPlace("6","3","0")
	Call LockFieldInit()
'	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
'	Call ggoOper.LockField(Document, "N")

	Call SetTaxBillNoMgmtMeth
	Call SetDefaultVal

	Call SetToolbar("11101000000011")
	Call InitVariables
	Call CookiePage(0)

'	Call ChangeTabs(TAB1)

	gIsTab     = "Y" : gTabMaxCnt = 2

	TabDiv(1).style.display = "none"

End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================
Sub btnPostFlag_OnClick()

	if frm1.chkTaxNo.checked and Not gBlnTaxbillnoMgmtMeth Then
		Call DisplayMsgBox("205626", "x", "x", "x")
		'세금계산서 방법이 설정되지 않았거나 2개이상 설정되었습니다.
		Exit Sub
	End if       

	If BtnSpreadCheck = False Then Exit Sub

	Dim strVal

	frm1.txtInsrtUserId.value = Parent.gUsrID 
	   
	If LayerShowHide(1) = False Then Exit Sub

	strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag         
	strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtBillNo.value)     
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
Sub rdoVatCalcType1_OnClick()
     lgBlnFlgChgValue = True

     Call ggoOper.SetReqAttr(window.document.frm1.txtVatType, "D")

     '이전매출채권을 참조한 경우에는 부가세 계산방법이 개별인 경우 부가세유형 및 부가세 포함여부를 수정할 수 없다.
     If frm1.txtHRefFlag.value = "B" Then
          ' 부가세 계산방법이 개별인 경우에는 부가세 포함여부 사용자가 수정 불가 
          Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag1, "Q")
          Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncFlag2, "Q")
          frm1.rdoVatIncFlag1.checked = True
     End if
End Sub

'========================================
Sub rdoVatCalcType2_OnClick()
     lgBlnFlgChgValue = True
     Call ggoOper.SetReqAttr(window.document.frm1.txtVatType, "N")

     ' 부가세 계산방법이 통합인 경우에는 부가세 포함여부 사용자가 변경 가능 
     Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag1, "N")
     Call ggoOper.SetReqAttr(window.document.frm1.rdoVatIncflag2, "N")
End Sub

'========================================
Sub rdoVatIncFlag1_OnClick()
     '부가세 포함여부 
     lgBlnFlgChgValue = True 
End Sub

'========================================
Sub rdoVatIncFlag2_OnClick()
 '부가세포함여부 
 lgBlnFlgChgValue = True
End Sub

'========================================
Sub txtBillDt_Change()
 If Trim(frm1.txtBillDt.Text) <> "" Then Call CalcPlanIncomeDt()
 lgBlnFlgChgValue = True
End Sub

'수금만기일 
'==========================================
Sub txtPlanIncomeDt_Change()
 lgBlnFlgChgValue = True
End Sub

' 2004.05.18 SMJ B/L 여부 추가 
Sub rdoBlFlagY_OnClick()
      'B/L 여부 

     lgBlnFlgChgValue = True 
     frm1.txtHExportFlag.value = "Y"
End Sub

'========================================
Sub rdoBlFlagN_OnClick()
   'B/L 여부 
   lgBlnFlgChgValue = True
   frm1.txtHExportFlag.value = "N"
End Sub

'환율 
'==========================================
Sub txtXchgRate_Change()

 lgBlnFlgChgValue = True
End Sub

'결제기간 
'==========================================
Sub txtPayDur_Change()
 lgBlnFlgChgValue = True
End Sub

'==========================================
Sub txtBillDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillDt.Focus
	End If
End Sub

'==========================================
Sub txtPlanIncomeDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPlanIncomeDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPlanIncomeDt.Focus
	End If
End Sub

'==========================================
Sub txtSoldToPartyCd_OnChange()
 If Trim(frm1.txtSoldToPartyCd.value) <> "" Then Call SoldToPartyLookUp()
End Sub

'==========================================
Sub txtBillToPartyCd_OnChange()
 If Trim(frm1.txtBillToPartyCd.value) = "" Then
  'frm1.txtBillToPartyNm.value = ""
 Else
  ' 발행처 변경시 관련 세금신고 사업장 Fetch
  Call GetTaxBizArea("BP")
 End if
End Sub

'==========================================
Sub txtDocCur1_OnChange()
	Call CurrencyOnChange
	
End Sub

'==========================================
Function CurrencyOnChange()

 If Trim(frm1.txtDocCur1.value) = "" Then Exit Function
  Call CurFormatNumericOCX

 If UCase(Trim(frm1.txtDocCur1.value)) = UCase(Parent.gCurrency) Then
  frm1.txtXchgRate.Text = 1
  Call ggoOper.SetReqAttr(frm1.txtXchgRate, "Q")
 Else
  frm1.txtXchgRate.Text = 0 
  Call ggoOper.SetReqAttr(frm1.txtXchgRate, "N")
 End If

End Function

'==========================================
Sub txtPayTermsCd_OnChange()

 '---입금유형 
 frm1.txtPayTypeCd.value = ""
 frm1.txtPayTypeNm.value = ""

 '---결제기간 
 frm1.txtPayDur.text = 0

End Sub

'==========================================
Function txtVatType_OnChange()
	Dim strCode
	 
	strCode = Trim(frm1.txtVatType.value)
	If strCode <> "" Then
		strCode = " " & FilterVar(strCode, "''", "S") & ""
		If Not GetCodeName("" & FilterVar("B9001", "''", "S") & "", strCode, "default", "default", 1, "" & FilterVar("CF", "''", "S") & "", 7) Then
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
End Function

'==========================================
Function txtTaxBizAreaCd_OnChange()
	If Trim(frm1.txtTaxBizAreaCd.value) = "" Then
		frm1.txtTaxBizAreaNm.value = ""
	Else
		Call GetTaxBizArea("NM")
		txtTaxBizAreaCd_OnChange = False
		If frm1.txtTaxBizAreaCd.value <> "" Then frm1.txtPaytermsTxt.focus
	End if
End Function

'==========================================
Sub txtSalesGrpCd_OnChange()
 If Trim(frm1.txtSalesGrpCd.value) = "" Then
  frm1.txtSalesGrpNm.value = ""
 Else
  If Trim(frm1.txtToBizAreaCd.value) = "" Then
   frm1.txtToBizAreaCd.value = frm1.txtSalesGrpCd.value
   frm1.txtToBizAreaNm.value = frm1.txtSalesGrpNm.value
  End if
  '영업그룹과 관련된 세금신고사업장을 Fetch한다.
  Call GetTaxBizArea("BA")
 End if
End Sub

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

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If
    
   If Not chkFieldByCell(frm1.txtConBillNo, "A", 1) Then Exit Function 

'    If Not chkField(Document, "1") Then         
'       Exit Function
'    End If
    
'   Call ggoOper.ClearField(Document, "2")          
'   Call ggoOper.LockField(Document, "N")          
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

	With frm1
		'총수금액 
		.txtIncomeAmt.Text = 0
		'환율 
		.txtXchgRate.Text = 1
		'매출채권금액 
		.txtBillAmt.Text = 0
		'매출채권자국금액 
		.txtBillAmtLoc.Text = 0
		'VAT금액 
		.txtVatAmt.Text = 0
		'VAT자국금액 
		.txtVatLocAmt.Text = 0
		'총수금자국금액 
		.txtIncomeLocAmt.Text = 0
	End With

    Call SetToolbar("11101000000011")

	Call LockFieldNew
	Call LockFieldInit
	
	frm1.rdoVatCalcType1.checked = True 	

'    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call SetDefaultVal

    FncNew = True                

End Function

'========================================
Function FncDelete() 
    
    FncDelete = False              
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
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

    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

	If ValidDateCheck(frm1.txtBillDt, frm1.txtPlanIncomeDt) = False Then Exit Function

	If CheckField = False Then Exit Function

'    If Not chkField(Document, "2") Then                             
'       Exit Function
'    End If

    Call DbSave                                                    
    
    FncSave = True                                                          
    
End Function

'========================================
Function FncCopy() 
 Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If
	    
	lgIntFlgMode = Parent.OPMD_CMODE            
	    
	Call ggoOper.ClearField(Document, "1")
	Call LockFieldNew
	Call LockFieldInit

'	Call PostFlagRelease()
'	Call ggoOper.LockField(Document, "N")         
	Call InitVariables               
	Call SetToolbar("11101000000011")
	    
	frm1.txtBillNo.value = ""
	frm1.txtBillAmt.Text = 0
	frm1.txtBillAmtLoc.Text = 0

	Call CurrencyOnChange()

	frm1.btnPostFlag.disabled = True
	lgBlnFlgChgValue = True

End Function

'========================================
Function FncCancel() 
    On Error Resume Next                                                    
End Function

'========================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================
Function FncPrev() 
    Dim strVal
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")  
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    End If

 If   LayerShowHide(1) = False Then
         Exit Function 
    End If
 
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
 Call parent.FncExport(Parent.C_SINGLE)
End Function

'========================================
Function FncFind() 
 Call parent.FncFind(Parent.C_SINGLE, True)
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

   
  If   LayerShowHide(1) = False Then
             Exit Function 
        End If

    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003       
    strVal = strVal & "&txtBillNo=" & Trim(frm1.txtBillNo.value)   
    
 Call RunMyBizASP(MyBizASP, strVal)
 
    DbDelete = True                                                         

End Function

'========================================
Function DbDeleteOk()              
 'pis
 lgBlnFlgChgValue = False
 Call MainNew()
End Function

'========================================
Function DbQuery() 
    
    Err.Clear                                                               

    DbQuery = False                                                         
    
   
  If   LayerShowHide(1) = False Then
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
 
    lgIntFlgMode = Parent.OPMD_UMODE            
    
	Call SetToolbar("11111000110111")          
 
	'수금만기일 2999-12-31이면 안보이게 
	if UniConvDateToYYYYMMDD(frm1.txtPlanIncomeDt.Text,Parent.gDateFormat,"-") = "2999-12-31" then
		frm1.txtPlanIncomeDt.Text = ""    
	end if
 
	If UNICDbl(frm1.txtSts.value) < 3 Then
		frm1.btnPostFlag.disabled = False
	Else
		frm1.btnPostFlag.disabled = True
	End If

	' 전표가 발행되지 않은 경우 
	If frm1.rdoPostFlagN.checked Then
		Call LockFieldQuery
		Call chkTaxBillNoCreatedMeth 
	Else
		Call LockFieldAll
	End If
	
	If frm1.txtHExportFlag.value = "Y" Then		
		frm1.rdoBlFlagY.checked = True
	Else
		frm1.rdoBlFlagN.checked = True
	End If
	

	lgBlnFlgChgValue = False
End Function

'========================================
Function BillQueryOk()              
	IF Trim(frm1.txtVatCalcType.value) = "1" Then
		frm1.rdoVatCalcType1.checked = True
		Call ProtectVatTypeTag
	Else
		frm1.rdoVatCalcType2.checked = True
	End If
	 
	If Trim(frm1.txtVatIncFlag.value) = "1" Then
		frm1.rdoVatIncFlag1.checked = True
	Else
		frm1.rdoVatIncFlag2.checked = True
	End if

	'세금신고 사업장 Fetch
	Call GetTaxBizArea("*")

	' 이전매출 참조시 주문처, 발행처, 화폐단위 Protect 처리 
	Call ProtectBillRefTag

	lgBlnFlgChgValue =True
End Function

'========================================
Function DbSave() 
    Err.Clear                

	DbSave = False               

	If LayerShowHide(1) = False Then Exit Function 

    Dim strVal

	With frm1
		.txtMode.value = Parent.UID_M0002           
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = Parent.gUsrID 
		.txtUpdtUserId.value = Parent.gUsrID

		If .chkTaxNo.checked = True Then
			.txtchkTaxNo.value = "Y"
		Else
			.txtchkTaxNo.value = "N"
		End If
			 
		If .rdoVatCalcType1.checked = True Then
			.txtVatCalcType.value = .rdoVatCalcType1.value
		Else
			.txtVatCalcType.value = .rdoVatCalcType2.value
		End If

		If .rdoVatIncFlag1.checked = True Then
			.txtVatIncFlag.value = .rdoVatIncFlag1.value
		Else
			.txtVatIncFlag.value = .rdoVatIncFlag2.value
		End If

		If .chkRefBillNoFlg.checked Then
			frm1.txtRefBillNoFlg.value = "Y"
		Else
			frm1.txtRefBillNoFlg.value = "N"
		End If

		If .rdoBlFlagY.checked Then
			.txtHExportFlag.value = "Y"
		Else
			.txtHExportFlag.value = "N"
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

