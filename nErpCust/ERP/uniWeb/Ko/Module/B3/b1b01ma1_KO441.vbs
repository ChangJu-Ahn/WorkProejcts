Const BIZ_PGM_QRY_ID = "b1b01mb1.asp"
Const BIZ_PGM_SAVE_ID = "b1b01mb2.asp"
Const BIZ_PGM_DEL_ID = "b1b01mb3.asp"
Const BIZ_PGM_LOOKUPHSCD_ID = "b1b01mb5.asp"
Const BIZ_PGM_JUMPITEMIMG_ID = "b1b02ma1"
Const BIZ_PGM_JUMPITEMBYPLANT_ID = "b1b11ma1"	

Dim IsOpenPop
Dim lgRdoOldVal1
Dim lgRdoOldVal2
Dim lgRdoOldVal3					

Dim arrCollectVatType

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 

    IsOpenPop = False		
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	With frm1	
		.cboItemAcct.value = "10"
		.rdoPhantomType2.checked = True
		lgRdoOldVal1 = 2
		.rdoUnifyPurFlg2.checked = True
		lgRdoOldVal2 = 2
		.rdoValidFlg1.checked = True
		lgRdoOldVal3 = 1
		
		.rdoPhoto2.checked = True 
	
		.txtValidFromDt.Text  = StartDate
		.txtValidToDt.Text	= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	End With
End Sub

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtitemcd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(1) = ""							' Item Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
		
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & "  "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtUnit.focus
	
End Function

'------------------------------------------  OpenItemGroup()  --------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  AND LEAF_FLG = " & FilterVar("Y", "''", "S") & "  AND VALID_TO_DT >=  " & FilterVar(BaseDate , "''", "S") & "" 			
	arrParam(5) = "품목그룹"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"		
    
    arrHeader(0) = "품목그룹"		
    arrHeader(1) = "품목그룹명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function


'------------------------------------------  OpenBasisItemCd()  ------------------------------------------
'	Name : OpenBasisItemCd()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBasisItemCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtBasisItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtBasisItemCd.value)	' Plant Code
	arrParam(1) = ""								' Item Code
	arrParam(2) = ""								' ----------
	arrParam(3) = ""								' ----------
	arrParam(4) = BaseDate

    arrField(0) = 1 								' Field명(0) : "ITEM_CD"
    arrField(1) = 2 								' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBasisItemCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtBasisItemCd.focus
		
End Function

'------------------------------------------  OpenWeightUnit()  -------------------------------------------
'	Name : OpenWeightUnit()
'	Description : WeightUnit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWeightUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtWeightUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtWeightUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & " "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWeightUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWeightUnit.focus
	
End Function

'------------------------------------------  OpenGrossWeightUnit()  -------------------------------------------
'	Name : OpenWeightUnit()
'	Description : WeightUnit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGrossWeightUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGrossWeightUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtGrossWeightUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & " "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetGrossWeightUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtGrossWeightUnit.focus
	
End Function

'------------------------------------------  OpenHsCd()  -------------------------------------------------
'	Name : OpenHsCd()
'	Description : HS Cd PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenHsCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	If IsOpenPop = True Or UCase(frm1.txtHSCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HS팝업"	
	arrParam(1) = "B_HS_CODE"				
	arrParam(2) = Trim(frm1.txtHSCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "HS코드"
	
    arrField(0) = "HS_CD"	
    arrField(1) = "HS_NM"
    arrField(2) = "HS_SPEC"	
    arrField(3) = "HS_UNIT"
    	
    
    arrHeader(0) = "HS코드"		
    arrHeader(1) = "HS명"
    arrHeader(2) = "HS규격"
    arrHeader(3) = "HS단위"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetHSCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtHSCd.focus
	
End Function

'===========================================================================
' Function Name : OpenBillHdr
' Function Desc : OpenBillHdr Reference Popup
'===========================================================================
Function OpenBillHdr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	
	If frm1.txtVatType.readOnly = True Then
		IsOpenPop = False
		Exit Function
	End If

	arrParam(1) = "B_MINOR ,B_CONFIGURATION "	' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtVatType.value)	' Code Condition
	arrParam(3) = ""							' Name Condition
	arrParam(4) = "B_MINOR.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " " _
					& " AND B_MINOR.MINOR_CD=B_CONFIGURATION.MINOR_CD " _
					& " AND B_MINOR.MAJOR_CD=B_CONFIGURATION.MAJOR_CD "	_
					& " AND B_CONFIGURATION.SEQ_NO = 1 "					' Where Condition
	arrParam(5) = "VAT유형"					' TextBox 명칭 
		
	arrField(0) = "B_MINOR.MINOR_CD"			' Field명(0)
	arrField(1) = "B_MINOR.MINOR_NM"			' Field명(1)
	arrField(2) = "F5" & parent.gColSep & "B_CONFIGURATION.REFERENCE"				' Field명(2)
	    	    
	arrHeader(0) = "VAT유형"				' Header명(0)
	arrHeader(1) = "VAT유형명"				' Header명(1)
	arrHeader(2) = "VAT율"					' Header명(2)

	arrParam(0) = arrParam(5)					' 팝업 명칭 

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBillHdr(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtVatType.focus
	
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetItemCd()  ------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement 		
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetUnit(byval arrRet)
	frm1.txtUnit.Value		= arrRet(0)		
	lgBlnFlgChgValue		= True
End Function

'------------------------------------------  SetItemGroup()  ---------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value	= arrRet(0)		
	frm1.txtItemGroupNm.value   = arrRet(1)
	lgBlnFlgChgValue			= True
End Function

'------------------------------------------  SetBasisItemCd()  -------------------------------------------
'	Name : SetBasisItemCd()
'	Description : BasisItemCd Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetBasisItemCd(byval arrRet)
	
	lgBlnFlgChgValue = True
	
	If Not ChkBaseItem(frm1.txtItemCd1.value,arrRet(0)) Then Exit Function
	
	frm1.txtBasisItemCd.Value    = UCase(arrRet(0))
	frm1.txtBasisItemNm.Value    = arrRet(1)		
	
End Function

'------------------------------------------  SetWeightUnit()  --------------------------------------------
'	Name : SetWeightUnit()
'	Description : WeightUnit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWeightUnit(byval arrRet)
	frm1.txtWeightUnit.Value= arrRet(0)		
	lgBlnFlgChgValue		= True
End Function

'------------------------------------------  SetGrossWeightUnit()  --------------------------------------------
'	Name : SetGrossWeightUnit()
'	Description : WeightUnit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetGrossWeightUnit(byval arrRet)
	frm1.txtGrossWeightUnit.Value= arrRet(0)		
	lgBlnFlgChgValue		= True
End Function

'------------------------------------------  SetHSCd()  --------------------------------------------------
'	Name : SetHSCd()
'	Description : HSCd Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetHSCd(byval arrRet)
	frm1.txtHSCd.Value		= arrRet(0)
	frm1.txtHSUnit.value	= arrRet(3)
	
	lgBlnFlgChgValue		= True
End Function

Sub SetCookieVal()
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm")
	End If	

	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
End Sub

'------------------------------------------  SetBillHdr()  -----------------------------------------------
'	Name : SetBillHdr()
'	Description : Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetBillHdr(Byval arrRet)
	frm1.txtVatType.value = arrRet(0)
	frm1.txtVatTypeNm.value = arrRet(1)
	frm1.txtVatRate.Text = arrRet(2)
	lgBlnFlgChgValue = true

End Function
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'=============================================  2.5.1 JumpItemByPlant()  ======================================
'=	Event Name : JumpItemByPlant
'=	Event Desc :
'========================================================================================================

Function JumpItemByPlant()
	Dim intRet
	
	'------ Check previous data area ------
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRet = DisplayMsgBox("900017",parent.VB_YES_NO, "X", "X")												
		If intRet = vbNo Then				 
			Exit Function
		End If
	End If
	
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value
	
	PgmJump(BIZ_PGM_JUMPITEMBYPLANT_ID)

End Function

'=============================================  2.5.2 JumpItemImage()  ======================================
'=	Event Name : JumpItemImage
'=	Event Desc :
'========================================================================================================

Function JumpItemImage()
	Dim intRet
	
	'------ Check previous data area ------
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRet = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")												
		If intRet = vbNo Then
			Exit Function
		End If
	End If

	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value
	
	PgmJump(BIZ_PGM_JUMPITEMIMG_ID)
	
End Function

'=============================================  2.5.2 LookUpHsCd()  ======================================
'=	Event Name : LookUpHsCd
'=	Event Desc :
'========================================================================================================

Sub LookUpHsCd()
	Err.Clear                                                               
    LayerShowHide(1)
		
    Dim strVal
    strVal = BIZ_PGM_LOOKUPHSCD_ID & "?txtMode=" & parent.UID_M0001				
    strVal = strVal & "&txtHsCd=" & Trim(frm1.txtHsCd.value)			
		    
	Call RunMyBizASP(MyBizASP, strVal)									
		
End Sub
'=============================================  2.5.2 LookUpHsNotOk()  ======================================
'=	Event Name : LookUpHsNotOk
'=	Event Desc :
'========================================================================================================

Sub LookUpHsNotOk()
	If IsOpenPop = False Then
		Call DisplayMsgBox("126700", "X", "X", "X")
	End If
	frm1.txtHSCd.focus
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : txtVatType_OnChange
'   Event Desc : 부가세타입 내용이 변경되었을때 부가세율 계산 
'==========================================================================================
Sub txtVatType_OnChange()

	Dim VatType, VatTypeNm, VatRate

	VatType = Trim(frm1.txtVatType.value)
	
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

	frm1.txtVatTypeNm.value = VatTypeNm
	frm1.txtVatRate.text = VatRate
End Sub

'================================== =====================================================
' Function Name : InitCollectType
' Function Desc : 소비세유형코드/명/율 저장하기 
' 여기부터 키보드에서 소비세유형코드를 변경시 소비세유형명,소비세율,매입금액,NetAmount를 변경시키는 함수 
'========================================================================================
Sub InitCollectType()
	Dim i
	Dim iCodeArr, iNameArr, iRateArr
	
    On Error Resume Next

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & "  And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = REPLACE(iRateArr(i),".",parent.gComNumDec)
	Next
End Sub

'========================================================================================
' Function Name : GetCollectTypeRef====>ado이용추가 
' Function Desc : 
'========================================================================================
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

'===========================================  2.5.2 ChkBaseItem(strData1, strData2)  ====================
'=	Event Name : ChkBaseItem(strData1, strData2)
'=	Event Desc : 기준품목과 품목 동일 여부 체크 
'========================================================================================================

Function ChkBaseItem(strData1, strData2)
	
	ChkBaseItem = False
	
	If UCase(Trim(strData1)) = UCase(Trim(strData2)) Then
		Call DisplayMsgBox("127421", "X", "기준품목", "품목")
		
		frm1.txtBasisItemCd.value = ""
		frm1.txtBasisItemNm.value = "" 
		frm1.txtBasisItemCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	ChkBaseItem = True
	
End Function

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
Sub rdoPhantomType1_OnClick()
	If lgRdoOldVal1 = 1 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 1
End Sub

Sub rdoPhantomType2_OnClick()
	If lgRdoOldVal1 = 2 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 2   	
End Sub

Sub rdoUnifyPurFlg1_OnClick()
	If lgRdoOldVal2 = 1 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal2 = 1
End Sub

Sub rdoUnifyPurFlg2_OnClick()
	If lgRdoOldVal2 = 2 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal2 = 2   	
End Sub

Sub rdoValidFlg1_OnClick()
	If lgRdoOldVal3 = 1 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal3 = 1
End Sub

Sub rdoValidFlg2_OnClick()
	If lgRdoOldVal3 = 2 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal3 = 2   	
End Sub

Sub cboItemAcct_onchange()
	if frm1.cboItemAcct.value >= "30" And frm1.cboItemAcct.value <= "50" Then
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType1,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType2,"Q")
		frm1.rdoPhantomType2.checked = True
	Else
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType1,"N")
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType2,"N")
	End IF  
	lgBlnFlgChgValue = True
End Sub

Sub cboItemClass_onchange()
	lgBlnFlgChgValue = True
End Sub

Sub txtWeight_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtGrossWeight_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCBM_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtVatRate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtHsCd_onKeyPress()
	frm1.txtHSUnit.value = ""
	lgBlnFlgChgValue = True
End Sub


'네패스 이경석과장 요청으로 MES목코드 중복으로 등록하지 못하게 변경...2009.09.03...kbs
Function txtCBMInfo_Onchange()
    Dim strItemCd, StrMesItemCd, StrItemCnt
    Dim IntRetCD 

    On Error Resume Next
    Err.Clear
    
    
    txtCBMInfo_Onchange = False

    StrMesItemCd = Trim(frm1.txtCBMInfo.value)
    strItemCd    = Trim(frm1.txtItemCd1.value)


    Call CommonQueryRs(" Count(*) ", " B_ITEM ", " CBM_DESCRIPTION = " & FilterVar(StrMesItemCd, "''", "S")  & " And ITEM_CD <> " & FilterVar(strItemCd, "''", "S")  & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    If Err.number <> 0 Then
	MsgBox Err.Description 
	Err.Clear 
	Exit Function
    End If

    If lgF0 = "" Then Exit Function

    StrItemCnt = Split(lgF0, Chr(11))

    if StrItemCnt(0) > "0" Then
	IntRetCD = DisplayMsgBox("990033", parent.VB_YES_NO, StrMesItemCd, "X")			
	If IntRetCD = vbNo Then
		frm1.txtCBMInfo.value = frm1.hCBMInfo.value
		Exit Function
	End If
    End If

    frm1.hCBMInfo.value = frm1.txtCBMInfo.value

    txtCBMInfo_Onchange = True

End Function



'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
	lgBlnFlgChgValue = True	
End Sub

Sub txtHSCd_OnChange()
	lgBlnFlgChgValue = True	
	If frm1.txtHSCd.value <> "" Then
		LookUpHsCd()
	Else
		frm1.txtHSUnit.value = ""
	End If
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                        
    
    Err.Clear                                                               

	'-----------------------
    'Check previous data area
    '----------------------- 

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
	frm1.txtItemNm.value = ""
        frm1.hCBMInfo.value = ""
	
    Call ggoOper.ClearField(Document, "2")										

    Call SetDefaultVal    
    Call InitVariables															



	'-----------------------
    'Check condition area
    '----------------------- 
	If Not chkFieldByCell(frm1.txtItemCd, "A", 1) Then Exit Function
 
	'-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then
		Exit Function
	End If

    FncQuery = True																
        
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                         
    
	'-----------------------
    'Check previous data area
    '-----------------------

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")	           
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    
    Call ggoOper.ClearField(Document, "A")                                      '⊙: Clear Condition Field
	Call pvLockField("N")
	Call ggoOper.SetReqAttr(frm1.txtUnit, "N")
	Call ggoOper.SetReqAttr(frm1.rdoPhantomType1,"N")
	Call ggoOper.SetReqAttr(frm1.rdoPhantomType2,"N")
    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal    
    Call InitVariables															'⊙: Initializes local global variables
    frm1.txtItemCd1.focus
    Set gActiveElement = document.activeElement
      
    FncNew = True																'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim intRetCd

    FncDelete = False												
    
	'-----------------------
    'Precheck area
    '-----------------------

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If

	'-----------------------
    'Delete function call area
    '-----------------------%>
    IntRetCd = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		           
	If IntRetCd = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then
		Exit Function
	End If													
    
    FncDelete = True                                                    
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 

    Dim IntRetCD 

    FncSave = False                                                         
    Err.Clear                                                               
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                         
        Exit Function
    End If
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             
       Exit Function
    End If

    If txtCBMInfo_Onchange = False Then
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
       Exit Function
    End If
			                                               
    FncSave = True                                                         
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

    Dim IntRetCD

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")				
       If IntRetCD = vbNo Then
           Exit Function
       End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE
    
    ' 조건부 필드를 삭제한다.
    Call ggoOper.ClearField(Document, "1")
    Call pvLockField("N")
    Call ggoOper.SetReqAttr(frm1.txtUnit, "N")
    Call SetToolbar("11101000000011")
    
    frm1.txtItemCd1.value    = ""
    frm1.txtItemDesc.value   = ""
    frm1.txtValidFromDt.Text = StartDate
    frm1.txtValidToDt.Text   = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
    
    frm1.txtItemCd1.focus
    Set gActiveElement = document.activeElement  

    lgBlnFlgChgValue = True

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    On Error Resume Next                                                  
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
     On Error Resume Next                                                   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    On Error Resume Next                                                   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                            
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    '------------------------------------
    'Data Sheet 초기화 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										
    
    Call SetDefaultVal
    Call InitVariables															
	
    Err.Clear                                                               
    
    LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)			
	strVal = strVal & "&PrevNextFlg=" & "P"    
	
	Call RunMyBizASP(MyBizASP, strVal)										
	
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim strVal
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                            
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet 초기화 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")									
    
    Call SetDefaultVal
    Err.Clear         
    
    LayerShowHide(1)
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)			
	strVal = strVal & "&PrevNextFlg=" & "N"    
	
	Call RunMyBizASP(MyBizASP, strVal)										

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         
End Function
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
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
'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Err.Clear                                                               
    
    DbDelete = False													
    
    LayerShowHide(1)						
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003					
    strVal = strVal & "&txtItemCd1=" & Trim(frm1.txtItemCd1.value)		
	
	Call RunMyBizASP(MyBizASP, strVal)									
	
    DbDelete = True                                                     

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()														
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                               
    
    DbQuery = False                                                         
    
    LayerShowHide(1)							
    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)			
	strVal = strVal & "&PrevNextFlg=" & ""

	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbQuery = True                                                          

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														
	
    '-----------------------
    'Reset variables area
    '-----------------------    
    Dim LayerN1

	Set LayerN1 = window.document.all("MousePT").style
	
    lgIntFlgMode = parent.OPMD_UMODE											
    lgBlnFlgChgValue = false
	frm1.txtItemNm1.focus 
	Set gActiveElement = document.activeElement 								
    Call pvLockField("Q")
	Call SetToolbar("11111000111111")
	
	If frm1.cboItemAcct.value >= "30" And frm1.cboItemAcct.value <= "50" Then
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType1,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType2,"Q")
	Else
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType1,"N")
		Call ggoOper.SetReqAttr(frm1.rdoPhantomType2,"N")
	End If
	
	If frm1.txtItemByPlantFlg.value = "Y"  Then
		Call ggoOper.SetReqAttr(frm1.txtUnit, "Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtUnit, "N")
	End If

	frm1.hCBMInfo.value = frm1.txtCBMInfo.value

End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															

	If frm1.txtHSCd.value = "" Then
		frm1.txtHSUnit.value = ""
	End IF
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function       
	
	LayerShowHide(1)
		
    Dim strVal

	With frm1
		.txtMode.value = parent.UID_M0002										
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUsrId.value = parent.gUsrID 
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	End With
	
    DbSave = True                                                         
    
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															

    dim LayerN1

    
	Set LayerN1 = window.document.all("MousePT").style
	
    frm1.txtItemCd.value = frm1.txtItemCd1.value 
    Call InitVariables
    
    Call MainQuery()

End Function

'========================================================================================
' Function Name : pvLockField
' Function Desc : ggoOperLockField 대용 
'========================================================================================
Function pvLockField(byVal pvFlag) 
	If pvFlag = "Q" Then
		Call LockHTMLField(frm1.txtItemCd1,"P")
		Call LockHTMLField(frm1.txtCBMInfo,"P")     '2008-04-04 11:26오전 :: hanc
		Call LockObjectField(frm1.txtValidFromDt,"P")
	ElseIf pvFlag = "N" Then
		Call LockHTMLField(frm1.txtItemCd1,"R")
		Call LockHTMLField(frm1.txtCBMInfo,"R")     '2008-04-04 11:26오전 :: hanc
		Call LockObjectField(frm1.txtValidFromDt,"R")
	End If
End Function
