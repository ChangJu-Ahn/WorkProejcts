<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111ma4
'*  4. Program Name         : 단가확정 
'*  5. Program Desc         : 단가확정 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/05/15
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit				

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "m3111mb4.asp"	
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

'이성룡 추가 
Dim C_SelCheck
Dim C_ConfirmYN

Dim C_PoNo
Dim C_PoSeq
Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemSpec
Dim C_PoDt
Dim C_PoQty
Dim C_PoUnit
Dim C_PoPrice1
Dim C_PoPrice2
Dim C_Check	
Dim C_PoCurrency
Dim C_PoAmt
Dim C_NetPoAmt
Dim C_VatAmt
Dim C_IOFlg 
Dim C_IOFlg_cd
Dim C_VatType
Dim C_VatNm
Dim C_VatRate
Dim C_SupplierCd
Dim C_SupplierNm
Dim C_IvQty			' 2005-10-21 매입수량 > 0


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop  
'이성룡 추가 (단가적용규칙)
Dim lsPriceType     
Dim lsClickBtnPrice   
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim StartDate,EndDate
EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE     
    lgBlnFlgChgValue = False      
    lgIntGrpCount = 0             
    lgStrPrevKey = ""             
    lgLngCurRows = 0              
    frm1.vspdData.MaxRows = 0
End Sub

'==========================================   Selection_Sel()  ======================================
'	Name : Selection_Sel()
'	Description : 일괄선택버튼의 Event 합수 
'=========================================================================================================
Sub Selection_Sel(ByVal pFlag)
	Dim index,Count
	Dim strColValue
	
	frm1.vspdData.ReDraw = false
	frm1.vspdData.Col = C_SelCheck
	
	Count = frm1.vspdData.MaxRows 
	
	With frm1.vspdData
	
		If Trim(pFlag) = "ON" Then '일괄선택 버튼 클릭시 
			For index = 1 to Count
				.Row = index
				.text = "1"
			Next
		Else					'일괄선택취소 버튼 클릭시 
			For index = 1 to Count
				.Row = index
				.text = "0"
			Next

		End If
		
	End With
	
	frm1.vspdData.ReDraw = true
	lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : btnSelect_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnSelect_OnClick()
	If frm1.vspdData.Maxrows > 0 then
	    Call Selection_Sel("ON")
	End If
End Sub

'==========================================================================================
'   Event Name : btnDisSelect_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnDisSelect_OnClick()
	If frm1.vspdData.Maxrows > 0 then
	    Call Selection_Sel("OFF")
	End If
End Sub

'==========================================================================================
'   Event Name : btnCallPrice_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnCallPrice_OnClick()

	Dim index
	
	If frm1.vspdData.Maxrows <= 0 then
		Exit Sub
	End if
		'이성룡 추가 
	Call SetPriceType2
	
	call Selection()
	
	For index = 1 to  frm1.vspdData.Maxrows
	    frm1.vspdData.row = index
	    frm1.vspdData.Col = C_SelCheck
	    
	    If frm1.vspdData.Text = "1" then
			frm1.vspdData.Col = 0
			ggoSpread.UpdateRow index
	    Else
			'frm1.vspdData.Col = 0
			'ggoSpread.EditUndo
	    End If
	    
	Next 
	
End Sub

' === 2005.07.06 단가 일괄 불러오기 관련 수정 ===========================================
Sub btnCallPrice_Ok()
Dim lRow	
	With frm1
	For lRow = 1 To .vspdData.MaxRows
				
		.vspdData.Row = lRow
		.vspdData.Col = C_Check
	
		If .vspdData.Text <> "0" Then
			Call vspdData_Change(C_PoPrice2, lRow)
		End If
	
	Next
	End With
End Sub
' === 2005.07.06 단가 일괄 불러오기 관련 수정 ===========================================

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
    frm1.txtGroupCd.Value	= Parent.gPurGrp
    frm1.txtStampDt.Text	= EndDate
    frm1.txtFrDt.Text		= StartDate
    frm1.txtToDt.Text		= EndDate
    Call SetToolbar("1110000000001111")
    frm1.txtGroupCd.focus
	Set gActiveElement = document.activeElement
	
	'이성룡 추가(하단버튼 화면 초기화시 Disabled)
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
	frm1.btnCallPrice.disabled = True
	
	If lsPriceType = "T" then
		frm1.rdoPrcTypeFlg(0).checked = true
	Else
		frm1.rdoPrcTypeFlg(1).checked = true
	End If
	'이성룡 추가 끝 

End Sub

'==========================================  2.2.1 SetPriceType()  ========================================
'	Name : SetPriceType()
'	Description : 화면초기화 할때 단가적용규치 Setting
'=========================================================================================================
Sub SetPriceType()

	Dim IntRetCd
	
	IntRetCD = CommonQueryRs("MINOR_CD", "B_CONFIGURATION", "(MAJOR_CD = 'M0001' AND REFERENCE = 'Y' )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	lsPriceType = TRIM(REPLACE(lgF0,CHR(11),""))
	
	frm1.txtPrcType.value = lsPriceType			'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	
End Sub


'==========================================  2.2.2 SetPriceType()  ========================================
'	Name : SetPriceType2()
'	Description : 단가를 가져오기 전에 단가적용규칙을 가져온다.
'=========================================================================================================
Sub SetPriceType2()

	If frm1.rdoPrcTypeflg1.checked = true then
		lsPriceType = "T"
		frm1.txtPrcType.value = "T"				'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	Else
		lsPriceType = "N"
		frm1.txtPrcType.value = "N"				'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	End if
	
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	'이성룡 추가(단가규칙 적용)
	C_SelCheck	= 1
	C_ConfirmYN	= 2
	
	C_PoNo		= 3
	C_PoSeq		= 4
	C_PlantCd	= 5
	C_PlantNm	= 6
	C_ItemCd	= 7
	C_ItemNm	= 8
	C_ItemSpec	= 9
	C_PoDt		= 10
	C_PoQty		= 11
	C_PoUnit	= 12
	C_PoPrice1	= 13
	C_PoPrice2	= 14
	C_Check		= 15
	C_PoCurrency= 16
	C_PoAmt		= 17
	C_NetPoAmt  = 18
	C_VatAmt    = 19
	C_IOFlg     = 20
	C_IOFlg_cd  = 21
	C_VatType   = 22
	C_VatNm     = 23
	C_VatRate   = 24
	C_SupplierCd= 25
	C_SupplierNm= 26
	C_IvQty		= 27
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData		
		ggoSpread.Spreadinit "V200512010",,Parent.gAllowDragDropSpread  

		.ReDraw = false

		.MaxCols = C_IvQty + 1
		.Col = .MaxCols:    .ColHidden = True
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
		
		'이성룡 추가 
		ggoSpread.SSSetCheck	C_SelCheck, "선택",5,,,true
		ggoSpread.SSSetEdit		C_ConfirmYN,"확정여부",10
		
		ggoSpread.SSSetCheck	C_Check, "",3,,,true
		ggoSpread.SSSetEdit		C_PoNo,"발주번호",20
		ggoSpread.SSSetEdit		C_PoSeq,"발주순번",10
		ggoSpread.SSSetEdit		C_PlantCd, "공장", 10
		ggoSpread.SSSetEdit		C_PlantNm, "공장명", 20
		ggoSpread.SSSetEdit		C_ItemCd, "품목", 10
		ggoSpread.SSSetEdit		C_ItemNm, "품목명", 20
		ggoSpread.SSSetEdit		C_ItemSpec, "품목규격", 20
		ggoSpread.SSSetDate		C_PoDt,"발주일", 10, 2, Parent.gDateFormat
		SetSpreadFloat			C_PoQty, "발주수량",15,1,3
		ggoSpread.SSSetEdit		C_PoUnit, "단위", 10
		ggoSpread.SSSetFloat	C_PoPrice1, "가단가", 15	,"C" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_PoPrice2, "진단가", 15	,"C" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PoCurrency, "화폐", 10
		ggoSpread.SSSetFloat 	C_PoAmt, "금액",18,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat 	C_NetPoAmt, "발주순금액",18,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat 	C_VatAmt, "VAT금액",15,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit		C_IOFlg, "VAT포함여부", 15
		ggoSpread.SSSetEdit		C_IOFlg_cd, "C_IOFlg_cd", 10
		ggoSpread.SSSetEdit		C_VatType, "VAT", 7,,,4,2
		ggoSpread.SSSetEdit		C_VatNm, "VAT명", 20 
		SetSpreadFloat			C_VatRate, "VAT율(%)",15,1,5
		ggoSpread.SSSetEdit		C_SupplierCd, "공급처", 10
		ggoSpread.SSSetEdit		C_SupplierNm, "공급처명", 20
		SetSpreadFloat			C_IvQty, "매입수량",15,1,3

		Call ggoSpread.SSSetColHidden(C_IOFlg_cd,C_IOFlg_cd,True)	
		Call ggoSpread.SSSetColHidden(C_IvQty,C_IvQty,True)	

		.ReDraw = true
		Call SetSpreadLock 
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = false
    ggoSpread.SpreadLock		-1, -1
    ggoSpread.SpreadUnLock	C_PoPrice2, -1, C_PoPrice2, -1
    ggoSpread.SpreadUnLock	C_Check, -1, C_Check, -1
    '이성룡 추가(단가정책 컬럼 추가)
    ggoSpread.SpreadUnLock	C_SelCheck, -1, C_SelCheck, -1
    
    ggoSpread.SSSetRequired	C_PoPrice2, -1, -1
    .vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow)
    With frm1
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, lRow, lRow
    '이성룡 추가 
    ggoSpread.SSSetProtected	C_ConfirmYN, lRow, lRow
    
    ggoSpread.SSSetProtected	C_ReqNo, lRow, lRow
    ggoSpread.SSSetProtected	C_ItemCd, lRow, lRow
    ggoSpread.SSSetProtected	C_ItemNm, lRow, lRow
    ggoSpread.SSSetRequired		C_PlantCd, lRow, lRow
    ggoSpread.SSSetRequired		C_Popup, lRow, lRow
    ggoSpread.SSSetProtected	C_PlantNm, lRow, lRow
    ggoSpread.SSSetProtected	C_ReqQty, lRow, lRow
    ggoSpread.SSSetProtected	C_Unit, lRow, lRow
    ggoSpread.SSSetProtected	C_NetPoAmt, lRow, lRow
    ggoSpread.SSSetProtected	C_VatAmt, lRow, lRow
    ggoSpread.SSSetProtected	C_IOFlg, lRow, lRow
    ggoSpread.SSSetProtected	C_VatType, lRow, lRow
    ggoSpread.SSSetProtected	C_VatNm, lRow, lRow
    ggoSpread.SSSetProtected	C_VatRate, lRow, lRow 
    ggoSpread.SSSetProtected	C_ReqDt, lRow, lRow
    ggoSpread.SSSetProtected	C_PlanDt, lRow, lRow
    ggoSpread.SSSetProtected	C_ReqState, lRow, lRow
    .vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			'이성룡 추가 
			C_SelCheck	= iCurColumnPos(1)
			C_ConfirmYN	= iCurColumnPos(2)
			C_PoNo		= iCurColumnPos(3)
			C_PoSeq		= iCurColumnPos(4)
			C_PlantCd	= iCurColumnPos(5)
			C_PlantNm	= iCurColumnPos(6)
			C_ItemCd	= iCurColumnPos(7)
			C_ItemNm	= iCurColumnPos(8)
			C_ItemSpec	= iCurColumnPos(9)
			C_PoDt		= iCurColumnPos(10)
			C_PoQty		= iCurColumnPos(11)
			C_PoUnit	= iCurColumnPos(12)
			C_PoPrice1	= iCurColumnPos(13)
			C_PoPrice2	= iCurColumnPos(14)
			C_Check		= iCurColumnPos(15)
			C_PoCurrency= iCurColumnPos(16)
			C_PoAmt		= iCurColumnPos(17)
			C_NetPoAmt  = iCurColumnPos(18)
			C_VatAmt    = iCurColumnPos(19)
			C_IOFlg     = iCurColumnPos(20)
			C_IOFlg_cd  = iCurColumnPos(21)
			C_VatType   = iCurColumnPos(22)
			C_VatNm     = iCurColumnPos(23)
			C_VatRate   = iCurColumnPos(24)
			C_SupplierCd= iCurColumnPos(25)
			C_SupplierNm= iCurColumnPos(26)
			C_IvQty		= iCurColumnPos(27)
	End Select
End Sub	

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	Dim strRet
	Dim arrParam
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	IsOpenPop = True
		
	Redim arrParam(2)
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag

	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus
	End If	
End Function

'------------------------------------------  OpenSupplier()  ---------------------------------------------
'	Name : OpenSupplier()
'	Description : OpenSupplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"				
	arrParam(1) = "B_BIZ_PARTNER"			

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"				
    arrField(1) = "BP_NM"				
    
    arrHeader(0) = "공급처"			
    arrHeader(1) = "공급처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
	End If	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	If UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(1) = "" 						
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	arrParam(4) = EndDate						' Current Date
			
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    arrField(2) = 4
    arrField(3) = 5
    arrField(4) = 8
    
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
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
	End If	
	
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenGroup()  -------------------------------------------------
'	Name : OpenGroup()
'	Description : OpenGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)		
		frm1.txtGroupCd.focus
	End If	

End Function 

'==========================================   lookupPriceForSelection()  =============================
'	Name : lookupPriceForSelection()
'	Description :
'=====================================================================================================
Function lookupPriceForSelection()
    Err.Clear
    Dim strVal
    Dim lColSep,lRowSep
    Dim lRow        
    Dim lGrpCnt     
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If
    
	lgBlnFlgChgValue = true
    
    If LayerShowHide(1) = False Then Exit Function
    
    '이성룡 주석처리 
	'Call RunMyBizASP(MyBizASP, strVal)
	

	With frm1		
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1

    strVal = ""
    
    '-----------------------
    'Data manipulate area
    '-----------------------

	.txtMode.value = "lookupPriceForSelection"	

	For lRow = 1 To .vspdData.MaxRows
				
		.vspdData.Row = lRow
		.vspdData.Col = C_Check
	
		If .vspdData.Text <> "0" Then
					
			frm1.vspdData.Row = lRow
			frm1.vspdData.Col = C_SupplierCd
			strVal = strVal & Trim(frm1.vspdData.Value) & Parent.gColSep
			frm1.vspdData.Col = C_ItemCd
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_PlantCd
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_PoUnit
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_PoCurrency
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_PoPrice1
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			strVal = strVal & lRow & Parent.gRowSep
					
			lGrpCnt = lGrpCnt + 1

			frm1.vspdData.Col = C_PoPrice2
			frm1.vspdData.Text = 0
		End If
	Next
	
	If strVal <> "" Then
		If LayerShowHide(1) = False Then Exit Function
		
		.hdnMaxRows.value = .vspdData.MaxRows
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End If	
	End With
End Function 
'==========================================   lookupPrice()  ======================================
'	Name : lookupPrice()
'	Description :
'==================================================================================================
Function lookupPrice(ByVal Row)
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If
    
    Dim strVal
    
	lgBlnFlgChgValue = true

	'frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = Row
    strVal = BIZ_PGM_ID & "?txtMode=" & "lookupPrice"
    strVal = strVal & "&txtStampDt=" & Trim(frm1.txtStampDt.text)
	frm1.vspdData.Col = C_SupplierCd
    strVal = strVal & "&txtBpCd=" & Trim(frm1.vspdData.Value)
	frm1.vspdData.Col = C_ItemCd
    strVal = strVal & "&txtItemCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_PlantCd
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_PoUnit
    strVal = strVal & "&txtUnit=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_PoCurrency
    strVal = strVal & "&txtCurrency=" & Trim(frm1.vspdData.text)
    strVal = strVal & "&txtRow=" & Row
    strVal = strVal & "&txtPrcType=" & Trim(lsPriceType)
	frm1.vspdData.Col = C_PoPrice2
	frm1.vspdData.Text = 0
	
    If LayerShowHide(1) = False Then 
		Exit Function
	End If
	Call RunMyBizASP(MyBizASP, strVal)

End Function 

'==========================================   Selection()  ======================================
'	Name : Selection()
'	Description : 일괄선택버튼의 Event 합수 
'================================================================================================
Sub Selection()
	Dim index,Count
	Dim lookupflg

	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
	
	'이성룡 추가 
	frm1.txtPriceMsg.value = "FALSE"
	
	frm1.vspdData.Row = index
	frm1.vspdData.Col = C_Check
	
	frm1.vspdData.text = "0"
	
	frm1.vspdData.Col = C_SelCheck	
		
		If frm1.vspdData.text = "1" then
		
			frm1.vspdData.Row = index
			frm1.vspdData.Col = C_Check
		
		
			if frm1.vspdData.Text = "1" then
				frm1.vspdData.Text = "0"
			else
				frm1.vspdData.Text = "1"
				lookupflg = true
			End if
		
		End If
				
	Next 
	
	frm1.vspdData.ReDraw = true
	
	lgBlnFlgChgValue = true
	
	If lookupflg Then
		If Not chkField(Document, "2") Then
		   Exit sub
		End If
		Call lookupPriceForSelection()
	End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	If frm1.vspdData.MaxRows > 0 Then
		Call SetPopupMenuItemInf("0001111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    		
End Sub
'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_PoCurrency,C_PoPrice1,"C" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_PoCurrency,C_PoPrice2,"C" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_PoCurrency,C_PoAmt,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_PoCurrency,C_NetPoAmt,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_PoCurrency,C_VatAmt,"A" ,"I","X","X")

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim PoQty
    Dim PoPrice
    Dim vat,VatRt
    Dim PoAmt
    Dim PoCurrency
    ggoSpread.Source = frm1.vspdData

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = 0
	
	
	If Col = C_Check And ggoSpread.UpdateFlag = frm1.vspdData.Text Then
		ggoSpread.EditUndo
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_PoPrice1,"C" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_PoPrice2,"C" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_PoAmt,"A" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_NetPoAmt,"A" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_VatAmt,"A" ,"X","X")

		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_PoPrice1,"C" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_PoPrice2,"C" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_PoAmt,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_NetPoAmt,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_VatAmt,"A" ,"I","X","X")

	ElseIf Col = C_Check And ggoSpread.UpdateFlag <> frm1.vspdData.Text Then
		Call SetPriceType2	
		Call lookupPrice(Row)
		ggoSpread.UpdateRow Row
		frm1.vspdData.Col = C_SelCheck
		frm1.vspdData.Text = "1"
	ElseIf Col = C_PoPrice2 Then

		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_Check
	    
	    frm1.vspdData.Col = C_PoCurrency 
	    PoCurrency = frm1.vspdData.Text		
	    	
		frm1.vspdData.Col = C_PoQty       '수량 
        PoQty = frm1.vspdData.Text		
		
		frm1.vspdData.Col = C_PoPrice2    '진단가 
		PoPrice = frm1.vspdData.Text		
		
		frm1.vspdData.Col = C_PoAmt       '발주금액 
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(PoQty) * Parent.UNICDbl(PoPrice),PoCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo ,"X")
		
		PoAmt = frm1.vspdData.Text    '발주금액 
        
        frm1.vspdData.Col = C_VatRate
        VatRt = frm1.vspdData.Text    'vat율		
		
		frm1.vspdData.Col = C_IOFlg_cd    '포함여부 
        
        if frm1.vspdData.Text = "2" then '포함	
            vat = ( UNICDbl(PoAmt) * UNICDbl(VatRt) ) / ( UNICDbl(VatRt) + 100) 'CInt(DocAmt * VatRt / (VatRt + 100))        

            frm1.vspdData.Col = C_VatAmt    'vat금액 
            frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(vat,PoCurrency,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo , "X")
            vat = UNICDbl(frm1.vspdData.Text)
         
            frm1.vspdData.Col = C_NetPoAmt    'net금액 
            frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency( UNICDbl(PoAmt) - vat,PoCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo ,"X")
        else                             '별도 
            vat = ( UNICDbl(PoAmt) * UNICDbl(VatRt) ) / 100 'CInt(DocAmt * VatRt / (VatRt + 100))        
            
            frm1.vspdData.Col = C_VatAmt    'vat금액 
            frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(vat,PoCurrency,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo , "X")
        
            frm1.vspdData.Col = C_NetPoAmt    'net금액 
            frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency( UNICDbl(PoAmt),PoCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo ,"X")
                   
        end if 	
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_PoPrice1,"C" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_PoPrice2,"C" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_PoAmt,"A" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_NetPoAmt,"A" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_PoCurrency,C_VatAmt,"A" ,"X","X")

		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_PoPrice1,"C" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_PoPrice2,"C" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_PoAmt,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_NetPoAmt,"A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_PoCurrency,C_VatAmt,"A" ,"I","X","X")

		ggoSpread.UpdateRow Row
	elseif Col <> C_Check then
		ggoSpread.UpdateRow Row
	End if
		
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
	Call CheckMinNumSpread(frm1.vspdData, Col, Row) 
End Sub

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '과부족허용율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999.9999"
    End Select
End Sub
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                
    Call ggoOper.LockField(Document, "N")                              
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet     
    Call SetPriceType                                              
    Call SetDefaultVal
    Call InitVariables   

End Sub

'==========================================================================================
'   Event Name : OCX Event
'   Event Desc :
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End if
End Sub

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End if
End Sub

Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtStampDt_DblClick(Button)
	if Button = 1 then
		frm1.txtStampDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtStampDt.Focus
	End if
End Sub
Sub txtStampDt_Change()
	lgBlnFlgChgValue = true	
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
 
    FncQuery = False                                                 
    
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")					'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)					
    frm1.txtStampDt.Text = EndDate
    Call InitVariables
    
    '-----------------------
    'Check Price Type(이성룡 단가적용규칙 추가)
    '-----------------------
    If lsPriceType <> "T" And lsPriceType <> "N" Then
    	Call DisplayMsgBox("171214", "X", "X", "X")      '☜ : No data is found. 
    	Exit Function
    End If			
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
    
	with frm1
		
		if (UniConvDateToYYYYMMDD(.txtFrDt.Text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.Text,Parent.gDateFormat,"")) and Trim(.txtFrDt.Text)<>"" and Trim(.txtToDt.Text)<>"" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")
			Exit Function
		End if   
		
	End with
	
	'----------------------------------------------------------------
    'Set Parameter to Hidden area (Added By Lee Sung Yong 2005/01/28)
    '----------------------------------------------------------------
    
    With frm1
    
        
	.hdnSupplier.value = Trim(.txtSupplierCd.value)
	.hdnGroup.value = Trim(.txtGroupCd.value)
	.hdnPoNo.Value = Trim(.txtPoNo.value)
	.hdnFrDt.Value = Trim(.txtFrDt.value)
	.hdnToDt.Value = Trim(.txtToDt.value)
	.hdnItemCd.Value = Trim(.txtItemCd.value)
    '-----------------------
    'Check Price Type2(이성룡 단가확정여부추가)
    '-----------------------
    If .rdoCfmFlg1.checked = true Then
		.hdnCfmFlg.value = "T"
	Else
		.hdnCfmFlg.value = "F"
    End If											
    
    end with    	
	
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
	Set gActiveElement = document.activeElement
    FncQuery = True			
    											

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                      
    
    Err.Clear                                                           
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                             
    Call ggoOper.ClearField(Document, "2")                             
    Call ggoOper.LockField(Document, "N")                              
    Call SetDefaultVal
    Call InitVariables                                                 
    
    
    
	Set gActiveElement = document.activeElement
    FncNew = True                                                      
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                  
    
    Err.Clear                                                        
    'On Error Resume Next   
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                         
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                          
       Exit Function
    End If
	
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncSave = True                                               
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                

	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_PoCurrency,C_PoPrice1,"C" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_PoCurrency,C_PoPrice2,"C" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_PoCurrency,C_PoAmt,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_PoCurrency,C_NetPoAmt,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_PoCurrency,C_VatAmt,"A" ,"I","X","X")

	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncExport(Parent.C_SINGLEMULTI)							
	Set gActiveElement = document.activeElement
End Function
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                       
	Set gActiveElement = document.activeElement
End Function
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	
	Dim IntRetCD
	
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    Err.Clear           
    
    If LayerShowHide(1) = False Then Exit Function
    
	Dim strVal
    
    With frm1
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtSupplierCd=" & .hdnSupplier.value
	    strVal = strVal & "&txtGroupCd=" & .hdnGroup.value
	    strVal = strVal & "&txtPoNo=" & .hdnPoNo.value
		strVal = strVal & "&txtFrDt=" & .hdnFrDt.Value
		strVal = strVal & "&txtToDt=" & .hdnToDt.Value
		'이성룡 추가(단가확정여부)
		strVal = strVal & "&txtCfmFlg=" & .hdnCfmFlg.Value
		'품목추가 
		strVal = strVal & "&txtItemCd=" & .hdnItemCd.value
	else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)
	    strVal = strVal & "&txtGroupCd=" & Trim(.txtGroupCd.value)
	    strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
		strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.Text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
		'이성룡 추가(단가확정여부)
		strVal = strVal & "&txtCfmFlg=" & .hdnCfmFlg.Value		
		'품목추가 
		strVal = strVal & "&txtItemCd=" & .txtItemCd.value
	End if

		.hdnmaxrows.value = .vspdData.MaxRows	
		
	Call RunMyBizASP(MyBizASP, strVal)						
        
    End With
    
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
    lgIntFlgMode = Parent.OPMD_UMODE								
    
	lgBlnFlgChgValue = False
		
    Call ggoOper.LockField(Document, "Q")	
	Call SetToolbar("11101001000111")
	
	'이성룡 추가 (조회후 하단버튼 활성화)
	frm1.btnSelect.disabled = False
	frm1.btnDisSelect.disabled = False
	frm1.btnCallPrice.disabled = False
		
	frm1.txtStampDt.Text = EndDate
	frm1.vspdData.focus
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
	
End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt
	Dim strVal
    Dim lColSep,lRowSep

	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
    Dim ii

    If LayerShowHide(1) = False Then Exit Function
    
    DbSave = False                                                       
    
	With frm1
		.txtMode.value = Parent.UID_M0002
		lGrpCnt = 1
		strVal = ""
		lColSep = parent.gColSep
		lRowSep = parent.gRowSep

		iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
		ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

		iTmpCUBufferCount = -1
		iTmpDBufferCount = -1

		strCUTotalvalLen = 0
		strDTotalvalLen  = 0
  
		For lRow = 1 To .vspdData.MaxRows
					
			If Trim(GetSpreadText(.vspdData,0,lRow,"X","X")) = ggoSpread.UpdateFlag Then
				.vspdData.Col = C_IvQty
				' 2005-10-21 매입수량 > 0
				If UNICDbl(.vspdData.Value) > 0 Then
					Call LayerShowHide(0)
					strVal = ""
					Call displaymsgbox("174201", "x", lRow & "행", "x")
					Exit Function
				End If

				strVal = Trim(GetSpreadText(.vspdData,C_PoNo,lRow,"X","X")) & lColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PoSeq,lRow,"X","X")) & lColSep
				strVal = strVal & UNIConvNum(GetSpreadText(.vspdData,C_PoPrice2,lRow,"X","X"),0) & lColSep
				strVal = strVal & UNIConvNum(GetSpreadText(.vspdData,C_NetPoAmt,lRow,"X","X"),0) & lColSep
				strVal = strVal & UNIConvNum(GetSpreadText(.vspdData,C_VatAmt,lRow,"X","X"),0) & lColSep
        		strVal = strVal & lRow & lRowSep
						
				lGrpCnt = lGrpCnt + 1
			End If

			Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
			    Case ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			End Select   
			
		Next
		
		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
		End If   
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)						
	End With
	
    DbSave = True                                               
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()												
    lgBlnFlgChgValue = False
    frm1.rdoCfmflg1.checked = True
	Call InitVariables
	Call MainQuery()
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute 
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>단가확정</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="구매그룹"  NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
														   <INPUT TYPE=TEXT ID="txtGroupNm" NAME="arrCond" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발주번호"  NAME="txtPoNo" SIZE=26 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처"  NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
														   <INPUT TYPE=TEXT ID="txtSupplierNm" NAME="arrCond" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>발주일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주일 NAME="txtFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주일 NAME="txtToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">
														 <INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14" ALT="품목명"></TD>
									<TD CLASS="TD5" NOWRAP>확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="확정여부" NAME="rdoCfmflg" id = "rdoCfmflg1" Value="Y" tag="1X"><label for="rdoCfmflg1">&nbsp;확정&nbsp;</label>
													 	   <INPUT TYPE=radio Class="Radio" ALT="확정여부" NAME="rdoCfmflg" id = "rdoCfmflg2" Value="N" checked tag="1X"><label for="rdoCfmflg2">&nbsp;미확정&nbsp;</label></TD>																				 	   
								</TR>				 	   
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>단가적용기준일</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=단가적용기준일 NAME="txtStampDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>단가적용규칙</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="확정여부" NAME="rdoPrcTypeflg" id = "rdoPrcTypeflg1" Value="Y" tag="1X"><label for="rdoPrcTypeflg1">&nbsp;진단가&nbsp;</label>
												 	   <INPUT TYPE=radio Class="Radio" ALT="확정여부" NAME="rdoPrcTypeflg" id = "rdoPrcTypeflg2" Value="N" checked tag="1X"><label for="rdoPrcTypeflg2">&nbsp;최신단가&nbsp;</label></TD>								

							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
    	<td <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
      <td WIDTH="100%">
		<table <%=LR_SPACE_TYPE_30%>>
			<tr> 
				<TD WIDTH=10>&nbsp;</TD>
				<td WIDTH="*" align="left">
				<button name="btnSelect" class="clsmbtn" >일괄선택</button>&nbsp;
				<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>&nbsp;
				<BUTTON NAME="btnCallPrice" CLASS="CLSMBTN">단가불러오기</BUTTON>
				</td>
				<TD WIDTH=10>&nbsp;</TD>
			</tr>
		</table>
      </td>
    </tr>
    <TR>
		<TD WIDTH=100% HEIGHT=<%=bizsize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=bizsize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCfmFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCfmFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPrcType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPriceMsg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPriceVar" tag="24">

</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
</HTML>
