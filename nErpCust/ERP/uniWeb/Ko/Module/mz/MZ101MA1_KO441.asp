<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   ***************************************** !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =======================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'==========================================  1.1.2 공통 Include   =====================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID 				= "MZ101MB1_KO441.asp"
const C_SHEETMAXROWS_D  = 100						'☜ : MB단과 꼭 일치시킬것.

'=== 첫번째 spread 상수 ===
Dim C1_RCPT_NO
Dim C1_RCPT_DT
Dim C1_PLANT_CD
Dim C1_ITEM_CD
Dim C1_ITEM_POP
Dim C1_ITEM_NM
Dim C1_SPEC
Dim C1_UNIT
Dim C1_UNIT_POP
Dim C1_RCPT_QTY
Dim C1_PRICE
Dim C1_AMT
Dim C1_ISSUE_QTY
Dim C1_BAL_QTY
Dim C1_RCPT_DOC_NO
Dim C1_CLOSE_YN

'=== 두번째 spread 상수 ===
Dim C2_SEQ
Dim C2_ISSUE_DT
Dim C2_ISSUE_TYPE
Dim C2_ISSUE_TYPE_NM
Dim C2_ISSUE_QTY
Dim C2_ISSUE_DOC_NO

Dim lgStrPrevKey2
Dim StartDate, EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)

Dim lgSortKey1
Dim lgSortKey2
Dim lgPageNo1
Dim lgSpdHdrClicked
'===================================================================================================================================
Dim IsOpenPop
Dim lgIntFlgModeM
Dim lgActiveRow
'===================================================================================================================================
Sub InitVariables()

	lgIntFlgMode = Parent.OPMD_CMODE				   'Indicates that current mode is Create mode
	lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False					'Indicates that no value changed

	lgSortKey1 = 2
	lgSortKey2 = 2
	lgPageNo = 0
	lgPageNo1 = 0
	lgActiveRow=0
	frm1.vspdData.MaxRows = 0
	frm1.vspdData1.MaxRows = 0

End Sub
'===================================================================================================================================
Sub SetDefaultVal()
	frm1.txtFrDt.text = startDate
	frm1.txtToDt.text = endDate
	frm1.txtPlantCd.focus
	Set gActiveSpdSheet = frm1.vspdData
	Set gActiveElement = document.activeElement
End Sub
'===================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'===================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)

	If pvSpdNo = "A" Then
		C1_RCPT_NO 			= 1
		C1_RCPT_DT 			= 2
		C1_PLANT_CD 		= 3
		C1_ITEM_CD 			= 4
		C1_ITEM_POP 		= 5
		C1_ITEM_NM 			= 6
		C1_SPEC 				= 7
		C1_UNIT 				= 8
		C1_UNIT_POP			= 9
		C1_RCPT_QTY 		= 10
		C1_PRICE 				= 11
		C1_AMT 					= 12
		C1_ISSUE_QTY 		= 13
		C1_BAL_QTY			= 14
		C1_RCPT_DOC_NO 	= 15
		C1_CLOSE_YN 		= 16	
	else
		C2_SEQ 					= 1	
		C2_ISSUE_DT 		= 2	
		C2_ISSUE_TYPE 	= 3	
		C2_ISSUE_TYPE_NM= 4
		C2_ISSUE_QTY 		= 5	
		C2_ISSUE_DOC_NO = 6	
	End If

End Sub
'===================================================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Then

		With frm1.vspdData

			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20081321",,Parent.gAllowDragDropSpread

			.ReDraw  = false

			.MaxCols = C1_CLOSE_YN + 1
			.Col = .MaxCols:		.ColHidden = True
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit   C1_RCPT_NO     , "입고번호", 15,,, 18, 2
			ggoSpread.SSSetDate		C1_RCPT_DT		, "입고일",10, 2, Parent.gDateFormat
			ggoSpread.SSSetEdit		C1_PLANT_CD		, "공장",10
			ggoSpread.SSSetEdit   C1_ITEM_CD     , "품목", 10,,, 18, 2
	  	ggoSpread.SSSetButton	C1_ITEM_POP		
			ggoSpread.SSSetEdit		C1_ITEM_NM		, "품목명",20
			ggoSpread.SSSetEdit		C1_SPEC				, "규격",20
			ggoSpread.SSSetEdit   C1_UNIT    		, "단위", 10,,, 10, 2
	  	ggoSpread.SSSetButton	C1_UNIT_POP		
			SetSpreadFloatLocal		C1_RCPT_QTY		, "입고수량",15,1,3	 
			SetSpreadFloatLocal		C1_PRICE			, "단가",15,1,4
			SetSpreadFloatLocal		C1_AMT				, "금액",15,1,2
			SetSpreadFloatLocal		C1_ISSUE_QTY	, "출고수량",15,1,3
			SetSpreadFloatLocal		C1_BAL_QTY	, "출고수량",15,1,3
			ggoSpread.SSSetEdit   C1_RCPT_DOC_NO, "입고면장번호", 20,,, 50, 2
	 	 	ggoSpread.SSSetCheck	C1_CLOSE_YN		,	"마감여부",		10,	,	,	True

      Call ggoSpread.SSSetColHidden(C1_PLANT_CD,C1_PLANT_CD,True)

			.ReDraw  = True

		End With
	Else

		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData1

			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20080321",,Parent.gAllowDragDropSpread

			.ReDraw = false

			.MaxCols = C2_ISSUE_DOC_NO + 1
			.Col = .MaxCols:		.ColHidden = True

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit		C2_SEQ					, "순번", 10
			ggoSpread.SSSetDate		C2_ISSUE_DT			, "출고일",10, 2, Parent.gDateFormat
	  	ggoSpread.SSSetCombo	C2_ISSUE_TYPE		,	"출고유형",15,		2
	  	ggoSpread.SSSetCombo	C2_ISSUE_TYPE_NM,	"출고유형",15,		2
			SetSpreadFloatLocal		C2_ISSUE_QTY		, "출고수량",15,1,3
			ggoSpread.SSSetEdit   C2_ISSUE_DOC_NO, "출고면장번호", 20,,, 50, 2

			Call ggoSpread.SSSetColHidden(C2_SEQ,	C2_SEQ,	True)
			Call ggoSpread.SSSetColHidden(C2_ISSUE_TYPE,	C2_ISSUE_TYPE,	True)

			.ReDraw = true
		End With
	End If

End Sub

'===================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	select case gActiveSpdSheet.id
	case "A"

		frm1.vspdData.ReDraw = False

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C1_RCPT_DT	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C1_PLANT_CD, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C1_ITEM_CD	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C1_ITEM_NM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C1_SPEC, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C1_UNIT	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C1_RCPT_QTY	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C1_PRICE	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C1_AMT	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C1_ISSUE_QTY, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C1_BAL_QTY, pvStartRow, pvEndRow

		frm1.vspdData.ReDraw = True
	case "B"

		frm1.vspdData1.ReDraw = False
		ggoSpread.Source = frm1.vspdData1

		ggoSpread.SSSetRequired		C2_ISSUE_DT	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C2_ISSUE_TYPE_NM	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C2_ISSUE_QTY	, pvStartRow, pvEndRow

		frm1.vspdData1.ReDraw = True
	End Select
End Sub

'===================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C1_RCPT_NO 			= iCurColumnPos(1)
			C1_RCPT_DT 			= iCurColumnPos(2)
			C1_PLANT_CD 		= iCurColumnPos(3)
			C1_ITEM_CD 			= iCurColumnPos(4)
			C1_ITEM_POP 		= iCurColumnPos(5)
			C1_ITEM_NM 			= iCurColumnPos(6)
			C1_SPEC 				= iCurColumnPos(7)
			C1_UNIT 				= iCurColumnPos(8)
			C1_UNIT_POP			= iCurColumnPos(9)
			C1_RCPT_QTY 		= iCurColumnPos(10)
			C1_PRICE 				= iCurColumnPos(11)
			C1_AMT 					= iCurColumnPos(12)
			C1_ISSUE_QTY 		= iCurColumnPos(13)
			C1_BAL_QTY			= iCurColumnPos(14)
			C1_RCPT_DOC_NO 	= iCurColumnPos(15)
			C1_CLOSE_YN 		= iCurColumnPos(16)

		Case "B"
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
			C2_SEQ 					= iCurColumnPos(1)	
			C2_ISSUE_DT 		= iCurColumnPos(2)	
			C2_ISSUE_TYPE 	= iCurColumnPos(3)	
			C2_ISSUE_TYPE_NM= iCurColumnPos(4)	
			C2_ISSUE_QTY 		= iCurColumnPos(5)	
			C2_ISSUE_DOC_NO = iCurColumnPos(6)	

	End Select
End Sub
'===================================================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	with frm1
		If pvSpdNo = "A" Then
			ggoSpread.Source = frm1.vspdData
			.vspdData.ReDraw = False

       ggoSpread.SSSetProtected		C1_RCPT_NO, -1, -1
       ggoSpread.SSSetRequired		C1_RCPT_DT, -1, -1
       ggoSpread.SSSetProtected		C1_PLANT_CD, -1, -1
       ggoSpread.SSSetRequired		C1_ITEM_CD, -1, -1
       ggoSpread.SSSetProtected		C1_ITEM_NM, -1, -1
       ggoSpread.SSSetProtected		C1_SPEC, -1, -1
       ggoSpread.SSSetRequired		C1_UNIT, -1, -1
       ggoSpread.SSSetRequired		C1_RCPT_QTY, -1, -1
       ggoSpread.SSSetRequired		C1_PRICE, -1, -1
       ggoSpread.SSSetRequired		C1_AMT, -1, -1
       ggoSpread.SSSetProtected		C1_ISSUE_QTY, -1, -1
       ggoSpread.SSSetProtected		C1_BAL_QTY, -1, -1

			.vspdData.ReDraw = True
		Else
			ggoSpread.Source = frm1.vspdData1
			.vspdData1.ReDraw = False
       ggoSpread.SSSetRequired		C2_ISSUE_DT, -1, -1
       ggoSpread.SSSetRequired		C2_ISSUE_TYPE_NM, -1, -1
       ggoSpread.SSSetRequired		C2_ISSUE_QTY, -1, -1
			.vspdData1.ReDraw = True
		End IF
	End With
End Sub

'===================================================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," b_user_defined_minor "," ud_major_cd = " & FilterVar("ZZ902", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    iCodeArr = lgF0
    iNameArr = lgF1
    
    ggoSpread.SetCombo replace(iCodeArr,Chr(11),vbTab), C2_ISSUE_TYPE
    ggoSpread.SetCombo replace(iNameArr,Chr(11),vbTab), C2_ISSUE_TYPE_NM
End Sub

'===================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "PLANT_CD"
	arrField(1) = "PLANT_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)
		frm1.txtPlantNm.value= arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If

End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명				
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function


'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd2()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd2(pRow)
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(GetSpreadText(frm1.vspdData,C1_ITEM_CD,pRow,"X","X"))		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명				
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		call frm1.vspdData.SetText(C1_ITEM_CD,pRow,arrRet(0))
		call frm1.vspdData.SetText(C1_ITEM_NM,pRow,arrRet(1))
		Call vspdData_Change(C1_ITEM_CD, pRow)
	End If	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
Function OpenUnit(pRow)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위"					
	arrParam(1) = "B_Unit_OF_MEASURE"		
	
	frm1.vspdData.Col=C1_UNIT
	frm1.vspdData.Row=pRow
	arrParam(2) = Trim(frm1.vspdData.text)	
	
	arrParam(4) = ""						
	arrParam(5) = "단위"					
	
    arrField(0) = "Unit"					
    arrField(1) = "Unit_Nm"					
    
    arrHeader(0) = "단위"				
    arrHeader(1) = "단위명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col=C1_UNIT
		frm1.vspdData.text= arrRet(0)	
		Call vspdData_Change(C1_UNIT, pRow)
	End If	
End Function

'===================================================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
					ByVal dColWidth , ByVal HAlign , _
					ByVal iFlag )

   Select Case iFlag
		Case 2															  '금액 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 3															  '수량 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo			,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 4															  '단가 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 5															  '환율 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 6															  '환율 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
	End Select

End sub

'===================================================================================================================================
Sub Form_Load()
	call LoadInfTB19029
	call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")								   '⊙: Lock  Suitable  Field
	call InitSpreadSheet("A")
	call InitSpreadSheet("B")												 '⊙: Setup the Spread sheet
	call InitVariables
	call SetDefaultVal
	call InitComboBox()
	call SetToolbar("1100110100001111")										'⊙: 버튼 툴바 제어 
End Sub
'===================================================================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrDt.Focus
	End if
End Sub
'===================================================================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End if
End Sub
'===================================================================================================================================
Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'===================================================================================================================================
Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'===================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###
 	gMouseClickStatus = "SPC"

 	Set gActiveSpdSheet = frm1.vspdData

 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")		 '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0001111111")		 '화면별 설정 
	End IF

	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		ElseIf lgSortKey1 = 2 Then

 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If

 		lgIntFlgModeM = Parent.OPMD_CMODE

 	End If

	if lgActiveRow<>Row Then
		call FncQuery2(Row)
	End if
	lgActiveRow=Row
End Sub
'===================================================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

	Dim iActiveRow
	Dim lngStartRow
	Dim iStrChildRow
	Dim i, K
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData1
	ggoSpread.Source = frm1.vspdData1

	With frm1

		If .vspdData1.MaxRows = 0 Then Exit Sub

		If Row <= 0 AND Col <> 0 Then
 			ggoSpread.Source = .vspdData1

 			.vspdData.Row = .vspdData.ActiveRow
 			.vspdData.Col = .vspdData.MaxCols
 			
		Else
 		End If
 	End With

End Sub
'===================================================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If lgSpdHdrClicked = 1 Then
		Exit Sub
	End If
	if frm1.vspddata.row = 0 then exit sub
End Sub
'===================================================================================================================================
sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	if frm1.vspddata1.row = 0 then exit sub

End Sub
'===================================================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row)

	With frm1
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.UpdateRow Row
	End with

	call vspdData_Change(1,frm1.vspdData.ActiveRow)
End Sub
'===================================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

	With frm1
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
	End with
	Select case Col
	case C1_RCPT_QTY,C1_PRICE			
			Call frm1.vspdData.SetText(C1_AMT,	Row,	UniConvNum(GetSpreadText(frm1.vspdData,C1_RCPT_QTY,Row,"X","X"),0) * UniConvNum(GetSpreadText(frm1.vspdData,C1_PRICE,Row,"X","X"),0))
	End Select
End Sub

'===================================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If y<20 Then
		lgSpdHdrClicked = 1
	End If

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'===================================================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
	If y<20 Then
		lgSpdHdrClicked = 1
	End If

	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub
'===================================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'===================================================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'===================================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
	Call GetSpreadColumnPos("A")
End Sub
'===================================================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	Call GetSpreadColumnPos("B")
End Sub
'===================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

End Sub
'===================================================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName

 	If Row <= 0 Then
		Exit Sub
 	End If

  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If
End Sub
'===================================================================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	
	With frm1.vspdData1
		.Row = Row
		.Col = Col
		Select case Col
		case C2_ISSUE_TYPE_NM
			intIndex = .Value
			.Col = C2_ISSUE_TYPE
			.Value = intIndex
		End Select
	End With
End Sub
'===================================================================================================================================
Sub FncSplitColumn()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub
'===================================================================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()

End Sub
'===================================================================================================================================
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
	Dim iActiveRow
	Dim iConvActiveRow
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim lRow
	Dim index
	Dim strFlag
	Dim strParentRowNo
	Dim i
	ggoSpread.Source = gActiveSpdSheet

	If gActiveSpdSheet.ID = "A" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("A")
		Call ggoSpread.ReOrderingSpreadData
	ElseIf gActiveSpdSheet.ID = "B" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("B")
		Call InitComboBox()
		frm1.vspdData1.Redraw = False
		Call ggoSpread.ReOrderingSpreadData("F")
	End If
End Sub
'===================================================================================================================================
Sub vspdData_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
	Row = 0: Row2 = -1: NewRow = 0
	ggoSpread.SwapRange Col, Row, Col2, Row2, NewCol, NewRow, Cancel
End Sub
'===================================================================================================================================
Sub vspdData_GotFocus()
	ggoSpread.Source = frm1.vspdData
End Sub
'===================================================================================================================================
Sub vspdData1_GotFocus()
	ggoSpread.Source = frm1.vspdData1
End Sub
'===================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
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
	End If  
End Sub
'===================================================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
  	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey2 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery2(frm1.vspdData.ActiveRow) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
'===================================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Select Case Col
		Case C1_ITEM_POP
			Call OpenItemCd2(Row)
		Case C1_UNIT_POP
			call OpenUnit(Row)
	End Select
End Sub
'===================================================================================================================================
Function FncSave()
	FncSave = False
	
	Dim IntRetCD

	Err.Clear

	If CheckRunningBizProcess = True Then
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = false  Then
      IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
      Exit Function
  End If
	
	If Not chkField(Document, "1") Then
	   		Exit Function
	End If

	If Not chkField(Document, "2") Then
	   		Exit Function
	End If

	If DbSave = False then
		Exit Function
	End If

	FncSave = True
End Function
'===================================================================================================================================
 Function FncQuery()
	Dim IntRetCD

	FncQuery = False														'⊙: Processing is NG

	On Error Resume Next
	Err.Clear															   '☜: Protect system from crashing

	ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = true Then
      IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
      if IntRetCD=vbNo Then Exit Function
  End If

	Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	Call InitVariables

	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If

	If (UniConvDateToYYYYMMDD(frm1.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(frm1.txtToDt.text,Parent.gDateFormat,"")) and Trim(frm1.txtFrDt.text)<>"" and Trim(frm1.txtToDt.text)<>"" then
		Call DisplayMsgBox("17a003", "X","입고일", "X")
		Exit Function
	End if

	If Dbquery = False then Exit Function
	FncQuery = True																'⊙: Processing is OK

End Function
'===================================================================================================================================
Function FncNew()
	Dim IntRetCD

	FncNew = False														  '⊙: Processing is NG

	Err.Clear

	ggoSpread.Source = frm1.vspdData1

	Call ggoOper.ClearField(Document, "1")										 '⊙: Clear Condition Field
	Call ggoOper.ClearField(Document, "Q")

	Call InitVariables

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
	ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData

	Call SetDefaultVal
	Call SetToolbar("11100000000000")

	FncNew = True														   '⊙: Processing is OK

End Function
'===================================================================================================================================
Function FncDeleteRow()		'###그리드 컨버전 주의부분###
	FncDeleteRow = false

	Dim lDelRows
	Dim iDelRowCnt, i
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iparentrow

	select case gActiveSpdSheet.id
	case "A"

		If frm1.vspdData.MaxRows < 1 then
			Exit function
		End if
		
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
		
	case "B"

		If frm1.vspdData1.MaxRows < 1 then
			Exit function
		End if
		
    With Frm1.vspdData1 
    	.focus
    	ggoSpread.Source = frm1.vspdData1 
    	lDelRows = ggoSpread.DeleteRow
    End With

		Call vspdData_Change(0, frm1.vspdData.Row)

	end select
	FncDeleteRow = true
End Function
'===================================================================================================================================
Function FncDelete()
	Dim lDelRows
	Dim iDelRowCnt, i
	if frm1.vspdData.Maxrows < 1 then exit function

	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		lDelRows = ggoSpread.DeleteRow

	End With
End Function
'===================================================================================================================================
Function FncCopy()
	FncCopy = false

	Dim lRow, lRow2
	Dim lngRangeFrom, lngRangeTo
	Dim strFlag
	Dim i, k

	With frm1

		If CheckDataExist = False Then
			Exit function
		End If

		.vspdData1.ReDraw = False

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.CopyRow

		.vspdData1.ReDraw = True
		.vspdData1.focus
	End With
	FncCopy = true
	Set gActiveSpdSheet = frm1.vspdData1

End Function
'===================================================================================================================================
Function FncInsertRow(ByVal pvRowCnt)	'###그리드 컨버전 주의부분###
	FncInsertRow = false

	On Error Resume Next

	Dim lRow
	Dim lRow2
	Dim lconvRow
	Dim strMark
	Dim iInsertRow
	Dim IntRetCD
	Dim imRow
	Dim strInspUnitIndctnCd
	Dim iParentRowNo,iparentrow, iStrPrNo, iStrPoNo, iStrPoSeq
	Dim iStrPoQty, iStrPoUnit, iStrPoDt, iStrRcptQty, iStrTracking_no
	Dim iStrPlantCd, iStrSpplCd

	select case gActiveSpdSheet.id
	case "A"

		with frm1	
			
			.vspdData.ReDraw = False
	
			Dim i
			Dim iCnt
			for i=1 To .vspdData.MAxRows
					if GetSpreadText(frm1.vspdData,0,i,"X","X") = ggoSpread.InsertFlag Then
						Call DisplayMsgBox("17a012", "X", "이 프로그램 ", "한 건이상 처리")
						exit function
					End if
			next
	
			If IsNumeric(Trim(pvRowCnt)) Then
				imRow = CInt(pvRowCnt)
			Else
				imRow = AskSpdSheetAddRowCount()
				If imRow = "" Then
					Exit Function
				End If
			End If
	
			.vspdData.focus
			ggoSpread.Source = .vspdData
			ggoSpread.InsertRow .vspdData.ActiveRow, imRow			
			call frm1.vspdData.SetText(C1_RCPT_DT,.vspdData.ActiveRow,EndDate)
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
			.vspdData.ReDraw = True
		End with
		
		
	case "B"

		with frm1
			If .vspdData.MaxRows <= 0 Then
				Exit Function
			End If
	
			.vspdData1.ReDraw = False
	
			If IsNumeric(Trim(pvRowCnt)) Then
				imRow = CInt(pvRowCnt)
			Else
				imRow = AskSpdSheetAddRowCount()
				If imRow = "" Then
					Exit Function
				End If
			End If
	
			.vspdData1.focus
			ggoSpread.Source = .vspdData1
			ggoSpread.InsertRow .vspdData1.ActiveRow, imRow			
			call frm1.vspdData1.SetText(C2_ISSUE_DT,.vspdData1.ActiveRow,EndDate)
	    SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1
			.vspdData1.ReDraw = True
		End with

	End Select



	FncInsertRow = true

	Set gActiveSpdSheet = document.activeElement
End Function
'===================================================================================================================================
Function FncCancel()
	FncCancel = false
	Dim lRow
	Dim i,k,iCnt
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iActiveRow
	Dim iConvActiveRow
	Dim strFlag

	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End If

	if isEmpty(gActiveSpdSheet) then
		Set gActiveSpdSheet = frm1.vspdData1
	end if

	select case gActiveSpdSheet.id
	case "A"
		with frm1	
			.vspdData.ReDraw = False	
			.vspdData.focus
			ggoSpread.Source = .vspdData
			ggoSpread.EditUndo
			.vspdData.ReDraw = True
		End with
	case "B"
		with frm1	
			.vspdData1.ReDraw = False	
			.vspdData1.focus
			ggoSpread.Source = .vspdData1
			ggoSpread.EditUndo
			.vspdData1.ReDraw = True
		End with
	End Select
	
	FncCancel = true
End Function
'===================================================================================================================================
Function FncPrint()
	FncPrint = False
	Call Parent.FncPrint()
	FncPrint = True
End Function
'===================================================================================================================================
Function FncExcel()
	FncExcel = False
 	Call parent.FncExport(Parent.C_MULTI)
 	FncExcel = True
End Function
'===================================================================================================================================
Function FncFind()
	FncFind = False
	Call parent.FncFind(Parent.C_MULTI, False)										 '☜:화면 유형, Tab 유무 
	FncFind = True
End Function
'===================================================================================================================================
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		  '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function
'===================================================================================================================================
Function DbQuery()
	Dim LngLastRow
	Dim LngMaxRow
	Dim LngRow
	Dim strTemp
	Dim StrNextKey
	Dim pP21018		 'As New P21018ListIndReqSvr

	DbQuery = False

	If LayerShowHide(1) = False Then Exit Function

	Err.Clear															   '☜: Protect system from crashing

	Dim strVal

	With frm1

		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
			if frm1.rdoBal1.checked Then
				strVal = strVal & "&rdoBal="
			elseIf frm1.rdoBal2.checked Then
				strVal = strVal & "&rdoBal=Y"
			else
				strVal = strVal & "&rdoBal=N"
			End if
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
			if frm1.rdoBal1.checked Then
				strVal = strVal & "&rdoBal="
			elseIf frm1.rdoBal2.checked Then
				strVal = strVal & "&rdoBal=Y"
			else
				strVal = strVal & "&rdoBal=N"
			End if
		End If

		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey					  '☜: Next key tag
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	End With

	DbQuery = True

End Function
'===================================================================================================================================
Function 	FncQuery2(Row)
	FncQuery2 = False
	frm1.vspdData1.MaxRows = 0
	call DbQuery2(Row)
	FncQuery2 = true
End Function

Function DbQuery2(Row)
	Dim LngLastRow
	Dim LngMaxRow
	Dim LngRow
	Dim strTemp
	Dim StrNextKey
	Dim pP21018		 'As New P21018ListIndReqSvr

	DbQuery2 = False

	If LayerShowHide(1) = False Then Exit Function

	Err.Clear															   '☜: Protect system from crashing

	Dim strVal
	
	With frm1

		strVal = BIZ_PGM_ID & "?txtMode=DbQuery2"
		strVal = strVal & "&txtRcptNo=" & GetSpreadText(.vspdData,C1_RCPT_NO,.vspdData.ActiveRow,"X","X")
		strVal = strVal & "&lgStrPrevKey2="   & lgStrPrevKey2					  '☜: Next key tag
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	End With

	DbQuery2 = True

End Function
'===================================================================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	Dim i, lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows
	'-----------------------
	'Reset variables area
	'-----------------------

  lgIntFlgMode = parent.OPMD_UMODE    
	Call ggoOper.LockField(Document, "N")									'⊙: This function lock the suitable field
	Call SetSpreadLock("A")
	call FncQuery2(1)
	call SetToolbar("1100111100001111")										'⊙: 버튼 툴바 제어 
	DbQueryOk = true

End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'===================================================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel
	
    DbSave = False                                                          
    
     If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
				.txtHDcmd.value				= GetSpreadText(.vspdData,0,.vspdData.ActiveRow,"X","X")
    		.txtRCPT_NO.value 		= GetSpreadText(.vspdData,C1_RCPT_NO,.vspdData.ActiveRow,"X","X")
    		.txtRCPT_DT.value 		= GetSpreadText(.vspdData,C1_RCPT_DT,.vspdData.ActiveRow,"X","X")
    		.txtPLANT_CD.value 		= .txtPlantCd.value
    		.txtITEM_CD.value 		= GetSpreadText(.vspdData,C1_ITEM_CD,.vspdData.ActiveRow,"X","X")
    		.txtUNIT.value 				= GetSpreadText(.vspdData,C1_UNIT,.vspdData.ActiveRow,"X","X")
    		.txtRCPT_QTY.value 		= UniConvNum(GetSpreadText(.vspdData,C1_RCPT_QTY,.vspdData.ActiveRow,"X","X"),0)
    		.txtPRICE.value	 			= UniConvNum(GetSpreadText(.vspdData,C1_PRICE,.vspdData.ActiveRow,"X","X"),0)
    		.txtAMT.value 				= UniConvNum(GetSpreadText(.vspdData,C1_AMT,.vspdData.ActiveRow,"X","X"),0)
    		.txtRCPT_DOC_NO.value = GetSpreadText(.vspdData,C1_RCPT_DOC_NO,.vspdData.ActiveRow,"X","X")
    		.txtCLOSE_YN.value 		= GetSpreadText(.vspdData,C1_CLOSE_YN,.vspdData.ActiveRow,"X","X")
    		
       For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
        
           Select Case .vspdData1.Text
        
               Case ggoSpread.InsertFlag                                      '☜: Insert
                    strVal = strVal & "C" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_SEQ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_DT,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_TYPE,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_QTY,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_DOC_NO,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
                    strVal = strVal & "U" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_SEQ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_DT,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_TYPE,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_QTY,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData1,C2_ISSUE_DOC_NO,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                    strDel = strDel & "D" & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData1,C2_SEQ,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData1,C2_ISSUE_DT,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData1,C2_ISSUE_TYPE,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData1,C2_ISSUE_QTY,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData1,C2_ISSUE_DOC_NO,lRow,"X","X") & parent.gColSep
                    strDel = strDel & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

     .txtMode.value        = parent.UID_M0002
     .txtFlgMode.value     = lgIntFlgMode       
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function
'===================================================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
	Call MainQuery()
End Function
'===================================================================================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>원료과세수불등록</font></td>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant() ">
														   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=18 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=34 tag="14"></TD>
								</TR>
								<tr>
									<TD CLASS="TD5" NOWRAP>입고일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입고일 NAME="txtFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="12N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입고일 NAME="txtToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="12N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
									</TD>
									<TD CLASS="TD5" NOWRAP>잔량유무</TD>
									<TD CLASS="TD6" NOWRAP>
											<input type=radio CLASS="RADIO" name="rdoBal" id="rdoBal1" value="" tag = "11" checked>
												<label for="rdoBal1">전체</label>&nbsp;
											<input type=radio CLASS="RADIO" name="rdoBal" id="rdoBal2" value="Y" tag = "11">
												<label for="rdoBal2">잔량有</label>&nbsp;
											<input type=radio CLASS="RADIO" name="rdoBal" id="rdoBal3" value="N" tag = "11">
												<label for="rdoBal3">잔량無</label>&nbsp;									
										</TD>
								</tr>
							</TABLE>
						</FIELDSET>
					</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="A"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="B"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtHDcmd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRCPT_NO" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRCPT_DT" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPLANT_CD" tag="24">
<INPUT TYPE=HIDDEN NAME="txtITEM_CD" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUNIT" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRCPT_QTY" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPRICE" tag="24">
<INPUT TYPE=HIDDEN NAME="txtAMT" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRCPT_DOC_NO" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCLOSE_YN" tag="24">


</FORM>

	<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>

</BODY>
</HTML>
