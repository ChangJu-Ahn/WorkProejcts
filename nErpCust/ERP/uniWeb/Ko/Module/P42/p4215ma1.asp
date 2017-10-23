
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Master Production Scheduling
'*  3. Program ID			: p4215ma1.asp
'*  4. Program Name			: 오더 Document 조회 
'*  5. Program Desc			: 오더 Document 조회 
'*  6. Business ASP List	: p4215mb1.asp (Query Order Document)
'*  7. Modified date(First)	: 
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: Hong, EunSook
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				: 
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID = "p4215mb1.asp"			'☆: Head Query 비지니스 로직 ASP명 

Const C_SHEETMAXROWS = 30

Dim C_OprNo				'= 1
Dim C_WcCd				'= 2
Dim C_WcNm				'= 3
Dim C_JobCd				'= 4
Dim C_JobNm				'= 5
Dim C_InsideFlg			'= 6
Dim C_OrderStatus		'= 7
Dim C_DtlPlanStartDt	'= 8
Dim C_DtlPlanComptDt	'= 9
Dim C_DtlReleaseDt		'= 10
Dim C_GoodQty			'= 11
Dim C_BadQty			'= 12
Dim C_ProdtOrderUnit	'= 13
Dim C_Document			'= 14
Dim C_ItemCd			'= 15
Dim C_ItemNm			'= 16
Dim C_ProdtOrderQty		'= 17
Dim C_PlanStartDt		'= 18
Dim C_PlanComptDt		'= 19
Dim C_Routing			'= 20
Dim C_TrackingNo		'= 21
Dim C_orgDocument		'= 22

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim ihGridCnt                     'hidden Grid Row Count
Dim intItemCnt                    'hidden Grid Row Count
Dim IsOpenPop						' Popup


'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
     
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
    lgSortKey = 1

End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
        
    'Dim sList
    Call InitSpreadPosVariables()
    
    With frm1
    
    'sList = "1" & Chr(11) & "2" & Chr(12) & "3"
       
    ggoSpread.Source = .vspdData
    ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
    
    .vspdData.ReDraw = False
    
    .vspdData.MaxCols = C_orgDocument + 1
    .vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_OprNo,			"공정", 10
    ggoSpread.SSSetEdit		C_WcCd,				"작업장", 10
    ggoSpread.SSSetEdit		C_WcNm,				"작업장명", 20
    ggoSpread.SSSetEdit		C_JobCd,			"공정작업", 10
    ggoSpread.SSSetEdit		C_JobNm,			"공정작업명", 20
    ggoSpread.SSSetEdit 	C_InsideFlg,		"공정타입",10
    ggoSpread.SSSetEdit		C_OrderStatus,		"지시상태", 10
    ggoSpread.SSSetDate		C_DtlPlanStartDt,	"착수예정일", 11, 2, parent.gDateFormat    
    ggoSpread.SSSetDate		C_DtlPlanComptDt,	"완료예정일", 11, 2, parent.gDateFormat    
    ggoSpread.SSSetDate		C_DtlReleaseDt,		"작업지시일", 11, 2, parent.gDateFormat    
	ggoSpread.SSSetFloat	C_GoodQty,			"양품수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"    
	ggoSpread.SSSetFloat	C_BadQty,			"불량수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
    ggoSpread.SSSetEdit		C_ProdtOrderUnit,	"오더단위", 8,,,3,2
    ggoSpread.SSSetEdit		C_Document,			" ", 100,,,100
    ggoSpread.SSSetEdit		C_ItemCd,			"품목", 18
    ggoSpread.SSSetEdit		C_ItemNm,			"품목명 ", 25
    ggoSpread.SSSetEdit		C_ProdtOrderQty,	"오더수량", 10
    ggoSpread.SSSetEdit		C_PlanStartDt,		"착수예정일", 10
    ggoSpread.SSSetEdit		C_PlanComptDt,		"완료예정일", 10
    ggoSpread.SSSetEdit		C_Routing,			"라우팅", 10
    ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No.", 25
    ggoSpread.SSSetEdit		C_orgDocument,		" ", 100,,,100

    ihGridCnt = 0               'Hidden Counter
    intItemCnt = 0
    ggoSpread.Source = .vspdData
	
	'Call ggoSpread.MakePairsColumn(C_ItemCode,C_ItemSpec)
 	Call ggoSpread.SSSetColHidden(C_Document ,C_orgDocument , True)
 	Call ggoSpread.SSSetColHidden(.vspdData.MaxCols ,.vspdData.MaxCols , True)
	
	.vspdData.ReDraw = true
	
    ggoSpread.SSSetSplit2(2) 

    End With
    
    Call SetSpreadLock()
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData 
    
    .Redraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.SSSetProtected C_OprNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcNm, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_JobCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_JobNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_InsideFlg, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_OrderStatus, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DtlPlanStartDt, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_DtlPlanComptDt,pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DtlReleaseDt,pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_GoodQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BadQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ProdtOrderUnit, pvStartRow, pvEndRow
    
    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
    .EditMode = True
    
    .Redraw = True
    
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	
	C_OprNo				= 1
	C_WcCd				= 2
	C_WcNm				= 3
	C_JobCd				= 4
	C_JobNm				= 5
	C_InsideFlg			= 6
	C_OrderStatus		= 7
	C_DtlPlanStartDt	= 8
	C_DtlPlanComptDt	= 9
	C_DtlReleaseDt		= 10
	C_GoodQty			= 11
	C_BadQty			= 12
	C_ProdtOrderUnit	= 13
	C_Document			= 14
	C_ItemCd			= 15
	C_ItemNm			= 16
	C_ProdtOrderQty		= 17
	C_PlanStartDt		= 18
	C_PlanComptDt		= 19
	C_Routing			= 20
	C_TrackingNo		= 21
	C_orgDocument		= 22

End Sub


'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 		C_OprNo				= iCurColumnPos(1)
		C_WcCd				= iCurColumnPos(2)
		C_WcNm				= iCurColumnPos(3)
		C_JobCd				= iCurColumnPos(4)
		C_JobNm				= iCurColumnPos(5)
		C_InsideFlg			= iCurColumnPos(6)
		C_OrderStatus		= iCurColumnPos(7)
		C_DtlPlanStartDt	= iCurColumnPos(8)
		C_DtlPlanComptDt	= iCurColumnPos(9)
		C_DtlReleaseDt		= iCurColumnPos(10)
		C_GoodQty			= iCurColumnPos(11)
		C_BadQty			= iCurColumnPos(12)
		C_ProdtOrderUnit	= iCurColumnPos(13)
		C_Document			= iCurColumnPos(14)
		C_ItemCd			= iCurColumnPos(15)
		C_ItemNm			= iCurColumnPos(16)
		C_ProdtOrderQty		= iCurColumnPos(17)
		C_PlanStartDt		= iCurColumnPos(18)
		C_PlanComptDt		= iCurColumnPos(19)
		C_Routing			= iCurColumnPos(20)
		C_TrackingNo		= iCurColumnPos(21)
		C_orgDocument		= iCurColumnPos(22)

 	End Select
 
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
    
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "공장"						' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"						' Field명(0)
    arrField(1) = "PLANT_NM"						' Field명(1)
    
    arrHeader(0) = "공장"						' Header명(0)
    arrHeader(1) = "공장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "OP"
	arrParam(4) = "RL"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  OpenOprCd()  -------------------------------------------------
'	Name : OpenOprCd()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprCd()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	If frm1.txtProdOrderNo.value = "" Then
		Call DisplayMsgBox("971012","X" , "제조오더번호","X")
		frm1.txtProdOrderNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	iCalledAspName = AskPRAspName("P4112PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4112PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Or UCase(frm1.txtOprCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdOrderNo.value
	arrParam(2) = "" 'frm1.txtOprCd.value

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetOprCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprCd.focus
	
End Function

'------------------------------------------  OpenProdRef()  -------------------------------------------------
'	Name : OpenProdRef()
'	Description : Production Reference
'---------------------------------------------------------------------------------------------------------
Function OpenProdRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4411RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4411RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
		
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)			<%'☆: 조회 조건 데이타 %>
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		<%'☜: 조회 조건 데이타 %>
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  -------------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Issue Reference
'--------------------------------------------------------------------------------------------------------- 
Function OpenRcptRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)			<%'☆: 조회 조건 데이타 %>
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		<%'☜: 조회 조건 데이타 %>

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenConsumRef()  -------------------------------------------------
'	Name : OpenConsumRef()
'	Description : Part Consumption Reference
'--------------------------------------------------------------------------------------------------------- 
Function OpenConsumRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4412RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4412RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)			<%'☆: 조회 조건 데이타 %>
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		<%'☜: 조회 조건 데이타 %>

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function


'==========================================  2.4.3 Set Return Value()  =============================================
'	Name : Set Return Value()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 


'================================== ReleaseButton() ===================================
' Function Name : ReleaseButton
' Function Desc : 저장이 성공하고나 실패했을경우 버튼을 다시 활성화시킨다.
'========================================================================================

'Sub ReleaseButton(pButton)
'	pButton.disabled = False
'End Sub


'================================== ProtectButton() ===================================
' Function Name : ProtectButton
' Function Desc : 저장시 버튼을 disalbe시킨다.
'========================================================================================

'Sub ProtectButton(pButton)
'	pButton.disabled = True
'End Sub

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
    frm1.txtPlantCd.Value    = arrRet(0)		
    frm1.txtPlantNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetOprCd()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOprCd(byval arrRet)
	frm1.txtOprCd.Value    = arrRet(0)		
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables

	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

End Sub

'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================

Sub vspdData_EditChange(ByVal Col , ByVal Row )
                
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
 		
 	End If
	
	'------ Developer Coding part (Start)
 	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_Document
		frm1.txtDocument.Value = replace(.Text,chr(7), chr(13) &chr(10))
	End With
	'------ Developer Coding part (End)
	
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub
 
'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspddata_KeyPress(index , KeyAscii )
 
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()

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

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	 '----------  Coding part  -------------------------------------------------------------

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

 
'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub 


'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
   
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					'⊙: Clear Contents  Field 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData  
    Call InitVariables	
    '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkfield(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If															'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
   
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================

Function FncPaste() 
     ggoSpread.SpreadPaste
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing   
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
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)											'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                        '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function


'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 

End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'******************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    
    DbQuery = False                                                         				'⊙: Processing is NG
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtOprCd=" & Trim(frm1.hOprCd.value)				'☆: 조회 조건 데이타 
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtOprCd=" & Trim(frm1.txtOprCd.value)				'☆: 조회 조건 데이타	
	End If
    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()															'☆: 조회 성공후 실행로직 
	
	Call SetToolBar("11000000000111")											'⊙: 버튼 툴바 제어 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
        
	 With frm1.vspdData

		.Row = 1

		.Col = C_Document
		frm1.txtDocument.Value = replace(.Text,chr(7),chr(13) &chr(10))
		.Col = C_ItemCd
		frm1.txtItemCd.Value = .Text
		.Col = C_ItemNm
		frm1.txtItemNm.Value = .Text
		.Col = C_ProdtOrderQty
		If .Text = "" or isnull(.text) or isempty(.text) then
			frm1.txtOrderQty.Value = 0
		Else
			frm1.txtOrderQty.Value = .Text
		End if
		.Col = C_PlanStartDt
		frm1.txtPlanStartDt.Text = .Text
		.Col = C_PlanComptDt
		frm1.txtPlanComptDt.Text = .Text
		.Col = C_Routing
		frm1.txtRouting.value = .Text 
		.Col = C_TrackingNo
		frm1.txtTrackingNo.value = .Text 
	
    End With

End Function


'========================================================================================
' Function Name : uFncDtlQuery
' Function Desc : This function is data query and display
'========================================================================================

Function uFncDtlQuery(Byval lRow) 
    

End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

'	Call InitVariables
'	Call FncQuery()

End Function

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



'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>오더Document조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenProdRef()">실적내역</A> | <A href="vbscript:OpenRcptRef()">입고내역</A> | <A href="vbscript:OpenConsumRef()">부품소비내역</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
									<TD CLASS=TD5 NOWRAP>공정</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOprCd" SIZE=10 MAXLENGTH=3 tag="11xxxU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprCd()"></TD>
								</TR>													
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD><!-- 첫번째 탭 내용 -->
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>품목</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="24" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>오더수량</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/p4215ma1_I793539428_txtOrderQty.js'></script>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>작업계획일자</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/p4215ma1_I932335698_txtPlanStartDt.js'></script>
								&nbsp;~&nbsp;
								<script language =javascript src='./js/p4215ma1_I733530098_txtPlanComptDt.js'></script>
							</TD>								
							<TD CLASS=TD5 NOWRAP>라우팅</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRouting" SIZE=10 MAXLENGTH=9 tag="24" ALT="라우팅"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="24" ALT="Tracking No."></TD>
							<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
						</TR>	
						<TR>
							<TD HEIGHT="60%" colspan=4>
								<script language =javascript src='./js/p4215ma1_I141483849_vspdData.js'></script>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 HEIGHT="40%" NOWRAP>Document</TD>
                         	<TD CLASS=TD656 HEIGHT="40%" valign="middle" colspan=3>
                         		<TEXTAREA  NAME="txtDocument" tag="24xxxU" rows=6 cols=80 wrap="soft" ALT="Document"></TEXTAREA>
                         	</TD>
						</TR>					
					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOprCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
