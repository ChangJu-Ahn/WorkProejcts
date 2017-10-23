
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p4111pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Production Order Reference ASP											*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2002/12/10																*
'*  9. Modifier (First)     : Kim GyoungDon																*
'* 10. Modifier (Last)      : RYU SUNG WON																*
'* 11. Comment              :																			*
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)  
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin		                *
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--####################################################################################################
'#						1. 선 언 부																		#
'#####################################################################################################-->

<!--********************************************  1.1 Inc 선언  *****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--============================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--============================================  1.1.2 공통 Include  ===================================
'=====================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'********************************************  1.2 Global 변수/상수 선언  *******************************
'*	Description : 1. Constant는 반드시 대문자 표기														*
'********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p4111pb1.asp"			'☆: 비지니스 로직 ASP명 
'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================
Dim C_ProdOrderNo
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_PlanStartDt
Dim C_PlanEndDt
Dim C_OrderStatus
Dim C_OrderStatusDesc
Dim C_OrderQty
	
Dim C_ProdQtyInOrderUnit
Dim C_GoodQtyInOrderUnit
Dim C_RcptQtyInOrderUnit
Dim C_OrderUnit

Dim C_OrderQtyInBaseUnit
Dim C_ProdQtyInBaseUnit
Dim C_GoodQtyInBaseUnit
Dim C_RcptQtyInBaseUnit
Dim C_BaseUnit

Dim C_RoutingNo
Dim C_SLCD
Dim C_SLNM
Dim C_ReWorkFlag
Dim C_OrderType
Dim C_OrderTypeDesc
Dim C_TrackingNo
Dim C_PlanOrderNo
	
'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
	
'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
Dim arrReturn
Dim lgPlantCD
Dim strFromStatus
Dim strToStatus
Dim strThirdStatus
Dim IsOpenPop
Dim arrParent
Dim IsFormLoaded
	
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable값 정의  ============================
'========================================================================================================
'----------------  공통 Global 변수값 정의  -------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						2. Function 부																	#
'#																										#
'#	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기술					#
'#	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.							#
'#						 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)			#
'########################################################################################################
'*******************************************  2.1 변수 초기화 함수  *************************************
'*	기능: 변수초기화																					*
'*	Description : Global변수 처리, 변수초기화 등의 작업을 한다.											*
'********************************************************************************************************
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ProdOrderNo		= 1
	C_ItemCd			= 2
	C_ItemNm			= 3
	C_Spec				= 4
	C_PlanStartDt		= 5
	C_PlanEndDt			= 6
	C_OrderStatus		= 7
	C_OrderStatusDesc	= 8
	C_OrderQty			= 9
		
	C_ProdQtyInOrderUnit= 10
	C_GoodQtyInOrderUnit= 11
	C_RcptQtyInOrderUnit= 12
	C_OrderUnit			= 13

	C_OrderQtyInBaseUnit= 14
	C_ProdQtyInBaseUnit	= 15
	C_GoodQtyInBaseUnit	= 16
	C_RcptQtyInBaseUnit	= 17
	C_BaseUnit			= 18

	C_RoutingNo			= 19
	C_SLCD				= 20
	C_SLNM				= 21
	C_ReWorkFlag		= 22
	C_OrderType			= 23
	C_OrderTypeDesc		= 24
	C_TrackingNo		= 25
	C_PlanOrderNo		= 26
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
	vspdData.MaxRows = 0
	lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
	lgStrPrevKey = ""										'initializes Previous Key		
    lgIntFlgMode = PopupParent.OPMD_CMODE								'Indicates that current mode is Create mode	

	Self.Returnvalue = Array("")
End Function

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter를 Variable에 Setting한다.											=
'========================================================================================================
Function InitSetting()

	Dim ArgArray						<%'Arguments로 넘겨받은 Array%>
	'ArrParent = window.dialogArguments
	
	ArgArray  = ArrParent(1)

	lgPlantCD = ArgArray(0)
	txtFromDt.Text = ArgArray(1)
	txtToDt.Text = ArgArray(2)
	strFromStatus = ArgArray(3)
	strToStatus = ArgArray(4)
	If (ArgArray(3) = ArgArray(4)) and (ArgArray(3) <> "" and ArgArray(4) <> "") Then
		cboOrderStatus.value = ArgArray(3)
	End If
	
	If Trim(strToStatus) <> "" Then

        If Len(Trim(strToStatus)) > 2 Then

            strToStatus = Mid(Trim(ArgArray(4)),1,2)

            strThirdStatus = Mid(Trim(ArgArray(4)),3,2)

        Else

            strThirdStatus = Trim(strToStatus)

        End If

    End If

	txtProdOrderNo.value = ArgArray(5)
	txtTrackingNo.value = ArgArray(6)
	txtItemCd.value = ArgArray(7)
	cboOrderType.value = ArgArray(8)

End Function

'==========================================   2.1.3 InitComboBox()  =====================================
'=	Name : InitComboBox()																				=
'=	Description : ComboBox에 Value를 Setting한다.														=
'========================================================================================================
Sub InitComboBox()

    Dim iCodeArr 
    Dim iNameArr

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call SetCombo2(cboOrderType, iCodeArr, iNameArr, Chr(11))
    ggoSpread.Source = vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderTypeDesc
	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call SetCombo2(cboOrderStatus, iCodeArr, iNameArr, Chr(11))
    ggoSpread.Source = vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc
	
	cboOrderType.value = ""
	cboOrderStatus.value = "" 	 
    
End Sub

'==========================================  2.1.4 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display in Spread(s)
'========================================================================================================= 
Sub InitSpreadComboBox()
	Dim iCodeArr 
    Dim iNameArr

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    ggoSpread.Source = vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderTypeDesc
	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    ggoSpread.Source = vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc
	
	cboOrderType.value = ""
	cboOrderStatus.value = "" 
End Sub
'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex
	
	With vspdData
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_OrderStatus
			intIndex = .value
			.Col = C_OrderStatusDesc
			.value = intindex
			.Row = intRow
			.col = C_OrderType
			intIndex = .value
			.Col = C_OrderTypeDesc
			.value = intindex			
		Next	
	End With
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================%>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE","PA") %>
	<% Call loadBNumericFormatA("Q", "P", "NOCOOKIE","PA") %>
End Sub
	
'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	
	vspdData.MaxCols = C_PlanOrderNo + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit		C_ProdOrderNo,		"오더번호", 18,,,18,2
	ggoSpread.SSSetEdit		C_ItemCd,			"품목", 18
	ggoSpread.SSSetEdit		C_ItemNm,			"품목명", 25
	ggoSpread.SSSetEdit		C_Spec,				"규격", 25
	ggoSpread.SSSetDate		C_PlanStartDt,		"착수예정일", 10, 2, PopupParent.gDateFormat
	ggoSpread.SSSetDate		C_PlanEndDt,		"완료예정일", 10, 2, PopupParent.gDateFormat
	ggoSpread.SSSetCombo	C_OrderStatus,		"지시상태", 10
	ggoSpread.SSSetCombo	C_OrderStatusDesc,	"지시상태", 10
	ggoSpread.SSSetFloat	C_OrderQty,			"오더수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "실적수량", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit, "양품수량", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"		
	ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit, "입고수량", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_OrderUnit,		"오더단위", 8

	ggoSpread.SSSetFloat	C_OrderQtyInBaseUnit, "기준수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit, "실적수량", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit, "양품수량", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"		
	ggoSpread.SSSetFloat	C_RcptQtyInBaseUnit, "입고수량", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"		
    ggoSpread.SSSetEdit		C_BaseUnit,			"기준단위", 8

	ggoSpread.SSSetEdit		C_RoutingNo,		"라우팅", 10,,,,2
	ggoSpread.SSSetEdit		C_SLCD,				"창고", 10,,,,2
	ggoSpread.SSSetEdit		C_SLNM,				"창고명", 20
	ggoSpread.SSSetEdit		C_ReWorkFlag,		"재작업", 6
	ggoSpread.SSSetCombo	C_OrderType,		"지시구분", 10
	ggoSpread.SSSetCombo	C_OrderTypeDesc,	"지시구분", 10
	ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No.", 25,,,,2
	ggoSpread.SSSetEdit		C_PlanOrderNo,		"계획오더 No.", 18,,,,2

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_OrderStatus,C_OrderStatus, True)
	Call ggoSpread.SSSetColHidden(C_OrderType,C_OrderType, True)
	
	ggoSpread.SSSetSplit2(1)
	vspdData.ReDraw = True
	Call SetSpreadLock()
End Sub
	
'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ProdOrderNo		= iCurColumnPos(1)
			C_ItemCd			= iCurColumnPos(2)
			C_ItemNm			= iCurColumnPos(3)
			C_Spec				= iCurColumnPos(4)
			C_PlanStartDt		= iCurColumnPos(5)
			C_PlanEndDt			= iCurColumnPos(6)
			C_OrderStatus		= iCurColumnPos(7)
			C_OrderStatusDesc	= iCurColumnPos(8)
			C_OrderQty			= iCurColumnPos(9)
				
			C_ProdQtyInOrderUnit= iCurColumnPos(10)
			C_GoodQtyInOrderUnit= iCurColumnPos(11)
			C_RcptQtyInOrderUnit= iCurColumnPos(12)
			C_OrderUnit			= iCurColumnPos(13)

			C_OrderQtyInBaseUnit= iCurColumnPos(14)
			C_ProdQtyInBaseUnit	= iCurColumnPos(15)
			C_GoodQtyInBaseUnit	= iCurColumnPos(16)
			C_RcptQtyInBaseUnit	= iCurColumnPos(17)
			C_BaseUnit			= iCurColumnPos(18)

			C_RoutingNo			= iCurColumnPos(19)
			C_SLCD				= iCurColumnPos(20)
			C_SLNM				= iCurColumnPos(21)
			C_ReWorkFlag		= iCurColumnPos(22)
			C_OrderType			= iCurColumnPos(23)
			C_OrderTypeDesc		= iCurColumnPos(24)
			C_TrackingNo		= iCurColumnPos(25)
			C_PlanOrderNo		= iCurColumnPos(26)
    End Select    
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
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call initData(1,1)
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If CheckRunningBizProcess = True Then Exit Sub
    If OldLeft <> NewLeft Then Exit Sub
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then Exit Sub
		End If
    End if    
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intRowCnt
	Dim intColCnt
	Dim intSelCnt

	If vspdData.MaxRows > 0 Then
		intSelCnt = 0
		Redim arrReturn(0)
		
		vspdData.Row = vspdData.ActiveRow

		If vspdData.SelModeSelected = True Then
			vspdData.Col = C_ProdOrderNo
			arrReturn(0) = vspdData.Text
		End If

		Self.Returnvalue = arrReturn
	End If		
	Self.Close()
End Function
	
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub txtFromDt_KeyDown(keycode, shift)
	If keycode=27 Then
 		Call Self.Close()
		Exit Sub
	ElseIf Keycode = 13 Then
		Call FncQuery()
	End If
End Sub	

Sub txtToDt_KeyDown(keycode, shift)
	If keycode=27 Then
 		Call Self.Close()
		Exit Sub
	ElseIf Keycode = 13 Then
		Call FncQuery()
	End If
End Sub	

Sub vspdData_KeyPress(keyAscii)
	If keyAscii=13 and vspdData.ActiveRow > 0 Then
 		Call OkClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub	


'*******************************************  2.4 POP-UP 처리함수  **************************************
'*	기능: POP-UP																						*
'*	Description : POP-UP Call하는 함수 및 Return Value setting 처리										*
'********************************************************************************************************
'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        txtFromDt.Action = 7
        Call SetFocusToDocument("P")
		txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        txtToDt.Action = 7
        Call SetFocusToDocument("P")
		txtToDt.Focus
    End If
End Sub

'===========================================  2.4.1 POP-UP Open 함수()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
'------------------------------------------  OpenItemInfo()  --------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strItemCd)

	Dim arrRet
	Dim arrParam(5), arrField(16)
	Dim iCalledAspName, IntRetCD

	IsOpenPop = True
	
	arrParam(0) = Trim(lgPlantCD)				' Plant Code
	arrParam(1) = Trim(strItemCd)				' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 	'ITEM_CD				' Field명(0)
	arrField(1) = 2 	'ITEM_NM				' Field명(1)
	arrField(2) = 26 	'UNIT_OF_ORDER_MFG
	arrField(3) = 4		'BASIC_UNIT
	arrField(4) = 28	'ORDER_LT_MFG
	arrField(5) = 33	'MIN_MRP_QTY
	arrField(6) = 34	'MAX_MRP_QTY
	arrField(7) = 35	'ROND_QTY
	arrField(8) = 36	'PROD_MGR	-- ?
	arrField(9) = 15	'MAJOR_SL_CD
	arrField(10) = 13	'PHANTOM_FLG
	arrField(11) = 25	'TRACKING_FLG
	arrField(12) = 17	'VALID_FLG
	arrField(13) = 18	'VALID_FROM_DT
	arrField(14) = 19	'VALID_TO_DT
	arrField(15) = 49	'INSPEC_MGR
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtItemCd.Focus
	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(lgPlantCD)
	arrParam(1) = Trim(strCode)
	arrParam(2) = ""
	arrParam(3) = txtFromDt.Text
	arrParam(4) = txtToDt.Text	
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	txtTrackingNo.Focus
	
End Function

'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(txtItemGroupCd.className) = UCase(PopUpParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
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
	txtItemGroupCd.focus
 
End Function

'=======================================  2.4.2 POP-UP Return값 설정 함수  ==============================
'=	Name : Set???()																						=
'=	Description : Reference 및 POP-UP의 Return값을 받는 부분											=
'========================================================================================================
'------------------------------------------  SetItemInfo()  ---------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet)
	txtItemCd.Value		= arrRet(0)
	txtItemNm.Value		= arrRet(1)
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	txtTrackingNo.Value = arrRet(0)
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	txtItemGroupCd.Value    = arrRet(0)  
	txtItemGroupNm.Value    = arrRet(1)  
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개별 프로그램마다 필요한 개발자 정의 Procedure(Sub, Function, Validation & Calulation 관련 함수)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						3. Event 부																		#
'#	기능: Event 함수에 관한 처리																		#
'#	설명: Window처리, Single처리, Grid처리 작업.														#
'#		  여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.								#
'#		  각 Object단위로 Grouping한다.																	#
'########################################################################################################
'********************************************  3.1 Window처리  ******************************************
'*	Window에 발생 하는 모든 Even 처리																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()

	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'⊙: Load table , B_numeric_format		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    		
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
	Call InitVariables											'⊙: Initializes local global variables
	Call InitSpreadSheet()
	Call InitComboBox()
	Call InitSetting()
	Call FncQuery()
	
	IsFormLoaded = true											'After Loading the Form, the OrderStatus Variables can be Changed.
End Sub
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery()
    FncQuery = False
    Call InitVariables
	If DbQuery = False Then	
		Exit Function
	End If
	FncQuery = False
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = vspdData
	Call SetPopupMenuItemInf("0000111111")
	
    If vspdData.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If    
End Sub

'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then Exit Function
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################
'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	On Error Resume Next
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkfield(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
	    
	If ValidDateCheck(txtFromDt, txtToDt) = False Then Exit Function
	    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
	    
    Call LayerShowHide(1)
	    
    Dim strVal
		
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtFromDt=" & Trim(hProdFromDt.value)
		strVal = strVal & "&txtToDt=" & Trim(hProdToDt.value)
		strVal = strVal & "&txtFromStstus=" & strFromStatus
		strVal = strVal & "&txtToStstus=" & strToStatus
		strVal = strVal & "&txtThirdStstus=" & strThirdStatus
		strVal = strVal & "&txtProdOrderNo=" & Trim(hProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(hItemCd.value)
		strVal = strVal & "&cboOrderType=" & Trim(hOrderType.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(hItemGroupCd.value)
	Else
		If Trim(cboOrderStatus.value) <> "" Then
			strFromStatus	= Trim(cboOrderStatus.value)
			strToStatus		= Trim(cboOrderStatus.value)
			strThirdStatus  = Trim(cboOrderStatus.value)
		ElseIf Trim(cboOrderStatus.value) = "" And IsFormLoaded = True Then		'After Loading the Form, the OrderStatus Variables can be Changed.
			strFromStatus	= ""
			strToStatus		= ""
			strThirdStatus  = ""
		End If

		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtFromDt=" & Trim(txtFromDt.text)
		strVal = strVal & "&txtToDt=" & Trim(txtToDt.text)
		strVal = strVal & "&txtFromStstus=" & strFromStatus
		strVal = strVal & "&txtToStstus=" & strToStatus
		strVal = strVal & "&txtThirdStstus=" & strThirdStatus
		strVal = strVal & "&txtProdOrderNo=" & Trim(txtProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)
		strVal = strVal & "&cboOrderType=" & Trim(cboOrderType.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(txtItemGroupCd.value)
	End If    

    Call RunMyBizASP(MyBizASP, strVal)					'☜: 비지니스 ASP 를 가동 
		
    DbQuery = True                                      '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRows)															'☆: 조회 성공후 실행로직 
	If lgIntFlgMode = PopupParent.OPMD_CMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = PopupParent.OPMD_UMODE	
	Call InitData(LngMaxRows,1)
    vspddata.Focus												'⊙: Indicates that current mode is Update mode
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5 NOWRAP>제조오더 번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더 번호"></TD>
						<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=20 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo txtTrackingNo.value"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>품목그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5>작업계획일자</TD>
						<TD CLASS=TD6>
							<script language =javascript src='./js/p4111pa1_I219408941_txtFromDt.js'></script>
							&nbsp;~&nbsp;
							<script language =javascript src='./js/p4111pa1_I747224920_txtToDt.js'></script>
						</TD>
						<TD CLASS=TD5 NOWRAP>지시구분</TD>
						<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderType" ALT="지시구분" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
					</TR>
					<TR>
						
						<TD CLASS=TD5></TD>
						<TD CLASS=TD6></TD>
						<TD CLASS=TD5 NOWRAP>지시상태</TD>
						<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderStatus" ALT="지시상태" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p4111pa1_vspdData_vspdData.js'></script>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrderType" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hToStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
