<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Resource Consumption Management
'*  3. Program ID			: p4713ma1.asp
'*  4. Program Name			: Resource Consumption By Resource
'*  5. Program Desc			: Resource Consumption By Resource
'*  6. Comproxy List		: 
'*	   Biz ASP  List		: 
'*  7. Modified date(First)	: 2001/12/12
'*  8. Modified date(Last)	: 2002/07/18
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Kang Seong Moon
'* 11. Comment				:
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'#########################################################################################################-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID	= "p4713mb1.asp"								'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_QRY1_ID	= "p4713mb2.asp"								'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "p4713mb3.asp"								'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_LOOK_HDR	= "p4713mb4.asp"								'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_LOOK_DTL	= "p4713mb5.asp"								'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_LOOK_RST	= "p4713mb6.asp"								'☆: Head Query 비지니스 로직 ASP명 

<!-- #Include file="../../inc/lgvariables.inc" -->	

' Grid 1(vspdData1) - Operation
Dim C_ProdtOrderNo			'= 1
Dim C_ProdtOrderNoPopup		'= 2
Dim C_OprNo					'= 3
Dim C_OprNoPopup			'= 4
Dim C_ConsumedDt			'= 5
Dim C_ConsumedTime			'= 6
Dim C_ItemCd				'= 7
Dim C_ItemNm				'= 8
Dim C_Spec					'= 9
Dim C_RoutNo				'= 10
Dim C_ProdtOrderQty			'= 11
Dim C_ProdtOrderUnit		'= 12
Dim C_ProdQtyInOrderUnit	'= 13
Dim C_GoodQtyInOrderUnit	'= 14
Dim C_BadQtyInOrderUnit		'= 15
Dim C_JobCd					'= 16
Dim C_JobNm					'= 17
Dim C_WcCd					'= 18
Dim C_WcNm					'= 19
Dim C_PlanStartDt			'= 20
Dim C_PlanComptDt			'= 21
Dim C_ReleaseDt				'= 22
Dim C_RealStartDt			'= 23
Dim C_OrderStatus			'= 24
Dim C_OrderStatusDesc		'= 25
Dim C_TrackingNo			'= 26
Dim C_OrderType				'= 27
Dim C_OrderTypeDesc			'= 28

' Grid 2(vspdData2) - Operation
Dim C_ResourceCd2			'= 1
Dim C_ResourceNm2			'= 2
Dim C_ResourceTypeNm2		'= 3
Dim C_ResourceGroupCd2		'= 4
Dim C_ResourceGroupNm2		'= 5
Dim	C_Rank2					'= 6
Dim	C_BOR_Efficiency2		'= 7
Dim C_ValidFromDt2			'= 8
Dim C_ValidToDt2			'= 9

Dim strDate
Dim strYear
Dim strMonth
Dim strDay
Dim BaseDate

BaseDate = "<%=GetsvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim ihGridCnt                    'hidden Grid Row Count
Dim intItemCnt					'hidden Grid Row Count
Dim IsOpenPop					'Popup
Dim gSelframeFlg
Dim lgSortKey2
'########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	
End Sub

Sub InitComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobNm
	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderTypeDesc

End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	If lngStartRow = 0 or lngStartRow = "" Then lngStartRow = 1

	ggoSpread.Source = frm1.vspdData1
	With frm1.vspdData1
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_JobCd
			intIndex = .value
			.Col = C_JobNm
			.value = intindex
			.col = C_OrderStatus
			intIndex = .value
			.Col = C_OrderStatusDesc
			.value = intindex
			.col = C_OrderType
			intIndex = .value
			.Col = C_OrderTypeDesc
			.value = intindex			
		Next	
	End With
End Sub

'==========================================  2.2.6 InitRowData()  ========================================== 
'	Name : InitRowData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitRowData(ByVal lngStartRow)

	Dim intIndex

	ggoSpread.Source = frm1.vspdData1
	With frm1.vspdData1
		.Row = lngStartRow
		.col = C_JobCd
		intIndex = .value
		.Col = C_JobNm
		.value = intindex
		.col = C_OrderStatus
		intIndex = .value
		.Col = C_OrderStatusDesc
		.value = intindex
		.col = C_OrderType
		intIndex = .value
		.Col = C_OrderTypeDesc
		.value = intindex
	End With
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
    frm1.txtConsumedDtFrom.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
    frm1.txtConsumedDtTo.text   = strDate
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet(ByVal pvSpdNo)
	Call InitSpreadPosVariables(pvSpdNo)
	
	Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
	
	if pvSpdNo = "*" or pvSpdNo = "A" then
		'------------------------------------------
		' Grid 1 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData1
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols = C_OrderTypeDesc + 1
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit		C_ProdtOrderNo, "오더번호", 18,,,18,2
		ggoSpread.SSSetButton 	C_ProdtOrderNoPopup
		ggoSpread.SSSetEdit		C_OprNo, "공정", 6,,,3,2
		ggoSpread.SSSetButton 	C_OprNoPopup
		ggoSpread.SSSetTime 	C_ConsumedTime,	"자원소비시간",	13,2 ,1 ,1
		ggoSpread.SSSetDate		C_ConsumedDt,	"자원소비일",	13,	2,	parent.gDateFormat
		ggoSpread.SSSetEdit		C_ItemCd, "품목", 18
		ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25
		ggoSpread.SSSetEdit		C_Spec, "규격", 25
		ggoSpread.SSSetFloat	C_ProdtOrderQty, "오더수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_ProdtOrderUnit, "오더단위", 8
		ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "실적수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit, "양품수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQtyInOrderUnit, "불량수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetCombo	C_JobCd, "작업", 10
		ggoSpread.SSSetCombo	C_JobNm, "작업명", 20
		ggoSpread.SSSetEdit		C_WcCd, "작업장", 10,,,7,2
		ggoSpread.SSSetEdit		C_WcNm, "작업장명", 20
		ggoSpread.SSSetDate 	C_PlanStartDt, "착수예정일", 10, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_PlanComptDt, "완료예정일", 10, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_RealStartDt, "실착수일", 10, 2, parent.gDateFormat
		ggoSpread.SSSetCombo	C_OrderStatus, "지시상태", 10
		ggoSpread.SSSetCombo	C_OrderStatusDesc, "지시상태", 10
		ggoSpread.SSSetEdit		C_RoutNo, "라우팅", 8
		ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 25
		ggoSpread.SSSetCombo	C_OrderType, "지시구분", 10
		ggoSpread.SSSetCombo	C_OrderTypeDesc, "지시구분", 10
		' Hidden Columns
		ggoSpread.SSSetDate 	C_ReleaseDt, "작업지시일", 10, 2, parent.gDateFormat
		.ReDraw = true
		Call ggoSpread.MakePairsColumn(C_ProdtOrderNo, C_ProdtOrderNoPopup)
		Call ggoSpread.MakePairsColumn(C_OprNo, C_OprNoPopup)
		Call ggoSpread.SSSetColHidden(C_OrderStatus ,C_OrderStatus , True)
		Call ggoSpread.SSSetColHidden(C_OrderType ,C_OrderType , True)
		Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
		ggoSpread.SSSetSplit2(6)
		ihGridCnt = 0               'Hidden Counter
		intItemCnt = 0
		End With
		
	end if
		
	if pvSpdNo = "*" or pvSpdNo = "B" then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols = C_ValidToDt2 +1										'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		Call GetSpreadColumnPos("B")
		ggoSpread.SSSetEdit		C_ResourceCd2,		"자원코드",		10
		ggoSpread.SSSetEdit		C_ResourceNm2,		"자원명",		20
		ggoSpread.SSSetEdit		C_ResourceTypeNm2,	"자원구분",		10
		ggoSpread.SSSetEdit		C_ResourceGroupCd2, "자원그룹",		10
		ggoSpread.SSSetEdit		C_ResourceGroupNm2, "자원그룹명",	20
		ggoSpread.SSSetFloat	C_Rank2,			"순서",			10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BOR_Efficiency2,	"효율",			10, "7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate		C_ValidFromDt2,		"시작일",		11,	2,	parent.gDateFormat
		ggoSpread.SSSetDate		C_ValidToDt2,		"종료일",		11,	2,	parent.gDateFormat
		
		'Call ggoSpread.MakePairsColumn(,)
		Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
		ggoSpread.SSSetSplit2(1)	
		.ReDraw = true
		End With
	end if
    
    Call SetSpreadLock()
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    With frm1

	ggoSpread.Source = frm1.vspdData1
	.vspdData1.ReDraw = False
	ggoSpread.SpreadLock C_ProdtOrderNo, -1, C_ConsumedDt
	ggoSpread.SpreadLock C_ItemCd, -1, C_OrderTypeDesc
	ggoSpread.SpreadLock frm1.vspdData1.MaxCols, -1, frm1.vspdData1.MaxCols
	
	ggoSpread.SpreadUnLock  C_ConsumedTime, -1 ,C_ConsumedTime
	ggoSpread.SSSetRequired  C_ConsumedTime, -1 , C_ConsumedTime
	.vspdData1.ReDraw = True

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()

    End With

End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, Byval Flag)
    
    With frm1.vspdData1 

	ggoSpread.Source = frm1.vspdData1
	    
    .Redraw = False
    
    If Flag = "C" Then  
		ggoSpread.SSSetRequired C_ProdtOrderNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_OprNo,				pvStartRow, pvEndRow
	Else
		ggoSpread.SSSetProtected C_ProdtOrderNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OprNo,				pvStartRow, pvEndRow
	End If
	
    ggoSpread.SSSetProtected C_JobCd,					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_JobNm,					pvStartRow, pvEndRow

    If Flag = "C" Then
		ggoSpread.SSSetRequired C_ConsumedDt,			pvStartRow, pvEndRow
	Else
		ggoSpread.SSSetProtected C_ConsumedDt,			pvStartRow, pvEndRow
	End If
    
    ggoSpread.SSSetRequired C_ConsumedTime,				pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcCd,					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcNm,					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemCd, 					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemNm,					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Spec,					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RoutNo,					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ProdtOrderQty, 			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ProdtOrderUnit, 			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ProdQtyInOrderUnit, 		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_GoodQtyInOrderUnit, 		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BadQtyInOrderUnit, 		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlanStartDt,				pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlanComptDt, 			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ReleaseDt,				pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RealStartDt,				pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderStatus,				pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderStatusDesc,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RoutNo, 					pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_TrackingNo, 				pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderType,				pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderTypeDesc,			pvStartRow, pvEndRow

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
Sub InitSpreadPosVariables(ByVal pvSpdNo)

	if pvSpdNo = "*" or pvSpdNo = "A" then
		' Grid 1(vspdData1) - Operation
		C_ProdtOrderNo			= 1
		C_ProdtOrderNoPopup		= 2
		C_OprNo					= 3
		C_OprNoPopup			= 4
		C_ConsumedDt			= 5
		C_ConsumedTime			= 6
		C_ItemCd				= 7
		C_ItemNm				= 8
		C_Spec					= 9
		C_RoutNo				= 10
		C_ProdtOrderQty			= 11
		C_ProdtOrderUnit		= 12
		C_ProdQtyInOrderUnit	= 13
		C_GoodQtyInOrderUnit	= 14
		C_BadQtyInOrderUnit		= 15
		C_JobCd					= 16
		C_JobNm					= 17
		C_WcCd					= 18
		C_WcNm					= 19
		C_PlanStartDt			= 20
		C_PlanComptDt			= 21
		C_ReleaseDt				= 22
		C_RealStartDt			= 23
		C_OrderStatus			= 24
		C_OrderStatusDesc		= 25
		C_TrackingNo			= 26
		C_OrderType				= 27
		C_OrderTypeDesc			= 28
	end if
	
	if pvSpdNo = "*" or pvSpdNo = "A" then
		' Grid 2(vspdData2) - Operation
		C_ResourceCd2			= 1
		C_ResourceNm2			= 2
		C_ResourceTypeNm2		= 3
		C_ResourceGroupCd2		= 4
		C_ResourceGroupNm2		= 5
		C_Rank2					= 6
		C_BOR_Efficiency2		= 7
		C_ValidFromDt2			= 8
		C_ValidToDt2			= 9	
	end if

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
  	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
  	Case "A"
 		ggoSpread.Source = frm1.vspdData1 
  		
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_ProdtOrderNo			= iCurColumnPos(1)
		C_ProdtOrderNoPopup		= iCurColumnPos(2)
		C_OprNo					= iCurColumnPos(3)
		C_OprNoPopup			= iCurColumnPos(4)
		C_ConsumedDt			= iCurColumnPos(5)
		C_ConsumedTime			= iCurColumnPos(6)
		C_ItemCd				= iCurColumnPos(7)
		C_ItemNm				= iCurColumnPos(8)
		C_Spec					= iCurColumnPos(9)
		C_RoutNo				= iCurColumnPos(10)
		C_ProdtOrderQty			= iCurColumnPos(11)
		C_ProdtOrderUnit		= iCurColumnPos(12)
		C_ProdQtyInOrderUnit	= iCurColumnPos(13)
		C_GoodQtyInOrderUnit	= iCurColumnPos(14)
		C_BadQtyInOrderUnit		= iCurColumnPos(15)
		C_JobCd					= iCurColumnPos(16)
		C_JobNm					= iCurColumnPos(17)
		C_WcCd					= iCurColumnPos(18)
		C_WcNm					= iCurColumnPos(19)
		C_PlanStartDt			= iCurColumnPos(20)
		C_PlanComptDt			= iCurColumnPos(21)
		C_ReleaseDt				= iCurColumnPos(22)
		C_RealStartDt			= iCurColumnPos(23)
		C_OrderStatus			= iCurColumnPos(24)
		C_OrderStatusDesc		= iCurColumnPos(25)
		C_TrackingNo			= iCurColumnPos(26)
		C_OrderType				= iCurColumnPos(27)
		C_OrderTypeDesc			= iCurColumnPos(28)
	Case "B"
 		ggoSpread.Source = frm1.vspdData2
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_ResourceCd2			= iCurColumnPos(1)
		C_ResourceNm2			= iCurColumnPos(2)
		C_ResourceTypeNm2		= iCurColumnPos(3)
		C_ResourceGroupCd2		= iCurColumnPos(4)
		C_ResourceGroupNm2		= iCurColumnPos(5)
		C_Rank2					= iCurColumnPos(6)
		C_BOR_Efficiency2		= iCurColumnPos(7)
		C_ValidFromDt2			= iCurColumnPos(8)
		C_ValidToDt2			= iCurColumnPos(9)
  	End Select
  
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
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
	arrField(0) = "PLANT_CD"					' Field명(0)
	arrField(1) = "PLANT_NM"					' Field명(1)
	
	arrHeader(0) = "공장"					' Header명(0)
	arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdtOrderNo()  -------------------------------------------------
'	Name : OpenProdtOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdtOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtProdtOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtConsumedDtFrom.Text
	arrParam(2) = frm1.txtConsumedDtTo.Text
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdtOrderNo.value) 
	arrParam(6) = ""
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = ""
	
	iCalledAspName = AskPRAspName("p4111pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdtOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdtOrderNo.focus
	
End Function

'------------------------------------------  OpenWcCd()  ------------------------------------------------
'	Name : OpenWcCd()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"											' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")	' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
	
End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "자원팝업"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtResourceCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			
	arrParam(5) = "자원"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "자원"		
    arrHeader(1) = "자원명"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtResourceCd.focus
		
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode			' Item Code
	arrParam(2) = "12!MO"			' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field명(0)
	arrField(1) = 2 '"ITEM_NM"			' Field명(1)
    
    iCalledAspName = AskPRAspName("b1b11pa3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenGridProdtOrderNo()  -------------------------------------------------
'	Name : OpenGridProdtOrderNo()
'	Description : Grid Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenGridProdtOrderNo(Byval strProdtOrderNo, Byval Row)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtProdtOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtConsumedDtFrom.Text
	arrParam(2) = frm1.txtConsumedDtTo.Text
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = strProdtOrderNo
	arrParam(6) = ""
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = ""
	
	iCalledAspName = AskPRAspName("p4111pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGridProdtOrderNo(arrRet, Row)
	End If	
End Function

'------------------------------------------  OpenGridOprNo()  -------------------------------------------------
'	Name : OpenGridOprNo()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGridOprNo(Byval strProdtOrderNo, Byval strOprNo, Byval Row)
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
'	If IsOpenPop = True Or UCase(frm1.txtOprCd.className) = "PROTECTED" Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If strProdtOrderNo = "" Then
		Call DisplayMsgBox("971012","X" , "제조오더번호","X")
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = strProdtOrderNo
	arrParam(2) = ""
	
	iCalledAspName = AskPRAspName("p4112pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4112pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGridOprNo(arrRet, Row)
	End If	
End Function

'------------------------------------------  OpenPartRef()  -------------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPartRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
   	With frm1.vspdData1
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4311ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4311ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P43/p4311ra1.asp", Array(arrParam(0), arrParam(1), arrParam(2)), _
	'	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenOprRef()  -------------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
   	With frm1.vspdData1
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4111ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P41/p4111ra1.asp", Array(arrParam(0), arrParam(1)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
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

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)	'☆: 조회 조건 데이타 

   	With frm1.vspdData1
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	IsOpenPop = True

	iCalledAspName = AskPRAspName("p4411ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4411ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P44/p4411ra1.asp", Array(arrParam(0), arrParam(1), arrParam(2)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
    	With frm1.vspdData1
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4511ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4511ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P45/p4511ra1.asp", Array(arrParam(0), arrParam(1)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
    	With frm1.vspdData1
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4412ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4412ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P44/p4412ra1.asp", Array(arrParam(0), arrParam(1)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
    frm1.txtPlantCd.Value    = arrRet(0)		
    frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	frm1.txtResourceCd.Value    = arrRet(0)		
	frm1.txtResourceNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdtOrderNo()  -------------------------------------------
'	Name : SetProdtOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdtOrderNo(byval arrRet)
    frm1.txtProdtOrderNo.Value    = arrRet(0)
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  -------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
    With frm1
	.txtItemCd.value = arrRet(0)
	.txtItemNm.value = arrRet(1)
    End With
End Function

'------------------------------------------  SetGridProdtOrderNo()  --------------------------------------
'	Name : SetGridProdtOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetGridProdtOrderNo(byval arrRet, Byval Row)
	With frm1.vspdData1
		.Row = Row
		.Col = C_ProdtOrderNo
   		.Text = arrRet(0)
   	End With
    Call LookUpOrderHeader(arrRet(0), Row)
End Function

'------------------------------------------  SetGridOprNo()  ---------------------------------------------
'	Name : SetGridOprNo()
'	Description : Operation Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetGridOprNo(byval arrRet, Byval Row)
	Dim strProdtOrderNo
	With frm1.vspdData1
		.Row = Row
		.Col = C_ProdtOrderNo
		strProdtOrderNo = .Value
		.Col = C_OprNo
   		.Text = arrRet(0)
   	End With
    Call LookUpOrderDetail(strProdtOrderNo, arrRet(0), Row)
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'-------------------------------------  LookUpOrderHeader()  -----------------------------------------
'	Name : LookUpOrderHeader()
'	Description : LookUp Order Header
'---------------------------------------------------------------------------------------------------------
Function LookUpOrderHeader(Byval strProdtOrderNo, Byval Row)
 
   Dim strVal

	If strProdtOrderNo = "" Then Exit Function
	
    strVal = BIZ_PGM_LOOK_HDR & "?txtMode=" & parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtProdtOrderNo=" & Trim(strProdtOrderNo)	'☜: 조회 조건 데이타 
    strVal = strVal & "&txtRow=" & Row								'☜: 조회 조건 데이타 

    Call RunMyBizASP(MyBizASP, strVal)								'☜: 비지니스 ASP 를 가동 
	
End Function

'-------------------------------------  LookUpOrderHeaderSuccess()  -----------------------------------------
'	Name : LookUpOrderHeaderSuccess()
'	Description : LookUp Order Header Success
'---------------------------------------------------------------------------------------------------------
Function LookUpOrderHeaderSuccess(Byval Row)
	Dim strProdtOrderNo
	
	Call InitRowData(Row)
	
	With frm1.vspdData1
	ggoSpread.Source = frm1.vspdData1
	.Row = Row
	.Col = C_ProdtOrderNo
	strProdtOrderNo = Trim(.Value)
	.Col = C_OprNo
		If .Value <> "" then
			Call LookUpOrderDetail(strProdtOrderNo,Trim(.Value),Row)
		End if
	End With
End Function

'-------------------------------------  LookUpOrderHeaderFail()  -----------------------------------------
'	Name : LookUpOrderHeaderFail()
'	Description : LookUp Order Header Fail
'---------------------------------------------------------------------------------------------------------
Function LookUpOrderHeaderFail(Byval Row)
    Call InitRow(Row, "H")
End Function

'-------------------------------------  LookUpOrderDetail()  -----------------------------------------
'	Name : LookUpOrderDetail()
'	Description : LookUp Order Detail
'---------------------------------------------------------------------------------------------------------
Function LookUpOrderDetail(Byval strProdtOrderNo, Byval strOprNo, Byval Row)
    
   Dim strVal

	If strProdtOrderNo = "" Then Exit Function
	
    strVal = BIZ_PGM_LOOK_DTL & "?txtMode=" & parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtProdtOrderNo=" & Trim(strProdtOrderNo)	'☜: 조회 조건 데이타 
    strVal = strVal & "&txtOprNo=" & Trim(strOprNo)					'☜: 조회 조건 데이타 
    strVal = strVal & "&txtRow=" & Row								'☜: 조회 조건 데이타 

    Call RunMyBizASP(MyBizASP, strVal)								'☜: 비지니스 ASP 를 가동 
	
End Function

'-------------------------------------  LookUpOrderDetailSuccess()  -----------------------------------------
'	Name : LookUpOrderDetailSuccess()
'	Description : LookUp Order Header
'---------------------------------------------------------------------------------------------------------
Function LookUpOrderDetailSuccess(Byval Row)
	Dim strProdtOrderNo, strConsumedDt 
	Call InitRowData(Row)
	ggoSpread.Source = frm1.vspdData1
	With frm1.vspddata1
		.Row = Row
		.Col = C_ProdtOrderNo
		strProdtOrderNo = .Value
		.Col = C_ItemCd
		frm1.h2ItemCd.Value = .Value
		.Col = C_OprNo
		frm1.h2OprNo.Value = .Value
		.Col = C_RoutNo
		frm1.h2RoutNo.Value = .Value
		.Col = C_ConsumedDt
		strConsumedDt = .Text
    End With
    Call LookUpProductionResults( strProdtOrderNo ,frm1.h2OprNo.Value, strConsumedDt ,Row)
End Function

'-------------------------------------  LookUpOrderDetailFail()  -----------------------------------------
'	Name : LookUpOrderDetailFail()
'	Description : LookUp Order Detail Fail
'---------------------------------------------------------------------------------------------------------
Function LookUpOrderDetailFail(Byval Row)
    Call InitRow(Row, "D")
End Function

'-------------------------------------  LookUpProductionResults()  -----------------------------------------
'	Name : LookUpProductionResults()
'	Description : Look Up Production Results
'---------------------------------------------------------------------------------------------------------
Function LookUpProductionResults(Byval strProdtOrderNo, Byval strOprNo, Byval strConsumedDt, Byval Row)
    
   Dim strVal

	If strProdtOrderNo = "" Then Exit Function
	
    strVal = BIZ_PGM_LOOK_RST & "?txtMode=" & parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtProdtOrderNo=" & Trim(strProdtOrderNo)	'☜: 조회 조건 데이타 
    strVal = strVal & "&txtOprNo=" & Trim(strOprNo)					'☜: 조회 조건 데이타 
    strVal = strVal & "&txtConsumedDt=" & Trim(strConsumedDt)				'☜: 조회 조건 데이타 
    strVal = strVal & "&txtRow=" & Row								'☜: 조회 조건 데이타 

    Call RunMyBizASP(MyBizASP, strVal)								'☜: 비지니스 ASP 를 가동 
	
End Function

'-------------------------------------  LookUpProductionResultsSuccess()  -----------------------------------------
'	Name : LookUpProductionResultsSuccess()
'	Description : LookUp Production Results Success
'---------------------------------------------------------------------------------------------------------
Function LookUpProductionResultsSuccess(Byval Row)
	Call InitRowData(Row)
	ggoSpread.Source = frm1.vspdData1
	With frm1.vspddata1
		.Row = Row		
		.Col = C_ItemCd
		If .Text = "" Then
			frm1.vspdData2.MaxRows = 0
			Exit Function
		End If
		frm1.h2ItemCd.value = .Text
		.Col = C_RoutNo
		If .Text = "" Then
			frm1.vspdData2.MaxRows = 0
			Exit Function
		End If
		frm1.h2RoutNo.value = .Text
		.Col = C_OprNo
		frm1.h2OprNo.value = .Text
    End With
    
    Call DbDtlquery
End Function

'-------------------------------------  LookUpProductionResultsFail()  -----------------------------------------
'	Name : LookUpProductionResultsFail()
'	Description : LookUp Production Results Fail
'---------------------------------------------------------------------------------------------------------
Function LookUpProductionResultsFail(Byval Row)
    Call InitRow(Row, "R")
    
    ggoSpread.Source = frm1.vspdData1
	With frm1.vspddata1
		.Row = Row		
		.Col = C_ItemCd
		If .Text = "" Then
			frm1.vspdData2.MaxRows = 0
			Exit Function
		End If
		frm1.h2ItemCd.value = .Text
		.Col = C_RoutNo
		If .Text = "" Then
			frm1.vspdData2.MaxRows = 0
			Exit Function
		End If
		frm1.h2RoutNo.value = .Text
		.Col = C_OprNo
		frm1.h2OprNo.value = .Text
    End With
    
    Call DbDtlQuery
End Function


'-------------------------------------  InitRow()  -----------------------------------------
'	Name : InitRow()
'	Description : Initialize Row
'---------------------------------------------------------------------------------------------------------
Function InitRow(Byval Row, Byval strFlag)

	frm1.h2ItemCd.Value = ""
	frm1.h2OprNo.Value = ""
	frm1.h2RoutNo.Value = ""

	ggoSpread.Source = frm1.vspdData1
	With frm1.vspddata1
		.Row = Row
		
	    If strFlag = "H" Then
			.Col = C_ProdtOrderNo
			.value = ""
			.Col = C_OprNo
			.value = ""
			.Col = C_ItemCd
			.value = ""
			.Col = C_ItemNm
			.value = ""
			.Col = C_RoutNo
			.value = ""
			.Col = C_ProdtOrderQty
			.value = ""
			.Col = C_ProdtOrderUnit
			.value = ""
			.Col = C_TrackingNo
			.value = ""
			.Col = C_OrderType
			.value = ""
			.Col = C_OrderTypeDesc
			.value = ""
			.Col = C_JobCd
			.value = ""
			.Col = C_JobNm
			.value = ""
			.Col = C_WcCd
			.value = ""
			.Col = C_WcNm
			.value = ""
			.Col = C_ProdQtyInOrderUnit
			.value = ""
			.Col = C_GoodQtyInOrderUnit
			.value = ""
			.Col = C_BadQtyInOrderUnit
			.value = ""
			.Col = C_PlanStartDt
			.value = ""
			.Col = C_PlanComptDt
			.value = ""
			.Col = C_ReleaseDt
			.value = ""
			.Col = C_RealStartDt
			.value = ""
			.Col = C_OrderStatus
			.value = ""
			.Col = C_OrderStatusDesc
			.value = ""
						   
	    ElseIf strFlag = "D" Then

			.Col = C_JobCd
			.value = ""
			.Col = C_JobNm
			.value = ""
			.Col = C_WcCd
			.value = ""
			.Col = C_WcNm
			.value = ""
			.Col = C_ProdQtyInOrderUnit
			.value = ""
			.Col = C_GoodQtyInOrderUnit
			.value = ""
			.Col = C_BadQtyInOrderUnit
			.value = ""
			.Col = C_PlanStartDt
			.value = ""
			.Col = C_PlanComptDt
			.value = ""
			.Col = C_ReleaseDt
			.value = ""
			.Col = C_RealStartDt
			.value = ""
			.Col = C_OrderStatus
			.value = ""
			.Col = C_OrderStatusDesc
			.value = ""
		
		Else
			.Col = C_ProdQtyInOrderUnit
			.value = ""
			.Col = C_GoodQtyInOrderUnit
			.value = ""
			.Col = C_BadQtyInOrderUnit
			.value = ""
		
		End If
		
    End With

End Function


'==============================================================================
' Function : ConvToSec()
' Description : 저장시에 각 시간 데이터들을 초로 환산 
'==============================================================================
Function ConvToSec(ByVal Str)

	If Str = "" Then
		ConvToSec = 0
	ElseIf Len(Trim(Str)) = 8 Then
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
	Else
		ConvToSec = -999999
	End If

End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Row = " & lRow & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Col = " & lCol & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Action = 0" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.focus" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Row = " & lRow & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Col = " & lCol & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Action = 0" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
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
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     				'⊙: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                          			'⊙: Lock  Suitable  Field
    Call InitSpreadSheet("*")                                                    				'⊙: Setup the Spread sheet

	Call SetDefaultVal
    Call InitVariables		'⊙: Initializes local global variables
    Call InitComboBox()
    Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.Value = parent.gPlant
		frm1.txtPlantNm.Value = parent.gPlantNm
		frm1.txtResourceCd.focus
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
'   Event Name : vspdData1_Change
'   Event Desc :
'==========================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

	Dim strProdtOrderNo, strOprNo, strConsumedDt

    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
    
	With frm1.vspdData1

		Select Case Col
    
		    Case C_ProdtOrderNo
				
				.Col = C_ProdtOrderNo
				strProdtOrderNo = .Value
				If strProdtOrderNo = "" Then
					Call InitRow(Row, "H")
					Exit Sub
				End If
				Call LookUpOrderHeader(strProdtOrderNo, Row)
				
		    Case C_OprNo
				.Col = C_OprNo
				strOprNo = .Value
				
				If strOprNo = "" Then
					Call InitRow(Row, "D")
					Exit Sub
				End If
				.Col = C_ProdtOrderNo
				strProdtOrderNo = .Value
				If strProdtOrderNo = "" Then
					Call DisplayMsgBox("971012","X", "오더번호","X")
					.Col = C_OprNo
					.Value = ""
					Exit Sub
				End If
				Call LookUpOrderDetail(strProdtOrderNo, strOprNo, Row)
				
			Case C_ConsumedDt
				.Col = C_ConsumedDt
				strConsumedDt = .Text
				If strConsumedDt = "" Then
					Call InitRow(Row, "D")
					Exit Sub
				End If
				.Col = C_ProdtOrderNo
				strProdtOrderNo = .Value
				If strProdtOrderNo = "" Then
					Call DisplayMsgBox("971012","X", "오더번호","X")
					.Col = C_OprNo
					.Value = ""
					Exit Sub
				End If
				
				.Col = C_OprNo
				strOprNo = .Value
				If strOprNo = "" Then
					Call DisplayMsgBox("971012","X", "공정","X")
					.Col = C_ConsumedDt
					.Value = ""
					Exit Sub
				End If
								
				Call LookUpProductionResults(strProdtOrderNo, strOprNo, strConsumedDt , Row)
				
		End Select
    
    End With
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim strProdtOrderNo

	With frm1.vspdData1
		ggoSpread.Source = frm1.vspdData1
		If Row < 1 Then Exit Sub

		.Row = Row

		Select Case Col

		    Case C_ProdtOrderNoPopup
				.Col = C_ProdtOrderNo
				Call OpenGridProdtOrderNo(.Text, Row)
				Call SetActiveCell(frm1.vspdData1,C_ProdtOrderNo,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
		    Case C_OprNoPopup

				.Col = C_ProdtOrderNo
				strProdtOrderNo = .Text
				.Col = C_OprNo
				Call OpenGridOprNo(strProdtOrderNo, .Text, Row)
				Call SetActiveCell(frm1.vspdData1,C_ProdtOrderNo,Row,"M","X","X")
				Set gActiveElement = document.activeElement

		End Select

    End With
    
End Sub
'========================================================================================
' Function Name : vspdDat1a_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
  	
  	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
  	Else
  		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
  	End If
  	
  	gMouseClickStatus = "SPC"   
     
  	Set gActiveSpdSheet = frm1.vspdData1
     
  	If frm1.vspdData1.MaxRows = 0 Then
  		Exit Sub
  	End If
  	
  	If Row <= 0 Then
  		ggoSpread.Source = frm1.vspdData1 
  		If lgSortKey = 1 Then
  			ggoSpread.SSSort Col					'Sort in Ascending
  			lgSortKey = 2
  		Else
  			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
  			lgSortKey = 1
  		End If
  		
  		With frm1.vspdData1
			
			.Row = .ActiveRow
			.Col = C_ItemCd
			If .Text = "" Then
				frm1.vspdData2.MaxRows = 0
				Exit Sub
			End If
			frm1.h2ItemCd.value = .Text
			.Col = C_RoutNo
			If .Text = "" Then
				frm1.vspdData2.MaxRows = 0
				Exit Sub
			End If
			frm1.h2RoutNo.value = .Text
			.Col = C_OprNo
			frm1.h2OprNo.value = .Text
		End With
		If DbDtlQuery = False Then
		    Call RestoreToolBar()	
		    Exit Sub
		End If  
 	Else
 		'------ Developer Coding part (Start)
 	 	With frm1.vspdData1
			If .MaxRows <= 0 Then Exit Sub
			If Row < 1 Then Exit Sub
			.Row = Row
			.Col = C_ItemCd
			If .Text = "" Then
				frm1.vspdData2.MaxRows = 0
				Exit Sub
			End If
			frm1.h2ItemCd.value = .Text
			.Col = C_RoutNo
			If .Text = "" Then
				frm1.vspdData2.MaxRows = 0
				Exit Sub
			End If
			frm1.h2RoutNo.value = .Text
			.Col = C_OprNo
			frm1.h2OprNo.value = .Text
		End With
		If DbDtlQuery = False Then
		    Call RestoreToolBar()	
		    Exit Sub
		End If  
		'------ Developer Coding part (End)
  	End If

End Sub

'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
  	
  	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
  	
  	gMouseClickStatus = "SP2C"   
     
  	Set gActiveSpdSheet = frm1.vspdData2
     
  	If frm1.vspdData2.MaxRows = 0 Then
  		Exit Sub
  	End If
  	
  	If Row <= 0 Then
  		ggoSpread.Source = frm1.vspdData2 
  		If lgSortKey2 = 1 Then
  			ggoSpread.SSSort Col					'Sort in Ascending
  			lgSortKey2 = 2
  		Else
  			ggoSpread.SSSort Col, lgSortKey2		'Sort in Descending
  			lgSortKey2 = 1
  		End If
 	Else
  		'------ Developer Coding part (Start)
 	 	'------ Developer Coding part (End)
 	
  	End If

End Sub


'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData1
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData2
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
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
    dim pvSpdNo
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    pvSpdNo = gActiveSpdSheet.id
    Call InitSpreadSheet(pvSpdNo)
	call InitComboBox()
	select case pvSpdNo 
    case "A"
		ggoSpread.Source = frm1.vspdData1
    case "B"
		ggoSpread.Source = frm1.vspdData2
    end select
    Call ggoSpread.ReOrderingSpreadData
	Call InitData(1)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
' 	If NewCol = C_XXX or Col = C_XXX Then
 '		Cancel = True
 '		Exit Sub
 '	End If
     ggoSpread.Source = frm1.vspdData1
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
 	 ggoSpread.Source = frm1.vspdData2
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("B")
End Sub 

'========================================================================================
' Function Name : vspdData1_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub 
  
'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'		========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SP2C" Then
        gMouseClickStatus = "SP2CR"
     End If
End Sub 


'==========================================================================================
'   Event Name :vspdData1_DblClick
'   Event Desc :
'==========================================================================================

Sub vspdData1_DblClick(index , ByVal Col , ByVal Row )
     
    ggoSpread.Source = frm1.vspdData1(index)
End Sub


'==========================================================================================
'   Event Name :vspdData1_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspdData1_KeyPress(index , KeyAscii )

End Sub


'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData1 

    If Row >= NewRow Then
        Exit Sub
    End If

	 '----------  Coding part  -------------------------------------------------------------

    End With

End Sub


'==========================================================================================
'   Event Name : vspdData1_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================

Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData1
	
		.Row = Row
    
		Select Case Col
			Case  1
				.Col = Col
				intIndex = .Value
				.Col = C_BillFG
				.Value = intIndex
		End Select
	End With
End Sub


'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If 
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub


'================================================ txtBox_onChange() ===============================
'   Event Name : txtBox_onChange()
'   Event Desc : 
'==========================================================================================


'=======================================================================================================
'   Event Name : txtConsumedDtFrom_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtConsumedDtFrom_DblClick(Button)
    If Button = 1 Then
        frm1.txtConsumedDtFrom.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtConsumedDtFrom.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtConsumedDtTo_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtConsumedDtTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtConsumedDtTo.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtConsumedDtTo.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtConsumedDtFrom_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtConsumedDtFrom_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtConsumedDtTo_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtConsumedDtTo_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
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
   
    ggoSpread.Source = frm1.vspdData1										'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then									'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
   
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    If ValidDateCheck(frm1.txtConsumedDtFrom, frm1.txtConsumedDtTo) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
  	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    Call InitVariables															'⊙: Initializes local global variables
    
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
    End If  															'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
   
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 

    Dim IntRetCD , LngRows
    
    FncSave = False                                             '⊙: Processing is NG
    
    Err.Clear													'☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
       Exit Function
    End If
    
    With frm1.vspdData1
     
    For LngRows = 1 To .MaxRows
      .Row = LngRows
      .Col = 0
      If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then
      
        .Col = C_ConsumedDt
        	 
			If CompareDateByFormat(.text,"<%=strDate%>","자원소비일","현재일","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
			  Exit Function               
			End If  
			
		.Col = C_ConsumedTime					
			If .Text = "<%=ConvToTimeFormat(0)%>" Then
				Call DisplayMsgBox("189715", "x", "x", "x")
				Exit Function
			End If
	   End If	
    Next    
    
    End With
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function						'☜: Save db data
    
    FncSave = True                                              '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	If frm1.vspdData1.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData1.focus
    Set gActiveElement = document.activeElement 
	frm1.vspdData1.EditMode = True
	    
	frm1.vspdData1.ReDraw = False
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.CopyRow
    frm1.vspdData1.Col = C_ProdtOrderNo
    frm1.vspdData1.Text = ""
    'SetSpreadColor frm1.vspdData1.ActiveRow, "C"
    SetSpreadColor frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow, "C"
    frm1.vspdData1.ReDraw = True
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
	If frm1.vspdData1.MaxRows < 1 Then Exit Function	
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    Call initData(frm1.vspdData1.ActiveRow)
End Function



'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim intRetCD
	Dim imRow
	Dim i
	
	On Error Resume Next
	FncInsertRow = false
	
	If IsNumeric(Trim(pvRowCnt)) Then
 		imRow = Cint(pvRowCnt)
 	Else
	 	imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
 			Exit Function
 		End If
 	End if
	
    With frm1
	.vspdData1.focus
	Set gActiveElement = document.activeElement 
	ggoSpread.Source = .vspdData1
	.vspdData1.ReDraw = False
	'ggoSpread.InsertRow
	ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
	SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow -1, "C"
	for i=0 to imRow - 1
		.vspdData1.Row = .vspdData1.ActiveRow + i
		.vspdData1.Col = C_ConsumedDt
		.vspdData1.text = strDate
		.vspdData1.Col = C_ConsumedTime
		.vspdData1.text = "<%=ConvToTimeFormat(0)%>"
	next
	.vspdData1.ReDraw = True
	'SetSpreadColor .vspdData1.ActiveRow, "C"
	frm1.vspdData2.MaxRows = 0
    End With

    If Err.number = 0 Then FncInsertRow = True

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

    Dim lDelRows

    ggoSpread.Source = frm1.vspdData1
    If frm1.vspdData1.MaxRows < 1 Then Exit Function
    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                           '☜: Protect system from crashing
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
    Call parent.FncExport(parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         '☜:화면 유형, Tab 유무 
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



 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'******************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    
    Err.Clear							'☜: Protect system from crashing

    DbQuery = False                                                         			'⊙: Processing is NG
    
    
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtResourceCD=" & Trim(frm1.hResourceCD.value)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.hProdtOrderNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtConsumedDtFrom=" & Trim(frm1.hConsumedDtFrom.text)
		strVal = strVal & "&txtConsumedDtTo=" & Trim(frm1.hConsumedDtTo.text)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtResourceCD=" & Trim(frm1.txtResourceCD.value)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.txtProdtOrderNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtConsumedDtFrom=" & Trim(frm1.txtConsumedDtFrom.text)
		strVal = strVal & "&txtConsumedDtTo=" & Trim(frm1.txtConsumedDtTo.text)
	End If


    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

	
    DbQuery = True                                                          	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(LngMaxRow)													'☆: 조회 성공후 실행로직 
	
	Dim LngRow
	
	Call SetToolBar("11001111001111")											'⊙: 버튼 툴바 제어 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field

	If frm1.vspdData1.MaxRows <= 0 Then Exit Function

	Call InitData(LngMaxRow)

    With frm1.vspdData1
		.Redraw = False	
		.Row = 1
		.Col = C_ItemCd
		frm1.h2ItemCd.value = .Text
		.Col = C_RoutNo
		frm1.h2RoutNo.value = .Text
		.Col = C_OprNo
		frm1.h2OprNo.value = .Text
		.Redraw = True
	End With
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement	
	End If
	
	If DbDtlQuery = False Then
		Call RestoreToolBar()	
		Exit Function 
	End If
	
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
	
End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery가 실패일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryNotOk()

Call SetToolBar("11001101001111")											'⊙: 버튼 툴바 제어 

frm1.txtPlantCd.focus


End Function


'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal

	DbDtlQuery = False   

	Call LayerShowHide(1)
    
    With frm1
		.vspddata2.MaxRows = 0
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(.h2ItemCd.Value)
		strVal = strVal & "&txtOprNo=" & Trim(.h2OprNo.Value)
		strVal = strVal & "&txtRoutNo=" & Trim(.h2RoutNo.Value)
		strVal = strVal & "&txtResourceCd=" & Trim(.hResourceCd.Value)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
    End With

	Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    DbDtlQuery = True

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows   
    Dim strVal, strDel
	Dim ChkTimeVal
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size

    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0
    
    DbSave = False                                                          '⊙: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1
		.txtMode.value = parent.UID_M0002										'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
		
    '-----------------------
    'Data manipulate area
    '-----------------------

	With frm1.vspdData1


    For IntRows = 1 To .MaxRows
    
		.Row = IntRows
		.Col = 0
		
		Select Case .Text
	    
		    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
		    
				strVal = ""
				
				If .Text = ggoSpread.InsertFlag Then
					strVal = strVal & "C" & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
				Else
					strVal = strVal & "U" & iColSep				'⊙: U=Update
				End If
				strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
				.Col = C_ProdtOrderNo
				strVal = strVal & Trim(.Text) & iColSep
				.Col = C_OprNo
				strVal = strVal & Trim(.Text) & iColSep
				strVal = strVal & UCase(Trim(frm1.txtResourceCd.value)) & iColSep
				.Col = C_ConsumedDt
				strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
				.Col = C_ConsumedTime
				
				ChkTimeVal = ConvToSec(Trim(.Text))
				If ChkTimeVal = -999999	Then
					Call DisplayMsgBox("970029", vbInformation, "자원소비시간", "", I_MKSCRIPT)
					Call SheetFocus(arrVal(1),8,I_MKSCRIPT)
					Response.End	
				Else
					strVal = strVal & ChkTimeVal & parent.gRowSep
				End If
				

		    Case ggoSpread.DeleteFlag
				
				strDel = "" 
				
				strDel = strDel & "D" & iColSep				'⊙: D=Delete
				strDel = strDel & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
				.Col = C_ProdtOrderNo
				strDel = strDel & Trim(.Text) & iColSep
				.Col = C_OprNo
				strDel = strDel & Trim(.Text) & iColSep
				strDel = strDel & UCase(Trim(frm1.txtResourceCd.value)) & iColSep
				.Col = C_ConsumedDt
				strDel = strDel & UNIConvDate(Trim(.Text)) & parent.gRowSep
				
		End Select
		
		
		.Col = 0
		Select Case .Text
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			    
		         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
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
			         
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
			          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
			       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
			         
		         iTmpDBuffer(iTmpDBufferCount) =  strDel         
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
		End Select
		
    Next

    End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 비지니스 ASP 를 가동 

    End With

    DbSave = True                                                           ' ⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
    lgLngCurRows = 0                            'initializes Deleted Rows Count

	ggoSpread.source = frm1.vspdData1
    frm1.vspdData1.MaxRows = 0
	
	Call RemovedivTextArea
	Call MainQuery
	
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
</SCRIPT>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : 시간 형식으로 변경 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
				
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</Script>

<!-- #Include file="../../inc/uni2KCM.inc" -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자원소비등록(자원별)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPartRef()">부품내역</A> | <A href="vbscript:OpenOprRef()">공정내역</A> | <A href="vbscript:OpenProdRef()">실적내역</A> | <A href="vbscript:OpenRcptRef()">입고내역</A> | <A href="vbscript:OpenConsumRef()">부품소비내역</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>

									<TD CLASS=TD5 NOWRAP>자원소비일</TD>
									<TD CLASS=TD6>
										<script language =javascript src='./js/p4713ma1_I565424257_txtConsumedDtFrom.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p4713ma1_I164507341_txtConsumedDtTo.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>자원</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="자원"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWCCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()"> <INPUT TYPE=TEXT ID="txtWCNm" NAME="arrCond" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>제조오더 번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdtOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더 번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdtOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdtOrderNo()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR HEIGHT="70%">
							<TD WIDTH="100%" colspan=4>
								<script language =javascript src='./js/p4713ma1_A_vspdData1.js'></script>
							</TD>
						</TR>
						<TR HEIGHT="30%">
							<TD WIDTH="100%" colspan=4>
								<script language =javascript src='./js/p4713ma1_B_vspdData2.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hResourceCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdtOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hConsumedDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="hConsumedDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="h2ItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="h2RoutNo" tag="24">
<INPUT TYPE=HIDDEN NAME="h2OprNo" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
