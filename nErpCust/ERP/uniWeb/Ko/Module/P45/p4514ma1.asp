
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4514ma1.asp
'*  4. Program Name         : 입고대기현황조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 2001/11/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->							<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우 -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                      

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY1_ID						= "p4514mb1.asp"		'☆: Release Production Order Query ASP명 
Const BIZ_PGM_QRY2_ID						= "p4514mb2.asp"		'☆: Operation Query ASP명 
Const BIZ_PGM_QRY3_ID						= "p4514mb3.asp"		'☆: Reservation Query ASP명 

Const C_SHEETMAXROWS = 30

'-------------------------------
' Column Constants : Spread 1 
'-------------------------------
Dim C_ItemCd					'= 1
Dim C_ItemNm					'= 2
Dim C_Spec						' =3
Dim C_WaitQtyInOrderUnit		'= 4
Dim C_OrderUnit					'= 5
Dim C_ProdQtyInOrderUnit		'= 6
Dim C_GoodQtyInOrderUnit		'= 7
Dim C_BadQtyInOrderUnit			'= 8
Dim C_ReceiptQtyInOrderUnit		'= 9
Dim C_WaitQtyInBaseUnit			'= 10
Dim C_BaseUnit					'= 11
Dim C_ProdQtyInBaseUnit			'= 12
Dim C_GoodQtyInBaseUnit			'= 13
Dim C_BadQtyInBaseUnit			'= 14
Dim C_ReceiptQtyInBaseUnit		'= 15
Dim C_ItemGroupCd
Dim C_ItemGroupNm

'-------------------------------
' Column Constants : Spread 2 
'-------------------------------
Dim C_WCCd						'= 1
Dim C_WcNm						'= 2
Dim C_WaitQtyInOrderUnit1		'= 3
Dim C_OrderUnit1				'= 4
Dim C_ProdQtyInOrderUnit1		'= 5
Dim C_GoodQtyInOrderUnit1		'= 6
Dim C_BadQtyInOrderUnit1		'= 7
Dim C_ReceiptQtyInOrderUnit1	'= 8
Dim C_WaitQtyInBaseUnit1		'= 9
Dim C_BaseUnit1					'= 10
Dim C_ProdQtyInBaseUnit1		'= 11
Dim C_GoodQtyInBaseUnit1		'= 12
Dim C_BadQtyInBaseUnit1			'= 13
Dim C_ReceiptQtyInBaseUnit1		'= 14

'-------------------------------
' Column Constants : Spread 3 
'-------------------------------
Dim C_ProdtOrderNo				'= 1
Dim C_OprNo						'= 2
Dim C_WaitQtyInOrderUnit2		'= 3
Dim C_OrderUnit2				'= 4
Dim C_ProdQtyInOrderUnit2		'= 5
Dim C_GoodQtyInOrderUnit2		'= 6
Dim C_BadQtyInOrderUnit2		'= 7
Dim C_ReceiptQtyInOrderUnit2	'= 8
Dim C_WaitQtyInBaseUnit2		'= 9
Dim C_BaseUnit2					'= 10
Dim C_ProdQtyInBaseUnit2		'= 11
Dim C_GoodQtyInBaseUnit2		'= 12
Dim C_BadQtyInBaseUnit2			'= 13
Dim C_ReceiptQtyInBaseUnit2		'= 14
Dim C_TrackingNo				'= 15
Dim C_SlCd						'= 16
Dim C_SlNm						'= 17

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgStrPrevKey2		
Dim lgStrPrevKey3

Dim lgOldRow1
Dim lgOldRow2
		
Dim lgLngCurRows
Dim lgSortKey 
Dim lgSortKey2 
Dim lgSortKey3 
 
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop					 'Popup
Dim gSelframeFlg

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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgOldRow1 = 0
	lgOldRow2 = 0
	
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call InitSpreadPosVariables(pvSpdNo)
    With frm1
    if pvSpdNo = "*" or pvSpdNo= "A" then
		'-------------------------------------------
		' Spread 1 Setting
		'-------------------------------------------
		ggoSpread.Source = .vspdData1
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.vspdData1.MaxCols = C_ItemGroupNm + 1
		.vspdData1.MaxRows = 0
		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit		C_ItemCd,					"품목", 18
		ggoSpread.SSSetEdit		C_ItemNm,					"품목명", 25
		ggoSpread.SSSetEdit		C_Spec,						"규격", 25
		ggoSpread.SSSetFloat	C_WaitQtyInOrderUnit,		"입고대기",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_OrderUnit,				"오더단위", 8,,,3,2
		ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit,		"실적수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit,		"양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQtyInOrderUnit,		"불량수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_ReceiptQtyInOrderUnit,	"입고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_WaitQtyInBaseUnit,		"입고대기",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_BaseUnit,					"기준단위", 8,,,3,2
		ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit,		"실적수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit,		"양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQtyInBaseUnit,			"불량수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_ReceiptQtyInBaseUnit,		"입고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_ItemGroupCd,				"품목그룹",	15
		ggoSpread.SSSetEdit		C_ItemGroupNm,				"품목그룹명", 30
		'Call ggoSpread.MakePairsColumn(,)
 		Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols ,.vspdData1.MaxCols , True)
		ggoSpread.SSSetSplit2(2)
    end if

    if pvSpdNo = "*" or pvSpdNo= "B" then
		'-------------------------------------------
		' Spread 2 Setting
		'-------------------------------------------
		ggoSpread.Source = .vspdData2
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.vspdData2.MaxCols = C_ReceiptQtyInBaseUnit1 + 1
		.vspdData2.MaxRows = 0
		Call GetSpreadColumnPos("B")
		ggoSpread.SSSetEdit		C_WCCd,						"작업장", 10
		ggoSpread.SSSetEdit		C_WCNm,						"작업장명", 20
		ggoSpread.SSSetFloat	C_WaitQtyInOrderUnit1,		"입고대기",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_OrderUnit1,				"오더단위", 8,,,3,2
		ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit1,		"실적수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit1,		"양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQtyInOrderUnit1,		"불량수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_ReceiptQtyInOrderUnit1,	"입고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_WaitQtyInBaseUnit1,		"입고대기",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_BaseUnit1,				"기준단위", 8,,,3,2
		ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit1,		"실적수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit1,		"양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQtyInBaseUnit1,		"불량수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_ReceiptQtyInBaseUnit1,	"입고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		'Call ggoSpread.MakePairsColumn(,)
 		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols ,.vspdData2.MaxCols , True)
		ggoSpread.SSSetSplit2(1)
    end if

    if pvSpdNo = "*" or pvSpdNo= "C" then
		'-------------------------------------------
		' Spread 3 Setting
		'-------------------------------------------
		ggoSpread.Source = .vspdData3
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.vspdData3.MaxCols = C_SLNm + 1
		.vspdData3.MaxRows = 0
		Call GetSpreadColumnPos("C")
		ggoSpread.SSSetEdit		C_ProdtOrderNo,				"오더번호", 18
		ggoSpread.SSSetEdit		C_OprNo,					"공정", 8,,,3		
		ggoSpread.SSSetFloat	C_WaitQtyInOrderUnit2,		"입고대기",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_OrderUnit2,				"오더단위", 8,,,3,2
		ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit2,		"실적수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit2,		"양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQtyInOrderUnit2,		"불량수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_ReceiptQtyInOrderUnit2,	"입고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_WaitQtyInBaseUnit2,		"입고대기",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_BaseUnit2,				"기준단위", 8,,,3,2
		ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit2,		"실적수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit2,		"양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQtyInBaseUnit2,		"불량수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_ReceiptQtyInBaseUnit2,	"입고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_TrackingNo,				"Tracking No.", 25
		ggoSpread.SSSetEdit		C_SLCd,						"창고", 10
		ggoSpread.SSSetEdit		C_SLNm,						"창고명", 20
		
		'Call ggoSpread.MakePairsColumn(,)
 		Call ggoSpread.SSSetColHidden(.vspdData3.MaxCols ,.vspdData3.MaxCols , True)
		ggoSpread.SSSetSplit2(2)
    end if
		
    End With
    Call SetSpreadLock()
    
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
		
	'-------------------------
	' Set Lock Prop :Spread 1 		
	'-------------------------
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
			
	'-------------------------
	' Set Lock Prop :Spread 2 		
	'-------------------------
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
		
	'-------------------------
	' Set Lock Prop :Spread 3 		
	'-------------------------
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'================================== 2.2.6 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================

Sub SetSpreadColor(ByVal lRow)
	
End Sub
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

End Sub


'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)

    if pvSpdNo = "*" or pvSpdNo= "A" then
		'-------------------------------
		' Column Constants : Spread 1 
		'-------------------------------
		C_ItemCd					= 1
		C_ItemNm					= 2
		C_Spec						= 3
		C_WaitQtyInOrderUnit		= 4
		C_OrderUnit					= 5
		C_ProdQtyInOrderUnit		= 6
		C_GoodQtyInOrderUnit		= 7
		C_BadQtyInOrderUnit			= 8
		C_ReceiptQtyInOrderUnit		= 9
		C_WaitQtyInBaseUnit			= 10
		C_BaseUnit					= 11
		C_ProdQtyInBaseUnit			= 12
		C_GoodQtyInBaseUnit			= 13
		C_BadQtyInBaseUnit			= 14
		C_ReceiptQtyInBaseUnit		= 15
		C_ItemGroupCd				= 16
		C_ItemGroupNm				= 17
    end if

    if pvSpdNo = "*" or pvSpdNo= "B" then
		'-------------------------------
		' Column Constants : Spread 2 
		'-------------------------------
		C_WCCd						= 1
		C_WcNm						= 2
		C_WaitQtyInOrderUnit1		= 3
		C_OrderUnit1				= 4
		C_ProdQtyInOrderUnit1		= 5
		C_GoodQtyInOrderUnit1		= 6
		C_BadQtyInOrderUnit1		= 7
		C_ReceiptQtyInOrderUnit1	= 8
		C_WaitQtyInBaseUnit1		= 9
		C_BaseUnit1					= 10
		C_ProdQtyInBaseUnit1		= 11
		C_GoodQtyInBaseUnit1		= 12
		C_BadQtyInBaseUnit1			= 13
		C_ReceiptQtyInBaseUnit1		= 14
    end if

    if pvSpdNo = "*" or pvSpdNo= "C" then
		'-------------------------------
		' Column Constants : Spread 3 
		'-------------------------------
		C_ProdtOrderNo				= 1
		C_OprNo						= 2
		C_WaitQtyInOrderUnit2		= 3
		C_OrderUnit2				= 4
		C_ProdQtyInOrderUnit2		= 5
		C_GoodQtyInOrderUnit2		= 6
		C_BadQtyInOrderUnit2		= 7
		C_ReceiptQtyInOrderUnit2	= 8
		C_WaitQtyInBaseUnit2		= 9
		C_BaseUnit2					= 10
		C_ProdQtyInBaseUnit2		= 11
		C_GoodQtyInBaseUnit2		= 12
		C_BadQtyInBaseUnit2			= 13
		C_ReceiptQtyInBaseUnit2		= 14
		C_TrackingNo				= 15
		C_SlCd						= 16
		C_SlNm						= 17
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
  		C_ItemCd					= iCurColumnPos(1)
		C_ItemNm					= iCurColumnPos(2)
		C_Spec						= iCurColumnPos(3)
		C_WaitQtyInOrderUnit		= iCurColumnPos(4)
		C_OrderUnit					= iCurColumnPos(5)
		C_ProdQtyInOrderUnit		= iCurColumnPos(6)
		C_GoodQtyInOrderUnit		= iCurColumnPos(7)
		C_BadQtyInOrderUnit			= iCurColumnPos(8)
		C_ReceiptQtyInOrderUnit		= iCurColumnPos(9)
		C_WaitQtyInBaseUnit			= iCurColumnPos(10)
		C_BaseUnit					= iCurColumnPos(11)
		C_ProdQtyInBaseUnit			= iCurColumnPos(12)
		C_GoodQtyInBaseUnit			= iCurColumnPos(13)
		C_BadQtyInBaseUnit			= iCurColumnPos(14)
		C_ReceiptQtyInBaseUnit		= iCurColumnPos(15)
		C_ItemGroupCd				= iCurColumnPos(16)
		C_ItemGroupNm				= iCurColumnPos(17)
				
	Case "B"
 		ggoSpread.Source = frm1.vspdData2
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_WCCd						= iCurColumnPos(1)
		C_WcNm						= iCurColumnPos(2)
		C_WaitQtyInOrderUnit1		= iCurColumnPos(3)
		C_OrderUnit1				= iCurColumnPos(4)
		C_ProdQtyInOrderUnit1		= iCurColumnPos(5)
		C_GoodQtyInOrderUnit1		= iCurColumnPos(6)
		C_BadQtyInOrderUnit1		= iCurColumnPos(7)
		C_ReceiptQtyInOrderUnit1	= iCurColumnPos(8)
		C_WaitQtyInBaseUnit1		= iCurColumnPos(9)
		C_BaseUnit1					= iCurColumnPos(10)
		C_ProdQtyInBaseUnit1		= iCurColumnPos(11)
		C_GoodQtyInBaseUnit1		= iCurColumnPos(12)
		C_BadQtyInBaseUnit1			= iCurColumnPos(13)
		C_ReceiptQtyInBaseUnit1		= iCurColumnPos(14)
		
	Case "C"
 		ggoSpread.Source = frm1.vspdData3 
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_ProdtOrderNo				= iCurColumnPos(1)
		C_OprNo						= iCurColumnPos(2)
		C_WaitQtyInOrderUnit2		= iCurColumnPos(3)
		C_OrderUnit2				= iCurColumnPos(4)
		C_ProdQtyInOrderUnit2		= iCurColumnPos(5)
		C_GoodQtyInOrderUnit2		= iCurColumnPos(6)
		C_BadQtyInOrderUnit2		= iCurColumnPos(7)
		C_ReceiptQtyInOrderUnit2	= iCurColumnPos(8)
		C_WaitQtyInBaseUnit2		= iCurColumnPos(9)
		C_BaseUnit2					= iCurColumnPos(10)
		C_ProdQtyInBaseUnit2		= iCurColumnPos(11)
		C_GoodQtyInBaseUnit2		= iCurColumnPos(12)
		C_BadQtyInBaseUnit2			= iCurColumnPos(13)
		C_ReceiptQtyInBaseUnit2		= iCurColumnPos(14)
		C_TrackingNo				= iCurColumnPos(15)
		C_SlCd						= iCurColumnPos(16)
		C_SlNm						= iCurColumnPos(17)
 
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

'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명 
	
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

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분: From To를 입력할 것 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
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

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "CL"
	arrParam(4) = "RLST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
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
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
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
	frm1.txtItemGroupCd.focus
 
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
    If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""	

    iCalledAspName = AskPRAspName("p4600pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4600pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
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

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSLCd()

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

	arrParam(0) = "창고팝업"											' 팝업 명칭 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtSLCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "창고"												' TextBox 명칭 
	
   	arrField(0) = "SL_CD"													' Field명(0)
   	arrField(1) = "SL_NM"													' Field명(1)
    
   	arrHeader(0) = "창고"												' Header명(0)
   	arrHeader(1) = "창고명"												' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtSLCd.focus
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
	End With
End Function

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function    

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
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
	
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
	Call InitSpreadSheet("*")														'⊙: Setup the Spread sheet
	
	Call InitComboBox
	Call SetDefaultVal
	Call InitVariables																'⊙: Initializes local global variables
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolbar("11000000000011")
		
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtItemCd.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
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
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

  	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
  	
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
  		
  		lgOldRow2 = 0
			
		lgStrPrevKey2 = ""
		lgStrPrevKey3 = ""
		frm1.vspdData2.MaxRows = 0
		frm1.vspdData3.MaxRows = 0
			
		frm1.vspdData1.Col = C_ItemCd
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
			
		frm1.KeyItemCd.value =  Trim(frm1.vspdData1.Text)
			
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
					
		lgOldRow1 = frm1.vspdData1.Row 
  		
 	Else
  		'------ Developer Coding part (Start)
		If lgOldRow1 <> Row Then
			
			lgOldRow2 = 0
			
			lgStrPrevKey2 = ""
			lgStrPrevKey3 = ""
			frm1.vspdData2.MaxRows = 0
			frm1.vspdData3.MaxRows = 0
			
			frm1.vspdData1.Col = C_ItemCd
			frm1.vspdData1.Row = Row
			
			frm1.KeyItemCd.value =  Trim(frm1.vspdData1.Text)
			
			If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
					
			lgOldRow1 = Row
		End If
	 	'------ Developer Coding part (End)
 	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
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

  		lgStrPrevKey3 = ""
			
		frm1.vspdData3.MaxRows = 0

		frm1.vspdData2.Row = frm1.vspdData2.ActiveRow		
		frm1.vspdData2.Col = C_WCCd
		frm1.KeyWcCd.value  = Trim(frm1.vspdData2.Text) 
			
		If DbQuery3 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
			
		lgOldRow2 = Row
 	Else
  		'------ Developer Coding part (Start)
		If lgOldRow2 <> Row Then
			
			lgStrPrevKey3 = ""
			
			frm1.vspdData3.MaxRows = 0

			frm1.vspdData2.Row = Row		
			frm1.vspdData2.Col = C_WCCd
			frm1.KeyWcCd.value  = Trim(frm1.vspdData2.Text) 
			
			If DbQuery3 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
			
			lgOldRow2 = Row
		End If
		'------ Developer Coding part (End)
  	End If

End Sub

'==========================================================================================
'   Event Name : vspdData3_Click
'   Event Desc :
'==========================================================================================
Sub vspdData3_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 

  	gMouseClickStatus = "SP3C"   
     
  	Set gActiveSpdSheet = frm1.vspdData3
     
  	If frm1.vspdData3.MaxRows = 0 Then
  		Exit Sub
  	End If
  	
  	If Row <= 0 Then
  		ggoSpread.Source = frm1.vspdData3 
  		If lgSortKey3 = 1 Then
  			ggoSpread.SSSort Col					'Sort in Ascending
  			lgSortKey3 = 2
  		Else
  			ggoSpread.SSSort Col, lgSortKey3		'Sort in Descending
  			lgSortKey3 = 1
  		End If
 	Else
  		'------ Developer Coding part (Start)
 	 	'------ Developer Coding part (End)
 	
  	End If

End Sub


'==========================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc :
'==========================================================================================

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData1
		
		ggoSpread.Source = frm1.vspdData1
		
		.Row = Row
		.Col = C_Select
		
		If ButtonDown = 1 Then
			ggoSpread.UpdateRow Row
		Else
			ggoSpread.SSDeleteFlag Row,Row
		End If			
	
	End With
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
		If lgStrPrevKey1 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If		
		End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub
'==========================================================================================
'   Event Name : vspdData3_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then
		If lgStrPrevKey3 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery3 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
 '========================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub 

Sub vspdData2_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SP2C" Then
        gMouseClickStatus = "SP2CR"
     End If
End Sub 

Sub vspdData3_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SP3C" Then
        gMouseClickStatus = "SP3CR"
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
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData3
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
  
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
' 	If NewCol = C_WCCd or Col = C_WCCd Then
' 		Cancel = True
' 		Exit Sub
' 	End If
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
' 	If NewCol = C_WCCd or Col = C_WCCd Then
 '		Cancel = True
 '		Exit Sub
  '	End If
	ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub 

Sub vspdData3_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("C")
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
    select case pvSpdNo
    case "A"
		ggoSpread.Source = frm1.vspdData1
    case "B"
		ggoSpread.Source = frm1.vspdData2
    case "C"
		ggoSpread.Source = frm1.vspdData3
    end select
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
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	'If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function	
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
   
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		 Call RestoreToolBar()
		 Exit Function
	End If	 																'☜: Query db data
       
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
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	On Error Resume Next    
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
	If frm1.vspdData1.MaxRows <= 0 Then Exit Function	

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.EditUndo            
	
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
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
    Call parent.FncExport(parent.C_SINGLEMULTI)									'☜: 화면 유형 
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
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                               '☜:화면 유형, Tab 유무 
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
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : Spread 1 조회 및 Scroll
'========================================================================================
Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)
		strVal = strVal & "&txtSlCd=" & Trim(frm1.hSlCd.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
	Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)
		strVal = strVal & "&txtSlCd=" & Trim(frm1.txtSlCd.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
	End If    

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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
   
		frm1.vspdData1.Col = C_ItemCd
		frm1.vspdData1.Row = 1
		frm1.KeyItemCd.value = Trim(frm1.vspdData1.Text)
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		If DbQuery2 = False Then
			 Call RestoreToolBar()	
			 Exit Function 
        End If  
		
		lgOldRow1 = 1
		
    End If

    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field
    
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 
'========================================================================================

Function DbQuery2() 
    
    DbQuery2 = False                                    
    
    Call LayerShowHide(1)
 
    Dim strVal
	
	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'☜: 조회 조건 데이타 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.KeyItemCd.value)			'☜: 조회 조건 데이타 
	strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)					'☜: 조회 조건 데이타 
	strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'☜: 조회 조건 데이타 
	strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.hProdOrderNo.value)	'☜: 조회 조건 데이타 
	strVal = strVal & "&txtSlCd=" & Trim(frm1.hSlCd.value)					'☜: 조회 조건 데이타 

    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery2 = True                                                          	

End Function

'========================================================================================
' Function Name : DbQuery2Ok
' Function Desc : Spread 2 And Spread 3 Data 조회 
'========================================================================================

Function DbQuery2Ok() 
	
	frm1.vspdData2.Col = C_WcCd
	frm1.vspdData2.Row = 1
	lgOldRow2 = 1
	
	frm1.KeyWcCd.value = Trim(frm1.vspdData2.Text)
	
	If DbQuery3 = False Then
			Call RestoreToolBar()	
            Exit Function 
    End If  
	
End Function

'========================================================================================
' Function Name : DbQuery3
' Function Desc : Spread 2 Scroll 
'========================================================================================

Function DbQuery3() 
    
    DbQuery3 = False                                    
    
    Call LayerShowHide(1)
 
    Dim strVal
	
	strVal = BIZ_PGM_QRY3_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'☜: 조회 조건 데이타 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.KeyItemCd.value)			'☜: 조회 조건 데이타 
	strVal = strVal & "&txtWcCd=" & Trim(frm1.KeyWcCd.value)				'☜: 조회 조건 데이타 
	strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'☜: 조회 조건 데이타 
	strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.hProdOrderNo.value)	'☜: 조회 조건 데이타 
	strVal = strVal & "&txtSlCd=" & Trim(frm1.hSlCd.value)					'☜: 조회 조건 데이타 
    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery3 = True                                                          	

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>입고대기현황조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="라우팅명"></TD>
			 						<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>제조오더 번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더 번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>							
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>입고창고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="입고창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSlCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=20 tag="14" ALT="입고창고명"></TD>								
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWCCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()"> <INPUT TYPE=TEXT ID="txtWCNm" NAME="arrCond" tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p4514ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4514ma1_B_vspdData2.js'></script>
								</TD>
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4514ma1_C_vspdData3.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hSlCd" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="KeyWcCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
