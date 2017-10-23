
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 공정별계획관리 
'*  3. Program ID           : p4114ma1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/05/10
'*  8. Modified date(Last)  : 2003/05/20
'*  9. Modifier (First)     : Chen, Jaehyun
'* 10. Modifier (Last)      : Chen, Jaehyun
'* 11. Comment              :
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit								'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID		= "p4114mb1.asp"								'☆: Query Order Header
Const BIZ_PGM_QRY2_ID		= "p4114mb2.asp"								'☆: Query Order Detail
Const BIZ_PGM_SAVE_ID		= "p4114mb3.asp"                                '☆: Manage Order Detail
Const BIZ_PGM_SAVE2_ID		= "p4114mb5.asp"                                '☆: Manage Order Detail
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Operation 
Dim C_OdrNo					'= 1
Dim C_ProductCd	    		'= 2
Dim C_ProductCdNm			'= 3
Dim C_Spec					'= 4
Dim C_RoutNo				'= 5
Dim C_PlanStartDt1			'= 6
Dim C_PlanEndDt1			'= 7
Dim C_OrderQuantity			'= 8
Dim C_OrderUnit				'= 9
Dim C_TrackingNo			'= 10
Dim C_OrderStatus			'= 11
Dim C_OrderStatusNm			'= 12
Dim C_OrderSelect			'= 13
Dim C_OrderSelectNm			'= 14
Dim C_ItemGroupCd			'= 15
Dim C_ItemGroupNm			'= 16

' Grid 2(vspdData2) - Operation 
Dim C_OperationCd2			'= 1
Dim C_WCCd2					'= 2
Dim C_WCCdPopup2			'= 3
Dim C_WCNm2					'= 4
Dim C_WorkCd2				'= 5
Dim C_WorkNm2				'= 6
Dim C_PlanStartDt2			'= 7
Dim C_PlanEndDt2		    '= 8
Dim C_BpCd2					'= 9
Dim C_BpCdPopup2			'= 10
Dim C_BpNm2					'= 11
Dim C_CCFCost2				'= 12
Dim C_CCFAmt2				'= 13
Dim C_Currency2				'= 14
Dim C_CurrencyPopup			'= 15
Dim C_Tax2					'= 16
Dim C_TaxPopup				'= 17
Dim C_RoutOrder2			'= 18
Dim C_MilestoneFlg2			'= 19
Dim C_InspFlg2				'= 20
' Hidden
Dim C_InsideFlg2			'= 21
Dim C_ProdtOrderNo2			'= 22
Dim C_OrderStatus2			' =23

' Grid 3(vspddata3) - Hidden 
Dim C_OperationCd3			'= 1
Dim C_WCCd3					'= 2
Dim C_WCCdPopup3			'= 3
Dim C_WCNm3					'= 4
Dim C_WorkCd3				'= 5
Dim C_WorkNm3				'= 6
Dim C_PlanStartDt3			'= 7
Dim C_PlanEndDt3			'= 8
Dim C_BpCd3					'= 9
Dim C_Dummy1   				'= 10
Dim C_BpNm3					'= 11
Dim C_CCFCost3				'= 12
Dim C_CCFAmt3				'= 13
Dim C_Currency3				'= 14
Dim C_Dummy2				'= 15
Dim C_Tax3					'= 16
Dim C_Dummy3				'= 17
Dim C_RoutOrder3			'= 18
Dim C_MilestoneFlg3			'= 19
Dim C_InspFlg3				'= 20
' Hidden
Dim C_InsideFlg3			'= 21
Dim C_ProdtOrderNo3			'= 22
Dim C_OrderStatus3			'= 23

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size를 조사할 변수 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgIntPrevKey
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgCurrRow

Dim lgSortKey1
Dim lgSortKey2

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgRow         
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""							'initializes Previous Key
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgRow = 0
    
    lgSortKey1 = 1
    lgSortKey2 = 1

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

	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"

	frm1.txtProdFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -10, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtProdToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 20, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	
	Call InitSpreadPosVariables(pvSpdNo)
	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		
		With frm1.vspdData1 
		
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021206", , Parent.gAllowDragDropSpread
	
			.ReDraw = false
	
			.MaxCols = C_ItemGroupNm +1
			.MaxRows = 0	
			
			Call GetSpreadColumnPos("A")
		
			ggoSpread.SSSetEdit		C_OdrNo, "오더번호", 18,,,,2
			ggoSpread.SSSetEdit		C_ProductCd, "품목", 18,,,,2
			ggoSpread.SSSetEdit		C_ProductCdNm, "품목명", 25	
			ggoSpread.SSSetEdit		C_Spec,	"규격", 25	
			ggoSpread.SSSetEdit		C_RoutNo, "라우팅",8,,,,2				
			ggoSpread.SSSetDate 	C_PlanStartDt1, "착수예정일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_PlanEndDt1, "완료예정일", 11, 2, parent.gDateFormat	
			ggoSpread.SSSetFloat	C_OrderQuantity, "오더수량", 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"    
			ggoSpread.SSSetEdit		C_OrderUnit, "오더단위", 8	
			ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.",25
			ggoSpread.SSSetEdit		C_OrderStatus, "지시상태", 8
			ggoSpread.SSSetEdit		C_OrderStatusNm, "지시상태", 10
			ggoSpread.SSSetEdit		C_OrderSelect, "지시구분", 8
			ggoSpread.SSSetEdit		C_OrderSelectNm, "지시구분", 10	
			ggoSpread.SSSetEdit 	C_ItemGroupCd, "품목그룹",	15
			ggoSpread.SSSetEdit		C_ItemGroupNm, "품목그룹명", 30
	
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
 			Call ggoSpread.SSSetColHidden( C_OrderStatus, C_OrderStatus, True)
			Call ggoSpread.SSSetColHidden( C_OrderSelect, C_OrderSelect, True)
	
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("A")
			
			.ReDraw = true
			
		End With
		
	End If
	
	'------------------------------------------
	' Grid 2 - Component Spread Setting
	'------------------------------------------
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
	
		With frm1.vspdData2

			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20060809", , Parent.gAllowDragDropSpread	
			.ReDraw = false
			
			.MaxCols = C_OrderStatus2 +1										'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit		C_OperationCd2, "공정", 8	
			ggoSpread.SSSetEdit		C_WCCd2, "작업장", 12,,,7,2
			ggoSpread.SSSetButton 	C_WCCdPopup2
			ggoSpread.SSSetEdit		C_WCNm2, "작업장명", 20
			ggoSpread.SSSetCombo	C_WorkCd2, "작업", 8
			ggoSpread.SSSetCombo	C_WorkNm2, "작업명", 20
			ggoSpread.SSSetDate 	C_PlanStartDt2, "착수예정일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_PlanEndDt2, "완료예정일", 11, 2, parent.gDateFormat	
			ggoSpread.SSSetEdit		C_BpCd2, "외주처", 10,,,10,2
			ggoSpread.SSSetButton 	C_BpCdPopup2
			ggoSpread.SSSetEdit		C_BpNm2, "외주처명", 20
			ggoSpread.SSSetCombo	C_MilestoneFlg2, "Milestone", 8
			ggoSpread.SSSetCombo	C_InspFlg2, "공정검사여부", 12
			ggoSpread.SSSetFloat	C_CCFCost2,"외주단가", 15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_CCFAmt2,"외주금액", 15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_Currency2, "통화", 7,,,7,2
			ggoSpread.SSSetButton	C_CurrencyPopup	
			ggoSpread.SSSetEdit 	C_Tax2,"VAT유형", 8,,,7,2
			ggoSpread.SSSetButton	C_TaxPopup
			ggoSpread.SSSetEdit		C_RoutOrder2, "공정순서", 6
			ggoSpread.SSSetEdit		C_InsideFlg2, "사내/외", 6	
			ggoSpread.SSSetEdit		C_ProdtOrderNo2, "오더번호", 18	
			ggoSpread.SSSetEdit		C_OrderStatus2, "지시상태", 6
			
			
			Call ggoSpread.MakePairsColumn(C_WCCd2, C_WCCdPopup2)
			Call ggoSpread.MakePairsColumn(C_BpCd2, C_BpCdPopup2)
			Call ggoSpread.MakePairsColumn(C_Currency2, C_CurrencyPopup)
			Call ggoSpread.MakePairsColumn(C_Tax2, C_TaxPopup)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden( C_RoutOrder2, C_RoutOrder2, True)
 			Call ggoSpread.SSSetColHidden( C_InsideFlg2, C_OrderStatus2, True)
			
			ggoSpread.SSSetSplit2(3)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true
    
		End With
		
	End If	
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
	
		With frm1.vspdData3
			
			.MaxCols = C_OrderStatus3 +1										'☜: 최대 Columns의 항상 1개 증가시킴 

			.MaxRows = 0
			ggoSpread.Source = frm1.vspdData3

			.ReDraw = false
			ggoSpread.Spreadinit
			ggoSpread.SSSetEdit		C_OperationCd3, "공정", 12	
			ggoSpread.SSSetEdit		C_WCCd3, "작업장", 8,,,7
			ggoSpread.SSSetButton 	C_WCCdPopup3
			ggoSpread.SSSetEdit		C_WCNm3, "작업장명", 20
			ggoSpread.SSSetEdit		C_WorkCd3, "작업", 20
			ggoSpread.SSSetEdit		C_WorkNm3, "작업명", 20
			ggoSpread.SSSetDate 	C_PlanStartDt3, "착수예정일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_PlanEndDt3, "완료예정일", 11, 2, parent.gDateFormat	
			ggoSpread.SSSetEdit		C_BpCd3, "외주처", 12,,,12
			ggoSpread.SSSetButton 	C_Dummy1
			ggoSpread.SSSetEdit		C_BpNm3, "외주처명", 20
			ggoSpread.SSSetFloat	C_CCFCost3,"외주단가", 15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_CCFAmt3,"외주금액", 15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"			
			ggoSpread.SSSetEdit 	C_Currency3, "통화", 7,,,7,2
			ggoSpread.SSSetButton 	C_Dummy2
			ggoSpread.SSSetEdit 	C_Tax3,"세금", 12
			ggoSpread.SSSetButton 	C_Dummy3
			ggoSpread.SSSetEdit 	C_RoutOrder3,"공정순서", 6
			ggoSpread.SSSetEdit		C_MilestoneFlg3, "Milestone", 8
			ggoSpread.SSSetEdit		C_InspFlg3, "검사", 8
			ggoSpread.SSSetEdit 	C_InsideFlg3,"사내/외", 6
			ggoSpread.SSSetEdit		C_ProdtOrderNo3, "오더번호", 18
			ggoSpread.SSSetEdit		C_OrderStatus3, "지시상태", 6
			
			Call SetSpreadLock("C")
			
			.ReDraw = true
    
		End With
	End If	
	
End Sub

'2008-05-19 5:40오후 :: hanc
Function RegProd() 

    Dim IntRows 
    Dim strVal  
	
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

    RegProd = False
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
    	frm1.vspdData1.Row =frm1.vspdData1.ActiveRow
    	frm1.vspdData1.Col = C_OdrNo
		.txtOpr.value   =   frm1.vspdData1.Text 
	End With

  	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE2_ID)

    RegProd = True
    
End Function

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    With frm1
		If pvSpdNo = "A" Then
			'--------------------------------
			'Grid 1
			'--------------------------------
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		'--------------------------------
		'Grid 2
		'--------------------------------
		If pvSpdNo = "B" Then
			ggoSpread.Source = frm1.vspdData2

			.vspdData2.ReDraw = False
			ggoSpread.SpreadLock C_OperationCd2,-1,C_OperationCd2
			ggoSpread.SpreadUnLock C_WCCd2,-1,C_WCCd2
			ggoSpread.SpreadUnLock C_WCCdPopup2,-1,C_WCCdPopup2
			ggoSpread.SpreadLock C_WCNm2,-1,C_WCNm2
			ggoSpread.SpreadLock C_BpCd2,-1,C_BpCd2
			ggoSpread.SpreadLock C_BpNm2,-1,C_BpNm2
			ggoSpread.SpreadLock C_CCFCost2,-1,C_CCFCost2
			ggoSpread.SpreadLock C_CCFAmt2,-1,C_CCFAmt2
			ggoSpread.SpreadLock C_Currency2,-1,C_Currency2
			ggoSpread.SpreadLock C_Tax2,-1,C_Tax2  
			ggoSpread.SpreadLock C_ProdtOrderNo2,-1,C_ProdtOrderNo2
			ggoSpread.SpreadLock C_OrderStatus2,-1,C_OrderStatus2
			ggoSpread.SpreadLock frm1.vspdData2.MaxCols, -1, frm1.vspdData2.MaxCols

			ggoSpread.SSSetRequired	 C_WCCd2,				-1
			ggoSpread.SSSetRequired	 C_PlanStartDt2,		-1
			ggoSpread.SSSetRequired  C_PlanEndDt2,			-1
			ggoSpread.SSSetRequired  C_MilestoneFlg2,		-1
			ggoSpread.SSSetRequired  C_InspFlg2,			-1

			.vspdData2.Redraw = True
		End If
		'--------------------------------
		'Grid 3
		'--------------------------------
		If pvSpdNo = "C" Then
			ggoSpread.Source = frm1.vspdData3

			.vspdData3.ReDraw = False
			ggoSpread.SpreadLock C_OperationCd3,-1,C_OperationCd3
			ggoSpread.SpreadUnLock C_WCCd3,-1,C_WCCd3
			ggoSpread.SpreadUnLock C_WCCdPopup3,-1,C_WCCdPopup3
			ggoSpread.SpreadLock C_WCNm3,-1,C_WCNm3
			ggoSpread.SpreadLock C_BpCd3,-1,C_BpCd3
			ggoSpread.SpreadLock C_BpNm3,-1,C_BpNm3
			ggoSpread.SpreadLock C_CCFCost3,-1,C_CCFCost3
			ggoSpread.SpreadLock C_CCFAmt3,-1,C_CCFAmt3
			ggoSpread.SpreadLock C_Currency3,-1,C_Currency3
			ggoSpread.SpreadLock C_Tax3,-1,C_Tax3  
			ggoSpread.SpreadLock C_ProdtOrderNo3,-1,C_ProdtOrderNo3
			ggoSpread.SpreadLock C_OrderStatus3,-1,C_OrderStatus3
			ggoSpread.SpreadLock frm1.vspdData3.MaxCols, -1, frm1.vspdData3.MaxCols

			ggoSpread.SSSetRequired	 C_WCCd3,				-1
			ggoSpread.SSSetRequired	 C_PlanStartDt3,		-1
			ggoSpread.SSSetRequired  C_PlanEndDt3,			-1
			ggoSpread.SSSetRequired  C_MilestoneFlg3,		-1
			ggoSpread.SSSetRequired  C_InspFlg3,			-1

			.vspdData3.Redraw = True
		End If
		
    End With

End Sub


'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData2
    
		.Redraw = False
		ggoSpread.Source = frm1.vspdData2
		'ggoSpread.SSSetRequired	 C_OperationCd2,		pvStartRow, pvEndRow    
		'ggoSpread.SSSetRequired	 C_WCCd2,				pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected C_WCNm2,				pvStartRow, pvEndRow
		'ggoSpread.SSSetRequired	 C_PlanStartDt2,		pvStartRow, pvEndRow
		'ggoSpread.SSSetRequired  C_PlanEndDt2,			pvStartRow, pvEndRow 
		'ggoSpread.SSSetRequired  C_MilestoneFlg2,		pvStartRow, pvEndRow
		'ggoSpread.SSSetRequired  C_InspFlg2,			pvStartRow, pvEndRow
		
		ggoSpread.SSSetRequired C_OperationCd2,			pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected C_WCCd2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_WCNm2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_WorkCd2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_WorkNm2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanStartDt2,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanEndDt2,			pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_MilestoneFlg2,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspFlg2,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BpCd2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BpNm2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CCFCost2,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CCFAmt2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Currency2,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Tax2,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_C_ProdtOrderNo2,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OrderStatus2,		pvStartRow, pvEndRow
		
		.Redraw = True
    
    End With

End Sub



'============================= 2.2.5 SetSpread2Color() ===================================
' Function Name : SetSpread2Color
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpread2Color(ByVal LngRow)
	
	ggoSpread.Source = frm1.vspdData2
	
	With frm1.vspdData2

		.Row = LngRow
		.Col = C_OrderStatus2
				
		If UCase(Trim(.Text)) = "" Or UCase(Trim(.Text)) = "OP" Or UCase(Trim(.Text)) = "RL" Then
			
			ggoSpread.SpreadLock	 C_OperationCd2,		LngRow, C_OperationCd2, LngRow
			ggoSpread.SpreadUnLock	 C_WorkCd2,				LngRow, C_WorkNm2, LngRow
			ggoSpread.SpreadUnLock	 C_WCCd2,				LngRow, C_WCCd2, LngRow
			ggoSpread.SSSetRequired	 C_WCCd2,				LngRow, LngRow
			ggoSpread.SpreadUnLock	 C_PlanStartDt2,		LngRow, C_PlanStartDt2, LngRow
			ggoSpread.SSSetRequired	 C_PlanStartDt2,		LngRow, LngRow
			ggoSpread.SpreadUnLock	 C_PlanEndDt2,			LngRow, C_PlanEndDt2, LngRow
			ggoSpread.SSSetRequired  C_PlanEndDt2,			LngRow, LngRow
			ggoSpread.SpreadUnLock	 C_MilestoneFlg2,		LngRow, C_MilestoneFlg2, LngRow
			ggoSpread.SSSetRequired  C_MilestoneFlg2,		LngRow, LngRow
			ggoSpread.SpreadUnLock	 C_InspFlg2,			LngRow, C_InspFlg2, LngRow
			ggoSpread.SSSetRequired  C_InspFlg2,			LngRow, LngRow
					
			.Col = C_InsideFlg2
			If .Text = "Y" Then
				ggoSpread.SSSetProtected C_BpCd2,			LngRow, LngRow
				ggoSpread.SSSetProtected C_BpNm2,			LngRow, LngRow
				ggoSpread.SSSetProtected C_CCFAmt2,			LngRow, LngRow
				ggoSpread.SSSetProtected C_Currency2,		LngRow, LngRow
				ggoSpread.SSSetProtected C_Tax2,			LngRow, LngRow
				ggoSpread.SpreadLock C_BpCdPopup2,			LngRow, C_BpCdPopup2, LngRow
				ggoSpread.SpreadLock C_CurrencyPopup,		LngRow, C_CurrencyPopup, LngRow
				ggoSpread.SpreadLock C_TaxPopup,			LngRow, C_TaxPopup, LngRow
			Else
			    ggoSpread.SpreadUnLock C_BpCd2,				LngRow, C_BpCd2, LngRow
				ggoSpread.SpreadUnLock C_CCFCost2,			LngRow, C_CCFCost2, LngRow
				ggoSpread.SpreadUnLock C_CCFAmt2,			LngRow, C_CCFAmt2, LngRow
				ggoSpread.SpreadUnLock C_Currency2,			LngRow, C_Currency2, LngRow
				ggoSpread.SpreadUnLock C_Tax2,				LngRow, C_Tax2, LngRow
							
				ggoSpread.SSSetRequired C_BpCd2,			LngRow, LngRow
				ggoSpread.SSSetRequired C_CCFCost2,			LngRow, LngRow
				ggoSpread.SSSetRequired C_CCFAmt2,			LngRow, LngRow
				ggoSpread.SSSetRequired C_Currency2,		LngRow, LngRow
				ggoSpread.SSSetRequired C_Tax2,				LngRow, LngRow
				ggoSpread.SpreadUnLock C_BpCdPopup2,		LngRow, C_BpCdPopup2, LngRow
				ggoSpread.SpreadUnLock C_CurrencyPopup,		LngRow, C_CurrencyPopup, LngRow
				ggoSpread.SpreadUnLock C_TaxPopup,			LngRow, C_TaxPopup, LngRow
			End If

			.Col = C_RoutOrder2
	
			If .Text = "S" Or .Text = "L" Then
				ggoSpread.SSSetProtected C_InspFlg2,		LngRow, LngRow
				ggoSpread.SSSetProtected C_MilestoneFlg2,	LngRow, LngRow
			Else
				ggoSpread.SSSetRequired C_InspFlg2,			LngRow, LngRow
				ggoSpread.SSSetRequired C_MilestoneFlg2,	LngRow, LngRow
			End If
					
		Else
					
			ggoSpread.SpreadLock C_WCCd2,			LngRow, C_WCCd2, LngRow
			ggoSpread.SpreadLock C_WCCdPopup2,		LngRow, C_WCCdPopup2, LngRow
			ggoSpread.SpreadLock C_WorkCd2,		LngRow, C_WorkCd2, LngRow
			ggoSpread.SpreadLock C_WorkNm2,		LngRow, C_WorkNm2, LngRow

			ggoSpread.SSSetProtected	C_WCCd2,			LngRow, LngRow
			ggoSpread.SSSetProtected	C_PlanStartDt2,		LngRow, LngRow
			ggoSpread.SSSetProtected	C_PlanEndDt2,		LngRow, LngRow
			ggoSpread.SSSetProtected	C_MilestoneFlg2,	LngRow, LngRow
			ggoSpread.SSSetProtected	C_InspFlg2,			LngRow, LngRow
					
		End If
		
	End With

End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderType, lgF0, lgF1, Chr(11))
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & "  AND MINOR_CD <> 'CL' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderStatus, lgF0, lgF1, Chr(11))
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboJobCd, lgF0, lgF1, Chr(11))
	
	frm1.cboOrderType.value = ""
    frm1.cboOrderStatus.value = ""
    frm1.cboJobCd.value = ""
    
End Sub

'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()

    Dim strVal	
	Dim strCboCd
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	'****************************
	'List Milestone & Insp. Flag
	'****************************
	strCboCd =  "N" & vbTab & "Y"
	
	'****************************
	'List Minor code(Job Code)
	'****************************
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo strCboCd, C_MilestoneFlg2
	ggoSpread.SetCombo strCboCd, C_InspFlg2
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_WorkCd2
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_WorkNm2
    
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData2
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_WorkCd2
			intIndex = .value
			.Col = C_WorkNm2
			.value = intindex
		Next	
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)	
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Operation 
		C_OdrNo				= 1
		C_ProductCd	    	= 2
		C_ProductCdNm		= 3
		C_Spec				= 4
		C_RoutNo			= 5
		C_PlanStartDt1		= 6
		C_PlanEndDt1		= 7
		C_OrderQuantity		= 8
		C_OrderUnit			= 9
		C_TrackingNo		= 10
		C_OrderStatus       = 11
		C_OrderStatusNm     = 12
		C_OrderSelect		= 13
		C_OrderSelectNm		= 14
		C_ItemGroupCd		= 15
		C_ItemGroupNm		= 16
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Operation 
		C_OperationCd2		= 1
		C_WCCd2				= 2
		C_WCCdPopup2		= 3
		C_WCNm2				= 4
		C_WorkCd2           = 5
		C_WorkNm2			= 6
		C_PlanStartDt2		= 7
		C_PlanEndDt2		= 8
		C_BpCd2             = 9
		C_BpCdPopup2		= 10
		C_BpNm2				= 11
		C_CCFCost2			= 12
		C_CCFAmt2			= 13
		C_Currency2			= 14
		C_CurrencyPopup		= 15
		C_Tax2				= 16
		C_TaxPopup			= 17
		C_RoutOrder2		= 18
		C_MilestoneFlg2		= 19
		C_InspFlg2			= 20
		' Hidden
		C_InsideFlg2		= 21
		C_ProdtOrderNo2		= 22
		C_OrderStatus2		= 23
	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then	
		' Grid 3(vspddata3) - Hidden 
		C_OperationCd3		= 1
		C_WCCd3				= 2
		C_WCCdPopup3		= 3
		C_WCNm3				= 4
		C_WorkCd3           = 5
		C_WorkNm3			= 6
		C_PlanStartDt3		= 7
		C_PlanEndDt3		= 8
		C_BpCd3             = 9
		C_Dummy1   			= 10
		C_BpNm3				= 11
		C_CCFCost3			= 12
		C_CCFAmt3			= 13
		C_Currency3			= 14
		C_Dummy2			= 15
		C_Tax3				= 16
		C_Dummy3			= 17
		C_RoutOrder3		= 18
		C_MilestoneFlg3		= 19
		C_InspFlg3			= 20
		C_InsideFlg3		= 21
		C_ProdtOrderNo3		= 22
		C_OrderStatus3		= 23
	End If
	
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
 
			C_OdrNo				= iCurColumnPos(1)
			C_ProductCd	    	= iCurColumnPos(2)
			C_ProductCdNm		= iCurColumnPos(3)
			C_Spec				= iCurColumnPos(4)
			C_RoutNo			= iCurColumnPos(5)
			C_PlanStartDt1		= iCurColumnPos(6)
			C_PlanEndDt1		= iCurColumnPos(7)
			C_OrderQuantity		= iCurColumnPos(8)
			C_OrderUnit			= iCurColumnPos(9)
			C_TrackingNo		= iCurColumnPos(10)
			C_OrderStatus       = iCurColumnPos(11)
			C_OrderStatusNm     = iCurColumnPos(12)
			C_OrderSelect		= iCurColumnPos(13)
			C_OrderSelectNm		= iCurColumnPos(14)
			C_ItemGroupCd		= iCurColumnPos(15)
			C_ItemGroupNm		= iCurColumnPos(16)

		Case "B"	
			ggoSpread.Source = frm1.vspdData2
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
			C_OperationCd2		= iCurColumnPos(1)
			C_WCCd2				= iCurColumnPos(2)
			C_WCCdPopup2		= iCurColumnPos(3)
			C_WCNm2				= iCurColumnPos(4)
			C_WorkCd2           = iCurColumnPos(5)
			C_WorkNm2			= iCurColumnPos(6)
			C_PlanStartDt2		= iCurColumnPos(7)
			C_PlanEndDt2		= iCurColumnPos(8)
			C_BpCd2             = iCurColumnPos(9)
			C_BpCdPopup2		= iCurColumnPos(10)
			C_BpNm2				= iCurColumnPos(11)
			C_CCFCost2			= iCurColumnPos(12)
			C_CCFAmt2			= iCurColumnPos(13)
			C_Currency2			= iCurColumnPos(14)
			C_CurrencyPopup		= iCurColumnPos(15)
			C_Tax2				= iCurColumnPos(16)
			C_TaxPopup			= iCurColumnPos(17)
			C_RoutOrder2		= iCurColumnPos(18)
			C_MilestoneFlg2		= iCurColumnPos(19)
			C_InspFlg2			= iCurColumnPos(20)
			C_InsideFlg2		= iCurColumnPos(21)
			C_ProdtOrderNo2		= iCurColumnPos(22)
			C_OrderStatus2		= iCurColumnPos(23)
				
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
	Frm1.txtPlantCd.Focus
	
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
	arrParam(4) = "OP"
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
	Frm1.txtProdOrderNo.Focus
	
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
    
	If lgIntFlgMode = Parent.OPMD_CMODE Then
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
		.Col = C_OdrNo
		arrParam(1) = .Text
	End With
	
	iCalledAspName = AskPRAspName("P4311RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4311RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenAltOprRef()  -------------------------------------------------
'	Name : OpenAltOprRef()
'	Description : Altered Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenAltOprRef()
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
	
	iCalledAspName = AskPRAspName("P4114RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	frm1.vspdData1.Row =frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = C_OdrNo
                
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	arrParam(1) = Trim(frm1.vspdData1.Text)			'☜: 조회 조건 데이타 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
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
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field명(0)
	arrField(1) = 2 '"ITEM_NM"					' Field명(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

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
Function OpenTrackingInfo(Byval strCode)
	
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtProdFromDt.Text
	arrParam(4) = frm1.txtProdToDt.Text	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
    
End Function

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConWC()

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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") ' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
	
End Function

'------------------------------------------  OpenConWC2()  -------------------------------------------------
'	Name : OpenConWC2()
'	Description : Condition Work Center PopUp for Grid 2
'---------------------------------------------------------------------------------------------------------
Function OpenConWC2(Byval strCode, Byval Row)

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
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = ""'strName												' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")	' Where Condition

	With frm1.VspdData2														'check RoutOrder 
	   .row = Row
	   .col = C_RoutOrder2	   
		If  Trim(UCase(.Text)) <> "I" then
			arrParam(4) = arrParam(4) & "And INSIDE_FLG = " & FilterVar("Y", "''", "S") & " "
		End if	
	End With 

	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    arrField(2) = "INSIDE_FLG"												' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    arrHeader(2) = "사내/외구분"												' Header명(2)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConWC2(arrRet,Row)
	End If	
	
End Function

'------------------------------------------  OpenBizPartner()  -------------------------------------------------
'	Name : OpenBizparener()
'	Description : BpPopup2
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizPartner(ByVal str, ByVal Row)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "외주처팝업"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(str)
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "외주처"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    arrField(2) = ""	'"BP_TYPE"
    arrField(3) = ""	
        
    arrHeader(0) = "BP"		
    arrHeader(1) = "BP명"		
    arrHeader(2) = ""	'"Bp 구분"		
    arrHeader(3) = ""
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(arrRet, Row)
	End If	
	
End Function

'------------------------------------------  OpenCurrency()  ---------------------------------------------
'	Name : OpenCurrency()
'	Description : Currency Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenCurrency(ByVal strCurCd, ByVal Row)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "통화팝업"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = Trim(strCurCd)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "통화"			
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
  
    
    arrHeader(0) = "통화"		
    arrHeader(1) = "통화명"		
    
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet, Row)
	End If	
	
End Function

'-----------------------------------------------  OpenTax()  ---------------------------------------------
'	Name : OpenTax()
'	Description : Tax Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTaxType(Byval strVat, ByVal Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VAT형태"						' 팝업 명칭 
	arrParam(1) = "B_MINOR" 						' TABLE 명칭 
	arrParam(2) = Trim(strVat)						' Code Condition
		
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & ""		' and b_minor.minor_cd=b_configuration.minor_cd " Where Condition
	arrParam(5) = "VAT형태"						' TextBox 명칭 
	
    arrField(0) = "b_minor.MINOR_CD"				' Field명(0)
    arrField(1) = "b_minor.MINOR_NM"

    
    arrHeader(0) = "VAT형태"					' Header명(0)
    arrHeader(1) = "VAT형태명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTaxType(arrRet, Row)
	End If	
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

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
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

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetCurrency()  ----------------------------------------------
'	Name : SetCurrency()
'	Description : RoutingNo Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrency(Byval arrRet, Byval Row)
	With frm1
		.vspdData2.Row = Row
		.vspdData2.Col = C_Currency2
		.vspdData2.Text = UCase(arrRet(0))
		
		Call vspdData2_Change(.vspdData2.Col, .vspdData2.Row)
	
	End With
End Function

'------------------------------------------  SetTaxType()  -----------------------------------------------
'	Name : SetTaxType()
'	Description : RoutingNo Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetTaxType(Byval arrRet, Byval Row)
	With frm1
		.vspdData2.Row = Row
		.vspdData2.Col = C_Tax2
		.vspdData2.Text = UCase(arrRet(0))
		
		Call vspdData2_Change(.vspdData2.Col, .vspdData2.Row)
	
	End With
End Function

'------------------------------------------  SetConWC()  --------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetConWC2()  ----------------------------------------------
'	Name : SetConWC2()
'	Description : Work Center Popup for Grid 2 에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConWC2(Byval arrRet, Byval Row)
	With frm1
	
	   .vspdData2.Row = Row
	   .vspdData2.Col = C_RoutOrder2 
	   If UCase(.vspdData2.Text) <> "I" And UCase(arrRet(2)) = "N" then
	       Call DisplayMsgBox("181415", "x", "x", "x")
	       Exit Function
	   End if
	   
		.vspdData2.Col = C_WCCd2
		.vspdData2.Text = UCase(arrRet(0))
		.vspdData2.Col = C_WCNm2
		.vspdData2.Text = UCase(arrRet(1))
		.vspdData2.Col = C_InsideFlg2
		.vspdData2.Text = UCase(arrRet(2))
		
		If UCase(arrRet(2)) = "Y" Then
		
		.vspdData2.Col = C_BpCd2
		.vspdData2.Text = ""
		.vspdData2.Col = C_BpNm2
		.vspdData2.Text = ""
		.vspdData2.Col = C_CCFCost2
		.vspdData2.Text = ""
		.vspdData2.Col = C_CCFAmt2
		.vspdData2.Text = ""
		.vspdData2.Col = C_Currency2
		.vspdData2.Text = ""
		.vspdData2.Col = c_Tax2
		.vspdData2.Text = ""
		
		End if
		
	    .vspdData2.ReDraw = False 
	    Call SetFieldProp(Row,UCase(arrRet(2))) 
	    .vspdData2.ReDraw = True
		
		.vspdData2.Col = C_WCCd2
		Call vspdData2_Change(.vspdData2.Col, .vspdData2.Row)
	  
	End With
	
End Function

'------------------------------------------  SetBizPartner()  --------------------------------------------------
'	Name : SetBizPartner()
'	Description : RoutingNo Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizPartner(Byval arrRet, ByVal Row)
	With frm1
		.vspdData2.Row = Row
		.vspdData2.Col = C_BpCd2
		.vspdData2.Text = UCase(arrRet(0))
		
		.vspdData2.Row = Row
		.vspdData2.Col = C_BpNm2
		.vspdData2.Text = arrRet(1)
		
		.vspdData2.Col = C_BpCd2
		Call vspdData2_Change(.vspdData2.Col, .vspdData2.Row)		' 변경이 일어났다고 알려줌 
		 
	End With
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	
	frm1.txtTrackingNo.Value = arrRet(0)
	
End Function    

'------------------------------------------  SetFieldProp()  -----------------------------------------
'	Name: SetFieldProp()
'	Description : WorkCenter type setting
'---------------------------------------------------------------------------------------------------------
Function SetFieldProp(ByVal lRow, ByVal sType)

	ggoSpread.Source = frm1.vspdData2

	If sType = "N" Then			'외주 공정이면 
	  with frm1.VspdData2
	    .col = C_RoutOrder2 
	    .row = lRow
	    If UCase(.text) = "I" Then
			ggoSpread.SpreadUnLock	C_BpCd2,		lRow, C_BpCdPopup2, lRow
			ggoSpread.SpreadUnLock	C_CCFCost2,		lRow,C_CCFCost2, lRow
			ggoSpread.SpreadUnLock	C_CCFAmt2,		lRow,C_CCFAmt2, lRow
			ggoSpread.SpreadUnLock	C_Currency2,	lRow,C_CurrencyPopup, lRow
			ggoSpread.SpreadUnLock	C_Tax2,		lRow,C_TaxPopup, lRow
			ggoSpread.SSSetRequired C_BpCd2,		lRow,lRow
			ggoSpread.SSSetRequired	C_CCFCost2,	lRow, lRow
			ggoSpread.SSSetRequired	C_CCFAmt2,	lRow, lRow
			ggoSpread.SSSetRequired	C_Currency2,	lRow, lRow
			ggoSpread.SSSetRequired	C_Tax2,		lRow, lRow
		 Else
			ggoSpread.SpreadLock		C_BpCd2,			lRow, C_BpCdPopup2,  lRow
			ggoSpread.SSSetProtected	C_BpCd2,			lRow, lRow
			ggoSpread.SSSetProtected	C_CCFCost2,	lRow, lRow
			ggoSpread.SSSetProtected	C_CCFAmt2,	lRow, lRow
			ggoSpread.SSSetProtected	C_Currency2,		lRow, lRow
			ggoSpread.SSSetProtected	C_CurrencyPopup,		lRow, lRow
			ggoSpread.SSSetProtected	C_Tax2,		lRow, lRow
 			ggoSpread.SSSetProtected	C_TaxPopup,		lRow, lRow
		End If
	  End With 

	ElseIf sType = "Y" Then		'사내 공정이면 
		ggoSpread.SpreadLock		C_BpCd2,			lRow, C_BpCdPopup2,  lRow
		ggoSpread.SSSetProtected	C_BpCd2,			lRow, lRow
		ggoSpread.SSSetProtected	C_CCFCost2,	lRow, lRow
		ggoSpread.SSSetProtected	C_CCFAmt2,	lRow, lRow
		ggoSpread.SSSetProtected	C_Currency2,		lRow, lRow
		ggoSpread.SSSetProtected	C_CurrencyPopup,		lRow, lRow
		ggoSpread.SSSetProtected	C_Tax2,		lRow, lRow
 		ggoSpread.SSSetProtected	C_TaxPopup,		lRow, lRow

	End If
	
End Function

'------------------------------------------  SetFieldProp2()  -----------------------------------------
'	Name: SetFieldProp2()
'	Description : WorkCenter type setting for Hidden Grid
'---------------------------------------------------------------------------------------------------------
Function SetFieldProp2(ByVal lRow, ByVal sType, Byval oType)

	ggoSpread.Source = frm1.vspdData3
	
	If sType = "N" and oType = "I" Then			'외주 공정이면 
		ggoSpread.SpreadUnLock	C_BpCd3,		lRow, C_BpCd3, lRow
		ggoSpread.SpreadUnLock	C_CCFCost3,		lRow,C_CCFCost3, lRow
		ggoSpread.SpreadUnLock	C_CCFAmt3,		lRow,C_CCFAmt3, lRow
		ggoSpread.SpreadUnLock	C_Currency3,	lRow,C_Currency3, lRow
		ggoSpread.SpreadUnLock	C_Tax3,			lRow,C_Tax3, lRow
		ggoSpread.SSSetRequired C_BpCd3,		lRow,lRow
		ggoSpread.SSSetRequired	C_CCFCost3,		lRow, lRow
		ggoSpread.SSSetRequired	C_CCFAmt3,		lRow, lRow
		ggoSpread.SSSetRequired	C_Currency3,	lRow, lRow
		ggoSpread.SSSetRequired	C_Tax3,			lRow, lRow

	ElseIf sType = "Y" Then		'사내 공정이면 
		ggoSpread.SpreadLock		C_BpCd3,	lRow, C_BpCd3,  lRow
		ggoSpread.SSSetProtected	C_BpCd3,	lRow, lRow
		ggoSpread.SSSetProtected	C_CCFCost3,	lRow, lRow
		ggoSpread.SSSetProtected	C_CCFAmt3,	lRow, lRow
		ggoSpread.SSSetProtected	C_Currency3,lRow, lRow
		ggoSpread.SSSetProtected	C_Tax3,		lRow, lRow

	End If
	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'==========================================  2.5.2 LookupWc()  =======================================
'	Name : LookUpWc()
'	Description : Lookup WorkCenter using Keyboard
'===================================================================================================== 
Sub LookUpWc(ByVal Str, ByVal Row, ByVal Row1, Byval strOrderNo, Byval strOprNo)
	
	Dim strVal
	Dim strSelect, strWhere
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	
	If Str = "" Then Exit Sub
	
	strSelect = " A.WC_CD,    A.WC_NM,    A.INSIDE_FLG    "
	
	strWhere = " A.PLANT_CD =   " & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	strWhere = strWhere & " AND A.WC_CD = " & FilterVar(Str, "''", "S")
	
	If 	CommonQueryRs2by2(strSelect, " P_WORK_CENTER A  (NOLOCK) ", strWhere, lgF0) = False Then
		Call DisplayMsgBox("182100","X", Frm1.vspdData2.Text,"X")
		Call LookUpWcNotOk(Row)    
		Exit Sub
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	With frm1.vspdData2
		.Row = Row
		.Col = C_WCNm2
		.text = lgF0(2)
		.Col = C_InsideFlg2
		.text = lgF0(3)   			
	
	End With
	
	Call LookUpWcOk(lgF0(1), lgF0(2), lgF0(3), Row, Row1, strOrderNo, strOprNo)
	
End Sub

Function LookUpWcOk(ByVal WcCd, ByVal WcNm,ByVal InsideFlg, ByVal Row, ByVal Row1, Byval strOrderNo, Byval strOprNo)

	Dim lRows, lFoundRow
	Dim strHndOprNo, strHndOrderNo, strRoutOrder
	
	With frm1.vspdData2
		
		If CLng(Row1) <> CLng(frm1.vspdData1.ActiveRow) Then

			For lRows = 1 To frm1.vspdData3.MaxRows
		
			    frm1.vspdData3.Row = lRows
			    frm1.vspdData3.Col = C_ProdtOrderNo3
			    strHndOrderNo = frm1.vspdData3.Text
			    frm1.vspdData3.Col = C_OperationCd3
			    strHndOprNo = frm1.vspdData3.Text

			    If Trim(strHndOprNo) = Trim(strOprNo) And Trim(strHndOrderNo) = Trim(strOrderNo) Then
					lFoundRow = lRows
					Exit For
			    End If    
			Next

			With frm1.vspdData3
				
				ggoSpread.Source = frm1.vspdData3
				
				.Row = lFoundRow
				.Col = C_InsideFlg3
				.Text = InsideFlg
				.Col = C_WcNm3
				.Text = WcNm
				.Col = C_RoutOrder3
				strRoutOrder = UCase(.Text)
				If strRoutOrder <> "I"And UCase(InsideFlg) = "N" Then
					Call DisplayMsgBox("181415", "x", "x", "x")
				End if

				.ReDraw = False

				.Col = C_WCCd3
				.Text = WcCd

				If UCase(InsideFlg) = "Y" Then
					.Col = C_BpCd3
					.Text = ""
					.Col = C_BpNm3
					.Text = ""
					.Col = C_CCFCost3
					.Text = ""
					.Col = C_CCFAmt3
					.Text = ""
					.Col = C_Currency3
					.Text = ""
					.Col = c_Tax3
					.Text = ""
					Call SetFieldProp2(lFoundRow,"Y",strRoutOrder)
				Else
					Call SetFieldProp2(lFoundRow,"N",strRoutOrder)
				End If				
		
				.ReDraw = True
				
			End With
			
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.UpdateRow lRows
			
		Else
	
			.Row = Row
			.Col = C_InsideFlg2
			.Text = InsideFlg
			.Col = C_WcNm2
			.Text = WcNm
			.Row = Row
			.Col = C_RoutOrder2 
			If UCase(.Text) <> "I"And UCase(InsideFlg) = "N" Then
				Call DisplayMsgBox("181415", "x", "x", "x")
			End if
				
			.ReDraw = False	
	
			If UCase(InsideFlg) = "Y" Then
				.Col = C_BpCd2
				.Text = ""
				.Col = C_BpNm2
				.Text = ""
				.Col = C_CCFCost2
				.Text = ""
				.Col = C_CCFAmt2
				.Text = ""
				.Col = C_Currency2
				.Text = ""
				.Col = c_Tax2
				.Text = ""
				Call SetFieldProp(Row,"Y")
			Else
				Call SetFieldProp(Row,"N")
			End If
		
			.ReDraw = True

			ggoSpread.Source = frm1.vspdData2
			ggoSpread.UpdateRow Row
			CopyToHSheet Row
		
		End If
		
	End With

	IsOpenPop = False
		
End Function

Function LookUpWcNotOk(Byval Row)

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.UpdateRow Row
		CopyToHSheet Row


	IsOpenPop = False
	
End Function


'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################%>
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	    
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

    Call InitSpreadSheet("*")                                                    '⊙: Setup the Spread sheet

    Call InitVariables                                                      '⊙: Initializes local global variables

    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitComboBox
    Call InitSpreadComboBox
    Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
    
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus 
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
'=======================================================================================================
'   Event Name : txtProdFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtProdToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================0
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdToDt.Focus
    End If
End Sub

'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtProdFromDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdFromDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtProdToDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdToDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1
    
 	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		
 		lgOldRow = Row
		If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then		
			frm1.vspdData1.Col = 1
			frm1.vspdData1.Row = row
		
			lgOldRow = Row
		
			If DbDtlQuery(Row) = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If		
	 	'------ Developer Coding part (End)
	
 	End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0101111111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	End If
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		'ggoSpread.Source = frm1.vspdData2 
 		'If lgSortKey2 = 1 Then
 		'	ggoSpread.SSSort Col					'Sort in Ascending
 		'	lgSortKey2 = 2
 		'Else
 		'	ggoSpread.SSSort Col, lgSortKey2		'Sort in Descending
 		'	lgSortKey2 = 1
 		'End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
	
 	End If
 	
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If frm1.vspdData1.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If
	
	Call vspdData1_Click(NewCol, NewRow)
End Sub


'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)         

	Dim strItemCd, strWcCd
	Dim strHndItemCd, strHndOprNo
	Dim i, lActRow
	Dim strReqDt, strEndDt
	Dim	dblPrice, dblAmt, dblOrderQty
	Dim dtPlanStartDt, dtPlanEndDt
	Dim	dtPlanStartDtDtl, dtPlanEndDtDtl
	Dim strOrderNo,strOprNo
	
	With frm1.vspdData2

		.Row = Row
		Select Case Col
		
			Case C_OperationCd2
				If Trim(GetSpreadText(frm1.vspdData2,Col,Row,"X","X")) <> "" Then
					If CheckValidOprNo(Trim(GetSpreadText(frm1.vspdData2,Col,Row,"X","X")), Row) = False Then
						Call frm1.vspdData2.SetText(Col,Row,"")
						Exit Sub
					Else
						Call SetSpread2Color(Row)
						CopyToHSheet Row	
					End If
				End If
		
		    Case C_WcCd2
					ggoSpread.Source = frm1.vspdData2
					frm1.vspdData2.Col = Col	
					frm1.vspdData2.Row = Row
					strWcCd = frm1.vspdData2.text
					If frm1.vspdData2.Text <> "" Then
						IsOpenPop = True
						ggoSpread.Source = frm1.vspdData1
						lActRow = frm1.vspdData1.ActiveRow
						ggoSpread.Source = frm1.vspdData2
						frm1.vspdData2.Row = Row						
						frm1.vspdData2.Col = C_ProdtOrderNo2
						strOrderNo = frm1.vspdData2.text
						frm1.vspdData2.Col = C_OperationCd2
						strOprNo = frm1.vspdData2.text
						Call LookUpWc(strWcCd,Row,lActRow,strOrderNo,strOprNo)
					End If
		
			Case C_Bpcd2
		
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.UpdateRow Row
					CopyToHSheet Row

		    Case C_CCFCost2

				If Col = C_CCFCost2 Then
					.Col = C_CCFCost2
					
					dblPrice = UNICDbl(.Text)
					If dblPrice <= 0 Then
						
					End If
					.Col = C_CCFAmt2
					dblAmt = UNICDbl(.Text)
					If dblAmt = 0 Then
						ggoSpread.Source = frm1.vspdData1
						frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
						frm1.vspdData1.Col = C_OrderQuantity
						dblOrderQty = UNICDbl(frm1.vspdData1.Text)
						dblAmt = dblPrice * dblOrderQty
						ggoSpread.Source = frm1.vspdData2
						frm1.vspdData2.Row = Row
						frm1.vspdData2.Col = C_CCFAmt2
						frm1.vspdData2.Text = dblAmt						
					End If					
				End If

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				Call FixDecimalPlaceByCurrency(frm1.vspdData2,Row,C_Currency2,C_CCFCost2, "C" ,"X","X")
				
				CopyToHSheet Row
				

		    Case C_CCFAmt2

				If Col = C_CCFCost2 Then
					.Col = C_CCFCost2
					
					dblPrice = UNICDbl(.Text)
					If dblPrice <= 0 Then
						
					End If
					.Col = C_CCFAmt2
					dblAmt = UNICDbl(.Text)
					If dblAmt = 0 Then
						ggoSpread.Source = frm1.vspdData1
						frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
						frm1.vspdData1.Col = C_OrderQuantity
						dblOrderQty = UNICDbl(frm1.vspdData1.Text)
						dblAmt = dblPrice * dblOrderQty
						ggoSpread.Source = frm1.vspdData2
						frm1.vspdData2.Row = Row
						frm1.vspdData2.Col = C_CCFAmt2
						frm1.vspdData2.Text = dblAmt						
					End If					
				End If

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				Call FixDecimalPlaceByCurrency(frm1.vspdData2,Row,C_Currency2,C_CCFAmt2, "A" ,"X","X")
				
				CopyToHSheet Row				

		    Case C_Tax2

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row

		    Case C_Currency2

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row

				Call FixDecimalPlaceByCurrency(frm1.vspdData2,Row,C_Currency2,C_CCFCost2, "C" ,"X","X")
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,Row,Row,C_Currency2,C_CCFCost2, "C" ,"I","X","X")

				Call FixDecimalPlaceByCurrency(frm1.vspdData2,Row,C_Currency2,C_CCFAmt2, "A" ,"X","X")
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,Row,Row,C_Currency2,C_CCFAmt2, "A" ,"I","X","X")

				CopyToHSheet Row

		    Case C_PlanStartDt2

				.Col = C_PlanStartDt2
				dtPlanStartDtDtl = .Text
				.Col = C_PlanEndDt2
				dtPlanEndDtDtl = .Text
				If dtPlanStartDtDtl <> "" and dtPlanEndDtDtl <> "" Then  
					If CompareDateByFormat(dtPlanStartDtDtl,dtPlanEndDtDtl,"","","970025",parent.gDateFormat,Parent.gComDateType,False) = False Then  'If CDate(dtPlanStartDtDtl) > CDate(dtPlanEndDtDtl) Then  
						Call DisplayMsgBox("189206", "x", "x", "x")
						ggoSpread.Source = frm1.vspdData2
						.Col = C_PlanStartDt2
						.Text = ""
						ggoSpread.Source = frm1.vspdData2
						ggoSpread.UpdateRow Row
						CopyToHSheet Row
						Exit Sub
					End If
				End If
				ggoSpread.Source = frm1.vspdData1
				frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
				frm1.vspdData1.Col = C_PlanStartDt1
				dtPlanStartDt = frm1.vspdData1.Text
						    
				If CompareDateByFormat(dtPlanStartDt,dtPlanStartDtDtl,"","","970025",parent.gDateFormat,Parent.gComDateType,False) = False Then	'If CDate(dtPlanStartDt) > CDate(dtPlanStartDtDtl) Then  
					Call DisplayMsgBox("189304", "x", "x", "x")
					ggoSpread.Source = frm1.vspdData2
					.Col = C_PlanStartDt2
					.Text = ""
				End If

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row

				CopyToHSheet Row
		    
		    Case C_PlanEndDt2
		    
				.Col = C_PlanStartDt2
				dtPlanStartDtDtl = .Text
				.Col = C_PlanEndDt2
				dtPlanEndDtDtl = .Text
				
				If dtPlanStartDtDtl <> "" and dtPlanEndDtDtl <> "" Then
					If CompareDateByFormat(dtPlanStartDtDtl,dtPlanEndDtDtl,"","","970025",parent.gDateFormat,Parent.gComDateType,False) = False Then		'If CDate(dtPlanStartDtDtl) > CDate(dtPlanEndDtDtl) Then  
						Call DisplayMsgBox("189207", "x", "x", "x")
						ggoSpread.Source = frm1.vspdData2
						.Col = C_PlanEndDt2
						.Text = ""
						ggoSpread.Source = frm1.vspdData2
						ggoSpread.UpdateRow Row
						CopyToHSheet Row
						Exit Sub
					End If
				End If

				ggoSpread.Source = frm1.vspdData1
				frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
				frm1.vspdData1.Col = C_PlanEndDt1
				dtPlanEndDt = frm1.vspdData1.Text

				If CompareDateByFormat(dtPlanEndDtDtl,dtPlanEndDt,"","","970025",parent.gDateFormat,Parent.gComDateType,False) = False Then		'If CDate(dtPlanEndDt) < CDate(dtPlanEndDtDtl) Then
					Call DisplayMsgBox("189305", "x", "x", "x")
					ggoSpread.Source = frm1.vspdData2
					.Col = C_PlanEndDt2
					.Text = ""
				End If

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row

				CopyToHSheet Row
				
			Case C_InspFlg2
				
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row
				
			Case C_MilestoneFlg2
					
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row	
				
			Case C_WorkCd2, C_WorkNm2
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row	
				
		End Select

	End With
	
	
End Sub

'========================================================================================================
'   Event Name : vspdData2_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData2_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_CCFCost2
            Call EditModeCheck(frm1.vspdData2, Row, C_Currency2, C_CCFCost2, "C" ,"I", Mode, "X", "X")
        Case C_CCFAmt2
            Call EditModeCheck(frm1.vspdData2, Row, C_Currency2, C_CCFAmt2, "A" ,"I", Mode, "X", "X")        
    End Select
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim	strFlag
	
	With frm1.vspdData2
	
		.Row = Row
		Select Case Col

			Case  C_InspFlg2
			
				.Col = Col
				strFlag = .Text
				.Col = C_MilestoneFlg2

				If strFlag = "Y" and .Text = "N" Then
					Call DisplayMsgBox("189307", "x", "x", "x")
					.Col = Col
					.Text = "N"
					Exit Sub
				End If
				
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row
								
			Case  C_MilestoneFlg2
			
				.Col = Col
				strFlag = .Text
				.Col = C_InspFlg2
				If strFlag = "N" and .Text = "Y" Then
					Call DisplayMsgBox("189307", "x", "x", "x")
					.Col = Col
					.Text = "Y"
					Exit Sub
				End If
				
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row
				
			Case C_WorkCd2, C_WorkNm2
				.Col = Col
				intIndex = .Value
				If Col = C_WorkCd2 Then
					.Col = C_WorkNm2
				Else
					.Col = C_WorkCd2
				End If
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row	

		End Select
		
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_DblClick(ByVal Col , ByVal Row)
	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'==========================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_DblClick(ByVal Col , ByVal Row)
	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub


'==========================================================================================
'   Event Name : vspdData_DragDropBlock
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData2_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

Dim strCode
Dim strName

    With frm1.vspdData2
    
		ggoSpread.Source = frm1.vspdData2
		If Row < 1 Then Exit Sub

		Select Case Col

            Case C_WcCdPopup2
				.Col = C_WcCd2
				.Row = Row
				strCode = .Text
				Call OpenConWc2(strCode, Row)
				Call SetActiveCell(frm1.vspdData2, C_WcCd2, Row, "M","X","X")
				Set gActiveElement = document.activeElement
				
			Case C_BpcdPopup2
				.Col = C_BpCd2
				.Row = Row
				strCode = .Text
				Call OpenBizPartner(strCode, Row)
				Call SetActiveCell(frm1.vspdData2, C_BpCd2, Row, "M","X","X")
				Set gActiveElement = document.activeElement
            
		    Case C_CurrencyPopup
				.Col = C_Currency2
				.Row = Row
				strCode = .Text
				Call OpenCurrency(strCode, Row)
				Call SetActiveCell(frm1.vspdData2, C_Currency2, Row, "M","X","X")
				Set gActiveElement = document.activeElement
	    
		    Case C_TaxPopup
				.Col = C_Tax2
				.Row = Row
				strCode = .Text
				Call OpenTaxType(strCode, Row)
				Call SetActiveCell(frm1.vspdData2, C_Tax2, Row, "M","X","X")
				Set gActiveElement = document.activeElement

		End Select

	End With

End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
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
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgIntPrevKey <> 0 Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If LayerShowHide(1) = False Then Exit Sub
			If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If     
    End if
    
End Sub

'========================================================================================
' Function Name : vspdData1_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
  
'========================================================================================
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
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
    
	Dim LngRow, lngRows
	Dim strProdtOrderNo, strOperationNo
	Dim strHdnOrderNo, strHdnOprNo
	 
    If gActiveSpdSheet.Id = "B" Then

		frm1.vspdData2.Col = C_ProdtOrderNo2       
		frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
		strProdtOrderNo = frm1.vspdData2.text
		
		ggoSpread.Source = frm1.vspdData3
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("C")
		ggoSpread.ReOrderingSpreadData()
			
    End If

    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)  
    
    If gActiveSpdSheet.Id = "A" Then
    
		Call ggoSpread.ReOrderingSpreadData
    
    ElseIf gActiveSpdSheet.Id = "B" Then
		
		Call InitSpreadComboBox
		ggoSpread.Source = frm1.vspdData3
		Call CopyFromhSheet(strProdtOrderNo)
		
		ggoSpread.Source = frm1.vspdData2
		Call InitData(1)
		
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
    
    FncQuery = False											'⊙: Processing is NG
    
    Err.Clear													'☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData3							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")	'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then							'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData

	If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function	

    Call InitVariables
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function											'☜: Query db data
	End If
       
    FncQuery = True												'⊙: Processing is OK
    
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

    Dim IntRetCD 
    Dim	LngRows
	    
    FncSave = False												'⊙: Processing is NG
    
    Err.Clear                                                   '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData3							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData3							'⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function													'☜: Save db data
    
    FncSave = True												'⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
        
	If frm1.vspdData2.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData2.focus
    Set gActiveElement = document.activeElement 
	frm1.vspdData2.EditMode = True
	frm1.vspdData2.ReDraw = False
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.CopyRow
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.CopyRow
    Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData2, frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow, C_Currency2, C_CCFCost2, "C", "I", "X", "X")
    Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData2, frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow, C_Currency2, C_CCFAmt2, "A", "I", "X", "X")
    frm1.vspdData2.ReDraw = True
    SetSpreadColor frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow
   
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

Dim Row
Dim strMode
Dim	strProdtOrderNo
Dim	strOperationCd
Dim lngRows
Dim strHdnOrderNo, strHdnOprNo

	If frm1.vspdData2.MaxRows < 1 Then Exit Function	

    ggoSpread.Source = frm1.vspdData2	
    Row = frm1.vspdData2.ActiveRow
    frm1.vspdData2.Row = Row
    frm1.vspdData2.Col = 0
    strMode = frm1.vspdData2.Text
    frm1.vspdData2.Col = C_ProdtOrderNo2
    strProdtOrderNo = frm1.vspdData2.Text
    frm1.vspdData2.Col = C_OperationCd2
    strOperationCd = frm1.vspdData2.Text

	If strMode = ggoSpread.InsertFlag Then
		Call DeleteHSheet(strProdtOrderNo, strOperationCd)
		ggoSpread.Source = frm1.vspdData2
	    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	   
	Else
		'------------------------------------
		' Find First Row
		'------------------------------------ 
		For lngRows = 1 To frm1.vspdData3.MaxRows
		    frm1.vspdData3.Row = lngRows
		    frm1.vspdData3.Col = C_ProdtOrderNo3
		    strHdnOrderNo = frm1.vspdData3.Text
		    frm1.vspdData3.Col = C_OperationCd3
		    strHdnOprNo = frm1.vspdData3.Text
		    If strProdtOrderNo = strHdnOrderNo and strOperationCd = strHdnOprNo Then
		        Exit For
		    End If
		Next
		
		ggoSpread.Source = frm1.vspdData3
	    ggoSpread.EditUndo lngRows
	    
	    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,Frm1.vspdData2.ActiveRow,Frm1.vspdData2.ActiveRow,C_Currency2,C_CCFCost2, "C" ,"I","X","X")
	    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,Frm1.vspdData2.ActiveRow,Frm1.vspdData2.ActiveRow,C_Currency2,C_CCFAmt2, "A" ,"I","X","X")
		
	    Call CopyOneRowFromHSheet(lngRows, Row)
	End If

End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
	Dim imRow
	Dim pvRow
	
	Dim pvOprStatus
	
	On Error Resume Next
	
	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)

	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If
	
    With frm1
		
		If .vspdData2.ActiveRow = .vspdData2.MaxRows Then
			Call .vspdData2.GetText(C_OrderStatus2, .vspdData2.ActiveRow, pvOprStatus)
		Else
			Call .vspdData2.GetText(C_OrderStatus2, .vspdData2.ActiveRow + 1, pvOprStatus)
		End If	
		
		If pvOprStatus = "CL" Or pvOprStatus = "ST" Then
			Call DisplayMsgBox("189369", "X", "X", "X")
			Exit Function
		End If
		    
		.vspdData2.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False
		If frm1.vspdData2.selBlockRow = -1 Then
			ggoSpread.InsertRow 0, imRow
		Else
			ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
    	End If
    	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,Frm1.vspdData2.ActiveRow,Frm1.vspdData2.ActiveRow + imRow - 1,C_Currency2,C_CCFCost2, "C" ,"I","X","X")
    	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,Frm1.vspdData2.ActiveRow,Frm1.vspdData2.ActiveRow + imRow - 1,C_Currency2,C_CCFAmt2, "A" ,"I","X","X")
    	
    	For pvRow = .vspdData2.ActiveRow To .vspdData2.ActiveRow + imRow -1
    		.vspdData1.Row = .vspdData1.ActiveRow
			.vspdData2.Row = pvRow
			.vspdData1.Col = C_OdrNo
			.vspdData2.Col = C_ProdtOrderNo2
			.vspdData2.text = .vspdData1.text
			.vspdData1.Col = C_OrderStatus
			If UCase(Trim(.vspdData1.text)) = "RL" Or UCase(Trim(.vspdData1.text)) = "ST" Then
				.vspdData2.Col = C_OrderStatus2
				.vspdData2.text = "RL"
			Else
				.vspdData2.Col = C_OrderStatus2
				.vspdData2.text = "OP"
			End If
			.vspdData2.Col = C_MilestoneFlg2
			.vspdData2.text = "Y"
			.vspdData2.Col = C_InspFlg2
			.vspdData2.text= "N"
			.vspdData2.Col = C_RoutOrder2
			If pvRow = 1 Then
				If pvRow = .vspdData2.MaxRows Then
					.vspdData2.text = "S"
				Else
					.vspdData2.text = "F"
				End If
			ElseIf pvRow = .vspdData2.MaxRows Then
				.vspdData2.text = "L"
			Else
				.vspdData2.text = "I"
			End If
			
			.vspdData2.ReDraw = True
		Next
		
		SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow -1
    End With
    
    Set gActiveElement = document.ActiveElement
	If Err.number = 0 Then FncInsertRow = True
    
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt, i
    Dim pvOprStatus

    With frm1

		.vspdData1.Row = frm1.vspdData1.ActiveRow

		If .vspdData2.MaxRows < 1 Then Exit Function
		
		Call .vspdData2.GetText(C_OrderStatus2, .vspdData2.ActiveRow, pvOprStatus)
		
		If pvOprStatus = "CL" Or pvOprStatus = "ST" Then
			Call DisplayMsgBox("189369", "X", "X", "X")
			Exit Function
		End If

    End With
    
    Call DeleteMarkingHSheet()
    
	ggoSpread.Source = frm1.vspdData2
    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows

	'CopyToHSheet frm1.vspdData2.ActiveRow

End Function


'=======================================================================================================
'   Function Name : DeleteMarkingHSheet
'   Function Desc : DeleteMark the Row Which keys match with vapdData's Key and vspdData2's Key
'=======================================================================================================
Function DeleteMarkingHSheet()

	Dim lngRow2, lRow, lRows
	
	Dim strOrderNo, strOprNo
	Dim strHdnOrderNo, strHdnOprNo
	
	DeleteMarkingHSheet = False
	
	For lngRow2 = frm1.vspdData2.SelBlockRow To frm1.vspdData2.SelBlockRow2
	
        For lRows = 1 To frm1.vspdData3.MaxRows
            frm1.vspdData3.Row = lRows
            frm1.vspdData3.Col = C_ProdtOrderNo3
            strHdnOrderNo = frm1.vspdData3.Text
            frm1.vspdData3.Col = C_OperationCd3
            strHdnOprNo = frm1.vspdData3.Text
            frm1.vspdData2.Row = lngRow2
            frm1.vspdData2.Col = C_ProdtOrderNo2
            strOrderNo = frm1.vspdData2.Text
            frm1.vspdData2.Col = C_OperationCd2
            strOprNo = frm1.vspdData2.Text
            If strHdnOrderNo = strOrderNo And strHdnOprNo = strOprNo Then
				lRow = lRows
				Exit For
            End If    
		Next
	
		If lRow > 0 Then
			With frm1
    			ggoSpread.Source = .vspdData3
		 		.vspdData3.Col = 0
				.vspdData3.Text = ggoSpread.DeleteFlag
			End With
		End If
	Next
	
	DeleteMarkingHSheet = True
	
End Function    


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.fncPrint()                                                   '☜: Protect system from crashing
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
    Call parent.FncExport(Parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData2							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
    
    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    Err.Clear

    With frm1

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)
		strVal = strVal & "&txtProdOrdNo=" & Trim(.hProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.Value)
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.Value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.Value)
		strVal = strVal & "&txtProdfromDt=" & Trim(.hProdfromDt.Value)
		strVal = strVal & "&txtProdtoDt=" & Trim(.hProdtoDt.Value)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.hOrderType.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&cboJobCd=" & Trim(.hJobCd.Value)
		strVal = strVal & "&cboOrderStatus=" & Trim(frm1.hOrderStatus.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	Else
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)
		strVal = strVal & "&txtProdOrdNo=" & Trim(.txtProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
        strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.Value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.Value)
		strVal = strVal & "&txtProdfromDt=" & Trim(.txtProdfromDt.Text)
		strVal = strVal & "&txtProdtoDt=" & Trim(.txtProdtoDt.Text)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.cboOrderType.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&cboJobCd=" & Trim(.cboJobCd.Value)
		strVal = strVal & "&cboOrderStatus=" & Trim(frm1.cboOrderStatus.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Call SetToolBar("11001111000111")										'⊙: 버튼 툴바 제어 
	
	frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1

	lgOldRow = 1

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		If DbDtlQuery(frm1.vspdData1.Row) = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If

End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryNotOk()

	Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 

End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByVal LngRow) 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOrderNo
    
	boolExist = False
    With frm1

		.vspdData2.MaxRows = 0
	    .vspdData1.Row = LngRow
	    .vspdData1.Col = C_OdrNo
	    strOrderNo = .vspdData1.Text
    
	    If CopyFromHSheet(strOrderNo) = True Then
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2, 1, Frm1.vspdData2.MaxRows, C_Currency2,C_CCFCost2, "C", "I", "X", "X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2, 1, Frm1.vspdData2.MaxRows, C_Currency2,C_CCFAmt2, "A", "I", "X", "X")    	               
           Exit Function
        End If

		DbDtlQuery = False   
    
		.vspdData1.Row = .vspdData1.ActiveRow

		If LayerShowHide(1) = False Then Exit Function

		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strOrderNo)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strOrderNo)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbDtlQuery = True

End Function

Function DbDtlQueryOk()												'☆: 조회 성공후 실행로직 

	Dim LngRow

    '-----------------------
    'Reset variables area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2

	frm1.vspdData2.ReDraw = False
	
	'Call InitData(LngMaxRow)
	
	For LngRow = 1 To frm1.vspdData2.MaxRows
		
		Call SetSpread2Color(LngRow)
	
	Next
	
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgAfterQryFlg = True

	frm1.vspdData2.ReDraw = True

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData(Byval Row)

Dim strOprNo, strOrderNo
Dim strHndOprNo, strHndOrderNo
Dim lRows

    FindData = 0

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows

            .vspdData3.Row = lRows
            .vspdData3.Col = C_ProdtOrderNo3
            strHndOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OperationCd3
            strHndOprNo = .vspdData3.Text
            .vspdData2.Row = frm1.vspdData2.Row
            .vspdData2.Col = C_ProdtOrderNo2
            strOrderNo = .vspdData2.Text
            .vspdData2.Col = C_OperationCd2
            strOprNo = .vspdData2.Text
           
            If Trim(strHndOprNo) = Trim(strOprNo) And Trim(strHndOrderNo) = Trim(strOrderNo) Then
				FindData = lRows
				Exit Function
            End If    
        Next
        
    End With
    
End Function


'=======================================================================================================
'   Function Name : CopyFromHSheet
'   Function Desc : 
'====================================================================================================
Function CopyFromHSheet(ByVal strOrderNo)

Dim lngRows, LngRow
Dim boolExist
Dim iCols
Dim strHdnOrderNo
Dim strStatus
Dim iCurColumnPos

    boolExist = False
    
    CopyFromHSheet = boolExist
    
    ggoSpread.Source = frm1.vspdData2
 			
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
    With frm1

        Call SortHSheet()

        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_ProdtOrderNo3
            strHdnOrderNo = .vspdData3.Text
           
            If strOrderNo = strHdnOrderNo Then
                boolExist = True
                Exit For
            End If
        Next

        '------------------------------------
        ' Show Data
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            frm1.vspdData2.Redraw = False
            
            While lngRows <= .vspdData3.MaxRows

	             .vspdData3.Row = lngRows
                
                .vspdData3.Col = C_ProdtOrderNo3
				strHdnOrderNo = .vspdData3.Text
                
                If strOrderNo <> strHdnOrderNo Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
					If strOrderNo = strHdnOrderNo Then
						.vspdData2.MaxRows = .vspdData2.MaxRows + 1
						.vspdData2.Row = .vspdData2.MaxRows
						.vspdData2.Col = 0
						.vspdData3.Col = 0
						.vspdData2.Text = .vspdData3.Text
						
						For iCols = 1 To .vspdData3.MaxCols
						    .vspdData2.Col = iCurColumnPos(iCols)
						    .vspdData3.Col = iCols
						    .vspdData2.Text = .vspdData3.Text
						Next
						
						LngRow = .vspdData2.MaxRows
						ggoSpread.Source = frm1.vspdData2
						
						Call SetSpread2Color(LngRow)
					End If
                End If   
                
				Call FixDecimalPlaceByCurrency(frm1.vspdData2,.vspdData2.MaxRows,C_Currency2,C_CCFCost2, "C" ,"X","X")
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,.vspdData2.MaxRows,.vspdData2.MaxRows,C_Currency2,C_CCFCost2, "C" ,"I","X","X")

				Call FixDecimalPlaceByCurrency(frm1.vspdData2,.vspdData2.MaxRows,C_Currency2,C_CCFAmt2, "A" ,"X","X")
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,.vspdData2.MaxRows,.vspdData2.MaxRows,C_Currency2,C_CCFAmt2, "A" ,"I","X","X")                
                
                
                lngRows = lngRows + 1
                
            Wend
            frm1.vspdData2.Redraw = True

        End If
            
    End With        
    
    CopyFromHSheet = boolExist
    
End Function

'=======================================================================================================
'   Function Name : CopyOneRowFromHSheet
'   Function Desc : 
'====================================================================================================
Function CopyOneRowFromHSheet(ByVal SourceRow, ByVal TargetRow)

Dim iCols
Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
    With frm1
        '------------------------------------
        ' Show Data
        '------------------------------------ 
		.vspdData3.Row = SourceRow
		frm1.vspdData2.Redraw = False
		.vspdData2.Row = TargetRow
		.vspdData2.Col = 0
		.vspdData3.Col = 0
		.vspdData2.Text = .vspdData3.Text
		For iCols = 1 To .vspdData3.MaxCols
		    .vspdData2.Col = iCurColumnPos(iCols)
		    .vspdData3.Col = iCols
		    .vspdData2.Text = .vspdData3.Text
		Next
		ggoSpread.Source = frm1.vspdData2
		.vspdData2.Row = TargetRow
		.vspdData2.Col = C_InsideFlg2
		If .vspdData2.Text = "Y" Then
			ggoSpread.SSSetProtected C_BpCd2,			TargetRow, TargetRow
			ggoSpread.SSSetProtected C_CCFCost2,		TargetRow, TargetRow
			ggoSpread.SSSetProtected C_CCFAmt2,			TargetRow, TargetRow
			ggoSpread.SSSetProtected C_Currency2,		TargetRow, TargetRow
			ggoSpread.SSSetProtected C_Tax2,			TargetRow, TargetRow
			ggoSpread.SpreadLock C_BpCdPopup2,			TargetRow, C_BpCdPopup2, TargetRow
			ggoSpread.SpreadLock C_CurrencyPopup,		TargetRow, C_CurrencyPopup, TargetRow
			ggoSpread.SpreadLock C_TaxPopup,			TargetRow, C_TaxPopup, TargetRow
		Else
			ggoSpread.SpreadUnLock C_BpCd2,				TargetRow, C_BpCd2, TargetRow
			ggoSpread.SpreadUnLock C_CCFCost2,			TargetRow, C_CCFCost2, TargetRow
			ggoSpread.SpreadUnLock C_CCFAmt2,			TargetRow, C_CCFAmt2, TargetRow
			ggoSpread.SpreadUnLock C_Currency2,			TargetRow, C_Currency2, TargetRow
			ggoSpread.SpreadUnLock C_Tax2,				TargetRow, C_Tax2, TargetRow
			
			ggoSpread.SSSetRequired C_BpCd2,			TargetRow, TargetRow
			ggoSpread.SSSetRequired C_CCFCost2,			TargetRow, TargetRow
			ggoSpread.SSSetRequired C_CCFAmt2,			TargetRow, TargetRow
			ggoSpread.SSSetRequired C_Currency2,		TargetRow, TargetRow
			ggoSpread.SSSetRequired C_Tax2,				TargetRow, TargetRow
			ggoSpread.SpreadUnLock C_BpCdPopup2,		TargetRow, C_BpCdPopup2, TargetRow
			ggoSpread.SpreadUnLock C_CurrencyPopup,		TargetRow, C_CurrencyPopup, TargetRow
			ggoSpread.SpreadUnLock C_TaxPopup,			TargetRow, C_TaxPopup, TargetRow
		End If

		.vspdData2.Col = C_RoutOrder2
		If .vspdData2.Text = "S" Or .vspdData2.Text = "L" Then
			ggoSpread.SSSetProtected C_InspFlg2,		TargetRow, TargetRow
			ggoSpread.SSSetProtected C_MilestoneFlg2,	TargetRow, TargetRow
		Else
			ggoSpread.SSSetRequired C_InspFlg2,			TargetRow, TargetRow
			ggoSpread.SSSetRequired C_MilestoneFlg2,	TargetRow, TargetRow
		End If
                
		frm1.vspdData2.Redraw = True

    End With        
   
End Function

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 
'=======================================================================================================
Sub CopyToHSheet(ByVal Row)
Dim lRow
Dim iCols
Dim LngRow
Dim Stype 'for hidden grid function
Dim Otype 'for hidden grid function
Dim iCurColumnPos
	
	ggoSpread.Source = frm1.vspdData2
 			
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
	With frm1 
                
	    lRow = FindData(Row)
	    
	    If lRow > 0 Then
			LngRow = lRow
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            For iCols = 1 To 22 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text    
            Next
        Else
			.vspdData3.MaxRows = .vspdData3.MaxRows + 1
			LngRow = .vspdData3.MaxRows
            .vspdData3.Row = .vspdData3.MaxRows
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
       
            For iCols = 1 To 22 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
                
            Next
        
        End If
        
        .vspdData3.Col = C_InsideFlg3 
        Stype = UCase(.vspdData3.Text)
        .vspdData3.Col = C_RoutOrder3 
        Otype = UCase(.vspdData3.Text)
                
		Call FixDecimalPlaceByCurrency(frm1.vspdData3,lRow,C_Currency3,C_CCFCost3, "C" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData3,lRow,lRow,C_Currency3,C_CCFCost3, "C" ,"I","X","X")

		Call FixDecimalPlaceByCurrency(frm1.vspdData3,lRow,C_Currency3,C_CCFAmt3, "A" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData3,lRow,lRow,C_Currency3,C_CCFAmt3, "A" ,"I","X","X")

        call SetFieldProp2(lRow,Stype,Otype)
	End With
	
End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal strProdtOrderNo, Byval strOperationCd)

Dim boolExist
Dim lngRows
Dim strHndProdtOrderNo, strHndOperationCd
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_ProdtOrderNo3
			strHndProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OperationCd3
			strHndOperationCd = .vspdData3.Text

            If strProdtOrderNo = strHndProdtOrderNo and strOperationCd = strHndOperationCd Then
                boolExist = True
                Exit For
            End If    
        Next
       
        '------------------------------------
        ' Data Delete
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
				.vspdData3.Col = C_ProdtOrderNo3
				strHndProdtOrderNo = .vspdData3.Text
				.vspdData3.Col = C_OperationCd3
				strHndOperationCd = .vspdData3.Text
                
                If (strProdtOrderNo <> strHndProdtOrderNo) or (strOperationCd <> strHndOperationCd) Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If   

            Wend
            
            ggoSpread.Source = frm1.vspdData2
            
            frm1.vspdData2.Row = lgCurrRow
            frm1.vspdData2.Col = frm1.vspdData2.MaxCols
            ggoSpread.Source = frm1.vspdData2

            frm1.vspdData2.Redraw = True

        End If

    End With

    DeleteHSheet = True
End Function    

'======================================================================================================
' Function Name : SortHSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortHSheet()
    
    With frm1
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0 'SS_SORT_BY_ROW
       
        .vspdData3.SortKey(1) = C_ProdtOrderNo3
        .vspdData3.SortKey(2) = C_OperationCd3
        
        .vspdData3.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25 'SS_ACTION_SORT
        .vspdData3.BlockMode = False
    End With        
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows
    Dim strVal, strDel
    Dim	DblPrc, DblAmt, DblRout
    
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
	
    DbSave = False                                                          '⊙: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function

    With frm1
		.txtMode.value = Parent.UID_M0002											'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
	End With
	
    '-----------------------
    'Data manipulate area
    '-----------------------
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
   
	With frm1.vspdData3

		For IntRows = .MaxRows To 1 Step -1
    
			.Row = IntRows
			.Col = 0
			
			Select Case .Text
		    
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					
					strVal = ""
					
					If .Text = ggoSpread.InsertFlag Then
						strVal = strVal & "CREATE" & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
					Else
						strVal = strVal & "UPDATE" & iColSep				'⊙: U=Update
					End If
									    			      		            
					strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
					.Col = C_ProdtOrderNo3	' Production Order No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_OperationCd3	' Opr No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_WorkCd3
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					' Order Status
					.Col = C_OrderStatus3
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_PlanStartDt3
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					.Col = C_PlanEndDt3
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					' Run Time
					strVal = strVal & iColSep
					.Col = C_RoutOrder3
					DblRout = UCase(Trim(.Value))
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_InsideFlg3
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_MilestoneFlg3
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_WCCd3
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_InspFlg3
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_CCFCost3
					DblPrc = UNICDbl(.Value)
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
					.Col = C_CCFAmt3
					DblAmt = UNICDbl(.Value)
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
					.Col = C_Currency3
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_Tax3
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_BpCd3
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					'RowCount
					strVal = strVal & IntRows & iRowSep
					
					.Col = C_InsideFlg3
					If .Text = "N" Then

					    IF DblRout = "I" Then
							
							If DblPrc <= UNICDbl(0) Then
								Call DisplayMsgBox("189306", "x", "x", "x")
								Call LayerShowHide(0)
								.EditMode = True
								Call GetHiddenFocus(IntRows,C_CCFCost3)
								Exit Function
							ElseIf DblAmt <= UNICDbl(0) Then
								Call DisplayMsgBox("189306", "x", "x", "x")
								Call LayerShowHide(0)
								.EditMode = True
								Call GetHiddenFocus(IntRows,C_CCFAmt3)
								Exit Function
							End IF	
								
						End If
					End If
										
			    Case ggoSpread.DeleteFlag
					
					strDel = ""	
							
					strDel = strDel & "DELETE" & iColSep				'⊙: D=Delete
				
					strDel = strDel & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
					.Col = C_ProdtOrderNo3	' Production Order No.
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_OperationCd3	' Opr No.
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_WorkCd3
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					' Order Status
					strDel = strDel & iColSep
					.Col = C_PlanStartDt3
					strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep
					.Col = C_PlanEndDt3
					strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep
					' Run Time
					strDel = strDel & iColSep
					.Col = C_RoutOrder3
					DblRout = UCase(Trim(.Value))
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_InsideFlg3
					strDel = strDel & Trim(.Text) & iColSep
					.Col = C_MilestoneFlg3
					strDel = strDel & Trim(.Text) & iColSep
					.Col = C_WCCd3
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_InspFlg3
					strDel = strDel & Trim(.Text) & iColSep
					.Col = C_CCFCost3
					DblPrc = UNICDbl(.Value)
					strDel = strDel & UNIConvNum(Trim(.Text),0) & iColSep
					.Col = C_CCFAmt3
					DblAmt = UNICDbl(.Value)
					strDel = strDel & UNIConvNum(Trim(.Text),0) & iColSep
					.Col = C_Currency3
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_Tax3
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_BpCd3
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					'Row Count
					strDel = strDel & IntRows & iRowSep
					
					.Col = C_InsideFlg3
					If .Text = "N" Then

					    IF DblRout = "I" Then
							
							If DblPrc <= UNICDbl(0) Then
								Call DisplayMsgBox("189306", "x", "x", "x")
								Call GetHiddenFocus(IntRows,C_CCFCost3)
								Exit Function
							ElseIf DblAmt <= UNICDbl(0) Then
								Call DisplayMsgBox("189306", "x", "x", "x")
								Call GetHiddenFocus(IntRows,C_CCFAmt3)
								Exit Function
							End IF	
								
						End If
					End If

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
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 

    DbSave = True                                                           ' ⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0

	ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0
	ggoSpread.source = frm1.vspddata3
    frm1.vspdData3.MaxRows = 0
	
	Call RemovedivTextArea
	
	Call MainQuery
	
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
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 

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
'==============================================================================
' Function : GetHiddenFocus
' Description : 에러발생시 Hidden Spread Sheet를 찾아 SheetFocus에 값을 넘겨줌.
'==============================================================================
Function GetHiddenFocus(lRow, lCol)
	Dim lRows1, lRows2						'Quantity of the Hidden Data Keys Referenced by FindData Function
	Dim strHdnOrderNo, strHdnOprNo			'Variable of Hidden Keys
	Dim strOrderNo, strOprNo				'Variable of Visible Sheet Keys		
	
	If Trim(lCol) = "" Then
		lCol = C_OperationCd2				'If Value of Column is not passed, Assign Value of the First Column in Second Spread Sheet
	End If
	'Find Key Datas in Hidden Spread Sheet
	With frm1.vspdData3
		.Row = lRow
		.Col = C_ProdtOrderNo3
		strHdnOrderNo = Trim(.Text)
		.Col = C_OperationCd3
		strHdnOprNo = Trim(.Text)
	End With
	'Compare Key Datas to Visible Spread Sheets
	With frm1
		For lRows1 = 1 To .vspdData1.MaxRows
			.vspdData1.Row = lRows1
			.vspdData1.Col = C_OdrNo
			If Trim(.vspdData1.Text) = strHdnOrderNo Then
				.vspdData1.focus
				.vspdData1.Action = 0
				lgOldRow = lRows1			'※ If this line is omitted, program could not query Data When errors occur
				.vspdData2.MaxRows = 0
				ggoSpread.Source = .vspdData2
				If CopyFromHSheet(strHdnOrderNo) = True Then
				    For lRows2 = 1 To .vspdData2.MaxRows
						.vspdData2.Row = lRows2
						.vspdData2.Col = C_ProdtOrderNo2
						strOrderNo = .vspdData2.Text
						.vspdData2.Col = C_OperationCd2
						strOprNo = .vspdData2.Text
						'Find Key Datas in Second Sheet and then Focus the Cell 
						If Trim(strHdnOrderNo) = Trim(strOrderNo) And Trim(strHdnOprNo) = Trim(strOprNo) Then
							Call SheetFocus(lRows2, lCol)
							Exit Function
						End If
				    Next
				End If
			End If
		Next
	End With
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData2.focus
	frm1.vspdData2.Row = lRow
	frm1.vspdData2.Col = lCol
	frm1.vspdData2.Action = 0
	frm1.vspdData2.SelStart = 0
	frm1.vspdData2.SelLength = len(frm1.vspdData2.Text)
End Function

'==============================================================================
' Function : CheckValidOprNo
' Description : Opr No Check
'==============================================================================
Function CheckValidOprNo(ByVal pvOprNo, ByVal pvRow)
	Dim iIntCnt, iStrPrevOprNo, iStrNextOprNo
	Dim iStrValue
	
	CheckValidOprNo = True
	iStrPrevOprNo = "" : iStrNextOprNo = "zzz"

	For iIntCnt = pvRow - 1 To 1 Step -1
		Call frm1.vspdData2.GetText(0, iIntCnt, iStrValue)
		If iStrValue <> ggoSpread.DeleteFlag Then
			Call frm1.vspdData2.GetText(C_OperationCd2, iIntCnt, iStrPrevOprNo)
			If iStrPrevOprNo <> "" Then
				Exit For
			End If
		End If
	Next

	For iIntCnt = pvRow + 1 To frm1.vspdData2.MaxRows Step 1
		Call frm1.vspdData2.GetText(0, iIntCnt, iStrValue)
		If iStrValue <> ggoSpread.DeleteFlag Then
			Call frm1.vspdData2.GetText(C_OperationCd2, iIntCnt, iStrNextOprNo)
			If iStrNextOprNo <> "" Then
				Exit For
			Else
				iStrNextOprNo = "zzz"
			End If
		End If
	Next
		
	If pvOprNo >= iStrNextOprNo Or pvOprNo <= iStrPrevOprNo Then
		If iStrPrevOprNo = "" Then
			Call DisplayMsgBox("181220", "X", iStrNextOprNo, "X")
		ElseIf iStrNextOprNo = "zzz" Then
			Call DisplayMsgBox("181219", "X", iStrPrevOprNo, "X")
		Else
			Call DisplayMsgBox("181218", "X", iStrPrevOprNo, iStrNextOprNo)
		End If
		CheckValidOprNo = False
	End If
	
End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정별계획관리</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPartRef()">부품내역</A> | <A href="vbscript:OpenAltOprRef()">공정변경이력</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="TEXT" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>		
									<TD CLASS=TD5 NOWRAP>작업계획일자</TD>
									<TD CLASS="TD6">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtProdFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="시작일"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtProdtODt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료일"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="TEXT" NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>	
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="TEXT" NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>		
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
								    <TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="TEXT" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="TEXT" NAME="txtWCCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConWC()"> <INPUT TYPE=TEXT ID="txtWCNm" NAME="arrCond" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>지시구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderType" ALT="지시구분" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>지시상태</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderStatus" ALT="지시상태" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>작업</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboJobCd" ALT="작업" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>			 
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 ID = "A" width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 ID = "B" WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</Table>	
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD  HEIGHT=3></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD>&nbsp;</TD>
					<TD><!--<BUTTON NAME="btnRec" CLASS="CLSMBTN" ONCLICK="vbscript:RecMes()" >MES정보수신</BUTTON> -->
					&nbsp;<BUTTON NAME="btnReg" CLASS="CLSMBTN" ONCLICK="vbscript:RegProd()" >공정일괄추가</BUTTON>
					<TD>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtOpr" tag="24">
<INPUT TYPE=HIDDEN NAME="txtProdtOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="txtOprNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hWcCd" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderType" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdfromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hProdtoDt" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrderStatus" tag="24"><INPUT TYPE=HIDDEN NAME="hJobCd" tag="24" TABINDEX = "-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=100% name=vspdData3 ID = "C" width=100% TABINDEX = "-1"><PARAM NAME="MaxCols" VALUE="0" > <PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
