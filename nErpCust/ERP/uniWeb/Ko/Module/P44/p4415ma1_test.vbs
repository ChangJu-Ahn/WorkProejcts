
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

'Grid 1 - Order Header
Const BIZ_PGM_QRY1_ID	= "p4415mb1_test.asp"						'☆: Head Query 비지니스 로직 ASP명 
'Grid 2 - Production Results
Const BIZ_PGM_QRY2_ID	= "p4415mb2_test.asp"						'☆: 비지니스 로직 ASP명 
'Post Production Results
Const BIZ_PGM_SAVE_ID	= "p4415mb3_test.asp"						
'Shift Header
Const BIZ_PGM_SHIFT		= "p4415mb5_test.asp"						'☆: 비지니스 로직 ASP명 
'Jump (E)Production Order 
Const BIZ_PGM_JUMPREWORKRUN_ID = "p4111ma1"
'Jump (E)Resource Consumption (By Order)
Const BIZ_PGM_JUMPORDRSCCOMPT_ID = "p4712ma1"

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Order Header
Dim C_ProdtOrderNo			
Dim C_OprNo					
Dim C_ItemCd				
Dim C_ItemNm				
Dim C_Spec					
Dim C_ProdtOrderQty			
Dim C_ProdtOrderUnit		
Dim C_ProdQtyInOrderUnit	
Dim C_RemainQty				
Dim C_GoodQtyInOrderUnit	
Dim C_BadQtyInOrderUnit		
Dim C_InspGoodQtyInOrderUnit
Dim C_InspBadQtyInOrderUnit	
Dim C_RcptQtyInOrderUnit	
Dim C_PlanStartDt			
Dim C_PlanComptDt			
Dim C_ReleaseDt				
Dim C_RealStartDt			
Dim C_RoutNo				
Dim C_WcCd					
Dim C_WcNm					
Dim C_JobCd					
Dim C_JobDesc				
Dim C_RoutOrder						
Dim C_OrderStatus
Dim C_OrderStatusNm			
Dim C_MilestoneFlg			
Dim C_InsideFlag			
Dim C_InsideFlagNm			
Dim C_TrackingNo			
Dim C_ProdtOrderType		
Dim C_AutoRcptFlg			
Dim C_LotReq
Dim C_LotGenMthd				
Dim C_ProdInspReq			
Dim C_FinalInspReq
Dim C_ItemGroupCd
Dim C_ItemGroupNm
Dim C_ParentOrderNo
Dim C_ParentOprNo
Dim C_OrginalOrderNo
Dim C_OrginalOprNo
Dim C_ReworkPrevQty			

' Grid 2(vspdData2) - Results
Dim C_ReportDt				
Dim C_ReportType			
Dim C_ShiftId				
Dim C_ProdQty				
Dim C_ReasonCd				
Dim C_ReasonDesc			
Dim C_Remark1				
Dim C_LotNo					
Dim C_LotSubNo				
Dim C_RcptDocumentNo		
Dim C_IssueDocumentNo		
Dim C_InspReqNo				
Dim C_Insp_Good_Qty1			
Dim C_Insp_Bad_Qty1			
Dim C_Rcpt_Qty1		
' Hidden
Dim C_ProdtOrderNo1			
Dim C_OprNo1				
Dim C_Sequence				
Dim C_MilestoneFlg1			
Dim C_InsideFlag1			
Dim C_AutoRcptFlg1			
Dim C_LotReq1				
Dim C_LotGenMthd1
Dim C_ProdInspReq1			
Dim C_FinalInspReq1			
Dim C_RoutOrder1			

' Grid 3(vspdData3) - Hidden
Dim C_ReportDt2				
Dim C_ReportType2			
Dim C_ShiftId2				
Dim C_ProdQty2				
Dim C_ReasonCd2				
Dim C_ReasonDesc2			
Dim C_Remark2				
Dim C_LotNo2				
Dim C_LotSubNo2				
Dim C_RcptDocumentNo2		
Dim C_IssueDocumentNo2		
Dim C_InspReqNo2			
Dim C_Insp_Good_Qty2		
Dim C_Insp_Bad_Qty2			
Dim C_Rcpt_Qty2		
Dim C_ProdtOrderNo2			
Dim C_OprNo2				
Dim C_Sequence2				
Dim C_MilestoneFlg2			
Dim C_InsideFlag2			
Dim C_AutoRcptFlg2			
Dim C_LotReq2
Dim C_LotGenMthd2				
Dim C_ProdInspReq2			
Dim C_FinalInspReq2			
Dim C_RoutOrder2			

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntFlgMode								'Variable is for Operation Status
Dim lgIntPrevKey
Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgLngCurRows
Dim lgCurrRow
Dim lgShift
'==========================================  1.2.3 Global Variable값 정의  ==================================
'============================================================================================================
'----------------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgOldRow    
Dim lgSortKey1   
Dim lgSortKey2 
'++++++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""							'initializes Previous Key
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
    frm1.txtProdFromDt.text = StartDate
    frm1.txtProdToDt.text   = EndDate
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)
	
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "5", "0")
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
	
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20050825", ,Parent.gAllowDragDropSpread
    
			.ReDraw = false
    
			.MaxCols = C_ReworkPrevQty +1											'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
			        
			Call GetSpreadColumnPos("A")
	
			ggoSpread.SSSetEdit		C_ProdtOrderNo, "제조오더번호", 18
			ggoSpread.SSSetEdit		C_OprNo, "공정", 8
			ggoSpread.SSSetEdit		C_ItemCd, "품목", 18
			ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25
			ggoSpread.SSSetEdit		C_Spec, "규격", 25				'5
			ggoSpread.SSSetFloat	C_ProdtOrderQty,"오더수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ProdtOrderUnit, "오더단위", 8,,,3	
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit,"실적수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty,"잔량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit,"양품수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInOrderUnit,"불량수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_InspGoodQtyInOrderUnit,"품질양품",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_InspBadQtyInOrderUnit,"품질불량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit,"입고수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_PlanStartDt, "착수예정일", 11, 2, parent.gDateFormat	'15
			ggoSpread.SSSetDate 	C_PlanComptDt, "완료예정일", 11, 2, parent.gDateFormat	
			ggoSpread.SSSetDate 	C_ReleaseDt, "작업지시일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_RealStartDt, "실착수일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_RoutNo, "라우팅", 10 
			ggoSpread.SSSetEdit		C_WcCd, "작업장", 10			'20
			ggoSpread.SSSetEdit		C_WcNm, "작업장명", 20
			ggoSpread.SSSetEdit		C_JobCd, "작업", 8
			ggoSpread.SSSetEdit		C_JobDesc, "작업명", 20
			ggoSpread.SSSetEdit		C_RoutOrder, "작업순서", 8
			ggoSpread.SSSetEdit		C_OrderStatus, "지시상태", 10
			ggoSpread.SSSetEdit		C_OrderStatusNm, "지시상태", 10
			ggoSpread.SSSetEdit		C_MilestoneFlg, "Milestone", 10
			ggoSpread.SSSetEdit		C_InsideFlag, "사내/외", 10	
			ggoSpread.SSSetEdit		C_InsideFlagNm, "사내/외", 10
			ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 25,,,30
			ggoSpread.SSSetEdit		C_ProdtOrderType, "지시구분", 10
			ggoSpread.SSSetEdit		C_AutoRcptFlg, "", 10
			ggoSpread.SSSetEdit		C_LotReq, "", 10
			ggoSpread.SSSetEdit		C_LotGenMthd, "", 10
			ggoSpread.SSSetEdit		C_ProdInspReq, "공정검사", 8
			ggoSpread.SSSetEdit		C_FinalInspReq, "", 10				'39
			ggoSpread.SSSetEdit 	C_ItemGroupCd, "품목그룹",	15
			ggoSpread.SSSetEdit		C_ItemGroupNm, "품목그룹명", 30
			ggoSpread.SSSetEdit		C_ParentOrderNo,	"상위오더번호", 18
			ggoSpread.SSSetEdit		C_ParentOprNo,		"상위공정", 8
			ggoSpread.SSSetEdit		C_OrginalOrderNo,	"기존오더번호", 18
			ggoSpread.SSSetEdit		C_OrginalOprNo,		"기존공정", 8
			ggoSpread.SSSetFloat	C_ReworkPrevQty,	"재작업수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_RoutOrder, C_RoutOrder, True)
			Call ggoSpread.SSSetColHidden(C_OrderStatus, C_OrderStatus, True)
			Call ggoSpread.SSSetColHidden(C_InsideFlag, C_InsideFlag, True)
			Call ggoSpread.SSSetColHidden(C_AutoRcptFlg, C_AutoRcptFlg, True)
			Call ggoSpread.SSSetColHidden(C_LotReq, C_LotReq, True)
			Call ggoSpread.SSSetColHidden(C_LotGenMthd, C_LotGenMthd, True)
			Call ggoSpread.SSSetColHidden(C_FinalInspReq, C_FinalInspReq, True)
			Call ggoSpread.SSSetColHidden(C_ReworkPrevQty, C_ReworkPrevQty, True)
			         
			ggoSpread.SSSetSplit2(3)
			
			Call SetSpreadLock("A")
				
			.ReDraw = true
    
		End With
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20040808", ,Parent.gAllowDragDropSpread
    
			.ReDraw = false
    
			.MaxCols = C_RoutOrder1 +1										'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
    
			Call GetSpreadColumnPos("B")  

			ggoSpread.SSSetDate 	C_ReportDt, "실적일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetCombo	C_ReportType, "양/불", 6
			ggoSpread.SSSetCombo	C_ShiftId, "Shift", 8
			ggoSpread.SSSetFloat	C_ProdQty, "실적수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo	C_ReasonCd, "불량코드", 10
			ggoSpread.SSSetCombo	C_ReasonDesc, "불량이유", 20
			ggoSpread.SSSetEdit		C_Remark1, "비고", 20,,,120	
			ggoSpread.SSSetEdit		C_LotNo, "Lot No.", 20,,,25,2
			ggoSpread.SSSetFloat	C_LotSubNo, "순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_RcptDocumentNo, "입고번호", 18,,,16,2
			ggoSpread.SSSetEdit		C_IssueDocumentNo, "출고번호", 18,,,16,2	
			ggoSpread.SSSetEdit		C_InspReqNo, "검사의뢰번호", 18,,,18,2
			ggoSpread.SSSetFloat	C_Insp_Good_Qty1, "품질양품",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Insp_Bad_Qty1, "품질불량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Rcpt_Qty1, "입고수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ProdtOrderNo1, "", 18
			ggoSpread.SSSetEdit		C_OprNo1	, "", 8	
			ggoSpread.SSSetFloat	C_Sequence, "순번", 7, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_MilestoneFlg1, "Milestone", 10	
			ggoSpread.SSSetEdit		C_InsideFlag1, "사내/외", 10	
			ggoSpread.SSSetEdit		C_AutoRcptFlg1, "", 10	
			ggoSpread.SSSetEdit		C_LotReq1, "", 10 
			ggoSpread.SSSetEdit		C_LotGenMthd1, "", 10							
			ggoSpread.SSSetEdit		C_ProdInspReq1, "", 10	
			ggoSpread.SSSetEdit		C_FinalInspReq1, "", 10	
			ggoSpread.SSSetEdit		C_RoutOrder1, "", 10	
	
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_ProdtOrderNo1, C_ProdtOrderNo1, True)
			Call ggoSpread.SSSetColHidden(C_OprNo1, C_OprNo1, True)
			Call ggoSpread.SSSetColHidden(C_Sequence, C_Sequence, True)
			Call ggoSpread.SSSetColHidden(C_MilestoneFlg1, C_MilestoneFlg1, True)
			Call ggoSpread.SSSetColHidden(C_InsideFlag1, C_InsideFlag1, True)
			Call ggoSpread.SSSetColHidden(C_AutoRcptFlg1, C_AutoRcptFlg1, True)
			Call ggoSpread.SSSetColHidden(C_LotReq1, C_LotGenMthd1, True)
			Call ggoSpread.SSSetColHidden(C_Rcpt_Qty1, C_Rcpt_Qty1, True)  ' hidden for rcpt_qty 20030403 kjp
			Call ggoSpread.SSSetColHidden(C_ProdInspReq1, C_ProdInspReq1, True)
			Call ggoSpread.SSSetColHidden(C_FinalInspReq1, C_FinalInspReq1, True)
			Call ggoSpread.SSSetColHidden(C_RoutOrder1, C_RoutOrder1, True)
	
			ggoSpread.SSSetSplit2(5)
			
			Call SetSpreadLock("B")
	
			.ReDraw = true
    
		End With
    End If
    
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 3 - Hidden Setting
		'------------------------------------------
		With frm1.vspdData3
			
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit
			
			.ReDraw = false
					
			.MaxCols = C_RoutOrder2 +1										'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
			
			ggoSpread.SSSetDate 	C_ReportDt2,	"실적일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_ReportType2,	"양/불", 6
			ggoSpread.SSSetEdit		C_ShiftId2,		"Shift", 8	
			ggoSpread.SSSetFloat	C_ProdQty2,		"실적수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ReasonCd2,	"불량코드", 10
			ggoSpread.SSSetEdit		C_ReasonDesc2,	"불량이유", 20
			ggoSpread.SSSetEdit		C_Remark2,		"비고", 20,,,120	
			ggoSpread.SSSetEdit		C_LotNo2,		"Lot No.", 20,,,25,2		
			ggoSpread.SSSetFloat	C_LotSubNo2,	"순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_RcptDocumentNo2, "입고번호", 18,,,16
			ggoSpread.SSSetEdit		C_IssueDocumentNo2, "출고번호", 18,,,16	
			ggoSpread.SSSetEdit		C_InspReqNo2,	"검사의뢰번호", 18,,,18
			ggoSpread.SSSetFloat	C_Insp_Good_Qty2, "품질양품",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Insp_Bad_Qty2, "품질불량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Rcpt_Qty2,	"입고수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ProdtOrderNo2, "오더번호", 18
			ggoSpread.SSSetFloat	C_Sequence2,	"순번", 8, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_MilestoneFlg2, "Milestone", 10
			ggoSpread.SSSetEdit		C_InsideFlag2,	"사내/외", 10	
			ggoSpread.SSSetEdit		C_AutoRcptFlg2, "", 10	
			ggoSpread.SSSetEdit		C_LotReq2, "", 10	
			ggoSpread.SSSetEdit		C_LotGenMthd2, "", 10							
			ggoSpread.SSSetEdit		C_ProdInspReq2, "", 10	
			ggoSpread.SSSetEdit		C_FinalInspReq2, "", 10	
			ggoSpread.SSSetEdit		C_RoutOrder2, "", 10		
			
			.ReDraw = true
				
			Call SetSpreadLock("C")
    
		End With
    End If

End Sub

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
			
		If pvSpdNo = "B" Then
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData2
			.vspdData2.ReDraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData2.Redraw = True
		End If
			
		If pvSpdNo = "C" Then
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData3
			.vspdData3.ReDraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData3.Redraw = True
		End If
    End With

End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal strLotReq, ByVal strLotGenMthd, ByVal strProdInspReq, ByVal strFinalInspReq, ByVal strAutoRcptFlg, ByVal strInsideFlg, ByVal strRoutOrder)

    With frm1.vspdData2
    
		.Redraw = False
		ggoSpread.Source = frm1.vspdData2
		
		ggoSpread.SpreadUnLock C_ReportDt,pvStartRow,C_ReportDt,pvEndRow
		ggoSpread.SSSetRequired C_ReportDt,					pvStartRow, pvEndRow    
		ggoSpread.SpreadUnLock C_ReportType,pvStartRow,C_ReportType,pvEndRow
		ggoSpread.SSSetRequired C_ReportType,				pvStartRow, pvEndRow
		ggoSpread.SpreadUnLock C_ShiftId,pvStartRow,C_ShiftId,pvEndRow
		ggoSpread.SSSetRequired C_ShiftId,					pvStartRow, pvEndRow
		ggoSpread.SpreadUnLock C_ProdQty,pvStartRow,C_ProdQty,pvEndRow
		ggoSpread.SSSetRequired C_ProdQty,					pvStartRow, pvEndRow
		
		ggoSpread.SpreadUnLock C_Remark1,pvStartRow,C_Remark1,pvEndRow
		
		frm1.vspdData2.Col = C_ReportType
		If frm1.vspdData2.Text <> "B" Then
			ggoSpread.SSSetProtected C_ReasonCd,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_ReasonDesc,			pvStartRow, pvEndRow
		Else
			ggoSpread.SpreadUnLock C_ReasonCd,pvStartRow,C_ReasonCd,pvEndRow
			ggoSpread.SSSetRequired C_ReasonCd,				pvStartRow, pvEndRow
			ggoSpread.SpreadUnLock C_ReasonDesc,pvStartRow,C_ReasonDesc,pvEndRow
			ggoSpread.SSSetRequired C_ReasonDesc,			pvStartRow, pvEndRow
		End If
		
		If strRoutOrder = "S" Or strRoutOrder = "L" Then
			If strLotReq <> "Y" or strAutoRcptFlg <> "Y" Then
				ggoSpread.SSSetProtected C_LotNo,				pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_LotSubNo,			pvStartRow, pvEndRow
			Else
				If strLotGenMthd = "M" Then
					ggoSpread.SpreadUnLock C_LotNo,pvStartRow,C_LotNo,pvEndRow
					ggoSpread.SpreadUnLock C_LotSubNo,pvStartRow,C_LotSubNo,pvEndRow
					ggoSpread.SSSetRequired C_LotNo,				pvStartRow, pvEndRow
					ggoSpread.SSSetRequired C_LotSubNo,				pvStartRow, pvEndRow
				Else
					ggoSpread.SpreadUnLock C_LotNo,pvStartRow,C_LotNo,pvEndRow
					ggoSpread.SpreadUnLock C_LotSubNo,pvStartRow,C_LotSubNo,pvEndRow
				End If
			End If		
		Else
			ggoSpread.SSSetProtected C_LotNo,				pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_LotSubNo,			pvStartRow, pvEndRow	
		End If	
		ggoSpread.SSSetProtected C_ReportDt,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RcptDocumentNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IssueDocumentNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspReqNo,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Insp_Good_Qty1,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Insp_Bad_Qty1,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Rcpt_Qty1,				pvStartRow, pvEndRow
				
		.Redraw = True
    
    End With
   
End Sub

'========================== 2.2.6 InitSpreadComboBox()  ========================================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()

    Dim strCboCd
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
		
	'****************************
	'List Minor code(Job Code)
	'****************************
	strCboCd =  "G" & vbTab & "B"
	
	'****************************
	'List Minor code(Reason Code)
	'****************************
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3221", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo strCboCd, C_ReportType
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ReasonCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ReasonDesc
    
End Sub

'==========================================  2.2.6 InitShiftCombo()  =======================================
'	Name : InitShiftCombo()
'	Description : Combo Display
'===========================================================================================================
Function InitShiftCombo()
	
	Dim strPlantCd
	Dim strShiftArr
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	strPlantCd = FilterVar(UCase(frm1.hPlantCd.value), "''", "S")
	
	'****************************
	'List Minor code(Reason Code)
	'****************************
    Call CommonQueryRs(" SHIFT_CD "," P_SHIFT_HEADER "," PLANT_CD = " & strPlantCd ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ShiftId
    
    strShiftArr = Split(lgF0,Chr(11))
    
    lgShift = strShiftArr(0)
	
End Function

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
 Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex
	
	With frm1.vspdData2
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_ReasonCd
			intIndex = .value
			.Col = C_ReasonDesc
			.value = intindex
		Next	
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		 C_ProdtOrderNo				= 1
		 C_OprNo					= 2
		 C_ItemCd					= 3
		 C_ItemNm					= 4
		 C_Spec						= 5
		 C_ProdtOrderQty			= 6
		 C_ProdtOrderUnit			= 7
		 C_ProdQtyInOrderUnit		= 8
		 C_RemainQty				= 9
		 C_GoodQtyInOrderUnit		= 10
		 C_BadQtyInOrderUnit		= 11
		 C_InspGoodQtyInOrderUnit	= 12
		 C_InspBadQtyInOrderUnit	= 13
		 C_RcptQtyInOrderUnit		= 14
		 C_PlanStartDt				= 15
		 C_PlanComptDt				= 16
		 C_ReleaseDt				= 17
		 C_RealStartDt				= 18
		 C_RoutNo					= 19
		 C_WcCd						= 20
		 C_WcNm						= 21
		 C_JobCd					= 22
		 C_JobDesc					= 23
		 C_RoutOrder				= 24
		 C_OrderStatus				= 25
		 C_OrderStatusNm			= 26
		 C_MilestoneFlg				= 27
		 C_InsideFlag				= 28
		 C_InsideFlagNm				= 29
		 C_TrackingNo				= 30
		 C_ProdtOrderType			= 31
		 C_AutoRcptFlg				= 32
		 C_LotReq					= 33
		 C_LotGenMthd				= 34
		 C_ProdInspReq				= 35
		 C_FinalInspReq				= 36
		 C_ItemGroupCd				= 37
		 C_ItemGroupNm				= 38
		 C_ParentOrderNo			= 39
		 C_ParentOprNo				= 40
		 C_OrginalOrderNo			= 41
		 C_OrginalOprNo				= 42
		 C_ReworkPrevQty			= 43
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Results
		 C_ReportDt					= 1
		 C_ReportType				= 2
		 C_ShiftId					= 3
		 C_ProdQty					= 4
		 C_ReasonCd					= 5
		 C_ReasonDesc				= 6
		 C_Remark1					= 7
		 C_LotNo					= 8
		 C_LotSubNo					= 9
		 C_RcptDocumentNo			= 10
		 C_IssueDocumentNo			= 11
		 C_InspReqNo				= 12
		 C_Insp_Good_Qty1			= 13
		 C_Insp_Bad_Qty1			= 14
		 C_Rcpt_Qty1				= 15
		' Hidden
		 C_ProdtOrderNo1			= 16
		 C_OprNo1					= 17
		 C_Sequence					= 18
		 C_MilestoneFlg1			= 19
		 C_InsideFlag1				= 20
		 C_AutoRcptFlg1				= 21
		 C_LotReq1					= 22
		 C_LotGenMthd1				= 23
		 C_ProdInspReq1				= 24
		 C_FinalInspReq1			= 25
		 C_RoutOrder1				= 26
	End If		 
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		' Grid 3(vspdData3) - Hidden
		 C_ReportDt2				= 1
		 C_ReportType2				= 2
		 C_ShiftId2					= 3
		 C_ProdQty2					= 4
		 C_ReasonCd2				= 5
		 C_ReasonDesc2				= 6
		 C_Remark2					= 7
		 C_LotNo2					= 8
		 C_LotSubNo2				= 9
		 C_RcptDocumentNo2			= 10
		 C_IssueDocumentNo2			= 11
		 C_InspReqNo2				= 12
		 C_Insp_Good_Qty2			= 13
		 C_Insp_Bad_Qty2			= 14
		 C_Rcpt_Qty2				= 15
		 C_ProdtOrderNo2			= 16
		 C_OprNo2					= 17
		 C_Sequence2				= 18
		 C_MilestoneFlg2			= 19
		 C_InsideFlag2				= 20
		 C_AutoRcptFlg2				= 21
		 C_LotReq2					= 22
		 C_LotGenMthd2				= 23
		 C_ProdInspReq2				= 24
		 C_FinalInspReq2			= 25
		 C_RoutOrder2				= 26
	End If
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'=================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
			 C_ProdtOrderNo				= iCurColumnPos(1)
			 C_OprNo					= iCurColumnPos(2)
			 C_ItemCd					= iCurColumnPos(3)
			 C_ItemNm					= iCurColumnPos(4)
			 C_Spec						= iCurColumnPos(5)
			 C_ProdtOrderQty			= iCurColumnPos(6)
			 C_ProdtOrderUnit			= iCurColumnPos(7)
			 C_ProdQtyInOrderUnit		= iCurColumnPos(8)
			 C_RemainQty				= iCurColumnPos(9)
			 C_GoodQtyInOrderUnit		= iCurColumnPos(10)
			 C_BadQtyInOrderUnit		= iCurColumnPos(11)
			 C_InspGoodQtyInOrderUnit	= iCurColumnPos(12)
			 C_InspBadQtyInOrderUnit	= iCurColumnPos(13)
			 C_RcptQtyInOrderUnit		= iCurColumnPos(14)
			 C_PlanStartDt				= iCurColumnPos(15)
			 C_PlanComptDt				= iCurColumnPos(16)
			 C_ReleaseDt				= iCurColumnPos(17)
			 C_RealStartDt				= iCurColumnPos(18)
			 C_RoutNo					= iCurColumnPos(19)
			 C_WcCd						= iCurColumnPos(20)
			 C_WcNm						= iCurColumnPos(21)
			 C_JobCd					= iCurColumnPos(22)
			 C_JobDesc					= iCurColumnPos(23)
			 C_RoutOrder				= iCurColumnPos(24)
			 C_OrderStatus				= iCurColumnPos(25)
			 C_OrderStatusNm			= iCurColumnPos(26)
			 C_MilestoneFlg				= iCurColumnPos(27)
			 C_InsideFlag				= iCurColumnPos(28)
			 C_InsideFlagNm				= iCurColumnPos(29)
			 C_TrackingNo				= iCurColumnPos(30)
			 C_ProdtOrderType			= iCurColumnPos(31)
			 C_AutoRcptFlg				= iCurColumnPos(32)
			 C_LotReq					= iCurColumnPos(33)
			 C_LotGenMthd				= iCurColumnPos(34)
			 C_ProdInspReq				= iCurColumnPos(35)
			 C_FinalInspReq				= iCurColumnPos(36)
			 C_ItemGroupCd				= iCurColumnPos(37)
			 C_ItemGroupNm				= iCurColumnPos(38)
			 C_ParentOrderNo			= iCurColumnPos(39)
			 C_ParentOprNo				= iCurColumnPos(40)
			 C_OrginalOrderNo			= iCurColumnPos(41)
			 C_OrginalOprNo				= iCurColumnPos(42)
			 C_ReworkPrevQty			= iCurColumnPos(43)


		Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			' Grid 2(vspdData2) - Results
			 C_ReportDt					= iCurColumnPos(1)
			 C_ReportType				= iCurColumnPos(2)
			 C_ShiftId					= iCurColumnPos(3)
			 C_ProdQty					= iCurColumnPos(4)
			 C_ReasonCd					= iCurColumnPos(5)
			 C_ReasonDesc				= iCurColumnPos(6)
			 C_Remark1					= iCurColumnPos(7)
			 C_LotNo					= iCurColumnPos(8)
			 C_LotSubNo					= iCurColumnPos(9)
			 C_RcptDocumentNo			= iCurColumnPos(10)
			 C_IssueDocumentNo			= iCurColumnPos(11)
			 C_InspReqNo				= iCurColumnPos(12)
			 C_Insp_Good_Qty1			= iCurColumnPos(13)
			 C_Insp_Bad_Qty1			= iCurColumnPos(14)
			 C_Rcpt_Qty1				= iCurColumnPos(15)
			' Hidden
			 C_ProdtOrderNo1			= iCurColumnPos(16)
			 C_OprNo1					= iCurColumnPos(17)
			 C_Sequence					= iCurColumnPos(18)
			 C_MilestoneFlg1			= iCurColumnPos(19)
			 C_InsideFlag1				= iCurColumnPos(20)
			 C_AutoRcptFlg1				= iCurColumnPos(21)
			 C_LotReq1					= iCurColumnPos(22)
			 C_LotGenMthd1				= iCurColumnPos(23)
			 C_ProdInspReq1				= iCurColumnPos(24)
			 C_FinalInspReq1			= iCurColumnPos(25)
			 C_RoutOrder1				= iCurColumnPos(26)

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
'++++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  ------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""
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
		Call SetPlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
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
	arrParam(3) = "RL"
	arrParam(4) = "ST"
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

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
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
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
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

'------------------------------------------  OpenItemGroup()  -------------------------------------------------
'	Name : OpenItemGroup()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
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


'------------------------------------------  OpenWcCd()  ------------------------------------------------
'	Name : OpenWcCd()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
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

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
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
	arrParam(2) = Trim(frm1.txtItemCd.value)
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

'------------------------------------------  OpenPartRef()  ----------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPartRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4311RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4311RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
		frm1.vspdData1.Col = C_OprNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If	

	IsOpenPop = True
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenOprRef()  -----------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  ----------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Receipt Reference
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
End Function

'------------------------------------------  OpenConsumRef()  --------------------------------------------
'	Name : OpenConsumRef()
'	Description : Part Consumption Reference
'---------------------------------------------------------------------------------------------------------
Function OpenConsumRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4412RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4412RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent,arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenReworkRef()  --------------------------------------------
'	Name : OpenReworkRef()
'	Description : Rework Order History Reference
'---------------------------------------------------------------------------------------------------------
Function OpenReworkRef()

	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4413RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4413RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.hPlantCd.value)
	
	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ItemCd
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
		'opr_no
		arrParam(3) = ""									'☜: 조회 조건 데이타 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenBackFlushRef()  -----------------------------------------
'	Name : OpenBackFlushRef()
'	Description : BackFlush Simmulation Reference
'---------------------------------------------------------------------------------------------------------
Function OpenBackFlushRef()
	
	Dim arrRet
	Dim IntRows
	Dim strVal
	Dim iCalledAspName
	Dim strFlag
	
	If IsOpenPop = True Then Exit Function
	
	strVal = ""
		
	With frm1.vspdData3
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			strFlag = .Text
			.Col = C_ProdQty2		' Produced Qty
			
			If UNICDbl(.Text) > CDbl(0) and strFlag = ggoSpread.InsertFlag Then

				strVal = strVal & frm1.hPlantCd.value & parent.gColSep
				.Col = C_ProdtOrderNo2			
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_OprNo2
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_ProdQty2
				strVal = strVal & UniConvNum(.Text,0) & parent.gRowSep
			End If
		Next
	End With
	
	frm1.txtSpread.value = strVal
	
	iCalledAspName = AskPRAspName("P4400RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4400RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'===============================================================================
' Function Name : JumpReworkRun
' Function Desc : Jump to (E)Production Order(Single)
'========================================================================================
Function JumpReworkRun()
	
	Dim strProdtOrdNo, strOprNo
	Dim strItemCd
	Dim DblJumpQty, DblInspBadQty, DblBadQty, DblReworkQty
	Dim strTrackingNo
	
	If lgIntFlgMode = parent.OPMD_CMODE Then		
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData1 
		.Row = .ActiveRow
		.Col = C_InspBadQtyInOrderUnit
		DblInspBadQty = UNICDbl(.Text)
		.Col = C_BadQtyInOrderUnit	
		DblBadQty = UNICDbl(.Text)
		.Col = C_ReworkPrevQty	
		DblReworkQty = UNICDbl(.Text)
		
		DblJumpQty = DblInspBadQty + DblBadQty - DblReworkQty
		'Error Check -  Whether Defect Qty is greater than zero
		
		If DblInspBadQty + DblBadQty = Cdbl(0) Then
			Call DisplayMsgBox("189247", "x", "x", "x")
			Exit Function 
		End If
		
		If DblJumpQty <= 0 Then
			Call DisplayMsgBox("189248", "x", "x", "x")
			Exit Function 
		End If
		
		.Col = C_ProdtOrderNo
		strProdtOrdNo = UCase(Trim(.Text))
		.Col = C_OprNo
		strOprNo = UCase(Trim(.Text))
		.Col = C_ItemCd
		strItemCd = UCase(Trim(.Text))
		.Col = C_TrackingNo
		strTrackingNo = UCase(Trim(.Text))
		
	End With
	
End Function

'========================================================================================
' Function Name : JumpOrdRscComptRun
' Function Desc : Jump to (E)Production Order(Single)
'========================================================================================
Function JumpOrdRscComptRun()
	
	Dim strProdtOrdNo, strOprNo
	
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
	
	With frm1.vspdData1 
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		strProdtOrdNo = UCase(Trim(.Text))
		.Col = C_OprNo
		strOprNo = UCase(Trim(.Text))
		
	End With	
		
	WriteCookie "txtPlantCd", UCase(Trim(frm1.hPlantCd.value))
	WriteCookie "txtPlantNm", UCase(Trim(frm1.txtPlantNm.value))
	WriteCookie "txtProdOrderNo", strProdtOrdNo
	WriteCookie "txtOprNo", strOprNo
	
	PgmJump(BIZ_PGM_JUMPORDRSCCOMPT_ID)
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
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

'------------------------------------------  SetItemInfo()  -------------------------------------------
'	Name : SetItemGroup()
'	Description : Item Group Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetTrackingNo()  ----------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	
	frm1.txtTrackingNo.Value = arrRet(0)
	
End Function

'------------------------------------------  txtPlantCd_OnChange -----------------------------------------
'	Name : txtPlantCd_OnChange()
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlantCd_OnChange()
	
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

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
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
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdToDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtReportDT_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReportDT_DblClick(Button)
    If Button = 1 Then
        frm1.txtReportDT.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReportDT.Focus
    End If
End Sub
'------------------------------------------  txtReportDT_OnChange -----------------------------------------
'	Name : txtReportDT_OnChange()
'	Description : vspddata2의 ReportDt 업데이트 
'----------------------------------------------------------------------------------------------------------
Sub txtReportDT_Change()
	dim intRows
	if frm1.txtReportDt.text = "" then
	else
		with frm1.vspdData2
		for intRows = 1 to .maxRows    
			.Row = intRows
			.Col = 0
			if .text = ggoSpread.InsertFlag then
				.Col = C_ReportDt
				.text = frm1.txtReportDt.text
			End if
		next
		end with
		with frm1.vspdData3
		for intRows = 1 to .maxRows
			.Row = intRows
			.Col = 0
			if .text = ggoSpread.InsertFlag then
				.Col = C_ReportDt2
				.text = frm1.txtReportDt.text
			End if
		next 
		end with
	
	End if
End Sub
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'========================================================================================
' Function Name : JumpReworkRun
' Function Desc : Jump to (E)Production Order(Single)
'========================================================================================
Function JumpReworkRun()
	
	Dim strProdtOrdNo, strOprNo
	Dim strItemCd
	Dim DblJumpQty, DblInspBadQty, DblBadQty, DblReworkQty
	Dim strTrackingNo
	
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
	
	With frm1.vspdData1 
		.Row = .ActiveRow
		.Col = C_InspBadQtyInOrderUnit
		DblInspBadQty = UNICDbl(.Text)
		.Col = C_BadQtyInOrderUnit	
		DblBadQty = UNICDbl(.Text)
		.Col = C_ReworkPrevQty	
		DblReworkQty = UNICDbl(.Text)
		
		DblJumpQty = DblInspBadQty + DblBadQty - DblReworkQty
		'Error Check -  Whether Defect Qty is greater than zero
		
		If DblInspBadQty + DblBadQty = Cdbl(0) Then
			Call DisplayMsgBox("189247", "x", "x", "x")
			Exit Function 
		End If
		
		If DblJumpQty <= 0 Then
			Call DisplayMsgBox("189248", "x", "x", "x")
			Exit Function 
		End If
		
		.Col = C_ProdtOrderNo
		strProdtOrdNo = UCase(Trim(.Text))
		.Col = C_OprNo
		strOprNo = UCase(Trim(.Text))
		.Col = C_ItemCd
		strItemCd = UCase(Trim(.Text))
		.Col = C_TrackingNo
		strTrackingNo = UCase(Trim(.Text))
		
	End With	
		
	WriteCookie "txtPlantCd", UCase(Trim(frm1.hPlantCd.value))
	WriteCookie "txtPlantNm", UCase(Trim(frm1.txtPlantNm.value))
	WriteCookie "txtItemCd", strItemCd
	WriteCookie "txtProdOrderNo", strProdtOrdNo
	WriteCookie "txtOprNo", strOprNo
	WriteCookie "txtJumpQty", UniFormatNumber(DblJumpQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	WriteCookie "txtTrackingNo", strTrackingNo
	
	PgmJump(BIZ_PGM_JUMPREWORKRUN_ID)
	
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
'**********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  *********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*******************************************************************************************************

'******************************  3.2.1 Object Tag 처리  ************************************************
'	Window에 발생 하는 모든 Even 처리	
'*******************************************************************************************************

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
    '---------------------- 
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1
    
  	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
  	
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
 		
 		frm1.vspdData2.MaxRows = 0
			
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

			frm1.vspdData2.MaxRows = 0
			
			If DbDtlQuery(Row) = False Then	
	
				Call RestoreToolBar()
				Exit Sub
			End If	

		End If
	 	'------ Developer Coding part (End)
	
 	End If
 	
	
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    Dim strInsideFlag
    
    If Row = NewRow Then
        Exit Sub
    End If

	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If

	frm1.vspdData1.Row = NewRow
	
	If SetBlnInsertRow(frm1.vspdData1.Row) = True Then
		Call SetToolBar("11001101000111")										'⊙: 버튼 툴바 제어 
	ElseIf SetBlnInsertRow(frm1.vspdData1.Row) = False Then
		Call SetToolBar("11001000000111")										'⊙: 버튼 툴바 제어 
	End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2
	
	If SetBlnInsertRow(frm1.vspdData1.ActiveRow) = True Then
		Call SetPopupMenuItemInf("1001111111")         '화면별 설정 
	ElseIf SetBlnInsertRow(frm1.vspdData1.ActiveRow) = False Then	
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
    End If
    
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
 			
 	End If

End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
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

	Dim strItemCd
	Dim strHndItemCd, strHndOprNo
	Dim i
	Dim strReqDt, strEndDt
	Dim strProdtOrderNo, strOprNo
	Dim LngFindRow
	
	Dim strAutoRcptFlg, strLotReq, strRoutOrder, strLotGenMthd
	With frm1.vspdData2

		Select Case Col

		    Case C_ReportType
		       	
		       	ggoSpread.Source = frm1.vspdData2
				.Col = Col
				.Row = Row

				.Col = C_AutoRcptFlg1
				strAutoRcptFlg = .Text
				.Col = C_LotReq1
				strLotReq = .Text
				.Col = C_RoutOrder1
				strRoutOrder = .Text
				.Col = C_LotGenMthd1
				strLotGenMthd = .Text
				.ReDraw = False

				.Col = C_ReportType
				If Trim(.Text) = "G" Then
												
					ggoSpread.SpreadLock C_ReasonCd, Row, C_ReasonCd, Row
					ggoSpread.SpreadLock C_ReasonDesc, Row, C_ReasonDesc, Row
					ggoSpread.SSSetProtected C_ReasonCd, Row, Row
					ggoSpread.SSSetProtected C_ReasonDesc, Row, Row
                    
                    .Col = C_ReasonCd
					.Text = ""
					.Col = C_ReasonDesc
					.Text = ""

					If strLotReq <> "Y" or strAutoRcptFlg <> "Y" or (strRoutOrder <> "S" and strRoutOrder <> "L")  Then
						ggoSpread.SSSetProtected C_LotNo, Row, Row
						ggoSpread.SSSetProtected C_LotSubNo, Row, Row
					Else
						If strLotGenMthd = "M" Then
							ggoSpread.SpreadUnLock C_LotNo, Row, C_LotNo, Row
							ggoSpread.SpreadUnLock C_LotSubNo, Row, C_LotSubNo, Row
							ggoSpread.SSSetRequired C_LotNo, Row, Row
							ggoSpread.SSSetRequired C_LotSubNo, Row, Row
						Else
							ggoSpread.SpreadUnLock C_LotNo, Row, C_LotNo, Row
							ggoSpread.SpreadUnLock C_LotSubNo, Row, C_LotNo, Row
						End If
					End If
					
				Else
					ggoSpread.SpreadUnLock C_ReasonCd, Row, C_ReasonCd, Row
					ggoSpread.SpreadUnLock C_ReasonDesc, Row, C_ReasonDesc, Row
					ggoSpread.SSSetRequired C_ReasonCd, Row, Row
					ggoSpread.SSSetRequired C_ReasonDesc, Row, Row
					
					ggoSpread.SSSetProtected C_LotNo, Row, Row
					ggoSpread.SSSetProtected C_LotSubNo, Row, Row
					
				End If
			
				.ReDraw = True
			
				CopyToHSheet Row
		
				frm1.vspdData2.Row = Row
				frm1.vspdData2.Col = C_ProdtOrderNo1
				strProdtOrderNo = Trim(frm1.vspdData2.Text)
				frm1.vspdData2.Col = C_OprNo1
				strOprNo = Trim(frm1.vspdData2.Text)
				LngFindRow = FindRow(strProdtOrderNo, strOprNo)
				If LngFindRow > 0 Then
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.UpdateRow LngFindRow
				End If				
		    
		    Case C_ProdQty
		    
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row		    
		    
				frm1.vspdData2.Row = Row
				frm1.vspdData2.Col = C_ProdtOrderNo1
				strProdtOrderNo = Trim(frm1.vspdData2.Text)
				frm1.vspdData2.Col = C_OprNo1
				strOprNo = Trim(frm1.vspdData2.Text)
				LngFindRow = FindRow(strProdtOrderNo, strOprNo)
				If LngFindRow > 0 Then
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.UpdateRow LngFindRow
				End If
		    
		    Case C_ShiftId, C_ReasonCd, C_ReasonDesc, C_Remark1

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row
		    
				frm1.vspdData2.Row = Row
				frm1.vspdData2.Col = C_ProdtOrderNo1
				strProdtOrderNo = Trim(frm1.vspdData2.Text)
				frm1.vspdData2.Col = C_OprNo1
				strOprNo = Trim(frm1.vspdData2.Text)
				LngFindRow = FindRow(strProdtOrderNo, strOprNo)
				If LngFindRow > 0 Then
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.UpdateRow LngFindRow
				End If
		    
		    Case C_LotNo, C_RcptDocumentNo, C_IssueDocumentNo, C_InspReqNo, C_LotSubNo
		    
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row		    
				
				frm1.vspdData2.Row = Row
				frm1.vspdData2.Col = C_ProdtOrderNo1
				strProdtOrderNo = Trim(frm1.vspdData2.Text)
				frm1.vspdData2.Col = C_OprNo1
				strOprNo = Trim(frm1.vspdData2.Text)
				LngFindRow = FindRow(strProdtOrderNo, strOprNo)
				If LngFindRow > 0 Then
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.UpdateRow LngFindRow
				End If
				
		End Select

	End With

End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim strLotReq
	Dim strAutoRcptFlg
	Dim strRoutOrder
	Dim strProdtOrderNo, strOprNo
	Dim LngFindRow

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

	With frm1.vspdData2

		.Row = Row
		Select Case Col
		
			Case  C_ReasonCd
				.Col = Col
				intIndex = .Value
				.Col = C_ReasonDesc
				.Value = intIndex
			Case  C_ReasonDesc
				.Col = Col
				intIndex = .Value
				.Col = C_ReasonCd
				.Value = intIndex				
		End Select
		
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

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
			Call LayerShowHide(1)
			If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
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
'********************************************************************************************************
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
        IntRetCD = displaymsgbox("900013", parent.VB_YES_NO, "x", "x")	'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field

    Call InitVariables
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData

	'-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If														'☜: Query db data
	
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
        IntRetCD = displaymsgbox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData3							'⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
       Exit Function
    End If
    If Not chkfield(Document, "2") Then					'⊙: Check required field(Single area)
       Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function						'☜: Save db data
    
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
    frm1.vspdData2.ReDraw = True
    
    'SetSpreadColor frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow
   
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

Dim Row
Dim strMode
Dim	strProdtOrderNo, strOprNo
Dim	strSequence
Dim strChangeFlag
Dim LngRow, LngFindRow

	If frm1.vspdData2.MaxRows < 1 Then Exit Function	

    ggoSpread.Source = frm1.vspdData2	
    Row = frm1.vspdData2.ActiveRow
    frm1.vspdData2.Row = Row
    frm1.vspdData2.Col = 0
    strMode = frm1.vspdData2.Text
   
    frm1.vspdData2.Col = C_ProdtOrderNo1
    strProdtOrderNo = frm1.vspdData2.Text
    frm1.vspdData2.Col = C_OprNo1
    strOprNo = frm1.vspdData2.Text
    frm1.vspdData2.Col = C_Sequence
    strSequence = frm1.vspdData2.Text

	If strMode = ggoSpread.InsertFlag Then
		Call DeleteHSheet(strProdtOrderNo, strOprNo, strSequence)
	    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	End If

	strChangeFlag = "N"
	
	With frm1.vspdData2
	
		For LngRow = 1 to frm1.vspdData2.MaxRows
			frm1.vspdData2.Row = LngRow
			frm1.vspdData2.Col = 0
			strMode = frm1.vspdData2.Text
			If strMode = ggoSpread.InsertFlag Then
				strChangeFlag = "Y"
				Exit For
			End If
		Next
		
	End With

	If strChangeFlag = "N" Then
		LngFindRow = FindRow(strProdtOrderNo, strOprNo)
		If LngFindRow > 0 Then
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SSDeleteFlag LngFindRow,LngFindRow
		End If
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
	Dim strProdtOrderNo
	Dim strLotReq
	Dim strLotGenMthd
	Dim strProdInspReq
	Dim strFinalInspReq
	Dim strAutoRcptFlg
	Dim LngCompSeq, LngRowCnt
	Dim strRoutOrder
	Dim strOprNo
	Dim strInsideFlg
	
	On Error Resume Next
	
	FncInsertRow = False
	
	If SetBlnInsertRow(frm1.vspdData1.ActiveRow) = False Then Exit Function
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If

	With frm1
		.vspdData1.Row = .vspdData1.ActiveRow	    
		' Get Production Order No.
		.vspdData1.Col = C_ProdtOrderNo
		strProdtOrderNo = Trim(.vspdData1.value)
		.vspdData1.Col = C_OprNo
		strOprNo = Trim(.vspdData1.value)		
		.vspdData1.Col = C_LotReq
		strLotReq = Trim(.vspdData1.value)
		.vspdData1.Col = C_LotGenMthd
		strLotGenMthd = Trim(.vspdData1.value)
		.vspdData1.Col = C_ProdInspReq
		strProdInspReq = Trim(.vspdData1.value)
		.vspdData1.Col = C_FinalInspReq
		strFinalInspReq = Trim(.vspdData1.value)
		.vspdData1.Col = C_AutoRcptFlg
		strAutoRcptFlg = Trim(.vspdData1.value)
		.vspdData1.Col = C_RoutOrder
		strRoutOrder = Trim(.vspdData1.value)
		frm1.vspdData1.Col = C_InsideFlag
		strInsideFlag = Trim(.vspdData1.value)
		.vspdData2.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData2
		
		For LngRowCnt = 1 To .vspdData2.MaxRows
			.vspdData2.Row = LngRowCnt
			
		    .vspdData2.Col = C_Sequence
			
			If CInt(LngCompSeq) < CInt(.vspdData2.value) Then
				LngCompSeq = CInt(.vspdData2.value)
			End If
		Next
		
		.vspdData2.ReDraw = False
		ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
    	
    	For pvRow = .vspdData2.ActiveRow To .vspdData2.ActiveRow + imRow -1
			LngCompSeq = LngCompSeq + 1
			
			.vspdData2.Row = pvRow
			.vspdData2.Col = C_ProdtOrderNo1
			.vspdData2.value = strProdtOrderNo
			.vspdData2.Col = C_OprNo1
			.vspdData2.value = strOprNo
			.vspdData2.Col = C_Sequence
			.vspdData2.value = LngCompSeq
			.vspdData2.Col = C_ReportDt
			.vspdData2.Text = .txtReportDt.Text	
			.vspdData2.Col = C_ReportType
			.vspdData2.Text = "G"
			.vspdData2.Col = C_ShiftId
			.vspdData2.Text = lgShift
			.vspdData2.Col = C_ProdQty
			.vspdData2.value = "0"
			.vspdData2.Col = C_LotReq1
			.vspdData2.value = strLotReq
			.vspdData2.Col = C_LotGenMthd1
			.vspdData2.value = strLotGenMthd
			.vspdData2.Col = C_ProdInspReq1
			.vspdData2.value = strProdInspReq
			.vspdData2.Col = C_FinalInspReq1
			.vspdData2.value = strFinalInspReq
			.vspdData2.Col = C_AutoRcptFlg1
			.vspdData2.value = strAutoRcptFlg
			.vspdData2.Col = C_InsideFlag1
			.vspdData2.value = strInsideFlg
			.vspdData2.Col = C_RoutOrder1
			.vspdData2.value = strRoutOrder
'2008-03-19 12:18오전 :: hanc
			.vspdData2.Col = C_LotNo
			.vspdData2.value = "*"

				
		Next
		
		.vspdData2.ReDraw = True
		
		SetSpreadColor frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow + imRow -1, strLotReq, strLotGenMthd, strProdInspReq, strFinalInspReq, strAutoRcptFlg, strInsideFlg, strRoutOrder
				
		Set gActiveElement = document.ActiveElement
		
		If Err.number = 0 Then FncInsertRow = True
		
	End With
    
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
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

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData2							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = displaymsgbox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  ******************************
'	설명 : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
    
    DbQuery = False
    
    Call LayerShowHide(1)

    Err.Clear

    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)
		strVal = strVal & "&txtProdOrdNo=" & Trim(.hProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.Value)
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.Value)
		strVal = strVal & "&txtProdFromDt=" & Trim(.hProdFromDt.Value)
		strVal = strVal & "&txtProdTODt=" & Trim(.hProdTODt.Value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.Value)
		strVal = strVal & "&txtOrderType=" & Trim(.hOrderType.Value)
		strVal = strVal & "&txtrdoflag=" & Trim(.hrdoFlag.Value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		strVal = strVal & "&cboJobCd=" & Trim(.hJobCd.Value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	Else
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)
		strVal = strVal & "&txtProdOrdNo=" & Trim(.txtProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)		
		strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.Value)
		strVal = strVal & "&txtProdFromDt=" & Trim(.txtProdFromDt.Text)
		strVal = strVal & "&txtProdTODt=" & Trim(.txtProdTODt.Text)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.Value)
		strVal = strVal & "&txtOrderType=" & Trim(.cboOrderType.Value)
		If frm1.rdoCompleteFlg1.checked = True Then
			strVal = strVal & "&txtrdoflag=" & "Y"
		Else
			strVal = strVal & "&txtrdoflag=" & "N"
		End If
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)
		strVal = strVal & "&cboJobCd=" & Trim(.cboJobCd.Value)						'☆: 조회 조건 데이타 
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

	Dim strInsideFlag
	
	Call InitShiftCombo()
	Call InitData(LngMaxRow)
	Call SetFieldColor(True)
	
	frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1

	lgOldRow = 1
	
	frm1.vspdData1.Col = C_InsideFlag
	frm1.vspdData1.Row = 1
	strInsideFlag = frm1.vspdData1.value
	frm1.vspdData1.Col = C_MilestoneFlg
	frm1.vspdData1.Row = 1
	
	If frm1.vspdData1.value = "Y" and strInsideFlag = "Y" Then
		Call SetToolBar("11001101000111")										'⊙: 버튼 툴바 제어 
	Else
		Call SetToolBar("11001000000111")										'⊙: 버튼 툴바 제어 
	End If

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	
		If DbDtlQuery(frm1.vspdData1.Row) = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE

End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryNotOk()

	Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	Call SetFieldColor(False)
    
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(Byval LngRow) 

Dim strVal
Dim boolExist
Dim lngRows
Dim strProdtOrderNo
Dim strOprNo

	boolExist = False
    With frm1

	    .vspdData1.Row = LngRow
	    .vspdData1.Col = C_ProdtOrderNo
	    strProdtOrderNo = .vspdData1.Text
	    .vspdData1.Col = C_OprNo
	    strOprNo = .vspdData1.Text
        
        frm1.vspdData2.MaxRows = 0
        
	    If CopyFromHSheet(strProdtOrderNo, strOprNo) = True Then
			Exit Function
        End If

		DbDtlQuery = False   
    
		.vspdData1.Row = .vspdData1.ActiveRow

		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strProdtOrderNo)
			strVal = strVal & "&txtOprNo=" & Trim(strOprNo)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strProdtOrderNo)
			strVal = strVal & "&txtOprNo=" & Trim(strOprNo)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)												'☆: 조회 성공후 실행로직 

	Dim LngRow

    '-----------------------
    'Reset variables area
    '-----------------------
	frm1.vspdData2.ReDraw = Fales

	Call InitData(frm1.vspdData2.MaxRows)
   
    lgIntFlgMode = parent.OPMD_UMODE

	frm1.vspdData2.ReDraw = True

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData(ByVal Row)

Dim strProdtOrderNo, strOprNo, strSequence
Dim strHndProdtOrderNo, strHndOprNo, strHndSequence
Dim lRows

    FindData = 0

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows
        
            .vspdData3.Row = lRows
            .vspdData3.Col = C_ProdtOrderNo2
            strHndProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OprNo2
            strHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_Sequence2
            strHndSequence = .vspdData3.Text

            .vspdData2.Row = Row
            .vspdData2.Col = C_ProdtOrderNo1
            strProdtOrderNo = .vspdData2.Text
            .vspdData2.Col = C_OprNo1
            strOprNo = .vspdData2.Text
            .vspdData2.Col = C_Sequence
            strSequence = .vspdData2.Text
            
            If strHndProdtOrderNo = strProdtOrderNo and strHndOprNo = strOprNo and strHndSequence = strSequence Then
				FindData = lRows
				Exit Function
            End If    
        Next
        
    End With        
    
End Function


'=======================================================================================================
'   Function Name : CopyFromHSheet
'   Function Desc : 
'=======================================================================================================
Function CopyFromHSheet(ByVal strProdtOrderNo, strOprNo)

Dim lngRows
Dim boolExist
Dim iCols
Dim strHdnProdtOrderNo
Dim strHdnOprNo
Dim strStatus
Dim strLotReq
Dim strLotGenMthd
Dim strProdInspReq
Dim strFinalInspReq
Dim strAutoRcptFlg
Dim strInsideFlg
Dim strRoutOrder
Dim iCurColumnPos

	ggoSpread.Source = frm1.vspdData2
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

    boolExist = False
    
    CopyFromHSheet = boolExist
    
    With frm1

        Call SortHSheet()
        '------------------------------------
        ' Find First Row
        '------------------------------------
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_ProdtOrderNo2
            strHdnProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OprNo2
            strHdnOprNo = .vspdData3.Text
			
            If Trim(strProdtOrderNo) = Trim(strHdnProdtOrderNo) and Trim(strOprNo) = Trim(strHdnOprNo) Then
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
                
                .vspdData3.Col = C_ProdtOrderNo2
				strHdnProdtOrderNo = .vspdData3.Text
				.vspdData3.Col = C_OprNo2
				strHdnOprNo = .vspdData3.Text

                If Trim(strProdtOrderNo) = Trim(strHdnProdtOrderNo) and Trim(strOprNo) = Trim(strHdnOprNo) Then
					If Trim(strProdtOrderNo) = Trim(strHdnProdtOrderNo) Then
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
						
						.vspdData3.Col = 0

						If .vspdData3.Text = ggoSpread.InsertFlag Then
						
							.vspdData3.Col = C_AutoRcptFlg2
							strAutoRcptFlg = .vspdData3.Text
							.vspdData3.Col = C_LotReq2
							strLotReq = .vspdData3.Text
							.vspdData3.Col = C_LotGenMthd2
							strLotGenMthd = .vspdData3.Text
							.vspdData3.Col = C_ProdInspReq2
							strProdInspReq = .vspdData3.Text
							.vspdData3.Col = C_FinalInspReq2
							strFinalInspReq = .vspdData3.Text
							.vspdData3.Col = C_InsideFlag2
							strInsideFlg = .vspdData3.Text
							.vspdData3.Col = C_RoutOrder2
							strRoutOrder = .vspdData3.Text
							
							Call SetSpreadColor(.vspdData2.Row, .vspdData2.Row, strLotReq, strLotGenMthd, strProdInspReq, strFinalInspReq, strAutoRcptFlg, strInsideFlg, strRoutOrder)
						Else

						End If

					End If
				Else
					lngRows = .vspdData3.MaxRows + 1
                End If   
                
                lngRows = lngRows + 1
                
            Wend
            frm1.vspdData2.Redraw = True

        End If
            
    End With        
    
    CopyFromHSheet = boolExist
   
End Function

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 
'=======================================================================================================
Sub CopyToHSheet(ByVal Row)
Dim lRow
Dim iCols
Dim LngCurRow
Dim iCurColumnPos

	ggoSpread.Source = frm1.vspdData2
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	With frm1 
        
	    lRow = FindData(Row)

	    If lRow > 0 Then
			LngCurRow = lRow
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            For iCols = 1 To 26 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next
        Else
			.vspdData3.MaxRows = .vspdData3.MaxRows + 1
			LngCurRow = .vspdData3.MaxRows
            .vspdData3.Row = .vspdData3.MaxRows
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
       
       
            For iCols = 1 To 26 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next
        
        End If

		With .vspdData3

			.Redraw = False
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSSetRequired C_ReportDt2,				LngCurRow, LngCurRow
			ggoSpread.SSSetRequired C_ReportType2,				LngCurRow, LngCurRow
			ggoSpread.SSSetRequired C_ShiftId2,					LngCurRow, LngCurRow
			ggoSpread.SSSetRequired C_ProdQty2,					LngCurRow, LngCurRow
			
			.Col = C_ReportType2
			.Row = LngCurRow
			
			If Trim(.Text) = "G" Then
				.Col = C_ReasonCd2
				.Text = ""
				.Col = C_ReasonDesc2
				.Text = ""					

				ggoSpread.SpreadLock C_ReasonCd2, LngCurRow, C_ReasonCd2, LngCurRow
				ggoSpread.SpreadLock C_ReasonDesc2, LngCurRow, C_ReasonDesc2, LngCurRow
				ggoSpread.SSSetProtected C_ReasonCd2, LngCurRow, LngCurRow
				ggoSpread.SSSetProtected C_ReasonDesc2, LngCurRow, LngCurRow
				
				If UCase(Trim(GetSpreadText(frm1.vspdData3,C_AutoRcptFlg2,LngCurRow,"X","X"))) = "Y" _
						And UCase(Trim(GetSpreadText(frm1.vspdData3,C_LotReq2,LngCurRow,"X","X"))) = "Y" _
						And UCase(Trim(GetSpreadText(frm1.vspdData3,C_LotGenMthd2,LngCurRow,"X","X"))) = "M" _
						And (UCase(Trim(GetSpreadText(frm1.vspdData3,C_RoutOrder2,LngCurRow,"X","X"))) = "L" _
						Or UCase(Trim(GetSpreadText(frm1.vspdData3,C_RoutOrder2,LngCurRow,"X","X"))) = "S") Then
					ggoSpread.SpreadUnLock C_LotNo2,LngCurRow,C_LotNo2,LngCurRow
					ggoSpread.SpreadUnLock C_LotSubNo2,LngCurRow,C_LotSubNo2,LngCurRow
					ggoSpread.SSSetRequired C_LotNo2,					LngCurRow, LngCurRow
					ggoSpread.SSSetRequired C_LotSubNo2,					LngCurRow, LngCurRow
				Else
					ggoSpread.SpreadUnLock C_LotNo2,LngCurRow,C_LotNo2,LngCurRow
					ggoSpread.SpreadUnLock C_LotSubNo2,LngCurRow,C_LotSubNo2,LngCurRow	
				End If
			Else
				ggoSpread.SpreadUnLock C_ReasonCd2, LngCurRow, C_ReasonCd2, LngCurRow
				ggoSpread.SpreadUnLock C_ReasonDesc2, LngCurRow, C_ReasonDesc2, LngCurRow
				ggoSpread.SSSetRequired C_ReasonCd2, LngCurRow, LngCurRow
				ggoSpread.SSSetRequired C_ReasonDesc2, LngCurRow, LngCurRow
				ggoSpread.SSSetProtected C_LotNo2, LngCurRow, LngCurRow
				ggoSpread.SSSetProtected C_LotSubNo2, LngCurRow, LngCurRow
			End If
    
		End With

	End With
	
End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal strProdtOrderNo, Byval strOprNo, Byval strSequence)

Dim boolExist
Dim lngRows
Dim StrHndProdtOrderNo, strHndOprNo, strHndSequence
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
	    '------------------------------------
        ' Find First Row
        '------------------------------------
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows

            .vspdData3.Col = C_ProdtOrderNo2
			StrHndProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OprNo2
			strHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_Sequence2
			strHndSequence = .vspdData3.Text

            If strProdtOrderNo = StrHndProdtOrderNo and strHndOprNo = strOprNo and strSequence = strHndSequence Then
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
				.vspdData3.Col = C_ProdtOrderNo2
				StrHndProdtOrderNo = .vspdData3.Text
				.vspdData3.Col = C_OprNo2
				strHndOprNo = .vspdData3.Text
				.vspdData3.Col = C_Sequence2
				strHndSequence = .vspdData3.Text
                
                If (strProdtOrderNo <> StrHndProdtOrderNo) or (strOprNo <> strHndOprNo) or (strSequence <> strHndSequence) Then
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
'======================================================================================================
Function SortHSheet()
    
    With frm1
    
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0 'SS_SORT_BY_ROW
        
        .vspdData3.SortKey(1) = C_ProdtOrderNo2	' Production Order No
        .vspdData3.SortKey(2) = C_OprNo2		' Operation No        
        .vspdData3.SortKey(3) = C_Sequence2		' Sequence
        
        .vspdData3.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(3) = 1 'SS_SORT_ORDER_ASCENDING
        
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25 'SS_ACTION_SORT
        .vspdData3.BlockMode = False
        
    End With        
    
End Function

'=======================================================================================================
'   Function Name : FindRow
'   Function Desc : 
'=======================================================================================================
Function FindRow(Byval strProdtOrderNo, Byval strOprNo)

Dim lRows
Dim CompProdtOrderNo, CompOprNo

    FindRow = 0

    With frm1
        
        For lRows = 1 To .vspdData1.MaxRows
        
            .vspdData1.Row = lRows
            .vspdData1.Col = C_ProdtOrderNo
            CompProdtOrderNo = .vspdData1.Text
            .vspdData1.Col = C_OprNo
            CompOprNo = .vspdData1.Text
            If CompProdtOrderNo = strProdtOrderNo And CompOprNo = strOprNo Then
				FindRow = lRows
				Exit Function
            End If    
        Next
        
    End With
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows
    Dim strVal
   
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size
	

    DbSave = False
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			

	iTmpCUBufferCount = -1
	
	strCUTotalvalLen = 0
	
	With frm1.vspdData3

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0
	
			Select Case .Text
		    
			    Case ggoSpread.InsertFlag
					
					strVal = ""
					'=============>Refer BCP4B1_PCnfmRsltArr
					.Col = C_ProdtOrderNo2	' Production Order No.'0
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_OprNo2	' Operation No.			'1
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_ReportType2	' Report Type   '2
					strVal = strVal & Trim(.Text) & iColSep     
					.Col = C_ProdQty2		' Produced Qty  '3
					strVal = strVal & UNIConvNum(.Text,0) & iColSep
					If UNICDbl(.Value) <= 0 Then
						Call LayerShowHide(0)
						.EditMode = True
						strVal = ""
						Call displaymsgbox("189624", "x", "x", "x")
						Call GetHiddenFocus(IntRows, C_ProdQty2)
						Exit Function
					End If			
					.Col = C_ReportDt2		' Produced Date '4
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					If CompareDateByFormat(.Text, LocSvrDate,"실적일","현재일","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
					  Call LayerShowHide(0)
					   .EditMode = True
					   strVal = ""
					  Exit Function               
					End If	 
					.Col = C_ShiftId2		' Shift			'5
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ReasonCd2		' Reason Code	'6
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_LotNo2							'7
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_LotSubNo2						'8
					strVal = strVal & Trim(.Text) & iColSep					
					'9 item Document number
					If UCase(Trim(GetSpreadText(frm1.vspdData3,C_AutoRcptFlg2,IntRows,"X","X"))) = "Y" _
						And UCase(Trim(GetSpreadText(frm1.vspdData3,C_ReportType2,IntRows,"X","X"))) = "G" _
						And (UCase(Trim(GetSpreadText(frm1.vspdData3,C_RoutOrder2 ,IntRows,"X","X"))) = "S" _
							Or UCase(Trim(GetSpreadText(frm1.vspdData3,C_RoutOrder2 ,IntRows,"X","X"))) = "L") Then
						strVal = strVal & UCase(Trim(frm1.txtRcptNo.value)) & iColSep
					Else
						strVal = strVal & iColSep	
					End If
					
					.Col = C_Remark2						'10
					strVal = strVal & Trim(.Text) & iColSep    

					'prod_base_qty -because "" comes in		'11
					strVal = strVal & UNIConvNum("0",0) & iColSep
					' Subcontract Price	'12
					strVal = strVal  & iColSep  
					' Subcontract Amount	'13
					strVal = strVal & iColSep
					' Currency				'14
					strVal = strVal & iColSep

					strVal = strVal & IntRows & iRowSep		'15
					
					
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

			End Select
	    Next
	    
	End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
    DbSave = True
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()
   
    lgIntPrevKey = 0
    lgLngCurRows = 0

	ggoSpread.source = frm1.vspddata1
    frm1.vspdData1.MaxRows = 0
	ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0
	ggoSpread.source = frm1.vspddata3
    frm1.vspdData3.MaxRows = 0
	
	lgIntFlgMode = parent.OPMD_CMODE
	
	Call RemovedivTextArea
	Call MainQuery
	
End Function

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
'==============================================================================
' Function : GetHiddenFocus
' Description : 에러발생시 Hidden Spread Sheet를 찾아 SheetFocus에 값을 넘겨줌.
'==============================================================================
Function GetHiddenFocus(lRow, lCol)

	Dim lRows1, lRows2						'Quantity of the Hidden Data Keys Referenced by FindData Function
	Dim strHdnOrdNo, strHdnOprNo, strHdnSeqNo			'Variable of Hidden Keys
	Dim strSeqNo					'Variable of Visible Sheet Keys		
	
	If Trim(lCol) = "" Then
		lCol = C_ReportDt					'If Value of Column is not passed, Assign Value of the First Column in Second Spread Sheet
	End If
	'Find Key Datas in Hidden Spread Sheet
	With frm1.vspdData3
		.Row = lRow
		.Col = C_ProdtOrderNo2			
		strHdnOrdNo = Trim(.Text)
		.Col = C_OprNo2			
		strHdnOprNo = Trim(.Text)
		.Col = C_Sequence2				
		strHdnSeqNo = Trim(.Text)
	End With
	'Compare Key Datas to Visible Spread Sheets
	With frm1
		For lRows1 = 1 To .vspdData1.MaxRows
			.vspdData1.Row = lRows1
			.vspdData1.Col = C_ProdtOrderNo			
			If Trim(.vspdData1.Text) = strHdnOrdNo Then
				.vspdData1.Col = C_OprNo			
				If Trim(.vspdData1.Text) = strHdnOprNo Then
					.vspdData1.Col = C_ProdtOrderNo	
					.vspdData1.focus
					.vspdData1.Action = 0
					lgOldRow = lRows1			'※ If this line is omitted, program could not query Data When errors occur
					ggoSpread.Source = .vspdData2
					.vspdData2.MaxRows = 0
					If CopyFromHSheet(strHdnOrdNo, strHdnOprNo) = True Then
					    For lRows2 = 1 To .vspdData2.MaxRows
							.vspdData2.Row = lRows2
							.vspdData2.Col = C_Sequence
							strSeqNo = .vspdData2.Text
							'Find Key Datas in Second Sheet and then Focus the Cell 
							If Trim(strHdnSeqNo) = Trim(strSeqNo) Then
								Call SheetFocus(lRows2, lCol)
								Exit Function
							End If
					    Next
					End If
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
   
	Dim LngRow
	Dim strProdtOrderNo
	Dim strOprNo

    ggoSpread.Source = gActiveSpdSheet
    
    If gActiveSpdSheet.Id = "B" Then
		frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
		frm1.vspdData2.Col = C_ProdtOrderNo1
		strProdtOrderNo = Trim(frm1.vspdData2.Text)
		frm1.vspdData2.Col = C_OprNo1
		strOprNo = Trim(frm1.vspdData2.Text)
	End If
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
	
	If gActiveSpdSheet.Id = "A" Then
		Call ggoSpread.ReOrderingSpreadData()
	ElseIf gActiveSpdSheet.Id = "B" Then
		Call InitSpreadComboBox
		Call InitShiftCombo
		
	    ggoSpread.Source = frm1.vspdData3
        Call ggoSpread.RestoreSpreadInf()

		ggoSpread.Source = frm1.vspdData3
		Call InitSpreadSheet("C")
		Call ggoSpread.ReOrderingSpreadData()
		
		Call CopyFromHsheet(strProdtOrderNo,strOprNo)
				
		Call InitData(1)
	End If
   
End Sub 

'========================================================================================
' Function Name : SetBlnInsertRow
' Function Desc : Boolean Value whether insert row will be enabled, was passed in this function 
'========================================================================================
Function SetBlnInsertRow(ByVal pvRow)
	
	Dim strMileStoneFlg, strInsideFlag
	
	frm1.vspdData1.Row = pvRow
	frm1.vspdData1.Col = C_InsideFlag
	strInsideFlag = Trim(frm1.vspdData1.value)
	frm1.vspdData1.Col = C_MilestoneFlg
	strMileStoneFlg = Trim(frm1.vspdData1.value)
	
	If strMileStoneFlg <> "Y" OR strInsideFlag <> "Y" Then
		SetBlnInsertRow = False
	Else
		SetBlnInsertRow = True	
	End If
	
End Function

'==============================================================================
' Function : SetFieldColor
' Description : 중간 입력 필드의 Color를 맞춤. 
'==============================================================================
Function SetFieldColor(ByVal BlnQueryOk) 

	If BlnQueryOk  = True Then
		Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
		If UCase(Trim(GetSpreadText(frm1.vspdData1,C_AutoRcptFlg,1,"X","X"))) = "Y" Then
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"N")
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"Q")
		End If	
	
		frm1.txtReportDt.text	= LocSvrDate
		frm1.txtRcptNo.value = ""
	Else
		Call ggoOper.LockField(Document, "Q")                                   '⊙: Lock  Suitable  Field
	
		frm1.txtReportDt.text	= ""
		frm1.txtRcptNo.value = ""
	End If
End Function
