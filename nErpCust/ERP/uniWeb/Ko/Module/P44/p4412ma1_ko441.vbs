
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'Grid 1 - Order Header
Const BIZ_PGM_QRY1_ID	= "p4412mb1_ko441.asp"						'��: Head Query �����Ͻ� ���� ASP�� 
'Grid 2 - Production Results
Const BIZ_PGM_QRY2_ID	= "p4412mb2_ko441.asp"						'��: �����Ͻ� ���� ASP�� 
'Post Production Results
Const BIZ_PGM_SAVE_ID	= "p4412mb3_ko441.asp"						
'Shift Header
Const BIZ_PGM_SHIFT		= "p4400mb1.asp"						'��: �����Ͻ� ���� ASP�� 
'Jump (E)Production Order 
Const BIZ_PGM_JUMPREWORKRUN_ID = "p4111ma1"
'Jump (E)Resource Consumption (By Order)
Const BIZ_PGM_JUMPORDRSCCOMPT_ID = "p4712ma1"
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Order Header
Dim C_ProdtOrderNo			
Dim C_OprNo					
Dim C_ItemCd				
Dim C_ItemNm				
Dim C_Spec					
Dim C_ProdtOrderQty			
Dim C_ProdtOrderUnit		
Dim C_RemainQty				
Dim C_ProdQtyIn				
Dim C_ReportTypeIn			
Dim C_ReasonCdIn			
Dim C_ReasonDescIn			
Dim C_Remark				
Dim C_LotNoIn				
Dim C_LotSubNoIn	
Dim C_ProdQtyInOrderUnit	
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
Dim C_MachineCd	
Dim C_MachineNm				
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
Dim C_MachineCd1	
Dim C_MachineNm1	
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
Dim C_MachineCd2	
Dim C_MachineNm2			
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
Dim C_ProdInspReq2			
Dim C_FinalInspReq2			
Dim C_RoutOrder2			

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntFlgMode								'Variable is for Operation Status
Dim lgIntPrevKey
Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgLngCurRows
Dim lgCurrRow
Dim lgShift
Dim lgShiftCnt
'==========================================  1.2.3 Global Variable�� ����  ==================================
'============================================================================================================
'----------------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2
'++++++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""							'initializes Previous Key
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgOldRow = 0
	lgSortKey1   = 1
	lgSortKey2   = 1
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
    frm1.txtProdFromDt.text = StartDate
    frm1.txtProdToDt.text   = EndDate
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(pvSpdNo)

    Call InitSpreadPosVariables(pvSpdNo)
    
    Call AppendNumberPlace("6", "3", "0")
    Call AppendNumberPlace("7", "5", "0")
    
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
	
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20040913", ,Parent.gAllowDragDropSpread
    
			.ReDraw = false
    
			.MaxCols = C_ReworkPrevQty +1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ    
			.MaxRows = 0
    
			Call GetSpreadColumnPos("A")
			      
			ggoSpread.SSSetEdit		C_ProdtOrderNo,			"����������ȣ", 18
			ggoSpread.SSSetEdit		C_OprNo,				"����", 6
			ggoSpread.SSSetEdit		C_ItemCd,				"ǰ��", 18
			ggoSpread.SSSetEdit		C_ItemNm,				"ǰ���", 25
			ggoSpread.SSSetEdit		C_Spec,					"�԰�", 25
			ggoSpread.SSSetFloat	C_ProdtOrderQty,		"��������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ProdtOrderUnit,		"��������", 8,,,3	
			ggoSpread.SSSetFloat	C_RemainQty,			"�ܷ�",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyIn,			"��������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetCombo	C_ReportTypeIn,			"��/��", 6
			ggoSpread.SSSetCombo	C_ReasonCdIn,			"�ҷ��ڵ�", 10
			ggoSpread.SSSetCombo	C_ReasonDescIn,			"�ҷ�����", 20
			ggoSpread.SSSetEdit		C_Remark,				"���", 20,,,120
			ggoSpread.SSSetEdit		C_LotNoIn,				"Lot No.", 20,,,25,2
			ggoSpread.SSSetFloat	C_LotSubNoIn,			"����", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"	
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit,	"��������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit,	"��ǰ����",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInOrderUnit,	"�ҷ�����",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_InspGoodQtyInOrderUnit,"ǰ����ǰ",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_InspBadQtyInOrderUnit,"ǰ���ҷ�",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit,	"�԰�����",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_PlanStartDt,			"����������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_PlanComptDt,			"�ϷΌ����", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ReleaseDt,			"�۾�������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_RealStartDt,			"��������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_RoutNo,				"�����", 10
			ggoSpread.SSSetEdit		C_WcCd,					"�۾���", 10
			ggoSpread.SSSetEdit		C_WcNm,					"�۾����", 20
			ggoSpread.SSSetEdit		C_MachineCd,			"���TYPE", 10,,,120
			ggoSpread.SSSetEdit		C_MachineNm,			"����ڵ�", 15,,,120
			ggoSpread.SSSetEdit		C_JobCd,				"�۾�", 8
			ggoSpread.SSSetEdit		C_JobDesc,				"�۾���", 20
			ggoSpread.SSSetEdit		C_RoutOrder,			"�۾�����", 8
			ggoSpread.SSSetEdit		C_OrderStatus,			"���û���", 10
			ggoSpread.SSSetEdit		C_OrderStatusNm,		"���û���", 10
			ggoSpread.SSSetEdit		C_ProdInspReq,			"�����˻�", 8
			ggoSpread.SSSetEdit		C_MilestoneFlg,			"Milestone", 10
			ggoSpread.SSSetEdit		C_InsideFlag,			"�系/��", 10	
			ggoSpread.SSSetEdit		C_InsideFlagNm,			"�系/��", 10
			ggoSpread.SSSetEdit		C_TrackingNo,			"Tracking No.", 25,,,25
			ggoSpread.SSSetEdit		C_ProdtOrderType,		"���ñ���", 10
			ggoSpread.SSSetEdit		C_AutoRcptFlg, "", 10
			ggoSpread.SSSetEdit		C_LotReq, "", 10
			ggoSpread.SSSetEdit		C_LotGenMthd, "", 10
			ggoSpread.SSSetEdit		C_FinalInspReq, "", 10
			ggoSpread.SSSetEdit 	C_ItemGroupCd, "ǰ��׷�",	15
			ggoSpread.SSSetEdit		C_ItemGroupNm, "ǰ��׷��", 30
			ggoSpread.SSSetEdit		C_ParentOrderNo,		"����������ȣ", 18
			ggoSpread.SSSetEdit		C_ParentOprNo,			"��������", 8
			ggoSpread.SSSetEdit		C_OrginalOrderNo,		"����������ȣ", 18
			ggoSpread.SSSetEdit		C_OrginalOprNo,			"��������", 8
			ggoSpread.SSSetFloat	C_ReworkPrevQty,		"���۾�����",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
 
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_OrderStatus, C_OrderStatus, True)
			Call ggoSpread.SSSetColHidden(C_InsideFlag, C_InsideFlag, True)
			Call ggoSpread.SSSetColHidden(C_RoutOrder, C_RoutOrder, True)
			Call ggoSpread.SSSetColHidden(C_AutoRcptFlg, C_AutoRcptFlg, True)
			Call ggoSpread.SSSetColHidden(C_LotReq, C_LotGenMthd, True)
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
			ggoSpread.Spreadinit "V20040120", ,Parent.gAllowDragDropSpread
	
			.ReDraw = false
	
			.MaxCols = C_RoutOrder1 +1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
	
			Call GetSpreadColumnPos("B")
				    
			ggoSpread.SSSetDate 	C_ReportDt,			"������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetCombo	C_ReportType,		"��/��", 6
			ggoSpread.SSSetEdit		C_ShiftId,			"Shift", 8
			ggoSpread.SSSetFloat	C_ProdQty,			"��������",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo	C_ReasonCd,			"�ҷ��ڵ�", 10
			ggoSpread.SSSetCombo	C_ReasonDesc,		"�ҷ�����", 20
			ggoSpread.SSSetEdit		C_Remark1,			"���", 20,,,120
			ggoSpread.SSSetEdit		C_MachineCd1,			"���TYPE", 10,,,120
			ggoSpread.SSSetEdit		C_MachineNm1,			"����ڵ�", 15,,,120

			ggoSpread.SSSetEdit		C_LotNo,			"Lot No.", 20,,,25,2
			ggoSpread.SSSetFloat	C_LotSubNo,			"����", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_RcptDocumentNo,	"�԰���ȣ", 18,,,16,2
			ggoSpread.SSSetEdit		C_IssueDocumentNo,	"�����ȣ", 18,,,16,2	
			ggoSpread.SSSetEdit		C_InspReqNo,		"�˻��Ƿڹ�ȣ", 18,,,18,2	
			ggoSpread.SSSetFloat	C_Insp_Good_Qty1,	"ǰ����ǰ",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Insp_Bad_Qty1,	"ǰ���ҷ�",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Rcpt_Qty1,		"�԰�����",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ProdtOrderNo1,	"",18
			ggoSpread.SSSetEdit		C_OprNo1,			"",10
			ggoSpread.SSSetFloat	C_Sequence,			"����", 8, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_MilestoneFlg1,	"Milestone",10
			ggoSpread.SSSetEdit		C_InsideFlag1,		"�系/��", 10
			ggoSpread.SSSetEdit		C_AutoRcptFlg1,		"",10
			ggoSpread.SSSetEdit		C_LotReq1,			"",10
			ggoSpread.SSSetEdit		C_ProdInspReq1,		"",10
			ggoSpread.SSSetEdit		C_FinalInspReq1,	"",10
			ggoSpread.SSSetEdit		C_RoutOrder1,		"",10
			   
				
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_ProdtOrderNo1, C_ProdtOrderNo1, True)
			Call ggoSpread.SSSetColHidden(C_Sequence, C_Sequence, True)
			Call ggoSpread.SSSetColHidden(C_OprNo1, C_OprNo1, True)
			Call ggoSpread.SSSetColHidden(C_MilestoneFlg1, C_MilestoneFlg1, True)
			Call ggoSpread.SSSetColHidden(C_InsideFlag1, C_InsideFlag1, True)
			Call ggoSpread.SSSetColHidden(C_AutoRcptFlg1, C_AutoRcptFlg1, True)
			Call ggoSpread.SSSetColHidden(C_LotReq1, C_LotReq1, True)
			Call ggoSpread.SSSetColHidden(C_Rcpt_Qty1, C_Rcpt_Qty1, True)  ' hidden for rcpt_qty 20030403 kjp
			Call ggoSpread.SSSetColHidden(C_ProdInspReq1, C_ProdInspReq1, True)
			Call ggoSpread.SSSetColHidden(C_FinalInspReq1, C_FinalInspReq1, True)
			Call ggoSpread.SSSetColHidden(C_RoutOrder1, C_RoutOrder1, True)
	
			ggoSpread.SSSetSplit2(5)
	
			Call SetSpreadLock("B")
	
			.ReDraw = true
    
		End With
    End If
	
	
	If pvSpdNo = "*" Then
		With frm1.vspdData3
			'------------------------------------------
			' Grid 3 - Hidden Spread Setting
			'------------------------------------------
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit
	
			.ReDraw = false
				
			.MaxCols = C_RoutOrder2 +1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
    
			ggoSpread.SSSetDate 	C_ReportDt2,			"������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_ReportType2,			"��/��", 6
			ggoSpread.SSSetEdit		C_ShiftId2,				"Shift", 8	
			ggoSpread.SSSetFloat	C_ProdQty2,				"��������",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ReasonCd2,			"�ҷ��ڵ�", 10
			ggoSpread.SSSetEdit		C_ReasonDesc2,			"�ҷ�����", 15
			ggoSpread.SSSetEdit		C_Remark2,				"���", 20,,,120
			ggoSpread.SSSetEdit		C_MachineCd2,			"���TYPE", 10,,,120
			ggoSpread.SSSetEdit		C_MachineNm2,			"����ڵ�", 15,,,120
			ggoSpread.SSSetEdit		C_LotNo2,				"Lot No.", 20,,,25,2		
			ggoSpread.SSSetFloat	C_LotSubNo2,			"����", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_RcptDocumentNo2,		"�԰���ȣ", 18,,,16
			ggoSpread.SSSetEdit		C_IssueDocumentNo2,		"�����ȣ", 18,,,16	
			ggoSpread.SSSetEdit		C_InspReqNo2,			"�˻��Ƿڹ�ȣ", 18,,,18	
			ggoSpread.SSSetEdit		C_ProdtOrderNo2,		"������ȣ", 18
			ggoSpread.SSSetFloat	C_Sequence2,			"����", 8, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_InsideFlag2,			"�系/��", 10	
			ggoSpread.SSSetEdit		C_MilestoneFlg2,		"Milestone", 10	
			ggoSpread.SSSetFloat	C_Insp_Good_Qty2,		"ǰ����ǰ",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Insp_Bad_Qty2,		"ǰ���ҷ�",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Rcpt_Qty2,			"�԰�����",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
			Call SetSpreadLock("C")
			
			.ReDraw = true
    
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
			.vspdData1.ReDraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData1.ReDraw = True
		
		ElseIf pvSpdNo = "B" Then    
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		
		ElseIf pvSpdNo = "C" Then	
			'--------------------------------
			'Grid 3
			'--------------------------------
			ggoSpread.Source = frm1.vspdData3
			.vspdData3.Redraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData3.Redraw = True
	   End If
	End With

End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub


'========================== 2.2.6 InitSpreadComboBox()  ========================================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox(ByVal pvSpdNo)
	
    Dim strCboCd
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	'****************************
	'List Minor code(G/B & Reason)
	'****************************
	strCboCd =  "G" & vbTab & "B"
	
	'****************************
	'List Minor code(Reason Code)
	'****************************
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3221", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.SetCombo strCboCd, C_ReportTypeIn
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ReasonCdIn
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ReasonDescIn
	
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SetCombo strCboCd, C_ReportType
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ReasonCd
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ReasonDesc
		
	End If
	
End Sub

'==========================================  2.2.6 InitShiftCombo()  =======================================
'	Name : InitShiftCombo()
'	Description : Combo Display
'===========================================================================================================
Function InitShiftCombo()

    Dim strVal
    Dim i
	
	InitShiftCombo = False
	
    With frm1
	If .txtPlantCd.value = "" Then Exit	Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
'		frm1.txtPlantNm.value = ""
	Else
		For i = lgShiftCnt To 1 Step -1
			frm1.cboShift.remove(i) 
		Next
	End If
	
	strVal = BIZ_PGM_SHIFT & "?txtMode=" & parent.UID_M0001						'��: 
	strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
       
    End With
	
	InitShiftCombo = True
			
End Function


'==========================================  2.2.6 InitShiftComboOk()  =======================================
'	Name : InitShiftComboOk()
'	Description : Query
'===========================================================================================================
Sub InitShiftComboOk()
	frm1.cboShift.value = lgShift
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Sub
	End If												'��: Query db data

End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
 Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex
	
	With frm1.vspdData1
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_ReasonCdIn
			intIndex = .value
			.Col = C_ReasonDescIn
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
		' Grid 1(vspdData1) - Production Order
		C_ProdtOrderNo				= 1
		C_OprNo						= 2
		C_ItemCd					= 3
		C_ItemNm					= 4
		C_Spec						= 5
		C_ProdtOrderQty				= 6
		C_ProdtOrderUnit			= 7
		C_RemainQty					= 8
		C_ProdQtyIn					= 9
		C_ReportTypeIn				= 10
		C_ReasonCdIn				= 11
		C_ReasonDescIn				= 12
		C_Remark					= 13
		C_LotNoIn					= 14
		C_LotSubNoIn				= 15
		C_ProdQtyInOrderUnit		= 16
		C_GoodQtyInOrderUnit		= 17
		C_BadQtyInOrderUnit			= 18
		C_InspGoodQtyInOrderUnit	= 19
		C_InspBadQtyInOrderUnit		= 20
		C_RcptQtyInOrderUnit		= 21
		C_PlanStartDt				= 22
		C_PlanComptDt				= 23
		C_ReleaseDt					= 24
		C_RealStartDt				= 25
		C_RoutNo					= 26
		C_WcCd						= 27
		C_WcNm						= 28
		C_MachineCd			    	= 29    '2008-03-25 2:33���� :: hanc
		C_MachineNm 				= 30    '2008-03-25 2:34���� :: hanc
		C_JobCd						= 31
		C_JobDesc					= 32
		C_RoutOrder					= 33
		C_OrderStatus				= 34
		C_OrderStatusNm				= 35
		C_MilestoneFlg				= 36
		C_InsideFlag				= 37
		C_InsideFlagNm				= 38
		C_TrackingNo				= 39
		C_ProdtOrderType			= 40
		C_AutoRcptFlg				= 41
		C_LotReq					= 42
		C_LotGenMthd				= 43
		C_ProdInspReq				= 44
		C_FinalInspReq				= 45
		C_ItemGroupCd				= 46
		C_ItemGroupNm				= 47
		C_ParentOrderNo				= 48
		C_ParentOprNo				= 49
		C_OrginalOrderNo			= 50
		C_OrginalOprNo				= 51
		C_ReworkPrevQty				= 52
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
		C_MachineCd1				= 8     '2008-03-25 2:33���� :: hanc
		C_MachineNm1				= 9     '2008-03-25 2:34���� :: hanc
		C_LotNo						= 10
		C_LotSubNo					= 11
		C_RcptDocumentNo			= 12
		C_IssueDocumentNo			= 13
		C_InspReqNo					= 14
		C_Insp_Good_Qty1			= 15
		C_Insp_Bad_Qty1				= 16
		C_Rcpt_Qty1					= 17
		' Hidden
		C_ProdtOrderNo1				= 18
		C_OprNo1					= 19
		C_Sequence					= 20
		C_MilestoneFlg1				= 21
		C_InsideFlag1				= 22
		C_AutoRcptFlg1				= 23
		C_LotReq1					= 24
		C_ProdInspReq1				= 25
		C_FinalInspReq1				= 26
		C_RoutOrder1				= 27
	End If
	
	If pvSpdNo = "*" Then
		' Grid 3(vspdData3) - Hidden
		C_ReportDt2					= 1
		C_ReportType2				= 2
		C_ShiftId2					= 3
		C_ProdQty2					= 4
		C_ReasonCd2					= 5
		C_ReasonDesc2				= 6
		C_Remark2					= 7
		C_MachineCd2				= 8     '2008-03-25 2:33���� :: hanc
		C_MachineNm2				= 9     '2008-03-25 2:34���� :: hanc
		C_LotNo2					= 10
		C_LotSubNo2					= 11
		C_RcptDocumentNo2			= 12
		C_IssueDocumentNo2			= 13
		C_InspReqNo2				= 14
		C_Insp_Good_Qty2			= 15
		C_Insp_Bad_Qty2				= 16
		C_Rcpt_Qty2					= 17
		C_ProdtOrderNo2				= 18
		C_OprNo2					= 19
		C_Sequence2					= 20
		C_MilestoneFlg2				= 21
		C_InsideFlag2				= 22
		C_AutoRcptFlg2				= 23
		C_LotReq2					= 24
		C_ProdInspReq2				= 25
		C_FinalInspReq2				= 26
		C_RoutOrder2				= 27
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
			C_OprNo						= iCurColumnPos(2)
			C_ItemCd					= iCurColumnPos(3)
			C_ItemNm					= iCurColumnPos(4)
			C_Spec						= iCurColumnPos(5)
			C_ProdtOrderQty				= iCurColumnPos(6)
			C_ProdtOrderUnit			= iCurColumnPos(7)
			C_RemainQty					= iCurColumnPos(8)
			C_ProdQtyIn					= iCurColumnPos(9)
			C_ReportTypeIn				= iCurColumnPos(10)
			C_ReasonCdIn				= iCurColumnPos(11)
			C_ReasonDescIn				= iCurColumnPos(12)
			C_Remark					= iCurColumnPos(13)
			C_LotNoIn					= iCurColumnPos(14)
			C_LotSubNoIn				= iCurColumnPos(15)
			C_ProdQtyInOrderUnit		= iCurColumnPos(16)
			C_GoodQtyInOrderUnit		= iCurColumnPos(17)
			C_BadQtyInOrderUnit			= iCurColumnPos(18)
			C_InspGoodQtyInOrderUnit	= iCurColumnPos(19)
			C_InspBadQtyInOrderUnit		= iCurColumnPos(20)
			C_RcptQtyInOrderUnit		= iCurColumnPos(21)
			C_PlanStartDt				= iCurColumnPos(22)
			C_PlanComptDt				= iCurColumnPos(23)
			C_ReleaseDt					= iCurColumnPos(24)
			C_RealStartDt				= iCurColumnPos(25)
			C_RoutNo					= iCurColumnPos(26)
			C_WcCd						= iCurColumnPos(27)
			C_WcNm						= iCurColumnPos(28)
		    C_MachineCd 				= iCurColumnPos(29) '2008-03-25 2:35���� :: hanc
			C_MachineNm	    			= iCurColumnPos(30) '2008-03-25 2:35���� :: hanc
			C_JobCd						= iCurColumnPos(31)
			C_JobDesc					= iCurColumnPos(32)
			C_RoutOrder					= iCurColumnPos(33)
			C_OrderStatus				= iCurColumnPos(34)
			C_OrderStatusNm				= iCurColumnPos(35)
			C_MilestoneFlg				= iCurColumnPos(36)
			C_InsideFlag				= iCurColumnPos(37)
			C_InsideFlagNm				= iCurColumnPos(38)
			C_TrackingNo				= iCurColumnPos(39)
			C_ProdtOrderType			= iCurColumnPos(40)
			C_AutoRcptFlg				= iCurColumnPos(41)
			C_LotReq					= iCurColumnPos(42)
			C_LotGenMthd				= iCurColumnPos(43)
			C_ProdInspReq				= iCurColumnPos(44)
			C_FinalInspReq				= iCurColumnPos(45)
			C_ItemGroupCd				= iCurColumnPos(46)
			C_ItemGroupNm				= iCurColumnPos(47)
			C_ParentOrderNo				= iCurColumnPos(48)
			C_ParentOprNo				= iCurColumnPos(49)
			C_OrginalOrderNo			= iCurColumnPos(50)
			C_OrginalOprNo				= iCurColumnPos(51)
			C_ReworkPrevQty				= iCurColumnPos(52)
		
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
			 C_MachineCd1				= iCurColumnPos(8)  '2008-03-25 2:35���� :: hanc
			 C_MachineNm1				= iCurColumnPos(9)  '2008-03-25 2:35���� :: hanc
			 C_LotNo					= iCurColumnPos(10)
			 C_LotSubNo					= iCurColumnPos(11)
			 C_RcptDocumentNo			= iCurColumnPos(12)
			 C_IssueDocumentNo			= iCurColumnPos(13)
			 C_InspReqNo				= iCurColumnPos(14)
			 C_Insp_Good_Qty1			= iCurColumnPos(15)
			 C_Insp_Bad_Qty1			= iCurColumnPos(16)
			 C_Rcpt_Qty1				= iCurColumnPos(17)
			' Hidden
			 C_ProdtOrderNo1			= iCurColumnPos(18)
			 C_OprNo1					= iCurColumnPos(19)
			 C_Sequence					= iCurColumnPos(20)
			 C_MilestoneFlg1			= iCurColumnPos(21)
			 C_InsideFlag1				= iCurColumnPos(22)
			 C_AutoRcptFlg1				= iCurColumnPos(23)
			 C_LotReq1					= iCurColumnPos(24)
			 C_ProdInspReq1				= iCurColumnPos(25)
			 C_FinalInspReq1			= iCurColumnPos(26)
			 C_RoutOrder1				= iCurColumnPos(27)

    End Select    
End Sub    

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
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

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
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
		Call displaymsgbox("971012","X", "����","X")
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
		Call displaymsgbox("971012","X", "����","X")
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
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field��(0)
	arrField(1) = 2 '"ITEM_NM"					' Field��(1)
    
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

	arrParam(0) = "ǰ��׷��˾�"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "ǰ��׷�"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "ǰ��׷�"
	arrHeader(1) = "ǰ��׷��"
	    
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
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")		' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
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
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
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
		Call displaymsgbox("971012","X", "����","X")
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
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
		frm1.vspdData1.Col = C_OprNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
	End If	
		
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
		Call displaymsgbox("971012","X", "����","X")
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
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
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
		Call displaymsgbox("971012","X", "����","X")
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
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
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
		Call displaymsgbox("971012","X", "����","X")
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
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenReworkRef()  --------------------------------------------
'	Name : OpenReworkRef()
'	Description : Rework Order History Reference
'---------------------------------------------------------------------------------------------------------
Function OpenReworkRef()

	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
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
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
		'opr_no
		arrParam(3) = ""
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3)), _
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
		Call displaymsgbox("971012","X", "����","X")
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
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
		frm1.vspdData1.Col = C_OprNo
		arrParam(3) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
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
	
	If IsOpenPop = True Then Exit Function
	
	strVal = ""
	
	With frm1.vspdData1
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = C_ProdQtyIn		' Produced Qty
			If UNICDbl(.Text) > CDbl(0) Then

				strVal = strVal & frm1.hPlantCd.value & parent.gColSep
				.Col = C_ProdtOrderNo			
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_OprNo
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_ProdQtyIn
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

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)

    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With

End Function

'------------------------------------------  SetItemInfo()  -------------------------------------------
'	Name : SetItemGroup()
'	Description : Item Group Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetTrackingNo()  ----------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	
	frm1.txtTrackingNo.Value = arrRet(0)
	
End Function


'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtProdFromDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdFromDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtProdToDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdToDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	


'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportDT_DblClick(Button)
    If Button = 1 Then
        frm1.txtReportDT.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReportDT.Focus
    End If
End Sub

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
		Call DisplayMsgBox("971012","X", "����","X")
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
		
	WriteCookie"txtPlantCd", UCase(Trim(frm1.hPlantCd.value))
	WriteCookie"txtPlantNm", UCase(Trim(frm1.txtPlantNm.value))
	WriteCookie"txtItemCd", strItemCd
	WriteCookie"txtProdOrderNo", strProdtOrdNo
	WriteCookie"txtOprNo", strOprNo
	WriteCookie"txtJumpQty", UniFormatNumber(DblJumpQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	WriteCookie"txtTrackingNo", strTrackingNo
	
	PgmJump(BIZ_PGM_JUMPREWORKRUN_ID)
	
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
		Call DisplayMsgBox("971012","X", "����","X")
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

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  *********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� �����Ͽ� �ۼ��Ѵ�.
'*******************************************************************************************************

'******************************  3.2.1 Object Tag ó��  ************************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*******************************************************************************************************
Sub vspdData1_Click(ByVal Col , ByVal Row )
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
  	Else
  		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
  	End If
    
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
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
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
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData1_Change(ByVal Col, ByVal Row)

	Dim strInsideFlag
	Dim strMilestoneFlag	
	Dim strLotReq
	Dim strProdInspReq
	Dim strFinalInspReq
	Dim strAutoRcptFlg
	Dim	strRoutOrder
	Dim strLotGenMthd
	Dim dblProdQty, dblSubcontractPrice
	Dim dblProdtOrderQty, dblOrderQtyInBaseUnit
	
	With frm1.vspdData1
		
		Select Case Col

		    Case C_ProdQtyIn

				.Row = Row
				.Col = C_InsideFlag
				strInsideFlag = .value
				.Col = C_MilestoneFlg
				strMilestoneFlag = .value
				.Col = C_LotReq
				strLotReq = .value
				.Col = C_ProdInspReq
				strProdInspReq = .value
				.Col = C_FinalInspReq
				strFinalInspReq = .value
				.Col = C_AutoRcptFlg
				strAutoRcptFlg = .value
				.Col = C_RoutOrder
				strRoutOrder = .value
				.Col = C_LotGenMthd
				strLotGenMthd = .value
				.Col = C_ProdQtyIn

				If strMilestoneFlag = "Y" and strInsideFlag = "Y" and UNICDbl(.value) > 0 Then
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.SpreadUnLock C_ReportTypeIn, Row, C_ReportTypeIn, Row
					ggoSpread.SSSetRequired C_ProdQtyIn,Row,Row
					ggoSpread.SSSetRequired C_ReportTypeIn,Row,Row
					.Col = C_ReportTypeIn
					If Trim(.Text) = "G" Then
						ggoSpread.SpreadLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
						ggoSpread.SpreadLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
						ggoSpread.SSSetProtected C_ReasonCdIn, Row, Row
						ggoSpread.SSSetProtected C_ReasonDescIn, Row, Row
						.Col = C_ReasonCdIn
						.Text = ""
						.Col = C_ReasonDescIn
						.Text = ""
						
						If strLotReq <> "Y" or strAutoRcptFlg <> "Y" or (strRoutOrder <> "S" and strRoutOrder <> "L")  Then
							ggoSpread.SSSetProtected C_LotNoIn, Row, Row
							ggoSpread.SSSetProtected C_LotSubNoIn, Row, Row
						Else
							If strLotGenMthd = "M" Then
								ggoSpread.SpreadUnLock C_LotNoIn,Row,C_LotNoIn,Row
								ggoSpread.SpreadUnLock C_LotSubNoIn,Row,C_LotSubNoIn,Row							
								ggoSpread.SSSetRequired C_LotNoIn, Row, Row
								ggoSpread.SSSetRequired C_LotSubNoIn, Row, Row
							Else
								ggoSpread.SpreadUnLock C_LotNoIn,Row,C_LotNoIn,Row
								ggoSpread.SpreadUnLock C_LotSubNoIn,Row,C_LotSubNoIn,Row
							End If	
							
						End If
						
					Else
						ggoSpread.SpreadUnLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
						ggoSpread.SpreadUnLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
						ggoSpread.SSSetRequired C_ReasonCdIn, Row, Row
						ggoSpread.SSSetRequired C_ReasonDescIn, Row, Row
					End If
					ggoSpread.SpreadUnLock C_Remark, Row, C_Remark, Row
			
					ggoSpread.UpdateRow Row
				Else
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.SpreadUnLock C_ProdQtyIn,Row,C_ProdQtyIn,Row
					ggoSpread.SpreadLock C_ReportTypeIn,Row,C_ReportTypeIn,Row
					ggoSpread.SpreadLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
					ggoSpread.SpreadLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
					ggoSpread.SpreadLock C_Remark, Row, C_Remark, Row
					ggoSpread.SSSetProtected C_ReasonCdIn, Row, Row
					ggoSpread.SSSetProtected C_ReasonDescIn, Row, Row
					ggoSpread.SSSetProtected C_Remark, Row, Row

					.Col = C_ReportTypeIn
					.Text = "G"
					.Col = C_ReasonCdIn
					.Text = ""
					.Col = C_ReasonDescIn
					.Text = ""
					.Col = C_LotNoIn
					.Text = ""
					.Col = C_LotSubNoIn
					.Text = ""
					.Col = C_Remark
					.Text = ""
					
	   				ggoSpread.SSSetProtected C_LotNoIn, Row, Row
	   				ggoSpread.SSSetProtected C_LotSubNoIn, Row, Row
					
					ggoSpread.SSDeleteFlag Row,Row
					
				End If
	    	
		End Select

	End With

End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim strLotReq
	Dim strAutoRcptFlg
	Dim strRoutOrder
	Dim strLotGenMthd

	ggoSpread.Source = frm1.vspdData1

	With frm1.vspdData1

		.Row = Row
		Select Case Col
		
		    Case C_ReportTypeIn
		       	
				.Col = Col
				.Row = Row

				.Col = C_LotReq
				strLotReq = .Text
				.Col = C_AutoRcptFlg
				strAutoRcptFlg  = .Text
				.Col = C_RoutOrder
				strRoutOrder = .Text
				.Col = C_LotGenMthd
				strLotGenMthd = .Text
				.ReDraw = False

				.Col = C_ReportTypeIn
				If Trim(.Text) = "G" Then
				
					ggoSpread.SpreadLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
					ggoSpread.SpreadLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
					ggoSpread.SSSetProtected C_ReasonCdIn, Row, Row
					ggoSpread.SSSetProtected C_ReasonDescIn, Row, Row

					.Col = C_ReasonCdIn
					.Text = ""
					.Col = C_ReasonDescIn
					.Text = ""

					If strLotReq <> "Y" or strAutoRcptFlg <> "Y" or (strRoutOrder <> "S" and strRoutOrder <> "L")  Then
						ggoSpread.SSSetProtected C_LotNoIn, Row, Row
						ggoSpread.SSSetProtected C_LotSubNoIn, Row, Row
					Else
						If Trim(strLotGenMthd) = "M" Then
							ggoSpread.SpreadUnLock C_LotNoIn, Row, C_LotNoIn, Row
							ggoSpread.SpreadUnLock C_LotSubNoIn, Row, C_LotSubNoIn, Row
							ggoSpread.SSSetRequired C_LotNoIn, Row, Row
							ggoSpread.SSSetRequired C_LotSubNoIn, Row, Row
						Else
							ggoSpread.SpreadUnLock C_LotNoIn, Row, C_LotNoIn, Row
							ggoSpread.SpreadUnLock C_LotSubNoIn, Row, C_LotSubNoIn, Row
						End If
						
					End If
					
				Else
					
					.Col = C_LotNoIn
					.Text = ""
					.Col = C_LotSubNoIn
					.Text = ""
					
					ggoSpread.SpreadUnLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
					ggoSpread.SpreadUnLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
					ggoSpread.SSSetRequired C_ReasonCdIn, Row, Row
					ggoSpread.SSSetRequired C_ReasonDescIn, Row, Row
					
					ggoSpread.SSSetProtected C_LotNoIn, Row, Row
					ggoSpread.SSSetProtected C_LotSubNoIn, Row, Row
				End If
			
				.ReDraw = True
		
			Case  C_ReasonCdIn
				.Col = Col
				intIndex = .Value
				.Col = C_ReasonDescIn
				.Value = intIndex
			Case  C_ReasonDescIn
				.Col = Col
				intIndex = .Value
				.Col = C_ReasonCdIn
				.Value = intIndex				
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
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgIntPrevKey <> 0 Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call LayerShowHide(1)
			If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'========================================================================================
' Function Name : vspdData1_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
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
' Function Desc : �׸��� ��ġ ���� 
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
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

	Dim LngRow
	Dim strInsideFlag
	Dim strMilestoneFlag

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    Call InitSpreadComboBox(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
	
	If gActiveSpdSheet.Id = "A" Then
		Call InitData(1)
		ggoSpread.Source = frm1.vspdData1
		with frm1.vspdData1
			For LngRow = 1 To .MaxRows
				
				.Row = LngRow
				.Col = C_InsideFlag
				strInsideFlag = .value
				.Col = C_MilestoneFlg
				strMilestoneFlag = .value
				
				If Trim(strMilestoneFlag) = "Y" and Trim(strInsideFlag) = "Y" Then
					ggoSpread.SpreadUnLock C_ProdQtyIn,LngRow,C_ProdQtyIn,LngRow
				Else
					ggoSpread.SpreadLock C_ProdQty,LngRow,C_ProdQty,LngRow								
				End If
			Next
		end with
	Else
		lgOldRow = 0
		Call vspdData1_Click(frm1.vspdData1.ActiveCol, frm1.vspdData1.ActiveRow)	
    
    End If
    
End Sub 


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'#########################################################################################################


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    Dim IntRetCD 
    
    FncQuery = False											'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = displaymsgbox("900013", parent.VB_YES_NO, "x", "x")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function

   '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "3")						'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData

    Call InitVariables

	ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
	If InitShiftCombo = False Then Exit Function
	
    FncQuery = True												'��: Processing is OK
   
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
    
    FncSave = False												'��: Processing is NG
    
    Err.Clear                                                   '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = displaymsgbox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
    
    If Not chkfield(Document, "2") Then					'��: Check required field(Single area)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function						'��: Save db data
    
    FncSave = True												'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	Dim strMode
 
    If frm1.vspdData1.MaxRows < 1 Then Exit Function
    If gActiveSpdSheet.Id = "A" Then
		With frm1.vspdData1
			ggoSpread.Source = frm1.vspdData1
			.Row = .ActiveRow
			.Col = 0
			strMode = .Text
			If strMode = ggoSpread.UpdateFlag Then
				ggoSpread.EditUndo                                    '��: Protect system from crashing
				Call vspdData1_Change(C_ProdQtyIn, .ActiveRow)
			End If	
		End With
	End If	
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
	
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = displaymsgbox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc�Լ������� ȣ��Ǵ� ���� Function  ******************************
'	���� : 
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
		strVal = strVal & "&cboJobCd=" & Trim(.cboJobCd.Value)						'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)

    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Dim	lRow
	Dim strInsideFlag
	Dim strMilestoneFlag
	
	Call InitData(LngMaxRow)
	
	Call SetFieldColor(True)
	
	frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1

	lgOldRow = 1

    With frm1.vspdData1

		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData1
		
		For lRow = LngMaxRow To .MaxRows

			ggoSpread.Source = frm1.vspdData1
			
			.Row = lRow
			.Col = C_InsideFlag
			strInsideFlag = .value
			.Col = C_MilestoneFlg
			strMilestoneFlag = .value
			
			If Trim(strMilestoneFlag) = "Y" and Trim(strInsideFlag) = "Y" Then
				ggoSpread.SpreadUnLock C_ProdQtyIn,lRow,C_ProdQtyIn,lRow
			Else
				ggoSpread.SpreadLock C_ProdQtyIn,lRow,C_ProdQtyIn,lRow								
			End If

		Next
		
		.Redraw = True
    
    End With

	Call SetToolBar("11001001000111")										'��: ��ư ���� ���� 

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
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryNotOk()

	Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
	Call SetFieldColor(False)
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByVal LngRow) 

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
    
		.vspdData1.Row = LngRow

		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strProdtOrderNo)
			strVal = strVal & "&txtOprNo=" & Trim(strOprNo)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strProdtOrderNo)
			strVal = strVal & "&txtOprNo=" & Trim(strOprNo)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    End With

    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 

	Dim LngRow

    '-----------------------
    'Reset variables area
    '-----------------------
	frm1.vspdData2.ReDraw = False

	Call InitData(frm1.vspdData2.MaxRows)
   
    lgIntFlgMode = parent.OPMD_UMODE

	frm1.vspdData2.ReDraw = True

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData()

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

            .vspdData2.Row = frm1.vspdData2.ActiveRow
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
Dim strProdInspReq
Dim strFinalInspReq
Dim strAutoRcptFlg
Dim strInsideFlg
Dim strRoutOrder
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
        
	    lRow = FindData

	    If lRow > 0 Then
			LngCurRow = lRow
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            For iCols = 1 To 25 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCols
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
       
       
            For iCols = 1 To 25 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next
        
        End If

	End With
	
End Sub


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

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows 
    Dim strVal  
	
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

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
	
	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)						

	iTmpCUBufferCount = -1
	
	strCUTotalvalLen = 0
	
	With frm1.vspdData1

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0

			Select Case .Text
		    
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					
					strVal = ""
					
					.Col = C_ProdtOrderNo	' Production Order No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_OprNo	' Operation No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_ReportTypeIn	' Report Type
					strVal = strVal & Trim(.Text) & iColSep     '5
					.Col = C_ProdQtyIn		' Produced Qty
					strVal = strVal & UNIConvNum(.Text,0) & iColSep
					' Produced Date
					strVal = strVal & UNIConvDate(frm1.txtReportDT.Text) & iColSep
					If CompareDateByFormat(frm1.txtReportDT.Text, LocSvrDate,"������","������","970025",parent.gDateFormat,parent.gComDateType,True) = False Then	
						Call LayerShowHide(0)
					   .EditMode = True
					  Exit Function               
					End If 
					' Shift
					strVal = strVal & UCase(Trim(frm1.cboShift.value)) & iColSep
					.Col = C_ReasonCdIn		' Reason Code
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_LotNoIn			'C_LotNo
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_LotSubNoIn
					strVal = strVal & Trim(.Text) & iColSep
					'	item_document_no
					If UCase(Trim(GetSpreadText(frm1.vspdData1,C_AutoRcptFlg,IntRows,"X","X"))) = "Y" _
						And UCase(Trim(GetSpreadText(frm1.vspdData1,C_ReportTypeIn,IntRows,"X","X"))) = "G" _
						And (UCase(Trim(GetSpreadText(frm1.vspdData1,C_RoutOrder ,IntRows,"X","X"))) = "S" _
							Or UCase(Trim(GetSpreadText(frm1.vspdData1,C_RoutOrder ,IntRows,"X","X"))) = "L") Then
						strVal = strVal & UCase(Trim(frm1.txtRcptNo.value)) & iColSep
					Else
						strVal = strVal & iColSep	
					End If
					
					.Col = C_Remark
					strVal = strVal & Trim(.Text) & iColSep    '15		
					' Qty_base_unit 
					strVal = strVal & UNIConvNum("0",0) & iColSep
					' Subcontract Price
					strVal = strVal & iColSep   '10
					' Subcontract Amount
					strVal = strVal & iColSep
					' Currency
					strVal = strVal & iColSep
					strVal = strVal & IntRows & iRowSep
					
					If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
					                        
						Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
						objTEXTAREA.name = "txtCUSpread"
						objTEXTAREA.value = Join(iTmpCUBuffer,"")
						divTextArea.appendChild(objTEXTAREA)     
			 
						iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
						ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
						iTmpCUBufferCount = -1
						strCUTotalvalLen  = 0
					 End If
					   
					 iTmpCUBufferCount = iTmpCUBufferCount + 1
					  
					 If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
					    iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
					    ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
					 End If   
					     
					 iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
					 strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
					
			End Select
			
	    Next
	    
	End With
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
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
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
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
	Call DbQuery
	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
    On Error Resume Next
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
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
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
	If DbDtlQuery(lRow) = False Then	
		Call RestoreToolBar()
		Exit Function
	End If
End Function

'==============================================================================
' Function : SetFieldColor
' Description : �߰� �Է� �ʵ��� Color�� ����. 
'==============================================================================
Function SetFieldColor(BlnQueryOk) 

	If BlnQueryOk  = True Then
		Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
		If UCase(Trim(GetSpreadText(frm1.vspdData1,C_AutoRcptFlg,1,"X","X"))) = "Y" Then
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"N")
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"Q")
		End If	
	
		frm1.txtReportDt.text	= LocSvrDate
		frm1.txtRcptNo.value = ""
	Else
		Call ggoOper.LockField(Document, "Q")                                   '��: Lock  Suitable  Field
	
		frm1.txtReportDt.text = ""
		frm1.txtRcptNo.value = ""
		frm1.cboShift.value = ""
	End If
End Function