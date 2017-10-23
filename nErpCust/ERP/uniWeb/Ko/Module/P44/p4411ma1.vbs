'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'Grid 1 - Order Header
Const BIZ_PGM_QRY1_ID	= "p4411mb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 
'Grid 2 - Production Results
Const BIZ_PGM_QRY2_ID	= "p4411mb2.asp"								'��: �����Ͻ� ���� ASP�� 
'Post Production Results
Const BIZ_PGM_SAVE_ID	= "p4412mb3.asp"
'Shift Header
Const BIZ_PGM_SHIFT		= "p4400mb1.asp"								'��: �����Ͻ� ���� ASP�� 

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
Dim C_CurrencyCode			
Dim C_SubcontractPriceIn		
Dim C_SubcontractAmtIn		
Dim C_Remark				
Dim C_LotNoIn				
Dim C_RcptDocumentNoIn		
Dim C_IssueDocumentNoIn		
Dim C_InspReqNoIn				
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
Dim C_BpCd					
Dim C_BpNm					
Dim C_JobCd					
Dim C_JobDesc				
Dim C_RoutOrder			
Dim C_SubcontractPrice		
Dim C_SubcontractAmt		
Dim C_OrderStatus			
Dim C_OrderStatusNm		
Dim C_TaxType				
Dim C_InsideFlag			
Dim C_InsideFlagNm			
Dim C_TrackingNo			
Dim C_ProdtOrderType		
Dim C_AutoRcptFlg			
Dim C_LotReq			
Dim C_MilestoneFlg			
Dim C_ProdInspReq				
Dim C_FinalInspReq
Dim C_ItemGroupCd
Dim C_ItemGroupNm			
Dim C_OrderQtyInBaseUnit	

' Grid 2(vspdData2) - Results
Dim C_ReportDt				
Dim C_ReportType			
Dim C_ShiftId				
Dim C_ProdQty				
Dim C_ReasonCd				
Dim C_ReasonDesc			
Dim C_Remark1				
Dim C_LotNo					
Dim C_RcptDocumentNo		
' Hidden
Dim C_IssueDocumentNo		
Dim C_InspReqNo				
Dim C_SubcontractPrice1		
Dim C_SubcontractAmt1		
Dim C_CurrencyCode1			
Dim C_CurrencyPopup1		
Dim C_TaxType1				
Dim C_TaxPopup1				
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
Dim C_Insp_Good_Qty1		
Dim C_Insp_Bad_Qty1			
Dim C_Rcpt_Qty1				

' Grid 3(vspdData3) - Hidden
Dim C_ReportDt2				
Dim C_ReportType2			
Dim C_ShiftId2				
Dim C_ProdQty2				
Dim C_ReasonCd2				
Dim C_ReasonDesc2			
Dim C_Remark2				
Dim C_LotNo2				
Dim C_RcptDocumentNo2		
Dim C_IssueDocumentNo2		
Dim C_InspReqNo2				
Dim C_SubcontractPrice2		
Dim C_SubcontractAmt2		
Dim C_CurrencyCode2			
Dim C_CurrencyPopup2			
Dim C_TaxType2				
Dim C_TaxPopUp2				
Dim C_ProdtOrderNo2			
Dim C_OprNo2				
Dim C_Sequence2				
Dim C_MilestoneFlg2			
Dim C_InsideFlag2			
Dim C_AutoRcptFlg2			
Dim C_LotReq2				
Dim C_ProdInspReq2			
Dim C_FinalInspReq2			
Dim C_Insp_Good_Qty2			
Dim C_Insp_Bad_Qty2			
Dim C_Rcpt_Qty2				

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
Dim lgSortkey2
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
	lgSortKey1 = 1
	lgSortKey2 = 1
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
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)
	
	Call AppendNumberPlace("7", "5", "0")
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20030913", ,Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_OrderQtyInBaseUnit +1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_ProdtOrderNo,			"����������ȣ", 18
			ggoSpread.SSSetEdit		C_OprNo,				"����", 8
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
			ggoSpread.SSSetEdit		C_CurrencyCode,			"������ȭ", 8
			ggoSpread.SSSetFloat	C_SubcontractPriceIn,	"���ִܰ�",15,"C" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_SubcontractAmtIn,		"���ֱݾ�",15,"A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetEdit		C_Remark,				"���", 20,,,120
			ggoSpread.SSSetEdit		C_LotNoIn,				"Lot No.", 12,,,12,2	
			ggoSpread.SSSetEdit		C_RcptDocumentNoIn,		"�԰��ȣ", 18,,,16,2
			ggoSpread.SSSetEdit		C_IssueDocumentNoIn,	"����ȣ", 18,,,16,2	
			ggoSpread.SSSetEdit		C_InspReqNoIn,			"�˻��Ƿڹ�ȣ", 18,,,18,2
	
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit,	"��������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit,	"��ǰ����",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInOrderUnit,	"�ҷ�����",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_InspGoodQtyInOrderUnit,"ǰ����ǰ",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_InspBadQtyInOrderUnit,"ǰ���ҷ�",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit,	"�԰����",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_PlanStartDt,			"����������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_PlanComptDt,			"�ϷΌ����", 11, 2, parent.gDateFormat	
			
			ggoSpread.SSSetDate 	C_ReleaseDt,			"�۾�������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_RealStartDt,			"��������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_RoutNo,				"�����", 10
			ggoSpread.SSSetEdit		C_WcCd,					"�۾���", 10
			ggoSpread.SSSetEdit		C_WcNm,					"�۾����", 20
			ggoSpread.SSSetEdit		C_BpCd,					"�ŷ�ó", 10
			ggoSpread.SSSetEdit		C_BpNm,					"�ŷ�ó��", 20
			ggoSpread.SSSetEdit		C_JobCd,				"�۾�", 8
			ggoSpread.SSSetEdit		C_JobDesc,				"�۾���", 20
			ggoSpread.SSSetEdit		C_RoutOrder,			"�۾�����", 8
			ggoSpread.SSSetFloat	C_SubcontractPrice,		"���ִܰ�",15,"C" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_SubcontractAmt,		"���ֱݾ�",15,"A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetEdit		C_OrderStatus,			"���û���", 10
			ggoSpread.SSSetEdit		C_OrderStatusNm,		"���û���", 10
			ggoSpread.SSSetEdit		C_TaxType,				"VAT����", 8
					
			ggoSpread.SSSetEdit		C_InsideFlag,			"�系/��", 10	
			ggoSpread.SSSetEdit		C_InsideFlagNm,			"�系/��", 10
			ggoSpread.SSSetEdit		C_TrackingNo,			"Tracking No.", 25,,,25
			ggoSpread.SSSetEdit		C_ProdtOrderType,		"���ñ���", 10
			 
			ggoSpread.SSSetEdit		C_AutoRcptFlg,			"", 10
			ggoSpread.SSSetEdit		C_LotReq,				"", 10
			ggoSpread.SSSetEdit		C_MilestoneFlg,			"Milestone", 10
			ggoSpread.SSSetEdit		C_ProdInspReq,			"�����˻�", 8
			ggoSpread.SSSetEdit		C_FinalInspReq,			"", 10
			ggoSpread.SSSetEdit 	C_ItemGroupCd, "ǰ��׷�",	15
			ggoSpread.SSSetEdit		C_ItemGroupNm, "ǰ��׷��", 30
			ggoSpread.SSSetFloat	C_OrderQtyInBaseUnit,	"��������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
   
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_LotNoIn, C_LotNoIn, True)
			Call ggoSpread.SSSetColHidden(C_RcptDocumentNoIn, C_RcptDocumentNoIn, True)
			Call ggoSpread.SSSetColHidden(C_IssueDocumentNoIn, C_IssueDocumentNoIn, True)
			Call ggoSpread.SSSetColHidden(C_InspReqNoIn, C_InspReqNoIn, True)
			Call ggoSpread.SSSetColHidden(C_RoutOrder, C_RoutOrder, True)
			Call ggoSpread.SSSetColHidden(C_OrderStatus, C_OrderStatus, True)
			Call ggoSpread.SSSetColHidden(C_InsideFlag, C_InsideFlag, True)
			Call ggoSpread.SSSetColHidden(C_InsideFlagNm, C_InsideFlagNm, True)
			Call ggoSpread.SSSetColHidden(C_AutoRcptFlg, C_AutoRcptFlg, True)
			Call ggoSpread.SSSetColHidden(C_LotReq, C_LotReq, True)
			Call ggoSpread.SSSetColHidden(C_FinalInspReq, C_FinalInspReq, True)
			Call ggoSpread.SSSetColHidden(C_OrderQtyInBaseUnit, C_OrderQtyInBaseUnit, True)
			      
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
			ggoSpread.Spreadinit "V20030000", ,Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_Rcpt_Qty1 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetDate 	C_ReportDt,			"������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetCombo	C_ReportType,		"��/��", 6
			ggoSpread.SSSetEdit		C_ShiftId,			"Shift", 8
			ggoSpread.SSSetFloat	C_ProdQty,			"��������",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo	C_ReasonCd,			"�ҷ��ڵ�", 10
			ggoSpread.SSSetCombo	C_ReasonDesc,		"�ҷ�����", 20
			ggoSpread.SSSetEdit		C_Remark1,			"���", 20,,,120
			ggoSpread.SSSetEdit		C_LotNo,			"Lot No.", 12,,,12	
			ggoSpread.SSSetEdit		C_RcptDocumentNo,	"�԰��ȣ", 15,,,16,2
			ggoSpread.SSSetEdit		C_IssueDocumentNo,	"����ȣ", 15,,,16,2	
			ggoSpread.SSSetEdit		C_InspReqNo,		"�˻��Ƿڹ�ȣ", 12,,,18,2
			ggoSpread.SSSetFloat	C_SubcontractPrice1,"���ִܰ�",15,"C" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_SubcontractAmt1,	"���ֱݾ�",15,"A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetEdit		C_CurrencyCode1,	"������ȭ", 8
			ggoSpread.SSSetButton   C_CurrencyPopup1
			ggoSpread.SSSetEdit		C_TaxType1,			"VAT����", 8	
			ggoSpread.SSSetButton   C_TaxPopup1
	
			ggoSpread.SSSetEdit		C_ProdtOrderNo1,	"", 18	
			ggoSpread.SSSetEdit		C_OprNo1,			"", 10
			ggoSpread.SSSetFloat	C_Sequence,			"����", 8, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_MilestoneFlg1,	"Milestone", 10
			ggoSpread.SSSetEdit		C_InsideFlag1,		"�系/��", 10	
			ggoSpread.SSSetEdit		C_AutoRcptFlg1,		"", 10
			ggoSpread.SSSetEdit		C_LotReq1,			"", 10	
			ggoSpread.SSSetEdit		C_ProdInspReq1,		"", 10	
			ggoSpread.SSSetEdit		C_FinalInspReq1,	"", 10		
	
			ggoSpread.SSSetFloat	C_Insp_Good_Qty1,	"ǰ����ǰ",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Insp_Bad_Qty1,	"ǰ���ҷ�",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Rcpt_Qty1,		"�԰����",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_LotNo, C_LotNo, True)
			Call ggoSpread.SSSetColHidden(C_RcptDocumentNo, C_RcptDocumentNo, True)
			Call ggoSpread.SSSetColHidden(C_IssueDocumentNo, C_IssueDocumentNo, True)
			Call ggoSpread.SSSetColHidden(C_CurrencyCode1, C_CurrencyCode1, True)
			Call ggoSpread.SSSetColHidden(C_CurrencyPopup1, C_CurrencyPopup1, True)
			Call ggoSpread.SSSetColHidden(C_TaxType1, C_TaxType1, True)
			Call ggoSpread.SSSetColHidden(C_TaxPopup1, C_TaxPopup1, True)
	
			Call ggoSpread.SSSetColHidden(C_ProdtOrderNo1, C_ProdtOrderNo1, True)
			Call ggoSpread.SSSetColHidden(C_OprNo1, C_OprNo1, True)
			Call ggoSpread.SSSetColHidden(C_Sequence, C_Sequence, True)
			Call ggoSpread.SSSetColHidden(C_MilestoneFlg1, C_MilestoneFlg1, True)
			Call ggoSpread.SSSetColHidden(C_InsideFlag1, C_InsideFlag1, True)
	
			Call ggoSpread.SSSetColHidden(C_AutoRcptFlg1, C_AutoRcptFlg1, True)
			Call ggoSpread.SSSetColHidden(C_LotReq1, C_LotReq1, True)
			Call ggoSpread.SSSetColHidden(C_ProdInspReq1, C_ProdInspReq1, True)
			Call ggoSpread.SSSetColHidden(C_FinalInspReq1, C_FinalInspReq1, True)
			Call ggoSpread.SSSetColHidden(C_Rcpt_Qty1, C_Rcpt_Qty1, True)  ' hidden for rcpt_qty 20030403 kjp
	
			ggoSpread.SSSetSplit2(5)
			
			Call SetSpreadLock("B")
				
			.ReDraw = true
    
		End With
	End If	
	
	If pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 3 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData3
			
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit
			
			.ReDraw = false
			
			.MaxCols = C_Rcpt_Qty2 +1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			
			ggoSpread.SSSetDate 	C_ReportDt2,		"������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_ReportType2,		"��/��", 6
			ggoSpread.SSSetEdit		C_ShiftId2,			"Shift", 8	
			ggoSpread.SSSetFloat	C_ProdQty2,			"��������",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ReasonCd2,		"�ҷ��ڵ�", 10
			ggoSpread.SSSetEdit		C_ReasonDesc2,		"�ҷ�����", 15
			ggoSpread.SSSetEdit		C_Remark2,			"���", 20,,,120
			ggoSpread.SSSetEdit		C_LotNo2,			"Lot No.", 17,,,16		
			ggoSpread.SSSetEdit		C_RcptDocumentNo2,	"�԰��ȣ", 18,,,16
			ggoSpread.SSSetEdit		C_IssueDocumentNo2, "����ȣ", 18,,,16	
			ggoSpread.SSSetEdit		C_InspReqNo2,		"�˻��Ƿڹ�ȣ", 18,,,18	
			ggoSpread.SSSetFloat	C_SubcontractPrice2,"���ִܰ�",15,"C" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_SubcontractAmt2,	"���ֱݾ�",15,"A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetEdit		C_CurrencyCode2,	"������ȭ", 8
			ggoSpread.SSSetEdit		C_TaxType2,			"VAT����", 8
			ggoSpread.SSSetEdit		C_ProdtOrderNo2,	"������ȣ", 18
			ggoSpread.SSSetFloat	C_Sequence2,		"����",	8, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_InsideFlag2,		"�系/��", 10	
			ggoSpread.SSSetEdit		C_MilestoneFlg2,	"Milestone", 10	
			ggoSpread.SSSetFloat	C_Insp_Good_Qty2,	"ǰ����ǰ",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Insp_Bad_Qty2,	"ǰ���ҷ�",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Rcpt_Qty2,		"�԰����",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
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
			.vspdData1.ReDraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData1.ReDraw = True
		End If
		
		If pvSpdNo = "B" Then    
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "C" Then
			'--------------------------------
			'Grid 3
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
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub


'==========================================  2.2.6 InitSpreadComboBox()  =======================================
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

	If Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantNm.value = ""
		Exit Function
	Else
		For i = lgShiftCnt To 1 Step -1
			frm1.cboShift.remove(i) 
		Next
	End If
	
    With frm1
	
	strVal = BIZ_PGM_SHIFT & "?txtMode=" & parent.UID_M0001						'��: 
	strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    End With
	
	InitShiftCombo = True
	
End Function

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
		C_CurrencyCode				= 13
		C_SubcontractPriceIn		= 14
		C_SubcontractAmtIn			= 15
		C_Remark					= 16
		C_LotNoIn					= 17
		C_RcptDocumentNoIn			= 18
		C_IssueDocumentNoIn			= 19
		C_InspReqNoIn				= 20
		C_ProdQtyInOrderUnit		= 21
		C_GoodQtyInOrderUnit		= 22
		C_BadQtyInOrderUnit			= 23
		C_InspGoodQtyInOrderUnit	= 24
		C_InspBadQtyInOrderUnit		= 25
		C_RcptQtyInOrderUnit		= 26
		C_PlanStartDt				= 27
		C_PlanComptDt				= 28
		C_ReleaseDt					= 29
		C_RealStartDt				= 30
		C_RoutNo					= 31
		C_WcCd						= 32
		C_WcNm						= 33
		C_BpCd						= 34
		C_BpNm						= 35
		C_JobCd						= 36
		C_JobDesc					= 37
		C_RoutOrder					= 38
		C_SubcontractPrice			= 39
		C_SubcontractAmt			= 40
		C_OrderStatus				= 41
		C_OrderStatusNm				= 42
		C_TaxType					= 43
		C_InsideFlag				= 44
		C_InsideFlagNm				= 45
		C_TrackingNo				= 46
		C_ProdtOrderType			= 47
		C_AutoRcptFlg				= 48
		C_LotReq					= 49
		C_MilestoneFlg				= 50
		C_ProdInspReq				= 51
		C_FinalInspReq				= 52
		C_ItemGroupCd				= 53
		C_ItemGroupNm				= 54
		C_OrderQtyInBaseUnit		= 55
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
		C_LotNo						= 8
		C_RcptDocumentNo			= 9
		' Hidden
		C_IssueDocumentNo			= 10
		C_InspReqNo					= 11
		C_SubcontractPrice1			= 12
		C_SubcontractAmt1			= 13
		C_CurrencyCode1				= 14
		C_CurrencyPopup1			= 15
		C_TaxType1					= 16
		C_TaxPopup1					= 17
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
		C_Insp_Good_Qty1			= 27
		C_Insp_Bad_Qty1				= 28
		C_Rcpt_Qty1					= 29
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
		C_LotNo2					= 8
		C_RcptDocumentNo2			= 9
		C_IssueDocumentNo2			= 10
		C_InspReqNo2				= 11
		C_SubcontractPrice2			= 12
		C_SubcontractAmt2			= 13
		C_CurrencyCode2				= 14
		C_CurrencyPopup2			= 15
		C_TaxType2					= 16
		C_TaxPopUp2					= 17
		C_ProdtOrderNo2				= 18
		C_OprNo2					= 19
		C_Sequence2					= 20
		C_MilestoneFlg2				= 21
		C_InsideFlag2				= 22
		C_AutoRcptFlg2				= 23
		C_LotReq2					= 24
		C_ProdInspReq2				= 25
		C_FinalInspReq2				= 26
		C_Insp_Good_Qty2			= 27
		C_Insp_Bad_Qty2				= 28
		C_Rcpt_Qty2					= 29
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
			C_CurrencyCode				= iCurColumnPos(13)
			C_SubcontractPriceIn		= iCurColumnPos(14)
			C_SubcontractAmtIn			= iCurColumnPos(15)
			C_Remark					= iCurColumnPos(16)
			C_LotNoIn					= iCurColumnPos(17)
			C_RcptDocumentNoIn			= iCurColumnPos(18)
			C_IssueDocumentNoIn			= iCurColumnPos(19)
			C_InspReqNoIn				= iCurColumnPos(20)
			C_ProdQtyInOrderUnit		= iCurColumnPos(21)
			C_GoodQtyInOrderUnit		= iCurColumnPos(22)
			C_BadQtyInOrderUnit			= iCurColumnPos(23)
			C_InspGoodQtyInOrderUnit	= iCurColumnPos(24)
			C_InspBadQtyInOrderUnit		= iCurColumnPos(25)
			C_RcptQtyInOrderUnit		= iCurColumnPos(26)
			C_PlanStartDt				= iCurColumnPos(27)
			C_PlanComptDt				= iCurColumnPos(28)
			C_ReleaseDt					= iCurColumnPos(29)
			C_RealStartDt				= iCurColumnPos(30)
			C_RoutNo					= iCurColumnPos(31)
			C_WcCd						= iCurColumnPos(32)
			C_WcNm						= iCurColumnPos(33)
			C_BpCd						= iCurColumnPos(34)
			C_BpNm						= iCurColumnPos(35)
			C_JobCd						= iCurColumnPos(36)
			C_JobDesc					= iCurColumnPos(37)
			C_RoutOrder					= iCurColumnPos(38)
			C_SubcontractPrice			= iCurColumnPos(39)
			C_SubcontractAmt			= iCurColumnPos(40)
			C_OrderStatus				= iCurColumnPos(41)
			C_OrderStatusNm				= iCurColumnPos(42)
			C_TaxType					= iCurColumnPos(43)
			C_InsideFlag				= iCurColumnPos(44)
			C_InsideFlagNm				= iCurColumnPos(45)
			C_TrackingNo				= iCurColumnPos(46)
			C_ProdtOrderType			= iCurColumnPos(47)
			C_AutoRcptFlg				= iCurColumnPos(48)
			C_LotReq					= iCurColumnPos(49)
			C_MilestoneFlg				= iCurColumnPos(50)
			C_ProdInspReq				= iCurColumnPos(51)
			C_FinalInspReq				= iCurColumnPos(52)
			C_ItemGroupCd				= iCurColumnPos(53)
			C_ItemGroupNm				= iCurColumnPos(54)
			C_OrderQtyInBaseUnit		= iCurColumnPos(55)


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
			C_LotNo						= iCurColumnPos(8)
			C_RcptDocumentNo			= iCurColumnPos(9)
			' Hidden
			C_IssueDocumentNo			= iCurColumnPos(10)
			C_InspReqNo					= iCurColumnPos(11)
			C_SubcontractPrice1			= iCurColumnPos(12)
			C_SubcontractAmt1			= iCurColumnPos(13)
			C_CurrencyCode1				= iCurColumnPos(14)
			C_CurrencyPopup1			= iCurColumnPos(15)
			C_TaxType1					= iCurColumnPos(16)
			C_TaxPopup1					= iCurColumnPos(17)
			'Hidden
			C_ProdtOrderNo1				= iCurColumnPos(18)
			C_OprNo1					= iCurColumnPos(19)
			C_Sequence					= iCurColumnPos(20)
			C_MilestoneFlg1				= iCurColumnPos(21)
			C_InsideFlag1				= iCurColumnPos(22)
			C_AutoRcptFlg1				= iCurColumnPos(23)
			C_LotReq1					= iCurColumnPos(24)
			C_ProdInspReq1				= iCurColumnPos(25)
			C_FinalInspReq1				= iCurColumnPos(26)
			C_Insp_Good_Qty1			= iCurColumnPos(27)
			C_Insp_Bad_Qty1				= iCurColumnPos(28)
			C_Rcpt_Qty1					= iCurColumnPos(29)

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

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
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

'------------------------------------------  OpenBizPartner()  -------------------------------------------------
'	Name : OpenBizparener()
'	Description : BpPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó�˾�"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = frm1.txtBpCd.value 
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "����ó"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    arrField(2) = "BP_TYPE"
    arrField(3) = ""	
        
    arrHeader(0) = "BP"		
    arrHeader(1) = "BP��"		
    arrHeader(2) = "Bp ����"		
    arrHeader(3) = ""
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtBpCd.Value    = arrRet(0)		
		frm1.txtBpNm.Value    = arrRet(1)	
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBpCd.focus
	
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
		frm1.vspdData1.Col = C_OprNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
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

'------------------------------------------  txtPlantCd_OnChange -----------------------------------------
'	Name : txtPlantCd_OnChange()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlantCd_OnChange()

	Dim LngRows
	
	If frm1.txtPlantCd.value = "" Then

	Else
		
	End If	
End Sub

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
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*******************************************************************************************************

'******************************  3.2.1 Object Tag ó��  ************************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*******************************************************************************************************

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
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


'========================================================================================================
'   Event Name : vspdData2_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData1_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_SubcontractPriceIn
            Call EditModeCheck(frm1.vspdData1, Row, C_CurrencyCode, C_SubcontractPriceIn, "C" ,"I", Mode, "X", "X")
        Case C_SubcontractAmtIn
            Call EditModeCheck(frm1.vspdData1, Row, C_CurrencyCode, C_SubcontractAmtIn, "A" ,"I", Mode, "X", "X")        
    End Select
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
				.Col = C_ProdQtyIn
				dblProdQty = UNICDbl(.Text)
				.Col = C_ProdtOrderQty
				dblProdtOrderQty = UNICDbl(.Text)
				.Col = C_OrderQtyInBaseUnit
				dblOrderQtyInBaseUnit = UNICDbl(.Text)
				
				If dblProdQty * dblOrderQtyInBaseUnit > 0 Then
					dblProdQty = (dblProdQty * dblOrderQtyInBaseUnit) / dblProdtOrderQty
				End IF

				.Col = C_SubcontractPriceIn
				dblSubcontractPrice = UNICDbl(.Text)
				
				.Col = C_SubcontractAmtIn
				.Text = UNIFormatNumber(dblProdQty * dblSubcontractPrice,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)

				.ReDraw = True
				
				.Col = C_ProdQtyIn
				If strMilestoneFlag = "Y" and strInsideFlag = "N" and UNICDbl(.value) > 0 Then
				
					ggoSpread.Source = frm1.vspdData1
					ggoSpread.SpreadUnLock C_ProdQtyIn, Row, C_ReasonCdIn, Row
					ggoSpread.SpreadUnLock C_ReportTypeIn, Row, C_ReasonDescIn, Row
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
						Else
							ggoSpread.SpreadUnLock C_LotNoIn,Row,C_LotNoIn,Row
						End If
						If strAutoRcptFlg <> "Y" or (strRoutOrder <> "S" and strRoutOrder <> "L") Then
							ggoSpread.SSSetProtected C_RcptDocumentNoIn, Row, Row
						Else
							ggoSpread.SpreadUnLock C_RcptDocumentNoIn,Row,C_RcptDocumentNoIn,Row
						End If
						
					Else
						ggoSpread.SpreadUnLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
						ggoSpread.SpreadUnLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
						ggoSpread.SSSetRequired C_ReasonCdIn, Row, Row
						ggoSpread.SSSetRequired C_ReasonDescIn, Row, Row
					End If
					ggoSpread.SpreadUnLock C_SubcontractPriceIn, Row, C_SubcontractPriceIn, Row
					ggoSpread.SpreadUnLock C_SubcontractAmtIn, Row, C_SubcontractAmtIn, Row
					ggoSpread.SpreadUnLock C_Remark, Row, C_Remark, Row
					ggoSpread.SSSetRequired C_SubcontractPriceIn,Row,Row
					ggoSpread.SSSetRequired C_SubcontractAmtIn,Row,Row
					
					ggoSpread.UpdateRow Row
					
				Else

					ggoSpread.Source = frm1.vspdData1
					ggoSpread.SpreadUnLock C_ProdQtyIn,Row,C_ProdQtyIn,Row
					ggoSpread.SpreadLock C_ReportTypeIn,Row,C_ReportTypeIn,Row
					ggoSpread.SpreadLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
					ggoSpread.SpreadLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
					ggoSpread.SpreadLock C_Remark, Row, C_Remark, Row
					ggoSpread.SSSetProtected C_ReportTypeIn, Row, Row
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
					.Col = C_RcptDocumentNoIn
					.Text = ""
					.Col = C_Remark
					.Text = ""
					
					ggoSpread.SpreadLock C_SubcontractPriceIn,Row,C_SubcontractPriceIn,Row
					ggoSpread.SpreadLock C_SubcontractAmtIn,Row,C_SubcontractAmtIn,Row
					ggoSpread.SSSetProtected C_SubcontractPriceIn, Row, Row
					ggoSpread.SSSetProtected C_SubcontractAmtIn, Row, Row
					
					If strLotReq <> "Y" or strAutoRcptFlg <> "Y" or (strRoutOrder <> "S" and strRoutOrder <> "L")  Then
						ggoSpread.SSSetProtected C_LotNoIn, Row, Row
					Else
						ggoSpread.SpreadLock C_LotNoIn,Row,C_LotNoIn,Row
					End If
					If strAutoRcptFlg <> "Y" or (strRoutOrder <> "S" and strRoutOrder <> "L") Then
						ggoSpread.SSSetProtected C_RcptDocumentNoIn, Row, Row
					Else
						ggoSpread.SpreadLock C_RcptDocumentNoIn,Row,C_RcptDocumentNoIn,Row
					End If
					
					ggoSpread.SSDeleteFlag Row,Row
					
				End If

				.ReDraw = True

		    Case C_SubcontractPriceIn

				ggoSpread.Source = frm1.vspdData1
				frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
				frm1.vspdData1.Col = C_ProdtOrderQty
				dblProdtOrderQty = UNICDbl(frm1.vspdData1.Text)
				frm1.vspdData1.Col = C_OrderQtyInBaseUnit
				dblOrderQtyInBaseUnit = UNICDbl(frm1.vspdData1.Text)
				
				ggoSpread.Source = frm1.vspdData1
				.Col = C_ProdQtyIn
				dblProdQty = UNICDbl(.Text)
				
				dblProdQty = (dblProdQty * dblOrderQtyInBaseUnit) / dblProdtOrderQty
				
				.Col = C_SubcontractPriceIn
				dblSubcontractPrice = UNICDbl(.Text)
				.Col = C_SubcontractAmtIn
				.Text = UNIFormatNumber(dblProdQty * dblSubcontractPrice,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				ggoSpread.Source = frm1.vspdData1
				
'				ggoSpread.UpdateRow Row
		    
		    Case C_SubcontractAmtIn
		    
				ggoSpread.Source = frm1.vspdData1
				
'				ggoSpread.UpdateRow Row
				
		End Select

	End With

End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex

	ggoSpread.Source = frm1.vspdData1
'	ggoSpread.UpdateRow Row

	With frm1.vspdData1

		.Row = Row
		Select Case Col
		
		    Case C_ReportTypeIn
		       	
				.Col = Col
				.Row = Row
				
				.ReDraw = False
				
				If Trim(.Text) = "G" Then
				
					ggoSpread.SpreadLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
					ggoSpread.SpreadLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
					ggoSpread.SSSetProtected C_ReasonCdIn, Row, Row
					ggoSpread.SSSetProtected C_ReasonDescIn, Row, Row

					.Col = C_ReasonCdIn
					.Text = ""
					.Col = C_ReasonDescIn
					.Text = ""
				Else
					ggoSpread.SpreadUnLock C_ReasonCdIn, Row, C_ReasonCdIn, Row
					ggoSpread.SpreadUnLock C_ReasonDescIn, Row, C_ReasonDescIn, Row
					ggoSpread.SSSetRequired C_ReasonCdIn, Row, Row
					ggoSpread.SSSetRequired C_ReasonDescIn, Row, Row
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

		    Case C_CurrencyPopup1
				.Col = C_CurrencyCode1
				.Row = Row
				strCode = .Text
				Call OpenCurrency(strCode, Row)
				Call SetActiveCell(frm1.vspdData,C_CurrencyCode1,Row,"M","X","X")
				Set gActiveElement = document.activeElement
	    
		    Case C_TaxPopup1
				.Col = C_TaxType1
				.Row = Row
				strCode = .Text
				Call OpenTaxType(strCode, Row)
				Call SetActiveCell(frm1.vspdData,C_TaxType1,Row,"M","X","X")
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
'   Event Name : vspdData_TopLeftChange
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

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    Call InitSpreadComboBox(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
	
	If gActiveSpdSheet.Id = "A" Then
		Call InitData(1)
		
		frm1.vspdData1.Redraw = False
		
		For LngRow = 1 To frm1.vspdData1.MaxRows
			ggoSpread.Source = frm1.vspdData1			
			ggoSpread.SpreadUnLock C_ProdQtyIn, LngRow, C_ProdQtyIn, LngRow
		Next
		
		frm1.vspdData1.Redraw = True
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
    
    ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
	If InitShiftCombo = False Then Exit Function
   
    FncQuery = True												'��: Processing is OK
        
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
				ggoSpread.EditUndo                                                  '��: Protect system from crashing
				Call vspdData1_Change(C_ProdQtyIn, frm1.vspdData1.ActiveRow)
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

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  ******************************
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
		strVal = strVal & "&txtBpCd=" & Trim(.hBpCd.Value)
		strVal = strVal & "&txtProdFromDt=" & Trim(.hProdFromDt.Value)
		strVal = strVal & "&txtProdTODt=" & Trim(.hProdToDt.Value)
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
		strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.Value)
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

	Call InitData(LngMaxRow)
	frm1.txtReportDt.text	= LocSvrDate
	Call ggoOper.LockField(Document, "N")										'��: It's not Standard, This function unlock the suitable field

	frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1

	lgOldRow = 1

    With frm1.vspdData1

		.Redraw = False
		
		For lRow = LngMaxRow To .MaxRows
			
			ggoSpread.Source = frm1.vspdData1

			.Col = C_MilestoneFlg						
			.Row = lRow	

			If Trim(.text) = "Y" Then	
				ggoSpread.SpreadUnLock C_ProdQtyIn,lRow,C_ProdQtyIn,lRow
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
' Function Desc : DbQuery�� ������ ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryNotOk()

	Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
    frm1.txtReportDt.text	= ""
    frm1.cboShift.value = ""
	Call ggoOper.LockField(Document, "Q")										'��: It's not Standard, This function lock the suitable field
	
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

	    If CopyFromHSheet(strProdtOrderNo, strOprNo) = True Then
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2, 1, Frm1.vspdData2.MaxRows, C_CurrencyCode1,C_SubcontractPrice1, "C", "I", "X", "X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2, 1, Frm1.vspdData2.MaxRows, C_CurrencyCode1,C_SubcontractAmt1, "A", "I", "X", "X")    	               
           Exit Function
        End If

		DbDtlQuery = False   
    
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
            For iCols = 1 To .vspdData2.MaxCols 'vspdData2 �� ����Ÿ�� �����Ѵ�.
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
      
            For iCols = 1 To .vspdData2.MaxCols 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCols
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
		    
			    Case ggoSpread.UpdateFlag
					
					strVal = ""
					
					.Col = C_ProdtOrderNo	' Production Order No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_OprNo	' Operation No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_ReportTypeIn	' Report Type
					strVal = strVal & Trim(.Text) & iColSep     '5
					.Col = C_ProdQtyIn		' Produced Qty
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
					' Produced Date
					strVal = strVal & UNIConvDate(frm1.txtReportDT.Text) & iColSep
					If CompareDateByFormat(frm1.txtReportDT.Text, LocSvrDate,"������","������","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
					  Call LayerShowHide(0)
					   .EditMode = True
					   frm1.txtReportDT.focus
					   strVal = ""
					  Exit Function               
					End If 
					' Shift
					strVal = strVal & frm1.cboShift.value & iColSep
					.Col = C_ReasonCdIn			'Reason Code
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_LotNoIn		    'Lot No	
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					'Lot Sub No.
					strVal = strVal & iColSep
					.Col = C_RcptDocumentNoIn	'C_RcptDocumentNo
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_Remark
					strVal = strVal & Trim(.Text) & iColSep    '15
				    ' Qty_base_unit 
					strVal = strVal & UNIConvNum("0",0) & iColSep
					
					.Col = C_SubcontractPriceIn	' Subcontract Price
					If UNICDbl(.Value) <= 0 Then
						Call LayerShowHide(0)
						.EditMode = True
						Call Displaymsgbox("189306", "x", "x", "x")
						Call SheetFocus(IntRows,C_SubcontractPriceIn)
						strVal = ""
						Exit Function
					End If
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep   '10
					
					.Col = C_SubcontractAmtIn	' Subcontract Price
					If UNICDbl(.Value) <= 0 Then
						Call LayerShowHide(0)
						.EditMode = True
						Call Displaymsgbox("189306", "x", "x", "x")
						Call SheetFocus(IntRows,C_SubcontractAmtIn)
						strVal = ""
						Exit Function
					End If
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
					
					.Col = C_CurrencyCode		' Subcontract Price
					strVal = strVal & UCase(Trim(.Text)) & iColSep
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
	
	Call MainQuery()
	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
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
