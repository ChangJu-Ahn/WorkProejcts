
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_LOOKUP_ID	= "p4211mb0.asp"								' Lookup Item By Plant

'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p4211mb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 
'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p4211mb2.asp"								'��: �����Ͻ� ���� ASP�� 
'Save
Const BIZ_PGM_SAVE_ID	= "p4211mb3.asp"

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
' Grid 1(vspdData1) - Operation 
Dim C_OprNo				'= 1
Dim C_JobCd				'= 2
Dim C_JobDesc			'= 3
Dim C_WcCd				'= 4
Dim C_WcNm				'= 5
Dim C_PlanStartDt		'= 6
Dim C_PlanEndDt			'= 7
Dim C_OrderStatus		'= 8
Dim C_OrderStatusDesc	'= 9
Dim C_InsideFlag		'= 10
Dim C_InsideFlagDesc	'= 11
Dim C_MileStone			'= 12

' Grid 2(vspdData2) - Operation
Dim C_CompntCd			'= 1
Dim C_CompntCdPopup		'= 2
Dim C_CompntNm			'= 3
Dim C_Spec				'= 4
Dim C_RqrdQty			'= 5
Dim C_Unit				'= 6
Dim C_IssuedQty			'= 7
Dim C_RqrdDt			'= 8
Dim C_TrackingNo		'= 9
Dim C_MajorSLCd			'= 10
Dim C_MajorSLCdPopUp	'= 11
Dim C_MajorSLNm			'= 12
Dim C_ResrvStatus		'= 13
Dim C_ResrvDesc			'= 14
Dim C_IssueMeth			'= 15
Dim C_IssueMethDesc		'= 16
Dim C_ReqNo				'= 17
Dim C_ReqSeqNo			'= 18
' Hidden
Dim C_PlantCd			'= 19
Dim C_ProdtOrderNo		'= 20
Dim C_WcCd2				'= 21
Dim C_OprNo2			'= 22
Dim C_HndCompntCd		'= 23
Dim C_HdnOprStatus		'= 24

' Grid 3(vspdData3) - Hidden
Dim C_CompntCd3			'= 1
Dim C_CompntCdPopup3	'= 2
Dim C_CompntNm3			'= 3
Dim C_Spec3				'= 4
Dim C_RqrdQty3			'= 5
Dim C_Unit3				'= 6
Dim C_IssuedQty3		'= 7
Dim C_RqrdDt3			'= 8
Dim C_TrackingNo3		'= 9
Dim C_MajorSLCd3		'= 10
Dim C_MajorSLCdPopUp3	'= 11
Dim C_MajorSLNm3		'= 12
Dim C_ResrvStatus3		'= 13
Dim C_ResrvDesc3		'= 14
Dim C_IssueMeth3		'= 15
Dim C_IssueMethDesc3	'= 16
Dim C_ReqNo3			'= 17
Dim C_ReqSeqNo3			'= 18
Dim C_PlantCd3			'= 19
Dim C_ProdtOrderNo3		'= 20
Dim C_WcCd3				'= 21
Dim C_OprNo3			'= 22
Dim C_HndCompntCd3		'= 23
Dim C_HdnOprStatus3		'= 24
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size�� ������ ���� 
Dim lgIntFlgMode								'Variable is for Operation Status
Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgCurrRow
Dim lgFlgQueryCnt

Dim lgSortKey1
Dim lgSortKey2
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgRow         
'++++++++++++++++  Insert Your Code for Global Variables Assign  +++++++++++++++++++++++++++++++++++++

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
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""							'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()

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
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
    
			.ReDraw = false
    
			.MaxCols = C_MileStone +1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
    
			Call GetSpreadColumnPos("A")
    
			ggoSpread.SSSetEdit		C_OprNo,		"����", 10
			ggoSpread.SSSetCombo	C_JobCd,		"�۾�", 10
			ggoSpread.SSSetCombo	C_JobDesc,		"�۾���", 20
			ggoSpread.SSSetEdit		C_WcCd,			"�۾���", 10				
			ggoSpread.SSSetEdit		C_WcNm,			"�۾����", 20	
			ggoSpread.SSSetCombo	C_OrderStatus,	"���û���", 12
			ggoSpread.SSSetCombo	C_OrderStatusDesc, "���û���", 12
			ggoSpread.SSSetDate 	C_PlanStartDt,	"����������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_PlanEndDt,	"�ϷΌ����", 11, 2, parent.gDateFormat	
			ggoSpread.SSSetEdit		C_InsideFlag,	"�系/��", 12	
			ggoSpread.SSSetEdit		C_InsideFlagDesc, "�系/��", 12	
			ggoSpread.SSSetEdit		C_MileStone,	"Milestone", 12	
	
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( C_OrderStatus, C_OrderStatus, True)
			Call ggoSpread.SSSetColHidden( C_InsideFlag, C_InsideFlag, True)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
	
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
			ggoSpread.Spreadinit "V20030906", , Parent.gAllowDragDropSpread	
    
			.ReDraw = false
    
			.MaxCols = C_HdnOprStatus + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
    
			Call GetSpreadColumnPos("B")
	
			ggoSpread.SSSetEdit		C_OprNo2,		"����", 10	
			ggoSpread.SSSetEdit		C_CompntCd,		"��ǰ", 18,,,18,2
			ggoSpread.SSSetButton 	C_CompntCdPopup
			ggoSpread.SSSetEdit		C_CompntNm,		"��ǰ��", 25
			ggoSpread.SSSetEdit		C_Spec,			"�԰�", 25	
			ggoSpread.SSSetFloat	C_RqrdQty, 		"�ʿ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_Unit, 		"����", 7
			ggoSpread.SSSetFloat	C_IssuedQty,	"������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetDate 	C_RqrdDt, 		"�ʿ���", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit 	C_TrackingNo,	"Tracking No.", 25,,,25,2
			ggoSpread.SSSetEdit		C_MajorSLCd,	"���â��", 10,,,7,2
			ggoSpread.SSSetButton 	C_MajorSLCdPopup
			ggoSpread.SSSetEdit		C_MajorSLNm,	"���â���", 20
			ggoSpread.SSSetEdit		C_ResrvStatus,	"������", 10
			ggoSpread.SSSetEdit		C_ResrvDesc,	"������", 10
			ggoSpread.SSSetEdit		C_IssueMeth,	"�����", 10	
			ggoSpread.SSSetEdit		C_IssueMethDesc,"�����", 10
			ggoSpread.SSSetEdit		C_ReqNo,		"����", 6	
			ggoSpread.SSSetEdit		C_ReqSeqNo,		"����", 6
			ggoSpread.SSSetEdit		C_PlantCd,		"����", 6
			ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 18
			ggoSpread.SSSetEdit		C_WcCd2,		"�۾���", 10
			ggoSpread.SSSetEdit		C_OprNo2,		"����", 10	
			ggoSpread.SSSetEdit		C_HndCompntCd,	"��ǰ", 18
			ggoSpread.SSSetEdit		C_HdnOprStatus, "���û���", 8
			
			Call ggoSpread.MakePairsColumn(C_CompntCd, C_CompntCdPopup)
			Call ggoSpread.MakePairsColumn(C_MajorSLCd, C_MajorSLCdPopup)
			Call ggoSpread.SSSetColHidden( C_ResrvStatus, C_ResrvStatus, True)
			Call ggoSpread.SSSetColHidden( C_IssueMeth, C_IssueMeth, True)
 			Call ggoSpread.SSSetColHidden( C_ReqNo, C_HdnOprStatus, True)
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
 	
			ggoSpread.SSSetSplit2(2)	
			Call SetSpreadLock("B")
	
			.ReDraw = true
    
		End With
	End If	
	'------------------------------------------
	' Grid 3 - Hidden Spread Setting
	'------------------------------------------
	If pvSpdNo = "C" or pvSpdNo = "*" Then
		With frm1.vspdData3
			
			.MaxCols = C_HdnOprStatus3 +1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			ggoSpread.Source = frm1.vspdData3

			.ReDraw = false
			ggoSpread.Spreadinit
			ggoSpread.SSSetEdit		C_OprNo3,		"����", 10	
			ggoSpread.SSSetEdit		C_CompntCd3,	"��ǰ", 18
			ggoSpread.SSSetButton 	C_CompntCdPopup3
			ggoSpread.SSSetEdit		C_CompntNm3,	"��ǰ��", 25
			ggoSpread.SSSetEdit		C_Spec3,		"�԰�", 25
			ggoSpread.SSSetFloat	C_RqrdQty3,		"�ʿ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_IssuedQty3,	"������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_Unit3, 		"����", 7
			ggoSpread.SSSetDate 	C_RqrdDt3, 		"�ʿ���", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit 	C_TrackingNo3,	"Tracking No.", 25
			ggoSpread.SSSetEdit		C_MajorSLCd3,	"���â��", 10
			ggoSpread.SSSetButton 	C_MajorSLCdPopup3
			ggoSpread.SSSetEdit		C_MajorSLNm3,	"���â���", 20
			ggoSpread.SSSetEdit		C_ResrvStatus3, "������", 10
			ggoSpread.SSSetEdit		C_ResrvDesc3,	"������", 10
			ggoSpread.SSSetEdit		C_IssueMeth3,	"�����", 10	
			ggoSpread.SSSetEdit		C_IssueMethDesc3,"�����", 10
			ggoSpread.SSSetEdit		C_ReqNo3,		"����", 6	
			ggoSpread.SSSetEdit		C_ReqSeqNo3,	"����", 6	
			
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
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	
		
		If pvSpdNo = "B" Then
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData2

			.vspdData2.ReDraw = False
			ggoSpread.SpreadLock C_OprNo2,		-1,	C_OprNo2
			ggoSpread.SpreadLock C_CompntNm,	-1,	C_CompntNm
			ggoSpread.SpreadLock C_Spec,		-1,	C_Spec
			ggoSpread.SpreadLock C_Unit,		-1,	C_Unit
			ggoSpread.SpreadLock C_IssuedQty,	-1,	C_IssuedQty
			ggoSpread.SpreadLock C_TrackingNo,	-1,	C_TrackingNo
			ggoSpread.SpreadLock C_MajorSLNm,	-1,	C_MajorSLNm
			ggoSpread.SpreadLock C_ResrvStatus,	-1,	C_ResrvStatus
			ggoSpread.SpreadLock C_ResrvDesc,	-1,	C_ResrvDesc
			ggoSpread.SpreadLock C_IssueMeth,	-1,	C_IssueMeth	
			ggoSpread.SpreadLock C_IssueMethDesc,	-1,	C_IssueMethDesc
			ggoSpread.SpreadLock C_ReqNo,		-1,	C_ReqNo
			ggoSpread.SpreadLock C_ReqSeqNo,	-1,	C_ReqSeqNo
			
			ggoSpread.SpreadLock frm1.vspdData2.MaxCols, -1, frm1.vspdData2.MaxCols

			ggoSpread.SSSetRequired	 C_CompntCd, -1
			ggoSpread.SSSetRequired  C_RqrdQty,	-1
			ggoSpread.SSSetRequired  C_RqrdDt, -1
			ggoSpread.SSSetRequired  C_MajorSLCd, -1

			.vspdData2.Redraw = True
		End If
		
		If pvSpdNo = "C" Then
			'--------------------------------
			'Grid 3
			'--------------------------------
			ggoSpread.Source = frm1.vspdData3

			.vspdData3.ReDraw = False

			ggoSpread.SSSetRequired	 C_CompntCd3, -1
			ggoSpread.SSSetRequired  C_RqrdQty3,	-1
			ggoSpread.SSSetRequired  C_RqrdDt3, -1
			ggoSpread.SSSetRequired  C_MajorSLCd3, -1

			.vspdData2.Redraw = True
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
		ggoSpread.SSSetProtected C_OprNo2,				pvStartRow, pvEndRow    
		ggoSpread.SSSetRequired	 C_CompntCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CompntNm,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Spec,				pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_RqrdQty,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Unit,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IssuedQty,			pvStartRow, pvEndRow    
		ggoSpread.SSSetRequired  C_RqrdDt,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrackingNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_MajorSLCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MajorSLNm,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ResrvStatus,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ResrvDesc,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IssueMeth,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IssueMethDesc,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ReqNo,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ReqSeqNo,			pvStartRow, pvEndRow
		.Redraw = True
    
    End With

End Sub

'========================== 2.2.6 InitSpreadComboBox()  ========================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobDesc

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc


End Sub

'========================== 2.2.7 InitData()  =============================================
'	Name : InitData()
'	Description : Combo Display
'==========================================================================================
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex
	
	With frm1.vspdData1
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_OrderStatus
			intIndex = .value
			.Col = C_OrderStatusDesc
			.value = intindex
			.Row = intRow
			.col = C_JobCd
			intIndex = .value
			.Col = C_JobDesc
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
		C_OprNo				= 1
		C_JobCd				= 2
		C_JobDesc			= 3
		C_WcCd				= 4
		C_WcNm				= 5
		C_PlanStartDt		= 6
		C_PlanEndDt			= 7
		C_OrderStatus		= 8
		C_OrderStatusDesc	= 9
		C_InsideFlag		= 10
		C_InsideFlagDesc	= 11
		C_MileStone			= 12
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Operation
		C_CompntCd			= 1
		C_CompntCdPopup		= 2
		C_CompntNm			= 3
		C_Spec				= 4
		C_RqrdQty			= 5
		C_Unit				= 6
		C_IssuedQty			= 7
		C_RqrdDt			= 8
		C_TrackingNo		= 9
		C_MajorSLCd			= 10
		C_MajorSLCdPopUp	= 11
		C_MajorSLNm			= 12
		C_ResrvStatus		= 13
		C_ResrvDesc			= 14
		C_IssueMeth			= 15
		C_IssueMethDesc		= 16
		C_ReqNo				= 17
		C_ReqSeqNo			= 18
		' Hidden
		C_PlantCd			= 19
		C_ProdtOrderNo		= 20
		C_WcCd2				= 21
		C_OprNo2			= 22
		C_HndCompntCd		= 23
		C_HdnOprStatus		= 24
	End If
	
	If pvSpdNo = "*" Then
	' Grid 3(vspdData3) - Hidden
		C_CompntCd3			= 1
		C_CompntCdPopup3	= 2
		C_CompntNm3			= 3
		C_Spec3				= 4
		C_RqrdQty3			= 5
		C_Unit3				= 6
		C_IssuedQty3		= 7
		C_RqrdDt3			= 8
		C_TrackingNo3		= 9
		C_MajorSLCd3		= 10
		C_MajorSLCdPopUp3	= 11
		C_MajorSLNm3		= 12
		C_ResrvStatus3		= 13
		C_ResrvDesc3		= 14
		C_IssueMeth3		= 15
		C_IssueMethDesc3	= 16
		C_ReqNo3			= 17
		C_ReqSeqNo3			= 18
		C_PlantCd3			= 19
		C_ProdtOrderNo3		= 20
		C_WcCd3				= 21
		C_OprNo3			= 22
		C_HndCompntCd3		= 23
		C_HdnOprStatus3		= 24
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
			
			C_OprNo				= iCurColumnPos(1)
			C_JobCd				= iCurColumnPos(2)
			C_JobDesc			= iCurColumnPos(3)
			C_WcCd				= iCurColumnPos(4)
			C_WcNm				= iCurColumnPos(5)
			C_PlanStartDt		= iCurColumnPos(6)
			C_PlanEndDt			= iCurColumnPos(7)
			C_OrderStatus		= iCurColumnPos(8)
			C_OrderStatusDesc	= iCurColumnPos(9)
			C_InsideFlag		= iCurColumnPos(10)
			C_InsideFlagDesc	= iCurColumnPos(11)
			C_MileStone			= iCurColumnPos(12)
	
		Case "B"
 			ggoSpread.Source = frm1.vspdData2
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
 			C_CompntCd			= iCurColumnPos(1)
			C_CompntCdPopup		= iCurColumnPos(2)
			C_CompntNm			= iCurColumnPos(3)
			C_Spec				= iCurColumnPos(4)
			C_RqrdQty			= iCurColumnPos(5)
			C_Unit				= iCurColumnPos(6)
			C_IssuedQty			= iCurColumnPos(7)
			C_RqrdDt			= iCurColumnPos(8)
			C_TrackingNo		= iCurColumnPos(9)
			C_MajorSLCd			= iCurColumnPos(10)
			C_MajorSLCdPopUp	= iCurColumnPos(11)
			C_MajorSLNm			= iCurColumnPos(12)
			C_ResrvStatus		= iCurColumnPos(13)
			C_ResrvDesc			= iCurColumnPos(14)
			C_IssueMeth			= iCurColumnPos(15)
			C_IssueMethDesc		= iCurColumnPos(16)
			C_ReqNo				= iCurColumnPos(17)
			C_ReqSeqNo			= iCurColumnPos(18)
			C_PlantCd			= iCurColumnPos(19)
			C_ProdtOrderNo		= iCurColumnPos(20)
			C_WcCd2				= iCurColumnPos(21)
			C_OprNo2			= iCurColumnPos(22)
			C_HndCompntCd		= iCurColumnPos(23)
			C_HdnOprStatus		= iCurColumnPos(24)
 		
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

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"							' Field��(0)
    arrField(1) = "PLANT_NM"							' Field��(1)
    
    arrHeader(0) = "����"							' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
    
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
		Call DisplayMsgBox("971012","X", "����","X")
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
	arrParam(4) = "RLST"
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

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode, Byval strName, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 'ITEM_CD					' Field��(0)
	arrField(1) = 2 'ITEM_NM					' Field��(1)
	arrField(2) = 3	'SPECIFICATION
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet, Row)
	End If	
	
	Call SetFocusToDocument("M")

End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd(Byval strCode, Byval strName, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = strName													' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
   	arrField(0) = "SL_CD"													' Field��(0)
   	arrField(1) = "SL_NM"													' Field��(1)
   	arrHeader(0) = "â��"												' Header��(0)
   	arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSLCd(arrRet, Row)
	End If
	
End Function

'------------------------------------------  OpenStockRef()  -------------------------------------------
'	Name : OpenStockRef()
'	Description : Stock Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenStockRef()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
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

    frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	If frm1.vspdData2.Row < 1 Then 
	    IsOpenPop = False
	    Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4212RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4212RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	   
	arrParam(0) = Trim(UCase(frm1.txtPlantCd.value))		'��: ��ȸ ���� ����Ÿ 
	frm1.vspdData2.Col = C_CompntCd
	If frm1.vspdData2.Text = "" Then
	   IsOpenPop = False
	   Exit Function 
	End If
	arrParam(1) = Trim(UCase(frm1.vspdData2.Text))
	frm1.vspdData2.Col = C_CompntNm
	arrParam(2) = frm1.vspdData2.Text
	frm1.vspdData2.Col = C_MajorSLCd
	arrParam(3) = Trim(UCase(frm1.vspdData2.Text))
	frm1.vspdData2.Col = C_MajorSLNm
	arrParam(4) = frm1.vspdData2.Text
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet, Byval Row)

	Dim i

    With frm1.vspddata2

		For i = 1 to .MaxRows
			.Row = i
			.Col = C_CompntCd
			If .Text = arrRet(0) Then
				Call DisplayMsgBox("189504", "x", "x", "x")
				Exit Function
			End If
		Next
		
		.Row = Row
		.Col = C_CompntCd		
		.Text = arrRet(0)
		.Col = C_CompntNm
		.Text = arrRet(1)
		.Col = C_Spec
		.Text = arrRet(2)

		Call vspdData2_Change(C_CompntCd,  Row)

    End With

End Function

'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'	Name : LookUpItem ByPlant()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 

Function LookUpItemByPlant(Byval StrItemCd, Byval Row)
    
	Dim strVal
	Dim strSelect, strWhere
	Dim gComNum1000, gComNumDec, gAPNum1000, gAPNumDec
	
	gComNum1000 = parent.gComNum1000
	gComNumDec = parent.gComNumDec
	gAPNum1000 = parent.gAPNum1000
	gAPNumDec = parent.gAPNumDec

	If strItemCd = "" Then Exit Function
	
	frm1.vspdData2.Col = C_CompntCd
	frm1.vspdData2.Row = Row		
	
	strSelect = " A.ITEM_CD, A.BASIC_UNIT, A.ITEM_NM, A.SPEC, A.PHANTOM_FLG, B.VALID_FLG ITEM_VALID_FLG, B.PROCUR_TYPE,  "
	strSelect = strSelect & " B.VALID_FLG PLANT_VALID_FLG,   B.TRACKING_FLG, B.ORDER_UNIT_MFG, B.ORDER_LT_MFG,B.ISSUED_SL_CD, C.SL_NM, "
	strSelect = strSelect & " B.ISSUE_MTHD,   DBO.UFN_GETCODENAME( " & FilterVar("P1016", "''", "S") & " , B.ISSUE_MTHD ) AS  ISSUE_DESC  "
	
	strWhere = " A.ITEM_CD = B.ITEM_CD       AND B.ISSUED_SL_CD = C.SL_CD       AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	strWhere = strWhere & " AND B.ITEM_CD = " & FilterVar(Frm1.vspdData2.Text, "''", "S")
	
	If 	CommonQueryRs2by2(strSelect, " B_ITEM A (NOLOCK),    B_ITEM_BY_PLANT B (NOLOCK),  B_STORAGE_LOCATION C (NOLOCK) ", strWhere, lgF0) = False Then
		Call DisplayMsgBox("122700","X", Frm1.vspdData2.Text,"X")
		Call LookUpItemByPlantFail(Frm1.vspdData2.Text, Row)	    
		Exit Function
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	With frm1.vspdData2
		
		If lgF0(6) = "N"  Or lgF0(7) = "N" Then 'Invalid Item
			Call DisplayMsgBox("122619", "x", "x", "x") 
			Call LookUpItemByPlantFail(Frm1.vspdData2.Text, Row)
		Else
			If lgF0(5) = "Y" Then
				Call DisplayMsgBox("189214", "x", "x", "x")
			    Call LookUpItemByPlantFail(FilterVar(Frm1.vspdData2.Text, "''", "S"), Row)
			Else
				.Col = C_CompntNm
				.text = lgF0(3)
				.Col = C_Spec
				.text = lgF0(4)
				.Col = C_Unit
				.text = lgF0(2)

				If lgF0(10) = "N" Then 'TRACKING_FLG
					.Col = C_TrackingNo
					.Text = "*"
				Else
					.Col = C_TrackingNo		
					.Value = frm1.txtTrackingNo.Value
				End If

				.Col = C_MajorSLCd
				.text = lgF0(12)
				.Col = C_MajorSLNm
				.text = lgF0(13)
				.Col = C_IssueMeth
				.text = lgF0(14)
				.Col = C_IssueMethDesc
				.text = lgF0(15)    
			End If
		End If
	
	End With
	
	Call LookUpItemByPlantSuccess(Row)

End Function

Function LookUpItemByPlantFail(Byval strItemCd, Byval Row)

Dim	strOprNo

    With frm1.vspddata2
		.Row = Row
		.Col = C_CompntCd
		.text = ""
		.Col = C_CompntNm
		.text = ""
		.Col = C_Spec
		.text = ""
		.Col = C_Unit
		.text = ""
		.Col = C_TrackingNo
		.text = ""
		.Col = C_MajorSLCd
		.text = ""
		.Col = C_MajorSLNm
		.text = ""
		.Col = C_IssueMeth
		.text = ""
		.Col = C_IssueMethDesc
		.text = ""
		.Col = C_OprNo2
		strOprNo = .text
		
	End With
	
	Call DeleteHSheet(strOprNo, strItemCd)
	Call SetActiveCell(frm1.vspdData2, C_CompntCd, Row, "M","X","X")
	Set gActiveElement = document.activeElement
End Function

Function LookUpItemByPlantSuccess(Byval Row)
	
	Dim strCompntCd
	
	ggoSpread.Source = frm1.vspdData2
	frm1.vspdData2.Row = Row
	frm1.vspdData2.Col = C_CompntCd
	strCompntCd = frm1.vspdData2.Text

	ggoSpread.UpdateRow Row
	CopyToHSheet Row

	frm1.vspdData2.Col = C_HndCompntCd
	frm1.vspdData2.Text = strCompntCd
	
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet, Byval Row)

    With frm1.vspdData2
	   	.Row = Row
	   	.Col = C_MajorSLCD
	   	.Text = arrRet(0)
	   	.Col = C_MajorSLNM
	   	.Text = arrRet(1)
	End With

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	CopyToHSheet Row

End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'=======================================================================================================
'   Event Name : txtFromReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromReqDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFromReqDt_Change()	
	
End Sub
'=======================================================================================================
'   Event Name : txtToReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtToReqDt_Change()	
	
End Sub

'=======================================================================================================
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData1_onfocus()

End Sub

'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData2_onfocus()

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData1_Click(ByVal Col , ByVal Row )
    
    Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
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
			
		If DbDtlQuery = False Then		
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
			
			If DbDtlQuery = False Then	
	
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
		Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
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

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

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
	Dim	DblRqrdQty, DblIssuedQty
	Dim lNewRow, lOldRow

	lOldRow = frm1.vspdData1.ActiveRow
					
	With frm1.vspdData2

		Select Case Col

		    Case C_CompntCd

				.Row = Row
				.Col = C_CompntCd
				strItemCd = .Text
				
				If strItemCd = "" Then Exit Sub
				
				For i = 1 To .MaxRows
					If i <> Row Then
						.Row = i
						.Col = C_CompntCd
						If UCase(Trim(.Text)) = UCase(Trim(strItemCd)) Then
							Call DisplayMsgBox("189504", "x", "x", "x")
							.Row = Row
							.Text = ""
							Exit Sub
						End If
					End If						
				Next
				
				.Row = Row
				.Col = C_OprNo2
				strHndOprNo = .Text 				
				.Col = C_HndCompntCd
				strHndItemCd = .Text
			
				If strHndItemCd <> "" Then
					Call DeleteHSheet(strHndOprNo, strHndItemCd)
				End If

				.Row = Row
				.Col = C_HndCompntCd
				.Text = strItemCd
				
				Call LookUpItemByPlant(strItemCd, Row)

		    Case C_RqrdDt
				
				' �ʿ����� ������ �ϷΌ���� ���� �̷��� �� ����.
				.Row = Row
				.Col = C_RqrdDt
				strReqDt = .Text
				.Col = C_OprNo2
				strHndOprNo = .Text
				.Col = C_HndCompntCd
				strHndItemCd = .Text
				
				frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
				frm1.vspdData1.Col = C_PlanEndDt
				strEndDt = frm1.vspdData1.Text
			
				If UniConvDateAToB(strReqDt, parent.gDateFormat, parent.gServerDateFormat) > UniConvDateAToB(strEndDt, parent.gDateFormat, parent.gServerDateFormat) Then  
					Call DisplayMsgBox("189505", "x", "x", "x") '�ʿ����� ���԰����� �ϷΌ���Ϻ��� �̷��� �� �����ϴ�.
					lNewRow = frm1.vspdData1.ActiveRow
					If lNewRow <> lOldRow Then
						Call FixHiddenRow(strHndOprNo, strHndItemCd, C_RqrdDt, "") 
						Exit Sub
					Else
						.Row = Row
						If Row >= 1 Then
							.Col = C_RqrdDt
							.Text = ""
						Else
							Exit Sub
						End If
					End If
				End If

				.Col = C_CompntCd
			
				If .Text <> "" Then
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.UpdateRow Row
				End If
				
				CopyToHSheet Row

		    Case C_RqrdQty

				.Row = Row
				.Col = C_OprNo2
				strHndOprNo = .Text
				.Col = C_HndCompntCd
				strHndItemCd = .Text
				.Col = C_RqrdQty
				DblRqrdQty = .Text
				.Col = C_IssuedQty

				If UNICDbl(DblRqrdQty) < UNICDbl(.Text) Then  
					Call DisplayMsgBox("189521", "x", "x", "x")  '��ǰ �ʿ䷮�� ������� ���� ������ �� �����ϴ�.
					lNewRow = frm1.vspdData1.ActiveRow
					If lNewRow <> lOldRow Then
						Call FixHiddenRow(strHndOprNo, strHndItemCd, C_RqrdQty, "") 
						Exit Sub
					Else
						.Row = Row
						If Row >= 1 Then
							.Col = C_RqrdQty
							.Text = ""
						Else
							Exit Sub
						End If
					End If
				End If
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row

		    Case C_MajorSLCd

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				
				CopyToHSheet Row

		End Select

	End With

End Sub

'=======================================================================================================
'   Function Name : FixHiddenRow
'   Function Desc : 
'=======================================================================================================
Function FixHiddenRow(Byval strOprNo, Byval strItemCd, Byval Col, Byval strValue)

Dim strHndOprNo, strHndItemCd
Dim lRows

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows
        
            .vspdData3.Row = lRows
            .vspdData3.Col = C_OprNo3
            strHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_CompntCd3
            strHndItemCd = .vspdData3.Text

            If Trim(strHndOprNo) = Trim(strOprNo) And Trim(strHndItemCd) = Trim(strItemCd) Then
				.vspdData3.Col = Col
				.vspdData3.Text = strValue
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.UpdateRow lRows
				Exit Function
            End If    
        Next
        
    End With        
    
End Function

'========================================================================================
' Function Name : vspdData1_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
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
 
  
'========================================================================================
' Function Name : vspdData2_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
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

		    Case C_CompntCdPopup
				.Col = C_CompntCd
				.Row = Row
				strCode = .Text
				.Col = C_CompntNm
				.Row = Row
				strName = .Text
				Call OpenItemInfo(strCode, strName, Row)
				Call SetActiveCell(frm1.vspdData2, C_CompntCd, Row, "M","X","X")
				Set gActiveElement = document.activeElement
	    
		    Case C_MajorSLCdPopup
				.Col = C_MajorSLCd
				.Row = Row
				strCode = .Text
				.Col = C_MajorSLNm
				.Row = Row
				strName = .Text
				Call OpenSLCD(strCode, strName, Row)
				Call SetActiveCell(frm1.vspdData2, C_MajorSLCd, Row, "M","X","X")
				
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
    
    If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then

		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
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

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

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
	Dim strOprNo
	Dim strReqNo
	Dim strItemCd
	Dim strMode

	frm1.vspdData2.Col = C_OprNo2   
	frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	strOprNo = frm1.vspdData2.text

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    
    If gActiveSpdSheet.ID = "A" Then
		Call InitSpreadComboBox
		Call ggoSpread.ReOrderingSpreadData
    	Call InitData(1)
		
    ElseIf gActiveSpdSheet.ID = "B" Then
    
		ggoSpread.Source = frm1.vspdData3
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("C")
		ggoSpread.ReOrderingSpreadData()		

		ggoSpread.Source = frm1.vspdData2
		Call CopyFromHSheet(strOprNo)
	
		ggoSpread.SSSetProtected C_CompntCd,		1, frm1.vspdData2.MaxRows
		ggoSpread.SSSetProtected C_CompntCdPopup,	1, frm1.vspdData2.MaxRows

		With frm1.vspdData2
		For LngRow = 1 To .MaxRows
			.Row = LngRow
			.Col = C_HdnOprStatus
			If .Text = "CL" Then
				ggoSpread.SSSetProtected C_RqrdQty,			LngRow, LngRow
				ggoSpread.SSSetProtected C_RqrdDt,			LngRow, LngRow
				ggoSpread.SSSetProtected C_MajorSLCd,		LngRow, LngRow
				ggoSpread.SpreadLock C_MajorSLCdPopup,		LngRow, C_MajorSLCdPopup, LngRow
			Else
				ggoSpread.SSSetRequired C_RqrdQty,			LngRow, LngRow
				ggoSpread.SSSetRequired C_RqrdDt,			LngRow, LngRow
				ggoSpread.SSSetRequired C_MajorSLCd,		LngRow, LngRow
				ggoSpread.SpreadUnLock C_MajorSLCdPopup,	LngRow, C_MajorSLCdPopup, LngRow
			End If
		Next
		End With

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
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False											'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call InitVariables
	lgFlgQueryCnt = 0

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		 Call RestoreToolBar()
		 Exit Function												'��: Query db data		 
    End If   
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
    
    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
    
    With frm1.vspdData3
    
    For LngRows = 1 To .MaxRows
		.Row = LngRows
		.Col = C_RqrdQty3
		If .Value <= 0 Then
			Call DisplayMsgBox("189506", "x", "x", "x")
			Call GetHiddenFocus(LngRows, C_RqrdQty3)
			Exit Function
		End If    
    Next    
    
    End With
    '-----------------------
    'Save function call area
    '-----------------------
    
    If DbSave = False Then Exit Function													'��: Save db data
    
    FncSave = True												'��: Processing is OK
    
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
    SetSpreadColor frm1.vspdData2.ActiveRow
   
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 

Dim Row
Dim strMode
Dim	strOprNo
Dim	strItemCd
Dim strReqNo
Dim LngFindRow

	If frm1.vspdData2.MaxRows < 1 Then Exit Function	

    ggoSpread.Source = frm1.vspdData2	
    Row = frm1.vspdData2.ActiveRow
    frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
    frm1.vspdData2.Col = 0
    strMode = frm1.vspdData2.Text
    frm1.vspdData2.Col = C_OprNo2
    strOprNo = frm1.vspdData2.Text
    frm1.vspdData2.Col = C_CompntCd
    strItemCd = frm1.vspdData2.Text    
    frm1.vspdData2.Col = C_ReqNo
    strReqNo = frm1.vspdData2.Text

	If strMode = ggoSpread.InsertFlag Then
		Call DeleteHSheet(strOprNo, strItemCd)
	    ggoSpread.EditUndo                                                  '��: Protect system from crashing
	Else
		LngFindRow = FindRow(strOprNo, strReqNo)
		ggoSpread.Source = frm1.vspdData3
	    ggoSpread.EditUndo LngFindRow
	    
	    Call CopyOneRowFromHSheet(LngFindRow, Row)
	    
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
    
		.vspdData1.Row = .vspdData1.ActiveRow
		.vspdData1.Col = C_OrderStatus
    
		If .vspdData1.Text = "ST" Then
			Call DisplayMsgBox("189520", "x", "x", "x")
			Exit Function
		End IF    
   
		If .vspdData1.Text = "CL" Then
			Call DisplayMsgBox("189523", "x", "x", "x")
			Exit Function
		End IF    
		
		.vspdData2.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData2
		
		.vspdData2.ReDraw = False
		ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
   		
   		.vspdData1.Row = .vspdData1.ActiveRow
    	For pvRow = .vspdData2.ActiveRow To .vspdData2.ActiveRow + imRow -1
    		.vspdData2.Row = pvRow
			.vspdData2.Col = C_PlantCd
			.vspdData2.value = .txtPlantCd.value
			.vspdData2.Col = C_ProdtOrderNo
			.vspdData2.value = .txtProdOrderNo.value
			.vspdData1.Col = C_OprNo
			.vspdData2.Col = C_OprNo2
			.vspdData2.value = .vspdData1.value
			.vspdData1.Col = C_WcCd
			.vspdData2.Col = C_WcCd2
			.vspdData2.value = .vspdData1.value
		Next
		
		SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow -1
		.vspdData2.ReDraw = True
		
		Set gActiveElement = document.ActiveElement
	
		If Err.number = 0 Then FncInsertRow = True
		
    End With
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt, i

    With frm1

		.vspdData1.Row = frm1.vspdData1.ActiveRow
		.vspdData1.Col = C_OrderStatus
    
		If .vspdData1.Text = "ST" Then
			Call DisplayMsgBox("189520", "x", "x", "x")
			Exit Function
		End IF    
   
		If .vspdData1.Text = "CL" Then
			Call DisplayMsgBox("189523", "x", "x", "x")
			Exit Function
		End IF    
   
		If .vspdData2.MaxRows < 1 Then Exit Function

		Call DeleteMarkingHSheet()

    End With

	ggoSpread.Source = frm1.vspdData2
    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows

	CopyToHSheet frm1.vspdData2.ActiveRow

End Function

'=======================================================================================================
'   Function Name : DeleteMarkingHSheet
'   Function Desc : DeleteMark the Row Which keys match with vapdData's Key and vspdData2's Key
'=======================================================================================================
Function DeleteMarkingHSheet()

	Dim lRow, lRows
	
	Dim strInspItemCd
	Dim strInspSeries
	Dim strSampleNo
	Dim lngRow2
	Dim strHndOprNo, strOprNo, strHndItemCd, strItemCd	
	
	DeleteMarkingHSheet = False
	
	For lngRow2 = frm1.vspdData2.SelBlockRow To frm1.vspdData2.SelBlockRow2
	
        For lRows = 1 To frm1.vspdData3.MaxRows
            frm1.vspdData3.Row = lRows
            frm1.vspdData3.Col = C_OprNo3
            strHndOprNo = frm1.vspdData3.Text
            frm1.vspdData3.Col = C_CompntCd3
            strHndItemCd = frm1.vspdData3.Text
            frm1.vspdData2.Row = lngRow2
            frm1.vspdData2.Col = C_OprNo2
            strOprNo = frm1.vspdData2.Text
            frm1.vspdData2.Col = C_CompntCd
            strItemCd = frm1.vspdData2.Text
            If strHndOprNo = strOprNo And strHndItemCd = strItemCd Then
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
    Call parent.fncPrint()                                                   '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim strVal
    
    DbQuery = False
    
    lgFlgQueryCnt = lgFlgQueryCnt + 1
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing

    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdOrdNo=" & Trim(.hProdOrderNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	Else
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdOrdNo=" & Trim(.txtProdOrderNo.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Call InitData(LngMaxRow)

	frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1

	lgOldRow = 1
	
	If lgFlgQueryCnt = 1 Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
			Set gActiveElement = document.activeElement
			Call DbDtlQuery
		End If
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode

End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOprCd
	
	Call SetToolBar("11000000000111")

	boolExist = False
    With frm1

	    .vspdData1.Row = .vspdData1.ActiveRow
	    .vspdData1.Col = C_OprNo
	    strOprCd = .vspdData1.Text
    
	    If CopyFromHSheet(strOprCd) = True Then
	       Exit Function
        End If
		DbDtlQuery = False   
    
		.vspdData1.Row = .vspdData1.ActiveRow

		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdtOrderNo=" & Trim(.hProdOrderNo.Value)
			.vspdData1.Col = C_OprNo
			strVal = strVal & "&txtOprNo=" & Trim(.vspdData1.Text)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdtOrderNo=" & Trim(.txtProdOrderNo.Value)
			.vspdData1.Col = C_OprNo
			strVal = strVal & "&txtOprNo=" & Trim(.vspdData1.Text)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			
		End If
		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 
    End With
	
	
    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 
	
	Dim LngRow

	frm1.vspdData1.Col = C_InsideFlag
	If frm1.vspdData1.Text = "N" Then
		Call SetToolBar("11001001000111")										'��: ��ư ���� ���� 
	Else
		Call SetToolBar("11001111000111")										'��: ��ư ���� ���� 
	End IF

    '-----------------------
    'Reset variables area
    '-----------------------
	frm1.vspdData2.ReDraw = False
	
	ggoSpread.Source = frm1.vspdData2

	ggoSpread.SSSetProtected C_CompntCd,		1, frm1.vspdData2.MaxRows
	ggoSpread.SSSetProtected C_CompntCdPopup,	1, frm1.vspdData2.MaxRows

	With frm1.vspdData2

		For LngRow = 1 To .MaxRows
	
			.Row = LngRow
			.Col = C_HdnOprStatus
	
			If .Text = "CL" Then
				ggoSpread.SSSetProtected C_RqrdQty,			LngRow, LngRow
				ggoSpread.SSSetProtected C_RqrdDt,			LngRow, LngRow
				ggoSpread.SSSetProtected C_MajorSLCd,		LngRow, LngRow
				ggoSpread.SpreadLock C_MajorSLCdPopup,		LngRow, C_MajorSLCdPopup, LngRow
			Else
				ggoSpread.SSSetRequired C_RqrdQty,			LngRow, LngRow
				ggoSpread.SSSetRequired C_RqrdDt,			LngRow, LngRow
				ggoSpread.SSSetRequired C_MajorSLCd,		LngRow, LngRow
				ggoSpread.SpreadUnLock C_MajorSLCdPopup,	LngRow, C_MajorSLCdPopup, LngRow
			End If

		Next
	
	End With
   
	lgAfterQryFlg = True

	frm1.vspdData2.ReDraw = True

End Function

'============================================
'When No detailqryData
'===========================================
Function DbDtlQueryNotOk(ByVal LngMaxRow)	
	frm1.vspdData1.Col = C_InsideFlag
	If frm1.vspdData1.Text = "N" Then
		Call SetToolBar("11001001000111")										'��: ��ư ���� ���� 
	Else
		Call SetToolBar("11001101000111")										'��: ��ư ���� ���� 
	End IF
	
End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData(ByVal Row)

Dim strOprNo, strItemCd, strSeqNo
Dim strHndOprNo, strHndItemCd, strHndSeqNo
Dim lRows

    FindData = 0

    With frm1
        
        .vspdData2.Row = Row
        .vspdData2.Col = C_OprNo2
        strOprNo = .vspdData2.Text
        .vspdData2.Col = C_CompntCd
        strItemCd = .vspdData2.Text
        .vspdData2.Col = C_ReqSeqNo
        strSeqNo = .vspdData2.Text
        
        For lRows = 1 To .vspdData3.MaxRows
        
			.vspdData3.Row = lRows
			.vspdData3.Col = C_OprNo3
			strHndOprNo = .vspdData3.Text
			.vspdData3.Col = C_CompntCd3
			strHndItemCd = .vspdData3.Text
			.vspdData3.Col = C_ReqSeqNo3
			strHndSeqNo = .vspdData3.Text
			If Trim(strHndOprNo) = Trim(strOprNo) And Trim(strHndItemCd) = Trim(strItemCd) And Trim(strHndSeqNo) = Trim(strSeqNo) Then
            	FindData = lRows
				Exit Function
            End If    
        Next
        
    End With        
    
End Function


'=======================================================================================================
'   Function Name : FindRow
'   Function Desc : 
'=======================================================================================================
Function FindRow(ByVal strOprCd, ByVal strReqNo)

Dim lngRows
Dim strHdnOprCd
Dim strHdnReqNo

    FindRow = 0
    
    ggoSpread.Source = frm1.vspdData3
    
    With frm1
		'------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_OprNo3
            strHdnOprCd = .vspdData3.Text
            .vspdData3.Col = C_ReqNo3
            strHdnReqNo = .vspdData3.Text

            If strReqNo = "" Then
				If strOprCd = strHdnOprCd Then
				    FindRow = lngRows
				    Exit For
				End If
			Else
				If strOprCd = strHdnOprCd and strReqNo = strHdnReqNo Then
				    FindRow = lngRows
				    Exit For
				End If
			End If
        Next
            
    End With        
   
End Function

'=======================================================================================================
'   Function Name : CopyOneRowFromHSheet
'   Function Desc : 
'=======================================================================================================
Function CopyOneRowFromHSheet(ByVal LngSourceRow, ByVal LngTargetRow)

Dim iCols
Dim iCurColumnPos
Dim	strStatus

    CopyOneRowFromHSheet = False
    
    ggoSpread.Source = frm1.vspdData2
 			
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
    With frm1
		'------------------------------------
		' Show Data
		'------------------------------------ 
		.vspdData3.Row = LngSourceRow
		            
		frm1.vspdData2.Redraw = False
               
		.vspdData2.Row = LngTargetRow
		.vspdData2.Col = 0
		.vspdData3.Col = 0
		.vspdData2.Text = .vspdData3.Text
						
		For iCols = 1 To .vspdData3.MaxCols
		    .vspdData2.Col = iCurColumnPos(iCols)
		    .vspdData3.Col = iCols
		    .vspdData2.Text = .vspdData3.Text
		Next
						
		.vspdData3.Col = 0
		If .vspdData3.Text <> ggoSpread.InsertFlag Then 
			ggoSpread.SSSetProtected C_CompntCd,		LngTargetRow, LngTargetRow
			ggoSpread.SSSetProtected C_CompntCdPopup,	LngTargetRow, LngTargetRow
		End If
			
		.vspdData3.Col = C_HdnOprStatus3
		strStatus = .vspdData3.Text
		.vspdData3.Col = 0

		If strStatus = "CL" Then ' And .vspdData3.Text <> ggoSpread.InsertFlag
			ggoSpread.SSSetProtected C_RqrdQty3,		LngTargetRow, LngTargetRow
			ggoSpread.SSSetProtected C_RqrdDt3,			LngTargetRow, LngTargetRow
			ggoSpread.SSSetProtected C_MajorSLCd3,		LngTargetRow, LngTargetRow
			ggoSpread.SpreadLock C_MajorSLCdPopup3,		LngTargetRow, C_MajorSLCdPopup, LngTargetRow
		Else
			ggoSpread.SSSetRequired C_RqrdQty3,			LngTargetRow, LngTargetRow
			ggoSpread.SSSetRequired C_RqrdDt3,			LngTargetRow, LngTargetRow
			ggoSpread.SSSetRequired C_MajorSLCd3,		LngTargetRow, LngTargetRow
			ggoSpread.SpreadUnLock C_MajorSLCdPopup3,	LngTargetRow, C_MajorSLCdPopup, LngTargetRow
		End If

		frm1.vspdData2.Redraw = True

    End With        
    
    CopyOneRowFromHSheet = True
   
End Function

'=======================================================================================================
'   Function Name : CopyFromHSheet
'   Function Desc : 
'=======================================================================================================
Function CopyFromHSheet(ByVal strOprCd)

Dim lngRows
Dim boolExist
Dim iCols
Dim strHdnOprCd
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
        lngRows = FindRow(strOprCd, "")

        If lngRows > 0 Then
			boolExist = True
		End If    

		ggoSpread.Source = frm1.vspdData2

		'------------------------------------
		' Show Data
		'------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            frm1.vspdData2.Redraw = False
            
            While lngRows <= .vspdData3.MaxRows

	             .vspdData3.Row = lngRows
                
                .vspdData3.Col = C_OprNo3
				strHdnOprCd = .vspdData3.Text
                
                If strOprCd <> strHdnOprCd Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
               
					If strOprCd = strHdnOprCd Then
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
						If .vspdData3.Text <> ggoSpread.InsertFlag Then 
							ggoSpread.SSSetProtected C_CompntCd,		lngRows, lngRows
							ggoSpread.SSSetProtected C_CompntCdPopup,	lngRows, lngRows
						End If
			
						.vspdData3.Col = C_HdnOprStatus3
						strStatus = .vspdData3.Text
						.vspdData3.Col = 0

						If strStatus = "CL" Then ' And .vspdData3.Text <> ggoSpread.InsertFlag
							ggoSpread.SSSetProtected C_RqrdQty3,			lngRows, lngRows
							ggoSpread.SSSetProtected C_RqrdDt3,			lngRows, lngRows
							ggoSpread.SSSetProtected C_MajorSLCd3,		lngRows, lngRows
							ggoSpread.SpreadLock C_MajorSLCdPopup3,		lngRows, C_MajorSLCdPopup, lngRows
						Else
							ggoSpread.SSSetRequired C_RqrdQty3,			lngRows, lngRows
							ggoSpread.SSSetRequired C_RqrdDt3,			lngRows, lngRows
							ggoSpread.SSSetRequired C_MajorSLCd3,		lngRows, lngRows
							ggoSpread.SpreadUnLock C_MajorSLCdPopup3,	lngRows, C_MajorSLCdPopup, lngRows
						End If

					End If
                End If   
                
                lngRows = lngRows + 1
                
            Wend
            frm1.vspdData1.Col = C_InsideFlag
			If frm1.vspdData1.Text = "N" Then
				Call SetToolBar("11001001000111")										'��: ��ư ���� ���� 
			Else
				Call SetToolBar("11001111000111")										'��: ��ư ���� ���� 
			End IF
				
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
Dim iCurColumnPos

	ggoSpread.Source = frm1.vspdData2
 			
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
	With frm1 
        
	    lRow = FindData(Row)
	    
	    If lRow > 0 Then
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            For iCols = 1 To 20 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next
            
			.vspdData2.Col = C_CompntCd
			.vspdData3.Col = C_HndCompntCd3
			.vspdData3.Text = .vspdData2.Text
            
        Else
			.vspdData3.MaxRows = .vspdData3.MaxRows + 1
			
            .vspdData3.Row = .vspdData3.MaxRows
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
       
            For iCols = 1 To .vspdData2.MaxCols 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next

			.vspdData2.Col = C_CompntCd
			.vspdData3.Col = C_HndCompntCd3
			.vspdData3.Text = .vspdData2.Text
        
        End If

	End With
	
End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal strOprNo, Byval strItemCd)

Dim boolExist
Dim lngRows
Dim StrHndOprNo, strHndItemCd
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_OprNo3
			StrHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_CompntCd3
			strHndItemCd = .vspdData3.Text

            If strOprNo = StrHndOprNo and strItemCd = strHndItemCd Then
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
				.vspdData3.Col = C_OprNo3
				StrHndOprNo = .vspdData3.Text
				.vspdData3.Col = C_CompntCd3
				strHndItemCd = .vspdData3.Text
                
                If (strOprNo <> StrHndOprNo) or (strItemCd <> strHndItemCd) Then
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
        
        .vspdData3.SortKey(1) = C_OprNo3	' Operation No
        .vspdData3.SortKey(2) = C_CompntCd3	' Component Code
        
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
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen						'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

	Dim iTmpDBuffer							'������ ���� [����] 
	Dim iTmpDBufferCount					'������ ���� Position
	Dim iTmpDBufferMaxCount					'������ ���� Chunk Size

    DbSave = False                                                          '��: Processing is NG
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0

	With frm1.vspdData3

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0
			
			Select Case .Text
		    
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
			    	
			    	strVal = ""
			    	
			    	If .Text = ggoSpread.InsertFlag Then
						strVal = strVal & "CREATE" & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal = strVal & "UPDATE" & iColSep				'��: U=Update
			    	End If
			    		            
					strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
					.Col = C_ProdtOrderNo3	' Production Order No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					' Plan Order No.
					strVal = strVal & iColSep
					.Col = C_OprNo3			' Opr No.
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ReqSeqNo3		' Sequence
					strVal = strVal & Trim(.Text) & iColSep
					' Resvrd Status
					strVal = strVal & iColSep
					.Col = C_RqrdDt3		' Required Date
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					.Col = C_RqrdQty3		' Required Quantity
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
					.Col = C_ReqNo3			'  Required No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					' Resvrd Type
					strVal = strVal & iColSep
					.Col = C_TrackingNo3	' Tracking No.
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_Unit3			' Base Unit
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_IssueMeth3		' Issue Method
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_CompntCd3		' Child Item Cd
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_MajorSLCd3		'  Storage Location
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					.Col = C_WcCd3			'  Work Center
					strVal = strVal & UCase(Trim(.Text)) & iColSep
					'Row Count
					strVal = strVal & IntRows & parent.gRowSep
					

			    Case ggoSpread.DeleteFlag
					
					strDel = ""
					strDel = strDel & "DELETE" & iColSep				'��: D=Delete
					strDel = strDel & Trim(frm1.txtPlantCd.value) & iColSep
					.Col = C_ProdtOrderNo3	' Production Order No.
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					' Plan Order No.
					strDel = strDel & iColSep
					.Col = C_OprNo3			' Opr No.
					strDel = strDel & Trim(.Text) & iColSep
					.Col = C_ReqSeqNo3		' Sequence
					strDel = strDel & Trim(.Text) & iColSep
					' Resvrd Status
					strDel = strDel & iColSep
					.Col = C_RqrdDt3		' Required Date
					strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep
					.Col = C_RqrdQty3		' Required Quantity
					strDel = strDel & UNIConvNum(Trim(.Text),0) & iColSep
					.Col = C_ReqNo3			'  Required No.
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					' Resvrd Type
					strDel = strDel & iColSep
					.Col = C_TrackingNo3	' Tracking No.
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_Unit3			' Base Unit
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_IssueMeth3		' Issue Method
					strDel = strDel & Trim(.Text) & iColSep
					.Col = C_CompntCd3		' Child Item Cd
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_MajorSLCd3		'  Storage Location
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					.Col = C_WcCd3			'  Work Center
					strDel = strDel & UCase(Trim(.Text)) & iColSep
					'Row Count
					strDel = strDel & IntRows & parent.gRowSep
					
			End Select
			
			.Col = 0
			Select Case .Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			    
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
			         
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
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

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select

			
	    Next
	    
	End With
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: ���� �����Ͻ� ASP �� ���� 

    DbSave = True                                                           ' ��: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0

	ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0
	ggoSpread.source = frm1.vspddata3
    frm1.vspdData3.MaxRows = 0
	
	Call RemovedivTextArea
	Call DbDtlQuery
	
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
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
'----------  Coding part  -------------------------------------------------------------
'==============================================================================
' Function : GetHiddenFocus
' Description : �����߻��� Hidden Spread Sheet�� ã�� SheetFocus�� ���� �Ѱ���.
'==============================================================================
Function GetHiddenFocus(lRow, lCol)
	Dim lRows1, lRows2						'Quantity of the Hidden Data Keys Referenced by FindData Function
	Dim strHdnOprNo, strHdnItemCd			'Variable of Hidden Keys
	Dim strOprNo, strItemCd					'Variable of Visible Sheet Keys		
	
	If Trim(lCol) = "" Then
		lCol = C_CompntCd					'If Value of Column is not passed, Assign Value of the First Column in Second Spread Sheet
	End If
	'Find Key Datas in Hidden Spread Sheet
	With frm1.vspdData3
		.Row = lRow
		.Col = C_OprNo3
		strHdnOprNo = Trim(.Text)
		.Col = C_CompntCd3
		strHdnItemCd = Trim(.Text)
	End With
	'Compare Key Datas to Visible Spread Sheets
	With frm1
		For lRows1 = 1 To .vspdData1.MaxRows
			.vspdData1.Row = lRows1
			.vspdData1.Col = C_OprNo
			If Trim(.vspdData1.Text) = strHdnOprNo Then
				.vspdData1.focus
				.vspdData1.Action = 0
				lgOldRow = lRows1			'�� If this line is omitted, program could not query Data When errors occur
				ggoSpread.Source = .vspdData2
				.vspdData2.MaxRows = 0
				If CopyFromHSheet(strHdnOprNo) = True Then
				    For lRows2 = 1 To .vspdData2.MaxRows
						.vspdData2.Row = lRows2
						.vspdData2.Col = C_OprNo2
						strOprNo = .vspdData2.Text
						.vspdData2.Col = C_CompntCd
						strItemCd = .vspdData2.Text
						'Find Key Datas in Second Sheet and then Focus the Cell 
						If Trim(strHdnOprNo) = Trim(strOprNo) And Trim(strHdnItemCd) = Trim(strItemCd) Then
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
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData2.focus
	frm1.vspdData2.Row = lRow
	frm1.vspdData2.Col = lCol
	frm1.vspdData2.Action = 0
	frm1.vspdData2.SelStart = 0
	frm1.vspdData2.SelLength = len(frm1.vspdData2.Text)
End Function
