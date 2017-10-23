
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p4311mb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 
'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p4311mb2.asp"								'��: �����Ͻ� ���� ASP�� 

Const BIZ_PGM_SAVE_ID	= "p4311mb3.asp"

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Operation 
Dim C_ProdtOrderNo		'= 1
Dim C_CompntCd			'= 2
Dim C_CompntNm			'= 3
Dim C_Spec				'= 4
Dim C_RqrdQty			'= 5
Dim C_Unit				'= 6
Dim C_IssuedQty			'= 7
Dim C_RemainQty			'= 8
Dim C_TTlIssueQty		'= 9
Dim C_RqrdDt			'= 10
Dim C_ResrvStatus		'= 11
Dim C_ResrvStatusDesc	'= 12
Dim C_MajorSLCd			'= 13
Dim C_MajorSLNm			'= 14
Dim C_TrackingNo		'= 15
Dim C_OprNo				'= 16
Dim C_WcCD				'= 17
Dim C_ReqSeqNo			'= 18
Dim C_ReqNo				'= 19
Dim C_ParentItemCd		'= 20
Dim C_ParentItemNm		'= 21
Dim C_ParentItemSpec	'= 22
Dim C_JobNm				'= 23

' Grid 2(vspdData2) - Operation
Dim C_BlockIndicator	'= 1
Dim C_SLCd				'= 2
Dim C_SLNm				'= 3
Dim C_AllTrackingNo		'= 4
Dim C_LotNo				'= 5
Dim C_LotSubNo			'= 6
Dim C_OnHandQty			'= 7
Dim C_IssueQty			'= 8
Dim C_StkOnInspQty		'= 9
Dim C_StkOnTrnsQty		'= 10

' Grid 3(vspdData3) - Hidden
Dim C_CompntCd3			'= 1		' Child Item Cd
Dim C_ReqSeqNo3			'= 2		' Reserve Seq
Dim C_BlockIndicator3	'= 3		' Block Indicator
Dim C_SLCd3				'= 4		' Sl Cd
Dim C_SLNm3				'= 5		' Sl Nm
Dim C_AllTrackingNo3	'= 6		' Tracking No.
Dim C_LotNo3			'= 7		' Lot No.
Dim C_LotSubNo3			'= 8		' Lot Sub No.
Dim C_OnHandQty3		'= 9		' Good On Hand Qty
Dim C_IssueQty3			'= 10	' Issue Qty
Dim C_StkOnInspQty3		'= 11	' Inspection Qty
Dim C_StkOnTrnsQty3		'= 12	' Trans Qty
Dim C_PlantCd3			'= 13	' Plant Cd
Dim C_ProdtOrderNo3		'= 14	' Prodt Order No.
Dim C_OprNo3			'= 15	' Opr No.
Dim C_ReqNo3			'= 16	' MRP Req No.
Dim C_Unit3				'= 17	' Unit
Dim C_ReportDt3			'= 18	' Report Dt
Dim C_WcCD3				'= 19	' Wc Cd


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount							    'Group View Size�� ������ ���� 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgSortKey1
Dim lgSortKey2

Dim lgIntPrevKey
Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgCurrRow
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6		     'For InitCombobox 

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
         
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""							'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
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
    frm1.txtReqStartDt.Text = StartDate
    frm1.txtReqEndDt.Text = EndDate
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
			ggoSpread.Spreadinit "V20051006", , Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols =  C_JobNm + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ    
			.MaxRows =  0
			
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 18
			ggoSpread.SSSetEdit		C_CompntCd,		"��ǰ", 18
			ggoSpread.SSSetEdit		C_CompntNm,		"��ǰ��", 20
			ggoSpread.SSSetEdit		C_Spec,			"��ǰ�԰�", 20
			ggoSpread.SSSetFloat	C_RqrdQty, 		"�ʿ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_Unit, 		"����", 7
			ggoSpread.SSSetDate 	C_RqrdDt, 		"�ʿ���", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit 	C_TrackingNo,	"Tracking No.", 25
			ggoSpread.SSSetFloat	C_IssuedQty,	"��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_RemainQty,	"����ܷ�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_TTlIssueQty,	"������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetEdit		C_MajorSLCd,	"���â��", 10
			ggoSpread.SSSetEdit		C_MajorSLNm,	"���â���", 12
			ggoSpread.SSSetEdit		C_OprNo,		"����", 6
			ggoSpread.SSSetEdit		C_WcCD,			"�۾���", 10
			ggoSpread.SSSetEdit		C_ReqSeqNo,		"����", 6
			ggoSpread.SSSetEdit		C_ReqNo,		"����", 6
			ggoSpread.SSSetEdit		C_ResrvStatus,	"������", 10
			ggoSpread.SSSetEdit		C_ResrvStatusDesc, "������", 10	
			ggoSpread.SSSetEdit		C_ParentItemCd, "ǰ��", 15
			ggoSpread.SSSetEdit		C_ParentItemNm, "ǰ���", 20
			ggoSpread.SSSetEdit		C_ParentItemSpec, "ǰ��԰�", 20
			ggoSpread.SSSetEdit		C_JobNm,		"�۾���", 10
			
			'Call ggoSpread.MakePairsColumn(C_CompntNm, C_Spec)
			'Call ggoSpread.MakePairsColumn(C_ParentItemNm, C_ParentItemSpec)
 			Call ggoSpread.SSSetColHidden( .MaxCols	, .MaxCols	, True)
 			Call ggoSpread.SSSetColHidden( C_ReqSeqNo , C_ReqNo , True)
 			Call ggoSpread.SSSetColHidden( C_ResrvStatus , C_ResrvStatus , True)
			
			ggoSpread.SSSetSplit2(2)							'frozen ����߰� 
			
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
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_StkOnTrnsQty +1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
					
			Call GetSpreadColumnPos("B")
			
			Call AppendNumberPlace("6", "3", "0")
			
			ggoSpread.SSSetEdit		C_BlockIndicator,	"Block", 8
			ggoSpread.SSSetEdit		C_SLCd,				"â��", 7
			ggoSpread.SSSetEdit		C_SLNm,				"â���", 10
			ggoSpread.SSSetEdit		C_AllTrackingNo,	"Tracking No.", 25
			ggoSpread.SSSetEdit		C_LotNo,			"Lot No.", 13
			ggoSpread.SSSetFloat	C_LotSubNo,			"����", 11, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetFloat	C_OnHandQty,		"��ǰ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_IssueQty,			"������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_StkOnInspQty,		"�˻��߼�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_StkOnTrnsQty,		"�̵��߼�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden( .MaxCols	, .MaxCols	, True)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true
			
		End With
	End If    
    
    If pvSpdNo = "*" or pvSpdNo = "C" Then
		'------------------------------------------
		' Grid 3 - Hidden Spread Setting
		'------------------------------------------
		ggoSpread.Source = frm1.vspdData3
		frm1.vspdData3.MaxRows = 0
		frm1.vspdData3.MaxCols = C_WcCD3 + 1
		ggoSpread.Spreadinit
		ggoSpread.SSSetEdit		C_CompntCd3,		"��ǰ", 18
		ggoSpread.SSSetEdit		C_ReqSeqNo3,		"����", 6
		ggoSpread.SSSetEdit		C_BlockIndicator3,	"Block", 8
		ggoSpread.SSSetEdit		C_SLCd3,			"â��", 7
		ggoSpread.SSSetEdit		C_SLNm3,			"â���", 10
		ggoSpread.SSSetEdit		C_AllTrackingNo3,	"Tracking No.", 25
		ggoSpread.SSSetEdit		C_LotNo3,			"Lot No.", 13
		ggoSpread.SSSetFloat	C_LotSubNo3,		"����", 11, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat	C_OnHandQty3,		"��ǰ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_StkOnInspQty3,	"�˻��߼�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_StkOnTrnsQty3,	"�̵��߼�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_IssueQty3,		"������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PlantCd3,			"����", 15
		ggoSpread.SSSetEdit		C_ProdtOrderNo3,	"����������ȣ", 18
		ggoSpread.SSSetEdit		C_OprNo3,			"����", 6
		ggoSpread.SSSetEdit		C_ReqNo3,			"����", 6
		ggoSpread.SSSetEdit 	C_Unit3, 			"����", 7
		ggoSpread.SSSetDate 	C_ReportDt3, 		"�ʿ���", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_WcCD3,			"�۾���", 10
		
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
			ggoSpread.SpreadLock C_BlockIndicator, -1, C_BlockIndicator
			ggoSpread.SpreadLock C_SLCd, -1, C_SLCd
			ggoSpread.SpreadLock C_SLNm, -1, C_SLNm
			ggoSpread.SpreadLock C_AllTrackingNo, -1, C_AllTrackingNo
			ggoSpread.SpreadLock C_LotNo, -1, C_LotNo
			ggoSpread.SpreadLock C_LotSubNo, -1, C_LotSubNo
			ggoSpread.SpreadLock C_OnHandQty, -1, C_OnHandQty
			ggoSpread.SpreadLock C_StkOnInspQty, -1, C_StkOnInspQty
			ggoSpread.SpreadLock C_StkOnTrnsQty, -1, C_StkOnTrnsQty
			'ggoSpread.SpreadLock -1, -1
			.vspdData2.ReDraw = True
		End If	
	
    End With

End Sub


'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Operation 
		C_ProdtOrderNo		= 1
		C_CompntCd			= 2
		C_CompntNm			= 3
		C_CompntNm			= 4
		C_RqrdQty			= 5
		C_Unit				= 6
		C_IssuedQty			= 7
		C_RemainQty			= 8
		C_TTlIssueQty		= 9
		C_RqrdDt			= 10
		C_ResrvStatus		= 11
		C_ResrvStatusDesc	= 12
		C_MajorSLCd			= 13
		C_MajorSLNm			= 14
		C_TrackingNo		= 15
		C_OprNo				= 16
		C_WcCD				= 17
		C_ReqSeqNo			= 18
		C_ReqNo				= 19
		C_ParentItemCd		= 20
		C_ParentItemNm		= 21
		C_ParentItemSpec	= 22
		C_JobNm				= 23
	End If	
		
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
	
		' Grid 2(vspdData2) - Operation
		C_BlockIndicator	= 1
		C_SLCd				= 2
		C_SLNm				= 3
		C_AllTrackingNo		= 4
		C_LotNo				= 5
		C_LotSubNo			= 6
		C_OnHandQty			= 7
		C_IssueQty			= 8
		C_StkOnInspQty		= 9
		C_StkOnTrnsQty		= 10
	End If
	
	If pvSpdNo = "*" Then
		' Grid 3(vspdData3) - Hidden
		C_CompntCd3			= 1		' Child Item Cd
		C_ReqSeqNo3			= 2		' Reserve Seq
		C_BlockIndicator3	= 3		' Block Indicator
		C_SLCd3				= 4		' Sl Cd
		C_SLNm3				= 5		' Sl Nm
		C_AllTrackingNo3	= 6		' Tracking No.
		C_LotNo3			= 7		' Lot No.
		C_LotSubNo3			= 8		' Lot Sub No.
		C_OnHandQty3		= 9		' Good On Hand Qty
		C_IssueQty3			= 10	' Issue Qty
		C_StkOnInspQty3		= 11	' Inspection Qty
		C_StkOnTrnsQty3		= 12	' Trans Qty
		C_PlantCd3			= 13	' Plant Cd
		C_ProdtOrderNo3		= 14	' Prodt Order No.
		C_OprNo3			= 15	' Opr No.
		C_ReqNo3			= 16	' MRP Req No.
		C_Unit3				= 17	' Unit
		C_ReportDt3			= 18	' Report Dt
		C_WcCD3				= 19	' Wc Cd
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
 
			C_ProdtOrderNo		= iCurColumnPos(1)
			C_CompntCd			= iCurColumnPos(2)
			C_CompntNm			= iCurColumnPos(3)
			C_Spec				= iCurColumnPos(4)
			C_RqrdQty			= iCurColumnPos(5)
			C_Unit				= iCurColumnPos(6)
			C_IssuedQty			= iCurColumnPos(7)
			C_RemainQty			= iCurColumnPos(8)
			C_TTlIssueQty		= iCurColumnPos(9)
			C_RqrdDt			= iCurColumnPos(10)
			C_ResrvStatus		= iCurColumnPos(11)
			C_ResrvStatusDesc	= iCurColumnPos(12)
			C_MajorSLCd			= iCurColumnPos(13)
			C_MajorSLNm			= iCurColumnPos(14)
			C_TrackingNo		= iCurColumnPos(15)
			C_OprNo				= iCurColumnPos(16)
			C_WcCD				= iCurColumnPos(17)
			C_ReqSeqNo			= iCurColumnPos(18)
			C_ReqNo				= iCurColumnPos(19)
			C_ParentItemCd		= iCurColumnPos(20)
			C_ParentItemNm		= iCurColumnPos(21)
			C_ParentItemSpec	= iCurColumnPos(22)
			C_JobNm				= iCurColumnPos(23)
			
		Case "B"
 			ggoSpread.Source = frm1.vspdData2
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
 			C_BlockIndicator	= iCurColumnPos(1)
			C_SLCd				= iCurColumnPos(2)
			C_SLNm				= iCurColumnPos(3)
			C_AllTrackingNo		= iCurColumnPos(4)
			C_LotNo				= iCurColumnPos(5)
			C_LotSubNo			= iCurColumnPos(6)
			C_OnHandQty			= iCurColumnPos(7)
			C_IssueQty			= iCurColumnPos(8)
			C_StkOnInspQty		= iCurColumnPos(9)
			C_StkOnTrnsQty		= iCurColumnPos(10)
 			
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

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
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
	
    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
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
	Frm1.txtProdOrderNo.Focus
	
End Function

'------------------------------------------  OpenItemInfo1()  -------------------------------------------------
'	Name : OpenItemInfo1()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo1(Byval strCode)
	
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call Displaymsgbox("971012","X", "����","X")
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
	arrParam(1) = strCode			' Item Code
	arrParam(2) = ""				' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field��(0)
	arrField(1) = 2 '"ITEM_NM"			' Field��(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo1(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtChildItemCd.focus

End Function

'------------------------------------------  OpenItemInfo2()  -------------------------------------------------
'	Name : OpenItemInfo2()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo2(Byval strCode)
	
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call Displaymsgbox("971012","X", "����","X")
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
	arrParam(1) = strCode			' Item Code
	arrParam(2) = ""				' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field��(0)
	arrField(1) = 2 '"ITEM_NM"			' Field��(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo2(arrRet)
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

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd(Byval strCode)

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

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 	' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
   	arrField(0) = "SL_CD"													' Field��(0)
   	arrField(1) = "SL_NM"													' Field��(1)
   	arrHeader(0) = "â��"												' Header��(0)
   	arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtSLCd.focus
	
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
		Call DisplayMsgBox("971012","X", "����","X")
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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") ' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingInfo(Byval strCode)
	
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
	arrParam(2) = ""
	arrParam(3) = frm1.txtReqStartDt.Text
	arrParam(4) = frm1.txtReqEndDt.Text	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
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
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If
	
    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = C_ProdtOrderNo
	arrParam(1) = Trim(frm1.vspdData1.Text)		'��: ��ȸ ���� ����Ÿ 

	IsOpenPop = True	

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

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
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = C_ProdtOrderNo
	arrParam(1) = Trim(frm1.vspdData1.Text)		'��: ��ȸ ���� ����Ÿ 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

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
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4411RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4411RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = C_ProdtOrderNo
	arrParam(1) = Trim(frm1.vspdData1.Text)		'��: ��ȸ ���� ����Ÿ 
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetItemInfo1()  --------------------------------------------------
'	Name : SetItemInfo1()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo1(Byval arrRet)
    With frm1
	.txtChildItemCd.value = arrRet(0)
	.txtChildItemNm.value = arrRet(1)
    End With
End Function

'------------------------------------------  SetItemInfo1()  --------------------------------------------------
'	Name : SetItemInfo2()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo2(Byval arrRet)
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

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet)

    With frm1
        .txtSLCd.value = arrRet(0)  
	   	.txtSLNm.Value = arrRet(1)	   	
	End With

End Function

'------------------------------------------  SetConWC()  --------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	
	frm1.txtTrackingNo.Value = arrRet(0)
	
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
'------------------------------------------  ReqStartDt_KeyDown ----------------------------------------
'	Name : txtReqStartDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtReqEndDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtIssueDt.Focus
    End If
End Sub

Sub txtReqStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqStartDt.Focus
    End If
End Sub

Sub txtReqEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqEndDt.Focus
    End If
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
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
		Exit Sub
	End If
	
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
		
		Call DisableToolBar(parent.TBC_QUERY)
		
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
			
			Call DisableToolBar(parent.TBC_QUERY)
			
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
		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
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
	
	If SetTotalIssueQtyToGrid(Col, Row) = False Then
		Exit Sub
	End If
    ggoSpread.Source = frm1.vspdData2
	frm1.vspdData2.Col = C_IssueQty
	frm1.vspdData2.Row = Row
	If UNICDbl(frm1.vspdData2.Text) > 0 Then
		ggoSpread.UpdateRow Row
	Else
		ggoSpread.SSDeleteFlag Row,Row
	End If

	CopyToHSheet Row

End Sub

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
		If lgStrPrevKey <> "" And lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    Exit Sub
    
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
	Dim strMode, strProdtOrderNo, strItemCd, strSeq, strOprNo
	Dim lngHdnRow
	
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    
    select case gActiveSpdSheet.id
		case "A"
		Call ggoSpread.ReOrderingSpreadData()
	
	    case "B"
	    with frm1
			ggoSpread.Source = .vspdData3
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("C")
			ggoSpread.ReOrderingSpreadData()

	    	.vspdData1.Row = .vspdData1.ActiveRow
			.vspdData1.Col = C_ProdtOrderNo
			strProdtOrderNo = .vspdData1.Text
			.vspdData1.Col = C_CompntCd
			strItemCd = .vspdData1.Text
			.vspdData1.Col = C_ReqSeqNo
			strSeq = .vspdData1.Text
			.vspdData1.Col = C_OprNo
			strOprNo = .vspdData1.Text
			Call CopyFromHSheet(.vspdData1.Row, strProdtOrderNo,strItemCd,strSeq,strOprNo)
		end with
    End select
    
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
    Dim StrIssueDt
    
    FncQuery = False											'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = displaymsgbox("900013", parent.VB_YES_NO, "x", "x")	'��: Display Message
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
    
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtChildItemCd.value = "" Then
		frm1.txtChildItemNm.value = "" 
	End If
	If frm1.txtSLCd.value = "" Then
		frm1.txtSLNm.value = "" 
	End If
    
    If ValidDateCheck(frm1.txtReqStartDt, frm1.txtReqEndDt) = False Then 	Exit Function
    
	StrIssueDt = frm1.txtIssueDt.Text 
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
    frm1.txtIssueDt.Text = StrIssueDt
   
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
	End If												'��: Query db data
       
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
	Dim i, LngTtlIssQty, LngRemainQty
	    
    FncSave = False												'��: Processing is NG
    
    Err.Clear                                                   '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = displaymsgbox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If

    If Not chkfield(Document, "2") Then					'��: Check required field(Single area)
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then				'��: Check required field(Multi area)
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
	For i = 1 To frm1.vspdData1.MaxRows
		frm1.vspdData1.Row = i
		frm1.vspdData1.Col = C_RemainQty
		LngRemainQty = frm1.vspdData1.Text
		frm1.vspdData1.Col = C_TTlIssueQty
		LngTtlIssQty = frm1.vspdData1.Text
		
		If UNICDbl(LngRemainQty) < UNICDbl(LngTtlIssQty) Then
			frm1.vspdData1.focus
			frm1.vspdData1.Row = i
			frm1.vspdData1.Col = C_TTlIssueQty
			frm1.vspdData1.Action = 0
			frm1.vspdData1.SelStart = 0
			frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
			Call vspdData1_Click(C_TTlIssueQty, i)
			Call displaymsgbox("189515", "x", "x", "x")
			Call SetToolBar("11001000000111")										'��: ��ư ���� ���� 
			Exit Function
		End If
    Next
    
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
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	Dim lngCnclRow
	Dim lngRows
	Dim strItemCd, strSeq, strSlCd, strTrackingNo,strLotNo, strLotSubNo, strProdtOrderNo,strOprNo
	Dim strHdnItemCd, strHdnSeq, strHdnSlCd, strHdnTrackingNo,strHdnLotNo, strHdnLotSubNo
	Dim strHdnProdtOrderNo,strHdnOprNo
	
	If frm1.vspdData2.MaxRows < 1 Then Exit Function	
	
	ggoSpread.Source = frm1.vspdData1	
    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = C_CompntCd
    strItemCd = frm1.vspdData1.Text
	frm1.vspdData1.Col = C_ReqSeqNo
    strSeq = frm1.vspdData1.Text
    frm1.vspdData1.Col = C_ProdtOrderNo
    strProdtOrderNo =frm1.vspdData1.Text
	frm1.vspdData1.Col = C_OprNo
    strOprNo =frm1.vspdData1.Text
	
	ggoSpread.Source = frm1.vspdData2	
    lngCnclRow = frm1.vspdData2.ActiveRow
    frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	frm1.vspdData2.Col = C_SLCd
    strSlCd = frm1.vspdData2.Text
	frm1.vspdData2.Col = C_AllTrackingNo
    strTrackingNo = frm1.vspdData2.Text
	frm1.vspdData2.Col = C_LotNo
	strLotNo = frm1.vspdData2.Text
	frm1.vspdData2.Col = C_LotSubNo
    strLotSubNo = frm1.vspdData2.Text
	
	'------------------------------------
	' Find Row
	'------------------------------------ 
	For lngRows = 1 To frm1.vspdData3.MaxRows
		
		With frm1.vspdData3
			.Row = lngRows
			.Col = C_CompntCd3
			strHdnItemCd = Trim(.Text)
			.Col = C_ReqSeqNo3
			strHdnSeq = Trim(.Text)
			.Col = C_SLCd3
			strHdnSlCd = Trim(.Text)
			.Col = C_AllTrackingNo3
			strHdnTrackingNo = Trim(.Text)
			.Col = C_LotNo3
			strHdnLotNo = Trim(.Text)
			.Col = C_LotSubNo3
			strHdnLotSubNo = Trim(.Text)
			.Col = C_ProdtOrderNo3
			strHdnProdtOrderNo = Trim(.Text)
			.Col = C_OprNo3
			strHdnOprNo = Trim(.Text)
		End With
	    If strItemCd = strHdnItemCd and strSeq = strHdnSeq and strSlCd = strHdnSlCd _
			and strTrackingNo = strHdnTrackingNo and strLotNo = strHdnLotNo _
			and strLotSubNo = strHdnLotSubNo and strProdtOrderNo = strHdnProdtOrderNo _
			and strOprNo = strHdnOprNo Then
	        Exit For
	    End If
	Next
	
	ggoSpread.Source = frm1.vspdData3
	
	ggoSpread.EditUndo lngRows
	Call CopyOneRowFromHSheet(lngRows, lngCnclRow)
	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 
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
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��: Protect system from crashing
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
	
    ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = displaymsgbox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
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
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing

    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1	    
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2	    
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtReqStartDt=" & Trim(.hReqStartDt.Value)
		strVal = strVal & "&txtReqEndDt=" & Trim(.hReqEndDt.Value)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(.hProdOrderNo.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtChildItemCd=" & Trim(.hChildItemCd.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtSLCd=" & Trim(.hSLCd.Value)					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboProdMgr=" & Trim(.hProdMgr.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboInvMgr=" & Trim(.hInvMgr.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboJobCd=" & Trim(.hJobCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	Else
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1	    
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2	    
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)			'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&txtReqStartDt=" & Trim(.txtReqStartDt.Text)
		strVal = strVal & "&txtReqEndDt=" & Trim(.txtReqEndDt.Text)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(.txtProdOrderNo.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtChildItemCd=" & Trim(.txtChildItemCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtSLCd=" & Trim(.txtSLCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboProdMgr=" & Trim(.cboProdMgr.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboInvMgr=" & Trim(.cboInvMgr.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboJobCd=" & Trim(.cboJobCd.Value)	'��: ��ȸ ���� ����Ÿ 
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

	Call SetToolBar("11001001000111")										'��: ��ư ���� ���� 
	
	Call ggoOper.LockField(Document, "N")										'��: It's not Standard, This function lock the suitable field
    Call ggoOper.SetReqAttr(frm1.txtItemDocumentNo,"D")
    frm1.txtIssueDt.Text = LocSvrDate
 
	frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1

	lgOldRow = 1

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisableToolBar(parent.TBC_QUERY)
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		If DbDtlQuery(frm1.vspdData1.Row) = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE											'��: Indicates that current mode is Update mode

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
Dim strItemCd
Dim strSeq
Dim strOprNo
   
	boolExist = False
    With frm1

	    .vspdData1.Row = LngRow
	    .vspdData1.Col = C_ProdtOrderNo
	    strProdtOrderNo = .vspdData1.Text
	    .vspdData1.Col = C_CompntCd
	    strItemCd = .vspdData1.Text
	    .vspdData1.Col = C_ReqSeqNo
	    strSeq = .vspdData1.Text
	    .vspdData1.Col = C_OprNo
	    strOprNo = .vspdData1.Text
    
	    If CopyFromHSheet(LngRow, strProdtOrderNo,strItemCd,strSeq,strOprNo) = True Then
           Call RestoreToolBar
           Exit Function
        End If

		DbDtlQuery = False   
    
		.vspdData1.Row = LngRow

		Call LayerShowHide(1)       

		If lgIntFlgMode <> parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			.vspdData1.Col = C_CompntCd
			strVal = strVal & "&txtChildItemCd=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_MajorSLCd
			strVal = strVal & "&txtMajorSlCd=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_TrackingNo
			strVal = strVal & "&txtTrackingNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_ReqSeqNo
			strVal = strVal & "&txtReqSeqNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_ProdtOrderNo
			strVal = strVal & "&txtProdtOrderNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_OprNo
			strVal = strVal & "&txtOprNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_ReqNo
			strVal = strVal & "&txtMRPReqNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_Unit
			strVal = strVal & "&txtUnit=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_WcCD
			strVal = strVal & "&txtWcCd=" & Trim(.vspdData1.Text)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			.vspdData1.Col = C_CompntCd
			strVal = strVal & "&txtChildItemCd=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_MajorSLCd
			strVal = strVal & "&txtMajorSlCd=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_TrackingNo
			strVal = strVal & "&txtTrackingNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_ReqSeqNo
			strVal = strVal & "&txtReqSeqNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_ProdtOrderNo
			strVal = strVal & "&txtProdtOrderNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_OprNo
			strVal = strVal & "&txtOprNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_ReqNo
			strVal = strVal & "&txtMRPReqNo=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_Unit
			strVal = strVal & "&txtUnit=" & Trim(.vspdData1.Text)
			.vspdData1.Col = C_WcCD
			strVal = strVal & "&txtWcCd=" & Trim(.vspdData1.Text)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    End With

    DbDtlQuery = True

End Function

'========================================================================================
' Function Name : DbDtlQueryOk
' Function Desc : This function is detail data query and display
'========================================================================================
Function DbDtlQueryOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 

	Dim i

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	lgAfterQryFlg = True
	
	frm1.vspdData1.ReDraw = False
	frm1.vspdData2.ReDraw = False			
	
	 '-------------------------------------------------------------------------------------------------
    '     Set Property of VspdData2 
	'-------------------------------------------------------------------------------------------------	
	
	With frm1
		.vspdData1.Row = .vspdData1.ActiveRow 
		.vspdData1.Col = C_ResrvStatus
		If  .vspdData1.Text = "CL" Or .vspdData1.Text = "OP" Or .vspdData1.Text = "PL" Then
			
			ggoSpread.Source = frm1.vspdData2
			For i = 1 To .vspdData2.MaxRows
				ggoSpread.SSSetProtected C_IssueQty,	i, i
			Next 
			
		Else
		
			For i = 1 To .vspdData2.MaxRows
				.vspdData2.Row = i
		    
				.vspdData2.Col = C_BlockIndicator  
					IF UCase(Trim(.vspdData2.Text)) = "Y" Then
						ggoSpread.SSSetProtected C_IssueQty, i, i
					End If
					
				.vspdData1.Col = C_MajorSLCd
				.vspdData2.Col = C_SLCd
				    IF UCase(Trim(.vspdData1.Text)) <> UCase(Trim(.vspdData2.Text)) Then
				        ggoSpread.SSSetProtected C_IssueQty, i, i
				    End If
		    
				.vspdData1.Col = C_TrackingNo
				.vspdData2.Col = C_AllTrackingNo
					IF UCase(Trim(.vspdData1.Text)) <> UCase(Trim(.vspdData2.Text)) Then
				        ggoSpread.SSSetProtected C_IssueQty, i, i
				    End If
			Next		
		End If
	End With	
    
    Call RestoreToolBar()
     
	frm1.vspdData1.ReDraw = True
	frm1.vspdData2.ReDraw = True	

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData(Byval Row)
Dim strApNo
Dim strItemCd
Dim strSeq
Dim strSlCd
Dim strTrackingNo
Dim strLotNo
Dim strLotSubNo
Dim strProdtOrderNo
Dim strOprNo

Dim lRows

    FindData = 0

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows
        
            .vspdData3.Row = lRows
            .vspdData3.Col = 1
            strItemCd = .vspdData3.Text
            .vspdData3.Col = 2
            strSeq = .vspdData3.Text
            .vspdData3.Col = 4
            strSlCd = .vspdData3.Text
            .vspdData3.Col = 6
            strTrackingNo = .vspdData3.Text
            .vspdData3.Col = 7
            strLotNo = .vspdData3.Text
            .vspdData3.Col = 8
            strLotSubNo = .vspdData3.Text
            .vspdData3.Col = 14
            strProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = 15
            strOprNo = .vspdData3.Text
           
            .vspdData1.Row = frm1.vspdData1.ActiveRow
            .vspdData2.Row = Row
            .vspdData1.Col = C_CompntCd
           
            If strItemCd = .vspdData1.Text Then
                .vspdData1.Col = C_ReqSeqNo
                If strSeq = .vspdData1.Text Then
					.vspdData2.Col = C_SLCd
					If strSlCd = .vspdData2.Text Then
						.vspdData2.Col = C_AllTrackingNo
						If strTrackingNo = .vspdData2.Text Then
							.vspdData2.Col = C_LotNo
							If strLotNo = .vspdData2.Text Then
								.vspdData2.Col = C_LotSubNo
								If strLotSubNo = .vspdData2.Text Then
									.vspdData1.Col = C_ProdtOrderNo
									If strProdtOrderNo = .vspdData1.Text Then
										.vspdData1.Col = C_OprNo
										If strOprNo = .vspdData1.Text Then
											FindData = lRows
											Exit Function
										End If	
									End If
								End If
							End If
						End If
					End If
                End If
            End If    
        Next
        
    End With        
    
End Function


'=======================================================================================================
'   Function Name : CopyFromHSheet
'   Function Desc : 
'=======================================================================================================
Function CopyFromHSheet(ByVal LngRow ,ByRef strProdtOrderNo, ByRef strItemCd, ByRef strSeq, ByRef strOprNo)
Dim lngRows
Dim boolExist
Dim iCols
Dim strHdnProdtOrderNo
Dim strHdnItemCd
Dim strHdnSeq
Dim strHdnOprNo
Dim strIssueMthd
Dim iCurColumnPos

Dim i

    boolExist = False
    
    CopyFromHSheet = boolExist
    
    With frm1

        Call SortHSheet()
                        
        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_CompntCd3
            strHdnItemCd = .vspdData3.Text
            .vspdData3.Col = C_ReqSeqNo3
            strHdnSeq = .vspdData3.Text
            .vspdData3.col = C_ProdtOrderNo3
            strHdnProdtOrderNo = .vspdData3.Text
            .vspdData3.col = C_OprNo3
            strHdnOprNo = .vspdData3.Text
              
            If strItemCd = strHdnItemCd and strSeq = strHdnSeq and strProdtOrderNo = strHdnProdtOrderNo and strOprNo = strHdnOprNo Then
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
            
            ggoSpread.Source = frm1.vspdData2 
 		
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            While lngRows <= .vspdData3.MaxRows

	             .vspdData3.Row = lngRows
                
                .vspdData3.Col = C_CompntCd3
				strHdnItemCd = .vspdData3.Text
                .vspdData3.Col = C_ReqSeqNo3
				strHdnSeq = .vspdData3.Text
				.vspdData3.Col = C_ProdtOrderNo3
				strHdnProdtOrderNo = .vspdData3.Text
				.vspdData3.col = C_OprNo3
                strHdnOprNo = .vspdData3.Text
                
                If strItemCd <> strHdnItemCd and strSeq <> strHdnSeq and strProdtOrderNo <> strHdnProdtOrderNo and strOprNo <> strHdnOprNo Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
					If strItemCd = strHdnItemCd and strSeq = strHdnSeq and strProdtOrderNo = strHdnProdtOrderNo and strOprNo = strHdnOprNo Then
						.vspdData2.MaxRows = .vspdData2.MaxRows + 1
						.vspdData2.Row = .vspdData2.MaxRows
						.vspdData2.Col = 0
						.vspdData3.Col = 0
						.vspdData2.Text = .vspdData3.Text
                  
						For iCols = 1 To 10						'vspdData2 �κи� Copy
						    .vspdData2.Col = iCurColumnPos(iCols)
						    .vspdData3.Col = iCols + 2
						    .vspdData2.Text = .vspdData3.Text
						Next
					End If
                End If   
                .vspdData3.Col = .vspdData3.MaxCols
                .vspdData2.Col = .vspdData2.MaxCols
                .vspdData2.Text = CInt(.vspdData3.Text) - 1 
                
                lngRows = lngRows + 1
                
            Wend
           
           '-------------------------------------------------------------------------------------------------
           '     Set Property of VspdData2 
		   '-------------------------------------------------------------------------------------------------
			.vspdData1.Row = LngRow
			.vspdData1.Col = C_ResrvStatus
			If  .vspdData1.Text = "CL" Or .vspdData1.Text = "OP" Or .vspdData1.Text = "PL" Then
					
				ggoSpread.Source = frm1.vspdData2
				For i = 1 To .vspdData2.MaxRows
					ggoSpread.SSSetProtected C_IssueQty,	i, i
				Next 
					
			Else
				
				For i = 1 To .vspdData2.MaxRows
					.vspdData2.Row = i
				    
					.vspdData2.Col = C_BlockIndicator  
						IF UCase(Trim(.vspdData2.Text)) = "Y" Then
							ggoSpread.SSSetProtected C_IssueQty, i, i
						End If
							
					.vspdData1.Col = C_MajorSLCd
					.vspdData2.Col = C_SLCd
					    IF UCase(Trim(.vspdData1.Text)) <> UCase(Trim(.vspdData2.Text)) Then
					        ggoSpread.SSSetProtected C_IssueQty, i, i
					    End If
				    
					.vspdData1.Col = C_TrackingNo
					.vspdData2.Col = C_AllTrackingNo
						IF UCase(Trim(.vspdData1.Text)) <> UCase(Trim(.vspdData2.Text)) Then
					        ggoSpread.SSSetProtected C_IssueQty, i, i
					    End If
				Next		
			End If
           
            frm1.vspdData2.Redraw = True

        End If
            
    End With        
    
    CopyFromHSheet = boolExist
    
End Function

'=======================================================================================================
'   Function Name : CopyOneRowFromHSheet
'   Function Desc : 
'=======================================================================================================
Function CopyOneRowFromHSheet(ByVal SourceRow, ByVal TargetRow)
	
	Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

    frm1.vspdData2.Redraw = False
    '------------------------------------
    ' Init IssueQty
    '------------------------------------ 
    frm1.vspdData3.Row = SourceRow
    frm1.vspdData2.Row = TargetRow
    frm1.vspdData3.Col = 0
    frm1.vspdData2.Col = 0
    frm1.vspdData2.Text = frm1.vspdData3.Text
    frm1.vspdData3.Col = C_IssueQty3
    frm1.vspdData2.Col = iCurColumnPos(8)
	frm1.vspdData2.Text = frm1.vspdData3.text
    
    frm1.vspdData2.Redraw = True
    
    If SetTotalIssueQtyToGrid(frm1.vspdData2.Col,TargetRow) = False Then
		Exit Function
    End If
    
End function

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
       
            For iCols = 1 To 10 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols + 2
                .vspdData3.Text = .vspdData2.Text
            Next
        End If

	End With
	
End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal strItemSeq)
Dim boolExist
Dim lngRows
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_CompntCd3                

            If strItemSeq = .vspdData3.Text Then
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
                .vspdData3.Col = C_CompntCd3
                
                If strItemSeq <> .vspdData3.Text Then
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
        .vspdData3.SortKey(2) = C_OprNo3
        .vspdData3.SortKey(3) = C_CompntCd3
        .vspdData3.SortKey(4) = C_ReqSeqNo3
        .vspdData3.SortKey(5) = C_SLCd3
        
        
        .vspdData3.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(3) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(4) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(5) = 1 'SS_SORT_ORDER_ASCENDING
        
        
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 0
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

	Dim iTmpDBuffer							'������ ���� [����] 
	Dim iTmpDBufferCount					'������ ���� Position
	Dim iTmpDBufferMaxCount					'������ ���� Chunk Size
	
    DbSave = False                                                          '��: Processing is NG
    
	iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			

	iTmpCUBufferCount = -1
	
	strCUTotalvalLen = 0
    
    Call LayerShowHide(1)
    
    With frm1
		.txtMode.value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
		
	End With
	
    '-----------------------
    'Data manipulate area
    '-----------------------
	With frm1.vspdData3
	    
		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col =  C_IssueQty3			'10	Issue Qty
			
			If UNICDbl(.Text) > 0 Then
				
				strVal = ""
				.Col = C_PlantCd3			'1	Plant Cd
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_ProdtOrderNo3			'2	Prodt Order No.
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_OprNo3			'3	Opr No.
				strVal = strVal & Trim(.Text) & iColSep
				If CompareDateByFormat(frm1.txtIssueDt.Text, LocSvrDate,"�����","������","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
					  Call LayerShowHide(0)
					   .EditMode = True
					   strVal = ""
					  Exit Function               
				End If 
				strVal = strVal & UNIConvDate(frm1.txtIssueDt.text) & iColSep
				.Col = C_CompntCd3			'5	Child Item Cd
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_SLCd3			'6	Sl Cd
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_AllTrackingNo3			'7	Tracking No.
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_LotNo3 			'8	Lot No.
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_LotSubNo3			'9	Lot Sub No.
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_IssueQty3			'10	Issue Qty
				strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
				.Col = C_Unit3			'11	Unit
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_WcCD3			'12	Wc Cd
				strVal = strVal & UCase(Trim(.Text)) & iColSep
				.Col = C_ReqSeqNo3			'13	Reserve Seq
				strVal = strVal & Trim(.Text) & iColSep
				.Col = C_ReqNo3			'14 MRP Req No.
				strVal = strVal & UCase(Trim(.Text)) & iColSep
									'15	Document No
				strVal = strVal & UCase(frm1.txtItemDocumentNo.value) & iColSep
				
				strVal = strVal & "PI" & iColSep	'16 Trns Type
				
				strVal = strVal & "I01" & iColSep	'17 MoveType
				
				strVal = strVal & IntRows & iRowSep  '18 Row Count
				
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
				         				
			End IF
		Next

	End With
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
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
   
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0

	ggoSpread.source = frm1.vspddata1
    frm1.vspdData1.MaxRows = 0
	ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0
	ggoSpread.source = frm1.vspddata3
    frm1.vspdData3.MaxRows = 0
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
'----------  Coding part  -------------------------------------------------------------
'==============================================================================
' Function : SetTotalIssueQtyToGrid
' Description : vspdData2 ������ ���� �� ������°� vspdData1 ������ ���� 
'==============================================================================
Function SetTotalIssueQtyToGrid(ByVal Col, ByVal Row)

Dim LngvspdData1Row
Dim LngOnHandQty
Dim LngIssueQty
Dim LngTTlIssueQty
Dim i

SetTotalIssueQtyToGrid = False	
	'Because it take a lot of time during for ~ next operate, active row of vspdData1 is gotten  
	LngvspdData1Row = frm1.vspdData1.ActiveRow
	
    ggoSpread.Source = frm1.vspdData2
    
	With frm1.vspdData2
		.Row = Row
		.Col = C_OnHandQty
		LngOnHandQty = UNICDbl(.Text)
		.Col = C_IssueQty
		LngIssueQty =  UNICDbl(.Text)
		If LngOnHandQty < LngIssueQty Then
			Call Displaymsgbox("189516", "x", "x", "x")
			frm1.vspdData2.Row = Row
			frm1.vspdData2.Col = Col
			frm1.vspdData2.Text = 0
			Exit Function
		End If
		For i = 1 To .MaxRows
			.Row = i
			LngTTlIssueQty = LngTTlIssueQty + UNICDbl(.Text)
		Next
	End With

	With frm1.vspdData1
		.row = LngvspdData1Row
		.Col = C_TTlIssueQty
		.Text = UNIFormatNumber(LngTTlIssueQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	End With

SetTotalIssueQtyToGrid = True	

End Function

'==============================================================================
' Function : GetHiddenFocus
' Description : �����߻��� Hidden Spread Sheet�� ã�� SheetFocus�� ���� �Ѱ���.
'==============================================================================
Function GetHiddenFocus(lRow, lCol)
	Dim lRows1, lRows2						'Quantity of the Hidden Data Keys Referenced by FindData Function
	Dim strHdnItemCd, strHdnSeq, strHdnSlCd, strHdnTrackingNo			 'Variable of Hidden Keys
	Dim strHdnLotNo, strHdnLotSubNo, strHdnProdtOrderNo, strHdnOprNo	 'Variable of Hidden Keys
	Dim strItemCd, strSeq, strSlCd, strTrackingNo 						 'Variable of Visible Sheet Keys		
	Dim strLotNo, strLotSubNo, strProdtOrderNo, strOprNo				 'Variable of Visible Sheet Keys		
	
	If Trim(lCol) = "" Then
		lCol = C_BlockIndicator					'If Value of Column is not passed, Assign Value of the First Column in Second Spread Sheet
	End If
	'Find Key Datas in Hidden Spread Sheet
	
	With frm1.vspdData3
		.Row = lRow
		.Col = C_CompntCd3
		strHdnItemCd = Trim(.Text)
		.Col = C_ReqSeqNo3
		strHdnSeq = Trim(.Text)
		.Col = C_SLCd3
		strHdnSlCd = Trim(.Text)
		.Col = C_AllTrackingNo3
		strHdnTrackingNo = Trim(.Text)
		.Col = C_LotNo3
		strHdnLotNo = Trim(.Text)
		.Col = C_LotSubNo3
		strHdnLotSubNo = Trim(.Text)
		.Col = C_ProdtOrderNo3
		strHdnProdtOrderNo = Trim(.Text)
		.Col = C_OprNo3
		strHdnOprNo = Trim(.Text)
	End With
	'Compare Key Datas to Visible Spread Sheets
	With frm1		
		For lRows1 = 1 To .vspdData1.MaxRows
			.vspdData1.Row = lRows1
			.vspdData1.Col = C_ProdtOrderNo			
			strProdtOrderNo = Trim(.vspdData1.Text) 
			.vspdData1.Col = C_OprNo
			strOprNo = Trim(.vspdData1.Text)
			.vspdData1.Col = C_CompntCd	
			strItemCd = Trim(.vspdData1.Text)
			.vspdData1.Col = C_ReqSeqNo	
			strSeq = Trim(.vspdData1.Text)         
			
			If strProdtOrderNo = strHdnProdtOrderNo And strOprNo = strHdnOprNo _
			  And strItemCd = strHdnItemCd And strSeq = strHdnSeq Then
				.vspdData1.Col = C_ProdtOrderNo	
				.vspdData1.focus
				.vspdData1.Action = 0
				lgOldRow = lRows1			'�� If this line is omitted, program Could not query Data When errors occur
				.vspdData2.MaxRows = 0
				ggoSpread.Source = .vspdData2
				If CopyFromHSheet(lRows1, strProdtOrderNo,strItemCd,strSeq,strOprNo) = True Then
				    For lRows2 = 1 To .vspdData2.MaxRows
						.vspdData2.Row = lRows2
						.vspdData2.Col = C_SLCd			
						strSlCd  = Trim(.vspdData2.Text)
						.vspdData2.Col = C_AllTrackingNo	
						strTrackingNo  = Trim(.vspdData2.Text)
						.vspdData2.Col = C_LotNo				
						strLotNo   = Trim(.vspdData2.Text)
						.vspdData2.Col = C_LotSubNo			
						strLotSubNo   = Trim(.vspdData2.Text)
						'Find Key Datas in Second Sheet and then Focus the Cell 
						If strSlCd = strHdnSlCd And strTrackingNo = strHdnTrackingNo _
						  And strLotNo = strHdnLotNo And strLotSubNo = strHdnLotSubNo Then
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

'========================================================================================
' Function Name : ViewHidden
' Function Desc : Show Detail Field
'========================================================================================
Function ViewHidden(StrMnuID, MnuCount, StrImageSize )
    Dim ii

    For ii = 1 To MnuCount
        If document.all(StrMnuID & ii).style.display = "" Then 
           document.all(StrMnuID & ii).style.display = "none"
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/Smallplus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigPlus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
			End Select		
        Else
           document.all(StrMnuID & ii).style.display = ""
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/SmallMinus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigMinus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
			End Select
        End If
    Next    

End Function

