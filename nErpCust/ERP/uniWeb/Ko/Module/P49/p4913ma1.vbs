
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY0_ID						= "p4913mb0.asp"		'��: Master
Const BIZ_PGM_QRY1_ID						= "p4913mb1.asp"		'��: order �۾���Ȳ 
Const BIZ_PGM_QRY2_ID						= "p4913mb2.asp"		'��: �۾���Ȳ detail
Const BIZ_PGM_QRY3_ID						= "p4913mb3.asp"		'��: ������Ȳ 

Const BIZ_PGM_QRY4_ID						= "p4913mb4.asp"		'��: �������� 
Const BIZ_PGM_QRY5_ID						= "p4913mb5.asp"		'��: �����Ȳ 

Const BIZ_PGM_SAVE_ID						= "p4913mb01.asp"		'��: SAVE

Const BIZ_PGM_JUMPORDERRUN_ID				= "p4912ma1"

Const TAB1 = 1
Const TAB2 = 2

'-------------------------------
' Column Constants : Spread 1	������Ȳ 
'-------------------------------
Dim C_ProdtOrderNo
Dim C_OprNo
Dim C_TrackingNo
Dim C_ShiftCd
Dim C_ItemCd
Dim C_ItemNm
Dim C_ProdtOrderQty
Dim C_ProdtOrderUnit
Dim C_BadQty
Dim C_ExiProdQtyInOrderUnit
Dim C_ProdQtyInOrderUnit
Dim C_ProdQtyInOrderSum
Dim C_ExiGoodQtyInOrderUnit
Dim C_GoodQty
Dim C_GoodQtySum
Dim C_StApply
Dim C_StdTime
Dim C_IncTime
Dim C_DescTime
Dim C_OtTime
Dim C_EtcTime
Dim C_WkTime
Dim C_WkLossTime
Dim C_RealTime

'-------------------------------
' Column Constants : Spread 2
'-------------------------------

Dim C_PlantCd2
Dim C_WcCd2
Dim C_ReportDt2
Dim C_ProdtOrderNo2
Dim C_OprNo2
Dim C_Seq2
Dim C_ResourceCd2
Dim C_ResourceDesc2
Dim C_WorkType2
Dim C_WorkTypeDesc2
Dim C_WorkMan2
Dim C_WorkTime2

'-------------------------------
' Column Constants : Spread 3
'-------------------------------

Dim C_PlantCd3
Dim C_WcCd3
Dim C_ReportDt3
Dim C_ProdtOrderNo3
Dim C_OprNo3
Dim C_SeqNo3
Dim C_ItemCd3
Dim C_ResourceCd3
Dim C_ResourceDesc3
Dim C_StTime3
Dim C_EndTime3
Dim C_LossMan3
Dim C_WkLossQty3
Dim C_WkLossCd3
'Dim C_WkLossPopup3
Dim C_WkLossDesc3
Dim C_WkLossType3
Dim C_RtDeptCd3
Dim C_RtDeptPopup3
Dim C_RtDeptNm3
Dim C_Notes3

'-------------------------------
' Column Constants : Spread 4	�������� 
'-------------------------------

Dim	C_PlantCd4
Dim C_WcCd4
Dim C_ReportDt4
Dim C_SeqNo4
Dim C_EmpNo4
Dim C_EmpNoPopup4
Dim C_EmpNm4
Dim C_Time4
Dim C_Notes4

'-------------------------------
' Column Constants : Spread 5	�����Ȳ 
'-------------------------------

Dim C_PlantCd5
Dim C_WcCd5
Dim C_ReportDt5
Dim C_SeqNo5
Dim C_EmpNo5
Dim C_EmpNoPopup5
Dim C_EmpNm5
Dim C_ActCd5
Dim C_ActCdDesc5
Dim C_Notes5

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3

Dim lgSortKey1
Dim lgSortKey2
Dim lgSortKey3
Dim lgSortKey4
Dim lgSortKey5

Dim lgOldRow1
Dim lgOldRow2

Dim lgLngCurRows

Dim lgProdtOrderNo1
Dim lgProdtOrderNo2

Dim lgBlnFlgChgValue
Dim gRowVsp1

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop
Dim gSelframeFlg

Dim gwkTime
Dim grtTime
Dim gincTime
Dim gdescTime
Dim getcTime
Dim gotTime
Dim glossTime

Dim gfetcTime
Dim gflossTime
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

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgOldRow1 = 0
	lgOldRow2 = 0
	lgSortKey1 = 1
	lgSortKey2 = 1
	lgSortKey3 = 1
	lgSortKey4 = 1
	lgSortKey5 = 1

	gSelframeFlg = 1

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
	frm1.txtprodDt.text = StartDate

	If Trim(ReadCookie("txtPlantCd")) <> "" Then
		frm1.txtPlantCd.Value		= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value		= ReadCookie("txtPlantNm")
		frm1.txtprodDt.Text			= ReadCookie("txtprodDt")
		frm1.txtWcCd.Value			= ReadCookie("txtWcCd")
		frm1.txtWcNm.value			= ReadCookie("txtWcNm")

		lstrPgmID = ReadCookie("txtPGMID")
	End If	

	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtprodDt", ""
	WriteCookie "txtWcCd", ""	
	WriteCookie "txtWcNm", ""
	WriteCookie "txtPGMID", ""

End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    Call InitSpreadPosVariables(pvSpdNo)

    With frm1
		If pvSpdNo = "A" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 1 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData1
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.vspdData1.ReDraw = False

			.vspdData1.MaxCols = C_RealTime + 1
			.vspdData1.MaxRows = 0

			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_ProdtOrderNo,				"����������ȣ", 12
			ggoSpread.SSSetEdit		C_OprNo,					"������ȣ", 12
			ggoSpread.SSSetEdit		C_TrackingNo,				"Tracking No.", 12
			ggoSpread.SSSetEdit		C_ShiftCd,					"Shift", 8
			ggoSpread.SSSetEdit		C_ItemCd,					"ǰ��", 12
			ggoSpread.SSSetEdit		C_ItemNm,					"ǰ���", 18
			ggoSpread.SSSetFloat	C_ProdtOrderQty,			"��������", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ProdtOrderUnit,			"��������", 10
			ggoSpread.SSSetFloat	C_BadQty,					"�ҷ���", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ExiProdQtyInOrderUnit,	"�������Լ�", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit,		"���Լ�", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInOrderSum,		"���Դ���", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ExiGoodQtyInOrderUnit,	"�����ϼ���", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQty,					"�ϼ���", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtySum,				"�ϼ�����", 12,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetTime		C_StApply,					"ǥ�ؽð�", 13, 2, 1, 1
			ggoSpread.SSSetTime		C_StdTime,					"ǥ�ذ���", 13, 2, 1, 1

			ggoSpread.SSSetTime		C_IncTime,					"�����������", 13, 2, 1, 1
			ggoSpread.SSSetTime		C_DescTime,					"�����������", 13, 2, 1, 1
			ggoSpread.SSSetTime		C_OtTime,					"�ܾ�����", 13, 2, 1, 1
			ggoSpread.SSSetTime		C_EtcTime,					"��Ÿ����", 13, 2, 1, 1

			ggoSpread.SSSetEdit		C_WkTime,					"�۾�����", 13, 2', 1, 1
			ggoSpread.SSSetEdit		C_WkLossTime,				"���ǰ���", 13, 2', 1, 1
			ggoSpread.SSSetEdit		C_RealTime,					"�ǵ�����", 13, 2', 1, 1

 			Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
 			Call ggoSpread.SSSetColHidden(C_ShiftCd, C_ShiftCd, True)
 			

			Call SetSpreadLock("A")

			.vspdData1.ReDraw = True
		End If

		If pvSpdNo = "B" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 2 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

			.vspdData2.ReDraw = False

			.vspdData2.MaxCols = C_WorkTime2 + 1
			.vspdData2.MaxRows = 0

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit		C_PlantCd2,			"�����ڵ�", 12
			ggoSpread.SSSetEdit		C_WcCd2,			"�۾���", 12
			ggoSpread.SSSetDate		C_ReportDt2,		"�۾�����", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_ProdtOrderNo2,	"����������ȣ", 12
			ggoSpread.SSSetEdit		C_OprNo2,			"������ȣ", 12
			ggoSpread.SSSetEdit		C_Seq2,				"����", 12
			ggoSpread.SSSetCombo	C_ResourceCd2,		"�ڿ��ڵ�", 12
			ggoSpread.SSSetCombo	C_ResourceDesc2,	"�ڿ�", 18
			ggoSpread.SSSetCombo	C_WorkType2,		"����", 12
			ggoSpread.SSSetCombo	C_WorkTypeDesc2,	"����", 10
			ggoSpread.SSSetFloat	C_WorkMan2,			"�ο�", 12, 0, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetTime		C_WorkTime2,		"�ð�", 13, 2, 1, 1

			Call ggoSpread.SSSetColHidden(C_PlantCd2, C_Seq2, True)
			Call ggoSpread.SSSetColHidden(C_ResourceCd2, C_ResourceCd2, True)
			Call ggoSpread.SSSetColHidden(C_WorkType2, C_WorkType2, True)
 			Call ggoSpread.SSSetColHidden( .vspdData2.MaxCols, .vspdData2.MaxCols, True)

			Call SetSpreadLock("B")

			.vspdData2.Redraw = False
		End If

		If pvSpdNo = "C" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 3 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData3
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

			.vspdData3.ReDraw = False

			.vspdData3.MaxCols = C_Notes3 + 1
			.vspdData3.MaxRows = 0

			Call GetSpreadColumnPos("C")

			ggoSpread.SSSetEdit		C_PlantCd3,		"�����ڵ�", 12
			ggoSpread.SSSetEdit		C_WcCd3,		"�۾���", 12
			ggoSpread.SSSetEdit		C_ReportDt3,	"�۾�����", 12
			ggoSpread.SSSetEdit		C_ProdtOrderNo3,"����������ȣ", 12
			ggoSpread.SSSetEdit		C_OprNo3,		"������ȣ", 12
			ggoSpread.SSSetEdit		C_SeqNo3,		"����", 12
			ggoSpread.SSSetEdit		C_ItemCd3,		"ǰ���ڵ�", 12
			ggoSpread.SSSetCombo	C_ResourceCd3,	"�ڿ��ڵ�", 12
			ggoSpread.SSSetCombo	C_ResourceDesc3,"�ڿ�", 18
			ggoSpread.SSSetTime		C_StTime3,		"���۽ð�", 13, 2, 1, 1
			ggoSpread.SSSetTime		C_EndTime3,		"����ð�", 13, 2, 1, 1
			ggoSpread.SSSetFloat	C_LossMan3,		"�����ο�", 12, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetTime		C_WkLossQty3,	"���ǰ���", 13, 2, 1, 1
			ggoSpread.SSSetCombo	C_WkLossCd3,	"���Ǹ�", 12
			ggoSpread.SSSetCombo	C_WkLossDesc3,	"���Ǹ�", 12
			ggoSpread.SSSetEdit		C_WkLossType3,	"��������", 12
			ggoSpread.SSSetEdit		C_RtDeptCd3,	"å�Ӻμ�", 12
			ggoSpread.SSSetButton	C_RtDeptPopup3
			ggoSpread.SSSetEdit		C_RtDeptNm3,	"å�Ӻμ�", 12
			ggoSpread.SSSetEdit		C_Notes3,		"�߻�����", 50

			Call ggoSpread.SSSetColHidden(C_PlantCd3, C_ItemCd3, True)
			Call ggoSpread.SSSetColHidden(C_ResourceCd3, C_ResourceCd3, True)
			Call ggoSpread.SSSetColHidden(C_WkLossCd3, C_WkLossCd3, True)
			Call ggoSpread.SSSetColHidden(C_WkLossType3, C_WkLossType3, True)
 			Call ggoSpread.SSSetColHidden( .vspdData3.MaxCols, .vspdData3.MaxCols, True)

			.vspdData3.ReDraw = True

			Call SetSpreadLock("C")

		End If

		' TAB2
		If pvSpdNo = "D" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 4 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData4
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

			.vspdData4.ReDraw = False

			.vspdData4.MaxCols = C_Notes4 + 1
			.vspdData4.MaxRows = 0

			Call GetSpreadColumnPos("D")

			ggoSpread.SSSetEdit		C_PlantCd4,		"�����ڵ�", 14
			ggoSpread.SSSetEdit		C_WcCd4,		"�۾���", 14
			ggoSpread.SSSetDate		C_ReportDt4,	"�۾�����", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_SeqNo4,		"����", 10
			ggoSpread.SSSetEdit		C_EmpNo4,		"���", 14
			ggoSpread.SSSetButton	C_EmpNoPopup4
			ggoSpread.SSSetEdit		C_EmpNm4,		"����", 14
			ggoSpread.SSSetTime		C_Time4,		"�ð�", 13, 2, 1, 1
			ggoSpread.SSSetEdit		C_Notes4,		"��������", 50

			Call ggoSpread.SSSetColHidden(C_PlantCd4, C_SeqNo4, True)
 			Call ggoSpread.SSSetColHidden( .vspdData4.MaxCols, .vspdData4.MaxCols, True)

			.vspdData4.ReDraw = True

			Call SetSpreadLock("D")

		End If

		If pvSpdNo = "E" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 5 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData5
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

			.vspdData5.ReDraw = False

			.vspdData5.MaxCols = C_Notes5 + 1
			.vspdData5.MaxRows = 0

			Call GetSpreadColumnPos("E")

			ggoSpread.SSSetEdit		C_PlantCd5,		"�����ڵ�", 14
			ggoSpread.SSSetEdit		C_WcCd5,		"�۾���", 14
			ggoSpread.SSSetDate		C_ReportDt5,	"�۾�����", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_SeqNo5,		"����", 10
			ggoSpread.SSSetEdit		C_EmpNo5,		"���", 14
			ggoSpread.SSSetButton	C_EmpNoPopup5
			ggoSpread.SSSetEdit		C_EmpNm5,		"����", 14
			ggoSpread.SSSetCombo	C_ActCd5,		"�����", 14
			ggoSpread.SSSetCombo	C_ActCdDesc5,	"�����", 14
			ggoSpread.SSSetEdit		C_Notes5,		"����", 50

			Call ggoSpread.SSSetColHidden(C_PlantCd5, C_SeqNo5, True)
			Call ggoSpread.SSSetColHidden(C_ActCd5, C_ActCd5, True)
 			Call ggoSpread.SSSetColHidden( .vspdData5.MaxCols, .vspdData5.MaxCols, True)

			.vspdData5.ReDraw = True

			Call SetSpreadLock("E")

		End If

    End With

End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		Select Case pvSpdNo
			Case "A"
				'-------------------------
				' Set Lock Prop :Spread 1
				'-------------------------
				.vspdData1.ReDraw = False
				ggoSpread.Source = .vspdData1
				ggoSpread.SpreadLock -1, -1	' Set Lock Property : Spread 1
				.vspdData1.ReDraw = True
			Case "B"
				'-------------------------
				' Set Lock Prop :Spread 2
				'-------------------------
				.vspdData2.ReDraw = False
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadLock	 C_PlantCd2,		-1, C_PlantCd2
				ggoSpread.SpreadLock	 C_WcCd2,			-1, C_WcCd2
				ggoSpread.SpreadLock	 C_ReportDt2,		-1, C_ReportDt2
				ggoSpread.SpreadLock	 C_ProdtOrderNo2,	-1, C_ProdtOrderNo2
				ggoSpread.SpreadLock	 C_OprNo2,			-1, C_OprNo2
				ggoSpread.SpreadLock	 C_Seq2,			-1, C_Seq2
				ggoSpread.SpreadLock	 C_ResourceCd2,		-1, C_ResourceCd2
				ggoSpread.SpreadLock	 C_ResourceDesc2,	-1, C_ResourceDesc2
				ggoSpread.SpreadLock	 C_WorkType2,		-1, C_WorkType2
				ggoSpread.SpreadLock	 C_WorkTypeDesc2,	-1, C_WorkTypeDesc2
				ggoSpread.SSSetRequired	 C_WorkMan2,		-1, C_WorkMan2
				ggoSpread.SSSetRequired	 C_WorkTime2,		-1, C_WorkTime2
				.vspdData2.ReDraw = True
			Case "C"
				'-------------------------
				' Set Lock Prop :Spread 3
				'-------------------------
				.vspdData3.ReDraw = False
				ggoSpread.Source = .vspdData3
				ggoSpread.SpreadLock	 C_PlantCd3,		-1, C_PlantCd3							'
				ggoSpread.SpreadLock	 C_WcCd3,			-1, C_WcCd3							'
				ggoSpread.SpreadLock	 C_ReportDt3,		-1, C_ReportDt3							'
				ggoSpread.SpreadLock	 C_ProdtOrderNo3,	-1, C_ProdtOrderNo3
				ggoSpread.SpreadLock	 C_OprNo3,			-1, C_OprNo3
				ggoSpread.SpreadLock	 C_SeqNo3,			-1, C_SeqNo3
				ggoSpread.SpreadLock	 C_ItemCd3,			-1, C_ItemCd3
				ggoSpread.SpreadLock	 C_ResourceCd3,		-1, C_ResourceCd3
				ggoSpread.SpreadLock	 C_ResourceDesc3,	-1, C_ResourceDesc3
				ggoSpread.SSSetRequired	 C_LossMan3,		-1, C_LossMan3
				ggoSpread.SSSetRequired	 C_WkLossQty3,		-1, C_WkLossQty3
'				ggoSpread.SpreadLock	 C_WkLossDesc3,		-1, C_WkLossDesc3
				ggoSpread.SpreadLock	 C_RtDeptNm3,		-1, C_RtDeptNm3
				.vspdData3.ReDraw = True
			Case "D"
				'-------------------------
				' Set Lock Prop :Spread 4
				'-------------------------
				.vspdData4.ReDraw = False
				ggoSpread.Source = .vspdData4
				ggoSpread.SpreadLock	 C_EmpNo4,		-1, C_EmpNo4							'
				ggoSpread.SpreadLock	 C_EmpNoPopup4,	-1, C_EmpNoPopup4							'
				ggoSpread.SpreadLock	 C_EmpNm4,		-1, C_EmpNm4							'
				ggoSpread.SSSetRequired	 C_Time4,		-1, C_Time4
				.vspdData4.ReDraw = True
			Case "E"
				'-------------------------
				' Set Lock Prop :Spread 5
				'-------------------------
				.vspdData5.ReDraw = False
				ggoSpread.Source = .vspdData5
				ggoSpread.SpreadLock	 C_EmpNo5,		-1, C_EmpNo5							'
				ggoSpread.SpreadLock	 C_EmpNoPopup5,	-1, C_EmpNoPopup5							'
				ggoSpread.SpreadLock	 C_EmpNm5,		-1, C_EmpNm5							'
				ggoSpread.SSSetRequired	 C_ActCdDesc5,	-1, C_ActCdDesc5							'
				.vspdData5.ReDraw = True
		End Select
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc :
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal InOutType, ByVal pvSpdNo)

	Select Case pvSpdNo
'		-- TAB1 --
		Case "A"
			ggoSpread.Source = frm1.vspdData1
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			If InOutType = "N" Then
				ggoSpread.SSSetRequired		C_ResourceCd2,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ResourceDesc2,	pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkType2,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkTypeDesc2,	pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkMan2,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkTime2,		pvStartRow, pvEndRow
			Else
				ggoSpread.SSSetRequired		C_ResourceCd2,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ResourceDesc2,	pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkType2,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkTypeDesc2,	pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkMan2,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WorkTime2,		pvStartRow, pvEndRow
			End If

		Case "C"
			ggoSpread.Source = frm1.vspdData3
			If InOutType = "N" Then
				ggoSpread.SSSetRequired		C_ResourceCd3,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ResourceDesc3,	pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_LossMan3,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WkLossQty3,		pvStartRow, pvEndRow
				ggoSpread.SSSetProtected	C_RtDeptNm3,		pvStartRow, pvEndRow
			Else
				ggoSpread.SSSetRequired		C_ResourceCd3,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ResourceDesc3,	pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_LossMan3,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WkLossQty3,		pvStartRow, pvEndRow
				ggoSpread.SSSetProtected	C_RtDeptNm3,		pvStartRow, pvEndRow
			End If

'		-- TAB2 --
		Case "D"
			ggoSpread.Source = frm1.vspdData4
			If InOutType = "N" Then
				ggoSpread.SSSetRequired		C_PlantCd4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WcCd4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ReportDt4,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_SeqNo4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_EmpNo4,			pvStartRow, pvEndRow
				ggoSpread.SSSetProtected	C_EmpNm4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_Time4,			pvStartRow, pvEndRow
			Else
				ggoSpread.SSSetRequired		C_PlantCd4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WcCd4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ReportDt4,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_SeqNo4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_EmpNo4,			pvStartRow, pvEndRow
				ggoSpread.SSSetProtected	C_EmpNm4,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_Time4,			pvStartRow, pvEndRow
			End If

		Case "E"
			ggoSpread.Source = frm1.vspdData5
			If InOutType = "N" Then
				ggoSpread.SSSetRequired		C_PlantCd5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WcCd5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ReportDt5,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_SeqNo5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_EmpNo5,			pvStartRow, pvEndRow
				ggoSpread.SSSetProtected	C_EmpNm5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ActCdDesc5,		pvStartRow, pvEndRow
			Else
				ggoSpread.SSSetRequired		C_PlantCd5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_WcCd5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ReportDt5,		pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_SeqNo5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_EmpNo5,			pvStartRow, pvEndRow
				ggoSpread.SSSetProtected	C_EmpNm5,			pvStartRow, pvEndRow
				ggoSpread.SSSetRequired		C_ActCdDesc5,		pvStartRow, pvEndRow
			End If
	End Select
End Sub

'==========================================  2.2.6 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P4901' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_WorkType2
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_WorkTypeDesc2

End Sub

'=========================================================================================================
'	Name : InitSpreadComboBox2()
'	Description : After Query Combo Display
'=========================================================================================================
Sub InitSpreadComboBox2()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strPlantCd

	strPlantCd = UCase(Trim(frm1.txtPlantCd.Value))

	Call CommonQueryRs(" RESOURCE_CD,DESCRIPTION "," P_RESOURCE "," PLANT_CD = '" & strPlantCd & "' AND RESOURCE_TYPE = 'L' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ResourceCd2
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ResourceDesc2

    ggoSpread.Source = frm1.vspdData3
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ResourceCd3
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ResourceDesc3

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P4903' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_WkLossCd3
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_WkLossDesc3

End Sub

'=========================================================================================================
'	Name : InitSpreadComboBox5()
'	Description : After Query Combo Display
'=========================================================================================================
Sub InitSpreadComboBox5()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P4902' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	ggoSpread.Source = frm1.vspdData5
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ActCd5
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ActCdDesc5
End Sub


'==========================================  2.2.6 InitData()  ==========================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim	intIndex

	With frm1.vspdData1
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_OrderType
			intIndex = .value
			.Col = C_OrderTypeDesc
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
		'-------------------------------
		' Column Constants : Spread 1
		'-------------------------------
		C_ProdtOrderNo			= 1
		C_OprNo					= 2
		C_TrackingNo			= 3
		C_ShiftCd				= 4
		C_ItemCd				= 5
		C_ItemNm				= 6
		C_ProdtOrderQty			= 7
		C_ProdtOrderUnit		= 8
		C_BadQty				= 9
		C_ExiProdQtyInOrderUnit	= 10
		C_ProdQtyInOrderUnit	= 11
		C_ProdQtyInOrderSum		= 12
		C_ExiGoodQtyInOrderUnit	= 13
		C_GoodQty				= 14
		C_GoodQtySum			= 15
		C_StApply				= 16
		C_StdTime				= 17
		C_IncTime				= 18
		C_DescTime				= 19
		C_OtTime				= 20
		C_EtcTime				= 21
		C_WkTime				= 22
		C_WkLossTime			= 23
		C_RealTime				= 24

	End If

	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 2
		'-------------------------------

		C_PlantCd2			= 1
		C_WcCd2				= 2
		C_ReportDt2			= 3
		C_ProdtOrderNo2		= 4
		C_OprNo2			= 5
		C_Seq2				= 6
		C_ResourceCd2		= 7
		C_ResourceDesc2		= 8
		C_WorkType2			= 9
		C_WorkTypeDesc2		= 10
		C_WorkMan2			= 11
		C_WorkTime2			= 12

	End If

	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 3
		'-------------------------------
		C_PlantCd3			= 1
		C_WcCd3				= 2
		C_ReportDt3			= 3
		C_ProdtOrderNo3		= 4
		C_OprNo3			= 5
		C_SeqNo3			= 6
		C_ItemCd3			= 7
		C_ResourceCd3		= 8
		C_ResourceDesc3		= 9
		C_StTime3			= 10
		C_EndTime3			= 11
		C_LossMan3			= 12
		C_WkLossQty3		= 13
		C_WkLossCd3			= 14
		C_WkLossDesc3		= 15
		C_WkLossType3		= 16
		C_RtDeptCd3			= 17
		C_RtDeptPopup3		= 18
		C_RtDeptNm3			= 19
		C_Notes3			= 20

	End If

	' TAB2
	If pvSpdNo = "D" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 4
		'-------------------------------
		C_PlantCd4			= 1
		C_WcCd4				= 2
		C_ReportDt4			= 3
		C_SeqNo4			= 4
		C_EmpNo4			= 5
		C_EmpNoPopup4		= 6
		C_EmpNm4			= 7
		C_Time4				= 8
		C_Notes4			= 9
	End If

	If pvSpdNo = "E" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 5
		'-------------------------------
		C_PlantCd5			= 1
		C_WcCd5				= 2
		C_ReportDt5			= 3
		C_SeqNo5			= 4
		C_EmpNo5			= 5
		C_EmpNoPopup5		= 6
		C_EmpNm5			= 7
		C_ActCd5			= 8
		C_ActCdDesc5		= 9
		C_Notes5			= 10
	End If

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim	iCurColumnPos

 	Select Case Ucase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData1

 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ProdtOrderNo			= iCurColumnPos(1)
			C_OprNo					= iCurColumnPos(2)
			C_TrackingNo			= iCurColumnPos(3)
			C_ShiftCd				= iCurColumnPos(4)
			C_ItemCd				= iCurColumnPos(5)
			C_ItemNm				= iCurColumnPos(6)
			C_ProdtOrderQty			= iCurColumnPos(7)
			C_ProdtOrderUnit		= iCurColumnPos(8)
			C_BadQty				= iCurColumnPos(9)
			C_ExiProdQtyInOrderUnit	= iCurColumnPos(10)
			C_ProdQtyInOrderUnit	= iCurColumnPos(11)
			C_ProdQtyInOrderSum		= iCurColumnPos(12)
			C_ExiGoodQtyInOrderUnit	= iCurColumnPos(13)
			C_GoodQty				= iCurColumnPos(14)
			C_GoodQtySum			= iCurColumnPos(15)
			C_StApply				= iCurColumnPos(16)
			C_StdTime				= iCurColumnPos(17)
			C_IncTime				= iCurColumnPos(18)
			C_DescTime				= iCurColumnPos(19)
			C_OtTime				= iCurColumnPos(20)
			C_EtcTime				= iCurColumnPos(21)
			C_WkTime				= iCurColumnPos(22)
			C_WkLossTime			= iCurColumnPos(23)
			C_RealTime				= iCurColumnPos(24)

		Case "B"
			ggoSpread.Source = frm1.vspdData2

 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd2			= iCurColumnPos(1)
			C_WcCd2				= iCurColumnPos(2)
			C_ReportDt2			= iCurColumnPos(3)
			C_ProdtOrderNo2		= iCurColumnPos(4)
			C_OprNo2			= iCurColumnPos(5)
			C_Seq2				= iCurColumnPos(6)
			C_ResourceCd2		= iCurColumnPos(7)
			C_ResourceDesc2		= iCurColumnPos(8)
			C_WorkType2			= iCurColumnPos(9)
			C_WorkTypeDesc2		= iCurColumnPos(10)
			C_WorkMan2			= iCurColumnPos(11)
			C_WorkTime2			= iCurColumnPos(12)

		Case "C"
			ggoSpread.Source = frm1.vspdData3
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd3			= iCurColumnPos(1)
			C_WcCd3				= iCurColumnPos(2)
			C_ReportDt3			= iCurColumnPos(3)
			C_ProdtOrderNo3		= iCurColumnPos(4)
			C_OprNo3			= iCurColumnPos(5)
			C_SeqNo3			= iCurColumnPos(6)
			C_ItemCd3			= iCurColumnPos(7)
			C_ResourceCd3		= iCurColumnPos(8)
			C_ResourceDesc3		= iCurColumnPos(9)
			C_StTime3			= iCurColumnPos(10)
			C_EndTime3			= iCurColumnPos(11)
			C_LossMan3			= iCurColumnPos(12)
			C_WkLossQty3		= iCurColumnPos(13)
			C_WkLossCd3			= iCurColumnPos(14)
			C_WkLossDesc3		= iCurColumnPos(15)
			C_WkLossType3		= iCurColumnPos(16)
			C_RtDeptCd3			= iCurColumnPos(17)
			C_RtDeptPopup3		= iCurColumnPos(18)
			C_RtDeptNm3			= iCurColumnPos(19)
			C_Notes3			= iCurColumnPos(20)

		' TAB2
		Case "D"
			ggoSpread.Source = frm1.vspdData4

 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd4			= iCurColumnPos(1)
			C_WcCd4				= iCurColumnPos(2)
			C_ReportDt4			= iCurColumnPos(3)
			C_SeqNo4			= iCurColumnPos(4)
			C_EmpNo4			= iCurColumnPos(5)
			C_EmpNoPopup4		= iCurColumnPos(6)
			C_EmpNm4			= iCurColumnPos(7)
			C_Time4				= iCurColumnPos(8)
			C_Notes4			= iCurColumnPos(9)

		Case "E"
			ggoSpread.Source = frm1.vspdData5

 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd5			= iCurColumnPos(1)
			C_WcCd5				= iCurColumnPos(2)
			C_ReportDt5			= iCurColumnPos(3)
			C_SeqNo5			= iCurColumnPos(4)
			C_EmpNo5			= iCurColumnPos(5)
			C_EmpNoPopup5		= iCurColumnPos(6)
			C_EmpNm5			= iCurColumnPos(7)
			C_ActCd5			= iCurColumnPos(8)
			C_ActCdDesc5		= iCurColumnPos(9)
			C_Notes5			= iCurColumnPos(10)

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

'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim	arrRet
	Dim	arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"					' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"						' TextBox �� 

    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)

    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)

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
	Dim	arrRet
	Dim	arrParam(8)
	Dim	iCalledAspName

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("P4111PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtprodDt.Text
	arrParam(2) = ""
	arrParam(3) = "RL"
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

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenConWC()
	Dim	arrRet
	Dim	arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"												' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") 			' Where Condition
	arrParam(5) = "�۾���"													' TextBox ��Ī 

    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)

    arrHeader(0) = "�۾���"													' Header��(0)
    arrHeader(1) = "�۾����"												' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus

End Function

'------------------------------------------  SetConWC()  ----------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)
	frm1.txtWCNm.Value    = arrRet(1)
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
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
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtprodDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtprodDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtprodDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.UpdateRow frm1.vspdData1.ActiveRow

End Sub

'==========================================================================================
'   Event Name : vspdData3_Change
'   Event Desc :
'==========================================================================================
Sub vspdData3_Change(ByVal Col , ByVal Row )
	Dim p_DeptCd
	Dim IntRetCd
	Dim strDept_nm
	Dim strInternal_cd
	Dim OrgChangeDt
	Dim iWhereList

    ggoSpread.Source = frm1.vspdData3

	Select Case Col
		Case C_RtDeptCd3
   	        frm1.vspdData3.Col = C_RtDeptCd3
			frm1.vspdData3.Row = Row
            p_DeptCd = frm1.vspdData3.value

			If frm1.vspdData3.value = "" Then
				frm1.vspdData3.Col = C_RtDeptCd3
				frm1.vspdData3.Row = Row
				frm1.vspdData3.value = ""
				OrgChangeDt = ""
			Else
				If  OrgChangeDt > "" Then
					iWhereList = " DEPT_CD = " &  FilterVar( p_DeptCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= " & FilterVar( OrgChangeDt, "''", "S")  & ")"
				Else
					iWhereList = " DEPT_CD = " &  FilterVar( p_DeptCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
				End If

				If CommonQueryRs(" DEPT_NM,INTERNAL_CD "," B_ACCT_DEPT ",iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					Call  DisplayMsgBox("970000", "x","�μ��ڵ�","x")
					frm1.vspdData3.Col = C_RtDeptNm3
					frm1.vspdData3.Row = Row
					frm1.vspdData3.value = ""
					frm1.vspdData3.Col = C_RtDeptCd3
					frm1.vspdData3.Row = Row
					frm1.vspdData3.value = ""
					Set gActiveElement = document.ActiveElement

					Exit Sub
				Else
					frm1.vspdData3.Col = C_RtDeptNm3
					frm1.vspdData3.Row = Row
					frm1.vspdData3.value = Replace(lgF0,Chr(11), "")
				End If

			End If
	End Select 

    ggoSpread.UpdateRow Row

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.UpdateRow frm1.vspdData1.ActiveRow

End Sub

'==========================================================================================
'   Event Name : vspdData4_Change
'   Event Desc :
'==========================================================================================
Sub vspdData4_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData4
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name : vspdData5_Change
'   Event Desc :
'==========================================================================================
Sub vspdData5_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData5
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name : ReqValueCheck(ByVal Col , ByVal Row)
'   Event Desc :
'==========================================================================================
Function ReqValueCheck(ByVal Col , ByVal Row)
	ggoSpread.Source = frm1.vspdData2

	ReqValueCheck = False
	If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
		With frm1.vspdData2
			.Col = C_WorkMan2
			.Row = Row
			.Text = 0
			.Col = C_WorkTime2
			.Row = Row
			.Text = ConvToTimeFormat(0)
		End With
		ReqValueCheck = False
		Exit Function
	End If
	ReqValueCheck = True
End Function

'==========================================================================================
'   Event Name : CalcWorkTimeTotal
'   Event Desc :
'==========================================================================================
Function CalcWorkTimeTotal()
	Dim wkTime
	Dim rtTime
	Dim incTime
	Dim descTime
	Dim etcTime
	Dim dotTime
	Dim lossTime

	Dim a, b
	Dim i

'*************************************************************************
'  - ���ܾ����� = ���ܾ��ð� x ���ܾ��ο�   (��� �ܾ������� ���� ó��)
'  - �۾����� = �����ð� - �޾��ð� + ���ܾ����� + �����ð�(+) - �����ð�(-)
'  - �ǵ����� = �۾����� - �����ǰ��� 
'*************************************************************************
	wkTime   = 0
	rtTime   = 0
	incTime  = 0
	descTime = 0
	etcTime  = 0
	dotTime  = 0
	lossTime = 0

	gwkTime   = 0
	grtTime   = 0
	gincTime  = 0
	gdescTime = 0
	getcTime  = 0
	gotTime   = 0
	glossTime = 0

	ggoSpread.Source = frm1.vspdData2
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Col = C_WorkType2	' �۾����� 
			.Row = i
			Select Case .Text
				Case "P1"	' ���� 
					.Col = 0
					.Row = i
					If .Text = ggoSpread.DeleteFlag Then
						wkTime = 0
					Else
						.Col = C_WorkMan2
						.Row = i
						a = .Text
						.Col = C_WorkTime2
						.Row = i

						b = ConvtoSec(.Text)

						wkTime = (a * b)
					End If
					gwkTime = gwkTime + wkTime
					a = 0 : b = 0
				Case "P2"	' �޾� 
					.Col = 0
					.Row = i
					If .Text = ggoSpread.DeleteFlag Then
						rtTime = 0
					Else
						.Col = C_WorkMan2
						.Row = i
						a = .Text
						.Col = C_WorkTime2
						.Row = i
						b = ConvtoSec(.Text)

						rtTime = (a * b)
					End If
					grtTime = grtTime + rtTime
					a = 0 : b = 0
				Case "P3"		' ����(+)����(����)
					.Col = 0
					.Row = i
					If .Text = ggoSpread.DeleteFlag Then
						incTime = 0
					Else
						.Col = C_WorkMan2
						.Row = i
						a = .Text
						.Col = C_WorkTime2
						.Row = i
						b = ConvtoSec(.Text)

						incTime = (a * b)
					End If
					gincTime = gincTime + incTime
					a = 0 : b = 0
				Case "P4"		' ����(-)����(����)
					.Col = 0
					.Row = i
					If .Text = ggoSpread.DeleteFlag Then
						descTime = 0
					Else
						.Col = C_WorkMan2
						.Row = i
						a = .Text
						.Col = C_WorkTime2
						.Row = i
						b = ConvtoSec(.Text)

						descTime = (a * b)
					End If
					gdescTime = gdescTime + descTime
					a = 0 : b = 0
				Case "P5"		' �ܾ����� 
					.Col = 0
					.Row = i
					If .Text = ggoSpread.DeleteFlag Then
						otTime = 0
					Else
						.Col = C_WorkMan2
						.Row = i
						a = .Text
						.Col = C_WorkTime2
						.Row = i
						b = ConvtoSec(.Text)

						otTime = (a * b)
					End If
					gotTime = gotTime + otTime
					a = 0 : b = 0
				Case "P6"		' ��Ÿ���� 
					.Col = 0
					.Row = i
					If .Text = ggoSpread.DeleteFlag Then
						etcTime = 0
					Else
						.Col = C_WorkMan2
						.Row = i
						a = .Text
						.Col = C_WorkTime2
						.Row = i
						b = ConvtoSec(.Text)

						etcTime = (a * b)
					End If
					getcTime = getcTime + etcTime
					a = 0 : b = 0
			End Select
		Next
	End With
	' ���ǰ��� 
	ggoSpread.Source = frm1.vspdData3
	With frm1.vspdData3
		For i = 1 To .MaxRows
			.Col = 0
			.Row = i
			If .Text = ggoSpread.DeleteFlag Then
				lossTime = 0
			Else
				.Col = C_LossMan3
				.Row = i
				a = .Text

				.Col = C_WkLossQty3
				.Row = i
				b = ConvtoSec(.Text)

				lossTime = (a * b)
			End If
			glossTime = glossTime + lossTime
			a = 0 : b = 0
		Next
	End With

	With frm1.vspdData1
	' ����������� 
		.Col = C_IncTime
		.Row = .ActiveRow
		.Text = ConvToTimeFormat(gincTime)
	' ����������� 
		.Col = C_DescTime
		.Row = .ActiveRow
		.Text = ConvToTimeFormat(gdescTime)
	' �ܾ����� 
		.Col = C_OtTime
		.Row = .ActiveRow
		.Text = ConvToTimeFormat(gotTime)
	' ��Ÿ���� 
		.Col = C_EtcTime
		.Row = .ActiveRow
		.Text = ConvToTimeFormat(getcTime)

	' �۾����� 
		.Col = C_WkTime
		.Row = .ActiveRow
		.Text = ConvToTimeFormat(gwkTime - grtTime + gotTime + gincTime - gdescTime)

	' ���ǰ��� 
		.Col = C_WkLossTime
		.Row = .ActiveRow
		.Text = ConvToTimeFormat(glossTime)

	' �ǵ����� 
		.Col = C_RealTime
		.Row = .ActiveRow
		.Text = ConvToTimeFormat( ((gwkTime - grtTime + gotTime + gincTime - gdescTime) + getcTime) - glossTime )
	End With

End Function

'==========================================================================================
'   Event Name : CalcFormTotal
'   Event Desc :
'==========================================================================================
Function CalcFormTotal()
	Dim etcTime
	Dim lossTime
	Dim strValue

	Dim i

	etcTime  = 0
	lossTime = 0

	gfetcTime  = 0
	gflossTime = 0

	ggoSpread.Source = frm1.vspdData1
	With frm1.vspdData1
		For i = 1 To .MaxRows
			' ��Ÿ���� 
			.Col = C_EtcTime
			.Row = i
			etcTime = ConvtoSec(.Text)
			gfetcTime = gfetcTime + etcTime

			' �����ǰ��� 
			.Col = C_WkLossTime
			.Row = i
			lossTime = ConvtoSec(.Text)
			gflossTime = gflossTime + lossTime
		Next
	End With

	' ��Ÿ���� 
	frm1.fpDoubleSingle14.Value = ConvToTimeFormat(gfetcTime)
	' �����ǰ��� 
	frm1.fpDoubleSingle8.Value = ConvToTimeFormat(gflossTime)

'  - ���ܾ����� = ���ܾ��ð� x ���ܾ��ο�   (��� �ܾ������� ���� ó��)
'  - �۾����� = �����ð� - �޾��ð� + ���ܾ����� + �����ð�(+) - �����ð�(-)
'  - �ǵ����� = �۾����� - �����ǰ��� 

	strValue = 0
	With frm1
	' �۾����� 
		strValue = .fpDoubleSingle5.Value - .fpDoubleSingle12.Value + ConvToSec(.fpDoubleSingle11.Value) + ConvToSec(.fpDoubleSingle6.Value) - ConvToSec(.fpDoubleSingle13.Value)
		.fpDoubleSingle4.Value = ConvToTimeFormat(strValue)
	' �ǵ����� 
		.fpDoubleSingle15.Value = ConvToTimeFormat((strValue + ConvToSec(.fpDoubleSingle14.Value)) - ConvToSec(.fpDoubleSingle8.Value))
	End With
End Function

'==========================================================================================
'   Event Name : vspdData1_Click
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

		ggoSpread.Source = frm1.vspdData1                          '��: Preset spreadsheet pointer
		If ggoSpread.SSCheckChange = True Then                   '��: Check If data is chaged
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
			If IntRetCD = vbNo Then
				Exit Sub
			Else
				Call FncCancelVsp1()
			End If
		End If

 		lgOldRow2 = 0

		lgStrPrevKey2 = ""
		lgStrPrevKey3 = ""
		frm1.vspdData2.MaxRows = 0
		frm1.vspdData3.MaxRows = 0
		frm1.KeyProdtOrderNo2.value = ""
		frm1.KeyOprNo2.value = ""

		' ����������ȣ 
		frm1.vspdData1.Col = C_ProdtOrderNo
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.KeyProdtOrderNo2.value =  Trim(frm1.vspdData1.Text)

		frm1.vspdData1.Col = C_OprNo
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.KeyOprNo2.value =  Trim(frm1.vspdData1.Text)

		frm1.vspdData1.Col = C_ItemCd
		frm1.vspdData1.Row = Row
		frm1.KeyItemCd.value =  Trim(frm1.vspdData1.Text)

		IF DbQuery2 = False Then
			Call RestoreToolBar()
			Exit Sub
		End If

		lgOldRow1 = frm1.vspdData1.Row

	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow1 <> Row Then

			ggoSpread.Source = frm1.vspdData1                          '��: Preset spreadsheet pointer
			If ggoSpread.SSCheckChange = True Then                   '��: Check If data is chaged
				IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
				If IntRetCD = vbNo Then
					Exit Sub
				Else
					Call FncCancelVsp1()
				End If
			End If

			lgOldRow2 = 0

			lgStrPrevKey2 = ""
			lgStrPrevKey3 = ""
			frm1.vspdData2.MaxRows = 0
			frm1.vspdData3.MaxRows = 0

			' ����������ȣ 
			frm1.vspdData1.Col = C_ProdtOrderNo
			frm1.vspdData1.Row = Row
			frm1.KeyProdtOrderNo2.value =  Trim(frm1.vspdData1.Text)

			frm1.vspdData1.Col = C_OprNo
			frm1.vspdData1.Row = Row
			frm1.KeyOprNo2.value =  Trim(frm1.vspdData1.Text)

			frm1.vspdData1.Col = C_ItemCd
			frm1.vspdData1.Row = Row
			frm1.KeyItemCd.value =  Trim(frm1.vspdData1.Text)

			IF DbQuery2 = False Then
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

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------

	gMouseClickStatus = "SP2C"

	Set gActiveSpdSheet = frm1.vspdData2

 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If

End Sub

'==========================================================================================
'   Event Name : vspdData3_Click
'   Event Desc :
'==========================================================================================
Sub vspdData3_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
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
'   Event Name : vspdData4_Click
'   Event Desc :
'==========================================================================================
Sub vspdData4_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP4C"

	Set gActiveSpdSheet = frm1.vspdData4

 	If frm1.vspdData4.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData4
 		If lgSortKey4 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey4 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey4		'Sort in Descending
 			lgSortKey4 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'==========================================================================================
'   Event Name : vspdData5_Click
'   Event Desc :
'==========================================================================================
Sub vspdData5_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP5C"

	Set gActiveSpdSheet = frm1.vspdData5

 	If frm1.vspdData5.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData5
 		If lgSortKey5 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey5 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey5		'Sort in Descending
 			lgSortKey5 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'==========================================================================================
'   Event Name : vspdData4_ButtonClicked
'   Event Desc :
'==========================================================================================

Sub vspdData3_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------
	With frm1.vspdData3

    ggoSpread.Source = frm1.vspdData3

    If Row > 0 And Col = C_RtDeptPopup3 Then
        .Col = C_RtDeptCd3
        .Row = Row

        Call OpenDept(1)
        Call SetActiveCell(frm1.vspdData3, C_RtDeptCd3, Row,"M","X","X")
		Set gActiveElement = document.activeElement
    End If

    End With
End Sub

'========================================================================================================
' Name : OpenDept
' Desc : �μ� POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtDept_cd.value			' ���Ǻο��� ���� ��� Code Condition
	Else 'spread
		arrParam(0) = frm1.vspdData3.Text			' Grid���� ���� ��� Code Condition
	End If


	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtDept_cd.focus
		Else 'spread
			frm1.vspdData3.Col = C_RtDeptCd3
			frm1.vspdData3.action =0
		End If
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If

End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)

	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtDept_cd.value = arrRet(0)
			.txtDept_cd_Nm.value = arrRet(1)
			lgBlnFlgChgValue = True
			.txtDept_cd.focus
		Else 'spread
			.vspdData3.Col = C_RtDeptNm3
			.vspdData3.Text = arrRet(1)
			.vspdData3.Col = C_RtDeptCd3
			.vspdData3.Text = arrRet(0)
			.vspdData3.action =0
			ggoSpread.UpdateRow frm1.vspdData3.ActiveRow
		End If
	End With
End Function

'==========================================================================================
'   Event Name : vspdData4_ButtonClicked
'   Event Desc :
'==========================================================================================

Sub vspdData4_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------
	With frm1.vspdData4

    ggoSpread.Source = frm1.vspdData4

    If Row > 0 And Col = C_EmpNoPopUp4 Then
        .Col = C_EmpNo4
        .Row = Row

        Call OpenEmp(.Text, "vspdData4")
        Call SetActiveCell(frm1.vspdData4, C_EmpNo4, Row,"M","X","X")
		Set gActiveElement = document.activeElement
    End If

    End With
End Sub

'==========================================================================================
'   Event Name : vspdData4_ButtonClicked
'   Event Desc :
'==========================================================================================

Sub vspdData5_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------
	With frm1.vspdData5

    ggoSpread.Source = frm1.vspdData5

    If Row > 0 And Col = C_EmpNoPopUp5 Then
        .Col = C_EmpNo5
        .Row = Row

        Call OpenEmp(.Text, "vspdData5")
        Call SetActiveCell(frm1.vspdData5, C_EmpNo5, Row,"M","X","X")
		Set gActiveElement = document.activeElement
    End If

    End With
End Sub

'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
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

    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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

    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then
		If lgStrPrevKey3 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery3 = False Then
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

	gMouseClickStatus = "SP2C"

	Set gActiveSpdSheet = frm1.vspdData2

End Sub

'==========================================================================================
'   Event Name : vspdData3_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData3_MouseDown(Button,Shift,x,y)

	If Button <> "1" And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If

	gMouseClickStatus = "SP3C"

	Set gActiveSpdSheet = frm1.vspdData3

End Sub

'==========================================================================================
'   Event Name : vspdData4_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData4_MouseDown(Button,Shift,x,y)

	If Button <> "1" And gMouseClickStatus = "SP4C" Then
		gMouseClickStatus = "SP4CR"
	End If

	gMouseClickStatus = "SP4C"

	Set gActiveSpdSheet = frm1.vspdData4

End Sub

'==========================================================================================
'   Event Name : vspdData5_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData5_MouseDown(Button,Shift,x,y)

	If Button <> "1" And gMouseClickStatus = "SP5C" Then
		gMouseClickStatus = "SP5CR"
	End If

	gMouseClickStatus = "SP5C"

	Set gActiveSpdSheet = frm1.vspdData5

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
 	Dim	iColumnName

 	If Row <= 0 Then
		Exit Sub
 	End If

  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'========================================================================================
' Function Name : vspdData3_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData3_DblClick(ByVal Col, ByVal Row)
 	Dim	iColumnName

 	If Row <= 0 Then
		Exit Sub
 	End If

  	If frm1.vspdData3.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
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
' Function Name : vspdData3_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData3
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

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================
' Function Name : vspdData3_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData3_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : vspdData4_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData4_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : vspdData5_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData5_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData5
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub


'==========================================================================================
'   Event Name :vspddata2_ComboSelChange
'   Event Desc :Combo Change Event
'==========================================================================================
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim strProdtOrderNo, strOprNo
	Dim LngFindRow

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

	With frm1.vspdData2
		.Row = Row
		Select Case Col
			Case C_ResourceCd2
				.Col = Col
				intIndex = .Value
				.Col = C_ResourceDesc2
				.Value = intIndex
			Case C_ResourceDesc2
				.Col = Col
				intIndex = .Value
				.Col = C_ResourceCd2
				.Value = intIndex
			Case C_WorkType2
				.Col = Col
				intIndex = .Value
				.Col = C_WorkTypeDesc2
				.Value = intIndex
			Case C_WorkTypeDesc2
				.Col = Col
				intIndex = .Value
				.Col = C_WorkType2
				.Value = intIndex
		End Select
    End With
End Sub

'==========================================================================================
'   Event Name :vspddata3_ComboSelChange
'   Event Desc :Combo Change Event
'==========================================================================================
Sub vspdData3_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim strProdtOrderNo, strOprNo
	Dim LngFindRow

	ggoSpread.Source = frm1.vspdData3
	ggoSpread.UpdateRow Row

	With frm1.vspdData3
		.Row = Row
		Select Case Col
			Case C_ResourceCd3
				.Col = Col
				intIndex = .Value
				.Col = C_ResourceDesc3
				.Value = intIndex
			Case C_ResourceDesc3
				.Col = Col
				intIndex = .Value
				.Col = C_ResourceCd3
				.Value = intIndex

			Case C_WkLossCd3
				.Col = Col
				intIndex = .Value
				.Col = C_WkLossDesc3
				.Value = intIndex
			Case C_WkLossDesc3
				.Col = Col
				intIndex = .Value
				.Col = C_WkLossCd3
				.Value = intIndex
		End Select
    End With
End Sub

'==========================================================================================
'   Event Name :vspddata3_ComboSelChange
'   Event Desc :Combo Change Event
'==========================================================================================
Sub vspdData5_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim strProdtOrderNo, strOprNo
	Dim LngFindRow

	ggoSpread.Source = frm1.vspdData5
	ggoSpread.UpdateRow Row

	With frm1.vspdData5
		.Row = Row
		Select Case Col
			Case C_ActCd5
				.Col = Col
				intIndex = .Value
				.Col = C_ActCdDesc5
				.Value = intIndex
			Case C_ActCdDesc5
				.Col = Col
				intIndex = .Value
				.Col = C_ActCd5
				.Value = intIndex
		End Select
    End With
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
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    If gActiveSpdSheet = "A" Then Call InitSpreadComboBox()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.ReOrderingSpreadData
    If gActiveSpdSheet = "A" Then Call InitData(1)
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
    Dim	IntRetCD

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

	Select Case gSelframeFlg
		Case TAB1
			gMouseClickStatus = "SP2C"
			If FncQueryTAB1 = False Then
				Exit Function
			End If
		Case TAB2
			If FncQueryTAB2 = False Then
				Exit Function
			End If
	End Select

	Call DefaultFormValue

    FncQuery = True																'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncQueryTAB1
' Function Desc :
'========================================================================================
Function DefaultFormValue()

	With frm1
		.fpDoubleSingle1.Value = 0
		.fpDoubleSingle2.Value = 0
		.fpDoubleSingle3.Value = 0
		.fpDoubleSingle4.Value = 0
		.fpDoubleSingle5.Value = 0
		.fpDoubleSingle6.Value = 0
		.fpDoubleSingle7.Value = 0
		.fpDoubleSingle8.Value = 0
		.fpDoubleSingle9.Value = 0
		.fpDoubleSingle10.Value = 0
		.fpDoubleSingle11.Value = 0
		.fpDoubleSingle12.Value = 0
		.fpDoubleSingle13.Value = 0
		.fpDoubleSingle14.Value = 0
		.fpDoubleSingle15.Value = 0
	End With

	lgBlnFlgChgValue = False

End Function

'========================================================================================
' Function Name : FncQueryTAB1
' Function Desc :
'========================================================================================
Function FncQueryTAB1()
	FncQueryTAB1 = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If CheckChange = True Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")		'��: Display Message(There is no changed data.)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call InitVariables															'��: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If																	'��: Query db data

	FncQueryTAB1 = True
End Function

'========================================================================================
' Function Name : FncQueryTAB2
' Function Desc :
'========================================================================================
Function FncQueryTAB2()
	FncQueryTAB2 = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If CheckChange = True Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")		'��: Display Message(There is no changed data.)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

'    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field

	ggoSpread.Source = frm1.vspdData4
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData5
    ggoSpread.ClearSpreadData

    Call InitVariables															'��: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	gSelframeFlg = TAB2
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If																	'��: Query db data

	FncQueryTAB2 = True
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
    Dim	IntRetCD

    FncSave = False												'��: Processing is NG

    Err.Clear													'��: Protect system from crashing

'    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer

    If CheckChange = False And lgBlnFlgChgValue = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If

	If DefaultCheck = False Then Exit Function

	' �ܾ��ð� ��� 
	Call CalcWorkTimeTotal()

	'-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function				                                  '��: Save db data

    FncSave = True                                            '��: Processing is OK

End Function

'========================================================================================
' Function Name : CheckChange
' Function Desc : �ʼ��� Ȯ�� 
'========================================================================================
Function DefaultCheck()
	DefaultCheck = False

	Select Case gSelframeFlg
		Case TAB1
			ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer
			If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
			   Exit Function
			End If
			ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer
			If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
			   Exit Function
			End If
		Case TAB2
			ggoSpread.Source = frm1.vspdData4							'��: Preset spreadsheet pointer
			If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
			   Exit Function
			End If
			ggoSpread.Source = frm1.vspdData5							'��: Preset spreadsheet pointer
			If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
			   Exit Function
			End If
	End Select
	DefaultCheck = True
End Function

'========================================================================================
' Function Name : CheckChange
' Function Desc : ����� ����Ÿ Ȯ�� 
'========================================================================================
Function CheckChange()
	CheckChange = True

	Select Case gSelframeFlg
		Case TAB1
			ggoSpread.Source = frm1.vspdData1
			If ggoSpread.SSCheckChange = True Then
				Exit Function
			End If
			ggoSpread.Source = frm1.vspdData2
			If ggoSpread.SSCheckChange = True Then
				Exit Function
			End If
			ggoSpread.Source = frm1.vspdData3
			If ggoSpread.SSCheckChange = True Then
				Exit Function
			End If
		Case TAB2
			ggoSpread.Source = frm1.vspdData4
			If ggoSpread.SSCheckChange = True Then
				Exit Function
			End If
			ggoSpread.Source = frm1.vspdData5
			If ggoSpread.SSCheckChange = True Then
				Exit Function
			End If
	End Select
	CheckChange = False
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
	Select Case gMouseClickStatus
		' TAB1
		Case "SPC"

		Case "SP2C"
			If frm1.vspdData2.MaxRows <= 0 Then Exit Function

			ggoSpread.Source = frm1.vspdData2
			ggoSpread.EditUndo
		Case "SP3C"
			If frm1.vspdData3.MaxRows <= 0 Then Exit Function

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.EditUndo
		' TAB2
		Case "SP4C"
			If frm1.vspdData4.MaxRows <= 0 Then Exit Function

			ggoSpread.Source = frm1.vspdData4
			ggoSpread.EditUndo
		Case "SP5C"
			If frm1.vspdData5.MaxRows <= 0 Then Exit Function

			ggoSpread.Source = frm1.vspdData5
			ggoSpread.EditUndo
	End Select

End Function

Function FncCancelVsp1()
	If frm1.vspdData1.MaxRows <= 0 Then Exit Function

	ggoSpread.Source = frm1.vspdData1
	frm1.vspdData1.Row = gRowVsp1
	Call SheetFocus(gRowVsp1, 1)
	ggoSpread.EditUndo

End Function

Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
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

Function FncInsertRow(ByVal pvRowCnt)

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		FncInsertRow = False
		Exit Function
	End If

	If frm1.txtWcCd.value= "" Then
		Call DisplayMsgBox("971012","X", "�۾���","X")
		frm1.txtWcCd.focus
		Set gActiveElement = document.activeElement
		FncInsertRow = False
		Exit Function
	End If

	If frm1.txtprodDt.value= "" Then
		Call DisplayMsgBox("971012","X", "�۾�����","X")
		frm1.txtprodDt.focus
		Set gActiveElement = document.activeElement
		FncInsertRow = False
		Exit Function
	End If

	Call FncInsertRow2(pvRowCnt)

End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc :
'========================================================================================
Function FncInsertRow2(ByVal pvRowCnt)
    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '��: Clear error status

    FncInsertRow2 = False                                                         '��: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If

	Select Case gMouseClickStatus
		Case "SPC", "SP2C"
			With frm1
				.vspdData2.ReDraw = False
				.vspdData2.Focus

				ggoSpread.Source = .vspdData2

				If frm1.vspdData2.selBlockRow = -1 Then
					ggoSpread.InsertRow 0, iIntReqRows
				Else
					ggoSpread.InsertRow , iIntReqRows
				End If

				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData2,.vspdData2.ActiveRow,.vspdData2.ActiveRow + iIntReqRows - 1,C_CurCd,C_SubconPrc, "C" ,"I","X","X")

				Call SetSpreadColor(.vspdData2.ActiveRow, .vspdData2.ActiveRow + iIntReqRows - 1, "Y")
				For iIntCnt = .vspdData2.ActiveRow To .vspdData2.ActiveRow + iIntReqRows - 1
					.vspdData2.Row  = iIntCnt
					.vspdData2.Col  = C_PlantCd2
					.vspdData2.Text = UCase(Trim(frm1.txtPlantCd.Value))
					.vspdData2.Col  = C_WcCd2
					.vspdData2.Text = UCase(Trim(frm1.txtWcCd.Value))
					.vspdData2.Col  = C_ReportDt2
					.vspdData2.Text = Trim(frm1.txtprodDt.Text)
					.vspdData2.Col  = C_ProdtOrderNo2
					.vspdData2.Text = Trim(frm1.KeyProdtOrderNo2.value)
					.vspdData2.Col  = C_OprNo2
					.vspdData2.Text = Trim(frm1.KeyOprNo2.Value)
					.vspdData2.Col  = C_Seq2
					.vspdData2.Text = .vspdData2.MaxRows' + 1
					.vspdData2.Col  = C_WorkMan2
					.vspdData2.Text = 0
					.vspdData2.Col  = C_WorkTime2
					.vspdData2.Text = ConvToTimeFormat(0)
				Next

				Call ProtectMilestone(0)

				.vspdData2.ReDraw = True

				Call SetSpreadColor(frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow, "N", "B")
			End With

		Case "SP3C"
			With frm1
				.vspdData3.ReDraw = False
				.vspdData3.Focus

				ggoSpread.Source = .vspdData3

				If frm1.vspdData3.selBlockRow = -1 Then
					ggoSpread.InsertRow 0, iIntReqRows
				Else
					ggoSpread.InsertRow , iIntReqRows
				End If

				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData3,.vspdData3.ActiveRow,.vspdData3.ActiveRow + iIntReqRows - 1,C_CurCd,C_SubconPrc, "C" ,"I","X","X")

				Call SetSpreadColor(.vspdData3.ActiveRow, .vspdData3.ActiveRow + iIntReqRows - 1, "Y")
				For iIntCnt = .vspdData3.ActiveRow To .vspdData3.ActiveRow + iIntReqRows - 1
					.vspdData3.Row  = iIntCnt
					.vspdData3.Col  = C_PlantCd3
					.vspdData3.Text = UCase(Trim(frm1.txtPlantCd.Value))
					.vspdData3.Col  = C_WcCd3
					.vspdData3.Text = UCase(Trim(frm1.txtWcCd.Value))
					.vspdData3.Col  = C_ReportDt3
					.vspdData3.Text = Trim(frm1.txtprodDt.Text)
					.vspdData3.Col  = C_ProdtOrderNo3
					.vspdData3.Text = Trim(frm1.KeyProdtOrderNo2.value)
					.vspdData3.Col  = C_OprNo3
					.vspdData3.Text = Trim(frm1.KeyOprNo2.Value)
					.vspdData3.Col  = C_SeqNo3
					.vspdData3.Text = .vspdData3.MaxRows' + 1
					.vspdData3.Col  = C_ItemCd3
					.vspdData3.Text = frm1.KeyItemCd.Value

					.vspdData3.Col  = C_StTime3
					.vspdData3.Text = ConvToTimeFormat(0)
					.vspdData3.Col  = C_EndTime3
					.vspdData3.Text = ConvToTimeFormat(0)
					.vspdData3.Col  = C_LossMan3
					.vspdData3.Text = 0
					.vspdData3.Col  = C_WkLossQty3
					.vspdData3.Text = ConvToTimeFormat(0)
				Next

				Call ProtectMilestone(0)

				.vspdData3.ReDraw = True

				Call SetSpreadColor(frm1.vspdData3.ActiveRow, frm1.vspdData3.ActiveRow, "N", "C")
			End With

		Case "SP4C"
			With frm1
				.vspdData4.ReDraw = False
				.vspdData4.Focus

				ggoSpread.Source = .vspdData4

				If frm1.vspdData4.selBlockRow = -1 Then
					ggoSpread.InsertRow 0, iIntReqRows
				Else
					ggoSpread.InsertRow , iIntReqRows
				End If

				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,.vspdData4.ActiveRow,.vspdData4.ActiveRow + iIntReqRows - 1,C_CurCd,C_SubconPrc, "C" ,"I","X","X")

				Call SetSpreadColor(.vspdData4.ActiveRow, .vspdData4.ActiveRow + iIntReqRows - 1, "Y")
				For iIntCnt = .vspdData4.ActiveRow To .vspdData4.ActiveRow + iIntReqRows - 1
					.vspdData4.Row  = iIntCnt
					.vspdData4.Col  = C_PlantCd4
					.vspdData4.Text = UCase(Trim(frm1.txtPlantCd.Value))
					.vspdData4.Col  = C_WcCd4
					.vspdData4.Text = UCase(Trim(frm1.txtWcCd.Value))
					.vspdData4.Col  = C_ReportDt4
					.vspdData4.Text = Trim(frm1.txtprodDt.Text)
					.vspdData4.Col  = C_SeqNo4
					.vspdData4.Text = .vspdData4.MaxRows' + 1
					.vspdData4.Col  = C_Time4
					.vspdData4.Text = ConvToTimeFormat(0)
				Next

				Call ProtectMilestone(0)

				.vspdData4.ReDraw = True

				Call SetSpreadColor(frm1.vspdData4.ActiveRow, frm1.vspdData4.ActiveRow, "N", "D")
			End With

		Case "SP5C"
			With frm1
				.vspdData5.ReDraw = False
				.vspdData5.Focus

				ggoSpread.Source = .vspdData5

				If frm1.vspdData5.selBlockRow = -1 Then
					ggoSpread.InsertRow 0, iIntReqRows
				Else
					ggoSpread.InsertRow , iIntReqRows
				End If

				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData5,.vspdData5.ActiveRow,.vspdData5.ActiveRow + iIntReqRows - 1,C_CurCd,C_SubconPrc, "C" ,"I","X","X")

				Call SetSpreadColor(.vspdData5.ActiveRow, .vspdData5.ActiveRow + iIntReqRows - 1, "Y")
				For iIntCnt = .vspdData5.ActiveRow To .vspdData5.ActiveRow + iIntReqRows - 1
					.vspdData5.Row  = iIntCnt
					.vspdData5.Col  = C_PlantCd5
					.vspdData5.Text = UCase(Trim(frm1.txtPlantCd.Value))
					.vspdData5.Col  = C_WcCd5
					.vspdData5.Text = UCase(Trim(frm1.txtWcCd.Value))
					.vspdData5.Col  = C_ReportDt5
					.vspdData5.Text = Trim(frm1.txtprodDt.Text)
					.vspdData5.Col  = C_SeqNo5
					.vspdData5.Text = .vspdData5.MaxRows' + 1
				Next

				Call ProtectMilestone(0)

				.vspdData5.ReDraw = True

				Call SetSpreadColor(frm1.vspdData5.ActiveRow, frm1.vspdData5.ActiveRow, "N", "E")
			End With
	End Select

    If Err.number = 0 Then
       FncInsertRow2 = True                                                          '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow()
    Dim lDelRows
    Dim iIntCnt

    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
	Select Case	gMouseClickStatus
		Case "SPC", "SP2C"
			If frm1.vspdData2.maxrows < 1 Then Exit Function

			ggoSpread.Source = frm1.vspdData2
		Case "SP3C"
			If frm1.vspdData3.maxrows < 1 Then Exit Function

			ggoSpread.Source = frm1.vspdData3
		Case "SP4C"
			If frm1.vspdData4.maxrows < 1 Then Exit Function

			ggoSpread.Source = frm1.vspdData4
		Case "SP5C"
			If frm1.vspdData5.maxrows < 1 Then Exit Function

			ggoSpread.Source = frm1.vspdData5
	End Select

    lDelRows = ggoSpread.DeleteRow

	If gMouseClickStatus = "SP2C" or gMouseClickStatus = "SP3C" Then
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.UpdateRow frm1.vspdData1.ActiveRow
	End If
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
    Call parent.FncExport(parent.C_SINGLEMULTI)									'��: ȭ�� ���� 
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
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                               '��:ȭ�� ����, Tab ���� 
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
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()												'��: ���� ������ ���� ���� 

End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� :
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : TAB1 ��ȸ �� Scroll
'========================================================================================
Function DbQuery()

    DbQuery = False

    Call LayerShowHide(1)

	Dim	strVal

	Select Case gSelframeFlg
		Case TAB1
			If lgIntFlgMode = parent.OPMD_UMODE Then
				strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
				strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
				strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
				strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
				strVal = strVal & "&txtprodDt=" & Trim(frm1.hProdDt.value)
				strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)
			Else
				strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
				strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
				strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
				strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
				strVal = strVal & "&txtprodDt=" & Trim(frm1.txtprodDt.text)
				strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
			End If

			Call RunMyBizASP(MyBizASP, strVal)
			DbQuery = True

		Case TAB2
			If lgIntFlgMode = parent.OPMD_UMODE Then
				strVal = BIZ_PGM_QRY4_ID & "?txtMode=" & parent.UID_M0001
				strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
				strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
				strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
				strVal = strVal & "&txtprodDt=" & Trim(frm1.hProdDt.value)
				strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)
			Else
				strVal = BIZ_PGM_QRY4_ID & "?txtMode=" & parent.UID_M0001
				strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
				strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
				strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
				strVal = strVal & "&txtprodDt=" & Trim(frm1.txtprodDt.text)
				strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
			End If

			Call RunMyBizASP(MyBizASP, strVal)
			DbQuery = True
	End Select
End Function

'========================================================================================
' Function Name : DbQueryForm
' Function Desc :
'========================================================================================
Function DbQueryForm()

    DbQueryForm = False

'    Call LayerShowHide(1)

	Dim	strVal

		strVal = BIZ_PGM_QRY0_ID & "?txtMode0=" & parent.UID_M0001
'		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
'		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtprodDt=" & Trim(frm1.txtprodDt.text)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)

    Call RunMyBizASP(MyBizASP, strVal)

    DbQueryForm = True

End Function

'========================================================================================
' Function Name : DbQueryOkForm
' Function Desc :
'========================================================================================
Function DbQueryOkForm(Byval LngMaxRow)
	lgBlnFlgChgValue = False

	If lgIntFlgMode = parent.OPMD_UMODE Then
		frm1.vspdData1.Col = C_ProdtOrderNo
		frm1.vspdData1.Row = 1
		frm1.KeyProdtOrderNo2.value = Trim(frm1.vspdData1.Text)

		frm1.vspdData1.Col = C_OprNo
		frm1.vspdData1.Row = 1
		frm1.KeyOprNo2.value = Trim(frm1.vspdData1.Text)

		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement

		If DbQuery2 = False Then
			Call RestoreToolBar()
			Exit Function
		End If

'		lgOldRow1 = 1

	End If

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(Byval LngMaxRow)
    '-----------------------
    'Reset variables area
    '-----------------------

	Select Case gSelframeFlg
		Case TAB1
			lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode		

			If DbQueryForm = False Then
				Call RestoreToolBar()
		'		Exit Function
			End If
			' Resource Combo
			Call InitSpreadComboBox2()

			Call SetToolbar("11101111001011")										'��: ��ư ���� ���� 

		Case TAB2
			lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
			Call SetActiveCell(frm1.vspdData4,1,1,"M","X","X")
			Set gActiveElement = document.activeElement
			lgBlnFlgChgValue = False
			Call SetToolbar("11101111001011")										'��: ��ư ���� ���� 
	End Select

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
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
	strVal = strVal & "&txtprodDt=" & Trim(frm1.txtprodDt.Text)
	strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.KeyProdtOrderNo2.value)
	strVal = strVal & "&txtOprNo=" & Trim(frm1.KeyOprNo2.value)

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery2 = True

End Function

Function DbQuery3()
	DbQuery3 = False

	strVal = BIZ_PGM_QRY3_ID & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
	strVal = strVal & "&txtprodDt=" & Trim(frm1.txtprodDt.Text)
	strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.KeyProdtOrderNo2.value)
	strVal = strVal & "&txtOprNo=" & Trim(frm1.KeyOprNo2.value)

    Call RunMyBizASP(MyBizASP, strVal)

	DbQuery3 = True
End Function

'========================================================================================
' Function Name : DbQuery2Ok
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
'========================================================================================
Function DbQuery2Ok()

	If DbQuery3 = False Then
		Call RestoreToolBar()
		Exit Function
	End If

End Function

'========================================================================================
' Function Name : fnConvTime(ihour)
' Function Desc : 
'========================================================================================
Function fnConvTime(ihour)
	
	fnConvTime = ihour * 3600
	
End Function


'========================================================================================
' Function Name : DbSaveTab1()
' Function Desc : This function is data query and display
'========================================================================================
Function DbSaveTab1()
    Dim IntRows
    Dim strVal, strVal1, strVal2, strVal3
	Dim strDel, strDel1, strDel2, strDel3

	Dim iColSep, iRowSep

   	Dim TmpBufferVal, TmpBufferDel
   	Dim TmpBufferVal1, TmpBufferDel1
   	Dim TmpBufferVal2, TmpBufferDel2
   	Dim TmpBufferVal3, TmpBufferDel3

   	Dim iTotalStrVal, iTotalStrDel
   	Dim iTotalStrVal1, iTotalStrDel1
   	Dim iTotalStrVal2, iTotalStrDel2
   	Dim iTotalStrVal3, iTotalStrDel3

	Dim iValCnt, iDelCnt

	DbSaveTab1 = False
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep

    lGrpCnt = 1
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
    ReDim TmpBufferVal1(0) : ReDim TmpBufferDel1(0)
    ReDim TmpBufferVal2(0) : ReDim TmpBufferDel2(0)
    ReDim TmpBufferVal3(0) : ReDim TmpBufferDel3(0)

	'// txtSpread	: Master Form Data
	'// txtSpread1	: vspdData1 Data
	'// txtSpread2	: vspdData2 Data
	'// txtSpread3	: vspdData3 Data
'msgbox "DbSaveTab1---------Form"
	With frm1
		strVal = ""

		strVal = strVal & Trim(.fpDoubleSingle1.Value) & iColSep
		strVal = strVal & fnConvTime(Trim(.fpDoubleSingle5.Value)) & iColSep
		strVal = strVal & Trim(.fpDoubleSingle9.Value) & iColSep
		strVal = strVal & fnConvTime(Trim(.fpDoubleSingle12.Value)) & iColSep

		strVal = strVal & Trim(.fpDoubleSingle2.Value) & iColSep
		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle6.Value)) & iColSep
		strVal = strVal & Trim(.fpDoubleSingle10.Value) & iColSep
		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle13.Value)) & iColSep

		strVal = strVal & Trim(.fpDoubleSingle3.Value) & iColSep
		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle7.Value)) & iColSep
		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle11.Value)) & iColSep
		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle14.Value)) & iColSep

		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle4.Value)) & iColSep
		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle8.Value)) & iColSep
		strVal = strVal & ConvToSec(Trim(.fpDoubleSingle15.Value)) & iRowSep

		ReDim Preserve TmpBufferVal(iValCnt)
		TmpBufferVal(iValCnt) = strVal
		iValCnt = iValCnt + 1
		lGrpCnt = lGrpCnt + 1

		iTotalStrVal = Join(TmpBufferVal, "")

		.txtSpread.value = iTotalStrVal
	End With
'msgbox "DbSaveTab1---------vspdData1"
	' vspdData1
	With frm1.vspdData1
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					strVal1 = ""
					If .Text = ggoSpread.InsertFlag Then
						strVal1 = strVal1 & "C" & iColSep & IntRows & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal1 = strVal1 & "U" & iColSep & IntRows & iColSep				'��: U=Update
					End If
					.Col = C_ProdtOrderNo					' 2
					strVal1 = strVal1 & Trim(.Text) & iColSep
					.Col = C_OprNo							' 3
					strVal1 = strVal1 & Trim(.Text) & iColSep

					.Col = C_StdTime						' 4
					strVal1 = strVal1 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_IncTime						' 5
					strVal1 = strVal1 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_DescTime						' 6
					strVal1 = strVal1 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_OtTime							' 7
					strVal1 = strVal1 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_WkTime							' 8
					strVal1 = strVal1 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_EtcTime						' 9
					strVal1 = strVal1 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_WkLossTime						' 10
					strVal1 = strVal1 & ConvToSec(Trim(.Text)) & iRowSep

					ReDim Preserve TmpBufferVal1(iValCnt)
					TmpBufferVal1(iValCnt) = strVal1
					iValCnt = iValCnt + 1
					lGrpCnt = lGrpCnt + 1
				Case ggoSpread.DeleteFlag
					strDel1 = ""
					strDel1 = strDel1 & "D" & iColSep & IntRows & iColSep				'��: D=Delete
'					.Col = C_PlantCd2						'2
			        strDel1 = strDel1 & UCase(Trim(frm1.txtPlantCd.Value)) & iColSep
'					.Col = C_WcCd2							'3
					strDel1 = strDel1 & UCase(Trim(frm1.txtWcCd.Value)) & iColSep
'					.Col = C_ReportDt2						'4
					strDel1 = strDel1 & Trim(frm1.txtprodDt.Text) & iColSep
					.Col = C_ProdtOrderNo					'5
					strDel1 = strDel1 & Trim(.Text) & iColSep
					.Col = C_OprNo							'6
					strDel1 = strDel1 & Trim(.Text) & iRowSep
					ReDim Preserve TmpBufferDel1(iDelCnt)
					TmpBufferDel1(iDelCnt) = strDel1
					iDelCnt = iDelCnt + 1
					lGrpCnt = lGrpCnt + 1
			End Select
		Next
		iTotalStrVal1 = Join(TmpBufferVal1, "")
		iTotalStrDel1 = Join(TmpBufferDel1, "")

		frm1.txtSpread1.value = iTotalStrDel1 & iTotalStrVal1
	End With
'msgbox "DbSaveTab1---------vspdData2"
	' vspdData2
	With frm1.vspdData2
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					strVal2 = ""
					If .Text = ggoSpread.InsertFlag Then
						strVal2 = strVal2 & "C" & iColSep & IntRows & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal2 = strVal2 & "U" & iColSep & IntRows & iColSep				'��: U=Update
					End If

					.Col = C_ProdtOrderNo2					' 5
					strVal2 = strVal2 & Trim(.Text) & iColSep
					.Col = C_OprNo2							' 6
					strVal2 = strVal2 & Trim(.Text) & iColSep
					.Col = C_Seq2							' 7
					strVal2 = strVal2 & Trim(.Text) & iColSep
					.Col = C_ResourceCd2					' 8
					strVal2 = strVal2 & Trim(.Text) & iColSep
					.Col = C_WorkType2						' 9
					strVal2 = strVal2 & Trim(.Text) & iColSep
					.Col = C_WorkMan2						' 10
					strVal2 = strVal2 & Trim(.Text) & iColSep
					.Col = C_WorkTime2						' 11
					strVal2 = strVal2 & ConvToSec(Trim(.Text)) & iRowSep

					ReDim Preserve TmpBufferVal2(iValCnt)
					TmpBufferVal2(iValCnt) = strVal2
					iValCnt = iValCnt + 1
					lGrpCnt = lGrpCnt + 1

				Case ggoSpread.DeleteFlag
					strDel2 = ""
					strDel2 = strDel2 & "D" & iColSep & IntRows & iColSep				'��: D=Delete
					.Col = C_PlantCd2						'2
			        strDel2 = strDel2 & Trim(.Text) & iColSep
					.Col = C_WcCd2							'3
					strDel2 = strDel2 & Trim(.Text) & iColSep
					.Col = C_ReportDt2						'4
					strDel2 = strDel2 & Trim(.Text) & iColSep
					.Col = C_ProdtOrderNo2					'5
					strDel2 = strDel2 & Trim(.Text) & iColSep
					.Col = C_OprNo2							'6
					strDel2 = strDel2 & Trim(.Text) & iColSep
					.Col = C_Seq2							'7
					strDel2 = strDel2 & Trim(.Text) & iColSep
					.Col = C_ResourceCd2					'8
					strDel2 = strDel2 & Trim(.Text) & iRowSep

					ReDim Preserve TmpBufferDel2(iDelCnt)
					TmpBufferDel2(iDelCnt) = strDel2
					iDelCnt = iDelCnt + 1
					lGrpCnt = lGrpCnt + 1
			End Select
		Next
		iTotalStrVal2 = Join(TmpBufferVal2, "")
		iTotalStrDel2 = Join(TmpBufferDel2, "")

		frm1.txtSpread2.value = iTotalStrDel2 & iTotalStrVal2
	End With
'msgbox "DbSaveTab1---------vspdData3"
	' vspdData3
	With frm1.vspdData3
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					strVal3 = ""
					If .Text = ggoSpread.InsertFlag Then
						strVal3 = strVal3 & "C" & iColSep & IntRows & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal3 = strVal3 & "U" & iColSep & IntRows & iColSep				'��: U=Update
					End If
					.Col = C_PlantCd3						' 2
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_WcCd3							' 3
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_ReportDt3						' 4
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_ProdtOrderNo3					' 5
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_OprNo3							' 6
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_SeqNo3							' 7
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_ResourceCd3					' 8
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_ItemCd3						' 9
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_StTime3						' 10
					strVal3 = strVal3 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_EndTime3						' 11
					strVal3 = strVal3 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_LossMan3						' 12
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_WkLossQty3						' 13
					strVal3 = strVal3 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_WkLossCd3						' 14
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_WkLossType3					' 15
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_RtDeptCd3						' 16
					strVal3 = strVal3 & Trim(.Text) & iColSep
					.Col = C_Notes3							' 17
					strVal3 = strVal3 & Trim(.Text) & iRowSep

					ReDim Preserve TmpBufferVal3(iValCnt)
					TmpBufferVal3(iValCnt) = strVal3
					iValCnt = iValCnt + 1
					lGrpCnt = lGrpCnt + 1

				Case ggoSpread.DeleteFlag
					strDel3 = ""
					strDel3 = strDel3 & "D" & iColSep & IntRows & iColSep				'��: D=Delete
					.Col = C_PlantCd3						'2
			        strDel3 = strDel3 & Trim(.Text) & iColSep
					.Col = C_WcCd3							'3
					strDel3 = strDel3 & Trim(.Text) & iColSep
					.Col = C_ReportDt3						'4
					strDel3 = strDel3 & Trim(.Text) & iColSep
					.Col = C_ProdtOrderNo3					'5
					strDel3 = strDel3 & Trim(.Text) & iColSep
					.Col = C_OprNo3
					strDel3 = strDel3 & Trim(.Text) & iColSep
					.Col = C_SeqNo3							'6
					strDel3 = strDel3 & Trim(.Text) & iColSep
					.Col = C_ResourceCd3					'7
					strDel3 = strDel3 & Trim(.Text) & iRowSep

					ReDim Preserve TmpBufferDel3(iDelCnt)
					TmpBufferDel3(iDelCnt) = strDel3
					iDelCnt = iDelCnt + 1
					lGrpCnt = lGrpCnt + 1
			End Select
		Next
		iTotalstrVal3 = Join(TmpBufferVal3, "")
		iTotalStrDel3 = Join(TmpBufferDel3, "")

		frm1.txtSpread3.value = iTotalStrDel3 & iTotalstrVal3
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 

	DbSaveTab1 = True

End Function

'========================================================================================
' Function Name : UsrDbSave()
' Function Desc :
'========================================================================================
Function UsrDbSave()

End Function

'========================================================================================
' Function Name : DbSaveTab2()
' Function Desc : This function is data query and display
'========================================================================================
Function DbSaveTab2()
    Dim IntRows
    Dim strVal4, strVal5
	Dim strDel4, strDel5

	Dim iColSep, iRowSep

   	Dim TmpBufferVal4, TmpBufferDel4
   	Dim TmpBufferVal5, TmpBufferDel5
   	Dim iTotalStrVal4, iTotalStrDel4
   	Dim iTotalStrVal5, iTotalStrDel5
	Dim iValCnt, iDelCnt

	DbSaveTab2 = False
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep

    lGrpCnt = 1
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal4(0) : ReDim TmpBufferDel4(0)
    ReDim TmpBufferVal5(0) : ReDim TmpBufferDel5(0)

		'// txtSpread4	: vspdData4 Data
		'// txtSpread5	: vspdData5 Data

'// txtSpread4	: vspdData4 Data
	With frm1.vspdData4
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					strVal4 = ""
					If .Text = ggoSpread.InsertFlag Then
						strVal4 = strVal4 & "C" & iColSep & IntRows & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal4 = strVal4 & "U" & iColSep	& IntRows & iColSep				'��: U=Update
					End If
					.Col = C_PlantCd4						' 2
					strVal4 = strVal4 & Trim(.Text) & iColSep
					.Col = C_WcCd4							' 3
					strVal4 = strVal4 & Trim(.Text) & iColSep
					.Col = C_ReportDt4						' 4
					strVal4 = strVal4 & Trim(.Text) & iColSep
					.Col = C_SeqNo4							' 6
					strVal4 = strVal4 & Trim(.Text) & iColSep
					.Col = C_EmpNo4							' 7
					strVal4 = strVal4 & Trim(.Text) & iColSep
					.Col = C_Time4							' 8
					strVal4 = strVal4 & ConvToSec(Trim(.Text)) & iColSep
					.Col = C_Notes4							' 9
					strVal4 = strVal4 & Trim(.Text) & iRowSep

					ReDim Preserve TmpBufferVal4(iValCnt)

					TmpBufferVal4(iValCnt) = strVal4
					iValCnt = iValCnt + 1
					lGrpCnt = lGrpCnt + 1

				Case ggoSpread.DeleteFlag
					strDel4 = ""
					strDel4 = strDel4 & "D" & iColSep	& IntRows & iColSep				'��: D=Delete
					.Col = C_PlantCd4						' 2
					strDel4 = strDel4 & Trim(.Text) & iColSep
					.Col = C_WcCd4							' 3
					strDel4 = strDel4 & Trim(.Text) & iColSep
					.Col = C_ReportDt4						' 4
					strDel4 = strDel4 & Trim(.Text) & iColSep
					.Col = C_SeqNo4							' 6
					strDel4 = strDel4 & Trim(.Text) & iRowSep

					ReDim Preserve TmpBufferDel4(iDelCnt)

					TmpBufferDel4(iDelCnt) = strDel4
					iDelCnt = iDelCnt + 1
					lGrpCnt = lGrpCnt + 1
			End Select
		Next
		iTotalStrVal4 = Join(TmpBufferVal4, "")
		iTotalStrDel4 = Join(TmpBufferDel4, "")

		frm1.txtMaxRows.value = lGrpCnt-1
		frm1.txtSpread4.value = iTotalStrDel4 & iTotalStrVal4
	End With

'// txtSpread5	: vspdData5 Data
	With frm1.vspdData5
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					strVal5 = ""
					If .Text = ggoSpread.InsertFlag Then
						strVal5 = strVal5 & "C" & iColSep & IntRows & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal5 = strVal5 & "U" & iColSep	& IntRows & iColSep				'��: U=Update
					End If

					.Col = C_PlantCd5							'2
					strVal5 = strVal5 & Trim(.Text) & iColSep
					.Col = C_WcCd5								'3
					strVal5 = strVal5 & Trim(.Text) & iColSep
					.Col = C_ReportDt5							'4
					strVal5 = strVal5 & Trim(.Text) & iColSep
					.Col = C_SeqNo5								'6
					strVal5 = strVal5 & Trim(.Text) & iColSep
					.Col = C_EmpNo5								'7
					strVal5 = strVal5 & Trim(.Text) & iColSep
					.Col = C_ActCd5								'8
					strVal5 = strVal5 & Trim(.Text) & iColSep
					.Col = C_Notes5								'9
					strVal5 = strVal5 & Trim(.Text) & iRowSep

					ReDim Preserve TmpBufferVal5(iValCnt)

					TmpBufferVal5(iValCnt) = strVal5
					iValCnt = iValCnt + 1
					lGrpCnt = lGrpCnt + 1

				Case ggoSpread.DeleteFlag
					strDel5 = ""
					strDel5 = strDel5 & "D" & iColSep	& IntRows & iColSep				'��: D=Delete
					.Col = C_PlantCd5						' 2
					strDel5 = strDel5 & Trim(.Text) & iColSep
					.Col = C_WcCd5							' 3
					strDel5 = strDel5 & Trim(.Text) & iColSep
					.Col = C_ReportDt5						' 4
					strDel5 = strDel5 & Trim(.Text) & iColSep
					.Col = C_SeqNo5							' 6
					strDel5 = strDel5 & Trim(.Text) & iRowSep

					ReDim Preserve TmpBufferDel5(iDelCnt)

					TmpBufferDel5(iDelCnt) = strDel5
					iDelCnt = iDelCnt + 1
					lGrpCnt = lGrpCnt + 1

			End Select
		Next
		iTotalstrVal5 = Join(TmpBufferVal5, "")
		iTotalstrDel5 = Join(TmpBufferDel5, "")

		frm1.txtMaxRows.value = lGrpCnt-1
		frm1.txtSpread5.value = iTotalstrDel5 & iTotalstrVal5
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 

	DbSaveTab2 = True

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()

    DbSave = False														'��: Processing is NG
	LayerShowHide(1)

    'On Error Resume Next                                                   '��: Protect system from crashing

    With frm1
		.txtMode.Value			= parent.UID_M0002							'��: ���� ���� 
		.txtFlgMode.Value		= lgIntFlgMode								'��: �ű��Է�/���� ���� 
		.txtUpdtUserId.value	= parent.gUsrID
		.txtInsrtUserId.value	= parent.gUsrID
	End With

	Select Case gSelframeFlg
		Case TAB1
			If DbSaveTab1 = False Then Exit Function				                                  '��: Save db data
		Case TAB2
			If DbSaveTab2 = False Then Exit Function				                                  '��: Save db data
	End Select

	DbSave = True                                                           '��: Processing is OK

End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 

	Call InitVariables

    frm1.vspdData1.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
    frm1.vspdData3.MaxRows = 0

	frm1.KeyOprNo2.Value        = ""
	frm1.KeyProdtOrderNo2.value = ""
	frm1.KeyProdtOrderNo3.value = ""

	frm1.txtSpread.Value = ""
	frm1.txtSpread1.Value = ""
	frm1.txtSpread2.Value = ""
	frm1.txtSpread3.Value = ""

	lgBlnFlgChgValue = False
    Call MainQuery()

'	IsOpenPop = False
End Function

Function DbSaveFormOk()													'��: ���� ������ ���� ���� 
	Call InitVariables

    frm1.vspdData4.MaxRows = 0
    frm1.vspdData5.MaxRows = 0

	frm1.txtSpread4.Value = ""
	frm1.txtSpread5.Value = ""

	lgBlnFlgChgValue = False
    Call FncQueryTAB2()

'	IsOpenPop = False
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim	ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function


'==============================================================================
' Function :
' Description : Form event
'==============================================================================
Sub fpDoubleSingle1_Change()	' �����ð� 
	lgBlnFlgChgValue = True
End Sub

Sub fpDoubleSingle5_Change()
	lgBlnFlgChgValue = True
End Sub

Sub fpDoubleSingle9_Change()
	lgBlnFlgChgValue = True
End Sub

Sub fpDoubleSingle12_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line
'========================================================================================================
Function OpenEmp(ByVal strCode, ByVal gubun)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(strCode)						' Code Condition
	arrParam(1) = ""'frm1.txtName1.value			' Name Cindition
    arrParam(2) = lgUsrIntCd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	Select Case gubun
		Case "vspdData4"
			If arrRet(0) = "" Then
				frm1.vspdData4.focus
				Exit Function
			Else
				Call SetEmp(arrRet, gubun)
			End If
		Case "vspdData5"
			If arrRet(0) = "" Then
				frm1.vspdData5.focus
				Exit Function
			Else
				Call SetEmp(arrRet, gubun)
			End If
	End Select

End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SetEmp(Byval arrRet, Byval gubun)
	Select Case gubun
		Case "vspdData4"
			 With frm1.vspdData4
				.Col	= C_EmpNo4
				.Text	= arrRet(0)
				.Col	= C_EmpNm4
				.Text	= arrRet(1)
			 End With
		Case "vspdData5"
			 With frm1.vspdData5
				.Col	= C_EmpNo5
				.Text	= arrRet(0)
				.Col	= C_EmpNm5
				.Text	= arrRet(1)
			 End With
	 End Select
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenWorkDailyRef()
    Dim IntRetCd, strVal

	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "x", "x", "x")
			Exit Function
		End If
	End If

	If lgBlnFlgChgValue = True Then
		Call DisplayMsgBox("189217", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	If frm1.txtprodDt.Text = "" Then
		Call DisplayMsgBox("971012","X", "�۾�����","X")
		frm1.txtprodDt.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	If frm1.txtWcCd.Value = "" Then
		Call DisplayMsgBox("971012","X", "�۾���","X")
		frm1.txtWcCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	WriteCookie "txtPlantCd", UCase(Trim(frm1.txtPlantCd.value))
	WriteCookie "txtPlantNm", Trim(frm1.txtPlantNm.value)
	WriteCookie "txtprodDt", frm1.txtprodDt.Text
	WriteCookie "txtWcCd", UCase(Trim(frm1.txtWcCd.value))
	WriteCookie "txtWcNm", Trim(frm1.txtWcNm.value)
	WriteCookie "txtPGMID", "P4913MA1"
'	navigate BIZ_PGM_JUMPORDERRUN_ID
	PgmJump(BIZ_PGM_JUMPORDERRUN_ID)

End Function

'==============================================================================
' Function : ConvToTimeFormat
' Description : �ð� �������� ���� 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec

	Dim iVal2

	iVal2 = Fix(iVal)

	If iVal2 = 0 Then
		ConvToTimeFormat = "00:00:00"
	ElseIf iVal2 > 0 Then
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)

		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
	Else
		iVal2 = Replace(iVal2, "-", "")
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		ConvToTimeFormat = "-" & ConvToTimeFormat

	End If
End Function

'==============================================================================
' Function : ConvToSec()
' Description : ����ÿ� �� �ð� �����͵��� �ʷ� ȯ�� 
'==============================================================================
Function ConvToSec(ByVal Str)

	If Str = "" Then
		ConvToSec = 0
	ElseIf Str = "0" Then
		ConvToSec = 0
	ElseIf Len(Trim(Str)) = 8 Then
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
	ElseIf Len(Trim(Str)) = 9 Then
		Str = Replace(Str, "-", "")
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
		ConvToSec = "-" & ConvToSec
	End If

End Function