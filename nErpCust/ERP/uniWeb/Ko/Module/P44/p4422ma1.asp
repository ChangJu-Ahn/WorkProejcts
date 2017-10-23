<%@ LANGUAGE="VBSCRIPT" %> 
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4422ma1.asp
'*  4. Program Name         : ��������Ȳ��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/11/21
'*  8. Modified date(Last)  : 2003/05/26
'*  9. Modifier (First)     : Park, Bumsoo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                      

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY1_ID						= "p4422mb1.asp"		'��: Release Production Order Query ASP�� 
Const BIZ_PGM_QRY2_ID						= "p4422mb2.asp"		'��: Operation Query ASP�� 
Const BIZ_PGM_QRY3_ID						= "p4422mb3.asp"		'��: Reservation Query ASP�� 

'-------------------------------
' Column Constants : Spread 1 
'-------------------------------
Dim C_ItemCd					'= 1
Dim C_ItemNm					'= 2
Dim C_Spec						'
Dim C_WipQtyInOrderUnit			'= 3
Dim C_OrderUnit					'= 4
Dim C_PrevProdQtyInOrderUnit	'= 5
Dim C_PrevGoodQtyInOrderUnit	'= 6
Dim C_ProdQtyInOrderUnit		'= 7
Dim C_GoodQtyInOrderUnit		'= 8
Dim C_BadQtyInOrderUnit			'= 9
Dim C_WipQtyInBaseUnit			'= 10
Dim C_BaseUnit					'= 11
Dim C_PrevProdQtyInBaseUnit		'= 12
Dim C_PrevGoodQtyInBaseUnit		'= 13
Dim C_ProdQtyInBaseUnit			'= 14
Dim C_GoodQtyInBaseUnit			'= 15
Dim C_BadQtyInBaseUnit			'= 16
Dim C_ItemGroupCd				'= 17
Dim C_ItemGroupNm				'= 18

'-------------------------------
' Column Constants : Spread 2 
'-------------------------------
Dim C_WCCd						'= 1
Dim C_WcNm						'= 2
Dim C_WipQtyInOrderUnit1		'= 3
Dim C_OrderUnit1				'= 4
Dim C_PrevProdQtyInOrderUnit1	'= 5
Dim C_PrevGoodQtyInOrderUnit1	'= 6
Dim C_ProdQtyInOrderUnit1		'= 7
Dim C_GoodQtyInOrderUnit1		'= 8
Dim C_BadQtyInOrderUnit1		'= 9
Dim C_WipQtyInBaseUnit1			'= 10
Dim C_BaseUnit1					'= 11
Dim C_PrevProdQtyInBaseUnit1	'= 12
Dim C_PrevGoodQtyInBaseUnit1	'= 13
Dim C_ProdQtyInBaseUnit1		'= 14
Dim C_GoodQtyInBaseUnit1		'= 15
Dim C_BadQtyInBaseUnit1			'= 16

'-------------------------------
' Column Constants : Spread 3 
'-------------------------------
Dim C_ProdtOrderNo				'= 1
Dim C_OprNo						'= 2
Dim C_WipQtyInOrderUnit2		'= 3
Dim C_OrderUnit2				'= 4
Dim C_PrevProdQtyInOrderUnit2	'= 5
Dim C_PrevGoodQtyInOrderUnit2	'= 6
Dim C_ProdQtyInOrderUnit2		'= 7
Dim C_GoodQtyInOrderUnit2		'= 8
Dim C_BadQtyInOrderUnit2		'= 9
Dim C_WipQtyInBaseUnit2			'= 10
Dim C_BaseUnit2					'= 11
Dim C_PrevProdQtyInBaseUnit2	'= 12
Dim C_PrevGoodQtyInBaseUnit2	'= 13
Dim C_ProdQtyInBaseUnit2		'= 14
Dim C_GoodQtyInBaseUnit2		'= 15
Dim C_BadQtyInBaseUnit2			'= 16
Dim C_TrackingNo				'= 17
Dim C_OrderStatus				'= 18
Dim C_OrderStatusDesc			'= 19

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgStrPrevKey2		
Dim lgStrPrevKey3

Dim lgSortKey1
Dim lgSortKey2
Dim lgSortKey3

Dim lgOldRow1
Dim lgOldRow2
		
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop					 'Popup
Dim gSelframeFlg

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
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgOldRow1 = 0
	lgOldRow2 = 0
	
	lgSortKey1 = 1
	lgSortKey2 = 1
	lgSortKey3 = 1
	
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

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()     
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>

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
			ggoSpread.Spreadinit "V20030913", , Parent.gAllowDragDropSpread
			
			.vspdData1.Redraw = False
			
    		.vspdData1.MaxCols = C_ItemGroupNm + 1
			.vspdData1.MaxRows = 0
		
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit		C_ItemCd,					"ǰ��", 18
			ggoSpread.SSSetEdit		C_ItemNm,					"ǰ���", 25
			ggoSpread.SSSetEdit		C_Spec,						"�԰�", 25
			ggoSpread.SSSetFloat	C_WipQtyInOrderUnit,		"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_OrderUnit,				"��������", 8,,,3,2
			ggoSpread.SSSetFloat	C_PrevProdQtyInOrderUnit,	"��������������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQtyInOrderUnit,	"��������ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit,		"��������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit,		"��ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInOrderUnit,		"�ҷ�����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_WipQtyInBaseUnit,			"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_BaseUnit,					"���ش���", 8,,,3,2
			ggoSpread.SSSetFloat	C_PrevProdQtyInBaseUnit,	"��������������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQtyInBaseUnit,	"��������ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit,		"��������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit,		"��ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInBaseUnit,			"�ҷ�����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_ItemGroupCd,				"ǰ��׷�",	15
			ggoSpread.SSSetEdit		C_ItemGroupNm,				"ǰ��׷��", 30
		
			Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
		
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("A")
			
			.vspdData1.Redraw = True
			
		End If
				
		If pvSpdNo = "B" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 2 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			
			.vspdData2.Redraw = False
    
			.vspdData2.MaxCols = C_BadQtyInBaseUnit1 + 1
			.vspdData2.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit		C_WCCd,						"�۾���", 10
			ggoSpread.SSSetEdit		C_WCNm,						"�۾����", 20
			ggoSpread.SSSetFloat	C_WipQtyInOrderUnit1,		"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_OrderUnit1,				"��������", 8,,,3,2
			ggoSpread.SSSetFloat	C_PrevProdQtyInOrderUnit1,	"��������������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQtyInOrderUnit1,	"��������ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit1,		"��������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit1,		"��ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInOrderUnit1,		"�ҷ�����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_WipQtyInBaseUnit1,		"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_BaseUnit1,				"���ش���", 8,,,3,2
			ggoSpread.SSSetFloat	C_PrevProdQtyInBaseUnit1,	"��������������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQtyInBaseUnit1,	"��������ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit1,		"��������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit1,		"��ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInBaseUnit1,		"�ҷ�����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

			Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols, .vspdData2.MaxCols, True)

			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("B")
			
			.vspdData2.Redraw = True
			
		End If	
		
		If pvSpdNo = "C" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 3 Setting
			'-------------------------------------------
			
			ggoSpread.Source = .vspdData3
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			
			.vspdData3.Redraw = False
    
			.vspdData3.MaxCols = C_OrderStatusDesc + 1
			.vspdData3.MaxRows = 0
			
			Call GetSpreadColumnPos("C")
			
			ggoSpread.SSSetEdit		C_ProdtOrderNo,				"������ȣ", 18
			ggoSpread.SSSetEdit		C_OprNo,					"����", 8,,,3		
			ggoSpread.SSSetFloat	C_WipQtyInOrderUnit2,		"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_OrderUnit2,				"��������", 8,,,3,2
			ggoSpread.SSSetFloat	C_PrevProdQtyInOrderUnit2,	"��������������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQtyInOrderUnit2,	"��������ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit2,		"��������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit2,		"��ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInOrderUnit2,		"�ҷ�����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_WipQtyInBaseUnit2,		"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_BaseUnit2,				"���ش���", 8,,,3,2
			ggoSpread.SSSetFloat	C_PrevProdQtyInBaseUnit2,	"��������������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQtyInBaseUnit2,	"��������ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit2,		"��������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit2,		"��ǰ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQtyInBaseUnit2,		"�ҷ�����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_TrackingNo,				"Tracking No.", 25
			ggoSpread.SSSetCombo	C_OrderStatus,				"���û���", 10
			ggoSpread.SSSetCombo	C_OrderStatusDesc,			"���û���", 10
	
			Call ggoSpread.SSSetColHidden(.vspdData3.MaxCols, .vspdData3.MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_OrderStatus, C_OrderStatus, True)
			
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("C")
			
			.vspdData3.Redraw = True
			
		End If	
		
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	If pvSpdNo = "A" Then
		'-------------------------
		' Set Lock Prop :Spread 1 		
		'-------------------------	
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()

	ElseIf pvSpdNo = "B" Then
		'-------------------------
		' Set Lock Prop :Spread 2 		
		'-------------------------
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()

	ElseIf pvSpdNo = "C" Then
		'-------------------------
		' Set Lock Prop :Spread 3 		
		'-------------------------
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'================================== 2.2.6 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " and MINOR_CD <> " & FilterVar("OP", "''", "S") & " and MINOR_CD <> " & FilterVar("RL", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderStatus, lgF0, lgF1, Chr(11))
    ggoSpread.Source = frm1.vspdData3
	
	frm1.cboOrderStatus.value = "" 	 

End Sub

'==========================================  2.2.6 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc

End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData3
		For intRow = lngStartRow To .MaxRows

			.Row = intRow
			.col = C_OrderStatus
			intIndex = .value
			.Col = C_OrderStatusDesc
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
		C_ItemCd					= 1
		C_ItemNm					= 2
		C_Spec						= 3
		C_WipQtyInOrderUnit			= 4
		C_OrderUnit					= 5
		C_PrevProdQtyInOrderUnit	= 6
		C_PrevGoodQtyInOrderUnit	= 7
		C_ProdQtyInOrderUnit		= 8
		C_GoodQtyInOrderUnit		= 9
		C_BadQtyInOrderUnit			= 10
		C_WipQtyInBaseUnit			= 11
		C_BaseUnit					= 12
		C_PrevProdQtyInBaseUnit		= 13
		C_PrevGoodQtyInBaseUnit		= 14
		C_ProdQtyInBaseUnit			= 15
		C_GoodQtyInBaseUnit			= 16
		C_BadQtyInBaseUnit			= 17
		C_ItemGroupCd				= 18
		 C_ItemGroupNm				= 19
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 2 
		'-------------------------------
		C_WCCd						= 1
		C_WcNm						= 2
		C_WipQtyInOrderUnit1		= 3
		C_OrderUnit1				= 4
		C_PrevProdQtyInOrderUnit1	= 5
		C_PrevGoodQtyInOrderUnit1	= 6
		C_ProdQtyInOrderUnit1		= 7
		C_GoodQtyInOrderUnit1		= 8
		C_BadQtyInOrderUnit1		= 9
		C_WipQtyInBaseUnit1			= 10
		C_BaseUnit1					= 11
		C_PrevProdQtyInBaseUnit1	= 12
		C_PrevGoodQtyInBaseUnit1	= 13
		C_ProdQtyInBaseUnit1		= 14
		C_GoodQtyInBaseUnit1		= 15
		C_BadQtyInBaseUnit1			= 16
	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then	
		'-------------------------------
		' Column Constants : Spread 3 
		'-------------------------------
		C_ProdtOrderNo				= 1
		C_OprNo						= 2
		C_WipQtyInOrderUnit2		= 3
		C_OrderUnit2				= 4
		C_PrevProdQtyInOrderUnit2	= 5
		C_PrevGoodQtyInOrderUnit2	= 6
		C_ProdQtyInOrderUnit2		= 7
		C_GoodQtyInOrderUnit2		= 8
		C_BadQtyInOrderUnit2		= 9
		C_WipQtyInBaseUnit2			= 10
		C_BaseUnit2					= 11
		C_PrevProdQtyInBaseUnit2	= 12
		C_PrevGoodQtyInBaseUnit2	= 13
		C_ProdQtyInBaseUnit2		= 14
		C_GoodQtyInBaseUnit2		= 15
		C_BadQtyInBaseUnit2			= 16
		C_TrackingNo				= 17
		C_OrderStatus				= 18
		C_OrderStatusDesc			= 19
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
			
			C_ItemCd					= iCurColumnPos(1)
			C_ItemNm					= iCurColumnPos(2)
			C_Spec						= iCurColumnPos(3)
			C_WipQtyInOrderUnit			= iCurColumnPos(4)
			C_OrderUnit					= iCurColumnPos(5)
			C_PrevProdQtyInOrderUnit	= iCurColumnPos(6)
			C_PrevGoodQtyInOrderUnit	= iCurColumnPos(7)
			C_ProdQtyInOrderUnit		= iCurColumnPos(8)
			C_GoodQtyInOrderUnit		= iCurColumnPos(9)
			C_BadQtyInOrderUnit			= iCurColumnPos(10)
			C_WipQtyInBaseUnit			= iCurColumnPos(11)
			C_BaseUnit					= iCurColumnPos(12)
			C_PrevProdQtyInBaseUnit		= iCurColumnPos(13)
			C_PrevGoodQtyInBaseUnit		= iCurColumnPos(14)
			C_ProdQtyInBaseUnit			= iCurColumnPos(15)
			C_GoodQtyInBaseUnit			= iCurColumnPos(16)
			C_BadQtyInBaseUnit			= iCurColumnPos(17)
			C_ItemGroupCd				= iCurColumnPos(18)
			C_ItemGroupNm				= iCurColumnPos(19)
						
		Case "B"	
			ggoSpread.Source = frm1.vspdData2
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
 			C_WCCd						= iCurColumnPos(1)
			C_WcNm						= iCurColumnPos(2)
			C_WipQtyInOrderUnit1		= iCurColumnPos(3)
			C_OrderUnit1				= iCurColumnPos(4)
			C_PrevProdQtyInOrderUnit1	= iCurColumnPos(5)
			C_PrevGoodQtyInOrderUnit1	= iCurColumnPos(6)
			C_ProdQtyInOrderUnit1		= iCurColumnPos(7)
			C_GoodQtyInOrderUnit1		= iCurColumnPos(8)
			C_BadQtyInOrderUnit1		= iCurColumnPos(9)
			C_WipQtyInBaseUnit1			= iCurColumnPos(10)
			C_BaseUnit1					= iCurColumnPos(11)
			C_PrevProdQtyInBaseUnit1	= iCurColumnPos(12)
			C_PrevGoodQtyInBaseUnit1	= iCurColumnPos(13)
			C_ProdQtyInBaseUnit1		= iCurColumnPos(14)
			C_GoodQtyInBaseUnit1		= iCurColumnPos(15)
			C_BadQtyInBaseUnit1			= iCurColumnPos(16)
			
		Case "C"
			ggoSpread.Source = frm1.vspdData3
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ProdtOrderNo				= iCurColumnPos(1)
			C_OprNo						= iCurColumnPos(2)
			C_WipQtyInOrderUnit2		= iCurColumnPos(3)
			C_OrderUnit2				= iCurColumnPos(4)
			C_PrevProdQtyInOrderUnit2	= iCurColumnPos(5)
			C_PrevGoodQtyInOrderUnit2	= iCurColumnPos(6)
			C_ProdQtyInOrderUnit2		= iCurColumnPos(7)
			C_GoodQtyInOrderUnit2		= iCurColumnPos(8)
			C_BadQtyInOrderUnit2		= iCurColumnPos(9)
			C_WipQtyInBaseUnit2			= iCurColumnPos(10)
			C_BaseUnit2					= iCurColumnPos(11)
			C_PrevProdQtyInBaseUnit2	= iCurColumnPos(12)
			C_PrevGoodQtyInBaseUnit2	= iCurColumnPos(13)
			C_ProdQtyInBaseUnit2		= iCurColumnPos(14)
			C_GoodQtyInBaseUnit2		= iCurColumnPos(15)
			C_BadQtyInBaseUnit2			= iCurColumnPos(16)
			C_TrackingNo				= iCurColumnPos(17)
			C_OrderStatus				= iCurColumnPos(18)
			C_OrderStatusDesc			= iCurColumnPos(19)
				
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
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox �� 
	
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
		Call displaymsgbox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
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
	arrParam(2) = "12!MO"							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ���: From To�� �Է��� �� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
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
	arrParam(3) = "ST"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
    
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
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
	arrParam(3) = ""
	arrParam(4) = ""	
	
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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
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

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

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

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function    

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
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
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
	
	Call InitSpreadSheet("*")															'��: Setup the Spread sheet
	
	Call InitComboBox
	Call InitSpreadComboBox
	
	Call SetDefaultVal
	
	Call InitVariables																'��: Initializes local global variables
			
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11000000000011")
		
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtWCCd.focus
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

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================

Sub vspdData1_Click(ByVal Col , ByVal Row)
	
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
					
		lgOldRow1 = Row
 		
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

	Dim strProdOrdNo
	Dim strOprNo
	
	Call SetPopupMenuItemInf("0000111111")				'ȭ�麰 ���� 
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
 		
 		lgStrPrevKey3 = ""
			
		frm1.vspdData3.MaxRows = 0

		frm1.vspdData2.Row = frm1.vspdData2.ActiveRow		
		frm1.vspdData2.Col = C_WcCd
		frm1.KeyWcCd.value  = Trim(frm1.vspdData2.Text) 
		
		lgOldRow2 = frm1.vspdData2.Row
		
		If DbQuery3 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow2 <> Row Then
 		
			lgStrPrevKey3 = ""
			
			frm1.vspdData3.MaxRows = 0

			frm1.vspdData2.Row = Row		
			frm1.vspdData2.Col = C_WcCd
			frm1.KeyWcCd.value  = Trim(frm1.vspdData2.Text) 
			
			lgOldRow2 = Row
			
			If DbQuery3 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
			
		End If
	 	'------ Developer Coding part (End)
	
 	End If

End Sub

'==========================================================================================
'   Event Name : vspdData3_Click
'   Event Desc :
'==========================================================================================
Sub vspdData3_Click(ByVal Col , ByVal Row)

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP3C"
	
	Set gActiveSpdSheet = frm1.vspdData3

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    
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
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
    
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

End Sub

'==========================================================================================
'   Event Name : vspdData3_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData3_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
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
    Call GetSpreadColumnPos("C")
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
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
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
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
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
	End If															'��: Query db data
       
    FncQuery = True																'��: Processing is OK
   
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
' Function Desc : Spread 1 ��ȸ �� Scroll
'========================================================================================

Function DbQuery() 
    
    DbQuery = False                                                         			'��: Processing is NG
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.hProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&cboOrderStatus=" & Trim(frm1.hOrderStatus.value)	'��: ��ȸ ���� ����Ÿ 
	Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&cboOrderStatus=" & Trim(frm1.cboOrderStatus.value)	'��: ��ȸ ���� ����Ÿ 
	End If    

    Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          	'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()													'��: ��ȸ ������ ������� 
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
	
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field
    
    Call SetToolBar("11000000000011")
    
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
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.KeyItemCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)					'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.hProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&cboOrderStatus=" & Trim(frm1.hOrderStatus.value)	'��: ��ȸ ���� ����Ÿ 
		
    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery2 = True                                                          	

End Function

'========================================================================================
' Function Name : DbQuery2Ok
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
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
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.KeyItemCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtWcCd=" & Trim(frm1.KeyWcCd.value)				'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.hProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&cboOrderStatus=" & Trim(frm1.hOrderStatus.value)	'��: ��ȸ ���� ����Ÿ 
	
    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery3 = True                                                          	

End Function

'========================================================================================
' Function Name : DbQuery3Ok
' Function Desc : 
'========================================================================================

Function DbQuery3Ok(LngMaxRow) 
	Call InitData(LngMaxRow)
End Function


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
    If gActiveSpdSheet.Id = "C" Then Call InitSpreadComboBox
    
    ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.ReOrderingSpreadData()

	If gActiveSpdSheet.Id = "C" Then Call InitData(1)
    
End Sub 


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������Ȳ��ȸ</font></td>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="����ø�"></TD>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWCCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()"> <INPUT TYPE=TEXT ID="txtWCNm" NAME="arrCond" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�������� ��ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="�������� ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>							
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���û���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderStatus" ALT="���û���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
									<script language =javascript src='./js/p4422ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4422ma1_B_vspdData2.js'></script>
								</TD>
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4422ma1_C_vspdData3.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderStatus" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="KeyWcCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
