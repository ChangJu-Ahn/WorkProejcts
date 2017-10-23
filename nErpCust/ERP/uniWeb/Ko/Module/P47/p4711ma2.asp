<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name		: Production
'*  2. Function Name	: 
'*  3. Program ID		: p4711ma2.asp
'*  4. Program Name		: �ڿ��Һ�����ȸ 
'*  5. Program Desc		:
'*  6. Comproxy List	: +B19029LookupNumericFormat
'*  7. Modified date(First)	: 2001/11/29
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		:
'* 11. Comment		:
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'<Script LANGUAGE="vbscript"	  SRC="../../inc/incUni2KTV.vbs"></Script>
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<!--'==========================================  1.1.2 ���� Include   ======================================
'============================================================================================================-->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Const BIZ_PGM_CNCL_ID	= "p4711mb3.asp"	' Cancel Batch
Const BIZ_PGM_QRY1_ID	= "p4711mb4.asp"	' �ڿ��� Batch (�ڿ�)
Const BIZ_PGM_QRY2_ID	= "p4711mb5.asp"	' �ڿ��� Batch (�ڿ��� �ڿ��Һ�)
Const BIZ_PGM_QRY3_ID	= "p4711mb6.asp"	' �ڿ��� Batch (����)
Const BIZ_PGM_QRY4_ID	= "p4711mb7.asp"	' ������ Batch (������ �ڿ��Һ�)

<!-- #Include file="../../inc/lgvariables.inc" -->	

Const TAB1 = 1
Const TAB2 = 2

' Grid 1(vspdData1) - Resource
Dim C_ResourceCd2			'= 1
Dim C_ResourceNm2			'= 2
Dim C_ResourceTypeNm2		'= 3
Dim C_ResourceGroupCd2		'= 4
Dim C_ResourceGroupNm2		'= 5
Dim C_ValidFromDt2			'= 6
Dim C_ValidToDt2			'= 7

' Grid 2(vspdData2) - Resource Consumption
Dim C_ProdtOrderNo3			'= 1
Dim C_OprNo3				'= 2
Dim C_ConsumedDt3			'= 3
Dim C_ConsumedTime3			'= 4
Dim C_ItemCd3				'= 5
Dim C_ItemNm3				'= 6
Dim C_Spec3					'= 7
Dim C_ProdQtyInOrderUnit3	'= 8
Dim C_ProdtOrderUnit3		'= 9
Dim C_ReportType3			'= 10
Dim C_JobCd3				'= 11
Dim C_JobNm3				'= 12
Dim C_WcCd3					'= 13
Dim C_WcNm3					'= 14
Dim C_RoutNo3				'= 15

' Grid 3(vspdData3) - Porduction Order
Dim C_ProdtOrderNo4			'= 1
Dim C_OprNo4				'= 2
Dim C_ProdQtyInOrderUnit4	'= 3
Dim C_ProdtOrderUnit4		'= 4
Dim C_ReportType4			'= 5
Dim C_ItemCd4				'= 6
Dim C_ItemNm4				'= 7
Dim C_Spec4					'= 8
Dim C_JobCd4				'= 9
Dim C_JobNm4				'= 10
Dim C_WcCd4					'= 11
Dim C_WcNm4					'= 12
Dim C_RoutNo4				'= 13

' Grid 4(vspdData4) - Resource Consumption
Dim C_ResourceCd5			'= 1
Dim C_ResourceNm5			'= 2
Dim C_ResourceTypeNm5		'= 3
Dim C_ConsumedDt5			'= 4
Dim C_ConsumedTime5			'= 5
Dim C_ResourceGroupCd5		'= 6
Dim C_ResourceGroupNm5		'= 7
Dim C_ValidFromDt5			'= 8
Dim C_ValidToDt5			'= 9

Dim strDate
Dim BaseDate
Dim strYear
Dim strMonth
Dim strDay

BaseDate = "<%=GetsvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)

Dim IsOpenPop						 'Popup
Dim gSelframeFlg
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgStrPrevKey5
Dim lgStrPrevKey6
Dim lgStrPrevKey7
Dim lgStrPrevKey8
Dim lgSortKey2
Dim lgSortKey3
Dim lgSortKey4

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
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
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    gSelframeFlg = 0
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ===================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
    frm1.txtConsumedDtFrom.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
    frm1.txtConsumedDtTo.text   = strDate
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet(ByVal pvSpdNo)
	
	Call initSpreadPosVariables(pvSpdNo)  
			
	If pvSpdNo = "*" or pvSpdNo = "A" then
		'------------------------------------------
		' Grid 1 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData1
    		ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.ReDraw = false
			.MaxCols = C_ValidToDt2 + 1     
			.MaxRows = 0
			Call GetSpreadColumnPos("A")    
			ggoSpread.SSSetEdit		C_ResourceCd2, "�ڿ��ڵ�", 10
			ggoSpread.SSSetEdit		C_ResourceNm2, "�ڿ���", 20
			ggoSpread.SSSetEdit		C_ResourceTypeNm2, "�ڿ�����", 10
			ggoSpread.SSSetEdit		C_ResourceGroupCd2, "�ڿ��׷�", 10
			ggoSpread.SSSetEdit		C_ResourceGroupNm2, "�ڿ��׷��", 20
			ggoSpread.SSSetDate		C_ValidFromDt2, "������",	11,	2,	parent.gDateFormat
			ggoSpread.SSSetDate		C_ValidToDt2, "������",	11,	2,	parent.gDateFormat
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SSSetSplit2(1)	
			.ReDraw = true
		End With
	End if

	If pvSpdNo = "*" or pvSpdNo = "B" then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.ReDraw = false
			.MaxCols = C_RoutNo3 + 1
			.MaxRows = 0
			Call GetSpreadColumnPos("B")
			ggoSpread.SSSetEdit		C_ProdtOrderNo3,		"������ȣ", 18
			ggoSpread.SSSetEdit		C_OprNo3,				"����", 8
			ggoSpread.SSSetDate		C_ConsumedDt3,			"�ڿ��Һ���", 13, 2, parent.gDateFormat
			ggoSpread.SSSetTime 	C_ConsumedTime3,		"�ڿ��Һ�ð�",	13,2 ,1 ,1
			ggoSpread.SSSetEdit		C_ItemCd3,				"ǰ��", 18
			ggoSpread.SSSetEdit		C_ItemNm3,				"ǰ���", 25
			ggoSpread.SSSetEdit		C_Spec3,				"�԰�", 25
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit3,	"��������",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"		
		    ggoSpread.SSSetEdit		C_ProdtOrderUnit3,		"����", 8
			ggoSpread.SSSetEdit		C_ReportType3,			"��/��", 8
			ggoSpread.SSSetCombo	C_JobCd3,				"�۾�", 10
			ggoSpread.SSSetCombo	C_JobNm3,				"�۾���", 20
			ggoSpread.SSSetEdit		C_WcCd3,				"�۾���", 10
			ggoSpread.SSSetEdit		C_WcNm3,				"�۾����", 20
			ggoSpread.SSSetEdit		C_RoutNo3,				"�����", 10
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden(C_JobCd3, C_JobNm3, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SSSetSplit2(2)	
			.ReDraw = true
		End With
	End if

	If pvSpdNo = "*" or pvSpdNo = "C" then
		'------------------------------------------
		' Grid 3 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData3
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.ReDraw = false
			.MaxCols = C_RoutNo4 + 1
			.MaxRows = 0
			Call GetSpreadColumnPos("C")
			ggoSpread.SSSetEdit		C_ProdtOrderNo4,		"������ȣ", 18
			ggoSpread.SSSetEdit		C_OprNo4,				"����", 8
			ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit4,	"��������",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"		
		    ggoSpread.SSSetEdit		C_ProdtOrderUnit4,		"����", 8
			ggoSpread.SSSetEdit		C_ReportType4,			"��/��", 6
			ggoSpread.SSSetEdit		C_ItemCd4,				"ǰ��", 18
			ggoSpread.SSSetEdit		C_ItemNm4,				"ǰ���", 25
			ggoSpread.SSSetEdit		C_Spec4,				"�԰�", 25
			ggoSpread.SSSetCombo	C_JobCd4,				"�۾�", 10
			ggoSpread.SSSetCombo	C_JobNm4,				"�۾���", 20
			ggoSpread.SSSetEdit		C_WcCd4,				"�۾���", 10
			ggoSpread.SSSetEdit		C_WcNm4,				"�۾����", 20
			ggoSpread.SSSetEdit		C_RoutNo4,				"�����", 10
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden(C_ReportType4, C_ReportType4, True)
			Call ggoSpread.SSSetColHidden(C_JobCd4, C_JobNm4, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SSSetSplit2(1)	
			.ReDraw = true
		End With
	End if

	If pvSpdNo = "*" or pvSpdNo = "D" then
		'------------------------------------------
		' Grid 4 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData4
			ggoSpread.Source = frm1.vspdData4
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.ReDraw = false
			.MaxCols = C_ValidToDt5 + 1
			.MaxRows = 0
			Call GetSpreadColumnPos("D")
			ggoSpread.SSSetEdit		C_ResourceCd5,			"�ڿ��ڵ�", 10
			ggoSpread.SSSetEdit		C_ResourceNm5,			"�ڿ���", 20
			ggoSpread.SSSetEdit		C_ResourceTypeNm5,		"�ڿ�����", 10
			ggoSpread.SSSetDate		C_ConsumedDt5,			"�ڿ��Һ���", 13, 2, parent.gDateFormat
			ggoSpread.SSSetTime		C_ConsumedTime5,		"�ڿ��Һ�ð�",	13,2 ,1 ,1
			ggoSpread.SSSetEdit		C_ResourceGroupCd5,		"�ڿ��׷�", 10
			ggoSpread.SSSetEdit		C_ResourceGroupNm5,		"�ڿ��׷��", 20
			ggoSpread.SSSetDate		C_ValidFromDt5,			"������",	11,	2,	parent.gDateFormat
			ggoSpread.SSSetDate		C_ValidToDt5,			"������",	11,	2,	parent.gDateFormat
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SSSetSplit2(2)	
			.ReDraw = true
		End With
	End if
	
    Call SetSpreadLock()
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()

	ggoSpread.Source = frm1.vspdData3
	ggoSpread.SpreadLockWithOddEvenRowColor()

	ggoSpread.Source = frm1.vspdData4
	ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow, Byval Flag)
    

End Sub

'========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Call SetCombo(frm1.cboStatus, "R", "�����")
	Call SetCombo(frm1.cboStatus, "C", "��ҵ�")		'��: InitCombo ���� �ؾ� �Ǵµ� �ӽ÷� ���� ���� 
End Sub


'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	if pvSpdNo = "*" or pvSpdNo = "A" then
		' Grid 1(vspdData1) - Resource
		C_ResourceCd2			= 1
		C_ResourceNm2			= 2
		C_ResourceTypeNm2		= 3
		C_ResourceGroupCd2		= 4
		C_ResourceGroupNm2		= 5
		C_ValidFromDt2			= 6
		C_ValidToDt2			= 7
	End if

	if pvSpdNo = "*" or pvSpdNo = "B" then
		' Grid 2(vspdData2) - Resource Consumption
		C_ProdtOrderNo3			= 1
		C_OprNo3				= 2
		C_ConsumedDt3			= 3
		C_ConsumedTime3			= 4
		C_ItemCd3				= 5
		C_ItemNm3				= 6
		C_Spec3					= 7
		C_ProdQtyInOrderUnit3	= 8
		C_ProdtOrderUnit3		= 9
		C_ReportType3			= 10
		C_JobCd3				= 11
		C_JobNm3				= 12
		C_WcCd3					= 13
		C_WcNm3					= 14
		C_RoutNo3				= 15
	End if

	if pvSpdNo = "*" or pvSpdNo = "C" then
		' Grid 3(vspdData3) - Porduction Order
		C_ProdtOrderNo4			= 1
		C_OprNo4				= 2
		C_ProdQtyInOrderUnit4	= 3
		C_ProdtOrderUnit4		= 4
		C_ReportType4			= 5
		C_ItemCd4				= 6
		C_ItemNm4				= 7
		C_Spec4					= 8
		C_JobCd4				= 9
		C_JobNm4				= 10
		C_WcCd4					= 11
		C_WcNm4					= 12
		C_RoutNo4				= 13
	End if

	if pvSpdNo = "*" or pvSpdNo = "D" then
		' Grid 4(vspdData4) - Resource Consumption
		C_ResourceCd5			= 1
		C_ResourceNm5			= 2
		C_ResourceTypeNm5		= 3
		C_ConsumedDt5			= 4
		C_ConsumedTime5			= 5
		C_ResourceGroupCd5		= 6
		C_ResourceGroupNm5		= 7
		C_ValidFromDt5			= 8
		C_ValidToDt5			= 9
	End if

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
		C_ResourceCd2			= iCurColumnPos(1)
		C_ResourceNm2			= iCurColumnPos(2)
		C_ResourceTypeNm2		= iCurColumnPos(3)
		C_ResourceGroupCd2		= iCurColumnPos(4)
		C_ResourceGroupNm2		= iCurColumnPos(5)
		C_ValidFromDt2			= iCurColumnPos(6)
		C_ValidToDt2			= iCurColumnPos(7)
		
	Case "B"
		ggoSpread.Source = frm1.vspdData2 
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_ProdtOrderNo3			= iCurColumnPos(1)
		C_OprNo3				= iCurColumnPos(2)
		C_ConsumedDt3			= iCurColumnPos(3)
		C_ConsumedTime3			= iCurColumnPos(4)
		C_ItemCd3				= iCurColumnPos(5)
		C_ItemNm3				= iCurColumnPos(6)
		C_Spec3					= iCurColumnPos(7)
		C_ProdQtyInOrderUnit3	= iCurColumnPos(8)
		C_ProdtOrderUnit3		= iCurColumnPos(9)
		C_ReportType3			= iCurColumnPos(10)
		C_JobCd3				= iCurColumnPos(11)
		C_JobNm3				= iCurColumnPos(12)
		C_WcCd3					= iCurColumnPos(13)
		C_WcNm3					= iCurColumnPos(14)
		C_RoutNo3				= iCurColumnPos(15)
	
	Case "C"
		ggoSpread.Source = frm1.vspdData3 
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
  		C_ProdtOrderNo4			= iCurColumnPos(1)
		C_OprNo4				= iCurColumnPos(2)
		C_ProdQtyInOrderUnit4	= iCurColumnPos(3)
		C_ProdtOrderUnit4		= iCurColumnPos(4)
		C_ReportType4			= iCurColumnPos(5)
		C_ItemCd4				= iCurColumnPos(6)
		C_ItemNm4				= iCurColumnPos(7)
		C_Spec4					= iCurColumnPos(8)
		C_JobCd4				= iCurColumnPos(9)
		C_JobNm4				= iCurColumnPos(10)
		C_WcCd4					= iCurColumnPos(11)
		C_WcNm4					= iCurColumnPos(12)
		C_RoutNo4				= iCurColumnPos(13)
		
	Case "D"
		ggoSpread.Source = frm1.vspdData4
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
  		C_ResourceCd5			= iCurColumnPos(1)
		C_ResourceNm5			= iCurColumnPos(2)
		C_ResourceTypeNm5		= iCurColumnPos(3)
		C_ConsumedDt5			= iCurColumnPos(4)
		C_ConsumedTime5			= iCurColumnPos(5)
		C_ResourceGroupCd5		= iCurColumnPos(6)
		C_ResourceGroupNm5		= iCurColumnPos(7)
		C_ValidFromDt5			= iCurColumnPos(8)
		C_ValidToDt5			= iCurColumnPos(9)
		
  	End Select
  
End Sub


'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�.
'*********************************************************************************************************
'==========================================  2.3.1 Tab Click ó��  =================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'===================================================================================================================
'----------------  ClickTab1(): Header Tabó�� �κ� (Header Tab�� �ִ� ��츸 ���)  ----------------------------
Function ClickTab1()

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
		Call SetToolbar("11000000000111")
		Exit Function
    End If
	
	If gSelframeFlg = TAB1 Then Exit Function

    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field    
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call changeTabs(TAB1)	
	gSelframeFlg = TAB1
	lgIntFlgMode = parent.OPMD_CMODE
	
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function           
    End If 

End Function

Function ClickTab2()

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
		Call SetToolbar("11000000000111")
		Exit Function
    End If

	If gSelframeFlg = TAB2 Then Exit Function

    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field    
	ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData4
    ggoSpread.ClearSpreadData
	Call changeTabs(TAB2)	
	gSelframeFlg = TAB2
	lgIntFlgMode = parent.OPMD_CMODE
	
    If DbQuery = False Then   
		Call RestoreToolBar()
		Exit Function           
    End If 
    	
End Function

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenBatchRunNo()  -------------------------------------------------
'	Name : OpenBatchRunNo()
'	Description : Batch Run No. PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBatchRunNo()

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtBatchRunNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))
	arrParam(1) = Trim(frm1.txtPlantNm.value)
	arrParam(2) = UCase(Trim(frm1.txtBatchRunNo.value))

	iCalledAspName = AskPRAspName("p4711pa1")
 	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4711pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBatchRunNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBatchRunNo.focus
	
End Function

'------------------------------------------  OpenResourceGroup()  -------------------------------------------------
'	Name : OpenResourceGroup()
'	Description : ResourceGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResourceGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "�ڿ��׷��˾�"	
	arrParam(1) = "P_RESOURCE_GROUP"				
	arrParam(2) = Trim(frm1.txtResourceGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
				  			
	arrParam(5) = "�ڿ��׷�"			
	    
    arrField(0) = "RESOURCE_GROUP_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ��׷�"		
    arrHeader(1) = "�ڿ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetResourceGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtResourceGroupCd.focus
	
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
		Call DisplayMsgBox("971012","X", "����","X")
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

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)


	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "�ڿ��˾�"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtResourceCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")		
	arrParam(5) = "�ڿ�"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ�"		
    arrHeader(1) = "�ڿ���"
	
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode			' Item Code
	arrParam(2) = "12!MO"			' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field��(0)
	arrField(1) = 2 '"ITEM_NM"			' Field��(1)
	
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

'------------------------------------------  OpenErrorRef()  -------------------------------------------------
'	Name : OpenErrorRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenErrorRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	If frm1.txtBatchRunNo.value= "" Then
		Call DisplayMsgBox("971012","X", "�̷¹�ȣ","X")
		frm1.txtBatchRunNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	arrParam(0) = Trim(UCase(frm1.txtPlantCd.value))	'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtPlantNm.value)			'��: ��ȸ ���� ����Ÿ 
	arrParam(2) = Trim(frm1.txtBatchRunNo.value)		'��: ��ȸ ���� ����Ÿ 
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4711ra2")
 	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4711ra2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetBatchRunNo()  --------------------------------------------------
'	Name : SetBatchRunNo()
'	Description : ResourceGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBatchRunNo(byval arrRet)
	frm1.txtBatchRunNo.Value = arrRet(0)
	frm1.cboStatus.Value	 = arrRet(1)
End Function

'------------------------------------------  SetResourceGroup()  --------------------------------------------------
'	Name : SetResourceGroup()
'	Description : ResourceGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResourceGroup(byval arrRet)
	frm1.txtResourceGroupCd.Value    = arrRet(0)		
	frm1.txtResourceGroupNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	frm1.txtResourceCd.Value    = arrRet(0)		
	frm1.txtResourceNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdtOrderNo()  -------------------------------------------
'	Name : SetProdtOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdtOrderNo(byval arrRet)
    frm1.txtProdtOrderNo.Value    = arrRet(0)
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  -------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
    With frm1
	.txtItemCd.value = arrRet(0)
	.txtItemNm.value = arrRet(1)
    End With
End Function

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	
	Call AppendNumberPlace("6","11","4")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitComboBox
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtBatchRunNo.focus 
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If
       
	Call InitSpreadSheet("*")
       
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11000000000011")									'��: ��ư ���� ���� 
    
    gTabMaxCnt = 2
    gIsTab = "Y"
   
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
'*****************************************************************************************************

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

 	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 

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
 		
 		If gSelframeFlg = TAB1 Then
			With frm1.vspdData1
				.Row = .ActiveRow
				.Col = C_ResourceCd2
				frm1.hResourceCd.value = .Text
				frm1.vspdData2.MaxRows = 0
			End With
		
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		
		End If
 		
	Else
		If gSelframeFlg = TAB1 Then
			With frm1.vspdData1
				If .MaxRows <= 0 Then Exit Sub
				If Row < 1 Then Exit Sub
				.Row = Row
				.Col = C_ResourceCd2
				frm1.hResourceCd.value = .Text
				frm1.vspdData2.MaxRows = 0
			End With
			
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	
 	End If
 	
	
End Sub


'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
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

 	End If

End Sub


'==========================================================================================
'   Event Name : vspdData3_Click
'   Event Desc :
'==========================================================================================

Sub vspdData3_Click(ByVal Col, ByVal Row)
 	
 	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
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
 		
 		If gSelframeFlg = TAB2 Then
 			With frm1.vspdData3
				.Row = .ActiveRow
				.Col = C_ProdtOrderNo4
				frm1.hProdtOrderNo.value = .Text
				.Col = C_OprNo4
				frm1.hOprNo.value = .Text
				frm1.vspdData4.MaxRows = 0
			End With
			
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
			
		End If
 		
	Else
		If gSelframeFlg = TAB2 Then
			With frm1.vspdData3
				If .MaxRows <= 0 Then Exit Sub
				If Row < 1 Then Exit Sub
				.Row = Row
				.Col = C_ProdtOrderNo4
				frm1.hProdtOrderNo.value = .Text
				.Col = C_OprNo4
				frm1.hOprNo.value = .Text
				frm1.vspdData4.MaxRows = 0
			End With
			
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If

 	End If

End Sub


'========================================================================================
' Function Name : vspdData4_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData4_Click(ByVal Col, ByVal Row)
 	
 	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
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
 	End If
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

Sub vspdData4_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP4C" Then
       gMouseClickStatus = "SP4CR"
    End If
End Sub		

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
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

Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
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

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub 

Sub vspdData3_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("C")
End Sub 

Sub vspdData4_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("D")
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
	Dim pvSpdNo
	
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    pvSpdNo = gActiveSpdSheet.id
    Call InitSpreadSheet(pvSpdNo)  
    
    Select Case pvSpdNo
    case "A"
    	ggoSpread.Source = frm1.vspdData1
	case "B"
		ggoSpread.Source = frm1.vspdData2
    case "C"
    	ggoSpread.Source = frm1.vspdData3
    case "D"
    	ggoSpread.Source = frm1.vspdData4
    End Select 
	
	Call ggoSpread.ReOrderingSpreadData()
	
'	if pvSpdNo = "A" or pvSpdNo = "C" then
'		'-------------------------------------
'		'  DbDtlQuery
'		'-------------------------------------	
'		ggoSpread.Source = frm1.vspdData1
'		frm1.vspddata1.Row = 1
'		frm1.vspddata1.Col = C_ResourceCd2
'		frm1.hResourceCd.value = frm1.vspddata1.Text
'		frm1.vspdData2.MaxRows = 0
'		ggoSpread.Source = frm1.vspdData3
'		frm1.vspddata3.Row = 1
'		frm1.vspddata3.Col = C_ProdtOrderNo4
'		frm1.hProdtOrderNo.value = frm1.vspddata3.Text
'		frm1.vspddata3.Col = C_OprNo4
'		frm1.hOprNo.value = frm1.vspddata3.Text
'		frm1.vspdData4.MaxRows = 0
'		Call DbDtlQuery
'		
'		Set gActiveElement = document.ActiveElement
'	end if
	
End Sub 
 
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
     Exit Sub
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
	If gSelframeFlg = TAB1 Then   
		if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
			If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery = False Then	
					Call RestoreToolBar()
					Exit Sub
				End If	
			End If
		End If
    End If
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
     Exit Sub
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
	If gSelframeFlg = TAB1 Then   
		if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
			If lgStrPrevKey2 <> "" AND lgStrPrevKey3 <> "" AND lgStrPrevKey4 <> "" Then	 '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbDtlQuery = False Then	
					Call RestoreToolBar()
					Exit Sub
				End If	
			End If
		End If
    End If
    
End Sub

Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    Exit Sub
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
	If gSelframeFlg = TAB2 Then   
		if frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then
			If lgStrPrevKey5 <> "" AND lgStrPrevKey6 <> "" AND lgStrPrevKey7 <> "" Then						'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery = False Then	
					Call RestoreToolBar()
					Exit Sub
				End If	
			End If
		End If
    End If
    
End Sub

Sub vspdData4_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    Exit Sub
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
	If gSelframeFlg = TAB2 Then   
		if frm1.vspdData4.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData4,NewTop) Then
			If lgStrPrevKey8 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbDtlQuery = False Then	
					Call RestoreToolBar()
					Exit Sub
				End If	
			End If
		End If
    End If
    
End Sub

'=======================================================================================================
'   Event Name : txtConsumedDtFrom_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtConsumedDtFrom_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtConsumedDtTo_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtConsumedDtTo_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub



'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
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
   	If ValidDateCheck(frm1.txtConsumedDtFrom, frm1.txtConsumedDtTo) = False Then Exit Function
   	
   	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If

    Call InitVariables
   
    '-----------------------
    'Query function call area
    '-----------------------
    If gSelframeFlg = TAB1 or gSelframeFlg <> TAB2 Then
        Call ClickTab1
	Else
        Call ClickTab2
	End If

    FncQuery = True																'��: Processing is OK
    
End Function
'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_MULTI, True)
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

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			 '��: Data is changed.  Do you want to exit? 
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

    Err.Clear                                                              
    
    DbQuery = False                                                        
	
	Call LayerShowHide(1)													

	If gSelframeFlg = TAB1 Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
	Else
		strVal = BIZ_PGM_QRY3_ID & "?txtMode=" & parent.UID_M0001
	End If
	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value )
	strVal = strVal & "&txtBatchRunNo=" & Trim(frm1.txtBatchRunNo.value)
	strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.txtProdtOrderNo.value)
	strVal = strVal & "&txtResourceCd=" & Trim(frm1.txtResourceCd.value)
	strVal = strVal & "&txtConsumedDtFrom=" & Trim(frm1.txtConsumedDtFrom.text)
	strVal = strVal & "&txtConsumedDtTo=" & Trim(frm1.txtConsumedDtTo.text)
	strVal = strVal & "&txtResourceGroupCd=" & Trim(frm1.txtResourceGroupCd.value)
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
	strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)

	Call RunMyBizASP(MyBizASP, strVal)
	
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		If gSelframeFlg = TAB1 Then
			Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Else
			Call SetActiveCell(frm1.vspdData3,1,1,"M","X","X")
		End If
		Set gActiveElement = document.activeElement	
	End If
	
    lgIntFlgMode = parent.OPMD_UMODE										
      
    Call ggoOper.LockField(Document, "Q")
	Call SetToolbar("11000000000111")
	ggoSpread.Source = frm1.vspdData1
	frm1.vspddata1.Row = 1
	frm1.vspddata1.Col = C_ResourceCd2
	frm1.hResourceCd.value = frm1.vspddata1.Text
	ggoSpread.Source = frm1.vspdData3
	frm1.vspddata3.Row = 1
	frm1.vspddata3.Col = C_ProdtOrderNo4
	frm1.hProdtOrderNo.value = frm1.vspddata3.Text
	frm1.vspddata3.Col = C_OprNo4
	frm1.hOprNo.value = frm1.vspddata3.Text
	
	Call DbDtlQuery

	Set gActiveElement = document.ActiveElement
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 
	Dim strVal

    Err.Clear                                                              
    
    DbDtlQuery = False                                                        
	
	Call LayerShowHide(1)													

	If gSelframeFlg = TAB1 Then
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtResourceCd=" & Trim(frm1.hResourceCd.value)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.txtProdtOrderNo.value)
	Else
		strVal = BIZ_PGM_QRY4_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtResourceCd=" & Trim(frm1.txtResourceCd.value)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(frm1.hProdtOrderNo.value)
		strVal = strVal & "&txtOprNo=" & Trim(frm1.hOprNo.value)
	End If

	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value )
	strVal = strVal & "&txtBatchRunNo=" & Trim(frm1.txtBatchRunNo.value)
	strVal = strVal & "&txtConsumedDtFrom=" & Trim(frm1.txtConsumedDtFrom.text)
	strVal = strVal & "&txtConsumedDtTo=" & Trim(frm1.txtConsumedDtTo.text)
	strVal = strVal & "&txtResourceGroupCd=" & Trim(frm1.txtResourceGroupCd.value)
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
	strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
        
	Call RunMyBizASP(MyBizASP, strVal)										
	DbDtlQuery = True                                                          

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
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
						<TABLE CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ��Һ�����ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenErrorRef()">��������</A></TD>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU"  ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()" >&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>�̷¹�ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBatchRunNo" SIZE=18 MAXLENGTH=18 tag="12XXXU"  ALT="�̷¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBatchRunNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBatchRunNo()" >&nbsp;<SELECT NAME="cboStatus" ALT="Status" STYLE="Width: 90px;" tag="14"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�������� ��ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdtOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="�������� ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdtOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdtOrderNo()"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�ڿ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceGroupCd" SIZE=10 MAXLENGTH=10 tag="11xxxU"  ALT="�ڿ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenResourceGroup()" >&nbsp;<INPUT TYPE=TEXT NAME="txtResourceGroupNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ڿ��Һ���</TD>
									<TD CLASS=TD6>
										<script language =javascript src='./js/p4711ma2_I180849236_txtConsumedDtFrom.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p4711ma2_I562043440_txtConsumedDtTo.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWCCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()"> <INPUT TYPE=TEXT ID="txtWCNm" NAME="arrCond" tag="14"></TD>
								</TR>								
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<!-- DATA AREA -->
					<TD WIDTH="100%" HEIGHT="100%">
						<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
							<TR HEIGHT=23>
								<TD WIDTH="100%">
									<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH="100%" border=0>
										<TR>
											<TD WIDTH=10>&nbsp;</TD>
											<TD CLASS="CLSMTABP">
												<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
													<TR>
														<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ڿ���</font></td>
														<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
												    </TR>
												</TABLE>
											</TD>
											<TD CLASS="CLSMTABP">
												<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
													<TR>
														<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
														<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>������</font></td>
														<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
												    </TR>
												</TABLE>
											</TD>
											<TD WIDTH=300>&nbsp;</TD>
											<TD WIDTH=*>&nbsp;</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD WIDTH="100%" CLASS="TB3">
									<!-- ù��° �� ���� -->
									<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
										<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD WIDTH=100% HEIGHT=* valign=top>
													<TABLE <%=LR_SPACE_TYPE_20%>>
														<TR HEIGHT="60%">
															<TD WIDTH="40%" colspan=4>
																<script language =javascript src='./js/p4711ma2_A_vspdData1.js'></script>
															</TD>
															<TD WIDTH="60%" colspan=4>
																<script language =javascript src='./js/p4711ma2_B_vspdData2.js'></script>
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>										
										</TABLE>
									</DIV> 
									<!-- �ι�° �� ���� -->
									<DIV ID="TabDiv"  SCROLL="no" style="display:none">
										<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD WIDTH=100% HEIGHT=* valign=top>
													<TABLE <%=LR_SPACE_TYPE_20%>>
														<TR HEIGHT="60%">
															<TD WIDTH="40%" colspan=4>
																<script language =javascript src='./js/p4711ma2_C_vspdData3.js'></script>
															</TD>
															<TD WIDTH="70%" colspan=4>
																<script language =javascript src='./js/p4711ma2_D_vspdData4.js'></script>
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</DIV>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=20 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX = "-1"></IFRAME></TD>
	</TR>
</TABLE><TEXTAREA class=hidden name=txtSpread tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtlgMode" tag="24">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24">
<INPUT TYPE=hidden NAME="hResourceCd" tag="24">
<INPUT TYPE=hidden NAME="hProdtOrderNo" tag="24">
<INPUT TYPE=hidden NAME="hOprNo" tag="24">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

