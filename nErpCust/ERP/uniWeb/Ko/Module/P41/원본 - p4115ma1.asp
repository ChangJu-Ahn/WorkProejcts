
<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4115ma1.asp
'*  4. Program Name         : ��������Ȳ ��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2002/06/27
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Park, BumSoo
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
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                      

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY1_ID	= "p4115mb1.asp"
Const BIZ_PGM_QRY2_ID	= "p4115mb2.asp"
Const BIZ_PGM_QRY3_ID	= "p4115mb3.asp"

'-------------------------------
' Column Constants : Spread 1 
'-------------------------------
Dim C_OrdNo
Dim C_OprNo
Dim C_WcCd
Dim C_WcNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_JobCd
Dim C_JobNm
Dim C_ScheldStartDt
Dim C_ScheldEndDt
Dim C_OrdQty
Dim C_OrdUnit
Dim C_ProdQty
Dim C_GoodQty
Dim C_BadQty
Dim C_InspQty
Dim C_InspGoodQty
Dim C_InspBadQty
Dim C_RcptQty
Dim C_InsideFlg
Dim C_CurCd
Dim C_SubconPrc
Dim C_OrderStatus
Dim C_OrderStatusDesc
Dim C_OrderType
Dim C_OrderTypeDesc
Dim C_ItemGroupCd
Dim C_ItemGroupNm

'-------------------------------
' Column Constants : Spread 2 
'-------------------------------
Dim C_ChildItemCd
Dim C_ChildItemNm
Dim C_ChildSpec
Dim C_ReqDt
Dim C_ReqQty
Dim C_ResvUnit
Dim C_IssuedQty
Dim C_ConsumedQty
Dim C_SLCd
Dim C_SLNm

'-------------------------------
' Column Constants : Spread 3 
'-------------------------------
Dim C_Resource
Dim C_ResourceNm
Dim C_StartDt
Dim C_EndDt
Dim C_StartFlg
Dim C_EndFlg
Dim C_RscGrpCd

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
Dim lgStrPrevKey4
Dim lgStrPrevKey5

Dim lgSortKey1
Dim lgSortKey2
Dim lgSortKey3
		
Dim lgLngCurRows

Dim lgOldRow

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

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgStrPrevKey5 = ""
    
    lgOldRow = 0
    
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
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
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtProdFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -10, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtProdToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 20, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "P", "NOCOOKIE", "QA") %>
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
		
			.vspdData1.MaxCols = C_ItemGroupNm + 1
			.vspdData1.MaxRows = 0
		
			Call GetSpreadColumnPos("A")
				
			ggoSpread.SSSetEdit		C_OrdNo,		"������ȣ", 18
			ggoSpread.SSSetEdit		C_OprNo,		"����", 6
			ggoSpread.SSSetEdit		C_WCCd,			"�۾���", 8
			ggoSpread.SSSetEdit		C_WCNm,			"�۾����", 20
			ggoSpread.SSSetEdit		C_ItemCd,		"ǰ��", 18
			ggoSpread.SSSetEdit		C_ItemNm,		"ǰ���", 25
			ggoSpread.SSSetEdit		C_Spec,			"�԰�", 25
			ggoSpread.SSSetEdit		C_JobCd,		"�۾�", 6
			ggoSpread.SSSetEdit		C_JobNm,		"�۾���", 10
			ggoSpread.SSSetDate 	C_ScheldStartDt, "����������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ScheldEndDt,	"�ϷΌ����", 11, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_OrdQty,		"��������",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_OrdUnit,		"����", 8,,,3,2
			ggoSpread.SSSetFloat	C_ProdQty,		"��������",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GoodQty,		"��ǰ����",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BadQty,		"�ҷ�����",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspQty,		"�˻��߼�",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspGoodQty,	"ǰ���հݼ�",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspBadQty,	"ǰ���ҷ���",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RcptQty,		"�԰����",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_InsideFlg,	"�系/��", 10
			ggoSpread.SSSetEdit		C_CurCd,		"��ȭ", 8,,,3,2
			ggoSpread.SSSetFloat	C_SubconPrc,	"���ִܰ�", 15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo	C_OrderStatus,	"���û���", 10
			ggoSpread.SSSetCombo	C_OrderType,	"���ñ���", 10
			ggoSpread.SSSetCombo	C_OrderStatusDesc, "���û���", 10
			ggoSpread.SSSetCombo	C_OrderTypeDesc, "���ñ���", 10
			ggoSpread.SSSetEdit 	C_ItemGroupCd, "ǰ��׷�",	15
			ggoSpread.SSSetEdit		C_ItemGroupNm, "ǰ��׷��", 30
		
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden(C_InspQty, C_InspQty, True)
			Call ggoSpread.SSSetColHidden(C_OrderStatus, C_OrderStatus, True)
			Call ggoSpread.SSSetColHidden(C_OrderType, C_OrderType, True)
			Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
		
			ggoSpread.SSSetSplit2(3)
			
			Call SetSpreadLock("A")
			
			.vspdData1.ReDraw = True
		End If	
		
		If pvSpdNo = "B" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 2 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.vspdData2.ReDraw = false
			.vspdData2.MaxCols = C_SLNm + 1
			.vspdData2.MaxRows = 0
		
			Call GetSpreadColumnPos("B")
			ggoSpread.SSSetEdit		C_ChildItemCd,	"��ǰ", 18
			ggoSpread.SSSetEdit		C_ChildItemNm,	"��ǰ��", 25
			ggoSpread.SSSetEdit		C_ChildSpec,	"�԰�", 25
			ggoSpread.SSSetDate 	C_ReqDt,		"�ʿ���", 11, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_ReqQty,		"�ʿ����",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_ResvUnit,		"����", 10,,,3,2
			ggoSpread.SSSetFloat	C_IssuedQty,	"������",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ConsumedQty,	"�Һ����",15,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_SLCd,			"â��", 10
			ggoSpread.SSSetEdit		C_SLNm,			"â���", 20
		
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( .vspdData2.MaxCols, .vspdData2.MaxCols, True)
		
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("B")
			
			.vspdData2.ReDraw = True
		End If	
			
		If pvSpdNo = "C" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 3 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData3
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.vspdData3.ReDraw = False
			.vspdData3.MaxCols = C_RscGrpCd + 1
			.vspdData3.MaxRows = 0
		
			Call GetSpreadColumnPos("C")
			ggoSpread.SSSetEdit		C_Resource,		"�ڿ�", 10
			ggoSpread.SSSetEdit		C_ResourceNm,	"�ڿ���", 20
			ggoSpread.SSSetDate 	C_StartDt,		"��������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_EndDt,		"��������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_StartFlg,		"���۱���", 10
			ggoSpread.SSSetEdit		C_EndFlg,		"���ᱸ��", 10
			ggoSpread.SSSetEdit		C_RscGrpCd,		"�ڿ��׷�", 10
		
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( .vspdData3.MaxCols, .vspdData3.MaxCols, True)
		
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("C")
			
			.vspdData3.ReDraw = True
		End If
		
    End With
    
    
    
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		If pvSpdNo = "A" Then
			'-------------------------
			' Set Lock Prop :Spread 1 		
			'-------------------------
			ggoSpread.Source = .vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "B" Then
			'-------------------------
			' Set Lock Prop :Spread 2 		
			'-------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
				
		If pvSpdNo = "C" Then
			'-------------------------
			' Set Lock Prop :Spread 3 		
			'-------------------------
			.vspdData3.ReDraw = False
			ggoSpread.Source = .vspdData3
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
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
		C_OrdNo				= 1																	'��: Spread Sheet �� Columns �ε��� 
		C_OprNo				= 2
		C_WcCd				= 3
		C_WcNm				= 4
		C_ItemCd			= 5
		C_ItemNm			= 6
		C_Spec				= 7
		C_JobCd				= 8
		C_JobNm				= 9
		C_ScheldStartDt		= 10
		C_ScheldEndDt		= 11
		C_OrdQty			= 12
		C_OrdUnit			= 13
		C_ProdQty			= 14
		C_GoodQty			= 15
		C_BadQty			= 16
		C_InspQty			= 17
		C_InspGoodQty		= 18
		C_InspBadQty		= 19
		C_RcptQty			= 20
		C_InsideFlg			= 21
		C_CurCd				= 22
		C_SubconPrc			= 23
		C_OrderStatus		= 24
		C_OrderStatusDesc	= 25
		C_OrderType			= 26
		C_OrderTypeDesc		= 27
		C_ItemGroupCd		= 28
		C_ItemGroupNm		= 29
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then 
		'-------------------------------
		' Column Constants : Spread 2 
		'-------------------------------
		C_ChildItemCd		= 1
		C_ChildItemNm		= 2
		C_ChildSpec			= 3
		C_ReqDt				= 4
		C_ReqQty			= 5
		C_ResvUnit			= 6
		C_IssuedQty			= 7
		C_ConsumedQty		= 8
		C_SLCd				= 9
		C_SLNm				= 10
	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then 
		'-------------------------------
		' Column Constants : Spread 3 
		'-------------------------------
		C_Resource			= 1
		C_ResourceNm		= 2
		C_StartDt			= 3
		C_EndDt				= 4
		C_StartFlg			= 5
		C_EndFlg			= 6
		C_RscGrpCd			= 7
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
 			
			'-------------------------------
			' Column Constants : Spread 1 
			'-------------------------------
			C_OrdNo				= iCurColumnPos(1)																	'��: Spread Sheet �� Columns �ε��� 
			C_OprNo				= iCurColumnPos(2)
			C_WcCd				= iCurColumnPos(3)
			C_WcNm				= iCurColumnPos(4)
			C_ItemCd			= iCurColumnPos(5)
			C_ItemNm			= iCurColumnPos(6)
			C_Spec				= iCurColumnPos(7)
			C_JobCd				= iCurColumnPos(8)
			C_JobNm				= iCurColumnPos(9)
			C_ScheldStartDt		= iCurColumnPos(10)
			C_ScheldEndDt		= iCurColumnPos(11)
			C_OrdQty			= iCurColumnPos(12)
			C_OrdUnit			= iCurColumnPos(13)
			C_ProdQty			= iCurColumnPos(14)
			C_GoodQty			= iCurColumnPos(15)
			C_BadQty			= iCurColumnPos(16)
			C_InspQty			= iCurColumnPos(17)
			C_InspGoodQty		= iCurColumnPos(18)
			C_InspBadQty		= iCurColumnPos(19)
			C_RcptQty			= iCurColumnPos(20)
			C_InsideFlg			= iCurColumnPos(21)
			C_CurCd				= iCurColumnPos(22)
			C_SubconPrc			= iCurColumnPos(23)
			C_OrderStatus		= iCurColumnPos(24)
			C_OrderStatusDesc	= iCurColumnPos(25)
			C_OrderType			= iCurColumnPos(26)
			C_OrderTypeDesc		= iCurColumnPos(27)
			C_ItemGroupCd		= iCurColumnPos(28)
			C_ItemGroupNm		= iCurColumnPos(29)
	
		Case "B"
			ggoSpread.Source = frm1.vspdData2
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			'-------------------------------
			' Column Constants : Spread 2 
			'-------------------------------
			C_ChildItemCd		= iCurColumnPos(1)
			C_ChildItemNm		= iCurColumnPos(2)
			C_ChildSpec			= iCurColumnPos(3)
			C_ReqDt				= iCurColumnPos(4)
			C_ReqQty			= iCurColumnPos(5)
			C_ResvUnit			= iCurColumnPos(6)
			C_IssuedQty			= iCurColumnPos(7)
			C_ConsumedQty		= iCurColumnPos(8)
			C_SLCd				= iCurColumnPos(9)
			C_SLNm				= iCurColumnPos(10)
			
		Case "C"
			ggoSpread.Source = frm1.vspdData3
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
 			'-------------------------------
			' Column Constants : Spread 3 
			'-------------------------------
			C_Resource			= iCurColumnPos(1)
			C_ResourceNm		= iCurColumnPos(2)
			C_StartDt			= iCurColumnPos(3)
			C_EndDt				= iCurColumnPos(4)
			C_StartFlg			= iCurColumnPos(5)
			C_EndFlg			= iCurColumnPos(6)
			C_RscGrpCd			= iCurColumnPos(7)
	
 	End Select
	
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderType, lgF0, lgF1, Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderStatus, lgF0, lgF1, Chr(11))

	frm1.cboOrderType.value = ""
	frm1.cboOrderStatus.value = "" 	 
    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderTypeDesc

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    ggoSpread.Source = frm1.vspdData1
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

	With frm1.vspdData1
		For intRow = lngStartRow To .MaxRows

			.Row = intRow
			.col = C_OrderStatus
			intIndex = .value
			.Col = C_OrderStatusDesc
			.value = intindex
			.Row = intRow
			.col = C_OrderType
			intIndex = .value
			.Col = C_OrderTypeDesc
			.value = intindex			
		Next
	End With
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

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

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

	If arrRet(0) = "" Then
		Exit Function
	Else
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
		Call DisplayMsgBox("971012","X" , "����","X")
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
	arrParam(2) = "12!MO"						' Combo Set Data:"12!MO" -- ǰ����� ������ ���ޱ���: From To�� �Է��� �� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
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

'------------------------------------------  OpenWcPopup()  ----------------------------------------------
'	Name : OpenWcPopup()
'	Description : WcPopup
'---------------------------------------------------------------------------------------------------------
Function OpenWcPopup()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = Trim(frm1.txtWcCd.value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
	arrParam(5) = "�۾���"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "INSIDE_FLG"
    
    arrHeader(0) = "�۾���"		
    arrHeader(1) = "�۾����"		
    arrHeader(2) = "�۾���Ÿ��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWc(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWcCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

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
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdToDt.Text
	arrParam(3) = frm1.cboOrderStatus.value
	arrParam(4) = frm1.cboOrderStatus.value
	arrParam(5) = Trim(frm1.txtProdOrderNo.value) 
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = Trim(frm1.cboOrderType.value)
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
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
		Call DisplayMsgBox("971012","X", "����","X")
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

'------------------------------------------  OpenProdRef()  -------------------------------------------------
'	Name : OpenProdRef()
'	Description : Production Reference
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdRef()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = Parent.OPMD_CMODE Then
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
	
	iCalledAspName = AskPRAspName("P4411RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4411RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.vspdData1.MaxRows <= 0 Then Exit Function
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	
	arrParam(1) = Trim(frm1.KeyProdOrdNo.value) 
	arrParam(2) = Trim(frm1.KeyOprNo.value) 
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  -------------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Issue Reference
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = Parent.OPMD_CMODE Then
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

	If frm1.vspdData1.MaxRows <= 0 Then Exit Function
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		
	arrParam(1) = Trim(frm1.KeyProdOrdNo.value) 
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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

	If lgIntFlgMode = Parent.OPMD_CMODE Then
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
	
	If frm1.vspdData1.MaxRows <= 0 Then Exit Function
	
	iCalledAspName = AskPRAspName("P4412RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4412RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		
	arrParam(1) = Trim(frm1.KeyProdOrdNo.value) 
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
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

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetWc()  ----------------------------------------------------
'	Name : SetWc()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetWc(byval arrRet)
	frm1.txtWcCd.Value    = arrRet(0)		
	frm1.txtWcNm.Value    = arrRet(1)		
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
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
	Call InitSpreadSheet("*")															'��: Setup the Spread sheet
	Call InitSpreadComboBox
	Call InitComboBox
	Call SetDefaultVal
	Call InitVariables																'��: Initializes local global variables
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11000000000011")
		
	If frm1.txtPlantCd.value = "" Then
		If Parent.gPlant <> "" Then
			frm1.txtPlantCd.value = Parent.gPlant
			frm1.txtPlantNm.value = Parent.gPlantNm
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
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
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
'   Event Name : txtProdFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtProdFromDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtProdToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtProdToDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
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
 		
 		lgOldRow = frm1.vspdData1.ActiveRow
			
		frm1.vspdData2.MaxRows = 0 
		frm1.vspdData3.MaxRows = 0 
		lgStrPrevKey3 = ""
		lgStrPrevKey4 = ""
		lgStrPrevKey5 = ""
			
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_OrdNo
		frm1.KeyProdOrdNo.value = Trim(frm1.vspdData1.Text)
			
		frm1.vspdData1.Col = C_OprNo
		frm1.KeyOprNo.value = Trim(frm1.vspdData1.Text) 
			
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then
		
			lgOldRow = Row
			
			frm1.vspdData2.MaxRows = 0 
			frm1.vspdData3.MaxRows = 0 
			lgStrPrevKey3 = ""
			lgStrPrevKey4 = ""
			lgStrPrevKey5 = ""
			
			frm1.vspdData1.Row = Row
			frm1.vspdData1.Col = C_OrdNo
			frm1.KeyProdOrdNo.value = Trim(frm1.vspdData1.Text)
			
			frm1.vspdData1.Col = C_OprNo
			frm1.KeyOprNo.value = Trim(frm1.vspdData1.Text) 
			
			If DbQuery2 = False Then	
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
    If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Or lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
	End If
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
    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey3 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
	End If
    
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
    If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then
		If lgStrPrevKey4 <> "" Or lgStrPrevKey5 <> "" Then						'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery3 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
	End If
    
    
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

'========================================================================================
' Function Name : vspdData3_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData3_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData3.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
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
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

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
    If gActiveSpdSheet.Id = "A" Then Call InitSpreadCombobox
    Call ggoSpread.ReOrderingSpreadData
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
	
	If frm1.txtWcCd.value = "" Then
		frm1.txtWcNm.value = ""
	End If
	
	If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function	
	
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
	End If																'��: Query db data
       
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
	On Error Resume Next
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
    Call parent.FncExport(Parent.C_SINGLEMULTI)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001						
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)	
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)				
		strVal = strVal & "&txtProdFromDt=" & Trim(frm1.hProdFromDt.value)		
		strVal = strVal & "&txtProdToDt=" & Trim(frm1.hProdToDt.value)			
		strVal = strVal & "&cboOrderType=" & Trim(frm1.hOrderType.value)		
		strVal = strVal & "&cboOrderStatus=" & Trim(frm1.hOrderStatus.value)	
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
	Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001						
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)			
		strVal = strVal & "&txtProdFromDt=" & Trim(frm1.txtProdFromDt.text)			
		strVal = strVal & "&txtProdToDt=" & Trim(frm1.txtProdToDt.text)				
		strVal = strVal & "&cboOrderType=" & Trim(frm1.cboOrderType.value)
		strVal = strVal & "&cboOrderStatus=" & Trim(frm1.cboOrderStatus.value)	
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)	
		strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
	End If    

    Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          	'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(Byval LngMaxRow)
    '-----------------------
    'Reset variables area
    '-----------------------
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
		frm1.vspdData1.Col = C_OrdNo
		frm1.vspdData1.Row = 1
		frm1.KeyProdOrdNo.value = Trim(frm1.vspdData1.Text)
		
		frm1.vspdData1.Col = C_OprNo
		frm1.vspdData1.Row = 1
		frm1.KeyOprNo.value = Trim(frm1.vspdData1.Text)
		
		lgOldRow = 1
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Function
		End If
    End If

    lgIntFlgMode = Parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
    Call InitData(LngMaxRow)
    
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field
    Call SetToolBar("11000000000111")
    
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
'========================================================================================

Function DbQuery2() 
    
    DbQuery2 = False                                                         			'��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal

	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3					
	strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.KeyProdOrdNo.value)
	strVal = strVal & "&txtOprNo=" & Trim(frm1.KeyOprNo.value)	

    Call RunMyBizASP(MyBizASP, strVal)									

    DbQuery2 = True                                                     

End Function


'========================================================================================
' Function Name : DbQuery3
' Function Desc : Spread 2 Scroll 
'========================================================================================

Function DbQuery3() 
    
    DbQuery3 = False                                                         			'��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    
	strVal = BIZ_PGM_QRY3_ID & "?txtMode=" & Parent.UID_M0001						'��: 
	strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
	strVal = strVal & "&lgStrPrevKey5=" & lgStrPrevKey5
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.KeyProdOrdNo.value)
	strVal = strVal & "&txtOprNo=" & Trim(frm1.KeyOprNo.value)	
	
    Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    DbQuery3 = True                                                          	'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 

End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

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
					<TD WIDTH=* align=right><A href="vbscript:OpenProdRef()">��������</A> | <A href="vbscript:OpenRcptRef()">�԰���</A> | <A href="vbscript:OpenConsumRef()">��ǰ�Һ񳻿�</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="����ø�"></TD>
									<TD CLASS=TD5 NOWRAP>�۾���ȹ����</TD>								
									<TD CLASS=TD6>
										<script language =javascript src='./js/p4115ma1_I842118740_txtProdFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p4115ma1_I283148258_txtProdtODt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>	
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
									<TD CLASS=TD5 NOWRAP>���ñ���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderType" ALT="���ñ���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
			 						<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcPopup()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="14"></TD>
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
									<script language =javascript src='./js/p4115ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4115ma1_B_vspdData2.js'></script>
								</TD>
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4115ma1_C_vspdData3.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderType" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hProdToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24"><INPUT TYPE=HIDDEN NAME="KeyProdOrdNo" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyOprNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
