
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Production Order Management
'*  3. Program ID			: p4116ma1.asp
'*  4. Program Name			: Convert Production Order
'*  5. Program Desc			: Convert Production Order to Purchase Order (Outsourcing)
'*  6. Comproxy List		: 
'*	   Biz ASP  List		: +p4116mb1.asp		List Production Order Header
'*							  +p4116mb2.asp		Manage Conversion Production Order
'*							  
'*  7. Modified date(First)	: 2002/03/08
'*  8. Modified date(Last)	: 2003/05/20
'*  9. Modifier (First)		: Chen, Jaehyun
'* 10. Modifier (Last)		: Chen, Jaehyun
'* 11. Comment				:
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
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

Option Explicit									'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID			= "p4116mb1.asp"			'��: List Production Order Header
Const BIZ_PGM_SAVE_ID			= "p4116mb2.asp"			'��: Manage Production Order Header

Dim C_ProdtOrderNo	
Dim C_ItemCode			
Dim C_ItemName			
Dim C_Spec
Dim C_OrderQty			
Dim C_OrderUnit			
Dim C_OrderQtyInBaseUnit	
Dim C_BaseUnit				
Dim C_ProdQtyInOrderUnit	
Dim C_OrderStatus		
Dim C_RoutingNo			
Dim C_PlanStartDt		
Dim C_PlanEndDt			
Dim C_SLCD					
Dim C_SLNM					
Dim C_TrackingNo1		
Dim C_InitOprResult	
Dim C_PRRequireQty	
Dim C_PRRequireUnit	
Dim C_PRRequireDay	
Dim C_PRDelieveryDay
Dim C_PRNo					
Dim C_PurOrgCD			
Dim C_PurOrgCDPopup	
Dim C_PurOrgNM			
Dim C_PurSLCD				
Dim C_PurSLCDPopup	
Dim C_PurSLNM				
Dim C_Remark				
Dim C_DeptCD				
Dim C_DeptCDPopup		
Dim C_DeptNM				
Dim C_ReqPerson		
Dim C_OrderType
Dim C_ItemGroupCd
Dim C_ItemGroupNm

Dim LocSvrDate
LocSvrDate = "<%=GetSvrDate%>"
'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================

'========================================================================================================
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntGrpCount	' Group View Size�� ������ ���� 
Dim lgIntFlgMode		' Variable is for Operation Status
Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgBlnFlgChgValue
Dim lgBlnFlgClick
Dim lgBlnFlgCncl
Dim lgSortKey
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6									 '  For InitCombobox 
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop						' Popup
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

    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgBlnFlgChgValue = False 
    lgBlnFlgClick = False
    lgBlnFlgCncl = False
    lgSortKey = 1

End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display  - p1017
'=========================================================================================================
Sub InitComboBox()

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderType, lgF0, lgF1, Chr(11))

	'****************************
	'List Minor code(Order Status)
	'****************************
	'List Order status except Closed Order
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " And MINOR_CD <> " & FilterVar("CL", "''", "S") & " ", _
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderStatus, lgF0, lgF1, Chr(11))
 
 	frm1.cboOrderType.value = ""
	frm1.cboOrderStatus.value = ""

End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()	
	frm1.txtProdFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -10,LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtProdToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 20, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call InitSpreadPosVariables()        
    With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		
		.vspdData.ReDraw = False
		
		.vspdData.MaxCols = C_ItemGroupNm + 1
		.vspdData.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 18,,,18,2
		ggoSpread.SSSetEdit		C_ItemCode, "ǰ��", 18,,,18,2
		ggoSpread.SSSetEdit		C_ItemName, "ǰ���", 25
		ggoSpread.SSSetEdit		C_Spec,		"�԰�", 25
		ggoSpread.SSSetFloat	C_OrderQty,"��������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_OrderUnit, "��������", 8,,,3,2
		ggoSpread.SSSetFloat	C_OrderQtyInBaseUnit, "���ؼ���",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_BaseUnit, "���ش���", 8,,,3
		ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "��������",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_OrderStatus, "���û���", 10
		ggoSpread.SSSetEdit		C_RoutingNo, "�����", 10,,,10,2
		ggoSpread.SSSetDate 	C_PlanStartDt, "����������", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_PlanEndDt, "�ϷΌ����", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_SLCD, "â��", 10,,,7,2
		ggoSpread.SSSetEdit		C_SLNM, "â���", 20
		ggoSpread.SSSetEdit		C_TrackingNo1, "Tracking No", 25
		ggoSpread.SSSetFloat	C_InitOprResult, "�ʰ�������",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PRNo, "���ſ�û��ȣ", 10
		ggoSpread.SSSetDate 	C_PRRequireDay, "��û��", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_PRDelieveryDay, "�ʿ���", 11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_PRRequireQty, "��û����",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PRRequireUnit, "��û����", 5,,,3
		ggoSpread.SSSetEdit		C_PurOrgCD, "��������", 10
		ggoSpread.SSSetButton 	C_PurOrgCDPopup
		ggoSpread.SSSetEdit		C_PurOrgNM, "����������", 20
		ggoSpread.SSSetEdit		C_PurSLCD, "�԰�â��", 10
		ggoSpread.SSSetButton 	C_PurSLCDPopup
		ggoSpread.SSSetEdit		C_PurSLNM, "�԰�â���", 20
		ggoSpread.SSSetEdit		C_Remark, "���", 10
		ggoSpread.SSSetEdit		C_DeptCD, "��û�μ�", 10
		ggoSpread.SSSetButton 	C_DeptCDPopup
		ggoSpread.SSSetEdit		C_DeptNM, "��û�μ���", 10
		ggoSpread.SSSetEdit		C_ReqPerson, "��û��", 10
		ggoSpread.SSSetEdit		C_OrderType, "���ñ���", 10
		ggoSpread.SSSetEdit 	C_ItemGroupCd, "ǰ��׷�",	15
		ggoSpread.SSSetEdit		C_ItemGroupNm, "ǰ��׷��", 30
		
		'Call ggoSpread.MakePairsColumn(,)
 		Call ggoSpread.SSSetColHidden( C_InitOprResult, C_ReqPerson, True)
 		Call ggoSpread.SSSetColHidden( .vspdData.MaxCols, .vspdData.MaxCols, True)
		
		ggoSpread.SSSetSplit2(2)
		
		.vspdData.ReDraw = True
    
    End With
    
    Call SetSpreadLock()
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData 
    
		ggoSpread.Source = .vspdData
		.Redraw = False

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetProtected C_ProdtOrderNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemCode,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemName,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OrderQty,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OrderUnit,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OrderQtyInBaseUnit,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BaseUnit,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProdQtyInOrderUnit,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OrderStatus,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RoutingNo,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanStartDt,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlanEndDt,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SLCd,					pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SLNm,					pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrackingNo1,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InitOprResult,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PRRequireUnit,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PurOrgNm,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PurSLNm,					pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DeptNm,					pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OrderType,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemGroupCd,				pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemGroupNm,				pvStartRow, pvEndRow

		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
    
		.Redraw = True
    
    End With
End Sub


'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_ProdtOrderNo			= 1
	C_ItemCode				= 2
	C_ItemName				= 3
	C_Spec					= 4
	C_OrderQty				= 5
	C_OrderUnit				= 6
	C_OrderQtyInBaseUnit	= 7
	C_BaseUnit				= 8
	C_ProdQtyInOrderUnit    = 9
	C_OrderStatus			= 10
	C_RoutingNo				= 11
	C_PlanStartDt			= 12
	C_PlanEndDt				= 13
	C_SLCD					= 14
	C_SLNM					= 15
	C_TrackingNo1			= 16
	C_InitOprResult			= 17
	C_PRRequireQty			= 18
	C_PRRequireUnit			= 19
	C_PRRequireDay			= 20
	C_PRDelieveryDay		= 21
	C_PRNo					= 22
	C_PurOrgCD				= 23
	C_PurOrgCDPopup			= 24
	C_PurOrgNM				= 25
	C_PurSLCD				= 26
	C_PurSLCDPopup			= 27
	C_PurSLNM				= 28
	C_Remark				= 29
	C_DeptCD				= 30
	C_DeptCDPopup			= 31
	C_DeptNM				= 32
	C_ReqPerson				= 33
	C_OrderType				= 34
	C_ItemGroupCd			= 35
	C_ItemGroupNm			= 36

End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
		C_ProdtOrderNo			= iCurColumnPos(1)
		C_ItemCode				= iCurColumnPos(2)
		C_ItemName				= iCurColumnPos(3)
		C_Spec					= iCurColumnPos(4)
		C_OrderQty				= iCurColumnPos(5)
		C_OrderUnit				= iCurColumnPos(6)
		C_OrderQtyInBaseUnit	= iCurColumnPos(7)
		C_BaseUnit				= iCurColumnPos(8)
		C_ProdQtyInOrderUnit    = iCurColumnPos(9)
		C_OrderStatus			= iCurColumnPos(10)
		C_RoutingNo				= iCurColumnPos(11)
		C_PlanStartDt			= iCurColumnPos(12)
		C_PlanEndDt				= iCurColumnPos(13)
		C_SLCD					= iCurColumnPos(14)
		C_SLNM					= iCurColumnPos(15)
		C_TrackingNo1			= iCurColumnPos(16)
		C_InitOprResult			= iCurColumnPos(17)
		C_PRRequireQty			= iCurColumnPos(18)
		C_PRRequireUnit			= iCurColumnPos(19)
		C_PRRequireDay			= iCurColumnPos(20)
		C_PRDelieveryDay		= iCurColumnPos(21)
		C_PRNo					= iCurColumnPos(22)
		C_PurOrgCD				= iCurColumnPos(23)
		C_PurOrgCDPopup			= iCurColumnPos(24)
		C_PurOrgNM				= iCurColumnPos(25)
		C_PurSLCD				= iCurColumnPos(26)
		C_PurSLCDPopup			= iCurColumnPos(27)
		C_PurSLNM				= iCurColumnPos(28)
		C_Remark				= iCurColumnPos(29)
		C_DeptCD				= iCurColumnPos(30)
		C_DeptCDPopup			= iCurColumnPos(31)
		C_DeptNM				= iCurColumnPos(32)
		C_ReqPerson				= iCurColumnPos(33)
		C_OrderType				= iCurColumnPos(34)
		C_ItemGroupCd			= iCurColumnPos(35)
		C_ItemGroupNm			= iCurColumnPos(36)

 	End Select
 
End Sub


'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************

'------------------------------------------  OpenCondPlant()  -------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"					' �˾� ��Ī 
	arrParam(1) = "B_PLANT"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"						' TextBox ��Ī 
	
	arrField(0) = "PLANT_CD"						' Field��(0)
	arrField(1) = "PLANT_NM"						' Field��(1)
	
	arrHeader(0) = "����"					     ' Header��(0)
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
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
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

'------------------------------------------  OpenProdOrderNo()  ---------------------------------------
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdToDt.Text
	arrParam(3) = "OP"
	arrParam(4) = "RLST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = "" 'Trim(frm1.cboOrderType.value)
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
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

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingInfo()

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
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

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd2()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd2(Byval strCode, Byval Row)

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

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ""	' Where Condition
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
		Call SetSLCd2(arrRet, Row)
	End If
	
End Function


<!-- '------------------------------------------  OpenORG()  -------------------------------------------------
'	Name : OpenORG()
'	Description :
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True or UCase(frm1.txtOrgCd.ClassName)=UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"			
	arrParam(1) = "B_Pur_Org"			
	
	arrParam(2) = Trim(frm1.txtOrgCd.Value)
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "��������"				
	
    arrField(0) = "PUR_ORG"					
    arrField(1) = "PUR_ORG_NM"				
    
    arrHeader(0) = "��������"			
    arrHeader(1) = "����������"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtOrgCd.Value = arrRet(0)
		frm1.txtOrgNm.Value = arrRet(1)
		Call CopyTextboxtoGrid()      'copy text box to multigrid
		lgBlnFlgChgValue = True
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtOrgCd.focus
	
End Function

<!-- '------------------------------------------  OpenORG2()  -------------------------------------------------
'	Name : OpenORG2()
'	Description :
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenORG2(Byval strCode, Byval Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"			
	arrParam(1) = "B_Pur_Org"			
	
	arrParam(2) = Trim(strCode)
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "��������"				
	
    arrField(0) = "PUR_ORG"					
    arrField(1) = "PUR_ORG_NM"				
    
    arrHeader(0) = "��������"			
    arrHeader(1) = "����������"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.vspdData.row = Row
		frm1.vspdData.col = C_PurOrgCd
		frm1.vspdData.text= arrRet(0)
		frm1.vspdData.col = C_PurOrgNm
		frm1.vspdData.text= arrRet(1)
		
		Call CopyGridtoTextbox(Row)      'copy from multigrid to textbox
		lgBlnFlgChgValue = True
	End If
	
End Function

<!-- '------------------------------------------  OpenDept()  -------------------------------------------------
'	Name : OpenDept()
'	Description :  OpenDept PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenDept()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtDeptCd.ClassName)=UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��û�μ�"	
	arrParam(1) = "B_ACCT_DEPT"				
	
	arrParam(2) = Trim(frm1.txtDeptCd.Value)
'	arrParam(3) = Trim(frm1.txtDeptNm.Value)
	
	arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(parent.gChangeOrgId, "''", "S") & ""
	arrParam(5) = "��û�μ�"			
	
    arrField(0) = "DEPT_CD"	
    arrField(1) = "DEPT_NM"	
    
    arrHeader(0) = "��û�μ�"		
    arrHeader(1) = "��û�μ���"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtDeptCd.Value = arrRet(0)
		frm1.txtDeptNm.Value = arrRet(1)
		Call CopyTextboxtoGrid()      'copy text box to multigrid
		lgBlnFlgChgValue = True
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtDeptCd.focus
	
End Function

<!-- '------------------------------------------  OpenDept2()  -------------------------------------------------
'	Name : OpenDept2()
'	Description :  OpenDept PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenDept2(Byval strCode, Byval Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��û�μ�"	
	arrParam(1) = "B_ACCT_DEPT"				
	
	arrParam(2) = Trim(StrCode)
'	arrParam(3) = Trim(frm1.txtDeptNm.Value)
	
	arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(parent.gChangeOrgId, "''", "S") & ""
	arrParam(5) = "��û�μ�"			
	
    arrField(0) = "DEPT_CD"	
    arrField(1) = "DEPT_NM"	
    
    arrHeader(0) = "��û�μ�"		
    arrHeader(1) = "��û�μ���"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.row = Row
		frm1.vspdData.col = C_DeptCd
		frm1.vspdData.text= arrRet(0)
		frm1.vspdData.col = C_DeptNm
		frm1.vspdData.text= arrRet(1)
		Call CopyGridtoTextbox(Row)      'copy from multigrid to textbox
		lgBlnFlgChgValue = True
	End If	
	
End Function

'------------------------------------------  OpenStorage()  -------------------------------------------------
'	Name : OpenStorage()
'	Description :  OpenDept PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenStorage()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtStorageCd.className) = UCase(parent.UCN_PROTECTED) then Exit Function
	if Trim(frm1.txtPlantCd.value)="" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		Exit function	
	End if 
	
	IsOpenPop = True

	arrParam(0) = "�԰�â��"			
	arrParam(1) = "B_Storage_location,B_Plant"	
	
	arrParam(2) = Trim(frm1.txtStorageCD.Value)	
	
	arrParam(4) = "B_Storage_location.Plant_Cd=B_Plant.Plant_Cd And "	
	arrParam(4) = arrParam(4) & "B_Plant.Plant_Cd= " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ""
	arrParam(5) = "�԰�â��"					
	
    arrField(0) = "B_Storage_location.Sl_Cd"	
    arrField(1) = "B_Storage_location.Sl_Nm"	
    
    arrHeader(0) = "�԰�â��"				
    arrHeader(1) = "�԰�â���"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetStorage(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtStorageCD.focus
		
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
	
	iCalledAspName = AskPRAspName("P4311RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4311RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		<%'��: ��ȸ ���� ����Ÿ %>

   	With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

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
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
   	With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenConvHistory()  ------------------------------------------
'	Name : OpenConvHistory()
'	Description : Conversion History PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConvHistory()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4116RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4116RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
   	With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
    frm1.txtPlantCd.Value    = arrRet(0)		
    frm1.txtPlantNm.Value    = arrRet(1)
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

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function    

'------------------------------------------  SetStorage()  -----------------------------------------
'	Name : SetStorage()
'	Description : Storage Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetStorage(byval arrRet)
	frm1.txtStorageCd.Value    = arrRet(0)		
	frm1.txtStorageNm.Value    = arrRet(1)
	Call CopyTextboxtoGrid()      'copy text box to multigrid		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetSLCd2()  --------------------------------------------------
'	Name : SetSLCd2()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd2(byval arrRet, Byval Row)
    With frm1
	   	.vspdData.Row = Row
	   	.vspdData.Col = C_PurSLCD
	   	.vspdData.Text = arrRet(0)
	   	
	   	.vspdData.Col = C_PurSLNm
	   	.vspdData.Text = arrRet(1)	
	   	Call CopyGridtoTextbox(Row)      'copy from multigrid to textbox   	
	End With
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'========================================================================================
' Function Name : CopyTextboxToGrid
' Function Desc : Textbox �� �ִ� �����͸� Grid�� ���� 
'========================================================================================
Function CopyTextboxToGrid()
	
	Dim lRow
	
	If lgBlnFlgCncl = True Then Exit Function
	
	If frm1.vspdData.ActiveRow < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		
		.Redraw = False
		
		lRow = .ActiveRow
	    .Row = lRow
		.col = 0
		
		If .Text <> ggoSpread.UpdateFlag  Then
			.Text = ggoSpread.UpdateFlag
			ggoSpread.SpreadUnLock C_PRRequireDay, lRow, C_PRRequireDay, lRow
			ggoSpread.SpreadUnLock C_PRDelieveryDay, lRow, C_PRDelieveryDay, lRow
			ggoSpread.SpreadUnLock C_PRNo, lRow, C_PRNo, lRow
			ggoSpread.SpreadUnLock C_PRRequireQty, lRow, C_PRRequireQty, lRow
			ggoSpread.SpreadUnLock C_PurOrgCD, lRow, C_PurOrgCD, lRow
			ggoSpread.SpreadUnLock C_PurOrgCDPopup, lRow, C_PurOrgCDPopup, lRow
			ggoSpread.SpreadUnLock C_PurSLCD, lRow, C_PurSLCD, lRow
			ggoSpread.SpreadUnLock C_PurSLCDPopup, lRow, C_PurSLCDPopup, lRow
			ggoSpread.SpreadUnLock C_Remark, lRow, C_Remark, lRow
			ggoSpread.SpreadUnLock C_DeptCD, lRow, C_DeptCD, lRow
			ggoSpread.SpreadUnLock C_DeptCDPopup, lRow, C_DeptCDPopup, lRow
			ggoSpread.SpreadUnLock C_ReqPerson, lRow, C_ReqPerson, lRow
			ggoSpread.SSSetRequired C_PRRequireDay, lRow, lRow
			ggoSpread.SSSetRequired C_PRDelieveryDay, lRow, lRow
			ggoSpread.SSSetRequired C_PRRequireQty, lRow, lRow
			ggoSpread.SSSetRequired C_Remark, lRow, lRow
			ggoSpread.SSSetRequired C_PurOrgCD, lRow, lRow
			ggoSpread.SSSetRequired C_PurSLCD, lRow, lRow
		
		End If 
		.Col = C_PRNo
	    .Value = frm1.txtReqNo.value
	    .Col = C_PRRequireDay
	    .Text = frm1.txtReqDt.Text
	    .Col = C_PRDelieveryDay
	    .Text = frm1.txtDlvyDt.Text
	    .Col = C_PRRequireQty
	    .Text = frm1.txtReqQty.Text
	    .Col = C_PRRequireUnit
	    .Value = frm1.txtReqUnitCd.value
	    .Col = C_PurOrgCd
	    .Value = frm1.txtOrgCd.value
	    .Col = C_PurOrgNM
	    .Value = frm1.txtOrgNm.value
	    .Col = C_PurSLCD
	    .Value = frm1.txtStorageCd.value
		.Col = C_PurSLNM
	    .Value = frm1.txtStorageNm.value
	    .Col = C_Remark
	    .Value = frm1.txtRemark.value
	    .Col = C_DeptCD
	    .Value = frm1.txtDeptCd.value
		.Col = C_DeptNM
	    .Value = frm1.txtDeptNm.value
	    .Col = C_ReqPerson
	    .Value = frm1.txtEmpCd.value

		.Redraw = True
		
	End With

    lgBlnFlgChgValue = True

End Function

'========================================================================================
' Function Name : CopyGridToTextbox
' Function Desc : Grid�� �ִ� ������ Textbox �� �ű� 
'========================================================================================
Function CopyGridToTextbox(ByVal Row)

	With frm1.vspddata
		.Row = Row
				
		.Col = C_PRNo
		frm1.txtReqNo.value = .Text
		.Col = C_PRRequireQty
		frm1.txtReqQty.value = .Text
		.Col = C_PRRequireUnit
		frm1.txtReqUnitCd.value = .Text
		.Col = C_PurOrgCD
		frm1.txtOrgCd.value = .Text
		.Col = C_PurOrgNM
		frm1.txtOrgNm.value = .Text
		.Col = C_PurSLCD
		frm1.txtStorageCd.value = .Text
		.Col = C_PurSLNM
		frm1.txtstorageNm.value = .Text
		.Col = C_Remark
		frm1.txtRemark.value = .Text
		.Col = C_DeptCD
		frm1.txtDeptCd.value = .Text
		.Col = C_DeptNM
		frm1.txtDeptNm.value = .Text
		.Col = C_ReqPerson
		frm1.txtEmpCd.value = .Text
		.Col = C_PRRequireDay
		If .Text = "" Then
			frm1.txtReqDt.Text = UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)
		Else
			frm1.txtReqDt.Text = .Text
		End If
		.Col = C_PRDelieveryDay
		If .Text = "" Then	
			.Col = C_PlanEndDt
			If CompareDateByFormat( frm1.txtReqDt.Text,.text,"��û��","�ϷΌ����","970025",parent.gDateFormat,parent.gComDateType,False) = False Then			
				frm1.txtDlvyDt.Text = ""
			Else
				frm1.txtDlvyDt.Text = .Text
			End If
		Else
		    frm1.txtDlvyDt.Text = .Text 
		End If	
	End With
	
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
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     				'��: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "Q")                                          			'��: Lock  Suitable  Field
    Call InitSpreadSheet                                                    				'��: Setup the Spread sheet

	Call SetDefaultVal
    Call InitVariables																		'��: Initializes local global variables
    Call InitComboBox()

    Call SetToolBar("11000000000111")														'��: ��ư ���� ���� 
	
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.Value = parent.gPlant
		frm1.txtPlantNm.Value = parent.gPlantNm
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

'=======================================================================================================
'   Event Name : txtReqNo_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtReqNo_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
    If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtReqDt_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtReqDt_Change()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
    
	If lgBlnFlgClick <> True And CheckDateFormat(frm1.txtDlvyDt.text, parent.gDateFormat) = True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDlvyDt_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtDlvyDt_Change()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
    
	If lgBlnFlgClick <> True  And CheckDateFormat(frm1.txtReqDt.text, parent.gDateFormat) = True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtReqQty_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtReqQty_Change()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtReqUnitCd_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtReqUnitCd_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtOrgCd_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtOrgCd_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtOrgNm_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtOrgNm_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtStorageCd_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtStorageCd_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtStorageNm_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtStorageNm_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtRemark_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtRemark_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtDeptCd_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtDeptCd_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data
		Call CopyTextboxToGrid()
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtDeptNm_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtDeptNm_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtEmpCd_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtEmpCd_onChange()
	If frm1.vspdData.Row < 1 Then 
		Exit Sub
    End if
	
	If lgBlnFlgClick <> True	Then	'for protecting changing multi-grid data 
		Call CopyTextboxToGrid()
	End If
	
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

	msgbox "AA"
	Dim	DtPRRequireDay, DtPRDelieveryDay, DtInvCloseDt
	Dim strYear,strMonth,strDay
	Dim strOrderQty, strInitOprResult
			
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	With frm1.vspdData 
	
	.Redraw = False
	
    Select Case Col
    
        Case C_ItemCode

        Case C_PRRequireQty
			.Col = C_OrderQty
			strOrderQty = UNICDbl(.text)
			.Col = C_InitOprResult
			strInitOprResult = UNICDbl(.text)
			
			.Col = C_PRRequireQty
			If UNICDbl(.Value) > 0 Then
				ggoSpread.SpreadUnLock C_PRRequireDay, Row, C_PRRequireDay, Row
				ggoSpread.SpreadUnLock C_PRDelieveryDay, Row, C_PRDelieveryDay, Row
				ggoSpread.SpreadUnLock C_PRRequireQty, Row, C_PRRequireQty, Row
				ggoSpread.SpreadUnLock C_PRNo, Row, C_PRNo, Row
				ggoSpread.SpreadUnLock C_PurOrgCD, Row, C_PurOrgCD, Row
				ggoSpread.SpreadUnLock C_PurOrgCDPopup, Row, C_PurOrgCDPopup, Row
				ggoSpread.SpreadUnLock C_PurSLCD, Row, C_PurSLCD, Row 
				ggoSpread.SpreadUnLock C_PurSLCDPopup, Row, C_PurSLCDPopup, Row
				ggoSpread.SpreadUnLock C_Remark, Row, C_Remark, Row
				ggoSpread.SpreadUnLock C_DeptCD, Row, C_DeptCD, Row
				ggoSpread.SpreadUnLock C_DeptCDPopup, Row, C_DeptCDPopup, Row
				ggoSpread.SpreadUnLock C_ReqPerson, Row, C_ReqPerson, Row
				ggoSpread.SSSetRequired C_PRRequireDay, Row, Row
				ggoSpread.SSSetRequired C_PRDelieveryDay, Row, Row
				ggoSpread.SSSetRequired C_PRRequireQty, Row, Row
				ggoSpread.SSSetRequired C_PurOrgCD, Row, Row
				ggoSpread.SSSetRequired C_PurSLCD, Row, Row
				ggoSpread.SSSetRequired C_Remark, Row, Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
				.Col = C_PRRequireDay
				.Text= ""
				.Col = C_PRDelieveryDay
				.Text= ""
				.Col = C_PRNo
				.Text= ""
				.Col = C_PRRequireQty
				.Text= ""
				.Col = C_PRRequireUnit
				.Text= ""
				.Col = C_PurOrgCD
				.Text= ""
				.Col = C_PurSLCD
				.Text= ""
				.Col = C_Remark
				.Text= ""
				.Col = C_ReqPerson
				.Text= ""
				ggoSpread.SpreadLock C_PRRequireDay, Row, C_PRRequireDay, Row
				ggoSpread.SpreadLock C_PRDelieveryDay, Row, C_PRDelieveryDay, Row
				ggoSpread.SpreadLock C_PRNo, Row, C_PRNo, Row
				ggoSpread.SpreadLock C_PurOrgCD, Row, C_PurOrgCD, Row
				ggoSpread.SpreadLock C_PurOrgCDPopup, Row, C_PurOrgCDPopup, Row
				ggoSpread.SpreadLock C_PurSLCD, Row, C_PurSLCD, Row
				ggoSpread.SpreadLock C_PurSLCDPopup, Row, C_PurSLCDPopup, Row
				ggoSpread.SpreadLock C_Remark, Row, C_Remark, Row
				ggoSpread.SpreadLock C_DeptCD, Row, C_DeptCD, Row
				ggoSpread.SpreadLock C_DeptCDPopup, Row, C_DeptCDPopup, Row
				ggoSpread.SpreadLock C_ReqPerson, Row, C_ReqPerson, Row
				ggoSpread.SSSetProtected C_PRRequireDay, Row, Row
				ggoSpread.SSSetProtected C_PRDelieveryDay, Row, Row
				ggoSpread.SSSetProtected C_PRNo, Row, Row
				ggoSpread.SSSetProtected C_PurOrgCD, Row, Row
				ggoSpread.SSSetProtected C_PurOrgCDPopup, Row, Row
				ggoSpread.SSSetProtected C_PurSLCD, Row, Row
				ggoSpread.SSSetProtected C_PurSLCDPopup, Row, Row
				ggoSpread.SSSetProtected C_Remark, Row, Row
				ggoSpread.SSSetProtected C_DeptCD, Row, Row
				ggoSpread.SSSetProtected C_DeptCDPopup, Row, Row
				ggoSpread.SSSetProtected C_ReqPerson, Row, Row
			End IF
			
        Case C_PRRequireDay

			.Col = C_PRDelieveryDay
			DtPRDelieveryDay = .Text
			.Col = C_PRRequireDay
			DtPRRequireDay = .Text
			
			If (DtPRRequireDay <> "" AND DtPRDelieveryDay <> "") _
				And CheckDateFormat(DtPRDelieveryDay, parent.gDateFormat) = True And CheckDateFormat(DtPRRequireDay, parent.gDateFormat) = True Then
				If CompareDateByFormat(DtPRRequireDay,DtPRDelieveryDay,"��û��","�ʿ���","970025",parent.gDateFormat,parent.gComDateType,True) = False  Then  
					.Col = C_PRRequireDay
					.Text = ""
					Exit Sub
				End If
			End If				


        Case C_PRDelieveryDay
        
			.Col = C_PRDelieveryDay
			DtPRDelieveryDay = .Text
			.Col = C_PRRequireDay
			DtPRRequireDay = .Text
			If (DtPRDelieveryDay <> "" and DtPRRequireDay <> "")_
				And CheckDateFormat(DtPRDelieveryDay, parent.gDateFormat) = True And CheckDateFormat(DtPRRequireDay, parent.gDateFormat) = True Then
				If CompareDateByFormat(DtPRRequireDay,DtPRDelieveryDay,"��û��","�ʿ���","970025",parent.gDateFormat,parent.gComDateType,True) = False Then  
					.Col = C_PRDelieveryDay
					.Text = ""
					Exit Sub
				End If
		    End If
		    
		    If DtPRDelieveryDay <> "" _
				And CheckDateFormat(DtPRDelieveryDay, parent.gDateFormat) Then
				If CompareDateByFormat(UniConvDateAToB(LocSvrDate, parent.gServerDateFormat,parent.gDateFormat),DtPRDelieveryDay,"������","�ʿ���","970025",parent.gDateFormat,parent.gComDateType,False) = False Then  
					.Col = C_PRDelieveryDay
					.Text = ""
					Call DisplayMsgBox("172120","X","X","X")
					Exit Sub
				End If
		    End If

    End Select
    
    .Redraw = True
    
    End With

End Sub

'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================

Sub vspdData_EditChange(ByVal Col , ByVal Row )

    Dim DblNetAmt, DblVatAmt, DblNetLocAmt, DblVatLocAmt 

	With frm1.vspdData 

    End With
                
End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )	
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	End If
	
	lgBlnFlgClick = True						' This variable is for protect change of the vspdData 
	
    With frm1.vspdData
		'----------------------
		'Column Split
		'----------------------
		gMouseClickStatus = "SPC"
		Set gActiveSpdSheet = frm1.vspdData
 	
 		If Row <= 0 Then
 			ggoSpread.Source = frm1.vspdData 
 			If lgSortKey = 1 Then
 				ggoSpread.SSSort Col					'Sort in Ascending
 				lgSortKey = 2
 			Else
 				ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 				lgSortKey = 1
 			End If
 			
 			If frm1.vspdData.ActiveRow > 0 Then 
 				Call CopyGridToTextbox(frm1.vspdData.ActiveRow)
 			End If	
		Else
 			'------ Developer Coding part (Start)
 			Call CopyGridToTextbox(Row)
		 	'------ Developer Coding part (End)
 		End If
 		
 	
    End With
    lgBlnFlgClick = False
End Sub


'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If frm1.vspdData.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If
	
	Call vspdData_Click(NewCol, NewRow)
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub
 
'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspddata_KeyPress(index , KeyAscii )

End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	 '----------  Coding part  -------------------------------------------------------------

    End With

End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------
	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  1
				.Col = Col
				intIndex = .Value
				.Col = C_BillFG
				.Value = intIndex
		End Select
	End With
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If  
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	 With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
			If Row < 1 Then Exit Sub
		Select Case Col
			Case C_PurSLCDPopup
				.ReDraw = false
				.Row = Row
				.Col = C_PurSLCD
				Call OpenSLCD2(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_PurSLCD,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				.ReDraw = True
			
			Case C_DeptCDPopup
				.ReDraw = false
				.Col = C_DeptCD
				.Row = Row
				Call OpenDept2(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_DeptCD,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				.ReDraw = True
				
			Case C_PurOrgCDPopup
				.ReDraw = false
				.Col = C_PurOrgCD
				.Row = Row
				Call OpenOrg2(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_PurOrgCD,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				.ReDraw = True	
		End Select
	End With    
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
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
'   Event Name : txtToDt_DblClick(Button)
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
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� MainQuery�Ѵ�.
'=======================================================================================================
Sub txtProdFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� MainQuery�Ѵ�.
'=======================================================================================================
Sub txtProdToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub


'=======================================================================================================
'   Event Name : txtReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqDt.Action = 7
        Call CopyTextboxtoGrid()
        Call SetFocusToDocument("M")
		Frm1.txtReqDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDlyDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDlvyDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDlvyDt.Action = 7
        Call CopyTextboxtoGrid()
        Call SetFocusToDocument("M")
		Frm1.txtDlvyDt.Focus
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
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData										'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then									'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdToDt) = False Then Exit Function
   
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

	frm1.hPlantCd.value		= frm1.txtPlantCd.value
	frm1.hItemCd.value		= frm1.txtItemCd.value
	frm1.hProdOrderNo.value	= frm1.txtProdOrderNo.value
	frm1.hProdFromDt.value	= frm1.txtProdFromDt.Text
	frm1.hProdToDt.value	= frm1.txtProdToDt.Text
    frm1.hOrderStatus.value	= frm1.cboOrderStatus.value
	frm1.hTrackingNo.value	= frm1.txtTrackingNo.value

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function															'��: Query db data
       
    FncQuery = True															'��: Processing is OK
   
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
    Dim i 
    Dim strOrderQty
    Dim strInitOprResult
    Dim strRequireQty
    
    FncSave = False                                             '��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
   
   ggoSpread.Source = frm1.vspdData
   With frm1.vspdData
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then
				.Col = C_OrderQty
				strInitOprResult = .Text
				.Col = C_InitOprResult
				strOrderQty = .Text
				.Col = C_PRRequireQty
				strRequireQty = .Text 
				
				If UNICDbl(strRequireQty) <= 0  Then
					Call DisplayMsgBox("189804", "x", "x", "x")
					Exit Function
				End If
			End If
		Next
     End With 
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function						'��: Save db data
    
    FncSave = True                                              '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
            
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
	Dim lRow
	
    If frm1.vspdData.MaxRows < 1 Then Exit Function	
    
    lgBlnFlgCncl = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
    With frm1.vspdData
		lRow = .ActiveRow
		
		Call CopyGridToTextbox(lRow)
		
		.Col = 0
		.Row = lRow
		If	.Text <> ggoSpread.UpdateFlag  Then
			.Redraw = False
			ggoSpread.SpreadLock C_PRRequireDay, lRow, C_PRRequireDay, lRow
			ggoSpread.SpreadLock C_PRDelieveryDay, lRow, C_PRDelieveryDay, lRow
			ggoSpread.SpreadLock C_PRNo, lRow, C_PRNo, lRow
			ggoSpread.SpreadLock C_PRRequireQty, lRow, C_PRRequireQty, lRow
			ggoSpread.SpreadLock C_PurOrgCD, lRow, C_PurOrgCD, lRow
			ggoSpread.SpreadLock C_PurOrgCDPopup, lRow, C_PurOrgCDPopup, lRow
			ggoSpread.SpreadLock C_PurSLCD, lRow, C_PurSLCD, lRow
			ggoSpread.SpreadLock C_PurSLCDPopup, lRow, C_PurSLCDPopup, lRow
			ggoSpread.SpreadLock C_DeptCD, lRow, C_DeptCD, lRow
			ggoSpread.SpreadLock C_DeptCDPopup, lRow, C_DeptCDPopup, lRow
			ggoSpread.SpreadLock C_ReqPerson, lRow, C_ReqPerson, lRow
			ggoSpread.SpreadLock C_Remark, lRow, C_Remark, lRow
			
			ggoSpread.SSSetProtected C_PRRequireDay, lRow, lRow
			ggoSpread.SSSetProtected C_PRDelieveryDay, lRow, lRow
			ggoSpread.SSSetProtected C_PRNo, lRow, lRow
			ggoSpread.SSSetProtected C_PRRequireQty, lRow, lRow
			ggoSpread.SSSetProtected C_PurOrgCD, lRow, lRow
			ggoSpread.SSSetProtected C_PurOrgCDPopup, lRow, lRow
			ggoSpread.SSSetProtected C_PurSLCD, lRow, lRow
			ggoSpread.SSSetProtected C_PurSLCDPopup, lRow, lRow
			ggoSpread.SSSetProtected C_DeptCD, lRow, lRow
			ggoSpread.SSSetProtected C_DeptCDPopup, lRow, lRow
			ggoSpread.SSSetProtected C_ReqPerson, lRow, lRow
			ggoSpread.SSSetProtected C_Remark, lRow, lRow
			
			.Redraw = True
		End If
    End With
    
    lgBlnFlgCncl = False
    
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
    
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
    Call parent.FncExport(parent.C_SINGLEMULTI)												<%'��: ȭ�� ���� %>
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         <%'��:ȭ�� ����, Tab ���� %>
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
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
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

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'******************************************************************************************************%>
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    Err.Clear

    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtStartDt=" & Trim(frm1.hProdFromDt.value)
		strVal = strVal & "&txtEndDt=" & Trim(frm1.hProdToDt.value)
		strVal = strVal & "&txtOrderStatus=" & Trim(frm1.hOrderStatus.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.hOrderType.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtStartDt=" & Trim(frm1.txtProdFromDt.text)
		strVal = strVal & "&txtEndDt=" & Trim(frm1.txtProdToDt.text)
		strVal = strVal & "&txtOrderStatus=" & Trim(frm1.cboOrderStatus.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.cboOrderType.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
	End If    

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

 	Dim lRow
 	Dim LngRow
 	Dim Row

	lgBlnFlgClick = True
	
    Call ggoOper.LockField(Document, "N")
    Call ggoOper.SetReqAttr(frm1.txtReqNo,"D")
    Call ggoOper.SetReqAttr(frm1.txtDeptCd,"D")
    Call ggoOper.SetReqAttr(frm1.txtEmpCd,"D")
    
    Call SetToolBar("11001001000111")

	frm1.vspdData.ReDraw = False

    With frm1.vspdData

		If .MaxRows < 1 Then Exit Function
			
			Row = 1
			
			Call CopyGridToTextbox(Row)		

	End With
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	
	frm1.vspdData.ReDraw = True

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
	lgBlnFlgClick = False
	
End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery�� ���������� �ƴҰ�� 
'========================================================================================
Function DbQueryNotOk()	

	Call SetToolBar("11000000000111")														'��: ��ư ���� ���� 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_CMODE													'��: Indicates that current mode is Update mode

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows  
    Dim lGrpcnt 
    Dim strVal
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

    lGrpCnt = 1
    
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)				

	iTmpCUBufferCount = -1 
	
	strCUTotalvalLen = 0 
    
    DbSave = False                                                          '��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1
		.txtMode.value = parent.UID_M0002										'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
		
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1

	With frm1.vspdData


    For IntRows = 1 To .MaxRows
    
		.Row = IntRows
		.Col = 0

		Select Case .Text
	    
		    Case ggoSpread.UpdateFlag
				
				' =====> Interface Refers to the Const Bas P4A2 or IG1_Import_Group at this.MB2
				strVal = ""
				strVal = strVal & IntRows & iColSep			' 0. Row Number
				.Col = C_ProdtOrderNo
				strVal = strVal & Trim(.Text) & iColSep		' 1. Production Order No.
				.Col = C_Remark
				strVal = strVal & Trim(.Text) & iColSep		' 2. Remark
				.Col = C_PRNo
				strVal = strVal & Trim(.Text) & iColSep		' 3. PR No.
				.Col = C_PRRequireQty
				strVal = strVal & Trim(.Text) & iColSep		' 4. Req Qty 
				.Col = C_PRRequireUnit
				strVal = strVal & Trim(.Text) & iColSep		' 5. Req Unit
				.Col = C_PRRequireDay
				strVal = strVal & Trim(.Text) & iColSep		' 6. Req Dt
				.Col = C_DeptCD
				strVal = strVal & Trim(.Text) & iColSep		' 7. Req Dept
				.Col = C_ReqPerson
				strVal = strVal & Trim(.Text) & iColSep		' 8. Req Prsn
				.Col = C_PRDelieveryDay
				strVal = strVal & Trim(.Text) & iColSep		' 9. Dlvy Dt
				.Col = C_PurSLCD
				strVal = strVal & Trim(.Text) & iColSep		' 10. SL_CD
				strVal = strVal & "" & iColSep				' 11. Base Req Qty
				strVal = strVal & "" & iColSep				' 12. Base Req Unit
				strVal = strVal & "" & iColSep				' 13. Pur Grp
				.Col = C_TrackingNo1
				strVal = strVal & Trim(.Text) & iColSep		' 14. Tracking No.
				.Col = C_PurOrgCD
				strVal = strVal & Trim(.Text) & iRowSep		' 15. Pur Org
				
				lGrpCnt = lGrpCnt + 1

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
			         
		            
		End Select

    Next

    End With	
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	   
	End If   
	
	.txtMaxRows.value = lGrpCnt-1
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: �����Ͻ� ASP �� ���� 

    End With

    DbSave = True                                                           ' ��: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 

	Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�۾����ú�ȯ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenConvHistory()">��ȯ�̷�</A> | <A href="vbscript:OpenPartRef()">��ǰ����</A> | <A href="vbscript:OpenOprRef()">��������</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�۾���ȹ����</TD>
									<TD CLASS="TD6">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtProdFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������" id=OBJECT1></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtProdToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������" id=OBJECT2></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
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
									<TD CLASS=TD5 NOWRAP>���ñ���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderType" ALT="���ñ���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top><!-- ù��° �� ���� -->
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%" colspan=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
						</TR>
						<TR>
							<TD  colspan=4>
							<FIELDSET valign=top>
								<LEGEND>��ȯ�����Է�</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS="TD5">���ſ�û��ȣ</TD>
										<TD CLASS="TD6"><INPUT TYPE=TEXT ALT="��û��ȣ" NAME="txtReqNo"  SIZE=20 MAXLENGTH=18 tag="23NXXU"></TD>
										<TD CLASS="TD5" NOWRAP>��������</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="��������" MAXLENGTH=4 NAME="txtOrgCd" SIZE=10 MAXLENGTH=4 tag="23NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrg()">
														   <INPUT TYPE=TEXT Alt="��������" NAME="txtOrgNm" SIZE=20 tag="24X"></TD>
									</TR>	
									<TR>
										<TD CLASS="TD5" NOWRAP>��û��</TD>
										<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtReqDt" CLASS=FPDTYYYYMMDD tag="23N1" Title="FPDATETIME" ALT="��û��"></OBJECT>');</SCRIPT></TD>
										</TD>
										<TD CLASS="TD5" NOWRAP>�ʿ���</TD>
										<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyDt" CLASS=FPDTYYYYMMDD tag="23N1" Title="FPDATETIME" ALT="�ʿ���"></OBJECT>');</SCRIPT></TD>
										</TR>	
									<TR>	
										<TD CLASS="TD5" NOWRAP>��û��</TD>
										<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtReqQty" CLASS=FPDS100 tag="23X3Z" Title="FPDOUBLESINGLE" ALT="��û��"></OBJECT>');</SCRIPT></TD>										<TD CLASS="TD5" NOWRAP>��û����</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="��û����"  NAME="txtReqUnitCd" SIZE=10 MAXLENGTH=3 tag="24XXXX"></TD>
									</TR>	
									<TR>
										<TD CLASS="TD5" NOWRAP>�԰�â��</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="�԰�â��" NAME="txtStorageCd"  SIZE=10 MAXLENGTH=7 tag="23NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStorageCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenStorage()">
															   <INPUT TYPE=TEXT ALT="�԰�â��" NAME="txtstorageNm" SIZE=20 tag="24X"></TD>
										<TD CLASS="TD5" NOWRAP>���</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="���" NAME="txtRemark" SIZE=30 MAXLENGTH=20 tag="23XXXX"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>��û��</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="��û��"  NAME="txtEmpCd" MAXLENGTH=20 SIZE=20 tag="23N"></TD>
										<TD CLASS="TD5" NOWRAP>��û�μ�</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="��û�μ�" NAME="txtDeptCd" SIZE=10 MAXLENGTH=10  tag="23NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDept()">
														   <INPUT TYPE=TEXT Alt="��û�μ�" NAME="txtDeptNm" SIZE=20 tag="24x"></TD>
									</TR>
								</TABLE>	
							</FIELDSET>			
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderType" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hProdToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
