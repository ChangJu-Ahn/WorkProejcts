
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Master Production Scheduling
'*  3. Program ID			: p4215ma1.asp
'*  4. Program Name			: ���� Document ��ȸ 
'*  5. Program Desc			: ���� Document ��ȸ 
'*  6. Business ASP List	: p4215mb1.asp (Query Order Document)
'*  7. Modified date(First)	: 
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: Hong, EunSook
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				: 
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
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

Option Explicit                                                             '��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID = "p4215mb1.asp"			'��: Head Query �����Ͻ� ���� ASP�� 

Const C_SHEETMAXROWS = 30

Dim C_OprNo				'= 1
Dim C_WcCd				'= 2
Dim C_WcNm				'= 3
Dim C_JobCd				'= 4
Dim C_JobNm				'= 5
Dim C_InsideFlg			'= 6
Dim C_OrderStatus		'= 7
Dim C_DtlPlanStartDt	'= 8
Dim C_DtlPlanComptDt	'= 9
Dim C_DtlReleaseDt		'= 10
Dim C_GoodQty			'= 11
Dim C_BadQty			'= 12
Dim C_ProdtOrderUnit	'= 13
Dim C_Document			'= 14
Dim C_ItemCd			'= 15
Dim C_ItemNm			'= 16
Dim C_ProdtOrderQty		'= 17
Dim C_PlanStartDt		'= 18
Dim C_PlanComptDt		'= 19
Dim C_Routing			'= 20
Dim C_TrackingNo		'= 21
Dim C_orgDocument		'= 22

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim ihGridCnt                     'hidden Grid Row Count
Dim intItemCnt                    'hidden Grid Row Count
Dim IsOpenPop						' Popup


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

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
     
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
    lgSortKey = 1

End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************

'========================================================================================
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

Sub InitSpreadSheet()
        
    'Dim sList
    Call InitSpreadPosVariables()
    
    With frm1
    
    'sList = "1" & Chr(11) & "2" & Chr(12) & "3"
       
    ggoSpread.Source = .vspdData
    ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
    
    .vspdData.ReDraw = False
    
    .vspdData.MaxCols = C_orgDocument + 1
    .vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_OprNo,			"����", 10
    ggoSpread.SSSetEdit		C_WcCd,				"�۾���", 10
    ggoSpread.SSSetEdit		C_WcNm,				"�۾����", 20
    ggoSpread.SSSetEdit		C_JobCd,			"�����۾�", 10
    ggoSpread.SSSetEdit		C_JobNm,			"�����۾���", 20
    ggoSpread.SSSetEdit 	C_InsideFlg,		"����Ÿ��",10
    ggoSpread.SSSetEdit		C_OrderStatus,		"���û���", 10
    ggoSpread.SSSetDate		C_DtlPlanStartDt,	"����������", 11, 2, parent.gDateFormat    
    ggoSpread.SSSetDate		C_DtlPlanComptDt,	"�ϷΌ����", 11, 2, parent.gDateFormat    
    ggoSpread.SSSetDate		C_DtlReleaseDt,		"�۾�������", 11, 2, parent.gDateFormat    
	ggoSpread.SSSetFloat	C_GoodQty,			"��ǰ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"    
	ggoSpread.SSSetFloat	C_BadQty,			"�ҷ�����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
    ggoSpread.SSSetEdit		C_ProdtOrderUnit,	"��������", 8,,,3,2
    ggoSpread.SSSetEdit		C_Document,			" ", 100,,,100
    ggoSpread.SSSetEdit		C_ItemCd,			"ǰ��", 18
    ggoSpread.SSSetEdit		C_ItemNm,			"ǰ��� ", 25
    ggoSpread.SSSetEdit		C_ProdtOrderQty,	"��������", 10
    ggoSpread.SSSetEdit		C_PlanStartDt,		"����������", 10
    ggoSpread.SSSetEdit		C_PlanComptDt,		"�ϷΌ����", 10
    ggoSpread.SSSetEdit		C_Routing,			"�����", 10
    ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No.", 25
    ggoSpread.SSSetEdit		C_orgDocument,		" ", 100,,,100

    ihGridCnt = 0               'Hidden Counter
    intItemCnt = 0
    ggoSpread.Source = .vspdData
	
	'Call ggoSpread.MakePairsColumn(C_ItemCode,C_ItemSpec)
 	Call ggoSpread.SSSetColHidden(C_Document ,C_orgDocument , True)
 	Call ggoSpread.SSSetColHidden(.vspdData.MaxCols ,.vspdData.MaxCols , True)
	
	.vspdData.ReDraw = true
	
    ggoSpread.SSSetSplit2(2) 

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
    
    .Redraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.SSSetProtected C_OprNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcNm, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_JobCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_JobNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_InsideFlg, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_OrderStatus, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DtlPlanStartDt, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected C_DtlPlanComptDt,pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DtlReleaseDt,pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_GoodQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BadQty, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ProdtOrderUnit, pvStartRow, pvEndRow
    
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
	
	C_OprNo				= 1
	C_WcCd				= 2
	C_WcNm				= 3
	C_JobCd				= 4
	C_JobNm				= 5
	C_InsideFlg			= 6
	C_OrderStatus		= 7
	C_DtlPlanStartDt	= 8
	C_DtlPlanComptDt	= 9
	C_DtlReleaseDt		= 10
	C_GoodQty			= 11
	C_BadQty			= 12
	C_ProdtOrderUnit	= 13
	C_Document			= 14
	C_ItemCd			= 15
	C_ItemNm			= 16
	C_ProdtOrderQty		= 17
	C_PlanStartDt		= 18
	C_PlanComptDt		= 19
	C_Routing			= 20
	C_TrackingNo		= 21
	C_orgDocument		= 22

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
 
 		C_OprNo				= iCurColumnPos(1)
		C_WcCd				= iCurColumnPos(2)
		C_WcNm				= iCurColumnPos(3)
		C_JobCd				= iCurColumnPos(4)
		C_JobNm				= iCurColumnPos(5)
		C_InsideFlg			= iCurColumnPos(6)
		C_OrderStatus		= iCurColumnPos(7)
		C_DtlPlanStartDt	= iCurColumnPos(8)
		C_DtlPlanComptDt	= iCurColumnPos(9)
		C_DtlReleaseDt		= iCurColumnPos(10)
		C_GoodQty			= iCurColumnPos(11)
		C_BadQty			= iCurColumnPos(12)
		C_ProdtOrderUnit	= iCurColumnPos(13)
		C_Document			= iCurColumnPos(14)
		C_ItemCd			= iCurColumnPos(15)
		C_ItemNm			= iCurColumnPos(16)
		C_ProdtOrderQty		= iCurColumnPos(17)
		C_PlanStartDt		= iCurColumnPos(18)
		C_PlanComptDt		= iCurColumnPos(19)
		C_Routing			= iCurColumnPos(20)
		C_TrackingNo		= iCurColumnPos(21)
		C_orgDocument		= iCurColumnPos(22)

 	End Select
 
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
    
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
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"						' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"						' Field��(0)
    arrField(1) = "PLANT_NM"						' Field��(1)
    
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
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
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
	arrParam(3) = "OP"
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

'------------------------------------------  OpenOprCd()  -------------------------------------------------
'	Name : OpenOprCd()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprCd()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	If frm1.txtProdOrderNo.value = "" Then
		Call DisplayMsgBox("971012","X" , "����������ȣ","X")
		frm1.txtProdOrderNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	iCalledAspName = AskPRAspName("P4112PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4112PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Or UCase(frm1.txtOprCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdOrderNo.value
	arrParam(2) = "" 'frm1.txtOprCd.value

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetOprCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprCd.focus
	
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
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4411RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4411RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
		
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)			<%'��: ��ȸ ���� ����Ÿ %>
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		<%'��: ��ȸ ���� ����Ÿ %>
	
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

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)			<%'��: ��ȸ ���� ����Ÿ %>
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		<%'��: ��ȸ ���� ����Ÿ %>

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

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4412RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4412RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)			<%'��: ��ȸ ���� ����Ÿ %>
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		<%'��: ��ȸ ���� ����Ÿ %>

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function


'==========================================  2.4.3 Set Return Value()  =============================================
'	Name : Set Return Value()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 


'================================== ReleaseButton() ===================================
' Function Name : ReleaseButton
' Function Desc : ������ �����ϰ� ����������� ��ư�� �ٽ� Ȱ��ȭ��Ų��.
'========================================================================================

'Sub ReleaseButton(pButton)
'	pButton.disabled = False
'End Sub


'================================== ProtectButton() ===================================
' Function Name : ProtectButton
' Function Desc : ����� ��ư�� disalbe��Ų��.
'========================================================================================

'Sub ProtectButton(pButton)
'	pButton.disabled = True
'End Sub

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

'------------------------------------------  SetOprCd()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOprCd(byval arrRet)
	frm1.txtOprCd.Value    = arrRet(0)		
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

	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
	Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
	Call InitSpreadSheet                                                    '��: Setup the Spread sheet
	Call InitVariables                                                      '��: Initializes local global variables

	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtProdOrderNo.focus 
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

End Sub

'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================

Sub vspdData_EditChange(ByVal Col , ByVal Row )
                
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
 		
 	End If
	
	'------ Developer Coding part (Start)
 	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_Document
		frm1.txtDocument.Value = replace(.Text,chr(7), chr(13) &chr(10))
	End With
	'------ Developer Coding part (End)
	
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
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()

End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
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
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
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
   
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					'��: Clear Contents  Field 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData  
    Call InitVariables	
    '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing   
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
    Call parent.FncExport(parent.C_SINGLEMULTI)											'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                        '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
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
'******************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    
    DbQuery = False                                                         				'��: Processing is NG
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtOprCd=" & Trim(frm1.hOprCd.value)				'��: ��ȸ ���� ����Ÿ 
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtOprCd=" & Trim(frm1.txtOprCd.value)				'��: ��ȸ ���� ����Ÿ	
	End If
    Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          	'��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()															'��: ��ȸ ������ ������� 
	
	Call SetToolBar("11000000000111")											'��: ��ư ���� ���� 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
        
	 With frm1.vspdData

		.Row = 1

		.Col = C_Document
		frm1.txtDocument.Value = replace(.Text,chr(7),chr(13) &chr(10))
		.Col = C_ItemCd
		frm1.txtItemCd.Value = .Text
		.Col = C_ItemNm
		frm1.txtItemNm.Value = .Text
		.Col = C_ProdtOrderQty
		If .Text = "" or isnull(.text) or isempty(.text) then
			frm1.txtOrderQty.Value = 0
		Else
			frm1.txtOrderQty.Value = .Text
		End if
		.Col = C_PlanStartDt
		frm1.txtPlanStartDt.Text = .Text
		.Col = C_PlanComptDt
		frm1.txtPlanComptDt.Text = .Text
		.Col = C_Routing
		frm1.txtRouting.value = .Text 
		.Col = C_TrackingNo
		frm1.txtTrackingNo.value = .Text 
	
    End With

End Function


'========================================================================================
' Function Name : uFncDtlQuery
' Function Desc : This function is data query and display
'========================================================================================

Function uFncDtlQuery(Byval lRow) 
    

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

'	Call InitVariables
'	Call FncQuery()

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����Document��ȸ</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOprCd" SIZE=10 MAXLENGTH=3 tag="11xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprCd()"></TD>
								</TR>													
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD><!-- ù��° �� ���� -->
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>ǰ��</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="24" ALT="ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="24"></TD>
							<TD CLASS=TD5 NOWRAP>��������</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/p4215ma1_I793539428_txtOrderQty.js'></script>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>�۾���ȹ����</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/p4215ma1_I932335698_txtPlanStartDt.js'></script>
								&nbsp;~&nbsp;
								<script language =javascript src='./js/p4215ma1_I733530098_txtPlanComptDt.js'></script>
							</TD>								
							<TD CLASS=TD5 NOWRAP>�����</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRouting" SIZE=10 MAXLENGTH=9 tag="24" ALT="�����"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="24" ALT="Tracking No."></TD>
							<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
							<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
						</TR>	
						<TR>
							<TD HEIGHT="60%" colspan=4>
								<script language =javascript src='./js/p4215ma1_I141483849_vspdData.js'></script>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 HEIGHT="40%" NOWRAP>Document</TD>
                         	<TD CLASS=TD656 HEIGHT="40%" valign="middle" colspan=3>
                         		<TEXTAREA  NAME="txtDocument" tag="24xxxU" rows=6 cols=80 wrap="soft" ALT="Document"></TEXTAREA>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOprCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
