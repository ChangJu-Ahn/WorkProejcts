
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p4211pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Reservation Reference ASP											*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2005/12/29																*
'*  8. Modified date(Last)  : 2005/12/29																*
'*  9. Modifier (First)     : Chen, Jae Hyun															*
'* 10. Modifier (Last)      : Chen, Jae Hyun															*
'* 11. Comment              :																			*
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)  
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin		                *
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->

<!--********************************************  1.1 Inc ����  *****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--============================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--============================================  1.1.2 ���� Include  ===================================
'=====================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p4211pb1.asp"			'��: �����Ͻ� ���� ASP�� 
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
Dim C_prodt_order_no
Dim C_opr_no
Dim C_job_nm
Dim C_sl_cd 
Dim C_sl_nm
Dim C_tracking_no
Dim C_req_dt
Dim C_req_qty
Dim C_basic_unit
Dim C_issued_qty
Dim C_consumed_qty
Dim C_req_no
Dim C_item_cd
Dim C_item_nm
Dim C_spec
Dim C_plan_start_dt
Dim C_plan_compt_dt
Dim C_order_status
Dim C_prodt_order_qty
Dim C_prod_qty_in_order_unit
Dim C_good_qty_in_order_unit
Dim C_rcpt_qty_in_order_unit
Dim C_prodt_order_unit
Dim C_re_work_flg

	
'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
	
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
Dim arrReturn
Dim lgPlantCD
Dim strFromStatus
Dim strToStatus
Dim strThirdStatus
Dim IsOpenPop
Dim arrParent
Dim IsFormLoaded

Dim lgStrPrevKey1, lgStrPrevKey2, lgStrPrevKey3
	
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable�� ����  ============================
'========================================================================================================
'----------------  ���� Global ������ ����  -------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_prodt_order_no			= 1
	C_opr_no					= 2
	C_job_nm					= 3
	C_sl_cd						= 4
	C_sl_nm						= 5
	C_tracking_no				= 6
	C_req_dt					= 7
	C_req_qty					= 8
	C_basic_unit				= 9
	C_issued_qty				= 10
	C_consumed_qty				= 11
	C_req_no					= 12
	C_item_cd					= 13
	C_item_nm					= 14
	C_spec						= 15
	C_plan_start_dt				= 16
	C_plan_compt_dt				= 17
	C_order_status				= 18
	C_prodt_order_qty			= 19
	C_prod_qty_in_order_unit	= 20
	C_good_qty_in_order_unit	= 21
	C_rcpt_qty_in_order_unit	= 22
	C_prodt_order_unit			= 23
	C_re_work_flg				= 24
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	vspdData.MaxRows = 0
	lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
	lgStrPrevKey1 = ""	: lgStrPrevKey2 = "" : lgStrPrevKey3 = ""		'initializes Previous Key		
    lgIntFlgMode = PopupParent.OPMD_CMODE								'Indicates that current mode is Create mode	

	Self.Returnvalue = Array("")
End Function

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()

	Dim ArgArray						<%'Arguments�� �Ѱܹ��� Array%>
	Dim LocSvrDate
	
	ArgArray  = ArrParent(1)

	lgPlantCD = ArgArray(0)
	txtFromDt.Text = ArgArray(1)
	txtToDt.Text = ArgArray(2)
	strFromStatus = ArgArray(3)
	strToStatus = ArgArray(4)
	If (ArgArray(3) = ArgArray(4)) and (ArgArray(3) <> "" and ArgArray(4) <> "") Then
		cboOrderStatus.value = ArgArray(3)
	End If
	
	If Trim(strToStatus) <> "" Then

        If Len(Trim(strToStatus)) > 2 Then

            strToStatus = Mid(Trim(ArgArray(4)),1,2)

            strThirdStatus = Mid(Trim(ArgArray(4)),3,2)

        Else
            strThirdStatus = Trim(strToStatus)
        End If

    End If
	
	txtProdOrderNo.value = ArgArray(5)
	txtTrackingNo.value = ArgArray(6)
	txtItemCd.value = ArgArray(7)
	'txtChildItemCd.value = "2029H20SP-0001"
	'txtWcCd.value = "W120"	

	txtChildItemCd.value = ArgArray(8)
	txtWcCd.value = ArgArray(9)
	
	'Setting Date
	LocSvrDate = "<%=GetSvrDate%>"

	txtFromDt.text = UniConvDateAToB(UNIDateAdd ("M", -1, LocSvrDate, PopupParent.gServerDateFormat), PopupParent.gServerDateFormat, PopupParent.gDateFormat)
	txtToDt.text   = UniConvDateAToB(UNIDateAdd ("M", 1, LocSvrDate, PopupParent.gServerDateFormat), PopupParent.gServerDateFormat, PopupParent.gDateFormat)
	
End Function

'==========================================   2.1.3 InitComboBox()  =====================================
'=	Name : InitComboBox()																				=
'=	Description : ComboBox�� Value�� Setting�Ѵ�.														=
'========================================================================================================
Sub InitComboBox()

	Dim iCodeArr 
    Dim iNameArr

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call SetCombo2(cboOrderStatus, iCodeArr, iNameArr, Chr(11))

	cboOrderStatus.value = "" 	
End Sub

'==========================================  2.1.4 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display in Spread(s)
'========================================================================================================= 
Sub InitSpreadComboBox()
	On Error Resume Next
End Sub
'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow, ByVal iPos)
	On Error Resume Next
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================%>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE","PA") %>
	<% Call loadBNumericFormatA("Q", "P", "NOCOOKIE","PA") %>
End Sub
	
'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20051127",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	
	vspdData.MaxCols = C_re_work_flg + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit		C_prodt_order_no,	"������ȣ", 18
	ggoSpread.SSSetEdit		C_opr_no,			"����", 8
	ggoSpread.SSSetEdit		C_job_nm,			"�۾���", 10
	ggoSpread.SSSetEdit		C_sl_cd,			"â��", 10
	ggoSpread.SSSetEdit		C_sl_nm,			"â���", 20
	ggoSpread.SSSetEdit		C_tracking_no,		"Tracking No.", 25
	ggoSpread.SSSetDate		C_req_dt,			"�ʿ���", 10, 2, PopupParent.gDateFormat
	ggoSpread.SSSetFloat	C_req_qty,			"�ʿ����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_basic_unit,		"���ش���", 8
	ggoSpread.SSSetFloat	C_issued_qty,		"������", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_consumed_qty,		"�Һ����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_req_no,			"��û��ȣ", 18
	
	ggoSpread.SSSetEdit		C_item_cd,			"��ǰ��", 18
	ggoSpread.SSSetEdit		C_item_nm,			"��ǰ���", 25
	ggoSpread.SSSetEdit		C_spec,				"�԰�", 25
	ggoSpread.SSSetDate		C_plan_start_dt,	"����������", 10, 2, PopupParent.gDateFormat
	ggoSpread.SSSetDate		C_plan_compt_dt,	"�ϷΌ����", 10, 2, PopupParent.gDateFormat
	ggoSpread.SSSetEdit		C_order_status,		"���û���", 10

	ggoSpread.SSSetFloat	C_prodt_order_qty,"��������",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_prod_qty_in_order_unit, "��������", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_good_qty_in_order_unit, "��ǰ����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"		
	ggoSpread.SSSetFloat	C_rcpt_qty_in_order_unit, "�԰����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_prodt_order_unit,		"��������", 8
	ggoSpread.SSSetEdit		C_re_work_flg,			"���۾�", 8

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols, True)
	
	ggoSpread.SSSetSplit2(1)
	vspdData.ReDraw = True
	Call SetSpreadLock()
End Sub
	
'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_prodt_order_no			= iCurColumnPos(1)
			C_opr_no					= iCurColumnPos(2)
			C_job_nm					= iCurColumnPos(3)
			C_sl_cd						= iCurColumnPos(4)
			C_sl_nm						= iCurColumnPos(5)
			C_tracking_no				= iCurColumnPos(6)
			C_req_dt					= iCurColumnPos(7)
			C_req_qty					= iCurColumnPos(8)
			C_basic_unit				= iCurColumnPos(9)
			C_issued_qty				= iCurColumnPos(10)
			C_consumed_qty				= iCurColumnPos(11)
			C_req_no					= iCurColumnPos(12)
			C_item_cd					= iCurColumnPos(13)
			C_item_nm					= iCurColumnPos(14)
			C_spec						= iCurColumnPos(15)
			C_plan_start_dt				= iCurColumnPos(16)
			C_plan_compt_dt				= iCurColumnPos(17)
			C_order_status				= iCurColumnPos(18)
			C_prodt_order_qty			= iCurColumnPos(19)
			C_prod_qty_in_order_unit	= iCurColumnPos(20)
			C_good_qty_in_order_unit	= iCurColumnPos(21)
			C_rcpt_qty_in_order_unit	= iCurColumnPos(22)
			C_prodt_order_unit			= iCurColumnPos(23)
			C_re_work_flg				= iCurColumnPos(24)
    End Select    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call initData(1,1)
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If CheckRunningBizProcess = True Then Exit Sub
    If OldLeft <> NewLeft Then Exit Sub
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then Exit Sub
		End If
    End if    
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim strIssueQty, strConsumeQty

	If vspdData.MaxRows > 0 Then
	
		Redim arrReturn(4)
		
		vspdData.Row = vspdData.ActiveRow

		If vspdData.SelModeSelected = True Then
			vspdData.Col = C_prodt_order_no
			arrReturn(0) = vspdData.Text
			vspdData.Col = C_req_no
			arrReturn(1) = vspdData.Text
			vspdData.Col = C_tracking_no
			arrReturn(2) = vspdData.Text
			vspdData.Col = C_issued_qty
			strIssueQty = UniCDbl(vspdData.Text)
			vspdData.Col = C_consumed_qty
			strConsumeQty = UniCDbl(vspdData.Text)
			arrReturn(3) = UniFormatNumber(strIssueQty - strConsumeQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			vspdData.Col = C_basic_unit
			arrReturn(4) = vspdData.Text
		End If

		Self.Returnvalue = arrReturn
	End If		
	Self.Close()
End Function
	
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub txtFromDt_KeyDown(keycode, shift)
	If keycode=27 Then
 		Call Self.Close()
		Exit Sub
	ElseIf Keycode = 13 Then
		Call FncQuery()
	End If
End Sub	

Sub txtToDt_KeyDown(keycode, shift)
	If keycode=27 Then
 		Call Self.Close()
		Exit Sub
	ElseIf Keycode = 13 Then
		Call FncQuery()
	End If
End Sub	

Sub vspdData_KeyPress(keyAscii)
	If keyAscii=13 and vspdData.ActiveRow > 0 Then
 		Call OkClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub	


'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        txtFromDt.Action = 7
        Call SetFocusToDocument("P")
		txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        txtToDt.Action = 7
        Call SetFocusToDocument("P")
		txtToDt.Focus
    End If
End Sub

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
'------------------------------------------  OpenItemInfo()  --------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strItemCd)

	Dim arrRet
	Dim arrParam(5), arrField(16)
	Dim iCalledAspName, IntRetCD

	IsOpenPop = True
	
	arrParam(0) = Trim(lgPlantCD)				' Plant Code
	arrParam(1) = Trim(strItemCd)				' Item Code
	arrParam(2) = "15!MP"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 	'ITEM_CD				' Field��(0)
	arrField(1) = 2 	'ITEM_NM				' Field��(1)
	arrField(2) = 26 	'UNIT_OF_ORDER_MFG
	arrField(3) = 4		'BASIC_UNIT
	arrField(4) = 28	'ORDER_LT_MFG
	arrField(5) = 33	'MIN_MRP_QTY
	arrField(6) = 34	'MAX_MRP_QTY
	arrField(7) = 35	'ROND_QTY
	arrField(8) = 36	'PROD_MGR	-- ?
	arrField(9) = 15	'MAJOR_SL_CD
	arrField(10) = 13	'PHANTOM_FLG
	arrField(11) = 25	'TRACKING_FLG
	arrField(12) = 17	'VALID_FLG
	arrField(13) = 18	'VALID_FROM_DT
	arrField(14) = 19	'VALID_TO_DT
	arrField(15) = 49	'INSPEC_MGR
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtItemCd.Focus
	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(lgPlantCD)
	arrParam(1) = Trim(strCode)
	arrParam(2) = ""
	arrParam(3) = txtFromDt.Text
	arrParam(4) = txtToDt.Text	
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	txtTrackingNo.Focus
	
End Function

'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(txtItemGroupCd.className) = UCase(PopUpParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(txtItemGroupCd.Value))
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
	txtItemGroupCd.focus
 
End Function

'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
'------------------------------------------  SetItemInfo()  ---------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet)
	txtItemCd.Value		= arrRet(0)
	txtItemNm.Value		= arrRet(1)
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	txtTrackingNo.Value = arrRet(0)
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	txtItemGroupCd.Value    = arrRet(0)  
	txtItemGroupNm.Value    = arrRet(1)  
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()

	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'��: Load table , B_numeric_format		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    		
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	Call InitVariables											'��: Initializes local global variables
	Call InitSpreadSheet()
	Call InitComboBox()
	Call InitSetting()
	txtProdOrderNo.focus
	Set gActiveElement = document.activeElement 
	Call FncQuery()
	
	IsFormLoaded = true											'After Loading the Form, the OrderStatus Variables can be Changed.
End Sub
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery()
    FncQuery = False
    Call InitVariables
	If DbQuery = False Then	
		Exit Function
	End If
	FncQuery = False
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"					'SpreadSheet ������ vspdData�ϰ�� 
	Set gActiveSpdSheet = vspdData
	Call SetPopupMenuItemInf("0000111111")
	
    If vspdData.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If    
End Sub

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then Exit Function
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	On Error Resume Next
    Err.Clear                                                               <%'��: Protect system from crashing%>
	    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkfield(Document, "1") Then									'��: This function check indispensable field
	   Exit Function
	End If
	    
	If ValidDateCheck(txtFromDt, txtToDt) = False Then Exit Function
	    
    DbQuery = False                                                         <%'��: Processing is NG%>
	    
    Call LayerShowHide(1)
	    
    Dim strVal
		
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtFromDt=" & Trim(hProdFromDt.value)
		strVal = strVal & "&txtToDt=" & Trim(hProdToDt.value)
		strVal = strVal & "&txtFromStstus=" & strFromStatus
		strVal = strVal & "&txtToStstus=" & strToStatus
		strVal = strVal & "&txtThirdStstus=" & strThirdStatus
		strVal = strVal & "&txtProdOrderNo=" & Trim(hProdOrderNo.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(hItemGroupCd.value)
		strVal = strVal & "&txtChildItemCd=" & Trim(txtChildItemCd.value)
		strVal = strVal & "&txtWcCd=" & Trim(txtWcCd.value)
	Else
		If Trim(cboOrderStatus.value) <> "" Then
			strFromStatus	= Trim(cboOrderStatus.value)
			strToStatus		= Trim(cboOrderStatus.value)
			strThirdStatus  = Trim(cboOrderStatus.value)
		ElseIf Trim(cboOrderStatus.value) = "" And IsFormLoaded = True Then		'After Loading the Form, the OrderStatus Variables can be Changed.
			strFromStatus	= ""
			strToStatus		= ""
			strThirdStatus  = ""
		End If

		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtFromDt=" & Trim(txtFromDt.text)
		strVal = strVal & "&txtToDt=" & Trim(txtToDt.text)
		strVal = strVal & "&txtFromStstus=" & strFromStatus
		strVal = strVal & "&txtToStstus=" & strToStatus
		strVal = strVal & "&txtThirdStstus=" & strThirdStatus
		strVal = strVal & "&txtProdOrderNo=" & Trim(txtProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(txtItemGroupCd.value)
		strVal = strVal & "&txtChildItemCd=" & Trim(txtChildItemCd.value)
		strVal = strVal & "&txtWcCd=" & Trim(txtWcCd.value)
	End If    

    Call RunMyBizASP(MyBizASP, strVal)					'��: �����Ͻ� ASP �� ���� 
		
    DbQuery = True                                      '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(LngMaxRows)															'��: ��ȸ ������ ������� 
	If lgIntFlgMode = PopupParent.OPMD_CMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = PopupParent.OPMD_UMODE	
    vspddata.Focus												'��: Indicates that current mode is Update mode
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5 NOWRAP>��ǰ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChildItemCd" SIZE=18 MAXLENGTH=18 tag="14xxxU" ALT="��ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemNm" SIZE=20 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>�۾���</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=18 MAXLENGTH=10 tag="14xxxU" ALT="�۾���">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="14"></TD>
					</TR>
					
					<TR>
						<TD CLASS=TD5 NOWRAP>�������� ��ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="�������� ��ȣ"></TD>
						<TD CLASS=TD5>�ʿ���</TD>
						<TD CLASS=TD6>
							<script language =javascript src='./js/p4211pa1_I869256311_txtFromDt.js'></script>
							&nbsp;~&nbsp;
							<script language =javascript src='./js/p4211pa1_I748909616_txtToDt.js'></script>
						</TD>
						
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ǰ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=18 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>��ǰ��׷�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=20 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo txtTrackingNo.value"></TD>
						<TD CLASS=TD5 NOWRAP>���û���</TD>
						<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderStatus" ALT="���û���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p4211pa1_vspdData_vspdData.js'></script>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hToStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
