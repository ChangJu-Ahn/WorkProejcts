
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : ����BOM����																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p1714pa2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Production Order Reference ASP											*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2005-02-18 																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : yjw																		*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. History
'*                          :													                        *
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
Const BIZ_PGM_QRY_ID = "p1714pb1.asp"			'��: �����Ͻ� ���� ASP�� 
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

Dim C_Level
Dim C_ChildItemSeq
Dim C_ChildItemCd
Dim C_ChildItemNm
Dim C_ChildItemSpec
Dim C_ItemAcctNm
Dim C_ProcurTypeNm
Dim C_ChildItemQty
Dim C_ChildItemUnit
Dim C_PrntItemQty
Dim C_PrntItemUnit
Dim C_SafetyLt
Dim C_LossRate
Dim C_SupplyTypeNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_EcnNo
Dim C_EcnDesc
Dim C_ReasonNm
Dim C_DrawingPath

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

	C_Level			= 1
	C_ChildItemSeq	= 2
	C_ChildItemCd	= 3
	C_ChildItemNm	= 4
	C_ChildItemSpec	= 5
	C_ItemAcctNm	= 6
	C_ProcurTypeNm	= 7
	C_ChildItemQty	= 8
	C_ChildItemUnit	= 9
	C_PrntItemQty	= 10
	C_PrntItemUnit	= 11
	C_SafetyLt		= 12
	C_LossRate		= 13
	C_SupplyTypeNm	= 14
	C_ValidFromDt	= 15
	C_ValidToDt		= 16
	C_EcnNo			= 17
	C_EcnDesc		= 18
	C_ReasonNm		= 19
	C_DrawingPath	= 20

End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	vspdData.MaxRows = 0
	lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
	lgStrPrevKey = ""										'initializes Previous Key
    lgIntFlgMode = PopupParent.OPMD_CMODE								'Indicates that current mode is Create mode

	Self.Returnvalue = Array("")
End Function

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()
	Dim ArgArray						<%'Arguments�� �Ѱܹ��� Array%>

	ArgArray  = ArrParent(1)
	hBasePlantCd.Value  = ArgArray(0)
	txtReqTransNo.Value = ArgArray(1)

End Function

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
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False

	vspdData.MaxCols = C_DrawingPath + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit		C_Level,			"����", 8
	ggoSpread.SSSetEdit		C_ChildItemSeq,		"����", 8
	ggoSpread.SSSetEdit		C_ChildItemCd,		"��ǰ��", 14
	ggoSpread.SSSetEdit		C_ChildItemNm,		"��ǰ���", 18
	ggoSpread.SSSetEdit		C_ChildItemSpec,	"�԰�", 18
	ggoSpread.SSSetEdit		C_ItemAcctNm,		"ǰ�����", 12, 2
	ggoSpread.SSSetEdit		C_ProcurTypeNm,		"���ޱ���", 12, 2
	ggoSpread.SSSetEdit		C_ChildItemQty,		"��ǰ����ؼ�", 12, 1
	ggoSpread.SSSetEdit		C_ChildItemUnit,	"����", 8, 2
	ggoSpread.SSSetEdit		C_PrntItemQty,		"��ǰ����ؼ�", 12, 1
	ggoSpread.SSSetEdit		C_PrntItemUnit,		"����", 8, 2
	ggoSpread.SSSetEdit		C_SafetyLt,			"����L/T", 10, 1
	ggoSpread.SSSetEdit		C_LossRate,			"Loss��", 10, 1
	ggoSpread.SSSetEdit		C_SupplyTypeNm,		"�����󱸺�", 12, 2
	ggoSpread.SSSetEdit		C_ValidFromDt,		"������", 12, 2
	ggoSpread.SSSetEdit		C_ValidToDt,		"������", 12, 2
	ggoSpread.SSSetEdit		C_EcnNo,			"���躯���ȣ", 18
	ggoSpread.SSSetEdit		C_EcnDesc,			"���躯�泻��", 24
	ggoSpread.SSSetEdit		C_ReasonNm,			"���躯��ٰ�", 24
	ggoSpread.SSSetEdit		C_DrawingPath,		"������", 28

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols, True)

'	ggoSpread.SSSetSplit2(1)
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


			C_Level				= iCurColumnPos(1)
			C_ChildItemSeq		= iCurColumnPos(2)
			C_ChildItemCd		= iCurColumnPos(3)
			C_ChildItemNm		= iCurColumnPos(4)
			C_ChildItemSpec		= iCurColumnPos(5)
			C_ItemAcctNm		= iCurColumnPos(6)
			C_ProcurTypeNm		= iCurColumnPos(7)
			C_ChildItemQty		= iCurColumnPos(8)
			C_ChildItemUnit		= iCurColumnPos(9)
			C_PrntItemQty		= iCurColumnPos(10)
			C_PrntItemUnit		= iCurColumnPos(11)
			C_SafetyLt			= iCurColumnPos(12)
			C_LossRate			= iCurColumnPos(13)
			C_SupplyTypeNm		= iCurColumnPos(14)
			C_ValidFromDt		= iCurColumnPos(15)
			C_ValidToDt			= iCurColumnPos(16)
			C_EcnNo				= iCurColumnPos(17)
			C_EcnDesc			= iCurColumnPos(18)
			C_ReasonNm			= iCurColumnPos(19)
			C_DrawingPath		= iCurColumnPos(20)

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
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If CheckRunningBizProcess = True Then Exit Sub
    If OldLeft <> NewLeft Then Exit Sub

    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
	Dim intRowCnt
	Dim intColCnt
	Dim intSelCnt

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
'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================


'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------


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
	Call InitSetting()
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

    DbQuery = False                                                         <%'��: Processing is NG%>

    Call LayerShowHide(1)

    Dim strVal

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtReqTransNo=" & Trim(hReqTransNo.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtReqTransNo=" & Trim(txtReqTransNo.value)
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


'------------------------------------------  OpenReqTransNo()  -------------------------------------------
'	Name : OpenReqTransNo()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function OpenReqTransNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strPlantCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	strPlantCd = Trim(hBasePlantCd.value)

	' �˾� ��Ī 
	arrParam(0) = "�̰��Ƿڹ�ȣ"
	' TABLE ��Ī 
	arrParam(1) = "P_EBOM_TO_PBOM_MASTER A, B_ITEM B, B_PLANT C"
	' Code Condition
	arrParam(2) = Trim(txtReqTransNo.value)
	' Name Cindition
	arrParam(3) = ""
	' Where Condition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD AND A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S")
	' TextBox ��Ī 
	arrParam(5) = "�̰��Ƿڹ�ȣ"

    arrField(0) = "A.REQ_TRANS_NO"				' Field��(0)
    arrField(1) = "A.PLANT_CD"					' Field��(1)
    arrField(2) = "C.PLANT_NM"					' Field��(2)
    arrField(3) = "A.ITEM_CD"					' Field��(3)
    arrField(4) = "B.ITEM_NM"					' Field��(4)
    arrField(5) = "A.STATUS"					' Field��(5)

    arrHeader(0) = "�̰��Ƿڹ�ȣ"			' Header��(0)
    arrHeader(1) = "������"				' Header��(1)
    arrHeader(2) = "�������"				' Header��(2)
    arrHeader(3) = "ǰ��"					' Header��(3)
    arrHeader(4) = "ǰ���"					' Header��(4)
    arrHeader(5) = "�̰�����"				' Header��(5)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReqTransNo(arrRet)
	End If

	Call SetFocusToDocument("M")

	txtReqTransNo.focus

End Function

'------------------------------------------  SetReqTransNo()  --------------------------------------------------
'	Name : SetReqTransNo()
'	Description : SetReqTransNo
'---------------------------------------------------------------------------------------------------------
Function SetReqTransNo(Byval arrRet)
	txtReqTransNo.Value	= arrRet(0)
'	frm1.txtDestPlantCd.Value	= arrRet(1)
'	frm1.txtDestPlantNm.Value	= arrRet(2)
'	frm1.txtItemCd.Value		= arrRet(3)
'	frm1.txtItemNm.Value		= arrRet(4)
'	frm1.hStatus.Value			= arrRet(5)
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
						<TD CLASS=TD5 NOWRAP>�̰��Ƿڹ�ȣ</TD>
						<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtReqTransNo" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="�̰��Ƿڹ�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReqTransNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenReqTransNo()"></TD>
					</TR>

					<TR>
						<TD CLASS=TD5 NOWRAP>������</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDestPlantCd" SIZE=6 MAXLENGTH=4 tag="14" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtDestPlantNm" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>�������</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBasePlantCd" SIZE=6 MAXLENGTH=4 tag="14" ALT="�������">&nbsp;<INPUT TYPE=TEXT NAME="txtBasePlantNm" SIZE=25 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ǰ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=14 MAXLENGTH=18 tag="14" ALT="��ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=24 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>�԰�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSpec" SIZE=30 MAXLENGTH=100 tag="14" ALT="�԰�"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>�̰���û��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransDt" SIZE=14 MAXLENGTH=10 tag="14" ALT="�̰���û��"></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p1714pa1_vspdData_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hBasePlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqTransNo" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
