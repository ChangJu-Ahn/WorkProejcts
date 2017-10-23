
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p4600pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Tracking No. ASP															*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2002/01/02																*
'*  8. Modified date(Last)  : 2002/12/10																*
'*  9. Modifier (First)     : Park, BumSoo																*
'* 10. Modifier (Last)      : Ryu Sung Won																*
'* 11. Comment              :																			*
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin                           *
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--#####################################################################################################
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
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

Const BIZ_PGM_QRY_ID = "p4600pb1.asp"				 '��: �����Ͻ� ���� ASP�� 

Dim C_TrackingNo
Dim C_SoNo
Dim C_SoType
Dim C_SoTypeDesc
Dim C_SoldToParty
Dim C_BpNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_SoDt
Dim C_DlvyDt
Dim C_SoQty
Dim C_BaseUnit
Dim C_PlanQty
Dim C_ProdQty
Dim C_GrQty
Dim C_SalesGrp
Dim C_SalesGrpNm
	
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim arrReturn				<% '--- Return Parameter Group %>
Dim lgNextNo				<% '��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� %>
Dim lgPrevNo				<% ' "" %>
Dim lgPlantCD				<% '--- Plant Code %>
Dim strFromStatus
Dim strToStatus
Dim strThirdStatus
Dim IsOpenPop				<%'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� %>
Dim arrParent
Dim iDBSYSDate
Dim StartDate, EndDate

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

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
	C_TrackingNo	= 1
	C_SoNo			= 2
	C_SoType		= 3
	C_SoTypeDesc	= 4
	C_SoldToParty	= 5
	C_BpNm			= 6
	C_ItemCd		= 7
	C_ItemNm		= 8
	C_Spec			= 9
	C_SoDt			= 10
	C_DlvyDt		= 11
	C_SoQty			= 12
	C_BaseUnit		= 13
	C_PlanQty		= 14
	C_ProdQty		= 15
	C_GrQty			= 16
	C_SalesGrp		= 17
	C_SalesGrpNm	= 18
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	vspdData.MaxRows = 0
	lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
	lgStrPrevKey = ""										'initializes Previous Key		
    lgIntFlgMode = PopupParent.OPMD_CMODE					'Indicates that current mode is Create mode	
	<% '------ Coding part ------ %>
	Self.Returnvalue = Array("")
End Function

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()
	Dim ArgArray						'Arguments�� �Ѱܹ��� Array

	ArgArray			= ArrParent(1)
	lgPlantCD			= UCase(ArgArray(0))
	txtTrackingNo.value = ArgArray(1)
	txtItemCd.value		= UCase(ArgArray(2))
	txtDlvryDtFrom.Text	= ArgArray(3)
	txtDlvryDtTo.Text	= ArgArray(4)
End Function

'==========================================   2.1.3 InitComboBox()  =====================================
'=	Name : InitComboBox()																				=
'=	Description : ComboBox�� Value�� Setting�Ѵ�.														=
'========================================================================================================
Sub InitComboBox()
    
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
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
	
	vspdData.MaxCols = C_SalesGrpNm + 1
	vspdData.MaxRows = 0

    Call GetSpreadColumnPos("A")
    
	ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 25
	ggoSpread.SSSetEdit		C_SoNo, "���ֹ�ȣ", 18
	ggoSpread.SSSetEdit		C_SoType, "��������", 10
	ggoSpread.SSSetEdit		C_SoTypeDesc, "��������", 10
	ggoSpread.SSSetEdit		C_SoldToParty, "�ŷ�ó", 10
	ggoSpread.SSSetEdit		C_BpNm, "�ŷ�ó��", 20
	ggoSpread.SSSetEdit		C_ItemCd, "ǰ��", 18
	ggoSpread.SSSetEdit		C_ItemNm, "ǰ���", 25
	ggoSpread.SSSetEdit		C_Spec, "�԰�", 25
	ggoSpread.SSSetDate		C_SoDt, "������", 10, 2, gDateFormat
	ggoSpread.SSSetDate		C_DlvyDt, "������", 10, 2, gDateFormat
	ggoSpread.SSSetFloat	C_SoQty, "���ּ���",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_PlanQty, "��������",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQty, "��������", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GrQty, "�԰����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
    ggoSpread.SSSetEdit		C_BaseUnit, "����", 8
	ggoSpread.SSSetEdit		C_SalesGrp, "�����׷�", 10
	ggoSpread.SSSetEdit		C_SalesGrpNm, "�����׷��", 20

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SoType,C_SoType, True)
    
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
			C_TrackingNo	= iCurColumnPos(1)
			C_SoNo			= iCurColumnPos(2)
			C_SoType		= iCurColumnPos(3)
			C_SoTypeDesc	= iCurColumnPos(4)
			C_SoldToParty	= iCurColumnPos(5)
			C_BpNm			= iCurColumnPos(6)
			C_ItemCd		= iCurColumnPos(7)
			C_ItemNm		= iCurColumnPos(8)
			C_Spec			= iCurColumnPos(9)
			C_SoDt			= iCurColumnPos(10)
			C_DlvyDt		= iCurColumnPos(11)
			C_SoQty			= iCurColumnPos(12)
			C_BaseUnit		= iCurColumnPos(13)
			C_PlanQty		= iCurColumnPos(14)
			C_ProdQty		= iCurColumnPos(15)
			C_GrQty			= iCurColumnPos(16)
			C_SalesGrp		= iCurColumnPos(17)
			C_SalesGrpNm	= iCurColumnPos(18)
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
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
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

	If vspdData.MaxRows > 0 Then
			
		intSelCnt = 0
		Redim arrReturn(0)
		
		vspdData.Row = vspdData.ActiveRow

		If vspdData.SelModeSelected = True Then
			vspdData.Col = C_TrackingNo
			arrReturn(0) = vspdData.Text
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
	'	Self.Returnvalue = Array("")
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

Sub txtDlvryDtFrom_KeyDown(keycode, shift)
	If keycode=27 Then
 		Call Self.Close()
		Exit Sub
	ElseIf Keycode = 13 Then
		Call FncQuery()
	End If
End Sub	

Sub txtDlvryDtTo_KeyDown(keycode, shift)
	If keycode=27 Then
 		Call Self.Close()
		Exit Sub
	ElseIf Keycode = 13 Then
		Call FncQuery()
	End If
End Sub	

Sub txtSoDtFrom_KeyDown(keycode, shift)
	If keycode=27 Then
 		Call Self.Close()
		Exit Sub
	ElseIf Keycode = 13 Then
		Call FncQuery()
	End If
End Sub	

Sub txtSoDtTo_KeyDown(keycode, shift)
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
Sub txtDlvryDtFrom_DblClick(Button)
    If Button = 1 Then
        txtDlvryDtFrom.Action = 7
        Call SetFocusToDocument("M")
		txtDlvryDtFrom.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDlvryDtTo_DblClick(Button)
    If Button = 1 Then
        txtDlvryDtTo.Action = 7
        Call SetFocusToDocument("M")
		txtDlvryDtTo.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtSoDtFrom_DblClick(Button)
    If Button = 1 Then
        txtSoDtFrom.Action = 7
        Call SetFocusToDocument("M")
		txtSoDtFrom.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtSoDtTo_DblClick(Button)
    If Button = 1 Then
        txtSoDtTo.Action = 7
        Call SetFocusToDocument("M")
		txtSoDtTo.Focus
    End If
End Sub

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================

'------------------------------------------  OpenItemInfo()  ------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strItemCd)

	Dim arrRet
	Dim arrParam(5), arrField(16)
	Dim iCalledAspName, IntRetCD
	
	IsOpenPop = True
	
	arrParam(0) = Trim(lgPlantCD)				' Plant Code
	arrParam(1) = Trim(strItemCd)				' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 	'ITEM_CD				' Field��(0)
	arrField(1) = 2 	'ITEM_NM				' Field��(1)
	arrField(2) = 26 	'UNIT_OF_ORDER_MFG
	arrField(3) = 4		'BASIC_UNIT
	arrField(4) = 28	'ORDER_LT_MFG
	arrField(5) = 33	'MIN_MRP_QTY
	arrField(6) = 34	'MAX_MRP_QTY
	arrField(7) = 35	'ROND_QTY
	arrField(8) = 39	'PROD_MGR	-- ?
	arrField(9) = 15	'MAJOR_SL_CD
	arrField(10) = 13	'PHANTOM_FLG
	arrField(11) = 25	'TRACKING_FLG
	arrField(12) = 17	'VALID_FLG
	arrField(13) = 18	'VALID_FROM_DT
	arrField(14) = 19	'VALID_TO_DT
	arrField(15) = 49	'INSPEC_MGR

	iCalledAspName = AskPRAspName("b1b11pa3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	txtItemCd.focus

End Function

'===========================================================================
' Function Name : OpenSoldToParty
' Function Desc : OpenSoldToParty Reference Popup
'===========================================================================
Function OpenSoldToParty()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_BIZ_PARTNER"								' TABLE ��Ī 
	arrParam(2) = Trim(txtSoldToParty.value)					' Code Condition
	arrParam(3) = ""											' Name Cindition
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "	' Where Condition
	arrParam(5) = "�ŷ�ó"									' TextBox ��Ī 
		
	arrField(0) = "BP_CD"										' Field��(0)
	arrField(1) = "BP_NM"										' Field��(1)
		
	arrHeader(0) = "�ŷ�ó"									' Header��(0)
	arrHeader(1) = "�ŷ�ó��"								' Header��(1)

	arrParam(0) = arrParam(5)									' �˾� ��Ī 
	arrParam(3) = ""											' ��: [Condition Name Delete]
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSoldToParty(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	txtSoldToParty.focus
	
End Function


'===========================================================================
' Function Name : OpenSalesGrp
' Function Desc : OpenSalesGrp Reference Popup
'===========================================================================
Function OpenSalesGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����׷�"					' �˾� ��Ī 
	arrParam(1) = "B_SALES_GRP"						' TABLE ��Ī 
	arrParam(2) = Trim(txtSalesGrp.value)			' Code Condition
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					' Where Condition
	arrParam(5) = "�����׷�"					' TextBox ��Ī 
		
	arrField(0) = "SALES_GRP"						' Field��(0)
	arrField(1) = "SALES_GRP_NM"					' Field��(1)
	    
	arrHeader(0) = "�����׷�"					' Header��(0)
	arrHeader(1) = "�����׷��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSalesGrp(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	txtSalesGrp.focus
	
End Function


Function OpenSoType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"					
	arrParam(1) = "S_SO_TYPE_CONFIG"				
	arrParam(2) = Trim(txtSoType.value)
	arrParam(3) = ""		'Trim(txtSoTypeNm.value)
	arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "��������"

	arrField(0) = "ED10" & PopupParent.gColSep & "SO_TYPE"			
	arrField(1) = "ED20" & PopupParent.gColSep & "SO_TYPE_NM"		
	arrField(2) = "ED9" & PopupParent.gColSep & "EXPORT_FLAG"		
	arrField(3) = "ED9" & PopupParent.gColSep & "RET_ITEM_FLAG"	
	arrField(4) = "ED15" & PopupParent.gColSep & "AUTO_DN_FLAG"	
	    
	arrHeader(0) = "��������"					
	arrHeader(1) = "�������¸�"					
	arrHeader(2) = "���⿩��"					
	arrHeader(3) = "��ǰ����"					
	arrHeader(4) = "�ڵ����ϻ�������"

	arrParam(3) = ""			'��: [Condition Name Delete]
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=570px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSoType(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	txtSoType.focus
	
End Function



'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet)
	txtItemCd.Value		= arrRet(0)
	txtItemNm.Value		= arrRet(1)
End Function

'------------------------------------------  SetSoldToParty()  -----------------------------------------
'	Name : SetSoldToParty()
'	Description : Sold-To-Party Popup���� Return�Ǵ� �� setting
'-------------------------------------------------------------------------------------------------------
Function SetSoldToParty(Byval arrRet)
	txtSoldToParty.value = arrRet(0)
	txtSoldToPartyNm.value = arrRet(1)
End Function

'------------------------------------------  SetSalesGrp()  -----------------------------------------
'	Name : SetSalesGrp()
'	Description : Sales Group Popup���� Return�Ǵ� �� setting
'-------------------------------------------------------------------------------------------------------
Function SetSalesGrp(Byval arrRet)
	txtSalesGrp.value = arrRet(0)
	txtSalesGrpNm.value = arrRet(1)
End Function

'------------------------------------------  SetSoType()  -----------------------------------------
'	Name : SetSoType()
'	Description : Sales Order Type Popup���� Return�Ǵ� �� setting
'-------------------------------------------------------------------------------------------------------
Function SetSoType(Byval arrRet)
	txtSoType.value = arrRet(0)
	txtSoTypeNm.value = arrRet(1)
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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
	iDBSYSDate = "<%=GetSvrDate%>"

	StartDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
	EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++	
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'��: Load table , B_numeric_format		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    		
	Call ggoOper.LockField(Document, "N")						'��: Lock  Suitable  Field 
	Call InitVariables											'��: Initializes local global variables
	Call InitSpreadSheet()
	Call InitComboBox()
	Call InitSetting()
	Call FncQuery()
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
	Call DbQuery()
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
	
    Err.Clear                                                               '��: Protect system from crashing
	    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkfield(Document, "1") Then									'��: This function check indispensable field
	   Exit Function
	End If
	    
	If ValidDateCheck(txtDlvryDtFrom, txtDlvryDtTo) = False Then Exit Function
	If ValidDateCheck(txtSoDtFrom, txtSoDtTo) = False Then Exit Function
	    
    DbQuery = False                                                         '��: Processing is NG
	    
    Call LayerShowHide(1)
	    
    Dim strVal

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtTrackingNo=" & Trim(hTrackingNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(hItemCd.value)
		strVal = strVal & "&txtSoldToParty=" & Trim(hSoldToParty.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(hSalesGrp.value)
		strVal = strVal & "&txtSoDtFrom=" & Trim(hSoDtFrom.value)
		strVal = strVal & "&txtSoDtTo=" & Trim(hSoDtTo.value)
		strVal = strVal & "&txtDlvryDtFrom=" & Trim(hDlvryDtFrom.value)
		strVal = strVal & "&txtDlvryDtTo=" & Trim(hDlvryDtTo.value)
		strVal = strVal & "&txtSoType=" & Trim(hSoType.value)
		strVal = strVal & "&txtrdoflag=" & Trim(hrdoflag.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtTrackingNo=" & Trim(txtTrackingNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)
		strVal = strVal & "&txtSoldToParty=" & Trim(txtSoldToParty.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(txtSalesGrp.value)
		strVal = strVal & "&txtSoDtFrom=" & Trim(txtSoDtFrom.text)
		strVal = strVal & "&txtSoDtTo=" & Trim(txtSoDtTo.text)
		strVal = strVal & "&txtDlvryDtFrom=" & Trim(txtDlvryDtFrom.text)
		strVal = strVal & "&txtDlvryDtTo=" & Trim(txtDlvryDtTo.text)
		strVal = strVal & "&txtSoType=" & Trim(txtSoType.value)
		If rdoCloseFlg1.checked = True Then
			strVal = strVal & "&txtrdoflag=" & "O"
		ElseIf rdoCloseFlg2.checked = True Then
			strVal = strVal & "&txtrdoflag=" & "C"
		Else 
			strVal = strVal & "&txtrdoflag=" & ""
		End If
	End If    

    Call RunMyBizASP(MyBizASP, strVal)
		
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()															'��: ��ȸ ������ ������� 
    
    If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
	
    lgIntFlgMode = PopupParent.OPMD_UMODE	
    vspddata.Focus												'��: Indicates that current mode is Update mode
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
						<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."></TD>
						<TD CLASS=TD5 NOWRAP>ǰ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5>��������</TD>
						<TD CLASS=TD6>
							<script language =javascript src='./js/p4600pa1_I960114854_txtSoDtFrom.js'></script>
							&nbsp;~&nbsp;
							<script language =javascript src='./js/p4600pa1_I525442874_txtSoDtTo.js'></script>
						</TD>
						<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSoldToParty" SIZE=10 MAXLENGTH=9 tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoldToParty" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSoldToParty()">&nbsp;<INPUT TYPE=TEXT NAME="txtSoldToPartyNm" SIZE=15 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5>��������</TD>
						<TD CLASS=TD6>
							<script language =javascript src='./js/p4600pa1_I948300367_txtDlvryDtFrom.js'></script>
							&nbsp;~&nbsp;
							<script language =javascript src='./js/p4600pa1_I916953229_txtDlvryDtTo.js'></script>
						</TD>
						<TD CLASS=TD5 NOWRAP>�����׷�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=9 tag="11XXXU" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSalesGrp()">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=15 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoType" TYPE="Text" MAXLENGTH="4" SIZE=6 tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoType()">&nbsp;<INPUT NAME="txtSoTypeNm" TYPE="Text" SIZE=20 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoCloseFlg" ID="rdoCloseFlg1" CLASS="RADIO" tag="11" Value="N" CHECKED><LABEL FOR="rdoCloseFlg1">Open</LABEL>
						     				 <INPUT TYPE="RADIO" NAME="rdoCloseFlg" ID="rdoCloseFlg2" CLASS="RADIO" tag="11" Value="Y"><LABEL FOR="rdoCloseFlg2">����</LABEL>
						     				 <INPUT TYPE="RADIO" NAME="rdoCloseFlg" ID="rdoCloseFlg3" CLASS="RADIO" tag="11" Value="C"><LABEL FOR="rdoCloseFlg3">��ü</LABEL></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p4600pa1_vspdData_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hSoldToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="hSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hSoDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="hSoDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="hDlvryDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="hDlvryDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="hSoType" tag="24">
<INPUT TYPE=HIDDEN NAME="hrdoflag" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
