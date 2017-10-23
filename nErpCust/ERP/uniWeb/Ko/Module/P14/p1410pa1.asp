
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p1410pa1.asp																*
'*  4. Program Name         : ECN PopUp																	*
'*  5. Program Desc         : Look up ECN No															*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2003/03/06																*
'*  8. Modified date(Last)  :																			*
'*  9. Modifier (First)     : Ryu Sung Won																*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
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

Const BIZ_PGM_QRY_ID = "p1410pb1.asp"			<% '��: �����Ͻ� ���� ASP�� %>

Const C_SHEETMAXROWS = 100

Dim C_EcnNo
Dim C_EcnDesc
Dim C_ReasonCd
Dim C_ReasonNm
Dim C_Status
Dim C_EBomFlg
Dim C_EBomDt
Dim C_MBomFlg
Dim C_MBomDt
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_Remark
	
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
	C_EcnNo			= 1
	C_EcnDesc		= 2
	C_ReasonCd		= 3
	C_ReasonNm		= 4
	C_Status		= 5
	C_EBomFlg		= 6
	C_EBomDt		= 7
	C_MBomFlg		= 8
	C_MBomDt		= 9
	C_ValidFromDt	= 10
	C_ValidToDt		= 11
	C_Remark		= 12
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
	<% '------ Coding part ------ %>
	Self.Returnvalue = Array("")
End Function

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()
	Dim ArgArray						'Arguments�� �Ѱܹ��� Array

	ArgArray  = ArrParent(1)
	txtECNNo.value		= UCase(ArgArray(0))
	txtReasonCd.value	= ArgArray(1)
	'cboStatus.value		= UCase(ArgArray(2))
	'cboEBomFlg.value	= ArgArray(3)
	'cboMBomFlg.value	= ArgArray(4)
	
	iDBSYSDate = "<%=GetSvrDate%>"
	txtValidDt.text = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
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
	
	vspdData.MaxCols = C_Remark + 1
	vspdData.MaxRows = 0

    Call GetSpreadColumnPos("A")
    
	ggoSpread.SSSetEdit		C_EcnNo,		"���躯���ȣ", 18
	ggoSpread.SSSetEdit		C_EcnDesc,		"���躯�泻��", 25
	ggoSpread.SSSetEdit		C_ReasonCd,		"���躯��ٰ�", 12
	ggoSpread.SSSetEdit		C_ReasonNm,		"���躯��ٰŸ�", 20
	ggoSpread.SSSetEdit		C_Status,		"���躯�����", 12
	ggoSpread.SSSetEdit		C_EBomFlg,		"����BOM�ݿ�����", 14, 2
	ggoSpread.SSSetEdit		C_EBomDt,		"����BOM�ݿ���", 14, 2
	ggoSpread.SSSetEdit		C_MBomFlg,		"����BOM�ݿ�����", 14, 2
	ggoSpread.SSSetEdit		C_MBomDt,		"����BOM�ݿ���", 14, 2
	ggoSpread.SSSetEdit		C_ValidFromDt,	"������", 12, 2
	ggoSpread.SSSetEdit		C_ValidToDt,	"������", 12, 2
	ggoSpread.SSSetEdit		C_Remark,		"���", 50
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
    
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
            
            C_EcnNo			= iCurColumnPos(1)
			C_EcnDesc		= iCurColumnPos(2)
			C_ReasonCd		= iCurColumnPos(3)
			C_ReasonNm		= iCurColumnPos(4)
			C_Status		= iCurColumnPos(5)
			C_EBomFlg		= iCurColumnPos(6)
			C_EBomDt		= iCurColumnPos(7)
			C_MBomFlg		= iCurColumnPos(8)
			C_MBomDt		= iCurColumnPos(9)
			C_ValidFromDt	= iCurColumnPos(10)
			C_ValidToDt		= iCurColumnPos(11)
			C_Remark		= iCurColumnPos(12)
			
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
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End If
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
		Redim arrReturn(10)
		
		vspdData.Row = vspdData.ActiveRow

		If vspdData.SelModeSelected = True Then
			vspdData.Col = C_EcnNo
			arrReturn(0) = vspdData.Text
			vspdData.Col = C_EcnDesc
			arrReturn(1) = vspdData.Text
			vspdData.Col = C_ReasonCd
			arrReturn(2) = vspdData.Text
			vspdData.Col = C_ReasonNm
			arrReturn(3) = vspdData.Text
			vspdData.Col = C_Status
			arrReturn(4) = vspdData.Text
			vspdData.Col = C_EBomFlg
			arrReturn(5) = vspdData.Text
			vspdData.Col = C_EBomDt
			arrReturn(6) = vspdData.Text
			vspdData.Col = C_MBomFlg
			arrReturn(7) = vspdData.Text
			vspdData.Col = C_MBomDt
			arrReturn(8) = vspdData.Text
			vspdData.Col = C_ValidFromDt
			arrReturn(9) = vspdData.Text
			vspdData.Col = C_ValidToDt
			arrReturn(10) = vspdData.Text
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

Sub vspdData_KeyPress(keyAscii)
	If keyAscii=13 and vspdData.ActiveRow > 0 Then
 		Call OkClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub	

Sub txtValidDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call FncQuery()
	End If
End Sub
'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************

'=======================================================================================================
'   Event Name : txtValidDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidDt_DblClick(Button)
    If Button = 1 Then
        txtValidDt.Action = 7
        Call SetFocusToDocument("P")
		txtValidDt.Focus
    End If
End Sub


'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================

'------------------------------------------  OpenReasonPopup()  ------------------------------------------
'	Name : OpenReasonPopup()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
  
	'---------------------------------------------
	' Parameter Setting
	'--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "���躯��ٰ��˾�"				' �˾� ��Ī 
	arrParam(1) = "B_MINOR"								' TABLE ��Ī 
	arrParam(2) = UCase(Trim(txtReasonCd.value))		' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1402", "''", "S") & ""
	
	arrParam(5) = "���躯��ٰ�"					' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"							' Field��(0)
    arrField(1) = "MINOR_NM"							' Field��(1)
        
    arrHeader(0) = "���躯��ٰ�"					' Header��(0)
    arrHeader(1) = "���躯��ٰŸ�"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReasonInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtReasonCd.focus
	
End Function


'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================

'------------------------------------------  SetReasonInfo()  ---------------------------------------------
'	Name : SetReasonInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function SetReasonInfo(byval arrRet)
	txtReasonCd.Value	= arrRet(0)
	txtReasonNm.Value	= arrRet(1)	
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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    		
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	
	Call InitVariables											'��: Initializes local global variables
	Call InitSpreadSheet()
	'Call InitComboBox()
	Call InitSetting()
	Call FncQuery()

	txtECNNo.focus
	Set gActiveElement = document.activeElement
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
	
	Dim strcboStatus, strcboEBomFlg, strcboMBomFlg
	
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
		strVal = strVal & "&txtEcnNo="		& Trim(hECNNo.value)
		strVal = strVal & "&txtEcnDesc="	& Trim(hECNDesc.value)
		strVal = strVal & "&txtReasonCd="	& Trim(hReasonCd.value)
		strVal = strVal & "&txtValidDt="	& hValidDt.value
		strVal = strVal & "&cboStatus="		& Trim(hStatus.value)
		strVal = strVal & "&cboEBomFlg="	& Trim(hEBomFlg.value)
		strVal = strVal & "&cboMBomFlg="	& Trim(hMBomFlg.value)
		
		strVal = strVal & "&lgIntFlgMode="	& lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey="	& lgStrPrevKey
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtEcnNo="		& Trim(txtECNNo.value)
		strVal = strVal & "&txtEcnDesc="	& Trim(txtECNDesc.value)
		strVal = strVal & "&txtReasonCd="	& Trim(txtReasonCd.value)
		strVal = strVal & "&txtValidDt="	& txtValidDt.text
		
		If cboStatus1.checked = True then
			strcboStatus = ""
		ElseIf cboStatus2.checked = True then
			strcboStatus = "1"
		Else			
			strcboStatus = "2"
		End IF
		
		If cboEBomFlg1.checked = True then
			strcboEBomFlg = ""
		ElseIf cboEBomFlg2.checked = True then
			strcboEBomFlg = "Y"
		Else			
			strcboEBomFlg = "N"
		End IF
		
		If cboMBomFlg1.checked = True then
			strcboMBomFlg = ""
		ElseIf cboMBomFlg2.checked = True then
			strcboMBomFlg = "Y"
		Else			
			strcboMBomFlg = "N"
		End IF	
			
		strVal = strVal & "&cboStatus=" & strcboStatus
		strVal = strVal & "&cboEBomFlg=" & strcboEBomFlg
		strVal = strVal & "&cboMBomFlg=" & strcboMBomFlg

		strVal = strVal & "&lgIntFlgMode="	& lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey="	& lgStrPrevKey

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
						<TD CLASS=TD5 NOWRAP>���躯���ȣ</TD>
						<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtECNNo" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="���躯���ȣ">&nbsp;<INPUT TYPE=TEXT NAME="txtECNDesc" SIZE=50 MAXLENGTH=100 tag="11XXXX" ALT="���躯�泻��"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>���躯��ٰ�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReasonCd" SIZE=6 MAXLENGTH=2 tag="1XXXU" ALT="���躯��ٰ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReasonPopup" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenReasonPopup()">&nbsp;<INPUT TYPE=TEXT NAME="txtReasonNm" SIZE=20 tag="X4" ALT="���躯��ٰŸ�"></TD>
						<TD CLASS=TD5 NOWRAP>���躯�����</TD>
						<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboStatus" tag="1X" ID="cboStatus1" VALUE="1"><LABEL FOR="cboStatus1">��ü</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboStatus" tag="1X" CHECKED ID="cboStatus2" VALUE="2"><LABEL FOR="cboStatus2">Active</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboStatus" tag="1X" ID="cboStatus3" VALUE="3"><LABEL FOR="cboStatus3">Inactive</LABEL>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>����BOM�ݿ�����</TD>
						<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBomFlg" tag="1X" CHECKED ID="cboEBomFlg1" VALUE=""><LABEL FOR="cboEBomFlg1">��ü</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBomFlg" tag="1X" ID="cboEBomFlg2" VALUE="Y"><LABEL FOR="cboEBomFlg2">��</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBomFlg" tag="1X" ID="cboEBomFlg3" VALUE="N"><LABEL FOR="cboEBomFlg3">�ƴϿ�</LABEL>
						</TD>
						<TD CLASS=TD5 NOWRAP>����BOM�ݿ�����</TD>
						<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBomFlg" tag="1X" CHECKED ID="cboMBomFlg1" VALUE=""><LABEL FOR="cboMBomFlg1">��ü</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBomFlg" tag="1X" ID="cboMBomFlg2" VALUE="Y"><LABEL FOR="cboMBomFlg2">��</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBomFlg" tag="1X" ID="cboMBomFlg3" VALUE="N"><LABEL FOR="cboMBomFlg3">�ƴϿ�</LABEL>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>������</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/p1410pa1_I845325320_txtValidDt.js'></script>
						</TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p1410pa1_vspdData_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hEcnNo" tag="24"><INPUT TYPE=HIDDEN NAME="hECNDesc" tag="24"><INPUT TYPE=HIDDEN NAME="hReasonCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hValidDt" tag="24"><INPUT TYPE=HIDDEN NAME="hStatus" tag="24"><INPUT TYPE=HIDDEN NAME="hEBomFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hMBomFlg" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
