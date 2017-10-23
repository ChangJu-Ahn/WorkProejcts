<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1104ma1.asp
'*  4. Program Name         : Entry Shift
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2000/04/12
'*  8. Modified date(Last)  : 2002/12/18
'*  9. Modifier (First)     : Mr  KimGyoungDon
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1104mb1.asp"											'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "p1104mb2.asp"											'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID = "p1104mb3.asp"											'��: �����Ͻ� ���� ASP�� 

Dim C_StartDay
Dim C_StartTime
Dim C_EndDay
Dim C_EndTime
Dim C_OverRunFlg
Dim C_MustComplete
Dim C_hStartDay
Dim C_hEndDay

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_StartDay		= 1
	C_StartTime		= 2
	C_EndDay		= 3
	C_EndTime		= 4
	C_OverRunFlg	= 5
	C_MustComplete	= 6
	C_hStartDay		= 7
	C_hEndDay		= 8
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKeyIndex = 0                           'initializes Previous Key
    lgStrPrevKeyIndex1 = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                                       '��: initializes sort direction
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	 frm1.txtValidFromDt.text  = StartDate
	 frm1.txtValidToDt.text	= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_hEndDay + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCombo 	C_StartDay, 		"���ۿ���", 19
		ggoSpread.SSSetTime		C_StartTime,		"���۽ð�", 19, 2, 1, 1
		ggoSpread.SSSetCombo	C_EndDay,			"�������", 19
		ggoSpread.SSSetTime		C_EndTime,			"����ð�", 19, 2, 1, 1
		ggoSpread.SSSetCombo 	C_OverRunFlg, 		"�ܾ����ɿ���", 20
		ggoSpread.SSSetCombo 	C_MustComplete, 	"Shift Break ��뿩��", 20
		ggoSpread.SSSetCombo 	C_hStartDay, 		"���ۿ���", 15
		ggoSpread.SSSetCombo	C_hEndDay,			"�������", 15
	
		Call ggoSpread.SSSetColHidden(C_hStartDay, C_hEndDay, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(1)										'frozen ����߰� 
	
		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1

		.vspdData.ReDraw = False
    
		ggoSpread.SpreadLock C_StartDay, -1, C_StartDay
		ggoSpread.SpreadLock C_StartTime, -1, C_StartTime

		ggoSpread.SSSetRequired 	C_EndDay,		-1
		ggoSpread.SSSetRequired 	C_EndTime,		-1
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1	
		.vspdData.ReDraw = True
	
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired 	C_StartDay,		pvStartRow, pvEndRow	
		ggoSpread.SSSetRequired 	C_StartTime,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_EndDay,		pvStartRow, pvEndRow	
		ggoSpread.SSSetRequired 	C_EndTime,		pvStartRow, pvEndRow		  
	
		.vspdData.ReDraw = True
    
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_StartDay		= iCurColumnPos(1)
			C_StartTime		= iCurColumnPos(2)
			C_EndDay		= iCurColumnPos(3)
			C_EndTime		= iCurColumnPos(4)
			C_OverRunFlg	= iCurColumnPos(5)
			C_MustComplete	= iCurColumnPos(6)
			C_hStartDay		= iCurColumnPos(7)
			C_hEndDay		= iCurColumnPos(8)
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
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim strCboCd
	
	strCboCd = "" & "Y" & vbTab & "N"
	ggoSpread.SetCombo strCboCd, C_OverRunFlg  
	ggoSpread.SetCombo strCboCd, C_MustComplete
	
	strCboCd = ""
	strCboCd = "1" & vbTab & "2" & vbTab & "3" & vbTab & "4" & vbTab & "5" & vbTab & "6" & vbTab & "7" 
	
	ggoSpread.SetCombo strCboCd, C_hStartDay
	ggoSpread.SetCombo strCboCd, C_hEndDay
		
	strCboCd = ""	
	strCboCd = "�Ͽ���" & vbTab & "������" & vbTab & "ȭ����" & vbTab & "������" & vbTab & "�����" & vbTab & "�ݿ���" & vbTab & "�����" 

	ggoSpread.SetCombo strCboCd, C_StartDay
	ggoSpread.SetCombo strCboCd, C_EndDay
	  
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenShiftCd()  -------------------------------------------------
'	Name : OpenShiftCd()
'	Description : Shift Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenShiftCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Shift�˾�"											' �˾� ��Ī 
	arrParam(1) = "P_SHIFT_HEADER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtShiftCd1.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")	' Where Condition
	arrParam(5) = "Shift"												' TextBox ��Ī 
	 
    arrField(0) = "SHIFT_CD"												' Field��(0)
    arrField(1) = "DESCRIPTION"												' Field��(1)
    
    arrHeader(0) = "Shift"												' Header��(0)
    arrHeader(1) = "Shift��"											' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetShiftCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtShiftCd1.focus
	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetShiftCd()  --------------------------------------------------
'	Name : SetShiftCd()
'	Description : Condition Shift Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetShiftCd(byval arrRet)
	frm1.txtShiftCd1.Value    = arrRet(0)		
	frm1.txtShiftNm1.Value    = arrRet(1)		
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function LookUpItemByPlant(strCode)
	Dim strVal
    
    With frm1
    
    strVal = BIZ_PGM_ITEM_ID & "?txtCode=" & strCode						'��: ���� �� 
    strVal = strVal & "&txtClnrType=" & Trim(.txtClnrType.value)
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
End Function

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.Col = C_hStartDay
			intIndex = .value
			.col = C_StartDay
			.value = intindex
			
			.Row = intRow
			.Col = C_hEndDay
			intIndex = .value
			.col = C_EndDay
			.value = intindex
			
		Next	
	End With
	
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
   
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call InitComboBox
    Call SetToolbar("11101101001011")										'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtShiftCd1.focus 
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
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspddata_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("1101110111")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspddata_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspddata_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
'Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
'	Dim intIndex

'	With frm1.vspdData
	
'		.Row = Row
    
'		Select Case Col
'			Case  C_StartDay
'				.Col = Col
'				intIndex = .Value
'				.Col = C_hStartDay
'				.Value = intIndex
'			Case  C_EndDay
'				.Col = Col
'				intIndex = .Value
'				.Col = C_hEndDay
'				.Value = intIndex
'		End Select
    
'    End With

'End Sub

'===========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : COPY�� �ϸ� ���� �Ҵ���� �ʴ� ���׷� vspddata_ComboSelChange �Լ� ������� ���.2003-09-08
'===========================================================================================================
Sub vspdData_Change(Col , Row)

	Dim iDx
       
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Select Case Col
		Case  C_StartDay
			iDx = Frm1.vspdData.value
			Frm1.vspdData.Col = C_hStartDay
			Frm1.vspdData.value = iDx
		Case  C_EndDay
			iDx = Frm1.vspdData.value
			Frm1.vspdData.Col = C_hEndDay
			Frm1.vspdData.value = iDx
		Case Else
	End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> 0 Or lgStrPrevKeyIndex1 <> ""  Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False															'��: Processing is NG

    Err.Clear																	'��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtShiftCd1.value = "" Then
		frm1.txtShiftNm1.value = ""
	End If
		
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
    Call SetDefaultVal
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
		Exit Function           
    End If     													'��: Query db data

    FncQuery = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    'On Error Resume Next														'��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    frm1.txtShiftCd1.value = ""
    frm1.txtShiftNm1.value = "" 
    Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables	
    Call SetToolbar("11101101001011")															'��: Initializes local global variables
    frm1.txtShiftCd2.focus 
    Set gActiveElement = document.activeElement 
    
    FncNew = True																'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False															'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    On Error Resume Next														'��: Protect system from crashing

    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then											'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")									'��:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then
		Exit Function
	End If
	
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then   
		Exit Function           
    End If      
    
    FncDelete = True															'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    'On Error Resume Next														'��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")								'��: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If frm1.txtPlantCd.value = "" Then
			Call DisplayMsgBox("970029", "X","����", "X")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement  
			Exit Function
		End If
	End If
    If Not chkField(Document, "2") Then
		Exit Function
	End If
	
	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck  Then                                  '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		LayerShowHide(0)
		Exit Function           
    End If     																	'��: Save db data
    
    FncSave = True																'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
	Dim IntRetCD
	
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    frm1.vspdData.EditMode = True
    ggoSpread.CopyRow

    Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
    
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo															'��: Protect system from crashing
	Call InitData(1)
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim iIntReqRows, iIntCnt

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If

	With frm1
	
		.vspdData.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData
		.vspdData.EditMode = True
    
		.vspdData.ReDraw = False
 
		ggoSpread.InsertRow , iIntReqRows
    
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)

		For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + iIntReqRows - 1
			.vspdData.Row = iIntCnt
			.vspdData.Col = C_OverRunFlg
			.vspdData.Text = "N"
    
			.vspdData.Col = C_MustComplete
			.vspdData.Text = "N"
		Next
    
		.vspdData.ReDraw = True
    
    End With
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
   
    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
	 
    With frm1.vspdData 
    
    .focus
    Set gActiveElement = document.activeElement 
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow
    
    End With
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
   Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)							'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)	                   '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
	
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    LayerShowHide(1) 
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtShiftCd=" & Trim(.txtShiftCd2.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtShiftCd=" & Trim(.txtShiftCd1.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	If LngMaxRow > 0 Then
		Call InitData(LngMaxRow)
	End If	
    
    lgBlnFlgChgValue = False
    
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    Call SetToolbar("11111111001111")										'��: ��ư ���� ���� 
	
	frm1.txtShiftNm2.focus
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim lRow        
	Dim strVal, strDel
	Dim rtnCheck, iStrSDay, iStrSTime, iStrEDay, iStrETime
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt
	Dim iColSep
	
    DbSave = False                                                          '��: Processing is NG
    
    '-----------------------
    'Check Valid Date
    '-----------------------
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function  
    '-----------------------
    'Save
    '-----------------------
	
	LayerShowHide(1) 
		
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = Parent.gColSep
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
										  
    '-----------------------							  
    'Data manipulate area								  
    '-----------------------							  
    For lRow = 1 To .vspdData.MaxRows					  
														  
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag										'��: �ű� 
				
				strVal = ""
				
				strVal = strVal & "C" & iColSep & lRow & iColSep			'��: C=Create

                .vspdData.Col = C_hStartDay	'1
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrSday = Trim(.vspdData.Text)

                .vspdData.Col = C_StartTime	'2
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrSTime = Trim(.vspdData.Text)

                .vspdData.Col = C_hEndDay	'3
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrEDay = Trim(.vspdData.Text)

                .vspdData.Col = C_EndTime	'4
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrETime = Trim(.vspdData.Text)

                .vspdData.Col = C_OverRunFlg	'5
                strVal = strVal & Trim(.vspdData.Text) & iColSep

                .vspdData.Col = C_MustComplete	'6
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

				rtnCheck = ChkValidData(iStrSday, iStrSTime, iStrEDay, iStrETime) 

				If rtnCheck = 1 Then
					   Call DisplayMsgBox("972002", "X", "�������", "���ۿ���")
					   Call SheetFocus(lRow, 3)
					   Exit Function
				ElseIf rtnCheck = 2 Then
					   Call DisplayMsgBox("972002", "X", "����ð�", "���۽ð�")
					   Call SheetFocus(lRow, 4)
					   Exit Function
				ElseIf rtnCheck = -1 Then
					   Call DisplayMsgBox("970029", "X", "���۽ð�", "X")
					   Call SheetFocus(lRow, 3)
					   Exit Function
				ElseIf rtnCheck = -2 Then
					   Call DisplayMsgBox("970029", "X", "����ð�", "X")
					   Call SheetFocus(lRow, 4)
					   Exit Function
				End If		

				ReDim Preserve TmpBufferVal(iValCnt)
				TmpBufferVal(iValCnt) = strVal
				iValCnt = iValCnt + 1

            Case ggoSpread.UpdateFlag
				
				strVal = ""
				
				strVal = strVal & "U" & iColSep & lRow & iColSep			'��: U=Update

                .vspdData.Col = C_hStartDay	'1
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrSday = Trim(.vspdData.Text)
                              
                .vspdData.Col = C_StartTime	'2
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrSTime = Trim(.vspdData.Text)
                              
                .vspdData.Col = C_hEndDay	'3
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrEday = Trim(.vspdData.Text)
                              
                .vspdData.Col = C_EndTime	'4
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                iStrETime = Trim(.vspdData.Text)
                              
                .vspdData.Col = C_OverRunFlg	'5
                strVal = strVal & Trim(.vspdData.Text) & iColSep
               
                .vspdData.Col = C_MustComplete	'6
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                              
				rtnCheck = ChkValidData(iStrSday, iStrSTime, iStrEDay, iStrETime) 
				
				If rtnCheck = 1 Then
					   Call DisplayMsgBox("972002", "X", "�������", "���ۿ���")
					   Call SheetFocus(lRow, 3)
					   Exit Function
				ElseIf rtnCheck = 2 Then
					   Call DisplayMsgBox("972002", "X", "����ð�", "���۽ð�")
					   Call SheetFocus(lRow, 4)
					   Exit Function
				ElseIf rtnCheck = -1 Then
					   Call DisplayMsgBox("970029", "X", "���۽ð�", "X")
					   Call SheetFocus(lRow, 3)
					   Exit Function
				ElseIf rtnCheck = -2 Then
					   Call DisplayMsgBox("970029", "X", "����ð�", "X")
					   Call SheetFocus(lRow, 4)
					   Exit Function
				End If	
				
				ReDim Preserve TmpBufferVal(iValCnt)
				TmpBufferVal(iValCnt) = strVal
				iValCnt = iValCnt + 1					   
                
            Case ggoSpread.DeleteFlag										'��: ���� 
				
				strDel = ""
				
				strDel = strDel & "D" & iColSep & lRow & iColSep

                .vspdData.Col = C_hStartDay	'1
                strDel = strDel & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_StartTime	'1
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBufferDel(iDelCnt)
				TmpBufferDel(iDelCnt) = strDel
				iDelCnt = iDelCnt + 1
                
        End Select
                
    Next
	
	iTotalStrVal = Join(TmpBufferVal, "")
	iTotalStrDel = Join(TmpBufferDel, "")
	
	.txtSpread.value = iTotalStrDel & iTotalStrVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()															'��: ���� ������ ���� ���� 
   
	frm1.txtShiftCd1.value = frm1.txtShiftCd2.value 
    
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
	
    Call MainQuery()
   
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Dim strVal
	
	DbDelete = False														'��: Processing is NG
	
	LayerShowHide(1) 
		
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)				'��: ���� ���� ����Ÿ 
    strVal = strVal & "&txtShiftCd=" & Trim(frm1.txtShiftCd2.value)				'��: ���� ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    DbDelete = True                                                         '��: Processing is NG 
	
End Function

Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
End Function

'==============================================================================
' Function : ChkValidData
' Description : Start Day�� End Day Check
'==============================================================================
Function ChkValidData(SDay, STime, EDay, ETime)
	ChkValidData = 0

	If CInt(SDay) > CInt(EDay) Then
		ChkValidData = 1
		Exit Function
	End If

	If Len(Trim(STime)) <> 8 and Len(Trim(STime)) <> 0 Then
		ChkValidData = -1
		Exit Function
	End IF

	If Len(Trim(ETime)) <> 8 and Len(Trim(ETime)) <> 0 Then
		ChkValidData = -2
		Exit Function
	End IF

	If CInt(SDay) = CInt(EDay) Then
		If STime > ETime Then
			ChkValidData = 2
			Exit Function
		End If	
	End If
End Function

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Shift���</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Shift</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShiftCd1" SIZE=5 MAXLENGTH=2 tag="12XXXU" ALT="Shift"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenShiftCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtShiftNm1" SIZE=30 MAXLENGTH=40 tag="14" ALT="Shift ��"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>Shift</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShiftCd2" SIZE=5 MAXLENGTH=2 tag="23XXXU" ALT="Shift">&nbsp;<INPUT TYPE=TEXT NAME="txtShiftNm2" SIZE=30 MAXLENGTH=40 tag="21" ALT="Shift ��"></TD>
								<TD CLASS="TD5" NOWRAP>��ȿ�Ⱓ</TD>
								<TD CLASS="TD6">
									<script language =javascript src='./js/p1104ma1_I696465939_txtValidFromDt.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/p1104ma1_I802009737_txtValidToDt.js'></script>											
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" COLSPAN = 4>
									<script language =javascript src='./js/p1104ma1_I536739426_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hShiftCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
