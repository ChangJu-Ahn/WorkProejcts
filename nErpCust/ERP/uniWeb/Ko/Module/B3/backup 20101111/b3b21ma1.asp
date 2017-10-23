<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b21ma1.asp
'*  4. Program Name         : Characteristic
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/02/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "b3b21mb1.asp"					'☆: Detail Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "b3b21mb2.asp"
Const BIZ_PGM_DEL_ID = "b3b21mb3.asp"

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgStrPrevKey2
Dim IsOpenPop	
Dim lgRdoOldVal
Dim lgLastOpNo, lgLastOpRowNo

Dim BaseDate
Dim StartDate

Dim C_CharValueCd
Dim C_CharValueNm

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_CharValueCd	= 1
	C_CharValueNm	= 2
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                                       '⊙: initializes sort direction
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	
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

		.MaxCols = C_CharValueNm + 1
		.MaxRows = 0
	
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_CharValueCd, "사양값", 20,,,16,2
		ggoSpread.SSSetEdit		C_CharValueNm, "사양값명", 96,,,30,1
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True
	End With
    
	ggoSpread.SSSetSplit2(1)										'frozen 기능추가 
    
	Call SetSpreadLock()
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    frm1.vspdData.ReDraw = False
	ggoSpread.SpreadLock	C_CharValueCd,	-1, C_CharValueCd
	ggoSpread.SSSetRequired C_CharValueNm,	-1
	ggoSpread.SSSetProtected frm1.vspdData.MaxCols, -1, -1
	frm1.vspdData.ReDraw = True
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetRequired 	C_CharValueCd, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_CharValueNm, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected 	frm1.vspdData.MaxCols, pvStartRow, pvEndRow
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
			C_CharValueCd	= iCurColumnPos(1)
			C_CharValueNm	= iCurColumnPos(2)
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
	Dim iIntCnt
	
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()

    frm1.vspdData.redraw = False
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()

	frm1.vspdData.Redraw = True
End Sub

'==========================================  InitData()  =================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================= 
Sub InitData(ByVal lngStartRow)

End Sub

'==========================================  OpenCharCd()  ===============================================
'	Name : OpenCharCd()
'	Description : Open Popup
'========================================================================================================= 
Function OpenCharCd()

	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtCharCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = UCase(Trim(frm1.txtCharCd.value))		' Characteristic Code
	arrParam(1) = ""									' Characteristic Name
	arrParam(2) = ""									' ----------
	arrParam(3) = ""									' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 									' Field명(0) : "Characteristic_CD"
    arrField(1) = 2 									' Field명(1) : "Characteristic_NM"
    
	iCalledAspName = AskPRAspName("B3B30PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B30PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=600px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetCharCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCharCd.focus
	
End Function

'==========================================  OpenCharValueCd()  ==========================================
'	Name : OpenCharValueCd()
'	Description : Open Popup
'========================================================================================================= 
Function OpenCharValueCd()

	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtCharValueCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtCharCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "사양항목", "X")
		frm1.txtCharCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtCharCd.value))			' Char Code
	arrParam(1) = UCase(Trim(frm1.txtCharValueCd.value))	' CharValue Name
	arrParam(2) = ""										' ----------
	arrParam(3) = ""										' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 										' Field명(0) : "Characteristic_CD"
    arrField(1) = 2 										' Field명(1) : "Characteristic_NM"

	iCalledAspName = AskPRAspName("B3B32PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B32PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=500px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetCharValueCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtCharValueCd.focus
	
End Function

'==========================================  SetCharCd()  ================================================
'	Name : SetCharCd()
'	Description : Set Popup Values
'========================================================================================================= 
Function SetCharCd(byval arrRet)
	frm1.txtCharCd.Value	= arrRet(0)	
	frm1.txtCharNm.Value   = arrRet(1)
	
	frm1.txtCharCd.focus
	Set gActiveElement = document.activeElement
End Function

'==========================================  SetCharValueCd()  ===========================================
'	Name : SetCharValueCd()
'	Description : Set Popup Values
'========================================================================================================= 
Function SetCharValueCd(byval arrRet)
	frm1.txtCharValueCd.Value	= arrRet(0)	
	frm1.txtCharValueNm.Value   = arrRet(1)
	
	frm1.txtCharValueCd.focus
	Set gActiveElement = document.activeElement
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field

	'----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet															'⊙: Setup the Spread sheet
	Call SetDefaultVal
	Call InitVariables																'⊙: Initializes local global variables
	Call SetToolbar("11101101001011")
	
	frm1.txtCharCd.focus()
	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub txtCharValueDigit_Change()
	lgBlnFlgChgValue = True
End Sub
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	
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
    Dim IntRetCD

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001110111")
	Else 	
		Call SetPopupMenuItemInf("1101111111") 
	End If		

	'If lgIntFlgMode = parent.OPMD_CMODE Then
	'	Call SetPopupMenuItemInf("0000110111")
	'Else 	
	'	If frm1.vspdData.MaxRows = 0 Then 
	'		Call SetPopupMenuItemInf("0000110111")
	'	Else
	'		Call SetPopupMenuItemInf("0001111111") 
	'	End if			
	'End If	
		
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		ggoSpread.Source = frm1.vspdData
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspddata_KeyPress(index , KeyAscii)
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData
		If Row >= NewRow Then
		    Exit Sub
		End If
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
    if frm1.vspdData.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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
    
    FncQuery = False                                                        '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtCharCd.value = "" Then
		frm1.txtCharNm.value = ""
	End If
		
	If frm1.txtCharValueCd.value = "" Then
		frm1.txtCharValueNm.value = ""
	End If
	
    If frm1.txtCharCd1.value = "" Then
		frm1.txtCharNm1.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     									'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    Dim slPlantCd
    Dim slPlantNm
    
    FncNew = False                                                          '⊙: Processing is NG
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    Call SetToolbar("11101101001011")
    
    frm1.txtCharCd1.focus 
    Set gActiveElement = document.activeElement 
    
    FncNew = True                                                           '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '⊙: "Will you destory previous data"	
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
    If DbDelete = False Then   
		Exit Function           
    End If     						'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		LayerShowHide(0)
		Exit Function           
    End If     				                                                  '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    With frm1.vspdData
		If .MaxRows < 1 Then Exit Function
    
		.Focus

		.EditMode = True
	
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
    
		ggoSpread.CopyRow
    
		Call SetSpreadColor(.ActiveRow, .ActiveRow)

		.ReDraw = True
	End With
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
	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.EditUndo                                                  '☜: Protect system from crashing

	'frm1.vspdData.Col = C_HiddenInsideFlg
	
	'If UCase(Trim(frm1.vspdData.Text)) = "N" Then
	'	Call SetFieldProp(frm1.vspdData.Row, "N")
	'Else
	'	Call SetFieldProp(frm1.vspdData.Row, "Y")
	'End IF
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If
	
    With frm1
        .vspdData.ReDraw = False
		.vspdData.Focus
    
		ggoSpread.Source = .vspdData

        ggoSpread.InsertRow , iIntReqRows
	
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)

        .vspdData.ReDraw = True
	End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 
    Dim iIntCnt
    Dim iChrFlag

    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData
    lDelRows = ggoSpread.DeleteRow
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         '☜:화면 유형, Tab 유무 
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
    DbDelete = False														'⊙: Processing is NG
    
    LayerShowHide(1)
		
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtCharCd1=" & Trim(frm1.txtCharCd1.value)				'☜: 삭제 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
    
    DbQuery = False															    '⊙: Processing is NG
    LayerShowHide(1)
   
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'☜: 
		strVal = strVal & "&txtCharCd=" & Trim(frm1.hCharCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd=" & Trim(frm1.hCharValueCd.value)	'☆: 조회 조건 데이타 
		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'☜: 
		strVal = strVal & "&txtCharCd=" & Trim(frm1.txtCharCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd=" & Trim(frm1.txtCharValueCd.value)	'☆: 조회 조건 데이타 
		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=0"
    End If

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    DbQuery = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(ByVal isValue)													'☆: 조회 성공후 실행로직 
	Dim i
    '-----------------------
    'Reset variables area
    '-----------------------

    If isValue Then
		Call ggoOper.SetReqAttr(frm1.txtCharValueDigit, "Q")
		Call SetToolbar("11101111001111")
	Else
		Call ggoOper.SetReqAttr(frm1.txtCharValueDigit, "N")
		Call SetToolbar("11111111001111")
	End If

	With frm1
		ggoOper.SetReqAttr	.txtCharCd1, "Q"
	End With

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE

    lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim IntRows 
    Dim strVal, strDel
	Dim lGrpCnt
	Dim TmpBufferVal, TmpBufferDel
	Dim iValCnt, iDelCnt
	Dim iTotalStrVal, iTotalStrDel
	
	On Error Resume Next
	Err.Clear
	
    DbSave = False                                                          '⊙: Processing is NG
	
    LayerShowHide(1)
	
	lGrpCnt = 1
	iValCnt = 0: iDelCnt = 0
	ReDim TmpBufferVal(0): ReDim TmpBufferDel(0)

    With frm1
		.txtMode.Value = parent.UID_M0002											'☜: 저장 상태 
		.txtFlgMode.Value = lgIntFlgMode									'☜: 신규입력/수정 상태 
		
    '-----------------------
    'Data manipulate area
    '-----------------------
    For IntRows = 1 To .vspdData.MaxRows
		.vspdData.Row = IntRows
		.vspdData.Col = 0
		Select Case .vspdData.Text

            Case ggoSpread.InsertFlag												'☜: 신규 
				
				strVal = ""
				
				strVal = strVal & "C" & parent.gColSep & IntRows & parent.gColSep					'☜: C=Create

                .vspdData.Col = C_CharValueCd	'1
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_CharValueNm	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
				
				ReDim Preserve TmpBufferVal(iValCnt)
				
				TmpBufferVal(iValCnt) = strVal
				
				iValCnt = iValCnt + 1
				
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
				
				strVal = ""
				
				strVal = strVal & "U" & parent.gColSep	& IntRows & parent.gColSep					'☜: U=Update
				
                .vspdData.Col = C_CharValueCd	'1
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_CharValueNm	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBufferVal(iValCnt)
				
				TmpBufferVal(iValCnt) = strVal
				
				iValCnt = iValCnt + 1
                                               
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag																'☜: 삭제 
				
				strDel = ""
				
				strDel = strDel & "D" & parent.gColSep	& IntRows & parent.gColSep
				
                .vspdData.Col = C_CharValueCd	'1
                strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_CharValueNm	'2
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBufferDel(iDelCnt)
				
				TmpBufferDel(iDelCnt) = strDel
				
				iDelCnt = iDelCnt + 1
                                
                lGrpCnt = lGrpCnt + 1
        End Select
    Next
	
	iTotalStrDel = Join(TmpBufferDel, "")
	iTotalStrVal = Join(TmpBufferVal, "")
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = iTotalStrDel & iTotalStrVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 
    DbSave = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	frm1.txtCharCd.value = frm1.txtCharCd1.value
'	frm1.txtCharCd1.value = frm1.txtCharCd1.value
		
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
	IsOpenPop = False
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
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
'       					6. Tag부 
'#########################################################################################################-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사양항목등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>사양항목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharCd" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="사양항목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtCharNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="사양항목명"></TD>
									<TD CLASS=TD5 NOWRAP>사양값</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharValueCd" SIZE=18 MAXLENGTH=16 tag="11XXXU" ALT="사양값"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharValue" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharValueCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtCharValueNm" SIZE=20 MAXLENGTH=30 tag="14" ALT="사양값명"></TD>										
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
								<TD CLASS=TD5 NOWRAP>사양항목</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharCd1" SIZE=20 MAXLENGTH=18 tag="22XXXU" ALT="사양항목">&nbsp;<INPUT TYPE=TEXT NAME="txtCharNm1" SIZE=20 MAXLENGTH=40 tag="22XXXX" ALT="사양항목명"></TD>
								<TD CLASS=TD5 NOWRAP>사양값 자리수</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/b3b21ma1_I583314108_txtCharValueDigit.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<script language =javascript src='./js/b3b21ma1_I949479144_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hCharCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hCharValueCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
