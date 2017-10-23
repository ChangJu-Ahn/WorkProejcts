<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Manufacturing Instruction Detail Entry
'*  3. Program ID           : P1206ma1
'*  4. Program Name         : Manufacturing Instruction Detail Entry
'*  5. Program Desc         : 
'*  6. Comproxy List        : this program use source of HR Module
'*  7. Modified date(First) : 2002/03/19
'*  8. Modified date(Last)  : 2002/12/02
'*  9. Modifier (First)     : Chen, Jae Hyun
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

Const BIZ_PGM_QRY_ID = "p1206mb1.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID= "p1206mb2.asp"											'☆: 비지니스 로직 ASP명 

Dim C_WICd
Dim C_WIDesc
Dim C_ValidStartDt
Dim C_ValidEndDt

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop

Dim BaseDate, StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_WICd			= 1
	C_WIDesc		= 2
	C_ValidStartDt	= 3
	C_ValidEndDt	= 4
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    
    lgIntGrpCount = 100                         'initializes Group View Size

    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                                       '⊙: initializes sort direction
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtValidDt.text = StartDate
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

		.MaxCols = C_ValidEndDt + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit	C_WICd, "단위작업", 18,,,10,2
		ggoSpread.SSSetEdit	C_WIDesc, "단위작업내역", 76,,,100
		ggoSpread.SSSetDate C_ValidStartDt, "유효시작일", 10, 2, parent.gDateFormat
		ggoSpread.SSSetDate C_ValidEndDt, "유효종료일", 10, 2, parent.gDateFormat
    
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(1)										'frozen 기능추가 
    
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
		ggoSpread.SpreadLock	C_WICd,			-1, C_WICd
		ggoSpread.SpreadUnLock	C_WIDesc,		-1, C_WIDesc
		ggoSpread.SpreadLock	C_ValidStartDt, -1, C_ValidStartDt
		ggoSpread.SpreadUnLock	C_ValidEndDt,	-1, C_ValidEndDt
    
		ggoSpread.SSSetRequired	C_WIDesc,		-1
		ggoSpread.SSSetRequired	C_ValidEndDt,	-1
		ggoSpread.SSSetProtected .vspdData.MaxCols,	-1
    
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
		ggoSpread.SSSetRequired		C_WICd,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_WIDesc,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ValidStartDt,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ValidEndDt,	pvStartRow, pvEndRow
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
			C_WICd			= iCurColumnPos(1)
			C_WIDesc		= iCurColumnPos(2)
			C_ValidStartDt	= iCurColumnPos(3)
			C_ValidEndDt	= iCurColumnPos(4)
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
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'------------------------------------------  OpenWICdPopup()  -------------------------------------------------
'	Name : OpenWICDPopup()
'	Description : WcPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWICD(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "단위작업팝업"	
	arrParam(1) = "P_MFG_INSTRUCTION_DETAIL"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "VALID_END_DT >=  " & FilterVar(BaseDate, "''", "S") & " AND VALID_START_DT <= " &  " " & FilterVar(BaseDate, "''", "S") & ""
	arrParam(5) = "단위작업"			
	
    arrField(0) = "MFG_INSTRUCTION_DTL_CD"	
    arrField(1) = "MFG_INSTRUCTION_DTL_DESC"	
    
    arrHeader(0) = "단위작업"		
    arrHeader(1) = "단위작업내역"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWICD(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWICD.focus
	
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetWICD()
'	Description : Work Instruction Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWICD(Byval arrRet)
	With frm1
		.txtWICD.Value = arrRet(0)
		.txtWINM.Value = arrRet(1) 
	End With
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    '----------  Coding part  -------------------------------------------------------------

   ' Call InitComboBox
    Call SetToolbar("11001101001011")										'⊙: 버튼 툴바 제어 
    Call SetDefaultVal
    Call InitVariables
    
	frm1.txtWICd.focus 
	Set gActiveElement = document.activeElement 
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("1101111111")

    If frm1.vspdData.MaxRows <= 0 Or Col < 1 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim strEndDt, strStartDt
	With frm1.vspdData
		Select Case Col

		    Case C_ValidEndDt
				.Row = Row
				.Col = C_ValidEndDt
				strEndDt = .Text
				
				.Col = C_ValidStartDt
				strStartDt = .Text
				
				If strEndDt = "" or strStartDt = "" Then Exit Sub	
				
				If CompareDateByFormat(strStartDt,strEndDt,"유효시작일","유효종료일","970025",parent.gDateFormat,parent.gComDateType,True) = False Then  'If CDate(strStartDt) > CDate(strEndDt) Then  
					.Col = C_ValidEndDt
					.Text = ""
					Exit Sub
				End If
				
			Case C_ValidStartDt
				.Row = Row
				.Col = C_ValidEndDt
				strEndDt = .Text
				
				.Col = C_ValidStartDt
				strStartDt = .Text
				
				If strEndDt = "" or strStartDt = "" Then Exit Sub	
				
				If CompareDateByFormat(strStartDt,strEndDt,"유효시작일","유효종료일","970025",parent.gDateFormat,parent.gComDateType,True) = False Then  'If CDate(strStartDt) > CDate(strEndDt) Then  
					.Col = C_ValidStartDt
					.Text = ""
					Exit Sub
				End If
				
		End Select
		
	End With

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row >= NewRow Then
        Exit Sub
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
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
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtValidDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtValidDt.Focus
	End If 
End Sub


'=======================================================================================================
'   Event Name : txtValidDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtValidDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidDt_Change() 
	'lgBlnFlgChgValue = True 
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
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtWICd.value = "" Then
		frm1.txtWINm.value = ""
	End If
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	Call ggoSpread.ClearSpreadData
    Call InitVariables

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
    End If     
    
    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
     
    FncNew = True																	'⊙: Processing is OK

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆: 아래 메세지를 DB화 해서 이 라인으로 대체 
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
    End If         
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    Dim strStartDt, strEndDt, i
   
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False AND ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    
    ggoSpread.Source = frm1.vspdData

    If Not chkField(Document, "2") Then
		Exit Function
	End If
	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    With frm1.vspdData
    
    For i = 0 to .Maxrows
		.Row = i + 1
		.Col = 0
		If  .Text = ggoSpread.InsertFlag or .Text = ggoSpread.UpdateFlag Then  
				
			.Col = C_ValidEndDt
			strEndDt = .Text
					
			.Col = C_ValidStartDt
			strStartDt = .Text
					
			If CompareDateByFormat(strStartDt,strEndDt,"유효시작일","유효종료일","970025",parent.gDateFormat,parent.gComDateType,True) = False Then  'If CDate(strStartDt) > CDate(strEndDt) Then  
				.Col = C_ValidStartDt
				.Text = ""
				Exit Function
			End If
		End If	
    Next
    
    End With
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If          '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
		
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement 
	frm1.vspdData.EditMode = True
	frm1.vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
    
    frm1.vspdData.Col = C_WICD
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_ValidStartDt
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Text = frm1.txtValidDt.text
      
    frm1.vspdData.ReDraw = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData

    ggoSpread.EditUndo
    
    Call SetToolbar("11001111001111")
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim iIntReqRows, iIntCnt

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If

	With frm1
		.vspdData.Focus
		Set gActiveElement = document.activeElement
		ggoSpread.Source = .vspdData
		.vspdData.EditMode = True
		.vspdData.ReDraw = False
    
		ggoSpread.InsertRow , iIntReqRows
    
		.vspdData.ReDraw = True
    
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)
		For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + iIntReqRows - 1
			.vspdData.Row = iIntCnt
			.vspdData.Col = C_ValidStartDt
			.vspdData.Text = .hValidDt.Value
    
			.vspdData.Col = C_ValidEndDt
			.vspdData.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
		Next
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
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
	Set gActiveElement = document.activeElement
    ggoSpread.Source = frm1.vspdData
    
	lDelRows = ggoSpread.DeleteRow
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                               '☜: Protect system from crashing
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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

    DbQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtWICd=" & UCase(Trim(.hWICd.value))				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtValidDt=" & Trim(.hValidDt.Value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & UCase(Trim(lgStrPrevKey))
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtWICd=" & UCase(Trim(.txtWICd.value))				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtValidDt=" & Trim(.txtValidDt.Text)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & UCase(Trim(lgStrPrevKey))
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	End If   
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = false
    
    Call SetToolbar("11001111001111")
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim iColSep
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt

    DbSave = False                                                          '⊙: Processing is NG
	     
    LayerShowHide(1)
		
	With frm1
		
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
	    
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
 
        Select Case .vspdData.Text

			   Case ggoSpread.InsertFlag                                      '☜: Update
                     
														   strVal = ""
                                                           strVal = strVal & "C" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_WICd				   : strVal = strVal & UCase(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WIDesc	           : strVal = strVal & .vspdData.Text & parent.gColSep
                    .vspdData.Col = C_ValidStartDt	       : strVal = strVal & .vspdData.Text & parent.gColSep
                    .vspdData.Col = C_ValidEndDt	       : strVal = strVal & .vspdData.Text & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    ReDim Preserve TmpBufferVal(iValCnt)
                    TmpBufferVal(iValCnt) = strVal
                    iValCnt = iValCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
														   strVal = ""
                                                           strVal = strVal & "U" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_WICd		           : strVal = strVal & UCase(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WIDesc	           : strVal = strVal & .vspdData.Text & parent.gColSep
                    .vspdData.Col = C_ValidEndDt	       : strVal = strVal & .vspdData.Text & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    ReDim Preserve TmpBufferVal(iValCnt)
                    TmpBufferVal(iValCnt) = strVal
                    iValCnt = iValCnt + 1
                    
               Case ggoSpread.DeleteFlag                                      '☜: Delete
														   strDel = ""
                                                           strDel = strDel & "D" & parent.gColSep
                                                           strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_WICd			     : strDel = strDel & UCase(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    ReDim Preserve TmpBufferDel(iDelCnt)
                    TmpBufferDel(iDelCnt) = strDel
                    iDelCnt = iDelCnt + 1
                   
	    End Select
              
    Next

	iTotalStrDel = Join(TmpBufferDel, "")
	iTotalStrVal = Join(TmpBufferVal, "")
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = iTotalStrDel & iTotalStrVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	'frm1.txtRoutNo.value = frm1.txtRoutingNo.value 
	
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

	DbDelete = False														'⊙: Processing is NG
	
	LayerShowHide(1)
		
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtRoutingNo=" & Trim(frm1.txtRoutingNo.value)				'☜: 삭제 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG 
End Function
Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>단위작업등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>단위작업</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtWICd" SIZE=18 MAXLENGTH=10 tag="11NXXU" ALT="단위작업"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWICd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenWICD frm1.txtWICd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtWINm" SIZE=60 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p1206ma1_I997490954_txtValidDt.js'></script>
									</TD>
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
								<TD HEIGHT="100%" COLSPAN = 2>
								<script language =javascript src='./js/p1206ma1_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hWICd" tag="24"><INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
