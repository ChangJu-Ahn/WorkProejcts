<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1112MA1
'*  4. Program Name         : 검사항목등록 
'*  5. Program Desc         : 검사항목등록 
'*  6. Component List       : PQBG030,PQBG040
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit															'☜: indicates that All variables must be declared in advance
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_QRY_ID = "q1112mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_SAVE_ID = "q1112mb2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_InspItemCd '= 1
Dim C_InspItemNm '= 2
Dim C_InspItemClassCd '= 3
Dim C_InspItemClassPopup '= 4
Dim C_InspItemClassNm '= 5
Dim C_InspCharNm '= 6
'--------- Hidden --------
Dim C_InspCharCd '= 7
'-------------------------

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   	'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    		'Indicates that no value changed
    lgIntGrpCount = 0                           			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           			'initializes Previous Key
    lgLngCurRows = 0                            		'initializes Deleted Rows Count  
    lgSortKey    = 1                            '⊙: initializes sort direction
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021225", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
		
		.MaxCols = C_InspCharCd + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
		
		Call ggoSpread.SSSetEdit(C_InspItemCd,		"검사항목코드", 20, 0, -1, 5, 2)
		Call ggoSpread.SSSetEdit(C_InspItemNm,		"검사항목명", 20, 0, -1, 40)
		Call ggoSpread.SSSetEdit(C_InspItemClassCd, "검사항목분류코드", 20, 0, -1, 2, 2)
		Call ggoSpread.SSSetButton(C_InspItemClassPopup)
		Call ggoSpread.SSSetEdit(C_InspItemClassNm, "검사항목분류명", 20, 0, -1, 40)
		Call ggoSpread.SSSetCombo(C_InspCharCd,		"표시속성코드", 5, 0, False)
		Call ggoSpread.SSSetCombo(C_InspCharNm,		"표시속성", 20, 0, False)

 		Call ggoSpread.MakePairsColumn(C_InspItemClassCd, C_InspItemClassPopup)
 		
 		Call ggoSpread.SSSetColHidden(C_InspCharCd, C_InspCharCd, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		.ReDraw = true
	
		Call SetSpreadLock
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	With frm1.vspdData
		.ReDraw = False
		Call ggoSpread.SpreadLock(C_InspItemCd, -1, C_InspItemCd)
		Call ggoSpread.SpreadLock(C_InspItemClassNm, -1, C_InspItemClassNm)
		Call ggoSpread.SSSetRequired(C_inspItemNm, -1)
		Call ggoSpread.SSSetRequired(C_InspItemClassCd, -1)
		Call ggoSpread.SSSetRequired(C_InspCharNm, -1)
		Call ggoSpread.SpreadLock(frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols)
		.ReDraw = True
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False
		Call ggoSpread.SSSetRequired(C_InspItemCd, pvStartRow, pvEndRow)
		Call ggoSpread.SSSetRequired(C_inspItemNm, pvStartRow, pvEndRow)
		Call ggoSpread.SSSetRequired(C_InspItemClassCd, pvStartRow, pvEndRow)
		Call ggoSpread.SSSetProtected(C_InspItemClassNm, pvStartRow, pvEndRow)
		Call ggoSpread.SSSetRequired(C_InspCharNm, pvStartRow, pvEndRow)
		.vspdData.ReDraw = True
	End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim strCboCd
	Dim strCboNm

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0023", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	
	With frm1.vspdData
		strCboCd = lgF0 
		strCboNm = lgF1
	
		strCboCd=replace(strCboCd,Chr(11),vbTab)
		strCboNm=replace(strCboNm,Chr(11),vbTab)
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd, C_InspCharCd

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboNm, C_InspCharNm
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_InspItemCd = 1
	C_InspItemNm = 2
	C_InspItemClassCd = 3
	C_InspItemClassPopup = 4
	C_InspItemClassNm = 5
	C_InspCharNm = 6
	'-------- Hidden --------
	C_InspCharCd = 7
	'------------------------	
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
			
			C_InspItemCd = iCurColumnPos(1)
			C_InspItemNm = iCurColumnPos(2)
			C_InspItemClassCd = iCurColumnPos(3)
			C_InspItemClassPopup = iCurColumnPos(4)
			C_InspItemClassNm = iCurColumnPos(5)
			C_InspCharNm = iCurColumnPos(6)
			'-------- Hidden --------
			C_InspCharCd = iCurColumnPos(7)
			'------------------------
 	End Select
End Sub

'------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : InspItemPlant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	OpenInspItem = false
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "검사항목 팝업"					' 팝업 명칭 
	arrParam(1) = "Q_INSPECTION_ITEM"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtInspitemCd.Value)			' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "검사항목"			
	
	arrField(0) = "INSP_ITEM_CD"							' Field명(0)
	arrField(1) = "INSP_ITEM_NM"						' Field명(1)
	
	arrHeader(0) = "검사항목코드"						' Header명(0)
	arrHeader(1) = "검사항목명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtInspItemCd.Value    = arrRet(0)		
		frm1.txtInspItemNm.Value    = arrRet(1)	
	End If	
	
	frm1.txtInspItemCd.Focus
	Set gActiveElement = document.activeElement				
	
	OpenInspItem = true
End Function

'------------------------------------------  OpenInspItemClass()  -------------------------------------------------
'	Name : OpenInspItemClass()
'	Description : Inspection Item Class PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItemClass(ByVal IRow, Byval strCode)
	OpenInspItemClass = false
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "검사항목분류팝업"					' 팝업 명칭 
	arrParam(1) = "B_MINOR"						' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("Q0003", "''", "S") & ""					' Where Condition
	arrParam(5) = "검사항목분류"			
	
	arrField(0) = "MINOR_CD"						' Field명(0)
	arrField(1) = "MINOR_NM"						' Field명(1)
	
	arrHeader(0) = "검사항목분류코드"					' Header명(0)
	arrHeader(1) = "검사항목분류명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	With frm1
		
		If arrRet(0) <> "" Then
			.vspdData.Col = C_InspItemClassCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_InspItemClassNm
			.vspdData.Text = arrRet(1)
			Call vspdData_Change(.vspdData.Col, .vspdData.Row)						 ' 변경이 읽어났다고 알려줌 
		End If	

		Call SetActiveCell(frm1.vspdData,C_InspItemClassCd,IRow,"M","X","X")
		Set gActiveElement = document.activeElement

	End With
	
	OpenInspItemClass = true
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	
	Call InitSpreadSheet                                                    '⊙: setup the Spread sheet
	Call InitComboBox
	Call InitVariables                                                      '⊙: Initializes local global variables
	
	Call SetToolBar("11001101001011")		'⊙: 버튼 툴바 제어 
	frm1.txtInspItemCd.focus
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
    
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
	
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
 	
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

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
	Call ggoOper.LockField(Document, "Q")	
 	'------ Developer Coding part (End) 	
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row	
	frm1.vspdData.Col = Col
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
		Select Case Col
			Case  C_InspCharNm
				.Col = Col
				intIndex = .Value
				.Col = C_InspCharCd
				.Value = intIndex
		End Select
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_InspItemClassPopUp Then
			.Col = C_InspItemClassCd
			.Row = Row
			If OpenInspItemClass(Row, .Text) = False Then	Exit Sub
		End If
	
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then Exit Sub
	End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
    
	FncQuery = False                                                        '⊙: Processing is NG
    
	Err.Clear                                                               '☜: Protect system from crashing

	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
	'Check previous data area
	'-----------------------
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
    
	'-----------------------
	'Erase contents area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Call InitVariables
	
	'-----------------------
	'Query function call area
	'-----------------------
	
	If DbQuery = False then	Exit Function
																				'☜: Query db data
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
		
	ggoSpread.Source = frm1.vspdData

    '-----------------------
	'Check previous data area
	'-----------------------
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Contents  Field
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call InitVariables                                                      '⊙: Initializes local global variables
	
	Call SetToolBar("11001101001011")		'⊙: 버튼 툴바 제어 
	frm1.txtInspItemCd.focus
	Set gActiveElement = document.activeElement 
	FncNew = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
	FncDelete = False                                                       '⊙: Processing is NG
	Err.Clear            
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then                                                '☜: Delete db data
		Exit Function                                                        '☜:
	End If
    
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Contents  Field
    
	FncDelete = True
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	FncSave = False                                                         '⊙: Processing is NG
	Err.Clear                                                               '☜: Protect system from crashing
	On Error Resume Next                                                    '☜: Protect system from crashing
	
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If

    '-----------------------
	'Check content area
	'-----------------------
	ggoSpread.Source = frm1.vspdData							 '⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then					 '⊙: Check required field(Multi area)
    	Exit Function
    End If
	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then	
		Exit Function
	End If			                                                  '☜: Save db data

	FncSave = True                                                          '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	With frm1
		If .vspdData.MaxRows < 1 then
	    		Exit function
    	End if
		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData	
		ggoSpread.CopyRow
		
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow)
	    
	    .vspdData.Row = .vspdData.ActiveRow
	    .vspdData.Col = C_InspItemCd
	    .vspdData.Text = ""
	    frm1.vspdData.ReDraw = True                                   					            '☜: Protect system from crashing
	End With
	FncCopy = true
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = false
	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End if
	ggoSpread.Source = frm1.vspdData	
	ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	FncCancel = true
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = false
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If
	End If	
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
    	ggoSpread.InsertRow .vspdData.ActiveRow, imRow
    	Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1)
		.vspdData.ReDraw = True
    End With
    
    FncInsertRow = true
    
    Set gActiveElement = document.ActiveElement    
    If Err.number = 0 Then FncInsertRow = True  
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false
	Dim lDelRows
	Dim iDelRowCnt, i
    
    With frm1
		If .vspdData.MaxRows < 1 then
			Exit function
		End if	
		.vspdData.focus
		ggoSpread.Source = .vspdData 
		lDelRows = ggoSpread.DeleteRow
	End With
	FncDeleteRow = true
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    FncPrev = false                                                 '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    FncNext = false                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
	FncExit = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
	Err.Clear                                                               					'☜: Protect system from crashing
	Call LayerShowHide(1)
	DbQuery = False
	
	With frm1	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001  				 '☜:
			strVal = strVal & "&txtInspItemCd=" & .hInspItemCd.value			'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey					
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001   			'☜:
			strVal = strVal & "&txtInspItemCd=" & Trim(.txtInspItemCd.Value)			'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
				
		Call RunMyBizASP(MyBizASP, strVal)							'☜: 비지니스 ASP 를 가동 
		
		DbQuery = True                                                          					'⊙: Processing is NG
	End With
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    Call SetToolBar("11001111001111")		'⊙: 버튼 툴바 제어 

	Call ggoOper.LockField(Document, "Q")	

	Dim posActiveRow
	If frm1.vspdData.MaxRows <= 100 Then
		posActiveRow = 1		
		Call SetActiveCell(frm1.vspdData,C_InspItemNm,posActiveRow,"M","X","X")		
		Set gActiveElement = document.activeElement
	End If	
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpInsCnt
	Dim lGrpDelCnt 
	Dim strVal 
	Dim strDel

	Dim iColSep
	Dim iRowSep
	Dim iInsertFlag
	Dim iUpdateFlag
	Dim iDeleteFlag
	Dim arrVal
	Dim arrDel
			
	Call LayerShowHide(1)
	
	DbSave = False
	
	On Error Resume Next                                                   '☜: Protect system from crashing

	iColSep     = Parent.gColSep
	iRowSep     = Parent.gRowSep
	iInsertFlag = ggoSpread.InsertFlag
	iUpdateFlag = ggoSpread.UpdateFlag
	iDeleteFlag = ggoSpread.DeleteFlag

	With frm1
		.txtMode.value = Parent.UID_M0002
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpInsCnt = 0
		lGrpDelCnt = 0

    	ReDim arrVal(0)
		ReDim arrDel(0)

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row = lRow
			.vspdData.Col = 0
			Select Case .vspdData.Text
				Case iInsertFlag

					strVal = "C" & iColSep _
							     & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemNm,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemClassCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspCharCd,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1

				Case iUpdateFlag

					strVal = "U" & iColSep _
							 	 & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemNm,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemClassCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspCharCd,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)					
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1
					
				Case iDeleteFlag
					
					strDel = "D" & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & iColSep _
								 & iColSep & iColSep & iColSep _
								 & CStr(lRow) & iRowSep
									
					ReDim Preserve arrDel(lGrpDelCnt)
					arrDel(lGrpDelCnt) = strDel
					lGrpDelCnt = lGrpDelCnt + 1
					
			End Select
		Next
		
		strVal = Join(arrVal,"")
		strDel = Join(arrDel,"")
		
		.txtMaxRows.value = lGrpInsCnt + lGrpDelCnt
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)						'☜: 비지니스 ASP 를 가동 
	End With
	DbSave = True                                                           						'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()																			'☆: 저장 성공후 실행 로직 
	DbSaveOk = false
  	Call InitVariables
  	frm1.vspdData.MaxRows = 0
	Call MainQuery()
	DbSaveOk = true
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete = false
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사항목 등록</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWPAP>검사항목</TD>
									<TD CLASS="TD656" NOWPAP>
										<INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="검사항목" tag="11XXXU" ><IMG align=top height=20 name=btnInspItemCd onclick=vbscript:OpenInspItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtInspItemNm" tag="14" >
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
						<TABLE WIDTH="100%" HEIGHT="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/q1112ma1_I148279279_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" tabindex=-1 ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hInspItemCd" tag="24" tabindex=-1 >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
