<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1111MA1
'*  4. Program Name         : 측정기 정보등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG010,PD6G020
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_QRY_ID = "q1111mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_SAVE_ID = "q1111mb2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_MeasmtEquipmtCd			'= 1															'☆: Spread Sheet의 Column별 상수 
Dim C_MeasmtEquipmtNm			'= 2

Dim IsOpenPop        

'========================================================================================================= 
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE			'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False					'Indicates that no value changed
    lgIntGrpCount = 0							'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""							'initializes Previous Key
    lgLngCurRows = 0							'initializes Deleted Rows Count
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

'========================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021225", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
		
		.MaxCols = C_MeasmtEquipmtNm + 1
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		Call ggoSpread.SSSetEdit(C_MeasmtEquipmtCd, "측정기코드", 20, 0, -1, 10, 2)
		Call ggoSpread.SSSetEdit(C_MeasmtEquipmtNm, "측정기명", 70, 0, -1, 40)
 		
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		.ReDraw = true

		Call SetSpreadLock
	End With
End Sub

'========================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	With frm1.vspdData
		.ReDraw = False
		Call ggoSpread.SpreadLock(C_MeasmtEquipmtCd, -1, C_MeasmtEquipmtCd)
		Call ggoSpread.SSSetRequired(C_MeasmtEquipmtNm, -1)
		Call ggoSpread.SpreadLock(frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols)
		.ReDraw = True
	End With
End Sub

'========================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal QueryStatus)
	With frm1.vspdData
		.ReDraw = False
		
		If QueryStatus = 1 Then 
			Call ggoSpread.SSSetProtected(C_MeasmtEquipmtCd, pvStartRow, pvEndRow)
			Call ggoSpread.SSSetRequired(C_MeasmtEquipmtNm, pvStartRow, pvEndRow)
		Else
			Call ggoSpread.SSSetRequired(C_MeasmtEquipmtCd, pvStartRow, pvEndRow)
			Call ggoSpread.SSSetRequired(C_MeasmtEquipmtNm, pvStartRow, pvEndRow)
		End If
		
		.ReDraw = True		
	End With
End Sub

'=========================================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_MeasmtEquipmtCd = 1
	C_MeasmtEquipmtNm = 2
End Sub

'=========================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
 			C_MeasmtEquipmtCd = iCurColumnPos(1)
			C_MeasmtEquipmtNm = iCurColumnPos(2)
 	End Select
End Sub

'=========================================================================================================
'	Name : OpenMeasmtEquipmt
'	Description : InspItemPlant PopUp
'=========================================================================================================
Function OpenMeasmtEquipmt()
	OpenMeasmtEquipmt = false
	 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "측정기 팝업"								' 팝업 명칭 
	arrParam(1) = "Q_Measurement_Equipment"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtMeasmtEquipmtCd.Value)			' Code Condition
	arrParam(3) = ""											' Name Cindition
	arrParam(4) = ""											' Where Condition
	arrParam(5) = "측정기"			
	
	arrField(0) = "MEASMT_EQUIPMT_CD"							' Field명(0)
	arrField(1) = "MEASMT_EQUIPMT_NM"							' Field명(1)
	
	arrHeader(0) = "측정기코드"								' Header명(0)
	arrHeader(1) = "측정기명"								' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtMeasmtEquipmtCd.Value    = arrRet(0)
		frm1.txtMeasmtEquipmtNm.Value    = arrRet(1)
	End If	
	frm1.txtMeasmtEquipmtCd.focus
	Set gActiveElement = document.activeElement
	
	OpenMeasmtEquipmt = true
End Function

'========================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")									'⊙: Lock  Suitable  Field
	
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetToolBar("11101101001011")										'⊙: 버튼 툴바 제어 
	Call SetSingleFocus
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"
    
    If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1101110111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("1101011111")         '화면별 설정 
	End If
	
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
		If CheckRunningBizProcess = True Then Exit Sub
 	
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
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
	Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.MaxRows, 1)
	Call ggoSpread.SSSetProtected(C_MeasmtEquipmtCd, 1, frm1.vspdData.MaxRows)
	Call ggoOper.LockField(Document, "Q") 
 	'------ Developer Coding part (End) 	
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
	
	FncQuery = False																'⊙: Processing is NG
	
	Err.Clear                                                            			'☜: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData
	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		
		If IntRetCD = vbNo Then Exit Function
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
	If DbQuery = False then Exit Function
	
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
   
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")											'⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call InitVariables																'⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolBar("11001101001011")												'⊙: 버튼 툴바 제어 
    Call SetSingleFocus
    FncNew = True																	'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : SetSingleFocus
' Function Desc : 
'========================================================================================
Sub SetSingleFocus()
	frm1.txtMeasmtEquipmtCd.focus
	Set gActiveElement = document.activeElement 
End Sub

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
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then Exit Function
	
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	
	FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False                                                  					'⊙: Processing is NG

	Err.Clear                                                            	 			'☜: Protect system from crashing
	On Error Resume Next                                           						'☜: Protect system from crashing
	
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
	ggoSpread.Source = frm1.vspdData													'⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then Exit Function										'⊙: Check required field(Multi area)
    
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then Exit Function	
    
	FncSave = True                                      								'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	
	With frm1
		If .vspdData.MaxRows < 1 then Exit function

		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData	
		ggoSpread.CopyRow

		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow, 0)

	    .vspdData.Row = .vspdData.ActiveRow
	    .vspdData.Col = C_MeasmtEquipmtCd
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
	
	If frm1.vspdData.MaxRows < 1 then Exit function

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
    	Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1, 0)
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
		If .vspdData.MaxRows < 1 then Exit function

		.vspdData.focus
		ggoSpread.Source = .vspdData 
	     '----------  Coding part  -------------------------------------------------------------   
	
		lDelRows = ggoSpread.DeleteRow
	
	End With
	
	FncDeleteRow = true
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    FncPrev= false                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    FncNext= false                                                      '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 '☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then  Exit Sub

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

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
			If IntRetCD = vbNo Then Exit Function
		End If
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    	
	Call LayerShowHide(1)
	
	Err.Clear                                                               					'☜: Protect system from crashing
	
	DbQuery = False                                                        					 '⊙: Processing is NG

	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtMeasmtEquipmtCd=" & .hMeasmtEquipmtCd.value _
									& "&lgStrPrevKey=" & lgStrPrevKey _
									& "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtMeasmtEquipmtCd=" & Trim(.txtMeasmtEquipmtCd.value) _
									& "&lgStrPrevKey=" & lgStrPrevKey _
									& "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)
	End With
	
	DbQuery = True 
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()		
	DbQueryOk = false												'☆: 조회 성공후 실행로직 
	
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE			
	Call SetToolBar("11101111001111")		'⊙: 버튼 툴바 제어 
	
	Call ggoOper.LockField(Document, "Q")	
	
	Dim posActiveRow
	If frm1.vspdData.MaxRows <= 100 Then
		posActiveRow = 1
		Call SetActiveCell(frm1.vspdData,C_MeasmtEquipmtNm,posActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data Save and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpInsCnt
	Dim lGrpDelCnt 
	Dim strDel
	Dim strVal

	Dim iColSep
	Dim iRowSep
	Dim iInsertFlag
	Dim iUpdateFlag
	Dim iDeleteFlag
	Dim arrVal
	Dim arrDel
	
	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
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
							     & GetSpreadText(.vspdData,C_MeasmtEquipmtCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MeasmtEquipmtNm,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1

				Case iUpdateFlag

					strVal = "U" & iColSep _
							 	 & GetSpreadText(.vspdData,C_MeasmtEquipmtCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MeasmtEquipmtNm,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)					
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1
					
				Case iDeleteFlag

					strDel = "D" & iColSep _
								 & GetSpreadText(.vspdData,C_MeasmtEquipmtCd,lRow,"X","X") & iColSep _
								 & iColSep _
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
			
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	End With
    DbSave = True
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
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
	<%' SPACE AREA%>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<%' TAB, REFERENCE AREA %>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>측정기 정보 등록</font></td>
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
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWPAP>측정기</TD>
									<TD CLASS="TD656" NOWPAP>
										<INPUT TYPE=TEXT NAME="txtMeasmtEquipmtCd" SIZE="10" MAXLENGTH="10" ALT="측정기" tag="11XXXU" ><IMG align=top height=20 name=btnMeasmtEquipmt onclick=vbscript:OpenMeasmtEquipmt() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtMeasmtEquipmtNm" tag="14" >
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
									<script language =javascript src='./js/q1111ma1_I620839715_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hMeasmtEquipmtCd" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
