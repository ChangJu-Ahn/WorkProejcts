<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1114MA1
'*  4. Program Name         : 검사분류별 불량률단위등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG070,PQBG080
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

Option Explicit
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_QRY_ID = "q1114mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_SAVE_ID = "q1114mb2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspClassNm		'= 1					'☆: Spread Sheet의 Column별 상수 
Dim C_DefectRatioUnitCd '= 2
Dim C_InspClassCd		'= 3

Dim IsOpenPop        

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE     				'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False   		                'Indicates that no value changed
    lgIntGrpCount = 0                           			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           			'initializes Previous Key
    lgLngCurRows = 0                            		'initializes Deleted Rows Count
    lgSortKey    = 1                            '⊙: initializes sort direction
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
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
		
		ggoSpread.Spreadinit "V20021224", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
	
		.MaxCols = C_InspClassCd + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCombo C_InspClassNm, "검사분류명", 40, 0, False
		ggoSpread.SSSetCombo C_DefectRatioUnitCd, "불량률단위", 52, 0, False
		ggoSpread.SSSetCombo C_InspClassCd, "검사분류코드", 10, 0, False

 		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols , True)
 		Call ggoSpread.SSSetColHidden( C_InspClassCd, C_InspClassCd , True)

		.ReDraw = true
		
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
		Call ggoSpread.SpreadLock(C_InspClassNm, -1, C_InspClassNm)
		Call ggoSpread.SSSetRequired(C_DefectRatioUnitCd, -1)
		Call ggoSpread.SpreadLock(frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols)
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
		ggoSpread.SSSetRequired C_InspClassNm,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_DefectRatioUnitCd, pvStartRow, pvEndRow
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

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	With frm1.vspdData
		
		strCboCd = lgF0 
		strCboNm = lgF1
		
		strCboCd=replace(strCboCd,Chr(11),vbTab)
		strCboNm=replace(strCboNm,Chr(11),vbTab)
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd, C_InspClassCd

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboNm, C_InspClassNm

		Call CommonQueryRs(" DEFECT_RATIO_UNIT_CD "," Q_DEFECT_RATIO_UNIT ","DEFECT_RATIO_UNIT_CD >= ''" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		strCboCd = lgF0 
		
		strCboCd = Replace(strCboCd,Chr(11),vbTab)
		strCboCd = Trim(strCboCd)
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd, C_DefectRatioUnitCd
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_InspClassNm		= 1					'☆: Spread Sheet의 Column별 상수 
	C_DefectRatioUnitCd = 2
	C_InspClassCd		= 3
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
		
		C_InspClassNm		= iCurColumnPos(1)					'☆: Spread Sheet의 Column별 상수 
		C_DefectRatioUnitCd = iCurColumnPos(2)
		C_InspClassCd		= iCurColumnPos(3)
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"									' 팝업 명칭 
	arrParam(1) = "B_PLANT"									' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""											' Where Condition
	arrParam(5) = "공장"										' 조건필드의 라벨 명칭 
		
	arrField(0) = "PLANT_CD"									' Field명(0)
	arrField(1) = "PLANT_NM"									' Field명(1)
	
	arrHeader(0) = "공장코드"										' Header명(0)
	arrHeader(1) = "공장명"									' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	End If	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	
	Call InitVariables                                                      	'⊙: Initializes local global variables
	Call InitSpreadSheet                                                    	'⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolBar("11001101001011")		'⊙: 버튼 툴바 제어 
	
	Call SetSingleFocus
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row	
	 '----------  Coding part  -------------------------------------------------------------
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
	 	
	Call SetPopupMenuItemInf("1101011111")         '화면별 설정 
    
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
  
'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData		
		.Row = Row    
		Select Case Col
			Case  C_InspClassNm
				.Col = Col
				intIndex = .Value
				.Col = C_InspClassCd
				.Value = intIndex
		End Select
	End With
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

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	
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
    
    Call SetSingleFocus										'⊙: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")	
End Sub 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	FncQuery = False                                                        					'⊙: Processing is NG
	
	Err.Clear                                                            		   					'☜: Protect system from crashing
	'-----------------------
	'Check previous data area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If
	
	'-----------------------
	'Check condition area
	'-----------------------
	If frm1.txtPlantCd.Value = "" Then
		Call DisplayMsgBox("189220", "X", "X", "X")
		Exit Function  
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
	
	FncQuery = True
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                                          					'⊙: Processing is NG
	
	Err.Clear              
	
	ggoSpread.Source = frm1.vspdData
    '-----------------------
	'Check previous data area
	'-----------------------
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	End If
	'-----------------------
	'Erase condition area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                          		'⊙: Lock  Suitable  Field
	Call InitVariables                                                      						'⊙: Initializes local global variables
	Call SetDefaultVal
	Call SetToolBar("11001101001011")		'⊙: 버튼 툴바 제어 
	Call SetSingleFocus
	FncNew = True
End Function

'========================================================================================
' Function Name : SetSingleFocus
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Sub SetSingleFocus()
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
	FncDelete = False                                                       					'⊙: Processing is NG
	Err.Clear                                                               						'☜: Protect system from crashing

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
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")
	FncDelete = True                                                        					'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False                                                  		       			'⊙: Processing is NG

	Err.Clear      
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
	End If
                    
	If frm1.txtPlantCd.Value = "" Then
		Call DisplayMsgBox("189220", "X", "X", "X")
		Exit Function  
	End If
    
	ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
	If Not ggoSpread.SSDefaultCheck Then Exit Function

	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False Then Exit Function                                  			'☜: Save db data
    
    FncSave = True
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	With frm1.vspdData
		If .MaxRows < 1 Then Exit function
		
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow
	    .Row = .ActiveRow
	    .Col = C_InspClassCd
	    .Text = ""
	    .Col = C_InspClassNm
	    .Text = ""
	    .ReDraw = True                                   					            '☜: Protect system from crashing
	End With
	FncCopy = true
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel= false
	If frm1.vspdData.MaxRows < 1 Then Exit function
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
		If imRow = "" Then Exit Function
	End If
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.InsertRow .ActiveRow, imRow
    	SetSpreadColor .ActiveRow, .ActiveRow + imRow -1
		.ReDraw = True
   	End With
    
    Call SetActiveCell(frm1.vspdData,C_InspClassNm,frm1.vspdData.ActiveRow,"M","X","X")
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
    
    If frm1.vspdData.MaxRows < 1 Then Exit function
	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData 

	lDelRows = ggoSpread.DeleteRow
	
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
    FncNext = false                                                   '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)					 '☜: 화면 유형 
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
			If IntRetCD = vbNo Then Exit Function
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
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

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
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
								    & "&txtPlantCd=" & Trim(frm1.hPlantCd.value) _
								    & "&lgStrPrevKey=" & lgStrPrevKey _
								    & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtPlantCd=" & Trim(.txtPlantCd.value) _
									& "&lgStrPrevKey=" & lgStrPrevKey _
									& "&txtMaxRows=" & .vspdData.MaxRows
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
	lgIntFlgMode = Parent.OPMD_UMODE
	Call SetToolBar("11001111001111")		'⊙: 버튼 툴바 제어 

	Dim posActiveRow
	If frm1.vspdData.MaxRows <= 100 Then
		posActiveRow = 1
	ElseIf frm1.vspdData.MaxRows > 100 Then
		posActiveRow = frm1.vspdData.MaxRows - 99
	End If
	
	Call SetActiveCell(frm1.vspdData,C_DefectRatioUnitCd,posActiveRow,"M","X","X")
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
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
	Dim arrVal
	Dim arrDel
	
	Dim iColSep
	Dim iRowSep
	Dim iInsertFlag
	Dim iUpdateFlag
	Dim iDeleteFlag
	
	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing
	
	iColSep     = Parent.gColSep
	iRowSep     = Parent.gRowSep
	iInsertFlag = ggoSpread.InsertFlag
	iUpdateFlag = ggoSpread.UpdateFlag
	iDeleteFlag = ggoSpread.DeleteFlag                                                 	'☜: Protect system from crashing

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
							     & GetSpreadText(.vspdData,C_InspClassCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_DefectRatioUnitCd,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1

				Case iUpdateFlag

					strVal = "U" & iColSep _
							 	 & GetSpreadText(.vspdData,C_InspClassCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_DefectRatioUnitCd,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)					
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1

				Case iDeleteFlag

					strDel = "D" & iColSep _
								 & GetSpreadText(.vspdData,C_InspClassCd,lRow,"X","X") & iColSep _
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
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)				'☜: 비지니스 ASP 를 가동 
	End With
	DbSave = True                                                           		'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()	
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
					<TD WIDTH=10></TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사분류별 불량률단위</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD656">
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU" ><IMG align=top height=20 name=btnPlantCd onclick=vbscript:OpenPlant() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH="40" tag="14" >
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="22" TITLE="SPREAD"> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
