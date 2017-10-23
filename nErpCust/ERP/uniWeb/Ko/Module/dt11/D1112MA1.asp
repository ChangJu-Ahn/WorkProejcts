<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Digital Tax
'*  2. Function Name        : 
'*  3. Program ID           : D1112MA1
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID      = "D1112MB1.asp"
Const BIZ_PGM_SAVE_ID = "D1112MB2.asp" 

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim 	C1_default_flag
Dim 	C1_name
Dim 	C1_email
Dim 	C1_dept_nm
Dim 	C1_ph_no
Dim 	C1_remarks
Dim 	C1_bp_cd
Dim 	C1_bp_seq


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
		
		.MaxCols = C1_bp_seq + 1
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCheck	C1_default_flag	, "주담당자" ,	8,,,1		
		ggoSpread.SSSetEdit	C1_name		, "이름" ,	20	,,,	30
		ggoSpread.SSSetEdit	C1_email	, "E-Mail" ,	20	,,,	40
		ggoSpread.SSSetEdit	C1_dept_nm	, "부서명" ,	20	,,,	40
		ggoSpread.SSSetEdit	C1_ph_no	, "전화번호" ,	20	,,,	20
		ggoSpread.SSSetEdit	C1_remarks	, "비고" ,	20	,,,	1000
		ggoSpread.SSSetEdit	C1_bp_cd	, "거래처" ,	20		
		ggoSpread.SSSetEdit	C1_bp_seq	, "순번" ,	20		

		
 		Call ggoSpread.SSSetColHidden(C1_bp_cd, C1_bp_seq, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols   , .MaxCols   , True)
 		
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

		ggoSpread.SpreadUnLock	C1_default_flag	, -1, 	C1_default_flag
		ggoSpread.SpreadUnLock	C1_name	, -1, 	C1_name
		ggoSpread.SpreadUnLock	C1_email	, -1, 	C1_email
		ggoSpread.SpreadUnLock	C1_dept_nm	, -1, 	C1_dept_nm
		ggoSpread.SpreadUnLock	C1_ph_no	, -1, 	C1_ph_no
		ggoSpread.SpreadUnLock	C1_remarks	, -1, 	C1_remarks
		ggoSpread.SpreadLock	C1_bp_cd	, -1, 	C1_bp_cd
		ggoSpread.SpreadLock	C1_bp_seq	, -1, 	C1_bp_seq
		
		ggoSpread.SSSetRequired  C1_name , -1
        ggoSpread.SSSetRequired C1_email , -1
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
			ggoSpread.SSSetRequired  C1_name , pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C1_email ,pvStartRow, pvEndRow
		
		.ReDraw = True		
	End With
End Sub


Sub InitGridComboBox()
	Dim IntRetCD1

	on error resume next

End Sub


'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows

		Next
	End With
End Sub


'=========================================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()

    C1_default_flag	=	1
	C1_name	=	2
	C1_email	=	3
	C1_dept_nm	=	4
	C1_ph_no	=	5
	C1_remarks	=	6
	C1_bp_cd	=	7
	C1_bp_seq	=	8


	

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
			
            C1_default_flag	=	iCurColumnPos(1)
			C1_name	=	iCurColumnPos(2)
			C1_email	=	iCurColumnPos(3)
			C1_dept_nm	=	iCurColumnPos(4)
			C1_ph_no	=	iCurColumnPos(5)
			C1_remarks	=	iCurColumnPos(6)
			C1_bp_cd	=	iCurColumnPos(7)
			C1_bp_seq	=	iCurColumnPos(8)


 	End Select
End Sub


'========================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")									'⊙: Lock  Suitable  Field
	
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call InitGridComboBox
	Call SetToolBar("11100001001011")										'⊙: 버튼 툴바 제어 
	Call SetSingleFocus

End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0  Then
	
		End If
    End With
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


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------   
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row
	End With
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

    If Not chkFieldByCell(frm1.BizPartnerID,"A",gPageNo) Then Exit Function

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
    Call SetToolBar("11000001001011")												'⊙: 버튼 툴바 제어 
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
	Dim IntRetCD, rowCnt
	
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
	Dim pvCnt
	
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
    	For pvCnt = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow -1
			.vspdData.Row	= pvCnt
			.vspdData.Col	= C1_bp_cd
			.vspdData.Text	= UCase(.txtHBpCd.value)
		Next 
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

	DbQuery = False
	
    If Trim(frm1.BizPartnerID.value) = "" Then 
        frm1.txtUsrNm1.value = ""
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & "&BizPartnerID=" & Trim(frm1.BizPartnerID.value)
	
	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)																'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()		

    lgIntFlgMode = Parent.OPMD_UMODE

	DbQueryOk = false												'☆: 조회 성공후 실행로직 
	
	'-----------------------
	'Reset variables area
	'-----------------------
    Call InitData
	lgIntFlgMode = Parent.OPMD_UMODE			
	Call SetToolBar("11101111001111")		'⊙: 버튼 툴바 제어 
	
	Call ggoOper.LockField(Document, "Q")	
	
    frm1.txtHBpCd.value = frm1.BizPartnerID.value

    Set gActiveElement = document.activeElement
	DbQueryOk = true
End Function

Function DBQueryNotOk()		
	Call SetToolBar("11101111001111")		'⊙: 버튼 툴바 제어 
	frm1.txtHBpCd.value = frm1.BizPartnerID.value
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
        Dim checkCount
        checkCount = 0
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C1_default_flag

            If .vspdData.text = "1" Then
                checkCount = checkCount + 1
            End If 

            If checkCount > 1 Then 
            	Call LayerShowHide(0)
                IntRetCD = DisplayMsgBox("205932", "X", "주담당자", "X") 
                Exit Function
            End If

			.vspdData.Col = 0
			
			
			Select Case .vspdData.Text

				Case iInsertFlag

					strVal = "C" & iColSep _
							     & GetSpreadText(.vspdData,C1_bp_cd       ,lRow,"X","X") & iColSep _
							     & GetSpreadText(.vspdData,C1_bp_seq    ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_name     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_email     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_dept_nm     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_ph_no     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_default_flag     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_remarks     ,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1

				Case iUpdateFlag

					strVal = "U" & iColSep _
							     & GetSpreadText(.vspdData,C1_bp_cd       ,lRow,"X","X") & iColSep _
							     & GetSpreadText(.vspdData,C1_bp_seq    ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_name     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_email     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_dept_nm     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_ph_no     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_default_flag     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_remarks     ,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)					
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1
					
				Case iDeleteFlag

					strDel = "D" & iColSep _
							     & GetSpreadText(.vspdData,C1_bp_cd       ,lRow,"X","X") & iColSep _
							     & GetSpreadText(.vspdData,C1_bp_seq    ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_name     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_email     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_dept_nm     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_ph_no     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_default_flag     ,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C1_remarks     ,lRow,"X","X") & iColSep _
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


'=========================================================================================================
'    Name : OpenUsrId()
'    Description : User PopUp
'=========================================================================================================
Function OpenUsrId(Byval strCode ,Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
  
    arrParam(0) = "거래처 팝업"                                             ' 팝업 명칭 
    arrParam(1) = "B_BIZ_PARTNER (nolock)"                                  ' TABLE 명칭 
    arrParam(2) = strCode                                                   ' Code Condition
    arrParam(3) = ""                                                        ' Name Cindition
    arrParam(4) = ""                                                        ' Where Condition
    arrParam(5) = "거래처"            
    
    arrField(0) = "bp_cd"                                                  ' Field명(0)
    arrField(1) = "bp_nm"                                                  ' Field명(1)
    
    arrHeader(0) = "거래처"                                            ' Header명(0)
    arrHeader(1) = "거래처명"                                              ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUsrId(arrRet,iWhere)        'return value setting
    End If    
	frm1.BizPartnerID.focus
	Set gActiveElement = document.activeElement

End Function

'=========================================================================================================
'    Name : SetUsrId()
'    Description : User Master Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetUsrId(Byval arrRet,ByVal iWhere)

    With frm1
    
       if iWhere = 0 Then
            .BizPartnerID.value = arrRet(0)
            .txtUsrNm1.value = arrRet(1)
       else
			.vspdData.Col  = C1_bp_cd
			.vspdData.Text = arrRet(0)
			.vspdData.Col  = C1_bp_nm
			.vspdData.Text = arrRet(1)
       
       end if     
    End With
    
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>거래처담당자관리</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
									 </TR>
								</TABLE>
							</TD>
                            <TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
							<TD WIDTH=10>&nbsp;</TD>
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
 											<TD CLASS="TD5" NOWRAP>거래처</TD>
											<TD CLASS="TD656" NOWRAP colspan = 3>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="BizPartnerID" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUsrId frm1.BizPartnerID.value,0">
												<INPUT TYPE=TEXT AlT="거래처" ID="txtUsrNm1" NAME="txtUsrNm1" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
 											
										</TR>
										
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
						</TR>
						<TR>
							<TD WIDTH=100% HEIGHT=* valign=top>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR HEIGHT="100%">
										<TD  WIDTH="100%" colspan=4><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData ID = "A" width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
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
			<TR HEIGHT="20">
				<TD WIDTH="100%" >
  					<TABLE <%=LR_SPACE_TYPE_30%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD WIDTH=*>&nbsp;</TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
  					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
        <TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex=-1></TEXTAREA>
		<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserDN" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserInfo" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtbtnFlag" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtChangeStatus" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtHBpCd" tag="24" TABINDEX="-1">
	</FORM>
	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=280 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
</BODY>
</HTML>
