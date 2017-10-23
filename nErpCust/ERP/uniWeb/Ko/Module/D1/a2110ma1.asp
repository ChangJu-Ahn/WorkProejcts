<%@ LANGUAGE="VBSCRIPT" %>
<!--*******************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 기준 
'*  3. Program ID           : A1103ma1,A1103mb1
'*  4. Program Name         : 회계자동기표환경설정 
'*  5. Program Desc         : 
'*  6. Component List       : +B11011 (Manage)
'                             +B11018 (조회용)
'*  7. Modified date(First) : 2000/03/22
'*  8. Modified date(Last)  : 2004/12/09
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "a2110mb1.asp"								'☆: 비지니스 로직 ASP명 

Dim C_MO_CD														'☆: Spread Sheet의 Column별 상수 
Dim C_MO_NM
Dim C_GL_POSTING_FG
Dim C_GL_POSTING_FG_NM  
Dim C_BATCH_FG
Dim C_BATCH_FG_NM
Dim C_INV_POST_FG


'========================================================================================================
Sub initSpreadPosVariables()  
    C_MO_CD             = 1										'☆: Spread Sheet의 Column별 상수 
    C_MO_NM             = 2 
    C_GL_POSTING_FG     = 3
    C_GL_POSTING_FG_NM  = 4
    C_BATCH_FG          = 5
    C_BATCH_FG_NM       = 6
    C_INV_POST_FG       = 7 
End Sub

Dim igBlnFlgChgValue															' Variable is for Dirty flag
Dim igIntFlgMode																' Variable is for Operation Status

Dim igStrNextKey
Dim igPageNo

Dim isOpenPop
Dim igSortKey


'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE											'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False													'Indicates that no value changed
    lgIntGrpCount = 0															'initializes Group View Size
    
    lgStrPrevKey = ""
    lgSortKey = 1
End Sub

'********************************************************************************************************* 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()

End Sub

'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData   
	ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

    With frm1.vspdData
        .MaxCols = C_INV_POST_FG + 1
        .MaxRows = 0

        .ReDraw = False 

		Call GetSpreadColumnPos("A")

        ggoSpread.SSSetcombo C_MO_CD           ,"업무구분코드"    , 21
        ggoSpread.SSSetcombo C_MO_NM           ,"업무구분명"      , 35
        ggoSpread.SSSetCombo C_BATCH_FG        ,""                    ,   7
        ggoSpread.SSSetCombo C_BATCH_FG_NM     ,"일괄처리여부"    , 25
        ggoSpread.SSSetCombo C_GL_POSTING_FG   ,""                    , 7
        ggoSpread.SSSetCombo C_GL_POSTING_FG_NM,"회계전표처리여부", 25
		ggoSpread.SSSetEdit	 C_INV_POST_FG     ,""                    ,	2

		Call ggoSpread.MakePairsColumn(C_MO_CD,C_MO_NM,"1")
		Call ggoSpread.MakePairsColumn(C_BATCH_FG,C_BATCH_FG_NM,"1")
		Call ggoSpread.MakePairsColumn(C_GL_POSTING_FG,C_GL_POSTING_FG_NM,"1")

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_BATCH_FG,C_BATCH_FG,True)
		Call ggoSpread.SSSetColHidden(C_GL_POSTING_FG,C_GL_POSTING_FG,True)
		Call ggoSpread.SSSetColHidden(C_INV_POST_FG,C_INV_POST_FG,True)		

        .ReDraw = True
    End With

    Call SetSpreadLock 
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
 Sub SetSpreadLock()
    With frm1    
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock         C_MO_CD            , -1 , C_MO_CD
        ggoSpread.SpreadLock         C_MO_NM            , -1 , C_MO_NM 
        ggoSpread.SSSetRequired      C_BATCH_FG_NM      , -1 , C_BATCH_FG_NM
        ggoSpread.SSSetRequired      C_GL_POSTING_FG_NM , -1 , C_GL_POSTING_FG_NM
        ggoSpread.SSSetProtected	 C_INV_POST_FG		, -1 , -1        
        ggoSpread.SSSetProtected	.vspdData.MaxCols	, -1 , -1
        .vspdData.ReDraw = True
    End With    
End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired     C_MO_CD                 , pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_MO_NM                 , pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_BATCH_FG_NM           , pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_GL_POSTING_FG_NM      , pvStartRow, pvEndRow
        .vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor_AfterQry()
	Dim ii

    With frm1
		.vspddata.ReDraw = False
		For ii = 1 To .vspddata.Maxrows
			.vspddata.Row = ii
			.vspddata.col = C_INV_POST_FG
			If Trim(.vspddata.Text) <> "M" Then
				ggoSpread.SSSetProtected C_BATCH_FG_NM ,ii,ii								
			End If
		Next
        .vspdData.ReDraw = True		
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
' Name : InitComboBox()
' Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim IntRetCD1
	Dim IntRetCD2
	Dim IntRetCD3
  
	On Error Resume Next

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1026", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
  
	If intRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_BATCH_FG 
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_BATCH_FG_NM
	End If  
 
	IntRetCD2= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1027", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
	If intRetCD2 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_GL_POSTING_FG
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_GL_POSTING_FG_NM
	End If  
 
	IntRetCD3= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("B0001", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
	If intRetCD3 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_MO_CD
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_MO_NM
	End If  
End Sub

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
' Name : OpenItemInfo()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode, Byval iWhere)'
	Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    'IsOpenPop = True

    arrParam(0) = "업무구분 팝업"										' 팝업 명칭 
    arrParam(1) = "B_MINOR"													' TABLE 명칭 
    arrParam(2) = strCode													' Code Condition
    arrParam(3) = ""														' Name Cindition
    arrParam(4) = "MAJOR_CD = " & FilterVar("B0001", "''", "S") & " "		' Where Condition
    arrParam(5) = "업무구분"   

    arrField(0) = "MINOR_CD"												' Field명(0)
    arrField(1) = "MINOR_NM"												' Field명(1)

    arrHeader(0) = "업무구분코드"										' Header명(0)
    arrHeader(1) = "업무구분명"											' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
 
    If arrRet(0) = "" Then
        frm1.txtMO_CD.focus
        Exit Function
    Else
        Call SetItemInfo(arrRet, iWhere)
    End If 
End Function

 '------------------------------------------  SetItemInfo()  --------------------------------------------------
' Name : SetItemInfo()
' Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet, Byval iWhere)'
    With frm1
        If iWhere = 0 Then
        .txtMO_CD.focus
        .txtMO_CD.value = arrRet(0)
        .txtMO_NM.value = arrRet(1)
        End if 
    End With
End Function

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

            C_MO_CD             = iCurColumnPos(1)
            C_MO_NM             = iCurColumnPos(2)
            C_GL_POSTING_FG     = iCurColumnPos(3)
            C_GL_POSTING_FG_NM  = iCurColumnPos(4)
            C_BATCH_FG          = iCurColumnPos(5)
            C_BATCH_FG_NM       = iCurColumnPos(6)
            C_INV_POST_FG       = iCurColumnPos(7)            
    End Select    
End Sub

'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029() 
    
    Call ggoOper.LockField(Document, "N")															'⊙: Load table , B_numeric_format
    'Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)       

    Call InitSpreadSheet																			'⊙: Setup the Spread sheet
    Call InitVariables																				'⊙: Initializes local global variables
    Call initcombobox																				' 업무구분  Combo

    Call SetDefaultVal
    Call SetToolbar("110010010000111")																'⊙: 버튼 툴바 제어 

    frm1.txtMO_CD.focus   
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub


'******************************  3.2.1 Object Tag 처리  *********************************************
' Window에 발생 하는 모든 Even 처리 
'********************************************************************************************************* 
Sub vspdData_Click(Col, Row)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then
		Exit Sub
   	End If

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
        Exit Sub
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
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
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    If OldLeft <> NewLeft Then
		Exit Sub
    End If

	If CheckRunningBizProcess = True Then
		Exit Sub
	End If

    If frm1.vspdData.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData,NewTop) Then '☜: 재쿼리 체크 
        If lgStrPrevKey <> "" Then       '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
            DbQuery
        End If
    End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")          '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call InitVariables

    If Trim(frm1.txtMO_CD.value) = "" Then
        frm1.txtMO_NM.value = ""
    End If

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         '⊙: This function check indispensable field
		Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery

    FncQuery = True
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 

    FncSave = False

    Err.Clear
    On Error Resume Next

    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
		Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave

    FncSave = True
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
    Dim  IntRetCD

    On Error Resume Next
    Err.Clear

    FncCopy = False

    If Frm1.vspdData.MaxRows < 1 Then
		Exit Function
    End If

    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
        .ReDraw = False
		If .ActiveRow > 0 Then
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow

			.ReDraw = True
			.Focus
		End If
	End With
	
	With Frm1
        .vspdData.Col  = C_MO_CD
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""

        .vspdData.Col  = C_MO_NM
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
	End With

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    Dim iDx

    On Error Resume Next
    Err.Clear

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    FncCancel = False

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.EditUndo  

    Call InitData()

    If Err.number = 0 Then	
		FncCancel = True
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow

    On Error Resume Next
    Err.Clear

    FncInsertRow = False

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
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
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
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

    On Error Resume Next
    Err.Clear

    FncDeleteRow = False 

    If Frm1.vspdData.MaxRows < 1 then
		Exit function
	End if

    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With

    If Err.number = 0 Then
		FncDeleteRow = True
    End If

    Set gActiveElement = document.ActiveElement
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
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call Parent.FncExport(Parent.C_MULTI)            '☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, True)                                         '☜:화면 유형, Tab 유무 
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
	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next
    Err.Clear

    FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		              '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then
       FncExit = True
    End If

    Set gActiveElement = document.ActiveElement
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
    Call LayerShowHide(1)

    Err.Clear                                                               '☜: Protect system from crashing

    Dim strVal

    With frm1
        If lgIntFlgMode = Parent.OPMD_UMODE Then
          strVal = BIZ_PGM_ID   & "?txtMode="     & Parent.UID_M0001		'☜: 
          strVal = strVal       & "&txtMO_CD="    & Trim(.hMO_CD.value)		'☆: 조회 조건 데이타 
          strVal = strVal       & "&iStrNextKey=" & igStrNextKey
          strVal = strVal       & "&iPageNo="     & igPageNo
        Else
          strVal = BIZ_PGM_ID   & "?txtMode="     & Parent.UID_M0001		'☜: 
          strVal = strVal       & "&txtMO_CD="    & Trim(.txtMO_CD.value)	'☆: 조회 조건 데이타 
          strVal = strVal       & "&iStrNextKey=" & igStrNextKey
          strVal = strVal       & "&iPageNo="     & igPageNo
        End If
    End With
    
    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    lgIntFlgMode = Parent.OPMD_UMODE 

    Call ggoOper.LockField(Document, "Q")
    Call InitData
    Call SetToolbar("110010010001111")
    Call SetSpreadColor_AfterQry

	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim pB21011     'As New P21011ManageIndReqSvr
    Dim IRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
    Dim strVal,strDel

    DbSave = False                                                          '⊙: Processing is NG
    Call LayerShowHide(1)                                                   '☜: Protect system from crashing
    On Error Resume Next                                                   '☜: Protect system from crashing

    With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1

		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For IRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = IRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag           '☜: 신규 
					strVal = strVal & "C" & Parent.gColSep & IRow & Parent.gColSep			'☜: C=Create, Row위치 정보 
		            .vspdData.Col = C_MO_CD           
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep 
		            .vspdData.Col = C_BATCH_FG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_GL_POSTING_FG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.UpdateFlag           '☜: 수정 
					strVal = strVal & "U" & Parent.gColSep & IRow & Parent.gColSep			'☜: U=Update, Row위치 정보 
		            .vspdData.Col = C_MO_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col =  C_BATCH_FG          '10
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col =  C_GL_POSTING_FG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag           '☜: 삭제 

					strDel = strDel & "D" & Parent.gColSep & IRow & Parent.gColSep
		            .vspdData.Col = C_MO_CD
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col =  C_BATCH_FG          '10
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col =  C_GL_POSTING_FG
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
 
		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = strDel & strVal
 
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)													'☜: 비지니스 ASP 를 가동 
	End With
 
    DbSave = True																			'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()																			'☆: 저장 성공후 실행 로직 
    On Error Resume Next																	'☜: If process fails
    Err.Clear																				'☜: Clear error status

    Call InitVariables
    Call ggoOper.ClearField(Document, "2")													'☜: Clear Contents  Field    

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	Set gActiveElement = document.ActiveElement

    Call Dbquery
End Function

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData()
    Dim intRow
    Dim intIndex 

    With frm1.vspdData
		For intRow = 1 To .MaxRows
		    .Row = intRow
		    .col = C_MO_CD               :   intIndex    = .value
		    .col = C_MO_NM               :   .value      = intindex
		    .col = C_Batch_Fg            :   intIndex    = .value
		    .col = C_Batch_Fg_NM         :   .value      = intindex
		    .Col = C_Gl_Posting_Fg       :   intIndex    = .value
		    .col = C_Gl_Posting_Fg_NM    :   .value      = intindex
		Next 
    End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'======================================================================================================= 
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex

    With frm1.vspdData
        .Row = Row

        Select Case Col
			Case  C_MO_CD
			    .Col = Col
			    intIndex = .Value
			    .Col = C_MO_NM
			    .Value = intIndex
			Case C_GL_POSTING_FG 
			    .Col = Col
			    intIndex = .Value
			    .Col = C_GL_POSTING_FG_NM
			    .Value = intIndex
			Case C_GL_POSTING_FG_NM
			    .Col = Col
			    intIndex = .Value
			    .Col = C_GL_POSTING_FG
			    .Value = intIndex
			Case C_BATCH_FG
			    .Col = Col
			    intIndex = .Value
			    .Col = C_BATCH_FG_NM
			    .Value = intIndex
			Case C_BATCH_FG_NM
			    .Col = Col
			    intIndex = .Value
			    .Col = C_BATCH_FG
			    .Value = intIndex
        End Select
    End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->

</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>회계자동기표환경설정</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
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
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>업무구분</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtMO_CD" MAXLENGTH="2" SIZE=15 ALT ="업무구분"   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenItemInfo(txtMO_CD.value,0)">&nbsp;
															 <INPUT NAME="txtMO_NM" MAXLENGTH="30" SIZE=30 ALT ="업무구분명" tag="14">
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/a2110ma1_I262107080_vspdData.js'></script>
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
	<!--<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>-->
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"   tag="24"    tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"           tag="24"    tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"        tag="24"    tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hMO_CD"            tag="24"    tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

