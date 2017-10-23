<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Basic Info. - Accounting
'*  2. Function Name        : Common Info.
'*  3. Program ID           : B1320MA1
'*  4. Program Name         : 은행한도정보등록 
'*  5. Program Desc         : Register of Bank Loan Limit
'*  6. Component List       : PB2SA10
'*  7. Modified date(First) : 2000/03/22
'*  8. Modified date(Last)  : 2001/02/26
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Song, Mun Gil / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit 

			'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->			

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "b1320mb1.asp"												'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_BANK_LOAN_TYPE_NM
Dim C_BANK_LOAN_TYPE_CD
Dim C_Currency 
Dim C_CurrencyPopup 
Dim C_BANK_LOAN_LIMIT
Dim C_BANK_LOAN_USE
Dim C_BANK_LOAN_LOAN 



'Coding 부분 


Dim lgStrLoanPrevKey
Dim lgStrCurrencyPrevKey
Dim ggAmtofMoneyNo

Dim IsOpenPop          
Dim lgRetFlag

'========================================================================================================
Sub initSpreadPosVariables()         '1.2 변수에 Constants 값을 할당 
	C_BANK_LOAN_TYPE_NM		= 1
    C_BANK_LOAN_TYPE_CD		= 2
	C_Currency				= 3
	C_CurrencyPopup			= 4
	C_BANK_LOAN_LIMIT		= 5
	C_BANK_LOAN_USE			= 6
	C_BANK_LOAN_LOAN		= 7


	
End Sub

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrLoanPrevKey = ""
    lgStrCurrencyPrevKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
    lgSortKey = 1
    
    lgStrPrevKey = ""						    'initializes Previous Key
    lgPageNo     = "0"

End Sub


'========================================================================================================= 
Sub SetDefaultVal()
End Sub

'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()

'dim ggAmtofMoneyNo
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData   
	ggoSpread.Spreadinit "V20021106",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 

    With frm1.vspdData	
		.MaxCols = C_BANK_LOAN_LOAN + 1
		.MaxRows = 0
		.ReDraw = False 
	
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCombo C_BANK_LOAN_TYPE_NM, "은행한도유형", 18
		ggoSpread.SSSetCombo C_BANK_LOAN_TYPE_CD, "은행한도유형", 18
		ggoSpread.SSSetEdit C_Currency, "통화", 13,,,3,2
		ggoSpread.SSSetButton C_CurrencyPopup
		ggoSpread.SSSetFloat C_BANK_LOAN_USE, "사용금액", 28, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_BANK_LOAN_LIMIT,"한도금액", 28, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_BANK_LOAN_LOAN, "무역금융금액", 28, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_BANK_LOAN_TYPE_CD,C_BANK_LOAN_TYPE_CD,True)
		call ggoSpread.MakePairsColumn(C_Currency,C_CurrencyPopup)

        .ReDraw = true

    End With	
    Call SetSpreadLock    
End Sub

'================================== 2.2.4 SetSpreadLock() =============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_BANK_LOAN_TYPE_NM,-1,C_BANK_LOAN_TYPE_NM
    ggoSpread.SpreadLock C_Currency,-1,C_Currency
    ggoSpread.SpreadLock C_CurrencyPopup, -1, C_CurrencyPopup
    ggoSpread.SpreadLock C_BANK_LOAN_USE,-1,C_BANK_LOAN_USE
    ggoSpread.SpreadLock C_BANK_LOAN_LOAN,-1,C_BANK_LOAN_LOAN
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    
    .vspdData.ReDraw = True

    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1    
       .vspdData.ReDraw = False
                               
    ggoSpread.SSSetRequired	C_BANK_LOAN_TYPE_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_Currency, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BANK_LOAN_USE,pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BANK_LOAN_LOAN,pvStartRow, pvEndRow
    
       .vspdData.ReDraw = True
    End With
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1
		
	On error resume next

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("B9015", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	If intRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_BANK_LOAN_TYPE_CD		
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_BANK_LOAN_TYPE_NM
	End If		
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

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
			
			.Col = C_BANK_LOAN_TYPE_CD
			intIndex = .value
			.col = C_BANK_LOAN_TYPE_NM
			.value = intindex
					
		Next	
	End With
End Sub

'========================================================================================================= 
Function OpenPopup(strCode, iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	Case 0
		arrParam(0) = "은행 팝업"						' 팝업 명칭 
		arrParam(1) = "B_BANK"									' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtbank_cd.Value)				' Code Condition
		arrParam(3) = ""										' Name Cindition
		arrParam(4) = ""										' Where Condition
		arrParam(5) = "은행"							' TextBox 명칭 
	
	    arrField(0) = "BANK_CD"									' Field명(0)
		arrField(1) = "BANK_NM"									' Field명(1)
    
	    arrHeader(0) = "은행코드"							' Header명(0)
		arrHeader(1) = "은행명"						' Header명(1)
    
	Case 1
		arrParam(0) = "통화 팝업"						
		arrParam(1) = "b_currency"						
		arrParam(2) = strCode
		arrParam(3) = ""								
		arrParam(4) = ""								
		arrParam(5) = "통화"							
			
		arrField(0) = "currency"						
		arrField(1) = "currency_desc"					
		    
		arrHeader(0) = "통화"							
		arrHeader(1) = "비고"							

	Case Else
		Exit Function
	End Select
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
		Case 0
			frm1.txtbank_cd.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


 '------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)'
	With frm1
	Select Case iWhere
		Case 0
			.txtbank_cd.focus
			.txtbank_cd.value = arrRet(0)
			.txtbank_nm.value = arrRet(1)
		Case 1
			.vspdData.Col  = C_Currency
			.vspdData.Text = arrRet(0)
			Call vspdData_Change(.vspdData.Col, .vspdData.Row)
		End Select
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
			C_BANK_LOAN_TYPE_NM          = iCurColumnPos(1)
			C_BANK_LOAN_TYPE_CD          = iCurColumnPos(2)
			C_Currency       = iCurColumnPos(3)    
			C_CurrencyPopup        = iCurColumnPos(4)
			C_BANK_LOAN_LIMIT      = iCurColumnPos(5)
			C_BANK_LOAN_USE = iCurColumnPos(6)
			C_BANK_LOAN_LOAN    = iCurColumnPos(7)
			
    End Select    
End Sub

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029 
    'call initSpreadPosVariables() 
    Call ggoOper.LockField(Document, "N")
	'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitSpreadSheet
    Call InitVariables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("110011010010111")
    frm1.txtbank_cd.focus

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
	
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

    
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
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

'    If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
'      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
'         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
'      End If
'    End If
    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_CURRENCYPOPUP Then
		    .Col = Col
		    .Row = Row

			.Col = C_CURRENCY
		    Call OpenPopup(.Text, 1)
			Call SetActiveCell(frm1.vspdData,Col - 1,frm1.vspdData.ActiveRow ,"M","X","X")
		End If
    End With
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
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '-- 이전 Start
    '----------  Coding part  -------------------------------------------------------------   
    'if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크 
	'	If lgStrLoanPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
	'		Call DbQuery
	'	End If
    'End if
    '-- 이전 End

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then	
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
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
    
		Select Case Col
		
			Case  C_BANK_LOAN_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_BANK_LOAN_TYPE_CD
				.Value = intIndex

		End Select

	End With

End Sub

'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables
    Call InitComboBox
    					
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function


'========================================================================================
Function FncNew() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
Function FncDelete() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False
    
    Err.Clear
    'On Error Resume Next
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If
    
    If Not chkField(Document, "1") Then
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
	'call initSpreadPosVariables()
	If frm1.vspdData.MaxRows < 1 Then Exit Function
		
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_Bank_Loan_Type_Nm
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_Currency
    frm1.vspdData.Text = ""
    
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
    Call InitData                                                     
End Function


'========================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	  Dim lRow, IntRetCD
	  Dim IntRetCD1,strYear,strMonth,strDay, strYYYYMM, strYYYYMMDD
    
    On Error Resume Next
    Err.Clear

    
	FncInsertRow = False
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
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imrow -1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	    
    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
    End With
    
End Function



'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
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

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	
	FncExit = False
    
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
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
    
    Call LayerShowHide(1)                                                   '☜: Protect system from crashing
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
			strVal = strVal & "&txtbank_cd=" & Trim(.hBankCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrLoanPrevKey=" & lgStrLoanPrevKey
			strVal = strVal & "&lgStrCurrencyPrevKey=" & lgStrCurrencyPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else    
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
			strVal = strVal & "&txtbank_cd=" & Trim(.txtbank_cd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrLoanPrevKey=" & lgStrLoanPrevKey
			strVal = strVal & "&lgStrCurrencyPrevKey=" & lgStrCurrencyPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End if

    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    lgIntFlgMode = Parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
    
	Call InitData
	
	Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field

	Call SetToolbar("110011110011111")										'버튼 툴바 제어 
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    Dim  pB13031     'As New P21011ManageIndReqSvr
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal,strDel
	
    DbSave = False                                                          '⊙: Processing is NG
    
    'On Error Resume Next                                                   '☜: Protect system from crashing
    
    Call LayerShowHide(1)                                                   '☜: Protect system from crashing
	
	lgRetFlag = False
	
	
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
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag													'☜: 신규 
				
				strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep  					'☜: C=Create, Row위치 정보 
                .vspdData.Col = C_BANK_LOAN_TYPE_CD						
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_Currency
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_BANK_LOAN_LIMIT
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
                .vspdData.Col = C_BANK_LOAN_USE
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
                .vspdData.Col = C_BANK_LOAN_LOAN
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep              
                    
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag													'☜: 수정 
					
				strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep  					'☜: U=Update, Row위치 정보 
                .vspdData.Col = C_BANK_LOAN_TYPE_CD						'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_Currency
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_BANK_LOAN_LIMIT
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
                .vspdData.Col = C_BANK_LOAN_USE
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
                .vspdData.Col = C_BANK_LOAN_LOAN
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep              

                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag											'☜: 삭제 

				strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep
                .vspdData.Col = C_BANK_LOAN_TYPE_CD						'1
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_Currency
                strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep									
                               
                lGrpCnt = lGrpCnt + 1
                
                lgRetFlag = True
        End Select
                
    Next
    	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 

	End With

    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitVariables
	Call Dbquery
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->

</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>은행한도정보등록</font></td>
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
					<TD HEIGHT="20" WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>은행</TD>
									<TD CLASS="TD656" COLSPAN=3><INPUT NAME="txtbank_cd" MAXLENGTH="10" SIZE=10 ALT ="은행코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBank_Cd.value,0)">&nbsp;
										 				   <INPUT NAME="txtbank_nm" MAXLENGTH="30" SIZE=30 STYLE="TEXT-ALIGN:left" ALT ="은행코드명" tag="24X"></TD>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/b1320ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hBankCd" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

