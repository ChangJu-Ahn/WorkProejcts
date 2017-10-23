<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name		: 
*  2. Function Name		: Multi Sample
*  3. Program ID		: VB000MA1.ASP
*  4. Program Name		: 
*  5. Program Desc		: 경영정보계획 
*  6. Comproxy List		:
*  7. Modified date(First)	: 2005/02/02
*  8. Modified date(Last)	: 
*  9. Modifier (First)	: Cho Ig Sung
* 10. Modifier (Last)	: 
* 11. Comment			: EIS
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=						4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "GB000MB1.asp"									'비지니스 로직 ASP명 
Const CookieSplit = 1233

Dim IsOpenPop  

Dim C_YYYYMM										'Spread Sheet의 Column별 상수 
Dim C_BS_PL_FLAG		
Dim C_BS_PL_POPUP
Dim C_BS_PL_NM
Dim C_SUMMARY
Dim C_PLAN_AMT		

'========================================================================================================
'=						4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=						4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
'========================================================================================================
' Name : InitSpreadPosVariables() 
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()
	C_YYYYMM		= 1		
	C_BS_PL_FLAG	= 2	
	C_BS_PL_POPUP	= 3	
	C_BS_PL_NM	= 4	
	C_SUMMARY		= 5	
	C_PLAN_AMT		= 6
End Sub
'========================================================================================================
' Name : InitVariables() 
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode	= Parent.OPMD_CMODE						'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False									'⊙: Indicates that no value changed
	lgIntGrpCount	= 0										'⊙: Initializes Group View Size
	lgStrPrevKey	= ""									'⊙: initializes Previous Key
	lgSortKey		= 1										'⊙: initializes sort direction
End Sub

Sub SetDefaultVal()
	Dim StartDate, ServerDate

	StartDate	= Parent.gFiscStart
	ServerDate	= UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,gDateFormat)
	
	frm1.txtFromDt.text		= StartDate
	frm1.txtToDt.text		= ServerDate
	
	Call ggoOper.FormatDate(frm1.txtFromDt, gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtToDt, gDateFormat, 2)

End Sub

'========================================================================================================
' Name : LoadInfTB19029() 
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub
'========================================================================================================
' Name : CookiePage()
' Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

	Dim strFromDt, strToDt

	strFromDt	= frm1.txtFromDt.year & right("0" & frm1.txtFromDt.month,2)
	strToDt		= frm1.txtToDt.year & right("0" & frm1.txtToDt.month,2)

	lgKeyStream	= Trim(strFromDt) &  Parent.gColSep 
	lgKeyStream	= lgKeyStream & Trim(strToDt) &  Parent.gColSep		'You Must append one character( Parent.gColSep)
End Sub		

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()	
End Sub

Sub InitData()
End Sub

'======================================================================================================
'	Event Name : vspdData_ComboSelChange
'	Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()	
	With frm1.vspdData
 
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread	

		.ReDraw = false

		.MaxCols = C_PLAN_AMT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

		.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		Call  GetSpreadColumnPos("A")	

		ggoSpread.SSSetEdit		C_YYYYMM,			"계획년월",			10,,, 10,2
		ggoSpread.SSSetEdit		C_BS_PL_FLAG,		"제조손익분류",		15,,, 10,2
		ggoSpread.SSSetButton	C_BS_PL_POPUP	
		ggoSpread.SSSetEdit		C_BS_PL_NM,			"제조손익분류명",	25, 0 
		ggoSpread.SSSetEdit		C_SUMMARY,			"적요",				40,,,255

		ggoSpread.SSSetFloat	C_PLAN_AMT,			"계획금액", 20, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"		

		Call ggoSpread.MakePairsColumn(C_BS_PL_FLAG,C_BS_PL_POPUP)
				
		.ReDraw = true		
		Call SetSpreadLock	
	
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	With frm1
		
		.vspdData.ReDraw = False

		ggoSpread.SpreadLock		C_YYYYMM,			-1, C_YYYYMM
		ggoSpread.SpreadLock		C_BS_PL_FLAG,		-1, C_BS_PL_FLAG
		ggoSpread.SpreadLock		C_BS_PL_POPUP,		-1, C_BS_PL_POPUP
		ggoSpread.SSSetProtected	C_BS_PL_NM,			-1,	C_BS_PL_NM
'		ggoSpread.SSSetRequired		C_SUMMARY,			-1, C_SUMMARY
		ggoSpread.SSSetRequired		C_PLAN_AMT,			-1, C_PLAN_AMT
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired		C_YYYYMM		, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_BS_PL_FLAG	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_BS_PL_NM		, pvStartRow, pvEndRow
'		ggoSpread.SSSetRequired		C_SUMMARY		, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_PLAN_AMT		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	.vspdData.MaxCols, pvStartRow, pvEndRow

		.vspdData.ReDraw = True
		
	End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
	Dim iDx
	Dim iRow
	iPosArr = Split(iPosArr, Parent.gColSep)
	If IsNumeric(iPosArr(0)) Then
		iRow = CInt(iPosArr(0))
		For iDx = 1 To  frm1.vspdData.MaxCols - 1
			Frm1.vspdData.Col = iDx
			Frm1.vspdData.Row = iRow
			If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  Parent.UC_PROTECTED Then
				Frm1.vspdData.Col = iDx
				Frm1.vspdData.Row = iRow
				Frm1.vspdData.Action = 0 ' go to 
				Exit For
			End If
		Next
	End If	
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description	: 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				
			C_YYYYMM		= iCurColumnPos(1)		
			C_BS_PL_FLAG	= iCurColumnPos(2)	
			C_BS_PL_POPUP	= iCurColumnPos(3)	
			C_BS_PL_NM	= iCurColumnPos(4)	
			C_SUMMARY		= iCurColumnPos(5)	
			C_PLAN_AMT		= iCurColumnPos(6)  
	End Select	
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	Err.Clear
	Call LoadInfTB19029															'⊙: Load table , B_numeric_format
		
	Call  ggoOper.FormatField(Document, "1",ggStrIntegeralPart,ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")			'⊙: Lock Field
			
	Call InitSpreadSheet															'Setup the Spread sheet
	Call InitVariables															'Initializes local global variables
	Call InitComboBox
	
	Call SetDefaultVal

	Call SetToolbar("1100111100101111")				'버튼 툴바 제어 

	Call CookiePage(0) 
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End If
End Sub

Sub txtFromDt_Keypress(Key)
	If Key = 13 Then
		FncQuery()
	End If
End Sub

Sub txtToDt_Keypress(Key)
	If Key = 13 Then
		FncQuery()
	End If
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
	Dim IntRetCD 
	FncQuery = False 
	Err.Clear

	ggoSpread.Source = frm1.vspdData
	If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  Parent.VB_YES_NO,"x","x")	'☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	ggoSpread.ClearSpreadData
				
	If Not chkField(Document, "1") Then
		Exit Function
	End If
  
	If UniConvDateToYYYYMMDD(frm1.txtFromDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToDt.text, Parent.gDateFormat,"")Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
	End If
	
	Call InitVariables
	Call MakeKeyStream("X")
	
	Call  DisableToolBar( Parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreTooBar()
		Exit Function
	End If
	FncQuery = True
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
	Dim IntRetCD 
	FncSave = False
	Err.Clear
	
	ggoSpread.Source = frm1.vspdData
	If  ggoSpread.SSCheckChange = False Then
		IntRetCD =  DisplayMsgBox("900001","x","x","x")							'☜:There is no changed data. 
		Exit Function
	End If
	
	ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then										'☜: Check contents area
		Exit Function
	End If

	Call  DisableToolBar(Parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
		Exit Function
	End If				
	FncSave = True
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

	If Frm1.vspdData.MaxRows < 1 Then
		Exit Function
	End If
		
	ggoSpread.Source = Frm1.vspdData

	With Frm1.VspdData
		.ReDraw = False
		If .ActiveRow > 0 Then
			ggoSpread.CopyRow
						
			SetSpreadColor .ActiveRow, .ActiveRow

			.Col = C_BS_PL_FLAG
			.Text = ""
												
			.Col = C_BS_PL_NM
			.Text = ""

'			.Col = C_SUMMARY
'			.Text = ""

			.ReDraw = True

			.Col = C_YYYYMM
			.Focus
			.Action = 0 ' go to 
		End If
	End With

	Set gActiveElement = document.ActiveElement	

End Function
'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
	ggoSpread.Source = Frm1.vspdData 
	ggoSpread.EditUndo
	Call InitData
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow, iRow
	
	On Error Resume Next
	Err.Clear
 
	FncInsertRow = False 

	If IsNumeric(Trim(pvRowCnt)) Then
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
		ggoSpread.InsertRow .vspdData.ActiveRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1		
				
		.vspdData.ReDraw = True
	End With

	If Err.number = 0 Then
		FncInsertRow = True
	End If	

	Set gActiveElement = document.ActiveElement	
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
	Dim lDelRows
	If Frm1.vspdData.MaxRows < 1 then
		Exit function
	End if 
 
	With Frm1.vspdData 
		.focus

		ggoSpread.Source = frm1.vspdData 
		lDelRows =  ggoSpread.DeleteRow
	End With
	Set gActiveElement = document.ActiveElement	
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
	Call parent.FncExport( Parent.C_MULTI)										'☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
	Call parent.FncFind( Parent.C_MULTI, False)									'☜:화면 유형, Tab 유무 
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
'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
 
	ggoSpread.Source = frm1.vspdData 
	If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  Parent.VB_YES_NO,"x","x")	'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description	: 
'========================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description	: 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()	
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal
	DbQuery = False
	Err.Clear

	if LayerShowHide(1) = false then
		exit Function
	end if
 
	With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="			&  Parent.UID_M0001				
		strVal = strVal	& "&txtKeyStream="		& lgKeyStream						'☜: Query Key
		strVal = strVal	& "&txtMaxRows="		& .vspdData.MaxRows
		strVal = strVal	& "&lgStrPrevKey="		& lgStrPrevKey				'☜: Next key tag
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)												'☜: Run Biz Logic
	DbQuery = True
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	Dim lRow		
	Dim lGrpCnt	
	Dim strVal, strDel
 
	DbSave = False
	
	if LayerShowHide(1) = false then
		exit Function
	end if

	strVal = ""
	strDel = ""
	lGrpCnt = 1

	With Frm1
		For lRow = 1 To .vspdData.MaxRows
	
			.vspdData.Row = lRow
			.vspdData.Col = 0
		
			Select Case .vspdData.Text
		
				Case  ggoSpread.InsertFlag									'☜: Insert
													strVal = strVal & "C"  &  Parent.gColSep
													strVal = strVal & lRow &  Parent.gColSep
										
					.vspdData.Col = C_YYYYMM		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
					.vspdData.Col = C_BS_PL_FLAG	: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
					.vspdData.Col = C_SUMMARY		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
					.vspdData.Col = C_PLAN_AMT		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
					
				Case  ggoSpread.UpdateFlag									'☜: Update
													strVal = strVal & "U"  &  Parent.gColSep
													strVal = strVal & lRow &  Parent.gColSep
										
					.vspdData.Col = C_YYYYMM		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
					.vspdData.Col = C_BS_PL_FLAG	: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
					.vspdData.Col = C_SUMMARY		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
					.vspdData.Col = C_PLAN_AMT		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
				Case  ggoSpread.DeleteFlag									'☜: Delete
													strDel = strDel & "D"  &  Parent.gColSep
													strDel = strDel & lRow &  Parent.gColSep

					.vspdData.Col = C_YYYYMM		: strDel = strDel & Trim(.vspdData.Text) &  Parent.gColSep
					.vspdData.Col = C_BS_PL_FLAG	: strDel = strDel & Trim(.vspdData.Text) &  Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
			End Select
		Next
		.txtMode.value		=  Parent.UID_M0002
		.txtMaxRows.value	= lGrpCnt-1 
		.txtSpread.value	= strDel & strVal
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	DbSave = True															
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim IntRetCd
	FncDelete = False
	If lgIntFlgMode <>  Parent.OPMD_UMODE Then									'Check if there is retrived data
		Call  DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If
	
	IntRetCD =  DisplayMsgBox("900003",  Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function 
	End If
	
	Call  DisableToolBar( Parent.TBC_DELETE)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If			
	
	FncDelete = True
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()				
	lgIntFlgMode =  Parent.OPMD_UMODE	
	Call  ggoOper.LockField(Document, "Q")		'⊙: Lock field
	Call SetToolbar("110011110011111")
	Call InitData
	frm1.vspdData.focus								
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
		
	Call InitVariables				'⊙: Initializes local global variables
	Call  DisableToolBar( Parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If				'☜: Query db data
End Function


'========================================================================================================
' Name : OpenPopUp()		
' Desc : developer describe this line 
'========================================================================================================
Function OpenPopUp(ByVal strCode, Byval iWhere)

Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	Select Case iWhere
		Case 0
			arrParam(0) = "제조손익분류 팝업"		' 팝업 명칭 
			arrParam(1) = "B_MINOR"						' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "MAJOR_CD = 'A1023'"			' Where Condition
			arrParam(5) = "Minor Code"				' 조건필드의 라벨 명칭 

			arrField(0) = "MINOR_CD"				' Field명(0)
			arrField(1) = "MINOR_NM"				' Field명(1)
 
			arrHeader(0) = "Minor Code"				' Header명(0)
			arrHeader(1) = "Minor Code명"				' Header명(1)

	
	End Select
		
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'======================================================================================================
' Name : SetPopUp()			
' Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.vspdData.Col = C_BS_PL_FLAG
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_BS_PL_NM
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
		End Select
	End With
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
 
		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			Select Case Col
				Case C_BS_PL_POPUP
					.Col = Col - 1
					.Row = Row
					Call OpenPopUp (.text, 0)
			End Select
		
			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")  
		End If

	End With
End Sub

'========================================================================================================
'	Event Name : vspdData_Click
'	Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
	gMouseClickStatus = "SPC" 
	Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row
	
End Sub

'========================================================================================================
'	Event Name : vspdData_Change 
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim iDx
	Dim IntRetCD
	Dim intIndex 
		
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	Select Case Col
		Case  C_BS_PL_FLAG
			iDx = Frm1.vspdData.value
				Frm1.vspdData.Col = C_BS_PL_FLAG
	
			If Frm1.vspdData.value = "" Then
					Frm1.vspdData.Col	= C_BS_PL_NM
					Frm1.vspdData.value	= "" 
			Else
				IntRetCd =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = 'A1023' AND MINOR_CD = '" & trim(Frm1.vspdData.value) & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				
				If IntRetCd = false then
					Call  DisplayMsgBox("122300","X",trim(Frm1.vspdData.value),"X")	'%1 Minor코드정보가 없습니다.
					
					Frm1.vspdData.Col = C_BS_PL_NM
					Frm1.vspdData.value = ""
					Frm1.vspdData.Col = C_BS_PL_FLAG
					Frm1.vspdData.focus
				Else
						Frm1.vspdData.Col = C_BS_PL_NM
						Frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
				End if 
			End if  

	End Select	

	If Frm1.vspdData.CellType =  Parent.SS_CELL_TYPE_FLOAT Then
		If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
			Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
		End If
	End If
 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub
'========================================================================================================
'	Event Name : vspdData_DblClick 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
	if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'	Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'	Event Name : vspdData_MouseDown
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
		If Button = 2 And  gMouseClickStatus = "SPC" Then
			gMouseClickStatus = "SPCR"
		End If
End Sub	

'========================================================================================================
'	Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)	
		Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
'	Event Name : vspdData_TopLeftChange
'	Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
		If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
	
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
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
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>경영정보계획</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
								</TR>
							</TABLE>
						</TD>
						<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
  
		<TR HEIGHT=*>
			<TD width=100% CLASS="Tab11">
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS="TD5" NOWRAP>작업년월</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/gb000ma1_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
																<script language =javascript src='./js/gb000ma1_fpDateTime2_txtToDt.js'></script></TD>
										<TD CLASS="TD5" NOWRAP></TD>
										<TD CLASS="TD6" NOWRAP></TD>
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>

					<TR>
						<TD <%=HEIGHT_TYPE_03%>></TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_20%> >
								<TR>
									<TD HEIGHT=100% WIDTH=100% >
										<script language =javascript src='./js/gb000ma1_vaSpread_vspdData.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
	
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
		</TR>

	</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"		TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"	TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"	TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"	TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

