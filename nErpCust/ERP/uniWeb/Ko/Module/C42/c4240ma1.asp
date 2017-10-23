<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 회계가공비집계조회 
'*  3. Program ID           : c4214ma1.asp
'*  4. Program Name         : 회계가공비집계조회 
'*  5. Program Desc         : 회계가공비집계조회 
'*  6. Modified date(First) : 2005-08-30
'*  7. Modified date(Last)  : 2005-08-30
'*  8. Modifier (First)     : choe0tae 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4240mb1.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt
Dim lgStrPrevKey2
Dim lgRow, lgEOF1, lgEOF2

iDBSYSDate = "<%=GetSvrDate%>"
iStrFromDt = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)	

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol

Dim lgSTime		' -- 디버깅 타임체크 
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
		
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    
    lgStrPrevKey = ""	
    lgRow = 0
    lgStrPrevKey2 = ""	
    lgEOF1 = False
    lgEOF2 = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtSTART_DT.Text = Left(iStrFromDt, 7)
	frm1.txtEND_DT.Text = Left(iStrFromDt, 7)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("Q","C", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(Byval pMaxCols)
	Dim i, ret
	With frm1.vspdData

		.style.display = "none"
		.Redraw = False

		.MaxRows = 0
		.MaxCols = pMaxCols

		If frm1.rdoTYPE1.checked Then
			.ColHeaderRows = 4
		Else
			.ColHeaderRows = 5
		End If

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021122",, "" ',,parent.gAllowDragDropSpread 

		ggoSpread.SSSetEdit		1,	"원가요소" & vbCrLf & "Type"	, 4,,,,1	
		ggoSpread.SSSetEdit		2,	"원가요소"	, 5,,,,1	
		ggoSpread.SSSetEdit		3,	"원가요소명"	, 8	
		ggoSpread.SSSetEdit		4,	"계정"	, 10,,,,1	
		ggoSpread.SSSetEdit		5,	"계정명"	, 20
		ggoSpread.SSSetEdit		6,	"집계원천"	, 8,,,,1

		For i = 7 To pMaxCols				
			ggoSpread.SSSetFloat	i,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			.ColHidden = False
		Next
		
		.Col = pMaxCols
		.ColHidden = True
		
		ggoSpread.SSSetSplit2(6)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()

		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 
		.rowheight(-998) = 12	' 높이 재지정 
		.rowheight(-997) = 12	' 높이 재지정 
		
		.Redraw = True
	End With
End Sub
		

Sub SetGridHead(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol

	' -- 그리드 1 정의 
	With frm1.vspdData
		
		arrRows = Split(pData, Parent.gRowSep)

		'헤더를 ?줄로    
		'.ColHeaderRows = UBound(arrRows, 1)
		
		For i = 0 To UBound(arrRows, 1) -1
			arrCols = Split(arrRows(i), Parent.gColSep)
			iColCnt = UBound(arrCols, 1)
			.Row	= CDbl(arrCols(iColCnt))		' -- 마지작 컬럼에 행번호가 들어있다.
			iCol = 7
			For j = 0 To iColCnt 
				.Col = iCol
				Select Case j
					Case 0, 1, 2, 3, 4, 5, iColCnt
						.Text = arrCols(j)
						iCol = iCol + 1
					Case Else
						.Text = arrCols(j)
						 iCol = iCol + 1	: .Col = iCol	' -- 금액 
				End SElect
				
			Next
		Next
	End With
End Sub

Sub ReInitSpreadSheet()
	
	Dim ret, iRowSpan
	' -- 그리드 1 정의 
	With frm1.vspdData

		'.MaxCols = .DataColCnt -1
		.Col = .MaxCols
		.ColHidden = True

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		
		If frm1.rdoTYPE1.checked then
			iRowSpan = 4
		Else
			iRowSpan = 5
		End If
		
		ret = .AddCellSpan(1, -1000, 1, iRowSpan)
		ret = .AddCellSpan(2, -1000, 1, iRowSpan)
		ret = .AddCellSpan(3, -1000, 1, iRowSpan)
		ret = .AddCellSpan(4, -1000, 1, iRowSpan)
		ret = .AddCellSpan(5, -1000, 1, iRowSpan)
		ret = .AddCellSpan(6, -1000, 1, iRowSpan)
		
		.BlockMode = True
		.Col = 7 : .Row = -1000 : .RowMerge = 1
		.Col = 7 : .Row = -999	: .RowMerge = 1
		.Col = 7 : .Row = -998	: .RowMerge = 1
		.Col = 7 : .Row = -997	: .RowMerge = 1
		.BlockMode = False
		
	End With

End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 ' -- 그리드1에서 팝업 클릭시 
Function OpenPopUp(Byval iWhere)
	Dim arrRet, sTmp
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	Select Case iWhere
		Case 0
			arrParam(0) = "원가요소 팝업"
			arrParam(1) = "dbo.C_COST_ELMT_S"	
			arrParam(2) = Trim(.txtCOST_ELMT_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "원가요소" 

			arrField(0) = "COST_ELMT_CD"	
			arrField(1) = "COST_ELMT_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "원가요소"	
			arrHeader(1) = "원가요소명"
			arrHeader(2) = ""
			
		Case 3
			arrParam(0) = "계정 팝업"
			arrParam(1) = "dbo.a_acct"	
			arrParam(2) = Trim(.txtFROM_ACCT_CD.value)
			arrParam(3) = ""
			arrParam(4) = "temp_fg_3 in ('M2','M3','M4')"
			arrParam(5) = "계정" 

			arrField(0) = "ACCT_CD"
			arrField(1) = "ACCT_NM"		
			
			arrHeader(0) = "계정"
			arrHeader(1) = "계정명"
		Case 4
			arrParam(0) = "계정 팝업"
			arrParam(1) = "dbo.a_acct"	
			arrParam(2) = Trim(.txtto_ACCT_CD.value)
			arrParam(3) = ""
			arrParam(4) = "temp_fg_3 in ('M2','M3','M4')"
			arrParam(5) = "계정" 

			arrField(0) = "ACCT_CD"
			arrField(1) = "ACCT_NM"		
			
			arrHeader(0) = "계정"
			arrHeader(1) = "계정명"

		Case 1
			arrParam(0) = "Sender C/C 팝업"
			arrParam(1) = "dbo.c_mfc_alloc_by_cc_s a left outer join b_cost_center b on a.recv_cost_cd = b.cost_cd"	
			arrParam(2) = Trim(.txtCOST_CD.value)
			arrParam(3) = ""
			arrParam(4) = "dstb_order = 0"
			arrParam(5) = "Sender C/C" 

			arrField(0) = "a.recv_cost_cd"
			arrField(1) = "b.cost_nm"		
				
			arrHeader(0) = "Sender C/C"
			arrHeader(1) = "Sender C/C명"
			
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

	End With
End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1

		Select Case iWhere
		
			Case 0
				.txtCOST_ELMT_CD.value		= arrRet(0)
				.txtCOST_ELMT_NM.value		= arrRet(1)
				
			Case 3
				.txtFROM_ACCT_CD.value		= arrRet(0)
				.txtFROM_ACCT_NM.value		= arrRet(1)

			Case 4
				.txtTO_ACCT_CD.value		= arrRet(0)
				.txtTO_ACCT_NM.value		= arrRet(1)
				
			Case 1
				.txtCOST_CD.value		= arrRet(0)
				.txtCOST_NM.value		= arrRet(1)
				
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtSTART_DT, parent.gDateFormat,2)
	Call ggoOper.FormatDate(frm1.txtEND_DT, parent.gDateFormat,2)
    'Call InitSpreadSheet
    Call InitVariables
    
'	Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("110000000001111")	
    frm1.txtSTART_DT.focus
   	Set gActiveElement = document.activeElement			    
    
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
Function GetText4Grid(Byref pGrid, Byval pCol, Byval pRow)
	With pGrid
		If pGrid.MaxRows = 0 Then Exit Function
		If pRow = "" Then pRow = .ActiveRow
		.Col = pCol : .Row = pRow : GetText4Grid = Trim(.Text)
	End With
End Function

'==============================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub txtSTART_DT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtSTART_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtSTART_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtSTART_DT.Focus
    End If
End Sub

Sub txtEND_DT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtEND_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtEND_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtEND_DT.Focus
    End If
End Sub

Sub txtCOST_ELMT_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtACCT_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub




'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	If lgRow <> Row Then
		With frm1.vspdData 
			.Col = .MaxCols : .Row = Row
			If .Text = "0" Then
				lgStrPrevKey2 = "" : lgEOF2 = False
				'Call DBQuery2
			End If
		End With
	End If
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		With frm1.vspdData 
			.Col = .MaxCols : .Row = NewRow
			If .Text = "0" Then
				lgStrPrevKey2 = "" : lgEOF2 = False
				'Call DBQuery2
			End If
		End With
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" And lgEOF1 = False Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) And lgStrPrevKey2 <> "" And lgEOF2 = False Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		'Call DisableToolBar(Parent.TBC_QUERY)
		'If DBQuery2 = False Then
		'	Call RestoreToolBar()
		'	Exit Sub
		'End If
	End if    

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD , sStartDt, sEndDt
    
    FncQuery = False
    
    Err.Clear
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtSTART_DT.text,frm1.txtEND_DT.text,frm1.txtSTART_DT.Alt,frm1.txtEND_DT.Alt, _
	    	               "970024",frm1.txtSTART_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtSTART_DT.focus
	   Exit Function
	End If
	
	If ChkKeyField=false then Exit Function 
    
    Call ggoOper.ClearField(Document, "2")

	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	'frm1.vspdData2.style.display = "none"
	'frm1.vspdData2.MaxRows = 0
    'ggoSpread.Source = frm1.vspdData2
	'ggoSpread.ClearSpreadData

    Call InitVariables 	

    IF DbQuery = False Then
		Exit Function
	END IF
       
    FncQuery = True		
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False 
    
    Err.Clear     

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")   
			If IntRetCD = vbNo Then
				Exit Function
			End If
    End If
    

    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2") 
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

     
    Call ggoOper.LockField(Document, "N") 
    Call InitVariables 
    Call SetDefaultVal
    
    FncNew = True 

End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
    
    FncSave = True      
    
End Function


'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows = 0 then exit function 
	   
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow
    
    Dim iSeqNo
    
	frm1.vspdData.ReDraw = True
End Function


Function FncCancel() 
    Dim lDelRows

	lgBlnFlgChgValue = True
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD, iSeqNo, iSubSeqNo
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If	
	End If

	With frm1

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	End With
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function


Function FncDeleteRow() 
    Dim lDelRows
	
	lgBlnFlgChgValue = True
End Function
Function FncPrint()
    Call parent.FncPrint() 
End Function

Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    Dim sStartDt, sEndDt, sYear, sMon, sDay
    
    With frm1
    
		If lgIntFlgMode = Parent.OPMD_CMODE Then
		
			Call parent.ExtractDateFromSuper(.txtSTART_DT.Text, parent.gDateFormat,sYear,sMon,sDay)
			sStartDt = sYear & sMon 
		
			Call parent.ExtractDateFromSuper(.txtEND_DT.Text, parent.gDateFormat,sYear,sMon,sDay)
			sEndDt = sYear & sMon
		
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtStartDt=" & sStartDt
			strVal = strVal & "&txtEndDt=" & sEndDt	
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtCOST_ELMT_CD=" & Trim(.txtCOST_ELMT_CD.value)
			strVal = strVal & "&txtFROM_ACCT_CD=" & Trim(.txtFROM_ACCT_CD.value)
			strVal = strVal & "&txtTO_ACCT_CD=" & Trim(.txtTO_ACCT_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.txtCOST_CD.value)

			If .rdoTYPE1.checked then
				strVal = strVal & "&rdoTYPE=1"
			Else
				strVal = strVal & "&rdoTYPE=2"
			End If
		
			'If lgStrPrevKey = "" Then Call InitSpreadSheet
		
			strVal = strVal & "&txtGrid=A"
			'strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			sStartDt = Trim(.hSTART_DT.value)
			sEndDt = Trim(.hEND_DT.value)

			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtStartDt=" & sStartDt
			strVal = strVal & "&txtEndDt=" & sEndDt	
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtCOST_ELMT_CD=" & Trim(.hCOST_ELMT_CD.value)
			strVal = strVal & "&txtFROM_ACCT_CD=" & Trim(.hFROM_ACCT_CD.value)
			strVal = strVal & "&txtTO_ACCT_CD=" & Trim(.hTO_ACCT_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.hCOST_CD.value)
			strVal = strVal & "&rdoTYPE=" & Trim(.hTYPE.value)
			strVal = strVal & "&txtGrid=A"
			
		End If
				
		lgSTime = Time	' -- 디버깅용 
		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery = True 
    

End Function

' -- 헤더 그리드 클릭시 디테일 그리드 조회 
Function DbQuery2() 
	Dim strVal

    DbQuery2 = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    Dim sStartDt, sEndDt, sYear, sMon, sDay
    
    If lgStrPrevKey2 = "" Then Call InitSpreadSheet2
    
    With frm1
		
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtStartDt=" & Trim(.hSTART_DT.value)
		strVal = strVal & "&txtEndDt=" & Trim(.hEND_DT.value)	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey2
		strVal = strVal & "&txtACCT_CD=" & GetText4Grid(frm1.vspdData, 4, "")
		strVal = strVal & "&rdoTYPE=" & Trim(.hTYPE.value)
		strVal = strVal & "&txtGrid=B"
		'strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		
		lgSTime = Time	' -- 디버깅용 
		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery2 = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	
	frm1.vspdData.style.display = ""	'-- 그리드 보이게..
	
	Frm1.vspdData.Focus
   	
   	Call ReInitSpreadSheet
	
	lgIntFlgMode = Parent.OPMD_UMODE   	

    Set gActiveElement = document.ActiveElement   

	'Call SetQuerySpreadColor
	
	window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"
End Function

Function DbQueryOk2()	
	
	frm1.vspdData2.style.display = ""	'-- 그리드 보이게..
	
	Frm1.vspdData2.Focus
   	
   	Call ReInitSpreadSheet2
   	

    Set gActiveElement = document.ActiveElement   

	'Call SetQuerySpreadColor2
	lgRow = frm1.vspdData.ActiveRow
	
	window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"
End Function

'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor(Byval pGrpRow)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt

	With frm1.vspdData	
	.ReDraw = False
	arrRow = Split(pGrpRow, Parent.gRowSep)
	
	iLoopCnt = UBound(arrRow, 1)
	
	For i = 0 to iLoopCnt -1
		arrCol = Split(arrRow(i), Parent.gColSep)
	
		.Col = -1
		.Row = CDbl(arrCol(1))	' -- 행 
		
		Select Case arrCol(0)
			Case "1"
				iRow = .Row
				.Col = -1
			   ret = .AddCellSpan(4, iRow , 3, 1)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
			Case "2"
				iRow = .Row
				.Col = -1
				ret = .AddCellSpan(2, iRow , 5, 1)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
			Case "3"
				iRow = .Row
				.Col = -1
				ret = .AddCellSpan(1, iRow , 6, 1)
				.BackColor = RGB(204,255,255) 
				.ForeColor = vbBlack
			Case "4"  
				iRow = .Row
				.Col = -1
				'ret = .AddCellSpan(1, iLoopCnt + 1, C_MVMT_QTY-iCnt, 1)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
			Case "5" 
				'ret = .AddCellSpan(1, iLoopCnt + 1, C_MVMT_QTY, 1)
				.BackColor = RGB(255,240,245) 
				.ForeColor = vbBlack
		End Select
		'.BlockMode = False
	Next



	.ReDraw = True
	End With

End Sub

Sub SetQuerySpreadColor2(Byval pGrpRow)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt

	With frm1.vspdData2
	.ReDraw = False
	arrRow = Split(pGrpRow, Parent.gRowSep)
	
	iLoopCnt = UBound(arrRow, 1)
	
	For i = 0 to iLoopCnt -1
		arrCol = Split(arrRow(i), Parent.gColSep)
	
		.Col = -1
		.Row = CDbl(arrCol(1))	' -- 행 
		
		Select Case arrCol(0)
			Case "1"
				iRow = .Row
				.Col = -1
			   ret = .AddCellSpan(3, iRow , 1, 1)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
			Case "2"
				iRow = .Row
				.Col = -1
				ret = .AddCellSpan(1, iRow , 3, 1)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
			Case "3"
				iRow = .Row
				.Col = -1
				ret = .AddCellSpan(1, iRow , 6, 1)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
			Case "4"  
				iRow = .Row
				.Col = -1
				
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
			Case "5" 
				
				.BackColor = RGB(255,240,245) 
				.ForeColor = vbBlack
		End Select

	Next

	.ReDraw = True
	End With

End Sub

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 

    DbSave = True    
    
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()	
   
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true	


	If Trim(frm1.txtCOST_ELMT_Cd.value) <> "" Then
		' -- 변경값 체크 
		strWhere = " COST_ELMT_CD = " & FilterVar(frm1.txtCOST_ELMT_Cd.value, "''", "S")
		
		Call CommonQueryRs(" COST_ELMT_CD, COST_ELMT_NM "," C_COST_ELMT_S ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCOST_ELMT_CD.alt,"X")
			frm1.txtCOST_ELMT_Nm.value = ""
			frm1.txtCOST_ELMT_CD.focus
			ChkKeyField = False
			Exit Function
		End If
			frm1.txtCOST_ELMT_Cd.value = Replace(lgF0, Chr(11), "")
			frm1.txtCOST_ELMT_Nm.value = Replace(lgF1, Chr(11), "")

	Else
		frm1.txtCOST_ELMT_Nm.value = ""
	End If	
	
	If Trim(frm1.txtFROM_ACCT_CD.value) <> "" Then

		' -- 변경값 체크 
		strWhere = " acct_cd = " & FilterVar(frm1.txtFROM_ACCT_CD.value, "''", "S")
		strWhere = strWhere & "	and temp_fg_3 in ('M2','M3','M4')" 
		
		Call CommonQueryRs(" ACCT_CD, ACCT_NM "," A_ACCT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			frm1.txtFROM_ACCT_Nm.value = ""
		Else
			frm1.txtFROM_ACCT_CD.value = Replace(lgF0, Chr(11), "")
			frm1.txtFROM_ACCT_Nm.value = Replace(lgF1, Chr(11), "")
		End If 
	Else
		frm1.txtFROM_ACCT_Nm.value = ""
	End If	

	
	If Trim(frm1.txtTO_ACCT_CD.value) <> "" Then

		' -- 변경값 체크 
		strWhere = " acct_cd = " & FilterVar(frm1.txtTO_ACCT_CD.value, "''", "S")
		strWhere = strWhere & "	and temp_fg_3 in ('M2','M3','M4')" 
		
		Call CommonQueryRs(" ACCT_CD, ACCT_NM "," A_ACCT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtTO_ACCT_CD.alt,"X")
			frm1.txtTO_ACCT_Nm.value = ""
			frm1.txtTO_ACCT_CD.focus
			ChkKeyField = False
			Exit Function
		End If
			frm1.txtTO_ACCT_CD.value = Replace(lgF0, Chr(11), "")
			frm1.txtTO_ACCT_Nm.value = Replace(lgF1, Chr(11), "")
 
	Else
		frm1.txtTO_ACCT_Nm.value = ""
	End If	

	If Trim(frm1.txtCOST_CD.value) <> "" Then

		' -- 변경값 체크 
		strWhere = " CODE = " & FilterVar(frm1.txtCOST_CD.value, "''", "S")
		
		Call CommonQueryRs(" CODE, CD_NM ","dbo.ufn_c_getListOfPopup_C4002MA1('2') ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCOST_CD.alt,"X")
			frm1.txtCOST_NM.value = ""
			frm1.txtCOST_CD.focus
			ChkKeyField = False
			Exit Function
		End If
			frm1.txtCOST_CD.value = Replace(lgF0, Chr(11), "")
			frm1.txtCOST_NM.value = Replace(lgF1, Chr(11), "")
 
	Else
		frm1.txtCOST_NM.value = ""
	End If	
End Function 

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;&nbsp;</TD>
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
									<TD CLASS="TD5">작업년월</TD>
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSTART_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 기준년월" tag="12" id=txtSTART_DT></OBJECT>');</SCRIPT>&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtEND_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="종료 기준년월" tag="12" id=txtEND_DT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>원가요소</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtCOST_ELMT_CD" TYPE="Text" MAXLENGTH="20" tag="15XXXU" size="10" ALT="원가요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtCOST_ELMT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">Cost Center</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtCOST_CD" TYPE="Text" MAXLENGTH="10" tag="15XXX" size="10" ALT="Sender C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)">
									<input NAME="txtCOST_NM" TYPE="TEXT" MAXLENGTH="20" tag="14XXX" size="30">
									</TD>
									<TD CLASS="TD5">계정</TD>
									<TD CLASS="TD6" valign=top><input NAME="txtFROM_ACCT_CD" TYPE="Text" MAXLENGTH="20" tag="15XXXU" size="10" ALT="시작 계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(3)">
									<input NAME="txtFROM_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="30">&nbsp;~<br>
									<input NAME="txtTO_ACCT_CD" TYPE="Text" MAXLENGTH="20" tag="15XXXU" size="10" ALT="종료 계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(4)">
									<input NAME="txtTO_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="30">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">집계구분</TD>
									<TD CLASS="TD6" valign=top><INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE1 tag="15XXX" checked><LABEL FOR="rdoTYPE1">기간</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE2 tag="15XXX"><LABEL FOR="rdoTYPE2">월별</LABEL>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>    

							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData1 NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>    
<!--							<TR>
								<TD HEIGHT="35%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR> -->
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hSTART_DT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hEND_DT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCOST_ELMT_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hFROM_ACCT_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTO_ACCT_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCOST_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTYPE" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

