<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :제조원가명세서조회 
'*  3. Program ID           : c4201mb1.asp
'*  4. Program Name         :제조원가명세서조회 
'*  5. Program Desc         :제조원가명세서조회 
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

Const BIZ_PGM_ID = "c4201mb1.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt
Dim lgRow

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

Dim lgSTime	' -- 디버깅 타임체크 
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
	<%Call LoadInfTB19029A("Q","C", "NOCOOKIE", "QA") %>
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
Sub InitSpreadSheet()
	Dim i, ret

	With frm1.vspdData1

		.Redraw = False

		ggoSpread.Source = frm1.vspdData1
		'ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread	' 풀면 드랙&드롭시 컬럼이 바뀐다 
		ggoSpread.Spreadinit "V20021122",,""	' 풀면 드랙&드롭시 컬럼이 바뀐다 

		.MaxRows = 0

		if frm1.rdoTYPE_FLAG2.checked Or frm1.rdoTYPE_FLAG4.checked then 
			.MaxCols = 10
		Else
			.MaxCols = 8
		End If
		.Col = .MaxCols
		.ColHidden = True

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		.Col = 5: .Row = -1: .ColMerge = 1
		.Col = 6: .Row = -1: .ColMerge = 1

		ggoSpread.SSSetEdit		1,	"과목대분류", 10,,,,1	
		ggoSpread.SSSetEdit		2,	"과목대분류", 12
		ggoSpread.SSSetEdit		3,	"품목계정"	, 10,,,,1
		ggoSpread.SSSetEdit		4,	"품목계정"	, 12	
		ggoSpread.SSSetEdit		5,	"과목"	, 10,,,,1
		ggoSpread.SSSetEdit		6,	"과목"	, 20		
		
		If frm1.rdoTYPE_FLAG2.checked Or frm1.rdoTYPE_FLAG4.checked then 
			ggoSpread.SSSetEdit		7,	"계정"	, 12,,,,1	
			ggoSpread.SSSetEdit		8,	"계정명"	, 20
			ggoSpread.SSSetFloat	9,	"금액"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Else
			ggoSpread.SSSetFloat	7,	"금액"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		End If
		
		Call ggoSpread.SSSetColHidden(1 , 1, True)
		Call ggoSpread.SSSetColHidden(3 , 3, True)
		Call ggoSpread.SSSetColHidden(5 , 5, True)
		'ggoSpread.SSSetSplit2(6)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'ggoSpread.SSSetProtected	.MaxCols,-1,-1

		.Redraw = True
	End With
End Sub
		
Sub InitSpreadSheet2(Byval pMaxCols)
	Dim i, ret, iCol

	With frm1.vspdData2

		.Redraw = False

		ggoSpread.Source = frm1.vspdData2
		'ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread
		ggoSpread.Spreadinit "V20021122",,"" 


		.MaxRows = 0
		.MaxCols = pMaxCols

		'헤더를 2줄로    
		.ColHeaderRows = 2

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		.Col = 5: .Row = -1: .ColMerge = 1
		.Col = 6: .Row = -1: .ColMerge = 1

		ggoSpread.SSSetEdit		1,	"과목대분류", 10,,,,1	
		ggoSpread.SSSetEdit		2,	"과목대분류", 12
		ggoSpread.SSSetEdit		3,	"품목계정"	, 10,,,,1
		ggoSpread.SSSetEdit		4,	"품목계정"	, 12	
		ggoSpread.SSSetEdit		5,	"과목"	, 10,,,,1
		ggoSpread.SSSetEdit		6,	"과목"	, 20
		
		If frm1.rdoTYPE_FLAG2.checked Or frm1.rdoTYPE_FLAG4.checked then 
			ggoSpread.SSSetEdit		7,	"과목상세"	, 12,,,,1	
			ggoSpread.SSSetEdit		8,	"과목상세명"	, 20
			iCol = 9
		Else
			iCol = 7
		End If

		For i = iCol To pMaxCols
			ggoSpread.SSSetFloat	i,		""	, 15,		Parent.ggAmtOfMoneyNo		,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Next
		
		Call ggoSpread.SSSetColHidden(1 , 1, True)
		Call ggoSpread.SSSetColHidden(3 , 3, True)
		Call ggoSpread.SSSetColHidden(5 , 5, True)
		Call ggoSpread.SSSetColHidden(.MaxCols , .MaxCols, True)
		
		ret = .AddCellSpan(2, -1000 , 1, 2)
		ret = .AddCellSpan(4, -1000 , 1, 2)
		ret = .AddCellSpan(6, -1000 , 1, 2)
		
		If frm1.rdoTYPE_FLAG2.checked Or frm1.rdoTYPE_FLAG4.checked then 
			ret = .AddCellSpan(7, -1000 , 1, 2)
			ret = .AddCellSpan(8, -1000 , 1, 2)
		Else
			ret = .RemoveCellSpan(7, -1000)
			ret = .RemoveCellSpan(8, -1000)
		End If
		
		'ggoSpread.SSSetSplit2(6)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'ggoSpread.SSSetProtected	.MaxCols,-1,-1

		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 

		.Redraw = True
	End With
End Sub

Sub InitSpreadSheet3(Byval pMaxCols)
	Dim i, ret, iCol

	With frm1.vspdData3

		.Redraw = False

		ggoSpread.Source = frm1.vspdData3
		ggoSpread.Spreadinit "V20021122"',parent.gAllowDragDropSpread
		ggoSpread.Spreadinit "V20021122",,""

		.MaxRows = 0
		.MaxCols = pMaxCols

		'헤더를 2줄로    
		.ColHeaderRows = 3

		ggoSpread.SSSetEdit		1,	"품목계정"	, 10,,,,1
		ggoSpread.SSSetEdit		2,	"품목계정"	, 12	
		ggoSpread.SSSetEdit		3,	"과목"	, 10,,,,1
		ggoSpread.SSSetEdit		4,	"과목"	, 20	

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		
		If frm1.rdoTYPE_FLAG2.checked Or frm1.rdoTYPE_FLAG4.checked then 
			ggoSpread.SSSetEdit		5,	"과목상세"	, 12,,,,1	
			ggoSpread.SSSetEdit		6,	"과목상세명"	, 20
			iCol = 7
		Else
			iCol = 5
		End If

		For i = iCol To pMaxCols
			ggoSpread.SSSetFloat	i,		""	, 15,		Parent.ggAmtOfMoneyNo		,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Next
		
		Call ggoSpread.SSSetColHidden(1 , 1, True)
		Call ggoSpread.SSSetColHidden(3 , 3, True)
		Call ggoSpread.SSSetColHidden(.MaxCols , .MaxCols, True)
		'ggoSpread.SSSetSplit2(6)

		ret = .AddCellSpan(2, -1000 , 1, 3)
		ret = .AddCellSpan(4, -1000 , 1, 3)
		
		If frm1.rdoTYPE_FLAG2.checked Or frm1.rdoTYPE_FLAG4.checked then 
			ret = .AddCellSpan(5, -1000 , 1, 3)
			ret = .AddCellSpan(6, -1000 , 1, 3)
		Else
			ret = .RemoveCellSpan(5, -1000)
			ret = .RemoveCellSpan(6, -1000)
		End If
		
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'ggoSpread.SSSetProtected	.MaxCols,-1,-1

		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 

		.Redraw = True
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
		.Col = 5: .Row = -1: .ColMerge = 1
		.Col = 6: .Row = -1: .ColMerge = 1
		
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

'	.rowheight(-1000) = 20	' 높이 재지정 
End Sub

Sub ReInitSpreadSheet2()
	
	Dim ret, iRowSpan
	' -- 그리드 1 정의 
	With frm1.vspdData2

		'.MaxCols = .DataColCnt -1
		.Col = .MaxCols
		.ColHidden = True

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1

		If frm1.rdoTYPE1.checked then
			iRowSpan = 4
		Else
			iRowSpan = 5
		End If
		
		ret = .AddCellSpan(1, -1000, 1, iRowSpan)
		ret = .AddCellSpan(2, -1000, 1, iRowSpan)
		ret = .AddCellSpan(3, -1000, 1, iRowSpan)
				
		.BlockMode = True
		.Col = 4 : .Col2 = .MaxCols - 1
		.Row = -1000 : .Row2 = -1000 
		.RowMerge = 1
		.Col = 4 : .Col2 = .MaxCols - 1
		.Row = -999 : .Row2 = -999 
		.RowMerge = 1
		.Col = 4 : .Col2 = .MaxCols - 1
		.Row = -998 : .Row2 = -999 
		.RowMerge = 1
		.Col = 4 : .Col2 = .MaxCols - 1
		.Row = -997 : .Row2 = -999 
		.RowMerge = 1
		.BlockMode = False
		
	End With

'	.rowheight(-1000) = 20	' 높이 재지정 
End Sub

Sub SetGridHead2(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol

	' -- 그리드 1 정의 
	With frm1.vspdData2
		
		arrRows = Split(pData, Parent.gRowSep)

		'헤더를 ?줄로    
		.ColHeaderRows = UBound(arrRows, 1)
		
		For i = 0 To UBound(arrRows, 1) -1
			arrCols = Split(arrRows(i), Parent.gColSep)
			iColCnt = UBound(arrCols, 1)
			.Row	= CDbl(arrCols(iColCnt))		' -- 마지작 컬럼에 행번호가 들어있다.
			
			If frm1.rdoTYPE_FLAG1.checked Or frm1.rdoTYPE_FLAG3.checked Then
				iCol = 7	
			Else
				iCol = 9
			End if
			
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

Sub SetGridHead3(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol

	' -- 그리드 1 정의 
	With frm1.vspdData3
		
		arrRows = Split(pData, Parent.gRowSep)

		'헤더를 ?줄로    
		.ColHeaderRows = UBound(arrRows, 1)
    
		For i = 0 To UBound(arrRows, 1) -1
			arrCols = Split(arrRows(i), Parent.gColSep)
			iColCnt = UBound(arrCols, 1)
			.Row	= CDbl(arrCols(iColCnt))		' -- 마지작 컬럼에 행번호가 들어있다.

			If frm1.rdoTYPE_FLAG1.checked Or frm1.rdoTYPE_FLAG3.checked Then
				iCol = 5	
			Else
				iCol = 7
			End if

			For j = 0 To iColCnt 
				.Col = iCol
				Select Case j
					Case 0, 1, 2, iColCnt
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
			arrParam(0) = "공장 팝업"
			arrParam(1) = "dbo.b_plant"	
			arrParam(2) = Trim(.txtPLANT_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장" 

			arrField(0) = "plant_cd"	
			arrField(1) = "plant_nm"
			arrField(2) = ""		
    
			arrHeader(0) = "공장"	
			arrHeader(1) = "공장명"
			arrHeader(2) = ""
			
		Case 1
			arrParam(0) = "작업지시 C/C 팝업"
			arrParam(1) = "dbo.b_cost_center"	
			arrParam(2) = Trim(.txtCOST_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "작업지시 C/C" 

			arrField(0) = "cost_cd"
			arrField(1) = "cost_NM"		
			
			arrHeader(0) = "작업지시 C/C"
			arrHeader(1) = "작업지시 C/C명"

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
				.txtPLANT_CD.value		= arrRet(0)
				.txtPLANT_NM.value		= arrRet(1)
				
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
    
	'Call InitComboBox
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

'========================================================================================================

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

Sub txtPlantCD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtCostCD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

'=======================================================================================================
'   Event Name : rdoType_onClick
'   Event Desc : 
'=======================================================================================================
Sub  rdoTYPE1_onClick()
	DivPlantCd(0).style.display = "none"
	DivPlantCd(1).style.display = "none"
	
	DivCostCd(0).style.display = "none"
	DivCostCd(1).style.display = "none"
	
	DivSlim.style.display = ""
End Sub

Sub  rdoTYPE2_onClick()
	DivPlantCd(0).style.display = ""
	DivPlantCd(1).style.display = ""
	
	DivCostCd(0).style.display = "none"
	DivCostCd(1).style.display = "none"
	
	DivSlim.style.display = ""
End Sub

Sub  rdoTYPE3_onClick()
	DivSlim.style.display = "none"
	
	DivPlantCd(0).style.display = "none"
	DivPlantCd(1).style.display = "none"
	
	DivCostCd(0).style.display = ""
	DivCostCd(1).style.display = ""
	
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	'Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	'----------------------
	'Column Split
	'----------------------
	'gMouseClickStatus = "SPC"
	
	'Set gActiveSpdSheet = frm1.vspdData1
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

Sub vspdData3_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) And lgStrPrevKey <> "" Then
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
	    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub


Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub
'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    

End Sub

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    

End Sub

Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    

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
    
    Call ggoOper.ClearField(Document, "2")

	frm1.vspdData1.MaxRows = 0
    ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData

	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

	frm1.vspdData3.MaxRows = 0
    ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    Call InitVariables 	

	DivGrid(0).style.display = "none"
	DivGrid(1).style.display = "none"
	DivGrid(2).style.display = "none"

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
	If lgStrPrevKey <> "" Then
		lgStrPrevKey = lgStrPrevKey & parent.gColSep & "*"	' 현재 키값 & * 를 보낸다 
		Call DBQuery()
		Exit Function
	End If

	Call parent.FncExport(Parent.C_MULTI)

End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
'Sub FncSplitColumn()

 '   If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
  '     Exit Sub
 '   End If
'
  '  ggoSpread.Source = gActiveSpdSheet
   ' ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
'End Sub


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
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
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
		
			sStartDt = sYear & parent.gComDateType & sMon & parent.gComDateType & sDay
		
			If .txtEND_DT.text = "" then
				sEndDt = iDBSYSDate
			Else
				Call parent.ExtractDateFromSuper(.txtEND_DT.Text, parent.gDateFormat,sYear,sMon,sDay)
				sEndDt = sYear & parent.gComDateType & sMon & parent.gComDateType & sDay
				sEndDt = DateAdd("m", 1, sEndDt)-1
			End If
		Else
			sStartDt = .hSTART_DT.value 
			sEndDt	= .hEND_DT.value 
		End If
				
		strVal = BIZ_PGM_ID & "?txtMode=" & lgIntFlgMode
		strVal = strVal & "&txtStartDt=" & sStartDt
		strVal = strVal & "&txtEndDt=" & sEndDt	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		
		If lgIntFlgMode = Parent.OPMD_CMODE Then
			strVal = strVal & "&txtPLANT_CD=" & Trim(.txtPLANT_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.txtCOST_CD.value)

			' -- 집계단위 
			If .rdoTYPE1.checked then
				strVal = strVal & "&rdoTYPE=1"
			ElseIf .rdoTYPE2.checked then
				strVal = strVal & "&rdoTYPE=2"
			Else
				strVal = strVal & "&rdoTYPE=3"
			End If

			' -- 구분 
			If .rdoTYPE_FLAG1.checked then
				strVal = strVal & "&rdoTYPE_FLAG=1"
			ElseIf .rdoTYPE_FLAG2.checked then
				strVal = strVal & "&rdoTYPE_FLAG=2"
			ElseIf .rdoTYPE_FLAG3.checked then
				strVal = strVal & "&rdoTYPE_FLAG=3"
			Else
				strVal = strVal & "&rdoTYPE_FLAG=4"
			End If
			
		Else
			strVal = strVal & "&txtPLANT_CD=" & Trim(.hPLANT_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.hCOST_CD.value)
			strVal = strVal & "&rdoTYPE=" & Trim(.hTYPE.value)
			strVal = strVal & "&rdoTYPE_FLAG=" & Trim(.hTYPE_FLAG.value)
		End If

		lgSTime = Time	' -- 디버깅용 
		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery = True 
    

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	

	If frm1.rdoTYPE1.checked then
		DivGrid(0).style.display = ""
		DivGrid(1).style.display = "none"
		DivGrid(2).style.display = "none"
		Frm1.vspdData1.Focus
	ElseIf frm1.rdoTYPE2.checked then
		DivGrid(0).style.display = "none"
		DivGrid(1).style.display = ""
		DivGrid(2).style.display = "none"
		Frm1.vspdData2.Focus
	Else
		DivGrid(0).style.display = "none"
		DivGrid(1).style.display = "none"
		DivGrid(2).style.display = ""
		Frm1.vspdData3.Focus
	End If

    Set gActiveElement = document.ActiveElement   

	lgIntFlgMode = Parent.OPMD_UMODE
	
	'window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"
	
	'If lgStrPrevKey = "*" Then
	'	lgStrPrevKey = ""
	'	Call FncExcel() 
	'End If
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
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
			Case "4"  
				iRow = .Row
				.Col = -1
				'ret = .AddCellSpan(1, iLoopCnt + 1, C_MVMT_QTY-iCnt, 1)
				.BackColor = RGB(255,240,245) 
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
				.BackColor = RGB(204,255,255) 
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
									<TD CLASS="TD5">기준년월</TD>
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSTART_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 기준년월" tag="12" id=txtSTART_DT></OBJECT>');</SCRIPT>&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtEND_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="종료 기준년월" tag="12" id=txtEND_DT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>집계단위</TD>
									<TD CLASS="TD6" valign=top><INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE1 tag="15XXX" checked><LABEL FOR="rdoTYPE1">Company</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE2 tag="15XXX"><LABEL FOR="rdoTYPE2">공장</LABEL>
									<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE3 tag="15XXX"><LABEL FOR="rdoTYPE3">작업지시 C/C</LABEL>
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">구분</TD>
									<TD CLASS="TD6" valign=top><INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE_FLAG ID=rdoTYPE_FLAG1 tag="15XXX" checked><LABEL FOR="rdoTYPE_FLAG1">집계</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE_FLAG ID=rdoTYPE_FLAG2 tag="15XXX"><LABEL FOR="rdoTYPE_FLAG2">상세</LABEL>
									<DIV id=DivSlim><INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE_FLAG ID=rdoTYPE_FLAG3 tag="15XXX"><LABEL FOR="rdoTYPE_FLAG3">집계Siml</LABEL>
									<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE_FLAG ID=rdoTYPE_FLAG4 tag="15XXX"><LABEL FOR="rdoTYPE_FLAG4">상세Siml</LABEL></DIV>
									</TD>
									<TD CLASS="TD5" NOWRAP><DIV id=DivPlantCd style="display:none">공장</DIV>
									<DIV id=DivCostCd style="display:none">작업지시 C/C<DIV></TD>
									<TD CLASS="TD6" NOWRAP><DIV id=DivPlantCd style="display:none"><input NAME="txtPLANT_CD" TYPE="Text" MAXLENGTH="4" tag="15XXX" size="20" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPLANT_NM" TYPE="TEXT" MAXLENGTH="10" tag="14XXX" size="20"></DIV>
									<DIV id=DivCostCd style="display:none"><input NAME="txtCOST_CD" TYPE="Text" MAXLENGTH="10" tag="15XXX" size="20" ALT="작업지시 C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)">
									<input NAME="txtCOST_NM" TYPE="TEXT" MAXLENGTH="10" tag="14XXX" size="20"></DIV>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
					<DIV ID=divGrid style="display: none">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData1 NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</DIV>

					<DIV ID=divGrid style="display: none">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</DIV>

					<DIV ID=divGrid style="display: none">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData3 NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</DIV>
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
<INPUT TYPE=HIDDEN NAME="hPLANT_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCOST_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTYPE" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTYPE_FLAG" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

