<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 품목별원가수불장조회 
'*  3. Program ID           : c4202mb1.asp
'*  4. Program Name         : 품목별원가수불장조회 
'*  5. Program Desc         : 품목별원가수불장조회 
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

Const BIZ_PGM_ID = "c4202mb1.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt

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
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtSTART_DT.Text = Left(iStrFromDt, 7)
	frm1.txtEND_DT.Text = Left(iStrFromDt, 7)

    If parent.gPlant <> "" Then
		frm1.txtPlant_Cd.value = UCase(parent.gPlant)
		frm1.txtPlant_Nm.value = parent.gPlantNm
	End If
	
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
Sub InitSpreadSheet()
	Dim i, ret
	With frm1.vspdData
		

		frm1.vspdData2.style.display = "none"
		.style.display = ""

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021122",,""   ',parent.gAllowDragDropSpread 

		.MaxRows = 0
		.MaxCols = 37

		.Col  = 36 : .ColHidden = True
		.Col  = 37 : .ColHidden = True
		
		'헤더를 2줄로    
		.ColHeaderRows = 2

		ggoSpread.SSSetEdit		1,	"공장"	, 6,,,20,1	
		ggoSpread.SSSetEdit		2,	"품목계정"	, 6,,,20,1	
		ggoSpread.SSSetEdit		3,	"품목계정명"	, 8
		ggoSpread.SSSetEdit		4,	"품목"	, 15,,,20,1	
		ggoSpread.SSSetEdit		5,	"품목명"	, 15	
		
		ggoSpread.SSSetFloat	6,	"기초재고"	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	7,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	8,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	9,	"입고 (수량, 금액, 단가)"	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	10,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	11,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	12,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	13,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	14,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	15,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	16,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	17,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	18,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	19,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	20,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	21,	"출고 (수량, 금액, 단가)"	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	22,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	23,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	24,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	25,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	26,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	27,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	28,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	29,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	30,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	31,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	32,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		ggoSpread.SSSetFloat	33,	"기말재고"	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	34,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	35,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		'Call ggoSpread.SSSetColHidden(36,37,True)

		For i = 1 To 5
			ret = .AddCellSpan(i, -1000, 1, 2)
		Next
		
		ret = .AddCellSpan(6, -1000, 3, 2)	' -- 기초 
		ret = .AddCellSpan(9, -1000, 12, 1)	' -- 입고 
		ret = .AddCellSpan(21, -1000, 12, 1)	' -- 출고 
		ret = .AddCellSpan(33, -1000, 3, 2)	' -- 기말 

		ret = .AddCellSpan(9, -999, 3, 1)	' -- 입고 
		ret = .AddCellSpan(12, -999, 3, 1)	
		ret = .AddCellSpan(15, -999, 3, 1)	
		ret = .AddCellSpan(18, -999, 3, 1)	
		
		ret = .AddCellSpan(21, -999, 3, 1)	' -- 출고 
		ret = .AddCellSpan(24, -999, 3, 1)	
		ret = .AddCellSpan(27, -999, 3, 1)	
		ret = .AddCellSpan(30, -999, 3, 1)
		
		' 두번째 헤더 출력 글자 
		.Row = -999
		.Col = 9	: .Text = "구매입고"
		.Col = 12	: .Text = "생산입고"
		.Col = 15	: .Text = "예외입고"
		.Col = 18	: .Text = "이동입고"
		
		.Col = 21	: .Text = "생산출고"
		.Col = 24	: .Text = "판매출고"
		.Col = 27	: .Text = "예외출고"
		.Col = 30	: .Text = "이동출고"
		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 
		
		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1

		ggoSpread.SpreadLockWithOddEvenRowColor()
		
	End With
End Sub

Sub InitSpreadSheet2()
	Dim i, ret
	With frm1.vspdData2

		frm1.vspdData.style.display = "none"
		.style.display = ""
		.Redraw = False

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021122",,"" ',parent.gAllowDragDropSpread 

		.MaxRows = 0

		If .MaxCols > 1 Then
			ret = .RemoveCellSpan(.MaxCols-3, -1000)
		End If
		.MaxCols = 100

		'헤더를 2줄로    
		.ColHeaderRows = 4

		ggoSpread.SSSetEdit		1,	"공장"	, 6,,,20,1	
		ggoSpread.SSSetEdit		2,	"품목계정"	, 6,,,20,1	
		ggoSpread.SSSetEdit		3,	"품목계정명"	, 8
		ggoSpread.SSSetEdit		4,	"품목"	, 10,,,20,1	
		ggoSpread.SSSetEdit		5,	"품목명"	, 15

		For i = 6 To 98	Step 3 ' -- 숫자형 필드부터 
			ggoSpread.SSSetFloat	i,		""	, 10,		Parent.ggQtyNo		,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	i+1,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	i+2,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Next
		
		ggoSpread.SpreadLockWithOddEvenRowColor()

		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 
		.rowheight(-998) = 12	' 높이 재지정 
		.rowheight(-997) = 12	' 높이 재지정 
		
		.Redraw = True
	End With
End Sub
		
Sub ReInitSpreadSheet2()
	
	Dim ret
	' -- 그리드 1 정의 
	With frm1.vspdData2

		'.MaxCols = .DataColCnt -1
		.Col = .MaxCols
		.ColHidden = True

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		.Col = 5: .Row = -1: .ColMerge = 1
		
		ret = .AddCellSpan(1, -1000, 1, 4)
		ret = .AddCellSpan(2, -1000, 1, 4)
		ret = .AddCellSpan(3, -1000, 1, 4)
		ret = .AddCellSpan(4, -1000, 1, 4)
		ret = .AddCellSpan(5, -1000, 1, 4)
		
		ret = .AddCellSpan(6, -1000, 3, 4)
		ret = .AddCellSpan(.MaxCols-3, -1000, 3, 4)

		.BlockMode = True
		.Col = 6 : .Col2 = .MaxCols - 4
		.Row = -1000 : .Row2 = -1000 
		.RowMerge = 1
		.Col = 6 : .Col2 = .MaxCols - 4
		.Row = -999 : .Row2 = -999 
		.RowMerge = 1
		.Col = 6 : .Col2 = .MaxCols - 4
		.Row = -998 : .Row2 = -998 
		.RowMerge = 1
		.Col = 6 : .Col2 = .MaxCols - 4
		.Row = -997 : .Row2 = -997 
		.RowMerge = 1
		.BlockMode = False
		
	End With

'	.rowheight(-1000) = 20	' 높이 재지정 
End Sub

Sub SetGridHead(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol

	Dim pObj

	' -- 그리드를 2개로 늘리면서 취해진 소스변경 
	If frm1.rdoTYPE1.checked then
		Set pObj = frm1.vspdData
	Else
		Set pObj = frm1.vspdData2
	End If

	' -- 그리드 1 정의 
	With pObj
		
		arrRows = Split(pData, Parent.gRowSep)

		'헤더를 ?줄로    
		.ColHeaderRows = UBound(arrRows, 1)
    
		For i = 0 To UBound(arrRows, 1) -1
			arrCols = Split(arrRows(i), Parent.gColSep)
			iColCnt = UBound(arrCols, 1)
			.Row	= CDbl(arrCols(iColCnt))		' -- 마지작 컬럼에 행번호가 들어있다.
			iCol = 6
			For j = 0 To iColCnt 
				.Col = iCol
				Select Case j
					Case Else
						.Text = arrCols(j)
						 iCol = iCol + 1	: .Col = iCol	' -- 수량 
						.Text = arrCols(j)
						 iCol = iCol + 1	: .Col = iCol	' -- 금액 
						.Text = arrCols(j)
						 iCol = iCol + 1	: .Col = iCol	' -- 단가 
				End SElect
				
			Next
		Next
	End With
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
    
    ggoSpread.Source = frm1.vspdData
    
    .ReDraw = False

    .ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False

	
    .vspdData.ReDraw = True
    
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
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
			arrParam(1) = "dbo.B_PLANT"	
			arrParam(2) = Trim(.txtPLANT_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장" 

			arrField(0) = "PLANT_CD"	
			arrField(1) = "PLANT_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "공장"	
			arrHeader(1) = "공장명"
			arrHeader(2) = ""
			
		Case 1
			arrParam(0) = "품목계정 팝업"
			arrParam(1) = "dbo.B_MINOR"	
			arrParam(2) = Trim(.txtITEM_ACCT.value)
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S")
			arrParam(5) = "품목계정" 

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"		
			arrField(2) = ""	
    
			arrHeader(0) = "품목계정"
			arrHeader(1) = "품목계정명"
			arrHeader(2) = "C/C LEVEL"	

		Case 2
			arrParam(0) = "품목 팝업"
			arrParam(1) = "dbo.B_ITEM a left outer join dbo.b_item_by_plant b on a.item_cd = b.item_cd "	
			arrParam(2) = Trim(.txtITEM_CD.value)
			arrParam(3) = ""	
			If frm1.txtPLANT_CD.value <> "" then
				arrParam(4) = " b.PLANT_CD = " & FilterVar(frm1.txtPLANT_CD.value, "''", "S")
			End If
			arrParam(5) = "품목" 

			arrField(0) = "a.ITEM_CD"	
			arrField(1) = "a.ITEM_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "품목"	
			arrHeader(1) = "품목명"
			arrHeader(2) = ""
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
				.txtITEM_ACCT.value		= arrRet(0)
				.txtITEM_ACCT_NM.value	= arrRet(1)
				
			Case 2
				.txtITEM_CD.value		= arrRet(0)
				.txtITEM_NM.value		= arrRet(1)
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
End Sub

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

'==========================================================================================
'   Event Desc : 배부규칙 설정확인 버튼 클릭시 
'==========================================================================================

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

Sub txtPLANT_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtITEM_ACCT_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtITEM_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtPlant_Cd_OnChange()
	Dim sWhereSQL
	
	If Trim(frm1.txtPlant_Cd.value) <> "" Then

		' -- 변경값 체크 
		sWhereSQL = " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.value, "''", "S")
		
		Call CommonQueryRs(" PLANT_CD, PLANT_NM "," B_PLANT ", sWhereSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			'Call DisplayMsgBox("900014", "x", "x", "x")		
			frm1.txtPlant_Nm.value = ""
		Else
			frm1.txtPlant_Cd.value = Replace(lgF0, Chr(11), "")
			frm1.txtPlant_Nm.value = Replace(lgF1, Chr(11), "")
		End If 
	Else
		frm1.txtPlant_Nm.value = ""
	End If	
End Sub

Sub txtITEM_ACCT_OnChange()
	Dim sWhereSQL
	
	If Trim(frm1.txtITEM_ACCT.value) <> "" Then

		' -- 변경값 체크 
		sWhereSQL = " A.ITEM_ACCT = " & FilterVar(frm1.txtITEM_ACCT.value, "''", "S")
		
		Call CommonQueryRs(" a.ITEM_ACCT, b.minor_nm "," b_item_acct_inf a left outer join B_MINOR B on a.item_acct = b.minor_cd and b.major_cd = 'P1001' ", sWhereSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			'Call DisplayMsgBox("900014", "x", "x", "x")		
			frm1.txtITEM_ACCT_Nm.value = ""
		Else
			frm1.txtITEM_ACCT.value = Replace(lgF0, Chr(11), "")
			frm1.txtITEM_ACCT_Nm.value = Replace(lgF1, Chr(11), "")
		End If 
	Else
		frm1.txtITEM_ACCT_Nm.value = ""
	End If	
End Sub

Sub txtITEM_CD_OnChange()
	Dim sWhereSQL
	
	If Trim(frm1.txtITEM_CD.value) <> "" Then

		' -- 변경값 체크 
		sWhereSQL = " a.ITEM_CD = " & FilterVar(frm1.txtITEM_CD.value, "''", "S")
		if Trim(frm1.txtPLANT_CD.value) <> "" Then
			sWhereSQL = sWhereSQL & " AND b.plant_cd = " & FilterVar(frm1.txtPLANT_CD.value, "''", "S")
		End If
		
		Call CommonQueryRs(" ITEM_CD, ITEM_NM "," B_ITEM a left outer join b_item_by_plant b on a.item_cd = b.item_cd  ", sWhereSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			'Call DisplayMsgBox("900014", "x", "x", "x")		
			frm1.txtITEM_NM.value = ""
		Else
			frm1.txtITEM_CD.value = Replace(lgF0, Chr(11), "")
			frm1.txtITEM_NM.value = Replace(lgF1, Chr(11), "")
		End If 
	Else
		frm1.txtITEM_NM.value = ""
	End If	
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    

End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
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
    
    sStartDt= Replace(frm1.txtSTART_DT.text, parent.gComDateType, "")
    sEndDt	= Replace(frm1.txtEND_DT.text, parent.gComDateType, "")
    
    If sStartDt > sEndDt And sEndDt <> "" Then
	   Call DisplayMsgBox("970024", Parent.VB_INFORMATION, frm1.txtSTART_DT.alt,frm1.txtEND_DT.alt)   
	   Exit Function
	End If
    
    If ChkKeyField=False then Exit function
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables 	
'    Call InitSpreadSheet

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
End Function


Function FncCancel() 

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
		Call DBQuery
		Exit Function
	End If
    Call parent.FncExport(Parent.C_MULTI)
Exit Function

	Dim iExcelRow, arrData(), ret, iMaxRows, iMaxCols
	Dim iRow, iCol, iColHeadCnt
	Dim xlApp 'As Excel.Application
	Dim xlBook 'As Excel.Workbook
	Dim xlSheet 'As Excel.Worksheet	

	Set xlApp = CreateObject("excel.application")	
	If Err.Number <> 0 Then
	    Msgbox Err.Number & " : " & Err.Description
	    Exit Function
	End If

	' -- 워크북 
	'Set xlApp = New Excel.Application
	Set xlBook = xlApp.Workbooks.Add
	Set xlSheet = xlBook.Worksheets.Add

	With frm1.vspdData
	' -- 보이기 
	xlApp.Visible = True
	
	' -- 제목 찍기 
	xlSheet.Cells(2, 1 ).value = document.title 
	xlSheet.Cells(2, 1 ).Font.Size = 25

	xlSheet.Cells(5, 1 ).value = document.title 
	xlSheet.Cells(5, 1 ).Font.Size = 17

	' -- 조건절 찍기 
	iExcelRow = 10
	xlSheet.Cells(iExcelRow, 1 ).value = "기준년월"
	xlSheet.Cells(iExcelRow, 2 ).value = frm1.txtSTART_DT.Text
	xlSheet.Cells(iExcelRow, 5 ).value = "공장"
	xlSheet.Cells(iExcelRow, 6 ).value = frm1.txtPLANT_CD.value
	xlSheet.Cells(iExcelRow, 7 ).value = frm1.txtPLANT_NM.value

	iExcelRow = iExcelRow + 1
	xlSheet.Cells(iExcelRow, 1 ).value = "품목계정"
	xlSheet.Cells(iExcelRow, 2 ).value = frm1.txtITEM_ACCT.value
	xlSheet.Cells(iExcelRow, 3 ).value = frm1.txtITEM_ACCT_NM.value
	xlSheet.Cells(iExcelRow, 5 ).value = "품목"
	xlSheet.Cells(iExcelRow, 6 ).value = frm1.txtITEM_CD.value
	xlSheet.Cells(iExcelRow, 7 ).value = frm1.txtITEM_NM.value

	' -- 그리드 헤더 찍기 
	iExcelRow = iExcelRow + 3
	
	iColHeadCnt = .ColHeaderRows
	
	For iRow = 0 To iColHeadCnt - 1
		.Row = -1000 + iRow
		
		For iCol = 1 To .MaxCols 
			.Col = iCol
			
			'xlSheet.Cells(iExcelRow + iRow, iCol ).value = .Text
		Next
	Next

	iExcelRow = iExcelRow + 1
	
	iMaxRows = CLng(.MaxRows)
	iMaxCols = CLng(.MaxCols)
	
	ReDim arrData(iMaxRows, iMaxCols)
	ret = .GetArray(1, 0, arrData)

	xlSheet.Range(xlApp.Cells(iExcelRow, 1), xlApp.Cells(iExcelRow+iMaxRows, iMaxCols)).Value = arrData	
	
	
	End With    
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
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

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


		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtStartDt=" & sStartDt
		strVal = strVal & "&txtEndDt=" & sEndDt	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey


		If lgIntFlgMode = Parent.OPMD_CMODE Then
			strVal = strVal & "&txtPLANT_CD=" & Trim(.txtPLANT_CD.value)
			strVal = strVal & "&txtITEM_ACCT=" & Trim(.txtITEM_ACCT.value)
			strVal = strVal & "&txtITEM_CD=" & Trim(.txtITEM_CD.value)
		
			If .rdoTYPE1.checked then
				strVal = strVal & "&rdoTYPE=1"
				
				Call InitSpreadSheet
			Else
				strVal = strVal & "&rdoTYPE=2"
				
				Call InitSpreadSheet2
			End If
		
		Else
			strVal = strVal & "&txtPLANT_CD=" & Trim(.hPLANT_CD.value)
			strVal = strVal & "&txtITEM_ACCT=" & Trim(.hITEM_ACCT.value)
			strVal = strVal & "&txtITEM_CD=" & Trim(.hITEM_CD.value)
			strVal = strVal & "&rdoTYPE=" & Trim(.hTYPE.value)
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

	Dim pObj

	' -- 그리드를 2개로 늘리면서 취해진 소스변경 
	If frm1.rdoTYPE1.checked then
		Frm1.vspdData.Focus
	Else
		Frm1.vspdData2.Focus
	End If
	
   	
   	If frm1.rdoTYPE2.checked And lgIntFlgMode = Parent.OPMD_CMODE then
   		Call ReInitSpreadSheet2
   	End If
   	

    Set gActiveElement = document.ActiveElement   

	lgIntFlgMode = Parent.OPMD_UMODE
	
	window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"

	If lgStrPrevKey = "*" Then
		lgStrPrevKey = ""
		Call FncExcel() 
	End If
	
End Function

'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor(Byval pGrpRow)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt

	Dim pObj

	' -- 그리드를 2개로 늘리면서 취해진 소스변경 
	If frm1.rdoTYPE1.checked then
		Set pObj = frm1.vspdData
	Else
		Set pObj = frm1.vspdData2
	End If

	' -- 그리드 1 정의 
	With pObj

	.ReDraw = False
	arrRow = Split(pGrpRow, Parent.gRowSep)
	
	iLoopCnt = UBound(arrRow, 1)
	
	For i = 0 to iLoopCnt -1
		arrCol = Split(arrRow(i), Parent.gColSep)
	
		.Col = -1
		.Row = CDbl(arrCol(1))	' -- 행 
		
		Select Case arrCol(0)
			Case "1"
			   ret = .AddCellSpan(4, .Row , 2, 1)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
			Case "2"
				ret = .AddCellSpan(2, .Row , 4, 1)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
			Case "3"
				ret = .AddCellSpan(1, .Row , 5, 1)
				.BackColor = RGB(204,255,255) 
				.ForeColor = vbBlack
			Case "4"  
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
'check plant
	If Trim(frm1.txtPLANT_CD.value) <> "" Then
		strWhere = " plant_cd= " & FilterVar(frm1.txtPLANT_CD.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	 b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPLANT_CD.alt,"X")
			frm1.txtPLANT_CD.focus 
			frm1.txtPLANT_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPLANT_NM.value = strDataNm(0)
	Else
		frm1.txtPLANT_NM.value=""
	End If
'check item_acct
	If Trim(frm1.txtITEM_ACCT.value) <> "" Then
		strWhere = " minor_cd  = " & FilterVar(frm1.txtITEM_ACCT.value, "''", "S") & " "		
		strWhere = strWhere & "		and major_cd=" & filterVar("P1001","","S")
		
		Call CommonQueryRs(" minor_nm  ","	 b_minor ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_ACCT.alt,"X")
			frm1.txtITEM_ACCT.focus 
			frm1.txtITEM_ACCT_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_ACCT_NM.value = strDataNm(0)
	ELSE
		frm1.txtITEM_ACCT_NM.value=""
	End If
'check item
	If Trim(frm1.txtITEM_CD.value) <> "" Then
		strFrom = " B_ITEM a left outer join dbo.b_item_by_plant b on a.item_cd = b.item_cd "	

		strWhere = " a.item_cd  = " & FilterVar(frm1.txtITEM_CD.value, "''", "S") & " "		
		If frm1.txtPLANT_CD.value<>"" then 
			strWhere = strWhere & "		and b.plant_cd=" & filterVar(frm1.txtPLANT_CD.value,"","S")
		End If
		
		Call CommonQueryRs(" a.item_nm  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_CD.alt,"X")
			frm1.txtITEM_CD.focus 
			frm1.txtITEM_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_NM.value = strDataNm(0)
	ELSE
		frm1.txtITEM_NM.value=""
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
									<TD CLASS="TD5">기준년월</TD>
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSTART_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 기준년월" tag="12" id=txtSTART_DT></OBJECT>');</SCRIPT>&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtEND_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="종료 기준년월" tag="12" id=txtEND_DT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtPLANT_CD" TYPE="Text" MAXLENGTH="4" tag="15XXXU" size="20" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPLANT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_ACCT" TYPE="Text" MAXLENGTH="2" tag="15XXU" size="10" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)">
									<input NAME="txtITEM_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_CD" TYPE="Text" MAXLENGTH="18" tag="15XXXU" size="20" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(2)">
									<input NAME="txtITEM_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">출력구분</TD>
									<TD CLASS="TD6" valign=top><INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE1 tag="15XXX" checked><LABEL FOR="rdoTYPE1">집계</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoTYPE ID=rdoTYPE2 tag="15XXX"><LABEL FOR="rdoTYPE2">상세</LABEL>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>    
							</TABLE>						</FIELDSET>
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hSTART_DT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hEND_DT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPLANT_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hITEM_ACCT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hITEM_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTYPE" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

