<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :CC별원가분석(월별) 
'*  3. Program ID           : c4229ma1.asp
'*  4. Program Name         : CC별원가분석(월별)
'*  5. Program Desc         : CC별원가분석(월별)
'*  6. Modified date(First) : 2005-11-18
'*  7. Modified date(Last)  : 2005-11-18
'*  8. Modifier (First)     : HJO
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4229mb1.asp"                               'Biz Logic ASP

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3																		'☜: Tab의 위치 
Const TAB4 = 4
Const TAB5 = 5																		'☜: Tab의 위치 


Dim iDBSYSDate
Dim iStrFromDt
Dim iStrToDt

iDBSYSDate = "<%=GetSvrDate%>"
iStrToDt = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)	
iStrFromDt= UNIDateAdd("m", -1,iStrToDt, parent.gServerDateFormat)
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
Dim lgStrPrevKey2
Dim lgSTime		' -- 디버깅 타임체크 
Dim  gSelframeFlg
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================

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

	frm1.txtFrom_YYYYMM.Text =UniConvDateAToB(iStrFromDt, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtTo_YYYYMM.text =UniConvDateAToB(iStrToDt, parent.gServerDateFormat, parent.gDateFormat)
	
	
	Call ggoOper.FormatDate(frm1.txtFrom_YYYYMM, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtTo_YYYYMM, parent.gDateFormat, 2)
	
	frm1.txtCost_cd.focus 
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
	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "MA") %>
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
Sub InitSpreadSheet(byVal iTab, byVal iMaxCols)
	Dim i, ret
	
	
    'Call AppendNumberPlace("6","3","0")
    '--------------TAB1
    SELECT CASE ITAB
		CASE TAB1
			With frm1.vspdData
		
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False

			.MaxRows = 0
			.MaxCols = iMaxCols		
			'.ColHidden = True

			'헤더를 2줄로    
			.ColHeaderRows = 2

			.Col = -1: .Row = -1000: .RowMerge = 1
			.Col = 1: .Row = -1: .ColMerge = 1
			.Col = 2: .Row = -1: .ColMerge = 1
			
			ggoSpread.SSSetEdit		1,	"작업지시C/C"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"작업지시C/C명"	, 18		


			For i = 3 To iMaxCols	step 6
					ggoSpread.SSSetFloat	i,		""		, 13,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+1,		""		, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+2,		""		, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+3,		""		, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+4,		""		, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+5,		""	, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					
			Next		
			For i = 1 To 2
				ret = .AddCellSpan(i, -1000 , 1, 2)
			Next

			.Col = iMaxCols		
			.ColHidden = True
			
			.rowheight(-1000) = 12	' 높이 재지정 
			.rowheight(-999) = 12	' 높이 재지정 
			
			ggoSpread.SSSetSplit2(2) 
			.ReDraw = True		
			End With
		
			ggoSpread.SpreadLockWithOddEvenRowColor()
    
		CASE TAB2
			With frm1.vspdData2
		
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False


			.MaxRows = 0
			.MaxCols = iMaxCols
			.ColHeaderRows = 3
		
			.Col = -1: .Row = -1000: .RowMerge = 1
			
			.Col = 1: .Row = -1: .ColMerge = 1
			.Col = 2: .Row = -1: .ColMerge = 1
			.Col = 3: .Row = -1: .ColMerge = 1
			.Col = 4: .Row = -1: .ColMerge = 1
			.Col = 5: .Row = -1: .ColMerge = 1
		
			ggoSpread.SSSetEdit		1,	"원가요소"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"원가요소명", 18
			ggoSpread.SSSetEdit		3,	"계정"	, 10,,,,1
			ggoSpread.SSSetEdit		4,	"계정명"	, 18
			ggoSpread.SSSetEdit		5,	"구분"	, 8,2


			For i = 6 To iMaxCols						
					ggoSpread.SSSetFloat	i,	""	, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			Next		
			
			For i = 1 To 5
				ret = .AddCellSpan(i, -1000 , 1, 3)
			Next

			.rowheight(-1000) = 12	' 높이 재지정 
			.rowheight(-999) = 12	' 높이 재지정 
			.rowheight(-998) = 12	' 높이 재지정 
			
			.Col = iMaxCols		
			.ColHidden = True
			.ReDraw = True		
			End With
		
			ggoSpread.SpreadLockWithOddEvenRowColor()
    

		
		CASE TAB3
		
			With frm1.vspdData3
		
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False


			.MaxRows = 0
			.MaxCols =iMaxCols
			.ColHeaderRows = 2

			.Col = -1: .Row = -1000: .RowMerge = 1
			.Col = 1: .Row = -1: .ColMerge = 1
			.Col = 2: .Row = -1: .ColMerge = 1
			
		
			ggoSpread.SSSetEdit		1,	"작업지시C/C"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"작업지시C/C명"	, 18	


			For i = 3 To iMaxCols 	STEP 3
					ggoSpread.SSSetFloat	i,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+1,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+2,	""	, 15,		Parent.ggExchRateNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			Next		
			For i = 1 To 2
				ret = .AddCellSpan(i, -1000 , 1, 2)
			Next

			.rowheight(-1000) = 12	' 높이 재지정 
			.rowheight(-999) = 17	' 높이 재지정 
			.Col = iMaxCols		
			.ColHidden = True
			.ReDraw = True		
			End With
		
			ggoSpread.SpreadLockWithOddEvenRowColor()
    

		
		CASE TAB4
			With frm1.vspdData4
		
			ggoSpread.Source = frm1.vspdData4
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False


			.MaxRows = 0
			.MaxCols = iMaxCols
		
			ggoSpread.SSSetEdit		1,	"작업지시C/C"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"작업지시C/C명"	, 18


			For i = 3 To iMaxCols	
					ggoSpread.SSSetFloat	i,	""	, 13,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			Next		

			.Col = iMaxCols		
			.ColHidden = True
			.ReDraw = True		
			End With
		
			ggoSpread.SpreadLockWithOddEvenRowColor()
    

		CASE TAB5
			With frm1.vspdData5
		
			ggoSpread.Source = frm1.vspdData5
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False


			.MaxRows = 0
			.MaxCols = iMaxCols
			.ColHeaderRows = 2

			.Col = -1: .Row = -1000: .RowMerge = 1
			.Col = 1: .Row = -1: .ColMerge = 1
			.Col = 2: .Row = -1: .ColMerge = 1
			.Col = 3: .Row = -1: .ColMerge = 1
			.Col = 4: .Row = -1: .ColMerge = 1
		
			ggoSpread.SSSetEdit		1,	"작업지시C/C"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"작업지시C/C명"	, 18
			ggoSpread.SSSetEdit		3,	"배부요소"	, 10,,,,1
			ggoSpread.SSSetEdit		4,	"배부요소명"	, 18	


			For i = 5 To iMaxCols 	STEP 3				
					ggoSpread.SSSetFloat	i,	""	, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+1,	""	, 13,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+2,	""	, 13,		Parent.ggExchRateNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					
			Next		
			For i = 1 To 4
				ret = .AddCellSpan(i, -1000 , 1, 2)
			Next

			.rowheight(-1000) = 12	' 높이 재지정 
			.rowheight(-999) = 12	' 높이 재지정 
			.Col = iMaxCols		
			.ColHidden = True
			.ReDraw = True		
			End With		
			ggoSpread.SpreadLockWithOddEvenRowColor()		
	END SELECT 	
End Sub



'========================================================================================
' Function Name : SetGridHead
' Function Desc : set grid head row
'========================================================================================
Sub SetGridHead(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol
	
	Select Case 	gSelframeFlg
		CASE TAB1
		' -- 그리드 1 정의 
			With frm1.vspdData			
				arrRows = Split(pData, Parent.gRowSep)
				'헤더를 ?줄로    
				'.ColHeaderRows = UBound(arrRows, 1)
				For i = 0 To UBound(arrRows, 1) -1
					arrCols = Split(arrRows(i), Parent.gColSep)
					iColCnt = UBound(arrCols, 1)
					.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
					iCol =3
					For j = 0 To iColCnt 
						.Col = iCol
						Select Case j
							Case 0, 1,  2,iColCnt
								.Text = arrCols(j)
								iCol = iCol + 1
							Case Else
								.Text = arrCols(j)
								 iCol = iCol + 1	: .Col = iCol	' -- 금액 

						End SElect									
					Next
				Next
			End With		
		CASE TAB2	
			With frm1.vspdData2			
				arrRows = Split(pData, Parent.gRowSep)
				'헤더를 ?줄로    
				'.ColHeaderRows = UBound(arrRows, 1)
				For i = 0 To UBound(arrRows, 1) -1
					arrCols = Split(arrRows(i), Parent.gColSep)
					iColCnt = UBound(arrCols, 1)
					.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
					iCol =6
					For j = 0 To iColCnt 
						.Col = iCol
						Select Case j
							Case 0, 1,  2,3,4,5,iColCnt
								.Text = arrCols(j)
								iCol = iCol + 1
							Case Else
								.Text = arrCols(j)
								 iCol = iCol + 1	: .Col = iCol	' -- 금액 
						End SElect									
					Next
				Next
			End With
	
		CASE TAB3
	
			' -- 그리드 1 정의 
			With frm1.vspdData3
				
				arrRows = Split(pData, Parent.gRowSep)

				'헤더를 ?줄로    
				'.ColHeaderRows = UBound(arrRows, 1)

				For i = 0 To UBound(arrRows, 1) -1
					arrCols = Split(arrRows(i), Parent.gColSep)
					iColCnt = UBound(arrCols, 1)
					.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
					iCol =3
					For j = 0 To iColCnt 
						.Col = iCol
						Select Case j
							Case 0, 1,  2,iColCnt
								.Text = arrCols(j)
								iCol = iCol + 1
							Case Else
								.Text = arrCols(j)
								 iCol = iCol + 1	: .Col = iCol	' -- 금액 
						End SElect									
					Next
				Next
			End With
		CASE TAB4	
			With frm1.vspdData4			
				arrRows = Split(pData, Parent.gRowSep)
				'헤더를 ?줄로    
				'.ColHeaderRows = UBound(arrRows, 1)

				For i = 0 To UBound(arrRows, 1) -1
					arrCols = Split(arrRows(i), Parent.gColSep)
					iColCnt = UBound(arrCols, 1)
					.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
					iCol =3
					For j = 0 To iColCnt 
						.Col = iCol
						Select Case j
							Case 0, 1,  2,iColCnt
								.Text = arrCols(j)
								iCol = iCol + 1
							Case Else
								.Text = arrCols(j)
								 iCol = iCol + 1	: .Col = iCol	' -- 금액 
						End SElect									
					Next
				Next
			End With	
		CASE TAB5
				With frm1.vspdData5
					
					arrRows = Split(pData, Parent.gRowSep)
					'헤더를 ?줄로    
					'.ColHeaderRows = UBound(arrRows, 1)

					For i = 0 To UBound(arrRows, 1) -1
						arrCols = Split(arrRows(i), Parent.gColSep)
						iColCnt = UBound(arrCols, 1)
						.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
						iCol =5
						For j = 0 To iColCnt 
							.Col = iCol
							Select Case j
								Case 0, 1,  2,3,4, iColCnt
									.Text = arrCols(j)
									iCol = iCol + 1
								Case Else
									.Text = arrCols(j)
									 iCol = iCol + 1	: .Col = iCol	' -- 금액 
							End SElect									
						Next
					Next
			End With	
	END SELECT
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
	If Not chkField(Document, "1") Then
			   Exit Function
	End If
	
	With frm1
	
	arrParam(1) = " B_COST_CENTER "
	
	Select Case iWhere
		Case 0
			arrParam(0) = "작업지시C/C 팝업"
			arrParam(2) = Trim(.txtCost_cd.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "작업지시C/C" 

			arrField(0) ="ED10" & parent.gColsep &  "COST_CD"	
			arrField(1) ="ED20" & parent.gColsep &  "COST_NM"    
			arrHeader(0) = "작업지시C/C"	
			arrHeader(1) = "작업지시C/C명"

			
	
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
				.txtCost_cd.value		= arrRet(0)
				.txtCost_nm.value		= arrRet(1)				
				.txtCost_cd.focus
		
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
	
	Call ggoOper.FormatDate(frm1.txtFrom_YYYYMM,   parent.gDateFormat,2)
	Call ggoOper.FormatDate(frm1.txtTo_YYYYMM, parent.gDateFormat,2)
	
	Call SetDefaultVal

   
   Call ClickTab1()	
	 gIsTab     = "Y" 
	 gTabMaxCnt = 5

   Call SetToolbar("110000000001111")	
   
	    
   	
    
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


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub txtFrom_YYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub  txtFrom_YYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrom_YYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrom_YYYYMM.Focus
    End If
End Sub
Sub txtTo_YYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub  txtTo_YYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrom_YYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtTo_YYYYMM.Focus
    End If
End Sub


Sub txtCost_cd_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub
Sub txtCost_cd_onChange()
	If frm1.txtCost_cd.value ="" then frm1.txtCost_nm.value=""
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    'ggoSpread.Source = frm1.vspdData
    'Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
        End With

    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정		
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData         
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
   
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
   
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub
Sub vspdData4_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData4.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData4,NewTop) And lgStrPrevKey <> "" Then
    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub
Sub vspdData5_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData5.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData5,NewTop) And lgStrPrevKey <> "" Then
  
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
    
    sStartDt= Replace(frm1.txtFrom_YYYYMM.text, parent.gComDateType, "")
    sEndDt= Replace(frm1.txtTo_YYYYMM.text, parent.gComDateType, "")

	If ValidDateCheck(frm1.txtFrom_YYYYMM, frm1.txtTo_YYYYMM) = False Then 
		frm1.txtFrom_YYYYMM.focus 
		Exit Function
	End If
	
    IF ChkKeyField()=False Then Exit Function 

    
    Call ggoOper.ClearField(Document, "2")
    
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
    Call InitSpreadSheet(gActiveSpdSheet.id)      
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	
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
	Dim strVal, strNext

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    Dim sStartDt, sEndDt, sYear, sMon, sDay
    
    With frm1
		Call parent.ExtractDateFromSuper(.txtFrom_YYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)	
		sStartDt= (sYear&sMon)
		Call parent.ExtractDateFromSuper(.txtTo_YYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)
		sEndDt=sYear&sMon
		

		'If gSelframeFlg=TAB1 THEN 
			strNext=lgStrPrevKey		
		'Else
		'	strNext=lgStrPrevKey2			
		'End If

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & strNext
			strVal = strVal & "&txtFrom_YYYYMM=" & Trim(.hYYYYMM.value)
			strVal = strVal & "&txtTo_YYYYMM=" & Trim(.hYYYYMM2.value)				
			strVal = strVal & "&txtCost_cd=" & Trim(.hCost_cd.value)
			strVal = strVal & "&txtFrame=" & gSelframeFlg

		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & strNext
			strVal = strVal & "&txtFrom_YYYYMM=" & sStartDt
			strVal = strVal & "&txtTo_YYYYMM=" & sEndDt				
			strVal = strVal & "&txtCost_cd=" & Trim(.txtCost_cd.value)
			strVal = strVal & "&txtFrame=" & gSelframeFlg
		End If


		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	lgIntFlgMode = parent.OPMD_UMODE	

	SELECT CASE gSelframeFlg
	CASE TAB1 
		frm1.vspdData.style.display = ""	'-- 그리드 보이게..	
		Frm1.vspdData.Focus
	CASE TAB2

		frm1.vspdData2.style.display = ""	'-- 그리드 보이게..	
		Frm1.vspdData2.Focus
	CASE TAB3

		frm1.vspdData3.style.display = ""	'-- 그리드 보이게..	
		Frm1.vspdData3.Focus
	CASE TAB4
		frm1.vspdData4.style.display = ""	'-- 그리드 보이게..	
		Frm1.vspdData4.Focus
	CASE TAB5
		frm1.vspdData5.style.display = ""	'-- 그리드 보이게..	
		Frm1.vspdData5.Focus
	END SELECT 

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor(byVal arrStr)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt, strRowI
	
	Select case  gSelframeFlg
		case TAB1 
			 With frm1.vspdData
		
			.ReDraw = False
			
			arrRow = Split(arrStr, Parent.gRowSep)
			
			iLoopCnt = UBound(arrRow, 1)

			For i = 0 to iLoopCnt -1
				arrCol = Split(arrRow(i), Parent.gColSep)
			
				.Col = -1
				.Row = CDbl(arrCol(2))	' -- 행 
			
				Select Case arrCol(0)
					Case "%1"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1)  : .Col2=.MaxCols
						.BlockMode = True
					   ret = .AddCellSpan(1, iRow ,2, 1)
						.BackColor = RGB(250,250,210) 
						.ForeColor = vbBlack
						.BlockMode = False

				End Select
				.Col = 1: .Row = -1: .ColMerge = 1
				.Col = 2: .Row = -1: .ColMerge = 1
				strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
			Next

			frm1.txtTmp.value=frm1.txtTmp.value & strRowI
			.ReDraw = True
			End With
	CASE TAB2
		 With frm1.vspdData2
		
			.ReDraw = False
			
			arrRow = Split(arrStr, Parent.gRowSep)
			
			iLoopCnt = UBound(arrRow, 1)

			For i = 0 to iLoopCnt -1
				arrCol = Split(arrRow(i), Parent.gColSep)
			
				.Col = -1
				.Row = CDbl(arrCol(2))	' -- 행 
		
				Select Case arrCol(0)
					Case "%1"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(1, iRow ,4, 3)
						.BackColor = RGB(250,250,210) 
						.ForeColor = vbBlack
						.BlockMode = False
					Case "%2"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1)  : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(3, iRow,2, 3)
						.BackColor = RGB(204,255,153) 
						.ForeColor = vbBlack
						.BlockMode = False
				End Select
				.Col = 1: .Row = -1: .ColMerge = 1
				.Col = 2: .Row = -1: .ColMerge = 1
				.Col = 3: .Row = -1: .ColMerge = 1
				.Col = 4: .Row = -1: .ColMerge = 1
				.Col = 5: .Row = -1: .ColMerge = 1

				strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
			Next

			frm1.txtTmp.value=frm1.txtTmp.value & strRowI
			.ReDraw = True
			End With
	CASE TAB3
		 With frm1.vspdData3
		
			.ReDraw = False
			
			arrRow = Split(arrStr, Parent.gRowSep)
			
			iLoopCnt = UBound(arrRow, 1)

			For i = 0 to iLoopCnt -1
				arrCol = Split(arrRow(i), Parent.gColSep)
			
				.Col = -1
				.Row = CDbl(arrCol(2))	' -- 행 
			
				Select Case arrCol(0)
					Case "%1"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(1, iRow ,2, 1)
						.BackColor = RGB(250,250,210) 
						.ForeColor = vbBlack
						.BlockMode = False
			
				End Select
				.Col = 1: .Row = -1: .ColMerge = 1
				.Col = 2: .Row = -1: .ColMerge = 1

				strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
			Next

			frm1.txtTmp.value=frm1.txtTmp.value & strRowI
			.ReDraw = True
			End With
	
	CASE TAB4
		 With frm1.vspdData4
		
			.ReDraw = False
			
			arrRow = Split(arrStr, Parent.gRowSep)
			
			iLoopCnt = UBound(arrRow, 1)

			For i = 0 to iLoopCnt -1
				arrCol = Split(arrRow(i), Parent.gColSep)
			
				.Col = -1
				.Row = CDbl(arrCol(2))	' -- 행 
			
				Select Case arrCol(0)
					Case "%1"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(1, iRow ,2, 1)
						.BackColor = RGB(250,250,210) 
						.ForeColor = vbBlack
						.BlockMode = False

	'				
				End Select
				.Col = 1: .Row = -1: .ColMerge = 1
				.Col = 2: .Row = -1: .ColMerge = 1

				strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
			Next

			frm1.txtTmp.value=frm1.txtTmp.value & strRowI
			.ReDraw = True
			End With
	
	CASE TAB5
		 	 With frm1.vspdData5
		
			.ReDraw = False
			
			arrRow = Split(arrStr, Parent.gRowSep)
			
			iLoopCnt = UBound(arrRow, 1)

			For i = 0 to iLoopCnt -1
				arrCol = Split(arrRow(i), Parent.gColSep)
			
				.Col = -1
				.Row = CDbl(arrCol(2))	' -- 행 
			
				Select Case arrCol(0)
					Case "%1"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(1, iRow ,2, 1)
						.BackColor = RGB(250,250,210) 
						.ForeColor = vbBlack
						.BlockMode = False

					Case "%2"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1)  : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(3, iRow,2, 1)
						.BackColor = RGB(204,255,153) 
						.ForeColor = vbBlack
						.BlockMode = False
				End Select
	
				.Col = 1: .Row = -1: .ColMerge = 1
				.Col = 2: .Row = -1: .ColMerge = 1
				.Col = 3: .Row = -1: .ColMerge = 1
				.Col = 4: .Row = -1: .ColMerge = 1

				strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
			Next

			frm1.txtTmp.value=frm1.txtTmp.value & strRowI
			.ReDraw = True
			End With
	
	
	End SELECT

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
'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
   
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
 '   Call InitSpreadSheet(gSelframeFlg)
 
	frm1.vspdData.style.display="none"
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables 
End Function

Function ClickTab2()
   
	Call changeTabs(TAB2)	 
	gSelframeFlg = TAB2
	frm1.vspdData2.style.display="none"
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	Call InitVariables

End Function

Function ClickTab3()
   
	Call changeTabs(TAB3)	 
	gSelframeFlg = TAB3
	frm1.vspdData3.style.display="none"
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

	Call InitVariables

End Function
Function ClickTab4()
   
	Call changeTabs(TAB4)	 
	gSelframeFlg = TAB4
	frm1.vspdData4.style.display="none"
	ggoSpread.Source = frm1.vspdData4
	ggoSpread.ClearSpreadData

	Call InitVariables
  
End Function
Function ClickTab5()
   
	Call changeTabs(TAB5)	 
	gSelframeFlg = TAB5
	frm1.vspdData5.style.display="none"
	ggoSpread.Source = frm1.vspdData5
	ggoSpread.ClearSpreadData

	Call InitVariables
End Function

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
'check COST CD
	If Trim(frm1.txtCost_cd.value) <> "" Then
		strWhere = " cost_cd = " & FilterVar(frm1.txtCost_cd.value, "''", "S") & "  "

		Call CommonQueryRs(" Cost_nm ","	 b_cost_center  ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCost_cd.alt,"X")			
			frm1.txtCost_nm.value = ""
			ChkKeyField = False
			frm1.txtCost_cd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtCost_nm.value = strDataNm(0)
	Else
		frm1.txtCost_nm.value=""
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
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>C/C별제품제조원가</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>C/C별투입원가</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>C/C별매출이익율</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab4()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>C/C별불량현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab5()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>C/C별배부요소DATA</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;
					</TD>
					
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
									<TD CLASS="TD6" valign=top> <TABLE>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFrom_YYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 작업년월" tag="12" id=txtFrom_YYYYMM></OBJECT>');</SCRIPT>
												</TD>
												<TD>~</TD>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtTo_YYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="종료 작업년월" tag="12" id=txtTo_YYYYMM></OBJECT>');</SCRIPT>	
												
												</TD>
											</TR>
										 </TABLE>
									</TD>
									<TD CLASS="TD5" NOWRAP>작업지시C/C</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtCost_cd" TYPE="Text" MAXLENGTH="10" tag="15XXX" size="10" ALT="작업지시C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtCost_nm" TYPE="TEXT"  tag="14XXX" size="25">
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

						<DIV ID="TabDiv" SCROLL=no style="display:none;">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" SCROLL=no style="display:none;">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" SCROLL=no style="display:none;">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData3 NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" SCROLL=no style="display:none;">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData4 NAME=vspdData4 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
						</DIV>
												<DIV ID="TabDiv" SCROLL=no style="display:none;">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData5 NAME=vspdData5 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO  noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM2" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCost_cd" tag="24" TABINDEX= "-1">

<INPUT TYPE=HIDDEN NAME="txtTmp" tag="24" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

