<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 오더-공정별원가조회 
'*  3. Program ID           : c4207ma1.asp
'*  4. Program Name         : 오더-공정별원가조회 
'*  5. Program Desc         : 오더-공정별원가조회 
'*  6. Modified date(First) : 2005-10-10
'*  7. Modified date(Last)  : 
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

Const BIZ_PGM_ID = "c4208mb1.asp"                               'Biz Logic ASP

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
Dim lgErrRow, lgErrCol, lgOptionFlag, lgEOF

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
	Dim i, ret, iBas
	With frm1.vspdData
		
		.style.display = ""

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030825",, "" ', ,Parent.gAllowDragDropSpread
			
		iBas = 9	' -- 앞에 고정이 변할경우 대비 

		.MaxRows = 0
		.MaxCols = iBas+30		' -- Group/RowNum 

		.Col  = iBas+29 : .ColHidden = True
		.Col  = iBas+30 : .ColHidden = True
		
		'헤더를 2줄로    
		.ColHeaderRows = 8

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetEdit		1,	"공장"	, 6,,,,1	
		ggoSpread.SSSetEdit		2,	"품목계정"	, 5,,,,1	
		ggoSpread.SSSetEdit		3,	"조달구분"	, 8,,,,1	
		ggoSpread.SSSetEdit		4,	"품목"	, 8,,,,1
		ggoSpread.SSSetEdit		5,	"품목명"	, 10		
		ggoSpread.SSSetEdit		6,	"오더번호"	, 15,,,,1
		ggoSpread.SSSetEdit		7,	"공순"	, 5,,,,1
		ggoSpread.SSSetEdit		8,	"공정"	, 7,,,,1
		ggoSpread.SSSetEdit		9,	"공정명"	, 10				
		
		ggoSpread.SSSetFloat	iBas+1,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+2,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+3,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+4,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+5,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+6,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+7,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+8,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+9,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+10,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+11,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+12,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+13,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+14,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+15,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+16,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+17,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+18,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+19,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+20,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+21,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+22,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+23,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+24,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+25,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+26,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+27,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+28,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		'Call ggoSpread.SSSetColHidden(36,37,True)

		For i = 1 To iBas
			ret = .AddCellSpan(i, -1000, 1, 8)
		Next
		
		For i = iBas +1 To iBas + 28 Step 4
			Select Case i
				Case iBas + 9, iBas + 17
					ret = .AddCellSpan(i, -1000, 8, 1)	' -- 기초재공/전공정대재(당기착수) 등등 
				Case Else
					ret = .AddCellSpan(i, -1000, 4, 1)	' -- 기초재공/전공정대재(당기착수) 등등 
			End Select

			Select Case i
				Case iBas + 1, iBas + 25
					ret = .AddCellSpan(i, -999, 1, 2)	
				Case Else
					ret = .AddCellSpan(i, -999, 1, 7)	
			End Select
			
			ret = .AddCellSpan(i+1, -999, 3, 1)
			ret = .AddCellSpan(i+1, -998, 3, 1)

			If i = iBas + 1 Or i = iBas + 25 Then			
				ret = .AddCellSpan(i, -997, 1, 5)
			End If
			
			ret = .AddCellSpan(i+1, -997, 3, 1)
			ret = .AddCellSpan(i+1, -996, 3, 1)
			ret = .AddCellSpan(i+1, -995, 3, 1)
			ret = .AddCellSpan(i+1, -994, 3, 1)
			' -- 마지막 행은 Span이 없다.
		Next
		
		' 1번째 헤더 출력 글자 
		.Row = -1000
		.Col = iBas+1	: .Text = "기초재공"
		.Col = iBas+5	: .Text = "전공정대체(당기착수)"
		.Col = iBas+9	: .Text = "차공정대체"
		.Col = iBas+17	: .Text = "완성대체"
		.Col = iBas+25	: .Text = "기말재공"
		
		' 2번째 헤더 출력 글자 
		.Row = -999
		.Col = iBas+1	: .Text = "기초재공수량"
		.Col = iBas+2	: .Text = "기초재공금액"
		.Col = iBas+5	: .Text = "전공정대체(당기착수수량)" 
		.Col = iBas+6	: .Text = "전공정대체(당기착수)금액"
		.Col = iBas+9	: .Text = "차공정대체(기초재공분)수량"
		.Col = iBas+10	: .Text = "차공정대체(기초재공분)금액"
		.Col = iBas+13	: .Text = "차공정대체(당기착수분)수량"
		.Col = iBas+14	: .Text = "차공정대체(당기착수분)금액"
		.Col = iBas+17	: .Text = "완성대체(기초재공분)수량"
		.Col = iBas+18	: .Text = "완성대체(기초재공분)금액"
		.Col = iBas+21	: .Text = "완성대체(당기재공분)수량"
		.Col = iBas+22	: .Text = "완성대체(당기재공분)금액"
		.Col = iBas+25	: .Text = "기말재공수량"
		.Col = iBas+26	: .Text = "기말재공금액"


		' 3번째 헤더 출력 글자 
		.Row = -998
		For i = iBas +1 To iBas+28 Step 4
			.Col = i+1	: .Text = "재료비"
		Next

		' 4번째 헤더 출력 글자 
		.Row = -997
		.Col = iBas +1	: .Text = "기초재공환산수량"
		.Col = iBas +25	: .Text = "기말재공환산수량"

		For i = iBas +1 To iBas+28 Step 4
			.Col = i+1	: .Text = "반제품비"
		Next

		' 5번째 헤더 출력 글자 
		.Row = -996
		For i = iBas +1 To iBas+28 Step 4
			.Col = i+1	: .Text = "노무비"
		Next

		' 6번째 헤더 출력 글자 
		.Row = -995
		For i = iBas +1 To iBas+28 Step 4
			.Col = i+1	: .Text = "경비"
		Next

		' 7번째 헤더 출력 글자 
		.Row = -994
		For i = iBas +1 To iBas+28 Step 4
			.Col = i+1	: .Text = "외주가공비"
		Next

		' 8번째 헤더 출력 글자 
		.Row = -993
		For i = iBas +1 To iBas+28 Step 4
			.Col = i+1	: .Text = "합계"
			.Col = i+2	: .Text = "전공정원가"
			.Col = i+3	: .Text = "당공정(투입원가)"
		Next
		
		
		.rowheight(-993) = 20	' 높이 재지정 
		
		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		.Col = 5: .Row = -1: .ColMerge = 1
		.Col = 6: .Row = -1: .ColMerge = 1
		.Col = 7: .Row = -1: .ColMerge = 1
		.Col = 8: .Row = -1: .ColMerge = 1
		.Col = 9: .Row = -1: .ColMerge = 1

		ggoSpread.SSSetSplit2(iBas)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		
	End With
End Sub

Sub InitSpreadSheet2()
	Dim i, ret, iBas
	With frm1.vspdData
		
		.style.display = ""

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030825" ,, "" ', ,Parent.gAllowDragDropSpread
			
		iBas = 9	' -- 앞에 고정이 변할경우 대비 

		.MaxRows = 0
		.MaxCols = iBas+22		' -- Group/RowNum 

		.Col  = iBas+21 : .ColHidden = True
		.Col  = iBas+22 : .ColHidden = True
		
		'헤더를 2줄로    
		.ColHeaderRows = 8

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetEdit		1,	"공장"	, 6,,,,1	
		ggoSpread.SSSetEdit		2,	"품목계정"	, 5,,,,1	
		ggoSpread.SSSetEdit		3,	"조달구분"	, 8,,,,1	
		ggoSpread.SSSetEdit		4,	"품목"	, 8,,,,1
		ggoSpread.SSSetEdit		5,	"품목명"	, 10	
		ggoSpread.SSSetEdit		6,	"오더번호"	, 15,,,,1
		ggoSpread.SSSetEdit		7,	"공순"	, 5,,,,1
		ggoSpread.SSSetEdit		8,	"공정"	, 7,,,,1
		ggoSpread.SSSetEdit		9,	"공정명"	, 10			
		
		ggoSpread.SSSetFloat	iBas+1,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+2,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+3,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+4,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+5,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+6,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+7,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+8,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+9,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+10,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+11,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+12,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+13,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+14,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+15,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+16,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+17,	""	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+18,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+19,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+20,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		'Call ggoSpread.SSSetColHidden(36,37,True)

		For i = 1 To iBas
			ret = .AddCellSpan(i, -1000, 1, 8)
		Next
		
		For i = iBas +1 To iBas + 20 Step 4
			ret = .AddCellSpan(i, -1000, 4, 1)	' -- 기초재공/전공정대재(당기착수) 등등 

			Select Case i
				Case iBas + 1, iBas + 17
					ret = .AddCellSpan(i, -999, 1, 2)	
				Case Else
					ret = .AddCellSpan(i, -999, 1, 7)	
			End Select
			
			ret = .AddCellSpan(i+1, -999, 3, 1)
			ret = .AddCellSpan(i+1, -998, 3, 1)

			If i = iBas + 1 Or i = iBas + 17 Then			
				ret = .AddCellSpan(i, -997, 1, 5)
			End If
			
			ret = .AddCellSpan(i+1, -997, 3, 1)
			ret = .AddCellSpan(i+1, -996, 3, 1)
			ret = .AddCellSpan(i+1, -995, 3, 1)
			ret = .AddCellSpan(i+1, -994, 3, 1)
			' -- 마지막 행은 Span이 없다.
		Next
		
		' 1번째 헤더 출력 글자 
		.Row = -1000
		.Col = iBas+1	: .Text = "기초재공"
		.Col = iBas+5	: .Text = "전공정대체(당기착수)"
		.Col = iBas+9	: .Text = "차공정대체"
		.Col = iBas+13	: .Text = "완성대체"
		.Col = iBas+17	: .Text = "기말재공"
		
		' 2번째 헤더 출력 글자 
		.Row = -999
		.Col = iBas+1	: .Text = "기초재공수량"
		.Col = iBas+2	: .Text = "기초재공금액"
		.Col = iBas+5	: .Text = "전공정대체(당기착수수량)" 
		.Col = iBas+6	: .Text = "전공정대체(당기착수)금액"
		.Col = iBas+9	: .Text = "차공정대체수량"
		.Col = iBas+10	: .Text = "차공정대체(금액"
		.Col = iBas+13	: .Text = "완성대체수량"
		.Col = iBas+14	: .Text = "완성대체금액"
		.Col = iBas+17	: .Text = "기말재공수량"
		.Col = iBas+18	: .Text = "기말재공금액"


		' 3번째 헤더 출력 글자 
		.Row = -998
		For i = iBas +1 To iBas+20 Step 4
			.Col = i+1	: .Text = "재료비"
		Next

		' 4번째 헤더 출력 글자 
		.Row = -997
		.Col = iBas +1	: .Text = "기초재공환산수량"
		.Col = iBas +17	: .Text = "기말재공환산수량"

		For i = iBas +1 To iBas+20 Step 4
			.Col = i+1	: .Text = "반제품비"
		Next

		' 5번째 헤더 출력 글자 
		.Row = -996
		For i = iBas +1 To iBas+20 Step 4
			.Col = i+1	: .Text = "노무비"
		Next

		' 6번째 헤더 출력 글자 
		.Row = -995
		For i = iBas +1 To iBas+20 Step 4
			.Col = i+1	: .Text = "경비"
		Next

		' 7번째 헤더 출력 글자 
		.Row = -994
		For i = iBas +1 To iBas+20 Step 4
			.Col = i+1	: .Text = "외주가공비"
		Next

		' 8번째 헤더 출력 글자 
		.Row = -993
		For i = iBas +1 To iBas+20 Step 4
			.Col = i+1	: .Text = "합계"
			.Col = i+2	: .Text = "전공정원가"
			.Col = i+3	: .Text = "당공정(투입원가)"
		Next
		
		
		.rowheight(-993) = 20	' 높이 재지정 
		
		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		.Col = 5: .Row = -1: .ColMerge = 1
		.Col = 6: .Row = -1: .ColMerge = 1
		.Col = 7: .Row = -1: .ColMerge = 1
		.Col = 8: .Row = -1: .ColMerge = 1
		.Col = 9: .Row = -1: .ColMerge = 1

		ggoSpread.SSSetSplit2(iBas)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		
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
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			' -- 그리드1의 컬럼 정의 

		
    End Select    
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
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
			arrParam(0) = "조달구분 팝업"
			arrParam(1) = "dbo.B_MINOR"	
			arrParam(2) = Trim(.txtPROC_TYPE.value)
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1003", "''", "S")
			arrParam(5) = "조달구분" 

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"		
			arrField(2) = ""	
    
			arrHeader(0) = "조달구분"
			arrHeader(1) = "조달구분명"
			arrHeader(2) = "C/C LEVEL"	

		Case 3
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

		Case 4
			arrParam(0) = "공정/구매그룹 팝업"
			arrParam(1) = "dbo.ufn_C_getPopup_by_C4207MA1() "	
			arrParam(2) = Trim(.txtWC_CD.value)
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "공정/구매그룹" 

			arrField(0) = "HH10" & parent.gColSep & "WC_CD"	
			arrField(1) = "ED8" & parent.gColSep & "TYPE_FLG_NM"	
			arrField(2) = "ED14" & parent.gColSep & "WC_CD"
			arrField(3) = "ED20" & parent.gColSep & "WC_NM"		
    
			arrHeader(0) = ""	
			arrHeader(1) = "구분"	
			arrHeader(2) = "공정/구매그룹"
			arrHeader(3) = "공정/구매그룹명"
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
				.txtPROC_TYPE.value		= arrRet(0)
				.txtPROC_TYPE_NM.value	= arrRet(1)

			Case 3
				.txtITEM_CD.value		= arrRet(0)
				.txtITEM_NM.value		= arrRet(1)

			Case 4
				.txtWC_CD.value		= arrRet(0)
				.txtWC_NM.value		= arrRet(1)
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function
 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    Call CommonQueryRs(" OPTION_VALUE "," C_COST_CONFG_S ", "OPTION_CD=" & FilterVar("C4003", "''", "S") & " "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 <> "" Then
		lgOptionFlag = Replace(lgF0, Chr(11),"")
	End If
    'ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_GP_LEVEL
    
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
	'Call ggoOper.FormatDate(frm1.txtEND_DT, parent.gDateFormat,2)
    
    Call InitVariables
	Call InitComboBox		' -- lgOptionFlag 를 구해온다 
    
    If lgOptionFlag = "F" Then
		Call InitSpreadSheet
	Else
		Call InitSpreadSheet2
	End If
    
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
Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i

	Dim intCnt,IntRetCD

	MsgBox "미개발"
	Exit Function

    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If
    

    StrUrl = StrUrl & "VER_CD|"			& frm1.txtVER_CD.value 

     ObjName = AskEBDocumentName("C4002MA1", "ebr")
     
     if  strPrintType = "VIEW" then
		Call FncEBRPreview(ObjName, StrUrl)
     else
		Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	
     
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


Sub txtPlant_Cd_OnChange()
	Dim sWhereSQL
	
	If Trim(frm1.txtPlant_Cd.value) <> "" Then

		' -- 변경값 체크 
		sWhereSQL = " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.value, "''", "S")
		
		Call CommonQueryRs(" PLANT_CD, PLANT_NM "," B_PLANT ", sWhereSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			Call DisplayMsgBox("900014", "x", "x", "x")		
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
			Call DisplayMsgBox("900014", "x", "x", "x")		
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
		
		Call CommonQueryRs(" a.ITEM_CD, a.ITEM_NM "," B_ITEM a left outer join b_item_by_plant b on a.item_cd = b.item_cd  ", sWhereSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			Call DisplayMsgBox("900014", "x", "x", "x")		
			frm1.txtITEM_NM.value = ""
		Else
			frm1.txtITEM_CD.value = Replace(lgF0, Chr(11), "")
			frm1.txtITEM_NM.value = Replace(lgF1, Chr(11), "")
		End If 
	Else
		frm1.txtITEM_NM.value = ""
	End If	
End Sub

Sub txtPROC_TYPE_OnChange()
	Dim sWhereSQL
	
	If Trim(frm1.txtPROC_TYPE.value) <> "" Then

		' -- 변경값 체크 
		sWhereSQL = " MINOR_CD = " & FilterVar(frm1.txtPROC_TYPE.value, "''", "S") & " AND MAJOR_CD=" & FilterVar("P1003", "''", "S")
		
		Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR ", sWhereSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			Call DisplayMsgBox("900014", "x", "x", "x")		
			frm1.txtPROC_TYPE_NM.value = ""
		Else
			frm1.txtPROC_TYPE.value = Replace(lgF0, Chr(11), "")
			frm1.txtPROC_TYPE_NM.value = Replace(lgF1, Chr(11), "")
		End If 
	Else
		frm1.txtPROC_TYPE_NM.value = ""
	End If	
End Sub

Sub txtWC_CD_OnChange()
	Dim sWhereSQL
	
	If Trim(frm1.txtWC_CD.value) <> "" Then

		' -- 변경값 체크 
		sWhereSQL = " WC_CD = " & FilterVar(frm1.txtWC_CD.value, "''", "S")
		
		Call CommonQueryRs(" WC_CD, WC_NM "," dbo.ufn_C_getPopup_by_C4207MA1() ", sWhereSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		lgF0 = Replace(lgF0, Chr(11), "")
		lgF1 = Replace(lgF1, Chr(11), "")
	
		If lgF0 = "" Then
			Call DisplayMsgBox("900014", "x", "x", "x")		
			frm1.txtWC_NM.value = ""
		Else
			frm1.txtWC_CD.value = Replace(lgF0, Chr(11), "")
			frm1.txtWC_NM.value = Replace(lgF1, Chr(11), "")
		End If 
	Else
		frm1.txtWC_NM.value = ""
	End If	
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
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
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)

	
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	
End Sub

' -- 그리드1 팝업 버튼 클릭 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And (lgStrPrevKey <> "" AND lgStrPrevKey <> "*") Then
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
    
    If ChkKeyField=False then Exit Function 
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

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

	Call ExportExcel
	Exit Function
	
	If lgIntFlgMode = Parent.OPMD_UMODE And (lgStrPrevKey <> "*" And lgStrPrevKey <> "") Then
		' -- 전체 올 쿼리한다.
		Call ggoOper.ClearField(Document, "2")
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
	
		lgStrPrevKey = "*"
		Call DBQuery
	Else
		Call parent.FncExport(Parent.C_MULTI)
	End If
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
		Call parent.ExtractDateFromSuper(.txtSTART_DT.Text, parent.gDateFormat,sYear,sMon,sDay)
		
		sStartDt = sYear & parent.gComDateType & sMon & parent.gComDateType & sDay
		
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtStartDt=" & sStartDt
		strVal = strVal & "&txtEndDt=" & sEndDt	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPLANT_CD=" & Trim(.txtPLANT_CD.value)
		strVal = strVal & "&txtITEM_ACCT=" & Trim(.txtITEM_ACCT.value)
		strVal = strVal & "&txtPROC_TYPE=" & Trim(.txtPROC_TYPE.value)
		strVal = strVal & "&txtITEM_CD=" & Trim(.txtITEM_CD.value)
		strVal = strVal & "&txtWC_CD=" & Trim(.txtWC_CD.value)
		strVal = strVal & "&txtOptionFlag=" & lgOptionFlag				
		'strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		
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
	
	frm1.vspdData.style.display = ""	'-- 그리드 보이게..
	
	Frm1.vspdData.Focus
   	
    Set gActiveElement = document.ActiveElement   
	
	lgIntFlgMode = Parent.OPMD_UMODE
	'Call SetQuerySpreadColor
	
	window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"
	
	If 	lgStrPrevKey = "*" Then
		parent.FncExport(Parent.C_MULTI)
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

	With frm1.vspdData	
	.ReDraw = False
	arrRow = Split(pGrpRow, Parent.gRowSep)
	
	iLoopCnt = UBound(arrRow, 1)
	
	For i = 0 to iLoopCnt -1
		arrCol = Split(arrRow(i), Parent.gColSep)
	
		'.BlockMode = True
		.Col = -1 
		.Row = CDbl(arrCol(1)) * 6 - 5	' -- 행 
		'.Row2 = CDbl(arrCol(1)) * 6 	' -- 행 
		
		Select Case arrCol(0)
			Case "2"
				.Col = -1
			   ret = .AddCellSpan(7, .Row , 3, 6)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
			Case "3"
				.Col = -1
				ret = .AddCellSpan(6, .Row , 4, 6)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
			Case "4"
				.Col = -1
				ret = .AddCellSpan(4, .Row , 6, 6)
				.BackColor = RGB(204,255,255) 
				.ForeColor = vbBlack
			Case "5"  
				.Col = -1
				ret = .AddCellSpan(3, .Row, 7, 6)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
			Case "6" 
				ret = .AddCellSpan(2, .Row, 8, 6)
				.BackColor = RGB(255,240,245) 
				.ForeColor = vbBlack
			Case "7" 
				ret = .AddCellSpan(1, .Row, 9, 6)
				.BackColor = RGB(255,250,245) 
				.ForeColor = vbBlack
		End Select
		'.BlockMode = False
	Next

	.ReDraw = True
	End With

End Sub

' -- 집계 조회시 행번호와 소계라인이 같으므로 실재행을 찾는다.
Function FindRow(Byval pRow, Byval pGrpNo)
	Dim i, iMaxRows
	With frm1.vspdData
		iMaxRows = .MaxRows
		For i = pRow To iMaxRows
			.Row = i 
			.Col = .MaxCols -1 
			If .Text = pGrpNo Then
				FindRow = .Row
				Exit Function
			End If
		Next
	End With
End Function
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

Function ExportExcel()
	Dim iExcelRow
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
	'xlApp.Visible = True
	
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
	xlSheet.Cells(iExcelRow, 5 ).value = "조달구분"
	xlSheet.Cells(iExcelRow, 6 ).value = frm1.txtPROC_TYPE.value
	xlSheet.Cells(iExcelRow, 7 ).value = frm1.txtPROC_TYPE_NM.value

	iExcelRow = iExcelRow + 1
	xlSheet.Cells(iExcelRow, 1 ).value = "품목"
	xlSheet.Cells(iExcelRow, 2 ).value = frm1.txtITEM_CD.value
	xlSheet.Cells(iExcelRow, 3 ).value = frm1.txtITEM_NM.value
	xlSheet.Cells(iExcelRow, 5 ).value = "공정/구매그룹"
	xlSheet.Cells(iExcelRow, 6 ).value = frm1.txtWC_CD.value
	xlSheet.Cells(iExcelRow, 7 ).value = frm1.txtWC_NM.value
	
	' -- 그리드 헤더 찍기 
	iExcelRow = iExcelRow + 3
	
	iColHeadCnt = .ColHeaderRows
	
	For iRow = 0 To iColHeadCnt - 1
		.Row = -1000 + iRow
		
		For iCol = 1 To .MaxCols 
			.Col = iCol
			xlSheet.Cells(iExcelRow + iRow, iCol ).value = .Text
		Next
	Next
	
	iExcelRow = iExcelRow + iRow
	' -- 데이타 찍기 
	For iRow = 1 To .MaxRows
		.Row = iRow
		
		For iCol = 1 To .MaxCols 
			.Col = iCol
			xlSheet.Cells(iExcelRow + iRow, iCol ).value = .Text
		Next
	Next
	
	xlApp.Visible = True
	
	End With
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
	'check proc type
	If Trim(frm1.txtPROC_TYPE.value) <> "" Then
		strFrom = " b_minor  "
		strWhere = " minor_cd  = " & FilterVar(frm1.txtPROC_TYPE.value, "''", "S") & " "
		strWhere = strWhere & "		and MAJOR_CD =" & FilterVar("P1003", "''", "S")	
		
		Call CommonQueryRs(" minor_nm  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPROC_TYPE.alt,"X")
			frm1.txtPROC_TYPE.focus 
			frm1.txtPROC_TYPE_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPROC_TYPE_NM.value = strDataNm(0)
	ELSE
		frm1.txtPROC_TYPE_NM.value=""
	End If
'check item
	If Trim(frm1.txtITEM_CD.value) <> "" Then
		strFrom = " B_ITEM a left outer join dbo.b_item_by_plant b on a.item_cd = b.item_cd  "
		strWhere = " a.item_cd  = " & FilterVar(frm1.txtITEM_CD.value, "''", "S") & " "
		If frm1.txtPLANT_CD.value <> "" then
			strWhere = strWhere & " and b.PLANT_CD = " & FilterVar(frm1.txtPLANT_CD.value, "''", "S")
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


	'check WC
	If Trim(frm1.txtWC_CD.value) <> "" Then
		strFrom = "  dbo.ufn_C_getPopup_by_C4207MA1()  "	

		strWhere = " wc_cd  = " & FilterVar(frm1.txtWC_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs(" wc_nm  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtWC_CD.alt,"X")
			frm1.txtWC_CD.focus 
			frm1.txtWC_nm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtWC_nm.value = strDataNm(0)
	ELSE
		frm1.txtWC_nm.value=""
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
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSTART_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 기준년월" tag="12" id=txtSTART_DT></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtPLANT_CD" TYPE="Text" MAXLENGTH="4" tag="15XXXU" size="20" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPLANT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_ACCT" TYPE="Text" MAXLENGTH="2" tag="15XXXU" size="10" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)">
									<input NAME="txtITEM_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5" NOWRAP>조달구분</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtPROC_TYPE" TYPE="Text" MAXLENGTH="1" tag="15XXXU" size="10" ALT="조달구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(2)">
									<input NAME="txtPROC_TYPE_NM" TYPE="TEXT"  tag="14XXX" size="20"></TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_CD" TYPE="Text" MAXLENGTH="18" tag="15XXXU" size="20" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(3)">
									<input NAME="txtITEM_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5" NOWRAP>공정/구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtWC_CD" TYPE="Text" MAXLENGTH="7" tag="15XXXU" size="10" ALT="공정/구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(4)">
									<input NAME="txtWC_NM" TYPE="TEXT"  tag="14XXX" size="20"></TD>
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpreadI1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadI2" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadU1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadU2" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadD1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadD2" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hDstbFctr" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

