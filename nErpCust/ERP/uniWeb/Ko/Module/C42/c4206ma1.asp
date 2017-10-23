<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 공정별원가조회 
'*  3. Program ID           : c4206ma1.asp
'*  4. Program Name         : 공정별원가조회 
'*  5. Program Desc         : 공정별원가조회 
'*  6. Modified date(First) : 2005-10-20
'*  7. Modified date(Last)  :
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4206mb1.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt
Dim lgStrPrevKey2
Dim lgRow, lgEOF, lgEOF2

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
Dim lgOptionFlag

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
    lgEOF = False

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtYYYYMM.Text =UniConvDateAToB(iStrFromDt, parent.gServerDateFormat, parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, parent.gDateFormat, 2)

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

'------------------------------------------  C_COST_CONFG_S()  ----------------------------------------------
'	Name :getOptionValue
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub getOptionValue()
	DIm strWhere, tmpData
	
	strWhere = " OPTION_CD=" & FilterVar("C4003", "''", "S") 
    Call CommonQueryRs(" top 1 OPTION_VALUE "," C_COST_CONFG_S ",strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If lgF0 <> "" Then	
		tmpData=split(lgF0,chr(11))
		lgOptionFlag =tmpData(0)		
	Else	
		lgOptionFlag="F"
	End If
    
    
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
		ggoSpread.Spreadinit "V20030825", ,Parent.gAllowDragDropSpread
			
		iBas = 4	' -- 앞에 고정이 변할경우 대비 

		.MaxRows = 0
		.MaxCols = iBas+30		' -- Group/RowNum 

		.Col  = iBas+29 : .ColHidden = True
		.Col  = iBas+30 : .ColHidden = True
		
		'헤더를 2줄로    
		.ColHeaderRows = 8

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetEdit		1,	"공장"	, 9,,,20,1	
		ggoSpread.SSSetEdit		2,	"사내/외주구분"	, 10,,,,1
		ggoSpread.SSSetEdit		3,	"공정"	, 7,,,,1
		ggoSpread.SSSetEdit		4,	"공정명"	, 10			
		
		ggoSpread.SSSetFloat	iBas+1,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+2,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+3,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+4,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+5,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+6,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+7,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+8,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+9,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+10,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+11,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+12,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+13,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+14,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+15,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+16,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+17,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+18,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+19,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+20,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+21,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+22,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+23,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+24,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+25,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+26,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+27,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	iBas+28,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		
		For i = 1 To iBas
			ret = .AddCellSpan(i, -1000, 1, 8)
		Next
		
		For i = iBas +1 To iBas + 28 Step 4
			Select Case i
				Case iBas + 9, iBas + 17
					ret = .AddCellSpan(i, -1000, 8, 1)	' -- 기초재공/전공정대재(당기착수) 등등 
					'ret = .AddCellSpan(i, -1000, 4, 1)	' -- 기초재공/전공정대재(당기착수) 등등 
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

		ggoSpread.SSSetSplit2(iBas)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		
	End With
End Sub
		
Sub InitSpreadSheet2()
	Dim i, ret, iBas
	With frm1.vspdData
		
		.style.display = ""

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030825", ,Parent.gAllowDragDropSpread
			
		iBas = 4	' -- 앞에 고정이 변할경우 대비 

		.MaxRows = 0
		.MaxCols = iBas+22		' -- Group/RowNum 

		.Col  = iBas+21 : .ColHidden = True
		.Col  = iBas+22 : .ColHidden = True
		
		'헤더를 2줄로    
		.ColHeaderRows = 8

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetEdit		1,	"공장"	, 6,,,,1	
		ggoSpread.SSSetEdit		2,	"사내/외주구분"	, 5,,,,1	
		ggoSpread.SSSetEdit		3,	"공정"	, 7,,,,1
		ggoSpread.SSSetEdit		4,	"공정명"	, 10				
		
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
		.Col = iBas+10	: .Text = "차공정대체(금액)"
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

		ggoSpread.SSSetSplit2(iBas)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		
	End With
End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 ' -- 그리드1에서 팝업 클릭시 
Function OpenPopUp(Byval iWhere)
	Dim arrRet, sTmp, strFrom
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	Select Case iWhere
		Case 0
			arrParam(0) = "공장 팝업"
			arrParam(1) = "dbo.B_PLANT"	
			arrParam(2) = Trim(.txtPlantCD.value)
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
			strFrom = "(select  '공정' AS FLAG_NM , wc_cd as code, wc_nm as cd_nm	 from P_work_center "
			strFrom = strFrom & " union "
			strFrom = strFrom & "	select    '구매그룹' AS FLAG_NM, pur_grp as code, pur_grp_nm as cd_nm from b_pur_grp where usage_flg='Y') tmp"
					
			arrParam(0) = "공정/구매그룹팝업"						' 팝업 명칭 
			arrParam(1) =strFrom					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtWPCd.Value)	' Code Condition
			arrParam(3) =""'Trim(frm1.txtWPCd.Value)										' Name Cindition
			arrParam(4) =""							' Where Condition
			arrParam(5) = "공정/구매그룹"							' TextBox 명칭			
			
			arrField(0) = "HH" & Parent.gColSep & "code"					' Field명(1)
			arrField(1) ="ED10" & Parent.gColSep &  "FLAG_NM"					' Field명(0)
			arrField(2) = "ED10" & Parent.gColSep & "code"					' Field명(1)
			arrField(3) ="ED25" & Parent.gColSep &  "cd_nm"					' Field명(0)
    		
    		arrHeader(0) = "공정/구매그룹"						' Header명(0)
			arrHeader(1) = "공정/구매그룹구분"						' Header명(0)    
			arrHeader(2) = "공정/구매그룹"						' Header명(0)
			arrHeader(3) = "공정/구매그룹명"						' Header명(1)    

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
				.txtPlantCD.value		= arrRet(0)
				.txtPlantNm.value		= arrRet(1)				
			Case 1
				.txtWPCd.Value    = arrRet(2)		
				.txtWPNm.Value   = arrRet(3)							
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
	Call ggoOper.FormatDate(frm1.txtYYYYMM, parent.gDateFormat,2)

    Call InitVariables
    

    Call SetDefaultVal
    Call SetToolbar("110000000001111")	
    If parent.gPlant <> "" Then
		frm1.txtPlantCD.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtWpCd.focus 		
	Else
		frm1.txtPlantCD.focus 		
	End If
	
	Call getOptionValue
	
	
	If  (lgOptionFlag) = "F" Then	
		Call InitSpreadSheet
	Else	
		Call InitSpreadSheet2
	End If
	
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
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub txtYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtYYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYYYYMM.Focus
    End If
End Sub



Sub txtWpCd_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtPlantCD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" And lgEOF = False Then
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
    
    sStartDt= Replace(frm1.txtYYYYMM.text, parent.gComDateType, "")
    
    If ChkKeyField=false then exit Function 
  
    Call ggoOper.ClearField(Document, "2")

	frm1.vspdData.MaxRows = 0
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
    Err.Clear                                                                   '☜: Clear error status
     
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
'========================================================================================
' Function Name : FncExcel
' Function Desc : 
'========================================================================================

Function FncExcel() 
	lgStrPrevKey = "*"
	Call DBQuery
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
End Function
'========================================================================================
' Function Name : FncExcelUpLoad
' Function Desc : 
'========================================================================================
Function FncExcelUpLoad()
    Call parent.FncExport(Parent.C_MULTI)
End function
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
If lgOptionFlag = "F" Then
		Call InitSpreadSheet
	Else
		Call InitSpreadSheet2
	End If    
    Call getOptionValue
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
		Call parent.ExtractDateFromSuper(.txtYYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)
		
		sStartDt =trim(sYear & sMon)
		
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtYYYYMM=" & sStartDt		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtWpCd=" & Trim(.txtWpCd.value)
		strVal = strVal & "&txtPlantCD=" & Trim(.txtPlantCD.value)
		strVal = strVal & "&txtOptionValue=" & lgOptionFlag


'		lgSTime = Time	' -- 디버깅용 
		
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
	'window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"
	
	If 	lgStrPrevKey = "*" Then
		Call FncExcelUpLoad() 
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
	
		.Col = -1 
		.Row = CDbl(arrCol(1)) * 6 - 5	' -- 행 
		
		Select Case arrCol(0)

			Case "1"
				.Col = -1
				ret = .AddCellSpan(3, .Row , 2, 6)
				.BackColor = RGB(204,255,255) 
				.ForeColor = vbBlack
			Case "2"  
				.Col = -1
				ret = .AddCellSpan(2, .Row, 3, 6)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
			Case "3" 
				ret = .AddCellSpan(1, .Row,4, 6)
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
'check plant
	If Trim(frm1.txtPlantCD.value) <> "" Then
		strWhere = " plant_cd= " & FilterVar(frm1.txtPlantCD.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	 b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlantCD.alt,"X")
			frm1.txtPlantCD.focus 
			frm1.txtPlantNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlantNm.value = strDataNm(0)
	Else
		frm1.txtPlantNm.value=""
	End If
'check wc
	If Trim(frm1.txtWPCd.value) <> "" Then
		strFrom = "(select  '공정' AS FLAG_NM , wc_cd as code, wc_nm as cd_nm	 from P_work_center "
		strFrom = strFrom & " union "
		strFrom = strFrom & "	select    '구매그룹' AS FLAG_NM, pur_grp as code, pur_grp_nm as cd_nm from b_pur_grp where usage_flg='Y') tmp"

		strWhere = " code  = " & FilterVar(frm1.txtWPCd.value, "''", "S") & " "			
		
		Call CommonQueryRs(" cd_nm  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtWPCd.alt,"X")
			frm1.txtWPCd.focus 
			frm1.txtWPNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtWPNm.value = strDataNm(0)
	ELSE
		frm1.txtWPNm.value=""
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
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 기준년월" tag="12" id=txtYYYYMM></OBJECT>');</SCRIPT>&nbsp;
									
									</TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtPlantCD" TYPE="Text" MAXLENGTH="4" tag="15XXXU" size="10" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPlantNm" TYPE="TEXT"  tag="14XXX" size="25">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">공정/구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWPCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="공정/구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(1)">
									<INPUT TYPE=TEXT NAME="txtWPNm" SIZE=25 tag="14">
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6" valign=top>&nbsp;</TD>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hOptionFlag" tag="24" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

