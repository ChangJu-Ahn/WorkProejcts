<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 공제감면세액조정 
'*  3. Program ID           : W6101MA1
'*  4. Program Name         : W6101MA1.asp
'*  5. Program Desc         : 제48호 소득구분계산서 
'*  6. Modified date(First) : 2005/01/24
'*  7. Modified date(Last)  : 2005/01/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W6101MA1"
Const BIZ_PGM_ID		= "W6101MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W6101RA1.asp"											 '☆: 비지니스 로직 ASP명 
Const JUMP_PGM_ID		= "W8101MA1"
Const EBR_RPT_ID	    = "W6101OA1"

Const TAB1 = 1																	'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3
Const TAB4 = 4

Const TYPE_1	= 0		' 그리드를 구분짓기 위한 상수 
Const TYPE_2_1	= 1		
Const TYPE_2_2	= 2		 
Const TYPE_3	= 3		
Const TYPE_4	= 4		

' -- 그리드 컬럼 정의 
Dim C_1_W1
Dim C_1_W2
Dim C_1_W1_CD
Dim C_1_W3
Dim C_1_W4_1
Dim C_1_W4_2
Dim C_1_W4_3
Dim C_1_W4_4
Dim C_1_W4_5
Dim C_1_W4_6
Dim C_1_W5

Dim C_2_SEQ_NO
Dim C_2_W1
Dim C_2_W1_BT
Dim C_2_W1_NM
Dim C_2_W2
Dim C_2_W3_CD
Dim C_2_W3
Dim C_2_W4

Dim C_3_W_TYPE
Dim C_3_W1
Dim C_3_W2
Dim C_3_W3_1
Dim C_3_W4_1
Dim C_3_W5_1
Dim C_3_W6_1
Dim C_3_W7_1
Dim C_3_W8_1
Dim C_3_W3_2
Dim C_3_W4_2
Dim C_3_W5_2
Dim C_3_W6_2
Dim C_3_W7_2
Dim C_3_W8_2
Dim C_3_W9
Dim C_3_W10
Dim C_3_W11

Dim C_4_W1
Dim C_4_W2
Dim C_4_W1_CD
Dim C_4_W3
Dim C_4_W4_1
Dim C_4_W5_1
Dim C_4_W4_2
Dim C_4_W5_2
Dim C_4_W4_3
Dim C_4_W5_3
Dim C_4_W4_4
Dim C_4_W5_4
Dim C_4_W4_5
Dim C_4_W5_5
Dim C_4_W4_6
Dim C_4_W5_6
Dim C_4_W6
Dim C_4_W7
Dim C_4_DESC1

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(4)

Dim lgW_NM(8)

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_1_W1		= 1
	C_1_W2		= 2
	C_1_W1_CD	= 3
	C_1_W3		= 4
	C_1_W4_1	= 5
	C_1_W4_2	= 6
	C_1_W4_3	= 7
	C_1_W4_4	= 8
	C_1_W4_5	= 9
	C_1_W4_6	= 10
	C_1_W5		= 11

	C_2_SEQ_NO	= 1
	C_2_W1		= 2
	C_2_W1_BT	= 3
	C_2_W1_NM	= 4
	C_2_W2		= 5
	C_2_W3_CD	= 6
	C_2_W3		= 7
	C_2_W4		= 8
	
	C_3_W_TYPE	= 1
	C_3_W1		= 2
	C_3_W2		= 3
	C_3_W3_1	= 4
	C_3_W4_1	= 5
	C_3_W5_1	= 6
	C_3_W6_1	= 7
	C_3_W7_1	= 8
	C_3_W8_1	= 9
	C_3_W3_2	= 10
	C_3_W4_2	= 11
	C_3_W5_2	= 12
	C_3_W6_2	= 13
	C_3_W7_2	= 14
	C_3_W8_2	= 15
	C_3_W9		= 16
	C_3_W10		= 17
	C_3_W11		= 18

	C_4_W1		= 1
	C_4_W2		= 2
	C_4_W1_CD	= 3
	C_4_W3		= 4
	C_4_W4_1	= 5
	C_4_W5_1	= 6
	C_4_W4_2	= 7
	C_4_W5_2	= 8
	C_4_W4_3	= 9
	C_4_W5_3	= 10
	C_4_W4_4	= 11
	C_4_W5_4	= 12
	C_4_W4_5	= 13
	C_4_W5_5	= 14
	C_4_W4_6	= 15
	C_4_W5_6	= 16
	C_4_W6		= 17
	C_4_W7		= 18
	C_4_DESC1	= 19
	
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False

    lgCurrGrid = TYPE_1
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  콤보 박스 채우기  ====================================

Sub InitComboBox()
	' 조회조건(구분)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1047' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW11 ,lgF0  ,lgF1  ,Chr(11))

End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1063' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(TYPE_2_1)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_2_W3_CD
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_2_W3
		
		ggoSpread.Source = lgvspdData(TYPE_2_2)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_2_W3_CD
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_2_W3
	End If
End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1)		= frm1.vspdData0
	Set lgvspdData(TYPE_2_1)	= frm1.vspdData1
	Set lgvspdData(TYPE_2_2)	= frm1.vspdData2
	Set lgvspdData(TYPE_3)		= frm1.vspdData3
	Set lgvspdData(TYPE_4)		= frm1.vspdData4
		
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
	Call AppendNumberPlace("8","15","0")
	' 1번 그리드 

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_1_W5 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
 
  		'헤더를 3줄로    
		.ColHeaderRows = 3  
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_1_W1,		"(1)과목", 20,,,50,1
		ggoSpread.SSSetEdit		C_1_W2,		"(2)구분", 6, 2,,50,1
		ggoSpread.SSSetEdit		C_1_W1_CD,	"코드"	, 5,2,,50,1
		ggoSpread.SSSetFloat	C_1_W3,		"(3)합계"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_1_W4_1,	"(4-1)금액"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_1_W4_2,	"(4-2)금액"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_1_W4_3,	"(4-3)금액"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_1_W4_4,	"(4-4)금액"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_1_W4_5,	"(4-5)금액"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_1_W4_6,	"(4-6)금액"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_1_W5,		"(5)금액"	, 13, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

		' 그리드 헤더 합침 
		ret = .AddCellSpan(C_1_W1		, -1000, 1, 3)	' 순번 2행 합침 
		ret = .AddCellSpan(C_1_W2		, -1000, 1, 3)	
		ret = .AddCellSpan(C_1_W1_CD	, -1000, 1, 3)
		ret = .AddCellSpan(C_1_W3 		, -1000, 1, 3)
		ret = .AddCellSpan(C_1_W4_1		, -1000, 6, 1)
		ret = .AddCellSpan(C_1_W5		, -1000, 1, 3)
    
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_1_W4_1	: .Text = "감면분 또는 합병 승계사업해당분등"
		
		.Row = -998
		.Col = C_1_W4_1	: .Text = "(4-1)금액"
		.Col = C_1_W4_2	: .Text = "(4-2)금액"
		.Col = C_1_W4_3	: .Text = "(4-3)금액"
		.Col = C_1_W4_4	: .Text = "(4-4)금액"
		.Col = C_1_W4_5	: .Text = "(4-5)금액"
		.Col = C_1_W4_6	: .Text = "(4-6)금액"
		
		.rowheight(-999) = 12					
		'.rowheight(-998) = 15	' 높이 재지정	(2줄일 경우, 1줄은 15)
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_1_W4_4,C_1_W4_4,True)
		Call ggoSpread.SSSetColHidden(C_1_W4_5,C_1_W4_5,True)
		Call ggoSpread.SSSetColHidden(C_1_W4_6,C_1_W4_6,True)
		
		'Call InitSpreadComboBox
		
		.ReDraw = true	
				
	End With 

 
	' 2-1번 그리드 
	With lgvspdData(TYPE_2_1)
				
		ggoSpread.Source = lgvspdData(TYPE_2_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2_1,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_2_W4 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
	
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_2_SEQ_NO,	"순번", 10,,,15,1
		ggoSpread.SSSetEdit		C_2_W1	,	"코드", 7,,,10,1
		ggoSpread.SSSetButton 	C_2_W1_BT
		ggoSpread.SSSetEdit		C_2_W1_NM,	"(1)과목", 13,,,100,1
		ggoSpread.SSSetFloat	C_2_W2,		"(2)금액", 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetCombo	C_2_W3_CD,	"(3)코드"    , 10, 0
		ggoSpread.SSSetCombo	C_2_W3,		"(3)구분1"    , 10, 0
		ggoSpread.SSSetCombo	C_2_W4,		"(4)구분2"    , 10, 0
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_2_SEQ_NO,C_2_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_2_W3_CD,C_2_W3_CD,True)
				
		.ReDraw = true	
				
	End With 

	' 2-2번 그리드 

	With lgvspdData(TYPE_2_2)
				
		ggoSpread.Source = lgvspdData(TYPE_2_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2_2,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_2_W4 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
	
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_2_SEQ_NO,	"순번", 10,,,15,1
		ggoSpread.SSSetEdit		C_2_W1	,	"코드", 7,,,10,1
		ggoSpread.SSSetButton 	C_2_W1_BT
		ggoSpread.SSSetEdit		C_2_W1_NM,	"(1)과목", 13,,,100,1
		ggoSpread.SSSetFloat	C_2_W2,		"(2)금액", 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetCombo	C_2_W3_CD,	"(3)코드"    , 10, 0
		ggoSpread.SSSetCombo	C_2_W3,		"(3)구분1"    , 10, 0
		ggoSpread.SSSetCombo	C_2_W4,		"(4)구분2"    , 10, 0
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_2_SEQ_NO,C_2_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_2_W3_CD,C_2_W3_CD,True)

		.ReDraw = true	
				
	End With
	' 그리드 콤보 
	Call InitSpreadComboBox
	
	' 3번 그리드 

	With lgvspdData(TYPE_3)
				
		ggoSpread.Source = lgvspdData(TYPE_3)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_3,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_3_W11 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
	
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_3_W_TYPE,	"코드", 10,,,20,1
		ggoSpread.SSSetFloat	C_3_W1,		"(1)합계", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W2,		"(2)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W3_1,	"(3)감면1", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W4_1,	"(4)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W5_1,	"(5)감면2", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W6_1,	"(6)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W7_1,	"(7)감면3", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W8_1,	"(8)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W3_2,	"(3)감면4", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W4_2,	"(4)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W5_2,	"(5)감면5", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W6_2,	"(6)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W7_2,	"(7)감면6", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W8_2,	"(8)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W9,		"(9)기타", 15, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_3_W10,	"(10)비율", 7, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetEdit		C_3_W11,	"코드", 10,,,10,1
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_3_W3_2,C_3_W8_2,True)
		Call ggoSpread.SSSetColHidden(C_3_W11,C_3_W11,True)	
					
		.ReDraw = true	
				
	End With
		
	' 4번 그리드 

	With lgvspdData(TYPE_4)
				
		ggoSpread.Source = lgvspdData(TYPE_4)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_4,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_4_DESC1 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
 
  		'헤더를 3줄로    
		.ColHeaderRows = 3  
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_4_W1,		"(1)과목", 20,,,50,1
		ggoSpread.SSSetEdit		C_4_W2,		"(2)구분", 6,2,,50,1
		ggoSpread.SSSetEdit		C_4_W1_CD,	"코드"	, 5,2,,50,1
		ggoSpread.SSSetFloat	C_4_W3,		"(3)합계"	, 11, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_4_W4_1,	"(4)금액"	, 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W5_1,	"(5)비율"	, 6, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W4_2,	"(4)금액"	, 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W5_2,	"(5)비율"	, 6, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W4_3,	"(4)금액"	, 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W5_3,	"(5)비율"	, 6, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W4_4,	"(4)금액"	, 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W5_4,	"(5)비율"	, 6, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W4_5,	"(4)금액"	, 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W5_5,	"(5)비율"	, 6, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W4_6,	"(4)금액"	, 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W5_6,	"(5)비율"	, 6, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W6,		"(6)금액"	, 10, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_4_W7,		"(7)비율"	, 6, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetEdit		C_4_DESC1,	"비고"		, 7,,,50,1

		' 그리드 헤더 합침 
		ret = .AddCellSpan(C_4_W1		, -1000, 1, 3)	' 순번 2행 합침 
		ret = .AddCellSpan(C_4_W2		, -1000, 1, 3)	
		ret = .AddCellSpan(C_4_W1_CD	, -1000, 1, 3)
		ret = .AddCellSpan(C_4_W3 	, -1000, 1, 3)
		ret = .AddCellSpan(C_4_W4_1	, -1000,12, 1)
		ret = .AddCellSpan(C_4_W4_1	, -999 , 2, 1)
		ret = .AddCellSpan(C_4_W4_2	, -999 , 2, 1)
		ret = .AddCellSpan(C_4_W4_3	, -999 , 2, 1)
		ret = .AddCellSpan(C_4_W4_4	, -999 , 2, 1)
		ret = .AddCellSpan(C_4_W4_5	, -999 , 2, 1)
		ret = .AddCellSpan(C_4_W4_6	, -999 , 2, 1)
		ret = .AddCellSpan(C_4_W6		, -1000, 2, 1)
		ret = .AddCellSpan(C_4_W6 	, -999 , 1, 2)
		ret = .AddCellSpan(C_4_W7 	, -999 , 1, 2)
		ret = .AddCellSpan(C_4_DESC1	, -1000, 1, 3)
    
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_4_W4_1	: .Text = "감면분 또는 합병 승계사업해당분등"
		.Col = C_4_W6	: .Text = "기타분"
		
		.Row = -998
		.Col = C_4_W4_1	: .Text = "(4)금액"
		.Col = C_4_W5_1	: .Text = "(5)비율"
		.Col = C_4_W4_2	: .Text = "(4)금액"
		.Col = C_4_W5_2	: .Text = "(5)비율"
		.Col = C_4_W4_3	: .Text = "(4)금액"
		.Col = C_4_W5_3	: .Text = "(5)비율"
		.Col = C_4_W4_4	: .Text = "(4)금액"
		.Col = C_4_W5_4	: .Text = "(5)비율"
		.Col = C_4_W4_5	: .Text = "(4)금액"
		.Col = C_4_W5_5	: .Text = "(5)비율"
		.Col = C_4_W4_6	: .Text = "(4)금액"
		.Col = C_4_W5_6	: .Text = "(5)비율"

		.Row = -999
		.Col = C_4_W6		: .Text = "(6)금액"
		.Col = C_4_W7		: .Text = "(7)비율"
								
		'.rowheight(-1000) = 30	' 높이 재지정	(2줄일 경우, 1줄은 15)
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_4_W4_4,C_4_W5_6,True)

		
		'Call InitSpreadComboBox
		
		.ReDraw = true	
				
	End With 
					
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	Call GetFISC_DATE
	
	'Exit Sub
	' 기본 그리드 생성 
	Call MakeDefaultGrid("N")
		
End Sub

' 그리드 재구성: Query후, New/Form_load후 
Sub MakeDefaultGrid(Byval pMode)
	Dim ret, iRow, iMaxRows, arrF0, arrF1, iCol

	' 탭1번 그리드 
	ret = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1062' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	With lgvspdData(TYPE_1)
	
	If ret <> False Then
		arrF0 = Split(lgF0, chr(11))
		arrF1 = Split(lgF1, chr(11))
		iMaxRows = UBound(arrF0)
		
		.Redraw = False
		ggoSpread.Source = lgvspdData(TYPE_1)
		ggoSpread.InsertRow , iMaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			If pMode <> "N" Then 
				.Col = 0		: .value = iRow
			End If
			.Col = C_1_W1		: .value = arrF1(iRow-1)
			.Col = C_1_W1_CD	: .value = arrF0(iRow-1)
			
			If iRow = 4 Or iRow = 8 Or iRow = 11 Or iRow = 15 Or iRow = 18 Then
				.Col = C_1_W2	: .value = "개별분"
			ElseIf iRow = 5 Or iRow = 9 Or iRow = 12 Or iRow = 16 Or iRow = 19 Then
				.Col = C_1_W2	: .value = "공통분"
			ElseIf iRow = 6 Or iRow = 10 Or iRow = 13 Or iRow = 17 Or iRow = 20 Then
				.Col = C_1_W2	: .value = "계"
			End If
		Next

		.Col = C_1_W1
		ret = .AddCellSpan(C_1_W4_1 , 3 , 7, 1)
		ret = .AddCellSpan(C_1_W4_1 , 5 , 7, 3)
		ret = .AddCellSpan(C_1_W4_1 , 9 , 7, 2)
		ret = .AddCellSpan(C_1_W4_1 ,12 , 7, 3)
		ret = .AddCellSpan(C_1_W4_1 ,16 , 7, 2)
		ret = .AddCellSpan(C_1_W4_1 ,19 , 7, 6)
		ret = .AddCellSpan(C_1_W1 	, 4 , 1, 3)	: .Row = 4 : .TypeVAlign = 2
		ret = .AddCellSpan(C_1_W1 	, 8 , 1, 3)	: .Row = 8 : .TypeVAlign = 2
		ret = .AddCellSpan(C_1_W1 	, 11 , 1, 3): .Row = 11 : .TypeVAlign = 2
		ret = .AddCellSpan(C_1_W1 	, 15 , 1, 3): .Row = 15 : .TypeVAlign = 2
		ret = .AddCellSpan(C_1_W1 	, 18 , 1, 3): .Row = 18 : .TypeVAlign = 2
		
		Call SetSpreadLock(TYPE_1)
		
		.Redraw = True
		
		.SetActiveCell	C_1_W3, 1
	End If
	
	End With

	' 탭4 그리드 
	With lgvspdData(TYPE_4)
	
	If ret <> False Then
		arrF0 = Split(lgF0, chr(11))
		arrF1 = Split(lgF1, chr(11))
		iMaxRows = UBound(arrF0)
		
		.Redraw = False
		ggoSpread.Source = lgvspdData(TYPE_4)
		ggoSpread.InsertRow , iMaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			If pMode <> "N" Then 
				.Col = 0		: .value = iRow
			End If
			.Col = C_4_W1		: .value = arrF1(iRow-1)
			.Col = C_4_W1_CD	: .value = arrF0(iRow-1)
			
			If iRow = 4 Or iRow = 8 Or iRow = 11 Or iRow = 15 Or iRow = 18 Then
				.Col = C_4_W2	: .value = "개별분"
			ElseIf iRow = 5 Or iRow = 9 Or iRow = 12 Or iRow = 16 Or iRow = 19 Then
				.Col = C_4_W2	: .value = "공통분"
			ElseIf iRow = 6 Or iRow = 10 Or iRow = 13 Or iRow = 17 Or iRow = 20 Then
				.Col = C_4_W2	: .value = "계"
			End If
		Next

		' -- 중복 과목 SPAN
		ret = .AddCellSpan(C_4_W1 	, 4 , 1, 3)	: .Row = 4 : .TypeVAlign = 2
		ret = .AddCellSpan(C_4_W1 	, 8 , 1, 3)	: .Row = 8 : .TypeVAlign = 2
		ret = .AddCellSpan(C_4_W1 	, 11 , 1, 3): .Row = 11 : .TypeVAlign = 2
		ret = .AddCellSpan(C_4_W1 	, 15 , 1, 3): .Row = 15 : .TypeVAlign = 2
		ret = .AddCellSpan(C_4_W1 	, 18 , 1, 3): .Row = 18 : .TypeVAlign = 2
		
		' -- 중복과목 보더 
		.SetCellBorder C_4_W1, 4, C_4_DESC1, 6, 16, &H800000, 1 
		.SetCellBorder C_4_W1, 7, C_4_DESC1, 7, 4, &H800000, 1 
		.SetCellBorder C_4_W1, 8, C_4_DESC1, 11, 16, &H800000, 1 
		.SetCellBorder C_4_W1, 11, C_4_DESC1, 13, 16, &H800000, 1 
		.SetCellBorder C_4_W1, 14, C_4_DESC1, 14, 4, &H800000, 1 
		.SetCellBorder C_4_W1, 15, C_4_DESC1, 17, 16, &H800000, 1 
		.SetCellBorder C_4_W1, 18, C_4_DESC1, 20, 16, &H800000, 1 
		.SetCellBorder C_4_W1, 21, C_4_DESC1, 21, 4, &H800000, 1 
		
		Call SetSpreadLock(TYPE_4)
	 	ggoSpread.SSSetSplit2(2) 
		
		.Redraw = True
	End If
	
	End With
		
	' 3번 그리드 
	With lgvspdData(TYPE_3)
	
		.Redraw = False
		ggoSpread.Source = lgvspdData(TYPE_3)
		ggoSpread.InsertRow , 2	

		For iCol = C_3_W2 To C_3_W10 Step 2
			Call MakePercentCol( lgvspdData(TYPE_3), iCol, "", "", "")
		Next
			
		.Row = 1
		.Col = C_3_W_TYPE	: .value = "매출액 비율"
		.Col = C_3_W2		: .value = 1
		If pMode <> "N" Then 
			.Col = 0		: .value = 1
		End If
				
		.Row = 2
		.Col = C_3_W_TYPE	: .value = "개별손금비율"
		.Col = C_3_W2		: .value = 1
		If pMode <> "N" Then 
			.Col = 0		: .value = 2
		End If
		
		Call SetSpreadLock(TYPE_3)
		
		.Redraw = True	
	End With

	' 4번 그리드 변경 
	With lgvspdData(TYPE_4)
	
		.Redraw = False

		For iCol = C_4_W5_1 To C_4_W7 Step 2
			Call MakePercentCol( lgvspdData(TYPE_4), iCol, "", "", "")
		Next
		
		.Redraw = True
	End With	
End Sub

Sub SetSpreadLock(pType)

	With lgvspdData(pType)
	
		ggoSpread.Source = lgvspdData(pType)	

		Select Case pType
			Case TYPE_1 
				ggoSpread.SpreadLock C_1_W1, -1, C_1_W1_CD	' 전체 적용 
			
				ggoSpread.SpreadLock C_1_W5,  1, C_1_W5, 1			' 개별행 
				ggoSpread.SpreadLock C_1_W5,  2, C_1_W5, 2			' 개별행 
				ggoSpread.SpreadLock C_1_W3,  3, C_1_W5, 3			' 개별행 
				ggoSpread.SpreadLock C_1_W3,  4, C_1_W3, 4
				ggoSpread.SpreadLock C_1_W3,  5, C_1_W5, 5
				ggoSpread.SpreadLock C_1_W4_1,  6, C_1_W5, 6
				ggoSpread.SpreadLock C_1_W3,  7, C_1_W5, 7
				ggoSpread.SpreadLock C_1_W3,  8, C_1_W3, 8
				ggoSpread.SpreadLock C_1_W3,  9, C_1_W5, 9
				ggoSpread.SpreadLock C_1_W4_1, 10, C_1_W5, 10
				ggoSpread.SpreadLock C_1_W3,   11, C_1_W3, 11
				ggoSpread.SpreadLock C_1_W3,   12, C_1_W5, 12
				ggoSpread.SpreadLock C_1_W4_1, 13, C_1_W5, 13
				ggoSpread.SpreadLock C_1_W3,   14, C_1_W5, 14
				ggoSpread.SpreadLock C_1_W3,   15, C_1_W3, 15
				ggoSpread.SpreadLock C_1_W3,   16, C_1_W5, 16
				ggoSpread.SpreadLock C_1_W4_1, 17, C_1_W5, 17
				ggoSpread.SpreadLock C_1_W3,   18, C_1_W3, 18
				ggoSpread.SpreadLock C_1_W3,   19, C_1_W5, 19
				ggoSpread.SpreadLock C_1_W4_1, 20, C_1_W5, 20
				ggoSpread.SpreadLock C_1_W3,   21, C_1_W5, 21
				ggoSpread.SpreadLock C_1_W3,   22, C_1_W5, 22
				ggoSpread.SpreadLock C_1_W3,   23, C_1_W5, 23
				ggoSpread.SpreadLock C_1_W3,   24, C_1_W5, 24
				ggoSpread.SpreadLock C_1_W3,   25, C_1_W5, 25

			Case TYPE_2_1
				If .MaxRows > 0 Then

					.Row = .MaxRows			': .Col = C_2_W1
					.Col = C_2_SEQ_NO
					If .Text = "999999" Then	' 기부금한도초과가 없음 
						ggoSpread.SSSetRequired C_2_W1, 1, .MaxRows				
					Else
						ggoSpread.SSSetRequired C_2_W1, 1, .MaxRows	-3	
	'					ggoSpread.SSSetRequired C_2_W1, .MaxRows	-1, .MaxRows	-1
					End If				

				End If
				ggoSpread.SpreadLock C_2_W1_NM,   -1, C_2_W1_NM
			Case TYPE_2_2
				If .MaxRows > 0 Then

					.Row = .MaxRows			': .Col = C_2_W1
					.Col = C_2_SEQ_NO
					If .Text = "999999" Then	' 기부금한도초과가 없음 
						ggoSpread.SSSetRequired C_2_W1, 1, .MaxRows				
					Else
						ggoSpread.SSSetRequired C_2_W1, 1, .MaxRows	-3	
	'					ggoSpread.SSSetRequired C_2_W1, .MaxRows	-1, .MaxRows	-1
					End If				

				End If
			
				ggoSpread.SpreadLock C_2_W1_NM,   -1, C_2_W1_NM
			Case TYPE_3
				ggoSpread.SpreadLock C_3_W_TYPE,   -1, C_3_W10

			Case TYPE_4
				' 전체	: Col 락 
				ggoSpread.SpreadLock C_4_W1,  -1, C_4_W1_CD
				ggoSpread.SpreadLock C_4_W3,   1, C_4_W7, 20
				
				' 개별 
				ggoSpread.SpreadLock C_4_W3,  21, C_4_W4_1, 21
				ggoSpread.SpreadLock C_4_W4_2,  21, C_4_W4_2, 21
				ggoSpread.SpreadLock C_4_W4_3,  21, C_4_W4_3, 21
				ggoSpread.SpreadLock C_4_W4_4,  21, C_4_W4_2, 21
				ggoSpread.SpreadLock C_4_W4_5,  21, C_4_W4_3, 21
				ggoSpread.SpreadLock C_4_W4_6,  21, C_4_W4_2, 21
				
				ggoSpread.SpreadLock C_4_W5_1,  21, C_4_W5_1, 24
				ggoSpread.SpreadLock C_4_W5_2,  21, C_4_W5_2, 24
				ggoSpread.SpreadLock C_4_W5_3,  21, C_4_W5_3, 24
				ggoSpread.SpreadLock C_4_W5_4,  21, C_4_W5_4, 24
				ggoSpread.SpreadLock C_4_W5_5,  21, C_4_W5_5, 24
				ggoSpread.SpreadLock C_4_W5_6,  21, C_4_W5_6, 24
				ggoSpread.SpreadLock C_4_W7,  21, C_4_W7, 24
				
				ggoSpread.SpreadLock C_4_W6,  21, C_4_W6, 21
				ggoSpread.SpreadLock C_4_W3,  25, C_4_W7, 25
		End Select
		
	End With	
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)	

		Select Case pType
			Case TYPE_1
				ggoSpread.SSSetProtected C_W5, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_W8, pvStartRow, pvEndRow
				
			Case TYPE_2_1, TYPE_2_2
				ggoSpread.SSSetRequired C_2_W1, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_2_W1_NM, pvStartRow, pvEndRow
		End Select
			
	End With	
End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_2_1 To TYPE_2_2
		Call SetSpreadLock(iRow)
		
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then

				.Row = .MaxRows		
				.Col = C_2_SEQ_NO 
				''If .Text = "합계" Then	' 기부금한도초과가 없음 
				
				If .text="999999" Then 
					.Row = .MaxRows		: .Col = C_2_W1	: .Value = "합계"
					.Col = C_2_W1_BT	: .CellType = 1
					ggoSpread.SpreadLock C_2_W1,  .MaxRows	, C_2_W4, .MaxRows				
				Else
					.Row = .MaxRows - 2	: .Col = C_2_W1	: .Value = "소계"
					.Col = C_2_W1_BT	: .CellType = 1
					.Row = .MaxRows - 1	: .Col = C_2_W1	: .Value = "기부금한도초과"
					.Col = C_2_W1_BT	: .CellType = 1
					.Row = .MaxRows		: .Col = C_2_W1	: .Value = "합계"
					.Col = C_2_W1_BT	: .CellType = 1
					
					ggoSpread.SpreadLock C_2_W1,  .MaxRows - 2	, C_2_W4, .MaxRows - 2
					ggoSpread.SpreadLock C_2_W1,  .MaxRows		, C_2_W4, .MaxRows 
				End If				

			End If
		End With
	Next
End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W_TYPE	= iCurColumnPos(2)
            C_W13		= iCurColumnPos(3)
            C_W1		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W13		= iCurColumnPos(6)
            C_W3		= iCurColumnPos(7)
            C_W4		= iCurColumnPos(8)
            C_W5		= iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W9		= iCurColumnPos(13)
            C_W_TYPE	= iCurColumnPos(14)
            C_W1		= iCurColumnPos(15)
            C_W2		= iCurColumnPos(16)
    End Select    
End Sub

Function AllGridInit()
	Dim iRow
	
	For iRow = TYPE_1 To TYPE_4
		With lgvspdData(iRow)
		
			ggoSpread.Source = lgvspdData(iRow)
			ggoSpread.ClearSpreadData	' 삭제 
		End With
	Next
	
	Call MakeDefaultGrid("N")
End Function

Sub ChangeAllUpdateFlg(Byval Index)
	Dim iRow, iMaxRows
	With lgvspdData(Index)
		ggoSpread.Source = lgvspdData(Index)
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow
			ggoSpread.UpdateRow iRow
		Next
	End With
End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef4()	' 그리드4의 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrRs
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = ReadRefDoc(TAB4) & vbCrLf & vbCrLf

	' 변경될 위치를 알려줌 
	Dim iCol, iRow
	With lgvspdData(TYPE_4)
		iCol = .ActiveCol	: iRow = .ActiveRow

		.AllowMultiBlocks = True
		.SetSelection C_4_W3, 22, C_4_W3, 25  ' -- 처음 선택할때 
		'.AddSelection C_1_W4_1, -999, C_1_W4_6, -999 ' -- 개별행을 여러개 추가할때 
		
		IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
		
		.SetSelection iCol, iRow, iCol, iRow
		
		If IntRetCD = vbNo Then
			 Exit Function
		End If
	
		Dim IntRetCD1

		IntRetCD1 = CommonQueryRs("W22_3, W23_3, W24_3", "dbo.ufn_TB_48_GetRef4('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

		If IntRetCD1 <> False Then
			.Col = C_4_W3
			
			.Row = 22	: .value = lgF0
			.Row = 23	: .value = lgF1
			.Row = 24	: .value = lgF2
		Else
			Call DisplayMsgBox("W60006", parent.VB_INFORMATION, "", "")
		End If

	End With

	' 22의 비율 넣기 
	Call DuplicateRow(22)
	Call DuplicateRow(23)
	Call DuplicateRow(24)
	
	' 22행 이상의 값을 계산한다.			
	Call ReClacGrid4_22Over
	
	lgvspdData(TYPE_4).focus
End Function

Sub DuplicateRow(Byval pRow)
	Call PutGrid4(C_4_W5_1, pRow, GetGrid(TYPE_4, C_4_W5_1, 21))
	Call PutGrid4(C_4_W5_2, pRow, GetGrid(TYPE_4, C_4_W5_2, 21))
	Call PutGrid4(C_4_W5_3, pRow, GetGrid(TYPE_4, C_4_W5_3, 21))
	Call PutGrid4(C_4_W5_4, pRow, GetGrid(TYPE_4, C_4_W5_4, 21))
	Call PutGrid4(C_4_W5_5, pRow, GetGrid(TYPE_4, C_4_W5_5, 21))
	Call PutGrid4(C_4_W5_6, pRow, GetGrid(TYPE_4, C_4_W5_6, 21))
	Call PutGrid4(C_4_W7, pRow, GetGrid(TYPE_4, C_4_W7, 21))	
End Sub

Function GetRef2()	' 그리드1의 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrRs
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = ReadRefDoc(TAB1) & vbCrLf & vbCrLf

	' 변경될 위치를 알려줌 
	Dim iCol, iRow
	With lgvspdData(TYPE_1)
		iCol = .ActiveCol	: iRow = .ActiveRow

		.AllowMultiBlocks = True
		.SetSelection C_1_W3, 1, C_1_W3, 20  ' -- 처음 선택할때 
		'.AddSelection C_1_W4_1, -999, C_1_W4_6, -999 ' -- 개별행을 여러개 추가할때 
		
		IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
		
		.SetSelection iCol, iRow, iCol, iRow
		
		If IntRetCD = vbNo Then
			 Exit Function
		End If
	End With
	
	' 모든 탭의 그리드 초기화 한다.
	Call AllGridInit
		
    Dim IntRetCD1

	IntRetCD1 = CommonQueryRs("W1", "dbo.ufn_TB_48_GetRef2('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		arrRs = Split(lgF0, chr(11))
		
		With lgvspdData(TYPE_1)
			ggoSpread.Source = lgvspdData(TYPE_1)
			
			
			.Col = C_1_W3
			.Row = 1	: .value = UNICDbl(arrRs(0))
			.Row = 2	: .value = UNICDbl(arrRs(1))
			.Row = 6	: .value = UNICDbl(arrRs(2))
			.Row = 10	: .value = UNICDbl(arrRs(3))
			.Row = 13	: .value = UNICDbl(arrRs(4))
			.Row = 17	: .value = UNICDbl(arrRs(5))
			.Row = 20	: .value = UNICDbl(arrRs(6))
			
			Call vspdData_Change(TYPE_1, C_1_W3, 1)
			Call vspdData_Change(TYPE_1, C_1_W3, 2)
			
			Call ReCalcGrid(TYPE_1)
		End With
	End If
	
	lgvspdData(TYPE_1).focus
End Function

Function GetRef1()	' 감면사업등록 팝업 
	' 2. 팝업 
	Dim arrRet, sParam
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
    
	arrParam(0) = frm1.txtCO_CD.value
	arrParam(1) = frm1.txtFISC_YEAR.text
	arrParam(2) = frm1.cboREP_TYPE.value

	arrRet = window.showModalDialog(BIZ_REF_PGM_ID, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	Else
		Call SetColHead(arrRet)
	End If	
	
	lgvspdData(TYPE_1).focus
End Function


Function GetRef3()	' 그리드2의 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrRs
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = ReadRefDoc(TAB2) & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
			
    Dim IntRetCD1(1), arrW1_CD, arrW1, arrW2, iMaxRows, iRow

	IntRetCD1(0) = CommonQueryRs("W1, W1_NM, W2", "dbo.ufn_TB_48_GetRef3('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "', '1')", "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1(0) <> False Then
		arrW1_CD	= Split(lgF0, chr(11))
		arrW1		= Split(lgF1, chr(11))
		arrW2		= Split(lgF2, chr(11))
		iMaxRows	= UBound(arrW1_CD)
		
		With lgvspdData(TYPE_2_1)
			ggoSpread.Source = lgvspdData(TYPE_2_1)

			ggoSpread.ClearSpreadData	' 삭제 
			ggoSpread.InsertRow , iMaxRows
		
			For iRow = 0 To iMaxRows -1
				.Row = iRow + 1

				.Col = C_2_SEQ_NO	: .Value = iRow + 1
				.Col = C_2_W1		: .Value = arrW1_CD(iRow)
				.Col = C_2_W1_NM	: .Value = arrW1(iRow)
				.Col = C_2_W2		: .Value = arrW2(iRow)						
			Next
			
			ggoSpread.SSSetRequired C_2_W1, -1,-1
			ggoSpread.SSSetProtected C_2_W1_NM, -1,-1
			
			.Row = .MaxRows			: .Col = C_2_W1_NM
			If .Text = "" Then	' 기부금한도초과가 없음 
				.Row = .MaxRows		: .Col = C_2_W1_NM	: .Value = "합계"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO
				.Col = C_2_W1_BT	: .CellType = 1
				ggoSpread.SpreadLock C_2_W1,  .MaxRows	, C_2_W4, .MaxRows				
			ElseIf .Text = "합계" Then
				.Row = .MaxRows - 2	: .Col = C_2_W1_NM	: .Value = "소계"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO
				.Col = C_2_W1_BT	: .CellType = 1
				.Row = .MaxRows - 1	: .Col = C_2_W1_NM	: .Value = "기부금한도초과"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO + 1
				.Col = C_2_W1_BT	: .CellType = 1
				.Row = .MaxRows		: .Col = C_2_W1_NM	: .Value = "합계"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO + 2
				.Col = C_2_W1_BT	: .CellType = 1
				
				ggoSpread.SpreadLock C_2_W1,  .MaxRows - 2	, C_2_W4, .MaxRows - 2
				ggoSpread.SpreadLock C_2_W1,  .MaxRows		, C_2_W4, .MaxRows 
			
			End If
				
		End With
	End If
		
	IntRetCD1(1) = CommonQueryRs("W1, W1_NM, W2", "dbo.ufn_TB_48_GetRef3('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "', '2')", "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1(1) <> False Then
		arrW1_CD	= Split(lgF0, chr(11))
		arrW1		= Split(lgF1, chr(11))
		arrW2		= Split(lgF2, chr(11))
		iMaxRows	= UBound(arrW1_CD)

		With lgvspdData(TYPE_2_2)
			ggoSpread.Source = lgvspdData(TYPE_2_2)

			ggoSpread.ClearSpreadData	' 삭제 
			ggoSpread.InsertRow , iMaxRows
			
			For iRow = 0 To iMaxRows -1
				.Row = iRow + 1
				.Col = C_2_SEQ_NO	: .Value = iRow + 1
				.Col = C_2_W1		: .Value = arrW1_CD(iRow)
				.Col = C_2_W1_NM	: .Value = arrW1(iRow)
				.Col = C_2_W2		: .Value = arrW2(iRow)					
			Next
			ggoSpread.SSSetRequired C_2_W1, -1,-1
			ggoSpread.SSSetProtected C_2_W1_NM, -1,-1
			
			.Row = .MaxRows			: .Col = C_2_W1_NM
			If .Text = "" Then	' 기부금전기이월액이 없음 
				.Row = .MaxRows		: .Col = C_2_W1_NM	: .Value = "합계"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO
				.Col = C_2_W1_BT	: .CellType = 1
				ggoSpread.SpreadLock C_2_W1,  .MaxRows	, C_2_W4, .MaxRows	
			Else
				.Row = .MaxRows - 2	: .Col = C_2_W1_NM	: .Value = "소계"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO
				.Col = C_2_W1_BT	: .CellType = 1
				.Row = .MaxRows - 1	: .Col = C_2_W1_NM	: .Value = "기부금전기이월액"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO+1
				.Col = C_2_W1_BT	: .CellType = 1
				.Row = .MaxRows		: .Col = C_2_W1_NM	: .Value = "합계"
				.Col = C_2_SEQ_NO	: .Value = SUM_SEQ_NO+2
				.Col = C_2_W1_BT	: .CellType = 1
					
				ggoSpread.SpreadLock C_2_W1,  .MaxRows - 2	, C_2_W4, .MaxRows - 2
				ggoSpread.SpreadLock C_2_W1,  .MaxRows		, C_2_W4, .MaxRows 
			End If
				
		End With
	End If
	
	If IntRetCD1(0) = False And IntRetCD1(1) = False Then	
		Call DisplayMsgBox("900014", parent.VB_INFORMATION, "", "")
	End If

	lgvspdData(lgCurrGrid).focus
End Function

Sub RedrawGrid2TotalLine()

End Sub

' -- 세무서식 메시지가 |로 분리되어 있다.
Function ReadRefDoc(pTab)
	Dim arrRefDoc
	arrRefDoc	= Split(wgRefDoc, "|")
	ReadRefDoc	= arrRefDoc(pTab-1)
End Function

Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.

		
End Sub

' -- 감면사업 그리드에 표현 
Sub SetColHead(parrRet)
	Dim iCol, iMaxCols
	iMaxCols = UBound(parrRet)
	' -- 1번 그리드 
	With lgvspdData(TYPE_1)
		.Row = -999
		For iCol = 0 To iMaxCols -1
			
			.Col = C_1_W4_1 + iCol 
			.text = parrRet(iCol)
			lgW_NM(iCol) = parrRet(iCol)
		Next
	End With
	
	If parrRet(3) <> "" Or parrRet(4) <> "" Or parrRet(5) <> ""  Then 
		Call ShowColHidden(TYPE_1)
	Else
		Call NotShowColHidden(TYPE_1)
	End If
	
	' -- 3번 그리드 
	With lgvspdData(TYPE_3)
		.Row = 0
		For iCol = 0 To iMaxCols -1
			
			.Col = C_3_W3_1 + (iCol * 2)
			.text = parrRet(iCol)
		Next
	End With
	
	If parrRet(3) <> "" Or parrRet(4) <> "" Or parrRet(5) <> ""  Then 
		Call ShowColHidden(TYPE_3)
	Else
		Call NotShowColHidden(TYPE_3)
	End If	
	
	' -- 4번 그리드 
	With lgvspdData(TYPE_4)
		.Row = -999
		For iCol = 0 To iMaxCols -1
			
			.Col = C_4_W4_1 + (iCol * 2)
			.text = parrRet(iCol)
		Next
	End With
	
	If parrRet(3) <> "" Or parrRet(4) <> "" Or parrRet(5) <> ""  Then 
		Call ShowColHidden(TYPE_4)
	Else
		Call NotShowColHidden(TYPE_4)
	End If
End Sub

Sub ShowColHidden(pType)
	Select Case pType
		Case TYPE_1

			With lgvspdData(pType)
				.BlockMode = True
				.Row	= -1
				.Row2	= -1
				.Col	= C_1_W4_4
				.Col2	= C_1_W4_6
				.ColHidden = False
				.BlockMode = False
			End With
		
		Case TYPE_3
		
			With lgvspdData(pType)
				.BlockMode = True
				.Row	= -1
				.Row2	= -1
				.Col	= C_3_W3_2
				.Col2	= C_3_W8_2
				.ColHidden = False
				.BlockMode = False
			End With	

		Case TYPE_4
		
			With lgvspdData(pType)
				.BlockMode = True
				.Row	= -1
				.Row2	= -1
				.Col	= C_4_W4_4
				.Col2	= C_4_W5_6
				.ColHidden = False
				.BlockMode = False
			End With	
	End Select
End Sub

Sub NotShowColHidden(pType)
	Select Case pType
		Case TYPE_1

			With lgvspdData(pType)
				.BlockMode = True
				.Row	= -1
				.Row2	= -1
				.Col	= C_1_W4_4
				.Col2	= C_1_W4_6
				.ColHidden = True
				.BlockMode = False
			End With
		
		Case TYPE_3
		
			With lgvspdData(pType)
				.BlockMode = True
				.Row	= -1
				.Row2	= -1
				.Col	= C_3_W3_2
				.Col2	= C_3_W8_2
				.ColHidden = True
				.BlockMode = False
			End With	

		Case TYPE_4
		
			With lgvspdData(pType)
				.BlockMode = True
				.Row	= -1
				.Row2	= -1
				.Col	= C_4_W4_4
				.Col2	= C_4_W5_6
				.ColHidden = True
				.BlockMode = False
			End With					
	End Select
End Sub

Function SetComboVal()	' 감면분 데이타와 2개의 정해진 데이타로 콤보를 생성한다.
	Dim iRow, iMaxRows
	iMaxRows = UBOund(lgW_NM)
	
	For iRow = 0 To iMaxRows -1
		If Trim(lgW_NM(iRow)) = "" Then
			lgW_NM(iRow) = "기타"
			lgW_NM(iRow+1) = "공통분"
			Exit Function
		End If
	Next
End Function

' -- 탭별 링크 보여주기 
Function ShowTabLInk(pType)
	Dim pObj1, pObj2, i
	Set pObj1 = document.all("myTabRef")
	Set pObj2 = document.all("myTabRef2")
	
	For i = 0 To 3
		pObj1(i).style.display = "none"
		pObj2(i).style.display = "none"
	Next
	
	pObj1(pType-1).style.display = ""
	pObj2(pType-1).style.display = ""
End Function

Function ChkChgTab()
	ChkChgTab = False
	' 1. 감면 세액 로딩 체크 
	With lgvspdData(TYPE_1)
		.Col = C_1_W4_1
		.Row = -999
		If .Text = "" Then
			Call DisplayMsgBox("W60002", "X", "1. 감면사업 등록", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	End With
	ChkChgTab = True
End Function

Function ChkChgTab3()	' 3번 그리드 조건 체크 
	ChkChgTab3 = False
	If frm1.cboW11.value = "" Then
		Call DisplayMsgBox("W60002", "X", "3. 공통비용 배부율 등록", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	ChkChgTab3 = True
End Function

' -- 탭2의 구분2 콤보(다이나믹 ㅠ.ㅠ )
Function InitSpreadComboBox2()
	Dim sTmp
	
	Call SetComboVal  ' -- 감면분 배열에 기타/공통분은 추가함 
	
	sTmp = MakeSpreadCombo(lgW_NM)	' inc_CliGrid 에 정의됨 
	
	ggoSpread.Source = lgvspdData(TYPE_2_1)
	ggoSpread.SetCombo sTmp, C_2_W4
			
	ggoSpread.Source = lgvspdData(TYPE_2_2)
	ggoSpread.SetCombo sTmp, C_2_W4	

End Function

'====================================== 탭 함수 =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1	' 기본 그리드 
	Call ShowTabLInk(TAB1)

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetToolbar("1101000000000111")										<%'버튼 툴바 제어 %>
	Else
		Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
	End If
		
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	' Tab1 조건 체크후 이상없으면 진행 
	If Not ChkChgTab Then Exit Function

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetToolbar("1101011100000111")										<%'버튼 툴바 제어 %>
	Else
		Call SetToolbar("1100010100000111")										<%'버튼 툴바 제어 %>
	End If
	
	' 1. 탭2,3 클릭시 구분2 콤보 생성 
	With lgvspdData(TYPE_2_1)
		.Col = C_2_W4	: .Row = 1
		If .TypeComboBoxCount = 0 Then	'  콤보값이 없다면 
			Call  InitSpreadComboBox2
		End If
	End With
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_2_1
	Call ShowTabLInk(TAB2)
End Function

Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function
	
	' Tab1 조건 체크후 이상없으면 진행 
	If Not ChkChgTab Then Exit Function

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetToolbar("1101000000000111")										<%'버튼 툴바 제어 %>
	Else
		Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
	End If
	
	Call SetGrid4ByGrid2()	' -- 탭1,탭2를 데이타로 탭4를 생성 
	Call SetGrid3Data		' -- 탭3 생성 
	
	Call changeTabs(TAB3)
	gSelframeFlg = TAB3
	lgCurrGrid = TYPE_3
	Call ShowTabLInk(TAB3)
End Function

Function ClickTab4()

	If gSelframeFlg = TAB4 Then Exit Function

	' Tab1 조건 체크후 이상없으면 진행 
	If Not ChkChgTab Then Exit Function
	If Not ChkChgTab3 Then Exit Function

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
	Else
		Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
	End If
		
	Call changeTabs(TAB4)
	gSelframeFlg = TAB4
	lgCurrGrid = TYPE_4
	Call ShowTabLInk(TAB4)
	
	Call SetGrid4Data()
End Function

' -- 탭4번 그리드 테이타 출력 
Function SetGrid4Data()
	Dim iRow, iCol

	' 3번 그리드를 갱신한후 (이유: 탭1을 고치고 탭3을 안누른후 탭4로 올 경우)
	Call SetGrid4ByGrid2
	Call SetGrid3Data
	
	Dim dblSum, dblAmt(30)
	
	With lgvspdData(TYPE_4)
	
		' 05,09,12,16,19
		' ④금 액 에는 ( ③합계 × ⑤비율 )을 계산하여 입력하고 ⑥금액은 ( ③금액 - ④금 액 )를 계산하여 입력합니다.
		Call ReClacGrid4(5)
		Call ReClacGrid4(9)
		Call ReClacGrid4(12)
		Call ReClacGrid4(16)
		Call ReClacGrid4(19)
		
		' 03 각 열의 01 - 02
		Call PutGrid4(C_4_W3, 3, UNICDbl(GetGrid(TYPE_4, C_4_W3, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 2)) )
		Call PutGrid4(C_4_W4_1, 3, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 2)) )
		Call PutGrid4(C_4_W4_2, 3, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 2)) )
		Call PutGrid4(C_4_W4_3, 3, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 2)) )
		Call PutGrid4(C_4_W4_4, 3, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 2)) )
		Call PutGrid4(C_4_W4_5, 3, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 2)) )
		Call PutGrid4(C_4_W4_6, 3, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 2)) )
		Call PutGrid4(C_4_W6, 3, UNICDbl(GetGrid(TYPE_4, C_4_W6, 1)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 2)) )
		' 06 각 열의 04 + 05
		Call PutGrid4(C_4_W3, 6, UNICDbl(GetGrid(TYPE_4, C_4_W3, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W3, 5)) )
		Call PutGrid4(C_4_W4_1, 6, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 5)) )
		Call PutGrid4(C_4_W4_2, 6, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 5)) )
		Call PutGrid4(C_4_W4_3, 6, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 5)) )
		Call PutGrid4(C_4_W4_4, 6, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 5)) )
		Call PutGrid4(C_4_W4_5, 6, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 5)) )
		Call PutGrid4(C_4_W4_6, 6, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 5)) )
		Call PutGrid4(C_4_W6, 6, UNICDbl(GetGrid(TYPE_4, C_4_W6, 4)) + UNICDbl(GetGrid(TYPE_4, C_4_W6, 5)) )
		' 07 각 열의 03 - 06
		Call PutGrid4(C_4_W3, 7, UNICDbl(GetGrid(TYPE_4, C_4_W3, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 6)) )
		Call PutGrid4(C_4_W4_1, 7, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 6)) )
		Call PutGrid4(C_4_W4_2, 7, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 6)) )
		Call PutGrid4(C_4_W4_3, 7, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 6)) )
		Call PutGrid4(C_4_W4_4, 7, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 6)) )
		Call PutGrid4(C_4_W4_5, 7, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 6)) )
		Call PutGrid4(C_4_W4_6, 7, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 6)) )
		Call PutGrid4(C_4_W6, 7, UNICDbl(GetGrid(TYPE_4, C_4_W6, 3)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 6)) )
		' 10 각 열의 08 + 09
		Call PutGrid4(C_4_W3,10, UNICDbl(GetGrid(TYPE_4, C_4_W3, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W3, 9)) )
		Call PutGrid4(C_4_W4_1,10, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 9)) )
		Call PutGrid4(C_4_W4_2,10, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 9)) )
		Call PutGrid4(C_4_W4_3,10, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 9)) )
		Call PutGrid4(C_4_W4_4,10, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 9)) )
		Call PutGrid4(C_4_W4_5,10, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 9)) )
		Call PutGrid4(C_4_W4_6,10, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 9)) )
		Call PutGrid4(C_4_W6,10, UNICDbl(GetGrid(TYPE_4, C_4_W6, 8)) + UNICDbl(GetGrid(TYPE_4, C_4_W6, 9)) )
		' 13 각 열의 11 + 12
		Call PutGrid4(C_4_W3,13, UNICDbl(GetGrid(TYPE_4, C_4_W3,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W3,12)) )
		Call PutGrid4(C_4_W4_1,13, UNICDbl(GetGrid(TYPE_4, C_4_W4_1,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_1,12)) )
		Call PutGrid4(C_4_W4_2,13, UNICDbl(GetGrid(TYPE_4, C_4_W4_2,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2,12)) )
		Call PutGrid4(C_4_W4_3,13, UNICDbl(GetGrid(TYPE_4, C_4_W4_3,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3,12)) )
		Call PutGrid4(C_4_W4_4,13, UNICDbl(GetGrid(TYPE_4, C_4_W4_4,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_4,12)) )
		Call PutGrid4(C_4_W4_5,13, UNICDbl(GetGrid(TYPE_4, C_4_W4_5,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5,12)) )
		Call PutGrid4(C_4_W4_6,13, UNICDbl(GetGrid(TYPE_4, C_4_W4_6,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6,12)) )
		Call PutGrid4(C_4_W6,13, UNICDbl(GetGrid(TYPE_4, C_4_W6,11)) + UNICDbl(GetGrid(TYPE_4, C_4_W6,12)) )
		' 14 각 열의 07 + 10 - 13
		Call PutGrid4(C_4_W3,14, UNICDbl(GetGrid(TYPE_4, C_4_W3, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W3,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W3,13)) )
		Call PutGrid4(C_4_W4_1,14, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_1,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1,13)) )
		Call PutGrid4(C_4_W4_2,14, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2,13)) )
		Call PutGrid4(C_4_W4_3,14, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3,13)) )
		Call PutGrid4(C_4_W4_4,14, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_4,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4,13)) )
		Call PutGrid4(C_4_W4_5,14, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5,13)) )
		Call PutGrid4(C_4_W4_6,14, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6,13)) )
		Call PutGrid4(C_4_W6,14, UNICDbl(GetGrid(TYPE_4, C_4_W6, 7)) + UNICDbl(GetGrid(TYPE_4, C_4_W6,10)) - UNICDbl(GetGrid(TYPE_4, C_4_W6,13)) )
		' 17 각 열의 15 + 16
		Call PutGrid4(C_4_W3,17, UNICDbl(GetGrid(TYPE_4, C_4_W3,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W3,16)) )
		Call PutGrid4(C_4_W4_1,17, UNICDbl(GetGrid(TYPE_4, C_4_W4_1,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_1,16)) )
		Call PutGrid4(C_4_W4_2,17, UNICDbl(GetGrid(TYPE_4, C_4_W4_2,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2,16)) )
		Call PutGrid4(C_4_W4_3,17, UNICDbl(GetGrid(TYPE_4, C_4_W4_3,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3,16)) )
		Call PutGrid4(C_4_W4_4,17, UNICDbl(GetGrid(TYPE_4, C_4_W4_4,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_4,16)) )
		Call PutGrid4(C_4_W4_5,17, UNICDbl(GetGrid(TYPE_4, C_4_W4_5,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5,16)) )
		Call PutGrid4(C_4_W4_6,17, UNICDbl(GetGrid(TYPE_4, C_4_W4_6,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6,16)) )
		Call PutGrid4(C_4_W6,17, UNICDbl(GetGrid(TYPE_4, C_4_W6,15)) + UNICDbl(GetGrid(TYPE_4, C_4_W6,16)) )
		' 20 각 열의 18 + 19
		Call PutGrid4(C_4_W3,20, UNICDbl(GetGrid(TYPE_4, C_4_W3,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W3,19)) )
		Call PutGrid4(C_4_W4_1,20, UNICDbl(GetGrid(TYPE_4, C_4_W4_1,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_1,19)) )
		Call PutGrid4(C_4_W4_2,20, UNICDbl(GetGrid(TYPE_4, C_4_W4_2,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2,19)) )
		Call PutGrid4(C_4_W4_3,20, UNICDbl(GetGrid(TYPE_4, C_4_W4_3,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3,19)) )
		Call PutGrid4(C_4_W4_4,20, UNICDbl(GetGrid(TYPE_4, C_4_W4_4,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_4,19)) )
		Call PutGrid4(C_4_W4_5,20, UNICDbl(GetGrid(TYPE_4, C_4_W4_5,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5,19)) )
		Call PutGrid4(C_4_W4_6,20, UNICDbl(GetGrid(TYPE_4, C_4_W4_6,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6,19)) )
		Call PutGrid4(C_4_W6,20, UNICDbl(GetGrid(TYPE_4, C_4_W6,18)) + UNICDbl(GetGrid(TYPE_4, C_4_W6,19)) )
		' 21 각 열의 14 + 17 - 20
		Call PutGrid4(C_4_W3,21, UNICDbl(GetGrid(TYPE_4, C_4_W3,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W3,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W3,20)) )
		Call PutGrid4(C_4_W4_1,21, UNICDbl(GetGrid(TYPE_4, C_4_W4_1,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_1,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1,20)) )
		Call PutGrid4(C_4_W4_2,21, UNICDbl(GetGrid(TYPE_4, C_4_W4_2,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2,20)) )
		Call PutGrid4(C_4_W4_3,21, UNICDbl(GetGrid(TYPE_4, C_4_W4_3,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3,20)) )
		Call PutGrid4(C_4_W4_4,21, UNICDbl(GetGrid(TYPE_4, C_4_W4_4,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_4,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4,20)) )
		Call PutGrid4(C_4_W4_5,21, UNICDbl(GetGrid(TYPE_4, C_4_W4_5,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5,20)) )
		Call PutGrid4(C_4_W4_6,21, UNICDbl(GetGrid(TYPE_4, C_4_W4_6,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6,20)) )
		Call PutGrid4(C_4_W6,21, UNICDbl(GetGrid(TYPE_4, C_4_W6,14)) + UNICDbl(GetGrid(TYPE_4, C_4_W6,17)) - UNICDbl(GetGrid(TYPE_4, C_4_W6,20)) )
		' 21 각 열중 비율을 구한다.
		dblAmt(C_4_W3) = UNICDbl(GetGrid(TYPE_4, C_4_W3,21))
		dblAmt(C_4_W4_1) = UNICDbl(GetGrid(TYPE_4, C_4_W4_1,21))
		dblAmt(C_4_W4_2) = UNICDbl(GetGrid(TYPE_4, C_4_W4_2,21))
		dblAmt(C_4_W4_3) = UNICDbl(GetGrid(TYPE_4, C_4_W4_3,21))
		dblAmt(C_4_W4_4) = UNICDbl(GetGrid(TYPE_4, C_4_W4_4,21))
		dblAmt(C_4_W4_5) = UNICDbl(GetGrid(TYPE_4, C_4_W4_5,21))
		dblAmt(C_4_W4_6) = UNICDbl(GetGrid(TYPE_4, C_4_W4_6,21))
		dblAmt(C_4_W6) = UNICDbl(GetGrid(TYPE_4, C_4_W6,21))
		
		If dblAmt(C_4_W3) <> 0 Then
			dblAmt(C_4_W5_1) = dblAmt(C_4_W4_1) / dblAmt(C_4_W3)
			dblAmt(C_4_W5_2) = dblAmt(C_4_W4_2) / dblAmt(C_4_W3)
			dblAmt(C_4_W5_3) = dblAmt(C_4_W4_3) / dblAmt(C_4_W3)
			dblAmt(C_4_W5_4) = dblAmt(C_4_W4_4) / dblAmt(C_4_W3)
			dblAmt(C_4_W5_5) = dblAmt(C_4_W4_5) / dblAmt(C_4_W3)
			dblAmt(C_4_W5_6) = dblAmt(C_4_W4_6) / dblAmt(C_4_W3)
			dblAmt(C_4_W7) = dblAmt(C_4_W6) / dblAmt(C_4_W3)
		End If
		
		' 21의 비율 넣기 
		Call PutGrid4(C_4_W5_1, 21, dblAmt(C_4_W5_1))
		Call PutGrid4(C_4_W5_2, 21, dblAmt(C_4_W5_2))
		Call PutGrid4(C_4_W5_3, 21, dblAmt(C_4_W5_3))
		Call PutGrid4(C_4_W5_4, 21, dblAmt(C_4_W5_4))
		Call PutGrid4(C_4_W5_5, 21, dblAmt(C_4_W5_5))
		Call PutGrid4(C_4_W5_6, 21, dblAmt(C_4_W5_6))
		Call PutGrid4(C_4_W7, 21, dblAmt(C_4_W7))

		' 25행 결과 계산 
		Call PutGrid4(C_4_W3, 25, UNICDbl(GetGrid(TYPE_4, C_4_W3, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 24)) )
		Call PutGrid4(C_4_W4_1, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 24)) )
		Call PutGrid4(C_4_W4_2, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 24)) )
		Call PutGrid4(C_4_W4_3, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 24)) )
		Call PutGrid4(C_4_W4_4, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 24)) )
		Call PutGrid4(C_4_W4_5, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 24)) )
		Call PutGrid4(C_4_W4_6, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 24)) )
		Call PutGrid4(C_4_W6, 25, UNICDbl(GetGrid(TYPE_4, C_4_W6, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 24)) )
	
	End With
	
	'Call ReClacGrid4_22Over
End Function

' 탭4의 22행 이상 계산 : 금액불러오기 클릭시 호출 
Sub ReClacGrid4_22Over()
	Dim iRow
	For iRow = 22 To 24
		Call ReClacGrid4_22Over_2(iRow)
	Next 
	
	' 25행 결과 계산 
	Call PutGrid4(C_4_W3, 25, UNICDbl(GetGrid(TYPE_4, C_4_W3, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 24)) )
	Call PutGrid4(C_4_W4_1, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 24)) )
	Call PutGrid4(C_4_W4_2, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 24)) )
	Call PutGrid4(C_4_W4_3, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 24)) )
	Call PutGrid4(C_4_W4_4, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 24)) )
	Call PutGrid4(C_4_W4_5, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 24)) )
	Call PutGrid4(C_4_W4_6, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 24)) )
	Call PutGrid4(C_4_W6, 25, UNICDbl(GetGrid(TYPE_4, C_4_W6, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 24)) )
	
End Sub

Sub ReClacGrid4_22Over_2(Byval pRow)
	Dim dblAmt(30), iRow
	With lgvspdData(TYPE_4)
		dblAmt(C_4_W3) = UNICDbl(GetGrid(TYPE_4, C_4_W3,pRow))
		dblAmt(C_4_W5_1) = UNICDbl(GetGrid(TYPE_4, C_4_W5_1,pRow))
		dblAmt(C_4_W5_2) = UNICDbl(GetGrid(TYPE_4, C_4_W5_2,pRow))
		dblAmt(C_4_W5_3) = UNICDbl(GetGrid(TYPE_4, C_4_W5_3,pRow))
		dblAmt(C_4_W5_4) = UNICDbl(GetGrid(TYPE_4, C_4_W5_4,pRow))
		dblAmt(C_4_W5_5) = UNICDbl(GetGrid(TYPE_4, C_4_W5_5,pRow))
		dblAmt(C_4_W5_6) = UNICDbl(GetGrid(TYPE_4, C_4_W5_6,pRow))
		dblAmt(C_4_W7) = UNICDbl(GetGrid(TYPE_4, C_4_W7,pRow))
		' 3 * 5_1 = 4_1
		dblAmt(C_4_W4_1) = dblAmt(C_4_W3) * dblAmt(C_4_W5_1)
		dblAmt(C_4_W4_2) = dblAmt(C_4_W3) * dblAmt(C_4_W5_2)
		dblAmt(C_4_W4_3) = dblAmt(C_4_W3) * dblAmt(C_4_W5_3)
		dblAmt(C_4_W4_4) = dblAmt(C_4_W3) * dblAmt(C_4_W5_4)
		dblAmt(C_4_W4_5) = dblAmt(C_4_W3) * dblAmt(C_4_W5_5)
		dblAmt(C_4_W4_6) = dblAmt(C_4_W3) * dblAmt(C_4_W5_6)
		dblAmt(C_4_W6) = dblAmt(C_4_W3) * dblAmt(C_4_W7)
		
		Call PutGrid4(C_4_W4_1, pRow, dblAmt(C_4_W4_1))
		Call PutGrid4(C_4_W4_2, pRow, dblAmt(C_4_W4_2))
		Call PutGrid4(C_4_W4_3, pRow, dblAmt(C_4_W4_3))
		Call PutGrid4(C_4_W4_4, pRow, dblAmt(C_4_W4_4))
		Call PutGrid4(C_4_W4_5, pRow, dblAmt(C_4_W4_5))
		Call PutGrid4(C_4_W4_6, pRow, dblAmt(C_4_W4_6))
		Call PutGrid4(C_4_W6, pRow, dblAmt(C_4_W6))
		
	End With
End Sub

Sub ReClacGrid4(Byval pRow)
	Dim dblSum
	' 05,09,12,16,19
	' ④금 액 에는 ( ③합계 × ⑤비율 )을 계산하여 입력하고 ⑥금액은 ( ③금액 - ④금 액 )를 계산하여 입력합니다.
	Call PutGrid4(C_4_W4_1, pRow, UNICDbl(GetGrid(TYPE_4, C_4_W3, pRow)) * UNICDbl(GetGrid(TYPE_4, C_4_W5_1, pRow)) )
	Call PutGrid4(C_4_W4_2, pRow, UNICDbl(GetGrid(TYPE_4, C_4_W3, pRow)) * UNICDbl(GetGrid(TYPE_4, C_4_W5_2, pRow)) )
	Call PutGrid4(C_4_W4_3, pRow, UNICDbl(GetGrid(TYPE_4, C_4_W3, pRow)) * UNICDbl(GetGrid(TYPE_4, C_4_W5_3, pRow)) )
	Call PutGrid4(C_4_W4_4, pRow, UNICDbl(GetGrid(TYPE_4, C_4_W3, pRow)) * UNICDbl(GetGrid(TYPE_4, C_4_W5_4, pRow)) )
	Call PutGrid4(C_4_W4_5, pRow, UNICDbl(GetGrid(TYPE_4, C_4_W3, pRow)) * UNICDbl(GetGrid(TYPE_4, C_4_W5_5, pRow)) )
	Call PutGrid4(C_4_W4_6, pRow, UNICDbl(GetGrid(TYPE_4, C_4_W3, pRow)) * UNICDbl(GetGrid(TYPE_4, C_4_W5_6, pRow)) )
	
	dblSum = UNICDbl(GetGrid(TYPE_4, C_4_W4_1, pRow)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_2, pRow)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_3, pRow))
	dblSum = dblSum + UNICDbl(GetGrid(TYPE_4, C_4_W4_4, pRow)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_5, pRow)) + UNICDbl(GetGrid(TYPE_4, C_4_W4_6, pRow))
	
	Call PutGrid4(C_4_W6, pRow, UNICDbl(GetGrid(TYPE_4, C_4_W3, pRow)) - dblSum )
	
End Sub

' iRow1 행의 값을 읽어서 iRow2에 썸한다.
Function PutSum(Byval iRow1, Byval iRow2)
	With lgvspdData(TYPE_4)
		.Row = iRow2 
		.Col = C_4_W3	: .Value = GetGrid(TYPE_4, C_4_W3, 14)
	End With
End Function

' 탭4 그리드에 데이타 넣기 
Function PutGrid4(Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(TYPE_4)
		.Col = pCol	: .Row = pRow : .Value = pVal
	End With
End Function

' 해당 그리드에서 데이타가져오기 
Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : GetGrid = .Value
	End With
End Function

Function PutGrid(BYval pType, Byval pCol, BYval pRow, Byval pVal)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow 
		If pVal <> "0" Then .Text = pVal
	End With
End Function

Function PutGrid2(BYval pType, Byval pCol, BYval pRow, Byval pVal)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : .Value = pVal
	End With
End Function

' -- 탭3번 그리드 데이타 출력 
Function SetGrid3Data()
	Dim dblW_4_3 , dblW_4_4_1, dblW_4_4_2, dblW_4_4_3, dblW_4_4_4, dblW_4_4_5, dblW_4_4_6, dblW_4_6
	
	' -- 매출액 비율 
	With lgvspdData(TYPE_4)
		.Row = 1	' -- 매출액 
		.Col = C_4_W3	: dblW_4_3		= UNICDbl(.value)
		.Col = C_4_W4_1	: dblW_4_4_1	= UNICDbl(.value)
		.Col = C_4_W4_2	: dblW_4_4_2	= UNICDbl(.value)
		.Col = C_4_W4_3	: dblW_4_4_3	= UNICDbl(.value)
		.Col = C_4_W4_4	: dblW_4_4_4	= UNICDbl(.value)
		.Col = C_4_W4_5	: dblW_4_4_5	= UNICDbl(.value)
		.Col = C_4_W4_6	: dblW_4_4_6	= UNICDbl(.value)
		.Col = C_4_W6	: dblW_4_6		= UNICDbl(.value)
	End With
	
	If dblW_4_3 = 0 Then Exit Function
	
	With lgvspdData(TYPE_3)
		.Row = 1
		.Col = C_3_W1	: .value = dblW_4_3
		.Col = C_3_W3_1	: .value = dblW_4_4_1
		.Col = C_3_W5_1	: .value = dblW_4_4_2
		.Col = C_3_W7_1	: .value = dblW_4_4_3
		.Col = C_3_W3_2	: .value = dblW_4_4_4
		.Col = C_3_W5_2	: .value = dblW_4_4_5
		.Col = C_3_W7_2	: .value = dblW_4_4_6	
		.Col = C_3_W9	: .value = dblW_4_6
		
		.Col = C_3_W4_1	: .value = dblW_4_4_1 / dblW_4_3
		.Col = C_3_W6_1	: .value = dblW_4_4_2 / dblW_4_3
		.Col = C_3_W8_1	: .value = dblW_4_4_3 / dblW_4_3
		.Col = C_3_W4_2	: .value = dblW_4_4_4 / dblW_4_3
		.Col = C_3_W6_2	: .value = dblW_4_4_5 / dblW_4_3
		.Col = C_3_W8_2	: .value = dblW_4_4_6 / dblW_4_3	
		.Col = C_3_W10	: .value = dblW_4_6   / dblW_4_3
							
	End With	

	' 2행의 값으로 비율생성 
	Call SetGrid3DataRow2 
	
	' 탭4에 결과값 반영 
	With lgvspdData(TYPE_3)
		' -- 배부기준 - 매출액 
		Call PutGrid4(C_4_W5_1, 1, GetGrid(TYPE_3, C_3_W4_1, 1))
		Call PutGrid4(C_4_W5_2, 1, GetGrid(TYPE_3, C_3_W6_1, 1))
		Call PutGrid4(C_4_W5_3, 1, GetGrid(TYPE_3, C_3_W8_1, 1))
		Call PutGrid4(C_4_W5_4, 1, GetGrid(TYPE_3, C_3_W4_2, 1))
		Call PutGrid4(C_4_W5_5, 1, GetGrid(TYPE_3, C_3_W6_2, 1))
		Call PutGrid4(C_4_W5_6, 1, GetGrid(TYPE_3, C_3_W8_2, 1))
		Call PutGrid4(C_4_W7, 1, GetGrid(TYPE_3, C_3_W10, 1))
		
		Call PutGrid4(C_4_W5_1, 9, GetGrid(TYPE_3, C_3_W4_1, 1))
		Call PutGrid4(C_4_W5_2, 9, GetGrid(TYPE_3, C_3_W6_1, 1))
		Call PutGrid4(C_4_W5_3, 9, GetGrid(TYPE_3, C_3_W8_1, 1))
		Call PutGrid4(C_4_W5_4, 9, GetGrid(TYPE_3, C_3_W4_2, 1))
		Call PutGrid4(C_4_W5_5, 9, GetGrid(TYPE_3, C_3_W6_2, 1))
		Call PutGrid4(C_4_W5_6, 9, GetGrid(TYPE_3, C_3_W8_2, 1))
		Call PutGrid4(C_4_W7, 9, GetGrid(TYPE_3, C_3_W10, 1))
		
		Call PutGrid4(C_4_W5_1,16, GetGrid(TYPE_3, C_3_W4_1, 1))
		Call PutGrid4(C_4_W5_2,16, GetGrid(TYPE_3, C_3_W6_1, 1))
		Call PutGrid4(C_4_W5_3,16, GetGrid(TYPE_3, C_3_W8_1, 1))
		Call PutGrid4(C_4_W5_4,16, GetGrid(TYPE_3, C_3_W4_2, 1))
		Call PutGrid4(C_4_W5_5,16, GetGrid(TYPE_3, C_3_W6_2, 1))
		Call PutGrid4(C_4_W5_6,16, GetGrid(TYPE_3, C_3_W8_2, 1))
		Call PutGrid4(C_4_W7,16, GetGrid(TYPE_3, C_3_W10, 1))
		
		' -- 배부기준 - 개별손금 
		Dim iRow
		iRow = UNICDbl(frm1.cboW11.value)
		Call PutGrid4(C_4_W5_1, 5, GetGrid(TYPE_3, C_3_W4_1, iRow))
		Call PutGrid4(C_4_W5_2, 5, GetGrid(TYPE_3, C_3_W6_1, iRow))
		Call PutGrid4(C_4_W5_3, 5, GetGrid(TYPE_3, C_3_W8_1, iRow))
		Call PutGrid4(C_4_W5_4, 5, GetGrid(TYPE_3, C_3_W4_2, iRow))
		Call PutGrid4(C_4_W5_5, 5, GetGrid(TYPE_3, C_3_W6_2, iRow))
		Call PutGrid4(C_4_W5_6, 5, GetGrid(TYPE_3, C_3_W8_2, iRow))
		Call PutGrid4(C_4_W7, 5, GetGrid(TYPE_3, C_3_W10, iRow))
		
		Call PutGrid4(C_4_W5_1,12, GetGrid(TYPE_3, C_3_W4_1, iRow))
		Call PutGrid4(C_4_W5_2,12, GetGrid(TYPE_3, C_3_W6_1, iRow))
		Call PutGrid4(C_4_W5_3,12, GetGrid(TYPE_3, C_3_W8_1, iRow))
		Call PutGrid4(C_4_W5_4,12, GetGrid(TYPE_3, C_3_W4_2, iRow))
		Call PutGrid4(C_4_W5_5,12, GetGrid(TYPE_3, C_3_W6_2, iRow))
		Call PutGrid4(C_4_W5_6,12, GetGrid(TYPE_3, C_3_W8_2, iRow))
		Call PutGrid4(C_4_W7,12, GetGrid(TYPE_3, C_3_W10, iRow))
		
		Call PutGrid4(C_4_W5_1,19, GetGrid(TYPE_3, C_3_W4_1, iRow))
		Call PutGrid4(C_4_W5_2,19, GetGrid(TYPE_3, C_3_W6_1, iRow))
		Call PutGrid4(C_4_W5_3,19, GetGrid(TYPE_3, C_3_W8_1, iRow))
		Call PutGrid4(C_4_W5_4,19, GetGrid(TYPE_3, C_3_W4_2, iRow))
		Call PutGrid4(C_4_W5_5,19, GetGrid(TYPE_3, C_3_W6_2, iRow))
		Call PutGrid4(C_4_W5_6,19, GetGrid(TYPE_3, C_3_W8_2, iRow))
		Call PutGrid4(C_4_W7,19, GetGrid(TYPE_3, C_3_W10, iRow))		
	End With
End Function

' -- 탭3 그리드 2행 출력 
Function SetGrid3DataRow2()	
	Dim dblW_4_3 , dblW_4_4_1, dblW_4_4_2, dblW_4_4_3, dblW_4_4_4, dblW_4_4_5, dblW_4_4_6, dblW_4_6, iRow, i
	
	' 개별손금비율 
	
	With lgvspdData(TYPE_4)
		For i = 1 To 4
			Select Case i
				Case 1
					iRow = 2	' -- 탭4그리드의 행위치 
				Case 2
					iRow = 4
				Case 3
					iRow = 11
				Case 4
					iRow = 18
			End Select
			.Row = iRow	' -- 해당위치 
			.Col = C_4_W3	: dblW_4_3		= dblW_4_3	 + UNICDbl(.value)
			.Col = C_4_W4_1	: dblW_4_4_1	= dblW_4_4_1 + UNICDbl(.value)
			.Col = C_4_W4_2	: dblW_4_4_2	= dblW_4_4_2 + UNICDbl(.value)
			.Col = C_4_W4_3	: dblW_4_4_3	= dblW_4_4_3 + UNICDbl(.value)
			.Col = C_4_W4_4	: dblW_4_4_4	= dblW_4_4_4 + UNICDbl(.value)
			.Col = C_4_W4_5	: dblW_4_4_5	= dblW_4_4_5 + UNICDbl(.value)
			.Col = C_4_W4_6	: dblW_4_4_6	= dblW_4_4_6 + UNICDbl(.value)
			.Col = C_4_W6	: dblW_4_6		= dblW_4_6   + UNICDbl(.value)
		Next
	End With
	
	If dblW_4_3 = 0 Then Exit Function
	
	With lgvspdData(TYPE_3)
		.Row = 2
		.Col = C_3_W1	: .value = dblW_4_3
		.Col = C_3_W3_1	: .value = dblW_4_4_1
		.Col = C_3_W5_1	: .value = dblW_4_4_2
		.Col = C_3_W7_1	: .value = dblW_4_4_3
		.Col = C_3_W3_2	: .value = dblW_4_4_4
		.Col = C_3_W5_2	: .value = dblW_4_4_5
		.Col = C_3_W7_2	: .value = dblW_4_4_6	
		.Col = C_3_W9	: .value = dblW_4_6
		
		.Col = C_3_W4_1	: .value = dblW_4_4_1 / dblW_4_3
		.Col = C_3_W6_1	: .value = dblW_4_4_2 / dblW_4_3
		.Col = C_3_W8_1	: .value = dblW_4_4_3 / dblW_4_3
		.Col = C_3_W4_2	: .value = dblW_4_4_4 / dblW_4_3
		.Col = C_3_W6_2	: .value = dblW_4_4_5 / dblW_4_3
		.Col = C_3_W8_2	: .value = dblW_4_4_6 / dblW_4_3	
		.Col = C_3_W10	: .value = dblW_4_6   / dblW_4_3
							
	End With
	 
End Function

'============================================  조회조건 함수  ====================================
Function OpenAdItem(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "조정과목 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_ADJUST_ITEM"					<%' TABLE 명칭 %>
	
	lgvspdData(lgCurrGrid).Col = C_2_W1
	lgvspdData(lgCurrGrid).Row = lgvspdData(lgCurrGrid).ActiveRow
	arrParam(2) = lgvspdData(lgCurrGrid).value		<%' Code Condition%>
	
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = " USE_YN = '1' "							<%' Where Condition%>
	arrParam(5) = "조정과목"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "ITEM_CD"					<%' Field명(0)%>
    arrField(1) = "ITEM_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "조정과목"					<%' Header명(0)%>
    arrHeader(1) = "과목명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAdItem(arrRet,iWhere)
		
	End If	
	
End Function

Function SetAdItem(byval arrRet,Byval iWhere)
    With lgvspdData(iWhere)
		.Col = C_2_W1		: .Value = arrRet(0)
		.Col = C_2_W1_NM	: .Value = arrRet(1)
		
		ggoSpread.Source = lgvspdData(iWhere)
		ggoSpread.UpdateRow .ActiveRow
		lgBlnFlgChgValue = True
	End With
	
End Function

'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
 
	Call InitComboBox	' 먼저해야 한다. 기업의 회계기준일을 읽어오기 위해 
	Call InitData

	Call FncQuery
     
    
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	Call GetFISC_DATE
End Sub

' -- 탭3.공통비용의 배부율 콤보 변경시 
Sub cboW11_onChange()
	Dim pVal
	pVal = frm1.cboW11.value 
	
	With lgvspdData(TYPE_3)
		.Col = C_3_W11
		.Row = 1	: .value = pVal
		 ggoSpread.UpdateRow .Row
		.Row = 2	: .value = pVal
		 ggoSpread.UpdateRow .Row
	End With
End Sub

'============================================  그리드 이벤트   ====================================
' -- 0번 그리드 
Sub vspdData0_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_1
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData0_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData0_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_1
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_GotFocus()
	lgCurrGrid = TYPE_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData0_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_1
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData0_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_1
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData0_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_1
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData0_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_1
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 1번 그리드 
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2_1
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2_1
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2_1
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_2_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2_1
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2_1
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2_1
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_2_1
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 2번 그리드 
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2_2
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2_2
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_GotFocus()
	lgCurrGrid = TYPE_2_2
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2_2
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2_2
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_2_2
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 3번 그리드 
Sub vspdData3_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_3
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_3
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_3
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_GotFocus()
	lgCurrGrid = TYPE_3
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_3
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_3
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData3_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_3
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 4번 그리드 
Sub vspdData4_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_4
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_4
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData4_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_4
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_GotFocus()
	lgCurrGrid = TYPE_4
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData4_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_4
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData4_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_4
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData4_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_4
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData4_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_4
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)
	Dim iIdx, iRow, sW3, sW4, dblW2

	With lgvspdData(Index)
		Select Case Col
			Case C_2_W3_CD
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
			Case C_2_W3
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
				
				' 매출액/매출원가 일때, 공통분 선택시 에러출력 
				.Col = C_2_W3_CD	: sW3 = UNICDbl(.text)
				.Col = C_2_W4		: sW4 = .Text
				If (sW3 = 1 Or sW3 = 2) And sW4 = "공통분" Then
					Call DisplayMsgBox("W60004", parent.VB_INFORMATION, "", "X")	
					.Col = C_2_W3	: .Text = ""
				End If
			Case C_2_W4
				' 매출액/매출원가 일때, 공통분 선택시 에러출력 
				.Col = C_2_W3_CD	: sW3 = UNICDbl(.text)
				.Col = C_2_W4		: sW4 = .Text
				If (sW3 = 1 Or sW3 = 2) And sW4 = "공통분" Then
					Call DisplayMsgBox("W60004", parent.VB_INFORMATION, "", "X")	
					.Col = C_2_W4	: .Text = ""
				End If
		End Select
	End With
End Sub

' -- 그리드2에 의해 그리드4을 변경한다 
Function SetGrid4ByGrid2()
	Dim iRow, iCol, dblW2, sW4, iIdx, i
	Dim iMaxRows, iType
	
	' -- 그리드4 초기화 
	With lgvspdData(TYPE_4)
		.BlockMode = True
		.Col = C_4_W3	: .Row = 1
		.Col = C_4_W7	: .Row = 20
		.Text = ""
		.BlockMode = False
	End With

	' -- 그리드1을 4로 옮긴다.
	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows
			
		' Grid1 --> Grid4 로 복제 
		For iRow = 1 To 20
			Call PutGrid4(C_4_W3, iRow, GetGrid(TYPE_1, C_1_W3, iRow))
			Call PutGrid4(C_4_W4_1, iRow, GetGrid(TYPE_1, C_1_W4_1, iRow))
			Call PutGrid4(C_4_W4_2, iRow, GetGrid(TYPE_1, C_1_W4_2, iRow))
			Call PutGrid4(C_4_W4_3, iRow, GetGrid(TYPE_1, C_1_W4_3, iRow))
			Call PutGrid4(C_4_W4_4, iRow, GetGrid(TYPE_1, C_1_W4_4, iRow))
			Call PutGrid4(C_4_W4_5, iRow, GetGrid(TYPE_1, C_1_W4_5, iRow))
			Call PutGrid4(C_4_W4_6, iRow, GetGrid(TYPE_1, C_1_W4_6, iRow))
			Call PutGrid4(C_4_W6, iRow, GetGrid(TYPE_1, C_1_W5, iRow))
		Next
	End With
			
	' -- 탭2 데이타가 있을 경우 
	For iType = TYPE_2_1 To TYPE_2_2
	
		With lgvspdData(iType)
			iMaxRows = .MaxRows
			
			For i = 1 To iMaxRows
				.Row = i
				
				.Col = C_2_W3_CD	: iRow = UNICDbl(.Text)	' 마이너코드값이 행위치이다.
				.Col = C_2_W4		: iIdx = UNICDbl(.Value)

				.Col = C_2_W2		: dblW2	 = UNICDbl(.value) ' 금액 
				.Col = C_2_W4		: sW4	 = .Text
				
				If sW4 <> "" Then
					Call SetGrid4Amt(iType, dblW2, iRow, sW4, iIdx)
				End If
			Next
	
		End With
	Next

End Function

' -- 탭2 그리드에서 콤보 선택시 해당 금액과 탭1 그리드의 값을 읽어 탭4 그리드에 반영한다 
Function SetGrid4Amt(Byval pType, Byval dblW2, Byval iRow, Byval sW4, Byval iIdx)

	Select Case iRow
		Case 1	' -- 매출액 
			If pType = TYPE_2_1 Then
				Call SetGrid4ByPlusAmt(pType, dblW2, iRow, sW4, iIdx)
			Else
				Call SetGrid4ByMinusAmt(pType, dblW2, iRow, sW4, iIdx)
			End If
		Case 2	' -- 매출원가 
			If pType = TYPE_2_2 Then
				Call SetGrid4ByPlusAmt(pType, dblW2, iRow, sW4, iIdx)
			Else
				Call SetGrid4ByMinusAmt(pType, dblW2, iRow, sW4, iIdx)
			End If
		Case 4	' -- 판관비 
			If pType = TYPE_2_2 Then
				Call SetGrid4ByPlusAmt(pType, dblW2, iRow, sW4, iIdx)
			Else
				Call SetGrid4ByMinusAmt(pType, dblW2, iRow, sW4, iIdx)
			End If
		Case 8	' -- 영업외수익 
			If pType = TYPE_2_1 Then
				Call SetGrid4ByPlusAmt(pType, dblW2, iRow, sW4, iIdx)
			Else
				Call SetGrid4ByMinusAmt(pType, dblW2, iRow, sW4, iIdx)
			End If
		Case 11	' -- 영업외비용 
			If pType = TYPE_2_2 Then
				Call SetGrid4ByPlusAmt(pType, dblW2, iRow, sW4, iIdx)
			Else
				Call SetGrid4ByMinusAmt(pType, dblW2, iRow, sW4, iIdx)
			End If
		Case 15	' -- 특별이익 
			If pType = TYPE_2_1 Then
				Call SetGrid4ByPlusAmt(pType, dblW2, iRow, sW4, iIdx)
			Else
				Call SetGrid4ByMinusAmt(pType, dblW2, iRow, sW4, iIdx)
			End If
		Case 18	' -- 특별손실 
			If pType = TYPE_2_2 Then
				Call SetGrid4ByPlusAmt(pType, dblW2, iRow, sW4, iIdx)
			Else
				Call SetGrid4ByMinusAmt(pType, dblW2, iRow, sW4, iIdx)
			End If
			
	End Select
End Function

' -- 탭2 그리드에서 콤보 선택시 해당 금액을 탭4 그리드에 반영한다 
Function SetGrid4ByPlusAmt(Byval pType, Byval pAmt, Byval iRow, Byval pW4, Byval iIdx)
	Dim dblAmt, dblSum
	With lgvspdData(TYPE_4)
		.Row = iRow
		Select Case pW4
			Case "기타"
				' 행.3열 와 행.5열 에 가산 
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W3, iRow)) : Call PutGrid2(TYPE_4, C_4_W3, iRow, dblAmt + pAmt)
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W6, iRow)) : Call PutGrid2(TYPE_4, C_4_W6, iRow, dblAmt + pAmt)
			Case "공통분"
				If iRow = 1 Or iRow = 2 Then	Exit Function
				
				iRow = iRow + 1
				' 행.3열 와 행.5열 에 가산 
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W3, iRow)) : Call PutGrid2(TYPE_4, C_4_W3, iRow, dblAmt + pAmt)
			Case Else
				' 감면사업인 경우 
				' 행.3열 와 행.4_1+콤보인덱스 열에 가산 
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W3, iRow)) : Call PutGrid2(TYPE_4, C_4_W3, iRow, dblAmt + pAmt)
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W4_1+(2*iIdx), iRow)) : Call PutGrid2(TYPE_4, C_4_W4_1+(2*iIdx), iRow, dblAmt + pAmt)
		End Select
	End With
End Function

' -- 탭2 그리드에서 콤보 선택시 해당 금액을 탭4 그리드에 반영한다 
Function SetGrid4ByMinusAmt(Byval pType, Byval pAmt, Byval iRow, Byval pW4, Byval iIdx)
	Dim dblAmt, dblSum
	With lgvspdData(TYPE_4)
		.Row = iRow
		Select Case pW4
			Case "기타"
				' 행.3열 와 행.5열 에 가산 
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W3, iRow)) : Call PutGrid2(TYPE_4, C_4_W3, iRow, dblAmt - pAmt)
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W6, iRow)) : Call PutGrid2(TYPE_4, C_4_W6, iRow, dblAmt - pAmt)
			Case "공통분"
				If iRow = 1 Or iRow = 2 Then	Exit Function
				
				iRow = iRow + 1
				' 행.3열 와 행.5열 에 가산 
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W3, iRow)) : Call PutGrid2(TYPE_4, C_4_W3, iRow, dblAmt - pAmt)
			Case Else
				' 감면사업인 경우 
				' 행.3열 와 행.4_1+콤보인덱스 열에 가산 
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W3, iRow)) : Call PutGrid2(TYPE_4, C_4_W3, iRow, dblAmt - pAmt)
				dblAmt = UNICDbl(GetGrid(TYPE_4, C_4_W4_1+(2*iIdx), iRow)) : Call PutGrid2(TYPE_4, C_4_W4_1+(2*iIdx), iRow, dblAmt - pAmt)
		End Select
	End With
End Function

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum
	
	lgBlnFlgChgValue= True ' 변경여부 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(Index).text) < UNICDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row

	' --- 추가된 부분 
	With lgvspdData(Index)

	Dim dblAmt(30)
	Dim dblW3, dblW4_1, dblW4_2, dblW4_3, dblW4_4, dblW4_5, dblW4_6, dblW5, dblW2, iRet, sW1
	Dim dblW3_1, dblW3_2, dblW3_3, dblW3_4, dblW3_5, dblW3_6
	Dim dblW5_1, dblW5_2, dblW5_3, dblW5_4, dblW5_5, dblW5_6, dblW6, dblW7
		
	If Index = TYPE_1 Then	'1번 그리 

		Select Case Col
			Case C_1_W3, C_1_W4_1,  C_1_W4_2,  C_1_W4_3,  C_1_W4_4,  C_1_W4_5,  C_1_W4_6, C_1_W5
				' 음수 체크해 절대값으로 치환해 넣는다. 0으로?
				If Row <> 3 And Row <> 7 And Row <> 14 And Row <> 21 And Row <> 25 Then
					.Col = Col	: .Row = Row
					If UNICDbl(.value) < 0 Then
						Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "금액", "X")	 
						.value = ABS(UNICDbl(.value))
					End If
				End If
				
				If Row = 1 Or Row = 2 Then
				
					.Row = Row
					.Col = C_1_W3	: dblW3	  = UNICDbl(.value)
					.Col = C_1_W4_1	: dblW4_1 = UNICDbl(.value)
					.Col = C_1_W4_2	: dblW4_2 = UNICDbl(.value)
					.Col = C_1_W4_3	: dblW4_3 = UNICDbl(.value)
					.Col = C_1_W4_4	: dblW4_4 = UNICDbl(.value)
					.Col = C_1_W4_5	: dblW4_5 = UNICDbl(.value)
					.Col = C_1_W4_6	: dblW4_6 = UNICDbl(.value)
				
					dblW5 = dblW3 - (dblW4_1 + dblW4_2 + dblW4_3 + dblW4_4 + dblW4_5 + dblW4_6)
					.Col = C_1_W5	: .value = dblW5
					
				ElseIf Row = 4 Or Row = 8 Or Row = 11 Or Row = 15 Or Row = 18 Then
				
					.Row = Row
					.Col = C_1_W4_1	: dblW4_1 = UNICDbl(.value)
					.Col = C_1_W4_2	: dblW4_2 = UNICDbl(.value)
					.Col = C_1_W4_3	: dblW4_3 = UNICDbl(.value)
					.Col = C_1_W4_4	: dblW4_4 = UNICDbl(.value)
					.Col = C_1_W4_5	: dblW4_5 = UNICDbl(.value)
					.Col = C_1_W4_6	: dblW4_6 = UNICDbl(.value)
					.Col = C_1_W5	: dblW5   = UNICDbl(.value)
					dblW3_4 = dblW4_1 + dblW4_2 + dblW4_3 + dblW4_4 + dblW4_5 + dblW4_6 + dblW5
					.Col = C_1_W3	: .Row = Row	: .value = dblW3_4
					
					.Row = Row + 2	: .Col = C_1_W3		: dblW3_6 = UNICDbl(.value)
					dblW3_5 = dblW3_6 - dblW3_4
					.Row = Row + 1	: .Col = C_1_W3		: .value = dblW3_5
					
					
					'Call ReCalcGrid(TYPE_1)
				End If
				
			'Case C_1_W3
					
				Call ReCalcGrid(TYPE_1)

		End Select
	
	ElseIf Index = TYPE_2_1 Or Index = TYPE_2_2 Then
		' -- 익금산입/손금불산입, 손금산입/입금불산입 
		Select Case Col
			Case C_2_W2
							
				' -- 썸을 출력한다.
				If Chk1TotalLine(Index) Then
					' -- 1개의 토탈라인인 경우 
					Call FncSumSheet(lgvspdData(Index), C_2_W2, 1, .MaxRows - 1, true, .MaxRows, C_2_W2, "V")
				Else
					' -- 1개의 토탈라인인 경우 
					dblSum = FncSumSheet(lgvspdData(Index), C_2_W2, 1, .MaxRows - 3, true, .MaxRows - 2, C_2_W2, "V")
					
					' -- 기부금한도초과/기부금전기이월 값 
					dblW2 = UNICDbl(GetGrid(Index, C_2_W2, .MaxRows-1))
					dblSum = dblSum + dblW2
					
					Call PutGrid(Index, C_2_W2, .MaxRows, dblSum)
				End If
				
			Case C_2_W1
				' -- 코드입력시 코드명 가져옴 
				.Col = Col	: .Row = Row	: sW1 = .Text
				iRet = CommonQueryRs("ITEM_NM", "TB_ADJUST_ITEM", " ITEM_CD='" & sW1 & "' AND USE_YN='1'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

				If iRet <> False Then
					.Col = C_2_W1_NM	: .Row = Row : .Text = Replace(lgF0, Chr(11), "")
				Else
					' --WC0023
					Call DisplayMsgBox("WC0023", parent.VB_INFORMATION, "제15호 조정과목 코드", "X")	
					.Row = Row
					.Col = C_2_W1		: .Text = ""
					.Col = C_2_W1_NM	: .Text = ""
				End If
				
		End Select
		
	ElseIf Index = TYPE_4 Then
		
		Select Case Col
			Case C_4_W3, C_4_W4_1, C_4_W4_2, C_4_W4_3, C_4_W4_4, C_4_W4_5, C_4_W4_6
			
			.Row = Row
			.Col = C_4_W3	: dblAmt(C_4_W3) = UNICDbl(.value)
			.Col = C_4_W4_1	: dblAmt(C_4_W4_1) = UNICDbl(.value)
			.Col = C_4_W4_2	: dblAmt(C_4_W4_2) = UNICDbl(.value)
			.Col = C_4_W4_3	: dblAmt(C_4_W4_3) = UNICDbl(.value)
			.Col = C_4_W4_4	: dblAmt(C_4_W4_4) = UNICDbl(.value)
			.Col = C_4_W4_5	: dblAmt(C_4_W4_5) = UNICDbl(.value)
			.Col = C_4_W4_6	: dblAmt(C_4_W4_6) = UNICDbl(.value)
			
			.Col = C_4_W5_1	: .text	= ""
			.Col = C_4_W5_2	: .text	= ""
			.Col = C_4_W5_3	: .text	= ""
			.Col = C_4_W5_4	: .text	= ""
			.Col = C_4_W5_5	: .text	= ""
			.Col = C_4_W5_6	: .text	= ""
			.Col = C_4_W7	: .text	= ""
			
			.Col = C_4_W6	: .value = dblAmt(C_4_W3) - dblAmt(C_4_W4_1) - dblAmt(C_4_W4_2) - dblAmt(C_4_W4_3) - dblAmt(C_4_W4_4) - dblAmt(C_4_W4_5) - dblAmt(C_4_W4_6)

			' 25행 결과 계산 
			Call PutGrid4(C_4_W3, 25, UNICDbl(GetGrid(TYPE_4, C_4_W3, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W3, 24)) )
			Call PutGrid4(C_4_W4_1, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_1, 24)) )
			Call PutGrid4(C_4_W4_2, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_2, 24)) )
			Call PutGrid4(C_4_W4_3, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_3, 24)) )
			Call PutGrid4(C_4_W4_4, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_4, 24)) )
			Call PutGrid4(C_4_W4_5, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_5, 24)) )
			Call PutGrid4(C_4_W4_6, 25, UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W4_6, 24)) )
			Call PutGrid4(C_4_W6, 25, UNICDbl(GetGrid(TYPE_4, C_4_W6, 21)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 22)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 23)) - UNICDbl(GetGrid(TYPE_4, C_4_W6, 24)) )

		End Select
	End If
	
	End With
	
	Call ChangeAllUpdateFlg(Index)
	If Index < TYPE_4 Then
		Call ChangeAllUpdateFlg(TYPE_4)
	End If
End Sub

' -- 구분1,2 존재 체크 
Function ChkW3_W4(Byval pType)
	Dim sW3, sW4
	With lgvspdData(pType)
		.Row = .ActiveRow
		.Col = C_2_W3	: sW3 = .Text
		.Col = C_2_W4	: sW4 = .Text
		If sW3 <> "" And sW4 <> "" Then		' 둘다 존재할 경우 
			ChkW3_W4 = True
		Else
			ChkW3_W4 = False
		End If
	End With
End Function

' -- 구분1,2 존재 체크 
Function ChkOLDW2_W3_W4(Byval pType)
	Dim sW3, sW4, sW2
	With lgvspdData(pType)
		.Row = .ActiveRow
		.Col = C_2_W2_OLD	: sW2 = .Text
		.Col = C_2_W3_OLD	: sW3 = .Text
		.Col = C_2_W4_OLD	: sW4 = .Text
		If sW2 <> "" And sW3 <> "" And sW4 <> "" Then		' 둘다 존재할 경우 
			ChkOLDW2_W3_W4 = True
		Else
			ChkOLDW2_W3_W4 = False
		End If
	End With
End Function

' -- 현재 그리드의 합계 체크 
Function Chk1TotalLine(Byval pType)

	With lgvspdData(pType)
		.Row = .MaxRows	: .Col = C_2_SEQ_NO
		If UNICDbl(.value) = SUM_SEQ_NO Then
			' -- 합계 1개인 경우 
			Chk1TotalLine = True
		Else
			' -- 합계 2개인 경우 
			Chk1TotalLine = False
		End If
	End With
End Function

' -- 그리드 1을 재계산 
Function ReCalcGrid(pType)
	With lgvspdData(pType)
	
	Select Case pType
		Case TYPE_1
			Dim dblW3_1, dblW3_2
			Dim dblW3_3, dblW3_6, dblW3_7, dblW3_10, dblW3_13, dblW3_14, dblW3_17, dblW3_20, dblW3_21, dblW3_22, dblW3_23, dblW3_24, dblW3_25
			Dim dblW3_4, dblW3_5, dblW3_8, dblW3_9, dblW3_11, dblW3_12, dblW3_15, dblW3_16, dblW3_18, dblW3_19
			
			.Col = C_1_W3
			
			.Row = 1	: dblW3_1	= UNICDbl(.value)
			.Row = 2	: dblW3_2	= UNICDbl(.value)
			dblW3_3		= dblW3_1 - dblW3_2	' 매출이익 
			.Row = 3	: .value	= dblW3_3	
			
			.Row = 4	: dblW3_4	= UNICDbl(.value)
			.Row = 6	: dblW3_6	= UNICDbl(.value)
			
			dblW3_5		= dblW3_6 - dblW3_4	' 판매비 공통(코드05)
			dblW3_7		= dblW3_3 - dblW3_6	' 영업이익(코드07)
			
			.Row = 5	: .value	= dblW3_5	
			.Row = 7	: .Value	= dblW3_7
			
			.Row = 8	: dblW3_8	= UNICDbl(.value)
			.Row = 10	: dblW3_10	= UNICDbl(.value)
			
			dblW3_9		= dblW3_10 - dblW3_8	' 영업외수익(코드09)
			.Row = 9	: .value = dblW3_9
			
			.Row = 11	: dblW3_11	= UNICDbl(.value)
			.Row = 13	: dblW3_13	= UNICDbl(.value)
			
			dblW3_12	= dblW3_13 - dblW3_11	' 영업외비용(코드12)
			.Row = 12	: .Value	= dblW3_12
			
			dblW3_14	= dblW3_7 + dblW3_10 - dblW3_13 ' 경상이익		
			.Row = 14	: .Value	= dblW3_14
			
			.Row = 15	: dblW3_15	= UNICDbl(.value)
			.Row = 17	: dblW3_17	= UNICDbl(.value)
			
			dblW3_16	= dblW3_17 - dblW3_15	' 영업외비용(코드12)
			.Row = 16	: .Value	= dblW3_16
			
			.Row = 18	: dblW3_18	= UNICDbl(.value)
			.Row = 20	: dblW3_20	= UNICDbl(.value)
			
			dblW3_19	= dblW3_20 - dblW3_18	' 영업외비용(코드12)
			.Row = 19	: .Value	= dblW3_19
			
			.Row = 21	: dblW3_21	= dblW3_14 + dblW3_17 - dblW3_20	: .value = dblW3_21
			.Row = 22	: dblW3_22	= UNICDbl(.value)
			.Row = 23	: dblW3_23	= UNICDbl(.value)
			.Row = 24	: dblW3_24	= UNICDbl(.value)
			dblW3_25	= dblW3_21 - dblW3_22 - dblW3_23 - dblW3_24
			.Row = 25	: .value = dblW3_25
			
	End Select
	
	End With
End Function

' 2번 그리드 썸계산 
Function SetGridSum2()
	Dim dblW10, dblW11, dblW12, dblW13, dblW14, iRow
	Dim dblW10Sum, dblW11Sum, dblW12Sum, dblW13Sum
	
	With lgvspdData(TYPE_2)
		.Row = 1
		.Col = C_W10	: dblW10 = UNICDbl(.value)	: dblW10Sum = dblW10
		.Col = C_W11	: dblW11 = UNICDbl(.value)	: dblW11Sum = dblW11
		
		dblW12 = dblW10 - dblW11
		.Col = C_W12	: .Value = dblW12

		.Row = 2
		.Col = C_W10	: dblW10 = UNICDbl(.value)	: dblW10Sum = dblW10Sum + dblW10
		.Col = C_W11	: dblW11 = UNICDbl(.value)	: dblW11Sum = dblW11Sum + dblW11

		dblW13 = dblW11 - dblW10
		.Col = C_W13	: .Value = dblW13

		.Row = 3
		.Col = C_W10	: dblW10 = UNICDbl(.value)	: dblW10Sum = dblW10Sum + dblW10
		.Col = C_W11	: dblW11 = UNICDbl(.value)	: dblW11Sum = dblW11Sum + dblW11
		
		dblW14 = dblW11 - dblW10
		.Col = C_W14	: .Value = dblW14
		
		.Row = 4
		.Col = C_W10	: .value = dblW10Sum
		.Col = C_W11	: .value = dblW11Sum
		.Col = C_W12	: .value = dblW12
		
		.Row = 5
		.Col = C_W13	: .value = dblW13

	End With
End Function

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(Index)
   
    If lgvspdData(Index).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(Index)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	lgvspdData(Index).Row = Row
End Sub

Sub vspdData_ColWidthChange(Index, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(Index, ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If lgvspdData(Index).MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus(Index)
    ggoSpread.Source = lgvspdData(Index)
    lgCurrGrid = Index
End Sub

Sub vspdData_MouseDown(Index, Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	lgCurrGrid = Index
	ggoSpread.Source = lgvspdData(Index)
End Sub    

Sub vspdData_ScriptDragDropBlock(Index, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(Index, ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if lgvspdData(Index).MaxRows < NewTop + VisibleRowCnt(lgvspdData(Index),NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(Index, ByVal Col, ByVal Row, Byval ButtonDown)
	With lgvspdData(Index)
		If Row > 0 And Col = C_2_W1_BT Then
		    .Row = Row
		    .Col = C_2_W1_BT

		    Call OpenAdItem(Index)
		End If
    End With
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'============================================  툴바지원 함수  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                               <%'Protect system from crashing%>

	Call ClickTab1()
<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	'For i = TYPE_1 To TYPE_6
	'	ggoSpread.Source = lgvspdData(i)
	'	If ggoSpread.SSCheckChange = True Then
	'		blnChange = True
	'		Exit For
	'	End If
    'Next
    
    If lgBlnFlgChgValue Or blnChange Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables													<%'Initializes local global variables%>
   ' Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i, sMsg
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

    'If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

' ---------------------- 서식내 검증 -------------------------
Function  Verification()
	Dim dblW11, dblW12, dblW16, dblW14, dblW15, dblW13
	
	Verification = False

	Verification = True	
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100000000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If lgvspdData(lgCurrGrid).MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If lgvspdData(lgCurrGrid).ActiveRow > 0 Then
			lgvspdData(lgCurrGrid).focus
			lgvspdData(lgCurrGrid).ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor lgCurrGrid, lgvspdData(lgCurrGrid).ActiveRow, lgvspdData(lgCurrGrid).ActiveRow

			lgvspdData(lgCurrGrid).Col = C_W13
			lgvspdData(lgCurrGrid).Text = ""
    
			lgvspdData(lgCurrGrid).Col = C_W3
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).Col = C_W4
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).Col = C_W5
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
	FncCancel = False
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    
    If lgvspdData(lgCurrGrid).ActiveRow = lgvspdData(lgCurrGrid).MaxRows Then 
		Exit Function
	End If
    ggoSpread.EditUndo  
    If lgvspdData(lgCurrGrid).MaxRows = 1 Then
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		ggoSpread.ClearSpreadData
    End If
    FncCancel = True
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

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
   
	With lgvspdData(lgCurrGrid)

		.focus
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		
		iRow = .ActiveRow
		
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow ,1
			Call SetSpreadColor(lgCurrGrid, iRow, iRow) 
			.Col = C_2_SEQ_NO : .Row = iRow	: .Text = 1
			
			iRow = 2
			ggoSpread.InsertRow ,1
			Call SetSpreadColor(lgCurrGrid, iRow, iRow) 
			.Col = C_2_SEQ_NO : .Row = iRow	: .Text = SUM_SEQ_NO	
			Call SetReDrawTotalLine(lgCurrGrid)
										
		Else
			
			If iRow = .MaxRows Then
				ggoSpread.InsertRow iRow-1 , imRow 
				Call SetSpreadColor(lgCurrGrid,iRow, iRow + imRow - 1)
				Call MaxSpreadVal(lgvspdData(lgCurrGrid), C_2_SEQ_NO, iRow)
				'.vspdData.Col = C_SEQ_NO : .vspdData.Row = Row : iSeqNo = .vspdData.Value
			Else
				ggoSpread.InsertRow ,imRow
				Call SetSpreadColor(lgCurrGrid, iRow+1, iRow+1)
				Call MaxSpreadVal(lgvspdData(lgCurrGrid), C_2_SEQ_NO, iRow+1)
				'.vspdData.Col = C_SEQ_NO : .vspdData.Row = Row+1 : iSeqNo = .vspdData.Value
			End If   
		End If
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
    lgBlnFlgChgValue = False
    
End Function

Sub SetReDrawTotalLine(Byval pType)
	Dim ret
	
	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)
		.Row = .MaxRows
		.Col = C_2_SEQ_NO
		
		If UNICDbl(.value) = SUM_SEQ_NO Then
			' -- 과목코드/팝업/코드명위치 
			.BlockMode = True
			.Col = C_2_W1		: .Row = .MaxRows
			.Col2 = C_2_W1_NM	: .Row2 = .MaxRows
			.CellType = 1
			.BlockMode = False
			
			' -- 구분1,2위치 
			.BlockMode = True
			.Col = C_2_W3		: .Row = .MaxRows
			.Col2 = C_2_W4		: .Row2 = .MaxRows
			.CellType = 1
			.BlockMode = False
			
			.Col = C_2_W1	: .Row = .MaxRows : .Text = "합계" : .TypeHAlign = 2
			ret = .AddCellSpan(C_2_W1, .MaxRows, 3, 1)
			ggoSpread.SpreadLock C_2_W1, .MaxRows, C_2_W4
		Else
			' -- 과목코드/팝업/코드명위치 
			.BlockMode = True
			.Col = C_2_W1		: .Row = .MaxRows -2
			.Col2 = C_2_W1_NM	: .Row2 = .MaxRows -2
			.CellType = 1
			.BlockMode = False

			' -- 구분1,2위치 
			.BlockMode = True
			.Col = C_2_W3		: .Row = .MaxRows
			.Col2 = C_2_W4		: .Row2 = .MaxRows
			.CellType = 1
			.BlockMode = False
			
			.Col = C_2_W1	: .Row = .MaxRows -2 : .Text = "합계" : .TypeHAlign = 2
			ret = .AddCellSpan(C_2_W1, .MaxRows -2, 3, 1)
			ggoSpread.SpreadLock C_2_W1, .MaxRows-2, C_2_W4
			
			' -- 과목코드/팝업/코드명위치 
			.BlockMode = True
			.Col = C_2_W1		: .Row = .MaxRows
			.Col2 = C_2_W1_NM	: .Row2 = .MaxRows
			.CellType = 1
			.BlockMode = False

			' -- 구분1,2위치 
			.BlockMode = True
			.Col = C_2_W3		: .Row = .MaxRows
			.Col2 = C_2_W4		: .Row2 = .MaxRows
			.CellType = 1
			.BlockMode = False
			
			.Col = C_2_W1	: .Row = .MaxRows : .Text = "합계" : .TypeHAlign = 2
			ret = .AddCellSpan(C_2_W1, .MaxRows, 3, 1)
			ggoSpread.SpreadLock C_2_W1, .MaxRows, C_2_W4
		End If
	End With
End Sub

Function FncDeleteRow() 
	Dim iMaxRows, iRow, iAllDel, lDelRows, iSeqNo
	iAllDel = True
	
	FncDeleteRow = False
	If lgCurrGrid <> TYPE_2_1 And lgCurrGrid <> TYPE_2_2 Then Exit Function
	
	With lgvspdData(lgCurrGrid)	
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    lDelRows = ggoSpread.DeleteRow
    
    If .MaxRows > 1 Then
		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_2_SEQ_NO : iSeqNo = UNICDbl(.Value)
			.Col = 0 
			If .Text <> ggoSpread.DeleteFlag And iSeqNo <> 999999 Then  iAllDel = False
		Next
		
		If iAllDel Then
			lDelRows = ggoSpread.DeleteRow(.MaxRows)
		End If

	End If	
	End With
	lgBlnFlgChgValue = True
	FncDeleteRow = True
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function

'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        'strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function ReDrawGRidColHead()
	' -- 그리드 컬럼헤더를 재 갱신한다.
	Dim iRow, ret
	
	With lgvspdData(TYPE_1)
		.Redraw = False
		
		Call SetSpreadLock
		
		iRow = ReDrawW1("0", 1)
		iRow = ReDrawW1("1", iRow)

		.Row = iRow		
		.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2	
		.Col = C_W2_P	: .CellType = 1
		ret = .AddCellSpan(C_W1	, .Row, 3, 1)
		ggoSpread.SpreadLock C_W1, .Row, C_W10, .Row
				
		.Redraw = True
	End With
End Function

Function MakeCrLf(Byval iCnt)
	Dim i, sTmp
	If iCnt < 1 Then Exit Function
	For i = 1 to iCnt
		sTmp = sTmp & vbCrLf 
	Next
	MakeCrLf = sTmp
End Function

Function ReDrawW1(Byval pW1_CD, Byval pRow)
	Dim iRow, iMaxRows, iRowLoc , iRowSpanCnt, ret
	
	pRow = pRow 
	iRowLoc = pRow : iRowSpanCnt = 0

	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows
		.Row = pRow		: .Col = C_W1	

		Do Until False
			.Row = pRow	: .Col = C_W1_CD
			If Left(.Value, 1) = pW1_CD Then
				iRowSpanCnt = iRowSpanCnt + 1
			Else
				' -- 합계 
				.Row = pRow - 1
				.Col = C_W2		: .CellType = 1	: .Text = "합계"	: .TypeHAlign = 2	
				.Col = C_W2_P	: .CellType = 1
				ggoSpread.SpreadLock C_W1, .Row, C_W10, .Row
				ret = .AddCellSpan(C_W1	, iRowLOc, 1, iRowSpanCnt)
				Exit Do
			End If
			pRow = pRow + 1
		Loop
		
		ReDrawW1 = pRow 

		.Row = iRowLoc
		If pW1_CD = "0" Then
			.value = "자" & MakeCrLf(iRowSpanCnt/2) & "산"
		Else
			.value = "부" & MakeCrLf(iRowSpanCnt/2) & "채"
		End If
		If iRowSpanCnt > 1 Then
			.TypeEditMultiLine = True
		End If		
		.TypeHAlign = 2 : .TypeVAlign = 2
		
	End With
End Function
		
Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	If lgvspdData(TYPE_1).MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101000000000111")										<%'버튼 툴바 제어 %>

		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
		End If
	Else
		Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
	End If
	
	'Call SetSpreadTotalLine ' - 합계라인 재구성 
	
	'lgvspdData(lgCurrGrid).focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel, sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    For i = TYPE_1 To TYPE_4	' 전체 그리드 갯수 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
			' ----- 1번째 그리드 
			For lRow = 1 To .MaxRows
    
		       .Row = lRow
		       .Col = 0 : sTmp = Parent.gColSep

			  ' 모든 그리드 데이타 보냄     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = 1 To lMaxCols
						.Col = lCol : sTmp = sTmp & Trim(.Text) &  Parent.gColSep
					Next
					sTmp = sTmp & Trim(.Text) &  Parent.gRowSep
			  End If  

		       .Row = lRow	: .Col = 0
		    
		       Select Case .Text
		           Case  ggoSpread.InsertFlag                                      '☜: Insert
		                                              strVal = strVal & "C" & sTmp
		           Case  ggoSpread.UpdateFlag                                      '☜: Update
		                                              strVal = strVal & "U" & sTmp
		           Case  ggoSpread.DeleteFlag                                      '☜: Update
		                                              strDel = strDel & "D" & sTmp
		       End Select
		       
			Next
		
		End With

		document.all("txtSpread" & CStr(i)).value =  strDel & strVal
		strDel = "" : strVal = ""
	Next

	'Frm1.txtSpread.value      = strDel & strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	
    Call MainQuery()
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

Function ProgramJump
    Call PgmJump(JUMP_PGM_ID)
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
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE border=0 cellpadding=0 cellspacing=0 width=1024>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP" width=170>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>1. 손익계산서 소득구분</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" width=170>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>2. 세무조정 소득구분</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" width=170>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>3. 공통비용 배부율 등록</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" width=170>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab4()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>4. 소득구분계산서 작성</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><DIV id="myTabRef"><A href="vbscript:GetRef1">감면사업등록</A>|<A href="vbscript:GetRef2">손익계산서불러오기</A></DIV>
						<DIV id="myTabRef" STYLE="display:'none'"><A href="vbscript:GetRef3">금액불러오기</A></DIV>
						<DIV id="myTabRef" STYLE="display:'none'">&nbsp;</DIV>
						<DIV id="myTabRef" STYLE="display:'none'"><A href="vbscript:GetRef4">금액불러오기</A></DIV>
						</TD>
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
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
							     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							    </TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="50%" VALIGN=TOP HEIGHT=*>익금산입.손금불산입 
							     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=95% tag="23" TITLE="SPREAD" id=vspdData1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							    </TD>
							     <TD WIDTH="50%" VALIGN=TOP HEIGHT=*>손금산입.입금불산입 
							     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=95% tag="23" TITLE="SPREAD" id=vspdData2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							    </TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=100 COLSPAN=4>
							     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData3> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							    </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" HEIGHT=10>공통비용분의 배부율등록</TD>
								<TD CLASS="TD6"><SELECT NAME=cboW11 STYLE="WIDTH: 200" tag="23"><OPTION VALUE=""></SELECT></TD>
								<TD CLASS="TD5"></TD>
								<TD CLASS="TD6"></TD>
							</TR>
							<TR>
								<TD COLSPAN=4 HEIGHT=*>&nbsp;</TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
							     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData4 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData4> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							    </TD>
							</TR>
						</TABLE>
						</DIV>

					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR HEIGHT=20>   
        <TD WIDTH=100%>   
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                <TD WIDTH=50%>&nbsp;</TD>   
                <TD WIDTH=50%>   
                    <TABLE WIDTH=100%>                           
                        <TD WIDTH=* Align=right><DIV ID=myTabRef2>&nbsp;</A></DIV>
		          <DIV ID=myTabRef2 STYLE="display:'none'">&nbsp;</A></DIV>
		          <DIV ID=myTabRef2 STYLE="display:'none'">&nbsp;</A></DIV>
		          <DIV ID=myTabRef2 STYLE="display:'none'"><A href="Vbscript:ProgramJump()">제3호 법인세과세표준 및 세액조정계산서</A></DIV></TD>                                                                                     
                        <TD WIDTH=10>&nbsp;</TD>                           
                    </TABLE>   
                </TD>   
                </TR>   
            </TABLE>   
        </TD>   
    </TR> 
    <TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			    
		
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread4" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

