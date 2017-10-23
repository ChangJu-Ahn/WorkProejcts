
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 19호 가지급금등의인정이자조정명세서(갑)
'*  3. Program ID           : W3115MA1
'*  4. Program Name         : W3115MA1.asp
'*  5. Program Desc         : 19호 가지급금등의인정이자조정명세서(갑)
'*  6. Modified date(First) : 2005/01/24
'*  7. Modified date(Last)  : 2006/01/25
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : HJO
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
' 
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W3115MA1"
Const BIZ_PGM_ID		= "W3115MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W3115MB2.asp"
Const EBR_RPT_ID		= "W3115OA1"

' -- 1번 그리드 
Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_R1		' 3117의 SEQ_NO
Dim C_R2		' 3117의 (8)인정이자율 종류 
Dim C_R3		' 3117의 (5)차감계 (8)이 회사부담이자율인 경우만 
Dim C_R4		' 3117의 (5)의 구성비 
Dim C_R5		' 3117의 (6)이자수익에 대한 값 

' -- 2번 그리드 
Dim C_CHILD_SEQ_NO
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10

' -- 3번 그리드 
Dim C_W_TYPE
Dim C_SEQ_NO2
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16
Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W21
Dim C_W22
Dim C_W23

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgChgFlg
Dim lgFISC_START_DT, lgFISC_END_DT, lgRateConf
Dim lgblnYoon

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	lgCurrGrid	= 1
	lgChgFlg	= False

	'--1번그리드 
	C_SEQ_NO = 1
	C_W1 = 2
	C_W2 = 3
	C_W3 = 4
	C_W4 = 5
	C_W5 = 6
	C_R1 = 7
	C_R2 = 8
	C_R3 = 9
	C_R4 = 10
	C_R5 = 11

	'--2번그리드 
	C_CHILD_SEQ_NO	= 2
	C_W6		= 3
	C_W7		= 4
	C_W8		= 5
	C_W9		= 6
	C_W10		= 7

	'--3번그리드 
	C_W_TYPE = 1
	C_SEQ_NO2 = 2
	C_W11 = 3
	C_W12 = 4
	C_W13 = 5
	C_W14 = 6
	C_W15 = 7
	C_W16 = 8
	C_W17 = 9
	C_W18 = 10
	C_W19 = 11
	C_W20 = 12
	C_W21 = 13
	C_W22 = 14
	C_W23 = 15

	
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
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
     
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	' 1번 그리드 
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
	   'patch version
	    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	
	    .MaxCols = C_R5 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	    .MaxRows = 0
	    
		'헤더를 2줄로    
	    .ColHeaderRows = 2
	    'Call AppendNumberPlace("6","3","2")
	
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_W1,		"(1)성명",			10,,,50,1	
		ggoSpread.SSSetFloat	C_W2,		"(2)가지급금" & VbCrlf & "적수",	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W3,		"(3)가수금" & VbCrlf & "적수",	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W4,		"(4)금액{(2)-(3)}",	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetEdit		C_W5,		"(5)구성비",		6, 2,,10,2
	    ggoSpread.SSSetEdit		C_R1,		"Seq_no",		6, 2,,10,2
	    ggoSpread.SSSetEdit		C_R2,		"(8)인정이자율종류",		6, 2,,10,2
	    ggoSpread.SSSetFloat	C_R3,		"금액",	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetEdit		C_R4,		"구성비",		6, 2,,10,2
	    ggoSpread.SSSetEdit		C_R5,		"(6)이자수익",		6, 2,,10,2
	
		' 퍼센트 형 정의 
	    .Col = C_W5
	    .Row = -1
	    .CellType = 14
'	    .TypePercentDecimal = 1
	    .TypePercentMax = 100
	    .TypePercentMin = 0
	    .TypePercentDecPlaces = 2
	    
		' 퍼센트 형 정의 
	    .Col = C_R4
	    .Row = -1
	    .CellType = 14
'	    .TypePercentDecimal = 1
	    .TypePercentMax = 100
	    .TypePercentMin = 0
	    .TypePercentDecPlaces = 2
	    
		' 그리드 헤더 합침 정의 
		'ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)	' SEQ_NO 행 합침 
	    ret = .AddCellSpan(C_W1, -1000, 1, 2)	' 성명 
	    ret = .AddCellSpan(C_W2, -1000, 1, 2)	' 가지급금적수 
	    ret = .AddCellSpan(C_W3, -1000, 1, 2)	' 가수금적수 
	    ret = .AddCellSpan(C_W4, -1000, 2, 1)	' 차감적수 
	    ret = .AddCellSpan(C_R3, -1000, 2, 1)	' 회사부담 적용적수 
	    
	    ' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W4
		.Text = "차감적수"
		.Col = C_R3
		.Text = "회사부담 적용적수"
	
		' 두번째 헤더 출력 글자 
		.Row = -999
		.Col = C_W4
		.Text = "(4)금액" & VbCrlf & "{(2)-(3)}"
		.Col = C_W5
		.Text = "(5)구성비"
		.Col = C_R3
		.Text = "금액" & VbCrlf & "{(2)-(3)}"
		.Col = C_R4
		.Text = "구성비"
		.rowheight(-999) = 20	' 높이 재지정 
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_R1,C_R1,True)
		Call ggoSpread.SSSetColHidden(C_R2,C_R2,True)
		Call ggoSpread.SSSetColHidden(C_R5,C_R5,True)
		
		ggoSpread.SSSetSplit2(2) 						
		.ReDraw = true
	
    End With

 	' -----  2번 그리드 
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_W10 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	   .MaxRows = 0
	 
		'헤더를 2줄로    
	    .ColHeaderRows = 2
	    Call AppendNumberPlace("6","3","2")
	
		ggoSpread.SSSetEdit		C_SEQ_NO,	"부모순번", 5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"자식순번", 5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_W6,		"(6)이자율",		5, 2,,10,2
	    ggoSpread.SSSetFloat	C_W7,		"(7)적수" ,				12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W8,		"(8)인정이자{(7)X(6)}/365" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W9,		"(9)회사계상액",			12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetFloat	C_W10,		"(10)조정액" & VbCrLf & "{(8) - (9)}",			12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	
		' 퍼센트 형 정의 
	    .Col = C_W6
	    .Row = -1
	    .CellType = 14
'	    .TypePercentDecimal = 0
	    .TypePercentMax = 100
	    .TypePercentMin = 0
	    .TypePercentDecPlaces = 2
	    
		' 그리드 헤더 합침 정의 
		'ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)	' SEQ_NO 행 합침 
	    'ret = .AddCellSpan(C_CHILD_SEQ_NO, -1000, 1, 2)	' SEQ_NO 행 합침 
	    ret = .AddCellSpan(C_W6, -1000, 3, 1)	' 인정이자계산 
	    ret = .AddCellSpan(C_W9, -1000, 1, 2)	' 회사계상액 
	    ret = .AddCellSpan(C_W10, -1000, 1, 2)	' 조정개 
	    
	    ' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W6
		.Text = "인정이자계산"
	
		' 두번째 헤더 출력 글자 
		.Row = -999
		.Col = C_W6
		.Text = "(6)이자율"
		.Col = C_W7
		.Text = "(7)적수"
		.Col = C_W8
		.Text = "(8)인정이자" & VbCrLf & "{(7)X(6)}/365"
		.rowheight(-999) = 20	' 높이 재지정 
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_CHILD_SEQ_NO,C_CHILD_SEQ_NO,True)
					
		.ReDraw = true
	
    End With

 	' -----  3번 그리드 
	With frm1.vspdData3
	
		ggoSpread.Source = frm1.vspdData3	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_W23 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	   .MaxRows = 4
	 
	    Call AppendNumberPlace("6","3","2")
	
		ggoSpread.SSSetEdit		C_W_TYPE,	"구분", 5,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_SEQ_NO2,	"부모순번", 5,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_W11,		"성명",		15,,,20,	1
	    ggoSpread.SSSetEdit		C_W12,		"구성비",		6, 2,,10,2
	    ggoSpread.SSSetFloat	C_W13,		"계" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W14,		"(15)" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W15,		"(16)",		12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
		ggoSpread.SSSetFloat	C_W16,		"(17)" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W17,		"(18)" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W18,		"(19)" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W19,		"(20)",		12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
		ggoSpread.SSSetFloat	C_W20,		"(21)" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W21,		"(22)",		12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
		ggoSpread.SSSetFloat	C_W22,		"(23)" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W23,		"(24)" ,	12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
			
		' 퍼센트 형 정의 
	    .Col = C_W12
	    .Row = -1
	    .CellType = 14
'	    .TypePercentDecimal = 0
	    .TypePercentMax = 100
	    .TypePercentMin = 0
	    .TypePercentDecPlaces = 2

		' 이자율, 지급이자, 차입금적수, 이자율별 적용적수 설정 
		Call SetInitGrid3
		Call ChangeRowFlg(frm1.vspdData3)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO2,C_SEQ_NO2,True)
		ggoSpread.SSSetSplit2(5) 									
		.ReDraw = true
    
    End With
    
'	Call InitSpreadComboBox()
    Call SetSpreadLock 
           
End Sub


'============================================  그리드 함수  ====================================
Sub SetInitGrid3()
	DIm iCol
	With frm1.vspdData3
	
		ggoSpread.Source = frm1.vspdData3	
		' 퍼센트 형 정의 
		For iCol = C_W13 To .MaxCols - 1 
		    .Col = iCol	:	.Row = 1	:	.CellType = 14	:	.TypePercentMax = 100	:	.TypePercentMin = 0	:	.TypePercentDecPlaces = 2
		Next
		
		'타이틀 셋팅 
		.Col = C_W11	:	.Row = 1	:	.Text = "(11)이자율"
		.Col = C_W11	:	.Row = 2	:	.Text = "(12)지급이자"
		.Col = C_W11	:	.Row = 3	:	.Text = "(13)차입금적수"
		.Col = C_W11	:	.Row = 4	:	.Text = "(14)이자율별적용적수"
		
		' 타입 설정 
		.Col = C_W_TYPE	:	.Row = 1	:	.Text = "H"
		.Col = C_W_TYPE	:	.Row = 2	:	.Text = "H"
		.Col = C_W_TYPE	:	.Row = 3	:	.Text = "H"
		.Col = C_W_TYPE	:	.Row = 4	:	.Text = "H"
		
		'순번설정 
		.Col = C_SEQ_NO2	:	.Row = 1	:	.Text = "1"
		.Col = C_SEQ_NO2	:	.Row = 2	:	.Text = "2"
		.Col = C_SEQ_NO2	:	.Row = 3	:	.Text = "3"
		.Col = C_SEQ_NO2	:	.Row = 4	:	.Text = "4"
	End With
End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    .vspdData2.ReDraw = False
    .vspdData3.ReDraw = False

	' 1번 그리드 
    ggoSpread.Source = frm1.vspdData
        
'	ggoSpread.SSSetRequired C_W1, -1, -1
'	ggoSpread.SSSetRequired C_W2, -1, -1
'	ggoSpread.SSSetRequired C_W3, -1, -1
	ggoSpread.SpreadLock C_W1, -1, C_W1
	ggoSpread.SpreadLock C_W2, -1, C_W2
	ggoSpread.SpreadLock C_W3, -1, C_W3
    ggoSpread.SpreadLock C_W4, -1, C_W4
    ggoSpread.SpreadLock C_W5, -1, C_W5    
	ggoSpread.SpreadLock C_R3, -1, C_R3
    ggoSpread.SpreadLock C_R4, -1, C_R4
    
    ' 2번 그리드 
    ggoSpread.Source = frm1.vspdData2	

 	ggoSpread.SSSetRequired C_W6, -1, -1
 	ggoSpread.SSSetRequired C_W9, -1, -1
    ggoSpread.SpreadLock C_W7, -1, C_W7
    ggoSpread.SpreadLock C_W8, -1, C_W8
    ggoSpread.SpreadLock C_W10, -1, C_W10

	' 3번 그리드 
    ggoSpread.Source = frm1.vspdData3	

    ggoSpread.SpreadLock C_W11, -1, C_W11
    ggoSpread.SpreadLock C_W12, -1, C_W12
    ggoSpread.SpreadLock C_W13, -1, C_W13   
	ggoSpread.SSSetRequired C_W14, -1, -1
	ggoSpread.SSSetRequired C_W15, -1, -1
	ggoSpread.SSSetRequired C_W16, -1, -1
	ggoSpread.SSSetRequired C_W17, -1, -1
	ggoSpread.SSSetRequired C_W18, -1, -1
	ggoSpread.SSSetRequired C_W19, -1, -1
	ggoSpread.SSSetRequired C_W20, -1, -1
	ggoSpread.SSSetRequired C_W21, -1, -1
	ggoSpread.SSSetRequired C_W22, -1, -1
	ggoSpread.SSSetRequired C_W23, -1, -1

    .vspdData.ReDraw = True
    .vspdData2.ReDraw = True
    .vspdData3.ReDraw = True

    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal pGub)
    With frm1

	'If lgCurrGrid = 1 Then
		'.vspdData.ReDraw = False
 
		ggoSpread.Source = .vspdData

		If pGub <> "R" Then
			ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W2, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W3, pvStartRow, pvEndRow
		Else
			ggoSpread.SpreadLock C_W1, pvEndRow, C_W1
			ggoSpread.SpreadLock C_W2, pvEndRow, C_W2
			ggoSpread.SpreadLock C_W3, pvEndRow, C_W3
		End If
	    ggoSpread.SpreadLock C_W4, pvEndRow, C_W4
	    ggoSpread.SpreadLock C_W5, pvEndRow, C_W5    
		ggoSpread.SpreadLock C_R3, pvEndRow, C_R3
	    ggoSpread.SpreadLock C_R4, pvEndRow, C_R4
	    
		'.vspdData.ReDraw = True

    'End If
    
    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColorDetail2(ByVal pvStartRow, ByVal pvEndRow, ByVal pGub, byVal pR2)
	Dim iSeqNo, iChildSeqNo, iCnt, iRow, dblW9
	Dim iR2,tmpSeqNo,tmpSeqNo1
	
    With frm1
    
		' 2번 그리드 
		ggoSpread.Source = .vspdData2	
		.vspdData.Row=.vspdData.ActiveRow : .vspdData.Col = C_SEQ_NO :tmpSeqNo1= .vspdData.Value
		.vspdData.Row=.vspdData.ActiveRow : .vspdData.Col = C_R2 :iR2= .vspdData.Value
		
		
		If pGub <> "R" Then
		
		 	ggoSpread.SSSetRequired C_W6, pvStartRow, pvEndRow
		 	ggoSpread.SSSetRequired C_W7, pvStartRow, pvEndRow
		 	ggoSpread.SSSetRequired C_W9, pvStartRow, pvEndRow
	    	ggoSpread.SpreadLock C_W8, pvEndRow, C_W8
	    	ggoSpread.SpreadLock C_W10, pvEndRow, C_W10
		 	

'	 		For iRow = 1 To .vspdData2.MaxRows
'				.vspdData2.Row = iRow	:	.vspdData2.Col = C_CHILD_SEQ_NO	:	iChildSeqNo = .vspdData2.Text
'	 			.vspdData2.Col = C_SEQ_NO	:	.vspdData2.Row = iRow : tmpSeqNo=.vspdData2.Value
'				iCnt = CheckDetailData(.vspdData2, .vspdData2, iRow)
'				
'			 	If iCnt > 1 Then
'		 			If iChildSeqNo = 999999 Then
'				    	ggoSpread.SpreadLock C_W6, iRow, C_W8
'					    ggoSpread.SpreadUnLock C_W9, iRow, C_W9
'					 	ggoSpread.SSSetRequired C_W9, iRow, iRow
'					    ggoSpread.SpreadLock C_W10, iRow, C_W10
'		 			Else
'				 		ggoSpread.SSSetRequired C_W6, iRow, iRow
'					 	ggoSpread.SSSetRequired C_W7, iRow, iRow
'				    	ggoSpread.SpreadLock C_W9, iRow, C_W9
'				    End If
'			 	Else			 
'					
'				 	ggoSpread.SSSetRequired C_W6, iRow, iRow
'				 	ggoSpread.SSSetRequired C_W7, iRow, iRow
'				 	ggoSpread.SSSetRequired C_W9, pvStartRow, pvEndRow
'
'
'			 	End If
'	 		Next
		Else
			
	    	ggoSpread.SpreadLock C_W6, pvEndRow, C_W6
	    	ggoSpread.SpreadLock C_W7, pvEndRow, C_W7
	    	ggoSpread.SpreadLock C_W9, pvEndRow, C_W9
	    End If
	    ggoSpread.SpreadLock C_W8, pvEndRow, C_W8
	    ggoSpread.SpreadLock C_W10, pvEndRow, C_W10
    
    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor3(ByVal pvStartRow, ByVal pvEndRow, ByVal pGub, ByVal pType)
    With frm1
    
		' 3번 그리드 
		ggoSpread.Source = .vspdData3	

	    ggoSpread.SpreadLock C_W11, -1, C_W11
	    ggoSpread.SpreadLock C_W12, -1, C_W12
	    ggoSpread.SpreadLock C_W13, -1, C_W13   
		If pGub <> "R" Then
			ggoSpread.SSSetRequired C_W14, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W15, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W16, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W17, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W18, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W19, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W20, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W21, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W22, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W23, pvStartRow, pvEndRow
		Else
			ggoSpread.SpreadLock C_W14, pvEndRow, C_W14
			ggoSpread.SpreadLock C_W15, pvEndRow, C_W15
			ggoSpread.SpreadLock C_W16, pvEndRow, C_W16
			ggoSpread.SpreadLock C_W17, pvEndRow, C_W17
			ggoSpread.SpreadLock C_W18, pvEndRow, C_W18
			ggoSpread.SpreadLock C_W19, pvEndRow, C_W19
			ggoSpread.SpreadLock C_W20, pvEndRow, C_W20
			ggoSpread.SpreadLock C_W21, pvEndRow, C_W21
			ggoSpread.SpreadLock C_W22, pvEndRow, C_W22
			ggoSpread.SpreadLock C_W23, pvEndRow, C_W23
		End If
	
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W7		= iCurColumnPos(2)
            C_W9		= iCurColumnPos(3)
            C_W8		= iCurColumnPos(4)
            C_W9		= iCurColumnPos(6)
            C_W10		= iCurColumnPos(7)
            C_W11		= iCurColumnPos(8)
            C_W12       = iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W17		= iCurColumnPos(13)
            C_W18		= iCurColumnPos(14)
            C_W19		= iCurColumnPos(15)
            C_W20		= iCurColumnPos(16)
    End Select    
End Sub

Sub InsertRow2Head()
	' fncNew, onLoad시에 호출해서 기본적으로 3칸을 입력함 
	Dim ret, iRow, iLoop
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
			
		.ReDraw = False

		iRow = 1
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow, "I") 
		.Col = C_SEQ_NO : .Row = iRow: .Text = iRow	
		
		iRow = 2
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow, "I") 
		.Col = C_SEQ_NO : .Row = iRow: .Text = "999999"	
		
		.col = C_W1 : .CellType = 1 : .text = "계" : .TypeHAlign = 2
		.col = C_W5 : .CellType = 1 : .text = "100%" : .TypeHAlign = 2
		.col = C_R4 : .CellType = 1 : .text = "100%" : .TypeHAlign = 2
				
		ggoSpread.SpreadLock C_W1, iRow, C_R4, iRow
		
		.ReDraw = True		
		.focus
		.SetActiveCell C_W1, 1
					
	End With

	'Call InsertRow2Detail2(1)
	Call InsertRowHead3(0, 1)
	
	Call vspdData_Click(C_W1, 1)
End Sub

Sub InsertRow2Detail2(Byval pSeqNo)

	' 작업진행률 그리드 추가 
	Dim ret, iRow, iLoop, iLastRow
	Dim tmpR2
	
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_R2
	tmpR2=frm1.vspdData.Text
	
	
	With frm1.vspdData2
		
		.focus
		ggoSpread.Source = frm1.vspdData2

		'iLastRow = .MaxRows
		'.SetActiveCell C_W6, iLastRow	
		.Row = .ActiveRow
		.Col = C_CHILD_SEQ_NO
		If .Text = "999999" Then
			iRow = .ActiveRow - 1
		Else
			iRow = .ActiveRow
		End If
		
		.ReDraw = False
		'ggoSpread.ClearSpreadData

		ggoSpread.InsertRow iRow, 1
		
		iRow = iRow + 1
		.Row = iRow
		.Col = C_CHILD_SEQ_NO	: .Text = iRow
		.Col = C_SEQ_NO			: .Text = pSeqNo
		Call SetSpreadColorDetail2(iRow, iRow, "I",tmpR2) 
		'.RowHidden = True
		
		.SetActiveCell C_W6, iRow	
		.ReDraw = True		

	End With
	
End Sub

Sub InsertRowHead3(ByVal pRow, ByVal pSeqNo)
	' fncNew, onLoad시에 호출해서 기본적으로 3칸을 입력함 
	Dim ret, iRow, iLoop
	
	With frm1.vspdData3
		ggoSpread.Source = frm1.vspdData3
			
		.ReDraw = False
'		.ActiveRow = .MaxRows

		iRow = pRow + 4
		ggoSpread.InsertRow iRow ,1

		iRow = iRow + 1
		Call SetSpreadColor3(iRow, iRow, "I", "D") 
		.Col = C_W_TYPE : .Row = iRow: .Text = "D"
		.Col = C_SEQ_NO2 : .Row = iRow: .Text = pSeqNo
		
		If pSeqNo = 1 Then
			iRow = 6 : .Col=C_SEQ_NO2 : .Row=iRow
			If .Text <>"999999"	Then 
						
				ggoSpread.InsertRow 5,1
				.Col = C_W_TYPE : .Row = iRow: .Text = "D"
				.Col = C_SEQ_NO2 : .Row = iRow: .Text = "999999"	
			
				.col = C_W11 : .CellType = 1 : .text = "계" : .TypeHAlign = 2
						
				ggoSpread.SpreadLock C_W11, iRow, C_W23, iRow
				frm1.vspdData3.SetActiveCell C_W11, frm1.vspdData3.MaxRows-1
			
			End If
			'iRow = 6 : .Col=C_W11 : .Row=iRow
			
		End If
		
		.ReDraw = True		
					
	End With

'	Call vspdData3_Click(C_W11, 5)
End Sub


' -- 헤더쪽 그리드 재조정 
Sub RedrawSumRow(ByVal pGub)
	Dim iRow, iMaxRows, iSeqNo, ret
	
	With frm1.vspdData
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData
		.Redraw = false
		For iRow = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = iRow : iSeqNo = .Value
			
			If iSeqNo = 999999 Then ' 합계행 
				.col = C_W1 : .text = "계" : .TypeHAlign = 2
				
				ggoSpread.SpreadLock C_W1, iRow, C_R4, iRow
			Else
				ggoSpread.SpreadUnLock C_W1, iRow, C_R4, iRow
				Call SetSpreadColor(iRow, iRow, pGub)
			End If
		Next
		.Redraw = True
	End With
End Sub

' --  2번째 그리드 합계 재조정 
Sub RedrawSumRow2(ByVal pGub)
	Dim iRow, iMaxRows, iSeqNo, ret
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData2
		.Redraw = false
		For iRow = 1 to iMaxRows
			.Col = C_CHILD_SEQ_NO : .Row = iRow : iSeqNo = .value
			
			If iSeqNo = 999999 Then ' 합계행 
			
				.Col = C_W6		:	.CellType = 1 : .Text = "계"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail2(iRow,iRow, pGub,"S") 

			Else
				ggoSpread.SpreadUnLock C_W6, iRow, C_W10, iRow
				Call SetSpreadColorDetail2(iRow, iRow, pGub,"S")
			End If
		Next
		.Redraw = True
	End With
End Sub

' -- 3번째 그리드 합계 재조정 
Sub RedrawSumRow3(ByVal pGub)
	Dim iRow, iMaxRows, iSeqNo, ret, iWType
	
	With frm1.vspdData3
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData3
		.Redraw = false
		For iRow = 1 to iMaxRows
			.Col = C_W_TYPE	:	.Row = iRow : iWType = .text
			.Col = C_SEQ_NO2 : .Row = iRow : iSeqNo = .Value
			
			If iSeqNo = 999999 Then ' 합계행 
			
				.Col = C_W11	:	.CellType = 1 : .Text = "계"	: .TypeHAlign = 2	

				ggoSpread.SpreadLock C_W11, iRow, C_W23, iRow	
			Else
			
				Call SetSpreadColor3(iRow,iRow, pGub, iWType) 

'				ggoSpread.SpreadLock C_W11, iRow, C_W23, iRow	
			End If
		Next
		.Redraw = True
	End With
End Sub


'============================== 사용자 정의 함수  ========================================

' -- 행을 히든 처리 
Function ShowRowHidden(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
		iMaxRows = .MaxRows : iFirstRow = 0
		
		For iRow = 1 To iMaxRows
			.Col = C_SEQ_NO : .Row = iRow : iSeqNo = .value
			If iSeqNo = pSeqNo Then	' 같은 관계라면..
				.RowHidden = False
				If iFirstRow = 0 Then iFirstRow = iRow
			Else
				.RowHidden = True
			End If
		Next
		
		ShowRowHidden = iFirstRow
	End With
End Function
'vspdData_click 시 vspdData3위치 동기화 
Function ShowGrid3Row(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
		iMaxRows = .MaxRows -1
		
		For iRow = 5 To iMaxRows
			.Col = C_SEQ_NO2 : .Row = iRow : iSeqNo = .TEXT
			If iSeqNo = pSeqNo Then	' 같은 관계라면..
				ShowGrid3Row=iRow : .Col=C_W11 :.Focus
			End If
		Next		
	End With
End Function 

' -- 합계 행인지 체크(Header Grid)
Function CheckTotalRow(Byref pObj, Byval pRow, ByVal pCol) 
	CheckTotalRow = False
	pObj.Col = pCol : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = "999999" Then	 ' 합계 행 
		CheckTotalRow = True
	End If
End Function

' -- 합계 행인지 체크(Detail Grid)
Function CheckTotalRow2(Byref pObj, Byval pRow) 
	CheckTotalRow2 = False
	pObj.Col = C_CHILD_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = "999999" Then	 ' 합계 행 
		CheckTotalRow2 = True
	End If
End Function

' -- Detail Data가 존재하는지 체크 
Function CheckDetailData(Byref pObj, Byref pObjDe, Byval pRow) 
	Dim iSeqNo, iRow
	CheckDetailData = 0
	pObj.Col = C_SEQ_NO : pObj.Row = pRow	:	iSeqNo = Trim(pObj.Text)
	
	With pObjDe
		For iRow = 1 To .MaxRows
			.Row = iRow	:	.Col = C_SEQ_NO
			If Trim(.Text) = iSeqNo Then
				.Col = 0
				If .Text <> ggoSpread.DeleteFlag Then
					CheckDetailData = CheckDetailData + 1
				End If
			End If
		Next
	End With
End Function

' 1번 그리드의 성명을 하위 그리드에 적용한다.
Function SetG3HeaderW1(ByVal pCol, ByVal pRow, ByVal pTxt)
	Dim iSeq_no, iRow
	
	With Frm1.vspdData
		.Col = C_SEQ_NO	:	.Row = pRow	:	iSeq_no = .Text
	End With
	
	With Frm1.vspdData3
		For iRow = 5 To .MaxRows
			.Col = C_SEQ_NO2	:	.Row = iRow
			If iSeq_no = .Text Then
				.Col = C_W11	:	.Row = iRow	:	.Text = pTxt
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.UpdateRow iRow
			End If
		Next
	End With
End Function

' 1번 그리드의 구성비를 하위 그리드에 적용한다.
Function SetG3HeaderW5(ByVal pCol, ByVal pRow, ByVal pTxt)
	Dim iSeq_no, iRow
	
	With Frm1.vspdData
		.Col = C_SEQ_NO	:	.Row = pRow	:	iSeq_no = .Text
	End With
	
	With Frm1.vspdData3
		For iRow = 5 To .MaxRows
			.Col = C_SEQ_NO2	:	.Row = iRow
			If iSeq_no = .Text Then
				.Col = C_W12	:	.Row = iRow	:	.Text = pTxt
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.UpdateRow iRow
			End If
		Next
	End With
End Function

' 1번 그리드의 (4)금액을 계산하고 (5)구성비를 설정하며 합계를 설정한다.
Function SetG1W4_W5(ByVal Col, ByVal Row)
	Dim iRow, dblSum, dblW2, dblW3, dblW4, txtW5
	
	With Frm1.vspdData
		For iRow = 1 To .MaxRows - 1
			.Col = C_W2	:	.Row = iRow	:	dblW2 = UNICDbl(.Text)
			.Col = C_W3	:	.Row = iRow	:	dblW3 = UNICDbl(.Text)
			dblW4 = dblW2 - dblW3
			If dblW4 < 0 Then
				.Col = C_W4	:	.Row = iRow	:	.Text = 0
			Else
				.Col = C_W4	:	.Row = iRow	:	.Text = dblW2 - dblW3
			End If
		Next
		
		dblSum = FncSumSheet(Frm1.vspdData, C_W2, 1, .MaxRows - 1, true, .MaxRows, C_W2, "V")	' 가지급금 적수 합계 
		dblSum = FncSumSheet(Frm1.vspdData, C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' 가수금 적수 합계 
		dblSum = FncSumSheet(Frm1.vspdData, C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 금액 합계 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow .MaxRows

		If dblSum > 0 Then
			For iRow = 1 To .MaxRows - 1
				.Col = C_W4	:	.Row = iRow	:	dblW4 = UNICDbl(.Text)
				.Col = C_W5	:	.Row = iRow	:	.Value = (dblW4 / dblSum) 
				txtW5 = .Text
				Call SetG3HeaderW5(Col, iRow, txtW5)
			Next
		End If
		dblSum = FncSumSheet(Frm1.vspdData, C_R3, 1, .MaxRows - 1, true, .MaxRows, C_R3, "V")	' 회사부담 적용적수금액 합계 

		If dblSum > 0 Then
			For iRow = 1 To .MaxRows - 1
				.Col = C_R3	:	.Row = iRow	:	dblW4 = UNICDbl(.Text)
				.Col = C_R4	:	.Row = iRow	:	.Value = (dblW4 / dblSum) 
			Next
		End If
	End With
End Function


' 2번 그리드의 (6)이자율을 확인하고 다른필드들 계산을 한다.
Function SetGrid2(ByVal pCol, ByVal pRow)
	Dim iG1R2, iSeqNo, iG1Row

	' Header의 초기값 가져오기 
	With Frm1.vspdData
		iG1Row = .ActiveRow
		.Col = C_R2	:	.Row = iG1Row	:	iG1R2 = .Text
		.Col = C_SEQ_NO	:	.Row = iG1Row	:	iSeqNo = Trim(.Text)
	End With
	
	' 체크사항 및 계산부분 
	With Frm1.vspdData2
		' (6)이자율에 대한 체크 
'		If pCol = C_W6 And CheckW6(iG1R2, iSeqNo, iG1Row) <> True Then
'			.Col = pCol	:	.Row = pRow	:	.Text = ""
'			MsgBox "(6)이자율을 확인하십시오.", vbCritical
'			Exit Function
'		End If
		
		' (7)적수에 대한 계산 
		Call SetW7(iG1R2, iSeqNo, iG1Row)
		
		' (8)인정이자에 대한 계산(각 행별로 (7)X(6)/365
		Call SetW8(pCol, pRow, iSeqNo)
		 '(10)조정액 등록 
		Call SetW10(pCol, pRow, iG1R2, iSeqNo, iG1Row)
		' (9)회사계상액에 이자수령액을 입력 
		 Call SetW9(pCol, pRow, iG1R2, iSeqNo, iG1Row)
		 
		

	End With
End Function

Function CheckW6(ByVal pG1R2, ByVal pSeqNo, ByVal pG1Row)
	Dim iRow, iCSeq_no
	
	CheckW6 = True
	
	If pG1R2 = "1" And CheckDetailData(Frm1.vspdData, Frm1.vspdData2, pG1Row) > 1 Then
		CheckW6 = False
	Else
		
		With Frm1.vspdData2
			For iRow = 1 To .MaxRows - 1
				.Row = iRow	:	.Col = C_SEQ_NO
				If Trim(.Text) = pSeqNo Then
					.Col = C_CHILD_SEQ_NO	:	iCSeq_no = .Text
					.Col = C_W6
					If .Text = "" Or iCSeq_no = "999999"Then
					ElseIf lgRateConf > UNICDbl(.Value) Then
						CheckW6 = False
					ElseIf CheckExistW11(.Text) <> True Then
						CheckW6 = False
					End If
				End If
			Next
		End With
	End If
End Function

Function SetW7(ByVal pG1R2, ByVal pSeqNo, ByVal pG1Row)
	Dim dblW4, dblR3, dblR4, dblW6, iRow, dblW7, iRowConf

	dblW4 = 0	:	dblR3 = 0	:	dblR4 = 0
	With Frm1.vspdData
		.Col = C_W4	:	.Row = pG1Row	:	dblW4 = UNICDbl(.Text)
		.Col = C_R3	:	.Row = pG1Row	:	dblR3 = UNICDbl(.Text)
		.Col = C_R3	:	.Row = .MaxRows	:	dblR4 = UNICDbl(.Text)
	End With

	If pG1R2 = "1" Then
		If dblW4 > 0 Then	
			With Frm1.vspdData2
				For iRow = 1 To .MaxRows
					.Row = iRow	:	.Col = C_SEQ_NO
					If Trim(.Text) = pSeqNo Then
						.Col = C_W7
						.Text = dblW4
					End If
				Next
			End With
		End If
		
	ElseIf DblR3 > 0 And DblR4 > 0 Then
		dblW7 = 0
		With Frm1.vspdData2
			For iRow = 1 To .MaxRows
				.Row = iRow	:	.Col = C_SEQ_NO
				If Trim(.Text) = pSeqNo Then
					.Row = iRow	:	.Col = C_CHILD_SEQ_NO
					If .Text = 999999 Then
						.Row = iRow	:	.Col = C_W6	:	dblW6 = 0
					Else
						.Row = iRow	:	.Col = C_W6	:	dblW6 = UNICDbl(.Value)
					End If
					
					If dblW6 > lgRateConf Then
						.Col = C_W7
						.Text = GetG3W14(dblW6) * (dblR3 / dblR4)
						dblW7 = dblW7 + UNICDbl(.Text)
					Else
						iRowConf = .Row
					End If
				End If
			Next
			.Row = iRowConf	:	.Col = C_W7	:	.Text = dblW4 - dblW7
		End With
	End If

	'합계구하기...ㅜ.ㅜ 
	dblW7 = 0
	With Frm1.vspdData2
		For iRow = 1 To .MaxRows
			.Row = iRow	:	.Col = C_SEQ_NO
			If Trim(.Text) = pSeqNo Then
				.Row=iRow : .Col =0
				If .text<>"삭제" Then				
					.Row = iRow	:	.Col = C_CHILD_SEQ_NO
					If .Text <> 999999 Then
						.Col = C_W7
						dblW7 = dblW7 + UNICDbl(.Text)						
					Else
						iRowConf = .Row
					End If
				End If
			End If
		Next
		
		.Row = iRowConf	:	.Col = C_W7	:	.value = dblW7
	End With
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow iRowConf
End Function

Function SetW8(ByVal pCol, ByVal pRow, ByVal pSeqNo)
	Dim dblW6, iRow, dblW7, dblSum, iSumRow

	dblSum = 0
	With Frm1.vspdData2
		For iRow = 1 To .MaxRows
			.Row = iRow	:	.Col = C_SEQ_NO
			If Trim(.Text) = pSeqNo Then
				.Col=0
				If .Text <>"삭제" Then
					.Col = C_CHILD_SEQ_NO
					If Trim(.Text) = "999999" Then
						iSumRow = iRow
					Else

						.Col = C_W7	:	dblW7 = UNICDbl(.Text)
						.Col = C_W6	:	dblW6 = UNICDbl(.Value)
	
						.Col = C_W8
						If lgblnYoon Then
							' 윤년 
							.Text = (dblW7 * dblW6) / 366
						Else	
							' 평년 
							.Text = (dblW7 * dblW6) / 365
						End If
						dblSum = dblSum + UNICDbl(.Text)
					End If
				End If
			End If
		Next
		.Row = iSumRow	:	.Col = C_W8	:	.VALUE= dblSum
	End With
End Function

Function SetW9(ByVal pCol, ByVal pRow, ByVal pG1R2, ByVal pSeqNo, ByVal pG1Row)
	dim i
	
		with frm1.vspdData2
			for i=pRow to frm1.vspdData2.maxRows
				.col=2 :				.row = i
				if .value ="999999" then
					Call FncSumSheet(frm1.vspdData2, C_W9, 1,i-1, true, i, C_W9, "V")	' 합계 
					'call  vspdData2_Change( C_W9 , i  )
					
					Call FncSumSheet(frm1.vspdData2, C_W9+1, 1,i-1, true, i, C_W9+1, "V")	' 합계 
					exit for
				end if
			next
		end with
	'Call FncSumSheet(frm1.vspdData2, C_W9, 1, frm1.vspdData2.MaxRows-1, true, frm1.vspdData2.MaxRows, C_W9, "V")	' 합계 
End Function

Function SetW10(ByVal pCol, ByVal pRow, ByVal pG1R2, ByVal pSeqNo, ByVal pG1Row)
	Dim iRow, dblW8, dblW9, dblW10, iCnt

'	iCnt = CheckDetailData(frm1.vspdData2, frm1.vspdData2, pRow)
'	If iCnt > 1 Then
'		With Frm1.vspdData2
'			For iRow = 1 To .MaxRows
'				.Row = iRow	:	.Col = C_SEQ_NO
'				If Trim(.Text) = pSeqNo Then
'					.Row = iRow	:	.Col = C_CHILD_SEQ_NO
'	
'					If .Text = 999999 Then
'						.Row = iRow	:	.Col = C_W8	:	dblW8 = UNICDbl(.Text)
'						.Row = iRow	:	.Col = C_W9	:	dblW9 = UNICDbl(.Text)
'						.Row = iRow	:	.Col = C_W10
'						dblW10 = dblW8 - dblW9
'						If dblW10 < 0 Then
'							.Text = 0
'						Else
'							.Text = dblW10
'						End If
'					End If
'				End If
'			Next
'		End With
'	ElseIf iCnt = 1 Then
		With Frm1.vspdData2
		
			.Row = pRow	:	.Col = C_W8	:	dblW8 = UNICDbl(.Text)
			.Row = pRow	:	.Col = C_W9	:	dblW9 = UNICDbl(.Text)
			.Row = pRow	:	.Col = C_W10
			dblW10 = dblW8 - dblW9
		'	msgbox dblW8 
	'		msgbox  dblW9
			If dblW10 < 0 Then
				.Text = 0
			Else
				.Text = dblW10
			End If
		End With
		
		'Call FncSumSheet(frm1.vspdData2, C_W10, 1, frm1.vspdData2.MaxRows-1, true, frm1.vspdData2.MaxRows, C_W10, "V")	' 합계 
'	End If
End Function


Function CheckExistW11(ByVal pW6)
	Dim iCol
	
	CheckExistW11 = False
	If Frm1.vspdData3.MaxRows > 0 Then
		With Frm1.vspdData3
			For iCol = C_W14 To C_W23
				.Row = 1	:	.Col = iCol
				If Trim(.Text) = "" Then
				
				ElseIf .Text = pW6 Then
					CheckExistW11 = True
					Exit Function
				End If
			Next
		End With
	Else
		CheckExistW11 = True
	End If
End Function


' 3번 그리드의 (11)이자율을 확인하고 다른필드들 계산과 체크한다.
Function SetGrid3(ByVal pCol, ByVal pRow, ByVal pGubun)
	Dim dblSum, dblW14, iCol, iRow, dblRate
	Dim tmpVal
	
	
	' 체크사항 및 계산부분 
	With Frm1.vspdData3
		If pGubun = "C" Then
			' (11)이자율에 대한 체크 
			If pRow = 1 And CheckW11() <> True Then
				.Col = pCol	:	.Row = pRow	:	.Text = ""
				MsgBox "(11)이자율을 확인하십시오. (9% 보다 작을 수 없습니다)", vbCritical
				Exit Function
			End If
		End If

			
		If pRow = 2 Or pRow = 3 Then
			' (12)지급이자에 대한 합계 
			dblSum = FncSumSheet(frm1.vspdData3, 2, C_W14, C_W23, false, -1, -1, "H")	' 현재 열 합계 
			.Col = C_W13 : .Row = 2 : .Value = dblSum
			' (12)지급이자로  (13)차입금적수를 계산 
			If pRow = 2 Then
				.Col = pCol : .Row = 2 : dblW14 = UNICDbl(.Text)
				.Col = pCol : .Row = 1 : dblRate = UNICDbl(.value)
				.Col = pCol : .Row = 3
				If dblRate > 0 Then
					If lgblnYoon Then
						' 윤년 
						.Text = (dblW14 / dblRate) * 366
					Else	
						' 평년 
						.Text = (dblW14 / dblRate) * 365
					End If
				End If
			End If

			
			' (13)차입금적수를 (14)이자율별 적용적수에 입력하고 합계계산 
			.Col = pCol : .Row = 3 : dblW14 = .Text
			.Col = pCol : .Row = 4 : .Text = dblW14

			dblSum = FncSumSheet(frm1.vspdData3, 3, C_W14, C_W23, false, -1, -1, "H")	' 현재 열 합계 
			.Col = C_W13 : .Row = 3 : .value = dblSum
			
			' (14)이자율별적용적수의 합계 
			dblSum = FncSumSheet(frm1.vspdData3, 4, C_W14, C_W23, false, -1, -1, "H")	' 현재 열 합계 
			.Col = C_W13 : .Row = 4 : .Value = dblSum
		
			If pGubun = "R" Then
				 '(14)이자율별적용적수의 합계는 Grid1의 회사부담적용적수금액의 합계를 한도로 한다.
				If CheckW4_W14(UNICDbl(dblSum)) <> True Then
					dblSum = FncSumSheet(frm1.vspdData3, 4, C_W14, C_W23, false, -1, -1, "H")	' 현재 열 합계 
					'(4) >=(14) 일경우 에러 
					frm1.vspdData.Row=frm1.vspdData.MaxRows : frm1.vspdData.Col = C_W4
					tmpVal = UNICDbl(frm1.vspdData.Value)
					If tmpVal < dblSum Then 
						Call DisplayMsgBox("WC0015", "X", "(14)이자율별적용적수", "(4)금액")           '⊙: "Will you destory previous data"
						.Row=iRow : .Col = pCol : .Value=0
						.Row=pRow : .Col = pCol : .Value=0
						.Row=.MaxRows : .Col = pCol : .Value=0
						Exit Function 
					Else
					.Col = C_W13 : .Row = 4 : .Value = dblSum
					End If
					
				End If
			End If
		End If
		
		'배분계산 
		Call SetCalcDivision(pRow)

		' 당좌대출이자율인것만 합계하여 이자율별 적수에 넣어주고 합계를 구한다.
		'인명별 배분 계산 합계 
		For iCol = C_W13 To C_W23
			dblSum = FncSumSheet(frm1.vspdData3, iCol, 5, .MaxRows -1, false, -1, -1, "V")	' 현재 열 합계 
			.Col = iCol :	.Row = .MaxRows : .Value = dblSum
			.Col = iCol	:	.Row = 1
			If UNICDbl(.Value) = lgRateConf Then
				.Row = 4	:	.Text = dblSum
			End If
		Next
		'(4) >=(14) 일경우 에러 
		frm1.vspdData.Row=frm1.vspdData.MaxRows : frm1.vspdData.Col = C_W4
		tmpVal = UNICDbl(frm1.vspdData.Value)
		For iRow = 4 To .MaxRows
			dblSum = FncSumSheet(frm1.vspdData3, iRow, C_W14, C_W23, false, -1, -1, "H")	' 현재 열 합계 
			
			If iRow= 4 Then 
				If tmpVal < UNICDbl(dblSum) Then 
			
					Call DisplayMsgBox("WC0015", "X", "(14)이자율별적용적수", "(4)금액")           '⊙: "Will you destory previous data"
					.Row=iRow : .Col = pCol : .Value=0
					.Row=pRow : .Col = pCol : .Value=0
					.Row=.MaxRows : .Col = pCol : .Value=0
					Exit Function 
				End If
			End IF	
			.Col = C_W13 : .Row = iRow : .value = dblSum		
		Next
	End With
End Function

Function CheckW11()
	Dim iCol
	
	CheckW11 = True
	
	With Frm1.vspdData3
		For iCol = C_W14 To C_W23
			.Row = 1	:	.Col = iCol
			If .Text = "" Then
			ElseIf UNICDbl(.Value)  > 0 and lgRateConf > UNICDbl(.Value) Then
				CheckW11 = False
				Exit Function
			End If
		Next
	End With
End Function

Function CheckW4_W14(ByVal pW14Sum)
	Dim dblW4Sum, iCol, dblW14
	
	CheckW4_W14 = True
	
	With Frm1.vspdData
		.Row = .MaxRows	:	.Col = C_R3
		If .MaxRows < 1 Then
			Exit Function
		ElseIf .Text = "" Then
			Exit Function
		ElseIf UNICDbl(.Text) = 0 Then
			Exit Function
		Else
			dblW4Sum = UNICDbl(.Text)
		End If
	End With

	If dblW4Sum < pW14Sum Then
		dblW4Sum = pW14Sum - dblW4Sum
		CheckW4_W14 = False
		With Frm1.vspdData3
			For iCol = C_W23 To C_W14 Step -1
				.Row = 1	:	.Col = iCol
				If UNICDbl(.Value) > 0 Then
					.Row = 4	:	.Col = iCol
					If .Text = "" Then
						dblW14 = 0
					Else
						dblW14 = UNICDbl(.Text)
					End If
					
					dblW14 = dblW14 - dblW4Sum
					
					If dblW14 < 0 Then
						.Text = 0
						dblW4Sum = -1 * dblW14
					Else
						.Text = dblW14
						Exit For
					End If
				End If
			Next
		End With
	End If
End Function

Function SetCalcDivision(ByVal pRow)
	Dim dblW23_SUM, dblW12, dblR3, dblR4, dblW4, dblW14
	Dim iRow, iCol, iSeqNo, iR2, dblRate, dblRateAmtSum
	
	With Frm1.vspdData3
		' 이자율별 적용적수의 합계를 구성비별로 나눈다.
		If pRow < 4 Then
			Exit Function
		ElseIf .MaxRows = pRow Then
			Exit Function
		End If
			
		.Col = C_SEQ_NO2	:	.Row = pRow	:	iSeqNo = .Text
		
	End With

	With Frm1.vspdData
		For iRow = 1 To .MaxRows - 1
			.Row = iRow	:	.Col = C_SEQ_NO
			If iSeqNo = .Text Then
				Exit For
			End If
		Next
		.Col = C_R2	:	.Row = iRow	:	iR2 = .Text
		.Col = C_W4	:	.Row = iRow	:	dblW4 = UNICDbl(.Text)
		.Col = C_R3	:	.Row = iRow	:	dblR3 = UNICDbl(.Text)
		.Col = C_R3	:	.Row = .MaxRows	:	dblR4 = UNICDbl(.Text)
	End With
	
	With Frm1.vspdData3
		' 이자율별 적용적수를 회사부담이자율을 선택한 사람에 대해 차감계비로 나누어 등록한다.
		If iR2 = "2" Then
			dblRateAmtSum = 0
			If dblR4 > 0 Then
				For iCol = C_W14 To C_W23
					.Col = iCol	:	.Row = 1	:	dblRate = UNICdbl(.Value)
					If dblRate = 0 Then
						Exit For
					ElseIf dblRate > lgRateConf Then
						.Col = iCol	:	.Row = 4	:	dblW14 = UNICDbl(.Text)
						.Col = iCol	:	.Row = pRow
						.Text = dblW14 * dblR3 / dblR4
						dblRateAmtSum = dblRateAmtSum + UNICDbl(.Text)
						Call SetGrid2W7_FROM3(dblRate, iSeqNo, UNICDbl(.Text))
					End If
				Next
				For iCol = C_W14 To C_W23
					.Col = iCol	:	.Row = 1	:	dblRate = UNICdbl(.Value)
					If dblRate = lgRateConf Then
						.Col = iCol	:	.Row = pRow
						.Text = dblR3 - dblRateAmtSum
						Call SetGrid2W7_FROM3(dblRate, iSeqNo, UNICDbl(.Text))
					End If
				Next
			End If
		ElseIf iR2 = "1" Then
			' 9%에 (4)금액을 넣고 만다.
			For iCol = C_W14 To C_W23
				.Col = iCol	:	.Row = 1	:	dblRate = UNICdbl(.Value)
				If dblRate = lgRateConf Then
					.Col = iCol	:	.Row = pRow
					.Text = dblW4
					Call SetGrid2W7_FROM3(dblRate, iSeqNo, UNICDbl(.Text))
					Exit For
				End If
			Next
		End If
	End With
End Function

Function SetGrid2W7_FROM3(ByVal pRate, ByVal pSeqNo, ByVal dblAmt)
	Dim iRow, dblSum, iChildNo, iTotRow
	
	With Frm1.vspdData2
		dblSum = 0
		For iRow = 1 To .MaxRows
			.Row = iRow	:	.Col = C_SEQ_NO
			If .Text = pSeqNo Then
				.Col = C_CHILD_SEQ_NO	:	iChildNo = .Text
				.Col = C_W6
				if .Text = "" Then
				ElseIf iChildNo = 999999 Then
					iTotRow = iRow
				ElseIf UNICDbl(.Value) = pRate Then
					.Col = C_W7	:	.Text = dblAmt
					dblSum = dblSum + UNICDbl(dblAmt)

					' (8)인정이자에 대한 계산(각 행별로 (7)X(6)/365
					Call SetW8(C_W7, iRow, pSeqNo)
					
 					'(10)조정액 등록 
					Call SetW10(C_W7, iRow, 0, pSeqNo, 0)
				Else
					.Col = C_W7	:	dblSum = dblSum + UNICDbl(.Text)
				End If
			End If
		Next
		.Col = C_W7	:	.Row = iTotRow	:	.Text = dblSum
	End With

End Function

Function CheckDivision()
	Dim iCol, iRow, dblW23_SUM, dblW23_Calc
	
	With Frm1.vspdData3
		For iRow = 5 To .MaxRows - 1
			.Col = C_W13	:	.Row = iRow	:	dblW23_SUM = UNICDbl(.Text)
			dblW23_Calc = 0
			For iCol = C_W14 To C_W23
				.Col = iCol	:	.Row = iRow	:	dblW23_Calc = dblW23_Calc + UNICDbl(.Text)
			Next
			If dblW23_SUM <> dblW23_Calc Then
				dblW23_Calc = dblW23_Calc - dblW23_SUM
				For iCol = C_W23 To C_W14 Step - 1
					.Col = iCol	: .Row = 1
					If UNICDbl(.Value) > 0 Then
						.Col = iCol	: .Row = iRow
						dblW23_SUM = UNICDbl(.Text) - dblW23_Calc
						If dblW23_SUM < 0 Then
							.Text = 0
							dblW23_Calc =  dblW23_SUM *  -1
						ElseIf dblW23_SUM = 0 Then
							.Text = 0
							Exit For
						Else
							.Text = dblW23_SUM
							Exit For
						End IF
					End IF
				Next
			End If
		Next
	End With
End Function

Function GetG3W14(ByVal dblW6)
	Dim iCol
	
	With Frm1.vspdData3
		.Row = 1
		For iCol = C_W14 To C_W23
			.Col = iCol
			If dblW6 = UNICDbl(.Value) Then
				Exit For
			End If
		Next
		.Row = 4
		GetG3W14 = UNICDbl(.Text)
	End With
End Function
'============================== 레퍼런스 함수  ========================================

Sub GetFISC_DATE()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, datMonCnt, i, datNow

	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		lgFISC_START_DT = CDate(lgF0)
	Else
		lgFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		lgFISC_END_DT = CDate(lgF1)
	Else
		lgFISC_END_DT = ""
	End if

	call CommonQueryRs(" CONVERT(NUMERIC(5,2), REFERENCE)"," B_CONFIGURATION "," MAJOR_CD = 'W2006' AND MINOR_CD = '1' AND SEQ_NO = 1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		lgRateConf = UNICDbl(lgF0)
	Else
		lgRateConf = 0.9
	End if

	lgblnYoon = False
	datMonCnt = DateDiff("m", lgFISC_START_DT, lgFISC_END_DT)
	' 현재 법인의 당기기간안에 윤달이 있는지 체크해서 lgblnYOON를 변화시킨다.
	For i = 1 To datMonCnt
		datNow = DateAdd("m", i, lgFISC_START_DT)
		If Month(datNow) = 2 Then	' 2월을 가지는 당기기간이면 
			lgblnYoon = CheckIntercalaryYear(Year(datNow))
			Exit For
		End If
	Next
End Sub

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData

	ggoSpread.Source = Frm1.vspdData2
    ggoSpread.ClearSpreadData

	ggoSpread.Source = Frm1.vspdData3
    ggoSpread.ClearSpreadData

    
			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
End Function

Function GetRefOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr, iSeqNo, iLastRow, iRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_CMODE
        
    Call SetToolbar("1111111100010111")										<%'버튼 툴바 제어 %>
	
'	Call RedrawSumRow("R")
	Call RedrawSumRow("Q")
	Call ChangeRowFlg(frm1.vspdData)
	Call RedrawSumRow2("Q")
	Call ChangeRowFlg(frm1.vspdData2)
	Call RedrawSumRow3("Q")
	Call ChangeRowFlg(frm1.vspdData3)

	With frm1.vspdData
		.Col = C_SEQ_NO : .Row = .ActiveRow : iSeqNo = .value
			
		' 하위 그리드 표시루틴'
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)

	End With			
	
	With Frm1.vspdData3
		For iRow = 1 to .MaxRows
			Call SetGrid3(C_W14, iRow, "R")
		Next
	End With
	Call vspdData_Change(C_W2,1)
	frm1.vspdData.focus			
End Function

Function ChangeRowFlg(iObj)
	Dim iRow
	
	With iObj
		ggoSpread.Source = iObj
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .value = ggoSpread.InsertFlag
		Next
	End With
End Function

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

'    iCalledAspName = AskPRAspName("W5105RA1")
    
 '   If Trim(iCalledAspName) = "" Then
  '      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
   '     IsOpenPop = False
    '    Exit Function
    'End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
   
    'Call InsertRow2Head
    'Call InsertRow2Detail(1)
    
    Call SetToolbar("1110110100010111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData 		
    
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call GetFISC_DATE
End Sub


Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	Call GetFISC_DATE
End Sub


'==========================================================================================
Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	Call GetFISC_DATE
	'on load checking data exist or not	
	
	Call CommonQueryRs("count(SEQ_NO)"," TB_19A_1H "," CO_CD= '" & frm1.txtCO_CD.value & "' AND FISC_YEAR='" & frm1.txtFISC_YEAR.text & "' AND REP_TYPE='" & frm1.cboREP_TYPE.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If uniCDbl(lgF0)<> 0 Then 
		ggoSpread.Source=frm1.vspdData3
		ggoSpread.ClearSpreadData		    
		Call DbQuery
	End IF

End Sub


'============================================  1번 그리드 이벤트  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case C_W1		' 성명(법인명)
				.Col = Col
				If .Text = "주택자금" Then
					.Col = C_W7_NM : .Text = "상여" : intIndex = .Value
					.Col = C_W7 : .Value = intIndex		
				Else
					.Col = C_W7_NM : .Text = "기타사외유출" : intIndex = .Value
					.Col = C_W7 : .Value = intIndex		
				End If
			Case  C_W7
				.Col = Col
				intIndex = .Value
				.Col = C_W7_NM
				.Value = intIndex	
			Case  C_W7_NM
				.Col = Col
				intIndex = .Value
				.Col = C_W7
				.Value = intIndex		
			Case  C_W8
				.Col = Col
				intIndex = .Value
				.Col = C_W8_NM
				.Value = intIndex	
			Case  C_W8_NM
				.Col = Col
				intIndex = .Value
				.Col = C_W8
				.Value = intIndex		
		End Select
	End With

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum, iRow
	
	With Frm1.vspdData
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
		    
		Select Case Col
			Case C_W1		' 성명 
				.Col = C_W1
				Call SetG3HeaderW1(Col, Row, .Text)	' 현재 이름을 하위 그리드에 넣는다.
			Case C_W2, C_W3		' 가지급금 적수 & 가수금적수 
				Call SetG1W4_W5(Col, Row)
		End Select

	End With
	With Frm1.vspdData3
		For iRow = 1 to .MaxRows
			Call SetGrid3(C_W14, iRow, "C")
		Next
		ggoSpread.Source = frm1.vspdData3
	    ggoSpread.UpdateRow 4
	    ggoSpread.UpdateRow .MaxRows

	End With

End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row
	
	Dim iSeqNo, IntRetCD, iLastRow
	
    ggoSpread.Source = frm1.vspdData
  
    If Row = frm1.vspdData.MaxRows Then
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
	Else
		With frm1.vspdData
			.Col = C_SEQ_NO : .Row = Row : iSeqNo = .Value
			
			' 하위 그리드 표시루틴'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			frm1.vspdData2.SetActiveCell C_W6, iLastRow			
			.focus
			
			iLastRow=ShowGrid3Row(frm1.vspdData3,iSeqNo)
			frm1.vspdData3.SetActiveCell C_W11, iLastRow
		End With	
	End If
	frm1.vspdData.Focus
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	lgCurrGrid = 1
	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
'    Call GetSpreadColumnPos("A")
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
		
	'	Call DbDtlQuery(NewRow)
	Call vspdData_Click(newcol,newrow)
		
'        frm1.vspddata.Col = 0
'		lgStrPrevKey=""
    End If
End Sub
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
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



'============================================  2번 그리드 이벤트  ====================================
' 이자율 변경, 적수 변경, 회사계상액 변경 
' (6)이자율 : 3117 (8) 인정이자율 종류에 따라 
'						1. 당좌대출이자율일 경우 
'							W2006의 이자율 입력 
'						2. 회사부담이자율일 경우 
'							3번 그리드의 (11)이자율의 이자율을 차례대로 입력 
' (7)적수 : 
'						1. 당좌대출이자율일 경우 
'							인명별의 (4)의 금액입력 
'						2. 회사부담이자율일 경우 
'							1) (6) 이자율이 당좌대출이자율을 초과하는 경우 
'								각 행별로 (14)이자율별적용적수 X 3117의 (5)차감계의 비율(단, (8)이 회사부담이율인 항목의 비율임)
'							2) (6) 이자율이 당좌대출이자율인 경우 
'								(4) 금액 - ???
' (8)인정이자 : 각행별로 (7)X(8) / 365(윤년인겨우 366)
' (9)회사상계액 : 인정이자 프로그램의 (6)이자수령액을 인명별로 입력하며, 인명별로 적용이자율이 2개 이상인 경우 소계란에 (6)이자수령액을 입력함.
' (10)조정액 : 이자율이 하나인 경우는 각행의 ⑧ - ⑨ 를 계산하여 입력하고, 
'				이자율이 둘 이상인 경우는 소계의 ⑧ - ⑨ 를 계산하여 입력함.
'				단, 조정액은 "0"이하인 경우는 "0"을 입력함.
'				조정액을 인정이자 프로그램의 (7)구분별로 합계하여  (1)과목명에 "인정이자" (2) 금액에는 구분별 합계금액 
'				(3) 소득처분은 인정이자의 "(7)구분"을 입력하고 
'				조정내용은 " 특수관계자에 대한 가지급금  인정이자를 계산하고 익금산입하고 "소득처분"로 처분함 
'================================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData2
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
		
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.UpdateRow Row
		
	    Select Case Col
			Case C_W6, C_W7, C_W8, C_W9, C_W10
				Call SetGrid2(Col, Row)	' W6 Check, Others Setting
	    End Select

	End With
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	Dim tmpR2
	
	With frm1
		.vspdData.Row= .vspdData.ActiveRow : .vspdData.Col = C_R2
		tmpR2=.vspdData.Value
		ggoSpread.Source = .vspdData2
		.vspdData2.Row =Row : .vspdData2.Col =C_W6
		If .vspdData2.Text="계" Then Exit sub
		If .vspdData2.Row <1 Then Exit Sub
		
		If tmpR2 ="1" Then		'lock		
			
			
			ggoSpread.SpreadLock C_W6, Row, C_W7, Row
			ggoSpread.SSSetRequired C_W6, Row,Row	
			ggoSpread.SSSetRequired C_W7, Row, Row
		Else
			ggoSpread.SpreadunLock  C_W6, Row, C_W7,Row
		    
			ggoSpread.SSSetRequired C_W6, Row, Row
			ggoSpread.SSSetRequired C_W7, Row, Row		
		End IF	
	
	End With

End Sub

'============================================  3번 그리드 이벤트  ====================================
Sub vspdData3_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData3
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
		
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.UpdateRow Row
		
		Call SetGrid3(Col, Row, "C")	' W11 Check, Others Setting, Summary Set
		
		ggoSpread.Source = frm1.vspdData3
		If Row = 2 Then
		    ggoSpread.UpdateRow 3
		    ggoSpread.UpdateRow 4
		ElseIf Row > 4 Then
		    ggoSpread.UpdateRow frm1.vspdData3.MaxRows
		End If


	End With
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 3
	ggoSpread.Source = Frm1.vspdData3
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData3
   
    If frm1.vspdData3.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = frm1.vspdData3
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData3.Row = Row
	
	Dim iSeqNo, IntRetCD, iLastRow
	
    ggoSpread.Source = frm1.vspdData3
  
  
End Sub





'============================================  툴바지원 함수  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                                <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData2
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData3
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
    
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
'    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
	Dim blnChange
        
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>

    
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange <> False Then
		blnChange = True
'	    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
'	    Exit Function
	End If
	
	If  ggoSpread.SSDefaultCheck =False Then                                         '☜: Check contents area
	      Exit Function
	End If    

    ggoSpread.Source = frm1.vspdData2
	If ggoSpread.SSCheckChange <> False Then
		blnChange = True
'	    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
'	    Exit Function
	End If
	
	If   ggoSpread.SSDefaultCheck =false Then                                         '☜: Check contents area
	      Exit Function
	End If    

    ggoSpread.Source = frm1.vspdData3
	If ggoSpread.SSCheckChange <> False Then
		blnChange = True
'	    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
'	    Exit Function
	End If
	
	If   ggoSpread.SSDefaultCheck =False Then                                         '☜: Check contents area
	      Exit Function
	End If
	
    If Not blnChange Then
		Call DisplayMsgBox("900001", "X", "X", "X")           '⊙: "Will you destory previous data"
		Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD , blnChange

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData2
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData3
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
	
    If lgBlnFlgChgValue Or blnChange Then
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
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables
    Call InitData

    Call SetToolbar("1110111100010111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    Exit Function
       
	If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow, "I"

			.vspdData.Col = C_W9
			.vspdData.Text = ""
    
			.vspdData.Col = C_W10
			.vspdData.Text = ""
			
			.vspdData.Col = C_W11
			.vspdData.Text = ""
			
			.vspdData.Col = C_W12
			.vspdData.Text = ""
			
			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows

	Select Case lgCurrGrid 
		CAse  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 
				If CheckTotalRow(frm1.vspdData, .ActiveRow,C_SEQ_NO) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				ElseIf CheckDetailData(Frm1.vspdData, Frm1.vspdData2, .ActiveRow) > 0 Then
					MsgBox "하위 데이타가 존재하여 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.EditUndo
				End If
				Call SetG1W4_W5(C_W2, lDelRows)
					'lgCurrGrid=3
					'Call fncCancel()
					ggoSpread.Source = frm1.vspdData3
					lDelRows = ggoSpread.EditUndo
					
					If .MaxRows = 1 Then
						ggoSpread.Source = frm1.vspdData
						lDelRows = ggoSpread.EditUndo
					End If
			End With
		CAse 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True And CheckDetailData(Frm1.vspdData2, Frm1.vspdData2, .ActiveRow) > 1 Then
					MsgBox "다른 행이 존재해 합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				ElseIf CheckDetailData(Frm1.vspdData2, Frm1.vspdData2, .ActiveRow) = 2 Then
					lDelRows = ggoSpread.EditUndo
					Call SetGrid2(C_W6, lDelRows)
					lDelRows = ggoSpread.EditUndo
				Else
					lDelRows = ggoSpread.EditUndo
				End If
				Call SetGrid2(C_W6, lDelRows)
			End With    
 		CAse 3
 			Exit Function ' -- 3번 그리드는 삭제할 수 없다.
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If .MaxRows <= 0 Then
					Exit Function
				ElseIf .ActiveRow <= 4 Then
					Exit Function
				ElseIf CheckTotalRow(frm1.vspdData3, .ActiveRow,C_SEQ_NO2) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				'Else
				'	lDelRows = ggoSpread.EditUndo
				End If
				'Call SetGrid3(C_W11, lDelRows, "")
			End With 
			lgCurrGrid=1
	End Select
  
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, iLastRow
    Dim iStrNm, iG1R2

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
   
	Frm1.vspdData.Col = C_W2
	Frm1.vspdData.Row = Frm1.vspdData.ActiveRow
   	iStrNm = Frm1.vspdData.Text
	Frm1.vspdData.Col = C_R2
   	iG1R2 = Frm1.vspdData.Text

	With frm1	

		Select Case lgCurrGrid
			Case 1	' 1번 그리드 
		
			' 첫행일 경우 합계까지 넣는 루틴 
			If .vspdData.MaxRows = 0 Then
				Call InsertRow2Head
				Call SetToolbar("1110111100010111")
				Exit Function
			End If
		
			.vspdData.focus
			ggoSpread.Source = .vspdData
			
			iRow = .vspdData.ActiveRow	' 현재행 
			
			.vspdData.ReDraw = False
			
			If iRow = .vspdData.MaxRows Then
		
				' SEQ_NO 를 그리드에 넣는 로직 
				iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
				ggoSpread.InsertRow iRow-1 ,imRow	' 그리드 행 추가(사용자 행수 포함)
				SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1, "I"	' 그리드 색상변경 
				Call InsertRowHead3(iRow -1, iSeqNo)
		
				For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
					.vspdData.Row = iRow
					.vspdData.Col = C_SEQ_NO
					.vspdData.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.

				Next
				
			Else

				' SEQ_NO 를 그리드에 넣는 로직 
				iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
				ggoSpread.InsertRow ,imRow	' 그리드 행 추가(사용자 행수 포함)
				SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1, "I"	' 그리드 색상변경 
				Call InsertRowHead3(iRow, iSeqNo)
		
				For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
					.vspdData.Row = iRow
					.vspdData.Col = C_SEQ_NO
					.vspdData.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.

				Next			
			End If

			.vspdData.ReDraw = True	
						
			' 하위 그리드 표시루틴'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			
			frm1.vspdData2.SetActiveCell C_W6, iLastRow-1
			
			Call vspdData_Click(.vspdData.Col, .vspdData.ActiveRow)
			
		Case 2	' 2번 그리드 
			.vspdData2.focus
			ggoSpread.Source = .vspdData2		
					
			If iG1R2="1" Then Exit Function 
			
			.vspdData.Col = C_SEQ_NO : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = .vspdData.value

			' 첫행일 경우 합계까지 넣는 루틴 
			If .vspdData.MaxRows =  0 Then
				Exit Function
			ElseIf .vspdData.ActiveRow = .vspdData.MaxRows Then
				Exit Function
			ElseIf ShowRowHidden(frm1.vspdData2, iSeqNo) > 0 Then
				Call InsertRow2Detail2(iSeqNo)
				'Call ShowRowHidden(frm1.vspdData2, iSeqNo) If 절에서 수행하면 해당 iSeqNo의 행들이 보이게 된다.
			ElseIf CheckDetailData(frm1.vspdData2, frm1.vspdData2, .vspdData2.ActiveRow) = 0 Then
				iRow = .vspdData2.ActiveRow
				.vspdData2.Row = iRow	:	.vspdData2.Col= C_CHILD_SEQ_NO
				ggoSpread.InsertRow ,imRow
				MaxSpreadVal2 .vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
				SetSpreadColorDetail2 iRow+1,iRow+1, "I",iG1R2

				iRow = iRow + 1
				ggoSpread.InsertRow ,1
				.vspdData2.Row = iRow+1
				.vspdData2.Col = C_CHILD_SEQ_NO	: .vspdData2.Text = "999999"
				.vspdData2.Col = C_SEQ_NO			: .vspdData2.Text = iSeqNo
				.vspdData2.Col = C_W6				: .vspdData2.CellType = 1 : .vspdData2.text = "계" : .vspdData2.TypeHAlign = 2
				
				ggoSpread.SpreadLock C_W6, iRow+1, C_W10, iRow+1
				'ggoSpread.SpreadRequired C_W9, iRow+1, C_W9
				'SetSpreadColorDetail2 iRow+1,iRow+1, "I",iG1R2

				'For iRow = 1 to .vspdData2.MaxRows	' 추가된 그리드의 SEQ_NO를 변경한다.
				'	.vspdData2.Row = iRow
				'	.vspdData2.Col = C_SEQ_NO
				'	If .vspdData2.Text = iSeqNo Then
				'		.vspdData2.Col = C_W9	:	.vspdData2.Text = 0
				'		.vspdData2.Col = C_W10	:	.vspdData2.Text = 0
				'	End If
				'
				'Next			
				
				
			Else
				'.vspdData2.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
				iRow = .vspdData2.ActiveRow
				.vspdData2.Row = iRow	:	.vspdData2.Col= C_CHILD_SEQ_NO

'				If iRow = .vspdData2.MaxRows Then
				If .vspdData2.Text = "999999" Then
					ggoSpread.InsertRow iRow-1 , imRow 
					MaxSpreadVal2 .vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow , iSeqNo
					SetSpreadColorDetail2 iRow,iRow, "I",iG1R2
				Else
					ggoSpread.InsertRow ,imRow
					MaxSpreadVal2 .vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
					SetSpreadColorDetail2 iRow+1,iRow+1, "I",iG1R2
				End If	
			End If

		Case 3	' 3번 그리드 
			' 첫행일 경우 합계까지 넣는 루틴 
'			If .vspdData3.MaxRows = 4 Then
'				Call InsertRow2Head3
'				Call SetToolbar("1110111100011111")
				Exit Function
'			End If

			.vspdData3.focus
			ggoSpread.Source = .vspdData3

			iRow = .vspdData3.ActiveRow	' 현재행 
			
			.vspdData3.ReDraw = False
			
			If iRow = .vspdData3.MaxRows Then
		
				' SEQ_NO 를 그리드에 넣는 로직 
				iSeqNo = GetMaxSpreadVal(.vspdData3 , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
				ggoSpread.InsertRow iRow-1 ,imRow	' 그리드 행 추가(사용자 행수 포함)
'				SetSpreadColor3 .vspdData3.ActiveRow, .vspdData3.ActiveRow + imRow - 1	' 그리드 색상변경 
		
				For iRow = .vspdData3.ActiveRow to .vspdData3.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
					.vspdData3.Row = iRow
					.vspdData3.Col = C_SEQ_NO
					.vspdData3.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.

				Next				
			ElseIf iRow < 4 Then
				' SEQ_NO 를 그리드에 넣는 로직 
				iSeqNo = GetMaxSpreadVal(.vspdData3 , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
				ggoSpread.InsertRow iRow-1 ,imRow	' 그리드 행 추가(사용자 행수 포함)
'				SetSpreadColor3 .vspdData3.ActiveRow, .vspdData3.ActiveRow + imRow - 1	' 그리드 색상변경 
		
				For iRow = .vspdData3.ActiveRow to .vspdData3.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
					.vspdData3.Row = iRow
					.vspdData3.Col = C_SEQ_NO
					.vspdData3.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.

				Next				
			Else

				' SEQ_NO 를 그리드에 넣는 로직 
				iSeqNo = GetMaxSpreadVal(.vspdData3 , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
				ggoSpread.InsertRow ,imRow	' 그리드 행 추가(사용자 행수 포함)
'				SetSpreadColor3 .vspdData3.ActiveRow, .vspdData3.ActiveRow + imRow - 1	' 그리드 색상변경 
		
				For iRow = .vspdData3.ActiveRow to .vspdData3.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
					.vspdData3.Row = iRow
					.vspdData3.Col = C_SEQ_NO
					.vspdData3.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.

				Next			
			End If

			.vspdData3.ReDraw = True	
						
			Call vspdData3_Click(.vspdData3.Col, .vspdData3.ActiveRow)

		End Select
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

	Select Case lgCurrGrid 
		Case  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 
				If CheckTotalRow(frm1.vspdData, .ActiveRow,C_SEQ_NO) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				ElseIf CheckDetailData(Frm1.vspdData, Frm1.vspdData2, .ActiveRow) > 0 Then
					MsgBox "하위 데이타가 존재하여 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				Call SetG1W4_W5(C_W2, lDelRows)
				
			End With
		Case 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True And CheckDetailData(Frm1.vspdData2, Frm1.vspdData2, .ActiveRow) > 1 Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					frm1.vspdData.Col = C_R2	:	frm1.vspdData.Row = frm1.vspdData.ActiveRow			
					If frm1.vspdData.Text="1" Then Exit Function 
					lDelRows = ggoSpread.DeleteRow
				End If
				Call SetGrid2(C_W6, lDelRows)
			End With    
 		Case 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If .MaxRows <= 0 Then
					Exit Function
				ElseIf .ActiveRow <= 4 Then
					Exit Function
				ElseIf CheckTotalRow(frm1.vspdData3, .ActiveRow,C_SEQ_NO2) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				Call SetGrid3(C_W11, lDelRows, "")
			End With     
	End Select
	
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
    Dim IntRetCD , blnChange

	FncExit = False

  '-----------------------
    'Check previous data area
    '-----------------------
	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData2
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData3
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
	
    If lgBlnFlgChgValue Then
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
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr, iSeqNo, iLastRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
        
    Call SetToolbar("1111111100010111")										<%'버튼 툴바 제어 %>
	
	Call RedrawSumRow("Q")
	Call RedrawSumRow2("Q")
	Call RedrawSumRow3("Q")

	With frm1.vspdData
		.Col = C_SEQ_NO : .Row = .ActiveRow : iSeqNo = .Value
			
		' 하위 그리드 표시루틴'
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)

	End With		
	frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow , lCol, lGrpCnt, lMaxRows, lMaxCols
    Dim lStartRow, lEndRow , lChkAmt
    Dim strVal
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    lGrpCnt = 1
    
	With frm1.vspdData
		' ----- 1번째 그리드 
		ggoSpread.Source = frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0
		   
		   ' I/U/D 플래그 처리 
		   Select Case .Text
		       Case  ggoSpread.InsertFlag                                      '☜: Insert
		                                          strVal = strVal & "C"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
		                                           strVal = strVal & "U"  &  Parent.gColSep                                                 
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '☜: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		       
		  End Select
		 
	  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  

		Next
	End With

    frm1.txtSpread.value      = strVal    
    strVal = ""

 	With frm1.vspdData2
		' ----- 2번째 그리드 
		ggoSpread.Source = frm1.vspdData2
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0

		   ' I/U/D 플래그 처리 
		   Select Case .Text
		       Case  ggoSpread.InsertFlag                                      '☜: Insert
                                          strVal = strVal & "C"  &  Parent.gColSep	
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
                                          strVal = strVal & "U"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '☜: Delete
											strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With

    frm1.txtSpread2.value      = strVal
    strVal = ""
    	
	With frm1.vspdData3
		' ----- 3번째 그리드 
		ggoSpread.Source = frm1.vspdData3
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0
		   
		   ' I/U/D 플래그 처리 
		   Select Case .Text
		       Case  ggoSpread.InsertFlag                                      '☜: Insert
		                                          strVal = strVal & "C"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
		                                           strVal = strVal & "U"  &  Parent.gColSep                                                 
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '☜: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_W_TYPE To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With
	
    frm1.txtSpread3.value      = strVal
    strVal = ""	  
        
	frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
	frm1.txtFlgMode.value     = lgIntFlgMode


	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
	frm1.vspdData3.MaxRows = 0
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    	
    Call MainQuery()
End Function

Function DBSaveFail()													        <%' Save Failed %>

	'frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2    
	Call DisplayMsgBox("W30011", "X", "X", "X")     
    
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

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="No">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">금액불러오기</A>|<a href="vbscript:OpenRefMenu">소득금액합계표 조회</A>  
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=1>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 1. 가지급금등 인정이자 조정</LEGEND>
								       <TABLE CLASS="BasicTB" CELLSPACING=0>
								           <TR>				
								        	   <TD>
				                                   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=170 tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								        	   </TD>
								        	   <TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=170 tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								        	   </TD>
								           </TR>
								       </TABLE>				
								  </FIELDSET>
								  <BR>
									<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 2. 이자율별 차입금 적수계산</LEGEND>
								       <TABLE CLASS="BasicTB" CELLSPACING=0>
								           <TR>				
								        	   <TD WIDTH="100%">
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=150 tag="23" TITLE="SPREAD" id=vaSpread3> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								        	   </TD>
								           </TR>
								       </TABLE>				
								  </FIELDSET>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			    <TR>
				        <TD WIDTH=10>&nbsp;</TD>
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>가지급금등 인정이자 조정</LABEL>&nbsp;
				            <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><별지>이자율 차입금 적수계산</LABEL>&nbsp;
				        </TD>
				                                 
	
                </TR>
			
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" STYLE="Display:none"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" STYLE="Display:none"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24" STYLE="Display:none"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_CO_CD" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_FISC_YEAR" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_REP_TYPE" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
