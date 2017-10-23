<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기타 서식 
'*  3. Program ID           : w9109ma1
'*  4. Program Name         : w9109ma1.asp
'*  5. Program Desc         : 제 54호 주식변동상황명세서(갑)
'*  6. Modified date(First) : 2005/02/02
'*  7. Modified date(Last)  : 2006/02/03
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : HJO 
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

Const BIZ_MNU_ID = "w9109ma1"
Const BIZ_PGM_ID = "w9109mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W9109OA1"
Const TYPE_1	= 0
Const TYPE_2_1	= 1
Const TYPE_2_2	= 2
Const TYPE_3	= 3

Dim lgChkFlag  'checking validation of input data

Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6_1
Dim C_W6_2
Dim C_W89

Dim C_SEQ_NO
Dim C_W7
Dim C_W8
Dim C_W8_P
Dim C_W8_NM
Dim C_W9
Dim C_W9_P
Dim C_W9_NM
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13

Dim C_W16
Dim C_W17_1
Dim C_W17
Dim C_W17_P
Dim C_W17_NM
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W21
Dim C_W22
Dim C_W23
Dim C_W24
Dim C_W25
Dim C_W26
Dim C_W27
Dim C_W28
Dim C_W29
Dim C_W30
Dim C_W31
Dim C_W32
Dim C_W33
Dim C_W34
Dim C_W35
Dim C_W36
Dim C_W36_P
Dim C_W36_NM

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim	IsRunEvents, lgvspdData(3)

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	
	' 그리드1
	C_W1		= 0
	C_W2		= 1
	C_W3		= 2
	C_W4		= 3
	C_W5		= 4
	C_W6_1		= 5
	C_W6_2		= 6
	C_W89		= 7
	
	C_SEQ_NO	= 1
	C_W7		= 2
	C_W8		= 3
	C_W8_P		= 4
	C_W8_NM		= 5
	C_W9		= 6
	C_W9_P		= 7
	C_W9_NM		= 8
	C_W10		= 9
	C_W11		= 10
	C_W12		= 11
	C_W13		= 12
	
	C_W16		= 1
	C_W17_1		= 2
	C_W17		= 3
	C_W17_P		= 4
	C_W17_NM	= 5
	C_W18		= 6
	C_W19		= 7
	C_W20		= 8
	C_W21		= 9
	C_W22		= 10
	C_W23		= 11
	C_W24		= 12
	C_W25		= 13
	C_W26		= 14
	C_W27		= 15
	C_W28		= 16
	C_W29		= 17
	C_W30		= 18
	C_W31		= 19
	C_W32		= 20
	C_W33		= 21
	C_W34		= 22
	C_W35		= 23
	C_W36		= 24
	C_W36_P		= 25
	C_W36_NM	= 26
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

	IsRunEvents = False
	lgCurrGrid = TYPE_1
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
	Dim ret, iType
	
	Set lgvspdData(TYPE_1)		= frm1.txtData
	Set lgvspdData(TYPE_2_1)	= frm1.vspdData0
	Set lgvspdData(TYPE_2_2)	= frm1.vspdData1
	Set lgvspdData(TYPE_3)		= frm1.vspdData2
	
	Call initSpreadPosVariables()  
	
	Call AppendNumberPlace("6","3","2")

	' -- 변동상황 그리드(자본금)
	For iType = TYPE_2_1 To TYPE_2_2
		With lgvspdData(iType)
			
			ggoSpread.Source = lgvspdData(iType)	
			'patch version
			ggoSpread.Spreadinit "V20041222" & iType ,,parent.gForbidDragDropSpread    
			
			.ReDraw = false
				 
			.MaxCols = C_W13 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
			.Col = .MaxCols														'☆: 사용자 별 Hidden Column
			.ColHidden = True    

  			'헤더를 3줄로    
			.ColHeaderRows = 2  
							       
			.MaxRows = 0
			ggoSpread.ClearSpreadData

			ggoSpread.SSSetEdit		C_SEQ_NO,	"순번", 10,,,10,1
			ggoSpread.SSSetDate		C_W7,		"(7)일자"	, 10, 2, Parent.gDateFormat, -1
			ggoSpread.SSSetEdit		C_W8,		"(8)원인코드", 6,,,10,1
			ggoSpread.SSSetButton	C_W8_P	
			ggoSpread.SSSetEdit		C_W8_NM,		"원인명", 10,,,30,1
			ggoSpread.SSSetEdit		C_W9,		"(9)종류" , 6,,,10,1
			ggoSpread.SSSetButton	C_W9_P	
			ggoSpread.SSSetEdit		C_W9_NM,		"종류명" , 10,,,30,1
			ggoSpread.SSSetFloat	C_W10,		"(10)주식수" & vbCrLf & "(출자좌수)" , 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
			ggoSpread.SSSetFloat	C_W11,		"(11)주당" & vbCrLf & "액면가액", 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
			ggoSpread.SSSetFloat	C_W12,		"(12)주당발행" & vbCrLf & "(인수)가액", 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
			ggoSpread.SSSetFloat	C_W13,		"(13)증가(감소)" & vbCrLf & "자본금" , 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  


			ret = .AddCellSpan(C_SEQ_NO	, -1000, 1, 2)	
			ret = .AddCellSpan(C_W7		, -1000, 1, 2)	
			ret = .AddCellSpan(C_W8		, -1000, 1, 2)	
			ret = .AddCellSpan(C_W8_P	, -1000, 1, 2)
			ret = .AddCellSpan(C_W8_NM	, -1000, 1, 2)		
			ret = .AddCellSpan(C_W9		, -1000, 6, 1)	
			ret = .AddCellSpan(C_W13	, -1000, 1, 2)	

			' 첫번째 헤더 출력 글자 
			.Row = -1000
			.Col = C_W7		: .Text = "(7)일자"
			.Col = C_W8		: .Text = "(8)원인코드"
			.Col = C_W8_NM	: .Text = "원인명"
			.Col = C_W9		: .Text = "증가(감소)한 주식의 내용"
			.Col = C_W13	: .Text = "(13)증가(감소)" & vbCrLf & "자본금"
			
			.Row = -999
			.Col = C_W9		: .Text = "(9)종류"
			.Col = C_W9_NM	: .Text = "종류명"
			.Col = C_W10	: .Text = "(10)주식수" & vbCrLf & "(출자좌수)"
			.Col = C_W11	: .Text = "(11)주당" & vbCrLf & "액면가액"
			.Col = C_W12	: .Text = "(12)주당발행" & vbCrLf & "(인수)가액"
					
			.rowheight(-1000) = 10					
			.rowheight(-999) = 20					
				
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_SEQ_NO, C_SEQ_NO, True)
			
			.ReDraw = true

			.ScriptEnhanced = True
    
			.TextTip = 1
			.TextTipDelay = 10  ' Control displays text tips after 250 milliseconds
			ret = .SetTextTipAppearance("굴림체", "9", False, False, &HD2F0E1, &H800000)			

		End With
	Next
	
	' -- 변동상황 그리드(주식수/출자좌수)
	With lgvspdData(TYPE_3)

		ggoSpread.Source = lgvspdData(TYPE_3)	
		'patch version
		ggoSpread.Spreadinit "V20041222" & TYPE_3 ,,parent.gForbidDragDropSpread    

		.ReDraw = false
					 
		.MaxCols = C_W36_NM + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    

  		'헤더를 3줄로    
		.ColHeaderRows = 3  
								       
		.MaxRows = 0
		ggoSpread.ClearSpreadData
			
		ggoSpread.SSSetEdit		C_W16,		"(16)" & vbCrLf & "일" & vbCrLf & "련" & vbCrLf & "번" & vbCrLf & "호", 5,,,10,2
		ggoSpread.SSSetEdit		C_W17_1,	"구분", 5,,,5,1
		ggoSpread.SSSetEdit		C_W17,		"(17)구분", 10,,,50,1
		ggoSpread.SSSetButton	C_W17_P	
		ggoSpread.SSSetEdit		C_W17_NM,	"구분명", 10,,,50,1
		ggoSpread.SSSetEdit		C_W18,		"(18)성명(법인명)" , 7,,,50,1
		ggoSpread.SSSetEdit		C_W19,		"(19)주민등록번호" & vbCrLf & "(사업자등록번호)" , 10,,,20,1
		ggoSpread.SSSetFloat	C_W20,		"(20)주식수" & vbCrLf & "(좌수)", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W21,		"(21)지분율" , 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W22,		"(22)양수", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W23,		"(23)유상증자", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W24,		"(24)무상증자" , 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W25,		"(25)상속", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W26,		"(26)증여", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W27,		"(27)전환사채등" & vbCrLf & "출자전환" , 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W28,		"(28)기타", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W29,		"(29)양도", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W30,		"(30)상속" , 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W31,		"(31)증여", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W32,		"(32)감자", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W33,		"(33)기타" , 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W34,		"(34)주식수" & vbCrLf & "(좌수)", 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W35,		"(35)지분율" , 7,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetEdit		C_W36,		"(36)대주주와의" & vbCrLf & "관계코드", 7,,,10,1
		ggoSpread.SSSetButton	C_W36_P
		ggoSpread.SSSetEdit		C_W36_NM,	"관계명", 10,,,50,1

		ret = .AddCellSpan(C_W16	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W17_1	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W17	, -1000, 5, 2)	
		ret = .AddCellSpan(C_W20	, -1000, 2, 2)	
		ret = .AddCellSpan(C_W22	, -1000,14, 1)	
		ret = .AddCellSpan(C_W36	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W36_P	, -1000, 1, 3)
		ret = .AddCellSpan(C_W36_NM	, -1000, 1, 3)		
		
		ret = .AddCellSpan(C_W22	,  -999, 7, 1)	
		ret = .AddCellSpan(C_W29	,  -999, 5, 1)	


		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W16	: .Text = "(16)" & vbCrLf & "일" & vbCrLf & "련" & vbCrLf & "번" & vbCrLf & "호"
		.Col = C_W17	: .Text = "주주.출자자"
		.Col = C_W20	: .Text = "기 초"
		.Col = C_W22	: .Text = "변동상황 (주식수.출자좌수)"
		.Col = C_W34	: .Text = "기 말"
		.Col = C_W36	: .Text = "(36)대주주와의" & vbCrLf & "관계코드"
		
		.Row = -999
		.Col = C_W22	: .Text = "증가주식수(출자좌수)"
		.Col = C_W29	: .Text = "감소주식수(출자좌수)"
		
		.Row = -998
		.Col = C_W17	: .Text = "(17)구분"
		.Col = C_W17_NM	: .Text = "구분명"
		.Col = C_W18	: .Text = "(18)성명(법인명)"
		.Col = C_W19	: .Text = "(19)주민등록번호(사업자등록번호)"
		.Col = C_W20	: .Text = "(20)주식수(좌수)"
		.Col = C_W21	: .Text = "(21)지분율"
		.Col = C_W22	: .Text = "(22)양수"
		.Col = C_W23	: .Text = "(23)유상증자"
		.Col = C_W24	: .Text = "(24)무상증자"
		.Col = C_W25	: .Text = "(25)상속"
		.Col = C_W26	: .Text = "(26)증여"
		.Col = C_W27	: .Text = "(27)전환사채등출자전환"
		.Col = C_W28	: .Text = "(28)기타"
		.Col = C_W29	: .Text = "(29)양도"
		.Col = C_W30	: .Text = "(30)상속"
		.Col = C_W31	: .Text = "(31)증여"
		.Col = C_W32	: .Text = "(32)감자"
		.Col = C_W33	: .Text = "(33)기타"
		.Col = C_W34	: .Text = "(34)주식수(좌수)"
		.Col = C_W35	: .Text = "(35)지분율"
		.Col = C_W36	: .Text = "(36)대주주와의관계코드"
		.Col = C_W36_NM	: .Text = "관계명"

		.rowheight(-1000) = 10					
		.rowheight(-999) = 10					
		.rowheight(-998) = 30		
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W17_1, C_W17_1, True)
		
		Call MakePercentCol( lgvspdData(TYPE_3), C_W21, "0", "", "")
		Call MakePercentCol( lgvspdData(TYPE_3), C_W35, "0", "", "")
		
		Call ggoSpread.SSSetSplit2(C_W19)

		.ScriptEnhanced = True
    
		.TextTip = 1
		.TextTipDelay = 10  ' Control displays text tips after 250 milliseconds
		ret = .SetTextTipAppearance("굴림체", "9", False, False, &HD2F0E1, &H800000)			
		
	End With			
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)
	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)
	
		Select Case pType
			Case TYPE_2_1
				ggoSpread.SpreadLock C_W7,  1, C_W9_P, 2
				ggoSpread.SpreadLock C_W12,  1, C_W12, 2
				ggoSpread.SpreadLock C_W8_NM,  1, C_W8_NM, .MaxRows
				ggoSpread.SpreadLock C_W9_NM,  1, C_W9_NM, .MaxRows
				'ggoSpread.SSSetRequired  C_W7, 3, 3
				'ggoSpread.SSSetRequired  C_W8, 3, 3
				'ggoSpread.SSSetRequired  C_W9, 3, 3
				'ggoSpread.SSSetRequired  C_W10, 3, 3
				'ggoSpread.SSSetRequired  C_W11, 3, 3
				'ggoSpread.SSSetRequired  C_W12, 3, 3
				'ggoSpread.SSSetRequired  C_W13, 3, 3
				ggoSpread.SpreadLock C_W13,  -1, C_W13
			Case TYPE_2_2
				ggoSpread.SpreadLock C_W7,  6, C_W9_P, 7
				ggoSpread.SpreadLock C_W12,  6, C_W12, 7
				ggoSpread.SpreadLock C_W8_NM,  1, C_W8_NM, .MaxRows
				ggoSpread.SpreadLock C_W9_NM,  1, C_W9_NM, .MaxRows
				ggoSpread.SpreadLock C_W11,  6, C_W11, 7
				ggoSpread.SpreadLock C_W13,  -1, C_W13
			Case TYPE_3
				ggoSpread.SpreadLock C_W16,  1, C_W36_P, 1
				ggoSpread.SpreadLock C_W16,  -1, C_W16
				ggoSpread.SSSetRequired  C_W17, 3, .MaxRows
				ggoSpread.SpreadLock C_W17_NM,  -1, C_W17_NM
				ggoSpread.SSSetRequired  C_W36, 3, .MaxRows
				ggoSpread.SpreadLock C_W17,  1, C_W17_P, 2
				ggoSpread.SpreadLock C_W34,  -1, C_W35
				ggoSpread.SpreadLock C_W21,  -1, C_W21
				ggoSpread.SpreadLock C_W36_NM,  -1, C_W36_NM
				
		End Select

	End With
End Sub

Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
    With lgvspdData(pType)
	ggoSpread.Source = lgvspdData(pType)
	Select Case pType
		Case TYPE_3
			ggoSpread.SSSetProtected C_W16, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired  C_W17, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W17_NM, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired  C_W36, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W34, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W35, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W21, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W36_NM, pvStartRow, pvEndRow
			
	End Select
	    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = lgvspdData(TYPE_2_1)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_SEQ_NO	= iCurColumnPos(1)
			C_W7		= iCurColumnPos(2)
			C_W8		= iCurColumnPos(3)
			C_W8_P		= iCurColumnPos(4)
			C_W8_NM		= iCurColumnPos(5)
			C_W9		= iCurColumnPos(6)
			C_W9_P		= iCurColumnPos(7)
			C_W9_NM		= iCurColumnPos(8)
			C_W10		= iCurColumnPos(9)
			C_W11		= iCurColumnPos(10)
			C_W12		= iCurColumnPos(11)
			C_W13		= iCurColumnPos(12)
		Case "B"
            ggoSpread.Source = lgvspdData(TYPE_2_2)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_SEQ_NO	= iCurColumnPos(1)
			C_W7		= iCurColumnPos(2)
			C_W8		= iCurColumnPos(3)
			C_W8_P		= iCurColumnPos(4)
			C_W8_NM		= iCurColumnPos(5)
			C_W9		= iCurColumnPos(6)
			C_W9_P		= iCurColumnPos(7)
			C_W9_NM		= iCurColumnPos(8)
			C_W10		= iCurColumnPos(9)
			C_W11		= iCurColumnPos(10)
			C_W12		= iCurColumnPos(11)
			C_W13		= iCurColumnPos(12)
		Case "C"
            ggoSpread.Source = lgvspdData(TYPE_3)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_W16		= iCurColumnPos(1)
			C_W17_1		= iCurColumnPos(2)
			C_W17		= iCurColumnPos(3)
			C_W17_P		= iCurColumnPos(4)
			C_W17_NM	= iCurColumnPos(5)
			C_W18		= iCurColumnPos(6)
			C_W19		= iCurColumnPos(7)
			C_W20		= iCurColumnPos(8)
			C_W21		= iCurColumnPos(9)
			C_W22		= iCurColumnPos(10)
			C_W23		= iCurColumnPos(11)
			C_W24		= iCurColumnPos(12)
			C_W25		= iCurColumnPos(13)
			C_W26		= iCurColumnPos(14)
			C_W27		= iCurColumnPos(15)
			C_W28		= iCurColumnPos(16)
			C_W29		= iCurColumnPos(17)
			C_W30		= iCurColumnPos(18)
			C_W31		= iCurColumnPos(19)
			C_W32		= iCurColumnPos(20)
			C_W33		= iCurColumnPos(21)
			C_W34		= iCurColumnPos(22)
			C_W35		= iCurColumnPos(23)
			C_W36		= iCurColumnPos(24)
			C_W36_P		= iCurColumnPos(25)
			C_W36_NM	= iCurColumnPos(26)
    End Select    
End Sub

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Dim iType, iRow, iMaxRows, iStartRow
	iMaxRows = 7 : iStartRow = 1
	
	' 그리드2 추가 
	For iType = TYPE_2_1 To TYPE_2_2
		With lgvspdData(iType)
			ggoSpread.Source = lgvspdData(iType)
			ggoSpread.InsertRow , iMaxRows
			Call SetSpreadLock(iType) 
		
			For iRow = 1 To 7
				.Col = C_SEQ_NO
				.Row = iRow
				.Value = iStartRow
				iStartRow = iStartRow + 1
			Next
						
			Select Case iType
				Case TYPE_2_1
					Call ReTypeGrid(iType, 1)
				Case TYPE_2_2
					Call ReTypeGrid(iType, 6)
			End Select
		End With
	Next
	
	Call GetRef
	lgBlnFlgChgValue = False
	
End Sub

Sub ReTypeGrid(Byval pType, Byval pRow)
	Dim ret, iRow, iStartRow
	
	With lgvspdData(pType)
		' -- 일자/원인 
		.BlockMode = True
		.Col = C_W7		: .Row = pRow
		.Col2= C_W9_P	: .Row2 = pRow + 1
		.CellType = 1
		.BlockMode = False

		.Col = C_W9	
		.Row = pRow		: .Text = "01"
		.Row = pRow+1	: .Text = "02"
						
		ret = .AddCellSpan(C_W7	,  pRow, 4, 2)	
		.Col = C_W7	: .Row = pRow
		.TypeEditMultiLine = True
		.TypeHAlign = 2	: .TypeVAlign = 2
		
		Select Case pType
			Case TYPE_2_1
				iStartRow = 0
				
				.Text = "(14) 기 초"
				' (11)주당액면가액 합치기 
				ret = .AddCellSpan(C_W11	,  pRow, 1, 2)	
				.Col = C_W11	: .Row = pRow
				.TypeVAlign = 2
				ret = .AddCellSpan(C_W13	,  pRow, 1, 2)	
				.Col = C_W13	: .Row = pRow
				.TypeVAlign = 2

			Case TYPE_2_2
				iStartRow = 7

				.Text = "(15) 기 말"
				' (13)증가(감소)자본금 합치기 
				ret = .AddCellSpan(C_W13	,  pRow, 1, 2)	
				.Col = C_W13	: .Row = pRow
				.TypeVAlign = 2
				ret = .AddCellSpan(C_W11	,  pRow, 1, 2)	
				.Col = C_W11	: .Row = pRow
				.TypeVAlign = 2
							
		End Select

	End With	
End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 그리드1의 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2, iMaxRows, sTmp, iRow, arrADDR
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
'    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

'	sMesg = wgRefDoc & vbCrLf & vbCrLf

'	IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"

'	If IntRetCD = vbNo Then
'		 Exit Function
'	End If
	Call ggoOper.FormatDate(frm1.txtData(C_W6_1), parent.gDateFormat,1)
	Call ggoOper.FormatDate(frm1.txtData(C_W6_2), parent.gDateFormat,1)
	
    Dim IntRetCD1

	IntRetCD = CommonQueryRs("W1, W2"," dbo.ufn_TB_54_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		iMaxRows	= UBound(arrW1)

		With frm1
		
		For iRow = 0 To iMaxRows -1
			Select Case arrW1(iRow) 
				Case "10"
					If arrW2(iRow) <> "" Then Call txtW10_Click(CDbl(arrW2(iRow))-1)
				Case "6_1"				
					frm1.txtData(C_W6_1).text=arrW2(iRow)
				Case "6_2"
					frm1.txtData(C_W6_2).text=arrW2(iRow)  
				Case Else
					sTmp = "frm1.txtData(C_W" & arrW1(iRow) & ").Value = """ & CStr(arrW2(iRow)) & """"	
					Execute sTmp	' -- 변수에 들어 있는 명령을 실행한다. 
			End Select
		Next
		
		End With
		
		'Call SetReCalc1
	End If

	lgBlnFlgChgValue = True
End Function


' 해당 그리드에서 데이타가져오기 
Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : GetGrid = .text
	End With
End Function

' 해당 그리드에서 데이타가져오기 
Function PutGrid(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : .text = pVal
	End With
End Function

'============================================  그리드 팝업  ====================================

Function OpenW1082(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "원인 코드"						' 팝업 명칭 
	arrParam(1) = "B_MINOR"						' TABLE 명칭 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD='W1082'"								' Where Condition
	arrParam(5) = "원인 코드"

    arrField(0) = "MINOR_CD"					' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)

    arrHeader(0) = "원인 코드"						' Header명(0)
    arrHeader(1) = "원인 명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		lgvspdData(lgCurrGrid).focus
	    Exit Function
	Else
		Call SetW1082(arrRet)
	End If	

End Function

Sub SetW1082(Byref pArrRet)
	With lgvspdData(lgCurrGrid)
		.Row = .ActiveRow
		.Col = C_W8
		.Value = pArrRet(0)
		.Col = C_W8_NM
		.Value = pArrRet(1)
		
		Call vspdData_Change(lgCurrGrid, .Col, .Row) 
	End With
End Sub

Function OpenW1083(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "종 류"						' 팝업 명칭 
	arrParam(1) = "B_MINOR"						' TABLE 명칭 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD='W1083'"								' Where Condition
	arrParam(5) = "종류 코드"

    arrField(0) = "MINOR_CD"					' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)

    arrHeader(0) = "종류 코드"						' Header명(0)
    arrHeader(1) = "종류 명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		lgvspdData(lgCurrGrid).focus
	    Exit Function
	Else
		Call SetW1083(arrRet)
	End If	

End Function

Sub SetW1083(Byref pArrRet)
	With lgvspdData(lgCurrGrid)
		.Row = .ActiveRow
		.Col = C_W9
		.Value = pArrRet(0)
		.Col = C_W9_NM
		.Value = pArrRet(1)
		
		Call vspdData_Change(lgCurrGrid, .Col, .Row) 
	End With
End Sub

Function OpenW1034(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구 분"						' 팝업 명칭 
	arrParam(1) = "B_MINOR"						' TABLE 명칭 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD='W1034'"								' Where Condition
	arrParam(5) = "구분 코드"

    arrField(0) = "MINOR_CD"					' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)

    arrHeader(0) = "구분 코드"						' Header명(0)
    arrHeader(1) = "구분 명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		lgvspdData(TYPE_3).focus
	    Exit Function
	Else
		Call SetW1034(arrRet)
	End If	

End Function

Sub SetW1034(Byref pArrRet)
	With lgvspdData(TYPE_3)
		.Row = .ActiveRow
		.Col = C_W17
		.Value = pArrRet(0)
		.Col = C_W17_NM
		.Value = pArrRet(1)
		
		Call vspdData_Change(TYPE_3, .Col, .Row) 
	End With
End Sub

Function OpenW1035(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "(53)대주주와의 관계코드"						' 팝업 명칭 
	arrParam(1) = "B_MINOR"						' TABLE 명칭 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD='W1035'"								' Where Condition
	arrParam(5) = "관계 코드"

    arrField(0) = "MINOR_CD"					' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)

    arrHeader(0) = "관계 코드"						' Header명(0)
    arrHeader(1) = "관계 명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		lgvspdData(TYPE_3).focus
	    Exit Function
	Else
		Call SetW1035(arrRet)
	End If	

End Function

Sub SetW1035(Byref pArrRet)
	With lgvspdData(TYPE_3)
		.Row = .ActiveRow
		.Col = C_W36
		.Value = pArrRet(0)
		.Col = C_W36_NM
		.Value = pArrRet(1)
		
		Call vspdData_Change(TYPE_3, .Col, .Row) 
	End With
End Sub

Function CheckMinorCd(Byval pMajorCd, Byval pMinorCd)
	Dim iRet
	
	CheckMinorCd = False
	iRet = CommonQueryRs("MINOR_CD"," dbo.ufn_TB_MINOR('" & pMajorCd & "','" & C_REVISION_YM & "')", " MINOR_CD='" & pMinorCd & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	lgChkFlag=True
	
	If iRet = False Then
		Call DisplayMsgBox("W90001", parent.VB_INFORMATION, pMinorCd, "X")
		lgChkFlag=False 
		Exit Function
	End If
	CheckMinorCd = True
End Function

' 헤더 재계산 
Sub SetHeadReCalc()	
	Dim dblSum, dblData(40)
	
	If IsRunEvents Then Exit Sub	' 아래 .vlaue = 에서 이벤트가 발생해 재귀함수로 가는걸 막는다.
	
	IsRunEvents = True
	
	With frm1
		
	End With

	lgBlnFlgChgValue= True ' 변경여부 
	IsRunEvents = False	' 이벤트 발생금지를 해제함 
End Sub

'============================================  조회조건 함수  ====================================

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitVariables                                                      <%'Initializes local global variables%>
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    
    Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)


	Call InitData()
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
Sub txtData_DblClick(Button)

    If Button = 1 Then
        frm1.txtData.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtData.Focus
    End If
End Sub
Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	Call GetFISC_DATE
End Sub

Sub txtW10_Click(pIdx)
	With frm1
		.txtW10(0).checked = false
		.txtW10(1).checked = false
		.txtW10(2).checked = false
		.txtW10(pIdx).checked = true
		.txtData(C_W10).value = pIdx + 1
	End With
End Sub

Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, ret, datFISC_START_DT, datFISC_END_DT
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	ret = CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If ret Then
		datFISC_START_DT = CDate(lgF0)
		datFISC_END_DT = CDate(lgF1)
		lgMonGap = DateDiff("m", datFISC_START_DT, datFISC_END_DT)+1
	Else
		lgMonGap = 12
	End If
	
	ret = CommonQueryRs("W1"," dbo.ufn_TB_4_GetRate('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If ret Then
		lgW2019 = UNICDbl(lgF0)
	End If
End Sub

'==========================================================================================
' -- 0번 그리드 
Sub vspdData0_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2_1
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData0_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData0_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2_1
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_GotFocus()
	lgCurrGrid = TYPE_2_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData0_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2_1
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData0_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2_1
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData0_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2_1
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData0_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_2_1
	Call vspdData_ButtonClicked( lgCurrGrid, Col, Row, ButtonDown)
End Sub

' -- 1번 그리드 
Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2_2
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2_2
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_2_2
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2_2
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2_2
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_2_2
	Call vspdData_ButtonClicked( lgCurrGrid, Col, Row, ButtonDown)
End Sub

' -- 2번 그리드 
Sub vspdData2_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_3
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_3
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_3
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_GotFocus()
	lgCurrGrid = TYPE_3
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_3
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_3
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_3
	Call vspdData_ButtonClicked( lgCurrGrid, Col, Row, ButtonDown)
End Sub

Sub vspdData0_ScriptTextTipFetch(Col, Row, MultiLine, TipWidth, TipText, ShowTip)
	lgCurrGrid = TYPE_2_1
	Call vspdData_ScriptTextTipFetch(lgCurrGrid, Col, Row, MultiLine, TipWidth, TipText, ShowTip)
End Sub

Sub vspdData1_ScriptTextTipFetch(Col, Row, MultiLine, TipWidth, TipText, ShowTip)
	lgCurrGrid = TYPE_2_2
	Call vspdData_ScriptTextTipFetch(lgCurrGrid, Col, Row, MultiLine, TipWidth, TipText, ShowTip)
End Sub

Sub vspdData2_ScriptTextTipFetch(Col, Row, MultiLine, TipWidth, TipText, ShowTip)
	lgCurrGrid = TYPE_3
	Call vspdData_ScriptTextTipFetch(lgCurrGrid, Col, Row, MultiLine, TipWidth, TipText, ShowTip)
End Sub

Sub vspdData_ScriptTextTipFetch(Index, Col, Row, MultiLine, TipWidth, TipText, ShowTip)
	Dim pVal
	With lgvspdData(Index)
	.Col = Col : .Row = Row : pVal = Trim(.Text)
	
	Select Case Index
		Case TYPE_2_1
				If Col = C_W8 Then
					TipText = GetCodeNm("1", pVal)
					window.status = TipText
				ElseIf Col = C_W9 Then
					TipText = GetCodeNm("2", pVal)
					window.status = TipText
				Else
					TipText = ""
					window.status = ""
				End If
		Case TYPE_2_2
			If Col = C_W8 Then
				TipText = GetCodeNm("1", pVal)
				window.status = TipText
			ElseIf Col = C_W9 Then
				TipText = GetCodeNm("2", pVal)
				window.status = TipText
			Else
				window.status = ""
			End If
		Case TYPE_3
			If Col = C_W17 Then
				TipText = GetCodeNm("3", pVal)
				window.status = TipText
			ElseIf Col = C_W36 Then
				TipText = GetCodeNm("4", pVal)
				window.status = TipText
			Else 
				window.status = ""
			End If
	End Select
	
	ShowTip=true
	
	End With
End Sub

Function GetCodeNm(pType, pCode)
	Select Case pType
		Case "1"	
			Select Case pCode
				Case "01"
					GetCodeNm = "유사증자"
				Case "02"
					GetCodeNm = "무상증자"
				Case "03"
					GetCodeNm = "출자전환"
				Case "04"
					GetCodeNm = "주식배당"
				Case "05"
					GetCodeNm = "유상감자"
				Case "06"
					GetCodeNm = "무상감자"
				Case "07"
					GetCodeNm = "액면분할"
				Case "08"
					GetCodeNm = "주식병합"
				Case "09"
					GetCodeNm = "기타(자사주소각등)"
			End Select
		Case "2"
			Select Case pCode
				Case "01"
					GetCodeNm = "보통주"
				Case "02"
					GetCodeNm = "우선주"
			End Select
		Case "3"	
			Select Case pCode
				Case "01"
					GetCodeNm = "개인"
				Case "02"
					GetCodeNm = "영리내국법인"
				Case "03"
					GetCodeNm = "비영리내국법인"
				Case "04"
					GetCodeNm = "개인단체"
				Case "05"
					GetCodeNm = "외국투자자"
				Case "06"
					GetCodeNm = "외국법인"
			End Select
		Case "4"	
			Select Case pCode
				Case "00"
					GetCodeNm = "본인"
				Case "01"
					GetCodeNm = "배우자"
				Case "02"
					GetCodeNm = "자"
				Case "03"
					GetCodeNm = "부모"
				Case "04"
					GetCodeNm = "형제자매"
				Case "05"
					GetCodeNm = "손"
				Case "06"
					GetCodeNm = "조부모"
				Case "07"
					GetCodeNm = "02_06의 배우자"
				Case "08"
					GetCodeNm = "01~07이외의 친족"
				Case "09"
					GetCodeNm = "기타"
			End Select
	End Select
End Function

'==========================================================================================
Sub vspdData_Change(Byval pType, ByVal Col , ByVal Row )
	Dim dblSum, dblCol, dblW50, dblAmt(40)
	Dim pVal
	
	lgChkFlag=True
	
	With lgvspdData(pType)
		lgBlnFlgChgValue= True ' 변경여부 
		.Row = Row
		.Col = Col
		pVal = .value
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
		  If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
		     .text = .TypeFloatMin
		  End If
		End If
	
		ggoSpread.Source = lgvspdData(pType)
		ggoSpread.UpdateRow Row

		Select Case pType
			Case TYPE_2_1, TYPE_2_2
				.Row = Row : .Col = Col
				Select Case Col
					Case C_W8
						If pType=TYPE_2_2 and  (Row=6 or Row=7) Then 						
						Else
							If CheckMinorCd("W1082", .Text) = False Then
								.Text = ""
								lgChkFlag=False
								Exit Sub
							End If
						End IF
						.Col = C_W8_NM : .value = GetCodeNm("1", pVal)
					Case C_W9 
						If pType=TYPE_2_2 and  (Row=6 or Row=7) Then 						
						Else
							If CheckMinorCd("W1083", .Text) = False Then
								.Text = ""
								lgChkFlag=False
								Exit Sub
							End If
						End If
						.Col = C_W9_NM : .value = GetCodeNm("2", pVal)
						Call ReCalc_W15
					Case C_W10, C_W11
						If pType = TYPE_2_1 And (Row = 1 Or Row = 2) Then	' -- 기초 
							dblSum = UNICDbl(GetGrid(TYPE_2_1, C_W10, 1)) * UNICDbl(GetGrid(TYPE_2_1, C_W11, 1)) 
							dblSum = dblSum + ( UNICDbl(GetGrid(TYPE_2_1, C_W10, 2)) * UNICDbl(GetGrid(TYPE_2_1, C_W11, 1)) )
							Call PutGrid(TYPE_2_1, C_W13, 1, dblSum)
						ElseIf pType = TYPE_2_2 And (Row = 6 Or Row = 7) Then ' -- 기말 
							dblSum = UNICDbl(GetGrid(TYPE_2_2, C_W10, 6)) * UNICDbl(GetGrid(TYPE_2_2, C_W11, 6)) 
							dblSum = dblSum + ( UNICDbl(GetGrid(TYPE_2_2, C_W10, 7)) * UNICDbl(GetGrid(TYPE_2_2, C_W11, 6)) )
							Call PutGrid(TYPE_2_2, C_W13, 6, dblSum)
						Else
							dblSum = UNICDbl(GetGrid(pType, C_W10, Row)) * UNICDbl(GetGrid(pType, C_W11, Row)) 
							.Row = Row : .Col = C_W8
							If .Text = "07" Or .Text = "08" Then
								Call PutGrid(pType, C_W13, Row, "0")
							Else
								Call PutGrid(pType, C_W13, Row, dblSum)
							End If
						End If
						
						Call ReCalc_W15
					Case C_W13
						Call ReCalc_W15
				End Select
			
			Case TYPE_3
				.Row = Row	: .Col = Col
				Select Case Col
					Case C_W17					
						If instr(1, .Text,"계")<>0 Then						
							Exit Sub
						ElseIF CheckMinorCd("W1034", .Text) = False Then
							.Text = ""
							lgChkFlag=False
							Exit Sub
						End If
						.Col = C_W17_NM : .value = GetCodeNm("3", pVal)
					Case C_W19
						If chkCW19(TYPE_3,.text,Row) =false then 
							lgChkFlag=False
							Exit Sub						
						End IF
						
					Case C_W36
						If  len(trim(.text))=0 Then						
							Exit Sub
						ElseIf CheckMinorCd("W1035", .Text) = False Then
							.Text = ""
						End If
						.Col = C_W36_NM : .value = GetCodeNm("4", pVal)
					Case C_W20, C_W22, C_W23, C_W24, C_W25, C_W26, C_W27, C_W28, C_W29, C_W30, C_W31, C_W32, C_W33
						Call FncSumSheet(lgvspdData(TYPE_3), Col, 2, .MaxRows, true, 1, Col, "V")	' 합계(주식수)
						
						dblAmt(C_W20) = UNICDbl(GetGrid(TYPE_3, C_W20, Row))
						dblAmt(C_W22) = UNICDbl(GetGrid(TYPE_3, C_W22, Row))
						dblAmt(C_W23) = UNICDbl(GetGrid(TYPE_3, C_W23, Row))
						dblAmt(C_W24) = UNICDbl(GetGrid(TYPE_3, C_W24, Row))
						dblAmt(C_W25) = UNICDbl(GetGrid(TYPE_3, C_W25, Row))
						dblAmt(C_W26) = UNICDbl(GetGrid(TYPE_3, C_W26, Row))
						dblAmt(C_W27) = UNICDbl(GetGrid(TYPE_3, C_W27, Row))
						dblAmt(C_W28) = UNICDbl(GetGrid(TYPE_3, C_W28, Row))
						dblAmt(C_W29) = UNICDbl(GetGrid(TYPE_3, C_W29, Row))
						dblAmt(C_W30) = UNICDbl(GetGrid(TYPE_3, C_W30, Row))
						dblAmt(C_W31) = UNICDbl(GetGrid(TYPE_3, C_W31, Row))
						dblAmt(C_W32) = UNICDbl(GetGrid(TYPE_3, C_W32, Row))
						dblAmt(C_W33) = UNICDbl(GetGrid(TYPE_3, C_W33, Row))
						dblAmt(C_W34) = dblAmt(C_W20) + dblAmt(C_W22) + dblAmt(C_W23) + dblAmt(C_W24) + dblAmt(C_W25) + dblAmt(C_W26) + dblAmt(C_W27) + dblAmt(C_W28)
						dblAmt(C_W34) = dblAmt(C_W34) - dblAmt(C_W29) - dblAmt(C_W30) - dblAmt(C_W31) - dblAmt(C_W32) - dblAmt(C_W33)
						
						Call PutGrid(TYPE_3, C_W34, Row, dblAmt(C_W34))
						
						Call ReCalcGrid()
				End Select 
		End Select
			
	End With
	
End Sub


' -- 그리드 W10, W11
Function ReCalc_W15()
	Dim iType, iRow, i, dblAmt(2)
	
	With lgvspdData(TYPE_2_1)
		For iRow = 1 To 7
			.Row = iRow
			.Col = C_W9
				
			If .Text = "01" Then
				.Col = C_W8
				Select Case UNICDbl(.Text)
					Case "01", "02", "03", "04", "07"
						.Col = C_W10
						dblAmt(1) = dblAmt(1) + UNICDbl(.value)
					Case "05", "06", "08", "09"
						.Col = C_W10
						dblAmt(1) = dblAmt(1) - UNICDbl(.value)
					Case Else
						.Col = C_W10
						dblAmt(1) = dblAmt(1) + UNICDbl(.value)
				End Select
			ElseIf .Text = "02" Then
				.Col = C_W8
				Select Case UNICDbl(.Text)
					Case "01", "02", "03", "04", "07"
						.Col = C_W10
						dblAmt(2) = dblAmt(2) + UNICDbl(.value)
					Case "05", "06", "08", "09"
						.Col = C_W10
						dblAmt(2) = dblAmt(2) - UNICDbl(.value)
					Case Else
						.Col = C_W10
						dblAmt(2) = dblAmt(2) + UNICDbl(.value)
				End Select
			End If
				
			.Col = C_W13 
			dblAmt(0) = dblAmt(0) + UNICDbl(.value)
		Next
	End With

	With lgvspdData(TYPE_2_2)
		For iRow = 1 To 5
			.Row = iRow
			.Col = C_W9
				
			If .Text = "01" Then
				.Col = C_W8
				Select Case UNICDbl(.Text)
					Case "01", "02","03", "04", "07"
						.Col = C_W10
						dblAmt(1) = dblAmt(1) + UNICDbl(.value)
					Case "05", "06", "08", "09"
						.Col = C_W10
						dblAmt(1) = dblAmt(1) - UNICDbl(.value)
					Case Else
						.Col = C_W10
						dblAmt(1) = dblAmt(1) + UNICDbl(.value)
				End Select
			ElseIf .Text = "02" Then
				.Col = C_W8
				Select Case UNICDbl(.Text)
					Case "01", "02", "03", "04", "07"
						.Col = C_W10
						dblAmt(2) = dblAmt(2) + UNICDbl(.value)
					Case "05", "06", "08", "09"
						.Col = C_W10
						dblAmt(2) = dblAmt(2) - UNICDbl(.value)
					Case Else
						.Col = C_W10
						dblAmt(2) = dblAmt(2) + UNICDbl(.value)
				End Select
			End If
				
			.Col = C_W13 
			dblAmt(0) = dblAmt(0) + UNICDbl(.value)
		Next
	End With
	
	
	With lgvspdData(TYPE_2_2)
		ggoSpread.Source = lgvspdData(TYPE_2_2)
		.Row = 6	: .Col = C_W10	: .value = dblAmt(1)
					  .Col = C_W13	: .value = dblAmt(0)
		ggoSpread.UpdateRow .Row
		.Row = 7	: .Col = C_W10	: .value = dblAmt(2)	
		ggoSpread.UpdateRow .Row
		
	End With
	
End Function

Sub vspdData_Click(Byval pType, ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(pType)
   
    If lgvspdData(pType).MaxRows <=1 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(pType)      
      
       Exit Sub
    End If
    
	lgvspdData(pType).Row = Row
End Sub

Sub vspdData_ColWidthChange(Byval pType, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(pType)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(Byval pType, ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If lgvspdData(pType).MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus(Byval pType)
    ggoSpread.Source = lgvspdData(pType)

End Sub

Sub vspdData_MouseDown(Byval pType, Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Byval pType, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = lgvspdData(pType)
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(Byval pType, ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if lgvspdData(pType).MaxRows < NewTop + VisibleRowCnt(lgvspdData(pType),NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(Byval pType, ByVal Col, ByVal Row, Byval ButtonDown)
	With lgvspdData(pType)
		.Row = Row
		Select Case pType
			Case TYPE_2_1, TYPE_2_2
				If Col = C_W8_P Then
					.Col = Col - 1
					Call OpenW1082(.Value)
				ElseIf Col = C_W9_P Then
					.Col = Col - 1
					Call OpenW1083(.Value)
				End If
				
			Case TYPE_3	
				Select Case Col
					Case C_W17_P
						.Col = Col - 1
						Call OpenW1034(.text)
					Case C_W36_P
						.Col = Col - 1
						Call OpenW1035(.Value)					
				End Select
		End Select
	
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

    Call SetToolbar("1100110100000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function

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

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue Then
		ggoSpread.Source = lgvspdData(TYPE_3)
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
			If IntRetCD = vbNo Then
		  	Exit Function
			End If
		End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, dblSum, iType
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	
	If lgvspdData(TYPE_1)(C_W1).value = "" Then
		Call DisplayMsgBox("W90002", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

'	If Not chkField(Document, "2") Then                             '⊙: Check contents area
'	   Exit Function
'	End If
	For iType = TYPE_2_1 To TYPE_3
	
		If lgvspdData(iType).MaxRows > 0 Then
	
			ggoSpread.Source = lgvspdData(iType)
			If ggoSpread.SSCheckChange = True Then
				blnChange = True
			End If

			If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
			      Exit Function
			End If    
	
		End If
	Next

	If blnChange = False Then
	    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
	    Exit Function
	End If
				
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 

End Function

Function FncCancel() 
	With lgvspdData(TYPE_3)
		ggoSpread.Source = lgvspdData(TYPE_3)	
		.Row = .ActiveRow
		ggoSpread.EditUndo                                                   '☜: Protect system from crashing
    End With
    ' 삭제후 결과를 다른행에 반영한다.
    Call ReCalcGrid()
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
 
	With lgvspdData(TYPE_3)	' 포커스된 그리드 
			
		ggoSpread.Source = lgvspdData(TYPE_3)
			
		iRow = .ActiveRow
		.ReDraw = False
					
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow , 3
			Call SetSpreadLock(TYPE_3) 
			.Row = iRow		

			iRow = 1		: .Row = iRow
			.Col = C_W17_1	: .Value = "1"
			.Col = C_W16	: .Value = "1"
			.Col = C_W21	: .Text = "100%"
			.Col = C_W35	: .Text = "100%"
			
			iRow = 2		: .Row = iRow
			.Col = C_W17_1	: .Value = "2"
			.Col = C_W16	: .Value = "2"
			
			iRow = 3		: .Row = iRow
			.Col = C_W17_1	: .Value = "3"
			.Col = C_W16	: .Value = "3"
		
			Call SetTotalLine
		Else
				
			If iRow = 1 Or iRow = 2  Then	' -- 합계줄에서 InsertRow를 하면 하위에 추가한다.
				iRow = .MaxRows 
				ggoSpread.InsertRow iRow , imRow 

			Else
				ggoSpread.InsertRow ,imRow
			End If   
			
			SetSpreadColor TYPE_3, iRow+1, iRow + imRow
			Call SetDefaultVal( iRow+1, imRow)

		End If
	End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
         
End Function

' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(TYPE_3)	' 포커스된 그리드 

	ggoSpread.Source = lgvspdData(TYPE_3)
	
	If iAddRows = 1 Then ' 1줄만 넣는경우 
		.Row = iRow
		.Col = C_W17_1	: .Value = "3"
		MaxSpreadVal lgvspdData(TYPE_3), C_W16, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(TYPE_3), C_W16, iRow)	' 현재의 최대SeqNo를 구한다 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i	
			.Col = C_W32_1	: .Value = "3"
			.Col = C_W31	: .Value = iSeqNo : iSeqNo = iSeqNo + 1
		Next
	End If
	End With
End Function

Sub SetTotalLine()
	With lgvspdData(TYPE_3)
		.Row = 1
		.Col = C_W17	: .TypeHAlign = 2	: .value = "합 계"
		.Col = C_W17_P	: .CellType = 1
		.Col = C_W36_P	: .CellType = 1
		
		.Row = 2		
		.Col = C_W17	: .TypeHAlign = 2	
		.TypeEditMultiLine = True	
		.value = "소액주주" & vbCrLf & "(소액출자자)" & vbCrLf & "소계"
		.RowHeight(2) = 25		
		
		.Col = C_W17_P	: .CellType = 1
		.Col = C_W36_P	: .CellType = 1
		
	End With
End Sub

Function FncDeleteRow() 
    Dim lDelRows

    With lgvspdData(TYPE_3) 
    	.focus
    	ggoSpread.Source = lgvspdData(TYPE_3)
    	
    	If .MaxRows = 0 Then Exit Function
    	If (.ActiveRow = 1 Or .ActiveRow = 2) And .MaxRows > 2 Then
    		Call  DisplayMsgBox("WC0032", parent.VB_INFORMATION, "X", "X") 
    		Exit Function
    	End If
    	
    	lDelRows = ggoSpread.DeleteRow
    	
    	' 삭제후 결과를 다른행에 반영한다.
    	Call ReCalcGrid()
    	
    	lgBlnFlgChgValue = True
    End With
End Function

Function ReCalcGrid()
	Dim iRow, iMaxRows, dblW(30), dblW35Sum, dblW50Sum, dblW35, dblW50, dblW20Sum, dblW20, dblW34Sum, dblW34
	
	With lgvspdData(TYPE_3)
		.ReDraw  = False
		iMaxRows = .MaxRows
		
		Call FncSumSheet(lgvspdData(TYPE_3), C_W20, 2, .MaxRows, true, 1, C_W20, "V")	' 합계(주식수)
		Call FncSumSheet(lgvspdData(TYPE_3), C_W34, 2, .MaxRows, true, 1, C_W34, "V")	' 합계(주식수)
		
		dblW20Sum = UNICDbl(GetGrid(TYPE_3 , C_W20, 1))
		dblW34Sum = UNICDbl(GetGrid(TYPE_3 , C_W34, 1))
		
		For iRow = 2 To iMaxRows
			.Row = iRow

			.Col = 0		
			If .Text <> ggoSpread.DeleteFlag Then
				.Col = C_W20	: dblW20 = UNICDbl(.text)
				
				If dblW20Sum > 0 Then .Col = C_W21	: .value = dblW20 / dblW20Sum
				.Col = C_W34	: dblW34 = UNICDbl(.value)
				If dblW34Sum > 0 Then .Col = C_W35	: .value = dblW34 / dblW34Sum
				
				.Col = C_W20	: dblW(C_W20) = dblW(C_W20) + UNICDbl(.text)
				'.Col = C_W21	: dblW(C_W21) = dblW(C_W21) + UNICDbl(.text)
				.Col = C_W22	: dblW(C_W22) = dblW(C_W22) + UNICDbl(.text)
				.Col = C_W23	: dblW(C_W23) = dblW(C_W23) + UNICDbl(.text)
				.Col = C_W24	: dblW(C_W24) = dblW(C_W24) + UNICDbl(.text)
				.Col = C_W25	: dblW(C_W25) = dblW(C_W25) + UNICDbl(.text)
				.Col = C_W26	: dblW(C_W26) = dblW(C_W26) + UNICDbl(.text)
				.Col = C_W27	: dblW(C_W27) = dblW(C_W27) + UNICDbl(.text)
				.Col = C_W28	: dblW(C_W28) = dblW(C_W28) + UNICDbl(.text)
				.Col = C_W29	: dblW(C_W29) = dblW(C_W29) + UNICDbl(.text)
				.Col = C_W30	: dblW(C_W30) = dblW(C_W30) + UNICDbl(.text)
				.Col = C_W31	: dblW(C_W31) = dblW(C_W31) + UNICDbl(.text)
				.Col = C_W32	: dblW(C_W32) = dblW(C_W32) + UNICDbl(.text)
				.Col = C_W33	: dblW(C_W33) = dblW(C_W33) + UNICDbl(.text)
				.Col = C_W34	: dblW(C_W34) = dblW(C_W34) + UNICDbl(.text)
				'.Col = C_W35	: dblW(C_W35) = dblW(C_W35) + UNICDbl(.text)
			End If
			
		Next
		
		.Row = 1
		.Col = C_W20	: .value = dblW(C_W20) 
		'msgbox .value & "vlaue"
		'.Col = C_W21	: .Text = "100%"
		.Col = C_W22	: .value = dblW(C_W22)
		.Col = C_W23	: .value = dblW(C_W23)
		.Col = C_W24	: .value = dblW(C_W24)
		.Col = C_W25	: .value = dblW(C_W25)
		.Col = C_W26	: .value = dblW(C_W26) 
		.Col = C_W27	: .value = dblW(C_W27) 
		.Col = C_W28	: .value = dblW(C_W28)
		.Col = C_W29	: .value = dblW(C_W29) 
		.Col = C_W30	: .value = dblW(C_W30) 
		.Col = C_W31	: .value = dblW(C_W31) 
		.Col = C_W32	: .value = dblW(C_W32) 
		.Col = C_W33	: .value = dblW(C_W33) 
		.Col = C_W34	: .value = dblW(C_W34) 
		'.Col = C_W35	: .Text = "100%"

		ggoSpread.Source = lgvspdData(TYPE_3)
		ggoSpread.UpdateRow 1
				
		.ReDraw  = True
	End With
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
	
	If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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
        
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgvspdData(TYPE_3).MaxRows > 0 Or GetGrid(TYPE_2_1, C_W7, 3) <> "" Then
    
		lgIntFlgMode = parent.OPMD_UMODE
		Call SetSpreadLock(TYPE_2_1)
		Call SetSpreadLock(TYPE_2_2)
		Call SetSpreadLock(TYPE_3)
		
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101111100000111")										<%'버튼 툴바 제어 %>

		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
		End If
	Else
		Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>
	End If

End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow, lCol   
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel, lMaxRows, lMaxCols, iType, sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    lGrpCnt = 0
	   
	If fnchkSum =False Then 
		Call LayerShowHide(0)
		Exit Function 
	End If
	   
	' 그리드 부분 
	For iType = TYPE_2_1 To TYPE_3
	
		With lgvspdData(iType)
			ggoSpread.Source = lgvspdData(iType)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			strVal = ""
			strDel = ""					
			For lRow = 1 To lMaxRows
			    
			 .Row = lRow : .Col = 0 : sTmp = ""			
			 
			  ' 모든 그리드 데이타 보냄     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = 1 To lMaxCols						
						If iType=TYPE_3 and lCol = C_W19 then
							.Col = lCol
							
							If chkCW19(TYPE_3, .text,.Row) =False then	
								lgChkFlag=False
								Call LayerShowHide(0)
								Exit function			
							End IF	
						End IF
										
						If lgChkFlag=False then
							Call LayerShowHide(0)
							 .Row=lRow : .Col=lCol : .focus
							Exit function	
						End IF
						'Select Case lCol
						'	Case C_W3
						'		.Col = lCol : sTmp = sTmp & Trim(.Value) &  Parent.gColSep
						'	Case Else
								.Col = lCol : sTmp = sTmp & Trim(.Text) &  Parent.gColSep
								
						'End Select
					Next
					sTmp = sTmp & Trim(.Text) &  Parent.gRowSep
			  End If  

			   .Row = lRow : .Col = 0
			   
			   ' I/U/D 플래그 처리 
			   Select Case .Text
			       Case  ggoSpread.InsertFlag                                      '☜: Insert
			                                          strVal = strVal & "C"  &  Parent.gColSep & sTmp
			            lGrpCnt = lGrpCnt + 1
			                    
			       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
			                                          strVal = strVal & "U"  &  Parent.gColSep & sTmp                                                 
			            lGrpCnt = lGrpCnt + 1                                                 
			       Case  ggoSpread.DeleteFlag                                      '☜: Delete
			                                          strDel = strDel & "D"  &  Parent.gColSep & sTmp
			            lGrpCnt = lGrpCnt + 1  
			  End Select
			 
 
			Next
		End With
		
		document.all("txtSpread"&iType).value = strDel & strVal
	Next
	
    strVal = ""
    strDel = ""
    lgvspdData(TYPE_1)(C_W89).value = lgvspdData(TYPE_3).MaxRows
	
    With lgvspdData(TYPE_1)
    ' -- 헤더 저장 
		For lCol = C_W1 To C_W89
			Select Case lCol
				Case C_W4, C_W5,C_W6_1,C_W6_2
					strVal = strVal & Trim(lgvspdData(TYPE_1)(lCol).text) &  Parent.gColSep	' 콘트롤 
				Case Else
					strVal = strVal & Trim(lgvspdData(TYPE_1)(lCol).value) &  Parent.gColSep ' html input
			End Select
		Next
	End With
	frm1.txtSpread0.value   = strVal
	frm1.txtHeadMode.value	= lgIntFlgMode
	frm1.txtMode.value        =  Parent.UID_M0002
    
    strVal = ""
    strDel = ""
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	lgvspdData(TYPE_2_1).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_2_1)
    ggoSpread.ClearSpreadData

	lgvspdData(TYPE_2_2).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_2_2)
    ggoSpread.ClearSpreadData

	lgvspdData(TYPE_3).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_3)
    ggoSpread.ClearSpreadData    	
    
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

Function DbDeleteOk()
	Call FncNew()
End Function

Function fnchkSum()
	Dim tmpSum, tmpSum1,i, tmpSum2, tmpSum3
	fnchkSum=True
	
	If  lgvspdData(TYPE_3).MaxRows <=1 Then 
			Call DisplayMsgBox("WC0004", "X", "(10)주식수(출자좌수)", "(34)주식수(좌수)값")
			fnchkSum = False
			Exit Function 
	End IF
		
	tmpSum =0
	with lgvspdData(TYPE_2_1)
		
		' -- 기초주식수 체크 
		.Col = C_W10
		.Row = 1 : tmpSum2 = UNICDbl(.Value)
		.Row = 2 : tmpSum2 = tmpSum2 + UNICDbl(.Value)

    End With 

	with lgvspdData(TYPE_2_2)
		' -- 기말 체크 
		' -- 기초주식수 체크 
		.Col = C_W10
		.Row = 6 : tmpSum = UNICDbl(.Value)
		.Row = 7 : tmpSum = tmpSum + UNICDbl(.Value)

	End With
	
     With  lgvspdData(TYPE_3)
		
		.Row=1 : .Col = C_W16
		If .TEXT =1 THEN 
			.Row = 1 : .Col = C_W34
			tmpSum1=uniCDbl(.Value)
			
			.Row = 1 : .Col = C_W20
			tmpSum3=uniCDbl(.Value)
		Else
			.Row=.MaxRows : .Col =C_W34
			tmpSum1=uniCdbl(.Value)

			.Row=.MaxRows : .Col =C_W20
			tmpSum3=uniCdbl(.Value)
		End IF     
     End With
     
     If tmpSum <> tmpSum1 Then
		
		Call DisplayMsgBox("WC0004", "X", "(15) 기말 (10)주식수(출자좌수)", "(34)주식수(좌수)값")
		fnchkSum = False
		Exit Function 
     ElseIf tmpSum2 <> tmpSum3 Then
		
		Call DisplayMsgBox("WC0004", "X", "(14) 기초 (10)주식수(출자좌수)", "(20)주식수(좌수)값")
		fnchkSum = False
		Exit Function 
	End If
	fnchkSum=True
End Function 

Function chkCW19(byVal pType, byVal pTxt,byVal pRow)
	Dim tmpVal
	
	chkCW19=True
	
	With lgvspdData(pType)
		tmpVal= pTxt
		.Col=C_W19-3	: .Row =pRow
		If tmpVal = "" Then
			chkCW19=True					
			Exit Function
		ElseIf  instr(1, .Text,"계")<>0 Then	
			chkCW19=True					
			Exit Function
		ElseIf len(tmpVal)<4 Then 							
			Call DisplayMsgBox("WC0017", "x", "(19)주민등록번호(사업자등록번호)", "4")
			chkCW19=False
			Exit Function
		End IF
		If .Text="01" and len(tmpVal) <>13 Then
			Call DisplayMsgBox("126134", "x", "x", "x")
			chkCW19=False
			Exit Function						
		End If
	End with
	chkCW19=True
End Function 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT LANGUAGE=javascript FOR=txtData EVENT=Change>
<!--

//-->
</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=300 border=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" width=80% align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:GetRef"></A></TD>
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
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"> <% ' -- overflow=auto : 컨텐츠 구역을 브라우저 크기에 따라 스크롤바가 생성되게 한다 %>
					<TABLE <%=LR_SPACE_TYPE_20%> border="0" width="100%">
					   <TR HEIGHT=20>
							<TD WIDTH=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
								<TR HEIGHT=10>
								       <TD CLASS="TD51" WIDTH="12%">(1)법인명</TD>
									   <TD CLASS="TD61" WIDTH="20%"><INPUT TYPE=TEXT tag="24" style="width: 100%" id="txtData" name=txtData></TD>
									   <TD CLASS="TD51" WIDTH="12%">(2)사업자등록번호</TD>
									   <TD CLASS="TD61" WIDTH="20%"><INPUT TYPE=TEXT tag="24" style="width: 100%" id="txtData" name=txtData style="text-align: center"></TD>
									   <TD CLASS="TD51" WIDTH="12%">(3)대표자</TD>
									   <TD CLASS="TD61" WIDTH="25%"><INPUT TYPE=TEXT tag="24" style="width: 100%" id="txtData" name=txtData></OBJECT></TD>
				
								</TR>
								<TR HEIGHT=10>
								       <TD CLASS="TD51">(4)상장(등록)변경일</TD>
									   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="txtData" name=txtData CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="(4)상장(등록)변경일" tag="25" width = 50%></OBJECT>');</SCRIPT></TD>
									   <TD CLASS="TD51">(5)합병.분할일</TD>
									   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="txtData" name=txtData CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="(5)합병.분할일" tag="25" width = 50%></OBJECT>');</SCRIPT></TD>
									   <TD CLASS="TD51">(6)사업연도</TD>
									   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="txtData" name=txtData CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="(6)시작사업연도" tag="23" width = 45%></OBJECT>');</SCRIPT>&nbsp;~&nbsp; 
									   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="txtData" name=txtData CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="(6)끝사업연도" tag="23x" width = 45%></OBJECT>');</SCRIPT></TD>
								</TR>
								</TABLE>
							</TD>
						</TR>
					   <TR HEIGHT=175>
							<TD WIDTH=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> border="0" width="100%">
									<TR>
										<TD WIDTH=50%>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
										<TD WIDTH=50%>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
								</TABLE>								
							</TD>
						</TR>
					   <TR HEIGHT=250>
							<TD WIDTH=100%>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							<TD>
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
		<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			   <TR>
				        <TD WIDTH=10>&nbsp;</TD>
				        <TD width=30%><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>주주변동사항리스트</LABEL>&nbsp;
				        </TD>
				        <TD ROWSPAN=2>코드보기: <span title="유상증자(01), 무상증자(02), 출자전환(03), 주식배당(04), 유상감자(05), 무상감자(06), 액면분할(07), 주식병합(08), 기타(자사주소각)(09)">(8)원인코드</span></TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24" STYLE="display: 'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" STYLE="display: 'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" STYLE="display: 'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24" STYLE="display: 'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHeadMode" tag="24">
<INPUT TYPE=HIDDEN id="txtData" name=txtData tag="24" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

