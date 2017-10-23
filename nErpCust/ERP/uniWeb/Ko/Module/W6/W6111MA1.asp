
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 조특제3호연구및인력개발명세서 
'*  3. Program ID           : W6111MA1
'*  4. Program Name         : W6111MA1.asp
'*  5. Program Desc         : 조특제3호연구및인력개발명세서 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2007/2
'*  8. Modifier (First)     : 홍지영 
'*  9. Modifier (Last)      : leewolsan 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
    '<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
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

Const BIZ_MNU_ID		= "W6111MA1"
Const BIZ_PGM_ID		= "W6111Mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID		= "W6111OA1"

Const TAB1 = 1																	'☜: Tab의 위치 
Const TAB2 = 2

Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.

' -- 그리드 컬럼 정의 
Dim C_SEQ_NO	
Dim C_ACCT
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W1_TITLE
Dim C_W2_TITLE
Dim C_W3_TITLE
Dim C_W4_TITLE
Dim C_W5_TITLE


Dim C_COL1
Dim C_COL2
Dim C_COL3
Dim C_COL4
Dim C_COL5
Dim C_COL6
Dim C_COL7

Dim C_COL8
Dim C_COL9
Dim C_COL10
Dim C_COL11
Dim C_COL12
Dim C_COL13
Dim C_COL14
Dim C_COL15



Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(2)

Dim lgW2, lgMonth	' 설정률, 사업연도월수 
Dim lgFiscStartDt, lgFiscEndDt

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1	' -- 1번 그리드 
	C_ACCT		= 2	' 
	C_W1		= 3	'
	C_W2		= 4	' 
	C_W3		= 5 ' 
	C_W4		= 6	' 
	C_W5		= 7	'
	C_W6		= 8	'
	C_W1_TITLE	= 9	' 
	C_W2_TITLE	= 10' 
	C_W3_TITLE	= 11'
	C_W4_TITLE	= 12	'  
	C_W5_TITLE	= 13	' 

    C_COL1		= 1	' 
    C_COL2		= 2	' 
    C_COL3		= 3	' 
    C_COL4		= 4	'
    C_COL5		= 5	' 
    C_COL6		= 6	' 
    C_COL7		= 7	' 
    C_COL8		= 8	' 
    C_COL9		= 9	' 
    C_COL10		= 10	'
    C_COL11		= 11	' 
    C_COL12		= 12	' 
    C_COL13		= 13	' 
    C_COL14		= 14	' 
    C_COL15		= 15	' 
	 
	
	

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





'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt
	Dim w8_d1_s    '기간1의 시작일			row = 1  : col = c_col3
	Dim w8_d1_e    '기간1의 종료일			row = 1  : col = c_col5
    Dim w8_d2_s    '기간2의 시작일			row = 1  : col = c_col6
	Dim w8_d2_e    '기간2의 종료일			row = 1  : col = c_col8
    Dim w8_d3_s    '기간3의 시작일			row = 1  : col = c_col9
	Dim w8_d3_e    '기간3의 종료일			row = 1  : col = c_col11
	Dim w8_d4_s    '기간4의 시작일			row = 1  : col = c_col12
	Dim w8_d4_e    '기간4의 종료일			row = 1  : col = c_col14
	
	Dim w8_Amt1    '기간1의 발생합계액		row = 2  : col = c_col3
	Dim w8_Amt2    '기간2의 발생합계액		row = 2  : col = c_col6
	Dim w8_Amt3    '기간3의 발생합계액		row = 2  : col = c_col9
	Dim w8_Amt4    '기간4의 발생합계액		row = 2  : col = c_col12
	Dim w8_Sum     '기간1~4의 발생합계액	row = 2  : col = c_col15
	
	Dim w9         '직전4년간 연평균발생액 	row = 4  : col = c_col1
	Dim w10        '증가발생액          	row = 5  : col = c_col3
	Dim w15_11     '15당해년도 총발생금액공제-11대상금액          	row = 8  : col = c_col1
	Dim w15_12_View   '15-12 공제율        	row = 8  : col = c_col6
	Dim w15_12_Value  '15-12 공제율        	frm1.txt15_12Value.value
	Dim w15_13     '15-13공제세액        	row = 8  : col = c_col9
	Dim w15_14     '15-14비고           	row = 8  : col = c_col12
	
	
	Dim w16_11     '16증가발생금액공제-11대상금액     row = 9  : col = c_col1
	Dim w16_12_View   '16-12 공제율        	row = 9  : col = c_col6
	Dim w16_12_Value  '16-12 공제율        	frm1.txt16_12Value.value
	Dim w16_13     '16-13공제세액        	row = 9  : col = c_col9
	Dim w16_14     '16-14비고           	row = 9  : col = c_col12
	Dim w17        '17 당해연도에 공제받을세액        	row = 10  : col = c_col12
	Dim w17_14        '17-14 비고     	row = 10  : col = c_col12
	Dim w18_A         '18    석박사급인건비	row = 12  : col = c_col9
	Dim w18_B         '18    연구및 인력개발비 당기발생액	row = 13  : col = c_col9
	Dim w18_C         '18    석박사급 인건비에 대한 세액공제	row = 14  : col = c_col9
	Dim w18_A_14        '17-14 비고     	row = 12  : col = c_col12
	Dim w18_B_14        '17-14 비고     	row = 13  : col = c_col12
	Dim w18_C_14        '17-14 비고     	row = 14  : col = c_col12
	Dim wDESC       '⑬ ×( 48 / 직전 4년간의 사업연도 월수) × (1 / 4) × (당해연도 월수 / 12)     	row = 4  : col = c_col2
	
	Dim CompType 
	
	if pOpt = "S" then
			With frm1.vspdData1
				    ggoSpread.Source = frm1.vspdData1
					.row = 1  : .col = c_col3 : w8_d1_s  = .text
					.row = 1  : .col = c_col5 : w8_d1_e  = .text
					.row = 1  : .col = c_col6 : w8_d2_s  = .text
					.row = 1  : .col = c_col8 : w8_d2_e  = .text
					.row = 1  : .col = c_col9 : w8_d3_s  = .text
					.row = 1  : .col = c_col11 : w8_d3_e  = .text
					.row = 1  : .col = c_col12 : w8_d4_s  = .text
					.row = 1  : .col = c_col14 : w8_d4_e  = .text
					
					.row = 2  : .col = c_col3 :  w8_Amt1  = .text
					.row = 2  : .col = c_col6 :  w8_Amt2  = .text
					.row = 2  : .col = c_col9 :  w8_Amt3  = .text
					.row = 2  : .col = c_col12 : w8_Amt4  = .text
					.row = 2  : .col = c_col15 : w8_Sum  = .text
					
					.row = 4  : .col = c_col1 : w9  = .text
  				    .row = 5  : .col = c_col3 : w10  = .text
  				    .row = 8  : .col = c_col2:  w15_11  = .text
  				    .row = 8  : .col = c_col6 : w15_12_View  = .text
  				    
  				     w15_12_Value =	frm1.txt15_12Value.value
  				  
  				    .row = 8  : .col = c_col9 :  w15_13  = .text
  				    .row = 8  : .col = c_col12 : w15_14  = .text 
  				    .row = 9  : .col = c_col2 :  w16_11  = .text 
  				    .row = 9  : .col = c_col6 :  w16_12_View  = .text 
  				     w16_12_Value = frm1.txt16_12Value.value
  				    .row = 9  : .col = c_col9 : w16_13  = .text
  				    .row = 9  : .col = c_col12 : w16_14  = .text  
  				     CompType = frm1.txtCompType.value 
  				    if CompType = 2 then   '중소기업 
  				       .row = 10  : .col = c_col9 :  w17  = .text 
  				       .row = 10  : .col = c_col12 : w17_14  = .text  
  				    else
  				       
  				       .row = 11  : .col = c_col9 : w17  = .text 
  				       .row = 11  : .col = c_col12 : w17_14  = .text 
  				   end if
  				       .row = 12  : .col = c_col9 : w18_A  = .text
  				       .row = 13  : .col = c_col9 : w18_B  = .text 
  				       .row = 14  : .col = c_col9 : w18_C  = .text
  				       
  				       .row = 12  : .col = c_col12 : w18_A_14  = .text
  				       .row = 13  : .col = c_col12 : w18_B_14  = .text 
  				       .row = 14  : .col = c_col12 : w18_C_14  = .text    
  				       .row = 4  : .col = c_col2 : wDESC  = .text    
  			
					 
			End With	 
			
					lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep					 '0 법인코드 
					lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep '	 '1 년도 
					lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '2 신고구분 
					lgKeyStream = lgKeyStream & w8_d1_s &  parent.gColSep '						  '3 기간	
					lgKeyStream = lgKeyStream & w8_d1_e &  parent.gColSep						  '4
					lgKeyStream = lgKeyStream & w8_d2_s &  parent.gColSep						  '5
					lgKeyStream = lgKeyStream & w8_d2_e &  parent.gColSep						  '6
					lgKeyStream = lgKeyStream & w8_d3_s &  parent.gColSep						  '7	
					lgKeyStream = lgKeyStream & w8_d3_e &  parent.gColSep						  '8
					lgKeyStream = lgKeyStream & w8_d4_s &  parent.gColSep						  '9
					lgKeyStream = lgKeyStream & w8_d4_e &  parent.gColSep						  '10
					lgKeyStream = lgKeyStream & w8_Amt1 &  parent.gColSep						  '11
					lgKeyStream = lgKeyStream & w8_Amt2 &  parent.gColSep						  '12
					lgKeyStream = lgKeyStream & w8_Amt3 &  parent.gColSep						  '13	
					lgKeyStream = lgKeyStream & w8_Amt4 &  parent.gColSep						  '14	
					lgKeyStream = lgKeyStream & w8_Sum &  parent.gColSep						  '15
					lgKeyStream = lgKeyStream & w9     &  parent.gColSep						  '16	
					lgKeyStream = lgKeyStream & w10 &  parent.gColSep							  '17
					lgKeyStream = lgKeyStream & w15_11 &  parent.gColSep						  '18
					lgKeyStream = lgKeyStream & w15_12_View &  parent.gColSep					  '19
					lgKeyStream = lgKeyStream & w15_12_Value &  parent.gColSep					  '20
					lgKeyStream = lgKeyStream & w15_13 &  parent.gColSep						  '21
					lgKeyStream = lgKeyStream & w15_14 &  parent.gColSep						  '22	
					lgKeyStream = lgKeyStream & w16_11 &  parent.gColSep						  '23	
					lgKeyStream = lgKeyStream & w16_12_View &  parent.gColSep				       '24
					lgKeyStream = lgKeyStream & w16_12_Value &  parent.gColSep					  '25
					lgKeyStream = lgKeyStream & w16_13 &  parent.gColSep						  '26
					lgKeyStream = lgKeyStream & w16_14 &  parent.gColSep						   '27
					lgKeyStream = lgKeyStream & CompType &  parent.gColSep						   '28
					lgKeyStream = lgKeyStream & w17 &  parent.gColSep								'29
					lgKeyStream = lgKeyStream & w17_14 &  parent.gColSep							'30
					lgKeyStream = lgKeyStream & w18_A &  parent.gColSep							'31
					lgKeyStream = lgKeyStream & w18_A_14 &  parent.gColSep							'32
					lgKeyStream = lgKeyStream & w18_B &  parent.gColSep							'33
					lgKeyStream = lgKeyStream & w18_B_14 &  parent.gColSep							'34
					lgKeyStream = lgKeyStream & w18_C &  parent.gColSep							'35
					lgKeyStream = lgKeyStream & w18_C_14 &  parent.gColSep							'36
					lgKeyStream = lgKeyStream & wDESC &  parent.gColSep							'37
					lgKeyStream = lgKeyStream & Trim(frm1.txtyearMth.value) &  parent.gColSep							'38
					lgKeyStream = lgKeyStream & Trim(frm1.txt4yearMth.value) &  parent.gColSep							'39

	Else
	        lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
			lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
			lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '
	End if		
   
End Sub 




'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	' 조회조건(구분)
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))


End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1

    Call initSpreadPosVariables()  
	

	' 1번 그리드 
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
	
		ggoSpread.Spreadinit "V20061222" & TYPE_1,,parent.gForbidDragDropSpread    
    
    
		.ReDraw = false

		.MaxCols = C_W5_TITLE + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    

			       
		.MaxRows = 1
		ggoSpread.ClearSpreadData

	    .ColHeadersShow = False

		ggoSpread.SSSetEdit		C_SEQ_NO,	"순번"		, 10,,,6,1	
		ggoSpread.SSSetEdit		C_ACCT,	    "계정"		, 20,,,50,1	
		ggoSpread.SSSetFloat     C_W1,	    "(1)인건비"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W2,		"(2)재료비"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W3,		"(3)위탁개발비"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W4,		"(4)위탁훈련비"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W5,		"(5)기타"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W6,		"(6)계"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetEdit		C_W1_TITLE,	    "계정항목"		, 20,,,50,1	
		ggoSpread.SSSetEdit		C_W2_TITLE,	    "계정"		, 20,,,50,1	
		ggoSpread.SSSetEdit		C_W3_TITLE,	    "계정"		, 20,,,50,1	
		ggoSpread.SSSetEdit		C_W4_TITLE,	    "계정"		, 20,,,50,1	
		ggoSpread.SSSetEdit		C_W5_TITLE,	    "계정"		, 20,,,50,1	
		' 그리드 헤더 합침 

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W1_TITLE,C_W5_TITLE,True)
		
		Call SetSpreadLock(TYPE_1)
			
		Call SetHeader()   
	
		.ReDraw = true	
			
	End With 
	
   	With lgvspdData(TYPE_2)
 
	' 2번 그리드 
	
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_COL15 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    

		'헤더를 2줄로    
		.ColHeaderRows = 1  
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","3","2")

		ggoSpread.SSSetEdit		C_COL1,	    ""		, 55,,,30,1	' 
		ggoSpread.SSSetEdit     C_COL2,	     ""		, 5,,,150,1	' 
		ggoSpread.SSSetEdit  	C_COL3,		 ""		, 8,,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL4,		 ""		, 3,,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL5,		""		, 8,,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL6,		 ""		, 8,,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL7,		 ""		, 3,,,50,1	' ' 
		ggoSpread.SSSetEdit	    C_COL8,		 ""		, 8,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL9,		 ""		, 8,,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL10,	 ""		, 3,,,50,1	' ' 
		ggoSpread.SSSetEdit	    C_COL11,	 ""		, 8,,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL12,		""	, 8,,,50,1	' 
    	ggoSpread.SSSetEdit	    C_COL13,		""	, 3,,,50,1	' ' 
		ggoSpread.SSSetEdit	    C_COL14,		""	, 8,,,50,1	' 
		ggoSpread.SSSetEdit	    C_COL15,		""	, 15,,,15,1	' 



		' 그리드 헤더 합침 
		ret = .AddCellSpan(C_COL1 , -1000, 15, 1)	'

		
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_COL1	: .Text = "연구및인력개발소요자금 상당액(24)중 환입할 금액"
		
	
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
				
		Call SetSpreadLock(TYPE_2)
					
		.ReDraw = true	
			
	End With     
End Sub

Function SetHeader()
   Dim strW1T, strW2T, strW3T, strW4T, strW5T

	With lgvspdData(TYPE_1)
	     ggoSpread.Source = lgvspdData(TYPE_1)
	  if .Maxrows <= 0 then

	     ggoSpread.InsertRow 1
	  end if  
	     .col = 0 
	     .Row = 1 
	     .text =""
	    
	    
			.Col =  C_ACCT : .Row = 1 :.value = "계정"  : .TypeHAlign = 2
			.Col =  C_W1 : .Row = 1 :.CellType = 1:.value = "(1)인건비" 
			.Col =  C_W2 : .Row = 1 :.CellType = 1:.value = "(2)재료비"
			.Col =  C_W3 : .Row = 1 :.CellType = 1 : .Text = "(3)위탁개발비"
			.Col =  C_W4 : .Row = 1 :.CellType = 1 : .Text = "(4)위탁훈련비"
			.Col =  C_W5 : .Row = 1 :.CellType = 1 : .Text = "(5)기타"
			.Col =  C_W6 : .Row = 1 :.CellType = 1 : .Text = "(6)계"
			 ggoSpread.SSSetProtected C_ACCT, 1, 1

			 ggoSpread.SpreadLock C_W1, 1, C_W6,1
			 
			
	   if .Maxrows > 1 then		  
		
			.Col =  C_W1_Title : .Row = 2 :.CellType =1 : strW1T = .text
			
			.Col =  C_W2_Title : .Row = 2 :.CellType =1 :strW2T = .text
			.Col =  C_W3_Title : .Row = 2 :.CellType = 1 :strW3T = .text
			.Col =  C_W4_Title : .Row = 2 :.CellType = 1:strW4T = .text
			.Col =  C_W5_Title : .Row = 2 :.CellType =1 :strW5T = .text
	
			
		'msgbox strW1T
			'.Col =  C_W1 : .Row = 1 :.Text  = strW1T
			'.Col =  C_W2 : .Row = 1 :.Text  = strW2T
			'.Col =  C_W3 : .Row = 1 : .Text = strW3T
			'.Col =  C_W4 : .Row = 1 :.Text = strW4T
			'.Col =  C_W5 : .Row = 1 : .Text = strW5T
			'.Col =  C_W6 : .Row = 1 : .Text = "(계)"
			 ggoSpread.SSSetProtected C_ACCT, 1, 1
			 ggoSpread.SSSetProtected C_W6, 1, 1
			
	  end if	 

	     
	End  With
end Function
'============================================  그리드 함수  ====================================

Sub InitData()
dim iMaxRows
dim sCoCd, 	sFiscYear ,sRepType
    sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
        '현재사업일 
        Call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear  & "' AND REP_TYPE='" & sRepType  & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
  		lgFiscStartDt = CDate(lgF0)
  		lgFiscEndDt = CDate(lgF1)

        if sRepType = 2 then  '중간 예납인 경우 2로 
           frm1.txtyearMth.value  = 6
        else
            frm1.txtyearMth.value =  DateDiff("m", CDate(lgF0), CDate(lgF1)) + 1
        end if   
     
	iMaxRows = 14 '

	With frm1.vspdData1
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData1
		.maxrows  = iMaxRows

		.Redraw = True

		Call InitData2()
	  
	End With	
 
End Sub



Sub InitData2()
Dim iRow ,ret
Dim sFiscYear, sRepType, sCoCd
    sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value

	With frm1.vspdData1
		.Redraw = False
         
         
        
           ret = .AddCellSpan(C_COL1, 1 ,1, 2)	    ' 직전4년간 발생합계액 
           ret = .AddCellSpan(C_COL3, 2 ,3, 1)	    ' 기간1
		   ret = .AddCellSpan(C_COL6, 2 ,3, 1)		' 기간2
		   ret = .AddCellSpan(C_COL9, 2 ,3, 1)		' 기간3
		   ret = .AddCellSpan(C_COL12, 2 ,3, 1)	  	' 기간4
		   		   
	
		   ret = .AddCellSpan(C_COL2, 3 ,15, 1)	    ' (13)x (48/직전 4년간 사업연도 월수) * (1/4) * (당해연도 월수 /12)
		   ret = .AddCellSpan(C_COL2, 4 ,15, 1)	    ' (13)x (48/직전 4년간 사업연도 월수) * (1/4) * (당해연도 월수 /12)
		   ret = .AddCellSpan(C_COL1, 5 ,2, 1)		' 증가발생금액 
		   ret = .AddCellSpan(C_COL3, 5 ,15, 1)	    
		   ret = .AddCellSpan(C_COL1, 6 ,15, 1)	     '공제세액 
		   
		   ret = .AddCellSpan(C_COL2, 7 ,4, 1)	     '대상금액(7,10) 
		   ret = .AddCellSpan(C_COL6, 7 ,3, 1)	     '공제율 
		   ret = .AddCellSpan(C_COL9, 7 ,3, 1)	     '공제세액 
		   ret = .AddCellSpan(C_COL12, 7 ,4, 1)	     '비고 
		   
		   ret = .AddCellSpan(C_COL2, 8 ,4, 1)	     '대상금액(7,10) 
		   ret = .AddCellSpan(C_COL6, 8 ,3, 1)	     '공제율 
		   ret = .AddCellSpan(C_COL9, 8 ,3, 1)	     '공제세액 
		   ret = .AddCellSpan(C_COL12,8 ,4, 1)	     '비고 
		   
		   ret = .AddCellSpan(C_COL2, 9 ,4, 1)	     '대상금액(7,10) 
		   ret = .AddCellSpan(C_COL6, 9 ,3, 1)	     '공제율 
		   ret = .AddCellSpan(C_COL9, 9 ,3, 1)	     '공제세액 
		   ret = .AddCellSpan(C_COL12,9 ,4, 1)	     '비고 
		   
	   	   ret = .AddCellSpan(C_COL1, 10 ,1, 2)	     '당해년도에 공제받을 세액 
	   	   ret = .AddCellSpan(C_COL2, 10 ,7, 1)	     '중소기업(15)과(16)중 선택 
	   	   ret = .AddCellSpan(C_COL9, 10 ,3, 1)	     '공제세액 
	   	   ret = .AddCellSpan(C_COL12,10 ,4, 1)	     '비고 
	   	   ret = .AddCellSpan(C_COL2, 11 ,7, 1)	     '중소기업외 기업(16)
	   	   ret = .AddCellSpan(C_COL9, 11 ,3, 1)	     '공제세액 
	   	   ret = .AddCellSpan(C_COL12,11 ,4, 1)	     '비고 
	   	  
	   	   ret = .AddCellSpan(C_COL1,12 ,1, 3)	     '(18)석박사급 인건비 관련 세액공제 
	   	   ret = .AddCellSpan(C_COL2,12 ,7, 1)	     '석박사급인건비 
	   	   ret = .AddCellSpan(C_COL9,12 ,3, 1)	     '공제세액 
	   	   ret = .AddCellSpan(C_COL12,12 ,4, 1)	     '비고 
	   	   ret = .AddCellSpan(C_COL2,13 ,7, 1)	     '연구 및 인력개발비 당기 발생세액 
	   	   ret = .AddCellSpan(C_COL9,13 ,3, 1)	     '공제세액 
	   	   ret = .AddCellSpan(C_COL12,13 ,4, 1)	     '비고 
	   	   ret = .AddCellSpan(C_COL2,14 ,7, 1)	     '석박사급 인건비에 대한 세액공제 
	   	   ret = .AddCellSpan(C_COL9,14 ,3, 1)	     '공제세액 
	   	   ret = .AddCellSpan(C_COL12,14 ,4, 1)	     '비고 
	   	   
		
		'1라인 
		
		iRow = 0
		iRow = iRow + 1 : .Row = iRow
		.Col = C_COL1	: .Text = "(8)직전4년간 " & vbCr & " 발생합계액": .TypeEditMultiLine = true : .typevalign = 2
		
	    iRow = 0
		iRow = iRow + 1 : .Row = iRow
		.Col = C_COL2	: .value = "기간"
		.Col = C_COL3	: .CellType = 0	
		.Col = C_COL5	: .CellType = 0	
		.Col = C_COL6	: .CellType = 0	
		.Col = C_COL8	: .CellType = 0	
		.Col = C_COL9	: .CellType = 0	
		.Col = C_COL11	: .CellType = 0	
		.Col = C_COL12	: .CellType = 0	
		.Col = C_COL14	: .CellType = 0	
		
		 iRow = 0
		 iRow = iRow + 1: .Row = iRow
		'.Col = C_COL3	: .TEXT =  & "-01-01"	
		.Col = C_COL4	: .value = "~"
		'.Col = C_COL5	: .TEXT = "2002-12-31"
		'.Col = C_COL6	: .TEXT = "2002-12-31"	
		.Col = C_COL7	: .value = "~"
		'.Col = C_COL8	: .TEXT = "2002-12-31"
		'.Col = C_COL9	: .TEXT = "2001-12-31"	
		.Col = C_COL10	: .value = "~"
		'.Col = C_COL11	: .TEXT = "2003-12-31"	
		'.Col = C_COL12	: .TEXT = "2002-12-31"
		.Col = C_COL13	: .value = "~"
		'.Col = C_COL14	: .TEXT = "2004-12-31"	
		.Col = C_COL15	: .value = "합계"
		
		
		'2라인 
		
		iRow = 0
		iRow = iRow + 2 : .Row = iRow
		.Col = C_COL2	: .value = "금액"
    
		  ggoSpread.SSSetFloat     C_COL3,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
		  ggoSpread.SSSetFloat     C_COL6,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
		  ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
		  ggoSpread.SSSetFloat     C_COL12,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
		  ggoSpread.SSSetFloat     C_COL15,	    ""	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow

		  	
		'3라인   
		iRow = 0
		iRow = iRow + 3 : .Row = iRow
		.Col = C_COL1	: .value = "(9)직전4년간 " & vbCr & " 연평균발생액" : .TypeEditMultiLine = true : .typevalign = 2
	     .rowheight(iRow) = 20	
	     .Col = C_COL2	: .TypeHAlign = 2 : .typevalign = 2 	: .TypeEditMultiLine = true : .text = " (8) x (48/직전 4년간 사업연도 월수) * (1/4) * (당해연도 월수 /12) "
	
		'4라인 
		iRow = 0
		iRow = iRow + 4 : .Row = iRow   
		ggoSpread.SSSetFloat     C_COL1,	    ""	, 13, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
	    .Col = C_COL2	: .TypeHAlign = 2 : .typevalign = 2 	: .TypeEditMultiLine = true
	    '5라인 
		iRow = 0
		iRow = iRow + 5 : .Row = iRow
		.Col = C_COL1	: .value = "(10)증가발생금액((7)-(9))"   
         ggoSpread.SSSetFloat     C_COL3,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow

		
		'6라인 
		iRow = 0
		iRow = iRow + 6 : .Row = iRow
		.Col = C_COL1	: .value = "공   제   세   액"  :  .TypeHAlign = 2
			  
		'7라인 
		iRow = 0
		iRow = iRow + 7 : .Row = iRow
		.Col = C_COL1	: .value = "구분"  :  .TypeHAlign = 2
		.Col = C_COL2	: .value = "(11)대상금액((7),(10))"						: .TypeEditMultiLine = true : .typevalign = 2
		.Col = C_COL6	: .value = "(12)공제율"									: .TypeEditMultiLine = true : .typevalign = 2
		.Col = C_COL9	: .value = "(13)공제세액((11)x(12))"					: .TypeEditMultiLine = true : .typevalign = 2
		.Col = C_COL12	: .value = "(14)비   고"								: .TypeEditMultiLine = true : .typevalign = 2

		.Col = C_COL3	:  .TypeHAlign = 1 
		 
		'8라인 
		iRow = 0
		iRow = iRow + 8 : .Row = iRow
		.rowheight(iRow) = 20	
		.typevalign = 2
		
		.Col = C_COL1	: .value = "(15)당해연도" & vbCr & " 총발생금액 공제"	: .TypeEditMultiLine = true : .typevalign = 2
		
		
		'공제율 
		  call CommonQueryRs("REFERENCE_1,REFERENCE_2"," dbo.ufn_TB_Configuration('W4004', '" & C_REVISION_YM & "') "," Minor_cd= '3' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
  
         .Col = C_COL6	: .text =  replace(lgF1,Chr(11),"") : .TypeEditMultiLine = true : .typevalign = 2 : .typehalign = 2
          frm1.txt15_12Value.value =  replace(lgF0,Chr(11),"")
    
    
	    ggoSpread.SSSetFloat     C_COL2,	    ""	, 4, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
	    ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,,,iRow
	
	    
		'9라인 
		iRow = 0
		iRow = iRow + 9 : .Row = iRow
		.rowheight(iRow) = 20	
		
		.Col = C_COL1	: .value = "(16)증가발생" & vbCr & "  금액 공제" : .TypeEditMultiLine = true : .typevalign = 2
		
		'공제율 
		
		  call CommonQueryRs("REFERENCE_1,REFERENCE_2, COMP_TYPE1"," dbo.ufn_TB_Configuration('W4004', '" & C_REVISION_YM & "') , TB_COMPANY_HISTORY", " CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear  & "' AND REP_TYPE='" & sRepType  & "' and   Minor_cd=   ( Case   when  COMP_TYPE1 = 1 then 2 else 1 end ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
          frm1.txt16_12Value.value =  replace(lgF0,Chr(11),"")   

         .Col = C_COL6	: .text =   replace(lgF1,Chr(11),"") : .TypeEditMultiLine = true : .typevalign = 2 : .typehalign = 2
          frm1.txtCompType.value =  replace(lgF2,Chr(11),"")   ' 중소기업여부 1: 일반 2:중소기업 
          
          
          
		.TypeVAlign = 2
		ggoSpread.SSSetFloat     C_COL2,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
	    ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,,,iRow
	    
	
          
		'10라인 
		iRow = 0
		iRow = iRow + 10 : .Row = iRow
		.Col = C_COL1	: .value = "(17)당해년도에" & vbCr & " 공제받을 세액" : .TypeEditMultiLine = true : .typevalign = 2
		.Col = C_COL2	: .value = "중소기업{(15)과(16) 중 선택}" : .TypeEditMultiLine = true : .typevalign = 2
		 ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
		   
		'11라인 
		iRow = 0
		iRow = iRow + 11 : .Row = iRow
		.Col = C_COL2	: .value = "중소기업 외{(16)}" : .TypeEditMultiLine = true : .typevalign = 2
		 ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,,,iRow
		 

		
		
		'12라인 
		iRow = 0
		iRow = iRow + 12 : .Row = iRow
		.Col = C_COL1	: .value = "(18)석박사급  " & vbCr & " 인건비 " & vbCr & " 관련 세액공제" : .TypeEditMultiLine = true : .typevalign = 2
		.Col = C_COL2	: .value = "석박사급인건비" : .TypeEditMultiLine = true : .typevalign = 2
		 ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z",,,iRow
		
		 
		 '13라인 
		iRow = 0
		iRow = iRow + 13 : .Row = iRow
		.Col = C_COL2	: .value = "연구 및 인력개발비 당기 발생액" : .TypeEditMultiLine = true : .typevalign = 2
		 ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,,,iRow
		 
		 
		  '14라인 
		iRow = 0
		iRow = iRow + 14 : .Row = iRow
		.Col = C_COL2	: .value = "석박사급 인건비에 대한 세액공제" : .TypeEditMultiLine = true : .typevalign = 2
		 ggoSpread.SSSetFloat     C_COL9,	    ""	, 8, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,,,iRow
		   

		SetSpreadColor TYPE_2,   -1,-1
		 .Row = 12	: .RowHidden = True
		 .Row = 13	: .RowHidden = True
		 .Row = 14	: .RowHidden = True
		
		.Redraw = true	
	End With 
	
	 
End Sub


Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	Select Case pType
		Case TYPE_1
			ggoSpread.SSSetRequired C_ACCT, -1, C_ACCT
			ggoSpread.SpreadLock C_W6, -1, C_W6
			ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
			ggoSpread.SpreadLock C_W1, 1, C_W6,1
	
		
		Case TYPE_2
			'ggoSpread.SSSetRequired C_W9, -1, C_W9
			'ggoSpread.SpreadLock C_W22, -1, C_W22
			'ggoSpread.SpreadLock C_W23, -1, C_W23
			'ggoSpread.SpreadLock C_W25, -1, C_W25
	End Select

	
	
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(pType)
	
	Select Case pType
		Case TYPE_1
			

		    
               if  lgvspdData(pType).row <>   lgvspdData(pType).maxrows   or   lgvspdData(pType).row <> 1 then
			  
		            ggoSpread.SSSetRequired C_ACCT, 2, pvEndRow
		        
	           end if	 

		       ggoSpread.SSSetProtected C_ACCT, lgvspdData(pType).maxrows, lgvspdData(pType).maxrows
		       ggoSpread.SSSetProtected C_W1, lgvspdData(pType).maxrows, lgvspdData(pType).maxrows
		       ggoSpread.SSSetProtected C_W2, lgvspdData(pType).maxrows, lgvspdData(pType).maxrows
		       ggoSpread.SSSetProtected C_W3, lgvspdData(pType).maxrows, lgvspdData(pType).maxrows
		       ggoSpread.SSSetProtected C_W4, lgvspdData(pType).maxrows, lgvspdData(pType).maxrows
		       ggoSpread.SSSetProtected C_W5, lgvspdData(pType).maxrows, lgvspdData(pType).maxrows
		       ggoSpread.SSSetProtected C_W6, -1, -1
		
		Case TYPE_2
           With lgvspdData(pType)  
           
        
		
				 ggoSpread.SSSetProtected C_COL1, pvStartRow, pvEndRow 	
		 		 ggoSpread.SSSetProtected C_COL2, pvStartRow, pvEndRow 	
		 		 ggoSpread.SSSetProtected C_COL6, 3, 11	
		 		 ggoSpread.SSSetProtected C_COL3, 3, 11
		 		 ggoSpread.SSSetProtected C_COL9, 3, 11  
		 		 ggoSpread.SSSetProtected C_COL12, 7, 7  
		 		 ggoSpread.SSSetProtected C_COL15, 1, 1  
		 
		 		 ggoSpread.SSSetProtected C_COL4, 1, 1 
		 		 ggoSpread.SSSetProtected C_COL7, 1, 1 
		 		 ggoSpread.SSSetProtected C_COL10, 1, 1 
		 		 ggoSpread.SSSetProtected C_COL13, 1, 1 
		 		 
		 		 'ggoSpread.SSSetRequired C_COL3, 1, 1
		 		 'ggoSpread.SSSetRequired C_COL5, 1, 1	
		 		 'ggoSpread.SSSetRequired C_COL6, 1, 1	
		 		 'ggoSpread.SSSetRequired C_COL8, 1, 1
		 		 'ggoSpread.SSSetRequired C_COL9, 1, 1  
		 		 'ggoSpread.SSSetRequired C_COL11, 1, 1 
		 		 'ggoSpread.SSSetRequired C_COL12, 1, 1 
		 		 'ggoSpread.SSSetRequired C_COL14, 1, 1 
		 		 
		 		 
		 		 ggoSpread.SSSetProtected C_COL15, 1, 2 
		 		 ggoSpread.SSSetProtected C_COL9, 13, 14 
		 		 

		
				'.Col =  C_COL1	: .Row=4  :.BackColor  = &H00F9E8D1&
				'.Col =  C_COL3	: .Row=5  :.BackColor  = &H00F9E8D1&
				'.Col =  C_COL2	: .Row=8  :.BackColor  = &H00F9E8D1&
				'.Col =  C_COL2	: .Row=9  :.BackColor  = &H00F9E8D1&
				'.Col =  C_COL9	: .Row=8  :.BackColor  = &H00F9E8D1&
				'.Col =  C_COL9	: .Row=9  :.BackColor  = &H00F9E8D1&
				'.Col =  C_COL9	: .Row=11 :.BackColor  = &H00F9E8D1&
				'.Col =  C_COL9	: .Row=12 :.BackColor  = &H00F9E8D1&
				  
			
			end With

	End Select
	
End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_1 To TYPE_1
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_ACCT: .CellType = 1	: .Text = "(7)계"	: .TypeHAlign = 2
				ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
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

'============================== 레퍼런스 함수  ========================================

Function GetRef()	' 금액가져오기 링크 클릭시 


    Dim sFiscYear, sRepType, sCoCd, IntRetCD  ,sCOL
    Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol ,dblW18
    Dim sMesg

	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	    
	
	

	' 세무정보 조사 : 메시지가져오기.
	
	
	if wgConfirmFlg = "Y" then    '확정시 
	   Exit function
	end if   
	
	' 온로드시 레퍼런스메시지 가져온다.
     wgRefDoc = GetDocRef(sCoCd,sFiscYear, sRepType, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	  
    '
    
	With frm1.vspdData1
			.Row = 1 :.Col = C_COL3 : .text  = "" 
			.Row = 1 :.Col = C_COL5 : .text  = "" 
			.Row = 1 :.Col = C_COL6 : .text  = "" 
			.Row = 1 :.Col = C_COL8 : .text  = "" 
			.Row = 1 :.Col = C_COL9 : .text  = "" 
			.Row = 1 :.Col = C_COL11 : .text  = "" 
			.Row = 1 :.Col = C_COL12 : .text  = "" 
			.Row = 1 :.Col = C_COL14 : .text  = ""
	       
			.Row = 2 :.Col = C_COL3 : .text  = 0 
			.Row = 2 :.Col = C_COL6 : .text  = 0
			.Row = 2 :.Col = C_COL9 : .text  = 0 
			.Row = 2 :.Col = C_COL12 : .text  = 0 
	       
	       '전기 조특3호 
			
			 	Call CommonQueryRs(" W8_D1_S, W8_D1_E , W8_AMT1 ","dbo.ufn_TB_JT3A_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & 1 & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
				If replace(lgF0, chr(11), "")  <>  "" Then	 
				    	arrW1 = Split(lgF0, chr(11))
						arrW2 = Split(lgF1, chr(11))
						arrW3 = Split(lgF2, chr(11))
	
						.Redraw = False
						lgIntFlgMode = parent.OPMD_UMODE
						lgBlnFlgChgValue = True
						ggoSpread.Source = frm1.vspdData1
	
						 .Row = 1
						.Col = C_COL3 : .text  = arrW1(0)
						  Call vspdData1_Change(C_COL3 , 1)
						 .Row = 1   
						.Col = C_COL5 : .text  = arrW2(0)
						 Call vspdData1_Change(C_COL5 , 1 )
						 .Row = 2   
						.Col = C_COL3: .text  = arrW3(0)
						 Call vspdData1_Change(C_COL3 , 2 )
						 sCOL =  C_COL6
						 					
				end if		 
				   
				lgF0 = ""
			    lgF2 = ""
			    lgF3 = ""
			    lgF4 = ""
				
				
				
				Call CommonQueryRs(" W8_D2_S, W8_D2_E , W8_AMT2 ","dbo.ufn_TB_JT3A_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & 1 & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			    sCOL =  Split(lgF2, chr(11))

				If replace(lgF0, chr(11), "") <> "" Then	 
				    

						 arrW1 = Split(lgF0, chr(11))
						arrW2 = Split(lgF1, chr(11))
						arrW3 = Split(lgF2, chr(11))     
					     .Row = 1  
						 .Col = C_COL6 : .text  = arrW1(0)
						 Call vspdData1_Change(C_COL6 , 1)
						 .Row = 1   
						.Col = C_COL8 : .text  = arrW2(0)
						 Call vspdData1_Change(C_COL8 , 1 )
						 .Row = 2   
						.Col = C_COL6: .text  = arrW3(0)
						 Call vspdData1_Change(C_COL6 , 2)
						 sCOL =  C_COL9
						
			   END IF
			    lgF0 = ""
			    lgF2 = ""
			    lgF3 = ""
			    lgF4 = ""
			   
			    
			    Call CommonQueryRs(" W8_D3_S, W8_D3_E , W8_AMT3 ","dbo.ufn_TB_JT3A_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & 1 & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

				If replace(lgF0, chr(11), "")  <> "" Then	 
	
						arrW1 = Split(lgF0, chr(11))
						arrW2 = Split(lgF1, chr(11))
						arrW3 = Split(lgF2, chr(11))     
					       .Row = 1 
						 .Col = C_COL9 : .text  = arrW1(0)
						 Call vspdData1_Change(C_COL9 , 1)
						 .Row = 1   
						.Col = C_COL11 : .text  = arrW2(0)
						 Call vspdData1_Change(C_COL11 , 1 )
						 .Row = 2   
						.Col = C_COL9: .text  = arrW3(0)
						 Call vspdData1_Change(C_COL9 , 2 )
						  sCOL =  C_COL12
						
						
				   END IF	
				   
				    lgF0 = ""
					lgF2 = ""
					lgF3 = ""
					lgF4 = ""
			  
			  Call CommonQueryRs(" W8_D4_S, W8_D4_E , W8_AMT4 ","dbo.ufn_TB_JT3A_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & 1 & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
 
				If replace(lgF0, chr(11), "")  <>  "" Then	 
				    
						arrW1 = Split(lgF0, chr(11))
						arrW2 = Split(lgF1, chr(11))
						arrW3 = Split(lgF2, chr(11))     
						 
						   .Row = 1 
						 .Col = sCOL : .text  = arrW1(0)
						 Call vspdData1_Change(sCOL , 1)
						 .Row = 1   
						.Col = unicdbl(sCOL) + 2 : .text  = arrW2(0)
						 Call vspdData1_Change(sCOL + 2 , 1 )
						 .Row = 2   
						.Col = sCOL: .text  = arrW3(0)
						 Call vspdData1_Change(sCOL , 1 )
						
				else
				 		 .Row = 1 :.Col = sCOL	: .Text = DateAdd("yyyy",-1, lgFiscStartDt)
						 .Row = 1 :.Col = unicdbl(sCOL) + 2	:.Text = DateAdd("yyyy",-1, lgFiscEndDt)
			
				end if		 
				     
	
					
		.Redraw = True
	End With
	
End Function




Sub Fn_SumCal()	
	  Dim  dtDayS ,dtDayE, i4yearMth,iyearMth , strFISC_START_DT, strFISC_End_DT
	  Dim  dtDayGap1 ,dtDayGap2, dtDayGap3, dtDayGap4,dblW7,dblW9,dblW8,dblW10,dblW15_13,dblW16_13,dblW18_a,dblW18_b, dblW18_c 
	  Dim  sCoCd, sFiscYear, sRepType
	  
	    sCoCd		= frm1.txtCO_CD.value
		sFiscYear	= frm1.txtFISC_YEAR.text
		sRepType	= frm1.cboREP_TYPE.value    
	

	   '현재사업일 
		Call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear  & "' AND REP_TYPE='" & sRepType  & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
             
			 
           frm1.txtyearMth.value =  DateDiff("m", CDate(lgF0), CDate(lgF1)) + 1
         
	  With frm1.vspdData1
	        ggoSpread.Source = frm1.vspdData1
	        '*****4년간 월수 
					.Row = 1 :.Col = C_COL3	: dtDayS = .text 
					.Row = 1 :.Col = C_COL5	: dtDayE = .text 
	       
					         
					if Trim(dtDayS) <> "" and  Trim(dtDayE) <> ""  then
						          dtDayGap1 = DateDiff("m", CDate(replace(dtDayS,"",0)), CDate(replace(dtDayE,"",0))) + 1
					Else
						dtDayGap1 = 0
					end if  
					   
					 	 
					.Row = 1 :.Col = C_COL6	: dtDayS = .text 
					.Row = 1 :.Col = C_COL8	: dtDayE = .text 
					 if Trim(dtDayS)<> "" and  Trim(dtDayE) <> ""  then 
					    dtDayGap2 = DateDiff("m", CDate(replace(trim(dtDayS),"",0)), CDate(replace(trim(dtDayE),"",0))) + 1
					Else
						dtDayGap2 = 0
					 end if
	       
	  
					.Row = 1 :.Col = C_COL9	: dtDayS = .text 
					.Row = 1 :.Col = C_COL11	: dtDayE = .text 
					if Trim(dtDayS)<> "" and  Trim(dtDayE) <> ""  then 	
						 dtDayGap3 = DateDiff("m", CDate(dtDayS), CDate(dtDayE)) + 1
					Else
						dtDayGap3 = 0
					end if	 
					 
							.Row = 1 :.Col = C_COL12	: dtDayS = .text 
							.Row = 1 :.Col = C_COL14	: dtDayE = .text 
					if Trim(dtDayS) <> "" and  Trim(dtDayE) <> ""  then 		 		
							 dtDayGap4 = DateDiff("m", CDate(dtDayS), CDate(dtDayE)) + 1
					Else
						dtDayGap4 = 0
					 end if	 
					 
	      '*****4년간 월수 
				 i4yearMth =  unicdbl(dtDayGap1)+unicdbl(dtDayGap2)+unicdbl(dtDayGap3)+unicdbl(dtDayGap4)
				 iyearMth =   frm1.txtyearMth.value  '당기 월수 
				 frm1.txt4yearMth.value = i4yearMth
	            .Row = 3 :.Col = C_COL2	:  .text  = "(8) x (48/직전 4년간의 사업연도 월수) x (1/4) x (당해년도 월수/12) " 
				.Row = 4 :.Col = C_COL2	:  .text  = "(8) x  (48/"& i4yearMth &")  x (1/4) x ("& iyearMth &"/12)" : .TypeHAlign = 2 : .typevalign = 2 	
	      '*********2행 금액 합계 재계산 
	      
		        Call FncSumSheet(lgvspdData(lgCurrGrid), 2, C_COL3, C_COL14, true, 2 , C_COL15, "H")	' 합계 
		      
		  '*********(9)직전4년간 연평균 발생액    
		  
		      if frm1.txt4yearMth.value = 0 then
		           .Row = 4 :.Col = C_COL1	: .text =0
			     
			  else  
				   .Row = 2 :.Col = C_COL15	:  dblW8 = .text     
                   .Row = 4 :.Col = C_COL1	: .text  = fix(unicdbl(dblW8 * (48/UNICDbl(frm1.txt4yearMth.value)) * (1/4) * (UNICDbl(frm1.txtyearMth.value)/12)))
              end if     
	               
	      end  With   
	  
	     '********(10) 증가발생액 
	  
	                ggoSpread.Source = frm1.vspdData0
	                if  frm1.vspdData0.maxrows  > 1 then
	                    frm1.vspdData0.Row =  frm1.vspdData0.maxrows : frm1.vspdData0.Col =  C_W6 :  dblW7 = frm1.vspdData0.text   
	                end if    
	                ggoSpread.Source = frm1.vspdData1
	                frm1.vspdData1.Row = 4 :frm1.vspdData1.Col = C_COL1	:  dblW9 =  frm1.vspdData1.text  
	             
	                dblW10 = unicdbl(dblW7) - unicdbl(dblW9) 
	                'if dblW10 < 0 then
	                '   dblW10 = 0
	                'end if
	                 frm1.vspdData1.Row = 5 :frm1.vspdData1.Col = C_COL3	:  frm1.vspdData1.text  =  unicdbl(dblW10)
	                
	     '********(15) 당해연도 총 발생 금액공제        
	                frm1.vspdData1.Row = 8 :frm1.vspdData1.Col = C_COL2	:   frm1.vspdData1.text   =unicdbl(dblW7)
	               
	      
	     '********(16) 증가발생 금액공제           
	                frm1.vspdData1.Row =9 :frm1.vspdData1.Col = C_COL2	:   frm1.vspdData1.text   = unicdbl(dblW10)
	                
	                
	     '********(15-13)공제세액1           
	               frm1.vspdData1.Row =8 :frm1.vspdData1.Col = C_COL9	:   frm1.vspdData1.text   = Fix(unicdbl(frm1.txt15_12Value.value )*unicdbl(dblW7))
	               dblW15_13 =  frm1.vspdData1.text 
	           
	           
	     '********(16-13)공제세액1      
	               frm1.vspdData1.Row =9 :frm1.vspdData1.Col = C_COL9	:   frm1.vspdData1.text   = Fix(unicdbl(frm1.txt16_12Value.value )*unicdbl(dblW10))                    
                   dblW16_13 =  frm1.vspdData1.text 
         
         '********(17)당해년도에 공제받을 세액 

				   if frm1.txtCompType.value  = 2 then     '중소기업일경우 
				        frm1.vspdData1.Row =11 :frm1.vspdData1.Col = C_COL9	:   frm1.vspdData1.text   =0
		
						if  unicdbl(dblW16_13) > unicdbl(dblW15_13) then
							
						    frm1.vspdData1.Row =10 :frm1.vspdData1.Col = C_COL9	:   frm1.vspdData1.text   = dblW16_13
						else
						    frm1.vspdData1.Row =10 :frm1.vspdData1.Col = C_COL9	:   frm1.vspdData1.text   = dblW15_13
						end if
				   else
						 frm1.vspdData1.Row =10 :frm1.vspdData1.Col = C_COL9	:   frm1.vspdData1.text   =0
				         frm1.vspdData1.Row =11 :frm1.vspdData1.Col = C_COL9	:   frm1.vspdData1.text   = dblW16_13
                        
				   end if
		
		
		'********(18)석박사급 인건비 관려세액공제 
                    
				     ggoSpread.Source = frm1.vspdData0
	                if  frm1.vspdData0.maxrows  > 1 then
	                    frm1.vspdData0.Row =  frm1.vspdData0.maxrows : frm1.vspdData0.Col =  C_W6 :  dblW18_b = frm1.vspdData0.text   
	                end if    
	                ggoSpread.Source = frm1.vspdData1
	            
	                 frm1.vspdData1.Row = 13 :frm1.vspdData1.Col = C_COL9	:  frm1.vspdData1.text  =  unicdbl(dblW18_b)
	                 
	   ' ********(18)석박사급 인건비에대한 세액공제((17)중소기업외의 기업 * (18) 석박사급 인건비 / (18) 연구및 인력개발비 당기발생액 
                   
	                 ggoSpread.Source = frm1.vspdData1
	                 frm1.vspdData1.Row = 11 :frm1.vspdData1.Col = C_COL9	:  dblW16_13   = frm1.vspdData1.text     '중소기업 
	         
	                 frm1.vspdData1.Row = 12 :frm1.vspdData1.Col = C_COL9	:  dblW18_a    = frm1.vspdData1.text 
	                 frm1.vspdData1.Row = 13 :frm1.vspdData1.Col = C_COL9	:  dblW18_b    = frm1.vspdData1.text 
	             
	                 if unicdbl(dblW18_b) = 0 then
	                    dblW18_c = 0
	                 else
	                    dblW18_c =unicdbl(dblW16_13) * unicdbl(dblW18_a)  /  unicdbl(dblW18_b)
	                  end if   
	                 
                     frm1.vspdData1.Row = 14 :frm1.vspdData1.Col = C_COL9	:  frm1.vspdData1.text  =  unicdbl(dblW18_c) 
	       
	 
End Sub


 function Fn_CompanyYYMMDD()

  Dim sFiscYear, sRepType, sCoCd, iGap, IntRetCD
  dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value

    With frm1.vspdData1
		.Redraw = False
        ggoSpread.Source = frm1.vspdData1
        
        'frm1.cboREP_TYPE.value
        '현재사업일 
		
		Call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear  & "' AND REP_TYPE='" & sRepType  & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
             
			 
           frm1.txtyearMth.value =  DateDiff("m", CDate(lgF0), CDate(lgF1)) + 1
  
		'직전연도사업일 
		'Call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear - 1 & "' AND REP_TYPE='1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

             
			 .Row = 1 :.Col = C_COL3	: .text =  replace(lgF0, Chr(11),"")	:  .Text = DateAdd("yyyy",-4, lgFiscStartDt)
			 .Row = 1 :.Col = C_COL5	: .text =  replace(lgF1, Chr(11),"")	: .Text = DateAdd("yyyy",-4, lgFiscEndDt)
			 
			 
       '-2연도사업일 
		'call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear - 2 & "' AND REP_TYPE='1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		     .Row = 1 :.Col = C_COL6	: .text =  replace(lgF0, Chr(11),"")	: .Text = DateAdd("yyyy",-3, lgFiscStartDt)
			 .Row = 1 :.Col = C_COL8	: .text =  replace(lgF1, Chr(11),"")	:  .Text = DateAdd("yyyy",-3, lgFiscEndDt)
	
		'-3연도사업일 
		'call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear - 3 & "' AND REP_TYPE='1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
		      .Row = 1 :.Col = C_COL9	: .text =  replace(lgF0, Chr(11),"")	: .Text = DateAdd("yyyy",-2, lgFiscStartDt)
			  .Row = 1 :.Col = C_COL11	: .text =  replace(lgF1, Chr(11),"")	: .Text = DateAdd("yyyy",-2, lgFiscEndDt)
	
		'-4연도사업일 
		'call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear - 4 & "' AND REP_TYPE='1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
		      .Row = 1 :.Col = C_COL12	: .text =  replace(lgF0, Chr(11),"")	: .Text = DateAdd("yyyy",-1, lgFiscStartDt)
			  .Row = 1 :.Col = C_COL14	: .text =  replace(lgF1, Chr(11),"")	: .Text = DateAdd("yyyy",-1, lgFiscEndDt)
	
		.Redraw = True
	end 	With
		Call Fn_SumCal()			
		
		
End function

		
'====================================== 탭 함수 =========================================

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
    Call InitVariables 
   
	Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
                                                     <%'Initializes local global variables%>

    Call SetToolbar("1101110100100111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
	Call InitComboBox
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
   
    Call InitData 
	'Call InitData2()
	
	Call Fn_CompanyYYMMDD
	
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


'============================================  그리드 이벤트   ====================================
' -- 0번 그리드 
Sub vspdData0_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_1
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Click(ByVal Col, ByVal Row)
	'lgCurrGrid = TYPE_1
	'Call vspdData_Click(lgCurrGrid,  Col,  Row)
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

Sub vspdData0_KeyDown(KeyCode, shift)
    lgCurrGrid = TYPE_1
    Call vspdData_KeyDown(lgCurrGrid, KeyCode, shift)
End Sub

' -- 1번 그리드 
Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	'lgCurrGrid = TYPE_2
	'Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_2
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData1_KeyDown(KeyCode, shift)
    lgCurrGrid = TYPE_2
    Call vspdData_KeyDown(lgCurrGrid, KeyCode, shift)
End Sub

'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, dblFiscYear
	Dim dblW8 ,dblW9, dblW10, dblW11, dblW12, dblW13, dblW14, dblW15, dblW16, dblW17, dblW18
	Dim dblW19, dblW20, dblW21, dblW22, dblW23, dblW24, dblW25,iRow
	Dim dblW1,dblW2,dblW3,dblW4, dblW5
	Dim sFiscYear, sRepType, sCoCd
	
	lgBlnFlgChgValue= True ' 변경여부 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

   
    
	' --- 추가된 부분 
	With lgvspdData(Index)

	If Index = TYPE_1 Then	'1번 그리 
	   
	      if Row <> 1 then
				Select Case Col
		
					Case C_W1, C_W2, C_W3, C_W4, C_W5, C_W6
					    .Col = C_W1	: .Row = Row	: dblW1 = UNICDbl(.Value)
						.Col = C_W2	: .Row = Row	: dblW2 = UNICDbl(.Value)
						.Col = C_W3	: .Row = Row	: dblW3 = UNICDbl(.Value)
						.Col = C_W4	: .Row = Row	: dblW4 = UNICDbl(.Value)
					
						
						.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
						
						
						If dblSum < 0 Then
							Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "금액", "X")           '⊙: "%1 금액이 0보다 적습니다."
							.Value = 0
						End If
						If row <> 1 then 
						  Call FncSumSheet(lgvspdData(lgCurrGrid), Row, C_W1, C_W5, true, Row , C_W6, "H")	' 합계 
						
						 
						end if  
						
						
						
						Call CheckReCalc 
						
				End Select
		  end if		
				
		 If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
			If uniCDbl(lgvspdData(Index).text) < uniCDbl(lgvspdData(Index).TypeFloatMin) Then
			   lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
			End If
	     End If
	         ggoSpread.Source = lgvspdData(Index)
		 	 ggoSpread.UpdateRow Row
		 	 ggoSpread.UpdateRow .Maxrows
		 	 
		 	 if Row = 1 then
	             for iRow = 1 to .maxrows
	                  ggoSpread.UpdateRow iRow
	             next 
	         end if
	   
	ElseIf Index = TYPE_2 Then
			Select Case Col
			     
			Case C_COL3, C_COL5 ' -- 직전 4년시작 
				If Row = 1 Then
					' -- 날짜비교 
					If UNICDbl(frm1.txtFISC_YEAR.Text) - UNICDbl(Year(.Text)) <> 4 Then
						Call DisplayMsgBox("X",  parent.VB_INFORMATION, "기간이 잘못되었습니다(직전 4년)", "X")   
						.Text = ""
						Exit Sub
					End If
					
			    end if 
				Call Fn_SumCal
			Case C_COL6, C_COL8
				If Row = 1 Then
					' -- 날짜비교 
					If UNICDbl(frm1.txtFISC_YEAR.Text) - UNICDbl(Year(.Text)) <> 3 Then
						Call DisplayMsgBox("X",  parent.VB_INFORMATION, "기간이 잘못되었습니다(직전 3년)", "X")   
						.Text = ""
						Exit Sub
					End If
					
			    end if 
				Call Fn_SumCal
			Case C_COL9, C_COL11
				If Row = 1 Then
					' -- 날짜비교 
					If UNICDbl(frm1.txtFISC_YEAR.Text) - UNICDbl(Year(.Text)) <> 2 Then
						Call DisplayMsgBox("X",  parent.VB_INFORMATION, "기간이 잘못되었습니다(직전 2년)", "X")   
						.Text = ""
						Exit Sub
					End If
					Call Fn_SumCal
			    Else
			        Call Fn_SumCal
			    end if 

			Case C_COL12, C_COL14
				If Row = 1 Then
					' -- 날짜비교 
					If UNICDbl(frm1.txtFISC_YEAR.Text) - UNICDbl(Year(.Text)) <> 1 Then
						Call DisplayMsgBox("X",  parent.VB_INFORMATION, "기간이 잘못되었습니다(직전 1년)", "X")   
						.Text = ""
						Exit Sub
					End If
					'Call Fn_SumCal
			    end if 
					Call Fn_SumCal
			End Select
			
			' -- 2행이 변경될 경우 
			If Row = 2 Then
				.Row = Row : .Col = Col
				If UNICDbl(.Value) > 0 Then
					' 0 보다 크면 프로텍트 푼다.
				Else
					' 0 이면 프로텍트로 바꾸고 디폴트 날짜를 박는다.
					.Row = Row - 1 : .Col = Col
					Select Case Col
						Case C_COL3
							.Text = DateAdd("yyyy",-4, lgFiscStartDt)
							.Col = C_COL5
							.Text = DateAdd("yyyy",-4, lgFiscEndDt)
						Case C_COL6
							.Text = DateAdd("yyyy",-3, lgFiscStartDt)
							.Col = C_COL8
							.Text = DateAdd("yyyy",-3, lgFiscEndDt)
						Case C_COL9
							.Text = DateAdd("yyyy",-2, lgFiscStartDt)
							.Col = C_COL11
							.Text = DateAdd("yyyy",-2, lgFiscEndDt)
						Case C_COL12
							.Text = DateAdd("yyyy",-1, lgFiscStartDt)
							.Col = C_COL14
							.Text = DateAdd("yyyy",-1, lgFiscEndDt)
					End Select
					Call Fn_SumCal
				End If
			End If
			
			lgBlnFlgChgValue = true
			
	End If
	
	End With
	
End Sub

' -- 2번째 그리드 
Sub SetGridTYPE_2()
	Dim dblW9, dblW10, dblW11, dblW12, dblW13, dblW14, dblW15, dblW16, dblW17, dblW18
	Dim dblW19, dblW20, dblW21, dblW22, dblW23, dblW24, dblW25

	With lgvspdData(TYPE_2)
		.Row = .ActiveRow
		.Col = C_W19 : dblW19 = UNICDbl(.Value)
		.Col = C_W20 : dblW20 = UNICDbl(.Value)
		.Col = C_W21 : dblW21 = UNICDbl(.Value)
									
		' 합계변경 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W19, 1, .MaxRows - 1, true, .MaxRows, C_W19, "V")	' 합계 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W20, 1, .MaxRows - 1, true, .MaxRows, C_W20, "V")	' 합계 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W21, 1, .MaxRows - 1, true, .MaxRows, C_W21, "V")	' 합계 
					
		' W22 변경 
		dblW22 = dblW19 + dblW20 + dblW21
		.Col = C_W22	: .Row = .ActiveRow : .Value = dblW22
					
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W22, 1, .MaxRows - 1, true, .MaxRows, C_W22, "V")	' 합계 
					
		' W23 변경 
		.Col = C_W17	: .Row = .ActiveRow : dblW17 = UNICDbl(.value)
		dblW23 = dblW17 + dblW22
		.Col = C_W23	: .Row = .ActiveRow : .Value = dblW23

		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W23, 1, .MaxRows - 1, true, .MaxRows, C_W23, "V")	' 합계 
		
		.Row = .ActiveRow			
		.Col = C_W24	: dblW24 = UNICDbl(.Value)
		' W25 변경 
		dblW25= dblW23 - dblW24
		.Col = C_W25	: .Value = dblW25
					
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W25, 1, .MaxRows - 1, true, .MaxRows, C_W25, "V")	' 합계	
	End With
End Sub

' 2번 그리드에서 1번 그리드의 데이타를 찾아서 W16금액을 리턴한다 
Sub GetW16(Byval pYear , Byref pdblW16, Byref pdblW17)
	Dim iRow, iMaxRows
	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows - 1
		.Col = C_W9
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			If UNICDbl(.Value) = pYear Then
				.Col = C_W16 : pdblW16 = UNICDbl(.Value)
				.Col = C_W17 : pdblW17 = UNICDbl(.Value)
				Exit Sub
			End If
		Next
		pdblW16 = -1 : pdblW17 = -1
	End With
End Sub




Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    Call SetPopupMenuItemInf("1101011111") 

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
    ggoSpread.Source = frm1.vspdData
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
    
    'if lgvspdData(Index).MaxRows < NewTop + VisibleRowCnt(lgvspdData(Index),NewTop) Then	           
   ' 	If lgStrPrevKeyIndex <> "" Then                         
   '   	   Call DisableToolBar(Parent.TBC_QUERY)
'			If DbQuery = False Then
	'			Call RestoreTooBar()
'			    Exit Sub
'			End If  				
 '   	End If
  '  End if
End Sub

Sub vspdData_KeyDown(Index, KeyCode, shift)
	With lgvspdData(Index)
		If KeyCode = 46 Then
			.Col = .ActiveCol
			.Row = .ActiveRow
			.Text = ""
		End If
	End With
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
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
    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	Call MakeKeyStream("X")
     
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

	    
    For i = TYPE_1 To TYPE_1
    
		ggoSpread.Source = lgvspdData(i) 
		If ggoSpread.SSCheckChange = False and lgBlnFlgChgValue = False then
			Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	Next
	
	
    ggoSpread.Source = frm1.vspdData0
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  

    Call Verification()

    'If Verification = False Then Exit Function

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

' ----------------------  검증 -------------------------
Function  Verification()
	Dim dblData, sTitle,Row
	
	Verification = False

	With lgvspdData(TYPE_1)

		if .maxRows =1 then exit Function
	    Row =  1
	      				       
		.Col = C_W6			: .Row = .MaxRows 	: dblData		= UNICDbl(.Value )
		if dblData <> 0 then
			.Col = C_W1	: .Row = Row	: sTitle	= Trim(.Value )
		
			if  sTitle= "" then
			   Call DisplayMsgBox("X", "X" , "첫번째항목의 구분과 비목을 입력해 주십시오", "X")           '⊙: ""
			   Verification = False
			   Exit Function
			End if
		
			sTitle = ""
		
	
			.Col = C_W2	: .Row = Row	: sTitle  	= Trim(.Value )
			if   sTitle= "" then
			   Call DisplayMsgBox("X", "X" , "두번째항목의 구분과 비목을 입력해 주십시오", "X")           '⊙: ""
			   Verification = False
			   Exit Function
			End if
		
		
			sTitle = ""
		
			.Col = C_W3	: .Row = Row	: sTitle  	= Trim(.Value )
		
			if  sTitle= "" then
			   Call DisplayMsgBox("X", "X" , "세번째항목의 구분과 비목을 입력해 주십시오", "X")           '⊙: ""
			   Verification = False
			   Exit Function
			End if

			sTitle = ""
		

			.Col = C_W4	: .Row = Row	: sTitle  	= Trim(.Value )
			if sTitle= "" then
			   Call DisplayMsgBox("X", "X" , "네번째항목의 구분과 비목을 입력해 주십시오", "X")           '⊙: ""
			   Verification = False
			   Exit Function
			End if
        End if				
	    
	
		
	End With
	
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
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData
    Call SetHeader()
    Call SetToolbar("1101110100100111")
	call Fn_CompanyYYMMDD
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
    lgCurrGrid = TYPE_1
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
    
    if lgvspdData(lgCurrGrid).ActiveRow = lgvspdData(lgCurrGrid).maxrows  then exit function
    
    
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    
    Call CheckReCalc()				' 한라인이 취소되면 재계산 

End Function

' 재계산 
Function CheckReCalc()
	Dim dblSum
	
	With lgvspdData(lgCurrGrid)
		ggoSpread.Source = lgvspdData(lgCurrGrid)	
	
        if  lgvspdData(lgCurrGrid).maxrows =< 1 then exit function
        
         
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W1, 2, .MaxRows - 1, true, .MaxRows, C_W1, "V")	' 합계 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W2, 2, .MaxRows - 1, true, .MaxRows, C_W2, "V")	' 합계 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W3, 2, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' 합계 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W4, 2, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 합계 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W5, 2, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W6, 2, .MaxRows - 1, true, .MaxRows, C_W6, "V")	' 합계 
		 
        Call Fn_SumCal
        ggoSpread.UpdateRow .Maxrows
	End With
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
    lgCurrGrid = TYPE_1
    lgvspdData(lgCurrGrid).focus
	With lgvspdData(lgCurrGrid)	' 포커스된 그리드 
			
		ggoSpread.Source = lgvspdData(lgCurrGrid)
			
		iRow = .ActiveRow
		lgvspdData(lgCurrGrid).ReDraw = False
		
		If .MaxRows = 1 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow , 2
			Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
			.Row = iRow	+ 1	
			
			.Col = C_SEQ_NO : .Text = iRow	
			
			iRow = 3		: .Row = iRow
			.Col = C_SEQ_NO : .Text = SUM_SEQ_NO	
			.Col = C_ACCT	: .CellType = 1	: .Text = "(7)계"	: .TypeHAlign = 2
			ggoSpread.SpreadLock C_W9, iRow, C_W6, iRow
		
		Else
				
			If iRow = .MaxRows Then	' -- 마지막 합계줄에서 InsertRow를 하면 상위에 추가한다.
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor lgCurrGrid, iRow, iRow + imRow - 1

				Call SetDefaultVal(lgCurrGrid, iRow, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor lgCurrGrid, iRow+1, iRow + imRow

				Call SetDefaultVal(lgCurrGrid, iRow+1, imRow)
			End If   
		End If
	End With
	
    lgvspdData(lgCurrGrid).ReDraw = True
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


' GetREF 에서 적수 가져온뒤 호출됨 
Function InsertTotalLine(Index)
	With lgvspdData(Index)
	
	ggoSpread.Source = lgvspdData(Index)
	
	If .MaxRows = 0 Then	' 한줄 추가 
		ggoSpread.InsertRow ,1
		
		.Row = 1
		.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
		.Col = C_W9		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
		
		ggoSpread.SpreadLock C_W1, 1, C_W6, 1
	End If
	End With
End Function

' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal(Index, iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(lgCurrGrid)	' 포커스된 그리드 

	ggoSpread.Source = lgvspdData(lgCurrGrid)
	
	If iAddRows = 1 Then ' 1줄만 넣는경우 
		.Row = iRow
		MaxSpreadVal lgvspdData(lgCurrGrid), C_SEQ_NO, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(lgCurrGrid), C_SEQ_NO, iRow)	' 현재의 최대SeqNo를 구한다 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

	With lgvspdData(lgCurrGrid)
		.focus
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		if .Activerow <> 1 or  .Activerow <> .maxrows then
		   lDelRows = ggoSpread.DeleteRow
		end if   
	End With
	
	Call CheckReCalc()				' 한라인이 취소되면 재계산 
	

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
    If ggoSpread.SSCheckChange = True  OR lgBlnFlgChgValue =TRUE Then
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
        'Call DisplayMsgBox("900002", "X", "X", "X")
        'Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF
    Call MakeKeyStream("X")
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
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
		strVal = strVal     & "&txtMaxRows="         & .vspdData0.MaxRows 
        
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function


Function DBQueryFalse()
Call FncNew()
'Call Fn_CompanyYYMMDD

End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>

    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	
	If lgvspdData(TYPE_1).MaxRows > 0 Or _
		lgvspdData(TYPE_2).MaxRows > 0 Or _
		lgvspdData(TYPE_2).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg <> "Y" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
			Call SetSpreadLock(TYPE_1)
			Call SetSpreadLock(TYPE_2)
			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101111100000111")										<%'버튼 툴바 제어 %>
			
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
		End If
	Else
		Call SetToolbar("1100111100000111")										<%'버튼 툴바 제어 %>
	End If

	Call SetSpreadTotalLine ' - 합계라인 재구성 
		Call SetHeader()
	lgBlnFlgChgValue = False
	'lgvspdData(lgCurrGrid).focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    ,strHerder
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    strHerder = ""
    
		With lgvspdData(TYPE_1)
	
			ggoSpread.Source = lgvspdData(i)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
		
		
				 .Col = C_W1 :.Row = 1 : strHerder = strHerder & mid(Trim(.Text),4,len(Trim(.Text)))   &  Parent.gColSep
				 .Col = C_W2 :.Row = 1 : strHerder = strHerder & mid(Trim(.Text),4,len(Trim(.Text)))   &  Parent.gColSep
				 .Col = C_W3 :.Row = 1 : strHerder = strHerder & mid(Trim(.Text),4,len(Trim(.Text))) &  Parent.gColSep
				 .Col = C_W4 :.Row = 1 : strHerder = strHerder & mid(Trim(.Text),4,len(Trim(.Text)))  &  Parent.gColSep
				 .Col = C_W5 :.Row = 1 : strHerder = strHerder & mid(Trim(.Text),4,len(Trim(.Text)))  &  Parent.gColSep

			
			' ----- 1번째 그리드 
			For lRow = 2 To .MaxRows
    
		       .Row = lRow
		       .Col = 0
		       
		       IF lRow = .MaxRows  Then  '계일경우 헤더와 계정이 입력되지 않게 
		           strHerder = ""
		           strHerder = strHerder & ""  &  Parent.gColSep
		           strHerder = strHerder & ""  &  Parent.gColSep
		           strHerder = strHerder & ""  &  Parent.gColSep
		           strHerder = strHerder & ""  &  Parent.gColSep
		           strHerder = strHerder & ""  &  Parent.gColSep
		       End If
		    
		       Select Case .Text
		           Case  ggoSpread.InsertFlag                                      '☜: Insert
		                                              strVal = strVal & "C"  &  Parent.gColSep
		           Case  ggoSpread.UpdateFlag                                      '☜: Update
		                                              strVal = strVal & "U"  &  Parent.gColSep
		           Case  ggoSpread.DeleteFlag                                      '☜: Delete
		                                              strDel = strDel & "D"  &  Parent.gColSep
		       End Select
		       
			  ' 모든 그리드 데이타 보냄     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = C_SEQ_NO To C_W6
					    if lRow = .MaxRows And lCol = C_ACCT Then
					       .Col = lCol : strVal = strVal & ""  &  Parent.gColSep
					    Else
						   .Col = lCol : strVal = strVal & Trim(.Text)  &  Parent.gColSep
						   
						End if   
						
					Next
					
					strVal = strVal &  strHerder  &  Parent.gRowSep
			
					lGrpCnt = lGrpCnt + 1
			  End If  
			Next
		
		End With


	Call MakeKeyStream("S")
	
	Frm1.txtSpread.value      = strDel & strVal
    Frm1.txtMaxRows.value  =     lGrpCnt - 1

	Frm1.txtMode.value        =  Parent.UID_M0002
	frm1.txtFlgMode.value	  =  lgIntFlgMode
	frm1.txtKeyStream.value      =  lgKeyStream
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
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function



'========================================================================================
Function FncBtnPreview() 
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
	Dim ObjName
	
	StrEbrFile = "W6111OA1"
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
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
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">금액 불러오기</A>  
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
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=* VALIGN=TOP>
									<table <%=LR_SPACE_TYPE_20%> border="0" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD">&nbsp;당해연도의 연구 및 인력개발비발생명세</TD>
									   </TR>
									   <TR>
										   <TD width="100%" HEIGHT=100><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										  </TD>
									  </TR>
									   <TR>
										   <TD width="100%" HEIGHT=100%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										  </TD>
									  </TR>
									  <TR>
										  <TD height=*>&nbsp;</TD>
									  </TR>
									</TABLE>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			    
		
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
					<BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
<INPUT TYPE=HIDDEN NAME="txt4yearMth" alt = "4년간 개월수" tag="24">
<INPUT TYPE=HIDDEN NAME="txtyearMth" alt = "당월 개월수"  tag="24" >
<INPUT TYPE=HIDDEN NAME="txt15_12Value" tag="24">
<INPUT TYPE=HIDDEN NAME="txt16_12Value" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCompType" tag="24">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

