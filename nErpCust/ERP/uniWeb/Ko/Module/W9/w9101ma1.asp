<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기타서식 
'*  3. Program ID           : w9101mA1
'*  4. Program Name         : w9101mA1.asp
'*  5. Program Desc         : 제47호 주요계정명세서(갑)
'*  6. Modified date(First) : 2005/02/23
'*  7. Modified date(Last)  : 2005/02/23
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
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

Const BIZ_MNU_ID		= "w9101mA1"
Const BIZ_PGM_ID		= "w9101mB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "w9101mB2.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID		= "w9101OA1"


' -- 그리드 컬럼 정의 
Dim C_W1
Dim C_W1_NM1
Dim C_W1_NM2
Dim C_W2
Dim C_W2_CD
Dim C_W3
Dim C_W4
Dim C_W5

Dim C_01
Dim C_02
Dim C_03
Dim C_04
Dim C_05
Dim C_06
Dim C_07
Dim C_08
Dim C_09
Dim C_10
Dim C_11
Dim C_12
Dim C_13
Dim C_14
Dim C_15
Dim C_16
Dim C_17
Dim C_18
Dim C_19
Dim C_20
Dim C_21
'Dim C_22
'Dim C_23

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgFISC_START_DT, lgFISC_END_DT 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_W1		= 1
	C_W1_NM1	= 2
	C_W1_NM2	= 3
	C_W2		= 4
	C_W2_CD		= 5
	C_W3		= 6
	C_W4		= 7
	C_W5		= 8
	
	C_01 = 1
	C_02 = 2
	C_03 = 3
	C_04 = 4
	C_05 = 5
	C_06 = 6
	C_07 = 7
	C_08 = 8
	C_09 = 9
	C_10 = 10
	C_11 = 11
	C_12 = 12
	C_13 = 13
	C_14 = 14
	C_15 = 15
	C_16 = 16
	C_17 = 17
	C_18 = 18
	C_19 = 19
	C_20 = 20
	C_21 = 21
	'C_22 = 22
	'C_23 = 23
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



'============================================  콤보 박스 채우기  ====================================

Sub InitComboBox()
	' 조회조건(구분)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1

End Sub


Sub InitSpreadSheet()
	Dim ret, iRow
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
 	
	' 1번 그리드 

	With Frm1.vspdData
				
		ggoSpread.Source = Frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20041222_0" ,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W5 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W1,		"(1)구분",			5,,,10,1
		ggoSpread.SSSetEdit		C_W1_NM1,	"(1)구분",			11,2,,50,1
		ggoSpread.SSSetEdit		C_W1_NM2,	"(1)구분",			22,2,,50,1
		ggoSpread.SSSetEdit  	C_W2,		"(2)근거법 조항"		, 20,,,100,1	' 
		ggoSpread.SSSetEdit		C_W2_CD,	"코드",			5,2,,10,1
	    ggoSpread.SSSetFloat	C_W3,		"(3)회사계상금액",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	    ggoSpread.SSSetFloat	C_W4,		"(4)세무조정금액",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
		ggoSpread.SSSetFloat	C_W5,		"(5)차가감금액" & vbCrLf & "((3)-(4))",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 

	    ret = .AddCellSpan(C_W1_NM1, 0, 2, 1)
		.rowheight(0) = 20	' 높이 재지정 

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W1, C_W1, True)

		Call FncNew()
		Call SetSpreadLock()

		.ReDraw = true	
				
	End With 

 
	Call InitSpreadComboBox
	
					
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call GetFISC_DATE
	
	'Exit Sub
		
End Sub


Sub SetSpreadLock()

	With Frm1.vspdData
	
		ggoSpread.Source = Frm1.vspdData

		ggoSpread.SpreadLock C_W1,   -1, C_W2_CD
		ggoSpread.SpreadUnLock C_W3, -1, C_W5	' 전체 적용 

		ggoSpread.SpreadLock C_W3,   C_13, C_W4, C_13
		ggoSpread.SpreadLock C_W3,   C_16, C_W4, C_16
		ggoSpread.SpreadLock C_W3,   C_20, C_W4, C_21
		'ggoSpread.SpreadLock C_W3,   C_23, C_W4, C_23

		ggoSpread.SpreadLock C_W5,   C_01, C_W5, C_12
		ggoSpread.SpreadLock C_W5,   C_14, C_W5, C_15
		ggoSpread.SpreadLock C_W5,   C_17, C_W5, C_19
		'ggoSpread.SpreadLock C_W5,   C_22, C_W5, C_22

	End With	
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow

	With Frm1.vspdData
		ggoSpread.Source = Frm1.vspdData
		For iRow = pvStartRow To pvEndRow
			.Col = C_SEQ_NO
			.Row = iRow
			If .Text = 999999 Then
				ggoSpread.SpreadLock C_W1,   iRow, C_W5, iRow
			Else
				ggoSpread.SpreadUnLock C_W1, iRow, C_W_DESC, iRow	' 전체 적용 
				ggoSpread.SSSetRequired C_W1, iRow, iRow
				ggoSpread.SSSetRequired C_W1, iRow, iRow
				ggoSpread.SSSetRequired C_W1_NM, iRow, iRow
				ggoSpread.SpreadLock C_W5,   iRow, C_W5
			End If
		Next
			
	End With	
End Sub

Sub SetSpreadTotalLine()
	Dim ret
		
	ggoSpread.Source = Frm1.vspdData
	With Frm1.vspdData
		If .MaxRows > 0 Then
			.Row = .MaxRows
			ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' 순번 2행 합침 
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "계"	:	.TypeHAlign = 2
			SetSpreadColor 1, .MaxRows

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
            C_W1		= iCurColumnPos(2)
            C_W1_BT		= iCurColumnPos(3)
            C_W1_NM		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W3		= iCurColumnPos(6)
            C_W4		= iCurColumnPos(7)
            C_W5		= iCurColumnPos(8)
            C_W_DESC	= iCurColumnPos(9)
    End Select    
End Sub

Sub InitSpreadRow()
	Dim ret, iRow

	With Frm1.vspdData
		
		ggoSpread.Source = Frm1.vspdData
		
		'patch version
		If .MaxRows = 0 Then	.MaxRows = C_21

	    ret = .AddCellSpan(C_W1_NM1, C_01, 1, 5)
	    ret = .AddCellSpan(C_W1_NM1, C_06, 1, 4)
	    ret = .AddCellSpan(C_W1_NM1, C_11, 1, 5)
	    ret = .AddCellSpan(C_W1_NM1, C_16, 1, 3)
	    ret = .AddCellSpan(C_W1_NM1, C_19, 2, 1)
	    ret = .AddCellSpan(C_W1_NM1, C_21, 1, 2)
	   ' ret = .AddCellSpan(C_W1_NM1, C_22, 1, 2)
       
	    ' 첫번째 헤더 출력 글자 
		.Col = C_W1_NM1
		.Row = C_01	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "준" & VbCrlf & "비" & vbCrLf & "금" & vbCrLf & "충" & vbCrLf & "당" & vbCrLf & "금" & vbCrLf & "등"
		.Row = C_06	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "손" & VbCrlf & "금" & vbCrLf & "산" & vbCrLf & "입"
		.Row = C_10	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "익금불" & vbCrLf & "산입비"
		.Row = C_11	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "기" & VbCrlf & VbCrlf & "부" & vbCrLf & VbCrlf & "금"
		.Row = C_16	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "접" & VbCrlf & "대" & vbCrLf & "비"
		.Row = C_19	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(119) 외화자산ㆍ부채평가손익"
		.Row = C_21	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = "업무무관부동산" & VbCrlf & "등에 관련한" & VbCrlf & "차입금이자"
	    
		' 두번째 헤더 출력 글자 
		.Col = C_W1_NM2
		.Row = C_01	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(101) 고유목적사업준비금"
		.Row = C_02	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(102) 퇴직급여충당금"
		.Row = C_03	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(103) 퇴직보험료"
		.Row = C_04	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(104) 대손충당금"
		.Row = C_05	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(105) 대손금"
		.Row = C_06	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(106) 합병평가차익"
		.Row = C_07	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(107) 분할평가차익"
		.Row = C_08	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(108) 물적분할자산양도차익"
		.Row = C_09	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(109) 교환자산양도차익"
		.Row = C_10	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(110) 채무면제익등 이월결손금 보전액"
		.Row = C_11	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(111) 당연손금기부금"
		.Row = C_12	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(112) 50% 손금기부금"
		.Row = C_13	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(113) 지정기부금한도액"
		.Row = C_14	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(114) 지정기부금"
		.Row = C_15	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(115) 기타기부금"
		.Row = C_16	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(116) 접대비한도액"
		.Row = C_17	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(117) 접대비" & vbCrlf & "(118 포함)"
		.Row = C_18	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(118) 5만원 (경조사비는 "& vbCrLf &"10만원) 초과 접대비"
		.Row = C_21	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "(120) 업무무관부동산 등"
		
		' 근거법 조항 넣기 
		.Col = C_W2
		.Row = C_01	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.Text = "법인세법 제29조           조세특례제한법 제74조"
		.Row = C_02	:	.TypeHAlign = 0	:			.Text = "법인세법 제33조"
		.Row = C_03	:	.TypeHAlign = 0	:			.Text = "법인세법시행령 제44조의2"
		.Row = C_04	:	.TypeHAlign = 0	:			.Text = "법인세법 제34조"
		.Row = C_05	:	.TypeHAlign = 0	:			.Text = "법인세법 제34조"
		.Row = C_06	:	.TypeHAlign = 0	:			.Text = "법인세법 제44조"
		.Row = C_07	:	.TypeHAlign = 0	:			.Text = "법인세법 제46조"
		.Row = C_08	:	.TypeHAlign = 0	:			.Text = "법인세법 제47조"
		.Row = C_09	:	.TypeHAlign = 0	:			.Text = "법인세법 제50조"
		.Row = C_10	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "법인세법 제18조제8호"
		.Row = C_11	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.Text = "법인세법 제24조제2항   조세특례제한법 제73조 제1항 제1호"
		.Row = C_12	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:			.Text = "조세특례제한법  제73조 제1항 제2호 내지 제14호"
		.Row = C_13	:	.TypeHAlign = 0	:			.Text = "법인세법 제24조 제1항"
		.Row = C_14	:	.TypeHAlign = 0	:			.Text = "법인세법 제24조 제1항"
		.Row = C_15	:	.TypeHAlign = 0	:			.Text = "법인세법 제24조 제1항"
		.Row = C_16	:	.TypeHAlign = 0	:			.Text = "법인세법 제24조 제1항"
		.Row = C_17	:	.TypeHAlign = 0	:			.TypeVAlign = 2	:			.Text = "법인세법 제24조 제1항"
		.Row = C_18	:	.TypeHAlign = 0	:			.Text = "법인세법 제24조 제2항"
		.Row = C_19	:	.TypeHAlign = 0	:			.Text = "법인세법 제42조"
		.Row = C_21	:	.TypeVAlign = 2	:				.Text = "법인세법 제28조 제1항"

		
		' 기본코드값입력하기 
		.Col = C_W1
		.Row = C_01	:	.TypeHAlign = 0	:			.Text = "101"
		.Row = C_02	:	.TypeHAlign = 0	:			.Text = "102"
		.Row = C_03	:	.TypeHAlign = 0	:			.Text = "103"
		.Row = C_04	:	.TypeHAlign = 0	:			.Text = "104"
		.Row = C_05	:	.TypeHAlign = 0	:			.Text = "105"
		.Row = C_06	:	.TypeHAlign = 0	:			.Text = "106"
		.Row = C_07	:	.TypeHAlign = 0	:			.Text = "107"
		.Row = C_08	:	.TypeHAlign = 0	:			.Text = "108"
		.Row = C_09	:	.TypeHAlign = 0	:			.Text = "109"
		.Row = C_10	:	.TypeHAlign = 0	:			.Text = "110"
		.Row = C_11	:	.TypeHAlign = 0	:			.Text = "111"
		.Row = C_12	:	.TypeHAlign = 0	:			.Text = "112"
		.Row = C_13	:	.TypeHAlign = 0	:			.Text = "113"
		.Row = C_14	:	.TypeHAlign = 0	:			.Text = "114"
		.Row = C_15	:	.TypeHAlign = 0	:			.Text = "115"
		.Row = C_16	:	.TypeHAlign = 0	:			.Text = "116"
		.Row = C_17	:	.TypeHAlign = 0	:			.Text = "117"
		.Row = C_18	:	.TypeHAlign = 0	:			.Text = "118"
		.Row = C_19	:	.TypeHAlign = 0	:			.Text = "119"
		.Row = C_20	:	.TypeHAlign = 0	:			.Text = "1191"
		.Row = C_21	:	.TypeHAlign = 0	:			.Text = "120"




		' 홈텍스코드값입력하기 
		.Col = C_W2_CD
		.Row = C_01	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "53"
		.Row = C_02	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "12"
		.Row = C_03	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "71"
		.Row = C_04	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "13"
		.Row = C_05	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "72"
		.Row = C_06	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "55"
		.Row = C_07	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "56"
		.Row = C_08	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "57"
		.Row = C_09	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "58"
		.Row = C_10	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "59"
		.Row = C_11	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "41"
		.Row = C_12	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "64"
		.Row = C_13	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "66"
		.Row = C_14	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "42"
		.Row = C_15	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "73"
		.Row = C_16	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "49"
		.Row = C_17	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "65"
		.Row = C_18	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "61"
		.Row = C_19	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "74"
		.Row = C_20	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "75"
		.Row = C_21	:	.TypeHAlign = 2	:			.TypeVAlign = 2	:			.Text = "76"


		.rowheight(C_01) = 20	' 높이 재지정 
		.rowheight(C_10) = 20	' 높이 재지정 
		.rowheight(C_11) = 30	' 높이 재지정 
		.rowheight(C_12) = 20	' 높이 재지정 
		.rowheight(C_17) = 20	' 높이 재지정 
		.rowheight(C_18) = 20	' 높이 재지정 
		.rowheight(C_20) = 20	' 높이 재지정 
		.rowheight(C_21) = 30	' 높이 재지정 


		.Col = C_W3	:	.Row = C_01	:	.TypeVAlign = 2	:	.Row = C_10	:	.TypeVAlign = 2	:	.Row = C_11	:	.TypeVAlign = 2	:	.Row = C_17	:	.TypeVAlign = 2	:	.Row = C_20	:	.TypeVAlign = 2 : .Row = C_21	:	.TypeVAlign = 2: .Row = C_18	:	.TypeVAlign = 2
		.Col = C_W4	:	.Row = C_01	:	.TypeVAlign = 2	:	.Row = C_10	:	.TypeVAlign = 2	:	.Row = C_11	:	.TypeVAlign = 2	:	.Row = C_17	:	.TypeVAlign = 2	:	.Row = C_20	:	.TypeVAlign = 2 : .Row = C_21	:	.TypeVAlign = 2: .Row = C_18	:	.TypeVAlign = 2
		.Col = C_W5	:	.Row = C_01	:	.TypeVAlign = 2	:	.Row = C_10	:	.TypeVAlign = 2	:	.Row = C_11	:	.TypeVAlign = 2	:	.Row = C_17	:	.TypeVAlign = 2	:	.Row = C_20	:	.TypeVAlign = 2: .Row = C_21	:	.TypeVAlign = 2: .Row = C_18	:	.TypeVAlign = 2
		
		

		.Row = C_20	: .RowHidden = True
'		.Row = C_22	: .RowHidden = True
'		.Row = C_23	: .RowHidden = True

    	
	End With 

End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
     wgRefDoc = GetDocRef(sCoCd,sFiscYear, sRepType, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData

	Call InitSpreadRow()
	Call SetSpreadLock
			
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
	Call Fn_GridCalc()
	lgBlnFlgChgValue = True
	Frm1.vspdData.focus			
End Function

Function ChangeRowFlg(iObj)
	Dim iRow
	
	With iObj
		ggoSpread.Source = iObj
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
	End With
End Function



Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, iCnt
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
		
End Sub

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
  
	' 변경한곳 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	

	Call InitComboBox	' 먼저해야 한다. 기업의 회계기준일을 읽어오기 위해 

	Call InitData

	Call FncQuery()
	
     

End Sub

'============================================  사용자 함수  ====================================
Function Fn_GridCalc()
	Dim iRow, dblSum
	Dim dblW3, dblW4, dblW5
	
    ggoSpread.Source = Frm1.vspdData

	With Frm1.vspdData
		For iRow = C_01 To C_21
		' (5) = (3) - (4) : 113, 116, 120, 121, 123제외 
			If iRow <> C_13 And iRow <> C_16 And iRow <> C_20 And iRow <> C_21  Then ' 서색개정으로 제거됨 
				.Row = iRow	:	.Col = C_W3	:	dblW3 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W4	:	dblW4 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W5	:	dblW5 = dblW3 - dblW4
				.Text = dblW5
			End If
		Next
			
	End With

End Function


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

Sub txtw124_Change( )
    lgBlnFlgChgValue = True
End Sub

Sub txtw125_Change( )
    lgBlnFlgChgValue = True
End Sub

'============================================  그리드 이벤트   ====================================


'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIdx, iRow, sW3, sW4, dblW2

	With Frm1.vspdData
		Select Case Col
			Case C_W3_NM
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
			Case C_W3
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
		End Select
		
		

	End With
End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	lgBlnFlgChgValue= True ' 변경여부 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.UpdateRow Row

	' --- 추가된 부분 
	Call Fn_GridCalc()


	
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = Frm1.vspdData
   
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = Frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	Frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = Frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If Frm1.vspdData.MaxRows = 0 Then
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

	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = Frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
'    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With Frm1.vspdData
		If Row > 0 And Col = C_W1_BT Then
		    .Row = Row
		    .Col = C_W1_BT

		    Call OpenAdItem()
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

	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>
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
'    Call InitVariables													<%'Initializes local global variables%>
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
    Dim IntRetCD  , iRow

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
    Call InitSpreadRow
    Call SetSpreadLock
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
    lgIntFlgMode = parent.OPMD_CMODE
    With Frm1.vspdData
		For iRow = 1 To .MaxRows
				.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
    End With
	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
	Dim iActiveRow
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With Frm1.vspdData
	    ggoSpread.Source = Frm1.vspdData
	    iActiveRow = .ActiveRow

		If .ActiveRow > 0 and .ActiveRow <> .MaxRows Then
			.focus
			.ReDraw = False
		
			ggoSpread.CopyRow
			.Col = C_W1
			.Text = ""

			Call SetDefaultVal(iActiveRow + 1, 1)
			SetSpreadColor iActiveRow, iActiveRow + 1
			.ReDraw = True
			
			Call Fn_GridCalc()
    
		End If
	End With


    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows, iActiveRow, dblSum

	With Frm1.vspdData
		.focus
		iActiveRow = .ActiveRow
		ggoSpread.Source = Frm1.vspdData
		If CheckTotalRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
			Exit Function
		Else
			lDelRows = ggoSpread.EditUndo
		End If
		
	End With

	Call Fn_GridCalc()

End Function

' -- 합계 행인지 체크(Header Grid)
Function CheckTotalRow(Byref pObj, Byval pRow) 
	CheckTotalRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = "999999" And pObj.MaxRows > 1 Then	 ' 합계 행 
		CheckTotalRow = True
	End If
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo
    Dim ret

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

	With Frm1.vspdData
	
		.focus
		ggoSpread.Source = Frm1.vspdData
	
		iSeqNo = .MaxRows+1
	
		if .MaxRows = 0 then
		
			ggoSpread.InsertRow  imRow 
			.Col	= C_SEQ_NO	:	.Text	= 1
			SetSpreadColor 1, 1
			
			ggoSpread.InsertRow  imRow 
			.Row = .MaxRows
			.Col	= C_SEQ_NO	:	.Text	= 999999
			ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' 순번 2행 합침 
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "계"	:	.TypeHAlign = 2
			SetSpreadColor .MaxRows, .MaxRows
			.Row  = 1
			.ActiveRow = 1

		else
			iRow = .ActiveRow

			If iRow = .MaxRows Then	' -- 마지막 합계줄에서 InsertRow를 하면 상위에 추가한다.
				iRow = iRow - 1
				ggoSpread.InsertRow iRow, imRow 
				SetSpreadColor iRow, iRow + imRow + 1

				Call SetDefaultVal(iRow + 1, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor iRow, iRow + imRow + 1

				Call SetDefaultVal(iRow + 1, imRow)
			End If   
			.vspdData.Row  = iRow + 1
			.vspdData.ActiveRow = iRow +1
			
        End if 	
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With Frm1.vspdData
	
		If iAddRows = 1 Then ' 1줄만 넣는경우 
			.Row = iRow
			.Value = MaxSpreadVal(Frm1.vspdData, C_SEQ_NO, iRow)
		Else
			iSeqNo = MaxSpreadVal(Frm1.vspdData, C_SEQ_NO, iRow)	' 현재의 최대SeqNo를 구한다 
			
			For i = iRow to iRow + iAddRows -1
				.Row = i	:	.Col = C_SEQ_NO
				If .Text <> 999999 Then
					: .Value = iSeqNo : iSeqNo = iSeqNo + 1
				End If
			Next
		End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows, iActiveRow, dblSum

	With Frm1.vspdData
		.focus
		iActiveRow = .ActiveRow
		ggoSpread.Source = Frm1.vspdData
		If CheckTotalRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
			Exit Function
		Else
			lDelRows = ggoSpread.DeleteRow
		End If
		
	End With

	Call Fn_GridCalc()

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
	
    ggoSpread.Source = Frm1.vspdData
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
        'strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryFalse()
	Call FncNew()
End Function
		
Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		

	    Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
	End If
	
	Call InitSpreadRow()
	Call SetSpreadLock
'	Call SetSpreadTotalLine ' - 합계라인 재구성 
	
	Frm1.vspdData.focus			
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
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
		
		' ----- 1번째 그리드 
		For lRow = 1 To .MaxRows

	       .Row = lRow
	       .Col = 0
	    
	       Select Case .Text
	           Case  ggoSpread.InsertFlag                                      '☜: Insert
	                                              strVal = strVal & "C"  &  Parent.gColSep
	           Case  ggoSpread.UpdateFlag                                      '☜: Update
	                                              strVal = strVal & "U"  &  Parent.gColSep
		       Case  ggoSpread.DeleteFlag                                      '☜: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
		       Case Else
		                                          strVal = strVal & ""  &  Parent.gColSep
	       End Select
	       
		  ' 모든 그리드 데이타 보냄     
'		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = 1 To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
'		  End If  
		Next
	
	End With

	Frm1.txtSpread.value      = strVal
	strVal = ""

	Frm1.txtMode.value		=	Parent.UID_M0002
	Frm1.txtFlgMode.Value	=	lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
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
					<TD WIDTH=* align=right><A href="vbscript:GetRef">금액불러오기</A></TD>
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD >
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">
											<TABLE <%=LR_SPACE_TYPE_60%> border="1" height=100% width="100%">
												<TR>
													<TD CLASS="TD51" width="10%" ALIGN=CENTER>
														상여배당등 
													</TD>
													<TD CLASS="TD51" width="26%" ALIGN=CENTER>
														(121)소득처분 금액   (법인세법시행령제106조)
													</TD>
													<TD CLASS="TD51" width="4%" ALIGN=CENTER>
														97
													</TD>
													<TD CLASS="TD51" width="17%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW124" name=txtW124 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
													<TD CLASS="TD51" width="22%">
														(122)이익처분 금액      (상법제462조등)
													</TD>
													<TD CLASS="TD51" width="4%" ALIGN=CENTER>
														98
													</TD>
													<TD CLASS="TD51" width="17%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW125" name=txtW125 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
												</TR>
											</TABLE>

										</TD>
									</TR>
								</TABLE>
							    </TD>
							</TR>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
