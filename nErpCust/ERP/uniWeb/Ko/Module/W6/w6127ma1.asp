
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기타 서식 
'*  3. Program ID           : w6127ma1
'*  4. Program Name         : w6127ma1.asp
'*  5. Program Desc         : 제 55호 소득자료명세서 
'*  6. Modified date(First) : 2005/01/27
'*  7. Modified date(Last)  : 2006/02/08
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

Const BIZ_MNU_ID = "w6127ma1"
Const BIZ_PGM_ID = "w6127mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "w6127Oa1"
Dim C_W1_1
Dim C_W1_2
Dim C_W1_CD
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5

' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
Dim C_ROW_01
Dim C_ROW_02
Dim C_ROW_03
Dim C_ROW_04
Dim C_ROW_05
Dim C_ROW_06
Dim C_ROW_07
Dim C_ROW_08
Dim C_ROW_09
Dim C_ROW_10
Dim C_ROW_11
Dim C_ROW_12
Dim C_ROW_13
Dim C_ROW_14
Dim C_ROW_15
Dim C_ROW_16
Dim C_ROW_17
Dim C_ROW_18
Dim C_ROW_19
Dim C_ROW_20
Dim C_ROW_21
Dim C_ROW_22
Dim C_ROW_23
Dim C_ROW_24
Dim C_ROW_25

DIM lgCOMP_TYPE1
Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim	lgFISC_START_DT, lgFISC_END_DT, lgW2001, lgMonGap, lgW2019

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_W1_1				= 1
    C_W1_2				= 2
    C_W1_CD				= 3
    C_W2				= 4
    C_W3				= 5
    C_W4				= 6
    C_W5				= 7

	' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
	C_ROW_01			= 1
	C_ROW_02			= 2
	C_ROW_03			= 3
	C_ROW_04			= 4
	C_ROW_05			= 5
	C_ROW_06			= 6
	C_ROW_07			= 7
	C_ROW_08			= 8
	C_ROW_09			= 9
	C_ROW_10			= 10
	C_ROW_11			= 11
	C_ROW_12			= 12
	C_ROW_13			= 13
	C_ROW_14			= 14
	C_ROW_15			= 15
	C_ROW_16			= 16
	C_ROW_17			= 17
	C_ROW_18			= 18
	C_ROW_19			= 21
	C_ROW_20			= 22
	C_ROW_21			= 23
	C_ROW_22			= 24
	C_ROW_23			= 25
	C_ROW_24			= 19	' -- 행이 중간에 끼어듬 
	C_ROW_25			= 20

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

	call CommonQueryRs("REFERENCE_1"," ufn_TB_Configuration('W2001','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           
    lgW2001 = Split(lgF0 , chr(11))
    lgW2001(0) = UNICDbl(lgW2001(0))
    lgW2001(1) = UNICDbl(lgW2001(1))
End Sub

Sub InitSpreadSheet()
	Dim ret
		
	Call initSpreadPosVariables()  

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData	
		'patch version
		ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
			 
		.ReDraw = false
			 
		.MaxCols = C_W5 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    
				       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W1_1,		"(1)구분", 23,,,100,1
		ggoSpread.SSSetEdit		C_W1_CD,	"코드" , 7,2,,10,1
		ggoSpread.SSSetFloat	C_W2,		"(2)감면후세액", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W3,		"(3)최저한세", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W4,		"(4)조정감", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_W5,		"(5)조정후세액" , 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  

		ret = .AddCellSpan(C_W1_1	, 0, 2, 1)	' 수평 2열 합침 
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = true
			
		'Call SetSpreadLock 
	    
	End With

End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock()
    With frm1.vspdData

		.ReDraw = False
		ggoSpread.SpreadLock C_W1_1,  -1, C_W1_CD
		ggoSpread.SpreadLock C_W5,  C_ROW_01, C_W5, C_ROW_22
		
		ggoSpread.SpreadLock C_W3,  C_ROW_01, C_W4, C_ROW_03
		ggoSpread.SpreadLock C_W2,  C_ROW_04, C_W4, C_ROW_04
		ggoSpread.SpreadLock C_W2,  C_ROW_05, C_W2, C_ROW_06
		ggoSpread.SpreadLock C_W2,  C_ROW_07, C_W4, C_ROW_03
		ggoSpread.SpreadLock C_W3,  C_ROW_08, C_W4, C_ROW_12
		ggoSpread.SpreadLock C_W2,	C_ROW_10, C_W4, C_ROW_10
		ggoSpread.SpreadLock C_W2,	C_ROW_13, C_W2, C_ROW_14
		ggoSpread.SpreadLock C_W2,	C_ROW_15, C_W4, C_ROW_15
		ggoSpread.SpreadLock C_W3,	C_ROW_16, C_W4, C_ROW_16
		ggoSpread.SpreadLock C_W2,	C_ROW_17, C_W2, C_ROW_17
		ggoSpread.SpreadLock C_W2,	C_ROW_18, C_W4, C_ROW_18
		ggoSpread.SpreadLock C_W2,	C_ROW_19, C_W4, C_ROW_19
		ggoSpread.SpreadLock C_W2,	C_ROW_20, C_W4, C_ROW_20
		ggoSpread.SpreadLock C_W3,	C_ROW_21, C_W3, C_ROW_22
		ggoSpread.SpreadLock C_W2,	C_ROW_23, C_W4, C_ROW_23
		
		ggoSpread.SpreadLock C_W2,	C_ROW_24, C_W4, C_ROW_24
		ggoSpread.SpreadLock C_W2,	C_ROW_25, C_W4, C_ROW_25

		.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
 
  	'ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
      
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_W1_1				= iCurColumnPos(1)
            C_W1_2				= iCurColumnPos(2)
            C_W1_CD				= iCurColumnPos(3)
            C_W2				= iCurColumnPos(4)
            C_W3				= iCurColumnPos(5)
            C_W4				= iCurColumnPos(6)
            C_W5				= iCurColumnPos(7)
    End Select    
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call GetFISC_DATE
	
	' 그리드 행 추가 
	ggoSpread.Source = frm1.vspdData
	
	' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
	ggoSpread.InsertRow , C_ROW_23
	
	
	call CommonQueryRs("COMP_TYPE1 ,fisc_end_dt "," TB_COMPANY_HISTORY "," CO_CD = "&filterVar(frm1.txtCO_CD.value,"''","S")&" and FISC_YEAR="&filterVar(frm1.txtFISC_YEAR.text,"''","S")&" and REP_TYPE="&filterVar(frm1.cboREP_TYPE.value,"''","S")&" ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	lgCOMP_TYPE1=Replace(lgF0,chr(11),"")
	
		
	Call SpreadInitData()
End Sub

Sub SpreadInitData()
	Dim iRow , ret
	
	iRow = 1
	With frm1.vspdData
	
		.Redraw = False
	
		.RowHeight(-1) = 17
		.ColWidth(C_W1_1) = 7	: .ColWidth(C_W1_2) = 26
		
		.Col  = C_W1_1	: .Row  = -1	: .Col2 = C_W5	: .Row2 = -1
		.BlockMode = True
		.TypeEditMultiLine = True : .TypeVAlign = 2	
		.BlockMode = False
	
		ret = .AddCellSpan(C_W1_1	, C_ROW_01, 2, 1)	' 수평 2열 합침 
		
		ret = .AddCellSpan(C_W1_1	, C_ROW_02, 1, 2)	' 수직 2행 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_05, 1, 2)	' ""
		
		ret = .AddCellSpan(C_W1_1	, C_ROW_04, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_07, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_08, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_09, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_10, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_11, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_12, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_13, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_14, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_15, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_16, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_17, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_18, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_19, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_20, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_21, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_22, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_23, 2, 1)	' 수평 2열 합침 
		
		' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
		ret = .AddCellSpan(C_W1_1	, C_ROW_24, 2, 1)	' 수평 2열 합침 
		ret = .AddCellSpan(C_W1_1	, C_ROW_25, 2, 1)	' 수평 2열 합침 
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(101) 결 산 서 상 당 기 순 이 익"
		.Col = C_W1_CD	: .value = "01"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "소 득" & vbCrLf & "조 정" & vbCrLf & "금 액"	: .TypeHAlign = 2
		.Col = C_W1_2	: .value = "(102) 익 금 산 입"
		.Col = C_W1_CD	: .value = "02"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_2	: .value = "(103) 손 금 산 입"
		.Col = C_W1_CD	: .value = "03"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(104) 조 정 후 소 득 금 액" & vbCrLf & "        [(101)+(102)-(103)]"
		.Col = C_W1_CD	: .value = "04"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "최저한세 적용대상 특별비용"	: .TypeHAlign = 2
		.Col = C_W1_2	: .value = "(105) 준 비 금"
		.Col = C_W1_CD	: .value = "05"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_2	: .value = "(106) 특별상각및특례자산감가상각비"
		.Col = C_W1_CD	: .value = "06"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(107) 특 별 비 용 손 급 산 입 전" & vbCrLf & "        소 득 금 액 [(104)+(105)+(106)]"
		.Col = C_W1_CD	: .value = "07"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(108) 기 부 금 한 도 초 과 액"
		.Col = C_W1_CD	: .value = "08"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(109) 기 부 금 한 도 초 과" & vbCrLf & "        이 월 액 손 금 산 입"
		.Col = C_W1_CD	: .value = "09"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(110) 각 사 업 연 도 소 득 금 액"
		.Col = C_W1_CD	: .value = "10"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(111) 이 월 결 손 금"
		.Col = C_W1_CD	: .value = "11"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(112) 비 과 세 소 득"
		.Col = C_W1_CD	: .value = "12"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(113) 최 저 한 세 적 용 대 상" & vbCrLf & "        비  과  세  소  득"
		.Col = C_W1_CD	: .value = "13"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(114) 최 저 한 세 적 용 대 상" & vbCrLf & "        입  금  불  산  입"
		.Col = C_W1_CD	: .value = "14" 
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(115) 차 가 감 소 득 금 액" & vbCrLf & "        [(110)-(111)-(112)+(113)+(114)]"
		.Col = C_W1_CD	: .value = "15"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(116) 소 득 공 제"
		.Col = C_W1_CD	: .value = "16"
		
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(117) 최 저 한 세 적 용 대 상 소 득 공 제"
		.Col = C_W1_CD	: .value = "17"

		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(118) 과 세 표 준 금 액 [(115)-(116)+(117)]"
		.Col = C_W1_CD	: .value = "18"

		' -- 2006-01-02 : 200603 개정판 수정 (선박표준이익 추가)
		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(124) 선 박 표 준 이 익"
		.Col = C_W1_CD	: .value = "24"

		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(125) 과 세 표 준 금 액 [(118)+(124)]"
		.Col = C_W1_CD	: .value = "25"
		' --------------------------------------------------------

		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(119) 세  율"
		.Col = C_W1_CD	: .value = "19"
		Call MakePercentType( frm1.vspdData, C_W2, C_ROW_19, C_W5, C_ROW_19, 0, "", "")

		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(120) 산 출 세 액"
		.Col = C_W1_CD	: .value = "20"

		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(121) 감 면 세 액"
		.Col = C_W1_CD	: .value = "21"

		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(122) 세 액 공 제"
		.Col = C_W1_CD	: .value = "22"

		.Row = iRow		: iRow = iRow + 1	
		.Col = C_W1_1	: .value = "(123) 차 감 세 액 [(120)-(121)-(122)]"
		.Col = C_W1_CD	: .value = "23"


		.Redraw = True
		
		.SetActiveCell C_W2, 1
		Call SetSpreadLock
	End With
End Sub


'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 그리드1의 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2, iMaxRows, sTmp
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

	' 변경될 위치를 알려줌 
	Dim iCol, iRow
	With frm1.vspdData
		iCol = .ActiveCol	: iRow = .ActiveRow

		.AllowMultiBlocks = True
		.SetSelection C_W2, C_ROW_01, C_W2, C_ROW_03  ' -- 처음 선택할때 
		.AddSelection C_W3, C_ROW_05, C_W3, C_ROW_06	' -- 개별행을 여러개 추가할때 
		.AddSelection C_W2, C_ROW_08, C_W2, C_ROW_09
		.AddSelection C_W2, C_ROW_11, C_W2, C_ROW_11
		
		' -- 2006-01-02 : 200603개정판 적용 
		'.AddSelection C_W2,21, C_W2, 22
		.AddSelection C_W2, C_ROW_21, C_W2, C_ROW_22
		
		IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
		
		.SetSelection iCol, iRow, iCol, iRow
		
		If IntRetCD = vbNo Then
			 Exit Function
		End If
	End With
	
    Dim IntRetCD1

	IntRetCD = CommonQueryRs("W01"," TB_3 " , " CO_CD='" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If IntRetCD = False Then
		Call DisplayMsgBox("W60006", parent.VB_INFORMATION, "", "X") 
	End If
	
	IntRetCD = CommonQueryRs("W1, W2"," dbo.ufn_TB_4_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		iMaxRows	= UBound(arrW1)

		With frm1.vspdData
		
		For iRow = 0 To iMaxRows -1
			.Row = CDbl(arrW1(iRow))
			
			Select Case arrW1(iRow)
				Case "01", "02", "03", "08", "09", "11"	', "21", "22"
					.Col = C_W2 : .value = arrW2(iRow)
					
				' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
				Case "21", "22"
					.Col = C_W2 : .Row = CDbl(arrW1(iRow))+2
					.value = arrW2(iRow)
				Case Else
					.Col = C_W3 : .value = arrW2(iRow)
			End Select
		Next
		
		End With
		
		Call SetReCalc1
	End If
	
	
	frm1.vspdData.focus
End Function

Function GetRef2()
	With frm1.vspdData
		' W4, W5 를 초기화 
		.BlockMode = True
		.Col = C_W4	: .Row = 1
		'.Col2 = C_W5	: .Row2 = 23
		' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
		.Col2 = C_W5	: .Row2 = C_ROW_23
		.Text = ""
		.BlockMode = False
	End With	
	Call SetReCalc2("")
End Function

' -- 금액 불러오기/그리드 (2),(3) 수정시 계산 
Sub SetReCalc1()
	' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
	Dim dblAmt(25, 7)	' 7 = C_W5
	
	With frm1.vspdData
		
		' (2) 감면후 세액 공식 -------------------------------------------------------
		dblAmt(1, C_W2) = GetGrid(C_W2, C_ROW_01)
		dblAmt(2, C_W2) = GetGrid(C_W2, C_ROW_02)
		dblAmt(3, C_W2) = GetGrid(C_W2, C_ROW_03)
		
		dblAmt(4, C_W2) = dblAmt(1, C_W2) + dblAmt(2, C_W2) - dblAmt(3, C_W2)
		Call PutGrid(C_W2, C_ROW_04, dblAmt(4, C_W2))	'(104) 조정후소득금액 
		
		dblAmt(7, C_W2) = dblAmt(4, C_W2)
		Call PutGrid(C_W2, C_ROW_07, dblAmt(7, C_W2))	'(107) 특별비용손금산입전소득금액 
		
		dblAmt(8, C_W2) = GetGrid(C_W2, C_ROW_08)
		dblAmt(9, C_W2) = GetGrid(C_W2, C_ROW_09)
		
		dblAmt(10, C_W2) = dblAmt(7, C_W2) + dblAmt(8, C_W2) - dblAmt(9, C_W2)
		Call PutGrid(C_W2, C_ROW_10, dblAmt(10, C_W2))	'(110) 각사업연도소득금액 
		
		dblAmt(11, C_W2) = GetGrid(C_W2, C_ROW_11)
		dblAmt(12, C_W2) = GetGrid(C_W2, C_ROW_12)
		
		dblAmt(15, C_W2) = dblAmt(10, C_W2) - dblAmt(11, C_W2) - dblAmt(12, C_W2)
		Call PutGrid(C_W2, C_ROW_15, dblAmt(15, C_W2))	'(115) 차가감소득금액 
		
		dblAmt(16, C_W2) = GetGrid(C_W2, C_ROW_16)
		
		dblAmt(18, C_W2) = dblAmt(15, C_W2) - dblAmt(16, C_W2) 
		Call PutGrid(C_W2, C_ROW_18, dblAmt(18, C_W2))	'(118) 과세표준금액 

		' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
		dblAmt(24, C_W2) = GetGrid(C_W2, C_ROW_24)
		dblAmt(25, C_W2) = dblAmt(24, C_W2) + dblAmt(18, C_W2)
		Call PutGrid(C_W2, C_ROW_25, dblAmt(25, C_W2))	'(125) 과세표준금액 
		
		'If dblAmt(18, C_W2) * (12 / lgMonGap) > 100000000 Then
		If dblAmt(25, C_W2) * (12 / lgMonGap) > 100000000 Then
			Call PutGrid(C_W2, C_ROW_19, lgW2001(1))	'(120) 산출세액 
		Else
			Call PutGrid(C_W2, C_ROW_19, lgW2001(0))	'(120) 산출세액 
		End If

		' -- 18 과세표준금액이 25과세표준금액으로 대체됨(선박표준이익 추가되서 그럼)
		'If dblAmt(18, C_W2) < 0 Then
		If dblAmt(25, C_W2) < 0 Then
			dblAmt(20, C_W2) = 0
		'ElseIf dblAmt(18, C_W2) * (12 / lgMonGap) > 100000000 Then
		ElseIf dblAmt(25, C_W2) * (12 / lgMonGap) > 100000000 Then
			'dblAmt(20, C_W2) = (dblAmt(18, C_W2) * (12 / lgMonGap) * lgW2001(1) * lgMonGap/12 ) - (( 100000000 * (lgW2001(1) - lgW2001(0))) * lgMonGap/12)
			dblAmt(20, C_W2) = (dblAmt(25, C_W2) * (12 / lgMonGap) * lgW2001(1) * lgMonGap/12 ) - (( 100000000 * (lgW2001(1) - lgW2001(0))) * lgMonGap/12)
		Else
			'dblAmt(20, C_W2) = (dblAmt(18, C_W2) * (12 / lgMonGap) * lgW2001(0)) * lgMonGap / 12
			dblAmt(20, C_W2) = (dblAmt(25, C_W2) * (12 / lgMonGap) * lgW2001(0)) * lgMonGap / 12
		End If
		Call PutGrid(C_W2, C_ROW_20, dblAmt(20, C_W2))	'(120) 산출세액 
		
		dblAmt(21, C_W2) = GetGrid(C_W2, C_ROW_21)
		dblAmt(22, C_W2) = GetGrid(C_W2, C_ROW_22)
		
		dblAmt(23, C_W2) = dblAmt(20, C_W2) - dblAmt(21, C_W2) - dblAmt(22, C_W2)
		If dblAmt(23, C_W2) < 0 Then dblAmt(23, C_W2) = 0
		Call PutGrid(C_W2, C_ROW_23, dblAmt(23, C_W2))	'(115) 차가감소득금액		
		
		
		' (3) 최저한세 공식 -------------------------------------------------------
		dblAmt(4, C_W3) = GetGrid(C_W2, C_ROW_04)
		Call PutGrid(C_W3, C_ROW_04, dblAmt(4, C_W3))
		dblAmt(5, C_W3) = GetGrid(C_W3, C_ROW_05)
		dblAmt(6, C_W3) = GetGrid(C_W3, C_ROW_06)
		
		dblAmt(7, C_W3) = dblAmt(4, C_W3) + dblAmt(5, C_W3) + dblAmt(6, C_W3)
		Call PutGrid(C_W3, C_ROW_07, dblAmt(7, C_W3))	'(107) 특별비용손금산입전소득금액 
		
		dblAmt(8, C_W3) = GetGrid(C_W2, C_ROW_08)
		Call PutGrid(C_W3, C_ROW_08, dblAmt(8, C_W3))
		dblAmt(9, C_W3) = GetGrid(C_W2, C_ROW_09)
		Call PutGrid(C_W3, C_ROW_09, dblAmt(9, C_W3))
		
		dblAmt(10, C_W3) = dblAmt(7, C_W3) + dblAmt(8, C_W3) - dblAmt(9, C_W3)
		Call PutGrid(C_W3, C_ROW_10, dblAmt(10, C_W3))	'(110) 각사업연도소득금액 
		
		dblAmt(11, C_W3) = dblAmt(11, C_W2)
		Call PutGrid(C_W3, C_ROW_11, dblAmt(11, C_W3))	' (2)감면후 세액을 입력 
		dblAmt(12, C_W3) = dblAmt(12, C_W2)
		Call PutGrid(C_W3, C_ROW_12, dblAmt(12, C_W3))	' (2)감면후 세액을 입력 
			
		dblAmt(13, C_W3) = GetGrid(C_W3, C_ROW_13)
		dblAmt(14, C_W3) = GetGrid(C_W3, C_ROW_14)
		
		dblAmt(15, C_W3) = dblAmt(10, C_W3) - dblAmt(11, C_W3) - dblAmt(12, C_W3) + dblAmt(13, C_W3) + dblAmt(14, C_W3)
		Call PutGrid(C_W3, C_ROW_15, dblAmt(15, C_W3))	'(115) 차가감소득금액 
		
		dblAmt(16, C_W3) = dblAmt(16, C_W2)		
		Call PutGrid(C_W3, C_ROW_16, dblAmt(16, C_W3))	' (2)감면후 세액을 입력 
		
		dblAmt(17, C_W3) = GetGrid(C_W3, C_ROW_17)
		dblAmt(18, C_W3) = dblAmt(15, C_W3) - dblAmt(16, C_W3) + dblAmt(17, C_W3) 
		Call PutGrid(C_W3, C_ROW_18, dblAmt(18, C_W3))	'(118) 과세표준금액 

		' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
		dblAmt(24, C_W3) = GetGrid(C_W3, C_ROW_24)
		dblAmt(25, C_W3) = dblAmt(24, C_W3) + dblAmt(18, C_W3) 
		Call PutGrid(C_W3, C_ROW_25, dblAmt(25, C_W3))	'(125) 과세표준금액 

		' -- 200603 변경 중소기업 10%, 일반 기업(감면전 과세표준이 1000억원까지 13%, 1000억 초과분에 대해 15%)
		'If lgW2019 = 0.15 Then
		'	If dblAmt(25, C_W3) < 100000000000 Then
		'		lgW2019 = 0.13
		'	End If
		'End If

		'=================================================================
		'200703 
		'=================================================================
		dim BAmt 
		BAmt=100000000000

		If lgCOMP_TYPE1 = "1" Then '일반 기업 
			lgW2019 = 0.13
			
			If dblAmt(25, C_W3) < BAmt Then
				dblAmt(20, C_W3) = dblAmt(25, C_W3) * lgW2019
			else
				dblAmt(20, C_W3) = (BAmt *lgW2019) + ((dblAmt(25, C_W3) - Bamt) * 0.15)
			End If

		else
			dblAmt(20, C_W3) = dblAmt(25, C_W3) * lgW2019
		end if

		Call PutGrid(C_W3, C_ROW_19, lgW2019)	'(120) 산출세액 
		
		
		
		
		If dblAmt(20, C_W3) < 0 Then dblAmt(20, C_W3) = 0
		Call PutGrid(C_W3, C_ROW_20, dblAmt(20, C_W3))	'(120) 산출세액 
		
	End With
	lgBlnFlgChgValue = True
		
	Call AllUpdateFlg
End Sub

' 해당 그리드에서 데이타가져오기 
Function GetGrid(Byval pCol, Byval pRow)
	With frm1.vspdData
		.Col = pCol	: .Row = pRow : GetGrid = UNICDbl(.Value)
	End With
End Function

' 해당 그리드에서 데이타가져오기 
Function PutGrid(Byval pCol, Byval pRow, Byval pVal)
	With frm1.vspdData
		.Col = pCol	: .Row = pRow 
		If IsNumeric(pVal) Then
			If pVal <> 0 Then
				.Value = pVal
			Else 
				.Text = ""
			End If
		Else
			.Value = pVal
		End If
	End With
End Function

'  최저한세 링크 클릭시 호출됨 
Sub SetReCalc2(Byval pEvent)
	Dim dblAmt(25, 7)	' 7 = C_W5
	
	With frm1.vspdData
		lgBlnFlgChgValue = true
		If GetGrid(C_W2, C_ROW_23) > GetGrid(C_W3, C_ROW_20) Then
			'MsgBox "(2)감면후세액의 (123)차감금액이 (3)최저한세의 (120)산출세액보다 큽니다"
			Call SetReCalc5()
			Exit Sub
		End If
			
		' (5) 조정후세액 
		dblAmt(4, C_W5) = GetGrid(C_W2, C_ROW_04)
		Call PutGrid(C_W5, C_ROW_04, dblAmt(4, C_W5))
		
		dblAmt(5, C_W5) = GetGrid(C_W4, C_ROW_05)
		Call PutGrid(C_W5, C_ROW_05, dblAmt(5, C_W5))
		
		dblAmt(6, C_W5) = GetGrid(C_W4, C_ROW_06)
		Call PutGrid(C_W5, C_ROW_06, dblAmt(6, C_W5))
		
		dblAmt(7, C_W5) = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5)
		Call PutGrid(C_W5, C_ROW_07, dblAmt(7, C_W5))
		
		dblAmt(8, C_W5) = GetGrid(C_W2, C_ROW_08)
		Call PutGrid(C_W5, C_ROW_08, dblAmt(8, C_W5))
		
		dblAmt(9, C_W5) = GetGrid(C_W2, C_ROW_09)
		Call PutGrid(C_W5, C_ROW_09, dblAmt(9, C_W5))
		
		dblAmt(10, C_W5) = dblAmt(7, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5)
		Call PutGrid(C_W5, C_ROW_10, dblAmt(10, C_W5))
		
		dblAmt(11, C_W5) = GetGrid(C_W2, C_ROW_11)
		Call PutGrid(C_W5, C_ROW_11, dblAmt(11, C_W5))
		dblAmt(12, C_W5) = GetGrid(C_W2, C_ROW_12)
		Call PutGrid(C_W5, C_ROW_12, dblAmt(12, C_W5))
		dblAmt(13, C_W5) = GetGrid(C_W4, C_ROW_13)
		Call PutGrid(C_W5, C_ROW_13, dblAmt(13, C_W5))
		dblAmt(14, C_W5) = GetGrid(C_W4, C_ROW_14)
		Call PutGrid(C_W5, C_ROW_14, dblAmt(14, C_W5))
		
		dblAmt(15, C_W5) = dblAmt(10, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) + dblAmt(13, C_W5) + dblAmt(14, C_W5)
		Call PutGrid(C_W5, C_ROW_15, dblAmt(15, C_W5))
		
		dblAmt(16, C_W5) = GetGrid(C_W2, C_ROW_16)
		Call PutGrid(C_W5, C_ROW_16, dblAmt(16, C_W5))
		
		dblAmt(17, C_W5) = GetGrid(C_W4, C_ROW_17)
		Call PutGrid(C_W5, C_ROW_17, dblAmt(17, C_W5))
		
		dblAmt(18, C_W5) = dblAmt(15, C_W5) - dblAmt(16, C_W5) + dblAmt(17, C_W5) 
		Call PutGrid(C_W5, C_ROW_18, dblAmt(18, C_W5))

		' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
		dblAmt(24, C_W5) = GetGrid(C_W3, C_ROW_24)
		dblAmt(25, C_W5) = dblAmt(24, C_W5) + dblAmt(18, C_W5) 
		Call PutGrid(C_W5, C_ROW_25, dblAmt(25, C_W5))	'(125) 과세표준금액 

		'If dblAmt(18, C_W5) * (12 / lgMonGap) > 100000000 Then
		If dblAmt(25, C_W5) * (12 / lgMonGap) > 100000000 Then
			Call PutGrid(C_W5, C_ROW_19, lgW2001(1))	'(120) 산출세액 
		Else
			Call PutGrid(C_W5, C_ROW_19, lgW2001(0))	'(120) 산출세액 
		End If

		'If dblAmt(18, C_W5) < 0 Then
		If dblAmt(25, C_W5) < 0 Then
			dblAmt(20, C_W5) = 0
		'ElseIf dblAmt(18, C_W5) * (12 / lgMonGap) > 100000000 Then
		ElseIf dblAmt(25, C_W5) * (12 / lgMonGap) > 100000000 Then
			'dblAmt(20, C_W5) = (dblAmt(18, C_W5) * (12 / lgMonGap) * lgW2001(1) * lgMonGap/12 ) - (( 100000000 * (lgW2001(1) - lgW2001(0))) * lgMonGap/12)
			dblAmt(20, C_W5) = (dblAmt(25, C_W5) * (12 / lgMonGap) * lgW2001(1) * lgMonGap/12 ) - (( 100000000 * (lgW2001(1) - lgW2001(0))) * lgMonGap/12)
		Else
			'dblAmt(20, C_W5) = (dblAmt(18, C_W5) * (12 / lgMonGap) * lgW2001(0)) * lgMonGap / 12
			dblAmt(20, C_W5) = (dblAmt(25, C_W5) * (12 / lgMonGap) * lgW2001(0)) * lgMonGap / 12
		End If
		Call PutGrid(C_W5, C_ROW_20, dblAmt(20, C_W5))	'(120) 산출세액 
				
		'If dblAmt(18, C_W5) > 100000000 Then
		'	dblAmt(20, C_W5) = (dblAmt(18, C_W5) * lgW2001(1)) - (100000000 * (lgW2001(1) - lgW2001(0)))
		'Else
		'	dblAmt(20, C_W5) = (dblAmt(18, C_W5) * lgW2001(0)) 
		'End If
		'Call PutGrid(C_W5, 20, dblAmt(20, C_W5))
		
		dblAmt(21, C_W2) = GetGrid(C_W2, C_ROW_21)
		dblAmt(21, C_W4) = GetGrid(C_W4, C_ROW_21)
		dblAmt(21, C_W5) = dblAmt(21, C_W2) - dblAmt(21, C_W4)
		Call PutGrid(C_W5, C_ROW_21, dblAmt(21, C_W5))
		
		dblAmt(22, C_W2) = GetGrid(C_W2, C_ROW_22)
		dblAmt(22, C_W4) = GetGrid(C_W4, C_ROW_22)
		dblAmt(22, C_W5) = dblAmt(22, C_W2) - dblAmt(22, C_W4)
		Call PutGrid(C_W5, C_ROW_22, dblAmt(22, C_W5))
		
		dblAmt(23, C_W5) = dblAmt(20, C_W5) - dblAmt(21, C_W5) - dblAmt(22, C_W5)
		Call PutGrid(C_W5, C_ROW_23, dblAmt(23, C_W5))
		
	End With
	
	If pEvent = "" Then 
		Call SetReCalc3()	
	End If
	
	Call AllUpdateFlg
End Sub

	' 최저한세 클릭후 호출되는 계산식2 (그리드 Change이벤트에서는 불러지지 말아야한다)
Sub SetReCalc3()
	Dim dblAmt(25, 7)	' 7 = C_W5
	Dim dblA, dblB, dblX, iRow
	
	With frm1.vspdData
		' C_W4 전체 초기화 
		Call PutGrid(C_W4, C_ROW_05, 0)
		Call PutGrid(C_W4, C_ROW_06, 0)
		Call PutGrid(C_W4, C_ROW_13, 0)
		Call PutGrid(C_W4, C_ROW_14, 0)
		Call PutGrid(C_W4, C_ROW_17, 0)
		Call PutGrid(C_W4, C_ROW_21, 0)
		Call PutGrid(C_W4, C_ROW_22, 0)	
	
		' -- 06 (3) 최저한세 => 06. W4,W5 구하기 
		dblAmt(6, C_W3) = GetGrid(C_W3, C_ROW_06)
		dblAmt(6, C_W4) = GetGrid(C_W3, C_ROW_06)
		dblAmt(6, C_W5) = GetGrid(C_W3, C_ROW_06)
		Call PutGrid(C_W4, C_ROW_06, dblAmt(6, C_W4))
		Call PutGrid(C_W5, C_ROW_06, dblAmt(6, C_W5))
			
		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
			
		dblA = dblAmt(4, C_W5) + dblAmt(6, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) - dblAmt(16, C_W5)
			
		dblAmt(20, C_W3) = GetGrid(C_W3, C_ROW_20)
		dblAmt(21, C_W5) = GetGrid(C_W5, C_ROW_21)
		dblAmt(22, C_W5) = GetGrid(C_W5, C_ROW_22)
		dblB = dblAmt(20, C_W3) + dblAmt(21, C_W5) + dblAmt(22, C_W5)
			
		If dblA > 100000000 Then
			dblX = 100000000 * lgW2001(0) + (dblA - 100000000) * lgW2001(1)
		Else
			dblX = dblA * lgW2001(0) 
		End If
			
		If dblX > dblB Then
			' NO 코스 
			If dblA - dblAmt(6, C_W5) > 100000000 Then
				dblAmt(6, C_W4) = dblAmt(6, C_W3) - (( dblX - dblB) / lgW2001(1))
			ElseIf dblA < 100000000 Then	' NO 코스일때 
				dblAmt(6, C_W4) = dblAmt(6, C_W3) - (( dblX - dblB) / lgW2001(0))
			ElseIf dblB > (100000000 * lgW2001(0)) Then
				dblAmt(6, C_W4) = dblAmt(6, C_W3) - (( dblX - dblB) / lgW2001(1))
			Else
				dblAmt(6, C_W4) = dblAmt(6, C_W3) - (((dblX - (100000000 * lgW2001(0))) / lgW2001(1)) + ((100000000 * lgW2001(0)) - dblB) / lgW2001(0))
			End If	

			dblAmt(6, C_W5) = dblAmt(6, C_W4)

			Call PutGrid(C_W4, C_ROW_06, dblAmt(6, C_W4))
			Call PutGrid(C_W5, C_ROW_06, dblAmt(6, C_W5))	
			Call SetReCalc4()
			Exit Sub
		End If

							
		' -- 05 (3) 최저한세 => 05. W4,W5 구하기 
		dblAmt(5, C_W3) = GetGrid(C_W3, C_ROW_05)
		dblAmt(5, C_W4) = GetGrid(C_W3, C_ROW_05)
		dblAmt(5, C_W5) = GetGrid(C_W3, C_ROW_05)
		Call PutGrid(C_W4, C_ROW_05, dblAmt(5, C_W4))
		Call PutGrid(C_W5, C_ROW_05, dblAmt(5, C_W5))

		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(5, C_W5) = GetGrid(C_W5, C_ROW_05)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
			
		dblA = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5)  + dblAmt(8, C_W5) - dblAmt(9, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) - dblAmt(16, C_W5)
			
		dblAmt(20, C_W3) = GetGrid(C_W3, C_ROW_20)
		dblAmt(21, C_W5) = GetGrid(C_W5, C_ROW_21)
		dblAmt(22, C_W5) = GetGrid(C_W5, C_ROW_22)
		dblB = dblAmt(20, C_W3) + dblAmt(21, C_W5) + dblAmt(22, C_W5)
			
		If dblA > 100000000 Then
			dblX = 100000000 * lgW2001(0) + (dblA - 100000000) * lgW2001(1)
		Else
			dblX = dblA * lgW2001(0) 
		End If
			
		If dblX > dblB Then
			' NO 코스 
			If dblA - dblAmt(5, C_W5) > 100000000 Then
				dblAmt(5, C_W4) = dblAmt(5, C_W3) - (( dblX - dblB) / lgW2001(1))
			ElseIf dblA < 100000000 Then	' NO 코스일때 
				dblAmt(5, C_W4) = dblAmt(5, C_W3) - (( dblX - dblB) / lgW2001(0))
			ElseIf dblB > (100000000 * lgW2001(0)) Then
				dblAmt(5, C_W4) = dblAmt(5, C_W3) - (( dblX - dblB) / lgW2001(1))
			Else
				dblAmt(5, C_W4) = dblAmt(5, C_W3) - (((dblX - (100000000 * lgW2001(0))) / lgW2001(1)) + ((100000000 * lgW2001(0)) - dblB) / lgW2001(0))
			End If	
			dblAmt(5, C_W5) = dblAmt(5, C_W4)

			Call PutGrid(C_W4, C_ROW_05, dblAmt(5, C_W4))
			Call PutGrid(C_W5, C_ROW_05, dblAmt(5, C_W5))	
			Call SetReCalc4()
			Exit Sub
		End If

		
		' -- 14 (3) 최저한세 => 14. W4,W5 구하기 
		dblAmt(14, C_W3) = GetGrid(C_W3, C_ROW_14)
		dblAmt(14, C_W4) = GetGrid(C_W3, C_ROW_14)
		dblAmt(14, C_W5) = GetGrid(C_W3, C_ROW_14)
		Call PutGrid(C_W4, C_ROW_14, dblAmt(14, C_W4))
		Call PutGrid(C_W5, C_ROW_14, dblAmt(14, C_W5))

			
		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(5, C_W5) = GetGrid(C_W5, C_ROW_05)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(14, C_W5) = GetGrid(C_W5, C_ROW_14)
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
			
		dblA = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) + dblAmt(14, C_W5) - dblAmt(16, C_W5)
			
		dblAmt(20, C_W3) = GetGrid(C_W3, C_ROW_20)
		dblAmt(21, C_W5) = GetGrid(C_W5, C_ROW_21)
		dblAmt(22, C_W5) = GetGrid(C_W5, C_ROW_22)
		dblB = dblAmt(20, C_W3) + dblAmt(21, C_W5) + dblAmt(22, C_W5)
			
		If dblA > 100000000 Then
			dblX = 100000000 * lgW2001(0) + (dblA - 100000000) * lgW2001(1)
		Else
			dblX = dblA * lgW2001(0) 
		End If
			
		If dblX > dblB Then
			' NO 코스 
			If dblA - dblAmt(4, C_W5) > 100000000 Then
				dblAmt(14, C_W4) = dblAmt(14, C_W3) - (( dblX - dblB) / lgW2001(1))
			ElseIf dblA < 100000000 Then	' NO 코스일때 
				dblAmt(14, C_W4) = dblAmt(14, C_W3) - (( dblX - dblB) / lgW2001(0))
			ElseIf dblB > (100000000 * lgW2001(0)) Then
				dblAmt(14, C_W4) = dblAmt(14, C_W3) - (( dblX - dblB) / lgW2001(1))
			Else
				dblAmt(14, C_W4) = dblAmt(14, C_W3) - (((dblX - (100000000 * lgW2001(0))) / lgW2001(1)) + ((100000000 * lgW2001(0)) - dblB) / lgW2001(0))
			End If	
			dblAmt(14, C_W5) = dblAmt(14, C_W4)

			Call PutGrid(C_W4, C_ROW_14, dblAmt(14, C_W4))
			Call PutGrid(C_W5, C_ROW_14, dblAmt(14, C_W5))	
			Call SetReCalc4()
			Exit Sub
		End If

			
		' -- 22 (2) 최저한세 => 22. W4,W5 구하기 
		dblAmt(22, C_W2) = GetGrid(C_W2, C_ROW_22)
		dblAmt(22, C_W4) = GetGrid(C_W2, C_ROW_22)
		'dblAmt(22, C_W5) = GetGrid(C_W5, 22)
		Call PutGrid(C_W4, C_ROW_22, dblAmt(22, C_W4))
		'Call PutGrid(C_W5, 22, dblAmt(22, C_W5))
			
		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(5, C_W5) = GetGrid(C_W5, C_ROW_05)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(14, C_W5) = GetGrid(C_W5, C_ROW_14)
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
			
		dblA = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) + dblAmt(14, C_W5) - dblAmt(16, C_W5)
			
		dblAmt(20, C_W3) = GetGrid(C_W3, C_ROW_20)
		dblAmt(21, C_W5) = GetGrid(C_W5, C_ROW_21)
		dblB = dblAmt(20, C_W3) + dblAmt(21, C_W5)
			
		If dblA > 100000000 Then
			dblX = 100000000 * lgW2001(0) + (dblA - 100000000) * lgW2001(1)
		Else
			dblX = dblA * lgW2001(0) 
		End If
			
		If dblX > dblB Then
			' NO 코스 
			dblAmt(22, C_W4) = dblAmt(22, C_W2) - ( dblX - dblB)

			dblAmt(22, C_W5) = dblAmt(22, C_W4)

			Call PutGrid(C_W4, C_ROW_22, dblAmt(22, C_W4))
			Call PutGrid(C_W5, C_ROW_22, dblAmt(22, C_W5))	
			Call SetReCalc4()
			Exit Sub							
		End If
	

		' -- 21 (2) 최저한세 => 22. W4,W5 구하기 
		dblAmt(21, C_W2) = GetGrid(C_W2, C_ROW_21)
		dblAmt(21, C_W4) = GetGrid(C_W2, C_ROW_21)
		'dblAmt(21, C_W5) = GetGrid(C_W5, 21)
		Call PutGrid(C_W4, C_ROW_21, dblAmt(21, C_W4))
		'Call PutGrid(C_W5, 21, dblAmt(21, C_W5))
			
		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(5, C_W5) = GetGrid(C_W5, C_ROW_05)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(14, C_W5) = GetGrid(C_W5, C_ROW_14)
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
			
		dblA = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) + dblAmt(14, C_W5) - dblAmt(16, C_W5)
	
		dblAmt(20, C_W3) = GetGrid(C_W3, C_ROW_20)
		dblAmt(21, C_W5) = GetGrid(C_W5, C_ROW_21)
		dblB = dblAmt(20, C_W3) 
			
		If dblA > 100000000 Then
			dblX = 100000000 * lgW2001(0) + (dblA - 100000000) * lgW2001(1)
		Else
			dblX = dblA * lgW2001(0) 
		End If
			
		If dblX > dblB Then
			' NO 코스 
			dblAmt(21, C_W4) = dblAmt(21, C_W2) - ( dblX - dblB)

			dblAmt(21, C_W5) = dblAmt(21, C_W4)

			Call PutGrid(C_W4, C_ROW_21, dblAmt(21, C_W4))
			Call PutGrid(C_W5, C_ROW_21, dblAmt(21, C_W5))	
			Call SetReCalc4()
			Exit Sub			
		End If	

		' -- 17 (3) 최저한세 => 17. W4,W5 구하기 
		dblAmt(17, C_W3) = GetGrid(C_W3, C_ROW_17)
		dblAmt(17, C_W4) = GetGrid(C_W3, C_ROW_17)
		dblAmt(17, C_W5) = GetGrid(C_W3, C_ROW_17)
		Call PutGrid(C_W4, C_ROW_17, dblAmt(17, C_W4))
		Call PutGrid(C_W5, C_ROW_17, dblAmt(17, C_W5))
			
		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(5, C_W5) = GetGrid(C_W5, C_ROW_05)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(14, C_W5) = GetGrid(C_W5, C_ROW_14)
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
		dblAmt(17, C_W5) = GetGrid(C_W5, C_ROW_17)
			
		dblA = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) + dblAmt(14, C_W5) - dblAmt(16, C_W5) + dblAmt(17, C_W5)
		
		dblAmt(20, C_W3) = GetGrid(C_W3, C_ROW_20)
		dblB = dblAmt(20, C_W3) 
			
		If dblA > 100000000 Then
			dblX = 100000000 * lgW2001(0) + (dblA - 100000000) * lgW2001(1)
		Else
			dblX = dblA * lgW2001(0) 
		End If
			
		If dblX > dblB Then
			' NO 코스 
			If dblA - dblAmt(17, C_W5) > 100000000 Then
				dblAmt(17, C_W4) = dblAmt(17, C_W3) - (( dblX - dblB) / lgW2001(1))
			ElseIf dblA < 100000000 Then	' NO 코스일때 
				dblAmt(17, C_W4) = dblAmt(17, C_W3) - (( dblX - dblB) / lgW2001(0))
			ElseIf dblB > (100000000 * lgW2001(0)) Then
				dblAmt(17, C_W4) = dblAmt(17, C_W3) - (( dblX - dblB) / lgW2001(1))
			Else
				dblAmt(17, C_W4) = dblAmt(17, C_W3) - (((dblX - (100000000 * lgW2001(0))) / lgW2001(1)) + ((100000000 * lgW2001(0)) - dblB) / lgW2001(0))
			End If	
			dblAmt(17, C_W5) = dblAmt(17, C_W4)

			Call PutGrid(C_W4, C_ROW_17, dblAmt(17, C_W4))
			Call PutGrid(C_W5, C_ROW_17, dblAmt(17, C_W5))	
			Call SetReCalc4()
			Exit Sub
		End If

		' -- 13 (3) 최저한세 => 13. W4,W5 구하기 
		dblAmt(13, C_W3) = GetGrid(C_W3, C_ROW_13)
		dblAmt(13, C_W4) = GetGrid(C_W3, C_ROW_13)
		dblAmt(13, C_W5) = GetGrid(C_W3, C_ROW_13)
		Call PutGrid(C_W4, C_ROW_13, dblAmt(13, C_W4))
		Call PutGrid(C_W5, C_ROW_13, dblAmt(13, C_W5))

		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(5, C_W5) = GetGrid(C_W5, C_ROW_05)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(13, C_W5) = GetGrid(C_W5, C_ROW_13)
		dblAmt(14, C_W5) = GetGrid(C_W5, C_ROW_14)
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
		dblAmt(17, C_W5) = GetGrid(C_W5, C_ROW_17)
			
		dblA = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5)  + dblAmt(13, C_W5) + dblAmt(14, C_W5) - dblAmt(16, C_W5) + dblAmt(17, C_W5)
			
		dblAmt(20, C_W3) = GetGrid(C_W3, C_ROW_20)
		dblAmt(21, C_W5) = GetGrid(C_W5, C_ROW_21)
		dblB = dblAmt(20, C_W3) 
			
		If dblA > 100000000 Then
			dblX = 100000000 * lgW2001(0) + (dblA - 100000000) * lgW2001(1)
		Else
			dblX = dblA * lgW2001(0) 
		End If
			
		If dblX > dblB Then
			' NO 코스 
			If dblA - dblAmt(13, C_W5) > 100000000 Then
				dblAmt(13, C_W4) = dblAmt(13, C_W3) - (( dblX - dblB) / lgW2001(1))
			ElseIf dblA < 100000000 Then	' NO 코스일때 
				dblAmt(13, C_W4) = dblAmt(13, C_W3) - (( dblX - dblB) / lgW2001(0))
			ElseIf dblB > (100000000 * lgW2001(0)) Then
				dblAmt(13, C_W4) = dblAmt(13, C_W3) - (( dblX - dblB) / lgW2001(1))
			Else
				dblAmt(13, C_W4) = dblAmt(13, C_W3) - (((dblX - (100000000 * lgW2001(0))) / lgW2001(1)) + ((100000000 * lgW2001(0)) - dblB) / lgW2001(0))
			End If	
			dblAmt(13, C_W5) = dblAmt(13, C_W4)

			Call PutGrid(C_W4, C_ROW_13, dblAmt(13, C_W4))
			Call PutGrid(C_W5, C_ROW_13, dblAmt(13, C_W5))
			Call SetReCalc4()
			Exit Sub
		End If
			
		Call SetReCalc4()	
	End With	
End Sub

'  계산로직2를 수행하고  썸값 
Sub SetReCalc4()
	Dim dblAmt(25, 7)	' 7 = C_W5
	
	With frm1.vspdData
			
		' (3) 최저한세 금액 
		dblAmt(4, C_W5) = GetGrid(C_W5, C_ROW_04)
		dblAmt(5, C_W5) = GetGrid(C_W5, C_ROW_05)
		dblAmt(6, C_W5) = GetGrid(C_W5, C_ROW_06)
		
		dblAmt(7, C_W5) = dblAmt(4, C_W5) + dblAmt(5, C_W5) + dblAmt(6, C_W5)
		Call PutGrid(C_W5, C_ROW_07, dblAmt(7, C_W5))
		
		dblAmt(8, C_W5) = GetGrid(C_W5, C_ROW_08)
		dblAmt(9, C_W5) = GetGrid(C_W5, C_ROW_09)
		
		dblAmt(10, C_W5) = dblAmt(7, C_W5) + dblAmt(8, C_W5) - dblAmt(9, C_W5)
		Call PutGrid(C_W5, C_ROW_10, dblAmt(10, C_W5))
		
		dblAmt(11, C_W5) = GetGrid(C_W5, C_ROW_11)
		dblAmt(12, C_W5) = GetGrid(C_W5, C_ROW_12)
		dblAmt(13, C_W5) = GetGrid(C_W5, C_ROW_13)
		dblAmt(14, C_W5) = GetGrid(C_W5, C_ROW_14)
		
		dblAmt(15, C_W5) = dblAmt(10, C_W5) - dblAmt(11, C_W5) - dblAmt(12, C_W5) + dblAmt(13, C_W5) + dblAmt(14, C_W5)
		Call PutGrid(C_W5, C_ROW_15, dblAmt(15, C_W5))
		
		dblAmt(16, C_W5) = GetGrid(C_W5, C_ROW_16)
		dblAmt(17, C_W5) = GetGrid(C_W5, C_ROW_17)
		
		dblAmt(18, C_W5) = dblAmt(15, C_W5) - dblAmt(16, C_W5) + dblAmt(17, C_W5) 
		Call PutGrid(C_W5, C_ROW_18, dblAmt(18, C_W5))

		' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
		dblAmt(24, C_W5) = GetGrid(C_W3, C_ROW_24)
		dblAmt(25, C_W5) = dblAmt(24, C_W5) + dblAmt(18, C_W5) 
		Call PutGrid(C_W5, C_ROW_25, dblAmt(25, C_W5))	'(125) 과세표준금액 
		
		'If dblAmt(18, C_W5) > 100000000 Then
		If dblAmt(25, C_W5) > 100000000 Then
			'dblAmt(20, C_W5) = (dblAmt(18, C_W5) * lgW2001(1)) - (100000000 * (lgW2001(1) - lgW2001(0)))
			dblAmt(20, C_W5) = (dblAmt(25, C_W5) * lgW2001(1)) - (100000000 * (lgW2001(1) - lgW2001(0)))
		Else
			'dblAmt(20, C_W5) = (dblAmt(18, C_W5) * lgW2001(0)) 
			dblAmt(20, C_W5) = (dblAmt(25, C_W5) * lgW2001(0)) 
		End If
		Call PutGrid(C_W5, C_ROW_20, dblAmt(20, C_W5))

		dblAmt(21, C_W2) = GetGrid(C_W2, C_ROW_21)
		dblAmt(21, C_W4) = GetGrid(C_W4, C_ROW_21)
		dblAmt(21, C_W5) = dblAmt(21, C_W2) - dblAmt(21, C_W4)
		Call PutGrid(C_W5, C_ROW_21, dblAmt(21, C_W5))
		
		dblAmt(22, C_W2) = GetGrid(C_W2, C_ROW_22)
		dblAmt(22, C_W4) = GetGrid(C_W4, C_ROW_22)
		dblAmt(22, C_W5) = dblAmt(22, C_W2) - dblAmt(22, C_W4)
		Call PutGrid(C_W5, C_ROW_22, dblAmt(22, C_W5))
		
		dblAmt(23, C_W5) = dblAmt(20, C_W5) - dblAmt(21, C_W5) - dblAmt(22, C_W5)
		Call PutGrid(C_W5, C_ROW_23, dblAmt(23, C_W5))
				
	End With

End Sub

'  최저한세 링크 클릭시 호출됨 
Sub SetReCalc5()
	Dim dblAmt(25, 7)	' 7 = C_W5
	
	With frm1.vspdData

		' -- (104) 조정후 소득금액 (02)			감면후 금액을 (5)조정후세액으로 이동 
		dblAmt(4, C_W2) = GetGrid(C_W2, C_ROW_04)
		Call PutGrid(C_W5, C_ROW_04, dblAmt(4, C_W2))

		dblAmt(7, C_W2) = GetGrid(C_W2, C_ROW_07)
		Call PutGrid(C_W5, C_ROW_07, dblAmt(7, C_W2))

		dblAmt(10, C_W2) = GetGrid(C_W2, C_ROW_10)
		Call PutGrid(C_W5, C_ROW_10, dblAmt(10, C_W2))

		dblAmt(15, C_W2) = GetGrid(C_W2, C_ROW_15)
		Call PutGrid(C_W5, C_ROW_15, dblAmt(15, C_W2))

		dblAmt(18, C_W2) = GetGrid(C_W2, C_ROW_18)
		Call PutGrid(C_W5, C_ROW_18, dblAmt(18, C_W2))

		dblAmt(25, C_W2) = GetGrid(C_W2, C_ROW_25)
		Call PutGrid(C_W5, C_ROW_25, dblAmt(25, C_W2))

		dblAmt(19, C_W2) = GetGrid(C_W2, C_ROW_19)
		Call PutGrid(C_W5, C_ROW_19, dblAmt(19, C_W2))

		dblAmt(20, C_W2) = GetGrid(C_W2, C_ROW_20)
		Call PutGrid(C_W5, C_ROW_20, dblAmt(20, C_W2))

		dblAmt(21, C_W2) = GetGrid(C_W2, C_ROW_21)
		Call PutGrid(C_W5, C_ROW_21, dblAmt(21, C_W2))

		dblAmt(22, C_W2) = GetGrid(C_W2, C_ROW_22)
		Call PutGrid(C_W5, C_ROW_22, dblAmt(22, C_W2))

		dblAmt(23, C_W2) = GetGrid(C_W2, C_ROW_23)
		Call PutGrid(C_W5, C_ROW_23, dblAmt(23, C_W2))

	End With
End Sub

Sub AllUpdateFlg()
	Dim iRow, iMaxRows
	If lgBlnFlgChgValue = False Then Exit Sub
	With frm1.vspdData
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			ggoSpread.UpdateRow iRow
		Next
	End With
End Sub

'============================================  조회조건 함수  ====================================

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitVariables                                                      <%'Initializes local global variables%>
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

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

Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	Call GetFISC_DATE
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
		If frm1.cboREP_TYPE.value = "2" Then
			lgMonGap = 6
		Else
			lgMonGap = DateDiff("m", datFISC_START_DT, datFISC_END_DT)+1
		End If
	Else
		lgMonGap = 12
	End If
	
	ret = CommonQueryRs("W1"," dbo.ufn_TB_4_GetRate('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If ret Then
		lgW2019 = UNICDbl(lgF0)
	
	End If
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If

	lgBlnFlgChgValue = True
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	With frm1.vspdData

		Select Case Col
			Case C_W2, C_W3
					
				Call SetReCalc1
			
			Case C_W4
				'MsgBox "유자가 W4열을 수정하여 발생해야 합니다"
				Call SetReCalc2("1")
			Case C_W5
				.Col =C_W5
				If Row=25 Then 					
					If  compVal(uniCdbl(.Text)) =False Then 												
						Exit Sub
					End If
				End If
		End Select 
			
	End With

	
End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
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

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    

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

    Call SetToolbar("1100100000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue Then
		ggoSpread.Source = frm1.vspdData
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

Function FncSave() 
    Dim blnChange, dblSum
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    
		
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
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
                                                '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
     
End Function


Function FncDeleteRow() 

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
	
    ggoSpread.Source = frm1.vspdData	
    If lgBlnFlgChgValue = True Then
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
	lgIntFlgMode = parent.OPMD_UMODE
	
	Call SpreadInitData
	
	ggoSpread.Source = frm1.vspdData
	
	' -- 2006-01-02 : 200603개정판 적용 (2행 밀림)
	
	With frm1.vspdData
	For iDx = 1 To C_ROW_23
		.Row = iDx
		.Col = 0
		.Text =iDx 'ggoSpread.UpdateFlag
		
'		ggoSpread.UpdateRow iDx
	Next	
	End With
	
	' 세무정보 조사 : 컨펌되면 락된다.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 컨펌체크 : 그리드 락 
	If wgConfirmFlg = "N" Then
		Call SetToolbar("1101100000000111")		
	Else
		Call SetToolbar("1100000000000111")		
	End If

	frm1.vspdData.focus			
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
	Call InitVariables
	Call FncNew()
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
    Dim strVal, strDel, lMaxRows, lMaxCols
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
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
		                                          strDel = strDel & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = 1 To lMaxCols
				
					Select Case lRow
						Case C_ROW_19
							.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
						Case C_ROW_23
							If lCol =C_W5 Then
								.Col =lCol : .Row=lRow
								If compVal(uniCdbl(.text))=False Then 
									Call LayerShowHide(0)									
									Exit Function 								
								Else
								.Col =lCol : .Row=lRow
								.Col = lCol : strVal = strVal & Trim(.text) &  Parent.gColSep
								End If
							Else
								.Col = lCol : strVal = strVal & Trim(.value) &  Parent.gColSep
							End IF
						Case Else
							.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
					End Select
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With

	frm1.txtMode.value        =  Parent.UID_M0002
    frm1.txtSpread.value      = strDel & strVal
    strVal = ""
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function
'-------------------------------------------------------------------------------------------------------------------------------------
'compared the value of (123)'s 5 and (120)'s 3 
'-------------------------------------------------------------------------------------------------------------------------------------
Function compVal(byVal strVal)
	Dim tmpTxt
	dim tmp123
	compVal= True
	
	with frm1.vspdData
		.Row= 22 : .Col = C_W3
		tmpTxt=uniCdbl(.Text)
		
		.Row= 25 : .Col = C_W5
		tmp123=uniCdbl(.value)

		
		If tmp123<tmpTxt Then
			Call DisplayMsgBox("WC0017", "X", "조정 후 차감세액 금액", "최저한세 금액") 
			compVal=False
			.Row=25:.Col = C_W5
			Exit Function 
		End If	
	End With

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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:GetRef">금액 불러오기</A> | <A href="vbscript:GetRef2">최저한세 계산</A></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

