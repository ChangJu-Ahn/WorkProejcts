
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 준비금조정 
'*  3. Program ID           : W4105MA1
'*  4. Program Name         : W4105MA1.asp
'*  5. Program Desc         : 제5호 특별비용조정명세 
'*  6. Modified date(First) : 2005/01/18
'*  7. Modified date(Last)  : 2005/01/18
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

Const BIZ_MNU_ID		= "W4105MA1"
Const BIZ_PGM_ID		= "W4105mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W4105mb2.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID	    = "W4105OA1"

Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 

' -- 그리드 컬럼 정의 
Dim C_SEQ_NO
Dim C_W1_CD
Dim C_W1
Dim C_W2
Dim C_W2_CD
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(2)

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO	= 1
	C_W1_CD		= 2	' -- 1번 그리드 
	C_W1		= 3	' 구분 
	C_W2		= 4	' 근거법조항 
	C_W2_CD		= 5 ' 코드 
	C_W3		= 6	' 회사계상액 
	C_W4		= 7 ' 한도초과액 
	C_W5		= 8	' 차감액 
	C_W6		= 9	' 최저한세적용손금부인액 
	C_W7		= 10	' 손금불산입액계 
	C_W8		= 11	' 손금산입계 
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
	
    Call initSpreadPosVariables()  

	' 1번 그리드 
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W8 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    

		'헤더를 3줄로    
		.ColHeaderRows = 2    
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","4","0")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"순번", 10,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_W1_CD,	"MINOR_CD", 10,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit     C_W1,	    "(1)구 분", 40,,,60,1 
		ggoSpread.SSSetEdit		C_W2,		"근거법조항"	, 12,,,50,1 
		ggoSpread.SSSetEdit		C_W2_CD,	"코드", 5,2,,2,1 
		ggoSpread.SSSetFloat	C_W3,		"(2)회사계상액"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W4,		"(3)한도초과액"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W5,		"(4)차감액"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W6,		"(5)최저한세" & vbCrLf & "적용" & vbCrLf & "손금부인액"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W7,		"(6)손금불산입계" & vbCrLf & "[(3)+(5)]" , 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W8,		"(7)손금산입계" & vbCrLf & "[(2)-(6)]"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

		' 그리드 헤더 합침 
		ret = .AddCellSpan(C_SEQ_NO , -1000, 1, 2)
		ret = .AddCellSpan(C_W1_CD	, -1000, 1, 2)	' 순번 2행 합침 
		ret = .AddCellSpan(C_W1		, -1000, 1, 2)	
		ret = .AddCellSpan(C_W2		, -1000, 1, 2)
		ret = .AddCellSpan(C_W2_CD	, -1000, 1, 2)
		ret = .AddCellSpan(C_W3		, -1000, 3, 1)
		ret = .AddCellSpan(C_W6		, -1000, 1, 2)
		ret = .AddCellSpan(C_W7		, -1000, 1, 2)
		ret = .AddCellSpan(C_W8		, -1000, 1, 2)
    
		' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W3	: .Text = "조 정 명 세 서 내 역"
		
		.Row = -999
		.Col = C_W3	: .Text = "(2)회사계상액"
		.Col = C_W4	: .Text = "(3)한도초과액"
		.Col = C_W5	: .Text = "(4)차감액"
			
		.rowheight(-999) = 20	' 높이 재지정	(2줄일 경우, 1줄은 15)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W1_CD,C_W1_CD,True)
		Call ggoSpread.SSSetColHidden(C_W2,C_W2,True)
				
		Call SetSpreadLock(TYPE_1)
				
		.ReDraw = true	
			
	End With 
  
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
       
    ' 그리드 초기 데이타셋팅 
    Dim arrMinorCd, arrW1, arrW2, arrW2_CD, iMaxRows, iRow, iMinorCnt
	'call CommonQueryRs("MINOR_CD, MINOR_NM, REFERENCE_1, REFERENCE_2","ufn_TB_Configuration('W1054', '" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    'arrMinorCd	= Split(lgF0, Chr(11))
    'arrW1		= Split(lgF1, Chr(11))
    'arrW2		= Split(lgF2, Chr(11))
    'arrW2_CD	= Split(lgF3, Chr(11))
    
    'iMinorCnt = 18	' 마이너코드 총 갯수는 18개로 하드코딩 
	'iMaxRows = UBound(arrMinorCd)
	iMaxRows = 22
	
	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		ggoSpread.InsertRow , iMaxRows
		
		' 배열을 그리드에 삽입 
		For iRow = 1 To iMaxRows
			
			.Row = iRow
			.Col = C_SEQ_NO : .value = iRow
			.Col = C_W1_CD	: .value = 100 + iRow
		'	.Col = C_W1		: .value = arrW1(iRow-1)
		'	.Col = C_W2		: .value = arrW2(iRow-1)
		'	.Col = C_W2_CD	: .value = arrW2_CD(iRow-1)
			
			
		'	If Trim(arrW1(iRow-1)) <> "" Then
		'		ggoSpread.SpreadLock C_W1, iRow, C_W2_CD, iRow
		'		ggoSpread.SpreadLock C_W5, iRow, C_W5, iRow
		'		ggoSpread.SpreadLock C_W7, iRow, C_W8, iRow
		'	End If
			
		Next
		
		'If iMinorCnt > iMaxRows Then
		'	ggoSpread.InsertRow iMaxRows, iMinorCnt - iMaxRows + 4
		'	
		'	For iRow = iMaxRows To iMinorCnt + 4
		'		.Row = iRow
		'		.Col = C_SEQ_NO : .value = iRow			
		'	Next
		'End If

		'ggoSpread.SpreadLock C_W1, 1, C_W2_CD, 8
		'ggoSpread.SpreadLock C_W5, iRow, C_W5, iRow
		'ggoSpread.SpreadLock C_W7, iRow, C_W8, iRow
				
		.Row = 1
		.Col = C_W1		: .Value = "(101)주권상장중소기업 등의 사업손실준비금(제8조의2)"
		.Col = C_W2_CD	: .Value = "42"
		
		.Row = 2
		.Col = C_W1		: .Value = "(102)연구및인력개발준비금(제9조)"
		.Col = C_W2_CD	: .Value = "02"

		.Row = 3
		.Col = C_W1		: .Value = "(103)사회간접자본투자준비금(제28조)"
		.Col = C_W2_CD	: .Value = "08"

		.Row = 4
		.Col = C_W1		: .Value = "(104)부동산투자회사투자손실준비금(제55조의2)"
		.Col = C_W2_CD	: .Value = "45"

		.Row = 5
		.Col = C_W1		: .Value = "(105)100%손금산입고유목적 사업준비금(제74조제1항)"
		.Col = C_W2_CD	: .Value = "17"

		.Row = 6
		.Col = C_W1		: .Value = "(106)80%손금산입고유목적 사업준비금(제74조제2항)"
		.Col = C_W2_CD	: .Value = "18"

		.Row = 7
		.Col = C_W1		: .Value = "(107)주권상장기업 등의 자사주 처분손실준비금(제104의3)"
		.Col = C_W2_CD	: .Value = "43"

		.Row = 8
		.Col = C_W1		: .Value = "(108)문화사업준비금(제104의9)"
		.Col = C_W2_CD	: .Value = "47"

		.Row = 9
		.Col = C_W1		: .Value = "(109)"
		.Col = C_W2_CD	: .Value = "44"

		.Row = 10
		.Col = C_W1		: .Value = "(110)"
		.Col = C_W2_CD	: .Value = ""

		.Row = 11
		.Col = C_W1		: .Value = "(111)"
		.Col = C_W2_CD	: .Value = ""

		.Row = 12
		.Col = C_W1		: .Value = "(112)"
		.Col = C_W2_CD	: .Value = ""
		
		.Row = 13
		.Col = C_W1		: .Value = "(113)"

		.Row = 14
		.Col = C_W1		: .Value = "(114)"

		.Row = 15
		.Col = C_W1		: .Value = "(115)"

		.Row = 16
		.Col = C_W1		: .Value = "(116)"
		
		.Row = 17
		.Col = C_W1		: .Value = "(117)"

		.Row = 18
		.Col = C_W1		: .Value = "(118)"

		
		
		' 준비금계(119), 준비금 및 특별감가상각비 계(142)는 회색 처리 
		.Row = 19
		'.Col = C_W1_CD	: .Value = "119"
		.Col = C_W1		: .Value = "(119)준비금 계"
		.Col = C_W2_CD	: .Value = "19"

		.Row = 20
		'.Col = C_W1_CD	: .Value = "140"
		.Col = C_W1		: .Value = "(140)특별감가상각비 계"
		.Col = C_W2_CD	: .Value = "40"

		.Row = 21
		'.Col = C_W1_CD	: .Value = "141"
		.Col = C_W1		: .Value = "(141)특례자산감가상각비 계(제30조)"
		.Col = C_W2_CD	: .Value = "46"

		.Row = 22
		'.Col = C_W1_CD	: .Value = "142"
		.Col = C_W1		: .Value = "준비금 및 특별감가상각비 계(119+140+141)"
		.Col = C_W2_CD	: .Value = "41"

		Call SetSpreadLock_Query(TYPE_1)
				
		lgvspdData(TYPE_1).Redraw = True
	End With

End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	Select Case pType
		Case TYPE_1
			ggoSpread.SpreadLock C_SEQ_NO, -1, C_W1_CD
			ggoSpread.SpreadLock C_W5, -1, C_W5
			ggoSpread.SpreadLock C_W7, -1, C_W7
			ggoSpread.SpreadLock C_W8, -1, C_W8
	End Select
	
End Sub

Sub SetSpreadLock_Query(Byval pType)
	Dim iRow

	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		ggoSpread.SpreadLock C_SEQ_NO, -1, C_W1_CD
		ggoSpread.SpreadLock C_W1, 1, C_W1, 8
		ggoSpread.SpreadLock C_W2_CD, 1, C_W2_CD, 9
		ggoSpread.SpreadLock C_W2_CD,19, C_W2_CD, 22
		ggoSpread.SpreadLock C_W5, -1, C_W5
		ggoSpread.SpreadLock C_W7, -1, C_W7
		ggoSpread.SpreadLock C_W8, -1, C_W8
		ggoSpread.SpreadLock C_SEQ_NO, 20, C_W2_CD, 21
		ggoSpread.SpreadLock C_SEQ_NO, 19, C_W8, 19
		ggoSpread.SpreadLock C_SEQ_NO, 22, C_W8, 22
				
		lgvspdData(TYPE_1).Redraw = True
	End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.Source = lgvspdData(pType)

	Select Case pType
		Case TYPE_1
	End Select
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"

    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
			
    ' 그리드 초기 데이타셋팅 
	call CommonQueryRs("W5, W6","TB_31_1H "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		.Row = 1
		.Col = C_W3 : .Value = Replace(lgF0, chr(11), "")
		.Col = C_W4 : .Value = Replace(lgF1, chr(11), "")
		
		Call vspdData_Change(TYPE_1, C_W3, 1)
		Call vspdData_Change(TYPE_1, C_W4, 1)
		
		lgvspdData(TYPE_1).Redraw = True
	End With

	call CommonQueryRs("W4, W5","TB_31_2H "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		.Row = 3
		.Col = C_W3 : .Value = Replace(lgF0, chr(11), "")
		.Col = C_W4 : .Value = Replace(lgF1, chr(11), "")
		
		Call vspdData_Change(TYPE_1, C_W3, 3)
		Call vspdData_Change(TYPE_1, C_W4, 3)
		
		lgvspdData(TYPE_1).Redraw = True
	End With
	
End Function

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = "../W5/W5105RA1.ASP"
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
   

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function
'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000101111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
	Call InitData 
	
    Call MainQuery() 
    
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
	'Call GetFISC_DATE
End Sub

'============================================  그리드 이벤트   ====================================
' -- 0번 그리드 
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

' -- 1번 그리드 
Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
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

'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, dblAmt(20), dblW3, dblW4, dblW5, dblW6, dblW7, dblW8
	Dim sCoCd,sFiscYear,sRepType
	
	sCoCd		= "<%=wgCO_CD%>"
	sFiscYear	= "<%=wgFISC_YEAR%>"
	sRepType	= "<%=wgREP_TYPE%>"
	' 법인정보 출력 
	
	lgBlnFlgChgValue= True ' 변경여부 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(lgvspdData(Index).text) < CDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row

	' --- 추가된 부분 
	With lgvspdData(Index)

	lgvspdData(Index).Redraw = False
	
	If Index = TYPE_1 Then	'1번 그리 
	
		Select Case Col
			Case C_W3, C_W4, C_W6	' 회사계상액/한도초과액 
			
				.Col = C_W3
				
				IF .Row = 1 THEN 
					Call CommonQueryRs("W5"," TB_31_1H(NOLOCK) ","  CO_CD = '" & sCoCd & "' and FISC_YEAR = '" & sFiscYear & "' and REP_TYPE =  '" & sRepType & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					if UNICDbl(.Value) <> UNICDbl(Replace(lgF0,Chr(11),""))  then
						Call DisplayMsgBox("WC0004", parent.VB_INFORMATION, "특별비용 조정명세서", "중소기업투자준비금 조정명세서")
					   .Value = UNICDbl(Replace(lgF0,Chr(11),""))
					End if 
				End IF
				.Row = Row 
				.Col = C_W3 :dblAmt(C_W3) = UNICDbl(.Value)
				.Col = C_W4 :dblAmt(C_W4) = UNICDbl(.Value)
				.Col = C_W6 :dblAmt(C_W6) = UNICDbl(.Value)
				
				dblAmt(C_W5) = dblAmt(C_W3) - dblAmt(C_W4)
				If dblAmt(C_W5) < 0 Then
					Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, GetGrid(C_W4, -999), GetGrid(C_W3, -999))
					.Row = Row
					.Col = Col	: .value = 0
					.Col = Col	: dblAmt(Col) = 0
					dblAmt(C_W5) = dblAmt(C_W3) - dblAmt(C_W4)
				End If

				dblAmt(C_W7) = dblAmt(C_W4) + dblAmt(C_W6)	
				dblAmt(C_W8) = dblAmt(C_W3) - dblAmt(C_W7)
				If dblAmt(C_W8) < 0 Then
					Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, GetGrid(C_W7, 0), GetGrid(C_W3, -999))
					.Row = Row
					.Col = Col	: .value = 0
					.Col = Col	: dblAmt(Col) = 0
					dblAmt(C_W7) = dblAmt(C_W4) + dblAmt(C_W6)
					dblAmt(C_W8) = dblAmt(C_W3) - dblAmt(C_W7)
				End If
						
				.Col = C_W5 : .Value = dblAmt(C_W5)
				.Col = C_W7 : .Value = dblAmt(C_W7)
				.Col = C_W8 : .Value = dblAmt(C_W8)

				' -- 현재 변경된 컬럼의 준비금 누계 
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, 19 - 1, true, 19, Col, "V")	' 합계 
				
				' -- C_W5 차감액 준비금 누계 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W5, 1, 19 - 1, true, 19, C_W5, "V")	' 합계 
				' -- C_W7
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W7, 1, 19 - 1, true, 19, C_W7, "V")	' 합계 
				' -- C_W8
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W8, 1, 19 - 1, true, 19, C_W8, "V")	' 합계 
				
				ggoSpread.UpdateRow 19

				
				' -- 현재 변경된 컬럼의 준비금 누계 
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 19, .MaxRows- 1, true, .MaxRows, Col, "V")	' 합계 
				
				' -- C_W5 차감액 준비금 누계 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W5, 19, .MaxRows- 1 - 1, true, .MaxRows, C_W5, "V")	' 합계 
				' -- C_W7
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W7, 19, .MaxRows- 1 - 1, true, .MaxRows, C_W7, "V")	' 합계 
				' -- C_W8
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W8, 19, .MaxRows- 1 - 1, true, .MaxRows, C_W8, "V")	' 합계				
				
				ggoSpread.UpdateRow .MaxRows
		End Select
	
	End If
	
	lgvspdData(Index).Redraw = True
	
	End With
	
End Sub

Function GetGrid(Byval Col, Byval Row)
	With lgvspdData(TYPE_1)
		.Col = Col : .Row = Row : GetGrid = .value
	End With
End Function

' -- 2번째 그리드 
Sub SetGridTYPE_2()
	Dim dblSum, dblW11, dblW12, dblW13 , dblFiscYear, dblW26, dblW25, dblW24, dblW23, dblW22, dblW17, dblW15, dblW14
	Dim dblW27, dblW28, dblW29, dblW30, dblW31, dblW32, dblW33

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
	With lgvspdData(TYPE_2)
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


' 헤더 변경 
Sub SetHeadReCalc()	
	Dim dblSum, dblW16, dblW3, dblW2, dblW4, dblW5, dblW6, dblW7
	
	With lgvspdData(TYPE_1)
		.Col = C_W16 : .Row = .MaxRows : dblW16 = UNICDbl(.Value)
	End With	
	
	With frm1
		.txtW1.value = dblW16
		dblW2 = UNICDbl(.txtW2_VAL.value)
		dblW3 = UNICDbl(.txtW3_VAL.value)
		dblW4 = UNICDbl(UNIFormatNumber(dblW16 * dblW2 * dblW3 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit))  ' 소수점은 절하하는 이유로 해당함수를 사용함 
		.txtW4.value = dblW4
		dblW5 = UNICDbl(.txtW5.value)
		If (dblW5 - dblW4) > 0 Then
			dblW6 = dblW5 - dblW4
		Else	
			dblW6 = 0
		End If
		dblW7 = UNICDbl(.txtW7.value)
		
		.txtW6.value = dblW6
		.txtW8.value = dblW6 + dblW7
		.txtW9.value = dblW5 - dblW6 - dblW7
	End With
End Sub

' 잔액 컬럼이 변경될때 호출됨 
Sub SetW17_W18(Row)
	Dim dblSum, dblW4, dblW15, dblW16, dblW18
	
	With lgvspdData(TYPE_1)
		
		.Row = Row
		.Col = C_W11	: dblW11 = UNICDbl(.Value)	' 차변 
		.Col = C_W12	: dblW12 = UNICDbl(.Value)	' 차변 
		.Col = C_W13	: dblW13 = UNICDbl(.Value)	' 차변 
		
		.Col = C_W14	: dblW14 = UNICDbl(.Value)	' 차변 
		.Col = C_W15	: dblW15 = UNICDbl(.Value)	' 대변 
		.Col = C_W16	: dblW16 = UNICDbl(.Value)	' 차변 
		.Col = C_W18	: dblW18 = UNICDbl(.Value)	' 대변 
		
		.Col = C_W18	: dblW18 = dblW11 - dblW12 - dblW13				: .Value = dblW18
		.Col = C_W17	: dblW17 = dblW18 - dblW14 - dblW15 - dblW16	: .Value = dblW17

		Call FncSumSheet(lgvspdData(TYPE_2), C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' 합계 
		Call FncSumSheet(lgvspdData(TYPE_2), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' 합계 
	End With
	
End Sub

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(Index)
   
    If lgvspdData(Index).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    'If Row <= 0 Then
    '   ggoSpread.Source = lgvspdData(Index)
    '   
    '   If lgSortKey = 1 Then
    '       ggoSpread.SSSort Col               'Sort in ascending
    '       lgSortKey = 2
    '   Else
    '       ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
    '       lgSortKey = 1
    '   End If
       
    '   Exit Sub
    'End If

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
    'Call InitData                              
    															
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
	ggoSpread.Source = lgvspdData(TYPE_1)
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
	
	If blnChange = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If


    If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

' ---------------------- 서식내 검증 -------------------------
Function  Verification()

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

    Call SetToolbar("1100100000001111")

	'Call ClickTab1()
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

    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    
    Call CheckReCalc()				' 한라인이 취소되면 재계산 
    Call CheckW7Status(lgCurrGrid)	' 적수셀 상태 체크 
End Function

' 재계산 
Function CheckReCalc()
	Dim dblSum
	
	With lgvspdData(lgCurrGrid)
		ggoSpread.Source = lgvspdData(lgCurrGrid)	
	
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' 합계 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 합계 
	
		' 자신이 변경되면 다른 호출해야 할 부분을 아래에 기술 
		Call SetW5(lgCurrGrid, 1)
					
		' 일수 변경 요청 
		Call SetW6(lgCurrGrid, 1)
					
		Call SetW7(lgCurrGrid, 1)	' 적수 계산 
	End With
End Function

Function FncInsertRow(ByVal pvRowCnt) 
  
    
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
		lDelRows = ggoSpread.DeleteRow
	End With
	
	Call CheckReCalc()				' 한라인이 취소되면 재계산 
	
	Call CheckW7Status(lgCurrGrid)	' 적수셀 상태 체크 
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


Function DbQueryFalse()
	Call InitData
End Function


Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If lgvspdData(TYPE_1).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
			Call SetSpreadLock_Query(TYPE_1)
			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101100000011111")										<%'버튼 툴바 제어 %>
			
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100100000011111")										<%'버튼 툴바 제어 %>
		End If
	Else
		Call SetToolbar("1100100000111111")										<%'버튼 툴바 제어 %>
	End If
	
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
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    For i = TYPE_1 To TYPE_1	' 전체 그리드 갯수 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
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
		                                              strDel = strDel & "D"  &  Parent.gColSep
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
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">금액불러오기</A>|<A href="vbscript:OpenRefMenu">소득금액합계표조회</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w4105ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>

						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=* VALIGN=TOP><script language =javascript src='./js/w4105ma1_vspdData0_vspdData0.js'></script></TD>
									  </TR>
							
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
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

