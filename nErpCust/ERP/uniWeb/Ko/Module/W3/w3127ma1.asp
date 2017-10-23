
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 각 과목별 조정 
'*  3. Program ID           : W3127MA1
'*  4. Program Name         : W3127MA1.asp
'*  5. Program Desc         : 제26호 업무무관부동산등에 관련한 차입금이자조정명세서(을)
'*  6. Modified date(First) : 2005/01/05
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

Const BIZ_MNU_ID		= "W3127MA1"
Const BIZ_PGM_ID		= "w3127mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "w3127mb2.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID		= "W3127OA1"

Const TAB1 = 1																	'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.
Const TYPE_3	= 2		
Const TYPE_4A	= 3
Const TYPE_4B	= 4
Const TYPE_5	= 5
Const TYPE_6	= 6

' -- 그리드 컬럼 정의 
Dim C_SEQ_NO
Dim C_W_TYPE
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7

Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(6)
Dim lgFISC_START_DT, lgFISC_END_DT, lgW12_REF ' 자본금+미지급.....

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1	' -- 1번 그리드 
    C_W_TYPE	= 2	' 구분 
    C_W1		= 3	' 연월일 
    C_W2		= 4 ' 적요 
    C_W3		= 5	' 차변 
    C_W4		= 6	' 대변 
    C_W5		= 7	' 잔액 
    C_W6		= 8	' 일수 
    C_W7		= 9	' 적수 

	C_W10		= 5	' 자산총계 
	C_W11		= 6	' 부채총계 
	C_W12		= 7	' 자기자본 
	C_W13		= 8 ' 연일수 
	C_W14		= 9	' 적수 
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
    
    lgW12_REF = 0
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
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	Set lgvspdData(TYPE_3) = frm1.vspdData2
	Set lgvspdData(TYPE_4A) = frm1.vspdData3
	Set lgvspdData(TYPE_4B) = frm1.vspdData4
	Set lgvspdData(TYPE_5) = frm1.vspdData5
	Set lgvspdData(TYPE_6) = frm1.vspdData6
	
	lgvspdData(TYPE_1).ScriptEnhanced  = True
	lgvspdData(TYPE_2).ScriptEnhanced  = True
	lgvspdData(TYPE_3).ScriptEnhanced  = True
	lgvspdData(TYPE_4A).ScriptEnhanced  = True
	lgvspdData(TYPE_4B).ScriptEnhanced  = True
	lgvspdData(TYPE_5).ScriptEnhanced  = True
	lgvspdData(TYPE_6).ScriptEnhanced  = True
	
    Call initSpreadPosVariables()  

	' 1번-5번 그리드 

	For iRow = TYPE_1 To TYPE_5		' 그리드 정의 상수 
		With lgvspdData(iRow)
			
			ggoSpread.Source = lgvspdData(iRow)	
			'patch version
			ggoSpread.Spreadinit "V20041222_" & iRow,,parent.gForbidDragDropSpread    
    
			.ReDraw = false

			.MaxCols = C_W7 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
			.Col = .MaxCols									'☆: 사용자 별 Hidden Column
			.ColHidden = True    
				       
			.MaxRows = 0
			ggoSpread.ClearSpreadData

			'Call AppendNumberPlace("6","3","2")

			ggoSpread.SSSetEdit		C_SEQ_NO,	"순번"		, 10,,,6,1	' 히든컬럼 
			ggoSpread.SSSetEdit		C_W_TYPE,	"구 분"		, 10,,,6,1	' 히든컬럼 
			ggoSpread.SSSetDate		C_W1,		"(1)연월일"	, 10, 2, Parent.gDateFormat, -1
			ggoSpread.SSSetEdit		C_W2,		"(2)적 요"	, 15,,,50,1
			ggoSpread.SSSetFloat	C_W3,		"(3)차 변"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W4,		"(4)대 변"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W5,		"(5)잔 액"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W6,		"(6)일수"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
			ggoSpread.SSSetFloat	C_W7,		"(7)적 수"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
			Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
				
			Call SetSpreadLock(iRow)
				
			.ReDraw = true	
			
		End With 
	Next
 
	With lgvspdData(TYPE_6)
	
		' 자기자본 적수계산 그리드 정의 
 		ggoSpread.Source = lgvspdData(TYPE_6)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_6,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W14 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    

		ggoSpread.ClearSpreadData
		.MaxRows = 0
		'Call AppendNumberPlace("6","3","2")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"순번"		, 10,,,6,1	' 히든컬럼 
		ggoSpread.SSSetEdit		C_W_TYPE,	"구 분"		, 10,,,6,1	' 히든컬럼 
		ggoSpread.SSSetDate		C_W1,		"Blank"			, 10, 2, Parent.gDateFormat, -1
		ggoSpread.SSSetFloat	C_W2,		"Blank"			, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W10,		"(10)대차대조표" & vbCrLf & "자산총계", 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W11,		"(11)대차대조표" & vbCrLf & "부채총계", 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W12,		"(12)자기자본 [(10)-(11)]"	, 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W13,		"(13)사업연도" & vbCrLf & "일수", 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W14,		"(14)적 수"	, 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
		Call ggoSpread.SSSetColHidden(C_W2,C_W2,True)
		
		.rowheight(-1000) = 20	' 높이 재지정	(2줄일 경우, 1줄은 15)
		
		Call SetSpreadLock(TYPE_6)
			
		.ReDraw = true	 
    
    End With
     
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
    
	With lgvspdData(TYPE_6)
		ggoSpread.Source = lgvspdData(TYPE_6)
		
		ggoSpread.InsertRow ,1
		SetSpreadColor TYPE_6, 1, 1
		
		.Col = C_SEQ_NO : .Row = 1 : .text = 1
		.Col = C_W_TYPE : .Row = 1 : .text = TYPE_6
	End With

	Call GetFISC_DATE

End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	If pType = TYPE_6 Then	' 자기자본적수계산 
		ggoSpread.SpreadLock C_W13, -1, C_W13
		ggoSpread.SpreadLock C_W14, -1, C_W14
	Else
		ggoSpread.SpreadLock C_W5, -1, C_W5
		ggoSpread.SpreadLock C_W6, -1, C_W6
		ggoSpread.SpreadLock C_W7, -1, C_W7
		ggoSpread.SSSetRequired	 C_W1, -1, C_W1
		'ggoSpread.SSSetProtected C_W1, lgvspdData(pType).MaxRows, C_W1, lgvspdData(pType).MaxRows 
	End If
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(pType)

	If pType = TYPE_6 Then	' 자기자본적수계산 
		ggoSpread.SSSetProtected C_W13, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W14, pvStartRow, pvEndRow 
	Else
		ggoSpread.SSSetProtected C_W5, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W6, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W7, pvStartRow, pvEndRow 
		If lgvspdData(pType).MaxRows = pvEndRow Then
			ggoSpread.SSSetRequired	 C_W1, pvStartRow, pvEndRow -1
		Else
			ggoSpread.SSSetRequired	 C_W1, pvStartRow, pvEndRow
		End If
	End If

End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_1 To TYPE_5
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
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
            C_W21		= iCurColumnPos(3)
            C_W1		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W21		= iCurColumnPos(6)
            C_W3		= iCurColumnPos(7)
            C_W4		= iCurColumnPos(8)
            C_W5		= iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W17		= iCurColumnPos(13)
            C_W_TYPE	= iCurColumnPos(14)
            C_W1		= iCurColumnPos(15)
            C_W2		= iCurColumnPos(16)
    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================

Sub GetFISC_DATE()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd
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
	
	' 6. 자기자본 적수계산 그리드 W13 입력 
	With lgvspdData(TYPE_6)
		.Col = C_W13 : .Row = .MaxRows 
		If frm1.cboREP_TYPE.value = "2" Then
			.text = DateDiff("d", lgFISC_START_DT, DateAdd("m", 6, lgFISC_START_DT) - 1)+1
		Else
			.text = lgFISC_END_DT - lgFISC_START_DT + 1
		End If
	End With
End Sub

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
			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
	' 2. 대차대조표의 자산총계, 부채총계-미지급법인세, 자본금+미지급법인세+주식발행초과금+감자차익-주식발행할인차금-감자차손 가져오기 
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
    
'    With frm1
 '       If .vspdData.ActiveRow > 0 then 
  '          arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
   '         arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
    '    End If            
    'End With    

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

'====================================== 탭 함수 =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1	' 기본 그리드 
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_3
End Function

Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3)
	gSelframeFlg = TAB3
	lgCurrGrid = TYPE_5
End Function


'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110110100101111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData 
	Call fncQuery
     
    
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

' -- 3번 그리드 
Sub vspdData3_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_4A
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4A
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_4A
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_4A
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_GotFocus()
	lgCurrGrid = TYPE_4A
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_4A
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_4A
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_4A
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 4번 그리드 
Sub vspdData4_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_4B
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4B
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_4B
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData4_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_4B
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_GotFocus()
	lgCurrGrid = TYPE_4B
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData4_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_4B
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData4_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_4B
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData4_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_4B
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 5번 그리드 
Sub vspdData5_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_5
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData5_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_5
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData5_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_5
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData5_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_5
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData5_GotFocus()
	lgCurrGrid = TYPE_5
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData5_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_5
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData5_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_5
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData5_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_5
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 6번 그리드 
Sub vspdData6_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_6
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData6_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_6
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData6_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_6
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData6_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_6
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData6_GotFocus()
	lgCurrGrid = TYPE_6
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData6_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_6
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData6_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_6
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData6_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_6
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, datW1_DOWN, datW1, iRow, iMaxRows, dblW5_UP, dblW10, dblW11, dblW3, dblW4, dblW5
	Dim preSum
	
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
	
	If Index = TYPE_6 Then	' 자기자본적수계산 
		Select Case Col
			Case C_W10, C_W11
				.Col = C_W10 : dblW10 = UNICDbl(.text)
				.Col = C_W11 : dblW11 = UNICDbl(.text)
				
				If (dblW10 - dblW11) > lgW12_REF Then
					.Col = C_W12 : .text = dblW10 - dblW11
				Else
					.Col = C_W12 : .text = lgW12_REF
				End If
				
				Call SetW14()
			Case C_W12	' -- 2006.03.22 이벤트 추가
				Call SetW14()
		End Select 
		
	ElseIf Index=TYPE_4B Then	'대변차변이 다른 것과 반대
		Select Case Col
			Case C_W1	' 연월일 변경시 
				If lgFISC_END_DT = "" Then
					MsgBox "법인 기본정보의 사업연도 종료일이 비어있습니다"
					Exit Sub
				Else
					iMaxRows = .MaxRows
				
					' 1-1. 현재 입력한 연월일을 기준으로 다음행보다 크면 에러를 일으킨다.
					If Row + 1 <> iMaxRows Then
						.Row = Row		: .Col = C_W1	: datW1 = CDate(.Text)
						
						' 1.1 아래행이 있을 경우 
						.Row = Row+1	: .Col = C_W1	
						If .Text <> "" Then
							datW1_DOWN = CDate(.Text)

							If datW1 > datW1_DOWN Then ' 아래행보다 날짜가 이후면 에러 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '⊙: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
					'1-2.현재입력한 연월일을 기준으로 이전행보다 작으면 에러
					ElseIf Row-1 <> 0 Then 
						.Row=Row		:.Col = C_W1 : datW1=Cdate(.Text)
						.Row=Row-1	:.Col = C_W1
						If .text<>"" Then 
							datW1_DOWN =Cdate(.text)
							If datW1 < datW1_DOWN Then ' 아래행보다 날짜가 이전면 에러 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '⊙: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
						
					End If
					
					.Col = C_W3	: dblW3 = UNICDbl(.text)
					.Col = C_W4	: dblW4 = UNICDbl(.text)
					
					If dblW3 > 0 Or dblW4 > 0 Then
					
						' 2. 일수 변경 요청 
						Call SetW6(Index, Row)					
		
						' 2. 적수 변경 
						Call SetW7(Index, Row)	
					End If
				End If
	
			Case C_W3, C_W4		' 차/대변 

				' 1. 기준 연월일부터 입력체크 
				.Col = C_W1		: .Row = Row	
				If .Text = "" Then	
					Call DisplayMsgBox("W30002", parent.VB_INFORMATION, "X", "X")           '⊙: "일자를 먼저 입력하십시오.."
					.Col = Col	: .text =""
					Exit Sub				
				End If
							
				' 2. 음수 체크 
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.text)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '⊙: "%1 금액이 0보다 적습니다."
					.text = 0
				End If

				' 3. 컬럼 합계 계산 
				'dblSum = FncSumSheet(lgvspdData(Index), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' 합계 

				' 4. 상위행 데이타 체크 
				If Row > 1 Then
					.Row = Row -1	: .Col = C_W5	
					If .text = "" Then
						Call DisplayMsgBox("W30003", parent.VB_INFORMATION, "X", "X")           '⊙: "상위행 차/대변을 먼저 입력하십시오."
						Exit Sub
					End If
				End If
					
				' 5. 잔액 계산 
				
				.Row = Row
				.Col = C_W3	: dblW3 = UNICDbl(.text)
				.Col = C_W4	: dblW4 = UNICDbl(.text)
				
				dblW5 =   dblW4-dblW3	' 잔액 
				
				' 4.1 첫행인지 체크 
				If Row - 1 = 0 Then
					.Col = C_W5 : .text = dblW5
				Else
					' 첫행이 아닐때 
					.Row = Row -1
					.Col = C_W5	: dblW5_UP = UNICDbl(.text)	' 윗행 잔액 
					.Row = Row
					.Col = C_W5	: .text = dblW5_UP + dblW5		' 윗행 잔액+현재행 잔액 
				End If
				
				' 자신이 변경되면 다른 호출해야 할 부분을 아래에 기술 
				Call SetW5(Index, Row)
				
				' 일수 변경 요청 
				Call SetW6(Index, Row)
				
				Call SetW7(Index, Row)	' 적수 계산 
		End Select
	
	Else
		Select Case Col
			Case C_W1	' 연월일 변경시 
				If lgFISC_END_DT = "" Then
					MsgBox "법인 기본정보의 사업연도 종료일이 비어있습니다"
					Exit Sub
				Else
					iMaxRows = .MaxRows
					
					' 1-1. 현재 입력한 연월일을 기준으로 다음행보다 크면 에러를 일으킨다.
					If Row + 1 <> iMaxRows Then
						.Row = Row		: .Col = C_W1	: datW1 = CDate(.Text)
						
						' 1.1 아래행이 있을 경우 
						.Row = Row+1	: .Col = C_W1	
						If .Text <> "" Then
							datW1_DOWN = CDate(.Text)

							If datW1 > datW1_DOWN Then ' 아래행보다 날짜가 이후면 에러 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '⊙: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
					'1-2.현재입력한 연월일을 기준으로 이전행보다 작으면 에러
					ElseIf Row-1 <> 0 Then 
						.Row=Row		:.Col = C_W1 : datW1=Cdate(.Text)
						.Row=Row-1	:.Col = C_W1
						If .text<>"" Then 
							datW1_DOWN =Cdate(.text)
							If datW1 < datW1_DOWN Then ' 아래행보다 날짜가 이전면 에러 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '⊙: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
						
					End If
					
					.Col = C_W3	: dblW3 = UNICDbl(.text)
					.Col = C_W4	: dblW4 = UNICDbl(.text)
					
					If dblW3 > 0 Or dblW4 > 0 Then
					
						' 2. 일수 변경 요청 
						Call SetW6(Index, Row)					
		
						' 2. 적수 변경 
						Call SetW7(Index, Row)	
					End If
				End If
	
			Case C_W3, C_W4		' 차/대변 

				' 1. 기준 연월일부터 입력체크 
				.Col = C_W1		: .Row = Row	
				If .Text = "" Then	
					Call DisplayMsgBox("W30002", parent.VB_INFORMATION, "X", "X")           '⊙: "일자를 먼저 입력하십시오.."
					.Col = Col	: .text = 0
					'Exit Sub				
				End If
							
				' 2. 음수 체크 
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.text)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '⊙: "%1 금액이 0보다 적습니다."
					.text = 0
				End If

				' 3. 컬럼 합계 계산 
				'dblSum = FncSumSheet(lgvspdData(Index), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' 합계 

				' 4. 상위행 데이타 체크 
				If Row > 1 Then
					.Row = Row -1	: .Col = C_W5	
					If .text = "" Then
						Call DisplayMsgBox("W30003", parent.VB_INFORMATION, "X", "X")           '⊙: "상위행 차/대변을 먼저 입력하십시오."
						Exit Sub
					End If
				End If
					
				' 5. 잔액 계산 
				
				.Row = Row
				.Col = C_W3	: dblW3 = UNICDbl(.text)
				.Col = C_W4	: dblW4 = UNICDbl(.text)
					
				dblW5 =  dblW3 - dblW4	' 잔액 
				
				' 4.1 첫행인지 체크 
				If Row - 1 = 0 Then
					.Col = C_W5 : .text = dblW5
				Else
					' 첫행이 아닐때 
					.Row = Row -1
					.Col = C_W5	: dblW5_UP = UNICDbl(.text)	' 윗행 잔액 
					.Row = Row
					.Col = C_W5	: .Value = dblW5_UP + dblW5		' 윗행 잔액+현재행 잔액 
				End If
				
				' 자신이 변경되면 다른 호출해야 할 부분을 아래에 기술 
				Call SetW5(Index, Row)
				
				' 일수 변경 요청 
				Call SetW6(Index, Row)
				
				Call SetW7(Index, Row)	' 적수 계산 
		End Select
	
	End If
	
	End With
	
End Sub

' 일수 변경 
Sub SetW6(Index, Row)	
	Dim dblW5, dblW6, datW1, datW1_DOWN, dblSum, iRow, blnPrintLast
	
	With lgvspdData(Index)
		blnPrintLast = False
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		For iRow = .MaxRows-1 To 1 Step -1
			.Row = iRow
			.Col = C_W5	: dblW5 = UNICDbl(.text)	' 잔액값 읽기 
			.Col = C_W1	
		
			If .Text = "" Then		' 잔액이 엄거나, 연월일이 공란이면 합계후 종료한다.
				
			Else		

				datW1 = CDate(.Text)
		
				If blnPrintLast = False Then	' 마지막행 일수 계산안한경우 
					If frm1.cboREP_TYPE.value = "2" Then
						.Col = C_W6	: .text = DateDiff("d", datW1, DateAdd("m", 6, lgFISC_START_DT)-1)+1
					Else
						.Col = C_W6	: .text = DateDiff("d", datW1, lgFISC_END_DT)+1
					End If
					blnPrintLast = True
				Else
					.Col = C_W1	: .Row = iRow+1	
					
					If .Text <> "" Then	' 존재할때.
						datW1_DOWN = CDate(.Text)	' 현재 변경행의 일자를 기억 
						.Col = C_W6	: .Row = iRow	: .text = DateDiff("d", datW1,  datW1_DOWN)	
					End If
				End If
			
			ggoSpread.UpdateRow iRow	
			End If
		Next
		
		'dblSum = FncSumSheet(lgvspdData(Index), C_W6, 1, .MaxRows - 1, true, .MaxRows, C_W6, "V")	' 합계 

		'Call UpdateTotalLIne
		
	End With	

End Sub

Sub UpdateTotalLIne()
	ggoSpread.Source = lgvspdData(lgCurrGrid)
	ggoSpread.UpdateRow lgvspdData(lgCurrGrid).MaxRows
End Sub

' 잔액 컬럼이 변경될때 호출됨 
Sub SetW5(Index, Row)
	Dim dblSum, dblW3, dblW4, dblW5, dblW6, dblW7, iRow, iMaxRows, dblW5_UP, datW1
	
	With lgvspdData(Index)

		iMaxRows = .MaxRows
		.Row = Row	' 호출된 시점의 Row
		.Col = C_W5	: dblW5 = UNICDbl(.text)	' 현재 변경행의 잔액을 기억 
		.Col = C_W6 : dblW6 = UNICDbl(.text)
		.Col = C_W7 : dblW7 = dblW5 * dblW6	: .text = dblW7	' 적수 변경 
		
		.Row = Row + 1	' 하위 행 
		.Col = C_W3	: dblW3 = UNICDbl(.text)	' 차변 
		.Col = C_W4	: dblW4 = UNICDbl(.text)	' 대변 
						
		If Row = iMaxRows -1 then 'Or (dblW3 = 0 And dblW4 = 0) Then	' 합계 이전 행이거나, 다음행에 차/대변이 공란이면 
'			' 완료 하였으므로 합계를 구한다.
'			dblSum = FncSumSheet(lgvspdData(Index), C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 
'			Call UpdateTotalLIne
			Exit Sub
		End If

		.Col = C_W5	: dblW5 = dblW5 + (dblW3 - dblW4) 
		.text = dblW5

		'.Col = C_W1 : datW1 = .Text	' 연월일 
	End With
	
	' 자신이 변경되면 다른 호출해야 할 부분을 아래에 기술 
	Call SetW5(Index, Row + 1)	' 하위행을 변경하엿으므로, 재귀 루프를 시작한다. 즉 1-10행(합계11행)이 있을 경우, 5행을 고치면 6-10행까지 고쳐야 한다.	

End Sub

Sub SetW7(Index, Row)	' 적수가 변경될시 해야될 이벤트 

	' 자신이 변경되면 다른 호출해야 할 부분을 아래에 기술 
	Dim dblW5, dblW6, iRow, dblSum
	
	With lgvspdData(Index)
		For iRow = 1 To .MaxRows -1 
			.Row = iRow
			.Col = C_W5 : dblW5 = UNICDbl(.text)
			.Col = C_W6 : dblW6 = UNICDbl(.text)
			If dblW5 <> 0 And dblW6 <> 0 Then 
				.Col = C_W7 : .text = dblW5 * dblW6
			End If
		Next
		
		dblSum = FncSumSheet(lgvspdData(Index), C_W7, 1, .MaxRows - 1, true, .MaxRows, C_W7, "V")	' 합계 
		Call UpdateTotalLIne
	End With
End Sub


Sub SetW14() ' 적수계산 
	Dim dblW12, dblW13


	With lgvspdData(TYPE_6)
	
		.Col = C_W12 : 	dblW12 = UNICDbl(.text)
		.Col = C_W13 : dblW13 = UNICDbl(.text)
		.Col = C_W14 : .text = dblW12 * dblW13
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
       
'       If lgSortKey = 1 Then
 '          ggoSpread.SSSort Col               'Sort in ascending
  '         lgSortKey = 2
   '    Else
    '       ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
     '      lgSortKey = 1
      ' End If
       
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
    For i = TYPE_1 To TYPE_6
    
		ggoSpread.Source = lgvspdData(i)
		IF ggoSpread.SSDefaultCheck = False Then								  '☜: Check contents area
			Exit Function
		End If
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
		End If
		
	Next
	
	If blnChange = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

    For i = TYPE_1 To TYPE_5
		' 검증 '1 각 잔액의 합은 < 0 작을수 없다. WC0006
		With lgvspdData(i)
		If .MaxRows > 0 Then
			.Row = .MaxRows : .Col = C_W5
			If UNICDbl(.text) < 0 Then
				Select Case i
					Case TYPE_1
						sMsg = "1. 업무무관부당산의 적수"
					Case TYPE_2
						sMsg = "2. 업무무관동산의 적수"
					Case TYPE_3
						sMsg = "3. 타법인주식의 적수"
					Case TYPE_4A
						sMsg = "4. 가지급금의 적수" & vbCrLf & " ㄱ. 가지급금의 적수"
					Case TYPE_4B
						sMsg = "4. 가지급금의 적수" & vbCrLf & " ㄴ. 가수금의 적수"
					Case TYPE_5
						sMsg = "5. 기타 적수"
					Case TYPE_6
						sMsg = "6. 자기자본 적수계산"
				End Select
				sMsg = sMsg & "의 (5)잔액 합계 "
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, sMsg, "X")                          
				Exit Function
			End If
		End If
		End With
	Next
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True         
    Set gActiveElement = document.ActiveElement                                                 
    
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

    Call SetToolbar("1110110100001111")

	Call ClickTab1()
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

			lgvspdData(lgCurrGrid).Col = C_W21
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
	
	If lgCurrGrid = TYPE_6 Then Exit Function
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
	
'		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' 합계 
'		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 합계 
	
		' 자신이 변경되면 다른 호출해야 할 부분을 아래에 기술 
		Call SetW5(lgCurrGrid, 1)
					
		' 일수 변경 요청 
		Call SetW6(lgCurrGrid, 1)
					
		Call SetW7(lgCurrGrid, 1)	' 적수 계산 
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
   	 
	If lgCurrGrid = TYPE_6 Then	Exit Function	' 6번 그리드는 추가할수 없다.
	
	
	With lgvspdData(lgCurrGrid)	' 포커스된 그리드 
		
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		
		iRow = .ActiveRow
		lgvspdData(lgCurrGrid).ReDraw = False
				
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow ,2
			Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
			.Row = iRow		
			.Col = C_SEQ_NO : .Text = iRow	
			.Col = C_W_TYPE : .Text = lgCurrGrid
			
			If lgCurrGrid = TYPE_6 Then
				ggoSpread.SpreadLock C_W10, iRow, C_W14, iRow
				
			Else	
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
				.Col = C_W_TYPE : .Text = lgCurrGrid	
				.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
						
				ggoSpread.SpreadLock C_W1, iRow, C_W7, iRow
			End If
						
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
	
	Call CheckW7Status(lgCurrGrid)	' 적수셀 상태 체크 

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' -- TYPE_4A, TYPE_4B 그리드의 적수 Enable/Disable 체크 
Function CheckW7Status(Index)

	If Index = TYPE_4A Or Index = TYPE_4B Then
	
		With lgvspdData(Index)
	
		ggoSpread.Source = lgvspdData(Index)

		If lgvspdData(Index).MaxRows > 1 Then
			ggoSpread.SpreadLock C_W7, .MaxRows, C_W7, .MaxRows
		Else
			ggoSpread.SpreadUnLock C_W7, .MaxRows, C_W7, .MaxRows
		End If
	
		End With
	End If
End Function

' GetREF 에서 적수 가져온뒤 호출됨 
Function InsertTotalLine(Index)
	With lgvspdData(Index)
	
	ggoSpread.Source = lgvspdData(Index)
	
	If .MaxRows = 0 Then	' 한줄 추가 
		ggoSpread.InsertRow ,1
		
		.Row = 1
		.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
		.Col = C_W_TYPE : .Text = Index	
		.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
		
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
		.Col = C_W_TYPE : .Value = lgCurrGrid	' 현재 그리드 번호 
		MaxSpreadVal lgvspdData(lgCurrGrid), C_SEQ_NO, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(lgCurrGrid), C_SEQ_NO, iRow)	' 현재의 최대SeqNo를 구한다 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO : .text = iSeqNo : iSeqNo = iSeqNo + 1
			.Col = C_W_TYPE : .text = lgCurrGrid	' 현재 그리드 번호 
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

	If lgCurrGrid = TYPE_6 Then Exit Function
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
    If ggoSpread.SSCheckChange = True Then
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
        strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If lgvspdData(TYPE_1).MaxRows > 0 Or _
		lgvspdData(TYPE_2).MaxRows > 0 Or _
		lgvspdData(TYPE_3).MaxRows > 0 Or _
		lgvspdData(TYPE_4A).MaxRows > 0 Or _
		lgvspdData(TYPE_4B).MaxRows > 0 Or _
		lgvspdData(TYPE_5).MaxRows > 0 Or _
		lgvspdData(TYPE_6).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
			'Call SetSpreadLock(TYPE_1)
			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1111111100011111")										<%'버튼 툴바 제어 %>
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W7, -1
			Call SetToolbar("1110000000011111")										<%'버튼 툴바 제어 %>
		End If
	
	Else
		Call SetToolbar("1100110100111111")										<%'버튼 툴바 제어 %>
	End If
	
	Call SetSpreadTotalLine ' - 합계라인 재구성 
	
	Call ClickTab1()
	lgvspdData(lgCurrGrid).focus			
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
    
    For i = TYPE_1 To TYPE_6	' 전체 그리드 갯수 
    
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
		                                              strVal = strVal & "D"  &  Parent.gColSep
		       End Select
		       
			  ' 모든 그리드 데이타 보냄     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = C_SEQ_NO To lMaxCols
						If lCol=C_W1  Then 
							.Col = lCol : 
							If  trim(.Text)="계" or trim(.text)="Blank" Then 
								strVal = strVal & Trim("2999-12-31") &  Parent.gColSep
							Else
								strVal = strVal & Trim(.Text) &  Parent.gColSep
							End IF
						Else
							
						.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
						End If
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
			Next
		
		End With

	Next


	Frm1.txtSpread.value      = strVal '& strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	
	For iRow = TYPE_1 To TYPE_6
		lgvspdData(lgCurrGrid).MaxRows = 0
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		ggoSpread.ClearSpreadData
	Next
	
	lgvspdData(lgCurrGrid).MaxRows = 1
	
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()" width=200>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>업무무관 부동산/동산의 적수</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>타법인 주식/가지급금의 적수</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기타/자기자본 적수계산</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w3127ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;1. 업무무관 부동산의 적수							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData0_vspdData0.js'></script>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;2. 업무무관 동산의 적수							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData1_vspdData1.js'></script>
										</TD>
									</TR>
								</TABLE>
								</DIV>
						
								<DIV ID="TabDiv" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;3. 타법인주식의 적수							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="30%">
											<script language =javascript src='./js/w3127ma1_vspdData2_vspdData2.js'></script>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;4. 가지급금의 적수							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="60%">
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD HEIGHT="10">&nbsp;ㄱ. 가지급금등의 적수							
												</TD>
											</TR>
											<TR>
												<TD HEIGHT="45%">
													<script language =javascript src='./js/w3127ma1_vspdData3_vspdData3.js'></script>
												</TD>
											</TR>
											<TR>
												<TD HEIGHT="10">&nbsp;ㄴ. 가수금등의 적수							
												</TD>
											</TR>
											<TR>
												<TD HEIGHT="45%">
													<script language =javascript src='./js/w3127ma1_vspdData4_vspdData4.js'></script>
												</TD>
											</TR>								
											</TABLE>
										</TD>
									</TR>
								</TABLE>
								</DIV>

								<DIV ID="TabDiv" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;5. 기타적수							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData5_vspdData5.js'></script>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;6. 자기자본 적수계산							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData6_vspdData6.js'></script>
										</TD>
									</TR>
								</TABLE>
								</DIV>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>업무무관의부동산의 적수</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><별지>업무무관의동산의 적수</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check3" ><LABEL FOR="prt_check3"><별지>타법인주식의 적수</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check4" ><LABEL FOR="prt_check4"><별지>가지급금등의 적수</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check5" ><LABEL FOR="prt_check5"><별지>가수금등의 적수</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check6" ><LABEL FOR="prt_check6"><별지>기타 적수</LABEL>&nbsp;
				        
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

