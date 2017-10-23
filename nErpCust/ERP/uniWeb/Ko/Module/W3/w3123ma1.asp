
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 각 과목별 조정 
'*  3. Program ID           : W3127MA1
'*  4. Program Name         : W3127MA1.asp
'*  5. Program Desc         : 제26호 업무무관부동산등에 관련한 차입금이자조정명세서(을)
'*  6. Modified date(First) : 2005/01/05
'*  7. Modified date(Last)  : 2005/01/05
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W3123MA1"
Const BIZ_PGM_ID		= "w3123mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "w3123mb2.asp"											 '☆: 비지니스 로직 ASP명 

Const TAB1 = 1																	'☜: Tab의 위치 
Const TAB2 = 2

Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.

' -- 그리드 컬럼 정의 
Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(6)	' 멀티 그리드 처리 변수 
Dim lgblnYoon

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1	' -- 1번 그리드 
    C_W1		= 2	' 일자/이자율 
    C_W2		= 3 ' 금액/지급이자 
    C_W3		= 4	' 적요/적수 
    C_W4		= 5	' 이자율 
    C_W5		= 6	' 적수 
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
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
   
End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	' 그리드 셋팅 
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	
	lgvspdData(TYPE_1).ScriptEnhanced  = True
	lgvspdData(TYPE_2).ScriptEnhanced  = True
	
    Call initSpreadPosVariables()  

	' 1번 그리드(탭1)
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W5 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
				       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		Call AppendNumberPlace("6","2","2")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"순번"		, 10,,,6,1	' 히든컬럼 
		ggoSpread.SSSetDate		C_W1,		"(1)일 자"	, 10, 2, Parent.gDateFormat, -1
		ggoSpread.SSSetFloat	C_W2,		"(2)금 액"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetEdit		C_W3,		"(3)적 요"	, 20,,,50,1
		ggoSpread.SSSetFloat	C_W4,		"(4)이자율"	, 10, 6,	ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W5,		"(5)적 수"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		' 퍼센트 형 정의 
		.Col = C_W4
		.Row = -1
		.CellType = 14
		'.TypePercentDecimal = 2
		.TypePercentMax = 99
		.TypePercentMin = 0
		'.TypePercentDecPlaces = 0
    
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		Call SetSpreadLock(TYPE_1)
				
		.ReDraw = true	
			
	End With 
 
	' 2번 그리드(탭1)
	With lgvspdData(TYPE_2)
			
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W3 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
				       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","3","2")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"순번"		, 10,,,6,1	' 히든컬럼 
		ggoSpread.SSSetFloat	C_W1,		"(1)이자율"	, 15, 6,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W2,		"(2)지급이자", 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W3,		"(3)적 수"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		' 퍼센트 형 정의 
		.Col = C_W1
		.Row = -1
		.CellType = 14
		'.TypePercentDecimal = 2
		.TypePercentMax = 99
		.TypePercentMin = 0
		'.TypePercentDecPlaces = 0
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		Call SetSpreadLock(TYPE_2)
				
		.ReDraw = true	
			
	End With 
     
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call CheckFISC_DATE
End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	If pType = TYPE_2 Then	' 이자율별그리드 
		ggoSpread.SpreadLock C_SEQ_NO, -1, C_W3
	Else
		ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
		ggoSpread.SSSetRequired	 C_W2, -1, C_W2
		ggoSpread.SSSetRequired	 C_W4, -1, C_W4
		ggoSpread.SpreadLock C_W5, -1, C_W5
		'ggoSpread.SSSetProtected C_W1, lgvspdData(pType).MaxRows, C_W1, lgvspdData(pType).MaxRows 
	End If
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(pType)

	If pType = TYPE_2 Then	' 이자율별그리드 
		ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W1, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W2, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W3, pvStartRow, pvEndRow 
	Else
		ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow 
		ggoSpread.SSSetRequired	 C_W2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	 C_W4, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_W5, pvStartRow, pvEndRow 
		If pvEndRow = lgvspdData(pType).MaxRows Then
			ggoSpread.SpreadLock C_W1, pvEndRow, C_W5
		End If
	End If

End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
				.Col = C_W4		: .CellType = 1	: .Text = ""
				ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With
	Next

End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData0
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W_TYPE	= iCurColumnPos(2)
            C_W1		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W3		= iCurColumnPos(7)
            C_W4		= iCurColumnPos(8)
            C_W5		= iCurColumnPos(9)
    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 세무정보 조사 : 메시지가져오기.
	Call GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call FncNew()

	Call LayerShowHide(1)

	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
	' 2. 대차대조표의 자산총계, 부채총계-미지급법인세, 자본금+미지급법인세+주식발행초과금+감자차익-주식발행할인차금-감자차손 가져오기 
End Function

' 레퍼런스에서 넣었으므로 입력으로 변환해 준다.
Function ChangeRowFlg(Index)
	Dim iRow
	
	With lgvspdData(Index) 
		ggoSpread.Source = lgvspdData(Index)
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
	End With
End Function

' -- 금액 불러오기 후 적수계산 
Function ChangeW5()
	Dim iRow
	
	With lgvspdData(TYPE_1) 
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		For iRow = 1 To .MaxRows
			Call SetW5(TYPE_1, iRow, False)
		Next
		
		Call FncSumSheet(lgvspdData(TYPE_1), C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 
	End With
End Function


Sub CheckFISC_DATE()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		sFISC_START_DT = CDate(lgF0)
	Else
		sFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		sFISC_END_DT = CDate(lgF1)
	Else
		sFISC_END_DT = ""
	End if
	
	lgblnYoon = False
	datMonCnt = DateDiff("m", sFISC_START_DT, sFISC_END_DT)
	' 현재 법인의 당기기간안에 윤달이 있는지 체크해서 lgblnYOON를 변화시킨다.
	For i = 1 To datMonCnt
		datNow = DateAdd("m", i, sFISC_START_DT)
		If Month(datNow) = 2 Then	' 2월을 가지는 당기기간이면 
			lgblnYoon = CheckIntercalaryYear(Year(datNow))
			Exit For
		End If
	Next
End Sub

'====================================== 탭 함수 =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1	' 기본 그리드 
End Function

Function ClickTab2()	
	Dim i, blnChange

	If gSelframeFlg = TAB2 Then Exit Function

	' 1번 그리드에서 온 경우 그리드 필수조건을 체크한다.
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  
		     
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_2
	
	If lgBlnFlgChgValue Then	Call GridReCalc()	' 2번 그리드 재계산 
End Function


Function GridReCalc()
	Dim iRow, iMaxRows, dblW2, dblW4, dblW5, dblOldW4, oRs
	
	dblOldW4 = 0
	
	' 2번 그리드 초기화 
	ggoSpread.Source = lgvspdData(TYPE_2)
	ggoSpread.ClearSpreadData
  	
	With lgvspdData(TYPE_1)
	
		ggoSpread.Source = lgvspdData(TYPE_1)
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows -1 
			.Row = iRow
			.Col = C_W2	: dblW2 = UNICDbl(.Value)	' 금액 
			.Col = C_W4	: dblW4 = UNICDbl(.Value)	' 이자율 
			.Col = C_W5	: dblW5 = UNICDbl(.Value)	' 적수 
			
			.Col = 0	
			If .Text <> ggoSpread.DeleteFlag Then
				Call CheckExist(dblW4, dblW2, dblW5)
			End If

		Next
	End With
	
	With lgvspdData(TYPE_2)
	
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.SSSort C_W1, 2
		ggoSpread.InsertRow,1
		ggoSpread.SpreadLock C_W1, -1, C_W3, -1
			
		.Row = .MaxRows
		.Col = C_SEQ_NO	: .Value = SUM_SEQ_NO
		.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
		Call FncSumSheet(lgvspdData(TYPE_2), C_W2, 1, .MaxRows - 1, true, .MaxRows, C_W2, "V")	' 합계 
		Call FncSumSheet(lgvspdData(TYPE_2), C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' 합계 
		
		For iRow = 1 To .MaxRows -1
			.Row = iRow
			.Col = C_SEQ_NO	: .Value = iRow
		Next
		
	End With	
End Function

Function CheckExist(pdblW4, pdblW2, pdblW5)
	Dim iRow, iMaxRows, blnExist, iNowLoc
	blnExist = False
	
	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		
		iMaxRows = .MaxRows
		For iRow = 1 to iMaxRows 
			.Col = C_W1	: .Row = iRow
			If UNICDbl(.Value) = pdblW4 Then	
				blnExist = True
				' 중복되면 해당 데이타를 업데이트 시킨다.
				.Col = C_W2 : .Value = pdblW2 + UNICDbl(.Value)
				.Col = C_W3 : .Value = pdblW5 + UNICDbl(.Value)		
				Exit For		
			End If
		Next
		
		If Not blnExist Then  ' 같은게 없다면 
			ggoSpread.InsertRow,1
			.Row = iRow + 1
			.Col = C_W1	: .Value = pdblW4
			.Col = C_W2 : .Value = pdblW2
			.Col = C_W3 : .Value = pdblW5
			
			'MaxSpreadVal lgvspdData(TYPE_2), C_SEQ_NO, iRow + 1
		End If

	End With
	
End Function
'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

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
    
    Call CheckFISC_DATE
End Sub

Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	Call CheckFISC_DATE
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
	Dim dblSum, datW1_DOWN, datW1, iRow, iMaxRows, dblW2, dblW4, dblW5
	
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
	
		Select Case Col
			Case C_W1	' 연월일 변경시 
				iMaxRows = .MaxRows
						
				' 1. 현재 입력한 연월일을 기준으로 다음행보다 크면 에러를 일으킨다.
				If Row + 1 <> iMaxRows Then
					.Row = Row		: .Col = C_W1	: datW1 = CDate(.Text)
							
					' 1.1 아래행이 있을 경우 
					.Row = Row+1	: .Col = C_W1	
					If .Text <> "" Then
						datW1_DOWN = CDate(.Text)

						If datW1 > datW1_DOWN Then ' 아래행보다 날짜가 이후면 에러 
							Call DisplayMsgBox("WC0016", parent.VB_YES, "X", "X")           '⊙: "Will you destory previous data"
							Exit Sub						
						End If
					End If

				End If
				
				Call SetW5(Index,Row, True)
			
			Case C_W2, C_W4		' 금액 

				' 1. 음수 체크 
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '⊙: "%1 금액이 0보다 적습니다."
					.Value = 0
				End If

				' 2. 컬럼 합계 계산 
				If Col = C_W2 Then	' 2006-05-12 수정(컬럼하드코딩)
					dblSum = FncSumSheet(lgvspdData(Index), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' 합계 
				End If
				
				.Col = C_W2	: .Row = Row	: dblW2 = UNICDbl(.Value)
				.Col = C_W4	: .Row = Row	: dblW4 = UNICDbl(.Value)
					
				Call SetW5(Index, Row, True)
					
		End Select
	

	End With
	
End Sub

' -- 적수 계산 
Function SetW5(Index, Row, blnSum)
	Dim datW1, dblW2, dblW4, dblW5

	With lgvspdData(Index)
	
		.Col = C_W2	: .Row = Row	: dblW2 = UNICDbl(.Value)
		.Col = C_W4	: .Row = Row	: dblW4 = UNICDbl(.Value)
						
		If dblW4 <> 0 And dblW2 <> 0 Then
			.Col = C_W1	: .Row = Row	

			' 3. 적수계산 
			If lgblnYoon Then
				' 윤년 
				dblW5 = (dblW2 / (dblW4)) * 366
			Else	
				' 평년 
				dblW5 = (dblW2 / (dblW4)) * 365
			End If
							
			.Col = C_W5	: .Row = Row	: .Value = dblW5	'UNIFormatNumber(dblW5, ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
			
			' 2. 컬럼 합계 계산 
			If blnSum = True Then
				Call FncSumSheet(lgvspdData(Index), C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 
			End If
		End If
	End With
End Function


Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    Call SetPopupMenuItemInf("1101000000") 

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
    ggoSpread.Source = frm1.vspdData0
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
    Call ggoOper.LockField(Document, "N")
   
    ggoSpread.ClearSpreadData
    Call InitVariables													<%'Initializes local global variables%>
    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgvspdData(TYPE_2).MaxRows = 0 Or lgBlnFlgChgValue Then Call GridReCalc
    
    For i = TYPE_1 To TYPE_2
    
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
		End If
	Next
	
	If blnChange = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	 
	' 검증작업 
	'1. 차입금 적수계산 : (2)금액의 합계 < 0 이면 오류(WC0013)
	With lgvspdData(TYPE_1)
		.Row = .MaxRows : .Col = C_W2
		If .Value < 0 Then
			Call DisplayMsgBox("WC0013", "X", "X", "X")                          
			Exit Function
		End If 
	End With		

	'2. 이자율별 적수계산 : (2)금액의 합계 < 0 이면 오류(WC0013)
	With lgvspdData(TYPE_2)
		.Row = .MaxRows : .Col = C_W2
		If .Value < 0 Then
			Call DisplayMsgBox("WC0013", "X", "X", "X")                          
			Exit Function
		End If 
	End With	
		
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
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

    Call SetToolbar("1100110100000111")

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
	
	If lgCurrGrid = TYPE_2 Then Exit Function
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    
    lgBlnFlgChgValue = True
    Call CheckReCalc()				' 한라인이 취소되면 재계산 
    'Call CheckW7Status(lgCurrGrid)	' 적수셀 상태 체크 
End Function

' 재계산 
Function CheckReCalc()
	Dim dblSum
	
	With lgvspdData(lgCurrGrid)
		If .MaxRows = 0 Then Exit Function
		ggoSpread.Source = lgvspdData(lgCurrGrid)	
	
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W2, 1, .MaxRows - 1, true, .MaxRows, C_W2, "V")	' 합계 
		'dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 합계 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 
	
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
   
	If lgCurrGrid = TYPE_2 Then	Exit Function	' 2번 그리드는 추가할수 없다.
	
	With lgvspdData(lgCurrGrid)	' 포커스된 그리드 
		
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		
		iRow = .ActiveRow
		lgvspdData(lgCurrGrid).ReDraw = False
				
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow , 2
			
			Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
			.Row = iRow		
			.Col = C_SEQ_NO : .Text = iRow	
	
			iRow = 2		: .Row = iRow
			.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
			.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
						
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
	
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' GetREF 에서 적수 가져온뒤 호출됨 
Function InsertTotalLine(Index)
	With lgvspdData(Index)
	
	ggoSpread.Source = lgvspdData(Index)
	
	If .MaxRows > 0 Then	' 한줄 추가 
		ggoSpread.InsertRow .MaxRows,1
		SetSpreadColor Index, .MaxRows, .MaxRows
		
		.Row = .MaxRows
		.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
		.Col = C_W1		: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
		
		ggoSpread.SpreadLock C_W1, .MaxRows, C_W5, .MaxRows
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

	If lgCurrGrid = TYPE_2 Then Exit Function
	With lgvspdData(lgCurrGrid)
		.focus
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		lDelRows = ggoSpread.DeleteRow
	End With
	
	lgBlnFlgChgValue = True
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
	ggoSpread.Source = lgvspdData(TYPE_1)
	
	If lgvspdData(TYPE_1).MaxRows > 0 Or _
		lgvspdData(TYPE_2).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE

		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
			Call SetSpreadLock(TYPE_1)
			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101111100000111")										<%'버튼 툴바 제어 %>
		Else
		
			ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
		End If
	
		Call SetSpreadTotalLine ' - 합계라인 재구성 

		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.SpreadLock	C_W1, -1, C_W3, -1
				    
	Else
		Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>
	End If
	

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
    
    For i = TYPE_1 To TYPE_2	' 전체 그리드 갯수 
    
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

		If i = TYPE_1 Then
			Frm1.txtSpread.value      = strDel & strVal
			strVal = "" :  strDel = ""
		Else
			Frm1.txtSpread2.value      = strDel & strVal
		End If
	Next

	
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	
	For iRow = TYPE_1 To TYPE_2
		lgvspdData(lgCurrGrid).MaxRows = 0
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		ggoSpread.ClearSpreadData
	Next
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>차입금 적수 계산</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>이자율별 차입금적수계산</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">금액불러오기</A></TD>
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
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X1"></SELECT>
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="25X1" TITLE="SPREAD" id=vaSpread Index=0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</DIV>
						
								<DIV ID="TabDiv" SCROLL=no>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="25X1" TITLE="SPREAD" id=vaSpread Index=1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

