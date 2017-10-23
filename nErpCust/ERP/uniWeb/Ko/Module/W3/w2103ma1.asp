
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 수입금액조정 
'*  3. Program ID           : W1113MA1
'*  4. Program Name         : W1113MA1.asp
'*  5. Program Desc         : 수입배당금 입력 
'*  6. Modified date(First) : 2004/12/28
'*  7. Modified date(Last)  : 2004/12/28
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

Const BIZ_MNU_ID = "w2103ma1"
Const BIZ_PGM_ID = "w2103mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID = "w2103mb2.asp"

Const TAB1 = 1																	'☜: Tab의 위치 
Const TAB2 = 2

Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.

Dim C_SEQ_NO
Dim C_DOC_DATE
Dim C_DOC_AMT
Dim C_DEBIT_CREDIT
Dim C_DEBIT_CREDIT_NM
Dim C_SUMMARY_DESC
Dim C_COMPANY_NM
Dim C_STOCK_RATE
Dim C_ACQUIRE_AMT
Dim C_COMPANY_TYPE
Dim C_COMPANY_TYPE_NM
Dim C_HOLDING_TERM
Dim C_JUKSU
Dim C_OWN_RGST_NO
Dim C_CO_ADDR
Dim C_REPRE_NM
Dim C_STOCK_CNT

Dim C_MINOR_NM
Dim C_MINOR_CD
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W3_1
Dim C_W3_2
Dim C_W3_3
Dim C_W3_4
Dim C_W3_5
Dim C_W4

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO			= 1
    C_DOC_DATE			= 2
    C_DOC_AMT			= 3
    C_DEBIT_CREDIT		= 4
    C_DEBIT_CREDIT_NM	= 5
    C_SUMMARY_DESC		= 6
    C_COMPANY_NM		= 7
    C_STOCK_RATE		= 8
    C_ACQUIRE_AMT		= 9
    C_COMPANY_TYPE		= 10
    C_COMPANY_TYPE_NM	= 11
    C_HOLDING_TERM		= 12
    C_JUKSU				= 13
    C_OWN_RGST_NO		= 14
    C_CO_ADDR			= 15
    C_REPRE_NM			= 16
    C_STOCK_CNT			= 17
    
    C_MINOR_NM			= 2
    C_MINOR_CD			= 3
    C_W1				= 4
    C_W2				= 5
    C_W3				= 6
    C_W3_1				= 6
    C_W3_2				= 7
    C_W3_3				= 8
    C_W3_4				= 9
    C_W3_5				= 10
    C_W4				= 11
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

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20041222",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_STOCK_CNT + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    Call AppendNumberPlace("6","3","1")

    ggoSpread.SSSetEdit		C_SEQ_NO,		"순번", 10,,,100,1
	ggoSpread.SSSetDate		C_DOC_DATE,		"(1)일자",			10,		2,		Parent.gDateFormat,	-1
	ggoSpread.SSSetFloat	C_DOC_AMT,		"(2)금액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec 
    ggoSpread.SSSetCombo	C_DEBIT_CREDIT, "차/대"    , 10
    ggoSpread.SSSetCombo	C_DEBIT_CREDIT_NM, "차/대"    , 10
    ggoSpread.SSSetEdit		C_SUMMARY_DESC, "(3)적요", 15,,,200,1
    ggoSpread.SSSetEdit		C_COMPANY_NM,	"(4)회사명", 15,,,30
    ggoSpread.SSSetFloat	C_STOCK_RATE,	"(5)지분율" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetFloat	C_ACQUIRE_AMT,	"(6)취득가액" , 15,Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
    ggoSpread.SSSetCombo	C_COMPANY_TYPE, "회사구분", 10
    ggoSpread.SSSetCombo	C_COMPANY_TYPE_NM, "(7)회사구분", 10
    ggoSpread.SSSetFloat	C_HOLDING_TERM, "(8)당기보유기간", 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
    ggoSpread.SSSetFloat	C_JUKSU,	"적수" , 15,Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
    ggoSpread.SSSetEdit		C_OWN_RGST_NO,	"(9)사업자등록번호", 14, 2,,100,1
    ggoSpread.SSSetEdit		C_CO_ADDR,		"(10)소재지", 20,,,100,1
    ggoSpread.SSSetEdit		C_REPRE_NM,		"(11)대표자", 10,,,100,1
    ggoSpread.SSSetFloat	C_STOCK_CNT,	"(12)발행주식총수", 14, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"

	.Col = C_OWN_RGST_NO : .Row = -1 : .CellType = 4 : .TypePicMask = "999-99-99999" 
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_DEBIT_CREDIT_NM,C_DEBIT_CREDIT_NM,True)
	Call ggoSpread.SSSetColHidden(C_DEBIT_CREDIT,C_DEBIT_CREDIT,True)
	Call ggoSpread.SSSetColHidden(C_COMPANY_TYPE,C_COMPANY_TYPE,True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
	Call InitSpreadComboBox()
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With

	' 2번 그리드 
	With frm1.vspdData2
	
	ggoSpread.Source = frm1.vspdData2	
   'patch version
    ggoSpread.Spreadinit "V20041222_2",,parent.gAllowDragDropSpread    
    
	.ReDraw = false
	
	'헤더를 2줄로    
    .ColHeaderRows = 2   
    
    .MaxCols = C_W4 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    Call AppendNumberPlace("7","2","0")

    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번", 10,,,15,1
    ggoSpread.SSSetEdit		C_MINOR_NM,	"코드명", 10,,,100,1
    ggoSpread.SSSetEdit		C_MINOR_CD,	"코드", 10,,,10,1
	ggoSpread.SSSetEdit		C_W1,		"(1)법인구분", 10,,,100,1
	ggoSpread.SSSetEdit		C_W2,		"(2)배당법인구분", 15,,,100,1  
    ggoSpread.SSSetCombo	C_W3_1,		"ㄱ"    , 10, 2
    ggoSpread.SSSetCombo	C_W3_2,		"ㄴ"    , 10, 2
    ggoSpread.SSSetEdit		C_W3_3,		" ", 3, 2,,1,1
    ggoSpread.SSSetCombo	C_W3_4,		"ㄷ"    , 10, 2
    ggoSpread.SSSetCombo	C_W3_5,		"ㄹ"    , 10, 2
	ggoSpread.SSSetFloat	C_W4,		"(4)익금불산입율(%)" ,15,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
     
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_MINOR_CD,True)
	
	' 그리드 헤더 합침 정의 
	ret = .AddCellSpan(C_SEQ_NO		, -1000, 1, 2)	' SEQ_NO 합침 
	ret = .AddCellSpan(C_MINOR_NM	, -1000, 1, 2)	' SEQ_NO 합침 
	ret = .AddCellSpan(C_MINOR_CD	, -1000, 1, 2)	' SEQ_NO 합침 
	ret = .AddCellSpan(C_W1			, -1000, 1, 2)	' SEQ_NO 합침 
	ret = .AddCellSpan(C_W2			, -1000, 1, 2)	' SEQ_NO 합침 
	ret = .AddCellSpan(C_W3			, -1000, 5, 2)	' SEQ_NO 합침 
	ret = .AddCellSpan(C_W4			, -1000, 1, 2)	' SEQ_NO 합침 
	
     ' 첫번째 헤더 출력 글자 
	.Row = -1000
	.Col = C_W3
	.Text = "(3)지분율"
		
	' 두번째 헤더 출력 글자 
	'.Row = -999	
	'.Col = C_W3_1
	'.Text = "ㄱ(%)"
	'.Col = C_W3_2
	'.Text = "ㄴ"
	'.Col = C_W3_3	
	'.Text = " "
	'.Col = C_W3_4
	'.Text = "ㄷ(%)"
	'.Col = C_W3_5
	'.Text = "ㄹ"
	'
	.rowheight(-999) = 12	' 높이 재지정 
					
	Call InitSpreadComboBox2()
	
	.ReDraw = true
	
    Call SetSpreadLock2 
    
    End With   
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx

	' 차/대변 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " (MAJOR_CD='W1004') ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_COMPANY_TYPE
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_COMPANY_TYPE_NM
	End If
		  
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

End Sub

Sub InitSpreadComboBox2()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx, i, sVal, sVal2, sVal3
    
    ggoSpread.Source = frm1.vspddata2
    
    sVal = " " & vbTab
    For i = 0 to 100 Step 10
		sVal = sVal & CStr(i) & vbTab
    Next
    
    sVal2 = " " & vbTab & "초과" & vbTab & "이상" & vbTab 
    sVal3 = " " & vbTab & "이하" & vbTab & "미만" & vbTab

	ggoSpread.SetCombo sVal, C_W3_1
	ggoSpread.SetCombo sVal, C_W3_4
	ggoSpread.SetCombo sVal2, C_W3_2
	ggoSpread.SetCombo sVal3, C_W3_5

End Sub

' 환경변수1 %의 콤보값 인덱스 가져옴 
Function ReadCombo1(pVal)
	If pVal = "" Then
		ReadCombo1 = 0
	Else
		ReadCombo1 = (UNICDbl(pVal) / 10) + 1	' 공백포함 
	End If
End Function

' 환경변수2 초과등 콤보값 인덱스가져옴 
Function ReadCombo2(pVal)
	Select Case pVal
		Case ">"
			ReadCombo2 = 1
		Case ">="
			ReadCombo2 = 2
		Case "<"
			ReadCombo2 = 2
		Case "<="
			ReadCombo2 = 1
		Case Else
			ReadCombo2 = 0
	End Select
End Function

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
	ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
	'ggoSpread.SSSetRequired C_DOC_DATE, -1, -1
	'ggoSpread.SSSetRequired C_DOC_AMT, -1, -1
	ggoSpread.SSSetRequired C_COMPANY_NM, -1, -1
	ggoSpread.SSSetRequired C_STOCK_RATE, -1, -1
	ggoSpread.SSSetRequired C_DOC_AMT, -1, -1
	'ggoSpread.SSSetRequired C_ACQUIRE_AMT, -1, -1
	ggoSpread.SSSetRequired C_COMPANY_TYPE_NM, -1, -1
	ggoSpread.SSSetRequired C_JUKSU, -1, -1
	ggoSpread.SSSetRequired C_OWN_RGST_NO, -1, -1
	ggoSpread.SSSetRequired C_CO_ADDR, -1, -1
	ggoSpread.SSSetRequired C_REPRE_NM, -1, -1
	ggoSpread.SSSetRequired C_STOCK_CNT, -1, -1

    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadLock2()
    With frm1

    .vspdData2.ReDraw = False
    
	ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
	ggoSpread.SpreadLock C_W1, -1, C_W1, -1
	ggoSpread.SpreadLock C_W2, -1, C_W2, -1
	ggoSpread.SpreadLock C_W3_3, -1, C_W3_3, -1

    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
 
  	ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow
 	'ggoSpread.SSSetRequired C_DOC_DATE, pvStartRow, pvEndRow
 	ggoSpread.SSSetRequired C_DOC_AMT, pvStartRow, pvEndRow
 	ggoSpread.SSSetRequired C_COMPANY_NM, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_STOCK_RATE, pvStartRow, pvEndRow
	'ggoSpread.SSSetRequired C_ACQUIRE_AMT, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_COMPANY_TYPE_NM, pvStartRow, pvEndRow
	'ggoSpread.SSSetRequired C_HOLDING_TERM, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_JUKSU, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_OWN_RGST_NO, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_CO_ADDR, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_REPRE_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_STOCK_CNT, pvStartRow, pvEndRow
        
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO			= iCurColumnPos(1)
            C_DOC_DATE			= iCurColumnPos(2)
            C_DOC_AMT			= iCurColumnPos(3)
            C_DEBIT_CREDIT		= iCurColumnPos(4)
            C_DEBIT_CREDIT_NM	= iCurColumnPos(5)
            C_SUMMARY_DESC		= iCurColumnPos(6)
            C_COMPANY_NM		= iCurColumnPos(7)
            C_STOCK_RATE		= iCurColumnPos(8)
            C_ACQUIRE_AMT       = iCurColumnPos(9)
            C_COMPANY_TYPE		= iCurColumnPos(10)
            C_COMPANY_TYPE_NM	= iCurColumnPos(11)
            C_HOLDING_TERM		= iCurColumnPos(12)
            C_OWN_RGST_NO		= iCurColumnPos(13)
            C_CO_ADDR			= iCurColumnPos(14)
            C_REPRE_NM			= iCurColumnPos(15)
            C_STOCK_CNT			= iCurColumnPos(16)
    End Select    
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	lgCurrGrid = TYPE_1
	
End Sub

'============================================  조회조건 함수  ====================================

'============================== 레퍼런스 함수  ========================================

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg
	
	If gSelframeFlg = TAB2 Then Exit Function
	
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If

    'Call ggoOper.ClearField(Document, "2")	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call InitVariables 
    			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
	' 2. 대차대조표의 자산총계, 부채총계-미지급법인세, 자본금+미지급법인세+주식발행초과금+감자차익-주식발행할인차금-감자차손 가져오기 
	lgBlnFlgChgValue = True
End Function

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
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_2
	
End Function

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110110100101111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()

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


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_DEBIT_CREDIT
				.Col = Col
				intIndex = .Value
				.Col = C_DEBIT_CREDIT_NM
				.Value = intIndex	
			Case  C_DEBIT_CREDIT_NM
				.Col = Col
				intIndex = .Value
				.Col = C_DEBIT_CREDIT
				.Value = intIndex		
			Case C_COMPANY_TYPE
				.Col = Col
				intIndex = .Value
				.Col = C_COMPANY_TYPE_NM
				.Value = intIndex	
			Case C_COMPANY_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_COMPANY_TYPE
				.Value = intIndex	
		End Select
	End With
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblAmt1, dblAmt2, dblSum
	With frm1.vspdData
	
    .Row = Row
    .Col = Col

    If .CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(.text) < CDbl(.TypeFloatMin) Then
         .text = .TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    Select Case Col
		Case C_ACQUIRE_AMT, C_HOLDING_TERM
			.Col = C_ACQUIRE_AMT	: dblAmt1 = UNICDbl(.value)
			.Col = C_HOLDING_TERM	: dblAmt2 = UNICDbl(.value)
			
			if dblAmt2 > 365 then
				Call DisplayMsgBox("970028", parent.VB_INFORMATION, "당기보유기간이 365일", "X")     
				'.value = 0                     
			End If
			
			dblSum = dblAmt1 * dblAmt2
			.Col = C_JUKSU			: .value = dblSum
		Case C_DOC_AMT
			.Row = Row : .Col = Col
			If UNICDbl(.value) < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "금액", "X")     
				.value = 0                     
				Exit Sub
			End If
		
    End Select
    
	End With
End Sub

Function GetHead(Byval pCol)
	With frm1.vspdData
		.Col = pCol : .Row = 0	: GetHead = .Text
	End With
End Function

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

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

Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 

End Sub

'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col

    If Frm1.vspdData2.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData2.text) < CDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

End Sub


Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData2.Row = Row
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2

End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
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

    Call SetToolbar("1100110000001111")

	Call ClickTab1()
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
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
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
    Dim blnChange, dblSum
    Dim i
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    for i = 1 to frm1.vspdData.MaxRows
    frm1.vspdData.col = C_HOLDING_TERM
    frm1.vspdData.row = i
	
		if UNICDbl(frm1.vspdData.value) > 365 then
					Call DisplayMsgBox("970028", parent.VB_INFORMATION, "당기보유기간이 365일", "X")     
					Exit Function                    
		End If
    
    next    
    
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    

	ggoSpread.Source = frm1.vspdData2
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
	
	dblSum = FncSumSheet(frm1.vspdData, C_DOC_AMT, 1, frm1.vspdData.MaxRows, false, -1, -1, "V")
	
	If dblSum < 0 Then
		Call DisplayMsgBox("WC0013", parent.VB_INFORMATION, "(금액)", "X")                          
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
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			.vspdData.Col = C_DOC_AMT
			.vspdData.Text = ""
    
			.vspdData.Col = C_COMPANY_NM
			.vspdData.Text = ""
			
			.vspdData.Col = C_STOCK_RATE
			.vspdData.Text = ""
			
			.vspdData.Col = C_ACQUIRE_AMT
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
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
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
    
    If lgCurrGrid = TYPE_2 Then 
		Call InitGrid2
		Exit Function
	End If
	
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
   
	With frm1	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
			
		' SEQ_NO 를 그리드에 넣는 로직 
		iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
		ggoSpread.InsertRow ,imRow	' 그리드 행 추가(사용자 행수 포함)
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1	' 그리드 색상변경 
		
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
			.vspdData.Row = iRow
			.vspdData.Col = C_SEQ_NO
			.vspdData.Text = iSeqNo
			iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.
		Next				
		.vspdData.ReDraw = True	

		''SetSpreadColor .vspdData.ActiveRow    
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

Function InitGrid2()    
	Dim i, iRow, iCol, iMaxRows, ret, sField, sFrom , sWhere, arrMinorNm, arrMinorCd, arrSeqNo, arrRef, arrW1_W2, sOldW1, sOldW2
	Dim soldMinorNm, soldMinorCd, iSpanRowW1, iSpanRowW2, iSpanCntW1, iSpanCntW2
	
	If frm1.vspdData2.MaxRows > 0 Then Exit Function
	
	sField	= "	A.MINOR_NM, B.MINOR_CD, B.SEQ_NO, B.REFERENCE"
	sFrom	= " B_MINOR A " & vbCrLf
	sFrom	= sFrom	& " 	INNER JOIN B_CONFIGURATION B WITH (NOLOCK) ON A.MAJOR_CD=B.MAJOR_CD AND A.MINOR_CD=B.MINOR_CD "
	sWhere	= " A.MAJOR_CD='W2003' "

	Call CommonQueryRs(sField, sFrom, sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		soldMinorCd	= "" : iSpanCntW1 = 0 : iSpanCntW2 = 0 : iSpanRowW1 = 0 : iSpanRowW2 = 0 : iRow = 0
		arrMinorNm	= Split(lgF0 , Chr(11))
		arrMinorCd	= Split(lgF1 , Chr(11))
		arrSeqNo	= Split(lgF2 , Chr(11))
		arrRef		= Split(lgF3 , Chr(11))
		
		iMaxRows = UBound(arrMinorNm)
		
		For i = 0 to iMaxRows -1
			
			ggoSpread.InsertRow , 1
			.Row = iRow + 1 : iRow = iRow + 1 : .Col = C_SEQ_NO : .Value = iRow
			
			.Col = C_MINOR_NM	: .Value = arrMinorNm(i)
			.Col = C_MINOR_CD	: .Value = arrMinorCd(i)
			
			arrW1_W2 = Split(arrMinorNm(i), "|")	' 코드명을 | 로 분리한다.
			
			If sOldW1 <> arrW1_W2(0) Then	'C_W1 비교 
				iSpanCntW1 = 1 : iSpanRowW1 = .Row
			Else
				iSpanCntW1 = iSpanCntW1 + 1
				ret = .AddCellSpan(C_W1, iSpanRowW1, 1, iSpanCntW1)	' A1-A5 합침 
			End If
			
			If sOldW2 <> arrW1_W2(1) Then	' C_W2
				iSpanCntW2 = 1	: iSpanRowW2 = .Row
			Else 
				iSpanCntW2 = iSpanCntW2 + 1
				ret = .AddCellSpan(C_W2, iSpanRowW2, 1, iSpanCntW2)	' A1-A5 합침 
			End If		
			
			.Col = C_W1	: .Value = arrW1_W2(0)
			.Col = C_W2	: .Value = arrW1_W2(1)
				
			For iCol = 1 To 6	' 컬럼 갯수 
				.Col = C_W2 + iCol 
				Select Case iCol
					Case 1, 4
						.Value = ReadCombo1(arrRef(i)) : i = i + 1
					Case 2, 5
						.Value = ReadCombo2(arrRef(i)) : i = i + 1
					Case 6
						.Value = arrRef(i) ' For 문에서 i 값이 증가한다.
					Case 3
						.Value = "~"
				End Select

				
			Next

			sOldW1	= arrW1_W2(0)
			sOldW2	= arrW1_W2(1)
		Next
		
		Call SetSpreadLock2
		
	End With

End Function

' -- 그리드 span
Function SetGridSpan()
	Dim soldMinorCd, iSpanCntW1, iSpanCntW2, iSpanRowW1, iSpanRowW2, iRow, i, sMinorNm, arrW1_W2, sOldW1, sOldW2, ret, iMaxRows
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		soldMinorCd	= "" : iSpanCntW1 = 0 : iSpanCntW2 = 0 : iSpanRowW1 = 0 : iSpanRowW2 = 0 : iRow = 0

		
		iMaxRows = .MaxRows
		
		For i = 1 to iMaxRows 
			
			.Row = i : .Col = C_MINOR_NM : sMinorNm = .Text
			
			arrW1_W2 = Split(sMinorNm, "|")	' 코드명을 | 로 분리한다.
			
			If sOldW1 <> arrW1_W2(0) Then	'C_W1 비교 
				iSpanCntW1 = 1 : iSpanRowW1 = .Row
			Else
				iSpanCntW1 = iSpanCntW1 + 1
				ret = .AddCellSpan(C_W1, iSpanRowW1, 1, iSpanCntW1)	' A1-A5 합침 
			End If
			
			If sOldW2 <> arrW1_W2(1) Then	' C_W2
				iSpanCntW2 = 1	: iSpanRowW2 = .Row
			Else 
				iSpanCntW2 = iSpanCntW2 + 1
				ret = .AddCellSpan(C_W2, iSpanRowW2, 1, iSpanCntW2)	' A1-A5 합침 
			End If	

			.Col = C_W1	: .Value = arrW1_W2(0)
			.Col = C_W2	: .Value = arrW1_W2(1)
			
			sOldW1	= arrW1_W2(0)
			sOldW2	= arrW1_W2(1)
		Next
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
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
	
    ggoSpread.Source = frm1.vspdData	
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
        strVal = strVal     & "&txtCurrGrid="        & lgCurrGrid      
		
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
    If frm1.vspdData.MaxRows > 0 Or frm1.vspdData2.MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		Call SetGridSpan
		
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1111111100011111")										<%'버튼 툴바 제어 %>

		Else
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLock	1, -1, frm1.vspdData.MaxCols
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLock	1, -1, frm1.vspdData2.MaxCols
			
			Call SetToolbar("1110000000011111")										<%'버튼 툴바 제어 %>
		End If
	Else
		Call SetToolbar("1110110100001111")										<%'버튼 툴바 제어 %>
	End If
	'frm1.vspdData.focus			
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
    
	With Frm1
	
		With frm1.vspdData
		
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			For lRow = 1 To lMaxRows
    
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
				  	For lCol = 1 To lMaxCols
				  		.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				  	Next
				  	strVal = strVal & Trim(.Text) &  Parent.gRowSep
				End If  
			Next
       End With
       .txtSpread.value        =  strDel & strVal
       strDel = ""	: strVal = ""
       
       ' 2번 그리드 
		With frm1.vspdData2
			
			ggoSpread.Source = frm1.vspdData2
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
       
       frm1.txtSpread2.value        =  strDel & strVal
		.txtMode.value        =  Parent.UID_M0002
		'.txtUpdtUserId.value  =  Parent.gUsrID
		'.txtInsrtUserId.value =  Parent.gUsrID
		.txtCurrGrid.value     = lgCurrGrid
	End With

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>익금불산입율 등록</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:GetRef">금액불러오기</A>&nbsp;</TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w2103ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
									<script language =javascript src='./js/w2103ma1_vaSpread1_vspdData.js'></script>
								</DIV>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
									<script language =javascript src='./js/w2103ma1_vaSpread2_vspdData2.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

