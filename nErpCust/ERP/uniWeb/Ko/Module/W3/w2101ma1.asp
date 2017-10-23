
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 수입금액조정 
'*  3. Program ID           : W1111MA1
'*  4. Program Name         : W1111MA1.asp
'*  5. Program Desc         : 제16-2호 수입배당금 명세서 
'*  6. Modified date(First) : 2004/12/30
'*  7. Modified date(Last)  : 2004/12/30
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "W2101MA1"
Const BIZ_PGM_ID = "w2101mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID = "w2101mb2.asp"
Const EBR_RPT_ID		= "W2101OA1"
Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.


Dim C_SEQ_NO1
Dim C_W7
Dim C_W8
Dim C_W8_NM
Dim C_W9
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13

Dim C_SEQ_NO2
Dim C_W14
Dim C_W15
Dim C_W16
Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W21
Dim C_HEAD_SEQ_NO1

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(1)
Dim	lgTB_26_AMT, lgTB_3_AMT	

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_SEQ_NO1	= 1	' -- 1번 그리드 
    C_W7		= 2
    C_W8		= 3
    C_W8_NM		= 4
    C_W9		= 5
    C_W10		= 6
    C_W11		= 7
    C_W12		= 8
    C_W13		= 9	
 
 	C_SEQ_NO2	= 1  ' -- 2번 그리드 
    C_W14		= 2 
    C_W15		= 3
    C_W16		= 4
    C_W17		= 5
    C_W18		= 6
    C_W19		= 7
    C_W20		= 8
    C_W21		= 9
    C_HEAD_SEQ_NO1 = 10
    
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
    
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1003' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW2 ,lgF0  ,lgF1  ,Chr(11))    
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1

	' 1번 그리드 
	With lgvspdData(TYPE_1)
	
	ggoSpread.Source = lgvspdData(TYPE_1)	
   'patch version
    ggoSpread.Spreadinit "V20041222" & TYPE_1,,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_W13 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols									'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    ' 
    Call AppendNumberPlace("6","3","1")

    ggoSpread.SSSetEdit		C_SEQ_NO1,	"순번"		, 10,,,15,1	' 히든컬럼 
	ggoSpread.SSSetEdit		C_W7,		"(7)법인명"	, 20,,,50,1	
    ggoSpread.SSSetCombo	C_W8,		"(8)구분"		, 10
    ggoSpread.SSSetCombo	C_W8_NM,	"(8)구분"		, 10	
    ggoSpread.SSSetEdit		C_W9,		"(9)사업자등록번호", 20,,,20,1
    ggoSpread.SSSetEdit		C_W10,		"(10)소재지"	, 30,,,250,1
    ggoSpread.SSSetEdit		C_W11,		"(11)대표"		, 15,,,50,1
    ggoSpread.SSSetFloat	C_W12,		"(12)발행주식총수" , 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
    ggoSpread.SSSetFloat	C_W13,		"(13)지분율"	, 10, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","100" 

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_W8,C_W8,True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO1,C_SEQ_NO1,True)
				
	Call InitSpreadComboBox()
	
	.ReDraw = true
	
    'Call SetSpreadLock 
    
    End With

 	' -----  2번 그리드 
	With lgvspdData(TYPE_2)
	
	ggoSpread.Source = lgvspdData(TYPE_2)	
   'patch version
    ggoSpread.Spreadinit "V20041222_2" & TYPE_2,,parent.gAllowDragDropSpread    
    
	.ReDraw = false
    
    .MaxCols = C_HEAD_SEQ_NO1 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols									'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

	'헤더를 2줄로    
    .ColHeaderRows = 2
    'Call AppendNumberPlace("6","3","2")

    ggoSpread.SSSetEdit		C_SEQ_NO2,	"순번", 10,,,15,1	' 히든컬럼 
	ggoSpread.SSSetEdit		C_W14,		"(14)자회사 또는" & vbCrLf & "배당금지급 법인명", 20,,,50,1
    ggoSpread.SSSetFloat	C_W15,		"(15)배당금액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
    ggoSpread.SSSetCombo	C_W16,		"(16)익금불산입율" , 10, 1
	ggoSpread.SSSetFloat	C_W17,		"(17)익금불산입 대상금액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetFloat	C_W18,		"(18)소계",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
    ggoSpread.SSSetFloat	C_W19,		"(19)법$18의2 (법$18의3)",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
    ggoSpread.SSSetFloat	C_W20,		"(20)법$18의2 제1항제4호" ,15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W21,		"(21)익금 불산입액" & vbCrLf & "(17-18)" , 15,Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"","0",""
	ggoSpread.SSSetEdit		C_HEAD_SEQ_NO1,	"헤더순번", 10,,,15,1	' 히든컬럼 
 
    ret = .AddCellSpan(1, -1000, 1, 2)
    ret = .AddCellSpan(2, -1000, 1, 2)
    ret = .AddCellSpan(3, -1000, 1, 2)
    ret = .AddCellSpan(4, -1000, 1, 2)
    ret = .AddCellSpan(5, -1000, 1, 2)
    ret = .AddCellSpan(6, -1000, 3, 1)
    ret = .AddCellSpan(9, -1000, 1, 2) 
    ret = .AddCellSpan(10, -1000, 1, 2) 
    
    ' 첫번째 헤더 출력 글자 
	.Row = -1000
	.Col = 6
	.Text = "익금불산입차감금액"

	' 두번째 헤더 출력 글자 
	.Row = -999
	.Col = 6
	.Text = "(18)소계"
	.Col = 7
	.Text = "(19)법$18의2" & vbCrLf & "(법$18의3)"
	.Col = 8
	.Text = "(20)법$18의2" & vbCrLf & "제1항제4호"
	.rowheight(-999) = 20	' 높이 재지정 
	
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO2,C_SEQ_NO2,True)
	Call ggoSpread.SSSetColHidden(C_HEAD_SEQ_NO1,C_HEAD_SEQ_NO1,True)
				
	Call InitSpreadComboBox2()
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
       
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx

	' 회사구분 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " (MAJOR_CD='W1004') ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(TYPE_1)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W8
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W8_NM
	End If
		  
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

End Sub

Sub InitSpreadComboBox2()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx, i, sVal, sVal2, sVal3
    
    ggoSpread.Source = lgvspdData(TYPE_2)
    
    sVal = "100" & vbTab & "90" & vbTab & "60" & vbTab & "50" & vbTab & "30"

	ggoSpread.SetCombo sVal, C_W16

End Sub

Sub SetSpreadLock()

    lgvspdData(TYPE_2).ReDraw = False
    ggoSpread.Source = lgvspdData(TYPE_2)	

    ggoSpread.SpreadLock C_W17, -1, C_W17
    ggoSpread.SpreadLock C_W18, -1, C_W18
    ggoSpread.SpreadLock C_W21, -1, C_W21
	ggoSpread.SSSetRequired C_W14, -1, -1
	ggoSpread.SSSetRequired C_W15, -1, -1
	ggoSpread.SSSetRequired C_W16, -1, -1
	'ggoSpread.SSSetRequired C_W19, -1, -1
	ggoSpread.SpreadLock C_W14, lgvspdData(TYPE_2).MaxRows, C_W21
	lgvspdData(TYPE_2).ReDraw = True
	
	lgvspdData(TYPE_1).ReDraw = False
    ggoSpread.Source = lgvspdData(TYPE_1)
        
	ggoSpread.SSSetRequired C_W7, -1, -1
	ggoSpread.SSSetRequired C_W8, -1, -1
	ggoSpread.SSSetRequired C_W8_NM, -1, -1
	ggoSpread.SSSetRequired C_W9, -1, -1
	ggoSpread.SSSetRequired C_W10, -1, -1
	ggoSpread.SSSetRequired C_W11, -1, -1
	ggoSpread.SSSetRequired C_W12, -1, -1
	ggoSpread.SSSetRequired C_W13, -1, -1
	lgvspdData(TYPE_1).ReDraw = True
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
    With lgvspdData(pType)

	If pType = TYPE_1 Then
		.ReDraw = False
 
		ggoSpread.Source = lgvspdData(pType)
	
  		ggoSpread.SSSetRequired C_W7, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W8, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W8_NM, pvStartRow, pvEndRow 		
 		ggoSpread.SSSetRequired C_W9, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W10, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W11, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W12, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W13, pvStartRow, pvEndRow
		    
		.ReDraw = True
    Else
		.ReDraw = False
 
		ggoSpread.Source = lgvspdData(pType)
	
		ggoSpread.SSSetRequired C_W14, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W15, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W16, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_W17, pvStartRow, pvEndRow
  		ggoSpread.SSSetProtected C_W18, pvStartRow, pvEndRow
  		'ggoSpread.SSSetRequired C_W19, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_W21, pvStartRow, pvEndRow
  		
		.ReDraw = True    
    End If
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case pvSpdNo
       Case TYPE_1
            ggoSpread.Source = lgvspdData(TYPE_1)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_SEQ_NO1	= iCurColumnPos(1)	' -- 1번 그리드 
			C_W7		= iCurColumnPos(2)
			C_W8		= iCurColumnPos(3)
			C_W8_NM		= iCurColumnPos(4)
			C_W9		= iCurColumnPos(5)
			C_W10		= iCurColumnPos(6)
			C_W11		= iCurColumnPos(7)
			C_W12		= iCurColumnPos(8)
			C_W13		= iCurColumnPos(9)	
			
 		Case TYPE_2
 			ggoSpread.Source = lgvspdData(TYPE_2)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
  			C_SEQ_NO2	= iCurColumnPos(1)  ' -- 2번 그리드 
			C_W14		= iCurColumnPos(2) 
			C_W15		= iCurColumnPos(3)
			C_W16		= iCurColumnPos(4)
			C_W17		= iCurColumnPos(5)
			C_W18		= iCurColumnPos(6)
			C_W19		= iCurColumnPos(7)
			C_W20		= iCurColumnPos(8)
			C_W21		= iCurColumnPos(9)
			C_HEAD_SEQ_NO1 = iCurColumnPos(10)
               
    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function


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

    'Call ggoOper.ClearField(Document, "2")	
    ggoSpread.Source = lgvspdData(TYPE_1)
    ggoSpread.ClearSpreadData
    
    ggoSpread.Source = lgvspdData(TYPE_2)
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

Sub PutGrid(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	If pType = TYPE_1 Then
		With lgvspdData(TYPE_1)
			.Col = pCol	: .Row = pRow : .Value = pVal
		End With
	Else
		With lgvspdData(TYPE_2)
			.Col = pCol	: .Row = pRow : .Value = pVal
		End With
	End If
End Sub

Sub PutGridText(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	If pType = TYPE_1 Then
		With lgvspdData(TYPE_1)
			.Col = pCol	: .Row = pRow : .Text = pVal
		End With
	Else
		With lgvspdData(TYPE_2)
			.Col = pCol	: .Row = pRow : .Text = pVal
		End With
	End If
End Sub

' -- 배당내역불러오기 에서 호출됨 
Sub ReClacGrid2()
	Dim dblVal(30), iMaxRows, iRow
	
	With lgvspdData(TYPE_2)
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			
			'Call vspdData_Change(TYPE_2, C_W15, iRow)
			'Call vspdData_Change(TYPE_2, C_W19, iRow)
			'.Col = C_W15	: dblVal(C_W15) = UNICDbl(.value)
			'.Col = C_W16	: dblVal(C_W16) = UNICDbl(.Text)
			'If dblVal(C_W15) = 0 And dblVal(C_W16) = 0 Then Exit Sub
			'dblVal(C_W17) = dblVal(C_W15) * (dblVal(C_W16) * 0.01)
			'.Col = C_W17	: .value = dblVal(C_W17)		
		
		Next
	
	
	End With
End Sub

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100011111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

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

Sub InitData()
	Dim sCoCd , sFiscYear, sRepType
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	lgCurrGrid = TYPE_1
	
	Dim iRet
	sCoCd		= "<%=wgCO_CD%>"
	sFiscYear	= "<%=wgFISC_YEAR%>"
	sRepType	= "<%=wgREP_TYPE%>"
	' 법인정보 출력 
	iRet = CommonQueryRs("W1, W2, W3, W4, W5, W6"," dbo.ufn_TB_16_2_GetRef4('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If iRet = True Then
		With frm1
		.txtW1.value = Replace(lgF0, Chr(11), "")
		.cboW2.value = Replace(lgF1, Chr(11), "")
		.txtW3.value = Replace(lgF2, Chr(11), "")
		.txtW4.value = Replace(lgF3, Chr(11), "")
		.txtW5.value = Replace(lgF4, Chr(11), "")
		.txtW6.value = Replace(lgF5, Chr(11), "")
		
		Call cboW2_onChange
		End With
    End If
End Sub

Sub cboW2_onChange()
	With frm1
		.txtW2_NM.value =	.cboW2.options(.cboW2.selectedIndex).Text
	End With
	lgBlnFlgChgValue = True
End Sub

Sub txtW1_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW3_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW4_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW5_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW6_onChange()
	lgBlnFlgChgValue = True
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

' -- 1번 그리드 
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

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
Sub vspdData_ComboSelChange(Byval pType, ByVal Col, ByVal Row)
	Dim iIdx, dblVal(30)
	
	lgBlnFlgChgValue = True
	With lgvspdData(pType)
	
		Select Case pType
			Case TYPE_2
				If Col = C_W16 Then
					.Row = Row
					.Col = C_W15	: dblVal(C_W15) = UNICDbl(.value)
					.Col = C_W16	: dblVal(C_W16) = UNICDbl(.Text)
					If dblVal(C_W15) = 0 And dblVal(C_W16) = 0 Then Exit Sub
					dblVal(C_W17) = dblVal(C_W15) * (dblVal(C_W16) * 0.01)
					.Col = C_W17	: .value = dblVal(C_W17)
				End If
			
			Case TYPE_1
				If Col = C_W8_NM Then
					.Row = Row
					.Col = C_W8_NM	: iIdx = UNICDbl(.Value)
					.Col = C_W8		: .Value = iIdx
				End If
		End Select
	End With
End Sub
	
Sub vspdData_Change(Byval pType, ByVal Col , ByVal Row )
    lgvspdData(pType).Row = Row
    lgvspdData(pType).Col = Col

    If lgvspdData(TYPE_1).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(TYPE_1).text) < UNICDbl(lgvspdData(TYPE_1).TypeFloatMin) Then
         lgvspdData(TYPE_1).text = lgvspdData(TYPE_1).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(pType)
    ggoSpread.UpdateRow Row

	If pType = TYPE_1 Then Exit Sub
	
	Dim dblVal(30)
	
	With lgvspdData(TYPE_2)
	
		Select Case Col
			Case C_W15, C_W16	' W17계산 
				.Row = Row
				.Col = C_W15	: dblVal(C_W15) = UNICDbl(.value)
				.Col = C_W16	: dblVal(C_W16) = UNICDbl(.Text)
				If dblVal(C_W15) = 0 And dblVal(C_W16) = 0 Then Exit Sub
				dblVal(C_W17) = dblVal(C_W15) * (dblVal(C_W16) * 0.01)
				.Col = C_W17	: .value = dblVal(C_W17)
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' 합계 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' 합계 
			
				' C_W17 을 변경햇으므로 해당이벤트를 발생해 준다.
				Call vspdData_Change(pType, C_W17, Row)
				
			Case C_W17, C_W18
				.Row = Row
				.Col = C_W17	: dblVal(C_W17) = UNICDbl(.value)
				.Col = C_W18	: dblVal(C_W18) = UNICDbl(.value)
				
				dblVal(C_W21) = dblVal(C_W17) - dblVal(C_W18)
				.Col = C_W21	: .value = dblVal(C_W21)
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' 합계 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W21, 1, .MaxRows - 1, true, .MaxRows, C_W21, "V")	' 합계 
			
				' C_W21 을 변경햇으므로 해당이벤트를 발생해 준다.
				Call vspdData_Change(pType, C_W21, Row)
				
			Case C_W19, C_W20
				.Row = Row
				.Col = C_W19	: dblVal(C_W19) = UNICDbl(.value)
				.Col = C_W20	: dblVal(C_W20) = UNICDbl(.value)
				
				dblVal(C_W18) = dblVal(C_W19) + dblVal(C_W20)
				.Col = C_W18	: .value = dblVal(C_W18)
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' 합계 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' 합계		
			
				' C_W18 을 변경햇으므로 해당이벤트를 발생해 준다.
				Call vspdData_Change(pType, C_W18, Row)
				
			Case C_W14 
				' -- 그리드1의 (7)법인명을 검색해 SEQ_NO를 C_HEAD_SEQ_NO1에 넣는다 
				.Row = Row
				.Col = Col
				dblVal(C_HEAD_SEQ_NO1) = SearchW7(.Text)
				
				If dblVal(C_HEAD_SEQ_NO1) = -1 Then
					Call DisplayMsgBox("W20002", parent.VB_INFORMATION, .Text, "X")
					.Value = ""
					Exit Sub
				End If
				.Col = C_HEAD_SEQ_NO1	: .value = dblVal(C_HEAD_SEQ_NO1)
		End Select
		
		ggoSpread.UpdateRow .MaxRows
	End With
End Sub

Sub ReClacGrid2Sum()
	Dim iCol
	With lgvspdData(TYPE_2)
	For iCol = C_W15 To C_W21
		If iCol <> C_W16 Then
			Call FncSumSheet(lgvspdData(lgCurrGrid), iCol, 1, .MaxRows - 1, true, .MaxRows, iCol, "V")	' 합계 
		End If
	Next
	End With
End Sub

Function SearchW7(Byval pCoNm)
	Dim iMaxRows, iRow
	
	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_W7
			If UCase(.Text) = UCase(pCoNm) Then
				.Col = C_SEQ_NO1
				SearchW7 = UNICDbl(.Value)
				Exit Function
			End If
		Next

	End With
	SearchW7 = -1	' --없을경우 
End Function

Sub vspdData_Click(Byval pType, ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(pType)
   
    If lgvspdData(pType).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(pType)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	lgvspdData(TYPE_1).Row = Row
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

Sub vspdData_GotFocus(Byval pType )
    ggoSpread.Source = lgvspdData(pType)
End Sub

Sub vspdData_MouseDown(Byval pType, Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	ggoSpread.Source = lgvspdData(pType)
End Sub    

Sub vspdData_ScriptDragDropBlock(Byval pType, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = lgvspdData(TYPE_1)
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos(pType)
End Sub

Sub vspdData_TopLeftChange(Byval pType, ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if lgvspdData(TYPE_1).MaxRows < NewTop + VisibleRowCnt(lgvspdData(TYPE_1),NewTop) Then	           
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
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = lgvspdData(TYPE_1)
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
    
    
    Dim blnChange
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    

<%  '-----------------------
    'Precheck area
    '----------------------- %> 

    If Not chkField(Document, "2") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
    
    ggoSpread.Source = lgvspdData(TYPE_1)
    If ggoSpread.SSCheckChange = False Then
        blnChange = False
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    
	
    ggoSpread.Source = lgvspdData(TYPE_2)
    If ggoSpread.SSCheckChange = False Then
        blnChange = False
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  

    If lgBlnFlgChgValue = False  Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
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

    Call SetToolbar("1100110100011111")

	Call InitData
	
	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If lgvspdData(lgCurrGrid).MaxRows = 0 Then
       Exit Function
    End If
 
	With lgvspdData(lgCurrGrid)

		.focus
		.ReDraw = False
				
		Select Case lgCurrGrid
			Case TYPE_1
				ggoSpread.Source = lgvspdData(TYPE_1)
		
				ggoSpread.CopyRow
				SetSpreadColor .ActiveRow, .ActiveRow
				MaxSpreadVal lgvspdData(TYPE_1), C_SEQ_NO1, iRow

			Case TYPE_2
				ggoSpread.Source = lgvspdData(TYPE_2)
		
				ggoSpread.CopyRow
				SetSpreadColor .ActiveRow, .ActiveRow
				MaxSpreadVal lgvspdData(TYPE_2), C_SEQ_NO2, iRow
		End Select
	
		.ReDraw = True
	
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    ggoSpread.Source = lgvspdData(TYPE_1)	
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
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
 
	With lgvspdData(lgCurrGrid)	' 포커스된 그리드 
			
		ggoSpread.Source = lgvspdData(lgCurrGrid)
			
		iRow = .ActiveRow
		lgvspdData(lgCurrGrid).ReDraw = False
					
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1			
			If lgCurrGrid = TYPE_1 Then
				ggoSpread.InsertRow , 1
				Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
				
				.Col = C_SEQ_NO1 : .Text = iRow	
			Else
				ggoSpread.InsertRow , 2
				Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
				
				.Col = C_SEQ_NO2 : .Text = iRow	
			
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO2: .Text = SUM_SEQ_NO	
				.Col = C_W14	: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
				.Col = C_W16	: .CellType = 1	: .TypeHAlign = 1
				ggoSpread.SpreadLock C_W14, iRow, .MaxCols-1, iRow
			End If		
		
		Else
				
			If iRow = .MaxRows Then	' -- 마지막 합계줄에서 InsertRow를 하면 상위에 추가한다.
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor lgCurrGrid, iRow, iRow + imRow - 1

				'If lgCurrGrid = TYPE_1 Then
					Call SetDefaultVal(lgCurrGrid, iRow, imRow)
				'End If
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor lgCurrGrid, iRow+1, iRow + imRow

				'If lgCurrGrid = TYPE_1 Then
					Call SetDefaultVal(lgCurrGrid, iRow+1, imRow)
				'End If
			End If   
		End If
	End With
	

	'Call CheckW7Status(lgCurrGrid)	' 적수셀 상태 체크 

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetDefaultVal(Index, iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(Index)	' 포커스된 그리드 

	ggoSpread.Source = lgvspdData(Index)
	
	If iAddRows = 1 Then ' 1줄만 넣는경우 
		.Row = iRow
		MaxSpreadVal lgvspdData(Index), C_SEQ_NO1, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(Index), C_SEQ_NO2, iRow)	' 현재의 최대SeqNo를 구한다 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO2 : .Value = iSeqNo : iSeqNo = iSeqNo + 1
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows
	
	If lgCurrGrid = 1	Then
		With lgvspdData(TYPE_1) 
			.focus
			ggoSpread.Source = lgvspdData(TYPE_1) 
			lDelRows = ggoSpread.DeleteRow
		End With
    Else
		With lgvspdData(TYPE_2) 
			.focus
			ggoSpread.Source = lgvspdData(TYPE_2)
			lDelRows = ggoSpread.DeleteRow
		End With    
    End If
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
	
    ggoSpread.Source = lgvspdData(TYPE_1)	
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
        'strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
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
		
		Call SetSpreadLock()
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101111100011111")										<%'버튼 툴바 제어 %>
			
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100000000011111")										<%'버튼 툴바 제어 %>
		End If
	Else
		Call SetToolbar("1100110100011111")										<%'버튼 툴바 제어 %>
	End If
	
	lgvspdData(TYPE_1).focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
    Dim lRow, lCol, lMaxRows, lMaxCols , i 
 
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

		document.all("txtSpread" & CStr(i)).value =  strDel & strVal
		strDel = "" : strVal = ""
	Next
      
	frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
	frm1.txtFlgMode.value     = lgIntFlgMode
		

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	lgvspdData(TYPE_1).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_1)
    ggoSpread.ClearSpreadData

	lgvspdData(TYPE_2).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_2)
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
	Call InitVariables
	
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
						<a href="vbscript:GetRef">배당내역 불러오기</A> | <A href="vbscript:OpenRefMenu">소득금액합계표조회</A>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w2101ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<TABLE <%=LR_SPACE_TYPE_20%> BORDER=0>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%"> 1. 지주회사 또는 출자법인 현황</TD>
                            </TR>
                            <TR HEIGHT=10>
								<TD WIDTH="100%">
                                 <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
								     <TR>
								         <TD CLASS="TD51" width="17%" ALIGN=CENTER>(1)법인명</TD>
								         <TD CLASS="TD51" width="10%" ALIGN=CENTER>(2)구분</TD>
								         <TD CLASS="TD51" width="17%" ALIGN=CENTER>(3)사업자등록번호</TD>
								         <TD CLASS="TD51" width="25%" ALIGN=CENTER>(4)소재지</TD>
								         <TD CLASS="TD51" width="15%" ALIGN=CENTER>(5)대표자성명</TD>
								         <TD CLASS="TD51" width="10%" ALIGN=CENTER>(6)업태 업종</TD>
								    </TR>
								    <TR>
								  		<TD><INPUT NAME="txtW1" MAXLENGTH=25 TYPE="Text" ALT="법인명" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><SELECT NAME="cboW2" ALT="구분" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></SELECT><INPUT TYPE=HIDDEN NAME=txtW2_NM></TD>
								  		<TD><INPUT NAME="txtW3" MAXLENGTH=20  TYPE="Text" ALT="사업자등록번호" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><INPUT NAME="txtW4" MAXLENGTH=125  TYPE="Text" ALT="소재지" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><INPUT NAME="txtW5" MAXLENGTH=25  TYPE="Text" ALT="대표자성명" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><INPUT NAME="txtW6" MAXLENGTH=50  TYPE="Text" ALT="업태 업종" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								    </TR>
								</table>
								</TD>
							</TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">2. 자회사 또는 배당금 지급법인 현황</TD>
                            </TR>
                            <TR HEIGHT=40%>
                                <TD WIDTH="100%">
								<script language =javascript src='./js/w2101ma1_vaSpread1_vspdData0.js'></script></TD>
                            </TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">3. 수입배당금 및 익금불산입 금액 명세</TD>
                            </TR>   
                            <TR HEIGHT=60%>
                                <TD WIDTH="100%">
								<script language =javascript src='./js/w2101ma1_vaSpread1_vspdData1.js'></script></TD>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>자회사 또는 배당금지급법인 현황</LABEL>&nbsp;
							<INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><별지>수입배당금 및 익금불산입 금액 명세</LABEL>&nbsp;
				        
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
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

