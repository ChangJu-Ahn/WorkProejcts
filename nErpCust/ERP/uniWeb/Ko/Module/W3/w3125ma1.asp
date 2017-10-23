
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 과목별조정 
'*  3. Program ID           : W2125MA1
'*  4. Program Name         : W2125MA1.asp
'*  5. Program Desc         : 제26호 업무무관부동산등에 관련한 차입금이자조정명세서(갑)
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W3125MA1"
Const BIZ_PGM_ID		= "w3125mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "w3125mb2.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID		= "W3125OA1"

' -- 그리드 컬럼 정의 
Dim C_SEQ_NO
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W22
Dim C_W23
Dim C_W25
Dim C_W26
Dim C_W28
Dim C_W29
Dim C_W30
Dim C_W31


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim gCurrGrid, lgblnYoon

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1	' -- 1번 그리드 
    C_W18		= 2
    C_W19		= 3
    C_W20		= 4
    C_W22		= 5
    C_W23		= 6
    C_W25		= 7
    C_W26		= 8
    C_W28		= 9	
    C_W29		= 10	
    C_W30		= 11
    C_W31		= 12

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

    .MaxCols = C_W31 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols									'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

	'헤더를 2줄로    
    .ColHeaderRows = 2    ' 
    Call AppendNumberPlace("6","2","2")
	Call AppendNumberPlace("7","16","0")

    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번"		, 10,,,6,1	' 히든컬럼 
	ggoSpread.SSSetFloat	C_W18,		"(18)이자율"		, 10,		6					,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","100"
    ggoSpread.SSSetFloat	C_W19,		"(19)지급이자"		, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W20,		"(20)차입금적수"	, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W22,		"(22)지급이자"		, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W23,		"(23)차입금적수"	, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W25,		"(25)지급이자"		, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W26,		"(26)차입금적수"	, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W28,		"(28)지급이자"		, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W29,		"(29)차입금적수"	, 13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W30,		"(30)지급이자" & vbCrLf & "(19-22-25-28)"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetFloat	C_W31,		"(31)차입금적수" & vbCrLf & "(20-23-26-29)"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec


	' 그리드 헤더 합침 
    ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)	' 순번 2행 합침 
    ret = .AddCellSpan(C_W18	, -1000, 1, 2)	
    ret = .AddCellSpan(C_W19	, -1000, 1, 2)
    ret = .AddCellSpan(C_W20	, -1000, 1, 2)
    ret = .AddCellSpan(C_W22	, -1000, 2, 1)
    ret = .AddCellSpan(C_W25	, -1000, 2, 1)
    ret = .AddCellSpan(C_W28	, -1000, 2, 1) 
    ret = .AddCellSpan(C_W30	, -1000, 2, 1)
    
    ' 첫번째 헤더 출력 글자 
	.Row = -1000
	.Col = C_W22	: .Text = "(21)채권자불분"
	.Col = C_W25	: .Text = "(24)기준초과"
	.Col = C_W28	: .Text = "(27)건설자금이자등"
	.Col = C_W30	: .Text = "차 감"

	' 두번째 헤더 출력 글자 
	.Row = -999
	.Col = C_W22	: .Text = "(22)지급이자"
	.Col = C_W23	: .Text = "(23)차입금적수"
	.Col = C_W25	: .Text = "(25)지급이자"
	.Col = C_W26	: .Text = "(26)차입금적수"
	.Col = C_W28	: .Text = "(28)지급이자"
	.Col = C_W29	: .Text = "(29)차입금적수"
	.Col = C_W30	: .Text = "(30)지급이자" & vbCrLf & "(19-22-25-28)"
	.Col = C_W31	: .Text = "(31)차입금적수" & vbCrLf & "(20-23-26-29)"
	.rowheight(-999) = 20	' 높이 재지정	(2줄일 경우, 1줄은 15)
	
	
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
	Call InitSpreadComboBox()
	
	.ReDraw = true
	
    'Call SetSpreadLock 
    
    End With
      
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	Call CheckFISC_DATE
End Sub


Sub InitSpreadComboBox()

    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx

	' 회사구분 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " (MAJOR_CD='W1013') ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W19
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W20
	End If
		  
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

End Sub

Sub SetSpreadLock()

    ggoSpread.Source = frm1.vspdData	

    ggoSpread.SpreadLock C_W20, -1, C_W20
    ggoSpread.SpreadLock C_W30, -1, C_W30
    ggoSpread.SpreadLock C_W31, -1, C_W31
	ggoSpread.SSSetRequired C_W18, -1, -1
	ggoSpread.SSSetRequired C_W19, -1, -1

End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = frm1.vspdData
	
  	ggoSpread.SSSetRequired C_W18, pvStartRow, pvEndRow
 	ggoSpread.SSSetRequired C_W19, pvStartRow, pvEndRow
 	
 	ggoSpread.SSSetProtected C_W20, pvStartRow, pvEndRow 		
 	ggoSpread.SSSetProtected C_W30, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_W31, pvStartRow, pvEndRow
		    
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W18		= iCurColumnPos(2)
            C_W21		= iCurColumnPos(3)
            C_W19		= iCurColumnPos(4)
            C_W20		= iCurColumnPos(5)
            C_W21		= iCurColumnPos(6)
            C_W22		= iCurColumnPos(7)
            C_W23		= iCurColumnPos(8)
            C_W25       = iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W17		= iCurColumnPos(13)
            C_W18		= iCurColumnPos(14)
            C_W19		= iCurColumnPos(15)
            C_W20		= iCurColumnPos(16)
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

	Call fncNew()
	
	Dim strVal    
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
End Function


' 레퍼런스에서 넣었으므로 입력으로 변환해 준다.
Function ChangeRowFlg()
	Dim iRow
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
		Next
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


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
' GetREF 에서 적수 가져온뒤 호출됨 
Function InsertTotalLine()
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData
	
	If .MaxRows = 0 Then	' 한줄 추가 
		ggoSpread.InsertRow ,1
		
		.Row = 1
		.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
		.Col = C_W18	: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
		
		ggoSpread.SpreadLock C_W18, 1, C_W31, 1
	Else
		.Row = .MaxRows
		.Col = C_W18	: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
		
		ggoSpread.SpreadLock C_W18, .Row, C_W31, .Row		
	End If
	End With
End Function


Sub SetW4()
	Dim dblW2, dblW3
	dblW2 = UNICDbl(frm1.txtW2.value)
	dblW3 = UNICDbl(frm1.txtW3.value)

	frm1.txtW4.value = UNIFormatNumber((dblW2 + dblW3) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	
End Sub

Sub SetW7()
	Dim dblW5, dblW6
	dblW5 = UNICDbl(frm1.txtW5.value)
	dblW6 = UNICDbl(frm1.txtW6.value)
	
	If (dblW5 - dblW6) < 0 Then
		frm1.txtW7.value =  0
	Else
		frm1.txtW7.value = UNIFormatNumber((dblW5 - dblW6) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	End If
End Sub

Sub SetW8()
	Dim dblW4, dblW7
	dblW4 = UNICDbl(frm1.txtW4.value)
	dblW7 = UNICDbl(frm1.txtW7.value)

	If dblW4 < dblW7 Then
		frm1.txtW8.value = UNIFormatNumber(dblW4 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	Else
		frm1.txtW8.value = UNIFormatNumber(dblW7 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	End If
End Sub

Sub SetW9()
	Dim dblW1, dblW8, dblW5
	dblW1 = UNICDbl(frm1.txtW1.value)
	dblW8 = UNICDbl(frm1.txtW8.value)
	dblW5 = UNICDbl(frm1.txtW5.value)
	If dblW5 = 0 Then
		frm1.txtW9.value = 0
	Else
		frm1.txtW9.value = UNIFormatNumber((dblW1 * dblW8) / dblW5, ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	End If
End Sub

Sub SetW10()
	Dim dblW1, dblW9
	dblW1 = UNICDbl(frm1.txtW1.value)
	dblW9 = UNICDbl(frm1.txtW9.value)

	frm1.txtW10.value = UNIFormatNumber(dblW1 - dblW9 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
End Sub

Sub SetW13()
	'Dim dblW8, dblW9
	'dblW8 = UNICDbl(frm1.txtW8.value)
	'dblW9 = UNICDbl(frm1.txtW9.value)
	
	'If (dblW8 - dblW9) < 0 Then
	'	frm1.txtW13.value =  0
	'Else
	'	frm1.txtW13.value = UNIFormatNumber((dblW8 - dblW9) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	'End If
End Sub

Sub SetW14()
	Dim dblW11, dblW12, dblW13
	dblW11 = UNICDbl(frm1.txtW11.value)
	dblW12 = UNICDbl(frm1.txtW12.value)
	dblW13 = UNICDbl(frm1.txtW13.value)

	frm1.txtW14.value = UNIFormatNumber(dblW11 + dblW12 + dblW13 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
End Sub

Sub SetW15()
	Dim dblW31, dblW8
	With frm1.vspdData
		.Row = .MaxRows	: .Col = C_W31	: dblW31 = UNICDbl(.Value)
	End With

	dblW8 = UNICDbl(frm1.txtW8.value)

	frm1.txtW15.value = UNIFormatNumber(dblW31 - dblW8 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
End Sub

Sub SetW16()
	Dim dblW14, dblW15
	dblW14 = UNICDbl(frm1.txtW14.value)
	dblW15 = UNICDbl(frm1.txtW15.value)

	If dblW14 < dblW15 Then
		frm1.txtW16.value = UNIFormatNumber(dblW14 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	Else
		frm1.txtW16.value = UNIFormatNumber(dblW15 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	End If
End Sub

Sub SetW17()
	Dim dblW10, dblW16, dblW15
	dblW10 = UNICDbl(frm1.txtW10.value)
	dblW16 = UNICDbl(frm1.txtW16.value)
	dblW15 = UNICDbl(frm1.txtW15.value)

	If dblW15 = 0 Then
		frm1.txtW17.value = 0
	Else
		frm1.txtW17.value = UNIFormatNumber((dblW10 * dblW16) / dblW15 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	End If
End Sub

Sub SetW30()
	Dim dblSum
	
	' --- 추가된 부분 
	With Frm1.vspdData
		dblSum = FncSumSheet(frm1.vspdData, C_W30, 1, .MaxRows - 1, false, -1, -1, "V")	' 현재 컬럼 행합계 
		.Col = C_W30 : .Row = .MaxRows : .Value = dblSum
		

	End With
End Sub

Sub SetAllRecalc()
	Call SetW1_W5
	Call SetW4
	Call SetW7
	Call SetW8
	Call SetW9
	Call SetW10
	Call SetW13
	Call SetW14
	Call SetW15
	Call SetW16
	Call SetW17
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum, dblAmt, dblW18
	
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	' --- 추가된 부분 
	With Frm1.vspdData
	
	Select Case Col
		Case C_W19, C_W22, C_W25, C_W28 ' 지급이자 변경시 
			' 1. 자신의 썸 출력 
			Call FncSumSheet(frm1.vspdData, Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' 합계 
			
			' 2. 차입금적수 출력 
			.Row = Row
			.Col = Col		: dblAmt = UNICDbl(.Value)	' 현재컬럼값 
			.Col = C_W18	: dblW18 = UNICDbl(.Value)	' 이자율 

			If dblW18 = 0 Then Exit Sub
			' 3. 적수계산 
			If lgblnYoon Then
				' 윤년 
				dblSum = (dblAmt / (dblW18/100)) * 366
			Else	
				' 평년 
				dblSum = (dblAmt / (dblW18/100)) * 365
			End If
			
			.Col = Col + 1	: .Value = dblSum ' 현재셀 다음에 적수출력 
			
			Call FncSumSheet(frm1.vspdData,  Col + 1, 1, .MaxRows - 1, true, .MaxRows,  Col + 1, "V")	' 합계 

			' 차감 계산 
			Call SetW30_W31(Row)
			
			Call SetAllRecalc
		Case C_W18
			Call vspdData_Change(C_W19, Row)
	End Select
	
	End With
	
End Sub

Sub SetW30_W31(Row)
	Dim dblW19, dblW22, dblW25, dblW28, dblW20, dblW23, dblW26, dblW29
	
	With Frm1.vspdData
		.Row = Row
		.Col = C_W19	: dblW19 = UNICDbl(.Value)
		.Col = C_W22	: dblW22 = UNICDbl(.Value)
		.Col = C_W25	: dblW25 = UNICDbl(.Value)
		.Col = C_W28	: dblW28 = UNICDbl(.Value)
		.Col = C_W30	: .Value = dblW19 - dblW22 - dblW25 - dblW28

		.Col = C_W20	: dblW20 = UNICDbl(.Value)
		.Col = C_W23	: dblW23 = UNICDbl(.Value)
		.Col = C_W26	: dblW26 = UNICDbl(.Value)
		.Col = C_W29	: dblW29 = UNICDbl(.Value)
		.Col = C_W31	: .Value = dblW20 - dblW23 - dblW26 - dblW29
		
		Call FncSumSheet(frm1.vspdData,  C_W30, 1, .MaxRows - 1, true, .MaxRows,  C_W30, "V")	' 합계 
		CAll FncSumSheet(frm1.vspdData,  C_W31, 1, .MaxRows - 1, true, .MaxRows,  C_W31, "V")	' 합계 
		
		'Call SetW1_W5()
	End With
End Sub

Sub SetW1_w5()
	With Frm1.vspdData
		.Row = .MaxRows 
		.Col = C_W30	: frm1.txtW1.value = UNIFormatNumber(UNICDbl(.Value) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
		.Col = C_W31	: frm1.txtW5.value = UNIFormatNumber(UNICDbl(.Value) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	End With
End Sub

Sub SetAllW30_W31()
	Dim dblW19, dblW22, dblW25, dblW28, dblW20, dblW23, dblW26, dblW29, iRow, iMaxRows
	
	With Frm1.vspdData
		iMaxRows = .MaxRows - 1
		For iRow = 1 To iMaxRows
		
			.Row = iRow
			.Col = C_W19	: dblW19 = UNICDbl(.Value)
			.Col = C_W22	: dblW22 = UNICDbl(.Value)
			.Col = C_W25	: dblW25 = UNICDbl(.Value)
			.Col = C_W28	: dblW28 = UNICDbl(.Value)
			.Col = C_W30	: .Value = dblW19 - dblW22 - dblW25 - dblW28

			.Col = C_W20	: dblW20 = UNICDbl(.Value)
			.Col = C_W23	: dblW23 = UNICDbl(.Value)
			.Col = C_W26	: dblW26 = UNICDbl(.Value)
			.Col = C_W29	: dblW29 = UNICDbl(.Value)
			.Col = C_W31	: .Value = dblW20 - dblW23 - dblW26 - dblW29
		Next
		
		Call FncSumSheet(frm1.vspdData,  C_W30, 1, .MaxRows - 1, true, .MaxRows,  C_W30, "V")	' 합계 
		CAll FncSumSheet(frm1.vspdData,  C_W31, 1, .MaxRows - 1, true, .MaxRows,  C_W31, "V")	' 합계 

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

	gCurrGrid = 1
	ggoSpread.Source = Frm1.vspdData
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


Sub vspdData2_MouseDown(Button , Shift , x , y )
	gCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

End Sub

'============================================  툴바지원 함수  ====================================

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
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If 

	' 검증작업 
	'1. (31)금액의 합계 < 0 이면 오류(WC0013)
	With frm1.vspdData
		.Row = .MaxRows : .Col = C_W30
		If .Value < 0 Then
			Call DisplayMsgBox("WC0013", parent.VB_INFORMATION, "(30)지급이자 (19-22-25-28)", "X")                          
			Exit Function
		End If
		
		: .Col = C_W31
		If .Value < 0 Then
			Call DisplayMsgBox("WC0013", parent.VB_INFORMATION, "(31)차입금적수 (20-23-26-29)", "X")                          
			Exit Function
		End If 
	End With	
	'--------------
	'add logic 20060201 by Hjo
	'-------------
	Call  fncChkRd() 
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

    Call SetToolbar("1110110100001111")

	frm1.txtCO_CD.focus

    FncNew = True

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

			.vspdData.Col = C_W21
			.vspdData.Text = ""
    
			.vspdData.Col = C_W22
			.vspdData.Text = ""
			
			.vspdData.Col = C_W23
			.vspdData.Text = ""
			
			.vspdData.Col = C_W25
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
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
   
	With frm1.vspdData

		.focus
		ggoSpread.Source = frm1.vspdData
		
		'.vspdData.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
		iRow = .ActiveRow
		
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow ,1
			Call SetSpreadColor(iRow, iRow) 
			.Col = C_SEQ_NO : .Row = iRow	: .Text = iRow	
		
			iRow = 2
			ggoSpread.InsertRow ,1
			Call SetSpreadColor(iRow, iRow) 
			.Col = C_SEQ_NO : .Row = iRow	: .Text = SUM_SEQ_NO	
			.Col = C_W18	: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
					
			ggoSpread.SpreadLock C_W18, iRow, C_W31, iRow
						
		Else
			
			If iRow = .MaxRows Then
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor iRow, iRow + imRow - 1
				MaxSpreadVal frm1.vspdData, C_SEQ_NO, iRow	
				.vspdData.Col = C_SEQ_NO : .vspdData.Row = Row : iSeqNo = .vspdData.Value
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor iRow+1, iRow+1
				MaxSpreadVal frm1.vspdData, C_SEQ_NO, iRow+1
				.vspdData.Col = C_SEQ_NO : .vspdData.Row = Row+1 : iSeqNo = .vspdData.Value
			End If   
		End If
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Sub SetSpreadTotalLine()
	With frm1.vspdData
		.ReDraw = False
		.Row = .MaxRows
		.Col = C_W18	: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
						
		ggoSpread.SpreadLock C_W18, .MaxRows, C_W31, .MaxRows
		.ReDraw = True
	End With
End Sub

Function FncDeleteRow() 
    Dim lDelRows

	If gCurrGrid = 1	Then
		With frm1.vspdData 
			.focus
			ggoSpread.Source = frm1.vspdData 
			lDelRows = ggoSpread.DeleteRow
		End With
    Else
		With frm1.vspdData2 
			.focus
			ggoSpread.Source = frm1.vspdData2
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
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData
	
	If frm1.vspdData.MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE

		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg = "N" Then
			ggoSpread.SpreadUnLock	C_W18, -1, C_W31
			Call SetSpreadLock()
			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1111111100011111")										<%'버튼 툴바 제어 %>
		Else
		
			ggoSpread.SpreadLock	C_W18, -1, C_W31
			Call SetToolbar("1110000000011111")										<%'버튼 툴바 제어 %>
		End If
	
		Call SetSpreadTotalLine ' - 합계라인 재구성 
		
		'Call SetToolbar("1111111100111111")										<%'버튼 툴바 제어 %>
	Else
		Call SetToolbar("1110110100111111")										<%'버튼 툴바 제어 %>
	End If
	
	frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel, lMaxRows, lMaxCols , lCol
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    ' -- html은 그냥 넘어간다 
    
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
		 .Col = 0
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With
	
	With frm1
       .txtSpread.value      = strDel & strVal
		.txtMode.value        =  Parent.UID_M0002
		.txtMaxRows.value     = lGrpCnt-1 
		.txtSpread2.value      = strDel & strVal
		.txtFlgMode.value     = lgIntFlgMode
		
	End With

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
'========================================================================================
Function fncChkRd()
    Dim IntRetCD , i
    Dim sFiscYear, sRepType, sCoCd
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value 
    
  Call CommonQueryRs("  count(*)   ","  tb_26ah "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   If lgF0>0 Then
			fncChkRd=True
			lgIntFlgMode=parent.OPMD_UMODE 
	Else
			fncChkRd=False
			lgIntFlgMode=parent.OPMD_CMODE 
	End IF

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
									<TD CLASS="TD6"><script language =javascript src='./js/w3125ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
				<TD WIDTH=100% HEIGHT=* valign=top>
					<DIV ID="ViewDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto">
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
									   <TR>
										   <TD width="100%" COLSPAN=9 CLASS="CLSFLD"><br>&nbsp;1. 타법인주식 등에 관련한 차입금지급이자</TD>
									   </TR>
									   <TR>
										   <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(1)지급이자</TD>
										   <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=7>적&nbsp;&nbsp;&nbsp;&nbsp;수</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(9)손금불산입<br>지급이자<br>[(1)*(8)/(5)]</TD>
									   </TR>
									   <TR>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(2)타법인주식</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(3)기타</TD>
								           <TD CLASS="TD51" width="10%" ALIGN=CENTER>(4)계</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(5)차입금</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(6)자기자본<br>* (1,2,4,15)</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(7)차감</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(8)(4)와(7)중<br>적은금액</TD>
									  </TR>
									  <TR>
											<TD><script language =javascript src='./js/w3125ma1_txtW1_txtW1.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW2_txtW2.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW3_txtW3.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW4_txtW4.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW5_txtW5.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW6_txtW6.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW7_txtW7.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW8_txtW8.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW9_txtW9.js'></script></TD>
									  </TR>
								  </table>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%" COLSPAN=8 CLASS="CLSFLD"><br>&nbsp;2. 업무무관부동산등에 관련한 차입금지급이자</TD>
									   </TR>
									   <TR>
										   <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(10)지급이자</TD>
										   <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=4>적&nbsp;&nbsp;&nbsp;&nbsp;수</TD>
										   <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(15)차입금<BR>[(31)-(8)]</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(16)(14)와 (15)중<BR>적은금액</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(17)손금불산입지금이자<br>[(10)*(16)/15)]</TD>
									   </TR>
									   <TR>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(11)업무무관부동산</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(12)업무무관동산</TD>
								           <TD CLASS="TD51" width="10%" ALIGN=CENTER>(13)가지급금등</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(14)계<BR>[(11)+(12)+(13)]</TD>
									  </TR>
									  <TR>
											<TD><script language =javascript src='./js/w3125ma1_txtW10_txtW10.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW11_txtW11.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW12_txtW12.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW13_txtW13.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW14_txtW14.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW15_txtW15.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW16_txtW16.js'></script></TD>
											<TD><script language =javascript src='./js/w3125ma1_txtW17_txtW17.js'></script></TD>
									  </TR>
								  </table>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;3. 지급이자 및 차입금 적수계산</TD>
									   </TR>
									   <TR>
										   <TD width="100%"><script language =javascript src='./js/w3125ma1_vaSpread1_vspdData.js'></script>
										  </TD>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>지급이자 및 차입금 적수계산</LABEL>&nbsp;</TD>
				            
				            
				                                 
	
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
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

