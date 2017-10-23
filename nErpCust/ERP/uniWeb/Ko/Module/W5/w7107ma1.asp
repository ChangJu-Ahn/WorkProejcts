<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 소득금액조정 
'*  3. Program ID           : w7107mA1
'*  4. Program Name         : w7107mA1.asp
'*  5. Program Desc         : 이월결손금 등록 
'*  6. Modified date(First) : 2005/02/18
'*  7. Modified date(Last)  : 2005/02/18
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "w7107mA1"
Const BIZ_PGM_ID		= "w7107mB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W7107MB2.asp"

' -- 그리드 컬럼 정의 
Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10
Dim C_W11

Dim IsOpenPop  
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgFISC_START_DT, lgFISC_END_DT, lgFISC_CALC_DT

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO	= 1
	C_W1		= 2
	C_W2		= 3
	C_W3		= 4
	C_W4		= 5
	C_W5		= 6
	C_W6		= 7
	C_W7		= 8	
	C_W8		= 9
	C_W9		= 10
	C_W10		= 11
	C_W11		= 12
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
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

Function OpenAccount()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_WORK_6"					<%' TABLE 명칭 %>
	

	frm1.vspdData.Col = C_W1
    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>

	arrParam(3) = ""							<%' Name Cindition%>
'	arrParam(4) = " ACCT_CD IN (SELECT ACCT_CD FROM TB_ACCT_MATCH (NOLOCK) WHERE MATCH_CD = '18')"							<%' Where Condition%>
	arrParam(4) = " "							<%' Where Condition%>
	arrParam(5) = "계정"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "ACCT_CD"					<%' Field명(0)%>
    arrField(1) = "ACCT_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "계정코드"					<%' Header명(0)%>
    arrHeader(1) = "계정명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAccount(arrRet)
	End If	
	
End Function

Function SetAccount(byval arrRet)
    With frm1
		.vspdData.Col = C_W1
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_W1_NM
		.vspdData.Text = arrRet(1)
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
		lgBlnFlgChgValue = True
	End With
End Function


Sub InitSpreadSheet()
	Dim ret, iRow
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
	
	' 1번 그리드 

	With Frm1.vspdData
				
		ggoSpread.Source = Frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20041222_1",,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W11 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
 
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'헤더를 2줄로    
	    .ColHeaderRows = 2

	    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetDate		C_W1,		"(1)일자",			10, 2, parent.gDateFormat
	    ggoSpread.SSSetFloat	C_W2,		"(2)발생액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W3,		"(3)소급공제",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W4,		"(4)차감계",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W5,		"(5)기공제액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W6,		"(6)당기공제",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W7,		"(7)보전",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W8,		"(8)계",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W9,		"(9)기한내",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W10,		"(10)기한경과",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W11,		"(11)계",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 


	    ret = .AddCellSpan(0, -1000, 1, 2)
	    ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)
	    ret = .AddCellSpan(C_W1, -1000, 1, 2)
	    ret = .AddCellSpan(C_W2, -1000, 3, 1)
	    ret = .AddCellSpan(C_W5, -1000, 4, 1)
	    ret = .AddCellSpan(C_W9, -1000, 3, 1)
	    
	    ' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W2
		.Text = "이월결손금"
		.Col = C_W5
		.Text = "감소내역"
		.Col = C_W9
		.Text = "잔액"
	
		' 두번째 헤더 출력 글자 
		.Row = -999
		.Col = C_W2
		.Text = "(2)발생액"
		.Col = C_W3
		.Text = "(3)소급공제"
		.Col = C_W4
		.Text = "(4)차감계"
		.Col = C_W5
		.Text = "(5)기공제액"
		.Col = C_W6
		.Text = "(6)당기공제"
		.Col = C_W7
		.Text = "(7)보전"
		.Col = C_W8
		.Text = "(8)계"
		.Col = C_W9
		.Text = "(9)기한내"
		.Col = C_W10
		.Text = "(10)기한경과"
		.Col = C_W11
		.Text = "(11)계"

		.rowheight(-999) = 20	' 높이 재지정 
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		
		Call InitSpreadComboBox

		.ReDraw = true	

		Call SetSpreadLock()
				
	End With 
	
	
					
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

		ggoSpread.SpreadUnLock C_W1, -1, C_W11	' 전체 적용 
		ggoSpread.SpreadLock C_W4, -1, C_W4
		ggoSpread.SpreadLock C_W8, -1, C_W8
		ggoSpread.SpreadLock C_W11, -1, C_W11

		ggoSpread.SSSetRequired C_W1, -1, -1
		
	End With	
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow, sITEM_CD

	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
		For iRow = pvStartRow To pvEndRow
			.Col = C_SEQ_NO
			.Row = iRow
			If .Text = "999999" Then
				ggoSpread.SpreadLock C_W1,   iRow, C_W11, iRow

				.Col = C_W1	:	.CellType = 1	:	.Text = "합계"	:	.TypeHAlign = 2
			Else
				ggoSpread.SpreadUnLock C_W1, iRow, C_W11, iRow	' 전체 적용 
				ggoSpread.SSSetRequired C_W1, iRow, iRow
'				ggoSpread.SSSetRequired C_W2, iRow, iRow
				ggoSpread.SpreadLock C_W4, iRow, C_W4, iRow
				ggoSpread.SpreadLock C_W8, iRow, C_W8, iRow
				ggoSpread.SpreadLock C_W11, iRow, C_W11, iRow
			End If
		Next
			
	End With	
End Sub

' -- 헤더쪽 그리드 재조정 
Sub RedrawSumRow()
	Dim iRow
	
	iRow = 1
	
	ggoSpread.Source = Frm1.vspdData
	With Frm1.vspdData
		ggoSpread.SpreadUnLock C_W1, iRow, C_W7, .MaxRows - 1	' 전체 적용 
		ggoSpread.SSSetRequired C_W1, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W2, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W3, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W6, iRow, .MaxRows - 1

		ggoSpread.SpreadLock C_W1,   .MaxRows, C_W7, .MaxRows

		.Row = .MaxRows
		Call .AddCellSpan(C_W1, .MaxRows, 3, 1) 
		.Col = C_W1	:	.CellType = 1	:	.Text = "합계"	:	.TypeHAlign = 2
		.Col = C_W6_NM	:	.CellType = 1
		
		For iRow = 1 to .MaxRows - 1
			.Row = iRow	: .Col = C_W6
			If .Text = "20" Then
				ggoSpread.SpreadLock C_W4, iRow, C_W5, iRow
			Else
				ggoSpread.SSSetRequired C_W4, iRow, iRow
				ggoSpread.SSSetRequired C_W5, iRow, iRow
			End If
		Next
	End With	
End Sub


'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub



Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO		= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W2		= iCurColumnPos(3)
            C_W3		= iCurColumnPos(4)
            C_W4		= iCurColumnPos(5)
            C_W5		= iCurColumnPos(6)
            C_W6		= iCurColumnPos(7)
            C_W7		= iCurColumnPos(8)
            C_W4		= iCurColumnPos(9)
            C_W5		= iCurColumnPos(10)
            C_W6		= iCurColumnPos(11)
            C_W7		= iCurColumnPos(12)
    End Select    
End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
			
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
	lgBlnFlgChgValue = True
    '-----------------------
    'Reset variables area
    '-----------------------
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
	    lgIntFlgMode = parent.OPMD_CMODE
		

		Call SetToolbar("1110111100001111")										<%'버튼 툴바 제어 %>
		Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
'		Call RedrawSumRow
		Call Fn_GridCalc(0, 0)
		Call ChangeRowFlg(frm1.vspdData)
		
	End If

	frm1.vspdData.focus			
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

	call CommonQueryRs("reference"," b_configuration ","  major_cd = 'W2011' and minor_cd = '1' and seq_no = '1' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		lgFISC_CALC_DT = DateAdd("YYYY", Cint(lgF0) * -1, lgFISC_END_DT)
	Else
		lgFISC_CALC_DT = ""
	End if

End Sub

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110111100001111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
 
	Call InitComboBox	' 먼저해야 한다. 기업의 회계기준일을 읽어오기 위해 
	Call ggoOper.ClearField(Document, "1")	
	Call InitData

	
    
    
End Sub


'============================================  사용자 함수  ====================================
Function Fn_GridCalc(ByVal pCol, ByVal pRow)
	Dim iRow, dblSum, sW1, iNum, sW1_0
	Dim dblW2, dblW3, dblW4, dblW5, dblW6, dblW7, dblW8, dblW9, dblW10, dblW11
	
	If Frm1.vspdData.MaxRows <= 0 Then Exit Function

    ggoSpread.Source = Frm1.vspdData
    iRow = pRow

	With Frm1.vspdData
		' (4) = (2) - (3)
		If pRow = 0 Then iRow = Frm1.vspdData.ActiveRow
		
		' 윗열 사업연도 > 아래행 사업연도 오류 메세지 WC0033 (1)사업연도는 과거에서 최근순으로 입력하여야 합니다.
		If iRow > 1 And iRow < .MaxRows Then
			.Row = iRow - 1	:	.Col = C_W1
			sW1_0 = CDate(.Text)
			.Row = iRow	:	.Col = C_W1
			sW1 = CDate(.Text)
			If DateDiff("d", sW1_0, sW1) < 0 Then
				Call DisplayMsgBox("WC0016", "X", "X", "X")
			    .Col = pCol	:	.Text = ""
			    Exit Function
			End If
		End If

		.Row = iRow	:	.Col = C_W2	:	dblW2 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W3	:	dblW3 = UNICdbl(.Text)
		
		' (2) < (3) 오류메세지 WC0010 (3)소급공제금액은 (2)발생액보다 같거나 작아야 합니다.
		If dblW2 < dblW3 Then
		    Call DisplayMsgBox("WC0010", "X", "(3)소급공제금액", "(2)발생액")
		    .Col = pCol	:	.Text = 0
		    Exit Function
		Else
			.Row = iRow	:	.Col = C_W4	:	dblW4 = dblW2 - dblW3	:	.Text = dblW4
		End If
			
		' (8)계 = (5)+(6)+(7)
		.Row = iRow	:	.Col = C_W5	:	dblW5 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W6	:	dblW6 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W7	:	dblW7 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W8	:	dblW8 = dblW5 + dblW6 + dblW7	:	.Text = dblW8

		' (4) < (8) 오류메세지 WC0010 (8)감소내역 계는 (4)차감계보다 같거나 작아야 합니다.
		If dblW4 < dblW8 Then
		    Call DisplayMsgBox("WC0010", "X", "(8)감소내역 계", "(4)차감계")
		    .Col = pCol	:	.Text = 0
		    Exit Function
		End If

	
		.Col = C_W1
		If .Text <> "" And pCol <> C_W9 And pCol <> C_W10 Then
			sW1 = CDate(.Text)
			If DateDiff("d", lgFISC_CALC_DT, sW1) >= 0 Then
				'(9) = (4)-(8)
				.Row = iRow	:	.Col = C_W9	:	dblW9 = dblW4 - dblW8	:	.Text = dblW9
				.Row = iRow	:	.Col = C_W10	:	.Text = 0
			Else
				'(10) = (4)-(8)
				.Row = iRow	:	.Col = C_W9	:	.Text = 0
				.Row = iRow	:	.Col = C_W10	:	dblW10 = dblW4 - dblW8	:	.Text = dblW10
			End If
		End If

		' (11)계 = (9)+(10)
		.Row = iRow	:	.Col = C_W9		:	dblW9 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W10	:	dblW10 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W11	:	dblW11 = dblW9 + dblW10	:	.Text = dblW11

	End With

	dblSum = FncSumSheet(Frm1.vspdData, C_W2, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W2, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W3, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W3, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W4, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W4, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W5, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W5, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W6, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W6, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W7, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W7, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W8, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W8, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W9, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W9, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W10, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W10, "V")	' 합계 
	dblSum = FncSumSheet(Frm1.vspdData, C_W11, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W11, "V")	' 합계 

End Function


'============================================  이벤트 함수  ====================================
'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIdx, sW6

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

	Call Fn_GridCalc(Col, Row)    
    ggoSpread.UpdateRow Frm1.vspdData.MaxRows

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

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
    Call GetSpreadColumnPos("A")
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

		    Call OpenAccount()
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
	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
    
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

    Call SetToolbar("1110111100001111")

     
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
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange <> False Then
		blnChange = True
'	    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
'	    Exit Function
	End If

    If lgBlnFlgChgValue = False and blnChange = True Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

	
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    

    'If Verification = False Then Exit Function
    
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

    Call SetToolbar("1110111100001111")
    lgIntFlgMode = parent.OPMD_CMODE

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

		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
		
			ggoSpread.CopyRow

			.Row = .ActiveRow
			.Col = C_W1
			.Text = ""
			
  
			Call SetDefaultVal(iActiveRow + 1, 1)
			SetSpreadColor iActiveRow, iActiveRow + 1
			.ReDraw = True
			
			Call Fn_GridCalc(0, 0)
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   

End Function

Function FncCancel() 
    Dim lDelRows, dblSum 

		With Frm1.vspdData
			.focus

			ggoSpread.Source = Frm1.vspdData
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
				MsgBox "합계는 삭제할 수 없습니다.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.EditUndo
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
				If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
			End If
			
		End With

		Call Fn_GridCalc(0, 0)    
		dblSum = FncSumSheet(Frm1.vspdData, C_W3, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W3, "V")	' 합계 


End Function

' -- 합계 행인지 체크(Header Grid)
Function CheckTotRow(Byref pObj, Byval pRow) 
	CheckTotRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = 999999 Then	 ' 합계 행 
		CheckTotRow = True
	End If
End Function


' -- Detail Data가 존재하는지 체크 
Function CheckLastRow(Byref pObj, Byval pRow) 
	Dim iCnt, iRow, iMaxRow, iTmpRow
	CheckLastRow = 0
	iCnt = 0
	
	With pObj

		For iRow = 1 To .MaxRows
			.Row = iRow : .Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				iCnt = iCnt + 1
				iMaxRow = iRow
			End If
			.Col = C_SEQ_NO
			If .Text = 999999 Then
				iTmpRow = iRow
			End If
		Next
		.Col = C_SEQ_NO	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow = iMaxRow
		ElseIf iCnt = 1 Then
			CheckLastRow = iTmpRow
		End If
	End With
	
End Function



Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow

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

		if .MaxRows = 0 then
		
			ggoSpread.InsertRow  imRow 
			.Col	= C_SEQ_NO	:	.Text	= 1
			SetSpreadColor 1, 1
			
			ggoSpread.InsertRow  imRow 
			.Row = .MaxRows
			.Col	= C_SEQ_NO	:	.Text	= 999999
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "계"	:	.TypeHAlign = 2
			SetSpreadColor .MaxRows, .MaxRows
			.Row  = 1
			.ActiveRow = 1

		else
			iRow = .ActiveRow

			If iRow = .MaxRows Then	' -- 마지막 합계줄에서 InsertRow를 하면 상위에 추가한다.
				iRow = iRow - 1
				ggoSpread.InsertRow iRow, imRow 
				iRow = iRow + 1

				Call SetDefaultVal(iRow, imRow)
				SetSpreadColor iRow, iRow + imRow
			Else
				ggoSpread.InsertRow ,imRow

				Call SetDefaultVal(iRow + 1, imRow)
				SetSpreadColor iRow, iRow + imRow
			End If   
			.vspdData.Row  = iRow + 1
			.vspdData.ActiveRow = iRow +1
			
        End if 	
		
    End With

    Call SetToolbar("1111111100101111")

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
				.Row = i
				.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
			Next
		End If
	End With
End Function



Function FncDeleteRow() 
    Dim lDelRows, iActiveRow, dblSum 

	With Frm1.vspdData
		.focus

		ggoSpread.Source = Frm1.vspdData
		If .MaxRows <= 0 Then
			Exit Function
		ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "합계는 삭제할 수 없습니다.", vbCritical
			Exit Function
		Else
			lDelRows = ggoSpread.DeleteRow
			lgBlnFlgChgValue = True
			lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
			If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
		End If
		
	End With

	Call Fn_GridCalc(0, 0)    
	dblSum = FncSumSheet(Frm1.vspdData, C_W3, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W3, "V")	' 합계 
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


	    strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey             '☜: Next key tag
	    strVal = strVal     & "&txtMaxRows="		& Frm1.vspdData.MaxRows         '☜: Max fetched data

		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
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
		

		Call SetToolbar("1111111100101111")										<%'버튼 툴바 제어 %>
		Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
'		Call RedrawSumRow
		
	End If
	
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

	Frm1.txtSpread.value      = strVal
	strVal = ""

	Frm1.txtMode.value        =  Parent.UID_M0002
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w7107ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
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
							     <script language =javascript src='./js/w7107ma1_vspdData_vspdData.js'></script>
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
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                <TD WIDTH=50%>&nbsp;</TD>   
                <TD WIDTH=50%>
                </TD>   
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
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
