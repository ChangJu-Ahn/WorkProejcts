<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 소득금액조정 
'*  3. Program ID           : w5109mA1
'*  4. Program Name         : w5109mA1.asp
'*  5. Program Desc         : 제22호 기부금 명세서 
'*  6. Modified date(First) : 2005/02/16
'*  7. Modified date(Last)  : 2006/02/08
'*  8. Modifier (First)     : LSHSAT
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

Const BIZ_MNU_ID		= "w5109mA1"
Const BIZ_PGM_ID		= "w5109mB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W5109MB2.asp"

DIM EBR_RPT_ID	 

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
Dim C_W_DESC

Dim C_W9_CD
Dim C_W9_AMT
Dim C_W9_DESC


Dim IsOpenPop  
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgFISC_START_DT, lgFISC_END_DT

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
	C_W_DESC	= 10
	
	C_W9_CD		= 2
	C_W9_AMT	= 3
	C_W9_DESC	= 4
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
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1018', '" & C_REVISION_YM & "') ","  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
	
End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "dbo.ufn_TB_MINOR('W1008', '" & C_REVISION_YM & "')", "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W2
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W1
	End If

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
		ggoSpread.Spreadinit "V20041222_1",,parent.gForbidDragDropSpread 
    
		.ReDraw = false

		.MaxCols = C_W_DESC + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
 
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'헤더를 2줄로    
	    .ColHeaderRows = 2

		ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetCombo	C_W1,		"(1)유형",		10
	    ggoSpread.SSSetCombo	C_W2,		"(2)코드",		6
		ggoSpread.SSSetEdit		C_W3,		"(3)과목",			20,,,50,1
'	    ggoSpread.SSSetDate		C_W4,		"(4)연월",			10, 2, parent.gDateFormat
	    ggoSpread.SSSetMask		C_W4,		"(4)연월",			7, 2, "9999-99"
		ggoSpread.SSSetEdit		C_W5,		"(5)적요",			15,,,100,1
		ggoSpread.SSSetEdit		C_W6,		"(6)법인명등",			10,,,30,1
		'ggoSpread.SSSetEdit		C_W7,		"(7)사업자등록번호등",	10,,,14,1
		ggoSpread.SSSetMask		C_W7,		"(7)사업자등록번호"	, 15, 2,"999-99-99999"
	    ggoSpread.SSSetFloat	C_W8,		"(8)금액",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetEdit		C_W_DESC,	"비고",			18,,,100,1

	    ret = .AddCellSpan(0, -1000, 1, 2)
	    ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)
	    ret = .AddCellSpan(C_W1, -1000, 2, 1)
	    ret = .AddCellSpan(C_W3, -1000, 1, 2)
	    ret = .AddCellSpan(C_W4, -1000, 1, 2)
	    ret = .AddCellSpan(C_W5, -1000, 1, 2)
	    ret = .AddCellSpan(C_W6, -1000, 2, 1)
	    ret = .AddCellSpan(C_W8, -1000, 1, 2)
	    ret = .AddCellSpan(C_W_DESC, -1000, 1, 2) 
	    
	    ' 첫번째 헤더 출력 글자 
		.Row = -1000
		.Col = C_W1
		.Text = "구분"
	
		.Col = C_W6
		.Text = "기부처"
	
		' 두번째 헤더 출력 글자 
		.Row = -999
		.Col = C_W1
		.Text = "(1)유형"
		.Col = C_W2
		.Text = "(2)코드"
		.Col = C_W6
		.Text = "(6)법인명등"
		.Col = C_W7
		.Text = "(7)사업자등록번호등"
		.rowheight(-999) = 20	' 높이 재지정 
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		
'		Call InitSpreadComboBox

		.ReDraw = true	

'		Call SetSpreadLock()
				
	End With 
	

	' 2번 그리드 

	With Frm1.vspdData2
				
		ggoSpread.Source = Frm1.vspdData2
		'patch version
		ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread 
    
		.ReDraw = false

		.MaxCols = C_W9_DESC + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
 
		ggoSpread.ClearSpreadData
		.MaxRows = 8

		'헤더를 2줄로    
	    '  .ColHeaderRows = 0
	    .ColHeadersShow = False
	    .RowHeaderCols = 2
	    .RowHeadersShow = True

		ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_W9_CD,	"",			5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetFloat	C_W9_AMT,	"",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetEdit		C_W9_DESC,	"",			18,,,100,1

		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W9_CD,C_W9_CD,True)

		Call InitGrid2Header()		
		Call InitSpreadComboBox

		.ReDraw = true	

'		Call SetSpreadLock()
				
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

		ggoSpread.SpreadUnLock C_W1, -1, C_W_DESC ' 전체 적용 
		ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO	' 전체 적용 

		ggoSpread.SSSetRequired C_W1, -1, -1
		ggoSpread.SSSetRequired C_W8, -1, -1
		
	End With	

	With Frm1.vspdData2
	
		ggoSpread.Source = Frm1.vspdData2

		ggoSpread.SpreadUnLock C_W9_CD, -1, C_W9_DESC ' 전체 적용 
		ggoSpread.SpreadLock C_W9_AMT, -1, C_W9_AMT	' 전체 적용 

	End With	
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow

	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
'		ggoSpread.SpreadUnLock C_W1, pvStartRow, C_W_DESC, pvEndRow ' 전체 적용 
		
		ggoSpread.SpreadLock C_SEQ_NO,   pvStartRow, C_SEQ_NO, pvEndRow
		ggoSpread.SpreadLock C_W2,   pvStartRow, C_W2, pvEndRow
		ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
'		ggoSpread.SSSetRequired C_W3, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W4, pvStartRow, pvEndRow
'		ggoSpread.SSSetRequired C_W5, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W6, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W7, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W8, pvStartRow, pvEndRow

		If pvStartRow = 1 Then
			.Col = C_W2 : .Row = 1
'			If .Text = "20" Then
'				ggoSpread.SpreadLock C_W1,   1, C_W2, 1
'				ggoSpread.SpreadLock C_W6,   1, C_W7, 1
'				ggoSpread.SpreadUnLock C_W3, 1, C_W5, 1 ' 전체 적용 
'			End If
		End If

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
		ggoSpread.SSSetRequired C_W1_NM, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W2, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W3, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W6, iRow, .MaxRows - 1
		ggoSpread.SSSetRequired C_W6_NM, iRow, .MaxRows - 1

		ggoSpread.SpreadLock C_W1,   .MaxRows, C_W7, .MaxRows

		.Row = .MaxRows
		Call .AddCellSpan(C_W1, .MaxRows, 3, 1) 
		.Col = C_W1	:	.CellType = 1	:	.Text = "합계"	:	.TypeHAlign = 2
		.Col = C_W6_NM	:	.CellType = 1
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
            C_W8		= iCurColumnPos(9)
            C_W_DESC	= iCurColumnPos(10)
    End Select    
End Sub


Sub InitGrid2Header()
	Dim ret
	
	With Frm1.vspdData2
		ggoSpread.Source = Frm1.vspdData2
		'patch version
    
		.ReDraw = false

		If .MaxRows < 8 Then
			.MaxRows = 8
		End If

	    ret = .AddCellSpan(-1000, 1, 1, 7)
	    ret = .AddCellSpan(-1000, 8, 2, 1)
	    
	    ' 첫번째 헤더 출력 글자 
		.Row = 1
		.Col = -1000
		.Text = "(9)소계"
		.ColWidth(-1000) = 15
	
		.Row = 8
		.Text = "계"
		
		' 두번째 헤더 출력 글자		
		.Col = -999
		.Row = 1	:	.TypeHAlign = 0
		.Text = "가. 법 제24조 제2항 기부금(법정기부금, 코드10)"
		.Row = 2	:	.TypeHAlign = 0 : .RowHidden = True		' -- 200603 : 개정서식 적용으로 나. 가 사라짐, 아래 나/다/라 변경 : W5109MA1_HTF에서 20번을 제출안함
		.Text = "나. 조세특례제한법 제76조 기부금(정치자금, 코드20)"
		.Row = 3	:	.TypeHAlign = 0
		.Text = "나.조세특례제한법 제73조 제1항 제1호 기부금(코드60)"
		.Row = 4	:	.TypeHAlign = 0
		.Text = "다.조세특례제한법 제73조 제1항 제2호 내지 제15호 기부금(코드30)"
		.Row = 5	:	.TypeHAlign = 0
		.Text = "라. 법 제24조 제1항 기부금(지정기부금, 코드40)"
		.Row = 6	:	.TypeHAlign = 0
		.Text = "마.조세특례제한법 제73조 제2항 기부금(코드 70)"
		.Row = 7	:	.TypeHAlign = 0
		.Text = "바.기타 기부금(코드 50)"
		.rowheight(-1) = 15	' 높이 재지정 
		.ColWidth(-999) = 76
		
		' 기본코드값입력하기 
		.Col = C_SEQ_NO
		.Row = 1	:	.Text = 1
		.Row = 2	:	.Text = 2
		.Row = 3	:	.Text = 3
		.Row = 4	:	.Text = 4
		.Row = 5	:	.Text = 5
		.Row = 6	:	.Text = 6
		.Row = 7	:	.Text = 7
		.Row = 8	:	.Text = 8
		
		' 기본코드값입력하기 
		.Col = C_W9_CD
		.Row = 1	:	.Text = "10"
		.Row = 2	:	.Text = "20"
		.Row = 3	:	.Text = "60"
		.Row = 4	:	.Text = "30"
		.Row = 5	:	.Text = "40"
		.Row = 6	:	.Text = "70"
		.Row = 7	:	.Text = "50"
		.Row = 8	:	.Text = "99"
		
		.ReDraw = true	

		Call SetSpreadLock()
	End With
End Sub

'============================== 사용자 정의 함수  ========================================
Function SetCalcSumGrid()
	Dim dblC10, dblC20, dblC30, dblC40, dblC50, dblC60, dblC70, iRow
	
	dblC10 = 0
	dblC20 = 0
	dblC30 = 0
	dblC40 = 0
	dblC50 = 0
	dblC60 = 0
	dblC70 = 0

	ggoSpread.Source = Frm1.vspdData
	With Frm1.vspdData
	
		If .MaxRows > 0 Then
	
			For iRow = 1 To .MaxRows
				.Row = iRow	:	.Col = 0
				If .Text <> ggoSpread.DeleteFlag Then
					.Row = iRow	:	.Col = C_W2
					Select Case .Text
						Case "10"
							.Col = C_W8
							dblC10 = dblC10 + UNICDbl(.Text)
						Case "20"
							.Col = C_W8
							'dblC20 = dblC20 + UNICDbl(.Text)	'-- 200603 서식개정으로 제거됨
						Case "30"
							.Col = C_W8
							dblC30 = dblC30 + UNICDbl(.Text)
						Case "40"
							.Col = C_W8
							dblC40 = dblC40 + UNICDbl(.Text)
						Case "50"
							.Col = C_W8
							dblC50 = dblC50 + UNICDbl(.Text)
						Case "60"
							.Col = C_W8
							dblC60 = dblC60 + UNICDbl(.Text)
						Case "70"
							.Col = C_W8
							dblC70 = dblC70 + UNICDbl(.Text)
					End Select
				End If
			Next
		End If
	End With
	
	With Frm1.vspdData2
		.Row = 1	:	.Col = C_W9_AMT	:	.Text = dblC10
		.Row = 2	:	.Col = C_W9_AMT	:	.Text = dblC20
		.Row = 3	:	.Col = C_W9_AMT	:	.Text = dblC60
		.Row = 4	:	.Col = C_W9_AMT	:	.Text = dblC30
		.Row = 5	:	.Col = C_W9_AMT	:	.Text = dblC40
		.Row = 6	:	.Col = C_W9_AMT	:	.Text = dblC70
		.Row = 7	:	.Col = C_W9_AMT	:	.Text = dblC50
		.Row = 8	:	.Col = C_W9_AMT	:	.Text = dblC10 + dblC20 + dblC30 + dblC40 + dblC50 + dblC60 + dblC70
	End With
	
End Function

Function ChkW4Date(ByVal pCol, ByVal pRow)
	Dim iDate 
	
	With Frm1.vspdData
		.Col = pCol
		.Row = pRow
		
		iDate = UNIFormatDate(.Text + "-01")

		If iDate = "" Then
			Call UNIMsgBox("연월 형식이 아닙니다.", vbOKOnly, "알림")
			.Text = ""
			Exit Function
		End If

		If CompareDateByFormat(lgFISC_START_DT, iDate, "당기시작일", "연월" , "970023", parent.gDateFormat,parent.gComDateType,True) = False Then
			.Text = ""
			Exit Function
		End If

		If CompareDateByFormat(iDate, lgFISC_END_DT, "연월" , "당기종료일", "970025", parent.gDateFormat,parent.gComDateType,True) = False Then
			.Text = ""
			Exit Function
		End If

	End With
End Function

Function ChkW1(ByVal pCol, ByVal pRow)
	Dim iRow
	
	ChkW1 = True
	Exit Function	' -- 정치자금 제거됨 200603 개정서식
	
	With Frm1.vspdData
		For iRow = 1 To .MaxRows
			.Col = C_W2	:	.Row = iRow
			If iRow = 1 And .Text <> "20" Then
				Call UNIMsgBox("맨첫줄은 정치자금만 입력가능합니다.", vbOKOnly, "알림")
				.Text = ""
				.Col = C_W1	:	.Text = ""
				ChkW1 = False
				Exit Function
			ElseIf iRow > 1 And .Text = "20" Then
				Call UNIMsgBox("정치자금은 맨첫줄만 입력가능합니다.", vbOKOnly, "알림")
				.Text = ""
				.Col = C_W1	:	.Text = ""
				ChkW1 = False
				Exit Function
			End If
		Next
	End With
End Function
'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
			
	ggoSpread.Source = Frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    Call InitGrid2Header
			
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
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
	    lgIntFlgMode = parent.OPMD_CMODE
		

		Call SetToolbar("1110111100001111")										<%'버튼 툴바 제어 %>
		Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
'		Call RedrawSumRow
		Call ChangeRowFlg(frm1.vspdData)
		Call SetCalcSumGrid
		
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

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

   ' iCalledAspName = AskPRAspName("W5105RA1")
    
   ' If Trim(iCalledAspName) = "" Then
   '     IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
    '    IsOpenPop = False
    '    Exit Function
    'End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog("W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function


Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
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
	Call fncQuery   
    
End Sub


'============================================  이벤트 함수  ====================================
'============================================  이벤트 호출 함수  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIdx

	With Frm1.vspdData
		Select Case Col
			Case C_W1
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
			Case C_W2
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
    
    If Col = C_W1 Or Col = C_W8 Then
    	If ChkW1(Col, Row) = True Then Call SetCalcSumGrid
'		dblSum = FncSumSheet(Frm1.vspdData, C_W3, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W3, "V")	' 합계 
	ElseIf Col = C_W4 Then
		Call ChkW4Date(Col, Row)
    End If

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

Sub vspdData2_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	lgBlnFlgChgValue= True ' 변경여부 
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col

    If Frm1.vspdData2.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData2.text) < UNICDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
    End If
	
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
	If ggoSpread.SSCheckChange = False Then
		blnChange = True
	End If

    If lgBlnFlgChgValue = False and blnChange = True Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

	
    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
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
	Dim iDblW8, iRow
	
	Verification = False

	With Frm1.vspdData
		For iRow = 1 To .MaxRows
			.Row = iRow
		    .Col = C_W8 :	iDblW8 = unicdbl(.text)
			
		    '(8) < 0 이면 오류 (메세지 WC0010)
		    If iDblW8 < 0 Then
		        Call DisplayMsgBox("WC0006", "X", "기부금 금액", "X")
			    Exit Function
		    End If
		Next
	End With

	Verification = True	
End Function
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 '========================================================================================================
'	Name : SetPrintCond()
'	Description : Group Condition PopUp
'========================================================================================================
Sub SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE)
	varCo_Cd	 =  Trim(frm1.txtCo_Cd.value)
	varFISC_YEAR = Trim(frm1.txtFISC_YEAR.text)
	varREP_TYPE	 =  Trim(frm1.cboREP_TYPE.value)

End Sub  


'========================================================================================
Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE
	Dim StrUrl  , i

	Dim intCnt,IntRetCD


    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If
    

    Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE)
    Call CommonQueryRs("Count(*)"," TB_22H","CO_CD= '" & varCo_Cd & "' AND FISC_YEAR='" & varFISC_YEAR & "' AND REP_TYPE='" & varREP_TYPE & "'  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    if unicdbl(lgF0) > 18  then
      	 EBR_RPT_ID	    = "w5109OA2"
    else
         EBR_RPT_ID	    = "w5109OA1"
    end if
    
   
    StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
	StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
	StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE

     ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
     if  strPrintType = "VIEW" then
	 Call FncEBRPreview(ObjName, StrUrl)
     else
	 Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	
     
     Dim objChkBox, iCnt
     
     Set objChkBox = document.all("prt_check")
     
     If Not objChkBox is Nothing Then
     
		iCnt = GetEBRCheckBox("prt_check")
		
		for i =0 to  iCnt -1 
                  if document.all("prt_check"&i+1).checked = true then 
		    ObjName = AskEBDocumentName(EBR_RPT_ID & i+1, "ebr")

		      
				if  strPrintType = "VIEW" then
					Call FncEBRPreview(ObjName, StrUrl)
				else
					Call FncEBRPrint(EBAction,ObjName,StrUrl)
				end if	
		     end if	
		Next 

     End If

End Function    

Function GetEBRCheckBox(Byval pObjName)
	Dim oFrm, oNode, iCnt
	Set oFrm = document.frm1
	iCnt = 0
	For Each oNode In oFrm.elements
		If oNode.TagName = "INPUT" Then
			If LCase(oNode.name) = LCase(pObjName) Then
				iCnt = iCnt + 1
			End If
		End If
	Next
	GetEBRCheckBox = iCnt
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
'    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
	Call InitGrid2Header
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
			
			.Col = C_W2
			.Text = ""

			.Col = C_W8
			.Text = ""
    
			.ReDraw = True

			SetSpreadColor .ActiveRow, .ActiveRow
			Call SetDefaultVal(iActiveRow + 1, 1)
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
		Else
			lDelRows = ggoSpread.EditUndo(.ActiveRow)
			lgBlnFlgChgValue = True
		End If
		
		Call SetCalcSumGrid	' 합계 
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

		iRow = .ActiveRow

		If .MaxRows <= 0 Then

			iRow = 0
			ggoSpread.InsertRow ,imRow
	
			'.Row = iRow + 1	:	.Col = C_W1	:	.Text = "정치자금"
			'.Row = iRow + 1	:	.Col = C_W2	:	.Text = "20"
			SetSpreadColor iRow + 1, iRow + imRow
		Else

			ggoSpread.InsertRow ,imRow
			SetSpreadColor iRow + 1, iRow + imRow
	
		End If
		Call SetDefaultVal(iRow + 1, imRow)
'		.ActiveRow = iRow + 1
		
    End With

    Call SetToolbar("1111111100101111")
	
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' 그리드에 SEQ_NO 넣는 로직 
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
		Else
			lDelRows = ggoSpread.DeleteRow
			lgBlnFlgChgValue = True
		End If
		
		Call SetCalcSumGrid	' 합계 
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
'		Call RedrawSumRow
'		Call SetCalcSumGrid
		
	End If
    Call InitGrid2Header
	Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
	
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

	With Frm1.vspdData2

		ggoSpread.Source = Frm1.vspdData2
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
		
		' ----- 2번째 그리드 
		For lRow = 1 To .MaxRows

	       .Row = lRow

	       Select Case lgIntFlgMode
	           Case  parent.OPMD_CMODE                                     '☜: Insert
	                                              strVal = strVal & "C"  &  Parent.gColSep
	           Case  parent.OPMD_UMODE                                      '☜: Update
	                                              strVal = strVal & "U"  &  Parent.gColSep
	       End Select
	       
		  ' 모든 그리드 데이타 보냄     
			For lCol = 1 To lMaxCols
				.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
			Next
			strVal = strVal & Trim(.Text) &  Parent.gRowSep
		Next
	
	End With

	Frm1.txtSpread2.value      = strVal
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
					<TD WIDTH=* align=right><A href="vbscript:GetRef">금액불러오기</A>|<a href="vbscript:OpenRefMenu">소득금액합계표조회</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w5109ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
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
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=208>
							     <script language =javascript src='./js/w5109ma1_vspdData_vspdData.js'></script>
							    </TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/w5109ma1_vspdData2_vspdData2.js'></script>
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
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" style="display:'none'"></TEXTAREA>
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
