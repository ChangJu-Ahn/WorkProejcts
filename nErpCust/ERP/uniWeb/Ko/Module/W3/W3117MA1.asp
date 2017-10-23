
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 가지급금 및 가수금 적수계산 
'*  3. Program ID           : W3117MA1
'*  4. Program Name         : W3117MA1.asp
'*  5. Program Desc         : 가지급금 및 가수금 적수계산 
'*  6. Modified date(First) : 2005/01/20
'*  7. Modified date(Last)  : 2006/01/24
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : HJO
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
' 인사 DB추출 로직 미비.
' 연산문제 : 삭제된 로에 대한 처리 없이 저장됨..
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

Const BIZ_MNU_ID		= "W3117MA1"
Const BIZ_PGM_ID		= "W3117MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W3117MB2.asp"
Const EBR_RPT_ID		= "W3117OA1"

' -- 1번 합계 그리드 
Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W7_NM
Dim C_W8
Dim C_W8_NM

' -- 2,3번 가지급금/가수금 그리드 
Dim C_CHILD_SEQ_NO
Dim C_W9
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgOldCol, lgOldRow , lgChgFlg
Dim lgFISC_START_DT, lgFISC_END_DT, lgRateOver, lgDefaultRate

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	lgCurrGrid	= 1
	lgOldRow	= 0
	lgOldCol	= 2
	lgChgFlg	= False

	'--1번그리드 
	C_SEQ_NO = 1
	C_W1 = 2
	C_W2 = 3
	C_W3 = 4
	C_W4 = 5
	C_W5 = 6
	C_W6 = 7
	C_W7 = 8
	C_W7_NM = 9
	C_W8 = 10
	C_W8_NM = 11

	'--2번 3번그리드 
	C_CHILD_SEQ_NO	= 2
	C_W9		= 3
	C_W10		= 4
	C_W11		= 5
	C_W12		= 6
	C_W13		= 7
	C_W14		= 8
	C_W15		= 9
	C_W16		= 10

	
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
		frm1.vspdData.ScriptEnhanced = True
	   'patch version
	    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	
	    .MaxCols = C_W8_NM + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	    .MaxRows = 0
	    
	    'Call AppendNumberPlace("6","3","2")
	
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번",				5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetCombo	C_W1,		"(1)대여구분",		10
		ggoSpread.SSSetEdit		C_W2,		"(2)성명(법인명)",	15,,,50,1	
		ggoSpread.SSSetFloat	C_W3,		"(3)가지급금 적수",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W4,		"(4)가수금 적수",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W5,		"(5)차감계",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W6,		"(6)이자수익",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetCombo	C_W7,		"(7)소득처분", 		10
	    ggoSpread.SSSetCombo	C_W7_NM,	"(7)소득처분", 		10
	    ggoSpread.SSSetCombo	C_W8,		"(8)인정이자율종류", 10
	    ggoSpread.SSSetCombo	C_W8_NM,	"(8)인정이자율종류", 15
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W7,C_W7,True)
		Call ggoSpread.SSSetColHidden(C_W8,C_W8,True)
						
		.ReDraw = true
	
    End With

 	' -----  2번 그리드 
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_W16 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	   .MaxRows = 0
	 
	    'Call AppendNumberPlace("6","3","2")
	
		ggoSpread.SSSetEdit		C_SEQ_NO,	"부모순번", 5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"자식순번", 5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_W9,		"(9)성명(법인명)",	15,,,50,1	
	    ggoSpread.SSSetDate		C_W10,		"(10)일자",			10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_W11,		"(11)적요",			15,,,50,1
	    ggoSpread.SSSetFloat	C_W12,		"(12)차변금액" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W13,		"(13)대변금액" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W14,		"(14)잔액",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetEdit		C_W15,		"(15)일수",			10,1,,50,1
	    ggoSpread.SSSetFloat	C_W16,		"(16)적수",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_CHILD_SEQ_NO,C_CHILD_SEQ_NO,True)
					
		.ReDraw = true
	
    End With

 	' -----  3번 그리드 
	With frm1.vspdData3
	
		ggoSpread.Source = frm1.vspdData3	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_W16 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols									'☆: 사용자 별 Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	   .MaxRows = 0
	 
	    'Call AppendNumberPlace("6","3","2")
	
		ggoSpread.SSSetEdit		C_SEQ_NO,	"부모순번", 5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"자식순번", 5,,,6,1	' 히든컬럼 
	    ggoSpread.SSSetEdit		C_W9,		"(9)성명(법인명)",	15,,,50,1	
	    ggoSpread.SSSetDate		C_W10,		"(10)일자",			10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_W11,		"(11)적요",			15,,,50,1
	    ggoSpread.SSSetFloat	C_W12,		"(12)차변금액" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W13,		"(13)대변금액" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W14,		"(14)잔액",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetEdit		C_W15,		"(15)일수",			10,1,,50,1
	    ggoSpread.SSSetFloat	C_W16,		"(16)적수",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_CHILD_SEQ_NO,C_CHILD_SEQ_NO,True)
					
		.ReDraw = true
    
    End With
    
	Call InitSpreadComboBox()
    Call SetSpreadLock 
           
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim IntRetCD1

	' 대여구분 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo "주택자금" & vbTab & "기타", C_W1

	' 인정이자 소득처분 
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1060' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W7
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W7_NM
	End If

	' 인정이자율 종류 
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1059' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W8
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W8_NM
	End If

End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    .vspdData2.ReDraw = False
    .vspdData3.ReDraw = False

	' 1번 그리드 
    ggoSpread.Source = frm1.vspdData
        
	ggoSpread.SSSetRequired C_W1, -1, -1
	ggoSpread.SSSetRequired C_W2, -1, -1
'	ggoSpread.SSSetRequired C_W6, -1, -1
	ggoSpread.SSSetRequired C_W7, -1, -1
	ggoSpread.SSSetRequired C_W7_NM, -1, -1
	ggoSpread.SSSetRequired C_W8, -1, -1
	ggoSpread.SSSetRequired C_W8_NM, -1, -1
    ggoSpread.SpreadLock C_W3, -1, C_W3
    ggoSpread.SpreadLock C_W4, -1, C_W4
    ggoSpread.SpreadLock C_W5, -1, C_W5    
    
    ' 2번 그리드 
    ggoSpread.Source = frm1.vspdData2	

    ggoSpread.SpreadLock C_W9, -1, C_W9
	ggoSpread.SSSetRequired C_W10, -1, -1
    ggoSpread.SpreadLock C_W14, -1, C_W14
    ggoSpread.SpreadLock C_W15, -1, C_W15
    ggoSpread.SpreadLock C_W16, -1, C_W16

	' 3번 그리드 
    ggoSpread.Source = frm1.vspdData3	

    ggoSpread.SpreadLock C_W9, -1, C_W9
	ggoSpread.SSSetRequired C_W10, -1, -1
    ggoSpread.SpreadLock C_W14, -1, C_W14
    ggoSpread.SpreadLock C_W15, -1, C_W15
    ggoSpread.SpreadLock C_W16, -1, C_W16

    .vspdData.ReDraw = True
    .vspdData2.ReDraw = True
    .vspdData3.ReDraw = True

    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

	'If lgCurrGrid = 1 Then
		'.vspdData.ReDraw = False
 
		ggoSpread.Source = .vspdData
	
		ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W2, pvStartRow, pvEndRow
'		ggoSpread.SSSetRequired C_W6, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W7, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W7_NM, pvStartRow, pvEndRow
		If lgRateOver Then
			ggoSpread.SSSetRequired C_W8, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W8_NM, pvStartRow, pvEndRow
		Else
			ggoSpread.SpreadLock C_W8, pvEndRow, C_W8
			ggoSpread.SpreadLock C_W8_NM, pvEndRow, C_W8_NM
		End If
	    ggoSpread.SpreadLock C_W3, pvEndRow, C_W3
	    ggoSpread.SpreadLock C_W4, pvEndRow, C_W4
	    ggoSpread.SpreadLock C_W5, pvEndRow, C_W5    
		    
		'.vspdData.ReDraw = True

    'End If
    
    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColorDetail2(ByVal pvEndRow)
    With frm1
    
		' 2번 그리드 
		ggoSpread.Source = frm1.vspdData2	

		ggoSpread.SpreadLock C_W9, pvEndRow, C_W9
		ggoSpread.SSSetRequired C_W10, pvEndRow, pvEndRow
		ggoSpread.SpreadLock C_W14, pvEndRow, C_W14
		ggoSpread.SpreadLock C_W15, pvEndRow, C_W15
		ggoSpread.SpreadLock C_W16, pvEndRow, C_W16
    
    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColorDetail3(ByVal pvEndRow)
    With frm1
    
		' 2번 그리드 
		ggoSpread.Source = frm1.vspdData3	

		ggoSpread.SpreadLock C_W9, pvEndRow, C_W9
		ggoSpread.SSSetRequired C_W10, pvEndRow, pvEndRow
		ggoSpread.SpreadLock C_W14, pvEndRow, C_W14
		ggoSpread.SpreadLock C_W15, pvEndRow, C_W15
		ggoSpread.SpreadLock C_W16, pvEndRow, C_W16
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W7		= iCurColumnPos(2)
            C_W9		= iCurColumnPos(3)
            C_W8		= iCurColumnPos(4)
            C_W8_NM		= iCurColumnPos(5)
            C_W9		= iCurColumnPos(6)
            C_W10		= iCurColumnPos(7)
            C_W11		= iCurColumnPos(8)
            C_W12       = iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W17		= iCurColumnPos(13)
            C_W18		= iCurColumnPos(14)
            C_W19		= iCurColumnPos(15)
            C_W20		= iCurColumnPos(16)
    End Select    
End Sub


'============================== 사용자 정의 함수  ========================================
Sub InsertRow2Head()
	' fncNew, onLoad시에 호출해서 기본적으로 3칸을 입력함 
	Dim ret, iRow, iLoop
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
			
		.ReDraw = False

		iRow = 1
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		.Col = C_SEQ_NO : .Row = iRow: .Text = iRow	
		Call SetDefaultW8(iRow)		' 인정이자율 설정 
		
		iRow = 2
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		.Col = C_SEQ_NO : .Row = iRow: .Text = "999999"	
		Call .AddCellSpan(C_W1, .MaxRows, 2, 1)
		
		.col = C_W1 : .CellType = 1 : .text = "계" : .TypeHAlign = 2
				
		ggoSpread.SpreadLock C_W1, iRow, C_W8_NM, iRow
		
		.ReDraw = True		
		.focus
		'.SetActiveCell 2, 1
					
	End With

'	Call InsertRow2Detail2(1)
'	Call InsertRow2Detail3(1)
	
	Call vspdData_Click(C_W1, 1)
End Sub

Sub InsertRow2Detail2(Byval pSeqNo)

	' 작업진행률 그리드 추가 
	Dim ret, iRow, iLoop, iLastRow
	
	With frm1.vspdData2
		
		.focus
		ggoSpread.Source = frm1.vspdData2

		iLastRow = .MaxRows
		.SetActiveCell C_W9, iLastRow	
		
		.ReDraw = False
		'ggoSpread.ClearSpreadData

		iRow = 1
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = iRow
		.Col = C_SEQ_NO			: .Text = pSeqNo
		Call SetSpreadColorDetail2(iLastRow+iRow) 
		.RowHidden = True

		iRow = 2
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = "999999"
		.Col = C_SEQ_NO			: .Text = pSeqNo
		.Col = C_W9				: .Text = "계"	: .TypeHAlign = 0	
		Call .AddCellSpan(C_W9, .MaxRows, 7, 1)
		Call SetSpreadColorDetail2(iLastRow+iRow) 

		ggoSpread.SpreadLock C_W9, iLastRow+iRow, C_W16, iLastRow+iRow
		.RowHidden = True	
		
		'.vspdData2.SetActiveCell 2, 1	
		.ReDraw = True		

	End With
	
End Sub

Sub InsertRow2Detail3(Byval pSeqNo)

	' 작업진행률 그리드 추가 
	Dim ret, iRow, iLoop, iLastRow
	' 기타수입금액 그리드	
	With frm1.vspdData3
		
		.focus
		ggoSpread.Source = frm1.vspdData3

		iLastRow = .MaxRows
		.SetActiveCell C_W9, iLastRow	
		
		.ReDraw = False
		'ggoSpread.ClearSpreadData

		iRow = 1
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = iRow
		.Col = C_SEQ_NO			: .Text = pSeqNo
		Call SetSpreadColorDetail3(iLastRow+iRow) 
		.RowHidden = True

		iRow = 2
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = "999999"
		.Col = C_SEQ_NO			: .Text = pSeqNo
		.Col = C_W9				: .Text = "계"	: .TypeHAlign = 0	
		Call .AddCellSpan(C_W9, .MaxRows, 7, 1)
		Call SetSpreadColorDetail3(iLastRow+iRow) 

		ggoSpread.SpreadLock C_W9, iLastRow+iRow, C_W16, iLastRow+iRow
		.RowHidden = True	
		
		'.vspdData2.SetActiveCell 2, 1	
		.ReDraw = True		
	
	End With
End Sub

' -- 헤더쪽 그리드 재조정 
Sub RedrawSumRow()
	Dim iRow, iMaxRows, lSeqNo, ret
	
	With frm1.vspdData
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData
		
		For iRow = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = iRow : lSeqNo = .Value
			
			If lSeqNo = 999999 Then ' 합계행 
				ret = .AddCellSpan(C_W1, iRow, 2, 1)	' 합계 행의 계정과목 셀을 합침	
				
				.col = C_W1 : .CellType = 1 : .text = "계" : .TypeHAlign = 2

				ggoSpread.SpreadLock C_W1, iRow, C_W8_NM, iRow

			Else
				Call SetSpreadColor(iRow, iRow)
			End If
		Next
	End With
End Sub

' --  2번째 그리드 합계 재조정 
Sub RedrawSumRow2()
	Dim iRow, iMaxRows, lSeqNo, ret
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData2
		
		For iRow = 1 to iMaxRows
			.Col = C_CHILD_SEQ_NO : .Row = iRow : lSeqNo = .Value
			
			If lSeqNo = 999999 Then ' 합계행 
			
				.Col = C_W9			: .Text = "계"	: .TypeHAlign = 0	
				Call SetSpreadColorDetail2(iRow) 

				ggoSpread.SpreadLock C_W9, iRow, C_W16, iRow
				ret = .AddCellSpan(C_W9, iRow, 7, 1)	' 합계 행의 계정과목 셀을 합침	
				ggoSpread.SpreadLock 1, iRow, C_W16, iRow	
			End If
		Next
	End With
End Sub

' -- 3번째 그리드 합계 재조정 
Sub RedrawSumRow3()
	Dim iRow, iMaxRows, lSeqNo, ret
	
	With frm1.vspdData3
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData3
		
		For iRow = 1 to iMaxRows
			.Col = C_CHILD_SEQ_NO : .Row = iRow : lSeqNo = .Value
			
			If lSeqNo = 999999 Then ' 합계행 
			
				.Col = C_W9			: .Text = "계"	: .TypeHAlign = 0	
				Call SetSpreadColorDetail3(iRow) 

				ret = .AddCellSpan(C_W9, iRow, 7, 1)	' 합계 행의 계정과목 셀을 합침	
				ggoSpread.SpreadLock 1, iRow, C_W16, iRow	
			End If
		Next
	End With
End Sub

' -- 행을 히든 처리 
Function ShowRowHidden(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
	iMaxRows = .MaxRows : iFirstRow = 0
	
	For iRow = 1 To iMaxRows
		.Col = C_SEQ_NO : .Row = iRow : iSeqNo = .Value
		If iSeqNo = pSeqNo Then	' 같은 관계라면..
			.RowHidden = False
			If iFirstRow = 0 Then iFirstRow = iRow
		Else
			.RowHidden = True
		End If
	Next
	
	ShowRowHidden = iFirstRow
	End With
End Function

' -- 합계 행인지 체크 
Function CheckTotalRow(Byref pObj, Byval pRow) 
	CheckTotalRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If UNICDbl(pObj.Text) = 999999 Then	 ' 합계 행 
		CheckTotalRow = True
	End If
End Function

' -- 합계 행인지 체크 
Function CheckTotalRow2(Byref pObj, Byval pRow) 
	CheckTotalRow2 = False
	pObj.Col = C_CHILD_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If UNICDbl(pObj.Text) = 999999 Then	 ' 합계 행 
		CheckTotalRow2 = True
	End If
End Function

' -- Detail Data가 존재하는지 체크 
Function CheckDetailData(Byref pObj, Byref pObjDe, Byval pRow) 
	Dim iSeqNo, iRow
	CheckDetailData = 0
	pObj.Col = C_SEQ_NO : pObj.Row = pRow	:	iSeqNo = Trim(pObj.Text)
	
	With pObjDe
		For iRow = 1 To .MaxRows
			.Row = iRow	:	.Col = C_SEQ_NO
			If Trim(.Text) = iSeqNo Then
				.Col = 0
				If .Text <> ggoSpread.DeleteFlag Then
					CheckDetailData = CheckDetailData + 1
				End If
			End If
		Next
	End With
End Function

' -- 합계이외의 데이타가 있는지 존재하는지 체크 
Function CheckLastRow(Byref pObj, Byval pRow) 
	Dim iCnt, iRow, iMaxRow
	CheckLastRow = 0
	iCnt = 0
	
	With pObj

		For iRow = 1 To .MaxRows
			.Row = iRow : .Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				iCnt = iCnt + 1
				iMaxRow = iRow
			End If
		Next
		.Col = C_SEQ_NO	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow = iMaxRow
		End If
	End With
	
End Function

' -- 합계이외의 데이타가 있는지 존재하는지 체크 
Function CheckLastRow2(Byref pObj, Byval pRow) 
	Dim iCnt, iRow, iMaxRow, iSeqNo, iTmpRow
	CheckLastRow2 = 0
	iCnt = 0
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_SEQ_NO
	iSeqNo = frm1.vspdData.Text
	With pObj

		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_SEQ_NO
			If .Text = iSeqNo Then
				.Col = 0
				If .Text <> ggoSpread.DeleteFlag Then
					iCnt = iCnt + 1
					iMaxRow = iRow
				End If
				.Col = C_CHILD_SEQ_NO
				If .Text = 999999 Then
					iTmpRow = iRow
				End If
			End If
		Next
		.Col = C_CHILD_SEQ_NO	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow2 = iMaxRow
		ElseIf iCnt = 1 Then
			CheckLastRow2 = iTmpRow
		End If
	End With
	
End Function


' ----------- Grid 0 Process
Function Fn_GridCalc(ByVal pCol, ByVal pRow)
	Dim dblSum

	With Frm1.vspdData
		Select Case pCol
			Case C_W2		' 성명(법인명)
				.Col = C_W2	:	.Row = pRow
				Call SetW2ToChildGrid(.Text)	' 현재 이름을 하위 그리드에 넣는다.
		End Select

		' C_W3 : 가지급금 적수 
		dblSum = FncSumSheet(frm1.vspdData, C_W3, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W3 : .Row = .MaxRows : .Value = dblSum	

		' C_W4 : 가수금 적수 
		dblSum = FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W4 : .Row = .MaxRows : .Value = dblSum	

		' C_W5 : 차감계 
		dblSum = FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W5 : .Row = .MaxRows : .Value = dblSum	

		' C_W6 : 이자수익 
		dblSum = FncSumSheet(frm1.vspdData, C_W6, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W6 : .Row = .MaxRows : .Value = dblSum	
	End With
End Function

' ----------- Grid 2 Process
Function Fn_GridCalc2(ByVal pCol, ByVal pRow)
	Dim dblSum

	With Frm1.vspdData2
	    Select Case pCol
			Case 0, C_W10, C_W12, C_W13
				Call SetG2SumValue(C_W12, pRow)	' 차변 컬럼 행합계 
				Call SetG2SumValue(C_W13, pRow)	' 대변 컬럼 행합계 
				Call SetG2W14(pRow)			' 잔액 계산 
				If pRow <> 0 Then
					.Col = C_W14 : .Row = pRow
					If UNICDbl(.Text) < 0 Then
						Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "가지급금(14)잔액", "X")           '⊙: "일자를 순서 맞추기.."	
						Exit Function
					End If
				End If
				Call SetG2W15(pRow)			' 일수 계산 
				Call SetG2W16(pRow)			' 적수 계산 
				'Call SetG2SumValue(C_W14, pRow)	' 잔액 행합계 
				Call SetG2SumValue(C_W15, pRow)	' 일수 행합계 
				Call SetG2SumValue(C_W16, pRow)	' 적수 행합계 
				
				Call SetG2W11(pRow)				' 적요 반영 
				Call SetW3_W4()				' 상단 그리드 반영 
	    End Select
	End With
End Function

' ----------- Grid 3 Process
Function Fn_GridCalc3(ByVal pCol, ByVal pRow)
	Dim dblSum

	With Frm1.vspdData3
	    Select Case pCol
			Case 0, C_W10, C_W12, C_W13
				Call SetG3SumValue(C_W12, pRow)	' 차변 컬럼 행합계 
				Call SetG3SumValue(C_W13, pRow)	' 대변 컬럼 행합계 
				Call SetG3W14(pRow)			' 잔액 계산  
				If pRow <> 0 Then
					.Col = C_W14 : .Row = pRow
					If UNICDbl(.Text) < 0 Then
						Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "가수금(14)잔액", "X")           '⊙: "일자를 순서 맞추기.."	
						Exit Function
					End If
				End If
				Call SetG3W15(pRow)			' 일수 계산 
				Call SetG3W16(pRow)			' 적수 계산 
				'Call SetG3SumValue(C_W14, pRow)	' 잔액 행합계 
				Call SetG3SumValue(C_W15, pRow)	' 일수 행합계 
				Call SetG3SumValue(C_W16, pRow)	' 적수 행합계 
				
				Call SetG3W11(pRow)				' 적요 반영 
				Call SetW3_W4()				' 상단 그리드 반영 
	    End Select

	End With
End Function

' -- 현재 과목을 아래 그리드에 표시 
Sub	SetW2ToChildGrid(Byval pW2)
	Dim i, iMaxRows, iLastRow, iSeqNo
	
	frm1.vspdData.Col = C_SEQ_NO: frm1.vspdData.Row = frm1.vspdData.ActiveRow
	iSeqNo = frm1.vspdData.Text
	
	With frm1.vspdData2
		iMaxRows = .MaxRows 
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData2, i) = False Then 
					.Col = C_W9 : .Row = i : .text = pW2
					
				End If
			End If
		Next
	End With

	With frm1.vspdData3
		iMaxRows = .MaxRows
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData3, i) = False Then 
					.Col = C_W9 : .Row = i : .text = pW2 
				End If
			End If
		Next
	End With
	
End Sub 

Function SetDefaultW8(ByVal pRow)
	Dim iIndex
	
	With Frm1.vspdData
		.Row = pRow
		If lgRateOver = False Then
			.Col = C_W8 : .Text = lgDefaultRate :	iIndex = .Value
			.Col = C_W8_NM : .Value = iIndex
		End If
	End With
End Function


' --- W5에 데이타 계산 
Function SetW5(Byval pRow)
	Dim dblW3, dblW4
	With frm1.vspdData
		.Col = C_W3	: .Row = pRow	: dblW3 = UNICDbl(.Value)
		.Col = C_W4	: .Row = pRow	: dblW4 = UNICDbl(.Value)
		.Col = C_W5	: .Row = pRow
		.Value = (dblW3 - dblW4)
	End With
End Function

' --- Grid2 W14에 데이타 계산 
Function SetG2W14(Byval pRow)
	Dim iRow, bIsFirstRow
	Dim iSeqNo, iChildSeqNo
	Dim dblW12, dblW13, dblW14
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData2
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.text)
		.Col = C_CHILD_SEQ_NO	: .Row = pRow	: iChildSeqNo = UNICDbl(.text)
		bIsFirstRow = True : dblW14 = 0
		For iRow = 1 To .MaxRows-1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.text) = iSeqNo Then
				If bIsFirstRow = False Then
					.Col = C_SEQ_NO	: .Row = iRow - 1
					If iSeqNo = UNICDbl(.text) Then
						.Col = C_W14
						dblW14 = UNICDbl(.text)
					End If
				End If
				bIsFirstRow = False
				
				.Row = iRow 
				.Col = C_W12	: dblW12 = UNICDbl(.text)
				.Col = C_W13	: dblW13 = UNICDbl(.text)
				.Col = C_W14	: .Value = (dblW14 + dblW12 - dblW13)
				
			End If
		Next		
	End With
End Function

' --- Grid3 W14에 데이타 계산 
Function SetG3W14(Byval pRow)
	Dim iRow, bIsFirstRow
	Dim iSeqNo, iChildSeqNo
	Dim dblW12, dblW13, dblW14
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData3
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.text)
		.Col = C_CHILD_SEQ_NO	: .Row = pRow	: iChildSeqNo = UNICDbl(.text)
		bIsFirstRow = True : dblW14 = 0
		For iRow = 1 To .MaxRows-1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.text) = iSeqNo Then
				If bIsFirstRow = False Then
					.Col = C_SEQ_NO	: .Row = iRow - 1
					If iSeqNo = UNICDbl(.text) Then
						.Col = C_W14
						dblW14 = UNICDbl(.text)
					End If
				End If
				bIsFirstRow = False
				
				.Row = iRow 
				.Col = C_W12	: dblW12 = UNICDbl(.text)
				.Col = C_W13	: dblW13 = UNICDbl(.text)
				.Col = C_W14	: .Value = (dblW14 + dblW12 - dblW13)
			End If
		Next
	End With
End Function

' --- Grid2 W15(일수)에 데이타 계산 
Function SetG2W15(Byval pRow)
	Dim datW10, datW10_DOWN, dblSum, iRow, blnPrintLast
	Dim dblW12, dblW13, iSeqNo
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData2
		blnPrintLast = False
		
		.Col = C_W12	: .Row = pRow	: dblW12 = UNICDbl(.Text)
		.Col = C_W13	: .Row = pRow	: dblW13 = UNICDbl(.Text)
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Value)
		If dblW12 = 0 And dblW13 = 0 Then
			Exit Function
		End If
		
		For iRow = .MaxRows-1 To 1 Step -1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.Value) = iSeqNo Then
			
				.Row = iRow
				.Col = C_W10	
				If .Text = "" Then		' 일자가 공란이면 합계후 종료한다.
					
				Else		
	
					datW10 = CDate(.Text)

					If blnPrintLast = False Then	' 마지막행 일수 계산안한경우 
						If frm1.cboREP_TYPE.value = "2" Then
							.Col = C_W15	: .Value = DateDiff("d", datW10, DateAdd("m", 6, lgFISC_START_DT)-1)+1
						Else
							.Col = C_W15	: .Value = DateDiff("d", datW10, lgFISC_END_DT)+1
						End If
						'.Col = C_W15	: .Text = DateDiff("d", datW10, lgFISC_END_DT)+1
						blnPrintLast = True
					Else
						.Col = C_W10	: .Row = iRow+1	
						
						If .Text <> "" Then	' 존재할때.
							datW10_DOWN = CDate(.Text)	' 현재 변경행의 일자를 기억 
							.Col = C_W15	: .Row = iRow	: .Text = DateDiff("d", datW10,  datW10_DOWN)	
						Else
							.Col = C_W15	: .Row = iRow	: .Text = ""
						End If
					End If
				
				End If
			End If
		Next
		
		dblSum = FncSumSheet(frm1.vspdData2, C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' 합계 
	End With	

End Function

' --- Grid3 W15(일수)에 데이타 계산 
Function SetG3W15(Byval pRow)
	Dim datW10, datW10_DOWN, dblSum, iRow, blnPrintLast
	Dim dblW12, dblW13, iSeqNo
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData3
		blnPrintLast = False

		.Col = C_W12	: .Row = pRow	: dblW12 = UNICDbl(.Text)
		.Col = C_W13	: .Row = pRow	: dblW13 = UNICDbl(.Text)
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Value)
		If dblW12 = 0 And dblW13 = 0 Then
			Exit Function
		End If
		
		For iRow = .MaxRows-1 To 1 Step -1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.Value) = iSeqNo Then
				.Row = iRow
				.Col = C_W10	
			
				If .Text = "" Then		' 일자가 공란이면 합계후 종료한다.
					
				Else		
	
					datW10 = CDate(.Text)
			
					If blnPrintLast = False Then	' 마지막행 일수 계산안한경우 
						.Col = C_W15	: .Text = DateDiff("d", datW10, lgFISC_END_DT)+1
						blnPrintLast = True
					Else
						.Col = C_W10	: .Row = iRow+1	
						
						If .Text <> "" Then	' 존재할때.
							datW10_DOWN = CDate(.Text)	' 현재 변경행의 일자를 기억 
							.Col = C_W15	: .Row = iRow	: .Text = DateDiff("d", datW10,  datW10_DOWN)	
						Else
							.Col = C_W15	: .Row = iRow	: .Text = ""
						End If
					End If
				
				End If
			End If
		Next
		
		dblSum = FncSumSheet(frm1.vspdData3, C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' 합계 
	End With	
End Function

' --- Grid2 W16에 데이타 계산 
Function SetG2W16(Byval pRow)
	Dim dblW14, dblW15, iRow
	With frm1.vspdData2
		For iRow =  1 To .MaxRows
			.Col = C_W14	: .Row = iRow	: dblW14 = UNICDbl(.Text)
			.Col = C_W15	: .Row = iRow
			If Trim(.Text) <> "" Then
				dblW15 = UNICDbl(.Text)
				.Col = C_W16	: .Row = iRow
				.Text = (dblW14 * dblW15)
			End If
		Next
	End With
End Function

' --- Grid3 W16에 데이타 계산 
Function SetG3W16(Byval pRow)
	Dim dblW14, dblW15, iRow
	With frm1.vspdData3
		For iRow =  1 To .MaxRows
			.Col = C_W14	: .Row = iRow	: dblW14 = UNICDbl(.Text)
			.Col = C_W15	: .Row = iRow
			If Trim(.Text) <> "" Then
				dblW15 = UNICDbl(.Text)
				.Col = C_W16	: .Row = iRow
				.Text = (dblW14 * dblW15)
			End If
		Next
	End With
End Function

' 2번 그리드에 적요 반영 
Function SetG2W11(ByVal pRow)
	Dim dblW12, dblW13
	Dim datW10
	Dim strDesc
	
	strDesc = ""
	If pRow = 0 Then Exit Function
	With frm1.vspdData2
		.Row = pRow
		.Col = C_W12 : dblW12 = UNICDbl(.Text)
		.Col = C_W13 : dblW13 = UNICDbl(.Text)
		.Col = C_W10

		If dblW12 > 0 Then
			strDesc = "대여"
		ElseIf dblW13 > 0 Then
			strDesc = "상환"
		End If

		If Trim(.Value) <> "" Then
			datW10 = CDate(.Text)
			If Month(datW10) = 1 And Day(datW10) = 1 And dblW12 > 0 Then strDesc = "전기이월"
		End IF
		
		.Col = C_W11 : .Row = pRow : .Text = strDesc
	End With
End Function

' 3번 그리드에 적요 반영 
Function SetG3W11(ByVal pRow)
	Dim dblW12, dblW13
	Dim datW10
	Dim strDesc
	
	strDesc = ""
	If pRow = 0 Then Exit Function
	With frm1.vspdData3
		.Row = pRow
		.Col = C_W12 : dblW12 = UNICDbl(.Text)
		.Col = C_W13 : dblW13 = UNICDbl(.Text)
		.Col = C_W10

		If dblW12 > 0 Then
			strDesc = "대여"
		ElseIf dblW13 > 0 Then
			strDesc = "상환"
		End If

		If Trim(.Text) <> "" Then
			datW10 = CDate(.Text)
			If Month(datW10) = 1 And Day(datW10) = 1 And dblW12 > 0 Then strDesc = "전기이월"
		End IF
		
		.Col = C_W11 : .Row = pRow : .Text = strDesc
	End With
End Function

Function SetG2SumValue(ByVal pCol, ByVal pRow)
	Dim iRow
	Dim iSeqNo, iChlSeqNo, dblSum
	
	dblSum = 0
	
	If pRow = 0 Then Exit Function
	
	With frm1.vspdData2
		
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Text)
		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_SEQ_NO
			
			If iSeqNo = UNICDbl(.Text) Then
				.Col = pCol
			
				If .Text = "" Then		' 일자가 공란이면 합계후 종료한다.					
				Else		
	
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.Text) <> 999999 Then
						.Col = pCol : dblSum = dblSum + UNICDbl(.Text)
					Else
						IF pCol= C_W16 Then
							.Col = pCol : .Text = dblSum
							ggoSpread.Source = frm1.vspdData2
							ggoSpread.UpdateRow iRow
						Else
							.Col = pCol : .Text = ""
						End If
					End If
				
				End If
			End If
		Next
		
	End With	
End Function

Function SetG3SumValue(ByVal pCol, ByVal pRow)
	Dim iRow
	Dim iSeqNo, iChlSeqNo, dblSum
	
	dblSum = 0
	If pRow = 0 Then Exit Function
	
	With frm1.vspdData3
		
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Text)
		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_SEQ_NO
			
			If iSeqNo = UNICDbl(.Text) Then
				.Col = pCol
			
				If .Text = "" Then		' 일자가 공란이면 합계후 종료한다.
					
				Else		
	
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.Text) <> 999999 Then
						.Col = pCol : dblSum = dblSum + UNICDbl(.Text)
					Else
						.Col = pCol : .Text = dblSum
						ggoSpread.Source = frm1.vspdData3
						ggoSpread.UpdateRow iRow
					End If
				
				End If
			End If
		Next
		
	End With	
End Function

' 1번그리드 W3(가지급금적수), W4(가수금 적수)에 넣기 
Function SetW3_W4()
	Dim dblGrid2Sum, dblGrid3Sum, dblSum, iG1Row, iSeqNo, iRow
	
	iG1Row = frm1.vspdData.ActiveRow
	With frm1.vspdData
		.Col = C_SEQ_NO	: .Row = iG1Row	: iSeqNo = UNICDbl(.Value)
	End With
	
	With frm1.vspdData3
		If .MaxRows = 0 Then
			dblGrid3Sum = 0
		Else
			For iRow = 1 To .MaxRows
				.Row = iRow
				.Col = C_SEQ_NO
				If UNICDbl(.value) = iSeqNo Then
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.value) = 999999 Then
						.Col = C_W16 : dblGrid3Sum = UNICDbl(.Value)
						.Col = 0
						If .Text = ggoSpread.DeleteFlag Then dblGrid3Sum = 0
					End If
				End If
			Next
		End If
	End With
	
	With frm1.vspdData2
		If .MaxRows = 0 Then
			dblGrid2Sum = 0
		Else
			For iRow = 1 To .MaxRows
				.Row = iRow
				.Col = C_SEQ_NO
				If UNICDbl(.value) = iSeqNo Then
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.value) = 999999 Then
						.Col = C_W16 : dblGrid2Sum = UNICDbl(.Value)
						.Col = 0
						If .Text = ggoSpread.DeleteFlag Then dblGrid2Sum = 0
					End If
				End If
			Next
		End If
	End With

	With frm1.vspdData
		
		.Col = C_W3	: .Row = iG1Row	: .Value = dblGrid2Sum
		.Col = C_W4	: .Row = iG1Row	: .Value = dblGrid3Sum

		dblSum = FncSumSheet(frm1.vspdData, C_W3, 1, .MaxRows - 1, false, -1, -1, "V")	' 현재 컬럼 행합계 
		.Col = C_W3 : .Row = .MaxRows : .Value = dblSum
		dblSum = FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, false, -1, -1, "V")	' 현재 컬럼 행합계 
		.Col = C_W4 : .Row = .MaxRows : .Value = dblSum
	End With
	
	Call Setw5(iG1Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow iG1Row

	With frm1.vspdData
		dblSum = FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, false, -1, -1, "V")	' 현재 컬럼 행합계 
		.Col = C_W5 : .Row = .MaxRows : .Value = dblSum
		ggoSpread.UpdateRow .MaxRows
	End With
	
End Function

Function ChkG2W10(ByVal pRow)
	Dim datCurW10
	ChkG2W10 = False

	If Frm1.vspdData2.MaxRows <= 0 Then
		ChkG2W10 = True
		Exit Function
	End If
	
	With Frm1.vspdData2
		.Row = pRow : .Col = C_W10
		If Trim(.Value) = "" Then
			ChkG2W10 = True
		Else
			datCurW10 = CDate(.Text)
			If pRow -1  >= 1 Then
				.Row = pRow - 1
				If Trim(.Value) = "" Then
					ChkG2W10 = True
				ElseIf DateDiff("d", CDate(.Text), datCurW10) > 0 Then
					ChkG2W10 = True
				Else
					ChkG2W10 = False
					Exit Function
				End If
			End If
			If pRow + 1 < .MaxRows Then
				.Row = pRow + 1
				If Trim(.Value) = "" Then
					ChkG2W10 = True
				ElseIf DateDiff("d", datCurW10, CDate(.Text)) > 0 Then
					ChkG2W10 = True
				Else
					ChkG2W10 = False
				End If
			ElseIf pRow + 1 = .MaxRows Then
				ChkG2W10 = True
			End If
		End If
	End With
End Function

Function ChkG3W10(ByVal pRow)
	Dim datCurW10
	ChkG3W10 = False

	If Frm1.vspdData3.MaxRows <= 0 Then
		ChkG3W10 = True
		Exit Function
	End If
	
	With Frm1.vspdData3
		.Row = pRow : .Col = C_W10
		If Trim(.Value) = "" Then
			ChkG3W10 = True
		Else
			datCurW10 = CDate(.Text)
			If pRow -1  >= 1 Then
				.Row = pRow - 1
				If Trim(.Value) = "" Then
					ChkG3W10 = True
				ElseIf DateDiff("d", CDate(.Text), datCurW10) > 0 Then
					ChkG3W10 = True
				Else
					ChkG3W10 = False
					Exit Function
				End If
			End If
			If pRow + 1 < .MaxRows Then
				.Row = pRow + 1
				If Trim(.Value) = "" Then
					ChkG3W10 = True
				ElseIf DateDiff("d", datCurW10, CDate(.Text)) > 0 Then
					ChkG3W10 = True
				Else
					ChkG3W10 = False
				End If
			ElseIf pRow + 1 = .MaxRows Then
				ChkG3W10 = True
			End If
		End If
	End With
End Function

'============================== 레퍼런스 함수  ========================================

Sub GetFISC_DATE()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd
	Dim dblConfRate, dblRateLoan
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

	call CommonQueryRs(" ISNULL(MAX(W1), 0)"," TB_LOAN_CALC_SUM "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		dblRateLoan = UNICDbl(lgF0)
	Else
		dblRateLoan = 0
	End if

	call CommonQueryRs(" CONVERT(NUMERIC(5,2), REFERENCE) * 100"," B_CONFIGURATION "," MAJOR_CD = 'W2006' AND MINOR_CD = '1' AND SEQ_NO = 1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		dblConfRate = UNICDbl(lgF0)
	Else
		dblConfRate = 0
	End if
	
	If dblRateLoan >= 0 And dblRateLoan < dblConfRate Then
		lgRateOver = False		' 인정이자율을 당좌대출이자율로 셋팅한다.
		lgDefaultRate = "1"
	Else
		lgRateOver = True		' 인정이자율을 사용자가 선택한다.
	End If
End Sub

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 세무정보 조사 : 컨펌되면 락된다.
'	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

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

    iCalledAspName = AskPRAspName("W1111RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W1111RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    'Call InsertRow2Head
    'Call InsertRow2Detail(1)
    
    Call SetToolbar("1110110100100111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
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
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call GetFISC_DATE

End Sub

Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	Call GetFISC_DATE
End Sub



Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

	End With
End Sub

'==========================================================================================
Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	Call GetFISC_DATE

End Sub


'============================================  1번 그리드 이벤트  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case C_W1		' 성명(법인명)
				.Col = Col
				If .Text = "주택자금" Then
					.Col = C_W7_NM : .Text = "상여" : intIndex = .Value
					.Col = C_W7 : .Value = intIndex		
				Else
					.Col = C_W7_NM : .Text = "기타사외유출" : intIndex = .Value
					.Col = C_W7 : .Value = intIndex		
				End If
			Case  C_W7
				.Col = Col
				intIndex = .Value
				.Col = C_W7_NM
				.Value = intIndex	
			Case  C_W7_NM
				.Col = Col
				intIndex = .Value
				.Col = C_W7
				.Value = intIndex		
			Case  C_W8
				.Col = Col
				intIndex = .Value
				.Col = C_W8_NM
				.Value = intIndex	
			Case  C_W8_NM
				.Col = Col
				intIndex = .Value
				.Col = C_W8
				.Value = intIndex		
		End Select
	End With

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
		ggoSpread.UpdateRow .maxRows
		
		
		Call Fn_GridCalc(Col, Row)
	End With
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow Then
		Call vspdData_Click(Col, NewRow)
	End If
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
	
	Dim iSeqNo, IntRetCD, iLastRow
	
	If lgOldRow = Row  Then Exit Sub
	
    ggoSpread.Source = frm1.vspdData
  
    If Row = frm1.vspdData.MaxRows Then
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
		iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)

	Else
		With frm1.vspdData
			.Col = C_SEQ_NO : .Row = Row : iSeqNo = .Value
			
			' 하위 그리드 표시루틴'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			frm1.vspdData2.SetActiveCell C_W10, iLastRow
			
'			If iLastRow = 0 Then 
'				Call InsertRow2Detail2(iSeqNo)
'				iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
'			End If

			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			frm1.vspdData3.SetActiveCell C_W10, iLastRow
	
'			If iLastRow = 0 Then 
'				Call InsertRow2Detail3(iSeqNo)
'				iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
'			End If

			.focus
		End With	
	End If
  
    lgOldRow = Row	: lgOldCol = Col


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

	lgCurrGrid = 1
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



'============================================  2번 그리드 이벤트  ====================================

Sub vspdData2_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData2
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
		
		If ChkG2W10(Row) = False Then
			Call DisplayMsgBox("WC0016", parent.VB_INFORMATION, "X", "X")           '⊙: "일자를 순서 맞추기.."	
			.Row = Row : .text = ""		
			Exit Sub
		End If
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.UpdateRow Row

		Call Fn_GridCalc2(Col, Row)
	End With
	lgChgFlg = True ' 데이타 변경 
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

End Sub

'============================================  3번 그리드 이벤트  ====================================
Sub vspdData3_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData3
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
		
		If ChkG3W10(Row) = False Then
			Call DisplayMsgBox("WC0016", parent.VB_INFORMATION, "X", "X")           '⊙: "일자를 순서 맞추기.."	
			.Row = Row : .text = ""		
			Exit Sub
		End If
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.UpdateRow Row
		
		Call Fn_GridCalc3(Col, Row)
	End With
	lgChgFlg = True ' 데이타 변경 
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 3
	ggoSpread.Source = Frm1.vspdData3
End Sub

'============================================  툴바지원 함수  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                                <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData2
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData3
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
    Call InitVariables													<%'Initializes local global variables%>
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
	Dim blnChange
        
    FncSave = False           
    blnChange = False                                              
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    
'    If lgChgFlg = False Then
    
	    ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange <> False Then
			blnChange = True
		End If

	    ggoSpread.Source = frm1.vspdData2
		If ggoSpread.SSCheckChange <> False Then
			blnChange = True
'		    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
'		    Exit Function
		End If

	    ggoSpread.Source = frm1.vspdData3
		If ggoSpread.SSCheckChange <> False Then
			blnChange = True
		End If


		If blnChange = False Then
		    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		    Exit Function
		End If

	    ggoSpread.Source = frm1.vspdData
		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		      Exit Function
		End If    


	    ggoSpread.Source = frm1.vspdData2
		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		      Exit Function
		End If    


	    ggoSpread.Source = frm1.vspdData3
		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		      Exit Function
		End If    
		
'	End If
	
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

    Call InsertRow2Head
    
    Call SetToolbar("1110111100000111")

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

			.vspdData.Col = C_W9
			.vspdData.Text = ""
    
			.vspdData.Col = C_W10
			.vspdData.Text = ""
			
			.vspdData.Col = C_W11
			.vspdData.Text = ""
			
			.vspdData.Col = C_W12
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
    Dim lDelRows

	Select Case lgCurrGrid 
		CAse  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 
				If CheckTotalRow(frm1.vspdData, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				ElseIf CheckDetailData(Frm1.vspdData, Frm1.vspdData2, .ActiveRow)  > 0  Or CheckDetailData(Frm1.vspdData, Frm1.vspdData3, .ActiveRow)  > 0 Then
					MsgBox "하위 데이타가 존재하여 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.EditUndo

					lDelRows = ggoSpread.EditUndo
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
				End If
				Call SetW3_W4()
			End With
		CAse 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.EditUndo

					lDelRows = ggoSpread.EditUndo
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData2, lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
				End If
				Call SetW3_W4()
			End With    
 		CAse 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow2(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.EditUndo

					lDelRows = ggoSpread.EditUndo
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData3, lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
				End If
				Call SetW3_W4()
			End With     
	End Select
  
	lgChgFlg = True                                                '☜: Protect system from crashing
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, iLastRow
    Dim iStrNm

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
   
	Frm1.vspdData.Col = C_W2
	Frm1.vspdData.Row = Frm1.vspdData.ActiveRow
   	iStrNm = Frm1.vspdData.Text

	With frm1	

		' 첫행일 경우 합계까지 넣는 루틴 
		If .vspdData.MaxRows = 0 Then
			Call InsertRow2Head
			Call SetToolbar("1110111100000111")
			Exit Function
		End If
		
		Select Case lgCurrGrid
			Case 1	' 1번 그리드 
		
			.vspdData.focus
			ggoSpread.Source = .vspdData
			
			iRow = .vspdData.ActiveRow	' 현재행 
			
			.vspdData.ReDraw = False
			
			If iRow = .vspdData.MaxRows Then
		
				' SEQ_NO 를 그리드에 넣는 로직 
				iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
				ggoSpread.InsertRow iRow-1 ,imRow	' 그리드 행 추가(사용자 행수 포함)
				SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1	' 그리드 색상변경 
		
				For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
					.vspdData.Row = iRow
					.vspdData.Col = C_SEQ_NO
					.vspdData.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.

					Call SetDefaultW8(iRow)		' 인정이자율 설정 

				Next				

				'ggoSpread.InsertRow iRow-1 , imRow 
				'SetSpreadColor iRow, iRow + imRow - 1
				'MaxSpreadVal2 .vspdData, C_SEQ_NO, iRow	
				'vspdData.Col = C_SEQ_NO : .vspdData.Row = Row : iSeqNo = .vspdData.Value
			Else

				' SEQ_NO 를 그리드에 넣는 로직 
				iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' 최대SEQ_NO를 구해온다.
			
				ggoSpread.InsertRow ,imRow	' 그리드 행 추가(사용자 행수 포함)
				SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1	' 그리드 색상변경 
		
				For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' 추가된 그리드의 SEQ_NO를 변경한다.
					.vspdData.Row = iRow
					.vspdData.Col = C_SEQ_NO
					.vspdData.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO를 증가한다.

					Call SetDefaultW8(iRow)		' 인정이자율 설정 
				Next			
				'ggoSpread.InsertRow ,imRow
				'SetSpreadColor iRow+1, iRow+1
				'MaxSpreadVal .vspdData, C_SEQ_NO, iRow+1
				'.vspdData.Col = C_SEQ_NO : .vspdData.Row = Row+1 : iSeqNo = .vspdData.Value
			End If

			.vspdData.ReDraw = True	
						
			' 하위 그리드 표시루틴'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			frm1.vspdData2.SetActiveCell C_W7, iLastRow
			
			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			frm1.vspdData3.SetActiveCell C_W7, iLastRow
			
			Call vspdData_Click(.vspdData.Col, .vspdData.ActiveRow)
		Case 2	' 2번 그리드 
			.vspdData2.focus
			ggoSpread.Source = .vspdData2

			.vspdData.Col = C_SEQ_NO : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = .vspdData.Value

			' 첫행일 경우 합계까지 넣는 루틴 
			If .vspdData.ActiveRow = .vspdData.MaxRows Then
				Exit Function
			ElseIf ShowRowHidden(frm1.vspdData2, iSeqNo) = 0 Then
				Call InsertRow2Detail2(iSeqNo)
				Call ShowRowHidden(frm1.vspdData2, iSeqNo)
			Else
				'.vspdData2.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
				iRow = .vspdData2.ActiveRow
				
				If iRow = .vspdData2.MaxRows Then
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColorDetail2 iRow
					MaxSpreadVal2 .vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow , iSeqNo
				Else
					ggoSpread.InsertRow ,imRow
					SetSpreadColorDetail2 iRow+1
					MaxSpreadVal2 .vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
				End If	
			End If
			Call SetW2ToChildGrid(iStrNm)	' 현재 이름을 하위 그리드에 넣는다.

		Case 3	' 3번 그리드 
			.vspdData3.focus
			ggoSpread.Source = .vspdData3

			.vspdData.Col = C_SEQ_NO : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = .vspdData.Value

			' 첫행일 경우 합계까지 넣는 루틴 
			If .vspdData.ActiveRow = .vspdData.MaxRows Then
				Exit Function
			ElseIf ShowRowHidden(frm1.vspdData3, iSeqNo) = 0 Then
				Call InsertRow2Detail3(iSeqNo)
				Call ShowRowHidden(frm1.vspdData3, iSeqNo)
			Else
				'.vspdData3.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
				iRow = .vspdData3.ActiveRow
				
				If iRow = .vspdData3.MaxRows Then
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColorDetail3 iRow
					MaxSpreadVal2 .vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow	, iSeqNo
				Else
					ggoSpread.InsertRow ,imRow
					SetSpreadColorDetail3 iRow+1
					MaxSpreadVal2 .vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
				End If	
			End If
			Call SetW2ToChildGrid(iStrNm)	' 현재 이름을 하위 그리드에 넣는다.
		End Select
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

	Select Case lgCurrGrid 
		Case  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 
				If CheckTotalRow(frm1.vspdData, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				ElseIf CheckDetailData(Frm1.vspdData, Frm1.vspdData2, .ActiveRow)  > 0  Or CheckDetailData(Frm1.vspdData, Frm1.vspdData3, .ActiveRow)  > 0 Then
					MsgBox "하위 데이타가 존재하여 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.DeleteRow

					lDelRows = ggoSpread.DeleteRow
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
				End If
				Call SetW3_W4()
			End With
		Case 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.DeleteRow

					lDelRows = ggoSpread.DeleteRow
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData2, lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
				End If
				Call SetW3_W4()
			End With    
 		Case 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow2(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.DeleteRow

					lDelRows = ggoSpread.DeleteRow
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData3, lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
				End If
				Call SetW3_W4()
			End With     
	End Select
	
	lgChgFlg = True
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
	
    Dim iNameArr, iSeqNo, iLastRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
        
    Call SetToolbar("1111111100110111")										<%'버튼 툴바 제어 %>
	
	Call RedrawSumRow
	Call RedrawSumRow2
	Call RedrawSumRow3

	With frm1.vspdData
		.Col = C_SEQ_NO : .Row = .ActiveRow : iSeqNo = .Value
			
		' 하위 그리드 표시루틴'
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)

		' 하위 그리드 표시루틴'
		iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
	End With			
	frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow , lCol, lGrpCnt, lMaxRows, lMaxCols
    Dim lStartRow, lEndRow , lChkAmt
    Dim strVal
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
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
		                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
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

    frm1.txtSpread.value      = strVal
    strVal = ""

 	With frm1.vspdData2
		' ----- 2번째 그리드 
		ggoSpread.Source = frm1.vspdData2
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
                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 
		 '잔액검증 20060125 by HJO
		 .Col = C_W14 : .Row = lRow
			If UNICDbl(.Text) < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "가지지급금(14)잔액", "X")           '⊙: "일자를 순서 맞추기.."	
				Call  LayerShowHide(0)
				Exit Function
			Else
			.Col=0
			  ' 모든 그리드 데이타 보냄     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = C_SEQ_NO To lMaxCols
	'					If lCol = C_W10 Then
	'						.Col = lCol : strVal = strVal & CDate(.Text) &  Parent.gColSep
	'					Else
							.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
	'					End If
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
			End If
		Next
	End With

    frm1.txtSpread2.value      = strVal
    strVal = ""
    	
	With frm1.vspdData3
		' ----- 3번째 그리드 
		ggoSpread.Source = frm1.vspdData3
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
		                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 '잔액검증 20060125 by HJO
		 .Col = C_W14 : .Row = lRow
		If UNICDbl(.Text) < 0 Then
			Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "가수금(14)잔액", "X")           '⊙: "일자를 순서 맞추기.."	
			Call  LayerShowHide(0)
			Exit Function
		End If
		.Col=0
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With
	
    frm1.txtSpread3.value      = strVal
    strVal = ""	  
        
	frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
	frm1.txtFlgMode.value     = lgIntFlgMode

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
    
	frm1.vspdData3.MaxRows = 0
    ggoSpread.Source = frm1.vspdData3
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
<BODY TABINDEX="-1" SCROLL="">
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<!--<a href="vbscript:GetRef">인사데이타 불러오기</A>  -->
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
									<TD CLASS="TD6"><script language =javascript src='./js/w3117ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=1>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 1. </LEGEND>
                                   <script language =javascript src='./js/w3117ma1_vaSpread_vspdData.js'></script>
								  </FIELDSET>
								  <BR>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 2. </LEGEND>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 가. 가지급금</LEGEND>
                                   <script language =javascript src='./js/w3117ma1_vaSpread2_vspdData2.js'></script>
								  </FIELDSET>
								  <BR>
									<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 나. 가수금</LEGEND>
									<script language =javascript src='./js/w3117ma1_vaSpread3_vspdData3.js'></script>
								  </FIELDSET>
								  <BR>
								  </FIELDSET>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>가지급의 적수계산</LABEL>&nbsp;
				            <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><별지>가수금의 적수계산</LABEL>&nbsp;
				            
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" STYLE="Display:none"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" STYLE="Display:none"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24" STYLE="Display:none"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_CO_CD" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_FISC_YEAR" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_REP_TYPE" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

