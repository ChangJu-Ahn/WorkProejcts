<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 수입금액조정 
'*  3. Program ID           : W1111MA1
'*  4. Program Name         : W1111MA1.asp
'*  5. Program Desc         : 제16호 수입금액 조정명세서 
'*  6. Modified date(First) : 2005/01/03
'*  7. Modified date(Last)  : 2006/01/23
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

Const BIZ_MNU_ID = "w2105MA1"
Const BIZ_PGM_ID = "w2105mb1.asp"	
Const EBR_RPT_ID = "w2105OA1"										 '☆: 비지니스 로직 ASP명 

' -- 1번 수입금액조정계산 그리드 
Dim C_SEQ_NO
Dim C_W1_CD
Dim C_W1_NM
Dim C_W2_CD
Dim C_W2_NM
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_DESC1

' -- 2번 수입금액 조정명세 가. 작업진행률에 의한 수입금액 그리드 
Dim C_CHILD_SEQ_NO
Dim C_W2_NM2
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16

' -- 3번 수입금액 조정명세 나. 기타 수입금액 그리드 
Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_DESC2

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgOldCol, lgOldRow , lgChgFlg

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	lgCurrGrid	= 1
	lgOldRow	= 0
	lgOldCol	= 2
	lgChgFlg	= False

	'--1번수입금액조정계산그리드 
	C_SEQ_NO	= 1
	C_W1_CD		= 2
	C_W1_NM		= 3
	C_W2_CD		= 4
	C_W2_NM		= 5
	C_W3		= 6
	C_W4		= 7
	C_W5		= 8
	C_W6		= 9
	C_DESC1		= 10

	'--2번수입금액조정명세가.작업진행률에의한수입금액그리드 
	C_CHILD_SEQ_NO	= 2
	C_W2_NM2	= 3
	C_W7		= 4
	C_W8		= 5
	C_W9		= 6
	C_W10		= 7
	C_W11		= 8
	C_W12		= 9
	C_W13		= 10
	C_W14		= 11
	C_W15		= 12
	C_W16		= 13

	'--3번수입금액조정명세나.기타수입금액그리드 
	C_CHILD_SEQ_NO	= 2
	'C_W2_NM2	= 3
	C_W17		= 4
	C_W18		= 5
	C_W19		= 6
	C_W20		= 7
	C_DESC2		= 8
	
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
    lgOldRow = 0
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
   
    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_DESC1 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols									'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    ggoSpread.ClearSpreadData
    .MaxRows = 0
    
	'헤더를 2줄로    
    .ColHeaderRows = 2    
    ' 
    Call AppendNumberPlace("6","3","2")

    ggoSpread.SSSetEdit		C_SEQ_NO,	"순번"		, 5,,,6,1	' 히든컬럼 
	ggoSpread.SSSetEdit		C_W1_CD,	"(1)항목"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W1_NM,	"(1)항목"	, 15,,,50,1	
	ggoSpread.SSSetEdit		C_W2_CD,	"(2)과목"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W2_NM,	"(2)과목"	, 15,,,50,1	
	ggoSpread.SSSetFloat	C_W3,		"(3)결산서상" & vbCrLf & "수입금액"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	ggoSpread.SSSetFloat	C_W4,		"(4)가산"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	ggoSpread.SSSetFloat	C_W5,		"(5)차감"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetFloat	C_W6,		"(6)조정후 수입금액" & vbCrLf & "[(3) + (4) - (5)]", 15,	Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetEdit		C_DESC1,	"비 고", 20,,,20,1

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
	Call ggoSpread.SSSetColHidden(C_W1_CD,C_W1_CD,True)
	Call ggoSpread.SSSetColHidden(C_W2_CD,C_W2_CD,True)
					
	'Call InitSpreadComboBox()	콤보없음 

	' 그리드 헤더 합침 정의 
	ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)	' SEQ_NO 합침 
    ret = .AddCellSpan(C_W1_CD, -1000, 4, 1)	' 계정과목 셀 합침 
    ret = .AddCellSpan(C_W3, -1000, 1, 2)	' 결산서상수입금액 행 합침 
    ret = .AddCellSpan(C_W4, -1000, 2, 1)	' 조정 셀 합침 
    ret = .AddCellSpan(C_W6, -1000, 1, 2)	' 수정후수입금액 행 합침 
    ret = .AddCellSpan(C_DESC1, -1000, 1, 2)	' 비고 행 합침 

       
     ' 첫번째 헤더 출력 글자 
	.Row = -1000
	.Col = C_W1_CD
	.Text = "계 정 과 목"
	.Col = C_W4
	.Text = "조       정"
		
	' 두번째 헤더 출력 글자 
	.Row = -999	
	.Col = C_W1_NM
	.Text = "(1)항 목"
	.Col = C_W2_NM
	.Text = "(2)과 목"
	.Col = C_W4	
	.Text = "(4)가 산"
	.Col = C_W5
	.Text = "(5)차 감"

	.rowheight(-999) = 15	' 높이 재지정 
   	
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
 
	'헤더를 2줄로    
    .ColHeaderRows = 2
    'Call AppendNumberPlace("6","3","2")

	ggoSpread.SSSetEdit		C_SEQ_NO,	"부모순번", 5,,,6,1	' 히든컬럼 
    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"자식순번", 5,,,6,1	' 히든컬럼 
    ggoSpread.SSSetEdit		C_W2_NM2,	"과 목"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W7,		"(7)공사명", 10,,,50,1
	ggoSpread.SSSetEdit		C_W8,		"(8)도급자", 10,,,50,1
    ggoSpread.SSSetFloat	C_W9,		"(9)도급금액" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
	ggoSpread.SSSetFloat	C_W10,		"(10)당해사업 연도말" & vbCrLf & "총공사비 누적액" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetFloat	C_W11,		"(11)총공사 예정비",		13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
    ggoSpread.SSSetEdit		C_W12,		"(12)진행율",		10, 2,,10,2
    ggoSpread.SSSetFloat	C_W13,		"(13)누적익금" & vbCrLf & "산입액" & vbCrLf & "[(9) * (12)]" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W14,		"(14)전기말" & vbCrLf & "누적수입" & vbCrLf & "계상액" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W15,		"(15)당기회사" & vbCrLf & "수입계상액" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W16,		"(16)조정액" & vbCrLf & "[(13) - (14) - (15)]" ,14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""

	' 퍼센트 형 정의 
    .Col = C_W12
    .Row = -1
    .CellType = 14
    .TypeHAlign = 2
    '.TypePercentDecimal = 0
    .TypePercentMax = 999
    '.TypePercentMin = 0
    '.TypePercentDecPlaces = 0
    
	' 그리드 헤더 합침 정의 
	'ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)	' SEQ_NO 행 합침 
    'ret = .AddCellSpan(C_CHILD_SEQ_NO, -1000, 1, 2)	' SEQ_NO 행 합침 
    ret = .AddCellSpan(C_W2_NM2, -1000, 1, 2)	' 공사명 
    ret = .AddCellSpan(C_W7, -1000, 1, 2)	' 공사명 
    ret = .AddCellSpan(C_W8, -1000, 1, 2)	' 도급자 
    ret = .AddCellSpan(C_W9, -1000, 1, 2)	' 도급금액 
    ret = .AddCellSpan(C_W10, -1000, 3, 1)	' 작업진행률계산 
    ret = .AddCellSpan(C_W13, -1000, 1, 2)	' 입급산입액 
    ret = .AddCellSpan(C_W14, -1000, 1, 2)	' 전기말 
    ret = .AddCellSpan(C_W15, -1000, 1, 2)	' 당기회사 
    ret = .AddCellSpan(C_W16, -1000, 1, 2)	' 조정 
    
    ' 첫번째 헤더 출력 글자 
	.Row = -1000
	.Col = C_W10
	.Text = "작업진행률계산"

	' 두번째 헤더 출력 글자 
	.Row = -999
	.Col = C_W10
	.Text = "(10)당해사업 연도말" & vbCrLf & "총공사비 누적액"
	.Col = C_W11
	.Text = "(11)총공사 예정비"
	.Col = C_W12
	.Text = "(12)진행율" & vbCrLf & "[(10)/(11)]"
	.rowheight(-999) = 20	' 높이 재지정 
	
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_CHILD_SEQ_NO,True)
				
	'Call InitSpreadComboBox()
	
	.ReDraw = true
	
  'Call SetSpreadLock 
    
    End With

 	' -----  3번 그리드 
	With frm1.vspdData3
	
	ggoSpread.Source = frm1.vspdData3
   'patch version
    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
    
	.ReDraw = false
    
    .MaxCols = C_DESC2 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols									'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    ggoSpread.ClearSpreadData
   .MaxRows = 0
 
    'Call AppendNumberPlace("6","3","2")

    ggoSpread.SSSetEdit		C_SEQ_NO,	"부모순번", 5,,,6,1	' 히든컬럼 
    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"자식순번", 5,,,6,1	' 히든컬럼 
    ggoSpread.SSSetEdit		C_W2_NM2,	"과 목"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W17,		"(17)구 분", 20,,,50,1
	ggoSpread.SSSetEdit		C_W18,		"(18)근거법령", 20,,,50,1
    ggoSpread.SSSetFloat	C_W19,		"(19)수입금액" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
	ggoSpread.SSSetFloat	C_W20,		"(20)대응원가" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	ggoSpread.SSSetEdit		C_DESC2,	"비 고", 20,,,50,1
		
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_CHILD_SEQ_NO,True)
				
	'Call InitSpreadComboBox()
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
           
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    .vspdData2.ReDraw = False
    .vspdData3.ReDraw = False

	' 1번 그리드 
    ggoSpread.Source = frm1.vspdData
        
	ggoSpread.SpreadLock C_SEQ_NO, -1, C_W1_CD
	ggoSpread.SSSetRequired C_W1_NM, -1, -1
	ggoSpread.SSSetRequired C_W2_NM, -1, -1
	ggoSpread.SSSetRequired C_W3, -1, -1
    ggoSpread.SpreadLock C_W4, -1, C_W4
    ggoSpread.SpreadLock C_W5, -1, C_W5
    ggoSpread.SpreadLock C_W6, -1, C_W6    
    
    ' 2번 그리드 
    ggoSpread.Source = frm1.vspdData2	

    ggoSpread.SpreadLock C_W2_NM2, -1, C_W2_NM2
    ggoSpread.SpreadLock C_W12, -1, C_W12
    ggoSpread.SpreadLock C_W13, -1, C_W13
    ggoSpread.SpreadLock C_W16, -1, C_W16

	
	' 3번 그리드 
    ggoSpread.Source = frm1.vspdData3	

    'ggoSpread.SpreadLock C_W17, -1, C_W17
	ggoSpread.SpreadLock C_W2_NM2, -1, C_W2_NM2
    
	'ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
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
	
  		ggoSpread.SSSetProtected C_SEQ_NO, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_CHILD_SEQ_NO, pvEndRow, pvEndRow
		ggoSpread.SSSetRequired C_W1_NM, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W2_NM, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W3, pvStartRow, pvEndRow 		
 		
 		ggoSpread.SSSetProtected C_W4, -1, -1
 		ggoSpread.SSSetProtected C_W5, -1, -1
 		ggoSpread.SSSetProtected C_W6, -1, -1
		    
		'.vspdData.ReDraw = True

    'End If
    
    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColorDetail2(ByVal pvEndRow)
    With frm1
    
		' 2번 그리드 
		ggoSpread.Source = frm1.vspdData2	

		ggoSpread.SpreadLock C_SEQ_NO, -1, C_CHILD_SEQ_NO
		ggoSpread.SSSetProtected C_W2_NM2, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_W12, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_W13, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_W16, pvEndRow, pvEndRow
    
    End With
End Sub

' InsertRow/Copy 할때 호출됨 
Sub SetSpreadColorDetail3(ByVal pvEndRow)
    With frm1
    
		' 2번 그리드 
		ggoSpread.Source = frm1.vspdData3	

		ggoSpread.SpreadLock C_SEQ_NO, -1, C_CHILD_SEQ_NO
		ggoSpread.SSSetProtected C_W2_NM2, pvEndRow, pvEndRow

    
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

Sub InsertRow2Head()
	' fncNew, onLoad시에 호출해서 기본적으로 3칸을 입력함 
	Dim ret, iRow, iLoop, iSeqNo
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
			
		.ReDraw = False

		iRow = 1
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow)
		
		iRow = 2
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		.Col = C_SEQ_NO : .Row = iRow: .value = SUM_SEQ_NO
		.col = C_W1_CD : .text = "계" : .TypeHAlign = 2
				
		ggoSpread.SpreadLock C_W1_CD, iRow, C_DESC1-1, iRow
		ret = .AddCellSpan(C_W1_CD, iRow, 4, 1)	' 합계 행의 계정과목 셀을 합침 
		
		.ReDraw = True		
		.focus
		.SetActiveCell C_W1_NM, 1
					
	End With

	Call InsertRow2Detail2(iSeqNo)
	Call InsertRow2Detail3(iSeqNo)
	
	Call vspdData_Click(C_W1_NM, 1)
End Sub

Sub InsertRow2Detail2(Byval pSeqNo)

	' 작업진행률 그리드 추가 
	Dim ret, iRow, iLoop, iLastRow, sW2_NM
	
	sW2_NM = GetGrid(frm1.vspdData, C_W2_NM, frm1.vspdData.ActiveRow)
	
	With frm1.vspdData2
		
		.focus
		ggoSpread.Source = frm1.vspdData2

		iLastRow = .MaxRows
		.SetActiveCell C_W2_NM2, iLastRow	
		
		.ReDraw = False
		For iRow = 1 to 2
			If iRow Mod 2 = 0 Then	' 합계행 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .value = SUM_SEQ_NO
				.Col = C_SEQ_NO			: .Text = pSeqNo
				.Col = C_W2_NM2			: .Text = "계"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail2(iLastRow+iRow) 

				ggoSpread.SpreadLock C_W9, iLastRow+iRow, C_W16, iLastRow+iRow
				ret = .AddCellSpan(C_W2_NM2, iLastRow+iRow, 3, 1)	' 합계 행의 계정과목 셀을 합침	
				ggoSpread.SpreadLock 1, iLastRow+iRow, C_W16, iLastRow+iRow	
				.RowHidden = True	
			Else	 ' 그외 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .Text = iRow
				.Col = C_SEQ_NO			: .Text = pSeqNo
				.Col = C_W2_NM2			: .Text = sW2_NM
				Call SetSpreadColorDetail2(iLastRow+iRow) 
				.RowHidden = True
			End If
		Next

		.ReDraw = True		
		.SetActiveCell C_W2_NM2, iLastRow+1	

	End With
	
End Sub

Sub InsertRow2Detail3(Byval pSeqNo)

	' 작업진행률 그리드 추가 
	Dim ret, iRow, iLoop, iLastRow, sW2_NM
	
	sW2_NM = GetGrid(frm1.vspdData, C_W2_NM, frm1.vspdData.ActiveRow)
	
	' 기타수입금액 그리드	
	With frm1.vspdData3
		
		.focus
		ggoSpread.Source = frm1.vspdData3

		iLastRow = .MaxRows
		.SetActiveCell C_W2_NM2, iLastRow	
	
		.ReDraw = False
		For iRow = 1 to 2
			If iRow Mod 2 = 0 Then	' 합계행 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .Text = "999999"
				.Col = C_SEQ_NO			: .Text = pSeqNo
				.Col = C_W2_NM2			: .Text = "계"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail3(iLastRow+iRow) 
				
				'ggoSpread.SpreadLock C_W9, iLoop+iRow-1, C_W16, iLoop+iRow-1
				ret = .AddCellSpan(C_W2_NM2, iLastRow+iRow, 3, 1)	' 합계 행의 계정과목 셀을 합침	
				ggoSpread.SpreadLock 1, iLastRow+iRow, C_DESC2, iLastRow+iRow	
				.RowHidden = True	
			Else	 ' 그외 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .Text = iRow
				.Col = C_SEQ_NO			: .Text = pSeqNo	
				.Col = C_W2_NM2			: .Text = sW2_NM
				Call SetSpreadColorDetail3(iLastRow+iRow) 

				.RowHidden = True
			End If
		Next
	
		.ReDraw = True	
		.SetActiveCell C_W2_NM2, iLastRow+1	
	
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
				.col = C_W1_CD : .text = "계" : .TypeHAlign = 2
				
				ggoSpread.SpreadLock C_W1_CD, iRow, C_DESC1-1, iRow
				ret = .AddCellSpan(C_W1_CD, iRow, 4, 1)	' 합계 행의 계정과목 셀을 합침	
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
			
				.Col = C_W2_NM2			: .Text = "계"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail2(iRow) 

				ggoSpread.SpreadLock C_W9, iRow, C_W16, iRow
				ret = .AddCellSpan(C_W2_NM2, iRow, 3, 1)	' 합계 행의 계정과목 셀을 합침	
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
			
				.Col = C_W2_NM2			: .Text = "계"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail3(iRow) 

				ret = .AddCellSpan(C_W2_NM2, iRow, 3, 1)	' 합계 행의 계정과목 셀을 합침	
				ggoSpread.SpreadLock 1, iRow, C_DESC2, iRow	
			End If
		Next
	End With
End Sub

'============================== 사용자 정의 함수  ========================================
' -- 행을 히든 처리 
Function ShowRowHidden(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
	iMaxRows = .MaxRows : iFirstRow = 0
	pObj.ReDraw = False
	For iRow = 1 To iMaxRows
		.Col = C_SEQ_NO : .Row = iRow : iSeqNo = .Value
		If iSeqNo = pSeqNo Then	' 같은 관계라면..
			.RowHidden = False
			If iFirstRow = 0 Then iFirstRow = iRow
		Else
			.RowHidden = True
		End If
	Next
	pObj.ReDraw = True
	ShowRowHidden = iFirstRow
	End With
End Function

Function ShowRowHidden2(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
	iMaxRows = .MaxRows : iFirstRow = 0
	pObj.ReDraw = False
	For iRow = 1 To iMaxRows
		.Col = C_CHILD_SEQ_NO : .Row = iRow : iSeqNo = .Value
		If iSeqNo = pSeqNo Then	' 같은 관계라면..
			.RowHidden = False
			If iFirstRow = 0 Then iFirstRow = iRow
		Else
			.RowHidden = True
		End If
	Next
	pObj.ReDraw = True
	ShowRowHidden2 = iFirstRow
	End With
End Function

' -- 합계 행인지 체크 
Function CheckTotalRow(Byref pObj, Byval pRow) 
	CheckTotalRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If CDbl(pObj.Text) = 999999 Then	 ' 합계 행 
		CheckTotalRow = True
	End If
End Function

' -- 합계 행인지 체크 
Function CheckTotalRow2(Byref pObj, Byval pRow) 
	CheckTotalRow2 = False
	pObj.Col = C_CHILD_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If CDbl(pObj.Text) = 999999 Then	 ' 합계 행 
		CheckTotalRow2 = True
	End If
End Function

' -- 현재 과목을 아래 그리드에 표시 
Sub	SetW2ToChildGrid(Byval pW2)
	Dim i, iMaxRows, iLastRow, iSeqNo
	
	frm1.vspdData2.Col = C_SEQ_NO: frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	iSeqNo = frm1.vspdData2.Value
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		iMaxRows = .MaxRows 
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData2, i) = False Then 
					.Col = C_W2_NM2 : .Row = i : .Value = pW2 
					ggoSpread.UpdateRow .Row
				End If
			End If
		Next
	End With

	With frm1.vspdData3
		ggoSpread.Source = frm1.vspdData3
		iMaxRows = .MaxRows
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData3, i) = False Then 
					.Col = C_W2_NM2 : .Row = i : .Value = pW2 
					ggoSpread.UpdateRow .Row
				End If
			End If
		Next
	End With
	
End Sub 

' --- W12에 데이타 계산 
Function SetW12(Byval pRow)
	Dim dblW10, dblW11
	With frm1.vspdData2
		.Col = C_W10	: .Row = pRow	: dblW10 = CDbl(.Value)
		.Col = C_W11	: .Row = pRow	: dblW11 = CDbl(.Value)
		.Col = C_W12	: .Row = pRow
		If dblW11 > 0 Then
			.Value = dblW10/dblW11
		Else
			.Value = 0
		End If
	End With
End Function

' --- W13에 데이타 계산 
Function SetW13(Byval pRow)
	Dim dblW9, dblW12
	With frm1.vspdData2
		.Col = C_W9		: .Row = pRow	: dblW9 = CDbl(.Value)
		.Col = C_W12	: .Row = pRow	: dblW12 = CDbl(.Value)
		.Col = C_W13	: .Row = pRow
		If dblW12 > 0 Then
			.Value = (dblW9*dblW12)
		Else
			.Value = 0
		End If
	End With
End Function

' --- W16에 데이타 계산 
Function SetW16(Byval pRow)
	Dim dblW13, dblW14, dblW15
	With frm1.vspdData2
		.Col = C_W13	: .Row = pRow	: dblW13 = CDbl(.Value)
		.Col = C_W14	: .Row = pRow	: dblW14 = CDbl(.Value)
		.Col = C_W15	: .Row = pRow	: dblW15 = CDbl(.Value)
		.Col = C_W16	: .Row = pRow	: .Value = dblW13 - dblW14 - dblW15
	End With
End Function

' 1번그리드 W4(가산)에 넣기 
Function SetW4_W5()
	Dim dblGrid2Sum, dblW19Sum, dblW20Sum, dblSum
	
	' 
	With frm1.vspdData3
		dblW19Sum = GetSum(frm1.vspdData3, C_W19)
		dblW20Sum = GetSum(frm1.vspdData3, C_W20)
	End With
	
	With frm1.vspdData2
		dblGrid2Sum = GetSum(frm1.vspdData2, C_W16)
	End With

	With frm1.vspdData
		
		If 	dblGrid2Sum > 0 Then
			.Col = C_W4	: .Row = .ActiveRow	: .Value = dblGrid2Sum + dblW19Sum
			.Col = C_W5	: .Row = .ActiveRow	: .Value = dblW20Sum
		Else
			.Col = C_W5	: .Row = .ActiveRow	: .Value = ABS(dblGrid2Sum) + dblW20Sum
			.Col = C_W4	: .Row = .ActiveRow	: .Value = dblW19Sum
		End If

			dblSum = FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, false, -1, -1, "V")	' 현재 컬럼 행합계 
			.Col = C_W4 : .Row = .MaxRows : .Value = dblSum
			dblSum = FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, false, -1, -1, "V")	' 현재 컬럼 행합계 
			.Col = C_W5 : .Row = .MaxRows : .Value = dblSum
			
			Call SetW6(.ActiveRow)
	End With
	
	
End Function

' -- 현재 보이는 그리드의 특정컬럼의 합계행의 값을 읽어온다.
Function GetSum(Byref pGrid, Byval pCol)
	Dim iRow, iMaxRows, iSeqNo
	
	With pGrid
		iMaxRows = .MaxRows
		.Row = .ActiveRow	: .Col = C_SEQ_NO : iSeqNo = .Value
		For iRow = 1 To iMaxRows
			.Row = iRow : .Col = C_SEQ_NO
			If .Value = iSeqNo Then
				.Col = C_CHILD_SEQ_NO 
				If UNICDbl(.Value) = SUM_SEQ_NO Then
					.Col = pCol
					GetSum = UNICDbl(.Value)
					Exit Function
				End If
			End If
		Next
	End With
End Function

' -- W6(조정후수입금액)
Function SetW6(Byval Row)
	Dim dblW3, dblW4, dblW5, dblSum
	With frm1.vspdData
		.Col = C_W3	: .Row = Row	: dblW3 = CDbl(.value)
		.Col = C_W4	: .Row = Row	: dblW4 = CDbl(.value)
		.Col = C_W5	: .Row = Row	: dblW5 = CDbl(.value)
		.Col = C_W6	: .Row = Row	: .value = dblW3 + dblW4 - dblW5
		
		dblSum = FncSumSheet(frm1.vspdData, C_W6, 1, .MaxRows - 1, false, -1, -1, "V")	' 현재 컬럼 행합계 
		.Col = C_W6 : .Row = .MaxRows : .Value = dblSum
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow .ActiveRow
		ggoSpread.UpdateRow .MaxRows
	End With
End Function

'============================== 레퍼런스 함수  ========================================
Function GetRef()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD, iMaxRows, iRow
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

	arrParam(0) = frm1.txtCO_CD.value 
	arrParam(1) = frm1.txtFISC_YEAR.text 
	arrParam(2) = frm1.cboREP_TYPE.value 

    arrRet = window.showModalDialog("w2105ra1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
	If arrRet(0, 0) = "" Then
	    Exit Function
	End If	

	
	With frm1.vspdData
		.Redraw = False
		
		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.ClearSpreadData
		iMaxRows = UBound(arrRet, 1)
		lgCurrGrid = 1
		Call FncInsertRow(1)
		If iMaxRows > 1 Then Call FncInsertRow(iMaxRows-1)
		
		For iRow = 1 To iMaxRows
			.Row = iRow	
			Call vspdData_Click(C_W1_NM, iRow)
			.Row = iRow	
			'.Col = C_W1_CD	: .Value = arrRet(iRow, 3)
			.Col = C_W1_NM	: .Value = arrRet(iRow-1, 5)
			.Col = C_W2_CD	: .Value = arrRet(iRow-1, 3)
			.Col = C_W2_NM	: .Value = arrRet(iRow-1, 4)
			.Col = C_W3		: .Value = arrRet(iRow-1, 2)
			.Col = C_W6		: .Value = arrRet(iRow-1, 2)

			Call vspdData_Change(C_W2_NM, iRow)
			Call vspdData_Change(C_W3, iRow)
		Next
		.Row = 1
		.Col = 	C_W1_NM
		.Action = 0
		.Redraw = True
		
		Call vspdData_Click(C_W1_NM, 1)
	End With
	
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
    
    Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>

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
End Sub



'============================================  1번 그리드 이벤트  ====================================
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
	    
	Select Case Col
		Case C_W2_NM	' 과목명 
			.Col = C_W2_NM
			Call SetW2ToChildGrid(.Value)	' 현재 과목을 하위 그리드에 넣는다.
		Case C_W3		' 결산서상 수입금액 
			dblSum = FncSumSheet(frm1.vspdData, Col, 1, .MaxRows - 1, false, -1, -1, "V")
			.Col = Col : .Row = .MaxRows : .Value = dblSum	
			Call SetW6(Row)
	End Select
	End With
End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("0001011111") 

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

	' -- 커서 이동시 하단 변경 (_Click이벤트에서 이동: 200603)
	Dim iSeqNo, IntRetCD, iLastRow
	
	If lgOldRow = Row  Then Exit Sub
	
	ggoSpread.Source = frm1.vspdData
  
	If Row = frm1.vspdData.MaxRows Then
		iLastRow = ShowRowHidden2(frm1.vspdData2, iSeqNo)
		iLastRow = ShowRowHidden2(frm1.vspdData3, iSeqNo)

	Else
		With frm1.vspdData
			.Col = C_SEQ_NO : .Row = Row : iSeqNo = .Value
				
			' 하위 그리드 표시루틴'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			frm1.vspdData2.SetActiveCell C_W7, iLastRow
				
			If iLastRow = 0 Then 
				Call InsertRow2Detail2(iSeqNo)
				iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			End If

			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			frm1.vspdData3.SetActiveCell C_W17, iLastRow
	
			If iLastRow = 0 Then 
				Call InsertRow2Detail3(iSeqNo)
				iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			End If

			.focus
		End With	
	End If
  
	lgOldRow = Row	: lgOldCol = Col

End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	If Row <> NewRow And (Row > 0 And NewRow > 0)  Then
		window.setTimeout "vbscript:vspdData_Click " & NewCol & "," & NewRow & " ", 500
	End If
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
	Dim dblSum, dblW10, dblW11
	
	With Frm1.vspdData2
	.Row = Row
	.Col = Col

	If .CellType = parent.SS_CELL_TYPE_FLOAT Then
		If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
		   .text = .TypeFloatMin
		End If
	End If
		
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	
    Select Case Col
		Case C_W9, C_W10, C_W11, C_W12, C_W13, C_W14, C_W15
			.Col = C_W10	: dblW10 = UNICDbl(.value)
			.Col = C_W11	: dblW11 = UNICDbl(.value)
			If dblW11 < dblW10 And (dblW11 <> 0 And dblW10 <> 0) Then
				Call DisplayMsgBox("WC0010", "X", GetGrid(frm1.vspdData2, C_W10, -999), GetGrid(frm1.vspdData2, C_W11, -999)) 
				.Row = Row
				.Col = Col	: .value = 0
				frm1.vspdData2.focus
			End If
		
			dblSum = ufn_FncSumSheet(frm1.vspdData2, Col)
			Call SetW12(Row)			' 진행률 계산 
			Call SetW13(Row)			' 익금산입액 계산 
			Call SetW16(Row)			' 조정액 계산 
			dblSum = ufn_FncSumSheet(frm1.vspdData2, C_W16)
			
			Call SetW4_W5				' 상단 그리드 반영 

    End Select

	End With
	lgChgFlg = True ' 데이타 변경 
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

End Sub

' -- 현재 보이는 부모순번에 일치하는 모든행의 섬을 구한다.
Function ufn_FncSumSheet(Byref pGrid, Byval pCol)
	Dim dblSeqNo, iRow, iMaxRows, dblSum
	
	With pGrid
		.ReDraw = False
		iMaxRows = .MaxRows
		.Col = C_SEQ_NO	: dblSeqNO = UNICDbl(.value)
		
		For iRow = 1 To iMaxRows
			.Col = C_SEQ_NO : .Row = iRow
			If UNICDbl(.value) = dblSeqNo Then	' 부모순번이 같은 행 
				.Col = C_CHILD_SEQ_NO	
				If UNICDbl(.value) < SUM_SEQ_NO Then
					.Col = pCol	: dblSum = dblSum + UNICDbl(.value)
				ElseIf UNICDbl(.value) = SUM_SEQ_NO Then
					.Col = pCol	: .Value = dblSum	' 합계행에 데이타 출력 
					ggoSpread.UpdateRow .Row
					ufn_FncSumSheet = dblSum
					Exit For
				End If
			End If
		Next
		.ReDraw = True
	End With
End Function

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
		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.UpdateRow Row
	
    Select Case Col
		Case C_W19, C_W20
			dblSum = ufn_FncSumSheet(frm1.vspdData3, Col)
			
			Call SetW4_W5()
			
    End Select

	End With
	lgChgFlg = True ' 데이타 변경 
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 3
	ggoSpread.Source = Frm1.vspdData3
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
    
    If lgChgFlg = False Then
    
		If ggoSpread.SSCheckChange = False Then
		    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		    Exit Function
		End If
		
	End If
	
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
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

    Call InsertRow2Head
    'Call InsertRow2Detail(1)
    
    Call SetToolbar("1100111100000111")

	Call InitData
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
				Else
					lDelRows = ggoSpread.EditUndo
				End If
				
			End With
		CAse 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.EditUndo
				End If
			End With    
 		CAse 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow2(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.EditUndo
				End If
			End With     
	End Select
  
	lgChgFlg = True                                                '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, iLastRow, sW2_NM

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
    
    ' 멀티 행을 지원하지 않는다. 하위그리드가 상위그리드에 물려있어 복잡함 
	imRow = CInt(pvRowCnt)
		 
	' 첫행일 경우 합계까지 넣는 루틴 
	If frm1.vspdData.MaxRows = 0 Then
		Call InsertRow2Head
		Call SetToolbar("1100111100000111")
		Exit Function
	End If
		
	Select Case lgCurrGrid
		Case 1	' 1번 그리드 
		
		With frm1.vspdData
			
		.focus
		ggoSpread.Source = frm1.vspdData
			
		iRow = .ActiveRow	' 현재행 
			
		.ReDraw = False
			
		If iRow = .MaxRows Then
			ggoSpread.InsertRow iRow-1 , imRow 
			SetSpreadColor iRow, iRow+imRow	' 그리드 색상변경 
			iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow)
		Else
			ggoSpread.InsertRow ,imRow
			SetSpreadColor iRow+1, iRow+imRow	' 그리드 색상변경 
			iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow+1)
		End If	

		.ReDraw = True	
						
		' 하위 그리드 표시루틴'
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
		frm1.vspdData2.SetActiveCell C_W7, iLastRow
			
		If iLastRow = 0 Then Call InsertRow2Detail2(iSeqNo)

		iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
		frm1.vspdData3.SetActiveCell C_W7, iLastRow
			
		If iLastRow = 0 Then Call InsertRow2Detail3(iSeqNo)
			
		Call vspdData_Click(.Col, .ActiveRow)
		
		frm1.vspdData.SetActiveCell C_W1_NM, .ActiveRow
		End With
		
	Case 2	' 2번 그리드 
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_W2_NM
		sW2_NM = frm1.vspdData.value
		
		With frm1.vspdData2
		.focus
		ggoSpread.Source = frm1.vspdData2
			
		iRow = .ActiveRow
		.ReDraw = False
					
		iSeqNo = GetGRid(frm1.vspdData, C_SEQ_NO, frm1.vspdData.ActiveRow)	' 부모그리드의 위치순번 

		If .MaxRows = 0 Then
			Call InsertRow2Detail2(iSeqNo)			
		ElseIf iRow = .MaxRows And iRow > 0 Then
			ggoSpread.InsertRow iRow-1 , imRow 
			SetSpreadColorDetail2 iRow-1
			MaxSpreadVal2 frm1.vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow , iSeqNo
			.Row = iRow
		Else
			ggoSpread.InsertRow ,imRow
			SetSpreadColorDetail2 iRow+1
			MaxSpreadVal2 frm1.vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
			.Row = iRow+1
		End If
		.Col = C_W2_NM2 : .value = sW2_NM
		.ReDraw = True
		End With
		
	Case 3	' 3번 그리드 
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_W2_NM
		sW2_NM = frm1.vspdData.value
		
		With frm1.vspdData3
		.focus
		ggoSpread.Source = frm1.vspdData3

		iRow = .ActiveRow
		.ReDraw = False
		
		iSeqNo = GetGRid(frm1.vspdData, C_SEQ_NO, frm1.vspdData.ActiveRow)	' 부모그리드의 위치순번 
		
		If .MaxRows = 0 Then
			Call InsertRow2Detail3(iSeqNo)
		ElseIf iRow = .MaxRows And iRow > 0 Then
			ggoSpread.InsertRow iRow-1 , imRow 
			SetSpreadColorDetail3 iRow-1
			MaxSpreadVal2 frm1.vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow	, iSeqNo
			.Row = iRow
		Else
			ggoSpread.InsertRow iRow,imRow
			SetSpreadColorDetail3 iRow+1
			MaxSpreadVal2 frm1.vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
			.Row = iRow+1
		End If	
		.Col = C_W2_NM2 : .value = sW2_NM
		.ReDraw = True
		End With
	End Select

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' -- 그리드 좌표의 값 읽기 
Function GetGrid(Byref pGrid, Byval pCol, Byval pRow)
	With pGrid
		.Col = pCol : .Row = pRow : GetGrid = .Value
	End With
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
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				
			End With
		Case 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				lDelRows = ggoSpread.DeleteRow
			End With    
 		Case 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "합계 행은 삭제할 수 없습니다.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				lDelRows = ggoSpread.DeleteRow
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
    If frm1.vspdData.MaxRows > 0 Then
    
		lgIntFlgMode = parent.OPMD_UMODE
		    
		Call SetToolbar("1101111100000111")										<%'버튼 툴바 제어 %>
	
		Call RedrawSumRow
		Call RedrawSumRow2
		Call RedrawSumRow3

		With frm1.vspdData
			.Col = C_SEQ_NO : .Row = 1 : iSeqNo = .Value
				
			' 하위 그리드 표시루틴'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)

			' 하위 그리드 표시루틴'
			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
		End With	
	
		' 세무정보 조사 : 컨펌되면 락된다.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
		If wgConfirmFlg = "Y" Then
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLock -1, -1
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLock -1, -1
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SpreadLock -1, -1
			
			Call SetToolbar("1100000000000111")	
		Else
			Call vspdData_Click(C_W1_NM, 1)
		End If
	Else
		Call SetToolbar("1100110100000111")	
	End If
	lgOldRow =0
	
	frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow , lCol, lGrpCnt, lMaxRows, lMaxCols
    Dim lStartRow, lEndRow , lChkAmt
    Dim strVal, strDel
 
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

    frm1.txtSpread.value      = strDel & strVal
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
					.Col = C_W16 : lChkAmt = .Value
					If lChkAmt = 0 Then
		                                          strVal = strVal & "I"  &  Parent.gColSep ' 무시 추가된 코드 
		            Else
		                                          strVal = strVal & "C"  &  Parent.gColSep	
		            End If
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
					.Col = C_W16 : lChkAmt = .Value
					If lChkAmt = 0 Then
		                                          strVal = strVal & "I"  &  Parent.gColSep ' 무시 추가된 코드 
		            Else
		                                          strVal = strVal & "U"  &  Parent.gColSep
		            End If                                                   
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

    frm1.txtSpread2.value      = strDel & strVal
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
	
    frm1.txtSpread3.value      = strDel & strVal
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
						<a href="vbscript:GetRef">수입금액 조회</A>  
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
						<TABLE <%=LR_SPACE_TYPE_20%> BORDER=0>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">1. 수입금액 조정계산</TD>
                            </TR>
                            <TR HEIGHT=30%>
								<TD WIDTH="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="25" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">2. 수입금액 조정명세</TD>
                            </TR>
                            <TR HEIGHT=60%>
								<TD WIDTH="100%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> BORDER=0>
									<TR HEIGHT=10>
									    <TD WIDTH="100%">&nbsp;&nbsp;&nbsp;가. 작업진행률에 의한 수입금액</TD>
									</TR>
									<TR HEIGHT=60%>
										<TD WIDTH="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="25" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR HEIGHT=10>
									    <TD WIDTH="100%">&nbsp;&nbsp;&nbsp;나. 기타 수입금액</TD>
									</TR>
									<TR HEIGHT=40%>
										<TD WIDTH="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="25" TITLE="SPREAD" id=vaSpread3> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
								</TABLE>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지>수입금액조정계산</LABEL>&nbsp;
				                                 <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><별지>가.작업진행에 의한 수입금액</LABEL>&nbsp;
				                                 <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check3" ><LABEL FOR="prt_check3"><별지>나.기타수입금액</LABEL>&nbsp;</TD>
                </TR>
			
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24"></TEXTAREA>
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

