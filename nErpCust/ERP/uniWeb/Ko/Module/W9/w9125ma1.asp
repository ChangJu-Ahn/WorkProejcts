<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 해외현지법인명세서 
'*  3. Program ID           : w9125ma1
'*  4. Program Name         : w9125ma1.asp
'*  5. Program Desc         : 해외현지법인명세서 
'*  6. Modified date(First) : 2006/01/09
'*  7. Modified date(Last)  : 2007/03
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      :  LEE WOL SAN
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
<SCRIPT LANGUAGE="VBScript"  SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "w9125ma1"
Const BIZ_PGM_ID = "w9125mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "w9125OA1"		' -- 주의 : EBR이 3개라 순차적으로 뒤에 붙여줌 

' -- 서식명 
Dim C_C_CD
Dim C_C_NM
Dim C_C_CD2
Dim C_C_NM2

' -- 컬럼 정보 
Dim C_C01
Dim C_C01_POP
Dim C_C02
Dim C_C02_POP
Dim C_C03
Dim C_C03_POP
Dim C_C04
Dim C_C04_POP
Dim C_C05
Dim C_C05_POP
Dim C_C06
Dim C_C06_POP
Dim C_C07
Dim C_C07_POP
Dim C_C08
Dim C_C08_POP
Dim C_C09
Dim C_C09_POP
Dim C_C10
Dim C_C10_POP
Dim C_C11
Dim C_C11_POP
Dim C_C12
Dim C_C12_POP

' -- 행정보(서식)
Dim C_W6
Dim C_W6_1
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10
Dim C_W11_1
Dim C_W11_2
Dim C_W12
Dim C_W12_1
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16
Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W21


Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	' -- 서식명 
	
	C_C_NM			= 1
	C_C_NM2			= 2

	' -- 컬럼 정보 
	C_C01			= 3
	C_C01_POP		= 4
	C_C02			= 5
	C_C02_POP		= 6
	C_C03			= 7
	C_C03_POP		= 8
	C_C04			= 9
	C_C04_POP		= 10
	C_C05			= 11
	C_C05_POP		= 12
	C_C06			= 13
	C_C06_POP		= 14
	C_C07			= 15
	C_C07_POP		= 16
	C_C08			= 17
	C_C08_POP		= 18
	C_C09			= 19
	C_C09_POP		= 20
	C_C10			= 21
	C_C10_POP		= 22
	C_C11			= 23
	C_C11_POP		= 24
	C_C12			= 25
	C_C12_POP		= 26

	' -- 행정보(서식)
	C_W6			= 1
	C_W6_1			= 2
	C_W7			= 3
	C_W8			= 4
	C_W9			= 5
	C_W10			= 6
	C_W11_1			= 7
	C_W11_2			= 8
	C_W12			= 9
	C_W12_1			= 10
	C_W13			= 11
	C_W14			= 12
	C_W15			= 13
	C_W16			= 14 
	C_W17			= 15
	C_W18			= 16
	C_W19			= 17
	C_W20			= 18
	C_W21			= 19

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
	Dim ret, i
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","5","0")		' -- 금액 5자리 고정 :
	Call AppendNumberPlace("7","3","2")	' -- 금액 5.2자리 고정 :
	Call AppendNumberPlace("8","15","0")	' -- 금액 15자리 고정 : 
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
    
    .EditEnterAction = 2
    
	.ReDraw = false

    .MaxCols = C_C12_POP + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

     ggoSpread.SSSetEdit		C_C_NM,		"", 7
     ggoSpread.SSSetEdit		C_C_NM2,	"", 15
     
	 ggoSpread.SSSetEdit		C_C01,		"01", 20
	 ggoSpread.SSSetButton		C_C01_POP   
	 ggoSpread.SSSetEdit		C_C02,		"02", 20
	 ggoSpread.SSSetButton		C_C02_POP   
	 ggoSpread.SSSetEdit		C_C03,		"03", 20
	 ggoSpread.SSSetButton		C_C03_POP   
	 ggoSpread.SSSetEdit		C_C04,		"04", 20
	 ggoSpread.SSSetButton		C_C04_POP   
	 ggoSpread.SSSetEdit		C_C05,		"05", 20
	 ggoSpread.SSSetButton		C_C05_POP   
	 ggoSpread.SSSetEdit		C_C06,		"06", 20
	 ggoSpread.SSSetButton		C_C06_POP   
	 ggoSpread.SSSetEdit		C_C07,		"07", 20
	 ggoSpread.SSSetButton		C_C07_POP   
	 ggoSpread.SSSetEdit		C_C08,		"08", 20
	 ggoSpread.SSSetButton		C_C08_POP   
	 ggoSpread.SSSetEdit		C_C09,		"09", 20
	 ggoSpread.SSSetButton		C_C09_POP   
	 ggoSpread.SSSetEdit		C_C10,		"10", 20
	 ggoSpread.SSSetButton		C_C10_POP   
	 ggoSpread.SSSetEdit		C_C11,		"11", 20
	 ggoSpread.SSSetButton		C_C11_POP   
	 ggoSpread.SSSetEdit		C_C12,		"12", 20
	 ggoSpread.SSSetButton		C_C12_POP   

	ret = .AddCellSpan(C_C_NM, 0 , 2, 1)
	
    ggoSpread.SSSetSplit2(2)
	' 그리드 헤더 합침 정의 

	'Call ggoSpread.SSSetColHidden(C_C05, .MaxCols,True)

	.ReDraw = true

    End With   
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

    Dim IntRetCD1
    Dim iRow, iCol
	' 시부인 구분 
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1090' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData

	   iRow = C_W21
	   
		For iCol = C_C01 To C_C12 Step 2
			Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), iCol, iRow, iRow)
		Next		
	End If
	  	  
		  
End Sub



' Col, Row1~Row2 까지 콤보를 만든다. : 표준에 없어서 직접 정의함 
Sub Spread_SetCombo(pVal, pCol1, pRow1, pRow2)

	With  frm1.vspdData

		.BlockMode = True
		.Col = pCol1	: .Col2 = pCol1
		.Row = pRow1	: .Row2 = pRow2
		.CellType = 8	'SS_CELL_TYPE_COMBOBOX

		.TypeComboBoxList = pVal	

		.TypeComboBoxEditable = False
		.TypeComboBoxMaxDrop = 3
		' Select the first item in the list
		'.TypeComboBoxCurSel = 0
		' Set the width to display the widest item in the list
		'.TypeComboBoxWidth = 1 
		.BlockMode = False
	End With

End Sub


Sub SetSpreadLock()
	Dim i
	
    With frm1.vspdData

    .ReDraw = False
    
    ggoSpread.SpreadLock C_C_NM, -1, C_C_NM2

    .ReDraw = True

    End With
End Sub


Sub SetColorGrid(Byval pCol, Byval pBoolean)
	Dim i
	
    With frm1.vspdData

    .ReDraw = False

	For i = C_C01 To C_C12 Step 2
		If .ColHidden = False Then
		
			If pBoolean Then
				ggoSpread.SSSetRequired pCol, C_W6, C_W12
			Else
				ggoSpread.SpreadUnLock pCol, C_W6, pCol, C_W12
			End If
		End If
	Next

    .ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
          
    End Select    
End Sub

Sub InitData()
	Dim iMaxRows, iRow, ret
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	iMaxRows = C_W21 ' 하드코딩되는 행수 
	
	With frm1.vspdData
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow , iMaxRows

		' -- 셀 합침 
		For iRow = C_W6 To C_W14
			ret = .AddCellSpan(C_C_NM, iRow , 2, 1)
		Next
		
		ret = .AddCellSpan(C_C_NM, C_W21 , 2, 1)
		
		ret = .AddCellSpan(C_C_NM, C_W15 , 2, 1)
		ret = .AddCellSpan(C_C_NM, C_W17 , 1, 2)
		ret = .AddCellSpan(C_C_NM, C_W19 , 1, 2)
		
		ret = .AddCellSpan(C_C_NM, C_W11_1 , 2, 2)
		
		' -- 높이 재지정 
		.Rowheight(C_W7) = 20		' 현지법인명 
		.Rowheight(C_W9) = 30		' 현지법인소재지 
		.Rowheight(C_W21) = 15
		
		.Redraw = True
		
		Call InitData2
		
		Call SetSpreadLock

	End With	
End Sub

 ' -- DBQueryOk 에서도 불러준다.
Sub InitData2()
	Dim iRow  , iCol

	With frm1.vspdData
		.Redraw = False

		iRow = 0
		iRow = iRow + 1 : .Row = iRow    :  .TypeVAlign = 2
		.Col = C_C_NM	: .value = " (9)투 자 국"
		iRow = iRow + 1 : .Row = iRow    :  .TypeVAlign = 2
		.Col = C_C_NM	: .value = "    투자국명"
	  
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (10)현지법인명"				
		
		.BlockMode = True
		.Col = -1	: .Row = C_W7 : .TypeEditMultiLine = true : .TypeTextWordWrap = True
		.BlockMode = False
		
		
		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (11)현지법인고유번호"
        
        
        iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (12)현지법인소재지"	

		.BlockMode = True
		.Col = -1	: .Row = C_W9 : .TypeEditMultiLine = true : .TypeTextWordWrap = True
		.BlockMode = False
		
		iRow = iRow + 1 : .Row = iRow    :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (13)설립일자"

		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (14)사업연도"

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_C_NM	: .value = ""

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (15)업종"
		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "     업종상세"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (16)직원수"
		
		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (17)모법인파견직원수"
	    
		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .TypeEditMultiLine = true 	:.TypeHAlign = 2	: .value = "해외현지법인 관련 투자현황" & vbCrLf & "투자" 
		.foreColor = rgb(0,0,255)
		
		.Col = C_C_NM2	: .value = "(18)지분율"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM2	: .value = "(19)자본금" : .Rowhidden =true

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .TypeEditMultiLine = true 	:.TypeHAlign = 2	: .value = "대부" & vbCrLf & "투자"
		.Col = C_C_NM2	: .value = "(20)대여금" : .Rowhidden =true

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM2	: .value = "(21)수입이자" : .Rowhidden =true

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	:.TypeHAlign = 2	: .value = "청산"
		.Col = C_C_NM2	: .value = "(22)청산일"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM2	: .value = "(23)회수금액"


		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " (21)재무상황표 제출대상 여부"
        .RowHIdden = true

		For iCol = C_C01 To C_C12 Step 2
			ggoSpread.SSSetEdit		iCol,	"", 20, , C_W6, 3
			ggoSpread.SSSetEdit		iCol,	"", 20, , C_W6_1
			ggoSpread.SSSetEdit		iCol,	"", 20, , C_W7, 60
			ggoSpread.SSSetMask		iCol,	"", 20, 2, "9999-9999", C_W8
			ggoSpread.SSSetEdit		iCol,	"", 20, , C_W9, 70
			ggoSpread.SSSetDate		iCol,		"",	20,		2,		Parent.gDateFormat,	C_W19
			ggoSpread.SSSetEdit		iCol,	"", 20, , C_W12, 7
			ggoSpread.SSSetEdit		iCol,	"", 20, , C_W12_1
			
			ggoSpread.SSSetDate		iCol,		"",	20,		2,		Parent.gDateFormat,	C_W10
			ggoSpread.SSSetDate		iCol,		"",	20,		2,		Parent.gDateFormat,	C_W11_1
			ggoSpread.SSSetDate		iCol,		"",	20,		2,		Parent.gDateFormat,	C_W11_2
			
			ggoSpread.SSSetFloat	iCol, "", 20, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , , , , C_W13
			ggoSpread.SSSetFloat	iCol, "", 20, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , , , , C_W14
			ggoSpread.SSSetFloat	iCol, "", 20, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , , , , C_W15
			
			ggoSpread.SSSetFloat	iCol, "", 20, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , , , , C_W16
			ggoSpread.SSSetFloat	iCol, "", 20, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , , , , C_W17
			ggoSpread.SSSetFloat	iCol, "", 20, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , , , , C_W18
			ggoSpread.SSSetFloat	iCol, "", 20, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , , , , C_W20
			
			ggoSpread.SSSetProtected		iCol		, C_W6_1	,C_W6_1				
			ggoSpread.SSSetProtected		iCol		, C_W12_1	,C_W12_1
			ggoSpread.SSSetProtected		iCol		, C_W15	,C_W15
			.Row=C_W15: .BackColor = rgb(196,253,245)
		Next

		
		for iRow = C_W6 to  .maxrows 	
			.Row = iRow

		    Select Case iRow
				Case C_W6, C_W12, C_W12+4 '200703 
				
				
				Case Else
					For iCol = C_C01_POP To C_C12_POP Step 2	' -- 팝업버튼을 제거 
						.Col = iCol	: .CellType = 1
						ggoSpread.SSSetProtected	iCol, iRow, iRow
						
					
					Next
				
		    End Select
		Next    
		
		Call InitSpreadComboBox()	
		
		' -- 컬럼헤더에 위치 표기 및 히든 
		.Row = 0
		.Col = C_C01	: .Value = "01"
		.Col = C_C02	: .Value = "02"
		.Col = C_C03	: .Value = "03"
		.Col = C_C04	: .Value = "04"
		.Col = C_C05	: .Value = "05"	: .ColHidden = True
		.Col = C_C06	: .Value = "06"	: .ColHidden = True
		.Col = C_C07	: .Value = "07"	: .ColHidden = True
		.Col = C_C08	: .Value = "08"	: .ColHidden = True
		.Col = C_C09	: .Value = "09"	: .ColHidden = True
		.Col = C_C10	: .Value = "10"	: .ColHidden = True
		.Col = C_C11	: .Value = "11"	: .ColHidden = True
		.Col = C_C12	: .Value = "12"	: .ColHidden = True
		
		For iCol = C_C05_POP To C_C12_POP Step 2
			.Col = iCol	: .ColHidden = True
		Next
		

	End With
End Sub

Function GetValue4Grid(Byval pCol, Byval pRow)
	With frm1.vspdData
		.Col = pCol : .Row = pRow : GetValue4Grid = .Value
	End With
End Function

Function GetText4Grid(Byval pCol, Byval pRow)
	With frm1.vspdData
		.Col = pCol : .Row = pRow : GetText4Grid = .Text
	End With
End Function

Sub SetText4Grid(Byval pCol, Byval pRow, Byval pData)
	With frm1.vspdData
		.Col = pCol : .Row = pRow : .Text = pData
	End With
End Sub

Sub SetValue4Grid(Byval pCol, Byval pRow, Byval pData)
	With frm1.vspdData
		.Col = pCol : .Row = pRow : .Value = pData
	End With
End Sub

' -- mb 단에서 05 이상 데이타 존재시 사용함 
Sub ShowColumn(Byval pCol)
	With frm1.vspdData
		.Col = pCol	: .ColHidden = False
		.Col = .Col + 1 : .ColHidden = False
	End With	
End Sub
'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 
	
End Function

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	
	'Call InitData()	
    Call FncQuery
    call AutoSum()
    
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
'===============


Sub BtnIntCol()
Dim iCol
    With frm1.vspdData
        for iCol = C_C05  to C_C12 Step 2
            .col = iCol
            If .ColHidden Then
				.ColHidden = False
				.Col = .Col + 1
				.ColHidden = False
				Exit Sub
			End If
        Next
    
'		  Call InitData2()	

    
   	End With
End sub



'===========================================================================

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere, Byval iRow )
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd
	
	If IsOpenPop = True Then Exit Function
	
	if iRow = C_W12 then
		arrParam(0) = "업종번호"								' 팝업 명칭 
		arrParam(1) = "tb_std_income_rate" 								' TABLE 명칭 
		arrParam(2) = Trim(strCode)										' Code Condition
		arrParam(3) = ""												' Name Cindition
		If frm1.txtFISC_YEAR.text >= "2006" Then							' -- 2006년 중간예납부터 표준소득율코드 바뀜					
			arrParam(4) = " ATTRIBUTE_YEAR = '2005'"					' Where Condition

			arrField(0) = "STD_INCM_RT_CD"									' Field명(0)
			arrField(1) = "MIDDLE_NM"									' Field명(1)
			arrField(2) = "DETAIL_NM"									' Field명(1)
			arrField(3) = ""									' Field명(1)
					
			arrHeader(0) = " 번호"									' Header명(0)
			arrHeader(1) = "업태"									' Header명(1)
			arrHeader(2) = "업종"									' Header명(1)
			arrHeader(3) = ""									' Header명(1)

		Else
			arrParam(4) = " ATTRIBUTE_YEAR = '2003'"

			arrField(0) = "STD_INCM_RT_CD"									' Field명(0)
			arrField(1) = "BUSNSECT_NM"									' Field명(1)
			arrField(2) = "DETAIL_NM"									' Field명(1)
			arrField(3) = "FULL_DETAIL_NM"									' Field명(1)
					
			arrHeader(0) = " 번호"									' Header명(0)
			arrHeader(1) = "업태"									' Header명(1)
			arrHeader(2) = "업종"									' Header명(1)
			arrHeader(3) = "업종상세"									' Header명(1)

		End If
		arrParam(5) = "업종"									' 조건필드의 라벨 명칭 


     Elseif iRow = C_W6 then
               
          arrParam(0) = "투자국"								' 팝업 명칭 
	  	arrParam(1) = "B_COUNTRY"								' TABLE 명칭 
	  	arrParam(2) = Trim(strCode)										' Code Condition
	  	arrParam(3) = ""												' Name Cindition
	  	arrParam(4) = ""
	  	arrParam(5) = "코드"									' 조건필드의 라벨 명칭 
            
	  	arrField(0) = "COUNTRY_CD"									' Field명(0)
	  	arrField(1) = "COUNTRY_NM"									' Field명(1)
	  	arrField(2) = ""									' Field명(1)
	  	arrField(3) = ""									' Field명(1)
			
	  	arrHeader(0) = "코드"									' Header명(0)
	  	arrHeader(1) = "국가명"									' Header명(1)
	  	arrHeader(2) = ""									' Header명(1)
	  	arrHeader(3) = ""									' Header명(
               
     end if 				
	
	IsOpenPop = True
			
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=750px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere ,iRow)
	End If
End Function



'======================================================================================================
'   FUNCTION Name : 해외현지법인 관련 투자현황 
'   FUNCTION Desc : 
'=======================================================================================================

Function  OpenPopUp2(Byval strCode, Byval iWhere, Byval iRow )
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6, Param7

	If IsOpenPop = True Then Exit Function
    
   
		
    ggoSpread.source = frm1.vspdData
     frm1.vspdData.col=frm1.vspdData.ActiveCol -1 :  frm1.vspdData.Row=1

    if frm1.vspdData.text ="" then
		call DisplayMsgBox("971012","X", "투자국","X")
		frm1.vspdData.action =0
		Exit Function
		
    end if
    
     If Not chkField(Document, "2") Then                             '⊙: Check contents area
	   Exit Function
	End If
		
		 
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If   

	IsOpenPop = True
	Param1 =  frm1.txtCO_CD.value 
	Param2 =frm1.txtFISC_YEAR.text
	Param3 = frm1.cboREP_TYPE.value
	Param4 = (frm1.vspdData.ActiveCol)/2-1

	iCalledAspName = AskPRAspName("w9125pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"w9125pa1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6, Param7,"A"), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    
    	
	
	IsOpenPop = False


	If arrRet(0,0) = "" Then
		Exit Function
	Else
		frm1.txtReq_no.value = arrRet(0,0) 
		
			'call vspdData_Change ("1",frm1.vspdData.ActiveRow)
	End If
 


End Function



'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================



Function SetPopup(Byval arrRet,Byval iWhere , iRow)
	With frm1
		select Case iRow
			Case C_W6
				.vspdData.Col = iWhere-1
				.vspdData.Row = iRow
				.vspdData.Text = arrRet(0)
				.vspdData.Row=iRow +1
				.vspdData.Text =arrRet(1)
			Case C_W12
				.vspdData.Col = iWhere-1
				.vspdData.Row = iRow
				.vspdData.Text = arrRet(0)
				.vspdData.Row=iRow +1
				
				If frm1.txtFISC_YEAR.text >= "2006" Then							' -- 2006년 중간예납부터 표준소득율코드 바뀜					
					.vspdData.Text =arrRet(2)
				Else
					.vspdData.Text =arrRet(3)
				End If
		End Select	
	
	End With
	''
	Call vspdData_Change(iWhere-1,iRow)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow iRow
	
	lgBlnFlgChgValue = True

End Function


Sub vspdData_Change(ByVal Col , ByVal Row )
	dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6   ,IntRetCD, i, blnData, sWhere
	Dim dblSum, dblCol(1)
	
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
    
	lgBlnFlgChgValue = True
	
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	With frm1.vspdData

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	
	If Trim(frm1.vspdData.text) <> "" Then
		Select Case ROW
			Case C_W6						
			    
			       .Row = Row
			       .Col = Col

			        IntRetCD =  CommonQueryRs(" COUNTRY_CD,COUNTRY_NM   ","B_COUNTRY"," COUNTRY_CD = '" & Trim(frm1.vspdData.text) & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
						If IntRetCD = False Then
						    Call  DisplayMsgBox("970000","X","투자국","X")                         '☜ : 입력된자료가 있습니다.
						    .col = col :					    .row = row
						    frm1.vspdData.Text = ""
						    .row = row+1
						    frm1.vspdData.Text = ""
						Else
                             frm1.vspdData.Text = UCASE(Replace(lgF0,chr(11),""))
                             .Row=Row+1
                             frm1.vspdData.Text = UCASE(Replace(lgF1,chr(11),""))
						 
						End If

			Case C_W12
			             .Row = Row
						 .Col = Col

							If frm1.txtFISC_YEAR.text >= "2006" Then	' -- 2006.07.07 수정 
								sWhere = " AND ATTRIBUTE_YEAR = '2005' " 
							Else
								sWhere = " AND ATTRIBUTE_YEAR = '2003' " 
							End If
						 
						 IntRetCD =  CommonQueryRs(" Top 1 STD_INCM_RT_CD,FULL_DETAIL_NM   ","tb_std_income_rate"," STD_INCM_RT_CD = '" & Trim(frm1.vspdData.text) & "'" & sWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
						If IntRetCD = False Then
						    Call  DisplayMsgBox("970000","X","업종","X")                         '☜ : 입력된자료가 있습니다.
						    .col = Col
						    .row = row
						    frm1.vspdData.Text = ""
						    .row = row+1
						    frm1.vspdData.Text = ""
						   
						Else           
							frm1.vspdData.Text = UCASE(Replace(lgF0,chr(11),""))
                             .Row=Row+1
                             frm1.vspdData.Text = UCASE(Replace(lgF1,chr(11),""))
						End If
			
			Case C_W10	' -- 설립일자가 사업연도 시작/종료보다 크면 에러 
				If CompareDateByFormat(GetText4Grid(Col, Row), GetText4Grid(Col, C_W11_1), GetText4Grid(C_C_NM, Row), GetText4Grid(C_C_NM, C_W11_1), _
				    	               "970024", parent.gClientDateFormat, parent.gComDateType, true) = False Then
					.Col = Col	: .Row = Row : .Text = ""
				   frm1.vspdData.focus
				   Exit Sub
				End If

			Case C_W11_1, C_W11_2	' -- 설립일자가 사업연도 시작/종료보다 크면 에러 
				If CompareDateByFormat(GetText4Grid(Col, C_W11_1), GetText4Grid(Col, C_W11_2), GetText4Grid(C_C_NM, C_W11_1) & " 시작일", GetText4Grid(C_C_NM, C_W11_1) & " 종료일", _
				    	               "970024", parent.gClientDateFormat, parent.gComDateType, true) = False Then
					.Col = Col	: .Row = Row : .Text = ""
				   frm1.vspdData.focus
				   Exit Sub
				End If

				If CompareDateByFormat(GetText4Grid(Col, C_W10), GetText4Grid(Col, C_W11_1), GetText4Grid(C_C_NM, C_W10) , GetText4Grid(C_C_NM, C_W11_1) & " 시작일", _
				    	               "970024", parent.gClientDateFormat, parent.gComDateType, true) = False Then
					.Col = Col	: .Row = Row : .Text = ""
				   frm1.vspdData.focus
				   Exit Sub
				End If

				If CompareDateByFormat(GetText4Grid(Col, C_W10), GetText4Grid(Col, C_W11_2), GetText4Grid(C_C_NM, C_W10) , GetText4Grid(C_C_NM, C_W11_1) & " 종료일", _
				    	               "970024", parent.gClientDateFormat, parent.gComDateType, true) = False Then
					.Col = Col	: .Row = Row : .Text = ""
				   frm1.vspdData.focus
				   Exit Sub
				End If
								
		End Select
	End If
	
	
	Select Case Row
		Case C_W6, C_W7, C_W8, C_W9, C_W10, C_W11_1, C_W11_2, C_W12
			' -- 필수입력 체크 
			blnData = False
				
			For i = C_W6 To C_W12
				.Col = Col	: .Row = i
					
				if i<>2 then 	
					If Trim(.Text) <> "" Then	 blnData = True
				end if
				'msgbox  Trim(.Text)
			Next
				
			If blnData Then
				ggoSpread.SSSetRequired		Col		, C_W6	,C_W12
				'ggoSpread.SSSetRequired		Col		, C_W21	,C_W21
				ggoSpread.SSSetProtected		Col		, C_W6_1	,C_W6_1
				ggoSpread.SSSetProtected		Col		, C_W12_1	,C_W12_1
				ggoSpread.SSSetProtected		Col		, C_W15	,C_W15
				.Row=C_W15: .BackColor = rgb(196,253,245)
				.Row=C_W15+1: .BackColor = rgb(255,0,0)
				
			Else
				ggoSpread.SpreadUnLock		Col		,-1	, Col

			End If
			
			
				
	End Select
		
	End With
End Sub

Sub ChkRequired()
	Dim iCol, iRow, blnData
	
	With frm1.vspdData
	
	For iCol = C_C01 To C_C12 Step 2
		.Col = iCol

		blnData = False
				
		For iRow = C_W6 To C_W12
			.Row = iRow
			If Trim(.Text) <> "" Then blnData = True
		Next
				
		If blnData Then
			ggoSpread.SSSetRequired		iCol		, C_W6	,C_W12
			'ggoSpread.SSSetRequired		iCol		, C_W21	,C_W21
			ggoSpread.SSSetProtected		iCol		, C_W6_1	,C_W6_1
			ggoSpread.SSSetProtected		iCol		, C_W12_1	,C_W12_1
			ggoSpread.SSSetProtected		iCol		, C_W15	,C_W15
			.Row=C_W15: .BackColor = rgb(196,253,245)

		Else
			ggoSpread.SpreadUnLock		iCol		,-1	, iCol
		End If
		
	Next
	
	End With
End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
    
End Sub

Sub vspdData_KeyDown(KeyCode, shift)


	With frm1.vspdData
	
	if .ActiveRow=13 then
		call DisplayMsgBox("","X", "팝업버튼을 이용하세요","X")
		.Row=13: .text =""
    exit sub
    end if
    
	
    If KeyCode = 46 Then	' Del
		.Col = .ActiveCol	: .Row = .ActiveRow : .Text = ""
		Call vspdData_Change(.ActiveCol, .ActiveRow)
    End If
    End With
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

' -- 그리드1 팝업 버튼 클릭 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim sCode, sCode2
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
	
		Select Case Col
			Case C_C01_POP, C_C02_POP , C_C03_POP, C_C04_POP, C_C05_POP, C_C06_POP, C_C07_POP, C_C08_POP, C_C09_POP ,C_C10_POP, C_C11_POP, C_C12_POP
				.vspdData.Col = Col - 1
				.vspdData.Row = Row
				
				sCode = UCase(Trim(.vspdData.Text))
				if Row<= 10 then
					Call OpenPopup(sCode, Col, Row)
				else
					Call OpenPopup2(sCode, Col, Row)
				end if
				
		End Select
        Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   
        
    End With
    
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

Sub txtW15_Change()
	call AutoSum()
	lgBlnFlgChgValue=True	

End Sub
Sub txtW16_Change()
	lgBlnFlgChgValue=True
	call AutoSum()
End Sub

Sub txtW17_Change()
	lgBlnFlgChgValue=True
	'call AutoSum()
End Sub
Sub txtW18_Change()
	lgBlnFlgChgValue=True

End Sub

function AutoSum()
	dim w15
	dim w16
	dim w17
	
	dim w17_1

	w15 = cDbl(frm1.txtw15.value)
	w16 = cDbl(frm1.txtw16.value)
	w17 = cDbl(frm1.txtw17.value)
	
	
	w17_1 = cDbl(w15+w16-w17)
	
	 frm1.txtw17_1.value = w17_1
	
end Function


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

    Call SetToolbar("1100100000000111")

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
    If lgBlnFlgChgValue = True Then
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
   
    
    Call InitData()

    CALL DBQuery()
    
End Function

' -- 컬럼 헤더 리턴 
Function GetColName(Byval pCol)
	With frm1.vspdData
		.Col = pCol	: .Row = -999
		GetColName = .Value
	End With
End Function

Function FncSave() 
    Dim blnChange, dblSum, iCol, iRow
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                             '⊙: Check contents area
	   Exit Function
	End If
		
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If
dim tmp
	With frm1.vspdData
	
	For iCol = C_C01 To C_C12 Step 2
		.Col = iCol
		If Not .ColHidden Then
			For iRow = C_W6 To C_W21
				
				if iRow = 1 then
					tmp =Trim(GetSpreadText(frm1.vspdData,iCol,0,"X","X"))
				
				end if
				.Row = iRow
							
				If .BackColor = ggoSpread.RequiredColor Then
					Select Case iRow
						Case C_W6,C_W6_1, C_W7, C_W8, C_W9, C_W10, C_W11_1, C_W11_2, C_W12, C_W12_1,C_W21
							' -- 필수입력 체크 
							If Trim(.Text) = "" Then
								.Col = C_C_NM
								Call DisplayMsgBox("970021", "X", .Text, "X")                          <%'No data changed!!%>
								Call SetFocusToDocument("M")
								.focus
								.Col = iCol	: .Row = iRow	: .Action = 0
								Exit Function
							End If
					End Select
				End If
			 
				Select case iRow
					case C_W8 '현지 법인고유번호 체크 
		
						if .text <> "" then
						
							if len(Replace(trim(.text),"-",""))<>8 then
								UNIMsgBox "법인고유번호 길이를 확이 하십시요.", 48, "uniERPII"
								exit function
								
							end if
							
							if instr("128", left(.text,1) ) >  0 then
							else
								UNIMsgBox "법인고유번호 첫글자는  1,2,8 중에 하나 이어야 합니다.", 48, "uniERPII"
								exit function
							end if
							
							if Right(.text,4) <>  "00" & tmp then
					
								UNIMsgBox "법인고유번호 일련번호를 확인 하십시요.", 48, "uniERPII"
								exit function
							end if
							
								
							'if trim(mid(.text,2,3)) <> Trim(GetSpreadText(frm1.vspdData,iCol,1,"X","X"))   then
							'	UNIMsgBox "법인고유번호 두번째이후 3글자는  투자국값 이어야 합니다.", 48, "uniERPII"
							'		exit function
							'end if
						
						end if
					
				end Select
			Next
		End If
	Next
	
	End With
	
'	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
'	      Exit Function
'	End If    
	
'	If blnChange = False Then
 '       Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
  '      Exit Function
	'End If
	

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

End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
           
    	lDelRows = ggoSpread.DeleteRow
    End With
    
    lgBlnFlgChgValue = True
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
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function FncDelete()
Dim iRow 
Dim IntRetCd


    'frm1.vspdData.AddSelection C_W6, -1, C_W6, -1

    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("800442", parent.VB_YES_NO, "X", "X")			    <%'%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    'Call FncDeleteRow
       
    'Call FncSave
     If DbDelete= False Then
       Exit Function
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              
    
   lgBlnFlgChgValue = True
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
        strVal = strVal	& "&txtTmp="	& C_REVISION_YM  
			
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim iCol, i
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	
	If lgIntFlgMode <> parent.OPMD_UMODE  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE

		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg <>"Y" Then
			
			Call ChkRequired
			
			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>

		Else
			
		End If
		For i=1 to frm1.vspdData.MaxRows
			frm1.vspdData.Col=0
			frm1.vspdData.Row=i
			frm1.vspdData.Text=""
		
		Next

	Else
	
		Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
	End If

	lgBlnFlgChgValue = False
    
	'Call SetSpreadLock(TYPE_1)
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
    Dim strVal, strDel, lMaxRows, lMaxCols, arrVal(12)
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols : lGrpCnt = 0
		
		
		
		For lCol = C_C01 To C_C12 Step 2
			.Col = lCol

			If lgIntFlgMode = parent.OPMD_CMODE  Then
				strVal = "C"  &  Parent.gColSep
			Else
				strVal = "U"  &  Parent.gColSep
			End If
			.Row = 0
			strVal = strVal & Trim(.Text)  &  Parent.gColSep ' -- 컬럼번호 
			
			For lRow = 1 To .MaxRows
               .Row = lRow

				Select Case lRow
					Case C_W10, C_W11_1, C_W11_2, C_W19	' -- 날짜형 
						strVal = strVal & Trim(.Text) &  Parent.gColSep 
					
					Case C_W21
						If Trim(.Text) = "예" Then
							strVal = strVal & "Y" &  Parent.gRowSep
						Else
							strVal = strVal & "N" &  Parent.gRowSep 
						End If 
					Case C_W6_1, C_W12_1	
					Case C_W15
						strVal = strVal & Trim(frm1.txtW15.value) &  Parent.gColSep 
					Case C_W16
						strVal = strVal & Trim(frm1.txtW16.value) &  Parent.gColSep 
					Case C_W17
						strVal = strVal & Trim(frm1.txtW17.value) &  Parent.gColSep 
					Case C_W18
						strVal = strVal & Trim(frm1.txtW18.value) &  Parent.gColSep 			
					Case Else	' -- 숫자/문자형 
						strVal = strVal & Trim(.Value) &  Parent.gColSep 
				End Select
			
			Next
		
			arrVal(lGrpCnt) = strVal
			lGrpCnt = lGrpCnt + 1	
			  
        Next
        

        frm1.txtSpread.value        =  Join(arrVal, "")
		frm1.txtMode.value        =  Parent.UID_M0002
		
		

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
	Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
    Call MainQuery()
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	DbDelete = False			                                                 '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	With Frm1
		.txtMode.value        =  parent.UID_M0003                                '☜: Delete
	End With

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call fncNew()
End Function
function FncBtnPrint1(strPrintType)
	dim sWhere,sMaxSeq,vArrSeq
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE
	Dim StrUrl  , i

	Dim intCnt,IntRetCD
	
	sWhere = "CO_CD=" & FilterVar("<%=wgCO_CD%>", "''", "S") & vbCrLf
	sWhere = sWhere & " AND FISC_YEAR=" & FilterVar("<%=wgFISC_YEAR%>", "''", "S") & vbCrLf
	sWhere = sWhere & " AND REP_TYPE=" & FilterVar("<%=wgREP_TYPE%>", "''", "S") & vbCrLf
	sWhere = sWhere & " AND ISNULL(A.W6,'') <> ''"

	if  CommonQueryRs("distinct  seq_no "," TB_A125 A ",sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
		
	else
		lgF0 = "1" & chr(11) '빈화면을 출력할 수 있도록 함.
	
	end if

	vArrSeq = split(lgF0,chr(11))
	
	for i=0 to uBound(vArrSeq)-1
		StrUrl=""
			Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE) 
			StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
			StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
			StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE
			StrUrl = StrUrl & "|varseq_no|"       & vArrSeq(i)

			 ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")

			 if  strPrintType = "VIEW" then
			 Call FncEBRPreview(ObjName, StrUrl)
			 else
				If document.all("EBAction") is Nothing Then
					Dim pObj , pHTML
					
					pHTML = "<FORM NAME=EBAction TARGET=MyBizASP METHOD=POST>" & _
					"	<INPUT TYPE=HIDDEN NAME=uname>" & _
					"	<INPUT TYPE=HIDDEN NAME=dbname>" & _
					"	<INPUT TYPE=HIDDEN NAME=filename>" & _
					"	<INPUT TYPE=HIDDEN NAME=condvar>" & _
					"	<INPUT TYPE=HIDDEN NAME=date>	" & _
					"</FORM>" 

					Set pObj = document.all("MousePT")
					Call pObj.insertAdjacentHTML("afterBegin", pHTML)
				End If
			 
				Call FncEBRPrint(EBAction,ObjName,StrUrl)
			 end if	
     	next 
end function
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
					<TD WIDTH=* align=right><BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnIntCol()"   Flag=1>열추가</BUTTON></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD ALIGN=RIGHT>단위 : 원 
								</TD>
							</TR>
						     <TR>
								   <TD align=right>
							
											
											 <TABLE width="100%" bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   <TR>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER rowspan="2">(4)전기말 가동 법인수</TD>
									       <TD CLASS="TD51" width="40%" ALIGN=CENTER colspan="2">당기 투자 변동사항</TD>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER rowspan="2">(7)당기발가동법인수<BR>((4)+(5)-(6))</TD>
								           <TD CLASS="TD51" width="20%" ALIGN=CENTER rowspan="2">(8)재무상황표<BR>제출법인수</TD>
									  </TR>
									  <TR>
						
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER>(5)신설법인수</TD>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER>(6)청산(폐업) 법인수</TD>
						
									  </TR>
									  <TR>
											<TD CLASS="TD61" width="20%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW15" name=txtW15 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="11X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="20%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW16" name=txtW16 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="11X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<input name="txtW2_1_1" tag="14XZ0" type="hidden">										    </TD>
											<TD CLASS="TD61" width="20%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW17" name=txtW17 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="11X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="10%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW17_1" name=txtW17_1 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="14X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD61" width="20%" align=center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW18" name=txtW18 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="11X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
								  </table>
								  							   
									</TD>
							</TR>
					
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint1('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
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
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

