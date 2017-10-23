
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 국조제1호정상가격산출명세 
'*  3. Program ID           : w9117ma1
'*  4. Program Name         : w9117ma1.asp
'*  5. Program Desc         : 국조제1호정상가격산출명세 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/02/02
'*  8. Modifier (First)     : 홍지영 
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
<SCRIPT LANGUAGE="VBScript"  SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "w9117ma1"
Const BIZ_PGM_ID = "w9117mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W9117OA1"
Dim C_W1
Dim C_W1_NM
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10
Dim C_W2_POP
Dim C_W3_POP
Dim C_W4_POP
Dim C_W5_POP
Dim C_W6_POP
Dim C_W7_POP
Dim C_W8_POP
Dim C_W9_POP
Dim C_W10_POP
Dim C_W2_CD
Dim C_W3_CD
Dim C_W4_CD
Dim C_W5_CD
Dim C_W6_CD
Dim C_W7_CD
Dim C_W8_CD
Dim C_W9_CD
Dim C_W10_CD


Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
    C_W1				= 1
    C_W1_NM				= 2
    C_W2				= 3
    C_W2_POP			= 4
    C_W2_CD				= 5
    C_W3				= 6
    C_W3_POP			= 7
    C_W3_CD	    		= 8
    C_W4				= 9
    C_W4_POP			= 10
    C_W4_CD				= 11
    C_W5				= 12
    C_W5_POP			= 13
    C_W5_CD				= 14
    C_W6				= 15
    C_W6_POP			= 16
    C_W6_CD				= 17
    C_W7				= 18
    C_W7_POP			= 19
    C_W7_CD				= 20
    C_W8				= 21
    C_W8_POP			= 22
    C_W8_CD				= 23
    C_W9				= 24
    C_W9_POP			= 25
    C_W9_CD				= 26
    C_W10				= 27
    C_W10_POP			= 28
    C_W10_CD			= 29
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
    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
    
	.ReDraw = false


    
    .MaxCols = C_W10_CD + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

     ggoSpread.SSSetEdit		C_W1,		"코드", 10,,,10,1
	 ggoSpread.SSSetEdit		C_W1_NM,	"", 20,,,50,1
	 ggoSpread.SSSetEdit		C_W2,		"1", 20,,,50,1
	 ggoSpread.SSSetCombo		C_W2_CD,	"", 10, 0
	 ggoSpread.SSSetButton		C_W2_POP   
     ggoSpread.SSSetEdit		C_W3,		"2", 20,,,50,1
     ggoSpread.SSSetCombo		C_W3_CD,	"", 10, 0
     ggoSpread.SSSetButton		C_W3_POP
     ggoSpread.SSSetEdit		C_W4,		"3", 20,,,100,1 
     ggoSpread.SSSetCombo		C_W4_CD,	"", 10, 0
     ggoSpread.SSSetButton		C_W4_POP
     ggoSpread.SSSetEdit		C_W5,		"4", 20,,,50,1 
     ggoSpread.SSSetCombo		C_W5_CD,	"", 10, 0
     ggoSpread.SSSetButton		C_W5_POP
     ggoSpread.SSSetEdit		C_W6,		"5", 20,,,50,1 
     ggoSpread.SSSetCombo		C_W6_CD,		"", 10, 0
     ggoSpread.SSSetButton		C_W6_POP
     ggoSpread.SSSetEdit		C_W7,		"6", 20,,,50,1 
     ggoSpread.SSSetCombo		C_W7_CD,		"", 10, 0
     ggoSpread.SSSetButton		C_W7_POP
     ggoSpread.SSSetEdit		C_W8,		"7", 20,,,50,1 
     ggoSpread.SSSetCombo		C_W8_CD,		"", 10, 0
     ggoSpread.SSSetButton		C_W8_POP
     ggoSpread.SSSetEdit		C_W9,		"8", 20,,,50,1 
     ggoSpread.SSSetCombo		C_W9_CD,		"", 10, 0
     ggoSpread.SSSetButton		C_W9_POP
     ggoSpread.SSSetEdit		C_W10,		"9", 20,,,50,1 
     ggoSpread.SSSetCombo		C_W10_CD,		"", 10, 0
     ggoSpread.SSSetButton		C_W10_POP



	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
	Call ggoSpread.SSSetColHidden(C_W5,C_W10_POP,True)
    ggoSpread.SSSetSplit2(2)
	' 그리드 헤더 합침 정의 

	
	.ReDraw = true

	Call SetSpreadLock 

    End With   
End Sub


'============================================  그리드 함수  ====================================


Sub InitSpreadComboBox()

    Dim IntRetCD1
    Dim iRow
	' 시부인 구분 
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1053' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData

	   'iRow = 6
	   iRow = 7
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W2_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W3_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W4_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W5_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W6_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W7_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W8_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W9_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W10_CD, iRow, iRow)
		
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W2, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W3, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W4, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W5, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W6, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W7, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W8, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W9, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W10, iRow, iRow)
		

	End If
		  	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1052' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData

		
		 'iRow = 9
		 iRow = 10
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W2_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W3_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W4_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W5_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W6_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W7_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W8_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W9_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W10_CD, iRow, iRow)
		
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W2, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W3, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W4, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W5, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W6, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W7, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W8, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W10, iRow, iRow)
		

	End If
		  
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1089' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData

		
		 'iRow = 8
		 iRow = 9
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W2_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W3_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W4_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W5_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W6_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W7_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W8_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W9_CD, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF0, chr(11),  chr(9)), C_W10_CD, iRow, iRow)
		
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W2, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W3, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W4, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W5, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W6, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W7, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W8, iRow, iRow)
		Call Spread_SetCombo(Replace(lgF1, chr(11),  chr(9)), C_W10, iRow, iRow)
		

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
    With frm1

    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_W1, -1, C_W1_NM
 
    .vspdData.ReDraw = True

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
	Dim iMaxRows, iRow,i
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	iMaxRows = 11 ' 하드코딩되는 행수 
	With frm1.vspdData
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow , iMaxRows

		iRow = 0
		For i=0 to iMaxRows-1
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow
		next
'		iRow = iRow + 1 : .Row = iRow
'		.Col = C_W1		: .value = iRow
'		
		

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
		.Col = C_W1_NM	: .value = " 국   | (1)법인명(상호)"
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " 외   | 소재국가명"
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = "      | (2)소재국가"
		
		
		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " 특   | (3)대표자(성명)"
        
        
        iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " 수   | 업종명(입력)"
		
		iRow = iRow + 1 : .Row = iRow    :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " 관   | (4)업종"

		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " 계   | (5)신고인과의관계"

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " 자   | (6)소재지(소재)"

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " (7)대상거래"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " (8)정상가격산출방법"
		
		iRow = iRow + 1 : .Row = iRow: .TypeEditMultiLine = true    :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " (9)위의 방법 선택이유"
	    
		
	  for iRow = 1 to  .maxrows 	
			for iCol = 3 to  .maxcols - 1
			         Select Case iRow
			               
			      			  'Case 1 ,3 ,4, 7,  10
			      			  Case 1 ,4 , 8,  11                                          
									Select Case iCol
									          Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
									            .row =iRow : .Col = iCol  : .CellType = 1 :   .TypeVAlign = 2
									            .row =iRow : .Col = iCol+1  : .CellType = 1                     '셀을 Edit로 변경 
								                ggoSpread.SSSetProtected	iCol+1  ,iRow,iRow
								                
								             
										End Select	
											
								 		 				  
			                  'Case 6 , 8  ,9
			                  Case 7 , 9  ,10                                                                      
								 Select Case iCol
						
								        Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
								            
											   .row =iRow : .Col = iCol+1  : .CellType = 1                            
											  	ggoSpread.SSSetProtected	iCol+1  ,iRow,iRow
											  	
											Call ggoSpread.SSSetColHidden(iCol + 2 ,iCol + 2 ,True)
								 	   End Select	
							Case 2, 5							
								.row =iRow : .Col = iCol  : .CellType = 1 :   .TypeVAlign = 2
								.row =iRow : .Col = iCol+1  : .CellType = 1                   
								ggoSpread.SSSetProtected	iCol  ,iRow,iRow
							Case Else
							
																					
					  End Select 		
				Next	
				
		   .rowheight(iRow) = 18	
		   					  
		Next    
		
		Call InitSpreadComboBox()	
		
		
		.Col = C_W2  : .Col2 = .maxcols -1
		.Row =1
		.TypeMaxEditLen = 60
		.TypeVAlign = 2
			
		.Col = C_W2  : .Col2 = .maxcols -1
		.Row =3
		.TypeMaxEditLen = 3	
		.TypeVAlign = 2
		
		.Col = C_W2  : .Col2 = .maxcols -1
		.Row =4
		.TypeMaxEditLen = 30
		.TypeVAlign = 2
		
		
		.Col = C_W2  : .Col2 = .maxcols -1
		.Row =5
		.TypeMaxEditLen = 50
		.TypeVAlign = 2
		
		.Col = C_W2  : .Col2 = .maxcols -1
		.Row =6
		.TypeMaxEditLen = 7
		.TypeVAlign = 2
		
			
		.Col = C_W2  : .Col2 = .maxcols -1
		.Row =8
		.TypeMaxEditLen = 70
		
		.Col = C_W2  : .Col2 = .maxcols -1
		'.Row =10
		.Row =11
		.rowheight(11) = 26
		.TypeMaxEditLen = 50
		.TypeVAlign = 2
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
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()

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
dim iIdx

		With frm1.vspdData
			select Case Row
			       'Case 6 , 8, 9
			       Case 7 , 9, 10
			     
						Select Case Col
									Case C_W2 ,C_W3, C_W4,C_W5, C_W6,C_W7, C_W8 ,C_W9 ,C_W10
										.Col = Col	: .Row = Row
										iIdx = UNICDbl(.Value)

										.Col = Col + 2
										.Value = iIdx
									
								End Select
			End Select
		
		End With	
	End Sub
'===============
Sub BtnIntCol()
Dim iCol
    With frm1.vspdData
        for iCol = C_W5  to .maxCols - 1 
            .col = iCol
            	Select Case iCol
						Case C_W2 ,C_W3, C_W4,C_W5, C_W6,C_W7, C_W8 ,C_W9 ,C_W10
							if  .ColHidden = True then
							     Call ggoSpread.SSSetColHidden(iCol ,iCol+2,False)
							   
							     Exit For
							End if 
				End Select			
             
        Next
    
		  Call InitData2()	

    
   	End With
End sub



'===========================================================================

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere ,Byval iRow )
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0

		Case  C_W2_POP, C_W3_POP, C_W4_POP, C_W5_POP, C_W6_POP, C_W7_POP, C_W8_POP, C_W9_POP ,C_W10_POP
	          if iRow = 6 then
					arrParam(0) = "업종"								' 팝업 명칭 
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
	          
               Elseif iRow =3 then
               
                    arrParam(0) = "국가코드"								' 팝업 명칭 
					arrParam(1) = "ufn_TB_COUNTRY(" & FilterVar(C_REVISION_YM, "''", "S") & ")" 								' TABLE 명칭 
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
	
	
		Case Else
			Exit Function
	End Select

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
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================



Function SetPopup(Byval arrRet,Byval iWhere , iRow)
	With frm1
		Select Case iRow
			Case 3
				.vspdData.Col = iWhere-1
				.vspdData.Row = iRow-1
				.vspdData.Text = arrRet(1)
			Case 6
				.vspdData.Col = iWhere-1
				.vspdData.Row = iRow-1
				.vspdData.Text = arrRet(2)
		End Select
			.vspdData.Col = iWhere-1
			.vspdData.Row = iRow
			.vspdData.Text = arrRet(0)

	End With
	
	Call vspdData_Change(iWhere-1,iRow)
	Call vspdData_Change(iWhere-1,iRow-1)
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function



Sub vspdData_Change(ByVal Col , ByVal Row )
 dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6   ,IntRetCD, sWhere
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

	Dim dblSum, dblCol(1)
	
	With frm1.vspdData
	
		Select Case ROW
			Case 6
			             .Row = Row
						 .Col = Col

							If frm1.txtFISC_YEAR.text >= "2006" Then	' -- 2006.07.07 수정 
								sWhere = " AND ATTRIBUTE_YEAR = '2005' " 
							Else
								sWhere = " AND ATTRIBUTE_YEAR = '2003' " 
							End If
						 
						 IntRetCD =  CommonQueryRs(" Top 1 STD_INCM_RT_CD  ","tb_std_income_rate"," STD_INCM_RT_CD = '" & Trim(frm1.vspdData.text) & "'" & sWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
						If IntRetCD = False Then
						    Call  DisplayMsgBox("970000","X","업종","X")                         '☜ : 입력된자료가 있습니다.
						    .col = Col
						    .row = row
						    frm1.vspdData.Text = ""
						   
						Else
           
						 
						End If
			Case 3			
			
			    
			       .Row = Row
			       .Col = Col
			    
			        IntRetCD =  CommonQueryRs(" COUNTRY_CD   ","ufn_TB_COUNTRY(" & FilterVar(C_REVISION_YM, "''", "S") & ")"," COUNTRY_CD = '" & Trim(frm1.vspdData.text) & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
						If IntRetCD = False Then
						    Call  DisplayMsgBox("970000","X","국가코드","X")                         '☜ : 입력된자료가 있습니다.
						    .col = col
						    .row = row
						    frm1.vspdData.Text = ""
						   
						Else
                             frm1.vspdData.Text = UCASE(Replace(lgF0,chr(11),""))
						 
						End If
				
		End Select
		


	
	End With
End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	Dim strTemp
	Dim intPos1

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        
        If Row = 6 or  Row = 3  Then
           Select Case Col 
                  Case C_W2_POP , C_W3_POP, C_W4_POP, C_W5_POP, C_W6_POP, C_W7_POP, C_W8_POP, C_W9_POP ,C_W10_POP
                  .col = col
                  .row = 6
                  Call OpenPopup(.Text,Col ,Row)
           End Select        

        End If
        
       
        
    End With
    
    

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
    Dim blnChange, dblSum, iCol
    
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
	
    ggoSpread.Source = frm1.vspdData
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

    frm1.vspdData.AddSelection C_W1, -1, C_W1, -1

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

    Call FncDeleteRow
       
    Call FncSave
    
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
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim iCol
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	
	If lgIntFlgMode <> parent.OPMD_UMODE  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE

		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 컨펌체크 : 그리드 락 
		If wgConfirmFlg <>"Y" Then

			'2 디비환경값 , 로드시환경값 비교 
			Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>

		Else
	
		End If
	Else
		Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
	End If


           
     With frm1.vspdData
     
        for iCol = C_W5  to .maxCols - 1 


            .col = iCol
            .Row = 1
            
			 Select Case .Col
			        Case   C_W5, C_W6  ,C_W7,  C_W8 , C_W9 , C_W10
							if  Trim(.text) <> "" then
							     Call ggoSpread.SSSetColHidden(iCol ,iCol+2,False)
							Else
							      Call ggoSpread.SSSetColHidden(iCol ,iCol+2,True)     
							  
							End if 
			End Select				
			     
        Next
        
      
      Call InitData2()

    
   	End With

    

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
    Dim strVal, strDel, lMaxRows, lMaxCols
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
		
		   For lRow = 1 To .MaxRows
    
               .Row = lRow
               .Col = 0
            
				Select Case  .Text
				    Case  ggoSpread.InsertFlag                                      '☜: Insert				  
				        

						 strVal = strVal & "C"  &  Parent.gColSep '0
						 'if (lRow =6 or lRow =8 or lRow =9 )  then
						 if (lRow =7 or  lRow =9 or lRow =10 )  then
							.Col = C_W1        : strVal = strVal & Trim(.Text) &  Parent.gColSep '1
							.Col = C_W2_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '2
							.Col = C_W3_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '3
							.Col = C_W4_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '4
							.Col = C_W5_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '5
							.Col = C_W6_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '6
							.Col = C_W7_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '7
							.Col = C_W8_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '8
							.Col = C_W9_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '9	
							.Col = C_W10_CD     : strVal = strVal & Trim(.Text) &  Parent.gRowSep '10		                          
									                         
							
					  else  
							.Col = C_W1     : strVal = strVal & Trim(.Text) &  Parent.gColSep '1
							.Col = C_W2     : strVal = strVal & Trim(.Text) &  Parent.gColSep '2
							.Col = C_W3     : strVal = strVal & Trim(.Text) &  Parent.gColSep '3
							.Col = C_W4     : strVal = strVal & Trim(.Text) &  Parent.gColSep '4
							.Col = C_W5     : strVal = strVal & Trim(.Text) &  Parent.gColSep '5
							.Col = C_W6     : strVal = strVal & Trim(.Text) &  Parent.gColSep '6
							.Col = C_W7     : strVal = strVal & Trim(.Text) &  Parent.gColSep '7
							.Col = C_W8     : strVal = strVal & Trim(.Text) &  Parent.gColSep '8
							.Col = C_W9     : strVal = strVal & Trim(.Text) &  Parent.gColSep '9
							.Col = C_W10     : strVal = strVal & Trim(.Text) &  Parent.gRowSep '10			                         
							
							  
					  end if		
						
				        
				        lGrpCnt = lGrpCnt + 1
				        
				    Case  ggoSpread.UpdateFlag                                      '☜: Update
				         strVal = strVal & "U"  &  Parent.gColSep
				         'if (lRow=6 or lRow =8 or lRow =9  )  then
				         if (lRow=7  or  lRow =9 or lRow =10  )  then
							.Col = C_W1        : strVal = strVal & Trim(.Text) &  Parent.gColSep '1
							.Col = C_W2_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '2
							.Col = C_W3_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '3
							.Col = C_W4_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '4
							.Col = C_W5_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '5
							.Col = C_W6_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '6
							.Col = C_W7_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '7
							.Col = C_W8_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '8
							.Col = C_W9_CD     : strVal = strVal & Trim(.Text) &  Parent.gColSep '9	
							.Col = C_W10_CD     : strVal = strVal & Trim(.Text) &  Parent.gRowSep '10		                          
									                         
					  else  
							.Col = C_W1     : strVal = strVal & Trim(.Text) &  Parent.gColSep '1
							.Col = C_W2     : strVal = strVal & Trim(.Text) &  Parent.gColSep '2
							.Col = C_W3     : strVal = strVal & Trim(.Text) &  Parent.gColSep '3
							.Col = C_W4     : strVal = strVal & Trim(.Text) &  Parent.gColSep '4
							.Col = C_W5     : strVal = strVal & Trim(.Text) &  Parent.gColSep '5
							.Col = C_W6     : strVal = strVal & Trim(.Text) &  Parent.gColSep '6
							.Col = C_W7     : strVal = strVal & Trim(.Text) &  Parent.gColSep '7
							.Col = C_W8     : strVal = strVal & Trim(.Text) &  Parent.gColSep '8
							.Col = C_W9     : strVal = strVal & Trim(.Text) &  Parent.gColSep '9
							.Col = C_W10     : strVal = strVal & Trim(.Text) &  Parent.gRowSep '10			                         
							
							  
					  end if		
				                                    
				        lGrpCnt = lGrpCnt + 1
				        
				         Case  ggoSpread.DeleteFlag                                      '☜: Update
				         strVal = strVal & "D"  &  Parent.gColSep
				        
							.Col = C_W1        : strVal = strVal & Trim(.Text) &  Parent.gRowSep '1
					                    
				        lGrpCnt = lGrpCnt + 1 
				        
				 
				End Select							  
        Next
        
        frm1.txtSpread.value        =  strDel & strVal

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
					<TD WIDTH=* align=right></TD>
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
								   <TD align=right>
									<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnIntCol()"   Flag=1>열추가</BUTTON></TD>
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
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

