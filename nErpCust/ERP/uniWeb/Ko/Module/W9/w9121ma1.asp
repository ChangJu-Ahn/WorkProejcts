
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 별지국외특수관계자요약손익계산서 
'*  3. Program ID           : w9121ma1
'*  4. Program Name         : w9121ma1.asp
'*  5. Program Desc         : 별지국외특수관계자요약손익계산서 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2005/03/18
'*  8. Modifier (First)     : 홍지영 
'*  9. Modifier (Last)      : 홍지영 
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

Const BIZ_MNU_ID = "w9121ma1"
Const BIZ_PGM_ID = "w9121mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID		= "W9121OA1"

Dim C_W1
Dim C_W1_NM
Dim C_W1_NM2
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10



Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
    C_W1				= 1
    C_W1_NM				= 2
    C_W1_NM2			= 3
    C_W2				= 4
    C_W3				= 5
    C_W4				= 6
    C_W5				= 7
    C_W6				= 8
    C_W7				= 9
    C_W8				= 10
    C_W9				= 11
    C_W10				= 12
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

	Call AppendNumberPlace("8","12","0")	' -- 금액 16자리 고정 : 출하검사패치

    
    .MaxCols = C_W10 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

     ggoSpread.SSSetEdit		C_W1,		"코드", 10,,,10,1
	 ggoSpread.SSSetEdit		C_W1_NM,	"", 20,,,50,1
	 ggoSpread.SSSetEdit		C_W1_NM2,	"", 8,,,50,1
	 ggoSpread.SSSetFloat	    C_W2,		"1"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W3,		"2"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W4,		"3"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W5,		"4"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W6,		"5"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W7,		"6"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W8,		"7"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W9,		"8"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W10,		"9"	, 20, "8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 
	 
	


	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
	Call ggoSpread.SSSetColHidden(C_W3,C_W10,True)
    ggoSpread.SSSetSplit2(3)
	' 그리드 헤더 합침 정의 


	
					

	
	.ReDraw = true

	Call SetSpreadLock 

    End With   
End Sub


'============================================  그리드 함수  ====================================



Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_W1, -1, C_W1_NM2

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
	Dim iMaxRows, iRow
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	iMaxRows = 19 ' 하드코딩되는 행수 
	With frm1.vspdData
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow , iMaxRows

		iRow = 0
		
      for irow = 1 to iMaxRows
		 .Row = iRow
		.Col = C_W1		: .value = iRow
	  NEXT	

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
		iRow = iRow + 1 : .Row = iRow    :  .TypeVAlign = 2 :  .rowheight(iRow) = 18
		
		.Col = C_W1_NM	: .value = "  (1)명칭"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
	  
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = "  (2)소재지"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		
		iRow = iRow + 1 : .Row = iRow   :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = "사업연도" :.TypeVAlign = 2 
		Call .AddCellSpan(C_W1_NM, iRow ,1, 2) 
		.Col = C_W1_NM2	: .value = "(시작)" :.TypeVAlign = 2 
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		
		.Col = C_W1_NM2	: .value = "(종료)" :.TypeVAlign = 2 
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = "(4)주된사업"
		Call .AddCellSpan(C_W1_NM, iRow ,1, 2) 
		.Col = C_W1_NM2	: .value = "(업종명)" :.TypeVAlign = 2 
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		
		.Col = C_W1_NM2	: .value = "(업종코드)" :.TypeVAlign = 2 
		
		
		
		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2  :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = "(5)자금 또는 출자금액"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		
		
		
		

		iRow = iRow + 1 : .Row = iRow    :.TypeVAlign = 2  :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = " (6)특수관계의 구분"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18 : .TypeEditMultiLine = true
		.Col = C_W1_NM	: .value = "(7)주식등의 "& vbCrLf & "소유비율(%)"
		 Call .AddCellSpan(C_W1_NM, iRow ,1, 4) 
		.Col = C_W1_NM2	: .value = "소유  (계)"  :.TypeVAlign = 2  :.TypeHAlign = 1 
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		
		.Col = C_W1_NM2	: .value = "      (직접)" :.TypeVAlign = 2  :.TypeHAlign = 1
		
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		
		.Col = C_W1_NM2	: .value = "피소유(계)" :.TypeVAlign = 2  :.TypeHAlign = 1 
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		
		.Col = C_W1_NM2	: .value = "     (직접)" :.TypeVAlign = 2  :.TypeHAlign = 1 
		
		
		

		iRow = iRow + 1 : .Row = iRow  
		.Col = C_W1_NM	: .value = " 계정과목"		:.TypeVAlign = 2 :.TypeHAlign = 2
		.Col = C_W1_NM2	: .value = "코드"			:.TypeVAlign = 2 :.TypeHAlign = 2
		
		 
		 iRow = iRow + 1 : .Row = iRow     
		.Col = C_W1_NM	: .value = " I.매출액"	    :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "01"			    :.TypeVAlign = 2 :.TypeHAlign = 2
	
		 
		 iRow = iRow + 1 : .Row = iRow     
		.Col = C_W1_NM	: .value = " II.매출원가"	:.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "02"			    :.TypeVAlign = 2 :.TypeHAlign = 2
		
		 iRow = iRow + 1 : .Row = iRow    
		.Col = C_W1_NM	: .value = " III.매출총손익" :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "03" :.TypeVAlign = 2 :.TypeHAlign = 2
		
		 iRow = iRow + 1 : .Row = iRow   
		.Col = C_W1_NM	: .value = " IV.판매비와관리비"     :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "04" :.TypeVAlign = 2	:.TypeHAlign = 2
		
		 iRow = iRow + 1 : .Row = iRow						:.TypeVAlign = 2
		.Col = C_W1_NM	: .value = "V.영업손익"			    :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "05" :.TypeVAlign = 2	:.TypeHAlign = 2
		
		 iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = "VI.법인세비용차감전순손익"  :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "06" :.TypeVAlign = 2 :.TypeHAlign = 2
		
	
		
		
	
		
		
	    
		
	  for iRow = 1 to  .maxrows 	
			for iCol = C_W2 to  .maxcols - 1
			         Select Case iRow
			               
			      			  Case 1 ,2,  5, 6, 8  
									Select Case iCol
									          Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
									           .row =iRow : .Col = iCol  :.CellType = 1 : .TypeMaxEditLen = 60
									           ggoSpread.SSSetProtected	iCol  ,iRow,iRow
								                
								             
										End Select	
		                      Case 3 , 4 
									Select Case iCol
									          Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
									           .row =iRow : .Col = iCol  :.CellType = 0 
									     
								     End Select	          
' 2005.02.27 수정: 국조1호에서 가져온것								
'							  Case 8
'									Select Case iCol
'									          Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
'									           .row =iRow : .Col = iCol  :.CellType = 1 
'									     
'								     End Select	          	     
							
							  Case 9, 10 ,11 ,12
									Select Case iCol
									          Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
													.Row = iRow : .Col = iCol
													.CellType = 13
													.TypeNumberMax = 100 : .TypeNumberMin = 0
													.TypeNumberDecPlaces= 2
									           '.row =iRow : .Col = iCol  :.CellType = 13 :.TypeCurrencyDecPlaces= 1 : .TypePercentMax = 10
									           
								     End Select	 
							
							  Case 13								
									Select Case iCol
									          Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
									           .row =iRow : .Col = iCol  :.CellType = 1 : .TypeMaxEditLen = 10 : .TypeEditMultiLine = false
									           
								     End Select	 
					  End Select 		
				Next	
				
		 
		   					  
		Next    
		

	
		
		
		
	End With
End Sub

'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD , iCol ,strSelect,strFrom,strWhere,ii,jj,arrVal1,arrVal2,iRow
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg , i

	' 온로드시 레퍼런스메시지 가져온다.
	
	 frm1.vspdData.AllowMultiBlocks = True  
	 frm1.vspdData.AddSelection C_W2, 1, C_W2, 1
	 frm1.vspdData.AddSelection C_W2, 2, C_W2, 2
	 frm1.vspdData.AddSelection C_W2, 5, C_W2, 6
	 frm1.vspdData.AddSelection C_W2, 8, C_W2, 8
	 frm1.vspdData.AllowMultiBlocks = False
	  
	  
	if   CommonQueryRs2by2("*", "TB_KJ1" , "Co_Cd  = '" & sCoCd & "'and Fisc_year = '" & sFiscYear & "' and Rep_type = '" & sRepType & "'" , lgF2By2)  =false then
	  	 Call DisplayMsgBox("X", "X", "국조 1호를 먼저 입력해 주세요", "X")
	    Exit Function  
	end if    
			    
			      
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
    
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	frm1.vspdData.SetSelection C_W2, 1, C_W2, 1
	
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
 
     CALL GetRefB
End Function

'============================== 레퍼런스 함수  ========================================
Function GetRefB()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD , iCol ,strSelect,strFrom,strWhere,ii,jj,arrVal1,arrVal2,iRow
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
	 
	  
	  
	if   CommonQueryRs2by2("*", "TB_KJ1" , "Co_Cd  = '" & sCoCd & "'and Fisc_year = '" & sFiscYear & "' and Rep_type = '" & sRepType & "'" , lgF2By2)  =false then
	  	 Call DisplayMsgBox("X", "X", "국조 1호를 먼저 입력해 주세요", "X")
	    Exit Function  
	end if    
			    
			      
   
    strSelect = "w1,w2,w3,w4,w5,w6,w7,w8,w9,w10"
    strFrom = "dbo.ufn_TB_KJ_BJ1_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')"
    strWhere = ""
		
	with frm1.vspdData	
	    .Redraw = False

			If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			   
			Else 

						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			
					For iRow =1 to .MaxRows
					
								For ii = 0 to Ubound(arrVal1,1) - 1
										arrVal2 = Split(arrVal1(ii), chr(11))
							
									.Row = iRow :.Col = C_W1 
								  	

								  	if   .Value  =  Trim(arrVal2(1)) then  
								  	    .Col = C_W2 : .Value = Trim(arrVal2(2))
								  	    .Col = C_W3 : .Value = Trim(arrVal2(3))
								  	    .Col = C_W4 : .Value = Trim(arrVal2(4))
								  	    .Col = C_W5 : .Value = Trim(arrVal2(5))
								  	    .Col = C_W6 : .Value = Trim(arrVal2(6))
								  	    .Col = C_W7 : .Value = Trim(arrVal2(7))
								  	    .Col = C_W8 : .Value = Trim(arrVal2(8))
								  	    .Col = C_W9 : .Value = Trim(arrVal2(9))
								  	    .Col = C_W10 : .Value = Trim(arrVal2(10))
								  	  
								  	 End if   

							Next	
					Next		
						
			End If
					
				
			
			
			  for iCol = C_W2  to .maxCols - 1 
		            .col = iCol
		            .Row = 1
				
					   
						if  Trim(.text) <> "" then
						     Call ggoSpread.SSSetColHidden(iCol ,iCol+2,False)
					    
						Else
						      Call ggoSpread.SSSetColHidden(iCol ,iCol+2,True)     
						End if 
						
					     
		        Next
		.Redraw = True
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
    
    Call SetToolbar("1110100000101111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()
    Call getRefB()
    
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

End Sub



'===============

'===========================================================================






Sub vspdData_Change(ByVal Col , ByVal Row )
 dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6   ,IntRetCD , dblSum1 , dblSum2,iCol
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
	
		Select Case ROW
			   
			   Case "8"
			   
			   
			    for iCol = C_W2  to .maxCols - 1 


				.col = iCol

				if  Trim(.VALUE) <> "" then
				
					call CommonQueryRs("MINOR_CD"," B_MINOR "," MAJOR_CD= 'W1051' AND MINOR_NM = '" & .VALUE & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

					if Trim(replace(lgF0,chr(11),"")) = "" then
							Call DisplayMsgBox("122500", "X", """" & .VALUE & """", "X")
							.value = ""
					End if 
			   End If
			   Next
			   
			   Case "9"
			   
			   
			    for iCol = C_W2  to .maxCols - 1 


				.col = iCol
			
				if  Trim(.text) <> "" then
					If .value < 0 Then 
						Call DisplayMsgBox("189306", "X", "X", "X")
						.value = ""
					End If
				End if 
				
			     
				Next
			   
			   
			 
			   Case "10"
			   
				for iCol = C_W2  to .maxCols - 1 


				.col = iCol
			
				if  Trim(.text) <> "" then
					If .value < 0 Then 
						Call DisplayMsgBox("189306", "X", "X", "X")
						.value = ""
					End If
				End if 
				
			     
				Next
			   
			   Case "12"
			   
			   for iCol = C_W2  to .maxCols - 1 


				.col = iCol
			
				if  Trim(.text) <> "" then
					If .value < 0 Then 
						Call DisplayMsgBox("189306", "X", "X", "X")
						.value = ""
					End If
				End if 
				
			     
				Next
			   
			   Case "14"
			   
			   for iCol = C_W2  to .maxCols - 1 


				.col = iCol
			
				if  Trim(.text) <> "" then
					If .value < 0 Then 
						Call DisplayMsgBox("189306", "X", "X", "X")
						.value = ""
					End If
				End if 
				
			     
				Next
			   
			   Case "15"
			   
			   for iCol = C_W2  to .maxCols - 1 


				.col = iCol
			
				if  Trim(.text) <> "" then
					If .value < 0 Then 
						Call DisplayMsgBox("189306", "X", "X", "X")
						.value = ""
					End If
				End if 
				
			     
				Next
			   
			   Case "17"
			   
			   for iCol = C_W2  to .maxCols - 1 


				.col = iCol
			
				if  Trim(.text) <> "" then
					If .value < 0 Then 
						Call DisplayMsgBox("189306", "X", "X", "X")
						.value = ""
					End If
				End if 
				
			     
				Next
				
		
		End Select
		


	
	End With
End Sub



Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	Dim strTemp
	Dim intPos1

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        
    
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

    Call SetToolbar("1110100000001111")

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
       frm1.vspdData.AddSelection C_W1, -1, C_W1, -1

       Call FncDeleteRow
    
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
			Call SetToolbar("1111100000011111")										<%'버튼 툴바 제어 %>

		Else
	
		End If
	Else
		Call SetToolbar("1111100000011111")										<%'버튼 툴바 제어 %>
	End If

          
     With frm1.vspdData
     
        for iCol = C_W2  to .maxCols - 1 


            .col = iCol
            .Row = 1
            
			
			if  Trim(.text) <> "" then
			     Call ggoSpread.SSSetColHidden(iCol ,iCol,False)
			    
			Else
			      Call ggoSpread.SSSetColHidden(iCol ,iCol,True)     
			End if 
				
			     
        Next
        
     End With   
 
    Call InitData2()
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
    Dim strVal, strDel, lMaxRows, lMaxCols , IRow
 
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
               if  IRow <> 1 or lRow <> 5 or lRow <> 6 or lRow <> 8 then
               .Row = lRow
               .Col = 0
               
						Select Case  .Text
						    Case  ggoSpread.InsertFlag                                      '☜: Insert
		

								 strVal = strVal & "C"  &  Parent.gColSep '0
								
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
									
				
						            lGrpCnt = lGrpCnt + 1
						        
						    Case  ggoSpread.UpdateFlag                                      '☜: Update
						           strVal = strVal & "U"  &  Parent.gColSep
						       
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
									
							                 
						        lGrpCnt = lGrpCnt + 1
						        
						         Case  ggoSpread.DeleteFlag                                      '☜: Update
						         strVal = strVal & "D"  &  Parent.gColSep
						        
									.Col = C_W1        : strVal = strVal & Trim(.Text) &  Parent.gRowSep '1
							                    
						        lGrpCnt = lGrpCnt + 1 
						        
						 
						End Select
				END IF		
				  
							  
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
	
    Call MainQuery()
End Function

Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i

	Dim intCnt,IntRetCD


    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If
   Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE)
	
	 StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
	 StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
	 StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE
	 
	call CommonQueryRs("W5,W8"," TB_KJ1 "," CO_CD= '" & varCo_Cd & "' AND FISC_YEAR='" & varFISC_YEAR & "' AND REP_TYPE='" & varREP_TYPE & "' and w4 >'' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    if Trim(replace(lgF0,chr(11),"")) <> "" then
		if Trim(replace(lgF1,chr(11),"")) <> ""	then
			StrUrl = StrUrl & "|flag|"       & "1','2','3"
		Else
			StrUrl = StrUrl & "|flag|"       & "1','2"
		End if
	Else
		StrUrl = StrUrl & "|flag|"       & "1"		
    end if

    			

     EBR_RPT_ID	    = "W9121OA1"
     ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
     if  strPrintType = "VIEW" then
	     Call FncEBRPreview(ObjName, StrUrl)
     else
	     Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	
   

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
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">금액 불러오기</A>  
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
									<TD CLASS="TD6"><script language =javascript src='./js/w9121ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
								<TD  align=center CLASS="TD61">
									[국   외   특   수   관   계  자]
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/w9121ma1_vaSpread1_vspdData.js'></script>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

