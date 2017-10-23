
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 제8호부표5공제감면세액계산서(5)
'*  3. Program ID           : W6117MA1
'*  4. Program Name         : W6117MA1.asp
'*  5. Program Desc         : 제8호부표5공제감면세액계산서(5)
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID = "W6117MA1"	
Const BIZ_PGM_ID = "W6117Mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W6117OA1"



Dim C_SEQ_NO
Dim C_W9
Dim C_W10
Dim C_W11
dim C_W12
dim C_W13
dim C_W14
dim C_W15
dim strMode



Const C_SHEETMAXROWS = 7








Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	 C_SEQ_NO	= 1

	 C_W9		= 2
	 C_W10		= 3
	 C_W11		= 4
	 C_W12		= 5
	 C_W13		= 6
	 C_W14		= 7
	 C_W15		= 8

    
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
	Dim IntRetCD
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub SetDefaultVal()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

     

    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)



End Sub
Sub InitSpreadSheet()

    Call initSpreadPosVariables()  

			      	
				With frm1.vspdData
	
					ggoSpread.Source = frm1.vspdData	
					'patch version
					 ggoSpread.Spreadinit "V20041222",,parent.gAllowDragDropSpread    
					 
						.ReDraw = false

					    .MaxCols = C_W15 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
						.Col = .MaxCols														'☆: 사용자 별 Hidden Column
						.ColHidden = True    
						       
					 .MaxRows = 0
					 ggoSpread.ClearSpreadData

					 Call AppendNumberPlace("6","3","2")
					 .RowHeight(0) = 30 
					 
					 Call GetSpreadColumnPos("A")   
					 

					 ggoSpread.SSSetEdit     C_SEQ_NO, "순번", 10,,,100,1
					 ggoSpread.SSSetMask     C_W9,"사업연도", 10, 2 ,"9999"
					 ggoSpread.SSSetFloat    C_W10, "(10)외국납부"& vbCr & " 세액발생액" ,    12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W11, "(11)기공제액"					   ,    12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W12, "(12)미공제액"& vbCr & " ((10)-(11))",   12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W13, "(13)당기공제액"& vbCr & " ((13)≤(8))", 12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W14, "(14)공제누계"& vbCr & " ((11)+(13))",    12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W15, "(15)이월액"& vbCr& " ((10)-(14))",    12,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 

					
					
				
				

						Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
						Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
									
					
						.ReDraw = true

				 
						Call SetSpreadLock 		

					 
					 End With



			
		
    
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt
	Dim strW1    
			
	if 	pOpt = "S" then
		    lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
			lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
			lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '
			lgKeyStream = lgKeyStream & Trim(Frm1.txtw1.Value ) &  parent.gColSep		' 3   (1)국외원천소득 총액 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtw2.Value ) &  parent.gColSep		' 4   (2)감면을 적용받은 국외원천소득                     
			lgKeyStream = lgKeyStream & Trim(Frm1.txtw3.Value ) &  parent.gColSep		' 5   (3)감면비율 표시 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW3value.Value ) &  parent.gColSep	' 6   (3)감면비율 값 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW4.Value ) &  parent.gColSep		' 7   (4)차감되는 감면국외 원천소득 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW5.Value ) &  parent.gColSep		' 8   (5)외국납부세액 공제대상 국외원천소득 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW7_A.Value ) &  parent.gColSep		' 9   (7)계산내역 (a)
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW7_B.Value ) &  parent.gColSep		' 10  (7)계산내역 (b)
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW7_C.Value ) &  parent.gColSep		' 11  (7)계산내역 (c)
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW8.Value ) &  parent.gColSep		' 11  (8)공제한도        
			
	Else
	        lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
			lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
			lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '
	End if		
    
	
	 

End Sub 


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx



End Sub


Sub SetSpreadLock()


           With frm1
    


				.vspdData.ReDraw = False
                  'ggoSpread.SSSetRequired C_W9,	  -1, C_W9
                  
                  ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
                  If .vspdData.MaxRows > 0 Then
	                  ggoSpread.SSSetRequired  C_W9  , 1, (.vspdData.MaxRows - 2)
		              ggoSpread.SSSetProtected C_W9 , (.vspdData.MaxRows - 1), (.vspdData.MaxRows)
		          End If
				  ggoSpread.SpreadLock C_w12, -1, C_w12
				  ggoSpread.SpreadLock C_w14, -1, C_w14
			      ggoSpread.SpreadLock C_w15, -1, C_w15
			  	

				.vspdData.ReDraw = True

			End With
 
   
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
dim sumRow
    ggoSpread.Source = frm1.vspdData
    With frm1
	
	sumRow = CInt(.vspdData.MaxRows)
	
    .vspdData.ReDraw = False
    
	ggoSpread.SSSetProtected C_SEQ_NO , pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_W9  , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_w12 , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_w14 , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_w15 , pvStartRow, pvEndRow
		
    .vspdData.ReDraw = True
    
    End With
End Sub






Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
        
				C_SEQ_NO	= iCurColumnPos(1)
				C_W9		= iCurColumnPos(2)
				C_W10	= iCurColumnPos(3)	
				C_W11	= iCurColumnPos(4)
				C_W12	= iCurColumnPos(5)
				C_W13	= iCurColumnPos(6)
				C_W14	= iCurColumnPos(7)
				C_W15	= iCurColumnPos(8)
				
	

    End Select    
End Sub

'============================================  조회조건 함수  ====================================
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0

		Case 5
		
	
		Case Else
			Exit Function
	End Select

	IsOpenPop = True
			
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================



Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
			Case 5
			
		End Select
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function

Function CheckReCalc()
	Dim dblSum_Data

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
	    
        if  frm1.vspdData.maxrows =< 0 then exit function
        ggoSpread.UpdateRow .MaxRows
         Call CheckReCalc2()

		 dblSum_Data = FncSumSheet(frm1.vspdData, C_W10, 1, .MaxRows - 1, true, .MaxRows, C_W10, "V")	' 합계 
		 dblSum_Data = FncSumSheet(frm1.vspdData, C_W11, 1, .MaxRows - 1, true, .MaxRows, C_W11, "V")	' 합계 
		 dblSum_Data = FncSumSheet(frm1.vspdData, C_W12, 1, .MaxRows - 1, true, .MaxRows, C_W12, "V")	' 합계 
		 dblSum_Data = FncSumSheet(frm1.vspdData, C_W13, 1, .MaxRows - 1, true, .MaxRows, C_W13, "V")	' 합계 
		 dblSum_Data = FncSumSheet(frm1.vspdData, C_W14, 1, .MaxRows - 1, true, .MaxRows, C_W14, "V")	' 합계 
		 dblSum_Data = FncSumSheet(frm1.vspdData, C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' 합계 

	End With
End Function


Function CheckReCalc2()
	Dim dblSum
    Dim i ,dblW12, dblW12Sum,dblW10, dblW11,dblW13,dblW14,dblW12SumChk
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
	
	    dblW12Sum = 0
	          dblW12SumChk  = FncSumSheet(frm1.vspdData, C_W12, 1, .MaxRows - 1, true, .MaxRows, C_W12, "V")	' 합계 
	 
	         if  unicdbl(frm1.txtw8.value  ) < unicdbl(dblW12SumChk) then

	              dblSum = unicdbl(frm1.txtw8.value  )
	         else 
	              dblSum = unicdbl(dblW12SumChk)
	         end if     
     if frm1.vspdData.maxrows <> 2 and   frm1.vspdData.maxrows > 0 then
				for i = 1 to  frm1.vspdData.maxrows - 2
				    .Row = i
				    .col = C_W12
				     dblW12 = unicdbl(.text )
				    
				  
				     if  dblW12Sum +  unicdbl(dblW12 ) =< dblSum  then
                  
							.col = C_W13 : .Row = i : .text  = unicdbl(dblW12 ) 
							
				      else 
				         
							.col = C_W13
							
							.text  = unicdbl(dblSum) - unicdbl(dblW12Sum )
						
								if  .text < 0 then'
								     .text  =0
								end if     
				     end if
				             dblW12Sum = dblW12Sum +  unicdbl(dblW12 )
				            .Row = frm1.vspdData.maxrows -1
							.col = C_W13
							
							.text  = unicdbl(dblSum) - unicdbl(dblW12Sum )
								if  .text < 0 then'
									 .text  =0
								end if    
								
								.Col = C_W10 :  Frm1.vspdData.Row=i :  dblW10 = .text 
								.Col = C_W11 :  Frm1.vspdData.Row=i  : dblW11 = .text 
								.Col = C_W12 :  Frm1.vspdData.Row=i  : .text  = unicdbl(dblW10) - unicdbl(dblW11)			
				    
								.Col = C_W13 :  Frm1.vspdData.Row=i :  dblW13 = .text 
								.Col = C_W11 :  Frm1.vspdData.Row=i  : dblW11 = .text 
								.Col = C_W14 :  Frm1.vspdData.Row=i  : .text  = unicdbl(dblW11)+ unicdbl(dblW13)
					
								.Col = C_W10 :  Frm1.vspdData.Row=i :dblW10 = .text 
								.Col = C_W14 :  Frm1.vspdData.Row=i :dblW14 = .text 
								.Col = C_W15 :  Frm1.vspdData.Row=i  : .text  = unicdbl(dblW10) - unicdbl(dblW14)
					
								  .Col = C_W10 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1 :  dblW10 = .text 
								.Col = C_W11 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1  : dblW11 = .text 
								.Col = C_W12 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1   : .text  = unicdbl(dblW10) - unicdbl(dblW11)			
				    
								.Col = C_W13 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1  :  dblW13 = .text 
								.Col = C_W11 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1   : dblW11 = .text 
								.Col = C_W14 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1   : .text  = unicdbl(dblW11)+ unicdbl(dblW13)
					
								.Col = C_W10 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1  :dblW10 = .text 
								.Col = C_W14 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1  :dblW14 = .text 
								.Col = C_W15 :  Frm1.vspdData.Row=Frm1.vspdData.maxrows - 1   : .text  = unicdbl(dblW10) - unicdbl(dblW14)
					
				Next
				
				  
		else   
		       .Row = frm1.vspdData.maxrows -1
		       .col = C_W13
			   .text  = unicdbl(dblSum) - unicdbl(dblW12Sum )
			   if  .text < 0 then'
			        .text  =0
			   end if  
			         .Col = C_W11 : Frm1.vspdData.Row= .Row   : dblW11 = .text 
					.Col = C_W14 : Frm1.vspdData.Row= .Row  : .text  = unicdbl(dblW11)+ unicdbl(dblW13)
					
					.Col = C_W10 :  Frm1.vspdData.Row=.Row  :dblW10 = .text 
					.Col = C_W15 : Frm1.vspdData.Row= .Row  : .text  = unicdbl(dblW10)- unicdbl(dblW14)  
		end if		

	End With
End Function



'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
	Call  AppendNumberPlace("7", "2", "0")
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call SetDefaultVal

    Call FncQuery
End Sub
'============================================  사용자 함수  ====================================
Function  Fn_SumCal 
	
	Frm1.txtW3value.value  = unicdbl(Frm1.txtw3.value)/100
	
    frm1.txtw4.value = unicdbl(frm1.txtw2.value) * unicdbl(frm1.txtw3value.value)
    frm1.txtw5.value = unicdbl(frm1.txtw1.value) - unicdbl(frm1.txtw4.value)
    frm1.txtw7_B.value = unicdbl(frm1.txtw5.value)
    if frm1.txtw7_C.value  = 0 then
       frm1.txtw8.value = 0 
    else
       frm1.txtw8.value = unicdbl(frm1.txtw7_A.value)  * (unicdbl(frm1.txtw7_B.value) /unicdbl(frm1.txtw7_C.value))
       
    end if   
    if frm1.txtw8.value < 0 then
       frm1.txtw8.value = 0
    end if
   Call  CheckReCalc()

End Function


Function MaxSpreadVal_S(Byref objSpread, ByVal intCol, byval Row)

	Dim iRows
	Dim MaxValue
	Dim tmpVal

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows -2
		objSpread.row = iRows
	    objSpread.col = intCol

		If objSpread.Text = "" Then
		   tmpVal = 0
		Else
  		   tmpVal = cdbl(objSpread.value)
		End If

		If tmpval > MaxValue And tmpval < SUM_SEQ_NO Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1

	objSpread.row	= row
	objSpread.col	= intCol
	objSpread.text	= MaxValue
	MaxSpreadVal_S = MaxValue
end Function


'============================================  이벤트 함수  ====================================

Function GetRef()	
    Dim IntRetCD , i
    Dim sMesg
    Dim sFiscYear, sRepType, sCoCd
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6,BackColor
		sCoCd		= frm1.txtCO_CD.value
		sFiscYear	= frm1.txtFISC_YEAR.text
		sRepType	= frm1.cboREP_TYPE.value
	
		if wgConfirmFlg = "Y" then
		  
		    Exit function
		end if

		 sMesg = wgRefDoc & vbCrLf & vbCrLf
		 BackColor = frm1.txtW7_C.BackColor
         frm1.txtW7_C.BackColor =&H009BF0A2&
         frm1.txtW7_A.BackColor =&H009BF0A2&
         IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
		 frm1.txtW7_C.BackColor = BackColor
		 frm1.txtW7_A.BackColor = BackColor
		If IntRetCD = vbNo Then
			Exit Function
			
		End If
		Dim arrW1 ,arrW2 
		'3호 서식의 (120)산출세액 
		'3호 서식의 (113)산출세액 
		IntRetCD = CommonQueryRs("W7A,W7C","dbo.ufn_TB_8_5_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

 	
		if  IntRetCD = False  then
    	     IntRetCD = DisplayMsgBox("W60006", "x", "(120) 산출세액"  , "X")     
		else 
		   
		    arrW1 = REPLACE(lgF0, chr(11),"")
		    arrW2 = REPLACE(lgF1, chr(11),"")
	
		    frm1.txtW7_A.value =  unicdbl(arrW1)
		       Call txtW7_A_change
		    frm1.txtW7_C.value =  unicdbl(arrW2)
		       Call txtW7_C_change
       end if    
    	
End Function




Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub





Sub txtw1_Change( )
    lgBlnFlgChgValue = True
    Call Fn_SumCal
End Sub


Sub txtw2_Change( )
    lgBlnFlgChgValue = True
    Call Fn_SumCal
End Sub

Sub txtw3_Change( )
    lgBlnFlgChgValue = True
    Frm1.txtW3value.value  = unicdbl(Frm1.txtw3.value)/100
    Call Fn_SumCal   
    
End Sub

Sub txtW7_A_change( )
    lgBlnFlgChgValue = True
    Call Fn_SumCal
End Sub

Sub txtW7_C_change( )
    lgBlnFlgChgValue = True
    Call Fn_SumCal
End Sub

Sub SetSpreadTotalLine()
	Dim iRow

		ggoSpread.Source =  frm1.vspdData
		With  frm1.vspdData
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W9: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
				ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
				
				.Row = .MaxRows -1
				.Col = C_W9: .CellType = 1	: .Text = "당기"	: .TypeHAlign = 2
				ggoSpread.SSSetProtected C_W9, .MaxRows-1, .MaxRows-1
				 'Call SetSpreadColor(-1,-1)  
			End If
		End With

End Sub 


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	
End Sub

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		
	End With
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
     Dim iDx
    Dim IntRetCD
    Dim i , j,txtW9 ,txtW9_Row1 , txtW9_Row2, dblW10, dblW11,dblW13,dblW14
   
 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
  '------ Developer Coding part (Start ) -------------------------------------------------------------- 

  '--------------------'그리드에 입력된 내역이 기존데이터에 있을때 체크----------------------------------
    Select Case Col
        Case C_W9
             Frm1.vspdData.Col = C_W9 : Frm1.vspdData.Row = Row : txtW9 = Frm1.vspdData.text 
          
             if unicdbl(txtW9) < unicdbl(frm1.txtFISC_YEAR.text) - 5 or  unicdbl(frm1.txtFISC_YEAR.text)  <= unicdbl(txtW9)   then
                  IntRetCD = DisplayMsgBox("X", "X", "년도는 당기년도의 5년 내이어야 합니다.", "X")
                  Frm1.vspdData.Col = C_W9 : Frm1.vspdData.Row = Row : Frm1.vspdData.value = "" 
             end if
  
     
              
          
        
            
        Case C_W10 , C_W11 ,C_W14 ,C_W13
            
            Frm1.vspdData.Col = C_W10 : Frm1.vspdData.Row = Row : dblW10 = Frm1.vspdData.text 
            Frm1.vspdData.Col = C_W11 : Frm1.vspdData.Row = Row : dblW11 = Frm1.vspdData.text 
            Frm1.vspdData.Col = C_W13 : Frm1.vspdData.Row = Row : dblW13 = Frm1.vspdData.text 
            Frm1.vspdData.Col = C_W14 : Frm1.vspdData.Row = Row : dblW14 = Frm1.vspdData.text 
            Frm1.vspdData.Col = C_W12 : Frm1.vspdData.Row = Row : Frm1.vspdData.text  = unicdbl(dblW10) - unicdbl(dblW11)
            Frm1.vspdData.Col = C_W15 : Frm1.vspdData.Row = Row : Frm1.vspdData.text  = unicdbl(dblW10) - unicdbl(dblW14)
            Frm1.vspdData.Col = C_W14 : Frm1.vspdData.Row = Row : Frm1.vspdData.text  = unicdbl(dblW11) + unicdbl(dblW13)
           

        
            Call CheckReCalc()
      
     
           
                          
        
    End Select
    
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If uniCDbl(Frm1.vspdData.text) < uniCDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub




Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101000000") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
   ' If Row <= 0 Then
   '    ggoSpread.Source = frm1.vspdData
       
   '    If lgSortKey = 1 Then
   '        ggoSpread.SSSort Col               'Sort in ascending
   '        lgSortKey = 2
   '    Else
   '        ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
   '        lgSortKey = 1
   '    End If
       
   '    Exit Sub
   ' End If

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


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
      
    	If lgStrPrevKey <> "" And lgStrPrevKey2 <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End If
End Sub







'============================================  툴바지원 함수  ====================================
'=====================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

     Call ggoOper.ClearField(Document, "2")
    Call SetDefaultVal
    Call InitVariables               

    Call SetToolbar("1100110100000111")          '⊙: 버튼 툴바 제어 
    FncNew = True                

End Function

'=====================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
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
	
	Call MakeKeyStream("X")
     
    CALL DBQuery()
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '☜: Processing is NG
    
    
    <%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    
    
    If lgIntFlgMode <>  parent.OPMD_UMODE  Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call MakeKeyStream("X")

    
    If DbDelete= False Then
       Exit Function
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function





Function FncSave() 
   dim IntRetCD
    FncSave = False                                                         
    
    
    
    

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.maxrows = 0 then 
       IntRetCD =  DisplayMsgBox("WC0002","x","x","x")                           '☜:There is no changed data. 
	    Exit Function
    end if 
    
    If ggoSpread.SSCheckChange = False  or lgBlnFlgChgValue = False Then
                 ggoSpread.Source = frm1.vspdData
			If  ggoSpread.SSCheckChange = False Then
			    IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
			    Exit Function
			End If
       
    End If
    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  
    
    Call MakeKeyStream("S")
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
		If .vspdData.ActiveRow > 0 and .vspdData.ActiveRow <> .vspdData.maxrows Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			
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
    if frm1.vspdData.maxrows -1 > frm1.vspdData.activerow then      '당기부분보다 큰 라인만 취소     
       ggoSpread.EditUndo                                                  '☜: Protect system from crashing
     
    end if 

   if frm1.vspdData.maxrows = 2 then           '                    '하단 고정계  합계/당기부분 
       ggoSpread.EditUndo                                                 
       ggoSpread.EditUndo 
    end if 
    Call CheckReCalc()                                               '재계산 
    
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID,ii

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncInsertRow = False                                                         '☜: Processing is NG
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
       ' imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If

 

	With frm1	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		 IF C_SHEETMAXROWS  <= .vspdData.MaxRows THEN
		    EXIT Function
		 END IF
		 
		'.vspdData.ReDraw = False
		iSeqNo = .vspdData.MaxRows+1
	
       
        
		if 	.vspdData.MaxRows = 0 then
		
			
			.vspdData.row = .vspdData.MaxRows
		     ggoSpread.InsertRow  imRow 
		     
		    .vspdData.Col	= C_SEQ_NO
			.vspdData.Text	= SUM_SEQ_NO - 1
			.vspdData.Col	= C_W9
			.vspdData.Text	= "합계"

			 
			  ggoSpread.InsertRow  imRow 
			 
		     SetSpreadColor 1, 1
		     
		      .vspdData.row = .vspdData.MaxRows
		     .vspdData.Col	= C_SEQ_NO
			.vspdData.Text	= SUM_SEQ_NO
			
			 Call SetSpreadTotalLine
			  'ggoSpread.SSSetProtected C_W9 , .vspdData.MaxRows-1, .vspdData.MaxRows-1
			
		else
				'.vspdData.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
				
		     
				iRow = .vspdData.ActiveRow

				If iRow >= .vspdData.MaxRows-1 Then
				
				    .vspdData.ActiveRow  = .vspdData.MaxRows -2
				     
					ggoSpread.InsertRow .vspdData.MaxRows -2 , imRow 
					SetSpreadColor .vspdData.MaxRows -2, .vspdData.MaxRows -2
    
			
					For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
					
						Call MaxSpreadVal_S(frm1.vspdData, C_SEQ_NO, ii)
						
					Next
					Call SetSpreadColor(iRow , (iRow-1) + imRow)   
				Else
				
				  
			
		            ggoSpread.InsertRow ,imRow
		            For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
					
						Call MaxSpreadVal_S(frm1.vspdData, C_SEQ_NO, ii)
						
					Next
					Call SetSpreadColor(iRow + 1, (iRow+1) +  imRow - 1)   

					
				End If
        end if 	
    End With
    
   
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
   If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
    if frm1.vspdData.maxrows -1 > frm1.vspdData.activerow then
       lDelRows = ggoSpread.DeleteRow                                              '☜: Protect system from crashing
      
		Call CheckReCalc()
    end if 
 
    lgBlnFlgChgValue = True
    
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
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
       
        
        
        
        	strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
            strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
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
	lgBlnFlgChgValue = false
    '-----------------------
    'Reset variables area
    '-----------------------
    
    Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    
    Call InitData
	'1 컨펌체크 
	If wgConfirmFlg = "Y" Then

	    Call SetToolbar("1100000000000111")	
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		  Call SetToolbar("1101111100000111")									<%'버튼 툴바 제어 %>
	
	End If
	
    lgIntFlgMode = parent.OPMD_UMODE
    Call SetSpreadLock
    Call SetSpreadTotalLine

	frm1.vspdData.focus			
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
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key	
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
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
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    ggoSpread.Source = frm1.vspdData 
	With Frm1
	

	     .txtFlgMode.value = lgIntFlgMode	
		 strMode	   = .txtFlgMode.value
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        

        
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W9          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W10          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W11          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W12          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W13          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W14          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W15          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
    

                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep
                                                         'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W9          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W10          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W11          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W12          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W13          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W14          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W15          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
    

                    
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	Frm1.txtSpread.value      = strDel & strVal

    Frm1.txtMaxRows.value  =     lGrpCnt - 1

	Frm1.txtMode.value        =  Parent.UID_M0002
	frm1.txtFlgMode.value	  =  lgIntFlgMode
	frm1.txtKeyStream.value      =  lgKeyStream
		
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
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%> </TD>
				</TR>
				
				
				
				
					<TR>
					<TD WIDTH=800 valign=top HEIGHT="100" >
					   
					      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>1.공제한도액 계산 </LEGEND>
									<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   
										
										<TR>
											<TD CLASS="TD51" align =center width = 20% >
												(1)국외원천소득총액 
											</TD>
											
										    <TD CLASS="TD51" align =center width = 20%  >
												(2)감면을<br> 적용받은<br> 국외원천소득 
											</TD>
											<TD CLASS="TD51" align =center width = 20%  >
												(3)<br>감면비율 
											</TD>
											<TD CLASS="TD51" align =center width = 20% >
												(4)차감되는 감면국외<br> 원천소득((2) x (3))
											</TD>
											<TD CLASS="TD51" align =center width = 20%  >
												(5)외국납부세액공제대상<br> 국외원천소득((1) - (4))
											</TD>
											
										</TR>
										
										<TR>
											
											<TD CLASS="TD61" align =center>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2" name=txtW2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center  >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X7Z" width = 90%></OBJECT>');</SCRIPT>%
												
											</TD>
											<TD CLASS="TD61" align =left >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =left >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
											<TR>
											<TD CLASS="TD51" align =center  >
												(6)계산기준 
											</TD>
											
										    <TD CLASS="TD51" align =center ColSPAN=3 >
												(7)계산내역 
											</TD>
											<TD CLASS="TD51" align =center  >
												(8)공제한도 
											</TD>
											
										</TR>
										<TR>
											       <TD CLASS="TD51" COLSPAN=2>
											       <TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
														<TR>
															<TD ALIGN=CENTER WIDTH=35%>산출세액 x</TD>
															<TD ALIGN=CETER WIDTH=45%>
															<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																<TR>
																	<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT></TD>
																	<TD ALIGN=CENTER>국외원천소득</TD>
																	<TD ROWSPAN=3 ></TD>
																</TR>
																<TR>
																	<TD HEIGHT=1 BGCOLOR=BLACK></TD>
																</TR>
																<TR>
																	<TD ALIGN=CENTER>과세표준</TD>
																</TR>
															</TABLE>	
															</TD>
															<TD WIDTH=20%>&nbsp;</TD>
														</TR>
													</TABLE></TD>															
												  
											
											     <TD CLASS="TD61" COLSPAN=2>
											       <TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
														<TR>
															<TD ALIGN=CENTER WIDTH=35%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7_A" name=txtW7_A CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
															<TD ALIGN=CETER WIDTH=45%>
																<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																	<TR>
																		<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT>x&nbsp;&nbsp;&nbsp;</TD>
																		<TD ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7_B" name=txtW7_B CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
																	
																	</TR>
																	<TR>
																		<TD HEIGHT=1 BGCOLOR=BLACK></TD>
																	</TR>
																	<TR>
																		<TD ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7_C" name=txtW7_C CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
																	</TR>
																</TABLE>	
															</TD>
															<TD WIDTH=20%>&nbsp;</TD>
														</TR>
													</TABLE></TD>															
												
												   <TD CLASS="TD61" COLSPAN =2> <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
											
											
											
											
											</TR>
										
										
										
										
						
									</TABLE>
						   </FIELDSET>				
						   			
					</TD>
				</TR>
				<TR>
				    
						<TD WIDTH=100%  valign=top>
						   
										<TABLE <%=LR_SPACE_TYPE_20%>>
										            <TR>
														<TD COLSPAN=3>
															 2.총급여액 및 퇴직급여추계액 명세 
														</TD>
														
													</TR>
											       
													<TR>
														<TD HEIGHT="100%" COLSPAN=3>
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
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>


<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtW3value"     TAG="24">

<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" tabindex="-1"></TEXTAREA>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname"    TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"   TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar"  TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date"     TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

