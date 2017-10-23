
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 17호 조정후 수입금액명세서 
'*  3. Program ID           : W2107MA1
'*  4. Program Name         : W2107MA1.asp
'*  5. Program Desc         : 17호 조정후 수입금액명세서 
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

Const BIZ_MNU_ID = "W2107MA1"	
Const BIZ_PGM_ID = "W2107Mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID  = "W2107OA1"

Dim C_SEQ_NO
Dim C_IND_CLASS
Dim C_IND_TYPE 
Dim C_CODE
Dim C_RATE_NO
dim C_RATE_POPUP
Dim C_TATAL_AMT
Dim C_DOMESTIC_IN_AMT
Dim C_DOMESTIC_OUT_AMT
Dim C_EXPORT_AMT


Dim  C_SEQ_NO_2
dim  C_ITEM
dim  C_ITEM_CD
Dim  C_AMT
dim  C_REMARK

Const C_SHEETMAXROWS = 16









Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables(spd)

   if spd = "ALL" or spd ="A" then
		C_SEQ_NO			= 1
		C_IND_CLASS			= 2
		C_IND_TYPE			= 3
		C_CODE				= 4
		C_RATE_NO			= 5
		C_RATE_POPUP		= 6
		C_TATAL_AMT			= 7
		C_DOMESTIC_IN_AMT	= 8
		C_DOMESTIC_OUT_AMT	= 9
		C_EXPORT_AMT		= 10
   end if 
    
   if spd = "ALL" or spd ="B" then 

		C_SEQ_NO_2			= 1
		C_ITEM              = 2
		C_ITEM_CD           = 3
		C_AMT       		= 4
		C_REMARK       		= 5
	end if	
    
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

Sub SetDefaultVal()
    frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

    'call fncnew()
 
    

End Sub

'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet(strSPD)

    Call initSpreadPosVariables(strSPD)  
	 Call AppendNumberPlace("6","14","0")
	 Call AppendNumberPlace("8","15","0")
			  if (strSPD = "ALL" or strSPD ="A") then
			      	
				With frm1.vspdData
	
					ggoSpread.Source = frm1.vspdData	
					'patch version
					 ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
					 
						.ReDraw = false

					    .MaxCols = C_EXPORT_AMT + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
						.Col = .MaxCols														'☆: 사용자 별 Hidden Column
						.ColHidden = True    
						       
					 .MaxRows = 0
					 ggoSpread.ClearSpreadData
					 .ColHeaderRows(3)
				
					
					 
					 Call GetSpreadColumnPos("A")    
					
					 ggoSpread.SSSetEdit     C_SEQ_NO,			"", 3,,,100,1
					 ggoSpread.SSSetEdit     C_IND_CLASS,		"(1)업  태", 15,,,15
					 ggoSpread.SSSetEdit     C_IND_TYPE,		"(2)종  목", 20,,,15
					 ggoSpread.SSSetEdit     C_CODE	,			"코드",		 4,2,,2
					 ggoSpread.SSSetEdit     C_RATE_NO,			"(3)기준(단순)경비율번호", 13,,,7
					 ggoSpread.SSSetButton	 C_RATE_POPUP
					 ggoSpread.SSSetFloat	 C_TATAL_AMT,		"(4)계",	        16,		"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec   ,,,,"0"
					 ggoSpread.SSSetFloat	 C_DOMESTIC_IN_AMT, "국내생산품",		15,		"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec   ,,,,"0"
					 ggoSpread.SSSetFloat	 C_DOMESTIC_OUT_AMT,"수입상품",			15,		"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec   ,,,,"0"
					 ggoSpread.SSSetFloat	 C_EXPORT_AMT,		"수출상품",		    15,		"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec   ,,,,"0"

					 .AddCellSpan  1, -1000, 1, 3
					 .AddCellSpan  2, -1000, 1, 3
					 .AddCellSpan  3, -1000, 1, 3
					 .AddCellSpan  4, -1000, 1, 3
					 .AddCellSpan  5, -1000, 1, 3
					 .AddCellSpan  7, -1000, 4, 1
					 .AddCellSpan  7, -999, 1, 2
					 .AddCellSpan  8, -999, 2, 1
					 .AddCellSpan  10, -999, 1, 3
					 
					 
					 .col = 7
					 .row =-1000
					 .text =  "수입금액"
					 .col = 8
					 .row =-999
					 .text = "내  수" 
					 .col = 7
					 .row =-999 
					 .text = "(4)계[(5)＋(6)＋(7)]"
					 .col = 8
					 .row =-998
					 .text = "(5)국 내 생 산 품"
					 .col = 9
					 .row =-998 
					 .text = "(6)수 입 상 품"
					 .col = 10
					 .row =-999 
					 .text = "(7)수     출"
					 
					 
					 
					 


						Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
						'Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
									'
					
						.ReDraw = true
						 lgActiveSpd = "A"
				
						Call SetSpreadLock 		
	
				
					 
					 End With
			end if		 
				 	
    
          if strSPD = "ALL" or strSPD ="B" then

             ' Call GetSpreadColumnPos("B")  

					With frm1.vspdData2

							 ggoSpread.Source = frm1.vspdData2	
							'patch version
							 ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
							 
								.ReDraw = false

							 .MaxCols = C_REMARK + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
								.Col = .MaxCols														'☆: 사용자 별 Hidden Column
								.ColHidden = True    
								       
							 .MaxRows = 0
							 ggoSpread.ClearSpreadData
							
							 Call AppendNumberPlace("6","3","2")
							   
							  
							 ggoSpread.SSSetEdit     C_SEQ_NO_2, "순번", 10,,,100,1
							 ggoSpread.SSSetEdit     C_ITEM, "(14)구  분", 20,,,20
							 ggoSpread.SSSetEdit     C_ITEM_CD, "(15)코드", 8,2,,10
							 ggoSpread.SSSetFloat	 C_AMT, "(16)금  액",		15,	  	"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec 
							 ggoSpread.SSSetEdit     C_REMARK, "비  고", 20,,,16

							Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	     					Call ggoSpread.SSSetColHidden(C_SEQ_NO_2,C_SEQ_NO_2,True)
							

	
						.ReDraw = true
					     lgActiveSpd = "B"
			
					      Call SetSpreadLock 		
					End With  
		 end if
			
		
    
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx



End Sub


Sub SetSpreadLock()

  If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "A"
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "A"
           With frm1
    


				.vspdData.ReDraw = False
					ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
					'ggoSpread.SpreadLock C_IND_CLASS, -1, C_IND_CLASS
					ggoSpread.SpreadLock C_CODE		, -1, C_CODE
					'ggoSpread.SpreadLock C_IND_TYPE, -1, C_IND_TYPE
					ggoSpread.SpreadLock C_TATAL_AMT, -1, C_TATAL_AMT

				.vspdData.ReDraw = True

			End With
        Case  "B"
        
			Call SetSpreadColor2
        
   END SELECT     
   
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    ggoSpread.Source = frm1.vspdData
    With frm1

    .vspdData.ReDraw = False
     ggoSpread.SSSetProtected C_SEQ_NO , pvStartRow, pvEndRow
     'ggoSpread.SSSetProtected C_IND_CLASS , pvStartRow, pvEndRow
     'ggoSpread.SSSetProtected C_IND_TYPE , pvStartRow, pvEndRow
     ggoSpread.SSSetProtected C_CODE     , pvStartRow, pvEndRow
     ggoSpread.SSSetProtected C_RATE_NO, 11, 12
     ggoSpread.SSSetProtected C_RATE_POPUP, 11, 12
     ggoSpread.SSSetProtected C_TATAL_AMT, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected C_DOMESTIC_IN_AMT, 12, 12
     ggoSpread.SSSetProtected C_DOMESTIC_OUT_AMT, 12, 12
     ggoSpread.SSSetProtected C_EXPORT_AMT, 12, 12
     
     ggoSpread.SSSetProtected C_IND_CLASS , frm1.vspdData.maxrows -1 , frm1.vspdData.maxrows 
     ggoSpread.SSSetProtected C_IND_TYPE ,  frm1.vspdData.maxrows -1 , frm1.vspdData.maxrows
        
    .vspdData.ReDraw = True
    
    End With
    
End Sub

' -- 200603 개정 
Sub SetSpreadColor2()
    ggoSpread.Source = frm1.vspdData2
    With frm1

    .vspdData2.ReDraw = False

	ggoSpread.SpreadLock C_SEQ_NO_2, -1, C_SEQ_NO_2
	
	ggoSpread.SSSetProtected	C_ITEM	, 1, 12
	ggoSpread.SSSetProtected	C_ITEM	, 18, 18
	ggoSpread.SSSetProtected	C_ITEM_CD, 1, 18
	
	ggoSpread.SpreadLock C_SEQ_NO_2, 18, C_AMT, 18
        
    .vspdData2.ReDraw = True
    
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
          
            
            C_SEQ_NO			= iCurColumnPos(1)
			C_IND_CLASS			= iCurColumnPos(2)
			C_IND_TYPE			= iCurColumnPos(3)
			C_CODE				= iCurColumnPos(4)
			C_RATE_NO			= iCurColumnPos(5)
			C_RATE_POPUP		= iCurColumnPos(6)
			C_TATAL_AMT			= iCurColumnPos(7)
			C_DOMESTIC_IN_AMT	= iCurColumnPos(8)
			C_DOMESTIC_OUT_AMT	= iCurColumnPos(9)
			C_EXPORT_AMT		= iCurColumnPos(10)
			 
		Case "B"
             ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_SEQ_NO_2		= iCurColumnPos(1)
			C_ITEM			= iCurColumnPos(2)
			C_ITEM_CD		= iCurColumnPos(3)
			C_AMT			= iCurColumnPos(4)
			C_REMARK		= iCurColumnPos(5)		

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
			arrParam(0) = "표준소득율"								' 팝업 명칭 
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
			arrParam(5) = "표준소득율"									' 조건필드의 라벨 명칭 
	
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
				
				.vspdData.Col = C_RATE_NO
				.vspdData.Text = arrRet(0)
				
				Call vspdData_Change(C_RATE_NO, frm1.vspdData.activerow )	 ' 변경이 읽어났다고 알려줌 
		
		End Select
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function

Function FncCalSum(byval Row)
dim  w4
dim  w6
      
      
      
        w4=  unicdbl(FncSumSheet(frm1.vspdData,C_DOMESTIC_IN_AMT,Row,Row, false, -1, -1, "V")) +_
             unicdbl(FncSumSheet(frm1.vspdData,C_DOMESTIC_OUT_AMT,Row, Row , false, -1, -1, "V"))+_
             unicdbl(FncSumSheet(frm1.vspdData,C_EXPORT_AMT,Row, Row, false, -1, -1, "V"))
  
        frm1.vspdData.Row = Row 
        frm1.vspdData.Col = C_TATAL_AMT
        frm1.vspdData.text = w4
        
         Call vspdData_Change(C_TATAL_AMT,Row )	
         Call vspdData_Change(C_TATAL_AMT, frm1.vspdData.maxrows)	
End Function

Function GetValue4Grid(Byref pGrid, Byval pCol, Byval pRow)
	With pGrid
		.Col = pCol : .Row = pRow : GetValue4Grid = .Value
	End With
End Function

Function GetText4Grid(Byref pGrid, Byval pCol, Byval pRow)
	With pGrid
		.Col = pCol : .Row = pRow : GetText4Grid = Trim(.Text)
	End With
End Function

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
    
    Call InitSpreadSheet("ALL")                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
   
    Call SetToolbar("1100100000000111")  

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal
	
	Call FncQuery
    
End Sub


'============================================  이벤트 함수  ====================================





Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol
	Dim sMesg

	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	' 세무정보 조사 : 메시지가져오기.
	
	
	if wgConfirmFlg = "Y" then    '확정시 
	   Exit function
	end if   
	
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
	Call selectColor(frm1.txtW9)
    Call selectColor(frm1.txtW10)
    Call selectColor(frm1.txtW11)

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	Call ggoOper.LockField(Document, "N") 
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If


   '***참조 
   '***TB_WORK_8
       

	call CommonQueryRs("W9,W10, W11","dbo.ufn_TB_17_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

       frm1.txtW9.value =  unicdbl(lgF0) 
       frm1.txtW10.value = unicdbl(lgF1)
       frm1.txtW11.value = unicdbl(lgF2)    
   
      lgBlnFlgChgValue = TRUE

End Function





Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub txtw8_Change( )

   	lgBlnFlgChgValue = True
   
    Frm1.txtw12.value = unicdbl(Frm1.txtw9.value) + unicdbl(Frm1.txtw10.value) + unicdbl(Frm1.txtw11.value)
	Frm1.txtw14.value = unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw13.value)

End Sub



Sub txtw9_Change( )

 
   	lgBlnFlgChgValue = True
    Frm1.txtw12.value = unicdbl(Frm1.txtw9.value) + unicdbl(Frm1.txtw10.value) + unicdbl(Frm1.txtw11.value)
	Frm1.txtw14.value = unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw13.value)

End Sub

Sub txtw10_Change( )

 
    lgBlnFlgChgValue = True
    Frm1.txtw12.value = unicdbl(Frm1.txtw9.value) + unicdbl(Frm1.txtw10.value) + unicdbl(Frm1.txtw11.value)
	Frm1.txtw14.value = unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw13.value)

End Sub

Sub txtw11_Change()

 
    lgBlnFlgChgValue = True
    Frm1.txtw12.value = unicdbl(Frm1.txtw9.value) + unicdbl(Frm1.txtw10.value) + unicdbl(Frm1.txtw11.value)
	Frm1.txtw14.value = unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw13.value)

End Sub

Sub txtw12_Change()
    lgBlnFlgChgValue = True
    Frm1.txtw12.value = unicdbl(Frm1.txtw9.value) + unicdbl(Frm1.txtw10.value) + unicdbl(Frm1.txtw11.value)
	Frm1.txtw14.value = unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw13.value)

End Sub

Sub txtw13_Change()
    lgBlnFlgChgValue = True
    Frm1.txtw12.value = unicdbl(Frm1.txtw9.value) + unicdbl(Frm1.txtw10.value) + unicdbl(Frm1.txtw11.value)
	Frm1.txtw14.value = unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw13.value)

End Sub


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        
        If Row > 0 And Col = C_RATE_POPUP Then
            .Col = Col - 1
            .Row = Row
            Call OpenPopup(.Text, 5)

        End If
    End With
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
	Dim intRow
	Dim intIndex 
	

	
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
     Dim iDx
    Dim IntRetCD, sWhere
    Dim i
    Dim w13,w5,w6,w7
 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
  '------ Developer Coding part (Start ) -------------------------------------------------------------- 
  
  '--------------------'그리드에 입력된 내역이 기존데이터에 있을때 체크----------------------------------
    Select Case Col
        Case C_RATE_NO
			If frm1.txtFISC_YEAR.text >= "2006" Then
				sWhere = " AND ATTRIBUTE_YEAR = '2005' " 
			Else
				sWhere = " AND ATTRIBUTE_YEAR = '2003' " 
			End If
        
            IntRetCD =  CommonQueryRs(" BUSNSECT_NM ,DETAIL_NM  ","tb_std_income_rate"," STD_INCM_RT_CD = '" & Trim(frm1.vspdData.text) & "'" & sWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
            If IntRetCD = False Then
               Call  DisplayMsgBox("970000","X","표준소득률","X")                         '☜ : 입력된자료가 있습니다.
               frm1.vspdData.text =""
            Else
           
              
            End If
        Case C_DOMESTIC_IN_AMT 
            
            w5 = FncSumSheet(frm1.vspdData,C_DOMESTIC_IN_AMT, 1, frm1.vspdData.MaxRows - 1, false, -1, -1, "V")
           
            frm1.vspdData.Row = frm1.vspdData.MaxRows
			frm1.vspdData.Col = C_DOMESTIC_IN_AMT
			frm1.vspdData.text = w5

             Call FncCalSum(Row)
    
         Case C_DOMESTIC_OUT_AMT  
             w6= FncSumSheet(frm1.vspdData,C_DOMESTIC_OUT_AMT, 1, frm1.vspdData.MaxRows - 1, false, -1, -1, "V")
           
            frm1.vspdData.Row = frm1.vspdData.MaxRows
			frm1.vspdData.Col = C_DOMESTIC_OUT_AMT
			frm1.vspdData.text = w6

             Call FncCalSum(Row)
        
        Case C_EXPORT_AMT  
              w7 = FncSumSheet(frm1.vspdData,C_EXPORT_AMT, 1, frm1.vspdData.MaxRows - 1, false, -1, -1, "V")
           
            frm1.vspdData.Row = frm1.vspdData.MaxRows
			frm1.vspdData.Col = C_EXPORT_AMT
			frm1.vspdData.text = w7
			 Call FncCalSum(Row)
         
        Case C_TATAL_AMT
        
            w13 = FncSumSheet(frm1.vspdData,C_TATAL_AMT, 1, frm1.vspdData.MaxRows - 1, false, -1, -1, "V")
           
            frm1.vspdData.Row = frm1.vspdData.MaxRows
			frm1.vspdData.Col = C_TATAL_AMT
			frm1.vspdData.text = w13
            frm1.txtw13.text = w13
    End Select
    
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	lgBlnFlgChgValue = True
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub


Sub vspdData2_Change(ByVal Col , ByVal Row )
     Dim iDx
    Dim IntRetCD
    Dim i
    Dim w17
 
  '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	With Frm1.vspdData2

		.Row = Row
		.Col = Col
  '--------------------'그리드에 입력된 내역이 기존데이터에 있을때 체크----------------------------------
     Select Case Col
        Case C_AMT
         
        
            w17 = FncSumSheet(frm1.vspdData2,C_AMT, 1, .MaxRows-1, false, -1, -1, "V")
            
            .Row = .MaxRows
            .Col = C_AMT
            .Value = w17
			
    End Select
    
    End With
    lgBlnFlgChgValue = True
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 

	
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row
	ggoSpread.UpdateRow frm1.vspdData2.MaxRows
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
   

	frm1.vspdData2.Row = Row
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
    Call GetSpreadColumnPos("B")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
   
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("All")      
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

Function FncNew() 
    Dim IntRetCD 
    dim row
   Dim sFiscYear, sRepType, sCoCd
	
	
	
    
    FncNew = False                                                          
    

	If lgBlnFlgChgValue = true  Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

    Call ggoOper.ClearField(Document, "2")

    Call InitVariables               

    Call SetToolbar("1100100000000111")          '⊙: 버튼 툴바 제어 
    

       
    
  
	With frm1	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
		 ggoSpread.InsertRow ,12
         SetSpreadColor -1, -1
                        
           
      FOR ROW = 1 TO     .vspdData.maxrows  	
	
		.vspdData.ReDraw = True		
		.vspdData.col = c_seq_no
		.vspdData.row = row 
		.vspdData.text = 100 + row
		.vspdData.col = c_code
		.vspdData.row = row 
		if row >=10 then
		   .vspdData.text = row
		else
			.vspdData.text = "0" & row
		end if 
		

		
		.vspdData.col = 2 
		.vspdData.row = 11 
		.vspdData.text = "기타"
		.vspdData.row = 12 
		.vspdData.text = "합계" 
		.vspdData.col = c_code
		.vspdData.row = 12 
		.vspdData.text = "99" 
	  next	
	  
	    sCoCd		= frm1.txtCO_CD.value
		sFiscYear	= frm1.txtFISC_YEAR.text
		sRepType	= frm1.cboREP_TYPE.value

	
	  if  .vspdData.maxrows   > 0 then
	      .vspdData.col  = C_SEQ_NO
	      .vspdData.Row  = 1

	      if Trim(.vspdData.text) = "101" then 
	         IntRetCD =  CommonQueryRs("IND_CLASS,IND_TYPE,HOME_TAX_MAIN_IND"," TB_COMPANY_HISTORY "," CO_CD = '" & sCoCd &"' AND FISC_YEAR = '" & sFiscYear &"' AND   REP_TYPE = '" & sRepType &"' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	         IF IntRetCD = TRUE THEN
	          .vspdData.col =  C_IND_CLASS
	          .vspdData.text = REPLACE(lgF0,CHR(11),"")
	          .vspdData.col =  C_IND_TYPE
	          .vspdData.text = REPLACE(lgF1,CHR(11),"")
	          .vspdData.col =  C_RATE_NO
	          .vspdData.text = REPLACE(lgF2,CHR(11),"")
	          
	         END IF 
		       
	      end if
	   end if
	  
    End With
        
   '-- 수입금액차액검토 
   Dim iRow, iMaxRows, arrW14, arrW15
   
   	call CommonQueryRs("MINOR_NM, MINOR_CD"," ufn_TB_MINOR('W1092','" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 <> "" Then
		arrW14 = Split(lgF0, Chr(11))
		arrW15 = Split(lgF1, Chr(11))
		
		iMaxRows = UBound(arrW14, 1)
		
		With frm1.vspdData2

			.focus
			ggoSpread.Source = frm1.vspdData2
		
			.ReDraw = False
			 ggoSpread.InsertRow ,iMaxRows
			 SetSpreadColor2
		
			For iRow = 1 To iMaxRows
				.Row = iRow
				.Col = C_SEQ_NO_2	: .Value = iRow
				.Col = C_ITEM		: .Text = arrW14(iRow-1)
				.Col = C_ITEM_CD	: .Text = arrW15(iRow-1)
			Next
			
			.ReDraw = True
		End With
	
	End If

	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData

    FncNew = True                

End Function



Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If  lgBlnFlgChgValue = true Then
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
    DIM IntRetCD 
    Dim strsql
    Dim dblTotal
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If 	lgBlnFlgChgValue = False  Then
                 ggoSpread.Source = frm1.vspdData2
			If  ggoSpread.SSCheckChange = False Then
			    IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
			    Exit Function
			End If
       
    End If
    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  
    
    ggoSpread.Source = frm1.vspdData2
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If
	



	

	 dblTotal = unicdbl(GetValue4Grid(frm1.vspdData2, C_AMT, frm1.vspdData2.MaxRows))
	 
	 if unicdbl(frm1.txtW14.text) <> dblTotal  then
	     Call  DisplayMsgBox("WC0004","X","(13)차액 ["&  unicdbl(frm1.txtW14.text)   & "]","차이내역의 합계 [" & dblTotal  & "] 금액" )  
	    
	    exit function
	 end if
	 
	 
	 
	 
		frm1.vspdData.row =  frm1.vspdData.maxrows 
		frm1.vspdData.col =  C_TATAL_AMT
		dblTotal = unicdbl(frm1.vspdData.text)
	 
            	
		if dblTotal > 10000000 and unicdbl(frm1.txtW12.text) <=  0 then
	        Call  DisplayMsgBox("WC0006","X",frm1.txtW12.alt ,0)  
	    
			exit function
		end if

        
	 
	 
	

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

    If Frm1.vspdData2.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData2

	With frm1
		If .vspdData2.ActiveRow > 0 Then
			.vspdData2.focus
			.vspdData2.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow

			
			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData2	
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
    
    
    if  imRow > 8 then 
        imRow = 8
    end if
    
    
  
	With frm1	
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
		
		.vspdData2.ReDraw = False
		if .vspdData2.MaxRows  < 8 then
		   'iSeqNo = .vspdData.MaxRows+1
            ggoSpread.InsertRow ,imRow
           ' SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
		
             MaxSpreadVal .vspdData2, C_SEQ_NO_2, .vspdData2.ActiveRow
             SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow
        end if   		
            
           
            	
		.vspdData2.ReDraw = True		
	
    End With
    
   
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData2
    	.focus
    	ggoSpread.Source = frm1.vspdData2 
    	lDelRows = ggoSpread.DeleteRow
    	
    	cALL vspdData2_Change(C_AMT,.ACTIVEROW)
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
    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '☜: Processing is NG
    
    
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    
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

    If DbDelete= False Then
       Exit Function
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function


'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
       
        
        
        
        	With frm1
			
		
					strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
					strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
					strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
					strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 
	
									

    End With
    

		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function


Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
										<%'버튼 툴바 제어 %>
    Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call InitData
	'1 컨펌체크 
	If wgConfirmFlg = "Y" Then

	    Call SetToolbar("1100000000000111")	
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		  Call SetToolbar("1101100000000111")									<%'버튼 툴바 제어 %>
	
	End If
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
    Call SetSpreadColor(-1,-1)
    Call SetSpreadColor2
   
  								<%'버튼 툴바 제어 %>
    lgBlnFlgChgValue =  False

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
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        

        
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C"  &  Parent.gColSep
                                                  strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_IND_CLASS          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_IND_TYPE           : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_CODE               : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_RATE_NO            : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_TATAL_AMT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep   
                    .vspdData.Col = C_DOMESTIC_IN_AMT    : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_DOMESTIC_OUT_AMT   : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_EXPORT_AMT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep   

                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep
                                                  strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_IND_CLASS          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_IND_TYPE          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_CODE             : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_RATE_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_TATAL_AMT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep   
                    .vspdData.Col = C_DOMESTIC_IN_AMT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_DOMESTIC_OUT_AMT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_EXPORT_AMT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep   
                    
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D"  &  Parent.gColSep
                                                  strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
		 .txtMode.value        =  Parent.UID_M0002
	
		 .txtFlgMode.value     = lgIntFlgMode
		'.txtUpdtUserId.value  =  Parent.gUsrID
		'.txtInsrtUserId.value =  Parent.gUsrID
		 .txtMaxRows.value     = lGrpCnt-1 
		 .txtSpread.value      = strDel & strVal
		
	End With	
	ggoSpread.Source = frm1.vspdData2 	
	
	 strVal = ""
	 strDel = ""
	 lGrpCnt = 1
	With Frm1
		For lRow = 1 To .vspdData2.MaxRows
    
           .vspdData2.Row = lRow
           .vspdData2.Col = 0
           
           Select Case .vspdData2.Text
              

        
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C"  &  Parent.gColSep
                                                  strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO_2          : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_ITEM          : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_ITEM_CD       : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_AMT          : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_REMARK       : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep
                                                  strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO_2          : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_ITEM          : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_ITEM_CD       : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_AMT          : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_REMARK       : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
                    
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
		.txtMaxRows2.value     = lGrpCnt-1 
		.txtSpread2.value      = strDel & strVal
 
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

	strVal = BIZ_PGM_ID & "?txtMode=" &  parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
	strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 
	strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey

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
	Call SetToolbar("1100100000000111")          '⊙: 버튼 툴바 제어 
End Function



'========================================================================================



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						
					</TD>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto">
							<TABLE  CLASS="TB3" CELLSPACING=0 BORDER=0>	
							    <TR>
								    
										<TD WIDTH=100% HEIGHT="250" valign=top>
										     <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>1.업종별 수입금액명세서 </LEGEND>
											
														<TABLE <%=LR_SPACE_TYPE_20%>>
															
																	<TR>
																		<TD HEIGHT="100%" COLSPAN=3>
																			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=250 tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
																		</TD>
																		
																	</TR>
														
														</TABLE>
											 </FIELDSET>
										</TD>
								</TR>
	
				
								<TR>
									<TD WIDTH=100% valign=top >
												   
									      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>2.부가가치세 과세표준과 수입금액 차액검토 </LEGEND>
													<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
																   
																	
														<TR>
															<TD CLASS="TD51" align =center   >
																(8)과세(일반 )
															</TD>
															<TD CLASS="TD51" align =center >
																(9)과세(영세율) 
															</TD>
															<TD CLASS="TD51" align =center>
																(10)면세수입금액 
															<TD CLASS="TD51" align =center>
																(11)합계[(8)＋(9)+(10)]
															</TD>
															<TD CLASS="TD51" align =center>
																(12)수입금액 
															</TD>
															<TD CLASS="TD51" align =center>
																(13)차액[(11)－(12)]
															</TD>
														</TR>
														<TR>
															<TD CLASS="TD61" align =center width = 10%><input type=hidden id="txtW8" name=txtW8 tag="24X2Z">
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW9" name=txtW9 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW11" name=txtW11 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12" name=txtW12 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="(11)합계[(8)+(9)+(10)]" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW13" name=txtW13 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW14" name=txtW14 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
															</TD>
														</TR>
																	
													
													</TABLE>
										   </FIELDSET>				
													   			
									</TD>
								</TR>
											
											
								<TR>
									<TD WIDTH=100% HEIGHT=100% valign=top>
									   	<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>[차액내역]</LEGEND>		
													<TABLE width = 100%  HEIGHT=150 bgcolor = #696969  border = 0 cellpadding =0 cellspacing = 0 ID="Table2">	
											
																     
															<TR>		
																	<TD WIDTH=100% HEIGHT="150" valign=top>
																		<TABLE <%=LR_SPACE_TYPE_20%>>
																			<TR>
																				<TD >
																					<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=150 tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
																				</TD>
																			</TR>
																		</TABLE>
																	</TD>
															</TR>
													</TABLE>
											</FIELDSET>		
														  			
									</TD>
								</TR>			
									
							</TABLE>
						 </div>	
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

<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" tabindex="-1" style="display: 'none'"></TEXTAREA>
<TEXTAREA CLASS=hidden NAME=txtSpread2 tag="24" tabindex="-1" style="display: 'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24">
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

