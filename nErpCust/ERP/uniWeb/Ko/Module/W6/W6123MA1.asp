
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 제8호부표2공제감면세액계산서(2)
'*  3. Program ID           : W6123MA1
'*  4. Program Name         : W6123MA1.asp
'*  5. Program Desc         : 제8호부표2공제감면세액계산서(2)
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/02/08
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID = "W6123MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_MNU_ID = "W6123MA1"
Const EBR_RPT_ID		= "W6123OA1"

Dim C_W1           ' (1)구분코드 
Dim C_W1_POP	   ' 팝업	
Dim C_W1_Nm		   ' (1)구분
Dim C_W1_Ref	   ' 근거법 조항 
Dim C_W2		   ' (2)계산내역 
Dim C_W2_POP       ' 팝업 
Dim C_W3		   ' (3)감면대상세액 
dim C_W4		   ' (4)최저한세 적용감면배제금액 
dim C_W5		   ' (5)감면세액	
dim C_W6		   ' 사유발생일 
dim C_W2_A		   ' 감면소득 
dim C_W2_B		   ' 3호서식 113 과세표준 
dim C_W2_C_VIEW	   ' 감면율	
Dim C_W2_C_VALUE
Dim C_W2_D         '산출세액 


dim strMode



Const C_SHEETMAXROWS = 100









Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_W1		= 1   ' (1)구분코드 
	C_W1_Nm		= 2   ' 구분명 
	C_W1_Ref	= 3   ' 근거법 조항 
	C_W2		= 4   ' (2)계산내역 
	C_W2_POP    = 5   ' 팝업 
	C_W3		= 6   ' (3)감면대상세액 
	C_W4		= 7   ' (4)최저한세 적용감면배제금액 
	C_W5		= 8   ' (5)감면세액	
	C_W6		= 9   ' 사유발생일 
	C_W2_A		= 10  ' 감면소득 
	C_W2_B		= 11  ' 3호서식 113 과세표준 
	C_W2_C_VIEW	= 12  ' 감면율	표시 
	C_W2_C_VALUE= 13  ' 감면율	값 
	C_W2_D      = 14   ' 산출세액 

    
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
			

	        lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
			lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
			lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '

   


    
 

End Sub 
'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub




Sub SetDefaultVal()


End Sub
Sub InitSpreadSheet()

    Call initSpreadPosVariables()  

			      	
				With frm1.vspdData
	
					ggoSpread.Source = frm1.vspdData	
					'patch version
					 ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
					 
						.ReDraw = false

					    .MaxCols = C_W2_D + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
						.Col = .MaxCols														'☆: 사용자 별 Hidden Column
						.ColHidden = True    
						       
					 .MaxRows = 0
					 ggoSpread.ClearSpreadData
                     .rowheight(0) = 28
					 Call AppendNumberPlace("6","3","2")
				 
					 Call GetSpreadColumnPos("A")    
					 
			
					 'ggoSpread.SSSetCombo    C_W1,		"(1)구분코드", 8
					 'ggoSpread.SSSetCombo    C_W1_Nm,	"(1)구분", 35
					 ggoSpread.SSSetEdit     C_W1,	"(1)구분", 35,,,50,1
					 ggoSpread.SSSetEdit     C_W1_Nm,	"(1)구분", 35,,,50,1
				     ggoSpread.SSSetEdit     C_W1_Ref,	"근거법조항", 15,,,100,1
					 ggoSpread.SSSetEdit     C_W2,		"(2)계산내역",  40,,,100,1
					 ggoSpread.SSSetButton   C_W2_POP	 
					 ggoSpread.SSSetFloat    C_W3,		"(3) 감면대상세액",				15,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W4,		"(4) 최저한세적용감면배제금액", 15,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W5,		"(5)감면세액((3)-(4))",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0"
					 ggoSpread.SSSetDate     C_W6,		"(6)사유발생일"					, 11,2,  parent.gDateFormat
					 ggoSpread.SSSetFloat    C_W2_A,	"감면소득액"				    , 15,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetFloat    C_W2_B,	"3호서식과세표준", 15,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 ggoSpread.SSSetEdit     C_W2_C_VIEW,	"감면율표시", 10,,,100,1
					 ggoSpread.SSSetEdit     C_W2_C_VALUE,	"감면율", 10,,,100,1
					
					 ggoSpread.SSSetFloat    C_W2_D,		"3호서식산출세액",				15,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					 
    

						Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
						Call ggoSpread.SSSetColHidden(C_W1, C_W1, True)
						Call ggoSpread.SSSetColHidden(C_W2_A, C_W2_D, True)
						Call ggoSpread.SSSetColHidden(C_W2_POP,C_W2_POP, True)
	
		
					
						.ReDraw = true

				 
					Call SetSpreadLock 		
	
				
					 
					 End With


    
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim IntRetCD1

	'IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", " ufn_TB_MINOR('W1017', '" & C_REVISION_YM & "')", " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	'If IntRetCD1 <> False Then
	'	ggoSpread.Source = frm1.vspdData
	'	ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W1
		'ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W1_Nm
	'End If

End Sub

Sub SetSpreadLock()


           With frm1
    


				.vspdData.ReDraw = False
                  
					ggoSpread.SpreadLock C_w1, -1, C_w1
					'ggoSpread.SpreadLock C_w2, -1, C_w2
					'ggoSpread.SpreadLock C_w3, -1, C_w3
					
				.vspdData.ReDraw = True

			End With
 
   
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
dim sumRow
    ggoSpread.Source = frm1.vspdData
    With frm1

    .vspdData.ReDraw = False

   
        
		 ggoSpread.SSSetProtected C_w1 , pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired C_w1_Nm , pvStartRow, pvEndRow
		 'ggoSpread.SSSetProtected C_w1_ref , pvStartRow, pvEndRow
		  'ggoSpread.SSSetProtected C_w2 , pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired C_w2 , pvStartRow, pvEndRow
		
		 'ggoSpread.SSSetProtected C_w3 , pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected C_w5 , pvStartRow, pvEndRow


   
        
    .vspdData.ReDraw = True
    
    End With
End Sub






Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
        
				
				C_W1		= iCurColumnPos(1)           ' (1)구분코드 
				C_W1_Nm		= iCurColumnPos(2)    ' 구분명 
			
				C_W1_Ref	= iCurColumnPos(3)   ' 근거법 조항 
				C_W2		= iCurColumnPos(4)   ' (2)계산내역 
				C_W2_POP    = iCurColumnPos(5)   ' 팝업 
				C_W3		= iCurColumnPos(6)   ' (3)감면대상세액 
				C_W4		= iCurColumnPos(7)   ' (4)최저한세 적용감면배제금액 
				C_W5		= iCurColumnPos(8)   ' (5)감면세액	
				C_W6		= iCurColumnPos(9)   ' 사유발생일 
				C_W2_A		= iCurColumnPos(10)   ' 감면소득 
				C_W2_B		= iCurColumnPos(11)   ' 3호서식 113 과세표준 
				C_W2_C_VIEW	= iCurColumnPos(12)   ' 감면율	표시 
				C_W2_C_VALUE= iCurColumnPos(13)   ' 감면율	값 
				C_W2_D      = iCurColumnPos(14)   ' 3호서식 120 산출세액 

	

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
	Dim strMajor , strMinor
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0

		Case 3
			arrParam(0) = "구분"								' 팝업 명칭 
			arrParam(1) = "ufn_B_Configuration('W1017')" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""										' Where Condition
			arrParam(5) = "세액감면유형"									' 조건필드의 라벨 명칭 
            
			arrField(0) = "Minor_Cd "									' Field명(0)
			arrField(1) = "Minor_Nm"									' Field명(1)
			arrField(2) = "Reference_1"									' Field명(1)
			arrField(3) = ""									' Field명(1)
			
			arrHeader(0) = "구분"									' Header명(0)
			arrHeader(1) = "세액감면유형"									' Header명(1)
			arrHeader(2) = "근거법 조항"									' Header명(1)
			
			
	    Case 5
	
		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	if iWhere = 5 then	
	        Frm1.vspdData.Row = Frm1.vspdData.activerow
	        Frm1.vspdData.Col = C_W1
			strMinor = Trim(Frm1.vspdData.Text )
			if  Trim(strMinor) = "" then
			    IntRetCD = DisplayMsgBox("X", "X", "구분을 입력 해 주세요", "X") 
			    IsOpenPop = False
			    Exit Function    
			end if
			

			
			
			Dim sFiscYear, sRepType, sCoCd, IntRetCD
			sCoCd		= frm1.txtCO_CD.value
			sFiscYear	= frm1.txtFISC_YEAR.text
			sRepType	= frm1.cboREP_TYPE.value	
			
			  IntRetCD = CommonQueryRs("W16","TB_3", " CO_CD = '" & sCoCd & "' and FISC_YEAR = '" & sFiscYear & "' and REP_TYPE = '" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			if  IntRetCD = False  then
    
			    IntRetCD = DisplayMsgBox("W60006", "x", "(120) 산출세액"  , "X")     
			     IsOpenPop = False
			    Exit Function
			end if    
		
			
			'참조한 감면율 메이저 코드 
			call CommonQueryRs("Reference_2"," ufn_TB_Configuration('W1017') "," Minor_cd = '"& strMinor  &"' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   '팝업에서 참조할 메이저 코드 
            strMajor =Trim(Replace(lgF0,Chr(11),""))  
          
			If strMajor = "" Then
				IsOpenPop = False
				Call DisplayMsgBox("X", "X", "계산기준표를 참조하시고, (2)계산내역에 직접 결과금액을 입력하십시오 ", "X") 
				Exit Function
			End If
			
     	    arrRet = window.showModalDialog("w6123ra1.asp?sCoCd=" & sCoCd & "&sFiscYear="&sFiscYear&"&sRepType="&sRepType&"&strMajor="&strMajor, Array(window.parent), _
	             "dialogWidth=850px; dialogHeight=160px; center: Yes; help: No; resizable: No; status: No;")
	else
		     arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	end if		

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
			Case 3
				.vspdData.Col = C_W1
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_W1_Nm
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_W1_Ref
				.vspdData.Text = arrRet(2)
				
				'Call vspdData_Change(C_W1, frm1.vspdData.activerow )	 ' 변경이 읽어났다고 알려줌 
			Case 5
				.vspdData.Col = C_W3           '감면세액                    '
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_W2_A         '감면소득 
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_W2_B         '과세표쥰 
				.vspdData.Text = arrRet(2)
			    .vspdData.Col = C_W2_C_VALUE
				.vspdData.Text = arrRet(3)
				.vspdData.Col = C_W2_C_VIEW
				.vspdData.Text = arrRet(4)

				.vspdData.Col = C_W2_D       '산출세액 
				.vspdData.Text = arrRet(5)
				
				.vspdData.Col = C_W2           '내역 
				.vspdData.Text =  formatnumber(arrRet(5),0) & " x (" & formatnumber(arrRet(1),0)  &" / " & formatnumber(arrRet(2),0) &") x" & arrRet(4)
			
			
	
				
		
				Call vspdData_Change(C_W3, frm1.vspdData.activerow )	 ' 변경이 읽어났다고 알려줌 	
		
		End Select
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
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
    Call InitComboBox
    Call InitSpreadComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
     

    Call FncQuery
End Sub


'============================================  이벤트 함수  ====================================
Function CheckReCalc()
	Dim dblSum
	
	With frm1.vspdData
		ggoSpread.Source =frm1.vspdData
	
        if  frm1.vspdData.maxrows =< 0 then exit function
        
         


		dblSum = FncSumSheet(frm1.vspdData, C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' 합계 
		dblSum = FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' 합계 
		dblSum = FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' 합계 
		   ggoSpread.Source = frm1.vspdData
           ggoSpread.UpdateRow frm1.vspdData.maxrows

	End With
End Function





'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        
        If Row > 0 And Col = C_W1_POP Then
            .Col = Col - 2
            .Row = Row
            Call OpenPopup(.Text, 3)

        End If
        
         If Row > 0 And Col = C_W2_POP Then
            .Col = Col - 1
            .Row = Row
            Call OpenPopup(.Text, 5)

        End If
        
        
    End With
End Sub

Sub vspdData_ComboSelChange( ByVal Col, ByVal Row)
	Dim iIdx
	 
	With  frm1.vspdData
	  ggoSpread.Source = frm1.vspdData

		Select Case Col
			Case C_W1 , C_W1_Nm
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
				
		
		End Select
	End With
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
    Dim i
    Dim w13,w5,w6,w7 ,w3, w4
 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
  '------ Developer Coding part (Start ) -------------------------------------------------------------- 

  '--------------------'그리드에 입력된 내역이 기존데이터에 있을때 체크----------------------------------
    Select Case Col
        Case C_W1 , C_W1_nm
        
' -- 200603 개정 : 직접 입력하게 바꿈           
'            Frm1.vspdData.col = C_W1
'            call CommonQueryRs("Reference_1"," ufn_B_Configuration('W1017') "," Minor_CD = '" & Trim( Frm1.vspdData.text) &"'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'            Frm1.vspdData.col = C_W1_Ref
'            Frm1.vspdData.row = row
'            Frm1.vspdData.text = replace(lgF0,chr(11),"")
'            Frm1.vspdData.col = C_W2
'            Frm1.vspdData.text =""
            
'            Frm1.vspdData.col = C_W3
'            Frm1.vspdData.text =0
            
'            Frm1.vspdData.col = C_W5
'            Frm1.vspdData.text =0
'           call vspdData_Change(C_W3,row)
           
        Case C_W4 ,C_W3

            Frm1.vspdData.Row = Row
            Frm1.vspdData.col = C_W3
            w3 = Frm1.vspdData.text
      
            Frm1.vspdData.col = C_W4
            w4 = Frm1.vspdData.text
            Frm1.vspdData.col = C_W5

            Frm1.vspdData.text = unicdbl(w3) - unicdbl(w4)
            Call CheckReCalc
             
        
                          
        
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
    Call SetPopupMenuItemInf("11010000000") 

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
    Call GetSpreadColumnPos("B")
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




Sub SetSpreadTotalLine()
	Dim iRow

		ggoSpread.Source =  frm1.vspdData
		With  frm1.vspdData
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1: .CellType = 1	: .Text = "999999"	: .TypeHAlign = 2
				.Col = C_W1_Nm: .CellType = 1	: .Text = "계"	: .TypeHAlign = 2
				ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
				
			
			End If
		End With

End Sub 


'============================================  툴바지원 함수  ====================================
'=====================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

     Call ggoOper.ClearField(Document, "2")

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
     Call MakeKeyStream("X")															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
	
     
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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("800442", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    
    
    If lgIntFlgMode=True  Then                                            'Check if there is retrived data
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
    
    If ggoSpread.SSCheckChange = False   Then
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
    
   Call MakeKeyStream("X")
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
    if frm1.vspdData.maxrows <> frm1.vspdData.activerow then
       ggoSpread.EditUndo                                                  '☜: Protect system from crashing
     
    end if 

   if frm1.vspdData.maxrows = 1 then
      ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    end if 
   Call CheckReCalc    
    
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
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If

	With frm1	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		'.vspdData.ReDraw = False
		iSeqNo = .vspdData.MaxRows+1
	
      
        
		if 	.vspdData.MaxRows = 0 then
		
		     ggoSpread.InsertRow  imRow 
		     SetSpreadColor 1, 1
		
		     ggoSpread.InsertRow  imRow 
		     .row = .vspdData.MaxRows

			Call SetSpreadTotalLine
			


    
    
			
		else
				'.vspdData.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
				
		     
				iRow = .vspdData.ActiveRow
		
				If iRow = .vspdData.MaxRows Then
				    .vspdData.ActiveRow  = .vspdData.MaxRows -1
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColor iRow, iRow
    
				
					For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
					
						Call MaxSpreadVal(frm1.vspdData, C_W1, ii)
						
					Next
					Call SetSpreadColor(iRow , (iRow-1) + imRow)   
				Else
				
				  
			
		            ggoSpread.InsertRow ,imRow
		            For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
					
						Call MaxSpreadVal(frm1.vspdData, C_W1, ii)
						
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
    if frm1.vspdData.maxrows <> frm1.vspdData.activerow then
       lDelRows = ggoSpread.DeleteRow                                              '☜: Protect system from crashing
       Call CheckReCalc
       
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
     Call SetSpreadColor( -1, -1)
    Call SetSpreadTotalLine
    Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call InitData
	'1 컨펌체크 
	If wgConfirmFlg = "Y" Then

	    Call SetToolbar("1100000000000111")	
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		  Call SetToolbar("1101111100000111")									<%'버튼 툴바 제어 %>
	
	End If

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

	strVal = BIZ_PGM_ID & "?txtMode=" &  parent.UID_M0003                                '☜: Delete
	strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
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
                    .vspdData.Col = C_w1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_w1_Nm       : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_w3          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W4          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W5          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W6          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_a          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_b          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_c_view          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_c_value          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_d          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep   

					' -- 2006-3: 개정서식 반영  : 기존 팝업 방식을 제거하고 직접 입력하게 수정
                    .vspdData.Col = C_W1_ref          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep   
                    .vspdData.Col = C_W2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep   
 

                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_w1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_w1_Nm       : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_w3          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W4          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W5          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W6          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_a          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_b          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_c_view          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_c_value          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2_d          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep   
 
					' -- 2006-3: 개정서식 반영  : 기존 팝업 방식을 제거하고 직접 입력하게 수정
                    .vspdData.Col = C_W1_ref          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep   
                    .vspdData.Col = C_W2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep   
                    
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strVal = strVal & "D"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_w1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
		 .txtMode.value        =  Parent.UID_M0002
		'.txtUpdtUserId.value  =  Parent.gUsrID
		 frm1.txtKeyStream.value      =  lgKeyStream
		.txtMaxRows.value     = lGrpCnt-1 
		.txtSpread.value      = strDel & strVal

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

Function GetRef()	'계산기준 별표 링크 클릭시 

	Call window.open("W6123RA1.htm", BIZ_MNU_ID, _
	"Width=640px,Height=500px,center= Yes,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes")

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
					<TD >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" ><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;<a href="vbscript:GetRef">계산기준 별표1</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w6123ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
				    
						<TD WIDTH=100%  valign=top>
						   
										<TABLE <%=LR_SPACE_TYPE_20%>>
										            <TR>
														<TD COLSPAN=3>
															
														</TD>
														
													</TR>
											       
													<TR>
														<TD HEIGHT="100%" COLSPAN=3>
															<script language =javascript src='./js/w6123ma1_vaSpread1_vspdData.js'></script>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><별지><%=Request("strASPMnuMnuNm")%></LABEL>&nbsp;
				           
				        </TD>
				            
                </TR>
		    
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
<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" tabindex="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"     TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

