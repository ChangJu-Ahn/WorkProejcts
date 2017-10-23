<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 자동추출 계정과목 
'*  3. Program ID           : WB112MA2
'*  4. Program Name         : WB112MA2.asp
'*  5. Program Desc         : 자동추출 계정과목 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "WB112MA1"
Const BIZ_PGM_ID = "WB112MB1.asp"											 '☆: 비지니스 로직 ASP명 
 

Const TYPE_1 = 0
Const TYPE_2 = 1

Dim C_MAP_CD
Dim C_MAP_NM
Dim C_ACCT_CD
Dim C_ACCT_POP
Dim C_ACCT_NM
Dim C_ACCT_GP_CD
Dim C_ACCT_GP_POP
Dim C_ACCT_GP_NM

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid , lgvspdData(1)   
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_MAP_CD		= 1
	C_MAP_NM		= 2	

	
	 C_ACCT_CD      =  1
	 C_ACCT_POP		=  2
	 C_ACCT_NM		=  3
	 C_ACCT_GP_CD	=  4
	 C_ACCT_GP_POP	=  5
	 C_ACCT_GP_NM	=  6

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
    
    lgCurrGrid = TYPE_1
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

	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	
	' -- 1번 그리드 
	With lgvspdData(TYPE_1)
	
	ggoSpread.Source = lgvspdData(TYPE_1)	
   'patch version
    ggoSpread.Spreadinit "V20041222" & TYPE_1,,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_MAP_NM + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
			ggoSpread.SSSetEdit     C_MAP_CD, "코드", 4,,,100,1
			ggoSpread.SSSetEdit     C_MAP_NM, "맵핑계정명",   25,,,100,1
			

	.ReDraw = true
	
    Call SetSpreadLock(TYPE_1)
    
    End With

	' -- 2번 그리드 
	With lgvspdData(TYPE_2)
	
	ggoSpread.Source = lgvspdData(TYPE_2)	
   'patch version
    ggoSpread.Spreadinit "V20041222" & TYPE_2,,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_ACCT_GP_NM + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
		ggoSpread.SSSetEdit     C_ACCT_CD, "계정코드", 10,,,100,1
		ggoSpread.SSSetButton   C_ACCT_POP
		ggoSpread.SSSetEdit     C_ACCT_NM, "계정명",   25,,,100,1
		ggoSpread.SSSetEdit     C_ACCT_GP_CD, "대표계정", 10,,,100,1
		ggoSpread.SSSetButton   C_ACCT_GP_POP
		ggoSpread.SSSetEdit     C_ACCT_GP_NM, "대표계정명",  25,,,100,1


		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_ACCT_GP_CD, C_ACCT_GP_NM, True)
	
	'Call ggoSpread.SSSetColHidden(C_SEQ_NO, C_SEQ_NO, True)
	
	.ReDraw = true
	
    Call SetSpreadLock(TYPE_2)
    
    End With
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub


Sub SetSpreadLock(Byval pType)
    With lgvspdData(pType)

    .ReDraw = False
    
      ggoSpread.Source = lgvspdData(pType)
      ggoSpread.SpreadLockWithOddEvenRowColor()
	
    .ReDraw = True

    End With
End Sub




Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
dim sumRow
    ggoSpread.Source = lgvspdData(Type_2)
    With lgvspdData(Type_2)

     .ReDraw = False
       ggoSpread.SSSetRequired      C_ACCT_CD , pvStartRow, pvEndRow	 
	   ggoSpread.SSSetProtected     C_ACCT_NM , pvStartRow, pvEndRow
	   ggoSpread.SSSetProtected     C_ACCT_GP_NM , pvStartRow, pvEndRow

        
  .ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
 
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = lgvspdData(TYPE_1)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_W1			= iCurColumnPos(1)
            C_W1_NM			= iCurColumnPos(2)
            C_W_CHK			= iCurColumnPos(3)
            C_W2			= iCurColumnPos(4)
            C_UPDT_USER		= iCurColumnPos(5)
            C_UPDT_DT		= iCurColumnPos(6)
            C_W3			= iCurColumnPos(7)
            C_W4			= iCurColumnPos(8)
    End Select    
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

End Sub

'============================================  조회조건 함수  ====================================




'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1000000000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()
	Call MainQuery()
     
    
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

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '   
   lgKeyStream = lgKeyStream &  (frm1.txtMapcd.Value)   &  parent.gColSep ' 

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

'==========================================================================================
' -- 0번 그리드 
Sub vspdData0_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData0_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData0_GotFocus()
	lgCurrGrid = TYPE_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub


' -- 1번 그리드 
Sub vspdData1_Click(ByVal Col, ByVal Row)

	lgCurrGrid = TYPE_2

	
    Call SetPopupMenuItemInf("1111111111")    

    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData1

	'Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData1_Change(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_change(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_2
	Call vspdData_GotFocus(lgCurrGrid)
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


Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, datW1_DOWN, datW1, iRow, iMaxRows, dblW2, dblW4, dblW5,IntRetCd,strMajor
	
	lgBlnFlgChgValue= True ' 변경여부 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(lgvspdData(Index).text) < CDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row
   
	' --- 추가된 부분 
	With lgvspdData(Index)
	
		 Select Case Col
        Case C_ACCT_CD
			 lgvspdData(lgCurrGrid).Row = Row
			 lgvspdData(lgCurrGrid).Col = Col
            IntRetCd =  CommonQueryRs(" ACCT_NM "," TB_WORK_6 ","  ACCT_CD='" & Trim(.text) &"' and  CO_CD = '"&  Trim(frm1.txtCO_CD.value) &"' and FISC_YEAR  = '"& Trim(Frm1.txtFISC_YEAR.Text) &"' and REP_TYPE = '"& Trim(frm1.cboREP_TYPE.Value) &"' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
			If IntRetCd = false then
			   Call  DisplayMsgBox("110100","X",Trim(.text) ,"X")                                  '☜: Please do Display first.
        
			      .Col = Col 
				  .text = ""
                  .Col = Col + 2
				  .text = ""
	              
			Else
			      .Col = Col + 2
				  .text = Trim(Replace(lgF0,Chr(11),""))
	              
			End if 
		 Case C_ACCT_GP_CD
				 if  Trim(frm1.txtMapcd.value) = 10 then
					     strMajor = "W1056"										' Where Condition
				 elseif  Trim(frm1.txtMapcd.value) = "06" then
						   strMajor = "W1084"		    
				 elseif  Trim(frm1.txtMapcd.value) = "07" then
				 		 strMajor = "W1085"
				 elseif  Trim(frm1.txtMapcd.value) = "34" then
						  strMajor = "W1086"
				end if  	
				lgvspdData(lgCurrGrid).Row = Row
			    lgvspdData(lgCurrGrid).Col = Col  
            IntRetCd =  CommonQueryRs(" MINOR_NM "," B_MINOR ","  MINOR_CD='" & Trim(.text) &"' AND MAJOR_CD='" & strMajor &"'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
			If IntRetCd = false then
			   'Call  DisplayMsgBox("110100","X",Trim(.text) ,"X")                                  '☜: Please do Display first.
        
			      .Col = Col 
				  .text = ""
                  .Col = Col + 2
				  .text = ""
	              
			Else
			      .Col = Col + 2
				  .text = Trim(Replace(lgF0,Chr(11),""))
	              
			End if 	
        
    End Select
	

	End With
	
End Sub
Sub vspdData_Click(Index, ByVal Col, ByVal Row)
    dim IntRetCD
	lgCurrGrid = Index

   
   
   
    
   
    Set gActiveSpdSheet = lgvspdData(Index)
	If Index = TYPE_2 Then 
		
		Exit Sub
	End If
	
	
	
	 ggoSpread.Source = lgvspdData(TYPE_2)	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Sub
		End If
    End If
    
	
    If lgvspdData(TYPE_1).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(Index)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    Else
		' 디테일 그리드 호출 
		Call LayerShowHide(1)
	
		Dim strVal, sMAP_CD
    
		With lgvspdData(TYPE_1)
			.Row = Row	: .Col = C_MAP_CD : sMAP_CD = .Text
			frm1.txtMapcd.value = sMAP_CD
		End With
		
		 if sMAP_CD  = "10"   or sMAP_CD  = "06"  or sMAP_CD =  "07"    or sMAP_CD  =  "34" then  '접대비, 매출, 대손충당금설정대상, 재고자산 

       	      Call ggoSpread.SSSetColHidden(C_ACCT_GP_CD, C_ACCT_GP_NM, False)
       	   

        ELSE
              Call ggoSpread.SSSetColHidden(C_ACCT_GP_CD, C_ACCT_GP_NM, True)      
        	
        end if  	
     
		
		lgvspdData(TYPE_2).MaxRows = 0
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.ClearSpreadData
		
		With frm1
			strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001		
			strVal = strVal     & "&txtCocd="			 & .txtCO_CD.value      '☜: Query Key    				         
		    strVal = strVal     & "&txtFISC_YEAR="       & .txtFISC_YEAR.Text      '☜: Query Key        
		    strVal = strVal     & "&cboREP_TYPE="        & .cboREP_TYPE.Value      '☜: Query Key   
		    strVal = strVal     & "&txtCurrGrid="        & TYPE_2      
			strVal = strVal     & "&sMAP_CD="			 & sMAP_CD      
			
			Call RunMyBizASP(MyBizASP, strVal)   
		End With  
    End If

	
	lgvspdData(Index).Row = Row
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

		Case 2
			arrParam(0) = "계정코드"								' 팝업 명칭 
			arrParam(1) = "(select Distinct ACCT_CD , ACCT_NM FROM TB_WORK_6 "
			arrParam(1) = arrParam(1) & " where CO_CD = '"&  frm1.txtCO_CD.value &"' and FISC_YEAR  = '"& Trim(Frm1.txtFISC_YEAR.Text) &"' and REP_TYPE = '"& Trim(Frm1.cboREP_TYPE.Value) &"'  ) t" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""										' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 
            
			arrField(0) = "t.ACCT_CD"									' Field명(0)
			arrField(1) = "t.ACCT_NM"									' Field명(1)
			arrField(2) = ""									' Field명(1)
			arrField(3) = ""									' Field명(1)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정명"									' Header명(1)
			arrHeader(2) = ""									' Header명(1)
	
	  
		Case 3
		
		
			arrParam(0) = "대표계정코드"								' 팝업 명칭 
			arrParam(1) = "B_MINOR" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode) 										' Code Condition
			arrParam(3) = ""												' Name Cindition
			if  Trim(frm1.txtMapcd.value) = 10 then
				     arrParam(4) = "MAJOR_CD = 'W1056'"										' Where Condition
			elseif  Trim(frm1.txtMapcd.value) = "06" then
					 arrParam(4) = "MAJOR_CD = 'W1084'"	      
			elseif  Trim(frm1.txtMapcd.value) = "07" then
					 arrParam(4) = "MAJOR_CD = 'W1085'"	
			elseif  Trim(frm1.txtMapcd.value) = "34" then
					 arrParam(4) = "MAJOR_CD = 'W1086'"			
			end if  	
				arrParam(5) = "대표계정코드"									' 조건필드의 라벨 명칭 
            
			arrField(0) = "MINOR_CD"									' Field명(0)
			arrField(1) = "mINOR_NM"									' Field명(1)
			arrField(2) = ""									' Field명(1)
			arrField(3) = ""									' Field명(1)
			
			arrHeader(0) = "대표계정코드"									' Header명(0)
			arrHeader(1) = "대표계정명"									' Header명(1)
			arrHeader(2) = ""									' Header명(1)
	
		Case Else
			Exit Function
	End Select

	IsOpenPop = True
			
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
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
			Case 2
				.vspdData1.Col = C_ACCT_CD
				.vspdData1.Text = arrRet(0)
				.vspdData1.Col = C_ACCT_NM
				.vspdData1.Text = arrRet(1)
				
				Call vspdData1_Change(C_ACCT_CD, frm1.vspdData1.activerow )	 ' 변경이 읽어났다고 알려줌 
			Case 3
				.vspdData1.Col = C_ACCT_GP_CD
				.vspdData1.Text = arrRet(0)
				.vspdData1.Col = C_ACCT_GP_NM
				.vspdData1.Text = arrRet(1)
				
				Call vspdData1_Change(C_ACCT_GP_CD, frm1.vspdData1.activerow )	 ' 변경이 읽어났다고 알려줌 	
		
		End Select
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData1
        ggoSpread.Source = frm1.vspdData1
        
        If Row > 0 And Col = C_ACCT_POP Then
            .Col = Col - 1
            .Row = Row
            Call OpenPopup(.Text, 2)

        End If
        
         If Row > 0 And Col = C_ACCT_GP_POP Then
            .Col = Col - 1
            .Row = Row
            Call OpenPopup(.Text, 3)

        End If
        
    End With
End Sub

Sub vspdData_ColWidthChange(Index, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_GotFocus(Index)
    ggoSpread.Source = lgvspdData(Index)
    lgCurrGrid = Index
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

    Call SetToolbar("1000000000000111")

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
    ggoSpread.Source = lgvspdData(TYPE_2)
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
    Dim blnChange, dblSum,RetFlag
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = lgvspdData(TYPE_2)
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

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
 
End Function

Function FncCancel() 
                                                 '☜: Protect system from crashing
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

	With Frm1.vspdData1
	
		.focus
		ggoSpread.Source = Frm1.vspdData1

		iRow = .ActiveRow
		ggoSpread.InsertRow ,imRow
		SetSpreadColor iRow + 1, iRow + imRow
		
		

		
		 if frm1.txtMapcd.value = "10"or frm1.txtMapcd.value ="06" or frm1.txtMapcd.value = "07" or frm1.txtMapcd.value = "34" then  '접대비, 매출, 대손충당금설정대상, 재고자산 
       	     ggoSpread.SSSetRequired      C_ACCT_GP_CD , iRow + 1, iRow + 1	
         end if

		
    End With

	
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function
Function FncCancel() 
    ggoSpread.Source = lgvspdData(TYPE_2)
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncDeleteRow() 
    Dim lDelRows


		With lgvspdData(TYPE_2)
			.focus
			ggoSpread.Source = lgvspdData(TYPE_2)
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
	
    ggoSpread.Source = lgvspdData(TYPE_2)	
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
        strVal = strVal     & "&txtCurrGrid="        & TYPE_1      
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    
    

  
    
    
    
    If lgvspdData(TYPE_1).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		'Call SetGridSpan
		
		Call SetToolbar("1000000000000111")										<%'버튼 툴바 제어 %>

	End If
	lgvspdData(TYPE_1).focus
	
	Call vspdData0_Click(lgvspdData(TYPE_1).ActiveCol,lgvspdData(TYPE_1).ActiveRow)
End Function

Function DbQueryOk2()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgvspdData(TYPE_1).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		'Call SetGridSpan
		
		     Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

				'1 컨펌체크 
				If wgConfirmFlg = "Y" Then

				    Call SetToolbar("1000000000000111")	
					
				Else
				   '2 디비환경값 , 로드시환경값 비교 
					 Call SetToolbar("1000111100000111")										<%'버튼 툴바 제어 %>
	
	End If								<%'버튼 툴바 제어 %>

	End If
	lgvspdData(TYPE_1).focus			
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
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          

	if LayerShowHide(1) = false then
	exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With lgvspdData(TYPE_2)

       For lRow = 1 To .MaxRows
    
           .Row = lRow
           .Col = 0
        
           Select Case .Text
 
               Case  ggoSpread.InsertFlag                                      '☜: Update추가 
													  strVal = strVal & "C"  &  parent.gColSep
													  strVal = strVal & lRow &  parent.gColSep
                    .Col = C_ACCT_CD		: strVal = strVal & Trim(.Text) &  parent.gColSep
                    .Col = C_ACCT_NM		: strVal = strVal & Trim(.Text) &  parent.gColSep
                    .Col = C_ACCT_GP_CD		: strVal = strVal & Trim(.Text) &  parent.gRowSep



                     lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
													  strVal = strVal & "U"  &  parent.gColSep
													  strVal = strVal & lRow &  parent.gColSep
                    .Col = C_ACCT_CD		: strVal = strVal & Trim(.Text) &  parent.gColSep
                    .Col = C_ACCT__NM		: strVal = strVal & Trim(.Text) &  parent.gColSep
                    .Col = C_ACCT_GP_CD		: strVal = strVal & Trim(.Text) &  parent.gRowSep
           
                    lGrpCnt = lGrpCnt + 1
                                        
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D"  &  parent.gColSep
                                                  strDel = strDel & lRow &  parent.gColSep
                   .Col = C_ACCT_CD    : strDel = strDel & Trim(.Text) &  parent.gRowSep	'삭제시 key만								
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

       frm1.txtMode.value        =  parent.UID_M0002
       frm1.txtKeyStream.value   =  lgKeyStream
	   frm1.txtCurrGrid.value    = TYPE_2
	   frm1.txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
    DbSave = True                                                      
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	
    'Call MainQuery()
  	ggoSpread.Source = lgvspdData(TYPE_2)
	ggoSpread.ClearSpreadData
   Call vspdData0_Click(lgvspdData(TYPE_1).ActiveCol,lgvspdData(TYPE_1).ActiveRow)
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=300>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/wb112ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
							<TR HEIGHT="100%">
								<TD WIDTH=40%>
									<script language =javascript src='./js/wb112ma1_vaSpread1_vspdData0.js'></script>
								</TD>
								<TD WIDTH=60%>
									<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT=100%>
										<script language =javascript src='./js/wb112ma1_vaSpread2_vspdData1.js'></script>
										</TD>
									</TR>
									
								</TD>
							</TR>
						</TABLE>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMapcd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

