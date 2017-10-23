<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2101ma1
'*  4. Program Name         : 예산계정등록 
'*  5. Program Desc         : Register of Budget Account/Accout Group
'*  6. Comproxy List        : FU0011, FU0018
'*  7. Modified date(First) : 2000.09.14
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "f2101mb1.asp"			'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns


Dim C_CTRL_FG
Dim C_CTRL_FG_NM
Dim C_BDG_CD
Dim C_BDG_NM
Dim C_ACCT_CD
Dim C_ACCT_PB
Dim C_ACCT_NM
Dim C_GP_CD
Dim C_GP_PB
Dim C_GP_NM
Dim C_CTRL_UNIT
Dim C_CTRL_UNIT_NM
Dim C_TRANS_FG
Dim C_DIVERT_FG
Dim C_ADD_FG

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 


 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================


 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 

Sub initSpreadPosVariables()
	C_CTRL_FG = 1
	C_CTRL_FG_NM = 2
	C_BDG_CD = 3
	C_BDG_NM = 4
	C_ACCT_CD = 5
	C_ACCT_PB = 6
	C_ACCT_NM = 7
	C_GP_CD = 8
	C_GP_PB = 9 
	C_GP_NM = 10
	C_CTRL_UNIT = 11
	C_CTRL_UNIT_NM = 12
	C_TRANS_FG =  13
	C_DIVERT_FG = 14
	C_ADD_FG = 15
End Sub

Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False                    'Indicates that no value changed
	lgIntGrpCount = 0                           'initializes Group View Size
    
	lgStrPrevKey = ""                           'initializes Previous Key
	lgLngCurRows = 0                            'initializes Deleted Rows Count
    
	lgSortKey = 1
	lgPageNo = 0
End Sub

 '******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	' 현재 Page의 Form Element들을 Clear한다. 
	Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()
   
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
	
	With frm1.vspdData
		.ReDraw = False
	
		.MaxCols = C_ADD_FG + 1
		.ColsFrozen = C_BDG_CD
		.MaxRows = 0

        Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCombo  C_CTRL_FG     , "구분"    , 08
		ggoSpread.SSSetCombo  C_CTRL_FG_NM  , "구분"    , 12
		ggoSpread.SSSetEdit   C_BDG_CD      , "예산코드", 15, , , 18, 2
		ggoSpread.SSSetEdit   C_BDG_NM      , "예산명"  , 20, , , 30
		ggoSpread.SSSetEdit   C_ACCT_CD     , "계정코드", 15, , , 20, 2
		ggoSpread.SSSetButton C_ACCT_PB
		ggoSpread.SSSetEdit   C_ACCT_NM     , "계정명"  , 20, , , 30
		ggoSpread.SSSetEdit   C_GP_CD       , "그룹코드", 15, , , 20, 2
		ggoSpread.SSSetButton C_GP_PB
		ggoSpread.SSSetEdit   C_GP_NM       , "그룹명"  , 20, , , 30
		ggoSpread.SSSetCombo  C_CTRL_UNIT   , "구분"    , 08
		ggoSpread.SSSetCombo  C_CTRL_UNIT_NM, "예산통제단위", 15    
		
		ggoSpread.SSSetCheck  C_TRANS_FG    , "전용여부", 15, , True
		ggoSpread.SSSetCheck  C_DIVERT_FG   , "이월여부", 15, , True
		ggoSpread.SSSetCheck  C_ADD_FG      , "추가여부", 15, , True

		Call ggoSpread.MakePairsColumn(C_ACCT_CD,C_ACCT_PB)
		Call ggoSpread.MakePairsColumn(C_GP_CD,C_GP_PB)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_CTRL_FG,C_CTRL_FG,True)
		Call ggoSpread.SSSetColHidden(C_CTRL_UNIT,C_CTRL_UNIT,True)

		.ReDraw = True

		Call SetSpreadLock(.ActiveRow, "Query")
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal Row, ByVal FGubun)
    With frm1.vspdData
		.ReDraw = False
 
		Select Case FGubun
			Case "Query"
				ggoSpread.SpreadLock C_BDG_CD    , -1, C_BDG_CD
				ggoSpread.SpreadLock C_CTRL_FG_NM, -1, C_CTRL_FG_NM
				ggoSpread.SpreadLock C_ACCT_CD   , -1, C_ACCT_CD
				ggoSpread.SpreadLock C_ACCT_PB   , -1, C_ACCT_PB
				ggoSpread.SpreadLock C_ACCT_NM   , -1, C_ACCT_NM
				ggoSpread.SpreadLock C_GP_CD     , -1, C_GP_CD
				ggoSpread.SpreadLock C_GP_PB     , -1, C_GP_PB
				ggoSpread.SpreadLock C_GP_NM     , -1, C_Gp_NM
			Case "Insert"		
				ggoSpread.SpreadUnLock C_BDG_CD    , Row, C_BDG_CD    , Row
				ggoSpread.SpreadUnLock C_CTRL_FG_NM, Row, C_CTRL_FG_NM, Row
				ggoSpread.SpreadUnLock C_ACCT_CD   , Row, C_ACCT_CD   , Row
				ggoSpread.SpreadUnLock C_ACCT_PB   , Row, C_ACCT_PB   , Row
				ggoSpread.SpreadLock   C_ACCT_NM   , Row, C_ACCT_NM   , Row
				ggoSpread.SpreadUnLock C_GP_CD     , Row, C_GP_CD     , Row
				ggoSpread.SpreadUnLock C_GP_PB     , Row, C_GP_PB     , Row
				ggoSpread.SpreadLock   C_GP_NM     , Row, C_Gp_NM     , Row
			Case "Acct"
				ggoSpread.SpreadUnLock C_ACCT_CD   , Row, C_ACCT_CD   , Row
				ggoSpread.SpreadUnLock C_ACCT_PB   , Row, C_ACCT_PB   , Row
				ggoSpread.SpreadLock   C_GP_CD     , Row, C_GP_CD     , Row
				ggoSpread.SpreadLock   C_GP_PB     , Row, C_GP_PB     , Row
			Case "Group"
				ggoSpread.SpreadLock   C_ACCT_CD   , Row, C_ACCT_CD   , Row
				ggoSpread.SpreadLock   C_ACCT_PB   , Row, C_ACCT_PB   , Row
				ggoSpread.SpreadUnLock C_GP_CD     , Row, C_GP_CD     , Row
				ggoSpread.SpreadUnLock C_GP_PB     , Row, C_GP_PB     , Row
		End Select
		
		ggoSpread.SSSetProtected	.MaxCols,-1,-1
		.ReDraw = True
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow, ByVal FGubun)
    With frm1
		ggoSpread.Source = frm1.vspdData
		.vspdData.ReDraw = False

		' 필수 입력 항목으로 설정 
		Select Case FGubun
			Case "Query"
				ggoSpread.SSSetRequired  C_CTRL_UNIT_NM, pvStartRow, pvEndRow
			Case "Insert"
				ggoSpread.SSSetRequired  C_CTRL_FG_NM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired  C_CTRL_UNIT_NM, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_ACCT_NM, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_GP_NM, pvStartRow, pvEndRow
			Case "Acct"
				ggoSpread.SSSetRequired  C_ACCT_CD, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_GP_CD, pvStartRow, pvEndRow
			Case "Group"
				ggoSpread.SSSetProtected C_ACCT_CD, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired  C_GP_CD, pvStartRow, pvEndRow
		End Select
		
		.vspdData.ReDraw = True
    End With
End Sub

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet,arrTempImport,i,pQuery
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
   		Case 0
			arrParam(0) = frm1.txtBDG_CD.Alt					' 팝업 명칭 
			arrParam(1) = "F_BDG_ACCT"    						' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = frm1.txtBDG_CD.Alt					' 조건필드의 라벨 명칭 

			arrField(0) = "BDG_CD"	     						' Field명(0)
			arrField(1) = "GP_ACCT_NM"			    				' Field명(1)
    
			arrHeader(0) = "예산코드"						' Header명(0)
			arrHeader(1) = "예산명"							' Header명(1)
		Case 1
'			Call CommonQueryRs(" GP_CD "," F_BDG_ACCT "," ACCT_CTRL_FG = 'G' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	         

			arrParam(0) = "계정코드 팝업"					' 팝업 명칭 
			arrParam(1) = "A_ACCT A"    							' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "A.BDG_CTRL_FG = " & FilterVar("Y", "''", "S") & "  AND A.BDG_CTRL_GP_LVL = 0"	' Where Condition

			If lgF0 <> "" Then  
				arrTempImport = Split(lgF0, chr(11))
	        
 				For i = 0 To UBound(arrTempImport, 1) - 1 
					arrParam(4) = arrParam(4) & " AND A.GP_CD <> "&" " & FilterVar(arrTempImport(i), "''", "S") & ""
				next
			End If
		
			arrParam(5) = "계정"						' 조건필드의 라벨 명칭 

			arrField(0) = "A.ACCT_CD"	     						' Field명(0)
			arrField(1) = "A.ACCT_NM"			    				' Field명(1)
    
			arrHeader(0) = "계정코드"						' Header명(0)
			arrHeader(1) = "계정명"						' Header명(1)
		Case 2
			arrParam(0) = "계정그룹 팝업"					' 팝업 명칭 
			arrParam(1) = "A_ACCT_GP B"	    					' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "B.BDG_CTRL_FG = " & FilterVar("Y", "''", "S") & " "									' Where Condition
			arrParam(5) = "계정그룹코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "B.GP_CD"	    						' Field명(0)
			arrField(1) = "B.GP_NM"			    				' Field명(1)
			
			arrHeader(0) = "계정그룹코드"					' Header명(0)
			arrHeader(1) = "계정그롭코드명"					' Header명(0)
	End Select	

	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		ggoSpread.Source = frm1.vspdData
		Exit Function
	Else
		With frm1
			Select Case iWhere
				Case 0
				    .txtBDG_CD.value = arrRet(0)
					.txtBDG_NM.value = arrRet(1)
					.txtBDG_CD.focus
				Case 1
					.vspdData.Col  = C_ACCT_CD
					.vspdData.Text = arrRet(0)
					.vspdData.Col  = C_ACCT_NM
					.vspdData.Text = arrRet(1)
					Call vspdData_Change(.vspdData.Col, .vspdData.Row )	
				Case 2
					.vspdData.Col  = C_GP_CD
					.vspdData.Text = arrRet(0)
					.vspdData.Col  = C_GP_NM
					.vspdData.Text = arrRet(1)
					Call vspdData_Change(.vspdData.Col, .vspdData.Row )	
			End Select
		End With
	End If	
End Function

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2000", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboFG ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Spread()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2000", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_CTRL_FG			'COLM_DATA_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_CTRL_FG_NM
        	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2010", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_CTRL_UNIT			'KEY_DATA_TYPE_1
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_CTRL_UNIT_NM
End Sub

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
	 	 	.Row = intRow
			
			.Col = C_CTRL_FG
			intIndex = .value
			.col = C_CTRL_FG_NM
			.value = intindex
			
			.Col = C_CTRL_UNIT
			intIndex = .value
			.col = C_CTRL_UNIT_NM
			.value = intindex
		Next	
	End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
 '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format'    
	Call ggoOper.LockField(Document, "Q")
    Call InitSpreadSheet                          '⊙: Setup the Spread Sheet
    Call InitVariables                            '⊙: Initializes local global Variables
    Call InitComboBox
    Call InitComboBox_Spread
    Call SetDefaultVal
	Call SetToolbar("1100110100001111")
    frm1.txtBDG_CD.focus 
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_CTRL_FG = iCurColumnPos(1)
            C_CTRL_FG_NM = iCurColumnPos(2)
            C_BDG_CD = iCurColumnPos(3)
            C_BDG_NM = iCurColumnPos(4)
            C_ACCT_CD = iCurColumnPos(5)
            C_ACCT_PB = iCurColumnPos(6)
            C_ACCT_NM = iCurColumnPos(7)
            C_GP_CD = iCurColumnPos(8)
            C_GP_PB = iCurColumnPos(9)
            C_GP_NM = iCurColumnPos(10)
            C_CTRL_UNIT = iCurColumnPos(11)
            C_CTRL_UNIT_NM = iCurColumnPos(12)
            C_TRANS_FG = iCurColumnPos(13)
            C_DIVERT_FG = iCurColumnPos(14)
            C_ADD_FG = iCurColumnPos(15)
    End Select    
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If	

	gMouseClickStatus = "SPC"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData
    
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData	
		.Row = Row
   
		Select Case Col
			Case  C_CTRL_FG_NM
				.Col     = Col
				intIndex = .Value
				.Col     = C_CTRL_FG
				.Value   = intIndex
		
				Select Case .Text
					Case "A"	 '계정선택 
						.Col  = C_GP_CD
						.Text = ""
						.Col  = C_GP_NM
						.Text = ""
						Call SetSpreadLock(Row, "Acct")
						Call SetSpreadColor(Row, Row, "Acct")
					Case "G"	 '계정그룹선택 
						.Col  = C_ACCT_CD
						.Text = ""
						.Col  = C_ACCT_NM
						.Text = ""
						Call SetSpreadLock(Row, "Group")
						Call SetSpreadColor(Row, Row, "Group")
				End Select
			Case C_CTRL_UNIT_NM
				.Col     = Col
				intIndex = .Value
				.Col     = C_CTRL_UNIT
				.Value   = intIndex
		End Select
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_BDG_CD Or NewCol <= C_BDG_CD Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgPageNo <> "" Then                         
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
End Sub

'==========================================================================================
' Event Name : vspdData_ButtonClicked
' Event Desc : 버튼 컬럼을 클릭할 경우 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_ACCT_PB Then
			.Col = C_ACCT_CD
			.Row = Row
			
			Call OpenPopup(.Text, 1)		
	    Elseif Row > 0 and Col = C_GP_PB Then
	        .Col = Col
			.Row = Row
			
			Call OpenPopup(.Text, 2)
		End If
	End With
End Sub

'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 

'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
    
    FncQuery = False          '⊙: Processing is NG
    Err.Clear                 '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		if IntRetCD = vbNo Then
			Exit Function
		End If
    End If
   
    '-----------------------
    'Erase contents area
    '-----------------------
	'Call InitSpreadSheet                          '⊙: Setup the Spread Sheet
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables							  '⊙: Initializes local global variables
    Call InitComboBox_Spread
    
 	If Not ChkField(Document, "1") Then	'⊙: This function check indispensable field
		Exit Function
    End If
    
    Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
    FncNew = False                  '⊙: Processing is NG
    Err.Clear                       '☜: Protect system from crashing
    'On Error Resume Next            '☜: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ' 변경된 내용이 있는지 확인한다.
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015",Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")     '⊙: Clear Condition Field	
    Call InitVariables                         '⊙: Initializes local global variables
    Call SetDefaultVal
    
    Call FncSetToolBar("New")
    FncNew = True                              '⊙: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
    
    FncDelete = False            '⊙: Processing is NG
    Err.Clear                    '☜: Protect system from crashing
    'On Error Resume Next        '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ' Update 상태인지를 확인한다.
    If lgIntFlgMode <> Parent.OPMD_UMODE Then        'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then											  '☜: Delete db data
		Exit Function                        
    End If
    
    '-----------------------
    'Erase condition area
    '-----------------------
	Call ggoOper.ClearField(Document, "1")								  '⊙: Clear Condition Field
    FncDelete = True													 '⊙: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim intRow
    
    FncSave = False            '⊙: Processing is NG
    Err.Clear                  '☜: Protect system from crashing
    'On Error Resume Next       '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then              '⊙: Check required field(Multi area)
		Exit Function
    End If

	'-----------------------------------------------------------
	'통제구분을 계정(A)으로 선택한 경우, 계정 필수 입력 체크 
	'통제구분을 그룹(G)으로 선택한 경우, 계정그룹 필수 입력 체크 
	'-----------------------------------------------------------	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
			
			Select Case ggoSpread.SSCheckFlag(intRow)
				Case 1, 2	'Insert/Update
					.Col = C_CTRL_FG_NM
					If Trim(.Text) = "" Then
						.Row = 0
						Call DisplayMsgBox("970021", "X", .Text, "X")
						.Row = intRow
						.Action = 0
						Exit Function
					End If
					
					.Col = C_CTRL_FG
					If Trim(.Text) = "A" Then
						.Col = C_ACCT_CD
						If Trim(.Text) = "" Then
							.Row = 0
							Call DisplayMsgBox("970021", "X", .Text,"X")
							.Row = intRow
							.Action = 0		'Parent.SS_ACTION_ACTIVE_CELL
							Exit Function
						End If
					ElseIf Trim(.Text) = "G" Then
						.Col = C_GP_CD
						If Trim(.Text) = "" Then
							.Row = 0
							Call DisplayMsgBox("970021", "X", .Text,"X")
							.Row = intRow
							.Action = 0		'Parent.SS_ACTION_ACTIVE_CELL
							Exit Function
						End If
					End If
					
					.Col = C_CTRL_UNIT_NM
					If Trim(.Text) = "" Then
						.Row = 0
						Call DisplayMsgBox("970021", "X", .Text,"X")
						.Row = intRow
						.Action = 0
						Exit Function
					End If
				Case Else
			End Select
		Next
	End With
	
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                  '☜: Save db data

	FncSave = True                                                           '⊙: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	 
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    
    Call SetSpreadLock(frm1.vspdData.ActiveRow, "Insert")
	Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow, "Insert")

    With frm1
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Col = C_CTRL_FG
		
		Select Case .vspdData.Text 
			Case "A"	'계정코드 
			    Call SetSpreadLock(frm1.vspdData.ActiveRow, "Acct")
				Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow, "Acct")
				
			Case "G"	'계정그룹 
			    Call SetSpreadLock(frm1.vspdData.ActiveRow, "Group")
				Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow, "Group")
		End Select
    End With
    
	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	Call InitData
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(Byval pvRowcnt) 
    Dim IntRetCD
    Dim imRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear   

    FncInsertRow = False                                                         '☜: Processing is NG
    If IsNumeric(Trim(pvRowcnt)) Then 
		imRow  = Cint(pvRowcnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If                              
   
	With frm1
		.vspdData.focus
		.vspdData.ReDraw = False

		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow,imRow
		
		Call SetSpreadLock(.vspdData.ActiveRow, "Insert")
		Call SetSpreadColor(.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1, "Insert")
		
		.vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	    
    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
    End With
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                                     '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                 '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox_Spread
   	Call ggoSpread.ReOrderingSpreadData()
   	Call SetSpreadColor(-1, -1, "Query")
   	Call InitData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal

	Call LayerShowHide(1)
 
    DbQuery = False
    Err.Clear                '☜: Protect system from crashing
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then			
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal		& "&txtBDG_CD=" & Trim(.htxtBDG_CD.value)	'조회 조건 데이타 
			strVal = strVal		& "&cbofg=" & Trim(.hcbofg.value)	'조회 조건 데이타 
			strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey
		Else			
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal		& "&txtBDG_CD=" & Trim(.txtBDG_CD.value)	'조회 조건 데이타 
			strVal = strVal		& "&cbofg=" & Trim(.cbofg.value)	'조회 조건 데이타 
			strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey
		End If
		strVal = strVal		& "&lgPageNo		=" & lgPageNo
		strVal = strVal		& "&txtMaxRows		=" & .vspdData.MaxRows

	    Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	Call InitData
	Call SetSpreadLock(-1, "Query")
	Call SetSpreadColor(-1, -1, "Query")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE	'⊙: Indicates that current mode is Update mode    
	    
    Call FncSetToolBar("Query")    
    Set gActiveElement = document.activeElement 
End Function

'========================================================================================
' Function Name : DbSave()
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal,strDel
	Dim iColSep 
	
	Call LayerShowHide(1)
	
    DbSave = False				'⊙: Processing is NG
    'On Error Resume Next		'☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""
		iColSep = Parent.gColSep
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
			
		    Select Case .vspdData.Text
  				Case ggoSpread.InsertFlag	   											'☜: 신규 
					strVal = strVal & "C" & iColSep & lRow & iColSep					'☜: U=Create
				    .vspdData.Col = C_CTRL_FG
				    strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_NM
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_ACCT_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_GP_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_CTRL_UNIT
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_TRANS_FG
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_DIVERT_FG
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_ADD_FG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag												'☜: 수정 
					strVal = strVal & "U" & iColSep & lRow & iColSep					'☜: U=Update
				    .vspdData.Col = C_CTRL_FG
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_BDG_NM
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_ACCT_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_GP_CD
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_CTRL_UNIT
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_TRANS_FG
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_DIVERT_FG
		            strVal = strVal & Trim(.vspdData.Text) & iColSep
		            .vspdData.Col = C_ADD_FG
		            
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag												'☜: 삭제 
					strDel = strDel & "D" & iColSep & lRow & iColSep					'☜: U=Delete
		            .vspdData.Col = C_BDG_CD
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		 Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'☜: 비지니스 ASP 를 가동 
	End With

    DbSave = True                           '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
    Call InitVariables
	ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData
	
    Call MainQuery()
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	On Error Resume Next
End Function

Sub txtBDG_CD_onChange()
	frm1.txtBDG_NM.value = ""
End Sub

'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
		Case "NEW"
			Call SetToolbar("1100110100001111")
		Case "QUERY"
			Call SetToolbar("1100111100111111")
	End Select
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- 
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  
-->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
						<FIELDSET CLASS="CLSFLD" >
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>예산</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBDG_CD" MAXLENGTH="18" SIZE=10  ALT="예산코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtBDG_CD.Value,0)">&nbsp;<INPUT NAME="txtBDG_NM" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" ALT="예산명" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>구분</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboFG" ALT="구분" STYLE="WIDTH: 100px" tag="11" ONCELLCHANGE ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" id=vaSpread1 tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="htxtBDG_CD" tag="24">
<INPUT TYPE=hidden NAME="hcbofg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
