<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 표준원가관리 
'*  3. Program ID           : c2716ma1
'*  4. Program Name         : 품목별 가공비 정보 등록 
'*  5. Program Desc         : 품목별 인건비/제조경비 금액 및 원가요소 정보를 설정한다.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/08/24
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      : Cho Ig Sung
'* 11. Comment              :
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>

<Script Language="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "c2716mb1.asp"

'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
Dim C_ItemCd 
Dim C_ItemPop
Dim C_ItemNm 
Dim C_LaborCost
Dim C_LaborCostElmtCd 
Dim C_LaborCostElmtPop
Dim C_LaborCostElmtNm 
Dim C_Expense 
Dim C_ExpenseCostElmtCd
Dim C_ExpenseCostElmtPop 
Dim C_ExpenseCostElmtNm 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim lgQueryFlag					' 신규조회 및 추가조회 구분 Flag

 
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================


'========================================================================================================
Sub initSpreadPosVariables()  

 C_ItemCd				= 1
 C_ItemPop				= 2
 C_ItemNm				= 3
 C_LaborCost			= 4
 C_LaborCostElmtCd		= 5													'☆: Spread Sheet의 Column별 상수 
 C_LaborCostElmtPop		= 6
 C_LaborCostElmtNm		= 7
 C_Expense				= 8	
 C_ExpenseCostElmtCd	= 9
 C_ExpenseCostElmtPop	= 10												'☆: Spread Sheet의 Column별 상수  
 C_ExpenseCostElmtNm	= 11

End Sub


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""			                'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgSortKey = 1
	    
End Sub


'======================================================================================================== 
Sub SetDefaultVal()
End Sub


'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub



'=========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    

	With frm1.vspdData
	
    .MaxCols = C_ExpenseCostElmtNm+1									'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols											         	'☆: 사용자 별 Hidden Column
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021123",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()

	.ReDraw = false
    
    Call GetSpreadColumnPos("A")
    
	'SSSetEdit(Col, Header, ColWidth , HAlign , Row , Length)
	ggoSpread.SSSetEdit C_ItemCd, "품목코드", 12,,,18,2
	ggoSpread.SSSetButton C_ItemPop    
	ggoSpread.SSSetEdit C_ItemNm, "품목명", 20    
	ggoSpread.SSSetFloat C_LaborCost, "노무비", 15, Parent.ggUnitCostNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit C_LaborCostElmtCd, "노무비원가요소코드", 20,,,6,2
	ggoSpread.SSSetButton C_LaborCostElmtPop
	ggoSpread.SSSetEdit C_LaborCostElmtNm, "노무비원가요소명",20
	ggoSpread.SSSetFloat C_Expense, "제조경비", 15, Parent.ggUnitCostNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit C_ExpenseCostElmtCd, "경비원가요소코드", 20,,,6,2
	ggoSpread.SSSetButton C_ExpenseCostElmtPop
	ggoSpread.SSSetEdit C_ExpenseCostElmtNm, "경비원가요소명", 20
	
	call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemPop)
	call ggoSpread.MakePairsColumn(C_LaborCostElmtCd,C_LaborCostElmtPop)
	call ggoSpread.MakePairsColumn(C_ExpenseCostElmtCd,C_ExpenseCostElmtPop)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub


'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_ItemCd, -1, C_ItemCd    
	ggoSpread.SpreadLock C_ItemPop, -1, C_ItemPop
	ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
	ggoSpread.SpreadLock C_LaborCostElmtNm, -1, C_LaborCostElmtNm
	ggoSpread.SpreadLock C_ExpenseCostElmtNm, -1, C_ExpenseCostElmtNm
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_ItemCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected	C_LaborCostElmtNm, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected	C_ExpenseCostElmtNm, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd				= iCurColumnPos(1)
			C_ItemPop				= iCurColumnPos(2)
			C_ItemNm				= iCurColumnPos(3)    
			C_LaborCost				= iCurColumnPos(4)
			C_LaborCostElmtCd		= iCurColumnPos(5)
			C_LaborCostElmtPop		= iCurColumnPos(6)
			C_LaborCostElmtNm		= iCurColumnPos(7)
			C_Expense				= iCurColumnPos(8)
			C_ExpenseCostElmtCd		= iCurColumnPos(9)
			C_ExpenseCostElmtPop    = iCurColumnPos(10)
			C_ExpenseCostElmtNm		= iCurColumnPos(11)
			
    End Select    
End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'=======================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				'팝업 명칭 
	arrParam(1) = "B_PLANT"						'TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	'Code Condition
	arrParam(3) = ""							'Name Cindition
	arrParam(4) = ""							'Where Condition
	arrParam(5) = "공장"					'TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					'Field명(0)
    arrField(1) = "PLANT_NM"					'Field명(1)
    
    arrHeader(0) = "공장코드"					'Header명(0)
    arrHeader(1) = "공장명"					'Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
		
End Function

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
		Case 0
			If Trim(frm1.txtPlantCd.Value) = "" Then
				IsOpenPop = False
				IntRetCD = DisplayMsgBox("125000","x","x","x") '공장을 먼저 입력하세요 
				frm1.txtPlantCd.focus
				Exit Function
			End If

			'arrParam(0) = "품목팝업"	  				' 팝업 명칭 
			'arrParam(1) = "B_ITEM a, B_PLANT b, B_ITEM_BY_PLANT c"						' TABLE 명칭 
			'arrParam(2) = strCode						' Code Condition
			'arrParam(3) = ""							' Name Cindition
			'arrParam(4) = "a.ITEM_CD = c.ITEM_CD AND b.PLANT_CD = c.PLANT_CD AND a.PHANTOM_FLG <> 'Y' AND c.PROCUR_TYPE <> 'O' AND b.PLANT_CD = '" & Trim(frm1.txtPlantCd.Value) & "'"	' Where Condition
			'arrParam(5) = "품목"    			' 조건필드의 라벨 명칭 

			'arrField(0) = "a.Item_Cd"						' Field명(0)
			'arrField(1) = "a.Item_Nm"						' Field명(1)
  
			'arrHeader(0) = "품목코드"	  				' Header명(0)
			'arrHeader(1) = "품목명"						' Header명(1)
	
			arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
			arrParam(1) = strCode							' Item Code
			arrParam(2) = "12!MM"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
			arrParam(3) = ""							' Default Value
	

			arrField(0) = 1 								' Field명(0) :"ITEM_CD"
			arrField(1) = 2 								' Field명(1) :"ITEM_NM"
    			
		Case 1
			arrParam(0) = "원가요소팝업"	  				' 팝업 명칭 
			arrParam(1) = "C_Cost_Elmt"					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "Cost_Elmt_type = " & FilterVar("L", "''", "S") & " "	' Where Condition
			arrParam(5) = "원가요소"    				' 조건필드의 라벨 명칭 

			arrField(0) = "Cost_Elmt_Cd"				' Field명(0)
			arrField(1) = "Cost_Elmt_Nm"				' Field명(1)
  
			arrHeader(0) = "원가요소코드"  				' Header명(0)
			arrHeader(1) = "원가요소명"					' Header명(1)
    
		Case 2
			arrParam(0) = "원가요소팝업"	  				' 팝업 명칭 
			arrParam(1) = "C_Cost_Elmt"					' TABLE 명칭 
			arrParam(2) = strCode					' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "Cost_Elmt_type = " & FilterVar("E", "''", "S") & " "		' Where Condition
			arrParam(5) = "원가요소"    				' 조건필드의 라벨 명칭 

			arrField(0) = "Cost_Elmt_Cd"				' Field명(0)
			arrField(1) = "Cost_Elmt_Nm"				' Field명(1)
  
			arrHeader(0) = "원가요소코드"  				' Header명(0)
			arrHeader(1) = "원가요소명"					' Header명(1)
	End Select
		
	if iWhere = 0 then
		arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent, arrParam, arrField), _
				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	end if
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'==========================================  2.4.3 SetPopup()  =============================================
'	Name : SetPopup()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.vspdData.Col = C_ItemCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_ItemNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)

			Case 1
				.vspdData.Col = C_LaborCostElmtCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_LaborCostElmtNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
				
			Case 2
				.vspdData.Col = C_ExpenseCostElmtCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_ExpenseCostElmtNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			
		End Select

		lgBlnFlgChgValue = True
	End With
	
End Function
'=======================================================================================================
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function

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
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
	'Call InitCombo
    Call SetDefaultVal
    Call SetToolbar("110011010010111")										'⊙: 버튼 툴바 제어 
    frm1.txtPlantCd.focus
   	Set gActiveElement = document.activeElement		
     
End Sub

'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '**************************  3.2 HTML Form Element & Object Event처리  *******************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'******************************************************************************************

 '******************************  3.2.1 Object Tag 처리  **********************************
'	Window에 발생 하는 모든 Even 처리	
'******************************************************************************************

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	


    gMouseClickStatus = "SPC"	'Split 상태코드 

    Set gActiveSpdSheet = frm1.vspdData
    
    
    If Row <= 0 Then
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


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
	
End Sub


'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================

Sub vspdData_Change(ByVal Col, ByVal Row)

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
	
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	

 
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_ItemPop
				.vspdData.Col = Col
				.vspdData.Row = Row
				
				.vspdData.Col = 1
				Call OpenPopup(.vspdData.Text, 0)

			Case C_LaborCostElmtPop
				.vspdData.Col = Col
				.vspdData.Row = Row
				
				.vspdData.Col = 5
				Call OpenPopup(.vspdData.Text, 1)

			Case C_ExpenseCostElmtPop        
				.vspdData.Col = Col
				.vspdData.Row = Row
				  
				.vspdData.Col = 9
				Call OpenPopup(.vspdData.Text, 2)
		End Select
       Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub



'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
    	If lgStrPrevKey <> "" Then                  '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
	      	DbQuery
    	End If

    End if
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
'	Call InitSpreadSheet
    'Call InitComboBox
    Call InitVariables 															'⊙: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    if frm1.txtPlantCd.value = "" then
		frm1.txtPlantNm.value = ""
    end if
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    IF DbQuery = False Then																'☜: Query db data
		Exit function
	END IF
	 
    FncQuery = True																'⊙: Processing is OK
    
End Function


'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") '☜ 바뀐부분    
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    
    FncNew = True                                                           '⊙: Processing is OK

End Function


'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    If Not chkField(Document, "1") Then               '⊙: Check required field(Single area)
       Exit Function
    End If
    
  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
     
    '-----------------------
    'Save function call area
    '-----------------------
    IF DbSave = False Then				                                                  '☜: Save db data
		Exit Function
	END If
	
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

'=======================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
	if frm1.vspdData.maxrows < 1 then exit function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
    
     
    frm1.vspdData.Col = C_ItemCd
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_ItemNm
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function


'========================================================================================

Function FncCancel() 

	if frm1.vspdData.maxrows < 1 then exit function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function


'=======================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 

    Dim IntRetCD
    Dim imRow
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
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement  
   
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
	End If 

End Function


'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    
    if frm1.vspdData.maxrows < 1 then exit function
    
    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
    
    Set gActiveElement = document.ActiveElement   
    
End Function


'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function


'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'=======================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 '☜: 화면 유형 
End Function


'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'=======================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub


'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 


'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
    Err.Clear                                                               '☜: Protect system from crashing
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'☜:조회표시 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'☜:조회표시 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    
    DbQuery = True

End Function


'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

    Call SetToolbar("110011110011111")										'⊙: 버튼 툴바 제어 
    
   	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   

End Function



'========================================================================================

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep
    Dim iRowSep     
	
    DbSave = False                                                          '⊙: Processing is NG
    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
    On Error Resume Next                                                   '☜: Protect system from crashing
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	

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
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = 0
        
			Select Case .vspdData.Text

	            Case ggoSpread.InsertFlag											'☜: 신규 
					strVal = strVal & "C" & iColSep & lRow & iColSep				'☜: U=Create
					.vspdData.Col = C_ItemCd		'1
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_LaborCost	'3
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
					.vspdData.Col = C_LaborCostElmtCd	'5
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_Expense	'7
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
					.vspdData.Col = C_ExpenseCostElmtCd	'9
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag											'☜: 수정 
					strVal = strVal & "U" & iColSep & lRow & iColSep				'☜: U=Update
					.vspdData.Col = C_ItemCd		'1
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_LaborCost	'3
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
					.vspdData.Col = C_LaborCostElmtCd	'5
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_Expense	'7
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
					.vspdData.Col = C_ExpenseCostElmtCd	'9
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag											'☜: 삭제 
					strDel = strDel & "D" & iColSep & lRow & iColSep				'☜: D=Delete
					.vspdData.Col = C_ItemCd										'1
					strDel = strDel & Trim(.vspdData.Text) & iRowSep
					lGrpCnt = lGrpCnt + 1
                
	        End Select
                
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)									'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

	Call InitVariables
	frm1.vspdData.maxrows = 0
	Call MainQuery()
		
End Function



'========================================================================================

Function DbDelete() 
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별가공비정보등록</font></td>
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
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD656"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
										 <INPUT TYPE=TEXT ID="txtPlantNm" NAME="txtPlantNm" SIZE=30 tag="14X">
									</TD>
									<!-- <TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>  -->
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

