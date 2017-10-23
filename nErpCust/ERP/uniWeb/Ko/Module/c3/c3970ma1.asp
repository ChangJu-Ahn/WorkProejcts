<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : C3970MA1.asp
'*  4. Program Name         : MCS
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003-02-19
'*  8. Modified date(Last)  : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "c3970mb1.asp"							'☆: 비지니스 로직 ASP명 

'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

' Grid 1(vspdData1) 
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_MCSItem
Dim C_MCSDTLItem
Dim C_MCSDTLItemNm
Dim C_Amount
Dim C_AcctSeq
Dim C_Seq
Dim C_Type

dim	strFromYYYYMM 
Dim strToYyyyMm 

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size를 조사할 변수 
Dim lgIntFlgMode								'Variable is for Operation Status
Dim lgIsOpenPop

Dim lgStrPrevKey1
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2
Dim lgRadio

Dim strDate
Dim iDBSYSDate


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
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
    
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	Dim strYear,strMonth,strDay
	
	LocSvrDate = "<%=GetSvrDate%>"
	
	lgRadio	= "S"

	Call ggoOper.FormatDate(frm1.txtFromYYYYMM, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtToYYYYMM, Parent.gDateFormat, 2)

	frm1.txtFromYYYYMM.text	= UniConvDateAToB(LocSvrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtFromYYYYMM, Parent.gDateFormat, 2)
	frm1.txtToYYYYMM.text	= UniConvDateAToB(LocSvrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtToYYYYMM, Parent.gDateFormat, 2)

	
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet()


	Call InitSpreadPosVariables()


	With frm1.vspdData1 
			
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
		.ReDraw = false
					
		.MaxCols = C_Type + 1    
		.MaxRows = 0    
			

		Call GetSpreadColumnPos()


		ggoSpread.SSSetEdit 	C_ItemAcct,		"품목계정"	,10 
		ggoSpread.SSSetEdit 	C_ItemAcctNm,   "품목계정명",10
		ggoSpread.SSSetEdit 	C_MCSItem,      "항목"		,25 
		ggoSpread.SSSetEdit 	C_MCSDTLItem,   "세부항목"	,10
		ggoSpread.SSSetEdit 	C_MCSDTLItemNm, "세부항목명",25
		ggoSpread.SSSetFloat 	C_Amount,		"금액"		,30,parent.ggAmtofMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_AcctSeq,		"SEQ"		,5
		ggoSpread.SSSetEdit 	C_Seq,			"SUB_SEQ"	,5
		ggoSpread.SSSetEdit 	C_Type,			"Type"	,5	

	
		Call ggoSpread.MakePairsColumn(C_ItemAcct, C_ItemAcctNm )

		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
		
		Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
		Call ggoSpread.SSSetColHidden( C_MCSDTLItemNM, C_MCSDTLItemNm, True)			
		
		Call ggoSpread.SSSetColHidden( C_AcctSeq, C_AcctSeq, True)			
		Call ggoSpread.SSSetColHidden( C_Seq, C_Seq, True)
		Call ggoSpread.SSSetColHidden( C_Type, C_Type, True)
					
		ggoSpread.SSSetSplit2(4)
			
		Call SetSpreadLock()
			
		.ReDraw = true    
    
	End With
	


    
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData1
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub



'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()

		' Grid 1(vspdData1) - Order Header

	C_ItemAcct		= 1
	C_ItemAcctNm	= 2
	C_MCSItem		= 3
	C_MCSDTLItem	= 4
	C_MCSDTLItemNm 	= 5		
	C_Amount		= 6
	C_AcctSeq		= 7
	C_Seq			= 8
	C_Type			= 9

End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos()
	Dim iCurColumnPos
 	ggoSpread.Source = frm1.vspdData1
		
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	
	C_ItemAcct			= iCurColumnPos(1)
	C_ItemAcctNm		= iCurColumnPos(2)
	C_MCSItem			= iCurColumnPos(3)
	C_MCSDTLItem		= iCurColumnPos(4)
	C_MCSDTLItemNm		= iCurColumnPos(5)
	C_Amount			= iCurColumnPos(6)
	C_AcctSeq			= iCurColumnPos(7)
	C_Seq				= iCurColumnPos(8)
	C_Type				= iCurColumnPos(9)
	

End Sub    


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
'=======================================================================================================
'	Name : OpenWorkStep()
'	Description : Condition Plant PopUp
'=======================================================================================================
Function OpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
		Case 1
			arrParam(0) = "품목계정팝업"			'팝업 명칭 
			arrParam(1) = "(select minor_cd,minor_nm from B_MINOR where MAJOR_CD =" & FilterVar("P1001", "''", "S") & " union all select minor_cd,minor_nm from b_minor where major_cd =" & FilterVar("C2111", "''", "S") & " and minor_cd in (" & FilterVar("MC", "''", "S") & "," & FilterVar("BFO", "''", "S") & "," & FilterVar("EFO", "''", "S") & ")) a "						'TABLE 명칭 
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = ""							'Where Condition
			arrParam(5) = "품목계정"				'TextBox 명칭 
	
			arrField(0) = "a.minor_cd"					'Field명(0)
			arrField(1) = "a.minor_nm"					'Field명(1)
    
			arrHeader(0) = "품목계정"				'Header명(0)
			arrHeader(1) = "품목계정명"				'Header명(1)

		Case 2
			arrParam(0) = "수불유형팝업"			'팝업 명칭 
			arrParam(1) = "B_MINOR"						'TABLE 명칭 
			arrParam(2) = Trim(frm1.txtMovTypeCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = "major_cd = " & FilterVar("I0001", "''", "S") & " "							'Where Condition
			arrParam(5) = "수불유형"				'TextBox 명칭 
	
			arrField(0) = "minor_cd"					'Field명(1)
			arrField(1) = "minor_nm"					'Field명(1)
			
			arrHeader(0) = "수불유형"				'Header명(1)
			arrHeader(1) = "수불유형명"				'Header명(2)
			

	End Select
    

    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopup(arrRet,iWhere)
	End If
		
End Function


Function SetPopup(byval arrRet,byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtItemAcctCd.Value = arrRet(0)		
				.txtItemAcctNm.Value = arrRet(1)		
			Case 2
				.txtMovTypeCd.Value = arrRet(0)		
				.txtMovTypeNm.Value = arrRet(1)		
		End Select

		lgBlnFlgChgValue = True
	End With
End Function

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
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
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet()                                               '⊙: Setup the Spread sheet
   
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetToolBar("11000000000111")										'⊙: 버튼 툴바 제어 
 
    frm1.txtFromYyyymm.focus
	Set gActiveElement = document.activeElement
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'=======================================================================================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromYyyymm_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFromYyyymm.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtFromYyyymm.focus
	End If 
End Sub

Sub txtToYyyymm_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtToYyyymm.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtToYyyymm.focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFromYyyymm_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

Sub txtToYyyymm_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then							'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
   
    End If
    
    lgOldRow = frm1.vspdData1.ActiveRow
			

					
    
End Sub



'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub



'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub



'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos()
End Sub


'==========================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'==========================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

End Sub


'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData1 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub



Function Radio1_onChange()
	
	IF lgRadio = "S" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, True)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()
	
	lgRadio = "S"
	
	lgBlnFlgChgValue = True
End Function

Function Radio2_onChange()

	IF lgRadio = "D" Then
		Exit Function
	ENd IF

	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, False)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, False)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()


	lgRadio = "D"
	
	lgBlnFlgChgValue = True
End Function

Function Radio3_onChange()
	
	IF lgRadio = "S1" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, True)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()
	
	lgRadio = "S1"
	
	lgBlnFlgChgValue = True
End Function

Function Radio4_onChange()

	IF lgRadio = "D1" Then
		Exit Function
	ENd IF

	ggoSpread.Source = frm1.vspdData1	
	
	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, False)			
	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, False)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()


	lgRadio = "D1"
	
	lgBlnFlgChgValue = True
End Function
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

    FncQuery = False														'⊙: Processing is NG
    Err.Clear																'☜: Protect system from crashing

	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    Call InitVariables														'⊙: Initializes local global variables
	
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'⊙: This function check indispensable field
       Exit Function
    End If
    
    
    If CompareDateByFormat(frm1.txtFromYYYYMM.Text,frm1.txtToYYYYMM.Text,frm1.txtFromYYYYMM.Alt,frm1.txtToYYYYMM.Alt, _
	 "970024", frm1.txtFromYYYYMM.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtFromYYYYMM.Focus
		Exit Function
	End If    
    '-----------------------
    'Query function call area
    '-----------------------



    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'☜: Query db data
	End If
	
    FncQuery = True															'⊙: Processing is OK
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next													'☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'☜: Protect system from crashing
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'******************  5.2 Fnc함수명에서 호출되는 개발 Function  **************************
'	설명 : 
'**************************************************************************************** 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal
    Dim strYear,strMonth,strDay
    
    DbQuery = False

	Call LayerShowHide(1)
    
    Call ExtractDateFrom(frm1.txtFromYyyyMm.Text,frm1.txtFromYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strFromYYYYMM = strYear & strMonth

    Call ExtractDateFrom(frm1.txtToYyyyMm.Text,frm1.txtToYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strToYYYYMM = strYear & strMonth

	With frm1
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'☜: 
			strVal = strVal & "&txtFromYyyymm="		& strFromYYYYMM			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtToYyyymm="		& strToYYYYMM			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemAcctCd="		& .txtItemAcctCd.Value
			strVal = strVal & "&txtMovTypeCd="		& .txtMovTypeCd.Value
			strVal = strVal & "&txtRadio="			& lgRadio
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'☜: 
			strVal = strVal & "&txtFromYyyymm="		& strFromYYYYMM			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtToYyyymm="		& strToYYYYMM			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemAcctCd="		& .txtItemAcctCd.Value
			strVal = strVal & "&txtMovTypeCd="		& .txtMovTypeCd.Value
			strVal = strVal & "&txtRadio="			& lgRadio
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
		End If
	End With
    
    Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11000000000111")														'⊙: 버튼 툴바 제어 
	lgIntFlgMode = parent.OPMD_UMODE														'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	
		
End Function


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub 

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>제조원가명세서조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>					
					<TD>&nbsp;</TD>					
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
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/c3970ma1_fpDateTime1_txtFromYYYYMM.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/c3970ma1_fpDateTime2_txtToYYYYMM.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_Sum Checked tag = 2 value="01" onclick=radio1_onchange()><LABEL FOR=Rb_Sum>집계</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Dtl tag = 2 value="02" onclick=radio2_onchange()><LABEL FOR=Rb_Dtl>상세</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Sum1 tag = 2 value="03" onclick=radio3_onchange()><LABEL FOR=Rb_Sum1>집계Sim</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Dtl1 tag = 2 value="04" onclick=radio4_onchange()><LABEL FOR=Rb_Dtl1>상세Sim</LABEL></TD>										        							
								</TR>
								<TR>								
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItemAcctCd" SIZE=9 MAXLENGTH=10 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
										 <INPUT TYPE=TEXT ID="txtItemAcctNm" NAME="txtItemAcctNm" SIZE=25 tag="14X">
									</TD>
									<TD CLASS="TD5">수불유형</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtMovTypeCd" SIZE=9 MAXLENGTH=10 tag="11XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovTypeCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(2)">
										 <INPUT TYPE=TEXT ID="txtMovTypeNm" NAME="txtMovTypeNm" SIZE=25 tag="14X">
									</TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% VALIGN=top COLSPAN=4>
						<script language =javascript src='./js/c3970ma1_vaSpread1_vspdData1.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>

	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hMovTypeCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

