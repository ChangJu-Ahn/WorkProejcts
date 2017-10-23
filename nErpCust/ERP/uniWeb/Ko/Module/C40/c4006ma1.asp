<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4006MA1.asp
'*  4. Program Name			:완성품환산율등록 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4006Mb1.asp
'*						
'*  7. Modified date(First)	: 2005/08/29
'*  8. Modified date(Last)	: 2005/11/03
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: HJO
'* 11. Comment				: 
'* 12. History              : 
'*                          : 
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID	= "c4006mb1.asp"			'☆:  비지니스 로직 ASP명 
Const BIZ_PGM_ID2= "c4006mb2.asp"			'☆:  비지니스 로직 ASP명 


Dim LocSvrDate

LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)													

Dim C_WcCd				'= 1
Dim C_WcPop		'=2
Dim C_WcNm				'= 3
Dim C_ItemCd
Dim C_ItemCdPop
Dim C_ItemNM
Dim C_OrderNo				'= 4
Dim C_OrderNoPop	'=5
Dim C_ProdRate				'= 6


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop						' Popup


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

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
	lgBlnFlgChgValue = False
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
    lgSortKey = 1

End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    
    Call InitSpreadPosVariables()
    
    With frm1
           
    ggoSpread.Source = .vspdData
    ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
    
    Call AppendNumberPlace("6","3","0")
      
    
    .vspdData.ReDraw = False
    
    .vspdData.MaxCols = C_ProdRate + 1
    .vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	
    
    ggoSpread.SSSetEdit		C_WcCd,				"공정", 10,,,7
    ggoSpread.SSSetButton C_WcPop
    ggoSpread.SSSetEdit		C_WcNm,				"공정명", 30
    ggoSpread.SSSetEdit		C_ItemCd,				"품목", 10,,,18
    ggoSpread.SSSetButton C_ItemCdPop
    ggoSpread.SSSetEdit		C_ItemNm,				"품목명", 30    
    ggoSpread.SSSetEdit		C_OrderNo,			"제조번호", 18,,,18
    ggoSpread.SSSetButton C_OrderNoPop
    ggoSpread.SSSetFloat		C_ProdRate,			"완성품환산율", 15,"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,,	"Z"  ,"0","100"
   
 	Call ggoSpread.SSSetColHidden(.vspdData.MaxCols ,.vspdData.MaxCols , True)
		
    ggoSpread.SSSetSplit2(2) 
	.vspdData.ReDraw = False
	
    End With
   
    Call SetSpreadLock()
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData    
	'ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData 
    
    .Redraw = False

    ggoSpread.Source = frm1.vspdData    
    ggoSpread.SSSetRequired C_WcCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_ProdRate, pvStartRow, pvEndRow
    
    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
    .EditMode = True
    
    .Redraw = True
    
    End With
End Sub
'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetQuerySpreadColor
' Function Desc :  This method set color and protect  in spread sheet celles, after Query
'========================================================================================

Sub SetQuerySpreadColor(ByVal lRow, byVal mRow)
    
    With frm1
		.vspdData.ReDraw = False
  
		ggoSpread.SSSetProtected C_WcCd, lRow, mRow
		ggoSpread.SSSetProtected C_WcPop, lRow, mRow
		ggoSpread.SSSetProtected C_WcNm, lRow, mRow
		ggoSpread.SSSetProtected C_ItemNm,lRow, mRow
		ggoSpread.SSSetProtected C_OrderNo, lRow, mRow
		ggoSpread.SSSetProtected C_OrderNoPop, lRow,mRow
		ggoSpread.SSSetRequired C_ProdRate, lRow,mRow
		.vspdData.ReDraw = True
	End With
End Sub
'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_WcCd				= 1
	C_WcPop				=2
	C_WcNm				= 3
	C_ItemCd				= 4
	C_ItemCdPop				=5
	C_ItemNm				= 6
	C_OrderNo				= 7
	C_OrderNoPop			=8
	C_ProdRate				= 9
End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		C_WcCd				= iCurColumnPos(1)
 		C_WcPop				= iCurColumnPos(2)
		C_WcNm				= iCurColumnPos(3)
		C_ItemCd				= iCurColumnPos(4)
 		C_ItemCdPop				= iCurColumnPos(5)
		C_ItemNm				= iCurColumnPos(6)
		C_OrderNo				= iCurColumnPos(7)
		C_OrderNoPop				= iCurColumnPos(8)
		C_ProdRate				= iCurColumnPos(9)		
 	End Select 
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
    
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

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"						' Header명(0)
    arrHeader(1) = "공장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , frm1.txtPlantCd.alt,"X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  OpenWcCd()  -------------------------------------------------
'	Name : OpenWcCd()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCd()
	Dim arrRet
	Dim strWhere
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , frm1.txtPlantCd.alt,"X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	strWhere= " plant_cd=" & FilterVar(Trim(frm1.txtPlantCd.Value),"''","S")	
	strWhere = strWhere & " and  convert(varchar(7),valid_from_dt,120)<=	" & FilterVar(frm1.txtYYYYMM.text,"''","S")
	strWhere = strWhere & "	 and convert(varchar(7),valid_to_dt,120) >=" & FilterVar(frm1.txtYYYYMM.text,"''","S")
	
	arrParam(0) = "공정팝업"						' 팝업 명칭 
	arrParam(1) = "p_work_center"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWcCd.Value)	' Code Condition
	arrParam(3) =""										' Name Cindition
	arrParam(4) =strWhere							' Where Condition
	arrParam(5) = "공정"							' TextBox 명칭 
	
    arrField(0) ="ED10" & Parent.gColSep &  "WC_CD"					' Field명(0)
    arrField(1) = "ED31" & Parent.gColSep & "WC_NM"					' Field명(1)
    
    
    arrHeader(0) = "공정"						' Header명(0)
    arrHeader(1) = "공정명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWcCd.focus
	
End Function



'==========================================  2.4.3 Set Return Value()  =============================================
'	Name : Set Return Value()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
    frm1.txtPlantCd.Value    = arrRet(0)		
    frm1.txtPlantNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetWcCd()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWcCd(byval arrRet)
	frm1.txtWcCd.Value    = arrRet(0)		
	frm1.txtWcNm.Value   = arrRet(1)
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
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables

	'----------  Coding part  -------------------------------------------------------------
	'Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	Call SetToolBar("11001111001111")											'⊙: 버튼 툴바 제어	

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If
		frm1.txtYYYYMM.Text=LocSvrDate		
		Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
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


'=========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    with frm1.vspdData
		.Col = Col
		.Row = Row
		Select Case Col
		Case C_WcCd
			Call checkWcCd(Row,.Text)    
		Case C_ItemCd
			Call checkItemCd(Row,.Text)    
		Case C_OrderNo    
		    Call checkProdOrderNo(Row, .Text)
		End Select
	End With
    
End Sub
'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================

Sub vspdData_EditChange(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row        

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================

Sub vspdData_DblClick(ByVal Col , ByVal Row )
Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub
'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
   End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	If frm1.vspdData.MaxRows <0 Then Exit Sub
	
	With frm1.vspdData 		
		ggoSpread.Source = frm1.vspdData
	    Select Case Col
			Case  C_WcPop 
				.Col = Col :		    .Row = Row
	
				Call OpenSpreadPopup(C_WcPop, Row, .Text)		
				Call SetActiveCell(frm1.vspdData,C_ItemCd,Row,"M","X","X")			
			Case C_ItemCdPop
				.Col=Col :				.Row=Row
				Call OpenSpreadPopup(C_ItemCdPop, Row, .Text)
				Call SetActiveCell(frm1.vspdData,C_OrderNo,Row,"M","X","X")			
			Case C_OrderNoPop 
				.Col = C_OrderNo :			.Row = Row

				Call OpenSpdOrderNoPop(Row, .Text)
				Call SetActiveCell(frm1.vspdData,C_ProdRate,Row,"M","X","X")
		End Select
    Call vspdData_Change(Col,Row)
	End With	
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
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
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call SetQuerySpreadColor(1,1)
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : txtYYYYMM_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtYYYYMM.Action = 7
        Call SetFocusToDocument("P")
		Frm1.txtYYYYMM.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtYYYYMM_KeyDown
'   Event Desc : 
'=======================================================================================================
Sub  txtYYYYMM_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
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
'********************************************************************************************************* %>
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData										'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = True Then									'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    

	IF ChkKeyField()=False Then Exit Function 
    '-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field   
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables		

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If															'☜: Query db data
       
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
    Dim IntRetCD 
    Dim iRow
    FncSave = False                                           '⊙: Processing is NG
    
    Err.Clear                                                 '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 

    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
       Exit Function
    End If

    For iRow=1  to frm1.vspdData.MaxRows			
        frm1.vspdData.Row = iRow
        frm1.vspdData.Col = 0			
		Select Case frm1.vspdData.Text
			Case ggoSpread.InsertFlag				
				frm1.vspdData.Col = C_WcCd				
				If  checkWcCd(iRow,frm1.vspdData.Text)=False Then Exit Function 
				
				frm1.vspdData.Col = C_ItemCd				
				If frm1.vspdData.Text <>"" and frm1.vspdData.Text <>"*" Then 
					If  checkItemCd(iRow,frm1.vspdData.Text)=False Then Exit Function 
				End If
				frm1.vspdData.Col = C_OrderNo
				If  frm1.vspdData.Text <>"" and   frm1.vspdData.Text<>"*"  Then
					If  checkProdOrderNo(iRow, frm1.vspdData.Text)=False Then Exit Function 
				End If
		End Select	
	Next
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function				                                  '☜: Save db data
    
    FncSave = True                                            '⊙: Processing is OK
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow

			SetSpreadColor .ActiveRow, .ActiveRow			

            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	With frm1 
		.vspdData.Col = C_WcCd :.vspdData.Text = ""
		.vspdData.Col = C_WcNM :.vspdData.Text = ""

	End With
		
	Call SetActiveCell(frm1.vspdData,C_WcCd,frm1.vspdData.ActiveRow,"M","X","X")
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
    	
    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================

Function FncPaste() 
     ggoSpread.SpreadPaste
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow, i
    Dim iIntIndex
    
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
        ggoSpread.InsertRow ,imRow

        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1		
			
        .vspdData.ReDraw = True
        lgBlnFlgChgValue = True  
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
	Dim lDelRows, lDelRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    	
    End With

    lgBlnFlgChgValue = True 
   
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												<%'☜: 화면 유형 %>
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

    Dim IntRetCD
    
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
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
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================

Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function


'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================

Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function
'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'******************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    
    DbQuery = False    
    Call LayerShowHide(1)
 
    Dim strVal
    Dim sStartDt,sYear,sMon,sDay
    
    Call parent.ExtractDateFromSuper(frm1.txtYYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)	
	sStartDt= (sYear&sMon)
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtYYYYMM=" & Trim(frm1.hYYYYMM.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)				'☆: 조회 조건 데이타		
	Else
		strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtYYYYMM=" & Trim(sStartDt)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)				'☆: 조회 조건 데이타	
	End If

    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()													'☆: 조회 성공후 실행로직	
	
	
	Call SetToolBar("11001111001111")											'⊙: 버튼 툴바 제어	
   '-----------------------
    'Reset variables area
    '-----------------------

    Call SetQuerySpreadColor(-1,-1)

    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
		
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode


End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
	Dim lRow        
    Dim lGrpCnt    
    Dim strVal
	Dim	strChangeFlag
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size

	
    DbSave = False                                                          	'⊙: Processing is NG
    
    Call LayerShowHide(1)
    
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0

    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows
		
		strVal = ""
		
        .vspdData.Row = lRow
        .vspdData.Col = 0
			
		Select Case .vspdData.Text
		
			Case ggoSpread.UpdateFlag
				strVal = strVal & "U" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case ggoSpread.InsertFlag
				strVal = strVal & "C" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case ggoSpread.DeleteFlag
				strVal = strVal & "D" & iColSep			'☜: C=Create
				strChangeFlag = "Y"
			Case Else				
				strChangeFlag = "N"
		End Select

		If strChangeFlag = "Y" Then 
			strVal = strVal &lRow & iColSep	
			.vspdData.Col = C_WcCd
			strVal = strVal & Trim(.vspdData.Text) & iColSep																				
			.vspdData.Col = C_ItemCd
			strVal = strVal & Trim(.vspdData.Text) & iColSep					
			.vspdData.Col = C_OrderNo
			strVal = strVal & Trim(.vspdData.Text) & iColSep			
			.vspdData.Col = C_ProdRate
			strVal = strVal & cdbl(Trim(.vspdData.Text)/100) & iColSep 
			'row count
			strVal = strVal & lRow & parent.gRowSep			

		End If
        
        .vspdData.Col = 0
		Select Case .vspdData.Text
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
				    
		         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
					objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
				 
		           iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
				       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
				      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
				         
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
				         
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strVal) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
				          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
				       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
				         
		         iTmpDBuffer(iTmpDBufferCount) =  strVal         
		         strDTotalvalLen = strDTotalvalLen + Len(strVal)
				         
		End Select
                
    Next
    
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   	

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'☜: 비지니스 ASP 를 가동 
		
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0	
    lgBlnFlgChgValue = False    
    
    Call RemovedivTextArea
    Call MainQuery()

End Function


'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'----------  Coding part  -------------------------------------------------------------
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function

'===========================================================================================================
' Description : checkWcCd ;check valid wccd
'===========================================================================================================
Function checkWcCd(ByVal pvLngRow, ByVal pvStrData)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrWcCdInf
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	
	checkWcCd = False
	
	iStrSelectList = " WC_NM "
	iStrFromList   = " dbo.P_WORK_CENTER "
	iStrWhereList  = " PLANT_CD = " & FilterVar((frm1.txtPlantCd.value), "''", "S")
	iStrWhereList =  iStrWhereList & " AND WC_CD =  " & FilterVar(pvStrData , "''", "S") 

	Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("970000","X",frm1.txtWcCd.alt,"X")
		frm1.vspdData.Row=pvLngRow :frm1.vspdData.Col = C_WcNm : frm1.vspdData.Text =""
		Call SetActiveCell(frm1.vspdData,C_WcCd,pvLngRow,"M","X","X")			
		checkWcCd = False
		Exit Function
	End If	
	With frm1.vspdData
		iArrWcCdInf = split(lgF0,chr(11))
		.Row = pvLngRow
		.Col = C_WcNm	:  .text = Trim(iArrWcCdInf(0))			
	End With
	checkWcCd = True
End Function

'===========================================================================================================
' Description : checkItemCd ;check valid wccd
'===========================================================================================================
Function checkItemCd(ByVal pvLngRow, ByVal pvStrData)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrWcCdInf
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	
	checkItemCd = False
	
	iStrSelectList = " a.Item_NM "
	iStrFromList   = "  b_item a  "
	iStrFromList   = iStrFromList   & "	inner join b_item_by_plant b on a.item_cd=b.item_cd "
	iStrFromList   = iStrFromList   & "	inner join b_item_acct_inf c on c.item_acct=b.item_acct  "
	iStrWhereList  = " b.PLANT_CD = " & FilterVar((frm1.txtPlantCd.value), "''", "S")
	iStrWhereList =  iStrWhereList & " AND c.item_acct_group in ('1final','2semi') "

	Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("970000","X","품목","X")
		frm1.vspdData.Row=pvLngRow :frm1.vspdData.Col = C_ItemNm : frm1.vspdData.Text =""
		Call SetActiveCell(frm1.vspdData,C_ItemCd,pvLngRow,"M","X","X")			
		checkItemCd = False
		Exit Function
	End If	
	With frm1.vspdData
		iArrWcCdInf = split(lgF0,chr(11))
		.Row = pvLngRow
		.Col = C_ItemNm	:  .text = Trim(iArrWcCdInf(0))			
	End With
	checkItemCd = True
End Function
'===========================================================================================================
' Description : checkProdOrderNo  ; check valid prod order no 
'===========================================================================================================
Function checkProdOrderNo(ByVal pvLngRow, ByVal pvStrData)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrProdNoInf
		
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	
	checkProdOrderNo = False
	
	iStrSelectList = " a.prodt_order_no  "
	iStrFromList   = " p_production_order_header a, b_item b,b_storage_location c,b_item_by_plant d  "
	iStrWhereList = "  a.item_cd = b.item_cd and	a.plant_cd = d.plant_cd and	a.item_cd = d.item_cd and	a.sl_cd = c.sl_cd "	
	iStrWhereList  =iStrWhereList & " AND a.PLANT_CD = " & FilterVar(trim(frm1.txtPlantCd.value), "''", "S") & " AND a.prodt_order_no =  " & FilterVar(pvStrData , "''", "S") & ""
	'iStrWhereList = iStrWhereList & " and 	a.order_status in (  'OP', 'RL', 'RL' ) "
	
	Call CommonQueryRs(" a.prodt_order_no ",iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("970000","X",frm1.txtProdOrderNo.alt,"X")
		Call SetActiveCell(frm1.vspdData,C_OrderNo,pvLngRow,"M","X","X")			
		checkProdOrderNo = False
		Exit Function
	End If	
	
	With frm1.vspdData
		iArrProdNoInf = split(lgF0,chr(11))
		.Row = pvLngRow
		.Col = C_OrderNo	: .text = Trim(iArrProdNoInf(0))			
	End With
	checkProdOrderNo = True
End Function

'===========================================================================================================
' Description :spread popup button 
'===========================================================================================================
Function OpenSpreadPopup(ByVal pvLngCol, ByVal pvLngRow, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenSpreadPopup = False
	
	If IsOpenPop Then Exit Function

	IsOpenPop = True
	
	Select Case pvLngCol
		Case C_WcPop
			iArrParam(1) = "p_work_center"			<%' TABLE 명칭 %>
			iArrParam(2) = pvStrData					<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = "plant_cd=" & FilterVar(trim(frm1.txtPlantCd.value), "''", "S") & "  AND convert(char(7),valid_from_dt,120) <=" & FilterVar(frm1.txtYYYYMM.Text, "''", "S") & " "	<%' Where Condition%>
			iArrParam(4) = iArrParam(4) & "  AND convert(char(7),valid_to_dt,120) >=" & FilterVar(frm1.txtYYYYMM.Text, "''", "S") & " "	<%' Where Condition%>			
			iArrParam(5) = "공정"						<%' TextBox 명칭 %>
				
			iArrField(0) = "ED10" & Parent.gColSep & "WC_CD"
			iArrField(1) = "ED30" & Parent.gColSep & "WC_NM"
			
			    
			iArrHeader(0) = "공정"
			iArrHeader(1) = "공정명"
			
		Case C_ItemCdPop
			iArrParam(1) = " (select a.item_cd, a.item_nm, b.item_acct, c.item_acct_group "
			iArrParam(1) = iArrParam(1) & "	from b_item a	"
			iArrParam(1) = iArrParam(1) & "	inner join b_item_by_plant b on a.item_cd=b.item_cd	"
			iArrParam(1) = iArrParam(1) & "	inner join b_item_acct_inf c on c.item_acct=b.item_acct	"
			iArrParam(1) = iArrParam(1) & "	where b.plant_cd =" & FilterVar(trim(frm1.txtPlantCd.value), "''", "S")
			iArrParam(1) = iArrParam(1) & "		and c.item_acct_group in ('1final','2semi') ) A "			<%' TABLE 명칭 %>
			
			iArrParam(2) = pvStrData					<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = ""
			
			iArrParam(5) = "품목"						<%' TextBox 명칭 %>
				
			iArrField(0) = "ED10" & Parent.gColSep & "ITEM_CD"
			iArrField(1) = "ED30" & Parent.gColSep & "ITEM_NM"
			iArrField(2) = "ED10" & Parent.gColSep & "ITEM_ACCT"
			iArrField(3) = "ED10" & Parent.gColSep & "ITEM_ACCT_GROUP"			
			    
			iArrHeader(0) = "품목"
			iArrHeader(1) = "품목명"
			iArrHeader(2) = "계정"
			iArrHeader(3) = "계정그룹"
			
		Case C_OrderNoPop
			OpenSpreadPopup = OpenSpdOrderNoPop(pvLngRow, pvStrData)
			Exit Function		
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvLngCol, pvLngRow)
	End If	

End Function
'===========================================================================================================
' Description : set spread popup 
'===========================================================================================================
Function SetSpreadPopup(Byval pvArrRet,ByVal pvLngCol, ByVal pvLngRow)
	SetSpreadPopup = False

	With frm1.vspdData
		.Row = pvLngRow
		
		Select Case pvLngCol
			Case C_WcPop
				.Col = C_WcCd	: .Text = pvArrRet(0)
				.Col = C_WcNm	: .Text = pvArrRet(1)
			Case C_ItemCdPop
				.Col = C_ItemCd	: .Text = pvArrRet(0)
				.Col = C_ItemNm	: .Text = pvArrRet(1)		
			Case C_OrderNoPop
				.Col = C_OrderNo : .Text = PvArrRet(0)
		End Select
	End With

	SetSpreadPopup = True
End Function
'===========================================================================================================
' Description : spread orderno pop
'===========================================================================================================
Function OpenSpdOrderNoPop(ByVal pvLngRow, ByVal pvStrData)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , frm1.txtPlantCd.alt,"X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	If IsOpenPop then Exit Function 
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	msgbox Trim(pvStrData)

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	'arrParam(3) = "OP"
	arrParam(3) = ""
	'arrParam(4) = "RL"
	arrParam(4) = ""
	arrParam(5) = Trim(pvStrData)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		Exit Function
	Else
	OpenSpdOrderNoPop = SetSpreadPopup(arrRet, C_OrderNoPop, pvLngRow)
	End If 
	
End Function
'========================================================================================================
Sub btnCopyPrev_OnClick()

	If BtnSpreadCheck = False Then Exit Sub

	Err.Clear                                                        

	If  CheckExistData1() Then 
		Call CheckExistData2()
	End If	
	frm1.txtProdOrderNo.focus()

End Sub
'===========================================================================================================
' Description : CheckExistData ;Check Exist about the previous data 
'===========================================================================================================
Function CheckExistData1()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iTmp
	Dim IntRetCD
	
	Dim PrevDate
	
	CheckExistData1=FALSE
	
	PrevDate	= UNIDateAdd("m", -1, frm1.txtYYYYMM.Text, parent.gDateFormat)
	frm1.txtYYYYMM2.value = replace(left(PrevDate,7),"-","")
		
	iStrSelectList = " top 1 yyyymm "
	iStrFromList   = " c_prod_rate_by_ors_s"
	iStrWhereList  =iStrWhereList & " yyyymm = " & FilterVar(replace(left(PrevDate,7),"-",""), "''", "S")	

	Err.Clear

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		CheckExistData1=TRUE
		Exit Function 
	Else   
		If Err.number = 0 Then   'Data is not exist.
			 Call DisplayMsgBox("236306","X" , "X","X")
			 CheckExistData1=FALSE
		Else								'Err.
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If

End Function
'===========================================================================================================
' Description : CheckExistData2;Check exist about current data
'===========================================================================================================
Function CheckExistData2()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iTmp
	Dim IntRetCD
	
	
	iStrSelectList = " top 1 yyyymm "
	iStrFromList   = " c_prod_rate_by_ors_s"
	iStrWhereList  =iStrWhereList & " yyyymm = " & FilterVar(replace(frm1.txtYYYYMM.Text,"-",""), "''", "S")	

	Err.Clear

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		IntRetCD = DisplayMsgBox("900007", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then 
			Exit Function
		ELSE
			Call CopyPrevData()		
		END IF
	Else   
		If Err.number = 0 Then   'Data is not exist.
			Call CopyPrevData()
		Else								'Err.
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If
	
End Function
'========================================================================================================
' Description : CopyPrevData;Copy data
'===========================================================================================================
Sub CopyPrevData()
	
	Dim iStrVal

	iStrVal = BIZ_PGM_ID & "?txtMode=" & "btnCopyPrev"					
	iStrVal = iStrVal & "&txtYYYYMM1=" & Trim(frm1.txtYYYYMM.Text)
	iStrVal = iStrVal & "&txtYYYYMM2=" & Trim(frm1.txtYYYYMM2.value)		

	Call RunMyBizASP(MyBizASP, iStrVal)          

End Sub

'========================================================================================================
' Description : BtnSpreadCheck;Check changed data before anyother event
'===========================================================================================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData 

	 '--case multi -- %>
	 'when changed data exist asking what to do  %>
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	 'nothing changed  %>
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function


'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		
'check plant
	If Trim(frm1.txtPlantCd.value) <> "" Then
		strWhere = " plant_cd= " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	 b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlantCd.alt,"X")			
			frm1.txtPlantnm.value = ""
			ChkKeyField = False
			frm1.txtPlantCd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlantNM.value = strDataNm(0)
	End If
'check wc cd	
	If Trim(frm1.txtWcCd.value) <> "" Then
		strWhere = " Wc_Cd = " & FilterVar(frm1.txtWcCd.value, "''", "S") & " "		
		
		Call CommonQueryRs(" wc_Nm ","	 p_work_center ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtWcCd.alt,"X")			
			frm1.txtWcNM.value = ""
			ChkKeyField = False
			frm1.txtWcCd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtWcNM.value = strDataNm(0)
	Else
		frm1.txtWcNm.value=""
	End If
'check prod order no	
	If Trim(frm1.txtProdOrderNo.value) <> "" Then
		strFrom = " p_production_order_header a, b_item b,b_storage_location c,b_item_by_plant d "
		strWhere = " a.prodt_order_no = " & FilterVar(frm1.txtProdOrderNo.value, "''", "S") & " "		
		strWhere =strWhere & " and a.item_cd = b.item_cd and	a.plant_cd = d.plant_cd and	a.item_cd = d.item_cd and	a.sl_cd = c.sl_cd"	
		strWhere =strWhere & " AND a.PLANT_CD = " & FilterVar(trim(frm1.txtPlantCd.value), "''", "S") 
		'strWhere =strWhere & " and 	a.order_status in (  'OP', 'RL', 'RL' ) "
		
		Call CommonQueryRs(" a.prodt_order_no ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtProdOrderNo.alt,"X")			
			frm1.txtProdOrderNo.value = ""
			ChkKeyField = False
			frm1.txtProdOrderNo.focus 
			Exit Function
		End If	
		strDataNm = split(lgF0,chr(11))
		frm1.txtProdOrderNo.value = strDataNm(0)
	End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>완성품환산율등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
								<TD CLASS=TD5 NOWRAP>작업년월</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME tag="12" ALT="작업년월" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
									<TD CLASS=TD5 NOWRAP>공정</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=25 tag="14"></TD>
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
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>	
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
	 <TABLE <%=LR_SPACE_TYPE_30%>>
	     <TR>
	   <TD WIDTH=10>&nbsp;</TD>
	   <TD>
	    <BUTTON NAME="btnCopyPrev" CLASS="CLSSBTN">전월COPY</BUTTON>&nbsp;
	    </TD>
	    </TR>
	   </TABLE>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=bizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24"><INPUT TYPE=HIDDEN NAME="txtYYYYMM2" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
