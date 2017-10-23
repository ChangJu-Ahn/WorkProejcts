
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID  = "p4911mb1.asp"								'☆: 비지니스 로직(Qeury) ASP명 

Const BIZ_PGM_SAVE_ID = "p4911mb2.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_WcCd
Dim C_WcPopup
Dim C_WcNm
Dim C_ItemCd
Dim C_ItemPopup
Dim C_ItemNm
Dim C_StApply
Dim C_StFixed
Dim C_AstFixed
Dim C_FrDt
Dim C_EndDt
Dim C_UseFlg

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2. Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntGrpCount              ' GroupView Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgSortKey
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop
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

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------

    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count

End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'===========================================================================================================
Sub SetDefaultVal()
    frm1.txtBaseDt.text = StartDate
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ====================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'============================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	With frm1.vspdData

	.ReDraw = false
    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

    .MaxCols = C_UseFlg + 1
    .MaxRows = 0

	Call AppendNumberPlace("6", "3", "0")
	Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_WcCd, "작업장", 10, 0,,7,2
    ggoSpread.SSSetButton	C_WcPopup
    ggoSpread.SSSetEdit		C_WcNm, "작업장명", 16, 0,,40,2
    ggoSpread.SSSetEdit		C_ItemCd, "품목", 18, 0,,18,2
    ggoSpread.SSSetButton	C_ItemPopup
    ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25
    ggoSpread.SSSetTime		C_StApply, "적용ST", 12, 2, 1, 1
    ggoSpread.SSSetTime		C_StFixed, "불변ST", 12, 2, 1, 1
    ggoSpread.SSSetTime		C_AstFixed, "고객사ST", 12, 2, 1, 1
    ggoSpread.SSSetDate		C_FrDt, "적용시작일", 11, 2, parent.gDateFormat
    ggoSpread.SSSetDate		C_EndDt, "적용종료일", 11, 2, parent.gDateFormat
    ggoSpread.SSSetCheck	C_UseFlg, "사용여부", 10,,,1

	Call ggoSpread.MakePairsColumn(C_WcCd, C_WcPopup)
	Call ggoSpread.MakePairsColumn(C_ItemCd, C_ItemPopup)

	Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
	ggoSpread.SSSetSplit2(3)											'frozen 기능 추가 

	.ReDraw = true

	Call SetSpreadLock

    End With

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_WcCd,			-1, C_WcCd
		ggoSpread.SpreadLock C_WcPopup,			-1, C_WcPopup
		ggoSpread.SpreadLock C_WcNm,			-1, C_WcNm
		ggoSpread.SpreadLock C_ItemCd,			-1, C_ItemCd
		ggoSpread.SpreadLock C_ItemPopup,		-1, C_ItemPopup
		ggoSpread.SpreadLock C_ItemNm,			-1, C_ItemNm
		ggoSpread.SpreadLock C_FrDt,			-1, C_FrDt
		ggoSpread.SSSetRequired C_EndDt,		-1, C_EndDt

		ggoSpread.SSSetProtected .vspdData.MaxCols, -1
		.vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc :
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal InOutType)

	ggoSpread.Source = frm1.vspdData
    If InOutType = "N" Then
		ggoSpread.SSSetRequired 	C_WCCd, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_WcNm, 		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ItemCd,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ItemNm,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_FrDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_EndDt,		pvStartRow, pvEndRow
	Else
		ggoSpread.SSSetRequired 	C_WCCd, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_WcNm, 		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ItemCd,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ItemNm,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_FrDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_EndDt,		pvStartRow, pvEndRow
	End If
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_WcCd				= 1
	C_WcPopup			= 2
	C_WcNm				= 3
	C_ItemCd			= 4
	C_ItemPopup			= 5
	C_ItemNm			= 6
	C_StApply			= 7
	C_StFixed			= 8
	C_AstFixed			= 9
	C_FrDt				= 10
	C_EndDt				= 11
	C_UseFlg			= 12
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
  	Dim iCurColumnPos

 	Select Case Ucase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_WcCd				= iCurColumnPos(1)
			C_WcPopup			= iCurColumnPos(2)
			C_WcNm				= iCurColumnPos(3)
			C_ItemCd			= iCurColumnPos(4)
			C_ItemPopup			= iCurColumnPos(5)
			C_ItemNm			= iCurColumnPos(6)
			C_StApply			= iCurColumnPos(7)
			C_StFixed			= iCurColumnPos(8)
			C_AstFixed			= iCurColumnPos(9)
			C_FrDt				= iCurColumnPos(10)
			C_EndDt				= iCurColumnPos(11)
			C_UseFlg			= iCurColumnPos(12)
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
'++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 

    arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)

    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus

End Function

'------------------------------------------  OpenItemInfo()  ------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'------------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(strCode)					' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value

	arrField(0) = 1 '"ITEM_CD"					' Field명(0)
	arrField(1) = 2 '"ITEM_NM"					' Field명(1)

    iCalledAspName = AskPRAspName("b1b11pa3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet, iWhere)
	End If

	Select Case iWhere
		Case 0				' frm1
			Call SetFocusToDocument("M")
			frm1.txtItemCd.focus
		Case 1				' Grid
			frm1.vspdData.focus
	End Select
End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet, Byval iWhere)
	Select Case iWhere
		Case 0				' frm1
			With frm1
				.txtItemCd.Value    = arrRet(0)
				.txtItemNm.Value    = arrRet(1)
			End With
		Case 1				' Grid
			 With frm1.vspdData
				.Col	= C_ItemCd
				.Text	= arrRet(0)
				.Col	= C_ItemNm
				.Text	= arrRet(1)
			 End With
	End Select
End Function

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenConWC(ByVal strCode, ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"												' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = Trim(strCode)												' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") 			' Where Condition
	arrParam(5) = "작업장"													' TextBox 명칭 

    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)

    arrHeader(0) = "작업장"													' Header명(0)
    arrHeader(1) = "작업장명"												' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet, iWhere)
	End If

	Select Case iWhere
		Case 0				' frm1
			Call SetFocusToDocument("M")
			frm1.txtWCCd.focus
		Case 1				' Grid
			frm1.vspdData.focus
	End Select

End Function

'------------------------------------------  SetConWC()  ----------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet, ByVal iWhere)
	Select Case iWhere
		Case 0
			With frm1
				.txtWCCd.Value    = arrRet(0)
				.txtWCNm.Value    = arrRet(1)
			End With
		Case 1
			With frm1.vspdData
				.Col	= C_WcCd
				.Text	= arrRet(0)
				.Col	= C_WcNm
				.Text	= arrRet(1)
			End With
	End Select
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
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
'**********************************************************************************************************

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

'******************************  3.2.1 Object Tag 처리  **************************************************
'	Window에 발생 하는 모든 Even 처리 
'*********************************************************************************************************

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

  	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 

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
 	Else
  		'------ Developer Coding part (Start)
 	 	'------ Developer Coding part (End)
  	End If

End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc :
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc :
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_SubconPrc
            Call EditModeCheck(frm1.vspdData, Row, C_CurCd, C_SubconPrc, "C" ,"I", Mode, "X", "X")
    End Select
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

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------
	With frm1.vspdData

    ggoSpread.Source = frm1.vspdData

    If Row > 0 And Col = C_WcPopUp Then
        .Col = C_WcCd
        .Row = Row

        Call OpenConWC(.Text, 1)
        Call SetActiveCell(frm1.vspdData, C_WcCd, Row,"M","X","X")
		Set gActiveElement = document.activeElement

    ElseIf Row > 0 And Col = C_ItemPopup Then
		.Col = C_ItemCd
        .Row = Row

		Call OpenItemInfo(.Text, 1)
        Call SetActiveCell(frm1.vspdData, C_ItemCd, Row,"M","X","X")
		Set gActiveElement = document.activeElement
    End If

    End With
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange
'   Event Desc :Combo Change Event
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case C_UseFlg
				Col = Col
				intIndex = .Value
		End Select
    End With
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
     Call ggoSpread.ReOrderingSpreadData
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

    Dim strCboCd
    Dim strCboNm
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    strCboCd = "" & vbTab
    strCboNm = "" & vbTab

    '****************************
    'MileStone Flag Setting
    '****************************

    strCboCd = ""
    strCboCd = "Y" & vbTab & "N"

    ggoSpread.SetCombo strCboCd, C_UseFlg

    '****************************
    'Insp Flag Setting
    '****************************

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )
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
	'----------  Coding part  -------------------------------------------------------------
    End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData ,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if

End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
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
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function ********************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False                                                        '⊙: Processing is NG

    Err.Clear                                                               '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                   '⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If

	If frm1.txtWCCd.value = "" Then
		frm1.txtWCNm.value = ""
	End If

	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then											'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Exit Function																'☜: Query db data
	End If

    FncQuery = True																'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()
    Dim IntRetCD
    Dim slPlantCd
    Dim slPlantNm

    FncNew = False                                                          '⊙: Processing is NG

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------

    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field

    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables

	frm1.txtPlantCd.focus
    Set gActiveElement = document.activeElement

    FncNew = True                                                           '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                      '⊙: Processing is NG

    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '☆:
        Exit Function
    End If

    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------
		Exit Function
	End If

    If DbDelete = False Then
		Exit Function
    End If     						'☜: Delete db data

    FncDelete = True                                                        '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave()
    Dim IntRetCD
	Dim starDate
	Dim finaDate

    FncSave = False                                                         '⊙: Processing is NG

    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing

    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If

    With frm1
		For lRow = 1 To .vspdData.MaxRows

			.vspdData.Row = lRow
			.vspdData.Col= C_FrDt
			starDate =  Trim(.vspdData.Text)

			.vspdData.Col= C_EndDt
			finaDate =  Trim(.vspdData.Text)

			If UNIConvDate(finaDate) < UNIConvDate(starDate) Then
				Call DisplayMsgBox("972002", "X", "적용종료일", "적용시작일")
				Call SheetFocus(lRow, C_EndDt)
				Exit Function
			End If
		Next
	End With

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		LayerShowHide(0)
		Exit Function
    End If     				                                                  '☜: Save db data

    FncSave = True                                                          '⊙: Processing is OK
End Function

Function FncCopy()

    If frm1.vspdData.maxrows < 1 Then Exit Function

    frm1.vspdData.focus
    Set gActiveElement = document.activeElement
	frm1.vspdData.EditMode = True

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData

    ggoSpread.CopyRow

	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow, "N"

    With frm1
		.vspdData.Col = C_FrDt
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Text = frm1.txtBaseDt.Text

		.vspdData.Col = C_EndDt
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Text = frm1.txtBaseDt.Text

		frm1.vspdData.ReDraw = True

    End With
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

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.EditUndo                                                  '☜: Protect system from crashing

End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '☜: Clear error status

    FncInsertRow = False                                                         '☜: Processing is NG

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		FncInsertRow = False
		Exit Function
	End If

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If

    With frm1

        .vspdData.ReDraw = False
		.vspdData.Focus

		ggoSpread.Source = .vspdData

        If frm1.vspdData.selBlockRow = -1 Then
            ggoSpread.InsertRow 0, iIntReqRows
        Else
            ggoSpread.InsertRow , iIntReqRows
        End If

	    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + iIntReqRows - 1,C_CurCd,C_SubconPrc, "C" ,"I","X","X")

		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1, "Y")
		For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + iIntReqRows - 1
			.vspdData.Col = C_FrDt
			.vspdData.Text = Trim(frm1.txtBaseDt.Text)
			.vspdData.Col = C_EndDt
			.vspdData.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")	'Trim(frm1.txtBaseDt.Text)
			.vspdData.Col = C_UseFlg
			.vspdData.Text = 1
		Next

		Call ProtectMilestone(0)

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
    Dim iIntCnt

    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData
    lDelRows = ggoSpread.DeleteRow

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()															'☜: Protect system from crashing
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev()
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext()
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_MULTI)											'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)                                     '☜:화면 유형, Tab 유형 
End Function

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

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc :
'========================================================================================
Function FncSplitColumn()

    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit

    If gMouseClickStatus = "SPCRP" Then
       iColumnLimit  = 14

       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
          Exit Function
       End If

       Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE

       ggoSpread.Source = Frm1.vspdData

       ggoSpread.SSSetSplit(ACol)

       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow

       Frm1.vspdData.Action = 0

       Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If

End Function

'========================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  ******************************
'	설명 :
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery()
    Dim LngLastRow
    Dim LngMaxRow
    Dim LngRow
    Dim strTemp
    Dim StrNextKey

    DbQuery = False

    Call LayerShowHide(1)

    Err.Clear

	Dim strVal

    With frm1

    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&txtBaseDt=" & Trim(.hBaseDt.value)
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtBaseDt=" & Trim(.txtBaseDt.Text)
		strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	End If

    Call RunMyBizASP(MyBizASP, strVal)

    End With

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------

    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call SetToolbar("11101111001011")										'⊙: 버튼 툴바 제어 

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	Dim strVal
	Dim strPlantCd
	Dim strWcCd
	Dim strItemCd
	Dim strFrDt

    DbDelete = False														'⊙: Processing is NG

    LayerShowHide(1)

    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)			'☜: 삭제 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbDelete = True                                                         '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()
    Dim IntRows
    Dim strVal
	Dim strDel
	Dim strInsideFlg, strMilestoneFlg, strInspFlg, strValidFromDt, strValidToDt, strQueueTime, strSetupTime, strWaitTime, strFixRunTime
	Dim strRunTime, strMoveTime, strOverLapOpr, strSubconPrc
	Dim iColSep, iRowSep

    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규]
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

'	Dim iFormLimitByte						'102399byte
'
'	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

   	Dim TmpBufferVal, TmpBufferDel
   	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt

    DbSave = False                                                          '⊙: Processing is NG

    LayerShowHide(1)

    'On Error Resume Next                                                   '☜: Protect system from crashing

    With frm1
		.txtMode.Value			= parent.UID_M0002							'☜: 저장 상태 
		.txtFlgMode.Value		= lgIntFlgMode								'☜: 신규입력/수정 상태 
		.txtUpdtUserId.value	= parent.gUsrID
		.txtInsrtUserId.value	= parent.gUsrID
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep

    lGrpCnt = 1
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)

	With frm1.vspdData

    For IntRows = 1 To .MaxRows

		.Row = IntRows
		.Col = 0

		Select Case .Text

		    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

				strVal = ""

				If .Text = ggoSpread.InsertFlag Then
					strVal = strVal & "C" & iColSep & IntRows & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
				Else
					strVal = strVal & "U" & iColSep	& IntRows & iColSep				'⊙: U=Update
				End If

		        .Col = C_WcCd								'2
				strVal = strVal & Trim(.Text) & iColSep

		        .Col = C_ItemCd								'3
		        strVal = strVal & Trim(.Text) & iColSep

		        .Col = C_StApply							'4
		        strVal = strVal & ConvToSec(Trim(.Text)) & iColSep

		        .Col = C_StFixed							'5
		        strVal = strVal & ConvToSec(Trim(.Text)) & iColSep

    			.Col = C_AstFixed							'6
		        strVal = strVal & ConvToSec(Trim(.Text)) & iColSep

		        .Col = C_FrDt								'7
		        strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep

		        .Col = C_EndDt								'8
		        strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep

				.Col = C_UseFlg								'9
				strVal = strVal & Trim(.Text) & iRowSep

				ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = strVal
                iValCnt = iValCnt + 1                                                                                
                lGrpCnt = lGrpCnt + 1

		    Case ggoSpread.DeleteFlag

				strDel = ""

				strDel = strDel & "D" & iColSep	& IntRows & iColSep				'⊙: D=Delete

				.Col = C_WcCd		'2
				strDel = strDel & Trim(.Text) & iColSep

				.Col = C_ItemCd		'3
				strDel = strDel & Trim(.Text) & iColSep

				.Col = C_FrDt		'4
				strDel = strDel & Trim(.Text) & iRowSep

                ReDim Preserve TmpBufferDel(iDelCnt)
                
                TmpBufferDel(iDelCnt) = strDel
                iDelCnt = iDelCnt + 1 
                lGrpCnt = lGrpCnt + 1

		End Select
    Next
	End With

	iTotalStrVal = Join(TmpBufferVal, "")
	iTotalStrDel = Join(TmpBufferDel, "")
	
	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = iTotalStrDel & iTotalStrVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'☜: 비지니스 ASP 를 가동 

    DbSave = True                                                           '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
	IsOpenPop = False
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	With frm1.vspdData
		.focus
		.Row = lRow
		.Col = lCol
		.Action = 0
		.SelStart = 0
		.SelLength = len(.Text)
	End With
End Function


'==============================================================================
' Function : ConvToTimeFormat
' Description : 시간 형식으로 변경 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec

	Dim iVal2

	iVal2 = Fix(iVal)

	If iVal2 = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)

		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)

	End If
End Function

'==============================================================================
' Function : ConvToSec()
' Description : 저장시에 각 시간 데이터들을 초로 환산 
'==============================================================================
Function ConvToSec(ByVal Str)

	If Str = "" Then
		ConvToSec = 0
	ElseIf Len(Trim(Str)) = 8 Then
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
	Else
		ConvToSec = -999999
	End If

End Function