
'==========================================================================================================
Const BIZ_PGM_QRY_ID = "p4612mb1.asp"									'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "p4612mb2.asp"									'☆: Save 비지니스 로직 ASP명 

'=========================================================================================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

Dim C_Select
Dim C_ProdOrderNo
Dim C_ItemCode
Dim C_ItemName
Dim C_Spec
Dim C_ProdtOrderQty
Dim C_ProdtOrderUnit
Dim C_RemainQty
Dim C_PlanStartDt
Dim C_PlanComptDt
Dim C_SchdStartDt
Dim C_SchdComptDt
Dim C_RoutingNo
Dim C_TrackingNo
Dim C_ProdQtyInOrderUnit
Dim C_GoodQtyInOrderUnit
Dim C_BadQtyInOrderUnit
Dim C_InspGoodQtyInOrderUnit
Dim C_InspBadQtyInOrderUnit
Dim C_RcptQtyInOrderUnit
Dim C_UnRcptQtyInOrderUnit
Dim C_BaseUnit
Dim C_OrderQtyInBaseUnit
Dim C_ProdQtyInBaseUnit
Dim C_GoodQtyInBaseUnit
Dim C_BadQtyInBaseUnit
Dim C_InspGoodQtyInBaseUnit
Dim C_InspBadQtyInBaseUnit
Dim C_RcptQtyInBaseUnit
Dim C_UnRcptQtyInBaseUnit
Dim C_ReleaseDt
Dim C_RealStartDt
Dim C_RealComptDt
Dim C_OrderStatus
Dim C_OrderType
Dim C_Prod_Mgr
Dim C_ItemGroupCd
Dim C_ItemGroupNm


'=========================================================================================================
'	Insert Your Code for Global Variables Assign
'=========================================================================================================
Dim ihGridCnt						'hidden Grid Row Count
Dim intItemCnt						'hidden Grid Row Count
Dim IsOpenPop						'Popup
Dim gSelframeFlg

Dim lgButtonSelection

'=========================================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0					'initializes Group View Size
    lgStrPrevKey = ""					'initializes Previous Key
    lgLngCurRows = 0					'initializes Deleted Rows Count
	lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"	
End Sub
'=========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
    frm1.txtProdFromDt.text = StartDate
    frm1.txtProdToDt.text   = EndDate
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
End Sub


'==========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
    With frm1.vspdData
	.ReDraw = false
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

	.MaxCols = C_ItemGroupNm + 1
	.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	ggoSpread.SSSetCheck	C_Select,			"", 2,,,1
	ggoSpread.SSSetEdit		C_ProdOrderNo,		"오더번호", 18
	ggoSpread.SSSetEdit		C_ItemCode,			"품목",		18
	ggoSpread.SSSetEdit		C_ItemName,			"품목명",	25
	ggoSpread.SSSetEdit		C_Spec,				"규격",		25
	ggoSpread.SSSetFloat	C_ProdtOrderQty,	"오더수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_ProdtOrderUnit,	"오더단위", 8
    ggoSpread.SSSetFloat	C_RemainQty,		"오더잔량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetDate		C_PlanStartDt,		"착수예정일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_PlanComptDt,		"완료예정일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate 	C_SchdStartDt,		"착수계획일정", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate 	C_SchdComptDt,		"완료계획일정", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_RoutingNo,		"라우팅",	10
	ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "실적수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit, "양품수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit, "입고수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetDate		C_ReleaseDt,		"작업지시일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_RealStartDt,		"실착수일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_RealComptDt,		"실완료일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No.", 25
	ggoSpread.SSSetEdit		C_OrderStatus,		"지시상태", 12
	ggoSpread.SSSetEdit		C_OrderType,		"지시구분", 12	
	ggoSpread.SSSetEdit		C_Prod_Mgr,			"생산담당자", 10
	ggoSpread.SSSetEdit 	C_ItemGroupCd,		"품목그룹",	15
	ggoSpread.SSSetEdit		C_ItemGroupNm,		"품목그룹명", 30
	
	'hidden below-------
	ggoSpread.SSSetFloat	C_BadQtyInOrderUnit, "", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_InspGoodQtyInOrderUnit, "", 10
	ggoSpread.SSSetEdit		C_InspBadQtyInOrderUnit, "", 10
	ggoSpread.SSSetFloat	C_UnRcptQtyInOrderUnit, "", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_BaseUnit, "", 10
	ggoSpread.SSSetEdit		C_OrderQtyInBaseUnit, "", 10
	ggoSpread.SSSetEdit		C_ProdQtyInBaseUnit, "", 10
	ggoSpread.SSSetEdit		C_GoodQtyInBaseUnit, "", 10
	ggoSpread.SSSetFloat	C_BadQtyInBaseUnit, "", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_InspGoodQtyInBaseUnit, "", 10
	ggoSpread.SSSetEdit		C_InspBadQtyInBaseUnit, "", 10
	ggoSpread.SSSetEdit		C_RcptQtyInBaseUnit, "", 10
	ggoSpread.SSSetFloat	C_UnRcptQtyInBaseUnit, "", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	
  	Call ggoSpread.SSSetColHidden(C_OrderStatus, C_OrderStatus, True)
'	Call ggoSpread.SSSetColHidden(C_OrderType, C_OrderType, True)
	Call ggoSpread.SSSetColHidden(C_BadQtyInOrderUnit, C_UnRcptQtyInBaseUnit, True)
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	ggoSpread.SSSetSplit2(3)											'frozen 기능 추가 
	
	.ReDraw = True

	Call SetSpreadLock()

    End With
    
End Sub

'==========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock()

    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_ProdOrderNo, -1
		ggoSpread.SpreadLock C_ItemCode, -1
		ggoSpread.Spreadlock C_ItemName, -1
		ggoSpread.Spreadlock C_ProdtOrderQty, -1
		ggoSpread.SpreadLock C_ProdtOrderUnit, -1
		ggoSpread.Spreadlock C_RemainQty, -1
		ggoSpread.Spreadlock C_PlanStartDt, -1
		ggoSpread.Spreadlock C_PlanComptDt, -1
		ggoSpread.Spreadlock C_SchdStartDt, -1
		ggoSpread.Spreadlock C_SchdComptDt, -1
		ggoSpread.SpreadLock C_RoutingNo, -1
		ggoSpread.SpreadLock C_ProdQtyInOrderUnit, -1
		ggoSpread.SpreadLock C_RcptQtyInOrderUnit, -1
		ggoSpread.SpreadLock C_GoodQtyInOrderUnit, -1
		ggoSpread.SpreadLock C_ReleaseDt, -1
		ggoSpread.SpreadLock C_RealStartDt, -1
		ggoSpread.SpreadLock C_RealComptDt, -1
		ggoSpread.SpreadLock C_TrackingNo, -1
		ggoSpread.SpreadLock C_OrderStatus, -1
		ggoSpread.SpreadLock C_OrderType, -1
		ggoSpread.SpreadLock C_Prod_Mgr, -1
		ggoSpread.SpreadLock C_ItemGroupCd, -1
		ggoSpread.SpreadLock C_ItemGroupNm, -1
		.vspdData.ReDraw = True
    End With

End Sub

'==========================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData

	    .Redraw = False
	
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.SSSetProtected C_ProdOrderNo,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ItemCode,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ItemName,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ProdtOrderQty,		pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ProdtOrderUnit,		pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_RemainQty,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_PlanStartDt,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_PlanComptDt,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_SchdStartDt,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_SchdComptDt,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_RoutingNo,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ProdQtyInOrderUnit,	pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_RcptQtyInOrderUnit,	pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_GoodQtyInOrderUnit,	pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ReleaseDt,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_RealStartDt,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_RealComptDt,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_TrackingNo,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_OrderStatus,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_OrderType,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_Prod_Mgr,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ItemGroupCd,			pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_ItemGroupNm,			pvStartRow, pvEndRow
	    
	    .Col = 1
	    .Row = .ActiveRow
	    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
	    .EditMode = True
	    
	    .Redraw = True
    
    End With
    
End Sub


'==========================================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'==========================================================================================================
Sub InitSpreadPosVariables()

	C_Select					= 1
	C_ProdOrderNo				= 2
	C_ItemCode					= 3
	C_ItemName					= 4
	C_Spec						= 5
	C_ProdtOrderQty				= 6
	C_ProdtOrderUnit			= 7
	C_RemainQty					= 8
	C_PlanStartDt				= 9
	C_PlanComptDt				= 10
	C_SchdStartDt				= 11
	C_SchdComptDt				= 12
	C_RoutingNo					= 13
	C_ProdQtyInOrderUnit		= 14
	C_GoodQtyInOrderUnit		= 15
	C_RcptQtyInOrderUnit		= 16
	C_ReleaseDt					= 17
	C_RealStartDt				= 18
	C_RealComptDt				= 19
	C_TrackingNo				= 20
	C_OrderStatus				= 21
	C_OrderType					= 22
	C_BadQtyInOrderUnit			= 23
	C_InspGoodQtyInOrderUnit	= 24
	C_InspBadQtyInOrderUnit		= 25
	C_UnRcptQtyInOrderUnit		= 26
	C_BaseUnit					= 27
	C_OrderQtyInBaseUnit		= 28
	C_ProdQtyInBaseUnit			= 29
	C_GoodQtyInBaseUnit			= 30
	C_BadQtyInBaseUnit			= 31
	C_InspGoodQtyInBaseUnit		= 32
	C_InspBadQtyInBaseUnit		= 33
	C_RcptQtyInBaseUnit			= 34
	C_UnRcptQtyInBaseUnit		= 35
	C_Prod_Mgr					= 36
	C_ItemGroupCd				= 37
	C_ItemGroupNm				= 38
 
End Sub
 
'==========================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
  	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
  	Case "A"
 		ggoSpread.Source = frm1.vspdData 
  		
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_Select					= iCurColumnPos(1)
		C_ProdOrderNo				= iCurColumnPos(2)
		C_ItemCode					= iCurColumnPos(3)
		C_ItemName					= iCurColumnPos(4)
		C_Spec						= iCurColumnPos(5)
		C_ProdtOrderQty				= iCurColumnPos(6)
		C_ProdtOrderUnit			= iCurColumnPos(7)
		C_RemainQty					= iCurColumnPos(8)
		C_PlanStartDt				= iCurColumnPos(9)
		C_PlanComptDt				= iCurColumnPos(10)
		C_SchdStartDt				= iCurColumnPos(11)
		C_SchdComptDt				= iCurColumnPos(12)
		C_RoutingNo					= iCurColumnPos(13)
		C_ProdQtyInOrderUnit		= iCurColumnPos(14)
		C_GoodQtyInOrderUnit		= iCurColumnPos(15)
		C_RcptQtyInOrderUnit		= iCurColumnPos(16)
		C_ReleaseDt					= iCurColumnPos(17)
		C_RealStartDt				= iCurColumnPos(18)
		C_RealComptDt				= iCurColumnPos(19)
		C_TrackingNo				= iCurColumnPos(20)
		C_OrderStatus				= iCurColumnPos(21)
		C_OrderType					= iCurColumnPos(22)
		C_BadQtyInOrderUnit			= iCurColumnPos(23)
		C_InspGoodQtyInOrderUnit	= iCurColumnPos(24)
		C_InspBadQtyInOrderUnit		= iCurColumnPos(25)
		C_UnRcptQtyInOrderUnit		= iCurColumnPos(26)
		C_BaseUnit					= iCurColumnPos(27)
		C_OrderQtyInBaseUnit		= iCurColumnPos(28)
		C_ProdQtyInBaseUnit			= iCurColumnPos(29)
		C_GoodQtyInBaseUnit			= iCurColumnPos(30)
		C_BadQtyInBaseUnit			= iCurColumnPos(31)
		C_InspGoodQtyInBaseUnit		= iCurColumnPos(32)
		C_InspBadQtyInBaseUnit		= iCurColumnPos(33)
		C_RcptQtyInBaseUnit			= iCurColumnPos(34)
		C_UnRcptQtyInBaseUnit		= iCurColumnPos(35) 
		C_Prod_Mgr					= iCurColumnPos(36) 
		C_ItemGroupCd				= iCurColumnPos(37) 
		C_ItemGroupNm				= iCurColumnPos(38)  		
		
  	End Select
  
End Sub
 
'==========================================================================================================
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'==========================================================================================================
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "공장팝업"					' 팝업 명칭 
	arrParam(1) = "B_PLANT"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "공장"						' TextBox 명칭 
	
	arrField(0) = "PLANT_CD"						' Field명(0)
	arrField(1) = "PLANT_NM"						' Field명(1)
	
	arrHeader(0) = "공장"						' Header명(0)
	arrHeader(1) = "공장명"						' Header명(1)
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1) 		
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'==========================================================================================================
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'==========================================================================================================
Function OpenItemInfo(Byval strCode)
	
	Dim arrRet
	Dim iCalledAspName
	Dim arrParam(5), arrField(6)

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
	arrParam(1) = strCode						' Item Code
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
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)	
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
	
End Function

'==========================================================================================================
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'==========================================================================================================
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdToDt.Text
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = Trim(frm1.cboOrderType.value)
	
	iCalledAspName = AskPRAspName("p4111pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtProdOrderNo.Value    = arrRet(0)
		frm1.txtTrackingNo.focus 
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function

'==========================================================================================================
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'==========================================================================================================
Function OpenTrackingInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(4)
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
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtProdFromDt.Text
	arrParam(4) = frm1.txtProdToDt.Text	
	
	iCalledAspName = AskPRAspName("p4600pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4600pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtTrackingNo.Value = arrRet(0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
			
End Function

'==========================================================================================================
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'==========================================================================================================
Function OpenPartRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	frm1.vspdData.Row =frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ProdOrderNo

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
	arrParam(1) = Trim(frm1.vspdData.Text)				'☜: 조회 조건 데이타 
	
	iCalledAspName = AskPRAspName("p4311ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4311ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'==========================================================================================================
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'==========================================================================================================
Function OpenOprRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	frm1.vspdData.Row =frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ProdOrderNo
                
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	arrParam(1) = Trim(frm1.vspdData.Text)			'☜: 조회 조건 데이타 

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4111ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==========================================================================================================
'	Name : OpenProdRef()
'	Description : Production Reference
'==========================================================================================================
Function OpenProdRef()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	Trim(frm1.txtProdOrderNo.value)					'☜: 조회 조건 데이타 
	
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4411ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4411ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==========================================================================================================
'	Name : OpenRcptRef()
'	Description : Receipt Reference PopUp
'==========================================================================================================
Function OpenRcptRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	Trim(frm1.txtProdOrderNo.value)					'☜: 조회 조건 데이타 
	
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4511ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4511ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==========================================================================================================
'	Name : OpenConsumRef()
'	Description : Consumption Reference PopUp
'==========================================================================================================
Function OpenConsumRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	Trim(frm1.txtProdOrderNo.value)					'☜: 조회 조건 데이타 
	
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4412ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4412ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function


Function btnAutoSel_onClick()

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "전체선택"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "전체선택취소"
	End If

	Dim index,Count
	Dim strFlag
	
	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Select
		
		strFlag = frm1.vspdData.Value
		
		If lgButtonSelection = "SELECT" Then
			frm1.vspdData.Value = 1
			frm1.vspdData.Col = 0 
			ggoSpread.UpdateRow Index
		Else
			frm1.vspdData.Value = 0
			frm1.vspdData.Col = 0 
			'ggoSpread.SSDeleteFlag Index
			frm1.vspdData.Text=""
		End if

	Next 
	
	frm1.vspdData.ReDraw = true

End Function

'==========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData
		.Row = Row
		.Col = C_Select
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
	End With
End Sub

'==========================================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'==========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
  	
  	Call SetPopupMenuItemInf("0000111111")			'화면별 설정 
  	
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
  			ggoSpread.SSSort Col, lgSortKey			'Sort in Descending
  			lgSortKey = 1
  		End If
 		
  	End If
	
	'------ Developer Coding part (Start)
		With frm1
		'----------------------
		'Column Split
		'----------------------
		gMouseClickStatus = "SPC"

		.vspdData.Row = .vspdData.ActiveRow

		' 오더단위 
		.vspddata.Col = C_ProdtOrderUnit
		.txtOrderUnit.Value = .vspdData.Text
		' 오더수량 
		.vspddata.Col = C_ProdtOrderQty
		.txtOrderQty.Value = .vspdData.Text
		' 총생산량 
		.vspddata.Col = C_ProdQtyInOrderUnit
		.txtProdQty.Value = .vspdData.Text
		' 양품수량 
		.vspddata.Col = C_GoodQtyInOrderUnit
		.txtGoodQty.Value = .vspdData.Text
		' 불량수량 
		.vspddata.Col = C_BadQtyInOrderUnit
		.txtBadQty.Value = .vspdData.Text
		' 입고수량 
		.vspddata.Col = C_RcptQtyInOrderUnit
		.txtRcptQty.Value = .vspdData.Text
		' 입고대기수량 
		.vspddata.Col = C_UnRcptQtyInOrderUnit
		.txtUnRcptQty.Value = .vspdData.Text
	
		' 기준단위 
		.vspddata.Col = C_BaseUnit
		.txtBaseUnit.Value = .vspdData.Text
		' 오더수량 
		.vspddata.Col = C_OrderQtyInBaseUnit
		.txtOrderQty1.Value = .vspdData.Text
		' 총생산량 
		.vspddata.Col = C_ProdQtyInBaseUnit
		.txtProdQty1.Value = .vspdData.Text
		' 양품수량 
		.vspddata.Col = C_GoodQtyInBaseUnit
		.txtGoodQty1.Value = .vspdData.Text
		' 불량수량 
		.vspddata.Col = C_BadQtyInBaseUnit
		.txtBadQty1.Value = .vspdData.Text
		' 입고수량 
		.vspddata.Col = C_RcptQtyInBaseUnit
		.txtRcptQty1.Value = .vspdData.Text
		' 입고대기수량 
		.vspddata.Col = C_UnRcptQtyInBaseUnit
		.txtUnRcptQty1.Value = .vspdData.Text
		
		' 착수예정일 
		.vspddata.Col = C_PlanStartDt
		.txtPlanStratDt.text = .vspdData.Text
		' 완료예정일 
		.vspddata.Col = C_PlanComptDt
		.txtPlanEndDt.Text	= .vspdData.Text
		' 착수시작일정 
		.vspddata.Col = C_SchdStartDt
		.txtPlannedStratDt.Text = .vspdData.Text
		' 착수완료일정 
		.vspddata.Col = C_SchdComptDt
		.txtPlannedEndDt.Text = .vspdData.Text
		' 작업지시일 
		.vspddata.Col = C_ReleaseDt
		.txtReleaseDt.Text	= .vspdData.Text
		' 실착수일 
		.vspddata.Col = C_RealStartDt
		.txtRealStratDt.Text = .vspdData.Text
		' 지시상태 
		.vspddata.Col = C_OrderStatus
		.txtOrderStatus.value = .vspdData.Text

		End With
 	 	'------ Developer Coding part (End)
 	
	
End Sub

'==========================================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'==========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
  
'==========================================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'==========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
 	If NewCol = C_Select or Col = C_Select Then
 		Cancel = True
 		Exit Sub
 	End If
 
     ggoSpread.Source = frm1.vspdData
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("A")
End Sub 
  
'==========================================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'==========================================================================================================
Sub PopSaveSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
     Call ggoSpread.SaveSpreadColumnInf()
End Sub 
  
'==========================================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'==========================================================================================================
Sub PopRestoreSpreadColumnInf()
      ggoSpread.Source = gActiveSpdSheet
     Call ggoSpread.RestoreSpreadInf()
     Call InitSpreadSheet
     Call ggoSpread.ReOrderingSpreadData
End Sub 
  
'==========================================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================================
Sub vspddata_DblClick(index , ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspddata(index)
End Sub

'==========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub

'==========================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'==========================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7       
        Call SetFocusToDocument("M")
		frm1.txtProdFromDt.Focus
    End If
End Sub

'==========================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'==========================================================================================================
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7      
        Call SetFocusToDocument("M")
		frm1.txtProdToDt.Focus
    End If
End Sub

'==========================================================================================================
'   Event Name : txtProdFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'==========================================================================================================
Sub txtProdFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================================
'   Event Name : txtProdToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'==========================================================================================================
Sub txtProdToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'==========================================================================================================
Function FncQuery() 

    Dim IntRetCD 
    
    FncQuery = False                                            '⊙: Processing is NG
    
    Err.Clear                                                   '☜: Protect system from crashing

	If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function

    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    Call InitVariables											'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then							'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If
       
    FncQuery = True												'⊙: Processing is OK
   
End Function

'==========================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'==========================================================================================================
Function FncSave() 

    Dim IntRetCD 
    
    FncSave = False                                             '⊙: Processing is NG
    
    Err.Clear                                                   '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then				'⊙: Check required field(Multi area)
       Exit Function
    End If    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then				                                    '☜: Save db data
		Exit Function
	End If
	
	FncSave = True                                              '⊙: Processing is OK
    
End Function

'==========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'==========================================================================================================
Function FncCopy() 
    ggoSpread.SpreadCopy
End Function

'==========================================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'==========================================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'==========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'==========================================================================================================
Function FncCancel()
	If frm1.vspdData.MaxRows < 1 Then Exit Function	 
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

'==========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'==========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 

    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows
    lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
    
    If lTempRows <= 16 Then
        Call DbQuery
    End If
End Function

'==========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'==========================================================================================================
Function FncPrint()													'☜: Protect system from crashing
    Call parent.FncPrint()
End Function

'==========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'==========================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)							 '☜: 화면 유형 
End Function

'==========================================================================================================
' Function Name : FncFind
' Function Desc : 
'==========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                          '☜:화면 유형, Tab 유무 
End Function

'==========================================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'==========================================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    If gMouseClickStatus = "SPCRP" Then
       iColumnLimit  = 33
       
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

'==========================================================================================================
' Function Name : FncExit
' Function Desc : 
'==========================================================================================================
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

'==========================================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'==========================================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'==========================================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'==========================================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
		ggoSpread.LoadLayout
    End If
End Function

'==========================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'==========================================================================================================
Function DbDelete() 

    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtPlanOrderNo=" & Trim(frm1.txtPlanOrderNo.value)		'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtProdStartDt=" & Trim(frm1.txtProdStartDt.value)		'☜: 삭제 조건 데이타 
	strVal = strVal & "&txtProdEndDt=" & Trim(frm1.txtProdEndDt.value)			'☜: 삭제 조건 데이타 
    strVal = strVal & "&cboOrderType=" & Trim(frm1.cboOrderType.value)			'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG

End Function

'==========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'==========================================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
	Call FncNew()
End Function

'==========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function DbQuery() 
    
    Err.Clear

    DbQuery = False
    
    Call LayerShowHide(1)
 
    Dim strVal
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtFromDt=" & Trim(frm1.hProdFromDt.value)
		strVal = strVal & "&txtToDt=" & Trim(frm1.hProdToDt.value)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.hOrderType.value)
		strVal = strVal & "&cboProdMgr=" & Trim(frm1.hProdMgr.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtFromDt=" & Trim(frm1.txtProdFromDt.text)
		strVal = strVal & "&txtToDt=" & Trim(frm1.txtProdToDt.text)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.cboOrderType.value)
		strVal = strVal & "&cboProdMgr=" & Trim(frm1.cboProdMgr.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
	End If

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'==========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'==========================================================================================================
Function DbQueryOk()
	
	Call SetToolbar("11001001000111")

    Call ggoOper.LockField(Document, "Q")

	If frm1.vspdData.MaxRows <= 0 Then Exit Function
    
	If lgIntFlgMode = parent.OPMD_CMODE Then		    
    
		With frm1

			.vspdData.Row = 1
			' 오더단위 
			.vspddata.Col = C_ProdtOrderUnit
			.txtOrderUnit.Value = .vspdData.Text
			' 오더수량 
			.vspddata.Col = C_ProdtOrderQty
			.txtOrderQty.Value = .vspdData.Text
			' 총생산량 
			.vspddata.Col = C_ProdQtyInOrderUnit
			.txtProdQty.Value = .vspdData.Text
			' 양품수량 
			.vspddata.Col = C_GoodQtyInOrderUnit
			.txtGoodQty.Value = .vspdData.Text
			' 불량수량 
			.vspddata.Col = C_BadQtyInOrderUnit
			.txtBadQty.Value = .vspdData.Text
			' 입고수량 
			.vspddata.Col = C_RcptQtyInOrderUnit
			.txtRcptQty.Value = .vspdData.Text
			' 입고대기수량 
			.vspddata.Col = C_UnRcptQtyInOrderUnit
			.txtUnRcptQty.Value = .vspdData.Text
	
			' 기준단위 
			.vspddata.Col = C_BaseUnit
			.txtBaseUnit.Value = .vspdData.Text
			' 오더수량 
			.vspddata.Col = C_OrderQtyInBaseUnit
			.txtOrderQty1.Value = .vspdData.Text
			' 총생산량 
			.vspddata.Col = C_ProdQtyInBaseUnit
			.txtProdQty1.Value = .vspdData.Text
			' 양품수량 
			.vspddata.Col = C_GoodQtyInBaseUnit
			.txtGoodQty1.Value = .vspdData.Text
			' 불량수량 
			.vspddata.Col = C_BadQtyInBaseUnit
			.txtBadQty1.Value = .vspdData.Text
			' 입고수량 
			.vspddata.Col = C_RcptQtyInBaseUnit
			.txtRcptQty1.Value = .vspdData.Text
			' 입고대기수량 
			.vspddata.Col = C_UnRcptQtyInBaseUnit
			.txtUnRcptQty1.Value = .vspdData.Text

			' 착수예정일 
			.vspddata.Col = C_PlanStartDt
			.txtPlanStratDt.text = .vspdData.Text
			' 완료예정일 
			.vspddata.Col = C_PlanComptDt
			.txtPlanEndDt.Text	= .vspdData.Text
			' 착수시작일정 
			.vspddata.Col = C_SchdStartDt
			.txtPlannedStratDt.Text = .vspdData.Text
			' 착수완료일정 
			.vspddata.Col = C_SchdComptDt
			.txtPlannedEndDt.Text = .vspdData.Text
			' 작업지시일 
			.vspddata.Col = C_ReleaseDt
			.txtReleaseDt.Text	= .vspdData.Text
			' 실착수일 
			.vspddata.Col = C_RealStartDt
			.txtRealStratDt.Text = .vspdData.Text
			' 지시상태 
			.vspddata.Col = C_OrderStatus
			.txtOrderStatus.value = .vspdData.Text
			
		End With   
		
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement 

	End If
	
	frm1.btnAutoSel.disabled = False
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE										'⊙: Indicates that current mode is Update mode
    
End Function

'==========================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'==========================================================================================================
Function DbSave() 
	
    Dim lRow           
    Dim strVal
    
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

    DbSave = False                                                          '⊙: Processing is NG
    
    Call LayerShowHide(1)
	
	frm1.txtMode.value = parent.UID_M0002									'☜: 저장 상태 
	frm1.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
	frm1.txtUpdtUserId.value = parent.gUsrID
	frm1.txtInsrtUserId.value  = parent.gUsrID
	
	'-----------------------
	'Data manipulate area
	'-----------------------
		
	iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'한번에 설정한 버퍼의 크기 설정 
	strCUTotalvalLen  = parent.C_CHUNK_ARRAY_COUNT	
    
	'102399byte
	iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
	'버퍼의 초기화 
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)						
	
	iTmpCUBufferCount = -1 
	
	strCUTotalvalLen = 0 
	
    With frm1
    
		For lRow = 1 To .vspdData.MaxRows
		   	.vspdData.Row = lRow
			.vspdData.Col = 0
			.vspdData.Col = C_Select
		   	If .vspdData.Value = 1 Then
				
		   		strVal = ""
				' Plant Code
				strVal = strVal & UCase(Trim(.txtPlantCd.value)) & iColSep
				.vspdData.Col = C_ProdOrderNo
				strVal = strVal & Trim(.vspdData.Text) & iColSep
				strVal = strVal & lRow & iRowSep
				
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
				
		   	End If
   		Next
	
		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)
		End If  
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)							'☜: 비지니스 ASP 를 가동 
	
    End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'==========================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'==========================================================================================================
Function DbSaveOk()															'☆: 저장 성공후 실행 로직 

	Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
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

'==========================================================================================================
'	SheetFocus Define
'==========================================================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function
