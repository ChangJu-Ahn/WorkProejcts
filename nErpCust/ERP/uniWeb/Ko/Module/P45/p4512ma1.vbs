'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p4512mb1.asp"						'☆: 비지니스 로직(Qeury) ASP명 
Const BIZ_PGM_SAVE_ID = "p4512mb2.asp"						'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Dim c_ProdtOrderNo			'= 1
Dim C_ItemCd				'= 2
Dim C_ItemNm				'= 3
Dim C_Spec
Dim C_RcptQty				'= 5
Dim C_Unit					'= 6
Dim C_PostDt				'= 7
Dim C_DocumentDt			'= 8
Dim c_WcCd					'= 9
Dim c_WcNm					'= 10
Dim c_SlCd					'= 11
Dim c_SlNm					'= 12
Dim c_MoveType				'= 13
Dim c_DocumentNo			'= 14
Dim C_OprNO					'= 15
Dim C_Seq					'= 16
Dim C_ReportType			'= 17
Dim C_Year					'= 18
Dim C_ProdtOrderUnit		'= 19
Dim C_ProdtOrderQty			'= 20
Dim C_ProdQtyInOrderUnit	'= 21
Dim C_GoodQtyInOrderUnit	'= 22
Dim C_RcptQtyInOrderUnit	'= 23
Dim C_BaseUnit				'= 24
Dim C_OrderQtyInBaseUnit	'= 25
Dim C_ProdQtyInBaseUnit		'= 26
Dim C_GoodQtyInBaseUnit		'= 27
Dim C_RcptQtyInBaseUnit		'= 28
Dim C_PlanStartDt			'= 29
Dim C_PlanComptDt			'= 30
Dim C_SchdStartDt			'= 31
Dim C_SchdComptDt			'= 32
Dim C_ReleaseDt				'= 33
Dim C_RealStartDt			'= 34
Dim C_RealComptDt			'= 35
Dim C_RcptQtyInOrdRslt		'= 36
Dim C_OrderStatus			'= 37
Dim C_LotNo					'= 38
Dim C_LotSubNo				'= 39
Dim C_TrackingNo			'= 40
Dim C_ItemGroupCd
Dim C_ItemGroupNm
Dim C_lot_no


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgSortKey
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  --------------------------------------------------------------
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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
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
    frm1.txtFromDt.text = StartDate
    frm1.txtToDt.text   = EndDate
End Sub


'=============================================== 2.2.3 InitSpreadSheet() =================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
    With frm1.vspdData
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

	.ReDraw = false

	.MaxCols = C_lot_no + 1
	.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	ggoSpread.SSSetEdit		c_ProdtOrderNo, "오더번호", 18
	ggoSpread.SSSetEdit		C_ItemCd, "품목", 18
	ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25
	ggoSpread.SSSetEdit		C_Spec,	"규격", 25
	ggoSpread.SSSetFloat	C_RcptQty, "입고수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_Unit,	"기준단위", 8
	ggoSpread.SSSetDate		C_PostDt, "입고일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_DocumentDt, "전표발생일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		c_WcCd, "작업장", 10
	ggoSpread.SSSetEdit		c_WcNm, "작업장명", 20
	ggoSpread.SSSetEdit		c_SlCd, "창고", 10
	ggoSpread.SSSetEdit		c_SlNm, "창고명", 20
	ggoSpread.SSSetEdit		c_MoveType,	"이동유형", 8
	ggoSpread.SSSetEdit		c_DocumentNo, "입고번호", 18
	
	ggoSpread.SSSetDate		C_PlanStartDt, "착수예정일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_PlanComptDt, "완료예정일", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_ProdtOrderUnit, "오더단위", 8
	ggoSpread.SSSetFloat	C_ProdtOrderQty, "오더수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "실적수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit, "양품수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit, "입고수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_BaseUnit, "기준단위", 8
	ggoSpread.SSSetFloat	C_OrderQtyInBaseUnit, "기준수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit, "실적수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit, "양품수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInBaseUnit, "입고수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInOrdRslt, "", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
	ggoSpread.SSSetDate		C_SchdStartDt, "", 10, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_SchdComptDt, "", 10, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_ReleaseDt, "",   10, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_RealStartDt, "", 10, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_RealComptDt, "", 10, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_OrderStatus, "", 8
	ggoSpread.SSSetEdit		C_OprNO, "", 6
	ggoSpread.SSSetEdit		C_Seq, "",   6
	ggoSpread.SSSetEdit		C_ReportType, "", 6
	ggoSpread.SSSetEdit		C_Year, "", 6
	ggoSpread.SSSetEdit		C_LotNo, "Lot No.", 12,,,12
	Call AppendNumberPlace("6", "3", "0")
	ggoSpread.SSSetFloat		C_LotSubNo, "순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
	ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 25
	ggoSpread.SSSetEdit 	C_ItemGroupCd, "품목그룹",	15
	ggoSpread.SSSetEdit		C_ItemGroupNm, "품목그룹명", 30
	ggoSpread.SSSetEdit		C_lot_no, "LOT NO", 20


	'Call ggoSpread.MakePairsColumn(,)
	Call ggoSpread.SSSetColHidden(C_PlanStartDt, C_Year, True)
	Call ggoSpread.SSSetColHidden(C_lot_no, C_lot_no, True)
	Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
	ggoSpread.SSSetSplit2(2)											'frozen 기능 추가 

	.ReDraw = true

	Call SetSpreadLock

    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
	ggoSpread.SpreadLock c_ProdtOrderNo, -1
	ggoSpread.SpreadLock C_ItemCd, -1
	ggoSpread.SpreadLock C_ItemNm, -1
	ggoSpread.SpreadLock C_RcptQty, -1
	ggoSpread.SpreadLock C_Unit, -1
	ggoSpread.SpreadLock C_PostDt, -1
	ggoSpread.SpreadLock C_DocumentDt, -1
	ggoSpread.SpreadLock c_WcCd, -1
	ggoSpread.SpreadLock c_WcNm, -1
	ggoSpread.SpreadLock c_SlCd, -1
	ggoSpread.SpreadLock c_SlNm, -1
	ggoSpread.SpreadLock c_MoveType, -1
	ggoSpread.SpreadLock c_DocumentNo, -1
	ggoSpread.SpreadLock C_LotNo, -1
	ggoSpread.SpreadLock C_LotSubNo, -1
	ggoSpread.SpreadLock C_TrackingNo, -1
	ggoSpread.SpreadLock C_ItemGroupCd, -1
	ggoSpread.SpreadLock C_ItemGroupNm, -1
	ggoSpread.SpreadLock C_lot_no, -1
	
	
	.vspdData.ReDraw = True
	
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SSSetProtected  c_ProdtOrderNo,	pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_ItemCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_ItemNm,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_RcptQty,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_Unit,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_PostDt,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_DocumentDt,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  c_WcCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  c_WcNm,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  c_SlCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  c_SlNm,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  c_MoveType,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  c_DocumentNo,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_LotNo,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_LotSubNo,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected  C_TrackingNo,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemGroupCd,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemGroupNm,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_lot_no,		pvStartRow, pvEndRow
	
    .vspdData.ReDraw = True
    
    End With
End Sub


'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()

	c_ProdtOrderNo			= 1
	C_ItemCd				= 2
	C_ItemNm				= 3
	C_Spec					= 4
	C_RcptQty				= 5
	C_Unit					= 6
	C_PostDt				= 7
	C_DocumentDt			= 8
	c_WcCd					= 9
	c_WcNm					= 10
	c_SlCd					= 11
	c_SlNm					= 12
	c_MoveType				= 13
	c_DocumentNo			= 14
	C_PlanStartDt			= 15
	C_PlanComptDt			= 16
	C_ProdtOrderUnit		= 17
	C_ProdtOrderQty			= 18
	C_ProdQtyInOrderUnit	= 19
	C_GoodQtyInOrderUnit	= 20
	C_RcptQtyInOrderUnit	= 21
	C_BaseUnit				= 22
	C_OrderQtyInBaseUnit	= 23
	C_ProdQtyInBaseUnit		= 24
	C_GoodQtyInBaseUnit		= 25
	C_RcptQtyInBaseUnit		= 26
	C_RcptQtyInOrdRslt		= 27
	C_SchdStartDt			= 28
	C_SchdComptDt			= 29
	C_ReleaseDt				= 30
	C_RealStartDt			= 31
	C_RealComptDt			= 32
	C_OrderStatus			= 33
	C_OprNO					= 34
	C_Seq					= 35
	C_ReportType			= 36
	C_Year					= 37
	C_LotNo					= 38
	C_LotSubNo				= 39
	C_TrackingNo			= 40
	C_ItemGroupCd			= 41
	C_ItemGroupNm			= 42
	C_lot_no     			= 43
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
  			
		c_ProdtOrderNo			= iCurColumnPos(1)
		C_ItemCd				= iCurColumnPos(2)
		C_ItemNm				= iCurColumnPos(3)
		C_Spec					= iCurColumnPos(4)
		C_RcptQty				= iCurColumnPos(5)
		C_Unit					= iCurColumnPos(6)
		C_PostDt				= iCurColumnPos(7)
		C_DocumentDt			= iCurColumnPos(8)
		c_WcCd					= iCurColumnPos(9)
		c_WcNm					= iCurColumnPos(10)
		c_SlCd					= iCurColumnPos(11)
		c_SlNm					= iCurColumnPos(12)
		c_MoveType				= iCurColumnPos(13)
		c_DocumentNo			= iCurColumnPos(14)
		C_PlanStartDt			= iCurColumnPos(15)
		C_PlanComptDt			= iCurColumnPos(16)
		C_ProdtOrderUnit		= iCurColumnPos(17)
		C_ProdtOrderQty			= iCurColumnPos(18)
		C_ProdQtyInOrderUnit	= iCurColumnPos(19)
		C_GoodQtyInOrderUnit	= iCurColumnPos(20)
		C_RcptQtyInOrderUnit	= iCurColumnPos(21)
		C_BaseUnit				= iCurColumnPos(22)
		C_OrderQtyInBaseUnit	= iCurColumnPos(23)
		C_ProdQtyInBaseUnit		= iCurColumnPos(24)
		C_GoodQtyInBaseUnit		= iCurColumnPos(25)
		C_RcptQtyInBaseUnit		= iCurColumnPos(26)
		C_RcptQtyInOrdRslt		= iCurColumnPos(27)
		C_SchdStartDt			= iCurColumnPos(28)
		C_SchdComptDt			= iCurColumnPos(29)
		C_ReleaseDt				= iCurColumnPos(30)
		C_RealStartDt			= iCurColumnPos(31)
		C_RealComptDt			= iCurColumnPos(32)
		C_OrderStatus			= iCurColumnPos(33)
		C_OprNO					= iCurColumnPos(34)
		C_Seq					= iCurColumnPos(35)
		C_ReportType			= iCurColumnPos(36)
		C_Year					= iCurColumnPos(37)
		C_LotNo					= iCurColumnPos(38)
		C_LotSubNo				= iCurColumnPos(39)
		C_TrackingNo			= iCurColumnPos(40)
		C_ItemGroupCd			= iCurColumnPos(41)
		C_ItemGroupNm			= iCurColumnPos(42)
		C_lot_no     			= iCurColumnPos(43)

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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"					' 팝업 명칭 
	arrParam(1) = "B_PLANT"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "공장"						' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"						' Field명(0)
    arrField(1) = "PLANT_NM"						' Field명(1)
    
    arrHeader(0) = "공장"						' Header명(0)
    arrHeader(1) = "공장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
		
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode)
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
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
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
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "ST"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = ""
	
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
		Call SetProdOrderNo(arrRet)
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

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval strCode, Byval iWhere)
    If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""	
	
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
		Call SetTrackingNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSLCd()

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

	arrParam(0) = "창고팝업"											' 팝업 명칭 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtSLCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "창고"												' TextBox 명칭 
	
    	arrField(0) = "SL_CD"												' Field명(0)
    	arrField(1) = "SL_NM"												' Field명(1)
    
    	arrHeader(0) = "창고"											' Header명(0)
    	arrHeader(1) = "창고명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtSLCd.focus
	
End Function

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenConWC()
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

	arrParam(0) = "작업장팝업"											' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
		
End Function

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

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	
   	With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
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

'------------------------------------------  OpenProdRef()  ----------------------------------------------
'	Name : OpenProdRef()
'	Description : Production Reference
'---------------------------------------------------------------------------------------------------------
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

    arrParam(0) = Trim(frm1.txtPlantCd.value)	'☆: 조회 조건 데이타 
	
	With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
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

'------------------------------------------  OpenRcptRef()  ----------------------------------------------
'	Name : OpenRcptRef()
'	Description : Receipt Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
   	With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
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

'------------------------------------------  OpenConsumRef()  --------------------------------------------
'	Name : OpenConsumRef()
'	Description : Consumption Reference PopUp
'---------------------------------------------------------------------------------------------------------
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
	
   	With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
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

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetConPlant()  -----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  ---------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetTrackingNo()  ----------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function

'------------------------------------------  SetSLCd()  ----------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetConWC()  ---------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
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

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
  	
  	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0101111111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	End If
	
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

	'------ Developer Coding part (Start)
		With frm1
			'----------------------
			'Column Split
			'----------------------
			.vspddata.Row = .vspdData.ActiveRow
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
			' 입고수량 
			.vspddata.Col = C_RcptQtyInOrderUnit
			.txtRcptQty.Value = .vspdData.Text
		
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
			' 입고수량 
			.vspddata.Col = C_RcptQtyInBaseUnit
			.txtRcptQty1.Value = .vspdData.Text
		
			' 착수예정일 
			.vspddata.Col = C_PlanStartDt
			.txtPlanStratDt.text = .vspdData.Text
			' 완료예정일 
			.vspddata.Col = C_PlanComptDt
			.txtPlanEndDt.Text	= .vspdData.Text
			' 착수시작일정 
'			.vspddata.Col = C_SchdStartDt				' remove from kjpark at 2003.04.07
'			.txtPlannedStratDt.Text = .vspdData.Text	' remove from kjpark at 2003.04.07
			' 착수완료일정 
'			.vspddata.Col = C_SchdComptDt				' remove from kjpark at 2003.04.07
'			.txtPlannedEndDt.Text = .vspdData.Text		' remove from kjpark at 2003.04.07
			' 작업지시일 
			.vspddata.Col = C_ReleaseDt
			.txtReleaseDt.Text	= .vspdData.Text
			' 실착수일 
			.vspddata.Col = C_RealStartDt
			.txtRealStratDt.Text = .vspdData.Text
			' 실완료일 
'			.vspddata.Col = C_RealComptDt				' remove from kjpark at 2003.04.07
'			.txtRealEndDt.Text	= .vspdData.Text		' remove from kjpark at 2003.04.07
			' 지시상태 
			.vspddata.Col = C_OrderStatus
			.txtOrderStatus.value = .vspdData.Text
		End With
		'------ Developer Coding part (End)

End Sub
'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
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


'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
 	'If NewCol = C_Select or Col = C_Select Then
 	'	Cancel = True
 	'	Exit Sub
 	'End If
 
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
     Call ggoSpread.ReOrderingSpreadData
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
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
		Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" and lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
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

    FncQuery = False                                            '⊙: Processing is NG

    Err.Clear                                                   '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
	If frm1.txtWCCd.value = "" Then
		frm1.txtWCNm.value = "" 
	End If
	
	If frm1.txtSlCd.value = "" Then
		frm1.txtSlNm.value = "" 
	End If	
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
	Call InitVariables
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
    
    FncSave = False                                           '⊙: Processing is NG
    
    Err.Clear                                                 '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
       Exit Function
    End If    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then				                                  '☜: Save db data
		Exit Function
	End If
	
    FncSave = True                                            '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next                                                '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
    On Error Resume Next                                                 '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    
	If frm1.vspdData.MaxRows < 1 Then Exit Function	     
    
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement
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
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
       iColumnLimit  = 37
       
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
	
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
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
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtSlCd=" & Trim(.hSlCd.value)
		strVal = strVal & "&txtFromDt=" & .hFromDt.value
		strVal = strVal & "&txtToDt=" & .hToDt.value
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		strVal = strVal & "&txtSlCd=" & Trim(.txtSlCd.value)
		strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)
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
	
	Call SetToolbar("11001011000111")										'⊙: 버튼 툴바 제어 

    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
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
			' 입고수량 
			.vspddata.Col = C_RcptQtyInOrderUnit
			.txtRcptQty.Value = .vspdData.Text
			
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
			' 입고수량 
			.vspddata.Col = C_RcptQtyInBaseUnit
			.txtRcptQty1.Value = .vspdData.Text
			
			' 착수예정일 
			.vspddata.Col = C_PlanStartDt
			.txtPlanStratDt.text = .vspdData.Text
			' 완료예정일 
			.vspddata.Col = C_PlanComptDt
			.txtPlanEndDt.Text	= .vspdData.Text
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

	End If

    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
	
End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is to execute transaction.
'========================================================================================
Function DbSave() 
    Dim lRow              
	Dim strDel
	
	Dim iColSep, iRowSep
    
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size

    DbSave = False                                                          '⊙: Processing is NG
    
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
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
		'102399byte
		iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
		'버퍼의 초기화 
		ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

		iTmpDBufferCount = -1
	
		strDTotalvalLen  = 0

    
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
			   
			If .vspdData.Text = ggoSpread.DeleteFlag Then
				
				strDel = ""
				strDel = strDel & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
		        .vspdData.Col = C_ProdtOrderNo			'2
		        strDel = strDel & Trim(.vspdData.Text) & iColSep
		        .vspdData.Col = C_OprNo					'3
		        strDel = strDel & Trim(.vspdData.Text) & iColSep
		        .vspdData.Col = C_DocumentDt			'6
		        strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & iColSep
		        .vspdData.Col = C_ReportType			'5
		        strDel = strDel & Trim(.vspdData.Text) & iColSep
		        .vspdData.Col = C_RcptQtyInOrdRslt		'8
		        strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		        .vspdData.Col = C_RcptQty				'7
		        strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
		        .vspdData.Col = C_Seq					'4
		        strDel = strDel & Cint(Trim(.vspdData.Text)) & iColSep
		        .vspdData.Col = C_DocumentNo			'10
		        strDel = strDel & Trim(.vspdData.Text) & iColSep
		        .vspdData.Col = C_Year					'9
		        strDel = strDel & Trim(.vspdData.Text) & iColSep
		        .vspdData.Col = C_SlCd					'11
		        strDel = strDel & Trim(.vspdData.Text) & iColSep
		        strDel = strDel & lRow & iRowSep

                '2008-05-26 11:26오후 :: hanc
    			.vspdData.Col = C_LOT_NO
'MsgBox Trim(.vspdData.Text)
'    			If Trim(.vspdData.Text) = "" OR  Trim(.vspdData.Text) = "*" OR  Trim(.vspdData.Text) = "1" Then
'    			Else
'    			    MsgBox "인터페이스 DATA는 삭제 할 수 없습니다"
'    			    Call LayerShowHide(0)
'    			    Exit Function
'    			End If
    			    		        
		        If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
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
			         
			    iTmpDBuffer(iTmpDBufferCount) =  strDel         
			    strDTotalvalLen = strDTotalvalLen + Len(strDel)


		    End If
		    
		Next
	
		If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name = "txtDSpread"
		   objTEXTAREA.value = Join(iTmpDBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
		End If
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)									'☜: 비지니스 ASP 를 가동 
		
	End With
	
    DbSave = True																	'⊙: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
    Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 

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
