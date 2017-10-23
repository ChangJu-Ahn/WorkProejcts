'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p4512mb1.asp"						'��: �����Ͻ� ����(Qeury) ASP�� 
Const BIZ_PGM_SAVE_ID = "p4512mb2.asp"						'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
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


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgSortKey
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  --------------------------------------------------------------
Dim IsOpenPop          
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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
	ggoSpread.SSSetEdit		c_ProdtOrderNo, "������ȣ", 18
	ggoSpread.SSSetEdit		C_ItemCd, "ǰ��", 18
	ggoSpread.SSSetEdit		C_ItemNm, "ǰ���", 25
	ggoSpread.SSSetEdit		C_Spec,	"�԰�", 25
	ggoSpread.SSSetFloat	C_RcptQty, "�԰����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_Unit,	"���ش���", 8
	ggoSpread.SSSetDate		C_PostDt, "�԰���", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_DocumentDt, "��ǥ�߻���", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		c_WcCd, "�۾���", 10
	ggoSpread.SSSetEdit		c_WcNm, "�۾����", 20
	ggoSpread.SSSetEdit		c_SlCd, "â��", 10
	ggoSpread.SSSetEdit		c_SlNm, "â���", 20
	ggoSpread.SSSetEdit		c_MoveType,	"�̵�����", 8
	ggoSpread.SSSetEdit		c_DocumentNo, "�԰��ȣ", 18
	
	ggoSpread.SSSetDate		C_PlanStartDt, "����������", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_PlanComptDt, "�ϷΌ����", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_ProdtOrderUnit, "��������", 8
	ggoSpread.SSSetFloat	C_ProdtOrderQty, "��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit, "��ǰ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit, "�԰����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_BaseUnit, "���ش���", 8
	ggoSpread.SSSetFloat	C_OrderQtyInBaseUnit, "���ؼ���", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit, "��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit, "��ǰ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInBaseUnit, "�԰����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
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
	ggoSpread.SSSetFloat		C_LotSubNo, "����", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
	ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 25
	ggoSpread.SSSetEdit 	C_ItemGroupCd, "ǰ��׷�",	15
	ggoSpread.SSSetEdit		C_ItemGroupNm, "ǰ��׷��", 30
	ggoSpread.SSSetEdit		C_lot_no, "LOT NO", 20


	'Call ggoSpread.MakePairsColumn(,)
	Call ggoSpread.SSSetColHidden(C_PlanStartDt, C_Year, True)
	Call ggoSpread.SSSetColHidden(C_lot_no, C_lot_no, True)
	Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
	ggoSpread.SSSetSplit2(2)											'frozen ��� �߰� 

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
 
'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
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

	arrParam(0) = "�����˾�"					' �˾� ��Ī 
	arrParam(1) = "B_PLANT"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"						' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"						' Field��(0)
    arrField(1) = "PLANT_NM"						' Field��(1)
    
    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)
    
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field��(0)
	arrField(1) = 2 '"ITEM_NM"					' Field��(1)
    
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
		Call DisplayMsgBox("971012","X", "����","X")
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

	arrParam(0) = "ǰ��׷��˾�"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "ǰ��׷�"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "ǰ��׷�"
	arrHeader(1) = "ǰ��׷��"
	    
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
		Call DisplayMsgBox("971012","X", "����","X")
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtSLCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
	
    	arrField(0) = "SL_CD"												' Field��(0)
    	arrField(1) = "SL_NM"												' Field��(1)
    
    	arrHeader(0) = "â��"											' Header��(0)
    	arrHeader(1) = "â���"											' Header��(1)
    
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

    arrParam(0) = Trim(frm1.txtPlantCd.value)	'��: ��ȸ ���� ����Ÿ 
	
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	
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
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetConPlant()  -----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  ---------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
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
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function

'------------------------------------------  SetSLCd()  ----------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetConWC()  ---------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  **************************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
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
		Call SetPopupMenuItemInf("0101111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
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
			' �������� 
			.vspddata.Col = C_ProdtOrderUnit
			.txtOrderUnit.Value = .vspdData.Text
			' �������� 
			.vspddata.Col = C_ProdtOrderQty
			.txtOrderQty.Value = .vspdData.Text
			' �ѻ��귮 
			.vspddata.Col = C_ProdQtyInOrderUnit
			.txtProdQty.Value = .vspdData.Text
			' ��ǰ���� 
			.vspddata.Col = C_GoodQtyInOrderUnit
			.txtGoodQty.Value = .vspdData.Text
			' �԰���� 
			.vspddata.Col = C_RcptQtyInOrderUnit
			.txtRcptQty.Value = .vspdData.Text
		
			' ���ش��� 
			.vspddata.Col = C_BaseUnit
			.txtBaseUnit.Value = .vspdData.Text
			' �������� 
			.vspddata.Col = C_OrderQtyInBaseUnit
			.txtOrderQty1.Value = .vspdData.Text
			' �ѻ��귮 
			.vspddata.Col = C_ProdQtyInBaseUnit
			.txtProdQty1.Value = .vspdData.Text
			' ��ǰ���� 
			.vspddata.Col = C_GoodQtyInBaseUnit
			.txtGoodQty1.Value = .vspdData.Text
			' �԰���� 
			.vspddata.Col = C_RcptQtyInBaseUnit
			.txtRcptQty1.Value = .vspdData.Text
		
			' ���������� 
			.vspddata.Col = C_PlanStartDt
			.txtPlanStratDt.text = .vspdData.Text
			' �ϷΌ���� 
			.vspddata.Col = C_PlanComptDt
			.txtPlanEndDt.Text	= .vspdData.Text
			' ������������ 
'			.vspddata.Col = C_SchdStartDt				' remove from kjpark at 2003.04.07
'			.txtPlannedStratDt.Text = .vspdData.Text	' remove from kjpark at 2003.04.07
			' �����Ϸ����� 
'			.vspddata.Col = C_SchdComptDt				' remove from kjpark at 2003.04.07
'			.txtPlannedEndDt.Text = .vspdData.Text		' remove from kjpark at 2003.04.07
			' �۾������� 
			.vspddata.Col = C_ReleaseDt
			.txtReleaseDt.Text	= .vspdData.Text
			' �������� 
			.vspddata.Col = C_RealStartDt
			.txtRealStratDt.Text = .vspdData.Text
			' �ǿϷ��� 
'			.vspddata.Col = C_RealComptDt				' remove from kjpark at 2003.04.07
'			.txtRealEndDt.Text	= .vspdData.Text		' remove from kjpark at 2003.04.07
			' ���û��� 
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
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 


'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
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
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
     Call ggoSpread.SaveSpreadColumnInf()
End Sub 
  
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
      ggoSpread.Source = gActiveSpdSheet
     Call ggoSpread.RestoreSpreadInf()
     Call InitSpreadSheet
     Call ggoSpread.ReOrderingSpreadData
End Sub 

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
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
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
		Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" and lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'#########################################################################################################


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function ********************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                            '��: Processing is NG

    Err.Clear                                                   '��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	'��: Display Message
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
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
	Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then							'��: This function check indispensable field
       Exit Function
    End If
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()	
		Exit Function
	End If

    FncQuery = True												'��: Processing is OK
    
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
    
    FncSave = False                                           '��: Processing is NG
    
    Err.Clear                                                 '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '��: Check required field(Multi area)
       Exit Function
    End If    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then				                                  '��: Save db data
		Exit Function
	End If
	
    FncSave = True                                            '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next                                                '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
    On Error Resume Next                                                 '��: Protect system from crashing
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
Function FncPrint()															'��: Protect system from crashing
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
	
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  ******************************
'	���� : 
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
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	Call SetToolbar("11001011000111")										'��: ��ư ���� ���� 

    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	
	If lgIntFlgMode = parent.OPMD_CMODE Then		
	
		With frm1

			.vspdData.Row = 1
			' �������� 
			.vspddata.Col = C_ProdtOrderUnit
			.txtOrderUnit.Value = .vspdData.Text
			' �������� 
			.vspddata.Col = C_ProdtOrderQty
			.txtOrderQty.Value = .vspdData.Text
			' �ѻ��귮 
			.vspddata.Col = C_ProdQtyInOrderUnit
			.txtProdQty.Value = .vspdData.Text
			' ��ǰ���� 
			.vspddata.Col = C_GoodQtyInOrderUnit
			.txtGoodQty.Value = .vspdData.Text
			' �԰���� 
			.vspddata.Col = C_RcptQtyInOrderUnit
			.txtRcptQty.Value = .vspdData.Text
			
			' ���ش��� 
			.vspddata.Col = C_BaseUnit
			.txtBaseUnit.Value = .vspdData.Text
			' �������� 
			.vspddata.Col = C_OrderQtyInBaseUnit
			.txtOrderQty1.Value = .vspdData.Text
			' �ѻ��귮 
			.vspddata.Col = C_ProdQtyInBaseUnit
			.txtProdQty1.Value = .vspdData.Text
			' ��ǰ���� 
			.vspddata.Col = C_GoodQtyInBaseUnit
			.txtGoodQty1.Value = .vspdData.Text
			' �԰���� 
			.vspddata.Col = C_RcptQtyInBaseUnit
			.txtRcptQty1.Value = .vspdData.Text
			
			' ���������� 
			.vspddata.Col = C_PlanStartDt
			.txtPlanStratDt.text = .vspdData.Text
			' �ϷΌ���� 
			.vspddata.Col = C_PlanComptDt
			.txtPlanEndDt.Text	= .vspdData.Text
			' �۾������� 
			.vspddata.Col = C_ReleaseDt
			.txtReleaseDt.Text	= .vspdData.Text
			' �������� 
			.vspddata.Col = C_RealStartDt
			.txtRealStratDt.Text = .vspdData.Text
			' ���û��� 
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
    
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
	
End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is to execute transaction.
'========================================================================================
Function DbSave() 
    Dim lRow              
	Dim strDel
	
	Dim iColSep, iRowSep
    
	Dim strDTotalvalLen						'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpDBuffer							'������ ���� [����] 
	Dim iTmpDBufferCount					'������ ���� Position
	Dim iTmpDBufferMaxCount					'������ ���� Chunk Size

    DbSave = False                                                          '��: Processing is NG
    
    Call LayerShowHide(1)
    
	With frm1
	
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------		
		iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
		'�ѹ��� ������ ������ ũ�� ���� 
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
		'102399byte
		iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
		'������ �ʱ�ȭ 
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

                '2008-05-26 11:26���� :: hanc
    			.vspdData.Col = C_LOT_NO
'MsgBox Trim(.vspdData.Text)
'    			If Trim(.vspdData.Text) = "" OR  Trim(.vspdData.Text) = "*" OR  Trim(.vspdData.Text) = "1" Then
'    			Else
'    			    MsgBox "�������̽� DATA�� ���� �� �� �����ϴ�"
'    			    Call LayerShowHide(0)
'    			    Exit Function
'    			End If
    			    		        
		        If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
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

			    If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
			       iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			       ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			    End If   
			         
			    iTmpDBuffer(iTmpDBufferCount) =  strDel         
			    strDTotalvalLen = strDTotalvalLen + Len(strDel)


		    End If
		    
		Next
	
		If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name = "txtDSpread"
		   objTEXTAREA.value = Join(iTmpDBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
		End If
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)									'��: �����Ͻ� ASP �� ���� 
		
	End With
	
    DbSave = True																	'��: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
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
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function
