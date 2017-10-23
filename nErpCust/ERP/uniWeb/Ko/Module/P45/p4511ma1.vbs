
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID	= "p4511mb1.asp"								'��: �����Ͻ� ����(Qeury) ASP�� 
Const BIZ_PGM_SAVE_ID	= "p4511mb2.asp"								'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Dim C_Select				
Dim C_ProdtOrderNo			
Dim C_ItemCd				
Dim C_ItemNm				
Dim C_Spec					
Dim C_ReportDt			
Dim C_ShiftCd				
Dim C_ReportType		
Dim C_ProdQty				
Dim C_ProdtOrderUnit
Dim C_RcptQty				
Dim C_BaseUnit			
Dim C_SlCd					
Dim C_SlCdPopup			
Dim C_LotReqFlg
Dim C_LotGenMthd
Dim C_LotNo					
Dim C_LotSubNo			
Dim C_OprNo					
Dim C_WcCd					
Dim C_Seq					
Dim C_PlanStartDt			
Dim C_PlanComptDt			
Dim C_ProdtOrderQty		
Dim C_ProdQtyInOrderUnit	
Dim C_GoodQtyInOrderUnit	
Dim C_RcptQtyInOrderUnit	
Dim C_OrderQtyInBaseUnit	
Dim C_ProdQtyInBaseUnit		
Dim C_GoodQtyInBaseUnit		
Dim C_RcptQtyInBaseUnit		
Dim C_SchdStartDt			
Dim C_SchdComptDt			
Dim C_ReleaseDt				
Dim C_RealStartDt			
Dim C_RealComptDt			
Dim C_OrderStatus			
Dim C_TrackingNo
Dim C_ItemGroupCd
Dim C_ItemGroupNm


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
Dim lgButtonSelection
Dim lgRedrewFlg
'++++++++++++++++  Insert Your Code for Global Variables Assign  +++++++++++++++++++++++++++++++++++++++++ 

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
    lgStrPrevKey2 = ""                          'initializes Previous Key2
    lgLngCurRows = 0                            'initializes Deleted Rows Count
   	lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "��ü����"
	    
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
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "��ü����"
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'================================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

    With frm1.vspdData
    .ReDraw = false

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030801", , Parent.gAllowDragDropSpread
	
	.MaxCols = C_ItemGroupNm + 1
	.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	ggoSpread.SSSetCheck	C_Select, "", 2,,,1
	ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 18
	ggoSpread.SSSetEdit		C_ItemCd, "ǰ��", 18
	ggoSpread.SSSetEdit		C_ItemNm, "ǰ���", 25
	ggoSpread.SSSetEdit		C_Spec, "�԰�", 25
	ggoSpread.SSSetDate		C_ReportDt, "������", 10, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_ShiftCd, "Shift", 8
	ggoSpread.SSSetEdit		C_ReportType, "��/��", 6
	ggoSpread.SSSetFloat	C_ProdQty, "���귮", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_ProdtOrderUnit, "��������", 8
	ggoSpread.SSSetFloat	C_RcptQty, "�԰�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_BaseUnit, "���ش���", 8
	ggoSpread.SSSetEdit		C_SlCd, "â��", 10,,,,2
	ggoSpread.SSSetButton	C_SlCdPopup
	ggoSpread.SSSetEdit		C_LotReqFlg, "", 10 'dummy
	ggoSpread.SSSetEdit		C_LotGenMthd, "Lot �ο����", 10
	ggoSpread.SSSetEdit		C_LotNo, "Lot No.", 20,,,25,2
	
	Call AppendNumberPlace("6", "3", "0")
	ggoSpread.SSSetFloat	C_LotSubNo, "����", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
	ggoSpread.SSSetEdit		C_OprNo, "����", 6
	ggoSpread.SSSetEdit		C_WcCd, "�۾���", 10
	ggoSpread.SSSetEdit		C_Seq, "����", 6
	ggoSpread.SSSetDate		C_PlanStartDt, "����������", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_PlanComptDt, "�ϷΌ����", 11, 2, parent.gDateFormat
	ggoSpread.SSSetFloat	C_ProdtOrderQty, "��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInOrderUnit, "��ǰ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInOrderUnit, "�԰����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_OrderQtyInBaseUnit, "���ؼ���", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInBaseUnit, "��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQtyInBaseUnit, "��ǰ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_RcptQtyInBaseUnit, "�԰����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetDate		C_SchdStartDt, "������ȹ����", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_SchdComptDt, "�Ϸ��ȹ����", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_ReleaseDt, "�۾�������", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_RealStartDt, "��������", 11, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_RealComptDt, "�ǿϷ���", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_OrderStatus, "���û���", 12
	ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 25
	ggoSpread.SSSetEdit 	C_ItemGroupCd, "ǰ��׷�",	15
	ggoSpread.SSSetEdit		C_ItemGroupNm, "ǰ��׷��", 30

	Call ggoSpread.MakePairsColumn(C_SlCd, C_SlCdPopup)
	Call ggoSpread.SSSetColHidden(C_Seq, C_OrderStatus , True)
	Call ggoSpread.SSSetColHidden(C_LotReqFlg, C_LotGenMthd , True)
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols , True)
	ggoSpread.SSSetSplit2(3)											'frozen ��� �߰� 
	
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
	ggoSpread.SpreadLock C_ProdtOrderNo, -1, C_ProdtOrderNo
	ggoSpread.SpreadLock C_ItemCd, -1, C_ItemCd
	ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
	ggoSpread.SpreadLock C_Spec, -1, C_Spec
	ggoSpread.SpreadLock C_ReportDt, -1, C_ReportDt    
	ggoSpread.SpreadLock C_ShiftCd, -1, C_ShiftCd
	ggoSpread.SpreadLock C_ReportType, -1, C_ReportType
	ggoSpread.SpreadLock C_ProdQty, -1, C_ProdQty	
	ggoSpread.SpreadLock C_ProdtOrderUnit, -1, C_ProdtOrderUnit
	ggoSpread.SpreadLock C_RcptQty, -1, C_RcptQty
	ggoSpread.SpreadLock C_BaseUnit, -1, C_BaseUnit
	ggoSpread.SpreadLock C_SlCd, -1, C_SlCd
	ggoSpread.SpreadLock C_SlCdPopup, -1, C_SlCdPopup
	ggoSpread.SpreadLock C_LotNo, -1, C_LotNo
	ggoSpread.SpreadLock C_LotSubNo, -1, C_LotSubNo
	ggoSpread.SpreadLock C_OprNo, -1, C_OprNo
	ggoSpread.SpreadLock C_WcCd, -1, C_WcCd
	ggoSpread.SpreadLock C_Seq, -1, C_Seq
	ggoSpread.SpreadLock C_PlanStartDt, -1,C_PlanStartDt
	ggoSpread.SpreadLock C_PlanComptDt, -1,C_PlanComptDt
	ggoSpread.SpreadLock C_TrackingNo, -1,C_TrackingNo
	ggoSpread.SpreadLock C_ItemGroupCd, -1,C_ItemGroupCd
	ggoSpread.SpreadLock C_ItemGroupNm, -1,C_ItemGroupNm
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
	ggoSpread.SSSetProtected C_ProdtOrderNo,	pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemNm,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Spec,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ReportDt,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ShiftCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ReportType,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ProdQty,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ProdtOrderUnit,	pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_RcptQty,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_BaseUnit,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_OprNo,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_WcCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_SlCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_SlCdPopup,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Seq,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PlanStartDt,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PlanComptDt,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_TrackingNo,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemGroupCd,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemGroupNm,		pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_Select				= 1
	C_ProdtOrderNo			= 2
	C_ItemCd				= 3
	C_ItemNm				= 4
	C_Spec					= 5
	C_ReportDt				= 6
	C_ShiftCd				= 7
	C_ReportType			= 8
	C_ProdQty				= 9
	C_ProdtOrderUnit		= 10
	C_RcptQty				= 11
	C_BaseUnit				= 12
	C_SlCd					= 13
	C_SlCdPopup				= 14
	C_LotReqFlg				= 15
	C_LotGenMthd			= 16
	C_LotNo					= 17
	C_LotSubNo				= 18
	C_OprNo					= 19
	C_WcCd					= 20
	C_Seq					= 21
	C_PlanStartDt			= 22
	C_PlanComptDt			= 23
	C_ProdtOrderQty			= 24
	C_ProdQtyInOrderUnit	= 25
	C_GoodQtyInOrderUnit	= 26
	C_RcptQtyInOrderUnit	= 27
	C_OrderQtyInBaseUnit	= 28
	C_ProdQtyInBaseUnit		= 29
	C_GoodQtyInBaseUnit		= 30
	C_RcptQtyInBaseUnit		= 31
	C_SchdStartDt			= 32
	C_SchdComptDt			= 33
	C_ReleaseDt				= 34
	C_RealStartDt			= 35
	C_RealComptDt			= 36
	C_OrderStatus			= 37
	C_TrackingNo			= 38
	C_ItemGroupCd			= 39
	C_ItemGroupNm			= 40
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
		C_Select				= iCurColumnPos(1)
		C_ProdtOrderNo			= iCurColumnPos(2)
		C_ItemCd				= iCurColumnPos(3)
		C_ItemNm				= iCurColumnPos(4)
		C_Spec					= iCurColumnPos(5)
		C_ReportDt				= iCurColumnPos(6)
		C_ShiftCd				= iCurColumnPos(7)
		C_ReportType			= iCurColumnPos(8)
		C_ProdQty				= iCurColumnPos(9)
		C_ProdtOrderUnit		= iCurColumnPos(10)
		C_RcptQty				= iCurColumnPos(11)
		C_BaseUnit				= iCurColumnPos(12)
		C_SlCd					= iCurColumnPos(13)
		C_SlCdPopup				= iCurColumnPos(14)
		C_LotReqFlg				= iCurColumnPos(15)
		C_LotGenMthd			= iCurColumnPos(16)
		C_LotNo					= iCurColumnPos(17)
		C_LotSubNo				= iCurColumnPos(18)
		C_OprNo					= iCurColumnPos(19)
		C_WcCd					= iCurColumnPos(20)
		C_Seq					= iCurColumnPos(21)
		C_PlanStartDt			= iCurColumnPos(22)
		C_PlanComptDt			= iCurColumnPos(23)
		C_ProdtOrderQty			= iCurColumnPos(24)
		C_ProdQtyInOrderUnit	= iCurColumnPos(25)
		C_GoodQtyInOrderUnit	= iCurColumnPos(26)
		C_RcptQtyInOrderUnit	= iCurColumnPos(27)
		C_OrderQtyInBaseUnit	= iCurColumnPos(28)
		C_ProdQtyInBaseUnit		= iCurColumnPos(29)
		C_GoodQtyInBaseUnit		= iCurColumnPos(30)
		C_RcptQtyInBaseUnit		= iCurColumnPos(31)
		C_SchdStartDt			= iCurColumnPos(32)
		C_SchdComptDt			= iCurColumnPos(33)
		C_ReleaseDt				= iCurColumnPos(34)
		C_RealStartDt			= iCurColumnPos(35)
		C_RealComptDt			= iCurColumnPos(36)
		C_OrderStatus			= iCurColumnPos(37)
		C_TrackingNo			= iCurColumnPos(38)
		C_ItemGroupCd			= iCurColumnPos(39)
		C_ItemGroupNm			= iCurColumnPos(40)
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

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"							' Field��(0)
    arrField(1) = "PLANT_NM"							' Field��(1)
    
    arrHeader(0) = "����"							' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
	
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

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'----------------------------------------------------------------------------------------------------------------
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

'--------------------------------------  OpenTrackingInfo()  ---------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 			' Where Condition
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
	
    arrField(0) = "SL_CD"													' Field��(0)
    arrField(1) = "SL_NM"													' Field��(1)
    
    arrHeader(0) = "â��"												' Header��(0)
    arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtSLCd.focus
	
End Function
'------------------------------------------  OpenSLCd2()  -------------------------------------------------
'	Name : OpenSLCd2()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd2(Byval strCode, Byval Row)

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
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
   	arrField(0) = "SL_CD"													' Field��(0)
   	arrField(1) = "SL_NM"													' Field��(1)
   	arrHeader(0) = "â��"												' Header��(0)
   	arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSLCd2(arrRet, Row)
	End If
	
End Function

'------------------------------------------  OpenOprRef()  -------------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'-----------------------------------------------------------------------------------------------------------
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

'------------------------------------------  OpenProdRef()  -------------------------------------------------
'	Name : OpenProdRef()
'	Description : Production Reference
'------------------------------------------------------------------------------------------------------------
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	
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

'------------------------------------------  OpenRcptRef()  -------------------------------------------------
'	Name : OpenRcptRef()
'	Description : Receipt Reference PopUp
'------------------------------------------------------------------------------------------------------------
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	
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

'------------------------------------------  SetProdOrderNo()  --------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetTrackingNo()  -------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function

'------------------------------------------  SetConWC()  -------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConWC(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetSLCd2()  --------------------------------------------------
'	Name : SetSLCd2()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd2(byval arrRet, Byval Row)

    With frm1
	   	.vspdData.Row = Row
	   	.vspdData.Col = C_SLCD
	   	.vspdData.Text = arrRet(0)	   	
	End With

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

Function btnAutoSel_onClick()

	lgRedrewFlg = False

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "��ü����"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "��ü�������"
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

	lgRedrewFlg = True

End Function

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *************************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
  	
  	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
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
 
 	If NewCol = C_Select or Col = C_Select Then
 		Cancel = True
 		Exit Sub
 	End If
 
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
			If Row < 1 Then Exit Sub
		Select Case Col
			Case C_Select
				If lgRedrewFlg = True Then .ReDraw = false
				.Row = Row
				.Col = C_Select
				
				If ButtonDown = 1 Then
					ggoSpread.UpdateRow Row
					ggoSpread.SpreadUnLock C_SlCd, Row , C_SlCd ,Row
					ggoSpread.SpreadUnLock C_SlCdPopup, Row , C_SlCdPopup ,Row
					ggoSpread.SSSetRequired  C_SlCd,			Row, Row
					
					.Col = C_LotReqFlg								'Lot ����ǰ Check!
					If Trim(.Text) = "Y" Then
						.Col = C_LotGenMthd
						If Trim(.Text) = "M" Then
							ggoSpread.SpreadUnLock C_LotNo, Row, C_LotSubNo, Row
							ggoSpread.SSSetRequired  C_LotNo,			Row, Row
						Else
							ggoSpread.SpreadUnLock C_LotNo, Row, C_LotSubNo, Row
						End If	
					End If
					
				Else
					If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
						Exit Sub
					End If
					
					.Col = C_LotNo
					.Text = ""
					.Col = C_LotSubNo
					.Text = 0
					ggoSpread.SSDeleteFlag Row,Row
					ggoSpread.SSSetProtected C_SlCd,			Row, Row
					ggoSpread.SSSetProtected C_SlCdPopup,			Row, Row
					ggoSpread.SSSetProtected C_LotNo, Row, Row
					ggoSpread.SpreadLock C_SlCd, Row , C_SlCd, Row
					ggoSpread.SpreadLock C_SlCdPopup, Row , C_SlCdPopup, Row
					ggoSpread.SpreadLock C_LotNo, Row, C_LotSubNo, Row			
				End If			

				If lgRedrewFlg = True Then .ReDraw = True
			
			Case C_SlCdPopup
				.Col = C_SLCD
				.Row = Row
				Call OpenSLCD2(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_SLCD,Row,"M","X","X")
				Set gActiveElement = document.activeElement
		End Select
	End With
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
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
'   Event Name : txtRcptDT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtRcptDT_DblClick(Button)
    If Button = 1 Then
        frm1.txtRcptDT.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtRcptDT.Focus
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
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
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
    On Error Resume Next                                                   '��: Protect system from crashing    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 

    Dim IntRetCD 
    
    FncSave = False												'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing
   
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then				'��: Check required field(Multi area)
       Exit Function
    End If
        
    '-----------------------
    'Save function call area
    '-----------------------
    
    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True												'��: Processing is OK

End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function	 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                             '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 

End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                                  '��: Protect system from crashing
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                             '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                             '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
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

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'**********************************************************************************************************

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
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtSlCd=" & Trim(.hSlCd.value)
		strVal = strVal & "&txtFromDt=" & Trim(.hFromDt.value)
		strVal = strVal & "&txtToDt=" & Trim(.hToDt.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows 
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
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
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
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
Function DbQueryOk()
	
	Dim LngRow
	
	Call SetToolbar("11001001000111")

    Call ggoOper.LockField(Document, "N")
	frm1.txtRcptDT.text = LocSvrDate
	
		With frm1.vspdData
		
		.ReDraw = false

			.Row = 1
			' �������� 
			.Col = C_ProdtOrderUnit
			frm1.txtOrderUnit.Value = .Text
			' �������� 
			.Col = C_ProdtOrderQty
			frm1.txtOrderQty.Value = .Text
			' �ѻ��귮 
			.Col = C_ProdQtyInOrderUnit
			frm1.txtProdQty.Value = .Text
			' ��ǰ���� 
			.Col = C_GoodQtyInOrderUnit
			frm1.txtGoodQty.Value = .Text
			' �԰���� 
			.Col = C_RcptQtyInOrderUnit
			frm1.txtRcptQty.Value = .Text
			
			' ���ش��� 
			.Col = C_BaseUnit
			frm1.txtBaseUnit.Value = .Text
			' �������� 
			.Col = C_OrderQtyInBaseUnit
			frm1.txtOrderQty1.Value = .Text
			' �ѻ��귮 
			.Col = C_ProdQtyInBaseUnit
			frm1.txtProdQty1.Value = .Text
			' ��ǰ���� 
			.Col = C_GoodQtyInBaseUnit
			frm1.txtGoodQty1.Value = .Text
			' �԰���� 
			.Col = C_RcptQtyInBaseUnit
			frm1.txtRcptQty1.Value = .Text
			
			' ���������� 
			.Col = C_PlanStartDt
			frm1.txtPlanStratDt.text = .Text
			' �ϷΌ���� 
			.Col = C_PlanComptDt
			frm1.txtPlanEndDt.Text	= .Text
			' �۾������� 
			.Col = C_ReleaseDt
			frm1.txtReleaseDt.Text	= .Text
			' �������� 
			.Col = C_RealStartDt
			frm1.txtRealStratDt.Text = .Text
			' ���û��� 
			.Col = C_OrderStatus
			frm1.txtOrderStatus.value = .Text
				
		.ReDraw = True	
					
		End With 
	
	frm1.btnAutoSel.disabled = False
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
	
End Function

Function DbQueryNotOk()														'��: ��ȸ ������ ������� 
	Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	
End Function	
'========================================================================================
' Function Name : DbSave
' Function Desc : This function is to execute transaction.
'========================================================================================
Function DbSave() 

    Dim lRow    
	Dim strVal
	Dim strDate											'Issued Date
	Dim strReportDate									'Report Date
	
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size
	
	DbSave = False                                                          '��: Processing is NG
    
    Call LayerShowHide(1)
    
    frm1.txtMode.value = parent.UID_M0002
	frm1.txtUpdtUserId.value = parent.gUsrID
	frm1.txtInsrtUserId.value = parent.gUsrID
		
	'-----------------------
	'Data manipulate area
	'-----------------------
	iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
    
	'102399byte
	iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
	'������ �ʱ�ȭ 
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)		

	iTmpCUBufferCount = -1
	
	strCUTotalvalLen = 0
    
    strDate = frm1.txtRcptDT.text
    
	With frm1.vspdData
	
		For lRow = 1 To .MaxRows
		
		    .Row = lRow
		    .Col = 0
		
			.Col = C_ReportDt
			strReportDate = .Text
			
			.Col = C_Select
		    
			If .Value = 1 Then
			
				If strReportDate <> "" Then
					
					If CompareDateByFormat(strReportDate, strDate, "������", "�԰���","970023", parent.gDateFormat, parent.gComDateType,True) = False Then
						  Call LayerShowHide(0)
						  .EditMode = True
						  strVal = ""
						  Exit Function               
					End If
					
					If CompareDateByFormat(strDate, LocSvrDate,"�԰���","������","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
					  Call LayerShowHide(0)
					   .EditMode = True
					   strVal = ""
					  Exit Function               
					End If 
					
					strVal = ""
					
					'//Ref. ConstBas\P0\BCP4B3_PProdGoodsRcpt.bas
				    .Col = C_ProdtOrderNo			
				    strVal = strVal & Trim(.Text) & iColSep	'ProdtOrderNo
				    .Col = C_OprNo					
				    strVal = strVal & Trim(.Text) & iColSep	'OprNo
				    .Col = C_ItemCd
					strVal = strVal & Trim(.Text) & iColSep	'ItemCd
				    .Col = C_Seq					
				    strVal = strVal & CInt(Trim(.Text)) & iColSep			'Seq
				    .Col = C_ReportType	
				    strVal = strVal & Trim(.Text) & iColSep					'ReportType
					.Col = C_RcptQty				
				    strVal = strVal & UNIConvNum(.Text,0) & iColSep	'QtyInBaseUnit
				    .Col = C_BaseUnit	
				    strVal = strVal & Trim(.Text) & iColSep					'BaseUnit	
				    .Col = C_LotNo					
				    strVal = strVal & UCase(Trim(.Text)) & iColSep	'LotNo
				    .Col = C_LotSubNo
				    strVal = strVal & UNIConvNum(.Text,0) & iColSep	'LotSubNo
				    .Col = C_TrackingNo
				    strVal = strVal & Trim(.Text) & iColSep					'TrackingNo
				    .Col = C_SlCd					
				    strVal = strVal & Trim(.Text) & iColSep	'SLCD
				    .Col = C_WcCd
				    strVal = strVal & UCase(Trim(.Text)) &	iColSep			'WCCD
					.Col = C_OrderStatus
				    strVal = strVal & UCase(Trim(.Text)) & iColSep			'OrderStatus
				    
				    '------------------------------------------------
				    '//		Insert another txtSpread value
				    '------------------------------------------------
				    				    
				    strVal = strVal & UNIConvDate(strDate) & iColSep		'RcptDate
					strVal = strVal & frm1.txtRcptNo.value & iColSep		'RcptNo
					strVal = strVal & lRow & iRowSep						'Count (to trace error row)
					
					If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
			                            
			           Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
			           objTEXTAREA.name = "txtCUSpread"
			           objTEXTAREA.value = Join(iTmpCUBuffer,"")
			           divTextArea.appendChild(objTEXTAREA)     
			 
			           iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
			           ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			           iTmpCUBufferCount = -1
			           strCUTotalvalLen  = 0
			        End If
			       
			        iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			        If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
			           iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
			           ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			        End If   
			         
			        iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			        strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			        
				End If
				
			End If
			            
		Next
		
		If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)
		End If   

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)							'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()															'��: ���� ������ ���� ���� 
   
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
