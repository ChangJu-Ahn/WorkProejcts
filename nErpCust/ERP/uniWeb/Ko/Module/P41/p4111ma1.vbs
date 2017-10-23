
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_INC_CLS_DT		= "p4100mb1.asp"			'��: LookUp Plant for Inventory Close Date
Const BIZ_PGM_QRY_ID			= "p4111mb1.asp"			'��: LookUp Production Order Header
Const BIZ_PGM_SAVE_ID			= "p4111mb2.asp"			'��: Manage Production Order
Const BIZ_PGM_LOOKUP_ID			= "p4111mb0.asp"			'��: LookUp Item By Plant
Const BIZ_PGM_MAJOR_ROUT		= "p4111mb4.asp"			'��: LookUp Major Routing
Const BIZ_PGM_RELEASE_ID		= "p4111mb3.asp"			'��: Release Production Order
Const BIZ_PGM_JUMPORDERRUN_ID	= "p4110ma1.asp"
Const BIZ_PGM_LOOKUP_DATE		= "p4111mb5.asp"			'��: LookUp Planned Date


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgInvCloseDt				'������� 
Dim lgMajorRout, lgCostCd, lgCostNm
Dim lgDtValidFromDt
Dim lgDtValidToDt
Dim lgOPMDMode
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop			'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim lgCboKeyPress
Dim lgOldIndex
Dim lgOldIndex2
Dim lgQueryType

Dim lgCalType					'Calendar Type
Dim lgPlannedDate
Dim lgReworkMode

'#########################################################################################################
'						2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'====================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			'��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False				'��: Indicates that no value changed
    lgIntGrpCount = 0					'��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgReworkMode = "N"    
	frm1.btnRelease.disabled = True
	
End Sub


'************************************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *******************************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'***************************************************************************************************************************************************** 
'========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=================================================================================================== 
Sub SetDefaultVal()
	lgBlnFlgChgValue = False	
	frm1.cboReWork.value = "N"
	frm1.txtProdMgr.value = ""
	txtPlantCd_onChange()  
End Sub

'========================================  2.2.1 SetCookieVal()  ======================================
'	Name : SetCookieVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=================================================================================================== 
Sub SetCookieVal()
   	
	lgBlnFlgChgValue = False
	
	frm1.cboReWork.value = "Y"
	frm1.txtPlantCd.Value	= ReadCookie("txtPlantCd")
	frm1.txtPlantNm.value	= ReadCookie("txtPlantNm")
	frm1.txtItemCd.Value	= ReadCookie("txtItemCd")

	frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement

	gLookUpEnable = True
	Call LookUpItemByPlant()
	gLookUpEnable = False
		
	frm1.txtParentOrderNo.Value	= ReadCookie("txtProdOrderNo")
	frm1.txtParentOprNo.Value	= ReadCookie("txtOprNo")
	frm1.txtOrderQty.Text = ReadCookie("txtJumpQty")
	frm1.txtTrackingNo.value = ReadCookie("txtTrackingNo")
	Call txtPlantCd_onchange()

	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""	
	WriteCookie "txtProdOrderNo", ""
	WriteCookie "txtOprNo", ""
	WriteCookie "txtTrackingNo", ""
	WriteCookie "txtJumpQty", ""
		
End Sub
'***********************************************************  2.3 Operation ó���Լ�  *****************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�. 
'*********************************************************************************************************


'***********************************************************  2.4 POP-UP ó���Լ�  ********************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************

'======================================= 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenCondPlant()  -----------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()

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
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  ------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()

	Dim arrRet
	Dim arrParam(5), arrField(15)
	Dim iCalledAspName

	If IsOpenPop = True or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1029!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 	'ITEM_CD		' Field��(0)
	arrField(1) = 2 	'ITEM_NM		' Field��(1)
	arrField(2) = 26 	'UNIT_OF_ORDER
	arrField(3) = 4		'BASIC_UNIT
	arrField(4) = 28	'ORDER_LT
	arrField(5) = 33	'MIN_MRP_QTY
	arrField(6) = 34	'MAX_MRP_QTY
	arrField(7) = 35	'ROUND_QTY
	arrField(8) = 15	'MAJOR_SL_CD
	arrField(9) = 13	'PHANTOM_FLG
	arrField(10) = 25	'TRACKING_FLG
	arrField(11) = 17	'VALID_FLG
	arrField(12) = 18	'VALID_FROM_DT
	arrField(13) = 19	'VALID_TO_DT
	arrField(14) = 49	'INSPEC_MGR
	arrField(15) = 3	'SPECIFICATION

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
	
End Function

 '------------------------------------------  OpenSLCd()  ----------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True or UCase(frm1.txtSLCd.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		'Call DisplayMsgBox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtSLCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")	' Where Condition
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

'------------------------------------------  OpenProdOrderNo()  ---------------------------------------
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
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "OP"
	arrParam(4) = "OP"
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

'------------------------------------------  OpenParentOrderNo()  ---------------------------------------
'	Name : OpenParentOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenParentOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtParentOrderNo.className) = "PROTECTED" Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "ST"
	arrParam(4) = "CL"
	arrParam(5) = Trim(frm1.txtParentOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""		
	arrParam(8) = ""
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetParentOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtParentOrderNo.focus
	
End Function

'------------------------------------------  OpenParentOprNo()  -------------------------------------------------
'	Name : OpenParentOprNo()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenParentOprNo()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If frm1.txtParentOrderNo.value = "" Then
		Call DisplayMsgBox("971012","X" , "����������ȣ","X")
		frm1.txtParentOrderNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	iCalledAspName = AskPRAspName("P4112PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4112PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Or UCase(frm1.txtParentOprNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtParentOrderNo.value
	arrParam(2) = "" 'frm1.txtOprCd.value
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetParentOprNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtParentOprNo.focus
	
End Function


'====================  OpenRoutingNo  ======================================
' Function Name : OpenRoutingNo
' Function Desc : OpenRoutingNo Reference Popup
'==========================================================================
Function OpenRoutingNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True or UCase(frm1.txtRouting.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "ǰ��","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
		
	arrParam(0) = "����� �˾�"					' �˾� ��Ī 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtRouting.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  "And P_ROUTING_HEADER.ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") 	
	arrParam(5) = "�����"			
	
    arrField(0) = "ROUT_NO"							' Field��(0)
    arrField(1) = "DESCRIPTION"						' Field��(1)
    arrField(2) = "BOM_NO"							' Field��(1)
    arrField(3) = "MAJOR_FLG"						' Field��(1)
   
    arrHeader(0) = "�����"						' Header��(0)
    arrHeader(1) = "����ø�"					' Header��(1)
    arrHeader(2) = "BOM Type"					' Header��(1)
    arrHeader(3) = "�ֶ����"					' Header��(1)        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetRoutingNo(arrRet)
                Call LookUpRouting()
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtRouting.focus
	
End Function

 '------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtUnit.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"
	arrParam(2) = Trim(frm1.txtUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & ""
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtUnit.focus
	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = frm1.txtPlanStartDt.Text
	arrParam(4) = frm1.txtPlanEndDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function

'------------------------------------------  OpenOprRef()  --------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprRef()
	
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
'		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "x", "x", "x")
			Exit Function
'		End If
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtProdOrderNo1.value)	'��: ��ȸ ���� ����Ÿ 

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenPartRef()  -------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPartRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
'		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "x", "x", "x")
			Exit Function
'		End If
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4311RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4311RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtProdOrderNo1.value)	'��: ��ȸ ���� ����Ÿ 
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenStockRef()  -------------------------------------------
'	Name : OpenStockRef()
'	Description : Stock Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenStockRef()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "ǰ��","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4212RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4212RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(UCase(frm1.txtPlantCd.value))
	arrParam(1) = Trim(UCase(frm1.txtItemCd.value))
	arrParam(2) = Trim(frm1.txtItemNm.value)
	arrParam(3) = Trim(UCase(frm1.txtSLCd.value))
	arrParam(4) = Trim(frm1.txtSLNm.value)

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenCostCtr()  ----------------------------------------------
'	Name : OpenCostCtr()
'	Description : Cost Center Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCostCtr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(Frm1.txtCostCd.className) = "PROTECTED" Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	IsOpenPop = True 

	arrParam(0) = "Cost Center �˾�"			' �˾� ��Ī 
	arrParam(1) = "B_COST_CENTER"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtCostCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "B_COST_CENTER.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				" AND B_COST_CENTER.COST_TYPE ='M'" & _
				" AND B_COST_CENTER.DI_FG ='D'"			' Where Condition
	arrParam(5) = "Cost Center"					' TextBox ��Ī 
	
    arrField(0) = "COST_CD"							' Field��(0)
    arrField(1) = "COST_NM"							' Field��(1)
    
    arrHeader(0) = "Cost Center"				' Header��(0)
    arrHeader(1) = "Cost Center ��"				' Header��(1)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCostCtr(arrRet)
	End If	
    
End Function

'----------------------------------------  LookUpInvClsDt()  -------------------------------------------
'	Name : LookUpInvClsDt()
'	Description : LookUp Inventory Close Date
'---------------------------------------------------------------------------------------------------------
Function LookUpInvClsDt()

	Dim strVal
	
	Call LayerShowHide(1)
    
       strVal = BIZ_PGM_INC_CLS_DT & "?txtMode=" & parent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
       strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	'��: ��ȸ ���� ����Ÿ 
        
    Call RunMyBizASP(MyBizASP, strVal)								'��: �����Ͻ� ASP �� ���� 
	
End Function

'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'	Name : LookUpItem ByPlant()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 
Function LookUpItemByPlant()
    
	Dim strVal
	
	If gLookUpEnable = False Then Exit Function
	
    If LayerShowHide(1) = False Then Exit Function
    
    strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & parent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)	'��: ��ȸ ���� ����Ÿ 
        
    Call RunMyBizASP(MyBizASP, strVal)						'��: �����Ͻ� ASP �� ���� 
	
End Function

'-------------------------------------  LookUpItemByPlant Fail()  ---------------------------------------
'	Name : LookUpItemByPlantFail()
'	Description : LookUp Item By Plant Fail
'--------------------------------------------------------------------------------------------------------- 
Function LookUpItemByPlantFail()

    With frm1
	
		.txtItemCd.Value		= ""
		.txtItemNm.Value		= ""
		.txtUnit.value			= ""
		.txtProdLT.value		= ""
		.txtMaxLotQty.value		= ""
		.txtMinLotQty.value		= ""
		.txtRoundingQty.value	= ""
		.txtSLCd.value			= ""
		.txtSLNm.value			= ""
		.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		    
	End With
	
End Function

'-------------------------------------  LookUpItemByPlant Success()  ---------------------------------------
'	Name : LookUpItemByPlantSuccess()
'	Description : LookUp Item By Plant Success
'--------------------------------------------------------------------------------------------------------- 
Function LookUpItemByPlantSuccess()
	' when rework order is operating, Item Unit should be value from production results(p4413ma1/p4416ma1)
	If Trim(ReadCookie("txtOrderUnit")) <> "" Then
		frm1.txtUnit.value = ReadCookie("txtOrderUnit")
	End If
	
	Call LookUpMajorRouting()
End Function

'-------------------------------------  LookUpMajorRouting()  -----------------------------------------
'	Name : LookUpMajorRouting()
'	Description : LookUp Major Routing
'--------------------------------------------------------------------------------------------------------- 
Function LookUpMajorRouting()
    
	Dim strVal
	
	If gLookUpEnable = False Then Exit Function
	
    If LayerShowHide(1) = False Then Exit Function
    
    strVal = BIZ_PGM_MAJOR_ROUT & "?txtMode=" & parent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)	'��: ��ȸ ���� ����Ÿ 
        
    Call RunMyBizASP(MyBizASP, strVal)								'��: �����Ͻ� ASP �� ���� 
	
End Function

'-------------------------------------  LookUpMajorRoutingSuccess()  -----------------------------------------
'	Name : LookUpMajorRoutingSuccess()
'	Description : LookUp Major Routing
'--------------------------------------------------------------------------------------------------------- 
Function LookUpMajorRoutingSuccess()
	frm1.txtRouting.value = lgMajorRout
	frm1.txtCostCd.value = lgCostCd
	frm1.txtCostNm.value = lgCostNm
	If lgReworkMode = "Y" Then
		Call LookUpInvClsDt()
	End If
End Function

'-------------------------------------  LookUpRouting()  -----------------------------------------
'	Name : LookUpRouting()
'	Description : LookUp Major Routing
'--------------------------------------------------------------------------------------------------------- 
Function LookUpRouting()


    If 	CommonQueryRs("A.ROUT_NO, A.COST_CD, B.COST_NM ", "P_ROUTING_HEADER A , B_COST_CENTER B ", _
				" A.PLANT_CD *= B.PLANT_CD AND A.COST_CD *= B.COST_CD AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & _
				" AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd.value, "''", "S") & " AND A.ROUT_NO = " & FilterVar(frm1.txtRouting.value, "''", "S") , _
				lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then
		Exit Function
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1, Chr(11))
	lgF2 = Split(lgF2, Chr(11))
	
	frm1.txtRouting.value = lgF0(0)
	frm1.txtCostCd.value = lgF1(0)
	frm1.txtCostNm.value = lgF2(0)
	
End Function

'-------------------------------------  LookUpDate()  -----------------------------------------
'	Name : LookUpDate()
'	Description : LookUp Major Routing
'--------------------------------------------------------------------------------------------------------- 
Function LookUpDate(Byval strType)
    
	Dim strVal
	Dim LngProdLt
	Dim TempLt
	
	If gLookUpEnable = False Then Exit Function
	
    If LayerShowHide(1) = False Then Exit Function
	
	If Trim(frm1.txtProdLT.value) = "" Then
		TempLt = 0
	Else
		TempLt = frm1.txtProdLT.value	
	End If	
	
    If strType = "START_DATE" Then
		LngProdLt = 0 - CInt(TempLt)
	Else
		LngProdLt = CInt(TempLt)
	End If

	If LngProdLt = 0 Then
		If strType = "START_DATE" Then
			lgPlannedDate = Trim(frm1.txtPlanStartDt.text)
		Else
			lgPlannedDate = Trim(frm1.txtPlanEndDt.text)
		End If
		Call LookUpDateSuccess(strType)
		Call LayerShowHide(0)
		Exit Function
	End If

    strVal = BIZ_PGM_LOOKUP_DATE & "?txtMode=" & parent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCalType=" & Trim(lgCalType)
    strVal = strVal & "&txtProdLT=" & LngProdLt
    If strType = "START_DATE" Then
		strVal = strVal & "&txtPlanDt=" & Trim(frm1.txtPlanStartDt.text)
    Else
		strVal = strVal & "&txtPlanDt=" & Trim(frm1.txtPlanEndDt.text)
    End If
    strVal = strVal & "&txtType=" & strType
       
    Call RunMyBizASP(MyBizASP, strVal)								'��: �����Ͻ� ASP �� ���� 
	
End Function

'-------------------------------------  LookUpDateSuccess()  -----------------------------------------
'	Name : LookUpDateSuccess()
'	Description : LookUp Major Routing
'--------------------------------------------------------------------------------------------------------- 
Function LookUpDateSuccess(Byval strType)
    If strType = "START_DATE" Then
		frm1.txtPlanEndDt.text = lgPlannedDate
    Else
		frm1.txtPlanStartDt.text = lgPlannedDate
    End If
	
End Function

'Add 2005-09-27
Sub ProtectCostCd()

	If UCase(Trim(Frm1.hOprCostFlag.value)) = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtCostCd, "N")  
	Else
		Frm1.txtCostCd.value = ""
		Frm1.txtCostNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtCostCd, "Q")  
	End If
End Sub

'---------------------------------------------  ReleaseOrder()  ------------------------------------------
'	Name : ReleaseOrder()
'	Description : ReleaseOrder
'--------------------------------------------------------------------------------------------------------- 
Function ReleaseOrder()

	Dim IntRetCD, strVal
	Dim iColSep
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "x", "x", "x")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		Call DisplayMsgBox("189217", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	iColSep = Parent.gColSep
	
    If LayerShowHide(1) = False Then Exit Function
	strVal = ""
   	strVal = strVal & "CREATE" & iColSep
	strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
    strVal = strVal & UCase(Trim(frm1.txtProdOrderNo1.Value)) & iColSep        
	strVal = strVal & 0 & parent.gRowSep
	
	frm1.txtSpread.value = strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_RELEASE_ID)										'��: �����Ͻ� ASP �� ���� 
	
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetUnit(byval arrRet)
	frm1.txtUnit.Value    = arrRet(0)		
	lgBlnFlgChgValue = True
End Function
'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)
	frm1.txtProdOrderNo.focus
	Set gActiveElement = document.activeElement	
End Function
'------------------------------------------  SetParentOrderNo()  --------------------------------------------------
'	Name : SetParentOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetParentOrderNo(byval arrRet)
	frm1.txtParentOrderNo.Value    = arrRet(0)
	frm1.txtParentOrderNo.focus
	Set gActiveElement = document.activeElement	
End Function

'------------------------------------------  SetOprCd()  --------------------------------------------------
'	Name : SetParentOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetParentOprNo(byval arrRet)
	frm1.txtParentOprNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetRoutingNo()  --------------------------------------------------
'	Name : SetRoutingNo()
'	Description : RoutingNo Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRoutingNo(Byval arrRet)
	frm1.txtRouting.value = arrRet(0)
	frm1.txtBomNo.value = arrRet(2)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet)
	
	If arrRet(9) = "Y" Then 'PHANTOM_FLG
		Call DisplayMsgBox("189214", "x", "x", "x")
		Exit Function
	End If

	If arrRet(11) = "N" Then 'VALID_FLG
		Call DisplayMsgBox("122729", "x", "x", "x")
		Exit Function	
	End If
	
	If arrRet(10) = "N" Then 'TRACKING_FLG
		frm1.txtTrackingNo.ReadOnly = True
		frm1.txtTrackingNo.classname = "protected"
		frm1.txtTrackingNo.tabindex = "-1"
	Else
		frm1.txtTrackingNo.ReadOnly = False
		frm1.txtTrackingNo.classname = "required"
		frm1.txtTrackingNo.tabindex = "1"
	End If	

	frm1.txtItemCd.Value		= arrRet(0)
	frm1.txtItemNm.Value		= arrRet(1)
	frm1.txtUnit.value			= arrRet(2)
	frm1.txtProdLT.value		= arrRet(4)
	frm1.txtMaxLotQty.value		= arrRet(6)
	frm1.txtMinLotQty.value		= arrRet(5)
	frm1.txtRoundingQty.value	= arrRet(7)
	frm1.txtSLCd.value			= arrRet(8)
	frm1.txtBaseUnit.value 		= arrRet(3)
	frm1.txtSpecification.value = arrRet(15)
	lgDtValidFromDt				= arrRet(12)
	lgDtValidToDt				= arrRet(13)

	LookUpMajorRouting()

	lgBlnFlgChgValue = True

End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
    frm1.txtTrackingNo.Value = arrRet(0)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtProdOrderNo.focus
	Set gActiveElement = document.activeElement
	
	Call LookUpInvClsDt()
	
	Call txtPlantCd_onchange()
	
End Function

'------------------------------------------  SetCostCtr()  -----------------------------------------------
'	Name : SetCostCtr()
'	Description : Cost Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetCostCtr(byval arrRet)
	frm1.txtCostCd.value = arrRet(0)
	frm1.txtCostNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function



'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Function JumpOrderRun()

    Dim IntRetCd, strVal
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "x", "x", "x")
			Exit Function
		End If
	End If

	If lgBlnFlgChgValue = True Then
		Call DisplayMsgBox("189217", "x", "x", "x")
		Exit Function
	End If
		
	If frm1.txtStatus.value = "CL" Then
		Call DisplayMsgBox("189222", "x", "x", "x")
		Exit Function
	End If		

	If frm1.cboReWork.value = "Y" Then
		Call DisplayMsgBox("189218", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
		
	WriteCookie "txtPlantCd", UCase(Trim(frm1.txtPlantCd.value))
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	WriteCookie "txtItemCd", UCase(Trim(frm1.txtItemCd.value))
	WriteCookie "txtItemNm", Trim(frm1.txtItemNm.value)
	WriteCookie "txtSpecification", Trim(frm1.txtSpecification.value)
	
	WriteCookie "txtProdOrderNo", UCase(Trim(frm1.txtProdOrderNo1.value))
	WriteCookie "txtPlanOrderNo", UCase(Trim(frm1.txtPlanOrderNo.value))
	WriteCookie "txtOrderQty", UCase(Trim(frm1.txtOrderQty.value)) 
	WriteCookie "txtOrderUnit", UCase(Trim(frm1.txtUnit.value))
	WriteCookie "txtPlanStartDt", UCase(Trim(frm1.txtPlanStartDt.text))
	WriteCookie "txtPlanEndDt", UCase(Trim(frm1.txtPlanEndDt.text))
	WriteCookie "txtInvCloseDt", lgInvCloseDt
'	WriteCookie "txtPGMID", "P4111MA1"		
	navigate BIZ_PGM_JUMPORDERRUN_ID	
	
End Function

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtPlantCd_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtPlantCd_onChange()
	 Dim IntRetCd

    If  frm1.txtPlantCd.value = "" Then
        frm1.txtPlantCd.Value = ""
        frm1.txtPlantNm.Value = ""
        frm1.hOprCostFlag.value = ""
    Else
		
		Call LookUpInvClsDt()
		
        IntRetCD =  CommonQueryRs(" a.plant_nm, b.opr_cost_flag "," b_plant a (nolock), p_plant_configuration b (nolock) ", _
							" a.plant_cd = b.plant_cd and a.plant_cd = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & "" , _
							lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False   Then
			frm1.txtPlantNm.Value=""
			frm1.hOprCostFlag.value = ""
        Else
            frm1.txtPlantNm.Value= Trim(Replace(lgF0,Chr(11),""))
            frm1.hOprCostFlag.Value= Trim(Replace(lgF1,Chr(11),""))
        End If
        
        Call ProtectCostCd()
		
     End If

End Sub


'==========================================================================================
'   Event Name : txtItemCd_onChange()
'   Event Desc :
'==========================================================================================
Sub txtItemCd_onChange()
	With frm1
		If .txtItemCd.value = "" Then
			.txtItemCd.Value		= ""
			.txtItemNm.Value		= ""
			.txtUnit.value			= ""
			.txtBaseUnit.value		= ""
			.txtProdLT.value		= ""
			.txtMaxLotQty.value		= ""
			.txtMinLotQty.value		= ""
			.txtRoundingQty.value	= ""
			.txtSLCd.value			= ""
			.txtSLNm.value			= ""
			.txtRouting.value 		= ""
			.txtTrackingNo.value 	= ""
			.txtSpecification.value = ""
			.txtItemCd.focus
			Set gActiveElement = document.activeElement
		Else	
			.txtRouting.value 		= ""
			.txtTrackingNo.value 	= ""
			Call LookUpItemByPlant()
		End If
	End With
End Sub


'==========================================================================================
'   Event Name : cboReWork_onChange()
'   Event Desc :
'==========================================================================================
Sub cboReWork_onChange()
    lgBlnFlgChgValue = True
    If frm1.cboReWork.value = "Y" And lgIntFlgMode = parent.OPMD_CMODE Then
		Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"D")
		Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"D")
	ElseIf frm1.cboReWork.value = "N" And lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtParentOrderNo.value = ""
		Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"Q")
		Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"Q")
	End If	
End Sub

'==========================================================================================
'   Event Name : txtUnit_onChange()
'   Event Desc :
'==========================================================================================
Sub txtUnit_onChange()
	lgBlnFlgChgValue = True
    frm1.txtBaseOrderQty.value = 0
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtPlanStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanStartDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt_OnBlur()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtPlanStartDt_OnBlur()
	If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
         Exit Sub
	End If
   	If frm1.txtPlanEndDt.text = "" and frm1.txtPlanStartDt.text <> "" Then Call LookUpDate("START_DATE")
End Sub
Sub txtPlanStartDt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPlanEndDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtPlanEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanEndDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPlanEndDt_OnBlur()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtPlanEndDt_OnBlur()
	If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
         Exit Sub
	End If
    If frm1.txtPlanStartDt.text = "" and frm1.txtPlanEndDt.text <> "" Then Call LookUpDate("END_DATE")
End Sub
Sub txtPlanEndDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
'   Event Name : txtOrderQty_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtOrderQty_Change()
	frm1.txtBaseOrderQty.value = 0
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtBaseOrderQty_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtBaseOrderQty_Change()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtCostCd_onChange()
'   Event Desc :
'==========================================================================================
Sub txtCostCd_onChange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtRouting_onChange()
'   Event Desc : 2005-10-04 Add
'==========================================================================================
Sub txtRouting_onChange()
	lgBlnFlgChgValue = True
	If UCase(Trim(frm1.hOprCostFlag.value)) = "Y" And Trim(frm1.txtRouting.value) <> ""   Then
		Call LookUpRouting()	
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
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If

    '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '----------------------- 

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
     
    '-----------------------
    'Query function call area
    '----------------------- 

	lgQueryType = "CURR"

    If DbQuery = False Then Exit Function																'��: Query db data
      
    FncQuery = True																'��: Processing is OK
        
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()

    Dim IntRetCD 
    Dim strPlantCd, strPlantNm    
    
    FncNew = False																'��: Processing is NG
	'-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    strPlantCd	= frm1.txtPlantCd.value
    strPlantNm	= frm1.txtPlantNm.value
    
    Call ggoOper.ClearField(Document, "1")                                      '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
   	Call ggoOper.SetReqAttr(frm1.txtProdOrderNo1,"D")
   	Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"Q")
   	Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"Q")
   	Call ggoOper.SetReqAttr(frm1.txtTrackingNo, "Q")							
    Call InitVariables															'��: Initializes local global variables
    Call SetToolBar("11101000000011")
   

    If strPlantCd <> "" Then
		frm1.txtPlantCd.value = strPlantCd
		frm1.txtPlantNm.value = strPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 		
	Else
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtItemCd.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
	End If
    
    Call SetDefaultVal
    
    FncNew = True																'��: Processing is OK

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    
    Dim IntRetCD 
    
    FncDelete = False														'��: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "x", "x", "x")
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------

    IntRetCd = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	If IntRetCd = vbNo Then
		Exit Function
	End If

    lgIntFlgMode = 0000

    If DbDelete = False Then Exit Function													'��: Delete db data
    
    FncDelete = True											'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 

    Dim IntRetCD 
    Dim strInvCloseDt
    
    FncSave = False                                           '��: Processing is NG
    
    Err.Clear                                                 '��: Protect system from crashing
    
    If lgBlnFlgChgValue = False Then                          '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '��: Display Message(There is no changed data.)
        Exit Function
    End If
	'-----------------------
    'Check content area
    '-----------------------
    
	If frm1.txtOrderQty.Value = "" Then frm1.txtOrderQty.Value = 0
	If frm1.txtBaseOrderQty.Value = "" Then frm1.txtBaseOrderQty.Value = 0

	If frm1.txtOrderQty.Value = 0 Then
		Call DisplayMsgBox("189208", "x", "x", "x")
		frm1.txtOrderQty.focus
		Set gActiveElement = document.activeElement
		Exit Function	
    End If
    
	If frm1.txtOrderQty.Value < 0 Then  
		Call DisplayMsgBox("189208", "x", "x", "x")
		frm1.txtOrderQty.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
    
	If frm1.txtBaseOrderQty.Value < 0 Then  
		Call DisplayMsgBox("189208", "x", "x", "x")
		frm1.txtBaseOrderQty.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If Trim(frm1.cboReWork.value) = "Y" Then
		If (Trim(frm1.txtParentOrderNo.value)  = ""  XOR _
			Trim(frm1.txtParentOprNo.value) = "") Then
			Call DisplayMsgBox("189249","X", "X", "X")
			frm1.txtParentOrderNo.focus
			Set gActiveElement = document.activeElement
			Exit Function
		End If
	End If	

    If Not chkfield(Document, "2") Then					'��: Check required field(Single area)
       Exit Function
    End If

	If UniConvDateAToB(frm1.txtPlanStartDt.Text, parent.gDateFormat, parent.gServerDateFormat) > UniConvDateAToB(frm1.txtPlanEndDt.Text, parent.gDateFormat, parent.gServerDateFormat) Then  
		Call DisplayMsgBox("189207", "x", "x", "x")
		frm1.txtPlanEndDt.focus

		Set gActiveElement = document.activeElement
		Exit Function
    End If
 
	strInvCloseDt = UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat)

 	' lgInvCloseDt -> parent.gServerDateFormat
	'If UniConvDateAToB(frm1.txtPlanStartDt.Text, parent.gDateFormat, parent.gServerDateFormat) <= strInvCloseDt Then  
	'	Call DisplayMsgBox("189204", "x", "x", "x")
	'	frm1.txtPlanStartDt.focus
	'	Set gActiveElement = document.activeElement
	'	Exit Function
    'End If
    
	'If UniConvDateAToB(frm1.txtPlanEndDt.Text, parent.gDateFormat, parent.gServerDateFormat) <= strInvCloseDt Then  
	'	Call DisplayMsgBox("189205", "x", "x", "x")
	'	frm1.txtPlanEndDt.focus
	'	Set gActiveElement = document.activeElement
	'    Exit Function
    'End If

    If DbSave = False Then Exit Function				                                  '��: Save db data

    FncSave = True                                            '��: Processing is OK
   
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

	Dim IntRetCD
	Dim strPlantCd, strPlantNm

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												'��: Indicates that current mode is Crate mode

	strPlantCd = frm1.txtPlantCd.value
	strPlantNm = frm1.txtPlantNm.value
    
     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
    
	frm1.txtPlantCd.value = strPlantCd
	frm1.txtPlantNm.value = strPlantNm
    
    Call ggoOper.LockField(Document, "N")									'��: This function lock the suitable field
    
	frm1.txtPlanStartDt.Text = ""
	frm1.txtPlanEndDt.Text = ""
	frm1.txtPlannedStartDt.Text = ""
	frm1.txtPlannedEndDt.Text = ""
	frm1.txtReleaseDt.Text = ""
	frm1.txtStatus.value = ""
	frm1.txtPlanOrderNo.value = ""
    frm1.txtProdOrderNo1.value = ""

   	Call ggoOper.SetReqAttr(frm1.txtProdOrderNo1,"D")
   	Call ggoOper.SetReqAttr(frm1.txtRemark,"D")
   	frm1.cboReWork.value = "N"
   	Call ggoOper.SetReqAttr(frm1.txtParentOrderNo, "Q")
   	Call ggoOper.SetReqAttr(frm1.txtParentOprNo, "Q")
   	
   	Call txtItemCd_onChange
   	
    frm1.txtProdOrderNo1.focus
    Set gActiveElement = document.activeElement  
    lgBlnFlgChgValue = True
    
    Call SetToolBar("11101000000011")
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint()                                         '��: Protect system from crashing
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 

    Dim IntRetCD 
    
    FncPrev = False                                                        '��: Processing is NG
    
    Err.Clear                                                              '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '----------------------- 

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
     '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

  
    '-----------------------
    'Query function call area
    '----------------------- 

	lgQueryType = "PREV"

    If DbQuery = False Then Exit Function																'��: Query db data
      
    FncPrev = True																'��: Processing is OK

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 

    Dim IntRetCD 
    
    FncNext = False                                                        '��: Processing is NG
    
    Err.Clear                                                              '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '----------------------- 

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
   Call InitVariables															'��: Initializes local global variables
   
    '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
  
    '-----------------------
    'Query function call area
    '----------------------- 

	lgQueryType = "NEXT"

    If DbQuery = False Then Exit Function																'��: Query db data
      
    FncNext = True																'��: Processing is OK

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)											'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)									'��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

	Dim IntRetCD
	
	FncExit = False
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
    
    Err.Clear																'��: Protect system from crashing

    DbDelete = False														'��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function

    Dim strVal
	Dim iColSep
    
    strVal = ""
    
    iColSep = Parent.gColSep
    
    With frm1

	.txtMode.value = parent.UID_M0003												'��: �����Ͻ� ó�� ASP �� ���� 
	.txtFlgMode.value = lgIntFlgMode
	strVal = strVal & "DELETE" & iColSep
	' Plant Code
	strVal = strVal & UCase(Trim(.txtPlantCd.value)) & iColSep
	' Production Order No.
	strVal = strVal & UCase(Trim(.txtProdOrderNo1.value)) & iColSep
	' Item Code
	strVal = strVal & UCase(Trim(.txtItemCd.value)) & iColSep
	' Re-Work Flag
	strVal = strVal & Trim(.cboReWork.value) & iColSep
	' Order Quantity
	strVal = strVal & UNIConvNum(Trim(.txtOrderQty.value),0) & iColSep
	' Order Unit
	strVal = strVal & UCase(Trim(.txtUnit.value)) & iColSep
	' Base Quantity
	strVal = strVal & iColSep
	' Basic Unit
	strVal = strVal & UCase(Trim(.txtBaseUnit.value)) & iColSep
	' S/L Code
	strVal = strVal & UCase(Trim(.txtSLCd.value)) & iColSep
	' Routing No.
	strVal = strVal & UCase(Trim(.txtRouting.value)) & iColSep
	' Planned Start Date
	strVal = strVal & UNIConvDate(Trim(.txtPlanStartDt.value)) & iColSep
	' Planned End Date
	strVal = strVal & UNIConvDate(Trim(.txtPlanEndDt.value)) & iColSep
	' BOM Type
	strVal = strVal & UCase(Trim(.txtBOMNo.value)) & iColSep
	' Tracking No.
	If Trim(.txtTrackingNo.Value) = "" Then
		strVal = strVal & "*" & iColSep								'��: Tracking No.
	Else
		strVal = strVal & UCase(Trim(.txtTrackingNo.value)) & iColSep
	End If	
	' Remark
	strVal = strVal & Trim(.txtRemark.value) & iColSep
	
	strVal = strVal & Trim(UCase(.txtParentOrderNo.value)) & iColSep
	
	strVal = strVal & Trim(UCase(.txtParentOprNo.value)) & iColSep
	
	strVal = strVal & Trim(UCase(.txtCostCd.value)) & iColSep
	
	strVal = strVal & 0 & parent.gRowSep
	
	.txtSpread.value = strVal
    
    End With

	lgOPMDMode = "DELETE"

    Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)							'��: �����Ͻ� ASP �� ���� 

    DbDelete = True

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call FncNew()
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Err.Clear																'��: Protect system from crashing
    
    DbQuery = False															'��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
   	
   	If lgQueryType = "CURR" Then
		strVal = strVal & "&txtQueryType=" & ""
   	ElseIf lgQueryType = "PREV" Then
		strVal = strVal & "&txtQueryType=" & "P"
   	ElseIf lgQueryType = "NEXT" Then
		strVal = strVal & "&txtQueryType=" & "N"
	Else
		strVal = strVal & "&txtQueryType=" & "R"
	End If

    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbQuery = True															'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

    Call SetToolBar("11111000111111")

	Call ggoOper.SetReqAttr(frm1.txtProdOrderNo1,"Q")
	
    If frm1.txtStatus.Value = "PL" or frm1.txtStatus.Value = "OP" Then
		' Order Quantity
		Call ggoOper.SetReqAttr(frm1.txtOrderQty,"N")
		' Order Unit
		Call ggoOper.SetReqAttr(frm1.txtUnit,"N")
		' Planned Start Date
		Call ggoOper.SetReqAttr(frm1.txtPlanStartDt,"N")
			' Planned End Date
		Call ggoOper.SetReqAttr(frm1.txtPlanEndDt,"N")
			' Storage Location
		Call ggoOper.SetReqAttr(frm1.txtSLCd,"N")
		' Routing No.
		Call ggoOper.SetReqAttr(frm1.txtRouting,"N")
		' Tracking No.
		If Trim(frm1.txtTrackingNo.value) = "*" or Trim(frm1.txtTrackingNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtTrackingNo,"Q")
		Else
			Call ggoOper.SetReqAttr(frm1.txtTrackingNo,"N")		
		End If
		' Re-Work Flag
		Call ggoOper.SetReqAttr(frm1.cboReWork,"Q")
		' Remark
		Call ggoOper.SetReqAttr(frm1.txtRemark,"D")
		' ParentOrderNo
		Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"Q")
		' ParentOprNo
		Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"Q")
		
		Call ProtectCostCd()
		
		Call SetToolBar("11111000111111")
		
		frm1.btnRelease.disabled = False
		
    Else
		' Order Quantity
		Call ggoOper.SetReqAttr(frm1.txtOrderQty,"Q")
		' Order Unit
		Call ggoOper.SetReqAttr(frm1.txtUnit,"Q")
		' Planned Start Date
		Call ggoOper.SetReqAttr(frm1.txtPlanStartDt,"Q")
		' Planned End Date
		Call ggoOper.SetReqAttr(frm1.txtPlanEndDt,"Q")
		' Storage Location
		Call ggoOper.SetReqAttr(frm1.txtSLCd,"Q")
		' Routing No.
		Call ggoOper.SetReqAttr(frm1.txtRouting,"Q")
		' Tracking No.
		Call ggoOper.SetReqAttr(frm1.txtTrackingNo,"Q")
		' Re-Work Flag
		Call ggoOper.SetReqAttr(frm1.cboReWork,"Q")
		' Remark
		Call ggoOper.SetReqAttr(frm1.txtRemark,"Q")
		' ParentOrderNo
		Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"Q")
		' ParentOprNo
		Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"Q")
		
		Call ggoOper.SetReqAttr(frm1.txtCostCd,"Q")
		
		Call SetToolBar("11100000111111")

		frm1.btnRelease.disabled = True

    End If
    
	lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery�� ������ ��� MyBizASP ���� ȣ��Ǵ� Function,
'========================================================================================
Function DbQueryNotOk()
	Call SetToolBar("11101000001111")
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.SetReqAttr(frm1.txtTrackingNo, "Q")
	frm1.cboReWork.value = "N"
	Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"Q")
	Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"Q")
	lgIntFlgMode = parent.OPMD_CMODE
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 

    Err.Clear																'��: Protect system from crashing

    DbSave = False															'��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function

    Dim strVal
    Dim iColSep
    
    strVal = ""
    
    iColSep = Parent.gColSep

    With frm1
	
	.txtMode.value = parent.UID_M0002					'��: �����Ͻ� ó�� ASP �� ���� 
	.txtFlgMode.value = lgIntFlgMode
			
	If lgIntFlgMode = parent.OPMD_CMODE Then
		strVal = strVal & "CREATE" & iColSep
	Else 
		strVal = strVal & "UPDATE" & iColSep
	End If
	
	' Plant Code
	strVal = strVal & UCase(Trim(.txtPlantCd.value)) & iColSep
	' Production Order No.
	strVal = strVal & UCase(Trim(.txtProdOrderNo1.value)) & iColSep
	' Item Code
	strVal = strVal & UCase(Trim(.txtItemCd.value)) & iColSep
	' Re-Work Flag
	strVal = strVal & Trim(.cboReWork.value) & iColSep
	' Order Quantity
	strVal = strVal & UNIConvNum(Trim(.txtOrderQty.Text),0) & iColSep
	' Order Unit
	strVal = strVal & UCase(Trim(.txtUnit.value)) & iColSep
	' Base Quantity
	strVal = strVal & UNIConvNum("0",0) &iColSep
	' Basic Unit
	strVal = strVal & UCase(Trim(.txtBaseUnit.value)) & iColSep
	' S/L Code
	strVal = strVal & UCase(Trim(.txtSLCd.value)) & iColSep
	' Routing No.
	strVal = strVal & UCase(Trim(.txtRouting.value)) & iColSep
	' Planned Start Date
	strVal = strVal & UNIConvDate(Trim(.txtPlanStartDt.Text)) & iColSep
	' Planned End Date
	strVal = strVal & UNIConvDate(Trim(.txtPlanEndDt.Text)) & iColSep
	' BOM Type
	strVal = strVal & UCase(Trim(.txtBOMNo.value)) & iColSep
	' Tracking No.
	If Trim(.txtTrackingNo.Value) = "" Then
		strVal = strVal & "*" & iColSep								'��: Tracking No.
	Else
		strVal = strVal & UCase(Trim(.txtTrackingNo.value)) & iColSep
	End If	
	' Remark
	strVal = strVal & Trim(.txtRemark.value) & iColSep
	' Parent Order No
	strVal = strVal & UCase(Trim(.txtParentOrderNo.value)) & iColSep
	' Parent Opr No
	strVal = strVal & UCase(Trim(.txtParentOprNo.value)) & iColSep
	' Add 2005-09-28
	strVal = strVal & UCase(Trim(.txtCostCd.value)) & iColSep
	
	strVal = strVal & 0 & parent.gRowSep
	
	.txtSpread.value = strVal
		
    End With


	lgOPMDMode = "UPDATE"
	
	
    Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)							'��: �����Ͻ� ASP �� ���� 

    DbSave = True
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk(ByVal BlnRelease)															'��: ���� ������ ���� ���� 
	
    Call InitVariables

    lgBlnFlgChgValue = False
	
	 '-----------------------
    'Erase contents area
    '----------------------- 

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
     
    '-----------------------
    'Query function call area
    '----------------------- 

	If BlnRelease = True Then
		lgQueryType = "RELEASE"
	Else
		lgQueryType = "CURR"	
	End If	

    If DbQuery = False Then Exit Function

End Function
