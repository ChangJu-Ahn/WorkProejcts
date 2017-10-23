
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID	= "p4312mb1.asp"						'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID	= "p4312mb2.asp"						'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Dim C_CompntCd 
Dim C_CompntNm
Dim C_Spec 
Dim C_IssueDt 
Dim C_IssueQty 
Dim C_Unit 
Dim C_WcCd 
Dim C_LotNo 
Dim C_LotSubNo 
Dim C_SlipNo 
Dim C_Req_No 
Dim C_SeqNo 
Dim C_SubSeqNo 
Dim C_SLCd 
Dim C_Year 

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgIntGrpCount				' Group View Size�� ������ ���� 
Dim lgIntFlgMode					' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop  
Dim lgSortKey     
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

    lgIntFlgMode = parent.OPMD_CMODE			'Indicates that current mode is Create mode
    lgIntGrpCount = 0					'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""					'initializes Previous Key
    lgLngCurRows = 0					'initializes Deleted Rows Count
    lgSortKey  = 1
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'============================================================================================================
Sub InitSpreadSheet()
	
    Call InitSpreadPosVariables()
	
    With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
    
		.ReDraw = false
	
		.MaxCols = C_Year+1
		.MaxRows = 0
	
		Call AppendNumberPlace("6", "3", "0")
	
		Call GetSpreadColumnPos("A")
	
		ggoSpread.SSSetEdit		C_CompntCd, "��ǰ", 18
		ggoSpread.SSSetEdit		C_CompntNm, "��ǰ��", 25
		ggoSpread.SSSetEdit		C_Spec, "�԰�", 25
		ggoSpread.SSSetDate		C_IssueDt, "�����", 11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_IssueQty, "������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_Unit, "����", 8
		ggoSpread.SSSetEdit		C_WcCd, "�۾���", 10
		ggoSpread.SSSetEdit		C_LotNo, "Lot ��ȣ", 12
		ggoSpread.SSSetFloat	C_LotSubNo, "����", 6, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit		C_SlipNo, "����ȣ", 18
		ggoSpread.SSSetEdit		C_Req_No, "���ù�ȣ", 18
		ggoSpread.SSSetEdit		C_SeqNo, "����", 10
		ggoSpread.SSSetEdit		C_SubSeqNo, "����", 10
		ggoSpread.SSSetEdit		C_SLCd, "���â��", 10
		ggoSpread.SSSetEdit		C_Year, "��ǥ�߻��⵵", 10	

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SLCd, C_SLCd, True)
		Call ggoSpread.SSSetColHidden(C_Year, C_Year, True)
		Call ggoSpread.SSSetColHidden(C_Req_No, C_Req_No, True)
		Call ggoSpread.SSSetColHidden(C_SeqNo, C_SeqNo, True)
		Call ggoSpread.SSSetColHidden(C_SubSeqNo, C_SubSeqNo, True)

		ggoSpread.SSSetSplit2(2)							'frozen ����߰� 

		.ReDraw = true

		Call SetSpreadLock 

    End With
       
End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()
	C_CompntCd	= 1
	C_CompntNm	= 2
	C_Spec		= 3
	C_IssueDt	= 4
	C_IssueQty	= 5
	C_Unit		= 6
	C_WcCd		= 7
	C_LotNo		= 8
	C_LotSubNo	= 9
	C_SlipNo	= 10
	C_Req_No	= 11
	C_SeqNo		= 12
	C_SubSeqNo	= 13
	C_SLCd		= 14
	C_Year		= 15

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'=================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_CompntCd	= iCurColumnPos(1)
			C_CompntNm	= iCurColumnPos(2)
			C_Spec		= iCurColumnPos(3)
			C_IssueDt	= iCurColumnPos(4)
			C_IssueQty	= iCurColumnPos(5)
			C_Unit		= iCurColumnPos(6)
			C_WcCd		= iCurColumnPos(7)
			C_LotNo		= iCurColumnPos(8)
			C_LotSubNo	= iCurColumnPos(9)
			C_SlipNo	= iCurColumnPos(10)
			C_Req_No	= iCurColumnPos(11)
			C_SeqNo		= iCurColumnPos(12)
			C_SubSeqNo	= iCurColumnPos(13)
			C_SLCd		= iCurColumnPos(14)
			C_Year		= iCurColumnPos(15)
	End Select 
End Sub    

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
	With frm1
   		.vspdData.ReDraw = False
		ggoSpread.SpreadLock -1, -1
		.vspdData.ReDraw = True
	End With
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
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

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		'Call DisplayMsgBox("189220", "x", "x", "x")
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

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
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

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
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
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode			' Item Code
	arrParam(2) = ""				' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field��(0)
	arrField(1) = 2 '"ITEM_NM"			' Field��(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCompntCd.focus

End Function

'------------------------------------------  OpenOprRef()  -------------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
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
	
    If frm1.txtProdOrderNo.value= "" Then
		Call DisplayMsgBox("971012","X", "����������ȣ","X")
		frm1.txtProdOrderNo.focus 
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		'��: ��ȸ ���� ����Ÿ 
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenProdRef()  -------------------------------------------------
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
	
    If frm1.txtProdOrderNo.value= "" Then
		Call DisplayMsgBox("971012","X", "����������ȣ","X")
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4411RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4411RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		'��: ��ȸ ���� ����Ÿ 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  -------------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Issue Reference
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

    If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
    If frm1.txtProdOrderNo.value= "" Then
		Call DisplayMsgBox("971012","X", "����������ȣ","X")
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)		'��: ��ȸ ���� ����Ÿ 
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenConsumRef()  -------------------------------------------------
'	Name : OpenConsumRef()
'	Description : Part Consumption Reference
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
	
    If frm1.txtProdOrderNo.value= "" Then
		Call DisplayMsgBox("971012","X", "����������ȣ","X")
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4412RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4412RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet)
    With frm1
	.txtCompntCd.value = arrRet(0)
	.txtCompntNm.value = arrRet(1)
    End With
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
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
'********************************************************************************************************* 

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

'******************************  3.2.1 Object Tag ó��  *********************************************
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
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )
	
	
	If frm1.vspdData.MaxRows = 0 Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0101111111")
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
 		Exit Sub
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    'Exit Sub
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" AND lgStrPrevKey1 <> "" AND lgStrPrevKey2 <> ""  Then				'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
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

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then					'��: This function check indispensable field
       Exit Function
    End If

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtCompntCd.value = "" Then
		frm1.txtCompntNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If												'��: Query db data

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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 

    Dim IntRetCD 
    
    FncSave = False                                             '��: Processing is NG
    
    Err.Clear                                                   '��: Protect system from crashing
    
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
    If DbSave = False Then Exit Function				                                    '��: Save db data
    
    FncSave = True                                              '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow
    
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function	 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 

    With frm1
	.vspdData.focus
	Set gActiveElement = document.activeElement
	ggoSpread.Source = .vspdData
	'.vspdData.EditMode = True
	.vspdData.ReDraw = False
	ggoSpread.InsertRow
	.vspdData.ReDraw = True
	SetSpreadColor .vspdData.ActiveRow
    End With
    
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
	If frm1.vspdData.MaxRows < 1 Then Exit Function	     
    
    With frm1.vspdData 
    
    .focus
    Set gActiveElement = document.activeElement
    ggoSpread.Source = frm1.vspdData 
    
	.Col = 8'ggoSpread.SSGetColsIndex(8)
	If .Text = "A" Then
	    Call DisplayMsgBox("183104", "x", "x", "x")              '��: you must release this line if you change msg into code
	    Exit Function
	End If
    
	lDelRows = ggoSpread.DeleteRow
   
    End With
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
	
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim strVal

    Err.Clear                                                               '��: Protect system from crashing

    DbQuery = False
    
    Call LayerShowHide(1)
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)
		strVal = strVal & "&txtCompntCd=" & Trim(.hCompntCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
    	strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.value)
		strVal = strVal & "&txtCompntCd=" & Trim(.txtCompntCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
    	strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
    End IF
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    
End Function


'========================================================================================
' Function Name : HeaderQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function HeaderQueryOk()											'��: ��ȸ ������ ������� 

	lgIntFlgMode = parent.OPMD_UMODE									'��: Indicates that current mode is Update mode
	
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DBQueryOk()									'��: ��ȸ ������ ������� 

	Call SetToolBar("11001011000111")					'��: ��ư ���� ���� 
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE							'��: Indicates that current mode is Update mode
    Call ggoOper.LockField(Document, "Q")				'��: This function lock the suitable field
    ggoSpread.Source = frm1.vspdData
    
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
    
	iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '������ �ʱ�ȭ 
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpDBufferCount = -1
	
	strDTotalvalLen  = 0
    
    Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows
    
		.vspdData.Row = lRow
		.vspdData.Col = 0
	
		If .vspdData.Text = ggoSpread.DeleteFlag Then
				
			strDel = ""
			strDel = strDel & UCase(Trim(frm1.txtPlantCd.Value)) & iColSep				'��: C=Create
			strDel = strDel & UCase(Trim(frm1.txtProdOrderNo.Value)) & iColSep				'��: C=Create
			.vspdData.Col = C_SlipNo
			strDel = strDel & Trim(.vspdData.Text) & iColSep
			.vspdData.Col = C_SeqNo
			strDel = strDel & Trim(.vspdData.Text) & iColSep
			.vspdData.Col = C_SubSeqNo
			strDel = strDel & Trim(.vspdData.Text) & iColSep
			.vspdData.Col = C_IssueQty
			strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
			.vspdData.Col = C_Req_No
			strDel = strDel & Trim(.vspdData.Text) & iColSep
			.vspdData.Col = C_CompntCd
			strDel = strDel & Trim(.vspdData.Text) & iColSep 	
			.vspdData.Col = C_SLCd
			strDel = strDel & Trim(.vspdData.Text) & iColSep 	
			.vspdData.Col = C_Year
			strDel = strDel & Trim(.vspdData.Text) & iColSep
			'Row Count
			strDel = strDel & lRow & iRowSep	
			
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
	

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'��: �����Ͻ� ASP �� ���� 
	
	DbSave = True                                                           					'��: Processing is NG
		
	
	End With
    
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

	DbDelete = False														'��: Processing is NG
	
	On Error Resume Next
	
    DbDelete = True 
	
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

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'----------  Coding part  -------------------------------------------------------------
'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


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
Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.ReOrderingSpreadData()

	Call InitData(1)
    
End Sub 

