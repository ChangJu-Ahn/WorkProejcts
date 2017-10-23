'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p4313mb1.asp"											'��: �����Ͻ� ����(Qeury) ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Dim C_ProdtOrderNo	
Dim C_CompntCd		
Dim C_CompntNm
Dim C_Spec		
Dim C_RqrdQty		
Dim C_Unit			
Dim C_IssuedQty		
Dim C_RemainQty	
Dim C_ConsumedQty	
Dim C_TTlIssueQty	
Dim C_RqrdDt		
Dim C_ResrvStatus	
Dim C_ResrvStatusDesc
Dim C_IssueMthd		
Dim C_IssueMthdDesc	
Dim C_MajorSLCd		
Dim C_MajorSLNm		
Dim C_OprNo			
Dim C_WcCD			
Dim C_ReqSeqNo		
Dim C_ReqNo			
Dim C_ParentItemCd	
Dim C_ParentItemNm	
Dim C_ParentItemSpec
Dim C_TrackingNo	
Dim C_JobNm

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgOldRow
Dim lgLngCnt
Dim lgAfterQryFlg

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
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

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""			'initializes Previous Key
    lgStrPrevKey2 = ""
    lgLngCurRows = 0		'initializes Deleted Rows Count
    
    lgLngCnt = 0
    lgOldRow = 0
    lgAfterQryFlg = False
    lgSortKey    = 1
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
	frm1.txtReqStartDt.Text = StartDate
	frm1.txtReqEndDt.Text = EndDate
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    
	
	With frm1.vspdData 
		 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20051001",,parent.gAllowDragDropSpread    
		
		.ReDraw = false
	
		.MaxCols = C_JobNm + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ    
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 18
		ggoSpread.SSSetEdit		C_CompntCd, "��ǰ", 18
		ggoSpread.SSSetEdit		C_CompntNm, "��ǰ��", 25
		ggoSpread.SSSetEdit		C_Spec,		"��ǰ�԰�", 25
		ggoSpread.SSSetFloat	C_RqrdQty, 	"�ʿ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_Unit, 	"����", 7
		ggoSpread.SSSetDate 	C_RqrdDt, 	"�ʿ���", 11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_IssuedQty, "������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	'�������� 
		ggoSpread.SSSetFloat	C_RemainQty, "��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	'����ܷ� 
		ggoSpread.SSSetFloat	C_ConsumedQty, "�Һ����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	'�߰� 
		ggoSpread.SSSetFloat	C_TTlIssueQty, "������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit		C_MajorSLCd, "���â��", 10
		ggoSpread.SSSetEdit		C_MajorSLNm, "���â���", 20
		ggoSpread.SSSetEdit		C_OprNo, "����", 6
		ggoSpread.SSSetEdit		C_WcCD, "�۾���", 10
		ggoSpread.SSSetEdit		C_ReqSeqNo, "����", 6
		ggoSpread.SSSetEdit		C_ReqNo, "����", 6
		ggoSpread.SSSetEdit		C_ResrvStatus, "������", 10
		ggoSpread.SSSetEdit		C_ResrvStatusDesc, "������", 10
		ggoSpread.SSSetCombo	C_IssueMthd, "�����", 10
		ggoSpread.SSSetCombo	C_IssueMthdDesc, "�����", 10	
		ggoSpread.SSSetEdit		C_ParentItemCd, "ǰ��", 18
		ggoSpread.SSSetEdit		C_ParentItemNm, "ǰ���", 25
		ggoSpread.SSSetEdit		C_ParentItemSpec, "ǰ��԰�", 25
		ggoSpread.SSSetEdit 	C_TrackingNo, "Tracking No.", 25
		ggoSpread.SSSetEdit 	C_JobNm,		"�۾���", 10
		
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_ReqSeqNo, C_ReqSeqNo, True)
		Call ggoSpread.SSSetColHidden(C_ReqNo, C_ReqNo, True)
		Call ggoSpread.SSSetColHidden(C_ResrvStatus, C_ResrvStatus, True)
		Call ggoSpread.SSSetColHidden(C_IssueMthd, C_IssueMthd, True)
		Call ggoSpread.SSSetColHidden(C_TTlIssueQty, C_TTlIssueQty, True)
	
		ggoSpread.SSSetSplit2(2)							'frozen ����߰� 
	
		Call SetSpreadLock 
    
		.ReDraw = true
    
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
	
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
End Sub


'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitSpreadComboBox()
	
	Dim strData, strDesc
	Dim iCodeArr 
    Dim iNameArr
 
 	'Production Manager
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1016", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_IssueMthd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_IssueMthdDesc
	
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex
	
	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_IssueMthd
			intIndex = .value
			.Col = C_IssueMthdDesc
			.value = intindex
		Next	
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	
	C_ProdtOrderNo		= 1
    C_CompntCd			= 2
	C_CompntNm			= 3
	C_Spec				= 4
	C_RqrdQty			= 5
	C_Unit				= 6
	C_IssuedQty			= 7
	C_RemainQty			= 8
	C_ConsumedQty		= 9
	C_TTlIssueQty		= 10
	C_RqrdDt			= 11
	C_ResrvStatus		= 12
	C_ResrvStatusDesc	= 13
	C_IssueMthd			= 14
	C_IssueMthdDesc		= 15
	C_MajorSLCd			= 16
	C_MajorSLNm			= 17
	C_OprNo				= 18
	C_WcCD				= 19
	C_ReqSeqNo			= 20
	C_ReqNo				= 21
	C_ParentItemCd		= 22
	C_ParentItemNm		= 23
	C_ParentItemSpec	= 24
	C_TrackingNo		= 25
	C_JobNm				= 26
	
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
 
 				C_ProdtOrderNo	= iCurColumnPos(1)
				C_CompntCd		= iCurColumnPos(2)
				C_CompntNm		= iCurColumnPos(3)
				C_Spec			= iCurColumnPos(4)
				C_RqrdQty		= iCurColumnPos(5)
				C_Unit			= iCurColumnPos(6)
				C_IssuedQty		= iCurColumnPos(7)
				C_RemainQty		= iCurColumnPos(8)
				C_ConsumedQty	= iCurColumnPos(9)
				C_TTlIssueQty	= iCurColumnPos(10)
				C_RqrdDt		= iCurColumnPos(11)
				C_ResrvStatus	= iCurColumnPos(12)
				C_ResrvStatusDesc= iCurColumnPos(13)
				C_IssueMthd		= iCurColumnPos(14)
				C_IssueMthdDesc	= iCurColumnPos(15)
				C_MajorSLCd		= iCurColumnPos(16)
				C_MajorSLNm		= iCurColumnPos(17)
				C_OprNo			= iCurColumnPos(18)
				C_WcCD			= iCurColumnPos(19)
				C_ReqSeqNo		= iCurColumnPos(20)
				C_ReqNo			= iCurColumnPos(21)
				C_ParentItemCd	= iCurColumnPos(22)
				C_ParentItemNm	= iCurColumnPos(23)
				C_ParentItemSpec= iCurColumnPos(24)
				C_TrackingNo	= iCurColumnPos(25)
				C_JobNm			= iCurColumnPos(26)

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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
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

'------------------------------------------  OpenItemInfo1()  -------------------------------------------------
'	Name : OpenItemInfo1()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo1(Byval strCode)
    
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
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


'------------------------------------------  OpenItemInfo2()  -------------------------------------------------
'	Name : OpenItemInfo2()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo2(Byval strCode)
	
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call Displaymsgbox("971012","X", "����","X")
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
		Call SetItemInfo2(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

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


'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") ' Where Condition
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
Function OpenSLCd(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
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

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtReqStartDt.Text
	arrParam(4) = frm1.txtReqEndDt.Text	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus

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
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If
	
    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
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
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ProdtOrderNo
	arrParam(1) = Trim(frm1.vspdData.Text)		'��: ��ȸ ���� ����Ÿ 
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If
	
    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
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
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)			' Plant Code
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ProdtOrderNo
	arrParam(1) = Trim(frm1.vspdData.Text)		'��: ��ȸ ���� ����Ÿ 
	arrParam(2) = ""									' Operation No.
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
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
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
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
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ProdtOrderNo
	arrParam(1) = Trim(frm1.vspdData.Text)		'��: ��ȸ ���� ����Ÿ 

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
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

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo1()  --------------------------------------------------
'	Name : SetItemInfo1()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo1(Byval arrRet)
    With frm1
	.txtChildItemCd.value = arrRet(0)
	.txtChildItemNm.value = arrRet(1)
    End With
End Function

'------------------------------------------  SetItemInfo1()  --------------------------------------------------
'	Name : SetItemInfo2()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo2(Byval arrRet)
    With frm1
	.txtItemCd.value = arrRet(0)
	.txtItemNm.value = arrRet(1)
    End With
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function


'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet)

    With frm1
        .txtSLCd.value = arrRet(0)  
	   	.txtSLNm.Value = arrRet(1)	   	
	End With

End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	
	frm1.txtTrackingNo.Value = arrRet(0)
	
End Function   


'------------------------------------------  SetConWC()  --------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
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
'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtReqStartDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtReqEndDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================

Sub txtReqStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqStartDt.Focus
    End If
End Sub

Sub txtReqEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqEndDt.Focus
    End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	
 	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
 	
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
 	
' 	If Row <= 0 Then
'       ggoSpread.Source = frm1.vspdData
'       Exit Sub
'    End If

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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
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

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtCompntCd.value = "" Then
		frm1.txtCompntNm.value = "" 
	End If
	
	If frm1.txtSLCd.value = "" Then
		frm1.txtSLNm.value = "" 
	End If
	
	If ValidDateCheck(frm1.txtReqStartDt, frm1.txtReqEndDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If


    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If																'��: Query db data

    FncQuery = True																'��: Processing is OK
    
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
	On Error Resume Next
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
    Call InitSpreadComboBox
    
    ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)
    
End Sub 


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
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

	Dim strVal     
    
    DbQuery = False
  
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing
        
    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1	    
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2	  	    
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtReqStartDt=" & Trim(.hReqStartDt.Value)
		strVal = strVal & "&txtReqEndDt=" & Trim(.hReqEndDt.Value)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(.hProdOrderNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtChildItemCd=" & Trim(.hCompntCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtSLCd=" & Trim(.hSLCd.Value)						'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboProdMgr=" & Trim(.hProdMgr.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboInvMgr=" & Trim(.hInvMgr.Value)					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboJobCd=" & Trim(.hJobCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&rdoIssueFlag=" & Trim(frm1.hIssueFlag.value)
		strVal = strVal & "&cboIssueMthd=" & Trim(frm1.hIssueMthd.value)	
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
		strVal =  BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1	    
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2	      
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)				'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&txtReqStartDt=" & Trim(.txtReqStartDt.Text)
		strVal = strVal & "&txtReqEndDt=" & Trim(.txtReqEndDt.Text)
		strVal = strVal & "&txtProdtOrderNo=" & Trim(.txtProdOrderNo.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtChildItemCd=" & Trim(.txtCompntCd.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtSLCd=" & Trim(.txtSLCd.Value)					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboProdMgr=" & Trim(.cboProdMgr.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboInvMgr=" & Trim(.cboInvMgr.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboJobCd=" & Trim(.cboJobCd.Value)	'��: ��ȸ ���� ����Ÿ 
		If frm1.rdoIssueFlag1.checked = True Then
			strVal = strVal & "&rdoIssueFlag=" & frm1.rdoIssueFlag1.value 
		ElseIf frm1.rdoIssueFlag2.checked = True Then
			strVal = strVal & "&rdoIssueFlag=" & frm1.rdoIssueFlag2.value 
		Else 
			strVal = strVal & "&rdoIssueFlag=" & frm1.rdoIssueFlag3.value 
		End If
		strVal = strVal & "&cboIssueMthd=" & Trim(frm1.cboIssueMthd.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	End IF	
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)											'��: ��ȸ ������ ������� 
	
	Call SetToolBar("11000000000111")										'��: ��ư ���� ���� 
    '-----------------------
    'Reset variables area
    '-----------------------
    
    Call InitData(LngMaxRow)
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	
End Function



'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function


'========================================================================================
' Function Name : ViewHidden
' Function Desc : Show Detail Field
'========================================================================================
Function ViewHidden(StrMnuID, MnuCount, StrImageSize )
    Dim ii

    For ii = 1 To MnuCount
        If document.all(StrMnuID & ii).style.display = "" Then 
           document.all(StrMnuID & ii).style.display = "none"
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/Smallplus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigPlus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
			End Select		
        Else
           document.all(StrMnuID & ii).style.display = ""
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/SmallMinus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigMinus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
			End Select
        End If
    Next    

End Function