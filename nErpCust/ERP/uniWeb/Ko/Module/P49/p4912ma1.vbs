
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p4912mb1.asp"								'��: �����Ͻ� ����(Qeury) ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Dim C_ProdtOrderNo
Dim C_OprNo
Dim C_ReportDt
Dim C_ItemCd
Dim C_ItemNm
Dim C_ProdtOrderQty
Dim C_ProdtOrderUnit
Dim C_BadQty
Dim C_ProdQtyInOrderUnit
Dim C_GoodQty
Dim C_StdTime
Dim C_WkTime
Dim C_IdTime
Dim C_EtcTime
Dim C_WkLossTime

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2. Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntGrpCount              ' GroupView Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgSortKey
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop
Dim lstrPgmID
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
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count

End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'===========================================================================================================
Sub SetDefaultVal()
    frm1.txtFromDt.text = StartDate
    frm1.txtToDt.text   = EndDate

	If Trim(ReadCookie("txtPlantCd")) <> "" Then
		frm1.txtPlantCd.Value		= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value		= ReadCookie("txtPlantNm")
		frm1.txtWcCd.Value			= ReadCookie("txtWcCd")
		frm1.txtWcNm.value			= ReadCookie("txtWcNm")
		frm1.txtProdOrderNo.value	= ReadCookie("txtProdOrderNo")

		frm1.txtFromDt.Text			= ReadCookie("txtprodDt")		
'		frm1.txtToDt.Text			= ReadCookie("txtPlanEndDt")
		lstrPgmID = ReadCookie("txtPGMID")
	End If	

	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtWcCd", ""	
	WriteCookie "txtWcNm", ""	
	WriteCookie "txtProdOrderNo", ""	
	WriteCookie "txtFromDt", ""
	WriteCookie "txtPGMID", ""

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

    .MaxCols = C_WkLossTime + 1
    .MaxRows = 0

	Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_ProdtOrderNo,		"����������ȣ", 14
    ggoSpread.SSSetEdit		C_OprNo,			"������ȣ", 12
    ggoSpread.SSSetDate		C_ReportDt,			"�۾���", 11, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_ItemCd,			"ǰ��", 12
	ggoSpread.SSSetEdit		C_ItemNm,			"ǰ���", 16

	ggoSpread.SSSetFloat	C_ProdtOrderQty,	"��������", 12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_ProdtOrderUnit,	"��������", 10
	ggoSpread.SSSetFloat	C_BadQty,			"�ҷ���", 12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit, "���Լ�", 12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_GoodQty,			"�ϼ���", 12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

	ggoSpread.SSSetTime		C_StdTime,			"ǥ�ذ���", 12, 2, 1, 1
	ggoSpread.SSSetTime		C_WkTime,			"�۾�����", 12, 2, 1, 1
	ggoSpread.SSSetTime		C_IdTime,			"��������", 12, 2, 1, 1
	ggoSpread.SSSetTime		C_EtcTime,			"��Ÿ����", 12, 2, 1, 1
	ggoSpread.SSSetTime		C_WkLossTime,		"���ǰ���", 12, 2, 1, 1

	'Call ggoSpread.MakePairsColumn(,)
	Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
	ggoSpread.SSSetSplit2(2)											'frozen ��� �߰� 

	.ReDraw = true

	Call SetSpreadLock

    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		.vspdData.ReDraw = True
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column
'=========================================================================================================
Sub InitSpreadPosVariables()

	C_ProdtOrderNo			= 1
	C_OprNo					= 2
	C_ReportDt				= 3
	C_ItemCd				= 4
	C_ItemNm				= 5
	C_ProdtOrderQty			= 6
	C_ProdtOrderUnit		= 7
	C_BadQty				= 8
	C_ProdQtyInOrderUnit	= 9
	C_GoodQty				= 10
	C_StdTime				= 11
	C_WkTime				= 12
	C_IdTime				= 13
	C_EtcTime				= 14
	C_WkLossTime			= 15

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

  		C_ProdtOrderNo			= iCurColumnPos(1)
  		C_OprNo					= iCurColumnPos(2)
  		C_ReportDt				= iCurColumnPos(3)
  		C_ItemCd				= iCurColumnPos(4)
  		C_ItemNm				= iCurColumnPos(5)
  		C_ProdtOrderQty			= iCurColumnPos(6)
  		C_ProdtOrderUnit		= iCurColumnPos(7)
  		C_BadQty				= iCurColumnPos(8)
  		C_ProdQtyInOrderUnit	= iCurColumnPos(9)
  		C_GoodQty				= iCurColumnPos(10)
  		C_StdTime				= iCurColumnPos(11)
  		C_WkTime				= iCurColumnPos(12)
  		C_IdTime				= iCurColumnPos(13)
  		C_EtcTime				= iCurColumnPos(14)
  		C_WkLossTime			= iCurColumnPos(15)

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

'------------------------------------------  OpenProdOrderNo()  -----------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------------
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
	arrParam(4) = "CL"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(7) = ""	'Trim(frm1.txtItemCd.value)
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
	arrParam(4) = "PLANT_CD = " & FilterVar(Ucase(Trim(frm1.txtPlantCd.value)),"''","S") 			' Where Condition
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

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetProdOrderNo()  ----------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)
End Function

'------------------------------------------  SetConWC()  ----------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------------
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

' 	If NewCol = C_XXX or Col = C_XXX Then
' 		Cancel = True
' 		Exit Sub
' 	End If
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
'   Event Name : txtPlantCd_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtPlantCd_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
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

'=======================================================================================================
'   Event Name : txtWcCd_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtWcCd_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtProdOrderNo_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtProdOrderNo_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : DefaultSumValue
' Function Desc : 
'========================================================================================
Sub DefaultSumValue()
	With frm1
		.fpDoubleSingle1.Value = 0
		.fpDoubleSingle2.Value = 0
		.fpDoubleSingle3.Value = 0
		.fpDoubleSingle4.Value = "00:00:00"
		.fpDoubleSingle5.Value = "00:00:00"
		.fpDoubleSingle6.Value = "00:00:00"
		.fpDoubleSingle7.Value = "00:00:00"
		.fpDoubleSingle8.Value = "00:00:00"
	End With
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

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If

	If frm1.txtWCCd.value = "" Then
		frm1.txtWCNm.value = ""
	End If

	If ValidDateCheck(frm1.txtFromDt, frm1.txtTODt) = False Then Exit Function

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	Call DefaultSumValue()
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then											'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Exit Function																'��: Query db data
	End If

    FncQuery = True																'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()
    On Error Resume Next                                                    '��: Protect system from crashing
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
    On Error Resume Next                                                    '��: Protect system from crashing
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
    On Error Resume Next                                                    '��: Protect system from crashing
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
    Call parent.FncExport(parent.C_MULTI)											'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)                                     '��:ȭ�� ����, Tab ���� 
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
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.value)
		strVal = strVal & "&txtFromDt=" & .hFromDt.value
		strVal = strVal & "&txtToDt=" & .hToDt.value
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.value)
		strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.value)
		strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
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
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 

	Call SetToolbar("11000000000111")										'��: ��ư ���� ���� 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
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

