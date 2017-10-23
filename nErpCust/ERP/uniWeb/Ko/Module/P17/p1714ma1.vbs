
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID	= "p1714mb1.asp"								'��: �����Ͻ� ����(Qeury) ASP�� 
Const BIZ_PGM_SAVE_ID	= "p1714mb2.asp"								'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID2	= "p1714mb3.asp"								'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_Select
Dim C_ReqTransNo					'�̰���û��ȣ 
Dim C_DestPlantCd					'������ 
Dim C_DestPlantNm					'������� 
Dim C_BasePlantCd					'������� 
Dim C_BasePlantNm					'�������� 
Dim C_ItemCd						'��ǰ�� 
Dim C_ItemNm						'��ǰ��� 
Dim C_Spec							'�԰� 
Dim C_ReqTransDt					'�̰���û�� 
Dim C_TransDt						'�̰��� 
Dim C_BomDesc						'BOM���� 
Dim C_ValidFromDt					'��ȿ�� 
Dim C_ValidToDt						'��ȿ�� 
Dim C_DrawingPath					'������ 
Dim C_TransStatus					'�̰����� 
Dim C_BomNo							'BOM
Dim C_MajorFlg						'�� BOM ���� 
Dim C_ReturnDesc					'�ݷ����� 

'Dim C_Row

Dim isClicked
Dim iCol
Dim iRow
Dim IsOpenPop
Dim iStrFree

Dim lgButtonSelection
Dim lgRedrewFlg
Dim gbtnAuto

Dim gFlg

'========================================================================================================
' Name : InitSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
    C_Select				= 1
    C_ReqTransNo			= 2 	'�̰���û��ȣ 
    C_DestPlantCd			= 3 	'������ 
    C_DestPlantNm			= 4 	'������� 
    C_BasePlantCd			= 5 	'������� 
    C_BasePlantNm			= 6 	'�������� 
    C_ItemCd				= 7 	'��ǰ�� 
    C_ItemNm				= 8 	'��ǰ��� 
    C_Spec					= 9 	'�԰� 
    C_ReqTransDt			= 10	'�̰���û�� 
    C_TransDt				= 11	'�̰��� 
    C_BomDesc				= 12	'BOM���� 
    C_ValidFromDt			= 13	'��ȿ�� 
    C_ValidToDt				= 14	'��ȿ�� 
    C_DrawingPath			= 15	'������ 
    C_TransStatus			= 16	'�̰����� 
    C_BomNo					= 17	'BOM
    C_MajorFlg				= 18	'��BOM���� 
    C_ReturnDesc			= 19	'�ݷ����� 

'    C_Row	                = 18

End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKeyIndex = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                               '��: initializes sort direction

	lgButtonSelection = "DESELECT"
	With frm1
		.btnAutoSel1.disabled = True
		.btnAutoSel1.value = "��ü����"
		.btnAutoSel2.disabled = True
		.btnAutoSel3.disabled = True
	End With
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	With frm1
		.btnAutoSel1.disabled = True
		.btnAutoSel1.value = "��ü����"
		.btnAutoSel2.disabled = True
		.btnAutoSel3.disabled = True
	End With
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	'============================================================================================
	'��: Spreadsheet vspdData
	'============================================================================================

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData

		ggoSpread.Spreadinit "V20050130",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_ReturnDesc + 1
		.MaxRows = 0

		Call GetSpreadColumnPos()

		ggoSpread.SSSetCheck		C_Select,		"", 2,,,1

		ggoSpread.SSSetEdit			C_ReqTransNo,			"�̰��Ƿڹ�ȣ", 18
		ggoSpread.SSSetEdit			C_DestPlantCd,			"������", 10,,,4,2
		ggoSpread.SSSetEdit			C_DestPlantNm,			"�������", 15

		ggoSpread.SSSetEdit 		C_BasePlantCd,			"�������", 10,,,4,2
		ggoSpread.SSSetEdit 		C_BasePlantNm,			"��������", 15
		ggoSpread.SSSetEdit			C_ItemCd,				"��ǰ��", 12,,,18,2
		ggoSpread.SSSetEdit			C_ItemNm,				"��ǰ���", 20
		ggoSpread.SSSetEdit			C_Spec,					"�԰�", 20
		ggoSpread.SSSetDate 		C_ReqTransDt,			"�̰���û��", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate			C_TransDt,				"�̰���", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit			C_BomDesc,				"BOM����", 30,,, 100
 		ggoSpread.SSSetDate			C_ValidFromDt,			"������", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate			C_ValidToDt,			"������", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit			C_DrawingPath,			"������", 30,,, 100
		ggoSpread.SSSetEdit			C_TransStatus,			"�̰�����", 10

		ggoSpread.SSSetEdit			C_BomNo,				"BOM", 10
		ggoSpread.SSSetEdit			C_MajorFlg,				"��BOM����", 10
		ggoSpread.SSSetEdit			C_ReturnDesc,			"�ݷ�����", 30,,, 100

		Call ggoSpread.SSSetColHidden(C_BomNo, C_MajorFlg, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True

    End With

	Call SetSpreadLock
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()

    With frm1
    	ggoSpread.Source = frm1.vspdData

		.vspdData.ReDraw = False

		ggoSpread.SSSetProtected	-1, -1

		ggoSpread.SpreadUnLock	C_Select, -1, C_Select

		.vspdData.ReDraw = True

    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc :
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal Level, ByVal QueryStatus)

	ggoSpread.Source = frm1.vspdData

    frm1.vspdData.ReDraw = False

    frm1.vspdData.ReDraw = True

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    C_Select				= iCurColumnPos(1)
    C_ReqTransNo			= iCurColumnPos(2) 	'�̰���û��ȣ 
    C_DestPlantCd			= iCurColumnPos(3) 	'������ 
    C_DestPlantNm			= iCurColumnPos(4) 	'������� 
    C_BasePlantCd			= iCurColumnPos(5) 	'������� 
    C_BasePlantNm			= iCurColumnPos(6) 	'�������� 
    C_ItemCd				= iCurColumnPos(7) 	'��ǰ�� 
    C_ItemNm				= iCurColumnPos(8) 	'��ǰ��� 
    C_Spec					= iCurColumnPos(9) 	'�԰� 
    C_ReqTransDt			= iCurColumnPos(10)	'�̰���û�� 
    C_TransDt				= iCurColumnPos(11)	'�̰��� 
    C_BomDesc				= iCurColumnPos(12)	'BOM���� 
    C_ValidFromDt			= iCurColumnPos(13)	'��ȿ�� 
    C_ValidToDt				= iCurColumnPos(14)	'��ȿ�� 
    C_DrawingPath			= iCurColumnPos(15)	'������ 
    C_TransStatus			= iCurColumnPos(16)	'�̰����� 
    C_BomNo					= iCurColumnPos(17)	'BOM
    C_MajorFlg				= iCurColumnPos(18)	'��BOM���� 
    C_ReturnDesc			= iCurColumnPos(19)	'�ݷ����� 

'    C_Row	                = iCurColumnPos(18)

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   :
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
'
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
'
'	Dim i, iStrArr, iStrNmArr
'    Dim strCbo
'    Dim strCboCd
'    Dim strCboNm
'	'****************************
'    'List Minor code(�����󱸺�)
'    '****************************
'    'strCboCd = "" & vbTab & ""
'    'strCboNm = "" & vbTab
'
'	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2201", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
'    iStrArr = Split(lgF0, Chr(11))
'    iStrNmArr = Split(lgF1, Chr(11))
'
'	If Err.number <> 0 Then
'		MsgBox Err.Description
'		Err.Clear
'		Exit Sub
'	End If
'
'	For i = 0 to UBound(iStrArr) - 1
'        strCboCd = strCboCd & iStrArr(i) & vbTab
'        strCboNm = strCboNm & iStrNmArr(i) & vbTab
'	Next
'
'	iStrFree = iStrNmArr(1)
'
'    ggoSpread.SetCombo strCboCd, C_SupplyFlg 'parent.ggoSpread.SSGetColsIndex()              'Supply Flag setting
'    ggoSpread.SetCombo strCboNm, C_SupplyFlgNm 'parent.ggoSpread.SSGetColsIndex()            'Supply Flag Nm Setting
'
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'=========================================================================================================
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	If frm1.vspdData.MaxRows <= 1 Then Exit Sub
	If lngStartRow = 1 Then lngStartRow = 2

	With frm1.vspdData

'		.ReDraw = False
'		If lgStrBOMHisFlg = "Y" Then
'			ggoSpread.SpreadUnLock	C_ECNNo,		lngStartRow, C_ECNNo, .MaxRows
'			ggoSpread.SpreadUnLock	C_ECNNoPopup,	lngStartRow, C_ECNNoPopup, .MaxRows
'			ggoSpread.SSSetRequired	C_ECNNo,		lngStartRow, .MaxRows
'			ggoSpread.SpreadUnLock	C_ECNDesc,		lngStartRow, C_ECNDesc, .MaxRows
'			ggoSpread.SSSetRequired	C_ECNDesc,		lngStartRow, .MaxRows
'			ggoSpread.SpreadUnLock	C_ReasonCd,		lngStartRow, C_ReasonCd, .MaxRows
'			ggoSpread.SpreadUnLock	C_ReasonCdPopup,	lngStartRow, C_ReasonCdPopup, .MaxRows
'			ggoSpread.SSSetRequired	C_ReasonCd,		lngStartRow, .MaxRows
'
'		Else
'			ggoSpread.SSSetProtected C_ECNNo,		lngStartRow, .MaxRows
'			ggoSpread.SSSetProtected C_ECNNoPopup,	lngStartRow, .MaxRows
'			ggoSpread.SSSetProtected C_ECNDesc,		lngStartRow, .MaxRows
'			ggoSpread.SSSetProtected C_ReasonCd,	lngStartRow, .MaxRows
'			ggoSpread.SSSetProtected C_ReasonCdPopup,lngStartRow, .MaxRows
'
'		End If
'
'		.ReDraw = True
	End With
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Design Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConBasePlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������˾�"									' �˾� ��Ī 
	arrParam(1) = "B_PLANT A, P_PLANT_CONFIGURATION B"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBasePlantCd.value)						' Code Condition
	arrParam(3) = ""													' Name Cindition
	arrParam(4) = "B.PLANT_CD = A.PLANT_CD AND B.ENG_BOM_FLAG = 'Y'"	' Where Condition
	arrParam(5) = "�������"				' TextBox ��Ī 

    arrField(0) = "A.PLANT_CD"					' Field��(0)
    arrField(1) = "A.PLANT_NM"					' Field��(1)

    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConBasePlant(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtDestPlantCd.focus

End Function


'------------------------------------------  OpenCondPlant2()  -------------------------------------------------
'	Name : OpenCondDestPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConDestPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�������˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtDestPlantCd.value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "������"					' TextBox ��Ī 

    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)

    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConDestPlant(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd(ByVal str, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If

	If frm1.txtDestPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "������", "X")
		frm1.txtDestPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtDestPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(str)	' Item Code

	arrField(0) = 1		'ITEM_CD
    arrField(1) = 2 	'ITEM_NM
    arrField(2) = 5		'ITEM_ACCT
    arrField(3) = 9 	'PROC_TYPE
    arrField(4) = 4 	'BASIC_UNIT
    arrField(5) = 51	'SINGLE_ROUT_FLG
    arrField(6) = 52	'Major_Work_Center
    arrField(7) = 13	'Phantom_flg
    arrField(8) = 18	'valid_from_dt
    arrField(9) = 19	'valid_to_dt
    arrField(10) = 3	' Field��(1) : "SPECIFICATION"

	iCalledAspName = AskPRAspName("B1B11PA4")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet, iPos)
	End If

	If iPos = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtItemCd.focus
	Else
		Call SetActiveCell(frm1.vspdData,C_ChildItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End If

End Function

'------------------------------------------  SetConBasePlant()  ----------------------------------------------
'	Name : SetConBasePlant()
'	Description : Condition Base Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConBasePlant(byval arrRet)
	frm1.txtBasePlantCd.Value    = arrRet(0)
	frm1.txtBasePlantNm.Value    = arrRet(1)

	Call txtBasePlantCd_OnChange()
End Function

'------------------------------------------  SetConDestPlant()  ----------------------------------------------
'	Name : SetConDestPlant()
'	Description : Condition Destination Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConDestPlant(byval arrRet)
	frm1.txtDestPlantCd.Value    = arrRet(0)
	frm1.txtDestPlantNm.Value    = arrRet(1)

	Call txtDestPlantCd_OnChange()
End Function

'------------------------------------------  SetItemCd()  ----------------------------------------------
'	Name : SetItemCd()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byval arrRet, ByVal iPos)
	If iPos = 0 Then
		frm1.txtItemCd.Value    = arrRet(0)
		frm1.txtItemNm.Value    = arrRet(1)
	Else
		With frm1.vspdData
			.Col = C_ChildItemCd
			.Row = .ActiveRow
			.Text = arrRet(0)

			Call LookUpItemByPlant(arrRet(0), .Row)

		End With

		lgBlnFlgChgValue = True
	End IF

End Function

'==========================================================================================
'   Function Name :SetFieldProp
'   Function Desc :���� Case�� ���� Field���� �Ӽ��� �����Ѵ�.
'==========================================================================================

Function SetFieldProp(ByVal lRow, ByVal Level, ByVal ProcType)

End Function


Sub SetCookieVal()
	
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm")
	End If	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtBasePlantCd.Value = ReadCookie("txtPlantCd")
		frm1.txtBasePlantNm.value = ReadCookie("txtPlantNm")
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""

End Sub

'==========================================================================================
'   Function Name : SetSelectAll
'   Function Desc : ��ü���� 
'==========================================================================================
Function btnAutoSel1_onClick()

	lgRedrewFlg = False

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		With frm1
			.btnAutoSel1.value = "��ü����"
		End With
	Else
		lgButtonSelection = "SELECT"
		With frm1
			.btnAutoSel1.value = "��ü�������"
		End With
	End If

	Dim index,Count
	Dim strFlag

	frm1.vspdData.ReDraw = False

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
			frm1.vspdData.Text=""
		End if
	Next

	frm1.btnAutoSel2.disabled = False
	frm1.btnAutoSel3.disabled = False

	frm1.vspdData.ReDraw = True

	lgRedrewFlg = True
End Function

'==========================================================================================
'   Function Name :
'   Function Desc : �̰� 
'==========================================================================================
Function btnAutoSel2_onClick()
	Dim IntRetCD
	Dim index, Count
	Dim strFlag

	strFlag = ""

	frm1.vspdData.ReDraw = False

	Count = frm1.vspdData.MaxRows

	For index = 1 to Count
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Select
		strFlag = frm1.vspdData.Value

		If strFlag = 1 Then Exit For
	Next

	frm1.vspdData.ReDraw = True

	If strFlag = 1 Then
        IntRetCD = DisplayMsgBox("P17141", parent.VB_YES_NO, "x", "x")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		Else
			gbtnAuto = "btnPrc"
			frm1.hgubun.Value = gbtnAuto
			frm1.hStartDate.Value = StartDate
			If DbSave = False Then
				Exit Function
			End If
		End If
	Else
		ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer
		If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
			IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
			Exit Function
		End If
	End If
End Function

'==========================================================================================
'   Function Name :
'   Function Desc : �ݷ� 
'==========================================================================================
Function btnAutoSel3_onClick()
	Dim index, Count
	Dim strFlag

	frm1.vspdData.ReDraw = False

	Count = frm1.vspdData.MaxRows

	For index = 1 to Count
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Select
		strFlag = frm1.vspdData.Value

		If strFlag = 1 Then Exit For
	Next

	frm1.vspdData.ReDraw = True

	If strFlag = 1 Then
        IntRetCD = DisplayMsgBox("P17142", parent.VB_YES_NO, "x", "x")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		Else
			gbtnAuto = "btnCancel"
			frm1.hgubun.Value = gbtnAuto
			gFlg = "check"
			If DbSave = False Then
				Exit Function
			End If
		End If
	Else
		ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer
		If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
			IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
			Exit Function
		End If
	End If
End Function

'==========================================================================================
'   Function Name :LookUpItemByPlant
'   Function Desc :������ ǰ���� Item Acct�� �д´�.
'==========================================================================================
Sub LookUpItemByPlant(ByVal strItemCd, ByVal IRow)

    Err.Clear															'��: Protect system from crashing

	Dim strSelect
	If strItemCd = "" Then Exit Sub

	frm1.vspdData.Col = C_ChildItemCd
	frm1.vspdData.Row = IRow

	strSelect = " b.ITEM_NM, a.ITEM_ACCT, dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ", a.ITEM_ACCT) ITEM_ACCT_NM, a.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", a.PROCUR_TYPE) PROCUR_TYPE_NM, b.SPEC, b.BASIC_UNIT, dbo.ufn_GetItemAcctGrp(a.ITEM_ACCT) ITEM_ACCT_GRP "

	If 	CommonQueryRs2by2(strSelect, " B_ITEM_BY_PLANT a, B_ITEM b ", " a.ITEM_CD = b.ITEM_CD AND a.PLANT_CD = " & _
	    FilterVar(frm1.txtDestPlantCd.Value, "''", "S") & " AND a.ITEM_CD = " & FilterVar(strItemCd, "''", "S"), lgF0) = False Then
		Call DisplayMsgBox("122700", "X", strItemCd, "X")
		Call LookUpItemByPlantNotOk()
		Exit Sub
	End If

	lgF0 = Split(lgF0, Chr(11))

	Call LookUpItemByPlantOk(lgF0(1), lgF0(2), lgF0(3), lgF0(4), lgF0(5), lgF0(6), lgF0(7), IRow, lgF0(8))
End Sub

'==========================================================================================
'   Function Name :LookUpItemByPlantOk
'   Function Desc :������ ǰ���� ���翩�θ� Check�Ը� �д´�.
'==========================================================================================
Function LookUpItemByPlantOk(ByVal strItemNm, ByVal strItemAcct, ByVal strItemAcctNm, ByVal strProcType, ByVal strProcTypeNm, ByVal strSpec, ByVal strBasicUnit, ByVal IRow , ByVal strItemAcctGrp)
End Function

Function LookUpItemByPlantNotOk()
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
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
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

' 	If NewCol = C_Select or Col = C_Select Then
' 		Cancel = True
' 		Exit Sub
' 	End If

     ggoSpread.Source = frm1.vspdData
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos()
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange
'   Event Desc :Combo Change Event
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

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
					ggoSpread.SpreadUnLock	C_ReturnDesc, Row, C_Select, Row
					ggoSpread.UpdateRow Row
				Else
					.Col = C_ReturnDesc
					.Text = ""
					ggoSpread.SpreadLock	C_ReturnDesc, Row, C_Select, Row
					ggoSpread.EditUndo Row
				End If

				If lgRedrewFlg = True Then .ReDraw = True
		End Select
	End With
	frm1.btnAutoSel2.disabled = False
	frm1.btnAutoSel3.disabled = False
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

'-------------------------------------------------------------------------------
' Function Name : txtDestPlantCd_OnChange()
' Function Desc :
'-------------------------------------------------------------------------------
Sub txtDestPlantCd_OnChange()
	Dim strPlant

	ggoSpread.Source = frm1.vspdData

	If Trim(frm1.txtDestPlantCd.value) <> "" Then
		Call CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtDestPlantCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		strPlant = Replace(lgF0, Chr(11), "")

		If Trim(strPlant) = "" Then
			frm1.txtDestPlantNm.Value = ""
		Else
			frm1.txtDestPlantNm.Value = Trim(strPlant)
			frm1.txtItemCd.focus
		End If
	End If
End Sub
'-------------------------------------------------------------------------------
' Function Name : txtBasePlantCd_OnChange()
' Function Desc : Design Plant
'-------------------------------------------------------------------------------
Sub txtBasePlantCd_OnChange()
	Dim strPlant

	If Trim(frm1.txtBasePlantCd.value) <> "" Then
'		Call CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtBasePlantCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call CommonQueryRs("A.PLANT_NM", "B_PLANT A, P_PLANT_CONFIGURATION B", "B.PLANT_CD = A.PLANT_CD AND B.ENG_BOM_FLAG = 'Y' AND A.PLANT_CD = " & FilterVar(frm1.txtBasePlantCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		strPlant = Replace(lgF0, Chr(11), "")

		If Trim(strPlant) = "" Then
			frm1.txtBasePlantNm.Value = ""
		Else
			frm1.txtBasePlantNm.Value = Trim(strPlant)
		End If
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False                                                        '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData										'��: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then									'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")		'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	If frm1.txtDestPlantCd.value = "" Then
		frm1.txtDestPlantNm.value = ""
	End If

	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If

	If frm1.txtBasePlantCd.value = "" Then
		frm1.txtBasePlantNm.value = ""
	End If

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
    If Not chkfield(Document, "1") Then											'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Exit Function																'��: Query db data
	End If

    FncQuery = True															'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()
	Dim IntRetCD
    Dim slPlantCd
    Dim slPlantNm

    FncNew = False                                                          '��: Processing is NG

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    slPlantCd = frm1.txtBasePlantCd.value
    slPlantNm = frm1.txtBasePlantNm.value

    Call ggoOper.ClearField(Document, "A")                                         '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field

    frm1.txtBasePlantCd.value = slPlantCd
    frm1.txtBasePlantNm.value = slplantNm

    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call txtBasePlantCd_OnChange()

    Call SetToolbar("11101101001011")

    frm1.txtDestPlantCd.focus
    Set gActiveElement = document.activeElement

    FncNew = True                                                           '��: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                      '��: Processing is NG

    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '��:
        Exit Function
    End If

    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------
		Exit Function
	End If


    If DbDelete = False Then
		Exit Function
    End If     						'��: Delete db data

    FncDelete = True                                                        '��: Processing is OK

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave()
    Dim IntRetCD

    FncSave = False                                                         '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '��: No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------

    If Not chkField(Document, "2") Then
       Exit Function
    End If

'	ggoSpread.Source = frm1.vspdData
'    If Not ggoSpread.SSDefaultCheck("Y") Then                                  '��: Delete�� Logic ����(������)
'       Exit Function
'    End If

    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer
    If Not ggoSpread.SSDefaultCheck         Then				'��: Check required field(Multi area)
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		Exit Function
    End If     				                                                  '��: Save db data

    FncSave = True                                                          '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
	Dim iRow
	Dim strPrntLevel
	Dim strLevel
	Dim Level
	Dim PrntItemCd
	Dim PrntItemAcct
	Dim PrntBomNo
	Dim PrntProcType
	Dim PrntBasicUnit
	Dim PrntItemAcctGrp
	Dim i

    If Not chkField(Document, "2") Then									'��: This function check indispensable field
       Exit Function
    End If

    With frm1

		.vspdData.focus
		Set gActiveElement = document.activeElement

		ggoSpread.Source = .vspdData

		.vspdData.ReDraw = False

		iRow =  .vspdData.ActiveRow

'		Call .vspdData.GetText(C_Level, iRow, strPrntLevel)
'		Call .vspdData.GetText(C_ChildItemCd, iRow, PrntItemCd)
'		Call .vspdData.GetText(C_ItemAcct, iRow, PrntItemAcct)
'		Call .vspdData.GetText(C_BomType, 1, PrntBomNo)
'		Call .vspdData.GetText(C_ProcType, iRow, PrntProcType)
'		Call .vspdData.GetText(C_ItemAcctGrp, iRow, PrntItemAcctGrp)
'
		If frm1.rdoSrchType1.checked = True And strPrntLevel <> "0" Then					'�ܴܰ��̸� 
			Call DisplayMsgBox("182722", "X", "X", "X")
			Exit Function
		End If

		If Not(PrntItemAcctGrp = "1FINAL"  Or PrntItemAcctGrp = "2SEMI") And PrntBomNo = "1" Then
			Call DisplayMsgBox("182618", "X", "X", "X")
			Exit Function
		End If

'		Call .vspdData.GetText(C_ChildItemUnit, iRow, PrntBasicUnit)

		If strPrntLevel = "" Then
			strLevel = ".1"
			level = 1
		ElseIf iRow < 1 Then
			strLevel = "0"
			Level = 0
		Else
			Level = Replace(strPrntLevel, ".","") + 1

			For i = 1 To Level
				strLevel = strLevel & "."
			Next

			strLevel = strLevel & Level
		End If


		If frm1.vspdData.maxrows < 1 Then Exit Function

		frm1.vspdData.focus
		Set gActiveElement = document.activeElement
		frm1.vspdData.EditMode = True

		ggoSpread.Source = frm1.vspdData

		ggoSpread.CopyRow

		Call SetSpreadColor(iRow + 1, iRow + 1, Level, 0)

    '------------------------------------------------------
    ' Default Value Setting
    '------------------------------------------------------
		iRow = .vspdData.ActiveRow
'		Call .vspdData.SetText(C_Level,			iRow, strLevel)
'		Call .vspdData.SetText(C_Seq,			iRow, "")
'		Call .vspdData.SetText(C_ChildItemCd,	iRow, "")
'		Call .vspdData.SetText(C_ChildItemNm,	iRow, "")
'		Call .vspdData.SetText(C_Spec,			iRow, "")
'		Call .vspdData.SetText(C_ItemAcctNm,	iRow, "")
'		Call .vspdData.SetText(C_ProcTypeNm,	iRow, "")
'		Call .vspdData.SetText(C_BomType,		iRow, "")
'		Call .vspdData.SetText(C_DrawingPath,	iRow, "")
'		Call .vspdData.SetText(C_PrntBasicUnit,	iRow, PrntBasicUnit)
'		Call .vspdData.SetText(C_HdrItemCd,		iRow, PrntItemCd)
'		Call .vspdData.SetText(C_HdrBomNo,		iRow, PrntBomNo)

		.vspdData.Col = C_SupplyFlg
		If .vspdData.text = "" Then
			.vspdData.text = "F"
		End If

'		Call .vspdData.SetText(C_ValidFromDt, iRow, "")
'		Call .vspdData.SetText(C_ValidToDt, iRow, UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31"))

		If lgStrBOMHisFlg = "Y" Then
'			Call .vspdData.SetText(C_ECNNo,	iIntCnt, frm1.txtECNNo.value)
'			Call .vspdData.SetText(C_ECNDesc,	iIntCnt, frm1.txtECNDesc.value)
'			Call .vspdData.SetText(C_ReasonCd,	iIntCnt, frm1.txtReasonCd.value)
'			Call .vspdData.SetText(C_ReasonNm,	iIntCnt, frm1.txtReasonNm.value)
		End If

'		If Trim(PrntProcType) = "O" Then					'����ǰ���� ���ְ���ǰ�� ��� 
'			ggoSpread.SpreadUnLock		C_SupplyFlgNm,		iRow+1, C_SupplyFlgNm,iRow+1
'			ggoSpread.SSSetRequired		C_SupplyFlgNm,		iRow+1, iRow+1
'		Else
'			ggoSpread.SSSetProtected	C_SupplyFlgNm,		iRow+1, iRow+1
'		End If

'		Call InitData(.vspdData.ActiveRow)

		.vspdData.ReDraw = True

    End With

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	Dim strLevel, strChildLevel
	Dim TempChildLevel

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_Level
		strLevel = Clng(Replace(.Text, ".", ""))

		Do
			ggoSpread.EditUndo
			'.Col = C_ECNNo
			'Call LookupECN(.Text, 1)	'2003-09-13
			If .MaxRows = 0 Then Exit Do

			.Col = C_Level
			.Row = .ActiveRow
			If Trim(.Text) = "" Then
				strChildLevel = Clng(0)
			Else
				strChildLevel = Clng(Replace(Trim(.Text) , ".", ""))
			End If
		Loop While (strLevel < strChildLevel)
    End With
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim iIntReqRows, iIntCnt
	Dim iRow
	Dim strPrntLevel
	Dim strLevel
	Dim Level
	Dim PrntItemCd
	Dim PrntBomNo
	Dim PrntItemAcct
	Dim PrntProcType
	Dim PrntItemAcctGrp
	Dim i
	Dim PrntBasicUnit

    On Error Resume Next
    Err.Clear

    FncInsertRow = False                                                         '��: Processing is NG
	iIntReqRows = 1

	If Trim(frm1.txtDestPlantCd.value) = "" Then
		Call DisplayMsgBox("189220", "X", "X", "X")
		Exit Function
	End If

    If Not chkField(Document, "2") Then									'��: This function check indispensable field
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If

    With frm1

		.vspdData.Focus
		Set gActiveElement = document.activeElement

		ggoSpread.Source = .vspdData

		iRow = .vspdData.ActiveRow

'		Call .vspdData.GetText(C_Level, iRow, strPrntLevel)

		If iRow >= 1 Then

			If IsNumeric(Trim(pvRowCnt)) Then
				iIntReqRows = CInt(pvRowCnt)
			Else
				iIntReqRows = AskSpdSheetAddRowCount()
				If iIntReqRows = "" Then
				    Exit Function
				End If
			End If

'			Call .vspdData.GetText(C_ChildItemCd, iRow, PrntItemCd)
'			Call .vspdData.GetText(C_ItemAcct, iRow, PrntItemAcct)
'			Call .vspdData.GetText(C_BomType, 1, PrntBomNo)
'			Call .vspdData.GetText(C_ProcType, iRow, PrntProcType)
'			Call .vspdData.GetText(C_ItemAcctGrp, iRow, PrntItemAcctGrp)

'			If strPrntLevel <> "0" Then					'�ܴܰ��̸� 
'				Call DisplayMsgBox("182722", "X", "X", "X")
'				Exit Function
'			End If
'
'			If Not (PrntItemAcctGrp = "1FINAL" Or PrntItemAcctGrp = "2SEMI")  And PrntBomNo = "1" Then
'				Call DisplayMsgBox("182618", "X", "X", "X")
'				Exit Function
'			End If
'
'			Call .vspdData.GetText(C_ChildItemUnit, iRow, PrntBasicUnit)
'
'			If strPrntLevel = "" Then
'				strLevel = ".1"
'				level = 1
'			Else
'				Level = Replace(strPrntLevel, ".","") + 1
'
'				For i = 1 To Level
'					strLevel = strLevel & "."
'				Next
'
'				strLevel = strLevel & Level
'			End If
		Else
'			strLevel = "0"
'			PrntBomNo = UCase(Trim(frm1.txtBomNo.value))	'2003-09-08
		End If

        ggoSpread.InsertRow , iIntReqRows

		.vspdData.EditMode = True
		.vspdData.ReDraw = False

		If lgIntFlgMode = parent.OPMD_CMODE And iRow < 1 Then
			Call SetSpreadColor(1, 1, 0, 0)

'			.vspdData.Col = C_Level
'			.vspdData.Text = strLevel
'			Call .vspdData.SetText(C_BomType,.vspdData.ActiveRow, PrntBomNo)	'2003-09-08
'			Call .vspdData.SetText(C_HdrBomNo,.vspdData.ActiveRow, PrntBomNo)	'2003-09-08
		Else
			For iIntCnt = iRow + 1 To iRow + iIntReqRows
				.vspdData.Row = iIntCnt

'				Call .vspdData.SetText(C_Level,			iIntCnt, strLevel)
'				Call .vspdData.SetText(C_PrntBasicUnit,	iIntCnt, PrntBasicUnit)
'				Call .vspdData.SetText(C_HdrItemCd,		iIntCnt, PrntItemCd)
'				Call .vspdData.SetText(C_HdrBomNo,		iIntCnt, PrntBomNo)
'				Call .vspdData.SetText(C_BomType,		iIntCnt, PrntBomNo)
'				Call .vspdData.SetText(C_HdrProcType,	iIntCnt, PrntProcType)
'				Call .vspdData.SetText(C_SupplyFlg,		iIntCnt, "F")
'				Call .vspdData.SetText(C_ValidFromDt,	iIntCnt, StartDate)
'				Call .vspdData.SetText(C_ValidToDt,		iIntCnt, UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"))
''				If lgStrBOMHisFlg = "Y" Then
''					Call .vspdData.SetText(C_ECNNo,	iIntCnt, frm1.txtECNNo.value)
''					Call .vspdData.SetText(C_ECNDesc,	iIntCnt, frm1.txtECNDesc.value)
''					Call .vspdData.SetText(C_ReasonCd,	iIntCnt, frm1.txtReasonCd.value)
''					Call .vspdData.SetText(C_ReasonNm,	iIntCnt, frm1.txtReasonNm.value)
''				End If

			Next

'			Call SetSpreadColor(iRow + 1, iRow + iIntReqRows, Level, 0)

'			For i = iRow + 1 To iRow + iIntReqRows
'				Call .vspdData.SetText(C_SupplyFlgNm, i, iStrFree)
'			Next
'
'			If Trim(PrntProcType)= "O" Then					'����ǰ���� ���ְ���ǰ�� ��� 
'				ggoSpread.SpreadUnLock C_SupplyFlgNm,	iRow + 1, C_SupplyFlgNm, iRow + iIntReqRows
'				ggoSpread.SSSetRequired	C_SupplyFlgNm,	iRow + 1, iRow + iIntReqRows
'			Else
'				ggoSpread.SSSetProtected C_SupplyFlgNm,	iRow + 1 , iRow + iIntReqRows
'			End If

		End If

		.vspdData.ReDraw = True

	End With
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
    Dim iIntCnt
    Dim iStrFlag, iStrLevel

    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	For iIntCnt = frm1.vspdData.SelBlockRow  To frm1.vspdData.SelBlockRow2
'		Call frm1.vspdData.GetText(C_Level, iIntCnt, iStrLevel)
'		Call frm1.vspdData.GetText(0, iIntCnt, iStrFlag)
		If iStrFlag <> ggoSpread.InsertFlag And CInt(Replace(iStrLevel, ".", "")) = 1 Then
'			ggoSpread.EditUndo
			Call frm1.vspdData.SetText(0, iIntCnt, ggoSpread.DeleteFlag)
		End If
	Next
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
   Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLEMULTI)							'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLEMULTI, False)	                   '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc :
'========================================================================================
Function FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display(��ܺ�)
'========================================================================================
Function DbQuery()

    Dim LngLastRow
    Dim LngMaxRow
    Dim LngRow
    Dim strTemp
    Dim StrNextKey
    Dim iStrBasePlantCd, iStrDestPlantCd, iStrItemCd, iStrReqTransNo
    Dim strQueryType

    DbQuery = False

    LayerShowHide(1)

    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal

	iStrBasePlantCd = UCase(Trim(frm1.txtBasePlantCd.value))
	iStrDestPlantCd = UCase(Trim(frm1.txtDestPlantCd.value))
	iStrItemCd      = UCase(Trim(frm1.txtItemCd.value))
	iStrReqTransNo  = UCase(Trim(frm1.txtReqTransNo.value))

'	strQueryType = UCase(Trim(frm1.txtQueryType.value))						'��: A : ����BOM QUERY, B : ����BOM QUERY, * : ALL

	strQueryType = "B"						'��: A : ����BOM QUERY, B : ����BOM QUERY, * : ALL

    With frm1
		strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001
		strVal = strVal & "&QueryType="			& strQueryType
		strVal = strVal & "&txtBasePlantCd="	& iStrBasePlantCd				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtDestPlantCd="	& iStrDestPlantCd				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd="			& iStrItemCd					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtReqTransNo="		& iStrReqTransNo 				'��: ��ȸ ���� ����Ÿ 
'		strVal = strVal & "&txtSerchType="		& "1"
'		strVal = strVal & "&txtBaseBomNo="		& "E"
'		strVal = strVal & "&txtDestBomNo="		& "1"
'		strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
'		strVal = strVal & "&txtMaxRows1="		& .vspdData1.MaxRows
'       strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex				'��: Next key tag
'       strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1			'��: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

    End With

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
'Function DbQueryOk(LngMaxRow)										'��: ��ȸ ������ ������� 
Function DbQueryOk()										'��: ��ȸ ������ ������� 

	Dim lRow
	Dim i
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE								'��: Indicates that current mode is Update mode

	Call ggoOper.LockField(Document, "N")							'��: This function lock the suitable field

	With frm1
		.hBasePlantCd.value = UCase(Trim(.txtBasePlantCd.value))
		.hDestPlantCd.value = UCase(Trim(.txtDestPlantCd.value))
		.hItemCd.value      = UCase(Trim(.txtItemCd.value))
		.hReqTransNo.value  = UCase(Trim(.txtReqTransNo.value))
	End With

    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If

    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode

	frm1.btnAutoSel1.disabled = False

End Function

Function DbQueryNotOk()
    lgIntFlgMode = parent.OPMD_CMODE								'��: Indicates that current mode is Update mode

	Call SetToolbar("11001111001001")

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()

    Dim lRow
	Dim strVal
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

	With frm1.vspdData

		For lRow = 1 To .MaxRows
			.Row = lRow
'		    .Col = 0
			.Col = C_Select

			If .Value = 1 Then
					strVal = ""

					'//Ref. ConstBas\P0\BCP4B3_PProdGoodsRcpt.bas
					.Col = C_BomNo
					strVal = strVal & Trim(.Text) & iColSep				'BOM_NO
					.Col = C_ReqTransNo
					strVal = strVal & Trim(.Text) & iColSep				'�̰��Ƿڹ�ȣ 
					.Col = C_DestPlantCd
					strVal = strVal & Trim(.Text) & iColSep				'������ 
					.Col = C_BasePlantCd
					strVal = strVal & Trim(.Text) & iColSep				'������� 
					.Col = C_ItemCd
					strVal = strVal & Trim(.Text) & iColSep				'ǰ�� 
					.Col = C_ReqTransDt
					strVal = strVal & Trim(.Text) & iColSep				'�̰���û�� 
					.Col = C_TransDt
					strVal = strVal & Trim(.Text) & iColSep				'�̰��� 
					.Col = C_BomDesc
					strVal = strVal & Trim(.Text) & iColSep				'BOM���� 
					.Col = C_ValidFromDt
					strVal = strVal & Trim(.Text) & iColSep				'������ 
					.Col = C_ValidToDt
					strVal = strVal & Trim(.Text) & iColSep				'������ 
					.Col = C_DrawingPath
					strVal = strVal & Trim(.Text) & iColSep				'������ 
					.Col = C_TransStatus
					strVal = strVal & Trim(.Text) & iColSep				'�̰����� 
					.Col = C_ReturnDesc
					If gbtnAuto = "btnPrc" Then
						strVal = strVal & "" & iColSep				'�ݷ����� 
					Else
						strVal = strVal & Trim(.Text) & iColSep				'�ݷ����� 
					End IF	

					'------------------------------------------------
					'//		Insert another txtSpread value
					'------------------------------------------------

					strVal = strVal & lRow & iRowSep						  'Count (to trace error row)

					If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 

					   Set objTEXTAREA = document.createElement("TEXTAREA")   '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
					   objTEXTAREA.name = "txtCUSpread"
					   objTEXTAREA.value = Join(iTmpCUBuffer,"")
					   divTextArea.appendChild(objTEXTAREA)

					   iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT      '�ӽ� ���� ���� �ʱ�ȭ 
					   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
					   iTmpCUBufferCount = -1
					   strCUTotalvalLen  = 0
					End If

					iTmpCUBufferCount = iTmpCUBufferCount + 1

					If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                            '������ ���� ����ġ�� ������ 
					   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
					   ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
					End If

					iTmpCUBuffer(iTmpCUBufferCount) =  strVal
					strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			End If
		Next

		If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)
		End If

		If gFlg = "check" Then
			Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID2)							'��: �����Ͻ� ASP �� ���� 
		Else
			Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)							'��: �����Ͻ� ASP �� ���� 
		End If

	End With

    DbSave = True                                                           '��: Processing is NG

	gFlg = ""

End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 

	Call InitVariables
    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbCheckOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbCheckOk()													'��: ���� ������ ���� ���� 

	gFlg = "check"
	Call RemovedivTextArea

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		Exit Function
    End If     				                                                  '��: Save db data

End Function

'========================================================================================
' Function Name : DbErrorPrcOk
' Function Desc :
'========================================================================================
Function DbErrorPrcOk()													'��:
	Call RemovedivTextArea

	Call DisplayMsgBox("187214","X", "","X")
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()												'��: ���� ������ ���� ���� 
'	Call InitVariables
'	Call FncNew()
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


'------------------------------------------  OpenReqTransNo()  -------------------------------------------
'	Name : OpenReqTransNo()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function OpenReqTransNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strPlantCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	strPlantCd = Trim(frm1.txtDestPlantCd.value)

	' �˾� ��Ī 
	arrParam(0) = "�̰��Ƿڹ�ȣ"
	' TABLE ��Ī 
	arrParam(1) = "P_EBOM_TO_PBOM_MASTER A, B_ITEM B, B_PLANT C"
	' Code Condition
	arrParam(2) = Trim(frm1.txtReqTransNo.value)
	' Name Cindition
	arrParam(3) = ""
	' Where Condition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD AND A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S")
	' TextBox ��Ī 
	arrParam(5) = "�̰��Ƿڹ�ȣ"

    arrField(0) = "A.REQ_TRANS_NO"				' Field��(0)
    arrField(1) = "A.PLANT_CD"					' Field��(1)
    arrField(2) = "C.PLANT_NM"					' Field��(2)
    arrField(3) = "A.ITEM_CD"					' Field��(3)
    arrField(4) = "B.ITEM_NM"					' Field��(4)
    arrField(5) = "A.STATUS"					' Field��(5)

    arrHeader(0) = "�̰��Ƿڹ�ȣ"			' Header��(0)
    arrHeader(1) = "������"				' Header��(1)
    arrHeader(2) = "�������"				' Header��(2)
    arrHeader(3) = "ǰ��"					' Header��(3)
    arrHeader(4) = "ǰ���"					' Header��(4)
    arrHeader(5) = "�̰�����"				' Header��(5)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReqTransNo(arrRet)
	End If

	Call SetFocusToDocument("M")

	frm1.txtReqTransNo.focus

End Function

'------------------------------------------  SetReqTransNo()  --------------------------------------------------
'	Name : SetReqTransNo()
'	Description : SetReqTransNo
'---------------------------------------------------------------------------------------------------------
Function SetReqTransNo(Byval arrRet)
	frm1.txtReqTransNo.Value	= arrRet(0)
'	frm1.txtDestPlantCd.Value	= arrRet(1)
'	frm1.txtDestPlantNm.Value	= arrRet(2)
'	frm1.txtItemCd.Value		= arrRet(3)
'	frm1.txtItemNm.Value		= arrRet(4)
'	frm1.hStatus.Value			= arrRet(5)
End Function

'------------------------------------------  OpenBomDetail()  --------------------------------------------
'	Name : OpenBomDetail()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function OpenBomDetail()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim strReqTransNo

	If IsOpenPop = True Then Exit Function

	If frm1.txtDestPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtDestPlantCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	iCalledAspName = AskPRAspName("p1714pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p1714pa2", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	frm1.vspdData.Col = C_ReqTransNo
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	If frm1.vspdData.Row = 0 Then
		strReqTransNo = ""
	Else
		strReqTransNo = Trim(frm1.vspdData.Text)'Trim(frm1.txtReqTransNo.value)
	End If

	arrParam(0) = frm1.txtDestPlantCd.value
	arrParam(1) = strReqTransNo	'Trim(frm1.txtPlantCd.value)
'	arrParam(1) = ""
'	arrParam(2) = ""

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=740px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	Call SetFocusToDocument("M")
End Function


'------------------------------------------  OpenErrorList()  --------------------------------------------
'	Name : OpenErrorList()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function OpenErrorList2()
	Dim arrRet
	Dim arrParam(5), arrField(5), arrHeader(5)
	Dim strReqTransNo

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	frm1.vspdData.Col = C_ReqTransNo
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	If frm1.vspdData.Row = 0 Then
		strReqTransNo = ""
	Else
		strReqTransNo = Trim(frm1.vspdData.Text)'Trim(frm1.txtReqTransNo.value)
	End If

	' �˾� ��Ī 
	arrParam(0) = "Error List"
	' TABLE ��Ī 
	arrParam(1) = "P_TRANS_BOM_ERROR A, P_EBOM_TO_PBOM_MASTER B, P_EBOM_TO_PBOM_DETAIL C, B_PLANT D, B_ITEM E"
	' Code Condition
	arrParam(2) = strReqTransNo'Trim(frm1.txtReqTransNo.value)
	' Name Cindition
	arrParam(3) = ""
	' Where Condition
	arrParam(4) = ""
	arrParam(4) = "A.PRNT_PLANT_CD = B.PLANT_CD AND A.PRNT_ITEM_CD = B.ITEM_CD AND A.PRNT_BOM_NO = B.BOM_NO AND A.REQ_TRANS_NO = B.REQ_TRANS_NO"
	arrParam(4) = arrParam(4) & " AND B.PLANT_CD = C.PRNT_PLANT_CD AND B.ITEM_CD = C.PRNT_ITEM_CD AND B.BOM_NO = C.PRNT_BOM_NO AND B.REQ_TRANS_NO = C.REQ_TRANS_NO"
	arrParam(4) = arrParam(4) & " AND A.CHILD_ITEM_SEQ = C.CHILD_ITEM_SEQ AND B.PLANT_CD = D.PLANT_CD AND C.CHILD_ITEM_CD = E.ITEM_CD"
	arrParam(4) = arrParam(4) & " AND A.REQ_TRANS_NO = " & FilterVar(strReqTransNo, "''", "S")

	' TextBox ��Ī 
	arrParam(5) = "�̰��Ƿڹ�ȣ"

    arrField(0) = "A.REQ_TRANS_NO"		' Field��(0)
    arrField(1) = "B.PLANT_CD"			' Field��(1)
    arrField(2) = "D.PLANT_NM"			' Field��(2)
    arrField(3) = "C.CHILD_ITEM_CD"		' Field��(3)
    arrField(4) = "E.ITEM_NM"			' Field��(4)
	arrField(5) = "A.ERROR_DESC"		' Field��(5)

    arrHeader(0) = "�̰��Ƿڹ�ȣ"	' Header��(0)
    arrHeader(1) = "������"		' Header��(1)
    arrHeader(2) = "�������"		' Header��(2)
    arrHeader(3) = "��ǰ��"			' Header��(3)
    arrHeader(4) = "��ǰ���"		' Header��(4)
	arrHeader(5) = "Error Message"	' Header��(5)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	Call SetFocusToDocument("M")
End Function


'------------------------------------------  OpenErrorList()  --------------------------------------------
'	Name : OpenErrorList()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function OpenErrorList()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim strReqTransNo

	If IsOpenPop = True Then Exit Function

	If frm1.txtDestPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtDestPlantCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	iCalledAspName = AskPRAspName("p1714pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p1714pa2", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	frm1.vspdData.Col = C_ReqTransNo
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	If frm1.vspdData.Row = 0 Then
		strReqTransNo = ""
	Else
		strReqTransNo = Trim(frm1.vspdData.Text)'Trim(frm1.txtReqTransNo.value)
	End If

	arrParam(0) = frm1.txtDestPlantCd.value
	arrParam(1) = strReqTransNo	'Trim(frm1.txtPlantCd.value)
'	arrParam(1) = ""
'	arrParam(2) = ""

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	Call SetFocusToDocument("M")
End Function