Const BIZ_PGM_QRY_ID = "b1b11mb8.asp"			 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "b1b11mb9.asp"	
Const BIZ_PGM_JUMPITEMBYPLANT_ID = "b1b11ma1"
Const BIZ_PGM_JUMPLOTCONTROL_ID = "b1b12ma1"

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_Unit
Dim C_Account
Dim C_ItemGroupCd
Dim C_ProcType
Dim C_SsQty
Dim C_MaxMrpQty
Dim C_DamperMax
Dim C_MinMrpQty
Dim C_FixedMrpQty
Dim C_LineNo
Dim C_RoundQty
Dim C_ReqRoundFlg
Dim C_ScrapRateMfg
Dim C_ScrapRatePur
Dim C_InspecLtMfg
Dim C_InspecLtPur
Dim C_InvCheckFlg
Dim C_InvMgrNm
Dim C_InvMgr
Dim C_MRPMgrNm
Dim C_MRPMgr
Dim C_ProdMgrNm
Dim C_ProdMgr
Dim C_MPSMgrNm
Dim C_MPSMgr
Dim C_InspecMgrNm
Dim C_InspecMgr
Dim C_StdTime
Dim C_VarLT
Dim C_LotFlg
Dim C_CalType
Dim C_CalTypePopup
Dim C_ValidFlg
Dim C_AtpLt
Dim C_EtcFlg1
Dim C_EtcFlg2
Dim C_OverRcptFlg
Dim C_OverRcptRate
Dim C_DamperMin
Dim C_DamperFlg
Dim C_Location	

Dim lgNextNo
Dim lgPrevNo
Dim lgStrPrevKey1
Dim lgOldRow
Dim IsOpenPop
Dim gSelframeFlg							 '���� TAB�� ��ġ�� ��Ÿ���� Flag 
Dim gblnWinEvent							 '~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()  
	C_ItemCd		= 1
	C_ItemNm		= 2
	C_Spec			= 3
	C_Unit			= 4
	C_Account		= 5
	C_ItemGroupCd	= 6
	C_ProcType		= 7
	C_SsQty			= 8
	C_MaxMrpQty		= 9
	C_DamperMax		= 10
	C_MinMrpQty		= 11
	C_FixedMrpQty	= 12
	C_LineNo		= 13
	C_RoundQty		= 14
	C_ReqRoundFlg	= 15
	C_ScrapRateMfg	= 16
	C_ScrapRatePur	= 17
	C_InspecLtMfg	= 18
	C_InspecLtPur	= 19
	C_InvCheckFlg	= 20
	C_InvMgrNm		= 21
	C_InvMgr		= 22
	C_MRPMgrNm		= 23
	C_MRPMgr		= 24
	C_ProdMgrNm		= 25
	C_ProdMgr		= 26
	C_MPSMgrNm		= 27
	C_MPSMgr		= 28
	C_InspecMgrNm	= 29
	C_InspecMgr		= 30
	C_StdTime		= 31
	C_VarLT			= 32
	C_LotFlg		= 33
	C_CalType		= 34
	C_CalTypePopup	= 35
	C_ValidFlg		= 36
	C_AtpLt			= 37
	C_EtcFlg1		= 38
	C_EtcFlg2		= 39
	C_OverRcptFlg	= 40
	C_OverRcptRate	= 41
	C_DamperMin		= 42
	C_DamperFlg		= 43
	C_Location		= 44
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE								'��: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False								'��: Indicates that no value changed
	lgIntGrpCount = 0
	lgOldRow = 0											'��: Initializes Group View Size
	lgStrPrevKey1 = ""
	lgSortKey = 1
	
	gblnWinEvent = False
End Function

'======================================== 2.2.3 InitSpreadSheet() =======================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	
	Dim i

	Call initSpreadPosVariables()    

    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20050122",, parent.gAllowDragDropSpread    

		.ReDraw = False

		.MaxCols = C_Location + 1
		.MaxRows = 0

		Call AppendNumberPlace("6", "9", "0")
		Call AppendNumberPlace("7", "2", "2")
		Call AppendNumberPlace("8", "2", "0")
		Call AppendNumberPlace("9", "3", "0")

		Call GetSpreadColumnPos("A")
    
		ggoSpread.SSSetEdit	C_ItemCd, "ǰ��", 15
		ggoSpread.SSSetEdit	C_ItemNm, "ǰ���", 20
		ggoSpread.SSSetEdit	C_Spec, "�԰�", 20
		ggoSpread.SSSetEdit	C_Unit, "����", 6
		ggoSpread.SSSetEdit	C_Account, "ǰ�����", 10
		ggoSpread.SSSetEdit	C_ItemGroupCd, "ǰ��׷�", 10
		ggoSpread.SSSetEdit	C_ProcType, "���ޱ���", 10
		ggoSpread.SSSetFloat	C_SsQty, "�������", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
	
		ggoSpread.SSSetFloat	C_MaxMrpQty, "�ִ��������", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,, "Z"
		ggoSpread.SSSetFloat	C_DamperMax, "���� L/T", 9, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000,parent.gComNumDec,,,, "0", "99"
		ggoSpread.SSSetFloat	C_MinMrpQty, "�ּҿ�������", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_FixedMrpQty, "������������", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_LineNo, "���Ҽ�", 8, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,, "1", "99"
	
		ggoSpread.SSSetFloat	C_RoundQty, "�ø���", 8, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetCombo	C_ReqRoundFlg, "�ҿ䷮�ø�����", 13
		ggoSpread.SSSetFloat	C_ScrapRateMfg, "����ǰ��ҷ���", 15, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_ScrapRatePur, "����ǰ��ҷ���", 15, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_InspecLtMfg, "�����˻� LT", 13, "9", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_InspecLtPur, "���Ű˻� LT", 13, "9", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetCombo	C_InvCheckFlg, "�������üũ", 10
	
		ggoSpread.SSSetCombo	C_InvMgrNm, "�������", 13
		ggoSpread.SSSetCombo	C_InvMgr, "�������", 10
		ggoSpread.SSSetCombo	C_MRPMgrNm, "MRP�����", 13
		ggoSpread.SSSetCombo	C_MRPMgr, "MRP�����", 10
		ggoSpread.SSSetCombo	C_ProdMgrNm, "��������", 13
		ggoSpread.SSSetCombo	C_ProdMgr, "��������", 10
		ggoSpread.SSSetCombo	C_MPSMgrNm, "�����˻�����", 13
		ggoSpread.SSSetCombo	C_MPSMgr, "�����˻�����", 10
		ggoSpread.SSSetCombo	C_InspecMgrNm, "���Ű˻�����", 13
		ggoSpread.SSSetCombo	C_InspecMgr, "���Ű˻�����", 10
		ggoSpread.SSSetFloat	C_StdTime, "ǥ�� ST", 8, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_VarLT, "����L/T", 8, "9", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit	C_LotFlg, "LOT��������", 6
		ggoSpread.SSSetEdit	C_CalType, "Į����Ÿ��", 10,,, 2, 2
		ggoSpread.SSSetButton	C_CalTypePopup  
	
		ggoSpread.SSSetCombo	C_ValidFlg, "��ȿ����", 10
	
		ggoSpread.SSSetFloat	C_AtpLt, "ATP L/T", 8, "9", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit	C_EtcFlg1, "C_EtcFlg1", 6
		ggoSpread.SSSetEdit	C_EtcFlg2, "C_EtcFlg2", 6
	
		ggoSpread.SSSetCombo	C_OverRcptFlg, "���԰���뿩��", 13
		ggoSpread.SSSetFloat	C_OverRcptRate, "���԰������", 13, "9", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_DamperMin, "Damper �ּ���", 13, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetCombo	C_DamperFlg, "Damper����", 10,,False
		ggoSpread.SSSetEdit	C_Location, "Location", 30,,, 20
	
		Call ggoSpread.MakePairsColumn(C_CalType, C_CalTypePopup)

		Call ggoSpread.SSSetColHidden(C_InvMgr, C_InvMgr, True)
		Call ggoSpread.SSSetColHidden(C_MRPMgr, C_MRPMgr, True)
		Call ggoSpread.SSSetColHidden(C_ProdMgr, C_ProdMgr, True)
		Call ggoSpread.SSSetColHidden(C_MPSMgr, C_MPSMgr, True)
		Call ggoSpread.SSSetColHidden(C_InspecMgr, C_InspecMgr, True)
		Call ggoSpread.SSSetColHidden(C_LotFlg, C_LotFlg, True)
		Call ggoSpread.SSSetColHidden(C_EtcFlg1, C_EtcFlg2, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SSSetSplit2(1)										'frozen ����߰� 
    
		.ReDraw = True

		Call SetSpreadLock 
		
    End With

End Sub

'================================== 2.2.4 SetSpreadLock() ==============================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
     With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.SpreadLock		C_ItemCd, -1, C_ItemCd
		ggoSpread.SpreadLock		C_ItemNm, -1, C_ItemNm
		ggoSpread.SpreadLock		C_Spec, -1, C_Spec
		ggoSpread.SpreadLock		C_Unit, -1, C_Unit
		ggoSpread.SpreadLock		C_Account, -1, C_Account
		ggoSpread.SpreadLock		C_ItemGroupCd, -1, C_ItemGroupCd
		ggoSpread.SpreadLock		C_ProcType, -1, C_ProcType
	
		ggoSpread.SSSetRequired		C_InvCheckFlg, -1
		ggoSpread.SSSetRequired		C_ValidFlg , -1
		ggoSpread.SSSetRequired		C_CalType, -1
		ggoSpread.SSSetProtected	.vspdData.MaxCols , -1
	
		.vspdData.ReDraw = True
	End With
End Sub

Sub SetCookieVal()

	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm") 
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""

End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1
		.vspdData.ReDraw = False

		ggoSpread.SSSetProtected		C_ItemCd, pvStartRow, pvEndRow	
		ggoSpread.SSSetProtected		C_ItemNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected		C_Spec, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected		C_Unit, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected		C_Account, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected		C_ItemGroupCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected		C_ProcType, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_InvCheckFlg, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_ValidFlg, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_CalType, pvStartRow, pvEndRow
	
		ggoSpread.SpreadUnLock			C_CalTypePopup, lRow, C_CalTypePopup	

		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_Unit			= iCurColumnPos(4)
			C_Account		= iCurColumnPos(5)
			C_ItemGroupCd	= iCurColumnPos(6)
			C_ProcType		= iCurColumnPos(7)
			C_SsQty			= iCurColumnPos(8)
			C_MaxMrpQty		= iCurColumnPos(9)
			C_DamperMax		= iCurColumnPos(10)
			C_MinMrpQty		= iCurColumnPos(11)
			C_FixedMrpQty	= iCurColumnPos(12)
			C_LineNo		= iCurColumnPos(13)
			C_RoundQty		= iCurColumnPos(14)
			C_ReqRoundFlg	= iCurColumnPos(15)
			C_ScrapRateMfg	= iCurColumnPos(16)
			C_ScrapRatePur	= iCurColumnPos(17)
			C_InspecLtMfg	= iCurColumnPos(18)
			C_InspecLtPur	= iCurColumnPos(19)
			C_InvCheckFlg	= iCurColumnPos(20)
			C_InvMgrNm		= iCurColumnPos(21)
			C_InvMgr		= iCurColumnPos(22)
			C_MRPMgrNm		= iCurColumnPos(23)
			C_MRPMgr		= iCurColumnPos(24)
			C_ProdMgrNm		= iCurColumnPos(25)
			C_ProdMgr		= iCurColumnPos(26)
			C_MPSMgrNm		= iCurColumnPos(27)
			C_MPSMgr		= iCurColumnPos(28)
			C_InspecMgrNm	= iCurColumnPos(29)
			C_InspecMgr		= iCurColumnPos(30)
			C_StdTime		= iCurColumnPos(31)
			C_VarLT			= iCurColumnPos(32)
			C_LotFlg		= iCurColumnPos(33)
			C_CalType		= iCurColumnPos(34)
			C_CalTypePopup	= iCurColumnPos(35)
			C_ValidFlg		= iCurColumnPos(36)
			C_AtpLt			= iCurColumnPos(37)
			C_EtcFlg1		= iCurColumnPos(38)
			C_EtcFlg2		= iCurColumnPos(39)
			C_OverRcptFlg	= iCurColumnPos(40)
			C_OverRcptRate	= iCurColumnPos(41)
			C_DamperMin		= iCurColumnPos(42)
			C_DamperFlg		= iCurColumnPos(43)
			C_Location		= iCurColumnPos(44)
    End Select
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
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()

    Call InitSpreadSheet()
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()

	Call InitData(1)
	
	frm1.vspdData.Focus
	Set gActiveElement = document.activeElement 
End Sub

'============================= 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'========================================================================================= 
 Sub InitSpreadComboBox()
	Dim i, iStrArr, iStrNmArr
    Dim strCbo  
    Dim strCboCd
    Dim strCboNm 
    
	'-----------------------------------------------------------------------------------------------------
	' List �ҿ䷮�ø�����/�������üũ/��ȿ����/���԰���뿩��/
	'-----------------------------------------------------------------------------------------------------
	strCbo = ""    
	strCbo = strCbo & "Y" & vbTab & "N" 
    
	ggoSpread.SetCombo strCbo,C_ReqRoundFlg
	ggoSpread.SetCombo strCbo,C_InvCheckFlg
	ggoSpread.SetCombo strCbo,C_ValidFlg
	ggoSpread.SetCombo strCbo,C_DamperFlg  
	
	strCbo = ""    
	strCbo = strCbo & "Y" & vbTab & "N" 
	ggoSpread.SetCombo strCbo,C_OverRcptFlg

   '****************************
    'List Minor code(MRP �����)
    '****************************
    strCboCd = "" & vbTab & ""
    strCboNm = "" & vbTab 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1011' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iStrArr = Split(lgF0, Chr(11))
    iStrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iStrArr) - 1
        strCboCd = strCboCd & iStrArr(i) & vbTab
        strCboNm = strCboNm & iStrNmArr(i) & vbTab
	Next

    ggoSpread.SetCombo strCboCd, C_MRPMgr             'MRP ����� setting
    ggoSpread.SetCombo strCboNm,C_MRPMgrNm 'parent.ggoSpread.SSGetColsIndex()             Name Setting
	
    '****************************
    'List Minor code(�������)
    '****************************
    strCboCd = "" & vbTab & ""
    strCboNm = "" & vbTab & "" 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'I0004' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iStrArr = Split(lgF0, Chr(11))
    iStrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iStrArr) - 1
        strCboCd = strCboCd & iStrArr(i) & vbTab
        strCboNm = strCboNm & iStrNmArr(i) & vbTab
	Next

    ggoSpread.SetCombo strCboCd, C_InvMgr             '������� setting
    ggoSpread.SetCombo strCboNm,C_InvMgrNm 'parent.ggoSpread.SSGetColsIndex()             Name Setting

    '****************************
    'List Minor code(��������)
    '****************************
    strCboCd = "" & vbTab & ""
    strCboNm = "" & vbTab & "" 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1015' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iStrArr = Split(lgF0, Chr(11))
    iStrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iStrArr) - 1
        strCboCd = strCboCd & iStrArr(i) & vbTab
        strCboNm = strCboNm & iStrNmArr(i) & vbTab
	Next

    ggoSpread.SetCombo strCboCd, C_ProdMgr             '�������� setting
    ggoSpread.SetCombo strCboNm,C_ProdMgrNm 'parent.ggoSpread.SSGetColsIndex()             Name Setting

    '****************************
    'List Minor code(�����˻�����)
    '****************************
    strCboCd = "" & vbTab & ""
    strCboNm = "" & vbTab & "" 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1012' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iStrArr = Split(lgF0, Chr(11))
    iStrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iStrArr) - 1
        strCboCd = strCboCd & iStrArr(i) & vbTab
        strCboNm = strCboNm & iStrNmArr(i) & vbTab
	Next

    ggoSpread.SetCombo strCboCd, C_MPSMgr             '�����˻����� setting
    ggoSpread.SetCombo strCboNm,C_MPSMgrNm 'parent.ggoSpread.SSGetColsIndex()             Name Setting
     
    '****************************
    'List Minor code(���Ű˻�����)
    '****************************
    strCboCd = "" & vbTab & ""
    strCboNm = "" & vbTab & "" 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'Q0002' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iStrArr = Split(lgF0, Chr(11))
    iStrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iStrArr) - 1
        strCboCd = strCboCd & iStrArr(i) & vbTab
        strCboNm = strCboNm & iStrNmArr(i) & vbTab
	Next

    ggoSpread.SetCombo strCboCd, C_InspecMgr             '���Ű˻����� setting
	ggoSpread.SetCombo strCboNm, C_InspecMgrNm 'parent.ggoSpread.SSGetColsIndex()             Name Setting
	
End Sub

'==========================================  2.2.6 InitData()  ===========================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex1,intIndex2,intIndex3,intIndex4,intIndex5, intindex6
	
	With frm1.vspdData
		'.ReDraw = False
		
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_MRPMgr
			intIndex1 = .value
			.Col = C_MRPMgrNm
			.value = intindex1
			
			.col = C_ProdMgr
			intIndex2 = .value
			.Col = C_ProdMgrNm
			.value = intindex2

			.col = C_MPSMgr
			intIndex3 = .value
			.Col = C_MPSMgrNm
			.value = intindex3

			.col = C_InspecMgr
			intIndex4 = .value
			.Col = C_InspecMgrNm
			.value = intindex4

			.col = C_InvMgr
			intIndex5 = .value
			.Col = C_InvMgrNm
			.value = intindex5
			
		Next	
		'.ReDraw = True
	End With
End Sub

'------------------------------------------  OpenConPlant()  --------------------------------------------
'	Name : OpenConPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

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
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenConItemCd()  ---------------------------------------------
'	Name : OpenConItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------- 
Function OpenConItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = Trim(frm1.txtItemCd.value) 						
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    
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
		Call SetItemInfo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus	

End Function

'------------------------------------------  OpenItemGroup()  ---------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------- 
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = 'N' AND VALID_TO_DT >= " & "'" & BaseDate & "'" 			
	arrParam(5) = "ǰ��׷�"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	

    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function

'------------------------------------------  OpenCalType()  -----------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'---------------------------------------------------------------------------------------------------------- 
Function OpenCalType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	frm1.vspdData.Col = C_Caltype
		 
	arrParam(0) = "Į���� Ÿ�� �˾�"			' �˾� ��Ī 
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.vspdData.Text)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "Į���� Ÿ��"					' TextBox ��Ī 
	
    arrField(0) = "CAL_TYPE"						' Field��(0)
    arrField(1) = "CAL_TYPE_NM"						' Field��(1)
    
    arrHeader(0) = "Į���� Ÿ��"				' Header��(0)
    arrHeader(1) = "Į���� Ÿ�Ը�"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCalType(arrRet)
	End If	
    
End Function

'------------------------------------------  SetConPlant()  ---------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
	End With
End Function

'------------------------------------------  SetItemGroup()  ----------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value		= arrRet(0)	
	frm1.txtItemGroupNm.Value		= arrRet(1)	
	lgBlnFlgChgValue		= True
End Function

'------------------------------------------  SetCalType()  ------------------------------------------------
'	Name : SetCalType()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------- 
Function SetCalType(byval arrRet)
	With frm1.vspdData
		.Col = C_Caltype
		.Text = arrRet(0)
		Call vspdData_Change(.Col, .Row)		' ������ �Ͼ�ٰ� �˷��� 
	End With
End Function

'=============================================  2.5.1 JumpItemByPlant()  ================================
'=	Event Name : JumpItemByPlant
'=	Event Desc :
'========================================================================================================
Function JumpItemByPlant()
	Dim intRet
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
		If ggoSpread.SSCheckChange = False Then    
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then    
		IntRet = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")												'��: "Will you destory previous data"
		If intRet = vbNo Then				 
			Exit Function
		End If
	End If
	
	WriteCookie "txtPlantCd", Trim(UCase(frm1.txtPlantCd.value))
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	With frm1.vspdData 
	
	    ggoSpread.Source = frm1.vspdData
	    .Col = C_ItemCd
	    .Row = .Activerow
	    WriteCookie "txtItemCd", Trim(.value)
		
		.Col = C_ItemNm
		WriteCookie "txtItemNm", .value
	End With
	
	PgmJump(BIZ_PGM_JUMPITEMBYPLANT_ID)

End Function

'=============================================  2.5.3 LotControl()  =====================================
'=	Event Name : LotControl	Jump																		=	=
'=	Event Desc :																						=
'========================================================================================================
Function LotControl()
	Dim IntRetCD
    
	 '------ Check previous data area ------ 
	If lgIntFlgMode = parent.OPMD_CMODE Then
		ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
		If ggoSpread.SSCheckChange = False Then    
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then    
		IntRetCD = DisplayMsgBox("900017",parent.VB_YES_NO,"X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
			.Col = C_LotFlg
			.Row = .Activerow
		
		If .value = "N" Then
			Call DisplayMsgBox("122814", "X", "X", "X")
			Exit Function
		End If 
	End With
	
	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	With frm1.vspdData 
	
	    ggoSpread.Source = frm1.vspdData
	    .Col = C_ItemCd
	    .Row = .Activerow
	    WriteCookie "txtItemCd", Trim(.value)
		
		.Col = C_ItemNm
		WriteCookie "txtItemNm", Trim(.value)
	End With
	
	PgmJump(BIZ_PGM_JUMPLOTCONTROL_ID)

End Function

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
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
			Case  C_MRPMgrNm
				.Col = Col
				intIndex = .Value
				.Col = C_MRPMgr
				.Value = intIndex
			Case  C_ProdMgrNm
				.Col = Col
				intIndex = .Value
				.Col = C_ProdMgr
				.Value = intIndex
			Case  C_MPSMgrNm
				.Col = Col
				intIndex = .Value
				.Col = C_MPSMgr
				.Value = intIndex
			Case  C_InspecMgrNm
				.Col = Col
				intIndex = .Value
				.Col = C_InspecMgr
				.Value = intIndex
			Case  C_InvMgrNm
				.Col = Col
				intIndex = .Value
				.Col = C_InvMgr
				.Value = intIndex
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
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop)	Then
    
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	 '----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
   
		If Row > 0 And Col = C_CalTypePopup Then
			.Col = Col
			.Row = Row

			Call OpenCalType
			Call SetActiveCell(frm1.vspdData,C_Caltype,Row,"M","X","X")
			Set gActiveElement = document.activeElement
		End If
    
    End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SPC"					'SpreadSheet ������ vspdData�ϰ�� 

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If
	
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
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

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 1.����ð�(runtime)�� �˾��޴��� ���ؼ� �������� �ٲ���.
'				 2.Mouse�� Ư��Cell�� ����("SPC")�ϰ� ������ ��ư("SPCR")�� ������ �˾��� ���δ�.
'				   �˾����� Ư�� �޴� item�� ����("SPCRP") ���� Į���� freeze�Ѵ�.
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=========================================  5.1.1 FncQuery()  ===========================================
'=	Event Name : FncQuery																				=
'=	Event Desc : This function is related to Query Button of Main ToolBar								=
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
	
    FncQuery = False															'��: Processing is NG

    Err.Clear																    '��: Protect system from crashing
	
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then                   '��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
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
	 
	If frm1.txtItemGroupCd.value = "" Then
		frm1.txtItemGroupNm.value = ""
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables
    
  	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     									'��: Query db data

    FncQuery = True																'��: Processing is OK
    Set gActiveElement = document.ActiveElement   
End Function

'===========================================  5.1.4 FncSave()  ==========================================
'=	Event Name : FncSave																				=
'=	Event Desc : This function is related to Save Button of Main ToolBar								=
'========================================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    On Error Resume Next														'��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")        '��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then              '��: Check required field(Multi area)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		LayerShowHide(0)
		Exit Function           
    End If     									'��: Save db data
    
    FncSave = True																'��: Processing is OK
    Set gActiveElement = document.ActiveElement   
           
End Function

'===========================================  5.1.6 FncCancel()  ========================================
'=	Event Name : FncCancel																				=
'=	Event Desc : This function is related to Cancel Button of Main ToolBar								=
'========================================================================================================

Function FncCancel() 
 
	If frm1.vspdData.Maxrows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData	
	ggoSpread.EditUndo                                                    '��: Protect system from crashing
	
	frm1.vspdData.Redraw = False
	Call InitData(1)
	frm1.vspdData.Redraw = True
    
    Set gActiveElement = document.ActiveElement   
End Function

'============================================  5.1.9 FncPrint()  ========================================
'=	Event Name : FncPrint																				=
'=	Event Desc : This function is related to Print Button of Main ToolBar								=
'========================================================================================================

Function FncPrint()
	Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function

'===========================================  5.1.12 FncExcel()  ========================================
'=	Event Name : FncExcel																				=
'=	Event Desc : This function is related to Excel Button of Main ToolBar								=
'========================================================================================================

Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
    Set gActiveElement = document.ActiveElement   
End Function


'===========================================  5.1.13 FncFind()  =========================================
'=	Event Name : FncFind																				=
'=	Event Desc :																						=
'========================================================================================================

Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, True)
    Set gActiveElement = document.ActiveElement   
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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X", "X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
    Set gActiveElement = document.ActiveElement   
    
End Function

'=============================================  5.2.1 DbQuery()  ========================================
'=	Event Name : DbQuery																				=
'=	Event Desc : This function is data query and display												=
'========================================================================================================
Function DbQuery()
	
	Dim strAvailableItem
	Err.Clear															'��: Protect system from crashing

	DbQuery = False														'��: Processing is NG

	LayerShowHide(1)								'��: �۾������� ǥ��	
	
	Dim strVal
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ				
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboAccount=" & Trim(frm1.hItemAccunt.value)
		strVal = strVal & "&cboProcType=" & Trim(frm1.hProcType.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value	)
		strVal = strVal & "&rdoAvailableItem=" & frm1.hAvailableItem.value	
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboAccount=" & Trim(frm1.cboAccount.value)
		strVal = strVal & "&cboProcType=" & Trim(frm1.cboProcType.value)
		
		If frm1.rdoAvailableItem1.checked = True then
			strAvailableItem = "A"
		ElseIf frm1.rdoAvailableItem2.checked = True then
			strAvailableItem = "Y"
		Else			
			strAvailableItem = "N"
		End IF
		strVal = strVal & "&rdoAvailableItem=" & strAvailableItem
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		
	End If	

	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

	DbQuery = True														'��: Processing is NG
End Function


'=============================================  5.2.4 DbQueryOk()  ======================================
'=	Event Name : DbQueryOk																				=
'=	Event Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű�	=
'========================================================================================================

Function DbQueryOk(LngMaxRow)	
	Dim i
	Dim lRow   												 '��: ��ȸ ������ ������� 
	 '------ Reset variables area ------ 
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
	
	lgIntFlgMode = parent.OPMD_UMODE											 '��: Indicates that current mode is Update mode 
	
    Call InitData(LngMaxRow)
    
	Call ggoOper.LockField(Document, "Q")								 '��: This function lock the suitable field 
	
		
	Call SetToolbar("11001001000111")	
	
	lgBlnFlgChgValue = False
	
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
 Function DbSave()
    Dim lRow        
   	Dim strVal
	Dim iStrMaxMrpQty, iStrDamperMax, iStrMinMrpQty, iStrFixedMrpQty, iStrLineNo, iStrRoundQty
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size
	
    DbSave = False                                                          '��: Processing is NG
    
    LayerShowHide(1)
		
    On Error Resume Next                                                   '��: Protect system from crashing
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
    
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
	
    '-----------------------
    'Data manipulate area
    '-----------------------
	For lRow = 1 To .vspdData.MaxRows
    
	    .vspdData.Row = lRow
	    .vspdData.Col = 0

		If .vspdData.Text = ggoSpread.UpdateFlag Then
			
			strVal = ""
			
			.vspdData.Col = C_ItemCd	
			strVal = strVal & Trim(.vspdData.Text) & iColSep
					        
			.vspdData.Col = C_SsQty	
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep       
					           
			.vspdData.Col = C_MaxMrpQty
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
			iStrMaxMrpQty = Trim(.vspdData.Text)

			.vspdData.Col = C_DamperMax
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
			iStrDamperMax = Trim(.vspdData.Text)

			'----------------------------
			' �ִ������ ���� L/T
			'----------------------------
			If UNICDbl(iStrMaxMrpQty) = 0  And UNICDbl(iStrDamperMax) <> 0 Then
					Call DisplayMsgBox("122722", vbOKOnly, "", "")
					Call SheetFocus(lRow, C_DamperMax)
					Exit Function
			ElseIf UNICDbl(iStrMaxMrpQty) <> 0  And UNICDbl(iStrDamperMax) = 0 Then
					Call DisplayMsgBox("122723", vbOKOnly , "", "")
					Call SheetFocus(lRow, C_DamperMax)
					Exit Function
			End If

			.vspdData.Col = C_MinMrpQty
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
			iStrMinMrpQty = Trim(.vspdData.Text)
					    
			.vspdData.Col = C_FixedMrpQty	
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
			iStrFixedMrpQty = Trim(.vspdData.Text)
					         
			.vspdData.Col = C_LineNo	
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep       
			iStrLineNo = Trim(.vspdData.Text)
					           
			.vspdData.Col = C_RoundQty
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
			iStrRoundQty = Trim(.vspdData.Text)

			'----------------------------
			' MRP Order Qty Check
			'----------------------------
			If Not (UNICDbl(iStrFixedMrpQty) = 0  Or parent.UNICDbl(iStrMinMrpQty) = 0) Then
				If UNICDbl(iStrFixedMrpQty) < UNICDbl(iStrMinMrpQty) Then
					Call DisplayMsgBox("972002", vbOKOnly, "������������", "�ּҿ�������")
					Call SheetFocus(lRow, C_MinMrpQty)
					Exit Function
				End If
			End If
		
			If Not (UNICDbl(iStrMaxMrpQty) = 0  Or parent.UNICDbl(iStrFixedMrpQty) = 0) Then
				If UNICDbl(iStrMaxMrpQty) < UNICDbl(iStrFixedMrpQty) Then 
					Call DisplayMsgBox("972002", vbOKOnly, "�ִ��������", "������������")
					Call SheetFocus(lRow, C_MaxMrpQty)
					Exit Function
				End If
			End If
		
			If Not (UNICDbl(iStrMaxMrpQty) = 0  Or parent.UNICDbl(iStrMinMrpQty) = 0) Then
				If UNICDbl(iStrMaxMrpQty) < UNICDbl(iStrMinMrpQty) Then 		
					Call DisplayMsgBox("972002", vbOKOnly, "�ִ��������", "�ּҿ�������")
					Call SheetFocus(lRow, C_MaxMrpQty)
					Exit Function
				End If
			End If
		
			If UNICDbl(iStrFixedMrpQty) <> 0 And UNICDbl(iStrLineNo) > 1 Then
				Call DisplayMsgBox("122721", vbOKOnly, "", "")
				Call SheetFocus(lRow, C_FixedMrpQty)
				Exit Function
			End If
		
			'----------------------------------
			' �ִ����, �ּҼ����� �ø��� �� 
			'----------------------------------
			If UNICDbl(iStrRoundQty) <> 0 Then
				If UNICDbl(iStrMinMrpQty) <> 0 Then
					If UNICDbl(iStrMinMrpQty) < UNICDbl(iStrRoundQty) Then 		
						Call DisplayMsgBox("972004", vbOKOnly, "�ø���", "�ּҿ�������")
						Call SheetFocus(lRow, C_RoundQty)
						Exit Function
					End If
				ElseIf UNICDbl(iStrMaxMrpQty) <> 0 Then
					If UNICDbl(iStrMaxMrpQty) < UNICDbl(iStrRoundQty) Then 		
						Call DisplayMsgBox("972004", vbOKOnly, "�ø���", "�ִ��������")
						Call SheetFocus(lRow, C_RoundQty)
						Exit Function
					End If
				End If
			End If


			.vspdData.Col = C_ReqRoundFlg
			If Trim(.vspdData.Text) <> "" Then
				strVal = strVal & Trim(.vspdData.Text) & iColSep
			Else
				strVal = strVal & "N" & iColSep
			End If
					    
			.vspdData.Col = C_ScrapRateMfg	
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
					         
			.vspdData.Col = C_ScrapRatePur	
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep       
					           
			.vspdData.Col = C_InspecLtMfg
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
					            
			.vspdData.Col = C_InspecLtPur
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep

			.vspdData.Col = C_InvCheckFlg
			If Trim(.vspdData.Text) <> "" Then
				strVal = strVal & Trim(.vspdData.Text) & iColSep
			Else
				strVal = strVal & "N" & iColSep
			End If
					    
			.vspdData.Col = C_InvMgr	
			strVal = strVal & Trim(.vspdData.Text) & iColSep
					         
			.vspdData.Col = C_MRPMgr	
			strVal = strVal & Trim(.vspdData.Text) & iColSep       
					           
			.vspdData.Col = C_ProdMgr
			strVal = strVal & Trim(.vspdData.Text) & iColSep
					            
			.vspdData.Col = C_MPSMgr
			strVal = strVal & Trim(.vspdData.Text) & iColSep
					            
			.vspdData.Col = C_InspecMgr
			strVal = strVal & Trim(.vspdData.Text) & iColSep
					    
			.vspdData.Col = C_StdTime
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
					            
			.vspdData.Col = C_VarLT
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
					            
			.vspdData.Col = C_CalType
			strVal = strVal & Trim(.vspdData.Text) & iColSep

			.vspdData.Col = C_ValidFlg	
			If Trim(.vspdData.Text) <> "" Then
				strVal = strVal & Trim(.vspdData.Text) & iColSep
			Else
				strVal = strVal & "N" & iColSep
			End If     
					     
			.vspdData.Col = C_AtpLt
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep

			.vspdData.Col = C_OverRcptFlg
			If Trim(.vspdData.Text) <> "" Then
				strVal = strVal & Trim(.vspdData.Text) & iColSep
			Else
				strVal = strVal & "N" & iColSep
			End If
					    
			.vspdData.Col = C_OverRcptRate
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep
					     
			.vspdData.Col = C_DamperMin
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep

			.vspdData.Col = C_DamperFlg
			If Trim(.vspdData.Text) <> "" Then
				strVal = strVal & Trim(.vspdData.Text) & iColSep
			Else
				strVal = strVal & "N" & iColSep
			End If
					    
			.vspdData.Col = C_Location
			strVal = strVal & Trim(.vspdData.Text) & iColSep

			strVal = strVal & lRow & iRowSep								'��: Row Number
			
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
	Next
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If 
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True																	'��: Processing is NG

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
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function


Function SheetFocus(lRow, lCol)
	frm1.vspdData.Focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function