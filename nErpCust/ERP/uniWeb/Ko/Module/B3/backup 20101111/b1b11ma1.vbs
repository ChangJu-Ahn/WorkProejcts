Const BIZ_PGM_QRY_ID = "b1b11mb1.asp"			
Const BIZ_PGM_SAVE_ID = "b1b11mb2.asp"			
Const BIZ_PGM_DEL_ID = "b1b11mb3.asp"			
Const BIZ_PGM_LOOKUP_ID = "b1b11mb4.asp"		
Const BIZ_PGM_JUMPITEMBYPLANTDETAIL_ID = "b1b11ma4"
Const BIZ_PGM_JUMPLOTCONT_ID = "b1b12ma1"
Const BIZ_PGM_JUMPALTITEM_ID = "b1b13ma1"	

Dim IsOpenPop
Dim gblnWinEvent							'~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
Dim lgRdoMrpOldVal
Dim lgRdoRndOldVal
Dim lgRdoColOldVal
Dim lgRdoLotOldVal
Dim lgRdoMpsOldVal
Dim lgRdoTrkOldVal
Dim lgRdoRecOldVal
Dim lgRdoPrdOldVal
Dim lgRdoFinOldVal
Dim lgRdoIssOldVal

Dim blncboPrcCtrlIndIsM

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE						'��: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False								'��: Indicates that no value changed
	lgIntGrpCount = 0										'��: Initializes Group View Size
	
	gblnWinEvent = False
End Function
	
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
Sub SetDefaultVal()

	With frm1
		.cboAccount.value		= "10"
		.cboProcType.value		= "M"
		.cboMatType.value		= "10"
		.cboProdEnv.value		= "MX"
		.cboIssueType.value		= "A"
		.cboABCFlg.value		= "A"
		.cboPrcCtrlInd.value	= "S"
		.cboOrderFrom.value		= ""
		.cboLotSizing.value		= "L"
		
		.rdoMRPFlg1.checked				= True
		.rdoCollectFlg2.checked			= True				'�ܰ��� ���� �ʵ�� ��� 
		.rdoMPSItem1.checked 			= True
		.rdoTrackingItem2.checked		= True
		.rdoLotNoFlg2.checked			= True
		.rdoFinalInspType2.checked		= True
		.rdoIssueInspType2.checked		= True
		.rdoMfgInspType2.checked		= True
		.rdoPurInspType2.checked		= True 
		
		.txtValidFromDt.Text	= StartDate
		.txtValidToDt.Text		= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	End With
		
	lgRdoMrpOldVal				= 1
	lgRdoColOldVal				= 2
	lgRdoMpsOldVal				= 1
	lgRdoTrkOldVal				= 2
	lgRdoRndOldVal				= 2
	lgRdoLotOldVal				= 2
	lgRdoRecOldVal				= 2
	lgRdoPrdOldVal				= 2
	lgRdoFinOldVal				= 2
	lgRdoIssOldVal				= 2		
	
	blncboPrcCtrlIndIsM = False	
End Sub

'------------------------------------------  OpenConPlant()  -------------------------------------------------
'	Name : OpenConPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
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
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenConItemCd()  -------------------------------------------------
'	Name : OpenConItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = Trim(frm1.txtItemCd.value) 						
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    arrField(2) = 13 							' Field��(1) : "ITEM_NM"
    
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

 '------------------------------------------  OpenConItemCd1()  -------------------------------------------------
'	Name : OpenConItemCd1()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	If UCase(frm1.txtItemCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd1.value)	' Item Code
	arrParam(1) = "" 						
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	arrParam(4) = BaseDate						' Current Date
			
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    arrField(2) = 4
    arrField(3) = 5
    arrField(4) = 8
    
	iCalledAspName = AskPRAspName("B1B01PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo1(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd1.focus

End Function

'------------------------------------------  OpenMfgUnit()  ----------------------------------------------
'	Name : OpenMfgUnit()	������������ 
'	Description : Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMfgUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtMfgOrderUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtMfgOrderUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "Dimension <> " & FilterVar("TM", "''", "S") & "  "			
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetMfgUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtMfgOrderUnit.focus
	
End Function

'------------------------------------------  OpenPurUnit()  ---------------------------------------------
'	Name : OpenPurUnit() (���ſ�������)
'	Description : Unit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPurUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPurOrderUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtPurOrderUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "Dimension <> " & FilterVar("TM", "''", "S") & "  "						
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPurUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPurOrderUnit.focus
	
End Function


 '------------------------------------------  OpenPurOrg()  -------------------------------------------------
'	Name : OpenPurOrg()	�������� 
'	Description : PurOrg PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPurOrg()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPurOrg.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���������˾�"	
	arrParam(1) = "B_PUR_ORG"				
	arrParam(2) = Trim(frm1.txtPurOrg.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "��������"
	
    arrField(0) = "PUR_ORG"	
    arrField(1) = "PUR_ORG_NM"	
    
    arrHeader(0) = "��������"		
    arrHeader(1) = "����������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPurOrg(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPurOrg.focus
	
End Function

'------------------------------------------  OpenIssueUnit()  -------------------------------------------------
'	Name : OpenIssueUnit()	(������)
'	Description : IssueUnit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenIssueUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtIssueUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtIssueUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "Dimension <> " & FilterVar("TM", "''", "S") & "  "					
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetIssueUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtIssueUnit.focus
	
End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()	�԰�â�� 
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSLCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X" , "����", "X")
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtSLCd.Value)									' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")  	' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
	
    arrField(0) = "SL_CD"													' Field��(0)
    arrField(1) = "SL_NM"													' Field��(1)
    
    arrHeader(0) = "â��"												' Header��(0)
    arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtSLCd.focus
	
End Function

'------------------------------------------  OpenIssueSLCd()  -------------------------------------------------
'	Name : OpenIssueSLCd()	���â�� 
'	Description : Issue Storage Location PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenIssueSLCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtIssueSLCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X" , "����", "X")
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtIssueSLCd.Value)									' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")	' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
	
    arrField(0) = "SL_CD"													' Field��(0)
    arrField(1) = "SL_NM"													' Field��(1)
    
    arrHeader(0) = "â��"												' Header��(0)
    arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetIssueSLCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtIssueSLCd.focus
		
End Function

'------------------------------------------  OpenWorkCenter()  -------------------------------------------------
'	Name : OpenWorkCenter()	�۾��� 
'	Description : Work Center Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWorkCenter()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtWorkCenter.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "�۾����˾�"												' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"												' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWorkCenter.Value)								' Code Condition
	arrParam(3) = ""															' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")  & _
				  " And VALID_TO_DT >=  " & FilterVar(BaseDate , "''", "S") & "" & _
				  " And INSIDE_FLG = " & FilterVar("Y", "''", "S") & "  "										' Where Condition
	arrParam(5) = "�۾���"													' TextBox ��Ī 
	
    arrField(0) = "WC_CD"														' Field��(0)
    arrField(1) = "WC_NM"														' Field��(1)
    arrField(2) = "INSIDE_FLG"													' Field��(0)
    arrField(3) = "WC_MGR"														' Field��(1)
    arrField(4) = "CAL_TYPE"													' Field��(0)
    
    arrHeader(0) = "�۾���"													' Header��(0)
    arrHeader(1) = "�۾����"												' Header��(1)
    arrHeader(2) = "�系�ܱ���"												' Header��(0)
    arrHeader(3) = "�۾�������"											' Header��(1)
    arrHeader(4) = "Į����Ÿ��"												' Header��(0)
    
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWorkCenter(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWorkCenter.focus
	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus		
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
		.txtPhantomFlg.value = arrRet(2)
		.txtItemCd.focus
	End With
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo1(Byval arrRet)
	With frm1
		.txtItemCd1.value = arrRet(0)
		.txtItemNm1.value = arrRet(1)
		.cboAccount.value = arrRet(3)
		.txtIssueUnit.value = arrRet(2)  
		.txtMfgOrderUnit.value = arrRet(2)
		.txtPurOrderUnit.value = arrRet(2)
		.txtBasicUnit.value = arrRet(2)
		.txtPhantomFlg.value = UCase(Trim(arrRet(4)))
		
		Call SetFieldForPhantom
		Call SetFieldForAccout(0)
		
		If .cboAccount.value >= "30" And .cboAccount.value <= "50" Then
			.cboProcType.value = "P"
		End If
		
		Call SetFieldForProcType(0)
				 
	End With
	
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetMfgUnit()  --------------------------------------------------
'	Name : SetMfgUnit()
'	Description : MfgUnit Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetMfgUnit(byval arrRet)
	frm1.txtMfgOrderUnit.Value    = arrRet(0)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetPurUnit()  --------------------------------------------------
'	Name : SetPurUnit()
'	Description : Unit Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPurUnit(byval arrRet)
	frm1.txtPurOrderUnit.Value    = arrRet(0)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetPurOrg()  --------------------------------------------------
'	Name : SetPurOrg()
'	Description : PurOrg Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPurOrg(byval arrRet)
	frm1.txtPurOrg.Value    = arrRet(0)
	frm1.txtPurOrgNm.Value    = arrRet(1)				
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetIssueUnit()  --------------------------------------------------
'	Name : SetIssueUnit()
'	Description : IssueUnit Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetIssueUnit(byval arrRet)
	frm1.txtIssueUnit.Value    = arrRet(0)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)
	
	If Trim(frm1.txtIssueSLCd.value) = "" Then
		frm1.txtIssueSLCd.value = arrRet(0)		
	End If
	
	lgBlnFlgChgValue = True
End Function
	
'------------------------------------------  SetIssueSLCd()  --------------------------------------------------
'	Name : SetIssueSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetIssueSLCd(byval arrRet)
	frm1.txtIssueSLCd.Value    = arrRet(0)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWorkCenter(byval arrRet)
	frm1.txtWorkCenter.Value= arrRet(0)		
	frm1.txtWcNm.Value		= arrRet(1)		
	lgBlnFlgChgValue = True
End Function

Sub SetCookieVal()
	If ReadCookie("MainFormFlg") = "LOT" Or ReadCookie("MainFormFlg") = "ALTITEM" Then
		frm1.txtPlantCd.value = ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm")
		
		WriteCookie "txtPlantCd", ""
		WriteCookie "txtPlantNm", ""
		WriteCookie "txtItemCd", ""
		WriteCookie "txtItemNm", ""
		WriteCookie "MainFormFlg",""
	
	Else	
		If ReadCookie("txtItemCd") <> "" Then
			frm1.txtItemCd.Value = ReadCookie("txtItemCd")
			frm1.txtItemNm.value = ReadCookie("txtItemNm")
		End If	
		If ReadCookie("txtPlantCd") <> "" Then
			frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
			frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
		End If	

		WriteCookie "txtPlantCd", ""
		WriteCookie "txtPlantNm", ""
		WriteCookie "txtItemCd", ""
		WriteCookie "txtItemNm", ""
	End If
End Sub


'=============================================  2.5.1 AltItem()  ======================================
'=	Event Name : AltItem	Jump																			=
'=	Event Desc :																						=
'========================================================================================================
Function AltItem()
	Dim IntRetCD
    
	 '------ Check previous data area ------ 
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
	
	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 

	PgmJump(BIZ_PGM_JUMPALTITEM_ID)
End Function

'=============================================  2.5.2 PlantItemDetail()  ======================================
'=	Event Name : Item by Plant Detail	Jump																			=
'=	Event Desc :																						=
'========================================================================================================
Function PlantItemDetail()
	Dim IntRetCD
    
	 '------ Check previous data area ------ 
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
	
	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 

	PgmJump(BIZ_PGM_JUMPITEMBYPLANTDETAIL_ID)

End Function

'=============================================  2.5.3 LotControl()  ======================================
'=	Event Name : LotControl	Jump																			=
'=	Event Desc :																						=
'========================================================================================================
Function LotControl()
	Dim IntRetCD
    
	 '------ Check previous data area ------ 
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If lgBlnFlgChgValue = False Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017",parent.VB_YES_NO, "X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
	
	If frm1.rdoLotNoFlg1.checked = False Then
		Call DisplayMsgBox("122814", "X", "X", "X")
		Exit Function
	End If 

	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 

	PgmJump(BIZ_PGM_JUMPLOTCONT_ID)

End Function

'==========================================  2.5.5 ChkValidData() =======================================
'=	Event Name : ChkValidData																				=
'=	Event Desc :																						=
'========================================================================================================
Function ChkValidData()
	Dim tmpVal1
	Dim tmpVal2
	Dim tmpVal3
	
	ChkValidData = False
	
	With frm1
		
		'----------------------------
		  ' ����, ���ޱ��� üũ 
		  '----------------------------
		If .cboAccount.value < "30" Then
			If .cboProcType.value = "P" Then
				Call DisplayMsgBox("122717", "X", "X", "X")
				.cboProcType.focus
				Set gActiveElement = document.activeElement 
				Exit Function
			End If							
		Else
			If .cboProcType.value <> "P" Then
				Call DisplayMsgBox("122718", "X", "X", "X")
				.cboProcType.focus
				Set gActiveElement = document.activeElement 
				Exit Function
			End If							
		End If
		

		If .rdoTrackingItem1.checked = True Then
			IF .cboProdEnv.value <> "MO" Then
				Call DisplayMsgBox("122719", "X", "X", "X")
				.cboProdEnv.focus
				Set gActiveElement = document.activeElement 
				Exit Function
			End If	
		End If			
		
		If .rdoLotNoFlg1.checked = True Then
			IF .cboIssueType.value <> "M" Then
				Call DisplayMsgBox("122720", "X", "X", "X")
				.cboIssueType.focus
				Set gActiveElement = document.activeElement 
				Exit Function
			End If	
		End If			
		'----------------------------
		' �ܰ� Check
		'----------------------------
		If .txtPhantomFlg.value = "N" Then
			If .cboPrcCtrlInd.value = "S" And UNICDbl(.txtStdPrice.Text) = 0  Then
				Call DisplayMsgBox("970022", "X" , "ǥ�شܰ�", "0")
				.txtStdPrice.Focus
				Set gActiveElement = document.activeElement 
				Exit Function 
'			ElseIf .cboPrcCtrlInd.value = "M" and UNICDbl(.txtMoveAvgPrice.Text) = 0 Then
'				Call DisplayMsgBox("970022", "X" , "�̵���մܰ�", "0")
'				.txtMoveAvgPrice.Focus
'				Set gActiveElement = document.activeElement
'				Exit Function 
			End If	
		End IF

		If .cboOrderFrom.value = "R" And UNICDbl(.txtReorderPoint.Text) = 0  Then
			Call DisplayMsgBox("970022", "X" , "������", "0")
			.txtReorderPoint.Focus
			Set gActiveElement = document.activeElement 
			Exit Function 
		End If	

		'----------------------------
		' Valid Date Check
		'----------------------------
		If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function     
		
	End With
	
	ChkValidData = True

End Function

'==========================================  2.5.6 LookUpItem() =======================================
'=	Event Name : LookUpItem																				=
'=	Event Desc :																						=
'========================================================================================================
 Sub LookUpItem()
	
	If gLookUpEnable = False Then Exit Sub
	
	LayerShowHide(1)
		
	Dim strVal

	strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 	
	strVal = strVal & "&txtItemCd1=" & Trim(frm1.txtItemCd1.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&PrevNextFlg=" & ""	
	strVal = strVal & "&lgCurDate=" & StartDate
	
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

End Sub

'==========================================  2.5.7 LookUpItemOk() =======================================
'=	Event Name : LookUpItemOk																				=
'=	Event Desc :																						=
'========================================================================================================
 Sub LookUpItemOk()
	
	Call SetFieldForPhantom
	Call SetFieldForAccout(0)
	
	If frm1.cboAccount.value >= "30" And frm1.cboAccount.value <= "50" Then
		frm1.cboProcType.value = "P"
	End If
	
	Call SetFieldForProcType(0)
End Sub
Sub LookUpItemNotOk()
	
End Sub

'==========================================  2.5.8 SetPrcCtrlInd() ======================================
'=	Event Name : SetPrcCtrlInd																			=
'=	Event Desc : �ܰ����п� ���� �� ����																					=
'========================================================================================================
 Sub SetPrcCtrlInd()	'�ܰ�����	M/S
	If frm1.cboProcType.value = "M" Or frm1.cboProcType.value = "O" Then	'�系����ǰ 
		Call ggoOper.SetReqAttr(frm1.txtStdPrice,"N")			'����� 
'		Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice,"Q")		'ȸ�� 
		frm1.cboPrcCtrlInd.value = "S"
	End If
	If frm1.cboPrcCtrlInd.value = "M" Then									'���ְ���ǰ 
		If blncboPrcCtrlIndIsM = False Then
			Call ggoOper.SetReqAttr(frm1.txtStdPrice,"Q")		'ȸ�� 
'			Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice,"N")	'����� 
		Else
			Call ggoOper.SetReqAttr(frm1.txtStdPrice,"Q")		'ȸ�� 
'			Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice,"Q")	'�����			
		End If
	ElseIf frm1.cboPrcCtrlInd.value = "S" Then								'����ǰ 
		Call ggoOper.SetReqAttr(frm1.txtStdPrice,"N")			'����� 
'		Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice,"Q")		'ȸ�� 
	End If
End Sub

'==========================================  2.5.9 SetFieldForProcType() ======================================
'=	Event Name : SetFieldForProcType																			=
'=	Event Desc : ���ޱ��п� ���� �� ����																					=
'========================================================================================================
 Sub SetFieldForProcType(ByVal iVal)
 	If frm1.txtPhantomFlg.value <> "Y" Then
		IF frm1.cboProcType.value = "M" Then
			Call ggoOper.SetReqAttr(frm1.txtMfgOrderUnit, "N")
			Call ggoOper.SetReqAttr(frm1.txtMfgOrderLT, "N")		
		
			Call ggoOper.SetReqAttr(frm1.txtPurOrderUnit, "D") 	
			Call ggoOper.SetReqAttr(frm1.txtPurOrderLT, "D") 
			Call ggoOper.SetReqAttr(frm1.txtPurOrg, "D") 
	
			If iVal <> 1 Then
				If frm1.txtPhantomFlg.value <> "Y" And frm1.cboAccount.value < "30" Then
					Call ggoOper.SetReqAttr(frm1.rdoCollectFlg1, "N")
					Call ggoOper.SetReqAttr(frm1.rdoCollectFlg2, "N")  
					
					If frm1.rdoCollectFlg1.checked = True Then
						Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "N")
					Else 
						Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "Q")
					End If
				
				End If
			End If
			
		Else
			Call ggoOper.SetReqAttr(frm1.txtMfgOrderUnit, "D")
			Call ggoOper.SetReqAttr(frm1.txtMfgOrderLT, "D")		
		
			Call ggoOper.SetReqAttr(frm1.txtPurOrderUnit, "N") 	
			Call ggoOper.SetReqAttr(frm1.txtPurOrderLT, "N") 
			Call ggoOper.SetReqAttr(frm1.txtPurOrg, "N")
	
			If iVal <> 1 Then
				frm1.rdoCollectFlg2.checked = True
				lgRdoColOldVal = 2 
		
				frm1.txtWorkCenter.value = ""

				Call ggoOper.SetReqAttr(frm1.rdoCollectFlg1, "Q")
				Call ggoOper.SetReqAttr(frm1.rdoCollectFlg2, "Q")  
				Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "Q")
	
			End If
	
		End If
	Else
		Exit Sub
	End If
End Sub

'==========================================  2.5.10 SetFieldForPhantom() ======================================
'=	Event Name : SetFieldForPhantom																			=
'=	Event Desc : Phantom���п� ���� �� ����																					=
'========================================================================================================
 Sub SetFieldForPhantom()
	IF frm1.txtPhantomFlg.value = "Y" Then
		If lgIntFlgMode = parent.OPMD_CMODE Then	
			frm1.rdoMRPFlg2.checked = True
			lgRdoMrpOldVal = 2 
		End If
		
		frm1.rdoCollectFlg2.checked = True
		lgRdoColOldVal = 2 
		frm1.txtWorkCenter.value = ""
		frm1.txtWcNm.value = ""
		frm1.rdoMPSItem2.checked = True
		frm1.rdoMRPFlg2.checked = True

		Call ggoOper.SetReqAttr(frm1.rdoCollectFlg1, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoCollectFlg2, "Q") 							
		Call ggoOper.SetReqAttr(frm1.cboIssueType, "D")
		Call ggoOper.SetReqAttr(frm1.txtSLCd, "D")
		Call ggoOper.SetReqAttr(frm1.txtIssueSLCd, "D")		
		Call ggoOper.SetReqAttr(frm1.txtIssueUnit, "D")
		Call ggoOper.SetReqAttr(frm1.txtCycleCntPerd, "D")
		Call ggoOper.SetReqAttr(frm1.txtMfgOrderUnit, "D")		
		Call ggoOper.SetReqAttr(frm1.txtMfgOrderLT, "D")
		Call ggoOper.SetReqAttr(frm1.txtPurOrderUnit, "D")
		Call ggoOper.SetReqAttr(frm1.txtPurOrderLT, "D")
		Call ggoOper.SetReqAttr(frm1.txtPurOrg, "D")
		Call ggoOper.SetReqAttr(frm1.cboABCFlg, "D")
		Call ggoOper.SetReqAttr(frm1.cboProdEnv, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoMPSItem1, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoMPSItem2, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoMRPFlg1, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoMRPFlg2, "Q")
		Call ggoOper.SetReqAttr(frm1.cboLotSizing, "Q")
		Call ggoOper.SetReqAttr(frm1.cboPrcCtrlInd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtStdPrice, "Q") 	 	 	 	
		
	Else
		If lgIntFlgMode = parent.OPMD_CMODE Then
			frm1.rdoMRPFlg1.checked = True
			lgRdoMrpOldVal = 1 
		End If
		
		Call ggoOper.SetReqAttr(frm1.rdoCollectFlg1, "D")
		Call ggoOper.SetReqAttr(frm1.rdoCollectFlg2, "D") 							
		Call ggoOper.SetReqAttr(frm1.cboIssueType, "N")
		Call ggoOper.SetReqAttr(frm1.txtSLCd, "N")
		Call ggoOper.SetReqAttr(frm1.txtIssueSLCd, "N")		
		Call ggoOper.SetReqAttr(frm1.txtIssueUnit, "N")
		Call ggoOper.SetReqAttr(frm1.txtCycleCntPerd, "N")
		Call ggoOper.SetReqAttr(frm1.cboABCFlg, "N") 	
		Call ggoOper.SetReqAttr(frm1.cboProdEnv, "N")
		Call ggoOper.SetReqAttr(frm1.rdoMPSItem1, "N")
		Call ggoOper.SetReqAttr(frm1.rdoMPSItem2, "N")
		Call ggoOper.SetReqAttr(frm1.rdoMRPFlg1, "N")
		Call ggoOper.SetReqAttr(frm1.rdoMRPFlg2, "N")
		Call ggoOper.SetReqAttr(frm1.cboLotSizing, "N") 
		
		If frm1.cboPrcCtrlInd.value  = "S" Then
			Call ggoOper.SetReqAttr(frm1.txtStdPrice, "N")
			
		Else 
			Call ggoOper.SetReqAttr(frm1.txtStdPrice, "Q")
			
		End If
		
		Call ggoOper.SetReqAttr(frm1.cboPrcCtrlInd , "N")
	End If

End Sub

'==========================================  2.5.11 SetFieldForAccout() ======================================
'=	Event Name : SetFieldForAccout																			=
'=	Event Desc : ǰ������� ���� �� ����																					=
'========================================================================================================
 Sub SetFieldForAccout(ByVal iVal)
	If frm1.txtPhantomFlg.value <> "Y" Then
		If iVal <> 1 Then
			If	frm1.cboAccount.value = "10" Then
				
				frm1.rdoMPSItem1.checked = True
				
				Call ggoOper.SetReqAttr(frm1.rdoMPSItem1, "N")
				Call ggoOper.SetReqAttr(frm1.rdoMPSItem2, "N")	
			Else
				
				frm1.rdoMPSItem2.checked = True
				
				If frm1.cboAccount.value = "33" Then
					frm1.cboLotSizing.value = "N"
				End If
				
				If frm1.cboAccount.value >= "30" And frm1.cboAccount.value <= "50" Then
					Call ggoOper.SetReqAttr(frm1.rdoMPSItem1, "Q")
					Call ggoOper.SetReqAttr(frm1.rdoMPSItem2, "Q")	
				Else
					Call ggoOper.SetReqAttr(frm1.rdoMPSItem1, "N")
					Call ggoOper.SetReqAttr(frm1.rdoMPSItem2, "N")	
				End If
			End If
		Else
			If frm1.cboAccount.value >= "30" And frm1.cboAccount.value <= "50" Then
				Call ggoOper.SetReqAttr(frm1.rdoMPSItem1, "Q")
				Call ggoOper.SetReqAttr(frm1.rdoMPSItem2, "Q")	
			Else
				Call ggoOper.SetReqAttr(frm1.rdoMPSItem1, "N")
				Call ggoOper.SetReqAttr(frm1.rdoMPSItem2, "N")	
			End If
		End If

		
		if frm1.cboAccount.value  >= "30" And frm1.cboAccount.value <= "50" Then
			frm1.rdoCollectFlg2.checked = True
			lgRdoColOldVal = 2 
			Call ggoOper.SetReqAttr(frm1.rdoCollectFlg1, "Q")
			Call ggoOper.SetReqAttr(frm1.rdoCollectFlg2, "Q")  
			Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "Q")
	
		Else
			If frm1.txtPhantomFlg.value <> "Y" And (frm1.cboProcType.value <> "O" And frm1.cboProcType.value <> "P") Then
				Call ggoOper.SetReqAttr(frm1.rdoCollectFlg1, "N")
				Call ggoOper.SetReqAttr(frm1.rdoCollectFlg2, "N")  
				
				If frm1.rdoCollectFlg1.checked = True Then
					Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "N")
				Else 
					Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "Q")
				End If

			Else 
				frm1.rdoCollectFlg2.checked = True
				lgRdoColOldVal = 2 
				Call ggoOper.SetReqAttr(frm1.rdoCollectFlg1, "Q")
				Call ggoOper.SetReqAttr(frm1.rdoCollectFlg2, "Q")  
				Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "Q")
			End If
		End If
	Else
		Exit Sub
	End If
End Sub

'==========================================  2.5.12 SetLotSizing() ======================================
'=	Event Name : SetLotSizing																			=
'=	Event Desc : LotSizing�� ���� �� ����																					=
'========================================================================================================

Sub SetLotSizing()
	If frm1.txtPhantomFlg.value <> "Y" Then
		If frm1.cboLotSizing.value = "P"  then
			Call ggoOper.SetReqAttr(frm1.txtRoundPeriod, "N")
		Else 
			Call ggoOper.setReqAttr(frm1.txtRoundPeriod, "Q")
			frm1.txtRoundPeriod.Value = 0
		End IF
	Else
		Exit Sub
	End IF
End Sub

'==========================================  2.5.13 SetOrderCreate() ======================================
'=	Event Name : SetOrderCreate																			=
'=	Event Desc : �����������ο� ���� �� ����																					=
'========================================================================================================
Sub SetOrderCreate()
	If frm1.txtPhantomFlg.value <> "Y" Then
		If frm1.rdoMRPFlg1.checked = True then
			Call ggoOper.SetReqAttr(frm1.cboLotSizing, "N")
		Else 
			Call ggoOper.setReqAttr(frm1.cboLotSizing, "Q")
			frm1.cboLotSizing.Value = "L"
		End IF
	End If
	
	If frm1.rdoMRPFlg2.checked = True then
		Call ggoOper.SetReqAttr(frm1.cboOrderFrom, "N")
		If frm1.cboOrderFrom.value = "" Then 
			frm1.cboOrderFrom.value = "R"
		End If
	Else
		Call ggoOper.SetReqAttr(frm1.cboOrderFrom, "Q")
		frm1.cboOrderFrom.value = ""
	End If
End Sub

'********************************************************************************************************
'*	Text Box OnChange																					*
'********************************************************************************************************
 Sub txtItemCd1_OnChange()
	
	If frm1.txtItemCd1.value <> "" Then
		Call LookUpItem()
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
 Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
 Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
 Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtSlCd_OnChange()
	
	If Trim(frm1.txtSLCd.value) <> "" Then
		If Trim(frm1.txtIssueSLCd.value) = "" Then
			frm1.txtIssueSLCd.value = Trim(frm1.txtSLCd.value)		
		End If
	End If
			
End Sub  

Sub txtCycleCntPerd_Change() 
	lgBlnFlgChgValue = True
End Sub

Sub txtMfgOrderLT_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMoveAvgPrice_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPurOrderLT_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtReorderPoint_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRoundPeriod_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtStdPrice_Change()
	lgBlnFlgChgValue = True
End Sub

'********************************************************************************************************
'*	Combo OnChange																						*
'********************************************************************************************************
Sub cboABCFlg_OnChange()
	lgBlnFlgChgValue = True		
End Sub

Sub cboAccount_OnChange() 
	lgBlnFlgChgValue = True
	Call SetFieldForAccout(0)
End Sub

Sub sub_cboProcType_OnChange() 
	Call SetFieldForProcType(0)

	Dim i
	For i = 0 To CInt(frm1.cboPrcCtrlInd.length) - 1
		frm1.cboPrcCtrlInd.Remove(0)
	Next
	
	For i = 0 To CInt(frm1.cboMatType.length) - 1
		frm1.cboMatType.Remove(0)
	Next	

	If frm1.cboProcType.value = "M" Then		'�系����ǰ 
		'-----------------------------------------------------------------------------------------------------
		' List Minor code	'�ܰ����� 
		'-----------------------------------------------------------------------------------------------------	
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1018", "''", "S") & "  AND MINOR_CD <> " & FilterVar("M", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboPrcCtrlInd, lgF0, lgF1, Chr(11))	

		'-----------------------------------------------------------------------------------------------------
		' List Minor code for Procurement Type(����Type)
		'-----------------------------------------------------------------------------------------------------
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1008", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboMatType, lgF0, lgF1, Chr(11))
		
		Call ggoOper.SetReqAttr(frm1.txtStdPrice, "N")  '����� 
'		Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice, "Q")  'ȸ�� 
	ElseIf frm1.cboProcType.value = "O" Then	'���ְ���ǰ 
		'-----------------------------------------------------------------------------------------------------
		' List Minor code	'�ܰ����� 
		'-----------------------------------------------------------------------------------------------------	
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1018", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboPrcCtrlInd, lgF0, lgF1, Chr(11))	

		'-----------------------------------------------------------------------------------------------------
		' List Minor code for Procurement Type(����Type)
		'-----------------------------------------------------------------------------------------------------
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1008", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboMatType, lgF0, lgF1, Chr(11))	
	ElseIf frm1.cboProcType.value = "P" Then	'����ǰ 
		'-----------------------------------------------------------------------------------------------------
		' List Minor code	'�ܰ����� 
		'-----------------------------------------------------------------------------------------------------	
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1018", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboPrcCtrlInd, lgF0, lgF1, Chr(11))	

		'-----------------------------------------------------------------------------------------------------
		' List Minor code for Procurement Type(����Type)
		'-----------------------------------------------------------------------------------------------------
		Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1008", "''", "S") & "  AND MINOR_CD <> " & FilterVar("10", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboMatType, lgF0, lgF1, Chr(11))	
	End If
	Call SetPrcCtrlInd
End Sub

Sub cboProcType_OnChange()		'���ޱ���	M/O/P
	lgBlnFlgChgValue = True		
	Call SetFieldForProcType(0)

	If frm1.cboProcType.value = "M" Then		'�系����ǰ		
		Call ggoOper.SetReqAttr(frm1.txtStdPrice, "N")			'����� 
'		Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice, "Q")		'ȸ�� 
		frm1.cboPrcCtrlInd.value = "S"
	ElseIf frm1.cboProcType.value = "O" Then	'���ְ���ǰ 
		
	ElseIf frm1.cboProcType.value = "P" Then	'����ǰ	

	End If
	Call SetPrcCtrlInd
End Sub

Sub cboMatType_OnChange() 
	lgBlnFlgChgValue = True
End Sub

Sub cboIssueType_OnChange() 
	lgBlnFlgChgValue = True		
End Sub

Sub cboOrderFrom_OnChange() 
	lgBlnFlgChgValue = True
	If frm1.cboOrderFrom.value = "R" Then
		Call ggoOper.SetReqAttr(frm1.txtReorderPoint, "N")
	Else
		frm1.txtReorderPoint = "0"
		Call ggoOper.SetReqAttr(frm1.txtReorderPoint, "Q")
	End If
End Sub 

Sub cboLotSizing_OnChange() 
	lgBlnFlgChgValue = True
	Call SetLotSizing()	
End Sub 

Sub cboPrcCtrlInd_OnChange()
	Call SetPrcCtrlInd()
	lgBlnFlgChgValue = True		
End Sub 

Sub cboProdEnv_OnChange() 
	lgBlnFlgChgValue = True		
End Sub

'********************************************************************************************************
'*	Radio OnClick																						*
'********************************************************************************************************
Function rdoCollectFlg1_OnClick()
	If lgRdoColOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoColOldVal =1
	
	Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "N")
End Function

Function rdoCollectFlg2_OnClick() 
	If lgRdoColOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoColOldVal = 2
	
	Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "Q")
	
	frm1.txtWorkCenter.value = ""
	frm1.txtWcNm.value = ""
End Function

Function rdoLotNoFlg1_OnClick() 
	If lgRdoLotOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoLotOldVal = 1
End Function

Function rdoLotNoFlg2_OnClick() 
	If lgRdoLotOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoLotOldVal = 2
End Function

Function rdoMPSItem1_OnClick() 
	If lgRdoMpsOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoMpsOldVal = 1
End Function

Function rdoMPSItem2_OnClick() 
	If lgRdoMpsOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoMpsOldVal = 2
End Function

Function sub_rdoMRPFlg1_OnClick() 
	Call ggoOper.SetReqAttr(frm1.cboLotSizing, "N")
	frm1.cboLotSizing.value = "L"
	Call ggoOper.SetReqAttr(frm1.cboOrderFrom, "Q")
	frm1.cboOrderFrom.value = ""
	frm1.txtReorderPoint = "0"
	Call ggoOper.SetReqAttr(frm1.txtReorderPoint, "Q")
End Function

Function rdoMRPFlg1_OnClick() 
	If lgRdoMrpOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoMrpOldVal = 1
	Call ggoOper.SetReqAttr(frm1.cboLotSizing, "N")
	frm1.cboLotSizing.value = "L"
	Call ggoOper.SetReqAttr(frm1.cboOrderFrom, "Q")
	frm1.cboOrderFrom.value = ""
	frm1.txtReorderPoint = "0"
	Call ggoOper.SetReqAttr(frm1.txtReorderPoint, "Q")
End Function

Function rdoMRPFlg2_OnClick() 
	If lgRdoMrpOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoMrpOldVal = 2
	Call ggoOper.SetReqAttr(frm1.cboLotSizing , "Q")
	frm1.cboLotSizing.value = "L"
	Call ggoOper.SetReqAttr(frm1.cboOrderFrom, "N")
	frm1.cboOrderFrom.value = "R"
	Call ggoOper.SetReqAttr(frm1.txtReorderPoint, "N")
End Function

Function rdoTrackingItem1_OnClick()
	If lgRdoTrkOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoTrkOldVal = 1
End Function  

Function rdoTrackingItem2_OnClick()
	If lgRdoTrkOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoTrkOldVal = 2
End Function  

Function rdoFinalInspType1_OnClick()
	If lgRdoFinOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoFinOldVal = 1
End Function  

Function rdoFinalInspType2_OnClick()
	If lgRdoFinOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoFinOldVal = 2
End Function  

Function rdoIssueInspType1_OnClick()
	If lgRdoRndOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoIssOldVal = 1
End Function 

Function rdoIssueInspType2_OnClick()
	If lgRdoIssOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoIssOldVal = 2
End Function 

Function rdoMfgInspType1_OnClick()
	If lgRdoPrdOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoPrdOldVal = 1
End Function

Function rdoMfgInspType2_OnClick()
	If lgRdoPrdOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoPrdOldVal = 2
End Function

Function rdoPurInspType1_OnClick()
	If lgRdoRecOldVal = 1 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoRecOldVal = 1
End Function

Function rdoPurInspType2_OnClick()
	If lgRdoRecOldVal = 2 Then Exit Function
	
	lgBlnFlgChgValue = True		
	lgRdoRecOldVal = 2
End Function

'=========================================  5.1.1 FncQuery()  ===========================================
'=	Event Name : FncQuery																				=
'=	Event Desc : This function is related to Query Button of Main ToolBar								=
'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													 '��: Processing is NG 

	Err.Clear															 '��: Protect system from crashing 

	 '------ Check previous data area ------ 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			 '��: "Will you destory previous data" 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	 '------ Erase contents area ------ 
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
	Call ggoOper.ClearField(Document, "2")								 '��: Clear Contents  Field 
	Call SetDefaultVal
	Call InitVariables													 '��: Initializes local global variables 
	
	 '------ Check condition area ------ 
	 If Not ChkFieldByCell(frm1.txtPlantCd,"A",1) Then Exit Function
	 If Not ChkFieldByCell(frm1.txtItemCd,"A",1) Then Exit Function

	 '------ Query function call area ------ 
	If DbQuery = False Then   
		Exit Function           
    End If 

	FncQuery = True														 '��: Processing is OK 
End Function
	

'===========================================  5.1.2 FncNew()  ===========================================
'=	Event Name : FncNew																					=
'=	Event Desc : This function is related to New Button of Main ToolBar									=
'========================================================================================================
Function FncNew()
	Dim IntRetCD 
	Dim strPlantCd
	Dim strPlantNm
	
	FncNew = False                                                         '��: Processing is NG															 '��: Protect system from crashing 

	 '------ Check previous data area ------ 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	 '------ Erase condition area ------ 
	 '------ Erase contents area ------ 
	strPlantCd = frm1.txtPlantCd.value
	strPlantNm = frm1.txtPlantNm.value
	
	Call ggoOper.ClearField(Document, "A")									'��: Clear Condition Field
	Call pvLockField("N")													'��: Lock  Suitable  Field
	
	frm1.txtPlantCd.value = strplantcd
	frm1.txtPlantNm.value = strplantnm
	
	Call SetDefaultVal
	Call InitVariables														'��: Initializes local global variables
	Call SetToolbar("11101000000011")												 '��: ��ư ���� ���� 
	
	Call ggoOper.SetReqAttr(frm1.txtReorderPoint, "Q")		'ȸ�� 
	frm1.txtItemCd1.focus
	Set gActiveElement = document.activeElement
	
	FncNew = True															'��: Processing is OK
End Function
	

'===========================================  5.1.3 FncDelete()  ========================================
'=	Event Name : FncDelete																				=
'=	Event Desc : This function is related to Delete Button of Main ToolBar								=
'========================================================================================================
Function FncDelete()
	Dim IntRetCD
	
	FncDelete = False												 '��: Processing is NG 
	
	 '------ Precheck area ------ 
	If lgIntFlgMode <> parent.OPMD_UMODE Then								 'Check if there is retrived data 
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If
	
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '��: "Will you destory previous data"	
	If IntRetCD = vbNo Then
		Exit Function	
	End If
	
	If DbDelete = False Then   
		Exit Function           
    End If 
	
	FncDelete = True												 '��: Processing is OK 
End Function


'===========================================  5.1.4 FncSave()  ==========================================
'=	Event Name : FncSave																				=
'=	Event Desc : This function is related to Save Button of Main ToolBar								=
'========================================================================================================
Function FncSave()
	Dim IntRetCD
	
	FncSave = False													 '��: Processing is NG 
	Err.Clear														 '��: Protect system from crashing 
	
	 '------ Precheck area ------ 
	If lgBlnFlgChgValue = False Then								 'Check if there is retrived data 
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")					 '��: No data changed!! 
	    Exit Function
	End If
		
	 '------ Check contents area ------ 
	If Not chkField(Document, "2") Then								'��: Check contents area 
		Exit Function
	End If

	If (frm1.cboProcType.value = "M" Or frm1.cboProcType.value = "O") And frm1.cboMatType.value <> "10" then
		IntRetCD = DisplayMsgBox("122725", "X", "X", "X")					 '����Ÿ���� VMI, ���������� ��쿡�� ���ޱ����� ����ǰ�� ��쿡�� �����մϴ�.' 
		Exit Function
	End if
	 '------ Save function call area ------ 
	If DbSave = False Then   
		Exit Function           
    End If     								 '��: Save db data 
	
	FncSave = True													 '��: Processing is OK 
End Function

'===========================================  5.1.5 FncCopy()  ==========================================
'=	Event Name : FncCopy																				=
'=	Event Desc : This function is related to Copy Button of Main ToolBar								=
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    	
	lgIntFlgMode = parent.OPMD_CMODE													'��: Indicates that current mode is Crate mode

	 '------ ���Ǻ� �ʵ带 �����Ѵ�. ------ 
														'��: Focus�� ù��° Tab���� �̵���Ų��	
	Call SetPrcCtrlInd()
	Call SetToolbar("11101000000011")												 '��: ��ư ���� ���� 
	
	lgBlnFlgChgValue = True
	
	frm1.txtItemCd.value = ""
	frm1.txtItemNm.value = ""
	frm1.txtItemCd1.value = ""
	frm1.txtItemNm1.value = ""
	frm1.txtValidFromDt.Text	= StartDate
	frm1.txtValidToDt.Text		= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")
	
	Call ggoOper.SetReqAttr(frm1.txtItemCd1, "N")
	Call ggoOper.SetReqAttr(frm1.txtValidFromDt, "N")
	
	frm1.txtItemCd1.focus
	Set gActiveElement = document.activeElement
	
	
	If frm1.cboPrcCtrlInd.value = "M" Then
		frm1.txtStdPrice.Text = ""
		frm1.txtPrevStdPrice.Text = ""	
'		Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice,"N")	'����� 
	Else
		frm1.txtMoveAvgPrice.Text = ""	
		frm1.txtPrevStdPrice.Text = ""	
		Call ggoOper.SetReqAttr(frm1.txtStdPrice,"N")	'����� 
	End If
End Function

'============================================  5.1.9 FncPrint()  ========================================
'=	Event Name : FncPrint																				=
'=	Event Desc : This function is related to Print Button of Main ToolBar								=
'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function


'============================================  5.1.10 FncPrev()  ========================================
'=	Event Name : FncPrev																				=
'=	Event Desc : This function is related to Previous Button											=
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                             '��: 
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables

	Err.Clear															'��: Protect system from crashing

	LayerShowHide(1)
		
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&PrevNextFlg=" & "P"								'��: ��ȸ ���� ����Ÿ 
	
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
End Function

'============================================  5.1.11 FncNext()  ========================================
'=	Event Name : FncNext																				=
'=	Event Desc : This function is related to Next Button												=
'========================================================================================================
Function FncNext()

    Dim strVal
    Dim IntRetCD

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                             '��: 
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables

	Err.Clear															'��: Protect system from crashing

	LayerShowHide(1)
										'��: �۾������� ǥ��	

	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&PrevNextFlg=" & "N"								'��: ��ȸ ���� ����Ÿ 
	
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
End Function

'===========================================  5.1.12 FncExcel()  ========================================
'=	Event Name : FncExcel																				=
'=	Event Desc : This function is related to Excel Button of Main ToolBar								=
'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLE)
End Function

'===========================================  5.1.13 FncFind()  =========================================
'=	Event Name : FncFind																				=
'=	Event Desc :																						=
'========================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
 Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'=============================================  5.2.1 DbQuery()  ========================================
'=	Event Name : DbQuery																				=
'=	Event Desc : This function is data query and display												=
'========================================================================================================
Function DbQuery()
	Err.Clear															'��: Protect system from crashing

	DbQuery = False														'��: Processing is NG

	LayerShowHide(1)
		
	Dim strVal

	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&PrevNextFlg=" & ""	
	
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

	DbQuery = True														'��: Processing is NG
End Function
	
'=============================================  5.2.2 DbSave()  =========================================
'=	Event Name : DbSave																					=
'=	Event Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ���										=
'========================================================================================================
Function DbSave()
	Dim rtnVal
	
	Err.Clear															'��: Protect system from crashing

	DbSave = False														'��: Processing is NG
	rtnVal = ChkValidData
	
	If rtnVal = False Then 
		Exit Function
	End If
	LayerShowHide(1)
	
	Dim strVal

	With frm1
		.txtMode.value = parent.UID_M0002										'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	End With

	DbSave = True														'��: Processing is NG
End Function

'=============================================  5.2.3 DbDelete()  =======================================
'=	Event Name : DbDelete																				=
'=	Event Desc : This function delete data																=
'========================================================================================================
Function DbDelete()
	Err.Clear															'��: Protect system from crashing

	DbDelete = False													'��: Processing is NG

	LayerShowHide(1)
		
	Dim strVal

	strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003					'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ���� ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ���� ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

	DbDelete = True														'��: Processing is NG
End Function
	
'=============================================  5.2.4 DbQueryOk()  ======================================
'=	Event Name : DbQueryOk																				=
'=	Event Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű�	=
'========================================================================================================
Function DbQueryOk()													 '��: ��ȸ ������ ������� 
	 '------ Reset variables area ------ 
	lgIntFlgMode = parent.OPMD_UMODE											 '��: Indicates that current mode is Update mode 
	Call pvLockField("Q")														 '��: This function lock the suitable field 
	Call SetToolbar("11111000111111")
	Call SetPrcCtrlInd()
	Call SetLotSizing()
	Call SetFieldForPhantom()
	Call SetFieldForAccout(1)
	Call SetFieldForProcType(1)
	Call SetOrderCreate
	
	If frm1.rdoCollectFlg1.checked = True Then
		Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtWorkCenter, "Q")
	End If
	
	If frm1.rdoMRPFlg1.checked = True Then
		Call ggoOper.SetReqAttr(frm1.cboLotSizing , "N")
		Call ggoOper.SetReqAttr(frm1.cboOrderFrom , "Q")
		Call ggoOper.SetReqAttr(frm1.txtReorderPoint , "Q")
	Else
		If frm1.cboOrderFrom.value = "R" Then
			Call ggoOper.SetReqAttr(frm1.cboLotSizing, "Q")
			Call ggoOper.SetReqAttr(frm1.cboOrderFrom , "N")
			Call ggoOper.SetReqAttr(frm1.txtReorderPoint , "N")
		Else
			Call ggoOper.SetReqAttr(frm1.cboLotSizing, "Q")
			Call ggoOper.SetReqAttr(frm1.cboOrderFrom , "N")
			Call ggoOper.SetReqAttr(frm1.txtReorderPoint , "Q")
		End	If
	End If
	
'	Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice, "Q")  'ȸ�� 
	
	If frm1.cboPrcCtrlInd.value = "M" then
		blncboPrcCtrlIndIsM = True
		Call ggoOper.SetReqAttr(frm1.txtStdPrice,"Q")		'ȸ�� 
		Call ggoOper.SetReqAttr(frm1.txtPrevStdPrice,"Q")	'ȸ�� 
'		Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice,"Q")	'ȸ�� 
	Else
		blncboPrcCtrlIndIsM = False
		Call ggoOper.SetReqAttr(frm1.txtStdPrice,"N")		'��� 
		Call ggoOper.SetReqAttr(frm1.txtPrevStdPrice,"Q")	'ȸ�� 
'		Call ggoOper.SetReqAttr(frm1.txtMoveAvgPrice,"Q")	'ȸ�� 
	End IF
	
	frm1.cboAccount.focus
	Set gActiveElement = document.activeElement 

	IF frm1.txtPhantomFlg.value = "Y" Then
		frm1.cboOrderFrom.value = ""
		Call ggoOper.SetReqAttr(frm1.txtStdPrice,"Q")
		Call ggoOper.SetReqAttr(frm1.cboOrderFrom,"Q")		'ȸ�� 
		Call ggoOper.SetReqAttr(frm1.txtReorderPoint,"Q")	'ȸ�� 
	End If
	
	lgBlnFlgChgValue = False	
End Function

'=============================================  5.2.5 DbSaveOk()  =======================================
'=	Event Name : DbSaveOk																				=
'=	Event Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű�	=
'========================================================================================================
Function DbSaveOk()
	frm1.txtItemCd.value = frm1.txtItemCd1.value 														'��: ���� ������ ���� ���� 
	Call InitVariables
	Call MainQuery()
End Function

'=============================================  5.2.6 DbDeleteOk()  =====================================
'=	Event Name : DbSaveOk																				=
'=	Event Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű�	=
'========================================================================================================
Function DbDeleteOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : pvLockField
' Function Desc : ggoOperLockField ��� 
'========================================================================================
Function pvLockField(byVal pvFlag) 
	If pvFlag = "Q" Then
		Call LockHTMLField(frm1.txtItemCd1,"P")
		Call LockObjectField(frm1.txtValidFromDt,"P")
	ElseIf pvFlag = "N" Then
		Call LockHTMLField(frm1.txtItemCd1,"R")
		Call LockObjectField(frm1.txtValidFromDt,"R")
	End If
End Function
