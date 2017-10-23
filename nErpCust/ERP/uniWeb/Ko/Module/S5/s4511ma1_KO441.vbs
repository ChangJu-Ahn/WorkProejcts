Option Explicit                               

Const BIZ_PGM_ID = "s4511mb1.asp"												'��: Head Query �����Ͻ� ���� ASP��
Const BIZ_PGM_JUMP_ID = "s4512ma1"

Const C_PopShipToParty	= 1		' ��ǰó
Const C_PopInvMgr		= 2		' �������
Const C_PopTransMeth	= 3		' ��۹��

Const TAB1 = 1                  '��: Tab�� ��ġ 
Const TAB2 = 2

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgBlnFlgChgValue1			'2002.12.18 ��ǰó������ ���� ����
Dim lgBlnFlgChgValue2			'2002.12.18 ������� ���� ����
Dim lgIntGrpCount              ' Group View Size�� ������ ����
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim IsOpenPop						' Popup
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i   '@@@CommonQueryRs �� ���� ���� 

'=========================================
Sub FormatField()
    With frm1
        ' ��¥ OCX Foramt ����
        Call FormatDATEField(.txtPlanned_gi_dt)
        Call FormatDATEField(.txtArriv_dt)
    End With
End Sub

'=========================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' ��¥ OCX
        Call LockObjectField(.txtPlanned_gi_dt, "R")
        Call LockObjectField(.txtArriv_dt, "O")

        If pvFlag = "N" Then
			Call LockHTMLField(.txtDnNo, "O")	
			Call LockHTMLField(.chkSoNo, "O")	
        End If
    End With

End Sub

'=========================================
Sub LockFieldQuery()
    With frm1
		Call LockHTMLField(.txtDnNo, "P")	
		Call LockHTMLField(.chkSoNo, "P")	
    End With
End Sub

'=====================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    

    lgBlnFlgChgValue1 = False                    
    lgBlnFlgChgValue2 = False                    
End Sub

'=====================================================
Sub SetDefaultVal()	
	frm1.txtConDnNo.focus
	frm1.txtPlanned_gi_dt.Text = EndDate
	frm1.txtDlvy_dt.value = EndDate				' Design�� �ʵ� ���̸� �����ϱ� ���� Text �ʵ带 �����
	
	lgBlnFlgChgValue = False
	lgBlnFlgChgValue1 = False
	lgBlnFlgChgValue2 = False
End Sub

'========================================
Function OpenSORef()
	Dim iCalledAspName
	Dim arrRet
	Dim PlannedGiDate
	
	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call DisplayMsgBox("204150", "X", "X", "X")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	iCalledAspName = AskPRAspName("S4511AA1_KO441")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4511AA1_KO441", "x")
		IsOpenPop = False
		exit Function
	end if

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If Err.Number <> 0 Then	Err.Clear 
	Else
		Call SetSORef(arrRet)
		Call SetToolbar("11101000000111")					'��: ��ư ���� ����
	End If	
End Function

'========================================
Function OpenConDnReqNo()
	Dim iCalledAspName
	Dim strRet
	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True

	iCalledAspName = AskPRAspName("S4511PA1_KO441")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4511PA1_KO441", "x")
		IsOpenPop = False
		exit Function
	end if
		
	strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=N", Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtConDnNo.focus
			
	If strRet <> "" Then
		frm1.txtConDnNo.value = strRet		
	End If	

End Function

'========================================
' Popup
'=========================================
Function OpenPopUp(Byval pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop Then Exit Function

	IsOpenPop = True

	With frm1
		Select Case pvIntWhere

			'��ǰó
			Case C_PopShipToParty
				If .txtShip_to_party.readOnly Then
					IsOpenPop = False
					Exit Function
				End If

				iArrParam(1) = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"								
				iArrParam(2) = Trim(.txtShip_to_party.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ")"						
	
				iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
				iArrField(2) = "ED10" & Parent.gColSep & "BP.CONTRY_CD"
				iArrField(3) = "ED20" & Parent.gColSep & "CT.COUNTRY_NM"
    
				iArrHeader(0) = .txtShip_to_party.alt					
				iArrHeader(1) = .txtShip_to_partyNm.alt					
				iArrHeader(2) = "����"
				iArrHeader(3) = "������"

				.txtShip_to_party.focus

			'�����
			Case C_PopInvMgr
				If .txtInvMgr.readOnly Then
					IsOpenPop = False
					Exit Function
				End If

				iArrParam(1) = "dbo.B_MINOR"
				iArrParam(2) = Trim(.txtInvMgr.value)
				iArrParam(3) = ""											
				iArrParam(4) = "MAJOR_CD = " & FilterVar("I0004", "''", "S") & ""
				
				iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
							
				iArrHeader(0) = .txtInvMgr.alt						
				iArrHeader(1) = .txtInvMgrNm.alt						

				.txtInvMgr.focus
				
			'��۹��
			Case C_PopTransMeth
				If .txtTrans_meth.readOnly Then
					IsOpenPop = False
					Exit Function
				End If

				iArrParam(1) = "dbo.B_MINOR"
				iArrParam(2) = Trim(.txtTrans_meth.value)
				iArrParam(3) = ""											
				iArrParam(4) = "MAJOR_CD = " & FilterVar("B9009", "''", "S") & ""
				
				iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
							
				iArrHeader(0) = .txtTrans_meth.alt						
				iArrHeader(1) = .txtTrans_meth_nm.alt						

				.txtTrans_meth.focus

		End Select
	End With
	
	iArrParam(0) = iArrHeader(0)							' �˾� Title
	iArrParam(5) = iArrHeader(0)							' ��ȸ���� ��Ī

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then
		OpenPopUp = SetPopUp(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function OpenZip()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtShip_to_party.value) = "" Then
		MsgBox "��ǰó�� ���� �Է��ϼ���", vbInformation, parent.gLogoName
		frm1.txtShip_to_party.focus 
		IsOpenPop = False			
		Exit Function
	End IF

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtZIP_cd.value)
	arrParam(1) = ""
	arrParam(2) = Trim(frm1.txtHCntryCd.value)

	arrRet = window.showModalDialog("../../comasp/ZipPopup.asp", Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	frm1.txtZIP_cd.focus

	If arrRet(0) <> "" Then
		With frm1
			.txtZIP_cd.value = arrRet(0)
			.txtADDR1_Dlv.value = arrRet(1)		
			.txtSTP_Inf_No.value = ""
			lgBlnFlgChgValue1 = True
		End With
	End If	
			
End Function

'========================================
Function OpenTransCo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ȸ��"							
	arrParam(1) = "B_MAJOR A , B_MINOR B"						
	arrParam(2) = ""										
	arrParam(3) = ""									
	arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9031", "''", "S") & " "				
	arrParam(5) = "���ȸ��"							

	arrField(0) = "B.MINOR_CD"								
	arrField(1) = "B.MINOR_NM"								

	arrHeader(0) = "���ȸ��"							
	arrHeader(1) = "���ȸ���"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtTransCo.focus
	
	If arrRet(0) <> "" Then
		Call SetTransCo(arrRet)
	End If
End Function

'========================================
Function OpenVehicleNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = "������ȣ"							
	arrParam(1) = "B_MAJOR A , B_MINOR B"						
	arrParam(2) = ""			
	arrParam(3) = ""									
	arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9032", "''", "S") & " "				
	arrParam(5) = "������ȣ"							

	arrField(0) = "B.MINOR_CD"								
	arrField(1) = "B.MINOR_NM"								

	arrHeader(0) = "����������ȣ"							
	arrHeader(1) = "������ȣ"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtVehicleNo.focus
	
	If arrRet(0) <> "" Then
		Call SetVehicleNo(arrRet)
	End If
End Function

'========================================
Function SetSORef(ByRef prArrRet) 

	With frm1
		.txtShip_to_party.value = prArrRet(0)			'��ǰó
		.txtShip_to_partyNm.value = prArrRet(1)			'��ǰó��
		.txtDlvy_dt.value = prArrRet(2)					'������	
		.txtMovType.value = prArrRet(3)					'��������
		.txtMovTypeNm.value = prArrRet(4)				'�������¸�
		.txtSo_no.value = prArrRet(5)					'���ֹ�ȣ
		.txtTempSoNo.value = prArrRet(5)					'���ֹ�ȣ
		.chkSoNo.checked  = True
'		.chkSoNo.disabled = True
		.txtSales_Grp.value = prArrRet(6)				'�����׷�
		.txtSales_GrpNm.value = prArrRet(7)				'�����׷��	
		.txtSo_Type.value = prArrRet(8)					'��������	
		.txtSo_TypeNm.value = prArrRet(9)				'�������¸�
		.txtTrans_meth.value = prArrRet(10)				'��۹��
		.txtTrans_meth_nm.value = prArrRet(11)			'��۹����
		.txtPlantCd.value = prArrRet(12)				'����
		.txtPlantNm.value = prArrRet(13)				'�����
		
		Call txtShip_to_party_Onchange()

		.txtDnNo.focus
	End With	
	lgBlnFlgChgValue = True

End Function

'========================================
Function SetPopUp(ByVal pvArrRet,ByVal pvIntWhere)

	With frm1
		Select Case pvIntWhere
			Case C_PopShipToParty
				.txtShip_to_party.value = pvArrRet(0)
				.txtShip_to_partyNm.value = pvArrRet(1)
				.txtHCntryCd.value = pvArrRet(2)
		
			Case C_PopInvMgr
				.txtInvMgr.value = pvArrRet(0)
				.txtInvMgrNm.value = pvArrRet(1) 

			Case C_PopTransMeth
				.txtTrans_meth.value = pvArrRet(0)
				.txtTrans_meth_nm.value = pvArrRet(1) 
		End Select
	End With
	
	lgBlnFlgChgValue = True		
End Function

'========================================
Function SetShipToPlceRef(Byval arrRet)
	on error resume next 
	
	frm1.txtSTP_Inf_No.value = arrRet(0)			'��ǰó��������ȣ	
	frm1.txtZIP_cd.value = arrRet(1)				'�����ȣ
	frm1.txtADDR1_Dlv.value = arrRet(2)				'��ǰ�ּ�1
	frm1.txtADDR2_Dlv.value = arrRet(3)				'��ǰ�ּ�2	
	frm1.txtADDR3_Dlv.value = arrRet(4)				'��ǰ�ּ�3
	frm1.txtReceiver.value = arrRet(5)				'�μ��ڸ�
	frm1.txtTel_No1.value = arrRet(6)				'��ȭ��ȣ1
	frm1.txtTel_No2.value = arrRet(7)				'��ȭ��ȣ2
	lgBlnFlgChgValue1 = True

End Function

'========================================
Function SetTrnsMethRef(Byval arrRet)	 
	
	frm1.txtTrnsp_Inf_No.value = arrRet(0)			'���������ȣ
	frm1.txtTransCo.value = arrRet(1)				'���ȸ��
	frm1.txtDriver.value = arrRet(2)				'�����ڸ�
	frm1.txtVehicleNo.value = arrRet(3)				'������ȣ	
	frm1.txtSender.value = arrRet(4)				'�ΰ��ڸ�
	lgBlnFlgChgValue2 = True

End Function

'========================================
Function SetTransCo(arrRet)	
	frm1.txtTransCo.value = arrRet(1)
	IF frm1.txtTrnsp_Inf_No.value <> "" Then
		frm1.txtTrnsp_Inf_No.value = ""
	End If
		
	lgBlnFlgChgValue2 = True
End Function

'========================================
Function SetVehicleNo(arrRet)
	frm1.txtVehicleNo.value = arrRet(1)
	IF frm1.txtTrnsp_Inf_No.value <> "" Then
		frm1.txtTrnsp_Inf_No.value = ""
	End If
		
	lgBlnFlgChgValue2 = True		
End Function

' ���ó���� ��� ���� �Ұ���� �����Ѵ�.
'========================================
Sub PostFlagProtect()
	With ggoOper
		Call .SetReqAttr(frm1.txtPlanned_gi_dt, "Q")		'�������
		Call .SetReqAttr(frm1.txtShip_to_place, "Q")		'��ǰ���
		Call .SetReqAttr(frm1.txtTrans_meth, "Q")		'��۹��
		Call .SetReqAttr(frm1.txtShip_to_party, "Q")		'��ǰó
		Call .SetReqAttr(frm1.txtRemark, "Q")			'���
		Call .SetReqAttr(frm1.txtArriv_dt, "Q")			'������ǰ��
		Call .SetReqAttr(frm1.txtArriv_Tm, "Q")			'��ǰ�ð�
		Call .SetReqAttr(frm1.txtInvMgr, "Q")			'�����
	End With
End Sub

'=====================================================
Sub PostFlagRelease()
	With ggoOper
		Call .SetReqAttr(frm1.txtPlanned_gi_dt, "N")		'�������
		Call .SetReqAttr(frm1.txtShip_to_party, "N")		'��ǰó

		Call .SetReqAttr(frm1.txtShip_to_place, "D")		'��ǰ���
		Call .SetReqAttr(frm1.txtTrans_meth, "D")		'��۹��
		Call .SetReqAttr(frm1.txtRemark, "D")			'���
		Call .SetReqAttr(frm1.txtArriv_dt, "D")			'������ǰ��
		Call .SetReqAttr(frm1.txtArriv_Tm, "D")			'��ǰ�ð�
		Call .SetReqAttr(frm1.txtInvMgr, "D")			'�����
	End With
End Sub

'=====================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877
	
	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtConDnNo.value

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtConDnNo.value =  arrVal(0)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If

End Function

'========================================
Function JumpChgCheck()

	Dim IntRetCD

	'************ �̱��� ��� **************
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call CookiePage(1)
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function

'========================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call FormatField()
    Call LockFieldInit("L")
    Call SetDefaultVal
    Call InitVariables                                                      
    Call SetToolbar("11100000000011")										'��: ��ư ���� ����
	Call CookiePage(0)
	Call ChangeTabs(TAB1)
	gIsTab = "Y" : gTabMaxCnt = 2
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================
Sub chkSoNo_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================
Function btnShipToPlceRef_OnClick()
	Dim iCalledAspName
	Dim arrRet
	Dim ShipToPartyCd
	
	On Error Resume Next
	
    If Trim(frm1.txtShip_to_party.value) = "" Then                                      'Check if there is retrived data
        Call DisplayMsgBox("204256", "X", "X", "X")  '�� �ٲ�κ�
        Exit Function
    End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("S4111RA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111RA1", "x")
		IsOpenPop = False
		exit Function
	end if

	ShipToPartyCd = Trim(frm1.txtShip_to_party.value)
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent , ShipToPartyCd),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
		Call SetShipToPlceRef(arrRet)
		Call SetToolbar("11101000000111")					'��: ��ư ���� ����
	End If	

End Function

'========================================
Function btnTrnsMethRef_OnClick()
	Dim iCalledAspName
	Dim arrRet

	
	On Error Resume Next

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("S4111RA2")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111RA2", "x")
		IsOpenPop = False
		exit Function
	end if
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent , ""),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
		Call SetTrnsMethRef(arrRet)
		Call SetToolbar("11101000000111")					'��: ��ư ���� ����
	End If	

End Function

'========================================
Sub txtPlanned_gi_dt_Change()
	lgBlnFlgChgValue = True
End Sub

'������ǰ��
Sub txtArriv_dt_Change()
	lgBlnFlgChgValue = True
End Sub

'��ǰó ������, ������� ���濩��
'========================================
Sub txtZip_cd_Onchange()
	lgBlnFlgChgValue1 = True
	IF frm1.txtSTP_Inf_No.value <> "" Then
		 frm1.txtSTP_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtReceiver_Onchange()
	lgBlnFlgChgValue1 = True
	IF frm1.txtSTP_Inf_No.value <> "" Then
		 frm1.txtSTP_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtADDR1_Dlv_Onchange()
	lgBlnFlgChgValue1 = True
	IF frm1.txtSTP_Inf_No.value <> "" Then
		 frm1.txtSTP_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtADDR2_Dlv_Onchange()
	lgBlnFlgChgValue1 = True
	IF frm1.txtSTP_Inf_No.value <> "" Then
		 frm1.txtSTP_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtADDR3_Dlv_Onchange()
	lgBlnFlgChgValue1 = True
	IF frm1.txtSTP_Inf_No.value <> "" Then
		 frm1.txtSTP_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtTel_No1_Onchange()
	lgBlnFlgChgValue1 = True
	IF frm1.txtSTP_Inf_No.value <> "" Then
		 frm1.txtSTP_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtTel_No2_Onchange()
	lgBlnFlgChgValue1 = True
	IF frm1.txtSTP_Inf_No.value <> "" Then
		 frm1.txtSTP_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtTransCo_Onchange()		
	lgBlnFlgChgValue2 = True	
	IF frm1.txtTrnsp_Inf_No.value <> "" Then
		 frm1.txtTrnsp_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtSender_Onchange()	
	lgBlnFlgChgValue2 = True
	IF frm1.txtTrnsp_Inf_No.value <> "" Then
		 frm1.txtTrnsp_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtVehicleNo_Onchange()
	lgBlnFlgChgValue2 = True
	IF frm1.txtTrnsp_Inf_No.value <> "" Then
		 frm1.txtTrnsp_Inf_No.value = ""
	End If
End Sub

'========================================
Sub txtDriver_Onchange()
	lgBlnFlgChgValue2 = True
	IF frm1.txtTrnsp_Inf_No.value <> "" Then
		 frm1.txtTrnsp_Inf_No.value = ""
	End If
End Sub

' ��ǰó onchange�� �����ڵ� ���������� ����
'========================================
Sub txtShip_to_Party_Onchange()
	lgBlnFlgChgValue = True
	If Trim(frm1.txtShip_to_party.value) = "" Then
		Dim iArrRs
		Redim iArrRs(8)
		Call SetSTPInfo(iArrRs)
	Else
		If GetContryCd Then	Call GetShipToPartyInfo
	End If
End Sub

'========================================
Function GetContryCd
	Dim iContryCd
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs 

	GetContryCd = False
	
	iStrSelectList = "BP_NM, CONTRY_CD"
	iStrFromList = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"
	iStrWhereList = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") " & _
					"AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ") " & _
					"AND BP_CD =  " & FilterVar(frm1.txtship_to_party.value, "''", "S") & ""

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList, iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		frm1.txtShip_to_partyNm.value = iArrRs(1)
		frm1.txtHCntryCd.value = iArrRs(2)
		
		GetContryCd = True
	Else
		Call DisplayMsgBox("126121","X","X","X")
		frm1.txtHCntryCd.value = ""	
		frm1.txtShip_to_partyNm.value = "" 
		
		Redim iArrRs(8)
		Call SetSTPInfo(iArrRs)
	End If
End Function

'********************************************************************************
' 2003.02.05 SMJ
' ��ǰó ������ ���������� default�� �ֱ��� ��ǰó������ ���������� ����
' ��ǰó ������ ���� ���� �ŷ�ó�� ������ default�� ������
'********************************************************************************
Sub GetShipToPartyInfo
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs
	
	iStrSelectList = "top 1  STP_INFO_NO, ZIP_CD, ADDR1, ADDR2, ADDR3, RECEIVER, TEL_NO1, TEL_NO2 "
	iStrFromList = " S_SHIP_TO_PARTY_INFO "
	iStrWhereList = " SHIP_TO_PARTY =  " & FilterVar(frm1.txtship_to_party.value, "''", "S") & "  ORDER BY STP_INFO_NO DESC"
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		Call SetSTPInfo(iArrRs)
		lgBlnFlgChgValue1 = True
	Else
		iStrSelectList = "'' STP_INFO_NO, ZIP_CD, ADDR1, ADDR2, '' ADDR3, '' RECEIVER, TEL_NO1, TEL_NO2 "
		iStrFromList = " B_BIZ_PARTNER "
		iStrWhereList = " BP_CD =  " & FilterVar(frm1.txtship_to_party.value, "''", "S") & " "
	
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrRs = Split(iStrRs, Chr(11))
			Call SetSTPInfo(iArrRs)
			lgBlnFlgChgValue1 = True
		Else
			Redim iArrRs(8)
			Call SetSTPInfo(iArrRs)
			
			If Err.number <> 0 Then Err.Clear 
		End If
	End If
End Sub

' ��ǰó ������ ����
'========================================
Sub SetSTPInfo(ByRef prArrSTPInfo)
	With frm1
		.txtSTP_Inf_No.value = prArrSTPInfo(1)			'��ǰó��������ȣ	
		.txtZIP_cd.value = prArrSTPInfo(2)				'�����ȣ
		.txtADDR1_Dlv.value = prArrSTPInfo(3)			'��ǰ�ּ�1
		.txtADDR2_Dlv.value = prArrSTPInfo(4)			'��ǰ�ּ�2	
		.txtADDR3_Dlv.value = prArrSTPInfo(5)			'��ǰ�ּ�3
		.txtReceiver.value = prArrSTPInfo(6)			'�μ��ڸ�
		.txtTel_No1.value = prArrSTPInfo(7)				'��ȭ��ȣ1
		.txtTel_No2.value = prArrSTPInfo(8)				'��ȭ��ȣ2
	End With
End Sub

'========================================
Sub txtPlanned_gi_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPlanned_gi_dt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtPlanned_gi_dt.Focus
	End If
End Sub

'========================================
Sub txtArriv_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtArriv_dt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtArriv_dt.Focus
	End If
End Sub

'========================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If Not chkFieldByCell(frm1.txtConDnNo, "A", TAB1) Then Exit Function

    If (lgBlnFlgChgValue = True Or lgBlnFlgChgValue1 = True or lgBlnFlgChgValue2 = True ) Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X") '�� �ٲ�κ�
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										
    Call InitVariables															
    Call SetDefaultVal
        
    Call DbQuery																'��: Query db data
    
    FncQuery = True																
        
End Function

'=====================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False                                                          
    
    If (lgBlnFlgChgValue = True Or lgBlnFlgChgValue1 = True or lgBlnFlgChgValue2 = True) Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") '�� �ٲ�κ�
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")
    Call LockFieldInit("N")
	Call PostFlagRelease()
	Call SetDefaultVal
    Call InitVariables															
    Call SetToolbar("11100000000011")

    FncNew = True																

End Function

'=====================================================
Function FncDelete() 
    
    FncDelete = False														
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")  '�� �ٲ�κ�
        Exit Function
    End If
    
	Dim IntRetCD
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")  '�� �ٲ�κ�
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
    Call DbDelete
    
    FncDelete = True                                                        
    
End Function

'=====================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    
	If (lgBlnFlgChgValue = False and lgBlnFlgChgValue1 = False and lgBlnFlgChgValue2 = False) Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")  '�� �ٲ�κ�
        Exit Function
    End If
    
	With frm1
        ' �Է��ʼ� �׸� Check
        If Not chkFieldByCell(.txtShip_to_party, "A", TAB1) Then Exit Function
        If Not chkFieldByCell(.txtPlanned_gi_dt, "A", TAB1) Then Exit Function
        
		' �ʵ���� Check
		If Not ChkFieldLengthByCell(.txtRemark, "A", TAB1) Then Exit Function
		If Not ChkFieldLengthByCell(.txtArriv_Tm, "A", TAB1) Then Exit Function
		If Not ChkFieldLengthByCell(.txtShip_to_place, "A", TAB1) Then Exit Function

		If Not ChkFieldLengthByCell(.txtReceiver, "A", TAB2) Then Exit Function
		If Not ChkFieldLengthByCell(.txtADDR1_Dlv, "A", TAB2) Then Exit Function
		If Not ChkFieldLengthByCell(.txtADDR2_Dlv, "A", TAB2) Then Exit Function
		If Not ChkFieldLengthByCell(.txtADDR3_Dlv, "A", TAB2) Then Exit Function
		If Not ChkFieldLengthByCell(.txtTel_No1, "A", TAB2) Then Exit Function
		If Not ChkFieldLengthByCell(.txtTel_No2, "A", TAB2) Then Exit Function
		If Not ChkFieldLengthByCell(.txtSender, "A", TAB2) Then Exit Function
		If Not ChkFieldLengthByCell(.txtDriver, "A", TAB2) Then Exit Function
    End With

    CAll DbSave

    FncSave = True                                                          
    
End Function

'=====================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=====================================================
Function FncPrev() 
    On Error Resume Next                                                    
End Function

'=====================================================
Function FncNext() 
    On Error Resume Next                                                    
End Function

'=====================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLE)
End Function

'=====================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE, False)
End Function

'=====================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True OR lgBlnFlgChgValue1 = True Or lgBlnFlgChgValue2 = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")   '�� �ٲ�κ�
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    FncExit = True
End Function

'=====================================================
Function DbDelete() 
    Err.Clear                                                               
    
    DbDelete = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							
    strVal = strVal & "&txtConDnNo=" & Trim(frm1.txtConDnNo.value)		
    
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbDelete = True                                                         

End Function

'=====================================================
Function DbDeleteOk()														
	Call MainNew()
End Function

'=====================================================
Function DbQuery() 
    
    Err.Clear                                                               
    
    DbQuery = False                                                         

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							
    strVal = strVal & "&txtConDnNo=" & Trim(frm1.txtConDnNo.value)			
    
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbQuery = True                                                          

End Function

'=====================================================
Function DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE
	lgBlnFlgChgValue = False
	lgBlnFlgChgValue1 = False
	lgBlnFlgChgValue2 = False
    
    Call LockFieldQuery()
	Call SetToolbar("11111000000111")

	If frm1.txtGoods_mv_no.value <> "" then
		Call PostFlagProtect()
	Else
		Call PostFlagRelease()
	End if
	
End Function

'=====================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Dim strVal

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = Parent.gUsrID 
		.txtUpdtUserId.value = Parent.gUsrID
		.txtlgBlnChgValue1.value = lgBlnFlgChgValue1		
		.txtlgBlnChgValue2.value = lgBlnFlgChgValue2		
		
		If .chkSoNo.checked = True Then
			.txtChkSoNo.value = "Y"
		Else
			.txtChkSoNo.value = "N"
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           
    
End Function

'=====================================================
Function DbSaveOk()
    
    Call InitVariables
    Call MainQuery()

End Function
