Option Explicit                               

Const BIZ_PGM_ID = "s4111mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "s4112ma1_KO441"

Const C_PopShipToParty	= 1		' 납품처 
Const C_PopInvMgr		= 2		' 재고담당자 
Const C_PopTransMeth	= 3		' 운송방법 

Const TAB1 = 1                  '☜: Tab의 위치 
Const TAB2 = 2

Const	C_REF2_DN_REQ_NO	=	0
Const	C_REF2_SHIP_TO_PARTY	=	1
Const	C_REF2_BP_NM	=	2
Const	C_REF2_SALES_GRP	=	3
Const	C_REF2_SALES_GRP_NM	=	4
Const	C_REF2_SALES_ORG	=	5
Const	C_REF2_SALES_ORG_NM	=	6
Const	C_REF2_MOV_TYPE	=	7
Const	C_REF2_MOV_TYPE_NM	=	8
Const	C_REF2_DLVY_DT	=	9
Const	C_REF2_PROMISE_DT	=	10
Const	C_REF2_ACTUAL_GI_DT	=	11
Const	C_REF2_COST_CD	=	12
Const	C_REF2_BIZ_AREA	=	13
Const	C_REF2_BIZ_AREA_NM	=	14
Const	C_REF2_TRANS_METH	=	15
Const	C_REF2_TRANS_METH_NM	=	16
Const	C_REF2_GOODS_MV_NO	=	17
Const	C_REF2_CI_FLAG	=	18
Const	C_REF2_POST_FLAG	=	19
Const	C_REF2_SO_TYPE	=	20
Const	C_REF2_SO_TYPE_NM	=	21
Const	C_REF2_SO_NO	=	22
Const	C_REF2_SHIP_TO_PLCE	=	23
Const	C_REF2_INSRT_USER_ID	=	24
Const	C_REF2_INSRT_DT	=	25
Const	C_REF2_UPDT_USER_ID	=	26
Const	C_REF2_UPDT_DT	=	27
Const	C_REF2_EXT1_QTY	=	28
Const	C_REF2_EXT2_QTY	=	29
Const	C_REF2_EXT3_QTY	=	30
Const	C_REF2_EXT1_AMT	=	31
Const	C_REF2_EXT2_AMT	=	32
Const	C_REF2_EXT3_AMT	=	33
Const	C_REF2_EXT1_CD	=	34
Const	C_REF2_EXT2_CD	=	35
Const	C_REF2_EXT3_CD	=	36
Const	C_REF2_TEMP_SO_NO	=	37
Const	C_REF2_VAT_FLAG	=	38
Const	C_REF2_AR_FLAG	=	39
Const	C_REF2_CUR	=	40
Const	C_REF2_XCHG_RATE	=	41
Const	C_REF2_XCHG_RATE_OP	=	42
Const	C_REF2_NET_AMT	=	43
Const	C_REF2_NET_AMT_LOC	=	44
Const	C_REF2_VAT_AMT	=	45
Const	C_REF2_VAT_AMT_LOC	=	46
Const	C_REF2_EXCEPT_DN_FLAG	=	47
Const	C_REF2_REMARK	=	48
Const	C_REF2_ARRIVAL_DT	=	49
Const	C_REF2_ARRIVAL_TIME	=	50
Const	C_REF2_STP_INFO_NO	=	51
Const	C_REF2_ZIP_CD	=	52
Const	C_REF2_ADDR1	=	53
Const	C_REF2_ADDR2	=	54
Const	C_REF2_ADDR3	=	55
Const	C_REF2_RECEIVER	=	56
Const	C_REF2_TEL_NO1	=	57
Const	C_REF2_TEL_NO2	=	58
Const	C_REF2_TRANS_INFO_NO	=	59
Const	C_REF2_TRANS_CO	=	60
Const	C_REF2_DRIVER	=	61
Const	C_REF2_VEHICLE_NO	=	62
Const	C_REF2_SENDER	=	63
Const	C_REF2_STO_FLAG	=	64
Const	C_REF2_CASH_DC_AMT	=	65
Const	C_REF2_TAX_DC_AMT	=	66
Const	C_REF2_TAX_BASE_AMT	=	67
Const	C_REF2_CASH_DC_AMT_LOC	=	68
Const	C_REF2_TAX_DC_AMT_LOC	=	69
Const	C_REF2_TAX_BASE_AMT_LOC	=	70
Const	C_REF2_SO_AUTO_FLAG	=	71
Const	C_REF2_PLANT_CD	=	72
Const	C_REF2_PLANT_NM	=	73
Const	C_REF2_INV_MGR	=	74
Const	C_REF2_INV_MGR_NM	=	75
Const	C_REF2_CONTRY_CD	=	76
Const	C_REF2_RET_ITEM_FLAG	=	77
Const	C_REF2_REL_BILL_FLAG	=	78
Const	C_REF2_EXPORT_FLAG	=	79


Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgBlnFlgChgValue1			'2002.12.18 납품처상세정보 변경 여부 
Dim lgBlnFlgChgValue2			'2002.12.18 운송정보 변경 여부 
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim IsOpenPop						' Popup
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i   '@@@CommonQueryRs 를 위한 변수 

'=========================================
Sub FormatField()
    With frm1
        ' 날짜 OCX Foramt 설정 
        Call FormatDATEField(.txtPlanned_gi_dt)
        Call FormatDATEField(.txtArriv_dt)
    End With
End Sub

'=========================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' 날짜 OCX
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
	frm1.txtDlvy_dt.value = EndDate				' Design상 필드 길이를 조정하기 위해 Text 필드를 사용함 
	
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

	iCalledAspName = AskPRAspName("S3111AA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S3111AA1", "x")
		gblnWinEvent = False
		exit Function
	end if

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If Err.Number <> 0 Then	Err.Clear 
	Else
		Call SetSORef(arrRet)
		Call SetToolbar("11101000000111")					'⊙: 버튼 툴바 제어 
	End If	
End Function

'========================================
Function OpenDnReqRef()
	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call DisplayMsgBox("205153", "X", "X", "X")
		Exit Function
	End IF

	iCalledAspName = AskPRAspName("S4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s4511ra1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	' Popup에서 Cancel한 경우 
	If UBOUND(arrRet, 1) = 0 Then	
		If Err.Number <> 0 Then	Err.Clear 
	Else
		Call SetDnReqRef(arrRet)
		Call SetToolbar("11101000000111")					'⊙: 버튼 툴바 제어 
	End if
End Function

'========================================
Function OpenConDnNo()
	Dim iCalledAspName
	Dim strRet
	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True

	iCalledAspName = AskPRAspName("S4111PA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111PA1", "x")
		gblnWinEvent = False
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

			'납품처 
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
				iArrHeader(2) = "국가"
				iArrHeader(3) = "국가명"

				.txtShip_to_party.focus

			'재고담당 
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
				
			'운송방법 
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
	
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

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
		MsgBox "납품처를 먼저 입력하세요", vbInformation, parent.gLogoName
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

	arrParam(0) = "운송회사"							
	arrParam(1) = "B_MAJOR A , B_MINOR B"						
	arrParam(2) = ""										
	arrParam(3) = ""									
	arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9031", "''", "S") & " "				
	arrParam(5) = "운송회사"							

	arrField(0) = "B.MINOR_CD"								
	arrField(1) = "B.MINOR_NM"								

	arrHeader(0) = "운송회사"							
	arrHeader(1) = "운송회사명"						

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

	arrParam(0) = "차량번호"							
	arrParam(1) = "B_MAJOR A , B_MINOR B"						
	arrParam(2) = ""			
	arrParam(3) = ""									
	arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9032", "''", "S") & " "				
	arrParam(5) = "차량번호"							

	arrField(0) = "B.MINOR_CD"								
	arrField(1) = "B.MINOR_NM"								

	arrHeader(0) = "차량관리번호"							
	arrHeader(1) = "차량번호"						

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
		.txtShip_to_party.value = prArrRet(0)			'납품처 
		.txtShip_to_partyNm.value = prArrRet(1)			'납품처명 
		.txtDlvy_dt.value = prArrRet(3)					'납기일	
		.txtMovType.value = prArrRet(4)					'출하형태 
		.txtMovTypeNm.value = prArrRet(5)				'출하형태명 
		.txtSo_no.value = prArrRet(6)					'수주번호 
		.txtSales_Grp.value = prArrRet(7)				'영업그룹 
		.txtSales_GrpNm.value = prArrRet(8)				'영업그룹명	
		.txtSo_Type.value = prArrRet(9)					'수주형태	
		.txtSo_TypeNm.value = prArrRet(10)				'수주형태명 
		.txtTrans_meth.value = prArrRet(11)				'운송방법 
		.txtTrans_meth_nm.value = prArrRet(12)			'운송방법명 
		.txtPlantCd.value = prArrRet(13)				'공장 
		.txtPlantNm.value = prArrRet(14)				'공장명 
		
		Call txtShip_to_party_Onchange()

		.txtDnNo.focus

		.txtHRefRoot.Value = "SO"

	End With	
	lgBlnFlgChgValue = True

End Function

'========================================
Function SetDnReqRef(ByRef prArrRet)

	With frm1


		.txtPlantCd.value = prArrRet(0, C_REF2_PLANT_CD)				'공장 
		.txtPlantNm.value = prArrRet(0, C_REF2_PLANT_NM)				'공장명 
'		.txtPlanned_gi_dt.Text = prArrRet(0, ) 
		.txtMovType.value = prArrRet(0, C_REF2_MOV_TYPE)					'출하형태 
		.txtMovTypeNm.value = prArrRet(0, C_REF2_MOV_TYPE_NM)				'출하형태명 
		.txtShip_to_party.value = prArrRet(0, C_REF2_SHIP_TO_PARTY)			'납품처 
		.txtShip_to_partyNm.value = prArrRet(0, C_REF2_BP_NM)			'납품처명 
		.txtSo_Type.value = prArrRet(0, C_REF2_SO_TYPE)					'수주형태	
		.txtSo_TypeNm.value = prArrRet(0, C_REF2_SO_TYPE_NM)				'수주형태명 
		.txtInvMgr.Value = prArrRet(0, C_REF2_INV_MGR) 
		.txtInvMgrNm.Value = prArrRet(0, C_REF2_INV_MGR_NM) 
		.txtSales_Grp.value = prArrRet(0, C_REF2_SALES_GRP)				'영업그룹 
		.txtSales_GrpNm.value = prArrRet(0, C_REF2_SALES_GRP_NM)				'영업그룹명	
		.txtTrans_meth.value = prArrRet(0, C_REF2_TRANS_METH)				'운송방법 
		.txtTrans_meth_nm.value = prArrRet(0, C_REF2_TRANS_METH_NM)			'운송방법명 
		.txtActGi_dt.Value = prArrRet(0, C_REF2_ACTUAL_GI_DT) 
		.txtSo_no.value = prArrRet(0, C_REF2_TEMP_SO_NO)					'수주번호 
		.txtTempSoNo.value = prArrRet(0, C_REF2_TEMP_SO_NO)					'수주번호 
		.chkSoNo.checked  = True
'		.chkSoNo.disabled = True

		.txtGoods_mv_no.Value = prArrRet(0, C_REF2_GOODS_MV_NO) 
		.txtArriv_dt.Value = prArrRet(0, C_REF2_ARRIVAL_DT) 
		.txtDlvy_dt.value = prArrRet(0, C_REF2_DLVY_DT)					'납기일	
		.txtArriv_Tm.Value = prArrRet(0, C_REF2_ARRIVAL_TIME) 
		.txtRemark.Value = prArrRet(0, C_REF2_REMARK) 


		.txtSTP_Inf_No.Value = prArrRet(0, C_REF2_STP_INFO_NO) 
'		.txtZIP_cd.Value = prArrRet(0, C_REF2_ZIP_CD) 
		.txtReceiver.Value = prArrRet(0, C_REF2_RECEIVER) 
'		.txtADDR1_Dlv.Value = prArrRet(0, C_REF2_ADDR1) 
'		.txtADDR2_Dlv.Value = prArrRet(0, C_REF2_ADDR2) 
'		.txtADDR3_Dlv.Value = prArrRet(0, C_REF2_ADDR3) 
		.txtShip_to_place.Value = prArrRet(0, C_REF2_SHIP_TO_PLCE) 
'		.txtTel_No1.Value = prArrRet(0, C_REF2_TEL_NO1) 
'		.txtTel_No2.Value = prArrRet(0, C_REF2_TEL_NO2) 
		.txtTrnsp_Inf_No.Value = prArrRet(0, C_REF2_TRANS_INFO_NO) 
		.txtTransCo.Value = prArrRet(0, C_REF2_TRANS_CO) 
		.txtSender.Value = prArrRet(0, C_REF2_SENDER) 
		.txtVehicleNo.Value = prArrRet(0, C_REF2_VEHICLE_NO) 
		.txtDriver.Value = prArrRet(0, C_REF2_DRIVER) 

		Call txtShip_to_party_Onchange()

		.txtDnNo.focus

		.txtHRefRoot.Value = "DR"

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
	
	frm1.txtSTP_Inf_No.value = arrRet(0)			'납품처상세정보번호	
	frm1.txtZIP_cd.value = arrRet(1)				'우편번호 
	frm1.txtADDR1_Dlv.value = arrRet(2)				'납품주소1
	frm1.txtADDR2_Dlv.value = arrRet(3)				'납품주소2	
	frm1.txtADDR3_Dlv.value = arrRet(4)				'납품주소3
	frm1.txtReceiver.value = arrRet(5)				'인수자명 
	frm1.txtTel_No1.value = arrRet(6)				'전화번호1
	frm1.txtTel_No2.value = arrRet(7)				'전화번호2
	lgBlnFlgChgValue1 = True

End Function

'========================================
Function SetTrnsMethRef(Byval arrRet)	 
	
	frm1.txtTrnsp_Inf_No.value = arrRet(0)			'운송정보번호 
	frm1.txtTransCo.value = arrRet(1)				'운송회사 
	frm1.txtDriver.value = arrRet(2)				'운전자명 
	frm1.txtVehicleNo.value = arrRet(3)				'차량번호	
	frm1.txtSender.value = arrRet(4)				'인계자명 
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

' 출고처리된 경우 수정 불가토록 변경한다.
'========================================
Sub PostFlagProtect()
	With ggoOper
		Call .SetReqAttr(frm1.txtPlanned_gi_dt, "Q")		'출고예정일 
		Call .SetReqAttr(frm1.txtShip_to_place, "Q")		'납품장소 
		Call .SetReqAttr(frm1.txtTrans_meth, "Q")		'운송방법 
		Call .SetReqAttr(frm1.txtShip_to_party, "Q")		'납품처 
		Call .SetReqAttr(frm1.txtRemark, "Q")			'비고 
		Call .SetReqAttr(frm1.txtArriv_dt, "Q")			'실제납품일 
		Call .SetReqAttr(frm1.txtArriv_Tm, "Q")			'납품시간 
		Call .SetReqAttr(frm1.txtInvMgr, "Q")			'재고담당 
	End With
End Sub

'=====================================================
Sub PostFlagRelease()
	With ggoOper
		Call .SetReqAttr(frm1.txtPlanned_gi_dt, "N")		'출고예정일 
		Call .SetReqAttr(frm1.txtShip_to_party, "N")		'납품처 

		Call .SetReqAttr(frm1.txtShip_to_place, "D")		'납품장소 
		Call .SetReqAttr(frm1.txtTrans_meth, "D")		'운송방법 
		Call .SetReqAttr(frm1.txtRemark, "D")			'비고 
		Call .SetReqAttr(frm1.txtArriv_dt, "D")			'실제납품일 
		Call .SetReqAttr(frm1.txtArriv_Tm, "D")			'납품시간 
		Call .SetReqAttr(frm1.txtInvMgr, "D")			'재고담당 
	End With
End Sub

'=====================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877
	
	Dim strTemp, arrVal

	If Kubun = 1 Then
		If frm1.txtConDnNo.Value = "" Then
			frm1.txtHRefRoot.value = ""
		End If

		WriteCookie CookieSplit , frm1.txtConDnNo.value & parent.gRowSep & frm1.txtHRefRoot.value

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

	'************ 싱글인 경우 **************
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
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call FormatField()
    Call LockFieldInit("L")
    Call SetDefaultVal
    Call InitVariables                                                      
    Call SetToolbar("11100000000011")										'⊙: 버튼 툴바 제어 
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
        Call DisplayMsgBox("204256", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("S4111RA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111RA1", "x")
		gblnWinEvent = False
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
		Call SetToolbar("11101000000111")					'⊙: 버튼 툴바 제어 
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
		gblnWinEvent = False
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
		Call SetToolbar("11101000000111")					'⊙: 버튼 툴바 제어 
	End If	

End Function

'========================================
Sub txtPlanned_gi_dt_Change()
	lgBlnFlgChgValue = True
End Sub

'실제납품일 
Sub txtArriv_dt_Change()
	lgBlnFlgChgValue = True
End Sub

'납품처 상세정보, 운송정보 변경여부 
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

' 납품처 onchange시 국가코드 가져오도록 수정 
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
' 납품처 정보를 수주참조시 default로 최근의 납품처정보를 가져오도록 수정 
' 납품처 정보가 없는 경우는 거래처의 정보를 default로 가져옴 
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

' 납품처 상세정보 설정 
'========================================
Sub SetSTPInfo(ByRef prArrSTPInfo)
	With frm1
		.txtSTP_Inf_No.value = prArrSTPInfo(1)			'납품처상세정보번호	
		.txtZIP_cd.value = prArrSTPInfo(2)				'우편번호 
		.txtADDR1_Dlv.value = prArrSTPInfo(3)			'납품주소1
		.txtADDR2_Dlv.value = prArrSTPInfo(4)			'납품주소2	
		.txtADDR3_Dlv.value = prArrSTPInfo(5)			'납품주소3
		.txtReceiver.value = prArrSTPInfo(6)			'인수자명 
		.txtTel_No1.value = prArrSTPInfo(7)				'전화번호1
		.txtTel_No2.value = prArrSTPInfo(8)				'전화번호2
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										
    Call InitVariables															
    Call SetDefaultVal
        
    Call DbQuery																'☜: Query db data
    
    FncQuery = True																
        
End Function

'=====================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False                                                          
    
    If (lgBlnFlgChgValue = True Or lgBlnFlgChgValue1 = True or lgBlnFlgChgValue2 = True) Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
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
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    End If
    
	Dim IntRetCD
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")  '☜ 바뀐부분 
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
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    End If
    
	With frm1
        ' 입력필수 항목 Check
        If Not chkFieldByCell(.txtShip_to_party, "A", TAB1) Then Exit Function
        If Not chkFieldByCell(.txtPlanned_gi_dt, "A", TAB1) Then Exit Function
        
		' 필드길이 Check
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
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
	Dim sTmp

	sTmp = frm1.txtHRefRoot.Value
    Call InitVariables
    Call MainQuery()
	frm1.txtHRefRoot.Value = sTmp
End Function
