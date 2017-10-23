<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
call LoadBasisGlobalInf()
																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
	Dim lgOpModeCRUD
	On Error Resume Next
				'☜: Protect system from crashing
	Err.Clear 
				'☜: Clear Error status

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")
				'☜: Read Operation Mode (CRUD)


	Select Case lgOpModeCRUD
	        'Case CStr(UID_M0001)                                                         '☜: Query
	             'Call SubBizQuery()
	        Case CStr(UID_M0002)
	             Call SubBizSave()
	        'Case CStr(UID_M0003)                                                         '☜: Delete
	            ' Call SubBizDelete()
	        'Case "CHECK"                                                                 '☜: Check	
	            ' Call SubBizCheck()    
	End Select

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================

Sub SubBizSave()
    Dim I1_b_biz_partner
    Dim I2_b_biz_partner
    Dim I3_b_pur_grp
    Dim I4_m_lc_hdr
	Dim I5_b_bank
    Dim I6_b_bank
    Dim I7_b_bank
    Dim I8_b_bank
    Dim I9_b_bank
    Dim I10_s_wks_user
        
    Const M468_I12_lc_no = 0
    Const M468_I12_lc_doc_no = 1
    Const M468_I12_lc_amend_seq = 2
    Const M468_I12_po_no = 3
    Const M468_I12_adv_no = 4
    Const M468_I12_pre_adv_ref = 5
    Const M468_I12_req_dt = 6
    Const M468_I12_adv_dt = 7
    Const M468_I12_open_dt = 8
    Const M468_I12_expiry_dt = 9
    Const M468_I12_amend_dt = 10
    Const M468_I12_manufacturer = 11
    Const M468_I12_agent = 12
    Const M468_I12_currency = 13
    Const M468_I12_doc_amt = 14
    Const M468_I12_xch_rate = 15
    Const M468_I12_xch_rate_op = 16
    Const M468_I12_loc_amt = 17
    Const M468_I12_bank_txt = 18
    Const M468_I12_incoterms = 19
    Const M468_I12_pay_method = 20
    Const M468_I12_pay_terms_txt = 21
    Const M468_I12_partial_ship = 22
    Const M468_I12_latest_ship_dt = 23
    Const M468_I12_shipment = 24
    Const M468_I12_doc1 = 25
    Const M468_I12_doc2 = 26
    Const M468_I12_doc3 = 27
    Const M468_I12_doc4 = 28
    Const M468_I12_doc5 = 29
    Const M468_I12_file_dt = 30
    Const M468_I12_file_dt_txt = 31
    Const M468_I12_insrt_user_id = 32
    Const M468_I12_insrt_dt = 33
    Const M468_I12_updt_user_id = 34
    Const M468_I12_updt_dt = 35
    Const M468_I12_ext1_qty = 36
    Const M468_I12_ext1_amt = 37
    Const M468_I12_ext1_cd = 38
    Const M468_I12_remark = 39
    Const M468_I12_lc_kind = 40
    Const M468_I12_lc_type = 41
    Const M468_I12_transhipment = 42
    Const M468_I12_transfer = 43
    Const M468_I12_delivery_plce = 44
    Const M468_I12_tolerance = 45
    Const M468_I12_loading_port = 46
    Const M468_I12_dischge_port = 47
    Const M468_I12_transport_comp = 48
    Const M468_I12_origin = 49
    Const M468_I12_origin_cntry = 50
    Const M468_I12_charge_txt = 51
    Const M468_I12_charge_cd = 52
    Const M468_I12_credit_core = 53
    Const M468_I12_lc_remn_doc_amt = 54
    Const M468_I12_lc_remn_loc_amt = 55
    Const M468_I12_fund_type = 56
    Const M468_I12_lmt_xch_rate = 57
    Const M468_I12_lmt_amt = 58
    Const M468_I12_inv_cnt = 59
    Const M468_I12_bl_awb_flg = 60
    Const M468_I12_freight = 61
    Const M468_I12_notify_party = 62
    Const M468_I12_consignee = 63
    Const M468_I12_insur_policy = 64
    Const M468_I12_pack_list = 65
    Const M468_I12_cert_origin_flg = 66
    Const M468_I12_l_lc_type = 67
    Const M468_I12_o_lc_kind = 68
    Const M468_I12_o_lc_doc_no = 69
    Const M468_I12_o_lc_amend_seq = 70
    Const M468_I12_o_lc_no = 71
    Const M468_I12_o_lc_type = 72
    Const M468_I12_o_lc_open_dt = 73
    Const M468_I12_o_lc_expiry_dt = 74
    Const M468_I12_o_lc_loc_amt = 75
    Const M468_I12_transport = 76
    Const M468_I12_biz_area = 77
    Const M468_I12_pay_dur = 78
    Const M468_I12_charge_flg = 79
    Const M468_I12_ext2_qty = 80
    Const M468_I12_ext3_qty = 81
    Const M468_I12_ext2_amt = 82
    Const M468_I12_ext3_amt = 83
    Const M468_I12_ext2_cd = 84
    Const M468_I12_ext3_cd = 85
    Const M468_I12_ext1_rt = 86
    Const M468_I12_ext2_rt = 87
    Const M468_I12_ext3_rt = 88
    Const M468_I12_ext1_dt = 89
    Const M468_I12_ext2_dt = 90
    Const M468_I12_ext3_dt = 91
            
            
	Dim iPM4G111
	Dim strConvDt
	
	Redim I4_m_lc_hdr(M468_I12_ext3_dt)
		
	On Error Resume Next
	Err.Clear 		

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		I5_b_bank = UCase(Trim(Request("txtOpenBank")))
		I6_b_bank = UCase(Trim(Request("txtAdvBank")))
		I7_b_bank = UCase(Trim(Request("txtPayBank")))
		I8_b_bank = UCase(Trim(Request("txtRenegoBank")))
		I9_b_bank = UCase(Trim(Request("txtConfirmBank")))

		If Len(Trim(Request("txtAdvDt"))) Then
			strConvDt = UNIConvDate(Request("txtAdvDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtAdvDt", 1, I_MKSCRIPT)

				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_adv_dt) = strConvDt
			End If
		End If
		I4_m_lc_hdr(M468_I12_adv_no) = UCase(Trim(Request("txtAdvNo")))

		I4_m_lc_hdr(M468_I12_agent) = UCase(Trim(Request("txtAgent")))

		If Len(Trim(Request("txtAmendDt"))) Then
			strConvDt = UNIConvDate(Request("txtAmendDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtAmendDt", 1, I_MKSCRIPT)

                Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_amend_dt) = strConvDt
			End If
		End If		

		I1_b_biz_partner = UCase(Trim(Request("txtBeneficiary")))
		I2_b_biz_partner = UCase(Trim(Request("txtApplicant")))
		
		I4_m_lc_hdr(M468_I12_bank_txt) = Trim(Request("txtBankTxt"))
		I4_m_lc_hdr(M468_I12_currency) = UCase(Trim(Request("txtCurrency")))
		I4_m_lc_hdr(M468_I12_doc1) = Trim(Request("txtDoc1"))
		I4_m_lc_hdr(M468_I12_doc2) = Trim(Request("txtDoc2"))
		I4_m_lc_hdr(M468_I12_doc3) = Trim(Request("txtDoc3"))
		I4_m_lc_hdr(M468_I12_doc4) = Trim(Request("txtDoc4"))
		I4_m_lc_hdr(M468_I12_doc5) = Trim(Request("txtDoc5"))

		If Len(Trim(Request("txtDocAmt"))) Then
			I4_m_lc_hdr(M468_I12_doc_amt) = UNIConvNum(Request("txtDocAmt"),0)
		End If
	
		If Len(Trim(Request("txtExpiryDt"))) Then
			strConvDt = UNIConvDate(Request("txtExpiryDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtExpiryDt", 1, I_MKSCRIPT)

				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_expiry_dt) = strConvDt
			End If
		End If	

		I4_m_lc_hdr(M468_I12_file_dt) = UNIConvNum(Request("txtFileDt"),0)
		I4_m_lc_hdr(M468_I12_file_dt_txt) = Trim(Request("txtFileDtTxt"))
		I4_m_lc_hdr(M468_I12_incoterms) = UCase(Trim(Request("txtIncoterms")))
		I4_m_lc_hdr(M468_I12_insrt_user_id) = UCase(gUsrID)
		
		If Len(Trim(Request("txtLatestShipDt"))) Then
			strConvDt = UNIConvDate(Request("txtLatestShipDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtLatestShipDt", 1, I_MKSCRIPT)

				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_latest_ship_dt) = strConvDt
			End If
		End If	

		If Len(Trim(Request("txtLCAmendSeq"))) Then
			I4_m_lc_hdr(M468_I12_lc_amend_seq) = UNIConvNum(Request("txtLCAmendSeq"),0)
		End If

		I4_m_lc_hdr(M468_I12_lc_doc_no) = UCase(Trim(Request("txtLCDocNo")))
		I4_m_lc_hdr(M468_I12_lc_kind) = "M"
		I4_m_lc_hdr(M468_I12_lc_no) = UCase(Trim(Request("txtLCNo1")))

		If Len(Trim(Request("txtLocAmt"))) Then
			I4_m_lc_hdr(M468_I12_loc_amt) = UNIConvNum(Request("txtLocAmt"),0)
		End If

		I4_m_lc_hdr(M468_I12_manufacturer) = UCase(Trim(Request("txtManufacturer")))

		If Len(Trim(Request("txtOpenDt"))) Then
						    
			strConvDt = UNIConvDate(Request("txtOpenDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtOpenDt", 1, I_MKSCRIPT)

				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_open_dt) = strConvDt
			End If
		Else
            I4_m_lc_hdr(M468_I12_open_dt) = UNIConvDate(UniConvYYYYMMDDToDate(gDateFormat,"1900","01","01"))
		End If	
		

		If Len(Trim(Request("txtReqDt"))) Then
			strConvDt = UNIConvDate(Request("txtReqDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtReqDt", 1, I_MKSCRIPT)
				
				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_req_dt) = strConvDt
			End If
		End If	

		I4_m_lc_hdr(M468_I12_partial_ship) = Request("rdoPartailShip")
		I4_m_lc_hdr(M468_I12_pay_terms_txt) = Trim(Request("txtPaytermstxt"))
		I4_m_lc_hdr(M468_I12_pay_method) = UCase(Trim(Request("txtPayTerms")))
		I4_m_lc_hdr(M468_I12_pre_adv_ref) = Trim(Request("txtPreAdvRef"))
		I4_m_lc_hdr(M468_I12_remark) = Trim(Request("txtRemark"))
		
		I3_b_pur_grp = UCase(Trim(Request("txtPurGrp")))
		
		I4_m_lc_hdr(M468_I12_shipment) = Trim(Request("txtShipment"))
		if Trim(Request("hdnPoNoCnt")) = "1" then
			I4_m_lc_hdr(M468_I12_po_no) = UCase(Trim(Request("txtPONo")))
		End if	
		I4_m_lc_hdr(M468_I12_updt_user_id) = UCase(gUsrID)
		I10_s_wks_user = UCase(gUsrID)
		
		If Len(Trim(Request("txtXchRate"))) Then
			I4_m_lc_hdr(M468_I12_xch_rate) = UNIConvNum(Request("txtXchRate"),0)
		End If

		If Len(Trim(Request("txtLmtXchRate"))) Then
			I4_m_lc_hdr(M468_I12_lmt_xch_rate) = UNIConvNum(Request("txtLmtXchRate"),0)
		End If

		I4_m_lc_hdr(M468_I12_lc_type) = UCase(Trim(Request("txtLCType")))
		I4_m_lc_hdr(M468_I12_transfer) = Request("rdoTransfer")
		I4_m_lc_hdr(M468_I12_transhipment) = Request("rdoTranshipment")

		If Len(Trim(Request("txtPayDur"))) Then
			I4_m_lc_hdr(M468_I12_pay_dur) = UNIConvNum(Request("txtPayDur"),0)
		End If

		I4_m_lc_hdr(M468_I12_delivery_plce) = Trim(Request("txtDeliveryPlce"))
	
		If Len(Trim(Request("txttolerance"))) Then
			I4_m_lc_hdr(M468_I12_tolerance) = UNIConvNum(Request("txttolerance"),0)
		End If

		I4_m_lc_hdr(M468_I12_loading_port) = Trim(Request("txtLoadingPort"))
		I4_m_lc_hdr(M468_I12_dischge_port) = Trim(Request("txtDischgePort"))
		I4_m_lc_hdr(M468_I12_transport) = UCase(Trim(Request("txtTransport")))
		I4_m_lc_hdr(M468_I12_transport_comp) = Trim(Request("txtTransportComp"))
		I4_m_lc_hdr(M468_I12_origin) = Trim(Request("txtOrigin"))
		I4_m_lc_hdr(M468_I12_origin_cntry) = UCase(Trim(Request("txtOriginCntry")))
		If Request("rdoChargeCd") = "Y" then
			I4_m_lc_hdr(M468_I12_charge_cd) = "AP"
		else
			I4_m_lc_hdr(M468_I12_charge_cd) = "BE"
		End if
		I4_m_lc_hdr(M468_I12_charge_txt) = Trim(Request("txtChargeTxt"))
		I4_m_lc_hdr(M468_I12_credit_core) = Trim(Request("txtCreditCore"))

		If Len(Trim(Request("txtInvCnt"))) Then
			I4_m_lc_hdr(M468_I12_inv_cnt) = UNIConvNum(Request("txtInvCnt"),0)
		End If

		If Len(Trim(Request("txtLmtAmt"))) Then
			I4_m_lc_hdr(M468_I12_lmt_amt) = UNIConvNum(Request("txtLmtAmt"),0)
		End If
		if Request("rdoBLAwFlg") = "Y" then
			I4_m_lc_hdr(M468_I12_bl_awb_flg) = "BL"
		else
			I4_m_lc_hdr(M468_I12_bl_awb_flg) = "AWB"
		end if
		I4_m_lc_hdr(M468_I12_freight) = UCase(Trim(Request("txtFreight")))
		I4_m_lc_hdr(M468_I12_notify_party) = UCase(Trim(Request("txtNotifyParty")))
		I4_m_lc_hdr(M468_I12_consignee) = Trim(Request("txtConsignee"))
		I4_m_lc_hdr(M468_I12_insur_policy) = Trim(Request("txtInsurPolicy"))

		If Len(Trim(Request("txtPackList"))) Then
			I4_m_lc_hdr(M468_I12_pack_list) = UNIConvNum(Request("txtPackList"),0)
		End If

		If Not IsEmpty(Request("chkCertOriginFlg")) Then
			I4_m_lc_hdr(M468_I12_cert_origin_flg) = "Y"
		Else
			I4_m_lc_hdr(M468_I12_cert_origin_flg) = "N"
		End If

		I4_m_lc_hdr(M468_I12_xch_rate_op) = UCase(Trim(Request("hdnXchRtOp")))

		I4_m_lc_hdr(M468_I12_fund_type) = Request("txtFundType")

    lgOpModeCRUD = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
	Set iPM4G111 = Server.CreateObject("PM4G111.cMMaintLcHdrS")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

	Dim sTemp
	If lgOpModeCRUD = OPMD_CMODE Then

		sTemp = iPM4G111.M_MAINT_LC_HDR_SVR(gStrGlobalCollection,"CREATE",I1_b_biz_partner,I2_b_biz_partner, _
            I3_b_pur_grp,I4_m_lc_hdr,I5_b_bank,I6_b_bank,I7_b_bank,I8_b_bank,I9_b_bank,I10_s_wks_user)

    ElseIf lgOpModeCRUD = OPMD_UMODE Then

		Call iPM4G111.M_MAINT_LC_HDR_SVR(gStrGlobalCollection,"UPDATE",I1_b_biz_partner,I2_b_biz_partner, _
            I3_b_pur_grp,I4_m_lc_hdr,I5_b_bank,I6_b_bank,I7_b_bank,I8_b_bank,I9_b_bank,I10_s_wks_user)

    end if        
    
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM4G111 = Nothing
    
    
    

		'-----------------------
		'Result data display area
		'-----------------------
    Response.Write "<Script language=vbs> " & vbCr
    Response.Write "If """ & ConvSPChars(Trim(sTemp)) & """ <> """"  Then " & vbCr   
	Response.Write "parent.frm1.txtLCNo.value = """ & ConvSPChars(UCase(Trim(sTemp))) & """" & vbCr
	Response.Write "parent.frm1.txtLCNo1.value = """ & ConvSPChars(UCase(Trim(sTemp))) & """" & vbCr
	Response.Write "End If" & vbCr
    Response.Write "Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "            
    


'		Response.Write "If """& lgIntFlgMode& """  = """ & OPMD_CMODE& """  Then" & vbCr
'			Response.Write ".frm1.txtLCNo.value = """ & ConvSPChars(M32111.ExportMLcHdrLcNo)& """" & vbCr
'		Response.Write "End If" & vbCr

	
End Sub
%>