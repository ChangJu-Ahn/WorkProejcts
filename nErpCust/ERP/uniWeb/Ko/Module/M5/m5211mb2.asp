<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
call LoadBasisGlobalInf()
 
 Dim lgOpModeCRUD
 On Error Resume Next
 Err.Clear 

 Call HideStatusWnd

 lgOpModeCRUD = Request("txtMode")

 Select Case lgOpModeCRUD
         Case CStr(UID_M0002)
              Call SubBizSave()
 End Select

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSave()

 On Error Resume Next
 Err.Clear                '☜: Protect system from crashing

    ' View Name : import m_bl_hdr
    Dim I3_m_bl_hdr
    Const M517_I3_bl_no = 0
    Const M517_I3_po_no = 1
    Const M517_I3_lc_no = 2
    Const M517_I3_lc_amend_seq = 3
    Const M517_I3_lc_doc_no = 4
    Const M517_I3_manufacturer = 5
    Const M517_I3_agent = 6
    Const M517_I3_bl_doc_no = 7
    Const M517_I3_receipt_plce = 8
    Const M517_I3_vessel_nm = 9
    Const M517_I3_voyage_no = 10
    Const M517_I3_forwarder = 11
    Const M517_I3_vessel_cntry = 12
    Const M517_I3_loading_port = 13
    Const M517_I3_dischge_port = 14
    Const M517_I3_delivery_plce = 15
    Const M517_I3_loading_dt = 16
    Const M517_I3_dischge_dt = 17
    Const M517_I3_transport = 18
    Const M517_I3_tranship_cntry = 19
    Const M517_I3_tranship_dt = 20
    Const M517_I3_final_dest = 21
    Const M517_I3_currency = 22
    Const M517_I3_doc_amt = 23
    Const M517_I3_xch_rate = 24
    Const M517_I3_loc_amt = 25
    Const M517_I3_incoterms = 26
    Const M517_I3_pay_method = 27
    Const M517_I3_pay_dur = 28
    Const M517_I3_packing_type = 29
    Const M517_I3_tot_packing_cnt = 30
    Const M517_I3_container_cnt = 31
    Const M517_I3_packing_txt = 32
    Const M517_I3_gross_weight = 33
    Const M517_I3_weight_unit = 34
    Const M517_I3_gross_volume = 35
    Const M517_I3_volume_unit = 36
    Const M517_I3_freight = 37
    Const M517_I3_freight_plce = 38
    Const M517_I3_bl_issue_cnt = 39
    Const M517_I3_bl_issue_plce = 40
    Const M517_I3_bl_issue_dt = 41
    Const M517_I3_origin = 42
    Const M517_I3_origin_cntry = 43
    Const M517_I3_posting_flg = 44
    Const M517_I3_ext1_qty = 45
    Const M517_I3_cash_doc_amt = 46
    Const M517_I3_ext1_cd = 47
    Const M517_I3_vat_type = 48
    Const M517_I3_vat_rate = 49
    Const M517_I3_vat_doc_amt = 50
    Const M517_I3_vat_loc_amt = 51
    Const M517_I3_lc_open_dt = 52
    Const M517_I3_lc_type = 53
    Const M517_I3_open_bank = 54
    Const M517_I3_pay_terms_txt = 55
    Const M517_I3_pay_type = 56
    Const M517_I3_net_weight = 57
    Const M517_I3_biz_area = 58
    Const M517_I3_tax_biz_area = 59
    Const M517_I3_cost_cd = 60
    Const M517_I3_pre_pay_no = 61
    Const M517_I3_pre_pay_doc_amt = 62
    Const M517_I3_pre_pay_loc_amt = 63
    Const M517_I3_loan_no = 64
    Const M517_I3_iv_type = 65
    Const M517_I3_trans_type = 66
    Const M517_I3_sppl_iv_no = 67
    Const M517_I3_sppl_iv_dt = 68
    Const M517_I3_setlmnt_dt = 69
    Const M517_I3_setlmnt_bank = 70
    Const M517_I3_fund_type = 71
    Const M517_I3_setlmnt_cur = 72
    Const M517_I3_setlmnt_xch_rt = 73
    Const M517_I3_setlmnt_doc_amt = 74
    Const M517_I3_setlmnt_loc_amt = 75
    Const M517_I3_usd_xch_rt = 76
    Const M517_I3_usd_xch_amt = 77
    Const M517_I3_xch_comm_doc_amt = 78
    Const M517_I3_trust_rcp_doc_amt = 79
    Const M517_I3_repay_doc_amt = 80
    Const M517_I3_repay_dt = 81
    Const M517_I3_accpt_no = 82
    Const M517_I3_lg_doc_no = 83
    Const M517_I3_lg_dt = 84
    Const M517_I3_lg_bank = 85
    Const M517_I3_lg_xch_rt = 86
    Const M517_I3_guar_doc_amt = 87
    Const M517_I3_guar_loc_amt = 88
    Const M517_I3_mst_bl_doc_no = 89
    Const M517_I3_charge_flg = 90
    Const M517_I3_loan_doc_amt = 91
    Const M517_I3_cash_loc_amt = 92
    Const M517_I3_loan_loc_amt = 93
    Const M517_I3_ref_iv_no = 94
    Const M517_I3_payee_cd = 95
    Const M517_I3_build_cd = 96
    Const M517_I3_bl_rcpt_dt = 97
    Const M517_I3_ext2_qty = 98
    Const M517_I3_ext3_qty = 99
    Const M517_I3_ext1_amt = 100
    Const M517_I3_ext2_amt = 101
    Const M517_I3_ext3_amt = 102
    Const M517_I3_ext2_cd = 103
    Const M517_I3_ext3_cd = 104
    Const M517_I3_ext1_rt = 105
    Const M517_I3_ext2_rt = 106
    Const M517_I3_ext3_rt = 107
    Const M517_I3_ext1_dt = 108
    Const M517_I3_ext2_dt = 109
    Const M517_I3_ext3_dt = 110
    Const M517_I3_xch_rate_op = 111
    Const M517_I3_remark = 112
 
 Dim iPM5G111
 Dim strConvDt
 Dim lgIntFlgMode
 Dim pvCommandSent
 Dim m_bl_hdr_bl_no
 Dim str_txtBeneficiary
 Dim str_txtApplicant
 Dim str_txtPurGrp
 
 Redim I3_m_bl_hdr(M517_I3_remark)
 
	I3_m_bl_hdr(M517_I3_bl_no)   = UCase(Trim(Request("txtBLNo1")))
	I3_m_bl_hdr(M517_I3_bl_doc_no)  = UCase(Trim(Request("txtBLDocNo")))

	If Trim(Request("hdnChkPoNo")) = "1" then
		I3_m_bl_hdr(M517_I3_po_no)  = UCase(Trim(Request("txtPONo")))
	End If
	If Trim(Request("hdnChkLcDocNo")) = "1" then
		I3_m_bl_hdr(M517_I3_lc_doc_no)  = UCase(Trim(Request("txtLCDocNo")))
		I3_m_bl_hdr(M517_I3_lc_no)   = UCase(Trim(Request("txtLCNo")))
	End If
	If Len(Trim(Request("txtLCAmendSeq"))) Then
		I3_m_bl_hdr(M517_I3_lc_amend_seq) = UNIConvNum(Request("txtLCAmendSeq"),0)
	End If

	If Len(Trim(Request("txtLoadingDt"))) Then
		strConvDt = UNIConvDate(Request("txtLoadingDt"))

		If strConvDt = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtLoadingDt", 1, I_MKSCRIPT)
			Exit Sub 
		Else
			I3_m_bl_hdr(M517_I3_loading_dt) = strConvDt
		End If
	End If  
	If Len(Trim(Request("txtDischgeDt"))) Then
		strConvDt = UNIConvDate(Request("txtDischgeDt"))

		If strConvDt = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtDischgeDt", 1, I_MKSCRIPT)
			Exit Sub
		Else
			I3_m_bl_hdr(M517_I3_dischge_dt) = strConvDt
		End If
	End If 
	If Len(Trim(Request("txtSetlmnt"))) Then
		strConvDt = UNIConvDate(Request("txtSetlmnt"))

		If strConvDt = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtSetlmnt", 1, I_MKSCRIPT)
			Exit Sub 
		Else
			I3_m_bl_hdr(M517_I3_setlmnt_dt) = strConvDt
		End If
	Else	'지불예정일을 입력하지 않은 경우 2999/12/31로 셋팅함(2003.09.22)
		I3_m_bl_hdr(M517_I3_setlmnt_dt) = "2999-12-31"
	End If 

	If Len(Trim(Request("txtBLIssueDt"))) Then
		strConvDt = UNIConvDate(Request("txtBLIssueDt"))

		If strConvDt = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtBLIssueDt", 1, I_MKSCRIPT)
			Exit Sub
		Else
			I3_m_bl_hdr(M517_I3_bl_issue_dt)= strConvDt
		End If
	End If  

  'I3_m_bl_hdr(M517_I3_vat_rate)  = UNIConvNum(Request("txtVatRate"),0)
  
  I3_m_bl_hdr(M517_I3_manufacturer) = UCase(Trim(Request("txtManufacturer")))
  I3_m_bl_hdr(M517_I3_agent)   = UCase(Trim(Request("txtAgent")))
  I3_m_bl_hdr(M517_I3_receipt_plce) = Trim(Request("txtReceiptPlce"))
  I3_m_bl_hdr(M517_I3_vessel_nm)  = Trim(Request("txtVesselNm"))
  I3_m_bl_hdr(M517_I3_voyage_no)  = UCase(Trim(Request("txtVoyageNo")))
  I3_m_bl_hdr(M517_I3_forwarder)  = UCase(Trim(Request("txtForwarder")))
  I3_m_bl_hdr(M517_I3_vessel_cntry) = UCase(Trim(Request("txtVesselCntry")))
  I3_m_bl_hdr(M517_I3_loading_port) = Trim(Request("txtLoadingPort"))
  I3_m_bl_hdr(M517_I3_dischge_port) = Trim(Request("txtDischgePort"))
  I3_m_bl_hdr(M517_I3_delivery_plce) = Trim(Request("txtDeliveryPlce"))
  I3_m_bl_hdr(M517_I3_transport)  = UCase(Trim(Request("txtTransport")))
  I3_m_bl_hdr(M517_I3_tranship_cntry) = UCase(Trim(Request("txtTranshipCntry")))
  If Len(Trim(Request("txtTranshipDt"))) Then
	strConvDt = UNIConvDate(Request("txtTranshipDt"))
    I3_m_bl_hdr(M517_I3_tranship_dt) = strConvDt
  End If  
  I3_m_bl_hdr(M517_I3_final_dest)  = Trim(Request("txtFinalDest"))
  I3_m_bl_hdr(M517_I3_currency)  = UCase(Trim(Request("txtCurrency")))
  I3_m_bl_hdr(M517_I3_doc_amt)  = UNIConvNum(Request("txtDocAmt"),0)
  If Len(Trim(Request("txtXchRate"))) Then
   I3_m_bl_hdr(M517_I3_xch_rate) = UNIConvNum(Request("txtXchRate"),0)
  End If     
  If Len(Trim(Request("txtLocAmt"))) Then
   I3_m_bl_hdr(M517_I3_loc_amt) = UNIConvNum(Request("txtLocAmt"),0)
  End If
  I3_m_bl_hdr(M517_I3_incoterms)  = UCase(Trim(Request("txtIncoterms")))
  I3_m_bl_hdr(M517_I3_pay_method)  = UCase(Trim(Request("txtPayMethod")))
  If Len(Trim(Request("txtPayDur"))) Then
   I3_m_bl_hdr(M517_I3_pay_dur) = UNIConvNum(Request("txtPayDur"),0)
  End If
  I3_m_bl_hdr(M517_I3_packing_type) = UCase(Trim(Request("txtPackingType")))
  If Len(Trim(Request("txtTotPackingCnt"))) Then
   I3_m_bl_hdr(M517_I3_tot_packing_cnt)= UNIConvNum(Request("txtTotPackingCnt"),0)
  End If
  I3_m_bl_hdr(M517_I3_packing_txt) = Trim(Request("txtPackingTxt"))
  If Len(Trim(Request("txtGrossWeight"))) Then
   I3_m_bl_hdr(M517_I3_gross_weight)= UNIConvNum(Request("txtGrossWeight"),0)
  End If
  if Len(Trim(Request("txtNetWeight"))) Then
   I3_m_bl_hdr(M517_I3_net_weight) = UNIConvNum(Request("txtNetWeight"),0)
  End If

  I3_m_bl_hdr(M517_I3_weight_unit) = UCase(Trim(Request("txtWeightUnit")))
  If Len(Trim(Request("txtContainerCnt"))) Then
   I3_m_bl_hdr(M517_I3_container_cnt) = UNIConvNum(Request("txtContainerCnt"),0)
  End If
  If Len(Trim(Request("txtGrossVolumn"))) Then
   I3_m_bl_hdr(M517_I3_gross_volume) = UNIConvNum(Request("txtGrossVolumn"),0)
  End If

  I3_m_bl_hdr(M517_I3_volume_unit) = UCase(Trim(Request("txtVolumnUnit")))
  I3_m_bl_hdr(M517_I3_freight)  = UCase(Trim(Request("txtFreight")))
  I3_m_bl_hdr(M517_I3_freight_plce) = Trim(Request("txtFreightPlce"))
  If Len(Trim(Request("txtBLIssueCnt"))) Then
   I3_m_bl_hdr(M517_I3_bl_issue_cnt)= UNIConvNum(Request("txtBLIssueCnt"),0)
  End If
  I3_m_bl_hdr(M517_I3_bl_issue_plce) = Trim(Request("txtBLIssuePlce"))
  I3_m_bl_hdr(M517_I3_origin)   = Trim(Request("txtOrigin"))
  I3_m_bl_hdr(M517_I3_origin_cntry) = UCase(Trim(Request("txtOriginCntry")))
  I3_m_bl_hdr(M517_I3_posting_flg) = UCase(Trim(Request("txtPost")))
  I3_m_bl_hdr(M517_I3_cash_doc_amt) = UNIConvNum(Request("txtCashAmt"),0)
  I3_m_bl_hdr(M517_I3_vat_type)  = UCase(Trim(Request("txtVatType")))
  I3_m_bl_hdr(M517_I3_pay_terms_txt) = Trim(Request("txtPayTermsTxt"))
  I3_m_bl_hdr(M517_I3_pay_type)  = UCase(Trim(Request("txtPayType")))
  I3_m_bl_hdr(M517_I3_tax_biz_area)  = UCase(Trim(Request("txtTaxBizArea")))

  I3_m_bl_hdr(M517_I3_pre_pay_no)  = UCase(Trim(Request("txtPrePayNo")))
  I3_m_bl_hdr(M517_I3_pre_pay_doc_amt)= UNIConvNum(Request("txtPrePayDocAmt"),0)
  I3_m_bl_hdr(M517_I3_loan_no)  = UCase(Trim(Request("txtLoanNo")))
  I3_m_bl_hdr(M517_I3_iv_type)  = UCase(Trim(Request("txtIvType")))
  I3_m_bl_hdr(M517_I3_sppl_iv_no)  = UCase(Trim(Request("txtIvNo")))

  I3_m_bl_hdr(M517_I3_loan_doc_amt) = UNIConvNum(Request("txtLoanAmt"),0)

  I3_m_bl_hdr(M517_I3_payee_cd)  = UCase(Trim(Request("txtPayeeCd")))
  I3_m_bl_hdr(M517_I3_build_cd)  = UCase(Trim(Request("txtBuildCd")))

  I3_m_bl_hdr(M517_I3_xch_rate_op) = UCase(Trim(Request("hdnDiv")))
  I3_m_bl_hdr(M517_I3_remark)  = Trim(Request("txtRemark")) '비고 
  lgIntFlgMode = CInt(Request("txtFlgMode"))        '☜: 저장시 Create/Update 판별 

  If lgIntFlgMode = OPMD_CMODE Then
	pvCommandSent = "CREATE"
  ElseIf lgIntFlgMode = OPMD_UMODE Then
	pvCommandSent = "UPDATE"
  End If

  Set iPM5G111 = Server.CreateObject("PM5G111.cMMaintImportBlHdrS")

  If CheckSYSTEMError(Err,True) = True Then
	Set iPM5G111 = Nothing
	Exit Sub
  End If
 
	str_txtBeneficiary = UCase(Trim(Request("txtBeneficiary")))
	str_txtApplicant   = UCase(Trim(Request("txtApplicant")))
	str_txtPurGrp      = UCase(Trim(Request("txtPurGrp")))

	 m_bl_hdr_bl_no = iPM5G111.M_MAINT_IMPORT_BL_HDR_SVR(gStrGlobalCollection, _
	                                        pvCommandSent, _
	                                       str_txtBeneficiary, _
	                                       str_txtApplicant, _
	                                       I3_m_bl_hdr, _
	                                       str_txtPurGrp)
  If CheckSYSTEMError(Err,True) = True Then
	Set iPM5G111 = Nothing
	Exit Sub
  End If

  Set iPM5G111 = Nothing

  Response.Write "<Script Language=VBScript>" & vbCr
  Response.Write " With parent"     & vbCr
  Response.Write "  If """ & lgIntFlgMode& """ = """ & OPMD_CMODE & """  Then" & vbCr
  Response.Write "   .frm1.txtBLNo.value = """ & ConvSPChars(m_bl_hdr_bl_no) & """" & vbCr
  Response.Write "  End If"    & vbCr
  Response.Write "  .DbSaveOk"    & vbCr
  Response.Write " End With"      & vbCr
  Response.Write "</Script>"      & vbCr

  Set iPM5G111 = Nothing

End Sub
%>
