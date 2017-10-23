<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

	On Error Resume Next
	
Call SubBizLOOKUP()

'=========================================================================================
Sub SubBizLOOKUP()

	Dim iPM3G19P
	
	Dim I1_m_pur_ord_hdr
    Dim E1_m_iv_type
    Dim E2_b_daily_exchange_rate
    Dim E3_b_minor_delivery
    Dim E4_b_minor_incoterms_nm
    Dim E5_b_minor_paymeth_nm
    Dim E6_b_minor_transport_nm
    Dim E7_b_minor_loading_nm
    Dim E8_b_minor_dischge_nm
    Dim E9_b_minor_origin_nm
    Dim E10_b_biz_partner_agent_nm
    Dim E11_b_biz_partner_manufacturer_nm
    Dim E12_b_biz_partner_applicant_nm
    Dim E13_b_bank_pay_bank_nm
    Dim E14_m_pur_ord_hdr
    Dim E15_b_biz_partner
    Dim E16_b_pur_grp
    Dim E17_b_pur_org
    
    '  View Name : export_delivery b_minor
    
    Const M450_E2_multi_divide = 0
    Const M450_E2_std_rate = 1
    
    Const M450_E3_minor_cd = 0
    Const M450_E3_minor_nm = 1
    Const M450_E3_minor_type = 2


    Const M450_E14_po_no = 0
    Const M450_E14_sppl_cd = 1
    Const M450_E14_payee_cd = 2
    Const M450_E14_build_cd = 3
    Const M450_E14_po_dt = 4
    Const M450_E14_po_cur = 5
    Const M450_E14_xch_rt = 6
    Const M450_E14_pay_meth = 7
    Const M450_E14_pay_dur = 8
    Const M450_E14_vat_type = 9
    Const M450_E14_vat_rt = 10
    Const M450_E14_tot_vat_doc_amt = 11
    Const M450_E14_tot_vat_loc_amt = 12
    Const M450_E14_tot_po_doc_amt = 13
    Const M450_E14_tot_po_loc_amt = 14
    Const M450_E14_sppl_sales_prsn = 15
    Const M450_E14_sppl_tel_no = 16
    Const M450_E14_release_flg = 17
    Const M450_E14_pur_org = 18
    Const M450_E14_manufacturer = 19
    Const M450_E14_agent = 20
    Const M450_E14_applicant = 21
    Const M450_E14_offer_dt = 22
    Const M450_E14_expiry_dt = 23
    Const M450_E14_transport = 24
    Const M450_E14_incoterms = 25
    Const M450_E14_delivery_plce = 26
    Const M450_E14_packing_cond = 27
    Const M450_E14_inspect_means = 28
    Const M450_E14_dischge_city = 29
    Const M450_E14_dischge_port = 30
    Const M450_E14_loading_port = 31
    Const M450_E14_origin = 32
    Const M450_E14_invoice_no = 33
    Const M450_E14_fore_dvry_dt = 34
    Const M450_E14_shipment = 35
    Const M450_E14_remark = 36
    Const M450_E14_lc_flg = 37
    Const M450_E14_merg_pur_flg = 38
    Const M450_E14_pur_biz_area = 39
    Const M450_E14_pur_cost_cd = 40
    Const M450_E14_pay_terms_txt = 41
    Const M450_E14_pay_type = 42
    Const M450_E14_cls_flg = 43
    Const M450_E14_import_flg = 44
    Const M450_E14_bl_flg = 45
    Const M450_E14_cc_flg = 46
    Const M450_E14_rcpt_flg = 47
    Const M450_E14_subcontra_flg = 48
    Const M450_E14_ret_flg = 49
    Const M450_E14_iv_flg = 50
    Const M450_E14_rcpt_type = 51
    Const M450_E14_issue_type = 52
    Const M450_E14_iv_type = 53
    Const M450_E14_sending_bank = 54
    Const M450_E14_charge_flg = 55
    Const M450_E14_ext1_qty = 56
    Const M450_E14_ext1_amt = 57
    Const M450_E14_ext1_rt = 58
    Const M450_E14_ext2_qty = 59
    Const M450_E14_ext2_amt = 60
    Const M450_E14_ext2_rt = 61
    Const M450_E14_ext3_cd = 62
    Const M450_E14_ext3_qty = 63
    Const M450_E14_ext3_amt = 64
    Const M450_E14_ext3_rt = 65
    Const M450_E14_tracking_no = 66
    Const M450_E14_so_no = 67
    Const M450_E14_inspect_method = 68
    Const M450_E14_insrt_user_id = 69
    Const M450_E14_insrt_dt = 70
    Const M450_E14_updt_user_id = 71
    Const M450_E14_updt_dt = 72
    Const M450_E14_ext1_cd = 73
    Const M450_E14_ext1_dt = 74
    Const M450_E14_ext2_cd = 75
    Const M450_E14_ext2_dt = 76
    Const M450_E14_ext3_dt = 77
    Const M450_E14_xch_rate_op = 78
    Const M450_E14_bp_cd = 79
    Const M450_E14_pur_grp = 80
    
    '  View Name : export b_biz_partner
    Const M450_E15_bp_cd = 0
    Const M450_E15_bp_type = 1
    Const M450_E15_bp_nm = 2

    '  View Name : export b_pur_grp
    Const M450_E16_pur_grp = 0
    Const M450_E16_pur_grp_nm = 1

    '  View Name : export b_pur_org
    Const M450_E17_pur_org = 0
    Const M450_E17_pur_org_nm = 1

    
    Dim lgCurrency
    
    On Error Resume Next
    Err.Clear																'☜: Protect system from crashing

	I1_m_pur_ord_hdr = Request("txtPONo")

	Set iPM3G19P = Server.CreateObject("PM3G19P.cMLookupPoHdrS")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM3G19P = Nothing		                                                 '☜: Unload Comproxy DLL
	   Exit Sub
    End If  

    Call iPM3G19P.M_LOOKUP_PO_HDR_SVR(gStrGlobalCollection,cstr(I1_m_pur_ord_hdr),E1_m_iv_type,E2_b_daily_exchange_rate, _
				E3_b_minor_delivery,E4_b_minor_incoterms_nm,E5_b_minor_paymeth_nm,E6_b_minor_transport_nm, _
				E7_b_minor_loading_nm,E8_b_minor_dischge_nm,E9_b_minor_origin_nm,E10_b_biz_partner_agent_nm, _
				E11_b_biz_partner_manufacturer_nm,E12_b_biz_partner_applicant_nm,E13_b_bank_pay_bank_nm, _
				E14_m_pur_ord_hdr,E15_b_biz_partner,E16_b_pur_grp,E17_b_pur_org)
				
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM3G19P = Nothing		                                                 '☜: Unload Comproxy DLL
	   Exit Sub
    End If  
    
    Set iPM3G19P = Nothing
    
	lgCurrency = ConvSPChars(E14_m_pur_ord_hdr(M450_E14_po_cur))
'-----------------------
'Result data display area
'-----------------------
	Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr

		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
		Response.Write ".txtCurrency.value = """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_po_cur)) & """" & vbCr
		Response.Write "parent.CurFormatNumericOCX" & vbCr
		'##########################
		Response.Write ".txtDocAmt.text				= """ & UNIConvNumDBToCompanyByCurrency(E14_m_pur_ord_hdr(M450_E14_tot_po_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
		Response.Write ".txtLocAmt.text				= """ & UNIConvNumDBToCompanyByCurrency(E14_m_pur_ord_hdr(M450_E14_tot_po_loc_amt),gCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
		Response.Write ".txtPONo.value				= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_po_no)) & """" & vbCr
		Response.Write ".txtCurrency.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_po_cur)) & """" & vbCr
		Response.Write ".txtPayTerms.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_pay_meth)) & """" & vbCr
		Response.Write ".txtPayTermsNm.value		= """ & ConvSPChars(E5_b_minor_paymeth_nm) & """" & vbCr	
		Response.Write ".txtPayDur.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_pay_dur)) & """" & vbCr
		Response.Write ".txtPurGrp.value			= """ & ConvSPChars(E16_b_pur_grp(M450_E16_pur_grp)) & """" & vbCr
		Response.Write ".txtPurGrpNm.value			= """ & ConvSPChars(E16_b_pur_grp(M450_E16_pur_grp_nm)) & """" & vbCr
		Response.Write ".txtBeneficiary.value		= """ & ConvSPChars(E15_b_biz_partner(M450_E15_bp_cd)) & """" & vbCr
		Response.Write ".txtBeneficiaryNm.value		= """ & ConvSPChars(E15_b_biz_partner(M450_E15_bp_nm)) & """" & vbCr
		Response.Write ".txtPurOrg.value			= """ & ConvSPChars(E17_b_pur_org(M450_E17_pur_org)) & """" & vbCr
		Response.Write ".txtPurOrgNm.value			= """ & ConvSPChars(E17_b_pur_org(M450_E17_pur_org_Nm)) & """" & vbCr
		Response.Write ".txtApplicant.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_applicant)) & """" & vbCr
		Response.Write ".txtApplicantNm.value		= """ & ConvSPChars(E12_b_biz_partner_applicant_nm) & """" & vbCr
		Response.Write ".txtAgent.value				= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_agent)) & """" & vbCr
		Response.Write ".txtAgentNm.value			= """ & ConvSPChars(E10_b_biz_partner_agent_nm) & """" & vbCr
		Response.Write ".txtManufacturer.value		= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_manufacturer)) & """" & vbCr
		Response.Write ".txtManufacturerNm.value	= """ & ConvSPChars(E11_b_biz_partner_manufacturer_nm) & """" & vbCr
		'Response.Write ".txtMultiDiv.value			= """ & ConvSPChars(E2_b_daily_exchange_rate(M450_E2_multi_divide)) & """" & vbCr
		'Response.Write ".txtXchRate.text			= """ & ConvSPChars(E2_b_daily_exchange_rate(M450_E2_std_rate)) & """" & vbCr
		Response.Write ".txtHLcFlg.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_lc_flg)) & """" & vbCr
		Response.Write ".hdnXchRtOp.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_xch_rate_op)) & """" & vbCr '13차 추가 
		
		'Call parent.setAmt
      	Response.Write "Call Parent.RefOk " & vbCr
	Response.Write "End With " & vbCr
 Response.Write "</Script>" & vbCr

	Set iPM3G19P = Nothing
End Sub
%>
