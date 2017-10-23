<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>

<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리														*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4211mb4.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : C/C Header Insert를 위한 S/O Header Data Query Transaction 처리용 ASP		*
'*  7. Modified date(First) : 2000/04/17																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kim Hyungsuk																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/17 : Coding Start												*
'********************************************************************************************************


Response.Buffer = True													'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

%>

<!-- #Include file="../../inc/IncServer.asp" -->

<%
    Dim lgOpModeCRUD

    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case UCase(Request("txtBasData"))
        Case CStr("SO")	     														'☜: 현재 수주헤더관련조회를 요청받음 
             Call SubBizSoNoHdr
        Case CStr("LC")																'☜: 현재 LC헤더관련조회를 요청받음 
             Call SubBizLCNoHdr
    End Select
'============================================================================================================
Sub SubBizSoNoHdr()

	Dim iS3G102
	Dim iCommandSent
	Dim I1_s_so_hdr
	Dim E1_s_so_hdr
	
    Const E1_so_no = 0
    Const E1_so_dt = 1
    Const E1_req_dlvy_dt = 2
    Const E1_cfm_flag = 3
    Const E1_price_flag = 4
    Const E1_cur = 5
    Const E1_xchg_rate = 6
    Const E1_net_amt = 7
    Const E1_net_amt_loc = 8
    Const E1_cust_po_no = 9
    Const E1_cust_po_dt = 10
    Const E1_sales_cost_center = 11
    Const E1_deal_type = 12
    Const E1_pay_meth = 13
    Const E1_pay_dur = 14
    Const E1_trans_meth = 15
    Const E1_vat_inc_flag = 16
    Const E1_vat_type = 17
    Const E1_vat_rate = 18
    Const E1_vat_amt = 19
    Const E1_vat_amt_loc = 20
    Const E1_origin_cd = 21
    Const E1_valid_dt = 22
    Const E1_contract_dt = 23
    Const E1_ship_dt_txt = 24
    Const E1_pack_cond = 25
    Const E1_inspect_meth = 26
    Const E1_incoterms = 27
    Const E1_dischge_city = 28
    Const E1_dischge_port_cd = 29
    Const E1_loading_port_cd = 30
    Const E1_beneficiary = 31
    Const E1_manufacturer = 32
    Const E1_agent = 33
    Const E1_remark = 34
    Const E1_pre_doc_no = 35
    Const E1_lc_flag = 36
    Const E1_rel_dn_flag = 37
    Const E1_rel_bill_flag = 38
    Const E1_ret_item_flag = 39
    Const E1_sp_stk_flag = 40
    Const E1_ci_flag = 41
    Const E1_export_flag = 42
    Const E1_so_sts = 43
    Const E1_insrt_user_id = 44
    Const E1_insrt_dt = 45
    Const E1_updt_user_id = 46
    Const E1_updt_dt = 47
    Const E1_ext1_qty = 48
    Const E1_ext2_qty = 49
    Const E1_ext3_qty = 50
    Const E1_ext1_amt = 51
    Const E1_ext2_amt = 52
    Const E1_ext3_amt = 53
    Const E1_ext1_cd = 54
    Const E1_maint_no = 55
    Const E1_ext3_cd = 56
    Const E1_pay_type = 57
    Const E1_pay_terms_txt = 58
    Const E1_dn_parcel_flag = 59
    Const E1_to_biz_area = 60
    Const E1_to_biz_grp = 61
    Const E1_biz_area = 62
    Const E1_to_biz_org = 63
    Const E1_to_biz_cost_center = 64
    Const E1_ship_dt = 65
    Const E1_auto_dn_flag = 66
    Const E1_ext2_cd = 67
    Const E1_bank_cd = 68
    Const E1_sales_grp = 69
    Const E1_sales_grp_nm = 70
    Const E1_so_type = 71
    Const E1_so_type_nm = 72
    Const E1_bill_to_party = 73
    Const E1_bill_to_party_type = 74
    Const E1_bill_to_party_nm = 75
    Const E1_ship_to_party = 76
    Const E1_ship_to_party_type = 77
    Const E1_ship_to_party_nm = 78
    Const E1_sold_to_party = 79
    Const E1_sold_to_party_type = 80
    Const E1_sold_to_party_nm = 81
    Const E1_payer = 82
    Const E1_payer_type = 83
    Const E1_payer_nm = 84
    Const E1_sales_org = 85
    Const E1_sales_org_nm = 86
    Const E1_bank_nm = 87
    Const E1_deal_type_nm = 88
    Const E1_vat_type_nm = 89
    Const E1_pay_meth_nm = 90
    Const E1_incoterms_nm = 91
    Const E1_pack_cond_nm = 92
    Const E1_inspect_meth_nm = 93
    Const E1_trans_meth_nm = 94
    Const E1_vat_inc_flag_nm = 95
    Const E1_pay_type_nm = 96
    Const E1_loading_port_nm = 97
    Const E1_dischge_port_nm = 98
    Const E1_origin_nm = 99
    Const E1_manufacturer_nm = 100
    Const E1_agent_nm = 101
    Const E1_beneficiary_nm = 102
    Const E1_currency_desc = 103
    Const E1_biz_area_nm = 104
    Const E1_to_biz_grp_nm = 105
    
	On Error Resume Next
	Err.Clear                                                               
	
	iCommandSent = "LOOKUP"
	'Update 2005-01-07 LSW
    'I1_s_so_hdr = FilterVar(Trim(Request("txtSONo")), "" , "SNM")
    I1_s_so_hdr = Trim(Request("txtSONo"))

    Set iS3G102 = Server.CreateObject ("PS3G102.cLookupSoHdrSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G102 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
	End If
	
	Call iS3G102.S_LOOKUP_SO_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_so_hdr, E1_s_so_hdr)
											
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G102 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
	End If
	
	Set iS3G102 = Nothing
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr

	Response.Write ".txtSalesGroup.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp))			& """" & vbCr
	Response.Write ".txtSalesGroupNm.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp_nm))       & """" & vbCr
	Response.Write ".txtCurrency.Value			= """ & ConvSPChars(E1_s_so_hdr(E1_cur))				& """" & vbCr
	Response.Write ".txtXchRate.Value			= """ & UNINumClientFormat(E1_s_so_hdr(E1_xchg_rate), ggExchRate.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtIncoterms.Value			= """ & ConvSPChars(E1_s_so_hdr(E1_incoterms))			& """" & vbCr
	Response.Write ".txtPayTerms.Value			= """ & ConvSPChars(E1_s_so_hdr(E1_pay_meth))			& """" & vbCr
	Response.Write ".txtSalesOrg.Value			= """ & ConvSPChars(E1_s_so_hdr(E1_sales_org))			& """" & vbCr
	Response.Write ".txtSalesOrgNm.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_sales_org_nm))		& """" & vbCr
	Response.Write ".txtApplicant.Value			= """ & ConvSPChars(E1_s_so_hdr(E1_sold_to_party))      & """" & vbCr
	Response.Write ".txtApplicantNm.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_sold_to_party_nm))   & """" & vbCr
	Response.Write ".txtBeneficiary.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_beneficiary))		& """" & vbCr
	Response.Write ".txtBeneficiaryNm.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_beneficiary_nm))     & """" & vbCr
	Response.Write ".txtLoadingPort.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_loading_port_cd))    & """" & vbCr
	Response.Write ".txtDischgePort.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_dischge_port_cd))    & """" & vbCr	
	Response.Write ".txtOrigin.Value			= """ & ConvSPChars(E1_s_so_hdr(E1_origin_cd))			& """" & vbCr
	Response.Write ".txtPackingType.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_pack_cond))			& """" & vbCr
	Response.Write ".txtPackingTypeNm.Value		= """ & ConvSPChars(E1_s_so_hdr(E1_pack_cond_nm))       & """" & vbCr

	Response.Write "parent.ReferenceQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
End Sub
'============================================================================================================
Sub SubBizLCNoHdr()

    Dim iS4G119
	Dim iCommandSent
    Dim I1_s_lc_hdr
	Dim E1_b_biz_partner
	Dim E2_b_bank
	Dim E3_b_bank
	Dim E4_b_bank
	Dim E5_b_bank
	Dim E6_b_bank
	Dim E7_b_sales_grp
	Dim E8_b_sales_org
	Dim E9_b_biz_partner
	Dim E10_b_biz_partner
	Dim E11_b_biz_partner
	Dim E12_b_biz_partner
	Dim E13_b_biz_partner
	Dim E14_b_minor
	Dim E15_b_minor
	Dim E16_b_minor
	Dim E17_b_minor
	Dim E18_b_minor
	Dim E19_b_minor
	Dim E20_b_minor
	Dim E21_b_country
	Dim E22_b_minor
	Dim E23_b_minor
	Dim E24_b_minor
	Dim E25_b_minor
	Dim E26_s_lc_hdr
	
	Const S357_I1_lc_no = 0    
	Const S357_I1_lc_kind = 1
	
	Const S357_E1_bp_nm = 0
	Const S357_E2_bank_cd = 0
	Const S357_E2_bank_nm = 1
	Const S357_E3_bank_cd = 0
	Const S357_E3_bank_nm = 1
	Const S357_E4_bank_cd = 0
	Const S357_E4_bank_nm = 1
	Const S357_E5_bank_cd = 0
	Const S357_E5_bank_nm = 1
	Const S357_E6_bank_cd = 0
	Const S357_E6_bank_nm = 1
	Const S357_E7_sales_grp_nm = 0
	Const S357_E7_sales_grp = 1
	Const S357_E8_sales_org_nm = 0
	Const S357_E8_sales_org = 1
	Const S357_E9_bp_nm = 0    
	Const S357_E9_bp_cd = 1
	Const S357_E10_bp_nm = 0   
	Const S357_E10_bp_cd = 1
	Const S357_E11_bp_nm = 0   
	Const S357_E12_bp_nm = 0   
	Const S357_E13_bp_nm = 0   
	Const S357_E14_minor_nm = 0
	Const S357_E15_minor_nm = 0
	Const S357_E16_minor_nm = 0
	Const S357_E17_minor_nm = 0
	Const S357_E18_minor_nm = 0
	Const S357_E19_minor_nm = 0
	Const S357_E20_minor_nm = 0
	Const S357_E21_country_nm = 0
	Const S357_E22_minor_nm = 0  
	Const S357_E23_minor_nm = 0  
	Const S357_E24_minor_nm = 0  
	Const S357_E25_minor_nm = 0  
	
	Const S357_E26_lc_no = 0    
	Const S357_E26_lc_doc_no = 1
	Const S357_E26_lc_amend_seq = 2
	Const S357_E26_so_no = 3
	Const S357_E26_adv_no = 4
	Const S357_E26_pre_adv_ref = 5
	Const S357_E26_adv_dt = 6
	Const S357_E26_open_dt = 7
	Const S357_E26_expiry_dt = 8
	Const S357_E26_amend_dt = 9
	Const S357_E26_manufacturer = 10
	Const S357_E26_agent = 11
	Const S357_E26_cur = 12
	Const S357_E26_lc_amt = 13
	Const S357_E26_xch_rate = 14
	Const S357_E26_lc_loc_amt = 15
	Const S357_E26_bank_txt = 16
	Const S357_E26_incoterms = 17
	Const S357_E26_pay_meth = 18
	Const S357_E26_payment_txt = 19
	Const S357_E26_latest_ship_dt = 20
	Const S357_E26_shipment = 21
	Const S357_E26_doc1 = 22
	Const S357_E26_doc2 = 23
	Const S357_E26_doc3 = 24
	Const S357_E26_doc4 = 25
	Const S357_E26_doc5 = 26
	Const S357_E26_file_dt = 27
	Const S357_E26_file_dt_txt = 28
	Const S357_E26_remark = 29
	Const S357_E26_lc_kind = 30
	Const S357_E26_lc_type = 31
	Const S357_E26_delivery_plce = 32
	Const S357_E26_amt_tolerance = 33
	Const S357_E26_loading_port = 34
	Const S357_E26_dischge_port = 35
	Const S357_E26_transport = 36
	Const S357_E26_transport_comp = 37
	Const S357_E26_origin = 38
	Const S357_E26_origin_cntry = 39
	Const S357_E26_charge_txt = 40
	Const S357_E26_charge_cd = 41
	Const S357_E26_credit_core = 42
	Const S357_E26_inv_cnt = 43
	Const S357_E26_bl_awb_flg = 44
	Const S357_E26_freight = 45
	Const S357_E26_notify_party = 46
	Const S357_E26_consignee = 47
	Const S357_E26_insur_policy = 48
	Const S357_E26_pack_list = 49
	Const S357_E26_l_lc_type = 50
	Const S357_E26_open_bank_txt = 51
	Const S357_E26_o_lc_doc_no = 52
	Const S357_E26_o_lc_amend_seq = 53
	Const S357_E26_o_lc_no = 54
	Const S357_E26_o_lc_expiry_dt = 55
	Const S357_E26_o_lc_loc_amt = 56
	Const S357_E26_o_lc_type = 57
	Const S357_E26_pay_dur = 58
	Const S357_E26_partial_ship_flag = 59
	Const S357_E26_biz_area = 60
	Const S357_E26_trnshp_flag = 61
	Const S357_E26_transfer_flag = 62
	Const S357_E26_cert_origin_flag = 63
	Const S357_E26_o_lc_amd_seq = 64
	Const S357_E26_sts = 65
	Const S357_E26_nego_amt = 66
	Const S357_E26_ext1_qty = 67
	Const S357_E26_ext2_qty = 68
	Const S357_E26_ext3_qty = 69
	Const S357_E26_ext1_amt = 70
	Const S357_E26_ext2_amt = 71
	Const S357_E26_ext3_qmt = 72
	Const S357_E26_ext1_cd = 73
	Const S357_E26_ext2_cd = 74
	Const S357_E26_ext3_cd = 75
	Const S357_E26_xch_rate_op = 76    
    
	On Error Resume Next
	Err.Clear                                                               

	ReDim I1_s_lc_hdr(1)		
	iCommandSent = "LOOKUP"
	I1_s_lc_hdr(S357_I1_lc_no) =  FilterVar(Trim(Request("txtLCNo")), "" , "SNM")

    Set iS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Set iS4G119 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
	End If
	
	Call iS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection,iCommandSent,I1_s_lc_hdr, _
	E1_b_biz_partner,E2_b_bank,E3_b_bank,E4_b_bank,E5_b_bank,E6_b_bank, _
	E7_b_sales_grp,E8_b_sales_org, _
	E9_b_biz_partner,E10_b_biz_partner,E11_b_biz_partner,E12_b_biz_partner,E13_b_biz_partner, _
	E14_b_minor,E15_b_minor,E16_b_minor,E17_b_minor,E18_b_minor,E19_b_minor,E20_b_minor,E21_b_country, _
    E22_b_minor,E23_b_minor,E24_b_minor,E25_b_minor,E26_s_lc_hdr )
											
	If CheckSYSTEMError(Err, True) = True Then
		Set iS4G119 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
	End If
	
	Set iS4G119 = Nothing


	'-----------------------
	'Display result data
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr

	Response.Write ".txtSalesGroup.Value		= """ & ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp))		& """" & vbCr
	Response.Write ".txtSalesGroupNm.Value		= """ & ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp_nm))   & """" & vbCr
	Response.Write ".txtCurrency.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_cur))				& """" & vbCr
	Response.write ".txtXchgRate.Text			= """ & UNINumClientFormat(E26_s_lc_hdr(S357_E26_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write ".txtIncoterms.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_incoterms))		& """" & vbCr
	Response.Write ".txtPayTerms.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_pay_meth))		& """" & vbCr
	Response.Write ".txtSalesOrg.Value			= """ & ConvSPChars(E8_b_sales_org(S357_E8_sales_org))		& """" & vbCr
	Response.Write ".txtSalesOrgNm.Value		= """ & ConvSPChars(E8_b_sales_org(S357_E8_sales_org_nm))	& """" & vbCr
	Response.Write ".txtApplicant.Value			= """ & ConvSPChars(E10_b_biz_partner(S357_E10_bp_cd))		& """" & vbCr
	Response.Write ".txtApplicantNm.Value		= """ & ConvSPChars(E10_b_biz_partner(S357_E10_bp_nm))		& """" & vbCr
	Response.Write ".txtBeneficiary.Value		= """ & ConvSPChars(E9_b_biz_partner(S357_E9_bp_cd))		& """" & vbCr
	Response.Write ".txtBeneficiaryNm.Value		= """ & ConvSPChars(E9_b_biz_partner(S357_E9_bp_nm))		& """" & vbCr
	Response.Write ".txtLoadingPort.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_loading_port))	& """" & vbCr
	Response.Write ".txtDischgePort.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_dischge_port))	& """" & vbCr
	Response.Write ".txtOrigin.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_origin))			& """" & vbCr
	Response.Write ".txtShipFinDt.Text		    = """ & UNIDateClientFormat(E26_s_lc_hdr("S357_E26_latest_ship_dt"))   & """" & vbCr

	Response.Write "parent.ReferenceQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
End Sub
'============================================================================================================
Sub SubBizSave()
    
End Sub
'============================================================================================================
Sub SubBizDelete()

End Sub
'============================================================================================================
Sub SubBizQueryMulti()
    
End Sub    
'============================================================================================================
Sub SubBizSaveMulti()        
    
End Sub    
'============================================================================================================
Sub SetErrorStatus()
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub
'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub

%>

