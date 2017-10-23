<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution                                                      *
'*  2. Function Name        :                                                                           *
'*  3. Program ID           : S5211MB4                                                                  *
'*  4. Program Name         :                                                                           *
'*  5. Program Desc         : 수출 B/L등록																*
'*  6. Comproxy List        : PS6G219.cSLkExportCcHdrSvr,PS3G102.cLookupSoHdrSvr                        *
'*                            PB5CS41.cLookupBizPartnerSvr                                              *
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/11/15																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Ahn Tae Hee												                *
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : 화면 design												*
'*							  2. 2000/04/18 : Coding Start												*
'*                            3. 2002/11/15 : UI 표준적용												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim strMode
                                                                                     '☜: Clear Error status
    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    strMode      = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
    Dim iCommandSent
    															
	Dim iS6G219		
    Dim I1_s_cc_hdr_cc_no
    Dim EG1_s_cc_hdr

   	Dim iS3G102
	Dim I1_s_so_hdr
	Dim E1_s_so_hdr
    
    Dim iB5CS41
	Dim imp_biz_partner_cd
    Dim E1_b_biz_partner
    
    Const E1_cc_no = 0    '  View Name : exp s_cc_hdr
    Const E1_iv_no = 1
    Const E1_iv_dt = 2
    Const E1_so_no = 3
    Const E1_manufacturer = 4
    Const E1_agent = 5
    Const E1_return_appl = 6
    Const E1_return_office = 7
    Const E1_reporter = 8
    Const E1_ed_no = 9
    Const E1_ed_dt = 10
    Const E1_ed_type = 11
    Const E1_inspect_type = 12
    Const E1_ep_type = 13
    Const E1_export_type = 14
    Const E1_incoterms = 15
    Const E1_pay_meth = 16
    Const E1_pay_dur = 17
    Const E1_dischge_cntry = 18
    Const E1_loading_port = 19
    Const E1_loading_cntry = 20
    Const E1_dischge_port = 21
    Const E1_vessel_nm = 22
    Const E1_transport = 23
    Const E1_trans_form = 24
    Const E1_inspect_req_dt = 25
    Const E1_device_plce = 26
    Const E1_lc_doc_no = 27
    Const E1_lc_amend_seq = 28
    Const E1_lc_no = 29
    Const E1_lc_type = 30
    Const E1_lc_open_dt = 31
    Const E1_open_bank = 32
    Const E1_gross_weight = 33
    Const E1_weight_unit = 34
    Const E1_tot_packing_cnt = 35
    Const E1_packing_type = 36
    Const E1_ship_fin_dt = 37
    Const E1_cur = 38
    Const E1_doc_amt = 39
    Const E1_fob_doc_amt = 40
    Const E1_xch_rate = 41
    Const E1_loc_amt = 42
    Const E1_fob_loc_amt = 43
    Const E1_freight_loc_amt = 44
    Const E1_insure_loc_amt = 45
    Const E1_el_doc_no = 46
    Const E1_el_app_dt = 47
    Const E1_ep_no = 48
    Const E1_ep_dt = 49
    Const E1_insp_cert_no = 50
    Const E1_insp_cert_dt = 51
    Const E1_quar_cert_no = 52
    Const E1_quar_cert_dt = 53
    Const E1_recomnd_no = 54
    Const E1_recomnd_dt = 55
    Const E1_trans_method = 56
    Const E1_trans_rep_cd = 57
    Const E1_trans_from_dt = 58
    Const E1_trans_to_dt = 59
    Const E1_customs = 60
    Const E1_final_dest = 61
    Const E1_remark1 = 62
    Const E1_remark2 = 63
    Const E1_remark3 = 64
    Const E1_origin = 65
    Const E1_origin_cntry = 66
    Const E1_usd_xch_rate = 67
    Const E1_biz_area = 68
    Const E1_ref_flag = 69
    Const E1_sts = 70
    Const E1_net_weight = 71
    Const E1_ext1_qty = 72
    Const E1_ext2_qty = 73
    Const E1_ext3_qty = 74
    Const E1_ext1_amt = 75
    Const E1_ext2_amt = 76
    Const E1_ext3_amt = 77
    Const E1_ext1_cd = 78
    Const E1_ext2_cd = 79
    Const E1_ext3_cd = 80
    Const E1_xch_rate_op = 81
    Const E2_bp_cd = 82  
    Const E2_bp_nm = 83    
    Const E3_bp_cd = 84  
    Const E3_bp_nm = 85
    Const E4_sales_grp = 86
    Const E4_sales_grp_nm = 87    
    Const E5_sales_org = 88 
    Const E5_sales_org_nm = 89
    Const E6_bp_nm = 90     
    Const E7_bp_nm = 91     
    Const E8_bp_nm = 92     
    Const E9_bp_nm = 93     
    Const E10_bank_nm = 94  
    Const E11_minor_nm = 95 
    Const E12_minor_nm = 96 
    Const E13_minor_nm = 97 
    Const E14_minor_nm = 98 
    Const E15_minor_nm = 99 
    Const E16_minor_nm = 100
    Const E17_minor_nm = 101
    Const E18_minor_nm = 102
    Const E19_country_nm = 103
    Const E20_minor_nm = 104  
    Const E21_country_nm = 105
    Const E22_minor_nm = 106  
    Const E23_country_nm = 107
    Const E24_minor_nm = 108  
    Const E25_minor_nm = 109  
    Const E26_minor_nm = 110  
    Const E27_minor_nm = 111  
    Const E28_bp_nm = 112     
    Const E29_minor_nm = 113

    'iS3G102
    Const S308_E1_so_no = 0
    Const S308_E1_so_dt = 1
    Const S308_E1_req_dlvy_dt = 2
    Const S308_E1_cfm_flag = 3
    Const S308_E1_price_flag = 4
    Const S308_E1_cur = 5
    Const S308_E1_xchg_rate = 6
    Const S308_E1_net_amt = 7
    Const S308_E1_net_amt_loc = 8
    Const S308_E1_cust_po_no = 9
    Const S308_E1_cust_po_dt = 10
    Const S308_E1_sales_cost_center = 11
    Const S308_E1_deal_type = 12
    Const S308_E1_pay_meth = 13
    Const S308_E1_pay_dur = 14
    Const S308_E1_trans_meth = 15
    Const S308_E1_vat_inc_flag = 16
    Const S308_E1_vat_type = 17
    Const S308_E1_vat_rate = 18
    Const S308_E1_vat_amt = 19
    Const S308_E1_vat_amt_loc = 20
    Const S308_E1_origin_cd = 21
    Const S308_E1_valid_dt = 22
    Const S308_E1_contract_dt = 23
    Const S308_E1_ship_dt_txt = 24
    Const S308_E1_pack_cond = 25
    Const S308_E1_inspect_meth = 26
    Const S308_E1_incoterms = 27
    Const S308_E1_dischge_city = 28
    Const S308_E1_dischge_port_cd = 29
    Const S308_E1_loading_port_cd = 30
    Const S308_E1_beneficiary = 31
    Const S308_E1_manufacturer = 32
    Const S308_E1_agent = 33
    Const S308_E1_remark = 34
    Const S308_E1_pre_doc_no = 35
    Const S308_E1_lc_flag = 36
    Const S308_E1_rel_dn_flag = 37
    Const S308_E1_rel_bill_flag = 38
    Const S308_E1_ret_item_flag = 39
    Const S308_E1_sp_stk_flag = 40
    Const S308_E1_ci_flag = 41
    Const S308_E1_export_flag = 42
    Const S308_E1_so_sts = 43
    Const S308_E1_insrt_user_id = 44
    Const S308_E1_insrt_dt = 45
    Const S308_E1_updt_user_id = 46
    Const S308_E1_updt_dt = 47
    Const S308_E1_ext1_qty = 48
    Const S308_E1_ext2_qty = 49
    Const S308_E1_ext3_qty = 50
    Const S308_E1_ext1_amt = 51
    Const S308_E1_ext2_amt = 52
    Const S308_E1_ext3_amt = 53
    Const S308_E1_ext1_cd = 54
    Const S308_E1_maint_no = 55
    Const S308_E1_ext3_cd = 56
    Const S308_E1_pay_type = 57
    Const S308_E1_pay_terms_txt = 58
    Const S308_E1_dn_parcel_flag = 59
    Const S308_E1_to_biz_area = 60
    Const S308_E1_to_biz_grp = 61
    Const S308_E1_biz_area = 62
    Const S308_E1_to_biz_org = 63
    Const S308_E1_to_biz_cost_center = 64
    Const S308_E1_ship_dt = 65
    Const S308_E1_auto_dn_flag = 66
    Const S308_E1_ext2_cd = 67
    Const S308_E1_bank_cd = 68
    Const S308_E1_sales_grp = 69
    Const S308_E1_sales_grp_nm = 70
    Const S308_E1_so_type = 71
    Const S308_E1_so_type_nm = 72
    Const S308_E1_bill_to_party = 73
    Const S308_E1_bill_to_party_type = 74
    Const S308_E1_bill_to_party_nm = 75
    Const S308_E1_ship_to_party = 76
    Const S308_E1_ship_to_party_type = 77
    Const S308_E1_ship_to_party_nm = 78
    Const S308_E1_sold_to_party = 79
    Const S308_E1_sold_to_party_type = 80
    Const S308_E1_sold_to_party_nm = 81
    Const S308_E1_payer = 82
    Const S308_E1_payer_type = 83
    Const S308_E1_payer_nm = 84
    Const S308_E1_sales_org = 85
    Const S308_E1_sales_org_nm = 86
    Const S308_E1_bank_nm = 87
    Const S308_E1_deal_type_nm = 88
    Const S308_E1_vat_type_nm = 89
    Const S308_E1_pay_meth_nm = 90
    Const S308_E1_incoterms_nm = 91
    Const S308_E1_pack_cond_nm = 92
    Const S308_E1_inspect_meth_nm = 93
    Const S308_E1_trans_meth_nm = 94
    Const S308_E1_vat_inc_flag_nm = 95
    Const S308_E1_pay_type_nm = 96
    Const S308_E1_loading_port_nm = 97
    Const S308_E1_dischge_port_nm = 98
    Const S308_E1_origin_nm = 99
    Const S308_E1_manufacturer_nm = 100
    Const S308_E1_agent_nm = 101
    Const S308_E1_beneficiary_nm = 102
    Const S308_E1_currency_desc = 103
    Const S308_E1_biz_area_nm = 104
    Const S308_E1_to_biz_grp_nm = 105
    'iB5CS41
    Const S074_E1_credit_rot_day = 53

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear

    iCommandSent = "QUERY"
                															
	I1_s_cc_hdr_cc_no = Trim(Request("txtCCNo"))
	'---------------------------------- C/C Header Data Query ----------------------------------

	Set iS6G219 = Server.CreateObject("PS6G219.cSLkExportCcHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
    End If   
  		
	Call iS6G219.S_LOOKUP_EXPORT_CC_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_cc_hdr_cc_no, _
	                                             EG1_s_cc_hdr)
	                                           	
	If CheckSYSTEMError(Err,True) = True Then
        Set iS6G219 = Nothing		                                                 '☜: Unload Comproxy DLL
        Response.End
    End If   

    Set iS6G219 = Nothing   
 
    Response.Write "<Script Language=VBScript>   " & vbCr
	Response.Write "With parent.frm1             " & vbCr
	Dim strDt
	Response.Write ".txtLCDocNo.value  = """ & ConvSPChars(EG1_s_cc_hdr(E1_lc_doc_no)) & """" & vbCr  
	If ConvSPChars(EG1_s_cc_hdr(E1_lc_doc_no))  <> "" Then
		Response.Write ".txtLCAmendSeq.value  = """ & ConvSPChars(EG1_s_cc_hdr(E1_lc_amend_seq)) & """" & vbCr  
	End IF
	Response.Write ".txtHLCNo.value  = """ & ConvSPChars(EG1_s_cc_hdr(E1_lc_no)) & """" & vbCr  
	Response.Write ".txtWeightUnit.value  = """ & ConvSPChars(EG1_s_cc_hdr(E1_weight_unit)) & """" & vbCr  
		
	Response.Write ".txtLoadingDt.text  = """ & UNIDateClientFormat(EG1_s_cc_hdr(E1_ship_fin_dt)) & """" & vbCr  

	Response.Write ".txtCurrency.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_cur)) & """" & vbCr  
	Response.Write ".txtCurrency1.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_cur)) & """" & vbCr  
	Response.Write " parent.CurFormatNumericOCX " & vbCr 
	Response.Write ".txtGrossWeight.value	 = """ & UNINumClientFormat(EG1_s_cc_hdr(E1_gross_weight), ggQty.DecPoint, 0) & """" & vbCr  
	Response.Write ".txtVesselNm.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_vessel_nm)) & """" & vbCr  
	Response.Write ".txtTotPackingCnt.value  = """ & UNINumClientFormat(EG1_s_cc_hdr(E1_tot_packing_cnt), ggQty.DecPoint, 0) & """" & vbCr  
'	Response.Write ".txtXchRate.text		 = """ & UNINumClientFormat(EG1_s_cc_hdr(E1_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr  
	Response.Write ".txtApplicant.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_cd)) & """" & vbCr  
	Response.Write ".txtApplicantNm.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_nm)) & """" & vbCr  
	Response.Write ".txtBeneficiary.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E3_bp_cd)) & """" & vbCr  
	Response.Write ".txtBeneficiaryNm.value  = """ & ConvSPChars(EG1_s_cc_hdr(E3_bp_nm)) & """" & vbCr  
	Response.Write ".txtAgent.value			 = """ & ConvSPChars(EG1_s_cc_hdr(E1_agent)) & """" & vbCr  
	Response.Write ".txtAgentNm.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E7_bp_nm)) & """" & vbCr  
	Response.Write ".txtManufacturer.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E1_manufacturer)) & """" & vbCr  
	Response.Write ".txtManufacturerNm.value = """ & ConvSPChars(EG1_s_cc_hdr(E6_bp_nm)) & """" & vbCr  
	Response.Write ".txtLoadingPort.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E1_loading_port)) & """" & vbCr  
	Response.Write ".txtLoadingPortNm.value  = """ & ConvSPChars(EG1_s_cc_hdr(E18_minor_nm)) & """" & vbCr  
	Response.Write ".txtDischgePort.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E1_dischge_port)) & """" & vbCr  
	Response.Write ".txtDischgePortNm.value  = """ & ConvSPChars(EG1_s_cc_hdr(E20_minor_nm)) & """" & vbCr  
	Response.Write ".txtOrigin.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_origin)) & """" & vbCr  
	Response.Write ".txtOriginNm.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E22_minor_nm)) & """" & vbCr  
	Response.Write ".txtOriginCntry.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E1_origin_cntry)) & """" & vbCr  
	Response.Write ".txtOriginCntryNm.value  = """ & ConvSPChars(EG1_s_cc_hdr(E23_country_nm)) & """" & vbCr  
	Response.Write ".txtPayTerms.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_pay_meth)) & """" & vbCr  
	Response.Write ".txtPayTermsNm.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E17_minor_nm)) & """" & vbCr  
	Response.Write ".txtPayDur.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_pay_dur)) & """" & vbCr  
	Response.Write ".txtIncoTerms.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_incoterms)) & """" & vbCr  
	Response.Write ".txtIncoTermsNm.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E16_minor_nm)) & """" & vbCr  
	Response.Write ".txtSalesGroup.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E4_sales_grp)) & """" & vbCr  
	Response.Write ".txtSalesGroupNm.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E4_sales_grp_nm)) & """" & vbCr  
	Response.Write ".txtPackingType.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E1_packing_type)) & """" & vbCr  
	Response.Write ".txtPackingTypeNm.value  = """ & ConvSPChars(EG1_s_cc_hdr(E26_minor_nm)) & """" & vbCr  
		
	If EG1_s_cc_hdr(E1_ref_flag) = "M" Then
		Response.Write ".txtRefFlg.value = ""M"" " & vbCr  
		Response.Write ".txtBilltoParty.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_cd)) & """" & vbCr  
		Response.Write ".txtBilltoPartyNm.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_nm)) & """" & vbCr  
		Response.Write ".txtPayer.value				 = """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_cd)) & """" & vbCr  
		Response.Write ".txtPayerNm.value			 = """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_nm)) & """" & vbCr  
		Response.Write ".txtToSalesGroup.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E4_sales_grp)) & """" & vbCr  
		Response.Write ".txtToSalesGroupNm.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E4_sales_grp_nm)) & """" & vbCr  
		Response.Write ".txtTaxBizArea.value		 = """ & ConvSPChars(EG1_s_cc_hdr(E1_biz_area)) & """" & vbCr  
		Response.Write ".txtVatRate.value = ""0""   " & vbCr 
		Response.Write ".btnPosting.Disabled = True " & vbCr 
	Else
		Response.Write ".txtRefFlg.value = ""C""    " & vbCr 
	End If	
		
	Response.Write "End With  " & vbCr 
    Response.Write "</Script> " & vbCr 
	If Len(Request("txtSONo")) then
                             
	    iCommandSent = "QUERY"
	   
	    I1_s_so_hdr = Trim(Request("txtSONo"))
                
        Set iS3G102 = Server.CreateObject ("PS3G102.cLookupSoHdrSvr")
 
	    If CheckSYSTEMError(Err, True) = True Then
		    Response.End				
	    End If
       
	    Call iS3G102.S_LOOKUP_SO_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_so_hdr, E1_s_so_hdr)
		   									
	    If CheckSYSTEMError(Err, True) = True Then
		    Set iS3G102 = Nothing		                                                 '☜: Unload Comproxy DLL
		    Response.End				
	    End If
   
	    Set iS3G102 = Nothing
        imp_biz_partner_cd = E1_s_so_hdr(S308_E1_sold_to_party) 
        iCommandSent = "LOOKUP"
    
        Set iB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    

        If CheckSYSTEMError(Err,True) = True Then
            Response.End				
        End If     
    
	    Call iB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection, iCommandSent, imp_biz_partner_cd, E1_b_biz_partner)           									 
 								 									 
        If CheckSYSTEMError(Err,True) = True Then
           Set iB5CS41 = Nothing		                                                 '☜: Unload Comproxy DLL
           Response.End				
        End If      
   
        Set iB5CS41 = Nothing   
        Response.Write "<Script Language=vbscript>" & vbCr
        Response.Write "With parent.frm1          "	& vbCr 
		Response.Write ".txtSONo.value				= """ & ConvSPChars(E1_s_so_hdr(S308_E1_so_no)) & """" & vbCr
		Response.Write ".txtBillToParty.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_bill_to_party)) & """" & vbCr
		Response.Write ".txtBillToPartyNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_bill_to_party_nm)) & """" & vbCr
		Response.Write ".txtPayer.value				= """ & ConvSPChars(E1_s_so_hdr(S308_E1_payer)) & """" & vbCr
		Response.Write ".txtPayerNm.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_payer_nm)) & """" & vbCr
		Response.Write ".txtToSalesGroup.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_sales_grp)) & """" & vbCr
		Response.Write ".txtToSalesGroupNm.value	= """ & ConvSPChars(E1_s_so_hdr(S308_E1_sales_grp_nm)) & """" & vbCr
		
		Response.Write ".txtCreditRot.value = """ & E1_b_biz_partner(S074_E1_credit_rot_day) & """" & vbCr
		Response.Write "Call parent.ReferenceQueryOk()	 " & vbCr			
		Response.Write "End With  " & vbCr
		Response.Write "</Script> " & vbCr
	
    Else
    
    	Response.Write "<Script Language=vbscript>      " & vbCr
		Response.Write "Call parent.ReferenceQueryOk()	" & vbCr			
		Response.Write "</Script>                       " & vbCr
    
    End if
 
 %>