<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S7111MB2	    															*
'*  4. Program Name         : NEGO 등록																	*
'*  5. Program Desc         : NEGO에 관련된 Default Data Query Transaction 처리용 ASP					*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/05/09																*
'*  8. Modified date(Last)  : 2000/05/09																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/05/09 : 화면 design												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->
<%

	'---------------------------------- L/C Header Data Query (Bill_Info_Svr)----------------------------------
	Dim PS7G139_pvCommand 
	Dim PS7G139_I1_s_bill_hdr 
	Dim PS7G139_E1_a_gl 
	Dim PS7G139_E2_a_temp_gl 
	Dim PS7G139_E3_a_batch 
	Dim PS7G139_E4_b_biz_partner
	Dim PS7G139_E5_s_bill_type_config 
	Dim PS7G139_E6_b_biz_area 
	Dim PS7G139_E7_b_biz_partner 
	Dim PS7G139_E8_b_biz_partner 
	Dim PS7G139_E9_b_minor 
	Dim PS7G139_E10_b_minor 
	Dim PS7G139_E11_s_bill_hdr 
	Dim PS7G139_E12_b_biz_partner 
	Dim PS7G139_E13_b_sales_org 
	Dim PS7G139_E14_b_sales_grp 
	Dim PS7G139_E15_b_sales_org 
	Dim PS7G139_E16_b_sales_grp 
	Dim PS7G139_E17_b_biz_partner 
	Dim PS7G139_E18_s_bl_info 
	Dim PS7G139_E19_b_biz_partner 
	Dim PS7G139_E20_b_biz_partner 
	Dim PS7G139_E21_b_biz_partner 
	Dim PS7G139_E22_b_minor 
	Dim PS7G139_E23_b_minor 
	Dim PS7G139_E24_b_minor 
	Dim PS7G139_E25_b_minor 
	Dim PS7G139_E26_b_minor 
	Dim PS7G139_E27_b_minor 
	Dim PS7G139_E28_b_minor 
	Dim PS7G139_E29_b_minor 
	Dim PS7G139_E30_b_country 
	Dim PS7G139_E31_b_country 
	Dim PS7G139_E32_b_country 	

    'Const S528_I1_bill_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_bill_hdr
    
    Const S528_E1_gl_no = 0    '[CONVERSION INFORMATION]  View Name : exp a_gl
    Const S528_E2_temp_gl_no = 0    '[CONVERSION INFORMATION]  View Name : exp a_temp_gl
    Const S528_E3_batch_no = 0    '[CONVERSION INFORMATION]  View Name : exp a_batch
    Const S528_E4_bp_cd = 0    '[CONVERSION INFORMATION]  View Name : exp_sold_to_party b_biz_partner
    
    Const S528_E4_bp_nm = 1
    Const S528_E4_credit_rot_day = 2
    
    Const S528_E5_bill_type = 0    '[CONVERSION INFORMATION]  View Name : exp s_bill_type_config    
    Const S528_E5_bill_type_nm = 1
    
    Const S528_E6_biz_area_nm = 0    '[CONVERSION INFORMATION]  View Name : exp b_biz_area
    Const S528_E7_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_apllicant_nm b_biz_partner
    Const S528_E8_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_beneficiary_nm b_biz_partner
    Const S528_E9_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_vat_type_nm b_minor    
    Const S528_E10_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_pay_meth_nm b_minor
    
    Const S528_E11_bill_no = 0    '[CONVERSION INFORMATION]  View Name : exp s_bill_hdr    
    Const S528_E11_post_flag = 1
    Const S528_E11_trans_type = 2
    Const S528_E11_bill_dt = 3
    Const S528_E11_cur = 4
    Const S528_E11_xchg_rate = 5
    Const S528_E11_xchg_rate_op = 6
    Const S528_E11_bill_amt = 7
    Const S528_E11_vat_type = 8
    Const S528_E11_vat_rate = 9
    Const S528_E11_vat_amt = 10
    Const S528_E11_pay_meth = 11
    Const S528_E11_pay_dur = 12
    Const S528_E11_tax_bill_no = 13
    Const S528_E11_tax_prt_cnt = 14
    Const S528_E11_accept_fob_amt = 15
    Const S528_E11_beneficiary = 16
    Const S528_E11_applicant = 17
    Const S528_E11_remark = 18
    Const S528_E11_bill_amt_loc = 19
    Const S528_E11_vat_calc_type = 20
    Const S528_E11_vat_amt_loc = 21
    Const S528_E11_tax_biz_area = 22
    Const S528_E11_pay_type = 23
    Const S528_E11_pay_terms_txt = 24
    Const S528_E11_collect_amt = 25
    Const S528_E11_collect_amt_loc = 26
    Const S528_E11_income_plan_dt = 27
    Const S528_E11_nego_amt = 28
    Const S528_E11_so_no = 29
    Const S528_E11_lc_no = 30
    Const S528_E11_lc_doc_no = 31
    Const S528_E11_lc_amend_seq = 32
    Const S528_E11_bl_flag = 33
    Const S528_E11_biz_area = 34
    Const S528_E11_cost_cd = 35
    Const S528_E11_to_biz_area = 36
    Const S528_E11_to_cost_cd = 37
    Const S528_E11_ref_flag = 38
    Const S528_E11_sts = 39
    Const S528_E11_ext1_qty = 40
    Const S528_E11_ext2_qty = 41
    Const S528_E11_ext3_qty = 42
    Const S528_E11_ext1_amt = 43
    Const S528_E11_ext2_amt = 44
    Const S528_E11_ext3_amt = 45
    Const S528_E11_ext1_cd = 46
    Const S528_E11_ext2_cd = 47
    Const S528_E11_ext3_cd = 48
    Const S528_E11_vat_auto_flag = 49
    Const S528_E11_vat_inc_flag = 50

    Const S528_E12_bp_cd = 0    '[CONVERSION INFORMATION]  View Name : exp_payer b_biz_partner
    Const S528_E12_bp_nm = 1
    
    Const S528_E13_sales_org = 0    '[CONVERSION INFORMATION]  View Name : exp_income b_sales_org
    Const S528_E13_sales_org_nm = 1
    
    Const S528_E14_sales_grp = 0    '[CONVERSION INFORMATION]  View Name : exp_income b_sales_grp
    Const S528_E14_sales_grp_nm = 1
    
    Const S528_E15_sales_org = 0    '[CONVERSION INFORMATION]  View Name : exp_billing b_sales_org
    Const S528_E15_sales_org_nm = 1
    
    Const S528_E16_sales_grp = 0    '[CONVERSION INFORMATION]  View Name : exp_billing b_sales_grp
    Const S528_E16_sales_grp_nm = 1
    
    Const S528_E17_bp_cd = 0    '[CONVERSION INFORMATION]  View Name : exp_bill_to_party b_biz_partner
    Const S528_E17_bp_nm = 1
    
    Const S528_E18_bl_doc_no = 0    '[CONVERSION INFORMATION]  View Name : exp s_bl_info
    Const S528_E18_ship_no = 1
    Const S528_E18_manufacturer = 2
    Const S528_E18_agent = 3
    Const S528_E18_receipt_plce = 4
    Const S528_E18_vessel_nm = 5
    Const S528_E18_voyage_no = 6
    Const S528_E18_forwarder = 7
    Const S528_E18_vessel_cntry = 8
    Const S528_E18_loading_port = 9
    Const S528_E18_dischge_port = 10
    Const S528_E18_delivery_plce = 11
    Const S528_E18_loading_plan_dt = 12
    Const S528_E18_latest_ship_dt = 13
    Const S528_E18_dischge_plan_dt = 14
    Const S528_E18_transport = 15
    Const S528_E18_tranship_cntry = 16
    Const S528_E18_tranship_dt = 17
    Const S528_E18_final_dest = 18
    Const S528_E18_incoterms = 19
    Const S528_E18_packing_type = 20
    Const S528_E18_tot_packing_cnt = 21
    Const S528_E18_container_cnt = 22
    Const S528_E18_packing_txt = 23
    Const S528_E18_gross_weight = 24
    Const S528_E18_weight_unit = 25
    Const S528_E18_gross_volumn = 26
    Const S528_E18_volumn_unit = 27
    Const S528_E18_freight = 28
    Const S528_E18_freight_plce = 29
    Const S528_E18_trans_price = 30
    Const S528_E18_trans_currency = 31
    Const S528_E18_trans_doc_amt = 32
    Const S528_E18_trans_xch_rate = 33
    Const S528_E18_trans_loc_amt = 34
    Const S528_E18_bl_issue_cnt = 35
    Const S528_E18_bl_issue_plce = 36
    Const S528_E18_bl_issue_dt = 37
    Const S528_E18_origin = 38
    Const S528_E18_origin_cntry = 39
    Const S528_E18_loading_dt = 40
    Const S528_E18_ext1_qty = 41
    Const S528_E18_ext2_qty = 42
    Const S528_E18_ext3_qty = 43
    Const S528_E18_ext1_amt = 44
    Const S528_E18_ext2_amt = 45
    Const S528_E18_ext3_amt = 46
    Const S528_E18_ext1_cd = 47
    Const S528_E18_ext2_cd = 48
    Const S528_E18_ext3_cd = 49
    
    Const S528_E19_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_manufacturer_nm b_biz_partner
    Const S528_E20_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_agent_nm b_biz_partner
    Const S528_E21_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_forwarder_nm b_biz_partner
    Const S528_E22_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_loading_port_nm b_minor
    Const S528_E23_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_discharge_port_nm b_minor
    Const S528_E24_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_transport_nm b_minor
    Const S528_E25_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_incoterms_nm b_minor
    Const S528_E26_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_origin_nm b_minor
    Const S528_E27_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_packing_type_nm b_minor
    Const S528_E28_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_freight_nm b_minor
    Const S528_E29_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_pay_type_nm b_minor
    Const S528_E30_country_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_vessel_cntry_nm b_country
    Const S528_E31_country_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_origin_cntry_nm b_country
    Const S528_E32_country_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_tranship_cntry_nm b_country


	'---------------------------------- L/C Header Data Query ----------------------------------
	
	Dim PS4G119_Command 
	Dim PS4G119_I1_s_lc_hdr 
	Dim PS4G119_E1_b_biz_partner 
	Dim PS4G119_E2_b_bank 
	Dim PS4G119_E3_b_bank 
	Dim PS4G119_E4_b_bank 
	Dim PS4G119_E5_b_bank 
	Dim PS4G119_E6_b_bank
	Dim PS4G119_E7_b_sales_grp 
	Dim PS4G119_E8_b_sales_org 
	Dim PS4G119_E9_b_biz_partner 
	Dim PS4G119_E10_b_biz_partner 
	Dim PS4G119_E11_b_biz_partner 
	Dim PS4G119_E12_b_biz_partner
	Dim PS4G119_E13_b_biz_partner 
	Dim PS4G119_E14_b_minor 
	Dim PS4G119_E15_b_minor 
	Dim PS4G119_E16_b_minor 
	Dim PS4G119_E17_b_minor 
	Dim PS4G119_E18_b_minor 
	Dim PS4G119_E19_b_minor 
	Dim PS4G119_E20_b_minor 
	Dim PS4G119_E21_b_country 
	Dim PS4G119_E22_b_minor 
	Dim PS4G119_E23_b_minor 
	Dim PS4G119_E24_b_minor 
	Dim PS4G119_E25_b_minor 
	Dim PS4G119_E26_s_lc_hdr

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



	'---------------------------------- Bill Data Query ----------------------------------
	Dim PS7G119_pvCommandSend 
	Dim PS7G119_I1_s_bill_hdr 
	Const S508_I1_bill_no = 0    
	Const S508_I1_except_flag = 1

	Dim PS7G119_E1_a_gl 
	Dim PS7G119_E2_a_temp_gl 
	Dim PS7G119_E3_a_batch 
	Dim PS7G119_E4_s_bill_type_config 
	Dim PS7G119_E5_b_biz_partner 
	Dim PS7G119_E6_b_biz_area 
	Dim PS7G119_E7_b_minor 
	Dim PS7G119_E8_b_sales_org 
	Dim PS7G119_E9_b_biz_partner 
	Dim PS7G119_E10_b_biz_partner 
	Dim PS7G119_E11_b_biz_partner 
	Dim PS7G119_E12_b_minor 
	Dim PS7G119_E13_b_sales_grp 
	Dim PS7G119_E14_b_biz_partner 
	Dim PS7G119_E15_b_sales_org 
	Dim PS7G119_E16_b_sales_grp 
	Dim PS7G119_E17_s_bill_hdr 
	Dim PS7G119_E18_b_minor

	Const S508_E4_bill_type = 0
	Const S508_E4_bill_type_nm = 1

	Const S508_E5_bp_cd = 0    
	Const S508_E5_bp_nm = 1
	Const S508_E5_credit_rot_day = 2

	Const S508_E8_sales_org = 0   
	Const S508_E8_sales_org_nm = 1

	Const S508_E9_bp_cd = 0    
	Const S508_E9_bp_nm = 1

	Const S508_E13_sales_grp = 0    
	Const S508_E13_sales_grp_nm = 1

	Const S508_E14_bp_cd = 0    
	Const S508_E14_bp_nm = 1

	Const S508_E15_sales_org = 0    
	Const S508_E15_sales_org_nm = 1

	Const S508_E16_sales_grp = 0    
	Const S508_E16_sales_grp_nm = 1

	Const S508_E17_bill_no = 0    
	Const S508_E17_post_flag = 1
	Const S508_E17_trans_type = 2
	Const S508_E17_bill_dt = 3
	Const S508_E17_cur = 4
	Const S508_E17_xchg_rate = 5
	Const S508_E17_xchg_rate_op = 6
	Const S508_E17_bill_amt = 7
	Const S508_E17_vat_type = 8
	Const S508_E17_vat_rate = 9
	Const S508_E17_vat_amt = 10
	Const S508_E17_pay_meth = 11
	Const S508_E17_pay_dur = 12
	Const S508_E17_tax_bill_no = 13
	Const S508_E17_tax_prt_cnt = 14
	Const S508_E17_accept_fob_amt = 15
	Const S508_E17_beneficiary = 16
	Const S508_E17_applicant = 17
	Const S508_E17_remark = 18
	Const S508_E17_bill_amt_loc = 19
	Const S508_E17_vat_calc_type = 20
	Const S508_E17_vat_amt_loc = 21
	Const S508_E17_tax_biz_area = 22
	Const S508_E17_pay_type = 23
	Const S508_E17_pay_terms_txt = 24
	Const S508_E17_collect_amt = 25
	Const S508_E17_collect_amt_loc = 26
	Const S508_E17_income_plan_dt = 27
	Const S508_E17_nego_amt = 28
	Const S508_E17_so_no = 29
	Const S508_E17_lc_no = 30
	Const S508_E17_lc_doc_no = 31
	Const S508_E17_lc_amend_seq = 32
	Const S508_E17_bl_flag = 33
	Const S508_E17_biz_area = 34
	Const S508_E17_cost_cd = 35
	Const S508_E17_to_biz_area = 36
	Const S508_E17_to_cost_cd = 37
	Const S508_E17_ref_flag = 38
	Const S508_E17_sts = 39
	Const S508_E17_except_flag = 40
	Const S508_E17_reverse_flag = 41
	Const S508_E17_ext1_qty = 42
	Const S508_E17_ext2_qty = 43
	Const S508_E17_ext3_qty = 44
	Const S508_E17_ext1_amt = 45
	Const S508_E17_ext2_amt = 46
	Const S508_E17_ext3_amt = 47
	Const S508_E17_ext1_cd = 48
	Const S508_E17_ext2_cd = 49
	Const S508_E17_ext3_cd = 50
	Const S508_E17_vat_auto_flag = 51
	Const S508_E17_vat_inc_flag = 52
	Const S508_E17_deposit_amt = 53
	Const S508_E17_deposit_amt_loc = 54
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd                                                                 '☜: Hide Processing message

Dim PS7G139																' Master L/C Header 조회용 Object
Dim PS7G119 																' Minor Code 조회용 Object
Dim PS4G119
Dim PB0C004

Dim strExchRateOp
Dim DocCurrency
Dim ApplyDt

Dim NegoDocAmt

Err.Clear																'☜: Protect system from crashing

If UCase(Request("txtBLFlag")) = "Y" Then	

	'---------------------------------- L/C Header Data Query ----------------------------------
	
    Set PS7G139 = Server.CreateObject("PS7G139.cSLkInfoSvr")
	
	if CheckSYSTEMError(Err,True) = True Then 	
		Response.End 
	end if
	    
    PS7G139_I1_s_bill_hdr = Trim(Request("txtBillNo"))
    PS7G139_pvCommand = "LOOKUP"    
    
    call PS7G139.S_BL_INFO_SVR(gStrGlobalCollection, PS7G139_pvCommand, PS7G139_I1_s_bill_hdr, _
		PS7G139_E1_a_gl , PS7G139_E2_a_temp_gl , PS7G139_E3_a_batch , _
		PS7G139_E4_b_biz_partner, PS7G139_E5_s_bill_type_config , PS7G139_E6_b_biz_area , _
		PS7G139_E7_b_biz_partner, PS7G139_E8_b_biz_partner, PS7G139_E9_b_minor, _ 
		PS7G139_E10_b_minor , PS7G139_E11_s_bill_hdr , PS7G139_E12_b_biz_partner, _
		PS7G139_E13_b_sales_org , PS7G139_E14_b_sales_grp , PS7G139_E15_b_sales_org , _
		PS7G139_E16_b_sales_grp , PS7G139_E17_b_biz_partner , PS7G139_E18_s_bl_info , _
		PS7G139_E19_b_biz_partner , PS7G139_E20_b_biz_partner , PS7G139_E21_b_biz_partner , _
		PS7G139_E22_b_minor , PS7G139_E23_b_minor , PS7G139_E24_b_minor , _
		PS7G139_E25_b_minor , PS7G139_E26_b_minor , PS7G139_E27_b_minor , _
		PS7G139_E28_b_minor , PS7G139_E29_b_minor , PS7G139_E30_b_country , _
		PS7G139_E31_b_country , PS7G139_E32_b_country  )

	If CheckSYSTEMError(Err,True) = True Then 	
		Set PS7G139 = Nothing
		Response.End 
	end if   

	'---------------------------------- L/C Header Data Query ----------------------------------
	
	Redim PS4G119_I1_s_lc_hdr(1)
	
	PS4G119_I1_s_lc_hdr(S357_I1_lc_kind) = "M"
	'2005-01-07 UPDATE LSW
	'PS4G119_I1_s_lc_hdr(S357_I1_lc_no) = FilterVar(Trim(Request("txtLCNo")), "", "SNM")
	PS4G119_I1_s_lc_hdr(S357_I1_lc_no) = Trim(Request("txtLCNo"))
	PS4G119_Command  = "LOOKUP"


	Set PS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")

	if CheckSYSTEMError(Err,True) = True Then 	
		Set PS7G139 = Nothing
		Response.End 
	end if

	Call PS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection, PS4G119_Command,  PS4G119_I1_s_lc_hdr, PS4G119_E1_b_biz_partner , PS4G119_E2_b_bank,_
		PS4G119_E3_b_bank ,PS4G119_E4_b_bank ,PS4G119_E5_b_bank ,PS4G119_E6_b_bank , _
		PS4G119_E7_b_sales_grp ,PS4G119_E8_b_sales_org ,PS4G119_E9_b_biz_partner ,PS4G119_E10_b_biz_partner , _
		PS4G119_E11_b_biz_partner ,PS4G119_E12_b_biz_partner ,PS4G119_E13_b_biz_partner ,PS4G119_E14_b_minor , _
		PS4G119_E15_b_minor ,PS4G119_E16_b_minor ,PS4G119_E17_b_minor ,PS4G119_E18_b_minor , _
		PS4G119_E19_b_minor ,PS4G119_E20_b_minor ,PS4G119_E21_b_country ,PS4G119_E22_b_minor ,PS4G119_E23_b_minor , _
		PS4G119_E24_b_minor ,PS4G119_E25_b_minor ,PS4G119_E26_s_lc_hdr)		 
	
	
    If cStr(Err.Description) = "B_MESSAGE 203400" then     		
		
		Set PS4G119 = Nothing
		
		PS4G119_I1_s_lc_hdr(S357_I1_lc_kind) = "L"
		'2005-01-07 UPDATE LSW		
		'PS4G119_I1_s_lc_hdr(S357_I1_lc_no) = FilterVar(Trim(Request("txtLCNo")), "", "SNM")
		PS4G119_I1_s_lc_hdr(S357_I1_lc_no) = Trim(Request("txtLCNo"))
		PS4G119_Command  = "LOOKUP"		
		Set PS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")			
		
		If CheckSYSTEMError(Err,True) = True Then 		
			Response.End 
		end if
	
		Call PS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection, PS4G119_Command, PS4G119_I1_s_lc_hdr, PS4G119_E1_b_biz_partner , PS4G119_E2_b_bank,_
			PS4G119_E3_b_bank ,PS4G119_E4_b_bank ,PS4G119_E5_b_bank ,PS4G119_E6_b_bank , _
			PS4G119_E7_b_sales_grp ,PS4G119_E8_b_sales_org ,PS4G119_E9_b_biz_partner ,PS4G119_E10_b_biz_partner , _
			PS4G119_E11_b_biz_partner ,PS4G119_E12_b_biz_partner ,PS4G119_E13_b_biz_partner ,PS4G119_E14_b_minor , _
			PS4G119_E15_b_minor ,PS4G119_E16_b_minor ,PS4G119_E17_b_minor ,PS4G119_E18_b_minor , _
			PS4G119_E19_b_minor ,PS4G119_E20_b_minor ,PS4G119_E21_b_country ,PS4G119_E22_b_minor ,PS4G119_E23_b_minor , _
			PS4G119_E24_b_minor ,PS4G119_E25_b_minor ,PS4G119_E26_s_lc_hdr)		
		
		if CheckSYSTEMError(Err,True) = True Then 			
			Response.End 
		end if
		
	End If	
	
	Dim lgCurrency
	lgCurrency = ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_cur))		
	
%>

<Script Language=VBScript>
With parent.frm1
	Dim strDt
		
	.txtCurrency.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_cur))%>"
	
	Call parent.CurFormatNumericOCX
		
	.txtBillNo.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_bill_no))%>"
	.txtBLNo.value = "<%=ConvSPChars(PS7G139_E18_s_bl_info(S528_E18_bl_doc_no))%>"
	.txtLCNo.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_lc_no))%>"
	.txtLCDocNo.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_lc_doc_no))%>"
	.txtLCAmendSeq.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_lc_amend_seq))%>"

	<%
	NegoDocAmt				= 0
	NegoDocAmt				= NegoDocAmt + UNICDbl(PS7G139_E11_s_bill_hdr(S528_E11_bill_amt), 0)
	NegoDocAmt				= NegoDocAmt + UNICDbl(PS7G139_E11_s_bill_hdr(S528_E11_vat_amt), 0 )
	NegoDocAmt				= NegoDocAmt - UNICDbl(PS7G139_E11_s_bill_hdr(S528_E11_collect_amt), 0 )
	NegoDocAmt				= NegoDocAmt - UNICDbl(PS7G139_E11_s_bill_hdr(S528_E11_nego_amt), 0 )
	%>
	.txtNegoDocAmt.text		= "<%=UNINumClientFormatByCurrency(NegoDocAmt, lgCurrency, ggAmtOfMoneyNo)%>"

	.txtBaseDocAmt.text = "<%=UNINumClientFormatByCurrency(PS7G139_E11_s_bill_hdr(S528_E11_bill_amt), lgCurrency, ggAmtOfMoneyNo)%>" 		
	.txtBaseCurrency.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_cur))%>"
	
	.txtLatestShipDt.text = "<%=UNIDateClientFormat(PS7G139_E18_s_bl_info(S528_E18_latest_ship_dt))%>"
	.txtOpenDt.text = "<%=UNIDateClientFormat(PS4G119_E26_s_lc_hdr(S357_E26_open_dt))%>"
	.txtExpireDt.text = "<%=UNIDateClientFormat(PS4G119_E26_s_lc_hdr(S357_E26_expiry_dt))%>"
	
	.txtOpenBank.value =  "<%=ConvSPChars(PS4G119_E2_b_bank(S357_E2_bank_cd))%>"
	.txtOpenBankNm.value =  "<%=ConvSPChars(PS4G119_E2_b_bank(S357_E2_bank_nm))%>"
	
	.txtIncoterms.value =  "<%=ConvSPChars(PS7G139_E18_s_bl_info(S528_E18_incoterms))%>"
	.txtPayTerms.value =  "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_pay_meth))%>"
	.txtPayTermsNm.value =  "<%=ConvSPChars(PS7G139_E10_b_minor(S528_E10_minor_nm))%>"
	.txtPayDur.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_pay_dur))%>"
	
	.txtAgent.value = "<%=ConvSPChars(PS7G139_E18_s_bl_info(S528_E18_agent))%>"
	.txtAgentNm.value = "<%=ConvSPChars(PS7G139_E20_b_biz_partner(S528_E20_bp_nm))%>"
	
	.txtManufacturer.value = "<%=ConvSPChars(PS7G139_E18_s_bl_info(S528_E18_manufacturer))%>"
	.txtManufacturerNm.value = "<%=ConvSPChars(PS7G139_E19_b_biz_partner(S528_E19_bp_nm))%>"
	


	.txtApplicant.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_applicant))%>"
	.txtApplicantNm.value = "<%=ConvSPChars(PS7G139_E7_b_biz_partner(S528_E7_bp_nm))%>"
	
	.txtBeneficiary.value = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_beneficiary))%>"
	.txtBeneficiaryNm.value = "<%=ConvSPChars(PS7G139_E8_b_biz_partner(S528_E8_bp_nm))%>"
	
	.txtSalesGroup.value = "<%=ConvSPChars(PS7G139_E14_b_sales_grp(S528_E14_sales_grp))%>"
	.txtSalesGroupNm.value = "<%=ConvSPChars(PS7G139_E14_b_sales_grp(S528_E14_sales_grp_nm))%>"
	.txtXchRate.text = "<%=ConvSPChars(PS7G139_E11_s_bill_hdr(S528_E11_xchg_rate))%>"
	
End With
</Script>
<%
	
	DocCurrency = PS7G139_E11_s_bill_hdr(S528_E11_cur)
	ApplyDt = PS7G139_E11_s_bill_hdr(S528_E11_bill_dt)
		
	Set PS7G139 = Nothing															'☜: ComProxy UnLoad
	Set PS4G119 = Nothing	
	
Else 

	Redim PS7G119_I1_s_bill_hdr(1)
	
	PS7G119_I1_s_bill_hdr(S508_I1_bill_no) = Trim(Request("txtBillNo"))
	PS7G119_I1_s_bill_hdr(S508_I1_except_flag) = "N"
	PS7G119_pvCommandSend = "QUERY"
		

	Set PS7G119  = Server.CreateObject("PS7G119.cSLkBillHdrSvr")
	if CheckSYSTEMError(Err,True) = True Then 		
		Response.End
	end if	
	
	call PS7G119.S_LOOKUP_BILL_HDR_SVR(gStrGlobalCollection, PS7G119_pvCommandSend ,PS7G119_I1_s_bill_hdr, _
		PS7G119_E1_a_gl ,PS7G119_E2_a_temp_gl ,PS7G119_E3_a_batch , _
		PS7G119_E4_s_bill_type_config ,PS7G119_E5_b_biz_partner ,PS7G119_E6_b_biz_area , _
		PS7G119_E7_b_minor ,PS7G119_E8_b_sales_org ,PS7G119_E9_b_biz_partner , _
		PS7G119_E10_b_biz_partner ,PS7G119_E11_b_biz_partner ,PS7G119_E12_b_minor , _
		PS7G119_E13_b_sales_grp ,PS7G119_E14_b_biz_partner ,PS7G119_E15_b_sales_org , _
		PS7G119_E16_b_sales_grp ,PS7G119_E17_s_bill_hdr ,PS7G119_E18_b_minor) 

	if CheckSYSTEMError(Err,True) = True Then 
		Set PS7G119  = Nothing				
		Response.End
	end if		
	'---------------------------------- L/C Header Data Query ----------------------------------
	
	Redim PS4G119_I1_s_lc_hdr(1)
	
	PS4G119_I1_s_lc_hdr(S357_I1_lc_kind) = "L"
	'2005-01-07 UPDATE LSW
	'PS4G119_I1_s_lc_hdr(S357_I1_lc_no) = FilterVar(Trim(Request("txtLCNo")), "", "SNM")
	PS4G119_I1_s_lc_hdr(S357_I1_lc_no) = Trim(Request("txtLCNo"))
	PS4G119_Command  = "LOOKUP"		

	Set PS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")		
	If CheckSYSTEMError(Err,True) = True Then 
		Set PS7G119  = Nothing			
		Response.End
	end if

	Call PS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection, PS4G119_Command,  PS4G119_I1_s_lc_hdr, PS4G119_E1_b_biz_partner , PS4G119_E2_b_bank,_
		PS4G119_E3_b_bank ,PS4G119_E4_b_bank ,PS4G119_E5_b_bank ,PS4G119_E6_b_bank , _
		PS4G119_E7_b_sales_grp ,PS4G119_E8_b_sales_org ,PS4G119_E9_b_biz_partner ,PS4G119_E10_b_biz_partner , _
		PS4G119_E11_b_biz_partner ,PS4G119_E12_b_biz_partner ,PS4G119_E13_b_biz_partner ,PS4G119_E14_b_minor , _
		PS4G119_E15_b_minor ,PS4G119_E16_b_minor ,PS4G119_E17_b_minor ,PS4G119_E18_b_minor , _
		PS4G119_E19_b_minor ,PS4G119_E20_b_minor ,PS4G119_E21_b_country ,PS4G119_E22_b_minor ,PS4G119_E23_b_minor , _
		PS4G119_E24_b_minor ,PS4G119_E25_b_minor ,PS4G119_E26_s_lc_hdr)		 
		
	If CheckSYSTEMError(Err,True) = True Then 
		Set PS4G119 = Nothing	
		Set PS7G119  = Nothing		
		Response.End
	end if
	
		'-----------------------
		'Result data display area
		'-----------------------
		
		
	lgCurrency = ConvSPChars(PS4G119_E26_s_lc_hdr(S508_E17_cur))
%>
<Script Language=VBScript>
	With parent.frm1
		Dim strDt

		.txtBillNo.value		= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_bill_no))%>"
		.txtLCNo.value			= "<%=ConvSPChars(PS4G119_E26_s_lc_hdr(S357_E26_lc_no))%>"
		.txtLCDocNo.value		= "<%=ConvSPChars(PS4G119_E26_s_lc_hdr(S357_E26_lc_doc_no))%>"
		.txtLCAmendSeq.value	= "<%=ConvSPChars(PS4G119_E26_s_lc_hdr(S357_E26_lc_amend_seq))%>"
		.txtCurrency.value		= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_cur))%>"				
		Call parent.CurFormatNumericOCX
		<%
		NegoDocAmt				= 0
		NegoDocAmt				= NegoDocAmt + UNICDbl(PS7G119_E17_s_bill_hdr(S508_E17_bill_amt), 0)
		NegoDocAmt				= NegoDocAmt + UNICDbl(PS7G119_E17_s_bill_hdr(S508_E17_vat_amt), 0 )
		NegoDocAmt				= NegoDocAmt - UNICDbl(PS7G119_E17_s_bill_hdr(S508_E17_collect_amt), 0 )
		NegoDocAmt				= NegoDocAmt - UNICDbl(PS7G119_E17_s_bill_hdr(S508_E17_nego_amt), 0 )
		%>
		.txtNegoDocAmt.text		= "<%=UNINumClientFormat(NegoDocAmt, ggAmtOfMoney.DecPoint, 0 )%>"
		.txtBaseCurrency.value	= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_cur))%>"
		.txtBaseDocAmt.text		= "<%=UNINumClientFormat(PS7G119_E17_s_bill_hdr(S508_E17_bill_amt), ggAmtOfMoney.DecPoint, 2 )%>"


		.txtLatestShipDt.text	= "<%=UNIDateClientFormat(PS4G119_E26_s_lc_hdr(S357_E26_latest_ship_dt))%>"
		.txtOpenDt.text			= "<%=UNIDateClientFormat(PS4G119_E26_s_lc_hdr(S357_E26_open_dt))%>"
		.txtExpireDt.text		= "<%=UNIDateClientFormat(PS4G119_E26_s_lc_hdr(S357_E26_expiry_dt))%>"
		.txtOpenBank.value		= "<%=ConvSPChars(PS4G119_E2_b_bank(S357_E2_bank_cd))%>"
		.txtOpenBankNm.value	= "<%=ConvSPChars(PS4G119_E2_b_bank(S357_E2_bank_nm))%>"
		
		.txtPayTerms.value		= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_pay_meth))%>"
		.txtPayTermsNm.value	= "<%=ConvSPChars(PS7G119_E7_b_minor)%>"
		
		.txtPayDur.value		= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_pay_dur))%>"
		.txtApplicant.value		= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_applicant))%>"
		.txtApplicantNm.value	= "<%=ConvSPChars(PS7G119_E11_b_biz_partner)%>"
		.txtBeneficiary.value	= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_beneficiary))%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(PS7G119_E10_b_biz_partner)%>"
		.txtSalesGroup.value	= "<%=ConvSPChars(PS7G119_E13_b_sales_grp(S508_E13_sales_grp))%>"
		.txtSalesGroupNm.value	= "<%=ConvSPChars(PS7G119_E13_b_sales_grp(S508_E13_sales_grp_nm))%>"
		.txtXchRate.text		= "<%=ConvSPChars(PS7G119_E17_s_bill_hdr(S508_E17_xchg_rate))%>"
		
	End With
</Script>
<%
		DocCurrency = PS7G119_E17_s_bill_hdr(S508_E17_cur)
		ApplyDt = PS7G119_E17_s_bill_hdr(S508_E17_bill_dt)
				
		Set PS7G119  = Nothing													'☜: ComProxy UnLoad
		Set PS4G119 = Nothing													'☜: ComProxy UnLoad
End If

'-----------------------------------------------------------------------------------------
'  환율연산자를 가져온다.
'-----------------------------------------------------------------------------------------



If DocCurrency <> gCurrency Then

	Dim I1_b_currency_currency
	Dim I2_b_currency_currency
	Dim I3_b_daily_exchange_rate_apprl_dt
	Dim E1_b_monthly_exchange_rate

	Const B276_E1_std_rate = 0
	Const B276_E1_multi_divide = 1

	I1_b_currency_currency = DocCurrency
	I3_b_daily_exchange_rate_apprl_dt = ApplyDt
	I2_b_currency_currency= gCurrency
	
	Set PB0C004 = Server.CreateObject("PB0C004.CB0C004")
	
	If CheckSYSTEMError(Err,True) = True Then 				
		Response.End
	end if	

	E1_b_monthly_exchange_rate = PB0C004.B_SELECT_EXCHANGE_RATE(gStrGlobalCollection, I1_b_currency_currency, I2_b_currency_currency, _ 
		I3_b_daily_exchange_rate_apprl_dt )

	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then 		
		Set PB0C004 = Nothing
		Response.End
	end if	
			
	strExchRateOp = E1_b_monthly_exchange_rate(B276_E1_multi_divide)
Else

	strExchRateOp = "*"
End If
%>
<Script Language=VBScript>
	With parent.frm1
		.txtExchRateOp.value = "<%=strExchRateOp%>"
	End With
	
	Call parent.BillQueryOk()	
</Script>
<%	
Set PB0C004 = Nothing
Response.End
%>

