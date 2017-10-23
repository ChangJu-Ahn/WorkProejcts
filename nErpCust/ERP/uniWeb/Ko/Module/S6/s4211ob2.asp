<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4211ob2.asp																*
'*  4. Program Name         : 데이터가져오기(Packing List 출력 통관참조에서)							*
'*  5. Program Desc         : 데이터가져오기(Packing List 출력 통관참조에서)							*
'*  7. Modified date(First) : 2000/12/08																*
'*  8. Modified date(Last)  : 2000/12/08																*
'*  9. Modifier (First)     : Kim Hyungsuk 																*
'* 10. Modifier (Last)      : Kim Hyungsuk 																*
'* 11. Comment              :																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%																

On Error Resume Next

Call LoadBasisGlobalInf()
Call HideStatusWnd

    Dim pS6G219
    Dim iCommandSent
    
    Dim I1_s_cc_hdr_cc_no 
    Dim EG1_s_cc_hdr 

    Const S446_E1_cc_no = 0
    Const S446_E1_iv_no = 1
    Const S446_E1_iv_dt = 2
    Const S446_E1_so_no = 3
    Const S446_E1_manufacturer = 4
    Const S446_E1_agent = 5
    Const S446_E1_return_appl = 6
    Const S446_E1_return_office = 7
    Const S446_E1_reporter = 8
    Const S446_E1_ed_no = 9
    Const S446_E1_ed_dt = 10
    Const S446_E1_ed_type = 11
    Const S446_E1_inspect_type = 12
    Const S446_E1_ep_type = 13
    Const S446_E1_export_type = 14
    Const S446_E1_incoterms = 15
    Const S446_E1_pay_meth = 16
    Const S446_E1_pay_dur = 17
    Const S446_E1_dischge_cntry = 18
    Const S446_E1_loading_port = 19
    Const S446_E1_loading_cntry = 20
    Const S446_E1_dischge_port = 21
    Const S446_E1_vessel_nm = 22
    Const S446_E1_transport = 23
    Const S446_E1_trans_form = 24
    Const S446_E1_inspect_req_dt = 25
    Const S446_E1_device_plce = 26
    Const S446_E1_lc_doc_no = 27
    Const S446_E1_lc_amend_seq = 28
    Const S446_E1_lc_no = 29
    Const S446_E1_lc_type = 30
    Const S446_E1_lc_open_dt = 31
    Const S446_E1_open_bank = 32
    Const S446_E1_gross_weight = 33
    Const S446_E1_weight_unit = 34
    Const S446_E1_tot_packing_cnt = 35
    Const S446_E1_packing_type = 36
    Const S446_E1_ship_fin_dt = 37
    Const S446_E1_cur = 38
    Const S446_E1_doc_amt = 39
    Const S446_E1_fob_doc_amt = 40
    Const S446_E1_xch_rate = 41
    Const S446_E1_loc_amt = 42
    Const S446_E1_fob_loc_amt = 43
    Const S446_E1_freight_loc_amt = 44
    Const S446_E1_insure_loc_amt = 45
    Const S446_E1_el_doc_no = 46
    Const S446_E1_el_app_dt = 47
    Const S446_E1_ep_no = 48
    Const S446_E1_ep_dt = 49
    Const S446_E1_insp_cert_no = 50
    Const S446_E1_insp_cert_dt = 51
    Const S446_E1_quar_cert_no = 52
    Const S446_E1_quar_cert_dt = 53
    Const S446_E1_recomnd_no = 54
    Const S446_E1_recomnd_dt = 55
    Const S446_E1_trans_method = 56
    Const S446_E1_trans_rep_cd = 57
    Const S446_E1_trans_from_dt = 58
    Const S446_E1_trans_to_dt = 59
    Const S446_E1_customs = 60
    Const S446_E1_final_dest = 61
    Const S446_E1_remark1 = 62
    Const S446_E1_remark2 = 63
    Const S446_E1_remark3 = 64
    Const S446_E1_origin = 65
    Const S446_E1_origin_cntry = 66
    Const S446_E1_usd_xch_rate = 67
    Const S446_E1_biz_area = 68
    Const S446_E1_ref_flag = 69
    Const S446_E1_sts = 70
    Const S446_E1_net_weight = 71
    Const S446_E1_ext1_qty = 72
    Const S446_E1_ext2_qty = 73
    Const S446_E1_ext3_qty = 74
    Const S446_E1_ext1_amt = 75
    Const S446_E1_ext2_amt = 76
    Const S446_E1_ext3_amt = 77
    Const S446_E1_ext1_cd = 78
    Const S446_E1_ext2_cd = 79
    Const S446_E1_ext3_cd = 80
    Const S446_E1_xch_rate_op = 81

    Const S446_E2_bp_cd = 82   
    Const S446_E2_bp_nm = 83
    Const S446_E3_bp_cd = 84
    Const S446_E3_bp_nm = 85
    Const S446_E4_sales_grp = 86
    Const S446_E4_sales_grp_nm = 87
    Const S446_E5_sales_org = 88
    Const S446_E5_sales_org_nm = 89
    Const S446_E6_bp_nm = 90    
    Const S446_E7_bp_nm = 91        
    Const S446_E8_bp_nm = 92
    Const S446_E9_bp_nm = 93
    Const S446_E10_bank_nm = 94
    Const S446_E11_minor_nm = 95
    Const S446_E12_minor_nm = 96
    Const S446_E13_minor_nm = 97
    Const S446_E14_minor_nm = 98    
    Const S446_E15_minor_nm = 99    
    Const S446_E16_minor_nm = 100    
    Const S446_E17_minor_nm = 101    
    Const S446_E18_minor_nm = 102    
    Const S446_E19_country_nm = 103    
    Const S446_E20_minor_nm = 104      
    Const S446_E21_country_nm = 105    
    Const S446_E22_minor_nm = 106      
    Const S446_E23_country_nm = 107
    Const S446_E24_minor_nm = 108      
    Const S446_E25_minor_nm = 109  
    Const S446_E26_minor_nm = 110      
    Const S446_E27_minor_nm = 111  
    Const S446_E28_bp_nm = 112     
    Const S446_E29_minor_nm = 113  
    Err.Clear 

    If Request("txtCCNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Response.End
	End If
	
	I1_s_cc_hdr_cc_no  = Trim(Request("txtCCNo"))
	iCommandSent = "LOOKUP"
	
	Set pS6G219 = Server.CreateObject("PS6G219.cSLkExportCcHdrSvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If  
    
    Call pS6G219.S_LOOKUP_EXPORT_CC_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_cc_hdr_cc_no, EG1_s_cc_hdr )
      
	If CheckSYSTEMError(Err,True) = True Then
       Set pS6G219 = Nothing
       Response.End
    End If      
    Set pS6G219 = Nothing
    
%>
<Script Language=VBScript>
	With parent.frm1
		.txtSeller.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E3_bp_cd))%>"
		.txtSellerNm.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E3_bp_nm))%>"
		.txtConsignee.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E2_bp_cd))%>"
		.txtConsigneeNm.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E2_bp_nm))%>"

		.txtDeDate.text = "<%=UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_ship_fin_dt))%>"

		.txtVessel.value =  "<%=ConvSPChars(EG1_s_cc_hdr(S446_E1_vessel_nm))%>"
		.txtFrom.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E1_loading_port))%>"
		.txtFromNm.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E18_minor_nm))%>"
		.txtTo.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E1_dischge_port))%>"
		.txtToNm.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E20_minor_nm))%>"
		.txtIVNo.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E1_iv_no))%>"

		.txtIVDate.text = "<%=UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_iv_dt))%>"

		.txtHCCNo.value = "<%=ConvSPChars(EG1_s_cc_hdr(S446_E1_cc_no))%>"
	End With
</Script>
<%
		Set S42119 = Nothing														'☜: Unload Comproxy

		Response.End																'☜: Process End
%>