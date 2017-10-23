<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

	On Error Resume Next													'☜: Protect system from crashing
	Err.Clear 																'☜: Clear Error status
	
	

	Dim lgCurrency

	Dim OBJ_PM52119																' S/O Header 조회용 Object
	
	Dim E1_ief_supplied_pp_flg
	Dim E2_ief_supplied_gl_type
	Dim E3_m_iv_hdr_gl_no
	Dim E4_ief_supplied_loan_flg
	Dim E5_m_iv_type_iv_type
	Dim E6_b_biz_partner_build
	Dim E7_b_biz_partner_payee
	Dim E8_b_biz_area_tax_biz_area
	Dim E9_b_biz_partner
	Dim E10_b_minor_incoterms
	Dim E12_m_bl_hdr
	Dim E13_b_pur_grp
	Dim E14_b_pur_org
	Dim E15_b_biz_partner
	Dim E16_b_biz_partner_manufacturer
	Dim E17_b_biz_partner_agent
	Dim E18_b_biz_partner_forwarder
	Dim E19_b_minor_loading_port
	Dim E20_b_minor_dischge_port
	Dim E21_b_minor_delivery_plce
	Dim E22_b_minor_transport
	Dim E23_b_minor_pay_method
	Dim E24_b_minor_packing_type
	Dim E25_b_minor_freight
	Dim E26_b_minor_origin
	Dim E27_b_minor_fund_type
	Dim E28_b_minor_pay_type
	Dim E29_b_bank_open_bank
	
	'View Name : export_beneficiary b_biz_partner
	Const M538_E9_bp_cd = 0
	Const M538_E9_bp_nm = 1

	'View Name : export m_bl_hdr
	Const M538_E12_bl_no = 0
	Const M538_E12_po_no = 1
	Const M538_E12_lc_no = 2
	Const M538_E12_lc_amend_seq = 3
	Const M538_E12_lc_doc_no = 4
	Const M538_E12_manufacturer = 5
	Const M538_E12_agent = 6
	Const M538_E12_bl_doc_no = 7
	Const M538_E12_receipt_plce = 8
	Const M538_E12_vessel_nm = 9
	Const M538_E12_voyage_no = 10
	Const M538_E12_forwarder = 11
	Const M538_E12_vessel_cntry = 12
	Const M538_E12_loading_port = 13
	Const M538_E12_dischge_port = 14
	Const M538_E12_delivery_plce = 15
	Const M538_E12_loading_dt = 16
	Const M538_E12_dischge_dt = 17
	Const M538_E12_transport = 18
	Const M538_E12_tranship_cntry = 19
	Const M538_E12_tranship_dt = 20
	Const M538_E12_final_dest = 21
	Const M538_E12_currency = 22
	Const M538_E12_doc_amt = 23
	Const M538_E12_xch_rate = 24
	Const M538_E12_loc_amt = 25
	Const M538_E12_incoterms = 26
	Const M538_E12_pay_method = 27
	Const M538_E12_pay_dur = 28
	Const M538_E12_packing_type = 29
	Const M538_E12_tot_packing_cnt = 30
	Const M538_E12_container_cnt = 31
	Const M538_E12_packing_txt = 32
	Const M538_E12_gross_weight = 33
	Const M538_E12_weight_unit = 34
	Const M538_E12_gross_volume = 35
	Const M538_E12_volume_unit = 36
	Const M538_E12_freight = 37
	Const M538_E12_freight_plce = 38
	Const M538_E12_bl_issue_cnt = 39
	Const M538_E12_bl_issue_plce = 40
	Const M538_E12_bl_issue_dt = 41
	Const M538_E12_origin = 42
	Const M538_E12_origin_cntry = 43
	Const M538_E12_insrt_user_id = 44
	Const M538_E12_insrt_dt = 45
	Const M538_E12_updt_user_id = 46
	Const M538_E12_updt_dt = 47
	Const M538_E12_posting_flg = 48
	Const M538_E12_ext1_qty = 49
	Const M538_E12_cash_doc_amt = 50
	Const M538_E12_ext1_cd = 51
	Const M538_E12_vat_type = 52
	Const M538_E12_vat_rate = 53
	Const M538_E12_vat_doc_amt = 54
	Const M538_E12_vat_loc_amt = 55
	Const M538_E12_lc_open_dt = 56
	Const M538_E12_lc_type = 57
	Const M538_E12_open_bank = 58
	Const M538_E12_pay_terms_txt = 59
	Const M538_E12_pay_type = 60
	Const M538_E12_net_weight = 61
	Const M538_E12_biz_area = 62
	Const M538_E12_tax_biz_area = 63
	Const M538_E12_cost_cd = 64
	Const M538_E12_pre_pay_no = 65
	Const M538_E12_pre_pay_doc_amt = 66
	Const M538_E12_pre_pay_loc_amt = 67
	Const M538_E12_loan_no = 68
	Const M538_E12_iv_type = 69
	Const M538_E12_trans_type = 70
	Const M538_E12_sppl_iv_no = 71
	Const M538_E12_sppl_iv_dt = 72
	Const M538_E12_setlmnt_dt = 73
	Const M538_E12_setlmnt_bank = 74
	Const M538_E12_fund_type = 75
	Const M538_E12_setlmnt_cur = 76
	Const M538_E12_setlmnt_xch_rt = 77
	Const M538_E12_setlmnt_doc_amt = 78
	Const M538_E12_setlmnt_loc_amt = 79
	Const M538_E12_usd_xch_rt = 80
	Const M538_E12_usd_xch_amt = 81
	Const M538_E12_xch_comm_doc_amt = 82
	Const M538_E12_trust_rcp_doc_amt = 83
	Const M538_E12_repay_doc_amt = 84
	Const M538_E12_repay_dt = 85
	Const M538_E12_accpt_no = 86
	Const M538_E12_lg_doc_no = 87
	Const M538_E12_lg_dt = 88
	Const M538_E12_lg_bank = 89
	Const M538_E12_lg_xch_rt = 90
	Const M538_E12_guar_doc_amt = 91
	Const M538_E12_guar_loc_amt = 92
	Const M538_E12_mst_bl_doc_no = 93
	Const M538_E12_charge_flg = 94
	Const M538_E12_cash_loc_amt = 95
	Const M538_E12_loan_doc_amt = 96
	Const M538_E12_loan_loc_amt = 97
	Const M538_E12_ref_iv_no = 98
	Const M538_E12_payee_cd = 99
	Const M538_E12_build_cd = 100
	Const M538_E12_bl_rcpt_dt = 101
	Const M538_E12_ext2_qty = 102
	Const M538_E12_ext3_qty = 103
	Const M538_E12_ext1_amt = 104
	Const M538_E12_ext2_amt = 105
	Const M538_E12_ext3_amt = 106
	Const M538_E12_ext2_cd = 107
	Const M538_E12_ext3_cd = 108
	Const M538_E12_ext1_rt = 109
	Const M538_E12_ext2_rt = 110
	Const M538_E12_ext3_rt = 111
	Const M538_E12_ext1_dt = 112
	Const M538_E12_ext2_dt = 113
	Const M538_E12_ext3_dt = 114
	Const M538_E12_xch_rate_op = 115

	'View Name : export b_pur_grp
	Const M538_E13_pur_grp = 0
	Const M538_E13_pur_grp_nm = 1

	'View Name : export b_pur_org
	Const M538_E14_pur_org = 0
	Const M538_E14_pur_org_nm = 1

	'View Name : export_applicant b_biz_partner
	Const M538_E15_bp_cd = 0
	Const M538_E15_bp_nm = 1


	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
	
	Dim iStrBlNo
	 
	Call HideStatusWnd
	
	'---------------------------------- B/L Header Data Query ----------------------------------
	Set OBJ_PM52119 = Server.CreateObject("PM5G1H9.cMLkImportBlHdrS")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	
	iStrBlNo = Trim(Request("txtBLNo"))
	
	Call OBJ_PM52119.M_LOOKUP_IMPORT_BL_HDR_SVR(gStrGlobalCollection, _
				            iStrBlNo, E1_ief_supplied_pp_flg, _
				            E2_ief_supplied_gl_type, E3_m_iv_hdr_gl_no, _
				            E4_ief_supplied_loan_flg, E5_m_iv_type_iv_type, _
				            E6_b_biz_partner_build, E7_b_biz_partner_payee, _
				            E8_b_biz_area_tax_biz_area, E9_b_biz_partner, _
				            E10_b_minor_incoterms, E12_m_bl_hdr, _
				            E13_b_pur_grp, E14_b_pur_org, _
				            E15_b_biz_partner, E16_b_biz_partner_manufacturer, _
				            E17_b_biz_partner_agent, E18_b_biz_partner_forwarder, _
				            E19_b_minor_loading_port, E20_b_minor_dischge_port, _
				            E21_b_minor_delivery_plce, E22_b_minor_transport, _
				            E23_b_minor_pay_method, E24_b_minor_packing_type, _
				            E25_b_minor_freight, E26_b_minor_origin, _
				            E27_b_minor_fund_type, E28_b_minor_pay_type, _
				            E29_b_bank_open_bank)
	
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
			Set OBJ_PM52119 = Nothing
			Response.End
	End If
	
	Set OBJ_PM52119 = Nothing														'☜: Unload Comproxy
	
	'-----------------------
	'Result data display area
	'-----------------------

	lgCurrency = ConvSPChars(E12_m_bl_hdr(M538_E12_currency))

	Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr

			'##### Rounding Logic #####
			'항상 거래화폐가 우선 
			Response.Write ".txtCurrency.value = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_currency)) & """" & vbCr
			Response.Write "parent.CurFormatNumericOCX" & vbCr
			'##########################

			'========= TAB 1 (수입신고) ==========
			Response.Write ".txtDischgeDt.Text		= """ & UNIDateClientFormat(E12_m_bl_hdr(M538_E12_dischge_dt)) & """" & vbCr							'도착일 
			Response.Write ".txtGrossWeight.Text	= """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_gross_weight), ggQty.DecPoint, 0) & """" & vbCr		'총중량 
			Response.Write "If """ & E12_m_bl_hdr(M538_E12_net_weight) & """  = 0 Then " & vbCr														'순중량 
				Response.Write ".txtNetWeight.Text	= """ & E12_m_bl_hdr(M538_E12_net_weight) & """" & vbCr
			Response.Write "Else"	 & vbCr
				Response.Write ".txtNetWeight.Text	= """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_net_weight), ggQty.DecPoint, 0) & """" & vbCr		
			Response.Write "End If" & vbCr
			
			Response.Write ".txtWeightUnit.value	= """ & ConvSPChars(E12_m_bl_hdr(M538_E12_weight_unit)) & """" & vbCr								 '중량단위 
			Response.Write ".txtPackingType.value	= """ & ConvSPChars(E12_m_bl_hdr(M538_E12_packing_type)) & """" & vbCr								 '포장형태 
			Response.Write ".txtPackingTypeNm.value = """ &ConvSPChars(E24_b_minor_packing_type) & """" & vbCr						 '포장형태명 
			Response.Write ".txtDischgePortCd.value = """ &ConvSPChars(E12_m_bl_hdr(M538_E12_dischge_port)) & """" & vbCr								 '도착항 
			Response.Write ".txtDischgePortNm.value = """ &ConvSPChars(E20_b_minor_dischge_port) & """" & vbCr						 '도착항명 
			Response.Write ".txtTransport.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_transport)) & """" & vbCr									 '운송방법 
			Response.Write ".txtTransportNm.value	= """ &ConvSPChars(E22_b_minor_transport) & """" & vbCr							 '운송방법명 
			Response.Write ".txtPayTerms.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_pay_method)) & """" & vbCr									 '결제방법 
			Response.Write ".txtPayTermsNm.value	= """ &ConvSPChars(E23_b_minor_pay_method) & """" & vbCr							 '결제방법명 
			Response.Write ".txtPayDur.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_pay_dur)) & """" & vbCr											 '결제기간 
			Response.Write ".txtIncoterms.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_incoterms)) & """" & vbCr									 '가격조건 
			Response.Write ".txtCurrency.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_currency)) & """" & vbCr										 '화폐단위 
			
			Response.Write "If """ &E12_m_bl_hdr(M538_E12_doc_amt) & """ = 0 Then" & vbCr																 '통관금액 
				'Response.Write ".txtDocAmt.Text		= """ & E12_m_bl_hdr(M538_E12_doc_amt) & """" & vbCr
				Response.Write ".txtDocAmt.Text		= 0 " & vbCr
				Response.Write ".txtLocAmt.Text		= 0 " & vbCr
			Response.Write "Else" & vbCr
				Response.Write ".txtDocAmt.Text		= 0 " & vbCr
				Response.Write ".txtLocAmt.Text		= 0 " & vbCr
				'Response.Write ".txtDocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E12_m_bl_hdr(M538_E12_doc_amt), lgCurrency,ggAmtOfMoneyNo,"X","X")& """" & vbCr
			Response.Write "End If" & vbCr
			
			Response.Write ".txtXchRate.Text		= """ &UNINumClientFormat(E12_m_bl_hdr(M538_E12_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr			 '환율 
			
			Response.Write "if """ &ConvSPChars(UCase(Trim(E12_m_bl_hdr(M538_E12_currency)))) & """ = ""USD"" then" & vbCr
				Response.Write ".txtUSDXchRate.Text = """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
			Response.Write "end if" & vbCr
			
			Response.Write ".txtBeneficiary.value	= """ &ConvSPChars(E9_b_biz_partner(M538_E9_bp_cd)) & """" & vbCr						 '수출자 
			Response.Write ".txtBeneficiaryNm.value = """ &ConvSPChars(E9_b_biz_partner(M538_E9_bp_nm)) & """" & vbCr					 '수출자명 
			Response.Write ".txtApplicant.value		= """ &ConvSPChars(E15_b_biz_partner(M538_E15_bp_cd)) & """" & vbCr							 '수입자 
			Response.Write ".txtApplicantNm.value	= """ &ConvSPChars(E15_b_biz_partner(M538_E15_bp_nm)) & """" & vbCr						 '수입자명 


			'========= TAB 2 (수입신고 기타) ==========
			Response.Write ".txtVesselCntry.value	= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_vessel_cntry)) & """" & vbCr								 '선박국적 
			Response.Write ".txtLoadingPort.value	= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_loading_port)) & """" & vbCr								 '선적항 
			Response.Write ".txtLoadingPortNm.value = """ &ConvSPChars(E19_b_minor_loading_port) & """" & vbCr						 '선적항명 
			Response.Write ".txtLoadingDt.Text		= """ &UNIDateClientFormat(E12_m_bl_hdr(M538_E12_loading_dt)) & """" & vbCr							 '선적일 
			Response.Write ".txtOrigin.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_origin)) & """" & vbCr											 '원산지 
			Response.Write ".txtOriginNm.value		= """ &ConvSPChars(E26_b_minor_origin) & """" & vbCr								 '원산지명 
			Response.Write ".txtOriginCntry.value	= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_origin_cntry)) & """" & vbCr								 '원산지국가 
			Response.Write ".txtLCDocNo.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_lc_doc_no)) & """" & vbCr										 'L/C번호 
			
			Response.Write "if Trim(""" &ConvSPChars(E12_m_bl_hdr(M538_E12_lc_doc_no)) & """) <> """" then" & vbCr										 'L/C순번 
				Response.Write ".txtLCAmendSeq.value = """ &ConvSPChars(E12_m_bl_hdr(M538_E12_lc_amend_seq)) & """" & vbCr
			Response.Write "End if" & vbCr
			
			Response.Write ".txtAgentCd.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_agent)) & """" & vbCr											 '대행자 
			Response.Write ".txtAgentNm.value		= """ &ConvSPChars(E17_b_biz_partner_agent) & """" & vbCr								 '대행자명 
			Response.Write ".txtManufacturerCd.value = """ &ConvSPChars(E12_m_bl_hdr(M538_E12_manufacturer)) & """" & vbCr							 '제조자 
			Response.Write ".txtManufacturerNm.value = """ &ConvSPChars(E16_b_biz_partner_manufacturer) & """" & vbCr					 '제조자명 
			Response.Write ".txtPurGrp.value		= """ &ConvSPChars(E13_b_pur_grp(M538_E13_pur_grp)) & """" & vbCr										 '수입담당 
			Response.Write ".txtPurGrpNm.value		= """ &ConvSPChars(E13_b_pur_grp(M538_E13_pur_grp_nm)) & """" & vbCr									 '수입담당명 
			Response.Write ".txtPurOrg.value		= """ &ConvSPChars(E14_b_pur_org(M538_E14_pur_org)) & """" & vbCr										 '수입부서 
			Response.Write ".txtPurOrgNm.value		= """ &ConvSPChars(E14_b_pur_org(M538_E14_pur_org_nm)) & """" & vbCr									 '수입부서명 


			'-------- Hidden Value ---------
			Response.Write ".txtLcNo.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_lc_no)) & """" & vbCr												 'LC관리번호 
			Response.Write ".txtLcType.value	= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_lc_type)) & """" & vbCr											 'LC유형 
			Response.Write ".txtLcOpenDt.value	= """ &UNIDateClientFormat(E12_m_bl_hdr(M538_E12_lc_open_dt)) & """" & vbCr								 'LC개설일 
			Response.Write ".txtPoNo.value		= """ &ConvSPChars(E12_m_bl_hdr(M538_E12_po_no)) & """" & vbCr												 'PO번호 
			
'???			Response.Write ".hdnDiv.value		= """ &ConvSPChars(MulDiv)& """" & vbCr														 '환율연산자 
			Response.Write "parent.dbBlQueryok()" & vbCr
			
		Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr

%>
