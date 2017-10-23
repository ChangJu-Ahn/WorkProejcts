<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
	call LoadBasisGlobalInf()
	call LoadInfTB19029B("I", "*","NOCOOKIE", "MB") 
	Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5212mb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입B/L내역등록 Query Transaction 처리용 ASP								*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2000/03/22																*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************

	Dim lgOpModeCRUD
	Dim lgCurrency

	On Error Resume Next					'☜: Protect system from crashing
	Err.Clear 								'☜: Clear Error status
				

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")	'☜: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD
	         Case CStr(UID_M0001)                                                         '☜: Query
	              Call SubBizQueryMulti()
	End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================

Sub SubBizQueryMulti()
	
	Dim iMax
	Dim PvArr
	Dim lGrpCnt
	Dim iTotstrData
	Const C_SHEETMAXROWS_D  = 100
	
	On Error Resume Next					'☜: Protect system from crashing
	Err.Clear																'☜: Protect system from crashing

	Dim iPM6G28C																' Master L/C Detail 조회용 Object
	Dim iM52119																' Master L/C Header 조회용 Object
	'====views for BL_HDR
	'View Name: export_beneficiary b_biz_partner
	Const M538_E9_bp_cd = 0
	Const M538_E9_bp_nm = 1
	'View Name: export m_bl_hdr
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
	Const M538_E12_pur_org = 115
	Const M538_E12_pur_grp = 116
	Const M538_E12_applicant = 117
	Const M538_E12_beneficiary = 118
		
	'View Name: export b_pur_grp
	Const M538_E13_pur_grp = 0
	Const M538_E13_pur_grp_nm = 1
		
	'View Name: export b_pur_org
	Const M538_E14_pur_org = 0
	Const M538_E14_pur_org_nm = 1
		
	'View Name: export_applicant b_biz_partner
	Const M538_E15_bp_cd = 0
	Const M538_E15_bp_nm = 1

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
		
	Dim strDt

'====views for BL_DTL
	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim intGroupCount          
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	Dim iCommandSent
	
		
	Dim I1_m_bl_hdr_bl_no
	Dim I2_m_bl_dtl_next
	Dim E1_m_bl_dtl_next
	Dim E2_m_bl_dtl_max
	Dim E3_m_bl_dtl_tot

	Dim EG1_export_group
	Const M534_EG1_E1_m_bl_dtl_bl_seq = 0
    Const M534_EG1_E1_m_bl_dtl_hs_cd = 1
    Const M534_EG1_E1_m_bl_dtl_qty = 2
    Const M534_EG1_E1_m_bl_dtl_price = 3
    Const M534_EG1_E1_m_bl_dtl_doc_amt = 4
    Const M534_EG1_E1_m_bl_dtl_loc_amt = 5
    Const M534_EG1_E1_m_bl_dtl_unit = 6
    Const M534_EG1_E1_m_bl_dtl_gross_weight = 7
    Const M534_EG1_E1_m_bl_dtl_net_weight = 8
    Const M534_EG1_E1_m_bl_dtl_volume_size = 9
    Const M534_EG1_E1_m_bl_dtl_cc_qty = 10
    Const M534_EG1_E1_m_bl_dtl_il_no = 11
    Const M534_EG1_E1_m_bl_dtl_il_seq = 12
    Const M534_EG1_E1_m_bl_dtl_ext1_qty = 13
    Const M534_EG1_E1_m_bl_dtl_ext1_amt = 14
    Const M534_EG1_E1_m_bl_dtl_ext1_cd = 15
    Const M534_EG1_E1_m_bl_dtl_biz_area = 16
    Const M534_EG1_E1_m_bl_dtl_cost_cd = 17
    Const M534_EG1_E1_m_bl_dtl_ext2_qty = 18
    Const M534_EG1_E1_m_bl_dtl_ext3_qtya = 19
    Const M534_EG1_E1_m_bl_dtl_ext2_amt = 20
    Const M534_EG1_E1_m_bl_dtl_ext3_amt = 21
    Const M534_EG1_E1_m_bl_dtl_ext2_cd = 22
    Const M534_EG1_E1_m_bl_dtl_ext3_cd = 23
    Const M534_EG1_E1_m_bl_dtl_ext1_rt = 24
    Const M534_EG1_E1_m_bl_dtl_ext2_rt = 25
    Const M534_EG1_E1_m_bl_dtl_ext3_rt = 26
    Const M534_EG1_E1_m_bl_dtl_ext1_dt = 27
    Const M534_EG1_E1_m_bl_dtl_ext2_dt = 28
    Const M534_EG1_E1_m_bl_dtl_ext3_dt = 29
    Const M534_EG1_E1_m_bl_dtl_tracking_no = 30
    Const M534_EG1_E2_b_hs_code_hs_nm = 31
    Const M534_EG1_E3_b_item_item_cd = 32
    Const M534_EG1_E3_b_item_item_nm = 33
    Const M534_EG1_E3_b_item_spec = 34
    Const M534_EG1_E3_b_item_item_acc = 35
    Const M534_EG1_E4_b_plant_plant_cd = 36
    Const M534_EG1_E4_b_plant_plant_nm = 37
    Const M534_EG1_E5_m_pur_ord_hdr_po_no = 38
    Const M534_EG1_E6_m_pur_ord_dtl_po_seq_no = 39
    Const M534_EG1_E6_m_pur_ord_dtl_tracking_no = 40
    Const M534_EG1_E6_m_pur_ord_dtl_po_qty = 41
    Const M534_EG1_E6_m_pur_ord_dtl_lc_qty = 42
    Const M534_EG1_E6_m_pur_ord_dtl_bl_qty = 43
    Const M534_EG1_E6_m_pur_ord_dtl_cc_qty = 44
    Const M534_EG1_E6_m_pur_ord_dtl_over_tol = 45
    Const M534_EG1_E6_m_pur_ord_dtl_under_tol = 46
    Const M534_EG1_E7_m_lc_hdr_lc_no = 47
    Const M534_EG1_E7_m_lc_hdr_lc_doc_no = 48
    Const M534_EG1_E7_m_lc_hdr_lc_amend_seq = 49
    Const M534_EG1_E8_m_lc_dtl_lc_seq = 50
    Const M534_EG1_E8_m_lc_dtl_over_tolerance = 51
    Const M534_EG1_E8_m_lc_dtl_under_tolerance = 52
    Const M534_EG1_E1_m_bl_dtl_remark = 53
		
    Dim str_txtblno

	lgStrPrevKey = Request("lgStrPrevKey")

'------- B/L Header Data Query -------------------------------------------------

'**수정(2003.06.11)-추가 100건을 조회시는 헤더는 조회하지 않는다**
	If Request("txtMaxRows") = 0 Then
		Set iM52119 = Server.CreateObject("PM5G1H9.cMLkImportBlHdrS")

		If CheckSYSTEMError(Err,True) = True Then
			Set iM52119 = Nothing
			Exit Sub
		End If
		str_txtblno = Trim(Request("txtBLNo"))
		
		Call iM52119.M_LOOKUP_IMPORT_BL_HDR_SVR(gStrGlobalCollection, _
								str_txtblno, E1_ief_supplied_pp_flg, _
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

		If CheckSYSTEMError2(Err,True,"","","","","") = true then 
			
			Set iM52119 = Nothing												
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			Exit Sub
																
		End If

		Set iM52119 = Nothing									


			lgCurrency = ConvSPChars(E12_m_bl_hdr(M538_E12_currency))
			Const strDefDate = "1899-12-30"


			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			'Response.Write "Dim strDt		" & vbCr
			'Response.Write "Dim strDefDate	" & vbCr
				
			'##### Rounding Logic #####
			'항상 거래화폐가 우선 
			Response.Write ".txtCurrency.value = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_currency)) & """" & vbCr
			Response.Write "parent.CurFormatNumericOCX	" & vbCr
			'##########################

			Response.Write ".txtHBLNo.value = """ & ConvSPChars(Request("txtBLNo")) & """" & vbCr
				
			Response.Write "strDefDate = """ & UNIDateClientFormat(strDefDate) &		 """" & vbCr
			'debug
			Response.Write ".hdnPONo.value = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_po_no)) & """" & vbCr
			Response.Write ".hdnLcNo.value = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_lc_no)) & """" & vbCr
			Response.Write ".txtBeneficiary.value = """ & ConvSPChars(E9_b_biz_partner(M538_E9_bp_cd)) & """"   & vbCr
			Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(E9_b_biz_partner(M538_E9_bp_nm)) & """" & vbCr
			
			Response.Write "strDt = """ & UNIDateClientFormat(E12_m_bl_hdr(M538_E12_bl_issue_dt)) & """" & vbCr

			Response.Write "If strDt <> strDefDate Then" & vbCr
			Response.Write "	.txtIssueDt.Text = strDt" & vbCr
			Response.Write "End If" & vbCr
				
			Response.Write ".txtXchRate.value     = """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
			Response.Write ".hdnDiv.value	      = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_xch_rate_op )) & """" & vbCr		'환율연산자 
			'Response.Write ".txtDocAmt.value      = """ & UNIConvNumDBToCompanyByCurrency(E12_m_bl_hdr(M538_E12_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X")& """" & vbCr
			Response.Write ".txtDocAmt.text      = """ & UNIConvNumDBToCompanyByCurrency(E12_m_bl_hdr(M538_E12_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	
			Response.Write ".txtMaxSeq.value      = 0 " & vbCr
			Response.Write ".hdnPayMethCd.Value   = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_pay_method)) & """" & vbCr
			Response.Write ".hdnPayMethNm.Value   = """ & ConvSPChars(E23_b_minor_pay_method) & """" & vbCr
			Response.Write ".hdnIncotermsCd.Value = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_incoterms)) & """" & vbCr
			Response.Write ".hdnIncotermsNm.Value = """ & ConvSPChars(E10_b_minor_incoterms) & """" & vbCr
			Response.Write ".hdnGrpCd.Value       = """ & ConvSPChars(E13_b_pur_grp(M538_E13_pur_grp)) & """" & vbCr
			Response.Write ".hdnGrpNm.Value       = """ & ConvSPChars(E13_b_pur_grp(M538_E13_pur_grp_nm)) & """" & vbCr
			Response.Write ".hdnIvNo.Value        = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_ref_iv_no)) & """" & vbCr
			Response.Write ".txtPost.Value        = """ & ConvSPChars(E12_m_bl_hdr(M538_E12_posting_flg)) & """" & vbCr
				
			Response.Write ".hdnGlNo.Value		  = """ & ConvSPChars(E3_m_iv_hdr_gl_no) & """" & vbCr
			Response.Write ".hdnGlType.value 	  = """ & ConvSPChars(E2_ief_supplied_gl_type) & """" & vbCr
				
			Response.Write "if """ & ConvSPChars(Trim(E1_ief_supplied_pp_flg)) & """ = ""Y"" then "	 & vbCr
			Response.Write "	.ChkPrepay.checked = true" & vbCr
			Response.Write "else"		   & vbCr
			Response.Write "	.ChkPrepay.checked = false" & vbCr
			Response.Write "End if"							& vbCr
				
			Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet"	& vbCr
			Response.Write "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE"	& vbCr
'				Response.Write "parent.CurFormatNumSprSheet"    & vbCr
			Response.Write "End With"	& vbCr
			Response.Write "</Script>"		& vbCr

		
	Else
		lgCurrency = request("txtCurrency")
	End If
		
		
'----- B/L Detail Data Query ----------------------------------------------------
		Set iPM6G28C = Server.CreateObject("PM6G28C.cMListImportBlDtlS")
		

		If CheckSYSTEMError(Err,True) = True Then
				Set iPM6G28C = Nothing
				Exit Sub
		End If

		I1_m_bl_hdr_bl_no = Request("txtBLNo")
		I2_m_bl_dtl_next = UNIConvNum(Request("lgStrPrevKey"),0)
		iCommandSent = "LIST"

		Call iPM6G28C.M_LIST_IMPORT_BL_DTL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, iCommandSent, I1_m_bl_hdr_bl_no, I2_m_bl_dtl_next, _ 
		                EG1_export_group, E1_m_bl_dtl_next, E2_m_bl_dtl_max, E3_m_bl_dtl_tot)

		if Request("txtQuerytype") <> "Query" And Err.Description = "B_MESSAGE175400" then
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			IF UBound(EG1_exp_group,1) <= 0 Then
				Set iPM6G28C = Nothing
				Exit Sub												'☜: ComProxy Unload	
			End If
		
		Else 
			If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
				Set iPM6G28C = Nothing												'☜: ComProxy Unload
				'Detail항목이 없을 경우 Header정보만 보여줌 
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
				Response.Write "parent.dbQueryOk" & chr(13)
				Response.Write "</Script>"
				Exit Sub															'☜: 비지니스 로직 처리를 종료함 
			End If
		End if
		
		iLngMaxRow = CInt(Request("txtMaxRows"))
		intGroupCount = UBound(EG1_export_group,1)
		
		if intGroupCount <> 0 then
			If EG1_export_group(intGroupCount,M534_EG1_E1_m_bl_dtl_bl_seq) = cstr(E1_m_bl_dtl_next) Then
		      	StrNextKey = ""
			Else
			  	StrNextKey = E1_m_bl_dtl_next
			End If
		End if
		
		ReDim PvArr(intGroupCount)
	

		For iLngRow = 0 To UBound(EG1_export_group,1)
			If  iLngRow < C_SHEETMAXROWS_D  Then
			Else
			   StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M534_EG1_E1_m_bl_dtl_bl_seq)) 
			   Exit For
			End If  
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E3_b_item_item_cd)) _
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E3_b_item_item_nm)) _					
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E3_b_item_spec)) _						
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_unit )) _						
			                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_qty), ggQty.DecPoint, 0) _
		 	                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_price), lgCurrency, ggUnitCostNo,"X","X") _
		 	                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _        
    		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_loc_amt),gCurrency,ggAmtOfMoneyNo,"X","X") _
			                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_gross_weight), ggQty.DecPoint, 0) _
			                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_volume_size), ggQty.DecPoint, 0) _        	
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_hs_cd)) _
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E2_b_hs_code_hs_nm)) _											
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_bl_seq)) _
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E5_m_pur_ord_hdr_po_no)) _
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E6_m_pur_ord_dtl_po_seq_no)) _
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E7_m_lc_hdr_lc_no))
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E8_m_lc_dtl_lc_seq ))

			if UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E6_m_pur_ord_dtl_lc_qty), ggQty.DecPoint, 0)	 > 0 then
				istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E8_m_lc_dtl_over_tolerance), ggExchRate.DecPoint, 0) _
				                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E8_m_lc_dtl_under_tolerance), ggExchRate.DecPoint, 0)
			else'발주tolerance로 수정(2003.07.01)
				istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E6_m_pur_ord_dtl_over_tol), ggExchRate.DecPoint, 0) _
				                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M534_EG1_E6_m_pur_ord_dtl_under_tol), ggExchRate.DecPoint, 0)
			end if
			
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E7_m_lc_hdr_lc_no)) _
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_tracking_no)) _
			                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_remark)) _
			                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _        
			                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,M534_EG1_E1_m_bl_dtl_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _        
			                    & Chr(11) & iLngMaxRow + iLngRow _								
			                    & Chr(11) & Chr(12)
			
			PvArr(lGrpCnt) = istrData
			lGrpCnt = lGrpCnt + 1
			istrData = ""
	    Next
	    
	    iTotstrData = Join(PvArr, "")
		
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & iTotstrData	    & """" & vbCr
    'Response.Write ".frm1.vspdData.ReDraw = False "			& vbCr
	'Response.Write ".SetSpreadColor -1	, -1	"				& vbCr
	'Response.Write ".frm1.vspdData.ReDraw = True	"		& vbCr	
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 

    Response.Write ".frm1.txtHBLNo.value = """ & ConvSPChars(Request("txtBLNo")) & """" & vbCr
    Response.write "	.DbQueryOk " & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>"	   & vbCr

		Set iPM6G28C = Nothing														'☜: Unload Comproxy
	
End Sub
'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
 Function ReturnZero(numData)

	if Len(Trim(numData)) < 1 then
		ReturnZero = "0"
	else
		ReturnZero = numData
	End if
	
End Function
%>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
