<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
	call LoadBasisGlobalInf()
	call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
 
 On Error Resume Next
 Err.Clear 
 
 Call HideStatusWnd
 
 Dim lgOpModeCRUD
 lgOpModeCRUD = Request("txtMode")

 Select Case lgOpModeCRUD
         Case CStr(UID_M0001)                                                         '☜: Query
              Call SubBizQuery()
         Case "LookupDailyExRt"
			  Call SubLookupDailyExRt()
 End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================

Sub SubBizQuery()
 
 On Error Resume Next
 Err.Clear 

 Dim iPM5G1H9                ' 수출 B/L Header 조회용 Object
 Dim strDt
 Dim strDefDate
  
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
 Const M538_E12_remark = 116
    
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
 Dim str_txtblno
 
 
 '---------------------------------- B/L Header Data Query ----------------------------------
 Set iPM5G1H9 = Server.CreateObject("PM5G1H9.cMLkImportBlHdrS")
 
 If CheckSYSTEMError(Err,True) = True Then
	Set iPM5G1H9 = Nothing
	Exit Sub
 End If
 
 str_txtblno = Trim(Request("txtBLNo"))
 Call iPM5G1H9.M_LOOKUP_IMPORT_BL_HDR_SVR(gStrGlobalCollection, _
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
	Set iPM5G1H9 = Nothing            '☜: ComProxy Unload
	Exit Sub               '☜: 비지니스 로직 처리를 종료함 
 End If

 Set iPM5G1H9 = Nothing            '☜: ComProxy Unload

 strDefDate = UNIDateClientFormat("1900-01-01")
 
 Response.Write "<Script Language=VBScript>" & vbCr
 Response.Write "With parent.frm1" & vbCr
 '##### Rounding Logic #####
 '항상 거래화폐가 우선 
 Response.Write ".txtCurrency.value =  """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_currency))) & """" & vbCr
 Response.Write "parent.CurFormatNumericOCX" & vbCr
 '##########################

 Response.Write ".txtBLNo1.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_bl_no))) & """" & vbCr
 Response.Write ".txtBLDocNo.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_bl_doc_no))) & """" & vbCr
 Response.Write ".txtPONo.value   = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_po_no))) & """" & vbCr
 
 If ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_po_no))) <> ""   Then
  Response.Write ".chkPoNoCnt.checked = true" & vbCr
 Else
  Response.Write ".chkPoNoCnt.checked = false" & vbCr
 End If


 If ConvSPChars(Trim(E1_ief_supplied_pp_flg)) = "Y" Then
  Response.Write ".ChkPrepay.checked = true" & vbCr
 Else
  Response.Write ".ChkPrepay.checked = false" & vbCr
 End If
  
 Response.Write ".txtLCDocNo.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_lc_doc_no))) & """" & vbCr
 Response.Write ".txtLCNo.value   = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_lc_no)))  & """" & vbCr
 
 If ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_lc_doc_no))) <> ""   Then
  Response.Write ".chkLcNoCnt.checked = true" & vbCr
 Else
  Response.Write ".chkLcNoCnt.checked = false" & vbCr
 End If

 strDt = UNIDateClientFormat(E12_m_bl_hdr(M538_E12_loading_dt))
 If strDt <> strDefDate Then
  Response.Write ".txtLoadingDt.text = """ & strDt & """" & vbCr
 End If
  
 strDt = UNIDateClientFormat(E12_m_bl_hdr(M538_E12_dischge_dt))
 If strDt <> strDefDate Then
  Response.Write ".txtDischgeDt.Text = """ & strDt & """" & vbCr
 End If
  
 strDt = UNIDateClientFormat(E12_m_bl_hdr(M538_E12_bl_issue_dt))
 If strDt <> strDefDate Then
  Response.Write ".txtBLIssueDt.Text = """ & strDt & """" & vbCr
 End If
 strDt = UNIDateClientFormat(E12_m_bl_hdr(M538_E12_setlmnt_dt))
 '지불예정일을 입력하지 않은 경우 2999/12/31로 셋팅함(2003.09.22)
 If strDt <> UNIDateClientFormat("2999-12-31") Then
  Response.Write ".txtSetlmnt.Text = """ & strDt & """" & vbCr
 Else
  Response.Write ".txtSetlmnt.Text = """"" & vbCr
 End If
  
 'Response.Write ".txtSetlmnt.Text		= """& UNIDateClientFormat(E12_m_bl_hdr(M538_E12_setlmnt_dt))	& """" & vbCr
 Response.Write ".txtTransport.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_transport)))		& """" & vbCr
 Response.Write ".txtTransportNm.value  = """ & ConvSPChars(Trim(E22_b_minor_transport))					& """" & vbCr
 Response.Write ".txtForwarder.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_forwarder)))		& """" & vbCr
 Response.Write ".txtForwarderNm.value  = """ & ConvSPChars(Trim(E18_b_biz_partner_forwarder))			& """" & vbCr
 Response.Write ".txtVesselNm.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_vessel_nm)))		& """" & vbCr
 Response.Write ".txtVoyageNo.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_voyage_no)))		& """" & vbCr
 Response.Write ".txtTotPackingCnt.value= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_tot_packing_cnt))) & """" & vbCr
 Response.Write ".txtVesselCntry.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_vessel_cntry)))	& """" & vbCr
 Response.Write ".txtBLIssuePlce.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_bl_issue_plce)))	& """" & vbCr
 Response.Write ".txtBLIssueCnt.text	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_bl_issue_cnt)))	& """" & vbCr
  
 Response.Write ".txtDocAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E12_m_bl_hdr(M538_E12_doc_amt), E12_m_bl_hdr(M538_E12_currency), ggAmtOfMoneyNo,"X","X") & """" & vbCr
 Response.Write ".txtLocAmt.text		= """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_loc_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
 Response.Write ".txtPayType.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_pay_type)))		& """" & vbCr
 Response.Write ".txtPayTypeNm.value	= """ & ConvSPChars(Trim(E28_b_minor_pay_type))					& """" & vbCr
 Response.Write ".txtPayTermsTxt.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_pay_terms_txt)))	& """" & vbCr
 Response.Write ".txtPayMethod.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_pay_method)))		& """" & vbCr
 Response.Write ".txtPayMethodNm.value  = """ & ConvSPChars(Trim(E23_b_minor_pay_method))				& """" & vbCr
 Response.Write ".txtPayDur.text		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_pay_dur)))		& """" & vbCr  
 Response.Write ".txtIncoterms.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_incoterms)))		& """" & vbCr
 'Response.Write ".txtLoanNo.value		= """& ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_loan_no)))		& """" & vbCr
 'Response.Write ".txtLoanAmt.text		= """& UNINumClientFormat(E12_m_bl_hdr(M538_E12_loan_doc_amt), ggAmtOfMoney.DecPoint, 0)& """" & vbCr
 Response.Write ".txtIvNo.value			= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_ref_iv_no)))		& """" & vbCr
  
 If ConvSPChars(E12_m_bl_hdr(M538_E12_posting_flg)) = "Y" Then
  Response.Write ".rdoPostingflg1.Checked = True" & vbCr
 ElseIf ConvSPChars(E12_m_bl_hdr(M538_E12_posting_flg)) = "N" Then
  Response.Write ".rdoPostingflg2.Checked = True" & vbCr
 End If
  
 Response.Write ".txtPost.Value			= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_posting_flg)))	& """" & vbCr
 Response.Write ".txtIvType.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_iv_type)))		& """" & vbCr
 Response.Write ".txtIvTypeNm.value		= """ & ConvSPChars(Trim(E5_m_iv_type_iv_type))					& """" & vbCr

 Response.Write ".txtPayeeCd.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_payee_cd)))		& """" & vbCr
 Response.Write ".txtPayeeNm.value		= """ & ConvSPChars(Trim(E7_b_biz_partner_payee))				& """" & vbCr
 Response.Write ".txtBuildCd.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_build_cd)))		& """" & vbCr
 Response.Write ".txtBuildNm.value		= """ & ConvSPChars(Trim(E6_b_biz_partner_build))				& """" & vbCr

 Response.Write ".txtBeneficiary.value  = """ & ConvSPChars(Trim(E9_b_biz_partner(M538_E9_bp_cd)))		& """" & vbCr
 Response.Write ".txtBeneficiaryNm.value= """ & ConvSPChars(Trim(E9_b_biz_partner(M538_E9_bp_nm)))		& """" & vbCr
 Response.Write ".txtPackingType.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_packing_type)))	& """" & vbCr
 Response.Write ".txtPackingTxt.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_packing_txt)))	& """" & vbCr 
 Response.Write ".txtPackingTypeNm.value= """ & ConvSPChars(Trim(E24_b_minor_packing_type))				& """" & vbCr  
 Response.Write ".txtGrossWeight.text	= """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_gross_weight), ggQty.DecPoint, 0) & """" & vbCr
 Response.Write ".txtNetWeight.Text		= """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_net_weight), ggQty.DecPoint, 0) & """" & vbCr
 Response.Write ".txtWeightUnit.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_weight_unit)))	& """" & vbCr
 Response.Write ".txtContainerCnt.value = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_container_cnt)))	& """" & vbCr
 Response.Write ".txtGrossVolumn.text	= """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_gross_volume), ggQty.DecPoint, 0) & """" & vbCr
 Response.Write ".txtVolumnUnit.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_volume_unit)))	& """" & vbCr
 Response.Write ".txtFreight.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_freight)))		& """" & vbCr
 Response.Write ".txtFreightPlce.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_freight_plce)))	& """" & vbCr
 Response.Write ".txtFreightNm.value	= """ & ConvSPChars(Trim(E25_b_minor_freight))					& """" & vbCr
 Response.Write ".txtFinalDest.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_final_dest)))		& """" & vbCr
 Response.Write ".txtDeliveryPlce.value = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_delivery_plce)))	& """" & vbCr
 Response.Write ".txtDeliveryPlceNm.value=""" & ConvSPChars(Trim(E21_b_minor_delivery_plce))				& """" & vbCr
 Response.Write ".txtReceiptPlce.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_receipt_plce)))	& """" & vbCr
 Response.Write ".txtTranshipCntry.value= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_tranship_cntry)))	& """" & vbCr
        
 'Response.Write "strDt = """& UNIDateClientFormat(E12_m_bl_hdr(M538_E12_tranship_dt))& """" & vbCr
 If strDt <> strDefDate Then
  Response.Write ".txtTranshipDt.text	= """ & UNIDateClientFormat(E12_m_bl_hdr(M538_E12_tranship_dt))	& """" & vbCr
 End If

 Response.Write ".txtLoadingPort.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_loading_port)))	& """" & vbCr
 Response.Write ".txtLoadingPortNm.value= """ & ConvSPChars(Trim(E19_b_minor_loading_port))				& """" & vbCr
 Response.Write ".txtDischgePort.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_dischge_port)))	& """" & vbCr
 Response.Write ".txtDischgePortNm.value= """ & ConvSPChars(Trim(E20_b_minor_dischge_port))				& """" & vbCr
  
 Response.Write ".txtOrigin.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_origin)))			& """" & vbCr
 Response.Write ".txtOriginNm.value		= """ & ConvSPChars(Trim(E26_b_minor_origin))					& """" & vbCr
 Response.Write ".txtOriginCntry.value  = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_origin_cntry)))	& """" & vbCr
 Response.Write ".txtAgent.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_agent)))			& """" & vbCr
 Response.Write ".txtAgentNm.value		= """ & ConvSPChars(Trim(E17_b_biz_partner_agent))				& """" & vbCr 
 Response.Write ".txtManufacturer.value = """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_manufacturer)))	& """" & vbCr
 Response.Write ".txtManufacturerNm.value=""" & ConvSPChars(Trim(E16_b_biz_partner_manufacturer))		& """" & vbCr 
 Response.Write ".txtPurGrp.value		= """ & ConvSPChars(Trim(E13_b_pur_grp(M538_E13_pur_grp)))		& """" & vbCr
 Response.Write ".txtPurGrpNm.value		= """ & ConvSPChars(Trim(E13_b_pur_grp(M538_E13_pur_grp_nm)))	& """" & vbCr
 Response.Write ".txtTaxBizArea.value	= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_tax_biz_area)))	& """" & vbCr
 Response.Write ".txtTaxBizAreaNm.value = """ & ConvSPChars(Trim(E8_b_biz_area_tax_biz_area))			& """" & vbCr
 Response.Write ".txtPurOrg.value		= """ & ConvSPChars(Trim(E14_b_pur_org(M538_E14_pur_org)))		& """" & vbCr
 Response.Write ".txtPurOrgNm.value		= """ & ConvSPChars(Trim(E14_b_pur_org(M538_E14_pur_org_nm)))	& """" & vbCr 
 Response.Write ".txtApplicant.value	= """ & ConvSPChars(Trim(E15_b_biz_partner(M538_E15_bp_cd)))		& """" & vbCr
 Response.Write ".txtApplicantNm.value  = """ & ConvSPChars(Trim(E15_b_biz_partner(M538_E15_bp_nm)))		& """" & vbCr
 Response.Write ".txtApplicantNm.value  = """ & ConvSPChars(Trim(E15_b_biz_partner(M538_E15_bp_nm)))		& """" & vbCr
 Response.Write ".txtremark.value		= """ & ConvSPChars(Trim(E12_m_bl_hdr(M538_E12_remark)))			& """" & vbCr
 Response.Write ".hdnLoanflg.value		= """ & ConvSPChars(Trim(E4_ief_supplied_loan_flg))				& """" & vbCr
 
 Response.Write ".txtGlNo.Value			= """ & ConvSPChars(Trim(E3_m_iv_hdr_gl_no))						& """" & vbCr
 Response.Write ".hdnGlType.value		= """ & ConvSPChars(Trim(E2_ief_supplied_gl_type))				& """" & vbCr
 Response.Write ".txtXchRate.text		= """ & UNINumClientFormat(E12_m_bl_hdr(M538_E12_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
    
 Response.Write "Call parent.DbQueryOk() " & vbCr             '☜: 조회가 성공 

 Response.Write ".txtHBLNo.value    = """ & ConvSPChars(Trim(Request("txtBLNo"))) & """" & vbCr
 Response.Write "End With"		& vbCr
 Response.Write "</Script>"		& vbCr

 Set iPM5G1H9 = Nothing              '☜: ComProxy UnLoad
End Sub
'============================================================================================================
' Name : SubLookupDailyExRt
' Desc :
'============================================================================================================
Sub SubLookupDailyExRt()
 
  
  On Error Resume Next
  Err.Clear                                                               '☜: Protect system from crashing
  
  Dim B17014
  Dim E1_b_daily_exchange_rate
  Const B253_E1_std_rate=0
  Const B253_E1_multi_divide=1
  
  Dim str_txtcurrency
  Dim str_txtblissuedt
  
  Set B17014 = Server.CreateObject("PB0C004.CB0C004")

  If CheckSYSTEMError(Err,True) = True Then
	Set B17014 = Nothing
	Exit Sub
  End If
  
  str_txtcurrency = Request("txtCurrency")
  str_txtblissuedt = UNIConvDate(Request("txtBlIssueDt"))
  E1_b_daily_exchange_rate = B17014.B_SELECT_EXCHANGE_RATE(gStrGlobalCollection, _
                         str_txtcurrency, _
                         gCurrency, _
                         str_txtblissuedt)

  If CheckSYSTEMError(Err,True) = True Then
	Set B17014 = Nothing
	Exit Sub
  End If

   Set B17014 = Nothing                                                   '☜: Unload Comproxy

  Response.Write "<Script Language=VBScript>" & vbCr
  Response.Write " With parent.frm1"    & vbCr
  Response.Write "IF " & Trim(E1_b_daily_exchange_rate(B253_E1_std_rate)) & " <> 0 THEN " & vbCr   
  Response.Write " .hdnDiv.value = """ & ConvSPChars(E1_b_daily_exchange_rate(B253_E1_multi_divide)) & """" & vbCr
  Response.Write " .txtXchRate.value = """ & UNINumClientFormat(E1_b_daily_exchange_rate(B253_E1_std_rate), ggExchRate.DecPoint, 0)    & """" & vbCr
  Response.Write "ElSE    " & vbCr
  Response.Write " .txtXchRate.value = 0 " & vbCr
  Response.Write "END IF " & vbCr
  Response.Write " End With"        & vbCr
  Response.Write "</Script>"        & vbCr

    Set B17014 = Nothing                                                   '☜: Unload Comproxy
End Sub
%>

