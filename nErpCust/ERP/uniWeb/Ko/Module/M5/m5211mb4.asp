<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
call LoadBasisGlobalInf()

'********************************************************************************************************
'*  1. Module Name          : Procuremant													            *
'*  2. Function Name        :																			*
'*  3. Program ID           : M5211mb4.asp												                *
'*  4. Program Name         :																			*
'*  5. Program Desc         : B/L Header Insert를 위한 P/O Header Data Query Transaction 처리용 ASP		*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/05/26																*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************


On Error Resume Next

Call HideStatusWnd

Dim M31119lc
Dim B1H028de
Dim B17014
Dim XchRate

 Dim E1_m_iv_type
 Dim E2_b_daily_exchange_rate
    Const M450_E2_multi_divide = 0
    Const M450_E2_std_rate = 1

 Dim E3_b_minor_delivery
    Const M450_E3_minor_cd = 0
    Const M450_E3_minor_nm = 1
    Const M450_E3_minor_type = 2

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

 Dim E15_b_biz_partner
    Const M450_E15_bp_cd = 0
    Const M450_E15_bp_type = 1
    Const M450_E15_bp_nm = 2

 Dim E16_b_pur_grp
    Const M450_E16_pur_grp = 0
    Const M450_E16_pur_grp_nm = 1

 Dim E17_b_pur_org
    Const M450_E17_pur_org = 0
    Const M450_E17_pur_org_nm = 1

    Dim E1_b_biz_partner_ssh
    Const B132_E1_bp_cd = 0
    Const B132_E1_bp_nm = 1
    Dim E2_b_biz_partner_sbi
    Const B132_E2_bp_cd = 0
    Const B132_E2_bp_nm = 1
    Dim E3_b_biz_partner_spa
    Const B132_E3_bp_cd = 0 
    Const B132_E3_bp_nm = 1
    Dim E4_b_biz_partner_mpa
    Const B132_E4_bp_cd = 0 
    Const B132_E4_bp_nm = 1
    Dim E5_b_biz_partner_mbi
    Const B132_E5_bp_cd = 0 
    Const B132_E5_bp_nm = 1
    Dim E6_b_biz_partner_mgs
    Const B132_E6_bp_cd = 0
    Const B132_E6_bp_nm = 1

 Dim str_pono
 
 Set M31119lc = Server.CreateObject("PM3G19P.cMLookupPoHdrS")

 If CheckSYSTEMError(Err,True) = True Then
	Set M52111 = Nothing
	Response.End
 End If
 str_pono = Request("txtPONo")
 Call M31119lc.M_LOOKUP_PO_HDR_SVR(gStrGlobalCollection, _
                   str_pono, E1_m_iv_type, _
                   E2_b_daily_exchange_rate, E3_b_minor_delivery, _
                   E4_b_minor_incoterms_nm, E5_b_minor_paymeth_nm, _
                   E6_b_minor_transport_nm, E7_b_minor_loading_nm, _
                   E8_b_minor_dischge_nm, E9_b_minor_origin_nm, _
                   E10_b_biz_partner_agent_nm, E11_b_biz_partner_manufacturer_nm, _
                   E12_b_biz_partner_applicant_nm, E13_b_bank_pay_bank_nm, _
                   E14_m_pur_ord_hdr, E15_b_biz_partner, _
                   E16_b_pur_grp, E17_b_pur_org)


 If CheckSYSTEMError(Err,True) = True Then
	Set M31119lc = Nothing
	Response.End
 End If

 Set M31119lc = Nothing

 '------------------------
 '지급처와 발행처를 Lookup
 '------------------------
 Set B1H028de = Server.CreateObject("PB5GS45.cBListDftBpFtnSvr")

 If CheckSYSTEMError(Err,True) = True Then
	Set B1H028de = Nothing
	Response.End
 End If

  Call B1H028de.B_LIST_DEFAULT_BP_FTN_SVR(gStrGlobalCollection, _
                     E15_b_biz_partner(M450_E15_bp_cd), _
					 E1_b_biz_partner_ssh, _
					 E2_b_biz_partner_sbi, _
					 E3_b_biz_partner_spa, _
					 E4_b_biz_partner_mpa, _
					 E5_b_biz_partner_mbi, _
					 E6_b_biz_partner_mgs)

 If CheckSYSTEMError(Err,True) = True Then
	Set B1H028de = Nothing
	Response.End
 End If

 Set B1H028de = Nothing
'-----------------------
'Result data display area
'-----------------------

 Response.Write "<Script Language=VBScript>" & vbcr
 Response.Write " With parent.frm1" & vbcr
 Response.Write "  '##### Rounding Logic #####" & vbcr
 Response.Write "  '항상 거래화폐가 우선" & vbcr
 Response.Write "  .txtCurrency.value = """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_po_cur))	& """" & vbcr
 Response.Write "  parent.CurFormatNumericOCX" & vbcr
 Response.Write "  '##########################" & vbcr

 Response.Write "  .txtPONo.value = """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_po_no))		& """" & vbcr
 Response.Write "  if """ & ConvSPChars(Trim(E14_m_pur_ord_hdr(M450_E14_po_no))) & """ <> """" then" & vbcr
 Response.Write "   .chkPoNoCnt.checked = true" & vbcr
 Response.Write "  else" & vbcr
 Response.Write "   .chkPoNoCnt.checked = false" & vbcr
 Response.Write "  End if" & vbcr
 Response.Write "  .chkLcNoCnt.checked = false" & vbcr
 Response.Write "  .txtTransport.value		= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_transport))		& """" & vbcr
 Response.Write "  .txtTransportNm.value	= """ & ConvSPChars(E6_b_minor_transport_nm)					& """" & vbcr
 Response.Write "  .txtDischgePort.value	= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_dischge_port))	& """" & vbcr
 Response.Write "  .txtLoadingPort.value	= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_loading_port))	& """" & vbcr
 Response.Write "  .txtAgent.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_agent))			& """" & vbcr
 Response.Write "  .txtManufacturer.value	= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_manufacturer))	& """" & vbcr
 Response.Write "  .txtDeliveryPlce.value	= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_delivery_plce))	& """" & vbcr
 Response.Write "  .txtPackingType.value	= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_packing_cond))	& """" & vbcr
 Response.Write "  .txtOrigin.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_origin))			& """" & vbcr
 Response.Write "  .txtPurGrp.value			= """ & ConvSPChars(E16_b_pur_grp(M450_E16_pur_grp))			& """" & vbcr
 Response.Write "  .txtPurGrpNm.value		= """ & ConvSPChars(E16_b_pur_grp(M450_E16_pur_grp_nm))			& """" & vbcr
  
 'Response.Write "  .txtXchRate.value		= """ & UNINumClientFormat(E14_m_pur_ord_hdr(M450_E14_xch_rt), ggExchRate.DecPoint, 0) & """" & vbcr
 Response.Write "     .hdnDiv.value         = """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_xch_rate_op))	& """" & vbcr
 Response.Write "  .txtIncoterms.value		= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_incoterms))		& """" & vbcr
 Response.Write "  .txtPayMethod.value		= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_pay_meth))		& """" & vbcr
 Response.Write "  .txtPayMethodNm.Value	= """ & ConvSPChars(E5_b_minor_paymeth_nm)						& """" & vbcr
 Response.Write "  .txtPayDur.value			= """ & UNINumClientFormat(E14_m_pur_ord_hdr(M450_E14_pay_dur), 0, 0) & """" & vbcr
 Response.Write "  .txtPurOrg.value			= """ & ConvSPChars(E17_b_pur_org(M450_E17_pur_org))			& """" & vbcr
 Response.Write "  .txtPurOrgNm.value		= """ & ConvSPChars(E17_b_pur_org(M450_E17_pur_org_nm))			& """" & vbcr
 Response.Write "  .txtPayeeCd.value		= """ & ConvSPChars(E4_b_biz_partner_mpa(B132_E4_bp_cd))		& """" & vbcr
 Response.Write "  .txtPayeeNm.value		= """ & ConvSPChars(E4_b_biz_partner_mpa(B132_E4_bp_nm))		& """" & vbcr
 Response.Write "  .txtBuildCd.value		= """ & ConvSPChars(E5_b_biz_partner_mbi(B132_E5_bp_cd))		& """" & vbcr
 Response.Write "  .txtBuildNm.value		= """ & ConvSPChars(E5_b_biz_partner_mbi(B132_E5_bp_nm))		& """" & vbcr
 Response.Write "  .txtBeneficiary.value	= """ & ConvSPChars(E15_b_biz_partner(M450_E15_bp_cd))			& """" & vbcr
 Response.Write "  .txtBeneficiaryNm.value	= """ & ConvSPChars(E15_b_biz_partner(M450_E15_bp_nm))			& """" & vbcr
 Response.Write "  .txtApplicant.value		= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_applicant))		& """" & vbcr
 Response.Write "  .txtApplicantNm.value	= """ & ConvSPChars(E12_b_biz_partner_applicant_nm)				& """" & vbcr
 Response.Write "  .txtIvType.value			= """ & ConvSPChars(E14_m_pur_ord_hdr(M450_E14_iv_type))		& """" & vbcr
 Response.Write "  .txtIvTypeNm.value		= """ & ConvSPChars(E1_m_iv_type)								& """" & vbcr
 Response.Write "  .hdnLoanflg.value		= ""N""" & vbcr
  
 Response.Write "  .txtBlIssueDt.text		= """ & Request("hdnBlIssueDt") & """" & vbcr
  
 Response.Write "if Trim(.txtPayeeCd.value) = """" then .txtPayeeCd.value = .txtBeneficiary.value"		& vbcr
 Response.Write "if Trim(.txtPayeeNm.value) = """" then .txtPayeeNm.value = .txtBeneficiaryNm.value"	& vbcr
 Response.Write "if Trim(.txtBuildCd.value) = """" then .txtBuildCd.value = .txtBeneficiary.value"		& vbcr
 Response.Write "if Trim(.txtBuildNm.value) = """" then .txtBuildNm.value = .txtBeneficiaryNm.value"	& vbcr
 Response.Write "  parent.GetPayDt()				" & vbcr
 Response.Write "  parent.ChangeCurOrDt()			" & vbcr
 Response.Write "  parent.dbRefQueryOK()			" & vbcr
 Response.Write "  parent.GetTaxBizArea("*")		" & vbcr
 Response.Write "  parent.CheckPrePayedAmtYN()		" & vbcr
 Response.Write " End With " & vbcr
 
 Response.Write "</Script> " & vbcr

Set M31119lc = Nothing 
Set B1H028de = Nothing 
%>
