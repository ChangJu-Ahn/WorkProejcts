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
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5211mb6.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : B/L Header Insert를 위한 L/C Header Data Query Transaction 처리용 ASP     *
'*  7. Modified date(First) : 2000/05/02																*
'*  8. Modified date(Last)  : 2003/05/26																*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************

On Error Resume Next

Call HideStatusWnd

Dim iPM4G119
Dim B1H028de
Dim B17014
Dim E1_ief_supplied_loan_flg 'not used
 
 Dim E2_b_minor_charge
    Const E2_minor_cd = 0
    Const E2_minor_nm = 1
    Const E2_minor_type = 2

 Dim E3_b_daily_exchange_rate 'not used

 Dim E4_b_minor_credit_core
    Const E4_minor_cd = 0
    Const E4_minor_nm = 1
    Const E4_minor_type = 2

 Dim E5_b_minor_fund_type
    Const E5_minor_cd = 0
    Const E5_minor_nm = 1
    Const E5_minor_type = 2
    
 Dim E6_b_minor_origin
    Const E6_minor_cd = 0
    Const E6_minor_nm = 1
    Const E6_minor_type = 2
    
 Dim E7_b_minor_freight
    Const E7_minor_cd = 0
    Const E7_minor_nm = 1
    Const E7_minor_type = 2
    
 Dim E8_b_minor_bl_awb_flg
    Const E8_minor_cd = 0
    Const E8_minor_nm = 1
    Const E8_minor_type = 2

 Dim E9_b_minor_dischge_port
    Const E9_minor_cd = 0
    Const E9_minor_nm = 1
    Const E9_minor_type = 2
    
 Dim E10_b_minor_loading_port
    Const E10_minor_cd = 0
    Const E10_minor_nm = 1
    Const E10_minor_type = 2
    
 Dim E11_b_minor_transport
    Const E11_minor_cd = 0
    Const E11_minor_nm = 1
    Const E11_minor_type = 2
    
 Dim E12_b_minor_pay_method
    Const E12_minor_cd = 0
    Const E12_minor_nm = 1
    Const E12_minor_type = 2
    
 Dim E13_b_minor_incoterms
    Const E13_minor_cd = 0
    Const E13_minor_nm = 1
    Const E13_minor_type = 2

 Dim E14_b_minor_delivery_plce
    Const E14_minor_cd = 0
    Const E14_minor_nm = 1
    Const E14_minor_type = 2
    
 Dim E15_b_minor_lc_type
    Const E15_minor_cd = 0
    Const E15_minor_nm = 1
    Const E15_minor_type = 2
    
 Dim E16_b_minor_o_lc_type
    Const E16_minor_cd = 0
    Const E16_minor_nm = 1
    Const E16_minor_type = 2
    
 Dim E17_b_minor_o_lc_kind
    Const E17_minor_cd = 0
    Const E17_minor_nm = 1
    Const E17_minor_type = 2
    
 Dim E18_m_lc_hdr
    Const E18_lc_no1 = 0
    Const E18_lc_doc_no1 = 1
    Const E18_lc_amend_seq1 = 2
    Const E18_po_no1 = 3
    Const E18_adv_no1 = 4
    Const E18_pre_adv_ref1 = 5
    Const E18_req_dt1 = 6
    Const E18_adv_dt1 = 7
    Const E18_open_dt1 = 8
    Const E18_expiry_dt1 = 9
    Const E18_amend_dt1 = 10
    Const E18_manufacturer1 = 11
    Const E18_agent1 = 12
    Const E18_currency1 = 13
    Const E18_doc_amt1 = 14
    Const E18_xch_rate1 = 15
    Const E18_xch_rate_op1 = 16
    Const E18_loc_amt1 = 17
    Const E18_bank_txt1 = 18
    Const E18_incoterms1 = 19
    Const E18_pay_method1 = 20
    Const E18_pay_terms_txt1 = 21
    Const E18_partial_ship1 = 22
    Const E18_latest_ship_dt1 = 23
    Const E18_shipment1 = 24
    Const E18_doc11 = 25
    Const E18_doc21 = 26
    Const E18_doc31 = 27
    Const E18_doc41 = 28
    Const E18_doc51 = 29
    Const E18_file_dt1 = 30
    Const E18_file_dt_txt1 = 31
    Const E18_insrt_user_id1 = 32
    Const E18_insrt_dt1 = 33
    Const E18_updt_user_id1 = 34
    Const E18_updt_dt1 = 35
    Const E18_ext1_qty1 = 36
    Const E18_ext1_amt1 = 37
    Const E18_ext1_cd1 = 38
    Const E18_remark1 = 39
    Const E18_lc_kind1 = 40
    Const E18_lc_type1 = 41
    Const E18_transhipment1 = 42
    Const E18_transfer1 = 43
    Const E18_delivery_plce1 = 44
    Const E18_tolerance1 = 45
    Const E18_loading_port1 = 46
    Const E18_dischge_port1 = 47
    Const E18_transport_comp1 = 48
    Const E18_origin1 = 49
    Const E18_origin_cntry1 = 50
    Const E18_charge_txt1 = 51
    Const E18_charge_cd1 = 52
    Const E18_credit_core1 = 53
    Const E18_lc_remn_doc_amt1 = 54
    Const E18_lc_remn_loc_amt1 = 55
    Const E18_fund_type1 = 56
    Const E18_lmt_xch_rate1 = 57
    Const E18_lmt_amt1 = 58
    Const E18_inv_cnt1 = 59
    Const E18_bl_awb_flg1 = 60
    Const E18_freight1 = 61
    Const E18_notify_party1 = 62
    Const E18_consignee1 = 63
    Const E18_insur_policy1 = 64
    Const E18_pack_list1 = 65
    Const E18_cert_origin_flg1 = 66
    Const E18_transport1 = 67
    Const E18_l_lc_type1 = 68
    Const E18_o_lc_kind1 = 69
    Const E18_o_lc_doc_no1 = 70
    Const E18_o_lc_amend_seq1 = 71
    Const E18_o_lc_no1 = 72
    Const E18_o_lc_type1 = 73
    Const E18_o_lc_open_dt1 = 74
    Const E18_o_lc_expiry_dt1 = 75
    Const E18_o_lc_loc_amt1 = 76
    Const E18_biz_area1 = 77
    Const E18_pay_dur1 = 78
    Const E18_charge_flg1 = 79
    Const E18_ext2_qty1 = 80
    Const E18_ext3_qty1 = 81
    Const E18_ext2_amt1 = 82
    Const E18_ext3_amt1 = 83
    Const E18_ext2_cd1 = 84
    Const E18_ext3_cd1 = 85
    Const E18_ext1_rt1 = 86
    Const E18_ext2_rt1 = 87
    Const E18_ext3_rt1 = 88
    Const E18_ext1_dt1 = 89
    Const E18_ext2_dt1 = 90
    Const E18_ext3_dt1 = 91
    Const E18_pur_org1 = 92
    Const E18_pur_grp1 = 93
    Const E18_applicant1 = 94
    Const E18_beneficiary1 = 95
    Const E18_open_bank1 = 96
    Const E18_adv_bank1 = 97
    Const E18_renego_bank1 = 98
    Const E18_pay_bank1 = 99
    Const E18_return_bank1 = 100
 
 Dim E19_b_bank_open_bank
 Const E19_bank_cd=0
 Const E19_bank_nm=1
  
 Dim E20_b_bank_adv_bank
 Const E20_bank_cd=0
 Const E20_bank_nm=1
 
 Dim E21_b_bank_pay_bank
 Const E21_bank_cd=0
 Const E21_bank_nm=1
 
 Dim E22_b_bank_renego_bank
 Const E22_bank_cd=0
 Const E22_bank_nm=1
 
 Dim E23_b_bank_return_bank
 Const E23_bank_cd=0
 Const E23_bank_nm=1
 
 Dim E24_b_pur_org
 Const E24_pur_org=0
 Const E24_pur_org_nm=1

 Dim E25_b_pur_grp
 Const E25_pur_grp=0
 Const E25_pur_grp_nm=1

 Dim E26_b_biz_partner_beneficiary
 Const E26_bp_cd=0
 Const E26_bp_nm=1

 Dim E27_b_biz_partner_applicant
 Const E27_bp_cd=0
 Const E27_bp_nm=1
 
 Dim E28_b_biz_partner_agent
 Const E28_bp_cd=0
 Const E28_bp_nm=1
 
 Dim E29_b_biz_partner_manufacturer
 Const E29_bp_cd=0
 Const E29_bp_nm=1
 
 Dim E30_b_biz_partner_notify_party
 Const E30_bp_cd=0
 Const E30_bp_nm=1

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
 
 Dim M446_I1_lc_no
 Const M446_I1_lc_no1=0
 Const M446_I1_lc_kind1=1
 
 Redim M446_I1_lc_no(1)
 M446_I1_lc_no(M446_I1_lc_no1) = Request("txtLCNo")
 M446_I1_lc_no(M446_I1_lc_kind1) = "M"

 Set iPM4G119 = Server.CreateObject("PM4G119.cMLookupLcHdrS")

 If CheckSYSTEMError(Err,True) = True Then
	Set iPM4G119 = Nothing
	Response.End
 End If

 Call iPM4G119.M_LOOKUP_LC_HDR_SVR(gStrGlobalCollection, _
           M446_I1_lc_no    , E1_ief_supplied_loan_flg , _
           E2_b_minor_charge    , E3_b_daily_exchange_rate , _
           E4_b_minor_credit_core   , E5_b_minor_fund_type , _
           E6_b_minor_origin    , E7_b_minor_freight , _
           E8_b_minor_bl_awb_flg   , E9_b_minor_dischge_port , _
           E10_b_minor_loading_port  , E11_b_minor_transport , _
           E12_b_minor_pay_method   , E13_b_minor_incoterms , _
           E14_b_minor_delivery_plce  , E15_b_minor_lc_type , _
           E16_b_minor_o_lc_type   , E17_b_minor_o_lc_kind , _
           E18_m_lc_hdr     , E19_b_bank_open_bank , _
           E20_b_bank_adv_bank    , E21_b_bank_pay_bank , _
           E22_b_bank_renego_bank   , E23_b_bank_return_bank , _
           E24_b_pur_org     , E25_b_pur_grp , _
           E26_b_biz_partner_beneficiary, E27_b_biz_partner_applicant , _
           E28_b_biz_partner_agent   , E29_b_biz_partner_manufacturer , _
           E30_b_biz_partner_notify_party )


 If CheckSYSTEMError(Err,True) = True Then
	Set iPM4G119 = Nothing
	Response.End
 End If

 Set iPM4G119 = Nothing

 '------------------------
 '지급처와 발행처를 Lookup
 '------------------------
 Set B1H028de = Server.CreateObject("PB5GS45.cBListDftBpFtnSvr")

 If CheckSYSTEMError(Err,True) = True Then
	Set B1H028de = Nothing
	Response.End
 End If

  Call B1H028de.B_LIST_DEFAULT_BP_FTN_SVR(gStrGlobalCollection, _
                     E26_b_biz_partner_beneficiary(E26_bp_cd), _
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
Const strDefDate = "1899-12-30"

Response.Write "<Script Language=VBScript>"    & vbcr
Response.Write " Dim strDefDate"      & vbcr
Response.Write " strDefDate = """ & UNIDateClientFormat(strDefDate) & """" & vbcr
Response.Write " With parent.frm1"      & vbcr
Response.Write "  '##### Rounding Logic #####"   & vbcr
Response.Write "  '항상 거래화폐가 우선"    & vbcr
Response.Write "  .txtCurrency.Value = """ & ConvSPChars(E18_m_lc_hdr(E18_currency1)) & """" & vbcr
Response.Write "  '##########################" & vbcr
Response.Write "  .txtLCDocNo.value   = """ & ConvSPChars(E18_m_lc_hdr(E18_lc_doc_no1)) & """" & vbcr
Response.Write "  .txtLCNo.value    = """ & ConvSPChars(Request("txtLCNo")) & """" & vbcr

Response.Write "  if """ & ConvSPChars(Trim(E18_m_lc_hdr(E18_lc_doc_no1))) & """ <> """" then" & vbcr
Response.Write "   .chkLcNoCnt.checked = true" & vbcr
Response.Write "  else" & vbcr
Response.Write "   .chkLcNoCnt.checked = false" & vbcr
Response.Write "  End if" & vbcr

Response.Write "  .chkPoNoCnt.checked   = false" & vbcr
Response.Write "  .txtTransport.Value   = """ & ConvSPChars(E18_m_lc_hdr(E18_transport1)) & """" & vbcr
Response.Write "  .txtTransportNm.Value   = """ & ConvSPChars(E11_b_minor_transport(E11_minor_nm)) & """" & vbcr
Response.Write "  .txtDischgePort.Value   = """ & ConvSPChars(E18_m_lc_hdr(E18_dischge_port1)) & """" & vbcr
Response.Write "  .txtDischgePortNm.Value  = """ & ConvSPChars(E9_b_minor_dischge_port(E9_minor_nm)) & """" & vbcr
'Response.Write "  '.txtXchRate.Value    = """ & ConvSPChars(E18_m_lc_hdr(E18_xch_rate1)) & """" & vbcr
Response.Write "     .hdnDiv.value           = """ & ConvSPChars(E18_m_lc_hdr(E18_xch_rate_op1)) & """" & vbcr
Response.Write "  .txtPayMethod.value   = """ & ConvSPChars(E18_m_lc_hdr(E18_pay_method1)) & """" & vbcr
Response.Write "  .txtPayMethodNm.value   = """ & ConvSPChars(E12_b_minor_pay_method(E12_minor_nm)) & """" & vbcr
Response.Write "  .txtPayDur.text    = """ & ConvSPChars(E18_m_lc_hdr(E18_pay_dur1)) & """" & vbcr
Response.Write "  .txtIncoterms.value   = """ & ConvSPChars(E18_m_lc_hdr(E18_incoterms1)) & """" & vbcr
Response.Write "  .txtPayeeCd.value    = """ & ConvSPChars(E4_b_biz_partner_mpa(B132_E4_bp_cd)) & """" & vbcr
Response.Write "  .txtPayeeNm.value    = """ & ConvSPChars(E4_b_biz_partner_mpa(B132_E4_bp_nm)) & """" & vbcr
Response.Write "  .txtBuildCd.value    = """ & ConvSPChars(E5_b_biz_partner_mbi(B132_E5_bp_cd)) & """" & vbcr
Response.Write "  .txtBuildNm.value    = """ & ConvSPChars(E5_b_biz_partner_mbi(B132_E5_bp_nm)) & """" & vbcr
Response.Write "  .txtBeneficiary.value   = """ & ConvSPChars(E26_b_biz_partner_beneficiary(E26_bp_cd)) & """" & vbcr
Response.Write "  .txtBeneficiaryNm.value  = """ & ConvSPChars(E26_b_biz_partner_beneficiary(E26_bp_nm)) & """" & vbcr
Response.Write "  .txtFreight.value    = """ & ConvSPChars(E18_m_lc_hdr(E18_freight1)) & """" & vbcr
Response.Write "  .txtFreightNm.value   = """ & ConvSPChars(E7_b_minor_freight(E7_minor_nm)) & """" & vbcr
Response.Write "  .txtDeliveryPlce.value   = """ & ConvSPChars(E18_m_lc_hdr(E18_delivery_plce1)) & """" & vbcr
Response.Write "  .txtDeliveryPlceNm.value  = """ & ConvSPChars(E14_b_minor_delivery_plce(E14_minor_nm)) & """" & vbcr
Response.Write "  .txtLoadingPort.value   = """ & ConvSPChars(E18_m_lc_hdr(E18_loading_port1)) & """" & vbcr
Response.Write "  .txtLoadingPortNm.Value  = """ & ConvSPChars(E10_b_minor_loading_port(E10_minor_nm)) & """" & vbcr
Response.Write "  .txtOrigin.value    = """ & ConvSPChars(E18_m_lc_hdr(E18_origin1)) & """" & vbcr
Response.Write "  .txtOriginNm.Value    = """ & ConvSPChars(E6_b_minor_origin(E6_minor_nm)) & """" & vbcr
Response.Write "  .txtOriginCntry.value   = """ & ConvSPChars(E18_m_lc_hdr(E18_origin_cntry1)) & """" & vbcr
Response.Write "  .txtAgent.value    = """ & ConvSPChars(E18_m_lc_hdr(E18_agent1)) & """" & vbcr
Response.Write "  .txtAgentNm.value    = """ & ConvSPChars(E28_b_biz_partner_agent(E28_bp_nm)) & """" & vbcr
Response.Write "  .txtManufacturer.value   = """ & ConvSPChars(E18_m_lc_hdr(E18_manufacturer1)) & """" & vbcr
Response.Write "  .txtManufacturerNm.value  = """ & ConvSPChars(E29_b_biz_partner_manufacturer(E29_bp_nm)) & """" & vbcr
Response.Write "  .txtPurGrp.value    = """ & ConvSPChars(E25_b_pur_grp(E25_pur_grp)) & """" & vbcr
Response.Write "  .txtPurGrpNm.value    = """ & ConvSPChars(E25_b_pur_grp(E25_pur_grp_nm)) & """" & vbcr
Response.Write "  .txtPurOrg.value    = """ & ConvSPChars(E24_b_pur_org(E24_pur_org)) & """" & vbcr
Response.Write "  .txtPurOrgNm.value    = """ & ConvSPChars(E24_b_pur_org(E24_pur_org_nm)) & """" & vbcr
Response.Write "  .txtApplicant.value   = """ & ConvSPChars(E27_b_biz_partner_applicant(E27_bp_cd)) & """" & vbcr
Response.Write "  .txtApplicantNm.value   = """ & ConvSPChars(E27_b_biz_partner_applicant(E27_bp_nm)) & """" & vbcr
Response.Write "  .hdnLoanflg.value    = """ & ConvSPChars(E1_ief_supplied_loan_flg) & """" & vbcr
  
Response.Write "  .txtBlIssueDt.text    = """ & Request("hdnBlIssueDt")  & """" & vbcr
  
Response.Write "  if Trim(.txtPayeeCd.value) = """" then .txtPayeeCd.value = .txtBeneficiary.value" & vbcr
Response.Write "  if Trim(.txtPayeeNm.value) = """" then .txtPayeeNm.value = .txtBeneficiaryNm.value" & vbcr
Response.Write "  if Trim(.txtBuildCd.value) = """" then .txtBuildCd.value = .txtBeneficiary.value" & vbcr
Response.Write "  if Trim(.txtBuildNm.value) = """" then .txtBuildNm.value = .txtBeneficiaryNm.value" & vbcr
Response.Write "  parent.GetPayDt()				"   & vbcr             '지불예정일 
Response.Write "  parent.ChangeCurOrDt()		"   & vbcr        '환율 
Response.Write "  parent.dbRefQueryOK()			"   & vbcr
Response.Write "  parent.GetTaxBizArea("*")		"   & vbcr
Response.Write "  parent.CheckPrePayedAmtYN()	"   & vbcr
'Response.Write "  parent.setLoan()				"   & vbcr
  
Response.Write " End With" & vbcr
 
Response.Write "</Script>" & vbcr
 Set iPM4G119 = Nothing              '☜: Unload Comproxy

 Response.End                '☜: Process End
%> 

