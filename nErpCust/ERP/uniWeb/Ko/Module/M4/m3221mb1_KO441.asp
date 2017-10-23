<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3221mb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C Amend등록 Query Transaction 처리용 ASP							*
'*  7. Modified date(First) : 2000/04/03																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Sun-Jung Lee																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/03 : Coding Start												*
'********************************************************************************************************

Dim lgOpModeCRUD
Dim iPM4G219
Dim lgCurrency

On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
	
	Call HideStatusWnd
	
	lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
	
	Select Case lgOpModeCRUD
	    Case CStr(UID_M0001)                                                         '☜: Query	         
	         Call  SubBizQueryMulti()
	         
	    Case CStr(UID_M0002, UID_M0003)                                              '☜: Save,Update
	         Call SubBizSaveMulti()
	End Select
	

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
Dim strDt
Dim strDefDate

Dim I_M_AmendHdr

Dim E1_b_daily_exchange_rate	
Dim E2_b_minor_charge_nm
Dim E3_b_minor_credit_core_nm
Dim E4_b_minor_fund_type_nm
Dim E5_b_minor_origin_nm
Dim E6_b_minor_freight_nm
Dim E7_b_minor_bl_awb_nm
Dim E8_b_minor_paymeth_nm
Dim E9_b_minor_incoterms_nm
Dim E10_b_minor_lc_type_nm
Dim E11_b_minor_at_transport_nm
Dim E12_b_minor_at_loading_nm
Dim E13_b_minor_at_dischge_nm
Dim E14_b_minor_be_transport_nm
Dim E15_b_minor_be_loading_nm
Dim E16_b_minor_m_lc_amend_hdr
Dim E17_m_lc_amend_hdr
	Const M445_E17_lc_amd_no = 0
	Const M445_E17_lc_no = 1
	Const M445_E17_lc_doc_no = 2
	Const M445_E17_lc_amend_seq = 3
	Const M445_E17_adv_no = 4
	Const M445_E17_pre_adv_ref = 5
	Const M445_E17_open_dt = 6
	Const M445_E17_be_expiry_dt = 7
	Const M445_E17_at_expiry_dt = 8
	Const M445_E17_manufacturer = 9
	Const M445_E17_agent = 10
	Const M445_E17_amend_dt = 11
	Const M445_E17_amend_req_dt = 12
	Const M445_E17_currency = 13
	Const M445_E17_be_doc_amt = 14
	Const M445_E17_at_doc_amt = 15
	Const M445_E17_at_xch_rate = 16
	Const M445_E17_inc_amt = 17
	Const M445_E17_dec_amt = 18
	Const M445_E17_be_loc_amt = 19
	Const M445_E17_at_loc_amt = 20
	Const M445_E17_be_partial_ship = 21
	Const M445_E17_at_partial_ship = 22
	Const M445_E17_be_latest_ship_dt = 23
	Const M445_E17_at_latest_ship_dt = 24
	Const M445_E17_open_bank = 25
	Const M445_E17_insrt_user_id = 26
	Const M445_E17_insrt_dt = 27
	Const M445_E17_updt_user_id = 28
	Const M445_E17_updt_dt = 29
	Const M445_E17_be_xch_rate = 30
	Const M445_E17_ext1_amt = 31
	Const M445_E17_ext1_cd = 32
	Const M445_E17_remark = 33
	Const M445_E17_lc_kind = 34
	Const M445_E17_remark2 = 35
	Const M445_E17_be_transhipment = 36
	Const M445_E17_at_transhipment = 37
	Const M445_E17_be_transfer = 38
	Const M445_E17_at_transfer = 39
	Const M445_E17_be_loading_port = 40
	Const M445_E17_at_loading_port = 41
	Const M445_E17_be_dischge_port = 42
	Const M445_E17_at_dischge_port = 43
	Const M445_E17_be_transport = 44
	Const M445_E17_at_transport = 45
	Const M445_E17_biz_area = 46
	Const M445_E17_charge_flg = 47
	Const M445_E17_ext1_qty = 48
	Const M445_E17_ext2_qty = 49
	Const M445_E17_ext3_qty = 50
	Const M445_E17_ext2_amt = 51
	Const M445_E17_ext3_amt = 52
	Const M445_E17_ext2_cd = 53
	Const M445_E17_ext3_cd = 54
	Const M445_E17_ext1_rt = 55
	Const M445_E17_ext2_rt = 56
	Const M445_E17_ext3_rt = 57
	Const M445_E17_ext1_dt = 58
	Const M445_E17_ext2_dt = 59
	Const M445_E17_ext3_dt = 60
Dim E18_m_lc_hdr
	'조회 성능개선(2003.05.23)
	Const M445_E18_lc_no = 0
	Const M445_E18_po_no = 1
	Const M445_E18_pay_method = 2
	Const M445_E18_incoterms = 3
	Const M445_E18_doc_amt = 4
	Const M445_E18_xch_rate_op = 5
	Const M445_E18_tot_amend_amt = 6

Dim E19_b_bank_issue_bank
	Const M445_E19_bank_cd = 0
	Const M445_E19_bank_nm = 1
Dim E20_b_bank_advise_bank
	Const M445_E20_bank_cd = 0
	Const M445_E20_bank_nm = 1
Dim E21_b_bank_renego_bank
Dim E22_b_bank_pay_bank
Dim E23_b_bank_confirm_bank
Dim E24_b_pur_org
	Const M445_E24_pur_org = 0
    Const M445_E24_pur_org_nm = 1
Dim E25_b_pur_grp
	Const M445_E25_pur_grp = 0
    Const M445_E25_pur_grp_nm = 1
Dim E26_b_biz_partner_beneficiary
	Const M445_E26_bp_cd = 0
    Const M445_E26_bp_nm = 1
Dim E27_b_biz_partner_applicant
	Const M445_E27_bp_cd = 0
    Const M445_E27_bp_nm = 1
Dim E28_b_biz_partner_agent
Dim E29_b_biz_partner_manufacturer
Dim E30_b_biz_partner_notify_party
    
On Error Resume Next                                                            '☜: Protect system from crashing
Err.Clear      
	
	Set iPM4G219 = Server.CreateObject("PM4G219.cMLookupLcAmendHdrS")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If
	
	Redim I_M_AmendHdr(1)
	I_M_AmendHdr(0) =  Request("txtLCAmdNo")
	I_M_AmendHdr(1) =  "M"
	
	Call iPM4G219.M_LOOKUP_LC_AMEND_HDR_SVR(gStrGlobalCollection, I_M_AmendHdr, E1_b_daily_exchange_rate, _
					E2_b_minor_charge_nm, E3_b_minor_credit_core_nm, E4_b_minor_fund_type_nm, E5_b_minor_origin_nm, E6_b_minor_freight_nm, _
					E7_b_minor_bl_awb_nm, E8_b_minor_paymeth_nm, E9_b_minor_incoterms_nm, E10_b_minor_lc_type_nm, E11_b_minor_at_transport_nm, _
					E12_b_minor_at_loading_nm, E13_b_minor_at_dischge_nm, E14_b_minor_be_transport_nm, E15_b_minor_be_loading_nm, E16_b_minor_m_lc_amend_hdr, _
					E17_m_lc_amend_hdr, E18_m_lc_hdr, E19_b_bank_issue_bank, E20_b_bank_advise_bank, E21_b_bank_renego_bank, E22_b_bank_pay_bank, E23_b_bank_confirm_bank, _
					E24_b_pur_org, E25_b_pur_grp, E26_b_biz_partner_beneficiary, E27_b_biz_partner_applicant, E28_b_biz_partner_agent, _
					E29_b_biz_partner_manufacturer, E30_b_biz_partner_notify_party)
	
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 	
		Set iPM4G219 = Nothing
		Exit Sub
	End If
	
	Set iPM4G219 = Nothing
	
	strDefDate	= UNIDateClientFormat("1900-01-01")
	lgCurrency	= ConvSPChars(E17_m_lc_amend_hdr(M445_E17_currency))
	
	Response.Write "<Script Language=vbscript>" & vbCr	    
	Response.Write " With parent "	& vbCr
	Response.Write " .frm1.txtCurrency.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_currency)) & """" & vbCr
	Response.Write " .frm1.txtLCAmdNo1.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_amd_no)) & """" & vbCr
	Response.Write " .frm1.txtHLCAmdNo.value       	= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_amd_no)) & """" & vbCr
	Response.Write " .frm1.txtLCDocNo.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_doc_no)) & """" & vbCr
	Response.Write " .frm1.txtLCAmendSeq.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_amend_seq)) & """" & vbCr
	Response.Write " .frm1.txtApplicant.value		= """ & ConvSPChars(E27_b_biz_partner_applicant(M445_E27_bp_cd)) & """" & vbCr
	Response.Write " .frm1.txtApplicantNm.value		= """ & ConvSPChars(E27_b_biz_partner_applicant(M445_E27_bp_nm)) & """" & vbCr

	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_amend_req_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtAmendReqDt.text = """ & strDt & """" & vbcr
	End If
			
	
	strDt =  UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_amend_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtAmendDt.text = """ & strDt & """" & vbcr
	End If
	
	If CDbl(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_inc_amt))) <> 0 then
		Response.Write " .frm1.rdoAtDocAmt1.Checked		= True" & vbCr
		Response.Write " .frm1.txtAmendAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E17_m_lc_amend_hdr(M445_E17_inc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	ElseIf CDbl(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_dec_amt))) <> 0 Then
		Response.Write " .frm1.rdoAtDocAmt2.Checked		= True" & vbCr
		Response.Write " .frm1.txtAmendAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E17_m_lc_amend_hdr(M445_E17_dec_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	End If		

	Response.Write " .frm1.txtAtDocAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E17_m_lc_amend_hdr(M445_E17_at_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_at_expiry_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtAtExpiryDt.text		= """ & UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_at_expiry_dt)) & """" & vbCr
	End If
	
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_expiry_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtBeExpiryDt.text		= """ & UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_expiry_dt)) & """" & vbCr
		Response.Write " .frm1.txtHExpiryDt.value		= """ & UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_expiry_dt)) & """" & vbCr
	End If

	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_latest_ship_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtBeLatestShipDt.text		= """ & UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_latest_ship_dt)) & """" & vbCr
		Response.Write " .frm1.txtHLatestShipDt.Value		= """ & UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_latest_ship_dt)) & """" & vbCr
	End If
			
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_at_latest_ship_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtAtLatestShipDt.text		= """ & UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_at_latest_ship_dt)) & """" & vbCr
	End If	
	
	If UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transhipment)))) = "Y" Then
		Response.Write " .frm1.rdoAtTranshipment1.Checked		= True" & vbCr
	ElseIf UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transhipment)))) = "N" Then
		Response.Write " .frm1.rdoAtTranshipment2.Checked		= True" & vbCr
	End If

	Response.Write " .frm1.txtHTranshipment.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transhipment)) & """" & vbCr
	Response.Write " .frm1.txtBeTranshipment.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_transhipment)) & """" & vbCr
	
	If UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_partial_ship)))) = "Y" Then
		Response.Write " .frm1.rdoAtPartialShip1.Checked		= True" & vbCr
	ElseIf UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_partial_ship)))) = "N" Then
		Response.Write " .frm1.rdoAtPartialShip2.Checked		= True" & vbCr
	End If

	Response.Write " .frm1.txtHPartialShip.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_partial_ship)) & """" & vbCr
	Response.Write " .frm1.txtBePartialShip.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_partial_ship)) & """" & vbCr
			
	If UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transfer)))) = "Y" Then
		Response.Write " .frm1.rdoAtTransfer1.Checked		= True" & vbCr
	ElseIf UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transfer)))) = "N" Then
		Response.Write " .frm1.rdoAtTransfer2.Checked		= True" & vbCr
	End If
		
	Response.Write " .frm1.txtHTransfer.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transfer)) & """" & vbCr
	Response.Write " .frm1.txtBeTransfer.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_transfer)) & """" & vbCr
	Response.Write " .frm1.txtXchRt.text		= """ & UNINumClientFormat(E17_m_lc_amend_hdr(M445_E17_at_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	
	Response.Write " .frm1.txtHTransport.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transfer)) & """" & vbCr
	Response.Write " .frm1.txtAtTransport.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_transport)) & """" & vbCr
	Response.Write " .frm1.txtBeTransport.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_transport)) & """" & vbCr
	Response.Write " .frm1.txtAtTransportNm.value		= """ & ConvSPChars(E11_b_minor_at_transport_nm) & """" & vbCr
	Response.Write " .frm1.txtBeTransportNm.value		= """ & ConvSPChars(E14_b_minor_be_transport_nm) & """" & vbCr
		
	Response.Write " .frm1.txtHLoadingPort.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_loading_port)) & """" & vbCr
	Response.Write " .frm1.txtAtLoadingPort.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_loading_port)) & """" & vbCr
	Response.Write " .frm1.txtBeLoadingPort.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_loading_port)) & """" & vbCr
	Response.Write " .frm1.txtAtLoadingPortNm.value		= """ & ConvSPChars(E12_b_minor_at_loading_nm) & """" & vbCr
	Response.Write " .frm1.txtBeLoadingPortNm.value		= """ & ConvSPChars(E15_b_minor_be_loading_nm) & """" & vbCr
		
	Response.Write " .frm1.txtHDischgePort.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_dischge_port)) & """" & vbCr
	Response.Write " .frm1.txtAtDischgePort.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_dischge_port)) & """" & vbCr
	Response.Write " .frm1.txtBeDischgePort.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_dischge_port)) & """" & vbCr			
	Response.Write " .frm1.txtAtDischgePortNm.value		= """ & ConvSPChars(E13_b_minor_at_dischge_nm) & """" & vbCr			
	Response.Write " .frm1.txtBeDischgePortNm.value		= """ & ConvSPChars(E16_b_minor_m_lc_amend_hdr) & """" & vbCr			
	
	Response.Write " .frm1.txtRemark.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_remark)) & """" & vbCr
	Response.Write " .frm1.txtCurrency.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_currency)) & """" & vbCr

	Response.Write " .frm1.txtBeDocAmt.value		= """ & UNIConvNumDBToCompanyByCurrency(E17_m_lc_amend_hdr(M445_E17_be_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_open_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtOpenDt.text = """ & strDt & """" & vbcr
	End If
	
	Response.Write " .frm1.txtAdvBank.value		= """ & ConvSPChars(E20_b_bank_advise_bank(M445_E20_bank_cd)) & """" & vbCr
	Response.Write " .frm1.txtAdvBankNm.value		= """ & ConvSPChars(E20_b_bank_advise_bank(M445_E20_bank_nm)) & """" & vbCr
	Response.Write " .frm1.txtOpenBank.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_open_bank)) & """" & vbCr
	Response.Write " .frm1.txtOpenBankNm.value		= """ & ConvSPChars(E19_b_bank_issue_bank(M445_E19_bank_nm)) & """" & vbCr
	Response.Write " .frm1.txtBeneficiary.value		= """ & ConvSPChars(E26_b_biz_partner_beneficiary(M445_E26_bp_cd)) & """" & vbCr
	Response.Write " .frm1.txtBeneficiaryNm.value		= """ & ConvSPChars(E26_b_biz_partner_beneficiary(M445_E26_bp_nm)) & """" & vbCr
	Response.Write " .frm1.txtPurGrp.value		= """ & ConvSPChars(E25_b_pur_grp(M445_E25_pur_grp)) & """" & vbCr
	Response.Write " .frm1.txtPurGrpNm.value		= """ & ConvSPChars(E25_b_pur_grp(M445_E25_pur_grp_nm)) & """" & vbCr
	Response.Write " .frm1.txtPurOrg.value		= """ & ConvSPChars(E24_b_pur_org(M445_E24_pur_org)) & """" & vbCr
	Response.Write " .frm1.txtPurOrgNm.value		= """ & ConvSPChars(E24_b_pur_org(M445_E24_pur_org_nm)) & """" & vbCr
	
	Response.Write " .frm1.txtLCNo.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_no)) & """" & vbCr
	Response.Write " .frm1.txtPONo.value		= """ & ConvSPChars(E18_m_lc_hdr(M445_E18_po_no)) & """" & vbCr
	Response.Write " .frm1.txtBeTransport.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_transport)) & """" & vbCr
	Response.Write " .frm1.txtIncAmt.value		= """ & UNINumClientFormat(E17_m_lc_amend_hdr(M445_E17_inc_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	Response.Write " .frm1.txtDecAmt.value		= """ & UNINumClientFormat(E17_m_lc_amend_hdr(M445_E17_dec_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	
	Response.Write " .DbQueryOk "           & vbCr
	Response.Write " End With"           & vbCr
    Response.Write "</Script>" & vbCr
End Sub

%>
