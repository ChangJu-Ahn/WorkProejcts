<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
Call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
     
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear  
    
    Const Command = "Delete"
    Dim PM4G211
    Dim strConvDt

    Dim I1_b_biz_partner_bp_cd 
    Dim I2_b_biz_partner_bp_cd 
    Dim I3_m_lc_amend_hdr 
    Const M435_I3_lc_amd_no = 0
    Const M435_I3_lc_no = 1
    Const M435_I3_lc_doc_no = 2
    Const M435_I3_lc_amend_seq = 3
    Dim I4_s_wks_user 
    Dim I5_b_pur_grp 
	Dim I6_m_lc_amend_hdr

    redim I3_m_lc_amend_hdr(1)
    
	If Trim(""&Request("txtLCAmdNo")) = "" Then										'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("229909", vbOKOnly, "", "", I_MKSCRIPT)           
		Exit Sub 
	End If
					
	Set PM4G211 = Server.CreateObject("PM4G211.cMMaintLcAmendHdrS")
	      
	If CheckSYSTEMError(Err,True) = True Then
		Set PM4G211 = Nothing
		Exit Sub
	End If
		
    I3_m_lc_amend_hdr(M435_I3_lc_amd_no)  = Request("txtLCAmdNo")
	    	       
	CALL PM4G211.M_MAINT_LC_AMEND_HDR_SVR(gStrGlobalCollection , Command, cstr(I1_b_biz_partner_bp_cd), _
	             cstr(I2_b_biz_partner_bp_cd),I3_m_lc_amend_hdr,I4_s_wks_user,I5_b_pur_grp,I6_m_lc_amend_hdr) 
		
	if CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set PM4G211 = Nothing												'☜: ComProxy Unload
			
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
 	end if

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " Parent.DBDeleteOK "           & vbCr
	Response.Write "</Script>"                  & vbCr 
End Sub

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
	Dim I6_m_lc_amend_hdr    
	
	Dim lgCurrency
	Dim  iPM4G219 
	Dim  lgStrPrevKey

	Redim I_M_AmendHdr(1)

	On Error Resume Next                                                            '☜: Protect system from crashing
	Err.Clear      
	
	lgStrPrevKey = "" & Request("lgStrPrevKey")
	
  	Set iPM4G219 = Server.CreateObject("PM4G219.cMLookupLcAmendHdrS")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If
	
	I_M_AmendHdr(0) =  Request("txtLCAmdNo")
	I_M_AmendHdr(1) =  "L"
		
	Call iPM4G219.M_LOOKUP_LC_AMEND_HDR_SVR (gStrGlobalCollection, I_M_AmendHdr, E1_b_daily_exchange_rate, _
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
	'@수정(2003.03.12)
	Response.Write "    .CurFormatNumericOCX " & vbCr
	Response.Write " .frm1.txtLCAmdNo1.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_amd_no)) & """" & vbCr
	Response.Write " .frm1.txtLCDocNo.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_doc_no)) & """" & vbCr
	Response.Write " .frm1.txtLCAmendSeq.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_amend_seq)) & """" & vbCr
	Response.Write " .frm1.txtBeneficiary.value 		= """ & ConvSPChars(E26_b_biz_partner_beneficiary(M445_E26_bp_cd)) & """" & vbCr
	Response.Write " .frm1.txtBeneficiaryNm.value		= """ & ConvSPChars(E26_b_biz_partner_beneficiary(M445_E26_bp_nm)) & """" & vbCr
  
	If CDbl(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_inc_amt))) <> 0 then
		Response.Write " .frm1.rdoAtDocAmt1.Checked		= True" & vbCr
		Response.Write " .frm1.txtAmendAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E17_m_lc_amend_hdr(M445_E17_inc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	ElseIf CDbl(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_dec_amt))) <> 0 Then
		Response.Write " .frm1.rdoAtDocAmt2.Checked		= True" & vbCr
		Response.Write " .frm1.txtAmendAmt.text		= """ & UNINumClientFormat(E17_m_lc_amend_hdr(M445_E17_dec_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	End If			 
		
	Response.Write " .frm1.txtAtXchRate.text  =  """ & UNINumClientFormat(E17_m_lc_amend_hdr(M445_E17_at_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write " .frm1.txtAtDocAmt.text   =  """ & UNIConvNumDBToCompanyByCurrency(E17_m_lc_amend_hdr(M445_E17_at_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write " .frm1.txtAtLocAmt.text   =  """ & UNINumClientFormat(E17_m_lc_amend_hdr(M445_E17_at_loc_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	Response.Write " .frm1.txtBeDocAmt.text   =  """ & UNIConvNumDBToCompanyByCurrency(E17_m_lc_amend_hdr(M445_E17_be_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write " .frm1.txtBeLocAmt.text   =  """ & UNINumClientFormat(E17_m_lc_amend_hdr(M445_E17_be_loc_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	       
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_amend_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtAmendDt.text		= """ & strDt & """" & vbCr			
	End If	
	        
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_amend_req_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtAmendReqDt.text		= """ & strDt & """" & vbCr		
	End If	
	         
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_at_expiry_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtAtExpireDt.text		= """ & strDt & """" & vbCr
		Response.Write " .frm1.txtHExpiryDt.value		= """ & strDt & """" & vbCr
	End If	
		
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_expiry_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtBeExpireDt.text = """ & strDt & """" & vbcr		 
	End If	
			
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_at_latest_ship_dt))					
	If strDt <> strDefDate Then 
	 	Response.Write " .frm1.txtAtLatestShipDt.text		= """ & strDt & """" & vbCr
	    Response.Write " .frm1.txtHLatestShipDt.value		= """ & strDt  & """" & vbCr
	End If

	strDt =  UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_be_latest_ship_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtBeLatestShipDt.text		= """ & strDt & """" & vbCr
	End If		
	
	If UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_partial_ship)))) = "Y" Then
		Response.Write " .frm1.rdoAtPartialShip1.Checked		= True" & vbCr
	ElseIf UCase(Trim(ConvSPChars(E17_m_lc_amend_hdr(M445_E17_at_partial_ship)))) = "N" Then		    
		Response.Write " .frm1.rdoAtPartialShip2.Checked		= True" & vbCr
	End If
				
	Response.Write " .frm1.txtBePartialShip.value	= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_partial_ship)) & """" & vbCr
	Response.Write " .frm1.txtOpenBank.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_open_bank)) & """" & vbCr
	Response.Write " .frm1.txtOpenBankNm.value		= """ & ConvSPChars(E19_b_bank_issue_bank(M445_E19_bank_nm)) & """" & vbCr
	Response.Write " .frm1.txtAdvBank.value		    = """ & ConvSPChars(E20_b_bank_advise_bank(M445_E20_bank_cd)) & """" & vbCr
	Response.Write " .frm1.txtAdvBankNm.value		= """ & ConvSPChars(E20_b_bank_advise_bank(M445_E20_bank_nm)) & """" & vbCr
				
	strDt = UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_open_dt))
	If strDt <> strDefDate Then 
		Response.Write " .frm1.txtOpenDt.text = """ & strDt & """" & vbcr
	End if
		
	Response.Write " .frm1.txtApplicant.value 	= """ & ConvSPChars(E27_b_biz_partner_applicant(M445_E27_bp_cd)) & """" & vbCr
	Response.Write " .frm1.txtApplicantNm.value = """ & ConvSPChars(E27_b_biz_partner_applicant(M445_E27_bp_nm)) & """" & vbCr
	Response.Write " .frm1.txtPurGrp.value		= """ & ConvSPChars(E25_b_pur_grp(M445_E25_pur_grp)) & """" & vbCr
	Response.Write " .frm1.txtPurGrpNm.value	= """ & ConvSPChars(E25_b_pur_grp(M445_E25_pur_grp_nm)) & """" & vbCr
	Response.Write " .frm1.txtPurOrg.value		= """ & ConvSPChars(E24_b_pur_org(M445_E24_pur_org)) & """" & vbCr
	Response.Write " .frm1.txtPurOrgNm.value	= """ & ConvSPChars(E24_b_pur_org(M445_E24_pur_org_nm)) & """" & vbCr
	Response.Write " .frm1.txtRemark.value	    = """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_remark )) & """" & vbCr

	Response.Write " .frm1.txtLCAmt.value	= """ & UNINumClientFormat(E18_m_lc_hdr(M445_E18_Doc_Amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	Response.Write " .frm1.txtHLCNo.value	= """ & ConvSPChars(E18_m_lc_hdr(M445_E18_lc_no)) & """" & vbCr
		
	Response.Write " .DbQueryOk "		    	  & vbCr 										'☜: 조회가 성공 
        
    Response.Write " .frm1.txtHLCAmdNo.value = """ & ConvSPChars(Request("txtLCAmdNo")) & """" & vbCr
	Response.Write "End With" & vbCr  
    Response.Write "</Script>" & vbCr
		
End Sub




'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    DIM  lgIntFlgMode
    DIM PM4G211
    Dim Command 
    Dim I1_b_biz_partner_bp_cd 
    Dim I2_b_biz_partner_bp_cd 
    Dim I3_m_lc_amend_hdr 
    Dim I4_s_wks_user 
    Dim I5_b_pur_grp 
	Dim I6_m_lc_amend_hdr
    Dim strConvDt
    
     Const M435_I3_lc_amd_no = 0
     Const M435_I3_lc_no = 1
     Const M435_I3_lc_doc_no = 2
     Const M435_I3_lc_amend_seq = 3
     Const M435_I3_adv_no = 4
     Const M435_I3_pre_adv_ref = 5
     Const M435_I3_open_dt = 6
     Const M435_I3_be_expiry_dt = 7
     Const M435_I3_at_expiry_dt = 8
     Const M435_I3_manufacturer = 9
     Const M435_I3_agent = 10
     Const M435_I3_amend_dt = 11
     Const M435_I3_amend_req_dt = 12
     Const M435_I3_currency = 13
     Const M435_I3_be_doc_amt = 14
     Const M435_I3_at_doc_amt = 15
     Const M435_I3_at_xch_rate = 16
     Const M435_I3_inc_amt = 17
     Const M435_I3_dec_amt = 18
     Const M435_I3_be_loc_amt = 19
     Const M435_I3_at_loc_amt = 20
     Const M435_I3_be_partial_ship = 21
     Const M435_I3_at_partial_ship = 22
     Const M435_I3_be_latest_ship_dt = 23
     Const M435_I3_at_latest_ship_dt = 24
     Const M435_I3_open_bank = 25
     Const M435_I3_be_xch_rate = 26
     Const M435_I3_ext1_amt = 27
     Const M435_I3_ext1_cd = 28
     Const M435_I3_remark = 29
     Const M435_I3_lc_kind = 30
     Const M435_I3_remark2 = 31
     Const M435_I3_be_transhipment = 32
     Const M435_I3_at_transhipment = 33
     Const M435_I3_be_transfer = 34
     Const M435_I3_at_transfer = 35
     Const M435_I3_be_loading_port = 36
     Const M435_I3_at_loading_port = 37
     Const M435_I3_be_dischge_port = 38
     Const M435_I3_at_dischge_port = 39
     Const M435_I3_be_transport = 40
     Const M435_I3_at_transport = 41
     Const M435_I3_biz_area = 42
     Const M435_I3_charge_flg = 43
     Const M435_I3_adv_bank = 44
     Const M435_I3_ext1_qty = 45
     Const M435_I3_ext2_qty = 46
     Const M435_I3_ext3_qty = 47
     Const M435_I3_ext2_amt = 48
     Const M435_I3_ext3_amt = 49
     Const M435_I3_ext2_cd = 50
     Const M435_I3_ext3_cd = 51
     Const M435_I3_ext1_rt = 52
     Const M435_I3_ext2_rt = 53
     Const M435_I3_ext3_rt = 54
     Const M435_I3_ext1_dt = 55
     Const M435_I3_ext2_dt = 56
     Const M435_I3_ext3_dt = 57
    
    Set PM4G211 = Server.CreateObject("PM4G211.cMMaintLcAmendHdrS")
    
   On Error Resume Next                                                            '☜: Protect system from crashing
	Err.Clear   

    lgIntFlgMode = CInt(Request("txtFlgMode"))	

       redim I3_m_lc_amend_hdr(60)
        
        '-----------------------
		I3_m_lc_amend_hdr(M435_I3_lc_amd_no)  = UCase(Trim(Request("txtLCAmdNo1")))
		I3_m_lc_amend_hdr(M435_I3_lc_doc_no)  = UCase(Request("txtLCDocNo"))

		if Len(Trim(Request("txtLCAmendSeq"))) Then
			I3_m_lc_amend_hdr(M435_I3_lc_amend_seq ) = UNIConvNum(Request("txtLCAmendSeq"),0)
		End If
		
		I2_b_biz_partner_bp_cd = UCase(Trim(Request("txtBeneficiary")))
		
	   If Len(Trim(Request("txtAmendReqDt"))) Then
			I3_m_lc_amend_hdr(M435_I3_amend_req_dt)    = UNIConvDate(Request("txtAmendReqDt"))
		End If		
		
		If Len(Trim(Request("txtAmendDt"))) Then
			I3_m_lc_amend_hdr(M435_I3_amend_dt)  = UNIConvDate(Request("txtAmendDt"))
		End If
		
		If Request("rdoAtDocAmt") = "I" Then
			If Len(Trim(Request("txtAmendAmt"))) Then
				I3_m_lc_amend_hdr(M435_I3_inc_amt) = UNIConvNum(Request("txtAmendAmt"),0)
			'	I3_m_lc_amend_hdr(M435_I3_at_doc_amt ) = UNIConvNum(Request("txtAtDocAmt"),0)
			End If
		ElseIf Request("rdoAtDocAmt") = "D" Then
			If Len(Trim(Request("txtAmendAmt"))) Then
				I3_m_lc_amend_hdr(M435_I3_dec_amt) = UNIConvNum(Request("txtAmendAmt"),0)
			'	I3_m_lc_amend_hdr(M435_I3_at_doc_amt ) = UNIConvNum(Request("txtAtDocAmt"),0)
			End If
		End If
		'수정 
		'M32211.ImportMLcAmendHdrBeDocAmt = UNIConvNum(Request("txtBeDocAmt"),0)   '추가함 
		I3_m_lc_amend_hdr(M435_I3_be_doc_amt)  = UNIConvNum(Request("txtBeDocAmt"),0)
		I3_m_lc_amend_hdr(M435_I3_at_doc_amt)  = UNIConvNum(Request("txtAtDocAmt"),0)
	    I3_m_lc_amend_hdr(M435_I3_at_xch_rate)  = UNIConvNum(Request("txtAtXchRate"),0)
		I3_m_lc_amend_hdr(M435_I3_currency ) = UCase(Trim(Request("txtCurrency")))
		
	
		If len(Trim(request("txtAtExpireDt"))) then
			I3_m_lc_amend_hdr(M435_I3_at_expiry_dt) = UNIConvDate(Request("txtAtExpireDt"))		
		End If
        
        If len(Trim(request("txtBeExpireDt"))) then
			I3_m_lc_amend_hdr(M435_I3_be_expiry_dt) = UNIConvDate(Request("txtBeExpireDt"))		
		End If
		
		If Len(Trim(Request("txtatLatestShipDt"))) Then
			I3_m_lc_amend_hdr(M435_I3_at_latest_ship_dt) = UNIConvDate(Request("txtatLatestShipDt"))
		End If

		If Len(Trim(Request("txtBeLatestShipDt"))) Then
			I3_m_lc_amend_hdr(M435_I3_be_latest_ship_dt) = UNIConvDate(Request("txtBeLatestShipDt"))
		End If
		   
		I3_m_lc_amend_hdr(M435_I3_at_partial_ship) = Request("rdoAtPartialShip")
		I3_m_lc_amend_hdr(M435_I3_be_partial_ship) =  UCase(Trim(Request("txtBePartialShip")))
		I3_m_lc_amend_hdr(M435_I3_open_bank)  = UCase(Trim(Request("txtOpenBank")))
		I3_m_lc_amend_hdr(M435_I3_adv_bank)  = UCase(Trim(Request("txtAdvBank")))
		
		If Len(Trim(Request("txtBeLatestShipDt"))) Then
			I3_m_lc_amend_hdr(M435_I3_open_dt)  = UNIConvDate(Request("txtOpenDt"))
		End If
		
		I2_b_biz_partner_bp_cd= Request("txtApplicant")
		I1_b_biz_partner_bp_cd = Request("txtBeneficiary")
		I5_b_pur_grp =  Request("txtPurGrp")
		I3_m_lc_amend_hdr(M435_I3_remark)  = UCase(Trim(Request("txtRemark")))   
		
		I3_m_lc_amend_hdr(M435_I3_lc_kind) = "L"
		I3_m_lc_amend_hdr(M435_I3_lc_no)  = UCase(Trim(Request("txtHLCNo"))) 
		
		If Len(Trim(Request("txtAtLocAmt"))) Then
			I3_m_lc_amend_hdr(M435_I3_at_loc_amt)  = UNIConvNum(Request("txtAtLocAmt"),0)
		End If
		
		If lgIntFlgMode = OPMD_CMODE Then
			Command = "Create"
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			Command = "Update"
		End If

	    CALL PM4G211.M_MAINT_LC_AMEND_HDR_SVR(gStrGlobalCollection , Command, I1_b_biz_partner_bp_cd,I2_b_biz_partner_bp_cd,I3_m_lc_amend_hdr,I4_s_wks_user,I5_b_pur_grp,I6_m_lc_amend_hdr) 
        
        if CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set PM4G211 = Nothing												'☜: ComProxy Unload
	        Exit Sub
																'☜: 비지니스 로직 처리를 종료함 
 	end if
               
	Response.Write "<Script Language=vbscript>" & vbCr
	
	If Trim(ConvSPChars(Trim(I6_m_lc_amend_hdr))) <> ""  Then
		Response.Write " parent.frm1.txtLCAmdNo.value = """ & ConvSPChars(Trim(I6_m_lc_amend_hdr)) & """" & vbCr
		Response.Write " parent.frm1.txtLCAmdNo1.value = """ & ConvSPChars(Trim(I6_m_lc_amend_hdr)) & """" & vbCr
	End If
	
	Response.Write " Parent.DBSaveOK "           & vbCr
	Response.Write "</Script>"                  & vbCr 
       
End Sub    



%>
