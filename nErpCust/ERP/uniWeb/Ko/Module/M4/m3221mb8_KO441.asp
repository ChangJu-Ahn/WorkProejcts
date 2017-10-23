<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%	
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
     
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
                                                   '☜: Query
    Call SubBizQuery()
         
'=======================================================================================================
 Sub SubBizQuery()
 
 Dim iPM4G119 
 dim lgCurrency
    
   Const M446_I1_lc_no = 0
   Const M446_I1_lc_kind = 1

   Dim I1_m_lc_hdr
   Dim E1_ief_supplied 
   Dim E2_b_minor 
   Dim E3_b_daily_exchange_rate 
   Dim E4_b_minor 
   Dim E5_b_minor 
   Dim E6_b_minor 
   Dim E7_b_minor 
   Dim E8_b_minor 
   Dim E9_b_minor 
   Dim E10_b_minor 
   Dim E11_b_minor 
   Dim E12_b_minor 
   Dim E13_b_minor 
   Dim E14_b_minor 
   Dim E15_b_minor 
   Dim E16_b_minor 
   Dim E17_b_minor 
   Dim E18_m_lc_hdr 
   Dim E19_b_bank 
   Dim E20_b_bank 
   Dim E21_b_bank 
   Dim E22_b_bank 
   Dim E23_b_bank 
   Dim E24_b_pur_org 
   Dim E25_b_pur_grp 
   Dim E26_b_biz_partner 
   Dim E27_b_biz_partner 
   Dim E28_b_biz_partner 
   Dim E29_b_biz_partner 
   Dim E30_b_biz_partner 
   
    Const EA_m_lc_hdr_lc_no1 = 0
    Const EA_m_lc_hdr_lc_doc_no1 = 1
    Const EA_m_lc_hdr_lc_amend_seq1 = 2
    Const EA_m_lc_hdr_po_no1 = 3
    Const EA_m_lc_hdr_adv_no1 = 4
    Const EA_m_lc_hdr_pre_adv_ref1 = 5
    Const EA_m_lc_hdr_req_dt1 = 6
    Const EA_m_lc_hdr_adv_dt1 = 7
    Const EA_m_lc_hdr_open_dt1 = 8
    Const EA_m_lc_hdr_expiry_dt1 = 9
    Const EA_m_lc_hdr_amend_dt1 = 10
    Const EA_m_lc_hdr_manufacturer1 = 11
    Const EA_m_lc_hdr_agent1 = 12
    Const EA_m_lc_hdr_currency1 = 13
    Const EA_m_lc_hdr_doc_amt1 = 14
    Const EA_m_lc_hdr_xch_rate1 = 15
    Const EA_m_lc_hdr_xch_rate_op1 = 16
    Const EA_m_lc_hdr_loc_amt1 = 17
    Const EA_m_lc_hdr_bank_txt1 = 18
    Const EA_m_lc_hdr_incoterms1 = 19
    Const EA_m_lc_hdr_pay_method1 = 20
    Const EA_m_lc_hdr_pay_terms_txt1 = 21
    Const EA_m_lc_hdr_partial_ship1 = 22
    Const EA_m_lc_hdr_latest_ship_dt1 = 23
    Const EA_m_lc_hdr_shipment1 = 24
    Const EA_m_lc_hdr_doc11 = 25
    Const EA_m_lc_hdr_doc21 = 26
    Const EA_m_lc_hdr_doc31 = 27
    Const EA_m_lc_hdr_doc41 = 28
    Const EA_m_lc_hdr_doc51 = 29
    Const EA_m_lc_hdr_file_dt1 = 30
    Const EA_m_lc_hdr_file_dt_txt1 = 31
    Const EA_m_lc_hdr_insrt_user_id1 = 32
    Const EA_m_lc_hdr_insrt_dt1 = 33
    Const EA_m_lc_hdr_updt_user_id1 = 34
    Const EA_m_lc_hdr_updt_dt1 = 35
    Const EA_m_lc_hdr_ext1_qty1 = 36
    Const EA_m_lc_hdr_ext1_amt1 = 37
    Const EA_m_lc_hdr_ext1_cd1 = 38
    Const EA_m_lc_hdr_remark1 = 39
    Const EA_m_lc_hdr_lc_kind1 = 40
    Const EA_m_lc_hdr_lc_type1 = 41
    Const EA_m_lc_hdr_transhipment1 = 42
    Const EA_m_lc_hdr_transfer1 = 43
    Const EA_m_lc_hdr_delivery_plce1 = 44
    Const EA_m_lc_hdr_tolerance1 = 45
    Const EA_m_lc_hdr_loading_port1 = 46
    Const EA_m_lc_hdr_dischge_port1 = 47
    Const EA_m_lc_hdr_transport_comp1 = 48
    Const EA_m_lc_hdr_origin1 = 49
    Const EA_m_lc_hdr_origin_cntry1 = 50
    Const EA_m_lc_hdr_charge_txt1 = 51
    Const EA_m_lc_hdr_charge_cd1 = 52
    Const EA_m_lc_hdr_credit_core1 = 53
    Const EA_m_lc_hdr_lc_remn_doc_amt1 = 54
    Const EA_m_lc_hdr_lc_remn_loc_amt1 = 55
    Const EA_m_lc_hdr_fund_type1 = 56
    Const EA_m_lc_hdr_lmt_xch_rate1 = 57
    Const EA_m_lc_hdr_lmt_amt1 = 58
    Const EA_m_lc_hdr_inv_cnt1 = 59
    Const EA_m_lc_hdr_bl_awb_flg1 = 60
    Const EA_m_lc_hdr_freight1 = 61
    Const EA_m_lc_hdr_notify_party1 = 62
    Const EA_m_lc_hdr_consignee1 = 63
    Const EA_m_lc_hdr_insur_policy1 = 64
    Const EA_m_lc_hdr_pack_list1 = 65
    Const EA_m_lc_hdr_cert_origin_flg1 = 66
    Const EA_m_lc_hdr_transport1 = 67
    Const EA_m_lc_hdr_l_lc_type1 = 68
    Const EA_m_lc_hdr_o_lc_kind1 = 69
    Const EA_m_lc_hdr_o_lc_doc_no1 = 70
    Const EA_m_lc_hdr_o_lc_amend_seq1 = 71
    Const EA_m_lc_hdr_o_lc_no1 = 72
    Const EA_m_lc_hdr_o_lc_type1 = 73
    Const EA_m_lc_hdr_o_lc_open_dt1 = 74
    Const EA_m_lc_hdr_o_lc_expiry_dt1 = 75
    Const EA_m_lc_hdr_o_lc_loc_amt1 = 76
    Const EA_m_lc_hdr_biz_area1 = 77
    Const EA_m_lc_hdr_pay_dur1 = 78
    Const EA_m_lc_hdr_charge_flg1 = 79
    Const EA_m_lc_hdr_ext2_qty1 = 80
    Const EA_m_lc_hdr_ext3_qty1 = 81
    Const EA_m_lc_hdr_ext2_amt1 = 82
    Const EA_m_lc_hdr_ext3_amt1 = 83
    Const EA_m_lc_hdr_ext2_cd1 = 84
    Const EA_m_lc_hdr_ext3_cd1 = 85
    Const EA_m_lc_hdr_ext1_rt1 = 86
    Const EA_m_lc_hdr_ext2_rt1 = 87
    Const EA_m_lc_hdr_ext3_rt1 = 88
    Const EA_m_lc_hdr_ext1_dt1 = 89
    Const EA_m_lc_hdr_ext2_dt1 = 90
    Const EA_m_lc_hdr_ext3_dt1 = 91
    Const EA_m_lc_hdr_pur_org1 = 92
    Const EA_m_lc_hdr_pur_grp1 = 93
    Const EA_m_lc_hdr_applicant1 = 94
    Const EA_m_lc_hdr_beneficiary1 = 95
    Const EA_m_lc_hdr_open_bank1 = 96
    Const EA_m_lc_hdr_adv_bank1 = 97
    Const EA_m_lc_hdr_renego_bank1 = 98
    Const EA_m_lc_hdr_pay_bank1 = 99
    Const EA_m_lc_hdr_return_bank1 = 100
	
	Dim strDt
	Dim strDefDate
	
	strDefDate="1900-01-01"
	
	On Error Resume Next

    Err.Clear 
    
    ReDim I1_m_lc_hdr(M446_I1_lc_kind)
    
    I1_m_lc_hdr(M446_I1_lc_no) 	     = Trim(Request("txtLCNo"))
    I1_m_lc_hdr(M446_I1_lc_kind) 	 = "L"
  
    Set iPM4G119 = Server.CreateObject("PM4G119.cMLookupLcHdrS")    
    
	If CheckSYSTEMError(Err,True) = true then 		
		Set iPM4G119 = Nothing
        Exit Sub
	End if

    Call iPM4G119.M_LOOKUP_LC_HDR_SVR(gStrGlobalCollection,I1_m_lc_hdr,E1_ief_supplied,E2_b_minor, _
    E3_b_daily_exchange_rate,E4_b_minor,E5_b_minor,E6_b_minor,E7_b_minor,E8_b_minor,E9_b_minor,E10_b_minor, _
    E11_b_minor,E12_b_minor, E13_b_minor, E14_b_minor, E15_b_minor,E16_b_minor,E17_b_minor,E18_m_lc_hdr, _
    E19_b_bank,E20_b_bank,E21_b_bank,E22_b_bank,E23_b_bank,E24_b_pur_org,E25_b_pur_grp,E26_b_biz_partner, _
    E27_b_biz_partner,E28_b_biz_partner,E29_b_biz_partner,E30_b_biz_partner)
    
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iPM4G119 = Nothing	
		Exit SUB								'☜: ComProxy Unload
        
	End if
	 
	Set iPM4G119 = Nothing
	
		lgCurrency		= ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_currency1))
	
	Response.Write "<Script Language=vbscript>" & vbCr	    
	Response.Write " With parent "	& vbCr
	Response.Write " .frm1.txtCurrency.value		= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_currency1)) & """" & vbCr
	'@수정(2003.03.12)
	Response.Write "    .CurFormatNumericOCX " & vbCr
	Response.Write " .frm1.txtLCDocNo.value		= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_doc_no1 )) & """" & vbCr
	Response.Write "if Trim(""" & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_doc_no1)) & """)  <> """" then " & vbCr
	Response.Write "  .frm1.txtLCAmendSeq.value 	= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_amend_seq1)) & """" & vbCr	
	Response.Write "end if " & vbCr	
	Response.Write " .frm1.txtBeneficiary.value 		= """ & ConvSPChars(E26_b_biz_partner(0)) & """" & vbCr
	Response.Write " .frm1.txtBeneficiaryNm.value		= """ & ConvSPChars(E26_b_biz_partner(1)) & """" & vbCr
'	Response.Write " strDt		= """ & UNIDateClientFormat(E18_m_lc_hdr(M445_E17_amend_req_dt)) & """" & vbCr

		Response.Write " .frm1.txtAtXchRate.text  = """ & UNINumClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_xch_rate1), ggExchRate.DecPoint, 0) & """" & vbCr
	
		Response.Write " .frm1.txtAtDocAmt.text   =  """ & UNIConvNumDBToCompanyByCurrency(E18_m_lc_hdr(EA_m_lc_hdr_doc_amt1),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
		Response.Write " .frm1.txtAtLocAmt.text   =  """ & UNINumClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_loc_amt1), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
		Response.Write " .frm1.txtBeDocAmt.text   =  """ & UNIConvNumDBToCompanyByCurrency(E18_m_lc_hdr(EA_m_lc_hdr_doc_amt1),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
		Response.Write " .frm1.txtBeLocAmt.text   =  """ & UNINumClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_loc_amt1), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	  
			strDt = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_expiry_dt1 ))
			
			If cDate(uniConvDate(strDt)) <> cDate(uniConvDate(strDefDate)) Then
				Response.Write " .frm1.txtAtExpireDt.text		= """ & strDt & """" & vbCr
				Response.Write " .frm1.txtHExpiryDt.value		= """ & strDt & """" & vbCr
					 
			End If	
		
		'	strDt = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_expiry_dt1 ))
					 
		'	If cDate(uniConvDate(strDt)) <> cDate(uniConvDate(strDefDate)) Then
				Response.Write " .frm1.txtBeExpireDt.text = """ & UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_expiry_dt1 )) & """" & vbcr		 
		
		'	End If	
			
			strDt = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_latest_ship_dt1 ))					
			If cDate(uniConvDate(strDt)) <> cDate(uniConvDate(strDefDate)) Then
			 	Response.Write " .frm1.txtAtLatestShipDt.text		= """ & strDt & """" & vbCr
			    Response.Write " .frm1.txtHLatestShipDt.value		= """ & strDt  & """" & vbCr
					
			End If
					
				
		'	strDt =  UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_latest_ship_dt1 ))
						
		'	IF cDate(uniConvDate(strDt)) <> cDate(uniConvDate(strDefDate)) Then
				Response.Write " .frm1.txtBeLatestShipDt.text		= """ & UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_latest_ship_dt1 )) & """" & vbCr
		'	End If		
	
		   	
		If ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_partial_ship1 )) = "Y" Then
			Response.Write " .frm1.rdoAtPartialShip1.Checked		= True" & vbCr
		ElseIf ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_partial_ship1 )) = "N" Then
		    
			Response.Write " .frm1.rdoAtPartialShip2.Checked		= True" & vbCr
		End If
				
	
		Response.Write " .frm1.txtBePartialShip.value		= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_partial_ship1 )) & """" & vbCr
		Response.Write " .frm1.txtOpenBank.value		= """ &ConvSPChars(E19_b_bank(0)) & """" & vbCr
		Response.Write " .frm1.txtOpenBankNm.value		= """ & ConvSPChars(E19_b_bank(1)) & """" & vbCr
		Response.Write " .frm1.txtAdvBank.value		= """ & ConvSPChars(E20_b_bank(0)) & """" & vbCr
		Response.Write " .frm1.txtAdvBankNm.value		= """ & ConvSPChars(E20_b_bank(1)) & """" & vbCr
				
		Response.Write " .frm1.txtOpenDt.text = """ & UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_open_dt1)) & """" & vbcr

        Response.Write " .frm1.txtApplicant.value 		= """ & ConvSPChars(E27_b_biz_partner(0)) & """" & vbCr
		Response.Write " .frm1.txtApplicantNm.value 	= """ & ConvSPChars(E27_b_biz_partner(1)) & """" & vbCr
		Response.Write " .frm1.txtPurGrp.value		= """ & ConvSPChars(E25_b_pur_grp(0)) & """" & vbCr
		Response.Write " .frm1.txtPurGrpNm.value		= """ & ConvSPChars(E25_b_pur_grp(1))  & """" & vbCr
		Response.Write " .frm1.txtPurOrg.value		= """ & ConvSPChars(E24_b_pur_org(0)) & """" & vbCr
		Response.Write " .frm1.txtPurOrgNm.value		= """ & ConvSPChars(E24_b_pur_org(1)) & """" & vbCr
		
		Response.Write " parent.RefOk "		    	  & vbCr 										'☜: 조회가 성공 
    	Response.Write "End With" & vbCr  
        Response.Write "</Script>" & vbCr
							
   
END SUB
	
%>	
