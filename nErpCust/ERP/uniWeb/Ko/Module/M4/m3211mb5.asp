<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 

	Dim lgOpModeCRUD
	
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")
				'☜: Read Operation Mode (CRUD)
	
'	Dim M32119	
'	Dim M32111
	dim lgCurrency

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)                                                         '☜: Query
	             Call SubBizQuery()
	        Case CStr(UID_M0002)
	             Call SubBizSave()
	        Case CStr(UID_M0003)                                                         '☜: Delete
	             Call SubBizDelete()
	        Case "ListLcDtl"                                                                 '☜: Check	
	             Call SubListLcDtl()    
	        Case "LookupDailyExRt"
				 Call SubLookupDailyExRt()     
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
																' Master L/C Header 조회용 Object
		
   If Request("txtLCNo") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	   Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
	   Exit Sub
   End If	
			
   Dim iPM4G119 
   Dim strDate
   Dim strDefDate
    
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
		Set iPM4G119 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End if
	
	Set iPM4G119 = Nothing					

		strDefDate=UNIDateClientFormat("1900-01-01")
		lgCurrency = ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_currency1))

		Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "on error resume next" & vbCr
			Response.Write "With parent.frm1" & vbCr

			'##### Rounding Logic #####
			'항상 거래화폐가 우선"
			Response.Write ".txtCurrency.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_currency1)) & """" & vbCr
			Response.Write "parent.CurFormatNumericOCX" & vbCr
			'##########################
			'Tab 1 
			Response.Write ".txtLCNo1.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_no1)) & """" & vbCr
			Response.Write ".txtLCDocNo.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_doc_no1)) & """" & vbCr
			Response.Write ".txtLCAmendSeq.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_amend_seq1)) & """" & vbCr
			Response.Write ".txtPONo.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_po_no1)) & """" & vbCr
			Response.Write ".txtLCType.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_l_lc_type1)) & """" & vbCr
			Response.Write ".txtLCTypeNm.value = """ & ConvSPChars(E15_b_minor(1)) & """" & vbCr
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_req_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtReqDt.text = """ & strDate & """" & vbCr
			End If
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_open_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtOpenDt.text = """ & strDate & """" & vbCr		
			End If
			
			Response.Write ".txtOpenBank.value = """ & ConvSPChars(E19_b_bank(0)) & """" & vbCr
			Response.Write ".txtOpenBankNm.value = """ & ConvSPChars(E19_b_bank(1)) & """" & vbCr	
			Response.Write ".txtAdvBank.value = """ & ConvSPChars(E20_b_bank(0)) & """" & vbCr
			Response.Write ".txtAdvBankNm.value = """ & ConvSPChars(E20_b_bank(1)) & """" & vbCr
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_expiry_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtExpiryDt.text = """ & strDate & """" & vbCr
			End If
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_latest_ship_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtLatestShipDt.text = """ & strDate & """" & vbCr
			End If
					
			Response.Write ".txtShipment.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_shipment1)) & """" & vbCr
			Response.Write ".txtCurrency.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_currency1)) & """" & vbCr
			
			If CDbl(E18_m_lc_hdr(EA_m_lc_hdr_doc_amt1)) <> 0 Then
				Response.Write ".txtDocAmt.text = """ & UNINumClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_doc_amt1), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
			Else
				Response.Write ".txtDocAmt.text = 0" & vbCr
			End If
	
			If CDbl(E18_m_lc_hdr(EA_m_lc_hdr_xch_rate1)) <> 0 Then 
				Response.Write ".txtXchRate.text = """ & UNINumClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_xch_rate1), ggExchRate.DecPoint, 0) & """" & vbCr
			Else
				Response.Write ".txtXchRate.text = 0" & vbCr
			End If
		
			If CDbl(E18_m_lc_hdr(EA_m_lc_hdr_loc_amt1)) <> 0 Then 
				Response.Write ".txtLocAmt.text = """ & UNINumClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_loc_amt1), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
			Else
				Response.Write ".txtLocAmt.text = 0" & vbCr
			End If
				
			Response.Write ".txtOLcKind.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_o_lc_kind1)) & """" & vbCr
			Response.Write ".txtOLcKindNm.value = """ & ConvSPChars(E17_b_minor(1)) & """" & vbCr
			
			If UCase(Trim(ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_partial_ship1)))) = "Y" Then
				Response.Write ".rdoPartailShip1.Checked = True" & vbCr
			ElseIf UCase(Trim(ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_partial_ship1)))) = "N" Then
				Response.Write ".rdoPartailShip2.Checked = True" & vbCr
			End If
			
			Response.Write ".txtFileDt.Text = """ & E18_m_lc_hdr(EA_m_lc_hdr_file_dt1) & """" & vbCr
			Response.Write ".txtFileDtTxt.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_file_dt_txt1)) & """" & vbCr
			Response.Write ".txtPayTerms.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_pay_method1)) & """" & vbCr
			Response.Write ".txtPayTermsNm.value = """ & ConvSPChars(E12_b_minor(1)) & """" & vbCr
			Response.Write ".txtPayDur.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_pay_dur1)) & """" & vbCr
			Response.Write ".txtPaytermstxt.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_pay_terms_txt1)) & """" & vbCr
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_adv_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtAdvDt.text = """ & strDate & """" & vbCr
			End If
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_amend_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtAmendDt.text = """ & strDate & """" & vbCr
			End If
				
			Response.Write ".txtPurGrp.value = """ & ConvSPChars(E25_b_pur_grp(0)) & """" & vbCr
			Response.Write ".txtPurGrpNm.value = """ & ConvSPChars(E25_b_pur_grp(1)) & """" & vbCr
			Response.Write ".txtBeneficiary.value = """ & ConvSPChars(E26_b_biz_partner(0)) & """" & vbCr
			Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(E26_b_biz_partner(1)) & """" & vbCr
			Response.Write ".txtPurOrg.value = """ & ConvSPChars(E24_b_pur_org(0)) & """" & vbCr
			Response.Write ".txtPurOrgNm.value = """ & ConvSPChars(E24_b_pur_org(1)) & """" & vbCr
			Response.Write ".txtApplicant.value = """ & ConvSPChars(E27_b_biz_partner(0)) & """" & vbCr
			Response.Write ".txtApplicantNm.value = """ & ConvSPChars(E27_b_biz_partner(1)) & """" & vbCr
			'Tab 2
			Response.Write ".txtDoc1.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_doc11)) & """" & vbCr
			Response.Write ".txtDoc2.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_doc21)) & """" & vbCr
			Response.Write ".txtDoc3.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_doc31)) & """" & vbCr
			Response.Write ".txtDoc4.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_doc41)) & """" & vbCr
			Response.Write ".txtDoc5.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_doc51)) & """" & vbCr
			Response.Write ".txtBankTxt.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_bank_txt1)) & """" & vbCr
			Response.Write ".txtRemark.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_remark1)) & """" & vbCr
			Response.Write ".txtOLCNo.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_o_lc_no1)) & """" & vbCr
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_o_lc_open_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtOLCOpenDt.text = """ & strDate & """" & vbCr
			End If
			
			strDate = UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_o_lc_expiry_dt1))
			If strDate <> strDefDate Then
				Response.Write ".txtOLCExpiryDt.text = """ & strDate & """" & vbCr
			End If
		
			Response.Write ".txtAgent.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_agent1)) & """" & vbCr
			Response.Write ".txtAgentNm.value = """ & ConvSPChars(E28_b_biz_partner(1)) & """" & vbCr
			Response.Write ".txtManufacturer.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_manufacturer1)) & """" & vbCr
			Response.Write ".txtManufacturerNm.value = """ & ConvSPChars(E29_b_biz_partner(1)) & """" & vbCr
			Response.Write ".txtMultiDiv.value = """ & ConvSPChars(E3_b_daily_exchange_rate(0)) & """" & vbCr
			Response.Write ".hdnXchRtOp.value = """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_xch_rate_op1)) & """" & vbCr    '13차 추가 
			 
			Response.Write "Call parent.DbQueryOk()	" & vbCr													'☜: 조회가 성공 
		Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr

		Set M32119 = Nothing														'☜: Unload Comproxy
End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSave()
	On Error Resume Next
	Err.Clear 
	
	Dim I1_b_biz_partner
    Dim I2_b_biz_partner
    Dim I3_b_pur_grp
    Dim I4_m_lc_hdr
	Dim I5_b_bank
    Dim I6_b_bank
    Dim I7_b_bank
    Dim I8_b_bank
    Dim I9_b_bank
    Dim I10_s_wks_user
        
    Const M468_I12_lc_no = 0
    Const M468_I12_lc_doc_no = 1
    Const M468_I12_lc_amend_seq = 2
    Const M468_I12_po_no = 3
    Const M468_I12_adv_no = 4
    Const M468_I12_pre_adv_ref = 5
    Const M468_I12_req_dt = 6
    Const M468_I12_adv_dt = 7
    Const M468_I12_open_dt = 8
    Const M468_I12_expiry_dt = 9
    Const M468_I12_amend_dt = 10
    Const M468_I12_manufacturer = 11
    Const M468_I12_agent = 12
    Const M468_I12_currency = 13
    Const M468_I12_doc_amt = 14
    Const M468_I12_xch_rate = 15
    Const M468_I12_xch_rate_op = 16
    Const M468_I12_loc_amt = 17
    Const M468_I12_bank_txt = 18
    Const M468_I12_incoterms = 19
    Const M468_I12_pay_method = 20
    Const M468_I12_pay_terms_txt = 21
    Const M468_I12_partial_ship = 22
    Const M468_I12_latest_ship_dt = 23
    Const M468_I12_shipment = 24
    Const M468_I12_doc1 = 25
    Const M468_I12_doc2 = 26
    Const M468_I12_doc3 = 27
    Const M468_I12_doc4 = 28
    Const M468_I12_doc5 = 29
    Const M468_I12_file_dt = 30
    Const M468_I12_file_dt_txt = 31
    Const M468_I12_insrt_user_id = 32
    Const M468_I12_insrt_dt = 33
    Const M468_I12_updt_user_id = 34
    Const M468_I12_updt_dt = 35
    Const M468_I12_ext1_qty = 36
    Const M468_I12_ext1_amt = 37
    Const M468_I12_ext1_cd = 38
    Const M468_I12_remark = 39
    Const M468_I12_lc_kind = 40
    Const M468_I12_lc_type = 41
    Const M468_I12_transhipment = 42
    Const M468_I12_transfer = 43
    Const M468_I12_delivery_plce = 44
    Const M468_I12_tolerance = 45
    Const M468_I12_loading_port = 46
    Const M468_I12_dischge_port = 47
    Const M468_I12_transport_comp = 48
    Const M468_I12_origin = 49
    Const M468_I12_origin_cntry = 50
    Const M468_I12_charge_txt = 51
    Const M468_I12_charge_cd = 52
    Const M468_I12_credit_core = 53
    Const M468_I12_lc_remn_doc_amt = 54
    Const M468_I12_lc_remn_loc_amt = 55
    Const M468_I12_fund_type = 56
    Const M468_I12_lmt_xch_rate = 57
    Const M468_I12_lmt_amt = 58
    Const M468_I12_inv_cnt = 59
    Const M468_I12_bl_awb_flg = 60
    Const M468_I12_freight = 61
    Const M468_I12_notify_party = 62
    Const M468_I12_consignee = 63
    Const M468_I12_insur_policy = 64
    Const M468_I12_pack_list = 65
    Const M468_I12_cert_origin_flg = 66
    Const M468_I12_l_lc_type = 67
    Const M468_I12_o_lc_kind = 68
    Const M468_I12_o_lc_doc_no = 69
    Const M468_I12_o_lc_amend_seq = 70
    Const M468_I12_o_lc_no = 71
    Const M468_I12_o_lc_type = 72
    Const M468_I12_o_lc_open_dt = 73
    Const M468_I12_o_lc_expiry_dt = 74
    Const M468_I12_o_lc_loc_amt = 75
    Const M468_I12_transport = 76
    Const M468_I12_biz_area = 77
    Const M468_I12_pay_dur = 78
    Const M468_I12_charge_flg = 79
    Const M468_I12_ext2_qty = 80
    Const M468_I12_ext3_qty = 81
    Const M468_I12_ext2_amt = 82
    Const M468_I12_ext3_amt = 83
    Const M468_I12_ext2_cd = 84
    Const M468_I12_ext3_cd = 85
    Const M468_I12_ext1_rt = 86
    Const M468_I12_ext2_rt = 87
    Const M468_I12_ext3_rt = 88
    Const M468_I12_ext1_dt = 89
    Const M468_I12_ext2_dt = 90
    Const M468_I12_ext3_dt = 91
            
            
	Dim iPM4G111
	Dim strConvDt
	Dim lgIntFlgMode
	
	Redim I4_m_lc_hdr(M468_I12_ext3_dt)
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))								'☜: 저장시 Create/Update 판별 
	
	on error resume next
		Err.Clear																'☜: Protect system from crashing

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------

		'Tab 1
		I4_m_lc_hdr(M468_I12_lc_doc_no) = UCase(Trim(Request("txtLCDocNo")))

		If Len(Trim(Request("txtLCAmendSeq"))) Then
			I4_m_lc_hdr(M468_I12_lc_amend_seq) = UNIConvNum(Request("txtLCAmendSeq"),0)
		End If
		
		If Trim(Request("txtPONoFlg")) = "Y" Then
			I4_m_lc_hdr(M468_I12_po_no) = UCase(Trim(Request("txtPONo")))
		End If	
		
		I4_m_lc_hdr(M468_I12_l_lc_type) = UCase(Trim(Request("txtLCType")))
		
		If Len(Trim(Request("txtReqDt"))) Then
			strConvDt = UNIConvDate(Request("txtReqDt"))

			If strConvDt = "" Then
				'Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				'Call LoadTab("parent.frm1.txtReqDt", 1, I_MKSCRIPT)
				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_req_dt) = strConvDt
			End If
		End If	

		If Len(Trim(Request("txtOpenDt"))) Then
			strConvDt = UNIConvDate(Request("txtOpenDt"))

			If strConvDt = "" Then
				'Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				'Call LoadTab("parent.frm1.txtOpenDt", 1, I_MKSCRIPT)
				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_open_dt) = strConvDt
			End If
		Else
		    I4_m_lc_hdr(M468_I12_open_dt) = UNIConvDate(UniConvYYYYMMDDToDate(gDateFormat,"1900","01","01"))    
		End If	
		
		I5_b_bank = UCase(Trim(Request("txtOpenBank")))
		I6_b_bank = UCase(Trim(Request("txtAdvBank")))

		If Len(Trim(Request("txtExpiryDt"))) Then
			strConvDt = UNIConvDate(Request("txtExpiryDt"))
			I4_m_lc_hdr(M468_I12_expiry_dt) = strConvDt
		End If	
		
		If Len(Trim(Request("txtLatestShipDt"))) Then
			strConvDt = UNIConvDate(Request("txtLatestShipDt"))
			I4_m_lc_hdr(M468_I12_latest_ship_dt) = strConvDt
		End If	
		
		I4_m_lc_hdr(M468_I12_shipment) = Trim(Request("txtShipment"))
		
		I4_m_lc_hdr(M468_I12_currency) = UCase(Trim(Request("txtCurrency")))
		
		If Len(Trim(Request("txtDocAmt"))) Then
			I4_m_lc_hdr(M468_I12_doc_amt) = UNIConvNum(Request("txtDocAmt"),0)
		End If

		If Len(Trim(Request("txtXchRate"))) Then
			I4_m_lc_hdr(M468_I12_xch_rate) = UNIConvNum(Request("txtXchRate"),0)
		End If
		
		If Len(Trim(Request("txtLocAmt"))) Then
			I4_m_lc_hdr(M468_I12_loc_amt) = UNIConvNum(Request("txtLocAmt"),0)
			'I4_m_lc_hdr(M468_I12_loc_amt) = Request("txtLocAmt")
		End If
		
		I4_m_lc_hdr(M468_I12_o_lc_kind) = UCase(Trim(Request("txtOLCKind")))
		I4_m_lc_hdr(M468_I12_partial_ship) = Request("rdoPartailShip")
		I4_m_lc_hdr(M468_I12_file_dt) = Request("txtFileDt")
		I4_m_lc_hdr(M468_I12_file_dt_txt) = Trim(Request("txtFileDtTxt"))
		I4_m_lc_hdr(M468_I12_pay_method) = UCase(Trim(Request("txtPayTerms")))
		
		If Len(Trim(Request("txtPayDur"))) Then
			I4_m_lc_hdr(M468_I12_pay_dur) = UNIConvNum(Request("txtPayDur"),0)
		End If

		I4_m_lc_hdr(M468_I12_pay_terms_txt) = Trim(Request("txtPaytermstxt"))

		If Len(Trim(Request("txtAdvDt"))) Then
			strConvDt = UNIConvDate(Request("txtAdvDt"))
			I4_m_lc_hdr(M468_I12_adv_dt) = strConvDt
		End If
		
		If Len(Trim(Request("txtAmendDt"))) Then
			strConvDt = UNIConvDate(Request("txtAmendDt"))

			If strConvDt = "" Then
				'Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				'Call LoadTab("parent.frm1.txtAmendDt", 1, I_MKSCRIPT)
				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_amend_dt) = strConvDt
			End If
		End If	
		
		I3_b_pur_grp = UCase(Trim(Request("txtPurGrp")))
		I1_b_biz_partner = UCase(Trim(Request("txtBeneficiary")))
		I2_b_biz_partner = UCase(Trim(Request("txtApplicant")))

	
		'Tab 2
		
		I4_m_lc_hdr(M468_I12_doc1) = Trim(Request("txtDoc1"))
		I4_m_lc_hdr(M468_I12_doc2) = Trim(Request("txtDoc2"))
		I4_m_lc_hdr(M468_I12_doc3) = Trim(Request("txtDoc3"))
		I4_m_lc_hdr(M468_I12_doc4) = Trim(Request("txtDoc4"))
		I4_m_lc_hdr(M468_I12_doc5) = Trim(Request("txtDoc5"))
		I4_m_lc_hdr(M468_I12_bank_txt) = Trim(Request("txtBankTxt"))
		I4_m_lc_hdr(M468_I12_remark) = Trim(Request("txtRemark"))
		I4_m_lc_hdr(M468_I12_o_lc_no) = Trim(Request("txtOLCNo"))
		
		I4_m_lc_hdr(M468_I12_o_lc_amend_seq) = UNIConvNum(Request("txtOLcAmendSeq"),0)
		
		If Len(Trim(Request("txtOLcOpenDt"))) Then
			strConvDt = UNIConvDate(Request("txtOLcOpenDt"))

			If strConvDt = "" Then
				'Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				'Call LoadTab("parent.frm1.txtOLcOpenDt", 1, I_MKSCRIPT)

				Set M32111 = Nothing											'☜: ComProxy UnLoad
				exit sub
				'Response.End													'☜: Process End
			Else
				I4_m_lc_hdr(M468_I12_o_lc_open_dt) = strConvDt
			End If
		End If	
		
		If Len(Trim(Request("txtOLcExpiryDt"))) Then
			strConvDt = UNIConvDate(Request("txtOLcExpiryDt"))

			If strConvDt = "" Then
				'Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				'Call LoadTab("parent.frm1.txtOLcExpiryDt", 1, I_MKSCRIPT)]
				Exit Sub
			Else
				I4_m_lc_hdr(M468_I12_o_lc_expiry_dt) = strConvDt
			End If
		End If		
		
		I4_m_lc_hdr(M468_I12_agent) = UCase(Trim(Request("txtAgent")))
		I4_m_lc_hdr(M468_I12_manufacturer) = UCase(Trim(Request("txtManufacturer")))
			
		I4_m_lc_hdr(M468_I12_insrt_user_id) = UCase(Trim(Request("txtInsrtUserId")))
		I4_m_lc_hdr(M468_I12_lc_kind) = "L"
		I4_m_lc_hdr(M468_I12_lc_no) = UCase(Trim(Request("txtLCNo1")))
		I4_m_lc_hdr(M468_I12_xch_rate_op) = UCase(Trim(Request("hdnXchRtOp")))

	Set iPM4G111 = Server.CreateObject("PM4G111.cMMaintLcHdrS")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	
	Dim sTemp
	
	If lgIntFlgMode = OPMD_CMODE Then
		sTemp = iPM4G111.M_MAINT_LC_HDR_SVR(gStrGlobalCollection,"CREATE",I1_b_biz_partner,I2_b_biz_partner, _
            I3_b_pur_grp,I4_m_lc_hdr,I5_b_bank,I6_b_bank,I7_b_bank,I8_b_bank,I9_b_bank,I10_s_wks_user)
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		Call iPM4G111.M_MAINT_LC_HDR_SVR(gStrGlobalCollection,"UPDATE",I1_b_biz_partner,I2_b_biz_partner, _
            I3_b_pur_grp,I4_m_lc_hdr,I5_b_bank,I6_b_bank,I7_b_bank,I8_b_bank,I9_b_bank,I10_s_wks_user)
    End if        
    
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM4G111 = Nothing
    
		'-----------------------
		'Result data display area
		'-----------------------
    Response.Write "<Script language=vbs> " & vbCr 
'				Response.Write "With parent" & vbCr
	Response.Write "If """ & ConvSPChars(Trim(sTemp)) & """ <> """"  Then " & vbCr
    Response.Write "parent.frm1.txtLCNo1.value = """ & ConvSPChars(sTemp) & """" & vbCr
    Response.Write "parent.frm1.txtLCNo.value = """ & ConvSPChars(sTemp) & """" & vbCr
    Response.Write "End If" & vbCr
'				Response.Write ".DbSaveOk" & vbCr
		
    Response.Write "Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "            
End Sub
'============================================================================================================
' Name : SubDelete
' Desc : Delete data from DB
'============================================================================================================
Sub SubBizDelete()
	Dim iPM4G111
	
	Dim I1_b_biz_partner
    Dim I2_b_biz_partner
    Dim I3_b_pur_grp
    Dim I4_m_lc_hdr
	Dim I5_b_bank
    Dim I6_b_bank
    Dim I7_b_bank
    Dim I8_b_bank
    Dim I9_b_bank
    Dim I10_s_wks_user
    
    Const M468_I12_lc_no = 0
    
    Redim I4_m_lc_hdr(91)
    
    On Error Resume Next
    Err.Clear																'☜: Protect system from crashing

	If Request("txtLCNo") = "" Then											'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		exit sub
	End If

    I4_m_lc_hdr(M468_I12_lc_no) = Trim(Request("txtLCNo"))

	Set iPM4G111 = Server.CreateObject("PM4G111.cMMaintLcHdrS")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

    Call iPM4G111.M_MAINT_LC_HDR_SVR(gStrGlobalCollection,"DELETE",I1_b_biz_partner,I2_b_biz_partner, _
            I3_b_pur_grp,I4_m_lc_hdr,I5_b_bank,I6_b_bank,I7_b_bank,I8_b_bank,I9_b_bank,I10_s_wks_user)       

	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM4G111 = Nothing

	Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.DbDeleteOk "      & vbCr   
    Response.Write "</Script> "            
End Sub															'☜: Process End
'============================================================================================================
' Name : SubListLcDtl
' Desc : 
'============================================================================================================
Sub SubListLcDtl()

    Dim iPM4G128

    Dim I1_m_lc_dtl_lc_seq
'    Dim I2_m_lc_hdr_lc_no
    Dim E1_m_lc_dtl_total_amt
    Dim E2_m_lc_dtl_max_lc_seq
    Dim E3_m_lc_dtl_next_lc_seq
    Dim EG1_export_group

     Const M474_I2_lc_no = 0
    Const M474_I2_lc_doc_no = 1

    Const M474_E1_doc_amt = 0
    Const M474_E1_loc_amt = 1

    Const M474_EG1_E1_m_pur_goods_mvmt_rcpt_no = 0
    '[CONVERSION INFORMATION]  View Name : export_item m_pur_ord_hdr
    Const M474_EG1_E2_m_pur_ord_hdr_po_no = 1
    '[CONVERSION INFORMATION]  View Name : export_item m_pur_ord_dtl
    Const M474_EG1_E3_m_pur_ord_dtl_po_seq_no = 2
    Const M474_EG1_E3_m_pur_ord_dtl_po_qty = 3
    Const M474_EG1_E3_m_pur_ord_dtl_lc_qty = 4
    Const M474_EG1_E3_m_pur_ord_dtl_after_lc_flg = 5
        '[CONVERSION INFORMATION]  View Name : export_item m_lc_dtl
    Const M474_EG1_E4_m_lc_dtl_lc_seq = 6
    Const M474_EG1_E4_m_lc_dtl_hs_cd = 7
    Const M474_EG1_E4_m_lc_dtl_qty = 8
    Const M474_EG1_E4_m_lc_dtl_price = 9
    Const M474_EG1_E4_m_lc_dtl_doc_amt = 10
    Const M474_EG1_E4_m_lc_dtl_loc_amt = 11
    Const M474_EG1_E4_m_lc_dtl_unit = 12
    Const M474_EG1_E4_m_lc_dtl_over_tolerance = 13
    Const M474_EG1_E4_m_lc_dtl_under_tolerance = 14
    Const M474_EG1_E4_m_lc_dtl_close_flg = 15
    Const M474_EG1_E4_m_lc_dtl_receipt_qty = 16
    Const M474_EG1_E4_m_lc_dtl_insrt_user_id = 17
    Const M474_EG1_E4_m_lc_dtl_insrt_dt = 18
    Const M474_EG1_E4_m_lc_dtl_updt_user_id = 19
    Const M474_EG1_E4_m_lc_dtl_updt_dt = 20
    Const M474_EG1_E4_m_lc_dtl_ext1_qty = 21
    Const M474_EG1_E4_m_lc_dtl_ext1_amt = 22
    Const M474_EG1_E4_m_lc_dtl_ext1_cd = 23
    Const M474_EG1_E4_m_lc_dtl_lc_kind = 24
    Const M474_EG1_E4_m_lc_dtl_bl_qty = 25
    Const M474_EG1_E4_m_lc_dtl_il_no = 26
    Const M474_EG1_E4_m_lc_dtl_il_seq = 27
    Const M474_EG1_E4_m_lc_dtl_remark2 = 28
    Const M474_EG1_E4_m_lc_dtl_biz_area = 29
    Const M474_EG1_E4_m_lc_dtl_ext2_qty = 30
    Const M474_EG1_E4_m_lc_dtl_ext3_qty = 31
    Const M474_EG1_E4_m_lc_dtl_ext2_amt = 32
    Const M474_EG1_E4_m_lc_dtl_ext3_amt = 33
    Const M474_EG1_E4_m_lc_dtl_ext2_cd = 34
    Const M474_EG1_E4_m_lc_dtl_ext3_cd = 35
    Const M474_EG1_E4_m_lc_dtl_ext1_rt = 36
    Const M474_EG1_E4_m_lc_dtl_ext2_rt = 37
    Const M474_EG1_E4_m_lc_dtl_ext3_rt = 38
    Const M474_EG1_E4_m_lc_dtl_ext1_dt = 39
    Const M474_EG1_E4_m_lc_dtl_ext2_dt = 40
    Const M474_EG1_E4_m_lc_dtl_ext3_dt = 41
    '[CONVERSION INFORMATION]  View Name : export_item b_item
    Const M474_EG1_E5_b_item_item_cd = 42
    Const M474_EG1_E5_b_item_item_nm = 43
    Const M474_EG1_E5_b_item_spec = 44
    Const M474_EG1_E5_b_item_item_acct = 45
    '[CONVERSION INFORMATION]  View Name : export_item b_hs_code
    Const M474_EG1_E6_b_hs_code_hs_cd = 46
    Const M474_EG1_E6_b_hs_code_hs_nm = 47
    '[CONVERSION INFORMATION]  View Name : export_item b_plant
    Const M474_EG1_E7_b_plant_plant_cd = 48
    Const M474_EG1_E7_b_plant_plant_nm = 49
    
    Const C_SHEETMAXROWS_D  = 100
    
    Dim str_txtLCNo
    On Error Resume Next
	Err.Clear 

    Set iPM4G128 = Server.CreateObject("PM4G128.cMListLcDtlS")

    If CheckSYSTEMError(Err, True) = True Then
        Set iPM4G128 = Nothing
        Exit Sub
    End If
    str_txtLCNo =  Trim(Request("txtLCNo"))
    Call iPM4G128.M_LIST_LC_DTL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, CSTR(I1_m_lc_dtl_lc_seq), str_txtLCNo, E1_m_lc_dtl_total_amt, _
          E2_m_lc_dtl_max_lc_seq, E3_m_lc_dtl_next_lc_seq, EG1_export_group)

    If CheckSYSTEMError2(Err, True, "", "", "", "", "") = True Then
        Set iPM4G128 = Nothing                                              '☜: ComProxy Unload
        Exit Sub
    End If

    Set iPM4G128 = Nothing

        Response.Write "<Script Language=VBScript>" & vbCr
        Response.Write "parent.DbJumpQueryOK()" & vbCr
        Response.Write "</Script>" & vbCr
End Sub

Sub SubLookupDailyExRt
	On Error Resume Next
    Err.Clear
    
	Dim iPB0C004
	Dim E_B_Daily_Exchange_Rate
		Const B253_E1_std_rate = 0
		Const B253_E1_multi_divide = 1
	Dim str_txtCurrency
	
    Set iPB0C004 = CreateObject("PB0C004.CB0C004")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPB0C004 = Nothing												'☜: ComPlus Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	str_txtCurrency = 	Request("txtCurrency")
    E_B_Daily_Exchange_Rate = iPB0C004.B_SELECT_EXCHANGE_RATE(gStrGlobalCollection, str_txtCurrency, gCurrency, UNIConvDate(Request("txtOpenDt")))

	If CheckSYSTEMError2(Err,True, ,"","","","") = true then 		
		Set iPB0C004 = Nothing												'☜: ComPlus Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	Set iPB0C004 = Nothing
	
	Response.Write "<Script Language=VBScript>" 	& vbCr
	Response.Write "	With parent.frm1" 			& vbCr
	Response.Write "IF " & Trim(E_B_Daily_Exchange_Rate(B253_E1_std_rate)) & " <> 0 THEN " & vbCr			
	Response.Write "	.hdnXchRtOp.value = """ & ConvSPChars(E_B_Daily_Exchange_Rate(B253_E1_multi_divide)) & """" & vbCr
	Response.Write "	.txtXchRate.value = """ & UNINumClientFormat(E_B_Daily_Exchange_Rate(B253_E1_std_rate), ggExchRate.DecPoint, 0)    & """" & vbCr
	Response.Write "ElSE    " & vbCr
	Response.Write "	.txtXchRate.value = 0 " & vbCr
	Response.Write "END IF	" & vbCr
	Response.Write "	End With"  						& vbCr
	Response.Write "</Script>"  						& vbCr


End Sub	    
%>
