<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3221mb3.asp																*
'*  4. Program Name         : 데이터가져오기(L/C Amend등록 L/C참조에서)									*
'*  5. Program Desc         : 데이터가져오기(L/C Amend등록 L/C참조에서)									*
'*  7. Modified date(First) : 2000/04/10																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Seo Jinkyung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/10 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")   
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
Call HideStatusWnd          

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 

Select Case strMode
	Case CStr(UID_M0001)														
		Dim S32119																' Master L/C Header 조회용 Object
		Dim iCommandSent 
		
		Dim I1_s_lc_hdr 
		
		Dim E1_b_biz_partner 
		Dim E2_b_bank  
		Dim E3_b_bank 
		Dim E4_b_bank  
		Dim E5_b_bank  
		Dim E6_b_bank 
		Dim E7_b_sales_grp  
		Dim E8_b_sales_org  
		Dim E9_b_biz_partner 
		Dim E10_b_biz_partner  
		Dim E11_b_biz_partner  
		Dim E12_b_biz_partner 
		Dim E13_b_biz_partner  
		Dim E14_b_minor  
		Dim E15_b_minor 
		Dim E16_b_minor  
		Dim E17_b_minor  
		Dim E18_b_minor 
		Dim E19_b_minor  
		Dim E20_b_minor  
		Dim E21_b_country 
		Dim E22_b_minor  
		Dim E23_b_minor  
		Dim E24_b_minor 
		Dim E25_b_minor  
		Dim E26_s_lc_hdr 


		Const S357_I1_lc_no = 0    'imp s_lc_hdr
		Const S357_I1_lc_kind = 1
		    
		Const S357_E1_bp_nm = 0    'exp_consignee b_biz_partner
		    
		Const S357_E2_bank_cd = 0    'exp_issue_bank b_bank
		Const S357_E2_bank_nm = 1
		    
		Const S357_E3_bank_cd = 0    'exp_advise_bank b_bank
		Const S357_E3_bank_nm = 1
		    
		Const S357_E4_bank_cd = 0    'exp_renego_bank b_bank
		Const S357_E4_bank_nm = 1

		Const S357_E5_bank_cd = 0    'exp_pay_bank b_bank
		Const S357_E5_bank_nm = 1

		Const S357_E6_bank_cd = 0    'exp_confirm_bank b_bank
		Const S357_E6_bank_nm = 1

		Const S357_E7_sales_grp_nm = 0    'exp b_sales_grp
		Const S357_E7_sales_grp = 1

		Const S357_E8_sales_org_nm = 0    'exp b_sales_org
		Const S357_E8_sales_org = 1

		Const S357_E9_bp_nm = 0    'exp_beneficiary b_biz_partner
		Const S357_E9_bp_cd = 1

		Const S357_E10_bp_nm = 0    'exp_applicant b_biz_partner
		Const S357_E10_bp_cd = 1

		Const S357_E11_bp_nm = 0    'exp_agent b_biz_partner

		Const S357_E12_bp_nm = 0    'exp_manufacturer b_biz_partner

		Const S357_E13_bp_nm = 0    'exp_notify_party b_biz_partner

		Const S357_E14_minor_nm = 0    'exp_incoterms_nm b_minor

		Const S357_E15_minor_nm = 0    'exp_pay_meth_nm b_minor

		Const S357_E16_minor_nm = 0    'exp_lc_type_nm b_minor

		Const S357_E17_minor_nm = 0    'exp_loading_port_nm b_minor
		    
		Const S357_E18_minor_nm = 0    'exp_discharge_port_nm b_minor
		    
		Const S357_E19_minor_nm = 0    'exp_transport_nm b_minor

		Const S357_E20_minor_nm = 0    'exp_origin_nm b_minor

		Const S357_E21_country_nm = 0    'exp_origin_cntry_nm b_country    
		    
		Const S357_E22_minor_nm = 0    'exp_charge_cd_nm b_minor

		Const S357_E23_minor_nm = 0    'exp_credit_core_nm b_minor

		Const S357_E24_minor_nm = 0    'exp_freight_nm b_minor

		Const S357_E25_minor_nm = 0    'exp_llc_type_nm b_minor

		Const S357_E26_lc_no = 0    'exp s_lc_hdr
		Const S357_E26_lc_doc_no = 1
		Const S357_E26_lc_amend_seq = 2
		Const S357_E26_so_no = 3
		Const S357_E26_adv_no = 4
		Const S357_E26_pre_adv_ref = 5
		Const S357_E26_adv_dt = 6
		Const S357_E26_open_dt = 7
		Const S357_E26_expiry_dt = 8
		Const S357_E26_amend_dt = 9
		Const S357_E26_manufacturer = 10
		Const S357_E26_agent = 11
		Const S357_E26_cur = 12
		Const S357_E26_lc_amt = 13
		Const S357_E26_xch_rate = 14
		Const S357_E26_lc_loc_amt = 15
		Const S357_E26_bank_txt = 16
		Const S357_E26_incoterms = 17
		Const S357_E26_pay_meth = 18
		Const S357_E26_payment_txt = 19
		Const S357_E26_latest_ship_dt = 20
		Const S357_E26_shipment = 21
		Const S357_E26_doc1 = 22
		Const S357_E26_doc2 = 23
		Const S357_E26_doc3 = 24
		Const S357_E26_doc4 = 25
		Const S357_E26_doc5 = 26
		Const S357_E26_file_dt = 27
		Const S357_E26_file_dt_txt = 28
		Const S357_E26_remark = 29
		Const S357_E26_lc_kind = 30
		Const S357_E26_lc_type = 31
		Const S357_E26_delivery_plce = 32
		Const S357_E26_amt_tolerance = 33
		Const S357_E26_loading_port = 34
		Const S357_E26_dischge_port = 35
		Const S357_E26_transport = 36
		Const S357_E26_transport_comp = 37
		Const S357_E26_origin = 38
		Const S357_E26_origin_cntry = 39
		Const S357_E26_charge_txt = 40
		Const S357_E26_charge_cd = 41
		Const S357_E26_credit_core = 42
		Const S357_E26_inv_cnt = 43
		Const S357_E26_bl_awb_flg = 44
		Const S357_E26_freight = 45
		Const S357_E26_notify_party = 46
		Const S357_E26_consignee = 47
		Const S357_E26_insur_policy = 48
		Const S357_E26_pack_list = 49
		Const S357_E26_l_lc_type = 50
		Const S357_E26_open_bank_txt = 51
		Const S357_E26_o_lc_doc_no = 52
		Const S357_E26_o_lc_amend_seq = 53
		Const S357_E26_o_lc_no = 54
		Const S357_E26_o_lc_expiry_dt = 55
		Const S357_E26_o_lc_loc_amt = 56
		Const S357_E26_o_lc_type = 57
		Const S357_E26_pay_dur = 58
		Const S357_E26_partial_ship_flag = 59
		Const S357_E26_biz_area = 60
		Const S357_E26_trnshp_flag = 61
		Const S357_E26_transfer_flag = 62
		Const S357_E26_cert_origin_flag = 63
		Const S357_E26_o_lc_amd_seq = 64
		Const S357_E26_sts = 65
		Const S357_E26_nego_amt = 66
		Const S357_E26_ext1_qty = 67
		Const S357_E26_ext2_qty = 68
		Const S357_E26_ext3_qty = 69
		Const S357_E26_ext1_amt = 70
		Const S357_E26_ext2_amt = 71
		Const S357_E26_ext3_qmt = 72
		Const S357_E26_ext1_cd = 73
		Const S357_E26_ext2_cd = 74
		Const S357_E26_ext3_cd = 75
		Const S357_E26_xch_rate_op = 76


		If Request("txtLCNo") = "" Then											
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
			Response.End
		End If
		
		
		
		'---------------------------------- L/C Header Data Query ----------------------------------

		Redim I1_s_lc_hdr(1)
				
		I1_s_lc_hdr(S357_I1_lc_no) = Trim(Request("txtLCNo"))	
		
		I1_s_lc_hdr(S357_I1_lc_kind) = "M"
		iCommandSent = "LOOKUP"
		
		
		Set S32119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")

        If CheckSYSTEMError(Err, True) = True Then
            Response.End
        End If
        
        Call S32119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_lc_hdr, _
				E1_b_biz_partner, E2_b_bank , E3_b_bank, E4_b_bank , _
				E5_b_bank , E6_b_bank, E7_b_sales_grp , E8_b_sales_org , _
				E9_b_biz_partner, E10_b_biz_partner , E11_b_biz_partner , E12_b_biz_partner, _
				E13_b_biz_partner , E14_b_minor , E15_b_minor, E16_b_minor , _
				E17_b_minor , E18_b_minor, E19_b_minor , E20_b_minor , _
				E21_b_country, E22_b_minor , E23_b_minor , E24_b_minor, _
				E25_b_minor , E26_s_lc_hdr )

        If CheckSYSTEMError(Err, True) = True Then
            Set S32119 = Nothing
            Response.End
        End If        		

%>
<Script Language=VBScript>
	With parent.frm1
		Dim strDt		
		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
		.txtAtCurrency.value			= "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
		.txtBeCurrency.value			= "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
		parent.CurFormatNumericOCX
		'##########################

		'Tab 1 : L/C Amend 정보 
		
		.txtLCNo.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_no))%>"
		.txtLCDocNo.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_doc_no))%>"
		.txtLCAmendSeq.value = "<%=Cint(E26_s_lc_hdr(S357_E26_lc_amend_seq)) + 1%>"


		.txtAmendDt.text = "<%=UNIConvDateAtoB(GetSvrDate, gServerDateFormat, gDateFormat)%>"
		.txtApplicant.value = "<%=ConvSPChars(E10_b_biz_partner(S357_E10_bp_cd))%>"
		.txtApplicantNm.value = "<%=ConvSPChars(E10_b_biz_partner(S357_E10_bp_nm))%>"		
		.txtBeneficiary.value = "<%=ConvSPChars(E9_b_biz_partner(S357_E9_bp_cd))%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(E9_b_biz_partner(S357_E9_bp_nm))%>"
		
		.txtBeCurrency.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
		.txtAtCurrency.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
		.txtAtDocAmt.text	= "<%=UNINumClientFormatByCurrency(E26_s_lc_hdr(S357_E26_lc_amt), E26_s_lc_hdr(S357_E26_cur), ggAmtOfMoneyNo)%>"
		.txtBeDocAmt.text	= "<%=UNINumClientFormatByCurrency(E26_s_lc_hdr(S357_E26_lc_amt), E26_s_lc_hdr(S357_E26_cur), ggAmtOfMoneyNo)%>"


		.txtAtXchrate.text = "<%=UNINumClientFormat(E26_s_lc_hdr(S357_E26_xch_rate), ggExchRate.DecPoint, 0)%>"
		.txtBeXchrate.text = "<%=UNINumClientFormat(E26_s_lc_hdr(S357_E26_xch_rate), ggExchRate.DecPoint, 0)%>"
		
		.txtAtExpireDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_expiry_dt))%>"

		.txtAtLatestShipDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_latest_ship_dt))%>"

		.txtBeExpireDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_expiry_dt))%>"

		.txtBeLatestShipDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_latest_ship_dt))%>"


		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_trnshp_flag))%>" = "Y" Then
			.rdoAtTranshipment1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_trnshp_flag))%>" = "N" Then
			.rdoAtTranshipment2.Checked = True
		End If

		.txtBeTranshipment.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_trnshp_flag))%>"
		
		If "<%=ConvSPChars(E26_s_lc_hdr(ExpSLcHdrPartialShipFlag))%>" = "Y" Then
			.rdoAtPartialShip1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(ExpSLcHdrPartialShipFlag))%>" = "N" Then
			.rdoAtPartialShip2.Checked = True
		End If
		
		.txtBePartialShip.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_partial_ship_flag))%>"
		
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transfer_flag))%>" = "Y" Then
			.rdoAtTransfer1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transfer_flag))%>" = "N" Then
			.rdoAtTransfer2.Checked = True
		End If
		
		.txtBeTransfer.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transfer_flag))%>"
		.txtAtTransport.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transport))%>"
		.txtAtTransportNm.value = "<%=ConvSPChars(E19_b_minor(S357_E19_minor_nm))%>"
		
		.txtBeTransport.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transport))%>"
		.txtBeTransportNm.value = "<%=ConvSPChars(E19_b_minor(S357_E19_minor_nm))%>"
		.txtAtLoadingPort.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_loading_port))%>"
		.txtAtLoadingPortNm.value = "<%=ConvSPChars(E17_b_minor(S357_E17_minor_nm))%>"
		
		.txtBeLoadingPort.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_loading_port))%>"
		.txtBeLoadingPortNm.value = "<%=ConvSPChars(E17_b_minor(S357_E17_minor_nm))%>"
		.txtAtDischgePort.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_dischge_port))%>"
		.txtAtDischgePortNm.value = "<%=ConvSPChars(E18_b_minor(S357_E18_minor_nm))%>"
		.txtBeDischgePort.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_dischge_port))%>"
		.txtBeDischgePortNm.value = "<%=ConvSPChars(E18_b_minor(S357_E18_minor_nm))%>"
		
		'Tab 2 : L/C Amend 기타 
		.txtAdvBank.value = "<%=ConvSPChars(E3_b_bank(S357_E3_bank_cd))%>"
		.txtAdvBankNm.value = "<%=ConvSPChars(E3_b_bank (S357_E3_bank_nm))%>"
		.txtSalesGroup.value = "<%=ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp))%>"
		.txtSalesGroupNm.value = "<%=ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp_nm))%>"
		.txtOpenBank.value = "<%=ConvSPChars(E2_b_bank(S357_E2_bank_cd))%>"
		.txtOpenBankNm.value = "<%=ConvSPChars(E2_b_bank(S357_E2_bank_nm))%>"
		
		.txtOpenDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_open_dt))%>"
	
		.txtManufacturer.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_manufacturer))%>"
		.txtManufacturerNm.value = "<%=ConvSPChars(E12_b_biz_partner(S357_E12_bp_nm))%>"
		.txtAgent.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_agent))%>"
		.txtAgentNm.value = "<%=ConvSPChars(E11_b_biz_partner(S357_E11_bp_nm))%>"
		.txtDoc1.value = "<%=""%>"
		
		Call parent.LCQueryOk()
	
	End With
</Script>
<%
		Set S32119 = Nothing														'☜: Unload Comproxy

		Response.End																'☜: Process End

	Case Else
		Response.End
End Select
%>