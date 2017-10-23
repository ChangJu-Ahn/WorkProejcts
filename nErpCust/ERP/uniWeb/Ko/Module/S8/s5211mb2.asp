<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution                                                      *
'*  2. Function Name        :                                                                           *
'*  3. Program ID           : S5211MB2                                                                  *
'*  4. Program Name         :                                                                           *
'*  5. Program Desc         : 수출 B/L등록																*
'*  6. Comproxy List        : PS3G102.cLookupSoHdrSvr,B5CS41.cLookupBizPartnerSvr			            *
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/11/15																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Ahn Tae Hee												                *
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : 화면 design												*
'*							  2. 2000/04/18 : Coding Start												*
'*                            3. 2002/11/15 : UI 표준적용												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim strMode
                                                                                     '☜: Clear Error status
    Call HideStatusWnd                                                               '☜: Hide Processing message

    strMode      = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
 
 	Dim iS3G102
	Dim iCommandSent
	Dim I1_s_so_hdr
	Dim E1_s_so_hdr
    
    Dim iS1C141
	Dim imp_biz_partner_cd
    Dim E1_b_biz_partner
    Const S308_E1_so_no = 0
    Const S308_E1_so_dt = 1
    Const S308_E1_req_dlvy_dt = 2
    Const S308_E1_cfm_flag = 3
    Const S308_E1_price_flag = 4
    Const S308_E1_cur = 5
    Const S308_E1_xchg_rate = 6
    Const S308_E1_net_amt = 7
    Const S308_E1_net_amt_loc = 8
    Const S308_E1_cust_po_no = 9
    Const S308_E1_cust_po_dt = 10
    Const S308_E1_sales_cost_center = 11
    Const S308_E1_deal_type = 12
    Const S308_E1_pay_meth = 13
    Const S308_E1_pay_dur = 14
    Const S308_E1_trans_meth = 15
    Const S308_E1_vat_inc_flag = 16
    Const S308_E1_vat_type = 17
    Const S308_E1_vat_rate = 18
    Const S308_E1_vat_amt = 19
    Const S308_E1_vat_amt_loc = 20
    Const S308_E1_origin_cd = 21
    Const S308_E1_valid_dt = 22
    Const S308_E1_contract_dt = 23
    Const S308_E1_ship_dt_txt = 24
    Const S308_E1_pack_cond = 25
    Const S308_E1_inspect_meth = 26
    Const S308_E1_incoterms = 27
    Const S308_E1_dischge_city = 28
    Const S308_E1_dischge_port_cd = 29
    Const S308_E1_loading_port_cd = 30
    Const S308_E1_beneficiary = 31
    Const S308_E1_manufacturer = 32
    Const S308_E1_agent = 33
    Const S308_E1_remark = 34
    Const S308_E1_pre_doc_no = 35
    Const S308_E1_lc_flag = 36
    Const S308_E1_rel_dn_flag = 37
    Const S308_E1_rel_bill_flag = 38
    Const S308_E1_ret_item_flag = 39
    Const S308_E1_sp_stk_flag = 40
    Const S308_E1_ci_flag = 41
    Const S308_E1_export_flag = 42
    Const S308_E1_so_sts = 43
    Const S308_E1_insrt_user_id = 44
    Const S308_E1_insrt_dt = 45
    Const S308_E1_updt_user_id = 46
    Const S308_E1_updt_dt = 47
    Const S308_E1_ext1_qty = 48
    Const S308_E1_ext2_qty = 49
    Const S308_E1_ext3_qty = 50
    Const S308_E1_ext1_amt = 51
    Const S308_E1_ext2_amt = 52
    Const S308_E1_ext3_amt = 53
    Const S308_E1_ext1_cd = 54
    Const S308_E1_maint_no = 55
    Const S308_E1_ext3_cd = 56
    Const S308_E1_pay_type = 57
    Const S308_E1_pay_terms_txt = 58
    Const S308_E1_dn_parcel_flag = 59
    Const S308_E1_to_biz_area = 60
    Const S308_E1_to_biz_grp = 61
    Const S308_E1_biz_area = 62
    Const S308_E1_to_biz_org = 63
    Const S308_E1_to_biz_cost_center = 64
    Const S308_E1_ship_dt = 65
    Const S308_E1_auto_dn_flag = 66
    Const S308_E1_ext2_cd = 67
    Const S308_E1_bank_cd = 68
    Const S308_E1_sales_grp = 69
    Const S308_E1_sales_grp_nm = 70
    Const S308_E1_so_type = 71
    Const S308_E1_so_type_nm = 72
    Const S308_E1_bill_to_party = 73
    Const S308_E1_bill_to_party_type = 74
    Const S308_E1_bill_to_party_nm = 75
    Const S308_E1_ship_to_party = 76
    Const S308_E1_ship_to_party_type = 77
    Const S308_E1_ship_to_party_nm = 78
    Const S308_E1_sold_to_party = 79
    Const S308_E1_sold_to_party_type = 80
    Const S308_E1_sold_to_party_nm = 81
    Const S308_E1_payer = 82
    Const S308_E1_payer_type = 83
    Const S308_E1_payer_nm = 84
    Const S308_E1_sales_org = 85
    Const S308_E1_sales_org_nm = 86
    Const S308_E1_bank_nm = 87
    Const S308_E1_deal_type_nm = 88
    Const S308_E1_vat_type_nm = 89
    Const S308_E1_pay_meth_nm = 90
    Const S308_E1_incoterms_nm = 91
    Const S308_E1_pack_cond_nm = 92
    Const S308_E1_inspect_meth_nm = 93
    Const S308_E1_trans_meth_nm = 94
    Const S308_E1_vat_inc_flag_nm = 95
    Const S308_E1_pay_type_nm = 96
    Const S308_E1_loading_port_nm = 97
    Const S308_E1_dischge_port_nm = 98
    Const S308_E1_origin_nm = 99
    Const S308_E1_manufacturer_nm = 100
    Const S308_E1_agent_nm = 101
    Const S308_E1_beneficiary_nm = 102
    Const S308_E1_currency_desc = 103
    Const S308_E1_biz_area_nm = 104
    Const S308_E1_to_biz_grp_nm = 105
    'iS1C141
    Const S074_E1_credit_rot_day = 53
	On Error Resume Next
	Err.Clear                               
	iCommandSent = "QUERY"
    I1_s_so_hdr =  Trim(Request("txtSONo"))
    
    Set iS3G102 = Server.CreateObject("PS3G102.cLookupSoHdrSvr")
 
	If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End If
 
	Call iS3G102.S_LOOKUP_SO_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_so_hdr, E1_s_so_hdr)
											
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G102 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.End 
	End If

	Set iS3G102 = Nothing
    imp_biz_partner_cd = E1_s_so_hdr(S308_E1_sold_to_party) 
    iCommandSent = "LOOKUP"
    Set iS1C141 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If     

	Call iS1C141.B_LOOKUP_BIZ_PARTNER_SVR (gStrGlobalCollection, iCommandSent, imp_biz_partner_cd, E1_b_biz_partner)           									 
 								 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iS1C141 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End 
    End If      
 
    Set iS1C141 = Nothing   
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent.frm1"			& vbCr

	Dim strDt

'	Response.Write ".txtXchRate.text			= """ & UNINumClientFormat(E1_s_so_hdr(S308_E1_xchg_rate), ggExchRate.DecPoint, 0) & """" & vbCr 
	Response.Write ".txtLoadingDt.text			= """ & UNIDateClientFormat(E1_s_so_hdr(S308_E1_ship_dt)) & """" & vbCr
			
	Response.Write ".txtCurrency.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_cur)) & """" & vbCr
	Response.Write ".txtCurrency1.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_cur)) & """" & vbCr
	Response.Write " parent.CurFormatNumericOCX " & vbCr 
	Response.Write ".txtTransport.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_trans_meth)) & """" & vbCr
	Response.Write ".txtTransportNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_trans_meth_nm)) & """" & vbCr
	Response.Write ".txtApplicant.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_sold_to_party)) & """" & vbCr
	Response.Write ".txtApplicantNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_sold_to_party_nm)) & """" & vbCr
	Response.Write ".txtLoadingPort.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_loading_port_cd)) & """" & vbCr
	Response.Write ".txtLoadingPortNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_loading_port_nm)) & """" & vbCr
	Response.Write ".txtIncoterms.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_incoterms)) & """" & vbCr
	Response.Write ".txtIncotermsNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_incoterms_nm)) & """" & vbCr
	Response.Write ".txtDischgePort.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_dischge_port_cd)) & """" & vbCr
	Response.Write ".txtDischgePortNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_dischge_port_cd)) & """" & vbCr
	Response.Write ".txtSalesGroup.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_sales_grp)) & """" & vbCr
	Response.Write ".txtSalesGroupNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_sales_grp_nm)) & """" & vbCr
	Response.Write ".txtBeneficiary.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_beneficiary)) & """" & vbCr
	Response.Write ".txtBeneficiaryNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_beneficiary_nm)) & """" & vbCr
	Response.Write ".txtAgent.value				= """ & ConvSPChars(E1_s_so_hdr(S308_E1_agent)) & """" & vbCr
	Response.Write ".txtAgentNm.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_agent_nm)) & """" & vbCr
	Response.Write ".txtManufacturer.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_manufacturer)) & """" & vbCr
	Response.Write ".txtManufacturerNm.value	= """ & ConvSPChars(E1_s_so_hdr(S308_E1_manufacturer_nm)) & """" & vbCr
	Response.Write ".txtPackingType.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_pack_cond)) & """" & vbCr
	Response.Write ".txtPackingTypeNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_pack_cond_nm)) & """" & vbCr
	Response.Write ".txtOrigin.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_origin_cd)) & """" & vbCr
	Response.Write ".txtOriginNm.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_origin_nm)) & """" & vbCr
	Response.Write ".txtBillToParty.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_bill_to_party)) & """" & vbCr
	Response.Write ".txtBillToPartyNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_bill_to_party_nm)) & """" & vbCr
	Response.Write ".txtPayTerms.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_pay_meth)) & """" & vbCr
	Response.Write ".txtPayTermsNm.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_pay_meth_nm)) & """" & vbCr
	Response.Write ".txtPayDur.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_pay_dur)) & """" & vbCr
	Response.Write ".txtPayer.value				= """ & ConvSPChars(E1_s_so_hdr(S308_E1_payer)) & """" & vbCr
	Response.Write ".txtPayerNm.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_payer_nm)) & """" & vbCr
	Response.Write ".txtToSalesGroup.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_to_biz_grp)) & """" & vbCr
	Response.Write ".txtToSalesGroupNm.value	= """ & ConvSPChars(E1_s_so_hdr(S308_E1_to_biz_grp_nm)) & """" & vbCr
	Response.Write ".txtRefFlg.value			= ""S"" " & vbCr

	Response.Write ".txtVatIncflag.value		= """ & ConvSPChars(E1_s_so_hdr(S308_E1_vat_inc_flag)) & """" & vbCr
	Response.Write ".txtVatType.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_vat_type)) & """" & vbCr
	Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(E1_s_so_hdr(S308_E1_vat_type_nm)) & """" & vbCr				
	Response.Write ".txtVATRate.value			= """ & UNINumClientFormat(E1_s_so_hdr(S308_E1_vat_rate), ggExchRate.DecPoint, 0) & """" & vbCr
		
		'약정회전일 
	Response.Write ".txtCreditRot.value = """ & E1_b_biz_partner(S074_E1_credit_rot_day) & """" & vbCr	

	Response.Write "Call parent.ReferenceQueryOk() "    & vbCr
	
	Response.Write "End With                  "    & vbCr
    Response.Write "</Script>                 "     & vbCr 
 
 %>