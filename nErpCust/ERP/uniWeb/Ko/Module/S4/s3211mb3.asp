<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211mb3.asp																*
'*  4. Program Name         : 데이터 가져오기(L/C등록 수주참조에서)										*
'*  5. Program Desc         : 데이터 가져오기(L/C등록 수주참조에서)										*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
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

Dim PB0C004
Dim strExchRateOp
Dim l_strExchRateOp
Dim iS3G102
Dim iCommandSent
Dim I1_s_so_hdr
Dim E1_s_so_hdr

Const B253_E1_std_rate = 0
const B253_E1_multi_divide = 1
	
Const E1_so_no = 0
Const E1_so_dt = 1
Const E1_req_dlvy_dt = 2
Const E1_cfm_flag = 3
Const E1_price_flag = 4
Const E1_cur = 5
Const E1_xchg_rate = 6
Const E1_net_amt = 7
Const E1_net_amt_loc = 8
Const E1_cust_po_no = 9
Const E1_cust_po_dt = 10
Const E1_sales_cost_center = 11
Const E1_deal_type = 12
Const E1_pay_meth = 13
Const E1_pay_dur = 14
Const E1_trans_meth = 15
Const E1_vat_inc_flag = 16
Const E1_vat_type = 17
Const E1_vat_rate = 18
Const E1_vat_amt = 19
Const E1_vat_amt_loc = 20
Const E1_origin_cd = 21
Const E1_valid_dt = 22
Const E1_contract_dt = 23
Const E1_ship_dt_txt = 24
Const E1_pack_cond = 25
Const E1_inspect_meth = 26
Const E1_incoterms = 27
Const E1_dischge_city = 28
Const E1_dischge_port_cd = 29
Const E1_loading_port_cd = 30
Const E1_beneficiary = 31
Const E1_manufacturer = 32
Const E1_agent = 33
Const E1_remark = 34
Const E1_pre_doc_no = 35
Const E1_lc_flag = 36
Const E1_rel_dn_flag = 37
Const E1_rel_bill_flag = 38
Const E1_ret_item_flag = 39
Const E1_sp_stk_flag = 40
Const E1_ci_flag = 41
Const E1_export_flag = 42
Const E1_so_sts = 43
Const E1_insrt_user_id = 44
Const E1_insrt_dt = 45
Const E1_updt_user_id = 46
Const E1_updt_dt = 47
Const E1_ext1_qty = 48
Const E1_ext2_qty = 49
Const E1_ext3_qty = 50
Const E1_ext1_amt = 51
Const E1_ext2_amt = 52
Const E1_ext3_amt = 53
Const E1_ext1_cd = 54
Const E1_maint_no = 55
Const E1_ext3_cd = 56
Const E1_pay_type = 57
Const E1_pay_terms_txt = 58
Const E1_dn_parcel_flag = 59
Const E1_to_biz_area = 60
Const E1_to_biz_grp = 61
Const E1_biz_area = 62
Const E1_to_biz_org = 63
Const E1_to_biz_cost_center = 64
Const E1_ship_dt = 65
Const E1_auto_dn_flag = 66
Const E1_ext2_cd = 67
Const E1_bank_cd = 68
Const E1_sales_grp = 69
Const E1_sales_grp_nm = 70
Const E1_so_type = 71
Const E1_so_type_nm = 72
Const E1_bp_cd1 = 73
Const E1_bp_type1 = 74
Const E1_bp_nm1 = 75
Const E1_bp_cd2 = 76
Const E1_bp_type2 = 77
Const E1_bp_nm2 = 78
Const E1_bp_cd3 = 79
Const E1_bp_type3 = 80
Const E1_bp_nm3 = 81
Const E1_bp_cd4 = 82
Const E1_bp_type4 = 83
Const E1_bp_nm4 = 84
Const E1_sales_org = 85
Const E1_sales_org_nm = 86
Const E1_bank_nm = 87
Const E1_deal_type_nm = 88
Const E1_vat_type_nm = 89
Const E1_pay_meth_nm = 90
Const E1_incoterms_nm = 91
Const E1_pack_cond_nm = 92
Const E1_inspect_meth_nm = 93
Const E1_trans_meth_nm = 94
Const E1_vat_inc_flag_nm = 95
Const E1_pay_type_nm = 96
Const E1_loading_port_nm = 97
Const E1_dischge_port_nm = 98
Const E1_origin_nm = 99
Const E1_manufacturer_nm = 100
Const E1_agent_nm = 101
Const E1_beneficiary_nm = 102
Const E1_currency_desc = 103
Const E1_biz_area_nm = 104
Const E1_to_biz_grp_nm = 105
	
iCommandSent = "QUERY"
'2003-01-20 수주의 LookUp패드가 filtervar함수가 필요없음 
I1_s_so_hdr = Trim(Request("txtSONo"))

Set iS3G102 = Server.CreateObject("PS3G102.cLookupSoHdrSvr")
    
If CheckSYSTEMError(Err, True) = True Then
	Set iS3G102 = Nothing		                                                 
	Response.End
	
End If
	
Call iS3G102.S_LOOKUP_SO_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_so_hdr, E1_s_so_hdr)
											
If CheckSYSTEMError(Err, True) = True Then
	Set iS3G102 = Nothing		                                                 
	Response.End
	
End If
	
Set iS3G102 = Nothing

If E1_s_so_hdr(E1_cur) <> gCurrency Then
	
	Set PB0C004 = Server.CreateObject("PB0C004.CB0C004") 
				
	l_strExchRateOp = PB0C004.B_SELECT_EXCHANGE_RATE(gStrGlobalCollection,E1_s_so_hdr(E1_cur), _
	gCurrency,UNIConvDateCompanyToDB(null,E1_s_so_hdr(E1_so_dt)))
		
	If CheckSYSTEMError(Err,True) = True Then
		Set PB0C004 = Nothing
		Response.End
	End If  
		
	Set PB0C004 = Nothing  
		
	strExchRateOp = l_strExchRateOp(B253_E1_multi_divide)
			
Else
	strExchRateOp = "*"
End If	
   
%>
<Script Language=VBScript>
	Dim strDt
	
	With parent.frm1

		.txtCurrency.value			= "<%=E1_s_so_hdr(E1_cur)%>"
		parent.CurFormatNumericOCX

		.txtSalesGroup.value = "<%=ConvSPChars(E1_s_so_hdr(E1_sales_grp))%>"
		.txtSalesGroupNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_sales_grp_nm))%>"
		.txtCurrency.value = "<%=ConvSPChars(E1_s_so_hdr(E1_cur))%>"
		.txtXchRate.value = "<%=UNINumClientFormat(E1_s_so_hdr(E1_xchg_rate),ggExchRate.DecPoint,0)%>"
		
		If .txtXchRate.text = "" Then .txtXchRate.text = "0"
				
		.txtLatestShipDt.text = "<%=UNIDateClientFormat(E1_s_so_hdr(E1_ship_dt))%>"
				
		.txtIncoterms.value = "<%=ConvSPChars(E1_s_so_hdr(E1_incoterms))%>"
		.txtIncotermsNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_incoterms_nm))%>"
		.txtPayTerms.value = "<%=ConvSPChars(E1_s_so_hdr(E1_pay_meth))%>"
		.txtPayTermsNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_pay_meth_nm))%>"
		.txtPayDur.text = "<%=ConvSPChars(E1_s_so_hdr(E1_pay_dur))%>"
		.txtLoadingPort.value = "<%=ConvSPChars(E1_s_so_hdr(E1_loading_port_cd))%>"
		.txtLoadingPortNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_loading_port_nm))%>"
		.txtDischgePort.value = "<%=ConvSPChars(E1_s_so_hdr(E1_dischge_port_cd))%>"
		.txtDischgePortNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_dischge_port_nm))%>"
		.txtTransport.value = "<%=ConvSPChars(E1_s_so_hdr(E1_trans_meth))%>"
		.txtTransportNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_trans_meth_nm))%>"
		.txtApplicant.value = "<%=ConvSPChars(E1_s_so_hdr(E1_bp_cd3))%>"
		.txtApplicantNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_bp_nm3))%>"
		.txtOrigin.value = "<%=ConvSPChars(E1_s_so_hdr(E1_origin_cd))%>"
		.txtOriginNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_origin_nm))%>"
		.txtAgent.value = "<%=ConvSPChars(E1_s_so_hdr(E1_agent))%>"
		.txtAgentNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_agent_nm))%>"
		.txtBeneficiary.value = "<%=ConvSPChars(E1_s_so_hdr(E1_beneficiary))%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_beneficiary_nm))%>"
		.txtManufacturer.value = "<%=ConvSPChars(E1_s_so_hdr(E1_manufacturer))%>"
		.txtManufacturerNm.value = "<%=ConvSPChars(E1_s_so_hdr(E1_manufacturer_nm))%>"
		.txtShipment.value = "<%=ConvSPChars(E1_s_so_hdr(E1_ship_dt_txt))%>"

		.txtDocAmt.value	= "<%=UNINumClientFormatByCurrency(E1_s_so_hdr(E1_net_amt), E1_s_so_hdr(E1_cur), ggAmtOfMoneyNo)%>"
		
		If .txtDocAmt.text = "" Then .txtDocAmt.text = "0"
		
		.txtExchRateOp.value = "<%=strExchRateOp%>"
		
		Call parent.SOQueryOK()
	End With
</Script>
<%
Response.End																'☜: Process End
%>