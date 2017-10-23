<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3111MA1
'*  4. Program Name         : 수주등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : S31111MaintSoHdrSvr, S31119LookupSoHdrSvr
'*  7. Modified date(First) : 2000/04/09
'*  8. Modified date(Last)  : 2002/06/04
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Lee Myung Wha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/09 : ..........
'*                            -2000/05/09 : 표준수정사항적용 
'*                            -2000/09/04 : 4Th Coding
'*                            -2001/12/18 : Date 표준 적용 
'**********************************************************************************************

	Dim lgOpModeCRUD
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
	         Call SubBizQuery()
        Case CStr(UID_M0002)
			 Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case "LookUp"                                                                 '☜: Check	
             Call SubLookUp() 
        Case "RETURNSOQUERY"
			 Call SubBizQuery()
		Case "SoTypeExp"
			 Call SubSoTypeExp()   
		Case "DNCheck"
			 Call SubDNCheck()   	
		Case "btnCONFIRM"
			 Call SubbtnCONFIRM()   
		Case "CheckCreditlimit"	
			 Call CheckCreditlimit()
		Case "PROJECTQUERY"
			 Call SubProjectRef()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
	Dim iS3G102
	Dim iCommandSent
	Dim I1_s_so_hdr
	Dim E1_s_so_hdr
	
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
    Const E1_bill_to_party = 73
    Const E1_bill_to_party_type = 74
    Const E1_bill_to_party_nm = 75
    Const E1_ship_to_party = 76
    Const E1_ship_to_party_type = 77
    Const E1_ship_to_party_nm = 78
    Const E1_sold_to_party = 79
    Const E1_sold_to_party_type = 80
    Const E1_sold_to_party_nm = 81
    Const E1_payer = 82
    Const E1_payer_type = 83
    Const E1_payer_nm = 84
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
    Const E1_dn_req_flag  = 106
    
	On Error Resume Next
	Err.Clear                                                               '☜: Protect system from crashing
	
	iCommandSent = "QUERY"
    I1_s_so_hdr = Trim(Request("txtConSo_no"))

	Set iS3G102 = Server.CreateObject ("PS3G102.cLookupSoHdrSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G102 = Nothing	
		Exit Sub
	End If
	
	Call iS3G102.S_LOOKUP_SO_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_so_hdr, E1_s_so_hdr)
											
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G102 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script Language=vbscript>"			& vbCr   
        Response.Write "parent.frm1.txtConSo_no.focus"		& vbCr    
        Response.Write "</Script>      "					& vbCr	                 '☜: Unload Comproxy DLL
        Exit Sub
	End If
	
	Set iS3G102 = Nothing
	
    if lgOpModeCRUD = CStr(UID_M0001) then
    	'-----------------------
    	'Display result data
    	'----------------------- 
    	Response.Write "<Script Language=vbscript>" & vbCr
    	Response.Write "With parent.frm1"			& vbCr
    	
    	'--첫번째 TAB 
    	
    	'##### Rounding Logic #####
    	'항상 거래화폐가 우선 
    	Response.Write ".txtDoc_cur.value			= """ & ConvSPChars(E1_s_so_hdr(E1_cur))						& """" & vbCr
    	Response.write " parent.CurFormatNumericOCX "																& vbCr
    	'##########################
    		
    	Response.write ".txtSoNo.value				= """ & ConvSPChars(E1_s_so_hdr(E1_so_no))						& """" & vbCr
    	Response.write ".txtSo_Type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_so_type))					& """" & vbCr
    	Response.write ".txtSo_TypeNm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_so_type_nm))					& """" & vbCr
    	'수주일 
    	Response.write ".txtSo_dt.Text				= """ & UNIDateClientFormat(E1_s_so_hdr(E1_so_dt))				& """" & vbCr
    	'고객주문일 
    	Response.write ".txtCust_po_dt.Text			= """ & UNIDateClientFormat(E1_s_so_hdr(E1_cust_po_dt))			& """" & vbCr
    	'납기일 
    	Response.write ".txtReq_dlvy_dt.Text		= """ & UNIDateClientFormat(E1_s_so_hdr(E1_req_dlvy_dt))		& """" & vbCr
    	
    	Response.write ".txtSold_to_party.value		= """ & ConvSPChars(E1_s_so_hdr(E1_sold_to_party))				& """" & vbCr
    	Response.write ".txtSold_to_partyNm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_sold_to_party_nm))			& """" & vbCr
    	Response.write ".txtSales_Grp.value			= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp))					& """" & vbCr
    	Response.write ".txtSales_GrpNm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp_nm))				& """" & vbCr
    	Response.write ".txtBill_to_party.value		= """ & ConvSPChars(E1_s_so_hdr(E1_bill_to_party))				& """" & vbCr
    	Response.write ".txtBill_to_partyNm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_bill_to_party_nm))			& """" & vbCr
    	Response.write ".txtShip_to_party.value		= """ & ConvSPChars(E1_s_so_hdr(E1_Ship_to_party))				& """" & vbCr
    	Response.write ".txtShip_to_partyNm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_Ship_to_party_nm ))			& """" & vbCr
    	Response.write ".txtTo_Biz_Grp.value		= """ & ConvSPChars(E1_s_so_hdr(E1_to_biz_grp))					& """" & vbCr
    	Response.write ".txtTo_Biz_GrpNm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_to_biz_grp_nm))				& """" & vbCr
    	Response.write ".txtPayer.value				= """ & ConvSPChars(E1_s_so_hdr(E1_payer))						& """" & vbCr
    	Response.write ".txtPayerNm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_payer_nm))					& """" & vbCr
    	Response.write ".txtDeal_Type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_deal_type))					& """" & vbCr
    	Response.write ".txtDeal_Type_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_deal_type_nm))				& """" & vbCr
    	Response.write ".txtCust_po_no.value		= """ & ConvSPChars(E1_s_so_hdr(E1_cust_po_no))					& """" & vbCr
    	Response.write ".txtPay_terms.value			= """ & ConvSPChars(E1_s_so_hdr(E1_pay_meth))					& """" & vbCr
    	Response.write ".txtPay_terms_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_pay_meth_nm))				& """" & vbCr
    	
    	If Trim(E1_s_so_hdr(E1_pay_type)) <> "" Then
    	Response.write ".txtPay_type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_pay_type))					& """" & vbCr
    	End If
    	Response.write ".txtPay_Type_nm.value		= """ & ConvSPChars(Trim(E1_s_so_hdr(E1_pay_type_nm)))				& """" & vbCr
    	'결제기간 
    	Response.write ".txtPay_dur.Text			= """ & UNINumClientFormat(E1_s_so_hdr(E1_pay_dur),0,0)			& """" & vbCr
    	Response.write ".txtVat_Type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_vat_type))					& """" & vbCr
    	Response.write ".txtVatTypeNm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_vat_type_nm))				& """" & vbCr
    	'부가세율 
    	Response.write ".txtVat_rate.text			= """ & UNINumClientFormat(E1_s_so_hdr(E1_vat_rate),ggExchRate.DecPoint, 0)	& """" & vbCr
    	Response.write ".txtDoc_cur.value			= """ & ConvSPChars(E1_s_so_hdr(E1_cur))						& """" & vbCr
    	
    	If E1_s_so_hdr(E1_cur) = gCurrency Then 
    	Response.write "parent.ggoOper.SetReqAttr .txtXchg_rate,  """ & "Q" & """" 									& vbCr
    	Else 
    	Response.write "parent.ggoOper.SetReqAttr .txtXchg_rate,  """ & "N" & """"									& vbCr
    	End If 
    
    	'수주금액 
    	'##### Rounding Logic #####
    	Response.write ".txtNet_amt.text			= """ & UNINumClientFormatByCurrency(E1_s_so_hdr(E1_net_amt), E1_s_so_hdr(E1_cur), ggAmtOfMoneyNo)	& """" & vbCr
    	'##########################
    	Response.write ".txtVat_Inc_Flag.value		= """ & ConvSPChars(E1_s_so_hdr(E1_vat_inc_flag))				& """" & vbCr
    	Response.write ".txtVat_Inc_Flag_Nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_vat_inc_flag_nm))			& """" & vbCr
    	'환율 
    	Response.write ".txtXchg_rate.value			= """ & UNINumClientFormat(E1_s_so_hdr(E1_xchg_rate), ggExchRate.DecPoint,0)	& """" & vbCr
    	'부가세액 
    	Response.write ".txtVat_amt.value			= """ & UNIConvNumDBToCompanyByCurrency(E1_s_so_hdr(E1_vat_amt_loc), E1_s_so_hdr(E1_cur), ggAmtOfMoneyNo, gTaxRndPolicyNo, "X")	& """" & vbCr
    	'원화금액 
    	Response.write ".txtNet_Amt_Loc.Text		= """ & UNINumClientFormat(E1_s_so_hdr(E1_net_amt_loc), ggAmtOfMoney.DecPoint,0)	& """" & vbCr
    	Response.write ".txtTrans_Meth.value		= """ & ConvSPChars(E1_s_so_hdr(E1_trans_meth))					& """" & vbCr
    	Response.write ".txtTrans_Meth_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_trans_meth_nm))				& """" & vbCr
    	
    	'If Trim(E1_s_so_hdr(E1_pay_terms_txt)) <> "" Then
    	Response.write ".txt_Payterms_txt.value		= """ & Trim(ConvSPChars(E1_s_so_hdr(E1_pay_terms_txt)))				& """" & vbCr
    	'End If
    	Response.write ".txtRemark.value			= """ & ConvSPChars(E1_s_so_hdr(E1_remark))						& """" & vbCr
    
    	'--두번째 TAB 
    	Response.write ".txtManufacturer.value		= """ & ConvSPChars(E1_s_so_hdr(E1_manufacturer))				& """" & vbCr
    	Response.write ".txtManufacturer_nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_manufacturer_nm))			& """" & vbCr
    	Response.write ".txtAgent.value				= """ & ConvSPChars(E1_s_so_hdr(E1_agent))						& """" & vbCr
    	Response.write ".txtAgent_nm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_agent_nm))					& """" & vbCr
    	Response.write ".txtBeneficiary.value		= """ & ConvSPChars(E1_s_so_hdr(E1_beneficiary))				& """" & vbCr
    	Response.write ".txtBeneficiary_nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_beneficiary_nm))				& """" & vbCr
    	Response.write ".txtOrigin.value			= """ & ConvSPChars(E1_s_so_hdr(E1_origin_cd))					& """" & vbCr
    	Response.write ".txtOriginNm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_origin_nm))					& """" & vbCr
    	'계약일 
    	Response.write ".txtContract_dt.Text		= """ & UNIDateClientFormat(E1_s_so_hdr(E1_contract_dt))		& """" & vbCr
    
    	'유효일 
    	Response.write ".txtValid_dt.Text			= """ & UNIDateClientFormat(E1_s_so_hdr(E1_valid_dt))			& """" & vbCr
    
    	'선적일 
    	Response.write ".txtship_dt.Text			= """ & UNIDateClientFormat(E1_s_so_hdr(E1_ship_dt))			& """" & vbCr
    	
    	Response.write ".txtShip_dt_txt.value		= """ & ConvSPChars(E1_s_so_hdr(E1_ship_dt_txt))				& """" & vbCr
    	Response.write ".txtIncoTerms.value			= """ & ConvSPChars(E1_s_so_hdr(E1_incoterms))					& """" & vbCr
    	Response.write ".txtIncoTerms_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_incoterms_nm))				& """" & vbCr
    	Response.write ".txtSending_Bank.value		= """ & ConvSPChars(E1_s_so_hdr(E1_bank_cd))					& """" & vbCr
    	Response.write ".txtSending_Bank_nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_bank_nm))					& """" & vbCr
    	Response.write ".txtPack_cond.value			= """ & ConvSPChars(E1_s_so_hdr(E1_pack_cond))					& """" & vbCr
    	Response.write ".txtPack_cond_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_pack_cond_nm))				& """" & vbCr
    	Response.write ".txtInspect_meth.value		= """ & ConvSPChars(E1_s_so_hdr(E1_inspect_meth))				& """" & vbCr
    	Response.write ".txtInspect_meth_nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_inspect_meth_nm))			& """" & vbCr
    	Response.write ".txtDischge_city.value		= """ & ConvSPChars(E1_s_so_hdr(E1_dischge_city))				& """" & vbCr
    	Response.write ".txtDischge_port_Cd.value	= """ & ConvSPChars(E1_s_so_hdr(E1_dischge_port_cd))			& """" & vbCr
    	Response.write ".txtDischge_port_Nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_dischge_port_nm))			& """" & vbCr
    	Response.write ".txtLoading_port_Cd.value	= """ & ConvSPChars(E1_s_so_hdr(E1_loading_port_cd))			& """" & vbCr
    	Response.write ".txtLoading_port_Nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_loading_port_nm))			& """" & vbCr
    	'Response.write ".txtHSONo.value				= """ & ConvSPChars(pS31119.ExpSSoHdrPreDocNo)				& """" & vbCr
    
    	'주문서관리번호 
    	Response.write ".txtMaintNo.value	= """ & ConvSPChars(E1_s_so_hdr(E1_maint_no))							& """" & vbCr
    	Response.write ".txtSoSts.value		= """ & ConvSPChars(E1_s_so_hdr(E1_so_sts))								& """" & vbCr
    	Response.write ".txtRetItemFlag.value	= """ & ConvSPChars(E1_s_so_hdr(E1_ret_item_flag))					& """" & vbCr
    	
    	Response.write ".RdoConfirm.value = """ & "Y" & """"														& vbCr
    	Response.write ".btnConfirm.value = """ & "확정처리" & """"												& vbCr
    
    	If E1_s_so_hdr(E1_cfm_flag) = "Y" Then
    		Response.write ".rdoCfm_flag1.checked = True" 															& vbCr
    		Response.write ".txtRadioFlag.value = .rdoCfm_flag1.value" 												& vbCr
    		
    		'출하버튼 처리 
    		If cstr(E1_s_so_hdr(E1_auto_dn_flag)) = "N" Then  'Or cdbl(E1_s_so_hdr(E1_so_sts)) = 1 Then
    			Response.write ".btnDNCheck.value = """ & "출하요청처리" & """"									& vbCr
				Response.write ".btnDNCheck.disabled = True"														& vbCr
    			    			
    		ElseIf cstr(E1_s_so_hdr(E1_auto_dn_flag)) = "Y"  Then   'And E1_s_so_hdr(E1_so_sts) <> 1 Then
				If E1_s_so_hdr(E1_dn_req_flag) = "N" And cdbl(E1_s_so_hdr(E1_so_sts)) = 2 Then			
					Response.write ".RdoDnReq.value = """ & "N" & """"												& vbCr
					Response.write ".btnDNCheck.value = """ & "출하요청처리" & """"								& vbCr
					Response.write ".btnDNCheck.disabled = False"													& vbCr

				ElseIf E1_s_so_hdr(E1_dn_req_flag) = "N" And cdbl(E1_s_so_hdr(E1_so_sts)) = 1 Then
					Response.write ".btnDNCheck.value = """ & "출하요청취소" & """"								& vbCr
					Response.write ".btnDNCheck.disabled = True"													& vbCr

				ElseIf E1_s_so_hdr(E1_dn_req_flag) = "Y" And cdbl(E1_s_so_hdr(E1_so_sts)) = 1 Then
					Response.write ".RdoDnReq.value = """ & "Y" & """"												& vbCr
					Response.write ".btnDNCheck.value = """ & "출하요청취소" & """"								& vbCr	
					Response.write ".btnDNCheck.disabled = False"													& vbCr
				End If
    			
    		End IF
    
    		Response.write ".RdoConfirm.value = """ & "N" & """"													& vbCr
    		Response.write ".btnConfirm.value = """ & "확정취소" & """"											& vbCr
    
    	ElseIf E1_s_so_hdr(E1_cfm_flag) = "N" Then
    
    		Response.write ".rdoCfm_flag2.checked = True" 															& vbCr
    		Response.write ".txtRadioFlag.value = .rdoCfm_flag2.value" 												& vbCr  
    		
    		Response.write ".btnDNCheck.disabled = True"															& vbCr
    
    	End If
    
    	If E1_s_so_hdr(E1_price_flag) = "Y" Then
    		Response.write ".rdoPrice_flag1.checked = True" 														& vbCr
    		Response.write ".txtRadioType.value = .rdoPrice_flag1.value"											& vbCr
    	ElseIf E1_s_so_hdr(E1_price_flag) = "N" Then
    		Response.write ".rdoPrice_flag2.checked = True" 														& vbCr
    		Response.write ".txtRadioType.value = .rdoPrice_flag2.value	"											& vbCr
    	End IF
       	
    	
    	Response.Write " parent.DbQueryOk" & vbCr
    	Response.Write " End With"          & vbCr
        Response.Write " </Script>"         & vbCr
        
        
    '수주 참조시에 조회 
    ElseIf lgOpModeCRUD = "RETURNSOQUERY" then 
        
        Response.Write "<Script Language=vbscript>" & vbCr
    	Response.Write "With parent.frm1"			& vbCr
    	
    	
    	'--첫번째 TAB 
    	Response.write ".txtSold_to_party.value		= """ & ConvSPChars(E1_s_so_hdr(E1_sold_to_party))						& """" & vbCr
    	Response.write ".txtSold_to_partyNm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_sold_to_party_nm))						& """" & vbCr
    	Response.write ".txtSales_Grp.value			= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp))					& """" & vbCr
    	Response.write ".txtSales_GrpNm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp_nm))				& """" & vbCr
    	Response.write ".txtBill_to_party.value		= """ & ConvSPChars(E1_s_so_hdr(E1_bill_to_party))						& """" & vbCr
    	Response.write ".txtBill_to_partyNm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_bill_to_party_nm))						& """" & vbCr
    	Response.write ".txtShip_to_party.value		= """ & ConvSPChars(E1_s_so_hdr(E1_ship_to_party))						& """" & vbCr
    	Response.write ".txtShip_to_partyNm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_ship_to_party_nm))					& """" & vbCr
    	Response.write ".txtTo_Biz_Grp.value		= """ & ConvSPChars(E1_s_so_hdr(E1_to_biz_grp))					& """" & vbCr
    	Response.write ".txtTo_Biz_GrpNm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_to_biz_grp_nm))				& """" & vbCr
    	Response.write ".txtPayer.value				= """ & ConvSPChars(E1_s_so_hdr(E1_payer))						& """" & vbCr
    	Response.write ".txtPayerNm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_payer_nm))						& """" & vbCr
    	Response.write ".txtDeal_Type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_deal_type))					& """" & vbCr
    	Response.write ".txtDeal_Type_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_deal_type_nm))				& """" & vbCr
    	Response.write ".txtCust_po_no.value		= """ & ConvSPChars(E1_s_so_hdr(E1_cust_po_no))					& """" & vbCr
    	Response.write ".txtPay_terms.value			= """ & ConvSPChars(E1_s_so_hdr(E1_pay_meth))					& """" & vbCr
    	Response.write ".txtPay_terms_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_pay_meth_nm))				& """" & vbCr
    	Response.write ".txtPay_type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_pay_type))					& """" & vbCr
    	Response.write ".txtPay_Type_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_pay_type_nm))				& """" & vbCr
    	'결제기간 
    	Response.write ".txtPay_dur.Text			= """ & UNINumClientFormat(E1_s_so_hdr(E1_pay_dur),0,0)			& """" & vbCr
    	Response.write ".txtVat_Type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_vat_type))					& """" & vbCr
    	Response.write ".txtVatTypeNm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_vat_type_nm))				& """" & vbCr
    	'부가세율 
    	Response.write ".txtVat_rate.text			= """ & UNINumClientFormat(E1_s_so_hdr(E1_vat_rate),ggExchRate.DecPoint, 0)	& """" & vbCr
    	Response.write ".txtDoc_cur.value			= """ & ConvSPChars(E1_s_so_hdr(E1_cur))						& """" & vbCr
    	
    	If E1_s_so_hdr(E1_cur) = gCurrency Then 
    	Response.write "parent.ggoOper.SetReqAttr .txtXchg_rate,  """ & "Q" & """" 									& vbCr
    	Else 
    	Response.write "parent.ggoOper.SetReqAttr .txtXchg_rate,  """ & "N" & """"									& vbCr
    	End If 
    
    	Response.write ".txtVat_Inc_Flag.value		= """ & ConvSPChars(E1_s_so_hdr(E1_vat_inc_flag))				& """" & vbCr
    	Response.write ".txtVat_Inc_Flag_Nm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_vat_inc_flag_nm))			& """" & vbCr
    	'환율 
    	Response.write ".txtXchg_rate.value			= """ & UNINumClientFormat(E1_s_so_hdr(E1_xchg_rate), ggExchRate.DecPoint,0)	& """" & vbCr
    	Response.write ".txtTrans_Meth.value		= """ & ConvSPChars(E1_s_so_hdr(E1_trans_meth))					& """" & vbCr
    	Response.write ".txtTrans_Meth_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_trans_meth_nm))				& """" & vbCr
    	Response.write ".txt_Payterms_txt.value		= """ & Trim(ConvSPChars(E1_s_so_hdr(E1_pay_terms_txt)))				& """" & vbCr
    	Response.write ".txtRemark.value			= """ & ConvSPChars(E1_s_so_hdr(E1_remark))						& """" & vbCr
    	Response.write ".txtRetItemFlag.value	    = """ & ConvSPChars(E1_s_so_hdr(E1_ret_item_flag))					& """" & vbCr
    	
    	Response.Write "End With"			& vbCr
        Response.Write "</Script>"			& vbCr
        Response.End 
    End if
    	
End Sub

'============================================================================================================
Sub SubBizSave()
	
	Dim iCommandSent
	Dim lgIntFlgMode
    Dim I1_s_so_hdr
    Dim E1_s_so_hdr
    Dim S_So_Type_Config	
	Dim Sold_To_Party		
	Dim B_Sales_Grp		
	Dim Bill_To_Party		
	Dim Ship_To_Party		
	Dim To_Grp_B_Sales_Grp	
	Dim Payer_To_Party		
	Dim B_Bank	
	Dim strProjectCode ' 프로젝트코드 
	
	Dim iS3G101	
	
	Const so_no = 0
    Const so_dt = 1
    Const req_dlvy_dt = 2
    Const cfm_flag = 3
    Const price_flag = 4
    Const cur = 5
    Const cust_po_no = 6
    Const cust_po_dt = 7
    Const deal_type = 8
    Const pay_meth = 9
    Const pay_dur = 10
    Const trans_meth = 11
    Const vat_inc_flag = 12
    Const vat_type = 13
    Const vat_rate = 14
    Const origin_cd = 15
    Const valid_dt = 16
    Const contract_dt = 17
    Const ship_dt_txt = 18
    Const pack_cond = 19
    Const inspect_meth = 20
    Const incoterms = 21
    Const dischge_city = 22
    Const dischge_port_cd = 23
    Const loading_port_cd = 24
    Const beneficiary = 25
    Const manufacturer = 26
    Const agent = 27
    Const remark = 28
    Const pay_type = 29
    Const ship_dt = 30
    Const pay_terms_txt = 31
    Const dn_parcel_flag = 32
    Const xchg_rate = 33
    Const to_biz_area = 34
    Const maint_no = 35
    Const ext1_qty = 36
    Const ext2_qty = 37
    Const ext3_qty = 38
    Const ext1_amt = 39
    Const ext2_amt = 40
    Const ext3_amt = 41
    Const ext1_cd = 42
    Const ext2_cd = 43
    Const ext3_cd = 44
    Const pre_doc_no = 45
    Const sto_flag = 46
    
    On Error Resume Next
    Err.Clear 
    
    ReDim I1_s_so_hdr(46) 
    
    Dim pvCB
    Dim pvICustomXML
    Dim prOCustomXML
    
	'-----------------------
    'Data manipulate area
    '-----------------------
	
    I1_s_so_hdr(So_No) = UCase(Trim(Request("txtSoNo")))
    I1_s_so_hdr(So_Dt) = UNIConvDate(Trim(Request("txtSo_dt")))
	I1_s_so_hdr(Req_Dlvy_Dt) = UNIConvDate(Trim(Request("txtReq_dlvy_dt")))
    I1_s_so_hdr(Cfm_Flag) = UCase(Trim(Request("txtRadioFlag")))
    I1_s_so_hdr(Price_Flag) = UCase(Trim(Request("txtRadioType")))

    I1_s_so_hdr(Cur) = UCase(Trim(Request("txtDoc_cur")))
    I1_s_so_hdr(Cust_Po_No) = UCase(Trim(Request("txtCust_po_no")))
	I1_s_so_hdr(Cust_Po_Dt) = UNIConvDate(Trim(Request("txtCust_po_dt")))
    I1_s_so_hdr(Deal_Type) = UCase(Trim(Request("txtDeal_Type")))
    I1_s_so_hdr(Pay_Meth) = UCase(Trim(Request("txtPay_terms")))

	If Len(Trim(Request("txtPay_dur"))) Then I1_s_so_hdr(Pay_Dur) = Trim(Request("txtPay_dur"))
	I1_s_so_hdr(Trans_Meth) = UCase(Trim(Request("txtTrans_Meth")))
    I1_s_so_hdr(Vat_Inc_Flag) = UCase(Trim(Request("txtVat_Inc_Flag")))	                    
    I1_s_so_hdr(Vat_Type) = UCase(Trim(Request("txtVat_Type")))
    I1_s_so_hdr(Vat_Rate) = UNIConvNum(Request("txtVat_rate"),0)
	
	I1_s_so_hdr(Origin_Cd) = UCase(Trim(Request("txtOrigin")))
	I1_s_so_hdr(Valid_Dt) = UNIConvDate(Trim(Request("txtValid_dt")))
	I1_s_so_hdr(Contract_Dt) = UNIConvDate(Trim(Request("txtContract_dt")))
	I1_s_so_hdr(Ship_Dt_Txt) = Trim(Request("txtShip_dt_txt"))
    I1_s_so_hdr(Pack_Cond) = UCase(Trim(Request("txtPack_cond")))							
	
	I1_s_so_hdr(Inspect_Meth) = UCase(Trim(Request("txtInspect_meth")))
    I1_s_so_hdr(Incoterms) = UCase(Trim(Request("txtIncoTerms")))			                
    I1_s_so_hdr(Dischge_City) = UCase(Trim(Request("txtDischge_city")))
    I1_s_so_hdr(Dischge_Port_Cd) = UCase(Trim(Request("txtDischge_port_Cd")))
    I1_s_so_hdr(Loading_Port_Cd) = UCase(Trim(Request("txtLoading_port_Cd")))
    
    I1_s_so_hdr(Beneficiary) = UCase(Trim(Request("txtBeneficiary")))
    I1_s_so_hdr(Manufacturer) = UCase(Trim(Request("txtManufacturer")))
    I1_s_so_hdr(Agent) = UCase(Trim(Request("txtAgent")))
	I1_s_so_hdr(Remark) = Trim(Request("txtRemark"))
    I1_s_so_hdr(Pay_Type) = UCase(Trim(Request("txtPay_type")))
	    
    I1_s_so_hdr(Ship_Dt) = UNIConvDate(Trim(Request("txtship_dt")))
	I1_s_so_hdr(Pay_Terms_Txt) = Trim(Request("txt_Payterms_txt"))
	I1_s_so_hdr(dn_parcel_flag) = ""
	If Len(Trim(Request("txtXchg_rate"))) Then I1_s_so_hdr(Xchg_Rate) = UNIConvNum(Trim(Request("txtXchg_rate")),0)
	I1_s_so_hdr(to_biz_area) = ""
	
	I1_s_so_hdr(Maint_No) = Trim(Request("txtMaintNo"))
	
	I1_s_so_hdr(ext1_qty) = 0
	I1_s_so_hdr(ext2_qty) = 0
	I1_s_so_hdr(ext3_qty) = 0
	I1_s_so_hdr(ext1_amt) = 0
	I1_s_so_hdr(ext2_amt) = 0
	I1_s_so_hdr(ext3_amt) = 0
	I1_s_so_hdr(ext1_cd) = ""
	I1_s_so_hdr(ext2_cd) = ""
	I1_s_so_hdr(ext3_cd) = "" 
	
	I1_s_so_hdr(Pre_Doc_No) = Trim(Request("txtHSONo"))
	I1_s_so_hdr(sto_flag) = "N"
	
	S_So_Type_Config	 = UCase(Trim(Request("txtSo_Type")))
	Sold_To_Party		 = UCase(Trim(Request("txtSold_to_party")))
    B_Sales_Grp			 = UCase(Trim(Request("txtSales_Grp")))
    Bill_To_Party		 = UCase(Trim(Request("txtBill_to_party")))
    Ship_To_Party		 = UCase(Trim(Request("txtShip_to_party")))
	To_Grp_B_Sales_Grp	 = UCase(Trim(Request("txtTo_Biz_Grp")))
	Payer_To_Party		 = UCase(Trim(Request("txtPayer")))
	strProjectCode		 = UCase(Trim(Request("txtProjectCd")))
	If Len(Trim(Request("txtSending_Bank"))) then B_Bank	= UCase(Trim(Request("txtSending_Bank")))	   
			 
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If
    
    Set iS3G101 = server.CreateObject("PS3G101.cSSoHdrSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Set iS3G101 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    pvCB = "F"
    
    E1_s_so_hdr = iS3G101.S_MAINT_SO_HDR_SVR(pvCB, gstrGlobalCollection, iCommandSent, I1_s_so_hdr, B_Sales_Grp, _
											To_Grp_B_Sales_Grp, B_Bank, Sold_To_Party, Bill_To_Party, _ 
											Ship_To_Party, S_So_Type_Config, Payer_To_Party, pvICustomXML, prOCustomXML, strProjectCode)
    
    If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS3G101 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"				& vbCr
	
	If E1_s_so_hdr <> "" then																		
	Response.Write ".frm1.txtConSo_no.value	= """ & ConvSPChars(E1_s_so_hdr) & """" & vbcr
	End If
	
	Response.Write ".DbSaveOk"                  & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr
	Response.End																				'☜: Process End

End Sub

'============================================================================================================
Sub SubBizDelete()
	
	Dim iCommandSent
	Dim I1_s_so_hdr
	Dim iS3G101	
	Dim pvCB
	
	ReDim I1_s_so_hdr(45)
	
	Const I1_so_no = 0
	
    On Error Resume Next
    Err.Clear 
	
    If Trim(Request("txtSoNo")) = "" Then										'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	
	iCommandSent = "DELETE"	
	I1_s_so_hdr(I1_so_no) = Trim(Request("txtSoNo"))
        
    Set iS3G101 = Server.CreateObject("PS3G101.cSSoHdrSvr")
  
    If CheckSYSTEMError(Err,True) = True Then
       Set iS3G101 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    pvCB = "F"
    
    Call iS3G101.S_MAINT_SO_HDR_SVR(pvCB, gstrGlobalCollection, iCommandSent, I1_s_so_hdr)
    
    If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS3G101 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
  
	Response.Write "<Script Language=vbscript>"	& vbCr
	Response.Write "With parent"				& vbCr
	Response.Write ".DbDeleteOk"                & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr
	Response.End		
	Err.Clear                                                               '☜: Protect system from crashing
    
End Sub

'============================================================================================================
Sub SubLookUp()
	
	Dim iCommandSent
	Dim I1_b_biz_partner
	Dim E1_b_biz_partner
	Dim E2_b_biz_partner
	Dim E3_b_biz_partner
	Dim E4_b_biz_partner
	Dim E5_b_biz_partner
	Dim E6_b_biz_partner
	Dim E7_b_biz_partner
	
	Dim iB5CS41
	Dim iB5GS45
	
	Const E1_Bp_Cd = 0
	Const E1_Bp_Nm = 4
	Const E1_Deal_Type = 28
	Const E1_Deal_Type_Nm = 127
	Const E1_Pay_Terms = 29
	Const E1_Pay_Terms_Nm = 134
	Const E1_Pay_Dur = 30
	Const E1_Vat_Type = 33
	Const E1_Vat_Type_Nm = 133
	Const E1_Vat_Rate = 34
	Const E1_Trans_Meth = 35
	Const E1_Trans_Meth_Nm = 126
	Const E1_Currency = 17
	Const E1_To_Grp = 57
	Const E1_To_Grp_Nm = 131
	Const E1_Biz_Grp = 26
    Const E1_Biz_Grp_Nm = 130
    Const E1_Vat_Inc_Flag = 32
    Const E1_Vat_Inc_Flag_Nm = 140
	
	If Trim(Request("txtSold_to_party")) = "" Then								'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("주문처값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	
	iCommandSent = "QUERY"
	I1_b_biz_partner = Trim(Request("txtSold_to_party"))   
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    	
    Set iB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iB5CS41 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If     
	
	Call iB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection, iCommandSent, I1_b_biz_partner, E1_b_biz_partner)           									 
									 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iB5CS41 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If      
    
    Set iB5CS41 = Nothing   
    
	Set iB5GS45 = Server.CreateObject("PB5GS45.cBListDftBpFtnSvr")    
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       Exit Sub
    End If     
	
	Call iB5GS45.B_LIST_DEFAULT_BP_FTN_SVR(gStrGlobalCollection, I1_b_biz_partner, _
											E2_b_biz_partner, E3_b_biz_partner, _
											E4_b_biz_partner, E5_b_biz_partner, _
											E6_b_biz_partner, E7_b_biz_partner)
											
									 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iB5GS45 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End
       Exit Sub
    End If      
    
    Set iB5GS45 = Nothing   
    
	Response.Write "<Script Language=vbscript>"		& vbCr
	Response.Write "With parent.frm1"				& vbCr
	
	'주문처 
	Response.Write ".txtSold_to_party.value		= """ & ConvSPChars(E1_b_biz_partner(E1_Bp_Cd))				& """" & vbCr
	Response.Write ".txtSold_to_partyNm.value	= """ & ConvSPChars(E1_b_biz_partner(E1_Bp_Nm))				& """" & vbCr
	'거래유형 
	Response.Write ".txtDeal_Type.value			= """ & ConvSPChars(E1_b_biz_partner(E1_Deal_Type))			& """" & vbCr
	Response.Write ".txtDeal_Type_nm.value		= """ & ConvSPChars(E1_b_biz_partner(E1_Deal_Type_Nm))		& """" & vbCr				
	'결제방법 
	Response.Write ".txtPay_terms.value			= """ & ConvSPChars(E1_b_biz_partner(E1_Pay_Terms))			& """" & vbCr
	Response.Write ".txtPay_terms_nm.value		= """ & ConvSPChars(E1_b_biz_partner(E1_Pay_Terms_Nm))		& """" & vbCr			
	'결제기간 
	Response.Write ".txtPay_dur.text			= """ & UNINumClientFormat(E1_b_biz_partner(E1_Pay_Dur),0,0) & """" & vbCr
	'부가세유형 
	Response.Write ".txtVat_Type.value			= """ & ConvSPChars(E1_b_biz_partner(E1_Vat_Type))			& """" & vbCr
	Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(E1_b_biz_partner(E1_Vat_Type_Nm))		& """" & vbCr
	'부가세율 
	Response.Write ".txtVat_rate.text			= """ & UNINumClientFormat(E1_b_biz_partner(E1_Vat_Rate),ggExchRate.DecPoint, 0) & """" & vbCr
	'운송방법 
	Response.Write ".txtTrans_Meth.value		= """ & ConvSPChars(E1_b_biz_partner(E1_Trans_Meth))		& """" & vbCr
	Response.Write ".txtTrans_Meth_nm.value		= """ & ConvSPChars(E1_b_biz_partner(E1_Trans_Meth_Nm))		& """" & vbCr
	'화폐 
	If Trim(E1_b_biz_partner(E1_Currency)) <> "" Then
		Response.Write ".txtDoc_cur.value			= """ & ConvSPChars(E1_b_biz_partner(E1_Currency))			& """" & vbCr
	End If
	
	'수금그룹 
	Response.Write ".txtTo_Biz_Grp.value		= """ & ConvSPChars(E1_b_biz_partner(E1_To_Grp))			& """" & vbCr
	Response.Write ".txtTo_Biz_GrpNm.value		= """ & ConvSPChars(E1_b_biz_partner(E1_To_Grp_Nm))			& """" & vbCr
	'영업그룹 
	Response.Write ".txtSales_Grp.value			= """ & ConvSPChars(E1_b_biz_partner(E1_Biz_Grp))			& """" & vbCr
	Response.Write ".txtSales_GrpNm.value		= """ & ConvSPChars(E1_b_biz_partner(E1_Biz_Grp_Nm))		& """" & vbCr
	'납품처 exp_ssh b_biz_partner
	Response.Write ".txtShip_to_party.value		= """ & ConvSPChars(E2_b_biz_partner(0))					& """" & vbCr
	Response.Write ".txtShip_to_partyNm.value	= """ & ConvSPChars(E2_b_biz_partner(1))					& """" & vbCr
	'발행처 exp_sbi b_biz_partner
	Response.Write ".txtBill_to_party.value		= """ & ConvSPChars(E3_b_biz_partner(0))					& """" & vbCr
	Response.Write ".txtBill_to_partyNm.value	= """ & ConvSPChars(E3_b_biz_partner(1))					& """" & vbCr
	'수금처 exp_spa b_biz_partner
	Response.Write ".txtPayer.value				= """ & ConvSPChars(E4_b_biz_partner(0))					& """" & vbCr
	Response.Write ".txtPayerNm.value			= """ & ConvSPChars(E4_b_biz_partner(1))					& """" & vbCr
	 
    '[CONVERSION INFORMATION]  SQL Result 저장 배열 

	'부가세구분 
	Response.Write ".txtVat_Inc_Flag.value = """ & ConvSPChars(E1_b_biz_partner(E1_Vat_Inc_Flag))			& """" & vbCr
	Response.Write ".txtVat_Inc_Flag_Nm.value = """ &ConvSPChars(E1_b_biz_partner(E1_Vat_Inc_Flag_Nm))		& """" & vbCr
		
	Response.Write "End With"							& vbCr
    Response.Write "parent.lgBlnFlgChgValue = true"		& vbcr
	Response.Write "Call parent.CurrencyOnChange"		& vbcr
	If Len(E1_b_biz_partner(E1_Vat_Type)) Then
	Response.Write "Call parent.SetVatType"				& vbcr
	End If	
	
	Response.Write "</Script>"							& vbCr
	Response.End		
	Err.Clear                                                               '☜: Protect system from crashing

End Sub

'========================================================================================================
Sub CheckCreditlimit()

	On Error Resume Next                                                             '☜: Protect system from crashing
    
    Err.Clear														'☜: Protect system from crashing

	Dim objPS3G113	
	Dim iArrData
	Dim iDblOverLimitAmt
	
	Redim iArrData(2)
    
    iArrData(0) = Trim(Request("txtCaller"))
    iArrData(1) = Trim(Request("txtSoNo"))
    iArrData(2) = Request("txtTotalAmt")
	
	Set objPS3G113 = Server.CreateObject("PS3G113.cChkCreditLimit")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If
  
    Call objPS3G113.ChkCreditLimitSvr(gStrGlobalCollection, iArrData, iDblOverLimitAmt)

	Set objPS3G113 = Nothing	
		
	If Err.number = 0 Then
		Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
		Response.Write("Call parent.ConfirmSO()" & vbCr)
		Response.Write("</SCRIPT>" & vbCr)

    Else
   
		' 여신한도가 초과된 경우(경고처리)
		If InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201929") > 0 Then
			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			Response.Write("Dim iReturnVal" & vbCr)
			' 여신한도를 %1 %2 만큼 초과하였습니다. 저장하시겠습니까?
			Response.Write("iReturnVal = parent.DisplayMsgBox(""201929"", parent.parent.VB_YES_NO, parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr )
			Response.Write("If iReturnVal = vbYes Then Call parent.ConfirmSO()" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
			
		ElseIf InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201722") > 0 Then

			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			'여신한도를 %1 %2 만큼 초과하였습니다.			
			Response.Write("Call parent.DisplayMsgBox(""201722"", ""X"", parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
		Else
			Call CheckSYSTEMError(Err,True)
		End If
	End If
	
	Response.End
End Sub

'============================================================================================================
Sub SubDNCheck()

    Dim iS3G117
    Dim I1_s_so_hdr
	Dim iStrDnFlag
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If Trim(Request("txtSoNo")) = "" Then											 '⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("출하요청 수주번호값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	
    I1_s_so_hdr = Trim(Request("txtSoNo"))
    iStrDnFlag = Trim(Request("RdoDnReq"))
    
    Set iS3G117 = Server.CreateObject("PS3G117.cSCreateDnBySoSvr")    
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iS3G117 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If     
	
    Call iS3G117.S_CREATE_DN_BY_SO_SVR (gStrGlobalCollection, I1_s_so_hdr, iStrDnFlag)
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iS3G117 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If      
    
    Set iS3G117 = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "parent.DbSaveOk"            & vbCr
	Response.Write "</Script>"                  & vbCr
	Response.End																				'☜: Process End

End Sub


'============================================================================================================
Sub SubbtnCONFIRM()

    Dim iS3G150
    Dim I1_s_so_hdr
	
	Const I1_so_no = 0			' I2_s_so_hdr
    Const I1_cfm_flag = 1
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	ReDim I1_s_so_hdr(1)
	
	I1_s_so_hdr(I1_so_no) = Trim(Request("txtSoNo"))
	I1_s_so_hdr(I1_cfm_flag) = Trim(Request("RdoConfirm"))
	
    Set iS3G150 = Server.CreateObject("PS3G150.cSConfirmSalesOrderSvr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iS3G150 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If      
    
	Call iS3G150.S_CONFIRM_SALES_ORDER_SVR(gStrGlobalCollection, I1_s_so_hdr)
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iS3G150 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If      
    
    Set iS3G150 = Nothing
    
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "parent.DbSaveOk"            & vbCr
	Response.Write "</Script>"                  & vbCr
	Response.End																				'☜: Process End

End Sub


'============================================================================================================
Sub SubSoTypeExp()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
	Dim MsgDisplayFlag
	Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Redim UNISqlId(0)
    Redim UNIValue(0, 0)
    																	  '아래에 보면 UNISqlId(1),UNISqlId(2), UNISqlId(3)의 where조건임을 알 수 있다.
    MsgDisplayFlag = False
	
    UNISqlId(0) = "S1911RA101"											  ' main query(spread sheet에 뿌려지는 query statement)
    
	strVal = ""
	
	'---수주타입 
    If Len(Trim(Request("txtSo_Type"))) Then
    	strVal	  = strVal & " " & FilterVar(Request("txtSo_Type"), "''", "S") & "  "
    End If
    
	UNIValue(0, 0)  = strVal
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
   
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg, gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg, vbInformation, I_MKSCRIPT)
    End If    
        
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("201600", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
	   MsgDisplayFlag = True
'       Exit Sub
    End If
    
	Response.Write "<Script Language=vbscript>"		& vbCr
	Response.Write "With parent.frm1"				& vbCr
	
	Response.write ".txtSo_TypeNm.value			= """ & ConvSPChars(rs0("SO_TYPE_NM"))		& """" & vbCr
	Response.write ".txtSoTypeExportFlag.value	= """ & ConvSPChars(rs0("EXPORT_FLAG"))		& """" & vbCr
	Response.write ".txtSoTypeRetItemFlag.value	= """ & ConvSPChars(rs0("RET_ITEM_FLAG"))	& """" & vbCr
	Response.write ".txtHDlvyLt.value			= """ & ConvSPChars(rs0("DLVY_LT"))			& """" & vbCr
	Response.write ".txtSoTypeCiFlag.value		= """ & ConvSPChars(rs0("CI_FLAG"))			& """" & vbCr

	Response.write "Call parent.fncSoTypeExpChange() "	& vbCr
	Response.Write "End With"						& vbCr
	Response.Write "</Script>"						& vbCr

	Response.End																				'☜: Process End

End Sub


'============================================================================================================
' Name : SubProjectRef
' Desc : Query Data from Db
'============================================================================================================
Sub SubProjectRef()
	
	Dim iS3G105
	Dim iCommandSent
	Dim I1_s_so_hdr
	Dim E1_s_so_hdr

    Const E1_prj_cd = 0
    Const E1_req_dlvy_dt = 1
    Const E1_bp_cd = 2
    Const E1_bp_nm = 3
    Const E1_cur = 4
    Const E1_xchg_rate = 5
    Const E1_net_amt = 6
    Const E1_net_amt_loc = 7
    Const E1_pay_meth = 8
    Const E1_pay_meth_nm = 9
    Const E1_vat_inc_flag = 10
    Const E1_vat_type = 11
    Const E1_vat_type_nm = 12
    Const E1_vat_rate = 13
    Const E1_vat_amt = 14
    Const E1_vat_amt_loc = 15
    Const E1_insrt_user_id = 16
    Const E1_insrt_dt = 17
    Const E1_updt_user_id = 18
    Const E1_updt_dt = 19
    Const E1_pay_type = 20
    Const E1_pay_type_nm = 21
    Const E1_sales_grp = 22
    Const E1_sales_grp_nm = 23
    
	On Error Resume Next
	Err.Clear                                                               '☜: Protect system from crashing
	
	iCommandSent = "QUERY"
    I1_s_so_hdr = Trim(Request("txtProjectCd"))

	Set iS3G105 = Server.CreateObject ("PS3G105.cLookupProjectSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G105 = Nothing	
		Exit Sub
	End If
	
	Call iS3G105.S_LOOKUP_PROJECT_SVR(gStrGlobalCollection, iCommandSent, I1_s_so_hdr, E1_s_so_hdr)
											
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G105 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script Language=vbscript>"			& vbCr   
        Response.Write "parent.frm1.txtConSo_no.focus"		& vbCr    
        Response.Write "</Script>      "					& vbCr	                 '☜: Unload Comproxy DLL
        Exit Sub
	End If
	
	Set iS3G105 = Nothing
	    
    '프로젝트 참조시에 조회 
    If lgOpModeCRUD = "PROJECTQUERY" then 
        
        Response.Write "<Script Language=vbscript>" & vbCr
    	Response.Write "With parent.frm1"			& vbCr
    	
    	
    	'--첫번째 TAB 
    	
    	Response.write ".txtProjectCd.value			= """ & ConvSPChars(E1_s_so_hdr(E1_prj_cd))						& """" & vbCr
    	'납기일 
    	Response.write ".txtReq_dlvy_dt.Text		= """ & UNIDateClientFormat(E1_s_so_hdr(E1_req_dlvy_dt))		& """" & vbCr
    	
    	Response.write ".txtSold_to_party.value		= """ & ConvSPChars(E1_s_so_hdr(E1_bp_cd))						& """" & vbCr
    	Response.write ".txtSold_to_partyNm.value	= """ & ConvSPChars(E1_s_so_hdr(E1_bp_nm))						& """" & vbCr
    	Response.write ".txtSales_Grp.value			= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp))					& """" & vbCr
    	Response.write ".txtSales_GrpNm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_sales_grp_nm))				& """" & vbCr
    	'결제방법 
    	Response.write ".txtPay_terms.value			= """ & ConvSPChars(E1_s_so_hdr(E1_pay_meth))					& """" & vbCr
    	Response.write ".txtPay_terms_nm.value		= """ & ConvSPChars(E1_s_so_hdr(E1_pay_meth_nm))				& """" & vbCr
    	Response.write ".txtVat_Type.value			= """ & ConvSPChars(E1_s_so_hdr(E1_vat_type))					& """" & vbCr
    	Response.write ".txtVatTypeNm.value			= """ & ConvSPChars(E1_s_so_hdr(E1_vat_type_nm))				& """" & vbCr
    	'부가세율 
    	Response.write ".txtVat_rate.text			= """ & UNINumClientFormat(E1_s_so_hdr(E1_vat_rate),ggExchRate.DecPoint, 0)	& """" & vbCr
    	Response.write ".txtDoc_cur.value			= """ & ConvSPChars(E1_s_so_hdr(E1_cur))						& """" & vbCr
    	
    	
    	Response.write ".txtNet_amt.text			= """ & UNINumClientFormat(E1_s_so_hdr(E1_net_amt),ggExchRate.DecPoint, 0)	& """" & vbCr
    	Response.write ".txtVat_amt.text			= """ & UNINumClientFormat(E1_s_so_hdr(E1_vat_amt),ggExchRate.DecPoint, 0)	& """" & vbCr
    	
    	If E1_s_so_hdr(E1_cur) = gCurrency Then 
    		Response.write "parent.ggoOper.SetReqAttr .txtXchg_rate,  """ & "Q" & """" 									& vbCr
    	Else 
    		Response.write "parent.ggoOper.SetReqAttr .txtXchg_rate,  """ & "N" & """"									& vbCr
    	End If 
    
    	Response.write ".txtVat_Inc_Flag.value		= """ & ConvSPChars(E1_s_so_hdr(E1_vat_inc_flag))				& """" & vbCr
    	If ConvSPChars(E1_s_so_hdr(E1_vat_inc_flag)) = "1" Then
    		Response.write ".txtVat_Inc_Flag_Nm.value	= """ & ConvSPChars("별도")			& """" & vbCr
    	ElseIf ConvSPChars(E1_s_so_hdr(E1_vat_inc_flag)) = "2" Then
    		Response.write ".txtVat_Inc_Flag_Nm.value	= """ & ConvSPChars("포함")			& """" & vbCr
    	Else	
    		Response.write ".txtVat_Inc_Flag_Nm.value	= """ & ConvSPChars("")			& """" & vbCr
    	End If	
    		
    	'환율 
    	Response.write ".txtXchg_rate.value			= """ & UNINumClientFormat(E1_s_so_hdr(E1_xchg_rate), ggExchRate.DecPoint,0)	& """" & vbCr
    	
    	Response.Write "End With"			& vbCr
        Response.Write "</Script>"			& vbCr
        Response.End 
    End if
    	
End Sub


'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	
End Sub


'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub


'============================================================================================================
Sub SubMakeSQLStatements(pMode)
    
    On Error Resume Next
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub


%>
