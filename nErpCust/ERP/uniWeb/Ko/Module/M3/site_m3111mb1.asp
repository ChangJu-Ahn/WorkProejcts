<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111mb1
'*  4. Program Name         : 발주등록 
'*  5. Program Desc         : 발주등록 
'*  6. Component List       : PM3G119.cMLookupPurOrdHdrS / PM3G111.cMMaintPurOrdHdrS / PM3G1R1.cMReleasePurOrdS / PM1G419.cMLkConfigProcessS / PB5CS41.cLookupBizPartnerSvr
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Min, HJ
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	Dim lgOpModeCRUD
	Dim ls_msg
	On Error Resume Next
																								'☜: Protect system from crashing
	Err.Clear 
																								'☜: Clear Error status
	Call HideStatusWnd
	lgOpModeCRUD	=	Request("txtMode")
																								'☜: Read Operation Mode (CRUD)	
	Select Case lgOpModeCRUD		
	   Case CStr(UID_M0001)																		'☜: Query
	      Call SubBizQuery()
	   Case CStr(UID_M0002)																		'☜: Save
	      Call SubBizSave()
	   Case CStr(UID_M0003)																		'☜: Delete
	      Call SubBizDelete()
	   Case "Release"  
		  Call SubReleaseCheck()
	   Case "UnRelease"  
		  Call SubReleaseCheck()
	   Case "SendingB2B"
		  Call SubSendingB2B()
	   Case "LookUpPoType"
		  Call SubLookUpPoType()
'	   Case "SupplierLookupAfterOnline"
'		  Call SubSupplierLookupAfterOnline()
	   Case "LookUpSupplier"
		  Call SubLookUpSupplier()
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
    Const M239_E22_release_flg = 0
    Const M239_E22_merg_pur_flg = 1
    Const M239_E22_po_no = 2
    Const M239_E22_po_dt = 3
    Const M239_E22_xch_rt = 4
    Const M239_E22_vat_type = 5
    Const M239_E22_vat_rt = 6
    Const M239_E22_tot_vat_doc_amt = 7
    Const M239_E22_tot_po_doc_amt = 8
    Const M239_E22_tot_po_loc_amt = 9
    Const M239_E22_pay_meth = 10
    Const M239_E22_pay_dur = 11
    Const M239_E22_pay_terms_txt = 12
    Const M239_E22_pay_type = 13
    Const M239_E22_sppl_sales_prsn = 14
    Const M239_E22_sppl_tel_no = 15
    Const M239_E22_remark = 16
    Const M239_E22_vat_inc_flag = 17
    Const M239_E22_offer_dt = 18
    Const M239_E22_fore_dvry_dt = 19
    Const M239_E22_expiry_dt = 20
    Const M239_E22_invoice_no = 21
    Const M239_E22_incoterms = 22
    Const M239_E22_transport = 23
    Const M239_E22_sending_bank = 24
    Const M239_E22_delivery_plce = 25
    Const M239_E22_applicant = 26
    Const M239_E22_manufacturer = 27
    Const M239_E22_agent = 28
    Const M239_E22_origin = 29
    Const M239_E22_packing_cond = 30
    Const M239_E22_inspect_means = 31
    Const M239_E22_dischge_city = 32
    Const M239_E22_dischge_port = 33
    Const M239_E22_loading_port = 34
    Const M239_E22_shipment = 35
    Const M239_E22_import_flg = 36
    Const M239_E22_bl_flg = 37
    Const M239_E22_cc_flg = 38
    Const M239_E22_rcpt_flg = 39
    Const M239_E22_subcontra_flg = 40
    Const M239_E22_ret_flg = 41
    Const M239_E22_iv_flg = 42
    Const M239_E22_rcpt_type = 43
    Const M239_E22_issue_type = 44
    Const M239_E22_iv_type = 45
    Const M239_E22_po_cur = 46
    Const M239_E22_xch_rate_op = 47
    Const M239_E22_cls_flg = 48
    Const M239_E22_ref_no = 49
    Const M239_E22_tot_vat_loc_amt = 50

	
	Const M239_E19_po_type_cd = 0
	Const M239_E19_po_type_nm = 1
	
	Const M239_E9_bp_cd = 0
	Const M239_E9_bp_nm = 1
	
	Const M239_E20_pur_grp = 0
	Const M239_E20_pur_grp_nm = 1
	
	Dim M31119
	Dim E1_b_bank_bank_nm
	Dim E2_b_minor_vat_type
	Dim E3_b_minor_pay_meth   
	Dim E4_b_minor_pay_type    
	Dim E5_b_minor_incoterms    
	Dim E6_b_minor_transport    
	Dim E7_b_minor_delivery_plce    
	Dim E8_b_minor_origin    
	Dim E9_b_biz_partner    
	Dim E10_b_biz_partner_applicant_nm    
	Dim E11_b_biz_partner_manufacturer_nm    
	Dim E12_b_minor_packing_cond    
	Dim E13_b_minor_inspect_means    
	Dim E14_b_minor_dischge_city    
	Dim E15_b_minor_dischge_port    
	Dim E16_b_minor_loading_port    
	Dim E17_b_configuration_reference
	Dim E18_b_currency_currency_desc
	Dim E19_m_config_process
	Dim E20_b_pur_grp
	Dim E21_b_biz_partner_agent_nm
	Dim E22_m_pur_ord_hdr
	Dim iPoNo
	Dim lgCurrency
    Err.Clear                                                               '☜: Protect system from crashing


	On Error Resume Next

    Set M31119 = Server.CreateObject("PM3G119.cMLookupPurOrdHdrS")    

	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End if
     
     iPoNo=Trim(Request("txtPoNo"))
     Call M31119.M_LOOKUP_PUR_ORD_HDR_SVR(gStrGlobalCollection, _
									  iPoNo, _
									  E1_b_bank_bank_nm, _
									  E2_b_minor_vat_type, _
									  E3_b_minor_pay_meth, _
									  E4_b_minor_pay_type, _
									  E5_b_minor_incoterms, _
									  E6_b_minor_transport, _
									  E7_b_minor_delivery_plce, _
									  E8_b_minor_origin, _
									  E9_b_biz_partner, _
									  E10_b_biz_partner_applicant_nm, _
									  E11_b_biz_partner_manufacturer_nm, _
									  E12_b_minor_packing_cond, _
									  E13_b_minor_inspect_means, _
									  E14_b_minor_dischge_city, _
									  E15_b_minor_dischge_port, _
									  E16_b_minor_loading_port, _
									  E17_b_configuration_reference, _
									  E18_b_currency_currency_desc, _
									  E19_m_config_process, _
									  E20_b_pur_grp, _
									  E21_b_biz_partner_agent_nm, _
									  E22_m_pur_ord_hdr)
    										      
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Response.Write "<Script Language=vbscript>"  	& vbCr
		Response.Write "parent.frm1.txtPoNo.focus" & vbCr
		Response.Write "call parent.ggoOper.ClearField(parent.Document, ""2"")   "                 & vbCr
		Response.Write "</Script>" 	& vbCr
		Set M31119 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If

	If UCase(Trim(E22_m_pur_ord_hdr(M239_E22_ret_flg))) = "Y" Then
       Call DisplayMsgBox("17a014", vbOKOnly, "반품발주건", "조회", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>"  	& vbCr
		Response.Write "parent.frm1.txtPoNo.focus" & vbCr
		Response.Write "call parent.ggoOper.ClearField(parent.Document, ""2"")   "                 & vbCr
		Response.Write "</Script>" 	& vbCr
	   Set M31119 = Nothing																	'☜: ComProxy UnLoad
	   Response.End																				'☜: Process End
	End If

   Set M31119 = Nothing																	'☜: ComProxy UnLoad

	'-----------------------
	'Result data display area
	'----------------------- 
	Dim strDefDate
	
	lgCurrency = ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_po_cur)))
	strDefDate = UniDateClientFormat("1899-12-31")

	Response.Write "<Script Language=vbscript>"															& vbCr
	Response.Write "With parent.frm1"																	& vbCr
	Response.Write "	.txtCurr.value	= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_po_cur)))	& """" 	& vbCr
	Response.Write "	parent.CurFormatNumericOCX	" 													& vbCr
	
	'첫번째 탭 
	Response.Write "	.txtPotypeCd.value	= """ & ConvSPChars(Trim(E19_m_config_process(M239_E19_po_type_cd)))	& """" & vbCr
	Response.Write "	.txtPotypeNm.value	= """ & ConvSPChars(Trim(E19_m_config_process(M239_E19_po_type_nm)))	& """" & vbCr
	
	If Trim(UCase(E22_m_pur_ord_hdr(M239_E22_release_flg))) = "Y" Then
		Response.Write "	.rdoRelease(1).Checked = true" & vbCr
	Else
		Response.Write "	.rdoRelease(0).Checked = true" & vbCr
	End If

	If UCase(Trim(E22_m_pur_ord_hdr(M239_E22_merg_pur_flg))) = "Y" Then
		Response.Write "	.rdoMergPurFlg(0).Checked = true" & vbCr
	Else
		Response.Write "	.rdoMergPurFlg(1).Checked = true" & vbCr
	End If

	Response.Write "	.txtPoNo2.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_po_no))) 			& """"	& vbCr
	Response.Write "	.txtPoDt.Text				= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_po_dt)) 	& """"	& vbCr				
	Response.Write "	.txtSupplierCd.value		= """ & ConvSPChars(Trim(E9_b_biz_partner(M239_E9_bp_cd))) 			& """"	& vbCr
	Response.Write "	.txtSupplierNm.value		= """ & ConvSPChars(Trim(E9_b_biz_partner(M239_E9_bp_nm))) 			& """"	& vbCr
	Response.Write "	.txtGroupCd.value			= """ & ConvSPChars(Trim(E20_b_pur_grp(M239_E20_pur_grp))) 			& """"	& vbCr
	Response.Write "	.txtGroupNm.value			= """ & ConvSPChars(Trim(E20_b_pur_grp(M239_E20_pur_grp_nm))) 		& """"	& vbCr					
	Response.Write "	.txtCurrNm.value			= """ & ConvSPChars(Trim(E18_b_currency_currency_desc)) 							& """"	& vbCr
	Response.Write "	.txtXch.value				= """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_xch_rt),ggExchRate.DecPoint,0) & """"	& vbCr
	Response.Write "	.txtVatType.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_vat_type))) 		& """"	& vbCr
	Response.Write "	.txtVatTypeNm.Value			= """ & ConvSPChars(Trim(E2_b_minor_vat_type)) 						& """"	& vbCr
	Response.Write "	.txtVatRt.Text				= """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_vat_rt),ggExchRate.DecPoint,0) & """"	& vbCr		
	Response.Write "	.txtVatAmt.Text             = """ & UNIConvNumDBToCompanyByCurrency(E22_m_pur_ord_hdr(M239_E22_tot_vat_doc_amt),lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo  , "X") & """"	& vbCr
	Response.Write "	.txtPoAmt.Text				= """ & UNIConvNumDBToCompanyByCurrency(E22_m_pur_ord_hdr(M239_E22_tot_po_doc_amt),lgCurrency,ggAmtOfMoneyNo , "X" , "X") & """"	& vbCr
	Response.Write "	.txtGrossPoAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(UniCDbl(E22_m_pur_ord_hdr(M239_E22_tot_po_doc_amt),0)+UniCDbl(E22_m_pur_ord_hdr(M239_E22_tot_vat_doc_amt),0),lgCurrency,ggAmtOfMoneyNo , "X" , "X") & """"	& vbCr
		
	Response.Write "	.txtPoLocAmt.Text			= """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_tot_po_loc_amt),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.txtGrossPoLocAmt.Text		= """ & UNINumClientFormat(UniCDbl(E22_m_pur_ord_hdr(M239_E22_tot_po_loc_amt),0)+UniCDbl(E22_m_pur_ord_hdr(M239_E22_tot_vat_loc_amt),0),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.txtPaytermCd.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_pay_meth))) 		& """"	& vbCr
	Response.Write "	.txtPaytermNm.Value			= """ & ConvSPChars(Trim(E3_b_minor_pay_meth)) 						& """"	& vbCr
	Response.Write "	.txtReference.Value			= """ & ConvSPChars(Trim(E17_b_configuration_reference)) 				& """"	& vbCr
	Response.Write "	.txtPayDur.Text				= """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_pay_dur),0,0) & """"	& vbCr
	Response.Write "	.txtPaytermstxt.Value		= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_pay_terms_txt))) 	& """"	& vbCr
	Response.Write "	.txtPayTypeCd.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_pay_type))) 		& """"	& vbCr
	Response.Write "	.txtPayTypeNm.Value			= """ & ConvSPChars(Trim(E4_b_minor_pay_type)) 						& """"	& vbCr
	Response.Write "	.txtSuppSalePrsn.Value		= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_sppl_sales_prsn))) & """"	& vbCr
	Response.Write "	.txtTel.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_sppl_tel_no))) 	& """"	& vbCr
	Response.Write "	.txtRemark.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_remark))) 		& """"	& vbCr
	
	If ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_vat_inc_flag))) = "2" Then 'vat 포함 여부 vat_inc_flag 
		Response.Write "	.rdoVatFlg2.Checked= true" 	& vbCr
		Response.Write "	.hdvatFlg.value=""2""" 		& vbCr
	Else
		Response.Write "	.rdoVatFlg1.Checked= true" 	& vbCr
		Response.Write "	.hdvatFlg.value=""1""" 		& vbCr
	End If		
	
	'==2003/01월 패치	
	Response.Write "	.txtDvryDt.Text				= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_fore_dvry_dt))	& """" & vbCr
	'두번째 탭 
	If E22_m_pur_ord_hdr(M239_E22_import_flg) = "Y" Then
	
		Response.Write "	.txtOffDt.Text				= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_offer_dt))		& """" & vbCr

		Response.Write "	.txtExpiryDt.Text			= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_expiry_dt))		& """" & vbCr
		Response.Write "	.txtInvNo.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_invoice_no)))				& """" & vbCr
		Response.Write "	.txtIncotermsCd.value		= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_incoterms)))				& """" & vbCr
		Response.Write "	.txtIncotermsNm.value		= """ & ConvSPChars(Trim(E5_b_minor_incoterms))								& """" & vbCr
		Response.Write "	.txtTransCd.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_transport)))				& """" & vbCr
		Response.Write "	.txtTransNm.value			= """ & ConvSPChars(Trim(E6_b_minor_transport))								& """" & vbCr
		Response.Write "	.txtBankCd.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_sending_bank)))			& """" & vbCr
		Response.Write "	.txtBankNm.value			= """ & ConvSPChars(Trim(E1_b_bank_bank_nm))											& """" & vbCr
		Response.Write "	.txtDvryplce.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_delivery_plce)))			& """" & vbCr
		Response.Write "	.txtDvryplceNm.value		= """ & ConvSPChars(Trim(E7_b_minor_delivery_plce))							& """" & vbCr
		Response.Write "	.txtApplicantCd.Value		= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_applicant)))				& """" & vbCr
		Response.Write "	.txtApplicantNm.Value		= """ & ConvSPChars(Trim(E10_b_biz_partner_applicant_nm))									& """" & vbCr
		Response.Write "	.txtManuCd.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_manufacturer)))			& """" & vbCr			
		Response.Write "	.txtManuNm.Value			= """ & ConvSPChars(Trim(E11_b_biz_partner_manufacturer_nm))									& """" & vbCr
		Response.Write "	.txtAgentCd.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_agent)))					& """" & vbCr
		Response.Write "	.txtAgentNm.Value			= """ & ConvSPChars(Trim(E21_b_biz_partner_agent_nm))									& """" & vbCr
		Response.Write "	.txtOrigin.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_origin)))					& """" & vbCr
		Response.Write "	.txtOriginNm.value			= """ & ConvSPChars(Trim(E8_b_minor_origin))									& """" & vbCr				
		Response.Write "	.txtPackingCd.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_packing_cond)))			& """" & vbCr
		Response.Write "	.txtPackingNm.value			= """ & ConvSPChars(Trim(E12_b_minor_packing_cond))							& """" & vbCr				
		Response.Write "	.txtInspectCd.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_inspect_means)))			& """" & vbCr
		Response.Write "	.txtInspectNm.Value			= """ & ConvSPChars(Trim(E13_b_minor_inspect_means))							& """" & vbCr
		Response.Write "	.txtDisCity.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_dischge_city)))			& """" & vbCr
		Response.Write "	.txtDisCityNm.Value			= """ & ConvSPChars(Trim(E14_b_minor_dischge_city))							& """" & vbCr
		Response.Write "	.txtDisPort.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_dischge_port)))			& """" & vbCr
		Response.Write "	.txtDisPortNm.value			= """ & ConvSPChars(Trim(E15_b_minor_dischge_port))							& """" & vbCr
		Response.Write "	.txtLoadPort.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_loading_port)))			& """" & vbCr
		Response.Write "	.txtLoadPortNm.Value		= """ & ConvSPChars(Trim(E16_b_minor_loading_port))							& """" & vbCr
		Response.Write "	.txtShipment.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_shipment)))				& """" & vbCr
	End if
		
		Response.Write "	.hdnImportflg.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_import_flg)))				& """" & vbCr
		Response.Write "	.hdnBLflg.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_bl_flg)))					& """" & vbCr
		Response.Write "	.hdnCcflg.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_cc_flg)))					& """" & vbCr
		Response.Write "	.hdnRcptflg.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_rcpt_flg)))				& """" & vbCr
		Response.Write "	.hdnSubcontraflg.value		= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_subcontra_flg)))			& """" & vbCr
		Response.Write "	.hdnRetflg.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_ret_flg)))				& """" & vbCr
		Response.Write "	.hdnIvflg.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_iv_flg)))					& """" & vbCr
		Response.Write "	.hdnRcptType.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_rcpt_type)))				& """" & vbCr
		Response.Write "	.hdnIssueType.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_issue_type)))				& """" & vbCr
		Response.Write "	.hdnIvType.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_iv_type)))				& """" & vbCr
		
		Response.Write "	.hdnxchrateop.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_xch_rate_op)))			& """" & vbCr 'Multi Divide 
		Response.Write "	.hdclsflg.value			    = """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_cls_flg)))			& """" & vbCr 'Multi Divide 
    Response.Write "	parent.lgNextNo = """"" 	& vbCr		' 다음 키 값 넘겨줌 
    Response.Write "	parent.lgPrevNo = """"" 	& vbCr		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 
    Response.Write "	parent.DbQueryOk" 	& vbCr
    Response.Write "End With" 	& vbCr
    Response.Write "</Script>" 	& vbCr

    Set M31119 = Nothing															'☜: Unload Comproxy
	
End Sub																				'☜: Process End
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSave()																	'☜: 저장 요청을 받음 
	
	Const M193_I2_po_no = 0										
	Const M193_I2_merg_pur_flg = 1                              
	Const M193_I2_pur_org = 2                                   
	Const M193_I2_pur_biz_area = 3                              
	Const M193_I2_pur_cost_cd = 4                               
	Const M193_I2_po_dt = 5                                     
	Const M193_I2_po_cur = 6                                    
	Const M193_I2_xch_rt = 7                                    
	Const M193_I2_pay_meth = 8                                  
	Const M193_I2_pay_dur = 9                                   
	Const M193_I2_pay_terms_txt = 10                            
	Const M193_I2_pay_type = 11                                 
	Const M193_I2_vat_type = 12                                 
	Const M193_I2_vat_rt = 13                                   
	Const M193_I2_tot_vat_doc_amt = 14                          
	Const M193_I2_tot_vat_loc_amt = 15                          
	Const M193_I2_tot_po_doc_amt = 16                           
	Const M193_I2_tot_po_loc_amt = 17                           
	Const M193_I2_sppl_sales_prsn = 18                          
	Const M193_I2_sppl_tel_no = 19                              
	Const M193_I2_release_flg = 20                              
	Const M193_I2_cls_flg = 21                                  
	Const M193_I2_import_flg = 22                               
	Const M193_I2_lc_flg = 23                                   
	Const M193_I2_bl_flg = 24                                   
	Const M193_I2_cc_flg = 25                                   
	Const M193_I2_rcpt_flg = 26                                 
	Const M193_I2_subcontra_flg = 27                            
	Const M193_I2_ret_flg = 28                                  
	Const M193_I2_iv_flg = 29                                   
	Const M193_I2_rcpt_type = 30                                
	Const M193_I2_issue_type = 31                               
	Const M193_I2_iv_type = 32                                  
	Const M193_I2_sppl_cd = 33                                  
	Const M193_I2_payee_cd = 34                                 
	Const M193_I2_build_cd = 35                                 
	Const M193_I2_remark = 36                                   
	Const M193_I2_manufacturer = 37                             
	Const M193_I2_agent = 38                                    
	Const M193_I2_applicant = 39                                
	Const M193_I2_offer_dt = 40                                 
	Const M193_I2_expiry_dt = 41                                
	Const M193_I2_transport = 42                                
	Const M193_I2_incoterms = 43                                
	Const M193_I2_delivery_plce = 44                            
	Const M193_I2_packing_cond = 45                             
	Const M193_I2_inspect_means = 46                            
	Const M193_I2_dischge_city = 47                             
	Const M193_I2_dischge_port = 48                             
	Const M193_I2_loading_port = 49                             
	Const M193_I2_origin = 50                                   
	Const M193_I2_sending_bank = 51                             
	Const M193_I2_invoice_no = 52                               
	Const M193_I2_fore_dvry_dt = 53                             
	Const M193_I2_shipment = 54                                 
	Const M193_I2_charge_flg = 55                               
	Const M193_I2_tracking_no = 56                              
	Const M193_I2_so_no = 57                                    
	Const M193_I2_inspect_method = 58                           
	Const M193_I2_ext1_cd = 59                                  
	Const M193_I2_ext1_qty = 60                                 
	Const M193_I2_ext1_amt = 61                                 
	Const M193_I2_ext1_rt = 62                                  
	Const M193_I2_ext1_dt = 63                                  
	Const M193_I2_ext2_cd = 64                                  
	Const M193_I2_ext2_qty = 65                                 
	Const M193_I2_ext2_amt = 66                                 
	Const M193_I2_ext2_rt = 67                                  
	Const M193_I2_ext2_dt = 68                                  
	Const M193_I2_ext3_cd = 69                                  
	Const M193_I2_ext3_qty = 70                                 
	Const M193_I2_ext3_amt = 71                                 
	Const M193_I2_ext3_rt = 72                                  
	Const M193_I2_ext3_dt = 73                                  
	Const M193_I2_xch_rate_op = 74                              
	Const M193_I2_vat_inc_flag = 75                             
	Const M193_I2_ref_no = 76                                   
	Const M193_I2_sto_flg = 77
	Const M193_I2_so_type = 78
	
	Dim M31111
	Dim lgIntFlgMode
	Dim iStrCommandSent
	Dim iStrPoNo
	Dim I1_b_company
	Dim I2_m_config_process
	Dim I3_b_biz_partner
	Dim I4_b_pur_grp
	Dim I5_m_pur_ord_hdr
	
	Redim I5_m_pur_ord_hdr(M193_I2_so_type)
	
	On Error Resume Next								
    Err.Clear																		'☜: Protect system from crashing
	
	If Len(Trim(Request("txtPoDt"))) Then
		If UNIConvDate(Request("txtPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtPoDt", 1, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

    '-----------------------
    'Data manipulate area
    '-----------------------
    '첫번째 탭 
    
    If lgIntFlgMode = OPMD_CMODE And Trim(Request("txtPoNo2")) <> "" Then
    	I5_m_pur_ord_hdr(M193_I2_po_no)				= Trim(Request("txtPoNo2"))
    End if
    
    I5_m_pur_ord_hdr(M193_I2_merg_pur_flg)			= UCase(Request("hdnMergPurFlg"))
	I2_m_config_process								= UCase(Trim(Request("txtPotypeCd")))
    I5_m_pur_ord_hdr(M193_I2_release_flg)			= UCase(Request("hdnrelease"))
    I5_m_pur_ord_hdr(M193_I2_po_dt)					= UNIConvDate(Request("txtPoDt"))
	I4_b_pur_grp									= UCase(Trim(Request("txtGroupCd")))
  	I3_b_biz_partner								= UCase(Trim(Request("txtSupplierCd")))
  	I5_m_pur_ord_hdr(M193_I2_tot_po_doc_amt)		= UNIConvNum(Request("txtPoAmt"),0)
  	I5_m_pur_ord_hdr(M193_I2_tot_vat_doc_amt)		= UNIConvNum(Request("txtVatAmt"),0)
  	I5_m_pur_ord_hdr(M193_I2_po_cur)				= UCase(Trim(Request("txtCurr")))
  	
    If Trim(Request("txtXch")) = "" Then
    	I5_m_pur_ord_hdr(M193_I2_xch_rt)			= "0"
    Else
    	I5_m_pur_ord_hdr(M193_I2_xch_rt)			= UNIConvNum(Request("txtXch"),0)
	End If 
	
  	I5_m_pur_ord_hdr(M193_I2_vat_type)				= UCase(Trim(Request("txtVatType")))
  	I5_m_pur_ord_hdr(M193_I2_pay_meth)				= UCase(Trim(Request("txtPaytermCd")))
  	
    If Trim(Request("txtPayDur")) = "" Then
	  	I5_m_pur_ord_hdr(M193_I2_pay_dur)			= "0"
	Else
	  	I5_m_pur_ord_hdr(M193_I2_pay_dur)			= UNIConvNum(Request("txtPayDur"),0)
	End If 
	
  	I5_m_pur_ord_hdr(M193_I2_vat_rt)				= UNIConvNum(Request("txtVatrt"),0)
  	I5_m_pur_ord_hdr(M193_I2_pay_terms_txt)			= Trim(Request("txtPayTermstxt"))
  	I5_m_pur_ord_hdr(M193_I2_pay_type)				= UCase(Trim(Request("txtPayTypeCd")))
  	I5_m_pur_ord_hdr(M193_I2_sppl_sales_prsn)		= Trim(Request("txtSuppSalePrsn"))
  	I5_m_pur_ord_hdr(M193_I2_sppl_tel_no)			= Trim(Request("txtTel"))
  	I5_m_pur_ord_hdr(M193_I2_remark)				= Trim(Request("txtRemark"))
  	
  	I5_m_pur_ord_hdr(M193_I2_vat_inc_flag)			= UCase(Trim(Request("hdvatFlg")))    '13차 vat 포함 구분 vat_inc_flag
	
	I1_b_company									= gCurrency

    '===== 내수의 경우에도 수입자는 항상 넘긴다.
    I5_m_pur_ord_hdr(M193_I2_applicant)				= Trim(UCase(Request("txtApplicantCd")))
    '====200301월 패치 
	I5_m_pur_ord_hdr(M193_I2_fore_dvry_dt)		= UNIConvDate(Request("txtDvryDt"))
	
	'두번째 탭 
    if UCase(Trim(Request("hdnImportflg"))) = "Y" then
    
		If Len(Trim(Request("txtOffDt"))) Then
			If UNIConvDate(Request("txtOffDt")) = "" Then
			    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			    Call LoadTab("parent.frm1.txtOffDt", 2, I_MKSCRIPT)
				Set M31111 = Nothing	
			    Exit Sub	
			End If
		End If
		If Len(Trim(Request("txtDvryDt"))) Then
			If UNIConvDate(Request("txtDvryDt")) = "" Then
			    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			    Call LoadTab("parent.frm1.txtDvryDt", 2, I_MKSCRIPT)
				Set M31111 = Nothing	
			    Exit Sub	
			End If
		End If
		If Len(Trim(Request("txtExpiryDt"))) Then
			If UNIConvDate(Request("txtExpiryDt")) = "" Then
			    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			    Call LoadTab("parent.frm1.txtExpiryDt", 2, I_MKSCRIPT)
				Set M31111 = Nothing	
			    Exit Sub	
			End If
		End If
		
				
	  	I5_m_pur_ord_hdr(M193_I2_offer_dt)			= UNIConvDate(Request("txtOffDt"))
	
		if Trim(Request("txtExpiryDt"))="" then
		  	I5_m_pur_ord_hdr(M193_I2_expiry_dt)		= UNIConvDate(Request("txtDvryDt"))
		else
		  	I5_m_pur_ord_hdr(M193_I2_expiry_dt)		= UNIConvDate(Request("txtExpiryDt"))
		end if 

	  	I5_m_pur_ord_hdr(M193_I2_invoice_no)		= Trim(Request("txtInvNo"))
	  	I5_m_pur_ord_hdr(M193_I2_incoterms)			= UCase(Trim(Request("txtIncotermsCd")))
	  	I5_m_pur_ord_hdr(M193_I2_transport)			= UCase(Trim(Request("txtTransCd")))
	  	I5_m_pur_ord_hdr(M193_I2_sending_bank)		= UCase(Trim(Request("txtBankCd")))
	  	I5_m_pur_ord_hdr(M193_I2_delivery_plce)		= UCase(Trim(Request("txtDvryPlce")))
	  	I5_m_pur_ord_hdr(M193_I2_manufacturer)		= UCase(Trim(Request("txtManuCd")))
	  	I5_m_pur_ord_hdr(M193_I2_agent)				= UCase(Trim(Request("txtAgentCd")))
	  	I5_m_pur_ord_hdr(M193_I2_origin)			= UCase(Trim(Request("txtOrigin")))
	  	I5_m_pur_ord_hdr(M193_I2_packing_cond)		= UCase(Trim(Request("txtPackingCd")))
	  	I5_m_pur_ord_hdr(M193_I2_inspect_means)		= UCase(Trim(Request("txtInspectCd")))
	  	I5_m_pur_ord_hdr(M193_I2_dischge_city)		= UCase(Trim(Request("txtDisCity")))
	  	I5_m_pur_ord_hdr(M193_I2_dischge_port)		= UCase(Trim(Request("txtDisPort")))
	  	I5_m_pur_ord_hdr(M193_I2_loading_port)		= UCase(Trim(Request("txtLoadPort")))
	  	I5_m_pur_ord_hdr(M193_I2_shipment)			= Trim(Request("txtShipMent"))
		
	End if     
    
  	I5_m_pur_ord_hdr(M193_I2_release_flg)			= "N"
    
    If lgIntFlgMode = OPMD_CMODE Then
		iStrCommandSent 							= "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
    	I5_m_pur_ord_hdr(M193_I2_po_no)				= UCase(Trim(Request("txtPoNo")))
		iStrCommandSent 							= "UPDATE"
    End If
	
	'-----------------------
	'Com Action Area
	'-----------------------
	
	'⊙: Lookup Pad 동작후 정상적인 데이타 이면, 저장 로직 시작 
    Set M31111 = Server.CreateObject("PM3G111.cMMaintPurOrdHdrS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set M31111 = Nothing 		
			Exit Sub
	End If

	 iStrPoNo = M31111.M_MAINT_PUR_ORD_HDR_SVR("F",gStrGlobalCollection, _
											  iStrCommandSent, _
											  I1_b_company, _
											  I2_m_config_process, _
											  I3_b_biz_partner, _
											  I4_b_pur_grp, _
											  I5_m_pur_ord_hdr)

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set M31111 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If

    Set M31111 = Nothing															'☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>"  								& vbCr
	Response.Write "With parent"	  											& vbCr	
	
	If lgIntFlgMode = OPMD_CMODE  Then 
		Response.Write "	.frm1.txtPoNo.Value	= """ & ConvSPChars(iStrPoNo) & """" & vbCr
	End If

	Response.Write "	.DbSaveOk" 	& vbCr
	Response.Write "End With" 		& vbCr
	Response.Write "</Script>"		& vbCr
				
    Set M31111 = Nothing															'☜: Unload Comproxy
	
End Sub																				'☜: Process End

'============================================================================================================
' Name : SubReleaseCheck()
' Desc : the case of Release,UnRelease
'============================================================================================================
Sub SubReleaseCheck()

	Dim M31211
	Dim strMode,lgIntFlgMode
	Dim txtSpread

	reDim IG1_import_group(0,2)
    Const M155_IG1_I1_select_char = 0 
    Const M155_IG1_I1_count = 1
    Const M155_IG1_I2_po_no = 2

	Dim prErrorPosition 
	Dim E3_m_pur_ord_hdr_po_no
	
    On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
	
    '-----------------------
    'Data manipulate area
    '-----------------------
	txtSpread = "U" & gColSep
    if strMode = "Release" then
		txtSpread = txtSpread & "Y" & gColSep
	else
		txtSpread = txtSpread & "N" & gColSep
	End if

	txtSpread = txtSpread & Trim(Request("txtPoNo")) & gColSep
	txtSpread = txtSpread & "1" & gRowSep

	'⊙: Lookup Pad 동작후 정상적인 데이타 이면, 저장 로직 시작 
	
    Set M31211 = Server.CreateObject("PM3G1R1.cMReleasePurOrdS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
			Set M31211 = Nothing 		
			Exit Sub
	End If

	'-----------------------
	'Com Action Area
	'-----------------------
	Call M31211.M_RELEASE_PUR_ORD_SVR(gStrGlobalCollection, _
									  , _
									  txtSpread, _
									  prErrorPosition)

    ls_msg = Trim(Cstr(Err.Description))
    
    Set M31211 = Nothing
    
    If Len(ls_msg) >= 6 Then
		If Right(ls_msg,6) = "TRGERR" Then
			Call DisplayMsgBox(Mid(ls_msg,Len(ls_msg) -11,6), "", "", "", I_MKSCRIPT)
			Response.End
		Else
			If CheckSYSTEMError2(Err,True,"","","","","") = true then 
				'Response.Write "<Script Language=VBScript>" & vbCr
				'Response.Write "parent.frm1.btnCfm.disabled = False" & vbCr
				'Response.Write "</Script>" & vbCr
				Response.End
			 End If
		End If
	Else		
		If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		    'Response.Write "<Script Language=VBScript>" & vbCr
		   ' Response.Write "parent.frm1.btnCfm.disabled = False" & vbCr
			'Response.Write "</Script>" & vbCr
			Response.End
		End If
	End If	 

	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" 					& vbCr
	Response.Write "With parent"									& vbCr	
    Response.Write "  If """ & strMode & """ = ""Release""  Then " 	& vbCr
	Response.Write "    .frm1.rdoRelease(0).Checked = true"        	& vbCr
	Response.Write"   Else "    									& vbCr
	Response.Write "    .frm1.rdoRelease(1).Checked = true"        	& vbCr
	Response.Write "  End if " & vbCr
	Response.Write ".DbSaveOk" & vbCr
	Response.Write "End With"   & vbCr
	Response.Write "</Script>" & vbCr
    
End Sub
'============================================================================================================
' Name : SubSendingB2B()
' Desc : the case of SendingB2B
'============================================================================================================
Sub SubSendingB2B()															'☜: B2B
   
'    Dim SDSOrder 
'    Dim arrParam
'	Dim strReturnMsg
'	Dim paramExt
'	Dim strServerName
'	
'    if UCase(Trim(gAPServer)) = "HEYJUDE" or UCase(Trim(gAPServer)) = "70.22.106.81" then
'    	strServerName = "Porsche"
'    elseif UCase(Trim(gAPServer)) = "PORSCHE" or UCase(Trim(gAPServer)) = "70.22.106.173" then
'    	strServerName = "HeyJude"
'    else
'    	Response.Write "Server Name" & UCase(Trim(gAPServer))
'    	Response.End
'    End if
'    
'	On Error Resume Next
'	Err.Clear
'	    
'    Set SDSOrder = Server.CreateObject("AAsdsorders10.csdsorders10")
'    
'    If CheckSYSTEMError(Err,True) = true then
'			Set SDSOrder = Nothing 		
'			Exit Sub
'	End If
'    
'    ReDim arrParam(0)
'        
'    arrParam(0) = Request("txtPoNo")
'    
'    strReturnMsg = SDSOrder.send(paramExt,strServerName,"sdsorders10", arrParam)
'    
'    If Trim(strReturnMsg) <> "" then
'		Set SDSOrder = Nothing
'		Call ServerMesgBox(strReturnMsg, vbInformation, I_MKSCRIPT)
'		Response.End
'    End if
'
'
'
'	Response.Write "<Script language=vbscript>" & vbCr
'	Response.Write "parent.SendingOK" & vbCr
'	Response.Write "</Script>" & vbcr
'
End SUb
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizDelete()																			'☜: 삭제 요청 
	
	On Error Resume Next
    Err.Clear																				'☜: Protect system from crashing
	
	Dim M31111
	Dim I5_m_pur_ord_hdr
	Const M193_I2_po_no = 0										

	Redim I5_m_pur_ord_hdr(76)

	I5_m_pur_ord_hdr(M193_I2_po_no)				= Trim(Request("txtPoNo"))

    Set M31111 = Server.CreateObject("PM3G111.cMMaintPurOrdHdrS")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set M31111 = Nothing 		
			Exit Sub
	End If
					   
	Call M31111.M_MAINT_PUR_ORD_HDR_SVR("F",gStrGlobalCollection, _
									  "DELETE", _
									  "", _
									  "", _
									  "", _
									  "", _
									  I5_m_pur_ord_hdr)

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set M31111 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If

    Set M31111 = Nothing															'☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Call parent.DbDeleteOk()" & vbCr
	Response.Write "</Script>" & vbCr
		
    Set M31111 = Nothing																	'☜: Unload Comproxy
												
End Sub
'============================================================================================================
' Name : SubLookUpPoType
' Desc : 
'============================================================================================================
Sub SubLookUpPoType()

	'EXPORTS View 상수 
	'View Name : exp m_config_process
	Const M369_E1_po_type_cd = 0
	Const M369_E1_po_type_nm = 1
	Const M369_E1_import_flg = 2
	Const M369_E1_bl_flg = 3
	Const M369_E1_cc_flg = 4
	Const M369_E1_rcpt_flg = 5
	Const M369_E1_subcontra_flg = 6
	Const M369_E1_ret_flg = 7
	Const M369_E1_iv_flg = 8
	Const M369_E1_usage_flg = 9
	Const M369_E1_rcpt_type = 10
	Const M369_E1_issue_type = 11
	Const M369_E1_iv_type = 12
	Const M369_E1_ext1_cd = 13
	Const M369_E1_ext2_cd = 14
	Const M369_E1_ext3_cd = 15
	Const M369_E1_ext4_cd = 16
	
	'View Name : exp_rcpt m_mvmt_type
	Const M369_E2_io_type_cd = 0
	Const M369_E2_io_type_nm = 1
	Const M369_E2_mvmt_cd = 2
	
	'View Name : exp_issue m_mvmt_type
	Const M369_E3_io_type_cd = 0
	Const M369_E3_io_type_nm = 1
	Const M369_E3_mvmt_cd = 2
	Const M369_E3_insrt_user_id = 3
	Const M369_E3_insrt_dt = 4
	Const M369_E3_updt_user_id = 5
	Const M369_E3_updt_dt = 6
	
	'View Name : exp m_iv_type
	Const M369_E4_iv_type_cd = 0
	Const M369_E4_iv_type_nm = 1
	Const M369_E4_trans_cd = 2

	Dim M14119
	Dim E1_m_config_process
	Dim E2_m_mvmt_type_rcpt
	Dim E3_m_mvmt_type_issue
	Dim E4_m_iv_type
	Dim iPoTypeCd
	
	Err.Clear
	On Error Resume Next

    Set M14119 = Server.CreateObject("PM1G419.cMLkConfigProcessS")    
    
	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End if
	
	iPoTypeCd = Trim(Request("txtPoTypeCd"))
    Call M14119.M_LOOKUP_CONFIG_PROCESS_SVR(gStrGlobalCollection, _
    									iPoTypeCd, _
    									E1_m_config_process, _
    									E2_m_mvmt_type_rcpt, _
    									E3_m_mvmt_type_issue, _
    									E4_m_iv_type)
    									
	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End if

	Set M14119 = Nothing												'☜: ComProxy Unload

	If E1_m_config_process(M369_E1_ret_flg) = "Y" Then
		Call DisplayMsgBox("17a014", vbOKOnly, "반품형태", "선택", I_MKSCRIPT)
		Set M14119 = Nothing

		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "with Parent.frm1" & vbCr
		Response.Write ".txtPotypeCd.value = """ & """" & vbCr
		Response.Write "end with"
		Response.Write"</Script>"

		Exit Sub
    End If
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with Parent.frm1" & vbCr

	Response.Write "	.hdnImportflg.Value		= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_import_flg))) 	& """" & vbCr
	Response.Write "	.hdnBLflg.Value			= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_bl_flg))) 		& """" & vbCr
	Response.Write "	.hdnCCflg.Value			= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_cc_flg)))		& """" & vbCr
	Response.Write "	.hdnRcptflg.Value		= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_rcpt_flg)))		& """" & vbCr
	Response.Write "	.hdnSubcontraflg.Value	= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_subcontra_flg)))	& """" & vbCr
	Response.Write "	.hdnRetflg.Value		= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_ret_flg)))		& """" & vbCr
	Response.Write "	.hdnIvflg.Value			= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_iv_flg)))		& """" & vbCr
	Response.Write "	.hdnRcpttype.Value		= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_rcpt_type)))		& """" & vbCr
	Response.Write "	.hdnIssueType.Value		= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_issue_type)))	& """" & vbCr
	Response.Write "	.hdnIvType.Value		= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_iv_type)))		& """" & vbCr
	Response.Write "	.txtPotypeNm.Value		= """ & ConvSPChars(Trim(E1_m_config_process(M369_E1_po_type_nm)))	& """" & vbCr

	if ConvSPChars(UCase(Trim(E1_m_config_process(M369_E1_import_flg)))) = "Y" then
	    Response.Write "Parent.ggoOper.SetReqAttr	.txtDvryDt, ""N"""	& vbCr	
	else     
		Response.Write "Parent.ggoOper.SetReqAttr	.txtDvryDt, ""D"""	& vbCr	
	end if	
	
	If UCase(Trim(Request("txtTabClickFlag"))) = "TRUE" Then
		Response.Write "	parent.lgOpenFlag	= True"	& vbCr	
		Response.Write "	Call parent.ClickTab2()" 	& vbCr
	End If
	
	Response.Write "end with" & vbCr	
	Response.Write "</Script>" & vbCr
	
	Set M14119 = Nothing
	
End Sub
'============================================================================================================
' Name : SubSupplierLookupAfterOnline
' Desc : 
'============================================================================================================
Sub SubSupplierLookupAfterOnline()
	Dim iSupplierCd
	On Error Resume Next
	Err.Clear

    Set B1H019 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
			Set B1H019 = Nothing 		
			Exit Sub
	End If
	
	iSupplierCd=Trim(Request("txtSupplierCd"))
	Call B1H019.B_LOOKUP_BIZ_PARTNER(gStrGlobalCollection, _
				                         iSupplierCd, _
				                         E1_b_biz_partner)
	
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
			Set B1H019 = Nothing 		
			Exit Sub
	End If

	Set B1H019 = Nothing 		


		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With Parent.frm1" & vbCr
		Response.Write "If Trim(.txtSupplierNm.Value) = """" Then " & vbCr 	
		Response.Write "  .txtSupplierNm.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_bp_nm))) & """" & vbCr
		Response.Write "End If" & vbCr
			
		Response.Write "If Trim(.txtCurr.Value) = """" Then " & vbCr 	
		Response.Write " .txtCurr.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_currency))) & """" & vbCr
		Response.Write "End if" & vbCr
			
		Response.Write "If Trim(.txtVatType.Value) = """" Then " & vbCr 	
		Response.Write " .txtVatType.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type))) & """" & vbCr
		Response.Write " .txtVatTypeNm.Value = """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type_nm))) & """" & vbCr
		Response.Write " .txtVatrt.text	= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write "End if" & vbCr
			
		Response.Write "If """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_inc_flag))) & """ = ""2"" Then" & vbCr	    'vat 포함 여부 
		Response.Write " .frm1.rdoVatFlg2.Checked= true "	& vbCr											'포함 
		Response.Write " .frm1.hdvatFlg.value 	= ""2"" " & vbCr
		Response.Write "Else " & vbCr
		Response.Write " .frm1.rdoVatFlg1.Checked= true "	& vbCr											'별도 
		Response.Write " .frm1.hdvatFlg.value 	= ""1"" " & vbCr
		Response.Write "End If" & vbCr	
		    
		Response.Write "If Trim(.txtPaytermCd.Value) = """" Then " & vbCr 	
		Response.Write " .txtPaytermCd.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_meth))) & """" & vbCr
		Response.Write " .txtPaytermNm.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_meth_nm))) & """" & vbCr
		Response.Write "End If" & vbCr
		
		Response.Write "If Trim(.txtPayDur.text) = """" Then " & vbCr 	
		Response.Write " .txtPayDur.text	= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_pay_dur),0 ,0) & """" & vbCr
		Response.Write "End If" & vbCr
		
		Response.Write "If Trim(.txtPayTermstxt.Value) = """" Then " & vbCr 	
		Response.Write " .txtPayTermstxt.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_terms_txt))) & """" & vbCr
		Response.Write "End If" & vbCr
		
		Response.Write "If Trim(.txtPayTypeCd.Value) = """" Then " & vbCr 	
		Response.Write " .txtPayTypeCd.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_type))) & """" & vbCr
		Response.Write " .txtPayTypeNm.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_type_nm))) & """" & vbCr
		Response.Write "End If " & vbCr

		Response.Write "If Trim(.txtSuppSalePrsn.Value) = """" Then " & vbCr 	
		Response.Write " .txtSuppSalePrsn.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_bp_prsn_nm))) & """" & vbCr
		Response.Write "End If " & vbCr
		
		Response.Write "If Trim(.txtTransCd.Value) = """" Then " & vbCr 	
		Response.Write " .txtTransCd.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_trans_meth))) & """" & vbCr
		Response.Write " .txtTransNm.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_trans_meth_nm))) & """" & vbCr
		Response.Write "End If" & vbCr
			
		Response.Write "End With " & vbCr	
		Response.Write "Parent.SetVatType() "	& vbCr													'부가세율을 현재 시점의 세율로 가져옴 
		Response.Write "</Script>" 		& vbCr

	
	Set B1H019 = Nothing
	
End Sub

'============================================================================================================
' Name : SubLookUpSupplier
' Desc : 
'============================================================================================================
Sub SubLookUpSupplier()
	
	On Error Resume Next
	Err.Clear
	
	Dim E1_b_biz_partner
	Const S074_E1_bp_cd = 0						
	Const S074_E1_bp_type = 1                   
	Const S074_E1_bp_rgst_no = 2                
	Const S074_E1_bp_full_nm = 3                
	Const S074_E1_bp_nm = 4                     
	Const S074_E1_bp_eng_nm = 5                 
	Const S074_E1_repre_nm = 6                  
	Const S074_E1_repre_rgst_no = 7             
	Const S074_E1_fnd_dt = 8                    
	Const S074_E1_zip_cd = 9                    
	Const S074_E1_addr1 = 10                    
	Const S074_E1_addr1_eng = 11                
	Const S074_E1_ind_type = 12                 
	Const S074_E1_ind_class = 13                
	Const S074_E1_trade_rgst_no = 14            
	Const S074_E1_contry_cd = 15                
	Const S074_E1_province_cd = 16              
	Const S074_E1_currency = 17                 
	Const S074_E1_tel_no1 = 18                  
	Const S074_E1_tel_no2 = 19                  
	Const S074_E1_fax_no = 20                   
	Const S074_E1_home_url = 21                 
	Const S074_E1_usage_flag = 22               
	Const S074_E1_bp_prsn_nm = 23               
	Const S074_E1_bp_contact_pt = 24            
	Const S074_E1_biz_prsn = 25                 
	Const S074_E1_biz_grp = 26                  
	Const S074_E1_biz_org = 27                  
	Const S074_E1_deal_type = 28                
	Const S074_E1_pay_meth = 29                 
	Const S074_E1_pay_dur = 30                  
	Const S074_E1_pay_day = 31                  
	Const S074_E1_vat_inc_flag = 32             
	Const S074_E1_vat_type = 33                 
	Const S074_E1_vat_rate = 34                 
	Const S074_E1_trans_meth = 35               
	Const S074_E1_trans_lt = 36                 
	Const S074_E1_sale_amt = 37                 
	Const S074_E1_capital_amt = 38              
	Const S074_E1_emp_cnt = 39                  
	Const S074_E1_bp_grade = 40                 
	Const S074_E1_comm_rate = 41                
	Const S074_E1_addr2 = 42                    
	Const S074_E1_addr2_eng = 43                
	Const S074_E1_addr3_eng = 44                
	Const S074_E1_pay_type = 45                 
	Const S074_E1_pay_terms_txt = 46            
	Const S074_E1_credit_mgmt_flag = 47         
	Const S074_E1_credit_grp = 48               
	Const S074_E1_vat_calc_type = 49            
	Const S074_E1_deposit_flag = 50             
	Const S074_E1_bp_group = 51                 
	Const S074_E1_clearance_id = 52             
	Const S074_E1_credit_rot_day = 53           
	Const S074_E1_gr_insp_type = 54             
	Const S074_E1_bp_alias_nm = 55              
	Const S074_E1_to_org = 56                   
	Const S074_E1_to_grp = 57                   
	Const S074_E1_pay_month = 58                
	Const S074_E1_expiry_dt = 59                
	Const S074_E1_pur_grp = 60                  
	Const S074_E1_pur_org = 61                  
	Const S074_E1_charge_lay_flag = 62          
	Const S074_E1_remark1 = 63                  
	Const S074_E1_remark2 = 64                  
	Const S074_E1_remark3 = 65                  
	Const S074_E1_close_day1 = 66               
	Const S074_E1_close_day2 = 67               
	Const S074_E1_close_day3 = 68               
	Const S074_E1_tax_biz_area = 69             
	Const S074_E1_cash_rate = 70                
	Const S074_E1_pay_type_out = 71             
	Const S074_E1_par_bank_cd1_bp = 72          
	Const S074_E1_bank_acct_no1_bp = 73         
	Const S074_E1_bank_cd1_bp = 74              
	Const S074_E1_par_bank_cd2_bp = 75          
	Const S074_E1_bank_cd2_bp = 76              
	Const S074_E1_bank_acct_no2_bp = 77         
	Const S074_E1_par_bank_cd3_bp = 78          
	Const S074_E1_bank_cd3_bp = 79              
	Const S074_E1_bank_acct_no3_bp = 80         
	Const S074_E1_par_bank_cd1 = 81             
	Const S074_E1_bank_cd1 = 82                 
	Const S074_E1_bank_acct_no1 = 83            
	Const S074_E1_par_bank_cd2 = 84             
	Const S074_E1_bank_cd2 = 85                 
	Const S074_E1_bank_acct_no2 = 86            
	Const S074_E1_par_bank_cd3 = 87             
	Const S074_E1_bank_cd3 = 88                 
	Const S074_E1_bank_acct_no3 = 89            
	Const S074_E1_pay_month2 = 90               
	Const S074_E1_pay_day2 = 91                 
	Const S074_E1_pay_month3 = 92               
	Const S074_E1_pay_day3 = 93                 
	Const S074_E1_close_day1_sales = 94         
	Const S074_E1_pay_month1_sales = 95         
	Const S074_E1_pay_day1_sales = 96           
	Const S074_E1_close_day2_sales = 97         
	Const S074_E1_pay_month2_sales = 98         
	Const S074_E1_pay_day2_sales = 99           
	Const S074_E1_close_day3_sales = 100        
	Const S074_E1_pay_month3_sales = 101        
	Const S074_E1_pay_day3_sales = 102          
	Const S074_E1_ext1_qty = 103                
	Const S074_E1_ext2_qty = 104                
	Const S074_E1_ext3_qty = 105                
	Const S074_E1_ext1_amt = 106                
	Const S074_E1_ext2_amt = 107                
	Const S074_E1_ext3_amt = 108                
	Const S074_E1_ext1_cd = 109                 
	Const S074_E1_ext2_cd = 110                 
	Const S074_E1_ext3_cd = 111                 
	Const S074_E1_in_out = 112                  
	Const S074_E1_card_co_cd = 113              
	Const S074_E1_card_mem_no = 114             
	Const S074_E1_pay_meth_pur = 115            
	Const S074_E1_pay_type_pur = 116            
	Const S074_E1_pay_dur_pur = 117             
	Const S074_E1_bank_cd = 118                 
	Const S074_E1_bank_acct_no = 119            
	Const S074_E1_ind_type_nm = 120             
	Const S074_E1_ind_class_nm = 121            
	Const S074_E1_bp_group_nm = 122             
	Const S074_E1_b_country_nm = 123            
	Const S074_E1_b_province_nm = 124           
	Const S074_E1_trans_meth_nm = 125           
	Const S074_E1_deal_type_nm = 126            
	Const S074_E1_bp_grade_nm = 127             
	Const S074_E1_s_credit_limit = 128          
	Const S074_E1_b_sales_grp_nm = 129          
	Const S074_E1_b_to_grp_nm = 130             
	Const S074_E1_b_pur_grp_nm = 131            
	Const S074_E1_vat_type_nm = 132       
	Const S074_E1_pay_meth_nm = 133       
	Const S074_E1_pay_type_nm = 134       
	Const S074_E1_tax_area_nm = 135       
	Const S074_E1_b_zip_code = 136        
	Const S074_E1_b_pur_org = 137         
	Const S074_E1_b_pur_org_nm = 138      
	Const S074_E1_vat_inc_flag_nm = 139   
	Const S074_E1_card_co_cd_nm = 140     
	Const S074_E1_pay_meth_pur_nm = 141   
	Const S074_E1_pay_type_pur_nm = 142   
	Const S074_E1_bank_cd_nm = 143        

	Dim iSupplierCd
	Dim B1H019
		
    Set B1H019 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
			Set B1H019 = Nothing 		
			Exit Sub
	End If
	
	iSupplierCd = Trim(Request("txtSupplierCd"))
	Call B1H019.B_LOOKUP_BIZ_PARTNER(gStrGlobalCollection, _
				                         iSupplierCd, _
				                         E1_b_biz_partner)
	
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
			Set B1H019 = Nothing 		
			Exit Sub
	End If

	Set B1H019 = Nothing 		

		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "with Parent.frm1" & vbCr
	
		Response.Write ".txtSupplierNm.Value		= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_bp_nm))) & """" & vbCr
		Response.Write ".txtCurr.Value				= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_currency))) & """" & vbCr
		Response.Write ".txtCurrNm.Value			= """" " & vbCr
		Response.Write ".txtVatType.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type))) & """" & vbCr
		Response.Write ".txtVatTypeNm.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type_nm))) & """" & vbCr
		Response.Write ".txtVatrt.text				= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write ".txtPaytermCd.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_meth_pur))) & """" & vbCr
		Response.Write ".txtPaytermNm.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_meth_pur_nm))) & """" & vbCr
		Response.Write ".txtPayDur.text				= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_pay_dur_pur),0 ,0) & """" & vbCr 
		Response.Write ".txtPayTermstxt.Value		= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_terms_txt))) & """" & vbCr
'		Response.Write ".txtPayTypeCd.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_type))) & """" & vbCr
'		Response.Write ".txtPayTypeNm.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_type_nm))) & """" & vbCr

		Response.Write ".txtPayTypeCd.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_type_pur))) & """" & vbCr
		Response.Write ".txtPayTypeNm.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pay_type_pur_nm))) & """" & vbCr

''
		Response.Write ".txtSuppSalePrsn.Value		= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_bp_prsn_nm))) & """" & vbCr
		Response.Write ".txtTransCd.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_trans_meth))) & """" & vbCr
		Response.Write ".txtTransNm.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_trans_meth_nm))) & """" & vbCr	'구매그룹 추가 
		Response.Write ".txtGroupCd.value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pur_grp))) & """" & vbCr
		Response.Write ".txtGroupNm.value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_b_pur_grp_nm))) & """" & vbCr
        Response.Write ".txtTel.value			    = """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_bp_contact_pt))) & """" & vbCr
        
	    if ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_inc_flag))) = "2" then
	    Response.Write " .rdoVatFlg2.Checked= true " & vbCr
		Response.Write " .hdvatFlg.value 	= ""2"" " & vbCr
	    Else
		Response.Write " .rdoVatFlg1.Checked= true"   & vbCr 
		Response.Write " .hdvatFlg.value 	= ""0"" " & vbCr
	    End if
  		
		Response.Write "End With" & vbCr	
		Response.Write "Parent.ChangeCurr()" & vbCr
		Response.Write "Parent.SetVatType()" & vbCr								'부가세율을 현재 시점의 세율로 가져옴 
		Response.Write "</Script>" & vbCr   				
	
	Set B1H019 = Nothing
	
End Sub
	
%>
































































