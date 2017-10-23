<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*							  M31119(Lookup_PO_Hdr)
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Shin jin hyun
'* 10. Modifier (Last)      : Min, HJ
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'* 14. Business Logic for m3111ma1(발주일반(무역)정보등록)
'**********************************************************************************************
	'Dim lgOpModeCRUD
	Dim ls_msg
	
	'이성룡 추가(단가Type)
	Dim iPriceType  	
	
	On Error Resume Next
	Err.Clear 
																
	Call HideStatusWnd
	
	lgOpModeCRUD	=	Request("txtMode")
	
	Select Case lgOpModeCRUD		
	   Case CStr(UID_M0001)																		'☜: Query
  	      Call SubBizQueryMulti()
	   Case CStr(UID_M0002)	
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
	   Case "lookupPrice"
		  Call SublookupPrice()
	   Case "lookupPriceForSelection"			
		  Call lookupPriceForSelection()
	   Case "LookUpSupplier"
		  Call SubLookUpSupplier()
	   Case "LookUpItemPlant"
		  Call SubLookUpItemPlant()
	   Case "LookUpItemPlantForUnit"
		  Call SubLookUpItemPlantForUnit()
	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Dim M31128																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim M31119
	Dim iMax
	Dim PvArr

	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	Const C_SHEETMAXROWS_D  = 100
	
	Dim E1_b_bank
	Dim E2_b_minor_vat_type
	Dim E3_b_minor_pay_meth
	Dim E4_b_minor_pay_type
	Dim E5_b_minor_incoterms
	Dim E6_b_minor_transport
	Dim E7_b_minor_delivery_plce
	Dim E8_b_minor_origin
	Dim E9_b_biz_partner
	Dim E10_b_biz_partner
	Dim E11_b_biz_partner
	Dim E12_b_minor_packing_cond
	Dim E13_b_minor_inspect_means
	Dim E14_b_minor_dischge_city
	Dim E15_b_minor_dischge_port
	Dim E16_b_minor_loading_port
	Dim E17_b_configuration_reference
	Dim E18_b_currency
	Dim E19_m_config_process
	Dim E20_b_pur_grp
	Dim E21_b_biz_partner
	Dim E22_m_pur_ord_hdr
   
    Const M192_EG1_E1_m_pur_req_pr_no = 0
    Const M192_EG1_E2_b_biz_area_biz_area_cd = 1
    Const M192_EG1_E2_b_biz_area_biz_area_nm = 2
    Const M192_EG1_E3_b_hs_code_hs_cd = 3
    Const M192_EG1_E3_b_hs_code_hs_nm = 4
    Const M192_EG1_E4_b_storage_location_sl_cd = 5
    Const M192_EG1_E4_b_storage_location_sl_type = 6
    Const M192_EG1_E4_b_storage_location_sl_nm = 7
    Const M192_EG1_E5_m_pur_ord_dtl_po_seq_no = 8
    Const M192_EG1_E5_m_pur_ord_dtl_dlvy_dt = 9
    Const M192_EG1_E5_m_pur_ord_dtl_po_qty = 10
    Const M192_EG1_E5_m_pur_ord_dtl_po_unit = 11
    Const M192_EG1_E5_m_pur_ord_dtl_po_prc = 12
    Const M192_EG1_E5_m_pur_ord_dtl_po_prc_flg = 13
    Const M192_EG1_E5_m_pur_ord_dtl_po_doc_amt = 14
    Const M192_EG1_E5_m_pur_ord_dtl_sl_cd = 15
    Const M192_EG1_E5_m_pur_ord_dtl_lc_qty = 16
    Const M192_EG1_E5_m_pur_ord_dtl_hs_cd = 17
    Const M192_EG1_E5_m_pur_ord_dtl_rcpt_biz_area = 18
    Const M192_EG1_E5_m_pur_ord_dtl_tracking_no = 19
    Const M192_EG1_E5_m_pur_ord_dtl_po_base_qty = 20
    Const M192_EG1_E5_m_pur_ord_dtl_po_base_unit = 21
    Const M192_EG1_E5_m_pur_ord_dtl_fr_trans_coef = 22
    Const M192_EG1_E5_m_pur_ord_dtl_to_trans_coef = 23
    Const M192_EG1_E5_m_pur_ord_dtl_po_loc_amt = 24
    Const M192_EG1_E5_m_pur_ord_dtl_rcpt_qty = 25
    Const M192_EG1_E5_m_pur_ord_dtl_iv_qty = 26
    Const M192_EG1_E5_m_pur_ord_dtl_bl_qty = 27
    Const M192_EG1_E5_m_pur_ord_dtl_cc_qty = 28
    Const M192_EG1_E5_m_pur_ord_dtl_po_sts = 29
    Const M192_EG1_E5_m_pur_ord_dtl_cls_flg = 30
    Const M192_EG1_E5_m_pur_ord_dtl_ref_po_no = 31
    Const M192_EG1_E5_m_pur_ord_dtl_ref_po_seq_no = 32
    Const M192_EG1_E5_m_pur_ord_dtl_over_tol = 33
    Const M192_EG1_E5_m_pur_ord_dtl_under_tol = 34
    Const M192_EG1_E5_m_pur_ord_dtl_so_no = 35
    Const M192_EG1_E5_m_pur_ord_dtl_so_seq_no = 36
    Const M192_EG1_E5_m_pur_ord_dtl_after_lc_flg = 37
    Const M192_EG1_E5_m_pur_ord_dtl_inspect_flg = 38
    Const M192_EG1_E5_m_pur_ord_dtl_ext1_cd = 39
    Const M192_EG1_E5_m_pur_ord_dtl_ext1_qty = 40
    Const M192_EG1_E5_m_pur_ord_dtl_ext1_amt = 41
    Const M192_EG1_E5_m_pur_ord_dtl_ext1_rt = 42
    Const M192_EG1_E5_m_pur_ord_dtl_ext2_cd = 43
    Const M192_EG1_E5_m_pur_ord_dtl_ext2_qty = 44
    Const M192_EG1_E5_m_pur_ord_dtl_ext2_amt = 45
    Const M192_EG1_E5_m_pur_ord_dtl_ext2_rt = 46
    Const M192_EG1_E5_m_pur_ord_dtl_ext3_cd = 47
    Const M192_EG1_E5_m_pur_ord_dtl_ext3_qty = 48
    Const M192_EG1_E5_m_pur_ord_dtl_ext3_amt = 49
    Const M192_EG1_E5_m_pur_ord_dtl_ext3_rt = 50
    Const M192_EG1_E5_m_pur_ord_dtl_vat_type = 51
    Const M192_EG1_E5_m_pur_ord_dtl_vat_rate = 52
    Const M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt = 53
    Const M192_EG1_E5_m_pur_ord_dtl_vat_loc_amt = 54
    Const M192_EG1_E5_m_pur_ord_dtl_ret_type = 55
    Const M192_EG1_E5_m_pur_ord_dtl_ref_mvmt_no = 56
    Const M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag = 57
    Const M192_EG1_E5_m_pur_ord_dtl_tot_rcpt_doc_amt = 58
    Const M192_EG1_E5_m_pur_ord_dtl_tot_rcpt_loc_amt = 59
    Const M192_EG1_E5_m_pur_ord_dtl_lot_no = 60
    Const M192_EG1_E5_m_pur_ord_dtl_lot_sub_no = 61
    Const M192_EG1_E5_m_pur_ord_dtl_ref_iv_no = 62
    Const M192_EG1_E5_m_pur_ord_dtl_ref_iv_seq = 63
    Const M192_EG1_E6_b_item_item_cd = 64
    Const M192_EG1_E6_b_item_item_nm = 65
    Const M192_EG1_E6_b_item_spec = 66
    Const M192_EG1_E7_b_plant_plant_cd = 67
    Const M192_EG1_E7_b_plant_plant_nm = 68
    Const M192_EG1_E8_b_minor_minor_nm = 69
    Const M192_EG1_E9_b_minor_minor_nm = 70
    
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
    Const M192_EG1_E5_m_pur_ord_dtl_remrk = 71
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------    

	Dim E1_m_pur_ord_dtl_po_seq_no

	Dim EG1_exp_group
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
    'Const M239_E22_cls_flg = 71
    Const M239_E22_cls_flg = 48			'변경 (이정태)
    Const M239_E22_ref_no = 49			'추가 (이정태)
    Const M239_E22_tot_vat_loc_amt = 50
    '
    Const M239_E19_po_type_cd = 0
    Const M239_E19_po_type_nm = 1
    '
    Const M239_E9_bp_cd = 0
    Const M239_E9_bp_nm = 1
    '
    Const M239_E20_pur_grp = 0
    Const M239_E20_pur_grp_nm = 1

	
	
	Dim iStrPoNo

    On Error Resume Next
    Err.Clear                                                               '☜: Protect system from crashing
    
	'====================
	'Call PO_HDR LOOK UP
	'====================
    Set M31119 = Server.CreateObject("PM3G119.cMLookupPurOrdHdrS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set M31119 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	iStrPoNo = Trim(Request("txtPoNo"))
	Call M31119.M_LOOKUP_PUR_ORD_HDR_SVR(gStrGlobalCollection, _
                        			  iStrPoNo, _
	                                  E1_b_bank, _
	                                  E2_b_minor_vat_type, _
	                                  E3_b_minor_pay_meth, _
	                                  E4_b_minor_pay_type, _
	                                  E5_b_minor_incoterms, _
	                                  E6_b_minor_transport, _
	                                  E7_b_minor_delivery_plce, _
	                                  E8_b_minor_origin, _
	                                  E9_b_biz_partner, _
	                                  E10_b_biz_partner, _
	                                  E11_b_biz_partner, _
	                                  E12_b_minor_packing_cond, _
	                                  E13_b_minor_inspect_means, _
	                                  E14_b_minor_dischge_city, _
	                                  E15_b_minor_dischge_port, _
	                                  E16_b_minor_loading_port, _
	                                  E17_b_configuration_reference, _
	                                  E18_b_currency, _
	                                  E19_m_config_process, _
	                                  E20_b_pur_grp, _
	                                  E21_b_biz_partner, _
	                                  E22_m_pur_ord_hdr)
   
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set M31119 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If

	If Ucase(Trim(E22_m_pur_ord_hdr(M239_E22_ret_flg))) = "Y" Then
       Call DisplayMsgBox("17a014", vbOKOnly, "반품발주건", "조회", I_MKSCRIPT)
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
	
	If trim(Ucase(E22_m_pur_ord_hdr(M239_E22_release_flg))) = "Y" Then
		Response.Write "	.rdoRelease(1).Checked = true" & vbCr
		Response.Write "	.rdoRelease(0).Checked = false" & vbCr
	Else
		Response.Write "	.rdoRelease(0).Checked = true" & vbCr
		Response.Write "	.rdoRelease(1).Checked = false" & vbCr
	End If
	
	If Ucase(Trim(E22_m_pur_ord_hdr(M239_E22_merg_pur_flg))) = "Y" Then
		Response.Write "	.rdoMergPurFlg(0).Checked = true" & vbCr
	Else
		Response.Write "	.rdoMergPurFlg(1).Checked = true" & vbCr
	End If

	Response.Write "	.txtPoNo2.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_po_no))) 			& """"	& vbCr
	'Response.Write "	.hdnPoNo.value				= """ & iStrPoNo 			& """"	& vbCr
	Response.Write "	.txtPoDt.Text				= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_po_dt)) 	& """"	& vbCr				
	Response.Write "	.txtSupplierCd.value		= """ & ConvSPChars(Trim(E9_b_biz_partner(M239_E9_bp_cd))) 			& """"	& vbCr
	Response.Write "	.txtSupplierNm.value		= """ & ConvSPChars(Trim(E9_b_biz_partner(M239_E9_bp_nm))) 			& """"	& vbCr
	Response.Write "	.txtGroupCd.value			= """ & ConvSPChars(Trim(E20_b_pur_grp(M239_E20_pur_grp))) 			& """"	& vbCr
	Response.Write "	.txtGroupNm.value			= """ & ConvSPChars(Trim(E20_b_pur_grp(M239_E20_pur_grp_nm))) 		& """"	& vbCr					
	'Response.Write "	.txtCurrNm.value			= """ & ConvSPChars(Trim(E18_b_currency_currency_desc)) 							& """"	& vbCr
	Response.Write "	.txtXch.value				= """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_xch_rt),ggExchRate.DecPoint,0) & """"	& vbCr
'	Response.Write "	.hdnxchrateop.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_xch_rate_op)))			& """" & vbCr 'Multi Divide 
	Response.Write "	.cboXchop.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_xch_rate_op)))			& """" & vbCr 'Multi Divide
	Response.Write "	.hdnxchrateop.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_xch_rate_op)))			& """" & vbCr 'Multi Divide  
	Response.Write "	.txtVatType.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_vat_type))) 		& """"	& vbCr
	Response.Write "	.txtVatTypeNm.Value			= """ & ConvSPChars(Trim(E2_b_minor_vat_type)) 						& """"	& vbCr
	Response.Write "	.txtVatRt.Text				= """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_vat_rt),ggExchRate.DecPoint,0) & """"	& vbCr		
	Response.Write "	.txtVatAmt.Text             = """ & UNIConvNumDBToCompanyByCurrency(E22_m_pur_ord_hdr(M239_E22_tot_vat_doc_amt),lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo  , "X") & """"	& vbCr
'	Response.Write "	.txtVatLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E22_m_pur_ord_hdr(M239_E22_tot_vat_loc_amt),lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo  , "X") & """"	& vbCr
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
	Response.Write "	.txtSuppSalePrsn.Value		= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_sppl_sales_prsn)))& """"	& vbCr
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
		Response.Write "	.txtBankNm.value			= """ & ConvSPChars(Trim(E1_b_bank))											& """" & vbCr
		Response.Write "	.txtDvryplce.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_delivery_plce)))			& """" & vbCr
		Response.Write "	.txtDvryplceNm.value		= """ & ConvSPChars(Trim(E7_b_minor_delivery_plce))							& """" & vbCr
		Response.Write "	.txtApplicantCd.Value		= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_applicant)))				& """" & vbCr
		Response.Write "	.txtApplicantNm.Value		= """ & ConvSPChars(Trim(E10_b_biz_partner))									& """" & vbCr
		Response.Write "	.txtManuCd.value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_manufacturer)))			& """" & vbCr			
		Response.Write "	.txtManuNm.Value			= """ & ConvSPChars(Trim(E11_b_biz_partner))									& """" & vbCr
		Response.Write "	.txtAgentCd.Value			= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_agent)))					& """" & vbCr
		Response.Write "	.txtAgentNm.Value			= """ & ConvSPChars(Trim(E21_b_biz_partner))									& """" & vbCr
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
		Response.Write "	.hdnSupplierCd.Value		= """ & ConvSPChars(Trim(E9_b_biz_partner(M239_E9_bp_cd)))				& """" & vbCr
		Response.Write "	.hdnPoDt.Value				= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_po_dt))			& """" & vbCr
		Response.Write "	.hdnCurr.value				= """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_po_cur))) & """"	& vbCr
		''cboXchop
		If ConvSPChars(E22_m_pur_ord_hdr(M239_E22_vat_inc_flag)) = "2" Then
		Response.Write "	.hdnVATINCFLG.value = ""2""" & vbCr	'포함 
	Else
		Response.Write "	.hdnVATINCFLG.value = ""*""" & vbCr	'별도 
	End If

	Response.Write "  parent.Setreference		  "			&	vbcr
	Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet"	& vbCr
	Response.Write "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE"	& vbCr
	Response.Write "	.hdclsflg.value			    = """ & ConvSPChars(Trim(E22_m_pur_ord_hdr(M239_E22_cls_flg)))			& """" & vbCr 'Multi Divide 
'	Response.Write "	parent.Setreference		  "			&	vbcr
	Response.Write "	parent.Changeflg()		  "			&	vbcr
'	Response.write "	parent.ChangeCurr() "	& vbCr
    Response.Write "End With" 	& vbCr
    Response.Write "</Script>" 	& vbCr

    Set M31119 = Nothing															'☜: Unload Comproxy
		
	'====================
	'Call PO_DTL List
	'====================
	lgStrPrevKey = Request("lgStrPrevKey")
	Set M31128 = Server.CreateObject("PM3G128.cMListPurOrdDtlS")     
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
		Exit Sub
	End If

	Call M31128.M_LIST_PUR_ORD_DTL_SVR(gStrGlobalCollection, _
						            C_SHEETMAXROWS_D, _
						            iStrPoNo, _
						            lgStrPrevKey, _
						            EG1_exp_group, _
						            E1_m_pur_ord_dtl_po_seq_no)
	  
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set M31128 = Nothing												'☜: ComProxy Unload
			'Detail항목이 없을 경우 Header정보만 보여줌 
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	End if
				
	Set M31128 = Nothing

	
   	iLngMaxRow = CLng(Request("txtMaxRows"))
			
	iMax = UBound(EG1_exp_group,1)
	ReDim PvArr(iMax)
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr

	For iLngRow = 0 To UBound(EG1_exp_group,1)
        
        If iLngRow >= C_SHEETMAXROWS_D Then
			StrNextKey = ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_seq_no))
        	Exit For
        End If
        
		'istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_seq_no))	'1
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E7_b_plant_plant_cd))		'2
        istrData = istrData & Chr(11) & ""														'3
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E7_b_plant_plant_nm))		'4
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E6_b_item_item_cd))			'5
        istrData = istrData & Chr(11) & ""														'6
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E6_b_item_item_nm))			'7
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E6_b_item_spec))	'품목규격 '8	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_qty),ggQty.DecPoint,0) '9	       
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_unit))	'10
        istrData = istrData & Chr(11) & ""														'11
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc),lgCurrency,ggUnitCostNo, "X" , "X")		'12		
        istrData = istrData & Chr(11) & ""
        If UCase(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc_flg))) = "F" Then				'13
			istrData = istrData & Chr(11) & "가단가"
		ElseIf UCase(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc_flg))) = "T" Then
			istrData = istrData & Chr(11) & "진단가"
		End If
'        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc))  '14
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc_flg))  '14
        
        If Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag)) = "2" Then				'13
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt)) + CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt)),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		Else
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		End If
		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '16
		istrData = istrData & Chr(11) & ""  '17		
		
		If Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag)) = "2" Then				'13
			istrData = istrData & Chr(11) & "포함"
		Else
			istrData = istrData & Chr(11) & "별도"
		End If
		
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))	'16
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_type))		'17
		istrData = istrData & Chr(11) & ""															'18
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E8_b_minor_minor_nm))			'19
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_rate),ggExchRate.DecPoint,0)	'20
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt),ggExchRate.DecPoint,0)	'20
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt),ggExchRate.DecPoint,0)	'21	
		istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_dlvy_dt))	'22
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E3_b_hs_code_hs_cd))			'23
		istrData = istrData & Chr(11) & ""															'24
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E3_b_hs_code_hs_nm))				'25
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E4_b_storage_location_sl_cd))			'26
		istrData = istrData & Chr(11) & ""															'27
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E4_b_storage_location_sl_nm))	'28
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_tracking_no))	'29
		istrData = istrData & Chr(11) & ""															'30
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_lot_no))			'31
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_lot_sub_no))		'32
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_ret_type))		'33
        istrData = istrData & Chr(11) & ""															'34
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E9_b_minor_minor_nm))			'35
    
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_over_tol),ggExchRate.DecPoint,0)	'36	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_under_tol),ggExchRate.DecPoint,0)	'37
        
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ""
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_seq_no))		'1  행복사시 문제로 위치 이동 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E1_m_pur_req_pr_no))				'38
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_ref_mvmt_no))		'39
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_ref_po_no))		'40
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_ref_po_seq_no))	'41
        istrData = istrData & Chr(11) & ""															'42
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_so_no))			'43
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_so_seq_no))		'44
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '17
        istrData = istrData & Chr(11) & "N"
        istrData = istrData & Chr(11) & ""			
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_remrk))				'71 
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)  
		PvArr(iLngRow) = istrData
		istrData=""
    Next

	istrData = Join(PvArr, "")
   
   Response.Write " If .frm1.vspdData.MaxRows < 1 then"						& vbCr
   Response.Write "	    If .frm1.hdnRelease.Value = ""Y"" then"				& vbCr
   Response.Write "	        For index = .C_SeqNo to .C_Stateflg"			& vbCr
   Response.Write "			        .ggoSpread.SpreadLock index , -1"		& vbCr
   Response.Write "		    Next"					& vbCr
   Response.Write "	    Else"						& vbCr
   Response.Write "		    .SetSpreadLock"			& vbCr
   Response.Write "		End If"						& vbCr
   Response.Write "	End if"							& vbCr
   Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
    Response.Write "	.frm1.txtDetailVatAmt.Text             = """ & UNIConvNumDBToCompanyByCurrency(E22_m_pur_ord_hdr(M239_E22_tot_vat_doc_amt),lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo , "X") & """"	& vbCr
    Response.Write "	.frm1.txtDetailNetAmt.Text				= """ & UNIConvNumDBToCompanyByCurrency(E22_m_pur_ord_hdr(M239_E22_tot_po_doc_amt),lgCurrency,ggAmtOfMoneyNo , "X" , "X") & """"	& vbCr
    Response.Write "	.frm1.txtDetailGrossAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(UniCDbl(E22_m_pur_ord_hdr(M239_E22_tot_po_doc_amt),0)+UniCDbl(E22_m_pur_ord_hdr(M239_E22_tot_vat_doc_amt),0),lgCurrency,ggAmtOfMoneyNo , "X" , "X") & """"	& vbCr

   Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
   Response.Write "	.lgStrPrevKey           = """ & StrNextKey   & """" & vbCr  
   Response.Write "	.frm1.txtPoNo2.value				= """ & ConvSPChars(Request("txtPoNo")) & """" & vbCr
'   Response.Write "If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> """" Then " & vbCr
'   Response.Write "	.SetSpreadLockAfterQuery "  & vbCr
'   Response.Write "	.DbQuery "  & vbCr	' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
'   Response.Write "Else " 			& vbCr
   Response.Write " .DbQueryOk "	& vbCr 
   'Response.Write "End If "		& vbCr
   Response.Write "End With"		& vbCr
   Response.Write "</Script>"		& vbCr    

	
   Set M31128 = Nothing
End Sub    																		'☜: Process End
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSave()	
		
    On Error Resume Next
    Err.Clear	
    
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
   
'm_pur_ord_dtl const
    Const L1_status = 0
    Const L1_seq_no = 1
    Const L1_plant_cd = 2           '공장 
    Const L1_popup1 = 3
    Const L1_plant_nm = 4          '공장명 
    Const L1_item_cd = 5           '품목 
    Const L1_popup2 = 6
    Const L1_item_nm = 7            '품목명 
    Const L1_sppl_spec = 8          '품목규격 
    Const L1_order_qty = 9          '발주수량 
    Const L1_order_unit = 10        '단위 
    Const L1_popup3 = 11
    Const L1_cost = 12              '단가 
    Const L1_cost_con = 13          '단가구분 
    Const L1_cost_con_cd = 14       '단가구분코드 
    Const L1_order_amt = 15         '금액 
    Const L1_io_flg = 16            'VAT포함여부 
    Const L1_io_flg_cd = 17         'VAT포함여부코드 
    Const L1_vat_type = 18          'VAT
    Const L1_popup7 = 19
    Const L1_vat_nm = 20            'VAT명 
    Const L1_vat_rate = 21          'VAT율(%)
    Const L1_vat_amt = 22           'VAT금액 
    Const L1_dlvy_dt = 23           '납기일 
    Const L1_hs_cd = 24             'HS부호 
    Const L1_popup5 = 25
    Const L1_hs_nm = 26             'HS명 
    Const L1_sl_cd = 27             '창고 
    Const L1_popup6 = 28
    Const L1_sl_nm = 29             '창고명 
    Const L1_tracking_no = 30       'Tracking No.
    Const L1_tracking_popup = 31
    Const L1_lot_no = 32            'Lot No.
    Const L1_lot_seq = 33           'Lot No.순번 
    Const L1_ret_cd = 34            '반품유형 
    Const L1_popup8 = 35
    Const L1_ret_nm = 36            '반품유형명 
    Const L1_over = 37              '과부족허용율(+)(%)
    Const L1_under = 38             '과부족허용율(-)(%)
    Const L1_bal_qty = 39           'Bal. Qty.
    Const L1_bal_doc_amt = 40       'Bal. Doc. Amt.
    Const L1_bal_loc_amt = 41       'Bal. Loc. Amt.
    Const L1_ex_rate = 42           'Xch. Rate
    Const L1_pr_no = 43             '구매요청번호 
    Const L1_mvmt_no = 44           '구매입고번호 
'-------------------------------------------------
'-------------------------------------------------
    Const L1_remrk = 45				'비고 
'-------------------------------------------------
'-------------------------------------------------    
    Const L1_po_no = 46             '발주번호 
    Const L1_po_seq_no = 47         '발주SeqNo
    Const L1_maint_seq = 48         'maintseq
    Const L1_so_no = 49
    Const L1_so_seq_no = 50
    Const L1_state_flg = 51
    Const L1_row_num = 52

		
	Dim M31111
	Dim lgIntFlgMode
	Dim lgBlnFlgChgValue
	Dim lgSSCheckValue
	Dim iStrCommandSent
	Dim iStrPoNo
	Dim E1_m_pur_ord_hdr_po_no
	Dim I1_b_company
	Dim I2_m_config_process
	Dim I3_b_biz_partner
	Dim I4_b_pur_grp
	Dim I5_m_pur_ord_hdr

	'Po Dtl 변수 
	Dim M31121																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim LngMaxRow
	Dim iErrorPosition
	Dim iStrSpread
	
	Dim itxtSpread
    	Dim itxtSpreadArr
    	Dim itxtSpreadArrCount

    	Dim iCUCount
    	Dim iDCount
    	Dim ii
	
	Redim I5_m_pur_ord_hdr(M193_I2_so_type)
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
    lgBlnFlgChgValue = Request("hdnchgValue")
   
    '-----------------------
    'Data manipulate area: Po header
    '-----------------------
		'첫번째 탭 
	    If lgIntFlgMode = OPMD_CMODE And Trim(Request("txtPoNo2")) <> "" Then			'신규가 아닐때 
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
		I5_m_pur_ord_hdr(M193_I2_xch_rate_op)			= Trim(Request("hdnxchrateop"))

	
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
		if Ucase(Trim(Request("hdnImportflg"))) = "Y" then
    
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
    
    'Hidden Field
  	'I5_m_pur_ord_hdr(M193_I2_import_flg)			= UCase(Trim(Request("hdnImportflg")))
  	'I5_m_pur_ord_hdr(M193_I2_bl_flg)				= UCase(Trim(Request("hdnBlflg")))
  	'I5_m_pur_ord_hdr(M193_I2_cc_flg)				= UCase(Trim(Request("hdnCcflg")))
  	'I5_m_pur_ord_hdr(M193_I2_rcpt_flg)				= UCase(Trim(Request("hdnRcptflg")))
  	'I5_m_pur_ord_hdr(M193_I2_subcontra_flg)		= UCase(Trim(Request("hdnSubcontraflg")))
  	'I5_m_pur_ord_hdr(M193_I2_ret_flg)				= UCase(Trim(Request("hdnRetflg")))
  	'I5_m_pur_ord_hdr(M193_I2_iv_flg)				= UCase(Trim(Request("hdnIvflg")))
  	'I5_m_pur_ord_hdr(M193_I2_rcpt_type)			= UCase(Trim(Request("hdnRcptType")))
  	'I5_m_pur_ord_hdr(M193_I2_issue_type)			= UCase(Trim(Request("hdnIssueType")))
  	'I5_m_pur_ord_hdr(M193_I2_iv_type)				= UCase(Trim(Request("hdnIvType")))
  	
  		I5_m_pur_ord_hdr(M193_I2_release_flg)			= "N"

  	'I5_m_pur_ord_hdr(M193_I2_cls_flg)				= "N"
  	'I5_m_pur_ord_hdr()								= "N"
  	'I5_m_pur_ord_hdr(M193_I2_charge_flg)			= "N"
    'M31111.ImpMPurOrdHdrClsFlg						= "N"
    'M31111.ImpMPurOrdHdrPoPrcFlg					= "F"
    'M31111.ImpMPurOrdHdrChargeFlg					= "N"

		
		'+++++++++++++++++++++++++++++++++
		'Data manipulate area: Po detail
		'+++++++++++++++++++++++++++++++
	
	lgSSCheckValue = Request("hdnSSCheckValue")
	
	 itxtSpread = ""
             
   	 iCUCount = Request.Form("txtCUSpread").Count
   	 iDCount  = Request.Form("txtDSpread").Count
             
    	itxtSpreadArrCount = -1
	ReDim itxtSpreadArr(iCUCount + iDCount)
	if lgSSCheckValue then		
		

		   
	
		  For ii = 1 To iDCount
        		itxtSpreadArrCount = itxtSpreadArrCount + 1
       			itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    		Next
    		For ii = 1 To iCUCount
        		itxtSpreadArrCount = itxtSpreadArrCount + 1
        		itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
   		 Next
    		itxtSpread = Join(itxtSpreadArr,"")
	End if
	
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
	
	LngMaxRow = Request("txtMaxRows")
	'=================================
	If lgIntFlgMode = OPMD_CMODE Then
			iStrCommandSent 							= "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
			I5_m_pur_ord_hdr(M193_I2_po_no)				= Ucase(Trim(Request("txtPoNo")))
			iStrCommandSent 							= "UPDATE"
	End If

	Set M31111 = Server.CreateObject("PM3GC11.cMMaintPurOrdCombi")    'Po Header
		If CheckSYSTEMError(Err,True) = true Then
			Set M31111 = Nothing 		
			Exit Sub
	End If	
		

	E1_m_pur_ord_hdr_po_no = M31111.M_MAINT_PUR_ORD_COMBI_SVR("F",gStrGlobalCollection, _
												  iStrCommandSent, _
												  I1_b_company, _
												  I2_m_config_process, _
												  I3_b_biz_partner, _
												  I4_b_pur_grp, _
												  I5_m_pur_ord_hdr, _
												  LngMaxRow, _
												  itxtSpread, _
												  ,_
												  iErrorPosition)
		
	ls_msg = Trim(Cstr(Err.Description))
	
	Set M31111 = Nothing
	
	If Len(ls_msg) >= 6 Then
		If Right(ls_msg,6) = "TRGERR" Then
			Call DisplayMsgBox(Mid(ls_msg,Len(ls_msg) -11,6), "", "", "", I_MKSCRIPT)
			Exit Sub
		Else
			If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
				Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
				Exit Sub
			End If
		End If
	Else
		If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
				Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
				Exit Sub
		End If
	End If	
	
	
'	If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
'			Set M31111 = Nothing
'			Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
'			Exit Sub
'	End If
		
	Response.Write "<Script Language=vbscript>"  								& vbCr
	Response.Write "With parent"	  		
		
	If lgIntFlgMode = OPMD_CMODE  Then 
		Response.Write "	.frm1.txtPoNo.Value	= """ & ConvSPChars(E1_m_pur_ord_hdr_po_no) & """" & vbCr
		Response.Write "	.frm1.txtPoNo2.Value	= """ & ConvSPChars(E1_m_pur_ord_hdr_po_no) & """" & vbCr
	End If
	Response.Write "	.DbSaveOk" 	& vbCr
	Response.Write "End With" 		& vbCr
	Response.Write "</Script>"		& vbCr
				
	'Set M31111 = Nothing						

End Sub		
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
				Response.Write "<Script Language=vbscript>" 					& vbCr
				Response.Write "With parent"									& vbCr	
				Response.Write ".frm1.btnCfm.disabled=false" & vbCr
				Response.Write "End With"   & vbCr
				Response.Write "</Script>" & vbCr
				Response.End							
			 End If
		End If
	Else		
		If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		    Response.Write "<Script Language=vbscript>" 					& vbCr
			Response.Write "With parent"									& vbCr	
			Response.Write ".frm1.btnCfm.disabled=false" & vbCr
			Response.Write "End With"   & vbCr
			Response.Write "</Script>" & vbCr
			Response.End							
		End If
	End If	
	
'	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
'		Set M31211 = Nothing												'☜: ComProxy Unload
'			Response.Write "<Script Language=VBScript>" & vbCr
'			Response.Write "parent.frm1.btnCfm.disabled = False" & vbCr
'			Response.Write "</Script>"  & vbCr
'		Exit Sub
'	 End If

    'Set M31211 = Nothing                                                   '☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" 					& vbCr
	Response.Write "With parent"									& vbCr	
    Response.Write "  If """ & strMode & """ = ""Release""  Then " 	& vbCr
	Response.Write "    .frm1.rdoRelease(0).Checked = true"        	& vbCr
	Response.Write"   Else "    									& vbCr
	Response.Write "    .frm1.rdoRelease(1).Checked = true"        	& vbCr
	Response.Write "  End if "  & vbCr
	Response.Write ".DbSaveOk"  & vbCr
	Response.Write "End With"   & vbCr
	Response.Write "</Script>"  & vbCr
    
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
	
	Dim M31111,iErrorPosition
	Dim iMaxRow, istrVal
	Dim I5_m_pur_ord_hdr
	Const M193_I2_po_no = 0										

	Redim I5_m_pur_ord_hdr(76)

	I5_m_pur_ord_hdr(M193_I2_po_no)				= Trim(Request("txtPoNo"))
	iMaxRow										= Trim(Request("txtMaxRows"))
	istrVal										= Trim(Request("txtSpread"))
	

   ' Set M31111 = Server.CreateObject("PM3G111.cMMaintPurOrdHdrS")
    Set M31111 = Server.CreateObject("PM3GC11.cMMaintPurOrdCombi")        
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set M31111 = Nothing 		
			Exit Sub
	End If
					   
	Call M31111.M_MAINT_PUR_ORD_COMBI_SVR("F",gStrGlobalCollection, _
									  "DELETE", _
									  "", _
									  "", _
									  "", _
									  "", _
									  I5_m_pur_ord_hdr, _
									  iMaxRow, _
									  istrVal, _
									  iErrorPosition)
	
	ls_msg = Trim(Cstr(Err.Description))
									  
	Set M31111 = Nothing
	
	If Len(ls_msg) >= 6 Then
		If Right(ls_msg,6) = "TRGERR" Then
			Call DisplayMsgBox(Mid(ls_msg,Len(ls_msg) -11,6), "", "", "", I_MKSCRIPT)
			Exit Sub
		Else
			If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
				Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
				Exit Sub
			End If
		End If
	Else
		If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
				Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
				Exit Sub
		End If
	End If
	

'	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
'		Set M31111 = Nothing												'☜: ComProxy Unload
'		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
'	 End If

    'Set M31111 = Nothing															'☜: Unload Comproxy

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
		Response.Write ".txtPotypeCd.value = """ &"""" & vbCr
		Response.Write "end with"
		Response.Write"</Script>"

		Exit Sub
    End If
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with Parent.frm1" & vbCr
	Response.Write "Dim lgTab" & vbCr

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

	if ConvSPChars(Ucase(Trim(E1_m_config_process(M369_E1_import_flg)))) = "Y" then
	    Response.Write "Parent.ggoOper.SetReqAttr	.txtDvryDt, ""N"""	& vbCr	
	    Response.Write "Parent.ggoOper.SetReqAttr	.txtOffDt, ""N"""	& vbCr
		Response.Write "Parent.ggoOper.SetReqAttr	.txtApplicantCd, ""N"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtApplicantNm, ""Q"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtIncotermsCd, ""N"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtIncotermsNm, ""Q"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtTransCd, ""N"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtTransNm, ""Q"""	& vbCr
	else     
		Response.Write "Parent.ggoOper.SetReqAttr	.txtDvryDt, ""D"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtOffDt, ""Q"""	& vbCr
		Response.Write "Parent.ggoOper.SetReqAttr	.txtApplicantCd, ""Q"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtApplicantNm, ""Q"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtIncotermsCd, ""Q"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtIncotermsNm, ""Q"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtTransCd, ""Q"""	& vbCr	
		Response.Write "Parent.ggoOper.SetReqAttr	.txtTransNm, ""Q"""	& vbCr
	end if	
	
	If UCase(trim(Request("txtTabClickFlag"))) = "TRUE" Then
		Response.Write "	parent.lgOpenFlag	= True"	& vbCr
		Response.Write "	Call parent.ClickTab"&Request("txtgSelframeFlg")&"()" 	& vbCr
	End If
	
	Response.Write "end with" & vbCr	
	Response.Write "</Script>" & vbCr
	
	Set M14119 = Nothing
	
End Sub
'============================================================================================================
' Name : SublookupPrice
' Desc : 
'============================================================================================================
Sub SublookupPrice()
	
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
	
	
	Dim iPM3G1P9
	
	Dim I1_m_supplier_item_price 
	Dim I2_b_biz_partner_bp_cd 
	Dim I3_b_item_item_cd 
	Dim I4_b_plant_plant_cd 
	Dim E1_m_supplier_item_price 
	Dim E2_b_item 
	Dim E3_b_plant 
	Dim E4_b_storage_location
	Dim E5_b_hs_code 
	Dim E6_m_supplier_item_by_plant 
	Dim E7_b_minor 
	
	Const M106_I1_pur_unit = 0    '  View Name : imp m_supplier_item_price
    Const M106_I1_pur_cur = 1
    Const M106_I1_valid_fr_dt = 2
    ReDim I1_m_supplier_item_price(M106_I1_valid_fr_dt)

    Const M106_E1_pur_prc = 0    '  View Name : exp m_supplier_item_price

    Const M106_E2_item_cd = 0    '  View Name : exp b_item
    Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 3
    Const M106_E2_vat_rate = 4
    ReDim E2_b_item(M106_E2_vat_rate)

    Const M106_E3_plant_cd = 0    '  View Name : exp b_plant
    Const M106_E3_plant_nm = 1
    ReDim E3_b_plant(M106_E3_plant_nm)
    
    Const M106_E4_sl_cd = 0    '  View Name : exp b_storage_location
    Const M106_E4_sl_nm = 1
    ReDim E4_b_storage_location(M106_E4_sl_nm)

    Const M106_E5_hs_cd = 0    '  View Name : exp b_hs_code
    Const M106_E5_hs_nm = 1
    ReDim E5_b_hs_code(M106_E5_hs_nm)

    Const M106_E6_pur_priority = 0    '  View Name : exp m_supplier_item_by_plant
    Const M106_E6_pur_unit = 1
    Const M106_E6_sppl_item_cd = 2
    Const M106_E6_sppl_item_nm = 3
    Const M106_E6_sppl_item_spec = 4
    Const M106_E6_maker_nm = 5
    Const M106_E6_valid_fr_dt = 6
    Const M106_E6_valid_to_dt = 7
    Const M106_E6_usage_flg = 8
    Const M106_E6_sppl_sales_prsn = 9
    Const M106_E6_sppl_tel_no = 10
    Const M106_E6_sppl_dlvy_lt = 11
    Const M106_E6_under_tol = 12
    Const M106_E6_over_tol = 13
    Const M106_E6_min_qty = 14
    Const M106_E6_def_flg = 15
    Const M106_E6_max_qty = 16
    ReDim E6_m_supplier_item_by_plant(M106_E6_max_qty)

    Const M106_E7_minor_nm = 0    '  View Name : exp_vat_nm b_minor
    Const M106_E7_minor_cd = 1
    ReDim E7_b_minor(M106_E7_minor_cd)

	

	
    Set iPM3G1P9 = Server.CreateObject("PM3G1P9.cMLookupPriceS")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1P9 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	I1_m_supplier_item_price(M106_I1_valid_fr_dt)	= UNIConvDate(Trim(Request("txtStampDt")))
	I2_b_biz_partner_bp_cd							= Trim(Request("txtBpCd"))
	I3_b_item_item_cd								= Trim(Request("txtItemCd"))
	I4_b_plant_plant_cd								= Trim(Request("txtPlantCd"))
	I1_m_supplier_item_price(M106_I1_pur_unit)		= Trim(Request("txtUnit"))
	I1_m_supplier_item_price(M106_I1_pur_cur)		= Trim(Request("txtCurrency"))
	
	Call iPM3G1P9.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, I1_m_supplier_item_price, I2_b_biz_partner_bp_cd, _
									I3_b_item_item_cd, I4_b_plant_plant_cd, iPriceType , E1_m_supplier_item_price, _
									E2_b_item, E3_b_plant, E4_b_storage_location, E5_b_hs_code, _
									E6_m_supplier_item_by_plant, E7_b_minor)
    

	
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1P9 = Nothing												'☜: ComProxy Unload
		'Response.Write "<Script Language=vbscript>" & vbCr		
        'Response.Write "Dim PoPrice1              " & vbCr
        'Response.Write "parent.frm1.vspdData.Row  = """ & ConvSPChars(Trim(Request("txtRow"))) & """" & vbCr 
       ' Response.Write "parent.frm1.vspdData.Col  = parent.C_Cost " & vbCr
       ' Response.Write "PoPrice1 = Parent.frm1.vspdData.Text " & vbCr
        'Response.Write "Parent.frm1.vspdData.Col  = Parent.C_Cost " & vbCr
		'Response.Write "Parent.frm1.vspdData.Text = PoPrice1 " & vbCr
        'Response.Write "Parent.vspdData_Change Parent.C_PoPrice2 , """ & Trim(Request("txtRow")) & """" & vbCr
        'Response.Write "</Script> " & vbCr
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With Parent" & vbCr
	Response.Write " Dim strRow " & vbCr
	Response.Write "	strRow	= """ & ConvSPChars(Trim(Request("txtRow"))) & """" & vbCr
	Response.Write "	.frm1.vspdData.Row  = """ & ConvSPChars(Trim(Request("txtRow"))) & """" & vbCr
	Response.Write "	.frm1.vspdData.Col  = Parent.C_Cost " & vbCr
	Response.Write "	.frm1.vspdData.Text = """ & UNINumClientFormat(E1_m_supplier_item_price(M106_E1_pur_prc) ,ggUnitCost.DecPoint,0)   & """" & vbCr
	'Response.Write "Parent.frm1.vspdData.Col  = Parent.C_CostConCd" & vbCr
	'Response.Write "Parent.frm1.vspdData.value = """"F""""  & vbCr
	'Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr		
	'Response.Write "	.ggoSpread.SpreadLock		C_CostCon,  """ & Trim(Request("txtRow")) & """,C_CostCon,""" & Trim(Request("txtRow")) & """" & vbCr
	'Response.Write "	.ggoSpread.SSSetProtected	C_CostCon,	""" & Trim(Request("txtRow")) & """,""" & Trim(Request("txtRow")) & """" & vbCr
    Response.Write "Parent.vspdData_Change Parent.C_Cost , """ & ConvSPChars(Trim(Request("txtRow"))) & """" & vbCr
    Response.Write "End With " & vbCr
    Response.Write "</Script>"                  & vbCr

	Set iPM3G1P9 = Nothing

End Sub

'============================================================================================================
' Name : lookupPriceForSelection
' Desc :
'============================================================================================================
Sub lookupPriceForSelection()

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear          

    Dim iPM3G1P9
    Dim iLngMaxRow
    Dim iLngRow
  
    Dim I1_m_supplier_item_price 
	Dim I2_b_biz_partner_bp_cd 
	Dim I3_b_item_item_cd 
	Dim I4_b_plant_plant_cd 
	Dim E1_m_supplier_item_price 
	Dim E2_b_item 
	Dim E3_b_plant 
	Dim E4_b_storage_location
	Dim E5_b_hs_code 
	Dim E6_m_supplier_item_by_plant 
	Dim E7_b_minor 
	
	Const M106_I1_pur_unit = 0    '  View Name : imp m_supplier_item_price
    Const M106_I1_pur_cur = 1
    Const M106_I1_valid_fr_dt = 2
    ReDim I1_m_supplier_item_price(M106_I1_valid_fr_dt)

    Const M106_E1_pur_prc = 0    '  View Name : exp m_supplier_item_price

    Const M106_E2_item_cd = 0    '  View Name : exp b_item
    Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 3
    Const M106_E2_vat_rate = 4
    ReDim E2_b_item(M106_E2_vat_rate)

    Const M106_E3_plant_cd = 0    '  View Name : exp b_plant
    Const M106_E3_plant_nm = 1
    ReDim E3_b_plant(M106_E3_plant_nm)
    
    Const M106_E4_sl_cd = 0    '  View Name : exp b_storage_location
    Const M106_E4_sl_nm = 1
    ReDim E4_b_storage_location(M106_E4_sl_nm)

    Const M106_E5_hs_cd = 0    '  View Name : exp b_hs_code
    Const M106_E5_hs_nm = 1
    ReDim E5_b_hs_code(M106_E5_hs_nm)

    Const M106_E6_pur_priority = 0    '  View Name : exp m_supplier_item_by_plant
    Const M106_E6_pur_unit = 1
    Const M106_E6_sppl_item_cd = 2
    Const M106_E6_sppl_item_nm = 3
    Const M106_E6_sppl_item_spec = 4
    Const M106_E6_maker_nm = 5
    Const M106_E6_valid_fr_dt = 6
    Const M106_E6_valid_to_dt = 7
    Const M106_E6_usage_flg = 8
    Const M106_E6_sppl_sales_prsn = 9
    Const M106_E6_sppl_tel_no = 10
    Const M106_E6_sppl_dlvy_lt = 11
    Const M106_E6_under_tol = 12
    Const M106_E6_over_tol = 13
    Const M106_E6_min_qty = 14
    Const M106_E6_def_flg = 15
    Const M106_E6_max_qty = 16
    ReDim E6_m_supplier_item_by_plant(M106_E6_max_qty)

    Const M106_E7_minor_nm = 0    '  View Name : exp_vat_nm b_minor
    Const M106_E7_minor_cd = 1
    ReDim E7_b_minor(M106_E7_minor_cd)
    
    If Len(Trim(Request("txtStampDt"))) Then
		If UNIConvDate(Request("txtStampDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtStampDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If

    Set iPM3G1P9 = Server.CreateObject("PM3G1P9.cMLookupPriceS")      
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
   If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1P9 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if

	
	iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

	Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus	,LngRow																'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																		'☜: Group Count

	arrTemp = Split(Request("txtSpread"), gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 
    '-----------------------
    'Data manipulate area
    '-----------------------
    
    lGrpCnt = 0
    
	ReDim returnValue(iLngMaxRow)

    For iLngRow = 1 To iLngMaxRow
    
		lGrpCnt = lGrpCnt +1														'☜: Group Count
		 
		arrVal = Split(arrTemp(iLngRow-1), gColSep)
	
		I1_m_supplier_item_price(M106_I1_valid_fr_dt) = UNIConvDate(Trim(Request("hdnPoDt")))
		
		I2_b_biz_partner_bp_cd						  = Trim(arrVal(0))
		I3_b_item_item_cd							  = Trim(arrVal(1))
		I4_b_plant_plant_cd							  = Trim(arrVal(2))
		I1_m_supplier_item_price(M106_I1_pur_unit)	  = Trim(arrVal(3))
		I1_m_supplier_item_price(M106_I1_pur_cur)	  = Trim(arrVal(4))
		 

		Call iPM3G1P9.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, I1_m_supplier_item_price, I2_b_biz_partner_bp_cd, _
									I3_b_item_item_cd, I4_b_plant_plant_cd, iPriceType , E1_m_supplier_item_price, _
									E2_b_item, E3_b_plant, E4_b_storage_location, E5_b_hs_code, _
									E6_m_supplier_item_by_plant, E7_b_minor)
		returnValue(iLngRow) = UNINumClientFormat(E1_m_supplier_item_price(M106_E1_pur_prc) ,ggUnitCost.DecPoint,0)
		 
		lGrpCnt = 0   
			
    Next


	Set iPM3G1P9 = Nothing

	Dim rowindex, rowCount, resultindex

	resultindex = 1  

	rowCount = Request("txtMaxRows")
	arrTemp = Split(Request("txtSpread"), gRowSep)
	iLngRow = 1
	
	For rowindex = 1 To rowCount

		arrVal = Split(arrTemp(iLngRow-1), gColSep)
		If iLngRow <= iLngMaxRow Then
			If CInt(rowindex) = Cint(arrVal(5)) Then
			
	Response.Write "<script language=vbscript>" & vbCr
	Response.Write " Parent.frm1.vspdData.Row  = """ & ConvSPChars(rowindex) & """" & vbCr
	Response.Write " Parent.frm1.vspdData.Col  = Parent.C_Cost "   & vbCr
	'Response.Write " Parent.frm1.vspdData.Text = """ & UNINumClientFormat(returnValue(resultindex),ggUnitCost.DecPoint,0 ) & """" & vbCr
	Response.Write " Parent.frm1.vspdData.Text = """ & ConvSPChars(returnValue(resultindex)) & """" & vbCr
	Response.Write " Parent.vspdData_Change Parent.C_Cost , """ & Cint(arrVal(5)) & """" & vbCr
	'Response.Write " Parent.frm1.vspdData.Col  = Parent.C_Cost "   & vbCr
	'Response.Write " Parent.frm1.vspdData.Text = """ & ConvSPChars(returnValue(resultindex)) & """" & vbCr
	'Response.Write " Parent.frm1.vspdData.Col  = Parent.C_Cost "   & vbCr
	'Response.Write " Parent.frm1.vspdData.Text = """ & ConvSPChars(returnValue(resultindex)) & """" & vbCr
	Response.Write "</script>" & vbCr

				resultindex = resultindex + 1
				iLngRow = iLngRow + 1
			End if
		End if
	Next	
	
' === 2005.07.15 단가 일괄 불러오기 관련 수정 ===========================================	
		Response.Write "<script language=vbscript>" & vbCr	
		Response.Write " Call parent.btnCallPrice_Ok() " & vbCr
		Response.Write "</script>" & vbCr	
' === 2005.07.15 단가 일괄 불러오기 관련 수정 ===========================================		
			
	Set iPM3G1P9 = Nothing

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
				                         E1_b_biz_partner, "Y")
	
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
		Response.Write "End If"& vbCr
			
		Response.Write "If Trim(.txtCurr.Value) = """" Then " & vbCr 	
		Response.Write " .txtCurr.Value	= """& ConvSPChars(Trim(E1_b_biz_partner(S074_E1_currency))) & """" & vbCr
		Response.Write "End if"& vbCr
	
		Response.Write "If UCase(Parent.Parent.gCurrency)=" & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_currency))) & " Then " & vbCr 
		Response.Write "Call Parent.ggoOper.SetReqAttr(frm1.cboXchop,""""Q"""")" & vbCr
		Response.Write "End if"& vbCr

		Response.Write "If Trim(.txtVatType.Value) = """" Then " & vbCr 	
		Response.Write " .txtVatType.Value	= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type))) & """" & vbCr
		Response.Write " .txtVatTypeNm.Value = """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type_nm))) & """" & vbCr
		Response.Write " .txtVatrt.text	= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write "End if"& vbCr
			
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
		Response.Write "End If"& vbCr
			
		Response.Write "End With " & vbCr	
		Response.Write "Parent.SetVatTypeHdr() "	& vbCr													'부가세율을 현재 시점의 세율로 가져옴 
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
	Const S074_E1_own_rgst_dt = 120            
	Const S074_E1_ind_type_nm = 121             
	Const S074_E1_ind_class_nm = 122            
	Const S074_E1_bp_group_nm = 123             
	Const S074_E1_b_country_nm = 124            
	Const S074_E1_b_province_nm = 125           
	Const S074_E1_trans_meth_nm = 126           
	Const S074_E1_deal_type_nm = 127            
	Const S074_E1_bp_grade_nm = 128             
	Const S074_E1_s_credit_limit = 129          
	Const S074_E1_b_sales_grp_nm = 130          
	Const S074_E1_b_to_grp_nm = 131             
	Const S074_E1_b_pur_grp_nm = 132            
	Const S074_E1_vat_type_nm = 133       
	Const S074_E1_pay_meth_nm = 134       
	Const S074_E1_pay_type_nm = 135       
	Const S074_E1_tax_area_nm = 136       
	Const S074_E1_b_zip_code = 137        
	Const S074_E1_b_pur_org = 138         
	Const S074_E1_b_pur_org_nm = 139      
	Const S074_E1_vat_inc_flag_nm = 140   
	Const S074_E1_card_co_cd_nm = 141     
	Const S074_E1_pay_meth_pur_nm = 142   
	Const S074_E1_pay_type_pur_nm = 143   
	Const S074_E1_bank_cd_nm = 144        

	Dim iSupplierCd,iStrCurrency
	Dim B1H019
		
    Set B1H019 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
			Set B1H019 = Nothing 		
			Exit Sub
	End If
	iStrCurrency= Trim(Request("txtCurr"))
	iSupplierCd = Trim(Request("txtSupplierCd"))
	
	Call B1H019.B_LOOKUP_BIZ_PARTNER(gStrGlobalCollection, _
				                         iSupplierCd, _
				                         E1_b_biz_partner, "Y")
	
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
	    Response.Write ".hdnSupplierCd.Value		= """ & ConvSPChars(Trim(Request("txtSupplierCd"))) & """" & vbCr
		Response.Write ".txtSupplierNm.Value		= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_bp_nm))) & """" & vbCr
		if ConvSPChars(Trim(E1_b_biz_partner(S074_E1_currency))) <> "" then      '화폐단위가 없을때 법인화폐정보 
            Response.Write ".txtCurr.Value = """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_currency))) & """" & vbCr
        Else
           Response.Write ".txtCurr.Value = .hdnCurr.value " & vbCr 
        End If 
		Response.Write ".txtCurrNm.Value			= """" " & vbCr
		Response.Write ".txtVatType.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type))) & """" & vbCr
		Response.Write ".txtVatTypeNm.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type_nm))) & """" & vbCr
		Response.Write ".txtVatrt.text				= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
	'	Response.Write ".hdntxtVatType.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type))) & """" & vbCr
	'	Response.Write ".hdntxtVatTypeNm.Value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_type_nm))) & """" & vbCr
	'	Response.Write ".hdntxtVatrt.value				= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
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
		'Response.Write ".txtGroupCd.value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pur_grp))) & """" & vbCr
		'Response.Write ".txtGroupNm.value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_b_pur_grp_nm))) & """" & vbCr
		
		if Trim(Request("txtGroupCd")) = "" then
			Response.Write ".txtGroupCd.value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_pur_grp))) & """" & vbCr
			Response.Write ".txtGroupNm.value			= """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_b_pur_grp_nm))) & """" & vbCr
		end if
    Response.Write ".txtTel.value			    = """ & ConvSPChars(Trim(E1_b_biz_partner(S074_E1_bp_contact_pt))) & """" & vbCr
        
    if ConvSPChars(Trim(E1_b_biz_partner(S074_E1_vat_inc_flag))) = "2" then
	    Response.Write " .rdoVatFlg2.Checked= true "  & vbCr
			Response.Write " .hdvatFlg.value 	= ""2"" " & vbCr
    Else
			Response.Write " .rdoVatFlg1.Checked= true"   & vbCr 
			Response.Write " .hdvatFlg.value 	= ""0"" " & vbCr
    End if
  		
		Response.Write "End With" & vbCr	
		Response.Write "Parent.ChangeCurr()" & vbCr
		Response.Write "Parent.SetVatTypeHdr()" & vbCr								'부가세율을 현재 시점의 세율로 가져옴 
		'Response.Write "Parent.SetVatTypeHdr()" & vbCr								'부가세율을 현재 시점의 세율로 가져옴 
		Response.Write "</Script>" & vbCr   				
	
	Set B1H019 = Nothing
	
End Sub
'============================================================================================================
' Name : SubLookUpItemPlant
' Desc : Look Up Item By Plant 
'============================================================================================================
Sub SubLookUpItemPlant()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim M31429
  	Dim SlCd,SlNm,Unit,trackingFlg
	Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																		'☜: Group Count
	
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow

	Dim I1_m_supplier_item_price
    Const M106_I1_pur_unit = 0    'View Name : imp m_supplier_item_price
    Const M106_I1_pur_cur = 1
    Const M106_I1_valid_fr_dt = 2

	Dim E1_m_supplier_item_price_pur_prc
	Dim E2_b_item
    'Const M106_E2_item_cd = 0    'View Name : exp b_item
    'Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 3
    Const M106_E2_vat_rate = 4
	Dim E3_b_plant
    Const M106_E3_plant_cd = 0    'View Name : exp b_plant
    Const M106_E3_plant_nm = 1
	Dim E4_b_storage_location
    Const M106_E4_sl_cd = 0    'View Name : exp b_storage_location
    Const M106_E4_sl_nm = 1
	Dim E5_b_hs_code
    Const M106_E5_hs_cd = 0    'View Name : exp b_hs_code
    Const M106_E5_hs_nm = 1
	Dim E6_m_supplier_item_by_plant
    Const M106_E6_pur_priority = 0    'View Name : exp m_supplier_item_by_plant
    Const M106_E6_pur_unit = 1
    Const M106_E6_sppl_item_cd = 2
    Const M106_E6_sppl_item_nm = 3
    Const M106_E6_sppl_item_spec = 4
    Const M106_E6_maker_nm = 5
    Const M106_E6_valid_fr_dt = 6
    Const M106_E6_valid_to_dt = 7
    Const M106_E6_usage_flg = 8
    Const M106_E6_sppl_sales_prsn = 9
    Const M106_E6_sppl_tel_no = 10
    Const M106_E6_sppl_dlvy_lt = 11
    Const M106_E6_under_tol = 12
    Const M106_E6_over_tol = 13
    Const M106_E6_min_qty = 14
    Const M106_E6_def_flg = 15
    Const M106_E6_max_qty = 16
	Dim E7_b_minor_vat
    Const M106_E7_minor_nm = 1    'View Name : exp_vat_nm b_minor
    Const M106_E7_minor_cd = 0

    Dim E1_b_pur_org
    Const P003_E1_pur_org = 0
    Const P003_E1_pur_org_nm = 1
    Const P003_E1_valid_fr_dt = 2
    Const P003_E1_valid_to_dt = 3
    Const P003_E1_usage_flg = 4

    Dim E2_b_item_group
    Const P003_E2_item_group_cd = 0
    Const P003_E2_item_group_nm = 1
    Const P003_E2_leaf_flg = 2

    Dim E3_for_issued_b_storage_location
    Const P003_E3_sl_cd = 0
    Const P003_E3_sl_type = 1
    Const P003_E3_sl_nm = 2

    Dim E4_for_major_b_storage_location
    Const P003_E4_sl_cd = 0
    Const P003_E4_sl_type = 1
    Const P003_E4_sl_nm = 2

    Dim E5_i_material_valuation
    Const P003_E5_prc_ctrl_indctr = 0
    Const P003_E5_moving_avg_prc = 1
    Const P003_E5_std_prc = 2
    Const P003_E5_prev_std_prc = 3

    Dim E6_b_item_by_plant
    Const P003_E6_procur_type = 0
    Const P003_E6_order_unit_mfg = 1
    Const P003_E6_order_lt_mfg = 2
    Const P003_E6_order_lt_pur = 3
    Const P003_E6_order_type = 4
    Const P003_E6_order_rule = 5
    Const P003_E6_req_round_flg = 6
    Const P003_E6_fixed_mrp_qty = 7
    Const P003_E6_min_mrp_qty = 8
    Const P003_E6_max_mrp_qty = 9
    Const P003_E6_round_qty = 10
    Const P003_E6_round_perd = 11
    Const P003_E6_scrap_rate_mfg = 12
    Const P003_E6_ss_qty = 13
    Const P003_E6_prod_env = 14
    Const P003_E6_mps_flg = 15
    Const P003_E6_issue_mthd = 16
    Const P003_E6_mrp_mgr = 17
    Const P003_E6_inv_check_flg = 18
    Const P003_E6_lot_flg = 19
    Const P003_E6_cycle_cnt_perd = 20
    Const P003_E6_inv_mgr = 21
    Const P003_E6_major_sl_cd = 22
    Const P003_E6_abc_flg = 23
    Const P003_E6_mps_mgr = 24
    Const P003_E6_recv_inspec_flg = 25
    Const P003_E6_inspec_lt_mfg = 26
    Const P003_E6_inspec_mgr = 27
    Const P003_E6_valid_from_dt = 28
    Const P003_E6_valid_to_dt = 29
    Const P003_E6_item_acct = 30
    Const P003_E6_single_rout_flg = 31
    Const P003_E6_prod_mgr = 32
    Const P003_E6_issued_sl_cd = 33
    Const P003_E6_issued_unit = 34
    Const P003_E6_order_unit_pur = 35
    Const P003_E6_var_lt = 36
    Const P003_E6_scrap_rate_pur = 37
    Const P003_E6_pur_org = 38
    Const P003_E6_prod_inspec_flg = 39
    Const P003_E6_final_inspec_flg = 40
    Const P003_E6_ship_inspec_flg = 41
    Const P003_E6_inspec_lt_pur = 42
    Const P003_E6_option_flg = 43
    Const P003_E6_over_rcpt_flg = 44
    Const P003_E6_over_rcpt_rate = 45
    Const P003_E6_damper_flg = 46
    Const P003_E6_damper_max = 47
    Const P003_E6_damper_min = 48
    Const P003_E6_reorder_pnt = 49
    Const P003_E6_std_time = 50
    Const P003_E6_add_sel_rule = 51
    Const P003_E6_add_sel_value = 52
    Const P003_E6_add_seq_rule = 53
    Const P003_E6_add_seq_atrid = 54
    Const P003_E6_rem_sel_rule = 55
    Const P003_E6_rem_sel_value = 56
    Const P003_E6_rem_seq_rule = 57
    Const P003_E6_rem_seq_atrid = 58
    Const P003_E6_llc = 59
    Const P003_E6_tracking_flg = 60
    Const P003_E6_valid_flg = 61
    Const P003_E6_work_center = 62
    Const P003_E6_order_from = 63
    Const P003_E6_cal_type = 64
    Const P003_E6_line_no = 65
    Const P003_E6_atp_lt = 66
    Const P003_E6_etc_flg1 = 67
    Const P003_E6_etc_flg2 = 68

    Dim E7_b_item
    Const P003_E7_item_cd = 0
    Const P003_E7_item_nm = 1
    Const P003_E7_formal_nm = 2
    Const P003_E7_spec = 3
    Const P003_E7_item_acct = 4
    Const P003_E7_item_class = 5
    Const P003_E7_hs_cd = 6
    Const P003_E7_hs_unit = 7
    Const P003_E7_unit_weight = 8
    Const P003_E7_unit_of_weight = 9
    Const P003_E7_basic_unit = 10
    Const P003_E7_draw_no = 11
    Const P003_E7_item_image_flg = 12
    Const P003_E7_phantom_flg = 13
    Const P003_E7_blanket_pur_flg = 14
    Const P003_E7_base_item_cd = 15
    Const P003_E7_proportion_rate = 16
    Const P003_E7_valid_flg = 17
    Const P003_E7_valid_from_dt = 18
    Const P003_E7_valid_to_dt = 19

    Dim E8_b_plant
    Const P003_E8_plant_cd = 0
    Const P003_E8_plant_nm = 1

	Dim LngMaxRow
	Dim LngRow
	Dim iStrSupplierCd
    Dim iStrSlCd

	'이성룡 추가  단가type 의 유무를 조사 
	If CheckPriceType(iPriceType) = False then
        Call DisplayMsgBox("171214", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Exit Sub
    End If	
	
	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	
	arrTemp = Split(Request("txtSpread"), gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 
    '-----------------------
    'Data manipulate area
    '-----------------------
    
    lGrpCnt = 0
 
 	Redim I1_m_supplier_item_price(2)

    Set M31429 = Server.CreateObject("PM3G1P9.cMLookupPriceS")

    If CheckSYSTEMError(Err,True) = true then
		Set M31429 = Nothing 		
		Exit Sub
	End If

	Dim B1b119
	Set	B1b119 = Server.CreateObject("PB3S106.cBLkUpItemByPlt")
	
	If CheckSYSTEMError(Err,True) = true Then
		Set B1b119 = Nothing 		
		Exit Sub
	End If

	I1_m_supplier_item_price(M106_I1_pur_cur)		= Trim(Request("txtCurr"))
	I1_m_supplier_item_price(M106_I1_valid_fr_dt)	= UNIConvDate(Request("txtPoDt"))
    iStrSupplierCd = Trim(Request("txtSupplierCd"))
    
    For LngRow = 1 To LngMaxRow  

		lGrpCnt = lGrpCnt +1	
		
		arrVal = Split(arrTemp(LngRow-1), gColSep)

		I1_m_supplier_item_price(M106_I1_pur_unit)		= Trim(arrVal(4))
		iStrSlCd										= Trim(arrVal(3))
		Call M31429.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, _
									   I1_m_supplier_item_price, _
									   iStrSupplierCd, _
									   Trim(arrVal(1)), _
									   Trim(arrVal(2)), _
									   iPriceType , _
										E1_m_supplier_item_price_pur_prc, _
										E2_b_item, _
										E3_b_plant, _
										E4_b_storage_location, _
										E5_b_hs_code, _
										E6_m_supplier_item_by_plant, _
										E7_b_minor_vat)
		Err.Clear
		If CheckSYSTEMError(Err,True) = true Then
			Set M31429 = Nothing 		
			Set B1b119 = Nothing 		
			Exit Sub
		End If
		
		If Trim(iStrSlCd) = "" Then
			SlCd = ConvSPChars(E4_b_storage_location(M106_E4_sl_cd))
			SlNm = ConvSPChars(E4_b_storage_location(M106_E4_sl_nm))
		Else
			SlCd = iStrSlCd
		End If
		Unit = ConvSPChars(E6_m_supplier_item_by_plant(M106_E6_pur_unit))
		
		Call B1b119.B_LOOK_UP_ITEM_BY_PLANT(gStrGlobalCollection, _
											Trim(arrVal(2)), _
											Trim(arrVal(1)), _
											E1_b_pur_org, _
											E2_b_item_group, _
											E3_for_issued_b_storage_location, _
											E4_for_major_b_storage_location, _
											E5_i_material_valuation, _
											E6_b_item_by_plant, _
											E7_b_item, _
											E8_b_plant)
		
		'-----------------------
		'Com action	result check area(OS,internal)
		'-----------------------
		'Lookup 실패해도 에러메세지는 뿌리지 않고 SKIP함. 200308
		'If CheckSYSTEMError(Err,True) = true Then
		'	Set B1b119 = Nothing
		'	Set M31429 = Nothing 		
		'	Exit Sub
		'End If
		
		Err.Clear

		If Trim(SlCd) = "" Then
			SlCd = ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_cd))
			SlNm = ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_nm))
		End if
		
		If Trim(""&Unit) = "" Then
			Unit = ConvSPChars(E6_b_item_by_plant(P003_E6_order_unit_pur))
		End if
	
		trackingFlg = E6_b_item_by_plant(P003_E6_tracking_flg)

		Response.Write "<Script language=vbs> " 				& vbCr         
		Response.Write " With Parent.frm1.vspdData"      		& vbCr
		Response.Write " 	.Row  	=  " & arrVal(0)   			& vbCr
		Response.Write " 	.Col 	= Parent.C_PlantNm "       	& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E8_b_plant(P003_E8_plant_nm)) & """" 	& vbCr
		Response.Write " 	.Col 	= Parent.C_ItemNm "        											& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E7_b_item(P003_E7_item_nm)) & """" 				& vbCr
		Response.Write " 	.Col 	= Parent.C_SpplSpec "        										& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E7_b_item(P003_E7_spec)) & """" 				& vbCr
		Response.Write " 	.Col 	= Parent.C_OrderUnit "        										& vbCr
		Response.Write "   If Trim(.text) = """" Then "         										& vbCr	
		Response.Write "      .text   = """ & Unit & """" 												& vbCr
		Response.Write "   End If  " 																	& vbCr	
			
		Response.Write "	Dim strPrNo" & VbCr
		Response.Write "	.Col 	= parent.C_PrNo   " 				& vbCr
		Response.Write "	strPrNo = .text  " 							& vbCr
		Response.Write "	.Col 	= parent.C_TrackingNo   " 			& vbCr
		
		'Response.Write "	If strPrNo="""" and .text <> """" Then "	& vbCr
		'수정(2003.05)
		Response.Write "	If strPrNo="""" Then "	& vbCr
		Response.Write "		If  """ & Trim(UCase(trackingFlg)) & """ <> ""Y"" Then " 									& vbCr	
		Response.Write "  	  		parent.ggoSpread.spreadlock parent.C_TrackingNo, .Row, parent.C_TrackingNoPop, .Row " 	& vbCr	
		Response.Write "      		.Col 	= Parent.C_TrackingNo "    														& vbCr	
		Response.Write "  			.text   = ""*""" 																		& vbCr
		Response.Write "		Else   " 																					& vbCr
		Response.Write "   			parent.ggoSpread.spreadUnlock parent.C_TrackingNo, .Row, parent.C_TrackingNoPop, .Row   " & vbCr
		Response.Write "   			parent.ggoSpread.sssetrequired parent.C_TrackingNo, .Row, .Row   " 						& vbCr
		Response.Write "   			.Col 	= parent.C_TrackingNo   " 						& vbCr
		Response.Write "   			.text = """"   " 										& vbCr
		Response.Write "		End If "             & vbCr
		Response.Write "	End If" & vbCr
		'2003.06.05(포맷팅 함수 수정)
		Response.Write " 	.Col 	= Parent.C_Cost "       																	& vbCr
		'Response.Write " 	.text   = """ & UNINumClientFormat(E1_m_supplier_item_price_pur_prc(0),ggUnitCost.DecPoint,0) & """" 	& vbCr
		Response.Write " 	.text   = """ & UNIConvNumDBToCompanyByCurrency(E1_m_supplier_item_price_pur_prc(0), ConvSPChars(I1_m_supplier_item_price(M106_I1_pur_cur)), ggUnitCostNo,"X","X") & """" 	& vbCr
		
		If Trim(ConvSPChars(E1_m_supplier_item_price_pur_prc(1))) <> "" Then
		Response.Write " 	.Col 	= Parent.C_CostConCd "       	& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E1_m_supplier_item_price_pur_prc(1)) & """" 	& vbCr				
		End If
		Response.Write " 	.Col 	= Parent.C_Over "        																	& vbCr
		Response.Write " 	.text   = """ & UNINumClientFormat(E6_m_supplier_item_by_plant(M106_E6_over_tol),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write " 	.Col 	= Parent.C_Under "        																	& vbCr
		Response.Write " 	.text   = """ & UNINumClientFormat(E6_m_supplier_item_by_plant(M106_E6_under_tol),ggExchRate.DecPoint,0) & """" & vbCr

		Response.Write " 	.Col 	= C_PrNo "        																			& vbCr
		Response.Write "   If Trim(.text) = """" Then "         																& vbCr	
		Response.Write "      .Col   = parent.C_OrderUnit " 																	& vbCr
		Response.Write "      .text   = """ & Unit & """" 																		& vbCr
		Response.Write "   End If  " 																							& vbCr	
		
		Response.Write "   If  """ & Request("hdnImportFlg") & """ = ""Y"" Then " 												& vbCr	
		Response.Write "      .Col 	= Parent.C_HSCd "    																		& vbCr	
		Response.Write "      .text   = """ & ConvSPChars(E7_b_item(P003_E7_hs_cd)) & """" 										& vbCr
		Response.Write "      .Col 	= Parent.C_HSNm "    																		& vbCr	
		Response.Write "      .text   = """ & ConvSPChars(E5_b_hs_code(M106_E5_hs_nm)) & """" 									& vbCr
		Response.Write "   End If   " 																							& vbCr
		
		Response.Write "      .Col 	= Parent.C_SlCd "    																		& vbCr	
		Response.Write "      .text   = """ & SlCd & """" 																		& vbCr
		Response.Write "      .Col 	= Parent.C_SlNm "    																		& vbCr	
		Response.Write "      .text   = """ & SlNm & """" 																		& vbCr
		
		'2002-03-09 추가 품목 Spec, 품목별 부가세 
		'수정(2003.05)
		Response.Write "      .Col 	= Parent.C_SpplSpec "    & vbCr	
		Response.Write "      .text   = """ & ConvSPChars(E2_b_item(M106_E2_spec)) & """" & vbCr

		Response.Write "   If  """ & trim(E2_b_item(M106_E2_vat_type)) & """ <> """" Then " & vbCr	
		Response.Write "      .Col 	= Parent.C_VatNm "    & vbCr				'품목별 부가세명 
		Response.Write "      If Trim(.Text) = """" then "    & vbCr				'품목별 부가세명 
		Response.Write "      	.Col 	= Parent.C_VatType "    & vbCr				'품목별 부가세 
		Response.Write "      	.text   = """ & ConvSPChars(E2_b_item(M106_E2_vat_type)) & """" & vbCr
		Response.Write "      	.Col 	= Parent.C_VatNm "    & vbCr				'품목별 부가세명 
		Response.Write "      	.text   = """ & ConvSPChars(E7_b_minor_vat(M106_E7_minor_nm)) & """" & vbCr
		Response.Write "      	.Col 	= Parent.C_VatRate "    & vbCr				'품목별 부가세명 
		Response.Write "      	.text   = """ & UNINumClientFormat(E2_b_item(M106_E2_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write "	  	Call Parent.setVatType(.Row) " & vbCr
		Response.Write "   	  End If   	" 																							& vbCr
		Response.Write "   End If   	" & vbCr
			
		Response.Write "	  Call parent.vspdData_Change(parent.C_OrderUnit , .Row ) " & vbCr
		Response.Write "	  Call parent.vspdData_Change(parent.C_Cost , .Row ) " & vbCr
		
		Response.Write " End With "             & vbCr		
		Response.Write "</Script> "             & vbCr		
	
'		Set B1b119 = Nothing
	
   Next
	
	Set B1b119 = Nothing
	Set M31429 = Nothing

	
'    For LngRow = 1 To iLngMaxRow
    
'	    Response.Write "<Script language=vbs> " & vbCr         
'		Response.Write "	Call Parent.vspdData_Change(parent.C_OrderUnit , " & ConvSPChars(arrVal(0)) & " ) "  & vbCr
'		Response.Write " </Script> "

'    Next
    

End Sub
'============================================================================================================
' Name : SubLookUpItemPlantForUnit
' Desc : Look Up Item By Plant For Unit
'============================================================================================================
Sub SubLookUpItemPlantForUnit()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt	
	Dim LngMaxRow
	Dim M31429																	'☜: Group Count
	Dim LngRow
	
	Dim I1_m_supplier_item_price
    Const M106_I1_pur_unit = 0    'View Name : imp m_supplier_item_price
    Const M106_I1_pur_cur = 1
    Const M106_I1_valid_fr_dt = 2

	Dim E1_m_supplier_item_price_pur_prc
	Dim E2_b_item
    'Const M106_E2_item_cd = 0    'View Name : exp b_item
    'Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 3
    Const M106_E2_vat_rate = 4
	
	Dim E3_b_plant
    Const M106_E3_plant_cd = 0    'View Name : exp b_plant
    Const M106_E3_plant_nm = 1
	Dim E4_b_storage_location
    Const M106_E4_sl_cd = 0    'View Name : exp b_storage_location
    Const M106_E4_sl_nm = 1
	Dim E5_b_hs_code
    Const M106_E5_hs_cd = 0    'View Name : exp b_hs_code
    Const M106_E5_hs_nm = 1
	Dim E6_m_supplier_item_by_plant
    Const M106_E6_pur_priority = 0    'View Name : exp m_supplier_item_by_plant
    Const M106_E6_pur_unit = 1
    Const M106_E6_sppl_item_cd = 2
    Const M106_E6_sppl_item_nm = 3
    Const M106_E6_sppl_item_spec = 4
    Const M106_E6_maker_nm = 5
    Const M106_E6_valid_fr_dt = 6
    Const M106_E6_valid_to_dt = 7
    Const M106_E6_usage_flg = 8
    Const M106_E6_sppl_sales_prsn = 9
    Const M106_E6_sppl_tel_no = 10
    Const M106_E6_sppl_dlvy_lt = 11
    Const M106_E6_under_tol = 12
    Const M106_E6_over_tol = 13
    Const M106_E6_min_qty = 14
    Const M106_E6_def_flg = 15
    Const M106_E6_max_qty = 16
	Dim E7_b_minor_vat
    Const M106_E7_minor_nm = 1    'View Name : exp_vat_nm b_minor
    Const M106_E7_minor_cd = 0
	Dim iStrSupplierCd
	
	'이성룡 추가  단가type 의 유무를 조사 
	If CheckPriceType(iPriceType) = False then
        Call DisplayMsgBox("171214", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Exit Sub
    End If		
	
	LngMaxRow = CInt(Request("txtMaxRows"))	

    Set M31429 = Server.CreateObject("PM3G1P9.cMLookupPriceS")
    
    If CheckSYSTEMError(Err,True) = true then
		Set M31429 = Nothing 		
		Exit Sub
	End If


	arrTemp = Split(Request("txtSpread"), gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 0
    
    redim I1_m_supplier_item_price(2)
    
    For LngRow = 1 To LngMaxRow
    
		lGrpCnt = lGrpCnt +1														'☜: Group Count
	
		arrVal = Split(arrTemp(LngRow-1), gColSep)
	
		I1_m_supplier_item_price(M106_I1_pur_unit)		= Trim(arrVal(3))
		I1_m_supplier_item_price(M106_I1_pur_cur)		= Trim(Request("txtCurr"))
		I1_m_supplier_item_price(M106_I1_valid_fr_dt)	= UNIConvDate(Request("txtPoDt"))
		
		iStrSupplierCd = Trim(Request("txtSupplierCd"))
		Call M31429.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, _
									   I1_m_supplier_item_price, _
									   iStrSupplierCd, _
									   Trim(arrVal(1)), _
									   Trim(arrVal(2)), _
									   iPriceType , _
									   E1_m_supplier_item_price_pur_prc, _
									   E2_b_item, _
									   E3_b_plant, _
									   E4_b_storage_location, _
									   E5_b_hs_code, _
									   E6_m_supplier_item_by_plant, _
									   E7_b_minor_vat)
		Err.Clear
		If CheckSYSTEMError(Err,True) = true Then
			Set M31429 = Nothing 		
			Exit Sub
		End If

		Set M31429 = Nothing

		Response.Write "<Script language=vbs> " & vbCr         
		Response.Write " With Parent.frm1.vspdData "      	& vbCr
		Response.Write " .Row  =  " & arrVal(0)   			& vbCr  '이부분 주의... 처리요!!!
		Response.Write " .Col 	= Parent.C_Cost "       	& vbCr
		'2003.06.05(포맷팅 함수 수정)
		'Response.Write " .text   = """ & UNINumClientFormat(E1_m_supplier_item_price_pur_prc(0),ggUnitCost.DecPoint,0) & """" & vbCr
		Response.Write " .text   = """ & UNIConvNumDBToCompanyByCurrency(E1_m_supplier_item_price_pur_prc(0), ConvSPChars(I1_m_supplier_item_price(M106_I1_pur_cur)), ggUnitCostNo,"X","X") & """" 	& vbCr
		
		If Trim(ConvSPChars(E1_m_supplier_item_price_pur_prc(1))) <> "" Then
		Response.Write " 	.Col 	= Parent.C_CostConCd "       	& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E1_m_supplier_item_price_pur_prc(1)) & """" 	& vbCr
		End If		
		Response.Write " .Col 	= Parent.C_Over "        	& vbCr
		Response.Write "   If Trim(.text) = """" Then "     & vbCr	
		Response.Write "      .text   = """ & UNINumClientFormat(E6_m_supplier_item_by_plant(M106_E6_over_tol),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write "   End If   " & vbCr
		Response.Write "      .Col 	= Parent.C_Under "    	& vbCr	
		Response.Write "   If Trim(.text) = """" Then "     & vbCr	
		Response.Write "      .text   = """ & UNINumClientFormat(E6_m_supplier_item_by_plant(M106_E6_under_tol),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write "   End If   " & vbCr
		Response.Write "   If  """ & Request("hdnImportFlg") & """ = ""Y"" Then " & vbCr	
		Response.Write "      .Col 	= Parent.C_HSCd "    			& vbCr	
		Response.Write "      If Trim(.text) = """" Then "         	& vbCr	
		Response.Write "      .text   = """ & ConvSPChars(E5_b_hs_code(M106_E5_hs_cd)) & """" & vbCr
		Response.Write "   	  End If   " & vbCr
		Response.Write "      .Col 	= Parent.C_HSNm "    			& vbCr	
		Response.Write "      If Trim(.text) = """" Then "         & vbCr	
		Response.Write "      .text   = """ & ConvSPChars(E5_b_hs_code(M106_E5_hs_nm)) & """" & vbCr
		Response.Write "   	  End If   " & vbCr
		Response.Write "      .Col 	= Parent.C_SlCd "    & vbCr	
		Response.Write "      If Trim(.text) = """" Then "         & vbCr
		Response.Write "      .text   = """ & ConvSPChars(E4_b_storage_location(M106_E4_sl_cd)) & """" & vbCr
		Response.Write "   	  End If   " & vbCr
		Response.Write "      .Col 	= Parent.C_SlNm "    & vbCr	
		Response.Write "      If Trim(.text) = """" Then "         & vbCr
		Response.Write "      .text   = """ & ConvSPChars(E4_b_storage_location(M106_E4_sl_nm)) & """" & vbCr
		Response.Write "   	  End If   " & vbCr
		Response.Write "   End If   " & vbCr
	
		'2002-03-09 추가 품목 Spec, 품목별 부가세 
		Response.Write "      .Col 	= Parent.C_SpplSpec "    & vbCr	
		Response.Write "      .text   = """ & ConvSPChars(E2_b_item(M106_E2_spec)) & """" & vbCr
		
		Response.Write "   If  """ & ConvSPChars(E2_b_item(M106_E2_vat_type)) & """ <> """" Then " & vbCr	
		Response.Write "      .Col 	= Parent.C_VatType "    & vbCr				'품목별 부가세 
		Response.Write "      .text   = """ & ConvSPChars(E2_b_item(M106_E2_vat_type)) & """" & vbCr
		Response.Write "      .Col 	= Parent.C_VatNm "    & vbCr				'품목별 부가세명 
		Response.Write "      .text   = """ & ConvSPChars(E7_b_minor_vat(M106_E7_minor_nm)) & """" & vbCr
		Response.Write "      .Col 	= Parent.C_VatRate "    & vbCr				'품목별 부가세명 
		Response.Write "      .text   = """ & UNINumClientFormat(E2_b_item(M106_E2_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write "	  Call Parent.setVatType(.Row) " & vbCr
		Response.Write "   End If   " & vbCr
		
		Response.Write "	  Call Parent.vspdData_Change(parent.C_Cost , .Row ) " & vbCr
		
		Response.Write " End With "             & vbCr		
		Response.Write "</Script> "             & vbCr		

    Next
	
	Set M31429 = Nothing
   
End Sub

'==============================================================================
' Function : CheckPriceType(이성룡 추가:단가타입 유무조사)
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function CheckPriceType(Byref lgPriceType) 

	lgPriceType = ""
    
	'lgObjRs =  Nothing

	lgStrSQL = ""
	lgStrSQL = "SELECT MINOR_CD FROM B_CONFIGURATION " & _
				"WHERE MAJOR_CD = 'M0001'" & _
				"AND REFERENCE = 'Y'"

    Call SubOpenDB(lgObjConn)
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    
		CheckPriceType = False
		Exit Function
	Else
		lgPriceType = lgObjRs(0)
		CheckPriceType = True
	End If

End Function

%>
