<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3112MB7
'*  4. Program Name         : 반품발주내역등록 
'*  5. Program Desc         : 반품발주내역등록 
'*  6. Component List       : PM3G119.cMLookupPurOrdHdrS/PM3G128.cMListPurOrdDtlS/PM3G121.cMMaintPurOrdDtlS/PM3G1R1.cMReleasePurOrdS/PM3G1P9.cMLookupPriceS/PB3S106.cBLkUpItemByPlt
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
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

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
    
    Dim lgOpModeCRUD
 	Dim iPriceType    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	
    lgOpModeCRUD  = Request("txtMode") 
										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
                
		Case CStr("Release"),CStr("UnRelease")			    '☜: 확정,확정취소 요청을 받음 
			 Call SubRelease()								'===> 미완성 
		
		Case CStr("LookUpItemPlant")
			 Call SubLookUpItemPlant()
		
		Case CStr("LookUpItemPlantForUnit")
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
    Const M239_E22_cls_flg = 48
    Const M239_E22_ref_no = 49			'추가 (이정태)
    Const M239_E22_tot_vat_loc_amt = 50
    
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
	
	iStrPoNo = UCase(Trim(Request("txtPoNo")))
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
		Response.Write "<Script Language=vbscript>"	& vbCr
		Response.Write "parent.frm1.txtPoNo.focus"	& vbCr
		Response.Write "</Script>" & vbCr
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	End If

	Set M31119 = Nothing												'☜: ComProxy Unload
	
	If UCase(Trim(E22_m_pur_ord_hdr(M239_E22_ret_flg))) <> "Y" then
	   Call DisplayMsgBox("17a014", vbInformation, "일반발주등록건", "조회", I_MKSCRIPT)
       Set M31119 = Nothing																	'☜: ComProxy UnLoad
		Response.Write "<Script Language=vbscript>"	& vbCr
		Response.Write "parent.frm1.txtPoNo.focus"	& vbCr
		Response.Write "</Script>" & vbCr
	   Exit Sub
    End if 
	
	lgCurrency = ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur))
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "if .frm1.vspdData.MaxRows = 0 then " & vbCr
	Response.Write "	.frm1.txtCurr.value 		= """ & ConvSPChars(UCase(Trim(E22_m_pur_ord_hdr(M239_E22_po_cur))))	 			& """" & vbCr
	Response.Write " 	parent.CurFormatNumericOCX " &vbCr
	Response.Write "	.frm1.txtCurrNm.value		= """ & ConvSPChars(E18_b_currency)       							& """" & vbCr
	Response.Write "	.frm1.txtGrossAmt.value		= """ & UNIConvNumDBToCompanyByCurrency(E22_m_pur_ord_hdr(M239_E22_tot_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbCr
	Response.Write "	.frm1.txtSupplierCd.value	= """ & ConvSPChars(UCase(Trim(E9_b_biz_partner(M239_E9_bp_cd))))              	& """" & vbCr
	Response.Write "	.frm1.txtSupplierNm.value	= """ & ConvSPChars(E9_b_biz_partner(M239_E9_bp_nm))              	& """" & vbCr
	Response.Write "	.frm1.txtGroupCd.value		= """ & ConvSPChars(UCase(Trim(E20_b_pur_grp(M239_E20_pur_grp))))                	& """" & vbCr
	Response.Write "	.frm1.txtGroupNm.value		= """ & ConvSPChars(E20_b_pur_grp(M239_E20_pur_grp_nm))              	& """" & vbCr
	Response.Write "	.frm1.txtPoTypeCd.value		= """ & ConvSPChars(UCase(Trim(E19_m_config_process(M239_E19_po_type_cd))))       	& """" & vbCr
	Response.Write "	.frm1.txtPoTypeNm.value		= """ & ConvSPChars(E19_m_config_process(M239_E19_po_type_nm))       	& """" & vbCr
	Response.Write "	.frm1.txtRelease.value		= """ & ConvSPChars(UCase(Trim(E22_m_pur_ord_hdr(M239_E22_release_flg)))) 			& """" & vbCr
	Response.Write "	.frm1.txthdnPoNo.value		= """ & ConvSPChars(UCase(Trim(E22_m_pur_ord_hdr(M239_E22_po_no))))               	& """" & vbCr
	Response.Write "	.frm1.txtPoNo.value			= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_no))               	& """" & vbCr
	Response.Write "	.frm1.txtPoDt.text			= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_po_dt))       	& """" & vbCr
	Response.Write "	.frm1.hdnDlvydt.value		= """ & UNIDateClientFormat(E22_m_pur_ord_hdr(M239_E22_fore_dvry_dt)) & """" & vbCr
	Response.Write "	.frm1.hdnImportFlg.value	= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_import_flg))          	& """" & vbCr
	Response.Write "	.frm1.hdnSubcontraflg.value = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_subcontra_flg))     	& """" & vbCr
	Response.Write "	.frm1.hdnClsflg.value		= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_cls_flg))             	& """" & vbCr
	Response.Write "	.frm1.hdnReleaseflg.value	= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_release_flg)) 			& """" & vbCr
	Response.Write "	.frm1.hdnRcptflg.value		= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_rcpt_flg))            	& """" & vbCr
	Response.Write "	.frm1.hdnRcptType.value		= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_rcpt_type))           	& """" & vbCr
	Response.Write "	.frm1.hdnRetflg.value		= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_ret_flg))             	& """" & vbCr
	Response.Write "	.frm1.hdnIVFlg.value		= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_iv_flg))              	& """" & vbCr
	Response.Write "	.frm1.hdnMvmtType.value		= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_rcpt_type))             	& """" & vbCr
	
	' ### 환율 & 환율연산자 #####
	Response.Write "	.frm1.hdnXch.value			= """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_xch_rt),ggExchRate.DecPoint,0)  & """" & vbCr
	Response.Write "    .frm1.hdnRefPoNo.Value		= """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_ref_no)) & """" & vbCr
	
	If ConvSPChars(E22_m_pur_ord_hdr(M239_E22_xch_rate_op)) <> "" Then
		Response.Write "	.frm1.hdnXchRateOp.value = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_xch_rate_op))           & """" & vbCr
	Else
		Response.Write "	.frm1.hdnXchRateOp.value = ""*""" & vbCr '환율 연산자가 없는 경우 Default value *
	End If

	' ### VAT Append #####
	Response.Write "	.frm1.hdnVATType.value   = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_vat_type))             				  & """" & vbCr
	'Response.Write "	.frm1.hdnVATRate.value   = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_vat_rt),ggExchRate.DecPoint,0) & """" & vbCr
	
	' ### 대물정산 반품시 부가세율 0 처리 2002.04.10 L.I.P
	Response.Write "If """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_iv_flg)) & """ = ""N"" And """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_ret_flg)) & """ = ""Y"" Then " & vbCr
	Response.Write "	.frm1.hdnVATRate.value   = 0 " & vbCr
	Response.Write "Else " & vbCr
	Response.Write "	.frm1.hdnVATRate.value   = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_vat_rt),ggExchRate.DecPoint,0) & """" & vbCr
	Response.Write "End If " & vbCr
	
	If ConvSPChars(E22_m_pur_ord_hdr(M239_E22_vat_inc_flag)) = "2" Then
		Response.Write "	.frm1.hdnVATINCFLG.value = ""2""" & vbCr	'포함 
	Else
		Response.Write "	.frm1.hdnVATINCFLG.value = ""*""" & vbCr	'별도 
	End If
	Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet"	& vbCr
	Response.Write "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE"	& vbCr
	Response.Write " End if "	& vbCr

	Response.Write " End With "	& vbCr
    Response.Write "</Script>" & vbCr

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
	
	if Request("txtQuerytype") <> "Query" And Err.Description = "B_MESSAGE173200" then
		Set M31128 = Nothing
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
		Exit Sub												'☜: ComProxy Unload
	Else 
		If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set M31128 = Nothing												'☜: ComProxy Unload
			'Detail항목이 없을 경우 Header정보만 보여줌 
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			Exit Sub															'☜: 비지니스 로직 처리를 종료함 
		End If
		
	End if

	iLngMaxRow = CLng(Request("txtMaxRows"))											'Save previous Maxrow                                                
	GroupCount = UBound(EG1_exp_group,1)
	ReDim PvArr(GroupCount)

	If EG1_exp_group(GroupCount,M192_EG1_E5_m_pur_ord_dtl_po_seq_no) = E1_m_pur_ord_dtl_po_seq_no Then
		StrNextKey = ""
	Else
		StrNextKey = E1_m_pur_ord_dtl_po_seq_no
	End If
	
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	
	For iLngRow = 0 To GroupCount
        
        if iLngRow >= C_SHEETMAXROWS_D Then
			StrNextKey = EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_seq_no)
			Exit For
        End If

 		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_seq_no))	'1
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E7_b_plant_plant_cd))			'2
        istrData = istrData & Chr(11) & ""																			'3
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E7_b_plant_plant_nm))			'4
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E6_b_item_item_cd))				'5
        istrData = istrData & Chr(11) & ""																			'6
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E6_b_item_item_nm))				'7
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E6_b_item_spec))	'품목규격	'8	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_qty),ggQty.DecPoint,0) '9	       
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_unit))		'10
        istrData = istrData & Chr(11) & ""																			'11
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc),lgCurrency,ggUnitCostNo,"X" , "X")	'12		
        If ConvSPChars(UCase(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc_flg)))) = "F" Then						'13
			istrData = istrData & Chr(11) & "가단가"
		ElseIf ConvSPChars(UCase(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc_flg)))) = "T" Then
			istrData = istrData & Chr(11) & "진단가"
		End If
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_prc_flg))	'14
		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt)),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
        
        If ConvSPChars(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))) = "2" Then							'13
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency((CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt)) + CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt))),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		Else
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		End If
		
		If ConvSPChars(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))) = "2" Then							'13
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency((CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt)) + CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt))),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		Else
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		End If
		
		
		If ConvSPChars(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))) = "2" Then							'13
			istrData = istrData & Chr(11) & "포함"
			'<--포함구분코드 -->
			'istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))
		Else
			istrData = istrData & Chr(11) & "별도"
			'<--포함구분코드 -->
			'istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))
		End If
		
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))	'16
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_type))		'17
		istrData = istrData & Chr(11) & ""																			'18
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E8_b_minor_minor_nm))			'19
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_rate),ggExchRate.DecPoint,0)	'20
		istrData = istrData & Chr(11) & UNINumClientFormatByTax(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt),lgCurrency,ggAmtOfMoneyNo)		'21
		istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_dlvy_dt))	'22
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E3_b_hs_code_hs_cd))			'23
		istrData = istrData & Chr(11) & ""																			'24
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E3_b_hs_code_hs_nm))			'25
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E4_b_storage_location_sl_cd))	'26
		istrData = istrData & Chr(11) & ""																			'27
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E4_b_storage_location_sl_nm))	'28
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_ret_type))		'33
        istrData = istrData & Chr(11) & ""																			'34
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E9_b_minor_minor_nm))			'35
    
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_tracking_no))	'29
		istrData = istrData & Chr(11) & ""																			'30
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_lot_no))		'31
		istrData = istrData & Chr(11) & ""																			'34
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_lot_sub_no))	'32
		
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_over_tol),ggExchRate.DecPoint,0)	'36	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_under_tol),ggExchRate.DecPoint,0)	'37
        
        istrData = istrData & Chr(11) & ""																			'38
        istrData = istrData & Chr(11) & ""																			'39
        istrData = istrData & Chr(11) & ""																			'40
        istrData = istrData & Chr(11) & ""																			'41
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E1_m_pur_req_pr_no))				'43
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E5_m_pur_ord_dtl_ref_mvmt_no))	'44
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E5_m_pur_ord_dtl_ref_po_no))		'45
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E5_m_pur_ord_dtl_ref_po_seq_no))	'46
        istrData = istrData & Chr(11) & ""																			'47
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E5_m_pur_ord_dtl_so_no))			'48
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E5_m_pur_ord_dtl_so_seq_no))		'49
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E5_m_pur_ord_dtl_ref_iv_no))		'50
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M192_EG1_E5_m_pur_ord_dtl_ref_iv_seq))	'51
        
        If ConvSPChars(Trim(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_inc_flag))) = "2" Then							'13
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency((CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt)) + CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_vat_doc_amt))),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		Else
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  '15
		End If                
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt)),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(CDbl(EG1_exp_group(iLngRow, M192_EG1_E5_m_pur_ord_dtl_po_doc_amt)),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
        
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
   Response.Write "	    If .frm1.txtRelease.Value = ""Y"" then"				& vbCr
   Response.Write "	        For index = .C_SeqNo to .C_Stateflg"			& vbCr
   Response.Write "			        .ggoSpread.SpreadLock index , -1"		& vbCr
   Response.Write "		    Next"					& vbCr
   Response.Write "	    Else"						& vbCr
   Response.Write "		    .SetSpreadLock"			& vbCr
   Response.Write "		End If"						& vbCr
   Response.Write "	End if"							& vbCr

    Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
    Response.Write "	.lgStrPrevKey           = """ & StrNextKey   & """" & vbCr  
    Response.Write " .frm1.txthdnPoNo.value		= """ & ConvSPChars(Request("txtPoNo")) & """" & vbCr
    
	'Response.Write "If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> """" Then " & vbCr
	'Response.Write "	.DbQuery "  & vbCr	' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
	'Response.Write "Else " 			& vbCr
    Response.Write "    .DbQueryOk "	& vbCr 
	'Response.Write "End If "		& vbCr
    Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr    
		
    Set M31128 = Nothing
		
End Sub    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
	Err.Clear	
	
	Dim M31121																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim LngMaxRow
	Dim iErrorPosition
	
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
    Const L1_remrk = 45
'-------------------------------------------------
'-------------------------------------------------
    Const L1_po_no = 45             '발주번호 
    Const L1_po_seq_no = 46         '발주SeqNo
    Const L1_maint_seq = 47         'maintseq
    Const L1_so_no = 48
    Const L1_so_seq_no = 49
    Const L1_state_flg = 50
    Const L1_row_num = 51
	Dim iStrPoNo
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim iDCount
    Dim ii
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
	
	LngMaxRow = CLng(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

    Set M31121 = Server.CreateObject("PM3G121.cMMaintPurOrdDtlS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = true Then 		
		Set M31121 = Nothing												'☜: ComPlus Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

	iStrPoNo   = UCase(Trim(Request("txthdnPoNo")))
	Call M31121.M_MAINT_PUR_ORD_DTL_SVR("F",gStrGlobalCollection, _
						             	LngMaxRow, _
						             	gCurrency, _
						             	itxtSpread, _
						             	iStrPoNo, _
						             	iErrorPosition)

	If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
		Set M31121 = Nothing
		Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
		Exit Sub
	End If

    Set M31121 = Nothing                                                   '☜: Unload Comproxy  
   
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.frm1.txtPoNo.Value = """ & ConvSPChars(Request("txthdnPoNo")) & """" & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "   & vbCr	         
       
End Sub    


'============================================================================================================
' Name : SubRelease
' Desc : 발주확정 
'============================================================================================================
Sub SubRelease()
														
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
	
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set M31211 = Nothing												'☜: ComProxy Unload
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.btnCfmSel.disabled = False" & vbCr
			Response.Write "</Script>" & vbCr
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If

    Set M31211 = Nothing                                                   '☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" 					& vbCr
	Response.Write "With parent"									& vbCr	
	Response.Write ".DbSaveOk" & vbCr
	Response.Write "End With"   & vbCr
	Response.Write "</Script>" & vbCr

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
    Const M106_E2_item_cd = 0    'View Name : exp b_item
    Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 0
    Const M106_E2_vat_rate = 1
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


    For LngRow = 1 To LngMaxRow  

		lGrpCnt = lGrpCnt +1	
		
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
			Set B1b119 = Nothing 		
			Exit Sub
		End If
		

		SlCd = ConvSPChars(E4_b_storage_location(M106_E4_sl_cd))
		SlNm = ConvSPChars(E4_b_storage_location(M106_E4_sl_nm))
		Unit = ConvSPChars(E6_m_supplier_item_by_plant(M106_E6_pur_unit))
		
		'Err.Clear

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
		If CheckSYSTEMError(Err,True) = true Then
			Set B1b119 = Nothing
			Set M31429 = Nothing 		
			Exit Sub
		End If
		
		Err.Clear

		If Trim(SlCd) = "" Then
			SlCd = ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_cd))
			SlNm = ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_nm))
		End if
		
		If Trim(""&Unit) = "" Then
			Unit = ConvSPChars(E6_b_item_by_plant(P003_E6_order_unit_pur))
		End if

	
		'b_item_by_plant확인... 못가져오고 있음...
		trackingFlg = E6_b_item_by_plant(P003_E6_tracking_flg)

		Response.Write "<Script language=vbs> " 				& vbCr         
		Response.Write " With Parent.frm1.vspdData"      		& vbCr
		Response.Write " 	.Row  	=  " & arrVal(0)   			& vbCr
		Response.Write " 	.Col 	= Parent.C_PlantNm "       	& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E8_b_plant(P003_E8_plant_nm)) & """" 	& vbCr
		Response.Write " 	.Col 	= Parent.C_ItemNm "        											& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E7_b_item(P003_E7_item_nm)) & """" 				& vbCr
		Response.Write " 	.Col 	= Parent.C_SpplSpec "        											& vbCr
		Response.Write " 	.text   = """ & ConvSPChars(E7_b_item(P003_E7_spec)) & """" 				& vbCr
		Response.Write VbCr
		Response.Write " 	.Col 	= Parent.C_OrderUnit "        										& vbCr
		Response.Write "   If Trim(.text) = """" Then "         										& vbCr	
		Response.Write "      .text   = """ & Unit & """" 												& vbCr
		Response.Write "   End If  " 																	& vbCr	
		Response.Write VbCr
		Response.Write "   If  """ & trackingFlg & """ <> ""Y"" Then " 									& vbCr	
		Response.Write "  	  	parent.ggoSpread.spreadlock parent.C_TrackingNo, .Row, parent.C_TrackingNoPop, .Row " 	& vbCr	
		Response.Write "      	.Col 	= Parent.C_TrackingNo "    														& vbCr	
		Response.Write "      	.text   = ""*""" 																		& vbCr
		Response.Write "   Else   " 																					& vbCr
		Response.Write "   		parent.ggoSpread.spreadUnlock parent.C_TrackingNo, .Row, parent.C_TrackingNoPop, .Row   " & vbCr
		Response.Write "   		parent.ggoSpread.sssetrequired parent.C_TrackingNo, .Row, .Row   " 						& vbCr
		Response.Write "      	.Col 	= Parent.C_TrackingNo "    														& vbCr	
		Response.Write "      	.text   = """"" 																		& vbCr
		Response.Write "   End If "             & vbCr

		Response.Write "If """ & UCase(Trim(E6_b_item_by_plant(P003_E6_lot_flg))) & """ = ""N"" then " & vbCr
		Response.Write "	parent.ggoSpread.spreadlock parent.C_Lot_No, .Row, parent.C_Lot_Seq, .Row " & vbCr
		Response.Write "	.Col 	= parent.C_Lot_No " & vbCr
		Response.Write "	.text	= ""*""" & vbCr
		Response.Write "Else " & vbCr
		Response.Write "	parent.ggoSpread.spreadUnlock parent.C_Lot_No, .Row, parent.C_Lot_Seq, .Row " & vbCr
		Response.Write "	.Col 	= parent.C_Lot_No " & vbCr
		Response.Write "	.text = """"" & vbCr
		Response.Write "End If " & vbCr

		Response.Write " 	.Col 	= Parent.C_Cost "       																	& vbCr
		Response.Write " 	.text   = """ & UNINumClientFormat(E1_m_supplier_item_price_pur_prc(0),ggUnitCost.DecPoint,0) & """" 	& vbCr
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
		Response.Write "      .Col 	= Parent.C_SpplSpec "    & vbCr	
		Response.Write "      .text   = """ & ConvSPChars(E2_b_item(M106_E2_spec)) & """" & vbCr

		Response.Write "   If  """ & Trim(E2_b_item(M106_E2_vat_type)) & """ <> """" Then " & vbCr	
		Response.Write "      .Col 	= Parent.C_VatType "    & vbCr				'품목별 부가세 
		Response.Write "      .text   = """ & ConvSPChars(E2_b_item(M106_E2_vat_type)) & """" & vbCr
		Response.Write "      .Col 	= Parent.C_VatNm "    & vbCr				'품목별 부가세명 
		Response.Write "      .text   = """ & ConvSPChars(E7_b_minor_vat(M106_E7_minor_nm)) & """" & vbCr
		Response.Write "      .Col 	= Parent.C_VatRate "    & vbCr				'품목별 부가세명 
		Response.Write "      .text   = """ & UNINumClientFormat(E2_b_item(M106_E2_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
		Response.Write "	  Call Parent.setVatType(.Row) " & vbCr
		Response.Write "   End If   " & vbCr
		Response.Write VbCr
		
		Response.Write "	  Call parent.vspdData_Change(parent.C_Cost , .Row ) " & vbCr
		
		Response.Write " End With "             & vbCr		
		Response.Write "</Script> "             & vbCr		
	
'		Set B1b119 = Nothing
	
   Next
	
	Set B1b119 = Nothing
	Set M31429 = Nothing

	
    For LngRow = 1 To iLngMaxRow
    
	    Response.Write "<Script language=vbs> " & vbCr         
		Response.Write "	Call Parent.vspdData_Change(parent.C_OrderUnit , " & ConvSPChars(arrVal(0)) & " ) " & vbCr
		Response.Write " </Script> "

    Next
    

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
   ' Const M106_E2_item_cd = 0    'View Name : exp b_item
   ' Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 0
    Const M106_E2_vat_rate = 1
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
		Response.Write " .text   = """ & UNINumClientFormat(E1_m_supplier_item_price_pur_prc(0),ggUnitCost.DecPoint,0) & """" & vbCr
		
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
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	Dim strHTML
	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
%>
