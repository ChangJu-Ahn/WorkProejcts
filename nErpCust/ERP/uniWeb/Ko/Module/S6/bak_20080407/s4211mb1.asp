<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4211mb1.asp																*
'*  4. Program Name         : 통관등록																	*
'*  5. Program Desc         : 통관등록																	*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kim Hyungsuk																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
  
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	Call HideStatusWnd
 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
Sub SubBizQuery()

    Dim pS6G219
    Dim iCommandSent
    
    Dim I1_s_cc_hdr_cc_no 
    Dim EG1_s_cc_hdr 
    Dim l2_s_cc_dtl_bill_count

    Const S446_E1_cc_no = 0    
    Const S446_E1_iv_no = 1
    Const S446_E1_iv_dt = 2
    Const S446_E1_so_no = 3
    Const S446_E1_manufacturer = 4
    Const S446_E1_agent = 5
    Const S446_E1_return_appl = 6
    Const S446_E1_return_office = 7
    Const S446_E1_reporter = 8
    Const S446_E1_ed_no = 9
    Const S446_E1_ed_dt = 10
    Const S446_E1_ed_type = 11
    Const S446_E1_inspect_type = 12
    Const S446_E1_ep_type = 13
    Const S446_E1_export_type = 14
    Const S446_E1_incoterms = 15
    Const S446_E1_pay_meth = 16
    Const S446_E1_pay_dur = 17
    Const S446_E1_dischge_cntry = 18
    Const S446_E1_loading_port = 19
    Const S446_E1_loading_cntry = 20
    Const S446_E1_dischge_port = 21
    Const S446_E1_vessel_nm = 22
    Const S446_E1_transport = 23
    Const S446_E1_trans_form = 24
    Const S446_E1_inspect_req_dt = 25
    Const S446_E1_device_plce = 26
    Const S446_E1_lc_doc_no = 27
    Const S446_E1_lc_amend_seq = 28
    Const S446_E1_lc_no = 29
    Const S446_E1_lc_type = 30
    Const S446_E1_lc_open_dt = 31
    Const S446_E1_open_bank = 32
    Const S446_E1_gross_weight = 33
    Const S446_E1_weight_unit = 34
    Const S446_E1_tot_packing_cnt = 35
    Const S446_E1_packing_type = 36
    Const S446_E1_ship_fin_dt = 37
    Const S446_E1_cur = 38
    Const S446_E1_doc_amt = 39
    Const S446_E1_fob_doc_amt = 40
    Const S446_E1_xch_rate = 41
    Const S446_E1_loc_amt = 42
    Const S446_E1_fob_loc_amt = 43
    Const S446_E1_freight_loc_amt = 44
    Const S446_E1_insure_loc_amt = 45
    Const S446_E1_el_doc_no = 46
    Const S446_E1_el_app_dt = 47
    Const S446_E1_ep_no = 48
    Const S446_E1_ep_dt = 49
    Const S446_E1_insp_cert_no = 50
    Const S446_E1_insp_cert_dt = 51
    Const S446_E1_quar_cert_no = 52
    Const S446_E1_quar_cert_dt = 53
    Const S446_E1_recomnd_no = 54
    Const S446_E1_recomnd_dt = 55
    Const S446_E1_trans_method = 56
    Const S446_E1_trans_rep_cd = 57
    Const S446_E1_trans_from_dt = 58
    Const S446_E1_trans_to_dt = 59
    Const S446_E1_customs = 60
    Const S446_E1_final_dest = 61
    Const S446_E1_remark1 = 62
    Const S446_E1_remark2 = 63
    Const S446_E1_remark3 = 64
    Const S446_E1_origin = 65
    Const S446_E1_origin_cntry = 66
    Const S446_E1_usd_xch_rate = 67
    Const S446_E1_biz_area = 68
    Const S446_E1_ref_flag = 69
    Const S446_E1_sts = 70
    Const S446_E1_net_weight = 71
    Const S446_E1_ext1_qty = 72
    Const S446_E1_ext2_qty = 73
    Const S446_E1_ext3_qty = 74
    Const S446_E1_ext1_amt = 75
    Const S446_E1_ext2_amt = 76
    Const S446_E1_ext3_amt = 77
    Const S446_E1_ext1_cd = 78
    Const S446_E1_ext2_cd = 79
    Const S446_E1_ext3_cd = 80
    Const S446_E1_xch_rate_op = 81

    Const S446_E2_bp_cd = 82
    Const S446_E2_bp_nm = 83
    
    Const S446_E3_bp_cd = 84
    Const S446_E3_bp_nm = 85
    
    Const S446_E4_sales_grp = 86
    Const S446_E4_sales_grp_nm = 87
    
    Const S446_E5_sales_org = 88
    Const S446_E5_sales_org_nm = 89

    Const S446_E6_bp_nm = 90    
    
    Const S446_E7_bp_nm = 91    
        
    Const S446_E8_bp_nm = 92    

    Const S446_E9_bp_nm = 93    

    Const S446_E10_bank_nm = 94 

    Const S446_E11_minor_nm = 95

    Const S446_E12_minor_nm = 96

    Const S446_E13_minor_nm = 97

    Const S446_E14_minor_nm = 98
    
    Const S446_E15_minor_nm = 99
    
    Const S446_E16_minor_nm = 100
    
    Const S446_E17_minor_nm = 101
    
    Const S446_E18_minor_nm = 102
    
    Const S446_E19_country_nm = 103
    
    Const S446_E20_minor_nm = 104
    
    Const S446_E21_country_nm = 105
    
    Const S446_E22_minor_nm = 106
    
    Const S446_E23_country_nm = 107

    Const S446_E24_minor_nm = 108
    
    Const S446_E25_minor_nm = 109

    Const S446_E26_minor_nm = 110
    
    Const S446_E27_minor_nm = 111

    Const S446_E28_bp_nm = 112

    Const S446_E29_minor_nm = 113
    
    Const S446_E30_carton_cnt = 114
    Const S446_E30_measurement=115
    


    On Error Resume Next
    Err.Clear 

    If Request("txtCCNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If
	
	I1_s_cc_hdr_cc_no  = Trim(Request("txtCCNo"))
	
	Select Case Request("txtPrevNext")
	Case "PREV"
			iCommandSent = "PREV"
	Case "NEXT"
			iCommandSent = "NEXT"
	Case Else 
			iCommandSent = "LOOKUP"
	End Select		
	
    Set pS6G219 = Server.CreateObject("PS6G219.cSLkExportCcHdrSvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
    
    Call pS6G219.S_LOOKUP_EXPORT_CC_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_cc_hdr_cc_no, EG1_s_cc_hdr, l2_s_cc_dtl_bill_count)

	If CheckSYSTEMError(Err,True) = True Then
       Set pS6G219 = Nothing
		Response.Write "<Script language=vbs>  " & vbCr   
		Response.Write " Parent.frm1.txtCCNo.focus  " & vbCr   		
		Response.Write "</Script>      " & vbCr      
       Exit Sub
    End If  
    
    Set pS6G219 = Nothing
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr
	
	Response.Write ".txtCCCurrency.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_cur))       & """" & vbCr
	Response.Write ".txtFobCurrency.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_cur))       & """" & vbCr
	Response.Write "parent.CurFormatNumericOCX" & vbCr

	'Tab 1
	Response.Write ".txtCCNo.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_cc_no))       & """" & vbCr
	Response.Write ".txtCCNo1.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_cc_no))       & """" & vbCr
	Response.Write ".txtSONo.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_so_no))       & """" & vbCr
	Response.Write ".txtIVNo.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_iv_no))       & """" & vbCr
	Response.Write ".txtLCDocNo.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_lc_doc_no))   & """" & vbCr
	
	If EG1_s_cc_hdr(S446_E1_lc_doc_no)   = ""	and		Trim(EG1_s_cc_hdr(S446_E1_lc_amend_seq)) = 0 Then 
	   	Response.Write ".txtLCAmendSeq.Value		= """"" & vbCr
	else
	   	Response.Write ".txtLCAmendSeq.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_lc_amend_seq))       & """" & vbCr
	End If   

	Response.Write ".txtLCNo.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_lc_no))       & """" & vbCr
	Response.Write ".txtEDNo.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_ed_no))       & """" & vbCr

	Response.Write ".txtIVDt.text				= """ & ConvSPChars(UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_iv_dt)))		& """" & vbCr
	
	Response.Write ".txtEDType.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_ed_type))     & """" & vbCr
	Response.Write ".txtEDTypeNm.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E12_minor_nm))   & """" & vbCr

	Response.Write ".txtEDDt.text				= """ & UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_ed_dt))		& """" & vbCr

	Response.Write ".txtWeightUnit.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_weight_unit))   & """" & vbCr

	Response.Write ".txtShipFinDt.text			= """ & UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_ship_fin_dt))		& """" & vbCr

	Response.Write ".txtGrossWeight.text		= """ & UNINumClientFormat(EG1_s_cc_hdr(S446_E1_gross_weight), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtNetWeight.text			= """ & UNINumClientFormat(EG1_s_cc_hdr(S446_E1_net_weight), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtVesselNm.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_vessel_nm))     & """" & vbCr
	Response.Write ".txtTotPackingCnt.Value		= """ & UNINumClientFormat(EG1_s_cc_hdr(S446_E1_tot_packing_cnt), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtCurrency.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_cur))			 & """" & vbCr		
	Response.Write ".txtXchRate.Value			= """ & UNINumClientFormat(EG1_s_cc_hdr(S446_E1_xch_rate), ggExchRate.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtCCCurrency.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_cur))			 & """" & vbCr		
    Response.Write ".txtDocAmt.Text				= """ & UNINumClientFormatByCurrency(EG1_s_cc_hdr(S446_E1_doc_amt),EG1_s_cc_hdr(S446_E1_cur),ggAmtOfMoneyNo)	& """" & vbCr	
	Response.Write ".txtLocAmt.text				= """ & UNINumClientFormat(EG1_s_cc_hdr(S446_E1_loc_amt), ggAmtOfMoney.DecPoint, 0)		& """" & vbCr
	Response.Write ".txtFobCurrency.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_cur))			 & """" & vbCr		
    Response.Write ".txtFOBDocAmt.Text			= """ & UNINumClientFormatByCurrency(EG1_s_cc_hdr(S446_E1_fob_doc_amt),EG1_s_cc_hdr(S446_E1_cur),ggAmtOfMoneyNo)	& """" & vbCr	
	Response.Write ".txtFOBLocAmt.text			= """ & UNINumClientFormat(EG1_s_cc_hdr(S446_E1_fob_loc_amt), ggAmtOfMoney.DecPoint, 0)				& """" & vbCr
	Response.Write ".txtApplicant.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E2_bp_cd))			 & """" & vbCr		
	Response.Write ".txtApplicantNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E2_bp_nm))			 & """" & vbCr		
	Response.Write ".txtBeneficiary.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E3_bp_cd))			 & """" & vbCr		
	Response.Write ".txtBeneficiaryNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E3_bp_nm))			 & """" & vbCr		
	Response.Write ".txtAgent.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_agent))			 & """" & vbCr		
	Response.Write ".txtAgentNm.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E7_bp_nm))			 & """" & vbCr		
	Response.Write ".txtManufacturer.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_manufacturer))		 & """" & vbCr		
	Response.Write ".txtManufacturerNm.Value	= """ & ConvSPChars(EG1_s_cc_hdr(S446_E6_bp_nm))			 & """" & vbCr		

	Response.Write ".txtBillCount.Value			= """ & ConvSPChars(l2_s_cc_dtl_bill_count)					 & """" & vbCr
	
	Response.Write ".txtCarton.value		=""" & UNINumClientFormat(EG1_s_cc_hdr(S446_E30_carton_cnt), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtMeasurement.value		=""" & UNINumClientFormat(EG1_s_cc_hdr(S446_E30_measurement), ggQty.DecPoint, 0)   & """" & vbCr
							
	'Tab 2
	Response.Write ".txtUSDXchRate.text			= """ & UNINumClientFormat(EG1_s_cc_hdr(S446_E1_usd_xch_rate), ggExchRate.DecPoint, 0)				& """" & vbCr
	Response.Write ".txtReturnAppl.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_return_appl))			 & """" & vbCr		
	Response.Write ".txtReturnApplNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E9_bp_nm))				 & """" & vbCr		
	Response.Write ".txtReturnOffice.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_return_office))		 & """" & vbCr		
	Response.Write ".txtReturnOfficeNm.Value	= """ & ConvSPChars(EG1_s_cc_hdr(S446_E11_minor_nm))			 & """" & vbCr		
	Response.Write ".txtLoadingPort.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_loading_port))			 & """" & vbCr		
	Response.Write ".txtLoadingPortNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E18_minor_nm))			 & """" & vbCr		
	Response.Write ".txtLoadingCntry.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_loading_cntry))		 & """" & vbCr		
	Response.Write ".txtLoadingCntryNm.Value	= """ & ConvSPChars(EG1_s_cc_hdr(S446_E19_country_nm))			 & """" & vbCr		
	Response.Write ".txtDischgePort.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_dischge_port))			 & """" & vbCr		
	Response.Write ".txtDischgePortNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E20_minor_nm))			 & """" & vbCr		
	Response.Write ".txtDischgeCntry.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_dischge_cntry))		 & """" & vbCr		
	Response.Write ".txtDischgeCntryNm.Value	= """ & ConvSPChars(EG1_s_cc_hdr(S446_E21_country_nm))			 & """" & vbCr		
	Response.Write ".txtOrigin.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_origin))				 & """" & vbCr		
	Response.Write ".txtOriginNm.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E22_minor_nm))			 & """" & vbCr		
	Response.Write ".txtOriginCntry.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_origin_cntry))			 & """" & vbCr		
	Response.Write ".txtOriginCntryNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E23_country_nm))			 & """" & vbCr		
	Response.Write ".txtFinalDest.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_final_dest))			 & """" & vbCr		
	Response.Write ".txtReporter.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_reporter))				 & """" & vbCr		
	Response.Write ".txtReporterNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E8_bp_nm))				 & """" & vbCr		
	Response.Write ".txtPayTerms.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_pay_meth))				 & """" & vbCr		
	Response.Write ".txtPayTermsNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E17_minor_nm))			 & """" & vbCr		
	Response.Write ".txtPayDur.text				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_pay_dur))				 & """" & vbCr		
	Response.Write ".txtIncoTerms.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_incoterms))			 & """" & vbCr		
	Response.Write ".txtIncoTermsNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E16_minor_nm))			 & """" & vbCr		
	Response.Write ".txtSalesGroup.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E4_sales_grp))			 & """" & vbCr		
	Response.Write ".txtSalesGroupNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E4_sales_grp_nm))			 & """" & vbCr		

	'Tab 3
	Response.Write ".txtEPNo.Value				= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_ep_no))				 & """" & vbCr		

	Response.Write ".txtEPDt.text				= """ & UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_ep_dt))		 & """" & vbCr		

	Response.Write ".txtCustoms.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_customs))				 & """" & vbCr		
	Response.Write ".txtCustomsNm.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E29_minor_nm))			 & """" & vbCr		
	Response.Write ".txtTransForm.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_trans_form))			 & """" & vbCr		
	Response.Write ".txtTransFormNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E25_minor_nm))			 & """" & vbCr		
	Response.Write ".txtPackingType.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_packing_type))			 & """" & vbCr		
	Response.Write ".txtPackingTypeNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E26_minor_nm))			 & """" & vbCr		
	Response.Write ".txtTransRepCd.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_trans_rep_cd))			 & """" & vbCr		
	Response.Write ".txtTransRepNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E28_bp_nm))				 & """" & vbCr		
	Response.Write ".txtTransMeth.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_trans_method))			 & """" & vbCr		
	Response.Write ".txtTransMethNm.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E27_minor_nm))			 & """" & vbCr		
	
	Response.Write ".txtTransFromDt.text		= """ & UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_trans_from_dt))	 & """" & vbCr		
	
	Response.Write ".txtTransToDt.text			= """ & UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_trans_to_dt))		 & """" & vbCr		

	Response.Write ".txtInspCertNo.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_insp_cert_no))				 & """" & vbCr		

	Response.Write ".txtInspCertDt.text			= """ & UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_insp_cert_dt))		 & """" & vbCr		
	
	Response.Write ".txtQuarCertNo.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_quar_cert_no))				 & """" & vbCr		

	Response.Write ".txtQuarCertDt.text			= """ & UNIDateClientFormat(EG1_s_cc_hdr(S446_E1_quar_cert_dt))		 & """" & vbCr		
		
	Response.Write ".txtDevicePlce.Value		= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_device_plce))				 & """" & vbCr		
	Response.Write ".txtRemark1.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_remark1))			 & """" & vbCr		
	Response.Write ".txtRemark2.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_remark2))			 & """" & vbCr		
	Response.Write ".txtRemark3.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_remark3))			 & """" & vbCr		
	Response.Write ".txtRefFlg.Value			= """ & ConvSPChars(EG1_s_cc_hdr(S446_E1_ref_flag))			 & """" & vbCr		

	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write "parent.ProtectXchRate" & vbCr
	Response.Write ".txtHCCNo.Value		= """ & ConvSPChars(Request("txtCCNo"))	 & """" & vbCr		
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
End Sub
'============================================================================================================
Sub SubBizSave()
    Dim pS6G211
    Dim iCommandSent
    Dim itxtFlgMode
	Dim strConvDt
	
	Dim I1_s_cc_hdr
	Dim I2_b_biz_partner_bp_cd
	Dim I3_b_biz_partner_bp_cd
	Dim I4_b_sales_grp_sales_grp
	Dim I5_s_lc_hdr_lc_no
	Dim I6_s_so_hdr_so_no
	Dim I7_s_wks_user_user_id	
	Dim E1_s_cc_hdr
	
    Const S440_I1_cc_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_cc_hdr
    Const S440_I1_iv_no = 1
    Const S440_I1_iv_dt = 2
    Const S440_I1_manufacturer = 3
    Const S440_I1_agent = 4
    Const S440_I1_return_appl = 5
    Const S440_I1_return_office = 6
    Const S440_I1_reporter = 7
    Const S440_I1_ed_no = 8
    Const S440_I1_ed_dt = 9
    Const S440_I1_ed_type = 10
    Const S440_I1_inspect_type = 11
    Const S440_I1_ep_type = 12
    Const S440_I1_export_type = 13
    Const S440_I1_incoterms = 14
    Const S440_I1_pay_meth = 15
    Const S440_I1_pay_dur = 16
    Const S440_I1_dischge_cntry = 17
    Const S440_I1_loading_port = 18
    Const S440_I1_loading_cntry = 19
    Const S440_I1_dischge_port = 20
    Const S440_I1_vessel_nm = 21
    Const S440_I1_transport = 22
    Const S440_I1_trans_form = 23
    Const S440_I1_inspect_req_dt = 24
    Const S440_I1_device_plce = 25
    Const S440_I1_gross_weight = 26
    Const S440_I1_weight_unit = 27
    Const S440_I1_tot_packing_cnt = 28
    Const S440_I1_packing_type = 29
    Const S440_I1_ship_fin_dt = 30
    Const S440_I1_cur = 31
    Const S440_I1_xch_rate = 32
    Const S440_I1_freight_loc_amt = 33
    Const S440_I1_insure_loc_amt = 34
    Const S440_I1_el_doc_no = 35
    Const S440_I1_el_app_dt = 36
    Const S440_I1_ep_no = 37
    Const S440_I1_ep_dt = 38
    Const S440_I1_insp_cert_no = 39
    Const S440_I1_insp_cert_dt = 40
    Const S440_I1_quar_cert_no = 41
    Const S440_I1_quar_cert_dt = 42
    Const S440_I1_recomnd_no = 43
    Const S440_I1_recomnd_dt = 44
    Const S440_I1_trans_method = 45
    Const S440_I1_trans_rep_cd = 46
    Const S440_I1_trans_from_dt = 47
    Const S440_I1_trans_to_dt = 48
    Const S440_I1_customs = 49
    Const S440_I1_final_dest = 50
    Const S440_I1_remark1 = 51
    Const S440_I1_remark2 = 52
    Const S440_I1_remark3 = 53
    Const S440_I1_origin = 54
    Const S440_I1_origin_cntry = 55
    Const S440_I1_usd_xch_rate = 56
    Const S440_I1_ref_flag = 57
    Const S440_I1_ext1_qty = 58
    Const S440_I1_ext2_qty = 59
    Const S440_I1_ext3_qty = 60
    Const S440_I1_ext1_amt = 61
    Const S440_I1_ext2_amt = 62
    Const S440_I1_ext3_amt = 63
    Const S440_I1_ext1_cd = 64
    Const S440_I1_ext2_cd = 65
    Const S440_I1_ext3_cd = 66

    Const S440_E1_cc_no = 0

    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status        
    
	ReDim I1_s_cc_hdr(S440_I1_ext3_cd)
	ReDim E1_s_cc_hdr(S440_E1_cc_no)

	'Tab 1		
	I1_s_cc_hdr(S440_I1_cc_no)   = UCase(Trim(Request("txtCCNo1")))		
	I1_s_cc_hdr(S440_I1_iv_no)   = UCase(Trim(Request("txtIVNo")))			
	I5_s_lc_hdr_lc_no			 = UCase(Trim(Request("txtLCNo")))			
	I1_s_cc_hdr(S440_I1_ed_no)   = UCase(Trim(Request("txtEDNo")))			
			
	If Trim(Request("txtSONoFlg")) = "Y" Then
		I6_s_so_hdr_so_no = UCase(Trim(Request("txtSONo")))	
	Else 
		I6_s_so_hdr_so_no = ""
	End If		

	I1_s_cc_hdr(S440_I1_iv_dt) = UNIConvDate(Request("txtIVDt"))	
	I1_s_cc_hdr(S440_I1_ed_type) = UCase(Trim(Request("txtEDType")))
	I1_s_cc_hdr(S440_I1_ed_dt) = UNIConvDate(Request("txtEDDt"))	
	I1_s_cc_hdr(S440_I1_weight_unit) = UCase(Trim(Request("txtWeightUnit")))
	I1_s_cc_hdr(S440_I1_ship_fin_dt) = UNIConvDate(Request("txtShipFinDt"))		
	
	if Len(Trim(Request("txtGrossWeight"))) then
		I1_s_cc_hdr(S440_I1_gross_weight) = UNIConvNum(Request("txtGrossWeight"),0)
	end if

	I1_s_cc_hdr(S440_I1_vessel_nm) = UCase(Trim(Request("txtVesselNm")))		

	if Len(Trim(Request("txtTotPackingCnt"))) then
		I1_s_cc_hdr(S440_I1_tot_packing_cnt) = UNIConvNum(Request("txtTotPackingCnt"),0)
	end if
	
	I1_s_cc_hdr(S440_I1_cur) = UCase(Trim(Request("txtCurrency")))		

	if Len(Trim(Request("txtXchRate"))) then
		I1_s_cc_hdr(S440_I1_xch_rate) = UNIConvNum(Request("txtXchRate"),0)
	end if

	I2_b_biz_partner_bp_cd = UCase(Trim(Request("txtApplicant")))		
	I3_b_biz_partner_bp_cd = UCase(Trim(Request("txtBeneficiary")))		
	I1_s_cc_hdr(S440_I1_agent) = UCase(Trim(Request("txtAgent")))		
	I1_s_cc_hdr(S440_I1_manufacturer) = UCase(Trim(Request("txtManufacturer")))					
		
	'Tab 2
	if Len(Trim(Request("txtUSDXchRate"))) then
		I1_s_cc_hdr(S440_I1_usd_xch_rate) = UNIConvNum(Request("txtUSDXchRate"),0)
	end if

	I1_s_cc_hdr(S440_I1_return_appl)    = UCase(Trim(Request("txtReturnAppl")))		
	I1_s_cc_hdr(S440_I1_return_office)  = UCase(Trim(Request("txtReturnOffice")))		
	I1_s_cc_hdr(S440_I1_loading_port)   = UCase(Trim(Request("txtLoadingPort")))		
	I1_s_cc_hdr(S440_I1_loading_cntry)  = UCase(Trim(Request("txtLoadingCntry")))		
	I1_s_cc_hdr(S440_I1_dischge_port)   = UCase(Trim(Request("txtDischgePort")))		
	I1_s_cc_hdr(S440_I1_dischge_cntry)  = UCase(Trim(Request("txtDischgeCntry")))		
	I1_s_cc_hdr(S440_I1_origin)			= UCase(Trim(Request("txtOrigin")))		
	I1_s_cc_hdr(S440_I1_origin_cntry)   = UCase(Trim(Request("txtOriginCntry")))		
	I1_s_cc_hdr(S440_I1_final_dest)     = UCase(Trim(Request("txtFinalDest")))		
	I1_s_cc_hdr(S440_I1_reporter)		= UCase(Trim(Request("txtReporter")))		
	I1_s_cc_hdr(S440_I1_pay_meth)		= UCase(Trim(Request("txtPayTerms")))		

	if Len(Trim(Request("txtPayDur"))) then
		I1_s_cc_hdr(S440_I1_pay_dur) = UNIConvNum(Request("txtPayDur"),0)
	end if

	I1_s_cc_hdr(S440_I1_incoterms)		= UCase(Trim(Request("txtIncoterms")))		
	I4_b_sales_grp_sales_grp		    = UCase(Trim(Request("txtSalesGroup")))		

	'Tab 3
	I1_s_cc_hdr(S440_I1_ep_no)			= UCase(Trim(Request("txtEPNo")))			
	
	if Len(Trim(Request("txtEPDt"))) then
		strConvDt = UNIConvDate(Request("txtEPDt"))		
		I1_s_cc_hdr(S440_I1_ep_dt) = strConvDt
	end if
	
	I1_s_cc_hdr(S440_I1_customs)		= UCase(Trim(Request("txtCustoms")))		
	I1_s_cc_hdr(S440_I1_trans_form)		= UCase(Trim(Request("txtTransForm")))		
	I1_s_cc_hdr(S440_I1_packing_type)   = UCase(Trim(Request("txtPackingType")))		
	I1_s_cc_hdr(S440_I1_trans_rep_cd)   = UCase(Trim(Request("txtTransRepCd")))		
	I1_s_cc_hdr(S440_I1_trans_method)	= UCase(Trim(Request("txtTransMeth")))		

	if Len(Trim(Request("txtTransFromDt"))) then
		strConvDt = UNIConvDate(Request("txtTransFromDt"))		
		I1_s_cc_hdr(S440_I1_trans_from_dt) = strConvDt
	end if

	if Len(Trim(Request("txtTransToDt"))) then
		strConvDt = UNIConvDate(Request("txtTransToDt"))		
		I1_s_cc_hdr(S440_I1_trans_to_dt) = strConvDt
	end if

	I1_s_cc_hdr(S440_I1_insp_cert_no)		= UCase(Trim(Request("txtInspCertNo")))		

	if Len(Trim(Request("txtInspCertDt"))) then
		strConvDt = UNIConvDate(Request("txtInspCertDt"))		
		I1_s_cc_hdr(S440_I1_insp_cert_dt) = strConvDt
	end if
	
	I1_s_cc_hdr(S440_I1_quar_cert_no)		= UCase(Trim(Request("txtQuarCertNo")))			
	
	if Len(Trim(Request("txtQuarCertDt"))) then
		strConvDt = UNIConvDate(Request("txtQuarCertDt"))		
		I1_s_cc_hdr(S440_I1_quar_cert_dt) = strConvDt
	end if

	I1_s_cc_hdr(S440_I1_device_plce)	= UCase(Trim(Request("txtDevicePlce")))		
	I1_s_cc_hdr(S440_I1_remark1)		= UCase(Trim(Request("txtRemark1")))		
	I1_s_cc_hdr(S440_I1_remark2)		= UCase(Trim(Request("txtRemark2")))		
	I1_s_cc_hdr(S440_I1_remark3)		= UCase(Trim(Request("txtRemark3")))		
	
	I7_s_wks_user_user_id				= UCase(Trim(Request("txtInsrtUserId")))
	I1_s_cc_hdr(S440_I1_ref_flag)		= UCase(Trim(Request("txtRefFlg")))		
	

	itxtFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

    If itxtFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf itxtFlgMode = OPMD_UMODE Then
    	iCommandSent = "UPDATE"
    End If

    Set pS6G211 = Server.CreateObject("PS6G211.cSExportCcHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Set pS6G211 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

    E1_s_cc_hdr(S440_E1_cc_no) =  pS6G211.S_MAINT_EXPORT_CC_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_cc_hdr, _
											I2_b_biz_partner_bp_cd, I3_b_biz_partner_bp_cd, I4_b_sales_grp_sales_grp, _
											I5_s_lc_hdr_lc_no, I6_s_so_hdr_so_no, I7_s_wks_user_user_id)
    
	If CheckSYSTEMError(Err,True) = True Then
       Set pS6G211 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set pS6G211 = Nothing	

 			
	Response.Write "<Script language=vbs> " & vbCr         
	Response.Write "With Parent "               & vbCr
    If itxtFlgMode = OPMD_CMODE Then
		Response.Write "   .frm1.txtCCNo.value = """   & ConvSPChars(E1_s_cc_hdr(S440_E1_cc_no))    & """" & vbCr 
	End If
    Response.Write " .DbSaveOk " & vbCr
    Response.Write "End With"     & vbCr      
    Response.Write "</Script> "  
    
End Sub
'============================================================================================================
Sub SubBizDelete()
    Dim pS6G211
    Dim iCommandSent
    Dim itxtFlgMode
	
	Dim I1_s_cc_hdr
	Dim I2_b_biz_partner_bp_cd
	Dim I3_b_biz_partner_bp_cd
	Dim I4_b_sales_grp_sales_grp
	Dim I5_s_lc_hdr_lc_no
	Dim I6_s_so_hdr_so_no
	Dim I7_s_wks_user_user_id	
	Dim E1_s_cc_hdr

    Const S440_I1_cc_no = 0
                
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status
   
    ReDim I1_s_cc_hdr(S440_I1_cc_no)

    If Request("txtCCNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	    Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If

	I1_s_cc_hdr(S440_I1_cc_no) = Trim(Request("txtCCNo"))

    iCommandSent = "DELETE"
    
    Set pS6G211 = Server.CreateObject("PS6G211.cSExportCcHdrSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
               

	call pS6G211.S_MAINT_EXPORT_CC_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_cc_hdr, _
											I2_b_biz_partner_bp_cd, I3_b_biz_partner_bp_cd, I4_b_sales_grp_sales_grp, _
											I5_s_lc_hdr_lc_no, I6_s_so_hdr_so_no, I7_s_wks_user_user_id)
    
    
    If CheckSYSTEMError(Err,True) = True Then
		Set pS6G211 = Nothing
		Exit Sub
	End If     
	Set pS6G211 = Nothing
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbDeleteOk "    & vbCr   
    Response.Write "</Script> "  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
Sub SubBizQueryMulti()
  
End Sub    
'============================================================================================================
Sub SubBizSaveMulti()        
    
End Sub    
'============================================================================================================
Sub SetErrorStatus()

End Sub
'============================================================================================================
Sub CommonOnTransactionCommit()

End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()

End Sub

%>

