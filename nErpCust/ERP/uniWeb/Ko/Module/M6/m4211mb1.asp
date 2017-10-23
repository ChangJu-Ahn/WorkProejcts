<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
		
	Dim lgOpModeCRUD
	Dim lgCurrency
	
	On Error Resume Next
	Err.Clear 
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
	
	Call HideStatusWnd
	
	lgOpModeCRUD = Request("txtMode")													'☜: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)                                                         '☜: Query
			    Call SubBizQuery()
	        Case Else
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    
	Dim OBJ_PM6G119																' Master L/C Header 조회용 Object

	Dim I1_m_cc_hdr 
    Dim E1_m_cc_hdr 
        
	Const M418_E1_cc_no = 0    
	Const M418_E1_id_no = 1
	Const M418_E1_ip_no = 2
	Const M418_E1_po_no = 3
	Const M418_E1_bl_doc_no = 4
	Const M418_E1_il_doc_no = 5
	Const M418_E1_reporter = 6
	Const M418_E1_tax_payer = 7
	Const M418_E1_manufacturer = 8
	Const M418_E1_agent = 9
	Const M418_E1_customs = 10
	Const M418_E1_cc_type = 11
	Const M418_E1_id_type = 12
	Const M418_E1_id_dt = 13
	Const M418_E1_id_req_dt = 14
	Const M418_E1_ip_type = 15
	Const M418_E1_import_type = 16
	Const M418_E1_pay_method = 17
	Const M418_E1_pay_dur = 18
	Const M418_E1_dischge_dt = 19
	Const M418_E1_tranship_dt = 20
	Const M418_E1_loading_dt = 21
	Const M418_E1_loading_port = 22
	Const M418_E1_dischge_port = 23
	Const M418_E1_vessel_nm = 24
	Const M418_E1_vessel_cntry = 25
	Const M418_E1_transport = 26
	Const M418_E1_loading_cntry = 27
	Const M418_E1_tranship_cntry = 28
	Const M418_E1_input_no = 29
	Const M418_E1_input_dt = 30
	Const M418_E1_exam_txt = 31
	Const M418_E1_device_no = 32
	Const M418_E1_device_plce = 33
	Const M418_E1_origin = 34
	Const M418_E1_origin_cntry = 35
	Const M418_E1_collect_type = 36
	Const M418_E1_incoterms = 37
	Const M418_E1_currency = 38
	Const M418_E1_doc_amt = 39
	Const M418_E1_cif_doc_amt = 40
	Const M418_E1_xch_rate = 41
	Const M418_E1_loc_amt = 42
	Const M418_E1_cif_loc_amt = 43
	Const M418_E1_gross_weight = 44
	Const M418_E1_weight_unit = 45
	Const M418_E1_packing_cnt = 46
	Const M418_E1_packing_type = 47
	Const M418_E1_packing_mark = 48
	Const M418_E1_packing_no = 49
	Const M418_E1_il_app_dt = 50
	Const M418_E1_ip_dt = 51
	Const M418_E1_doc = 52
	Const M418_E1_remark1 = 53
	Const M418_E1_remark2 = 54
	Const M418_E1_remark3 = 55
	Const M418_E1_cif_chg_rate = 56
	Const M418_E1_tariff_rate = 57
	Const M418_E1_exam_dt = 58
	Const M418_E1_output_dt = 59
	Const M418_E1_customs_exp_dt = 60
	Const M418_E1_payment_no = 61
	Const M418_E1_payment_dt = 62
	Const M418_E1_dvry_dt = 63
	Const M418_E1_taxbill_no = 64
	Const M418_E1_taxbill_dt = 65
	Const M418_E1_usd_xch_rate = 66
	Const M418_E1_insrt_user_id = 67
	Const M418_E1_insrt_dt = 68
	Const M418_E1_updt_user_id = 69
	Const M418_E1_updt_dt = 70
	Const M418_E1_ext1_qty = 71
	Const M418_E1_ext1_amt = 72
	Const M418_E1_ext1_cd = 73
	Const M418_E1_net_weight = 74
	Const M418_E1_biz_area = 75
	Const M418_E1_lc_no = 76
	Const M418_E1_lc_doc_no = 77
	Const M418_E1_lc_amend_seq = 78
	Const M418_E1_lc_open_dt = 79
	Const M418_E1_lc_type = 80
	Const M418_E1_open_bank = 81
	Const M418_E1_charge_flg = 82
	Const M418_E1_vat_type = 83
	Const M418_E1_vat_rate = 84
	Const M418_E1_vat_doc_amt = 85
	Const M418_E1_vat_loc_amt = 86
	Const M418_E1_tax_rate = 87
	Const M418_E1_tariff_tax = 88
	Const M418_E1_ext2_qty = 89
	Const M418_E1_ext3_qty = 90
	Const M418_E1_ext2_amt = 91
	Const M418_E1_ext3_amt = 92
	Const M418_E1_ext2_cd = 93
	Const M418_E1_ext3_cd = 94
	Const M418_E1_ext1_rt = 95
	Const M418_E1_ext2_rt = 96
	Const M418_E1_ext3_rt = 97
	Const M418_E1_ext1_dt = 98
	Const M418_E1_ext2_dt = 99
	Const M418_E1_ext3_dt = 100

	Const M418_E1_pur_grp = 101
	Const M418_E2_pur_grp_nm = 102
	Const M418_E1_pur_org = 103
	Const M418_E3_pur_org_nm = 104
	Const M418_E1_beneficiary = 105
	Const M418_E1_beneficiary_nm = 106
	Const M418_E1_applicant = 107
	Const M418_E5_applicant_nm = 108
	Const M418_E4_manufacturer_nm = 109
	Const M418_E6_agent_nm = 110
	Const M418_E7_tax_payer_nm = 111
	Const M418_E8_reporter_nm = 112
	Const M418_E11_bank_nm = 113
	Const M418_E12_transport_nm = 114
	Const M418_E13_pay_method_nm = 115
	Const M418_E14_packing_type_nm = 116
	Const M418_E15_orgin_nm = 117
	Const M418_E16_loaing_port_nm = 118
	Const M418_E17_dischae_port_nm = 119
	Const M418_E18_customs_nm = 120
	Const M418_E19_cc_type_nm = 121
	Const M418_E20_id_type_nm = 122
	Const M418_E21_ip_type_nm = 123
	Const M418_E22_import_type_nm = 124
    Const M418_E23_collect_type_nm = 125
        
	On Error Resume Next														'☜: Protect system from crashing
	Err.Clear																	'☜: Protect system from crashing

	'---------------------------------- 통관 Header Data Query ----------------------------------
	If Request("txtCCNo") = "" Then									
		Call DisplayMsgBox("700112", vbInformation,	"", "",	I_MKSCRIPT)
		Exit Sub 
	End If
	
	Set OBJ_PM6G119 = Server.CreateObject("PM6G119.cMLkImportCcHdrS")
		
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

	I1_m_cc_hdr = UCase(Trim(Request("txtCCNo")))
													
	Call OBJ_PM6G119.M_LOOKUP_IMPORT_CC_HDR_SVR(gStrGlobalCollection, I1_m_cc_hdr, E1_m_cc_hdr)
        
	If CheckSYSTEMError(Err,True) = True Then
		Set OBJ_PM6G119 = Nothing
		Exit Sub
	End If
	
	Set OBJ_PM6G119 = Nothing														'☜: Unload Comproxy
	
	'-----------------------
	'Result data display area
	'-----------------------
	lgCurrency = ConvSPChars(E1_m_cc_hdr(M418_E1_currency))

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write "Dim strDt" & vbCr
	Response.Write "Dim strDefDate" & vbCr
	'========= TAB 1 (수입신고) ==========
	
	'##### Rounding Logic #####
	'항상 거래화폐가 우선 
	Response.Write ".txtCurrency.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_currency))  & """"          & vbCr
	Response.Write "Parent.CurFormatNumericOCX" & vbCr
	'##########################
	Response.Write "strDefDate = """ & UNIDateClientFormat("1900-01-01") & """" & vbCr
		
	Response.Write ".txtInPutCCNo.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_cc_no)) & """"              & vbCr		'통관관리번호 
	Response.Write ".txtIDNo.value              = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_id_no)) & """"              & vbCr		'신고번호 
	Response.Write ".txtIDDt.Text               = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_id_dt)) & """"      & vbCr		'신고일 
	Response.Write ".txtIDReqDt.Text            = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_id_req_dt)) & """" & vbCr		'신고요청일 
	Response.Write ".txtIPNo.value              = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_ip_no)) & """"              & vbCr		'면허번호 
	
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_ip_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtIPDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write ".txtBLDocNo.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_bl_doc_no)) & """"          & vbCr		'B/L번호 
	Response.Write "If """ & ConvSPChars(E1_m_cc_hdr(M418_E1_bl_doc_no)) & """  <> """" Then"                       & vbCr
	Response.Write "	.chkBLNo.checked        = True"             & vbCr	
	Response.Write "	.txtChkBLNo.value       = ""Y"""            & vbCr
	Response.Write "End If"                                         & vbCr	
	Response.Write ".txtCustoms.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_customs)) & """"             & vbCr		'세관 
	Response.Write ".txtCustomsNm.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E18_customs_nm)) & """"        & vbCr		'세관명 
	Response.Write ".txtDischgeDt.Text          = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_dischge_dt)) & """" & vbCr		'도착일 
	Response.Write ".txtInputNo.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_input_no)) & """"           & vbCr		'반입번호 
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_input_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtPutDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write ".txtCollectType.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_collect_type)) & """"       & vbCr		'징수형태 
	Response.Write ".txtCollectTypeNm.value     = """ & ConvSPChars(E1_m_cc_hdr(M418_E23_collect_type_nm)) & """"   & vbCr      '징수형태명 
	Response.Write ".txtReporterCd.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_reporter)) & """"           & vbCr		'신고자 
	Response.Write ".txtReporterNm.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E8_reporter_nm)) & """"        & vbCr		'신고자명 
	Response.Write ".txtTaxPayerCd.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_tax_payer)) & """"          & vbCr		'납세의무자 
	Response.Write ".txtTaxPayerNm.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E7_tax_payer_nm)) & """"        & vbCr		'납세의무자명 
	Response.Write ".txtCCtypeCd.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_cc_type)) & """"            & vbCr		'통관계획 
	Response.Write ".txtCCtypeNm.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E19_cc_type_nm)) & """"        & vbCr		'통관계획명 
	Response.Write ".txtIDType.value            = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_id_type)) & """"            & vbCr		'신고구분 
	Response.Write ".txtIDTypeNm.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E20_id_type_nm)) & """"        & vbCr		'신고구분명 
	Response.Write ".txtIPType.value            = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_ip_type)) & """"            & vbCr		'거래구분 
	Response.Write ".txtIPTypeNm.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E21_ip_type_nm)) & """"         & vbCr		'거레구분명 
	Response.Write ".txtImportType.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_import_type)) & """"        & vbCr		'수입종류 
	Response.Write ".txtImportTypeNm.value      = """ & ConvSPChars(E1_m_cc_hdr(M418_E22_import_type_nm)) & """"    & vbCr	    '수입종류명 
	Response.Write ".txtGrossWeight.Text        = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_gross_weight), ggQty.DecPoint, 0) & """" & vbCr'총중량 
	Response.Write "If """ & E1_m_cc_hdr(M418_E1_net_weight) & """ = 0 Then"                                        & vbCr		'순중량 
	Response.Write "	.txtNetWeight.Text      = """ & E1_m_cc_hdr(M418_E1_net_weight) & """"                      & vbCr
	Response.Write "Else"                       & vbCr
	Response.Write "	.txtNetWeight.Text      = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_net_weight), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "End If"                     & vbCr
	Response.Write ".txtWeightUnit.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_weight_unit)) & """"        & vbCr		'중량단위 
	Response.Write ".txtTotPackingCnt.text      = """ & E1_m_cc_hdr(M418_E1_packing_cnt) & """"                     & vbCr		'총포장갯수 
	Response.Write ".txtPackingType.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_packing_type)) & """"       & vbCr		'포장형태 
	Response.Write ".txtPackingTypeNm.value     = """ & ConvSPChars(E1_m_cc_hdr(M418_E14_packing_type_nm)) & """"   & vbCr      '포장형태명 
	Response.Write ".txtDischgePortCd.value     = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_dischge_port)) & """"       & vbCr		'도착항 
	Response.Write ".txtDischgePortNm.value     = """ & ConvSPChars(E1_m_cc_hdr(M418_E17_dischae_port_nm)) & """"   & vbCr      '도착항명 
	Response.Write ".txtTransport.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_transport)) & """"          & vbCr		'운송방법 
	Response.Write ".txtTransportNm.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E12_transport_nm)) & """"      & vbCr		'운송방법명 
	Response.Write ".txtPayTerms.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_pay_method)) & """"         & vbCr		'결제방법 
	Response.Write ".txtPayTermsNm.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E13_pay_method_nm)) & """"     & vbCr		'결제방법명 
	Response.Write ".txtPayDur.value            = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_pay_dur)) & """"            & vbCr		'결제기간 
	Response.Write ".txtIncoterms.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_incoterms)) & """"          & vbCr		'가격조건 
	Response.Write ".txtCurrency.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_currency)) & """"           & vbCr		'화폐단위 
	Response.Write "If """ & E1_m_cc_hdr(M418_E1_doc_amt) & """ = 0 Then "                      & vbCr							'통관금액 
	Response.Write "	.txtDocAmt.Text         = """ & E1_m_cc_hdr(M418_E1_doc_amt) & """"     & vbCr
	Response.Write "Else"                                                                       & vbCr
	Response.Write "	.txtDocAmt.text         = """ & UNIConvNumDBToCompanyByCurrency(E1_m_cc_hdr(M418_E1_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write "End If"                                                                     & vbCr		
	Response.Write ".txtXchRate.Text            = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_xch_rate), ggExchRate.DecPoint, 0) & """"    & vbCr	'환율 
	Response.Write ".txtLocAmt.Text             = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_loc_amt), ggAmtOfMoney.DecPoint, 0) & """"   & vbCr	'원화금액 
	Response.Write ".txtUSDXchRate.Text         = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_usd_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr	'USD환율 
	Response.Write "If """ & E1_m_cc_hdr(M418_E1_cif_doc_amt) & """ = 0 Then "                      & vbCr											'CIF금액 
	Response.Write "	.txtCIFDocAmt.Text      = """ & E1_m_cc_hdr(M418_E1_cif_doc_amt) & """"     & vbCr
	Response.Write "Else"                                                                           & vbCr
	Response.Write "	.txtCIFDocAmt.text      = """ & UNIConvNumDBToCompanyByCurrency(E1_m_cc_hdr(M418_E1_cif_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write "End If"                                                                         & vbCr
	Response.Write ".txtCIFLocAmt.Text          = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_cif_loc_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr  'CIF원화금액 
	Response.Write ".txtBeneficiary.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_beneficiary)) & """"       & vbCr	'수출자 
	Response.Write ".txtBeneficiaryNm.value     = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_beneficiary_nm)) & """"    & vbCr	'수출자명 
	Response.Write ".txtApplicant.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_applicant)) & """"         & vbCr	'수입자 
	Response.Write ".txtApplicantNm.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E5_applicant_nm)) & """"      & vbCr	'수입자명 

	'========= TAB 2 (수입신고 기타) ==========
	Response.Write ".txtVesselNm.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_vessel_nm)) & """"          & vbCr	'Vessel명 
	Response.Write ".txtVesselCntry.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_vessel_cntry)) & """"       & vbCr	'선박국적 
	Response.Write ".txtLoadingPort.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_loading_port)) & """"       & vbCr	'선적항 
	Response.Write ".txtLoadingPortNm.value     = """ & ConvSPChars(E1_m_cc_hdr(M418_E16_loaing_port_nm)) & """"    & vbCr	'선적항명 
	Response.Write ".txtLoadingCntry.value      = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_loading_cntry)) & """"      & vbCr	'적출국가 
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_loading_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtLoadingDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write ".txtDeviceNo.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_device_no)) & """"          & vbCr	'장치확인번호 
	Response.Write ".txtDevicePlce.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_device_plce)) & """"        & vbCr	'반입장소 
	Response.Write ".txtPackingNo.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_packing_no)) & """"         & vbCr	'포장번호 
	Response.Write ".txtExamTxt.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_exam_txt)) & """"           & vbCr	'조사란 
	Response.Write ".txtOrigin.value            = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_origin)) & """"             & vbCr	'원산지 
	Response.Write ".txtOriginNm.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E15_orgin_nm)) & """"          & vbCr	'원산지명 
	Response.Write ".txtOriginCntry.value       = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_origin_cntry)) & """"       & vbCr	'원산지국가 
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_exam_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtInspectDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_output_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtOutputDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write ".txtPaymentNo.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_payment_no)) & """"         & vbCr	'납부서번호 
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_customs_exp_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtCustomsExpDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_payment_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtPaymentDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_dvry_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtDvryDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write ".txtTaxBillNo.value         = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_taxbill_no)) & """"         & vbCr	'계산서번호 
		
	Response.Write "strDt = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_taxbill_dt)) & """" & vbCr
	Response.Write "If strDt <> strDefDate Then " & vbCr
			Response.Write ".txtTaxBillDt.text = strDt " & vbCr
	Response.Write "end if " & vbCr
		
	Response.Write ".txtTariffTax.Text          = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_tariff_tax), ggAmtOfMoney.DecPoint, 0) & """"    & vbCr	'관세 
	Response.Write ".txtTariffRate.Text         = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_tariff_rate), ggExchRate.DecPoint, 0) & """"     & vbCr	'관세율 
	Response.Write ".txtVatType.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_vat_type)) & """"										& vbCr	'VAT유형 
	Response.Write ".txtVatRate.Text            = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_vat_rate), ggExchRate.DecPoint, 0) & """"        & vbCr	'관세율 
	Response.Write ".txtVatAmt.Text             = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_vat_loc_amt), ggAmtOfMoney.DecPoint, 0) & """"   & vbCr	'VAT금액 
	Response.Write ".txtLCDocNo.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_lc_doc_no)) & """"          & vbCr	'L/C번호 
	Response.Write ".txtLCAmendSeq.value        = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_lc_amend_seq)) & """"       & vbCr	'L/C순번 
	Response.Write ".txtAgentCd.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_agent)) & """"              & vbCr	'대행자 
	Response.Write ".txtAgentNm.value           = """ & ConvSPChars(E1_m_cc_hdr(M418_E6_agent_nm)) & """"           & vbCr	'대행자명 
	Response.Write ".txtManufacturerCd.value    = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_manufacturer)) & """"       & vbCr	'제조자 
	Response.Write ".txtManufacturerNm.value    = """ & ConvSPChars(E1_m_cc_hdr(M418_E4_manufacturer_nm)) & """"    & vbCr	'제조자명 
	Response.Write ".txtPurGrp.value            = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_pur_grp)) & """"            & vbCr	'수입담당 
	Response.Write ".txtPurGrpNm.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E2_pur_grp_nm)) & """"         & vbCr	'수입담당명 
	Response.Write ".txtPurOrg.value            = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_pur_org)) & """"            & vbCr	'수입부서 
	Response.Write ".txtPurOrgNm.value          = """ & ConvSPChars(E1_m_cc_hdr(M418_E3_pur_org_nm)) & """"         & vbCr	'수입부서명 
			
	'-------- Hidden Value ---------
	Response.Write ".txtLcNo.value              = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_lc_no)) & """"              & vbCr	'LC관리번호 
	Response.Write ".txtLcType.value            = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_lc_type)) & """"            & vbCr	'LC유형 
	Response.Write ".txtLcOpenDt.value          = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_lc_open_dt)) & """" & vbCr	'LC개설일 
	Response.Write ".txtPoNo.value              = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_po_no)) & """"              & vbCr	'PO번호 

	Response.Write "Call parent.DbQueryOk()	" & vbCr												'☜: 조회가 성공 

	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr
	
End Sub

%>

