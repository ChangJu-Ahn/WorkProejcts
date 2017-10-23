<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%

	Dim lgOpModeCRUD

	On Error Resume Next															'☜: Protect system from crashing
	Err.Clear 																		'☜: Clear Error status
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
	
	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")											'☜: Read Operation Mode (CRUD)
	Select Case lgOpModeCRUD
	        Case CStr(UID_M0002)
	             Call SubBizSave()
		        
	End Select

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================

Sub SubBizSave()

	Dim OBJ_PM42111																' Master L/C Header Save용 Object
	Dim UtxtCCno
	Dim lgIntFlgMode
	Dim expMCcHdrCcNo
	Dim iCommandSent
		
    Dim I1_b_pur_grp
    Dim I2_b_biz_partner
    Dim I3_b_biz_partner
    Dim I4_s_wks_user
    Dim I5_m_cc_hdr

		
	'IMPORTS View 상수 
	Const M410_I5_cc_no = 0														'View Name : import m_cc_hdr
	Const M410_I5_id_no = 1
	Const M410_I5_reporter = 2
	Const M410_I5_tax_payer = 3
	Const M410_I5_manufacturer = 4
	Const M410_I5_agent = 5
	Const M410_I5_customs = 6
	Const M410_I5_cc_type = 7
	Const M410_I5_id_type = 8
	Const M410_I5_id_dt = 9
	Const M410_I5_id_req_dt = 10
	Const M410_I5_ip_type = 11
	Const M410_I5_import_type = 12
	Const M410_I5_pay_method = 13
	Const M410_I5_pay_dur = 14
	Const M410_I5_dischge_dt = 15
	Const M410_I5_tranship_dt = 16
	Const M410_I5_loading_dt = 17
	Const M410_I5_loading_port = 18
	Const M410_I5_dischge_port = 19
	Const M410_I5_vessel_nm = 20
	Const M410_I5_vessel_cntry = 21
	Const M410_I5_transport = 22
	Const M410_I5_loading_cntry = 23
	Const M410_I5_tranship_cntry = 24
	Const M410_I5_bl_doc_no = 25
	Const M410_I5_input_no = 26
	Const M410_I5_input_dt = 27
	Const M410_I5_exam_txt = 28
	Const M410_I5_device_no = 29
	Const M410_I5_device_plce = 30
	Const M410_I5_origin = 31
	Const M410_I5_origin_cntry = 32
	Const M410_I5_collect_type = 33
	Const M410_I5_incoterms = 34
	Const M410_I5_currency = 35
	Const M410_I5_doc_amt = 36
	Const M410_I5_cif_doc_amt = 37
	Const M410_I5_xch_rate = 38
	Const M410_I5_loc_amt = 39
	Const M410_I5_cif_loc_amt = 40
	Const M410_I5_tariff_rate = 41
	Const M410_I5_gross_weight = 42
	Const M410_I5_net_weight = 43
	Const M410_I5_weight_unit = 44
	Const M410_I5_packing_cnt = 45
	Const M410_I5_packing_type = 46
	Const M410_I5_packing_mark = 47
	Const M410_I5_packing_no = 48
	Const M410_I5_po_no = 49
	Const M410_I5_il_doc_no = 50
	Const M410_I5_il_app_dt = 51
	Const M410_I5_ip_no = 52
	Const M410_I5_ip_dt = 53
	Const M410_I5_doc = 54
	Const M410_I5_biz_area = 55
	Const M410_I5_cif_chg_rate = 56
	Const M410_I5_exam_dt = 57
	Const M410_I5_output_dt = 58
	Const M410_I5_customs_exp_dt = 59
	Const M410_I5_payment_no = 60
	Const M410_I5_payment_dt = 61
	Const M410_I5_dvry_dt = 62
	Const M410_I5_taxbill_no = 63
	Const M410_I5_taxbill_dt = 64
	Const M410_I5_usd_xch_rate = 65
	Const M410_I5_remark1 = 66
	Const M410_I5_remark2 = 67
	Const M410_I5_remark3 = 68
	Const M410_I5_lc_no = 69
	Const M410_I5_lc_doc_no = 70
	Const M410_I5_lc_amend_seq = 71
	Const M410_I5_lc_open_dt = 72
	Const M410_I5_lc_type = 73
	Const M410_I5_open_bank = 74
	Const M410_I5_vat_type = 75
	Const M410_I5_vat_rate = 76
	Const M410_I5_vat_doc_amt = 77
	Const M410_I5_vat_loc_amt = 78
	Const M410_I5_tax_rate = 79
	Const M410_I5_tariff_tax = 80
	Const M410_I5_charge_flg = 81
	Const M410_I5_ext1_qty = 82
	Const M410_I5_ext1_amt = 83
	Const M410_I5_ext1_cd = 84
	Const M410_I5_ext2_qty = 85
	Const M410_I5_ext3_qty = 86
	Const M410_I5_ext2_amt = 87
	Const M410_I5_ext3_amt = 88
	Const M410_I5_ext2_cd = 89
	Const M410_I5_ext3_cd = 90
	Const M410_I5_ext1_rt = 91
	Const M410_I5_ext2_rt = 92
	Const M410_I5_ext3_rt = 93
	Const M410_I5_ext1_dt = 94
	Const M410_I5_ext2_dt = 95
	Const M410_I5_ext3_dt = 96
	Const M410_I5_bl_no = 97
		
	Redim I5_m_cc_hdr(M410_I5_bl_no)
		
	On Error Resume Next
	Err.Clear																		'☜: Protect system from crashing

	lgIntFlgMode = CInt(Request("txtFlgMode"))							
		
	UtxtCCno = UCase(Trim(Request("txtCCNo")))

	'========= TAB 1 (수입신고) ==========
	'통관관리번호 
		
	'신고번호 
	I5_m_cc_hdr(M410_I5_id_no) = UCase(Trim(Request("txtIDNo")))
	'신고일 
	If Len(Trim(Request("txtIDDt"))) Then I5_m_cc_hdr(M410_I5_id_dt) = uniConvDate(Request("txtIDDt"))
	'신고요청일 
	If Len(Trim(Request("txtIDReqDt"))) Then I5_m_cc_hdr(M410_I5_id_req_dt) = uniConvDate(Request("txtIDReqDt"))
	'면허번호 
	I5_m_cc_hdr(M410_I5_ip_no) = UCase(Trim(Request("txtIPNo")))
	'면허일 
	If Len(Trim(Request("txtIPDt"))) Then I5_m_cc_hdr(M410_I5_ip_dt) = uniConvDate(Request("txtIPDt"))
	'B/L번호 
	If Request("txtChkBLNo") = "Y" Then I5_m_cc_hdr(M410_I5_bl_doc_no) = UCase(Trim(Request("txtBLDocNo")))
	If Request("txtChkBLNo") = "Y" Then I5_m_cc_hdr(M410_I5_bl_no) = UCase(Trim(Request("txtBlNo")))
		
				
	'세관 
	I5_m_cc_hdr(M410_I5_customs) = UCase(Trim(Request("txtCustoms")))
	'도착일 
	If Len(Trim(Request("txtDischgeDt"))) Then I5_m_cc_hdr(M410_I5_dischge_dt) = uniConvDate(Request("txtDischgeDt"))
	'반입번호 
	I5_m_cc_hdr(M410_I5_input_no) = UCase(Trim(Request("txtInputNo")))
	'반입일 
	If Len(Trim(Request("txtPutDt"))) Then I5_m_cc_hdr(M410_I5_input_dt) = uniConvDate(Request("txtPutDt"))
	'징수형태 
	I5_m_cc_hdr(M410_I5_collect_type) = UCase(Trim(Request("txtCollectType")))
	'신고자 
	I5_m_cc_hdr(M410_I5_reporter) = UCase(Trim(Request("txtReporterCd")))
	'납세의무자 
	I5_m_cc_hdr(M410_I5_tax_payer) = UCase(Trim(Request("txtTaxPayerCd")))
	'통관계획 
	I5_m_cc_hdr(M410_I5_cc_type) = UCase(Trim(Request("txtCCtypeCd")))
		
	'신고구분 
	I5_m_cc_hdr(M410_I5_id_type) = UCase(Trim(Request("txtIDType")))
	'거래구분 
	I5_m_cc_hdr(M410_I5_ip_type) = UCase(Trim(Request("txtIPType")))
	'수입종류 
	I5_m_cc_hdr(M410_I5_import_type) = UCase(Trim(Request("txtImportType")))
	'총중량 
	If Len(Trim(Request("txtGrossWeight"))) Then I5_m_cc_hdr(M410_I5_gross_weight) = UNIConvNum(Request("txtGrossWeight"),0)
	'순중량 
	If Len(Trim(Request("txtNetWeight"))) Then I5_m_cc_hdr(M410_I5_net_weight) = UNIConvNum(Request("txtNetWeight"),0)
	'중량단위 
	I5_m_cc_hdr(M410_I5_weight_unit) = UCase(Trim(Request("txtWeightUnit")))
	'총포장갯수 
	If Len(Trim(Request("txtTotPackingCnt"))) Then I5_m_cc_hdr(M410_I5_packing_cnt) = UNIConvNum(Request("txtTotPackingCnt"),0)
	'포장형태 
	I5_m_cc_hdr(M410_I5_packing_type) = UCase(Trim(Request("txtPackingType")))
	'도착항 
	I5_m_cc_hdr(M410_I5_dischge_port) = UCase(Trim(Request("txtDischgePortCd")))
	'운송방법 
	I5_m_cc_hdr(M410_I5_transport) = UCase(Trim(Request("txtTransport")))
	'결제방법 
	I5_m_cc_hdr(M410_I5_pay_method) = UCase(Trim(Request("txtPayTerms")))
	'결제기간 
	If Len(Trim(Request("txtPayDur"))) Then I5_m_cc_hdr(M410_I5_pay_dur) = UNIConvNum(Request("txtPayDur"),0)
	'가격조건 
	I5_m_cc_hdr(M410_I5_incoterms) = UCase(Trim(Request("txtIncoterms")))
	'화폐단위 
	I5_m_cc_hdr(M410_I5_currency) = UCase(Trim(Request("txtCurrency")))
	'통관금액 
	If Len(Trim(Request("txtDocAmt"))) Then 
		I5_m_cc_hdr(M410_I5_doc_amt) = UNIConvNum(Request("txtDocAmt"),0)
	Else
		I5_m_cc_hdr(M410_I5_doc_amt) = 0
	End If	
	'환율 
	If Len(Trim(Request("txtXchRate"))) Then I5_m_cc_hdr(M410_I5_xch_rate) = UNIConvNum(Request("txtXchRate"),0)
	'원화금액 
	If Len(Trim(Request("txtLocAmt"))) Then I5_m_cc_hdr(M410_I5_loc_amt) = UNIConvNum(Request("txtLocAmt"),0)

	'USD환율 
	If Len(Trim(Request("txtUSDXchRate"))) Then I5_m_cc_hdr(M410_I5_usd_xch_rate) = UNIConvNum(Request("txtUSDXchRate"),0)
	'CIF금액 
	If Len(Trim(Request("txtCIFDocAmt"))) Then I5_m_cc_hdr(M410_I5_cif_doc_amt) = UNIConvNum(Request("txtCIFDocAmt"),0)
	'CIF원화금액 
	If Len(Trim(Request("txtCIFLocAmt"))) Then I5_m_cc_hdr(M410_I5_cif_loc_amt) = UNIConvNum(Request("txtCIFLocAmt"),0)
	'수출자 
	I3_b_biz_partner = UCase(Trim(Request("txtBeneficiary")))
	'수입자 
	I2_b_biz_partner = UCase(Trim(Request("txtApplicant")))


	'========= TAB 2 (수입신고 기타) ==========
	'Vessel명 
	I5_m_cc_hdr(M410_I5_vessel_nm) = UCase(Trim(Request("txtVesselNm")))
	'선박국적 
	I5_m_cc_hdr(M410_I5_vessel_cntry) = UCase(Trim(Request("txtVesselCntry")))
	'선적항 
	I5_m_cc_hdr(M410_I5_loading_port) = UCase(Trim(Request("txtLoadingPort")))
	'적출국가 
	I5_m_cc_hdr(M410_I5_loading_cntry) = UCase(Trim(Request("txtLoadingCntry")))
	 '선적일 
	If Len(Trim(Request("txtLoadingDt"))) Then I5_m_cc_hdr(M410_I5_loading_dt) = uniConvDate(Request("txtLoadingDt"))
	 '장치확인번호 
	I5_m_cc_hdr(M410_I5_device_no) = UCase(Trim(Request("txtDeviceNo")))
	 '반입장소 
	I5_m_cc_hdr(M410_I5_device_plce) = UCase(Trim(Request("txtDevicePlce")))
	 '포장번호 
	I5_m_cc_hdr(M410_I5_packing_no) = UCase(Trim(Request("txtPackingNo")))
	 '조사란 
	I5_m_cc_hdr(M410_I5_exam_txt) = UCase(Trim(Request("txtExamTxt")))
	 '원산지 
	I5_m_cc_hdr(M410_I5_origin) = UCase(Trim(Request("txtOrigin")))
	 '원산지국가 
	I5_m_cc_hdr(M410_I5_origin_cntry) = UCase(Trim(Request("txtOriginCntry")))
	 '검사일 
	If Len(Trim(Request("txtInspectDt"))) Then I5_m_cc_hdr(M410_I5_exam_dt) = uniConvDate(Request("txtInspectDt"))
	 '반출일 
	If Len(Trim(Request("txtOutputDt"))) Then I5_m_cc_hdr(M410_I5_output_dt) = uniConvDate(Request("txtOutputDt"))		
	 '납부서번호 
	I5_m_cc_hdr(M410_I5_payment_no) = UCase(Trim(Request("txtPaymentNo")))
	 '세관만기일 
	If Len(Trim(Request("txtCustomsExpDt"))) Then I5_m_cc_hdr(M410_I5_customs_exp_dt) = uniConvDate(Request("txtCustomsExpDt"))
	 '납부일 
	If Len(Trim(Request("txtPaymentDt"))) Then I5_m_cc_hdr(M410_I5_payment_dt) = uniConvDate(Request("txtPaymentDt"))
	 '납기일 
	If Len(Trim(Request("txtDvryDt"))) Then I5_m_cc_hdr(M410_I5_dvry_dt) = uniConvDate(Request("txtDvryDt"))
	 '계산서번호 
	I5_m_cc_hdr(M410_I5_taxbill_no) = UCase(Trim(Request("txtTaxBillNo")))
	 '계산서발행일 
	If Len(Trim(Request("txtTaxBillDt"))) Then I5_m_cc_hdr(M410_I5_taxbill_dt) = uniConvDate(Request("txtTaxBillDt"))
	 '관세 
	If Len(Trim(Request("txtTariffTax"))) Then I5_m_cc_hdr(M410_I5_tariff_tax) = UNIConvNum(Request("txtTariffTax"),0)
	 '관세율 
	If Len(Trim(Request("txtTariffRate"))) Then I5_m_cc_hdr(M410_I5_tariff_rate) = UNIConvNum(Request("txtTariffRate"),0)
	 'VAT유형 
	I5_m_cc_hdr(M410_I5_vat_type) = UCase(Trim(Request("txtVatType")))
	 'VAT율 
	If Len(Trim(Request("txtVatRate"))) Then I5_m_cc_hdr(M410_I5_vat_rate) = UNIConvNum(Request("txtVatRate"),0)
	 'VAT금액 
	If Len(Trim(Request("txtVatAmt"))) Then I5_m_cc_hdr(M410_I5_vat_loc_amt) = UNIConvNum(Request("txtVatAmt"),0)
	 '가산금액 
	'If Len(Trim(Request("txtAddLocAmt"))) Then I5_m_cc_hdr()AddLocAmt = Trim(Request("txtAddLocAmt"))
	 '공제금액 
	'If Len(Trim(Request("txtReduLocAmt"))) Then I5_m_cc_hdr()ReduLocAmt = Trim(Request("txtReduLocAmt"))
	 'L/C번호 
	I5_m_cc_hdr(M410_I5_lc_doc_no) = UCase(Trim(Request("txtLCDocNo")))
	 'L/C순번 
	If Len(Trim(Request("txtLCAmendSeq"))) Then I5_m_cc_hdr(M410_I5_lc_amend_seq) = UNIConvNum(Request("txtLCAmendSeq"),0)
	 '대행자 
	I5_m_cc_hdr(M410_I5_agent) = UCase(Trim(Request("txtAgentCd")))
	 '제조자 
	I5_m_cc_hdr(M410_I5_manufacturer) = UCase(Trim(Request("txtManufacturerCd")))
	 '수입담당 
	I1_b_pur_grp = UCase(Trim(Request("txtPurGrp")))
		

	'-------- Hidden Value ---------
	'LC관리번호 
	I5_m_cc_hdr(M410_I5_lc_no) = UCase(Trim(Request("txtLcNo")))
	'LC유형 
	I5_m_cc_hdr(M410_I5_lc_type) = UCase(Trim(Request("txtLcType")))
	'LC개설일 
	If Len(Trim(Request("txtLcOpenDt"))) Then I5_m_cc_hdr(M410_I5_lc_open_dt) = uniConvDate(Request("txtLcOpenDt"))
	'PO번호 
	I5_m_cc_hdr(M410_I5_po_no) = UCase(Trim(Request("txtPoNo")))

	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
		I5_m_cc_hdr(M410_I5_cc_no) = UCase(Trim(Request("txtInPutCCNo")))

	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
		I5_m_cc_hdr(M410_I5_cc_no) = UCase(Trim(Request("txtCCNo")))

	End If

	Set OBJ_PM42111 = Server.CreateObject("PM6G111.cMMaintImportCcHdrS")

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

	expMCcHdrCcNo =  OBJ_PM42111.M_MAINT_IMPORT_CC_HDR_SVR(gStrGlobalCollection, iCommandSent, _
								I1_b_pur_grp, I2_b_biz_partner, I3_b_biz_partner, I5_m_cc_hdr)
                
	If CheckSYSTEMError2(Err, True,"","","","","") = True Then
	    Set OBJ_PM42111 = Nothing
		Exit Sub
	End If

	Set OBJ_PM42111 = Nothing														'☜: Unload Comproxy
 		
	'-----------------------
	'Result data display area
		'-----------------------
	Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent" & vbCr

			Response.Write "If """ & ConvSPChars(expMCcHdrCcNo) & """ = """"  Then" & vbCr
				Response.Write ".frm1.txtCCNo.value = """ & ConvSPChars(UtxtCCno) & """" & vbCr
			Response.Write "Else" & vbCr
				Response.Write ".frm1.txtCCNo.value = """ & ConvSPChars(expMCcHdrCcNo) & """" & vbCr
			Response.Write "End If" & vbCr

			Response.Write ".DbSaveOk" & vbCr
		Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr

	Exit sub
		
End Sub
	
%>
