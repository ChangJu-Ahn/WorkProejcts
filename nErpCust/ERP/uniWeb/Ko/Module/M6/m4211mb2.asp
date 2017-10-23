<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%

	Dim lgOpModeCRUD

	On Error Resume Next															'��: Protect system from crashing
	Err.Clear 																		'��: Clear Error status
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
	
	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")											'��: Read Operation Mode (CRUD)
	Select Case lgOpModeCRUD
	        Case CStr(UID_M0002)
	             Call SubBizSave()
		        
	End Select

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================

Sub SubBizSave()

	Dim OBJ_PM42111																' Master L/C Header Save�� Object
	Dim UtxtCCno
	Dim lgIntFlgMode
	Dim expMCcHdrCcNo
	Dim iCommandSent
		
    Dim I1_b_pur_grp
    Dim I2_b_biz_partner
    Dim I3_b_biz_partner
    Dim I4_s_wks_user
    Dim I5_m_cc_hdr

		
	'IMPORTS View ��� 
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
	Err.Clear																		'��: Protect system from crashing

	lgIntFlgMode = CInt(Request("txtFlgMode"))							
		
	UtxtCCno = UCase(Trim(Request("txtCCNo")))

	'========= TAB 1 (���ԽŰ�) ==========
	'���������ȣ 
		
	'�Ű��ȣ 
	I5_m_cc_hdr(M410_I5_id_no) = UCase(Trim(Request("txtIDNo")))
	'�Ű��� 
	If Len(Trim(Request("txtIDDt"))) Then I5_m_cc_hdr(M410_I5_id_dt) = uniConvDate(Request("txtIDDt"))
	'�Ű��û�� 
	If Len(Trim(Request("txtIDReqDt"))) Then I5_m_cc_hdr(M410_I5_id_req_dt) = uniConvDate(Request("txtIDReqDt"))
	'�����ȣ 
	I5_m_cc_hdr(M410_I5_ip_no) = UCase(Trim(Request("txtIPNo")))
	'������ 
	If Len(Trim(Request("txtIPDt"))) Then I5_m_cc_hdr(M410_I5_ip_dt) = uniConvDate(Request("txtIPDt"))
	'B/L��ȣ 
	If Request("txtChkBLNo") = "Y" Then I5_m_cc_hdr(M410_I5_bl_doc_no) = UCase(Trim(Request("txtBLDocNo")))
	If Request("txtChkBLNo") = "Y" Then I5_m_cc_hdr(M410_I5_bl_no) = UCase(Trim(Request("txtBlNo")))
		
				
	'���� 
	I5_m_cc_hdr(M410_I5_customs) = UCase(Trim(Request("txtCustoms")))
	'������ 
	If Len(Trim(Request("txtDischgeDt"))) Then I5_m_cc_hdr(M410_I5_dischge_dt) = uniConvDate(Request("txtDischgeDt"))
	'���Թ�ȣ 
	I5_m_cc_hdr(M410_I5_input_no) = UCase(Trim(Request("txtInputNo")))
	'������ 
	If Len(Trim(Request("txtPutDt"))) Then I5_m_cc_hdr(M410_I5_input_dt) = uniConvDate(Request("txtPutDt"))
	'¡������ 
	I5_m_cc_hdr(M410_I5_collect_type) = UCase(Trim(Request("txtCollectType")))
	'�Ű��� 
	I5_m_cc_hdr(M410_I5_reporter) = UCase(Trim(Request("txtReporterCd")))
	'�����ǹ��� 
	I5_m_cc_hdr(M410_I5_tax_payer) = UCase(Trim(Request("txtTaxPayerCd")))
	'�����ȹ 
	I5_m_cc_hdr(M410_I5_cc_type) = UCase(Trim(Request("txtCCtypeCd")))
		
	'�Ű��� 
	I5_m_cc_hdr(M410_I5_id_type) = UCase(Trim(Request("txtIDType")))
	'�ŷ����� 
	I5_m_cc_hdr(M410_I5_ip_type) = UCase(Trim(Request("txtIPType")))
	'�������� 
	I5_m_cc_hdr(M410_I5_import_type) = UCase(Trim(Request("txtImportType")))
	'���߷� 
	If Len(Trim(Request("txtGrossWeight"))) Then I5_m_cc_hdr(M410_I5_gross_weight) = UNIConvNum(Request("txtGrossWeight"),0)
	'���߷� 
	If Len(Trim(Request("txtNetWeight"))) Then I5_m_cc_hdr(M410_I5_net_weight) = UNIConvNum(Request("txtNetWeight"),0)
	'�߷����� 
	I5_m_cc_hdr(M410_I5_weight_unit) = UCase(Trim(Request("txtWeightUnit")))
	'�����尹�� 
	If Len(Trim(Request("txtTotPackingCnt"))) Then I5_m_cc_hdr(M410_I5_packing_cnt) = UNIConvNum(Request("txtTotPackingCnt"),0)
	'�������� 
	I5_m_cc_hdr(M410_I5_packing_type) = UCase(Trim(Request("txtPackingType")))
	'������ 
	I5_m_cc_hdr(M410_I5_dischge_port) = UCase(Trim(Request("txtDischgePortCd")))
	'��۹�� 
	I5_m_cc_hdr(M410_I5_transport) = UCase(Trim(Request("txtTransport")))
	'������� 
	I5_m_cc_hdr(M410_I5_pay_method) = UCase(Trim(Request("txtPayTerms")))
	'�����Ⱓ 
	If Len(Trim(Request("txtPayDur"))) Then I5_m_cc_hdr(M410_I5_pay_dur) = UNIConvNum(Request("txtPayDur"),0)
	'�������� 
	I5_m_cc_hdr(M410_I5_incoterms) = UCase(Trim(Request("txtIncoterms")))
	'ȭ����� 
	I5_m_cc_hdr(M410_I5_currency) = UCase(Trim(Request("txtCurrency")))
	'����ݾ� 
	If Len(Trim(Request("txtDocAmt"))) Then 
		I5_m_cc_hdr(M410_I5_doc_amt) = UNIConvNum(Request("txtDocAmt"),0)
	Else
		I5_m_cc_hdr(M410_I5_doc_amt) = 0
	End If	
	'ȯ�� 
	If Len(Trim(Request("txtXchRate"))) Then I5_m_cc_hdr(M410_I5_xch_rate) = UNIConvNum(Request("txtXchRate"),0)
	'��ȭ�ݾ� 
	If Len(Trim(Request("txtLocAmt"))) Then I5_m_cc_hdr(M410_I5_loc_amt) = UNIConvNum(Request("txtLocAmt"),0)

	'USDȯ�� 
	If Len(Trim(Request("txtUSDXchRate"))) Then I5_m_cc_hdr(M410_I5_usd_xch_rate) = UNIConvNum(Request("txtUSDXchRate"),0)
	'CIF�ݾ� 
	If Len(Trim(Request("txtCIFDocAmt"))) Then I5_m_cc_hdr(M410_I5_cif_doc_amt) = UNIConvNum(Request("txtCIFDocAmt"),0)
	'CIF��ȭ�ݾ� 
	If Len(Trim(Request("txtCIFLocAmt"))) Then I5_m_cc_hdr(M410_I5_cif_loc_amt) = UNIConvNum(Request("txtCIFLocAmt"),0)
	'������ 
	I3_b_biz_partner = UCase(Trim(Request("txtBeneficiary")))
	'������ 
	I2_b_biz_partner = UCase(Trim(Request("txtApplicant")))


	'========= TAB 2 (���ԽŰ� ��Ÿ) ==========
	'Vessel�� 
	I5_m_cc_hdr(M410_I5_vessel_nm) = UCase(Trim(Request("txtVesselNm")))
	'���ڱ��� 
	I5_m_cc_hdr(M410_I5_vessel_cntry) = UCase(Trim(Request("txtVesselCntry")))
	'������ 
	I5_m_cc_hdr(M410_I5_loading_port) = UCase(Trim(Request("txtLoadingPort")))
	'���ⱹ�� 
	I5_m_cc_hdr(M410_I5_loading_cntry) = UCase(Trim(Request("txtLoadingCntry")))
	 '������ 
	If Len(Trim(Request("txtLoadingDt"))) Then I5_m_cc_hdr(M410_I5_loading_dt) = uniConvDate(Request("txtLoadingDt"))
	 '��ġȮ�ι�ȣ 
	I5_m_cc_hdr(M410_I5_device_no) = UCase(Trim(Request("txtDeviceNo")))
	 '������� 
	I5_m_cc_hdr(M410_I5_device_plce) = UCase(Trim(Request("txtDevicePlce")))
	 '�����ȣ 
	I5_m_cc_hdr(M410_I5_packing_no) = UCase(Trim(Request("txtPackingNo")))
	 '����� 
	I5_m_cc_hdr(M410_I5_exam_txt) = UCase(Trim(Request("txtExamTxt")))
	 '������ 
	I5_m_cc_hdr(M410_I5_origin) = UCase(Trim(Request("txtOrigin")))
	 '���������� 
	I5_m_cc_hdr(M410_I5_origin_cntry) = UCase(Trim(Request("txtOriginCntry")))
	 '�˻��� 
	If Len(Trim(Request("txtInspectDt"))) Then I5_m_cc_hdr(M410_I5_exam_dt) = uniConvDate(Request("txtInspectDt"))
	 '������ 
	If Len(Trim(Request("txtOutputDt"))) Then I5_m_cc_hdr(M410_I5_output_dt) = uniConvDate(Request("txtOutputDt"))		
	 '���μ���ȣ 
	I5_m_cc_hdr(M410_I5_payment_no) = UCase(Trim(Request("txtPaymentNo")))
	 '���������� 
	If Len(Trim(Request("txtCustomsExpDt"))) Then I5_m_cc_hdr(M410_I5_customs_exp_dt) = uniConvDate(Request("txtCustomsExpDt"))
	 '������ 
	If Len(Trim(Request("txtPaymentDt"))) Then I5_m_cc_hdr(M410_I5_payment_dt) = uniConvDate(Request("txtPaymentDt"))
	 '������ 
	If Len(Trim(Request("txtDvryDt"))) Then I5_m_cc_hdr(M410_I5_dvry_dt) = uniConvDate(Request("txtDvryDt"))
	 '��꼭��ȣ 
	I5_m_cc_hdr(M410_I5_taxbill_no) = UCase(Trim(Request("txtTaxBillNo")))
	 '��꼭������ 
	If Len(Trim(Request("txtTaxBillDt"))) Then I5_m_cc_hdr(M410_I5_taxbill_dt) = uniConvDate(Request("txtTaxBillDt"))
	 '���� 
	If Len(Trim(Request("txtTariffTax"))) Then I5_m_cc_hdr(M410_I5_tariff_tax) = UNIConvNum(Request("txtTariffTax"),0)
	 '������ 
	If Len(Trim(Request("txtTariffRate"))) Then I5_m_cc_hdr(M410_I5_tariff_rate) = UNIConvNum(Request("txtTariffRate"),0)
	 'VAT���� 
	I5_m_cc_hdr(M410_I5_vat_type) = UCase(Trim(Request("txtVatType")))
	 'VAT�� 
	If Len(Trim(Request("txtVatRate"))) Then I5_m_cc_hdr(M410_I5_vat_rate) = UNIConvNum(Request("txtVatRate"),0)
	 'VAT�ݾ� 
	If Len(Trim(Request("txtVatAmt"))) Then I5_m_cc_hdr(M410_I5_vat_loc_amt) = UNIConvNum(Request("txtVatAmt"),0)
	 '����ݾ� 
	'If Len(Trim(Request("txtAddLocAmt"))) Then I5_m_cc_hdr()AddLocAmt = Trim(Request("txtAddLocAmt"))
	 '�����ݾ� 
	'If Len(Trim(Request("txtReduLocAmt"))) Then I5_m_cc_hdr()ReduLocAmt = Trim(Request("txtReduLocAmt"))
	 'L/C��ȣ 
	I5_m_cc_hdr(M410_I5_lc_doc_no) = UCase(Trim(Request("txtLCDocNo")))
	 'L/C���� 
	If Len(Trim(Request("txtLCAmendSeq"))) Then I5_m_cc_hdr(M410_I5_lc_amend_seq) = UNIConvNum(Request("txtLCAmendSeq"),0)
	 '������ 
	I5_m_cc_hdr(M410_I5_agent) = UCase(Trim(Request("txtAgentCd")))
	 '������ 
	I5_m_cc_hdr(M410_I5_manufacturer) = UCase(Trim(Request("txtManufacturerCd")))
	 '���Դ�� 
	I1_b_pur_grp = UCase(Trim(Request("txtPurGrp")))
		

	'-------- Hidden Value ---------
	'LC������ȣ 
	I5_m_cc_hdr(M410_I5_lc_no) = UCase(Trim(Request("txtLcNo")))
	'LC���� 
	I5_m_cc_hdr(M410_I5_lc_type) = UCase(Trim(Request("txtLcType")))
	'LC������ 
	If Len(Trim(Request("txtLcOpenDt"))) Then I5_m_cc_hdr(M410_I5_lc_open_dt) = uniConvDate(Request("txtLcOpenDt"))
	'PO��ȣ 
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

	Set OBJ_PM42111 = Nothing														'��: Unload Comproxy
 		
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
