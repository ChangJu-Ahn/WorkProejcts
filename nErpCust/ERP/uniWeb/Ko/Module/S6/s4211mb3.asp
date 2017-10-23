<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4211mb3.asp																*
'*  4. Program Name         : 데이터가져오기(통관등록 L/C참조에서)										*
'*  5. Program Desc         : 데이터가져오기(통관등록 L/C참조에서)										*
'*  7. Modified date(First) : 2000/04/10																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kim Hyungsuk																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/10 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
    Dim lgOpModeCRUD

    On Error Resume Next                                                             
    Err.Clear                                                                        

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB") 
    Call HideStatusWnd                                                               
    lgOpModeCRUD      = Request("txtMode")                                           
    Call SubBizQuery()
'============================================================================================================
Sub SubBizQuery()

    Dim iS4G119
	Dim iCommandSent
    Dim I1_s_lc_hdr
	Dim E1_b_biz_partner
	Dim E2_b_bank
	Dim E3_b_bank
	Dim E4_b_bank
	Dim E5_b_bank
	Dim E6_b_bank
	Dim E7_b_sales_grp
	Dim E8_b_sales_org
	Dim E9_b_biz_partner
	Dim E10_b_biz_partner
	Dim E11_b_biz_partner
	Dim E12_b_biz_partner
	Dim E13_b_biz_partner
	Dim E14_b_minor
	Dim E15_b_minor
	Dim E16_b_minor
	Dim E17_b_minor
	Dim E18_b_minor
	Dim E19_b_minor
	Dim E20_b_minor
	Dim E21_b_country
	Dim E22_b_minor
	Dim E23_b_minor
	Dim E24_b_minor
	Dim E25_b_minor
	Dim E26_s_lc_hdr
	
	Const S357_I1_lc_no = 0    
	Const S357_I1_lc_kind = 1
	
	Const S357_E1_bp_nm = 0
	Const S357_E2_bank_cd = 0
	Const S357_E2_bank_nm = 1
	Const S357_E3_bank_cd = 0
	Const S357_E3_bank_nm = 1
	Const S357_E4_bank_cd = 0
	Const S357_E4_bank_nm = 1
	Const S357_E5_bank_cd = 0
	Const S357_E5_bank_nm = 1
	Const S357_E6_bank_cd = 0
	Const S357_E6_bank_nm = 1
	Const S357_E7_sales_grp_nm = 0
	Const S357_E7_sales_grp = 1
	Const S357_E8_sales_org_nm = 0
	Const S357_E8_sales_org = 1
	Const S357_E9_bp_nm = 0
	Const S357_E9_bp_cd = 1
	Const S357_E10_bp_nm = 0
	Const S357_E10_bp_cd = 1
	Const S357_E11_bp_nm = 0
	Const S357_E12_bp_nm = 0
	Const S357_E13_bp_nm = 0
	Const S357_E14_minor_nm = 0
	Const S357_E15_minor_nm = 0
	Const S357_E16_minor_nm = 0
	Const S357_E17_minor_nm = 0
	Const S357_E18_minor_nm = 0
	Const S357_E19_minor_nm = 0
	Const S357_E20_minor_nm = 0
	Const S357_E21_country_nm = 0
	Const S357_E22_minor_nm = 0  
	Const S357_E23_minor_nm = 0  
	Const S357_E24_minor_nm = 0  
	Const S357_E25_minor_nm = 0  
	
	Const S357_E26_lc_no = 0    
	Const S357_E26_lc_doc_no = 1
	Const S357_E26_lc_amend_seq = 2
	Const S357_E26_so_no = 3
	Const S357_E26_adv_no = 4
	Const S357_E26_pre_adv_ref = 5
	Const S357_E26_adv_dt = 6
	Const S357_E26_open_dt = 7
	Const S357_E26_expiry_dt = 8
	Const S357_E26_amend_dt = 9
	Const S357_E26_manufacturer = 10
	Const S357_E26_agent = 11
	Const S357_E26_cur = 12
	Const S357_E26_lc_amt = 13
	Const S357_E26_xch_rate = 14
	Const S357_E26_lc_loc_amt = 15
	Const S357_E26_bank_txt = 16
	Const S357_E26_incoterms = 17
	Const S357_E26_pay_meth = 18
	Const S357_E26_payment_txt = 19
	Const S357_E26_latest_ship_dt = 20
	Const S357_E26_shipment = 21
	Const S357_E26_doc1 = 22
	Const S357_E26_doc2 = 23
	Const S357_E26_doc3 = 24
	Const S357_E26_doc4 = 25
	Const S357_E26_doc5 = 26
	Const S357_E26_file_dt = 27
	Const S357_E26_file_dt_txt = 28
	Const S357_E26_remark = 29
	Const S357_E26_lc_kind = 30
	Const S357_E26_lc_type = 31
	Const S357_E26_delivery_plce = 32
	Const S357_E26_amt_tolerance = 33
	Const S357_E26_loading_port = 34
	Const S357_E26_dischge_port = 35
	Const S357_E26_transport = 36
	Const S357_E26_transport_comp = 37
	Const S357_E26_origin = 38
	Const S357_E26_origin_cntry = 39
	Const S357_E26_charge_txt = 40
	Const S357_E26_charge_cd = 41
	Const S357_E26_credit_core = 42
	Const S357_E26_inv_cnt = 43
	Const S357_E26_bl_awb_flg = 44
	Const S357_E26_freight = 45
	Const S357_E26_notify_party = 46
	Const S357_E26_consignee = 47
	Const S357_E26_insur_policy = 48
	Const S357_E26_pack_list = 49
	Const S357_E26_l_lc_type = 50
	Const S357_E26_open_bank_txt = 51
	Const S357_E26_o_lc_doc_no = 52
	Const S357_E26_o_lc_amend_seq = 53
	Const S357_E26_o_lc_no = 54
	Const S357_E26_o_lc_expiry_dt = 55
	Const S357_E26_o_lc_loc_amt = 56
	Const S357_E26_o_lc_type = 57
	Const S357_E26_pay_dur = 58
	Const S357_E26_partial_ship_flag = 59
	Const S357_E26_biz_area = 60
	Const S357_E26_trnshp_flag = 61
	Const S357_E26_transfer_flag = 62
	Const S357_E26_cert_origin_flag = 63
	Const S357_E26_o_lc_amd_seq = 64
	Const S357_E26_sts = 65
	Const S357_E26_nego_amt = 66
	Const S357_E26_ext1_qty = 67
	Const S357_E26_ext2_qty = 68
	Const S357_E26_ext3_qty = 69
	Const S357_E26_ext1_amt = 70
	Const S357_E26_ext2_amt = 71
	Const S357_E26_ext3_qmt = 72
	Const S357_E26_ext1_cd = 73
	Const S357_E26_ext2_cd = 74
	Const S357_E26_ext3_cd = 75
	Const S357_E26_xch_rate_op = 76    
    
	On Error Resume Next
	Err.Clear                                                               

    If Request("txtLCNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If

	ReDim I1_s_lc_hdr(1)		
	iCommandSent = "LOOKUP"
	'Update 2005-01-07 LSW
	'I1_s_lc_hdr(S357_I1_lc_no) =  FilterVar(Trim(Request("txtLCNo")), "" , "SNM")
	I1_s_lc_hdr(S357_I1_lc_no) = Trim(Request("txtLCNo"))
    I1_s_lc_hdr(S357_I1_lc_kind) = Request("txtLCKind")

    Set iS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Set iS4G119 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
	End If
	
	Call iS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection,iCommandSent,I1_s_lc_hdr, _
	E1_b_biz_partner,E2_b_bank,E3_b_bank,E4_b_bank,E5_b_bank,E6_b_bank, _
	E7_b_sales_grp,E8_b_sales_org, _
	E9_b_biz_partner,E10_b_biz_partner,E11_b_biz_partner,E12_b_biz_partner,E13_b_biz_partner, _
	E14_b_minor,E15_b_minor,E16_b_minor,E17_b_minor,E18_b_minor,E19_b_minor,E20_b_minor,E21_b_country, _
    E22_b_minor,E23_b_minor,E24_b_minor,E25_b_minor,E26_s_lc_hdr )

	If CheckSYSTEMError(Err, True) = True Then
		Set iS4G119 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
	End If
	
	Set iS4G119 = Nothing
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr

	Response.write ".txtSONo.value				= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_so_no))         & """" & vbCr
	Response.write ".txtLCNo.value				= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_lc_no))         & """" & vbCr	
	Response.write ".txtLCDocNo.value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_lc_doc_no))     & """" & vbCr

	If E26_s_lc_hdr(S357_E26_lc_doc_no) = "" And  E26_s_lc_hdr(S357_E26_lc_amend_seq) = 0 Then
		Response.write ".txtLCAmendSeq.value		= """" "   & vbCr
	Else
		Response.write ".txtLCAmendSeq.value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_lc_amend_seq))        & """" & vbCr
	End If

	Response.Write ".txtCurrency.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_cur))				& """" & vbCr
	Response.Write ".txtCCCurrency.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_cur))				& """" & vbCr
	Response.Write ".txtFobCurrency.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_cur))				& """" & vbCr
	Response.write ".txtXchrate.Value			= """ & UNINumClientFormat(E26_s_lc_hdr(S357_E26_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write ".txtApplicant.Value			= """ & ConvSPChars(E10_b_biz_partner(S357_E10_bp_cd))		& """" & vbCr
	Response.Write ".txtApplicantNm.Value		= """ & ConvSPChars(E10_b_biz_partner(S357_E10_bp_nm))		& """" & vbCr
	Response.Write ".txtBeneficiary.Value		= """ & ConvSPChars(E9_b_biz_partner(S357_E9_bp_cd))		& """" & vbCr
	Response.Write ".txtBeneficiaryNm.Value		= """ & ConvSPChars(E9_b_biz_partner(S357_E9_bp_nm))		& """" & vbCr
	Response.Write ".txtManufacturer.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_manufacturer))	& """" & vbCr
	Response.Write ".txtManufacturerNm.Value	= """ & ConvSPChars(E12_b_biz_partner(S357_E12_bp_nm))		& """" & vbCr
	Response.Write ".txtAgent.Value				= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_agent))			& """" & vbCr
	Response.Write ".txtAgentNm.Value			= """ & ConvSPChars(E11_b_biz_partner(S357_E11_bp_nm))		& """" & vbCr
	Response.Write ".txtPayTerms.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_pay_meth))		& """" & vbCr
	Response.Write ".txtPayTermsNm.Value		= """ & ConvSPChars(E15_b_minor(S357_E15_minor_nm))			& """" & vbCr
	Response.Write ".txtPayDur.text				= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_pay_dur))			& """" & vbCr
	Response.Write ".txtIncoterms.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_incoterms))		& """" & vbCr
	Response.Write ".txtIncotermsNm.Value		= """ & ConvSPChars(E14_b_minor(S357_E14_minor_nm))			& """" & vbCr
	Response.Write ".txtSalesGroup.Value		= """ & ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp))		& """" & vbCr
	Response.Write ".txtSalesGroupNm.Value		= """ & ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp_nm))   & """" & vbCr
	Response.Write ".txtLoadingPort.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_loading_port))	& """" & vbCr
	Response.Write ".txtLoadingPortNm.Value		= """ & ConvSPChars(E17_b_minor(S357_E17_minor_nm))			& """" & vbCr
	Response.Write ".txtDischgePort.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_dischge_port))	& """" & vbCr
	Response.Write ".txtDischgePortNm.Value		= """ & ConvSPChars(E18_b_minor(S357_E18_minor_nm))			& """" & vbCr
	Response.Write ".txtOrigin.Value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_origin))			& """" & vbCr
	Response.Write ".txtOriginNm.Value			= """ & ConvSPChars(E20_b_minor(S357_E20_minor_nm))			& """" & vbCr
	Response.Write ".txtOriginCntry.Value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_origin_cntry))	& """" & vbCr
	Response.Write ".txtOriginCntryNm.Value		= """ & ConvSPChars(E21_b_country(S357_E21_country_nm))		& """" & vbCr
	Response.Write ".txtRefFlg.Value			= ""L"""												& vbCr
	

	Response.Write "parent.ReferenceQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
End Sub

%>

