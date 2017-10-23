<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211mb1.asp																*
'*  4. Program Name         : L/C 등록																	*
'*  5. Program Desc         : L/C 등록																	*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/07/12																*
'*  8. Modified date(Last)  : 2001/08/29																*
'*  9. Modifier (First)     : Kim Hyungsuk 																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/07/12 : Coding ReStart											*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")   
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
Call HideStatusWnd          

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim PB0C004
Dim PS4G119
Dim PS4G111

Dim I1_s_lc_hdr
ReDim I1_s_lc_hdr(1)

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

Dim txtPrevNext

Const B253_E1_std_rate = 0
const B253_E1_multi_divide = 1

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

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)								'☜: 현재 조회 요청을 받음 

Err.Clear	

	If Trim(Request("txtLCNo")) = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If	
	
	I1_s_lc_hdr(S357_I1_lc_no) = Trim(Request("txtLCNo"))
    I1_s_lc_hdr(S357_I1_lc_kind) = "M"
	
	txtPrevNext = Request("txtPrevNext")
	
    Set PS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")
		
	Call PS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection,txtPrevNext,I1_s_lc_hdr, _
	E1_b_biz_partner,E2_b_bank,E3_b_bank,E4_b_bank,E5_b_bank,E6_b_bank, _
	E7_b_sales_grp,E8_b_sales_org, _
	E9_b_biz_partner,E10_b_biz_partner,E11_b_biz_partner,E12_b_biz_partner,E13_b_biz_partner, _
	E14_b_minor,E15_b_minor,E16_b_minor,E17_b_minor,E18_b_minor,E19_b_minor,E20_b_minor,E21_b_country, _
    E22_b_minor,E23_b_minor,E24_b_minor,E25_b_minor,E26_s_lc_hdr )
    
    If CheckSYSTEMError(Err,True) = True Then
		Set PS4G119 = Nothing
%>
	<Script Language=VBScript>
		parent.frm1.txtLCNo.focus
	</Script>	
<%				
		Response.End
	End If  
	
	Set PS4G119 = Nothing  
	
%>

<Script Language=VBScript>
	With parent.frm1
		Dim strDt
		
		
	
		.txtCurrency.value			= "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
		parent.CurFormatNumericOCX
	
		' Tab 1 : L/C 금액정보 
		
		.txtLCNo.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_no))%>"
		.txtLCNo1.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_no))%>"
		.txtSONo.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_so_no))%>"
		'2002/8/1 --수주번호Check추가--
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_so_no))%>" <> "" Then
			.chkSONoFlg.checked = True
		Else 
			.chkSONoFlg.checked = False
		End If	
		.txtLCDocNo.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_doc_no))%>"
		.txtLCAmendSeq.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_amend_seq))%>"
		.txtAdvNo.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_adv_no))%>"
		.txtLCType.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_type))%>"
		.txtLCTypeNm.value = "<%=ConvSPChars(E16_b_minor(S357_E16_minor_nm))%>"
		
		.txtAdvDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_adv_dt))%>"
		
		.txtAdvBank.value = "<%=ConvSPChars(E3_b_bank(S357_E3_bank_cd))%>"
		.txtAdvBankNm.value = "<%=ConvSPChars(E3_b_bank(S357_E3_bank_nm))%>"
		
		.txtExpireDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_expiry_dt))%>"
		
		.txtOpenBank.value = "<%=ConvSPChars(E2_b_bank(S357_E2_bank_cd))%>"
		.txtOpenBankNm.value = "<%=ConvSPChars(E2_b_bank(S357_E2_bank_nm))%>"
		
		.txtOpenDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_open_dt))%>"

		.txtCurrency.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
			
		.txtDocAmt.text	= "<%=UNINumClientFormatByCurrency(E26_s_lc_hdr(S357_E26_lc_amt), E26_s_lc_hdr(S357_E26_cur), ggAmtOfMoneyNo)%>"
				
		.txtLocAmt.text = "<%=UniConvNumberDBToCompany(E26_s_lc_hdr(S357_E26_lc_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)%>"

		.txtXchRate.text = "<%=UNINumClientFormat(E26_s_lc_hdr(S357_E26_xch_rate), ggExchRate.DecPoint, 0)%>"
		.txtApplicant.value = "<%=ConvSPChars(E10_b_biz_partner(S357_E10_bp_cd))%>"
		.txtApplicantNm.value = "<%=ConvSPChars(E10_b_biz_partner(S357_E10_bp_nm))%>"
		.txtBeneficiary.value = "<%=ConvSPChars(E9_b_biz_partner(S357_E9_bp_cd))%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(E9_b_biz_partner(S357_E9_bp_nm))%>"
		.txttolerance.value = "<%=UNINumClientFormat(E26_s_lc_hdr(S357_E26_amt_tolerance),ggExchRate.DecPoint,0)%>"	
		.txtIncoterms.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_incoterms))%>"
		.txtIncotermsNm.value = "<%=ConvSPChars(E14_b_minor(S357_E14_minor_nm))%>"
		.txtSalesGroup.value = "<%=ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp))%>"
		.txtSalesGroupNm.value = "<%=ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp_nm))%>"
		.txtPayTerms.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_pay_meth))%>"
		.txtPayTermsNm.value = "<%=ConvSPChars(E15_b_minor(S357_E15_minor_nm))%>"
		.txtPayDur.text = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_pay_dur))%>"
		
		'Tab 2 : L/C 선적정보 
		
		.txtLatestShipDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_latest_ship_dt))%>"

		.txtTransport.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transport))%>"
		.txtTransportNm.value = "<%=ConvSPChars(E19_b_minor(S357_E19_minor_nm))%>"
		
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_trnshp_flag))%>" = "Y" Then
			.rdoTranshipment1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_trnshp_flag))%>" = "N" Then
			.rdoTranshipment2.Checked = True
		End If
		
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_partial_ship_flag))%>" = "Y" Then
			.rdoPartailShip1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_partial_ship_flag))%>" = "N" Then
			.rdoPartailShip2.Checked = True
		End If

		.txtLoadingPort.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_loading_port))%>"
		.txtLoadingPortNm.value = "<%=ConvSPChars(E17_b_minor(S357_E17_minor_nm))%>"
		.txtDischgePort.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_dischge_port))%>"
		.txtDischgePortNm.value = "<%=ConvSPChars(E18_b_minor(S357_E18_minor_nm))%>"
		.txtDeliveryPlce.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_delivery_plce))%>"
		
		'Tab 3 : 구비서류 
			
		.txtFileDt.text = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_file_dt))%>"
		.txtFileDtTxt.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_file_dt_txt))%>"
		.txtInvCnt.text = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_inv_cnt))%>"
		.txtPackList.text = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_pack_list))%>"
		
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cert_origin_flag))%>" = "Y" Then
			.chkCertOriginFlg.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cert_origin_flag))%>" = "N" Then
			.chkCertOriginFlg.Checked = False
		End If

		
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_bl_awb_flg))%>" = "Y" Then
			.rdoBLAwFlg1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_bl_awb_flg))%>" = "N" Then
			.rdoBLAwFlg2.Checked = True
		End If
		
		.txtFreight.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_freight))%>"
		.txtFreightNm.value = "<%=ConvSPChars(E24_b_minor(S357_E24_minor_nm))%>"
		.txtNotifyParty.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_notify_party))%>"
		.txtNotifyPartyNm.value = "<%=ConvSPChars(E13_b_biz_partner(S357_E13_bp_nm))%>"
		.txtConsignee.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_consignee))%>"
		.txtConsigneeNm.value = "<%=ConvSPChars(E1_b_biz_partner(S357_E1_bp_nm))%>"
		.txtInsurPolicy.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_insur_policy))%>"
		.txtDoc1.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_doc1))%>"
		.txtDoc2.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_doc2))%>"
		.txtDoc3.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_doc3))%>"
		.txtDoc4.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_doc4))%>"
		.txtDoc5.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_doc5))%>"
		
		'Tab 4 : 은행 및 기타 
			
		.txtPayBank.value = "<%=ConvSPChars(E5_b_bank(S357_E5_bank_cd))%>"
		.txtPayBankNm.value = "<%=ConvSPChars(E5_b_bank(S357_E5_bank_nm))%>"
		
		.txtRenegoBank.value = "<%=ConvSPChars(E4_b_bank(S357_E4_bank_cd))%>"
		.txtRenegoBankNm.value = "<%=ConvSPChars(E4_b_bank(S357_E4_bank_nm))%>"
		
		.txtConfirmBank.value = "<%=ConvSPChars(E6_b_bank(S357_E6_bank_cd))%>"
		.txtConfirmBankNm.value = "<%=ConvSPChars(E6_b_bank(S357_E6_bank_nm))%>"
		
		.txtBankTxt.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_bank_txt))%>"
		
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transfer_flag))%>" = "Y" Then
			.rdoTransfer1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transfer_flag))%>" = "N" Then
			.rdoTransfer2.Checked = True
		End If
		
		.txtCreditCore.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_credit_core))%>"
		.txtCreditCoreNm.value = "<%=ConvSPChars(E23_b_minor(S357_E23_minor_nm))%>"
		
		If "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_charge_cd))%>" = "AP" Then
			.rdoChargeCd1.Checked = True
		ElseIf "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_charge_cd))%>" = "BE" Then
			.rdoChargeCd2.Checked = True
		End If
		
		.txtChargeTxt.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_charge_txt))%>"
		.txtPaymentTxt.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_payment_txt))%>"
		.txtShipment.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_shipment))%>"
		.txtPreAdvRef.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_pre_adv_ref))%>"
		.txtTransportComp.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_transport_comp))%>"
		.txtOrigin.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_origin))%>"
		.txtOriginNm.value = "<%=ConvSPChars(E20_b_minor(S357_E20_minor_nm))%>"
		.txtOriginCntry.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_origin_cntry))%>"
		.txtOriginCntryNm.value = "<%=ConvSPChars(E21_b_country(S357_E21_country_nm))%>"
		.txtAgent.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_agent))%>"
		.txtAgentNm.value = "<%=ConvSPChars(E11_b_biz_partner(S357_E11_bp_nm))%>"
		.txtManufacturer.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_manufacturer))%>"
		.txtManufacturerNm.value = "<%=ConvSPChars(E12_b_biz_partner(S357_E12_bp_nm))%>"
		.txtRemark.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_remark))%>"

		.txtAmendDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_amend_dt))%>"

		Call parent.ProtectXchRate()
		Call parent.DbQueryOk()														'☜: 조회가 성공 

		.txtExchRateOp.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_xch_rate_op))%>"
		.txtHLCNo.value = "<%=ConvSPChars(Request("txtLCNo"))%>"
	End With
</Script>

<%

Response.End

Case CStr(UID_M0002)															'☜: 현재 Save 요청을 받음 

Err.Clear																'☜: Protect system from crashing

Dim CommandSent

Dim I1_S_s_lc_hdr
ReDim I1_S_s_lc_hdr(70)

Dim I6_b_bank

Dim I7_b_bank

Dim I2_b_biz_partner

Dim I3_b_biz_partner

Dim I4_b_sales_grp

Dim I8_b_bank

Dim I9_b_bank

Dim I10_b_bank

Dim E1_s_lc_hdr

Const S349_I1_lc_no = 0    
Const S349_I1_lc_doc_no = 1
Const S349_I1_so_no = 2
Const S349_I1_adv_no = 3
Const S349_I1_pre_adv_ref = 4
Const S349_I1_adv_dt = 5
Const S349_I1_open_dt = 6
Const S349_I1_expiry_dt = 7
Const S349_I1_manufacturer = 8
Const S349_I1_agent = 9
Const S349_I1_cur = 10
Const S349_I1_lc_amt = 11
Const S349_I1_xch_rate = 12
Const S349_I1_lc_loc_amt = 13
Const S349_I1_bank_txt = 14
Const S349_I1_incoterms = 15
Const S349_I1_pay_meth = 16
Const S349_I1_pay_dur = 17
Const S349_I1_payment_txt = 18
Const S349_I1_partial_ship_flag = 19
Const S349_I1_latest_ship_dt = 20
Const S349_I1_shipment = 21
Const S349_I1_doc1 = 22
Const S349_I1_doc2 = 23
Const S349_I1_doc3 = 24
Const S349_I1_doc4 = 25
Const S349_I1_doc5 = 26
Const S349_I1_file_dt = 27
Const S349_I1_file_dt_txt = 28
Const S349_I1_remark = 29
Const S349_I1_lc_kind = 30
Const S349_I1_lc_type = 31
Const S349_I1_trnshp_flag = 32
Const S349_I1_transfer_flag = 33
Const S349_I1_delivery_plce = 34
Const S349_I1_amt_tolerance = 35
Const S349_I1_loading_port = 36
Const S349_I1_dischge_port = 37
Const S349_I1_transport = 38
Const S349_I1_transport_comp = 39
Const S349_I1_origin = 40
Const S349_I1_origin_cntry = 41
Const S349_I1_charge_txt = 42
Const S349_I1_charge_cd = 43
Const S349_I1_credit_core = 44
Const S349_I1_inv_cnt = 45
Const S349_I1_bl_awb_flg = 46
Const S349_I1_freight = 47
Const S349_I1_notify_party = 48
Const S349_I1_consignee = 49
Const S349_I1_insur_policy = 50
Const S349_I1_pack_list = 51
Const S349_I1_cert_origin_flag = 52
Const S349_I1_l_lc_type = 53
Const S349_I1_open_bank_txt = 54
Const S349_I1_o_lc_doc_no = 55
Const S349_I1_o_lc_amd_seq = 56
Const S349_I1_o_lc_amend_seq = 57
Const S349_I1_o_lc_no = 58
Const S349_I1_o_lc_expiry_dt = 59
Const S349_I1_o_lc_loc_amt = 60
Const S349_I1_o_lc_type = 61
Const S349_I1_ext1_qty = 62
Const S349_I1_ext2_qty = 63
Const S349_I1_ext3_qty = 64
Const S349_I1_ext1_amt = 65
Const S349_I1_ext2_amt = 66
Const S349_I1_ext3_qmt = 67
Const S349_I1_ext1_cd = 68
Const S349_I1_ext2_cd = 69
Const S349_I1_ext3_cd = 70

lgIntFlgMode = CInt(Request("txtFlgMode"))								'☜: 저장시 Create/Update 판별 
	
		'Tab 1 : L/C 금액정보 
		
		I1_S_s_lc_hdr(S349_I1_lc_no) = Trim(UCase(Request("txtLCNo1")))
		
		If Trim(Request("txtSoNoFlg")) = "Y" Then
			I1_S_s_lc_hdr(S349_I1_so_no) = Trim(UCase(Request("txtSONo")))
		Else 
			I1_S_s_lc_hdr(S349_I1_so_no) = ""
		End If		

		I1_S_s_lc_hdr(S349_I1_lc_doc_no) = Trim(UCase(Request("txtLCDocNo")))
		I1_S_s_lc_hdr(S349_I1_adv_no) = Trim(UCase(Request("txtAdvNo")))
		I1_S_s_lc_hdr(S349_I1_lc_type) = Trim(UCase(Request("txtLCType")))
		
		If Len(Trim(Request("txtAdvDt"))) Then
			I1_S_s_lc_hdr(S349_I1_adv_dt) = UNIConvDate(Trim(Request("txtAdvDt")))
		End If
		
		I7_b_bank = Trim(UCase(Request("txtAdvBank")))
		I1_S_s_lc_hdr(S349_I1_expiry_dt) = UNIConvDate(Trim(Request("txtExpireDt")))
		I6_b_bank = Trim(UCase(Request("txtOpenBank")))
		I1_S_s_lc_hdr(S349_I1_open_dt) = UNIConvDate(Trim(Request("txtOpenDt")))
		I1_S_s_lc_hdr(S349_I1_cur) = Trim(UCase(Request("txtCurrency")))
		
		If Len(Trim(Request("txtDocAmt"))) Then
			I1_S_s_lc_hdr(S349_I1_lc_amt) = UNIConvNum(Trim(Request("txtDocAmt")),0)
		End If
		
		If Len(Trim(Request("txtLocAmt"))) Then
			I1_S_s_lc_hdr(S349_I1_lc_loc_amt) = UNIConvNum(Trim(Request("txtLocAmt")),0)
		End If	
		
		If Len(Trim(Request("txtXchRate"))) Then
			I1_S_s_lc_hdr(S349_I1_xch_rate) = UNIConvNum(Trim(Request("txtXchRate")),0)
		End If			
		
		I2_b_biz_partner = Trim(UCase(Request("txtBeneficiary")))
		
		If Len(Trim(Request("txttolerance"))) Then
			I1_S_s_lc_hdr(S349_I1_amt_tolerance) = UNIConvNum(Trim(Request("txttolerance")),0)
		End If
				
		I3_b_biz_partner = Trim(UCase(Request("txtApplicant")))
		I1_S_s_lc_hdr(S349_I1_incoterms) = Trim(UCase(Request("txtIncoterms")))
		
		I4_b_sales_grp = Trim(Request("txtSalesGroup"))
		I1_S_s_lc_hdr(S349_I1_pay_meth) = Trim(UCase(Request("txtPayTerms")))
		
		If Len(Trim(Request("txtPayDur"))) Then
			I1_S_s_lc_hdr(S349_I1_pay_dur) = UNIConvNum(Trim(Request("txtPayDur")),0)
		End If
		
		' Tap 2 : L/C 선적정보 
		
		I1_S_s_lc_hdr(S349_I1_latest_ship_dt) = UNIConvDate(Trim(Request("txtLatestShipDt")))
		I1_S_s_lc_hdr(S349_I1_transport) = Trim(UCase(Request("txtTransport")))
		I1_S_s_lc_hdr(S349_I1_trnshp_flag) = Trim(UCase(Request("rdoTranshipment")))
		I1_S_s_lc_hdr(S349_I1_partial_ship_flag) = Trim(UCase(Request("rdoPartailShip")))
		I1_S_s_lc_hdr(S349_I1_loading_port) = Trim(UCase(Request("txtLoadingPort")))
		I1_S_s_lc_hdr(S349_I1_dischge_port) = Trim(UCase(Request("txtDischgePort")))
		I1_S_s_lc_hdr(S349_I1_delivery_plce) = Trim(UCase(Request("txtDeliveryPlce")))
		
		'Tab 3 : 구비서류 
		
		I1_S_s_lc_hdr(S349_I1_file_dt) = UNIConvNum(Trim(Request("txtFileDt")),0)
		I1_S_s_lc_hdr(S349_I1_file_dt_txt) = Trim(Request("txtFileDtTxt"))
		
		If Len(Trim(Request("txtInvCnt"))) Then
			I1_S_s_lc_hdr(S349_I1_inv_cnt) = UNIConvNum(Trim(Request("txtInvCnt")),0)
		End If
		
		If Len(Trim(Request("txtPackList"))) Then
			I1_S_s_lc_hdr(S349_I1_pack_list) = UNIConvNum(Trim(Request("txtPackList")),0)
		End If
		
		If Not IsEmpty(Request("chkCertOriginFlg")) Then
			I1_S_s_lc_hdr(S349_I1_cert_origin_flag) = "Y"
		Else
			I1_S_s_lc_hdr(S349_I1_cert_origin_flag) = "N"
		End If
		
		I1_S_s_lc_hdr(S349_I1_bl_awb_flg) = Trim(Request("rdoBLAwFlg"))
		I1_S_s_lc_hdr(S349_I1_freight) = Trim(UCase(Request("txtFreight")))
		I1_S_s_lc_hdr(S349_I1_notify_party) = Trim(UCase(Request("txtNotifyParty")))
		I1_S_s_lc_hdr(S349_I1_consignee) = Trim(UCase(Request("txtConsignee")))
		I1_S_s_lc_hdr(S349_I1_insur_policy) = Trim(Request("txtInsurPolicy"))
		I1_S_s_lc_hdr(S349_I1_doc1) = Trim(Request("txtDoc1"))
		I1_S_s_lc_hdr(S349_I1_doc2) = Trim(Request("txtDoc2"))
		I1_S_s_lc_hdr(S349_I1_doc3) = Trim(Request("txtDoc3"))
		I1_S_s_lc_hdr(S349_I1_doc4) = Trim(Request("txtDoc4"))
		I1_S_s_lc_hdr(S349_I1_doc5) = Trim(Request("txtDoc5"))

		
		'Tab 4 : 은행및 기타 
		
		I8_b_bank = Trim(UCase(Request("txtPayBank")))
		I9_b_bank = Trim(UCase(Request("txtRenegoBank")))
		I10_b_bank = Trim(UCase(Request("txtConfirmBank")))
		I1_S_s_lc_hdr(S349_I1_bank_txt) = Trim(Request("txtBankTxt"))
		I1_S_s_lc_hdr(S349_I1_transfer_flag) = Request("rdoTransfer")
		I1_S_s_lc_hdr(S349_I1_credit_core) = Trim(Request("txtCreditCore"))
		
		If Request("rdoChargeCd") = "Y" Then
			I1_S_s_lc_hdr(S349_I1_charge_cd) = "AP"
		Else
			I1_S_s_lc_hdr(S349_I1_charge_cd) = "BE"
		End If	
		
		I1_S_s_lc_hdr(S349_I1_charge_txt) = Trim(Request("txtChargeTxt"))
		I1_S_s_lc_hdr(S349_I1_payment_txt) = Trim(Request("txtPaymentTxt"))
		I1_S_s_lc_hdr(S349_I1_shipment) = Trim(UCase(Request("txtShipment")))
		I1_S_s_lc_hdr(S349_I1_pre_adv_ref) = Trim(UCase(Request("txtPreAdvRef")))
		I1_S_s_lc_hdr(S349_I1_transport_comp) = Trim(UCase(Request("txtTransportComp")))
		I1_S_s_lc_hdr(S349_I1_origin) = Trim(UCase(Request("txtOrigin")))
		I1_S_s_lc_hdr(S349_I1_origin_cntry) = Trim(UCase(Request("txtOriginCntry")))
		I1_S_s_lc_hdr(S349_I1_agent) = Trim(UCase(Request("txtAgent")))
		I1_S_s_lc_hdr(S349_I1_manufacturer) = Trim(UCase(Request("txtManufacturer")))
		I1_S_s_lc_hdr(S349_I1_remark) = Trim(Request("txtRemark"))
				
		I1_S_s_lc_hdr(S349_I1_lc_kind) = "M"
		
		I1_S_s_lc_hdr(S349_I1_o_lc_amd_seq) = 0
		I1_S_s_lc_hdr(S349_I1_o_lc_amend_seq) = 0
		I1_S_s_lc_hdr(S349_I1_o_lc_loc_amt) = 0
		
		If lgIntFlgMode = OPMD_CMODE Then
			CommandSent = "CREATE"
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			CommandSent = "UPDATE"
		End If	
					
	Set PS4G111 = Server.CreateObject("PS4G111.cSLcHdrSvr")

    If CheckSYSTEMError(Err,True) = True Then
		Response.End
    End If   
			
	E1_s_lc_hdr = PS4G111.S_MAINT_LC_HDR_SVR( gStrGlobalCollection,CommandSent, _
								I1_S_s_lc_hdr,I2_b_biz_partner,I3_b_biz_partner,I4_b_sales_grp , _
								I6_b_bank,I7_b_bank,I8_b_bank,I9_b_bank,I10_b_bank) 
	
	If CheckSYSTEMError(Err,True) = True Then
		Set PS4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.End
	End If  
	
	Set PS4G111 = Nothing				
%>
	<Script Language=VBScript>
		With parent
			If "<%=ConvSPChars(E1_s_lc_hdr)%>" <> "" Then
				
				.frm1.txtLCNo.value = Replace("<%=ConvSPChars(E1_s_lc_hdr)%>","''","'")
			Else
				
				.frm1.txtLCNo.value = .frm1.txtLCNo1.value	
			End If 
			
			.DbSaveOk
		
		End With
	</Script>
<%

Response.End

Case CStr(UID_M0003)														'☜: 삭제 요청 
			
Err.Clear																'☜: Protect system from crashing

Dim I1_D_s_lc_hdr
ReDim I1_D_s_lc_hdr(0)

Dim I10_D_b_bank

Const S349_D_I1_lc_no = 0	

Dim E1_D_s_lc_hdr
Const S349_E1_D_lc_no = 0   
	
	If Trim(Request("txtLCNo")) = "" Then
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	
	I1_D_s_lc_hdr(S349_D_I1_lc_no) = Request("txtLCNo")
	
	if 	Trim(Request("txtConfirmBank")) <> "" then
		I10_D_b_bank =Trim(Request("txtConfirmBank"))
	Else
		I10_D_b_bank = ""
	End if
		
	Set PS4G111 = Server.CreateObject("PS4G111.cSLcHdrSvr")

	E1_D_s_lc_hdr = PS4G111.S_MAINT_LC_HDR_SVR(gStrGlobalCollection, "DELETE" , _
	    I1_D_s_lc_hdr,,,,,,,,I10_D_b_bank) 
			
	If CheckSYSTEMError(Err,True) = True Then
		Set PS4G111 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.End
	End If  
	
	Set PS4G111 = Nothing			
	
%>
<Script Language=VBScript>
	With parent
		.DbDeleteOk
	End With
</Script>
<%

	Response.End	

Case Else

	Response.End

End Select
%>
