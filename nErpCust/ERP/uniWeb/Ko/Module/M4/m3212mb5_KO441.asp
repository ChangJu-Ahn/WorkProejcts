<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212mb5.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C 내역등록 Query Transaction 처리용 ASP					*
'*  7. Modified date(First) : 2000/04/19																*
'*  8. Modified date(Last)  : 2000/04/19																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************	
	Dim lgOpModeCRUD
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")
	Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	Dim lgCurrency
	
	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)  
				Call SubBizQuery()
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================

Sub SubBizQuery()
	Dim iMax
	Dim PvArr
	Dim lGrpCnt
	Dim iTotstrData
	
	Dim iPM4G119 
	 
	Const M446_I1_lc_no = 0
	Const M446_I1_lc_kind = 1

	Dim I1_m_lc_hdr
	Dim E1_ief_supplied 
	Dim E2_b_minor 
	Dim E3_b_daily_exchange_rate 
	Dim E4_b_minor 
	Dim E5_b_minor 
	Dim E6_b_minor 
	Dim E7_b_minor 
	Dim E8_b_minor 
	Dim E9_b_minor 
	Dim E10_b_minor 
	Dim E11_b_minor 
	Dim E12_b_minor 
	Dim E13_b_minor 
	Dim E14_b_minor 
	Dim E15_b_minor 
	Dim E16_b_minor 
	Dim E17_b_minor 
	Dim E18_m_lc_hdr 
	Dim E19_b_bank 
	Dim E20_b_bank 
	Dim E21_b_bank 
	Dim E22_b_bank 
	Dim E23_b_bank 
	Dim E24_b_pur_org 
	Dim E25_b_pur_grp
		Const M476_E25_pur_grp = 0
		Const M476_E25_pur_grp_nm =1
	Dim E26_b_biz_partner
		Const M476_E26_bp_cd = 0
		Const M476_E26_bp_nm = 1
	Dim E27_b_biz_partner
		Const M476_E27_bp_cd = 0
		Const M476_E27_bp_nm = 1
	Dim E28_b_biz_partner 
	Dim E29_b_biz_partner 
	Dim E30_b_biz_partner 
   
	Const EA_m_lc_hdr_lc_no1 = 0
	Const EA_m_lc_hdr_lc_doc_no1 = 1
	Const EA_m_lc_hdr_lc_amend_seq1 = 2
	Const EA_m_lc_hdr_po_no1 = 3
	Const EA_m_lc_hdr_adv_no1 = 4
	Const EA_m_lc_hdr_pre_adv_ref1 = 5
	Const EA_m_lc_hdr_req_dt1 = 6
	Const EA_m_lc_hdr_adv_dt1 = 7
	Const EA_m_lc_hdr_open_dt1 = 8
	Const EA_m_lc_hdr_expiry_dt1 = 9
	Const EA_m_lc_hdr_amend_dt1 = 10
	Const EA_m_lc_hdr_manufacturer1 = 11
	Const EA_m_lc_hdr_agent1 = 12
	Const EA_m_lc_hdr_currency1 = 13
	Const EA_m_lc_hdr_doc_amt1 = 14
	Const EA_m_lc_hdr_xch_rate1 = 15
	Const EA_m_lc_hdr_xch_rate_op1 = 16
	Const EA_m_lc_hdr_loc_amt1 = 17
	Const EA_m_lc_hdr_bank_txt1 = 18
	Const EA_m_lc_hdr_incoterms1 = 19
	Const EA_m_lc_hdr_pay_method1 = 20
	Const EA_m_lc_hdr_pay_terms_txt1 = 21
	Const EA_m_lc_hdr_partial_ship1 = 22
	Const EA_m_lc_hdr_latest_ship_dt1 = 23
	Const EA_m_lc_hdr_shipment1 = 24
	Const EA_m_lc_hdr_doc11 = 25
	Const EA_m_lc_hdr_doc21 = 26
	Const EA_m_lc_hdr_doc31 = 27
	Const EA_m_lc_hdr_doc41 = 28
	Const EA_m_lc_hdr_doc51 = 29
	Const EA_m_lc_hdr_file_dt1 = 30
	Const EA_m_lc_hdr_file_dt_txt1 = 31
	Const EA_m_lc_hdr_insrt_user_id1 = 32
	Const EA_m_lc_hdr_insrt_dt1 = 33
	Const EA_m_lc_hdr_updt_user_id1 = 34
	Const EA_m_lc_hdr_updt_dt1 = 35
	Const EA_m_lc_hdr_ext1_qty1 = 36
	Const EA_m_lc_hdr_ext1_amt1 = 37
	Const EA_m_lc_hdr_ext1_cd1 = 38
	Const EA_m_lc_hdr_remark1 = 39
	Const EA_m_lc_hdr_lc_kind1 = 40
	Const EA_m_lc_hdr_lc_type1 = 41
	Const EA_m_lc_hdr_transhipment1 = 42
	Const EA_m_lc_hdr_transfer1 = 43
	Const EA_m_lc_hdr_delivery_plce1 = 44
	Const EA_m_lc_hdr_tolerance1 = 45
	Const EA_m_lc_hdr_loading_port1 = 46
	Const EA_m_lc_hdr_dischge_port1 = 47
	Const EA_m_lc_hdr_transport_comp1 = 48
	Const EA_m_lc_hdr_origin1 = 49
	Const EA_m_lc_hdr_origin_cntry1 = 50
	Const EA_m_lc_hdr_charge_txt1 = 51
	Const EA_m_lc_hdr_charge_cd1 = 52
	Const EA_m_lc_hdr_credit_core1 = 53
	Const EA_m_lc_hdr_lc_remn_doc_amt1 = 54
	Const EA_m_lc_hdr_lc_remn_loc_amt1 = 55
	Const EA_m_lc_hdr_fund_type1 = 56
	Const EA_m_lc_hdr_lmt_xch_rate1 = 57
	Const EA_m_lc_hdr_lmt_amt1 = 58
	Const EA_m_lc_hdr_inv_cnt1 = 59
	Const EA_m_lc_hdr_bl_awb_flg1 = 60
	Const EA_m_lc_hdr_freight1 = 61
	Const EA_m_lc_hdr_notify_party1 = 62
	Const EA_m_lc_hdr_consignee1 = 63
	Const EA_m_lc_hdr_insur_policy1 = 64
	Const EA_m_lc_hdr_pack_list1 = 65
	Const EA_m_lc_hdr_cert_origin_flg1 = 66
	Const EA_m_lc_hdr_transport1 = 67
	Const EA_m_lc_hdr_l_lc_type1 = 68
	Const EA_m_lc_hdr_o_lc_kind1 = 69
	Const EA_m_lc_hdr_o_lc_doc_no1 = 70
	Const EA_m_lc_hdr_o_lc_amend_seq1 = 71
	Const EA_m_lc_hdr_o_lc_no1 = 72
	Const EA_m_lc_hdr_o_lc_type1 = 73
	Const EA_m_lc_hdr_o_lc_open_dt1 = 74
	Const EA_m_lc_hdr_o_lc_expiry_dt1 = 75
	Const EA_m_lc_hdr_o_lc_loc_amt1 = 76
	Const EA_m_lc_hdr_biz_area1 = 77
	Const EA_m_lc_hdr_pay_dur1 = 78
	Const EA_m_lc_hdr_charge_flg1 = 79
	Const EA_m_lc_hdr_ext2_qty1 = 80
	Const EA_m_lc_hdr_ext3_qty1 = 81
	Const EA_m_lc_hdr_ext2_amt1 = 82
	Const EA_m_lc_hdr_ext3_amt1 = 83
	Const EA_m_lc_hdr_ext2_cd1 = 84
	Const EA_m_lc_hdr_ext3_cd1 = 85
	Const EA_m_lc_hdr_ext1_rt1 = 86
	Const EA_m_lc_hdr_ext2_rt1 = 87
	Const EA_m_lc_hdr_ext3_rt1 = 88
	Const EA_m_lc_hdr_ext1_dt1 = 89
	Const EA_m_lc_hdr_ext2_dt1 = 90
	Const EA_m_lc_hdr_ext3_dt1 = 91
	Const EA_m_lc_dtl_tot_item_amt = 92
	
	Const strDefDate = "1899-12-30"
	
	'===============================================================================
	'iPM4G128 변수 선언 
	'===============================================================================
	Dim iPM4G128
	Dim iLngRow
	Dim istrData
	Dim lngPORemain
	Dim iLngMaxRow
	Dim iStrPrevKey
	Dim StrNextKey  	
	
	Dim Str_I1_m_lc_hdr
	Dim I2_m_lc_dtl
	Dim E1_m_lc_dtl
	Dim E2_m_lc_dtl
    Dim E3_m_lc_dtl
    Dim EG1_export_group
	
    Const M474_I2_lc_no = 0
    Const M474_I2_lc_doc_no = 1

    Const M474_E1_doc_amt = 0
    Const M474_E1_loc_amt = 1

    Const M474_EG1_E1_m_pur_goods_mvmt_rcpt_no = 0
    'export_item m_pur_ord_hdr
    Const M474_EG1_E2_m_pur_ord_hdr_po_no = 1
    'export_item m_pur_ord_dtl
    Const M474_EG1_E3_m_pur_ord_dtl_po_seq_no = 2
    Const M474_EG1_E3_m_pur_ord_dtl_po_qty = 3
    Const M474_EG1_E3_m_pur_ord_dtl_lc_qty = 4
    Const M474_EG1_E3_m_pur_ord_dtl_after_lc_flg = 5
    'export_item m_lc_dtl
    Const M474_EG1_E4_m_lc_dtl_lc_seq = 6
    Const M474_EG1_E4_m_lc_dtl_hs_cd = 7
    Const M474_EG1_E4_m_lc_dtl_qty = 8
    Const M474_EG1_E4_m_lc_dtl_price = 9
    Const M474_EG1_E4_m_lc_dtl_doc_amt = 10
    Const M474_EG1_E4_m_lc_dtl_loc_amt = 11
    Const M474_EG1_E4_m_lc_dtl_unit = 12
    Const M474_EG1_E4_m_lc_dtl_over_tolerance = 13
    Const M474_EG1_E4_m_lc_dtl_under_tolerance = 14
    Const M474_EG1_E4_m_lc_dtl_close_flg = 15
    Const M474_EG1_E4_m_lc_dtl_receipt_qty = 16
    Const M474_EG1_E4_m_lc_dtl_insrt_user_id = 17
    Const M474_EG1_E4_m_lc_dtl_insrt_dt = 18
    Const M474_EG1_E4_m_lc_dtl_updt_user_id = 19
    Const M474_EG1_E4_m_lc_dtl_updt_dt = 20
    Const M474_EG1_E4_m_lc_dtl_ext1_qty = 21
    Const M474_EG1_E4_m_lc_dtl_ext1_amt = 22
    Const M474_EG1_E4_m_lc_dtl_ext1_cd = 23
    Const M474_EG1_E4_m_lc_dtl_lc_kind = 24
    Const M474_EG1_E4_m_lc_dtl_bl_qty = 25
    Const M474_EG1_E4_m_lc_dtl_il_no = 26
    Const M474_EG1_E4_m_lc_dtl_il_seq = 27
    Const M474_EG1_E4_m_lc_dtl_remark2 = 28
    Const M474_EG1_E4_m_lc_dtl_biz_area = 29
    Const M474_EG1_E4_m_lc_dtl_ext2_qty = 30
    Const M474_EG1_E4_m_lc_dtl_ext3_qty = 31
    Const M474_EG1_E4_m_lc_dtl_ext2_amt = 32
    Const M474_EG1_E4_m_lc_dtl_ext3_amt = 33
    Const M474_EG1_E4_m_lc_dtl_ext2_cd = 34
    Const M474_EG1_E4_m_lc_dtl_ext3_cd = 35
    Const M474_EG1_E4_m_lc_dtl_ext1_rt = 36
    Const M474_EG1_E4_m_lc_dtl_ext2_rt = 37
    Const M474_EG1_E4_m_lc_dtl_ext3_rt = 38
    Const M474_EG1_E4_m_lc_dtl_ext1_dt = 39
    Const M474_EG1_E4_m_lc_dtl_ext2_dt = 40
    Const M474_EG1_E4_m_lc_dtl_ext3_dt = 41
    'export_item b_item
    Const M474_EG1_E5_b_item_item_cd = 42
    Const M474_EG1_E5_b_item_item_nm = 43
    Const M474_EG1_E5_b_item_spec = 44
    Const M474_EG1_E5_b_item_item_acct = 45
    'export_item b_hs_code
    Const M474_EG1_E6_b_hs_code_hs_cd = 46
    Const M474_EG1_E6_b_hs_code_hs_nm = 47
    'export_item b_plant
    Const M474_EG1_E7_b_plant_plant_cd = 48
    Const M474_EG1_E7_b_plant_plant_nm = 49
    Const M474_EG1_E8_mvmt_no = 50
    Const M474_EG1_E8_tracking_no = 51
	
	On Error Resume Next
    Err.Clear 
    
    '**수정(2003.06.11)-추가 100건을 조회시는 헤더는 조회하지 않는다**
    If Request("txtMaxRows") = 0 Then
    
		ReDim I1_m_lc_hdr(M446_I1_lc_kind)
    
		I1_m_lc_hdr(M446_I1_lc_no) 	     = Trim(Request("txtLCNo"))
		I1_m_lc_hdr(M446_I1_lc_kind) 	 = "L"

    
		Set iPM4G119 = Server.CreateObject("PM4G119.cMLookupLcHdrS")    
    
		If CheckSYSTEMError(Err,True) = true then 		
			Set iPM4G119 = Nothing
		    Exit Sub
		End if
	
	
		Call iPM4G119.M_LOOKUP_LC_HDR_SVR(gStrGlobalCollection,I1_m_lc_hdr,E1_ief_supplied,E2_b_minor, _
		E3_b_daily_exchange_rate,E4_b_minor,E5_b_minor,E6_b_minor,E7_b_minor,E8_b_minor,E9_b_minor,E10_b_minor, _
		E11_b_minor,E12_b_minor, E13_b_minor, E14_b_minor, E15_b_minor,E16_b_minor,E17_b_minor,E18_m_lc_hdr, _
		E19_b_bank,E20_b_bank,E21_b_bank,E22_b_bank,E23_b_bank,E24_b_pur_org,E25_b_pur_grp,E26_b_biz_partner, _
		E27_b_biz_partner,E28_b_biz_partner,E29_b_biz_partner,E30_b_biz_partner)
    
		If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set iPM4G119 = Nothing												'☜: ComProxy Unload
		    Exit Sub
		End if
	
		Set iPM4G119 = Nothing					
	
		lgCurrency = ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_currency1))
	
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
					
		'##### Rounding Logic #####		<=========================================여기 
		'항상 거래화폐가 우선 
		Response.Write ".txtCurrency.value 		= """ & ConvSPChars(Trim(UCase(E18_m_lc_hdr(EA_m_lc_hdr_currency1)))) & """" & vbCr
		Response.Write "parent.CurFormatNumericOCX" & vbCr
		'##########################	
		'Response.Write "strDefDate = ""1899-12-30"" "& vbCr
		'Response.Write "strDefDate = """ & UNIDateClientFormat(strDefDate) & """" & vbCr
		'Response.Write ".txtHLCNo.value = """ & ConvSPChars(Request("EA_m_lc_hdr_lc_no1")) & """" & vbCr '잠깐수정 
		Response.Write ".txtHLCNo.value = """ & ConvSPChars(Request("txtLCNo")) & """" & vbCr 
		Response.Write ".txtLCDocNo.value 		= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_doc_no1)) & """" & vbCr
		Response.Write ".txtLCAmendSeq.value 	= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_lc_amend_seq1)) & """" & vbCr
		Response.Write ".txtBeneficiary.value 	= """ & ConvSPChars(E26_b_biz_partner(M476_E26_bp_cd)) & """" & vbCr
		Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(E26_b_biz_partner(M476_E26_bp_nm)) & """" & vbCr
		Response.Write ".txtOpenDt.text = """ & UNIDateClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_open_dt1)) & """" & vbcr
	
		If ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_doc_amt1))  <> 0 Then
			Response.Write ".txtDocAmt.text = """ & UNIConvNumDBToCompanyByCurrency(E18_m_lc_hdr(EA_m_lc_hdr_doc_amt1),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
		Else
			.SetDefaultVal()
		End If
	
		Response.Write ".txtHXchRate.value 		= """ & UNINumClientFormat(E18_m_lc_hdr(EA_m_lc_hdr_xch_rate1), ggExchRate.DecPoint, 0) & """" & vbCr			
		Response.Write ".hdnDiv.value 		= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_xch_rate_op1)) & """" & vbCr		
		Response.Write ".txtHPurGrp.value 		= """ & ConvSPChars(E25_b_pur_grp(M476_E25_pur_grp)) & """" & vbCr
		Response.Write ".txtHPurGrpNm.value 		= """ & ConvSPChars(E25_b_pur_grp(M476_E25_pur_grp_nm)) & """" & vbCr
		Response.Write ".txtHApplicant.value	= """ & ConvSPChars(E27_b_biz_partner(M476_E27_bp_cd)) & """" & vbCr
		Response.Write ".txtHApplicantNm.value = """ & ConvSPChars(E27_b_biz_partner(M476_E27_bp_nm)) & """" & vbCr
		Response.Write ".txtHPayTerms.value	= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_pay_method1)) & """" & vbCr
		Response.Write ".txtHPayTermsNm.value	= """ & ConvSPChars(E12_b_minor(1)) & """" & vbCr
		Response.Write ".txtHPONo.value		= """ & ConvSPChars(E18_m_lc_hdr(EA_m_lc_hdr_po_no1)) & """" & vbCr
		Response.Write ".txtMaxSeq.value = 0" & vbCr
		'총품목금액을 계산한 값을 저장(2003.05) - Lee Eun Hee
		Response.Write ".txtTotItemAmt.text 	= """ & UNIConvNumDBToCompanyByCurrency(E18_m_lc_hdr(EA_m_lc_dtl_tot_item_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	
		Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet"	& vbCr
		Response.Write "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE"	& vbCr
		Response.Write "End With" & vbCr

		Response.Write "</Script>" & vbCr	
	
	Else
		lgCurrency = request("txtCurrency")
	End If
		
	'------------ L/C Detail Data Query ---------------------------------------
		
	Const C_SHEETMAXROWS_D  = 100
    
    iStrPrevKey      = Trim(Request("lgStrPrevKey"))
    
    If Request("lgStrPrevKey") <> "" Then
		'LC_Seq값이 들어간다.
		Str_I1_m_lc_hdr = Trim(iStrPrevKey)
	Else
		'Str_I1_m_lc_hdr = ""
    End If
    
	I2_m_lc_dtl = Trim(Request("txtLCNo"))		
	
	Set iPM4G128 = Server.CreateObject("PM4G128.cMListLcDtlS")    
    
	If CheckSYSTEMError(Err,True) = true then 
		Set iPM4G128 = Nothing
        Exit Sub
	End if
	
    Call iPM4G128.M_LIST_LC_DTL_SVR(gStrGlobalCollection, Cint(C_SHEETMAXROWS_D), Str_I1_m_lc_hdr,I2_m_lc_dtl,E1_m_lc_dtl,E2_m_lc_dtl, _
                  E3_m_lc_dtl,EG1_export_group)
    
    
    if Request("txtQuerytype") <> "Query" And Err.Description = "B_MESSAGE173500" then
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			IF UBound(EG1_exp_group,1) <= 0 Then
				Set iPM4G128 = Nothing
				Exit Sub												'☜: ComProxy Unload	
			End If
		
		Else 
			If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
				Set iPM4G128 = Nothing												'☜: ComProxy Unload
				'Detail항목이 없을 경우 Header정보만 보여줌 
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
				Response.Write "parent.dbQueryOk" & chr(13)
				Response.Write "</Script>"
				Exit Sub															'☜: 비지니스 로직 처리를 종료함 
			End If
		End if
	
	Set iPM4G128 = Nothing
	
	iMax = UBound(EG1_export_group,1)
	ReDim PvArr(iMax)
	
	lGrpCnt = 0

	For iLngRow = 0 To UBound(EG1_export_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_lc_seq)) 
		   Exit For
		End If
		
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E5_b_item_item_cd))												'1
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E5_b_item_item_nm))												'2
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E5_b_item_spec))												'3
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_unit))												'4	
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_qty), ggQty.DecPoint, 0)						'5
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_price), lgCurrency, ggUnitCostNo,"X","X")	'6
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X")	'7
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_loc_amt), ggAmtOfMoney.DecPoint, 0)						'8
		istrData = istrData & Chr(11) & UNINumClientFormat((CDbl(EG1_export_group(iLngRow, M474_EG1_E3_m_pur_ord_dtl_po_qty)) - CDbl(EG1_export_group(iLngRow, M474_EG1_E3_m_pur_ord_dtl_lc_qty))), ggQty.DecPoint, 0)	'9
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_hs_cd))													'10
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E6_b_hs_code_hs_nm))												'11
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_lc_seq))												'12
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E2_m_pur_ord_hdr_po_no))											'13
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E3_m_pur_ord_dtl_po_seq_no)) 										'14
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E1_m_pur_goods_mvmt_rcpt_no))	
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E8_mvmt_no))	
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_over_tolerance), ggExchRate.DecPoint, 0)			'16
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_under_tolerance), ggExchRate.DecPoint, 0)		'17
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E3_m_pur_ord_dtl_after_lc_flg))	
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M474_EG1_E8_tracking_no))	
		'총금액계산을 위해 추가(2003.05.30) - C_OrgLocAmt, C_OrgLocAmt1
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_loc_amt), ggAmtOfMoney.DecPoint, 0)
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M474_EG1_E4_m_lc_dtl_loc_amt), ggAmtOfMoney.DecPoint, 0)				
		istrData = istrData & Chr(11) & iLngMaxRow + iLngRow																				'18
		istrData = istrData & Chr(11) & Chr(12)
		
		PvArr(lGrpCnt) = istrData
        lGrpCnt = lGrpCnt + 1
        istrData = ""

	Next
	
	iTotstrData = Join(PvArr, "")

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr     
	Response.Write " .ggoSpread.Source = .frm1.vspdData" & vbCr 
	Response.Write " .ggoSpread.SSShowData     """ & iTotstrData & """" & vbCr
	Response.Write ".lgStrPrevKey = """ & StrNextKey & """" & vbCr
	Response.Write " 	.frm1.txtHLCNo.value = """ & ConvSPChars(Request("txtLCNo")) & """" & vbCr
	Response.Write " 	.DbQueryOk" & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr
	
End Sub

%>

