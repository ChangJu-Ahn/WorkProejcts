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

	Dim lgOpModeCRUD
	Dim lgCurrency
	
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)                                                         '☜: Query
				 Call SubBizQueryMulti()
	End Select


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()															'☜: 현재 조회/Prev/Next 요청을 받음 
		Dim strDt
		Dim iPM4G228																	' Open L/C Detail 조회용 Object
		Dim iPM4G219																' Open L/C Header 조회용 Object
		Dim I_M_AmendHdr
		Dim lgCurrency
		Dim istrData
		Dim StrNextKey		' 다음 값 
		Dim lgStrPrevKey	' 이전 값 
		Dim iLngMaxRow		' 현재 그리드의 최대Row
		Dim iLngRow
		Dim GroupCount          
		Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
		Const C_SHEETMAXROWS_D  = 100
		
		Const strDefDate = "1899-12-30"
		

	Dim E1_b_daily_exchange_rate	
	Dim E2_b_minor_charge_nm
	Dim E3_b_minor_credit_core_nm
	Dim E4_b_minor_fund_type_nm
	Dim E5_b_minor_origin_nm
	Dim E6_b_minor_freight_nm
	Dim E7_b_minor_bl_awb_nm
	Dim E8_b_minor_paymeth_nm
	Dim E9_b_minor_incoterms_nm
	Dim E10_b_minor_lc_type_nm
	Dim E11_b_minor_at_transport_nm
	Dim E12_b_minor_at_loading_nm
	Dim E13_b_minor_at_dischge_nm
	Dim E14_b_minor_be_transport_nm
	Dim E15_b_minor_be_loading_nm
	Dim E16_b_minor_m_lc_amend_hdr
	Dim E17_m_lc_amend_hdr
		Const M445_E17_lc_amd_no = 0
		Const M445_E17_lc_no = 1
		Const M445_E17_lc_doc_no = 2
		Const M445_E17_lc_amend_seq = 3
		Const M445_E17_adv_no = 4
		Const M445_E17_pre_adv_ref = 5
		Const M445_E17_open_dt = 6
		Const M445_E17_be_expiry_dt = 7
		Const M445_E17_at_expiry_dt = 8
		Const M445_E17_manufacturer = 9
		Const M445_E17_agent = 10
		Const M445_E17_amend_dt = 11
		Const M445_E17_amend_req_dt = 12
		Const M445_E17_currency = 13
		Const M445_E17_be_doc_amt = 14
		Const M445_E17_at_doc_amt = 15
		Const M445_E17_at_xch_rate = 16
		Const M445_E17_inc_amt = 17
		Const M445_E17_dec_amt = 18
		Const M445_E17_be_loc_amt = 19
		Const M445_E17_at_loc_amt = 20
		Const M445_E17_be_partial_ship = 21
		Const M445_E17_at_partial_ship = 22
		Const M445_E17_be_latest_ship_dt = 23
		Const M445_E17_at_latest_ship_dt = 24
		Const M445_E17_open_bank = 25
		Const M445_E17_insrt_user_id = 26
		Const M445_E17_insrt_dt = 27
		Const M445_E17_updt_user_id = 28
		Const M445_E17_updt_dt = 29
		Const M445_E17_be_xch_rate = 30
		Const M445_E17_ext1_amt = 31
		Const M445_E17_ext1_cd = 32
		Const M445_E17_remark = 33
		Const M445_E17_lc_kind = 34
		Const M445_E17_remark2 = 35
		Const M445_E17_be_transhipment = 36
		Const M445_E17_at_transhipment = 37
		Const M445_E17_be_transfer = 38
		Const M445_E17_at_transfer = 39
		Const M445_E17_be_loading_port = 40
		Const M445_E17_at_loading_port = 41
		Const M445_E17_be_dischge_port = 42
		Const M445_E17_at_dischge_port = 43
		Const M445_E17_be_transport = 44
		Const M445_E17_at_transport = 45
		Const M445_E17_biz_area = 46
		Const M445_E17_charge_flg = 47
		Const M445_E17_ext1_qty = 48
		Const M445_E17_ext2_qty = 49
		Const M445_E17_ext3_qty = 50
		Const M445_E17_ext2_amt = 51
		Const M445_E17_ext3_amt = 52
		Const M445_E17_ext2_cd = 53
		Const M445_E17_ext3_cd = 54
		Const M445_E17_ext1_rt = 55
		Const M445_E17_ext2_rt = 56
		Const M445_E17_ext3_rt = 57
		Const M445_E17_ext1_dt = 58
		Const M445_E17_ext2_dt = 59
		Const M445_E17_ext3_dt = 60
	Dim E18_m_lc_hdr
	'조회 성능개선(2003.05.23)
	Const M445_E18_lc_no = 0
	Const M445_E18_po_no = 1
	Const M445_E18_pay_method = 2
	Const M445_E18_incoterms = 3
	Const M445_E18_doc_amt = 4
	Const M445_E18_xch_rate_op = 5
	Const M445_E18_tot_amend_amt = 6
	
	Dim E19_b_bank_issue_bank
		Const M445_E19_bank_cd = 0
		Const M445_E19_bank_nm = 1
	Dim E20_b_bank_advise_bank
		Const M445_E20_bank_cd = 0
		Const M445_E20_bank_nm = 1
	Dim E21_b_bank_renego_bank
	Dim E22_b_bank_pay_bank
	Dim E23_b_bank_confirm_bank
	Dim E24_b_pur_org
		Const M445_E24_pur_org = 0
	    Const M445_E24_pur_org_nm = 1
	Dim E25_b_pur_grp
		Const M445_E25_pur_grp = 0
	    Const M445_E25_pur_grp_nm = 1
	Dim E26_b_biz_partner_beneficiary
		Const M445_E26_bp_cd = 0
	    Const M445_E26_bp_nm = 1
	Dim E27_b_biz_partner_applicant
		Const M445_E27_bp_cd = 0
	    Const M445_E27_bp_nm = 1
	Dim E28_b_biz_partner_agent
	Dim E29_b_biz_partner_manufacturer
	Dim E30_b_biz_partner_notify_party
'*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*__*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_
' iPM4G228의 변수들 
	Dim EG1_export_group
		Const M443_EG1_E1_m_lc_amend_dtl_lc_amd_seq = 0
		Const M443_EG1_E1_m_lc_amend_dtl_po_no = 1
		Const M443_EG1_E1_m_lc_amend_dtl_po_seq = 2
		Const M443_EG1_E1_m_lc_amend_dtl_hs_cd = 3
		Const M443_EG1_E1_m_lc_amend_dtl_be_qty = 4
		Const M443_EG1_E1_m_lc_amend_dtl_at_qty = 5
		Const M443_EG1_E1_m_lc_amend_dtl_be_price = 6
		Const M443_EG1_E1_m_lc_amend_dtl_at_price = 7
		Const M443_EG1_E1_m_lc_amend_dtl_be_doc_amt = 8
		Const M443_EG1_E1_m_lc_amend_dtl_at_doc_amt = 9
		Const M443_EG1_E1_m_lc_amend_dtl_be_loc_amt = 10
		Const M443_EG1_E1_m_lc_amend_dtl_at_loc_amt = 11
		Const M443_EG1_E1_m_lc_amend_dtl_unit = 12
		Const M443_EG1_E1_m_lc_amend_dtl_insrt_user_id = 13
		Const M443_EG1_E1_m_lc_amend_dtl_insrt_dt = 14
		Const M443_EG1_E1_m_lc_amend_dtl_updt_user_id = 15
		Const M443_EG1_E1_m_lc_amend_dtl_updt_dt = 16
		Const M443_EG1_E1_m_lc_amend_dtl_ext1_qty = 17
		Const M443_EG1_E1_m_lc_amend_dtl_ext1_amt = 18
		Const M443_EG1_E1_m_lc_amend_dtl_amd_flg = 19
		Const M443_EG1_E1_m_lc_amend_dtl_lc_kind = 20
		Const M443_EG1_E1_m_lc_amend_dtl_il_no = 21
		Const M443_EG1_E1_m_lc_amend_dtl_il_seq = 22
		Const M443_EG1_E1_m_lc_amend_dtl_over_tolerance = 23
		Const M443_EG1_E1_m_lc_amend_dtl_under_tolerance = 24
		Const M443_EG1_E1_m_lc_amend_dtl_remark2 = 25
		Const M443_EG1_E1_m_lc_amend_dtl_biz_area = 26
		Const M443_EG1_E1_m_lc_amend_dtl_ext2_qty = 27
		Const M443_EG1_E1_m_lc_amend_dtl_ext3_qty = 28
		Const M443_EG1_E1_m_lc_amend_dtl_ext2_amt = 29
		Const M443_EG1_E1_m_lc_amend_dtl_ext3_amt = 30
		Const M443_EG1_E1_m_lc_amend_dtl_ext1_cd = 31
		Const M443_EG1_E1_m_lc_amend_dtl_ext2_cd = 32
		Const M443_EG1_E1_m_lc_amend_dtl_ext3_cd = 33
		Const M443_EG1_E1_m_lc_amend_dtl_ext1_rt = 34
		Const M443_EG1_E1_m_lc_amend_dtl_ext2_rt = 35
		Const M443_EG1_E1_m_lc_amend_dtl_ext3_rt = 36
		Const M443_EG1_E1_m_lc_amend_dtl_ext1_dt = 37
		Const M443_EG1_E1_m_lc_amend_dtl_ext2_dt = 38
		Const M443_EG1_E1_m_lc_amend_dtl_ext3_dt = 39
		'  View Name : export_item b_item
		Const M443_EG1_E2_b_item_item_cd = 40
		Const M443_EG1_E2_b_item_item_nm = 41
		Const M443_EG1_E2_b_item_spec = 42
		Const M443_EG1_E2_b_item_item_acct = 43
		'  View Name : export_item b_plant
		Const M443_EG1_E3_b_plant_plant_cd = 44
		Const M443_EG1_E3_b_plant_plant_nm = 45
		'  View Name : export_item b_hs_code
		Const M443_EG1_E4_b_hs_code_hs_nm = 46
		'  View Name : export_item m_pur_ord_hdr
		Const M443_EG1_E5_m_pur_ord_hdr_po_no = 47
		'  View Name : export_item m_pur_ord_dtl
		Const M443_EG1_E6_m_pur_ord_dtl_po_seq_no = 48
		Const M443_EG1_E6_m_pur_ord_dtl_tracking_no = 49
		Const M443_EG1_E6_m_pur_ord_dtl_po_qty = 50
		Const M443_EG1_E6_m_pur_ord_dtl_lc_qty = 51
		'  View Name : export_item m_lc_hdr
		Const M443_EG1_E7_m_lc_hdr_lc_no = 52
		Const M443_EG1_E7_m_lc_hdr_lc_doc_no = 53
		Const M443_EG1_E7_m_lc_hdr_lc_amend_seq = 54
		'  View Name : export_item m_lc_dtl
		Const M443_EG1_E8_m_lc_dtl_lc_seq = 55
	Dim E1_m_lc_amend_dtl
		Const M443_E1_at_doc_amt = 0
		Const M443_E1_at_loc_amt = 1

	Dim E2_m_lc_amend_dtl
	Dim E3_m_lc_amend_dtl
	Dim iGroupCount
	Dim str_txtLCAmdNo
	Dim str_NextKey
	
	On Error Resume Next
	Err.Clear																	'☜: Protect system from crashing


		If Request("txtLCAmdNo") = "" Then											'⊙: 조회를 위한 값이 들어왔는지 체크 
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
			Exit Sub
		End If
        
		lgStrPrevKey = Request("lgStrPrevKey")
		'**수정(2003.06.11)-추가 100건을 조회시는 헤더는 조회하지 않는다**
		If Request("txtMaxRows") = 0 Then
			'---------------------------------- L/C Amend Header Data Query ----------------------------------

			Set iPM4G219 = Server.CreateObject("PM4G219.cMLookupLcAmendHdrS")

			If CheckSYSTEMError(Err,True) = True Then
				Set iPM4G219 = Nothing
				Exit Sub
			End If
            
			Redim I_M_AmendHdr(1)
			I_M_AmendHdr(0) =  Request("txtLCAmdNo")
			I_M_AmendHdr(1) =  "M"
			
            Call iPM4G219.M_LOOKUP_LC_AMEND_HDR_SVR(gStrGlobalCollection, I_M_AmendHdr, E1_b_daily_exchange_rate, _
					E2_b_minor_charge_nm, E3_b_minor_credit_core_nm, E4_b_minor_fund_type_nm, E5_b_minor_origin_nm, E6_b_minor_freight_nm, _
					E7_b_minor_bl_awb_nm, E8_b_minor_paymeth_nm, E9_b_minor_incoterms_nm, E10_b_minor_lc_type_nm, E11_b_minor_at_transport_nm, _
					E12_b_minor_at_loading_nm, E13_b_minor_at_dischge_nm, E14_b_minor_be_transport_nm, E15_b_minor_be_loading_nm, E16_b_minor_m_lc_amend_hdr, _
					E17_m_lc_amend_hdr, E18_m_lc_hdr, E19_b_bank_issue_bank, E20_b_bank_advise_bank, E21_b_bank_renego_bank, E22_b_bank_pay_bank, E23_b_bank_confirm_bank, _
					E24_b_pur_org, E25_b_pur_grp, E26_b_biz_partner_beneficiary, E27_b_biz_partner_applicant, E28_b_biz_partner_agent, _
					E29_b_biz_partner_manufacturer, E30_b_biz_partner_notify_party)
			
			If IsEmpty(E17_m_lc_amend_hdr) Then
				Set iPM4G219 = Nothing
				Exit Sub
			End If

			If CheckSYSTEMError2(Err,True,"","","","","") = true then 	
				Set iPM4G219 = Nothing
				Exit Sub
			End If
	
			Set iPM4G219 = Nothing

			lgCurrency		= ConvSPChars(E17_m_lc_amend_hdr(M445_E17_currency))
		
			Response.Write "<Script Language=VBScript>" & vbCr		
			Response.Write "With parent.frm1" & vbCr
			'##### Rounding Logic #####
			'항상 거래화폐가 우선 
			Response.Write ".txtCurrency.value 		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_currency)) & """" & vbCr
			Response.Write "parent.CurFormatNumericOCX" & vbCr
			'##########################	
			Response.Write ".txtHLCAmdNo.value	    = """ & ConvSPChars(Request("txtLCAmdNo")) & """" & vbCr
			Response.Write ".txtLCDocNo.value		= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_doc_no)) & """" & vbCr
			Response.Write ".txtLCAmendSeq.value	= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_amend_seq)) & """" & vbCr
			Response.Write ".txtBeneficiary.value	= """ & ConvSPChars(E26_b_biz_partner_beneficiary(M445_E26_bp_cd)) & """" & vbCr
			Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(E26_b_biz_partner_beneficiary(M445_E26_bp_nm)) & """" & vbCr
		    'Response.Write ".txtDocAmt.value     = """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_be_doc_amt)) & """" & vbCr
		    '총품목금액을 계산한 값을 저장(2003.05) - Lee Eun Hee
			Response.Write ".txtDocAmt.text 		= """ & UNIConvNumDBToCompanyByCurrency(E18_m_lc_hdr(M445_E18_tot_amend_amt),ConvSPChars(E17_m_lc_amend_hdr(M445_E17_currency)),ggAmtOfMoneyNo,"X","X") & """" & vbCr
		    Response.Write ".txtAmendDt.text		= """ & UNIDateClientFormat(E17_m_lc_amend_hdr(M445_E17_amend_dt)) & """" & vbCr
			Response.Write ".txtMaxSeq.value 		= 0" & vbCr
			Response.Write ".txtLCNo.value 			= """ & ConvSPChars(E17_m_lc_amend_hdr(M445_E17_lc_no)) & """" & vbCr
			Response.Write ".hdnPONO.value 			= """ & ConvSPChars(E18_m_lc_hdr(M445_E18_po_no)) & """" & vbCr	
			Response.Write ".hdnIncotermsCd.Value 	= """ & ConvSPChars(E18_m_lc_hdr(M445_E18_incoterms)) & """" & vbCr
			Response.Write ".hdnIncotermsNm.Value 	= """ & ConvSPChars(E9_b_minor_incoterms_nm) & """" & vbCr
			Response.Write ".hdnPayMethCd.Value 	= """ & ConvSPChars(E18_m_lc_hdr(M445_E18_pay_method)) & """" & vbCr
			Response.Write ".hdnPayMethNm.Value 	= """ & ConvSPChars(E8_b_minor_paymeth_nm) & """" & vbCr
			Response.Write ".hdnGrpCd.Value 		= """ & ConvSPChars(E25_b_pur_grp(M445_E25_pur_grp)) & """" & vbCr
			Response.Write ".hdnGrpNm.Value 		= """ & ConvSPChars(E25_b_pur_grp(M445_E25_pur_grp_nm)) & """" & vbCr

			Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet"	& vbCr
			Response.Write "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE"	& vbCr

			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
		
		Else
			lgCurrency = request("txtCurrency")
		End If
		'---------------------------------- L/C Amend Detail Data Query ----------------------------------

		Set iPM4G228 = Server.CreateObject("PM4G228.cMListLcAmendDtlS")
		
		If CheckSYSTEMError(Err,True) = True Then
			Set iPM4G228 = Nothing
			Exit Sub
		End If
 
		str_txtLCAmdNo = Request("txtLCAmdNo")
		str_NextKey    = UNIConvNum(Request("lgStrPrevKey"),0)
		
   	  	Call iPM4G228.M_LIST_LC_AMEND_DTL_SVR(gStrGlobalCollection, _
		                                             C_SHEETMAXROWS_D, _
		                                             str_txtLCAmdNo, _
		 	                                         str_NextKey, _
					                            	 EG1_export_group, _
					                            	 E1_m_lc_amend_dtl, _
					                            	 E2_m_lc_amend_dtl,_
					                            	 E3_m_lc_amend_dtl)
 				                           
	   if Request("txtQuerytype") <> "Query" And Err.Description = "B_MESSAGE173700" then
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			IF UBound(EG1_exp_group,1) <= 0 Then
				Set iPM4G228 = Nothing
				Exit Sub												'☜: ComProxy Unload	
			End If
		Else 
			
			If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
				Set iPM4G228 = Nothing	
				'Detail항목이 없을 경우 Header정보만 보여줌 
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.write " parent.frm1.vspdData.MaxRows = 0" & chr(13)
				Response.Write " parent.dbQueryOk " & chr(13)
				Response.Write "</Script>"
				Exit Sub															'☜: 비지니스 로직 처리를 종료함 
			End If
		End if

 		
		If IsEmpty(EG1_export_group) Then
		   Set iPM4G228 = Nothing
			Exit Sub
		End If

		iLngMaxRow = CLng(Request("txtMaxRows"))
		iGroupCount = UBound(EG1_export_group, 1)
        		
		For iLngRow = 0 To UBound(EG1_export_group,1)

			If  iLngRow < C_SHEETMAXROWS_D  Then
			Else
			   iStrNextKey = Request("txtLCAmdNo")
			   Exit For
			End If  	

	 		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_amd_flg))
	        Select Case Trim(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_amd_flg))
				Case "C"
					istrData = istrData & Chr(11) & "품목추가"
				Case "U"
					istrData = istrData & Chr(11) & "내용변경"
				Case "D"
					istrData = istrData & Chr(11) & "품목삭제"
				Case Else
					istrData = istrData & Chr(11) & ""
			End Select
				
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M443_EG1_E2_b_item_item_cd))														'1
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E2_b_item_item_nm))														'1
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E2_b_item_spec))														'1
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M443_EG1_E1_m_lc_amend_dtl_unit))													'2
	        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_be_qty), ggQty.DecPoint, 0)							'3
	        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_at_qty), ggQty.DecPoint, 0)
			
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_at_price), lgCurrency, ggUnitCostNo,"X","X")
	        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_at_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X")
	        
	        'istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M443_EG1_E6_m_pur_ord_dtl_po_qty), ggQty.DecPoint, 0) - UNINumClientFormat(EG1_export_group(iLngRow, M443_EG1_E6_m_pur_ord_dtl_lc_qty), ggQty.DecPoint, 0)
	        istrData = istrData & Chr(11) & UNINumClientFormat(cdbl(EG1_export_group(iLngRow, M443_EG1_E6_m_pur_ord_dtl_po_qty))-cdbl(EG1_export_group(iLngRow, M443_EG1_E6_m_pur_ord_dtl_lc_qty)), ggQty.DecPoint, 0)
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_hs_cd))													'7
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E4_b_hs_code_hs_nm))														'8
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_lc_amd_seq))												'9
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E8_m_lc_dtl_lc_seq))
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_po_no))													'10
	        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_po_seq))													'11
	        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_over_tolerance), ggExchRate.DecPoint, 0)			'12
	        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_under_tolerance), ggExchRate.DecPoint, 0)			'13
	        istrData = istrData & Chr(11) & ""
	        istrData = istrData & Chr(11) & ""
	        '총금액계산을 위해 추가(2003.05.30) - C_OrgDocAmt, C_OrgDocAmt1
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_at_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X" ,"X")
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, M443_EG1_E1_m_lc_amend_dtl_at_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X" ,"X")
	        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow																						'14
	        istrData = istrData & Chr(11) & Chr(12)
	    Next
	    
	    If  lgStrPrevKey = E3_m_lc_amend_dtl  Then
			lgStrPrevKey = ""
		Else
			StrNextKey = E3_m_lc_amend_dtl
		End If
	    
   		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr	
		Response.Write " parent.ggoSpread.Source = .vspdData " & vbCr
		Response.Write " parent.ggoSpread.SSShowData """ & istrData & """" & vbCr
		Response.Write " parent.lgStrPrevKey = """ & StrNextKey  & """" & vbCr
	    Response.Write " parent.DbQueryOk" & vbCr 			
		Response.Write "End With" & vbCr
	    Response.Write "</Script>" & vbCr
	   
	
	Set iPM4G228 = Nothing														'☜: Unload Comproxy
	'Exit Sub																'☜: Process End

End Sub
%>
