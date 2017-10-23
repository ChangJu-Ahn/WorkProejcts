<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution                                                      *
'*  2. Function Name        :                                                                           *
'*  3. Program ID           : S5211MB1
'*  4. Program Name         : 수출 B/L등록                                                                          *
'*  5. Program Desc         : 수출 B/L등록																*
'*  6. Comproxy List        : PS7G131.dll,PS7G115.dll										            *
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/11/15																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Ahn Tae Hee												                *
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : 화면 design												*
'*							  2. 2000/04/18 : Coding Start												*
'*                            3. 2002/11/15 : UI 표준적용												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%	
Dim strMode
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Dim pvCB 
Call HideStatusWnd                                                               '☜: Hide Processing message

strMode      = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
Select Case strMode
    Case CStr(UID_M0001)                                                        '☜: Query
         Call SubBizQuery()
    Case CStr(UID_M0002)
         Call SubBizSave()
    Case CStr(UID_M0003)                                                        '☜: Delete
         Call SubBizDelete()
    Case CStr("PostFlag")		     											'☜: 확정 요청 
	     Call SubBizPostFlag     
End Select
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    Call HideStatusWnd
    Dim lgCurrency
    Dim lgArrGlFlag   
    Dim lgStrGlFlag
    Err.Clear                                                               '☜: Protect system from crashing
    
	Call SubOpenDB(lgObjConn)
	call SubMakeSQLStatements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
	    lgObjRs.Close
	    lgObjConn.Close
	    Set lgObjRs = Nothing
	    Set lgObjConn = Nothing
	    If Request("txtPrevNext") = "Q" Then
			'B/L정보가 없습니다.
		    Call DisplayMsgBox("205300", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		Elseif Request("txtPrevNext") = "P" Then
			'이전 자료가 없습니다.
		    Call DisplayMsgBox("200002", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		Else
			'이후 자료가 없습니다.
		    Call DisplayMsgBox("200003", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		End If
	Exit sub
	End If
	
	lgCurrency = ConvSPChars(lgObjRs("Cur"))	
	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr


		' Tab 1: 선적정보 1

	Response.Write ".txtCurrency.value	    = """ & lgCurrency         & """" & vbCr
	Response.Write " parent.CurFormatNumericOCX " & vbCr 		

    Response.Write ".txtBLNo.value         = """ & ConvSPChars(lgObjRs("bill_no"))                  & """" & vbCr
    Response.Write ".txtBLNo1.value        = """ & ConvSPChars(lgObjRs("bill_no"))                  & """" & vbCr
	Response.Write ".txtSONo.value		   = """ & ConvSPChars(Trim(lgObjRs("so_no")))              & """" & vbCr
	Response.Write ".txtBLDocNo.value	   = """ & ConvSPChars(lgObjRs("bl_doc_no"))                & """" & vbCr
	Response.Write ".txtLCDocNo.value	   = """ & ConvSPChars(lgObjRs("lc_doc_no"))                & """" & vbCr

	If Trim(ConvSPChars(lgObjRs("lc_doc_no"))) = ""Then
		Response.Write".txtLCAmendSeq.value = """"" & vbCr
	Else
		Response.Write".txtLCAmendSeq.value = """   & lgObjRs("lc_amend_seq")          & """" & vbCr
	End If
	Response.Write ".txtBLIssueDt.text		= """ & UNIDateClientFormat(lgObjRs("bl_issue_dt"))   & """" & vbCr		

	Response.Write ".txtDocAmt.Text			= """ & UNINumClientFormatByCurrency(lgObjRs("bill_amt"), lgCurrency, ggAmtOfMoneyNo)   & """" & vbCr

	Response.Write ".txtXchRate.text		= """ & UNINumClientFormat(lgObjRs("Xchg_rate"), ggExchRate.DecPoint, 0)   & """" & vbCr

	Response.Write ".txtLocAmt.Text			= """ & UniConvNumberDBToCompany(lgObjRs("bill_amt_loc"), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)   & """" & vbCr

	Response.Write ".txtTransport.value		= """ & ConvSPChars(lgObjRs("transport"))   & """" & vbCr
	Response.Write ".txtTransportNm.value	= """ & ConvSPChars(lgObjRs("transport_nm"))   & """" & vbCr
	Response.Write ".txtApplicant.value		= """ & ConvSPChars(lgObjRs("applicant"))   & """" & vbCr
	Response.Write ".txtApplicantNm.value	= """ & ConvSPChars(lgObjRs("applicant_nm"))   & """" & vbCr
	Response.Write ".txtLoadingPort.value	= """ & ConvSPChars(lgObjRs("loading_port"))   & """" & vbCr
	Response.Write ".txtLoadingPortNm.value = """ & ConvSPChars(lgObjRs("loading_port_nm"))   & """" & vbCr
	Response.Write ".txtIncoterms.value		= """ & ConvSPChars(lgObjRs("incoterms"))   & """" & vbCr
	Response.Write ".txtIncotermsNm.value	= """ & ConvSPChars(lgObjRs("incoterms_nm"))   & """" & vbCr
	Response.Write ".txtDischgePort.value	= """ & ConvSPChars(lgObjRs("dischge_port"))   & """" & vbCr
	Response.Write ".txtDischgePortNm.value = """ & ConvSPChars(lgObjRs("dischge_port_nm"))   & """" & vbCr
	Response.Write ".txtSalesGroup.value	= """ & ConvSPChars(lgObjRs("sales_grp"))   & """" & vbCr
	Response.Write ".txtSalesGroupNm.value	= """ & ConvSPChars(lgObjRs("sales_grp_nm"))   & """" & vbCr
		
	Response.Write ".txtLoadingDt.text		= """ & UNIDateClientFormat(lgObjRs("loading_dt"))   & """" & vbCr
	
	Response.Write ".txtBeneficiary.value	= """ & ConvSPChars(lgObjRs("beneficiary"))   & """" & vbCr
	Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(lgObjRs("beneficiary_nm"))   & """" & vbCr
	Response.Write ".txtFreight.value		= """ & ConvSPChars(lgObjRs("freight"))   & """" & vbCr
	Response.Write ".txtFreightNm.value		= """ & ConvSPChars(lgObjRs("freight_nm"))   & """" & vbCr
	Response.Write ".txtBLIssueCnt.text		= """ & ConvSPChars(lgObjRs("bl_issue_cnt"))   & """" & vbCr
	Response.Write ".txtBLIssuePlce.value	= """ & ConvSPChars(lgObjRs("bl_issue_plce"))   & """" & vbCr
		
		' Tab 2 : 선적정보 2
		
	Response.Write ".txtAgent.value				= """ & ConvSPChars(lgObjRs("agent"))   & """" & vbCr
	Response.Write ".txtAgentNm.value			= """ & ConvSPChars(lgObjRs("agent_nm"))   & """" & vbCr
	Response.Write ".txtManufacturer.value		= """ & ConvSPChars(lgObjRs("manufacturer"))   & """" & vbCr
	Response.Write ".txtManufacturerNm.value	= """ & ConvSPChars(lgObjRs("manufacturer_nm"))   & """" & vbCr
	Response.Write ".txtVesselNm.value			= """ & ConvSPChars(lgObjRs("vessel_nm"))   & """" & vbCr
	Response.Write ".txtVoyageNo.value			= """ & ConvSPChars(lgObjRs("voyage_no"))   & """" & vbCr
	Response.Write ".txtForwarder.value			= """ & ConvSPChars(lgObjRs("forwarder"))   & """" & vbCr
	Response.Write ".txtForwarderNm.value		= """ & ConvSPChars(lgObjRs("forwarder_nm"))   & """" & vbCr
	Response.Write ".txtVesselCntry.value		= """ & ConvSPChars(lgObjRs("vessel_cntry"))   & """" & vbCr
	Response.Write ".txtVesselCntryNm.value		= """ & ConvSPChars(lgObjRs("vessel_cntry_nm"))   & """" & vbCr
	Response.Write ".txtReceiptPlce.value		= """ & ConvSPChars(lgObjRs("receipt_plce"))   & """" & vbCr
	Response.Write ".txtDeliveryPlce.value		= """ & ConvSPChars(lgObjRs("delivery_plce"))   & """" & vbCr
	Response.Write ".txtFinalDest.value			= """ & ConvSPChars(lgObjRs("final_dest"))   & """" & vbCr
	Response.Write ".txtDischgeDt.text			= """ & UNIDateClientFormat(lgObjRs("dischge_plan_dt"))   & """" & vbCr
	Response.Write ".txtTranshipCntry.value		= """ & ConvSPChars(lgObjRs("tranship_cntry"))   & """" & vbCr
	Response.Write ".txtTranshipCntryNm.value	= """ & ConvSPChars(lgObjRs("tranship_cntry_nm"))   & """" & vbCr
	Response.Write ".txtTranshipDt.text			= """ & UNIDateClientFormat(lgObjRs("tranship_dt"))   & """" & vbCr
	Response.Write ".txtPackingType.value		= """ & ConvSPChars(lgObjRs("packing_type"))   & """" & vbCr
	Response.Write ".txtPackingTypeNm.value		= """ & ConvSPChars(lgObjRs("packing_type_nm"))   & """" & vbCr
	Response.Write ".txtTotPackingCnt.text		= """ & ConvSPChars(lgObjRs("tot_packing_cnt"))   & """" & vbCr
	Response.Write ".txtPackingTxt.value		= """ & ConvSPChars(lgObjRs("packing_txt"))   & """" & vbCr
	Response.Write ".txtContainerCnt.text		= """ & ConvSPChars(lgObjRs("container_cnt"))   & """" & vbCr
	Response.Write ".txtGrossWeight.text		= """ & UNINumClientFormat(lgObjRs("gross_weight"), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtWeightUnit.value		= """ & ConvSPChars(lgObjRs("weight_unit"))   & """" & vbCr
	Response.Write ".txtGrossVolumn.text		= """ & UNINumClientFormat(lgObjRs("gross_volumn"), ggQty.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtVolumnUnit.value		= """ & ConvSPChars(lgObjRs("volumn_unit"))   & """" & vbCr
	Response.Write ".txtOrigin.value			= """ & ConvSPChars(lgObjRs("origin"))   & """" & vbCr
	Response.Write ".txtOriginNm.value			= """ & ConvSPChars(lgObjRs("origin_nm"))   & """" & vbCr
	Response.Write ".txtOriginCntry.value		= """ & ConvSPChars(lgObjRs("origin_cntry"))   & """" & vbCr
	Response.Write ".txtOriginCntryNm.value		= """ & ConvSPChars(lgObjRs("origin_cntry_nm"))   & """" & vbCr
	Response.Write ".txtFreightPlce.value		= """ & ConvSPChars(lgObjRs("freight_plce"))   & """" & vbCr
		
		'Tab 3 : 매출정보 
		
	Response.Write ".txtTaxBizArea.value		= """ & ConvSPChars(lgObjRs("tax_biz_area"))   & """" & vbCr
	Response.Write ".txtTaxBizAreaNm.value		= """ & ConvSPChars(lgObjRs("tax_biz_area_nm"))   & """" & vbCr
	Response.Write ".txtBillType.value			= """ & ConvSPChars(lgObjRs("bill_type"))   & """" & vbCr
	Response.Write ".txtBillTypeNm.value		= """ & ConvSPChars(lgObjRs("bill_type_nm"))   & """" & vbCr
	Response.Write ".txtPayer.value				= """ & ConvSPChars(lgObjRs("payer"))   & """" & vbCr
	Response.Write ".txtPayerNm.value			= """ & ConvSPChars(lgObjRs("payer_nm"))   & """" & vbCr
	Response.Write ".txtBilltoParty.value		= """ & ConvSPChars(lgObjRs("bill_to_party"))   & """" & vbCr
	Response.Write ".txtBilltoPartyNm.value		= """ & ConvSPChars(lgObjRs("bill_to_party_nm"))   & """" & vbCr
	Response.Write ".txtToSalesGroup.value		= """ & ConvSPChars(lgObjRs("to_sales_grp"))   & """" & vbCr 
	Response.Write ".txtToSalesGroupNm.value	= """ & ConvSPChars(lgObjRs("to_sales_grp_nm"))   & """" & vbCr 
	Response.Write ".txtCurrency1.value			= """ & lgCurrency   & """" & vbCr
	Response.Write ".txtDocAmt1.Text			= """ & UNINumClientFormatByCurrency(lgObjRs("bill_amt"), lgCurrency, ggAmtOfMoneyNo)   & """" & vbCr
		
	Response.Write ".txtPayDt.text				= """ & UNIDateClientFormat(lgObjRs("income_plan_dt"))   & """" & vbCr

	Response.Write ".txtLocAmt1.Text			= """ & UniConvNumberDBToCompany(lgObjRs("bill_amt_loc"), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)   & """" & vbCr

	Response.Write ".txtPayType.value			= """ & ConvSPChars(lgObjRs("pay_type"))   & """" & vbCr
	Response.Write ".txtPayTypeNm.value			= """ & ConvSPChars(lgObjRs("pay_type_nm"))   & """" & vbCr
	Response.Write ".txtPayTerms.value			= """ & ConvSPChars(lgObjRs("pay_meth"))   & """" & vbCr
	Response.Write ".txtPayTermsNm.value		= """ & ConvSPChars(lgObjRs("pay_meth_nm"))   & """" & vbCr
	Response.Write ".txtMoney.Text				= """ & UNINumClientFormatByCurrency(lgObjRs("collect_amt"), lgCurrency, ggAmtOfMoneyNo)   & """" & vbCr

	Response.Write ".txtCollectLocAmt.Text		= """ & UniConvNumberDBToCompany(lgObjRs("collect_amt_loc"), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)   & """" & vbCr
		
	Response.Write ".txtPayDur.value			= """ & ConvSPChars(lgObjRs("pay_dur"))   & """" & vbCr
	Response.Write ".txtPayTermstxt.value		= """ & ConvSPChars(lgObjRs("pay_terms_txt"))   & """" & vbCr
	Response.Write ".txtRemark.value			= """ & ConvSPChars(lgObjRs("remark"))   & """" & vbCr		
	Response.Write ".txtRefFlg.value			= """ & ConvSPChars(lgObjRs("ref_flag"))   & """" & vbCr
	Response.Write ".txtStatusFlg.value			= """ & ConvSPChars(lgObjRs("sts"))   & """" & vbCr			
		
	Response.Write ".txtVatIncFlag.value		= """ & ConvSPChars(lgObjRs("vat_inc_flag"))   & """" & vbCr				
	Response.Write ".txtVatType.value			= """ & ConvSPChars(lgObjRs("vat_type"))   & """" & vbCr				
	Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(lgObjRs("vat_type_nm"))   & """" & vbCr
		
	Response.Write ".txtVATRate.value			= """ & UNINumClientFormat(lgObjRs("vat_rate"), ggExchRate.DecPoint, 0)   & """" & vbCr
 
		'약정회전일 
	Response.Write ".txtCreditRot.value = """ & lgObjRs("credit_rot_day")   & """" & vbCr

		If lgObjRs("post_flag") = "Y" AND Len(Trim(lgObjRs("gl_no"))) Then
		    lgArrGlFlag = Split(lgObjRs("gl_no"), Chr(11))
		    lgStrGlFlag = lgArrGlFlag(0)
		    If lgArrGlFlag(0) = "G" Then	
			 '회계전표번호 
			    Response.Write ".txtGLNo.value	    = """ & lgArrGlFlag(1) & """" & vbCr
		    ElseIf lgArrGlFlag(0) = "T" Then
			 '결의전표번호 
			    Response.Write ".txtTempGLNo.value	= """ & lgArrGlFlag(1) & """" & vbCr	
		    Else
			 'Batch번호 
			    Response.Write ".txtBatchNo.value	= """ & lgArrGlFlag(1) & """" & vbCr
		    End If
		Else
		    Response.Write ".txtGLNo.value	    = """""	 & vbCr
		    Response.Write ".txtTempGLNo.value  = """""	 & vbCr
		    Response.Write ".txtBatchNo.value	= """""	 & vbCr
		End If
		
		If lgObjRs("post_flag") = "Y" Then
			Response.Write ".rdoPostingflg1.Checked = True       "	 & vbCr
			Response.Write ".btnPosting.value = ""확정취소"" "	 & vbCr

			if lgStrGlFlag = "G" Or lgStrGlFlag = "T" Then
				Response.Write ".btnGLView.disabled = False "	 & vbCr
			Else
				Response.Write ".btnGLView.disabled = True  "	 & vbCr
			End if
		Else
			Response.Write ".rdoPostingflg2.Checked = True "	 & vbCr
			Response.Write ".btnPosting.value = ""확정"" "	 & vbCr
			Response.Write ".btnGLView.disabled = True "	     & vbCr
		End If

		If Trim(lgObjRs("ref_flag")) = "M" then
			Response.Write ".btnPosting.disabled = true  "	     & vbCr
		Else
			Response.Write ".btnPosting.disabled = False "	     & vbCr
		End If
		Response.Write ".txtHBLNo.value = """ & ConvSPChars(Request("txtBLNo")) & """" & vbCr

		'선수금 현황 버튼 Enable
		IF lgObjRs("PreRcpt_flag") = "Y" Then
			Response.Write ".btnPreRcptView.disabled = False  "	     & vbCr
		Else
			Response.Write ".btnPreRcptView.disabled = True   "	     & vbCr
		End If


	Response.Write "parent.DbQueryOk	"                            & vbCr					'☜: 조회가 성공 
'    Response.Write "Call parent.ProtectXchRate()   "                 & vbCr
    Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing
	
										
End Sub	
'============================================================================================================
Sub SubBizSave()

	Dim iS51131
	Dim lgIntFlgMode
    Dim iCommandSent
    
    Dim E1_s_bill_hdr_no   
    Dim I1_s_bill_type_config_type 'imp s_bill_type_config
    Dim I2_b_biz_partner_bp_cd     'imp_sold_to_party b_biz_partner
    Dim I3_s_bl_info
    Dim I5_s_bill_hdr 
    Dim I6_b_biz_partner           'imp_bill_to_party b_biz_partner
    Dim I7_b_biz_partner           'imp_payer b_biz_partner
    Dim I8_b_sales_grp             'imp_income b_sales_grp
    Dim I9_b_sales_grp             'imp_billing b_sales_grp
    
    Const S510_I3_bl_doc_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_bl_info
    Const S510_I3_ship_no = 1
    Const S510_I3_manufacturer = 2
    Const S510_I3_agent = 3
    Const S510_I3_receipt_plce = 4
    Const S510_I3_vessel_nm = 5
    Const S510_I3_voyage_no = 6
    Const S510_I3_forwarder = 7
    Const S510_I3_vessel_cntry = 8
    Const S510_I3_loading_port = 9
    Const S510_I3_dischge_port = 10
    Const S510_I3_delivery_plce = 11
    Const S510_I3_loading_plan_dt = 12
    Const S510_I3_latest_ship_dt = 13
    Const S510_I3_dischge_plan_dt = 14
    Const S510_I3_transport = 15
    Const S510_I3_tranship_cntry = 16
    Const S510_I3_tranship_dt = 17
    Const S510_I3_final_dest = 18
    Const S510_I3_incoterms = 19
    Const S510_I3_packing_type = 20
    Const S510_I3_tot_packing_cnt = 21
    Const S510_I3_container_cnt = 22
    Const S510_I3_packing_txt = 23
    Const S510_I3_gross_weight = 24
    Const S510_I3_weight_unit = 25
    Const S510_I3_gross_volumn = 26
    Const S510_I3_volumn_unit = 27
    Const S510_I3_freight = 28
    Const S510_I3_freight_plce = 29
    Const S510_I3_trans_price = 30
    Const S510_I3_trans_currency = 31
    Const S510_I3_trans_doc_amt = 32
    Const S510_I3_trans_xch_rate = 33
    Const S510_I3_trans_loc_amt = 34
    Const S510_I3_bl_issue_cnt = 35
    Const S510_I3_bl_issue_plce = 36
    Const S510_I3_bl_issue_dt = 37
    Const S510_I3_origin = 38
    Const S510_I3_origin_cntry = 39
    Const S510_I3_loading_dt = 40
    Const S510_I3_ext1_qty = 41
    Const S510_I3_ext2_qty = 42
    Const S510_I3_ext3_qty = 43
    Const S510_I3_ext1_amt = 44
    Const S510_I3_ext2_amt = 45
    Const S510_I3_ext3_amt = 46
    Const S510_I3_ext1_cd = 47
    Const S510_I3_ext2_cd = 48
    Const S510_I3_ext3_cd = 49


    Const S510_I5_bill_no = 0
    Const S510_I5_bill_dt = 1
    Const S510_I5_cur = 2
    Const S510_I5_xchg_rate = 3
    Const S510_I5_vat_type = 4
    Const S510_I5_vat_rate = 5
    Const S510_I5_pay_meth = 6
    Const S510_I5_pay_dur = 7
    Const S510_I5_tax_bill_no = 8
    Const S510_I5_beneficiary = 9
    Const S510_I5_applicant = 10
    Const S510_I5_post_flag = 11
    Const S510_I5_remark = 12
    Const S510_I5_vat_calc_type = 13
    Const S510_I5_tax_biz_area = 14
    Const S510_I5_pay_type = 15
    Const S510_I5_pay_terms_txt = 16
    Const S510_I5_collect_amt = 17
    Const S510_I5_collect_amt_loc = 18
    Const S510_I5_income_plan_dt = 19
    Const S510_I5_so_no = 20
    Const S510_I5_lc_no = 21
    Const S510_I5_lc_doc_no = 22
    Const S510_I5_lc_amend_seq = 23
    Const S510_I5_bl_flag = 24
    Const S510_I5_ref_flag = 25
    Const S510_I5_except_flag = 26
    Const S510_I5_reverse_flag = 27
    Const S510_I5_ext1_qty = 28
    Const S510_I5_ext2_qty = 29
    Const S510_I5_ext3_qty = 30
    Const S510_I5_ext1_amt = 31
    Const S510_I5_ext2_amt = 32
    Const S510_I5_ext3_amt = 33
    Const S510_I5_ext1_cd = 34
    Const S510_I5_ext2_cd = 35
    Const S510_I5_ext3_cd = 36
    Const S510_I5_vat_auto_flag = 37
    Const S510_I5_vat_inc_flag = 38

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear    
    lgIntFlgMode = CInt(Request("txtFlgMode"))								'☜: 저장시 Create/Update 판별 
 

    ReDim I3_s_bl_info(S510_I3_ext3_cd)
    ReDim I5_s_bill_hdr(S510_I5_vat_inc_flag) 
	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
		
	'Tab 1 : 선적정보 1
		
	I5_s_bill_hdr(S510_I5_bill_no)			= UCase(Trim(Request("txtBLNo1")))	
	
	If Trim(Request("txtSONoFlg")) = "Y" Then
		I5_s_bill_hdr(S510_I5_so_no)			= UCase(Trim(Request("txtSONo")))	
	End If

	I3_s_bl_info(S510_I3_bl_doc_no)	     		= UCase(Trim(Request("txtBLDocNo")))
	I5_s_bill_hdr(S510_I5_lc_doc_no)			= UCase(Trim(Request("txtLCDocNo")))	
		
	If Len(Trim(Request("txtLCAmendSeq"))) Then
		I5_s_bill_hdr(S510_I5_lc_amend_seq) = UNIConvNum(Request("txtLCAmendSeq"),0)
	End If

	I5_s_bill_hdr(S510_I5_lc_no)				= Trim(Request("txtHLCNo"))	
	I3_s_bl_info(S510_I3_bl_issue_dt)			= UNIConvDate(Request("txtBLIssueDt"))
	I5_s_bill_hdr(S510_I5_bill_dt)			    = UNIConvDate(Request("txtBLIssueDt"))
	I5_s_bill_hdr(S510_I5_cur)			     	= Trim(Request("txtCurrency"))	

	If Len(Trim(Request("txtXchRate"))) Then
		I5_s_bill_hdr(S510_I5_xchg_rate)        = UNIConvNum(Request("txtXchRate"),0)
	End If

	I3_s_bl_info(S510_I3_transport)			= Trim(Request("txtTransport"))	
	I5_s_bill_hdr(S510_I5_applicant)		= Trim(Request("txtApplicant"))	
	I2_b_biz_partner_bp_cd                  = Trim(Request("txtApplicant"))	
		
	I3_s_bl_info(S510_I3_loading_port)		= Trim(Request("txtLoadingPort"))	
	I3_s_bl_info(S510_I3_incoterms)			= Trim(Request("txtIncoterms"))	
	I3_s_bl_info(S510_I3_dischge_port)		= Trim(Request("txtDischgePort"))	
	I9_b_sales_grp	= Trim(Request("txtSalesGroup"))	
	I3_s_bl_info(S510_I3_loading_dt)		= UNIConvDate(Request("txtLoadingDt"))
	I5_s_bill_hdr(S510_I5_beneficiary)		= Trim(Request("txtBeneficiary"))	
	I3_s_bl_info(S510_I3_freight)			= Trim(Request("txtFreight"))	
		
	If Len(Trim(Request("txtBLIssueCnt"))) Then
		I3_s_bl_info(S510_I3_bl_issue_cnt)  = UNIConvNum(Request("txtBLIssueCnt"),0)
	End If
		
	I3_s_bl_info(S510_I3_bl_issue_plce)		= Trim(Request("txtBLIssuePlce"))	
	
	'Tab 2 : 선적정보 2
	
	I3_s_bl_info(S510_I3_agent)				= Trim(Request("txtAgent"))	
	I3_s_bl_info(S510_I3_manufacturer)		= Trim(Request("txtManufacturer"))	
	I3_s_bl_info(S510_I3_vessel_nm)			= Trim(Request("txtVesselNm"))	
	I3_s_bl_info(S510_I3_voyage_no)			= Trim(Request("txtVoyageNo"))	
	I3_s_bl_info(S510_I3_forwarder)			= Trim(Request("txtForwarder"))	
	I3_s_bl_info(S510_I3_vessel_cntry)		= Trim(Request("txtVesselCntry"))	
	I3_s_bl_info(S510_I3_receipt_plce)		= Trim(Request("txtReceiptPlce"))	
	I3_s_bl_info(S510_I3_delivery_plce)		= Trim(Request("txtDeliveryPlce"))	
	I3_s_bl_info(S510_I3_final_dest)		= Trim(Request("txtFinalDest"))	
	I3_s_bl_info(S510_I3_dischge_plan_dt)	= UNIConvDate(Request("txtDischgeDt"))
	I3_s_bl_info(S510_I3_tranship_cntry)	= Trim(Request("txtTranshipCntry"))	
	I3_s_bl_info(S510_I3_tranship_dt)		= UNIConvDate(Request("txtTranshipDt"))
	I3_s_bl_info(S510_I3_packing_type)		= Trim(Request("txtPackingType"))	
	If Len(Trim(Request("txtTotPackingCnt"))) Then
		I3_s_bl_info(S510_I3_tot_packing_cnt) = UNIConvNum(Request("txtTotPackingCnt"),0)
	End If

	I3_s_bl_info(S510_I3_packing_txt)		= Trim(Request("txtPackingTxt"))	

	If Len(Trim(Request("txtContainerCnt"))) Then
		I3_s_bl_info(S510_I3_container_cnt) = UNIConvNum(Request("txtContainerCnt"),0)
	End If

	If Len(Trim(Request("txtGrossWeight"))) Then
		I3_s_bl_info(S510_I3_gross_weight)  = UNIConvNum(Request("txtGrossWeight"),0)
	End If

	I3_s_bl_info(S510_I3_weight_unit)		= Trim(Request("txtWeightUnit"))	

	If Len(Trim(Request("txtGrossVolumn"))) Then
		I3_s_bl_info(S510_I3_gross_volumn)  = UNIConvNum(Request("txtGrossVolumn"),0)
	End If

	I3_s_bl_info(S510_I3_volumn_unit)		= Trim(Request("txtVolumnUnit"))	
	I3_s_bl_info(S510_I3_origin	)			= Trim(Request("txtOrigin"))	
	I3_s_bl_info(S510_I3_origin_cntry)		= Trim(Request("txtOriginCntry"))	
	I3_s_bl_info(S510_I3_freight_plce)		= Trim(Request("txtFreightPlce"))	
	
	'Tab 3 : 매출정보 
		
	I5_s_bill_hdr(S510_I5_tax_biz_area)	= Trim(Request("txtTaxBizArea"))
	I1_s_bill_type_config_type          	= Trim(Request("txtBillType"))	
	I7_b_biz_partner                		= Trim(Request("txtPayer"))	
	I6_b_biz_partner                        = Trim(Request("txtBilltoParty"))	
	I5_s_bill_hdr(S510_I5_post_flag)			= Request("rdoPostingflg")	   
	I8_b_sales_grp                       	= Trim(Request("txtToSalesGroup"))	
	I5_s_bill_hdr(S510_I5_income_plan_dt)	= UNIConvDate(Request("txtPayDt"))
	
	I5_s_bill_hdr(S510_I5_pay_type)		= Trim(Request("txtPayType"))	    
	I5_s_bill_hdr(S510_I5_pay_meth)		= Trim(Request("txtPayTerms"))	      
		
	If Len(Trim(Request("txtPayDur"))) Then
		I5_s_bill_hdr(S510_I5_pay_dur) = UNIConvNum(Request("txtPayDur"),0)
	End If
		
	I5_s_bill_hdr(S510_I5_pay_terms_txt)	= Request("txtPayTermstxt")	    
	I5_s_bill_hdr(S510_I5_remark)			= Request("txtRemark")	        
	

	I5_s_bill_hdr(S510_I5_bl_flag)			= "Y"
	I5_s_bill_hdr(S510_I5_except_flag)		= "N"
	I5_s_bill_hdr(S510_I5_ref_flag)			= Trim(Request("txtRefFlg"))
	I5_s_bill_hdr(S510_I5_reverse_flag)		= "N"

	I5_s_bill_hdr(S510_I5_vat_calc_type)	= "1"
	I5_s_bill_hdr(S510_I5_vat_inc_flag)		= "1"
	'<--부가세유형 -->
	I5_s_bill_hdr(S510_I5_vat_type) = Trim(Request("txtVatType"))
	'<--부가세율 -->

    If Len(Trim(Request("txtVatRate"))) Then I5_s_bill_hdr(S510_I5_vat_rate) = UNIConvNum(Trim(Request("txtVatRate")),0)
		
	I5_s_bill_hdr(S510_I5_vat_auto_flag)		= "N"
	
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If
    Set iS51131 = Server.CreateObject("PS7G131.cSBlInfoSvr")
    
	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    E1_s_bill_hdr_no= ""
    
    E1_s_bill_hdr_no =iS51131.S_MAINT_BL_INFO_SVR( gStrGlobalCollection,  iCommandSent, I1_s_bill_type_config_type , _
              I2_b_biz_partner_bp_cd , I3_s_bl_info , "" , I5_s_bill_hdr , I6_b_biz_partner , I7_b_biz_partner , _
              I8_b_sales_grp , I9_b_sales_grp )
	
	If CheckSYSTEMError(Err,True) = True Then
       'Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS51131 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS51131 = Nothing

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"           & vbCr

	If ConvSPChars(E1_s_bill_hdr_no) <>"" Then
	    Response.Write ".frm1.txtBLNo.value  = """ & ConvSPChars(E1_s_bill_hdr_no)     & """" & vbCr
	Else
	    Response.Write ".frm1.txtBLNo.value  = .frm1.txtBLNo1.value " & vbCr
	end if
    Response.Write ".DbSaveOk"                  & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr

  

End Sub
'============================================================================================================
Sub SubBizDelete()
	Dim iS51131
    Dim I6_s_bill_hdr 
    Const S500_I6_bill_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_bill_hdr
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear    
	Redim I6_s_bill_hdr(S500_I6_bill_no)
	I6_s_bill_hdr(S500_I6_bill_no) = Trim(Request("txtBLNo"))

    Set iS51131 = Server.CreateObject("PS7G131.cSBlInfoSvr")

	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

    Call iS51131.S_MAINT_BL_INFO_SVR(gStrGlobalCollection, "DELETE","","" , _
			                               "" , "" ,I6_s_bill_hdr , "" , _
                                           "" , ""  , ""  )

	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS51131 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS51131 = Nothing
	Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "Call parent.DbDeleteOk() " & vbCr
    Response.Write "</Script>"                  & vbCr
    
End Sub

Sub SubBizPostFlag()
	Dim iS7G115
	Dim itxtBLNo
 	On Error Resume Next
	Err.Clear 
    Set iS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")
   
	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	
	itxtBLNo = Trim(Request("txtBLNo"))
     pvCB = "F"
    Call iS7G115.S_POST_OPEN_AR_SVR(pvCB,gStrGlobalCollection,itxtBLNo)

	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS7G115 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS7G115 = Nothing 
	Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "parent.DbSaveOk()         " & vbCr
    Response.Write "</Script>"                  & vbCr
  
End Sub
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	
End Sub
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub

'============================================================================================================
Sub SubMakeSQLStatements(pMode)
    On Error Resume Next
    
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub

'============================================================================================================
Sub SetErrorStatus()
End Sub

'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub


%>

<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================================================================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBLInfo ( " & FilterVar(Request("txtBlNo"), "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , '" & Request("txtPrevNext") & "')"
	lgstrsql = strSelectList & strFromList
	
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
