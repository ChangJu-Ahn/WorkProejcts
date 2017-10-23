<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5311MB2
'*  4. Program Name         : 세금계산서등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : S51119LookupBillHdrSvr, S51139LookupBlInfoSvr
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/03/27
'*                            2001/12/19	Date표준적용 
'**********************************************************************************************
%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
	Dim strQueryMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

	Const lsN	= "N"			'매출채권						
	Const lsY	= "Y"			'b/l내역						

	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
 

	strQueryMode = Request("txtQueryMode")												'☜ : 현재 상태를 받음 

    Select Case strQueryMode
        Case CStr(lsN)																	'☜: Query
             Call SubBizQueryN()
        Case CStr(lsY)
             Call SubBizQueryY()
    End Select
'============================================
' Name : SubBizQueryN
' Desc : Query Data from Db
'============================================
Sub SubBizQueryN()
    Dim pS7G119
    Dim iCommandSent
        
    Dim I1_s_bill_hdr 
    Dim E1_a_gl 
    Dim E2_a_temp_gl 
    Dim E3_a_batch 
    Dim E4_s_bill_type_config 
    Dim E5_b_biz_partner 
    Dim E6_b_biz_area 
    Dim E7_b_minor 
    Dim E8_b_sales_org 
    Dim E9_b_biz_partner 
    Dim E10_b_biz_partner 
    Dim E11_b_biz_partner 
    Dim E12_b_minor 
    Dim E13_b_sales_grp 
    Dim E14_b_biz_partner 
    Dim E15_b_sales_org 
    Dim E16_b_sales_grp 
    Dim E17_s_bill_hdr 
    Dim E18_b_minor 

    Const S508_I1_bill_no = 0    'View Name : imp s_bill_hdr
    Const S508_I1_except_flag = 1

'   Const S508_E1_gl_no = 0    'View Name : exp a_gl

'   Const S508_E2_temp_gl_no = 0    'View Name : exp a_temp_gl

'   Const S508_E3_batch_no = 0    'View Name : exp a_batch

   Const S508_E4_bill_type = 0    'View Name : exp s_bill_type_config
   Const S508_E4_bill_type_nm = 1

   Const S508_E5_bp_cd = 0    'View Name : exp_sold_to_party b_biz_partner
   Const S508_E5_bp_nm = 1
   Const S508_E5_credit_rot_day = 2

'   Const S508_E6_biz_area_nm = 0    'View Name : exp b_biz_area

'   Const S508_E7_minor_nm = 0    'View Name : exp_pay_meth_nm b_minor

   Const S508_E8_sales_org = 0    'View Name : exp_billing b_sales_org
   Const S508_E8_sales_org_nm = 1

   Const S508_E9_bp_cd = 0    'View Name : exp_bill_to_party b_biz_partner
   Const S508_E9_bp_nm = 1

'   Const S508_E10_bp_nm = 0    'View Name : exp_beneficiary_nm b_biz_partner

'   Const S508_E11_bp_nm = 0    'View Name : exp_applicant_nm b_biz_partner

  ' Const S508_E12_minor_nm = 0    'View Name : exp_vat_type_nm b_minor

   Const S508_E13_sales_grp = 0    'View Name : exp_billing b_sales_grp
   Const S508_E13_sales_grp_nm = 1

   Const S508_E14_bp_cd = 0    'View Name : exp_payer b_biz_partner
   Const S508_E14_bp_nm = 1

   Const S508_E15_sales_org = 0    'View Name : exp_income b_sales_org
   Const S508_E15_sales_org_nm = 1

   Const S508_E16_sales_grp = 0    'View Name : exp_income b_sales_grp
   Const S508_E16_sales_grp_nm = 1

   Const S508_E17_bill_no = 0    'View Name : exp s_bill_hdr
   Const S508_E17_post_flag = 1
   Const S508_E17_trans_type = 2
   Const S508_E17_bill_dt = 3
   Const S508_E17_cur = 4
   Const S508_E17_xchg_rate = 5
   Const S508_E17_xchg_rate_op = 6
   Const S508_E17_bill_amt = 7
   Const S508_E17_vat_type = 8
   Const S508_E17_vat_rate = 9
   Const S508_E17_vat_amt = 10
   Const S508_E17_pay_meth = 11
   Const S508_E17_pay_dur = 12
   Const S508_E17_tax_bill_no = 13
   Const S508_E17_tax_prt_cnt = 14
   Const S508_E17_accept_fob_amt = 15
   Const S508_E17_beneficiary = 16
   Const S508_E17_applicant = 17
   Const S508_E17_remark = 18
   Const S508_E17_bill_amt_loc = 19
   Const S508_E17_vat_calc_type = 20
   Const S508_E17_vat_amt_loc = 21
   Const S508_E17_tax_biz_area = 22
   Const S508_E17_pay_type = 23
   Const S508_E17_pay_terms_txt = 24
   Const S508_E17_collect_amt = 25
   Const S508_E17_collect_amt_loc = 26
   Const S508_E17_income_plan_dt = 27
   Const S508_E17_nego_amt = 28
   Const S508_E17_so_no = 29
   Const S508_E17_lc_no = 30
   Const S508_E17_lc_doc_no = 31
   Const S508_E17_lc_amend_seq = 32
   Const S508_E17_bl_flag = 33
   Const S508_E17_biz_area = 34
   Const S508_E17_cost_cd = 35
   Const S508_E17_to_biz_area = 36
   Const S508_E17_to_cost_cd = 37
   Const S508_E17_ref_flag = 38
   Const S508_E17_sts = 39
   Const S508_E17_except_flag = 40
   Const S508_E17_reverse_flag = 41
   Const S508_E17_ext1_qty = 42
   Const S508_E17_ext2_qty = 43
   Const S508_E17_ext3_qty = 44
   Const S508_E17_ext1_amt = 45
   Const S508_E17_ext2_amt = 46
   Const S508_E17_ext3_amt = 47
   Const S508_E17_ext1_cd = 48
   Const S508_E17_ext2_cd = 49
   Const S508_E17_ext3_cd = 50
   Const S508_E17_vat_auto_flag = 51
   Const S508_E17_vat_inc_flag = 52
   Const S508_E17_deposit_amt = 53
   Const S508_E17_deposit_amt_loc = 54

'   Const S508_E18_minor_nm = 0    'View Name : exp_pay_type_nm b_minor

    On Error Resume Next
    Err.Clear 

	redim I1_s_bill_hdr(S508_I1_except_flag)

	iCommandSent = "QUERY"	
	I1_s_bill_hdr(S508_I1_bill_no)  = Trim(Request("txtBillNo"))

    Set pS7G119 = Server.CreateObject("PS7G119.cSLkBillHdrSvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  

    Call pS7G119.S_LOOKUP_BILL_HDR_SVR( gStrGlobalCollection, iCommandSent, I1_s_bill_hdr, _
							 E1_a_gl, E2_a_temp_gl, E3_a_batch, E4_s_bill_type_config, E5_b_biz_partner, _
							 E6_b_biz_area,  E7_b_minor, E8_b_sales_org, E9_b_biz_partner, E10_b_biz_partner, _
							 E11_b_biz_partner, E12_b_minor, E13_b_sales_grp, E14_b_biz_partner, E15_b_sales_org, _
							 E16_b_sales_grp, E17_s_bill_hdr,  E18_b_minor )
	If CheckSYSTEMError(Err,True) = True Then
       Set pS7G119 = Nothing
       Exit Sub
    End If  

    Set pS7G119 = Nothing

   
	'-----------------------
	'Result data display area
	'----------------------- 
 	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr
	
	Response.Write ".txtCurrency.Value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_cur))       & """" & vbCr
	Response.Write "parent.CurFormatNumericOCX" & vbCr

	Response.Write ".txtBillAmt.text			= 0 " & vbCr
	Response.Write ".txtVATAmt.text				= 0 " & vbCr

	Response.Write ".txtBillNo.Value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_bill_no))                  & """" & vbCr
	Response.Write ".txtBillToParty.Value		= """ & ConvSPChars(E9_b_biz_partner(S508_E9_bp_cd))              & """" & vbCr
    Response.Write ".txtBillToPartyNm.Value		= """ & ConvSPChars(E9_b_biz_partner(S508_E9_bp_nm))              & """" & vbCr
	Response.Write ".txtVATType.Value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_vat_type))              & """" & vbCr
    Response.Write ".txtVatTypeNm.Value			= """ & ConvSPChars(E12_b_minor)              & """" & vbCr
	Response.Write ".txtTaxBizAreacd.Value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_tax_biz_area))             & """" & vbCr
	Response.Write ".txtTaxBizAreaNm.Value		= """ & ConvSPChars(E6_b_biz_area)          & """" & vbCr
	Response.Write ".txtSalesGroup.Value		= """ & ConvSPChars(E13_b_sales_grp(S508_E13_sales_grp))          & """" & vbCr
    Response.Write ".txtSalesGroupNm.Value		= """ & ConvSPChars(E13_b_sales_grp(S508_E13_sales_grp_nm))       & """" & vbCr
	Response.Write ".txtVATRate.Text			= """ & UNINumClientFormat(E17_s_bill_hdr(S508_E17_vat_rate), ggExchRate.DecPoint, 0)              & """" & vbCr
	
		'VAT적용기준 
	If Trim(ConvSPChars(E17_s_bill_hdr(S508_E17_vat_calc_type))) = "1" Then
	    Response.Write ".rdoVATCalcType1.checked = True         "    & vbCr
	ElseIf Trim(ConvSPChars(E17_s_bill_hdr(S508_E17_vat_calc_type))) = "2" Then
	    Response.Write ".rdoVATCalcType2.checked = True         "    & vbCr
	End If

		'VAT포함여부 
	If TRIM (ConvSPChars(E17_s_bill_hdr(S508_E17_vat_inc_flag))) = "1" Then
	    Response.Write ".rdoVATIncFlag1.checked = True           "   & vbCr
	ElseIf TRIM (ConvSPChars(E17_s_bill_hdr(S508_E17_vat_inc_flag))) = "2" Then
	    Response.Write ".rdoVATIncFlag2.checked = True           "   & vbCr
	End If

		
	Response.Write "parent.BillQueryOk " & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr

End Sub
'============================================
' Name : SubBizQueryY
' Desc : Query Data from Db
'============================================
Sub SubBizQueryY()
    Dim iS7G139
    Dim iCommandSent

    Dim I1_s_bill_hdr
    Dim E1_a_gl
    Dim E2_a_temp_gl
    Dim E3_a_batch
    Dim E4_b_biz_partner
    Dim E5_s_bill_type_config
    Dim E6_b_biz_area
    Dim E7_b_biz_partner
    Dim E8_b_biz_partner
    Dim E9_b_minor
    Dim E10_b_minor
    Dim E11_s_bill_hdr
    Dim E12_b_biz_partner
    Dim E13_b_sales_org
    Dim E14_b_sales_grp
    Dim E15_b_sales_org
    Dim E16_b_sales_grp
    Dim E17_b_biz_partner
    Dim E18_s_bl_info
    Dim E19_b_biz_partner
    Dim E20_b_biz_partner
    Dim E21_b_biz_partner
    Dim E22_b_minor
    Dim E23_b_minor
    Dim E24_b_minor
    Dim E25_b_minor
    Dim E26_b_minor
    Dim E27_b_minor
    Dim E28_b_minor
    Dim E29_b_minor
    Dim E30_b_country
    Dim E31_b_country
    Dim E32_b_country

    'IMPORTS View 상수 
    
    'Const S528_I1_bill_no = 0    'View Name : imp s_bill_hdr

    'EXPORTS View 상수 

    Const S528_E1_gl_no = 0    'View Name : exp a_gl

    Const S528_E2_temp_gl_no = 0    'View Name : exp a_temp_gl

    Const S528_E3_batch_no = 0    'View Name : exp a_batch

    Const S528_E4_bp_cd = 0    'View Name : exp_sold_to_party b_biz_partner
    Const S528_E4_bp_nm = 1
    Const S528_E4_credit_rot_day = 2

    Const S528_E5_bill_type = 0    'View Name : exp s_bill_type_config
    Const S528_E5_bill_type_nm = 1

    Const S528_E6_biz_area_nm = 0    'View Name : exp b_biz_area

    Const S528_E7_bp_nm = 0    'View Name : exp_apllicant_nm b_biz_partner

    Const S528_E8_bp_nm = 0    'View Name : exp_beneficiary_nm b_biz_partner

    Const S528_E9_minor_nm = 0    'View Name : exp_vat_type_nm b_minor

    Const S528_E10_minor_nm = 0    'View Name : exp_pay_meth_nm b_minor

    Const S528_E11_bill_no = 0    'View Name : exp s_bill_hdr
    Const S528_E11_post_flag = 1
    Const S528_E11_trans_type = 2
    Const S528_E11_bill_dt = 3
    Const S528_E11_cur = 4
    Const S528_E11_xchg_rate = 5
    Const S528_E11_xchg_rate_op = 6
    Const S528_E11_bill_amt = 7
    Const S528_E11_vat_type = 8
    Const S528_E11_vat_rate = 9
    Const S528_E11_vat_amt = 10
    Const S528_E11_pay_meth = 11
    Const S528_E11_pay_dur = 12
    Const S528_E11_tax_bill_no = 13
    Const S528_E11_tax_prt_cnt = 14
    Const S528_E11_accept_fob_amt = 15
    Const S528_E11_beneficiary = 16
    Const S528_E11_applicant = 17
    Const S528_E11_remark = 18
    Const S528_E11_bill_amt_loc = 19
    Const S528_E11_vat_calc_type = 20
    Const S528_E11_vat_amt_loc = 21
    Const S528_E11_tax_biz_area = 22
    Const S528_E11_pay_type = 23
    Const S528_E11_pay_terms_txt = 24
    Const S528_E11_collect_amt = 25
    Const S528_E11_collect_amt_loc = 26
    Const S528_E11_income_plan_dt = 27
    Const S528_E11_nego_amt = 28
    Const S528_E11_so_no = 29
    Const S528_E11_lc_no = 30
    Const S528_E11_lc_doc_no = 31
    Const S528_E11_lc_amend_seq = 32
    Const S528_E11_bl_flag = 33
    Const S528_E11_biz_area = 34
    Const S528_E11_cost_cd = 35
    Const S528_E11_to_biz_area = 36
    Const S528_E11_to_cost_cd = 37
    Const S528_E11_ref_flag = 38
    Const S528_E11_sts = 39
    Const S528_E11_ext1_qty = 40
    Const S528_E11_ext2_qty = 41
    Const S528_E11_ext3_qty = 42
    Const S528_E11_ext1_amt = 43
    Const S528_E11_ext2_amt = 44
    Const S528_E11_ext3_amt = 45
    Const S528_E11_ext1_cd = 46
    Const S528_E11_ext2_cd = 47
    Const S528_E11_ext3_cd = 48
    Const S528_E11_vat_auto_flag = 49
    Const S528_E11_vat_inc_flag = 50

    Const S528_E12_bp_cd = 0    'View Name : exp_payer b_biz_partner
    Const S528_E12_bp_nm = 1

    Const S528_E13_sales_org = 0    'View Name : exp_income b_sales_org
    Const S528_E13_sales_org_nm = 1

    Const S528_E14_sales_grp = 0    'View Name : exp_income b_sales_grp
    Const S528_E14_sales_grp_nm = 1

    Const S528_E15_sales_org = 0    'View Name : exp_billing b_sales_org
    Const S528_E15_sales_org_nm = 1

    Const S528_E16_sales_grp = 0    'View Name : exp_billing b_sales_grp
    Const S528_E16_sales_grp_nm = 1

    Const S528_E17_bp_cd = 0    'View Name : exp_bill_to_party b_biz_partner
    Const S528_E17_bp_nm = 1

    Const S528_E18_bl_doc_no = 0    'View Name : exp s_bl_info
    Const S528_E18_ship_no = 1
    Const S528_E18_manufacturer = 2
    Const S528_E18_agent = 3
    Const S528_E18_receipt_plce = 4
    Const S528_E18_vessel_nm = 5
    Const S528_E18_voyage_no = 6
    Const S528_E18_forwarder = 7
    Const S528_E18_vessel_cntry = 8
    Const S528_E18_loading_port = 9
    Const S528_E18_dischge_port = 10
    Const S528_E18_delivery_plce = 11
    Const S528_E18_loading_plan_dt = 12
    Const S528_E18_latest_ship_dt = 13
    Const S528_E18_dischge_plan_dt = 14
    Const S528_E18_transport = 15
    Const S528_E18_tranship_cntry = 16
    Const S528_E18_tranship_dt = 17
    Const S528_E18_final_dest = 18
    Const S528_E18_incoterms = 19
    Const S528_E18_packing_type = 20
    Const S528_E18_tot_packing_cnt = 21
    Const S528_E18_container_cnt = 22
    Const S528_E18_packing_txt = 23
    Const S528_E18_gross_weight = 24
    Const S528_E18_weight_unit = 25
    Const S528_E18_gross_volumn = 26
    Const S528_E18_volumn_unit = 27
    Const S528_E18_freight = 28
    Const S528_E18_freight_plce = 29
    Const S528_E18_trans_price = 30
    Const S528_E18_trans_currency = 31
    Const S528_E18_trans_doc_amt = 32
    Const S528_E18_trans_xch_rate = 33
    Const S528_E18_trans_loc_amt = 34
    Const S528_E18_bl_issue_cnt = 35
    Const S528_E18_bl_issue_plce = 36
    Const S528_E18_bl_issue_dt = 37
    Const S528_E18_origin = 38
    Const S528_E18_origin_cntry = 39
    Const S528_E18_loading_dt = 40
    Const S528_E18_ext1_qty = 41
    Const S528_E18_ext2_qty = 42
    Const S528_E18_ext3_qty = 43
    Const S528_E18_ext1_amt = 44
    Const S528_E18_ext2_amt = 45
    Const S528_E18_ext3_amt = 46
    Const S528_E18_ext1_cd = 47
    Const S528_E18_ext2_cd = 48
    Const S528_E18_ext3_cd = 49

    Const S528_E19_bp_nm = 0    'View Name : exp_manufacturer_nm b_biz_partner

    Const S528_E20_bp_nm = 0    'View Name : exp_agent_nm b_biz_partner

    Const S528_E21_bp_nm = 0    'View Name : exp_forwarder_nm b_biz_partner

    Const S528_E22_minor_nm = 0    '  View Name : exp_loading_port_nm b_minor

    Const S528_E23_minor_nm = 0    'View Name : exp_discharge_port_nm b_minor

    Const S528_E24_minor_nm = 0    'View Name : exp_transport_nm b_minor

    Const S528_E25_minor_nm = 0    'View Name : exp_incoterms_nm b_minor

    Const S528_E26_minor_nm = 0    'View Name : exp_origin_nm b_minor

    Const S528_E27_minor_nm = 0    'View Name : exp_packing_type_nm b_minor

    Const S528_E28_minor_nm = 0    'View Name : exp_freight_nm b_minor

    Const S528_E29_minor_nm = 0    'View Name : exp_pay_type_nm b_minor

    Const S528_E30_country_nm = 0    'View Name : exp_vessel_cntry_nm b_country

    Const S528_E31_country_nm = 0    'View Name : exp_origin_cntry_nm b_country

    Const S528_E32_country_nm = 0    'View Name : exp_tranship_cntry_nm b_country

    iCommandSent = "LOOKUP"
    I1_s_bill_hdr = Trim(Request("txtBillNo"))
    Set iS7G139 = Server.CreateObject("PS7G139.cSLkInfoSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If    
           		
    Call iS7G139.S_BL_INFO_SVR( gStrGlobalCollection ,  iCommandSent ,   I1_s_bill_hdr ,  E1_a_gl , _
             E2_a_temp_gl ,  E3_a_batch ,  E4_b_biz_partner , E5_s_bill_type_config ,  E6_b_biz_area , _
             E7_b_biz_partner , E8_b_biz_partner ,  E9_b_minor ,  E10_b_minor , E11_s_bill_hdr , _
             E12_b_biz_partner ,  E13_b_sales_org , E14_b_sales_grp ,  E15_b_sales_org ,  E16_b_sales_grp , _
             E17_b_biz_partner ,  E18_s_bl_info ,  E19_b_biz_partner , E20_b_biz_partner ,  E21_b_biz_partner , _
             E22_b_minor ,  E23_b_minor ,  E24_b_minor ,  E25_b_minor , E26_b_minor ,  E27_b_minor ,  E28_b_minor , _
             E29_b_minor ,  E30_b_country ,  E31_b_country , E32_b_country )
           		
            				 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iS7G139 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If     				
    Set iS7G139 = Nothing  

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr

	'##### Rounding Logic #####
	'항상 거래화폐가 우선 
	Response.Write ".txtCurrency.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_cur))      & """" & vbCr
	Response.Write "parent.CurFormatNumericOCX" & vbCr

	Response.Write ".txtBillAmt.Text			= 0 " & vbCr				
	Response.Write ".txtVATAmt.text				= 0 " & vbCr				
	'##########################

	' Tab 1: 선적정보 1

	Response.Write ".txtBillToParty.value		= """ & ConvSPChars(E17_b_biz_partner(S528_E17_bp_cd))  & """" & vbCr		
	Response.Write ".txtBilltoPartyNm.value		= """ & ConvSPChars(E17_b_biz_partner(S528_E17_bp_nm))  & """" & vbCr				
	Response.Write ".txtTaxBizAreacd.value		= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_tax_biz_area))   & """" & vbCr					
	Response.Write ".txtTaxBizAreaNm.value		= """ & ConvSPChars(E6_b_biz_area(S528_E6_biz_area_nm))      & """" & vbCr		
	Response.Write ".txtVatType.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_vat_type))       & """" & vbCr
	Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(E9_b_minor(S528_E9_minor_nm))       & """" & vbCr
	
	If Trim(ConvSPChars(E11_s_bill_hdr(S528_E11_vat_calc_type))) = "1" Then
	    Response.Write ".rdoVATCalcType2.checked = True         "    & vbCr
	ElseIf Trim(ConvSPChars(E11_s_bill_hdr(S528_E11_vat_calc_type))) = "2" Then
	    Response.Write ".rdoVATCalcType1.checked = True         "    & vbCr
	End If

	Response.Write ".txtSalesGroup.value		= """ & ConvSPChars(E16_b_sales_grp(S528_E16_sales_grp))		& """" & vbCr
	Response.Write ".txtSalesGroupNm.value		= """ & ConvSPChars(E16_b_sales_grp(S528_E16_sales_grp_nm))		& """" & vbCr
	Response.Write ".txtVatRate.Text			= """ & UNINumClientFormat(E11_s_bill_hdr(S528_E11_vat_rate), ggAmtOfMoney.DecPoint, 0)      & """" & vbCr	

	Response.Write "parent.BillQueryOk "		& vbCr
	Response.Write "End With"					& vbCr
    Response.Write "</Script>"					& vbCr        
    
End Sub    

'============================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================
Sub SubBizQueryMulti()
    
End Sub    

'============================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================
Sub SubBizSaveMulti()        
    
End Sub    

'============================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

