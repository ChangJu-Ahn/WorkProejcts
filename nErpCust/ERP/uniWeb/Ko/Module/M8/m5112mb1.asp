<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%


'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5112mb1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : M51121(Maint)
'							  M51128(List)
'*  7. Modified date(First) : 2000/05/05
'*  8. Modified date(Last)  : 2001/09/12
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Ma Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/05/05 : ..........
'* 14. Business Logic of m5112ma1(매입상세등록)
'**********************************************************************************************
    Dim lgOpModeCRUD
    Dim lgCurrency
    Dim pvCB
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
    Call LoadBasisGlobalInf()


    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
		Case "Release"					    '☜: 확정,확정취소 요청을 받음 
			 Call SubRelease()
		Case "LookupUnitCost"
			 Call SubLookupUnitCost()			 
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iMax
	Dim PvArr
	Dim lGrpCnt
	Dim iTotstrData
	
	On Error Resume Next
    Err.Clear                                                               '☜: Protect system from crashing
    
	Dim lgStrPrevKey    ' 이전 값 
	'HEADER_____________________________________
	Dim iPM8G119
	Dim I1_m_iv_hdr_iv_no
    Dim E1_ief_supplied_flag
    Dim E2_m_config_process
    Dim E3_ief_supplied_flag
    Dim E4_b_biz_partner
    Dim E5_b_biz_partner
    Dim E6_b_biz_area_biz_area_nm
    Dim E7_b_biz_partner
    Dim E8_b_pur_grp
    Dim E9_m_iv_hdr
    Dim E10_m_iv_type
    Dim E11_b_minor
    Dim E12_b_minor
    Dim E13_b_minor
    Dim E14_b_currency
    Dim E15_m_pur_ord_hdr
    Dim E16_b_configuration
    
    'DETAIL___________________________________
    Dim M51128
	Dim I2_m_iv_dtl_iv_seq_no
	Dim EG1_exp_group
	Dim E1_m_iv_dtl	
		
	Dim iLngMaxRow
	Dim iLngRow
	Dim iStrPrevKey
	Dim istrData
	Dim istrTemp
    Dim iStrNextKey  	
	Dim iarrValue
	
	Const C_SHEETMAXROWS_D  = 100
	'HEADER_____________________________________
    Const M594_E2_po_type_cd = 0
    Const M594_E2_po_type_nm = 1
    
    Const M594_E10_iv_type_nm = 0		'exp m_iv_type
    Const M594_E10_iv_type_cd = 1
    Const M594_E10_import_flg = 2
    Const M594_E10_except_flg = 3
    Const M594_E10_ret_flg = 4

    Const M594_E9_iv_no = 0				'exp m_iv_hdr
    Const M594_E9_iv_dt = 1
    Const M594_E9_ap_post_dt = 2
    Const M594_E9_pay_dt = 3
    Const M594_E9_posted_flg = 4
    Const M594_E9_sppl_iv_no = 5
    Const M594_E9_payee_cd = 6
    Const M594_E9_build_cd = 7
    Const M594_E9_pur_org = 8
    Const M594_E9_iv_biz_area = 9
    Const M594_E9_tax_biz_area = 10
    Const M594_E9_iv_cost_cd = 11
    Const M594_E9_pay_meth = 12
    Const M594_E9_pay_dur = 13
    Const M594_E9_pay_terms_txt = 14
    Const M594_E9_pay_type = 15
    Const M594_E9_gross_doc_amt = 16
    Const M594_E9_gross_loc_amt = 17
    Const M594_E9_net_doc_amt = 18
    Const M594_E9_net_loc_amt = 19
    Const M594_E9_cash_doc_amt = 20
    Const M594_E9_cash_loc_amt = 21
    Const M594_E9_iv_cur = 22
    Const M594_E9_xch_rt = 23
    Const M594_E9_vat_type = 24
    Const M594_E9_vat_rt = 25
    Const M594_E9_tot_vat_doc_amt = 26
    Const M594_E9_tot_vat_loc_amt = 27
    Const M594_E9_tot_diff_doc_amt = 28
    Const M594_E9_tot_diff_loc_amt = 29
    Const M594_E9_pay_bank_cd = 30
    Const M594_E9_pay_acct_cd = 31
    Const M594_E9_pp_no = 32
    Const M594_E9_pp_doc_amt = 33
    Const M594_E9_pp_loc_amt = 34
    Const M594_E9_remark = 35
    Const M594_E9_loan_no = 36
    Const M594_E9_loan_doc_amt = 37
    Const M594_E9_loan_loc_amt = 38
    Const M594_E9_bl_no = 39
    Const M594_E9_bl_doc_no = 40
    Const M594_E9_lc_doc_no = 41
    Const M594_E9_ref_po_no = 42
    Const M594_E9_ext1_cd = 43
    Const M594_E9_gl_no = 44
    Const M594_E9_ext1_qty = 45
    Const M594_E9_ext1_amt = 46
    Const M594_E9_ext1_rt = 47
    Const M594_E9_ext1_dt = 48
    Const M594_E9_ext2_cd = 49
    Const M594_E9_ext2_qty = 50
    Const M594_E9_ext2_amt = 51
    Const M594_E9_ext2_rt = 52
    Const M594_E9_ext2_dt = 53
    Const M594_E9_ext3_cd = 54
    Const M594_E9_ext3_qty = 55
    Const M594_E9_ext3_amt = 56
    Const M594_E9_ext3_rt = 57
    Const M594_E9_ext3_dt = 58
    Const M594_E9_xch_rate_op = 59
    Const M594_E9_vat_inc_flag = 60

    Const M594_E8_pur_grp = 0			'exp b_pur_grp
    Const M594_E8_pur_grp_nm = 1

    Const M594_E7_bp_cd = 0				'exp b_biz_partner
    Const M594_E7_bp_rgst_no = 1
    Const M594_E7_bp_nm = 2

    Const M594_E15_rcpt_flg = 0			'exp m_pur_ord_hdr
    Const M594_E15_rcpt_type = 1
    Const M594_E15_issue_type = 2
    
    'DETAIL__________________________________________
    
	Const M590_EG1_E1_minor_nm = 0		'exp_item_vat_type b_minor
    Const M590_EG1_E2_po_no = 1			'exp_item m_pur_ord_hdr
    Const M590_EG1_E3_po_qty = 2		'exp_item m_pur_ord_dtl
    Const M590_EG1_E3_po_seq_no = 3
    Const M590_EG1_E3_po_prc = 4
    Const M590_EG1_E3_iv_qty = 5
    Const M590_EG1_E4_mvmt_no = 6		'exp_item m_pur_goods_mvmt
    Const M590_EG1_E4_mvmt_qty = 7
    Const M590_EG1_E4_gm_no = 8
    Const M590_EG1_E4_gm_year = 9
    Const M590_EG1_E4_gm_seq_no = 10
    Const M590_EG1_E4_gm_sub_seq_no = 11
    Const M590_EG1_E4_mvmt_rcpt_no = 12
    Const M590_EG1_E4_iv_qty = 13
    Const M590_EG1_E5_iv_seq_no = 14    'exp_item m_iv_dtl
    Const M590_EG1_E5_po_seq_no = 15
    Const M590_EG1_E5_iv_qty = 16
    Const M590_EG1_E5_iv_unit = 17
    Const M590_EG1_E5_iv_prc = 18
    Const M590_EG1_E5_iv_doc_amt = 19
    Const M590_EG1_E5_item_acct = 20
    Const M590_EG1_E5_iv_loc_amt = 21
    Const M590_EG1_E5_vat_type = 22
    Const M590_EG1_E5_vat_rt = 23
    Const M590_EG1_E5_vat_doc_amt = 24
    Const M590_EG1_E5_vat_loc_amt = 25
    Const M590_EG1_E5_diff_doc_amt = 26
    Const M590_EG1_E5_diff_loc_amt = 27
    Const M590_EG1_E5_po_no = 28
    Const M590_EG1_E5_ext1_cd = 29
    Const M590_EG1_E5_ext1_qty = 30
    Const M590_EG1_E5_ext1_amt = 31
    Const M590_EG1_E5_ext1_rt = 32
    Const M590_EG1_E5_ext2_cd = 33
    Const M590_EG1_E5_ext2_qty = 34
    Const M590_EG1_E5_ext2_amt = 35
    Const M590_EG1_E5_ext2_rt = 36
    Const M590_EG1_E5_ext3_cd = 37
    Const M590_EG1_E5_ext3_qty = 38
    Const M590_EG1_E5_ext3_amt = 39
    Const M590_EG1_E5_ext3_rt = 40
    Const M590_EG1_E5_tracking_no = 41
    Const M590_EG1_E5_vat_inc_flag = 42
    Const M590_EG1_E6_plant_cd = 43		'exp_item b_plant
    Const M590_EG1_E6_plant_nm = 44
    Const M590_EG1_E7_item_cd = 45		'exp_item b_item
    Const M590_EG1_E7_item_nm = 46
    Const M590_EG1_E7_spec = 47

    Const M590_EG1_E7_po_doc_amt = 48
    Const M590_EG1_E7_mvmt_doc_amt = 49
    Const M590_EG1_E7_amt_upt_flg = 50
    Const M590_EG1_E7_prc_ctrl_flg = 51
    Const M590_EG1_E7_vat_doc_amt = 52
    Const M590_EG1_E7_ret_flg = 53
    Const M590_EG1_E7_vat_amt_rvs_flg = 54
    
    Const M590_EG1_E7_m_num_wks_doc_amt = 55
    Const M590_EG1_E7_m_num_wks_vat_doc_amt = 56
    Const M590_EG1_E7_po_vat_inc_flag = 57
    '#후LC추가(2003.03.17)-Lee, Eun Hee
    Const M590_EG1_E8_after_lc_no = 58
    Const M590_EG1_E8_after_lc_seq_no = 59
    Const M590_EG1_E8_lc_flg = 60
    '#매입내역별로 환율 관리 (2003.09.21)
    Const M590_EG1_E8_xch_rt = 61
    '비고추가 (2005.12.19)
    Const M590_EG1_E8_remark = 62
    
	
    if Request("lgStrPrevKey") = "" then
		lgStrPrevKey = 0
    else
		lgStrPrevKey = CLng(Request("lgStrPrevKey"))
    end if
	
	'**수정(2003.06.11)-추가 100건을 조회시는 헤더는 조회하지 않는다**
    If Request("txtMaxRows") = 0 Then
    
    Set iPM8G119 = Server.CreateObject("PM8G119.cMLookupIvHdrS")    
    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPM8G119 = Nothing
		Exit Sub
	End If

    I1_m_iv_hdr_iv_no		= Request("txtIvNo")

    CALL iPM8G119.M_LOOKUP_IV_HDR_SVR(gStrGlobalCollection,cstr(I1_m_iv_hdr_iv_no), _
            E1_ief_supplied_flag, E2_m_config_process,E3_ief_supplied_flag,E4_b_biz_partner,E5_b_biz_partner, _
            E6_b_biz_area_biz_area_nm,E7_b_biz_partner,E8_b_pur_grp,E9_m_iv_hdr,E10_m_iv_type,E11_b_minor, _
            E12_b_minor,E13_b_minor,E14_b_currency,E15_m_pur_ord_hdr, E16_b_configuration)
            
	If CheckSYSTEMError(Err,True) = True Then
		Set iPM8G119 = Nothing
		Exit Sub
	End If            
	
	Set iPM8G119 = Nothing

	lgCurrency = E9_m_iv_hdr(M594_E9_iv_cur)
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.write " if .frm1.vspdData.MaxRows = 0 then	 " & vbCr	
		'##### Rounding Logic #####		<=========================================
		'항상 거래화폐가 우선 
	Response.Write "	.frm1.txtCur.value = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur))              & """" & vbCr
	Response.Write "    .CurFormatNumericOCX " & vbCr
	Response.Write "	.frm1.txtIvTypeCd.value   = """ & ConvSPChars(E10_m_iv_type(M594_E10_iv_type_cd))        & """" & vbCr
	Response.Write "	.frm1.txtIvTypeNm.value   = """ & ConvSPChars(E10_m_iv_type(M594_E10_iv_type_nm))        & """" & vbCr
	Response.Write "	.frm1.txtIvDt.text   = """ & UNIDateClientFormat(E9_m_iv_hdr(M594_E9_iv_dt))        & """" & vbCr
	Response.Write "	.frm1.txtSpplCd.value = """ & ConvSPChars(E7_b_biz_partner(M594_E7_bp_cd))              & """" & vbCr
	Response.Write "	.frm1.txtSpplNm.value = """ & ConvSPChars(E7_b_biz_partner(M594_E7_bp_nm))              & """" & vbCr
	
	Response.Write "	.frm1.txtGrpCd.value     = """ & ConvSPChars(E8_b_pur_grp(M594_E8_pur_grp))                & """" & vbCr
	Response.Write "	.frm1.txtGrpNm.value     = """ & ConvSPChars(E8_b_pur_grp(M594_E8_pur_grp_nm))              & """" & vbCr
	Response.Write "	.frm1.txtIvAmt.text		 = """ & UNIConvNumDBToCompanyByCurrency(E9_m_iv_hdr(M594_E9_gross_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)),ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write "	.frm1.txtnetAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E9_m_iv_hdr(M594_E9_net_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)), ggAmtOfMoneyNo,"X","X") & """" & vbCr '13차추가 
'	Response.Write "	.frm1.txtvatAmt.text     = """ & UNINumClientFormatByTax(E9_m_iv_hdr(M594_E9_tot_vat_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)),ggAmtOfMoneyNo)     & """" & vbCr	
	Response.Write "	.frm1.txtvatAmt.text     = """ & UniConvNumberDBToCompany(E9_m_iv_hdr(M594_E9_tot_vat_doc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & """" & vbCr'13차추가 
	
'	Response.Write "	.frm1.txtIvAmt.value   = """ & UNINumClientFormatByCurrency(E9_m_iv_hdr(M594_E9_gross_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)),ggAmtOfMoneyNo)      & """" & vbCr
'	Response.Write "	.frm1.txtnetAmt.value  = """ & UNINumClientFormatByCurrency(E9_m_iv_hdr(M594_E9_net_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)),ggAmtOfMoneyNo)    & """" & vbCr
'	Response.Write "	.frm1.txtvatAmt.value  = """ & UNINumClientFormatByTax(E9_m_iv_hdr(M594_E9_tot_vat_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)),ggAmtOfMoneyNo)     & """" & vbCr
	
	'Response.Write "	.frm1.txtXchRt.value   = """ & UNINumClientFormat(E15_m_pur_ord_hdr(XchRt,ggExchRate.DecPoint,0)   & """" & vbCr
	'Response.Write "	.frm1.txtCurNm.value   = """ & ConvSPChars(E14_b_currency)     & """" & vbCr
	

	Response.Write "	.frm1.hdnGlNo.value    = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_gl_no)) & """" & vbCr
	Response.Write "	.frm1.hdnGlType.value    = """ & ConvSPChars(E3_ief_supplied_flag) & """" & vbCr
	'------------------------
	Response.Write "	.frm1.hdnSppl.value    = """ & ConvSPChars(E7_b_biz_partner(M594_E7_bp_cd)) & """" & vbCr
	Response.Write "	.frm1.hdnGrp.value    = """ & ConvSPChars(E8_b_pur_grp(M594_E8_pur_grp)) & """" & vbCr
	Response.Write "	.frm1.hdnPostingFlg.value    = """ & UCase(ConvSPChars(E9_m_iv_hdr(M594_E9_posted_flg))) & """" & vbCr
	Response.Write "	.frm1.hdnIvType.value    = """ & ConvSPChars(E10_m_iv_type(M594_E10_iv_type_cd)) & """" & vbCr
	Response.Write "	.frm1.hdnIvTypeNm.value    = """ & ConvSPChars(E10_m_iv_type(M594_E10_iv_type_nm)) & """" & vbCr
	Response.Write "	.frm1.hdnMvmtType.value    = """" " & vbCr
	Response.Write "	.frm1.hdnImportflg.value    = """ & ConvSPChars(E10_m_iv_type(M594_E10_import_flg)) & """" & vbCr
	Response.Write "	.frm1.hdnExceptflg.value    = """ & ConvSPChars(E10_m_iv_type(M594_E10_except_flg)) & """" & vbCr
	Response.Write "	.frm1.hdnRetflg.value    = """ & ConvSPChars(E10_m_iv_type(M594_E10_ret_flg)) & """" & vbCr
	Response.Write "	.frm1.hdnVatRt.value    = """ & UNINumClientFormat(E9_m_iv_hdr(M594_E9_vat_rt), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write "	.frm1.hdnXch.value    = """ & UNINumClientFormat(E9_m_iv_hdr(M594_E9_xch_rt), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write "	.frm1.txtXchRt.text    = """ & UNINumClientFormat(E9_m_iv_hdr(M594_E9_xch_rt), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write "	.frm1.hdnDiv.value    = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_xch_rate_op)) & """" & vbCr '환율연산자 
	
	Response.Write "	.frm1.hdnPONo.value    = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_ref_po_no)) & """" & vbCr
	
	Response.Write "If """ & ConvSPChars(Trim(E1_ief_supplied_flag)) & """ = ""Y"" Then " & vbCr
	Response.Write "	.frm1.ChkPrepay.checked = true " & vbCr
	Response.Write "Else " & vbCr
	Response.Write "	.frm1.ChkPrepay.checked = False " & vbCr
	Response.Write "End If " & vbCr
	
	Response.Write "If """ & ConvSPChars(E9_m_iv_hdr(M594_E9_vat_inc_flag)) & """ = ""2"" Then " & vbCr
						'.frm1.rdoVatFlg2.Checked= true 
	Response.Write "	.frm1.hdvatFlg.value 	= ""2"" " & vbCr
	Response.Write "Else " & vbCr
						'.frm1.rdoVatFlg1.Checked= true 
	Response.Write "	.frm1.hdvatFlg.value 	= ""1"" " & vbCr
	Response.Write "End If " & vbCr
	
	Response.Write "	.frm1.hdnPostDt.value    = """ & UNIDateClientFormat(E9_m_iv_hdr(M594_E9_ap_post_dt)) & """" & vbCr
	Response.Write "	.frm1.txtIvNo.value    = """ & ConvSPChars(Request("txtIvNo")) & """" & vbCr
	Response.Write "	.frm1.hdnVatType.value    = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_vat_type)) & """" & vbCr
	Response.Write "	.frm1.hdnRcptFlg.value    = """ & ConvSPChars(E15_m_pur_ord_hdr(M594_E15_rcpt_flg)) & """" & vbCr
	Response.Write "	.frm1.hdnRcptType.value    = """ & ConvSPChars(E15_m_pur_ord_hdr(M594_E15_rcpt_type)) & """" & vbCr
	Response.Write "	.frm1.hdnIssueType.value    = """ & ConvSPChars(E15_m_pur_ord_hdr(M594_E15_issue_type)) & """" & vbCr
	Response.Write "	.frm1.hdnPoTypeCd.value    = """ & ConvSPChars(E2_m_config_process(M594_E2_po_type_cd)) & """" & vbCr
	Response.Write "	.frm1.hdnPoTypeNm.value    = """ & ConvSPChars(E2_m_config_process(M594_E2_po_type_nm)) & """" & vbCr
	'Local LC후 입출고참조한 경우 매입의 결제방법이 LC의 결제방법과 동일한 건만 조회되도록 변경 
	'LC_kind, Pay_Meth 추가함 (2003.03.24)
	Response.Write "	.frm1.hdnLcKind.value   = """ & ConvSPChars(E16_b_configuration)     & """" & vbCr
	Response.Write "	.frm1.hdnPayMeth.value   = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_pay_meth))     & """" & vbCr
	
	Response.Write "If UCase(gCurrency) = UCase(Trim(.frm1.txtCur.value)) Then " & vbCr  '화폐가 KRW 인경우는 매입자국금액,VAT자국금액 히든 
	Response.Write "	.frm1.vspdData.Col = .C_IvLocAmt	: .frm1.vspdData.ColHidden = True " & vbCr
	Response.Write "	.frm1.vspdData.Col = .C_NetLocAmt	: .frm1.vspdData.ColHidden = True " & vbCr
	Response.Write "	.frm1.vspdData.Col = .C_VatLocAmt	: .frm1.vspdData.ColHidden = True " & vbCr
	Response.Write "Else " & vbCr
	Response.Write "	.frm1.vspdData.Col = .C_IvLocAmt	: .frm1.vspdData.ColHidden = False " & vbCr
	Response.Write "	.frm1.vspdData.Col = .C_NetLocAmt	: .frm1.vspdData.ColHidden = False " & vbCr
	Response.Write "	.frm1.vspdData.Col = .C_VatLocAmt	: .frm1.vspdData.ColHidden = False " & vbCr
	Response.Write "End If " & vbCr
		'hdnRetflg(반품) ,hdnExceptflg(예외)인경우 
	'수정(2003.03.06)이은희 
	Response.Write "If Trim(UCase(.frm1.hdnRetflg.value)) = ""Y"" or Trim(UCase(.frm1.hdnExceptflg.Value)) = ""Y"" Then " & vbCr 
	Response.Write "   If  Trim(.lgStrPrevKey) = """" Then " & vbCr 
	Response.Write "       .frm1.vspdData.Col = .C_NetLocAmt	: .frm1.vspdData.ColHidden = True " & vbCr
	Response.Write "   End If " & vbCr
	Response.Write "End If " & vbCr
	
	Response.Write "	.frm1.txthdnIvNo.value    = """ & ConvSPChars(Request("txtIvNo")) & """" & vbCr
	
	'2003.03 KJH 전표번호 가져오는 로직 수정 
	Response.Write "	.SubGetGlNo" & vbCr
			
	Response.Write "    .CurFormatNumSprSheet " & vbCr
	Response.Write " End If " & vbCr
	Response.Write "End With " & vbCr
	Response.Write "</Script> " & vbCr
	
	Else
		lgCurrency = request("txtCurrency")
	End If	
	
	'DETAIL________________________________________________________________________												
    Set M51128 = Server.CreateObject("PM8G128.cMListIvDtlS")
		
	If CheckSYSTEMError(Err,True) = True Then
		Set M51128 = Nothing
		Exit Sub
	End If					
	
	I1_m_iv_hdr_iv_no			= Request("txtIvNo")
    if Trim(lgStrPrevKey) <> "" then
		I2_m_iv_dtl_iv_seq_no = lgStrPrevKey
	End if

	CALL M51128.M_LIST_IV_DTL_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D,cstr(I1_m_iv_hdr_iv_no),cstr(I2_m_iv_dtl_iv_seq_no), _
            EG1_exp_group, E1_m_iv_dtl)


	if Request("txtQuerytype") <> "Query" And Err.Description = "B_MESSAGE175200" then
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			IF UBound(EG1_exp_group,1) <= 0 Then
				Set M51128 = Nothing
				Exit Sub												'☜: ComProxy Unload	
			End If
		
	Else 
		If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set M51128 = Nothing												'☜: ComProxy Unload
			'Detail항목이 없을 경우 Header정보만 보여줌 
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
			Exit Sub															'☜: 비지니스 로직 처리를 종료함 
		End If
		
	End if
			
	Set M51128 = Nothing
				
	iLngMaxRow = CLng(Request("txtMaxRows"))
	iMax = UBound(EG1_exp_group,1)
	ReDim PvArr(iMax)
	
	lGrpCnt = 0
	
	For iLngRow = 0 To UBound(EG1_exp_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E5_iv_seq_no)) 
           Exit For
        End If  	
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E6_plant_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E6_plant_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_item_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_item_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_spec))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E5_iv_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E5_iv_unit))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_iv_prc), lgCurrency, ggUnitCostNo,"X","X")

'        <!--거래금액(거래금액,tot금액) 과 Net Amount-->
		if EG1_exp_group(iLngRow,M590_EG1_E5_vat_inc_flag) = "2" then
					istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(ccur(EG1_exp_group(iLngRow,M590_EG1_E5_iv_doc_amt)) + ccur(EG1_exp_group(iLngRow,M590_EG1_E5_vat_doc_amt)), lgCurrency,ggAmtOfMoneyNo,"X","X")
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_iv_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_iv_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")
		else
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_iv_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_iv_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")
			istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_iv_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")
		End if

		if EG1_exp_group(iLngRow,M590_EG1_E5_vat_inc_flag) = "2" then
			istrData = istrData & Chr(11) & "포함"
'			<!--포함구분코드 -->
			istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E5_vat_inc_flag))
		else 
			istrData = istrData & Chr(11) & "별도"
'			<!--포함구분코드 -->
			istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E5_vat_inc_flag))
		End if
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E5_vat_type))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E1_minor_nm)) 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E5_vat_rt),ggExchRate.DecPoint,0) 'vat율 
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_vat_doc_amt), lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo,"X")      
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_vat_doc_amt), lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo,"X")      
        
        if EG1_exp_group(iLngRow,M590_EG1_E5_vat_inc_flag) = "2" then
			istrData = istrData & Chr(11) & UniConvNumberDBToCompany(ccur(EG1_exp_group(iLngRow,M590_EG1_E5_iv_loc_amt)) + ccur(EG1_exp_group(iLngRow,M590_EG1_E5_vat_loc_amt)), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
			istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_group(iLngRow,M590_EG1_E5_iv_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
		else
			istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_group(iLngRow,M590_EG1_E5_iv_loc_amt),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
			istrData = istrData & Chr(11) & UniConvNumberDBToCompany(EG1_exp_group(iLngRow,M590_EG1_E5_iv_loc_amt),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
		End if
'        <!-- istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E5_vat_loc_amt), ggAmtOfMoney.DecPoint, 0) -->
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_vat_loc_amt), gCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo,"X")

		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E8_remark)) '비고 

		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E3_po_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E3_po_prc),lgCurrency, ggUnitCostNo,"X","X")
        
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E4_mvmt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E3_iv_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E5_po_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E3_po_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E4_mvmt_rcpt_no)) '입고번호 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E4_gm_no)) '재고처리번호 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E4_gm_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E5_iv_seq_no))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E5_iv_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E4_mvmt_no))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E4_iv_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E5_iv_qty),ggQty.DecPoint,0)
        
        '==========================================================================================================================================
		'10월패치 추가 - 10/10
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_vat_amt_rvs_flg))
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E7_vat_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")
        istrData = istrData & Chr(11) & " "
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M590_EG1_E5_tracking_no))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_iv_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")			        
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E5_vat_doc_amt), lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo,"X")            		

		'==========================================================================================================================================
		'10월패치 추가 - 10/10
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E7_po_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")					
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E7_mvmt_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")					
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E7_m_num_wks_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")					
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_amt_upt_flg))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_prc_ctrl_flg))
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E7_vat_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")      		      
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,M590_EG1_E7_m_num_wks_vat_doc_amt), lgCurrency, ggAmtOfMoneyNo,"X","X")      		      
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E5_iv_qty),ggQty.DecPoint,0)
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_ret_flg))
		'==========================================================================================================================================
		
		istrData = istrData & Chr(11) & " "
		istrData = istrData & Chr(11) & " "
		'--추가(2003.02.18)------------
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E7_po_vat_inc_flag))
		'#후LC추가(2003.03.17)-Lee, Eun Hee
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E8_after_lc_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E8_after_lc_seq_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,M590_EG1_E8_lc_flg))
		'#매입내역별로 환율 관리 (2003.09.21)
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow,M590_EG1_E8_xch_rt),ggExchRate.DecPoint,0) '환율 
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
        
        PvArr(lGrpCnt) = istrData
        lGrpCnt = lGrpCnt + 1
        istrData = ""
    Next
    
    iTotstrData = Join(PvArr, "")

	Response.Write "<Script language=vbs> " & vbCr  
	Response.Write "With parent" & vbCr 
    
    Response.Write " .ggoSpread.Source   = .frm1.vspdData  " & vbCr
    Response.Write  "    .frm1.vspdData.Redraw = False   "                         & vbCr      
    Response.Write " .ggoSpread.SSShowData  """ & iTotstrData	& """ , ""F""" & vbCr	
    Response.Write " .lgStrPrevKey   = """ & iStrNextKey & """" & vbCr  
    
    'Response.Write " .frm1.txthdnIvNo.value = """ & Request("txtIvNo") & """" & vbCr
    
	Response.Write "       if Trim(UCase(.frm1.hdnPostingflg.Value)) <> ""Y"" then " & vbCr
	Response.Write "			.QueryAtSetSpreadColor(1) " & vbCr		 
	Response.Write "       End IF " & vbCr
	Response.Write "       .DbQueryOk() " & vbCr 
	Response.Write  "    .frm1.vspdData.Redraw = True   " & vbCr   
	Response.Write "End With				" & vbCr

    Response.Write "</Script> "		

End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear 
    
	Dim iPM8G121
    Dim iErrorPosition
    Dim iUpdtUserId, ihdnIvNo, itxtSpread
    '-------------------
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii

    Dim iCUCount
    Dim iDCount
    
    Call CheckNoForDT(Trim(Request("txthdnIvNo")))
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")
    '---------------------         
             
	Set iPM8G121 = Server.CreateObject("PM8G121.cMMaintIvDtlS")    

	If CheckSYSTEMError(Err,True) = true then 		
		Set iPM8G121 = Nothing
		Exit Sub
	End If
	
	iUpdtUserId = gUsrID
	ihdnIvNo = Trim(Request("txthdnIvNo"))
	
	pvCB = "F"

	Call iPM8G121.M_MAINT_IV_DTL_SVR(pvCB,gStrGlobalCollection, gCurrency, _
	                                 iUpdtUserId, ihdnIvNo, _
	                                 itxtSpread, iErrorPosition)

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		Set iPM8G121 = Nothing
		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "</Script> "
		Exit Sub
	End If

    Set iPM8G121 = Nothing    
                 
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.frm1.txtIvNo.Value = """ & ConvSPChars(Request("txthdnIvNo")) & """" & vbCr    
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "            
End Sub    
'============================================================================================================
' Name : SubRelease
' Desc : 확정요청받을때 
'============================================================================================================
Sub SubRelease()

    On Error Resume Next
    Err.Clear                                                                       '☜: Protect system from crashing

    Dim I2_ief_supplied
    Dim IG1_imp_dtl_group
    Dim I3_m_batch_ap_post_wks
    
    Const M557_IG1_I1_count = 0    '  View Name : imp_dtl_no ief_supplied
    Const M557_IG1_I2_iv_no = 1    '  View Name : imp_dtl m_batch_ap_post_wks
    Const M557_IG1_I3_ap_dt = 2
    
    Const M557_I3_ap_dt_type = 0    '  View Name : imp_hdr m_batch_ap_post_wks
    Const M557_I3_ap_dt = 1
    Const M557_I3_import_flg = 2

    Dim PM8G211
    Dim lgIntFlgMode

    If Len(Trim(Request("txtPostDt"))) Then
        If UNIConvDate(Request("txtPostDt")) = "" Then
            Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
            Call LoadTab("parent.frm1.txtPostDt", 0, I_MKSCRIPT)
            Exit Sub
        End If
    End If
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                      '☜: 저장시 Create/Update 판별 
    '수정(2003.06.09)    
    ReDim IG1_imp_dtl_group(0, 2)
    
	IG1_imp_dtl_group(0, M557_IG1_I2_iv_no)  = Trim(Request("txtIvNo"))
    IG1_imp_dtl_group(0, M557_IG1_I1_count) = 1
    IG1_imp_dtl_group(0, M557_IG1_I3_ap_dt) = UNIConvDate(Request("hdnPostDt"))

	ReDim I3_m_batch_ap_post_wks(2)
	I3_m_batch_ap_post_wks(M557_I3_ap_dt_type) = ""
    'I3_m_batch_ap_post_wks(M557_I3_ap_dt) = UNIConvDate(Request("hdnPostDt"))
    I3_m_batch_ap_post_wks(M557_I3_import_flg) = Trim(Request("hdnImportFlg"))
    
    
    If Request("hdnPostingFlg") = "Y" Then
        I2_ief_supplied = "N"
    Else
        I2_ief_supplied = "Y"
    End If
	
   Set PM8G211 = Server.CreateObject("PM8G211.cMPostApS")

   If CheckSYSTEMError(Err, True) = True Then
        Set PM8G211 = Nothing
	    Exit Sub
    End If
    
	pvCB = "F"
	
    Call PM8G211.M_POST_AP_SVR(pvCB, gStrGlobalCollection,I2_ief_supplied, IG1_imp_dtl_group, I3_m_batch_ap_post_wks)

    If CheckSYSTEMError(Err, True) = True Then
        Set PM8G211 = Nothing
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.btnPosting.disabled = False" & vbCr
			Response.Write "</Script>" & vbCr
        Exit Sub
    End If
    '-----------------------
    'Result data display area
    '-----------------------
    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.MainQuery() " & vbCr                                '☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> " & vbCr

    Set PM8G211 = Nothing                                                   '☜: Unload Comproxy

End Sub

'============================================================================================================
' Name : SubLookupUnitCost
' Desc : 확정요청받을때 
'============================================================================================================
Sub SubLookupUnitCost()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Dim arrParam, Unit
	Dim iPM3G1P9, iPB3S106

	'iPB3S106의 변수들 
    Dim E1_b_pur_org
    Dim E2_b_item_group
    Dim E3_for_issued_b_storage_location
    Dim E4_for_major_b_storage_location
    Dim E5_i_material_valuation
    Dim E6_b_item_by_plant
    Const P003_E6_order_unit_pur = 35
    Dim E7_b_item
    Const P003_E7_item_cd = 0
    Const P003_E7_item_nm = 1
    Const P003_E7_formal_nm = 2
    Const P003_E7_spec = 3
    
    Dim E8_b_plant
	Const P003_E8_plant_cd = 0
    Const P003_E8_plant_nm = 1
    
	'iPM3G1P9의 변수들 
	Dim I1_m_supplier_item_price
    Const M106_I1_pur_unit = 0    'View Name : imp m_supplier_item_price
    Const M106_I1_pur_cur = 1
    Const M106_I1_valid_fr_dt = 2
    Dim I2_b_biz_partner_bp_cd
	Dim I3_b_item_item_cd
	Dim I4_b_plant_plant_cd
	Dim E1_m_supplier_item_price_pur_prc
	Dim E2_b_item
	Dim E3_b_plant
	Dim E4_b_storage_location
	Dim E5_b_hs_code
	Dim E6_m_supplier_item_by_plant
    Const M106_E6_pur_priority = 0    'View Name : exp m_supplier_item_by_plant
    Const M106_E6_pur_unit = 1
	Dim E7_b_minor_vat
    
    Redim I1_m_supplier_item_price(2)
	
	
	arrParam = Split(Trim(request("txtSpread")), gColSep)
	
	I4_b_plant_plant_cd		= Trim(arrParam(1))
	I3_b_item_item_cd		= Trim(arrParam(2))
	

	Set iPB3S106 = Server.CreateObject("PB3S106.cBLkUpItemByPlt") 
	Set iPM3G1P9 = Server.CreateObject("PM3G1P9.cMLookupPriceS")
		
	If CheckSYSTEMError(Err,True) = true then
		Set iPB3S106 = Nothing
		Set iPM3G1P9 = Nothing 		
		Exit Sub
	End If
	
	'-----------------------
	' Find Unit of ItemByPlant
	'-----------------------
	Call iPB3S106.B_LOOK_UP_ITEM_BY_PLANT(gStrGlobalCollection, I4_b_plant_plant_cd, I3_b_item_item_cd, _
											E1_b_pur_org, E2_b_item_group,E3_for_issued_b_storage_location, _
											E4_for_major_b_storage_location, E5_i_material_valuation, _
											E6_b_item_by_plant, E7_b_item, E8_b_plant)
	
	
	Err.Clear
	If CheckSYSTEMError(Err,True) = true Then
		Set iPB3S106 = Nothing 		
		Exit Sub
	End If

	Set iPB3S106 = Nothing

	Unit = Trim(ConvSPChars(E6_b_item_by_plant(P003_E6_order_unit_pur)))
	
	I2_b_biz_partner_bp_cd	= Trim(arrParam(4))
	I1_m_supplier_item_price(M106_I1_pur_unit)		= Unit
	I1_m_supplier_item_price(M106_I1_pur_cur)		= Trim(arrParam(5))
	I1_m_supplier_item_price(M106_I1_valid_fr_dt)	= UNIConvDate(Trim(arrParam(6)))
	'-----------------------
	' Find Cost and Unit of SupplierItem
	'-----------------------
	Call iPM3G1P9.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, _
									   I1_m_supplier_item_price, I2_b_biz_partner_bp_cd, _
									   I3_b_item_item_cd, I4_b_plant_plant_cd, _
									   E1_m_supplier_item_price_pur_prc, E2_b_item, _
									   E3_b_plant, E4_b_storage_location, E5_b_hs_code, _
									   E6_m_supplier_item_by_plant, E7_b_minor_vat)
	
	Err.Clear
	If CheckSYSTEMError(Err,True) = true Then
		Set iPM3G1P9 = Nothing 		
		Exit Sub
	End If

	
	
	If Trim(ConvSPChars(E6_m_supplier_item_by_plant(M106_E6_pur_unit))) <> "" Then
		
		Unit = Trim(ConvSPChars(E6_m_supplier_item_by_plant(M106_E6_pur_unit)))
		
		I1_m_supplier_item_price(M106_I1_pur_unit)		= Unit
		
		Call iPM3G1P9.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, _
									   I1_m_supplier_item_price, I2_b_biz_partner_bp_cd, _
									   I3_b_item_item_cd, I4_b_plant_plant_cd, _
									   E1_m_supplier_item_price_pur_prc, E2_b_item, _
									   E3_b_plant, E4_b_storage_location, E5_b_hs_code, _
									   E6_m_supplier_item_by_plant, E7_b_minor_vat)
											
		If CheckSYSTEMError(Err,True) = True Then
			Set iPM3G1P9 = Nothing 		
			Exit Sub
		End If
		
	End If
	
	Set iPM3G1P9 = Nothing
	
  
		Response.Write "<Script language=vbs> " & vbCr         
		Response.Write " With Parent.frm1.vspdData "      	& vbCr
		Response.Write " .Row  =  " & arrParam(0)   		& vbCr  '이부분 주의... 처리요!!!
		
		Response.Write " .Col 	= Parent.C_PlantNm "    			& vbCr
		Response.Write " .text  = """ & ConvSPChars(E8_b_plant(P003_E8_plant_nm)) & """" & vbCr
		Response.Write " .Col 	= Parent.C_ItemNm "    			& vbCr
		Response.Write " .text  = """ & ConvSPChars(E7_b_item(P003_E7_item_nm)) & """" & vbCr
		Response.Write " .Col 	= Parent.C_SpplSpec "    			& vbCr
		Response.Write " .text  = """ & ConvSPChars(E7_b_item(P003_E7_spec)) & """" & vbCr
		Response.Write " .Col 	= Parent.C_Unit "    			& vbCr
		Response.Write " .text  = """ & Unit & """" & vbCr
		Response.Write " .Col 	= Parent.C_Cost "       	& vbCr
		Response.Write " .text   = """ & UNINumClientFormat(E1_m_supplier_item_price_pur_prc(0),ggUnitCost.DecPoint,0) & """" & vbCr
			
		Response.Write " End With "             & vbCr		
		Response.Write "</Script> "             & vbCr 

End Sub


Sub CheckNoForDT(ByVal I1_no)
	
	'For 전자세금계산서 
	Dim pvObjRs
	Dim pvStrSQL
	
	Dim IntRetCD
    Dim lgObjRs
    Dim lgObjConn
    Dim lgObjComm
    Dim lgStrSQL

	Call SubOpenDB(lgObjConn) 
	
	pvStrSQL = " SELECT REFERENCE FROM B_CONFIGURATION WHERE MAJOR_CD = 'DT004' AND MINOR_CD = 'A' AND REFERENCE = 'Y' "

	If 	FncOpenRs("R",lgObjConn,pvObjRs,pvStrSQL,"X","X") = True Then
			
		Call SubCreateCommandObject(lgObjComm)
    
		With lgObjComm      
		            
		    .CommandText = "usp_dt_check_posting_status_tax"
		    .CommandType = adCmdStoredProc

		    .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

		    .Parameters.Append .CreateParameter("@tax_type"     ,adVarWChar,adParamInput,3, "MM")
			.Parameters.Append .CreateParameter("@tax_no"     ,adVarWChar,adParamInput,18, I1_no)
		    .Parameters.Append .CreateParameter("@usr_id"     ,adVarWChar,adParamInput,13, gUsrID)

		    .Execute ,, adExecuteNoRecords

		End With

		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		if  IntRetCD <> 0 then
		    Call DisplayMsgBox("205914", vbInformation, I1_no, "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Call SubCloseCommandObject(lgObjComm)
			Call SubCloseRs(pvObjRs)
			Call SubCloseDB(lgObjConn)
			Response.end
		end if

    
		Call SubCloseCommandObject(lgObjComm)

	End If
	
	
	Call SubCloseRs(pvObjRs)
	Call SubCloseDB(lgObjConn)
	
End Sub

	
%>
