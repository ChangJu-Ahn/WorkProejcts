<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5113mb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 매입지급내역등록 Query Transaction 처리용 ASP								*
'*  7. Modified date(First) : 2001/07/01																*
'*  8. Modified date(Last)  : 2004/03/23																*
'*  9. Modifier (First)     : Ma Jin Ha																	*
'* 10. Modifier (Last)      : Kim Jin Tae																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************
    Dim lgOpModeCRUD
    
    Call HideStatusWnd                                          '☜: Hide Processing message
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	
	On Error Resume Next
    Err.Clear	
    
    lgOpModeCRUD  = Request("txtMode") 
																'☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                    '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                    '☜: Save,Update, Delete
           Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                       
    '         Call SubBizDelete()
		Case "PostDtUpdate"					 
			 Call PostDtUpdate()
		Case "LookupDailyExRt"									'☜:화폐에 따른 화폐율 변경시 호출 
			 Call LookupDailyExRt()
		Case "Release"											'☜: 확정,확정취소 요청을 받음 
			 Call SubRelease()
		
    End Select  
   
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next
    Err.Clear
    
	Dim iCommandSent
	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	Const C_SHEETMAXROWS_D  = 100
	
	Dim iPM8G119
	Dim iPM8G119_E1_ief_supplied_flag_pp_flg
    Dim iPM8G119_E2_m_config_process
		Const M595_E2_po_type_cd = 0    
		Const M595_E2_po_type_nm = 1
    Dim iPM8G119_E3_ief_supplied_flag_gl_type
    Dim iPM8G119_E4_b_biz_partner_build_bp_nm
		Const M595_E4_bp_cd = 0			
		Const M595_E4_bp_nm = 1
    Dim iPM8G119_E5_b_biz_partner_payee_bp_nm
		Const M595_E5_bp_cd = 0			
		Const M595_E5_bp_nm = 1
    Dim iPM8G119_E6_b_biz_area_tax_biz_area_nm
    Dim iPM8G119_E7_b_biz_partner
		Const M595_E7_bp_cd = 0			
		Const M595_E7_bp_rgst_no = 1
		Const M595_E7_bp_nm = 2
    Dim iPM8G119_E8_b_pur_grp
		Const M595_E8_pur_grp = 0		
		Const M595_E8_pur_grp_nm = 1
    Dim iPM8G119_E9_m_iv_hdr
		Const M595_E9_iv_no = 0			
		Const M595_E9_iv_dt = 1
		Const M595_E9_ap_post_dt = 2
		Const M595_E9_pay_dt = 3
		Const M595_E9_posted_flg = 4
		Const M595_E9_sppl_iv_no = 5
		Const M595_E9_payee_cd = 6
		Const M595_E9_build_cd = 7
		Const M595_E9_pur_org = 8
		Const M595_E9_iv_biz_area = 9
		Const M595_E9_tax_biz_area = 10
		Const M595_E9_iv_cost_cd = 11
		Const M595_E9_pay_meth = 12
		Const M595_E9_pay_dur = 13
		Const M595_E9_pay_terms_txt = 14
		Const M595_E9_pay_type = 15
		Const M595_E9_gross_doc_amt = 16
		Const M595_E9_gross_loc_amt = 17
		Const M595_E9_net_doc_amt = 18
		Const M595_E9_net_loc_amt = 19
		Const M595_E9_cash_doc_amt = 20
		Const M595_E9_cash_loc_amt = 21
		Const M595_E9_iv_cur = 22
		Const M595_E9_xch_rt = 23
		Const M595_E9_vat_type = 24
		Const M595_E9_vat_rt = 25
		Const M595_E9_tot_vat_doc_amt = 26
		Const M595_E9_tot_vat_loc_amt = 27
		Const M595_E9_tot_diff_doc_amt = 28
		Const M595_E9_tot_diff_loc_amt = 29
		Const M595_E9_pay_bank_cd = 30
		Const M595_E9_pay_acct_cd = 31
		Const M595_E9_pp_no = 32
		Const M595_E9_pp_doc_amt = 33
		Const M595_E9_pp_loc_amt = 34
		Const M595_E9_remark = 35
		Const M595_E9_loan_no = 36
		Const M595_E9_loan_doc_amt = 37
		Const M595_E9_loan_loc_amt = 38
		Const M595_E9_bl_no = 39
		Const M595_E9_bl_doc_no = 40
		Const M595_E9_lc_doc_no = 41
		Const M595_E9_ref_po_no = 42
		Const M595_E9_ext1_cd = 43
		Const M595_E9_gl_no = 44
		Const M595_E9_ext1_qty = 45
		Const M595_E9_ext1_amt = 46
		Const M595_E9_ext1_rt = 47
		Const M595_E9_ext1_dt = 48
		Const M595_E9_ext2_cd = 49
		Const M595_E9_ext2_qty = 50
		Const M595_E9_ext2_amt = 51
		Const M595_E9_ext2_rt = 52
		Const M595_E9_ext2_dt = 53
		Const M595_E9_ext3_cd = 54
		Const M595_E9_ext3_qty = 55
		Const M595_E9_ext3_amt = 56
		Const M595_E9_ext3_rt = 57
		Const M595_E9_ext3_dt = 58
		Const M595_E9_xch_rate_op = 59
		Const M595_E9_vat_inc_flag = 60
    Dim iPM8G119_E10_m_iv_type
		Const M595_E10_iv_type_nm = 0		
		Const M595_E10_iv_type_cd = 1
		Const M595_E10_import_flg = 2
		Const M595_E10_except_flg = 3
		Const M595_E10_ret_flg = 4
    Dim iPM8G119_E11_b_minor_nm_vat 
    Dim iPM8G119_E12_b_minor_nm_pay_meth
    Dim iPM8G119_E13_b_minor_nm_pay_type
    Dim iPM8G119_E14_b_currency_currency_desc
    Dim iPM8G119_E15_m_pur_ord_hdr
		Const M595_E15_rcpt_flg = 0			
		Const M595_E15_rcpt_type = 1
		Const M595_E15_issue_type = 2
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'PM1G429 DLL의 변수들..
	Dim iPM1G429
	Dim iPM1G429_I1_m_iv_type
	Dim iPM1G429_E1_m_iv_type
		Const M358_E1_iv_type_cd = 0		'  View Name : exp m_iv_type
		Const M358_E1_iv_type_nm = 1
		Const M358_E1_trans_cd = 2
		Const M358_E1_import_flg = 3
		Const M358_E1_except_flg = 4
		Const M358_E1_ret_flg = 5
		Const M358_E1_usage_flg = 6
		Const M358_E1_ext1_cd = 7
		Const M358_E1_ext2_cd = 8
		Const M358_E1_ext3_cd = 9
		Const M358_E1_ext4_cd = 10
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'PM8G131 DLL의 변수들..
	Dim iPM8G131
	Dim iPM8G131_EG1_exp_group
		Const M588_EG1_E1_minor_cd = 0		'  View Name : exp_item_pay_type_cd b_minor
		Const M588_EG1_E2_minor_cd = 1		'  View Name : exp_item b_minor
		Const M588_EG1_E2_minor_nm = 2
		Const M588_EG1_E3_bank_cd = 3		'  View Name : exp_item b_bank
		Const M588_EG1_E3_bank_nm = 4
		Const M588_EG1_E4_iv_pay_seq = 5    '  View Name : exp_item m_iv_payment_dtl
		Const M588_EG1_E4_pay_type = 6
		Const M588_EG1_E4_cur = 7
		Const M588_EG1_E4_xch_rt = 8
		Const M588_EG1_E4_pay_doc_amt = 9
		Const M588_EG1_E4_pay_loc_amt = 10
		Const M588_EG1_E4_note_no = 11
		Const M588_EG1_E4_bank_cd = 12
		Const M588_EG1_E4_prpaym_no = 13
		Const M588_EG1_E4_bank_acct_no = 14
		Const M588_EG1_E4_loan_no = 15
		Const M588_EG1_E4_ext1_cd = 16
		Const M588_EG1_E4_ext1_qty = 17
		Const M588_EG1_E4_ext1_amt = 18
		Const M588_EG1_E4_ext1_rt = 19
		Const M588_EG1_E4_ext1_dt = 20
		Const M588_EG1_E4_ext2_cd = 21
		Const M588_EG1_E4_ext2_qty = 22
		Const M588_EG1_E4_ext2_amt = 23
		Const M588_EG1_E4_ext2_rt = 24
		Const M588_EG1_E4_ext2_dt = 25
		Const M588_EG1_E4_ext3_cd = 26
		Const M588_EG1_E4_ext3_qty = 27
		Const M588_EG1_E4_ext3_amt = 28
		Const M588_EG1_E4_ext3_rt = 29
		Const M588_EG1_E4_ext3_dt = 30
    Dim iPM8G131_E1_m_iv_payment_dtl
		
    Dim Str_L_Iv_Pay_Seq, iIvNo
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'PM8G128 DLL의 변수들..
	Dim iPM8G128
	Dim iPM8G128_EG1_exp_group
	Dim iPM8G128_E1_m_iv_dtl_iv_seq_no
    
    Dim TmpBuffer
	Dim IStr
	Dim ITotalStr
	Dim iMax
	Dim iIntLoopCount
	                                                          
    lgStrPrevKey = Request("lgStrPrevKey")
    iIvNo = Trim(Request("txtIvNo"))

	If Request("txtMaxRows") = 0 Then
	
		Set iPM8G119 = CreateObject("PM8G119.cMLookupIvHdrS")    
	
		If CheckSYSTEMError(Err,True) = true then 
			Set iPM8G119 = Nothing
			Exit Sub
		End If
					
		Call iPM8G119.M_LOOKUP_IV_HDR_SVR(gStrGlobalCollection, iIvNo, iPM8G119_E1_ief_supplied_flag_pp_flg, _
					  iPM8G119_E2_m_config_process, iPM8G119_E3_ief_supplied_flag_gl_type, iPM8G119_E4_b_biz_partner_build_bp_nm, _
					  iPM8G119_E5_b_biz_partner_payee_bp_nm, iPM8G119_E6_b_biz_area_tax_biz_area_nm, iPM8G119_E7_b_biz_partner, _
					  iPM8G119_E8_b_pur_grp, iPM8G119_E9_m_iv_hdr, iPM8G119_E10_m_iv_type, iPM8G119_E11_b_minor_nm_vat, iPM8G119_E12_b_minor_nm_pay_meth, _
					  iPM8G119_E13_b_minor_nm_pay_type, iPM8G119_E14_b_currency_currency_desc, iPM8G119_E15_m_pur_ord_hdr)	
	
		If CheckSYSTEMError2(Err,True, "","","","","") = True Then
			Set iPM8G119 = Nothing			
			Exit Sub
		End If
		Set iPM8G119 = Nothing
			
			
		Set iPM1G429 = CreateObject("PM1G429.cMLookupIvTypeSvr")    
	
		If CheckSYSTEMError(Err,True) = true then 
			Set iPM1G429 = Nothing	
			Exit Sub
		End If
			
		iPM1G429_I1_m_iv_type = ConvSPChars(iPM8G119_E10_m_iv_type(M595_E10_iv_type_cd))
			
		iPM1G429_E1_m_iv_type = iPM1G429.LOOKUP_IV_TYPE_SVR(gStrGlobalCollection, iPM1G429_I1_m_iv_type)

		If CheckSYSTEMError2(Err,True, "","","","","") = True Then
			Set iPM1G429 = Nothing			
			Exit Sub
		End If
		Set iPM1G429 = Nothing
			
			
		Const strDefDate = "1899-12-30"
	
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		'##### Rounding Logic #####		<=========================================여기 
		Response.Write "	.frm1.txtCur.value        = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_iv_cur))		   & """" & vbCr
		Response.Write "	.CurFormatNumericOCX "          &vbCr
		'##########################		<=========================================여기 
		Response.Write "	.frm1.hdnImportflg.value  = """ & ConvSPChars(iPM1G429_E1_m_iv_type(M358_E1_import_flg))   & """" & vbCr
		Response.Write "	.frm1.hdnGlNo.value       = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_gl_no))		   & """" & vbCr
		Response.Write "	.frm1.hdnGlType.value     = """ & ConvSPChars(iPM8G119_E3_ief_supplied_flag_gl_type)	   & """" & vbCr
	
		Response.Write "	.frm1.txthdnIvNo.value    = """ & ConvSPChars(Request("txtIvNo"))					       & """" & vbCr
		Response.Write "	.frm1.txtIvNo.value       = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_iv_no))		   & """" & vbCr
		Response.Write "	.frm1.txtSpplCd.value     = """ & ConvSPChars(iPM8G119_E7_b_biz_partner(M595_E7_bp_cd))    & """" & vbCr
		Response.Write "	.frm1.txtSpplNm.value     = """ & ConvSPChars(iPM8G119_E7_b_biz_partner(M595_E7_bp_nm))    & """" & vbCr
			 
		If iPM8G119_E9_m_iv_hdr(M595_E9_bl_doc_no) = "" Then
			Response.Write "	.frm1.txtBLIvNo.value = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_sppl_iv_no))    & """" & vbCr
		Else
			Response.Write "	.frm1.txtBLIvNo.value = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_bl_doc_no))     & """" & vbCr
		End If
	
		Response.Write "	.frm1.txtIvTypeCd.value   = """ & ConvSPChars(iPM8G119_E10_m_iv_type(M595_E10_iv_type_cd)) & """" & vbCr
		Response.Write "	.frm1.txtIvTypeNm.value   = """ & ConvSPChars(iPM8G119_E10_m_iv_type(M595_E10_iv_type_nm)) & """" & vbCr
		Response.Write "	.frm1.txtGrpCd.value      = """ & ConvSPChars(iPM8G119_E8_b_pur_grp(M595_E8_pur_grp))	   & """" & vbCr
		Response.Write "	.frm1.txtGrpNm.value      = """ & ConvSPChars(iPM8G119_E8_b_pur_grp(M595_E8_pur_grp_nm))   & """" & vbCr
		Response.Write "	.frm1.txtXchRt.value      = """ & UNINumClientFormat(iPM8G119_E9_m_iv_hdr(M595_E9_xch_rt), ggExchRate.DecPoint, 1)  & """" & vbCr
		Response.Write "	.frm1.hdnDiv.value		  = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_xch_rate_op))		& """" & vbCr
		Response.Write "	.frm1.txtIvDt.value       = """ & UNIDateClientFormat(iPM8G119_E9_m_iv_hdr(M595_E9_iv_dt))       & """" & vbCr
		Response.Write "	.frm1.txtDocAmt.value     = """ & UNIConvNumDBToCompanyByCurrency(iPM8G119_E9_m_iv_hdr(M595_E9_gross_doc_amt), ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_iv_cur)), ggAmtOfMoneyNo,"X","X") & """" & vbCr
		Response.Write "	.frm1.txtLocAmt.value     = """ & UniConvNumberDBToCompany(iPM8G119_E9_m_iv_hdr(M595_E9_gross_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)           & """" & vbCr
		Response.Write "	.frm1.hdnLoanAmt.value    = """ & UNIConvNumDBToCompanyByCurrency(iPM8G119_E9_m_iv_hdr(M595_E9_loan_loc_amt), ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_iv_cur)), ggAmtOfMoneyNo,"X","X")  & """" & vbCr
		Response.Write "	.frm1.txtPostDt.text      = """ & UNIDateClientFormat(iPM8G119_E9_m_iv_hdr(M595_E9_ap_post_dt)) & """" & vbCr
		Response.Write "	.frm1.hdnPostDt.value     = .frm1.txtPostDt.text " & vbCr
		Response.Write "	.frm1.txtPost.value       = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_posted_flg))			& """" & vbCr
		Response.Write "	.frm1.txtGlNo.value       = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_gl_no))				& """" & vbCr
		Response.Write "	.frm1.txtBlNo.value       = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_bl_no))				& """" & vbCr
		Response.Write "	.frm1.txtPayeeCd.value    = """ & ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_payee_cd))           & """" & vbCr
	
		If ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_posted_flg)) = "Y" Then 
			Response.Write "   .frm1.rdoApFlg(0).Checked= true "   & vbCr 
			Response.Write "   .frm1.hdnPostingFlg.value = ""Y"" " & vbCr 
		Else
			Response.Write "   .frm1.rdoApFlg(1).Checked= true "   & vbCr 
			Response.Write "   .frm1.hdnPostingFlg.value = ""N"" " & vbCr 
		End If 
			
		'2003.03 KJH 전표번호 가져오는 로직 수정 
		Response.Write "parent.SubGetGlNo" & vbCr
	
		Response.Write "  .CurFormatNumSprSheet " & vbCr
		Response.Write "End With "   & vbCr
		Response.Write "</Script> " & vbCr
	End If
		
		
	Set iPM8G131 = CreateObject("PM8G131.cMListIvPaymentDtlS")
		
	If CheckSYSTEMError(Err,True) = true then 
		Set iPM8G131 = Nothing
		Exit Sub
	End If
		
	 if Trim(lgStrPrevKey) <> "" then
		Str_L_Iv_Pay_Seq = lgStrPrevKey
	End if
		
	iIvNo = Trim(Request("txtIvNo"))
		
	Call iPM8G131.M_LIST_IV_PAYMENT_DTL_SVR(gStrGlobalCollection, Cint(C_SHEETMAXROWS_D), Str_L_Iv_Pay_Seq, iIvNo, _
				  iPM8G131_EG1_exp_group, iPM8G131_E1_m_iv_payment_dtl)
		
	If Request("txtQuerytype") <> "Query" And Err.Description = "B_MESSAGE175500" then
		Set iPM8G131 = Nothing
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & chr(13)
			Response.Write "parent.dbQueryOk" & chr(13)
			Response.Write "</Script>"
		Exit Sub												'☜: ComProxy Unload
	Else
				  
		If CheckSYSTEMError2(Err,True, "","","","","") = True Then
			Set iPM8G131 = Nothing			
			Response.Write "<Script Language=vbscript>" & vbCr
			'Detail항목이 없을 경우 Header정보만 보여줌 
			Response.Write "parent.dbQueryOk" & vbCr
			Response.Write "</Script>"
			Exit Sub
		End If
	End If
			
	Set iPM8G131 = Nothing
		
	'2005.04 KJH
	IMax = 	UBound(iPM8G131_EG1_exp_group, 1)
	Redim TmpBuffer(IMax)
	iIntLoopCount = 0
	
	For iLngRow = 0 To IMax
	
		If  iIntLoopCount < (C_SHEETMAXROWS_D ) Then
		Else
		   StrNextKey = ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_iv_pay_seq)) 
		   Exit For
		End If  
		istrData = ""		
        istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_pay_type))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E2_minor_nm))
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_pay_doc_amt), ConvSPChars(iPM8G119_E9_m_iv_hdr(M595_E9_iv_cur)), ggAmtOfMoneyNo,"X","X")
        istrData = istrData & Chr(11) & UniConvNumberDBToCompany(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_pay_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_xch_rt),ggExchRate.DecPoint,0)
		istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_bank_acct_no))
		istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E3_bank_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E3_bank_nm))
		istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_note_no))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_prpaym_no))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_loan_no))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(iPM8G131_EG1_exp_group(iLngRow, M588_EG1_E4_iv_pay_seq))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)  
		
		TmpBuffer(iIntLoopCount) = istrData
		iIntLoopCount = iIntLoopCount + 1
    Next
    
    ITotalStr = Join(TmpBuffer, "")
    
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
    Response.Write "	.ggoSpread.Source        = .frm1.vspdData " & vbCr
    Response.Write "	.ggoSpread.SSShowData      """ & ITotalStr	 & """" & vbCr	
    Response.Write "	.lgStrPrevKey            = """ & StrNextKey & """" & vbCr  
    Response.Write "	.frm1.txthdnIvNo.value   = """ & ConvSPChars(Request("txtIvNo")) & """" & vbCr
	'Response.Write "    .ChkExistIvDtlByIvNo(""" & iIvNo & """) " & vbCr 
	Response.Write "    .DbQueryOk() " & vbCr 
	Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr
    		
    End Sub
    
    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
	Sub SubBizSaveMulti()
		Dim iPM8G132
		Dim iErrorPosition
		Dim B_Company_Cur
		Dim I1_ief_supplied
		Dim iIvNo, itxtSpread
			
		On Error Resume Next
		Err.Clear
		
		Set iPM8G132 = CreateObject("PM8G132.cMMaintIvPayDtlS")    
	   
	    If CheckSYSTEMError(Err,True) = true Then 		
			Set iPM8G132 = Nothing
			Exit Sub														'☜: 비지니스 로직 처리를 종료함 
		End if
		
	
        I1_ief_supplied = Trim(Request("hdninterface_Account"))
        iIvNo = Trim(Request("txthdnIvNo"))
        itxtSpread = Trim(Request("txtSpread"))
        
        Call iPM8G132.M_MAINT_IV_PAYMENT_DTL_SVR(gStrGlobalCollection, I1_ief_supplied, Cstr(gCurrency), iIvNo, , itxtSpread, iErrorPosition)                   
        
        
        If CheckSYSTEMError2(Err,True, iErrorPosition & "행:" ,"","","","") = True then
		  	Set iPM8G132 = Nothing
		  	Exit Sub
		End If

		Set iPM8G132 = Nothing   
					

		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " parent.frm1.txtIvNo.Value = """ & ConvSPChars(iIvNo) & """" & vbCr    
		Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "</Script> "           
        
	End Sub



'============================================================================================================
' Name : PostDtUpdate
' Desc : 회계적용일update pad
'============================================================================================================
Sub PostDtUpdate()

	On Error Resume Next
    Err.Clear
    
    Dim iPM8G111
    Dim iCommandSent
    Dim pvCB
    Dim I5_m_iv_hdr
		Const M579_I5_ap_post_dt = 12
		Const M579_I5_iv_no = 13
		
    
    Set iPM8G111 = Server.CreateObject("PM8G111.cMMaintIvHdrS")
    
    If CheckSYSTEMError(Err,True) = true Then 		
		Exit Sub														    '☜: 비지니스 로직 처리를 종료함 
	End if	
	
	
	iCommandSent = "UPDATEGP"
	
	Redim I5_m_iv_hdr(59)
	I5_m_iv_hdr(M579_I5_iv_no) = Trim(Request("IvNo"))
	If Trim(Request("PostDt")) <> "" Then 
		I5_m_iv_hdr(M579_I5_ap_post_dt) = UniConvDate(Request("PostDt")) '확정일 
	End if
	pvCB = "F"
	Call iPM8G111.M_MAINT_IV_HDR_SVR(pvCB, gStrGlobalCollection,iCommandSent, , , , , I5_m_iv_hdr)
	
	
	If CheckSYSTEMError2(Err,True, ,"","","","") = true then 		
		Set iPM8G111 = Nothing												'☜: ComPlus Unload
		Exit Sub														    '☜: 비지니스 로직 처리를 종료함 
	End if
    Set iPM8G111 = Nothing
    
    Response.Write "<Script Language=vbscript>"
    Response.Write "If parent.frm1.hdnPostDt.value <> parent.frm1.txtPostDt.text Then" & vbCr
    Response.Write " parent.DbSaveOk " & vbCr
	Response.Write "End If " & vbCr
	Response.Write "</Script> " & vbCr
	
		
End Sub

'============================================================================================================
' Name : LookupDailyExRt
' Desc : 화폐에 따른 화폐율 변경시 호출 
'============================================================================================================
Sub LookupDailyExRt()

	On Error Resume Next
    Err.Clear
    
	Dim iPB0C004
	Dim iCurrency, iChargeDt
	Dim E_B_Daily_Exchange_Rate
		Const B253_E1_std_rate = 0
		Const B253_E1_multi_divide = 1
	
    Set iPB0C004 = CreateObject("PB0C004.CB0C004")

    If CheckSYSTEMError(Err,True) = true Then 		
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
    
    iCurrency = Request("Currency")
    iChargeDt = UNIConvDate(Request("ChargeDt"))
    E_B_Daily_Exchange_Rate = iPB0C004.B_SELECT_EXCHANGE_RATE(gStrGlobalCollection,iCurrency, gCurrency, iChargeDt)

	If CheckSYSTEMError2(Err,True, ,"","","","") = true then 		
		Set iPB0C004 = Nothing												'☜: ComPlus Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	Set iPB0C004 = Nothing
	
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write " <Script Language=vbscript>"    & vbCr
	Response.Write "With parent.frm1.vspdData "     & vbCr ' 환율을 구해 지급자국금액을 구한다 
	Response.Write "	Dim DocAmt, ExchRate "      & vbCr
	Response.Write "	.Row = " & Request("LRow")  & vbCr '.ActiveRow
	Response.Write "	.Col = parent.C_PayType "   & vbCr
	Response.Write "	If parent.CheckPayType(Trim(.Text)) <> ""PP"" Then " & vbCr '선수금경우가 아닌 경우 
	Response.Write "	.Col = parent.C_ExchRate " & vbCr
	Response.Write "	.Text = """ & UNINumClientFormat(E_B_Daily_Exchange_Rate(B253_E1_std_rate), ggExchRate.DecPoint,0) & """" & vbCr
	Response.Write "	ExchRate = .Text " & vbCr
	Response.Write "Else " & vbCr							'선수금경우인 경우 
	Response.Write "	.Col = parent.C_ExchRate " & vbCr
	Response.Write "	ExchRate = .Text " & vbCr
	Response.Write "End If " & vbCr
	
	Response.Write "	.Col = parent.C_PayDocAmt " & vbCr
	Response.Write "	DocAmt = .Text " & vbCr
	
	Response.Write "	.Col = parent.C_PayLocAmt " & vbCr
	Response.Write "If """ & ConvSPChars(E_B_Daily_Exchange_Rate(B253_E1_multi_divide)) & """ = ""*"" Then " & vbCr
	Response.Write "	.Text = parent.UNIConvNumPCToCompanyByCurrency(parent.UniCdbl(DocAmt)  * parent.UniCdbl(ExchRate),gCurrency, parent.parent.ggAmtOfMoneyNo, parent.parent.gLocRndPolicyNo , ""X"" ) " & vbCr
	Response.Write "Else " & vbCr
	Response.Write "	.Text = parent.UNIConvNumPCToCompanyByCurrency(parent.UniCdbl(DocAmt) / parent.UniCdbl(ExchRate),gCurrency, parent.parent.ggAmtOfMoneyNo, parent.parent.gLocRndPolicyNo , ""X"" ) " & vbCr
	Response.Write "End If " & vbCr
	Response.Write "Parent.TotalSum " & vbCr
	Response.Write "End With "        & vbCr
	Response.Write "</Script> "       & vbCr

    
End Sub

'============================================================================================================
' Name : SubRelease
' Desc : 확정,확정취소 요청을 받음 
'============================================================================================================
Sub SubRelease()
	
	Dim iPM8G211
	Dim L_SelectChar
	Dim I3_m_batch_ap_post_wks
	Dim pvCB
	
	Dim IG1_imp_dtl_group
	
	Const M557_I3_ap_dt_type = 0
    Const M557_I3_ap_dt = 1
    Const M557_I3_import_flg = 2
    
    Const M557_IG1_I1_count = 0
    Const M557_IG1_I2_iv_no = 1
    Const M557_IG1_I3_ap_dt = 2
    
	On Error Resume Next
    Err.Clear
    
    Set iPM8G211 = CreateObject("PM8G211.cMPostApS")    
    
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM8G211 = Nothing
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	'--------------------
    '수정(2003.06.09)
    ReDim IG1_imp_dtl_group(0, 2)
    
	IG1_imp_dtl_group(0, M557_IG1_I2_iv_no)  = Trim(Request("txtIvNo"))
    IG1_imp_dtl_group(0, M557_IG1_I1_count) = 1
    IG1_imp_dtl_group(0, M557_IG1_I3_ap_dt) = UNIConvDate(Trim(Request("txtPostDt")))

	ReDim I3_m_batch_ap_post_wks(2)
	I3_m_batch_ap_post_wks(M557_I3_ap_dt_type) = ""
    'I3_m_batch_ap_post_wks(M557_I3_ap_dt) = UNIConvDate(Trim(Request("txtPostDt")))
    I3_m_batch_ap_post_wks(M557_I3_import_flg) = Trim(Request("hdnImportFlg"))
    
	
	IF Request("hdnPostingFlg") = "Y" then
		L_SelectChar		= "N"
	Else
		L_SelectChar		= "Y"
	End if
	pvCB = "F"
	Call iPM8G211.M_POST_AP_SVR(pvCB, gStrGlobalCollection, L_SelectChar, IG1_imp_dtl_group, I3_m_batch_ap_post_wks)

	If CheckSYSTEMError2(Err,True, "","","","","") = True Then
	  	Set iPM8G211 = Nothing
	  	Exit Sub
	End If
	Set iPM8G211 = Nothing
		
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.MainQuery() "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> " & vbCr

    Set iPM8G211 = Nothing                   
	
		
End Sub

%>
