<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Procuremen
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : M51119(Lookup_PO_Hdr)
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2001/10
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Ma Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'* 14. Business Logic of m5111ma1(매입일반등록)
'**********************************************************************************************
	'Dim lgOpModeCRUD
	Dim pvCB
	
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	lgOpModeCRUD	=	Request("txtMode")           '☜: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)                                                         '☜: Query
	             Call SubBizQuery()
	        Case CStr(UID_M0002)
	             Call SubBizSave()
	        Case CStr(UID_M0003)                                                         '☜: Delete
	             Call SubBizDelete()
	        Case "Release", "UnRelease"
				 Call SubReleaseCheck()
	        Case "LookUpSupplier"                                                                 '☜: Check	
	             Call SubLookUpSupplier()    
			Case "LookUpPo"	             
	             Call SubLookUpPo()
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
    Err.Clear 

	Dim iPM8G119
	
	Dim I1_m_iv_hdr_iv_no
    Dim E1_ief_supplied_flag_pp_flg
    Dim E2_m_config_process
    Dim E3_ief_supplied_flag_gl_type
    Dim E4_b_biz_partner_build_bp_nm
    Dim E5_b_biz_partner_payee_bp_nm
    Dim E6_b_biz_area_tax_biz_area_nm
    Dim E7_b_biz_partner
    Dim E8_b_pur_grp
    Dim E9_m_iv_hdr
    Dim E10_m_iv_type
    Dim E11_b_minor_nm_vat
    Dim E12_b_minor_nm_pay_meth
    Dim E13_b_minor_nm_pay_type
    Dim E14_b_currency_currency_desc
    Dim E15_m_pur_ord_hdr

    Const M594_E2_po_type_cd = 0    '  View Name : exp m_config_process
    Const M594_E2_po_type_nm = 1

    Const M594_E4_bp_nm = 0
    Const M594_E4_bp_rgst_no = 1

    Const M594_E5_bp_cd = 0    '  View Name : exp_payee b_biz_partner
    Const M594_E5_bp_nm = 1

    Const M594_E7_bp_cd = 0    '  View Name : exp b_biz_partner
    Const M594_E7_bp_rgst_no = 1
    Const M594_E7_bp_nm = 2

    Const M594_E8_pur_grp = 0    '  View Name : exp b_pur_grp
    Const M594_E8_pur_grp_nm = 1


    Const M594_E9_iv_no = 0    '  View Name : exp m_iv_hdr
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

    Const M594_E9_issue_dt_fg = 86

    Const M594_E10_iv_type_nm = 0    '  View Name : exp m_iv_type
    Const M594_E10_iv_type_cd = 1
    Const M594_E10_import_flg = 2
    Const M594_E10_except_flg = 3
    Const M594_E10_ret_flg = 4

    Const M594_E15_rcpt_flg = 0    '  View Name : exp m_pur_ord_hdr
    Const M594_E15_rcpt_type = 1
    Const M594_E15_issue_type = 2
    
    I1_m_iv_hdr_iv_no	 = Request("txtIvNo")
    
    Set iPM8G119 = Server.CreateObject("PM8G119.cMLookupIvHdrS")    
    
	If CheckSYSTEMError(Err,True) = true then 		
		Set iPM8G119 = Nothing
        Exit Sub
	End if
    
    Call iPM8G119.M_LOOKUP_IV_HDR_SVR(gStrGlobalCollection,cstr(I1_m_iv_hdr_iv_no),E1_ief_supplied_flag_pp_flg,E2_m_config_process, _
			E3_ief_supplied_flag_gl_type,E4_b_biz_partner_build_bp_nm,E5_b_biz_partner_payee_bp_nm,E6_b_biz_area_tax_biz_area_nm, _
			E7_b_biz_partner,E8_b_pur_grp,E9_m_iv_hdr,E10_m_iv_type,E11_b_minor_nm_vat,E12_b_minor_nm_pay_meth, _
			E13_b_minor_nm_pay_type,E14_b_currency_currency_desc,E15_m_pur_ord_hdr)
    
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iPM8G119 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End if
	
	Set iPM8G119 = Nothing												'☜  
	
	'-----------------------
	'Result data display area
	'----------------------- 	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	'##### Rounding Logic #####
	'항상 거래화폐가 우선 
	Response.Write ".txtCur.value 			= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)) & """" & vbCr
	Response.Write " parent.CurFormatNumericOCX" & vbCr

	'##########################
	'반품여부 Flag추가(2003.08.01)
	Response.Write ".hdnRetflg.Value 		= """ & ConvSPChars(E10_m_iv_type(M594_E10_ret_flg)) & """" & vbCr	
	Response.Write ".hdnImportflg.Value 	= """ & ConvSPChars(E10_m_iv_type(M594_E10_import_flg)) & """" & vbCr
	Response.Write ".txtIvNo.value 			= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_iv_no)) & """" & vbCr           '조회 매입번호 
	Response.Write ".txtIvNo1.value 		= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_iv_no)) & """" & vbCr           '데이타 매입번호 
	Response.Write ".hdnIvNo.Value 			= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_iv_no)) & """" & vbCr         'hidden 매입번호 
	Response.Write ".txtIvTypeCd.value		= """ & ConvSPChars(E10_m_iv_type(M594_E10_iv_type_cd)) & """" & vbCr      '매입형태 
	Response.Write ".txtIvTypeNm.value		= """ & ConvSPChars(E10_m_iv_type(M594_E10_iv_type_nm)) & """" & vbCr     '매입형태명 
	Response.Write ".txtIvDt.text 			= """ & UNIDateClientFormat(E9_m_iv_hdr(M594_E9_iv_dt)) & """" & vbCr   '매입일 
	'지불예정일을 입력하지 않을 경우 "2999/12/31"로 셋팅함(2003.09.22)
	Response.Write "if """ & CDate(E9_m_iv_hdr(M594_E9_pay_dt)) & """= ""2999-12-31"" then " & vbCr
	Response.Write ".txtPayDt.text 			= """" " & vbCr '지불예정일 
	Response.Write "else" & vbCr
	Response.Write ".txtPayDt.text 			= """ & UNIDateClientFormat(E9_m_iv_hdr(M594_E9_pay_dt)) & """" & vbCr '지불예정일 
	Response.Write "End if	" & vbCr
	
	Response.Write ".txtPostDt.Text 		= """ & UNIDateClientFormat(E9_m_iv_hdr(M594_E9_ap_post_dt)) & """" & vbCr '확정일 
	Response.Write "if """ & ConvSPChars(E9_m_iv_hdr(M594_E9_posted_flg)) & """= ""Y"" then " & vbCr
	Response.Write ".rdoApFlg(0).Checked= true " & vbCr
	Response.Write ".hdnApFlg.value 	= ""Y"" " & vbCr
	Response.Write "else" & vbCr
	Response.Write ".rdoApFlg(1).Checked= true " & vbCr
	Response.Write ".hdnApFlg.value 	= ""N"" " & vbCr
	Response.Write "End if	" & vbCr	
		
	Response.Write "If """ & ConvSPChars(E9_m_iv_hdr(M594_E9_vat_inc_flag)) & """ = ""2"" Then" & vbCr
	Response.Write "  .rdoVatFlg2.Checked= true " & vbCr
	Response.Write "  .hdvatFlg.value 	=  ""2"" " & vbCr
	Response.Write "Else" & vbCr
	Response.Write "  .rdoVatFlg1.Checked= true " & vbCr
	Response.Write "  .hdvatFlg.value 	= ""1"" " & vbCr
	Response.Write "End If" & vbCr					
	
	Response.Write "If """ & ConvSPChars(Trim(E1_ief_supplied_flag_pp_flg)) & """ =  ""Y"" then" & vbCr
	Response.Write "  .ChkPrepay.checked = true" & vbCr
	Response.Write "else" & vbCr
	Response.Write "  .ChkPrepay.checked = false" & vbCr
	Response.Write "End If" & vbCr
				
	Response.Write ".txtGlNo.Value			= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_gl_no)) & """" & vbCr
	Response.Write ".hdnGlType.value 		= """ & ConvSPChars(E3_ief_supplied_flag_gl_type) & """" & vbCr
	Response.Write ".txtSpplCd.value 		= """ & ConvSPChars(E7_b_biz_partner(M594_E7_bp_cd)) & """" & vbCr
	Response.Write ".hdnSpplCd.value 		= """ & ConvSPChars(E7_b_biz_partner(M594_E7_bp_cd)) & """" & vbCr
		
	Response.Write ".txtSpplNm.value 		= """ & ConvSPChars(E7_b_biz_partner(M594_E7_bp_nm)) & """" & vbCr
	Response.Write ".txtSpplRegNo.value 	= """ & ConvSPChars(E7_b_biz_partner(M594_E7_bp_rgst_no)) & """" & vbCr
	Response.Write ".txtSpplIvNo.value 		= """ & ConvSPChars(Trim(E9_m_iv_hdr(M594_E9_sppl_iv_no))) & """" & vbCr
		
	Response.Write ".txtPayeeCd.value 		= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_payee_cd)) & """" & vbCr
	Response.Write ".txtPayeeNm.value 		= """ & ConvSPChars(E5_b_biz_partner_payee_bp_nm) & """" & vbCr
	Response.Write ".txtBuildCd.value 		= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_build_cd)) & """" & vbCr
	Response.Write ".txtBuildNm.value 		= """ & ConvSPChars(E4_b_biz_partner_build_bp_nm(M594_E4_bp_nm)) & """" & vbCr
		
		
	Response.Write ".txtGrpCd.value 		= """ & ConvSPChars(E8_b_pur_grp(M594_E8_pur_grp)) & """" & vbCr			
	Response.Write ".txtGrpNm.value 		= """ & ConvSPChars(E8_b_pur_grp(M594_E8_pur_grp_nm)) & """" & vbCr
'	.txtIvAmt.text 			= "<%=UNINumClientFormat(E9_m_iv_hdr(M594_E9_gross_doc_amt), ggAmtOfMoney.DecPoint, 0)"""" & vbCr
		
	Response.Write ".txtIvAmt.text 			= """ & UNIConvNumDBToCompanyByCurrency(E9_m_iv_hdr(M594_E9_gross_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)), ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write ".txtIvLocAmt.text 		= """ & UniConvNumberDBToCompany(E9_m_iv_hdr( M594_E9_gross_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & """" & vbCr
		
	Response.Write ".txtnetDocAmt.text      = """ & UNIConvNumDBToCompanyByCurrency(E9_m_iv_hdr(M594_E9_net_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)), ggAmtOfMoneyNo,"X","X") & """" & vbCr '13차추가 
	Response.Write ".txtnetLocAmt.text      = """ & UniConvNumberDBToCompany(E9_m_iv_hdr(M594_E9_net_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & """" & vbCr'13차추가 
	
	Response.Write ".hdnCur.value 			= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)) & """" & vbCr
	Response.Write ".txtCurNm.value 		= """ & ConvSPChars(E14_b_currency_currency_desc) & """" & vbCr
	Response.Write ".txtVatCd.value 		= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_vat_type)) & """" & vbCr
	Response.Write ".txtVatNm.value 		= """ & ConvSPChars(E11_b_minor_nm_vat) & """" & vbCr
	Response.Write ".txtVatRt.text  		= """ & UNINumClientFormat(E9_m_iv_hdr(M594_E9_vat_rt), ggExchRate.DecPoint, 0) & """" & vbCr			
	Response.Write ".txtVatAmt.text  		= """ & UNIConvNumDBToCompanyByCurrency(E9_m_iv_hdr(M594_E9_tot_vat_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)),ggAmtOfMoneyNo,gTaxRndPolicyNo,"X") & """" & vbCr
	Response.Write ".txtVatLocAmt.text      = """ & UNIConvNumDBToCompanyByCurrency(E9_m_iv_hdr(M594_E9_tot_vat_loc_amt), gCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo,"X") & """" &vbCr'13차추가 
		
	Response.Write ".txtPayMethCd.value 	= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_pay_meth)) & """" & vbCr
	Response.Write ".hdnPayMethCd.value 	= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_pay_meth)) & """" & vbCr
	Response.Write ".txtPayMethNm.value 	= """ & ConvSPChars(E12_b_minor_nm_pay_meth) & """" & vbCr
	Response.Write ".txtPayDur.text 		= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_pay_dur)) & """" & vbCr		
	Response.Write ".txtPayTypeCd.value 	= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_pay_type)) & """" & vbCr
	Response.Write ".txtPayTypeNm.value 	= """ & ConvSPChars(E13_b_minor_nm_pay_type) & """" & vbCr
	Response.Write ".txtXchRt.text  		= """ & UNINumClientFormat(E9_m_iv_hdr(M594_E9_xch_rt), ggExchRate.DecPoint, 0) & """" & vbCr
	'환율연산자 조회(2003.09.23)
	Response.Write ".hdnDiv.value           = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_xch_rate_op)) & """" & vbcr
	Response.Write ".txtPayTermsTxt.Value	= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_pay_terms_txt)) & """" & vbCr
	'Response.Write ".txtLoanNo.Value  		= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_loan_no)) & """" & vbCr
    'Response.Write ".txtLoanAmt.Text  		= """ & UNIConvNumDBToCompanyByCurrency(E9_m_iv_hdr(M594_E9_loan_doc_amt), ConvSPChars(E9_m_iv_hdr(M594_E9_iv_cur)), ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write ".txtBlDocNo.Value		= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_bl_doc_no)) & """" & vbCr
	Response.Write ".txtBlNo.Value			= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_bl_no)) & """" & vbCr

	Response.Write ".txtPoNo.value = """ & ConvSPChars(E9_m_iv_hdr(M594_E9_ref_po_no)) & """" & vbCr
	
	Response.Write "If """ & ConvSPChars(Trim(E9_m_iv_hdr(M594_E9_ref_po_no))) & """ <> """" Then " & vbCr
	Response.Write "  .chkPoNo.checked = True " & vbCr
	Response.Write "  .txtChkPoNo.value = ""Y"" " & vbCr
	Response.Write "End If " & vbCr
		
	Response.Write ".txtBizAreaCd.Value 	= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_tax_biz_area)) & """" & vbCr
	Response.Write ".txtBizAreaNm.Value 	= """ & ConvSPChars(E6_b_biz_area_tax_biz_area_nm) & """" & vbCr
	Response.Write ".txtMemo.value 			= """ & ConvSPChars(E9_m_iv_hdr(M594_E9_remark)) & """" & vbCr
	
    Response.Write "If """ & ConvSPChars(E9_m_iv_hdr(M594_E9_issue_dt_fg)) & """ = ""Y""  then " & vbCr
    Response.Write "  .rdoIssueDTFg1.Checked = true " & vbCr
    Response.Write "  .hdIssueDTFg.value  = ""Y""" & vbCr
    Response.Write "Else" & vbCr
    Response.Write "  .rdoIssueDTFg2.Checked = true " & vbCr
    Response.Write "  .hdIssueDTFg.value  = ""N""" & vbCr
    Response.Write "End If  " & vbCr

	Response.Write "End With" & vbCr
	'2003.03 KJH 전표번호 가져오는 로직 수정 
	Response.Write "parent.SubGetGlNo" & vbCr
	
	Response.Write "parent.DbQueryOk" & vbCr																'☜: 조화가 성공 
	
	Response.Write "</Script>" & vbCr


    Set M51119 = Nothing															'☜: Unload Comproxy
	
End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSave()		'☜: 저장 요청을 받음 
	
	Dim lgIntFlgMode
	Dim iCommandSent
	
	Dim pvIvNo
	
	On Error Resume Next
	Err.Clear																		'☜: Protect system from crashing
	
	'For 전자세금계산서 
	If lgIntFlgMode = OPMD_CMODE Then
		pvIvNo			= Trim(UCase(Request("txtIvNo1")))
	Else
		pvIvNo			= Trim(Request("hdnIvNo"))
	End If 
	
	Call CheckNoForDT(pvIvNo)
	
		
	If Len(Trim(Request("txtIvDt"))) Then
		If UNIConvDate(Request("txtIvDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtIvDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
	'지급예정일 필수입력 체크 하지 않음(2003.09.22)
	'If Len(Trim(Request("txtPayDt"))) Then
	'	If UNIConvDate(Request("txtPayDt")) = "" Then
	'	    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
	'	    Call LoadTab("parent.frm1.txtPayDt", 0, I_MKSCRIPT)
	'	    Exit Sub	
	'	End If
	'End If
	
	If Len(Trim(Request("txtPostDt"))) Then
		If UNIConvDate(Request("txtPostDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtPostDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If


	Dim iPM8G111		
    Dim I1_m_iv_type_iv_type_cd
    Dim I2_b_biz_partner_bp_cd
    Dim I3_b_pur_grp
    Dim I4_b_company
    Dim I5_m_iv_hdr
    Dim I6_m_user_id
    Dim E1_Return
    
    Redim I5_m_iv_hdr(60)
    
    Const M528_I4_iv_dt = 0    '  View Name : imp m_iv_hdr
    Const M528_I4_pay_dt = 1
    Const M528_I4_iv_cur = 2
    Const M528_I4_xch_rt = 3
    Const M528_I4_vat_type = 4
    Const M528_I4_pay_meth = 5
    Const M528_I4_pay_dur = 6
    Const M528_I4_vat_rt = 7
    Const M528_I4_remark = 8
    Const M528_I4_gross_doc_amt = 9
    Const M528_I4_tot_vat_doc_amt = 10
    Const M528_I4_sppl_iv_no = 11
    Const M528_I4_ap_post_dt = 12
    Const M528_I4_iv_no = 13
    Const M528_I4_posted_flg = 14
    Const M528_I4_payee_cd = 15
    Const M528_I4_build_cd = 16
    Const M528_I4_pur_org = 17
    Const M528_I4_iv_biz_area = 18
    Const M528_I4_tax_biz_area = 19
    Const M528_I4_iv_cost_cd = 20
    Const M528_I4_pay_terms_txt = 21
    Const M528_I4_pay_type = 22
    Const M528_I4_gross_loc_amt = 23
    Const M528_I4_net_doc_amt = 24
    Const M528_I4_net_loc_amt = 25
    Const M528_I4_cash_doc_amt = 26
    Const M528_I4_cash_loc_amt = 27
    Const M528_I4_tot_vat_loc_amt = 28
    Const M528_I4_tot_diff_doc_amt = 29
    Const M528_I4_tot_diff_loc_amt = 30
    Const M528_I4_pay_bank_cd = 31
    Const M528_I4_pay_acct_cd = 32
    Const M528_I4_pp_no = 33
    Const M528_I4_pp_doc_amt = 34
    Const M528_I4_pp_loc_amt = 35
    Const M528_I4_loan_no = 36
    Const M528_I4_loan_doc_amt = 37
    Const M528_I4_loan_loc_amt = 38
    Const M528_I4_bl_no = 39
    Const M528_I4_bl_doc_no = 40
    Const M528_I4_lc_doc_no = 41
    Const M528_I4_ref_po_no = 42
    Const M528_I4_ext1_cd = 43
    Const M528_I4_ext1_qty = 44
    Const M528_I4_ext1_amt = 45
    Const M528_I4_ext1_rt = 46
    Const M528_I4_ext1_dt = 47
    Const M528_I4_ext2_cd = 48
    Const M528_I4_ext2_qty = 49
    Const M528_I4_ext2_amt = 50
    Const M528_I4_ext2_rt = 51
    Const M528_I4_ext2_dt = 52
    Const M528_I4_ext3_cd = 53
    Const M528_I4_ext3_qty = 54
    Const M528_I4_ext3_amt = 55
    Const M528_I4_ext3_rt = 56
    Const M528_I4_ext3_dt = 57
    Const M528_I4_vat_inc_flag = 58
    Const M528_I4_xch_rate_op = 59        
    Const M528_I4_issue_dt_fg = 60        
    
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	Err.Clear	
	
        
    I1_m_iv_type_iv_type_cd    				= Trim(UCase(Request("txtIvTypeCd")))
    I5_m_iv_hdr(M528_I4_iv_dt)				= UniConvDate(Request("txtIvDt"))
    
    '지불예정일을 입력하지 않은 경우 2999/12/31로 셋팅함(2003.09.22)
    If Trim(Request("txtPayDt")) = "" Then
		I5_m_iv_hdr(M528_I4_pay_dt)			= "2999-12-31"
    Else
		I5_m_iv_hdr(M528_I4_pay_dt)			= UniConvDate(Request("txtPayDt"))
    End If
	I5_m_iv_hdr(M528_I4_ap_post_dt)			= UniConvDate(Request("txtPostDt"))
	I5_m_iv_hdr(M528_I4_posted_flg)			= Trim(UCase(Request("hdnApFlg")))
    I5_m_iv_hdr(M528_I4_vat_inc_flag)		= Trim(UCase(Request("hdvatFlg")))    'vat 포함 구분 
 
    I2_b_biz_partner_bp_cd					= Trim(UCase(Request("txtSpplCd")))
    I5_m_iv_hdr(M528_I4_payee_cd)			= Trim(UCase(Request("txtPayeeCd")))
    I5_m_iv_hdr(M528_I4_build_cd)			= Trim(UCase(Request("txtBuildCd")))
    
    I5_m_iv_hdr(M528_I4_sppl_iv_no)			= Trim(Request("txtSpplIvNo"))
    I3_b_pur_grp							= Trim(UCase(Request("txtGrpCd")))
    I5_m_iv_hdr(M528_I4_iv_cur)				= Trim(UCase(Request("txtCur")))
    I5_m_iv_hdr(M528_I4_vat_type)			= Trim(UCase(Request("txtVatCd")))

    if Trim(Request("txtVatRt")) <> "" then
		I5_m_iv_hdr(M528_I4_vat_rt)			= UniConvNum(Request("txtVatRt"),0)
	else
		I5_m_iv_hdr(M528_I4_vat_rt)			= "0"
	End if 

    I5_m_iv_hdr(M528_I4_pay_meth)				= Trim(UCase(Request("txtPayMethCd")))
    if Trim(Request("txtPayDur")) <> "" then
		I5_m_iv_hdr(M528_I4_pay_dur)			= UniConvNum(Request("txtPayDur"),0)
	else
		I5_m_iv_hdr(M528_I4_pay_dur)			= "0"
	end if
    I5_m_iv_hdr(M528_I4_pay_type)				= Trim(UCase(Request("txtPayTypeCd")))
    if Trim(Request("txtXchRt")) <> "" then
		I5_m_iv_hdr(M528_I4_xch_rt)				= UniConvNum(Request("txtXchRt"),0)
	else
		I5_m_iv_hdr(M528_I4_xch_rt)				= "0"
	end if
    I5_m_iv_hdr(M528_I4_pay_terms_txt)			= Trim(Request("txtPayTermsTxt"))
    I5_m_iv_hdr(M528_I4_tax_biz_area)			= Trim(UCase(Request("txtBizAreaCd")))
    I5_m_iv_hdr(M528_I4_remark)					= Trim(Request("txtMemo"))
    '추가 
    I5_m_iv_hdr(M528_I4_xch_rate_op)			= Trim(Request("hdnDiv"))
    
    I5_m_iv_hdr(M528_I4_issue_dt_fg)			= Trim(Request("hdIssueDTFg"))
    
    I4_b_company						= gCurrency
    I6_m_user_id			= Trim(Request("hdnUsrId"))
    
    If Request("txtChkPoNo") = "Y" Then 

		 I5_m_iv_hdr(M528_I4_ref_po_no) = Trim(Request("txtPoNo"))
	Else
		 I5_m_iv_hdr(M528_I4_ref_po_no) = ""
	End if
	    
    If lgIntFlgMode = OPMD_CMODE Then
		I5_m_iv_hdr(M528_I4_iv_no)			= Trim(UCase(Request("txtIvNo1")))
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		I5_m_iv_hdr(M528_I4_iv_no)			= Trim(Request("hdnIvNo"))
		iCommandSent = "UPDATE"
    End If
    
	Set iPM8G111 = Server.CreateObject("PM8G111.cMMaintIvHdrS")

	If CheckSYSTEMError(Err,True) = True Then
       Set iPM8G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	pvCB = "F"
	
    If lgIntFlgMode = OPMD_CMODE Then
		E1_Return = iPM8G111.M_MAINT_IV_HDR_SVR(pvCB, gStrGlobalCollection, Cstr(iCommandSent), _
												Cstr(I1_m_iv_type_iv_type_cd), _
												Cstr(I2_b_biz_partner_bp_cd), _
												Cstr(I3_b_pur_grp), _
												Cstr(I4_b_company), _
												I5_m_iv_hdr, Cstr(I6_m_user_id))
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		E1_Return = iPM8G111.M_MAINT_IV_HDR_SVR(pvCB, gStrGlobalCollection, cstr(iCommandSent), _
												Cstr(I1_m_iv_type_iv_type_cd), _
												Cstr(I2_b_biz_partner_bp_cd), _
												Cstr(I3_b_pur_grp),  _
												Cstr(I4_b_company), _
												I5_m_iv_hdr, cstr(I6_m_user_id))
    End If
    
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM8G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM8G111 = Nothing

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "If """ & lgIntFlgMode & """  = """ & OPMD_CMODE & """  Then " & vbCr
	Response.Write "  .frm1.txtIvNo.value	= """ & ConvSPChars(E1_Return) & """" & vbCr
	Response.Write "  .frm1.hdnIvNo.value	= """ & ConvSPChars(E1_Return) & """" & vbCr
	Response.Write "End If" & vbCr
	Response.Write " .DbSaveOk" & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr
End Sub

'============================================================================================================
' Name : SubReleaseCheck
' Desc : 
'============================================================================================================
Sub SubReleaseCheck()																'☜: 회계처리,회계처리취소 요청을 받음	

    On Error Resume Next
    Err.Clear                                                                       '☜: Protect system from crashing

    Dim I2_ief_supplied
    Dim IG1_imp_dtl_group
    Dim I3_m_batch_ap_post_wks
    
    Const M557_I3_ap_dt_type = 0
    Const M557_I3_ap_dt = 1
    Const M557_I3_import_flg = 2
    
    Const M557_IG1_I1_count = 0
    Const M557_IG1_I2_iv_no = 1
    Const M557_IG1_I3_ap_dt = 2

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
    
	
    If Trim(Request("hdnApFlg")) = "Y" Then
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
    Response.Write "<Script Language=VBScript>" & vbCr
    Response.Write "With parent" & vbCr
    Response.Write " .MainQuery()" & vbCr
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr

    Set PM8G211 = Nothing                                                   '☜: Unload Comproxy
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizDelete()
	Dim iPM8G111		
    Dim I1_m_iv_type_iv_type_cd
    Dim I2_b_biz_partner_bp_cd
    Dim I3_b_pur_grp
    Dim I4_b_company
    Dim I5_m_iv_hdr
    Dim I6_m_user_id
    Dim E1_Return
    
	Const M528_I4_iv_no = 13
	
	On Error Resume Next
	Err.Clear 
	
	'For 전자세금계산서	
	Call CheckNoForDT(Trim(Request("txtIvNo")))

	Set iPM8G111 = Server.CreateObject("PM8G111.cMMaintIvHdrS")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM8G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Redim I5_m_iv_hdr(59)
    
    I5_m_iv_hdr(M528_I4_iv_no) = Trim(Request("txtIvNo"))
	
	pvCB = "F"
	
    E1_Return = iPM8G111.M_MAINT_IV_HDR_SVR(pvCB, gStrGlobalCollection, "DELETE", cstr(I1_m_iv_type_iv_type_cd), _
				cstr(I2_b_biz_partner_bp_cd),cstr(I3_b_pur_grp),cstr(I4_b_company),I5_m_iv_hdr,cstr(I6_m_user_id))
    
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM8G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM8G111 = Nothing

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "Call parent.DbDeleteOk()" & vbCr
	Response.Write "</Script>" & vbCr
	        
End Sub
'============================================================================================================
' Name : SubLookUpSupplier
' Desc : 
'============================================================================================================
Sub SubLookUpSupplier()
	Dim BpType, BpCd
	Dim iPB5CS41
	Dim E1_b_biz_partner
	
	Const S074_E1_bp_rgst_no = 2
	Const S074_E1_currency = 17
	Const S074_E1_pay_meth = 29
	Const S074_E1_pay_dur = 30
	Const S074_E1_vat_type = 33
	Const S074_E1_pay_type = 45
	Const S074_E1_pay_terms_txt = 46
	Const S074_E1_vat_type_nm = 124                           '[부가세유형명]
	Const S074_E1_pay_meth_nm = 133       
	Const S074_E1_pay_type_nm = 134 
	'추가(구매용)
	Const S074_E1_pay_meth_pur = 115                          '결재방법(구매)
    Const S074_E1_pay_type_pur = 116                          '입출금유형(구매)
    Const S074_E1_pay_dur_pur = 117                           '결재기간(구매)
    '네임추가 
    Const S074_E1_pay_meth_pur_nm = 141                       '[결재방법명(구매)]
    Const S074_E1_pay_type_pur_nm = 142                       '[입출금유형명(구매)]

    On Error Resume Next
    Err.Clear
    
    Set iPB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")

    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing
        Exit Sub
    End If

    BpType = Trim(Request("txtBpType"))
    BpCd = Trim(Request("txtBpCd"))
    
    Call iPB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection,"QUERY",BpCd,E1_b_biz_partner) 
    
    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing
        Exit Sub
    End If
    
    Set iPB5CS41 = Nothing     
    
    Response.Write "<Script Language=VBScript>" & vbCr
    Response.Write "With parent.frm1" & vbCr
    Response.Write "If  """ & BpType & """ = ""1"" then " & vbCr       '공급처경우 
    Response.Write "    .txtCur.Value             = """   & ConvSPChars(E1_b_biz_partner(S074_E1_currency))    & """" & vbCr
    Response.Write "    .txtCurNm.Value           = """"" & vbCr
    Response.Write " parent.GetPayDt()"   & vbCr
    Response.Write " parent.ChangeCurr()" & vbCr
	'***2002.12월 패치********
    Response.Write "ElseIf """ & BpType & """ = ""2"" then" & vbCr  '지급처인 경우 
    Response.Write "  If .ChkPoNo.checked = False  then " & vbCr
    Response.Write "    .txtPayMethCd.Value       = """   & ConvSPChars(E1_b_biz_partner(S074_E1_pay_meth_pur))    & """" & vbCr
    Response.Write "    .txtPayMethNm.Value       = """	  & ConvSPChars(E1_b_biz_partner(S074_E1_pay_meth_pur_nm))      & """" & vbCr
    Response.Write "  End If"             & vbCr

    Response.Write "  .txtPayDur.Value            = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_dur_pur))       & """" & vbCr
    Response.Write "  .txtPayTermstxt.Value       = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_terms_txt)) & """" & vbCr
    Response.Write "  .txtPayTypeCd.Value         = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_type_pur))      & """" & vbCr
    Response.Write "  .txtPayTypeNm.Value         = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_type_pur_nm))      & """" & vbCr
    Response.Write "ElseIf """ & BpType & """  = ""3""  then" & vbCr  '세금계산서발행처인 경우 
    Response.Write "  .txtVatCd.Value             = """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type))      & """" & vbCr
    Response.Write "  .txtVatNm.Value             = """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type_nm))   & """" & vbCr
    'Response.Write "  .txtSpplRegNo.Value        = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_rgst_no))    & """" & vbCr
    Response.Write " parent.SetVatType()" & vbCr
    Response.Write "End If"               & vbCr
    Response.Write "End With"             & vbCr
    Response.Write "</Script>"            & vbCr
                                                           '☜: Process End
End Sub

'============================================================================================================
' Name : SubLookUpPo
' Desc :
'============================================================================================================
Sub SubLookUpPo()
	Const M239_E22_release_flg = 0
    Const M239_E22_merg_pur_flg = 1
    Const M239_E22_po_no = 2
    Const M239_E22_po_dt = 3
    Const M239_E22_xch_rt = 4
    Const M239_E22_vat_type = 5
    Const M239_E22_vat_rt = 6
    Const M239_E22_tot_vat_doc_amt = 7
    Const M239_E22_tot_po_doc_amt = 8
    Const M239_E22_tot_po_loc_amt = 9
    Const M239_E22_pay_meth = 10
    Const M239_E22_pay_dur = 11
    Const M239_E22_pay_terms_txt = 12
    Const M239_E22_pay_type = 13
    Const M239_E22_sppl_sales_prsn = 14
    Const M239_E22_sppl_tel_no = 15
    Const M239_E22_remark = 16
    Const M239_E22_vat_inc_flag = 17
    Const M239_E22_offer_dt = 18
    Const M239_E22_fore_dvry_dt = 19
    Const M239_E22_expiry_dt = 20
    Const M239_E22_invoice_no = 21
    Const M239_E22_incoterms = 22
    Const M239_E22_transport = 23
    Const M239_E22_sending_bank = 24
    Const M239_E22_delivery_plce = 25
    Const M239_E22_applicant = 26
    Const M239_E22_manufacturer = 27
    Const M239_E22_agent = 28
    Const M239_E22_origin = 29
    Const M239_E22_packing_cond = 30
    Const M239_E22_inspect_means = 31
    Const M239_E22_dischge_city = 32
    Const M239_E22_dischge_port = 33
    Const M239_E22_loading_port = 34
    Const M239_E22_shipment = 35
    Const M239_E22_import_flg = 36
    Const M239_E22_bl_flg = 37
    Const M239_E22_cc_flg = 38
    Const M239_E22_rcpt_flg = 39
    Const M239_E22_subcontra_flg = 40
    Const M239_E22_ret_flg = 41
    Const M239_E22_iv_flg = 42
    Const M239_E22_rcpt_type = 43
    Const M239_E22_issue_type = 44
    Const M239_E22_iv_type = 45
    Const M239_E22_po_cur = 46
    Const M239_E22_xch_rate_op = 47
    Const M239_E22_pur_org = 48
    Const M239_E22_pur_biz_area = 49
    Const M239_E22_pur_cost_cd = 50
    Const M239_E22_tot_vat_loc_amt = 51
    Const M239_E22_cls_flg = 52
    Const M239_E22_lc_flg = 53
    Const M239_E22_sppl_cd = 54
    Const M239_E22_payee_cd = 55
    Const M239_E22_build_cd = 56
    Const M239_E22_charge_flg = 55
    Const M239_E22_tracking_no = 56
    Const M239_E22_so_no = 57
    Const M239_E22_inspect_method = 58
    Const M239_E22_ref_no = 59
    Const M239_E22_ext1_cd = 60
    Const M239_E22_ext1_qty = 61
    Const M239_E22_ext1_amt = 62
    Const M239_E22_ext1_rt = 63
    Const M239_E22_ext1_dt = 64
    Const M239_E22_ext2_cd = 65
    Const M239_E22_ext2_qty = 66
    Const M239_E22_ext2_amt = 67
    Const M239_E22_ext2_rt = 68
    Const M239_E22_ext2_dt = 69
    Const M239_E22_ext3_cd = 70
    Const M239_E22_ext3_qty = 71
    Const M239_E22_ext3_amt = 72
    Const M239_E22_ext3_rt = 73
    Const M239_E22_ext3_dt = 74

    Const M239_E22_issue_dt_fg = 86

    
    Const M239_E19_po_type_cd = 0
    Const M239_E19_po_type_nm = 1
    
    Const M239_E9_bp_cd = 0
	Const M239_E9_bp_nm = 1
    
    Const M239_E20_pur_grp = 0
    Const M239_E20_pur_grp_nm = 1
    
    
    Dim M31119
    Dim E1_b_bank_bank_nm
    Dim E2_b_minor_vat_type
    Dim E3_b_minor_pay_meth
    Dim E4_b_minor_pay_type
    Dim E5_b_minor_incoterms
    Dim E6_b_minor_transport
    Dim E7_b_minor_delivery_plce
    Dim E8_b_minor_origin
    Dim E9_b_biz_partner
    Dim E10_b_biz_partner_applicant_nm
    Dim E11_b_biz_partner_manufacturer_nm
    Dim E12_b_minor_packing_cond
    Dim E13_b_minor_inspect_means
    Dim E14_b_minor_dischge_city
    Dim E15_b_minor_dischge_port
    Dim E16_b_minor_loading_port
    Dim E17_b_configuration_reference
    Dim E18_b_currency_currency_desc
    Dim E19_m_config_process
    Dim E20_b_pur_grp
    Dim E21_b_biz_partner_agent_nm
    Dim E22_m_pur_ord_hdr
    Dim iCommandSent, iPoNo
    
    On Error Resume Next
    Err.Clear
    
    iPoNo = Trim(Request("txtPoNo"))    
    Set M31119 = Server.CreateObject("PM3G119.cMLookupPurOrdHdrS")

    If CheckSYSTEMError(Err, True) = True Then
        Exit Sub
    End If
    
     Call M31119.M_LOOKUP_PUR_ORD_HDR_SVR(gStrGlobalCollection, _
                                      iPoNo, E1_b_bank_bank_nm, E2_b_minor_vat_type, _
                                      E3_b_minor_pay_meth, E4_b_minor_pay_type, E5_b_minor_incoterms, _
                                      E6_b_minor_transport, E7_b_minor_delivery_plce, E8_b_minor_origin, _
                                      E9_b_biz_partner, E10_b_biz_partner_applicant_nm, _
                                      E11_b_biz_partner_manufacturer_nm, E12_b_minor_packing_cond, _
                                      E13_b_minor_inspect_means, E14_b_minor_dischge_city, _
                                      E15_b_minor_dischge_port, E16_b_minor_loading_port, _
                                      E17_b_configuration_reference, E18_b_currency_currency_desc, _
                                      E19_m_config_process, E20_b_pur_grp, _
                                      E21_b_biz_partner_agent_nm, E22_m_pur_ord_hdr)
                                                  
    If CheckSYSTEMError2(Err, True, "", "", "", "", "") = True Then
        Set M31119 = Nothing                                                '☜: ComProxy Unload
        Exit Sub                                                            '☜: 비지니스 로직 처리를 종료함 
     End If

   'If UCase(Trim(E22_m_pur_ord_hdr(M239_E22_ret_flg))) = "Y" Then
   '   Call DisplayMsgBox("17a014", vbOKOnly, "반품발주건", "조회", I_MKSCRIPT)
   '   Set M31119 = Nothing                                                                 '☜: ComProxy UnLoad
   '   Exit Sub                                                            '☜: 비지니스 로직 처리를 종료함 
   'End If
   
   Set M31119 = Nothing                                                                 '☜: ComProxy UnLoad

    '-----------------------
    'LookUp Iv Name
    '-----------------------
    '===================
    Dim strIvTypeNm
    Dim strPayeeCd
	Dim strPayeeNm
	Dim strBuildCd
	Dim strBuildNm
	
	Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
	If  E22_m_pur_ord_hdr(M239_E22_iv_type) <> "" Or E22_m_pur_ord_hdr(M239_E22_iv_type) <> Null then  			
		lgStrSQL = "select iv_type_nm from m_iv_type " 
		lgStrSQL = lgStrSQL & " WHERE iv_type_cd =  " & FilterVar(UCase(E22_m_pur_ord_hdr(M239_E22_iv_type)), "''", "S") & "" 
		
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			strIvTypeNm = ""
		Else
			strIvTypeNm	= lgObjRs("iv_type_nm")
		End If
	End If
		
	lgStrSQL = "SELECT A.BP_CD, A.BP_NM  FROM B_BIZ_PARTNER A, B_BIZ_PARTNER_FTN B  " 
	lgStrSQL = lgStrSQL & " WHERE B.PARTNER_BP_CD = A.BP_CD AND B.DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.PARTNER_FTN = " & FilterVar("MPA", "''", "S") & " "
	lgStrSQL = lgStrSQL & " AND B.BP_CD =  " & FilterVar(E9_b_biz_partner(M239_E9_bp_cd), "''", "S") & ""  
		
	IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
		strPayeeCd = ""
		strPayeeNm = ""
	Else
		strPayeeCd	= lgObjRs("BP_CD")
		strPayeeNm	= lgObjRs("BP_NM")
	End If
		
	lgStrSQL = "SELECT A.BP_CD, A.BP_NM  FROM B_BIZ_PARTNER A, B_BIZ_PARTNER_FTN B  " 
	lgStrSQL = lgStrSQL & " WHERE B.PARTNER_BP_CD = A.BP_CD AND B.DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.PARTNER_FTN = " & FilterVar("MBI", "''", "S") & " "
	lgStrSQL = lgStrSQL & " AND B.BP_CD =  " & FilterVar(E9_b_biz_partner(M239_E9_bp_cd), "''", "S") & ""  
		
	IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
		strBuildCd = ""
		strBuildNm = ""
	Else
		strBuildCd	= lgObjRs("BP_CD")
		strBuildNm	= lgObjRs("BP_NM")
	End If
		
	Call SubCloseRs(lgObjRs)
	Call SubCloseDB(lgObjConn)
    
    '========================


    Dim StrtxtBuildCd
    If E22_m_pur_ord_hdr(M239_E22_build_cd) = "" Then
        StrtxtBuildCd = E9_b_biz_partner(M239_E9_bp_cd)
    Else
        StrtxtBuildCd = E22_m_pur_ord_hdr(M239_E22_build_cd)
    End If

    Response.Write "<Script Language=VBScript>" & vbCr
    Response.Write "With parent" & vbCr
    '##### Rounding Logic #####
        '항상 거래화폐가 우선 
    Response.Write ".frm1.txtCur.value          = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)) & """" & vbCR
    Response.Write "parent.CurFormatNumericOCX" & vbCr
    '##########################
    Response.Write ".frm1.txtIvAmt.text         = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtIvLocAmt.text      = """ & UNIConvNumDBToCompanyByCurrency(0, gCurrency, ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtVatAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, gTaxRndPolicyNo, "X") & """" & vbCr
    
    Response.Write ".frm1.txtnetDocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtnetLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(0, gCurrency, ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtVatLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, gTaxRndPolicyNo, "X") & """" & vbCr
    
    Response.Write ".frm1.txtGrpCd.Value        = """ & ConvSPChars(E20_b_pur_grp(M239_E20_pur_grp)) & """" & vbCr
    Response.Write ".frm1.txtGrpNm.Value        = """ & ConvSPChars(E20_b_pur_grp(M239_E20_pur_grp_nm)) & """" & vbCr
    Response.Write ".frm1.txtSpplCd.Value       = """ & ConvSPChars(E9_b_biz_partner(M239_E9_bp_cd)) & """" & vbCr
    Response.Write ".frm1.hdnSpplCd.value       = """ & ConvSPChars(E9_b_biz_partner(M239_E9_bp_cd)) & """" & vbCr
    Response.Write ".frm1.txtSpplNm.Value       = """ & ConvSPChars(E9_b_biz_partner(M239_E9_bp_nm)) & """" & vbCr

    'Response.Write ".frm1.txtPayeeCd.Value      = """ & ConvSPChars(E3_b_biz_partner(B132_E3_bp_cd)) & """" & vbCr
    'Response.Write ".frm1.txtPayeeNm.Value      = """ & ConvSPChars(E3_b_biz_partner(B132_E3_bp_nm)) & """" & vbCr

    'Response.Write ".frm1.txtBuildCd.Value      = """ & ConvSPChars(E2_b_biz_partner(B132_E2_bp_cd)) & """" & vbCr
    'Response.Write ".frm1.txtBuildNm.Value      = """ & ConvSPChars(E2_b_biz_partner(B132_E2_bp_nm)) & """" & vbCr
'=========
    Response.Write ".frm1.txtPayeeCd.Value      = """ & ConvSPChars(strPayeeCd) & """" & vbCr
    Response.Write ".frm1.txtPayeeNm.Value      = """ & ConvSPChars(strPayeeNm) & """" & vbCr

    Response.Write ".frm1.txtBuildCd.Value      = """ & ConvSPChars(strBuildCd) & """" & vbCr
    Response.Write ".frm1.txtBuildNm.Value      = """ & ConvSPChars(strBuildNm) & """" & vbCr
'=========
    Response.Write ".frm1.hdnCur.Value          = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)) & """" & vbCr
    
    Response.Write ".frm1.txtCurNm.Value        = """ & ConvSPChars(E18_b_currency_currency_desc) & """" & vbCr
    Response.Write ".frm1.txtVatCd.Value        = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_vat_type)) & """" & vbCr
    Response.Write ".frm1.txtVatNm.Value        = """ & ConvSPChars(E2_b_minor_vat_type) & """" & vbCr
    Response.Write ".frm1.txtVatRt.Text         = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_vat_rt), ggExchRate.DecPoint, 0) & """" & vbCr
    Response.Write ".frm1.txtPayMethCd.Value    = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_meth)) & """" & vbCr
    Response.Write ".frm1.hdnPayMethCd.Value    = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_meth)) & """" & vbCr
        
    Response.Write ".frm1.txtPayMethNm.Value    = """ & ConvSPChars(E3_b_minor_pay_meth) & """" & vbCr
    Response.Write ".frm1.txtPayDur.Text        = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_pay_dur), 0, 0) & """" & vbCr
    Response.Write ".frm1.txtPayTypeCd.Value    = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_type)) & """" & vbCr
    Response.Write ".frm1.txtPayTypeNm.Value    = """ & ConvSPChars(E4_b_minor_pay_type) & """" & vbCr
    Response.Write ".frm1.txtPayTermsTxt.Value  = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_terms_txt)) & """" & vbCr
    Response.Write ".frm1.txtIvTypeCd.Value     = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_iv_type)) & """" & vbCr
    Response.Write ".frm1.txtIvTypeNm.value     = """ & strIvTypeNm & """" & vbCr
    Response.Write ".frm1.txtXchRt.text         = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_xch_rt), ggExchRate.DecPoint, 0) & """" & vbCr
    Response.Write ".frm1.txtPoNo.Value         = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_no)) & """" & vbCr
    
    Response.Write ".frm1.hdnDiv.value           = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_xch_rate_op)) & """" & vbcr
    
    Response.Write "If """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_vat_inc_flag)) & """ = ""2""  then " & vbCr   'vat 포함 여부 
    Response.Write "  .frm1.rdoVatFlg2.Checked = true " & vbCr
    Response.Write "  .frm1.hdvatFlg.value  = ""2""" & vbCr
    Response.Write "Else" & vbCr
    Response.Write "  .frm1.rdoVatFlg1.Checked = true " & vbCr
    Response.Write "  .frm1.hdvatFlg.value  = ""1""" & vbCr
    Response.Write "End If  " & vbCr
    
    
    Response.Write "If """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_no)) & """  <> """" Then " & vbCr
    Response.Write "  .frm1.chkPoNo.checked = True" & vbCr
    Response.Write "  .frm1.txtChkPoNo.value = ""Y"" " & vbCr
    Response.Write "End If" & vbCr
    

    '===========
    '반품여부 Flag추가(2003.08.01)
    Response.Write "If """ & ConvSPChars(UCase(E22_m_pur_ord_hdr(M239_E22_ret_flg))) & """ = ""Y"" Then " & vbCr
    Response.Write "  .frm1.hdnRetflg.value = ""Y"" " & vbCr
    Response.Write "Else " & vbCr
    Response.Write "  .frm1.hdnRetflg.value = ""N"" " & vbCr
    Response.Write "End If" & vbCr
    '============
    Response.Write "If Trim(.frm1.txtPayeeCd.Value) = """"  then " & vbCr
    Response.Write "  .frm1.txtPayeeCd.Value =  .frm1.txtSpplCd.value" & vbCr
    Response.Write "End If" & vbCr
    Response.Write "If Trim(.frm1.txtPayeeNm.Value) = """" then " & vbCr
    Response.Write "  .frm1.txtPayeeNm.Value =  .frm1.txtSpplNm.value" & vbCr
    Response.Write "End If" & vbCr

    Response.Write "If Trim(.frm1.txtBuildCd.Value) = """" then " & vbCr
    Response.Write "  .frm1.txtBuildCd.Value =  .frm1.txtSpplCd.value " & vbCr
    Response.Write "End If" & vbCr
    Response.Write "If Trim(.frm1.txtBuildNm.Value) = """" then" & vbCr
    Response.Write "  .frm1.txtBuildNm.Value =  .frm1.txtSpplNm.value" & vbCr
    Response.Write "End If" & vbCr
    
    
    Response.Write "parent.GetPayDt()" & vbCr
    Response.Write "parent.GetTaxBizArea(""BP"")" & vbCr
    Response.Write "parent.ChangeTag(False)" & vbCr
    'parent.DbPoQueryOK()
    
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr
'---------------------------
	On Error Resume Next
	Err.Clear

	Dim BpType
	Dim iPB5CS41
	Dim E8_b_biz_partner
	
	Const S074_E1_bp_rgst_no = 2

    On Error Resume Next
    Err.Clear

    Set iPB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")

    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing	
        Exit Sub
    End If

    BpType = Request("txtBpType")
    
    Call iPB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection,"QUERY",StrtxtBuildCd,E8_b_biz_partner) 

    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing
        Exit Sub
    End If
    
    Response.Write "<Script Language=VBScript>" & vbCr
    Response.Write "Parent.frm1.txtSpplRegNo.Value	= """ & ConvSPChars(E8_b_biz_partner(S074_E1_bp_rgst_no)) & """" & vbCr
    Response.Write "parent.DbPoQueryOK()" & vbCr
	Response.Write "</Script>" & vbCr

    Set iPB5CS41 = Nothing                                                   '☜: Unload Comproxy
    Set M31119   = Nothing                                                   '☜: Unload Comproxy
    Set iPB5GS45 = Nothing

End Sub


Sub CheckNoForDT(ByVal I1_no)
	
	'For 전자세금계산서 
	Dim pvObjRs
	Dim pvStrSQL
	
	Dim lgObjComm
    Dim IntRetCD
    
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
