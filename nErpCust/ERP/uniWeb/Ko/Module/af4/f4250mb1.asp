<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 자금 
'*  3. Program ID           : f4201mb1
'*  4. Program Name         : 차입금등록(query)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/07/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :																	
'**********************************************************************************************

	On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status  
    
    Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")  
    
    Call HideStatusWnd                                                              '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	
	Dim strLoanNo
	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


	lgErrorStatus	= "NO"
	lgErrorPos		= ""                                                           '☜: Set to space
	lgOpModeCRUD	= Request("txtMode")					'☜ : 현재 상태를 받음 

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))	
	

	'------ Developer Coding part (Start ) ------------------------------------------------------------------Dim strMode

	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount
	Dim iCommandSent

    iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    iCommandSent	= Request("txtPrevNext")
    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------     
Select Case lgOpModeCRUD
     Case CStr(UID_M0001)                                                         '☜: Query
          Call SubBizQuery()

     Case CStr(UID_M0002)       
          Call SubBizSave()

     Case CStr(UID_M0003)                                                         '☜: Delete
          Call SubBizDelete()
End Select

Response.end
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    
    Dim strLoanTermYr
    Dim strLoanTermMnth
    
    Const A859_f_ln_pay_no = 0
    Const A859_f_ln_loan_no = 1
    Const A859_f_ln_loan_nm = 2
    Const A859_f_ln_loan_dt = 3
    Const A859_f_ln_due_dt = 4
    Const A859_f_ln_loan_plc_type = 5
    Const A859_f_ln_loan_bank_cd = 6
    Const A859_f_ln_loan_bank_nm = 7
    Const A859_f_ln_bp_cd = 8
    Const A859_f_ln_bp_nm = 9
    Const A859_f_ln_loan_fg = 10
    Const A859_f_ln_gl_no = 11
    Const A859_f_ln_temp_gl_no = 12
    Const A859_f_ln_loan_type = 13
    Const A859_f_ln_loan_type_nm = 14
    Const A859_f_ln_loan_amt = 15
    Const A859_f_ln_loan_loc_amt = 16
    Const A859_f_ln_loan_int_rate = 17
    Const A859_f_ln_rdp_amt = 18
    Const A859_f_ln_rdp_loc_amt = 19
    Const A859_f_ln_int_pay_amt = 20
    Const A859_f_ln_int_pay_loc_amt = 21
    Const A859_f_ln_loan_bal_amt = 22
    Const A859_f_ln_loan_bal_loc_amt = 23
    Const A859_f_ln_mean_type = 24
    Const A859_f_ln_mean_type_nm = 25
    Const A859_f_ln_bank_cd = 26
    Const A859_f_ln_bank_nm = 27
    Const A859_f_ln_bank_acct_no = 28
    Const A859_f_ln_pr_amt = 29
    Const A859_f_ln_pr_loc_amt = 30
    Const A859_f_ln_pi_amt = 31
    Const A859_f_ln_pi_loc_amt = 32
    Const A859_f_ln_dfr_amt = 33
    Const A859_f_ln_dfr_loc_amt = 34
    Const A859_f_ln_pay_xch_rate = 35
    Const A859_f_ln_loan_xch_rate = 36
    Const A859_f_ln_dfr_xch_rate = 37
    Const A859_f_ln_pay_dt = 38
    Const A859_f_ln_pay_plan_dt = 39
    Const A859_f_ln_dept_cd = 40
    Const A859_f_ln_dept_nm = 41
    Const A859_f_ln_doc_cur = 42
    Const A859_f_ln_bc_amt = 43
    Const A859_f_ln_bc_loc_amt = 44
    Const A859_f_ln_bp_amt = 45
    Const A859_f_ln_bp_loc_amt = 46
    Const A859_f_ln_int_pay_stnd = 47
    Const A859_f_ln_user_fld1 = 48
    Const A859_f_ln_user_fld2 = 49
    Const A859_f_ln_repay_desc = 50
    Const A859_f_ln_int_pay_perd = 51
    Const A859_f_ln_int_pay_perd_base = 52
    Const A859_f_ln_day_mthd = 53
    Const A859_f_ln_conf_fg = 54
    Const A859_f_ln_pay_mean_acct_cd = 55
    Const A859_f_ln_pay_mean_acct_nm = 56
    Const A859_f_ln_int_pay_acct_cd = 57
    Const A859_f_ln_int_pay_acct_nm = 58
    Const A859_f_ln_bc_acct_cd = 59
    Const A859_f_ln_bc_acct_nm = 60
    Const A859_f_ln_bp_acct_cd = 61
    Const A859_f_ln_bp_acct_nm = 62
    Const A859_f_ln_org_change_id = 63
	
	Const A859_I1_f_ln_repay_pay_no = 0

	Dim iPAFG425LOOKUP																'☆ : 조회용 Component Dll 사용 변수 

	Dim I1_f_ln_repay
	Dim E_f_ln_info
	Dim E_PrevNext_Code
	
	Redim I1_f_ln_repay(A859_I1_f_ln_repay_pay_no+4)
	I1_f_ln_repay(A859_I1_f_ln_repay_pay_no+1) = lgAuthBizAreaCd
	I1_f_ln_repay(A859_I1_f_ln_repay_pay_no+2) = lgInternalCd
	I1_f_ln_repay(A859_I1_f_ln_repay_pay_no+3) = lgSubInternalCd
	I1_f_ln_repay(A859_I1_f_ln_repay_pay_no+4) = lgAuthUsrID		

    I1_f_ln_repay(A859_I1_f_ln_repay_pay_no) = Trim(Request("txtKeyStream"))

    Set iPAFG425LOOKUP = server.CreateObject("PAFG425.bFLkUpRepaySvr")

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If    
    
    Call iPAFG425LOOKUP.F_LOOKUP_REPAY_SVR(gStrGlobalCollection, iCommandSent, I1_f_ln_repay, E_f_ln_info, E_PrevNext_Code)																			

    If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG425LOOKUP = Nothing
		Exit Sub
    End If
	Set iPAFG425LOOKUP = Nothing

	If Trim(E_PrevNext_Code(0)) = "900011" or Trim(E_PrevNext_Code(0)) = "900012" Then
		Call DisplayMsgBox(E_PrevNext_Code(0), VbOKOnly, "", "", I_MKSCRIPT)
	End If

    Response.Write "<Script Language=vbscript>  " & vbCr
    Response.Write " with parent.frm1"			  & vbCr     
    Response.Write " .hPayNo.value				= """ & ConvSPChars(E_f_ln_info(A859_f_ln_pay_no)) & """			" & vbCr
    Response.Write " .txtPayNo.value			= """ & ConvSPChars(E_f_ln_info(A859_f_ln_pay_no)) & """			" & vbCr
    Response.Write " .txtLoanNo.value			= """ & ConvSPChars(E_f_ln_info(A859_f_ln_loan_no)) & """			" & vbCr
    Response.Write " .txtLoanNm.value			= """ & ConvSPChars(E_f_ln_info(A859_f_ln_loan_nm)) & """			" & vbCr
    Response.Write " .txtLoanDt.Text			= """ & UNIDateClientFormat(E_f_ln_info(A859_f_ln_loan_dt)) & """	" & vbCr
    Response.Write " .txtDueDt.Text				= """ & UNIDateClientFormat(E_f_ln_info(A859_f_ln_due_dt)) & """	" & vbCr
    If ConvSPChars(E_f_ln_info(A859_f_ln_loan_plc_type)) = "BK" Then
		Response.Write " .txtLoanPlcType1.Checked	= """ & True & """												" & vbCr
		Response.Write " .txtLoanPlcCd.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_loan_bank_cd)) & """	" & vbCr
		Response.Write " .txtLoanPlcNm.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_loan_bank_nm)) & """	" & vbCr
	Else
		Response.Write " .txtLoanPlcType2.Checked	= """ & True & """												" & vbCr
		Response.Write " .txtLoanPlcCd.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bp_cd)) & """			" & vbCr
		Response.Write " .txtLoanPlcNm.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bp_nm)) & """			" & vbCr
	End If
    Response.Write " .cboLoanFg.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_loan_fg)) & """			" & vbCr
    Response.Write " .txtGlNo.value			= """ & ConvSPChars(E_f_ln_info(A859_f_ln_gl_no)) & """				" & vbCr
    Response.Write " .txtTempGlNo.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_temp_gl_no)) & """		" & vbCr
    Response.Write " .txtLoanType.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_loan_type)) & """			" & vbCr
    Response.Write " .txtLoantypeNm.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_loan_type_nm)) & """		" & vbCr
    Response.Write " .txtLoanAmt.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_loan_amt),			ggAmtOfMoney.DecPoint	,0) & """" & vbCr
    Response.Write " .txtLoanLocAmt.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_loan_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtLoanIntRate.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_loan_int_rate),	ggExchRate.DecPoint		,0)	& """" & vbCr
    Response.Write " .txtRdpAmt.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_rdp_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtRdpLocAmt.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_rdp_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtIntPayAmt.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_int_pay_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtIntPayLocAmt.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_int_pay_loc_amt),	ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtLoanBalAmt.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_loan_bal_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtLoanBalLocAmt.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_loan_bal_loc_amt),	ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtRcptTypeCd.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_mean_type)) & """				" & vbCr
    Response.Write " .txtRcptTypeNm.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_mean_type_nm)) & """			" & vbCr
    Response.Write " .txtBankCd.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bank_cd)) & """				" & vbCr
    Response.Write " .txtBankNm.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bank_nm)) & """				" & vbCr
    Response.Write " .txtBankAcctNo.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bank_acct_no)) & """			" & vbCr
    Response.Write " .txtPlanAmtPR.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pr_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtPlanLocAmtPR.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pr_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hPlanAmtPR.value		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pr_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hPlanLocAmtPR.value	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pr_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtPlanAmtPI.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pi_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtPlanLocAmtPI.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pi_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hPlanAmtPI.value		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pi_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hPlanLocAmtPI.value	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pi_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtDfrIntPay.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_dfr_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtDfrIntPayLoc.Text	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_dfr_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtXchRate.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_pay_xch_rate),		ggExchRate.DecPoint		,0)	& """" & vbCr
    Response.Write " .hXchRate.Text			= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_loan_xch_rate),	ggExchRate.DecPoint		,0)	& """" & vbCr
'    Response.Write " .hDfrXchRate.value		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_dfr_xch_rate),		ggExchRate.DecPoint		,0)	& """" & vbCr
    Response.Write " .txtPayPlanDt.Text		= """ & UNIDateClientFormat(E_f_ln_info(A859_f_ln_pay_dt)) & """		" & vbCr
    Response.Write " .hPayPlanDt.value		= """ & UNIDateClientFormat(E_f_ln_info(A859_f_ln_pay_plan_dt)) & """	" & vbCr
    Response.Write " .txtDeptCd.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_dept_cd)) & """				" & vbCr
    Response.Write " .txtDeptNm.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_dept_nm)) & """				" & vbCr
    Response.Write " .txtDocCur.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_doc_cur)) & """				" & vbCr
    Response.Write " .txtEtcPay.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bc_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtEtcPayLoc.Text		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bc_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hEtcPay.value			= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bc_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hEtcPayLoc.value		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bc_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtEtcBPPay.value		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bp_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .txtEtcBPPayLoc.value	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bp_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hEtcBPPay.value		= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bp_amt),			ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hEtcBPPayLoc.value	= """ & UNINumClientFormat(E_f_ln_info(A859_f_ln_bp_loc_amt),		ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
    Response.Write " .hIntPayStnd.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_int_pay_stnd)) & """			" & vbCr
	If ConvSPChars(E_f_ln_info(A859_f_ln_int_pay_stnd)) = "AI" Then
		Response.Write " .txtIntPayStnd1.Checked	= """ & True & """												" & vbCr
	ElseIf ConvSPChars(E_f_ln_info(A859_f_ln_int_pay_stnd)) = "DI" Then
		Response.Write " .txtIntPayStnd2.Checked	= """ & True & """												" & vbCr
	End If
    Response.Write " .txtUserFld1.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_user_fld1)) & """				" & vbCr
    Response.Write " .txtUserFld2.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_user_fld2)) & """				" & vbCr
    Response.Write " .txtRepayDesc.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_repay_desc)) & """			" & vbCr
    Response.Write " .hInt_pay_perd.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_int_pay_perd)) & """			" & vbCr
    Response.Write " .hint_pay_perd_base.value= """ & ConvSPChars(E_f_ln_info(A859_f_ln_int_pay_perd_base)) & """	" & vbCr
    Response.Write " .hDay_Mthd.value		= """ & ConvSPChars(E_f_ln_info(A859_f_ln_day_mthd)) & """				" & vbCr
    Response.Write " .hConfFg.value			= """ & ConvSPChars(E_f_ln_info(A859_f_ln_conf_fg)) & """				" & vbCr
    Response.Write " .txtRcptAcctCd.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_pay_mean_acct_cd)) & """		" & vbCr
    Response.Write " .txtRcptAcctNm.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_pay_mean_acct_nm)) & """		" & vbCr
    Response.Write " .txtIntPayAcctCd.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_int_pay_acct_cd)) & """			" & vbCr
    Response.Write " .txtIntPayAcctNm.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_int_pay_acct_nm)) & """			" & vbCr
    Response.Write " .txtChargeAcctCd.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bc_acct_cd)) & """			" & vbCr
    Response.Write " .txtChargeAcctNm.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bc_acct_nm)) & """			" & vbCr
    Response.Write " .txtEtcBPAcctCd.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bp_acct_cd)) & """			" & vbCr
    Response.Write " .txtEtcBPAcctNm.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_bp_acct_nm)) & """			" & vbCr
    Response.Write " .hOrgChangeId.value	= """ & ConvSPChars(E_f_ln_info(A859_f_ln_org_change_id)) & """			" & vbCr

    Response.Write "End with                " & vbCr
    Response.Write "Parent.DbQueryOk        " & vbCr
    Response.Write "</Script>               " & vbCr

End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data To Db
'============================================================================================================
Sub SubBizSave()

	Dim iPAFG425CU

	Dim I1_f_ln_repay
	Dim E1_b_auto_numbering 
	Dim E2_b_auto_numbering

	Dim lgIntFlgMode

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear        

    Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A854_I2_a_data_auth_data_BizAreaCd = 0
    Const A854_I2_a_data_auth_data_internal_cd = 1
    Const A854_I2_a_data_auth_data_sub_internal_cd = 2
    Const A854_I2_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A854_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A854_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A854_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A854_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

    Const A855_f_ln_loan_no = 0
    Const A855_f_ln_pay_no = 1
    Const A855_f_ln_pay_plan_dt = 2
    Const A855_f_ln_pay_dt = 3
    Const A855_f_ln_dept_cd = 4
    Const A855_f_ln_doc_cur = 5
    Const A855_f_ln_mean_type = 6
    Const A855_f_ln_pay_mean_acct_cd = 7
    Const A855_f_ln_bank_cd = 8
    Const A855_f_ln_bank_acct_no = 9
    Const A855_f_ln_pr_amt = 10
    Const A855_f_ln_pr_loc_amt = 11
    Const A855_f_ln_pi_amt = 12
    Const A855_f_ln_pi_loc_amt = 13
    Const A855_f_ln_dfr_amt = 14
    Const A855_f_ln_dfr_loc_amt = 15
    Const A855_f_ln_bc_amt = 16
    Const A855_f_ln_bc_loc_amt = 17
    Const A855_f_ln_bp_amt = 18
    Const A855_f_ln_bp_loc_amt = 19
    Const A855_f_ln_int_pay_acct_cd = 20
    Const A855_f_ln_bc_acct_cd = 21
    Const A855_f_ln_bp_acct_cd = 22
    Const A855_f_ln_org_change_id = 23
    Const A855_f_ln_user_fld1 = 24
    Const A855_f_ln_user_fld2 = 25
    Const A855_f_ln_repay_desc = 26
	Redim I1_f_ln_repay(A855_f_ln_repay_desc)
	
		I1_f_ln_repay(A855_f_ln_loan_no)			= UCase(Trim(Request("txtLoanNo")))
		I1_f_ln_repay(A855_f_ln_pay_no)				= UCase(Trim(Request("hPayNo")))
		I1_f_ln_repay(A855_f_ln_pay_plan_dt)		= UNIConvDate(Trim(Request("hPayPlanDt")))
		I1_f_ln_repay(A855_f_ln_pay_dt)				= UNIConvDate(Request("txtPayPlanDt"))
		I1_f_ln_repay(A855_f_ln_dept_cd)			= UCase(Trim(Request("txtDeptCd")))
		I1_f_ln_repay(A855_f_ln_mean_type)			= UCase(Trim(Request("txtRcptTypeCd")))
		I1_f_ln_repay(A855_f_ln_doc_cur)			= UCase(Trim(Request("txtDocCur")))
		I1_f_ln_repay(A855_f_ln_pay_mean_acct_cd)	= UCase(Trim(Request("txtRcptAcctCd")))
		I1_f_ln_repay(A855_f_ln_bank_cd)			= UCase(Trim(Request("txtBankCd")))
		I1_f_ln_repay(A855_f_ln_bank_acct_no)		= Trim(Request("txtBankAcctNo"))
		I1_f_ln_repay(A855_f_ln_pr_amt)				= UNIConvNum(Request("txtPlanAmtPR"),0)
		I1_f_ln_repay(A855_f_ln_pr_loc_amt)			= UNIConvNum(Request("txtPlanLocAmtPR"),0)
		I1_f_ln_repay(A855_f_ln_pi_amt)				= UNIConvNum(Request("txtPlanAmtPI"),0)
		I1_f_ln_repay(A855_f_ln_pi_loc_amt)			= UNIConvNum(Request("txtPlanLocAmtPI"),0)
		I1_f_ln_repay(A855_f_ln_dfr_amt)			= UNIConvNum(Request("txtDfrIntPay"),0)
		I1_f_ln_repay(A855_f_ln_dfr_loc_amt)		= UNIConvNum(Request("txtDfrIntPayLoc"),0)
		I1_f_ln_repay(A855_f_ln_bc_amt)				= UNIConvNum(Request("txtEtcPay"),0)
		I1_f_ln_repay(A855_f_ln_bc_loc_amt)			= UNIConvNum(Request("txtEtcPayLoc"),0)
		I1_f_ln_repay(A855_f_ln_bp_amt)				= UNIConvNum(Request("txtEtcBPPay"),0)
		I1_f_ln_repay(A855_f_ln_bp_loc_amt)			= UNIConvNum(Request("txtEtcBPPayLoc"),0)
		I1_f_ln_repay(A855_f_ln_int_pay_acct_cd)	= UCase(Trim(Request("txtIntPayAcctCd")))
		I1_f_ln_repay(A855_f_ln_bc_acct_cd)			= UCase(Trim(Request("txtChargeAcctCd")))
		I1_f_ln_repay(A855_f_ln_bp_acct_cd)			= UCase(Trim(Request("txtEtcBPAcctCd")))
		I1_f_ln_repay(A855_f_ln_org_change_id)		= Trim(Request("hOrgChangeId"))
		I1_f_ln_repay(A855_f_ln_user_fld1)			= Trim(Request("txtUserFld1"))
		I1_f_ln_repay(A855_f_ln_user_fld2)			= Trim(Request("txtUserFld2"))
		I1_f_ln_repay(A855_f_ln_repay_desc)			= Trim(Request("txtRepayDesc"))

    Set iPAFG425CU = server.CreateObject("PAFG425.cFMngRepaySvr")   

    If CheckSYSTEMError(Err,True) = True Then				
		Exit Sub
    End If
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
	        
    Select Case lgIntFlgMode   
		Case  OPMD_CMODE                                                             '☜ : Create
			I1_f_ln_repay(A855_f_ln_pay_no)				= UCase(Trim(Request("txtPayNo")))
			Call iPAFG425CU.F_MANAGE_REPAY_SVR(gStrGlobalCollection, "C",	I1_f_ln_repay, E1_b_auto_numbering,I2_a_data_auth)
			I1_f_ln_repay(A855_f_ln_pay_no) = E1_b_auto_numbering(0)
        Case  OPMD_UMODE           
			Call iPAFG425CU.F_MANAGE_REPAY_SVR(gStrGlobalCollection, "U",	I1_f_ln_repay,,I2_a_data_auth)		
    End Select

	If CheckSYSTEMError(Err,True) = True Then		
		Set iPAFG425CU = Nothing
		Exit Sub	
    End If
		 
    Set iPAFG425CU = Nothing	
	
	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk(""" & I1_f_ln_repay(A855_f_ln_pay_no)	& """)	" & vbCr
    Response.Write "</Script>									" & vbCr    

End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : DELETE Data 
'============================================================================================================
Sub SubBizDelete()

	Dim iPAFG425D
	Dim I1_f_ln_repay
	Dim E1_f_ln_repay

    Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A854_I2_a_data_auth_data_BizAreaCd = 0
    Const A854_I2_a_data_auth_data_internal_cd = 1
    Const A854_I2_a_data_auth_data_sub_internal_cd = 2
    Const A854_I2_a_data_auth_data_auth_usr_id = 3 
 
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear        

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A854_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A854_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A854_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A854_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	I1_f_ln_repay = Split(Request("txtKeyStream"), gColSep)

	Set iPAFG425D = server.CreateObject("PAFG425.cFMngRepaySvr")   
	    
	If CheckSYSTEMError(Err, True) = True Then					
	   Exit Sub
	End If    

	Call iPAFG425D.F_MANAGE_REPAY_SVR(gStrGlobalCollection,	"D",I1_f_ln_repay,	E1_f_ln_repay,I2_a_data_auth)
	    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG425D = Nothing
		Exit Sub
	End If
		 
	Set iPAFG425D = Nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
	Response.Write "</Script>                   " & vbCr

End Sub

Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
%>

